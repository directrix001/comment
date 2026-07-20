"""
file_matcher.py
Given the list of external-link filenames pulled from the workbook, and a
folder the user points at, suggest the best matching file for each link using
token-overlap scoring (not just raw string similarity) so that files renamed
slightly month to month (dates, versions, "final", "v2"...) still match.
"""
import os
import re
from difflib import SequenceMatcher

STOPWORDS = {
    'the', 'for', 'new', 'file', 'files', 'final', 'update', 'v1', 'v2', 'v3',
    'working', 'received', 'summary', 'consolidation', 'and', 'with', 'to',
}

SCENARIO_TAGS = ['2+10', '5+7', '7+5', '10+2']


def _tokens(name):
    base = os.path.splitext(name)[0].lower()
    raw = re.split(r'[^a-z0-9+]+', base)
    return [t for t in raw if t and t not in STOPWORDS]


def _scenario_tag(name):
    low = name.lower()
    for tag in SCENARIO_TAGS:
        if tag in low:
            return tag
    return None


def find_candidate_files(folder):
    """Recursively list every .xlsx / .xlsm / .xls file under folder."""
    found = []
    for root, _dirs, files in os.walk(folder):
        for f in files:
            if f.lower().endswith(('.xlsx', '.xlsm', '.xls')) and not f.startswith('~$'):
                found.append(os.path.join(root, f))
    return found


def score_match(target_filename, candidate_path, preferred_scenario=None):
    cand_name = os.path.basename(candidate_path)
    t_tokens, c_tokens = set(_tokens(target_filename)), set(_tokens(cand_name))
    if not t_tokens or not c_tokens:
        overlap = 0.0
    else:
        overlap = len(t_tokens & c_tokens) / len(t_tokens | c_tokens)

    ratio = SequenceMatcher(None, target_filename.lower(), cand_name.lower()).ratio()
    score = 0.6 * overlap + 0.4 * ratio

    # bonus if scenario tag matches the one the user is generating for
    if preferred_scenario:
        t_tag = _scenario_tag(target_filename)
        c_tag = _scenario_tag(cand_name)
        if t_tag and c_tag and t_tag == c_tag == preferred_scenario:
            score += 0.15
        elif c_tag == preferred_scenario:
            score += 0.05
    return score


def match_links_to_folder(links, folder, preferred_scenario=None, min_score=0.25):
    """
    links: dict[int -> ExternalLink] (from link_scanner.scan_workbook)
    Returns dict[index -> {'best': path_or_None, 'score': float, 'candidates': [(path, score), ...top5]}]
    """
    candidates = find_candidate_files(folder)
    results = {}
    for idx, link in links.items():
        scored = [
            (path, score_match(link.filename, path, preferred_scenario))
            for path in candidates
        ]
        scored.sort(key=lambda x: x[1], reverse=True)
        top = scored[:5]
        best_path, best_score = (top[0] if top else (None, 0.0))
        if best_score < min_score:
            best_path = None
        results[idx] = {
            'best': best_path,
            'score': best_score,
            'candidates': top,
        }
    return results
