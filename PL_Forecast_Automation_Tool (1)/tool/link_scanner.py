"""
link_scanner.py
Scans an .xlsx (raw OOXML, not openpyxl) to find:
  - every external file link (workbook-level) with its resolved target path
  - every worksheet + cell that references each external link (via [N] index in formulas)

Raw XML is used deliberately: openpyxl silently drops external-link info on
re-save, and cross-referencing formula text is more reliable done directly on
the XML than through openpyxl's formula model for this purpose.
"""
import re
import os
import zipfile
import urllib.parse
from dataclasses import dataclass, field


@dataclass
class ExternalLink:
    index: int                 # the [N] used in formulas, e.g. [31]
    xml_file: str               # externalLinkN.xml
    raw_target: str             # raw Target attribute from the .rels file
    resolved_path: str          # decoded, best-guess real path on disk
    filename: str                # just the file name part
    sheets_referenced: set = field(default_factory=set)   # sheet names inside the external file
    used_in: list = field(default_factory=list)            # list of (local_sheet_name, cell_ref, formula)


def _sheet_name_map(z):
    """Map internal sheetN.xml file -> (sheet display name, visibility)."""
    wb_xml = z.read('xl/workbook.xml').decode('utf-8', errors='replace')
    rels_xml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8', errors='replace')

    rid_to_target = dict(re.findall(r'<Relationship Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_xml))

    sheets = re.findall(r'<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"', wb_xml)
    mapping = {}
    for name, rid in sheets:
        target = rid_to_target.get(rid)
        if target:
            # normalize e.g. "worksheets/sheet36.xml" -> "xl/worksheets/sheet36.xml"
            norm = 'xl/' + target if not target.startswith('xl/') else target
            mapping[norm] = name.replace('&amp;', '&')
    return mapping


def _external_link_targets(z):
    """externalLinkN.xml -> resolved path, from the matching .rels file."""
    names = z.namelist()
    result = {}
    link_files = [n for n in names if re.match(r'xl/externalLinks/externalLink\d+\.xml$', n)]
    for lf in sorted(link_files, key=lambda p: int(re.search(r'\d+', os.path.basename(p)).group())):
        base = os.path.basename(lf)
        rels_path = f'xl/externalLinks/_rels/{base}.rels'
        target = None
        if rels_path in names:
            rels_xml = z.read(rels_path).decode('utf-8', errors='replace')
            # prefer externalLinkPath relationship; take the last one listed (often the more complete path)
            matches = re.findall(r'<Relationship [^>]*Target="([^"]+)"[^>]*/?>', rels_xml)
            if matches:
                # pick the longest-looking path (usually the full UNC/drive path, not the relative one)
                target = max(matches, key=len)
        result[lf] = target
    return result


def _decode_target(raw_target):
    if not raw_target:
        return '(unknown / missing path)'
    t = raw_target
    if t.startswith('file:///'):
        t = t[len('file:///'):]
    t = urllib.parse.unquote(t)
    t = t.replace('/', '\\') if '\\' not in t and ':' in t[:3] else t
    return t


def scan_workbook(xlsx_path):
    """
    Returns: dict[int -> ExternalLink], plus prints/returns unresolved (missing) links.
    """
    with zipfile.ZipFile(xlsx_path) as z:
        sheet_name_of = _sheet_name_map(z)
        ext_targets = _external_link_targets(z)

        # index -> ExternalLink, keyed by the numeric N in externalLinkN.xml
        links = {}
        for xml_file, target in ext_targets.items():
            idx = int(re.search(r'externalLink(\d+)\.xml$', xml_file).group(1))
            resolved = _decode_target(target)
            links[idx] = ExternalLink(
                index=idx,
                xml_file=xml_file,
                raw_target=target,
                resolved_path=resolved,
                filename=os.path.basename(resolved) if resolved else '(unknown)',
            )
            # sheet names cached inside the external workbook
            try:
                ext_xml = z.read(xml_file).decode('utf-8', errors='replace')
                sheet_names = re.findall(r'<sheetName val="([^"]+)"', ext_xml)
                links[idx].sheets_referenced.update(sheet_names)
            except KeyError:
                pass

        # now scan every worksheet's formulas for [N] references
        sheet_files = [n for n in z.namelist() if re.match(r'xl/worksheets/sheet\d+\.xml$', n)]
        ref_pattern = re.compile(r'\[(\d+)\]')
        for sf in sheet_files:
            local_name = sheet_name_of.get(sf, sf)
            xml = z.read(sf).decode('utf-8', errors='replace')
            # walk rows/cells with formulas
            for cell_match in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*>(?:<f[^>]*>([^<]*)</f>)?', xml):
                cell_ref, formula = cell_match.group(1), cell_match.group(2)
                if not formula or '[' not in formula:
                    continue
                for m in ref_pattern.finditer(formula):
                    idx = int(m.group(1))
                    if idx in links:
                        links[idx].used_in.append((local_name, cell_ref, formula))

    return links


def summarize(links):
    lines = []
    for idx in sorted(links):
        L = links[idx]
        sheets_using = sorted(set(u[0] for u in L.used_in))
        lines.append(
            f"[{idx}] {L.filename}\n"
            f"     path: {L.resolved_path}\n"
            f"     used on sheets: {', '.join(sheets_using) if sheets_using else '(no active formula found)'}\n"
            f"     reference count: {len(L.used_in)}"
        )
    return '\n'.join(lines)


if __name__ == '__main__':
    import sys
    path = sys.argv[1]
    links = scan_workbook(path)
    print(summarize(links))
    print(f"\nTotal external links: {len(links)}")
