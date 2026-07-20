"""
manual_scanner.py
Finds "manual input" cells: hardcoded numeric cells (a plain <v>, no <f>) that
sit inside a row/table that is otherwise built from formulas, and rows that are
100% hardcoded within a sheet's normal data range. These are exactly the kind of
cell the user described in Reclass/Relief: some months of a row are pulled from
formulas/links, some are typed straight in.

Only scans sheets given in TARGET_SHEETS by default (Reclass, Relief are the two
the user named) but any visible sheet name can be passed in.
"""
import re
import zipfile
from dataclasses import dataclass, field

TARGET_SHEETS_DEFAULT = ['Reclass', 'Relief']

# columns considered "data" columns (monthly / period columns) - adjust as needed
DATA_COL_RANGE = ('C', 'AJ')


@dataclass
class ManualCell:
    sheet: str
    cell: str
    row: int
    col: str
    row_label: str
    current_value: str


def _col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n


def _sheet_name_map(z):
    wb_xml = z.read('xl/workbook.xml').decode('utf-8', errors='replace')
    rels_xml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8', errors='replace')
    rid_to_target = dict(re.findall(r'<Relationship Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_xml))
    sheets = re.findall(r'<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"', wb_xml)
    mapping = {}
    for name, rid in sheets:
        target = rid_to_target.get(rid)
        if target:
            norm = 'xl/' + target if not target.startswith('xl/') else target
            mapping[name.replace('&amp;', '&')] = norm
    return mapping


def _shared_strings(z):
    if 'xl/sharedStrings.xml' not in z.namelist():
        return []
    xml = z.read('xl/sharedStrings.xml').decode('utf-8', errors='replace')
    items = re.findall(r'<si>(.*?)</si>', xml, re.S)
    out = []
    for it in items:
        texts = re.findall(r'<t[^>]*>(.*?)</t>', it, re.S)
        out.append(''.join(texts))
    return out


def scan_manual_cells(xlsx_path, target_sheets=None):
    target_sheets = target_sheets or TARGET_SHEETS_DEFAULT
    lo_col, hi_col = _col_to_num(DATA_COL_RANGE[0]), _col_to_num(DATA_COL_RANGE[1])

    results = []
    with zipfile.ZipFile(xlsx_path) as z:
        name_to_file = _sheet_name_map(z)
        shared = _shared_strings(z)

        for sheet_name in target_sheets:
            sf = name_to_file.get(sheet_name)
            if not sf or sf not in z.namelist():
                continue
            xml = z.read(sf).decode('utf-8', errors='replace')

            for row_match in re.finditer(r'<row r="(\d+)"[^>]*>(.*?)</row>', xml, re.S):
                row_num = int(row_match.group(1))
                row_xml = row_match.group(2)

                cells = re.findall(
                    r'<c r="([A-Z]+)(\d+)"(?:[^>]*t="([^"]*)")?[^>]*>(?:<f[^>]*>[^<]*</f>|<f[^>]*/>)?(?:<v>([^<]*)</v>)?</c>',
                    row_xml,
                )
                if not cells:
                    continue

                # row label = first non-empty text cell in columns A/B
                row_label = ''
                for col, _rn, ctype, val in cells:
                    if col in ('A', 'B') and ctype == 's' and val:
                        try:
                            row_label = shared[int(val)]
                        except (ValueError, IndexError):
                            pass
                        if row_label:
                            break

                has_formula_in_row = '<f>' in row_xml or '<f ' in row_xml
                if not has_formula_in_row:
                    continue  # not a formula-driven row, skip (not the pattern we're after)

                for col, _rn, ctype, val in cells:
                    if _col_to_num(col) < lo_col or _col_to_num(col) > hi_col:
                        continue
                    if ctype:  # has a type attr (string, bool, etc.) -> not a plain hardcoded number
                        continue
                    if val is None or val == '':
                        continue
                    # check this specific cell has no <f> (formula) tag
                    cell_full = re.search(
                        r'<c r="%s%d"[^>]*>(.*?)</c>' % (re.escape(col), row_num), row_xml
                    )
                    if cell_full and '<f' in cell_full.group(1):
                        continue  # it's a formula cell, skip
                    results.append(ManualCell(
                        sheet=sheet_name, cell=f'{col}{row_num}', row=row_num,
                        col=col, row_label=row_label, current_value=val,
                    ))
    return results


if __name__ == '__main__':
    import sys
    cells = scan_manual_cells(sys.argv[1])
    for c in cells:
        print(f'{c.sheet}!{c.cell}  ({c.row_label!r})  = {c.current_value}')
    print(f'\nTotal manual-input candidate cells: {len(cells)}')
