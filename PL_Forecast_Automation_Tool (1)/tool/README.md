# P&L Forecast Automation Tool

A small desktop app (Tkinter) that automates updating your FY P&L Summary
workbook each month: it re-points every external file link to this month's
source files, refreshes their cached values, and gives you a form for the
cells that have to be typed in by hand (like the Connectivity row in
Relief/Reclass) — all **without ever breaking a link**.

## Why it's safe

Your workbook has 38 external links (checked directly against the file you
uploaded). If you open this file in Excel, edit it, and hit save, Excel/most
tools that touch it (including plain openpyxl) will happily rewrite the whole
package — and **that's what breaks links**: cached values get wiped and the
file goes to `#REF!`/`#NAME?` until every source is found again.

This tool never re-saves the workbook as a whole. It edits only the exact XML
parts that need to change:
- the `Target` path inside each external link's `.rels` file (so Excel knows
  where the new source file is), and
- the cached numbers inside that same external link's XML (so the numbers
  look right immediately), and
- the `<v>` of the specific manual-input cell you typed a new number into.

Every formula, every sheet, every style, and every *other* link is copied
through byte-for-byte, untouched. If a link can't be matched to a file in your
input folder, the tool leaves it exactly as it was rather than guessing or
deleting it — it will never silently break a reference.

## Setup (one time)

1. Install Python 3.9+ (Windows: python.org, check "Add to PATH"). Tkinter
   ships with the standard Windows installer — nothing extra to do there.
2. Open a terminal in this folder and run:
   ```
   pip install openpyxl
   ```

## Running it

```
python app.py
```

### Step 1 — Base workbook & scenario
- **Browse…** to this month's P&L Summary workbook (the one you uploaded is
  the FC5+7 file, but any month's version works the same way).
- Click **Scan Links** — it reads all external links plus the manual-input
  candidate cells (see the "Manual Inputs" tab) directly from the file.
- Pick the forecasting cycle you're producing: **F2+10, F5+7, F7+5, F10+2**.
  This is used (a) to prefer files with the matching tag when several
  candidates look similar, and (b) to rename the output file to match.

### Step 2 — Input folder
- **Browse…** to the folder holding this month's received files (source
  workbooks for Reclass, Relief, HQ Tagetik, FX, etc).
- Click **Auto-Match Files** — every one of the 38 links gets matched against
  files in that folder (recursively) by comparing filename word overlap plus
  fuzzy string similarity, with a bonus for matching scenario tags
  (e.g. "F5+7"). Matches are colour-coded green (matched) / red (unmatched).
- Any link can be corrected: **double-click its row** to browse for the
  correct file by hand. Manual picks are locked in (100% confidence) and
  survive re-matching if you switch scenarios.

### Step 3 — Manual Inputs tab
- Lists every hardcoded cell the scanner found sitting inside an otherwise
  formula-driven row on the Reclass and Relief tabs (this is where the
  Connectivity row example lives). Double-click the **Value** column to type
  in this month's number.

### Step 4 — Generate
- Click **Generate Updated File**. If any link is still unmatched you'll get
  a clear warning naming which ones — you can still proceed (that link is
  simply left as-is, not broken) or go back and match it first.
- Choose a save folder. The file is named after the original with the
  scenario tag swapped in (e.g. `...F5+7...` → `...F7+5...`).
- Open the result in Excel. If Excel prompts to update links, choose
  **Update Values** so every formula recalculates against the newly linked
  files.

## Files in this folder

| File | Purpose |
|---|---|
| `app.py` | The GUI — run this. |
| `link_scanner.py` | Finds every external link and which sheet/cell uses it (raw XML, not openpyxl, so nothing is lost). |
| `file_matcher.py` | Scores candidate files in your input folder against each link's original filename. |
| `manual_scanner.py` | Finds hardcoded cells sitting in formula-driven rows (Reclass/Relief by default). |
| `xml_updater.py` | The safe, link-preserving edit engine described above. |

## Adjusting what counts as "manual input"

By default the manual scanner only looks at the **Reclass** and **Relief**
tabs (columns C–AJ), since that's where you pointed out the mixed
manual/linked rows. To scan other tabs too, open `manual_scanner.py` and add
sheet names to `TARGET_SHEETS_DEFAULT`, or pass a list into
`scan_manual_cells(path, target_sheets=[...])`.

## Known limitations

- This tool re-points links and refreshes cached numbers; it does **not**
  evaluate Excel formulas itself. Real recalculation happens when you open
  the file in Excel (or hit F9). This is intentional — recalculating
  Excel's formula language independently would risk mismatched results, so
  the tool defers to Excel, the source of truth, for that step.
- The 4 external links with no visible formula (old, unused legacy links,
  e.g. old Japanese file paths) are left untouched either way — they aren't
  referenced by any active cell.
- It only matches `.xlsx`/`.xlsm`/`.xls` files; if a source file has a very
  different name from the original link, use the double-click override.
