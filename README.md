# Target Village Dashboard

Open `index.html` in any browser (double-click — no server or internet needed; fonts load online if available).

- **Overview** – KPIs, margin-vs-turnout chart, category/priority mix, top opportunities
- **Villages** – sortable/filterable list of all 208 target villages; click any row for booth-level detail, map link, print sheet
- **Constituencies** – target villages vs whole-AC performance for the 12 ACs
- **Field tracker** – status / owner / notes per village (saved in the browser; export/import JSON to share)
- **Unmapped & notes** – villages with no booth match (with verification hints), shared booths, workbook notes
- **How to read** – definitions and the category/priority rules

Deep link to a village: `index.html#villages:v=<AC no>-<index>` (the URL updates when you open one).

## Rebuilding data
`data.js` is generated from `raw_data/*.xlsx` (target lists), `form20 up/2022/` (all-candidate booth results) and
`data-analysis/cleaned_data/booth_level/` (approx. coordinates). Sources are only read, never modified.

    ~/Downloads/anaconda3/bin/python build_data.py   # needs openpyxl + LibreOffice (soffice) for .xls

The build cross-checks BJP votes per booth between each workbook and Form-20 and prints any mismatch.
