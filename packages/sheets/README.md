# @mrnote/sheets

> **For the full experience, install the MR Office desktop app (free): https://github.com/dongmingyang2007-pixel/MROffice/releases** (or https://mroffice.ai) — it adds opening and saving Word, Excel and PowerPoint files, formulas with functions, charts, SmartArt editing, and saving PDF edits back into the file.

The spreadsheet editor of [MR Office](https://mroffice.ai) (its SpreadsheetWindow: formula engine, number formats, conditional formatting, charts, pivots, sort and filter, frozen panes, many sheets), in its browser-only mode over a local file.

| Feature | Status |
|---|---|
| New workbook; open `.mrsheet` (the editor's JSON), `.csv`, `.tsv` | Available, verified (a CSV opens with its values) |
| Formulas, formatting, charts, sorting and filtering, sheets | The editor's own browser code; not each verified one by one in this build |
| Save | `.mrsheet` keeps everything; a CSV/TSV saves back as CSV/TSV (values of the first sheet) |
| Known issue | Rows draw taller than their row headers: the Office substitute fonts (fetched at build time in MR Office, `scripts/sync-office-fonts.mjs` there) are not fetched by this repository |
| Open or save **Excel (.xlsx)** | Not in the open-source editor: the MR Office desktop app (free, [download](https://github.com/dongmingyang2007-pixel/MROffice/releases)), open and save Excel files |
| Background calculation of large workbooks, data connections, AI | Need the service (its engine host) or are AI; not included |
