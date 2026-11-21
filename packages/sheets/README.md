# @mrnote/sheets

> **For the full experience, install the MR Office desktop app (free): https://github.com/XinaTian/MROffice/releases** (or https://mroffice.ai) — it adds opening and saving Word, Excel and PowerPoint files, formulas with functions, charts, SmartArt editing, and saving PDF edits back into the file.

On npm as [`mroffice-sheets`](https://www.npmjs.com/package/mroffice-sheets): compiled, with its types and stylesheet, no build setup needed (`npm i mroffice-sheets mroffice-core`; built from this tree by `npm run pack:npm`, scripts/npm/).

The spreadsheet editor of [MR Office](https://mroffice.ai) (its SpreadsheetWindow: formula engine, number formats, conditional formatting, charts, pivots, sort and filter, frozen panes, many sheets), in its browser-only mode over a local file.

| Feature | Status |
|---|---|
| New workbook; open `.mrsheet` (the editor's JSON), `.csv`, `.tsv` | Available, verified (a CSV opens with its values) |
| Formulas, formatting, charts, sorting and filtering, sheets | The editor's own browser code; not each verified one by one in this build |
| Save | `.mrsheet` keeps everything; a CSV/TSV saves back as CSV/TSV (values of the first sheet) |
| Fonts | Cells are measured with the open fonts that stand in for Office's (Arimo for Arial, Carlito for Calibri...): `npm run fonts` fetches them for the playground; the npm packages load them from mroffice-assets |
| Open or save **Excel (.xlsx)** | Not in the open-source editor: the MR Office desktop app (free, [download](https://github.com/XinaTian/MROffice/releases)), open and save Excel files |
| Background calculation of large workbooks, data connections, AI | Need the service (its engine host) or are AI; not included |
