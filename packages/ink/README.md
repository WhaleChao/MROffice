# @mrnote/ink

> **For the full experience, install the MR Office desktop app (free): https://github.com/XinaTian/MROffice/releases** (or https://mroffice.ai) — it adds opening and saving Word, Excel and PowerPoint files, formulas with functions, charts, SmartArt editing, and saving PDF edits back into the file.

Handwriting, from [MR Office](https://mroffice.ai): a handwritten note is an MR Office document on handwriting paper, opened on the Draw tab. Pressure-sensitive pens, highlighters, pencils, erasers (stroke and point), lasso select, tape, stickers, laser pointer, ruler, ink replay, and ink turned into shapes. The page it writes on is the document editor's (@mrnote/docs).

## What works in the open-source build

| Feature | Status |
|---|---|
| New handwriting note (`.mrink`), draw with pens and highlighters, erase, lasso, ruler, add and delete pages | Available, verified (a stroke drawn, saved and reopened) |
| Ink to shape, ink replay, palm rejection, pen pressure and tilt | Available (in the browser) |
| Ink to text, ink to math | Not in the open-source build: recognition is done by a model in MR Office's service |
| Save | Available: writes `.mrink` (the document JSON, strokes included) |
| Live ink with collaborators, AI | Part of the MR Office service; not included |
