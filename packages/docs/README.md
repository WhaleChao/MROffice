# @mrnote/docs

> **For the full experience, install the MR Office desktop app (free): https://github.com/XinaTian/MROffice/releases** (or https://mroffice.ai) — it adds opening and saving Word, Excel and PowerPoint files, formulas with functions, charts, SmartArt editing, and saving PDF edits back into the file.

On npm as [`mroffice-docs`](https://www.npmjs.com/package/mroffice-docs): compiled, with its types and stylesheet, no build setup needed (`npm i mroffice-docs mroffice-core`; built from this tree by `npm run pack:npm`, scripts/npm/).

The document editor of [MR Office](https://mroffice.ai): a paged word processor (MR Office's NoteEditor on Tiptap/ProseMirror) with the Office-style ribbon, running in the browser on a local file.

## What works in the open-source build

| Feature | Status |
|---|---|
| New document; open `.mrdoc` (MR Office's document JSON) and `.html` | Available, verified |
| Paged editing: styles, fonts, lists, tables, headers and footers, sections, columns, page setup | Available |
| Insert: pictures, shapes, text boxes, equations, symbols, charts, SmartArt | Available (charts and SmartArt draw in the browser) |
| Review: comments and track changes (kept in the document), find and replace, word count | Available |
| Draw: pens, highlighters, erasers on the page | Available (see @mrnote/handwriting = packages/ink) |
| Save | Available: writes `.mrdoc` (the editor's own JSON, reopened as it was) |
| Open or save **Word (.docx)** | Not in the open-source editor: the MR Office desktop app (free, [download](https://github.com/XinaTian/MROffice/releases)), open and save Word files |
| Export PDF, mail merge to .docx, cloud fonts | Need the service; not available |
| AI (Copilot, AI pictures, rewrite), collaboration, version history | Part of the MR Office service; not included |
