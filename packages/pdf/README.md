# @mrnote/pdf

> **For the full experience, install the MR Office desktop app (free): https://github.com/XinaTian/MROffice/releases** (or https://mroffice.ai) — it adds opening and saving Word, Excel and PowerPoint files, formulas with functions, charts, SmartArt editing, and saving PDF edits back into the file.

On npm as [`mroffice-pdf`](https://www.npmjs.com/package/mroffice-pdf): compiled, with its types and stylesheet, no build setup needed (`npm i mroffice-pdf mroffice-core`; built from this tree by `npm run pack:npm`, scripts/npm/).

The PDF editor of [MR Office](https://mroffice.ai), running in the browser on a local file: MR Office's PDF window (pdf.js rendering, MR Office's ribbon and tools) in its browser-only mode, where the file and everything done to it stay in the page.

This package is generated from MR Office's own source (see the repository README).

## What works in the open-source build

| Feature | Status |
|---|---|
| Open a PDF from this computer and read it: pages, zoom, single / two-page and continuous views, thumbnails, outline and bookmarks, read mode | Available |
| Search the document's text (the magnifier in the title row: every match, by page, and a jump to it) | Available |
| Presentation (Play, full screen) and Print (the browser's print) | Available |
| Comments and markup on selected text (highlight, underline, strikethrough, squiggly), notes and the comments panel | Available while the file is open (kept in the page, not written into the file) |
| Copy text, read aloud (the browser's speech) | Available |
| Save / Download | The file as it was opened. Comments and other changes are not written into the PDF |
| Edit text, pictures and shapes in place; insert, rotate, delete or reorder pages; fill and flatten forms; compress, watermark and background, redaction, encryption, PDF/A; conversion to Word / Excel / PowerPoint, Markdown export; version history | Not included: in MR Office these run in its PDF engine on the server, which is not part of the open-source editors. Use MR Office on the web or the desktop app |
| OCR, screenshot text, translation, AI (Copilot) | Not included |
