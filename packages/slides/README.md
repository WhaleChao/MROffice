# @mrnote/slides

> **For the full experience, install the MR Office desktop app (free): https://github.com/XinaTian/MROffice/releases** (or https://mroffice.ai) — it adds opening and saving Word, Excel and PowerPoint files, formulas with functions, charts, SmartArt editing, and saving PDF edits back into the file.

On npm as [`mroffice-slides`](https://www.npmjs.com/package/mroffice-slides): compiled, with its types and stylesheet, no build setup needed (`npm i mroffice-slides mroffice-core`; built from this tree by `npm run pack:npm`, scripts/npm/).

The presentation editor of [MR Office](https://mroffice.ai) (its PresentationWindow: themes, layouts, shapes, tables, charts, SmartArt, animations and transitions, presenter view), in its browser-only mode over a local file.

| Feature | Status |
|---|---|
| New presentation; open and save `.mrslides` (the editor's JSON) | Available, verified (new deck opens with the ribbon; save writes the deck) |
| Editing slides | Available: the slide canvas with its placeholders, the slide rail, notes and zoom |
| Open or save **PowerPoint (.pptx)** | Not in the open-source editor: the MR Office desktop app (free, [download](https://github.com/XinaTian/MROffice/releases)), open and save PowerPoint files |
| SmartArt layout and chart pictures drawn by the engine host, video export, AI (Designer, Copilot) | Need the service or are AI; not included |
