# MR Office — English (UK)

<p align="center"><sub>🌐 <a href="../README.md">English</a> · <a href="../i18n/README.en-GB.md">English (UK)</a> · <a href="../i18n/README.zh.md">简体中文</a> · <a href="../i18n/README.zh-TW.md">繁體中文</a> · <a href="../i18n/README.ja.md">日本語</a> · <a href="../i18n/README.ko.md">한국어</a> · <a href="../i18n/README.de.md">Deutsch</a> · <a href="../i18n/README.fr.md">Français</a> · <a href="../i18n/README.es.md">Español</a> · <a href="../i18n/README.es-419.md">Español (Latinoamérica)</a> · <a href="../i18n/README.pt-BR.md">Português (Brasil)</a> · <a href="../i18n/README.it.md">Italiano</a> · <a href="../i18n/README.nl.md">Nederlands</a> · <a href="../i18n/README.sv.md">Svenska</a> · <a href="../i18n/README.da.md">Dansk</a> · <a href="../i18n/README.no.md">Norsk</a> · <a href="../i18n/README.fi.md">Suomi</a> · <a href="../i18n/README.id.md">Bahasa Indonesia</a> · <a href="../i18n/README.th.md">ไทย</a> · <a href="../i18n/README.vi.md">Tiếng Việt</a></sub></p>

<p align="center"><b>The open-source editors of MR Office — documents, spreadsheets, presentations, PDF, handwriting and code — running in the browser.</b></p>

<p align="center"><a href="https://mroffice.ai/en-GB"><b>▶ Try the live demo — no install</b></a> · <a href="https://github.com/XinaTian/MROffice/releases">⬇ Download MR Office from Releases</a> · <a href="https://mroffice.ai">🌐 mroffice.ai</a></p>

<p align="center"><a href="../LICENSE"><img alt="AGPL-3.0" src="https://img.shields.io/badge/license-AGPL--3.0-blue"></a> <a href="https://github.com/XinaTian/MROffice/releases"><img alt="release" src="https://img.shields.io/github/v/release/XinaTian/MROffice"></a> <a href="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml/badge.svg"></a> <a href="https://github.com/XinaTian/MROffice/stargazers"><img alt="GitHub stars" src="https://img.shields.io/github/stars/XinaTian/MROffice?style=social"></a></p>

<p align="center"><a href="https://mroffice.ai/en-GB"><img src="../docs/media/mr-office-demo.webp" alt="The open-source editors of MR Office — documents, spreadsheets, presentations, PDF, handwriting and code — running in the browser." width="100%"></a></p>

<p align="center"><a href="https://github.com/user-attachments/assets/3dbe2320-c587-497b-a554-73e469909387">🔊 Watch the 50-second tour with sound</a></p>

<table>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/en-GB"><img src="../docs/screenshots/docs.png" alt="Documents" width="100%"></a><br><b>Documents</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/en-GB"><img src="../docs/screenshots/sheets.png" alt="Spreadsheets" width="100%"></a><br><b>Spreadsheets</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/en-GB"><img src="../docs/screenshots/slides.png" alt="Presentations" width="100%"></a><br><b>Presentations</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/en-GB"><img src="../docs/screenshots/pdf.png" alt="PDF" width="100%"></a><br><b>PDF</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/en-GB"><img src="../docs/screenshots/ink.png" alt="Handwriting" width="100%"></a><br><b>Handwriting</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/en-GB"><img src="../docs/screenshots/code.png" alt="Code" width="100%"></a><br><b>Code</b></td>
</tr>
</table>

MR Office is a workspace for notes, AI and real files. This repository publishes the source code of its editors: the same editors MR Office runs, which you can build, run and embed on your own. They are open-source editors, not a complete open-source office suite: the reading and writing of Microsoft Office files stays in MR Office (see below).

## The editors

What each editor does in this repository:

- **Documents** — a paged word processor: styles, fonts, lists, tables, headers and footers, sections and columns, shapes and equations, cover pages and tables of contents, comments and tracked changes, and pen ink on the page. Saves its own format (.mrdoc); opens HTML, Markdown and plain text.
- **Spreadsheets** — cells and formulas, number formats, conditional formatting, sorting and filtering, frozen panes and several sheets, with sums and averages of the selection. Arithmetic between cells (=A1*B1+10) is calculated in the browser as you type; formulas with functions (SUM, IF, VLOOKUP …) are kept as written. Saves its own format (.mrsheet); opens and saves CSV and TSV.
- **Presentations** — slides with eight themes of its own, layouts, slide master and layout views, text, shapes, text boxes, tables, WordArt and equations, comments, sections, transitions and animations with preview, and a slide show. Saves its own format (.mrslides).
- **PDF** — read, search, present and print PDF files, copy and read aloud their text, and mark up selected text (highlight, underline, strikethrough, squiggly) with notes and a comments panel while a file is open; saving hands back the file as it was opened.
- **Handwriting** — handwritten notes with pressure-sensitive pens, highlighters, erasers, lasso and ruler, and ink turned into shapes. Saves its own format (.mrink).
- **Code** — a lightweight code editor with syntax highlighting for many languages, an outline, and find and replace; JavaScript/TypeScript, Python, SQL, Lua and shell run right in the browser.

Calculating formulae with functions, drawing charts and laying out SmartArt run in MR Office's own engine, in the MR Office desktop app: the open-source editors keep the formulae, charts and SmartArt a file holds, and show SmartArt as it was saved.

## Run the editors

**▶ [Try the live demo — no install](https://mroffice.ai/en-GB)**

The six editors are in `packages/` (`@mrnote/docs`, `sheets`, `slides`, `pdf`, `ink`, `code`, sharing `@mrnote/core`) and run together in `apps/playground`:

```bash
npm install
npm run fonts     # the open substitutes for Office's fonts
npm run build
npm run preview   # http://localhost:5178
```

`npm run dev` starts a development server at the same address, and `npm run smoke` checks a build. Each package's README lists what it does in detail and shows how to embed it.

## Desktop app

[`apps/desktop`](../apps/desktop) holds **MR Office Editors**: the source of the MR Office desktop app's shell (Electron), under an identity of its own. With the free MR Office component from Releases, it is the MR Office desktop app, with Word, Excel, PowerPoint and PDF files, sign-in and AI. Without the component, it runs the open-source editors on their own, which open .mrdoc, .mrsheet, .mrslides and .mrink files, CSV, HTML, Markdown and PDF. The shell and the editors are AGPL source code; the component is compiled, free to use and not open source. After `npm run build`, start it with:

```bash
npm run desktop
```

## For the full experience, install the MR Office desktop app (free)

The MR Office desktop app is free, has these editors and the full service, and works offline on your own computer. It adds what the open-source editors leave out:

- Opening and saving Word, Excel and PowerPoint files (.docx, .xlsx, .pptx)
- Formulae with functions (SUM, IF, VLOOKUP …), calculated
- Charts
- SmartArt editing
- Saving PDF edits back into the file

**[Download MR Office from Releases →](https://github.com/XinaTian/MROffice/releases)** · [or from mroffice.ai](https://mroffice.ai)

### Word, Excel, PowerPoint and PDF files

Opening and saving Word, Excel and PowerPoint files (.docx, .xlsx, .pptx), and saving PDF edits back into the file, are provided by the MR Office desktop app (free), which has all of these editors and works offline on your own computer:

## How the pieces fit

```mermaid
flowchart LR
  subgraph OSS["Open source · AGPL-3.0"]
    direction LR
    core["@mrnote/core"] --> E["The six editors<br/>Documents · Spreadsheets · Presentations · PDF · Handwriting · Code"]
    E --> PG["apps/playground<br/>Browser"]
    E --> DS["apps/desktop<br/>Desktop shell"]
    MCP["apps/mcp<br/>MCP server for AI apps"]
  end
  COMP["MR Office component: Word, Excel and PowerPoint files · free, not open source"] -.-> DS
```

## License

Open source under the GNU Affero General Public License v3.0 (LICENSE), with the additional terms in NOTICE.md: keep the "MR Office" attribution and link to https://mroffice.ai, and no rights to the MR Office name or logo. A commercial licence is available from support@mr-note.com. Third-party software and data are listed in THIRD_PARTY_NOTICES.md. Contributions: see CONTRIBUTING.md.

## More

Website and the full product: **https://mroffice.ai** · Questions and feedback: support@mr-note.com

<p align="center">If MR Office is useful to you, a ⭐ on GitHub helps other people find it.</p>
