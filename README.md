# MR Office

[English](README.md) · [English (UK)](i18n/README.en-GB.md) · [简体中文](i18n/README.zh.md) · [繁體中文](i18n/README.zh-TW.md) · [日本語](i18n/README.ja.md) · [한국어](i18n/README.ko.md) · [Deutsch](i18n/README.de.md) · [Français](i18n/README.fr.md) · [Español](i18n/README.es.md) · [Español (Latinoamérica)](i18n/README.es-419.md) · [Português (Brasil)](i18n/README.pt-BR.md) · [Italiano](i18n/README.it.md) · [Nederlands](i18n/README.nl.md) · [Svenska](i18n/README.sv.md) · [Dansk](i18n/README.da.md) · [Norsk](i18n/README.no.md) · [Suomi](i18n/README.fi.md) · [Bahasa Indonesia](i18n/README.id.md) · [ไทย](i18n/README.th.md) · [Tiếng Việt](i18n/README.vi.md)

<p align="center"><b>The open-source editors of MR Office — documents, spreadsheets, presentations, PDF, handwriting and code — running in the browser.</b></p>

<p align="center"><a href="https://mroffice.ai/en"><b>▶ Try the live demo — no install</b></a> · <a href="https://github.com/XinaTian/MROffice/releases">⬇ Download MR Office from Releases</a> · <a href="https://mroffice.ai">🌐 mroffice.ai</a></p>

<p align="center"><a href="LICENSE"><img alt="AGPL-3.0" src="https://img.shields.io/badge/license-AGPL--3.0-blue"></a> <a href="https://github.com/XinaTian/MROffice/releases"><img alt="release" src="https://img.shields.io/github/v/release/XinaTian/MROffice"></a> <a href="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml/badge.svg"></a></p>

<table>
<tr>
<td align="center" width="33%"><a href="https://mroffice.ai/en"><img src="docs/screenshots/docs.png" alt="Documents" width="100%"></a><br><b>Documents</b></td>
<td align="center" width="33%"><a href="https://mroffice.ai/en"><img src="docs/screenshots/sheets.png" alt="Spreadsheets" width="100%"></a><br><b>Spreadsheets</b></td>
<td align="center" width="33%"><a href="https://mroffice.ai/en"><img src="docs/screenshots/slides.png" alt="Presentations" width="100%"></a><br><b>Presentations</b></td>
</tr>
<tr>
<td align="center" width="33%"><a href="https://mroffice.ai/en"><img src="docs/screenshots/pdf.png" alt="PDF" width="100%"></a><br><b>PDF</b></td>
<td align="center" width="33%"><a href="https://mroffice.ai/en"><img src="docs/screenshots/ink.png" alt="Handwriting" width="100%"></a><br><b>Handwriting</b></td>
<td align="center" width="33%"><a href="https://mroffice.ai/en"><img src="docs/screenshots/code.png" alt="Code" width="100%"></a><br><b>Code</b></td>
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

Calculating formulas with functions, drawing charts and laying out SmartArt run in MR Office's own engine, in the MR Office desktop app: the open-source editors keep the formulas, charts and SmartArt a file holds, and show SmartArt as it was saved.

## Run the editors

**▶ [Try the live demo — no install](https://mroffice.ai/en)**

The six editors are in `packages/` (`@mrnote/docs`, `sheets`, `slides`, `pdf`, `ink`, `code`, sharing `@mrnote/core`) and run together in `apps/playground`:

```bash
npm install
npm run fonts     # the open substitutes for Office's fonts
npm run build
npm run preview   # http://localhost:5178
```

`npm run dev` starts a development server at the same address, and `npm run smoke` checks a build. Each package's README lists what it does in detail and shows how to embed it.

## For the full experience, install the MR Office desktop app (free)

The MR Office desktop app is free, has these editors and the full service, and works offline on your own computer. It adds what the open-source editors leave out:

- Opening and saving Word, Excel and PowerPoint files (.docx, .xlsx, .pptx)
- Formulas with functions (SUM, IF, VLOOKUP …), calculated
- Charts
- SmartArt editing
- Saving PDF edits back into the file

**[Download MR Office from Releases →](https://github.com/XinaTian/MROffice/releases)** · [or from mroffice.ai](https://mroffice.ai)

### Word, Excel, PowerPoint and PDF files

Opening and saving Word, Excel and PowerPoint files (.docx, .xlsx, .pptx), and saving PDF edits back into the file, are provided by the MR Office desktop app (free), which has all of these editors and works offline on your own computer:

## License

Open source under the GNU Affero General Public License v3.0 (LICENSE), with the additional terms in NOTICE.md: keep the "MR Office" attribution and link to https://mroffice.ai, and no rights to the MR Office name or logo. A commercial license is available from support@mr-note.com. Third-party software and data are listed in THIRD_PARTY_NOTICES.md. Contributions: see CONTRIBUTING.md.

## More

Website and the full product: **https://mroffice.ai** · Questions and feedback: support@mr-note.com
