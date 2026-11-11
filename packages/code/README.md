# @mrnote/code

> **For the full experience, install the MR Office desktop app (free): https://github.com/XinaTian/MROffice/releases** (or https://mroffice.ai) — it adds opening and saving Word, Excel and PowerPoint files, formulas with functions, charts, SmartArt editing, and saving PDF edits back into the file.

The code editor of [MR Office](https://mroffice.ai): CodeMirror 6 with MR Office's ribbon, syntax highlighting for many languages, an outline, find and replace, formatting, problems, and code that runs right in the browser.

This package is generated from MR Office's own source (see the repository README, "How this repository is made").

## Use

```tsx
import { EditorHostProvider } from "@mrnote/core";
import codeEditor from "@mrnote/code";

const { Component: CodeEditor } = codeEditor;

<EditorHostProvider locale="en">
  <CodeEditor
    file={{ id: "1", name: "hello.py", bytes: new TextEncoder().encode('print("hi")') }}
    registerSave={(save) => { saver = save; return () => {}; }}
    onSaveRequest={() => download(saver())}
  />
</EditorHostProvider>
```

`registerSave` receives a function that returns the current text as a file (`{ name, type, bytes }`); `onSaveRequest` fires on the editor's own Save (ribbon, Ctrl/Cmd+S).

## What works in the open-source build

| Feature | Status |
|---|---|
| Open and edit text and source files (.js .ts .py .sql .lua .sh .rb .go .c .cpp .cs .java .rs .tex .json .yaml .html .css .md and more) | Available |
| Syntax highlighting, folding, bracket matching, multiple cursors, outline, breadcrumbs | Available |
| Find / replace, go to line, quick open (over the files open in this page) | Available |
| Format document (JavaScript/TypeScript/JSON/CSS/HTML/Markdown/YAML via Prettier, SQL, Python) | Available |
| Problems (JavaScript/TypeScript diagnostics, Python lint) and inlay hints | Available |
| Run JavaScript and TypeScript (sandboxed in the browser) | Available, verified |
| Run Python (Pyodide, loaded from the jsDelivr CDN on first run) | Available, verified (needs internet the first time) |
| Run SQL (SQLite, sql.js), Lua and shell scripts (WebAssembly, served by the page) | Available, verified |
| Run Ruby, C#, Go | The editor offers them, but their WebAssembly runtimes are not wired into the playground |
| Run C, C++, Java, Rust, compile LaTeX | Need MR Office's execution service; not available in the open-source build |
| Save | The host saves: the playground downloads the file |
| AI (Copilot, explain and fix), version history, collaboration, notebooks | Part of the MR Office service; not included |

## Size

About 115,000 lines of the editor's own code (packages/code/src/web), on top of the shared layer in @mrnote/core. In the playground's production build the code editor is a lazily loaded chunk; the in-browser runtimes (Pyodide, sql.js, the shell and Lua WebAssembly) load only when a file of that language is run.
