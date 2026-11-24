# Changelog

## Unreleased

- Spreadsheet formulas recalculate in the editor view, as the signed-in
  person: the view's calculation calls go through this server (the view-only
  tool `widget_calculate`, or `/_mcp/api/doc/<id>/calculate` for the link
  page) to MR Office's calculation engine with the person's own token
  (`office.convert`). On a server people sign in to, the link page calculates
  while one of the owner's AI app connections is open. Without MR Office, the
  view keeps showing the file's last results, as before. Needs the open
  editors built with the host bridge's `answerApiRequest`.
- Every tool states all three hints (`readOnlyHint`, `destructiveHint`,
  `openWorldHint`) and a title, the view's own tools included, as connector
  directories require; `render_page` is titled "Render a page as an image".
- `export_document` no longer points a signed-in server's model at a
  `sign_in` tool it does not have; `get_document` no longer reports the save
  time.
- `MROFFICE_MCP_VIEW_META`: a JSON object of extra `_meta` for the editor view
  (an app's own keys, such as the origin it serves views from).
- A server reached over the network serves `/.well-known/<name>` from
  `<data dir>/well-known/<name>`, so a directory's ownership token can be put
  in place without a restart.

## 0.2.0 — 2026-11-21

- **MR Office as your AI's screen.** With the MR Office desktop app (Mac,
  Windows) running on the same computer, stdio hands the AI app's connection
  to the app's own MCP endpoint: your own `.docx`, `.xlsx`, `.pptx` and `.pdf`
  files open in the app's window, every edit the AI makes shows there as it is
  made (one Undo step each, saved back to the file like your typing), and the
  AI reads your own edits back. The setup line is unchanged (the command
  `npx -y mroffice-mcp` in the AI app's MCP settings).
  - Found by the app's `ai-screen.json` (its data folder; loopback port and
    token, this user's only); a file left by an app that is gone is not used.
  - The desktop app asks the person once per AI app (Allow once / Always
    allow / Don't allow).
  - The app's tools: `open_document` (a path), `list_documents`,
    `get_document` (addresses: block ids, A1 cells, slide and shape ids),
    `get_selection`, `edit_document` (the Automation API's `word.*`, `excel.*`,
    `powerpoint.*`, `pdf.*` methods), `get_changes`, `render_page`,
    `save_document`, `export_document`.
  - A desktop app that restarts, or that ended the connection (its permission
    taken back in Settings), is connected to again.
- `--desktop`: always the desktop app (started in the background when it is
  not running); `--no-desktop`: never. Without either, the app is used when it
  runs or is installed, and the server's own documents otherwise.
  `MROFFICE_MCP_DESKTOP=1|0`, `MROFFICE_DESKTOP_DISCOVERY`,
  `MROFFICE_MCP_DESKTOP_LAUNCH=0`.
- Needs MR Office desktop 0.1.6 or newer (MRNote Classic 0.1.10 or newer) for
  the screen; older apps are simply not found.

## 0.1.0 — 2026-11-19

- First release: create, open, update, read, picture and export documents,
  spreadsheets and PDFs in MR Office's editor, inline in apps that render MCP
  Apps views and by link elsewhere; hosted at https://mcp.mroffice.ai/mcp.
