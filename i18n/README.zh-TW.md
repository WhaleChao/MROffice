# MR Office — 繁體中文

<p align="center"><sub>🌐 <a href="../README.md">English</a> · <a href="../i18n/README.en-GB.md">English (UK)</a> · <a href="../i18n/README.zh.md">简体中文</a> · <a href="../i18n/README.zh-TW.md">繁體中文</a> · <a href="../i18n/README.ja.md">日本語</a> · <a href="../i18n/README.ko.md">한국어</a> · <a href="../i18n/README.de.md">Deutsch</a> · <a href="../i18n/README.fr.md">Français</a> · <a href="../i18n/README.es.md">Español</a> · <a href="../i18n/README.es-419.md">Español (Latinoamérica)</a> · <a href="../i18n/README.pt-BR.md">Português (Brasil)</a> · <a href="../i18n/README.it.md">Italiano</a> · <a href="../i18n/README.nl.md">Nederlands</a> · <a href="../i18n/README.sv.md">Svenska</a> · <a href="../i18n/README.da.md">Dansk</a> · <a href="../i18n/README.no.md">Norsk</a> · <a href="../i18n/README.fi.md">Suomi</a> · <a href="../i18n/README.id.md">Bahasa Indonesia</a> · <a href="../i18n/README.th.md">ไทย</a> · <a href="../i18n/README.vi.md">Tiếng Việt</a></sub></p>

<p align="center"><b>MR Office 的開源編輯器：文件、試算表、簡報、PDF、手寫和程式碼，全部在瀏覽器中執行。</b></p>

<p align="center"><a href="https://mroffice.ai/zh-TW"><b>▶ 線上試用</b></a> · <a href="https://mroffice.ai/zh-TW/office/download"><b>⬇ 下載</b></a> · <a href="https://mroffice.ai">🌐 mroffice.ai</a><br><a href="https://mroffice.ai/zh-TW/developers">🧑‍💻 開發人員</a> · <a href="https://mroffice.ai/zh-TW/developers/embed">🧩 嵌入</a> · <a href="https://mroffice.ai/zh-TW/developers/mcp">🤖 AI 應用程式的 MCP</a> · <a href="https://www.npmjs.com/package/mroffice-embed">📦 npm SDK</a> · <a href="https://mroffice.ai/zh-TW/developers/oauth">🔑 使用 MR Office 登入</a></p>

<p align="center"><a href="../LICENSE"><img alt="AGPL-3.0" src="https://img.shields.io/badge/license-AGPL--3.0-blue"></a> <a href="https://github.com/XinaTian/MROffice/releases"><img alt="release" src="https://img.shields.io/github/v/release/XinaTian/MROffice"></a> <a href="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml/badge.svg"></a> <a href="https://github.com/XinaTian/MROffice/stargazers"><img alt="GitHub stars" src="https://img.shields.io/github/stars/XinaTian/MROffice?style=social"></a></p>

<p align="center"><a href="https://mroffice.ai/zh-TW"><img src="../docs/media/mr-office-demo.webp" alt="MR Office 的開源編輯器：文件、試算表、簡報、PDF、手寫和程式碼，全部在瀏覽器中執行。" width="100%"></a></p>

<p align="center"><a href="https://demo.mroffice.ai/media/mr-office-demo.mp4">🔊 觀看 50 秒介紹影片（有聲）</a></p>

<table>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/zh-TW"><img src="../docs/screenshots/docs.png" alt="文件" width="100%"></a><br><b>文件</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/zh-TW"><img src="../docs/screenshots/sheets.png" alt="試算表" width="100%"></a><br><b>試算表</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/zh-TW"><img src="../docs/screenshots/slides.png" alt="簡報" width="100%"></a><br><b>簡報</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/zh-TW"><img src="../docs/screenshots/pdf.png" alt="PDF" width="100%"></a><br><b>PDF</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/zh-TW"><img src="../docs/screenshots/ink.png" alt="手寫" width="100%"></a><br><b>手寫</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/zh-TW"><img src="../docs/screenshots/code.png" alt="程式碼" width="100%"></a><br><b>程式碼</b></td>
</tr>
</table>

MR Office 是一個結合筆記、AI 與真實檔案的工作空間。本儲存庫公開它的編輯器原始碼：與 MR Office 中執行的是同一套編輯器，你可以自行建置、執行與嵌入。這是開源編輯器，不是一套完整開源的 Office 套件：Microsoft Office 檔案的讀寫由 MR Office 提供（見下文）。

## 開發人員

開發人員需要的一切都在 **[MR Office Developers](https://mroffice.ai/zh-TW/developers)**——自助開通，無需申請：用任何 MR Office 帳號登入即可開始。

### 把編輯器嵌入你的網頁應用程式（建議）

編輯器由 MR Office 託管；你的伺服器簽發短期工作階段權杖，你的網頁用它開啟編輯器。它編輯真正的 Word、Excel、PowerPoint 和 PDF 檔案（.docx、.xlsx、.pptx、.pdf），支援多人共同編輯、上傳、所有匯出格式、列印，並透過簽署的回呼把儲存的檔案交回你的伺服器。免費：本機開發每月 1,000 次，正式網域每月 100 次，最多 2 個應用程式。這是託管的商業服務，你自己的應用程式不受本儲存庫 AGPL 授權約束。

1. 在你的專案資料夾執行 `npx mroffice-embed init`（或在開發人員主控台建立應用程式）：在瀏覽器按一次「允許」，應用程式的金鑰就會寫入 `.env.local` / `.env`。
2. 你的伺服器用 `mroffice-embed/server` 的 `createSessionToken` 簽發工作階段權杖；秘密金鑰只留在伺服器上。
3. 你的網頁用 `mroffice-embed` 的 `open()` 開啟編輯器（或用指令碼 `https://mroffice.ai/embed/v1/embed.js`）。

```js
// server: GET /api/mroffice/session (check who is signed in first)
import { createSessionToken } from "mroffice-embed/server";
const token = await createSessionToken({
  appId: process.env.MROFFICE_APP_ID,
  secretKey: process.env.MROFFICE_SECRET_KEY, // mro_sk_..., server only
  file: { key: "contract-42", name: "contract.docx", url: signedUrlToTheFile },
  user: { id: user.id, name: user.name },
  mode: "edit",
});

// page
import { open } from "mroffice-embed";
open({ container: "#editor", token: () => fetch("/api/mroffice/session").then((r) => r.json()).then((b) => b.token) });
```

[嵌入指南](https://mroffice.ai/zh-TW/developers/embed) · [Markdown 版指南（給 AI 程式設計助理）](https://mroffice.ai/developers/embed.md) · [價格與限制](https://mroffice.ai/zh-TW/developers/pricing) · [開發人員主控台](https://mroffice.ai/zh-TW/developers/console) · [線上展示](https://mroffice.ai/zh-TW/developers/embed/demo)

### AI 應用程式：MCP 伺服器

AI 應用程式可透過 MCP 伺服器在真正的 MR Office 編輯器中開啟、讀取、編輯和匯出文件：連線到託管端點 `https://mcp.mroffice.ai/mcp`（用 MR Office 帳號登入，不需金鑰），或在自己的電腦執行 `npx -y mroffice-mcp`。原始碼在 `apps/mcp`。 [MCP 指南 →](https://mroffice.ai/zh-TW/developers/mcp)

### 使用 MR Office 帳號登入（OAuth）

MR Office 是 OAuth 2.1 授權伺服器（PKCE、用戶端 ID 中繼資料文件、動態註冊）：你的應用程式可以讓使用者用 MR Office 帳號登入，並在其允許的範圍內操作。 [OAuth 指南 →](https://mroffice.ai/zh-TW/developers/oauth)

### 自行架設開放原始碼編輯器

本儲存庫的六個編輯器可依 AGPL-3.0 整合到你自己的產品：若他人透過網路使用你修改過的版本，你必須提供完整原始碼，並保留「MR Office」標示（NOTICE.md）。它們開啟自有格式、CSV、HTML、Markdown 和 PDF；Word、Excel 和 PowerPoint 檔案不在開放原始碼編輯器中（上面的託管嵌入有支援）。商業授權請洽 support@mr-note.com。 [開放原始碼指南 →](https://mroffice.ai/zh-TW/developers/open-source)

編輯器已作為 React 元件發佈到 npm：`npm i mroffice-docs mroffice-core`（另有 `mroffice-sheets`, `mroffice-slides`, `mroffice-pdf`, `mroffice-ink`, `mroffice-code`），每個套件的 README 寫了嵌入方法。

## 編輯器一覽

各編輯器在本儲存庫中能做的事：

- **文件** — 分頁文書處理器：樣式、字型、清單、表格、頁首頁尾、分節與分欄、圖形和方程式、封面和目錄、註解與追蹤修訂，以及在頁面上手寫。儲存為自有格式（.mrdoc）；可開啟 HTML、Markdown 和純文字。
- **試算表** — 儲存格與公式、數值格式、設定格式化的條件、排序篩選、凍結窗格和多工作表，並顯示選取範圍的加總與平均。儲存格之間的四則運算（如 =A1*B1+10）在瀏覽器中隨輸入計算；含函數的公式（SUM、IF、VLOOKUP 等）依原樣保留。儲存為自有格式（.mrsheet）；可開啟和儲存 CSV、TSV。
- **簡報** — 具 8 個自有佈景主題的投影片：版面配置、投影片母片與版面配置檢視、文字、圖形、文字方塊、表格、文字藝術師和方程式、註解、章節、轉場與動畫（可預覽），以及投影片放映。儲存為自有格式（.mrslides）。
- **PDF** — 閱讀、搜尋、放映和列印 PDF，複製和朗讀其中的文字；檔案開啟期間可對選取文字做螢光標示、底線、刪除線、波浪線標記，並加上便箋、使用註解窗格；儲存時交回開啟時的原檔案。
- **手寫** — 支援壓感的鋼筆、螢光筆、橡皮擦、套索和尺規，筆跡可轉換為圖形。儲存為自有格式（.mrink）。
- **程式碼** — 輕量程式碼編輯器，支援多種語言的語法醒目提示、大綱和尋找取代；JavaScript/TypeScript、Python、SQL、Lua 和 Shell 可直接在瀏覽器中執行。

帶函數的公式計算、圖表繪製和 SmartArt 排版由 MR Office 自己的引擎完成，在 MR Office 桌面版中提供：開源編輯器會保留檔案中的公式、圖表和 SmartArt，並按儲存時的樣子顯示 SmartArt。

## 執行編輯器

**▶ [線上試用，無需安裝](https://mroffice.ai/zh-TW)**

六個編輯器位於 `packages/`（`@mrnote/docs`、`sheets`、`slides`、`pdf`、`ink`、`code`，共用 `@mrnote/core`），在 `apps/playground` 中一起執行：

```bash
npm install
npm run fonts     # the open substitutes for Office's fonts
npm run build
npm run preview   # http://localhost:5178
```

`npm run dev` 在同一位址啟動開發伺服器，`npm run smoke` 檢查建置結果。各套件的 README 詳列了它能做的事和嵌入方式。

## 桌面應用程式

[`apps/desktop`](../apps/desktop) 是 **MR Office Editors**：MR Office 桌面版外殼（Electron）的原始碼，使用獨立的應用程式身分。裝上 Releases 中免費的 MR Office 元件後，它就是 MR Office 桌面版：可以開啟和儲存 Word、Excel、PowerPoint 和 PDF 檔案，可以登入並使用 AI。不裝元件時，它只執行開源編輯器，開啟 .mrdoc、.mrsheet、.mrslides、.mrink 檔案以及 CSV、HTML、Markdown 和 PDF。外殼和編輯器是 AGPL 原始碼；元件是編譯好的，可以免費使用，但不開源。先執行 `npm run build`，再執行：

```bash
npm run desktop
```

## 完整功能請安裝 MR Office 桌面版（免費）

MR Office 桌面版免費，包含這些編輯器和完整服務，在你自己的電腦上離線可用。開源編輯器沒有的功能，桌面版都提供：

- 開啟和儲存 Word、Excel、PowerPoint 檔案（.docx、.xlsx、.pptx）
- 帶函數的公式（SUM、IF、VLOOKUP 等）計算
- 圖表
- SmartArt 編輯
- PDF 編輯後寫回檔案

**[下載 MR Office（免費） →](https://mroffice.ai/zh-TW/office/download)**

### Word、Excel、PowerPoint 和 PDF 檔案

開啟和儲存 Word、Excel、PowerPoint 檔案（.docx、.xlsx、.pptx），以及把 PDF 的修改寫回檔案，由免費的 MR Office 桌面版提供。桌面版包含以上全部編輯器，在你自己的電腦上離線也能用：

## 各部分的關係

```mermaid
flowchart LR
  subgraph OSS["開源 · AGPL-3.0"]
    direction LR
    core["@mrnote/core"] --> E["六個編輯器<br/>文件 · 試算表 · 簡報 · PDF · 手寫 · 程式碼"]
    E --> PG["apps/playground<br/>瀏覽器"]
    E --> DS["apps/desktop<br/>桌面外殼"]
    MCP["apps/mcp<br/>給 AI 應用的 MCP 伺服器"]
  end
  COMP["MR Office 元件：Word、Excel、PowerPoint 檔案 · 免費，不開源"] -.-> DS
```

## 授權條款

依據 GNU Affero 通用公眾授權條款 v3.0（LICENSE）開源，並附加 NOTICE.md 中的條款：保留「MR Office」署名及指向 https://mroffice.ai 的連結，且不授予 MR Office 名稱與標誌的使用權。如需商業授權，請聯絡 support@mr-note.com。第三方軟體與資料見 THIRD_PARTY_NOTICES.md。貢獻程式碼請先閱讀 CONTRIBUTING.md。

## 更多

官網與完整產品：**https://mroffice.ai** · 問題與意見回饋：support@mr-note.com

<p align="center">如果 MR Office 對你有用，在 GitHub 上按個 ⭐，能讓更多人發現它。</p>
