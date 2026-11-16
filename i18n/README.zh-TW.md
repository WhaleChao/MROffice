# MR Office — 繁體中文

[English](../README.md) · [English (UK)](../i18n/README.en-GB.md) · [简体中文](../i18n/README.zh.md) · [繁體中文](../i18n/README.zh-TW.md) · [日本語](../i18n/README.ja.md) · [한국어](../i18n/README.ko.md) · [Deutsch](../i18n/README.de.md) · [Français](../i18n/README.fr.md) · [Español](../i18n/README.es.md) · [Español (Latinoamérica)](../i18n/README.es-419.md) · [Português (Brasil)](../i18n/README.pt-BR.md) · [Italiano](../i18n/README.it.md) · [Nederlands](../i18n/README.nl.md) · [Svenska](../i18n/README.sv.md) · [Dansk](../i18n/README.da.md) · [Norsk](../i18n/README.no.md) · [Suomi](../i18n/README.fi.md) · [Bahasa Indonesia](../i18n/README.id.md) · [ไทย](../i18n/README.th.md) · [Tiếng Việt](../i18n/README.vi.md)

<p align="center"><b>MR Office 的開源編輯器：文件、試算表、簡報、PDF、手寫和程式碼，全部在瀏覽器中執行。</b></p>

<p align="center"><a href="https://mroffice.ai/zh-TW"><b>▶ 線上試用，無需安裝</b></a> · <a href="https://github.com/XinaTian/MROffice/releases">⬇ 從 Releases 下載 MR Office</a> · <a href="https://mroffice.ai">🌐 mroffice.ai</a></p>

<p align="center"><a href="../LICENSE"><img alt="AGPL-3.0" src="https://img.shields.io/badge/license-AGPL--3.0-blue"></a> <a href="https://github.com/XinaTian/MROffice/releases"><img alt="release" src="https://img.shields.io/github/v/release/XinaTian/MROffice"></a> <a href="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml/badge.svg"></a></p>

<table>
<tr>
<td align="center" width="33%"><a href="https://mroffice.ai/zh-TW"><img src="../docs/screenshots/docs.png" alt="文件" width="100%"></a><br><b>文件</b></td>
<td align="center" width="33%"><a href="https://mroffice.ai/zh-TW"><img src="../docs/screenshots/sheets.png" alt="試算表" width="100%"></a><br><b>試算表</b></td>
<td align="center" width="33%"><a href="https://mroffice.ai/zh-TW"><img src="../docs/screenshots/slides.png" alt="簡報" width="100%"></a><br><b>簡報</b></td>
</tr>
<tr>
<td align="center" width="33%"><a href="https://mroffice.ai/zh-TW"><img src="../docs/screenshots/pdf.png" alt="PDF" width="100%"></a><br><b>PDF</b></td>
<td align="center" width="33%"><a href="https://mroffice.ai/zh-TW"><img src="../docs/screenshots/ink.png" alt="手寫" width="100%"></a><br><b>手寫</b></td>
<td align="center" width="33%"><a href="https://mroffice.ai/zh-TW"><img src="../docs/screenshots/code.png" alt="程式碼" width="100%"></a><br><b>程式碼</b></td>
</tr>
</table>

MR Office 是一個結合筆記、AI 與真實檔案的工作空間。本儲存庫公開它的編輯器原始碼：與 MR Office 中執行的是同一套編輯器，你可以自行建置、執行與嵌入。這是開源編輯器，不是一套完整開源的 Office 套件：Microsoft Office 檔案的讀寫由 MR Office 提供（見下文）。

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

## 完整功能請安裝 MR Office 桌面版（免費）

MR Office 桌面版免費，包含這些編輯器和完整服務，在你自己的電腦上離線可用。開源編輯器沒有的功能，桌面版都提供：

- 開啟和儲存 Word、Excel、PowerPoint 檔案（.docx、.xlsx、.pptx）
- 帶函數的公式（SUM、IF、VLOOKUP 等）計算
- 圖表
- SmartArt 編輯
- PDF 編輯後寫回檔案

**[從 Releases 下載 MR Office →](https://github.com/XinaTian/MROffice/releases)** · [或前往 mroffice.ai](https://mroffice.ai)

### Word、Excel、PowerPoint 和 PDF 檔案

開啟和儲存 Word、Excel、PowerPoint 檔案（.docx、.xlsx、.pptx），以及把 PDF 的修改寫回檔案，由免費的 MR Office 桌面版提供。桌面版包含以上全部編輯器，在你自己的電腦上離線也能用：

## 授權條款

依據 GNU Affero 通用公眾授權條款 v3.0（LICENSE）開源，並附加 NOTICE.md 中的條款：保留「MR Office」署名及指向 https://mroffice.ai 的連結，且不授予 MR Office 名稱與標誌的使用權。如需商業授權，請聯絡 support@mr-note.com。第三方軟體與資料見 THIRD_PARTY_NOTICES.md。貢獻程式碼請先閱讀 CONTRIBUTING.md。

## 更多

官網與完整產品：**https://mroffice.ai** · 問題與意見回饋：support@mr-note.com
