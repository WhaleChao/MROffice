# MR Office — 日本語

<p align="center"><sub>🌐 <a href="../README.md">English</a> · <a href="../i18n/README.en-GB.md">English (UK)</a> · <a href="../i18n/README.zh.md">简体中文</a> · <a href="../i18n/README.zh-TW.md">繁體中文</a> · <a href="../i18n/README.ja.md">日本語</a> · <a href="../i18n/README.ko.md">한국어</a> · <a href="../i18n/README.de.md">Deutsch</a> · <a href="../i18n/README.fr.md">Français</a> · <a href="../i18n/README.es.md">Español</a> · <a href="../i18n/README.es-419.md">Español (Latinoamérica)</a> · <a href="../i18n/README.pt-BR.md">Português (Brasil)</a> · <a href="../i18n/README.it.md">Italiano</a> · <a href="../i18n/README.nl.md">Nederlands</a> · <a href="../i18n/README.sv.md">Svenska</a> · <a href="../i18n/README.da.md">Dansk</a> · <a href="../i18n/README.no.md">Norsk</a> · <a href="../i18n/README.fi.md">Suomi</a> · <a href="../i18n/README.id.md">Bahasa Indonesia</a> · <a href="../i18n/README.th.md">ไทย</a> · <a href="../i18n/README.vi.md">Tiếng Việt</a></sub></p>

<p align="center"><b>MR Office のオープンソースエディター。文書、スプレッドシート、プレゼンテーション、PDF、手書き、コードがブラウザーで動きます。</b></p>

<p align="center"><a href="https://mroffice.ai/ja"><b>▶ オンラインで試す</b></a> · <a href="https://mroffice.ai/ja/office/download"><b>⬇ ダウンロード</b></a> · <a href="https://mroffice.ai">🌐 mroffice.ai</a><br><a href="https://mroffice.ai/ja/developers">🧑‍💻 開発者</a> · <a href="https://mroffice.ai/ja/developers/embed">🧩 埋め込み</a> · <a href="https://mroffice.ai/ja/developers/mcp">🤖 AI アプリ向け MCP</a> · <a href="https://www.npmjs.com/package/mroffice-embed">📦 npm SDK</a> · <a href="https://mroffice.ai/ja/developers/oauth">🔑 MR Office でサインイン</a></p>

<p align="center"><a href="../LICENSE"><img alt="AGPL-3.0" src="https://img.shields.io/badge/license-AGPL--3.0-blue"></a> <a href="https://github.com/XinaTian/MROffice/releases"><img alt="release" src="https://img.shields.io/github/v/release/XinaTian/MROffice"></a> <a href="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml/badge.svg"></a> <a href="https://github.com/XinaTian/MROffice/stargazers"><img alt="GitHub stars" src="https://img.shields.io/github/stars/XinaTian/MROffice?style=social"></a></p>

<p align="center"><a href="https://mroffice.ai/ja"><img src="../docs/media/mr-office-demo.webp" alt="MR Office のオープンソースエディター。文書、スプレッドシート、プレゼンテーション、PDF、手書き、コードがブラウザーで動きます。" width="100%"></a></p>

<p align="center"><a href="https://demo.mroffice.ai/media/mr-office-demo.mp4">🔊 50 秒の紹介動画を見る（音声あり）</a></p>

<table>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/ja"><img src="../docs/screenshots/docs.png" alt="文書" width="100%"></a><br><b>文書</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/ja"><img src="../docs/screenshots/sheets.png" alt="スプレッドシート" width="100%"></a><br><b>スプレッドシート</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/ja"><img src="../docs/screenshots/slides.png" alt="プレゼンテーション" width="100%"></a><br><b>プレゼンテーション</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/ja"><img src="../docs/screenshots/pdf.png" alt="PDF" width="100%"></a><br><b>PDF</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/ja"><img src="../docs/screenshots/ink.png" alt="手書き" width="100%"></a><br><b>手書き</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/ja"><img src="../docs/screenshots/code.png" alt="コード" width="100%"></a><br><b>コード</b></td>
</tr>
</table>

MR Office は、ノート、AI、実際のファイルをひとつにまとめたワークスペースです。このリポジトリでは、そのエディターのソースコードを公開しています。MR Office で動いているものと同じエディターで、自分でビルド、実行、組み込みができます。これはオープンソースのエディターであり、完全なオープンソースの Office スイートではありません。Microsoft Office ファイルの読み書きは MR Office が提供します（下記参照）。

## 開発者向け

開発者向けのすべては **[MR Office Developers](https://mroffice.ai/ja/developers)** にあります。セルフサービスで申請は不要です。MR Office アカウントでサインインすればすぐ始められます。

### Web アプリにエディターを埋め込む（推奨）

エディターは MR Office がホストします。サーバーが短期間有効なセッショントークンに署名し、ページがそれでエディターを開きます。本物の Word、Excel、PowerPoint、PDF ファイル（.docx、.xlsx、.pptx、.pdf）を編集でき、共同編集、アップロード、すべてのエクスポート形式、印刷に対応し、保存したファイルは署名付きコールバックでサーバーに返されます。無料プラン：localhost での開発は月 1,000 セッション、本番ドメインは月 100 セッション、アプリは最大 2 個。ホスト型の商用サービスなので、あなたのアプリはこのリポジトリの AGPL ライセンスに縛られません。

1. プロジェクトのフォルダーで `npx mroffice-embed init` を実行します（または開発者コンソールでアプリを作成）。ブラウザーで「許可」を 1 回押すと、アプリのキーが `.env.local` / `.env` に書き込まれます。
2. サーバーは `mroffice-embed/server` の `createSessionToken` でセッショントークンを発行します。シークレットキーはサーバーの外に出ません。
3. ページは `mroffice-embed` の `open()`（またはスクリプト `https://mroffice.ai/embed/v1/embed.js`）でエディターを開きます。

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

[埋め込みガイド](https://mroffice.ai/ja/developers/embed) · [Markdown 版ガイド（AI エージェント向け）](https://mroffice.ai/developers/embed.md) · [料金と上限](https://mroffice.ai/ja/developers/pricing) · [開発者コンソール](https://mroffice.ai/ja/developers/console) · [ライブデモ](https://mroffice.ai/ja/developers/embed/demo)

### AI アプリ: MCP サーバー

AI アプリは MCP サーバーを通じて、本物の MR Office エディターでドキュメントを開き、読み、編集し、エクスポートできます。ホスト型エンドポイント `https://mcp.mroffice.ai/mcp`（MR Office アカウントでサインイン、キー不要）に接続するか、自分のコンピューターで `npx -y mroffice-mcp` を実行します。ソースは `apps/mcp` にあります。 [MCP ガイド →](https://mroffice.ai/ja/developers/mcp)

### MR Office でサインイン（OAuth）

MR Office は OAuth 2.1 認可サーバーです（PKCE、クライアント ID メタデータドキュメント、動的登録）。あなたのアプリは MR Office アカウントでユーザーをサインインさせ、許可されたスコープで操作できます。 [OAuth ガイド →](https://mroffice.ai/ja/developers/oauth)

### オープンソースのエディターをセルフホストする

このリポジトリの 6 つのエディターは AGPL-3.0 のもとで自社製品に組み込めます。改変版をネットワーク越しに利用させる場合は完全なソースコードを提供し、「MR Office」の表示（NOTICE.md）を残す必要があります。独自形式、CSV、HTML、Markdown、PDF を開けます。Word、Excel、PowerPoint ファイルはオープンソースのエディターには含まれません（上のホスト型埋め込みが対応）。商用ライセンスは support@mr-note.com へ。 [オープンソースガイド →](https://mroffice.ai/ja/developers/open-source)

エディターは React コンポーネントとして npm にあります：`npm i mroffice-docs mroffice-core`（ほかに `mroffice-sheets`, `mroffice-slides`, `mroffice-pdf`, `mroffice-ink`, `mroffice-code`）。埋め込み方は各パッケージの README にあります。

## エディター

このリポジトリで各エディターができること：

- **文書** — ページ単位のワープロ：スタイル、フォント、リスト、表、ヘッダーとフッター、セクションと段組み、図形と数式、表紙と目次、コメントと変更履歴、ページへの手書き。独自形式（.mrdoc）で保存し、HTML、Markdown、テキストを開けます。
- **スプレッドシート** — セルと数式、表示形式、条件付き書式、並べ替えとフィルター、ウィンドウ枠の固定、複数シート、選択範囲の合計と平均。セル同士の四則演算（=A1*B1+10 など）は入力するとブラウザーで計算され、関数を使う数式（SUM、IF、VLOOKUP など）は書いたまま保持されます。独自形式（.mrsheet）で保存し、CSV と TSV を開いて保存できます。
- **プレゼンテーション** — 独自の 8 つのテーマを備えたスライド：レイアウト、スライド マスターとレイアウト表示、テキスト、図形、テキスト ボックス、表、ワードアート、数式、コメント、セクション、画面切り替えとアニメーション（プレビュー付き）、スライドショー。独自形式（.mrslides）で保存します。
- **PDF** — PDF の閲覧、検索、プレゼンテーション表示、印刷、テキストのコピーと読み上げ。開いている間は選択したテキストにハイライト、下線、取り消し線、波線を付け、メモとコメント パネルを使えます。保存すると開いたときのファイルを返します。
- **手書き** — 筆圧対応のペン、蛍光ペン、消しゴム、投げ縄、ルーラー。インクを図形に変換できます。独自形式（.mrink）で保存します。
- **コード** — 多くの言語の構文ハイライト、アウトライン、検索と置換を備えた軽量コードエディター。JavaScript/TypeScript、Python、SQL、Lua、シェルをブラウザーで実行できます。

関数を使う数式の計算、グラフの描画、SmartArt のレイアウトは MR Office 独自のエンジンで行われ、MR Office デスクトップ版で利用できます。オープンソースのエディターはファイル内の数式・グラフ・SmartArt を保持し、SmartArt は保存時の見た目で表示します。

## エディターを動かす

**▶ [ライブデモを試す（インストール不要）](https://mroffice.ai/ja)**

6 つのエディターは `packages/`（`@mrnote/docs`、`sheets`、`slides`、`pdf`、`ink`、`code`。共通部分は `@mrnote/core`）にあり、`apps/playground` でまとめて動きます：

```bash
npm install
npm run fonts     # the open substitutes for Office's fonts
npm run build
npm run preview   # http://localhost:5178
```

`npm run dev` は同じアドレスで開発サーバーを起動し、`npm run smoke` はビルドを確認します。各パッケージの README に、できることの詳細と組み込み方を載せています。

## デスクトップアプリ

[`apps/desktop`](../apps/desktop) には **MR Office Editors** があります。MR Office デスクトップアプリのシェル（Electron）のソースコードで、独自のアプリ ID を持ちます。Releases にある無料の MR Office コンポーネントを入れると MR Office デスクトップアプリになり、Word、Excel、PowerPoint、PDF ファイル、サインイン、AI が使えます。コンポーネントがなければ、オープンソースのエディターだけで動き、.mrdoc、.mrsheet、.mrslides、.mrink ファイルと CSV、HTML、Markdown、PDF を開きます。シェルとエディターは AGPL のソースコードで、コンポーネントはコンパイル済みで無料で使えますが、オープンソースではありません。`npm run build` の後、次のコマンドで起動します:

```bash
npm run desktop
```

## すべての機能を使うには MR Office デスクトップ版（無料）をインストールしてください

MR Office デスクトップ版は無料で、これらのエディターとすべてのサービスを備え、お使いのコンピューターでオフラインでも動作します。オープンソース版にない次の機能を利用できます：

- Word・Excel・PowerPoint ファイル（.docx、.xlsx、.pptx）を開いて保存
- 関数を使う数式（SUM、IF、VLOOKUP など）の計算
- グラフ
- SmartArt の編集
- PDF の編集内容をファイルに書き戻す

**[MR Office をダウンロード（無料） →](https://mroffice.ai/ja/office/download)**

### Word、Excel、PowerPoint、PDF ファイル

Word・Excel・PowerPoint ファイル（.docx、.xlsx、.pptx）を開いて保存する機能と、PDF の編集内容をファイルに書き戻す機能は、無料の MR Office デスクトップ版が提供します。デスクトップ版にはこれらのエディターがすべて含まれ、お使いのコンピューターでオフラインでも動作します：

## 構成

```mermaid
flowchart LR
  subgraph OSS["オープンソース · AGPL-3.0"]
    direction LR
    core["@mrnote/core"] --> E["6 つのエディター<br/>文書 · スプレッドシート · プレゼンテーション · PDF · 手書き · コード"]
    E --> PG["apps/playground<br/>ブラウザー"]
    E --> DS["apps/desktop<br/>デスクトップシェル"]
    MCP["apps/mcp<br/>AI アプリ向け MCP サーバー"]
  end
  COMP["MR Office コンポーネント：Word・Excel・PowerPoint ファイル · 無料、非オープンソース"] -.-> DS
```

## ライセンス

GNU Affero General Public License v3.0（LICENSE）のもとで、NOTICE.md の追加条項とともに公開しています。「MR Office」の表示と https://mroffice.ai へのリンクを残すこと、MR Office の名称とロゴの権利は付与されないことが条件です。商用ライセンスは support@mr-note.com までお問い合わせください。サードパーティのソフトウェアとデータは THIRD_PARTY_NOTICES.md に記載しています。コントリビューションについては CONTRIBUTING.md をご覧ください。

## 詳しく

Web サイトと製品版：**https://mroffice.ai** · ご質問・ご意見：support@mr-note.com

<p align="center">MR Office が役に立ったら、GitHub で ⭐ を付けてください。ほかの人が見つけやすくなります。</p>
