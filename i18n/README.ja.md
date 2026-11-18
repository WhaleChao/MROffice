# MR Office — 日本語

<p align="center"><sub>🌐 <a href="../README.md">English</a> · <a href="../i18n/README.en-GB.md">English (UK)</a> · <a href="../i18n/README.zh.md">简体中文</a> · <a href="../i18n/README.zh-TW.md">繁體中文</a> · <a href="../i18n/README.ja.md">日本語</a> · <a href="../i18n/README.ko.md">한국어</a> · <a href="../i18n/README.de.md">Deutsch</a> · <a href="../i18n/README.fr.md">Français</a> · <a href="../i18n/README.es.md">Español</a> · <a href="../i18n/README.es-419.md">Español (Latinoamérica)</a> · <a href="../i18n/README.pt-BR.md">Português (Brasil)</a> · <a href="../i18n/README.it.md">Italiano</a> · <a href="../i18n/README.nl.md">Nederlands</a> · <a href="../i18n/README.sv.md">Svenska</a> · <a href="../i18n/README.da.md">Dansk</a> · <a href="../i18n/README.no.md">Norsk</a> · <a href="../i18n/README.fi.md">Suomi</a> · <a href="../i18n/README.id.md">Bahasa Indonesia</a> · <a href="../i18n/README.th.md">ไทย</a> · <a href="../i18n/README.vi.md">Tiếng Việt</a></sub></p>

<p align="center"><b>MR Office のオープンソースエディター。文書、スプレッドシート、プレゼンテーション、PDF、手書き、コードがブラウザーで動きます。</b></p>

<p align="center"><a href="https://mroffice.ai/ja"><b>▶ ライブデモを試す（インストール不要）</b></a> · <a href="https://github.com/XinaTian/MROffice/releases">⬇ Releases から MR Office をダウンロード</a> · <a href="https://mroffice.ai">🌐 mroffice.ai</a></p>

<p align="center"><a href="../LICENSE"><img alt="AGPL-3.0" src="https://img.shields.io/badge/license-AGPL--3.0-blue"></a> <a href="https://github.com/XinaTian/MROffice/releases"><img alt="release" src="https://img.shields.io/github/v/release/XinaTian/MROffice"></a> <a href="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml/badge.svg"></a> <a href="https://github.com/XinaTian/MROffice/stargazers"><img alt="GitHub stars" src="https://img.shields.io/github/stars/XinaTian/MROffice?style=social"></a></p>

<p align="center"><a href="https://mroffice.ai/ja"><img src="../docs/media/mr-office-demo.webp" alt="MR Office のオープンソースエディター。文書、スプレッドシート、プレゼンテーション、PDF、手書き、コードがブラウザーで動きます。" width="100%"></a></p>

<p align="center"><a href="../docs/media/mr-office-demo.mp4">🔊 50 秒の紹介動画を見る（音声あり）</a></p>

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

**[Releases から MR Office をダウンロード →](https://github.com/XinaTian/MROffice/releases)** · [または mroffice.ai から](https://mroffice.ai)

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
