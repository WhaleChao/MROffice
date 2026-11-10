# MR Office — 日本語

[English](../README.md) · [English (UK)](../i18n/README.en-GB.md) · [简体中文](../i18n/README.zh.md) · [繁體中文](../i18n/README.zh-TW.md) · [日本語](../i18n/README.ja.md) · [한국어](../i18n/README.ko.md) · [Deutsch](../i18n/README.de.md) · [Français](../i18n/README.fr.md) · [Español](../i18n/README.es.md) · [Español (Latinoamérica)](../i18n/README.es-419.md) · [Português (Brasil)](../i18n/README.pt-BR.md) · [Italiano](../i18n/README.it.md) · [Nederlands](../i18n/README.nl.md) · [Svenska](../i18n/README.sv.md) · [Dansk](../i18n/README.da.md) · [Norsk](../i18n/README.no.md) · [Suomi](../i18n/README.fi.md) · [Bahasa Indonesia](../i18n/README.id.md) · [ไทย](../i18n/README.th.md) · [Tiếng Việt](../i18n/README.vi.md)

> ### ⬇ すべての機能を使うには MR Office デスクトップ版（無料）をインストールしてください
>
> MR Office デスクトップ版は無料で、これらのエディターとすべてのサービスを備え、お使いのコンピューターでオフラインでも動作します。オープンソース版にない次の機能を利用できます：
>
> - Word・Excel・PowerPoint ファイル（.docx、.xlsx、.pptx）を開いて保存
> - 関数を使う数式（SUM、IF、VLOOKUP など）の計算
> - グラフ
> - SmartArt の編集
> - PDF の編集内容をファイルに書き戻す
>
> **[Releases から MR Office をダウンロード →](https://github.com/dongmingyang2007-pixel/MROffice/releases)** · [または mroffice.ai から](https://mroffice.ai)

<p align="center"><a href="https://mroffice.ai"><b>🌐 mroffice.ai</b></a> · <a href="https://mroffice.ai">MR Office で試す →</a></p>

> MR Office のオープンソースエディター。文書、スプレッドシート、プレゼンテーション、PDF、手書き、コードがブラウザーで動きます。

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

## Word、Excel、PowerPoint、PDF ファイル

Word・Excel・PowerPoint ファイル（.docx、.xlsx、.pptx）を開いて保存する機能と、PDF の編集内容をファイルに書き戻す機能は、無料の MR Office デスクトップ版が提供します。デスクトップ版にはこれらのエディターがすべて含まれ、お使いのコンピューターでオフラインでも動作します：

- **MR Office** (無料のデスクトップアプリ): https://github.com/dongmingyang2007-pixel/MROffice/releases
- **公式サイト**: https://mroffice.ai

[すべてのバージョン →](https://github.com/dongmingyang2007-pixel/MROffice/releases)

## エディターを動かす

6 つのエディターは `packages/`（`@mrnote/docs`、`sheets`、`slides`、`pdf`、`ink`、`code`。共通部分は `@mrnote/core`）にあり、`apps/playground` でまとめて動きます：

```bash
npm install
npm run fonts     # the open substitutes for Office's fonts
npm run build
npm run preview   # http://localhost:5178
```

`npm run dev` は同じアドレスで開発サーバーを起動し、`npm run smoke` はビルドを確認します。各パッケージの README に、できることの詳細と組み込み方を載せています。

## ライセンス

GNU Affero General Public License v3.0（LICENSE）のもとで、NOTICE.md の追加条項とともに公開しています。「MR Office」の表示と https://mroffice.ai へのリンクを残すこと、MR Office の名称とロゴの権利は付与されないことが条件です。商用ライセンスは support@mr-note.com までお問い合わせください。サードパーティのソフトウェアとデータは THIRD_PARTY_NOTICES.md に記載しています。コントリビューションについては CONTRIBUTING.md をご覧ください。

## 詳しく

Web サイトと製品版：**https://mroffice.ai** · ご質問・ご意見：support@mr-note.com
