# MR Office — Bahasa Indonesia

<p align="center"><sub>🌐 <a href="../README.md">English</a> · <a href="../i18n/README.en-GB.md">English (UK)</a> · <a href="../i18n/README.zh.md">简体中文</a> · <a href="../i18n/README.zh-TW.md">繁體中文</a> · <a href="../i18n/README.ja.md">日本語</a> · <a href="../i18n/README.ko.md">한국어</a> · <a href="../i18n/README.de.md">Deutsch</a> · <a href="../i18n/README.fr.md">Français</a> · <a href="../i18n/README.es.md">Español</a> · <a href="../i18n/README.es-419.md">Español (Latinoamérica)</a> · <a href="../i18n/README.pt-BR.md">Português (Brasil)</a> · <a href="../i18n/README.it.md">Italiano</a> · <a href="../i18n/README.nl.md">Nederlands</a> · <a href="../i18n/README.sv.md">Svenska</a> · <a href="../i18n/README.da.md">Dansk</a> · <a href="../i18n/README.no.md">Norsk</a> · <a href="../i18n/README.fi.md">Suomi</a> · <a href="../i18n/README.id.md">Bahasa Indonesia</a> · <a href="../i18n/README.th.md">ไทย</a> · <a href="../i18n/README.vi.md">Tiếng Việt</a></sub></p>

<p align="center"><b>Editor sumber terbuka MR Office — dokumen, lembar kerja, presentasi, PDF, tulisan tangan, dan kode — yang berjalan di peramban.</b></p>

<p align="center"><a href="https://mroffice.ai/id"><b>▶ Coba online</b></a> · <a href="https://mroffice.ai/id/office/download"><b>⬇ Unduh</b></a> · <a href="https://mroffice.ai">🌐 mroffice.ai</a><br><a href="https://mroffice.ai/id/developers">🧑‍💻 Pengembang</a> · <a href="https://mroffice.ai/id/developers/embed">🧩 Sematkan</a> · <a href="https://mroffice.ai/id/developers/mcp">🤖 MCP untuk aplikasi AI</a> · <a href="https://www.npmjs.com/package/mroffice-embed">📦 SDK npm</a> · <a href="https://mroffice.ai/id/developers/oauth">🔑 Masuk dengan MR Office</a></p>

<p align="center"><a href="../LICENSE"><img alt="AGPL-3.0" src="https://img.shields.io/badge/license-AGPL--3.0-blue"></a> <a href="https://github.com/XinaTian/MROffice/releases"><img alt="release" src="https://img.shields.io/github/v/release/XinaTian/MROffice"></a> <a href="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml/badge.svg"></a> <a href="https://github.com/XinaTian/MROffice/stargazers"><img alt="GitHub stars" src="https://img.shields.io/github/stars/XinaTian/MROffice?style=social"></a></p>

<p align="center"><a href="https://mroffice.ai/id"><img src="../docs/media/mr-office-demo.webp" alt="Editor sumber terbuka MR Office — dokumen, lembar kerja, presentasi, PDF, tulisan tangan, dan kode — yang berjalan di peramban." width="100%"></a></p>

<p align="center"><a href="https://demo.mroffice.ai/media/mr-office-demo.mp4">🔊 Tonton tur 50 detik dengan suara</a></p>

<table>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/id"><img src="../docs/screenshots/docs.png" alt="Dokumen" width="100%"></a><br><b>Dokumen</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/id"><img src="../docs/screenshots/sheets.png" alt="Lembar kerja" width="100%"></a><br><b>Lembar kerja</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/id"><img src="../docs/screenshots/slides.png" alt="Presentasi" width="100%"></a><br><b>Presentasi</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/id"><img src="../docs/screenshots/pdf.png" alt="PDF" width="100%"></a><br><b>PDF</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/id"><img src="../docs/screenshots/ink.png" alt="Tulisan tangan" width="100%"></a><br><b>Tulisan tangan</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/id"><img src="../docs/screenshots/code.png" alt="Kode" width="100%"></a><br><b>Kode</b></td>
</tr>
</table>

MR Office adalah ruang kerja untuk catatan, AI, dan file sungguhan. Repositori ini menerbitkan kode sumber editornya: editor yang sama dengan yang dijalankan MR Office, yang bisa Anda bangun, jalankan, dan sematkan sendiri. Ini adalah editor sumber terbuka, bukan paket perkantoran sumber terbuka yang lengkap: membaca dan menulis file Microsoft Office tetap ada di MR Office (lihat di bawah).

## Untuk pengembang

Semua untuk pengembang ada di **[MR Office Developers](https://mroffice.ai/id/developers)** — layanan mandiri, tanpa pengajuan: masuk dengan akun MR Office mana pun dan mulai.

### Sematkan editor di aplikasi web Anda (disarankan)

MR Office menghosting editor; server Anda menandatangani token sesi berumur pendek dan halaman Anda membuka editor dengannya. Editor ini mengedit file Word, Excel, PowerPoint, dan PDF yang sesungguhnya (.docx, .xlsx, .pptx, .pdf), dengan pengeditan bersama, unggahan, semua format ekspor, pencetakan, dan hasil simpan yang dikirim ke server Anda lewat callback bertanda tangan. Paket gratis: 1.000 sesi per bulan di localhost, 100 di domain Anda sendiri, hingga 2 aplikasi. Ini layanan komersial yang dihosting, jadi aplikasi Anda tidak terikat lisensi AGPL repositori ini.

1. Di folder proyek Anda, jalankan `npx mroffice-embed init` (atau buat aplikasi di konsol pengembang): satu klik "Izinkan" di browser, dan kunci aplikasi ditulis ke `.env.local` / `.env`.
2. Server Anda membuat token sesi dengan `createSessionToken` dari `mroffice-embed/server`; kunci rahasia tidak pernah meninggalkan server.
3. Halaman Anda membuka editor dengan `open()` dari `mroffice-embed` (atau skrip `https://mroffice.ai/embed/v1/embed.js`).

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

[Panduan penyematan](https://mroffice.ai/id/developers/embed) · [Panduan dalam Markdown (untuk agen AI)](https://mroffice.ai/developers/embed.md) · [Harga dan batas](https://mroffice.ai/id/developers/pricing) · [Konsol pengembang](https://mroffice.ai/id/developers/console) · [Demo langsung](https://mroffice.ai/id/developers/embed/demo)

### Aplikasi AI: server MCP

Aplikasi AI membuka, membaca, mengedit, dan mengekspor dokumen di editor MR Office yang sesungguhnya melalui server MCP-nya: hubungkan endpoint yang dihosting `https://mcp.mroffice.ai/mcp` (masuk dengan akun MR Office Anda, tanpa kunci), atau jalankan `npx -y mroffice-mcp` di komputer Anda. Kode sumbernya ada di `apps/mcp`. [Panduan MCP →](https://mroffice.ai/id/developers/mcp)

### Masuk dengan MR Office (OAuth)

MR Office adalah server otorisasi OAuth 2.1 (PKCE, dokumen metadata ID klien, pendaftaran dinamis): aplikasi Anda dapat memasukkan orang dengan akun MR Office mereka dan bertindak dalam cakupan yang mereka izinkan. [Panduan OAuth →](https://mroffice.ai/id/developers/oauth)

### Host sendiri editor sumber terbuka

Enam editor di repositori ini dapat dibangun ke dalam produk Anda di bawah AGPL-3.0: jika orang menggunakan versi yang diubah melalui jaringan, Anda wajib menawarkan kode sumber lengkapnya dan mempertahankan atribusi "MR Office" (NOTICE.md). Editor ini membuka format miliknya sendiri, CSV, HTML, Markdown, dan PDF; file Word, Excel, dan PowerPoint tidak termasuk dalam editor sumber terbuka (penyematan yang dihosting di atas mendukungnya). Lisensi komersial tersedia melalui support@mr-note.com. [Panduan sumber terbuka →](https://mroffice.ai/id/developers/open-source)

Editor tersedia di npm sebagai komponen React: `npm i mroffice-docs mroffice-core` (serta `mroffice-sheets`, `mroffice-slides`, `mroffice-pdf`, `mroffice-ink`, `mroffice-code`); README tiap paket menjelaskan cara menyematkannya.

## Editor

Yang dilakukan setiap editor di repositori ini:

- **Dokumen** — pengolah kata berhalaman: gaya, font, daftar, tabel, header dan footer, bagian dan kolom, bentuk dan persamaan, halaman sampul dan daftar isi, komentar dan lacak perubahan, serta tinta di halaman. Menyimpan dalam formatnya sendiri (.mrdoc); membuka HTML, Markdown, dan teks biasa.
- **Lembar kerja** — sel dan rumus, format angka, pemformatan bersyarat, pengurutan dan filter, panel beku, dan banyak lembar, dengan jumlah dan rata-rata pilihan. Aritmetika antarsel (=A1*B1+10) dihitung di peramban saat diketik; rumus dengan fungsi (SUM, IF, VLOOKUP …) disimpan seperti yang ditulis. Menyimpan dalam formatnya sendiri (.mrsheet); membuka dan menyimpan CSV serta TSV.
- **Presentasi** — slide dengan delapan tema miliknya sendiri, tata letak, tampilan master slide dan tata letak, teks, bentuk, kotak teks, tabel, WordArt dan persamaan, komentar, bagian, transisi dan animasi dengan pratinjau, serta tayangan slide. Menyimpan dalam formatnya sendiri (.mrslides).
- **PDF** — membaca, mencari, menyajikan, dan mencetak file PDF, menyalin dan membacakan teksnya, serta menandai teks terpilih (sorot, garis bawah, coret, bergelombang) dengan catatan dan panel komentar selama file terbuka; menyimpan mengembalikan file seperti saat dibuka.
- **Tulisan tangan** — catatan tulisan tangan dengan pena peka tekanan, stabilo, penghapus, laso, dan penggaris, serta tinta yang diubah menjadi bentuk. Menyimpan dalam formatnya sendiri (.mrink).
- **Kode** — editor kode ringan dengan penyorotan sintaks untuk banyak bahasa, kerangka, serta cari dan ganti; JavaScript/TypeScript, Python, SQL, Lua, dan shell berjalan langsung di peramban.

Penghitungan rumus dengan fungsi, penggambaran bagan, dan tata letak SmartArt dijalankan oleh mesin MR Office sendiri, di aplikasi desktop MR Office: editor sumber terbuka menyimpan rumus, bagan, dan SmartArt dalam file, dan menampilkan SmartArt sebagaimana disimpan.

## Menjalankan editor

**▶ [Coba demo langsung — tanpa instalasi](https://mroffice.ai/id)**

Keenam editor ada di `packages/` (`@mrnote/docs`, `sheets`, `slides`, `pdf`, `ink`, `code`, yang berbagi `@mrnote/core`) dan berjalan bersama di `apps/playground`:

```bash
npm install
npm run fonts     # the open substitutes for Office's fonts
npm run build
npm run preview   # http://localhost:5178
```

`npm run dev` memulai server pengembangan di alamat yang sama, dan `npm run smoke` memeriksa hasil build. README setiap paket menjelaskan secara rinci apa yang dilakukannya dan cara menyematkannya.

## Aplikasi desktop

[`apps/desktop`](../apps/desktop) berisi **MR Office Editors**: kode sumber cangkang aplikasi desktop MR Office (Electron), dengan identitas aplikasinya sendiri. Dengan komponen MR Office gratis dari Releases, ini adalah aplikasi desktop MR Office, dengan file Word, Excel, PowerPoint, dan PDF, masuk akun, dan AI. Tanpa komponen itu, hanya editor sumber terbuka yang berjalan; editor ini membuka file .mrdoc, .mrsheet, .mrslides, dan .mrink, serta CSV, HTML, Markdown, dan PDF. Cangkang dan editor adalah kode sumber AGPL; komponennya sudah dikompilasi, gratis dipakai, dan bukan sumber terbuka. Setelah `npm run build`, jalankan dengan:

```bash
npm run desktop
```

## Untuk semua fitur, pasang aplikasi desktop MR Office (gratis)

Aplikasi desktop MR Office gratis, berisi editor-editor ini dan layanan lengkap, serta berjalan offline di komputer Anda sendiri. Aplikasi ini menambahkan yang tidak ada di editor sumber terbuka:

- Membuka dan menyimpan file Word, Excel, dan PowerPoint (.docx, .xlsx, .pptx)
- Menghitung rumus dengan fungsi (SUM, IF, VLOOKUP …)
- Bagan
- Pengeditan SmartArt
- Menyimpan hasil edit PDF kembali ke file

**[Unduh MR Office (gratis) →](https://mroffice.ai/id/office/download)**

### File Word, Excel, PowerPoint, dan PDF

Membuka dan menyimpan file Word, Excel, dan PowerPoint (.docx, .xlsx, .pptx), serta menyimpan hasil edit PDF kembali ke file, disediakan oleh aplikasi desktop MR Office (gratis), yang berisi semua editor ini dan berjalan offline di komputer Anda sendiri:

## Bagaimana bagian-bagiannya terhubung

```mermaid
flowchart LR
  subgraph OSS["Sumber terbuka · AGPL-3.0"]
    direction LR
    core["@mrnote/core"] --> E["Enam editor<br/>Dokumen · Lembar kerja · Presentasi · PDF · Tulisan tangan · Kode"]
    E --> PG["apps/playground<br/>Peramban"]
    E --> DS["apps/desktop<br/>Shell desktop"]
    MCP["apps/mcp<br/>Server MCP untuk aplikasi AI"]
  end
  COMP["Komponen MR Office: file Word, Excel, dan PowerPoint · gratis, bukan sumber terbuka"] -.-> DS
```

## Lisensi

Sumber terbuka di bawah GNU Affero General Public License v3.0 (LICENSE), dengan ketentuan tambahan di NOTICE.md: pertahankan atribusi "MR Office" dan tautan ke https://mroffice.ai, dan tidak ada hak atas nama atau logo MR Office. Lisensi komersial tersedia melalui support@mr-note.com. Perangkat lunak dan data pihak ketiga tercantum di THIRD_PARTY_NOTICES.md. Kontribusi: lihat CONTRIBUTING.md.

## Selengkapnya

Situs web dan produk lengkap: **https://mroffice.ai** · Pertanyaan dan masukan: support@mr-note.com

<p align="center">Jika MR Office bermanfaat bagi Anda, ⭐ di GitHub membantu orang lain menemukannya.</p>
