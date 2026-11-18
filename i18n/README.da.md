# MR Office — Dansk

<p align="center"><sub>🌐 <a href="../README.md">English</a> · <a href="../i18n/README.en-GB.md">English (UK)</a> · <a href="../i18n/README.zh.md">简体中文</a> · <a href="../i18n/README.zh-TW.md">繁體中文</a> · <a href="../i18n/README.ja.md">日本語</a> · <a href="../i18n/README.ko.md">한국어</a> · <a href="../i18n/README.de.md">Deutsch</a> · <a href="../i18n/README.fr.md">Français</a> · <a href="../i18n/README.es.md">Español</a> · <a href="../i18n/README.es-419.md">Español (Latinoamérica)</a> · <a href="../i18n/README.pt-BR.md">Português (Brasil)</a> · <a href="../i18n/README.it.md">Italiano</a> · <a href="../i18n/README.nl.md">Nederlands</a> · <a href="../i18n/README.sv.md">Svenska</a> · <a href="../i18n/README.da.md">Dansk</a> · <a href="../i18n/README.no.md">Norsk</a> · <a href="../i18n/README.fi.md">Suomi</a> · <a href="../i18n/README.id.md">Bahasa Indonesia</a> · <a href="../i18n/README.th.md">ไทย</a> · <a href="../i18n/README.vi.md">Tiếng Việt</a></sub></p>

<p align="center"><b>MR Offices open source-editorer – dokumenter, regneark, præsentationer, PDF, håndskrift og kode – i browseren.</b></p>

<p align="center"><a href="https://mroffice.ai/da"><b>▶ Prøv live-demoen – ingen installation</b></a> · <a href="https://github.com/XinaTian/MROffice/releases">⬇ Download MR Office fra Releases</a> · <a href="https://mroffice.ai">🌐 mroffice.ai</a></p>

<p align="center"><a href="../LICENSE"><img alt="AGPL-3.0" src="https://img.shields.io/badge/license-AGPL--3.0-blue"></a> <a href="https://github.com/XinaTian/MROffice/releases"><img alt="release" src="https://img.shields.io/github/v/release/XinaTian/MROffice"></a> <a href="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml/badge.svg"></a> <a href="https://github.com/XinaTian/MROffice/stargazers"><img alt="GitHub stars" src="https://img.shields.io/github/stars/XinaTian/MROffice?style=social"></a></p>

<p align="center"><a href="https://mroffice.ai/da"><img src="../docs/media/mr-office-demo.webp" alt="MR Offices open source-editorer – dokumenter, regneark, præsentationer, PDF, håndskrift og kode – i browseren." width="100%"></a></p>

<p align="center"><a href="https://github.com/user-attachments/assets/3dbe2320-c587-497b-a554-73e469909387">🔊 Se den 50 sekunder lange rundvisning med lyd</a></p>

<table>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/da"><img src="../docs/screenshots/docs.png" alt="Dokumenter" width="100%"></a><br><b>Dokumenter</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/da"><img src="../docs/screenshots/sheets.png" alt="Regneark" width="100%"></a><br><b>Regneark</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/da"><img src="../docs/screenshots/slides.png" alt="Præsentationer" width="100%"></a><br><b>Præsentationer</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/da"><img src="../docs/screenshots/pdf.png" alt="PDF" width="100%"></a><br><b>PDF</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/da"><img src="../docs/screenshots/ink.png" alt="Håndskrift" width="100%"></a><br><b>Håndskrift</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/da"><img src="../docs/screenshots/code.png" alt="Kode" width="100%"></a><br><b>Kode</b></td>
</tr>
</table>

MR Office er et arbejdsområde til noter, AI og rigtige filer. Dette repository udgiver kildekoden til dets editorer: de samme editorer, som MR Office kører, og som du selv kan bygge, køre og indlejre. Det er open source-editorer, ikke en komplet open source-kontorpakke: læsning og skrivning af Microsoft Office-filer bliver i MR Office (se nedenfor).

## Editorerne

Hvad hver editor gør i dette repository:

- **Dokumenter** — et sideopdelt tekstbehandlingsprogram: typografier, skrifttyper, lister, tabeller, sidehoveder og sidefødder, sektioner og spalter, figurer og ligninger, forsider og indholdsfortegnelser, kommentarer og registrering af ændringer samt håndskrift på siden. Gemmer i eget format (.mrdoc); åbner HTML, Markdown og ren tekst.
- **Regneark** — celler og formler, talformater, betinget formatering, sortering og filtrering, frosne ruder og flere ark, med sum og gennemsnit af markeringen. Regning mellem celler (=A1*B1+10) beregnes i browseren, mens du skriver; formler med funktioner (SUM, HVIS, LOPSLAG …) bevares, som de er skrevet. Gemmer i eget format (.mrsheet); åbner og gemmer CSV og TSV.
- **Præsentationer** — dias med otte egne temaer, layout, diasmaster- og layoutvisning, tekst, figurer, tekstfelter, tabeller, WordArt og ligninger, kommentarer, sektioner, overgange og animationer med eksempel samt diasshow. Gemmer i eget format (.mrslides).
- **PDF** — læse, søge i, præsentere og udskrive PDF-filer, kopiere og læse teksten højt og markere valgt tekst (fremhævning, understregning, gennemstregning, bølgelinje) med noter og en kommentarrude, mens filen er åben; når du gemmer, gives filen tilbage, som den blev åbnet.
- **Håndskrift** — håndskrevne noter med trykfølsomme penne, overstregningstuscher, viskelædere, lasso og lineal, og håndskrift der bliver til figurer. Gemmer i eget format (.mrink).
- **Kode** — en let kodeeditor med syntaksfremhævning til mange sprog, en disposition samt søg og erstat; JavaScript/TypeScript, Python, SQL, Lua og shell kører direkte i browseren.

Beregning af formler med funktioner, tegning af diagrammer og opsætning af SmartArt klares af MR Offices egen motor i MR Offices skrivebordsapp: open source-editorerne bevarer en fils formler, diagrammer og SmartArt og viser SmartArt, som det blev gemt.

## Kør editorerne

**▶ [Prøv live-demoen – ingen installation](https://mroffice.ai/da)**

De seks editorer ligger i `packages/` (`@mrnote/docs`, `sheets`, `slides`, `pdf`, `ink`, `code`, som deler `@mrnote/core`) og kører sammen i `apps/playground`:

```bash
npm install
npm run fonts     # the open substitutes for Office's fonts
npm run build
npm run preview   # http://localhost:5178
```

`npm run dev` starter en udviklingsserver på samme adresse, og `npm run smoke` tjekker et build. Hver pakkes README beskriver i detaljer, hvad den gør, og hvordan den indlejres.

## Skrivebordsapp

[`apps/desktop`](../apps/desktop) indeholder **MR Office Editors**: kildekoden til MR Offices skrivebordsapps skal (Electron) med sin egen app-identitet. Med den gratis MR Office-komponent fra Releases er det MR Offices skrivebordsapp med Word-, Excel-, PowerPoint- og PDF-filer, login og AI. Uden komponenten kører kun editorerne med åben kildekode; de åbner .mrdoc-, .mrsheet-, .mrslides- og .mrink-filer samt CSV, HTML, Markdown og PDF. Skallen og editorerne er AGPL-kildekode; komponenten er kompileret, gratis at bruge og ikke open source. Efter `npm run build` starter du den med:

```bash
npm run desktop
```

## Få alle funktioner: installér MR Offices skrivebordsapp (gratis)

MR Offices skrivebordsapp er gratis, indeholder disse editorer og hele tjenesten og virker offline på din egen computer. Den tilføjer det, som open source-editorerne mangler:

- Åbne og gemme Word-, Excel- og PowerPoint-filer (.docx, .xlsx, .pptx)
- Beregne formler med funktioner (SUM, HVIS, LOPSLAG …)
- Diagrammer
- Redigering af SmartArt
- Gemme ændringer i en PDF tilbage i filen

**[Download MR Office fra Releases →](https://github.com/XinaTian/MROffice/releases)** · [eller fra mroffice.ai](https://mroffice.ai)

### Word-, Excel-, PowerPoint- og PDF-filer

At åbne og gemme Word-, Excel- og PowerPoint-filer (.docx, .xlsx, .pptx) og gemme ændringer i en PDF tilbage i filen klares af MR Offices gratis skrivebordsapp. Den indeholder alle disse editorer og virker offline på din egen computer:

## Sådan hænger delene sammen

```mermaid
flowchart LR
  subgraph OSS["Open source · AGPL-3.0"]
    direction LR
    core["@mrnote/core"] --> E["De seks editorer<br/>Dokumenter · Regneark · Præsentationer · PDF · Håndskrift · Kode"]
    E --> PG["apps/playground<br/>Browser"]
    E --> DS["apps/desktop<br/>Skrivebordsskal"]
    MCP["apps/mcp<br/>MCP-server til AI-apps"]
  end
  COMP["MR Office-komponenten: Word-, Excel- og PowerPoint-filer · gratis, ikke open source"] -.-> DS
```

## Licens

Open source under GNU Affero General Public License v3.0 (LICENSE) med de yderligere vilkår i NOTICE.md: bevar angivelsen "MR Office" og linket til https://mroffice.ai, og ingen rettigheder til navnet eller logoet MR Office. En kommerciel licens fås via support@mr-note.com. Software og data fra tredjeparter står i THIRD_PARTY_NOTICES.md. Bidrag: se CONTRIBUTING.md.

## Mere

Websted og hele produktet: **https://mroffice.ai** · Spørgsmål og feedback: support@mr-note.com

<p align="center">Hvis MR Office er nyttigt for dig, hjælper en ⭐ på GitHub andre med at finde det.</p>
