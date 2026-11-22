# MR Office — Nederlands

<p align="center"><sub>🌐 <a href="../README.md">English</a> · <a href="../i18n/README.en-GB.md">English (UK)</a> · <a href="../i18n/README.zh.md">简体中文</a> · <a href="../i18n/README.zh-TW.md">繁體中文</a> · <a href="../i18n/README.ja.md">日本語</a> · <a href="../i18n/README.ko.md">한국어</a> · <a href="../i18n/README.de.md">Deutsch</a> · <a href="../i18n/README.fr.md">Français</a> · <a href="../i18n/README.es.md">Español</a> · <a href="../i18n/README.es-419.md">Español (Latinoamérica)</a> · <a href="../i18n/README.pt-BR.md">Português (Brasil)</a> · <a href="../i18n/README.it.md">Italiano</a> · <a href="../i18n/README.nl.md">Nederlands</a> · <a href="../i18n/README.sv.md">Svenska</a> · <a href="../i18n/README.da.md">Dansk</a> · <a href="../i18n/README.no.md">Norsk</a> · <a href="../i18n/README.fi.md">Suomi</a> · <a href="../i18n/README.id.md">Bahasa Indonesia</a> · <a href="../i18n/README.th.md">ไทย</a> · <a href="../i18n/README.vi.md">Tiếng Việt</a></sub></p>

<p align="center"><b>De open-source editors van MR Office, de editor die het dichtst bij Microsoft Office staat — documenten, spreadsheets, presentaties, PDF en code — in de browser.</b></p>

<p align="center"><a href="https://mroffice.ai/nl"><b>▶ Online proberen</b></a> · <a href="https://mroffice.ai/nl/office/download"><b>⬇ Downloaden</b></a> · <a href="https://mroffice.ai">🌐 mroffice.ai</a><br><a href="https://mroffice.ai/nl/developers">🧑‍💻 Ontwikkelaars</a> · <a href="https://mroffice.ai/nl/developers/embed">🧩 Insluiten</a> · <a href="https://mroffice.ai/nl/developers/mcp">🤖 MCP voor AI-apps</a> · <a href="https://www.npmjs.com/package/mroffice-embed">📦 npm-SDK</a> · <a href="https://mroffice.ai/nl/developers/oauth">🔑 Inloggen met MR Office</a></p>

<p align="center"><a href="../LICENSE"><img alt="AGPL-3.0" src="https://img.shields.io/badge/license-AGPL--3.0-blue"></a> <a href="https://github.com/XinaTian/MROffice/releases"><img alt="release" src="https://img.shields.io/github/v/release/XinaTian/MROffice"></a> <a href="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml/badge.svg"></a> <a href="https://github.com/XinaTian/MROffice/stargazers"><img alt="GitHub stars" src="https://img.shields.io/github/stars/XinaTian/MROffice?style=social"></a></p>

<p align="center"><a href="https://mroffice.ai/nl"><img src="../docs/media/mr-office-demo.webp" alt="De open-source editors van MR Office, de editor die het dichtst bij Microsoft Office staat — documenten, spreadsheets, presentaties, PDF en code — in de browser." width="100%"></a></p>

<p align="center"><a href="https://demo.mroffice.ai/media/mr-office-demo.mp4">🔊 Bekijk de rondleiding van 50 seconden met geluid</a></p>

<table>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/nl"><img src="../docs/screenshots/docs.png" alt="Documenten" width="100%"></a><br><b>Documenten</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/nl"><img src="../docs/screenshots/sheets.png" alt="Spreadsheets" width="100%"></a><br><b>Spreadsheets</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/nl"><img src="../docs/screenshots/slides.png" alt="Presentaties" width="100%"></a><br><b>Presentaties</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/nl"><img src="../docs/screenshots/pdf.png" alt="PDF" width="100%"></a><br><b>PDF</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/nl"><img src="../docs/screenshots/code.png" alt="Code" width="100%"></a><br><b>Code</b></td>
</tr>
</table>

MR Office is een werkruimte voor notities, AI en echte bestanden. Deze repository publiceert de broncode van de editors: dezelfde editors die MR Office gebruikt, die je zelf kunt bouwen, uitvoeren en embedden. Samen met de gratis MR Office desktop-app of de gehoste embed openen en opslaan ze echte Word-, Excel- en PowerPoint-bestanden.

## Voor ontwikkelaars

Alles voor ontwikkelaars staat op **[MR Office Developers](https://mroffice.ai/nl/developers)** — selfservice, geen aanvraag: log in met een willekeurig MR Office-account en begin.

### De editors in je webapp insluiten (aanbevolen)

MR Office host de editors; je server ondertekent een kortlevend sessietoken en je pagina opent de editor. Hij bewerkt echte Word-, Excel-, PowerPoint- en PDF-bestanden (.docx, .xlsx, .pptx, .pdf), met co-editing, uploads, elk exportformaat, afdrukken en opslaan teruggestuurd naar je server in een ondertekende callback. Gratis plan: 1.000 sessies per maand op localhost, 100 op je eigen domeinen, tot 2 apps. Als gehoste service legt het geen AGPL-verplichtingen op aan je eigen app.

1. Voer in de map van je project `npx mroffice-embed init` uit (of maak een app in de ontwikkelaarsconsole): één klik op "Toestaan" in de browser en de sleutels van de app staan in `.env.local` / `.env`.
2. Je server maakt een sessietoken met `createSessionToken` uit `mroffice-embed/server`; de geheime sleutel verlaat de server nooit.
3. Je pagina opent de editor met `open()` uit `mroffice-embed` (of het script `https://mroffice.ai/embed/v1/embed.js`).

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

[Insluitgids](https://mroffice.ai/nl/developers/embed) · [De gids als Markdown (voor AI-agents)](https://mroffice.ai/developers/embed.md) · [Prijzen en limieten](https://mroffice.ai/nl/developers/pricing) · [Ontwikkelaarsconsole](https://mroffice.ai/nl/developers/console) · [Live demo](https://mroffice.ai/nl/developers/embed/demo)

### AI-apps: de MCP-server

AI-apps openen, lezen, bewerken en exporteren documenten in de echte MR Office-editor via de MCP-server: verbind het gehoste eindpunt `https://mcp.mroffice.ai/mcp` (inloggen met je MR Office-account, geen sleutel) of voer `npx -y mroffice-mcp` uit op je eigen computer. De broncode staat in `apps/mcp`. [MCP-gids →](https://mroffice.ai/nl/developers/mcp)

### Inloggen met MR Office (OAuth)

MR Office is een OAuth 2.1-autorisatieserver (PKCE, client-ID-metadatadocumenten, dynamische registratie): je app kan mensen laten inloggen met hun MR Office-account en handelen met de scopes die zij toestaan. [OAuth-gids →](https://mroffice.ai/nl/developers/oauth)

### De open-source editors zelf hosten

De vijf editors in deze repository kunnen onder de AGPL-3.0 in je eigen product worden gebouwd: als mensen een gewijzigde versie via een netwerk gebruiken, moet je hun de volledige broncode aanbieden en de vermelding "MR Office" behouden (NOTICE.md). Ze openen hun eigen formaten, CSV, HTML, Markdown en PDF; gebruik voor Word-, Excel- en PowerPoint-bestanden de gehoste embed hierboven. Een commerciële licentie is verkrijgbaar via support@mr-note.com. [Open-sourcegids →](https://mroffice.ai/nl/developers/open-source)

De editors staan op npm als React-componenten: `npm i mroffice-docs mroffice-core` (en `mroffice-sheets`, `mroffice-slides`, `mroffice-pdf`, `mroffice-code`); de README van elk pakket laat zien hoe je het insluit.

## De editors

Wat elke editor in deze repository doet:

- **Documenten** — een paginagebonden tekstverwerker: lettertypen, alinea-opmaak, lijsten en stijlen op Start; voorblad en lege pagina's, pagina-einden, tabellen, afbeeldingen, vormen, SmartArt, grafieken, koppelingen, bladwijzers, opmerkingen, kop- en voetteksten en paginanummers, vergelijkingen, symbolen, tekstvakken en WordArt op Invoegen. Slaat het eigen formaat op (.mrdoc); opent HTML, Markdown en platte tekst.
- **Spreadsheets** — cellen en formules met lettertypen, uitlijning, getalnotaties, voorwaardelijke opmaak, celstijlen en het invoegen of verwijderen van cellen op Start; tabellen, afbeeldingen, vormen, pictogrammen, selectievakjes, grafieken en sparklines op Invoegen; meerdere bladen, met sommen en gemiddelden van de selectie. Rekenkunde tussen cellen (=A1*B1+10) wordt tijdens het typen in de browser berekend; formules met functies (SUM, IF, VLOOKUP …) blijven zoals geschreven. Slaat het eigen formaat op (.mrsheet); opent en slaat CSV en TSV op.
- **Presentaties** — dia's met lay-outs en secties, tekst, lettertypen en alinea's, en vormen met hun rangschikking op Start; tabellen, afbeeldingen, pictogrammen, SmartArt, grafieken, tekstvakken, WordArt, kop- en voetteksten, vergelijkingen, symbolen en video op Invoegen; sprekersnotities, opmerkingen en presenteren. Slaat het eigen formaat op (.mrslides).
- **PDF** — PDF-bestanden lezen, presenteren en afdrukken: weergaven met één en twee pagina's, continu scrollen, draaien, miniaturen en leesmodus, zoeken in en kopiëren van hun tekst, en opmerkingen terwijl een bestand open is; opslaan geeft het bestand terug zoals het werd geopend.
- **Code** — een lichte code-editor met syntaxiskleuring voor veel talen, een overzicht, zoeken, snel openen, een bestandsverkenner en een problemenlijst.

Formules met functies, grafieken en SmartArt worden berekend en getekend door de engine van MR Office, in de gratis desktop-app en de gehoste service; de editors hier behouden elke formule, grafiek en SmartArt die een bestand bevat, en tonen SmartArt zoals het is opgeslagen.

## De editors draaien

**▶ [Probeer de live demo – zonder installatie](https://mroffice.ai/nl)**

De vijf editors staan in `packages/` (`@mrnote/docs`, `sheets`, `slides`, `pdf`, `code`, met gedeelde `@mrnote/core`) en draaien samen in `apps/playground`:

```bash
npm install
npm run fonts     # the open substitutes for Office's fonts
npm run build
npm run preview   # http://localhost:5178
```

`npm run dev` start een ontwikkelserver op hetzelfde adres en `npm run smoke` controleert een build. De README van elk pakket beschrijft in detail wat het doet en hoe je het insluit.

## Desktop-app

[`apps/desktop`](../apps/desktop) bevat **MR Office Editors**: de bron van de shell van de MR Office-desktopapp (Electron), onder een eigen identiteit. Met de gratis MR Office-component uit Releases is het de MR Office-desktopapp, met Word-, Excel-, PowerPoint- en PDF-bestanden, aanmelden en AI. Zonder de component draait het de open-source editors op zichzelf, die .mrdoc-, .mrsheet- en .mrslides-bestanden, CSV, HTML, Markdown en PDF openen. De shell en de editors zijn AGPL-broncode; de component is een gratis download uit Releases. Na `npm run build` start je het met:

```bash
npm run desktop
```

## Voor alle functies: installeer de desktop-app van MR Office (gratis)

De MR Office desktop-app is gratis, bevat deze editors en de volledige service, en werkt offline op je eigen computer. Bovenop de editors voegt hij toe:

- Word-, Excel- en PowerPoint-bestanden (.docx, .xlsx, .pptx) openen en opslaan
- Formules met functies (SOM, ALS, VERT.ZOEKEN …) berekenen
- Grafieken
- SmartArt bewerken
- Bewerkingen van een pdf terugschrijven in het bestand

**[MR Office downloaden (gratis) →](https://mroffice.ai/nl/office/download)**

### Word-, Excel-, PowerPoint- en pdf-bestanden

Werk met echte Word-, Excel- en PowerPoint-bestanden (.docx, .xlsx, .pptx) en sla PDF-bewerkingen direct op in het bestand met de gratis MR Office desktop-app, die al deze editors bevat en offline op je eigen computer werkt:

## Hoe de onderdelen samenhangen

```mermaid
flowchart LR
  subgraph OSS["Open source · AGPL-3.0"]
    direction LR
    core["@mrnote/core"] --> E["De vijf editors<br/>Documenten · Spreadsheets · Presentaties · PDF · Code"]
    E --> PG["apps/playground<br/>Browser"]
    E --> DS["apps/desktop<br/>Desktopshell"]
    MCP["apps/mcp<br/>MCP-server voor AI-apps"]
  end
  COMP["MR Office-component: Word-, Excel- en PowerPoint-bestanden · gratis"] -.-> DS
```

## Licentie

Open source onder de GNU Affero General Public License v3.0 (LICENSE), met de aanvullende voorwaarden in NOTICE.md: behoud de vermelding "MR Office" en de link naar https://mroffice.ai, en geen rechten op de naam of het logo van MR Office. Een commerciële licentie is verkrijgbaar via support@mr-note.com. Software en gegevens van derden staan in THIRD_PARTY_NOTICES.md. Bijdragen: zie CONTRIBUTING.md.

## Meer

Website en het volledige product: **https://mroffice.ai** · Vragen en feedback: support@mr-note.com

<p align="center">Heb je iets aan MR Office? Een ⭐ op GitHub helpt anderen het te vinden.</p>
