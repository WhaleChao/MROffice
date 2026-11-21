# MR Office — Norsk

<p align="center"><sub>🌐 <a href="../README.md">English</a> · <a href="../i18n/README.en-GB.md">English (UK)</a> · <a href="../i18n/README.zh.md">简体中文</a> · <a href="../i18n/README.zh-TW.md">繁體中文</a> · <a href="../i18n/README.ja.md">日本語</a> · <a href="../i18n/README.ko.md">한국어</a> · <a href="../i18n/README.de.md">Deutsch</a> · <a href="../i18n/README.fr.md">Français</a> · <a href="../i18n/README.es.md">Español</a> · <a href="../i18n/README.es-419.md">Español (Latinoamérica)</a> · <a href="../i18n/README.pt-BR.md">Português (Brasil)</a> · <a href="../i18n/README.it.md">Italiano</a> · <a href="../i18n/README.nl.md">Nederlands</a> · <a href="../i18n/README.sv.md">Svenska</a> · <a href="../i18n/README.da.md">Dansk</a> · <a href="../i18n/README.no.md">Norsk</a> · <a href="../i18n/README.fi.md">Suomi</a> · <a href="../i18n/README.id.md">Bahasa Indonesia</a> · <a href="../i18n/README.th.md">ไทย</a> · <a href="../i18n/README.vi.md">Tiếng Việt</a></sub></p>

<p align="center"><b>MR Offices redigeringsverktøy med åpen kildekode – dokumenter, regneark, presentasjoner, PDF, håndskrift og kode – i nettleseren.</b></p>

<p align="center"><a href="https://mroffice.ai/no"><b>▶ Prøv på nett</b></a> · <a href="https://mroffice.ai/no/office/download"><b>⬇ Last ned</b></a> · <a href="https://mroffice.ai">🌐 mroffice.ai</a><br><a href="https://mroffice.ai/no/developers">🧑‍💻 Utviklere</a> · <a href="https://mroffice.ai/no/developers/embed">🧩 Bygg inn</a> · <a href="https://mroffice.ai/no/developers/mcp">🤖 MCP for KI-apper</a> · <a href="https://www.npmjs.com/package/mroffice-embed">📦 npm-SDK</a> · <a href="https://mroffice.ai/no/developers/oauth">🔑 Logg på med MR Office</a></p>

<p align="center"><a href="../LICENSE"><img alt="AGPL-3.0" src="https://img.shields.io/badge/license-AGPL--3.0-blue"></a> <a href="https://github.com/XinaTian/MROffice/releases"><img alt="release" src="https://img.shields.io/github/v/release/XinaTian/MROffice"></a> <a href="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml/badge.svg"></a> <a href="https://github.com/XinaTian/MROffice/stargazers"><img alt="GitHub stars" src="https://img.shields.io/github/stars/XinaTian/MROffice?style=social"></a></p>

<p align="center"><a href="https://mroffice.ai/no"><img src="../docs/media/mr-office-demo.webp" alt="MR Offices redigeringsverktøy med åpen kildekode – dokumenter, regneark, presentasjoner, PDF, håndskrift og kode – i nettleseren." width="100%"></a></p>

<p align="center"><a href="https://demo.mroffice.ai/media/mr-office-demo.mp4">🔊 Se den 50 sekunder lange omvisningen med lyd</a></p>

<table>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/no"><img src="../docs/screenshots/docs.png" alt="Dokumenter" width="100%"></a><br><b>Dokumenter</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/no"><img src="../docs/screenshots/sheets.png" alt="Regneark" width="100%"></a><br><b>Regneark</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/no"><img src="../docs/screenshots/slides.png" alt="Presentasjoner" width="100%"></a><br><b>Presentasjoner</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/no"><img src="../docs/screenshots/pdf.png" alt="PDF" width="100%"></a><br><b>PDF</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/no"><img src="../docs/screenshots/ink.png" alt="Håndskrift" width="100%"></a><br><b>Håndskrift</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/no"><img src="../docs/screenshots/code.png" alt="Kode" width="100%"></a><br><b>Kode</b></td>
</tr>
</table>

MR Office er et arbeidsområde for notater, KI og ekte filer. Dette repositoriet publiserer kildekoden til redigeringsverktøyene: de samme som MR Office bruker, og som du selv kan bygge, kjøre og bygge inn. Dette er redigeringsverktøy med åpen kildekode, ikke en komplett kontorpakke med åpen kildekode: lesing og skriving av Microsoft Office-filer blir værende i MR Office (se nedenfor).

## For utviklere

Alt for utviklere finnes på **[MR Office Developers](https://mroffice.ai/no/developers)** — selvbetjening, ingen søknad: logg på med en hvilken som helst MR Office-konto og sett i gang.

### Bygg inn redigeringsprogrammene i nettappen din (anbefalt)

MR Office er vert for redigeringsprogrammene; serveren din signerer et kortlivet økttoken, og siden din åpner redigeringsprogrammet med det. Det redigerer ekte Word-, Excel-, PowerPoint- og PDF-filer (.docx, .xlsx, .pptx, .pdf), med samredigering, opplastinger, alle eksportformater, utskrift og lagringer som sendes til serveren din i en signert tilbakekalling. Gratisplan: 1 000 økter i måneden på localhost, 100 på dine egne domener, opptil 2 apper. Det er en vertsbasert kommersiell tjeneste, så din egen app er ikke bundet av AGPL-lisensen til dette depotet.

1. Kjør `npx mroffice-embed init` i prosjektmappen (eller opprett en app i utviklerkonsollen): ett klikk på «Tillat» i nettleseren, så skrives appens nøkler til `.env.local` / `.env`.
2. Serveren din lager et økttoken med `createSessionToken` fra `mroffice-embed/server`; den hemmelige nøkkelen forlater aldri serveren.
3. Siden din åpner redigeringsprogrammet med `open()` fra `mroffice-embed` (eller skriptet `https://mroffice.ai/embed/v1/embed.js`).

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

[Veiledning for innbygging](https://mroffice.ai/no/developers/embed) · [Veiledningen som Markdown (for KI-agenter)](https://mroffice.ai/developers/embed.md) · [Priser og grenser](https://mroffice.ai/no/developers/pricing) · [Utviklerkonsoll](https://mroffice.ai/no/developers/console) · [Live-demo](https://mroffice.ai/no/developers/embed/demo)

### KI-apper: MCP-serveren

KI-apper åpner, leser, redigerer og eksporterer dokumenter i det ekte MR Office-redigeringsprogrammet via MCP-serveren: koble til det vertsbaserte endepunktet `https://mcp.mroffice.ai/mcp` (logg på med MR Office-kontoen din, ingen nøkkel), eller kjør `npx -y mroffice-mcp` på din egen datamaskin. Kildekoden ligger i `apps/mcp`. [MCP-veiledning →](https://mroffice.ai/no/developers/mcp)

### Logg på med MR Office (OAuth)

MR Office er en OAuth 2.1-autorisasjonsserver (PKCE, metadatadokumenter for klient-ID, dynamisk registrering): appen din kan logge folk på med MR Office-kontoen deres og handle med tillatelsene de gir. [OAuth-veiledning →](https://mroffice.ai/no/developers/oauth)

### Kjør redigeringsprogrammene med åpen kildekode selv

De seks redigeringsprogrammene i dette depotet kan bygges inn i ditt eget produkt under AGPL-3.0: bruker andre en endret versjon over et nettverk, må du tilby dem den fullstendige kildekoden og beholde merkingen «MR Office» (NOTICE.md). De åpner egne formater, CSV, HTML, Markdown og PDF; Word-, Excel- og PowerPoint-filer er ikke med i redigeringsprogrammene med åpen kildekode (den vertsbaserte innbyggingen over har dem). En kommersiell lisens fås via support@mr-note.com. [Veiledning for åpen kildekode →](https://mroffice.ai/no/developers/open-source)

npm-pakker forberedes under navnene `mroffice-core`, `mroffice-docs`, `mroffice-sheets`, `mroffice-slides`, `mroffice-pdf`, `mroffice-ink` og `mroffice-code`; til de er publisert, bygger du redigeringsprogrammene fra dette depotet (se nedenfor).

## Redigeringsverktøyene

Hva hvert verktøy gjør i dette repositoriet:

- **Dokumenter** — et sidebasert tekstbehandlingsprogram: stiler, skrifter, lister, tabeller, topp- og bunntekster, inndelinger og spalter, figurer og formler, forsider og innholdsfortegnelser, kommentarer og sporede endringer, samt håndskrift på siden. Lagrer i eget format (.mrdoc); åpner HTML, Markdown og ren tekst.
- **Regneark** — celler og formler, tallformater, betinget formatering, sortering og filtrering, frosne ruter og flere ark, med sum og gjennomsnitt for markeringen. Regning mellom celler (=A1*B1+10) beregnes i nettleseren mens du skriver; formler med funksjoner (SUMMER, HVIS, FINN.RAD …) beholdes slik de er skrevet. Lagrer i eget format (.mrsheet); åpner og lagrer CSV og TSV.
- **Presentasjoner** — lysbilder med åtte egne temaer, oppsett, lysbildemal- og oppsettvisning, tekst, figurer, tekstbokser, tabeller, WordArt og formler, kommentarer, inndelinger, overganger og animasjoner med forhåndsvisning, og lysbildefremvisning. Lagrer i eget format (.mrslides).
- **PDF** — lese, søke i, presentere og skrive ut PDF-filer, kopiere og lese opp teksten, og markere valgt tekst (utheving, understreking, gjennomstreking, bølgelinje) med notater og et kommentarpanel mens filen er åpen; lagring gir tilbake filen slik den ble åpnet.
- **Håndskrift** — håndskrevne notater med trykkfølsomme penner, markeringspenner, viskelær, lasso og linjal, og håndskrift som blir figurer. Lagrer i eget format (.mrink).
- **Kode** — et lett koderedigeringsverktøy med syntaksutheving for mange språk, en disposisjon og søk og erstatt; JavaScript/TypeScript, Python, SQL, Lua og skall kjører rett i nettleseren.

Beregning av formler med funksjoner, tegning av diagrammer og oppsett av SmartArt gjøres av MR Offices egen motor i MR Offices skrivebordsapp: redigeringsprogrammene med åpen kildekode beholder formlene, diagrammene og SmartArt i en fil og viser SmartArt slik det ble lagret.

## Kjør redigeringsverktøyene

**▶ [Prøv live-demoen – ingen installasjon](https://mroffice.ai/no)**

De seks verktøyene ligger i `packages/` (`@mrnote/docs`, `sheets`, `slides`, `pdf`, `ink`, `code`, som deler `@mrnote/core`) og kjører sammen i `apps/playground`:

```bash
npm install
npm run fonts     # the open substitutes for Office's fonts
npm run build
npm run preview   # http://localhost:5178
```

`npm run dev` starter en utviklingsserver på samme adresse, og `npm run smoke` kontrollerer et bygg. README-filen for hver pakke beskriver i detalj hva den gjør og hvordan den bygges inn.

## Skrivebordsapp

[`apps/desktop`](../apps/desktop) inneholder **MR Office Editors**: kildekoden til skallet i MR Offices skrivebordsapp (Electron), med sin egen app-identitet. Med den gratis MR Office-komponenten fra Releases er det MR Offices skrivebordsapp, med Word-, Excel-, PowerPoint- og PDF-filer, innlogging og AI. Uten komponenten kjører bare redigeringsprogrammene med åpen kildekode; de åpner .mrdoc-, .mrsheet-, .mrslides- og .mrink-filer samt CSV, HTML, Markdown og PDF. Skallet og redigeringsprogrammene er AGPL-kildekode; komponenten er kompilert, gratis å bruke og ikke åpen kildekode. Etter `npm run build` starter du den med:

```bash
npm run desktop
```

## For alle funksjoner: installer MR Offices skrivebordsapp (gratis)

MR Offices skrivebordsapp er gratis, har disse redigeringsprogrammene og hele tjenesten, og fungerer frakoblet på din egen datamaskin. Den legger til det redigeringsprogrammene med åpen kildekode mangler:

- Åpne og lagre Word-, Excel- og PowerPoint-filer (.docx, .xlsx, .pptx)
- Beregne formler med funksjoner (SUMMER, HVIS, FINN.RAD …)
- Diagrammer
- Redigering av SmartArt
- Lagre endringer i en PDF tilbake i filen

**[Last ned MR Office (gratis) →](https://mroffice.ai/no/office/download)**

### Word-, Excel-, PowerPoint- og PDF-filer

Å åpne og lagre Word-, Excel- og PowerPoint-filer (.docx, .xlsx, .pptx) og lagre endringer i en PDF tilbake i filen gjøres av MR Offices gratis skrivebordsapp. Den har alle disse redigeringsprogrammene og fungerer frakoblet på din egen datamaskin:

## Slik henger delene sammen

```mermaid
flowchart LR
  subgraph OSS["Åpen kildekode · AGPL-3.0"]
    direction LR
    core["@mrnote/core"] --> E["De seks redigeringsprogrammene<br/>Dokumenter · Regneark · Presentasjoner · PDF · Håndskrift · Kode"]
    E --> PG["apps/playground<br/>Nettleser"]
    E --> DS["apps/desktop<br/>Skrivebordsskall"]
    MCP["apps/mcp<br/>MCP-server for KI-apper"]
  end
  COMP["MR Office-komponenten: Word-, Excel- og PowerPoint-filer · gratis, ikke åpen kildekode"] -.-> DS
```

## Lisens

Åpen kildekode under GNU Affero General Public License v3.0 (LICENSE), med tilleggsvilkårene i NOTICE.md: behold navngivelsen «MR Office» og lenken til https://mroffice.ai, og ingen rettigheter til navnet eller logoen MR Office. En kommersiell lisens fås via support@mr-note.com. Programvare og data fra tredjeparter står i THIRD_PARTY_NOTICES.md. Bidrag: se CONTRIBUTING.md.

## Mer

Nettsted og hele produktet: **https://mroffice.ai** · Spørsmål og tilbakemeldinger: support@mr-note.com

<p align="center">Hvis MR Office er nyttig for deg, hjelper en ⭐ på GitHub andre å finne det.</p>
