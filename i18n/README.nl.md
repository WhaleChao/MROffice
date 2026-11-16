# MR Office — Nederlands

[English](../README.md) · [English (UK)](../i18n/README.en-GB.md) · [简体中文](../i18n/README.zh.md) · [繁體中文](../i18n/README.zh-TW.md) · [日本語](../i18n/README.ja.md) · [한국어](../i18n/README.ko.md) · [Deutsch](../i18n/README.de.md) · [Français](../i18n/README.fr.md) · [Español](../i18n/README.es.md) · [Español (Latinoamérica)](../i18n/README.es-419.md) · [Português (Brasil)](../i18n/README.pt-BR.md) · [Italiano](../i18n/README.it.md) · [Nederlands](../i18n/README.nl.md) · [Svenska](../i18n/README.sv.md) · [Dansk](../i18n/README.da.md) · [Norsk](../i18n/README.no.md) · [Suomi](../i18n/README.fi.md) · [Bahasa Indonesia](../i18n/README.id.md) · [ไทย](../i18n/README.th.md) · [Tiếng Việt](../i18n/README.vi.md)

<p align="center"><b>De opensource-editors van MR Office — documenten, spreadsheets, presentaties, pdf, handschrift en code — in de browser.</b></p>

<p align="center"><a href="https://demo.mroffice.ai"><b>▶ Probeer de live demo – zonder installatie</b></a> · <a href="https://github.com/XinaTian/MROffice/releases">⬇ MR Office downloaden via Releases</a> · <a href="https://mroffice.ai">🌐 mroffice.ai</a></p>

<p align="center"><a href="../LICENSE"><img alt="AGPL-3.0" src="https://img.shields.io/badge/license-AGPL--3.0-blue"></a> <a href="https://github.com/XinaTian/MROffice/releases"><img alt="release" src="https://img.shields.io/github/v/release/XinaTian/MROffice"></a> <a href="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml/badge.svg"></a></p>

<table>
<tr>
<td align="center" width="33%"><a href="https://demo.mroffice.ai"><img src="../docs/screenshots/docs.png" alt="Documenten" width="100%"></a><br><b>Documenten</b></td>
<td align="center" width="33%"><a href="https://demo.mroffice.ai"><img src="../docs/screenshots/sheets.png" alt="Spreadsheets" width="100%"></a><br><b>Spreadsheets</b></td>
<td align="center" width="33%"><a href="https://demo.mroffice.ai"><img src="../docs/screenshots/slides.png" alt="Presentaties" width="100%"></a><br><b>Presentaties</b></td>
</tr>
<tr>
<td align="center" width="33%"><a href="https://demo.mroffice.ai"><img src="../docs/screenshots/pdf.png" alt="PDF" width="100%"></a><br><b>PDF</b></td>
<td align="center" width="33%"><a href="https://demo.mroffice.ai"><img src="../docs/screenshots/ink.png" alt="Handschrift" width="100%"></a><br><b>Handschrift</b></td>
<td align="center" width="33%"><a href="https://demo.mroffice.ai"><img src="../docs/screenshots/code.png" alt="Code" width="100%"></a><br><b>Code</b></td>
</tr>
</table>

MR Office is een werkruimte voor notities, AI en echte bestanden. Deze repository publiceert de broncode van de editors: dezelfde editors die MR Office gebruikt, die je zelf kunt bouwen, draaien en insluiten. Het zijn opensource-editors, geen volledige opensource-kantoorsuite: het lezen en schrijven van Microsoft Office-bestanden blijft in MR Office (zie hieronder).

## De editors

Wat elke editor in deze repository doet:

- **Documenten** — een tekstverwerker met pagina's: stijlen, lettertypen, lijsten, tabellen, kop- en voetteksten, secties en kolommen, vormen en vergelijkingen, voorbladen en inhoudsopgaven, opmerkingen en wijzigingen bijhouden, en inkt op de pagina. Slaat op in een eigen formaat (.mrdoc); opent HTML, Markdown en platte tekst.
- **Spreadsheets** — cellen en formules, getalnotaties, voorwaardelijke opmaak, sorteren en filteren, vastgezette deelvensters en meerdere bladen, met som en gemiddelde van de selectie. Rekenen tussen cellen (=A1*B1+10) gebeurt tijdens het typen in de browser; formules met functies (SOM, ALS, VERT.ZOEKEN …) blijven staan zoals ze zijn geschreven. Slaat op in een eigen formaat (.mrsheet); opent en bewaart CSV en TSV.
- **Presentaties** — dia's met acht eigen thema's, indelingen, diamodel- en indelingsweergave, tekst, vormen, tekstvakken, tabellen, WordArt en vergelijkingen, opmerkingen, secties, overgangen en animaties met voorbeeld, en een diavoorstelling. Slaat op in een eigen formaat (.mrslides).
- **PDF** — pdf-bestanden lezen, doorzoeken, presenteren en afdrukken, tekst kopiëren en laten voorlezen, en geselecteerde tekst markeren (markeren, onderstrepen, doorhalen, golflijn) met notities en een opmerkingenvenster zolang het bestand open is; opslaan geeft het bestand terug zoals het werd geopend.
- **Handschrift** — handgeschreven notities met drukgevoelige pennen, markeerstiften, gummen, lasso en liniaal, en inkt die vormen wordt. Slaat op in een eigen formaat (.mrink).
- **Code** — een lichte code-editor met syntaxiskleuring voor veel talen, een overzicht, en zoeken en vervangen; JavaScript/TypeScript, Python, SQL, Lua en shell draaien direct in de browser.

Formules met functies berekenen, grafieken tekenen en SmartArt opmaken doet de eigen engine van MR Office, in de desktop-app van MR Office: de opensource-editors bewaren de formules, grafieken en SmartArt van een bestand en tonen SmartArt zoals het is opgeslagen.

## De editors draaien

**▶ [Probeer de live demo – zonder installatie](https://demo.mroffice.ai)**

De zes editors staan in `packages/` (`@mrnote/docs`, `sheets`, `slides`, `pdf`, `ink`, `code`, die `@mrnote/core` delen) en draaien samen in `apps/playground`:

```bash
npm install
npm run fonts     # the open substitutes for Office's fonts
npm run build
npm run preview   # http://localhost:5178
```

`npm run dev` start een ontwikkelserver op hetzelfde adres en `npm run smoke` controleert een build. De README van elk pakket beschrijft in detail wat het doet en hoe je het insluit.

## Voor alle functies: installeer de desktop-app van MR Office (gratis)

De desktop-app van MR Office is gratis, bevat deze editors en de volledige dienst, en werkt offline op je eigen computer. Hij voegt toe wat de opensource-editors niet hebben:

- Word-, Excel- en PowerPoint-bestanden (.docx, .xlsx, .pptx) openen en opslaan
- Formules met functies (SOM, ALS, VERT.ZOEKEN …) berekenen
- Grafieken
- SmartArt bewerken
- Bewerkingen van een pdf terugschrijven in het bestand

**[MR Office downloaden via Releases →](https://github.com/XinaTian/MROffice/releases)** · [of via mroffice.ai](https://mroffice.ai)

### Word-, Excel-, PowerPoint- en pdf-bestanden

Word-, Excel- en PowerPoint-bestanden (.docx, .xlsx, .pptx) openen en opslaan, en bewerkingen van een pdf terugschrijven in het bestand, doet de gratis desktop-app van MR Office. Die bevat al deze editors en werkt offline op je eigen computer:

## Licentie

Open source onder de GNU Affero General Public License v3.0 (LICENSE), met de aanvullende voorwaarden in NOTICE.md: behoud de vermelding "MR Office" en de link naar https://mroffice.ai, en geen rechten op de naam of het logo van MR Office. Een commerciële licentie is verkrijgbaar via support@mr-note.com. Software en gegevens van derden staan in THIRD_PARTY_NOTICES.md. Bijdragen: zie CONTRIBUTING.md.

## Meer

Website en het volledige product: **https://mroffice.ai** · Vragen en feedback: support@mr-note.com
