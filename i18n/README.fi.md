# MR Office — Suomi

<p align="center"><sub>🌐 <a href="../README.md">English</a> · <a href="../i18n/README.en-GB.md">English (UK)</a> · <a href="../i18n/README.zh.md">简体中文</a> · <a href="../i18n/README.zh-TW.md">繁體中文</a> · <a href="../i18n/README.ja.md">日本語</a> · <a href="../i18n/README.ko.md">한국어</a> · <a href="../i18n/README.de.md">Deutsch</a> · <a href="../i18n/README.fr.md">Français</a> · <a href="../i18n/README.es.md">Español</a> · <a href="../i18n/README.es-419.md">Español (Latinoamérica)</a> · <a href="../i18n/README.pt-BR.md">Português (Brasil)</a> · <a href="../i18n/README.it.md">Italiano</a> · <a href="../i18n/README.nl.md">Nederlands</a> · <a href="../i18n/README.sv.md">Svenska</a> · <a href="../i18n/README.da.md">Dansk</a> · <a href="../i18n/README.no.md">Norsk</a> · <a href="../i18n/README.fi.md">Suomi</a> · <a href="../i18n/README.id.md">Bahasa Indonesia</a> · <a href="../i18n/README.th.md">ไทย</a> · <a href="../i18n/README.vi.md">Tiếng Việt</a></sub></p>

<p align="center"><b>MR Officen avoimen lähdekoodin editorit – asiakirjat, laskentataulukot, esitykset, PDF, käsiala ja koodi – selaimessa.</b></p>

<p align="center"><a href="https://mroffice.ai/fi"><b>▶ Kokeile live-demoa – ei asennusta</b></a> · <a href="https://github.com/XinaTian/MROffice/releases">⬇ Lataa MR Office Releases-sivulta</a> · <a href="https://mroffice.ai">🌐 mroffice.ai</a></p>

<p align="center"><a href="../LICENSE"><img alt="AGPL-3.0" src="https://img.shields.io/badge/license-AGPL--3.0-blue"></a> <a href="https://github.com/XinaTian/MROffice/releases"><img alt="release" src="https://img.shields.io/github/v/release/XinaTian/MROffice"></a> <a href="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml/badge.svg"></a> <a href="https://github.com/XinaTian/MROffice/stargazers"><img alt="GitHub stars" src="https://img.shields.io/github/stars/XinaTian/MROffice?style=social"></a></p>

<p align="center"><a href="https://mroffice.ai/fi"><img src="../docs/media/mr-office-demo.webp" alt="MR Officen avoimen lähdekoodin editorit – asiakirjat, laskentataulukot, esitykset, PDF, käsiala ja koodi – selaimessa." width="100%"></a></p>

<p align="center"><a href="https://demo.mroffice.ai/media/mr-office-demo.mp4">🔊 Katso 50 sekunnin esittely äänen kanssa</a></p>

<table>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/fi"><img src="../docs/screenshots/docs.png" alt="Asiakirjat" width="100%"></a><br><b>Asiakirjat</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/fi"><img src="../docs/screenshots/sheets.png" alt="Laskentataulukot" width="100%"></a><br><b>Laskentataulukot</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/fi"><img src="../docs/screenshots/slides.png" alt="Esitykset" width="100%"></a><br><b>Esitykset</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/fi"><img src="../docs/screenshots/pdf.png" alt="PDF" width="100%"></a><br><b>PDF</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/fi"><img src="../docs/screenshots/ink.png" alt="Käsiala" width="100%"></a><br><b>Käsiala</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/fi"><img src="../docs/screenshots/code.png" alt="Koodi" width="100%"></a><br><b>Koodi</b></td>
</tr>
</table>

MR Office on työtila muistiinpanoille, tekoälylle ja oikeille tiedostoille. Tämä tietovarasto julkaisee sen editorien lähdekoodin: samat editorit, joita MR Office käyttää ja jotka voit itse koota, ajaa ja upottaa. Nämä ovat avoimen lähdekoodin editoreja, eivät täydellinen avoimen lähdekoodin toimisto-ohjelmisto: Microsoft Office -tiedostojen lukeminen ja kirjoittaminen jäävät MR Officeen (katso alla).

## Editorit

Mitä kukin editori tekee tässä tietovarastossa:

- **Asiakirjat** — sivupohjainen tekstinkäsittely: tyylit, fontit, luettelot, taulukot, ylä- ja alatunnisteet, osat ja palstat, muodot ja kaavat, kansilehdet ja sisällysluettelot, kommentit ja muutosten seuranta sekä käsin piirtäminen sivulle. Tallentaa omaan muotoonsa (.mrdoc); avaa HTML-, Markdown- ja tekstitiedostot.
- **Laskentataulukot** — solut ja kaavat, lukumuodot, ehdollinen muotoilu, lajittelu ja suodatus, kiinnitetyt ruudut ja useita taulukoita sekä valinnan summa ja keskiarvo. Solujen välinen laskenta (=A1*B1+10) lasketaan selaimessa kirjoitettaessa; funktioita käyttävät kaavat (SUMMA, JOS, PHAKU …) säilyvät kirjoitetussa muodossa. Tallentaa omaan muotoonsa (.mrsheet); avaa ja tallentaa CSV- ja TSV-tiedostot.
- **Esitykset** — dioja kahdeksalla omalla teemalla: asettelut, diapohja- ja asettelunäkymä, teksti, muodot, tekstiruudut, taulukot, WordArt ja kaavat, kommentit, osiot, siirtymät ja animaatiot esikatseluineen sekä diaesitys. Tallentaa omaan muotoonsa (.mrslides).
- **PDF** — PDF-tiedostojen lukeminen, haku, esittäminen ja tulostus, tekstin kopiointi ja ääneen lukeminen sekä valitun tekstin merkitseminen (korostus, alleviivaus, yliviivaus, aaltoviiva) muistilappujen ja kommenttiruudun kanssa tiedoston ollessa auki; tallennus palauttaa tiedoston sellaisena kuin se avattiin.
- **Käsiala** — käsin kirjoitetut muistiinpanot paineherkillä kynillä, korostuskynillä, pyyhekumeilla, lassolla ja viivaimella, ja käsiala muuttuu muodoiksi. Tallentaa omaan muotoonsa (.mrink).
- **Koodi** — kevyt koodieditori, jossa on syntaksikorostus monille kielille, jäsennys sekä etsi ja korvaa; JavaScript/TypeScript, Python, SQL, Lua ja komentotulkki toimivat suoraan selaimessa.

Funktioita sisältävien kaavojen laskennan, kaavioiden piirtämisen ja SmartArtin asettelun hoitaa MR Officen oma moottori MR Office -työpöytäsovelluksessa: avoimen lähdekoodin editorit säilyttävät tiedoston kaavat, kaaviot ja SmartArtin ja näyttävät SmartArtin sellaisena kuin se tallennettiin.

## Editorien ajaminen

**▶ [Kokeile live-demoa – ei asennusta](https://mroffice.ai/fi)**

Kuusi editoria ovat hakemistossa `packages/` (`@mrnote/docs`, `sheets`, `slides`, `pdf`, `ink`, `code`, yhteisenä `@mrnote/core`), ja ne toimivat yhdessä hakemistossa `apps/playground`:

```bash
npm install
npm run fonts     # the open substitutes for Office's fonts
npm run build
npm run preview   # http://localhost:5178
```

`npm run dev` käynnistää kehityspalvelimen samaan osoitteeseen, ja `npm run smoke` tarkistaa koontiversion. Jokaisen paketin README kertoo tarkemmin, mitä se tekee ja miten sen voi upottaa.

## Työpöytäsovellus

[`apps/desktop`](../apps/desktop) sisältää **MR Office Editorsin**: MR Officen työpöytäsovelluksen kuoren (Electron) lähdekoodin omalla sovellustunnisteellaan. Kun Releases-sivun ilmainen MR Office -komponentti on asennettu, se on MR Officen työpöytäsovellus: Word-, Excel-, PowerPoint- ja PDF-tiedostot, kirjautuminen ja tekoäly. Ilman komponenttia se ajaa vain avoimen lähdekoodin editoreja, jotka avaavat .mrdoc-, .mrsheet-, .mrslides- ja .mrink-tiedostot sekä CSV-, HTML-, Markdown- ja PDF-tiedostot. Kuori ja editorit ovat AGPL-lähdekoodia; komponentti on käännetty, ilmainen käyttää eikä avointa lähdekoodia. Kun olet ajanut `npm run build`, käynnistä se komennolla:

```bash
npm run desktop
```

## Kaikki ominaisuudet: asenna MR Officen työpöytäsovellus (maksuton)

MR Office -työpöytäsovellus on maksuton, sisältää nämä editorit ja koko palvelun ja toimii offline-tilassa omalla tietokoneellasi. Se tuo sen, mitä avoimen lähdekoodin editoreista puuttuu:

- Word-, Excel- ja PowerPoint-tiedostojen (.docx, .xlsx, .pptx) avaaminen ja tallentaminen
- Funktioita sisältävien kaavojen (SUMMA, JOS, PHAKU …) laskenta
- Kaaviot
- SmartArtin muokkaus
- PDF-muokkausten tallentaminen takaisin tiedostoon

**[Lataa MR Office Releases-sivulta →](https://github.com/XinaTian/MROffice/releases)** · [tai osoitteesta mroffice.ai](https://mroffice.ai)

### Word-, Excel-, PowerPoint- ja PDF-tiedostot

Word-, Excel- ja PowerPoint-tiedostojen (.docx, .xlsx, .pptx) avaamisen ja tallentamisen sekä PDF-muokkausten tallentamisen takaisin tiedostoon tarjoaa maksuton MR Office -työpöytäsovellus. Se sisältää kaikki nämä editorit ja toimii offline-tilassa omalla tietokoneellasi:

## Miten osat liittyvät toisiinsa

```mermaid
flowchart LR
  subgraph OSS["Avoin lähdekoodi · AGPL-3.0"]
    direction LR
    core["@mrnote/core"] --> E["Kuusi editoria<br/>Asiakirjat · Laskentataulukot · Esitykset · PDF · Käsiala · Koodi"]
    E --> PG["apps/playground<br/>Selain"]
    E --> DS["apps/desktop<br/>Työpöytäsovelluksen kuori"]
    MCP["apps/mcp<br/>MCP-palvelin tekoälysovelluksille"]
  end
  COMP["MR Office -komponentti: Word-, Excel- ja PowerPoint-tiedostot · ilmainen, ei avointa lähdekoodia"] -.-> DS
```

## Lisenssi

Avointa lähdekoodia GNU Affero General Public License v3.0 -lisenssillä (LICENSE) ja NOTICE.md-tiedoston lisäehdoin: säilytä maininta ”MR Office” ja linkki osoitteeseen https://mroffice.ai; MR Office -nimeen ja -logoon ei anneta oikeuksia. Kaupallisen lisenssin saa osoitteesta support@mr-note.com. Kolmansien osapuolten ohjelmistot ja aineistot on lueteltu tiedostossa THIRD_PARTY_NOTICES.md. Osallistuminen: katso CONTRIBUTING.md.

## Lisää

Verkkosivusto ja koko tuote: **https://mroffice.ai** · Kysymykset ja palaute: support@mr-note.com

<p align="center">Jos MR Office on sinulle hyödyllinen, ⭐ GitHubissa auttaa muita löytämään sen.</p>
