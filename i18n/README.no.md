# MR Office — Norsk

[English](../README.md) · [English (UK)](../i18n/README.en-GB.md) · [简体中文](../i18n/README.zh.md) · [繁體中文](../i18n/README.zh-TW.md) · [日本語](../i18n/README.ja.md) · [한국어](../i18n/README.ko.md) · [Deutsch](../i18n/README.de.md) · [Français](../i18n/README.fr.md) · [Español](../i18n/README.es.md) · [Español (Latinoamérica)](../i18n/README.es-419.md) · [Português (Brasil)](../i18n/README.pt-BR.md) · [Italiano](../i18n/README.it.md) · [Nederlands](../i18n/README.nl.md) · [Svenska](../i18n/README.sv.md) · [Dansk](../i18n/README.da.md) · [Norsk](../i18n/README.no.md) · [Suomi](../i18n/README.fi.md) · [Bahasa Indonesia](../i18n/README.id.md) · [ไทย](../i18n/README.th.md) · [Tiếng Việt](../i18n/README.vi.md)

> ### ⬇ For alle funksjoner: installer MR Offices skrivebordsapp (gratis)
>
> MR Offices skrivebordsapp er gratis, har disse redigeringsprogrammene og hele tjenesten, og fungerer frakoblet på din egen datamaskin. Den legger til det redigeringsprogrammene med åpen kildekode mangler:
>
> - Åpne og lagre Word-, Excel- og PowerPoint-filer (.docx, .xlsx, .pptx)
> - Beregne formler med funksjoner (SUMMER, HVIS, FINN.RAD …)
> - Diagrammer
> - Redigering av SmartArt
> - Lagre endringer i en PDF tilbake i filen
>
> **[Last ned MR Office fra Releases →](https://github.com/XinaTian/MROffice/releases)** · [eller fra mroffice.ai](https://mroffice.ai)

<p align="center"><a href="https://mroffice.ai"><b>🌐 mroffice.ai</b></a> · <a href="https://mroffice.ai">Prøv dem i MR Office →</a></p>

> MR Offices redigeringsverktøy med åpen kildekode – dokumenter, regneark, presentasjoner, PDF, håndskrift og kode – i nettleseren.

MR Office er et arbeidsområde for notater, KI og ekte filer. Dette repositoriet publiserer kildekoden til redigeringsverktøyene: de samme som MR Office bruker, og som du selv kan bygge, kjøre og bygge inn. Dette er redigeringsverktøy med åpen kildekode, ikke en komplett kontorpakke med åpen kildekode: lesing og skriving av Microsoft Office-filer blir værende i MR Office (se nedenfor).

## Redigeringsverktøyene

Hva hvert verktøy gjør i dette repositoriet:

- **Dokumenter** — et sidebasert tekstbehandlingsprogram: stiler, skrifter, lister, tabeller, topp- og bunntekster, inndelinger og spalter, figurer og formler, forsider og innholdsfortegnelser, kommentarer og sporede endringer, samt håndskrift på siden. Lagrer i eget format (.mrdoc); åpner HTML, Markdown og ren tekst.
- **Regneark** — celler og formler, tallformater, betinget formatering, sortering og filtrering, frosne ruter og flere ark, med sum og gjennomsnitt for markeringen. Regning mellom celler (=A1*B1+10) beregnes i nettleseren mens du skriver; formler med funksjoner (SUMMER, HVIS, FINN.RAD …) beholdes slik de er skrevet. Lagrer i eget format (.mrsheet); åpner og lagrer CSV og TSV.
- **Presentasjoner** — lysbilder med åtte egne temaer, oppsett, lysbildemal- og oppsettvisning, tekst, figurer, tekstbokser, tabeller, WordArt og formler, kommentarer, inndelinger, overganger og animasjoner med forhåndsvisning, og lysbildefremvisning. Lagrer i eget format (.mrslides).
- **PDF** — lese, søke i, presentere og skrive ut PDF-filer, kopiere og lese opp teksten, og markere valgt tekst (utheving, understreking, gjennomstreking, bølgelinje) med notater og et kommentarpanel mens filen er åpen; lagring gir tilbake filen slik den ble åpnet.
- **Håndskrift** — håndskrevne notater med trykkfølsomme penner, markeringspenner, viskelær, lasso og linjal, og håndskrift som blir figurer. Lagrer i eget format (.mrink).
- **Kode** — et lett koderedigeringsverktøy med syntaksutheving for mange språk, en disposisjon og søk og erstatt; JavaScript/TypeScript, Python, SQL, Lua og skall kjører rett i nettleseren.

Beregning av formler med funksjoner, tegning av diagrammer og oppsett av SmartArt gjøres av MR Offices egen motor i MR Offices skrivebordsapp: redigeringsprogrammene med åpen kildekode beholder formlene, diagrammene og SmartArt i en fil og viser SmartArt slik det ble lagret.

## Word-, Excel-, PowerPoint- og PDF-filer

Å åpne og lagre Word-, Excel- og PowerPoint-filer (.docx, .xlsx, .pptx) og lagre endringer i en PDF tilbake i filen gjøres av MR Offices gratis skrivebordsapp. Den har alle disse redigeringsprogrammene og fungerer frakoblet på din egen datamaskin:

- **MR Office** (gratis skrivebordsapp): https://github.com/XinaTian/MROffice/releases
- **Nettsted**: https://mroffice.ai

[Alle versjoner →](https://github.com/XinaTian/MROffice/releases)

## Kjør redigeringsverktøyene

De seks verktøyene ligger i `packages/` (`@mrnote/docs`, `sheets`, `slides`, `pdf`, `ink`, `code`, som deler `@mrnote/core`) og kjører sammen i `apps/playground`:

```bash
npm install
npm run fonts     # the open substitutes for Office's fonts
npm run build
npm run preview   # http://localhost:5178
```

`npm run dev` starter en utviklingsserver på samme adresse, og `npm run smoke` kontrollerer et bygg. README-filen for hver pakke beskriver i detalj hva den gjør og hvordan den bygges inn.

## Lisens

Åpen kildekode under GNU Affero General Public License v3.0 (LICENSE), med tilleggsvilkårene i NOTICE.md: behold navngivelsen «MR Office» og lenken til https://mroffice.ai, og ingen rettigheter til navnet eller logoen MR Office. En kommersiell lisens fås via support@mr-note.com. Programvare og data fra tredjeparter står i THIRD_PARTY_NOTICES.md. Bidrag: se CONTRIBUTING.md.

## Mer

Nettsted og hele produktet: **https://mroffice.ai** · Spørsmål og tilbakemeldinger: support@mr-note.com
