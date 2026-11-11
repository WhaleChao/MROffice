# MR Office — Svenska

[English](../README.md) · [English (UK)](../i18n/README.en-GB.md) · [简体中文](../i18n/README.zh.md) · [繁體中文](../i18n/README.zh-TW.md) · [日本語](../i18n/README.ja.md) · [한국어](../i18n/README.ko.md) · [Deutsch](../i18n/README.de.md) · [Français](../i18n/README.fr.md) · [Español](../i18n/README.es.md) · [Español (Latinoamérica)](../i18n/README.es-419.md) · [Português (Brasil)](../i18n/README.pt-BR.md) · [Italiano](../i18n/README.it.md) · [Nederlands](../i18n/README.nl.md) · [Svenska](../i18n/README.sv.md) · [Dansk](../i18n/README.da.md) · [Norsk](../i18n/README.no.md) · [Suomi](../i18n/README.fi.md) · [Bahasa Indonesia](../i18n/README.id.md) · [ไทย](../i18n/README.th.md) · [Tiếng Việt](../i18n/README.vi.md)

> ### ⬇ För alla funktioner: installera MR Offices skrivbordsapp (gratis)
>
> MR Offices skrivbordsapp är kostnadsfri, innehåller de här redigerarna och hela tjänsten och fungerar offline på din egen dator. Den lägger till det som redigerarna med öppen källkod saknar:
>
> - Öppna och spara Word-, Excel- och PowerPoint-filer (.docx, .xlsx, .pptx)
> - Beräkna formler med funktioner (SUMMA, OM, LETARAD …)
> - Diagram
> - Redigering av SmartArt
> - Spara ändringar i en PDF tillbaka i filen
>
> **[Ladda ned MR Office från Releases →](https://github.com/XinaTian/MROffice/releases)** · [eller från mroffice.ai](https://mroffice.ai)

<p align="center"><a href="https://mroffice.ai"><b>🌐 mroffice.ai</b></a> · <a href="https://mroffice.ai">Prova dem i MR Office →</a></p>

> MR Offices redigerare med öppen källkod – dokument, kalkylblad, presentationer, PDF, handskrift och kod – i webbläsaren.

MR Office är en arbetsyta för anteckningar, AI och riktiga filer. Det här arkivet publicerar källkoden till dess redigerare: samma redigerare som MR Office använder, som du själv kan bygga, köra och bädda in. Det är redigerare med öppen källkod, inte ett komplett kontorspaket med öppen källkod: läsning och skrivning av Microsoft Office-filer finns kvar i MR Office (se nedan).

## Redigerarna

Vad varje redigerare gör i det här arkivet:

- **Dokument** — en sidbaserad ordbehandlare: format, typsnitt, listor, tabeller, sidhuvuden och sidfötter, avsnitt och spalter, former och ekvationer, försättssidor och innehållsförteckningar, kommentarer och spårade ändringar, samt pennanteckningar på sidan. Sparar i eget format (.mrdoc); öppnar HTML, Markdown och oformaterad text.
- **Kalkylblad** — celler och formler, talformat, villkorsstyrd formatering, sortering och filtrering, låsta rutor och flera blad, med summa och medelvärde för markeringen. Räkning mellan celler (=A1*B1+10) sker i webbläsaren medan du skriver; formler med funktioner (SUMMA, OM, LETARAD …) behålls som de är skrivna. Sparar i eget format (.mrsheet); öppnar och sparar CSV och TSV.
- **Presentationer** — bilder med åtta egna teman, layouter, bildbakgrunds- och layoutvyer, text, former, textrutor, tabeller, WordArt och ekvationer, kommentarer, avsnitt, övergångar och animeringar med förhandsvisning samt bildspel. Sparar i eget format (.mrslides).
- **PDF** — läsa, söka i, presentera och skriva ut PDF-filer, kopiera och läsa upp texten, och markera vald text (överstrykning, understrykning, genomstrykning, vågig linje) med anteckningar och en kommentarspanel medan filen är öppen; när du sparar lämnas filen tillbaka som den öppnades.
- **Handskrift** — handskrivna anteckningar med tryckkänsliga pennor, överstrykningspennor, suddgummin, lasso och linjal, och penndrag som blir former. Sparar i eget format (.mrink).
- **Kod** — en lätt kodredigerare med syntaxmarkering för många språk, en disposition samt sök och ersätt; JavaScript/TypeScript, Python, SQL, Lua och skal körs direkt i webbläsaren.

Att beräkna formler med funktioner, rita diagram och ställa upp SmartArt görs av MR Offices egen motor, i MR Offices skrivbordsapp: redigerarna med öppen källkod behåller en fils formler, diagram och SmartArt och visar SmartArt som det sparades.

## Word-, Excel-, PowerPoint- och PDF-filer

Att öppna och spara Word-, Excel- och PowerPoint-filer (.docx, .xlsx, .pptx), och att spara ändringar i en PDF tillbaka i filen, sköts av MR Offices kostnadsfria skrivbordsapp. Den innehåller alla de här redigerarna och fungerar offline på din egen dator:

- **MR Office** (kostnadsfri skrivbordsapp): https://github.com/XinaTian/MROffice/releases
- **Webbplats**: https://mroffice.ai

[Alla versioner →](https://github.com/XinaTian/MROffice/releases)

## Kör redigerarna

De sex redigerarna finns i `packages/` (`@mrnote/docs`, `sheets`, `slides`, `pdf`, `ink`, `code`, som delar `@mrnote/core`) och körs tillsammans i `apps/playground`:

```bash
npm install
npm run fonts     # the open substitutes for Office's fonts
npm run build
npm run preview   # http://localhost:5178
```

`npm run dev` startar en utvecklingsserver på samma adress och `npm run smoke` kontrollerar ett bygge. Varje pakets README beskriver i detalj vad det gör och hur det bäddas in.

## Licens

Öppen källkod under GNU Affero General Public License v3.0 (LICENSE), med de ytterligare villkoren i NOTICE.md: behåll angivelsen ”MR Office” och länken till https://mroffice.ai, och inga rättigheter till namnet eller logotypen MR Office. En kommersiell licens finns via support@mr-note.com. Programvara och data från tredje part listas i THIRD_PARTY_NOTICES.md. Bidrag: se CONTRIBUTING.md.

## Mer

Webbplats och hela produkten: **https://mroffice.ai** · Frågor och synpunkter: support@mr-note.com
