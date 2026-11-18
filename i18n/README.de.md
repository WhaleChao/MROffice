# MR Office — Deutsch

<p align="center"><sub>🌐 <a href="../README.md">English</a> · <a href="../i18n/README.en-GB.md">English (UK)</a> · <a href="../i18n/README.zh.md">简体中文</a> · <a href="../i18n/README.zh-TW.md">繁體中文</a> · <a href="../i18n/README.ja.md">日本語</a> · <a href="../i18n/README.ko.md">한국어</a> · <a href="../i18n/README.de.md">Deutsch</a> · <a href="../i18n/README.fr.md">Français</a> · <a href="../i18n/README.es.md">Español</a> · <a href="../i18n/README.es-419.md">Español (Latinoamérica)</a> · <a href="../i18n/README.pt-BR.md">Português (Brasil)</a> · <a href="../i18n/README.it.md">Italiano</a> · <a href="../i18n/README.nl.md">Nederlands</a> · <a href="../i18n/README.sv.md">Svenska</a> · <a href="../i18n/README.da.md">Dansk</a> · <a href="../i18n/README.no.md">Norsk</a> · <a href="../i18n/README.fi.md">Suomi</a> · <a href="../i18n/README.id.md">Bahasa Indonesia</a> · <a href="../i18n/README.th.md">ไทย</a> · <a href="../i18n/README.vi.md">Tiếng Việt</a></sub></p>

<p align="center"><b>Die Open-Source-Editoren von MR Office – Dokumente, Tabellen, Präsentationen, PDF, Handschrift und Code – im Browser.</b></p>

<p align="center"><a href="https://mroffice.ai/de"><b>▶ Live-Demo ausprobieren – ohne Installation</b></a> · <a href="https://github.com/XinaTian/MROffice/releases">⬇ MR Office unter Releases herunterladen</a> · <a href="https://mroffice.ai">🌐 mroffice.ai</a></p>

<p align="center"><a href="../LICENSE"><img alt="AGPL-3.0" src="https://img.shields.io/badge/license-AGPL--3.0-blue"></a> <a href="https://github.com/XinaTian/MROffice/releases"><img alt="release" src="https://img.shields.io/github/v/release/XinaTian/MROffice"></a> <a href="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml/badge.svg"></a> <a href="https://github.com/XinaTian/MROffice/stargazers"><img alt="GitHub stars" src="https://img.shields.io/github/stars/XinaTian/MROffice?style=social"></a></p>

<p align="center"><a href="https://mroffice.ai/de"><img src="../docs/media/mr-office-demo.webp" alt="Die Open-Source-Editoren von MR Office – Dokumente, Tabellen, Präsentationen, PDF, Handschrift und Code – im Browser." width="100%"></a></p>

<p align="center"><a href="../docs/media/mr-office-demo.mp4">🔊 Die 50-Sekunden-Tour mit Ton ansehen</a></p>

<table>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/de"><img src="../docs/screenshots/docs.png" alt="Dokumente" width="100%"></a><br><b>Dokumente</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/de"><img src="../docs/screenshots/sheets.png" alt="Tabellen" width="100%"></a><br><b>Tabellen</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/de"><img src="../docs/screenshots/slides.png" alt="Präsentationen" width="100%"></a><br><b>Präsentationen</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/de"><img src="../docs/screenshots/pdf.png" alt="PDF" width="100%"></a><br><b>PDF</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/de"><img src="../docs/screenshots/ink.png" alt="Handschrift" width="100%"></a><br><b>Handschrift</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/de"><img src="../docs/screenshots/code.png" alt="Code" width="100%"></a><br><b>Code</b></td>
</tr>
</table>

MR Office ist ein Arbeitsbereich für Notizen, KI und echte Dateien. Dieses Repository veröffentlicht den Quellcode seiner Editoren: dieselben Editoren, die in MR Office laufen, zum Selbstbauen, Ausführen und Einbetten. Es sind Open-Source-Editoren, keine vollständige Open-Source-Office-Suite: Das Lesen und Schreiben von Microsoft-Office-Dateien bleibt in MR Office (siehe unten).

## Die Editoren

Was jeder Editor in diesem Repository kann:

- **Dokumente** — eine seitenbasierte Textverarbeitung: Formatvorlagen, Schriften, Listen, Tabellen, Kopf- und Fußzeilen, Abschnitte und Spalten, Formen und Formeln, Deckblätter und Inhaltsverzeichnisse, Kommentare und Änderungsverfolgung sowie Freihand auf der Seite. Speichert im eigenen Format (.mrdoc); öffnet HTML, Markdown und Text.
- **Tabellen** — Zellen und Formeln, Zahlenformate, bedingte Formatierung, Sortieren und Filtern, fixierte Bereiche und mehrere Blätter, mit Summe und Mittelwert der Auswahl. Rechnen zwischen Zellen (=A1*B1+10) geschieht beim Tippen im Browser; Formeln mit Funktionen (SUMME, WENN, SVERWEIS …) bleiben, wie sie geschrieben sind. Speichert im eigenen Format (.mrsheet); öffnet und speichert CSV und TSV.
- **Präsentationen** — Folien mit acht eigenen Designs, Layouts, Folienmaster- und Layoutansicht, Text, Formen, Textfeldern, Tabellen, WordArt und Formeln, Kommentaren, Abschnitten, Übergängen und Animationen mit Vorschau sowie Bildschirmpräsentation. Speichert im eigenen Format (.mrslides).
- **PDF** — PDF-Dateien lesen, durchsuchen, präsentieren und drucken, Text kopieren und vorlesen lassen und markierten Text, solange die Datei geöffnet ist, hervorheben, unter-, durch- oder wellenförmig unterstreichen, mit Notizen und einem Kommentarbereich; Speichern gibt die Datei so zurück, wie sie geöffnet wurde.
- **Handschrift** — handschriftliche Notizen mit drucksensitiven Stiften, Textmarkern, Radierern, Lasso und Lineal; Freihand wird zu Formen. Speichert im eigenen Format (.mrink).
- **Code** — ein schlanker Code-Editor mit Syntaxhervorhebung für viele Sprachen, Gliederung sowie Suchen und Ersetzen; JavaScript/TypeScript, Python, SQL, Lua und Shell laufen direkt im Browser.

Formeln mit Funktionen berechnen, Diagramme zeichnen und SmartArt anordnen übernimmt die eigene Engine von MR Office in der MR Office Desktop-App: Die Open-Source-Editoren bewahren die Formeln, Diagramme und SmartArt einer Datei und zeigen SmartArt so an, wie es gespeichert wurde.

## Die Editoren ausführen

**▶ [Live-Demo ausprobieren – ohne Installation](https://mroffice.ai/de)**

Die sechs Editoren liegen in `packages/` (`@mrnote/docs`, `sheets`, `slides`, `pdf`, `ink`, `code`, gemeinsam `@mrnote/core`) und laufen zusammen in `apps/playground`:

```bash
npm install
npm run fonts     # the open substitutes for Office's fonts
npm run build
npm run preview   # http://localhost:5178
```

`npm run dev` startet einen Entwicklungsserver unter derselben Adresse, `npm run smoke` prüft einen Build. Die README jedes Pakets beschreibt im Einzelnen, was es kann und wie man es einbettet.

## Desktop-App

[`apps/desktop`](../apps/desktop) enthält **MR Office Editors**: den Quellcode der Hülle der MR-Office-Desktop-App (Electron) mit einer eigenen App-Identität. Mit der kostenlosen MR-Office-Komponente aus den Releases ist es die MR-Office-Desktop-App, mit Word-, Excel-, PowerPoint- und PDF-Dateien, Anmeldung und KI. Ohne die Komponente laufen nur die Open-Source-Editoren; sie öffnen .mrdoc-, .mrsheet-, .mrslides- und .mrink-Dateien sowie CSV, HTML, Markdown und PDF. Hülle und Editoren sind AGPL-Quellcode; die Komponente ist kompiliert, kostenlos nutzbar und nicht Open Source. Nach `npm run build` starten Sie sie mit:

```bash
npm run desktop
```

## Für den vollen Funktionsumfang: die MR Office Desktop-App installieren (kostenlos)

Die MR Office Desktop-App ist kostenlos, enthält diese Editoren und den vollen Dienst und funktioniert offline auf Ihrem eigenen Computer. Sie bietet, was den Open-Source-Editoren fehlt:

- Word-, Excel- und PowerPoint-Dateien (.docx, .xlsx, .pptx) öffnen und speichern
- Formeln mit Funktionen (SUMME, WENN, SVERWEIS …) berechnen
- Diagramme
- SmartArt bearbeiten
- PDF-Änderungen in die Datei zurückschreiben

**[MR Office unter Releases herunterladen →](https://github.com/XinaTian/MROffice/releases)** · [oder auf mroffice.ai](https://mroffice.ai)

### Word-, Excel-, PowerPoint- und PDF-Dateien

Word-, Excel- und PowerPoint-Dateien (.docx, .xlsx, .pptx) öffnen und speichern sowie PDF-Änderungen in die Datei zurückschreiben bietet die kostenlose MR Office Desktop-App. Sie enthält alle diese Editoren und funktioniert offline auf Ihrem eigenen Computer:

## Wie die Teile zusammenpassen

```mermaid
flowchart LR
  subgraph OSS["Open Source · AGPL-3.0"]
    direction LR
    core["@mrnote/core"] --> E["Die sechs Editoren<br/>Dokumente · Tabellen · Präsentationen · PDF · Handschrift · Code"]
    E --> PG["apps/playground<br/>Browser"]
    E --> DS["apps/desktop<br/>Desktop-Shell"]
    MCP["apps/mcp<br/>MCP-Server für KI-Apps"]
  end
  COMP["MR-Office-Komponente: Word-, Excel- und PowerPoint-Dateien · kostenlos, nicht Open Source"] -.-> DS
```

## Lizenz

Open Source unter der GNU Affero General Public License v3.0 (LICENSE) mit den zusätzlichen Bedingungen in NOTICE.md: Die Nennung „MR Office“ und der Link auf https://mroffice.ai bleiben erhalten, Rechte am Namen und Logo von MR Office werden nicht eingeräumt. Eine kommerzielle Lizenz ist über support@mr-note.com erhältlich. Software und Daten Dritter stehen in THIRD_PARTY_NOTICES.md. Beiträge: siehe CONTRIBUTING.md.

## Mehr

Website und das vollständige Produkt: **https://mroffice.ai** · Fragen und Feedback: support@mr-note.com

<p align="center">Wenn dir MR Office nützt, hilft ein ⭐ auf GitHub anderen, es zu finden.</p>
