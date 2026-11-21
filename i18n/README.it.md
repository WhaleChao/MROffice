# MR Office — Italiano

<p align="center"><sub>🌐 <a href="../README.md">English</a> · <a href="../i18n/README.en-GB.md">English (UK)</a> · <a href="../i18n/README.zh.md">简体中文</a> · <a href="../i18n/README.zh-TW.md">繁體中文</a> · <a href="../i18n/README.ja.md">日本語</a> · <a href="../i18n/README.ko.md">한국어</a> · <a href="../i18n/README.de.md">Deutsch</a> · <a href="../i18n/README.fr.md">Français</a> · <a href="../i18n/README.es.md">Español</a> · <a href="../i18n/README.es-419.md">Español (Latinoamérica)</a> · <a href="../i18n/README.pt-BR.md">Português (Brasil)</a> · <a href="../i18n/README.it.md">Italiano</a> · <a href="../i18n/README.nl.md">Nederlands</a> · <a href="../i18n/README.sv.md">Svenska</a> · <a href="../i18n/README.da.md">Dansk</a> · <a href="../i18n/README.no.md">Norsk</a> · <a href="../i18n/README.fi.md">Suomi</a> · <a href="../i18n/README.id.md">Bahasa Indonesia</a> · <a href="../i18n/README.th.md">ไทย</a> · <a href="../i18n/README.vi.md">Tiếng Việt</a></sub></p>

<p align="center"><b>Gli editor open source di MR Office — documenti, fogli di calcolo, presentazioni, PDF, scrittura a mano e codice — nel browser.</b></p>

<p align="center"><a href="https://mroffice.ai/it"><b>▶ Prova online</b></a> · <a href="https://mroffice.ai/it/office/download"><b>⬇ Scarica</b></a> · <a href="https://mroffice.ai">🌐 mroffice.ai</a><br><a href="https://mroffice.ai/it/developers">🧑‍💻 Sviluppatori</a> · <a href="https://mroffice.ai/it/developers/embed">🧩 Incorpora</a> · <a href="https://mroffice.ai/it/developers/mcp">🤖 MCP per app di IA</a> · <a href="https://www.npmjs.com/package/mroffice-embed">📦 SDK npm</a> · <a href="https://mroffice.ai/it/developers/oauth">🔑 Accedi con MR Office</a></p>

<p align="center"><a href="../LICENSE"><img alt="AGPL-3.0" src="https://img.shields.io/badge/license-AGPL--3.0-blue"></a> <a href="https://github.com/XinaTian/MROffice/releases"><img alt="release" src="https://img.shields.io/github/v/release/XinaTian/MROffice"></a> <a href="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml"><img alt="CI" src="https://github.com/XinaTian/MROffice/actions/workflows/ci.yml/badge.svg"></a> <a href="https://github.com/XinaTian/MROffice/stargazers"><img alt="GitHub stars" src="https://img.shields.io/github/stars/XinaTian/MROffice?style=social"></a></p>

<p align="center"><a href="https://mroffice.ai/it"><img src="../docs/media/mr-office-demo.webp" alt="Gli editor open source di MR Office — documenti, fogli di calcolo, presentazioni, PDF, scrittura a mano e codice — nel browser." width="100%"></a></p>

<p align="center"><a href="https://demo.mroffice.ai/media/mr-office-demo.mp4">🔊 Guarda il tour di 50 secondi con l'audio</a></p>

<table>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/it"><img src="../docs/screenshots/docs.png" alt="Documenti" width="100%"></a><br><b>Documenti</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/it"><img src="../docs/screenshots/sheets.png" alt="Fogli di calcolo" width="100%"></a><br><b>Fogli di calcolo</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/it"><img src="../docs/screenshots/slides.png" alt="Presentazioni" width="100%"></a><br><b>Presentazioni</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/it"><img src="../docs/screenshots/pdf.png" alt="PDF" width="100%"></a><br><b>PDF</b></td>
</tr>
<tr>
<td align="center" width="50%"><a href="https://mroffice.ai/it"><img src="../docs/screenshots/ink.png" alt="Scrittura a mano" width="100%"></a><br><b>Scrittura a mano</b></td>
<td align="center" width="50%"><a href="https://mroffice.ai/it"><img src="../docs/screenshots/code.png" alt="Codice" width="100%"></a><br><b>Codice</b></td>
</tr>
</table>

MR Office è uno spazio di lavoro per note, IA e file reali. Questo repository pubblica il codice sorgente dei suoi editor: gli stessi editor che usa MR Office, che puoi compilare, eseguire e incorporare in autonomia. Sono editor open source, non una suite per ufficio open source completa: la lettura e la scrittura dei file di Microsoft Office restano in MR Office (vedi sotto).

## Per gli sviluppatori

Tutto per gli sviluppatori è su **[MR Office Developers](https://mroffice.ai/it/developers)**: self-service, senza richiesta. Accedi con qualsiasi account MR Office e inizia.

### Incorpora gli editor nella tua app web (consigliato)

MR Office ospita gli editor; il tuo server firma un token di sessione di breve durata e la tua pagina apre l'editor con esso. Modifica veri file Word, Excel, PowerPoint e PDF (.docx, .xlsx, .pptx, .pdf), con modifica condivisa, caricamenti, tutti i formati di esportazione, stampa e salvataggi inviati al tuo server con un callback firmato. Piano gratuito: 1000 sessioni al mese su localhost, 100 sui tuoi domini, fino a 2 app. È un servizio commerciale ospitato, quindi la tua app non è vincolata dalla licenza AGPL di questo repository.

1. Nella cartella del progetto esegui `npx mroffice-embed init` (oppure crea un'app nella console sviluppatori): un clic su "Consenti" nel browser e le chiavi dell'app vengono scritte in `.env.local` / `.env`.
2. Il tuo server crea un token di sessione con `createSessionToken` di `mroffice-embed/server`; la chiave segreta non lascia mai il server.
3. La tua pagina apre l'editor con `open()` di `mroffice-embed` (o lo script `https://mroffice.ai/embed/v1/embed.js`).

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

[Guida all'incorporamento](https://mroffice.ai/it/developers/embed) · [La guida in Markdown (per agenti IA)](https://mroffice.ai/developers/embed.md) · [Prezzi e limiti](https://mroffice.ai/it/developers/pricing) · [Console sviluppatori](https://mroffice.ai/it/developers/console) · [Demo dal vivo](https://mroffice.ai/it/developers/embed/demo)

### App di IA: il server MCP

Le app di IA aprono, leggono, modificano ed esportano documenti nel vero editor di MR Office tramite il suo server MCP: collega l'endpoint ospitato `https://mcp.mroffice.ai/mcp` (accesso con il tuo account MR Office, nessuna chiave) oppure esegui `npx -y mroffice-mcp` sul tuo computer. Il codice sorgente è in `apps/mcp`. [Guida MCP →](https://mroffice.ai/it/developers/mcp)

### Accedi con MR Office (OAuth)

MR Office è un server di autorizzazione OAuth 2.1 (PKCE, documenti di metadati dell'ID client, registrazione dinamica): la tua app può far accedere le persone con il loro account MR Office e agire con gli ambiti che consentono. [Guida OAuth →](https://mroffice.ai/it/developers/oauth)

### Ospita tu stesso gli editor open source

I sei editor di questo repository possono essere integrati nel tuo prodotto con licenza AGPL-3.0: se altre persone usano una versione modificata tramite rete, devi offrire loro il codice sorgente completo e mantenere l'attribuzione "MR Office" (NOTICE.md). Aprono i propri formati, CSV, HTML, Markdown e PDF; i file Word, Excel e PowerPoint non fanno parte degli editor open source (l'incorporamento ospitato qui sopra li supporta). Una licenza commerciale è disponibile scrivendo a support@mr-note.com. [Guida open source →](https://mroffice.ai/it/developers/open-source)

Sono in preparazione pacchetti npm con i nomi `mroffice-core`, `mroffice-docs`, `mroffice-sheets`, `mroffice-slides`, `mroffice-pdf`, `mroffice-ink` e `mroffice-code`; finché non saranno pubblicati, compila gli editor da questo repository (vedi sotto).

## Gli editor

Cosa fa ogni editor in questo repository:

- **Documenti** — un elaboratore di testi a pagine: stili, caratteri, elenchi, tabelle, intestazioni e piè di pagina, sezioni e colonne, forme ed equazioni, frontespizi e sommari, commenti e revisioni, e inchiostro sulla pagina. Salva nel proprio formato (.mrdoc); apre HTML, Markdown e testo semplice.
- **Fogli di calcolo** — celle e formule, formati numerici, formattazione condizionale, ordinamento e filtri, riquadri bloccati e più fogli, con somma e media della selezione. L'aritmetica tra celle (=A1*B1+10) è calcolata nel browser mentre scrivi; le formule con funzioni (SOMMA, SE, CERCA.VERT…) restano come sono scritte. Salva nel proprio formato (.mrsheet); apre e salva CSV e TSV.
- **Presentazioni** — diapositive con otto temi propri, layout, visualizzazioni schema e layout, testo, forme, caselle di testo, tabelle, WordArt ed equazioni, commenti, sezioni, transizioni e animazioni con anteprima, e presentazione. Salva nel proprio formato (.mrslides).
- **PDF** — leggere, cercare, presentare e stampare file PDF, copiarne e leggerne ad alta voce il testo, e annotare il testo selezionato (evidenziato, sottolineato, barrato, ondulato) con note e un riquadro dei commenti finché il file è aperto; il salvataggio restituisce il file così come è stato aperto.
- **Scrittura a mano** — note scritte a mano con penne sensibili alla pressione, evidenziatori, gomme, lazo e righello, e inchiostro trasformato in forme. Salva nel proprio formato (.mrink).
- **Codice** — un editor di codice leggero con evidenziazione della sintassi per molti linguaggi, una struttura e trova e sostituisci; JavaScript/TypeScript, Python, SQL, Lua e shell girano direttamente nel browser.

Il calcolo delle formule con funzioni, il disegno dei grafici e l'impaginazione degli SmartArt sono svolti dal motore di MR Office, nell'app desktop MR Office: gli editor open source conservano le formule, i grafici e gli SmartArt di un file e mostrano gli SmartArt come sono stati salvati.

## Eseguire gli editor

**▶ [Prova la demo online, senza installare nulla](https://mroffice.ai/it)**

I sei editor sono in `packages/` (`@mrnote/docs`, `sheets`, `slides`, `pdf`, `ink`, `code`, che condividono `@mrnote/core`) e girano insieme in `apps/playground`:

```bash
npm install
npm run fonts     # the open substitutes for Office's fonts
npm run build
npm run preview   # http://localhost:5178
```

`npm run dev` avvia un server di sviluppo allo stesso indirizzo e `npm run smoke` verifica una build. Il README di ogni pacchetto descrive nel dettaglio cosa fa e come incorporarlo.

## App desktop

[`apps/desktop`](../apps/desktop) contiene **MR Office Editors**: il codice sorgente della shell dell’app desktop di MR Office (Electron), con una propria identità di app. Con il componente MR Office gratuito delle Releases è l’app desktop di MR Office, con file Word, Excel, PowerPoint e PDF, accesso e IA. Senza il componente esegue solo gli editor open source, che aprono file .mrdoc, .mrsheet, .mrslides e .mrink, oltre a CSV, HTML, Markdown e PDF. Shell ed editor sono codice sorgente AGPL; il componente è compilato, gratuito e non open source. Dopo `npm run build`, avviala con:

```bash
npm run desktop
```

## Per tutte le funzioni, installa l'app desktop di MR Office (gratuita)

L'app desktop MR Office è gratuita, contiene questi editor e il servizio completo e funziona offline sul tuo computer. Aggiunge ciò che gli editor open source non hanno:

- Aprire e salvare file Word, Excel e PowerPoint (.docx, .xlsx, .pptx)
- Calcolare formule con funzioni (SOMMA, SE, CERCA.VERT …)
- Grafici
- Modifica degli SmartArt
- Salvare nel file le modifiche a un PDF

**[Scarica MR Office (gratis) →](https://mroffice.ai/it/office/download)**

### File Word, Excel, PowerPoint e PDF

L'apertura e il salvataggio di file Word, Excel e PowerPoint (.docx, .xlsx, .pptx), e il salvataggio nel file delle modifiche a un PDF, sono offerti dall'app desktop MR Office (gratuita), che contiene tutti questi editor e funziona offline sul tuo computer:

## Come si combinano le parti

```mermaid
flowchart LR
  subgraph OSS["Open source · AGPL-3.0"]
    direction LR
    core["@mrnote/core"] --> E["I sei editor<br/>Documenti · Fogli di calcolo · Presentazioni · PDF · Scrittura a mano · Codice"]
    E --> PG["apps/playground<br/>Browser"]
    E --> DS["apps/desktop<br/>App desktop"]
    MCP["apps/mcp<br/>Server MCP per le app di IA"]
  end
  COMP["Componente MR Office: file Word, Excel e PowerPoint · gratuito, non open source"] -.-> DS
```

## Licenza

Open source con la GNU Affero General Public License v3.0 (LICENSE), con i termini aggiuntivi di NOTICE.md: mantenere l'attribuzione "MR Office" e il link a https://mroffice.ai, nessun diritto sul nome o sul logo MR Office. Una licenza commerciale è disponibile scrivendo a support@mr-note.com. Software e dati di terze parti sono elencati in THIRD_PARTY_NOTICES.md. Contributi: vedi CONTRIBUTING.md.

## Altro

Sito web e prodotto completo: **https://mroffice.ai** · Domande e feedback: support@mr-note.com

<p align="center">Se MR Office ti è utile, una ⭐ su GitHub aiuta altre persone a trovarlo.</p>
