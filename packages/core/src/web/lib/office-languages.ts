/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The languages MR Office's own words are written in (its home page, its
 * sample files, its page title and description): every site language, the
 * regional variants reading their base (en-GB → en, es-419 → es). Pure: the
 * proxy, route handlers and components read it.
 */
import { pickLocale } from "../i18n/locale-tags";

export type OfficeLanguage =
  | "en" | "zh" | "zh-TW" | "ja" | "ko" | "de" | "fr" | "es" | "it" | "pt-BR"
  | "nl" | "da" | "no" | "sv" | "fi" | "id" | "vi" | "th";

export const OFFICE_LANGUAGES: readonly OfficeLanguage[] = [
  "en", "zh", "zh-TW", "ja", "ko", "de", "fr", "es", "it", "pt-BR",
  "nl", "da", "no", "sv", "fi", "id", "vi", "th",
];

/** The language MR Office writes in for a site locale. */
export function officeLanguage(locale: string): OfficeLanguage {
  // a variant reads its language's words (zh-HK → zh-TW, en-GB → en,
  // pt-PT → pt-BR, nb → no: i18n/locale-tags), anything else English
  return pickLocale(locale, OFFICE_LANGUAGES) ?? "en";
}

/** MR Office's words about itself, in each language: `description`, short,
 * for the installed app and the about page; `title` and `summary`, the home
 * page's name and description in search results and link previews, which
 * say what it opens and what an account adds, as MR Note's do. */
export const OFFICE_META: Record<
  OfficeLanguage,
  { readonly description: string; readonly title: string; readonly summary: string }
> = {
  en: { description: "The open-source editor closest to Microsoft Office: documents, spreadsheets, presentations and PDFs in your browser, with AI to help.",
    title: "MR Office — The open-source editor closest to Microsoft Office: Word, Excel, PowerPoint and PDF online",
    summary: "MR Office is the open-source editor closest to Microsoft Office: documents laid out page by page like Word, workbooks calculated with Excel's functions, decks with PowerPoint's themes and layouts, and .docx, .xlsx, .pptx and PDF files opened and saved as they are. It runs in your browser — no install, no sign-up to start. A free account adds saving, export and an AI assistant." },
  zh: { description: "最接近微软 Office 的开源编辑器：在浏览器里编辑文档、表格、演示文稿和 PDF，AI 随时帮忙。",
    title: "MR Office — 最接近微软 Office 的开源编辑器，在线编辑 Word、Excel、PPT 与 PDF",
    summary: "MR Office 是最接近微软 Office 的开源编辑器：文档像 Word 一样按页排版，表格用 Excel 的函数计算，演示文稿沿用 PowerPoint 的主题和版式，.docx、.xlsx、.pptx 与 PDF 原样打开、原样保存。在浏览器里就能用，不用安装，不注册也能开始；注册免费账号即可保存、导出，并有 AI 助手。" },
  "zh-TW": { description: "在瀏覽器中編輯文件、試算表、簡報和 PDF，AI 隨時幫忙。",
    title: "MR Office — 免費線上 Office，編輯 Word、Excel、PPT 與 PDF",
    summary: "MR Office 在瀏覽器中開啟和編輯 .docx、.xlsx、.pptx 與 PDF：無需安裝，不註冊也能開始；註冊免費帳號即可儲存、匯出，並有 AI 助理幫你寫作、算表和做簡報。" },
  ja: { description: "ブラウザで文書、表計算、プレゼンテーション、PDF を編集。AI がいつでも手伝います。",
    title: "MR Office — Word・Excel・PowerPoint・PDF を編集できる無料オンライン Office",
    summary: "MR Office なら .docx、.xlsx、.pptx、PDF をブラウザで開いて編集できます。インストール不要、登録なしで始められ、無料アカウントで保存・書き出しと、文章作成・表計算・スライド作りを手伝う AI アシスタントが使えます。" },
  ko: { description: "브라우저에서 문서, 스프레드시트, 프레젠테이션, PDF를 편집하고 AI의 도움을 받으세요.",
    title: "MR Office — Word, Excel, PowerPoint, PDF를 편집하는 무료 온라인 Office",
    summary: "MR Office에서 .docx, .xlsx, .pptx, PDF 파일을 브라우저로 열고 편집하세요. 설치도 가입도 없이 바로 시작하고, 무료 계정을 만들면 저장·내보내기와 함께 글쓰기·계산·슬라이드 작업을 돕는 AI 도우미를 쓸 수 있습니다." },
  de: { description: "Dokumente, Tabellen, Präsentationen und PDFs im Browser – mit KI, die hilft.",
    title: "MR Office – kostenloses Online-Office für Word, Excel, PowerPoint und PDF",
    summary: "MR Office öffnet und bearbeitet .docx-, .xlsx-, .pptx- und PDF-Dateien direkt im Browser – ohne Installation und ohne Anmeldung zum Start. Ein kostenloses Konto bringt Speichern, Export und einen KI-Assistenten, der beim Schreiben, Rechnen und Gestalten hilft." },
  fr: { description: "Documents, feuilles de calcul, présentations et PDF dans votre navigateur, avec l’aide de l’IA.",
    title: "MR Office — Office en ligne gratuit pour Word, Excel, PowerPoint et PDF",
    summary: "MR Office ouvre et modifie vos fichiers .docx, .xlsx, .pptx et PDF directement dans le navigateur, sans installation ni inscription pour commencer. Un compte gratuit ajoute l’enregistrement, l’export et un assistant IA qui vous aide à rédiger, calculer et mettre en page." },
  es: { description: "Documentos, hojas de cálculo, presentaciones y PDF en tu navegador, con ayuda de la IA.",
    title: "MR Office — Office en línea gratis para Word, Excel, PowerPoint y PDF",
    summary: "MR Office abre y edita archivos .docx, .xlsx, .pptx y PDF en tu navegador, sin instalar nada ni registrarte para empezar. Una cuenta gratuita añade guardado, exportación y un asistente de IA que te ayuda a redactar, calcular y diseñar." },
  it: { description: "Documenti, fogli di calcolo, presentazioni e PDF nel browser, con l’aiuto dell’IA.",
    title: "MR Office — Office online gratuito per Word, Excel, PowerPoint e PDF",
    summary: "MR Office apre e modifica file .docx, .xlsx, .pptx e PDF direttamente nel browser, senza installazioni né registrazione per iniziare. Un account gratuito aggiunge salvataggio, esportazione e un assistente IA che ti aiuta a scrivere, calcolare e impaginare." },
  "pt-BR": { description: "Documentos, planilhas, apresentações e PDFs no navegador, com a ajuda da IA.",
    title: "MR Office — Office online grátis para Word, Excel, PowerPoint e PDF",
    summary: "O MR Office abre e edita arquivos .docx, .xlsx, .pptx e PDF direto no navegador, sem instalar nada nem se cadastrar para começar. Uma conta gratuita adiciona salvamento, exportação e um assistente de IA que ajuda a escrever, calcular e criar apresentações." },
  nl: { description: "Documenten, spreadsheets, presentaties en pdf’s in je browser, met hulp van AI.",
    title: "MR Office — gratis online Office voor Word, Excel, PowerPoint en pdf",
    summary: "MR Office opent en bewerkt .docx-, .xlsx-, .pptx- en pdf-bestanden direct in je browser, zonder installatie en zonder account om te beginnen. Een gratis account voegt opslaan, exporteren en een AI-assistent toe die helpt bij schrijven, rekenen en vormgeven." },
  da: { description: "Dokumenter, regneark, præsentationer og PDF’er i browseren – med hjælp fra AI.",
    title: "MR Office – gratis online Office til Word, Excel, PowerPoint og PDF",
    summary: "MR Office åbner og redigerer .docx-, .xlsx-, .pptx- og PDF-filer direkte i browseren – uden installation og uden konto for at komme i gang. En gratis konto giver gem, eksport og en AI-assistent, der hjælper med at skrive, regne og designe." },
  no: { description: "Dokumenter, regneark, presentasjoner og PDF-er i nettleseren – med hjelp fra KI.",
    title: "MR Office – gratis nettbasert Office for Word, Excel, PowerPoint og PDF",
    summary: "MR Office åpner og redigerer .docx-, .xlsx-, .pptx- og PDF-filer rett i nettleseren – uten installasjon og uten konto for å komme i gang. En gratis konto gir lagring, eksport og en KI-assistent som hjelper deg å skrive, regne og designe." },
  sv: { description: "Dokument, kalkylblad, presentationer och PDF:er i webbläsaren – med hjälp av AI.",
    title: "MR Office – gratis Office online för Word, Excel, PowerPoint och PDF",
    summary: "MR Office öppnar och redigerar .docx-, .xlsx-, .pptx- och PDF-filer direkt i webbläsaren – utan installation och utan konto för att komma igång. Ett gratiskonto ger sparande, export och en AI-assistent som hjälper dig att skriva, räkna och formge." },
  fi: { description: "Asiakirjat, taulukot, esitykset ja PDF:t selaimessa – tekoäly auttaa.",
    title: "MR Office – ilmainen verkko-Office Wordille, Excelille, PowerPointille ja PDF:lle",
    summary: "MR Office avaa ja muokkaa .docx-, .xlsx-, .pptx- ja PDF-tiedostoja suoraan selaimessa – ilman asennusta ja ilman tiliä alkuun. Ilmainen tili tuo tallennuksen, viennin ja tekoälyavustajan, joka auttaa kirjoittamaan, laskemaan ja suunnittelemaan." },
  id: { description: "Dokumen, spreadsheet, presentasi, dan PDF di browser, dibantu AI.",
    title: "MR Office — Office online gratis untuk Word, Excel, PowerPoint, dan PDF",
    summary: "MR Office membuka dan mengedit file .docx, .xlsx, .pptx, dan PDF langsung di browser, tanpa instalasi dan tanpa daftar untuk mulai. Akun gratis menambahkan penyimpanan, ekspor, dan asisten AI yang membantu Anda menulis, menghitung, dan mendesain." },
  vi: { description: "Tài liệu, bảng tính, bản trình bày và PDF ngay trên trình duyệt, có AI hỗ trợ.",
    title: "MR Office — Office trực tuyến miễn phí cho Word, Excel, PowerPoint và PDF",
    summary: "MR Office mở và chỉnh sửa tệp .docx, .xlsx, .pptx và PDF ngay trên trình duyệt, không cần cài đặt hay đăng ký để bắt đầu. Tài khoản miễn phí cho bạn lưu, xuất tệp và trợ lý AI giúp viết, tính toán và thiết kế." },
  th: { description: "เอกสาร สเปรดชีต งานนำเสนอ และ PDF ในเบราว์เซอร์ พร้อม AI ช่วยเหลือ",
    title: "MR Office — Office ออนไลน์ฟรีสำหรับ Word, Excel, PowerPoint และ PDF",
    summary: "MR Office เปิดและแก้ไขไฟล์ .docx, .xlsx, .pptx และ PDF ได้ในเบราว์เซอร์ ไม่ต้องติดตั้งและไม่ต้องสมัครก็เริ่มได้ บัญชีฟรีเพิ่มการบันทึก การส่งออก และผู้ช่วย AI ที่ช่วยเขียน คำนวณ และออกแบบ" },
};
