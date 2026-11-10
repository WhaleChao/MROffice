/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { StockRole } from "../../components/console/editor/extensions/chartBlockShared";
import { contentLocaleOf, type ContentLocale } from "../i18n/content-locale";

export interface OfficeChartSampleWords {
  /** Office's sample category, "{n}" its number (or letter). */
  readonly category: string;
  /** Office's sample series, "{n}" its number. */
  readonly series: string;
  /** Office's word for a total (a treemap's root). */
  readonly total: string;
  /** The waterfall sample's four steps: opening, income, cost, closing. */
  readonly waterfall: readonly [string, string, string, string];
  /** A stock chart's five columns. */
  readonly stock: Readonly<Record<StockRole, string>>;
}

/** The words a chart whose reader is unknown is born with (see above). */
export const OFFICE_CHART_SAMPLE_FALLBACK_LOCALE: ContentLocale = "zh";

const NBSP = " ";

const WORDS: Readonly<Record<ContentLocale, OfficeChartSampleWords>> = {
  en: {
    category: "Category {n}",
    series: "Series {n}",
    total: "Total",
    waterfall: ["Opening", "Income", "Cost", "Closing"],
    stock: { volume: "Volume", open: "Open", high: "High", low: "Low", close: "Close" },
  },
  zh: {
    category: "类别 {n}",
    series: "系列 {n}",
    total: "总计",
    waterfall: ["期初", "收入", "成本", "期末"],
    stock: { volume: "成交量", open: "开盘", high: "盘高", low: "盘低", close: "收盘" },
  },
  "zh-TW": {
    category: "類別 {n}",
    series: "數列 {n}",
    total: "總計",
    waterfall: ["期初", "收入", "成本", "期末"],
    stock: { volume: "成交量", open: "開盤", high: "最高", low: "最低", close: "收盤" },
  },
  ja: {
    category: "カテゴリ {n}",
    series: "系列 {n}",
    total: "合計",
    waterfall: ["期首", "収入", "費用", "期末"],
    stock: { volume: "出来高", open: "始値", high: "高値", low: "安値", close: "終値" },
  },
  ko: {
    category: "범주 {n}",
    series: "계열 {n}",
    total: "합계",
    waterfall: ["기초", "수입", "비용", "기말"],
    stock: { volume: "거래량", open: "시가", high: "고가", low: "저가", close: "종가" },
  },
  de: {
    category: "Rubrik {n}",
    series: "Datenreihe {n}",
    total: "Gesamt",
    waterfall: ["Anfangsbestand", "Einnahmen", "Kosten", "Endbestand"],
    stock: { volume: "Volumen", open: "Eröffnungskurs", high: "Höchstkurs", low: "Tiefstkurs", close: "Schlusskurs" },
  },
  fr: {
    category: `Catégorie${NBSP}{n}`,
    series: "Série {n}",
    total: "Total",
    waterfall: ["Solde initial", "Recettes", "Coûts", "Solde final"],
    stock: { volume: "Volume", open: "Ouverture", high: "Haut", low: "Bas", close: "Fermeture" },
  },
  es: {
    category: "Categoría {n}",
    series: "Serie {n}",
    total: "Total",
    waterfall: ["Saldo inicial", "Ingresos", "Costes", "Saldo final"],
    stock: { volume: "Volumen", open: "Apertura", high: "Máximo", low: "Mínimo", close: "Cierre" },
  },
  it: {
    category: "Categoria {n}",
    series: "Serie {n}",
    total: "Totale",
    waterfall: ["Saldo iniziale", "Entrate", "Costi", "Saldo finale"],
    stock: { volume: "Volume", open: "Apertura", high: "Massimo", low: "Minimo", close: "Chiusura" },
  },
  nl: {
    category: "Categorie {n}",
    series: "Reeks {n}",
    total: "Totaal",
    waterfall: ["Beginsaldo", "Inkomsten", "Kosten", "Eindsaldo"],
    stock: { volume: "Volume", open: "Open", high: "Hoog", low: "Laag", close: "Slot" },
  },
  "pt-BR": {
    category: "Categoria {n}",
    series: "Série {n}",
    total: "Total",
    waterfall: ["Saldo inicial", "Receitas", "Custos", "Saldo final"],
    stock: { volume: "Volume", open: "Abertura", high: "Máxima", low: "Mínima", close: "Fechamento" },
  },
  da: {
    category: "Kategori {n}",
    series: "Serie {n}",
    total: "Samlet",
    waterfall: ["Primo", "Indtægter", "Omkostninger", "Ultimo"],
    stock: { volume: "Mængde", open: "Åbn", high: "Høj", low: "Lav", close: "Luk" },
  },
  no: {
    category: "Kategori {n}",
    series: "Serie {n}",
    total: "Totalt",
    waterfall: ["Inngående", "Inntekter", "Kostnader", "Utgående"],
    stock: { volume: "Volum", open: "Åpen", high: "Høy", low: "Lav", close: "Lukket" },
  },
  sv: {
    category: "Kategori {n}",
    series: "Serie {n}",
    total: "Totalt",
    waterfall: ["Ingående", "Intäkter", "Kostnader", "Utgående"],
    stock: { volume: "Volym", open: "Öppning", high: "Högt", low: "Lågt", close: "Stängt" },
  },
  fi: {
    category: "Luokka {n}",
    series: "Sarja {n}",
    total: "Summa",
    waterfall: ["Alkusaldo", "Tulot", "Kulut", "Loppusaldo"],
    stock: { volume: "Määrä", open: "Avaus", high: "Korkein", low: "Matalin", close: "Päätös" },
  },
  id: {
    category: "Kategori {n}",
    series: "Seri {n}",
    total: "Total",
    waterfall: ["Saldo awal", "Pendapatan", "Biaya", "Saldo akhir"],
    stock: { volume: "Volume", open: "Pembukaan", high: "Tertinggi", low: "Terendah", close: "Penutupan" },
  },
  th: {
    category: "ประเภท {n}",
    series: "ชุดข้อมูล {n}",
    total: "ผลรวม",
    waterfall: ["ยอดต้นงวด", "รายได้", "ต้นทุน", "ยอดปลายงวด"],
    stock: { volume: "ปริมาณ", open: "เปิด", high: "สูง", low: "ต่ำ", close: "ปิด" },
  },
  vi: {
    category: "Thể loại {n}",
    series: "Chuỗi {n}",
    total: "Tổng",
    waterfall: ["Đầu kỳ", "Thu nhập", "Chi phí", "Cuối kỳ"],
    stock: { volume: "Khối lượng", open: "Mở cửa", high: "Cao", low: "Thấp", close: "Đóng cửa" },
  },
};

/**
 * The sample words for an interface locale (`contentLocaleOf`: "en-GB" reads
 * "en", "es-419" "es", "zh-Hant" "zh-TW", an unknown language English). No
 * locale at all is the factory's fallback (see the module comment).
 */
export function officeChartSampleWords(
  locale: string | null | undefined,
): OfficeChartSampleWords {
  return WORDS[contentLocaleOf(locale, OFFICE_CHART_SAMPLE_FALLBACK_LOCALE)];
}

/** "Category 1" / "类别 1" / "Rubrik 1" ... */
export function officeChartSampleCategory(
  words: OfficeChartSampleWords,
  n: number | string,
): string {
  return words.category.replace("{n}", String(n));
}

/** "Series 1" / "系列 1" / "Datenreihe 1" ... */
export function officeChartSampleSeries(
  words: OfficeChartSampleWords,
  n: number | string,
): string {
  return words.series.replace("{n}", String(n));
}

export { pageInterfaceLocale } from "../i18n/content-locale";

/**
 * Office's placeholder chart title and axis title ("Chart Title" / "Axis
 * Title", 图表标题 / 坐标轴标题 -- the chart engine's strman_201f keys 0 and 1),
 * for a writer outside the React tree (an AI command) that puts them into a
 * chart. The catalogue's `presentation.ui.chart.defaultChartTitle` /
 * `defaultAxisTitle` and the API's `DEFAULT_CHART_TITLE_WORDS` say the same
 * (a test keeps them equal). No locale keeps English, as these commands
 * always wrote.
 */
const TITLE_WORDS: Readonly<Record<ContentLocale, { readonly chartTitle: string; readonly axisTitle: string }>> = {
  en: { chartTitle: "Chart Title", axisTitle: "Axis Title" },
  zh: { chartTitle: "图表标题", axisTitle: "坐标轴标题" },
  "zh-TW": { chartTitle: "圖表標題", axisTitle: "座標軸標題" },
  ja: { chartTitle: "グラフ タイトル", axisTitle: "軸ラベル" },
  ko: { chartTitle: "차트 제목", axisTitle: "축 제목" },
  de: { chartTitle: "Diagrammtitel", axisTitle: "Achsentitel" },
  fr: { chartTitle: "Titre du graphique", axisTitle: "Titre de l’axe" },
  es: { chartTitle: "Título del gráfico", axisTitle: "Título del eje" },
  it: { chartTitle: "Titolo del grafico", axisTitle: "Titolo asse" },
  nl: { chartTitle: "Grafiektitel", axisTitle: "Astitel" },
  "pt-BR": { chartTitle: "Título do Gráfico", axisTitle: "Título do Eixo" },
  da: { chartTitle: "Diagramtitel", axisTitle: "Aksetitel" },
  no: { chartTitle: "Diagramtittel", axisTitle: "Aksetittel" },
  sv: { chartTitle: "Diagramrubrik", axisTitle: "Axelrubrik" },
  fi: { chartTitle: "Kaavion otsikko", axisTitle: "Akselin otsikko" },
  id: { chartTitle: "Judul Bagan", axisTitle: "Judul Sumbu" },
  th: { chartTitle: "ชื่อแผนภูมิ", axisTitle: "ชื่อแกน" },
  vi: { chartTitle: "Tiêu đề biểu đồ", axisTitle: "Tiêu đề trục" },
};

export function officeChartPlaceholderTitles(
  locale: string | null | undefined,
): { readonly chartTitle: string; readonly axisTitle: string } {
  return TITLE_WORDS[contentLocaleOf(locale, "en")];
}

const REGION_FALLBACK: Readonly<Record<string, string>> = { CN: "中国", US: "美国", JP: "日本" };

/**
 * A map chart's sample region named in the reader's language ("China",
 * "中国", "Chine"), from the platform's own CLDR region names; the Chinese
 * name the sample always had when the platform cannot say.
 */
export function officeChartSampleRegionName(
  regionCode: "CN" | "US" | "JP",
  locale: string | null | undefined,
): string {
  if (typeof locale !== "string" || !locale.trim()) return REGION_FALLBACK[regionCode];
  try {
    const names = new Intl.DisplayNames([locale.trim().replace(/_/g, "-")], { type: "region" });
    return names.of(regionCode) || REGION_FALLBACK[regionCode];
  } catch {
    return REGION_FALLBACK[regionCode];
  }
}
