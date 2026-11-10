/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { OfficeThemeLabel } from "./office-theme-catalog";
import type {
  WordQuickStyleId,
  WordStyleBorder,
  WordStyleColor,
  WordStyleDefinition,
  WordStyleParagraphProps,
  WordStyleRunProps,
  WordStyleSet,
  WordStyleSetId,
} from "./office-word-style-sets";

/* ------------------------------------------------------------ colour words */

function themed(themeColor: string, adjust?: { tint?: string; shade?: string }): WordStyleColor {
  return {
    themeColor,
    ...(adjust?.tint ? { themeTint: adjust.tint } : {}),
    ...(adjust?.shade ? { themeShade: adjust.shade } : {}),
  };
}

const ACCENT_DEEP = themed("accent1", { shade: "BF" });
const ACCENT = themed("accent1");
const TEXT = themed("text1");
const TEXT_SOFT = themed("text1", { tint: "BF" });
const TEXT_MUTED = themed("text1", { tint: "A6" });
const TEXT2 = themed("text2");
const PAPER = themed("background1");

function rule(color: WordStyleColor, eighths: number, space = 1): WordStyleBorder {
  return {
    val: "single",
    sz: eighths,
    space,
    ...(color.themeColor ? { themeColor: color.themeColor } : {}),
    ...(color.themeTint ? { themeTint: color.themeTint } : {}),
    ...(color.themeShade ? { themeShade: color.themeShade } : {}),
  };
}

function fill(color: WordStyleColor, tint?: string): WordStyleColor {
  return {
    themeFill: color.themeColor,
    ...(tint ? { themeFillTint: tint } : color.themeShade ? { themeFillShade: color.themeShade } : {}),
  };
}

/* ------------------------------------------------------------- the recipe */

type HeadingRule = "none" | "under" | "over-under" | "band" | "bar" | "tint";

interface StyleSetRecipe {
  readonly id: WordStyleSetId;
  readonly label: OfficeThemeLabel;
  /** Body: size (half-points, default 11pt), space after (twips), line (240ths). */
  readonly body: { readonly sz?: number; readonly after: number; readonly line: number };
  /** Headings 1-3: sizes (half-points), colour, and what sets them apart. */
  readonly heading: {
    readonly sizes: readonly [number, number, number];
    readonly color: WordStyleColor;
    readonly bold?: boolean;
    readonly caps?: boolean;
    readonly spacing?: number;
    readonly body?: boolean;
    readonly center?: boolean;
    readonly before: number;
    readonly after: number;
    readonly rule?: HeadingRule;
  };
  readonly title: {
    readonly sz: number;
    readonly color?: WordStyleColor;
    readonly caps?: boolean;
    readonly bold?: boolean;
    readonly spacing?: number;
    readonly center?: boolean;
    readonly rule?: "none" | "under" | "over-under";
  };
  readonly subtitle: {
    readonly sz: number;
    readonly color: WordStyleColor;
    readonly italic?: boolean;
    readonly caps?: boolean;
    readonly spacing?: number;
  };
}

const RECIPES: readonly StyleSetRecipe[] = [
  {
    id: "default",
    label: { en: "Default" },
    body: { after: 160, line: 278 },
    heading: { sizes: [40, 32, 28], color: ACCENT_DEEP, before: 360, after: 80 },
    title: { sz: 56, spacing: -10 },
    subtitle: { sz: 28, color: TEXT_MUTED, spacing: 15 },
  },
  {
    id: "distinctive",
    label: { en: "Distinctive", zh: "独特", "zh-TW": "特殊", ja: "ユニーク", ko: "장식형", fr: "Soigné", es: "Elegante", de: "Elegant", it: "Caratteristico", "pt-BR": "Elegante" },
    body: { after: 200, line: 276 },
    heading: { sizes: [32, 26, 24], color: TEXT2, bold: true, before: 400, after: 120, rule: "under" },
    title: { sz: 52, color: TEXT2, spacing: 5, rule: "under" },
    subtitle: { sz: 26, color: ACCENT_DEEP, italic: true },
  },
  {
    id: "elegant",
    label: { en: "Elegant", zh: "典雅型", "zh-TW": "典雅", ja: "エレガント", ko: "꾸밈형", fr: "Élégant", es: "Elegante", it: "Elegante", "pt-BR": "Elegante" },
    body: { after: 160, line: 288 },
    heading: { sizes: [32, 26, 24], color: TEXT, caps: true, spacing: 20, center: true, before: 480, after: 160 },
    title: { sz: 60, caps: true, spacing: 40, center: true },
    subtitle: { sz: 24, color: TEXT_MUTED, caps: true, spacing: 30 },
  },
  {
    id: "fancy",
    label: { en: "Fancy", zh: "流行", "zh-TW": "新潮", ja: "ファンシー", ko: "꾸밈형", fr: "Recherché", es: "Sofisticado", de: "Ausgefallen", it: "Ricercato", "pt-BR": "Sofisticado" },
    body: { after: 200, line: 288 },
    heading: { sizes: [24, 22, 22], color: PAPER, bold: true, caps: true, spacing: 15, body: true, before: 240, after: 120, rule: "band" },
    title: { sz: 64, color: ACCENT_DEEP, caps: true, spacing: 10 },
    subtitle: { sz: 26, color: ACCENT, caps: true, spacing: 20 },
  },
  {
    id: "formal",
    label: { en: "Formal", zh: "正式", "zh-TW": "正式", ja: "文語体", ko: "정형", fr: "Officiel", de: "Formell", it: "Formale" },
    body: { after: 200, line: 252 },
    heading: { sizes: [28, 24, 22], color: TEXT, caps: true, spacing: 10, center: true, before: 360, after: 120, rule: "under" },
    title: { sz: 56, caps: true, spacing: 20, center: true, rule: "over-under" },
    subtitle: { sz: 24, color: TEXT_SOFT, caps: true, spacing: 20 },
  },
  {
    id: "manuscript",
    label: { en: "Manuscript", zh: "手稿", "zh-TW": "手寫", ja: "手書き", ko: "필사본", fr: "Manuscrit", es: "Manuscrito", de: "Manuskript", it: "Manoscritto", "pt-BR": "Manuscrito" },
    body: { after: 240, line: 480 },
    heading: { sizes: [32, 28, 24], color: TEXT, center: true, before: 240, after: 240 },
    title: { sz: 52, center: true },
    subtitle: { sz: 28, color: TEXT_SOFT, italic: true },
  },
  {
    id: "modern",
    label: { en: "Modern", zh: "现代", "zh-TW": "摩登", ja: "モダン", ko: "현대형", fr: "Moderne", es: "Moderno", it: "Moderno", "pt-BR": "Moderno" },
    body: { after: 200, line: 276 },
    heading: { sizes: [24, 22, 22], color: TEXT, bold: true, caps: true, spacing: 10, body: true, before: 240, after: 120, rule: "tint" },
    title: { sz: 64, color: ACCENT_DEEP, caps: true, bold: true },
    subtitle: { sz: 24, color: TEXT_MUTED, caps: true, spacing: 15 },
  },
  {
    id: "newsprint",
    label: { en: "Newsprint", zh: "新闻纸", "zh-TW": "新聞紙", ja: "ニュース", ko: "신문 용지", fr: "Papier journal", es: "Papel de periódico", de: "Zeitungspapier", it: "Carta di giornale", "pt-BR": "Papel Jornal" },
    body: { after: 200, line: 276 },
    heading: { sizes: [32, 28, 24], color: TEXT, bold: true, before: 360, after: 80 },
    title: { sz: 72, bold: true, spacing: -10, rule: "under" },
    subtitle: { sz: 28, color: TEXT_SOFT, italic: true },
  },
  {
    id: "perspective",
    label: { en: "Perspective", zh: "透视", "zh-TW": "透視圖", ja: "パースペクティブ", ko: "원근감", es: "Perspectiva", de: "Perspektive", it: "Prospettiva", "pt-BR": "Perspectiva" },
    body: { after: 200, line: 276 },
    heading: { sizes: [32, 26, 24], color: ACCENT_DEEP, caps: true, spacing: 10, before: 360, after: 120, rule: "bar" },
    title: { sz: 56, color: TEXT2, caps: true, spacing: 10 },
    subtitle: { sz: 26, color: ACCENT, spacing: 15 },
  },
  {
    id: "simple",
    label: { en: "Simple", zh: "简单", "zh-TW": "簡單", ja: "シンプル", ko: "단순형", es: "Sencillo", de: "Einfach", it: "Semplice", "pt-BR": "Simples" },
    body: { after: 200, line: 276 },
    heading: { sizes: [36, 28, 24], color: ACCENT, before: 320, after: 80 },
    title: { sz: 56, color: ACCENT, bold: true },
    subtitle: { sz: 28, color: TEXT_MUTED },
  },
  {
    id: "thatch",
    label: { en: "Thatch", zh: "茅草", "zh-TW": "茅草", ja: "ストロー", ko: "짚", fr: "Mailles", es: "Paja", de: "Stroh", it: "Intreccio", "pt-BR": "Colmo" },
    body: { after: 200, line: 288 },
    heading: { sizes: [28, 24, 22], color: ACCENT_DEEP, caps: true, spacing: 10, before: 240, after: 120, rule: "over-under" },
    title: { sz: 56, color: ACCENT_DEEP, caps: true, spacing: 20, rule: "over-under" },
    subtitle: { sz: 26, color: TEXT_SOFT, caps: true, spacing: 15 },
  },
  {
    id: "traditional",
    label: { en: "Traditional", zh: "传统", "zh-TW": "傳統", ja: "トラディショナル", ko: "전통형", fr: "Traditionnel", es: "Tradicional", de: "Traditionell", it: "Tradizionale", "pt-BR": "Tradicional" },
    body: { after: 0, line: 240 },
    heading: { sizes: [24, 24, 22], color: TEXT, bold: true, caps: true, before: 240, after: 120, rule: "under" },
    title: { sz: 56, bold: true, center: true },
    subtitle: { sz: 24, color: TEXT_SOFT, italic: true },
  },
  {
    id: "word2010",
    label: { en: "Word 2010" },
    body: { after: 200, line: 276 },
    heading: { sizes: [28, 26, 22], color: ACCENT_DEEP, bold: true, before: 480, after: 0 },
    title: { sz: 52, color: themed("text2", { shade: "BF" }), spacing: 5, rule: "under" },
    subtitle: { sz: 24, color: ACCENT, italic: true, spacing: 15 },
  },
];

/* ------------------------------------------------------------ the builder */

function paragraph(pPr: WordStyleParagraphProps, rPr: WordStyleRunProps, next: string | null = "Normal"): WordStyleDefinition {
  return { type: "paragraph", basedOn: "Normal", next, pPr, rPr };
}

function character(rPr: WordStyleRunProps): WordStyleDefinition {
  return { type: "character", basedOn: "Default Paragraph Font", next: null, pPr: {}, rPr };
}

function headingRule(kind: HeadingRule | undefined): Partial<WordStyleParagraphProps> {
  switch (kind) {
    case "under":
      return { pBdr: { bottom: rule(ACCENT, 8, 1) } };
    case "over-under":
      return { pBdr: { top: rule(ACCENT, 6, 1), bottom: rule(ACCENT, 6, 1) } };
    case "band":
      return { shd: fill(ACCENT_DEEP), pBdr: { top: rule(ACCENT_DEEP, 24, 0), bottom: rule(ACCENT_DEEP, 24, 0) } };
    case "bar":
      return { pBdr: { left: rule(ACCENT, 36, 6) } };
    case "tint":
      return { shd: fill(ACCENT, "33"), pBdr: { left: rule(ACCENT, 24, 4) } };
    default:
      return {};
  }
}

function buildSet(recipe: StyleSetRecipe): WordStyleSet {
  const h = recipe.heading;
  const bodySize = recipe.body.sz ?? 22;
  const font = h.body ? {} : { fontTheme: "major" as const, eastAsiaTheme: "major" as const };
  const headingRun = (sz: number, level: number): WordStyleRunProps => ({
    ...font,
    sz,
    color: h.color,
    ...(h.bold || level === 3 ? { b: true } : {}),
    ...(h.caps ? { caps: true } : {}),
    ...(h.spacing ? { spacing: h.spacing } : {}),
  });
  const headingPara = (level: number, before: number, after: number): WordStyleParagraphProps => ({
    keepNext: true,
    keepLines: true,
    before,
    after,
    outlineLvl: level - 1,
    ...(h.center ? { jc: "center" } : {}),
    ...(level === 1 ? headingRule(h.rule) : {}),
  });
  const lowerHeading = (level: number): WordStyleDefinition =>
    paragraph(
      { keepNext: true, keepLines: true, before: Math.round(h.before / 4), after: 0, outlineLvl: level - 1 },
      {
        ...(level <= 5 ? font : {}),
        color: level <= 5 ? (h.rule === "band" ? ACCENT_DEEP : h.color) : TEXT_MUTED,
        ...(level === 4 || level === 7 ? { i: true } : {}),
        ...(level === 5 || level === 8 ? { caps: true } : {}),
        ...(level === 6 || level === 9 ? { b: true } : {}),
      },
    );
  const titleRule =
    recipe.title.rule === "under"
      ? { pBdr: { bottom: rule(ACCENT, 8, 4) } }
      : recipe.title.rule === "over-under"
        ? { pBdr: { top: rule(ACCENT, 8, 4), bottom: rule(ACCENT, 8, 4) } }
        : {};
  const set: WordStyleSet = {
    id: recipe.id,
    label: recipe.label,
    docDefaults: {
      rPr: { fontTheme: "minor", eastAsiaTheme: "minor", kern: 2, sz: bodySize },
      pPr: { after: recipe.body.after, line: recipe.body.line, lineRule: "auto" },
    },
    styles: {
      normal: { type: "paragraph", basedOn: null, next: null, pPr: {}, rPr: {} },
      noSpacing: { type: "paragraph", basedOn: null, next: null, pPr: { after: 0, line: 240, lineRule: "auto" }, rPr: {} },
      heading1: paragraph(headingPara(1, h.before, h.after), headingRun(h.sizes[0], 1)),
      heading2: paragraph(headingPara(2, Math.round(h.before / 2), h.after), headingRun(h.sizes[1], 2)),
      heading3: paragraph(headingPara(3, Math.round(h.before / 2), h.after), {
        ...headingRun(h.sizes[2], 3),
        ...(h.rule === "band" ? { color: ACCENT_DEEP } : {}),
      }),
      heading4: lowerHeading(4),
      heading5: lowerHeading(5),
      heading6: lowerHeading(6),
      heading7: lowerHeading(7),
      heading8: lowerHeading(8),
      heading9: lowerHeading(9),
      title: paragraph(
        {
          after: 80,
          line: 240,
          lineRule: "auto",
          contextualSpacing: true,
          ...(recipe.title.center ? { jc: "center" } : {}),
          ...titleRule,
        },
        {
          fontTheme: "major",
          eastAsiaTheme: "major",
          kern: 28,
          sz: recipe.title.sz,
          ...(recipe.title.color ? { color: recipe.title.color } : {}),
          ...(recipe.title.caps ? { caps: true } : {}),
          ...(recipe.title.bold ? { b: true } : {}),
          ...(recipe.title.spacing ? { spacing: recipe.title.spacing } : {}),
        },
      ),
      subtitle: paragraph(recipe.title.center ? { jc: "center" } : {}, {
        sz: recipe.subtitle.sz,
        color: recipe.subtitle.color,
        ...(recipe.subtitle.italic ? { i: true } : {}),
        ...(recipe.subtitle.caps ? { caps: true } : {}),
        ...(recipe.subtitle.spacing ? { spacing: recipe.subtitle.spacing } : {}),
      }),
      quote: paragraph({ before: 160, jc: "center" }, { i: true, color: TEXT_SOFT }),
      intenseQuote: paragraph(
        {
          before: 360,
          after: 360,
          jc: "center",
          ind: { left: 864, right: 864 },
          pBdr: { top: rule(ACCENT_DEEP, 4, 10), bottom: rule(ACCENT_DEEP, 4, 10) },
        },
        { i: true, color: ACCENT_DEEP },
      ),
      listParagraph: paragraph({ ind: { left: 720 }, contextualSpacing: true }, {}, null),
      subtleEmphasis: character({ i: true, color: TEXT_SOFT }),
      emphasis: character({ i: true }),
      intenseEmphasis: character({ i: true, color: ACCENT_DEEP }),
      strong: character({ b: true }),
      subtleReference: character({ smallCaps: true, color: TEXT_SOFT }),
      intenseReference: character({ b: true, smallCaps: true, color: ACCENT_DEEP, spacing: 5 }),
      bookTitle: character({ b: true, i: true, spacing: 5 }),
    } satisfies Record<WordQuickStyleId, WordStyleDefinition>,
  };
  return Object.freeze(set);
}

let built: readonly WordStyleSet[] | null = null;

/** Every style set from MRAI's recipe, in the gallery's order. */
export function fallbackWordStyleSets(): readonly WordStyleSet[] {
  built ??= Object.freeze(RECIPES.map(buildSet));
  return built;
}
