/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ensurePdfEastAsianFace, PDF_EAST_ASIAN_CSS_FAMILY } from "./east-asian-face";
/**
 * Resolve a pdf.js text run's font into a STABLE, web-safe CSS family the inline
 * editor can render — one that matches the original glyph's serif/sans/mono
 * class and (crucially) survives `doc.destroy()`.
 *
 * Why this exists (BUG: "编辑前后字体无法保证相同"):
 * pdf.js sets `textContent.styles[fontName].fontFamily` to
 * `"<loadedName>", <generic>` (e.g. `"g_d0_f1", serif`). The FIRST token is a
 * synthetic @font-face family pdf.js injects into `document.fonts` for the
 * canvas — it is tied to the document and is removed when the page's document is
 * destroyed. The editor destroys the document right after extracting text, so an
 * overlay that rendered that raw family would fall back to the browser default
 * (sans) the moment the box is edited — a visible font change even though the
 * canvas already shows the real glyphs.
 *
 * The SECOND token, however, is a real CSS generic keyword (`serif` /
 * `sans-serif` / `monospace`) that pdf.js derived from the embedded font, and it
 * is persistent. So we classify the run (preferring an explicit recognised base
 * font name from the font object, else that generic keyword, else the loaded
 * name) into serif/sans/mono and map it to a concrete web-safe stack. The class
 * is the SAME one {@link fontClassFromFamily} (the flatten/export side) derives,
 * so the on-screen edit and the exported PDF use a matching face — font stays
 * consistent before and after editing.
 *
 * Pure + framework-free for unit testing (no pdf.js / DOM needed).
 */

export type FontClass = "serif" | "sans" | "mono";

/** Concrete web-safe stacks per class (real faces, not doc-bound synthetics). */
export const FONT_STACK: Record<FontClass, string> = {
  serif: '"Times New Roman", Times, serif',
  sans: "Helvetica, Arial, sans-serif",
  mono: '"Courier New", Courier, monospace',
};

export const CJK_SERIF_STACK =
  'SimSun, "Songti SC", STSong, "Noto Serif CJK SC", "Times New Roman", Times, serif';
export const CJK_SANS_STACK =
  '"Microsoft YaHei", "PingFang SC", SimHei, "Noto Sans CJK SC", Helvetica, Arial, sans-serif';

const CJK_CHAR_RE =
  /[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\u3040-\u30ff\uac00-\ud7af]/u;
const LATIN_CHAR_RE = /[A-Za-z\u00c0-\u024f]/u;

function cjkFontStackFromName(
  name?: string | null,
): { fontClass: FontClass; fontFamily: string } | null {
  const f = (name || "").toLowerCase();
  if (
    /simsun|songti|stsong|noto\s*serif\s*cjk|source\s*han\s*serif|宋体|宋|仿宋|fangsong/.test(
      f,
    )
  ) {
    return { fontClass: "serif", fontFamily: CJK_SERIF_STACK };
  }
  if (
    /simhei|msyh|microsoft\s*yahei|pingfang|hiragino\s*sans|noto\s*sans\s*cjk|source\s*han\s*sans|黑体|雅黑|苹方/.test(
      f,
    )
  ) {
    return { fontClass: "sans", fontFamily: CJK_SANS_STACK };
  }
  return null;
}

/**
 * Classify a font name / family string into serif / sans / mono. Order matters:
 * "sans-serif" contains the substring "serif", so sans is tested BEFORE serif.
 * Unknown / empty -> `null`; use {@link classifyFontName} when a safe default
 * is desired.
 */
const TEX_FONT_CLASS =
  /(^|[^a-z])(?:(cm(?:v|i|sl)?tt|lmmono|sftt|ectt|tctt)|(cmss|lmsans|sfss|ecss)|(cmr|cmbx|cmti|cmsl|cmmi|cmb|cmcsc|cmu|cmsy|cmex|cmfib|cmdunh|lmroman|lmmath|sfrm|sfbx|sfti|sfsl|sfcc|ecrm|ecbx|ecti|ecsl|eccc|tcrm|tcbx))(?=\d|[^a-z]|$)/u;

export function tryClassifyFontName(name?: string | null): FontClass | null {
  const cjk = cjkFontStackFromName(name);
  if (cjk) return cjk.fontClass;
  const f = (name || "").toLowerCase();
  // TeX's faces (Computer Modern, Latin Modern, the EC fonts): their names
  // say no class, and a TeX line written as a text box was set in the sans
  
  const tex = TEX_FONT_CLASS.exec(f);
  if (tex) return tex[2] ? "mono" : tex[3] ? "sans" : "serif";
  if (/mono|courier|consol|nimbusmon/.test(f)) return "mono";
  if (/sans|arial|helvetica|calibri|segoe|verdana|roboto|tahoma|nimbussan/.test(f))
    return "sans";
  // The URW and TeX Gyre clones say no class either: every LaTeX paper set in
  // Times is NimbusRomNo9L, and a part made italic there came out in Arial
  
  if (/serif|times|georgia|roman|garamond|minion|cambria|nimbusrom|termes|pagella|bonum|schola|palatino|bookman|stix|baskerville|caslon|utopia|libertine|century/.test(f))
    return "serif";
  return null;
}

/**
 * Classify a font name / family string into serif / sans / mono. Order matters:
 * "sans-serif" contains the substring "serif", so sans is tested BEFORE serif.
 * Unknown / empty -> "sans" (the safe default face). Kept byte-compatible with
 * `text-box-flatten.ts`'s `fontClassFromFamily` so the editor and the export
 * agree on a class.
 */
export function classifyFontName(name?: string | null): FontClass {
  const classified = tryClassifyFontName(name);
  if (classified) return classified;
  return "sans";
}

export interface FontHints {
  /**
   * The raw `styles[fontName].fontFamily` from pdf.js, e.g. `"g_d0_f1", serif`
   * or a `systemFontInfo.css` stack. The synthetic first token is ignored for
   * resolution but its trailing generic keyword is used as a signal.
   */
  styleFamily?: string | null;
  /**
   * The font object's real name (from `commonObjs`), subset-tag-stripped &
   * lower-cased by the caller when available, e.g. `helvetica`, `times-bold`.
   * The most reliable signal when present.
   */
  fontName?: string | null;
  /** Bold / italic flags resolved from the font object (best effort). */
  bold?: boolean;
  italic?: boolean;
}

export interface ResolvedFont {
  /** A stable, web-safe CSS font-family stack to render the editable box with. */
  fontFamily: string;
  /** The serif / sans / mono class the family belongs to. */
  fontClass: FontClass;
  bold: boolean;
  italic: boolean;
}

/** Strip a PDF subset prefix such as `ABCDEF+Helvetica-Bold`. */
export function stripPdfFontSubsetTag(name?: string | null): string {
  return String(name || "")
    .replace(/^[A-Z]{6}\+/, "")
    .trim();
}

/**
 * Normalize a pdf.js font-object name for persistence. This is the best export
 * hint we have: it often carries the original PostScript/base-14 face even when
 * the CSS family is a document-scoped synthetic name.
 */
export function normalizePdfSourceFontName(name?: string | null): string {
  return stripPdfFontSubsetTag(name).slice(0, 200);
}

/**
 * Build the CSS family used by the live overlay. Prefer the raw pdf.js
 * style-family while the PDF document is alive so embedded fonts can render
 * exactly; append the stable fallback stack so saved edits still degrade
 * predictably if the synthetic font face is unavailable later.
 */
export function pdfDisplayFontFamily(
  styleFamily: string | null | undefined,
  fallbackFamily: string,
  sourceFont?: {
    loadedName?: string;
    fallbackName?: string;
    systemFontInfo?: { css?: string } | null;
  } | null,
): string {
  // A text item's fontName is an object-cache key, not necessarily a loaded
  // FontFace. Follow the renderer's system substitution / loaded-name choice.
  // This alias is display-only and must never enter the stored document.
  const systemFamily = sourceFont?.systemFontInfo?.css?.trim();
  if (systemFamily) return systemFamily;
  if (sourceFont?.loadedName?.trim()) {
    return `${JSON.stringify(sourceFont.loadedName)}, ${sourceFont.fallbackName || fallbackFamily}`;
  }
  const raw = String(styleFamily || "").trim();
  if (!raw) return fallbackFamily;
  const rawLower = raw.toLowerCase();
  if (
    rawLower === "serif" ||
    rawLower === "sans-serif" ||
    rawLower === "monospace"
  ) {
    return `${fallbackFamily}, ${raw}`.slice(0, 400);
  }
  return `${raw}, ${fallbackFamily}`.slice(0, 400);
}

function cjkStackClass(fontFamily?: string | null): FontClass | null {
  if (!fontFamily) return null;
  if (fontFamily === CJK_SERIF_STACK) return "serif";
  if (fontFamily === CJK_SANS_STACK) return "sans";
  return cjkFontStackFromName(fontFamily)?.fontClass ?? null;
}

function cjkStackLatinFallback(fontFamily: string): string {
  return cjkStackClass(fontFamily) === "sans" ? FONT_STACK.sans : FONT_STACK.serif;
}

/**
 * Choose the stable font stack for the actual replacement text, not only for the
 * source run. Some PDFs expose a CJK fallback stack for every run on a CJK page;
 * if a box is pure Latin text, putting SimSun/Songti first visibly changes title
 * glyphs once the overlay becomes active. Keep CJK stacks for text that contains
 * CJK characters, but use the equivalent Latin serif/sans stack for pure Latin.
 */
/** The East Asian face the saved PDF sets a mixed line's CJK in (the server
 * draws characters cp1252 cannot carry in its Noto Sans SC subset, the rest
 * in the box's own face; native_edit_primitives.py `_ManagedTextBoxFaces`). */
export const EAST_ASIAN_FALLBACK_FAMILY = `"${PDF_EAST_ASIAN_CSS_FAMILY}", "Noto Sans SC"`;

/**
 * A Latin stack the browser falls back from to the same East Asian face the
 * PDF uses: Latin stays in the box's face and only the characters it lacks
 * are drawn in Noto Sans SC, as Word and WPS set a mixed line. Without it the
 * draft drew CJK in the system's face while the saved box drew it in Noto.
 * CJK stacks and stacks that already name it are left as they are.
 */
const CP1252_HIGH = new Set("€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ");

/** Whether the text has a character cp1252 cannot carry: the ones the saved
 * PDF draws in the East Asian face. */
export function needsEastAsianFace(text: string): boolean {
  for (const character of text) {
    const code = character.codePointAt(0)!;
    if (code < 0x80 || (code >= 0xa0 && code <= 0xff) || CP1252_HIGH.has(character)) continue;
    return true;
  }
  return false;
}

export function withEastAsianFallback(fontFamily: string, text?: string): string {
  if (!fontFamily || cjkStackClass(fontFamily) || /noto\s*sans\s*sc/i.test(fontFamily)) return fontFamily;
  if (text !== undefined && !needsEastAsianFace(text)) return fontFamily;
  const base = fontFamily;
  // The named face is the PDF's own file, registered on first need: the
  // stack is laid out again when it arrives (a system face until then).
  void ensurePdfEastAsianFace();
  const generic = /,\s*(sans-serif|serif|monospace|cursive|fantasy|system-ui)\s*$/i.exec(base);
  return generic
    ? `${base.slice(0, generic.index)}, ${EAST_ASIAN_FALLBACK_FAMILY}${base.slice(generic.index)}`
    : `${base}, ${EAST_ASIAN_FALLBACK_FAMILY}`;
}

export function fontFamilyForPdfText(
  text: string,
  fontFamily?: string | null,
): string | undefined {
  if (!fontFamily) return undefined;
  if (!cjkStackClass(fontFamily)) return fontFamily;
  return LATIN_CHAR_RE.test(text) && !CJK_CHAR_RE.test(text)
    ? cjkStackLatinFallback(fontFamily)
    : fontFamily;
}

/**
 * Extract the trailing CSS generic keyword from a pdf.js style family. For
 * `"g_d0_f1", serif` this returns `serif`; for a `systemFontInfo.css` stack like
 * `"Times New Roman", serif` it also returns the generic. Returns "" when no
 * recognised generic is present.
 */
export function genericFromStyleFamily(
  styleFamily?: string | null,
): "" | FontClass {
  const f = (styleFamily || "").toLowerCase();
  // Check the WHOLE string for an explicit CSS generic keyword. monospace before
  // serif; sans-serif (contains "serif") before serif.
  if (/\bmonospace\b/.test(f)) return "mono";
  if (/\bsans-serif\b/.test(f)) return "sans";
  if (/\bserif\b/.test(f)) return "serif";
  return "";
}

/**
 * Resolve a run's font into a stable family + class. Resolution order:
 *  1. An explicit recognised font name (from the font object) — most precise.
 *  2. The trailing CSS generic keyword pdf.js put in the style family (persistent
 *     even after the doc is destroyed).
 *  3. Classifying the (doc-bound) loaded name string as a last resort.
 * Always yields a concrete web-safe stack — never a doc-bound synthetic family.
 */
export function resolvePdfFontFamily(hints: FontHints): ResolvedFont {
  const bold = Boolean(hints.bold);
  const italic = Boolean(hints.italic);

  // 1) A real font name is the strongest signal.
  const name = (hints.fontName || "").toLowerCase().trim();
  if (name) {
    const cjk = cjkFontStackFromName(name);
    if (cjk) {
      return {
        fontFamily: cjk.fontFamily,
        fontClass: cjk.fontClass,
        bold,
        italic,
      };
    }
    const cls = tryClassifyFontName(name);
    if (cls) {
      return { fontFamily: FONT_STACK[cls], fontClass: cls, bold, italic };
    }
  }

  // 2) Trust the generic keyword pdf.js baked into the style family.
  const generic = genericFromStyleFamily(hints.styleFamily);
  if (generic) {
    return {
      fontFamily: FONT_STACK[generic],
      fontClass: generic,
      bold,
      italic,
    };
  }

  // 3) Last resort: classify whatever the style family string is (e.g. only a
  // loaded name). Defaults to sans.
  const cls = classifyFontName(hints.styleFamily);
  return { fontFamily: FONT_STACK[cls], fontClass: cls, bold, italic };
}
