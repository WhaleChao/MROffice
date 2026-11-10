/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Character-level (per-RUN) inline text formatting for presentation slide text
 * (PowerPoint parity). A phrase can be bold while the rest is normal; one line
 * can mix font size / colour / italic / underline / strikethrough.
 *
 * A presentation text block (or one of its list paragraphs) historically carries
 * a single flat `text` string plus a block/paragraph-level style applied to the
 * WHOLE span. This module adds an OPTIONAL, migration-safe `runs` array: each run
 * is `{ text, style? }` where `style` is a per-run SUBSET of the block style
 * (bold / italic / underline / strike / superscript / subscript / color /
 * highlight / fontSize / fontFamily). When a
 * block/paragraph carries no `runs` it renders exactly as before (a single span
 * with the block/paragraph style) — the array is additive and fully backward
 * compatible.
 *
 * THE INVARIANT: `textFromRuns(runs) === text`. The flat `text` is always the
 * concatenation of its runs' text, so every non-run-aware consumer (AI context,
 * export fallback, plain-text search) keeps working unchanged.
 *
 * Keep this module dependency-free (no React / no DOM) so it can be unit tested
 * in isolation and imported from the zod schema, the editor, and the render edge
 * alike.
 */

import { sanitizeBlockHref } from "./presentation-block-fx";
import { normalizeLinkTooltip } from "./presentation-link-settings";
import type { PresentationUnderlineStyle } from "./presentation-text-decoration";
import { normalizeOfficeTextPaint, type OfficeTextPaint } from "./office-text-paint";

/** A per-run style subset. Every field is optional; an empty style is dropped. */
export interface RunStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  /**
   * `a:rPr@u="dbl"`: the underline is drawn double (PPT-B363). A table cell's
   * run may carry any DrawingML kind (`presentation-text-decoration`).
   */
  underlineStyle?: PresentationUnderlineStyle;
  strike?: boolean;
  /** A table cell run's double strikethrough (`strike="dblStrike"`). */
  strikeStyle?: "double";
  /** `a:rPr@cap`: the text DISPLAYS in capitals; the text itself is kept (PPT-B362). */
  caps?: "all" | "small";
  /**
   * Baseline shift (PowerPoint parity). At most ONE of `superscript` /
   * `subscript` is ever set on a run — applying one CLEARS the other (they are
   * mutually exclusive baselines), enforced by {@link normalizeRunStyle}.
   */
  superscript?: boolean;
  subscript?: boolean;
  /** DrawingML proofing language (`a:rPr/@lang`) for this exact run. */
  proofingLanguage?: string;
  /** DrawingML `a:rPr/@noProof`; true disables checking for this exact run. */
  noProof?: boolean;
  color?: string;
  /**
   * Text HIGHLIGHT (background) colour — the "文字高亮底色" run background, a
   * #rrggbb hex. Distinct from the present-mode laser highlighter (an ephemeral
   * overlay): this is a persisted per-run background painted behind the glyphs.
   */
  highlight?: string;
  fontSize?: number;
  fontFamily?: string;
  /**
   * The face follows the theme: `fontFamily` is the theme's heading
   * ("heading", `a:latin typeface="+mj-lt"`) or body ("body", `+mn-lt`) font
   * as it last resolved, and a theme change (Design > Themes, Variants >
   * Fonts) gives the run the new theme's, as PowerPoint does for text set in
   * a theme font. Without it `fontFamily` is a face picked by name, which
   * stays whatever the theme (Microsoft: a theme's fonts are the font list's
   * first two entries; a specific font is not updated). The importer states
   * it (PPTX reading row 438); a face set by name clears it. Kept only beside
   * a `fontFamily` -- the spreadsheet's `themeFontRole`, one name for one
   * meaning.
   */
  themeFontRole?: "heading" | "body";
  






  symbolFont?: string;
  /**
   * Character-level text OUTLINE (stroke) colour — the per-run glyph stroke, a
   * #rrggbb hex. DISTINCT from `color` (the glyph fill): rendered via the
   * standard `-webkit-text-stroke` at a fixed 1px width. Additive — a run
   * without it paints no stroke.
   */
  textOutline?: string;
  /**
   * WordArt on these characters: their fill, outline and effects (DrawingML
   * states them on `a:rPr` -- `a:ln`, the fill and effect groups), so one
   * text box holds several WordArt styles, as PowerPoint applies a style,
   * Text Fill, Text Outline and Text Effects to the selected characters
   * (Microsoft Support, "Add or delete a fill, outline, or effect for text or
   * WordArt"). A block's own `textPaint` is what a run without one wears.
   * 3-D and Transform stay the text body's (`a:bodyPr`).
   */
  textPaint?: OfficeTextPaint;
  /**
   * Character Spacing on these characters, in points (`a:rPr@spc`, hundredths
   * of a point; Home > Character Spacing and the Font dialog's Spacing page
   * act on the selected characters, as PowerPoint's do). 0 is stated: the
   * characters of a box spaced as a whole set back to Normal. A block's own
   * `letterSpacing` is what a run without one wears.
   */
  letterSpacing?: number;
  /**
   * External or same-deck slide click target carried by this exact character range. PowerPoint
   * stores it on `<a:rPr><a:hlinkClick>` rather than on the text box itself.
   * Only http(s)/mailto and `slide://<stable-slide-id>` targets survive
   * {@link normalizeRunHyperlink}; unsafe action strings are ignored at every
   * load/render edge.
   */
  href?: string;
  /**
   * The link's ScreenTip (`a:hlinkClick@tooltip`, Insert Hyperlink >
   * ScreenTip...): kept only beside an `href` (PPT-B010).
   */
  linkTooltip?: string;
  /**
   * The link draws in its text's own colour instead of the theme's Hyperlink
   * colour (Office's `ahyp:hlinkClr val="tx"` on the run's `a:hlinkClick`):
   * kept only beside an `href`.
   */
  linkColor?: "text";
}

/**
 * The IDENTITY of a DrawingML text field (`a:fld`) — a slide number, a date or
 * a footer. The run's `text` beside it is only PowerPoint's CACHE of what the
 * field last drew: real PowerPoint refreshes that cache when it opens the file
 * (a probe whose cached text said `1` read back `Page 2 of 3` on slide 2), so
 * the identity is what has to survive a round trip and the text is what to
 * display (docs/office-graphics/SHAPE-IMPORT-AUDIT.md §17).
 */
export interface RunField {
  /** `a:fld@type` — `slidenum`, `datetime1`, … */
  type: string;
  /** `a:fld@id` — the braced GUID the package gave this field. */
  id: string;
}

/**
 * The equation a run IS (PowerPoint's Insert > Equation, Office Math `a14:m`).
 * The run's `text` beside it is the characters the equation draws in the line,
 * as the server reads them from the Office Math it writes (`POST
 * /api/v1/presentation/equation-text` answers them for a LaTeX); `latex` is the
 * equation itself, what the editor renders and edits. `display` says the zone
 * stands as its own line (`m:oMathPara`), as PowerPoint writes an equation
 * inserted into an empty paragraph.
 */
export interface RunMath {
  latex: string;
  display?: boolean;
}

/** The page's bound on one equation's LaTeX (the server's `UniversalRun`). */
export const MAX_RUN_MATH_LATEX_LENGTH = 65_536;

/** A single styled run within a paragraph (or the flat block text). */
export interface TextRun {
  text: string;
  /** Per-run overrides; absent / empty means "inherit the block/paragraph style". */
  style?: RunStyle;
  /**
   * Present only on the run an `a:fld` drew. A run carrying it is ATOMIC: it
   * never merges with a neighbour and never collapses away, because the marker
   * names that exact field and PowerPoint keeps the field while the text on
   * either SIDE of it is edited. Rewriting the run's own text is what DROPS the
   * field — that is PowerPoint's own behaviour, measured (§17), not a rule of
   * ours.
   */
  field?: RunField;
  /**
   * Present only on the run that is an equation. Atomic exactly like a field:
   * the zone is ONE element around the characters it draws, edited through the
   * equation editor and never by typing into it; cutting its characters drops
   * the equation to the plain text PowerPoint shows when a zone is broken.
   */
  math?: RunMath;
}

/** Whether a run is the atomic, un-mergeable kind (it carries a field). */
function runCarriesField(run: TextRun | undefined): boolean {
  const field = run?.field;
  return (
    !!field &&
    typeof field.type === "string" &&
    field.type.length > 0 &&
    typeof field.id === "string" &&
    field.id.length > 0
  );
}

/** Whether a run is an equation (`math` with a non-blank LaTeX). */
export function runCarriesMath(run: TextRun | undefined): boolean {
  const math = run?.math;
  return !!math && typeof math.latex === "string" && math.latex.trim().length > 0;
}

/**
 * Whether a run is atomic: a field or an equation. Such a run never merges
 * with a neighbour, never collapses to flat text, and loses its marker when a
 * cut goes through its characters.
 */
export function runIsAtomic(run: TextRun | undefined): boolean {
  return runCarriesField(run) || runCarriesMath(run);
}

/** The atomic markers of `run`, to copy onto a piece that is still the whole run. */
function atomicMarkers(run: TextRun): Pick<TextRun, "field" | "math"> {
  const markers: Pick<TextRun, "field" | "math"> = {};
  if (runCarriesField(run)) markers.field = run.field;
  if (runCarriesMath(run)) markers.math = run.math;
  return markers;
}

/**
 * Per-run text cap + max runs per paragraph/block (mirrors the schema). A text
 * box's run may be as long as its paragraph, which is as long as its whole text
 * (PPT-B392); a host that bounds a run more tightly passes `maxTextLength`.
 */
export const MAX_RUN_TEXT_LENGTH = 20_000;
/**
 * The run bound a spreadsheet cell's inline rich text keeps (its own schema and
 * the server's cell validator); the presentation's text-box lift is not its.
 */
// Excel's bound on a cell's text: a rich string Excel saved was refused here

export const MAX_CELL_RUN_TEXT_LENGTH = 32_767;



export const MAX_RUNS_PER_BLOCK = 1_000;
/** A spreadsheet cell's inline rich text keeps its own bound. */
export const MAX_CELL_RUNS = 200;
/** A run's character spacing bound, points either way (the block's own,
 * `presentation-line-spacing` MIN/MAX_LETTER_SPACING_PT). */
export const RUN_LETTER_SPACING_MAX_PT = 200;
/**
 * Bounds for a per-run font size, in its block's unit: DrawingML's 4000 pt
 * ceiling as design px (an imported box's points stay below it too). The
 * former 400 turned 500 pt typed into 300 pt in an authored box.
 */
export const MIN_RUN_FONT_SIZE = 1;
export const MAX_RUN_FONT_SIZE = 5333.33;
/** Step applied by the per-run font-size +/- stepper buttons (px). */
export const RUN_FONT_SIZE_STEP = 1;
/** Fallback size the +/- stepper assumes when a selection carries no explicit size. */
export const DEFAULT_RUN_FONT_SIZE = 16;

/**
 * Precision policy for persisted run font sizes.
 *
 * Presentation keeps the historical integer-pixel default. Spreadsheet rich
 * text opts into hundredths because SpreadsheetML stores point sizes as a
 * `double`; converting points to CSS pixels commonly produces values such as
 * 14.67 or 33.33 that must survive a round trip back to 11pt or 25pt.
 */
export interface RunNormalizationOptions {
  readonly fontSizePrecision?: "integer" | "hundredth";
  /** Run count cap; `MAX_RUNS_PER_BLOCK` when omitted. */
  readonly maxRuns?: number;
  /** Per-run text cap; `MAX_RUN_TEXT_LENGTH` when omitted. */
  readonly maxTextLength?: number;
}

/**
 * A curated, whitelisted per-run font family. Like the deck-theme whitelist, each
 * `css` is a COMPLETE, self-contained `font-family` stack (system fallbacks +
 * generic family) so a hand-edited deck can never smuggle an arbitrary
 * `font-family` value / remote `@import` through a run. `label` is the human name
 * shown in the dropdown; an empty `css` is never produced here — the picker maps
 * its "default / inherit" option to clearing the override (`fontFamily: ""`),
 * which {@link applyStylePatch} treats as "remove this key".
 *
 * Covers the system default, the common web-safe Latin faces, and the major CJK
 * (Simplified / Traditional Chinese, Japanese, Korean) families so a multilingual
 * deck can pick a face that actually carries the glyphs.
 */
export interface RunFontFamilyOption {
  /** Stable id used as the <select> option value + React key. */
  readonly id: string;
  /** Human label (also used as the i18n-free display string). */
  readonly label: string;
  /** A complete, safe CSS `font-family` stack to store on the run. */
  readonly css: string;
}

/** The family an option stands for: the first family of its CSS stack. */
export function runFontFamilyName(option: RunFontFamilyOption): string {
  return option.css.split(",")[0].trim().replace(/^['"]|['"]$/g, "");
}

export const RUN_FONT_FAMILIES: readonly RunFontFamilyOption[] = [
  {
    id: "system",
    label: "System default",
    css: "system-ui, -apple-system, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
  },
  { id: "arial", label: "Arial", css: "Arial, 'Helvetica Neue', Helvetica, sans-serif" },
  {
    id: "helvetica",
    label: "Helvetica",
    css: "'Helvetica Neue', Helvetica, Arial, sans-serif",
  },
  {
    id: "verdana",
    label: "Verdana",
    css: "Verdana, Geneva, 'DejaVu Sans', sans-serif",
  },
  {
    id: "tahoma",
    label: "Tahoma",
    css: "Tahoma, Geneva, Verdana, sans-serif",
  },
  {
    id: "trebuchet",
    label: "Trebuchet MS",
    css: "'Trebuchet MS', 'Segoe UI', Tahoma, sans-serif",
  },
  {
    id: "calibri",
    label: "Calibri",
    css: "Calibri, 'Segoe UI', Candara, 'Trebuchet MS', sans-serif",
  },
  {
    id: "times",
    label: "Times New Roman",
    css: "'Times New Roman', Times, serif",
  },
  { id: "georgia", label: "Georgia", css: "Georgia, Cambria, 'Times New Roman', serif" },
  {
    id: "cambria",
    label: "Cambria",
    css: "Cambria, Georgia, 'Times New Roman', serif",
  },
  {
    id: "garamond",
    label: "Garamond",
    css: "Garamond, 'Hoefler Text', 'Times New Roman', serif",
  },
  {
    id: "courier",
    label: "Courier New",
    css: "'Courier New', Courier, 'SFMono-Regular', Consolas, monospace",
  },
  {
    id: "consolas",
    label: "Consolas",
    css: "Consolas, 'SFMono-Regular', Menlo, 'Liberation Mono', monospace",
  },
  // CJK — Simplified Chinese
  {
    id: "pingfang-sc",
    label: "PingFang SC (苹方)",
    css: "'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', 'Noto Sans CJK SC', sans-serif",
  },
  {
    id: "microsoft-yahei",
    label: "Microsoft YaHei (微软雅黑)",
    css: "'Microsoft YaHei', 'PingFang SC', 'Noto Sans CJK SC', sans-serif",
  },
  {
    id: "simsun",
    label: "SimSun (宋体)",
    css: "SimSun, 'Songti SC', 'Noto Serif CJK SC', serif",
  },
  // CJK — Traditional Chinese
  {
    id: "pingfang-tc",
    label: "PingFang TC (蘋方)",
    css: "'PingFang TC', 'Microsoft JhengHei', 'Noto Sans CJK TC', sans-serif",
  },
  {
    id: "microsoft-jhenghei",
    label: "Microsoft JhengHei (微軟正黑體)",
    css: "'Microsoft JhengHei', 'PingFang TC', 'Noto Sans CJK TC', sans-serif",
  },
  // CJK — Japanese
  {
    id: "hiragino-kaku",
    label: "Hiragino Kaku Gothic (ヒラギノ)",
    css: "'Hiragino Kaku Gothic ProN', 'Yu Gothic', 'Noto Sans CJK JP', 'Meiryo', sans-serif",
  },
  {
    id: "yu-gothic",
    label: "Yu Gothic (游ゴシック)",
    css: "'Yu Gothic', 'Hiragino Kaku Gothic ProN', 'Noto Sans CJK JP', sans-serif",
  },
  // CJK — Korean
  {
    id: "malgun-gothic",
    label: "Malgun Gothic (맑은 고딕)",
    css: "'Malgun Gothic', 'Apple SD Gothic Neo', 'Noto Sans CJK KR', sans-serif",
  },
] as const;

/** Lookup of allowed run font-family CSS stacks (for matching the active value). */
const RUN_FONT_FAMILY_CSS_SET: ReadonlySet<string> = new Set(
  RUN_FONT_FAMILIES.map((option) => option.css),
);

/**
 * Resolve a (possibly absent / arbitrary) per-run `fontFamily` value onto the id
 * of the matching {@link RUN_FONT_FAMILIES} option, or `""` when it is absent or
 * unrecognized (so the dropdown shows "default / inherit"). Pure; used to reflect
 * the current selection's family back into the <select>.
 */
export function runFontFamilyId(css: string | undefined): string {
  if (typeof css !== "string" || !css) return "";
  const match = RUN_FONT_FAMILIES.find((option) => option.css === css);
  return match ? match.id : "";
}

/**
 * Map a font-family OPTION ID back to its CSS stack to store on the run. An empty
 * / unknown id resolves to `""`, which {@link applyStylePatch} treats as "clear
 * the override" (the run inherits the block / theme family). Pure.
 */
export function runFontFamilyCss(id: string | undefined): string {
  if (typeof id !== "string" || !id) return "";
  const match = RUN_FONT_FAMILIES.find((option) => option.id === id);
  return match ? match.css : "";
}

/** True when a CSS string is one of the whitelisted run font stacks. */
export function isAllowedRunFontFamily(css: unknown): boolean {
  return typeof css === "string" && RUN_FONT_FAMILY_CSS_SET.has(css);
}

/**
 * Parse a per-run font-size text input into a value for {@link applyRunFormat}'s
 * `fontSize` patch. An empty / whitespace-only field yields `undefined` (which
 * {@link applyStylePatch} treats as "clear the override" → inherit the block /
 * theme size). A non-finite value (`NaN`, `Infinity`) also yields `undefined`
 * so the control never persists a degenerate size. Any other value is rounded
 * and clamped to `[MIN_RUN_FONT_SIZE, MAX_RUN_FONT_SIZE]`. Pure.
 */
export function parseRunFontSizeInput(raw: string, options?: RunNormalizationOptions): number | undefined {
  if (typeof raw !== "string" || raw.trim() === "") return undefined;
  const token = raw.trim().toLowerCase();
  // Number("") is 0, and Number accepts exponent/hex/infinity spellings that
  // are not valid Office font-size field input. Authenticate the complete
  // decimal token plus its optional adjacent px suffix before converting it.
  const match = /^(\d+(?:\.\d+)?|\.\d+)(?:px)?$/.exec(token);
  if (!match) return undefined;
  const numericToken = match[1];
  const value = Number(numericToken);
  if (!Number.isFinite(value)) return undefined;
  return clampRunFontSize(value, options);
}

/** Round + clamp a numeric font size into the supported per-run range. */
export function clampRunFontSize(value: number, options?: RunNormalizationOptions): number {
  const scale = options?.fontSizePrecision === "hundredth" ? 100 : 1;
  return Math.min(MAX_RUN_FONT_SIZE, Math.max(MIN_RUN_FONT_SIZE, Math.round(value * scale) / scale));
}

/**
 * Compute the next per-run font size for a +/- stepper. `current` is the size of
 * the active selection (or `undefined` when it carries no explicit size, in which
 * case we step from {@link DEFAULT_RUN_FONT_SIZE} so the first click is sensible).
 * The result is always rounded + clamped to the supported range. Pure.
 */
export function stepRunFontSize(
  current: number | undefined,
  delta: number,
  options?: RunNormalizationOptions,
): number {
  const base =
    typeof current === "number" && Number.isFinite(current)
      ? current
      : DEFAULT_RUN_FONT_SIZE;
  return clampRunFontSize(base + delta, options);
}

const HEX_COLOR_RE = /^#[0-9a-fA-F]{6}$/;
const PROOFING_LANGUAGE_RE = /^[A-Za-z]{2,8}(?:-[A-Za-z0-9]{1,8}){0,5}$/;
const MAX_RUN_HYPERLINK_LENGTH = 2_048;

export type RunSlideAction = "first" | "last" | "next" | "previous";
export function runSlideActionTarget(value: unknown): RunSlideAction | null {
  return (["first", "last", "next", "previous"] as const).find(
    (target) => value === `slide-action://${target}`,
  ) ?? null;
}

/** Resolve at activation time, never freeze a relative link to a slide ID. */
export function runSlideActionIndex(value: unknown, sourceIndex: number, slideCount: number): number | null {
  const target = runSlideActionTarget(value);
  if (!target || !Number.isInteger(sourceIndex) || !Number.isInteger(slideCount) || sourceIndex < 0 || sourceIndex >= slideCount) return null;
  const index = target === "first" ? 0 : target === "last" ? slideCount - 1 : sourceIndex + (target === "next" ? 1 : -1);
  return index >= 0 && index < slideCount ? index : null;
}









export function shownSlideActionIndex(
  target: RunSlideAction,
  current: number,
  count: number,
  hidden: (index: number) => boolean,
): number {
  if (!Number.isInteger(count) || count <= 0) return -1;
  const scan = (from: number, step: 1 | -1) => {
    for (let index = from; index >= 0 && index < count; index += step) {
      if (!hidden(index)) return index;
    }
    return -1;
  };
  if (target === "first") return scan(0, 1);
  if (target === "last") return scan(count - 1, -1);
  return target === "next" ? scan(current + 1, 1) : scan(current - 1, -1);
}

/** Normalize one external or same-deck run hyperlink. */
export function normalizeRunHyperlink(
  value: unknown,
  knownSlideIds?: ReadonlySet<string>,
): string | undefined {
  if (runSlideActionTarget(value)) return value as string;
  const href = sanitizeBlockHref(value, knownSlideIds);
  return href && href.length <= MAX_RUN_HYPERLINK_LENGTH ? href : undefined;
}
/**
 * The whole hyperlink a collapsed caret stands in (PowerPoint: Ctrl+K with the
 * cursor in a link edits that link, its full text in Text to Display). `lines`
 * are the block's paragraphs (or its one flat run list) in text order, joined
 * by one newline each; `href` is the link the caret reads. The span is the
 * caret's run and its neighbours carrying the same link, within its line; a
 * caret at a link's edge belongs to the link it touches. Null when the caret
 * is in no run carrying `href`.
 */
export function hyperlinkSpanAtCaret(
  lines: ReadonlyArray<{ readonly text: string; readonly runs: readonly TextRun[] }>,
  caret: number,
  href: string,
): { start: number; end: number } | null {
  let base = 0;
  for (const line of lines) {
    const lineEnd = base + line.text.length;
    if (caret >= base && caret <= lineEnd) {
      let cursor = base;
      const spans = line.runs.map((run) => {
        const start = cursor;
        cursor += run.text.length;
        return { start, end: cursor, linked: run.style?.href === href && run.text.length > 0 };
      });
      let at = spans.findIndex((span) => span.linked && span.start < caret && caret < span.end);
      if (at < 0) at = spans.findIndex((span) => span.linked && span.end === caret);
      if (at < 0) at = spans.findIndex((span) => span.linked && span.start === caret);
      if (at < 0) return null;
      let first = at;
      let last = at;
      while (first > 0 && spans[first - 1]!.linked) first -= 1;
      while (last < spans.length - 1 && spans[last + 1]!.linked) last += 1;
      return { start: spans[first]!.start, end: spans[last]!.end };
    }
    base = lineEnd + 1;
  }
  return null;
}

/**
 * The ordered set of boolean style flags, used for normalize + equality.
 * `superscript`/`subscript` are mutually-exclusive baselines: both may appear
 * here for equality, but {@link normalizeRunStyle} drops `subscript` whenever
 * `superscript` is also set so a run never carries both at once.
 */
const BOOLEAN_STYLE_KEYS = [
  "bold",
  "italic",
  "underline",
  "strike",
  "superscript",
  "subscript",
  "noProof",
] as const;

/**
 * Normalize an arbitrary value into a clean {@link RunStyle}, or `undefined`
 * when it carries no usable formatting. Boolean flags survive only when `true`
 * (a `false`/absent flag is indistinguishable from "no override", so it is
 * dropped — this keeps style equality and the flat-collapse rule total). A
 * colour / highlight must be a #rrggbb hex; a font size must be a finite number,
 * rounded and clamped to the supported range; a font family is a bounded
 * non-empty string. `superscript` and `subscript` are mutually exclusive — when
 * both arrive `true`, superscript wins (subscript is dropped) so the run carries
 * a single, well-defined baseline (mirrors PowerPoint's toggle behaviour).
 */
export function normalizeRunStyle(
  value: unknown,
  options?: RunNormalizationOptions,
): RunStyle | undefined {
  if (!value || typeof value !== "object") return undefined;
  const record = value as Record<string, unknown>;
  const style: RunStyle = {};
  for (const key of BOOLEAN_STYLE_KEYS) {
    if (record[key] === true) style[key] = true;
  }
  // Mutually-exclusive baselines: never persist both. Superscript wins.
  if (style.superscript && style.subscript) delete style.subscript;
  if (record.underlineStyle === "double" && style.underline) style.underlineStyle = "double";
  if (record.caps === "all" || record.caps === "small") style.caps = record.caps;
  if (typeof record.color === "string" && HEX_COLOR_RE.test(record.color)) {
    style.color = record.color.toLowerCase();
  }
  if (typeof record.highlight === "string" && HEX_COLOR_RE.test(record.highlight)) {
    style.highlight = record.highlight.toLowerCase();
  }
  if (typeof record.textOutline === "string" && HEX_COLOR_RE.test(record.textOutline)) {
    style.textOutline = record.textOutline.toLowerCase();
  }
  if (typeof record.fontSize === "number" && Number.isFinite(record.fontSize)) {
    const bounded = Math.min(
      MAX_RUN_FONT_SIZE,
      Math.max(MIN_RUN_FONT_SIZE, record.fontSize),
    );
    style.fontSize =
      options?.fontSizePrecision === "hundredth"
        ? Math.round(bounded * 100) / 100
        : Math.round(bounded);
  }
  if (typeof record.fontFamily === "string" && record.fontFamily.trim()) {
    style.fontFamily = record.fontFamily.slice(0, 200);
    if (record.themeFontRole === "heading" || record.themeFontRole === "body") style.themeFontRole = record.themeFontRole;
  }
  if (
    typeof record.symbolFont === "string" && record.symbolFont.trim() && record.symbolFont.length <= 64 &&
    !record.symbolFont.startsWith("+") && !/[<>\u0000-\u001f\\"';{}]/.test(record.symbolFont)
  ) {
    style.symbolFont = record.symbolFont.trim();
  }
  if (
    typeof record.proofingLanguage === "string" &&
    record.proofingLanguage.length <= 64 &&
    PROOFING_LANGUAGE_RE.test(record.proofingLanguage.trim())
  ) {
    style.proofingLanguage = record.proofingLanguage.trim();
  }
  if (typeof record.letterSpacing === "number" && Number.isFinite(record.letterSpacing)) {
    // The block's own bounds (`MIN/MAX_LETTER_SPACING_PT`), hundredths of a point.
    style.letterSpacing =
      Math.round(Math.min(RUN_LETTER_SPACING_MAX_PT, Math.max(-RUN_LETTER_SPACING_MAX_PT, record.letterSpacing)) * 100) / 100;
  }
  if (record.textPaint != null) {
    // A malformed paint is dropped, never thrown: a run's style is read at
    // every render and save edge.
    try {
      const paint = normalizeOfficeTextPaint(record.textPaint);
      if (paint) style.textPaint = paint;
    } catch {
      /* not a paint */
    }
  }
  const href = normalizeRunHyperlink(record.href);
  if (href) style.href = href;
  const linkTooltip = href ? normalizeLinkTooltip(record.linkTooltip) : undefined;
  if (linkTooltip) style.linkTooltip = linkTooltip;
  if (href && record.linkColor === "text") style.linkColor = "text";
  return Object.keys(style).length > 0 ? style : undefined;
}

/** Structural equality of two (possibly absent) run styles. */
export function runStyleEquals(a: RunStyle | undefined, b: RunStyle | undefined): boolean {
  const left = a ?? {};
  const right = b ?? {};
  return (
    BOOLEAN_STYLE_KEYS.every((key) => !!left[key] === !!right[key]) &&
    // Display formats of a run (PPT-B362/B363): neighbours that differ in them
    
    (left.caps ?? undefined) === (right.caps ?? undefined) &&
    (left.underlineStyle ?? undefined) === (right.underlineStyle ?? undefined) &&
    (left.color ?? undefined) === (right.color ?? undefined) &&
    (left.highlight ?? undefined) === (right.highlight ?? undefined) &&
    (left.fontSize ?? undefined) === (right.fontSize ?? undefined) &&
    (left.fontFamily ?? undefined) === (right.fontFamily ?? undefined) &&
    (left.themeFontRole ?? undefined) === (right.themeFontRole ?? undefined) &&
    (left.symbolFont ?? undefined) === (right.symbolFont ?? undefined) &&
    (left.textOutline ?? undefined) === (right.textOutline ?? undefined) &&
    (left.proofingLanguage ?? undefined) ===
      (right.proofingLanguage ?? undefined) &&
    (left.href ?? undefined) === (right.href ?? undefined) &&
    (left.linkTooltip ?? undefined) === (right.linkTooltip ?? undefined) &&
    (left.linkColor ?? undefined) === (right.linkColor ?? undefined) &&
    (left.letterSpacing ?? undefined) === (right.letterSpacing ?? undefined) &&
    sameRunTextPaint(left.textPaint, right.textPaint)
  );
}

/** Whether two runs wear the same WordArt (none on both is the same). */
export function sameRunTextPaint(
  a: OfficeTextPaint | undefined,
  b: OfficeTextPaint | undefined,
): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  return JSON.stringify(a) === JSON.stringify(b);
}

/** Structural equality after canonical run normalization. */
export function runCollectionsEqual(
  a: readonly TextRun[] | undefined,
  b: readonly TextRun[] | undefined,
  options?: RunNormalizationOptions,
): boolean {
  const left = normalizeRuns(a, options);
  const right = normalizeRuns(b, options);
  if (!left || !right) return left === right;
  return (
    left.length === right.length &&
    left.every(
      (run, index) =>
        run.text === right[index]?.text &&
        runStyleEquals(run.style, right[index]?.style),
    )
  );
}

export type ParagraphRunsLike = Readonly<{
  text: string;
  runs?: readonly TextRun[];
}>;

export type ParagraphRunSelectionRange = Readonly<{
  paragraphIndex: number;
  from: number;
  to: number;
  runs: readonly TextRun[];
}>;

/**
 * Map a selection in newline-joined block text onto the exact character ranges
 * of its paragraphs. Newline separators carry no character formatting and are
 * therefore skipped. Pure and bounded by the already-bounded paragraph list.
 */
export function paragraphRunSelectionRanges(
  paragraphs: readonly ParagraphRunsLike[],
  start: number,
  end: number,
  options?: RunNormalizationOptions,
): ParagraphRunSelectionRange[] {
  const total = paragraphs.reduce(
    (length, paragraph, index) =>
      length + String(paragraph.text ?? "").length + (index > 0 ? 1 : 0),
    0,
  );
  const from = Math.min(Math.max(Math.round(start), 0), total);
  const to = Math.min(Math.max(Math.round(end), 0), total);
  if (from >= to) return [];
  const ranges: ParagraphRunSelectionRange[] = [];
  let paragraphStart = 0;
  paragraphs.forEach((paragraph, paragraphIndex) => {
    const text = String(paragraph.text ?? "");
    const paragraphEnd = paragraphStart + text.length;
    const localFrom = Math.max(from, paragraphStart) - paragraphStart;
    const localTo = Math.min(to, paragraphEnd) - paragraphStart;
    if (localFrom < localTo) {
      ranges.push({
        paragraphIndex,
        from: localFrom,
        to: localTo,
        runs: normalizeRuns(paragraph.runs, options) ?? runsFromText(text),
      });
    }
    paragraphStart = paragraphEnd + 1;
  });
  return ranges;
}

/** The flat text obtained by concatenating every run's text (the invariant). */
export function textFromRuns(runs: readonly TextRun[] | undefined): string {
  if (!runs || runs.length === 0) return "";
  return runs.map((run) => run.text).join("");
}

/** Wrap a flat string as a single unstyled run (the trivial runs form). */
export function runsFromText(text: string): TextRun[] {
  return [{ text: String(text ?? "") }];
}

/**
 * True when these runs carry no per-run variation the flat `text` + block style
 * cannot represent: a single run (or zero/one effective run) with no style.
 * Such a block renders identically to its plain `text`, so it should not persist
 * a `runs` field (byte-stable / no phantom array). A SINGLE run that carries a
 * style is NOT flat-equivalent (it overrides the block style for the whole span)
 * and is preserved.
 */
export function isFlatEquivalentRuns(
  runs: readonly TextRun[] | undefined,
  options?: RunNormalizationOptions,
): boolean {
  if (!runs || runs.length === 0) return true;
  if (runs.length > 1) return false;
  // A single run that IS a field is never flat-equivalent: collapsing it would
  // throw the field's identity away and leave a plain run holding a stale cached
  // page number. Every one of the 176 fields on the 132-package ruler is exactly
  // this shape — a placeholder whose whole label is the field.
  if (runIsAtomic(runs[0])) return false;
  return normalizeRunStyle(runs[0].style, options) === undefined;
}

/**
 * Normalize an arbitrary persisted value into a bounded `TextRun[]`, or
 * `undefined` when the value carries no real per-run variation. Empty-text runs
 * are dropped; adjacent runs with EQUAL style are merged (so the model never
 * stores redundant splits); count + per-run length are capped. A result that is
 * flat-equivalent (≤1 styled-less run) collapses to `undefined` so a block
 * without real inline structure stays a flat-text block (byte-stable).
 */
export function normalizeRuns(
  value: unknown,
  options?: RunNormalizationOptions,
): TextRun[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const cleaned: TextRun[] = [];
  for (const entry of value) {
    if (cleaned.length >= (options?.maxRuns ?? MAX_RUNS_PER_BLOCK)) break;
    if (!entry || typeof entry !== "object") continue;
    const record = entry as Record<string, unknown>;
    const text =
      typeof record.text === "string"
        ? record.text.slice(0, options?.maxTextLength ?? MAX_RUN_TEXT_LENGTH)
        : "";
    if (text === "") continue;
    const style = normalizeRunStyle(record.style, options);
    const field = normalizeRunField(record.field, text);
    const math = normalizeRunMath(record.math, text);
    const run: TextRun = style ? { text, style } : { text };
    if (field) run.field = field;
    if (math) run.math = math;
    cleaned.push(run);
  }
  const merged = mergeRuns(cleaned, options);
  if (isFlatEquivalentRuns(merged, options)) return undefined;
  return merged;
}

/**
 * Coalesce adjacent runs whose (normalized) style is equal into a single run,
 * dropping empty-text runs. Pure; returns a fresh array. The merged run keeps a
 * shared normalized style (or none). This is the canonical form `applyRunFormat`
 * and `normalizeRuns` reduce to, so equal-styled neighbours never linger.
 */
export function mergeRuns(
  runs: readonly TextRun[],
  options?: RunNormalizationOptions,
): TextRun[] {
  const out: TextRun[] = [];
  for (const run of runs) {
    const text = String(run.text ?? "");
    if (text === "") continue;
    const style = normalizeRunStyle(run.style, options);
    const field = normalizeRunField(run.field, text);
    const math = normalizeRunMath(run.math, text);
    const last = out[out.length - 1];
    if (
      last &&
      // A field run never merges, in either direction. The marker names ONE
      // `a:fld`, and PowerPoint keeps that field while the text on either side
      // of it is rewritten (§17); merging equally styled neighbours into it
      // would swallow their characters into the field's cached text, and
      // merging it into them would lose the field. This is also why the marker
      // lives on the run rather than as an index into the list: two equally
      // styled neighbours becoming one would shift every later index. An
      // equation is the same one element around its drawn characters.
      !field &&
      !math &&
      !runIsAtomic(last) &&
      runStyleEquals(normalizeRunStyle(last.style, options), style)
    ) {
      last.text += text;
      continue;
    }
    const merged: TextRun = style ? { text, style } : { text };
    if (field) merged.field = field;
    if (math) merged.math = math;
    out.push(merged);
  }
  return out;
}

/**
 * Validate one persisted `a:fld` identity marker, or `undefined`. The grammar is
 * the importer's own (`pptx_text_field_structure_refusal`): a bounded DrawingML
 * token for `@type`, a braced GUID for `@id`. A cached text carrying a NEWLINE
 * loses the marker, because the export spells a newline `a:br` and an `a:fld`
 * cannot hold one — which leaves exactly what PowerPoint leaves when a field's
 * visible text is replaced: a plain run.
 */
export function normalizeRunField(
  value: unknown,
  text: string,
): RunField | undefined {
  if (!value || typeof value !== "object") return undefined;
  if (/[\n\r\v]/.test(text)) return undefined;
  const record = value as Record<string, unknown>;
  const type = typeof record.type === "string" ? record.type : "";
  const id = typeof record.id === "string" ? record.id : "";
  if (!/^[A-Za-z][A-Za-z0-9._-]{0,63}$/.test(type)) return undefined;
  if (
    !/^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$/.test(
      id,
    )
  ) {
    return undefined;
  }
  return { type, id };
}

/**
 * Validate one persisted equation marker, or `undefined`: a non-blank LaTeX
 * within the server's bound, an optional display flag. A text carrying a
 * NEWLINE loses the marker, because an `a14:m` cannot hold a line break.
 */
export function normalizeRunMath(value: unknown, text: string): RunMath | undefined {
  if (!value || typeof value !== "object") return undefined;
  if (/[\n\r\v]/.test(text)) return undefined;
  const record = value as Record<string, unknown>;
  const latex = typeof record.latex === "string" ? record.latex : "";
  if (!latex.trim() || latex.length > MAX_RUN_MATH_LATEX_LENGTH) return undefined;
  return record.display === true ? { latex, display: true } : { latex };
}

/**
 * Split a single run at a character offset (relative to that run's text) into
 * `[before, after]`, each carrying a COPY of the run's style. An offset ≤0 or
 * ≥length returns one part + an empty part so callers can splice uniformly; the
 * empty part is dropped by {@link mergeRuns}. Pure; never mutates the input.
 */
export function splitRun(run: TextRun, offset: number): [TextRun, TextRun] {
  const text = String(run.text ?? "");
  const at = Math.min(Math.max(Math.round(offset), 0), text.length);
  const cloneStyle = (): RunStyle | undefined =>
    run.style ? { ...run.style } : undefined;
  const before = cloneStyle();
  const after = cloneStyle();
  const head: TextRun = before
    ? { text: text.slice(0, at), style: before }
    : { text: text.slice(0, at) };
  const tail: TextRun = after
    ? { text: text.slice(at), style: after }
    : { text: text.slice(at) };
  // A field is ONE element around ONE cached string; it cannot be two. A split
  // that leaves the whole text on one side (the caller's uniform `[part, empty]`
  // shape) keeps the field on that side and drops an empty part; a split that
  // really cuts the run through the middle drops the marker from BOTH halves —
  // which is what PowerPoint does to a field whose visible text is rewritten.
  if (runIsAtomic(run)) {
    if (at === 0) Object.assign(tail, atomicMarkers(run));
    else if (at === text.length) Object.assign(head, atomicMarkers(run));
  }
  return [head, tail];
}

/**
 * Apply a style PATCH to the character range `[start, end)` of a runs array,
 * splitting runs at the selection boundaries and re-styling the covered slice,
 * then merging adjacent-equal runs. The selection is clamped to the text length
 * (out-of-range / reversed selections are no-ops). A patch field of `false`
 * (or `undefined` / `""` / `null`) CLEARS that override on the covered runs; any
 * other value SETS it. Pure; returns a fresh, normalized array. Operates on the
 * runs form derived from the flat `text` when the caller has no runs yet (use
 * {@link runsFromText} first).
 *
 * @param runs   current runs (may be a single unstyled run for a flat block).
 * @param start  selection start char offset (inclusive) into `textFromRuns`.
 * @param end    selection end char offset (exclusive).
 * @param patch  partial style overrides to apply over the covered runs.
 */
export function applyRunFormat(
  runs: readonly TextRun[],
  start: number,
  end: number,
  /**
   * The style to apply, or a function of each covered run's own style — for
   * a step that depends on the run (Grow Font steps every run from its own
   * size rather than giving a mixed selection one size).
   */
  patch: Partial<RunStyle> | ((style: RunStyle | undefined) => Partial<RunStyle>),
  options?: RunNormalizationOptions,
): TextRun[] {
  const base = mergeRuns(runs, options);
  const total = textFromRuns(base).length;
  const from = Math.min(Math.max(Math.round(start), 0), total);
  const to = Math.min(Math.max(Math.round(end), 0), total);
  if (from >= to) return base.length ? base : [];
  const out: TextRun[] = [];
  let cursor = 0;
  for (const run of base) {
    const runStart = cursor;
    const runEnd = cursor + run.text.length;
    cursor = runEnd;
    // Run entirely outside the selection — keep as-is.
    if (runEnd <= from || runStart >= to) {
      out.push(run);
      continue;
    }
    // Compute the in-run boundaries of the covered slice.
    const localFrom = Math.max(from, runStart) - runStart;
    const localTo = Math.min(to, runEnd) - runStart;
    const before = run.text.slice(0, localFrom);
    const covered = run.text.slice(localFrom, localTo);
    const after = run.text.slice(localTo);
    const keepStyle = (): RunStyle | undefined =>
      run.style ? { ...run.style } : undefined;
    // Re-formatting a field does NOT end it: PowerPoint keeps the `a:fld` and
    // changes only its `a:rPr`. The marker therefore rides along with the piece
    // that is still the WHOLE run; a selection that cuts the run in two leaves
    // neither piece able to claim the field, and both lose it.
    const withField = (piece: TextRun, text: string): TextRun => {
      if (runIsAtomic(run) && text.length === run.text.length) {
        Object.assign(piece, atomicMarkers(run));
      }
      return piece;
    };
    if (before) {
      const s = keepStyle();
      out.push(withField(s ? { text: before, style: s } : { text: before }, before));
    }
    if (covered) {
      const nextStyle = applyStylePatch(
        run.style,
        typeof patch === "function" ? patch(run.style) : patch,
        options,
      );
      out.push(
        withField(
          nextStyle ? { text: covered, style: nextStyle } : { text: covered },
          covered,
        ),
      );
    }
    if (after) {
      const s = keepStyle();
      out.push(withField(s ? { text: after, style: s } : { text: after }, after));
    }
  }
  return mergeRuns(out, options);
}

/**
 * Make `[start, end)` ONE run carrying `field` (an `a:fld`: a slide number or
 * a date Insert > Slide Number / Date & Time puts at the cursor). The run
 * takes the style of the first run it covers. Pure; out-of-range or empty
 * ranges leave the runs as they are.
 */
export function applyRunField(
  runs: readonly TextRun[],
  start: number,
  end: number,
  field: RunField,
  options?: RunNormalizationOptions,
): TextRun[] {
  const base = mergeRuns(runs, options);
  const total = textFromRuns(base).length;
  const from = Math.min(Math.max(Math.round(start), 0), total);
  const to = Math.min(Math.max(Math.round(end), 0), total);
  if (from >= to) return base;
  const out: TextRun[] = [];
  let covered: TextRun | null = null;
  let cursor = 0;
  for (const run of base) {
    const runStart = cursor;
    const runEnd = cursor + run.text.length;
    cursor = runEnd;
    if (runEnd <= from || runStart >= to) {
      if (runStart >= to && covered) {
        out.push(covered);
        covered = null;
      }
      out.push(run);
      continue;
    }
    const localFrom = Math.max(from, runStart) - runStart;
    const localTo = Math.min(to, runEnd) - runStart;
    const before = run.text.slice(0, localFrom);
    const inside = run.text.slice(localFrom, localTo);
    const after = run.text.slice(localTo);
    const style = run.style ? { ...run.style } : undefined;
    if (before) out.push(style ? { text: before, style } : { text: before });
    if (covered) covered.text += inside;
    else covered = { ...(style ? { text: inside, style } : { text: inside }), field: { ...field } };
    if (after) {
      out.push(covered);
      covered = null;
      out.push(style ? { text: after, style: { ...style } } : { text: after });
    }
  }
  if (covered) out.push(covered);
  return mergeRuns(out, options);
}

/**
 * Read the per-run style that COVERS the selection `[start, end)` (or the run at
 * the collapsed cursor when `start === end`), restricted to a single field via
 * `pick`. Returns that field's value only when EVERY covered run agrees on it,
 * else `undefined` (a mixed selection). Used to reflect the active font family /
 * size back into the picker controls without changing how anything renders.
 *
 * The collapsed-cursor case reads the run immediately BEFORE the caret (or the
 * first run when the caret is at 0), matching the "typing here would inherit
 * that style" intuition. Pure; never mutates the input.
 */
export function runFieldAtSelection<K extends keyof RunStyle>(
  runs: readonly TextRun[],
  start: number,
  end: number,
  pick: K,
  options?: RunNormalizationOptions,
): RunStyle[K] | undefined {
  const base = mergeRuns(runs, options);
  const total = textFromRuns(base).length;
  const from = Math.min(Math.max(Math.round(start), 0), total);
  const to = Math.min(Math.max(Math.round(end), 0), total);
  // Collapsed cursor: read the run that owns the character just before the caret
  // (or the first run at the start of the text).
  if (from >= to) {
    const at = from <= 0 ? 0 : from - 1;
    let cursor = 0;
    for (const run of base) {
      const runEnd = cursor + run.text.length;
      if (at < runEnd) return normalizeRunStyle(run.style, options)?.[pick];
      cursor = runEnd;
    }
    return undefined;
  }
  // Ranged selection: the field is well-defined only when every covered run
  // carries the SAME value (including "all unset" → undefined).
  let cursor = 0;
  let resolved: RunStyle[K] | undefined;
  let seen = false;
  for (const run of base) {
    const runStart = cursor;
    const runEnd = cursor + run.text.length;
    cursor = runEnd;
    if (runEnd <= from || runStart >= to) continue;
    const value = normalizeRunStyle(run.style, options)?.[pick];
    if (!seen) {
      resolved = value;
      seen = true;
    } else if (value !== resolved) {
      return undefined;
    }
  }
  return resolved;
}

export interface RunProofingSelectionState {
  readonly value: Readonly<{
    proofingLanguage?: string;
    noProof: boolean;
  }> | null;
  readonly mixed: boolean;
}

/** Read exact proofing metadata across a non-empty character range. */
export function runProofingAtSelection(
  runs: readonly TextRun[],
  start: number,
  end: number,
): RunProofingSelectionState {
  const base = mergeRuns(runs);
  const total = textFromRuns(base).length;
  const from = Math.min(Math.max(Math.round(start), 0), total);
  const to = Math.min(Math.max(Math.round(end), 0), total);
  if (from >= to) return { value: null, mixed: false };
  let cursor = 0;
  const values = new Map<
    string,
    Readonly<{ proofingLanguage?: string; noProof: boolean }>
  >();
  for (const run of base) {
    const runStart = cursor;
    const runEnd = cursor + run.text.length;
    cursor = runEnd;
    if (runEnd <= from || runStart >= to) continue;
    const style = normalizeRunStyle(run.style);
    const value = Object.freeze({
      ...(style?.proofingLanguage
        ? { proofingLanguage: style.proofingLanguage }
        : {}),
      noProof: style?.noProof === true,
    });
    values.set(`${value.proofingLanguage ?? ""}\u0000${value.noProof}`, value);
  }
  if (values.size === 0) return { value: null, mixed: false };
  if (values.size > 1) return { value: null, mixed: true };
  return { value: values.values().next().value ?? null, mixed: false };
}

/**
 * Merge a style PATCH over a base run style, returning the new (normalized)
 * style or `undefined` when nothing is left. A patch value that is falsy
 * (`false` / `undefined` / `null` / `""`) CLEARS that key; any other value
 * SETS it. Mirrors the block-style toggle semantics in PresentationWindow.
 *
 * Setting one baseline shift CLEARS the other: a patch that turns `superscript`
 * on drops any existing `subscript` (and vice versa) BEFORE normalization, so
 * toggling subscript over a superscript run actually switches the baseline
 * rather than being silently dropped by the mutual-exclusivity rule.
 */
export function applyStylePatch(
  base: RunStyle | undefined,
  patch: Partial<RunStyle>,
  options?: RunNormalizationOptions,
): RunStyle | undefined {
  const next: Record<string, unknown> = { ...(base ?? {}) };
  if (patch.superscript) delete next.subscript;
  if (patch.subscript) delete next.superscript;
  // A face set by name is the run's own: it no longer follows the theme
  // (a patch that hands the run back to a theme font states its role).
  if (Object.hasOwn(patch, "fontFamily") && !Object.hasOwn(patch, "themeFontRole")) delete next.themeFontRole;
  for (const [key, value] of Object.entries(patch)) {
    if (value === false || value === undefined || value === null || value === "") {
      delete next[key];
    } else {
      next[key] = value;
    }
  }
  return normalizeRunStyle(next, options);
}

/** The boolean run-style flags a Ribbon control or shortcut can TOGGLE. */
export type ToggleableRunStyleKey =
  | "bold"
  | "italic"
  | "underline"
  | "strike"
  | "superscript"
  | "subscript";

/**
 * Resolve the PATCH a toggle shortcut (Ctrl+B / Ctrl+I / Ctrl+U) should apply to
 * the run range `[start, end)` for one boolean character-style flag, matching
 * PowerPoint / Word toggle behaviour:
 *
 *   - If EVERY covered run already carries the flag (the selection is uniformly
 *     ON), the patch CLEARS it (`{ [key]: false }` → `applyStylePatch` drops the
 *     key, turning the span back to default).
 *   - Otherwise (the flag is off somewhere, or the selection is mixed), the patch
 *     SETS it (`{ [key]: true }`), so a partially-bold selection becomes fully
 *     bold on the first press.
 *
 * A collapsed cursor (`start === end`) reads the run just before the caret (via
 * {@link runFieldAtSelection}) so toggling with no selection flips the style the
 * next typed character would inherit — `applyRunFormat` then applies it to the
 * whole block (its collapsed-selection contract). Pure; never mutates the input.
 */
export function resolveBooleanRunToggle(
  runs: readonly TextRun[],
  start: number,
  end: number,
  key: ToggleableRunStyleKey,
): Partial<RunStyle> {
  const uniformlyOn = runFieldAtSelection(runs, start, end, key) === true;
  return { [key]: !uniformlyOn };
}

/**
 * `runs` with the leading and trailing whitespace of their joined text cut
 * away, as a host's `text.trim()` cuts its label: the runs still concatenate
 * to the trimmed words, each keeping its own style.
 */
export function trimRunsLikeText(runs: readonly TextRun[]): TextRun[] {
  const text = runs.map((run) => run.text).join("");
  const start = text.length - text.trimStart().length;
  const end = text.trimEnd().length;
  const kept: TextRun[] = [];
  let at = 0;
  for (const run of runs) {
    const from = Math.max(start, at);
    const to = Math.min(end, at + run.text.length);
    if (to > from) kept.push({ ...run, text: run.text.slice(from - at, to - at) });
    at += run.text.length;
  }
  return kept;
}
