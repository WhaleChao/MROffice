/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How a presentation text body is drawn, as CSS: font size/family/line-height
 * of imported and authored text, insets and overflow, run styles, paragraph
 * indents, the shape text frame, and the autofit (shrink text on overflow)
 * measurements the canvas, thumbnails and slideshow all share.
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */

import officeLineMetrics from "@mrnote/core/web/lib/office-fonts/office-line-metrics.generated.json";
import { eastAsianMeasuredFirstBaselineEm } from "@mrnote/core/web/lib/text-engine/east-asian-lines.ts";
import { ptToPx } from "@mrnote/core/web/lib/office-units.ts";
import { officeTextOutlineCss } from "@mrnote/core/web/lib/office-text-outline/office-text-outline-css.ts";
import { runOutlineLine, runOutlineWidthPt } from "@mrnote/core/web/lib/presentation-runs.ts";
import { CSS_PX_PER_POINT, designPxToPoints, pointsToCssPx, storedFontSizeToCssPx, type PresentationFontSizeUnit } from "@mrnote/core/web/lib/presentation/text-units.ts";
import { normalizeDocumentShapeAdjustments } from "@mrnote/core/web/lib/document-shapes.ts";
import type { CSSProperties } from "react";
import {
  EMPTY_PRESENTATION_FONT_ALIASES,
  presentationFontFamilyWithAssets,
  presentationOfficeLightFontWeight,
  type PresentationFontAliases,
} from "@mrnote/core/web/lib/presentation-font-face-reads.ts";
import { officeTextPaintCss, type OfficeTextPaint } from "@mrnote/core/web/lib/office-text-paint.ts";
import {
  normalizeOfficeShapeShadow,
  OFFICE_SHADOW_DEFAULT_DIRECTION,
  type OfficeShapeShadow,
} from "@mrnote/core/web/lib/office-shape-paint.ts";
import { normalizeOfficeTextWarp } from "@mrnote/core/web/lib/office-text-warp.ts";
import { normalizePresentationUnderlineStyle, presentationTextDecorationCss } from "@mrnote/core/web/lib/presentation-text-decoration.ts";
import { smartArtTextFitScale } from "@mrnote/core/web/lib/office-smartart/text-fit.ts";
import {
  type SlideParagraph,
  normalizeParagraphAlign,
  paragraphStatesIndent,
} from "@mrnote/core/web/lib/presentation-paragraphs.ts";
import { type RunStyle, type TextRun, normalizeRuns } from "@mrnote/core/web/lib/presentation-point-runs.ts";
import {
  exactLineSpacingCss,
  pointsToPx,
  readSpacingFromStyle,
  spacingToCss,
} from "@mrnote/core/web/lib/presentation-line-spacing.ts";
import { normalizeShapeKind } from "@mrnote/core/web/lib/presentation-shapes.ts";
import { presentationShapeCatalogEntry } from "@mrnote/core/web/lib/presentation-shape-catalog.ts";
import { readVerticalAnchor, shapeLabelVerticalAlign } from "@mrnote/core/web/lib/presentation-vertical-anchor.ts";
import {
  authoredShapeTextRectFractions,
  shapeTextRectFractions,
  textBlockTextRectFractions,
} from "@mrnote/core/web/lib/presentation-shape-text-rect.ts";
import { normalizeColumnCount } from "@mrnote/core/web/lib/presentation-text-columns.ts";
import { readSmartArtBlockMarker } from "@mrnote/core/web/lib/presentation-smartart.ts";
import {
  geometryNumber,
  presentationBlockIsPlaceholder,
  hasLayoutGeometry,
  isImportedPptxBlock,
  isImportedPptxTextBlock,
} from "@mrnote/core/web/lib/presentation-block-reads.ts";
import {
  autofitModeForBlock,
  normalizedParagraphSpacingFields,
  normalizedSpacingFallbackBlocked,
  normalizeParagraphsWithRuns,
  ownsParagraphRichLayout,
  type SlideParagraphWithRuns,
  presentationTextUsesPoints,
} from "@mrnote/core/web/lib/presentation-text-buttons.ts";
import {
  applyAutofitScale,
  autofitFontScale,
  DEFAULT_AVG_CHAR_EM,
  estimateTextHeight,
  nativeAutofitRender,
  normalizeDrawnAutofitReceipt,
  textOverflowsAutofitFrame,
  type AutofitMetrics,
} from "@mrnote/core/web/lib/presentation-autofit.ts";
import { isTextAlign, textAlignCss } from "@mrnote/core/web/lib/presentation-text-format.ts";
import { autofitStepAtOrBelow, autofitStepsFor } from "../presentation-autofit-ladder";
import {
  measuredAutofit,
  requestAutofitMeasure,
  type AutofitMeasureContext,
} from "./autofit-measure-store";
import type {
  PresentationBlock,
  PresentationShapeBlock,
  PresentationTextBlock,
  PresentationTextStyle,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { asRecord } from "@mrnote/core/web/lib/as-record.ts";
import { cssFontFamilyList } from "@mrnote/core/web/lib/css-font-family.ts";
import {
  wordFontIsEastAsian,
  wordFontLineDescent,
  wordFontLineMetrics,
} from "@mrnote/core/web/components/console/editor/native-document/wordFontMetrics.ts";

/**
 * A shape's WordArt paint as CSS, or null when it has none or the stored paint
 * is not a valid one. The paint's reader refuses a malformed value by throwing,
 * and a render that threw took the whole editor to its error page for one
 * shape (PPT-B122); such a label draws as plain text instead.
 */
export function shapeTextPaintCss(block: PresentationBlock): CSSProperties | null {
  const paint = (block as { textPaint?: unknown }).textPaint;
  if (!paint) return null;
  try {
    return officeTextPaintCss(paint as OfficeTextPaint);
  } catch {
    return null;
  }
}

/**
 * A run's own WordArt paint as CSS, or null (none, or not a valid paint: a
 * malformed one draws as plain text rather than throwing, PPT-B122).
 */










export function presentationTextPaintOwnsShadow(paint: unknown): boolean {
  const shadow = paint && typeof paint === "object" ? (paint as { shadow?: unknown }).shadow : undefined;
  return Boolean(shadow && typeof shadow === "object");
}

export function runTextPaintCss(style: { textPaint?: unknown } | undefined): CSSProperties | null {
  const paint = style?.textPaint;
  if (!paint) return null;
  try {
    return officeTextPaintCss(paint as OfficeTextPaint);
  } catch {
    return null;
  }
}

/**
 * The text-style family lives in TWO layers in the wild: the editor writes
 * `block.style.*`, while the pptx importer lands the same style fields at the
 * block top level (each side's tests pin its own convention, and the exporter
 * now reads both). Converge at LOAD: hoist top-level values under `style`
 * (style's own value wins), so rendering and every ribbon control read one
 * layer. Pure; non-text records pass through untouched.
 */
/**
 * Whether a text box wraps its lines at its edge. `a:bodyPr@wrap="none"`
 * (imported as `style.textWrap: false`) runs the text on past the edge and
 * keeps only its hard breaks (PPT-B360).
 */













export function textBlockTabSizeCss(block: { style?: unknown }): CSSProperties {
  const stops = (block.style as { tabStops?: unknown } | undefined)?.tabStops;
  if (!Array.isArray(stops)) {
    const interval = (block.style as { defaultTabPx?: unknown } | undefined)?.defaultTabPx;
    return typeof interval === "number" && Number.isFinite(interval) && interval > 0 ? { tabSize: `${interval}px` } : {};
  }
  // A left stop is its bare position; another alignment is `{ pos, align }`
  // (laid out by `presentation-tab-stops`, which draws each tab measured).
  const first = stops
    .map((stop) => (typeof stop === "number" ? stop : (stop as { pos?: unknown } | null)?.pos))
    .find((pos): pos is number => typeof pos === "number" && Number.isFinite(pos) && pos > 0);
  return typeof first === "number" ? { tabSize: `${first}px` } : {};
}

/**
 * "Allow Latin text to wrap in the middle of a word" (`a:pPr@latinLnBrk="1"`
 * on every row, imported as `style.latinLineBreak`): PowerPoint breaks a Latin
 * word at any letter, as it already breaks East Asian text (measured, export
 * look probe-lnbrk). Absent, words wrap at spaces as before.
 */
export function textBlockLineBreakCss(block: { style?: unknown }): CSSProperties {
  return (block.style as { latinLineBreak?: unknown } | undefined)?.latinLineBreak === true
    ? { wordBreak: "break-all" }
    : {};
}

export function textBlockWraps(block: { style?: unknown }): boolean {
  return (block.style as { textWrap?: unknown } | undefined)?.textWrap !== false;
}








export function textBlockOverflowWrap(block: { style?: unknown }): "break-word" | "normal" {
  return textBlockWraps(block) ? "break-word" : "normal";
}

export function importedPptxTextOverflow(block: PresentationBlock): "visible" | "hidden" {
  const record = block as Record<string, unknown>;
  const style = asRecord(record.style);
  const raw = style && "verticalOverflow" in style
    ? style.verticalOverflow
    : record.verticalOverflow;
  return raw === "clip" || raw === "ellipsis" ? "hidden" : "visible";
}

/**
 * Whether a text block states its own internal margins (Format Shape > Text
 * Box > Left/Right/Top/Bottom margin). An imported one always draws them; an
 * authored one drew the editor's fixed padding whatever it stated, so 72 pt
 * and 0 pt left margins looked the same (PPT-B115).
 */
export function textBlockStatesInsets(block: PresentationBlock): boolean {
  return asRecord((block as Record<string, unknown>).textInsets) !== null;
}

/**
 * Whether a text block's words are drawn inside margins: its own, or -- an
 * authored text box that states none -- PowerPoint's defaults (0.1" at the
 * sides, 0.05" above and below), which the download and the PDF already give
 * it. The editor padded such a box 14/12 px at a 1.55 line while the show and
 * the thumbnails drew it flush at 1.5, so its words jumped up and left as the
 * show began. A SmartArt label keeps the band the engine sizes for it.
 */
export function textBlockDrawsInsets(block: PresentationBlock): boolean {
  if (textBlockStatesInsets(block)) return true;
  return (
    block.type === "text" &&
    !isImportedPptxBlock(block) &&
    hasLayoutGeometry(block) &&
    readSmartArtBlockMarker(block)?.role !== "label"
  );
}

/** PowerPoint's own text box margins, in EMU (`DEFAULT_PRESENTATION_TEXT_INSETS`). */
const OFFICE_DEFAULT_TEXT_INSETS = { left: 91_440, right: 91_440, top: 45_720, bottom: 45_720 } as const;

/** The padding of a block {@link textBlockDrawsInsets} says draws margins. */
export function textBlockInsetsCss(block: PresentationBlock, containingWidthEmu?: number): CSSProperties {
  return importedPptxTextInsetsCss(
    textBlockStatesInsets(block) || isImportedPptxBlock(block)
      ? block
      : ({ ...block, textInsets: OFFICE_DEFAULT_TEXT_INSETS } as PresentationBlock),
    containingWidthEmu,
  );
}

export function importedPptxTextInsetsCss(block: PresentationBlock, containingWidthEmu?: number): CSSProperties {
  const record = block as Record<string, unknown>;
  const insets = asRecord(record.textInsets);
  const width = geometryNumber(block, "width");
  // The box-wide indent (`style.indentLevel`, half an inch a level, which the
  // import states for a body whose every row has one `a:pPr@marL`) is added
  // to the left inset: DrawingML's body inset and `marL` are additive. The
  // inset padding below used to replace the indent `textBlockStyleCss` gives,
  
  // flush here).
  const rawLevel = asRecord(record.style)?.indentLevel ?? record.indentLevel;
  const indentEmu =
    typeof rawLevel === "number" && Number.isFinite(rawLevel) && rawLevel > 0
      ? Math.min(8, Math.round(rawLevel)) * 457_200
      : 0;
  if (width === null || width <= 0) return { padding: 0 };
  const referenceWidth = containingWidthEmu ?? width;
  // Vertical words (Text Direction > Stacked / Rotate all text, a vertical
  // text box): the frame that pads them writes vertically too, and a
  // browser resolves its percentage padding against the containing block's
  // HEIGHT there -- the canvas against the box's, the show against the
  // slide's -- so the same column sat in two places and neither at its
  // margins. Every surface lays a slide out in design pixels, so the margins
  // are stated in them.
  const writingMode = readSpacingFromStyle(asRecord(record.style) ?? undefined)?.writingMode;
  const vertical = writingMode !== undefined && writingMode !== "horizontal-tb";
  const percent = (value: unknown): string => {
    const numeric =
      typeof value === "number" && Number.isFinite(value) ? value : 0;
    const emu = Math.max(0, Math.min(width, numeric));
    return vertical ? `${emu / AUTOFIT_EMU_PER_PX}px` : `${(emu / referenceWidth) * 100}%`;
  };
  // A text block standing in for a non-rectangular shape lays its words out
  // in the geometry's text rectangle (`source.nativeTextRect`), the insets
  // measured from its edges, as PowerPoint does and as a shape label does
  // (`shapeTextRectFractions`). Absent, the rectangle is the frame.
  const rect = textBlockTextRectFractions(record.source);
  const height = geometryNumber(block, "height") ?? 0;
  const edge = {
    left: rect ? rect.left * width : 0,
    right: rect ? (1 - rect.right) * width : 0,
    top: rect ? rect.top * height : 0,
    bottom: rect ? (1 - rect.bottom) * height : 0,
  };
  const numeric = (value: unknown): number =>
    typeof value === "number" && Number.isFinite(value) ? value : 0;
  if (!insets) {
    if (!rect) {
      return indentEmu ? { padding: 0, paddingLeft: percent(indentEmu) } : { padding: 0 };
    }
    return {
      paddingTop: percent(edge.top),
      paddingRight: percent(edge.right),
      paddingBottom: percent(edge.bottom),
      paddingLeft: percent(edge.left + indentEmu),
    };
  }
  // CSS percentage padding — including top/bottom — is relative to the
  // containing block's inline width. Present mode applies padding directly
  // on an absolute slide child; the editor applies it inside the text frame.
  return {
    paddingTop: percent(edge.top + numeric(insets.top)),
    paddingRight: percent(edge.right + numeric(insets.right)),
    paddingBottom: percent(edge.bottom + numeric(insets.bottom)),
    paddingLeft: percent(edge.left + numeric(insets.left) + indentEmu),
  };
}

/**
 * The width an imported block's words are laid out in, in EMU: the frame (or
 * its geometry's text rectangle) less the insets and the box-wide indent
 * {@link importedPptxTextInsetsCss} pads it by. Null for a block without a
 * width, or one written vertically.
 */
export function importedPptxTextAreaWidthEmu(block: PresentationBlock): number | null {
  const record = block as Record<string, unknown>;
  const width = geometryNumber(block, "width");
  if (width === null || width <= 0) return null;
  const writingMode = readSpacingFromStyle(asRecord(record.style) ?? undefined)?.writingMode;
  if (writingMode !== undefined && writingMode !== "horizontal-tb") return null;
  const rawLevel = asRecord(record.style)?.indentLevel ?? record.indentLevel;
  const indentEmu =
    typeof rawLevel === "number" && Number.isFinite(rawLevel) && rawLevel > 0
      ? Math.min(8, Math.round(rawLevel)) * 457_200
      : 0;
  const rect = textBlockTextRectFractions(record.source);
  const insets = asRecord(record.textInsets);
  const clamp = (value: unknown) => Math.max(0, Math.min(width, typeof value === "number" && Number.isFinite(value) ? value : 0));
  const left = (rect ? rect.left * width : 0) + (insets ? clamp(insets.left) : 0) + indentEmu;
  const right = (rect ? (1 - rect.right) * width : 0) + (insets ? clamp(insets.right) : 0);
  return Math.max(0, width - clamp(left) - clamp(right));
}

export function importedPptxTextHasPreview(block: PresentationBlock): boolean {
  if (block.type !== "text") return false;
  const textBlock = block as PresentationTextBlock;
  const paragraphs = normalizeParagraphsWithRuns(textBlock.paragraphs);
  return Boolean(
    isImportedPptxTextBlock(block) ||
      // ONE question, asked of the one that draws.
      //
      // This gate decides whether the paragraph-aware preview layer is mounted
      // at all; `needsStructuredParagraphRender` (-> `ownsParagraphRichLayout`)
      // decides what that layer draws. They used to disagree: this side asked
      // only about bullets and runs, so a paragraph carrying its own
      // line spacing / space-before / space-after / alignment was stored and
      // then never drawn — the flat <textarea> underneath renders one
      // block-level line-height and cannot show a per-paragraph one.
      //
      // That is PowerPoint's most ordinary paragraph gesture: put the caret in
      // a paragraph, press 行距 ▸ 2.0. `applyParagraphFormatSelection` wrote
      // `lineSpacing` onto that one paragraph (correct — PowerPoint scopes it
      // to the caret's paragraph), the ribbon lit up, and the text on the
      // slide never moved.
      (paragraphs && ownsParagraphRichLayout(paragraphs)) ||
      normalizeRuns(textBlock.runs) ||
      normalizeColumnCount(textBlock.columnCount) > 1 ||
      // A middle or bottom anchor: the bare textarea always draws its text
      // at the top, so Top / Bottom left the canvas unchanged while the show
      // moved the title (PPT-B410). The preview layer places it as the show does.
      readVerticalAnchor(textBlock.style) !== "top",
  );
}

/**
 * PowerPoint for the web's AutoFit smart tag belongs to a real placeholder,
 * not every free text box. Require an explicit placeholder identity from the
 * scratch layout model or imported PresentationML projection; never infer one
 * merely because a block happens to contain text.
 */
function presentationTextBlockIsPlaceholder(
  block: PresentationBlock,
): block is PresentationTextBlock {
  return block.type === "text" && presentationBlockIsPlaceholder(block);
}

/**
 * CSS for a text block's optional formatting (font size / colour / weight).
 *
 * `themed` carries the deck-theme / slide-master contribution (colour, font
 * family, default font size, default bold) resolved via `resolveBlockTextStyle`
 * (or `resolveWithMaster`). It is applied ONLY as a fallback: an explicit
 * per-block `style.*` always wins over the corresponding theme/master value, and
 * a theme/master value applies only when the block has no explicit one. This
 * keeps the precedence block-override > master > theme > built-in default. With
 * no theme + no block style the result is `{}`, exactly as before (themeless
 * decks unchanged) — `resolveBlockTextStyle` returns no font-size / bold when the
 * theme omits them, so an existing deck's title/body sizing is untouched.
 */
export type { PresentationFontSizeUnit } from "@mrnote/core/web/lib/presentation/text-units.ts";

/**
 * DrawingML percentage line spacing is measured from Office's normal text
 * line, whose CJK baseline box is 1.2em. CSS unitless `line-height`, however,
 * measures directly from the font size. Keep authoring/export values untouched
 * and bridge that semantic difference while painting.
 */
export const IMPORTED_PPTX_NORMAL_LINE_HEIGHT = 1.2;

const OFFICE_LINE_FACES: ReadonlyMap<string, readonly number[]> = new Map(
  Object.entries((officeLineMetrics as { faces: Record<string, number[]> }).faces).map(([name, value]) => [name.toLowerCase(), value]),
);

/** A face as CSS reads it (a document's "Wingdings 3" quoted), else undefined. */
function cssFontFamilyOrUndefined(value: string | undefined): string | undefined {
  return value ? cssFontFamilyList(value) : value;
}

/**
 * Imported runs retain their exact source face for patchback/export. At the
 * browser edge, add bounded local fallbacks for the Office CJK faces used by
 * native decks so an unavailable Windows font does not silently fall through
 * to Safari's serif default and change wrapping/baselines.
 */
function importedPptxFontFamilyCss(
  value: string | undefined,
  unit: PresentationFontSizeUnit,
): string | undefined {
  if (!value || unit !== "point") return value;
  const primary = value
    .split(",")[0]
    ?.trim()
    .replace(/^['"]|['"]$/g, "")
    .toLowerCase();
  if (
    primary === "microsoft yahei ui" ||
    primary === "microsoft yahei" ||
    primary === "微软雅黑"
  ) {
    return '"Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", "PingFang SC", "Hiragino Sans GB", sans-serif';
  }
  if (primary === "dengxian light" || primary === "等线 light") {
    // The light family first; its fallbacks are drawn at the light weight
    // the text style asks for (`presentationOfficeLightFontWeight`).
    return '"DengXian Light", "等线 Light", "DengXian", "等线", "Microsoft YaHei UI", "Microsoft YaHei", "PingFang SC", sans-serif';
  }
  if (primary === "dengxian" || primary === "等线") {
    return '"DengXian", "等线", "Microsoft YaHei UI", "Microsoft YaHei", "PingFang SC", sans-serif';
  }
  return value;
}

function importedPptxLineHeightCss(
  value: number | undefined,
  unit: PresentationFontSizeUnit,
): number | undefined {
  // PowerPoint's single line is 1.2 of the type size in every face it was
  
  // New Roman, Calibri, Cambria, Georgia, Segoe UI, Avenir Next, Arial Black
  // and Khmer Sangam MN, in text boxes and placeholders, at 100 % and 90 %),
  
  // Calibri's 1.22 every Calibri deck's 2 % tall. A deck's embedded Poppins
  
  const single = IMPORTED_PPTX_NORMAL_LINE_HEIGHT;
  // A text box authored here states its spacing the same way: 行距 1.0 is
  // Office's single line, and a new deck's body draws at its master's 90%.
  // Painted as 0.9em, the body's Chinese lines touched (the server's PDF and
  // PowerPoint draw them at 0.9 x 1.2em). A box that states no spacing is
  // Office's single line too, as the PDF prints it and PowerPoint opens it:
  // drawn at the page's 1.5 (the editor's 1.55), the same words were a fifth
  // taller here than in every download, and choosing 行距 1.0 -- the
  // default -- tightened them.
  if (unit !== "point") {
    return Math.round((value ?? 1) * single * 1000) / 1000;
  }
  // An omitted a:lnSpc means 100% of Office's normal line, not "inherit the
  // host application's line-height". The console shell's global leading is
  // intentionally roomier, so leaving this undefined makes every subsequent
  // imported line drift farther down even though the source geometry is exact.
  const drawingMlLineSpacing = value ?? 1;
  return (
    Math.round(drawingMlLineSpacing * single * 1000) /
    1000
  );
}

function presentationFontSizeCss(
  value: number | undefined,
  unit: PresentationFontSizeUnit,
): number | undefined {
  if (value === undefined) return undefined;
  return storedFontSizeToCssPx(value, unit);
}

/**
 * Home > Text Shadow as PowerPoint writes it (and the API's
 * `_set_run_text_shadow`): black at 43 %, 3 pt blur, 3 pt at 45 degrees
 * (`a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl"`, alpha
 * 43137), in the shape shadow model's design px.
 */
export const PRESENTATION_HOME_TEXT_SHADOW: OfficeShapeShadow = {
  color: "#000000",
  blur: 4,
  distance: 4,
  opacity: 0.43137,
  direction: 45,
  align: "tl",
};












export function presentationTextShadowCss(effect?: unknown): string {
  let stated: OfficeShapeShadow | undefined;
  try {
    stated = normalizeOfficeShapeShadow(effect);
  } catch {
    stated = undefined;
  }
  const shadow = stated ?? PRESENTATION_HOME_TEXT_SHADOW;
  const radians = ((shadow.direction ?? OFFICE_SHADOW_DEFAULT_DIRECTION) * Math.PI) / 180;
  const round = (value: number) => Math.round(value * 100) / 100 + 0;
  const value = Number.parseInt(shadow.color.slice(1), 16);
  const alpha = Math.round(Math.min(1, Math.max(0, shadow.opacity)) * 10_000) / 10_000;
  return `${round(Math.cos(radians) * shadow.distance)}px ${round(Math.sin(radians) * shadow.distance)}px ${round(shadow.blur)}px rgba(${(value >> 16) & 255}, ${(value >> 8) & 255}, ${value & 255}, ${alpha})`;
}

export function textBlockStyleCss(
  style: PresentationTextStyle | undefined,
  themed?: {
    color?: string;
    fontFamily?: string;
    fontSize?: number;
    bold?: boolean;
    italic?: boolean;
  },
  fontSizeUnit: PresentationFontSizeUnit = "design-px", fontAliases: PresentationFontAliases = EMPTY_PRESENTATION_FONT_ALIASES,
): CSSProperties {
  const styleColor =
    style && typeof style.color === "string" && style.color
      ? style.color
      : undefined;
  const styleFontFamily =
    style && typeof (style as { fontFamily?: unknown }).fontFamily === "string"
      ? ((style as { fontFamily?: string }).fontFamily as string)
      : undefined;
  const styleFontSize =
    style && typeof style.fontSize === "number" ? style.fontSize : undefined;
  const effectiveFontSize = styleFontSize ?? themed?.fontSize;
  // Earlier locally-authored fixtures and third-party adapters used the CSS
  // spelling `textAlign`. Keep accepting that finite vocabulary at the render
  // edge while `align` remains the canonical persisted field.
  const effectiveTextAlign =
    style?.align ??
    (style as { textAlign?: unknown } | undefined)?.textAlign;
  // Weight: an explicit per-block bold flag (either true OR false) is authoritative
  // and suppresses the theme/master default; only when the block has no bold flag
  // does the theme/master `bold` apply. This preserves "block override wins" even
  // for an explicit `bold: false` against a bold master.
  const hasExplicitBold = typeof style?.bold === "boolean";
  const effectiveBold = hasExplicitBold ? style?.bold : themed?.bold;
  // Optional line / paragraph spacing (PowerPoint parity) layered over the
  // visual style. `spacingToCss` returns only the keys the block actually sets
  // (lineHeight + marginTop/marginBottom). Text without a stated line spacing
  // receives Office's single line box, imported or authored.
  const spacingCss = spacingToCss(
    readSpacingFromStyle(style as Record<string, unknown> | undefined),
  );
  const renderedLineHeight = importedPptxLineHeightCss(spacingCss.lineHeight, fontSizeUnit);
  if (renderedLineHeight !== undefined) {
    spacingCss.lineHeight = renderedLineHeight;
  }
  return {
    fontSize: presentationFontSizeCss(effectiveFontSize, fontSizeUnit),
    color: styleColor ?? themed?.color ?? undefined,
    fontFamily: cssFontFamilyOrUndefined(presentationFontFamilyWithAssets(importedPptxFontFamilyCss(
      styleFontFamily ?? themed?.fontFamily ?? undefined,
      fontSizeUnit,
    ), fontAliases)),
    fontWeight: effectiveBold
      ? 700
      : presentationOfficeLightFontWeight(styleFontFamily ?? themed?.fontFamily),
    // Slant as weight: a block's own italic flag wins, else the master's.
    fontStyle: (typeof style?.italic === "boolean" ? style.italic : themed?.italic) ? "italic" : undefined,
    textDecoration: style?.underline ? "underline" : undefined,
    // Capitals over the whole box (a SmartArt layout's own list style,
    // `a:defRPr@cap`): shown so, the words themselves untouched.
    ...((style as { caps?: unknown } | undefined)?.caps === "all"
      ? { textTransform: "uppercase" as const }
      : (style as { caps?: unknown } | undefined)?.caps === "small"
        ? { fontVariant: "small-caps" }
        : {}),
    // left/center/right/justify/distribute (PowerPoint parity); `left` (the
    // default) and unrecognized values resolve to no property (byte-clean),
    // `distribute` projects to justify + text-align-last (docs-engine rule).
    // A stated Left is drawn as left: it is a choice over whatever the box
    // would otherwise take (a centred layout, the editor's own rules), not
    // the absence of one.
    ...(effectiveTextAlign === "left" ? { textAlign: "left" as const } : textAlignCss(effectiveTextAlign)),
    // PPT 段落缩进(块级):0.5 英寸/级,与导出的 marL 一致(48px@96dpi)。
    ...(typeof (style as { indentLevel?: unknown } | undefined)?.indentLevel ===
    "number"
      ? {
          paddingLeft:
            ((style as { indentLevel: number }).indentLevel ?? 0) * 48,
        }
      : {}),
    // The words' shadow: as they state it (`textShadowEffect`, PPTX readings
    // row 805), else Home > Text Shadow's own.
    textShadow:
      (style as { textShadow?: unknown } | undefined)?.textShadow === true
        ? presentationTextShadowCss((style as { textShadowEffect?: unknown } | undefined)?.textShadowEffect)
        : undefined,
    ...spacingCss,
  };
}

/** EMU per CSS px at 96 DPI (914400 EMU/in ÷ 96 px/in). Slide design baseline. */
export const AUTOFIT_EMU_PER_PX = 9525;

/** Fallback body font size (px) when a block / theme sets no explicit size. */
const AUTOFIT_BASE_FONT_PX = 13;

/** Horizontal + vertical inner padding (px) the text box reserves around text. */
const AUTOFIT_BOX_PADDING_PX = 8;


/** A CJK or full-width glyph: about one em wide, twice a Latin letter. */
const WIDE_GLYPH = /[ᄀ-ᅟ⺀-〾ぁ-㏿㐀-䶿一-鿿ꥠ-꥿가-힣豈-﫿︰-﹏＀-｠￠-￦]/u;

/**
 * The text's average advance in em, its wide glyphs counted at 1 em: a
 * Chinese line wraps nearly twice as soon as the Latin estimate says.
 */
function autofitAverageCharEm(text: string): number | undefined {
  let wide = 0;
  let narrow = 0;
  for (const char of text) {
    if (char === "\n") continue;
    if (WIDE_GLYPH.test(char)) wide += 1;
    else narrow += 1;
  }
  if (wide === 0) return undefined;
  return (wide + narrow * DEFAULT_AVG_CHAR_EM) / (wide + narrow);
}









function autofitTextBlockPadding(block: PresentationBlock): { x: number; y: number } {
  const insets = asRecord(
    textBlockStatesInsets(block) || isImportedPptxBlock(block)
      ? (block as Record<string, unknown>).textInsets
      : textBlockDrawsInsets(block)
        ? OFFICE_DEFAULT_TEXT_INSETS
        : null,
  );
  const edge = (key: "left" | "right" | "top" | "bottom") => {
    const value = insets?.[key];
    return typeof value === "number" && Number.isFinite(value) && value > 0
      ? value / AUTOFIT_EMU_PER_PX
      : 0;
  };
  return { x: edge("left") + edge("right"), y: edge("top") + edge("bottom") };
}

/**
 * The largest size a text block's text draws at, in design px: each run's own
 * size where it states one, the box's size for the rest of the text.
 */
function autofitDrawnFontSize(
  block: PresentationTextBlock,
  blockFontSize: number,
): number {
  const unit: PresentationFontSizeUnit = presentationTextUsesPoints(block)
    ? "point"
    : "design-px";
  const statedRuns: unknown[] = [];
  const paragraphs = Array.isArray(block.paragraphs) ? block.paragraphs : [];
  let unrunText = paragraphs.length === 0 && !normalizeRuns(block.runs);
  if (paragraphs.length === 0) statedRuns.push(block.runs);
  for (const paragraph of paragraphs) {
    const runs = normalizeRuns((paragraph as { runs?: unknown }).runs);
    if (runs) statedRuns.push(runs);
    else if (typeof paragraph?.text === "string" && paragraph.text) unrunText = true;
  }
  let largest = unrunText ? blockFontSize : 0;
  for (const runs of statedRuns) {
    for (const run of normalizeRuns(runs) ?? []) {
      if (!run.text) continue;
      const size = presentationFontSizeCss(
        typeof run.style?.fontSize === "number" ? run.style.fontSize : undefined,
        unit,
      );
      largest = Math.max(largest, size !== undefined && size > 0 ? size : blockFontSize);
    }
  }
  return largest > 0 ? largest : blockFontSize;
}

/**
 * The box an authored shape's words are fitted in, in design px: the text
 * rectangle its geometry states less the label's insets, exactly where
 * `shapeTextFrameLayout` draws them (CSS padding percentages are of the
 * shape's width on every side). A text box turned into a shape keeps its
 * Shrink text on overflow, and the words kept shrinking only while it was a
 * text box (PPT-B315). An imported shape is drawn at its package's autofit
 * and is not measured here.
 */
function authoredShapeAutofitBox(
  block: PresentationShapeBlock,
): { width: number; height: number } | null {
  if (isImportedPptxBlock(block)) return null;
  const width = geometryNumber(block, "width");
  const height = geometryNumber(block, "height");
  if (width === null || height === null || width <= 0 || height <= 0) return null;
  const { textRect, nativeInsets } = shapeTextFrameLayout(block);
  const rect = textRect ?? { left: 0, top: 0, right: 1, bottom: 1 };
  const inset = (edge: "left" | "right" | "top" | "bottom", dimension: number) =>
    nativeInsets
      ? typeof nativeInsets[edge] === "number"
        ? Math.max(0, Math.min(0.5 * dimension, nativeInsets[edge] as number))
        : 0
      : 0.06 * width;
  const boxWidth = (rect.right - rect.left) * width - inset("left", width) - inset("right", width);
  const boxHeight = (rect.bottom - rect.top) * height - inset("top", height) - inset("bottom", height);
  return boxWidth > 0 && boxHeight > 0
    ? { width: boxWidth / AUTOFIT_EMU_PER_PX, height: boxHeight / AUTOFIT_EMU_PER_PX }
    : null;
}

/** Build the one canonical measurement payload used by render + smart tag. */
function autofitMetricsForBlock(
  block: PresentationBlock,
  bounds: { width: number; height: number },
  resolvedFontSize: number | undefined,
): AutofitMetrics | null {
  if (!hasLayoutGeometry(block)) return null;
  const shapeBox = block.type === "shape" ? authoredShapeAutofitBox(block as PresentationShapeBlock) : null;
  if (block.type !== "text" && !shapeBox) return null;
  const widthEmu = geometryNumber(block, "width") ?? bounds.width * 0.35;
  const heightEmu = geometryNumber(block, "height") ?? bounds.height * 0.18;
  const padding = shapeBox ? null : autofitTextBlockPadding(block);
  const boxWidth = shapeBox ? shapeBox.width : widthEmu / AUTOFIT_EMU_PER_PX - padding!.x;
  const boxHeight = shapeBox ? shapeBox.height : heightEmu / AUTOFIT_EMU_PER_PX - padding!.y;
  const blockFontSize =
    typeof resolvedFontSize === "number" && resolvedFontSize > 0
      ? resolvedFontSize
      : AUTOFIT_BASE_FONT_PX;
  // The text draws at its RUNS' sizes where they state one. Measuring the
  // box's own size alone fitted 32 M at 17 px while they drew at 96 px, so
  // Shrink text on overflow changed nothing that was visible (PPT-B095). The
  // largest size drawn is the conservative estimate the fit needs.
  const fontSize = autofitDrawnFontSize(
    block as PresentationTextBlock,
    blockFontSize,
  );
  const spacing = readSpacingFromStyle(
    (block as PresentationTextBlock).style as
      | Record<string, unknown>
      | undefined,
  );
  const lineSpacing = spacing?.lineSpacing;
  const text = blockBodyTextForAutofit(block as PresentationTextBlock);
  const avgCharEm = autofitAverageCharEm(text);
  return {
    text,
    boxWidth,
    boxHeight,
    fontSize,
    // The line drawn (Office's single line times the spacing, one when
    // unstated), so Shrink text on overflow measures the lines the box paints.
    lineHeight:
      (typeof lineSpacing === "number" && lineSpacing > 0 ? lineSpacing : 1) *
      IMPORTED_PPTX_NORMAL_LINE_HEIGHT,
    ...(avgCharEm !== undefined ? { avgCharEm } : {}),
  };
}

/**
 * Compute the PowerPoint-style autofit font SCALE for a positioned text block at
 * the render edge, returning `1` (no shrink) for any block that is not in
 * `shrink` mode, has no positioned geometry, or whose text already fits — so a
 * non-autofit block renders byte-for-byte as before.
 *
 * The fit is computed in the slide's DESIGN-PIXEL space: the block's EMU box is
 * converted to px via {@link AUTOFIT_EMU_PER_PX} (the 960px-wide slide baseline
 * the renderer scales from) and compared against the text's estimated height at
 * the resolved font size (`resolvedFontSize`, falling back to the body default).
 * Because both the box and the font size live in the same px space the ratio is
 * stable regardless of how large the slide is actually projected. Pure aside
 * from reading the block; delegates the actual fit math to
 * `lib/presentation-autofit` (`autofitFontScale`).
 */
function autofitScaleForBlock(
  block: PresentationBlock,
  bounds: { width: number; height: number },
  resolvedFontSize: number | undefined,
): number {
  return autofitRenderForBlock(block, bounds, resolvedFontSize).fontScale;
}

/**
 * What to draw this block's text at: the PACKAGE's own autofit where it still
 * describes this box, else this renderer's measured shrink.
 *
 * The package wins because real PowerPoint does not recompute autofit when it
 * opens a file — a `fontScale="62500"` shape came out of PowerPoint's own PDF
 * export at exactly 62.5% of its nominal size even though the text fitted at
 * 100% (SHAPE-IMPORT-AUDIT.md §35.1). Measuring our own shrink there would draw
 * a size PowerPoint does not. Where there is no receipt (a hand-authored block,
 * or one whose text was retyped so the edit dropped it) the measured path is
 * unchanged, so nothing that never carried one moves a pixel.
 *
 * A SHAPE block gets the same treatment: since the gate admits a label whose
 * `a:bodyPr` states an autofit, its receipt rides on the shape and the label has
 * to be drawn at the package's size. The MEASURED path takes an authored
 * shape too, in the text box `authoredShapeAutofitBox` models (PPT-B315).
 */
/**
 * Per slide block list: each authored SmartArt graphic's one label scale.
 * Keyed by the list itself, so an edit (which makes a new list) re-fits.
 */
const SMARTART_LABEL_FIT_CACHE = new WeakMap<
  readonly PresentationBlock[],
  Map<string, number>
>();

/**
 * The scale an authored SmartArt label is drawn at, or `null` for any other
 * block (or when the slide's blocks are not at hand).
 *
 * Office sizes a graphic's text together (Microsoft Support, "Learn more about
 * SmartArt graphics": when one shape's text shrinks, all the others shrink to
 * the same size). A label kept the style's 32 px in any shape, so a Radial
 * Cycle's satellites showed one character of "Item 2" and clipped the rest.
 * Every label of the graphic is measured in its own box, the metrics the
 * measured autofit uses, and one scale comes back for all of them (the shared
 * `smartArtTextFitScale`, which the Docs members and the sheet shapes use).
 * Only the drawing shrinks: the label's stored size is the graphic's largest,
 * so the words grow back once the long label is gone.
 */
export function smartArtLabelFitScale(
  block: PresentationBlock,
  siblings: readonly PresentationBlock[] | undefined,
): number | null {
  if (!siblings || block.type !== "text") return null;
  const marker = readSmartArtBlockMarker(block);
  if (marker?.role !== "label") return null;
  // A size set from Home is the label's own: it draws at that size and does
  // not take part in the graphic's shared shrink (PowerPoint stops autofitting
  // a shape whose text size was set by hand).
  if (smartArtLabelSizeIsOwn(block)) return null;
  let byGroup = SMARTART_LABEL_FIT_CACHE.get(siblings);
  if (!byGroup) {
    byGroup = new Map();
    SMARTART_LABEL_FIT_CACHE.set(siblings, byGroup);
  }
  const cached = byGroup.get(marker.groupId);
  if (cached !== undefined) return cached;
  const labels = siblings.flatMap((label) => {
    const labelMarker = readSmartArtBlockMarker(label);
    if (
      label.type !== "text" ||
      labelMarker?.role !== "label" ||
      labelMarker.groupId !== marker.groupId ||
      smartArtLabelSizeIsOwn(label) ||
      // An item the layout does not draw has no box to fit.
      (label as { smartArtHidden?: unknown }).smartArtHidden === true
    ) {
      return [];
    }
    const size = (label as PresentationTextBlock).style?.fontSize;
    const width = geometryNumber(label, "width");
    const height = geometryNumber(label, "height");
    if (typeof size !== "number" || width === null || height === null) {
      return [];
    }
    // A label laid out from Office's definition states PowerPoint's 90%
    // spacing and its box is already the words' area inside the shape's
    // insets: measured at that line, with no second padding.
    const stated = ((label as PresentationTextBlock).style as { lineSpacing?: unknown } | undefined)?.lineSpacing;
    const spaced = typeof stated === "number" && stated > 0;
    return [
      {
        text: blockBodyTextForAutofit(label as PresentationTextBlock),
        width: width / AUTOFIT_EMU_PER_PX - (spaced ? 0 : AUTOFIT_BOX_PADDING_PX),
        height: height / AUTOFIT_EMU_PER_PX - (spaced ? 0 : AUTOFIT_BOX_PADDING_PX),
        fontSize: presentationTextUsesPoints(label) ? ptToPx(size) : size,
        spacing: spaced ? (stated as number) : null,
        // Each label at its own spacing: an item the author spaced 150%
        // stands taller lines than PowerPoint's 90% beside it.
        ...(spaced ? { lineHeight: (stated as number) * IMPORTED_PPTX_NORMAL_LINE_HEIGHT } : {}),
        bold: (label as PresentationTextBlock).style?.bold === true,
      },
    ];
  });
  // The catalog engine draws its labels bold; a label from Office's
  // definition says its own weight and spacing.
  const officeSpacing = labels.find((label) => label.spacing !== null)?.spacing ?? null;
  const scale = smartArtTextFitScale(
    labels,
    officeSpacing !== null
      ? { bold: labels.some((label) => label.bold), lineHeight: officeSpacing * IMPORTED_PPTX_NORMAL_LINE_HEIGHT }
      : { bold: true },
  );
  byGroup.set(marker.groupId, scale);
  return scale;
}









export function smartArtLabelSizeIsOwn(block: PresentationBlock): boolean {
  const record = block as { smartArtAuthorTextSize?: unknown; autofit?: unknown };
  return record.smartArtAuthorTextSize === true || record.autofit === "none";
}

/** A SmartArt label with the group's fit scale made its own font size. */
export function smartArtLabelAtDrawnSize(
  block: PresentationBlock,
  scale: number,
): PresentationBlock {
  if (
    block.type !== "text" ||
    readSmartArtBlockMarker(block)?.role !== "label" ||
    smartArtLabelSizeIsOwn(block)
  ) {
    return block;
  }
  const scaled = (size: unknown) =>
    typeof size === "number" && Number.isFinite(size)
      ? Math.max(1, Math.round(size * scale * 100) / 100)
      : size;
  const text = block as PresentationTextBlock;
  const style = text.style ? { ...text.style, fontSize: scaled(text.style.fontSize) as number } : text.style;
  const runs = Array.isArray(text.runs)
    ? text.runs.map((run) =>
        typeof run.style?.fontSize === "number"
          ? { ...run, style: { ...run.style, fontSize: scaled(run.style.fontSize) as number } }
          : run,
      )
    : text.runs;
  return { ...text, ...(style ? { style } : {}), ...(runs ? { runs } : {}) } as PresentationBlock;
}

/**
 * A SmartArt label about to stop being SmartArt (Convert to Shapes, a copied
 * shape) keeps the band its words were drawn in. A label draws flush in the
 * box the engine sized for it ({@link textBlockDrawsInsets}), which is the
 * shape's text rectangle with the layout's margins already taken off
 * (PowerPoint's converted Basic Block List box states them as its own
 * insets, 247650 EMU, SmartArt probe B8); as a plain text box it would take
 * PowerPoint's default margins on top, 0.2" narrower, and its words would
 * wrap and stand elsewhere than the graphic drew them. A label that states
 * its own margins keeps them.
 */
export function smartArtLabelInItsBand(block: PresentationBlock): PresentationBlock {
  if (
    block.type !== "text" ||
    readSmartArtBlockMarker(block)?.role !== "label" ||
    textBlockStatesInsets(block)
  ) {
    return block;
  }
  return { ...block, textInsets: { left: 0, right: 0, top: 0, bottom: 0 } } as PresentationBlock;
}

export function autofitRenderForBlock(
  block: PresentationBlock,
  bounds: { width: number; height: number },
  resolvedFontSize: number | undefined,
  /** The slide's blocks, so a SmartArt label can be fitted with its graphic. */
  siblings?: readonly PresentationBlock[],
  /**
   * What the host draws the block with: given, a text block is measured on
   * the slide's own drawing of its words (`autofit-measure-store`) and drawn
   * at the step that fits; until then, and without it, at the estimate.
   */
  measure?: AutofitMeasureContext,
): { fontScale: number; lineSpacingMultiplier: number } {
  const receipt = normalizeDrawnAutofitReceipt(
    (block as { source?: Record<string, unknown> }).source
      ?.nativeAutofitReceipt,
  );
  const box = {
    widthEmu: geometryNumber(block, "width"),
    heightEmu: geometryNumber(block, "height"),
  };
  const native = nativeAutofitRender(receipt, box);
  if (native) return native;
  const smartArtScale = smartArtLabelFitScale(block, siblings);
  if (smartArtScale !== null) {
    return { fontScale: smartArtScale, lineSpacingMultiplier: 1 };
  }
  if (autofitModeForBlock(block) !== "shrink") {
    return { fontScale: 1, lineSpacingMultiplier: 1 };
  }
  // The package's own words in its own box, at full size: PowerPoint draws
  // them so (it refits only while someone types, §35.1), overflow and all.
  if (
    receipt?.mode === "normal" &&
    receipt.disposition === "native" &&
    box.widthEmu === receipt.originalWidthEmu &&
    box.heightEmu === receipt.originalHeightEmu
  ) {
    return { fontScale: 1, lineSpacingMultiplier: 1 };
  }
  if (measure && block.type === "text") {
    const measured = measuredAutofit(block, measure);
    if (measured) return measured;
    requestAutofitMeasure({ ...measure, block, bounds });
  }
  // The estimate, put on the step PowerPoint would take for those words.
  const metrics = autofitMetricsForBlock(block, bounds, resolvedFontSize);
  const estimate = metrics ? autofitFontScale("shrink", metrics) : 1;
  const step = autofitStepAtOrBelow(autofitStepsFor(block), estimate);
  return {
    fontScale: step.fontScale,
    lineSpacingMultiplier: 1 - step.lineSpacingReduction,
  };
}

/** Whether a placeholder needs the contextual Office Web AutoFit smart tag. */
export function autofitSmartTagVisibleForBlock(
  block: PresentationBlock,
  bounds: { width: number; height: number },
  resolvedFontSize: number | undefined,
): boolean {
  if (!presentationTextBlockIsPlaceholder(block)) return false;
  const metrics = autofitMetricsForBlock(block, bounds, resolvedFontSize);
  return metrics ? textOverflowsAutofitFrame(metrics) : false;
}

/**
 * The plain text used to measure a text block for autofit: the joined paragraph
 * text when the block is a list (so wrapped lines + hard breaks are counted), else
 * the flat `text`. Mirrors what the renderer actually lays out.
 */
function blockBodyTextForAutofit(block: PresentationTextBlock): string {
  const paragraphs = Array.isArray(block.paragraphs) ? block.paragraphs : null;
  if (paragraphs && paragraphs.length > 0) {
    return paragraphs
      .map((p) => (typeof p?.text === "string" ? p.text : ""))
      .join("\n");
  }
  return typeof block.text === "string" ? block.text : "";
}

/**
 * Layer an autofit-shrunk `fontSize` onto an already-resolved text style. When
 * `scale` is 1 (no shrink) the style is returned UNCHANGED (same reference) so a
 * non-autofit block is byte-for-byte identical. When shrinking, the effective
 * font size (the style's own size, else the body default) is multiplied by the
 * scale; an `em`-relative cascade below it (per-run `0.7em` etc.) then shrinks
 * proportionally.
 */
/** The CSS pixels of one point where sizes are stated in points, else none. */
function autofitPxPerPoint(unit: PresentationFontSizeUnit | undefined): number | undefined {
  return unit === "point" ? CSS_PX_PER_POINT : undefined;
}












export function reducedOfficeLineHeight(lineHeight: number, multiplier: number): number {
  if (!(multiplier > 0 && multiplier < 1)) return lineHeight;
  const reduced = lineHeight - (1 - multiplier) * IMPORTED_PPTX_NORMAL_LINE_HEIGHT;
  return Math.round(Math.max(reduced, lineHeight * 0.2) * 1000) / 1000;
}

export function withAutofitFontSize(
  style: CSSProperties,
  fit: number | { fontScale: number; lineSpacingMultiplier: number },
  /** `"point"` (an imported deck): the shrunk size lands on a whole point. */
  fontSizeUnit?: PresentationFontSizeUnit,
): CSSProperties {
  const scale = typeof fit === "number" ? fit : fit.fontScale;
  const lineSpacingMultiplier =
    typeof fit === "number" ? 1 : fit.lineSpacingMultiplier;
  const shrinks = scale > 0 && scale < 1;
  // `lnSpcReduction` is subtracted from the line spacing, and the standard's own
  // note -- confirmed against real PowerPoint, where a points-spaced paragraph's
  // pitch did not move at all while a percentage-spaced one went 24.0 -> 19.7 px
  // (§35.1) -- limits it to PERCENTAGE line spacing. A CSS `lineHeight` that is
  // a plain number IS the relative one; a string is a length and is left alone.
  const reduces =
    lineSpacingMultiplier > 0 &&
    lineSpacingMultiplier < 1 &&
    typeof style.lineHeight === "number";
  if (!shrinks && !reduces) return style;
  const base =
    typeof style.fontSize === "number" ? style.fontSize : AUTOFIT_BASE_FONT_PX;
  const next: CSSProperties = { ...style };
  if (shrinks) next.fontSize = applyAutofitScale(base, scale, autofitPxPerPoint(fontSizeUnit));
  if (reduces) {
    next.lineHeight = reducedOfficeLineHeight(style.lineHeight as number, lineSpacingMultiplier);
  }
  return next;
}

/**
 * The face a run's symbol-encoded characters (U+F000-U+F0FF) are drawn from:
 * its symbol font (`RunStyle.symbolFont`, `a:sym`), resolved as a run's own
 * face is (a served open clone -- Wine's Wingdings / Symbol -- behind the
 * Office name). Undefined when the run names none.
 */
export function runSymbolFontCss(
  style: RunStyle | undefined,
  fontSizeUnit: PresentationFontSizeUnit = "design-px", fontAliases: PresentationFontAliases = EMPTY_PRESENTATION_FONT_ALIASES,
): CSSProperties | undefined {
  if (typeof style?.symbolFont !== "string" || !style.symbolFont) return undefined;
  const fontFamily = cssFontFamilyOrUndefined(presentationFontFamilyWithAssets(importedPptxFontFamilyCss(
    style.symbolFont,
    fontSizeUnit,
  ), fontAliases));
  return fontFamily ? { fontFamily } : undefined;
}

/**
 * CSS overrides for a single per-run style, layered OVER the block/paragraph
 * style on a `<span>`. Only keys the run actually overrides are set, so an
 * unstyled run inherits the parent span entirely. Underline + strike compose
 * into a single `text-decoration` (PowerPoint allows both at once).
 *
 * `autofitScale` shrinks the run's own absolute size, which is where PowerPoint
 * applies `fontScale`. A relative size (the `0.7em` a super/subscript without
 * its own size gets) already follows its parent and is left alone.
 */
export function runStyleCss(
  style: RunStyle | undefined,
  fontSizeUnit: PresentationFontSizeUnit = "design-px", fontAliases: PresentationFontAliases = EMPTY_PRESENTATION_FONT_ALIASES,
  autofitScale: number = 1,
): CSSProperties {
  if (!style) return {};
  const css: CSSProperties = {};
  if (style.bold) css.fontWeight = 700;
  if (style.italic) css.fontStyle = "italic";
  const decorations: string[] = [];
  if (style.underline) decorations.push("underline");
  if (style.strike) decorations.push("line-through");
  // A hyperlink is underlined, as PowerPoint draws one.
  const linked = typeof style.href === "string" && style.href.length > 0;
  if (linked && !style.underline) decorations.unshift("underline");
  if (decorations.length) css.textDecoration = decorations.join(" ");
  // The line each is drawn as (Font dialog > Underline Style, Double
  // Strikethrough), by the drawing table cells use.
  Object.assign(
    css,
    presentationTextDecorationCss(
      style.underline ? normalizePresentationUnderlineStyle(style.underlineStyle) : undefined,
      style.strike && style.strikeStyle === "double" ? "double" : undefined,
    ),
  );
  // `a:rPr@cap`: shown in capitals, the text itself untouched (PPT-B362).
  if (style.caps === "all") css.textTransform = "uppercase";
  else if (style.caps === "small") css.fontVariant = "small-caps";
  if (typeof style.color === "string" && style.color) css.color = style.color;
  // A hyperlink takes the theme's Hyperlink colour (`hlink`, which the slide
  // surface states as --presentation-hyperlink), unless Office's colour source
  // says the run's own text colour (`ahyp:hlinkClr val="tx"`, `linkColor`).
  
  
  if (linked && style.linkColor !== "text") {
    css.color = `var(--presentation-hyperlink, ${css.color ?? "currentcolor"})`;
  }
  // Text HIGHLIGHT (`a:highlight`): a square fill exactly behind the run's
  // glyphs, which takes no room in the line -- PowerPoint neither rounds it nor
  
  // 'Zenodo' and every word after it drawn further right). The text over it is
  // raised above every highlight of the block (`PresentationRunsText`).
  if (typeof style.highlight === "string" && style.highlight) {
    css.backgroundColor = style.highlight;
  }
  // Character text OUTLINE: its own colour (distinct from the `color` fill)
  // and width (`a:ln@w`, absent Office's 0.75 pt), drawn as PowerPoint draws
  // it -- centred and round over the glyphs' merged outline, over the fill
  // (office-text-outline-css.ts); the CSS stroke stands in until then.
  if (typeof style.textOutline === "string" && style.textOutline) {
    const widthPt = runOutlineWidthPt(style);
    // Its join, cap and dash too (row 725; round and solid where none).
    Object.assign(css, officeTextOutlineCss(style.textOutline, widthPt, `${widthPt}pt ${style.textOutline}`, runOutlineLine(style)));
  }
  if (typeof style.fontSize === "number") {
    const size = presentationFontSizeCss(style.fontSize, fontSizeUnit);
    css.fontSize =
      typeof size === "number"
        ? applyAutofitScale(size, autofitScale, autofitPxPerPoint(fontSizeUnit))
        : size;
  }
  if (typeof style.fontFamily === "string" && style.fontFamily) {
    css.fontFamily = cssFontFamilyOrUndefined(presentationFontFamilyWithAssets(importedPptxFontFamilyCss(
      style.fontFamily,
      fontSizeUnit,
    ), fontAliases));
    // A light family's regular style is its light weight (see the block rule).
    if (!style.bold) {
      const lightWeight = presentationOfficeLightFontWeight(style.fontFamily);
      if (lightWeight !== undefined) css.fontWeight = lightWeight;
    }
  }
  // Character Spacing on these characters (`a:rPr@spc`), as the block's own
  // is drawn (points to px).
  if (typeof style.letterSpacing === "number") {
    css.letterSpacing = `${pointsToCssPx(style.letterSpacing)}px`;
  }
  // A shadow on these characters only (readings row 808), drawn as a box's.
  if (style.textShadow === true && !presentationTextPaintOwnsShadow((style as { textPaint?: unknown }).textPaint)) {
    css.textShadow = presentationTextShadowCss(style.textShadowEffect);
  }
  // WordArt on these characters (their fill, outline and effects): the run's
  // own paint, over its colour and outline as PowerPoint draws `a:rPr`'s fill
  // group and `a:ln`. The same paint, at rest and while typed (one renderer).
  const paint = runTextPaintCss(style);
  if (paint) Object.assign(css, paint);
  // Baseline shift. `superscript`/`subscript` are mutually exclusive (the model
  // never carries both). `vertical-align` lifts/drops the run and the glyphs
  // draw at 0.7 of the run's size, whether that size is stated or inherited:
  // the stated size stays the character's own (what Clear Formatting and the
  // Font Size box read), as PowerPoint keeps `sz` beside `baseline`. An
  // explicit 64 px superscript drew at 64 px (PPT-B101).
  if (style.superscript || style.subscript) {
    css.verticalAlign = style.superscript ? "super" : "sub";
    css.fontSize =
      typeof css.fontSize === "number"
        ? Math.round(css.fontSize * PRESENTATION_SCRIPT_SCALE * 100) / 100
        : `${PRESENTATION_SCRIPT_SCALE}em`;
    // A raised or lowered run does not make its line taller: PowerPoint
    
    // lines without them; the raised box pushed ours to 30.3.
    css.lineHeight = 0;
  }
  return css;
}









export function importedPptxRunLineBoxCss(
  runs: readonly TextRun[] | undefined,
  fontSizeUnit: PresentationFontSizeUnit, fontAliases: PresentationFontAliases = EMPTY_PRESENTATION_FONT_ALIASES,
  autofitScale: number = 1,
): CSSProperties {
  if (fontSizeUnit !== "point") return {};
  const normalized = normalizeRuns(runs);
  const firstRun = normalized?.find((run) => run.text.length > 0);
  const runCss = runStyleCss(firstRun?.style, fontSizeUnit, fontAliases, autofitScale);
  const sizes = (normalized ?? []).filter((run) => run.text.length > 0).map((run) => run.style?.fontSize);
  const smallest =
    sizes.length > 1 && sizes.every((size) => typeof size === "number" && size > 0)
      ? Math.min(...(sizes as number[]))
      : undefined;
  const strut =
    smallest !== undefined && smallest !== firstRun?.style?.fontSize
      ? runStyleCss({ fontSize: smallest }, fontSizeUnit, fontAliases, autofitScale).fontSize
      : runCss.fontSize;
  return {
    fontSize: strut,
    fontFamily: runCss.fontFamily,
  };
}








export function needsStructuredParagraphRender(
  paragraphs: readonly SlideParagraphWithRuns[],
): boolean {
  return (
    ownsParagraphRichLayout(paragraphs) ||
    paragraphs.some((paragraph) => paragraphStatesIndent(paragraph))
  );
}

/**
 * A flat body's style with its rows' shared indent level. Rows that all sit
 * at one level and state nothing else of their own are drawn by the flat
 * body, which indents by `style.indentLevel`: Increase Indent on a plain box
 * stored the level on its rows and the body drew them flush until a bullet
 * made them a list (PPT-B286).
 */
export function textStyleWithUniformIndent(
  block: PresentationTextBlock,
): PresentationTextBlock["style"] {
  const style = block.style;
  if ((style as { indentLevel?: unknown } | undefined)?.indentLevel !== undefined) return style;
  const paragraphs = normalizeParagraphsWithRuns(block.paragraphs);
  if (!paragraphs?.length || needsStructuredParagraphRender(paragraphs)) return style;
  const level = paragraphs[0].level;
  if (!(level > 0) || paragraphs.some((paragraph) => paragraph.level !== level)) return style;
  return { ...(style ?? {}), indentLevel: level } as PresentationTextBlock["style"];
}

/** EM of hanging indent applied per outline level for a bulleted/numbered list. */
export const PARAGRAPH_INDENT_EM = 1.5;

/** The size a slide's raised or lowered run is drawn at, of its own. */
export const PRESENTATION_SCRIPT_SCALE = 0.7;

/** Width reserved for the bullet/number marker in the hanging-indent gutter. */
export const PARAGRAPH_MARKER_EM = 1.4;

/**
 * A row's OWN `a:pPr@marL`/`@indent`, as CSS.
 *
 * ISO/IEC 29500-1 §21.1.2.2.7: a paragraph property stated closer to the text
 * takes precedence, so a direct `@marL` REPLACES the outline level's inherited
 * margin instead of adding to it — the caller therefore spreads this AFTER the
 * `level * PARAGRAPH_INDENT_EM` default, and a row that states nothing keeps
 * that default untouched. `@indent` offsets the FIRST line (negative = the
 * hanging indent PowerPoint's Paragraph dialog calls "Hanging"), which a
 * bulleted row already expresses through its marker gutter — the same reading
 * the SmartArt display metrics have always used, kept in one place so the two
 * cannot drift.
 */
export function paragraphOwnIndentCss(
  paragraph: SlideParagraph,
  hasMarker: boolean,
): CSSProperties {
  const style: CSSProperties = {};
  if (paragraph.marL !== undefined) {
    style.paddingInlineStart = paragraph.marL / AUTOFIT_EMU_PER_PX;
  }
  if (paragraph.indent !== undefined && !hasMarker) {
    // PowerPoint never starts a first line left of the box: a hanging
    // `@indent` past the row's margin (marL + indent < 0) draws the first line
    
    // `@marL`, every line at the edge in PowerPoint, the continuation lines
    // pushed in here; 51 such rows in 13 corpus decks).
    const floor = paragraph.marL ? -paragraph.marL : 0;
    style.textIndent = Math.max(paragraph.indent, floor) / AUTOFIT_EMU_PER_PX;
  }
  return style;
}

/** Put line/before/after spacing on every paragraph row, not the wrapper. */
export function splitParagraphListTextStyle(style: CSSProperties): {
  containerStyle: CSSProperties;
  paragraphStyle: CSSProperties;
} {
  const { lineHeight, marginTop, marginBottom, ...containerStyle } = style;
  const paragraphStyle: CSSProperties = {};
  if (lineHeight !== undefined) paragraphStyle.lineHeight = lineHeight;
  // Padding cannot collapse between adjacent block rows, so the authored
  // before + after values remain additive just like DrawingML paragraph space.
  if (marginTop !== undefined) paragraphStyle.paddingTop = marginTop;
  if (marginBottom !== undefined) paragraphStyle.paddingBottom = marginBottom;
  return { containerStyle, paragraphStyle };
}





















export function officePercentLeadingCss(lineHeight: unknown, fontFamily?: unknown): CSSProperties {
  const shift = officePercentLeading(lineHeight, fontFamily).shift;
  return shift === 0 ? {} : { position: "relative", top: `${shift}em` };
}

/** The last row's foot under {@link officePercentLeadingCss}: its flow ends at PowerPoint's last descent. */
export function officePercentLeadingTailCss(lineHeight: unknown, fontFamily?: unknown): CSSProperties {
  const tail = officePercentLeading(lineHeight, fontFamily).tail;
  return tail === 0 ? {} : { marginBottom: `${-tail}em` };
}











export function officePercentLeadingFlatRowCss(lineHeight: unknown, fontFamily?: unknown): CSSProperties {
  return { ...officePercentLeadingCss(lineHeight, fontFamily), ...officePercentLeadingTailCss(lineHeight, fontFamily) };
}

/**
 * The row's face as PowerPoint and the browser measure it, per em: its
 * descent, and half its ascent less its descent (where CSS puts the baseline
 * past the middle of the line box). Read from the Office files recorded for
 * Word (`wordFontMetrics`; a metric-compatible stand-in draws with the same
 * numbers), Calibri's for any other face; null for an East Asian face, whose
 * lines PowerPoint lays out by other rules.
 */
function officeRowFaceVerticals(fontFamily: unknown): { descent: number; cssHalf: number } | null {
  const fallback = { descent: 0.264, cssHalf: 0.35 };
  if (typeof fontFamily !== "string" || fontFamily.trim() === "") return fallback;
  const primary = fontFamily.split(",")[0].trim().replace(/^["']|["']$/g, "");
  if (wordFontIsEastAsian(primary)) return null;
  const line = wordFontLineMetrics(primary);
  const descent = wordFontLineDescent(primary, false);
  if (!line || descent === null || !(descent > 0) || !(line.box > descent)) return fallback;
  return { descent, cssHalf: (line.box - 2 * descent) / 2 };
}

/**
 * {@link officePercentLeadingCss}'s numbers, in em of the row's type size: how
 * far the glyphs move down from where CSS lays them (`shift`) and how much of
 * the last line's box PowerPoint leaves out under its descent (`tail`). A
 * slide table's cell lays its lines out by the same rule (look-fit
 * truth-probes/lookfit-cell-spacing-probe.pptx). Pure.
 */
export function officePercentLeadingEm(lineHeight: unknown, fontFamily?: unknown): { shift: number; tail: number } {
  return officePercentLeading(lineHeight, fontFamily);
}

function officePercentLeading(lineHeight: unknown, fontFamily?: unknown): { shift: number; tail: number } {
  if (typeof lineHeight !== "number" || !Number.isFinite(lineHeight) || lineHeight <= 0) {
    return { shift: 0, tail: 0 };
  }
  const round = (value: number) => Math.round(value * 10000) / 10000;
  // An East Asian face's line, at every spacing: PowerPoint leaves its own
  // descent share under the line (`officeLineFace`, at least 0.2 of a single
  // line whatever the face, and whatever Latin face the run states), its
  // first baseline at single less that share, below single at the line less
  // it (no three-quarter floor: MS PGothic at 90 % stood 0.83 em down, Yu
  // Gothic 0.79), above single at the larger of single's and 0.7547 L as a
  // Latin line's. CSS stood a 90 % MS PGothic title 0.075 em under
  
  // truth-probes/lookfit-cjk-spacing-probe.pptx and lookfit-cjk-faces-probe.pptx).
  const eastAsian = officeLineFace(fontFamily, false);
  if (eastAsian?.eastAsian) {
    const cssBaseline = lineHeight / 2 + eastAsian.cssHalf;
    const officeBaseline =
      lineHeight > IMPORTED_PPTX_NORMAL_LINE_HEIGHT + 0.001
        ? Math.max(IMPORTED_PPTX_NORMAL_LINE_HEIGHT - eastAsian.descent, 0.7547 * lineHeight)
        : lineHeight - eastAsian.descent;
    return {
      shift: round(officeBaseline - cssBaseline),
      tail: round(Math.max(0, lineHeight - officeBaseline - eastAsian.descent)),
    };
  }
  if (Math.abs(lineHeight - IMPORTED_PPTX_NORMAL_LINE_HEIGHT) <= 0.001) return { shift: 0, tail: 0 };
  if (lineHeight < IMPORTED_PPTX_NORMAL_LINE_HEIGHT) {
    // Below single (Office's own 90 % body and titles, an autofit's
    // `lnSpcReduction`) PowerPoint keeps the face's descent at the foot of
    // the line and takes what the line lost from above it -- the first
    // baseline L less the descent -- down to 0.7547 L, under which the
    // baseline goes on falling with the line whatever the face; and the last
    // descent stays under the last baseline. CSS shares the loss half above,
    // half below. PowerPoint's PDF, 30 pt at 95-50 % (MRAI-corpus/
    
    // Aptos 24.0, Arial 25.92, Times 24.96 pt, at 70 % all four 18.96 and at
    // 50 % 13.92 pt; under a bottom anchor the last baseline 7.92 pt over the
    // foot at every spacing. CSS had a 90 % Calibri body 2 pt low, a 70 %
    // one 3.5 pt.
    const face = officeRowFaceVerticals(fontFamily);
    if (!face) return { shift: 0, tail: 0 };
    const baseline = Math.max(lineHeight - face.descent, 0.7547 * lineHeight);
    const shift = baseline - (lineHeight / 2 + face.cssHalf);
    const tail = lineHeight - baseline - face.descent;
    return { shift: round(shift), tail: round(tail) };
  }
  // Above single the first baseline is three quarters of the line whatever the
  // face (PowerPoint's print of look-fit truth-probes/lookfit-box-spacing-
  // probe.pptx, 20 pt at 115-250 %: 20.9 / 26.9 / 36.0 / 44.9 pt under the top
  // in Calibri, Arial, Times New Roman and Georgia alike), and under the last
  // baseline PowerPoint leaves the face's share of a single line taken by its
  // descent (Calibri 0.262, Arial 0.228, Times New Roman 0.240, Georgia 0.228
  // em): `officeLineFace`, read from the font files. Calibri's 0.264 for every
  // face set Arial's last line 0.7 pt low at 20 pt, and its table rows that
  // much too tall.
  const face = officeLineFace(fontFamily);
  const cssBaseline = lineHeight / 2 + face.cssHalf;
  const officeBaseline = Math.max(IMPORTED_PPTX_NORMAL_LINE_HEIGHT - face.descent, 0.7547 * lineHeight);
  const shift = Math.max(0, officeBaseline - cssBaseline);
  const tail = Math.max(0, lineHeight - cssBaseline - shift - face.descent);
  return { shift: round(shift), tail: round(tail) };
}

/**
 * A face's numbers for a percentage spacing above single
 * (`office-line-metrics.generated.json`, scripts/measure-office-line-metrics.py,
 * read from the font files): the share of a single 1.2 em line its descent
 * takes in PowerPoint (1.2 x usWinDescent / (usWinAscent + usWinDescent) of the
 * Office file; an East Asian face's at least 0.24), half of ascent less
 * descent of the face the browser lays the line out with (where CSS puts the
 * baseline past the line's middle), and whether the face is East Asian (its
 * Office file's code pages). Calibri's for a face the table does not hold,
 * or none when `calibriWhenUnknown` is false.
 */
function officeLineFace(fontFamily: unknown): { descent: number; cssHalf: number; eastAsian: boolean };
function officeLineFace(
  fontFamily: unknown,
  calibriWhenUnknown: false,
): { descent: number; cssHalf: number; eastAsian: boolean } | undefined;
function officeLineFace(
  fontFamily: unknown,
  calibriWhenUnknown = true,
): { descent: number; cssHalf: number; eastAsian: boolean } | undefined {
  const primary = typeof fontFamily === "string" ? fontFamily.split(",")[0].trim().replace(/^["']|["']$/g, "").toLowerCase() : "";
  const stated = OFFICE_LINE_FACES.get(primary) ?? (calibriWhenUnknown ? OFFICE_LINE_FACES.get("calibri") : undefined);
  // An East Asian face PowerPoint for Windows sets apart from the rule keeps
  // what its measured first baseline leaves (`eastAsianMeasuredFirstBaselineEm`).
  const measured = stated?.[2] === 1 ? eastAsianMeasuredFirstBaselineEm(primary) : undefined;
  if (stated) return { descent: measured !== undefined ? Math.round((IMPORTED_PPTX_NORMAL_LINE_HEIGHT - measured) * 10000) / 10000 : stated[0]!, cssHalf: stated[1]!, eastAsian: stated[2] === 1 };
  return calibriWhenUnknown ? { descent: 0.264, cssHalf: 0.35, eastAsian: false } : undefined;
}

export function paragraphRowTextStyle(
  paragraph: SlideParagraphWithRuns,
  fallback: CSSProperties | undefined,
  fontSizeUnit: PresentationFontSizeUnit,
  






  lineSpacingMultiplier = 1,
): CSSProperties {
  const style: CSSProperties = { ...(fallback ?? {}) };
  for (const field of normalizedSpacingFallbackBlocked(
    paragraph.spacingFallbackBlocked,
  ) ?? []) {
    if (field === "lineSpacing") {
      // A row that resolved no line spacing of its own (the box's stands for
      // its first row only) draws at Office's single line, as an omitted
      // `a:lnSpc` does (`importedPptxLineHeightCss`) -- less the package's
      // autofit reduction, as a stated 100 % takes it. With none it took the
      // host's leading: the editor's 1.5, rows 30 pt where PowerPoint's are
      
      const single = importedPptxLineHeightCss(undefined, fontSizeUnit);
      style.lineHeight =
        typeof single === "number" && lineSpacingMultiplier > 0 && lineSpacingMultiplier < 1
          ? reducedOfficeLineHeight(single, lineSpacingMultiplier)
          : single;
    }
    if (field === "spaceBefore") delete style.paddingTop;
    if (field === "spaceAfter") delete style.paddingBottom;
  }
  const ownSpacing = normalizedParagraphSpacingFields(paragraph);
  const own = spacingToCss(ownSpacing);
  if (own.lineHeight !== undefined) {
    const rowLineHeight = importedPptxLineHeightCss(own.lineHeight, fontSizeUnit);
    style.lineHeight =
      typeof rowLineHeight === "number" && lineSpacingMultiplier > 0 && lineSpacingMultiplier < 1
        ? reducedOfficeLineHeight(rowLineHeight, lineSpacingMultiplier)
        : rowLineHeight;
  } else if (ownSpacing.lineSpacingPoints !== undefined) {
    // Line Spacing > Exactly: the row's lines are that many points tall, the
    // glyphs where PowerPoint puts them in each (`exactLineSpacingCss`), at
    // the row's largest type size (its runs', else the box's).
    const runSizes = (normalizeRuns(paragraph.runs) ?? [])
      .map((run) => run.style?.fontSize)
      .filter((value): value is number => typeof value === "number");
    const fallbackSize =
      typeof fallback?.fontSize === "number"
        ? fontSizeUnit === "point"
          ? designPxToPoints(fallback.fontSize)
          : fallback.fontSize
        : undefined;
    const typeSize =
      runSizes.length > 0 ? Math.max(...runSizes) : fallbackSize;
    Object.assign(
      style,
      exactLineSpacingCss(ownSpacing.lineSpacingPoints, typeSize),
    );
  }
  if (ownSpacing.lineSpacingPoints === undefined) {
    Object.assign(style, officePercentLeadingCss(style.lineHeight, style.fontFamily));
  }
  if (own.marginTop !== undefined) style.paddingTop = own.marginTop;
  if (own.marginBottom !== undefined) style.paddingBottom = own.marginBottom;
  const align = normalizeParagraphAlign(paragraph.align);
  if (align === "distribute") {
    style.textAlign = "justify";
    style.textAlignLast = "justify";
  } else if (align !== undefined) {
    style.textAlign = align;
    delete style.textAlignLast;
  }
  return style;
}

/**
 * PowerPoint keeps an oversized first-paragraph `spcBef` from displacing the
 * useful text when that leading space alone would overflow a fixed no-autofit
 * text frame. CSS padding has no equivalent overflow negotiation, so reproduce
 * that bounded case at the render edge while preserving the authored value for
 * editing and export. Ordinary paragraph spacing, multi-column text and text
 * that overflows even without the leading space remain unchanged.
 */
export function suppressImportedFirstParagraphSpaceBefore(
  block: PresentationBlock,
  bounds: { width: number; height: number },
  paragraphs: readonly SlideParagraphWithRuns[],
  paragraphStyle: CSSProperties | undefined,
  textStyle: CSSProperties,
): boolean {
  if (
    !isImportedPptxBlock(block) ||
    // A box that shrinks its text (Resize shape to fit text only moves the
    // frame, which is what PowerPoint drew it at).
    autofitModeForBlock(block) === "shrink" ||
    normalizeColumnCount((block as PresentationTextBlock).columnCount) > 1 ||
    paragraphs.length === 0
  ) {
    return false;
  }
  const widthEmu = geometryNumber(block, "width") ?? bounds.width;
  const heightEmu = geometryNumber(block, "height") ?? bounds.height;
  const insets = asRecord((block as Record<string, unknown>).textInsets);
  const insetEmu = (key: "left" | "right" | "top" | "bottom") => {
    const value = insets?.[key];
    return typeof value === "number" && Number.isFinite(value)
      ? Math.max(0, value)
      : 0;
  };
  const boxWidth =
    (widthEmu - insetEmu("left") - insetEmu("right")) /
    AUTOFIT_EMU_PER_PX;
  const boxHeight =
    (heightEmu - insetEmu("top") - insetEmu("bottom")) /
    AUTOFIT_EMU_PER_PX;
  if (!(boxWidth > 0) || !(boxHeight > 0)) return false;

  let firstSpaceBefore = 0;
  let estimatedHeight = 0;
  for (const [index, paragraph] of paragraphs.entries()) {
    const rowStyle = paragraphRowTextStyle(paragraph, paragraphStyle, "point");
    const spaceBefore =
      typeof rowStyle.paddingTop === "number" ? rowStyle.paddingTop : 0;
    const spaceAfter =
      typeof rowStyle.paddingBottom === "number" ? rowStyle.paddingBottom : 0;
    if (index === 0) firstSpaceBefore = spaceBefore;
    const runFontSizes = (normalizeRuns(paragraph.runs) ?? [])
      .map((run) => run.style?.fontSize)
      .filter((value): value is number => typeof value === "number")
      .map(pointsToPx);
    const blockFontSize =
      typeof textStyle.fontSize === "number"
        ? textStyle.fontSize
        : AUTOFIT_BASE_FONT_PX;
    const fontSize = Math.max(blockFontSize, ...runFontSizes);
    // An exact line height (`lineSpacingPoints`) arrives in px: as a multiple
    // of this row's size for the estimate.
    const exactLineHeightPx =
      typeof rowStyle.lineHeight === "string" && rowStyle.lineHeight.endsWith("px")
        ? Number.parseFloat(rowStyle.lineHeight)
        : Number.NaN;
    const lineHeight =
      typeof rowStyle.lineHeight === "number"
        ? rowStyle.lineHeight
        : Number.isFinite(exactLineHeightPx) && fontSize > 0
          ? exactLineHeightPx / fontSize
          : 1.2;
    const text = paragraph.text || " ";
    const cjkAdvance = /[\u3400-\u9fff\uf900-\ufaff]/u.test(text) ? 1 : undefined;
    estimatedHeight +=
      estimateTextHeight(text, boxWidth, fontSize, lineHeight, cjkAdvance) +
      spaceBefore +
      spaceAfter;
  }
  return (
    firstSpaceBefore > 0 &&
    estimatedHeight > boxHeight &&
    estimatedHeight - firstSpaceBefore <= boxHeight
  );
}

export { CSS_PX_PER_POINT } from "@mrnote/core/web/lib/presentation/text-units.ts";

/**
 * Centered text label overlaid on a SHAPE block (PowerPoint / WPS "add text to
 * shape"). Reads the shape's `text` (+ optional per-run `runs`) and its optional
 * `style` (color / size / bold / italic / align…). Absolutely positioned to fill
 * the shape box and center its content both axes so the text sits in the middle
 * of the rectangle / ellipse / diamond / callout. Returns `null` when the shape
 * carries no text (byte-stable for a plain shape). Shared by the present overlay
 * and the editor's non-editing render so the two never drift. A ZOOM tile uses
 * its own `ZoomTileLabel`, so this skips zoom tiles to avoid a double label.
 */
/**
 * Where a shape's words sit: the text rectangle its geometry states (DrawingML
 * `a:rect`), inset by the label's own insets (else 6%), with the vertical
 * anchor as the flex alignment of the one child. The label draws in it, and
 * the in-shape editor types in the SAME box at the same anchor, so entering
 * and leaving the words never moves them (PPT-B190).
 */
const SHAPE_TEXT_DEFAULT_INSET = 0.06;

export function shapeTextFrameLayout(block: PresentationShapeBlock): {
  vAlign: ReturnType<typeof shapeLabelVerticalAlign>;
  alignItems: "flex-start" | "center" | "flex-end";
  nativeInsets: Record<string, unknown> | null | undefined;
  shapeHeight: number;
  textRect: ReturnType<typeof shapeTextRectFractions>;
  frame: CSSProperties;
} {
  const style = block.style as PresentationTextStyle | undefined;
  // A shape label DEFAULTS to vertically centered (the editor's own choice for
  // a hand-drawn shape); an EXPLICIT `style.verticalAlign` overrides that, and
  // a native placeholder's absent anchor is its layout's `t`
  // (`shapeLabelVerticalAlign`). The flex `alignItems` controls the vertical
  // position of the single text child here.
  const vAlign = shapeLabelVerticalAlign(style, block.source);
  const alignItems =
    vAlign === "top"
      ? "flex-start"
      : vAlign === "bottom"
        ? "flex-end"
        : "center";
  const nativeInsets = asRecord((block as Record<string, unknown>).textInsets);
  const shapeWidth = geometryNumber(block, "width") ?? 0;
  const shapeHeight = geometryNumber(block, "height") ?? 0;
  const insetPercent = (edge: string, dimension: number) =>
    typeof nativeInsets?.[edge] === "number" && dimension > 0
      ? `${Math.max(0, Math.min(50, (nativeInsets[edge] as number) / dimension * 100))}%`
      : "6%";
  // The geometry's own text rectangle (DrawingML `a:rect`): PowerPoint lays the
  // label out inside it, the insets measured from its edges. Absent, it is the
  // frame and nothing below changes.
  const textRect = shapeTextRectFractions(
    block.source,
    normalizeShapeKind(block.shape),
  ) ?? (
    // No imported drawing (or handles dragged here): the preset's own text
    // rectangle from the shared catalog evaluator, as Docs and the sheet use.
    !asRecord(asRecord(block.source)?.nativeShapeDrawing) || normalizeDocumentShapeAdjustments(block.adjustments)
      ? authoredShapeTextRectFractions(
          presentationShapeCatalogEntry(normalizeShapeKind(block.shape))?.geometry === "preset"
            ? (presentationShapeCatalogEntry(normalizeShapeKind(block.shape)) as { drawingMlPreset: string }).drawingMlPreset
            : null,
          shapeWidth,
          shapeHeight,
          normalizeDocumentShapeAdjustments(block.adjustments),
        )
      : null
  );
  // A shape without stated insets keeps 6% of the text rectangle clear on
  // each side: of its HEIGHT above and below, of its WIDTH left and right.
  // As CSS padding the percentage read the frame's width on all four sides,
  // so widening a shape grew the space above and below its words and pushed
  // a centred label onto the bottom edge (PPT-B341).
  const rectWidth = textRect ? Math.max(0, textRect.right - textRect.left) : 1;
  const rectHeight = textRect ? Math.max(0, textRect.bottom - textRect.top) : 1;
  const rectEdge = (fraction: number, edge: string, dimension: number) => {
    const inset = nativeInsets
      ? typeof nativeInsets[edge] === "number" && dimension > 0
        ? Math.max(0, Math.min(50, (nativeInsets[edge] as number) / dimension * 100))
        : 0
      : SHAPE_TEXT_DEFAULT_INSET * 100 * (edge === "top" || edge === "bottom" ? rectHeight : rectWidth);
    return `${fraction * 100 + inset}%`;
  };
  return {
    vAlign,
    alignItems,
    nativeInsets,
    shapeHeight,
    textRect,
    frame: {
      position: "absolute",
      inset: 0,
      padding: 0,
      ...(textRect || !nativeInsets ? {
        top: rectEdge(textRect?.top ?? 0, "top", shapeHeight),
        right: rectEdge(1 - (textRect?.right ?? 1), "right", shapeWidth),
        bottom: rectEdge(1 - (textRect?.bottom ?? 1), "bottom", shapeHeight),
        left: rectEdge(textRect?.left ?? 0, "left", shapeWidth),
      } : {
        top: insetPercent("top", shapeHeight), right: insetPercent("right", shapeWidth),
        bottom: insetPercent("bottom", shapeHeight), left: insetPercent("left", shapeWidth),
      }),
    },
  };
}

/**
 * The width a shape's label is laid out in, in EMU: the frame
 * {@link shapeTextFrameLayout} places it in (the geometry's text rectangle
 * less the insets, or the default 6 % of it). Null without a width.
 */
export function shapeTextAreaWidthEmu(block: PresentationShapeBlock): number | null {
  const width = geometryNumber(block, "width");
  if (width === null || width <= 0) return null;
  const { frame } = shapeTextFrameLayout(block);
  const fraction = (value: unknown) => {
    const match = typeof value === "string" ? /^(-?[0-9.]+)%$/.exec(value) : null;
    return match ? Number(match[1]) / 100 : 0;
  };
  return Math.max(0, width * (1 - fraction(frame.left) - fraction(frame.right)));
}

/** PowerPoint's size for a shape's text that states none: 18 pt (24 slide px). */
export const SHAPE_LABEL_DEFAULT_FONT_PX = 24;

const SHAPE_LABEL_UNIFORM_RUN_KEYS = [
  "fontSize",
  "fontFamily",
  "color",
  "bold",
  "italic",
  "underline",
  "strike",
] as const;

/**
 * The run style every non-empty run of a shape's label agrees on (a key is
 * left out when runs differ or one leaves it unset). An imported label
 * states its size, face and colour on its runs, not on the shape's style.
 */
export function shapeLabelUniformRunStyle(block: PresentationBlock): Partial<RunStyle> {
  const record = block as { runs?: unknown; paragraphs?: unknown };
  const runs: TextRun[] = [...(normalizeRuns(record.runs) ?? [])];
  if (Array.isArray(record.paragraphs)) {
    for (const paragraph of record.paragraphs) {
      runs.push(...(normalizeRuns(asRecord(paragraph)?.runs) ?? []));
    }
  }
  const written = runs.filter((run) => run.text.length > 0);
  const uniform: Record<string, unknown> = {};
  if (written.length === 0) return uniform;
  for (const key of SHAPE_LABEL_UNIFORM_RUN_KEYS) {
    const first = (written[0]!.style as Record<string, unknown> | undefined)?.[key];
    if (first === undefined) continue;
    if (written.every((run) => (run.style as Record<string, unknown> | undefined)?.[key] === first)) {
      uniform[key] = first;
    }
  }
  return uniform as Partial<RunStyle>;
}

/** The alignment a shape's label is drawn with: its own, the one all its paragraphs state, else centred. */
export function shapeLabelTextAlignCss(block: PresentationBlock): CSSProperties {
  const style = block.style as { align?: unknown; textAlign?: unknown } | undefined;
  let align: unknown = style?.align ?? style?.textAlign;
  if (align === undefined) {
    const paragraphs = (block as { paragraphs?: unknown }).paragraphs;
    if (Array.isArray(paragraphs)) {
      const stated = paragraphs
        .map((paragraph) => asRecord(paragraph))
        .filter((paragraph) => typeof paragraph?.text === "string" && (paragraph.text as string).length > 0)
        .map((paragraph) => paragraph?.align);
      if (stated.length > 0 && stated.every((value) => value === stated[0])) align = stated[0];
    }
  }
  // `left` is the CSS default, so textAlignCss gives nothing for it; a shape
  // label's default is centre, so left must be stated.
  if (align === "left") return { textAlign: "left" };
  if (isTextAlign(align)) return textAlignCss(align);
  return { textAlign: "center" };
}













export function shapeLabelNeedsRichEditor(block: PresentationBlock): boolean {
  const record = block as {
    runs?: unknown;
    paragraphs?: unknown;
    textPaint?: unknown;
    textWarp?: unknown;
    columnCount?: unknown;
  };
  if (shapeTextPaintCss(block) !== null) return true;
  if (normalizeOfficeTextWarp(record.textWarp) !== null) return true;
  if (normalizeColumnCount(record.columnCount) > 1) return true;
  const paragraphs = normalizeParagraphsWithRuns(record.paragraphs);
  if (
    paragraphs &&
    (paragraphs.length > 1 ||
      needsStructuredParagraphRender(paragraphs) ||
      paragraphs[0].bullet !== "none" ||
      paragraphs[0].level > 0)
  ) {
    return true;
  }
  const runs = [
    ...(normalizeRuns(record.runs) ?? []),
    ...(paragraphs ?? []).flatMap((paragraph) => normalizeRuns(paragraph.runs) ?? []),
  ].filter((run) => run.text.length > 0);
  // An equation among the words is drawn (and walked through) as an
  // equation, which a plain field cannot do.
  if (runs.some((run) => Boolean((run as { math?: unknown }).math))) return true;
  if (runs.length < 2) return false;
  const first = JSON.stringify(runs[0].style ?? {});
  return runs.some((run) => JSON.stringify(run.style ?? {}) !== first);
}

/**
 * A shape's words while they are typed, drawn as they read at rest
 * (ShapeTextOverlay): the label's alignment, its runs' size, face and colour
 * where they agree, the package's autofit, PowerPoint's default size for a
 * label that states none, and its wrap. The editor is a plain textarea, so a
 * label whose runs genuinely differ shows the shape's own style there.
 */
export function shapeLabelEditorTextCss(
  block: PresentationBlock,
  fontSizeUnit: PresentationFontSizeUnit,
  fontAliases: PresentationFontAliases = EMPTY_PRESENTATION_FONT_ALIASES,
): CSSProperties {
  const style = {
    ...((block.style as Record<string, unknown> | undefined) ?? {}),
    ...shapeLabelUniformRunStyle(block),
  } as PresentationTextStyle;
  const base = textBlockStyleCss(style, undefined, fontSizeUnit, fontAliases);
  const sized =
    typeof base.fontSize === "number"
      ? withAutofitFontSize(base, autofitRenderForBlock(block, { width: 0, height: 0 }, base.fontSize), fontSizeUnit)
      : { ...base, fontSize: SHAPE_LABEL_DEFAULT_FONT_PX };
  const wraps = (block.style as { textWrap?: unknown } | undefined)?.textWrap !== false;
  return {
    ...sized,
    ...shapeLabelTextAlignCss(block),
    whiteSpace: wraps ? "pre-wrap" : "pre",
    overflowWrap: wraps ? "anywhere" : "normal",
  };
}
