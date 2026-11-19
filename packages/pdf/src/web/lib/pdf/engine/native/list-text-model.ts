/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeResourceDictionaryKind } from "./content-stream-transaction";
import type { PdfNativeParagraphRewriteEvidence } from "./paragraph-rewrite";
import { readMraiNativeParagraphRewriteEvidence } from "./paragraph-rewrite-evidence";
import { isMraiNativePredefinedUnicodeCMapName } from "./predefined-cmap-names";
import { PDF_EXISTING_RESOURCE_NAME } from "./embedded-image-model";
import { PdfNativeTransactionError } from "./transaction-error";
import { MRAI_NATIVE_ADAPTIVE_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION, MRAI_NATIVE_LIST_TEXT_MISSING_GLYPH_MESSAGE, MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION, MRAI_NATIVE_PROJECTED_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION, MRAI_NATIVE_TEXT_PROPERTIES_SCHEMA_VERSION, PDF_NATIVE_MAX_FIXED_FRAME_ITEMS, PDF_NATIVE_MAX_FIXED_FRAME_ITEM_INDEX_GAP, PDF_NATIVE_MAX_ITEM_TEXT_LENGTH } from "./vocabulary";
import type { PdfMatrix, PdfNativeDeviceColor, PdfNativeFixedFrameLinePlan, PdfNativeFontEncoding, PdfNativeFontEncodingKind, PdfNativeIndirectObjectRef, PdfNativeListTextFidelityEvidence, PdfNativeListTextSelectionEvidence, PdfNativeListTextSelectionItemEvidence, PdfNativeListTextSourceRun, PdfNativeParagraphReflowLineEvidence, PdfNativeParagraphReflowMutationEvidence, PdfNativeStandard14EncodingKind, PdfNativeStandard14LatinBaseFont, PdfNativeTextFontOption, PdfNativeTextLayoutWitnessEvidence, PdfNativeTextPropertyMutationEvidence, PdfNativeTextPropertyState } from "./list-text-fidelity";

/** A resource name an evidence record carries: the page's own (a font, a
 * colour space), read and written back as it is spelled, or one this engine
 * made (the narrow form, which this includes). See PDF_EXISTING_RESOURCE_NAME. */
export const RESOURCE_NAME = PDF_EXISTING_RESOURCE_NAME;

export const MAX_PROVEN_LINE_SPACING = 100_000;

/** Superscript and subscript ride on `Ts`; a rise beyond this is preserve-only. */
export const MAX_PROVEN_TEXT_RISE = 10_000;

export const PDF_DEVICE_COLOR_COMPONENT_LIMIT = 1;

// A writer that draws every glyph as its own show (Skia, wkhtmltopdf, WPS)
// makes a long line two hundred runs; the frame still has to hold them all.
/** Other font resources a route may prove beside its own, as editable
 * companions or read-only layout witnesses: a page of a word processor
 * routinely names a dozen. */
// A TeX page draws its text from twenty-odd fonts; every one of them has
// to be a proven witness for the free space beside a run to be measured.
export const PDF_NATIVE_MAX_SUPPORTING_FONTS = 31;

export const MRAI_NATIVE_MULTI_RUN_TEXT_SCHEMA_VERSION =
  "mrai.pdf-native-text.replace.v5" as const;

export const MRAI_NATIVE_PARAGRAPH_REFLOW_SCHEMA_VERSION =
  "mrai.pdf-native-text.replace.v7" as const;

export const MRAI_NATIVE_MULTICOLUMN_PARAGRAPH_REFLOW_SCHEMA_VERSION =
  "mrai.pdf-native-text.replace.v8" as const;

export const MRAI_NATIVE_ORTHOGONAL_PARAGRAPH_REFLOW_SCHEMA_VERSION =
  "mrai.pdf-native-text.replace.v9" as const;

export const MRAI_NATIVE_ORTHOGONAL_REFLOW_TRANSFORM_SCHEMA_VERSION =
  "mrai.pdf-native-text.orthogonal-transform.v1" as const;

export const MRAI_NATIVE_AFFINE_PARAGRAPH_REFLOW_SCHEMA_VERSION =
  "mrai.pdf-native-text.replace.v10" as const;

export const MRAI_NATIVE_AFFINE_REFLOW_TRANSFORM_SCHEMA_VERSION =
  "mrai.pdf-native-text.affine-transform.v1" as const;

export const MRAI_NATIVE_PARAGRAPH_COLUMN_TOPOLOGY_SCHEMA_VERSION =
  "mrai.pdf-native-text.column-topology.v1" as const;

export const MRAI_NATIVE_FIXED_FRAME_LINE_PLAN_SCHEMA_VERSION =
  "mrai.pdf-native-text.line-plan.v1" as const;

export const MRAI_NATIVE_ADJACENT_TEXT_OBJECT_LINE_PLAN_SCHEMA_VERSION =
  "mrai.pdf-native-text.line-plan.v2" as const;

export const MRAI_NATIVE_ADAPTIVE_FIXED_FRAME_LINE_PLAN_SCHEMA_VERSION =
  "mrai.pdf-native-text.line-plan.v3" as const;

export const MRAI_NATIVE_ADAPTIVE_ADJACENT_TEXT_OBJECT_LINE_PLAN_SCHEMA_VERSION =
  "mrai.pdf-native-text.line-plan.v4" as const;

export const MRAI_NATIVE_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION =
  "mrai.pdf-native-text.line-plan.v5" as const;

export function isMraiNativeMultiRunLinePlan(schema: unknown): boolean {
  return schema === MRAI_NATIVE_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION ||
    schema === MRAI_NATIVE_ADAPTIVE_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION || schema === MRAI_NATIVE_PROJECTED_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION;
}

export function isMraiNativeVerticalFontEncodingKind(
  kind: PdfNativeFontEncodingKind,
): boolean {
  return kind.startsWith("identity-v-") || kind.startsWith("cmap-v-");
}

export const PDF_NATIVE_STANDARD14_LATIN_BASE_FONTS = Object.freeze([
  "Helvetica",
  "Helvetica-Bold",
  "Helvetica-Oblique",
  "Helvetica-BoldOblique",
  "Times-Roman",
  "Times-Bold",
  "Times-Italic",
  "Times-BoldItalic",
  "Courier",
  "Courier-Bold",
  "Courier-Oblique",
  "Courier-BoldOblique",
] as const);

export const STANDARD14_LATIN_BASE_FONT_SET = new Set<string>(
  PDF_NATIVE_STANDARD14_LATIN_BASE_FONTS,
);

export function isStandard14LatinBaseFont(
  value: string,
): value is PdfNativeStandard14LatinBaseFont {
  return STANDARD14_LATIN_BASE_FONT_SET.has(value);
}

export function isMraiNativeStandard14EncodingKind(
  kind: string,
): kind is PdfNativeStandard14EncodingKind {
  return (
    kind === "standard14-winansi" ||
    kind === "standard14-standard" ||
    kind === "standard14-encoded"
  );
}

export function exactDeviceColor(
  kind: PdfNativeDeviceColor["kind"],
  components: readonly number[],
): PdfNativeDeviceColor {
  const expected = kind === "gray" ? 1 : kind === "rgb" ? 3 : 4;
  if (
    components.length !== expected ||
    components.some(
      (component) =>
        !Number.isFinite(component) ||
        component < 0 ||
        component > PDF_DEVICE_COLOR_COMPONENT_LIMIT,
    )
  ) {
    refuse("Native device-color operands are invalid.");
  }
  return kind === "gray"
    ? Object.freeze({
        kind,
        components: Object.freeze([components[0]] as const),
      })
    : kind === "rgb"
      ? Object.freeze({
          kind,
          components: Object.freeze([
            components[0],
            components[1],
            components[2],
          ] as const),
        })
      : Object.freeze({
          kind,
          components: Object.freeze([
            components[0],
            components[1],
            components[2],
            components[3],
          ] as const),
        });
}

export function refuse(message: string): never {
  // Routes refuse as their way of answering, hundreds of times per box;
  // what a refusal carries is its message, never where it was raised.
  const limit = Error.stackTraceLimit;
  Error.stackTraceLimit = 0;
  try {
    throw new PdfNativeTransactionError("unsupported-document", message);
  } finally {
    Error.stackTraceLimit = limit;
  }
}

/** Whether a refusal reports that the source font has no glyph for the text.
 *
 * A selection that weighed several candidate routes names the boundary each
 * one hit, so the missing glyph arrives inside that list rather than on its
 * own. The font extension replaces the run's font outright and proves its own
 * route from scratch, so it is the answer whenever one of those boundaries was
 * the missing glyph; losing the signal to the other routes' messages left the
 * extension unreachable on any real document. */
export function isMraiNativeListTextMissingGlyphRefusal(
  message: unknown,
): boolean {
  return (
    typeof message === "string" &&
    (message === MRAI_NATIVE_LIST_TEXT_MISSING_GLYPH_MESSAGE ||
      message.includes(`: ${MRAI_NATIVE_LIST_TEXT_MISSING_GLYPH_MESSAGE}`))
  );
}

export function isMraiNativeListTextMissingGlyphError(
  error: unknown,
): error is PdfNativeTransactionError {
  return (
    error instanceof PdfNativeTransactionError &&
    error.code === "unsupported-document" &&
    isMraiNativeListTextMissingGlyphRefusal(error.message)
  );
}

export function invalid(message: string): never {
  throw new PdfNativeTransactionError("invalid-request", message);
}

export function refsEqual(
  left: PdfNativeIndirectObjectRef,
  right: PdfNativeIndirectObjectRef,
): boolean {
  return (
    left.objectNumber === right.objectNumber &&
    left.generation === right.generation
  );
}

/** The text pdf.js extracts from a show whose spacing is kerning, not spaces.
 *
 * A producer may set a whole line as one `TJ` with no space glyph at all and
 * let the adjustments do the spacing; pdf.js turns an adjustment wider than
 * its tracking threshold into a literal space, so the text the reader reports
 * contains characters this source never drew. Rebuilding that projection is
 * what lets such a selection be recognised at all. The thresholds are pdf.js
 * own (TRACKING_SPACE_FACTOR 0.102 of the font size). */
/** pdf.js turns a wide `TJ` adjustment into a space it can report even though
 * the source draws no space glyph there. `TRACKING_SPACE_FACTOR` is the
 * threshold it applies (pdfjs-dist legacy/build/pdf.worker.mjs). */
export const PDFJS_TRACKING_SPACE_FACTOR = 0.102;

export function selectionVisualLineIndexes(
  items: readonly PdfNativeListTextSelectionItemEvidence[],
): readonly number[] {
  const indexes: number[] = [];
  let lineIndex = 0;
  let previous: PdfNativeListTextSelectionItemEvidence | undefined;
  // The line's own baseline: its tallest item so far. A formula's scripts
  // stand off it, a superscript above and a subscript below (TeX stacks
  // `a^*_B` at one pen, and sets `a^2_N` script after script): two scripts in
  // a row stand further apart than either stands from the line, and are
  // still one line. A new line stands off both the item before it and the
  // line's own baseline. (Mirror of replay's _selection_visual_line_indexes.)
  let anchor: PdfNativeListTextSelectionItemEvidence | undefined;
  // How far the item stands off `from`'s baseline: across that baseline,
  // not along the item's own up axis, which a synthetic oblique
  // (`9.5 0 2.0193 9.5 Tm`, a faux italic) leans along the line -- an
  // upright word after an italic one on its baseline is on its line.
  const apart = (from: PdfNativeListTextSelectionItemEvidence, item: PdfNativeListTextSelectionItemEvidence) => {
    const xAxisScale = Math.hypot(from.transform[0], from.transform[1]);
    const deltaX = item.transform[4] - from.transform[4];
    const deltaY = item.transform[5] - from.transform[5];
    const baselineDistance =
      Number.isFinite(xAxisScale) && xAxisScale > 0
        ? Math.abs((deltaY * from.transform[0] - deltaX * from.transform[1]) / xAxisScale)
        : Math.hypot(deltaX, deltaY);
    return baselineDistance > Math.max(item.height, from.height) * 0.55;
  };
  for (const item of items) {
    if (previous && apart(previous, item) && (!anchor || apart(anchor, item))) {
      lineIndex += 1;
      anchor = undefined;
    }
    if (!anchor || item.height > anchor.height) anchor = item;
    indexes.push(lineIndex);
    previous = item;
  }
  return Object.freeze(indexes);
}

export const CJK_VISUAL_LINE_BOUNDARY =
  /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u;

export const LONG_UNBROKEN_VISUAL_LINE_FRAGMENT = 24;

export function visualLineBoundaryNeedsSpace(
  previous: string,
  current: string,
): boolean {
  if (/\s$/u.test(previous) || /^\s/u.test(current)) return false;
  // A word broken at its hyphen across two lines ("low/mid/high-" and
  // "level") is one word: joined with a space, a paragraph set again drew
  
  if (/\p{L}[-\u2010]$/u.test(previous) && /^\p{L}/u.test(current)) return false;
  const previousLast = [...previous].at(-1) ?? "";
  const currentFirst = [...current][0] ?? "";
  if (
    CJK_VISUAL_LINE_BOUNDARY.test(previousLast) &&
    CJK_VISUAL_LINE_BOUNDARY.test(currentFirst)
  ) {
    return false;
  }
  if (
    !/\s/u.test(previous) &&
    !/\s/u.test(current) &&
    (Array.from(previous).length >= LONG_UNBROKEN_VISUAL_LINE_FRAGMENT ||
      Array.from(current).length >= LONG_UNBROKEN_VISUAL_LINE_FRAGMENT)
  ) {
    // A URL/hash/identifier may be painted as several absolute text objects
    // without gaining whitespace in the source. Reintroducing a synthetic
    // space makes an otherwise unchanged edit overflow its authenticated
    // frame and permanently changes copy/search text.
    return false;
  }
  return true;
}

/** A word gap between two runs of one line, as a share of the size: the
 * reader's space (pdf.js reads a gap past 0.102 of the size as one). */
const RUN_WORD_GAP_EM = 0.25;

export function summarizePdfNativeSelectionItems(
  items: readonly PdfNativeListTextSelectionItemEvidence[],
): string {
  const lineIndexes = selectionVisualLineIndexes(items);
  // Two runs of one line a word gap apart, the second off the first's
  // baseline, read with a space between them:
  // the reader ends its line at a word set in a leaned matrix of its own (a
  // superscript italic), reports the line end as an empty item and drops the
  // space before it; joined as they stood, "round" and "one" read
  // "roundone" when the box was opened again (WHO sitrep p1, the side
  
  const texts = items.map((item, index) => {
    const previous = index > 0 ? items[index - 1] : null;
    if (!previous || lineIndexes[index - 1] !== lineIndexes[index] || /\s$/u.test(previous.str) ||
      /^\s/u.test(item.str) || !/[\p{L}\p{N}]$/u.test(previous.str) || !/^[\p{L}\p{N}]/u.test(item.str)) return item.str;
    const axis = Math.hypot(previous.transform[0], previous.transform[1]);
    if (!(axis > 0)) return item.str;
    const gap = ((item.transform[4] - previous.transform[4]) * previous.transform[0] +
      (item.transform[5] - previous.transform[5]) * previous.transform[1]) / axis - previous.width;
    // Only where the run after stands off the line's baseline (a script):
    // runs on one baseline the reader has already read apart or together.
    const offBaseline = Math.abs(((item.transform[5] - previous.transform[5]) * previous.transform[0] -
      (item.transform[4] - previous.transform[4]) * previous.transform[1]) / axis);
    const size = Math.max(previous.height, item.height);
    return size > 0 && offBaseline > size * 0.15 && gap >= size * RUN_WORD_GAP_EM ? ` ${item.str}` : item.str;
  });
  return summarizeVisualRunTexts(texts, lineIndexes);
}

const STRONG_RTL_LETTER = /[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/gu;
const STRONG_LTR_LETTER = /[A-Za-z\u00C0-\u024F\u0370-\u03FF\u0400-\u04FF\u3040-\u30FF\u4E00-\u9FFF\uAC00-\uD7AF]/gu;






function runsInReadingOrder(texts: readonly string[], lineIndexes: readonly number[]): string[] {
  const result = [...texts];
  for (let start = 0; start < texts.length;) {
    let end = start + 1;
    while (end < texts.length && lineIndexes[end] === lineIndexes[start]) end += 1;
    const line = texts.slice(start, end).join("");
    const rtl = line.match(STRONG_RTL_LETTER)?.length ?? 0;
    if (end - start > 1 && rtl > 0 && rtl > (line.match(STRONG_LTR_LETTER)?.length ?? 0)) {
      texts.slice(start, end).reverse().forEach((text, offset) => {
        const [, leading, body, trailing] = /^(\s*)([\s\S]*?)(\s*)$/u.exec(text)!;
        result[start + offset] = trailing + body + leading;
      });
    }
    start = end;
  }
  return result;
}

export function summarizeVisualRunTexts(
  texts: readonly string[],
  lineIndexes: readonly number[],
): string {
  if (texts.length !== lineIndexes.length) {
    refuse("Visual text and line evidence differ.");
  }
  texts = runsInReadingOrder(texts, lineIndexes);
  let summary = "";
  let previousLine = -1;
  for (const [index, text] of texts.entries()) {
    if (!text) continue;
    if (summary && lineIndexes[index] !== previousLine) {
      if (visualLineBoundaryNeedsSpace(summary, text)) summary += " ";
    }
    summary += text;
    previousLine = lineIndexes[index];
  }
  return summary;
}

export function conformalTextLinearPart(matrix: PdfMatrix): boolean {
  const [a, b, c, d] = matrix;
  const scale = Math.hypot(a, b);
  return !(
    !Number.isFinite(scale) ||
    scale <= 0.000001 ||
    !nearlyEqual(Math.hypot(c, d), scale, 0.001) ||
    !nearlyEqual(a * c + b * d, 0, 0.001) ||
    a * d - b * c <= 0
  );
}

export function orthogonalTextRotation(matrix: PdfMatrix): 0 | 90 | 180 | 270 | null {
  if (!conformalTextLinearPart(matrix)) return null;
  const [a, b, c, d] = matrix;
  const scale = Math.hypot(a, b);
  const normalized = [a / scale, b / scale, c / scale, d / scale];
  if (
    nearlyEqual(normalized[0], 1, 0.001) &&
    nearlyEqual(normalized[1], 0, 0.001) &&
    nearlyEqual(normalized[2], 0, 0.001) &&
    nearlyEqual(normalized[3], 1, 0.001)
  ) {
    return 0;
  }
  if (
    nearlyEqual(normalized[0], 0, 0.001) &&
    nearlyEqual(normalized[1], 1, 0.001) &&
    nearlyEqual(normalized[2], -1, 0.001) &&
    nearlyEqual(normalized[3], 0, 0.001)
  ) {
    return 90;
  }
  if (
    nearlyEqual(normalized[0], -1, 0.001) &&
    nearlyEqual(normalized[1], 0, 0.001) &&
    nearlyEqual(normalized[2], 0, 0.001) &&
    nearlyEqual(normalized[3], -1, 0.001)
  ) {
    return 180;
  }
  if (
    nearlyEqual(normalized[0], 0, 0.001) &&
    nearlyEqual(normalized[1], -1, 0.001) &&
    nearlyEqual(normalized[2], 1, 0.001) &&
    nearlyEqual(normalized[3], 0, 0.001)
  ) {
    return 270;
  }
  return null;
}

export function exactSelectionItem(
  item: PdfNativeListTextSelectionItemEvidence,
): PdfNativeListTextSelectionItemEvidence {
  if (
    !Number.isSafeInteger(item.itemIndex) ||
    item.itemIndex < 0 ||
    item.str.length === 0 ||
    item.str.length > PDF_NATIVE_MAX_ITEM_TEXT_LENGTH ||
    !item.fontName.trim() ||
    item.fontName.length > 256 ||
    item.transform.length !== 6 ||
    item.transform.some((value) => !Number.isFinite(value)) ||
    !Number.isFinite(item.width) ||
    item.width <= 0 ||
    !Number.isFinite(item.height) ||
    item.height <= 0 ||
    typeof item.hasEOL !== "boolean"
  ) {
    invalid(
      "Native text editing requires complete source-bound pdf.js text items.",
    );
  }
  return Object.freeze({
    itemIndex: item.itemIndex,
    str: item.str,
    fontName: item.fontName,
    transform: Object.freeze([
      ...item.transform,
    ]) as PdfNativeListTextSelectionItemEvidence["transform"],
    width: item.width,
    height: item.height,
    hasEOL: item.hasEOL,
  });
}

export function exactSelection(
  selection: PdfNativeListTextSelectionEvidence,
  sourceSha256: string,
): PdfNativeListTextSelectionEvidence {
  if (
    !/^[a-f\d]{64}$/i.test(selection.sourceSha256) ||
    selection.sourceSha256.toLowerCase() !== sourceSha256.toLowerCase()
  ) {
    throw new PdfNativeTransactionError(
      "source-mismatch",
      "Text selection does not belong to the pinned PDF bytes.",
    );
  }
  if (
    !Number.isSafeInteger(selection.pageNumber) ||
    selection.pageNumber < 1 ||
    !Number.isSafeInteger(selection.selectionRevision) ||
    selection.selectionRevision < 1 ||
    selection.startOffset !== 0
  ) {
    invalid("Native text selection identity is incomplete or stale.");
  }
  const item = exactSelectionItem(selection.item);
  const items = selection.items?.map(exactSelectionItem);
  if (items) {
    const summary = summarizePdfNativeSelectionItems(items);
    const first = items[0];
    const last = items.at(-1);
    if (
      items.length < 2 ||
      items.length > PDF_NATIVE_MAX_FIXED_FRAME_ITEMS ||
      !first ||
      !last ||
      summary.length > 2_000 ||
      selection.endOffset !== summary.length ||
      item.itemIndex !== first.itemIndex ||
      item.str !== summary ||
      item.transform.some((value, index) => value !== first.transform[index]) ||
      item.width !== items.reduce((total, entry) => total + entry.width, 0) ||
      item.height !== first.height ||
      item.hasEOL !== last.hasEOL ||
      items.some(
        (entry, index) =>
          index > 0 &&
          (entry.itemIndex <= items[index - 1].itemIndex ||
            entry.itemIndex - items[index - 1].itemIndex >
              PDF_NATIVE_MAX_FIXED_FRAME_ITEM_INDEX_GAP),
      )
    ) {
      invalid(
        "Fixed-frame editing requires 2–64 ordered complete text items with stable line boundaries.",
      );
    }
  } else if (selection.endOffset !== item.str.length) {
    invalid(
      "Native text editing requires one complete, source-bound pdf.js text item.",
    );
  }
  return Object.freeze({
    sourceSha256: selection.sourceSha256.toLowerCase(),
    selectionRevision: selection.selectionRevision,
    pageNumber: selection.pageNumber,
    startOffset: selection.startOffset,
    endOffset: selection.endOffset,
    item,
    ...(items ? { items: Object.freeze(items) } : {}),
  });
}

export function nearlyEqual(left: number, right: number, tolerance = 0.01): boolean {
  return Math.abs(left - right) <= tolerance;
}

export function linePlanSelectionItemIndexes(
  linePlan: PdfNativeFixedFrameLinePlan,
  selectionItemCount: number,
): readonly number[] | null {
  const explicit = linePlan.lines.some(
    (line) => line.selectionItemIndex !== undefined,
  );
  if (!explicit) {
    return linePlan.lines.length === selectionItemCount
      ? Object.freeze(linePlan.lines.map((_, index) => index))
      : null;
  }
  if (
    !isMraiNativeMultiRunLinePlan(linePlan.schemaVersion) ||
    linePlan.lines.some(
      (line) =>
        !Number.isSafeInteger(line.selectionItemIndex) ||
        line.selectionItemIndex! < 0 ||
        line.selectionItemIndex! >= selectionItemCount,
    )
  ) {
    return null;
  }
  const indexes = linePlan.lines.map((line) => line.selectionItemIndex!);
  if (
    indexes[0] !== 0 ||
    indexes.at(-1) !== selectionItemCount - 1 ||
    indexes.some(
      (itemIndex, index) =>
        index > 0 &&
        itemIndex !== indexes[index - 1] &&
        itemIndex !== indexes[index - 1] + 1,
    )
  ) {
    return null;
  }
  return Object.freeze(indexes);
}

export function matchesPositionedLineItem(display: PdfNativeListTextSelectionItemEvidence,
  source: PdfNativeListTextSelectionItemEvidence): boolean {
  return display.itemIndex === source.itemIndex && display.fontName === source.fontName &&
    (display.str === source.str || source.str.startsWith(display.str) && /^ +$/.test(source.str.slice(display.str.length))) &&
    nearlyEqual(display.width, source.width, Math.max(.05, source.width * .005)) &&
    nearlyEqual(display.height, source.height, .05) &&
    display.transform.every((value, component) => nearlyEqual(value, source.transform[component]));
}

/**
 * Whether a positioned line proves `item` less the gap the reader read as its
 * leading space.
 *
 * pdf.js reads the gap between the last glyph it measured and the next one it
 * draws as a space. Where that next glyph starts a new item -- the font
 * changes, or the gap is wider than a space -- it reports the space as an item
 * of its own: zero-height, as wide as the gap, standing where the previous
 * glyph ended (compareWithLastPosition, addFakeSpaces). The viewer folds that
 * space into the item after it (normalizePdfJsNativeParagraphItems), so the
 * item starts at the gap and no show starts there. No show draws the space
 * either: it is the separator between two items, so the line proves the
 * reader's own item for the show -- one item further on, its text without
 * the space, starting where the gap ends and ending where the item does.
 * The gap is held to the source at both ends: it starts where `previous`, the
 * item the line before proves on the same visual line, ends, and it ends where
 * this line's show starts; each line's own proof binds its item to its show.
 * It is also wide enough for the reader to read as a space. Mirror of the
 * server's `_positioned_leading_gap_line_item`.
 */
export function positionedLeadingGapLineItem(
  item: PdfNativeListTextSelectionItemEvidence,
  line: PdfNativeListTextSelectionItemEvidence,
  previous: PdfNativeListTextSelectionItemEvidence | undefined,
): boolean {
  if (!previous || !/^ [^ ]/u.test(item.str) || line.str !== item.str.slice(1)) return false;
  const [a, b, c, d, x, y] = item.transform;
  const gap = line.transform[4] - x;
  return (
    line.itemIndex === item.itemIndex + 1 &&
    line.fontName === item.fontName &&
    line.height === item.height &&
    line.hasEOL === item.hasEOL &&
    a > 0 &&
    d > 0 &&
    Math.abs(b) <= POSITIONED_LEADING_GAP_AXIS_TOLERANCE &&
    Math.abs(c) <= POSITIONED_LEADING_GAP_AXIS_TOLERANCE &&
    line.transform[0] === a &&
    line.transform[1] === b &&
    line.transform[2] === c &&
    line.transform[3] === d &&
    line.transform[5] === y &&
    line.width > 0 &&
    gap >= PDFJS_TRACKING_SPACE_FACTOR * Math.hypot(c, d) &&
    nearlyEqual(line.transform[4] + line.width, x + item.width) &&
    nearlyEqual(previous.transform[4] + previous.width, x) &&
    nearlyEqual(previous.transform[5], y)
  );
}

/** The reader's items are upright where a gap is read along the baseline. */
export const POSITIONED_LEADING_GAP_AXIS_TOLERANCE = 0.001;

/**
 * For each line of a positioned paragraph, whether the item it proves is the
 * paragraph's item less a leading gap (positionedLeadingGapLineItem); null
 * when a line proves neither the paragraph's item nor that.
 */
export function positionedLineLeadingGaps(
  items: readonly PdfNativeListTextSelectionItemEvidence[],
  lineItems: readonly (PdfNativeListTextSelectionItemEvidence | undefined)[],
): readonly boolean[] | null {
  if (items.length !== lineItems.length) return null;
  const lineIndexes = selectionVisualLineIndexes(items);
  const gaps: boolean[] = [];
  for (const [index, item] of items.entries()) {
    const line = lineItems[index];
    if (!line) return null;
    if (matchesPositionedLineItem(item, line)) {
      gaps.push(false);
    } else if (
      index > 0 &&
      lineIndexes[index] === lineIndexes[index - 1] &&
      positionedLeadingGapLineItem(item, line, lineItems[index - 1])
    ) {
      gaps.push(true);
    } else {
      return null;
    }
  }
  return Object.freeze(gaps);
}

/**
 * The one item of a multi-item selection an edit changes, with that item's
 * new text: the edit's changed range (the box text less the prefix and
 * suffix it keeps) lies inside one item's text as the box joins them
 * (summarizeVisualRunTexts). Null when it spans items or touches the space
 * the box puts between two visual lines. An item left with nothing but the
 * reader's gap space is emptied.
 */
export function singleItemEditOfSelection(
  selection: PdfNativeListTextSelectionEvidence,
  replacementText: string,
): Readonly<{ selection: PdfNativeListTextSelectionEvidence; replacementText: string }> | null {
  const items = selection.items;
  if (!items || items.length < 2) return null;
  const lineIndexes = selectionVisualLineIndexes(items);
  const spans: Array<Readonly<{ index: number; start: number; end: number }>> = [];
  let summary = "";
  let previousLine = -1;
  for (const [index, item] of items.entries()) {
    if (!item.str) continue;
    if (summary && lineIndexes[index] !== previousLine && visualLineBoundaryNeedsSpace(summary, item.str)) summary += " ";
    spans.push(Object.freeze({ index, start: summary.length, end: summary.length + item.str.length }));
    summary += item.str;
    previousLine = lineIndexes[index];
  }
  if (summary !== summarizePdfNativeSelectionItems(items)) return null;
  const original = Array.from(summary), written = Array.from(replacementText);
  let prefix = 0;
  while (prefix < original.length && prefix < written.length && original[prefix] === written[prefix]) prefix += 1;
  let suffix = 0;
  while (suffix < original.length - prefix && suffix < written.length - prefix &&
    original[original.length - 1 - suffix] === written[written.length - 1 - suffix]) suffix += 1;
  if (prefix === original.length && prefix === written.length) return null;
  // Code-unit offsets of the changed range in the box text.
  const changeStart = original.slice(0, prefix).join("").length;
  const changeEnd = summary.length - original.slice(original.length - suffix).join("").length;
  const inserted = written.slice(prefix, written.length - suffix).join("");
  const owner = spans.find((span) => span.start <= changeStart && changeEnd <= span.end &&
    // A pure insertion at a boundary belongs to the item it extends.
    (changeStart < changeEnd || changeStart < span.end || span === spans.at(-1)));
  if (!owner) return null;
  const item = items[owner.index];
  let text = item.str.slice(0, changeStart - owner.start) + inserted + item.str.slice(changeEnd - owner.start);
  if (!text.trim()) text = "";
  return Object.freeze({
    selection: Object.freeze({ ...selection, items: undefined, item, startOffset: 0, endOffset: item.str.length }),
    replacementText: text,
  });
}

/**
 * The bytes the run itself draws in, inside the stream that draws them.
 * A replacement wide enough to move the rest of the line rewrites the later
 * operators that carry it, so the plan's operator span runs on past the run;
 * a delete or a rule belongs to the run alone and must not reach into the
 * text a line shift would only have moved.
 */
export function pdfNativeListTextSourceRunSpan(
  proof:
    | Readonly<{
        evidence: PdfNativeListTextFidelityEvidence;
        sourceRun?: PdfNativeListTextSourceRun;
      }>
    | null
    | undefined,
): Readonly<{ start: number; end: number }> | undefined {
  if (!proof) return undefined;
  // A paragraph rewrite's span runs from its first show to the end of its
  // last show's text object, over text that is not the paragraph's: it names
  
  // opened: its proof grants no native write at all.
  // Nor a box whose words are fixed (PdfNativeListTextSelectionProof.wordsFixed).
  if (proof.evidence.paragraphRewrite || (proof as { fallbackText?: unknown }).fallbackText ||
    (proof as { wordsFixed?: unknown }).wordsFixed) return undefined;
  const span = proof.sourceRun?.span ?? proof.evidence.textOperatorSpan;
  const showEnd = proof.evidence.lineShift?.showEnd;
  if (showEnd === undefined) return span;
  // The line shift reads the same stream the operator span does; the source
  // run is that span less the owning stream's offset into the page content.
  const end = showEnd + (span.end - proof.evidence.textOperatorSpan.end);
  return end > span.start && end < span.end
    ? Object.freeze({ start: span.start, end })
    : span;
}

export function unknownRecord(value: unknown): Record<string, unknown> | null {
  return typeof value === "object" && value !== null && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

export function readFiniteNumber(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

export function readBoundedEvidenceText(value: unknown): string | null {
  return typeof value === "string" && value.length > 0 && value.length <= 2_000
    ? value
    : null;
}

export function readCodeByteArray(value: unknown): readonly number[] | null {
  if (
    !Array.isArray(value) ||
    value.length === 0 ||
    value.length > 64 * 1_024 ||
    value.some(
      (entry) =>
        !Number.isSafeInteger(entry) ||
        Number(entry) < 0 ||
        Number(entry) > 255,
    )
  ) {
    return null;
  }
  return Object.freeze(value.map(Number));
}

export function readEvidenceParagraphReflow(
  value: unknown,
): PdfNativeParagraphReflowMutationEvidence | null {
  const record = unknownRecord(value);
  const frameLeft = readFiniteNumber(record?.frameLeft);
  const frameRight = readFiniteNumber(record?.frameRight);
  const downstreamTranslationY = readFiniteNumber(
    record?.downstreamTranslationY,
  );
  const orthogonalRecord = unknownRecord(record?.orthogonalTransform);
  const affineRecord = unknownRecord(record?.affineTransform);
  const localToPageLinear = orthogonalRecord?.localToPageLinear;
  const pageTranslationX = readFiniteNumber(
    orthogonalRecord?.pageTranslationX,
  );
  const pageTranslationY = readFiniteNumber(
    orthogonalRecord?.pageTranslationY,
  );
  const affineLocalToPageLinear = affineRecord?.localToPageLinear;
  const affinePageTranslationX = readFiniteNumber(
    affineRecord?.pageTranslationX,
  );
  const affinePageTranslationY = readFiniteNumber(
    affineRecord?.pageTranslationY,
  );
  const originalSelection = readEvidenceSelection(record?.originalSelection);
  const replacementSelection = readEvidenceSelection(
    record?.replacementSelection,
  );
  const originalOperatorBytes = readCodeByteArray(
    record?.originalOperatorBytes,
  );
  const replacementOperatorBytes = readCodeByteArray(
    record?.replacementOperatorBytes,
  );
  const readLines = (
    value: unknown,
  ): readonly PdfNativeParagraphReflowLineEvidence[] | null => {
    if (
      !Array.isArray(value) ||
      value.length < 1 ||
      value.length > PDF_NATIVE_MAX_FIXED_FRAME_ITEMS
    ) {
      return null;
    }
    const lines: PdfNativeParagraphReflowLineEvidence[] = [];
    for (const entry of value) {
      const line = unknownRecord(entry);
      const text = readBoundedEvidenceText(line?.text);
      const displayText = readBoundedEvidenceText(line?.displayText);
      const codeBytes = readCodeByteArray(line?.codeBytes);
      const textX = readFiniteNumber(line?.textX);
      const baseline = readFiniteNumber(line?.baseline);
      const advance = readFiniteNumber(line?.advance);
      if (
        !line ||
        !text ||
        !displayText ||
        !codeBytes ||
        textX === null ||
        baseline === null ||
        advance === null ||
        advance <= 0
      ) {
        return null;
      }
      lines.push(
        Object.freeze({
          text,
          displayText,
          codeBytes,
          textX,
          baseline,
          advance,
        }),
      );
    }
    return Object.freeze(lines);
  };
  const originalLines = readLines(record?.originalLines);
  const replacementLines = readLines(record?.replacementLines);
  const topologyRecord = unknownRecord(record?.columnTopology);
  const targetColumnLeft = readFiniteNumber(topologyRecord?.targetColumnLeft);
  const targetColumnRight = readFiniteNumber(topologyRecord?.targetColumnRight);
  const visiblePageBottom = readFiniteNumber(topologyRecord?.visiblePageBottom);
  const visiblePageTop = readFiniteNumber(topologyRecord?.visiblePageTop);
  const downstreamRoles = topologyRecord?.downstreamRoles;
  const columnTopology =
    topologyRecord?.schemaVersion ===
      MRAI_NATIVE_PARAGRAPH_COLUMN_TOPOLOGY_SCHEMA_VERSION &&
    targetColumnLeft !== null &&
    targetColumnRight !== null &&
    targetColumnRight > targetColumnLeft &&
    visiblePageBottom !== null &&
    visiblePageTop !== null &&
    visiblePageTop > visiblePageBottom &&
    Array.isArray(downstreamRoles) &&
    downstreamRoles.length >= 2 &&
    downstreamRoles.length <= 32 &&
    downstreamRoles.every(
      (role) => role === "translate" || role === "preserve",
    ) &&
    downstreamRoles.includes("translate") &&
    downstreamRoles.includes("preserve")
      ? Object.freeze({
          schemaVersion: MRAI_NATIVE_PARAGRAPH_COLUMN_TOPOLOGY_SCHEMA_VERSION,
          targetColumnLeft,
          targetColumnRight,
          visiblePageBottom,
          visiblePageTop,
          downstreamRoles: Object.freeze(
            downstreamRoles as ("translate" | "preserve")[],
          ),
        })
      : null;
  const orthogonalTransform =
    orthogonalRecord?.schemaVersion ===
      MRAI_NATIVE_ORTHOGONAL_REFLOW_TRANSFORM_SCHEMA_VERSION &&
    Array.isArray(localToPageLinear) &&
    localToPageLinear.length === 4 &&
    localToPageLinear.every(
      (value) => typeof value === "number" && Number.isFinite(value),
    ) &&
    orthogonalTextRotation(
      Object.freeze([
        Number(localToPageLinear[0]),
        Number(localToPageLinear[1]),
        Number(localToPageLinear[2]),
        Number(localToPageLinear[3]),
        0,
        0,
      ]) as PdfMatrix,
    ) !== null &&
    orthogonalTextRotation(
      Object.freeze([
        Number(localToPageLinear[0]),
        Number(localToPageLinear[1]),
        Number(localToPageLinear[2]),
        Number(localToPageLinear[3]),
        0,
        0,
      ]) as PdfMatrix,
    ) !== 0 &&
    pageTranslationX !== null &&
    pageTranslationY !== null &&
    downstreamTranslationY !== null &&
    nearlyEqual(
      pageTranslationX,
      Number(localToPageLinear[2]) * downstreamTranslationY,
      0.001,
    ) &&
    nearlyEqual(
      pageTranslationY,
      Number(localToPageLinear[3]) * downstreamTranslationY,
      0.001,
    )
      ? Object.freeze({
          schemaVersion:
            MRAI_NATIVE_ORTHOGONAL_REFLOW_TRANSFORM_SCHEMA_VERSION,
          localToPageLinear: Object.freeze(
            localToPageLinear.map(Number),
          ) as readonly [number, number, number, number],
          pageTranslationX,
          pageTranslationY,
        })
      : null;
  const affineMatrix =
    Array.isArray(affineLocalToPageLinear) &&
    affineLocalToPageLinear.length === 4 &&
    affineLocalToPageLinear.every(
      (value) => typeof value === "number" && Number.isFinite(value),
    )
      ? (Object.freeze([
          Number(affineLocalToPageLinear[0]),
          Number(affineLocalToPageLinear[1]),
          Number(affineLocalToPageLinear[2]),
          Number(affineLocalToPageLinear[3]),
          0,
          0,
        ]) as PdfMatrix)
      : null;
  const affineTransform =
    affineRecord?.schemaVersion ===
      MRAI_NATIVE_AFFINE_REFLOW_TRANSFORM_SCHEMA_VERSION &&
    affineMatrix !== null &&
    conformalTextLinearPart(affineMatrix) &&
    orthogonalTextRotation(affineMatrix) === null &&
    affinePageTranslationX !== null &&
    affinePageTranslationY !== null &&
    downstreamTranslationY !== null &&
    nearlyEqual(
      affinePageTranslationX,
      affineMatrix[2] * downstreamTranslationY,
      0.001,
    ) &&
    nearlyEqual(
      affinePageTranslationY,
      affineMatrix[3] * downstreamTranslationY,
      0.001,
    )
      ? Object.freeze({
          schemaVersion: MRAI_NATIVE_AFFINE_REFLOW_TRANSFORM_SCHEMA_VERSION,
          localToPageLinear: Object.freeze(
            affineMatrix.slice(0, 4),
          ) as readonly [number, number, number, number],
          pageTranslationX: affinePageTranslationX,
          pageTranslationY: affinePageTranslationY,
        })
      : null;
  if (
    !record ||
    frameLeft === null ||
    frameRight === null ||
    frameRight <= frameLeft ||
    !originalLines ||
    !replacementLines ||
    originalLines.length === replacementLines.length ||
    !originalOperatorBytes ||
    !replacementOperatorBytes ||
    !originalSelection ||
    !replacementSelection ||
    !Number.isSafeInteger(record.downstreamBlockCount) ||
    Number(record.downstreamBlockCount) < 1 ||
    Number(record.downstreamBlockCount) > 32 ||
    downstreamTranslationY === null ||
    downstreamTranslationY === 0 ||
    (record.columnTopology !== undefined && columnTopology === null) ||
    (record.orthogonalTransform !== undefined && orthogonalTransform === null) ||
    (record.affineTransform !== undefined && affineTransform === null) ||
    [columnTopology, orthogonalTransform, affineTransform].filter(Boolean)
      .length > 1 ||
    (columnTopology !== null &&
      (columnTopology.downstreamRoles.length !==
        Number(record.downstreamBlockCount) ||
        !nearlyEqual(columnTopology.targetColumnLeft, frameLeft, 0.001) ||
        !nearlyEqual(columnTopology.targetColumnRight, frameRight, 0.001)))
  ) {
    return null;
  }
  return Object.freeze({
    frameLeft,
    frameRight,
    originalLines,
    replacementLines,
    originalOperatorBytes,
    replacementOperatorBytes,
    originalSelection,
    replacementSelection,
    downstreamBlockCount: Number(record.downstreamBlockCount),
    downstreamTranslationY,
    ...(orthogonalTransform ? { orthogonalTransform } : {}),
    ...(affineTransform ? { affineTransform } : {}),
    ...(columnTopology ? { columnTopology } : {}),
  });
}

export function readEvidenceRef(value: unknown): PdfNativeIndirectObjectRef | null {
  const record = unknownRecord(value);
  if (
    !record ||
    !Number.isSafeInteger(record.objectNumber) ||
    Number(record.objectNumber) < 1 ||
    !Number.isSafeInteger(record.generation) ||
    Number(record.generation) < 0 ||
    Number(record.generation) > 65_535
  ) {
    return null;
  }
  return Object.freeze({
    objectNumber: Number(record.objectNumber),
    generation: Number(record.generation),
  });
}

export function readEvidenceFontEncoding(
  value: unknown,
): PdfNativeFontEncoding | null {
  const record = unknownRecord(value);
  if (!record || typeof record.kind !== "string") return null;
  const kind = record.kind;
  if (isMraiNativeStandard14EncodingKind(kind)) {
    return record.codeUnitBytes === 1 &&
      typeof record.baseFont === "string" &&
      isStandard14LatinBaseFont(record.baseFont)
      ? Object.freeze({
          kind,
          codeUnitBytes: 1,
          baseFont: record.baseFont,
        })
      : null;
  }
  if (record.kind === "simple-byte") {
    return record.codeUnitBytes === 1
      ? Object.freeze({ kind: "simple-byte", codeUnitBytes: 1 })
      : null;
  }
  if (
    (record.kind === "identity-h-cidfonttype2" ||
      record.kind === "identity-h-cidfonttype0" ||
      record.kind === "identity-v-cidfonttype2" ||
      record.kind === "identity-v-cidfonttype0") &&
    record.codeUnitBytes === 2
  ) {
    return Object.freeze({ kind: record.kind, codeUnitBytes: 2 });
  }
  if (
    (record.kind === "cmap-h-cidfonttype2" ||
      record.kind === "cmap-h-cidfonttype0" ||
      record.kind === "cmap-v-cidfonttype2" ||
      record.kind === "cmap-v-cidfonttype0") &&
    record.codeUnitBytes === 2 &&
    typeof record.cmapName === "string" &&
    isMraiNativePredefinedUnicodeCMapName(record.cmapName) &&
    record.cmapName.endsWith("-V") === record.kind.startsWith("cmap-v-")
  ) {
    return Object.freeze({
      kind: record.kind,
      codeUnitBytes: 2,
      cmapName: record.cmapName,
    });
  }
  return null;
}

export function readEvidenceResources(
  value: unknown,
): PdfNativeListTextFidelityEvidence["resources"] | null {
  const record = unknownRecord(value);
  const page = readEvidenceRef(record?.page);
  const font = readEvidenceRef(record?.font);
  if (!record || !page || !font) return null;
  const readNullableRef = (
    key: "descendantFont" | "descriptor" | "fontFile" | "toUnicode",
  ): PdfNativeIndirectObjectRef | null | undefined => {
    const entry = record[key];
    return entry === null ? null : (readEvidenceRef(entry) ?? undefined);
  };
  const descendantFont = readNullableRef("descendantFont");
  const descriptor = readNullableRef("descriptor");
  const fontFile = readNullableRef("fontFile");
  const toUnicode = readNullableRef("toUnicode");
  if (
    descendantFont === undefined ||
    descriptor === undefined ||
    fontFile === undefined ||
    toUnicode === undefined
  ) {
    return null;
  }
  // Where a font dictionary its owner writes in place of a reference lives.
  // Absent for every font that is an object of its own, which is nearly all
  // of them; present, it must name one resource and one owner spelling.
  const inlineRecord = unknownRecord(record.fontInline);
  if (record.fontInline !== undefined && !inlineRecord) return null;
  if (
    inlineRecord &&
    (typeof inlineRecord.resourceName !== "string" ||
      !RESOURCE_NAME.test(inlineRecord.resourceName) ||
      typeof inlineRecord.ownerKind !== "string" ||
      ![
        "absent-resources",
        "nested-resources",
        "resource-dictionary",
        "font-dictionary",
      ].includes(inlineRecord.ownerKind))
  ) {
    return null;
  }
  return Object.freeze({
    page,
    font,
    descendantFont,
    descriptor,
    fontFile,
    toUnicode,
    ...(inlineRecord
      ? {
          fontInline: Object.freeze({
            resourceName: inlineRecord.resourceName as string,
            ownerKind: inlineRecord.ownerKind as
              | PdfNativeResourceDictionaryKind
              | "font-dictionary",
          }),
        }
      : {}),
  });
}

export function readEvidenceLayoutWitnesses(
  value: unknown,
): readonly PdfNativeTextLayoutWitnessEvidence[] | null {
  if (value === undefined) return Object.freeze([]);
  // A record pins as many witness fonts as the planner proves beside its own
  // (a TeX page names twenty-odd, and the reader of a record must read every
  // record the planner writes).
  if (!Array.isArray(value) || value.length < 1 || value.length > PDF_NATIVE_MAX_SUPPORTING_FONTS) {
    return null;
  }
  const witnesses: PdfNativeTextLayoutWitnessEvidence[] = [];
  for (const entry of value) {
    const record = unknownRecord(entry);
    const fontEncoding = readEvidenceFontEncoding(record?.fontEncoding);
    const resources = readEvidenceResources(record?.resources);
    if (
      !record ||
      typeof record.resourceName !== "string" ||
      !RESOURCE_NAME.test(record.resourceName) ||
      !fontEncoding ||
      !resources ||
      witnesses.some(
        (witness) => witness.resourceName === record.resourceName,
      )
    ) {
      return null;
    }
    witnesses.push(
      Object.freeze({
        resourceName: record.resourceName,
        fontEncoding,
        resources,
      }),
    );
  }
  return Object.freeze(witnesses);
}

export function readEvidenceSpan(
  value: unknown,
): Readonly<{ start: number; end: number; sha256: string }> | null {
  const record = unknownRecord(value);
  if (
    !record ||
    !Number.isSafeInteger(record.start) ||
    Number(record.start) < 0 ||
    !Number.isSafeInteger(record.end) ||
    Number(record.end) <= Number(record.start) ||
    typeof record.sha256 !== "string" ||
    !/^[a-f\d]{64}$/i.test(record.sha256)
  ) {
    return null;
  }
  return Object.freeze({
    start: Number(record.start),
    end: Number(record.end),
    sha256: record.sha256.toLowerCase(),
  });
}

export function readEvidenceSelection(
  value: unknown,
): PdfNativeListTextSelectionEvidence | null {
  const record = unknownRecord(value);
  const itemRecord = unknownRecord(record?.item);
  if (
    !record ||
    !itemRecord ||
    typeof record.sourceSha256 !== "string" ||
    !/^[a-f\d]{64}$/i.test(record.sourceSha256) ||
    !Number.isSafeInteger(record.selectionRevision) ||
    Number(record.selectionRevision) < 1 ||
    !Number.isSafeInteger(record.pageNumber) ||
    Number(record.pageNumber) < 1 ||
    !Number.isSafeInteger(record.startOffset) ||
    Number(record.startOffset) < 0 ||
    !Number.isSafeInteger(record.endOffset) ||
    // An exact restore of an emptied line selects nothing.
    Number(record.endOffset) < (itemRecord.str === "" ? 0 : 1) ||
    !Number.isSafeInteger(itemRecord.itemIndex) ||
    Number(itemRecord.itemIndex) < 0 ||
    typeof itemRecord.str !== "string" ||
    typeof itemRecord.fontName !== "string" ||
    !Array.isArray(itemRecord.transform) ||
    itemRecord.transform.length !== 6 ||
    itemRecord.transform.some((entry) => readFiniteNumber(entry) === null) ||
    readFiniteNumber(itemRecord.width) === null ||
    readFiniteNumber(itemRecord.height) === null ||
    typeof itemRecord.hasEOL !== "boolean"
  ) {
    return null;
  }
  const baseSelection = {
    sourceSha256: record.sourceSha256.toLowerCase(),
    selectionRevision: Number(record.selectionRevision),
    pageNumber: Number(record.pageNumber),
    startOffset: Number(record.startOffset),
    endOffset: Number(record.endOffset),
    item: {
      itemIndex: Number(itemRecord.itemIndex),
      str: itemRecord.str,
      fontName: itemRecord.fontName,
      transform: itemRecord.transform.map(Number) as unknown as readonly [
        number,
        number,
        number,
        number,
        number,
        number,
      ],
      width: Number(itemRecord.width),
      height: Number(itemRecord.height),
      hasEOL: itemRecord.hasEOL,
    },
  };
  let selection: PdfNativeListTextSelectionEvidence = baseSelection;
  if (record.items !== undefined) {
    if (
      !Array.isArray(record.items) ||
      record.items.length > PDF_NATIVE_MAX_FIXED_FRAME_ITEMS
    )
      return null;
    const items: PdfNativeListTextSelectionItemEvidence[] = [];
    for (const entry of record.items) {
      const parsed = readEvidenceSelection({
        ...record,
        endOffset: unknownRecord(entry)?.str
          ? String(unknownRecord(entry)!.str).length
          : 0,
        item: entry,
        items: undefined,
      })?.item;
      if (!parsed) return null;
      items.push(parsed);
    }
    selection = { ...baseSelection, items: Object.freeze(items) };
  }
  try {
    return exactSelection(selection, selection.sourceSha256);
  } catch {
    return null;
  }
}

export function readEvidenceDeviceColor(value: unknown): PdfNativeDeviceColor | null {
  const record = unknownRecord(value);
  if (
    !record ||
    (record.kind !== "gray" &&
      record.kind !== "rgb" &&
      record.kind !== "cmyk") ||
    !Array.isArray(record.components)
  ) {
    return null;
  }
  try {
    return exactDeviceColor(record.kind, record.components.map(Number));
  } catch {
    return null;
  }
}

export function readEvidenceFontOption(
  value: unknown,
): PdfNativeTextFontOption | null {
  const record = unknownRecord(value);
  const encoding = readEvidenceFontEncoding(record?.encoding);
  const resources = readEvidenceResources(record?.resources);
  if (
    !record ||
    typeof record.resourceName !== "string" ||
    !RESOURCE_NAME.test(record.resourceName) ||
    typeof record.baseFont !== "string" ||
    record.baseFont.length === 0 ||
    record.baseFont.length > 200 ||
    typeof record.family !== "string" ||
    record.family.length === 0 ||
    record.family.length > 200 ||
    typeof record.bold !== "boolean" ||
    typeof record.italic !== "boolean" ||
    !encoding ||
    !resources
  ) {
    return null;
  }
  return Object.freeze({
    resourceName: record.resourceName,
    baseFont: record.baseFont,
    family: record.family,
    bold: record.bold,
    italic: record.italic,
    encoding,
    resources,
  });
}

export function readEvidencePropertyState(
  value: unknown,
): PdfNativeTextPropertyState | null {
  const record = unknownRecord(value);
  const font = readEvidenceFontOption(record?.font);
  const color = readEvidenceDeviceColor(record?.color);
  if (
    !record ||
    !font ||
    !color ||
    (record.alignment !== "left" &&
      record.alignment !== "center" &&
      record.alignment !== "right")
  ) {
    return null;
  }
  const values = [
    record.fontSize,
    record.textX,
    record.baseline,
    record.characterSpacing,
    record.wordSpacing,
    record.lineSpacing,
    record.leftIndent,
    record.rightIndent,
    record.firstLineIndent,
    record.horizontalScale,
    // A record written before superscript support carries no rise at all.
    record.textRise ?? 0,
  ].map(readFiniteNumber);
  const linear = Array.isArray(record.textMatrixLinear)
    ? record.textMatrixLinear.map(readFiniteNumber)
    : null;
  if (
    values.some((entry) => entry === null) ||
    values[0]! <= 0 ||
    // Zero is the default leading a run never set; a property edit keeps it.
    // `TD` sets the leading to -ty as a side effect, so a producer that steps
    // down the page leaves a negative one in force, and the edit records that
    // leading as found (mirror of the server's record schema).
    Math.abs(values[5]!) > MAX_PROVEN_LINE_SPACING ||
    values[6]! < 0 ||
    values[7]! < 0 ||
    values[9]! <= 0 ||
    Math.abs(values[10]!) > MAX_PROVEN_TEXT_RISE
  ) {
    return null;
  }
  // The named colour space the source set the run's colour through, the
  // components it set there, and a fill-and-stroke `Tr`: a history side
  // writes its wrapper back through them, as the edit that wrote it did
  // (mirror of the server's record schema). A record read without them would
  // rebuild the wrapper with a device colour and never find the one on the
  // page (an undo of a move under `/Cs8 cs 1 scn`).
  const colorSpaceResource =
    typeof record.colorSpaceResource === "string" &&
    RESOURCE_NAME.test(record.colorSpaceResource)
      ? record.colorSpaceResource
      : undefined;
  if (record.colorSpaceResource !== undefined && record.colorSpaceResource !== null && colorSpaceResource === undefined) {
    return null;
  }
  const colorSpaceComponents = Array.isArray(record.colorSpaceComponents)
    ? record.colorSpaceComponents.map(readFiniteNumber)
    : undefined;
  if (
    (record.colorSpaceComponents !== undefined && record.colorSpaceComponents !== null &&
      (!colorSpaceComponents || colorSpaceComponents.length < 1 || colorSpaceComponents.length > 32 ||
        colorSpaceComponents.some((entry) => entry === null || entry < 0 || entry > 1))) ||
    (record.renderingMode !== undefined && record.renderingMode !== null &&
      !(record.renderingMode === 0 || record.renderingMode === 1 || record.renderingMode === 2))
  ) {
    return null;
  }
  // The stroke width in force at the run (`w`), which a synthetic bold sets
  // and the writer puts back: carried as the record wrote it (0 for the
  // thinnest line a page may set before its text).
  const lineWidth = readFiniteNumber(record.lineWidth);
  if (record.lineWidth !== undefined && record.lineWidth !== null && (lineWidth === null || lineWidth < 0 || lineWidth > 1_000)) {
    return null;
  }
  // The stroke colour in force at the run, which a wrapper that strokes the
  // run in its own colour puts back after it, set the way the source set it
  // (mirror of the server's record schema). A history side read without it
  // would rebuild the wrapper without that restore and never find it.
  const strokeColor = record.strokeColor === undefined || record.strokeColor === null
    ? undefined
    : readEvidenceDeviceColor(record.strokeColor);
  const strokeColorSpaceResource =
    typeof record.strokeColorSpaceResource === "string" &&
    RESOURCE_NAME.test(record.strokeColorSpaceResource)
      ? record.strokeColorSpaceResource
      : undefined;
  const strokeColorSpaceComponents = Array.isArray(record.strokeColorSpaceComponents)
    ? record.strokeColorSpaceComponents.map(readFiniteNumber)
    : undefined;
  if (
    strokeColor === null ||
    (record.strokeColorSpaceResource !== undefined && record.strokeColorSpaceResource !== null &&
      strokeColorSpaceResource === undefined) ||
    (record.strokeColorSpaceComponents !== undefined && record.strokeColorSpaceComponents !== null &&
      (!strokeColorSpaceComponents || strokeColorSpaceComponents.length < 1 || strokeColorSpaceComponents.length > 32 ||
        strokeColorSpaceComponents.some((entry) => entry === null || entry < 0 || entry > 1))) ||
    ((strokeColorSpaceResource !== undefined || strokeColorSpaceComponents !== undefined) && strokeColor === undefined)
  ) {
    return null;
  }
  return Object.freeze({
    font,
    fontSize: values[0]!,
    color,
    ...(colorSpaceResource !== undefined ? { colorSpaceResource } : {}),
    ...(colorSpaceComponents !== undefined
      ? { colorSpaceComponents: Object.freeze(colorSpaceComponents as number[]) }
      : {}),
    ...(typeof record.renderingMode === "number" && record.renderingMode !== 0
      ? { renderingMode: record.renderingMode }
      : {}),
    ...(lineWidth !== null ? { lineWidth } : {}),
    ...(strokeColor !== undefined ? { strokeColor } : {}),
    ...(strokeColorSpaceResource !== undefined ? { strokeColorSpaceResource } : {}),
    ...(strokeColorSpaceComponents !== undefined
      ? { strokeColorSpaceComponents: Object.freeze(strokeColorSpaceComponents as number[]) }
      : {}),
    textX: values[1]!,
    baseline: values[2]!,
    alignment: record.alignment,
    characterSpacing: values[3]!,
    wordSpacing: values[4]!,
    lineSpacing: values[5]!,
    leftIndent: values[6]!,
    rightIndent: values[7]!,
    firstLineIndent: values[8]!,
    horizontalScale: values[9]!,
    textRise: values[10]!,
    textMatrixLinear: Object.freeze(
      linear && linear.length === 4 && !linear.some((value) => value === null)
        ? (linear as number[])
        : [1, 0, 0, 1],
    ) as readonly [number, number, number, number],
  });
}

export function readEvidencePropertyMutation(
  value: unknown,
): PdfNativeTextPropertyMutationEvidence | null {
  const record = unknownRecord(value);
  const original = readEvidencePropertyState(record?.original);
  const replacement = readEvidencePropertyState(record?.replacement);
  const frame = unknownRecord(record?.frame);
  const left = readFiniteNumber(frame?.left);
  const right = readFiniteNumber(frame?.right);
  const originalOperatorBytes = readCodeByteArray(
    record?.originalOperatorBytes,
  );
  const replacementOperatorBytes = readCodeByteArray(
    record?.replacementOperatorBytes,
  );
  const originalSelection = readEvidenceSelection(record?.originalSelection);
  const replacementSelection = readEvidenceSelection(
    record?.replacementSelection,
  );
  const offsetRecord = record?.offset === undefined ? null : unknownRecord(record.offset);
  const offsetX = offsetRecord ? readFiniteNumber(offsetRecord.x) : null;
  const offsetY = offsetRecord ? readFiniteNumber(offsetRecord.y) : null;
  // A run split by the formats of part of it: its parts end to end.
  const segments = record?.segments === undefined
    ? undefined
    : Array.isArray(record.segments)
      ? record.segments.map((value) => {
          const segment = unknownRecord(value);
          const codeStart = readFiniteNumber(segment?.codeStart);
          const codeEnd = readFiniteNumber(segment?.codeEnd);
          const state = readEvidencePropertyState(segment?.state);
          return codeStart === null || codeEnd === null || !state ||
            !Number.isSafeInteger(codeStart) || !Number.isSafeInteger(codeEnd) || codeEnd <= codeStart
            ? null
            : Object.freeze({ codeStart, codeEnd, state });
        })
      : null;
  if (
    !record ||
    !original ||
    !replacement ||
    left === null ||
    right === null ||
    right <= left ||
    !originalOperatorBytes ||
    !replacementOperatorBytes ||
    !originalSelection ||
    !replacementSelection ||
    (record.offset !== undefined && (offsetX === null || offsetY === null)) ||
    segments === null ||
    (segments !== undefined && (segments.length < 2 || segments.some((segment) => segment === null)))
  ) {
    return null;
  }
  return Object.freeze({
    original,
    replacement,
    frame: Object.freeze({ left, right }),
    originalOperatorBytes,
    replacementOperatorBytes,
    originalSelection,
    replacementSelection,
    ...(offsetX !== null && offsetY !== null ? { offset: Object.freeze({ x: offsetX, y: offsetY }) } : {}),
    ...(segments ? { segments: Object.freeze(segments as NonNullable<(typeof segments)[number]>[]) } : {}),
  });
}

/**
 * Parse persisted v4 fidelity without trusting a TypeScript cast. This parser
 * is deliberately structural; current-byte semantics are independently
 * re-established by the history planner before any mutation is authorized.
 */
export function readMraiNativeListTextFidelityEvidence(
  value: unknown,
): PdfNativeListTextFidelityEvidence | null {
  const record = unknownRecord(value);
  if (
    !record ||
    (record.schemaVersion !== MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION &&
      record.schemaVersion !== MRAI_NATIVE_MULTI_RUN_TEXT_SCHEMA_VERSION &&
      record.schemaVersion !== MRAI_NATIVE_TEXT_PROPERTIES_SCHEMA_VERSION &&
      record.schemaVersion !== MRAI_NATIVE_PARAGRAPH_REFLOW_SCHEMA_VERSION &&
      record.schemaVersion !==
        MRAI_NATIVE_MULTICOLUMN_PARAGRAPH_REFLOW_SCHEMA_VERSION &&
      record.schemaVersion !==
        MRAI_NATIVE_ORTHOGONAL_PARAGRAPH_REFLOW_SCHEMA_VERSION &&
      record.schemaVersion !==
        MRAI_NATIVE_AFFINE_PARAGRAPH_REFLOW_SCHEMA_VERSION) ||
    (record.layoutKind !== "list-body" &&
      record.layoutKind !== "paragraph-line") ||
    (record.wrapping !== "same-line-only" &&
      record.wrapping !== "fixed-frame-multiline" &&
      record.wrapping !== "paragraph-reflow")
  ) {
    return null;
  }
  const originalText = record.consolidatedRuns != null && record.originalText === ""
    ? "" : readBoundedEvidenceText(record.originalText);
  const originalDisplayText = record.consolidatedRuns != null && record.originalDisplayText === ""
    ? "" : readBoundedEvidenceText(record.originalDisplayText);
  // A consolidated leaf may empty its line (trailing contraction), and the
  // exact restore of such a line mirrors that empty side as its current one.
  const consolidatedRecord = unknownRecord(record.consolidatedRuns);
  // A run or line the edit deleted whole is empty on its replacement side in
  // every field too; the flag says so, and the shape is checked below.
  if (record.emptiedShow !== undefined && record.emptiedShow !== true) return null;
  const emptiedShow = record.emptiedShow === true;
  if (record.emptiedShowPen != null && (!emptiedShow || typeof record.emptiedShowPen !== "string" ||
    !/^-?\d{1,9}(?:\.\d{1,6})?$/u.test(record.emptiedShowPen))) return null;
  const emptiedShowPen = typeof record.emptiedShowPen === "string" ? record.emptiedShowPen : undefined;
  const emptyConsolidatedReplacement =
    (consolidatedRecord !== null || emptiedShow) &&
    record.replacementText === "" &&
    record.replacementDisplayText === "" &&
    Array.isArray(record.replacementCodeBytes) &&
    record.replacementCodeBytes.length === 0 &&
    record.replacementAdvance === 0;
  const emptyConsolidatedOriginal =
    consolidatedRecord !== null &&
    consolidatedRecord.history != null &&
    record.originalText === "" &&
    record.originalDisplayText === "" &&
    Array.isArray(record.originalCodeBytes) &&
    record.originalCodeBytes.length === 0 &&
    record.originalAdvance === 0;
  const replacementText = emptyConsolidatedReplacement
    ? ""
    : readBoundedEvidenceText(record.replacementText);
  const replacementDisplayText = emptyConsolidatedReplacement
    ? ""
    : readBoundedEvidenceText(record.replacementDisplayText);
  const originalCodeBytes = emptyConsolidatedOriginal
    ? Object.freeze([] as number[])
    : readCodeByteArray(record.originalCodeBytes);
  const replacementCodeBytes = emptyConsolidatedReplacement
    ? Object.freeze([] as number[])
    : readCodeByteArray(record.replacementCodeBytes);
  const fontSize = readFiniteNumber(record.fontSize);
  const textX = readFiniteNumber(record.textX);
  const baseline = readFiniteNumber(record.baseline);
  const lineSpacing = readFiniteNumber(record.lineSpacing);
  const originalAdvance = readFiniteNumber(record.originalAdvance);
  const replacementAdvance = readFiniteNumber(record.replacementAdvance);
  const textOperatorSpan = readEvidenceSpan(record.textOperatorSpan);
  const resources = readEvidenceResources(record.resources);
  const fontEncoding = readEvidenceFontEncoding(record.fontEncoding);
  const layoutWitnesses = readEvidenceLayoutWitnesses(
    record.layoutWitnesses,
  );
  if (
    originalText === null ||
    originalDisplayText === null ||
    ((originalText === "" || originalDisplayText === "") && !emptyConsolidatedOriginal) ||
    replacementText === null ||
    replacementDisplayText === null ||
    originalCodeBytes === null ||
    replacementCodeBytes === null ||
    typeof record.resourceName !== "string" ||
    !RESOURCE_NAME.test(record.resourceName) ||
    fontSize === null ||
    fontSize <= 0 ||
    textX === null ||
    baseline === null ||
    lineSpacing === null ||
    originalAdvance === null ||
    originalAdvance < 0 ||
    (originalAdvance === 0 && !emptyConsolidatedOriginal) ||
    replacementAdvance === null ||
    replacementAdvance < 0 ||
    (replacementAdvance === 0 && !emptyConsolidatedReplacement) ||
    !textOperatorSpan ||
    !resources ||
    !fontEncoding ||
    !layoutWitnesses ||
    layoutWitnesses.some(
      (witness) =>
        witness.resourceName === record.resourceName ||
        !refsEqual(witness.resources.page, resources.page),
    )
  ) {
    return null;
  }
  const selection =
    record.selection === null ? null : readEvidenceSelection(record.selection);
  if (record.selection !== null && !selection) return null;
  const rawSameLineLayoutState =
    record.sameLineLayoutState === undefined
      ? null
      : unknownRecord(record.sameLineLayoutState);
  const sameLineLayoutValues = rawSameLineLayoutState
    ? [
        rawSameLineLayoutState.fontSize,
        rawSameLineLayoutState.characterSpacing,
        rawSameLineLayoutState.wordSpacing,
        rawSameLineLayoutState.horizontalScale,
      ].map(readFiniteNumber)
    : null;
  if (
    record.sameLineLayoutState !== undefined &&
    (!rawSameLineLayoutState ||
      !sameLineLayoutValues ||
      sameLineLayoutValues.some((entry) => entry === null) ||
      sameLineLayoutValues[0]! <= 0 ||
      sameLineLayoutValues[3]! <= 0)
  ) {
    return null;
  }
  const propertyMutation = readEvidencePropertyMutation(
    record.propertyMutation,
  );
  const paragraphReflow = readEvidenceParagraphReflow(record.paragraphReflow);
  const paragraphSchema =
    record.schemaVersion === MRAI_NATIVE_PARAGRAPH_REFLOW_SCHEMA_VERSION ||
    record.schemaVersion ===
      MRAI_NATIVE_MULTICOLUMN_PARAGRAPH_REFLOW_SCHEMA_VERSION ||
    record.schemaVersion ===
      MRAI_NATIVE_ORTHOGONAL_PARAGRAPH_REFLOW_SCHEMA_VERSION ||
    record.schemaVersion === MRAI_NATIVE_AFFINE_PARAGRAPH_REFLOW_SCHEMA_VERSION;
  if (
    (record.schemaVersion === MRAI_NATIVE_TEXT_PROPERTIES_SCHEMA_VERSION) !==
      (propertyMutation !== null) ||
    (record.schemaVersion !== MRAI_NATIVE_TEXT_PROPERTIES_SCHEMA_VERSION &&
      record.propertyMutation !== undefined) ||
    paragraphSchema !== (paragraphReflow !== null) ||
    (!paragraphSchema && record.paragraphReflow !== undefined) ||
    (record.schemaVersion === MRAI_NATIVE_PARAGRAPH_REFLOW_SCHEMA_VERSION &&
      (paragraphReflow?.columnTopology !== undefined ||
        paragraphReflow?.orthogonalTransform !== undefined ||
        paragraphReflow?.affineTransform !== undefined)) ||
    (record.schemaVersion ===
      MRAI_NATIVE_MULTICOLUMN_PARAGRAPH_REFLOW_SCHEMA_VERSION &&
      (paragraphReflow?.columnTopology === undefined ||
        paragraphReflow?.orthogonalTransform !== undefined ||
        paragraphReflow?.affineTransform !== undefined)) ||
    (record.schemaVersion ===
      MRAI_NATIVE_ORTHOGONAL_PARAGRAPH_REFLOW_SCHEMA_VERSION &&
      (paragraphReflow?.orthogonalTransform === undefined ||
        paragraphReflow?.affineTransform !== undefined ||
        paragraphReflow?.columnTopology !== undefined)) ||
    (record.schemaVersion ===
      MRAI_NATIVE_AFFINE_PARAGRAPH_REFLOW_SCHEMA_VERSION &&
      (paragraphReflow?.affineTransform === undefined ||
        paragraphReflow?.orthogonalTransform !== undefined ||
        paragraphReflow?.columnTopology !== undefined))
  ) {
    return null;
  }
  let listAnchor: PdfNativeListTextFidelityEvidence["listAnchor"] = null;
  if (record.listAnchor !== null) {
    const anchor = unknownRecord(record.listAnchor);
    const markerOperatorSpan = readEvidenceSpan(anchor?.markerOperatorSpan);
    const markerX = readFiniteNumber(anchor?.markerX);
    const hangingIndent = readFiniteNumber(anchor?.hangingIndent);
    if (
      !anchor ||
      !readBoundedEvidenceText(anchor.markerText) ||
      markerX === null ||
      hangingIndent === null ||
      !markerOperatorSpan
    ) {
      return null;
    }
    listAnchor = Object.freeze({
      markerText: String(anchor.markerText),
      markerX,
      hangingIndent,
      markerOperatorSpan,
    });
  }
  let linePlan: PdfNativeFixedFrameLinePlan | null = null;
  if (record.linePlan !== null) {
    const plan = unknownRecord(record.linePlan);
    const frameLeft = readFiniteNumber(plan?.frameLeft);
    const frameRight = readFiniteNumber(plan?.frameRight);
    const schemas = new Set<string>([
      MRAI_NATIVE_FIXED_FRAME_LINE_PLAN_SCHEMA_VERSION,
      MRAI_NATIVE_ADJACENT_TEXT_OBJECT_LINE_PLAN_SCHEMA_VERSION,
      MRAI_NATIVE_ADAPTIVE_FIXED_FRAME_LINE_PLAN_SCHEMA_VERSION,
      MRAI_NATIVE_ADAPTIVE_ADJACENT_TEXT_OBJECT_LINE_PLAN_SCHEMA_VERSION,
      MRAI_NATIVE_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION,
      MRAI_NATIVE_ADAPTIVE_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION,
      MRAI_NATIVE_PROJECTED_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION,
    ]);
    if (
      !plan ||
      typeof plan.schemaVersion !== "string" ||
      !schemas.has(plan.schemaVersion) ||
      frameLeft === null ||
      frameRight === null ||
      frameRight <= frameLeft ||
      !Array.isArray(plan.lines) ||
      plan.lines.length < 2 ||
      plan.lines.length > PDF_NATIVE_MAX_FIXED_FRAME_ITEMS
    ) {
      return null;
    }
    const lines: PdfNativeFixedFrameLinePlan["lines"][number][] = [];
    for (const value of plan.lines) {
      const line = unknownRecord(value);
      const operatorSpan = readEvidenceSpan(line?.operatorSpan);
      const replacementSpan = readEvidenceSpan(line?.replacementSpan);
      const originalLineText = readBoundedEvidenceText(line?.originalText);
      const replacementLineText = (plan.schemaVersion === MRAI_NATIVE_ADAPTIVE_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION || plan.schemaVersion === MRAI_NATIVE_PROJECTED_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION) &&
        line?.replacementText === "" ? "" : readBoundedEvidenceText(line?.replacementText);
      if (
        !line ||
        !Number.isSafeInteger(line.itemIndex) ||
        Number(line.itemIndex) < 0 ||
        !originalLineText ||
        replacementLineText === null ||
        !["Tj", "Tj-literal", "'", "'-literal"].includes(
          String(line.operatorKind),
        ) ||
        !operatorSpan ||
        !replacementSpan
      ) {
        return null;
      }
      const numbers = [
        line.textX,
        line.baseline,
        line.originalAdvance,
        line.replacementAdvance,
      ].map(readFiniteNumber);
      const verticalCapacity =
        line.verticalCapacity === undefined
          ? null
          : readFiniteNumber(line.verticalCapacity);
      if (numbers.some((entry) => entry === null)) return null;
      const rawLayoutState =
        line.layoutState === undefined
          ? null
          : unknownRecord(line.layoutState);
      const layoutValues = rawLayoutState
        ? [
            rawLayoutState.fontSize,
            rawLayoutState.characterSpacing,
            rawLayoutState.wordSpacing,
            rawLayoutState.horizontalScale,
          ].map(readFiniteNumber)
        : null;
      if (
        line.layoutState !== undefined &&
        (!rawLayoutState ||
          !layoutValues ||
          layoutValues.some((entry) => entry === null) ||
          layoutValues[0]! <= 0 ||
          layoutValues[3]! <= 0)
      ) {
        return null;
      }
      if (
        line.verticalCapacity !== undefined &&
        (verticalCapacity === null || verticalCapacity <= 0)
      ) {
        return null;
      }
      let multiRunFields:
        | Readonly<{
            lineIndex: number;
            selectionItemIndex?: number;
            resourceName: string;
            fontEncoding: PdfNativeFontEncoding;
            resources: PdfNativeListTextFidelityEvidence["resources"];
          }>
        | undefined;
      if (
        isMraiNativeMultiRunLinePlan(plan.schemaVersion)
      ) {
        const lineFontEncoding = readEvidenceFontEncoding(line.fontEncoding);
        const lineResources = readEvidenceResources(line.resources);
        const selectionItemIndex =
          line.selectionItemIndex === undefined
            ? undefined
            : Number(line.selectionItemIndex);
        if (
          typeof line.resourceName !== "string" ||
          !RESOURCE_NAME.test(line.resourceName) ||
          !Number.isSafeInteger(line.lineIndex) ||
          Number(line.lineIndex) < 0 ||
          (selectionItemIndex !== undefined &&
            (!Number.isSafeInteger(selectionItemIndex) ||
              selectionItemIndex < 0 ||
              selectionItemIndex >= PDF_NATIVE_MAX_FIXED_FRAME_ITEMS)) ||
          !lineFontEncoding ||
          !lineResources
        ) {
          return null;
        }
        multiRunFields = Object.freeze({
          lineIndex: Number(line.lineIndex),
          ...(selectionItemIndex === undefined ? {} : { selectionItemIndex }),
          resourceName: line.resourceName,
          fontEncoding: lineFontEncoding,
          resources: lineResources,
        });
      } else if (line.selectionItemIndex !== undefined) {
        return null;
      }
      lines.push(
        Object.freeze({
          itemIndex: Number(line.itemIndex),
          originalText: originalLineText,
          replacementText: replacementLineText,
          textX: numbers[0]!,
          baseline: numbers[1]!,
          originalAdvance: numbers[2]!,
          replacementAdvance: numbers[3]!,
          ...(verticalCapacity === null ? {} : { verticalCapacity }),
          operatorKind: line.operatorKind as
            | "Tj"
            | "Tj-literal"
            | "'"
            | "'-literal",
          operatorSpan,
          replacementSpan,
          ...(layoutValues
            ? {
                layoutState: Object.freeze({
                  fontSize: layoutValues[0]!,
                  characterSpacing: layoutValues[1]!,
                  wordSpacing: layoutValues[2]!,
                  horizontalScale: layoutValues[3]!,
                }),
              }
            : {}),
          ...(multiRunFields ?? {}),
        }),
      );
    }
    if ((plan.schemaVersion === MRAI_NATIVE_ADAPTIVE_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION || plan.schemaVersion === MRAI_NATIVE_PROJECTED_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION) &&
      ((plan.schemaVersion === MRAI_NATIVE_ADAPTIVE_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION && !lines.some(line => line.replacementText === "")) || lines.some(line =>
        line.originalAdvance <= 0 || (line.replacementText === ""
          ? line.replacementAdvance !== 0 || !["Tj", "Tj-literal"].includes(line.operatorKind)
          : line.replacementAdvance <= 0)))) return null;
    linePlan = Object.freeze({
      schemaVersion:
        plan.schemaVersion as PdfNativeFixedFrameLinePlan["schemaVersion"],
      frameLeft,
      frameRight,
      lines: Object.freeze(lines),
    });
  }
  const singleItemSourceRunEvidence =
    record.schemaVersion === MRAI_NATIVE_MULTI_RUN_TEXT_SCHEMA_VERSION &&
    linePlan !== null &&
    isMraiNativeMultiRunLinePlan(linePlan?.schemaVersion) &&
    (linePlan.lines.length === 2 || linePlan.schemaVersion === MRAI_NATIVE_PROJECTED_MULTI_RUN_LINE_PLAN_SCHEMA_VERSION) &&
    linePlan.lines.every(
      (line) => line.selectionItemIndex === 0 && line.lineIndex === 0,
    ) &&
    selection !== null &&
    selection.items === undefined;
  const linePlanSelectionItems =
    selection?.items ??
    (singleItemSourceRunEvidence && selection
      ? Object.freeze([selection.item])
      : null);
  const vertical = isMraiNativeVerticalFontEncodingKind(fontEncoding.kind);
  const verticalFixedFrame =
    vertical &&
    record.wrapping === "fixed-frame-multiline" &&
    linePlan !== null &&
    selection?.items !== undefined &&
    selection.items.length >= 2 &&
    selection.items.length === linePlan.lines.length &&
    listAnchor === null &&
    propertyMutation === null &&
    paragraphReflow === null &&
    linePlan.lines.every(
      (line, index) =>
        line.itemIndex === selection.items![index].itemIndex &&
        nearlyEqual(line.baseline, linePlan.lines[0].baseline, 0.001) &&
        (index === 0 || linePlan.lines[index - 1].textX > line.textX) &&
        line.verticalCapacity !== undefined &&
        line.originalAdvance <= line.verticalCapacity + 0.001 &&
        line.replacementAdvance <= line.verticalCapacity + 0.001,
    ) &&
    linePlan.lines.slice(2).every((line, index) =>
      nearlyEqual(
        linePlan.lines[index].textX - linePlan.lines[index + 1].textX,
        linePlan.lines[index + 1].textX - line.textX,
        0.001,
      ),
    ) &&
    nearlyEqual(
      lineSpacing,
      linePlan.lines[0].textX - linePlan.lines[1].textX,
      0.001,
    ) &&
    nearlyEqual(linePlan.frameLeft, linePlan.lines.at(-1)!.textX, 0.001) &&
    nearlyEqual(
      linePlan.frameRight,
      linePlan.lines[0].textX + fontSize,
      0.001,
    ) &&
    linePlan.lines.every((line) => line.verticalCapacity !== undefined);
  if (
    (vertical &&
      !verticalFixedFrame &&
      (record.wrapping !== "same-line-only" ||
        linePlan !== null ||
        listAnchor !== null ||
        propertyMutation !== null ||
        paragraphReflow !== null ||
        selection === null ||
        (selection.items?.length ?? 1) !== 1)) ||
    (!verticalFixedFrame &&
      linePlan?.lines.some((line) => line.verticalCapacity !== undefined)) ||
    (record.wrapping === "same-line-only" && linePlan !== null) ||
    (sameLineLayoutValues !== null &&
      (record.schemaVersion !== MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION ||
        record.wrapping !== "same-line-only" ||
        linePlan !== null ||
        !nearlyEqual(sameLineLayoutValues[0]!, fontSize, 0.001))) ||
    (record.wrapping === "fixed-frame-multiline" && linePlan === null && record.positionedLines == null &&
      record.paragraphRewrite == null) ||
    (record.wrapping === "paragraph-reflow" && linePlan !== null) ||
    (record.schemaVersion === MRAI_NATIVE_MULTI_RUN_TEXT_SCHEMA_VERSION) !==
      (isMraiNativeMultiRunLinePlan(linePlan?.schemaVersion)) ||
    (record.schemaVersion === MRAI_NATIVE_MULTI_RUN_TEXT_SCHEMA_VERSION &&
      (!linePlan ||
        !linePlanSelectionItems ||
        (() => {
          const itemIndexes = linePlanSelectionItemIndexes(
            linePlan,
            linePlanSelectionItems.length,
          );
          return (
            !itemIndexes ||
            linePlan.lines.some(
              (line, index) =>
                line.itemIndex !==
                linePlanSelectionItems[itemIndexes[index]].itemIndex,
            )
          );
        })())) ||
    (record.schemaVersion === MRAI_NATIVE_TEXT_PROPERTIES_SCHEMA_VERSION &&
      (record.wrapping !== "same-line-only" ||
        linePlan !== null ||
        listAnchor !== null ||
        !selection)) ||
    (paragraphSchema &&
      (record.wrapping !== "paragraph-reflow" ||
        listAnchor !== null ||
        !selection))
  ) {
    return null;
  }
  let positionedRunGroup: PdfNativeListTextFidelityEvidence["positionedRunGroup"];
  if (record.positionedRunGroup !== undefined && record.positionedRunGroup !== null) {
    const group = unknownRecord(record.positionedRunGroup);
    if (!group || group.schemaVersion !== "mrai.pdf-native-positioned-runs.v1" || !Array.isArray(group.runs) ||
      group.runs.length < 2 || group.runs.length > 64 || record.schemaVersion !== MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION ||
      record.wrapping !== "same-line-only" || linePlan || listAnchor || paragraphReflow || propertyMutation || !selection || selection.items) return null;
    const runs: Array<{ evidence: PdfNativeListTextFidelityEvidence; replacement: readonly number[] }> = [];
    let total = 0;
    for (const value of group.runs) {
      const run = unknownRecord(value), rawEvidence = unknownRecord(run?.evidence);
      // A run that widens before the runs the group moves carries the room
      // past the group's end it grew into (positionedFrameCapacity); the
      // run's own reading bounds it.
      if (!run || !rawEvidence || rawEvidence.positionedRunGroup != null || rawEvidence.positionedLines != null ||
        rawEvidence.consolidatedRuns != null) return null;
      const evidence = readMraiNativeListTextFidelityEvidence(rawEvidence), replacement = readCodeByteArray(run.replacement);
      // The group carries its first run's font; a later run may be set in
      // another font of the group's size (a label whose faces switch
      // mid-word, TeX's braces and digits), and a run of nothing but white
      // space at another size (mirror of replay's group authority).
      if (!evidence || !replacement || !evidence.selection || evidence.selection.items ||
        evidence.schemaVersion !== record.schemaVersion || evidence.wrapping !== "same-line-only" || evidence.linePlan ||
        evidence.listAnchor || evidence.propertyMutation || evidence.paragraphReflow ||
        ((evidence.originalText ?? "").trim() !== "" && evidence.fontSize !== fontSize) ||
        (runs.length === 0 && (evidence.resourceName !== record.resourceName ||
          JSON.stringify(evidence.resources) !== JSON.stringify(resources) || JSON.stringify(evidence.fontEncoding) !== JSON.stringify(fontEncoding))) ||
        evidence.selection.pageNumber !== selection.pageNumber || evidence.selection.item.itemIndex !== selection.item.itemIndex) return null;
      total += replacement.length;
      if (total > 65536) return null;
      runs.push(Object.freeze({ evidence, replacement }));
    }
    if (runs.map(run => run.evidence.originalText).join("") !== originalText ||
      runs.map(run => run.evidence.replacementText).join("") !== replacementText ||
      JSON.stringify(runs.flatMap(run => [...run.evidence.originalCodeBytes])) !== JSON.stringify(originalCodeBytes) ||
      JSON.stringify(runs.flatMap(run => [...run.evidence.replacementCodeBytes])) !== JSON.stringify(replacementCodeBytes)) return null;
    const history = group.history == null ? null : unknownRecord(group.history);
    if (group.history != null && (!history || typeof history.originTransactionId !== "string" || !history.originTransactionId || history.originTransactionId.length > 200 ||
      !["undo", "redo"].includes(String(history.direction)))) return null;
    if (record.positionedRestore === true || (!history && runs.some(run => run.evidence.positionedRestore))) return null;
    positionedRunGroup = Object.freeze({ schemaVersion: "mrai.pdf-native-positioned-runs.v1", runs: Object.freeze(runs),
      ...(history ? { history: Object.freeze({ originTransactionId: history.originTransactionId as string, direction: history.direction as "undo" | "redo" }) } : {}) });
  }
  let positionedLines: PdfNativeListTextFidelityEvidence["positionedLines"];
  if (record.positionedLines != null) {
    const group = unknownRecord(record.positionedLines);
    if (!group || group.schemaVersion !== "mrai.pdf-native-positioned-lines.v1" ||
      !Array.isArray(group.lines) || group.lines.length < 2 || group.lines.length > 64 ||
      !selection?.items || selection.items.length !== group.lines.length || positionedRunGroup || linePlan ||
      propertyMutation || paragraphReflow || listAnchor || sameLineLayoutValues || record.positionedRestore === true ||
      record.schemaVersion !== MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION || record.wrapping !== "fixed-frame-multiline") return null;
    const history = group.history == null ? null : unknownRecord(group.history);
    if (group.history != null && (!history || typeof history.originTransactionId !== "string" || !history.originTransactionId || history.originTransactionId.length > 200 ||
      !["undo", "redo"].includes(String(history.direction)))) return null;
    const lines: Array<{ evidence: PdfNativeListTextFidelityEvidence; replacement: readonly number[] }> = [];
    let count = 0, bytes = 0;
    for (const [index, value] of group.lines.entries()) {
      const line = unknownRecord(value), raw = unknownRecord(line?.evidence);
      if (!line || !raw || raw.positionedLines != null) return null;
      const evidence = readMraiNativeListTextFidelityEvidence(raw), replacement = readCodeByteArray(line.replacement);
      // The paragraph is described by its first line; a line of a paragraph
      // set in more than one font is proven in the font that draws it, on the
      // paragraph's page (mirror of the server's record contract). Read as the
      // first line's font, a paragraph with an italic line was saved but its
      // history could not be read back, and the page stayed 未保存.
      if (!evidence || !replacement || !evidence.selection || evidence.selection.items ||
        evidence.schemaVersion !== record.schemaVersion || evidence.wrapping !== "same-line-only" ||
        evidence.linePlan || evidence.paragraphReflow || evidence.propertyMutation || evidence.listAnchor ||
        (evidence.positionedRestore && !history) ||
        (evidence.positionedRunGroup?.history && JSON.stringify(evidence.positionedRunGroup.history) !== JSON.stringify(history)) ||
        (evidence.consolidatedRuns?.history && JSON.stringify(evidence.consolidatedRuns.history) !== JSON.stringify(history)) ||
        JSON.stringify(evidence.resources.page) !== JSON.stringify(resources.page) ||
        (index === 0 && (evidence.resourceName !== record.resourceName || evidence.fontSize !== fontSize ||
          JSON.stringify(evidence.resources) !== JSON.stringify(resources) ||
          JSON.stringify(evidence.fontEncoding) !== JSON.stringify(fontEncoding))) ||
        evidence.selection.sourceSha256 !== selection.sourceSha256 || evidence.selection.pageNumber !== selection.pageNumber ||
        evidence.selection.selectionRevision !== selection.selectionRevision) return null;
      count += evidence.positionedRunGroup?.runs.length ?? 1;
      bytes += replacement.length;
      if (count > 64 || bytes > 65536) return null;
      lines.push(Object.freeze({ evidence, replacement }));
    }
    // A line proves its item, or the item less the gap the reader read as its
    // leading space; that space stays the separator before the line.
    const leadingGaps = positionedLineLeadingGaps(selection.items, lines.map(line => line.evidence.selection?.item));
    if (!leadingGaps) return null;
    // Each line as the reader reads it: its own text, or -- as the server
    // reads it -- one space for each run of white space and none for the
    // white space its show opens with when its item does not open with the
    // reader's gap space (a drop-cap line drawn as " Body", read as "Body").
    const replacementSummary = (projected: boolean) => summarizePdfNativeSelectionItems((selection.items ?? []).map((item, index) => {
      const own = lines[index].evidence.replacementDisplayText.trimEnd();
      const read = projected ? own.replace(/\s+/gu, " ").replace(/ $/u, "") : own;
      const lineItem = lines[index].evidence.selection?.item.str ?? "";
      return { ...item, str: (leadingGaps[index] ? " " : "") + (projected && !/^\s/u.test(lineItem) ? read.trimStart() : read) };
    }));
    if (lines.map(line => line.evidence.originalText).join("") !== originalText ||
      lines.map(line => line.evidence.replacementText).join("") !== replacementText ||
      summarizePdfNativeSelectionItems(selection.items) !== originalDisplayText ||
      (replacementSummary(false) !== replacementDisplayText && replacementSummary(true) !== replacementDisplayText) ||
      JSON.stringify(lines.flatMap(line => [...line.evidence.originalCodeBytes])) !== JSON.stringify(originalCodeBytes) ||
      JSON.stringify(lines.flatMap(line => [...line.evidence.replacementCodeBytes])) !== JSON.stringify(replacementCodeBytes)) return null;
    const rawFrame = group.frame == null ? null : unknownRecord(group.frame);
    const frameLeft = rawFrame ? readFiniteNumber(rawFrame.left) : null;
    const frameRight = rawFrame ? readFiniteNumber(rawFrame.right) : null;
    if (group.frame != null && (!rawFrame || frameLeft === null || frameRight === null || frameLeft >= frameRight)) return null;
    if (!rawFrame && lines.some(line => line.evidence.positionedFrameCapacity !== undefined)) return null;
    positionedLines = Object.freeze({ schemaVersion: "mrai.pdf-native-positioned-lines.v1", lines: Object.freeze(lines),
      ...(rawFrame ? { frame: Object.freeze({ left: frameLeft!, right: frameRight! }) } : {}),
      ...(history ? { history: Object.freeze({ originTransactionId: history.originTransactionId as string, direction: history.direction as "undo" | "redo" }) } : {}) });
  }
  let consolidatedRuns: PdfNativeListTextFidelityEvidence["consolidatedRuns"];
  if (record.consolidatedRuns != null) {
    const group = unknownRecord(record.consolidatedRuns);
    if (!group || group.schemaVersion !== "mrai.pdf-native-consolidated-runs.v1" || !Array.isArray(group.runs) ||
      group.runs.length < 1 || group.runs.length > 64 || record.schemaVersion !== MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION ||
      record.wrapping !== "same-line-only" || record.layoutKind !== "paragraph-line" || linePlan || listAnchor ||
      paragraphReflow || propertyMutation || positionedRunGroup || positionedLines || record.positionedRestore === true ||
      !sameLineLayoutValues || lineSpacing !== 0 || !Number.isSafeInteger(group.envelopeEnd) || (group.envelopeEnd as number) < 0) return null;
    const runs: Array<Readonly<{ start: number; end: number }>> = [];
    for (const value of group.runs) {
      const run = unknownRecord(value);
      if (!run || !Number.isSafeInteger(run.start) || !Number.isSafeInteger(run.end) || (run.start as number) < 0 ||
        (run.end as number) <= (run.start as number) || (runs.length > 0 && (run.start as number) < runs[runs.length - 1].end)) return null;
      runs.push(Object.freeze({ start: run.start as number, end: run.end as number }));
    }
    const envelopeEnd = group.envelopeEnd as number;
    const originalBytes = readCodeByteArray(group.originalBytes);
    if (runs[0].start !== textOperatorSpan.start || runs[runs.length - 1].end > textOperatorSpan.end ||
      envelopeEnd < runs[0].end || envelopeEnd > textOperatorSpan.end || !originalBytes ||
      originalBytes.length !== textOperatorSpan.end - textOperatorSpan.start) return null;
    const consolidatedHistory = group.history == null ? null : unknownRecord(group.history);
    if (group.history != null && (!consolidatedHistory || typeof consolidatedHistory.originTransactionId !== "string" ||
      !consolidatedHistory.originTransactionId || consolidatedHistory.originTransactionId.length > 200 ||
      !["undo", "redo"].includes(String(consolidatedHistory.direction)))) return null;
    consolidatedRuns = Object.freeze({ schemaVersion: "mrai.pdf-native-consolidated-runs.v1", runs: Object.freeze(runs), envelopeEnd, originalBytes,
      ...(consolidatedHistory ? { history: Object.freeze({ originTransactionId: consolidatedHistory.originTransactionId as string,
        direction: consolidatedHistory.direction as "undo" | "redo" }) } : {}) });
  }
  let paragraphRewrite: PdfNativeParagraphRewriteEvidence | undefined;
  if (record.paragraphRewrite != null) {
    const read = linePlan || positionedLines || positionedRunGroup || consolidatedRuns || paragraphReflow || propertyMutation
      ? null : readMraiNativeParagraphRewriteEvidence(record.paragraphRewrite, readMraiNativeListTextFidelityEvidence);
    if (!read) return null;
    paragraphRewrite = read;
  }
  if (emptiedShow && (!emptyConsolidatedReplacement || record.schemaVersion !== MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION ||
    record.wrapping !== "same-line-only" || !originalText.trim() || linePlan || listAnchor || propertyMutation ||
    paragraphReflow || positionedRunGroup || positionedLines || consolidatedRuns || paragraphRewrite ||
    record.positionedRestore === true || record.readerChunk != null || record.lineShift != null ||
    record.sameLineCapacity != null || record.positionedFrameCapacity != null)) return null;
  let sameLineCapacity: number | undefined;
  if (record.sameLineCapacity != null) {
    const capacity = readFiniteNumber(record.sameLineCapacity);
    if (capacity === null || capacity <= 0) return null;
    sameLineCapacity = capacity;
  }
  let positionedFrameCapacity: number | undefined;
  if (record.positionedFrameCapacity != null) {
    const capacity = readFiniteNumber(record.positionedFrameCapacity);
    if (capacity === null || capacity <= 0 || capacity > 1_000_000 || record.wrapping !== "same-line-only" || linePlan ||
      positionedRunGroup || positionedLines) return null;
    positionedFrameCapacity = capacity;
  }
  return Object.freeze({
    schemaVersion: record.schemaVersion as
      | typeof MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION
      | typeof MRAI_NATIVE_MULTI_RUN_TEXT_SCHEMA_VERSION
      | typeof MRAI_NATIVE_TEXT_PROPERTIES_SCHEMA_VERSION
      | typeof MRAI_NATIVE_PARAGRAPH_REFLOW_SCHEMA_VERSION
      | typeof MRAI_NATIVE_MULTICOLUMN_PARAGRAPH_REFLOW_SCHEMA_VERSION
      | typeof MRAI_NATIVE_ORTHOGONAL_PARAGRAPH_REFLOW_SCHEMA_VERSION
      | typeof MRAI_NATIVE_AFFINE_PARAGRAPH_REFLOW_SCHEMA_VERSION,
    layoutKind: record.layoutKind,
    originalText,
    originalDisplayText,
    originalCodeBytes,
    replacementText,
    replacementDisplayText,
    replacementCodeBytes,
    resourceName: record.resourceName,
    fontSize,
    textX,
    baseline,
    lineSpacing,
    originalAdvance,
    replacementAdvance,
    wrapping: record.wrapping,
    linePlan,
    ...(sameLineLayoutValues
      ? {
          sameLineLayoutState: Object.freeze({
            fontSize: sameLineLayoutValues[0]!,
            characterSpacing: sameLineLayoutValues[1]!,
            wordSpacing: sameLineLayoutValues[2]!,
            horizontalScale: sameLineLayoutValues[3]!,
          }),
        }
      : {}),
    fontEncoding,
    selection,
    textOperatorSpan,
    listAnchor,
    resources,
    ...(layoutWitnesses.length ? { layoutWitnesses } : {}),
    ...(propertyMutation ? { propertyMutation } : {}),
    ...(paragraphReflow ? { paragraphReflow } : {}),
    ...(positionedRunGroup ? { positionedRunGroup } : {}),
    ...(positionedLines ? { positionedLines } : {}),
    ...(consolidatedRuns ? { consolidatedRuns } : {}),
    ...(paragraphRewrite ? { paragraphRewrite } : {}),
    ...(positionedFrameCapacity !== undefined ? { positionedFrameCapacity } : {}),
    ...(sameLineCapacity !== undefined ? { sameLineCapacity } : {}),
    ...(record.positionedRestore === true ? { positionedRestore: true as const } : {}),
    ...(emptiedShow ? { emptiedShow: true as const } : {}),
    ...(emptiedShowPen !== undefined ? { emptiedShowPen } : {}),
  });
}
