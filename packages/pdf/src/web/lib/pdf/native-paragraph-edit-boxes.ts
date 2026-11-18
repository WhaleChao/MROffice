/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pdfTextListMarkerOf } from "./text-list-marker";
import { type PdfNativeListTextSelectionItemEvidence } from "./engine/native/list-text-fidelity";
import { PDF_NATIVE_MAX_FIXED_FRAME_ITEM_INDEX_GAP, PDF_NATIVE_MAX_FIXED_FRAME_ITEMS, PDF_NATIVE_MAX_ITEM_TEXT_LENGTH } from "./engine/native/vocabulary";
import { summarizePdfNativeSelectionItems } from "./engine/native/list-text-model";
import { multiplyMatrix, type PdfViewportLike } from "./text-boxes";
import type { PdfNativeKeptTextBlock } from "./native-kept-text-blocks";
import { nativeListTextItemEvidence } from "./native-list-text-item-evidence";

const UPRIGHT_ANGLE_EPSILON = 0.05;
const SAME_LINE_GAP_FACTOR = 1.5;
const PARAGRAPH_GAP_FACTOR = 0.4;
// A heading is read as such from 14 pt over a 12 pt body (1.167): the most
// common pair there is, and one a reader makes by enlarging a heading. At

const HEADING_FONT_RATIO = 1.15;
const PDFJS_SPLIT_SPACE_TOLERANCE = 0.05;
const PDFJS_SPLIT_SPACE_MAX_EM = 1.5;
// pdf.js reads a gap between two glyphs of at least this many ems as a space
// (its tracking bound) and keeps reading on within the in-flow bound.
const PDFJS_TRACKING_SPACE_EM = 0.102;
const PDFJS_SPACE_IN_FLOW_MAX_EM = 0.6;
// pdf.js keeps a run in the item it is building while its baseline stands
// within this share of the item's height (VERTICAL_SHIFT_RATIO).
const PDFJS_VERTICAL_SHIFT_RATIO = 0.25;
// Table cells: the least gap between two cells of a row (wider than a word
// space) and how closely two lines of one cell share its left edge.
const TABLE_CELL_MIN_GAP_EM = 0.5;
const TABLE_CELL_EDGE_EM = 0.02;
// A bullet, or a number or a letter label ("a.", "iv)") -- not a word that
// ends a sentence ("coupling.").
const TRAILING_LIST_MARKER = /^(?:[•◦▪‣·]|(?:\d+|[A-Za-z]{1,4})[.)])$/u;
const COMPACT_NUMERIC_CHART_LABEL = /^[+-]?(?:\d{1,4}(?:[.,]\d+)?|[.,]\d+)%?$/u;

export type PdfNativeParagraphEditBlockReason =
  | "too-many-items"
  | "noncontiguous-items"
  | "mixed-font-items"
  | "incomplete-line-boundaries"
  | "unsupported-geometry"
  | "not-text";

/**
 * A code its font names no character for is shown by the reader as its code
 * number; below 0x20 that is a control character. TeX's math extension fonts
 * are drawn that way (a radical, a brace, an operator), and no one can type
 * what such a box shows: it is a picture of glyphs, not text.
 */
const UNNAMED_CODE_CHARACTER = /[\u0000-\u0008\u000b-\u001f\u007f]/u;

export type PdfNativeParagraphEditBox = Readonly<{
  id: string;
  x: number;
  y: number;
  width: number;
  height: number;
  angle?: number;
  text: string;
  items: readonly PdfNativeListTextSelectionItemEvidence[];
  /** Raw pdf.js item identities owned by this visual paragraph boundary. */
  sourceItemIndexes: readonly number[];
  /**
   * One visual line can be fragmented by pdf.js even when the source owns a
   * single text-show operator (for example, repeated ASCII spaces). This
   * synthetic item is selection-only evidence for that exact visual line.
   * The native planner must still prove one matching source operator before
   * it authorizes a write.
   */
  selectionItem?: PdfNativeListTextSelectionItemEvidence;
  editable: boolean;
  disabledReason?: PdfNativeParagraphEditBlockReason;
}>;

type ProjectedNativeItem = Readonly<{
  evidence: PdfNativeListTextSelectionItemEvidence;
  visualFontKey: string;
  left: number;
  right: number;
  top: number;
  bottom: number;
  baselineX: number;
  baselineY: number;
  baselineInline: number;
  baselineBlock: number;
  inlineStart: number;
  inlineEnd: number;
  blockStart: number;
  blockEnd: number;
  fontSize: number;
  angle: number;
  vertical: boolean;
}>;

/** Render-only layout in the same baseline-minus-em convention as text boxes. */
export function nativeTextItemLayoutGeometry(
  item: PdfNativeListTextSelectionItemEvidence,
  viewport: PdfViewportLike,
) {
  const projected = projectNativeItem(item, viewport, item.fontName, false);
  if (!projected) return null;
  const radians = (projected.angle * Math.PI) / 180;
  const width = projected.inlineEnd - projected.inlineStart;
  const height = projected.blockEnd - projected.blockStart;
  const inlineCenter = (projected.inlineStart + projected.inlineEnd) / 2;
  const blockCenter = (projected.blockStart + projected.blockEnd) / 2;
  return Object.freeze({
    x:
      (inlineCenter * Math.cos(radians) -
        blockCenter * Math.sin(radians) -
        width / 2) /
      viewport.width,
    y:
      (inlineCenter * Math.sin(radians) +
        blockCenter * Math.cos(radians) -
        height / 2) /
      viewport.height,
    width: width / viewport.width,
    height: height / viewport.height,
    fontSize: projected.fontSize / viewport.height,
    angle: projected.angle,
  });
}

type NativeVisualLine = Readonly<{
  kind: "text" | "drop-cap";
  items: readonly ProjectedNativeItem[];
  left: number;
  right: number;
  top: number;
  bottom: number;
  baselineY: number;
  inlineStart: number;
  inlineEnd: number;
  blockStart: number;
  blockEnd: number;
  baselineBlock: number;
  bodyFontSize: number;
  angle: number;
  text: string;
}>;

type VisualColumn = {
  lines: NativeVisualLine[];
  inlineStart: number;
  inlineEnd: number;
  angle: number;
};

type MutableParagraph = {
  lines: NativeVisualLine[];
  left: number;
  right: number;
  top: number;
  bottom: number;
  inlineStart: number;
  inlineEnd: number;
  blockStart: number;
  blockEnd: number;
};

type PdfJsNormalizedParagraphItem = PdfNativeListTextSelectionItemEvidence &
  Readonly<{ sourceItemIndexes?: readonly number[];
    /** The width of the reader's space folded before the item's text: its
     * first glyph stands that far in from the item's origin. */
    foldedLeadWidth?: number }>;

function rawSourceItemIndexes(
  item: PdfNativeListTextSelectionItemEvidence,
): readonly number[] {
  return (
    (item as PdfJsNormalizedParagraphItem).sourceItemIndexes ?? [item.itemIndex]
  );
}

/**
 * pdf.js can expose an ASCII space inside one native `Tj` as a separate
 * zero-height item. Keep the raw items for DOM binding, but merge each exact,
 * geometry-proven split back into its native text run: after a standalone list
 * marker the space belongs to that marker, otherwise it prefixes the following
 * visible item. This retains native operator ownership without inventing gaps.
 */
/**
 * Whether the item opens its line: no inked item before it in the reader's
 * order ends on its baseline within an em and a half before it. A list
 * marker opens its line; the same digits and full stop after a decimal
 * point ("0.719." read as "0", ".", "719.") only end a number (a column
 * further left on the same baseline stands farther off than that).
 */
function itemOpensItsLine(
  items: readonly (PdfNativeListTextSelectionItemEvidence | null)[],
  at: number,
): boolean {
  const marker = items[at];
  if (!marker) return false;
  const height = Math.hypot(marker.transform[2], marker.transform[3]);
  for (let back = at - 1; back >= 0; back -= 1) {
    const before = items[back];
    if (!before || !before.str.trim()) continue;
    if (before.hasEOL) return true;
    const gap = marker.transform[4] - (before.transform[4] + before.width);
    return Math.abs(before.transform[5] - marker.transform[5]) > height * 0.5 || gap < -0.5 || gap > height * 1.5;
  }
  return true;
}

/**
 * Items the reader reports for a pass drawn again over text it already
 * reported: a Qt fake bold draws a word, then the same word 0.5 pt along
 * inside an empty /ActualText span, and pdf.js lists that pass as an item of
 * its own over the first one. It is the same text in the same place; the
 * first item's box edits both passes, so the pass gets no box of its own
 * (a box over it could not be opened and hid the editable one under it).
 * Kept to the shape of such a pass: the same font and glyph matrix, the
 * text the earlier item ends with, ending where it ends within a tenth of
 * an em.
 */
export function pdfJsOverprintItemIndexes(
  items: readonly (PdfNativeListTextSelectionItemEvidence | null)[],
): ReadonlySet<number> {
  const overprints = new Set<number>();
  for (let index = 1; index < items.length; index += 1) {
    const pass = items[index];
    if (!pass || pass.height <= 0 || !/\S/u.test(pass.str)) continue;
    const em = Math.hypot(pass.transform[2], pass.transform[3]);
    if (!(em > 0) || Math.abs(pass.transform[1]) > PDFJS_SPLIT_SPACE_TOLERANCE) continue;
    const tolerance = em * 0.1;
    for (let back = index - 1; back >= Math.max(0, index - 4); back -= 1) {
      const drawn = items[back];
      if (
        drawn &&
        drawn.height > 0 &&
        !overprints.has(drawn.itemIndex) &&
        drawn.fontName === pass.fontName &&
        drawn.str.endsWith(pass.str) &&
        [0, 1, 2, 3].every((component) =>
          Math.abs(drawn.transform[component] - pass.transform[component]) <= PDFJS_SPLIT_SPACE_TOLERANCE) &&
        Math.abs(drawn.transform[5] - pass.transform[5]) <= tolerance &&
        Math.abs(drawn.transform[4] + drawn.width - (pass.transform[4] + pass.width)) <= tolerance &&
        pass.width <= drawn.width + tolerance
      ) {
        overprints.add(pass.itemIndex);
        break;
      }
    }
  }
  return overprints;
}

export function normalizePdfJsNativeParagraphItems(
  items: readonly (PdfNativeListTextSelectionItemEvidence | null)[],
  /** Where the page's pinned boxes start (pdfNativeBoxOriginsOf): the
   * reader's space before the first item of a box is the gap between two
   * boxes, and is folded into neither. */
  boxOrigins: readonly Readonly<{ x: number; y: number }>[] = [],
): readonly (PdfNativeListTextSelectionItemEvidence | null)[] {
  const normalized = [...items];
  const opensBox = (item: PdfNativeListTextSelectionItemEvidence) => boxOrigins.some((origin) =>
    Math.abs(item.transform[4] - origin.x) <= PINNED_BOX_ORIGIN_TOLERANCE &&
    Math.abs(item.transform[5] - origin.y) <= PINNED_BOX_ORIGIN_TOLERANCE);
  for (let index = 0; index + 1 < items.length; index += 1) {
    const previous = index > 0 ? items[index - 1] : null;
    const space = items[index];
    const visible = items[index + 1];
    // A box the reader read before a write on its page starts at this item
    // (the page's box map): two e-mail addresses 1.46 em apart, each a box,
    // stood one after the other in the page's order once the name block
    // between them was set again -- the reader then read the gap between
    // them as a space, it was folded before the second address, whose box no
    // longer found its first item where it was pinned, and the two addresses
    
    if (space && visible && /^ +$/u.test(space.str) && space.height === 0 && opensBox(visible)) continue;
    // A space reported on a raised item's baseline and in its size (after an
    // exponent or an isotope's mass number, pdf.js keeps the superscript's
    // transform) still prefixes the item after it, on that item's own
    // baseline and in its size.
    const afterRaised = Boolean(
      space && visible && previous && previous.height > 0 && previous.itemIndex + 1 === space.itemIndex &&
      [0, 3, 5].every((component) => Math.abs(space.transform[component] - previous.transform[component]) <= PDFJS_SPLIT_SPACE_TOLERANCE) &&
      Math.abs(space.transform[5] - visible.transform[5]) > PDFJS_SPLIT_SPACE_TOLERANCE &&
      Math.abs(space.transform[5] - visible.transform[5]) <= Math.hypot(visible.transform[2], visible.transform[3]) * 0.55,
    );
    // The gap after glyphs a font names no character for (a TeX delimiter
    // drawn from its extension face, then a fraction raised beside it) is
    // measured from glyphs that are no text: the space it reads as belongs
    // to neither item, and stays a spacer between them.
    if (
      previous &&
      previous.itemIndex + 1 === space?.itemIndex &&
      UNNAMED_CODE_CHARACTER.test(previous.str)
    ) {
      continue;
    }
    const flush = Boolean(
      space &&
        visible &&
        Math.abs(space.transform[4] + space.width - visible.transform[4]) <=
          Math.max(PDFJS_SPLIT_SPACE_TOLERANCE, space.width * 0.005),
    );
    // A producer that sets each word with its own `Td` and draws the space
    // glyph between them as a show of its own: the reader reports that glyph
    // as a zero-height item where the word before it ends, and the pen then
    // jumps on to the next word. The reader reads the whole jump -- from the
    // word's end to the next word's origin -- as the space, as it reads a
    // space it makes itself (a gap of 0.102 em or more, within its in-flow
    // bound); the item after it is read as starting at the jump. The word
    // before may be set slanted (a synthetic oblique leans its up axis, never
    // its baseline): its width still runs along the line to the space.
    const em = visible ? Math.hypot(visible.transform[2], visible.transform[3]) : 0;
    const jump = space && visible ? visible.transform[4] - space.transform[4] : 0;
    const spacedAfterPrevious = Boolean(
      space &&
        visible &&
        previous &&
        previous.height > 0 &&
        previous.itemIndex + 1 === space.itemIndex &&
        Math.abs(previous.transform[1]) <= PDFJS_SPLIT_SPACE_TOLERANCE &&
        Math.abs(previous.transform[3] - space.transform[3]) <= PDFJS_SPLIT_SPACE_TOLERANCE &&
        Math.abs(previous.transform[5] - space.transform[5]) <= PDFJS_SPLIT_SPACE_TOLERANCE &&
        Math.abs(previous.transform[4] + previous.width - space.transform[4]) <=
          Math.max(PDFJS_SPLIT_SPACE_TOLERANCE, space.width * 0.005) &&
        jump > space.width &&
        jump >= em * PDFJS_TRACKING_SPACE_EM &&
        jump <= em * PDFJS_SPACE_IN_FLOW_MAX_EM,
    );
    if (
      !space ||
      !visible ||
      !/^ +$/u.test(space.str) ||
      space.height !== 0 ||
      space.hasEOL ||
      visible.height <= 0 ||
      visible.itemIndex !== space.itemIndex + 1 ||
      Math.abs(space.transform[1]) > PDFJS_SPLIT_SPACE_TOLERANCE ||
      Math.abs(visible.transform[1]) > PDFJS_SPLIT_SPACE_TOLERANCE ||
      (!afterRaised && (
        Math.abs(space.transform[0] - visible.transform[0]) > PDFJS_SPLIT_SPACE_TOLERANCE ||
        Math.abs(space.transform[3] - visible.transform[3]) > PDFJS_SPLIT_SPACE_TOLERANCE ||
        Math.abs(space.transform[5] - visible.transform[5]) > PDFJS_SPLIT_SPACE_TOLERANCE)) ||
      space.width >
        Math.max(
          PDFJS_SPLIT_SPACE_TOLERANCE,
          Math.hypot(visible.transform[2], visible.transform[3]) *
            PDFJS_SPLIT_SPACE_MAX_EM,
        ) ||
      !(flush || spacedAfterPrevious)
    ) {
      continue;
    }
    if (
      flush &&
      previous &&
      // A marker already merged with the space before it (a middle dot
      // between two gaps in a TeX line) is not set again with the space
      // after it: that space prefixes the item that follows instead.
      normalized[index - 1] === previous &&
      TRAILING_LIST_MARKER.test(previous.str.trim()) &&
      itemOpensItsLine(items, index - 1) &&
      previous.height > 0 &&
      previous.itemIndex + 1 === space.itemIndex &&
      Math.abs(previous.transform[1]) <= PDFJS_SPLIT_SPACE_TOLERANCE &&
      Math.abs(previous.transform[2]) <= PDFJS_SPLIT_SPACE_TOLERANCE &&
      Math.abs(previous.transform[0] - space.transform[0]) <=
        PDFJS_SPLIT_SPACE_TOLERANCE &&
      Math.abs(previous.transform[3] - space.transform[3]) <=
        PDFJS_SPLIT_SPACE_TOLERANCE &&
      Math.abs(previous.transform[5] - space.transform[5]) <=
        PDFJS_SPLIT_SPACE_TOLERANCE &&
      Math.abs(previous.transform[4] + previous.width - space.transform[4]) <=
        Math.max(PDFJS_SPLIT_SPACE_TOLERANCE, space.width * 0.005)
    ) {
      normalized[index - 1] = Object.freeze({
        ...previous,
        str: previous.str + space.str,
        width: previous.width + space.width,
        hasEOL: space.hasEOL,
        sourceItemIndexes: Object.freeze([
          ...rawSourceItemIndexes(previous),
          space.itemIndex,
        ]),
      });
      normalized[index] = null;
      index += 1;
      continue;
    }
    normalized[index] = Object.freeze({
      ...visible,
      itemIndex: space.itemIndex,
      str: space.str + visible.str,
      transform: Object.freeze([
        visible.transform[0],
        visible.transform[1],
        visible.transform[2],
        visible.transform[3],
        space.transform[4],
        visible.transform[5],
      ]) as PdfNativeListTextSelectionItemEvidence["transform"],
      width: flush
        ? space.width + visible.width
        : jump + visible.width,
      sourceItemIndexes: Object.freeze([space.itemIndex, visible.itemIndex]),
      foldedLeadWidth: flush ? space.width : jump,
    });
    normalized[index + 1] = null;
    index += 1;
  }
  foldSkewedFrameLineEnds(normalized);
  return Object.freeze(normalized);
}






















function foldSkewedFrameLineEnds(
  normalized: (PdfNativeListTextSelectionItemEvidence | null)[],
): void {
  let before: PdfNativeListTextSelectionItemEvidence | null = null;
  for (let index = 0; index < normalized.length; index += 1) {
    const after = normalized[index];
    if (!after) continue;
    const previous = before;
    before = after;
    if (
      !previous ||
      !previous.hasEOL ||
      !(previous.height > 0) ||
      !(after.height > 0) ||
      !/^\S/u.test(after.str) ||
      !/\S$/u.test(previous.str) ||
      rawSourceItemIndexes(after)[0] !==
        (rawSourceItemIndexes(previous).at(-1) ?? -2) + 1
    ) {
      continue;
    }
    const [a0, a1, a2, a3, ax, ay] = previous.transform;
    const [b0, b1, b2, b3, bx, by] = after.transform;
    const axis = Math.hypot(a0, a1);
    const nextAxis = Math.hypot(b0, b1);
    // Both runs run along the page's x axis, one way.
    if (
      !(axis > 0) ||
      !(nextAxis > 0) ||
      Math.abs(a1) > PDFJS_SPLIT_SPACE_TOLERANCE ||
      Math.abs(b1) > PDFJS_SPLIT_SPACE_TOLERANCE ||
      !(a0 > 0) ||
      !(b0 > 0) ||
      Math.abs(a0 - b0) > PDFJS_SPLIT_SPACE_TOLERANCE ||
      Math.abs(a3 - b3) > PDFJS_SPLIT_SPACE_TOLERANCE
    ) {
      continue;
    }
    // The reader compared in the next run's own frame: its transform is not
    // axis-aligned, and the two frames differ in their second axis.
    if (b2 === 0 || Math.abs(a2 - b2) <= PDFJS_SPLIT_SPACE_TOLERANCE) continue;
    // The pen after the upright run's last glyph (prevTransform), and the
    // slanted run's origin, each through its own matrix's inverse: the
    // reader ended the line only because these lie a line apart.
    const penX = ax + previous.width;
    const lastAcross = (a2 * penX + a3 * ay) / axis;
    const nextAcross = (b2 * bx + b3 * by) / nextAxis;
    if (!(Math.abs(nextAcross - lastAcross) > previous.height)) continue;
    // On the page the two share a baseline, as close as the reader keeps a
    // run in the item it is building, and the gap reads as one space.
    const gap = bx - penX;
    if (
      Math.abs(by - ay) > previous.height * PDFJS_VERTICAL_SHIFT_RATIO ||
      gap < axis * PDFJS_TRACKING_SPACE_EM ||
      gap > axis * PDFJS_SPLIT_SPACE_MAX_EM
    ) {
      continue;
    }
    const folded: PdfNativeListTextSelectionItemEvidence = Object.freeze({
      ...after,
      str: ` ${after.str}`,
      transform: Object.freeze([
        b0,
        b1,
        b2,
        b3,
        penX,
        by,
      ]) as PdfNativeListTextSelectionItemEvidence["transform"],
      width: gap + after.width,
      sourceItemIndexes: Object.freeze([...rawSourceItemIndexes(after)]),
    });
    normalized[index] = folded;
    before = folded;
  }
}

/**
 * The raw pdf.js items that are nothing but a line end. pdf.js ends a line
 * (compareWithLastPosition's appendEOL) on the item it is building; where it
 * has just flushed that item -- at a change of size or face, before it
 * compares the next glyph -- it pushes an empty item of its own to carry the
 * EOL instead. The viewer's item evidence drops empty items, and with them the
 * reader's line end: pass these to nativeParagraphEditBoxes, which gives the
 * line end back to the item it closes when both stand in one line of a box.
 */
export function pdfJsEmptyLineEndItemIndexes(
  rawItems: readonly unknown[],
): ReadonlySet<number> {
  const indexes = new Set<number>();
  for (const [index, value] of rawItems.entries()) {
    const item = value as { str?: unknown; hasEOL?: unknown } | null;
    if (item && typeof item === "object" && item.str === "" && item.hasEOL === true) indexes.add(index);
  }
  return indexes;
}

function finitePositive(value: number): boolean {
  return Number.isFinite(value) && value > 0;
}

function projectNativeItem(
  evidence: PdfNativeListTextSelectionItemEvidence,
  viewport: PdfViewportLike,
  visualFontKey: string,
  vertical: boolean,
): ProjectedNativeItem | null {
  const matrix = multiplyMatrix(viewport.transform, [...evidence.transform]);
  const fontSize = vertical
    ? Math.hypot(matrix[0], matrix[1])
    : Math.hypot(matrix[2], matrix[3]);
  const viewportScale = Math.hypot(
    viewport.transform[0],
    viewport.transform[1],
  );
  const advance =
    (vertical ? evidence.height : evidence.width) * viewportScale;
  if (
    !finitePositive(viewport.width) ||
    !finitePositive(viewport.height) ||
    !finitePositive(fontSize) ||
    !finitePositive(advance)
  ) {
    return null;
  }

  const inlineScale = vertical
    ? Math.hypot(matrix[2], matrix[3])
    : Math.hypot(matrix[0], matrix[1]);
  if (!finitePositive(inlineScale)) return null;
  const inlineX = vertical
    ? -matrix[2] / inlineScale
    : matrix[0] / inlineScale;
  const inlineY = vertical
    ? -matrix[3] / inlineScale
    : matrix[1] / inlineScale;
  const angleRadians = Math.atan2(inlineY, inlineX);
  const angle = (angleRadians * 180) / Math.PI;
  const cosine = inlineX;
  const sine = inlineY;
  const baselineEnd = {
    x: matrix[4] + cosine * advance,
    y: matrix[5] + sine * advance,
  };
  const topOffset = { x: sine * fontSize, y: -cosine * fontSize };
  const bottomOffset = {
    x: -sine * fontSize * 0.25,
    y: cosine * fontSize * 0.25,
  };
  const points = [
    { x: matrix[4] + topOffset.x, y: matrix[5] + topOffset.y },
    { x: baselineEnd.x + topOffset.x, y: baselineEnd.y + topOffset.y },
    { x: matrix[4] + bottomOffset.x, y: matrix[5] + bottomOffset.y },
    { x: baselineEnd.x + bottomOffset.x, y: baselineEnd.y + bottomOffset.y },
  ];
  const left = Math.min(...points.map((point) => point.x));
  const right = Math.max(...points.map((point) => point.x));
  const top = Math.min(...points.map((point) => point.y));
  const bottom = Math.max(...points.map((point) => point.y));
  const inlineCoordinates = points.map(
    (point) => point.x * cosine + point.y * sine,
  );
  const blockCoordinates = points.map(
    (point) => point.x * -sine + point.y * cosine,
  );
  const baselineInline = matrix[4] * cosine + matrix[5] * sine;
  const baselineBlock = matrix[4] * -sine + matrix[5] * cosine;
  if (
    !Number.isFinite(left) ||
    !Number.isFinite(right) ||
    !Number.isFinite(top) ||
    !Number.isFinite(bottom) ||
    right <= left ||
    bottom <= top
  ) {
    return null;
  }
  return Object.freeze({
    evidence,
    visualFontKey,
    left,
    right,
    top,
    bottom,
    baselineX: matrix[4],
    baselineY: matrix[5],
    baselineInline,
    baselineBlock,
    inlineStart: Math.min(...inlineCoordinates),
    inlineEnd: Math.max(...inlineCoordinates),
    blockStart: Math.min(...blockCoordinates),
    blockEnd: Math.max(...blockCoordinates),
    fontSize,
    angle,
    vertical,
  });
}

function dominantLineFontSize(items: readonly ProjectedNativeItem[]): number {
  const weights = new Map<number, number>();
  for (const item of items) {
    const key = Math.round(item.fontSize * 1_000) / 1_000;
    const characters = Math.max(1, Array.from(item.evidence.str.trim()).length);
    weights.set(key, (weights.get(key) ?? 0) + characters);
  }
  return [...weights.entries()].reduce(
    (best, entry) => (entry[1] > best[1] ? entry : best),
    [items[0].fontSize, -1] as [number, number],
  )[0];
}











function stackedScriptsInSourceOrder(
  ordered: readonly ProjectedNativeItem[],
): ProjectedNativeItem[] {
  const result: ProjectedNativeItem[] = [];
  for (let start = 0; start < ordered.length;) {
    let end = start + 1;
    let reach = ordered[start].inlineEnd;
    while (end < ordered.length &&
      ordered[end].inlineStart < reach - Math.min(ordered[end].fontSize, ordered[start].fontSize) * 0.01) {
      reach = Math.max(reach, ordered[end].inlineEnd);
      end += 1;
    }
    const cluster = ordered.slice(start, end);
    const stacked = cluster.length > 1 && cluster.some((item) =>
      Math.abs(item.baselineBlock - cluster[0].baselineBlock) >
        Math.min(item.fontSize, cluster[0].fontSize) * PDFJS_VERTICAL_SHIFT_RATIO);
    result.push(...(stacked
      ? [...cluster].sort((left, right) => left.evidence.itemIndex - right.evidence.itemIndex)
      : cluster));
    start = end;
  }
  return result;
}

const STRONG_RTL = /[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/gu;
const STRONG_LTR = /[A-Za-z\u00C0-\u024F\u0370-\u03FF\u0400-\u04FF\u3040-\u30FF\u4E00-\u9FFF\uAC00-\uD7AF]/gu;

/** Whether a line's letters are mostly of a right-to-left script. */
function rightToLeftLine(items: readonly ProjectedNativeItem[]): boolean {
  const text = items.map((item) => item.evidence.str).join("");
  const rtl = text.match(STRONG_RTL)?.length ?? 0;
  return rtl > 0 && rtl > (text.match(STRONG_LTR)?.length ?? 0);
}

function lineFromItems(
  items: readonly ProjectedNativeItem[],
  kind: NativeVisualLine["kind"] = "text",
): NativeVisualLine | null {
  const ordered = stackedScriptsInSourceOrder([...items].sort(
    (left, right) =>
      left.baselineInline - right.baselineInline ||
      left.evidence.itemIndex - right.evidence.itemIndex,
  ));
  // A right-to-left line reads its runs from the right: the Mozilla Arabic
  // sample's three words, each its own pdf.js run, were joined left to
  
  // direction inside it keep their own order, as bidi lays them. The items
  // stay in page order (the line's geometry is read from them).
  // (A space pdf.js reports before a run trails it once the runs read from
  // the right.)
  const text = rightToLeftLine(ordered)
    ? [...ordered].reverse().map((item) => {
      const [, leading, body, trailing] = /^(\s*)([\s\S]*?)(\s*)$/u.exec(item.evidence.str)!;
      return trailing + body + leading;
    }).join("")
    : ordered.map((item) => item.evidence.str).join("");
  if (!text.trim()) return null;
  return Object.freeze({
    kind,
    items: Object.freeze(ordered),
    left: Math.min(...ordered.map((item) => item.left)),
    right: Math.max(...ordered.map((item) => item.right)),
    top: Math.min(...ordered.map((item) => item.top)),
    bottom: Math.max(...ordered.map((item) => item.bottom)),
    baselineY:
      ordered.reduce((total, item) => total + item.baselineY, 0) /
      ordered.length,
    baselineBlock:
      ordered.reduce((total, item) => total + item.baselineBlock, 0) /
      ordered.length,
    inlineStart: Math.min(...ordered.map((item) => item.inlineStart)),
    inlineEnd: Math.max(...ordered.map((item) => item.inlineEnd)),
    blockStart: Math.min(...ordered.map((item) => item.blockStart)),
    blockEnd: Math.max(...ordered.map((item) => item.blockEnd)),
    bodyFontSize: dominantLineFontSize(ordered),
    angle: ordered[0].angle,
    text,
  });
}

function splitDropCapLine(
  items: readonly ProjectedNativeItem[],
): readonly NativeVisualLine[] {
  const ordered = [...items].sort(
    (left, right) =>
      left.baselineInline - right.baselineInline ||
      left.evidence.itemIndex - right.evidence.itemIndex,
  );
  const bodyFontSize = dominantLineFontSize(ordered);
  let prefixLength = 0;
  for (const item of ordered) {
    if (item.fontSize < bodyFontSize * 1.5) break;
    const candidate = ordered
      .slice(0, prefixLength + 1)
      .map((entry) => entry.evidence.str)
      .join("")
      .trim();
    if (Array.from(candidate).length > 3) break;
    prefixLength += 1;
  }
  if (prefixLength === 0 || prefixLength === ordered.length) {
    const line = lineFromItems(ordered);
    return line ? Object.freeze([line]) : Object.freeze([]);
  }

  const dropCapItems = ordered
    .slice(0, prefixLength)
    .filter((item) => item.evidence.str.trim());
  const bodyItems = ordered.slice(prefixLength);
  const dropCap = lineFromItems(dropCapItems, "drop-cap");
  const body = lineFromItems(bodyItems);
  return Object.freeze(
    [dropCap, body].filter((line): line is NativeVisualLine => line !== null),
  );
}









function runsCollideAcrossSizes(
  row: readonly ProjectedNativeItem[],
  item: ProjectedNativeItem,
): boolean {
  return row.some((entry) => {
    const larger = Math.max(entry.fontSize, item.fontSize);
    const smaller = Math.min(entry.fontSize, item.fontSize);
    if (larger <= smaller * 1.15) return false;
    const overlap = Math.min(entry.inlineEnd, item.inlineEnd) - Math.max(entry.inlineStart, item.inlineStart);
    return overlap > smaller * 0.1 &&
      Math.abs(entry.baselineBlock - item.baselineBlock) <= smaller * PDFJS_VERTICAL_SHIFT_RATIO;
  });
}










function runsOverlayAcrossSizes(
  row: readonly ProjectedNativeItem[],
  item: ProjectedNativeItem,
): boolean {
  return row.some((entry) => {
    const larger = Math.max(entry.fontSize, item.fontSize);
    const smaller = Math.min(entry.fontSize, item.fontSize);
    if (larger <= smaller * 1.15) return false;
    // A script (a raised "1/3", a stacked sub- and superscript) is a few
    // glyphs over its base: that is one line. A line of text in the smaller
    // size, three ems of the larger size or longer, is another line.
    const small = entry.fontSize < item.fontSize ? entry : item;
    if (small.inlineEnd - small.inlineStart <= larger * 3) return false;
    const overlap = Math.min(entry.inlineEnd, item.inlineEnd) - Math.max(entry.inlineStart, item.inlineStart);
    return overlap > smaller * 0.1 &&
      Math.abs(entry.baselineBlock - item.baselineBlock) > smaller * PDFJS_VERTICAL_SHIFT_RATIO;
  });
}

/**
 * TeX's radical sign hangs from its origin: the sign is drawn downward from
 * the height of its bar, three quarters of an em above the baseline of the
 * line it stands on, and the reader reports it up there. The glyph it
 * stands before -- the first of what it covers -- starts where the sign
 * ends, half an em to an em lower: that glyph's line is the sign's line.
 */
function radicandOf(
  item: ProjectedNativeItem,
  items: readonly ProjectedNativeItem[],
): ProjectedNativeItem | null {
  if (item.vertical || item.evidence.str.trim() !== "\u221a") return null;
  const size = item.fontSize;
  return items.find((other) =>
    other !== item && !other.vertical && other.evidence.str.trim().length > 0 &&
    Math.abs(other.angle - item.angle) < UPRIGHT_ANGLE_EPSILON &&
    other.baselineBlock - item.baselineBlock >= size * 0.5 &&
    other.baselineBlock - item.baselineBlock <= size &&
    Math.abs(other.inlineStart - item.inlineEnd) <= size * 0.25) ?? null;
}

/**
 * Radical signs the rows left off the line they stand on (visualLines): each
 * a row, and a box, of its own.
 */
const STRAY_RADICALS = new WeakSet<ProjectedNativeItem>();

function lineIsStrayRadical(line: NativeVisualLine): boolean {
  return line.items.length === 1 && STRAY_RADICALS.has(line.items[0]);
}

/**
 * Where an item's first glyph stands: after the reader's space folded before
 * it (normalizePdfJsNativeParagraphItems), which is the gap itself.
 */
function inkStart(item: ProjectedNativeItem): number {
  const lead = (item.evidence as PdfJsNormalizedParagraphItem).foldedLeadWidth ?? 0;
  return lead > 0 && item.evidence.width > 0
    ? item.inlineStart + (item.inlineEnd - item.inlineStart) * Math.min(1, lead / item.evidence.width)
    : item.inlineStart;
}










function figuresSetApart(
  left: readonly ProjectedNativeItem[],
  right: readonly ProjectedNativeItem[],
): boolean {
  let pair: readonly [ProjectedNativeItem, ProjectedNativeItem] | null = null;
  let nearest = Number.POSITIVE_INFINITY;
  for (const a of left) {
    for (const b of right) {
      const gap = Math.max(inkStart(a) - b.inlineEnd, inkStart(b) - a.inlineEnd);
      if (gap < nearest) {
        nearest = gap;
        pair = [a, b];
      }
    }
  }
  if (pair === null) return false;
  const size = Math.min(pair[0].fontSize, pair[1].fontSize);
  // (On one baseline, in one size: a formula's figures -- an exponent, a
  // fraction's parts -- stand off the line they belong to and stay with it.)
  return nearest > size * PDFJS_SPACE_IN_FLOW_MAX_EM &&
    Math.abs(pair[0].baselineBlock - pair[1].baselineBlock) <= size * PDFJS_SPLIT_SPACE_TOLERANCE &&
    Math.abs(pair[0].fontSize - pair[1].fontSize) <= size * 0.05 &&
    pair.every((item) => COMPACT_NUMERIC_CHART_LABEL.test(item.evidence.str.trim()));
}

/**
 * Justified to its frame, a line of few words keeps them far apart: a
 * caption's "EXPRESS EMPLOYMENT" set with `4.713 Tw` across a 229 pt column
 * stands five ems apart, and the reader, which keeps a word within
 * SAME_LINE_GAP_FACTOR ems of its line, read "EMPLOYMENT" as a box of its own
 * beside the paragraph it opens. Bold taken off the paragraph left that word
 * bold, and the paragraph, its first line no longer reaching the frame, was
 * set again ragged (walk 0444). Two rows are that one line where the page
 * says so on every count: they stand on one baseline in one size and one
 * face, the reader read the right one straight after the left one, neither
 * holds a word space of its own narrower than the one between them (a
 * justified line stretches all its spaces alike; a column's or a cell's
 * line keeps its words a space apart), and the line read next to them, a
 * line's pitch away, runs across the gap and ends where the right row ends
 * -- the frame the line was justified to.
 */
function joinRowsOfOneJustifiedLine(
  rows: ProjectedNativeItem[][],
  page: readonly ProjectedNativeItem[],
): void {
  const describe = (row: readonly ProjectedNativeItem[]) => {
    const byInline = [...row].sort((x, y) => x.inlineStart - y.inlineStart);
    const size = Math.max(...row.map((entry) => entry.fontSize));
    let widestInnerSpace = 0;
    let narrowestInnerSpace = Number.POSITIVE_INFINITY;
    let reach = byInline[0].inlineEnd;
    for (const entry of byInline.slice(1)) {
      const space = inkStart(entry) - reach;
      // (Closer than a word space: one word in two faces, a kerned pair.)
      if (space > size * PDFJS_TRACKING_SPACE_EM * 2) {
        widestInnerSpace = Math.max(widestInnerSpace, space);
        narrowestInnerSpace = Math.min(narrowestInnerSpace, space);
      }
      reach = Math.max(reach, entry.inlineEnd);
    }
    const indexes = row.map((entry) => entry.evidence.itemIndex);
    return {
      size,
      uniform: row.every((entry) => !entry.vertical && entry.fontSize >= size * 0.95),
      baseline: row.reduce((total, entry) => total + entry.baselineBlock, 0) / row.length,
      start: inkStart(byInline[0]),
      end: reach,
      first: Math.min(...indexes),
      last: Math.max(...indexes),
      // (A space the reader kept inside an item: a word space in flow.)
      spaced: row.some((entry) => /\S\s+\S/u.test(entry.evidence.str)),
      worded: row.some((entry) => /\p{L}/u.test(entry.evidence.str)),
      widestInnerSpace,
      narrowestInnerSpace,
    };
  };
  const described = new Map<ProjectedNativeItem[], ReturnType<typeof describe>>();
  const of = (row: ProjectedNativeItem[]) => {
    let known = described.get(row);
    if (!known) described.set(row, (known = describe(row)));
    return known;
  };
  const read = new Set(page.map((entry) => entry.evidence.itemIndex));
  for (let joined = true; joined;) {
    joined = false;
    // The row read straight after an item, or straight before one: no text
    // of the page in between, and at most the reader's own space item.
    const opening = new Map<number, ProjectedNativeItem[]>();
    const closing = new Map<number, ProjectedNativeItem[]>();
    for (const row of rows) {
      opening.set(of(row).first, row);
      closing.set(of(row).last, row);
    }
    const readAfter = (index: number) =>
      opening.get(index + 1) ?? (read.has(index + 1) ? undefined : opening.get(index + 2));
    const readBefore = (index: number) =>
      closing.get(index - 1) ?? (read.has(index - 1) ? undefined : closing.get(index - 2));
    for (const a of rows) {
      const left = of(a);
      const b = readAfter(left.last);
      if (!b || b === a || Math.abs(a[0].angle - b[0].angle) >= UPRIGHT_ANGLE_EPSILON) continue;
      const right = of(b);
      const size = Math.min(left.size, right.size);
      const gap = right.start - left.end;
      if (
        !left.uniform || !right.uniform ||
        // Words, both: a row of leader dots after its entry, a figure in
        // the next cell of its label's row are no words of a line.
        !left.worded || !right.worded ||
        Math.abs(left.size - right.size) > size * 0.05 ||
        Math.abs(left.baseline - right.baseline) > size * PDFJS_SPLIT_SPACE_TOLERANCE ||
        gap <= size * SAME_LINE_GAP_FACTOR ||
        left.spaced || right.spaced ||
        Math.min(left.narrowestInnerSpace, right.narrowestInnerSpace) < gap * 0.6 ||
        figuresSetApart(a, b)
      ) {
        continue;
      }
      const seamLeft = a.find((entry) => entry.evidence.itemIndex === left.last);
      const seamRight = b.find((entry) => entry.evidence.itemIndex === right.first);
      if (seamLeft?.evidence.fontName !== seamRight?.evidence.fontName) continue;
      const framed = [readAfter(right.last), readBefore(left.first)].some((other) => {
        if (!other || other === a || other === b) return false;
        if (Math.abs(other[0].angle - a[0].angle) >= UPRIGHT_ANGLE_EPSILON) return false;
        const line = of(other);
        const pitch = Math.abs(line.baseline - left.baseline);
        return line.uniform &&
          Math.abs(line.size - size) <= size * 0.05 &&
          pitch >= size * 0.5 && pitch <= size * 2 &&
          // A line of words a space apart, not a row of the same kind.
          (line.spaced || line.widestInnerSpace > 0) && line.widestInnerSpace < gap * 0.6 &&
          Math.abs(line.end - right.end) <= size * 0.1 &&
          Math.abs(line.start - left.start) <= size * 0.1;
      });
      if (!framed) continue;
      a.push(...b);
      a.sort((x, y) => x.baselineInline - y.baselineInline || x.evidence.itemIndex - y.evidence.itemIndex);
      rows.splice(rows.indexOf(b), 1);
      described.delete(a);
      described.delete(b);
      joined = true;
      break;
    }
  }
}

function visualLines(
  items: readonly ProjectedNativeItem[],
): readonly NativeVisualLine[] {
  const ordered = [...items].sort(
    (left, right) =>
      left.baselineBlock - right.baselineBlock ||
      left.baselineInline - right.baselineInline ||
      left.evidence.itemIndex - right.evidence.itemIndex,
  );
  const rows: ProjectedNativeItem[][] = [];
  for (const item of ordered) {
    let bestRow: ProjectedNativeItem[] | null = null;
    let bestDistance = Number.POSITIVE_INFINITY;
    for (const row of rows) {
      const sample = row[0];
      const baselineDistance = Math.abs(
        item.baselineBlock - sample.baselineBlock,
      );
      if (
        Math.abs(item.angle - sample.angle) >= UPRIGHT_ANGLE_EPSILON ||
        baselineDistance > Math.max(item.fontSize, sample.fontSize) * 0.55
      ) {
        continue;
      }
      const rowStart = Math.min(...row.map((entry) => entry.inlineStart));
      const rowEnd = Math.max(...row.map((entry) => entry.inlineEnd));
      const gap = Math.max(
        rowStart - item.inlineEnd,
        item.inlineStart - rowEnd,
        0,
      );
      if (
        gap >
        Math.min(item.fontSize, sample.fontSize) * SAME_LINE_GAP_FACTOR
      ) {
        continue;
      }
      if (runsCollideAcrossSizes(row, item) || runsOverlayAcrossSizes(row, item)) continue;
      if (figuresSetApart(row, [item])) continue;
      if (baselineDistance < bestDistance) {
        bestRow = row;
        bestDistance = baselineDistance;
      }
    }
    if (bestRow) bestRow.push(item);
    else rows.push([item]);
  }
  // Rows that are one line: a superscript or a subscript (an isotope's mass
  // number, an exponent, a footnote mark) whose baseline sorted before its
  // line's opened a row of its own, and the text after it joined that row
  // rather than the line's. Two rows within a line's own baseline distance
  // that stand beside each other are one row.
  for (let merged = true; merged;) {
    merged = false;
    for (let left = 0; left < rows.length && !merged; left += 1) {
      for (let right = left + 1; right < rows.length; right += 1) {
        const a = rows[left], b = rows[right];
        const sizeA = Math.max(...a.map((entry) => entry.fontSize));
        const sizeB = Math.max(...b.map((entry) => entry.fontSize));
        // A row's baseline is its body text's, not one its scripts pull up
        // or down: an equation's base and its raised "1/3" read 5 pt above
        // the base of the rest of the formula, and the two were kept apart
        
        const bodyBaseline = (row: readonly ProjectedNativeItem[], size: number) => {
          const body = row.filter((entry) => entry.fontSize >= size * 0.9);
          return body.reduce((total, entry) => total + entry.baselineBlock, 0) / body.length;
        };
        const baselineA = bodyBaseline(a, sizeA);
        const baselineB = bodyBaseline(b, sizeB);
        if (
          Math.abs(a[0].angle - b[0].angle) >= UPRIGHT_ANGLE_EPSILON ||
          Math.abs(baselineA - baselineB) > Math.max(sizeA, sizeB) * 0.55 ||
          // A line's own text and its superscript differ in size; two lines
          // of one size a fraction of an em apart are two lines -- unless
          // they stand on one baseline, as close as the reader keeps a run
          // in the item it is building (a slanted run set with a Tm of its
          // own a hair off the upright run's baseline opens a row of its own
          // when it sorts first).
          (Math.abs(sizeA - sizeB) <= Math.max(sizeA, sizeB) * 0.05 &&
            Math.abs(baselineA - baselineB) > Math.min(sizeA, sizeB) * PDFJS_VERTICAL_SHIFT_RATIO)
        ) {
          continue;
        }
        const startA = Math.min(...a.map((entry) => entry.inlineStart)), endA = Math.max(...a.map((entry) => entry.inlineEnd));
        const startB = Math.min(...b.map((entry) => entry.inlineStart)), endB = Math.max(...b.map((entry) => entry.inlineEnd));
        const gap = Math.max(startA - endB, startB - endA, 0);
        if (gap > Math.min(sizeA, sizeB) * SAME_LINE_GAP_FACTOR) continue;
        if (b.some((entry) => runsCollideAcrossSizes(a, entry) || runsOverlayAcrossSizes(a, entry))) continue;
        if (figuresSetApart(a, b)) continue;
        a.push(...b);
        a.sort((x, y) => x.baselineInline - y.baselineInline || x.evidence.itemIndex - y.evidence.itemIndex);
        rows.splice(right, 1);
        merged = true;
        break;
      }
    }
  }
  joinRowsOfOneJustifiedLine(rows, ordered);
  // A radical sign read onto another row than the glyph it stands before.
  // Its origin is nearer the line above its own than its own (radicandOf),
  // and rows are told by their baselines: the sign joined the line above,
  // which then read "…pre-softmax√" and, typed onto, was set again with the
  // sign after its last word, 250 pt from the "d model" it covers (Attention
  
  // the formula's drawing as TeX's taller signs are (UNNAMED_CODE_CHARACTER):
  // the line it stands on keeps the gap it stands in, and what moves that gap
  // moves the sign with it (the rewrite's carried gaps).
  for (const row of [...rows]) {
    for (const item of [...row]) {
      const radicand = radicandOf(item, ordered);
      if (!radicand || row.includes(radicand)) continue;
      STRAY_RADICALS.add(item);
      if (row.length < 2) continue;
      row.splice(row.indexOf(item), 1);
      rows.push([item]);
    }
  }
  const cells = rows.flatMap((row) => splitRowAtTableCells(row, ordered));
  // Rows stand in the order their body text does: a part set larger and
  // raised on the last row of a paragraph (a superscript the editor
  // enlarged) tops the row above it, and sorted by it the last row came
  // before that row, which then stood apart from its paragraph (Attention
  
  return Object.freeze(
    cells
      .flatMap(splitDropCapLine)
      .sort(
        (left, right) =>
          lineBodyExtent(left).blockStart - lineBodyExtent(right).blockStart ||
          left.inlineStart - right.inlineStart,
      ),
  );
}

/**
 * A table drawn cell by cell sets a row's cells on one baseline, often closer
 * together than a line would keep apart, but the reader reads each cell to
 * its end before the next: the first cell's second line comes between the
 * first cell's first line and the second cell. A row is cut into cells where
 * the reader, after the left of two neighbours standing wider apart than a
 * word space and before the right one, next reads text that starts exactly
 * where the left cell starts (its own next line) or where the right one
 * starts (the right cell's line above). Out-of-order runs of a formula
 * (accents, fraction parts, brace pieces, a word TeX set before the formula
 * to its left) do not continue a cell that way, so a line with inline
 * mathematics stays one line.
 */
function splitRowAtTableCells(
  row: readonly ProjectedNativeItem[],
  page: readonly ProjectedNativeItem[],
): ProjectedNativeItem[][] {
  const ordered = [...row].sort(
    (left, right) =>
      left.baselineInline - right.baselineInline ||
      left.evidence.itemIndex - right.evidence.itemIndex,
  );
  const own = new Set(ordered.map((item) => item.evidence.itemIndex));
  const parts: ProjectedNativeItem[][] = [[ordered[0]]];
  for (let index = 1; index < ordered.length; index += 1) {
    const before = ordered[index - 1];
    const after = ordered[index];
    const cellStart = parts[parts.length - 1][0];
    const em = Math.min(before.fontSize, after.fontSize);
    let next: ProjectedNativeItem | undefined;
    for (const item of page) {
      const at = item.evidence.itemIndex;
      if (
        at > before.evidence.itemIndex &&
        at < after.evidence.itemIndex &&
        !own.has(at) &&
        item.evidence.str.trim() !== "" &&
        (!next || at < next.evidence.itemIndex)
      ) {
        next = item;
      }
    }
    const cellBoundary =
      next !== undefined &&
      after.inlineStart - before.inlineEnd >= em * TABLE_CELL_MIN_GAP_EM &&
      Math.abs(next.angle - after.angle) < UPRIGHT_ANGLE_EPSILON &&
      (Math.abs(next.inlineStart - cellStart.inlineStart) <= em * TABLE_CELL_EDGE_EM ||
        Math.abs(next.inlineStart - after.inlineStart) <= em * TABLE_CELL_EDGE_EM);
    if (cellBoundary) parts.push([]);
    parts[parts.length - 1].push(after);
  }
  return parts;
}

/**
 * pdf.js reads a gap between two of its items as a space and reports it as a
 * white-space item of its own wherever it has no item to write the space
 * into: the gap is past its in-flow bound (a list label "(c)" set 0.74 em
 * before its text), or the item before was just flushed at a change of size
 * or face (a footnote mark after its word). It pushes that item where the pen
 * stood after the item before -- its prevTransform -- and gives it the gap
 * measured in the text matrix's own units, not the page's
 * (compareWithLastPosition, addFakeSpaces, pushWhitespace), so a producer
 * that sets its size in `Tm` rather than `Tf` gets a space item a tenth as
 * wide as its gap, which normalizePdfJsNativeParagraphItems's flush test
 * never folds. Such a spacer between two items that already form one line of
 * a box on their own is that line's word space: it is folded into the item
 * after it as its leading space, from where the item before ends and as wide
 * as the gap and the item together (the reading the planner proves as a
 * folded gap). A spacer the box does not hold on both sides stays out and
 * never bridges two boxes.
 */
function lineWithReaderGapSpaces(
  line: NativeVisualLine,
  spacers: readonly ProjectedNativeItem[],
  viewport: PdfViewportLike,
): NativeVisualLine {
  if (line.kind !== "text" || line.items.length < 2 || spacers.length === 0) return line;
  const spacerAt = new Map(spacers.map((spacer) => [spacer.evidence.itemIndex, spacer]));
  let changed = false;
  const items = line.items.map((after, index) => {
    const before = index > 0 ? line.items[index - 1] : null;
    if (!before || before.vertical || after.vertical) return after;
    const last = rawSourceItemIndexes(before.evidence).at(-1) ?? before.evidence.itemIndex;
    const spacer = spacerAt.get(last + 1)?.evidence;
    const previous = before.evidence;
    const next = after.evidence;
    if (
      !spacer ||
      rawSourceItemIndexes(next)[0] !== last + 2 ||
      previous.hasEOL ||
      // The gap after glyphs a font names no character for is measured from
      // glyphs that are no text: it stays a spacer between the two
      // (normalizePdfJsNativeParagraphItems keeps it out the same way).
      UNNAMED_CODE_CHARACTER.test(previous.str) ||
      !/\S$/u.test(previous.str) ||
      !/^\S/u.test(next.str)
    ) {
      return after;
    }
    // Both items and the spacer run along the page's x axis, one way.
    if (
      [previous, spacer, next].some(
        (item) =>
          Math.abs(item.transform[1]) > PDFJS_SPLIT_SPACE_TOLERANCE ||
          !(item.transform[0] > 0),
      )
    ) {
      return after;
    }
    // The spacer stands where the item before ends, on its baseline: the
    // pen the reader measured the gap from.
    if (
      Math.abs(previous.transform[4] + previous.width - spacer.transform[4]) >
        PDFJS_SPLIT_SPACE_TOLERANCE ||
      Math.abs(previous.transform[5] - spacer.transform[5]) > PDFJS_SPLIT_SPACE_TOLERANCE
    ) {
      return after;
    }
    // The gap is the space: it reaches the reader's tracking bound at the
    // size of the item before (the bound the reader judged it by). A narrower
    // gap is read as a space only for the white-space glyph drawn in it
    // (shouldAddWhitespace), a glyph the source still draws; that spacer
    // stays out, as it always has.
    const jump = next.transform[4] - spacer.transform[4];
    if (!(jump >= PDFJS_TRACKING_SPACE_EM * Math.hypot(previous.transform[2], previous.transform[3]))) {
      return after;
    }
    // The item keeps its own index: the box was judged on the items the
    // reader wrote text into, and the spacer is only named among its sources.
    const folded: PdfNativeListTextSelectionItemEvidence = Object.freeze({
      ...next,
      str: ` ${next.str}`,
      transform: Object.freeze([
        next.transform[0],
        next.transform[1],
        next.transform[2],
        next.transform[3],
        spacer.transform[4],
        next.transform[5],
      ]) as PdfNativeListTextSelectionItemEvidence["transform"],
      width: jump + next.width,
      sourceItemIndexes: Object.freeze([spacer.itemIndex, ...rawSourceItemIndexes(next)]),
    });
    const projectedFold = projectNativeItem(folded, viewport, after.visualFontKey, false);
    if (!projectedFold) return after;
    changed = true;
    return projectedFold;
  });
  return changed ? (lineFromItems(items, line.kind) ?? line) : line;
}

/**
 * The line a box offers as one synthetic selection item, or null when the box
 * must hand the planner its items one by one.
 *
 * One synthetic item stands for the runs pdf.js would have read as one item:
 * one font, one size, one baseline. A footnote mark in the line's own font
 * ("operations" then a raised, smaller "2") is a second run the planner has to
 * prove on its own, so that line keeps its items.
 *
 * A table-of-contents line is drawn `(Introduction)Tj ( )Tj 2 0 Td (....)Tj`:
 * pdf.js reports the space glyph as a zero-height item that does not touch
 * the leader after it, so normalizePdfJsNativeParagraphItems never folds it
 * and the line read `Introduction.....5`, which no source run spells. Such a
 * spacer is read back into the line here, between the two items it sits
 * between in the source and on their baseline; the items on either side
 * already formed the line on their own, so it never bridges two boxes. It is
 * read back only into a synthetic item: the planner refuses a zero-height
 * item of its own. The synthetic item spans from its first run to the end of
 * its last, gaps included, as pdf.js measures a run it merged itself.
 *
 * pdf.js reads runs as one item only when nothing else is drawn between them:
 * a show on another line ends the item it is building. A line whose runs the
 * reader reports with other text read between them (a chart's data labels,
 * drawn one per text object in no visual order) is no run it would have
 * merged, so that line keeps its items too.
 */
function singleRunLine(
  lines: readonly NativeVisualLine[],
  spacers: readonly ProjectedNativeItem[],
  pageItems: readonly ProjectedNativeItem[],
): NativeVisualLine | null {
  const line = lines.length === 1 ? lines[0] : null;
  if (
    !line ||
    line.kind !== "text" ||
    line.items.length < 2 ||
    line.items.some((item) => item.vertical)
  ) {
    return null;
  }
  const lineIndexes = new Set(line.items.map((item) => item.evidence.itemIndex));
  const firstIndex = Math.min(...lineIndexes), lastIndex = Math.max(...lineIndexes);
  if (
    pageItems.some(
      (item) =>
        item.evidence.itemIndex > firstIndex &&
        item.evidence.itemIndex < lastIndex &&
        !lineIndexes.has(item.evidence.itemIndex) &&
        item.evidence.str.trim() !== "",
    )
  ) {
    return null;
  }
  const first = line.items[0].evidence;
  // How far an item's origin stands off the first item's baseline: across
  // the direction the line is written in. Measured along the glyphs' up axis
  // it is the same on an upright line and on a turned one, and on a leaned
  // line (an italic drawn by its text matrix, `1 0 .21 1`) it grows with the
  // distance ALONG the line: every later piece of the one show read as off
  // its baseline, and the line kept its items.
  const xScale = Math.hypot(first.transform[0], first.transform[1]);
  const onRun = (item: PdfNativeListTextSelectionItemEvidence) =>
    item.fontName === first.fontName &&
    xScale > 0 &&
    Math.abs(
      ((item.transform[5] - first.transform[5]) * first.transform[0] -
        (item.transform[4] - first.transform[4]) * first.transform[1]) /
        xScale,
    ) <= PDFJS_SPLIT_SPACE_TOLERANCE;
  if (
    !line.items.every(
      ({ evidence }) =>
        onRun(evidence) &&
        Math.abs(evidence.height - first.height) <=
          PDFJS_SPLIT_SPACE_TOLERANCE,
    )
  ) {
    return null;
  }
  const added: ProjectedNativeItem[] = [];
  for (let index = 1; index < line.items.length; index += 1) {
    const before = line.items[index - 1];
    const after = line.items[index];
    for (const spacer of spacers) {
      const at = spacer.evidence.itemIndex;
      if (
        at > before.evidence.itemIndex &&
        at < after.evidence.itemIndex &&
        onRun(spacer.evidence) &&
        Math.abs(spacer.angle - line.angle) < UPRIGHT_ANGLE_EPSILON &&
        spacer.baselineInline >= before.baselineInline &&
        spacer.baselineInline <= after.baselineInline
      ) {
        added.push(spacer);
      }
    }
  }
  return added.length > 0
    ? (lineFromItems([...line.items, ...added], line.kind) ?? line)
    : line;
}

function horizontalOverlapRatio(
  left: NativeVisualLine,
  paragraph: MutableParagraph,
): number {
  const overlap =
    Math.min(left.inlineEnd, paragraph.inlineEnd) -
    Math.max(left.inlineStart, paragraph.inlineStart);
  const minimumWidth = Math.max(
    1,
    Math.min(
      left.inlineEnd - left.inlineStart,
      paragraph.inlineEnd - paragraph.inlineStart,
    ),
  );
  return Math.max(0, overlap) / minimumWidth;
}

function lineLooksLikeHeading(line: NativeVisualLine): boolean {
  return Array.from(line.text.trim()).length <= 24;
}

function lineIsPrivateUseOnly(line: NativeVisualLine): boolean {
  const text = line.text.replace(/\s+/gu, "");
  return text.length > 0 && /^[\u{E000}-\u{F8FF}]+$/u.test(text);
}

function lineLooksLikeListItem(line: NativeVisualLine): boolean {
  // A bracketed number opens an entry of a reference list ("[9] K. V.
  // Kokov, ..."), the way "9." opens a numbered item.
  // The editor's own bullets (●, ■, ◆, ❑, ✓, ➢ and their hollow forms) and
  // circled numbers open items too.
  return /^(?:\0\s*|[•◦▪‣·●○■□◆◇❑✓➢]\s*|\d+[.)]\s*|[①-⑳]\s*|\[\d{1,4}\]\s)/u.test(line.text.trimStart());
}

/** Where a list item's text starts after its marker: the item after the
 * marker's own (the marker and the blanks after it read as items of their
 * own), or past the marker and its blank where the two share one item. */
function listItemTextStart(line: NativeVisualLine): number | null {
  let markerSeen = false;
  for (const item of line.items) {
    const text = item.evidence.str;
    if (!text.trim()) continue;
    if (!markerSeen && /^(?:\0|[•◦▪‣·●○■□◆◇❑✓➢]|\d+[.)]|[①-⑳]|\[\d{1,4}\])$/u.test(text.trim())) {
      markerSeen = true;
      continue;
    }
    if (!markerSeen) {
      // A marker drawn in one show with its words (as a bullet put on in
      // the editor is): the words start after it and the blank.
      const opening = text.length - text.trimStart().length;
      const marker = pdfTextListMarkerOf(text.slice(opening));
      if (!marker) return null;
      let words = opening + marker.length;
      while (words < text.length && /\s/u.test(text[words]!)) words += 1;
      if (words >= text.length) return null;
      const share = words / Math.max(1, text.length);
      return item.inlineStart + (item.inlineEnd - item.inlineStart) * share;
    }
    // The text item may start with the blank after the marker.
    const lead = text.length - text.trimStart().length;
    if (!lead) return item.inlineStart;
    const share = lead / Math.max(1, text.length);
    return item.inlineStart + (item.inlineEnd - item.inlineStart) * share;
  }
  return null;
}

function lineLooksLikeCompactNumericChartLabel(
  line: NativeVisualLine,
): boolean {
  return (
    line.items.length === 1 &&
    line.items[0].evidence.hasEOL &&
    COMPACT_NUMERIC_CHART_LABEL.test(line.text.trim())
  );
}

function lineLooksLikeSectionHeading(line: NativeVisualLine): boolean {
  const text = line.text.trim();
  const punctuationText = text.replace(/["'”’）》】」』]+$/u, "");
  return (
    Array.from(text).length <= 8 &&
    !/[。！？.!?；;，,、]$/u.test(punctuationText) &&
    !/^\d+[.)]\s+/u.test(text)
  );
}

function isExactVerticalColumnFrame(
  lines: readonly NativeVisualLine[],
): boolean {
  if (lines.length < 2 || lines.some((line) => line.items.length !== 1)) {
    return false;
  }
  const projected = lines.map((line) => line.items[0]);
  const first = projected[0];
  const firstTransform = first.evidence.transform;
  const xScale = Math.hypot(firstTransform[0], firstTransform[1]);
  const yScale = Math.hypot(firstTransform[2], firstTransform[3]);
  if (
    !projected.every(
      (item, index) =>
        item.vertical &&
        item.visualFontKey === first.visualFontKey &&
        item.evidence.fontName === first.evidence.fontName &&
        [0, 1, 2, 3].every(
          (component) =>
            Math.abs(
              item.evidence.transform[component] - firstTransform[component],
            ) <= 0.001,
        ) &&
        (index === projected.length - 1 || item.evidence.hasEOL) &&
        (index === 0 ||
          (item.evidence.itemIndex > projected[index - 1].evidence.itemIndex &&
            item.evidence.itemIndex -
              projected[index - 1].evidence.itemIndex <=
              PDF_NATIVE_MAX_FIXED_FRAME_ITEM_INDEX_GAP)),
    ) ||
    !Number.isFinite(xScale) ||
    !Number.isFinite(yScale) ||
    xScale <= 0.000001 ||
    yScale <= 0.000001 ||
    firstTransform[0] * firstTransform[3] -
      firstTransform[1] * firstTransform[2] <=
      0
  ) {
    return false;
  }
  const columnSteps = projected.slice(1).map((item, index) => {
    const previous = projected[index].evidence.transform;
    const current = item.evidence.transform;
    const deltaX = current[4] - previous[4];
    const deltaY = current[5] - previous[5];
    return Object.freeze({
      column:
        -(
          deltaX * firstTransform[0] + deltaY * firstTransform[1]
        ) / xScale,
      inline:
        (deltaX * firstTransform[2] + deltaY * firstTransform[3]) /
        yScale,
    });
  });
  const step = columnSteps[0]?.column ?? 0;
  const tolerance = Math.max(0.001, step * 0.005);
  return (
    step > 0.05 &&
    step <= Math.max(first.fontSize, 1) * 4 &&
    columnSteps.every(
      (candidate) =>
        Math.abs(candidate.column - step) <= tolerance &&
        Math.abs(candidate.inline) <= Math.max(first.fontSize, 1) * 0.05,
    )
  );
}

function isLikelySectionHeadingBoundary(
  line: NativeVisualLine,
  adjacent: NativeVisualLine,
): boolean {
  if (!lineLooksLikeSectionHeading(line)) return false;
  if (line.bodyFontSize / adjacent.bodyFontSize >= HEADING_FONT_RATIO) {
    return true;
  }
  const lineLength = Array.from(line.text.trim()).length;
  const adjacentLength = Array.from(adjacent.text.trim()).length;
  return lineLength <= 4 && adjacentLength >= lineLength * 2;
}

function assignColumns(
  lines: readonly NativeVisualLine[],
): readonly VisualColumn[] {
  const columns: VisualColumn[] = [];
  for (const line of [...lines].sort(
    (left, right) => left.inlineStart - right.inlineStart,
  )) {
    let best: VisualColumn | null = null;
    let bestDistance = Number.POSITIVE_INFINITY;
    for (const column of columns) {
      if (Math.abs(line.angle - column.angle) >= UPRIGHT_ANGLE_EPSILON) {
        continue;
      }
      const overlap =
        Math.min(line.inlineEnd, column.inlineEnd) -
        Math.max(line.inlineStart, column.inlineStart);
      const minimumWidth = Math.max(
        1,
        Math.min(
          line.inlineEnd - line.inlineStart,
          column.inlineEnd - column.inlineStart,
        ),
      );
      const leftDistance = Math.abs(line.inlineStart - column.inlineStart);
      const fontSize = Math.max(line.bodyFontSize, 1);
      if (overlap / minimumWidth < 0.35 && leftDistance > fontSize * 4) {
        continue;
      }
      if (leftDistance < bestDistance) {
        best = column;
        bestDistance = leftDistance;
      }
    }
    if (best) {
      best.lines.push(line);
      best.inlineStart = Math.min(best.inlineStart, line.inlineStart);
      best.inlineEnd = Math.max(best.inlineEnd, line.inlineEnd);
    } else {
      columns.push({
        lines: [line],
        inlineStart: line.inlineStart,
        inlineEnd: line.inlineEnd,
        angle: line.angle,
      });
    }
  }
  return Object.freeze(columns);
}

/** The widest line step read as prose set with open leading, in ems of its
 * size: double spacing (2.4) and a little over, below most paragraph gaps. */
const OPEN_LEADING_MAX_STEP_EM = 3;

/** The step a column's prose is set at: the baseline step that recurs
 * between consecutive lines of one size (at least twice), null when none
 * does. */
function columnLineStep(lines: readonly NativeVisualLine[]): number | null {
  const counts = new Map<number, number>();
  for (let index = 1; index < lines.length; index += 1) {
    const previous = lines[index - 1];
    const line = lines[index];
    const fontSize = Math.max(previous.bodyFontSize, line.bodyFontSize, 1);
    if (Math.abs(previous.bodyFontSize - line.bodyFontSize) > fontSize * 0.05) continue;
    const step = line.baselineBlock - previous.baselineBlock;
    if (step <= 0 || step > fontSize * OPEN_LEADING_MAX_STEP_EM) continue;
    const key = Math.round(step * 2) / 2;
    counts.set(key, (counts.get(key) ?? 0) + 1);
  }
  let best: number | null = null;
  let bestCount = 1;
  for (const [step, count] of counts) {
    if (count > bestCount) {
      best = step;
      bestCount = count;
    }
  }
  return best;
}

/** Prose set with open leading (more than the usual gap between lines, as a
 * report set at 1.7 lines): the next line of one size at the step the
 * paragraph (or, from its first line, the column) is set at, after a line
 * that runs most of the measure, continues it. WPS reads such lines as one
 * paragraph; each read as its own, a line typed longer could not wrap into
 * the next (bug 0066). Short lines at a steady step (a table's labels, a
 * list of entries) stay apart. */
function continuesOpenLeading(
  paragraph: MutableParagraph,
  line: NativeVisualLine,
  columnStep: number | null,
  /** The step from `line` to the column's line after it (null for none). */
  nextStep: number | null,
): boolean {
  const previous = paragraph.lines.at(-1);
  if (!previous) return false;
  const fontSize = Math.max(previous.bodyFontSize, line.bodyFontSize, 1);
  if (Math.abs(previous.bodyFontSize - line.bodyFontSize) > fontSize * 0.05) return false;
  const step = line.baselineBlock - previous.baselineBlock;
  if (step <= 0 || step > fontSize * OPEN_LEADING_MAX_STEP_EM) return false;
  const ownStep = paragraph.lines.length >= 2
    ? previous.baselineBlock - paragraph.lines[paragraph.lines.length - 2].baselineBlock
    : null;
  // From its first line, the paragraph is set at the column's step, or at
  // one of its own the line after this one keeps (a paragraph given its own
  // line spacing, as the editor sets one).
  const near = (left: number | null, right: number) => left !== null && Math.abs(left - right) <= fontSize * 0.15;
  const expected = ownStep ?? (near(columnStep, step) ? columnStep : near(nextStep, step) ? step : null);
  if (expected === null || Math.abs(step - expected) > fontSize * 0.15) return false;
  const width = (entry: NativeVisualLine) => entry.inlineEnd - entry.inlineStart;
  const measure = Math.max(width(line), ...paragraph.lines.map(width));
  return width(previous) >= fontSize * 10 && width(previous) >= measure * 0.6;
}

/** A row's extent and baseline as its body text sets them: the items at
 * the row's body size (a part set larger or smaller on it left out). */
function lineBodyExtent(line: NativeVisualLine): Readonly<{ blockStart: number; blockEnd: number; baselineBlock: number }> {
  const body = line.items.filter((item) => Math.abs(item.fontSize - line.bodyFontSize) <= line.bodyFontSize * 0.05);
  if (!body.length || body.length === line.items.length) return line;
  return {
    blockStart: Math.min(...body.map((item) => item.blockStart)),
    blockEnd: Math.max(...body.map((item) => item.blockEnd)),
    baselineBlock: body.reduce((total, item) => total + item.baselineBlock, 0) / body.length,
  };
}

function canAppendLine(
  paragraph: MutableParagraph,
  line: NativeVisualLine,
  columnStep: number | null = null,
  nextStep: number | null = null,
): boolean {
  const previous = paragraph.lines.at(-1);
  if (!previous) return false;
  if (
    previous.items.some((item) => item.vertical) ||
    line.items.some((item) => item.vertical)
  ) {
    return isExactVerticalColumnFrame([...paragraph.lines, line]);
  }
  if (previous.kind === "drop-cap" || line.kind === "drop-cap") return false;
  // A line of nothing but private-use glyphs is a drawing piece, not prose:
  // TeX builds a tall bracket from extension pieces stacked closer than a
  // line apart, and joining them made a "paragraph" no rewrite can set.
  if (lineIsPrivateUseOnly(previous) || lineIsPrivateUseOnly(line)) return false;
  // (As is a radical sign left off the line it stands on.)
  if (lineIsStrayRadical(previous) || lineIsStrayRadical(line)) return false;
  if (Math.abs(previous.angle - line.angle) >= UPRIGHT_ANGLE_EPSILON) {
    return false;
  }
  const fontSize = Math.max(previous.bodyFontSize, line.bodyFontSize, 1);
  const gap = line.blockStart - previous.blockEnd;
  // Two rows are told apart by their body text: a part set larger and
  // raised on a row (a superscript the editor enlarged) reaches into the
  // row above by its own height, and read by it the paragraph's last row
  // split off as a box of its own, its words then typed apart from the
  
  const previousBody = lineBodyExtent(previous);
  const lineBody = lineBodyExtent(line);
  if (
    lineBody.blockStart - previousBody.blockEnd < -fontSize * 0.8 ||
    (gap > fontSize * PARAGRAPH_GAP_FACTOR && !continuesOpenLeading(paragraph, line, columnStep, nextStep))
  ) {
    return false;
  }
  // Two blocks set over each other (a producer that ran one paragraph into
  // the next, half a line apart): a line of prose whose baseline stands much
  // closer under the previous one than the paragraph's own line step, or
  // than any step prose is set at, is the next block's first line. WPS keeps
  // the two blocks apart; joined, editing either re-set both as one and the
  // second lost its place.
  const baselineStep = lineBody.baselineBlock - previousBody.baselineBlock;
  if (line.inlineEnd - line.inlineStart > fontSize * 2 && baselineStep > 0) {
    // The paragraph's own step: its closest, a row set lower for a larger
    // part of it standing further from the row above.
    const ownSteps = paragraph.lines.slice(1).map((entry, index) =>
      lineBodyExtent(entry).baselineBlock - lineBodyExtent(paragraph.lines[index]).baselineBlock)
      .filter((step) => step > 0);
    const ownStep = ownSteps.length ? Math.min(...ownSteps) : null;
    // Measured by the paragraph's own step where it has shown one: a new
    // last row of nothing but an enlarged part (a superscript set larger,
    // typed after the one before it) is no closer than the paragraph's step
    // to the row above, but its size made it look half a line close, and it
    
    if (ownStep !== null && ownStep > 0 ? baselineStep < ownStep * 0.75 : baselineStep < fontSize * 0.75) {
      return false;
    }
  }
  if (
    previous.bodyFontSize / line.bodyFontSize >= HEADING_FONT_RATIO &&
    lineLooksLikeHeading(previous)
  ) {
    return false;
  }
  const previousIsListItem = lineLooksLikeListItem(previous);
  const nextIsListItem = lineLooksLikeListItem(line);
  // An item's wrapped line, set at the item's text after its marker (a
  // hanging indent), continues the item: split from it, the item typed on
  // wrapped over its own second line (WHO sitrep p2, the side browser walk
  
  const itemTextStart = previousIsListItem && !nextIsListItem && paragraph.lines.length === 1
    ? listItemTextStart(previous) : null;
  const continuesItem = itemTextStart !== null && Math.abs(line.inlineStart - itemTextStart) <= fontSize * 0.5;
  if (previousIsListItem !== nextIsListItem && !continuesItem) {
    return false;
  }
  // Two marked lines in a row at different edges are two items of two
  // levels (a nested list): a bullet never continues the item before it
  // across a change of level, and the rewrite sets one frame with one edge.
  // Items at one edge keep grouping as they did.
  if (
    previousIsListItem &&
    nextIsListItem &&
    Math.abs(line.inlineStart - previous.inlineStart) > fontSize * 0.5
  ) {
    return false;
  }
  // A line set in from the paragraph's body edge by a first-line indent
  // (half an em to four) opens the next paragraph: producers set every
  // paragraph's first line in, and the reader's grouping would otherwise
  // run a page of prose into one frame whose lines start at more than one
  // left edge, which the rewrite refuses. The body edge is where the lines
  // after the first stand; a paragraph of one line so far is measured from
  // that line. A hanging indent runs the other way (its later lines stand
  // in, its first out) and is left whole, as is a short fragment such as a
  // superscript the reader lifted onto a line of its own.
  const bodyEdge =
    paragraph.lines.length >= 2
      ? Math.min(...paragraph.lines.slice(1).map((entry) => entry.inlineStart))
      : previous.inlineStart;
  const firstLineIndent = line.inlineStart - bodyEdge;
  // After one line only, a line set in may as well be the wrapped
  // continuation of a hanging indent whose marker the reader did not name;
  // it is read as one when the line before it filled the measure (it runs
  // at least as far as the new line, less two ems).
  const wrapsHangingIndent =
    paragraph.lines.length === 1 &&
    previous.inlineEnd >= line.inlineEnd - fontSize * 2;
  if (
    !previousIsListItem &&
    !wrapsHangingIndent &&
    firstLineIndent >= fontSize * 0.5 &&
    firstLineIndent <= fontSize * 4 &&
    line.inlineEnd - line.inlineStart > fontSize * 2
  ) {
    return false;
  }
  // A line that starts more than four ems in from the body edge is not the
  // paragraph's continuation: a displayed equation or a centred line set
  // between two paragraphs, which the rewrite cannot set in their frame.
  if (
    paragraph.lines.length >= 2 &&
    firstLineIndent > fontSize * 4 &&
    line.inlineEnd - line.inlineStart > fontSize * 2
  ) {
    return false;
  }
  // Nor is a line that starts more than four ems out from it: the prose
  // that resumes at the body edge after such a displayed line.
  if (firstLineIndent < -fontSize * 4) {
    return false;
  }
  // A hanging indent repeats where the next item starts: the paragraph's
  // first line stands out from the body edge, and a later line standing out
  // by the same amount opens the next item, marker glyph or not.
  const firstLine = paragraph.lines[0];
  const hang = bodyEdge - firstLine.inlineStart;
  if (
    paragraph.lines.length >= 2 &&
    hang >= fontSize * 0.5 &&
    Math.abs(bodyEdge - line.inlineStart - hang) <= fontSize * 0.25
  ) {
    return false;
  }
  // Space added between two lines after one that stops short of the
  // measure closes the paragraph: a form sets its questions a quarter line
  // apart and ends each on a short line, and a writer that spaces its
  // paragraphs does the same. Only a paragraph that has shown its own line
  // step (two lines) is measured, so a lone line is left to the rules above.
  if (paragraph.lines.length >= 2) {
    const beforePrevious = paragraph.lines[paragraph.lines.length - 2];
    const step = previous.baselineBlock - beforePrevious.baselineBlock;
    const nextStep = line.baselineBlock - previous.baselineBlock;
    const measureEnd = Math.max(...paragraph.lines.map((entry) => entry.inlineEnd));
    if (
      step > 0 &&
      nextStep > step * 1.25 &&
      nextStep - step >= fontSize * 0.2 &&
      previous.inlineEnd < measureEnd - fontSize * 2
    ) {
      return false;
    }
  }
  // The short last row of a wrapped paragraph (a title re-set in a narrow
  // frame ends on its last character) is no heading: it stands at the body
  // edge, one step under a line that filled the measure, in the same size.
  // Read as a heading it opened as a box of its own, apart from its
  
  const wrapTail = (() => {
    if (paragraph.lines.length < 2) return false;
    const beforePrevious = paragraph.lines[paragraph.lines.length - 2];
    const ownStep = previous.baselineBlock - beforePrevious.baselineBlock;
    const measureEnd = Math.max(...paragraph.lines.map((entry) => entry.inlineEnd));
    return ownStep > 0 && Math.abs(baselineStep - ownStep) <= fontSize * 0.15 &&
      Math.abs(line.inlineStart - bodyEdge) <= fontSize * 0.25 &&
      previous.inlineEnd >= measureEnd - fontSize * 1.5 &&
      Math.abs(line.bodyFontSize - previous.bodyFontSize) <= fontSize * 0.05;
  })();
  if (
    !wrapTail &&
    (isLikelySectionHeadingBoundary(previous, line) ||
      isLikelySectionHeadingBoundary(line, previous))
  ) {
    return false;
  }
  // (A heading never stands into the row above it: a larger row whose text
  // reaches into the body of the one above is the paragraph's own, a part
  // set larger on its last row.)
  if (
    line.bodyFontSize / previous.bodyFontSize >= HEADING_FONT_RATIO &&
    lineLooksLikeHeading(line) &&
    !(lineBody.blockStart < previousBody.blockEnd)
  ) {
    return false;
  }
  // A short lone line set smaller than the body under it is its heading
  // still (walk s7: "Abstract" made a 9.33 pt subscript over 11.96 pt
  // body read as the paragraph's first line, and every later click on it
  // opened the whole paragraph).
  if (
    paragraph.lines.length === 1 &&
    line.bodyFontSize / previous.bodyFontSize >= HEADING_FONT_RATIO &&
    lineLooksLikeHeading(previous)
  ) {
    return false;
  }
  const leftDelta = Math.abs(line.inlineStart - previous.inlineStart);
  if (
    leftDelta > fontSize * 0.5 &&
    lineLooksLikeCompactNumericChartLabel(previous) &&
    lineLooksLikeCompactNumericChartLabel(line)
  ) {
    return false;
  }
  return (
    leftDelta <= fontSize * 2.5 ||
    horizontalOverlapRatio(line, paragraph) >= 0.55
  );
}

function visualParagraphs(
  lines: readonly NativeVisualLine[],
): readonly MutableParagraph[] {
  const paragraphs: MutableParagraph[] = [];
  for (const column of assignColumns(lines)) {
    const columnParagraphs: MutableParagraph[] = [];
    // (In their body text's order, as visualLines sorts them.)
    const columnLines = [...column.lines].sort(
      (left, right) => lineBodyExtent(left).blockStart - lineBodyExtent(right).blockStart,
    );
    const columnStep = columnLineStep(columnLines);
    for (const [lineIndex, line] of columnLines.entries()) {
      const following = columnLines[lineIndex + 1];
      const nextStep = following && Math.abs(following.bodyFontSize - line.bodyFontSize) <= Math.max(line.bodyFontSize, 1) * 0.05
        ? lineBodyExtent(following).baselineBlock - lineBodyExtent(line).baselineBlock : null;
      let paragraph: MutableParagraph | undefined;
      for (let index = columnParagraphs.length - 1; index >= 0; index -= 1) {
        const candidate = columnParagraphs[index];
        if (canAppendLine(candidate, line, columnStep, nextStep)) {
          paragraph = candidate;
          break;
        }
      }
      if (paragraph) {
        paragraph.lines.push(line);
        paragraph.left = Math.min(paragraph.left, line.left);
        paragraph.right = Math.max(paragraph.right, line.right);
        paragraph.top = Math.min(paragraph.top, line.top);
        paragraph.bottom = Math.max(paragraph.bottom, line.bottom);
        paragraph.inlineStart = Math.min(
          paragraph.inlineStart,
          line.inlineStart,
        );
        paragraph.inlineEnd = Math.max(paragraph.inlineEnd, line.inlineEnd);
        paragraph.blockStart = Math.min(paragraph.blockStart, line.blockStart);
        paragraph.blockEnd = Math.max(paragraph.blockEnd, line.blockEnd);
      } else {
        columnParagraphs.push({
          lines: [line],
          left: line.left,
          right: line.right,
          top: line.top,
          bottom: line.bottom,
          inlineStart: line.inlineStart,
          inlineEnd: line.inlineEnd,
          blockStart: line.blockStart,
          blockEnd: line.blockEnd,
        });
      }
    }
    paragraphs.push(...columnParagraphs);
  }
  return Object.freeze(paragraphs);
}

function sourceStepIsContiguous(
  previous: ProjectedNativeItem,
  next: ProjectedNativeItem,
): boolean {
  const gap = next.evidence.itemIndex - previous.evidence.itemIndex;
  return gap >= 1 && gap <= PDF_NATIVE_MAX_FIXED_FRAME_ITEM_INDEX_GAP;
}

function paragraphOfItems(
  lines: readonly NativeVisualLine[],
  members: ReadonlySet<ProjectedNativeItem>,
): MutableParagraph | null {
  const owned = lines.flatMap((line) => {
    const items = line.items.filter((item) => members.has(item));
    if (items.length === line.items.length) return [line];
    const rebuilt = items.length > 0 ? lineFromItems(items, line.kind) : null;
    return rebuilt ? [rebuilt] : [];
  });
  if (owned.length === 0) return null;
  return {
    lines: owned,
    left: Math.min(...owned.map((line) => line.left)),
    right: Math.max(...owned.map((line) => line.right)),
    top: Math.min(...owned.map((line) => line.top)),
    bottom: Math.max(...owned.map((line) => line.bottom)),
    inlineStart: Math.min(...owned.map((line) => line.inlineStart)),
    inlineEnd: Math.max(...owned.map((line) => line.inlineEnd)),
    blockStart: Math.min(...owned.map((line) => line.blockStart)),
    blockEnd: Math.max(...owned.map((line) => line.blockEnd)),
  };
}

/**
 * The native writer owns a paragraph only as one run of source items that
 * reads in visual order (at most PDF_NATIVE_MAX_FIXED_FRAME_ITEM_INDEX_GAP
 * apart). Visual grouping cannot see source order, and real files draw what
 * reads as one paragraph out of that order: a cover title after the model
 * numbers beneath it, a full-width line after both columns above it, a side
 * heading after the column it sits beside, one cell per row of a contents
 * table. Refusing the whole paragraph leaves visible text that cannot be
 * opened, so it is offered as the source runs it is made of instead. In
 * reading order each item continues the run whose last item precedes it most
 * closely in the source. A run whose box would cover other text on the page
 * is cut where that text is read, or between the lines above and below it,
 * so a box made here never hides a box the reader can already open.
 */
function sourceContiguousParagraphs(
  paragraph: MutableParagraph,
  pageItems: readonly ProjectedNativeItem[],
): readonly MutableParagraph[] {
  const ordered = paragraph.lines.flatMap((line) => line.items);
  if (
    ordered.every(
      (item, index) =>
        index === 0 || sourceStepIsContiguous(ordered[index - 1], item),
    )
  ) {
    return [paragraph];
  }
  const readingIndex = new Map(ordered.map((item, index) => [item, index]));
  const lineOfItem = new Map(
    paragraph.lines.flatMap((line) =>
      line.items.map((item) => [item, line] as const),
    ),
  );
  const runs: ProjectedNativeItem[][] = [];
  for (const item of ordered) {
    let best: ProjectedNativeItem[] | undefined;
    for (const run of runs) {
      const last = run[run.length - 1];
      const line = lineOfItem.get(last);
      if (
        sourceStepIsContiguous(last, item) &&
        // A run that holds only part of a line must not reach the next one:
        // its box would span both lines' width and cover the rest of each.
        // A whole line may always carry on into the line below it.
        (line === lineOfItem.get(item) ||
          line?.items.every((entry) => run.includes(entry))) &&
        (!best ||
          last.evidence.itemIndex >
            best[best.length - 1].evidence.itemIndex)
      ) {
        best = run;
      }
    }
    if (best) best.push(item);
    else runs.push([item]);
  }
  const centre = (item: ProjectedNativeItem) =>
    [
      (item.inlineStart + item.inlineEnd) / 2,
      (item.blockStart + item.blockEnd) / 2,
    ] as const;
  const inside = (
    box: MutableParagraph | null,
    [inline, block]: readonly [number, number],
  ) =>
    box !== null &&
    inline > box.inlineStart &&
    inline < box.inlineEnd &&
    block > box.blockStart &&
    block < box.blockEnd;
  // Only a box of two or more items is mounted over the text layer, so only
  // such a run can hide other text; a cut is kept only when neither side
  // still covers that text.
  const covers = (
    run: readonly ProjectedNativeItem[],
    point: readonly [number, number],
  ) =>
    run.length > 1 &&
    inside(paragraphOfItems(paragraph.lines, new Set(run)), point);
  const exposed = pageItems.filter(
    (item) =>
      item.evidence.str.trim() &&
      Math.abs(item.angle - ordered[0].angle) < UPRIGHT_ANGLE_EPSILON &&
      inside(paragraph, centre(item)),
  );
  const uncoveringCut = (): readonly [number, number] | null => {
    for (const [runIndex, run] of runs.entries()) {
      if (run.length < 2) continue;
      const members = new Set(run);
      const box = paragraphOfItems(paragraph.lines, members);
      for (const item of exposed) {
        const point = centre(item);
        if (members.has(item) || !inside(box, point)) continue;
        const read = readingIndex.get(item);
        const cuts = [
          // Text read inside the run: cut where it is read.
          read === undefined
            ? -1
            : run.findIndex(
                (entry) => (readingIndex.get(entry) ?? -1) > read,
              ),
          // Text beside the run: part the lines below it from those above,
          // then the lines that reach its height from those above it.
          run.findIndex((entry) => entry.blockStart >= point[1]),
          run.findIndex((entry) => entry.blockEnd > point[1]),
        ];
        const cut = cuts.find(
          (at) =>
            at > 0 &&
            !covers(run.slice(0, at), point) &&
            !covers(run.slice(at), point),
        );
        if (cut !== undefined) return [runIndex, cut];
      }
    }
    return null;
  };
  for (let cut = uncoveringCut(); cut; cut = uncoveringCut()) {
    const [runIndex, at] = cut;
    const run = runs[runIndex];
    runs.splice(runIndex, 1, run.slice(0, at), run.slice(at));
  }
  return runs.flatMap((run) => {
    const owned = paragraphOfItems(paragraph.lines, new Set(run));
    return owned ? [owned] : [];
  });
}

/**
 * A paragraph's text is one field of the record that edits it, and bounded
 * (PDF_NATIVE_MAX_ITEM_TEXT_LENGTH). A paragraph past that bound -- a table
 * of contents read as one paragraph of seventy entries -- is cut at line
 * boundaries into paragraphs that fit, each edited on its own; text never
 * moves across a cut. A single line past the bound is left as it is.
 */
function paragraphsWithinTextLimit(
  paragraph: MutableParagraph,
): readonly MutableParagraph[] {
  const textOf = (lines: readonly NativeVisualLine[]) =>
    summarizePdfNativeSelectionItems(
      lines.flatMap((line) => line.items.map((item) => item.evidence)),
    ).length;
  if (textOf(paragraph.lines) <= PDF_NATIVE_MAX_ITEM_TEXT_LENGTH) {
    return [paragraph];
  }
  const chunks: NativeVisualLine[][] = [];
  let current: NativeVisualLine[] = [];
  for (const line of paragraph.lines) {
    if (
      current.length > 0 &&
      textOf([...current, line]) > PDF_NATIVE_MAX_ITEM_TEXT_LENGTH
    ) {
      chunks.push(current);
      current = [];
    }
    current.push(line);
  }
  if (current.length > 0) chunks.push(current);
  return chunks.flatMap((lines) => {
    const owned = paragraphOfItems(
      paragraph.lines,
      new Set(lines.flatMap((line) => line.items)),
    );
    return owned ? [owned] : [];
  });
}

function paragraphReason(
  lines: readonly NativeVisualLine[],
  angle: number,
  // The reader's pieces of one show on one baseline (singleRunLine): the
  // line is selected as that one run, so the paragraph geometry a turned
  // box of several runs must have is not asked of it.
  oneRun = false,
): PdfNativeParagraphEditBlockReason | undefined {
  const projectedItems = lines.flatMap((line) => line.items);
  const items = projectedItems.map((item) => item.evidence);
  const verticalColumnFrame = isExactVerticalColumnFrame(lines);
  const affineParagraphCandidate = (): boolean => {
    if (projectedItems.some((item) => item.vertical)) return false;
    const first = items[0]?.transform;
    if (!first) return false;
    const xScale = Math.hypot(first[0], first[1]);
    const yScale = Math.hypot(first[2], first[3]);
    if (
      xScale <= 0.000001 ||
      yScale <= 0.000001 ||
      Math.abs(xScale - yScale) > 0.001 ||
      Math.abs(first[0] * first[2] + first[1] * first[3]) > 0.001 ||
      first[0] * first[3] - first[1] * first[2] <= 0
    ) {
      return false;
    }
    const normalized = [
      first[0] / xScale,
      first[1] / xScale,
      first[2] / yScale,
      first[3] / yScale,
    ];
    if (
      items.some((item) => {
        const x = Math.hypot(item.transform[0], item.transform[1]);
        const y = Math.hypot(item.transform[2], item.transform[3]);
        return (
          x <= 0.000001 ||
          y <= 0.000001 ||
          [0, 1, 2, 3].some(
            (index) =>
              Math.abs(
                item.transform[index] / (index < 2 ? x : y) -
                  normalized[index],
              ) > 0.001,
          )
        );
      })
    ) {
      return false;
    }
    return lines.slice(1).every((line, index) => {
      const prior = lines[index].items[0]?.evidence;
      const next = line.items[0]?.evidence;
      if (!prior || !next) return false;
      const deltaX = next.transform[4] - prior.transform[4];
      const deltaY = next.transform[5] - prior.transform[5];
      const inlineDelta =
        deltaX * normalized[0] + deltaY * normalized[1];
      const blockDelta =
        deltaX * normalized[2] + deltaY * normalized[3];
      return (
        Math.abs(inlineDelta) <= Math.max(line.bodyFontSize, 1) * 0.75 &&
        blockDelta < -0.05
      );
    });
  };
  if (
    Math.abs(angle) >= UPRIGHT_ANGLE_EPSILON &&
    projectedItems.length > 1 &&
    !oneRun &&
    !verticalColumnFrame &&
    !affineParagraphCandidate()
  ) {
    return "unsupported-geometry";
  }
  if (items.length > PDF_NATIVE_MAX_FIXED_FRAME_ITEMS) {
    return "too-many-items";
  }
  if (items.some((item) => UNNAMED_CODE_CHARACTER.test(item.str))) {
    return "not-text";
  }
  if (
    items.some((item, index) => {
      if (index === 0) return false;
      const gap = item.itemIndex - items[index - 1].itemIndex;
      return gap < 1 || gap > PDF_NATIVE_MAX_FIXED_FRAME_ITEM_INDEX_GAP;
    })
  ) {
    return "noncontiguous-items";
  }
  if (projectedItems.some((item) => !item.visualFontKey)) {
    return "mixed-font-items";
  }
  if (
    lines.slice(0, -1).some((line, lineIndex) => {
      const last = line.items.at(-1)?.evidence;
      const firstNext = lines[lineIndex + 1]?.items[0]?.evidence;
      return (
        !last ||
        !firstNext ||
        firstNext.itemIndex <= last.itemIndex ||
        firstNext.itemIndex - last.itemIndex >
          PDF_NATIVE_MAX_FIXED_FRAME_ITEM_INDEX_GAP
      );
    })
  ) {
    return "incomplete-line-boundaries";
  }
  return undefined;
}

/** The key prefix of an item the writer's box mark holds
 * (pdfNativeTextHardGroups). */
const BOX_MARK_GROUP_PREFIX = "box:";

/**
 * The box key of a projected item (pdfNativeTextHardGroups, by raw item
 * index): a split space folded into the item it prefixes is that item's,
 * a space folded after a marker the marker's.
 */
function hardGroupOf(
  hardGroups: readonly (string | null)[] | null | undefined,
): (item: ProjectedNativeItem) => string | null {
  if (!hardGroups) return () => null;
  return (item) => {
    const sources = rawSourceItemIndexes(item.evidence);
    const visible = /^\s/u.test(item.evidence.str) ? sources.at(-1) : sources[0];
    return hardGroups[visible ?? item.evidence.itemIndex] ?? null;
  };
}












function markedBoxParagraph(lines: readonly NativeVisualLine[]): readonly MutableParagraph[] | null {
  if (lines.length === 0) return null;
  const angle = lines[0].angle;
  if (lines.some((line) => Math.abs(line.angle - angle) >= UPRIGHT_ANGLE_EPSILON ||
    line.items.some((item) => item.vertical))) return null;
  const ordered = [...lines].sort((left, right) =>
    lineBodyExtent(left).baselineBlock - lineBodyExtent(right).baselineBlock ||
    left.inlineStart - right.inlineStart);
  const rows: NativeVisualLine[] = [];
  let row: NativeVisualLine[] = [];
  const flush = () => {
    if (row.length === 0) return;
    const joined = row.length === 1 ? row[0] : lineFromItems(row.flatMap((line) => line.items), row[0].kind);
    if (joined) rows.push(joined);
    row = [];
  };
  for (const line of ordered) {
    const previous = row.at(-1);
    if (previous && Math.abs(lineBodyExtent(line).baselineBlock - lineBodyExtent(previous).baselineBlock) >
      Math.max(line.bodyFontSize, previous.bodyFontSize, 1) * 0.3) flush();
    row.push(line);
  }
  flush();
  if (rows.length === 0) return null;
  return [{
    lines: rows,
    left: Math.min(...rows.map((line) => line.left)),
    right: Math.max(...rows.map((line) => line.right)),
    top: Math.min(...rows.map((line) => line.top)),
    bottom: Math.max(...rows.map((line) => line.bottom)),
    inlineStart: Math.min(...rows.map((line) => line.inlineStart)),
    inlineEnd: Math.max(...rows.map((line) => line.inlineEnd)),
    blockStart: Math.min(...rows.map((line) => line.blockStart)),
    blockEnd: Math.max(...rows.map((line) => line.blockEnd)),
  }];
}

/** A box's rows in the order the page draws its items: a row ends where
 * the reader ended its line, or where the next item starts back at the left
 * and lower down (a line set under it). Within a row, a sum's limits and a
 * script stay with the line they are set in. */
function sourceOrderRows(items: readonly ProjectedNativeItem[]): NativeVisualLine[] | null {
  const ordered = [...items].sort((left, right) => left.evidence.itemIndex - right.evidence.itemIndex);
  const rows: ProjectedNativeItem[][] = [];
  for (const item of ordered) {
    const row = rows.at(-1);
    const previous = row?.at(-1);
    const size = Math.max(item.fontSize, previous?.fontSize ?? 0, 1);
    // Measured from the row's body (its largest text), not from a limit or
    // a script set below it.
    const body = row ? Math.max(...row.map((entry) => entry.fontSize)) : 0;
    const bodyBaseline = row ? Math.max(...row.filter((entry) => entry.fontSize >= body * 0.95)
      .map((entry) => entry.baselineBlock)) : 0;
    if (!row || !previous || previous.evidence.hasEOL ||
      (item.inlineStart < previous.inlineStart - size * 0.5 && item.baselineBlock - bodyBaseline > size * 0.5)) {
      rows.push([item]);
    } else {
      row.push(item);
    }
  }
  const lines = rows.map((row) => lineFromItems(row));
  return lines.every((line): line is NativeVisualLine => line !== null) ? lines : null;
}

/** Whether a paragraph reads its items in the order the page draws them. */
function readsInSourceOrder(paragraph: MutableParagraph): boolean {
  const order = paragraph.lines.flatMap((line) => line.items.map((item) => item.evidence.itemIndex));
  return order.every((index, at) => at === 0 || index > order[at - 1]);
}

/** The key prefix of the items of a centred stack (centredStacks). */
const CENTRED_STACK_GROUP_PREFIX = "stack:";
// How closely the lines of a centred label share their axis: a line centred
// with a trailing space stands half a space off the others.
const CENTRED_STACK_AXIS_EM = 0.2;
const CENTRED_STACK_STEP_TOLERANCE_EM = 0.15;
// Lines whose left edges differ by less are set flush.
const CENTRED_STACK_MIN_INSET_EM = 0.1;
// How far over and under a stack, in its own steps, text that is no stack
// of its own must keep for the stack to stand alone.
const CENTRED_STACK_REACH_STEPS = 1.5;






































function hangsAfterStackLine(line: ProjectedNativeItem, item: ProjectedNativeItem): boolean {
  const size = Math.max(line.fontSize, 1);
  return !line.vertical && !item.vertical && Math.abs(line.angle - item.angle) < UPRIGHT_ANGLE_EPSILON &&
    item.fontSize <= size * 0.85 &&
    Math.abs(item.baselineBlock - line.baselineBlock) <= size * 0.6 &&
    Math.abs(item.baselineBlock - line.baselineBlock) >= size * 0.1 &&
    item.inlineStart >= line.inlineEnd - size * 0.1 && item.inlineStart - line.inlineEnd <= size * 0.3 &&
    item.inlineEnd - item.inlineStart <= size * 1.5;
}

function centredStacks(
  items: readonly ProjectedNativeItem[],
  page: readonly ProjectedNativeItem[],
): ProjectedNativeItem[][] {
  const ordered = items.filter((item) => item.evidence.str.trim() !== "")
    .sort((left, right) => left.evidence.itemIndex - right.evidence.itemIndex);
  const centre = (item: ProjectedNativeItem) => (item.inlineStart + item.inlineEnd) / 2;
  const sameAngle = (left: ProjectedNativeItem, right: ProjectedNativeItem) =>
    !left.vertical && !right.vertical && Math.abs(left.angle - right.angle) < UPRIGHT_ANGLE_EPSILON;
  // The next line down, as close as a paragraph keeps its lines.
  const stepsUnder = (above: ProjectedNativeItem, below: ProjectedNativeItem) => {
    const size = Math.max(above.fontSize, below.fontSize, 1);
    return sameAngle(above, below) &&
      Math.abs(above.fontSize - below.fontSize) <= size * 0.05 &&
      below.baselineBlock - above.baselineBlock >= size * 0.75 &&
      below.blockStart - above.blockEnd <= size * PARAGRAPH_GAP_FACTOR;
  };
  type Run = { members: ProjectedNativeItem[]; step: number; centred: boolean; stands: boolean };
  const runs: Run[] = [];
  for (let at = 0; at < ordered.length;) {
    const first = ordered[at];
    const size = Math.max(first.fontSize, 1);
    const members = [first];
    // The stack's lines; its members are those and the marks hanging after them.
    const lines = [first];
    let low = centre(first), high = low, step: number | null = null;
    for (let next = at + 1; next < ordered.length; next += 1) {
      const previous = lines[lines.length - 1];
      const item = ordered[next];
      if (hangsAfterStackLine(previous, item)) {
        members.push(item);
        continue;
      }
      const middle = centre(item);
      const ownStep = item.baselineBlock - previous.baselineBlock;
      if (
        !stepsUnder(previous, item) ||
        Math.max(high, middle) - Math.min(low, middle) > size * CENTRED_STACK_AXIS_EM ||
        (step !== null && Math.abs(ownStep - step) > size * CENTRED_STACK_STEP_TOLERANCE_EM)
      ) {
        break;
      }
      step ??= ownStep;
      low = Math.min(low, middle);
      high = Math.max(high, middle);
      members.push(item);
      lines.push(item);
    }
    // (A mark read after the stack's last line is that line's only when the
    // stack is one: a run that ends in marks alone gives them back.)
    const run = step !== null &&
      !lines.every((item) => COMPACT_NUMERIC_CHART_LABEL.test(item.evidence.str.trim()));
    if (run) {
      const starts = lines.map((item) => item.inlineStart);
      runs.push({ members, step: step!, stands: true,
        centred: Math.max(...starts) - Math.min(...starts) >= size * CENTRED_STACK_MIN_INSET_EM });
    }
    at += run ? members.length : 1;
  }
  if (runs.length === 0) return [];
  const runOf = new Map(runs.flatMap((run) => run.members.map((member) => [member, run] as const)));
  const grouped = new Set(items);
  const others = page.filter((item) => item.evidence.str.trim() !== "");
  const stands = (run: Run) => {
    const { members, step } = run;
    const first = members[0], last = members[members.length - 1];
    const start = Math.min(...members.map((item) => item.inlineStart));
    const end = Math.max(...members.map((item) => item.inlineEnd));
    return !others.some((other) => {
      const its = runOf.get(other);
      if (its === run || !sameAngle(other, first)) return false;
      const em = Math.max(first.fontSize, other.fontSize, 1);
      const beside = Math.min(...members
        .filter((member) => Math.abs(other.baselineBlock - member.baselineBlock) < em * 0.8)
        .map((member) => Math.max(member.inlineStart - other.inlineEnd, other.inlineStart - member.inlineEnd, 0)));
      // (A line of another centred run may stand that near: three labels of
      // two rows each, their second rows a third of an em apart, are three
      // labels -- IPCC SPM's "Limited / Incomplete / Proactive adaptation".)
      if (beside <= em * PDFJS_SPACE_IN_FLOW_MAX_EM && !its?.centred) return true;
      // (Text of a box the file names is no line of this one's.)
      if (!grouped.has(other)) return false;
      if (beside <= em * SAME_LINE_GAP_FACTOR && !its?.centred) return true;
      if ((its?.centred && its.stands) || Math.min(other.inlineEnd, end) - Math.max(other.inlineStart, start) <= 0) return false;
      const over = first.baselineBlock - other.baselineBlock;
      const under = other.baselineBlock - last.baselineBlock;
      return (over >= em * 0.25 && over <= step * CENTRED_STACK_REACH_STEPS) ||
        (under >= em * 0.25 && under <= step * CENTRED_STACK_REACH_STEPS);
    });
  };
  for (let changed = true; changed;) {
    changed = false;
    for (const run of runs) {
      if (run.stands && !stands(run)) {
        run.stands = false;
        changed = true;
      }
    }
  }
  return runs.filter((run) => run.stands).map((run) => run.members);
}

/** Where the pinned boxes of a page start (a kept block's `origin`), for
 * normalizePdfJsNativeParagraphItems. */
export function pdfNativeBoxOriginsOf(
  keptBlocks: readonly PdfNativeKeptTextBlock[] | undefined,
  pageNumber: number,
): readonly Readonly<{ x: number; y: number }>[] {
  return (keptBlocks ?? []).flatMap((block) => block.pageNumber === pageNumber && block.origin ? [block.origin] : []);
}

/** How far, in PDF units, a kept block's first item may stand from the
 * origin its record names (a move is written in micro-units). */
const KEPT_BLOCK_ORIGIN_TOLERANCE = 1;

/** The items that draw a kept block's rows: one run of source items, in
 * order, row by row, each row on one baseline and below the one before it,
 * starting where the block's origin says when it names one. Null when the
 * page no longer draws the block so (rewritten since, undone, moved again,
 * or set otherwise). */
function keptBlockItems(
  rows: readonly string[],
  ordered: readonly ProjectedNativeItem[],
  claimed: ReadonlySet<ProjectedNativeItem>,
  origin?: Readonly<{ x: number; y: number }>,
): ProjectedNativeItem[][] | null {
  const flat = (text: string) => text.replace(/\s+/gu, "");
  const wanted = rows.map(flat);
  const first = wanted[0];
  for (let start = 0; start < ordered.length; start += 1) {
    if (!first.startsWith(flat(ordered[start].evidence.str)) || !flat(ordered[start].evidence.str)) continue;
    if (
      origin &&
      (Math.abs(ordered[start].evidence.transform[4] - origin.x) > KEPT_BLOCK_ORIGIN_TOLERANCE ||
        Math.abs(ordered[start].evidence.transform[5] - origin.y) > KEPT_BLOCK_ORIGIN_TOLERANCE)
    ) {
      continue;
    }
    const members: ProjectedNativeItem[] = [];
    const memberRows: ProjectedNativeItem[][] = [];
    let index = start;
    let previousBaseline: number | null = null;
    let matched = true;
    for (const row of wanted) {
      const rowItems: ProjectedNativeItem[] = [];
      let text = "";
      while (index < ordered.length && text.length < row.length) {
        const item = ordered[index];
        const previous = rowItems.at(-1) ?? members.at(-1);
        if (claimed.has(item) || (previous && !sourceStepIsContiguous(previous, item))) break;
        const piece = flat(item.evidence.str);
        if (!row.startsWith(text + piece)) break;
        text += piece;
        rowItems.push(item);
        index += 1;
      }
      const baseline = rowItems[0]?.baselineBlock;
      // A row's scripts stand off its baseline by up to half its body's
      // size: measured by their own smaller size, a superscript typed on
      // left the row unmatched, the block was read as several boxes and the
      
      const rowSize = Math.max(0, ...rowItems.map((item) => item.fontSize));
      if (
        text !== row ||
        baseline === undefined ||
        rowItems.some((item) => Math.abs(item.baselineBlock - baseline) > rowSize * 0.6) ||
        (previousBaseline !== null && !(baseline - previousBaseline > rowItems[0].fontSize / 4))
      ) {
        matched = false;
        break;
      }
      previousBaseline = baseline;
      members.push(...rowItems);
      memberRows.push(rowItems);
    }
    if (matched) return memberRows;
  }
  return null;
}

/**
 * The items a pinned box drew (native-page-box-map.ts): from the item at its
 * origin, the run of source items after it, none another box holds, whose
 * text is the box's, blanks aside. Its rows are not asked to stand one under
 * another: a box the reader once read so is that box however its rows lie (a
 * formula's pieces drawn out of order, a name block justified apart).
 */
function pinnedBoxItems(
  rows: readonly string[],
  ordered: readonly ProjectedNativeItem[],
  claimed: ReadonlySet<ProjectedNativeItem>,
  origin: Readonly<{ x: number; y: number }>,
): ProjectedNativeItem[] | null {
  const flat = (text: string) => text.replace(/\s+/gu, "");
  const wanted = flat(rows.join(""));
  if (!wanted) return null;
  for (let start = 0; start < ordered.length; start += 1) {
    const head = ordered[start];
    const piece = flat(head.evidence.str);
    if (!piece || claimed.has(head) || !wanted.startsWith(piece) ||
      Math.abs(head.evidence.transform[4] - origin.x) > PINNED_BOX_ORIGIN_TOLERANCE ||
      Math.abs(head.evidence.transform[5] - origin.y) > PINNED_BOX_ORIGIN_TOLERANCE) continue;
    const members = [head];
    let text = piece;
    let blanks: ProjectedNativeItem[] = [];
    for (let index = start + 1; index < ordered.length && text.length < wanted.length; index += 1) {
      const item = ordered[index];
      if (claimed.has(item) || !sourceStepIsContiguous(blanks.at(-1) ?? members[members.length - 1], item)) break;
      const next = flat(item.evidence.str);
      // A blank between two of its items is the box's; one after its last is not.
      if (!next) { blanks.push(item); continue; }
      if (!wanted.startsWith(text + next)) break;
      members.push(...blanks, item);
      blanks = [];
      text += next;
    }
    if (text === wanted) return members;
  }
  return null;
}

/** How far a pinned box's first item may stand from the origin its map
 * names: the reader's own number, read again from the same page. */
const PINNED_BOX_ORIGIN_TOLERANCE = 0.01;

/** Other words on the row of these items, within a word's space of them. */
function standsInsideLine(members: readonly ProjectedNativeItem[], items: readonly ProjectedNativeItem[]): boolean {
  const own = new Set(members);
  const head = members[0], tail = members[members.length - 1];
  const size = Math.max(Number(head.evidence.height) || 0, 1);
  const baseline = Number(head.evidence.transform[5]);
  const start = Number(head.evidence.transform[4]);
  const end = Number(tail.evidence.transform[4]) + (Number(tail.evidence.width) || 0);
  return items.some((item) => {
    if (own.has(item) || !item.evidence.str.trim() || Math.abs(Number(item.evidence.transform[5]) - baseline) > size * 0.3) return false;
    const left = Number(item.evidence.transform[4]), right = left + (Number(item.evidence.width) || 0);
    return Math.abs(left - end) <= size || Math.abs(start - right) <= size;
  });
}

/** The kept blocks the page still draws, each as one paragraph (at the
 * frame width its reader set, on an upright page whose space is the
 * block's), and the items they hold. */
function keptBlockParagraphs(
  blocks: readonly PdfNativeKeptTextBlock[] | undefined,
  pageNumber: number,
  projected: readonly ProjectedNativeItem[],
  viewport: PdfViewportLike,
  groupOf: (item: ProjectedNativeItem) => string | null,
  /** The page's lines, read once for every box that asks. */
  readPageLines: () => readonly NativeVisualLine[] = () => visualLines(projected),
): { paragraphs: MutableParagraph[]; claimed: Set<ProjectedNativeItem> } {
  const claimed = new Set<ProjectedNativeItem>();
  const paragraphs: MutableParagraph[] = [];
  if (!blocks?.length) return { paragraphs, claimed };
  const ordered = [...projected].sort((left, right) => left.evidence.itemIndex - right.evidence.itemIndex);
  for (const block of blocks) {
    // One row names a box only with where it stands.
    if (block.pageNumber !== pageNumber || block.rows.length < (block.origin ? 1 : 2)) continue;
    if (block.pinned) {
      // A box of the page's box map: the items it drew, as they were read.
      const members = block.origin ? pinnedBoxItems(block.rows, ordered, claimed, block.origin) : null;
      if (!members || members.some((member) => groupOf(member)?.startsWith(BOX_MARK_GROUP_PREFIX))) continue;
      // Its rows as the page's lines hold them (a digit set over another
      // reads as one line with the page's text around it, and as two on its
      // own: "43" came back "3 4"), in the order the reader reads them
      // (visualParagraphs): down the page, or along a turned text's block.
      // The page's lines move with what a write set beside the box (a title
      // indented took the name block's superscript star into a line of its
      // own: "∗Alan D. Sokal"), so the rows that read as the box's own text
      // are taken: the page's, else its items' own.
      const pageLines = readPageLines();
      const own = new Set(members);
      const byBlock = (lines: readonly NativeVisualLine[]) => [...lines].sort(
        (left, right) => lineBodyExtent(left).blockStart - lineBodyExtent(right).blockStart);
      const flat = (text: string) => text.replace(/\s+/gu, "");
      const wanted = flat(block.rows.join(""));
      const candidates = [
        paragraphOfItems(byBlock(pageLines.filter((line) => line.items.some((item) => own.has(item)))), own),
        paragraphOfItems(byBlock(visualLines(members)), own),
      ];
      const paragraph = candidates.find((candidate) => candidate &&
        flat(candidate.lines.map((line) => line.items.map((item) => item.evidence.str).join("")).join("")) === wanted) ??
        candidates[0];
      if (!paragraph) continue;
      for (const member of members) claimed.add(member);
      paragraphs.push(paragraph);
      continue;
    }
    const rows = keptBlockItems(block.rows, ordered, claimed, block.origin);
    const members = rows?.flat() ?? null;
    // A few words of a line, written on their own (a run a whole box could not
    // be proved for): the line goes on either side of them, and they are no box.
    // Taken out of it, a paragraph came apart into five (Chrome's Type3 text).
    if (rows && members && rows.length === 1 && !block.moved && !block.frame && standsInsideLine(members, ordered)) continue;
    // A box the writer marked in the page is the box the page draws: the
    // mark, written with the text, wins over the journal.
    if (!rows || !members || members.some((member) => groupOf(member)?.startsWith(BOX_MARK_GROUP_PREFIX))) continue;
    // Its lines are its rows, as the writer set them: a word of a row raised
    // as a superscript is read by the reader as a line of its own between
    // two rows, which put the block's items out of order and disabled it --
    // its rows opened one by one and the second set its typing over the
    
    const rowLines = rows.map((row) => lineFromItems(row));
    const paragraph = rowLines.every((line): line is NativeVisualLine => line !== null)
      ? paragraphOfItems(rowLines, new Set(members))
      : paragraphOfItems(visualLines(members), new Set(members));
    if (!paragraph) continue;
    const head = members[0];
    if (
      block.frame &&
      Math.abs(head.angle) < UPRIGHT_ANGLE_EPSILON &&
      Math.abs(Number(head.evidence.transform[4]) - block.frame.firstLeft) < 0.5
    ) {
      // The page draws the block in its own space: the frame's right edge,
      // where the reader left it, is where the box ends (WPS keeps a resized
      // block at its width).
      const right = multiplyMatrix(viewport.transform, [1, 0, 0, 1, block.frame.right, Number(head.evidence.transform[5])])[4];
      if (Number.isFinite(right) && right > paragraph.inlineEnd) {
        paragraph.right += right - paragraph.inlineEnd;
        paragraph.inlineEnd = right;
      }
    }
    for (const member of members) claimed.add(member);
    paragraphs.push(paragraph);
  }
  return { paragraphs, claimed };
}

/**
 * Build WPS-style paragraph controls while retaining exact pdf.js item
 * identities for the native content-stream writer.
 *
 * Visual lines are reconstructed independently of font-run fragmentation, so
 * an italic word, fallback subset or punctuation run cannot split the visible
 * paragraph boundary. The second pass respects column alignment, heading size
 * changes and paragraph gaps. A visual paragraph may contain multiple native
 * font resources; the source-bound planner, rather than the presentation
 * layer, proves each run's font/CMap/operator identity before permitting a
 * write. Unsupported structures never fall back to a synthetic
 * FreeText/overlay edit.
 */
export function nativeParagraphEditBoxes(
  input: Readonly<{
    pageNumber: number;
    items: readonly (PdfNativeListTextSelectionItemEvidence | null)[];
    viewport: PdfViewportLike;
    fontStyles?: Readonly<
      Record<
        string,
        Readonly<{
          fontFamily?: string;
          ascent?: number;
          descent?: number;
          vertical?: boolean;
        }>
      >
    >;
    /** pdfJsEmptyLineEndItemIndexes of the page's raw items. */
    emptyLineEnds?: ReadonlySet<number>;
    /** Blocks the editor set itself (pdfNativeKeptTextBlocks), newest
     * first: each one the page still draws stays one box. */
    keptBlocks?: readonly PdfNativeKeptTextBlock[];
    /** The box each raw item is drawn in, as the file says
     * (pdfNativeTextHardGroups): items of two keys never share a line or a
     * box. */
    hardGroups?: readonly (string | null)[] | null;
  }>,
): readonly PdfNativeParagraphEditBox[] {
  const normalizedItems = normalizePdfJsNativeParagraphItems(input.items,
    pdfNativeBoxOriginsOf(input.keptBlocks, input.pageNumber));
  const overprints = pdfJsOverprintItemIndexes(normalizedItems);
  const spacers: ProjectedNativeItem[] = [];
  const projected = normalizedItems.flatMap((item) => {
    if (!item || overprints.has(item.itemIndex)) return [];
    // pdf.js also emits zero-height whitespace spanning positioned table and
    // chart cells. It is layout separation, not source text owned by either
    // cell. Narrow split spaces were already folded into their exact run by
    // normalizePdfJsNativeParagraphItems; every remaining spacer is therefore
    // a hard visual boundary and must never bridge independent edit boxes.
    // It is kept aside only to be read back inside a line that its two
    // neighbours already form on their own (singleRunLine).
    const spacer = item.height === 0 && /^ +$/u.test(item.str);
    const style = input.fontStyles?.[item.fontName];
    const visualFontKey = style
      ? JSON.stringify([
          style.fontFamily?.trim().toLowerCase() ?? "",
          Number.isFinite(style.ascent) ? style.ascent : null,
          Number.isFinite(style.descent) ? style.descent : null,
          style.vertical === true,
        ])
      : `pdfjs-font:${item.fontName}`;
    const value = projectNativeItem(
      item,
      input.viewport,
      visualFontKey,
      style?.vertical === true,
    );
    if (value && spacer) {
      spacers.push(value);
      return [];
    }
    return value ? [value] : [];
  });
  if (projected.length === 0) return Object.freeze([]);

  // A block the editor set stays the one box it was set as; the page's other
  // text is grouped without it, so a row of the block that wrapped onto the
  // next paragraph's baseline is not read into that paragraph.
  const groupOf = hardGroupOf(input.hardGroups);
  let allLines: readonly NativeVisualLine[] | undefined;
  const pageLines = () => (allLines ??= visualLines(projected));
  const kept = keptBlockParagraphs(input.keptBlocks, input.pageNumber, projected, input.viewport, groupOf, pageLines);
  const free = kept.claimed.size ? projected.filter((item) => !kept.claimed.has(item)) : projected;
  // The boxes the file names (the editor's box marks, a tagged file's
  // paragraphs) are grouped each on its own: text of two never shares a
  // line or a box, however close it is set. A page without them is one part.
  const parts = new Map<string | null, ProjectedNativeItem[]>();
  // (A diagram's label, where the file names no box: centredStacks.)
  const stackOf = new Map<ProjectedNativeItem, string>();
  for (const stack of centredStacks(free.filter((item) => groupOf(item) === null), projected)) {
    for (const item of stack) stackOf.set(item, `${CENTRED_STACK_GROUP_PREFIX}${stack[0].evidence.itemIndex}`);
  }
  for (const item of free) {
    const key = stackOf.get(item) ?? groupOf(item);
    const part = parts.get(key);
    if (part) part.push(item);
    else parts.set(key, [item]);
  }
  const paragraphs = [
    ...kept.paragraphs,
    ...[...parts.entries()].flatMap(([key, part]) => {
      if (key?.startsWith(BOX_MARK_GROUP_PREFIX)) {
        // One mark, one box, its rows as the page's lines hold them (a sum's
        // limits read into the line beside them only with the page's text
        // around: alone, "n=1" was cut from the line it was set in).
        const own = new Set(part);
        const fromPage = paragraphOfItems(pageLines().filter((line) => line.items.some((item) => own.has(item))), own);
        const candidates = [sourceOrderRows(part), fromPage?.lines ?? null, visualLines(part)]
          .map((lines) => lines?.length ? markedBoxParagraph(lines) : null);
        // The rows that read in the order the page draws them, as a write sets
        // a box (a paragraph set again draws its text in its own order); else
        // the page's.
        const marked = candidates.find((candidate) => candidate && readsInSourceOrder(candidate[0])) ??
          candidates[1] ?? candidates[2];
        if (marked) return marked;
      }
      if (key?.startsWith(CENTRED_STACK_GROUP_PREFIX)) {
        // One label, one box: each of its items a row, a mark hanging
        // after a line on that line's row.
        const rowItems: ProjectedNativeItem[][] = [];
        for (const item of part) {
          const row = rowItems[rowItems.length - 1];
          if (row && hangsAfterStackLine(row[0], item)) row.push(item);
          else rowItems.push([item]);
        }
        const rows = rowItems.map((row) => lineFromItems(row));
        const stacked = rows.every((line): line is NativeVisualLine => line !== null) ? markedBoxParagraph(rows) : null;
        if (stacked) return stacked;
      }
      return visualParagraphs(visualLines(part))
        .flatMap((paragraph) => sourceContiguousParagraphs(paragraph, projected));
    }),
  ].flatMap(paragraphsWithinTextLimit);
  const boxes = paragraphs.map((grouped) => {
    // A turned line set with character or word spacing and a horizontal
    // scale: the reader splits its one show into pieces, and the scale makes
    // them fail the turned-paragraph test, which left the line the editor
    // had just rotated closed to every later edit (walk s1-c: Rotate Right
    // on a line with 1 Tc 3 Tw 120 Tz could not be opened, turned back or
    // deleted). One run is one run at any angle.
    const turnedRun = Math.abs(grouped.lines[0].angle) >= UPRIGHT_ANGLE_EPSILON
      ? singleRunLine(grouped.lines, spacers, projected) : null;
    const reason = paragraphReason(grouped.lines, grouped.lines[0].angle, turnedRun !== null);
    const runLine =
      reason === undefined ? (turnedRun ?? singleRunLine(grouped.lines, spacers, projected)) : null;
    const paragraph = runLine
      ? { ...grouped, lines: [runLine] }
      : {
          ...grouped,
          lines: grouped.lines.map((line) =>
            lineWithReaderGapSpaces(line, spacers, input.viewport),
          ),
        };
    const lineOfItem = new Map(
      paragraph.lines.flatMap((line, lineIndex) =>
        line.items.map((item) => [item.evidence, lineIndex] as const),
      ),
    );
    const ordered = paragraph.lines
      .flatMap((line) => line.items.map((item) => item.evidence))
      .sort((left, right) => left.itemIndex - right.itemIndex);
    // The reader ended its line right after an item and carried the EOL on
    // an empty item of its own (pdfJsEmptyLineEndItemIndexes), inside a line
    // of the box: TeX draws a big delimiter after an exponent in another
    // face, and the reader flushes the exponent's item, then ends its line
    // at the delimiter. The item closes that line to the reader as it would
    // had it carried the EOL itself, so the box says so: a paragraph
    // rewrite of the box then reads the gap after it as one its text does
    // not hold as a space (paragraph-rewrite's unreadGaps).
    const items = ordered.map((item, index) => {
      const next = ordered[index + 1];
      const last = rawSourceItemIndexes(item).at(-1) ?? item.itemIndex;
      return next &&
        !item.hasEOL &&
        input.emptyLineEnds?.has(last + 1) &&
        lineOfItem.get(item) === lineOfItem.get(next)
        ? Object.freeze({ ...item, hasEOL: true })
        : item;
    });
    const sourceItemIndexes = Object.freeze(
      Array.from(
        new Set(items.flatMap((item) => rawSourceItemIndexes(item))),
      ).sort((left, right) => left - right),
    );
    const angle = paragraph.lines[0].angle;
    const radians = (angle * Math.PI) / 180;
    const cosine = Math.cos(radians);
    const sine = Math.sin(radians);
    const orientedWidth = paragraph.inlineEnd - paragraph.inlineStart;
    const orientedHeight = paragraph.blockEnd - paragraph.blockStart;
    const inlineCenter = (paragraph.inlineStart + paragraph.inlineEnd) / 2;
    const blockCenter = (paragraph.blockStart + paragraph.blockEnd) / 2;
    const centerX = inlineCenter * cosine - blockCenter * sine;
    const centerY = inlineCenter * sine + blockCenter * cosine;
    const boxLeft = centerX - orientedWidth / 2;
    const boxTop = centerY - orientedHeight / 2;
    const runStart = runLine?.items[0].evidence;
    const runEnd = runLine?.items.at(-1)?.evidence;
    const runScale = runStart
      ? Math.hypot(runStart.transform[0], runStart.transform[1])
      : 0;
    const selectionItem =
      runLine && runStart && runEnd && runScale > 0
        ? Object.freeze({
            ...runStart,
            str: runLine.text,
            width:
              ((runEnd.transform[4] - runStart.transform[4]) *
                runStart.transform[0] +
                (runEnd.transform[5] - runStart.transform[5]) *
                  runStart.transform[1]) /
                runScale +
              runEnd.width,
            hasEOL: runEnd.hasEOL,
          })
        : undefined;
    const first = items[0].itemIndex;
    const last = items.at(-1)?.itemIndex ?? first;
    return Object.freeze({
      id: `pdf-native-paragraph-${input.pageNumber}-${first}-${last}`,
      x: boxLeft / input.viewport.width,
      y: boxTop / input.viewport.height,
      width: orientedWidth / input.viewport.width,
      height: orientedHeight / input.viewport.height,
      ...(angle ? { angle } : {}),
      text: paragraph.lines.map((line) => line.text).join("\n"),
      items: Object.freeze(items),
      sourceItemIndexes,
      ...(selectionItem ? { selectionItem } : {}),
      editable: reason === undefined,
      ...(reason ? { disabledReason: reason } : {}),
    });
  });
  boxes.sort((left, right) => left.y - right.y || left.x - right.x);
  return Object.freeze(boxes);
}

/**
 * A page's edit boxes from its text content, as the viewer reads them: each
 * item read as a selection names it, normalized, and grouped with the
 * page's empty line ends, hard groups and the blocks the editor kept. The
 * one reading of "which box is this text": the viewer's boxes and the
 * geometry a candidate preview is drawn with (native-candidate-text-
 * geometry.ts) both come from it, so what the editor places its caret by
 * is the box the page shows once the edit is written.
 */
export function nativeParagraphEditBoxesOfTextContent(input: Readonly<{
  pageNumber: number;
  content: Readonly<{
    items: readonly unknown[];
    styles?: Parameters<typeof nativeParagraphEditBoxes>[0]["fontStyles"];
    hardGroups?: readonly (string | null)[] | null;
  }>;
  viewport: PdfViewportLike;
  keptBlocks?: readonly PdfNativeKeptTextBlock[];
}>): Readonly<{
  items: readonly (PdfNativeListTextSelectionItemEvidence | null)[];
  boxes: readonly PdfNativeParagraphEditBox[];
}> {
  const items = normalizePdfJsNativeParagraphItems(
    input.content.items.map((item, index) => nativeListTextItemEvidence(item, index)),
    pdfNativeBoxOriginsOf(input.keptBlocks, input.pageNumber));
  const boxes = nativeParagraphEditBoxes({
    pageNumber: input.pageNumber,
    items,
    viewport: input.viewport,
    ...(input.content.styles ? { fontStyles: input.content.styles } : {}),
    emptyLineEnds: pdfJsEmptyLineEndItemIndexes(input.content.items),
    ...(input.keptBlocks ? { keptBlocks: input.keptBlocks } : {}),
    ...(input.content.hardGroups !== undefined ? { hardGroups: input.content.hardGroups } : {}),
  });
  return Object.freeze({ items, boxes });
}
