/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Per-paragraph TAB STOPS (Word/WPS 制表位) — the pure, deterministic data model
 * shared by the editor schema, the DOCX exporter, and the DOCX importer.
 *
 * A paragraph may carry an optional `tabStops` attribute: an array of custom tab
 * stops that a Tab character advances to. Each stop is `{ pos, align, leader }`:
 *   - `pos`    — the stop position in TWIPS (1/1440 inch). Word's ruler runs
 *                0..~15840 twips (11"); we clamp to that band.
 *   - `align`  — how text at the stop aligns: left / center / right / decimal.
 *   - `leader` — the fill character drawn from the previous text to the stop:
 *                none / dot / hyphen / underscore (Word's `w:leader`).
 *
 * The model is normalized to be deterministic: stops are clamped, sorted by
 * position, de-duplicated by position (last wins), and bounded in count. A
 * paragraph with NO tab stops carries no attribute at all, so it serializes,
 * exports, and renders byte-for-byte identically to before this feature.
 *
 * This module is intentionally free of React / ProseMirror / DOM imports so it
 * can be unit-tested in isolation and reused on both the export and import side.
 */

/**
 * Tab-stop alignment values (mirror of Word's `w:tab w:val`). `bar` is Word's
 * Bar tab (ECMA-376 17.18.84 ST_TabJc "bar"): it is NOT a stop a tab
 * character moves to; it draws a vertical line at its position through the
 * paragraph's lines, and tab positioning skips it.
 */
export const TAB_STOP_ALIGNMENTS = ["left", "center", "right", "decimal", "bar"] as const;
export type TabStopAlign = (typeof TAB_STOP_ALIGNMENTS)[number];

/** Tab-stop leader fills (mirror of Word's `w:tab w:leader`). */
export const TAB_STOP_LEADERS = ["none", "dot", "hyphen", "underscore"] as const;
export type TabStopLeader = (typeof TAB_STOP_LEADERS)[number];

/** One normalized tab stop. `leader` is always present ("none" when unset). */
export interface TabStop {
  /** Position in TWIPS (1/1440 inch), clamped to MIN_TAB_POS..MAX_TAB_POS. */
  pos: number;
  align: TabStopAlign;
  leader: TabStopLeader;
}

/**
 * Position bounds in TWIPS. Word's tab-stop ruler starts at 0 and runs to the
 * page width (~11in = 15840 twips). We require a strictly positive position
 * (a stop at 0 is meaningless — that is the line start) and cap at 15840.
 */
export const MIN_TAB_POS = 1;
export const MAX_TAB_POS = 15840;

/** Hard cap on stored stops per paragraph (Word's own dialog tops out near 50). */
export const MAX_TAB_STOPS = 50;

const TWIPS_PER_INCH = 1440;
const PX_PER_INCH = 96;
/** Twips-per-pixel at 96dpi: 1440/96 = 15. Exact integer, so px↔twips is clean. */
const TWIPS_PER_PX = TWIPS_PER_INCH / PX_PER_INCH;

const ALIGN_SET = new Set<string>(TAB_STOP_ALIGNMENTS);
const LEADER_SET = new Set<string>(TAB_STOP_LEADERS);

/** Twips → CSS pixels at 96dpi (1440 twips/inch, 96 px/inch). Rounded. */
export function twipsToPx(twips: number): number {
  if (!Number.isFinite(twips)) return 0;
  return Math.round(twips / TWIPS_PER_PX);
}

/** CSS pixels → twips at 96dpi. Rounded. */
export function pxToTwips(px: number): number {
  if (!Number.isFinite(px)) return 0;
  return Math.round(px * TWIPS_PER_PX);
}

/** Centimetres → twips (1in = 2.54cm = 1440 twips). Rounded. */
export function cmToTwips(cm: number): number {
  if (!Number.isFinite(cm)) return 0;
  return Math.round((cm / 2.54) * TWIPS_PER_INCH);
}

/** Twips → centimetres (1in = 2.54cm = 1440 twips). */
export function twipsToCm(twips: number): number {
  if (!Number.isFinite(twips)) return 0;
  return (twips / TWIPS_PER_INCH) * 2.54;
}

/** Inches → twips. Rounded. */
export function inchesToTwips(inches: number): number {
  if (!Number.isFinite(inches)) return 0;
  return Math.round(inches * TWIPS_PER_INCH);
}

/** Twips → inches. */
export function twipsToInches(twips: number): number {
  if (!Number.isFinite(twips)) return 0;
  return twips / TWIPS_PER_INCH;
}

/** Whitelisted alignment, defaulting to "left" for unknown / missing values. */
export function normalizeTabAlign(value: unknown): TabStopAlign {
  if (typeof value === "string") {
    // Word writes start/end for left/right in its modern serialization; fold
    // those (and the legacy num spelling) onto the model's values.
    if (value === "start") return "left";
    if (value === "end") return "right";
    if (value === "num") return "decimal";
    if (ALIGN_SET.has(value)) return value as TabStopAlign;
  }
  return "left";
}

/** Whitelisted leader, defaulting to "none". Word's `middleDot` folds to dot. */
export function normalizeTabLeader(value: unknown): TabStopLeader {
  if (typeof value === "string") {
    if (value === "middleDot") return "dot";
    if (LEADER_SET.has(value)) return value as TabStopLeader;
  }
  return "none";
}

/** Clamp a raw position to the valid twips band, or null when not a number. */
export function clampTabPos(value: unknown): number | null {
  const num = typeof value === "number" ? value : Number.parseFloat(String(value));
  if (!Number.isFinite(num)) return null;
  const rounded = Math.round(num);
  if (rounded < MIN_TAB_POS) return rounded <= 0 ? null : MIN_TAB_POS;
  return Math.min(rounded, MAX_TAB_POS);
}

/**
 * Normalize an arbitrary value into a clean, deterministic TabStop[] (or null
 * when there are no valid stops). The result is:
 *   - filtered to entries with a valid position,
 *   - clamped to the twips band,
 *   - de-duplicated by position (a later entry at the same position wins),
 *   - sorted ascending by position,
 *   - capped at MAX_TAB_STOPS.
 * Returning `null` (not `[]`) for the empty case keeps the attribute ABSENT on
 * paragraphs without tab stops, preserving byte-for-byte stability.
 */
export function normalizeTabStops(value: unknown): TabStop[] | null {
  const raw = toArray(value);
  if (!raw) return null;
  const byPos = new Map<number, TabStop>();
  for (const entry of raw) {
    if (!entry || typeof entry !== "object") continue;
    const record = entry as Record<string, unknown>;
    const pos = clampTabPos(record.pos);
    if (pos === null) continue;
    byPos.set(pos, {
      pos,
      align: normalizeTabAlign(record.align),
      leader: normalizeTabLeader(record.leader),
    });
  }
  if (byPos.size === 0) return null;
  const sorted = Array.from(byPos.values()).sort((a, b) => a.pos - b.pos);
  return sorted.slice(0, MAX_TAB_STOPS);
}

/**
 * A paragraph's tab stops with the one Word adds at a hanging indent: on the
 * first line of a paragraph whose first line hangs (a negative first-line
 * indent), the left indent is a left stop of its own (Word's "automatic tab
 * stop for hanging indent", on unless `w:noTabHangInd`). Without it the first
 * tab of a hanging paragraph went to the first custom stop and every tab after
 * it one stop further than Word's: R350-D07's interlinear glosses, one line in
 * Word (cat-Sg-Sg at its 1080-twip stop), ran past the margin and took two.
 * On every later line the pen starts at the left indent, past this stop.
 */
export function withHangingIndentStop(
  stops: TabStop[] | null,
  indentLeftPt: unknown,
  firstLineIndentPt: unknown,
): TabStop[] | null {
  const left = typeof indentLeftPt === "number" && Number.isFinite(indentLeftPt) ? indentLeftPt : 0;
  const first = typeof firstLineIndentPt === "number" && Number.isFinite(firstLineIndentPt) ? firstLineIndentPt : 0;
  if (!(first < 0) || !(left > 0)) return stops;
  const pos = clampTabPos(Math.round(left * 20));
  if (pos === null || stops?.some((stop) => stop.pos === pos)) return stops;
  const all = [...(stops ?? []), { pos, align: "left" as TabStopAlign, leader: "none" as TabStopLeader }];
  return all.sort((a, b) => a.pos - b.pos).slice(0, MAX_TAB_STOPS + 1);
}













export function hangingMarginTabTarget(penPx: number, nextCustomPx: number | null, hangPx: number | null): number | null {
  if (hangPx === null || !(hangPx > penPx + 0.01)) return null;
  if (nextCustomPx !== null && !(hangPx < nextCustomPx - 0.01)) return null;
  return hangPx;
}

/**
 * Accepts either a real array or the JSON string the editor serializes the
 * attribute to in its `data-tab-stops` attribute. Anything else → null.
 */
function toArray(value: unknown): unknown[] | null {
  if (Array.isArray(value)) return value;
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return null;
    try {
      const parsed = JSON.parse(trimmed);
      return Array.isArray(parsed) ? parsed : null;
    } catch {
      return null;
    }
  }
  return null;
}

/**
 * Serialize a normalized TabStop[] to the compact JSON string stored in the
 * editor's `data-tab-stops` HTML attribute. Returns null when there are no
 * stops so the attribute is omitted entirely (byte-stable). The stored shape is
 * the SAME normalized model, so a parse → normalize round-trip is idempotent.
 */
export function serializeTabStops(value: unknown): string | null {
  const stops = normalizeTabStops(value);
  if (!stops) return null;
  return JSON.stringify(stops);
}

/** Word's default tab interval (`w:defaultTabStop`, 0.5in) in twips. */
export const WORD_DEFAULT_TAB_INTERVAL_TWIPS = 720;

/** One stretch of a line between tabs, as the layout needs it. */
export interface TabLayoutSegment {
  /** Drawn width of the text up to the next tab (or the line end), in px. */
  widthPx: number;
  /** For a decimal stop: the width before the segment's first ".", in px. */
  decimalPx?: number;
}

/**
 * Word's absolute position tab (`w:ptab`, ISO/IEC 29500-1 §17.3.3.23; Word's
 * "Insert Alignment Tab"). Unlike a tab character it ignores every custom tab
 * stop and the default tab interval: it advances to a point fixed by its own
 * attributes -- the left edge, the centre or the right edge (`alignment`) of
 * the text margins or of the paragraph's indents (`relativeTo`) -- and the
 * text after it is aligned there as a stop of that alignment would align it.
 * `leader` fills the space it spans (`middleDot` is Word's centred dot).
 */
export const POSITIONAL_TAB_ALIGNMENTS = ["left", "center", "right"] as const;
export const POSITIONAL_TAB_BASES = ["margin", "indent"] as const;
export const POSITIONAL_TAB_LEADERS = ["none", "dot", "hyphen", "underscore", "middleDot"] as const;

export interface PositionalTab {
  alignment: (typeof POSITIONAL_TAB_ALIGNMENTS)[number];
  relativeTo: (typeof POSITIONAL_TAB_BASES)[number];
  leader: (typeof POSITIONAL_TAB_LEADERS)[number];
}

/** An exact `{ alignment, relativeTo, leader }`, or null for anything else. */
export function normalizePositionalTab(value: unknown): PositionalTab | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (keys.length !== 3 || !keys.every((key) => key === "alignment" || key === "relativeTo" || key === "leader")) {
    return null;
  }
  const { alignment, relativeTo, leader } = record;
  if (
    !(POSITIONAL_TAB_ALIGNMENTS as readonly unknown[]).includes(alignment) ||
    !(POSITIONAL_TAB_BASES as readonly unknown[]).includes(relativeTo) ||
    !(POSITIONAL_TAB_LEADERS as readonly unknown[]).includes(leader)
  ) {
    return null;
  }
  return { alignment, relativeTo, leader } as PositionalTab;
}

/** The line a positional tab is measured against, in px: the text column's
 * width (margin to margin) and the paragraph's left and right indents. */
export interface PositionalTabLine {
  widthPx: number;
  indentLeftPx?: number;
  indentRightPx?: number;
}

/**
 * How near a stop the pen may stand and still be ON it, in px. Word moves a
 * tab to the first stop right of the pen, and a pen standing on a stop goes to
 * the one after. Text widths are measured in fractions of a pixel, and an
 * advance is floored to half a pixel, so a pen "on" a stop reads a hair short
 * of it: without this reach the tab came out zero or half a pixel wide and
 * could not be seen (full-width text 16 px wide puts every third character on
 * a 0.5 in stop).
 */
export const TAB_STOP_REACH_PX = 0.75;

/** The custom stops a tab character can move to, left to right. */
function reachableStops(stops: readonly TabStop[] | null | undefined): TabStop[] {
  // A bar tab draws a line; it is never a position a tab character moves to.
  return (normalizeTabStops(stops ?? null) ?? []).filter((stop) => stop.align !== "bar");
}

/** Word's default tab interval in px for a stated `w:defaultTabStop`. */
function defaultIntervalPx(twips: number | undefined): number {
  return Math.max(1, (twips ?? WORD_DEFAULT_TAB_INTERVAL_TWIPS) / TWIPS_PER_PX);
}

/**
 * How far ONE tab advances the pen, in px, from where the pen stands on its
 * line: {@link wordTabAdvances}' rule for a single tab. `next` is the text
 * after the tab, up to the next tab or the end of the line. A paragraph that
 * wraps lays each tab out from its own place on its own line, which is this.
 */
export function wordTabAdvanceAt(input: {
  penPx: number;
  next: TabLayoutSegment;
  stops: readonly TabStop[] | null | undefined;
  defaultIntervalTwips?: number;
  positional?: PositionalTab | null;
  line?: PositionalTabLine;
  /** Word 2013 and later: a centre or right stop past the paragraph's right
   * edge ends its text at that edge (see `tabAdvance`). */
  stopsEndAtRightEdge?: boolean;
}): number {
  return tabAdvance(
    input.penPx,
    input.next,
    reachableStops(input.stops),
    defaultIntervalPx(input.defaultIntervalTwips),
    input.line ? (input.positional ?? null) : null,
    input.line,
    input.stopsEndAtRightEdge === true,
  );
}

function tabAdvance(
  pen: number,
  segment: TabLayoutSegment,
  stops: readonly TabStop[],
  interval: number,
  positional: PositionalTab | null,
  line: PositionalTabLine | undefined,
  stopsEndAtRightEdge = false,
): number {
  const width = Math.max(0, segment.widthPx);
  const stop = positional
    ? undefined
    : stops.find((candidate) => candidate.pos / TWIPS_PER_PX > pen + TAB_STOP_REACH_PX);
  let target: number;
  if (positional && line) {
    const from = positional.relativeTo === "indent" ? Math.max(0, line.indentLeftPx ?? 0) : 0;
    const to = line.widthPx - (positional.relativeTo === "indent" ? Math.max(0, line.indentRightPx ?? 0) : 0);
    const at = positional.alignment === "left" ? from : positional.alignment === "center" ? (from + to) / 2 : to;
    target =
      positional.alignment === "center" ? at - width / 2 : positional.alignment === "right" ? at - width : at;
  } else if (
    stop &&
    stopsEndAtRightEdge &&
    line &&
    (stop.align === "center" || stop.align === "right" || stop.align === "decimal") &&
    stop.pos / TWIPS_PER_PX > line.widthPx - Math.max(0, line.indentRightPx ?? 0) + TAB_STOP_REACH_PX
  ) {
    // Word 16.86 (compatibility mode 15): a centre, right or decimal stop
    // past the paragraph's right edge is a right stop AT that edge -- its
    // text ends there, on the line (corpus R49-D14 TA12i/TA12j: stops at 9500
    // twips with a 1000-twip right indent end "centred" / "right" at the
    // indent; R83-D10 TP2: "12.5" at a decimal stop at 9800 twips ends at the
    // 9026-twip margin, on its line).
    target = line.widthPx - Math.max(0, line.indentRightPx ?? 0) - width;
  } else if (stop) {
    const at = stop.pos / TWIPS_PER_PX;
    target =
      stop.align === "center"
        ? at - width / 2
        : stop.align === "right"
          ? at - width
          : stop.align === "decimal"
            ? at - Math.max(0, Math.min(width, segment.decimalPx ?? width))
            : at;
  } else {
    target = (Math.floor((pen + TAB_STOP_REACH_PX) / interval) + 1) * interval;
  }
  // Floored to half a pixel: a segment aligned to a right stop that ends
  // exactly at the margin must not spill onto the next line by a rounding
  // error (Word never moves text past its stop).
  let advance = Math.max(0, Math.floor((target - pen) * 2) / 2);
  if (positional && line) {
    // Text after a positional tab never runs past the paragraph's right
    // edge because of the tab: Word wraps only text wider than the room.
    const right = line.widthPx - Math.max(0, line.indentRightPx ?? 0);
    advance = Math.min(advance, Math.max(0, Math.floor((right - pen - width) * 2) / 2));
  }
  return advance;
}

/**
 * How far each tab on one line advances the pen, in px, as Word lays it out.
 *
 * `segments` holds the text around the tabs: `segments[0]` before the first
 * tab, `segments[i + 1]` after tab `i`, so there is one more segment than
 * tabs. A tab moves to the first custom stop right of the pen: a left stop
 * starts the next segment there, a centre stop centres it on the stop, a right
 * stop ends it there, and a decimal stop puts its first "." there (a segment
 * without one ends there, as Word does). A stop the segment cannot reach
 * without moving backwards leaves the tab empty. Past the last custom stop the
 * pen goes to the next default stop; Word drops the default stops left of a
 * custom one, which this gets by only reaching them past every custom stop.
 *
 * Tab `i` that is a positional tab (`positional[i]`, see
 * {@link PositionalTab}) goes to its own point on `line` instead and aligns
 * the next segment there the same way. Word moves one whose point the pen is
 * already past to that point on the next line; a band line does not break, so
 * such a tab is left empty like an unreachable stop.
 *
 * Positions are measured from the text column's left edge (Word's stops do
 * not move with the paragraph indent); `startPx` is where the line's text
 * starts, the indent included. The pen runs on from tab to tab, so this is
 * the layout of ONE drawn line; a paragraph that wraps uses
 * {@link wordTabAdvanceAt} per tab.
 */
export function wordTabAdvances(input: {
  segments: readonly TabLayoutSegment[];
  stops: readonly TabStop[] | null | undefined;
  startPx?: number;
  defaultIntervalTwips?: number;
  positional?: readonly (PositionalTab | null | undefined)[];
  line?: PositionalTabLine;
  stopsEndAtRightEdge?: boolean;
}): number[] {
  const { segments, line } = input;
  const stops = reachableStops(input.stops);
  const interval = defaultIntervalPx(input.defaultIntervalTwips);
  const advances: number[] = [];
  if (segments.length < 2) return advances;
  let pen = (input.startPx ?? 0) + Math.max(0, segments[0].widthPx);
  for (let index = 1; index < segments.length; index += 1) {
    const segment = segments[index];
    const positional = line ? (input.positional?.[index - 1] ?? null) : null;
    const advance = tabAdvance(pen, segment, stops, interval, positional, line, input.stopsEndAtRightEdge === true);
    advances.push(advance);
    pen += advance + Math.max(0, segment.widthPx);
  }
  return advances;
}

/**
 * The leader a tab draws: a positional tab's own, else the stop it lands on
 * ({@link wordTabStopAfter}); Word's centred dot draws as the dot leader.
 */
export function wordTabLeader(
  stops: readonly TabStop[] | null | undefined,
  penPx: number,
  positional?: PositionalTab | null,
): TabStopLeader {
  if (positional) return normalizeTabLeader(positional.leader);
  return wordTabStopAfter(stops, penPx)?.leader ?? "none";
}

/**
 * The stop a tab lands on, for drawing its leader: the first custom stop right
 * of `penPx`, or null past them all (a default stop draws no leader).
 */
export function wordTabStopAfter(
  stops: readonly TabStop[] | null | undefined,
  penPx: number,
): TabStop | null {
  const normalized = (normalizeTabStops(stops ?? null) ?? []).filter(
    (stop) => stop.align !== "bar",
  );
  return normalized.find((candidate) => candidate.pos / TWIPS_PER_PX > penPx + TAB_STOP_REACH_PX) ?? null;
}
