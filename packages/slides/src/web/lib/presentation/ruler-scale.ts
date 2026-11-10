/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PowerPoint's rulers (View > Ruler): numbered in the system's measurement
 * unit (inches or centimetres), 0 at the slide's centre -- or, while words
 * are being edited, at the left (top) of the text box, which then shows its
 * own indents and tab stops -- and, while an object is dragged or sized, its
 * extent marked along the scale.
 *
 * `presentationRulerTicks` is the pure scale; `drawPresentationRuler` paints
 * one band into a canvas.
 */

export type PresentationRulerUnit = "in" | "cm";

/** CSS px per unit at 100% (96 dpi). */
export const PRESENTATION_RULER_PX_PER_UNIT: Record<PresentationRulerUnit, number> = {
  in: 96,
  cm: 96 / 2.54,
};

/** Divisions of one unit: eighths of an inch, quarters of a centimetre. */
const DIVISIONS: Record<PresentationRulerUnit, number> = { in: 8, cm: 4 };

/** EMU in one unit, and in the smallest division the ruler snaps to. */
export const PRESENTATION_RULER_EMU_PER_UNIT: Record<PresentationRulerUnit, number> = {
  in: 914_400,
  cm: 360_000,
};

/** The ruler's snap: a marker or a stop lands on the nearest division. */
export function presentationRulerSnap(emu: number, unit: PresentationRulerUnit): number {
  const step = PRESENTATION_RULER_EMU_PER_UNIT[unit] / DIVISIONS[unit];
  return Math.round(Math.round(emu / step) * step);
}

/** The most `a:pPr@marL` / `@indent` reach (ST_TextMargin / ST_TextIndent). */
const MAX_TEXT_MARGIN_EMU = 51_206_400;

export type PresentationRulerIndentMarker = "firstLine" | "hanging" | "left";

/**
 * A paragraph's indents after one of the ruler's markers was dropped at
 * `atEmu` (from the text area's left), as PowerPoint's ruler moves them
 * (Microsoft Support, "Adjust the indent in a bulleted or numbered list on
 * the ruler"): the first-line marker moves the first line (the bullet) alone;
 * the hanging marker -- the left marker's pointed top -- moves the text and
 * leaves the first line where it is; the left marker's box moves both,
 * keeping their distance. `marL` is the text's margin, `indent` the first
 * line's offset from it (negative: hanging), in EMU. Neither the text nor the
 * first line goes left of the text area.
 */
export function presentationRulerIndentDrop(
  marker: PresentationRulerIndentMarker,
  current: { marL: number; indent: number },
  atEmu: number,
): { marL: number; indent: number } {
  const clamp = (value: number) => Math.min(MAX_TEXT_MARGIN_EMU, Math.max(0, Math.round(value)));
  const first = current.marL + current.indent;
  if (marker === "firstLine") {
    const nextFirst = clamp(atEmu);
    return { marL: current.marL, indent: nextFirst - current.marL };
  }
  if (marker === "hanging") {
    const marL = clamp(atEmu);
    return { marL, indent: first - marL };
  }
  // the box: both move, the first line no further left than the text area
  let marL = clamp(atEmu);
  if (marL + current.indent < 0) marL = clamp(-current.indent);
  return { marL, indent: current.indent };
}

export type PresentationTabAlign = "left" | "center" | "right" | "decimal";

/** A box's tab stops as the page holds them: a bare number (design px) for a left stop. */
export type PresentationStoredTabStop = number | { pos: number; align: PresentationTabAlign };

/** The stored stops read as `{pos, align}` (unknown entries left out). */
export function presentationTabStopsRead(value: unknown): Array<{ pos: number; align: PresentationTabAlign }> {
  if (!Array.isArray(value)) return [];
  return value.flatMap((stop) => {
    const pos = typeof stop === "number" ? stop : (stop as { pos?: unknown } | null)?.pos;
    const raw = typeof stop === "object" && stop ? (stop as { align?: unknown }).align : "left";
    const align: PresentationTabAlign =
      raw === "center" || raw === "right" || raw === "decimal" ? raw : "left";
    return typeof pos === "number" && Number.isFinite(pos) && pos > 0 ? [{ pos, align }] : [];
  });
}

/**
 * The stops written back: ascending, one per position (the later wins), a
 * left stop as its bare position as the importer reads one, at most 20 (the
 * page's limit), positions at the importer's 0.01 px.
 */
export function presentationTabStopsWrite(
  stops: ReadonlyArray<{ pos: number; align: PresentationTabAlign }>,
): PresentationStoredTabStop[] {
  const byPos = new Map<number, PresentationTabAlign>();
  for (const stop of stops) {
    const pos = Math.round(stop.pos * 100) / 100;
    if (pos > 0 && pos <= 4096) byPos.set(pos, stop.align);
  }
  return [...byPos.entries()]
    .sort((a, b) => a[0] - b[0])
    .slice(0, 20)
    .map(([pos, align]) => (align === "left" ? pos : { pos, align }));
}

/** The least room between two marks (px), and between two numbers. */
const MIN_TICK_GAP_PX = 4;
const MIN_LABEL_GAP_PX = 26;

export interface PresentationRulerTick {
  /** Along the band, in px from its start. */
  pos: number;
  /** Tick height class: a whole unit, a half, or a smaller division. */
  kind: "unit" | "half" | "minor";
  /** The number written at a whole unit, when one is written. */
  label?: string;
}

/**
 * The marks along a band `length` px long whose 0 stands at `originPx` (px
 * from the band's start), at `pxPerUnit` screen px per unit. Divisions too
 * close to tell apart are left out, and only every 1, 2, 5 or 10 units is
 * numbered when the numbers would crowd; the numbers count away from 0 both
 * ways, as PowerPoint's do.
 */
export function presentationRulerTicks(
  unit: PresentationRulerUnit,
  originPx: number,
  pxPerUnit: number,
  length: number,
): PresentationRulerTick[] {
  if (!(pxPerUnit > 0) || !(length > 0) || !Number.isFinite(originPx)) return [];
  const divisions = DIVISIONS[unit];
  let step = 1;
  while (step < divisions && (pxPerUnit / divisions) * step < MIN_TICK_GAP_PX) step *= 2;
  const labelEvery = [1, 2, 5, 10, 20, 50].find((n) => n * pxPerUnit >= MIN_LABEL_GAP_PX) ?? 100;
  const ticks: PresentationRulerTick[] = [];
  const first = Math.ceil(((0 - originPx) / pxPerUnit) * divisions);
  const last = Math.floor(((length - originPx) / pxPerUnit) * divisions);
  for (let index = first; index <= last; index += 1) {
    if (index % step !== 0) continue;
    const pos = originPx + (index / divisions) * pxPerUnit;
    if (index % divisions === 0) {
      const units = Math.abs(index / divisions);
      ticks.push({
        pos,
        kind: "unit",
        ...(units % labelEvery === 0 && units !== 0 ? { label: String(units) } : {}),
      });
    } else if ((index * 2) % divisions === 0) {
      ticks.push({ pos, kind: "half" });
    } else {
      ticks.push({ pos, kind: "minor" });
    }
  }
  return ticks;
}

export interface PresentationRulerPaint {
  orientation: "horizontal" | "vertical";
  /** The band's length and breadth, CSS px. */
  length: number;
  breadth: number;
  unit: PresentationRulerUnit;
  /** Where 0 stands along the band, and screen px per unit. */
  originPx: number;
  pxPerUnit: number;
  /** The white stretch: the slide, or the text box being edited. */
  extent: { start: number; end: number };
  /** Objects being moved or sized: their extents, marked on the scale. */
  marks?: ReadonlyArray<{ start: number; end: number }>;
  /** The paragraph's indents in the box being edited (horizontal band). */
  indents?: { firstLine: number; left: number; right: number } | null;
  /** The box's tab stops (horizontal band). */
  tabs?: ReadonlyArray<{ pos: number; align: PresentationTabAlign; ghost?: boolean }>;
}

const BAND = "#eef1f4";
const PAPER = "#ffffff";
const INK = "#475569";
const MARK = "rgba(30, 64, 175, 0.22)";
const MARK_EDGE = "#1e40af";
const INDENT = "#334155";

/**
 * Paint one band into `canvas` (sized to the band at the screen's pixel
 * ratio). Everything is drawn along the band's own axis; a vertical band is
 * the horizontal drawing transposed, its numbers read upright.
 */
export function drawPresentationRuler(
  canvas: HTMLCanvasElement,
  paint: PresentationRulerPaint,
): void {
  const ratio = typeof window === "undefined" ? 1 : window.devicePixelRatio || 1;
  const horizontal = paint.orientation === "horizontal";
  const width = Math.max(0, Math.round(horizontal ? paint.length : paint.breadth));
  const height = Math.max(0, Math.round(horizontal ? paint.breadth : paint.length));
  const pixelWidth = Math.round(width * ratio);
  const pixelHeight = Math.round(height * ratio);
  if (canvas.width !== pixelWidth) canvas.width = pixelWidth;
  if (canvas.height !== pixelHeight) canvas.height = pixelHeight;
  canvas.style.width = `${width}px`;
  canvas.style.height = `${height}px`;
  const context = canvas.getContext("2d");
  if (!context) return;
  context.setTransform(ratio, 0, 0, ratio, 0, 0);
  context.clearRect(0, 0, width, height);
  // Draw along x with the band's breadth along y; a vertical band is the
  // same drawing transposed (its marks stand at its right, by the slide).
  if (!horizontal) context.transform(0, 1, 1, 0, 0, 0);
  const length = paint.length;
  const breadth = paint.breadth;
  context.fillStyle = BAND;
  context.fillRect(0, 0, length, breadth);
  const start = Math.max(0, paint.extent.start);
  const end = Math.min(length, paint.extent.end);
  if (end > start) {
    context.fillStyle = PAPER;
    context.fillRect(start, 2, end - start, breadth - 4);
  }
  for (const mark of paint.marks ?? []) {
    const a = Math.max(0, mark.start);
    const b = Math.min(length, mark.end);
    if (b <= a) continue;
    context.fillStyle = MARK;
    context.fillRect(a, 2, b - a, breadth - 4);
    context.fillStyle = MARK_EDGE;
    context.fillRect(Math.round(a), 0, 1, breadth);
    context.fillRect(Math.round(b) - 1, 0, 1, breadth);
  }
  context.fillStyle = INK;
  context.font = "9px system-ui, -apple-system, 'Segoe UI', sans-serif";
  context.textAlign = "center";
  context.textBaseline = "middle";
  for (const tick of presentationRulerTicks(paint.unit, paint.originPx, paint.pxPerUnit, length)) {
    const x = Math.round(tick.pos) + 0.5;
    if (tick.label) {
      if (horizontal) {
        context.fillText(tick.label, x, breadth / 2);
      } else {
        // the number upright on the transposed band
        context.save();
        context.translate(x, breadth / 2);
        context.transform(0, 1, 1, 0, 0, 0);
        context.fillText(tick.label, 0, 0);
        context.restore();
      }
      continue;
    }
    const size = tick.kind === "unit" ? breadth * 0.5 : tick.kind === "half" ? breadth * 0.32 : breadth * 0.18;
    context.fillRect(x - 0.5, breadth - 2 - size, 1, size);
  }
  if (paint.indents) {
    const { firstLine, left, right } = paint.indents;
    context.fillStyle = INDENT;
    // First-line indent: a triangle pointing down, at the top edge.
    triangle(context, firstLine, 1, 4, 5, "down");
    // Hanging indent: a triangle pointing up, over the left indent's box.
    triangle(context, left, breadth - 4, 4, 5, "up");
    context.fillRect(left - 4, breadth - 4, 8, 3);
    void right;
  }
  for (const tab of paint.tabs ?? []) {
    drawTabGlyph(context, Math.round(tab.pos) + 0.5, breadth - 3, tab.align, tab.ghost ? 0.4 : 1);
  }
}

function triangle(
  context: CanvasRenderingContext2D,
  x: number,
  y: number,
  halfWidth: number,
  height: number,
  direction: "up" | "down",
): void {
  context.beginPath();
  if (direction === "down") {
    context.moveTo(x - halfWidth, y);
    context.lineTo(x + halfWidth, y);
    context.lineTo(x, y + height);
  } else {
    context.moveTo(x - halfWidth, y);
    context.lineTo(x + halfWidth, y);
    context.lineTo(x, y - height);
  }
  context.closePath();
  context.fill();
}

/**
 * A tab stop's glyph as Office draws it: an L for a left stop, an inverted T
 * for a centre stop, a reversed L for a right stop, the T with a dot for a
 * decimal stop. `y` is the glyph's foot.
 */
export function drawTabGlyph(
  context: CanvasRenderingContext2D,
  x: number,
  y: number,
  align: PresentationTabAlign,
  alpha = 1,
): void {
  context.save();
  context.globalAlpha = alpha;
  context.fillStyle = INDENT;
  context.fillRect(x - 0.75, y - 6, 1.5, 6);
  if (align === "left") context.fillRect(x - 0.75, y - 1.5, 5, 1.5);
  else if (align === "right") context.fillRect(x - 4.25, y - 1.5, 5, 1.5);
  else context.fillRect(x - 3.75, y - 1.5, 7.5, 1.5);
  if (align === "decimal") context.fillRect(x + 2, y - 4, 1.5, 1.5);
  context.restore();
}

