/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type {
  PdfPositionedTextLayoutLine,
  PdfTextLayoutRun,
} from "./text-run-layout";

/** Authoring layout, NEVER a source/selection proof or permission to write.
 * A receiver must separately authenticate the existing resource/box transaction.
 * Offsets are UTF-16 so soft wraps never alter the original string. */
export type PdfPortableTextLine = readonly [
  start: number,
  end: number,
  left: number,
  baseline: number,
  advance: number,
  contentAdvance: number,
];
export interface PdfPortableTextLayout {
  readonly version: 1;
  readonly textSha256: string;
  readonly width: number;
  readonly height: number;
  readonly fontSize: number;
  readonly lines: readonly PdfPortableTextLine[];
}
export const PDF_PORTABLE_LAYOUT_MAX_LINES = 2_000;
export const PDF_PORTABLE_LAYOUT_POINT_TOLERANCE = 0.01;



const BROWSER_LAYOUT_UNIT = 1 / 64;

/** Punctuation that may not open a line and so hangs past a CJK row's end
 * instead (Word's and WPS's 允许标点溢出边界, on by default). */
const HANGING_PUNCTUATION = /[、。，．：；！？）］】」』〉》〕｝…‥]$/u;

/** How far a row may reach past its frame: by one em when it ends with
 * punctuation that hangs (a line set in a wider face kept 「、」 at its end
 * 1 pt past the frame, and the save of every such edit was refused). Mirror
 * of the server's hanging_allowance. */
export function pdfPortableHangingAllowance(row: string, fontSize: number): number {
  return HANGING_PUNCTUATION.test(row.trimEnd()) ? fontSize : 0;
}
const keys = new Set([
  "version",
  "textSha256",
  "width",
  "height",
  "fontSize",
  "lines",
]);
const scalar = (value: unknown): value is number =>
  typeof value === "number" && Number.isFinite(value);
const boundary = (text: string, offset: number) =>
  !(
    offset > 0 &&
    offset < text.length &&
    text.charCodeAt(offset - 1) >= 0xd800 &&
    text.charCodeAt(offset - 1) <= 0xdbff &&
    text.charCodeAt(offset) >= 0xdc00 &&
    text.charCodeAt(offset) <= 0xdfff
  );

/** One column's left edge to the next's, for a frame this wide in `count`
 * columns `gap` apart (the writer's column_step). */
export function pdfPortableColumnStep(width: number, count: number, gap: number): number {
  return (width - (count - 1) * gap) / count + gap;
}

/** Decode stored authoring data synchronously, not a verified text identity.
 * Like a parsed source binding, this must be revalidated by the native writer;
 * only createPdfPortableTextLayout computes the current text digest here. */
export function readUnverifiedPdfPortableTextLayout(
  value: unknown,
  text: string,
  options: {
    /** A box set in runs: its first baseline may sit at its smallest run's
     * size (this fraction of the layout's font size), not the box's. */
    minimumBaselineScale?: number;
    /** A box in columns: their count and the spacing between them (ems of
     * the layout's font size). */
    columns?: Readonly<{ count: number; gapEm: number }>;
  } = {},
): PdfPortableTextLayout | null {
  if (!value || typeof value !== "object") return null;
  const data = value as Record<string, unknown>;
  if (
    typeof data.textSha256 !== "string" ||
    !scalar(data.width) ||
    !scalar(data.height) ||
    !scalar(data.fontSize)
  )
    return null;
  return readPdfPortableTextLayout(value, {
    text,
    textSha256: data.textSha256,
    width: data.width,
    height: data.height,
    fontSize: data.fontSize,
    ...(options.columns && options.columns.count > 1
      ? { columnStep: pdfPortableColumnStep(data.width, options.columns.count, options.columns.gapEm * data.fontSize) }
      : {}),
    ...(options.minimumBaselineScale !== undefined
      ? { minimumBaseline: data.fontSize * options.minimumBaselineScale }
      : {}),
  });
}

/** Capture positions actually used for the appearance, without rediscovering
 * lines by matching bounds, repeated strings or a second wrapping algorithm. */
export async function createPdfPortableTextLayout<
  Run extends PdfTextLayoutRun,
>(input: {
  text: string;
  width: number;
  height: number;
  fontSize: number;
  align: "left" | "center" | "right";
  positionedLines: readonly (PdfPositionedTextLayoutLine<Run> & {
    contentAdvance?: number;
    /** The row's own baseline (a box set in runs: below its largest run). */
    baseline?: number;
  })[];
  /** A box set in runs: the smallest run's size, its lowest first baseline. */
  minimumBaseline?: number;
  /** A box in columns: one column's left edge to the next's. */
  columnStep?: number;
}): Promise<PdfPortableTextLayout | null> {
  if (
    input.text.length > 20_000 ||
    input.positionedLines.length > PDF_PORTABLE_LAYOUT_MAX_LINES
  )
    return null;
  const subtle = globalThis.crypto?.subtle;
  if (!subtle || /[\uD800-\uDFFF]/u.test(input.text)) return null;
  const digest = await subtle.digest(
    "SHA-256",
    new TextEncoder().encode(input.text),
  );
  const textSha256 = Array.from(new Uint8Array(digest), (byte) =>
    byte.toString(16).padStart(2, "0"),
  ).join("");
  let offset = 0;
  const lines: PdfPortableTextLine[] = [];
  for (const positioned of input.positionedLines) {
    const text = positioned.line.segments
      .map((segment) => segment.text)
      .join("");
    while (
      /[\r\n]/.test(input.text[offset] ?? "") &&
      offset < input.text.length
    )
      offset++;
    if (input.text.slice(offset, offset + text.length) !== text) {
      if (process.env.NODE_ENV === "development") {
        try { console.warn("[pdf-text-box] portable row text differs", { row: text.slice(0, 40), source: input.text.slice(offset, offset + text.length).slice(0, 40) }); } catch { /* never throws */ }
      }
      return null;
    }
    let advance = positioned.line.width * positioned.horizontalScale;
    let contentAdvance = positioned.contentAdvance ?? advance;
    const spare = Math.max(0, positioned.availableWidth - contentAdvance);
    const left =
      positioned.left +
      (input.align === "center"
        ? spare / 2
        : input.align === "right"
          ? spare
          : 0);
    // A row the browser laid out in the box that reads past it by no more
    // than its measuring unit is that unit's rounding, not overflow: it is
    // held to the box, as the browser holds it. A box exactly as wide as
    
    // the rounding, and the move was lost.
    const excess = left + contentAdvance - input.width;
    if (excess > PDF_PORTABLE_LAYOUT_POINT_TOLERANCE && excess <= BROWSER_LAYOUT_UNIT) {
      contentAdvance -= excess;
      advance = Math.max(contentAdvance, advance - excess);
    }
    lines.push([
      offset,
      offset + text.length,
      left,
      positioned.baseline ?? positioned.top + input.fontSize,
      advance,
      contentAdvance,
    ]);
    offset += text.length;
  }
  const expected = {
    text: input.text,
    textSha256,
    width: input.width,
    height: input.height,
    fontSize: input.fontSize,
    ...(input.minimumBaseline !== undefined ? { minimumBaseline: input.minimumBaseline } : {}),
    ...(input.columnStep !== undefined ? { columnStep: input.columnStep } : {}),
  };
  return readPdfPortableTextLayout(
    {
      version: 1,
      textSha256,
      width: input.width,
      height: input.height,
      fontSize: input.fontSize,
      lines,
    },
    expected,
  );
}

/** Structural validation shared by serialization/reopen. Digest is supplied
 * from the exact current text, not trusted from the received layout itself. */
export function readPdfPortableTextLayout(
  value: unknown,
  expected: {
    text: string;
    textSha256: string;
    width: number;
    height: number;
    fontSize: number;
    /** The lowest a first baseline may sit; the font size by default. */
    minimumBaseline?: number;
    /** A box in columns (Word's and PowerPoint's Columns): one column's left
     * edge to the next's. The rows fill a column to the frame's foot before
     * the next: a row starts again at the top only where it opens a later
     * column (mirror of the writer's read_portable_text_layout). */
    columnStep?: number;
  },
): PdfPortableTextLayout | null {
  if (/[\uD800-\uDFFF]/u.test(expected.text)) return null;
  const minimumBaseline = expected.minimumBaseline ?? expected.fontSize;
  if (!scalar(minimumBaseline) || minimumBaseline <= 0) return null;
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const data = value as Record<string, unknown>;
  if (
    Object.keys(data).length !== keys.size ||
    Object.keys(data).some((key) => !keys.has(key)) ||
    data.version !== 1 ||
    data.textSha256 !== expected.textSha256 ||
    !/^[a-f0-9]{64}$/.test(expected.textSha256) ||
    !Array.isArray(data.lines) ||
    !data.lines.length ||
    data.lines.length > PDF_PORTABLE_LAYOUT_MAX_LINES ||
    expected.text.length > 20_000
  )
    return null;
  for (const key of ["width", "height", "fontSize"] as const) {
    if (
      !scalar(data[key]) ||
      !scalar(expected[key]) ||
      expected[key] <= 0 ||
      expected[key] > (key === "fontSize" ? 512 : 1_000_000) ||
      Math.abs(data[key] - expected[key]) > PDF_PORTABLE_LAYOUT_POINT_TOLERANCE
    )
      return null;
  }
  const step = expected.columnStep;
  if (step !== undefined && (!scalar(step) || step <= 0 || step > expected.width + PDF_PORTABLE_LAYOUT_POINT_TOLERANCE)) return null;
  let offset = 0,
    baseline = -Infinity,
    column = 0;
  const lines: PdfPortableTextLine[] = [];
  for (const raw of data.lines) {
    if (!Array.isArray(raw) || raw.length !== 6 || !raw.every(scalar))
      return null;
    const [start, end, left, y, advance, contentAdvance] = raw as number[];
    // The column a row stands in, and whether it opens a later one.
    const rowColumn = step !== undefined ? Math.floor((left + PDF_PORTABLE_LAYOUT_POINT_TOLERANCE) / step) : 0;
    if (rowColumn < column) return null;
    const opensColumn = rowColumn > column;
    column = rowColumn;
    if (process.env.NODE_ENV === "development") {
      // Development only: which rule a row broke (never throws).
      try {
        const row = expected.text.slice(start, end);
        const broke = left + contentAdvance > expected.width + PDF_PORTABLE_LAYOUT_POINT_TOLERANCE +
          pdfPortableHangingAllowance(row, expected.fontSize) ? "wider than the frame"
          : y > expected.height + PDF_PORTABLE_LAYOUT_POINT_TOLERANCE ? "below the frame"
          : y <= baseline && step === undefined ? "baseline not below the previous"
          : y < minimumBaseline - PDF_PORTABLE_LAYOUT_POINT_TOLERANCE ? "first baseline too high"
          : /[\r\n]/.test(row) ? "row holds a line break"
          : contentAdvance > advance + PDF_PORTABLE_LAYOUT_POINT_TOLERANCE ? "content wider than advance"
          : (!/\s$/u.test(row) && Math.abs(advance - contentAdvance) > PDF_PORTABLE_LAYOUT_POINT_TOLERANCE) ? "advance differs from content"
          : null;
        if (broke) console.warn("[pdf-text-box] portable row", broke, { row: row.slice(0, 40), left, y, advance, contentAdvance, width: expected.width, height: expected.height });
      } catch { /* never throws */ }
    }
    if (
      !Number.isSafeInteger(start) ||
      !Number.isSafeInteger(end) ||
      start < offset ||
      end < start ||
      end > expected.text.length ||
      !boundary(expected.text, start) ||
      !boundary(expected.text, end) ||
      !/^[\r\n]*$/.test(expected.text.slice(offset, start)) ||
      /[\r\n]/.test(expected.text.slice(start, end)) ||
      (expected.text.slice(start, end).trim().length > 0 &&
        contentAdvance <= 0) ||
      left < 0 ||
      advance < 0 ||
      advance > 1_000_000 ||
      contentAdvance < 0 ||
      contentAdvance > advance + PDF_PORTABLE_LAYOUT_POINT_TOLERANCE ||
      (!/\s$/u.test(expected.text.slice(start, end)) &&
        Math.abs(advance - contentAdvance) >
          PDF_PORTABLE_LAYOUT_POINT_TOLERANCE) ||
      left + contentAdvance >
        expected.width + PDF_PORTABLE_LAYOUT_POINT_TOLERANCE +
          pdfPortableHangingAllowance(expected.text.slice(start, end), expected.fontSize) ||
      (y <= baseline && !opensColumn) ||
      y < minimumBaseline - PDF_PORTABLE_LAYOUT_POINT_TOLERANCE ||
      y > expected.height + PDF_PORTABLE_LAYOUT_POINT_TOLERANCE
    )
      return null;
    lines.push(Object.freeze([start, end, left, y, advance, contentAdvance]));
    offset = end;
    baseline = y;
  }
  if (!/^[\r\n]*$/.test(expected.text.slice(offset))) return null;
  return Object.freeze({
    version: 1,
    textSha256: expected.textSha256,
    width: expected.width,
    height: expected.height,
    fontSize: expected.fontSize,
    lines: Object.freeze(lines),
  });
}
