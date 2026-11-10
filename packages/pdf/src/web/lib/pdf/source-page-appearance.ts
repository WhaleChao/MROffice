/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The characters a PDF page draws, each with its own origin, box and colour,
 * read from the reader's own engine (PDF.js): the page's operator list is
 * walked with the text state PDF.js draws with, so a character sits exactly
 * where the reader shows it and where the reader's text items say it starts.
 *
 * Read-only. Nothing here writes a PDF or grants a writer any authority.
 */

export interface PdfPageSourceAppearance {
  readonly sourceSha256: string;
  readonly workingRevision: number;
  readonly pageNumber: number;
  /** Native PDF coordinates to the coordinates of `origin` / `quad`. */
  readonly pageTransform: readonly number[];
  readonly text: string;
  readonly characters: readonly Readonly<{
    text: string; start: number; end: number;
    origin: readonly number[];
    /** Upper-left, upper-right, lower-left, lower-right, on the glyph's own axes. */
    quad: readonly number[];
    fontName: string; fontSize: number;
    /** Opaque RGB in 0..1, or empty when the text is not painted one colour. */
    color: readonly number[];
    bidi: number; direction: readonly number[];
  }>[];
}

export interface PdfSourceOperatorList {
  readonly fnArray: ArrayLike<number>;
  readonly argsArray: ArrayLike<unknown>;
}

export interface PdfSourceFont {
  readonly name?: string;
  readonly loadedName?: string;
  readonly ascent?: number;
  readonly descent?: number;
  readonly vertical?: boolean;
  readonly fontMatrix?: ArrayLike<number>;
  readonly defaultVMetrics?: ArrayLike<number>;
}

type Matrix = [number, number, number, number, number, number];
interface State {
  ctm: Matrix; tm: Matrix; x: number; y: number; lineX: number; lineY: number;
  font: string; size: number; direction: 1 | -1;
  charSpacing: number; wordSpacing: number; hscale: number; rise: number; leading: number; mode: number;
  fill: readonly number[]; stroke: readonly number[];
}

const MAX_OPERATORS = 2_000_000;
const MAX_CHARACTERS = 400_000;
const identity = (): Matrix => [1, 0, 0, 1, 0, 0];
const finite = (value: unknown): value is number => typeof value === "number" && Number.isFinite(value);
/** `left` applied first, then `right` (PDF row-vector order: left × right). */
const concat = (left: ArrayLike<number>, right: ArrayLike<number>): Matrix => [
  left[0] * right[0] + left[1] * right[2],
  left[0] * right[1] + left[1] * right[3],
  left[2] * right[0] + left[3] * right[2],
  left[2] * right[1] + left[3] * right[3],
  left[4] * right[0] + left[5] * right[2] + right[4],
  left[4] * right[1] + left[5] * right[3] + right[5],
];
function matrixOf(value: unknown): Matrix | null {
  const list = Array.isArray(value) || ArrayBuffer.isView(value) ? Array.from(value as ArrayLike<unknown>) : null;
  return list?.length === 6 && list.every(finite) ? (list as Matrix) : null;
}
function colorOf(args: readonly unknown[]): readonly number[] {
  const first = args[0];
  if (typeof first === "string" && /^#[0-9a-f]{6}$/i.test(first)) {
    return [1, 3, 5].map((at) => Number.parseInt(first.slice(at, at + 2), 16) / 255);
  }
  if (args.length >= 3 && args.slice(0, 3).every(finite)) {
    return (args.slice(0, 3) as number[]).map((value) => Math.max(0, Math.min(255, value)) / 255);
  }
  return [];
}
const RTL = /[֐-ࣿיִ-﷿ﹰ-ﻼ]|[\ud802-\ud803][\udc00-\udfff]|\ud83a[\udc00-\uddff]|\ud83b[\ude00-\udeff]/;

/**
 * Walk one page's operators and return every character it shows. An operator
 * list this walk cannot follow (unbalanced state, an oversized page) throws:
 * the caller then has no source geometry, which every consumer already
 * treats as "use the plain editing path".
 */
export function pdfPageSourceAppearanceFromOperators(input: {
  sourceSha256: string;
  workingRevision?: number;
  pageNumber: number;
  operators: PdfSourceOperatorList;
  ops: Readonly<Record<string, number>>;
  font: (name: string) => PdfSourceFont | null;
}): PdfPageSourceAppearance {
  const { operators, ops } = input;
  const count = operators.fnArray.length;
  if (count > MAX_OPERATORS || count !== operators.argsArray.length) {
    throw new Error("PDF source appearance exceeds its operator budget.");
  }
  const state: State = {
    ctm: identity(), tm: identity(), x: 0, y: 0, lineX: 0, lineY: 0, font: "", size: 0, direction: 1,
    charSpacing: 0, wordSpacing: 0, hscale: 1, rise: 0, leading: 0, mode: 0, fill: [0, 0, 0], stroke: [0, 0, 0],
  };
  const stack: State[] = [];
  const characters: Array<PdfPageSourceAppearance["characters"][number]> = [];
  let text = "";
  const save = () => {
    if (stack.length >= 4096) throw new Error("PDF source appearance state is too deep.");
    stack.push({ ...state, ctm: [...state.ctm] as Matrix, tm: [...state.tm] as Matrix });
  };
  const restore = () => {
    const previous = stack.pop();
    if (previous) Object.assign(state, previous);
  };
  const move = (x: number, y: number) => {
    state.x = state.lineX += x;
    state.y = state.lineY += y;
  };
  const setFont = (name: unknown, size: unknown) => {
    if (typeof name !== "string" || !finite(size)) return;
    state.font = name;
    state.direction = size < 0 ? -1 : 1;
    state.size = Math.abs(size);
  };
  const show = (glyphs: unknown) => {
    if (!Array.isArray(glyphs)) return;
    const font = input.font(state.font);
    const size = state.size;
    if (!font || size === 0) return;
    const fontMatrix = font.fontMatrix && finite(font.fontMatrix[0]) && font.fontMatrix[0] !== 0 ? font.fontMatrix[0] : 0.001;
    const vertical = font.vertical === true;
    const hscale = state.hscale * state.direction;
    const widthScale = size * fontMatrix;
    const ascent = finite(font.ascent) && font.ascent !== 0 ? font.ascent : 0.8;
    const descent = finite(font.descent) && font.descent !== 0 ? font.descent : -0.2;
    const up = size * state.direction;
    const toPage = concat(state.tm, state.ctm);
    const point = (tx: number, ty: number) => [
      toPage[0] * tx + toPage[2] * ty + toPage[4],
      toPage[1] * tx + toPage[3] * ty + toPage[5],
    ];
    const fontName = font.name ?? font.loadedName ?? state.font;
    const paint = state.mode === 1 || state.mode === 5 ? state.stroke : state.mode === 3 || state.mode === 7 ? [] : state.fill;
    // The size the page shows: the text matrix and CTM scale the nominal one.
    const pageSize = size * Math.hypot(toPage[2], toPage[3]);
    let x = 0;
    for (const glyph of glyphs) {
      if (finite(glyph)) {
        x += (vertical ? 1 : -1) * glyph * size / 1000;
        continue;
      }
      if (!glyph || typeof glyph !== "object") continue;
      const record = glyph as { unicode?: unknown; width?: unknown; isSpace?: unknown; vmetric?: ArrayLike<number> };
      const spacing = (record.isSpace === true ? state.wordSpacing : 0) + state.charSpacing;
      let width = finite(record.width) ? record.width : 0;
      let originX: number, originY: number, across: readonly [number, number], rise: readonly [number, number];
      let top: number, bottom: number;
      if (vertical) {
        const vmetric = record.vmetric ?? font.defaultVMetrics;
        const vx = -(record.vmetric ? record.vmetric[1] : width * 0.5) * widthScale;
        const vy = (vmetric ? vmetric[2] : 0) * widthScale;
        width = vmetric ? -vmetric[0] : width;
        const advance = width * widthScale;
        originX = state.x + vx * hscale;
        originY = state.y + state.rise - (x + vy);
        // A vertical glyph's box runs down the column; its "width" is one em.
        across = [size * hscale, 0];
        rise = [0, 1];
        top = 0;
        bottom = -advance;
        x += advance - spacing * state.direction;
      } else {
        const advance = width * widthScale;
        originX = state.x + x * hscale;
        originY = state.y + state.rise;
        across = [advance * hscale, 0];
        rise = [0, 1];
        top = ascent * up;
        bottom = descent * up;
        x += advance + spacing * state.direction;
      }
      const unicode = typeof record.unicode === "string" ? record.unicode : "";
      const parts = Array.from(unicode);
      if (!parts.length) continue;
      if (characters.length + parts.length > MAX_CHARACTERS) {
        throw new Error("PDF source appearance exceeds its character budget.");
      }
      // One glyph that reads as several characters (a ligature) shares its
      // box between them in equal parts, left to right.
      for (let part = 0; part < parts.length; part += 1) {
        const from = part / parts.length, to = (part + 1) / parts.length;
        const ax = originX + across[0] * from, bx = originX + across[0] * to;
        const quad = [
          ...point(ax + rise[0] * top, originY + rise[1] * top),
          ...point(bx + rise[0] * top, originY + rise[1] * top),
          ...point(ax + rise[0] * bottom, originY + rise[1] * bottom),
          ...point(bx + rise[0] * bottom, originY + rise[1] * bottom),
        ];
        const origin = point(ax, originY);
        const base = point(bx, originY);
        const length = Math.hypot(base[0] - origin[0], base[1] - origin[1]);
        const start = text.length;
        text += parts[part];
        characters.push(Object.freeze({
          text: parts[part], start, end: text.length,
          origin: Object.freeze(origin), quad: Object.freeze(quad),
          fontName, fontSize: pageSize, color: Object.freeze([...paint]),
          bidi: RTL.test(parts[part]) ? 1 : 0,
          direction: Object.freeze(length > 0 ? [(base[0] - origin[0]) / length, (base[1] - origin[1]) / length] : [1, 0]),
        }));
      }
    }
    if (vertical) state.y -= x;
    else state.x += x * hscale;
  };

  for (let index = 0; index < count; index += 1) {
    const op = operators.fnArray[index];
    const raw = operators.argsArray[index];
    const args: readonly unknown[] = Array.isArray(raw) ? raw : raw && ArrayBuffer.isView(raw) ? Array.from(raw as unknown as ArrayLike<unknown>) : [];
    if (op === ops.save) save();
    else if (op === ops.restore) restore();
    else if (op === ops.transform) {
      const matrix = matrixOf(args);
      if (matrix) state.ctm = concat(matrix, state.ctm);
    } else if (op === ops.paintFormXObjectBegin) {
      save();
      const matrix = matrixOf(args[0]);
      if (matrix) state.ctm = concat(matrix, state.ctm);
    } else if (op === ops.paintFormXObjectEnd) restore();
    else if (op === ops.beginAnnotation) save();
    else if (op === ops.endAnnotation) restore();
    else if (op === ops.beginText) {
      state.tm = identity();
      state.x = state.y = state.lineX = state.lineY = 0;
    } else if (op === ops.setTextMatrix) {
      const matrix = matrixOf(args[0]) ?? matrixOf(args);
      if (matrix) {
        state.tm = matrix;
        state.x = state.y = state.lineX = state.lineY = 0;
      }
    } else if (op === ops.moveText) {
      if (finite(args[0]) && finite(args[1])) move(args[0], args[1]);
    } else if (op === ops.setLeadingMoveText) {
      if (finite(args[0]) && finite(args[1])) {
        state.leading = args[1];
        move(args[0], args[1]);
      }
    } else if (op === ops.nextLine) move(0, state.leading);
    else if (op === ops.setLeading) {
      if (finite(args[0])) state.leading = -args[0];
    } else if (op === ops.setFont) setFont(args[0], args[1]);
    else if (op === ops.setCharSpacing) {
      if (finite(args[0])) state.charSpacing = args[0];
    } else if (op === ops.setWordSpacing) {
      if (finite(args[0])) state.wordSpacing = args[0];
    } else if (op === ops.setHScale) {
      if (finite(args[0])) state.hscale = args[0] / 100;
    } else if (op === ops.setTextRise) {
      if (finite(args[0])) state.rise = args[0];
    } else if (op === ops.setTextRenderingMode) {
      if (finite(args[0])) state.mode = args[0];
    } else if (op === ops.setGState) {
      if (Array.isArray(args[0])) {
        for (const entry of args[0]) {
          if (Array.isArray(entry) && entry[0] === "Font" && Array.isArray(entry[1])) setFont(entry[1][0], entry[1][1]);
        }
      }
    } else if (op === ops.setFillRGBColor) state.fill = colorOf(args);
    else if (op === ops.setStrokeRGBColor) state.stroke = colorOf(args);
    else if (op === ops.setFillColorN || op === ops.setFillTransparent) state.fill = [];
    else if (op === ops.setStrokeColorN || op === ops.setStrokeTransparent) state.stroke = [];
    else if (op === ops.showText || op === ops.showSpacedText) show(args[0]);
    else if (op === ops.nextLineShowText) {
      move(0, state.leading);
      show(args[0]);
    } else if (op === ops.nextLineSetSpacingShowText) {
      if (finite(args[0])) state.wordSpacing = args[0];
      if (finite(args[1])) state.charSpacing = args[1];
      move(0, state.leading);
      show(args[2]);
    }
  }
  return Object.freeze({
    sourceSha256: input.sourceSha256,
    workingRevision: input.workingRevision ?? 0,
    pageNumber: input.pageNumber,
    pageTransform: Object.freeze(identity()),
    text,
    characters: Object.freeze(characters),
  });
}
