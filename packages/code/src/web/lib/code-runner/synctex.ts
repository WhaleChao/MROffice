/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * SyncTeX: the map a LaTeX compile writes (main.synctex.gz, returned with the
 * receipt by lib/code-runner/output-files) between source lines and places on
 * the PDF's pages, read in the browser for Overleaf's two jumps: from a line
 * of code to where it is typeset (forward), and from a point on a page to the
 * line that typeset it (inverse).
 *
 * The format is the one synctex(5) describes. A preamble names every file
 * TeX read (`Input:<tag>:<path>`, more of them may follow in the content) and
 * the scale (`Unit:`, `Magnification:`, `X Offset:`, `Y Offset:`, which a
 * `Post scriptum:` section may restate). The content is one `{<page>` ...
 * `}<page>` sheet per page holding nested boxes: `[` vbox `]`, `(` hbox `)`,
 * void boxes `v` and `h`, and the records inside a box: `x` (current
 * position), `k` (kern), `g` (glue) and `$` (math). Each carries a link
 * `<tag>,<line>[,<column>]` and a point `<h>,<v>` (`=` repeats the previous
 * v); boxes add `<width>,<height>,<depth>`. A box's point is where its
 * baseline starts: it spans h..h+width across and v-height..v+depth down.
 * Values are TeX's scaled points times Unit, measured from the page's top
 * left. LuaTeX measures a vbox from its top instead, so only hboxes are used.
 *
 * Coordinates handed out here are PDF points (big points) from the page's top
 * left. The parsed map lives in memory only, cached per received file.
 */

import { compiledProjectPath } from "./latex-project";

/** Scaled points per big point (65536 sp = 1 pt, 72.27 pt = 72 bp). */
const SCALED_POINTS_PER_BIG_POINT = 65_781.76;
/** A gzip bomb's ceiling: real maps of long theses stay far below it. */
const MAX_SYNCTEX_TEXT_BYTES = 64 * 1024 * 1024;
/** How far from the asked-for line a forward search looks for one typeset. */
const MAX_FORWARD_LINE_DISTANCE = 20;
const MAX_FORWARD_RECTS = 64;

export interface SynctexRect {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

interface SynctexBox {
  readonly tag: number;
  readonly line: number;
  readonly h: number;
  readonly v: number;
  readonly width: number;
  readonly height: number;
  readonly depth: number;
  /** The records directly inside it: tag, line and h, three per record. */
  readonly records: number[];
}

interface SynctexSheet {
  readonly page: number;
  readonly boxes: SynctexBox[];
}

export interface SynctexMap {
  /** Input tag to the path TeX read it as. */
  readonly inputs: ReadonlyMap<number, string>;
  readonly sheets: readonly SynctexSheet[];
  readonly unit: number;
  readonly magnification: number;
  readonly xOffset: number;
  readonly yOffset: number;
}

export interface SynctexForwardResult {
  readonly page: number;
  readonly rects: readonly SynctexRect[];
}

export interface SynctexInverseResult {
  /** The project path ("chapters/intro.tex", "main.tex"). */
  readonly file: string;
  readonly line: number;
}

/** Inflate gzip bytes with the browser's own DecompressionStream. */
export async function gunzip(bytes: Uint8Array): Promise<Uint8Array> {
  if (typeof DecompressionStream === "undefined") {
    throw new Error("The SyncTeX file can't be read here.");
  }
  const input = new Uint8Array(bytes.byteLength);
  input.set(bytes);
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(input);
      controller.close();
    },
  }).pipeThrough(new DecompressionStream("gzip") as ReadableWritablePair<Uint8Array, Uint8Array>);
  const reader = stream.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > MAX_SYNCTEX_TEXT_BYTES) {
      await reader.cancel();
      throw new Error("The SyncTeX file is too large to read.");
    }
    chunks.push(value);
  }
  const output = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return output;
}

/** Reads the integers of a record body: `1,12:4736286,=:0,0,0`. */
function readNumbers(body: string, into: number[], lastV: number): number {
  into.length = 0;
  let value = 0;
  let sign = 1;
  let digits = false;
  for (let index = 0; index < body.length; index += 1) {
    const code = body.charCodeAt(index);
    if (code >= 48 && code <= 57) {
      value = value * 10 + (code - 48);
      digits = true;
    } else if (code === 45 && !digits) {
      sign = -1;
    } else if (code === 61) {
      // "=": the v of the previous full point.
      into.push(lastV);
      digits = false;
      value = 0;
      sign = 1;
    } else {
      if (digits) into.push(sign * value);
      value = 0;
      sign = 1;
      digits = false;
      if (code === 58) into.push(Number.NaN); // ":" separates link, point, size
    }
  }
  if (digits) into.push(sign * value);
  return into.length;
}

/**
 * Split a record into link, point and size, each as numbers. `numbers` holds
 * them with NaN between the parts.
 */
function recordParts(numbers: readonly number[]): number[][] {
  const parts: number[][] = [[]];
  for (const value of numbers) {
    if (Number.isNaN(value)) parts.push([]);
    else parts[parts.length - 1]!.push(value);
  }
  return parts;
}

function preambleNumber(line: string, label: string): number | null {
  if (!line.startsWith(label)) return null;
  const value = Number(line.slice(label.length).trim());
  return Number.isFinite(value) ? value : null;
}

/** Parse the text of a .synctex file. */
export function parseSynctex(text: string): SynctexMap {
  const inputs = new Map<number, string>();
  const sheets: SynctexSheet[] = [];
  let unit = 1;
  let magnification = 1000;
  let xOffset = 0;
  let yOffset = 0;
  let postScriptum = false;
  let sheet: SynctexSheet | null = null;
  // The open boxes of the current sheet; null for a vbox (records in it are
  // not on a line of text).
  const stack: (SynctexBox | null)[] = [];
  let formDepth = 0;
  let lastV = 0;
  const numbers: number[] = [];

  let start = 0;
  while (start < text.length) {
    let end = text.indexOf("\n", start);
    if (end < 0) end = text.length;
    const line = text.slice(start, end);
    start = end + 1;
    if (line.length === 0) continue;
    const kind = line.charCodeAt(0);

    if (sheet === null || formDepth > 0) {
      if (line.startsWith("Input:")) {
        const separator = line.indexOf(":", 6);
        const tag = Number(line.slice(6, separator));
        if (separator > 6 && Number.isSafeInteger(tag)) {
          inputs.set(tag, line.slice(separator + 1));
        }
        continue;
      }
      if (line.startsWith("Post scriptum:")) {
        postScriptum = true;
        continue;
      }
      const mag = preambleNumber(line, "Magnification:");
      if (mag !== null && mag > 0) {
        magnification = mag;
        continue;
      }
      const scale = postScriptum ? null : preambleNumber(line, "Unit:");
      if (scale !== null && scale > 0) {
        unit = scale;
        continue;
      }
      const x = preambleNumber(line, "X Offset:");
      if (x !== null) {
        xOffset = x;
        continue;
      }
      const y = preambleNumber(line, "Y Offset:");
      if (y !== null) {
        yOffset = y;
        continue;
      }
      if (kind === 60) {
        // "<": a form (pdfTeX's \pdfxform), drawn wherever it is used.
        formDepth += 1;
        continue;
      }
      if (kind === 62 && formDepth > 0) {
        formDepth -= 1;
        continue;
      }
      if (kind === 123 && formDepth === 0) {
        // "{<page>"
        const page = Number(line.slice(1));
        if (Number.isSafeInteger(page) && page > 0) {
          sheet = { page, boxes: [] };
          stack.length = 0;
        }
      }
      continue;
    }

    switch (kind) {
      case 73: {
        // "Input:" for a file TeX opened while this page was being written.
        const separator = line.startsWith("Input:") ? line.indexOf(":", 6) : -1;
        const tag = Number(line.slice(6, separator));
        if (separator > 6 && Number.isSafeInteger(tag)) {
          inputs.set(tag, line.slice(separator + 1));
        }
        break;
      }
      case 125: // "}<page>"
        sheets.push(sheet);
        sheet = null;
        stack.length = 0;
        break;
      case 60: // "<" a form nested in a sheet
        formDepth += 1;
        break;
      case 40: // "(" hbox
      case 104: {
        // "h" void hbox
        readNumbers(line.slice(1), numbers, lastV);
        const [link = [], point = [], size = []] = recordParts(numbers);
        if (link.length < 2 || point.length < 2) {
          if (kind === 40) stack.push(null);
          break;
        }
        lastV = point[1]!;
        const box: SynctexBox = {
          tag: link[0]!,
          line: link[1]!,
          h: point[0]!,
          v: point[1]!,
          width: size[0] ?? 0,
          height: size[1] ?? 0,
          depth: size[2] ?? 0,
          records: [],
        };
        sheet.boxes.push(box);
        if (kind === 40) stack.push(box);
        break;
      }
      case 91: // "[" vbox
        stack.push(null);
        break;
      case 41: // ")"
      case 93: // "]"
        stack.pop();
        break;
      case 120: // "x"
      case 107: // "k"
      case 103: // "g"
      case 36: {
        // "$"
        readNumbers(line.slice(1), numbers, lastV);
        const [link = [], point = []] = recordParts(numbers);
        if (link.length < 2 || point.length < 2) break;
        lastV = point[1]!;
        const parent = stack[stack.length - 1];
        if (parent) parent.records.push(link[0]!, link[1]!, point[0]!);
        break;
      }
      default:
        // "v" void vboxes, "f" form references and "!" byte offsets.
        break;
    }
  }
  return { inputs, sheets, unit, magnification, xOffset, yOffset };
}

const parsedMaps = new WeakMap<Uint8Array, Promise<SynctexMap | null>>();

/**
 * The map in a received main.synctex.gz, parsed once per file and kept while
 * the file is (null when it cannot be read).
 */
export function loadSynctexMap(bytes: Uint8Array): Promise<SynctexMap | null> {
  let parsed = parsedMaps.get(bytes);
  if (!parsed) {
    parsed = gunzip(bytes)
      .then((text) => parseSynctex(new TextDecoder("utf-8", { fatal: false }).decode(text)))
      .catch(() => null);
    parsedMaps.set(bytes, parsed);
  }
  return parsed;
}

function bigPointsPerUnit(map: SynctexMap): number {
  return (map.unit * (map.magnification / 1000)) / SCALED_POINTS_PER_BIG_POINT;
}

function offsets(map: SynctexMap): { x: number; y: number } {
  const scale = map.unit / SCALED_POINTS_PER_BIG_POINT;
  return { x: map.xOffset * scale, y: map.yOffset * scale };
}

/** The box's extent in PDF points from the page's top left. */
function boxRect(map: SynctexMap, box: SynctexBox): SynctexRect {
  const scale = bigPointsPerUnit(map);
  const offset = offsets(map);
  let left = box.h;
  let right = box.h + box.width;
  // A record past the box's own width (an \hbox to 0pt) still shows on it.
  for (let index = 2; index < box.records.length; index += 3) {
    const h = box.records[index]!;
    if (h < left) left = h;
    if (h > right) right = h;
  }
  return {
    left: left * scale + offset.x,
    right: right * scale + offset.x,
    top: (box.v - box.height) * scale + offset.y,
    bottom: (box.v + box.depth) * scale + offset.y,
  };
}

/** The input tags whose path is `projectPath` in the project. */
function tagsOf(map: SynctexMap, projectPath: string): Set<number> {
  const tags = new Set<number>();
  for (const [tag, path] of map.inputs) {
    if (compiledProjectPath(path) === projectPath) tags.add(tag);
  }
  return tags;
}

/** Whether TeX read the project file `projectPath` for this document. */
export function synctexHasInput(map: SynctexMap, projectPath: string): boolean {
  return tagsOf(map, projectPath).size > 0;
}

/**
 * Tag, then line, to the boxes typeset from it (sheet and box index, two per
 * box), for the project's files only. Built on the first forward search.
 */
type SynctexLineIndex = Map<number, Map<number, number[]>>;
const lineIndexes = new WeakMap<SynctexMap, SynctexLineIndex>();

function lineIndex(map: SynctexMap): SynctexLineIndex {
  const cached = lineIndexes.get(map);
  if (cached) return cached;
  const projectTags = new Set<number>();
  for (const [tag, path] of map.inputs) {
    if (compiledProjectPath(path) !== null) projectTags.add(tag);
  }
  const index: SynctexLineIndex = new Map();
  const add = (tag: number, line: number, sheet: number, box: number) => {
    if (!projectTags.has(tag)) return;
    let lines = index.get(tag);
    if (!lines) index.set(tag, (lines = new Map()));
    let boxes = lines.get(line);
    if (!boxes) lines.set(line, (boxes = []));
    if (boxes[boxes.length - 2] !== sheet || boxes[boxes.length - 1] !== box) {
      boxes.push(sheet, box);
    }
  };
  map.sheets.forEach((sheet, sheetIndex) => {
    sheet.boxes.forEach((box, boxIndex) => {
      // A line of text answers for the records in it; a box with nothing
      // inside answers for its own link when it shows something (a picture,
      // a rule), not when it is an empty strut or anchor.
      if (
        box.records.length === 0 &&
        box.width > 0 &&
        box.height + box.depth > 0
      ) {
        add(box.tag, box.line, sheetIndex, boxIndex);
      }
      for (let record = 0; record < box.records.length; record += 3) {
        add(box.records[record]!, box.records[record + 1]!, sheetIndex, boxIndex);
      }
    });
  });
  lineIndexes.set(map, index);
  return index;
}

function contains(outer: SynctexRect, inner: SynctexRect): boolean {
  const slack = 0.01;
  return (
    inner.left >= outer.left - slack &&
    inner.right <= outer.right + slack &&
    inner.top >= outer.top - slack &&
    inner.bottom <= outer.bottom + slack
  );
}

function linesOutward(line: number): number[] {
  const order = [line];
  for (let distance = 1; distance <= MAX_FORWARD_LINE_DISTANCE; distance += 1) {
    order.push(line + distance);
    if (line - distance >= 1) order.push(line - distance);
  }
  return order;
}

/**
 * Where line `line` of the project file `projectPath` is typeset: the first
 * page that has it and the lines of text on it. A line TeX typeset nothing
 * from (a blank line, a comment) goes to the nearest one that it did,
 * looking further down first.
 */
export function synctexForward(
  map: SynctexMap,
  projectPath: string,
  line: number,
): SynctexForwardResult | null {
  const tags = tagsOf(map, projectPath);
  if (tags.size === 0 || !Number.isSafeInteger(line) || line < 1) return null;
  const index = lineIndex(map);
  for (const candidate of linesOutward(line)) {
    const found: number[] = [];
    for (const tag of tags) {
      const boxes = index.get(tag)?.get(candidate);
      if (boxes) found.push(...boxes);
    }
    if (found.length === 0) continue;
    let firstSheet = -1;
    for (let at = 0; at < found.length; at += 2) {
      const sheet = map.sheets[found[at]!]!;
      if (firstSheet < 0 || sheet.page < map.sheets[firstSheet]!.page) firstSheet = found[at]!;
    }
    const sheet = map.sheets[firstSheet]!;
    const rects: SynctexRect[] = [];
    for (let at = 0; at < found.length && rects.length < MAX_FORWARD_RECTS * 4; at += 2) {
      if (found[at] === firstSheet) rects.push(boxRect(map, sheet.boxes[found[at + 1]!]!));
    }
    // A box inside a matching line of text adds nothing to its highlight.
    const outermost = rects.filter(
      (rect, rectIndex) =>
        !rects.some(
          (other, otherIndex) =>
            otherIndex !== rectIndex &&
            contains(other, rect) &&
            (!contains(rect, other) || otherIndex < rectIndex),
        ),
    );
    return { page: sheet.page, rects: outermost.slice(0, MAX_FORWARD_RECTS) };
  }
  return null;
}

function gap(low: number, high: number, at: number): number {
  return at < low ? low - at : at > high ? at - high : 0;
}

/**
 * The source line that typeset the point (`x`, `y`, PDF points from the top
 * left) of page `page`: the innermost line of text under it, else the one
 * beside or nearest to it, and in it the last record at or before the point.
 * Files outside the project (a class, a package) are passed over.
 */
export function synctexInverse(
  map: SynctexMap,
  page: number,
  x: number,
  y: number,
): SynctexInverseResult | null {
  const sheet = map.sheets.find((candidate) => candidate.page === page);
  if (!sheet || !Number.isFinite(x) || !Number.isFinite(y)) return null;
  const projectPaths = new Map<number, string | null>();
  const projectPathOf = (tag: number) => {
    if (!projectPaths.has(tag)) {
      const path = map.inputs.get(tag);
      projectPaths.set(tag, path ? compiledProjectPath(path) : null);
    }
    return projectPaths.get(tag) ?? null;
  };
  const scale = bigPointsPerUnit(map);
  const offset = offsets(map);
  // The line of text the point is on comes first (a point in the margin
  // belongs to the line beside it), then the one nearest across; of nested
  // boxes, the innermost.
  const candidates = sheet.boxes.map((box) => {
    const rect = boxRect(map, box);
    return {
      box,
      down: gap(rect.top, rect.bottom, y),
      across: gap(rect.left, rect.right, x),
      area: Math.max(0, rect.right - rect.left) * Math.max(0, rect.bottom - rect.top),
    };
  });
  candidates.sort(
    (left, right) =>
      left.down - right.down || left.across - right.across || left.area - right.area,
  );
  for (const { box } of candidates) {
    // A line of a paragraph is linked to the source line that ended the
    // paragraph (often the blank line after it), and so are the records TeX
    // writes at the line's very start and end. Between the first and last
    // record from another line lies what the words were typeset from.
    const records = box.records;
    let from = 0;
    let to = records.length;
    while (from < to && records[from] === box.tag && records[from + 1] === box.line) from += 3;
    while (to > from && records[to - 3] === box.tag && records[to - 2] === box.line) to -= 3;
    if (from >= to) {
      from = 0;
      to = records.length;
    }
    let best: { tag: number; line: number; h: number } | null = null;
    let first: { tag: number; line: number; h: number } | null = null;
    for (let index = from; index < to; index += 3) {
      const record = {
        tag: records[index]!,
        line: records[index + 1]!,
        h: records[index + 2]! * scale + offset.x,
      };
      if (record.line < 1 || !projectPathOf(record.tag)) continue;
      if (record.h <= x && (!best || record.h >= best.h)) best = record;
      if (!first || record.h < first.h) first = record;
    }
    const chosen = best ?? first;
    if (chosen) {
      return { file: projectPathOf(chosen.tag)!, line: chosen.line };
    }
    const own = projectPathOf(box.tag);
    if (own && box.line >= 1) return { file: own, line: box.line };
  }
  return null;
}
