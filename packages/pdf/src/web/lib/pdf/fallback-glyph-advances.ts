/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What a PDF line's characters advance by when the browser cannot take the
 * advance from the embedded font.
 *
 * While a text line is edited, the reader shows the browser's own setting of
 * the draft (in the font pdf.js loaded from the file) between a keystroke and
 * the exact preview. A subset font very often has no glyph for the space, and
 * a producer that had no glyph for a character writes `.notdef` (read back as
 * U+0000): the PDF still gives each an advance (its /W entry), but the
 * browser sets the space in a fallback font, at that font's width, and gives
 * U+0000 no width at all. Every gap of the line then changed for as long as
 * the draft showed, and changed back when the preview arrived.
 *
 * The renderer's operator list holds each glyph as the PDF sets it: its
 * Unicode, its advance and whether the font file holds it. The advances of
 * the ones it does not hold are collected here per font, and the editor sets
 * its draft and its input by them.
 */

import { pdfRangeGlyphRect } from "./range-glyph-rect";
import type { PdfTextSourceAppearance } from "./source-text-appearance";
import { characterBoundaryIndexes } from "./text-caret";

type OperatorIds = Readonly<Record<string, number | undefined>>;

interface OperatorListLike {
  readonly fnArray: readonly number[];
  readonly argsArray: readonly unknown[];
}

/** Per character, its advance in em. */
export type PdfFallbackGlyphAdvances = ReadonlyMap<string, number>;

function isControl(character: string): boolean {
  const code = character.codePointAt(0) ?? 0;
  return code < 0x20 && character !== "\n" && character !== "\t";
}

/**
 * Per font (pdf.js's loaded name), the advance of every character the PDF
 * sets that the browser cannot set from the font file: one the file has no
 * glyph for, or a control character. `fontMatrixOf` gives a font's glyph
 * space scale (0.001 for all but Type 3 fonts). A character set at two
 * different advances in one font is left out: no single width is right.
 */
export function collectPdfFallbackGlyphAdvances(
  operatorList: OperatorListLike,
  ops: OperatorIds,
  fontMatrixOf: (fontName: string) => number | null = () => 0.001,
): Map<string, Map<string, number>> {
  const fonts = new Map<string, Map<string, number>>();
  const conflicted = new Map<string, Set<string>>();
  const shows = new Map<number, number>();
  // Which argument holds the glyphs.
  for (const [name, index] of [
    ["showText", 0],
    ["showSpacedText", 0],
    ["nextLineShowText", 0],
    ["nextLineSetSpacingShowText", 2],
  ] as const) {
    const id = ops[name];
    if (typeof id === "number") shows.set(id, index);
  }
  let font: string | null = null;
  const { fnArray, argsArray } = operatorList;
  for (let index = 0; index < fnArray.length; index += 1) {
    const fn = fnArray[index];
    const args = argsArray[index];
    if (fn === ops.setFont) {
      font = Array.isArray(args) && typeof args[0] === "string" ? args[0] : null;
      continue;
    }
    const at = shows.get(fn);
    if (at === undefined || font === null || !Array.isArray(args)) continue;
    const glyphs = args[at];
    if (!Array.isArray(glyphs)) continue;
    const scale = fontMatrixOf(font);
    if (scale === null || !Number.isFinite(scale) || scale <= 0) continue;
    for (const glyph of glyphs) {
      if (!glyph || typeof glyph !== "object") continue;
      const { unicode, width, isInFont } = glyph as {
        unicode?: unknown;
        width?: unknown;
        isInFont?: unknown;
      };
      if (typeof unicode !== "string" || [...unicode].length !== 1) continue;
      if (typeof width !== "number" || !Number.isFinite(width) || width < 0) continue;
      if (isInFont !== false && !isControl(unicode)) continue;
      const advance = width * scale;
      let known = fonts.get(font);
      if (!known) fonts.set(font, (known = new Map()));
      const refused = conflicted.get(font);
      if (refused?.has(unicode)) continue;
      const previous = known.get(unicode);
      if (previous !== undefined && Math.abs(previous - advance) > 1e-6) {
        known.delete(unicode);
        if (refused) refused.add(unicode);
        else conflicted.set(font, new Set([unicode]));
        continue;
      }
      known.set(unicode, advance);
    }
  }
  for (const [name, known] of fonts) if (known.size === 0) fonts.delete(name);
  return fonts;
}

// pdf.js names each loaded font uniquely for the document it came from
// (`g_d3_f5`), so one table serves every open document. Bounded: a font
// re-registered moves to the end, the oldest fall off.
const REGISTERED = new Map<string, PdfFallbackGlyphAdvances>();
/** Per font, the characters its pages set at more than one advance. */
const CONFLICTED = new Map<string, Set<string>>();
const REGISTERED_LIMIT = 512;










export function registerPdfFallbackGlyphAdvances(
  fonts: ReadonlyMap<string, ReadonlyMap<string, number>>,
): void {
  for (const [name, advances] of fonts) {
    const merged = new Map(REGISTERED.get(name) ?? []);
    const refused = CONFLICTED.get(name) ?? new Set<string>();
    for (const [character, advance] of advances) {
      if (refused.has(character)) continue;
      const known = merged.get(character);
      if (known !== undefined && Math.abs(known - advance) > 1e-6) {
        merged.delete(character);
        refused.add(character);
        continue;
      }
      merged.set(character, advance);
    }
    if (refused.size) CONFLICTED.set(name, refused);
    REGISTERED.delete(name);
    REGISTERED.set(name, merged);
  }
  while (REGISTERED.size > REGISTERED_LIMIT) {
    const oldest = REGISTERED.keys().next().value;
    if (oldest === undefined) break;
    REGISTERED.delete(oldest);
    CONFLICTED.delete(oldest);
  }
}

/** The advances known for the first font of a CSS font-family list. */
export function pdfFallbackGlyphAdvancesFor(
  fontFamily: string | null | undefined,
): PdfFallbackGlyphAdvances | null {
  if (!fontFamily) return null;
  const first = fontFamily.split(",")[0]?.trim().replace(/^["']|["']$/g, "");
  return (first && REGISTERED.get(first)) || null;
}

/**
 * The CSS `word-spacing` that makes the browser's space as wide as the PDF's:
 * the PDF's advance less what the browser's own space measures in that font
 * list. Null when the PDF's space is the font's own (nothing to correct).
 */
export function pdfSpaceWordSpacingPx(
  advances: PdfFallbackGlyphAdvances | null,
  fontSizePx: number,
  measuredSpacePx: number | null,
): number | null {
  const space = advances?.get(" ");
  if (
    space === undefined ||
    measuredSpacePx === null ||
    !Number.isFinite(fontSizePx) ||
    fontSizePx <= 0 ||
    !Number.isFinite(measuredSpacePx)
  ) {
    return null;
  }
  return space * fontSizePx - measuredSpacePx;
}

/**
 * Set `element`'s content to `text`, giving each control character the PDF
 * sets with an advance (U+0000 for a missing glyph) a box of that width: the
 * browser gives such a character none. Plain text when there is none.
 */
export function fillPdfTextDraft(
  element: HTMLElement,
  text: string,
  advances: PdfFallbackGlyphAdvances | null,
  fontSizePx: number,
): void {
  const widths = advances
    ? [...advances].filter(([character]) => isControl(character))
    : [];
  if (
    widths.length === 0 ||
    !Number.isFinite(fontSizePx) ||
    fontSizePx <= 0 ||
    !widths.some(([character]) => text.includes(character))
  ) {
    // The same text already there is left as it is: written again, the
    // draft was laid out again for the next read of it, on each of the
    // layout passes a key makes.
    const only = element.childNodes.length === 1 ? element.firstChild : null;
    if (only?.nodeType === 3 && only.nodeValue === text) return;
    element.textContent = text;
    return;
  }
  const known = new Map(widths);
  const nodes: Node[] = [];
  let run = "";
  const flush = () => {
    if (run) nodes.push(element.ownerDocument.createTextNode(run));
    run = "";
  };
  for (const character of text) {
    const advance = known.get(character);
    if (advance === undefined) {
      run += character;
      continue;
    }
    flush();
    const box = element.ownerDocument.createElement("span");
    box.dataset.pdfFallbackGlyph = "true";
    box.style.display = "inline-block";
    box.style.width = `${advance * fontSizePx}px`;
    nodes.push(box);
  }
  flush();
  element.replaceChildren(...nodes);
}










export function pdfTextDraftNodes(
  doc: Document,
  text: string,
  advances: PdfFallbackGlyphAdvances | null | undefined,
): Node[] {
  const known = advances ? new Map([...advances].filter(([character]) => isControl(character))) : null;
  if (!known?.size || ![...known.keys()].some((character) => text.includes(character))) {
    return text ? [doc.createTextNode(text)] : [];
  }
  const nodes: Node[] = [];
  let run = "";
  for (const character of text) {
    const advance = known.get(character);
    if (advance === undefined) {
      run += character;
      continue;
    }
    if (run) nodes.push(doc.createTextNode(run));
    run = "";
    const box = doc.createElement("span");
    box.dataset.pdfFallbackGlyph = "true";
    box.style.display = "inline-block";
    box.style.width = `${advance}em`;
    nodes.push(box);
  }
  if (run) nodes.push(doc.createTextNode(run));
  return nodes;
}

const MEASURED_SPACE = new Map<string, number>();

/**
 * How wide the browser sets a space in a CSS font list at a size, measured
 * in the page itself: what `pdfSpaceWordSpacingPx` corrects from. (Not on a
 * canvas: a generic family resolves by the page's language there, and to
 * another font, with another space, on a canvas.) Null where text cannot be
 * measured.
 */
export function measureBrowserSpacePx(fontFamily: string, fontSizePx: number): number | null {
  if (!Number.isFinite(fontSizePx) || fontSizePx <= 0) return null;
  const key = `${fontSizePx}px ${fontFamily}`;
  const known = MEASURED_SPACE.get(key);
  if (known !== undefined) return known;
  if (typeof document === "undefined" || !document.body) return null;
  const probe = document.createElement("span");
  // Sixteen spaces: a sixteenth of the rounding of one.
  probe.textContent = " ".repeat(16);
  probe.setAttribute("aria-hidden", "true");
  Object.assign(probe.style, {
    position: "absolute",
    left: "-9999px",
    top: "0",
    visibility: "hidden",
    whiteSpace: "pre",
    wordSpacing: "0px",
    letterSpacing: "0px",
    fontFamily,
    fontSize: `${fontSizePx}px`,
  });
  document.body.appendChild(probe);
  const width = probe.getBoundingClientRect().width / 16;
  probe.remove();
  if (!Number.isFinite(width) || width <= 0) return null;
  // Kept only once every face of the list that sets a space is loaded:
  // measured while a web face was still coming, the fallback's space was
  // kept, and once the face came every space was set off by the difference.
  let settled = true;
  try {
    settled = !document.fonts || document.fonts.check(`${fontSizePx}px ${fontFamily}`, " ");
  } catch {
    settled = false;
  }
  if (!settled) return width;
  if (MEASURED_SPACE.size > 256) MEASURED_SPACE.clear();
  MEASURED_SPACE.set(key, width);
  return width;
}

/** Whether `text` holds a character the draft sets as a placeholder box. */
export function pdfTextDraftHasFallbackGlyphs(
  text: string,
  advances: PdfFallbackGlyphAdvances | null,
): boolean {
  if (!advances) return false;
  for (const [character] of advances) {
    if (isControl(character) && text.includes(character)) return true;
  }
  return false;
}

/**
 * The draft's characters as the browser set them, in the page's pixel
 * coordinates, in the shape the source caret draws from: the textarea
 * underneath gives a U+0000 no width while the draft gives it the PDF's
 * advance, so between a keystroke and the exact preview the system caret
 * stood up to a few glyphs left of the text. Placeholder boxes are one
 * character each; `canonical` maps a draft offset to the input's own.
 * `angle` is the draft's turn on the page (CSS degrees, clockwise): a line
 * of a page turned by a quarter is set turned, and its boxes, upright on
 * screen, are its characters' own; any other turn has no such boxes (null).
 */
/** What a draft measured last: each character's box from the draft's corner,
 * for the words the next key left where they were. */
type MeasuredDraft = Readonly<{ text: string; width: number; style: string;
  boxes: ReadonlyMap<number, Readonly<{ left: number; top: number; right: number; bottom: number }>> }>;
const measuredDrafts = new WeakMap<HTMLElement, MeasuredDraft>();
/** Shorter than this, measuring all of it is nothing. */
const DRAFT_PREFIX_WORTH_KEEPING = 32;
/** The offset before which the words the draft had last time stand as they
 * stood: up to the last break the words in common have (what follows it can
 * wrap with the key typed), proved by the boxes at both ends of that span --
 * a line set centred or justified, or a frame that changed, moves them. */
function keptDraftPrefix(last: MeasuredDraft | undefined, text: string, width: number, style: string,
  measure: (offset: number) => Readonly<{ left: number; top: number }> | null): number {
  if (!last || last.width !== width || last.style !== style) return 0;
  let common = 0;
  const shorter = Math.min(last.text.length, text.length);
  while (common < shorter && last.text.charCodeAt(common) === text.charCodeAt(common)) common += 1;
  if (common < DRAFT_PREFIX_WORTH_KEEPING) return 0;
  let kept = common - 1;
  while (kept > 0 && !/\s/u.test(text[kept - 1] ?? "")) kept -= 1;
  // No break in common (one long word, words without spaces): a few characters
  // back, where a key's rewrap of CJK ends.
  if (kept === 0) kept = common - 8;
  if (kept < DRAFT_PREFIX_WORTH_KEEPING) return 0;
  const boxes = [...last.boxes.keys()].filter(start => start < kept);
  if (!boxes.length) return 0;
  const lastStart = Math.max(...boxes);
  for (const start of [0, lastStart]) {
    const was = last.boxes.get(start), now = measure(start);
    if (!was || !now || Math.abs(was.left - now.left) > 0.01 || Math.abs(was.top - now.top) > 0.01) return 0;
  }
  return kept;
}

export function pdfTextDraftAppearance(
  draft: HTMLElement,
  pageHost: HTMLElement,
  pageWidth: number,
  pageHeight: number,
  text: string,
  identity: Readonly<{ sourceSha256: string; pageNumber: number }>,
  canonical: (offset: number) => number = (offset) => offset,
  angle = 0,
): PdfTextSourceAppearance | null {
  const host = pageHost.getBoundingClientRect();
  if (!(host.width > 0 && host.height > 0 && pageWidth > 0 && pageHeight > 0)) return null;
  const turn = (((angle % 360) + 360) % 360);
  if (!Number.isFinite(turn) || Math.abs(turn - Math.round(turn / 90) * 90) > 0.001) return null;
  const quarter = Math.round(turn / 90) % 4;
  // The line's own axes on screen: along the text, and from its top down.
  const along = ([[1, 0], [0, 1], [-1, 0], [0, -1]] as const)[quarter];
  const down = [-along[1], along[0]] as const;
  // No layout engine (a unit test's DOM): no boxes to draw the caret from.
  if (typeof draft.ownerDocument.defaultView?.Range?.prototype.getBoundingClientRect !== "function") return null;
  const sx = pageWidth / host.width;
  const sy = pageHeight / host.height;
  const doc = draft.ownerDocument;
  const characters: PdfTextSourceAppearance["characters"][number][] = [];
  let offset = 0;
  const corner = draft.getBoundingClientRect();
  const relative = (rect: DOMRect) => ({ left: rect.left - corner.left, top: rect.top - corner.top,
    right: rect.right - corner.left, bottom: rect.bottom - corner.top });
  const style = draft.getAttribute("style") ?? "";
  // A character's box by its offset, read now (to prove the words kept).
  const measureAt = (wanted: number): Readonly<{ left: number; top: number }> | null => {
    let at = 0;
    const walk = (node: Node): Readonly<{ left: number; top: number }> | null => {
      if (node.nodeType === Node.TEXT_NODE) {
        const content = node.textContent ?? "";
        if (wanted < at + content.length) {
          const boundaries = characterBoundaryIndexes(content);
          const index = boundaries.indexOf(wanted - at);
          if (index < 0 || index + 1 >= boundaries.length) return null;
          const range = doc.createRange();
          range.setStart(node, boundaries[index] ?? 0);
          range.setEnd(node, boundaries[index + 1] ?? content.length);
          const rect = pdfRangeGlyphRect(range);
          range.detach();
          return relative(rect);
        }
        at += content.length;
        return null;
      }
      if (node instanceof HTMLElement && node.dataset.pdfFallbackGlyph === "true") {
        if (wanted === at) return relative(node.getBoundingClientRect());
        at += 1;
        return null;
      }
      for (const child of Array.from(node.childNodes)) {
        const found = walk(child);
        if (found) return found;
      }
      return null;
    };
    return walk(draft);
  };
  const last = measuredDrafts.get(draft);
  const kept = keptDraftPrefix(last, text, corner.width, style, measureAt);
  const boxes = new Map<number, Readonly<{ left: number; top: number; right: number; bottom: number }>>();
  const keptRect = (start: number): DOMRect | null => {
    if (start >= kept) return null;
    const box = last?.boxes.get(start);
    return box ? new DOMRect(box.left + corner.left, box.top + corner.top, box.right - box.left, box.bottom - box.top) : null;
  };
  const push = (start: number, end: number, rect: DOMRect, value: string) => {
    boxes.set(start, relative(rect));
    const [advance, height] = along[0] ? [rect.width, rect.height] : [rect.height, rect.width];
    if (!(advance >= 0 && height > 0)) return;
    const left = (rect.left - host.left) * sx;
    const right = (rect.right - host.left) * sx;
    const top = (rect.top - host.top) * sy;
    const bottom = (rect.bottom - host.top) * sy;
    // The box's corner at the line's start or end and top or bottom.
    const corner = (end: boolean, below: boolean) => {
      const x = (end ? along[0] : -along[0]) + (below ? down[0] : -down[0]) > 0 ? right : left;
      const y = (end ? along[1] : -along[1]) + (below ? down[1] : -down[1]) > 0 ? bottom : top;
      return [x, y] as const;
    };
    const origin = corner(false, true);
    characters.push(
      Object.freeze({
        text: value,
        start: canonical(start),
        end: canonical(end),
        origin: Object.freeze([...origin]),
        quad: Object.freeze([...corner(false, false), ...corner(true, false), ...origin, ...corner(true, true)]),
        fontName: "",
        fontSize: along[0] ? bottom - top : right - left,
        color: Object.freeze([]),
        bidi: 0,
        direction: Object.freeze([...along]),
      }),
    );
  };
  const visit = (node: Node): void => {
    if (node.nodeType === Node.TEXT_NODE) {
      const content = node.textContent ?? "";
      const boundaries = characterBoundaryIndexes(content);
      for (let index = 0; index < boundaries.length - 1; index += 1) {
        const start = boundaries[index] ?? 0;
        const end = boundaries[index + 1] ?? content.length;
        const stood = keptRect(offset + start);
        if (stood) { push(offset + start, offset + end, stood, content.slice(start, end)); continue; }
        const range = doc.createRange();
        range.setStart(node, start);
        range.setEnd(node, end);
        const rect = pdfRangeGlyphRect(range);
        range.detach();
        push(offset + start, offset + end, rect, content.slice(start, end));
      }
      offset += content.length;
      return;
    }
    if (node instanceof HTMLElement && node.dataset.pdfFallbackGlyph === "true") {
      push(offset, offset + 1, keptRect(offset) ?? node.getBoundingClientRect(), text[offset] ?? "\u0000");
      offset += 1;
      return;
    }
    for (const child of Array.from(node.childNodes)) visit(child);
  };
  visit(draft);
  if (characters.length === 0 || offset !== text.length) return null;
  measuredDrafts.set(draft, { text, width: corner.width, style, boxes });
  const xs = characters.flatMap((character) => character.quad.filter((_, index) => index % 2 === 0));
  const ys = characters.flatMap((character) => character.quad.filter((_, index) => index % 2 === 1));
  return Object.freeze({
    sourceSha256: identity.sourceSha256,
    pageNumber: identity.pageNumber,
    text,
    pageTransform: Object.freeze([1, 0, 0, 1, 0, 0]),
    characters: Object.freeze(characters),
    bounds: Object.freeze({
      left: Math.min(...xs),
      top: Math.min(...ys),
      right: Math.max(...xs),
      bottom: Math.max(...ys),
    }),
  });
}
