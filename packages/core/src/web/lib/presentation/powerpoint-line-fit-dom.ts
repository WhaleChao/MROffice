/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Gives a slide paragraph's text column the width at which the browser ends
 * its lines where PowerPoint does (`./powerpoint-line-fit`). The column's
 * words are read from the DOM as drawn -- each run's own face, size, weight,
 * letter spacing and case -- and measured on a canvas with the same fonts.
 * The width is set as the column's inline-end margin and checked against the
 * lines the browser then lays out; anything that does not come out as
 * PowerPoint's lines is undone, so the paragraph is never worse than the
 * browser alone.
 */

import { officeFontMetricMatchesWidths, officeFontSubstitution, powerPointSubstitutes } from "../office-fonts/substitutes";
import substituteTable from "../office-fonts/substitutes.generated.json";
import { notedFontFaceCount, notedFontFaces, type DeclaredFontFace } from "../font-face-declarations";
import {
  POWERPOINT_ADVANCE_QUANTUM_PX,
  browserWidthRange,
  lineFitTokens,
  powerPointFitWidth,
  powerPointLines,
  type LineFitGlyph,
  type LineFitLine,
  type LineFitToken,
} from "./powerpoint-line-fit";

/** Set on a column whose margin this module wrote. */
export const POWERPOINT_LINE_FIT_ATTRIBUTE = "data-ppt-line-fit";

interface DomGlyph extends LineFitGlyph {
  /** Canvas font and letter spacing of the run it is in. */
  readonly font: string;
  /** Whether the run is drawn kerned (`font-kerning`), as the canvas must measure it. */
  readonly kerned: boolean;
  readonly spacing: number;
  /** Where it is in the DOM (null for a `<br>`). */
  readonly node: Text | null;
  readonly offset: number;
}

let measureContext: CanvasRenderingContext2D | null | undefined;
function context(doc: Document): CanvasRenderingContext2D | null {
  if (measureContext === undefined) {
    try {
      measureContext = doc.createElement("canvas").getContext("2d");
    } catch {
      measureContext = null;
    }
  }
  return measureContext;
}

const advanceCache = new Map<string, number>();
function advance(ctx: CanvasRenderingContext2D, font: string, ch: string): number {
  const key = `${font}\u0000${ch}`;
  ctx.fontKerning = "normal";
  const cached = advanceCache.get(key);
  if (cached !== undefined) return cached;
  ctx.font = font;
  const width = ctx.measureText(ch).width;
  if (advanceCache.size > 20_000) advanceCache.clear();
  advanceCache.set(key, width);
  return width;
}

function transformed(text: string, transform: string): string {
  if (transform === "uppercase") return text.toUpperCase();
  if (transform === "lowercase") return text.toLowerCase();
  return text;
}

/** Elements whose drawing is not text the canvas can measure. */
const OPAQUE =
  "img,svg,canvas,video,input,textarea,math,.katex,[data-presentation-math],[data-presentation-equation],[data-office-text-marker],[data-cell-math-run]";

/** The column's glyphs as drawn, or null when it holds something else. */
function readGlyphs(column: HTMLElement, ctx: CanvasRenderingContext2D, anyFace = false, families?: Set<string>): DomGlyph[] | null {
  const view = column.ownerDocument.defaultView;
  if (!view) return null;
  if (column.querySelector(OPAQUE)) return null;
  const glyphs: DomGlyph[] = [];
  const styles = new Map<Element, { font: string; spacing: number; transform: string; kerned: boolean } | null>();
  const styleOf = (element: Element) => {
    if (styles.has(element)) return styles.get(element)!;
    const css = view.getComputedStyle(element);
    if (families) for (const family of familiesOf(css.fontFamily)) families.add(family);
    let value: { font: string; spacing: number; transform: string; kerned: boolean } | null = null;
    const smallCaps = css.fontVariantCaps && css.fontVariantCaps !== "normal";
    const hidden = css.display === "none" || css.visibility === "hidden";
    // PowerPoint's rule needs PowerPoint's glyphs: a run drawn by a stand-in
    // (the Office face missing here) or a system fallback breaks by widths
    
    // p16: an Aptos Display title the stand-in fitted on one line).
    // Lines the text engine laid out (`ENGINE_LINE_STARTS_ATTRIBUTE`) are
    // PowerPoint's whatever face draws them: only the browser's own widths
    // are measured here, to end its lines at them.
    const real = hidden || anyFace ? true : stackDrawnByRealFace(column, css.fontFamily);
    if (real !== true) {
      styles.set(element, null);
      return null;
    }
    if (!smallCaps && !hidden) {
      const spacing = css.letterSpacing === "normal" ? 0 : parseFloat(css.letterSpacing);
      value = {
        font: `${css.fontStyle} ${css.fontWeight} ${css.fontStretch === "100%" ? "" : css.fontStretch} ${css.fontSize} ${css.fontFamily}`.replace(/\s+/g, " "),
        spacing: Number.isFinite(spacing) ? spacing : 0,
        transform: css.textTransform,
        kerned: css.fontKerning !== "none",
      };
    }
    styles.set(element, value);
    return value;
  };
  const walker = column.ownerDocument.createTreeWalker(column, 0x1 | 0x4);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (node.nodeType === 1) {
      if ((node as Element).tagName === "BR") {
        glyphs.push({ ch: "\n", advance: 0, font: "", kerned: true, spacing: 0, node: null, offset: 0 });
      }
      continue;
    }
    const text = node as Text;
    const parent = text.parentElement;
    if (!parent || !text.data) continue;
    const style = styleOf(parent);
    if (!style) return null;
    const shown = transformed(text.data, style.transform);
    // A case change that changes the length cannot be mapped back.
    if (shown.length !== text.data.length) return null;
    for (let offset = 0; offset < shown.length; offset += 1) {
      const ch = shown[offset];
      // A surrogate pair, a combining mark: not modelled.
      const code = ch.charCodeAt(0);
      if (
        (code >= 0xd800 && code <= 0xdfff) ||
        (code >= 0x0300 && code <= 0x036f) ||
        (code >= 0x200b && code <= 0x200f) ||
        code === 0x2028 ||
        code === 0x2029
      ) {
        return null;
      }
      glyphs.push({
        ch,
        advance: ch === "\n" ? 0 : advance(ctx, style.font, ch) + style.spacing,
        font: style.font,
        kerned: style.kerned,
        spacing: style.spacing,
        node: text,
        offset,
      });
    }
  }
  return glyphs;
}

/**
 * A run as the slide draws it, for the canvas to measure: no common ligature
 * (`presentationTextShapingCss`; a canvas has no `font-variant-ligatures`).
 * A zero-width non-joiner between letters of the scripts whose ligatures are
 * optional (Latin, Greek, Cyrillic) blocks their ligatures and nothing else:
 * kerning skips it and it has no advance. Other scripts are measured as they
 * are (their joining and required ligatures stay in the drawing too).
 */
export function unligatedForMeasure(text: string): string {
  // Combining marks stay on their letters (a non-joiner would keep "e" + U+0301
  // from composing into "é").
  return text.length > 1 && /^[\u0000-\u058f\u1d00-\u1fff\u2000-\u206f]*$/.test(text) && !/[\u0300-\u036f\u0483-\u0489\u0591-\u05c7\u1dc0-\u1dff\u20d0-\u20ff]/.test(text)
    ? [...text].join("\u200c")
    : text;
}

/** The browser's own width of glyphs `[start, end)`: kerned, run by run. */
function exactMeasure(ctx: CanvasRenderingContext2D, glyphs: readonly DomGlyph[]) {
  return (start: number, end: number): number => {
    let width = 0;
    let index = start;
    while (index < end) {
      const font = glyphs[index].font;
      const kerned = glyphs[index].kerned;
      let text = "";
      let count = 0;
      while (index < end && glyphs[index].font === font && glyphs[index].kerned === kerned) {
        if (glyphs[index].ch !== "\n") {
          text += glyphs[index].ch;
          count += 1;
        }
        index += 1;
      }
      if (text) {
        ctx.font = font;
        // The browser's own measure of the run: kerned only where it draws kerned.
        ctx.fontKerning = kerned ? "normal" : "none";
        width += ctx.measureText(unligatedForMeasure(text)).width + count * glyphs[index - 1].spacing;
      }
    }
    return width;
  };
}

/** The top of the line box glyph `index` is drawn on. */
function glyphTop(doc: Document, glyph: DomGlyph): number | null {
  if (!glyph.node) return null;
  const range = doc.createRange();
  range.setStart(glyph.node, glyph.offset);
  range.setEnd(glyph.node, glyph.offset + 1);
  const rect = range.getClientRects()[0];
  return rect ? rect.top : null;
}

/** Whether the browser ends its lines where `lines` say. */
function browserLinesMatch(
  doc: Document,
  glyphs: readonly DomGlyph[],
  tokens: readonly LineFitToken[],
  lines: readonly LineFitLine[],
): boolean {
  let previousTop: number | null = null;
  for (const line of lines) {
    const first = tokens[line.first].start;
    const lastInk = tokens[line.last].inkEnd - 1;
    if (lastInk < first) {
      previousTop = null;
      continue;
    }
    const top = glyphTop(doc, glyphs[first]);
    const bottom = glyphTop(doc, glyphs[lastInk]);
    if (top === null || bottom === null) return false;
    // The line's first and last words stand on one line box ...
    if (Math.abs(top - bottom) > 0.5) return false;
    // ... below the previous line's.
    if (previousTop !== null && top <= previousTop + 0.5) return false;
    previousTop = top;
  }
  return true;
}

/** What one column needs, read without writing anything. */
type ColumnPlan =
  | { readonly column: HTMLElement; readonly kind: "release" }
  | { readonly column: HTMLElement; readonly kind: "keep" }
  | {
      readonly column: HTMLElement;
      readonly kind: "fit";
      readonly margin: number;
      readonly glyphs: readonly DomGlyph[];
      readonly tokens: readonly LineFitToken[];
      readonly lines: readonly LineFitLine[];
    };

/** What a column was last fitted from: its words, width and face. */
const fittedFrom = new WeakMap<HTMLElement, string>();
/** The families a column's glyphs were measured in when it was last planned. */
const fittedFamilies = new WeakMap<HTMLElement, ReadonlySet<string>>();








const familyEpochs = new Map<string, number>();
let anyFontEpoch = 0;
let fontArrivals = 0;
const watchedFonts = new WeakSet<object>();
function watchFonts(doc: Document): void {
  const fonts = doc.fonts as (FontFaceSet & EventTarget) | undefined;
  if (fonts && !watchedFonts.has(fonts) && typeof fonts.addEventListener === "function") {
    watchedFonts.add(fonts);
    fonts.addEventListener("loadingdone", (event) => {
      const faces = (event as Event & { fontfaces?: readonly FontFace[] }).fontfaces;
      fontArrivals += 1;
      if (!faces || faces.length === 0) anyFontEpoch = fontArrivals;
      else for (const face of faces) familyEpochs.set(faceFacts(face, doc).family, fontArrivals);
      advanceCache.clear();
    });
  }
}
/** The latest arrival that can change a column measured in `families`. */
function familiesEpoch(doc: Document, families: ReadonlySet<string> | undefined): number {
  watchFonts(doc);
  if (!families) return fontArrivals;
  let epoch = anyFontEpoch;
  for (const family of families) epoch = Math.max(epoch, familyEpochs.get(family) ?? 0);
  return epoch;
}
/** The families named by a computed `font-family` list. */
function familiesOf(stack: string): string[] {
  return stack.split(",").map(familyKey).filter(Boolean);
}

/**
 * Whether faces that just arrived (a `loadingdone` event's) can change how
 * `columns` were fitted: a column not fitted yet, or one measured in a
 * family among them. A host listening for font loads refits only then.
 */
export function powerPointLineFitStaleAfter(columns: Iterable<HTMLElement>, event: Event): boolean {
  const arrived = arrivedFaces(event);
  if (!arrived) return true;
  for (const column of columns) {
    const families = fittedFamilies.get(column);
    if (!families) return true;
    for (const family of families) {
      const ranges = arrived.get(family);
      if (ranges === undefined) continue;
      // A face for some characters (a unicode-range slice: an East Asian
      // family arrives as a hundred of them) draws none of a column that
      // holds none of them: every cell of a large table measured again as
      
      if (ranges === null || textInRanges(column.textContent ?? "", ranges)) return true;
    }
  }
  return false;
}

/** By family, the code point ranges of the faces a `loadingdone` brought
 * (null: the whole of Unicode), or null when the event names no faces. */
type ArrivedFaces = Map<string, (readonly [number, number])[] | null>;
const arrivedByEvent = new WeakMap<Event, ArrivedFaces | null>();
function arrivedFaces(event: Event): ArrivedFaces | null {
  if (arrivedByEvent.has(event)) return arrivedByEvent.get(event)!;
  const faces = (event as Event & { fontfaces?: readonly FontFace[] }).fontfaces;
  let arrived: ArrivedFaces | null = null;
  if (faces && faces.length > 0) {
    arrived = new Map();
    const doc = typeof document === "undefined" ? null : document;
    for (const face of faces) {
      const { family, ranges } = faceFacts(face, doc);
      const known = arrived.get(family);
      if (known === null) continue;
      arrived.set(family, ranges === null ? null : [...(known ?? []), ...ranges]);
    }
  }
  arrivedByEvent.set(event, arrived);
  return arrived;
}

/**
 * A face's family (as `familyKey` gives it) and code point ranges, read once
 * per face: from what its builder declared (lib/font-face-declarations), else
 * from the face itself, once. In Safari 16.0-16.3 each read of a FontFace's
 * descriptors restyles the page, and a large family's faces arrive by the
 * hundred, each named again in every `loadingdone` after.
 */
const factsByFace = new WeakMap<FontFace, { family: string; ranges: (readonly [number, number])[] | null }>();
let declaredIndex: { doc: Document; count: number; byFace: WeakMap<FontFace, DeclaredFontFace> } | null = null;
function declaredOf(face: FontFace, doc: Document | null): DeclaredFontFace | undefined {
  if (!doc) return undefined;
  const count = notedFontFaceCount();
  if (!declaredIndex || declaredIndex.doc !== doc || declaredIndex.count !== count) {
    const byFace = new WeakMap<FontFace, DeclaredFontFace>();
    for (const noted of notedFontFaces(doc) ?? []) if (noted.declared) byFace.set(noted.face, noted.declared);
    declaredIndex = { doc, count, byFace };
  }
  return declaredIndex.byFace.get(face);
}
function faceFacts(face: FontFace, doc: Document | null): { family: string; ranges: (readonly [number, number])[] | null } {
  const known = factsByFace.get(face);
  if (known) return known;
  const declared = declaredOf(face, doc);
  const facts = declared
    ? { family: familyKey(declared.family), ranges: unicodeRanges(declared.unicodeRange) }
    : { family: familyKey(face.family), ranges: unicodeRanges(face.unicodeRange) };
  factsByFace.set(face, facts);
  return facts;
}

/** A face's `unicode-range` as code point ranges, null for all of Unicode or
 * one this reader does not follow. */
function unicodeRanges(descriptor: string | undefined): (readonly [number, number])[] | null {
  if (!descriptor) return null;
  const ranges: [number, number][] = [];
  for (const part of descriptor.split(",")) {
    const match = /^\s*U\+([0-9a-f?]{1,6})(?:-([0-9a-f]{1,6}))?\s*$/i.exec(part);
    if (!match) return null;
    const [, start, end] = match;
    const low = parseInt(start.replace(/\?/g, "0"), 16);
    const high = end ? parseInt(end, 16) : parseInt(start.replace(/\?/g, "f"), 16);
    if (!Number.isFinite(low) || !Number.isFinite(high)) return null;
    if (low <= 0 && high >= 0x10ffff) return null;
    ranges.push([low, high]);
  }
  return ranges.length ? ranges : null;
}

function textInRanges(text: string, ranges: readonly (readonly [number, number])[]): boolean {
  for (const character of text) {
    const code = character.codePointAt(0)!;
    for (const [low, high] of ranges) if (code >= low && code <= high) return true;
  }
  return false;
}

const GENERIC_FAMILIES = new Set([
  "serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui", "ui-serif",
  "ui-sans-serif", "ui-monospace", "ui-rounded", "emoji", "math", "fangsong",
  "-apple-system", "blinkmacsystemfont",
]);
function familyKey(name: string): string {
  return name.trim().replace(/^["']|["']$/g, "").toLowerCase();
}
/** Whether a face of this name is installed here: a FontFace of a local()
 * source alone loads only when the machine has it (as the Office font
 * registry probes it). Unknown until the probe answers. */
const installed = new Map<string, boolean>();
const probing = new Set<string>();
const awaitingProbe = new Set<HTMLElement>();
function installedHere(family: string, name: string, column: HTMLElement): boolean | undefined {
  const known = installed.get(family);
  if (known !== undefined) return known;
  awaitingProbe.add(column);
  if (probing.has(family) || typeof FontFace === "undefined") return typeof FontFace === "undefined" ? false : undefined;
  probing.add(family);
  const settle = (found: boolean) => {
    if (installed.has(family)) return;
    installed.set(family, found);
    probing.delete(family);
    // Fit again what waited for the answer: columns measured in this family.
    fontArrivals += 1;
    familyEpochs.set(family, fontArrivals);
    const waiting = [...awaitingProbe];
    awaitingProbe.clear();
    requestPowerPointLineFit(waiting);
  };
  try {
    new FontFace("mrai-line-fit-probe", `local("${name.replace(/["\\]/g, "\\$&")}")`)
      .load()
      .then(() => settle(true), () => settle(false));
    setTimeout(() => settle(false), 3000);
  } catch {
    settle(false);
  }
  return undefined;
}









function stackDrawnByRealFace(column: HTMLElement, stack: string): boolean | undefined {
  for (const raw of stack.split(",")) {
    const family = familyKey(raw);
    if (!family) continue;
    if (GENERIC_FAMILIES.has(family)) return false;
    if (drawnWithOfficeWidths(family)) return true;
    return installedHere(family, raw.trim().replace(/^["']|["']$/g, ""), column);
  }
  return false;
}

/** The open faces that set exactly the advance widths of the Office face they
 * stand in for (`metric: "exact"`: Arimo for Arial, Carlito for Calibri,
 * Tinos for Times New Roman, ...): lower-case stand-in name -> Office face. */
let exactStandIns: Map<string, string> | null = null;












export function drawnWithOfficeWidths(family: string): boolean {
  const key = familyKey(family);
  if (!exactStandIns) {
    exactStandIns = new Map();
    for (const entry of Object.values(substituteTable as Record<string, { office?: string; substitutes?: string[]; metric?: string }>)) {
      if (officeFontMetricMatchesWidths(entry.metric) && entry.office && entry.substitutes?.length) {
        const standIn = familyKey(entry.substitutes[0]);
        if (!exactStandIns.has(standIn)) exactStandIns.set(standIn, familyKey(entry.office));
      }
    }
  }
  // A stand-in named directly is the Office face it stands in for (Carlito:
  // Calibri), though the table also lists it as a family of its own.
  const office =
    exactStandIns.get(key) ?? (officeFontMetricMatchesWidths(officeFontSubstitution(key)?.metric) ? key : undefined);
  if (!office) return false;
  const drawnBy = powerPointSubstitutes(key)[0];
  return !drawnBy || familyKey(drawnBy) === office;
}

/**
 * Where the text engine ended a row's lines (`PresentationEngineTextBody`'s
 * layout of the body being typed in): the offsets in the column's text at
 * which each line after the first starts. Set on the column, read here, so
 * the browser's lines end where the engine's drawing of the body at rest
 * ends them.
 */
export const ENGINE_LINE_STARTS_ATTRIBUTE = "data-engine-line-starts";

/** The engine's lines as runs of tokens, or null when a start falls inside a token. */
function engineLines(tokens: readonly LineFitToken[], starts: string): LineFitLine[] | null {
  const offsets = starts.split(",").map(Number).filter((value) => Number.isFinite(value) && value > 0);
  const lines: LineFitLine[] = [];
  let first = 0;
  for (const offset of offsets) {
    const token = tokens.findIndex((candidate, index) => index > first && candidate.start === offset);
    if (token < 0) return null;
    lines.push({ first, last: token - 1 });
    first = token;
  }
  if (first < tokens.length) lines.push({ first, last: tokens.length - 1 });
  return lines;
}

function planColumn(column: HTMLElement): ColumnPlan {
  // A block-level column (a table cell's paragraph) or a row's growing flex
  // item: anything else has no line width of its own.
  const doc = column.ownerDocument;
  const view = doc.defaultView;
  const release = { column, kind: "release" } as const;
  if (!view || !column.isConnected) return release;
  const css = view.getComputedStyle(column);
  if (css.whiteSpace === "pre" || css.whiteSpace === "nowrap") return release;
  if (css.writingMode && css.writingMode !== "horizontal-tb") return release;
  // A flex item as wide as its words (a centred row moving its marker with
  // them) has no line width of its own to fit.
  const parentCss = column.parentElement ? view.getComputedStyle(column.parentElement) : null;
  const flexParent = /flex/.test(parentCss?.display ?? "");
  if (flexParent) {
    if (/column/.test(parentCss!.flexDirection)) {
      // In a column the item's width is the cross size: its own only when
      // it is stretched across the container.
      const align = css.alignSelf === "auto" ? parentCss!.alignItems : css.alignSelf;
      if (align !== "stretch" && align !== "normal") return release;
    } else if (css.flexGrow === "0") {
      return release;
    }
  } else if (css.display !== "block") {
    return release;
  }
  // The width the row gives the column, before any fitting margin: read
  // through the margin rather than by removing it, so a column that keeps
  // its fit is not laid out again.
  const fitted = column.hasAttribute(POWERPOINT_LINE_FIT_ATTRIBUTE);
  const shown = parseFloat(css.width);
  const width = shown + (fitted ? parseFloat(css.marginInlineEnd) || 0 : 0);
  if (!(width > 0)) return release;
  const indent = parseFloat(css.textIndent) || 0;
  // Unchanged since it was last fitted: nothing to read again.
  const engineStarts = column.getAttribute(ENGINE_LINE_STARTS_ATTRIBUTE);
  const signature = `${familiesEpoch(doc, fittedFamilies.get(column))}|${width}|${indent}|${css.font}|${css.letterSpacing}|${engineStarts ?? ""}|${column.textContent}`;
  if (fittedFrom.get(column) === signature) return { column, kind: "keep" };
  const ctx = context(doc);
  if (!ctx) return release;
  const families = new Set<string>(familiesOf(css.fontFamily));
  const glyphs = readGlyphs(column, ctx, engineStarts !== null, families);
  fittedFamilies.set(column, families);
  // Recorded with the families now known: the next plan reads the same epoch.
  fittedFrom.set(column, `${familiesEpoch(doc, families)}|${width}|${indent}|${css.font}|${css.letterSpacing}|${engineStarts ?? ""}|${column.textContent}`);
  if (!glyphs || glyphs.length === 0) return release;
  const tokens = lineFitTokens(glyphs);
  if (!tokens) return release;
  const lines = engineStarts !== null ? engineLines(tokens, engineStarts) : powerPointLines(glyphs, tokens, width, indent);
  if (!lines) return release;
  const range = browserWidthRange(tokens, lines, exactMeasure(ctx, glyphs), indent);
  if (!range) return release;
  const target = powerPointFitWidth(width, range);
  if (target === null) return release;
  // Rounding moves a line by at most half a step a glyph.
  const longest = Math.max(...lines.map((line) => tokens[line.last].inkEnd - tokens[line.first].start));
  // The engine's lines may differ from the browser's by more than rounding
  // (kerning, a stand-in's widths): a few percent of the line, no more.
  const reach = engineStarts !== null ? width * 0.03 : (longest * POWERPOINT_ADVANCE_QUANTUM_PX) / 2 + 0.1;
  if (Math.abs(target - width) > reach) return release;
  return { column, kind: "fit", margin: width - target, glyphs, tokens, lines };
}

function releaseColumn(column: HTMLElement): void {
  if (!column.hasAttribute(POWERPOINT_LINE_FIT_ATTRIBUTE)) return;
  column.style.marginInlineEnd = "";
  column.removeAttribute(POWERPOINT_LINE_FIT_ATTRIBUTE);
}

/**
 * Fit paragraph text columns, all read first, then all written, then all
 * checked: a page of paragraphs is laid out a couple of times, not once a
 * paragraph. Returns how many now end their lines as PowerPoint does by a
 * width change.
 */
export function fitColumnsToPowerPointLines(columns: Iterable<HTMLElement>): number {
  const plans = [...columns].map(planColumn);
  const written: Extract<ColumnPlan, { kind: "fit" }>[] = [];
  for (const plan of plans) {
    if (plan.kind === "keep") continue;
    if (plan.kind === "release") {
      releaseColumn(plan.column);
      continue;
    }
    plan.column.style.marginInlineEnd = `${plan.margin}px`;
    plan.column.setAttribute(POWERPOINT_LINE_FIT_ATTRIBUTE, String(-plan.margin));
    written.push(plan);
  }
  const failed = written.filter(
    (plan) => !browserLinesMatch(plan.column.ownerDocument, plan.glyphs, plan.tokens, plan.lines),
  );
  for (const plan of failed) releaseColumn(plan.column);
  return written.length - failed.length + plans.filter((plan) => plan.kind === "keep" && plan.column.hasAttribute(POWERPOINT_LINE_FIT_ATTRIBUTE)).length;
}

/**
 * Fit one paragraph text column. Returns true when it now ends its lines as
 * PowerPoint does by a width change, false when it was left to the browser.
 */
export function fitColumnToPowerPointLines(column: HTMLElement): boolean {
  return fitColumnsToPowerPointLines([column]) > 0;
}

const pending = new Set<HTMLElement>();
let flushQueued = false;

/**
 * Ask for `columns` to be fitted with every other column asked for in the
 * same task: one batch after the render that drew them (a microtask, before
 * the browser paints), so a slide and its thumbnails are laid out a couple
 * of times in all. `after-paint`: in the batch run once the frame has painted
 * (`requestPowerPointLineFitAfterPaint`), for columns a reader does not wait on.
 */
export function requestPowerPointLineFit(columns: Iterable<HTMLElement>, when: "now" | "after-paint" = "now"): void {
  if (when === "after-paint") {
    requestPowerPointLineFitAfterPaint(columns);
    return;
  }
  for (const column of columns) pending.add(column);
  if (flushQueued || pending.size === 0) return;
  flushQueued = true;
  queueMicrotask(() => {
    flushQueued = false;
    const batch = [...pending].filter((column) => column.isConnected);
    pending.clear();
    fitColumnsToPowerPointLines(batch);
  });
}

const pendingAfterPaint = new Set<HTMLElement>();
let afterPaintQueued = false;

/**
 * Columns that are not what the reader looks at (a slide's thumbnails):
 * fitted the same way, in one batch once the frame now being built has
 * painted (a frame, then a task), so their reading of the layout does not
 * hold the slide's own paint. Where nothing paints frames (jsdom), as soon
 * as the columns asked for at once.
 */
function requestPowerPointLineFitAfterPaint(columns: Iterable<HTMLElement>): void {
  for (const column of columns) pendingAfterPaint.add(column);
  if (afterPaintQueued || pendingAfterPaint.size === 0) return;
  afterPaintQueued = true;
  const flush = () => {
    afterPaintQueued = false;
    const batch = [...pendingAfterPaint].filter((column) => column.isConnected);
    pendingAfterPaint.clear();
    fitColumnsToPowerPointLines(batch);
  };
  if (typeof requestAnimationFrame !== "function") {
    queueMicrotask(flush);
    return;
  }
  requestAnimationFrame(() => setTimeout(flush, 0));
}
