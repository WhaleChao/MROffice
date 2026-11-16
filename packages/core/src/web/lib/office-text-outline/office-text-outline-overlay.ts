/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The overlay that draws WordArt text outlines as Office draws them
 * (office-text-outline-css.ts says why): for each element that states an
 * outline (`--office-text-outline-color` / `-width` in its own style), the
 * glyphs of its text are placed where the browser laid each character, their
 * merged outlines (outline-client.ts) stroked once, centred and round, as
 * one SVG picture, which the element's `::after` shows over its words in the
 * element's own coordinates -- so a slide's zoom, a turned shape and the
 * element's own transform carry it with the words, and the words stay the
 * text that is typed in and selected.
 *
 * Only an element whose text changed (or whose characters moved within it,
 * a line wrapping anew) is measured again; the rest keep their picture.
 */
import { measureText, onTextEngineChange, resolveFace } from "../text-engine/measure";
import { splitFamilyList } from "../text-engine/face-index";
import { localOnlyFaceState } from "../text-engine/verify";
import { layoutProbeParent } from "../layout-probe-root";
import { glyphPath, onGlyphPaths, releaseOutlineGlyphs, type GlyphPath } from "./outline-client";
import {
  OFFICE_TEXT_OUTLINE_COLOR,
  OFFICE_TEXT_OUTLINE_FILLED,
  OFFICE_TEXT_OUTLINE_DASH,
  OFFICE_TEXT_OUTLINE_IMAGE,
  OFFICE_TEXT_OUTLINE_JOIN,
  OFFICE_TEXT_OUTLINE_RINGS,
  OFFICE_TEXT_OUTLINE_WIDTH,
  setOfficeTextOutlinePendingProbe,
  noteOfficeTextOutlineGlyphs,
  setOfficeTextOutlineCanvasBand,
  setOfficeTextOutlineGlyphLookup,
  setOfficeTextOutlineSvgGlyphs,
  officeTextOutlineStroke,
  officeTextOutlineStrokeOfMarks,
  type OfficeTextOutlineStroke,
  type OfficeTextOutlineSvgGlyph,
} from "./office-text-outline-css";

const MARK = OFFICE_TEXT_OUTLINE_WIDTH;
const BOX = ["--office-text-outline-x", "--office-text-outline-y", "--office-text-outline-w", "--office-text-outline-h"] as const;
const PT = 96 / 72;
/** Skia's synthetic oblique: x leans by a quarter of y. */
const SYNTHETIC_SKEW = 0.25;

const STYLE_ID = "mrnote-office-text-outline";
const STYLESHEET = `
[style*="${OFFICE_TEXT_OUTLINE_IMAGE}"]{-webkit-text-stroke-width:0!important;${OFFICE_TEXT_OUTLINE_RINGS}:opacity(1)!important}
[style*="${OFFICE_TEXT_OUTLINE_FILLED}"],[style*="${OFFICE_TEXT_OUTLINE_FILLED}"] *,textarea[data-office-text-outline-filled]{-webkit-text-fill-color:transparent!important}
[style*="${OFFICE_TEXT_OUTLINE_FILLED}"][style*="background-clip"],[style*="${OFFICE_TEXT_OUTLINE_FILLED}"] [style*="background-clip"],textarea[data-office-text-outline-filled][style*="background-clip"]{background-image:none!important}
textarea[data-office-text-outline-mirrored]{-webkit-text-stroke-width:0!important;${OFFICE_TEXT_OUTLINE_RINGS}:opacity(1)!important}
[data-office-text-warp-source][style*="${OFFICE_TEXT_OUTLINE_IMAGE}"]::after,[data-office-text-warp-source] [style*="${OFFICE_TEXT_OUTLINE_IMAGE}"]::after{visibility:hidden}
[style*="${OFFICE_TEXT_OUTLINE_IMAGE}"]::after{content:"";position:absolute;pointer-events:none;user-select:none;-webkit-text-stroke:0;left:var(${BOX[0]});top:var(${BOX[1]});width:var(${BOX[2]});height:var(${BOX[3]});background:var(${OFFICE_TEXT_OUTLINE_IMAGE}) no-repeat 0 0/100% 100%}
`;

interface State {
  /** The inputs and the geometry the picture was drawn from. */
  drawn: string;
  /** Characters' arrangement within the element, scale-free. */
  shape: string;
  /** Waiting on glyphs or faces: drawn again when they come. */
  waiting: boolean;
  /** The overlay made the element its picture's containing block. */
  positioned?: boolean;
}

let started: Document | null = null;
const tracked = new Map<HTMLElement, State>();
const dirty = new Set<HTMLElement>();
let frame = 0;
let sweepTimer: ReturnType<typeof setInterval> | null = null;
let emptySince = 0;
/** The style each element was left with by the overlay's own writes: a
 * mutation that leaves it so is ours, not a change to answer (observer
 * records arrive after the write, so a flag set around it cannot tell). */
const leftStyle = new WeakMap<Element, string | null>();
function wrote(element: Element): void {
  leftStyle.set(element, element.getAttribute("style"));
}

export function startOfficeTextOutlineOverlay(doc: Document): void {
  if (started === doc || !doc.defaultView) return;
  started = doc;
  setOfficeTextOutlinePendingProbe(() => frame !== 0 || dirty.size > 0 || [...tracked.values()].some((state) => state.waiting));
  if (!doc.getElementById(STYLE_ID)) {
    const style = doc.createElement("style");
    style.id = STYLE_ID;
    style.textContent = STYLESHEET;
    doc.head.append(style);
  }
  const observer = new MutationObserver(onMutations);
  observer.observe(doc.documentElement, {
    subtree: true,
    childList: true,
    characterData: true,
    attributes: true,
    attributeFilter: ["style", "class"],
  });
  onGlyphPaths(() => {
    markWaiting();
    noteOfficeTextOutlineGlyphs();
  });
  // Every answer the face resolution gave as pending may have come: a face
  // file read, a face verified (or found installed), the order the browser
  // takes faces in. Words that waited on any of them are drawn then (a
  // document's static words are not redrawn by anything else).
  onTextEngineChange(() => {
    markWaiting();
    noteOfficeTextOutlineGlyphs();
  });
  doc.fonts?.addEventListener?.("loadingdone", () => {
    fontBoxes.clear();
    for (const element of tracked.keys()) dirty.add(element);
    schedule();
  });
  doc.defaultView.addEventListener("resize", () => sweep(), { passive: true });
  // A textarea's words change and scroll with no change to the page.
  const onField = (event: Event) => {
    const target = event.target;
    if (target instanceof HTMLTextAreaElement && mirrors.has(target)) {
      dirty.add(target);
      schedule();
    }
  };
  doc.addEventListener("input", onField, true);
  doc.addEventListener("scroll", onField, { capture: true, passive: true });
  scanAll(doc);
}

/** Tests: draw every marked element now (no frame wait). */
export function drawOfficeTextOutlinesNow(doc: Document): void {
  scanAll(doc);
  if (frame) cancelAnimationFrame(frame);
  frame = 0;
  run();
}

function scanAll(doc: Document): void {
  for (const element of doc.querySelectorAll<HTMLElement>(`[style*="${MARK}"]`)) dirty.add(element);
  schedule();
}

function marked(element: Element): element is HTMLElement {
  return element instanceof HTMLElement && (element.getAttribute("style") ?? "").includes(MARK);
}

/** The element whose outline draws this node's text: its nearest marked ancestor. */
function ownerOf(node: Node | null): HTMLElement | null {
  for (let at: Node | null = node; at; at = at.parentNode) {
    if (at instanceof Element && marked(at)) return at as HTMLElement;
  }
  return null;
}

function onMutations(records: MutationRecord[]): void {
  let any = false;
  for (const record of records) {
    if (record.type !== "attributes") noteScope(record.target);
    if (record.type === "attributes") {
      const target = record.target as Element;
      if (leftStyle.has(target) && leftStyle.get(target) === target.getAttribute("style")) continue;
      if (marked(target) || tracked.has(target as HTMLElement)) {
        dirty.add(target as HTMLElement);
        any = true;
      }
      continue;
    }
    if (record.type === "characterData") {
      const owner = ownerOf(record.target);
      if (owner) {
        dirty.add(owner);
        any = true;
      }
      continue;
    }
    const owner = ownerOf(record.target);
    if (owner) {
      dirty.add(owner);
      any = true;
    }
    for (const node of record.addedNodes) {
      if (!(node instanceof Element)) continue;
      if (marked(node)) dirty.add(node as HTMLElement);
      for (const inner of node.querySelectorAll<HTMLElement>(`[style*="${MARK}"]`)) dirty.add(inner);
      any = true;
    }
    if (record.removedNodes.length) any = true;
  }
  if (any || shapeScopes.size) schedule();
}

/** Where characters may have moved within an outlined element with no
 * change to it: a change nearby (a line before it wrapping anew), or
 * anywhere at all (the sweep). */
let shapeCheck = false;
const shapeScopes = new Set<Node>();

function noteScope(node: Node): void {
  if (tracked.size) shapeScopes.add(node);
}

/** The block a change lays out anew: the nearest ancestor that is not
 * inline (lines rewrap within it; what follows it only moves). */
function blockOf(node: Node): Node {
  for (let at: Node | null = node.nodeType === Node.ELEMENT_NODE ? node : node.parentNode; at; at = at.parentNode) {
    if (!(at instanceof Element)) return at;
    const display = getComputedStyle(at).display;
    if (display !== "inline" && display !== "contents") return at;
  }
  return node;
}

function schedule(): void {
  if (frame || !started?.defaultView) return;
  frame = started.defaultView.requestAnimationFrame(() => {
    frame = 0;
    run();
  });
  if (!sweepTimer) sweepTimer = setInterval(sweep, 400);
}

function markWaiting(): void {
  let any = false;
  for (const [element, state] of tracked) {
    if (state.waiting) {
      dirty.add(element);
      any = true;
    }
  }
  if (any) schedule();
}

/** Characters can move within an element with no change to it (its box
 * resized, a font arrived): its shape is compared now and then. */
function sweep(): void {
  if (!tracked.size && !dirty.size) {
    if (sweepTimer && Date.now() - emptySince > 30_000) {
      clearInterval(sweepTimer);
      sweepTimer = null;
      // No outlined words left on the page: what was kept for them goes.
      releaseOutlineGlyphs();
      fontBoxes.clear();
    }
    return;
  }
  if (started?.visibilityState === "hidden") return;
  shapeCheck = true;
  schedule();
}

function run(): void {
  if (shapeCheck || shapeScopes.size) {
    const scopes = shapeCheck ? null : [...new Set([...shapeScopes].filter((node) => node.isConnected).map(blockOf))];
    shapeCheck = false;
    shapeScopes.clear();
    for (const [element, state] of tracked) {
      if (dirty.has(element)) continue;
      if (!element.isConnected) {
        tracked.delete(element);
        if (element instanceof HTMLTextAreaElement) dropMirror(element);
        continue;
      }
      if (element instanceof HTMLTextAreaElement) {
        // Its words change with no change to the page (a value set by script).
        const mirror = mirrors.get(element);
        if (element.value !== state.drawn || !mirror?.isConnected) dirty.add(element);
        continue;
      }
      if (scopes && !scopes.some((scope) => scope.contains(element))) continue;
      if (shapeOf(element) !== state.shape) dirty.add(element);
    }
  }
  const work = [...dirty];
  dirty.clear();
  drawAll(work);
  if (!tracked.size) emptySince = Date.now();
}

/** The element's fragments relative to its first, scale-free: what changes
 * when its characters move within it. */
function shapeOf(element: HTMLElement): string {
  const rects = element.getClientRects();
  if (!rects.length) return "";
  const first = rects[0];
  const unit = first.height || 1;
  // Shown or hidden is part of it: words laid out hidden are drawn once shown.
  let shape = `${getComputedStyle(element).visibility === "hidden" ? "h" : ""}${rects.length}`;
  for (const rect of rects) {
    shape += `|${Math.round(((rect.left - first.left) / unit) * 100)},${Math.round(((rect.top - first.top) / unit) * 100)},${Math.round((rect.width / unit) * 100)}`;
  }
  return shape;
}

function clear(element: HTMLElement): void {
  const state = tracked.get(element);
  tracked.delete(element);
  if (!element.style.getPropertyValue(OFFICE_TEXT_OUTLINE_IMAGE) && !state?.positioned && !element.style.getPropertyValue(OFFICE_TEXT_OUTLINE_FILLED)) return;
  element.style.removeProperty(OFFICE_TEXT_OUTLINE_IMAGE);
  element.style.removeProperty(OFFICE_TEXT_OUTLINE_FILLED);
  for (const name of BOX) element.style.removeProperty(name);
  if (state?.positioned && element.style.position === "relative") element.style.removeProperty("position");
  wrote(element);
}

/** Whether making the element `position: relative` moves nothing: no box
 * inside it is placed against an ancestor (absolutely or fixed). */
function positionable(element: HTMLElement): boolean {
  for (const inner of element.querySelectorAll("*")) {
    const position = getComputedStyle(inner).position;
    if (position === "absolute" || position === "fixed") return false;
  }
  return true;
}

// --- geometry -----------------------------------------------------------------

/** The element's linear map to the viewport (2-D), or null under a 3-D one. */
function linearChain(element: Element, memo: Map<Element, readonly number[] | null>): readonly number[] | null {
  const known = memo.get(element);
  if (known !== undefined) return known;
  const parent = element.parentElement;
  const outer = parent ? linearChain(parent, memo) : [1, 0, 0, 1];
  if (!outer) {
    memo.set(element, null);
    return null;
  }
  const style = getComputedStyle(element);
  let m = [1, 0, 0, 1];
  const times = (n: readonly number[]) => {
    m = [m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1], m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3]];
  };
  // CSS applies translate, rotate, scale, then transform.
  const rotate = style.rotate;
  if (rotate && rotate !== "none") {
    const parts = rotate.trim().split(/\s+/);
    if (parts.length !== 1 && !(parts.length === 2 && parts[0] === "z") && !(parts.length === 4 && parts[0] === "0" && parts[1] === "0")) {
      memo.set(element, null);
      return null;
    }
    const angle = parseAngle(parts[parts.length - 1]);
    if (angle === null) {
      memo.set(element, null);
      return null;
    }
    const sign = parts.length === 4 && Number(parts[2]) < 0 ? -1 : 1;
    const c = Math.cos(angle * sign), s = Math.sin(angle * sign);
    times([c, s, -s, c]);
  }
  const scale = style.scale;
  if (scale && scale !== "none") {
    const [sx, sy = sx] = scale.trim().split(/\s+/).map(Number);
    times([sx, 0, 0, sy]);
  }
  const transform = style.transform;
  if (transform && transform !== "none") {
    const matrix = new DOMMatrixReadOnly(transform);
    if (!matrix.is2D) {
      memo.set(element, null);
      return null;
    }
    times([matrix.a, matrix.b, matrix.c, matrix.d]);
  }
  const zoom = Number.parseFloat((style as CSSStyleDeclaration & { zoom?: string }).zoom ?? "1");
  if (Number.isFinite(zoom) && zoom > 0 && zoom !== 1) times([zoom, 0, 0, zoom]);
  const total = [
    outer[0] * m[0] + outer[2] * m[1],
    outer[1] * m[0] + outer[3] * m[1],
    outer[0] * m[2] + outer[2] * m[3],
    outer[1] * m[2] + outer[3] * m[3],
  ];
  memo.set(element, total);
  return total;
}

function parseAngle(value: string): number | null {
  const match = /^(-?[\d.]+(?:e-?\d+)?)(deg|rad|turn|grad)$/.exec(value);
  if (!match) return value === "0" ? 0 : null;
  const n = Number(match[1]);
  return match[2] === "deg" ? (n * Math.PI) / 180 : match[2] === "rad" ? n : match[2] === "turn" ? n * 2 * Math.PI : (n * Math.PI) / 200;
}

interface FontBox {
  /** The text box's height (the primary face's content area), px. */
  height: number;
  /** Its baseline below its top, px. */
  baseline: number;
}
const fontBoxes = new Map<string, FontBox>();

/** A text box's height and baseline in a font, as this browser lays it. */
function fontBox(style: CSSStyleDeclaration, doc: Document): FontBox | null {
  const font = `${style.fontStyle} ${style.fontWeight} ${style.fontStretch} ${style.fontSize} ${style.fontFamily}`;
  const known = fontBoxes.get(font);
  if (known) return known;
  const host = doc.createElement("div");
  host.style.cssText = "position:absolute;left:0;top:0;white-space:pre;line-height:normal;";
  const span = doc.createElement("span");
  span.style.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
  span.style.fontStretch = style.fontStretch;
  span.textContent = "Hxg";
  const marker = doc.createElement("span");
  marker.style.cssText = "display:inline-block;width:0;height:0;vertical-align:baseline;";
  span.append(marker);
  host.append(span);
  layoutProbeParent(doc).append(host);
  const range = doc.createRange();
  range.selectNodeContents(span.firstChild!);
  const text = range.getBoundingClientRect();
  const base = marker.getBoundingClientRect();
  host.remove();
  if (!(text.height > 0)) return null;
  const box = { height: text.height, baseline: base.top - text.top };
  let loaded = true;
  try {
    loaded = doc.fonts ? doc.fonts.check(`${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`) : true;
  } catch {
    loaded = true;
  }
  if (loaded) fontBoxes.set(font, box);
  return box;
}

// --- drawing ------------------------------------------------------------------

interface Glyph {
  readonly path: GlyphPath;
  readonly id: string;
  /** Font units (y up) to the element's px: a b c d e f, as SVG's matrix(). */
  readonly m: readonly [number, number, number, number, number, number];
  /** px per font unit. */
  readonly scale: number;
  /** Extra width a synthesized bold adds to the band, px. */
  readonly bold: number;
  /** The colour the words are filled with here (null: no fill). */
  readonly fill?: string | null;
}

/** Whether a character stands upright in vertical text (`text-orientation`;
 * `mixed` by Unicode's Vertical_Orientation, its U and Tu ranges). */
function uprightIn(orientation: string, codePoint: number): boolean {
  if (orientation === "upright") return true;
  if (orientation.startsWith("sideways")) return false;
  return (
    (codePoint >= 0x1100 && codePoint <= 0x11ff) ||
    (codePoint >= 0x2e80 && codePoint <= 0xa4cf) ||
    (codePoint >= 0xa960 && codePoint <= 0xa97f) ||
    (codePoint >= 0xac00 && codePoint <= 0xd7ff) ||
    (codePoint >= 0xf900 && codePoint <= 0xfaff) ||
    (codePoint >= 0xfe30 && codePoint <= 0xfe4f) ||
    (codePoint >= 0xff01 && codePoint <= 0xff60) ||
    (codePoint >= 0xffe0 && codePoint <= 0xffe7) ||
    (codePoint >= 0x20000 && codePoint <= 0x3fffd)
  );
}

/** Characters vertical text draws with a vertical alternate glyph (GSUB
 * `vert`: CJK punctuation, brackets, the long vowel mark): not the glyph the
 * horizontal map names, so not outlined from it. */
function verticalAlternate(codePoint: number): boolean {
  return (
    (codePoint >= 0x3000 && codePoint <= 0x303f) ||
    (codePoint >= 0xfe10 && codePoint <= 0xfe1f) ||
    (codePoint >= 0xff01 && codePoint <= 0xff0f) ||
    (codePoint >= 0xff1a && codePoint <= 0xff20) ||
    (codePoint >= 0xff3b && codePoint <= 0xff40) ||
    (codePoint >= 0xff5b && codePoint <= 0xff65) ||
    codePoint === 0x30fc ||
    codePoint === 0xff70 ||
    (codePoint >= 0x2010 && codePoint <= 0x2026)
  );
}

interface Fallback {
  readonly text: string;
  readonly x: number;
  readonly y: number;
  readonly font: string;
  readonly fill?: string | null;
}

const svgEscape = (value: string) =>
  value.replace(/[&<>"']/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;" })[ch]!);
const p6 = (value: number) => String(Number(value.toPrecision(6)));
const n2 = (value: number) => {
  const rounded = Math.round(value * 100) / 100;
  return Object.is(rounded, -0) ? "0" : String(rounded);
};

let segmenter: Intl.Segmenter | null = null;
function graphemes(text: string): { segment: string; index: number }[] {
  segmenter ??= typeof Intl !== "undefined" && "Segmenter" in Intl ? new Intl.Segmenter(undefined, { granularity: "grapheme" }) : null;
  if (!segmenter) return Array.from(text, (segment, index) => ({ segment, index }));
  return Array.from(segmenter.segment(text), ({ segment, index }) => ({ segment, index }));
}

/** Scripts the browser shapes (joins, reorders, places marks): not placed
 * character by character. */
/**
 * Whether the page draws a character in the served file the overlay reads
 * its outline from: not where the family's installed font draws it (a
 * machine with Office's Calibri draws Calibri where the served file is
 * Carlito -- the same widths, other outlines: the band would not follow the
 * glyphs filled under it). There the words are filled from the served
 * file too, under the band (the page's own words see-through, still the
 * text that is selected and typed in), so fill and band are one glyph.
 */
export function drawsServedFile(resolved: { face: { localNames: readonly string[] }; installed?: true }): boolean | "pending" {
  if (resolved.installed) return false;
  if (!resolved.face.localNames.length) return true;
  const local = localOnlyFaceState(resolved.face as Parameters<typeof localOnlyFaceState>[0]);
  return local === "pending" ? "pending" : local === "missing";
}

const SHAPED = /[֐-ࣿऀ-෿ༀ-࿿က-႟ក-៿יִ-﷿ﹰ-﻿]/u;















function untransformed<T>(elements: readonly HTMLElement[], measure: () => T): T {
  const lifted: { node: HTMLElement; name: string; value: string; priority: string }[] = [];
  const seen = new Set<HTMLElement>();
  try {
    for (const element of elements) {
      for (let at: HTMLElement | null = element; at && !seen.has(at); at = at.parentElement) {
        seen.add(at);
        const style = getComputedStyle(at);
        for (const name of ["transform", "rotate", "scale", "translate"] as const) {
          const value = style.getPropertyValue(name);
          if (!value || value === "none") continue;
          lifted.push({ node: at, name, value: at.style.getPropertyValue(name), priority: at.style.getPropertyPriority(name) });
        }
      }
    }
    for (const { node, name } of lifted) node.style.setProperty(name, "none", "important");
    return measure();
  } finally {
    const restored = new Set<HTMLElement>();
    for (const { node, name, value, priority } of lifted.reverse()) {
      if (value) node.style.setProperty(name, value, priority);
      else node.style.removeProperty(name);
      restored.add(node);
    }
    for (const node of restored) wrote(node);
  }
}

/** Whether the element is drawn measured with its transforms lifted: one
 * under a 3-D turn. */
function turnedIn3D(element: HTMLElement, memo: Map<Element, readonly number[] | null>): boolean {
  return element.isConnected && marked(element) && linearChain(element, memo) === null;
}

function drawCaught(element: HTMLElement): void {
  try {
    drawFlat(element);
  } catch {
    // Whatever the overlay cannot read, the element's own stroke draws.
    clear(element);
  }
}

/** Draws the elements: the flat ones as they stand, then every one under a
 * 3-D turn inside one lift of all their transforms. */
function drawAll(elements: readonly HTMLElement[]): void {
  // Sorted before anything is written, so the chains are read in one go.
  const memo = new Map<Element, readonly number[] | null>();
  const turned: HTMLElement[] = [];
  const flat: HTMLElement[] = [];
  for (const element of elements) {
    let inTurn = false;
    try {
      inTurn = turnedIn3D(element, memo);
    } catch {
      inTurn = false;
    }
    (inTurn ? turned : flat).push(element);
  }
  for (const element of flat) drawCaught(element);
  if (!turned.length) return;
  try {
    untransformed(turned, () => {
      for (const element of turned) drawCaught(element);
    });
  } catch {
    for (const element of turned) clear(element);
  }
}

// --- a textarea's words ----------------------------------------------------------

/** The words a textarea draws, mirrored beside it in a box of its own the
 * overlay can measure (a textarea's text is not in the page), its own
 * words transparent: the textarea keeps the keys, the caret and the
 * selection, and the mirror's picture draws the outline over its words. */
const mirrors = new Map<HTMLTextAreaElement, HTMLDivElement>();
const fieldOfMirror = new WeakMap<Element, HTMLTextAreaElement>();
/** Set on a textarea while its mirror's picture fills its words too. */
const FIELD_FILLED = "data-office-text-outline-filled";
const MIRRORED = "data-office-text-outline-mirrored";
const MIRROR = "data-office-text-outline-mirror";
const COPIED = [
  "font-family", "font-size", "font-weight", "font-style", "font-stretch", "font-kerning", "font-variant",
  "font-feature-settings", "font-variation-settings", "font-synthesis", "font-variant-ligatures", "font-size-adjust",
  "letter-spacing", "word-spacing", "line-height", "text-align", "text-indent", "text-transform", "text-rendering",
  "white-space", "word-break", "overflow-wrap", "tab-size", "direction", "unicode-bidi", "writing-mode",
  "text-orientation", "hyphens", "line-break", "padding-top", "padding-right", "padding-bottom", "padding-left",
  "border-top-width", "border-right-width", "border-bottom-width", "border-left-width",
  OFFICE_TEXT_OUTLINE_COLOR, OFFICE_TEXT_OUTLINE_WIDTH,
] as const;

function dropMirror(field: HTMLTextAreaElement): void {
  const mirror = mirrors.get(field);
  mirrors.delete(field);
  if (mirror) {
    tracked.delete(mirror);
    mirror.remove();
  }
  if (field.hasAttribute(MIRRORED)) field.removeAttribute(MIRRORED);
  if (field.hasAttribute(FIELD_FILLED)) field.removeAttribute(FIELD_FILLED);
}

function mirrorField(field: HTMLTextAreaElement, style: CSSStyleDeclaration): void {
  const parent = field.parentElement;
  if (!parent || style.display === "none") {
    dropMirror(field);
    return;
  }
  let mirror = mirrors.get(field);
  if (!mirror || mirror.parentElement !== parent) {
    mirror?.remove();
    mirror = field.ownerDocument.createElement("div");
    mirror.setAttribute(MIRROR, "");
    mirror.setAttribute("aria-hidden", "true");
    mirror.append(field.ownerDocument.createElement("div"));
    mirrors.set(field, mirror);
    fieldOfMirror.set(mirror, field);
  }
  if (mirror.previousElementSibling !== field) field.after(mirror);
  const box = mirror.style;
  for (const name of COPIED) box.setProperty(name, style.getPropertyValue(name));
  box.setProperty("position", "absolute");
  box.setProperty("box-sizing", "content-box");
  box.setProperty("margin", "0");
  box.setProperty("border-style", "solid");
  box.setProperty("border-color", "transparent");
  box.setProperty("overflow", "hidden");
  box.setProperty("pointer-events", "none");
  box.setProperty("user-select", "none");
  box.setProperty("color", "transparent");
  box.setProperty("-webkit-text-fill-color", "transparent");
  box.setProperty("-webkit-text-stroke", "0 transparent");
  box.setProperty("background", "none");
  box.setProperty("left", `${field.offsetLeft}px`);
  box.setProperty("top", `${field.offsetTop}px`);
  const padX = (Number.parseFloat(style.paddingLeft) || 0) + (Number.parseFloat(style.paddingRight) || 0);
  const padY = (Number.parseFloat(style.paddingTop) || 0) + (Number.parseFloat(style.paddingBottom) || 0);
  box.setProperty("width", `${Math.max(0, field.clientWidth - padX)}px`);
  box.setProperty("height", `${Math.max(0, field.clientHeight - padY)}px`);
  box.setProperty("z-index", style.zIndex === "auto" ? "auto" : style.zIndex);
  box.setProperty("transform", style.transform);
  box.setProperty("transform-origin", style.transformOrigin);
  const words = mirror.firstElementChild as HTMLDivElement;
  // The trailing break a textarea shows as an empty last line.
  const text = field.value.endsWith("\n") ? `${field.value}\u200b` : field.value;
  if (words.textContent !== text) words.textContent = text;
  words.style.setProperty("transform", field.scrollTop || field.scrollLeft ? `translate(${-field.scrollLeft}px, ${-field.scrollTop}px)` : "none");
  wrote(mirror);
  wrote(words);
  tracked.set(field, { drawn: field.value, shape: "", waiting: false });
}

function drawFlat(element: HTMLElement): void {
  if (element instanceof HTMLTextAreaElement) {
    if (!element.isConnected || !marked(element)) {
      dropMirror(element);
      tracked.delete(element);
      return;
    }
    const fieldStyle = getComputedStyle(element);
    const fieldWidth = Number.parseFloat(fieldStyle.getPropertyValue(OFFICE_TEXT_OUTLINE_WIDTH));
    if (!(fieldWidth > 0)) {
      dropMirror(element);
      tracked.delete(element);
      return;
    }
    // Its fill read as its own (the mirror's picture fills it otherwise).
    const fieldFilled = element.hasAttribute(FIELD_FILLED);
    if (fieldFilled) element.removeAttribute(FIELD_FILLED);
    mirrorField(element, fieldStyle);
    const mirror = mirrors.get(element);
    if (mirror) {
      drawFlat(mirror);
      if (mirror.style.getPropertyValue(OFFICE_TEXT_OUTLINE_FILLED)) element.setAttribute(FIELD_FILLED, "");
      // The textarea's own stand-in stroke gives way once the mirror draws.
      const drawn = Boolean(mirror.style.getPropertyValue(OFFICE_TEXT_OUTLINE_IMAGE));
      if (drawn !== element.hasAttribute(MIRRORED)) {
        if (drawn) element.setAttribute(MIRRORED, "");
        else element.removeAttribute(MIRRORED);
      }
    }
    return;
  }
  // The words' own fill see-through while the picture fills them
  // (drawsServedFile): lifted while the element is read, so the fills read
  // are its own.
  const wasFilled = element.style.getPropertyValue(OFFICE_TEXT_OUTLINE_FILLED) !== "";
  if (wasFilled) element.style.removeProperty(OFFICE_TEXT_OUTLINE_FILLED);
  let filled = false;
  try {
    filled = drawElement(element);
  } finally {
    if (filled) element.style.setProperty(OFFICE_TEXT_OUTLINE_FILLED, "1");
    if (filled || wasFilled) wrote(element);
  }
}

function drawElement(element: HTMLElement): boolean {
  if (!element.isConnected || !marked(element)) {
    clear(element);
    return false;
  }
  const doc = element.ownerDocument;
  const style = getComputedStyle(element);
  const color = style.getPropertyValue(OFFICE_TEXT_OUTLINE_COLOR).trim();
  const widthPt = Number.parseFloat(style.getPropertyValue(OFFICE_TEXT_OUTLINE_WIDTH));
  // Its join and dash (Format Text Effects > Text Outline), round and solid
  // where it states none.
  const stroke = officeTextOutlineStrokeOfMarks(
    style.getPropertyValue(OFFICE_TEXT_OUTLINE_JOIN), style.getPropertyValue(OFFICE_TEXT_OUTLINE_DASH),
  );
  const tag = element.tagName;
  if (!/^#[0-9a-f]{6}([0-9a-f]{2})?$/i.test(color) || !(widthPt > 0) || tag === "TEXTAREA" || tag === "INPUT"
    || !/^(horizontal-tb|vertical-rl|vertical-lr)$/.test(style.writingMode) || style.display === "none") {
    clear(element);
    return false;
  }
  const memo = new Map<Element, readonly number[] | null>();
  const L = linearChain(element, memo);
  if (!L) {
    clear(element);
    return false;
  }
  const det = L[0] * L[3] - L[1] * L[2];
  if (!(Math.abs(det) > 1e-9)) {
    clear(element);
    return false;
  }
  const inverse = [L[3] / det, -L[1] / det, -L[2] / det, L[0] / det];
  const axisAligned = Math.abs(L[1]) < 1e-6 && Math.abs(L[2]) < 1e-6 && L[0] > 0 && L[3] > 0;
  // The ::after's containing block: the element's padding box (the first
  // fragment's, for an inline element), its top left the local origin.
  const inline = style.display === "inline";
  const rects = element.getClientRects();
  if (!rects.length) {
    clear(element);
    return false;
  }
  // The picture is placed against the element's own padding box: an
  // element placed in flow becomes `position: relative` (which moves nothing
  // of its own) unless a box inside it is placed against an ancestor.
  let positioned = tracked.get(element)?.positioned ?? false;
  if (style.position === "static") {
    if (!positionable(element)) {
      clear(element);
      return false;
    }
    element.style.setProperty("position", "relative");
    wrote(element);
    positioned = true;
  }
  const borderLeft = Number.parseFloat(style.borderLeftWidth) || 0;
  const borderTop = Number.parseFloat(style.borderTopWidth) || 0;
  let originX: number, originY: number;
  if (axisAligned) {
    const box = inline ? rects[0] : element.getBoundingClientRect();
    originX = box.left + borderLeft * L[0];
    originY = box.top + borderTop * L[3];
  } else {
    if (inline && rects.length > 1) {
      clear(element);
      return false;
    }
    const box = element.getBoundingClientRect();
    const cx = element.offsetWidth / 2 - borderLeft, cy = element.offsetHeight / 2 - borderTop;
    originX = (box.left + box.right) / 2 - (L[0] * cx + L[2] * cy);
    originY = (box.top + box.bottom) / 2 - (L[1] * cx + L[3] * cy);
  }
  const toLocal = (x: number, y: number): [number, number] => {
    const dx = x - originX, dy = y - originY;
    return [inverse[0] * dx + inverse[2] * dy, inverse[1] * dx + inverse[3] * dy];
  };

  const glyphs: Glyph[] = [];
  const fallbacks: Fallback[] = [];
  let waiting = false;
  let unsupported = false;
  // Whether an installed font draws any of the words (drawsServedFile), and
  // whether every fill is one the picture can draw: a colour, or a gradient
  // or pattern clipped to the words (background-clip: text) of an element
  // that is not turned.
  let installedDrawn = false;
  // Words not shown yet (a document laid out hidden first): looked at again
  // once they are (shapeOf states the element's visibility).
  let unseen = false;
  let fillable = !element.querySelector(`[style*="${OFFICE_TEXT_OUTLINE_WIDTH}:"]`);
  const fills = new Map<Element, string | null>();
  const clips: ClipFill[] = [];
  const clipIndex = new Map<Element, number>();
  // A textarea's mirror fills its words as the textarea does.
  const field = fieldOfMirror.get(element) ?? null;
  const fillOf = (parent: Element, own: CSSStyleDeclaration): string | null => {
    if (fills.has(parent)) return fills.get(parent)!;
    const source = field ? getComputedStyle(field) : own;
    let clipped: Element | null = null;
    if (field) {
      if (clipsToWords(source)) clipped = field;
    } else {
      for (let at: Element | null = parent; at; at = at === element ? null : at.parentElement) {
        if (clipsToWords(at === parent ? own : getComputedStyle(at))) {
          clipped = at;
          break;
        }
      }
    }
    let fill: string | null = null;
    if (clipped) {
      let index = clipIndex.get(clipped);
      if (index === undefined) {
        const clip = axisAligned ? clipFillOf(clipped, clipped === field ? element : null, toLocal, L) : null;
        if (clip) {
          index = clips.length;
          clips.push(clip);
          clipIndex.set(clipped, index);
        } else fillable = false;
      }
      if (index !== undefined) fill = `${CLIP_FILL}${index}`;
    } else {
      fill = source.webkitTextFillColor || source.color;
      if (/^(transparent|rgba\([^)]*,\s*0\))$/.test(fill.replace(/\s+/g, " ").trim())) fill = null;
    }
    fills.set(parent, fill);
    return fill;
  };
  const range = doc.createRange();
  const walker = doc.createTreeWalker(element, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null) {
    const parent = node.parentElement;
    if (!parent || ownerOf(node) !== element) continue;
    const text = node.data;
    if (!text.trim()) continue;
    if (SHAPED.test(text)) {
      unsupported = true;
      break;
    }
    const own = parent === element ? style : getComputedStyle(parent);
    if (own.visibility === "hidden" || own.display === "none") {
      unseen = true;
      continue;
    }
    const box = fontBox(own, doc);
    if (!box) continue;
    const families = splitFamilyList(own.fontFamily);
    const weight = Number.parseInt(own.fontWeight, 10) || 400;
    const italic = own.fontStyle !== "normal";
    const sizePx = Number.parseFloat(own.fontSize);
    const fontCss = `${own.fontStyle} ${own.fontWeight} ${own.fontSize} ${own.fontFamily}`;
    for (const { segment, index } of graphemes(text)) {
      if (!segment.trim()) continue;
      const codePoint = segment.codePointAt(0)!;
      range.setStart(node, index);
      range.setEnd(node, index + segment.length);
      const charRects = range.getClientRects();
      if (!charRects.length) continue;
      const r = charRects[0];
      const vertical = own.writingMode !== "horizontal-tb";
      let left: number, top: number, w: number;
      if (axisAligned) {
        [left, top] = toLocal(r.left, r.top);
        w = r.width / L[0];
      } else {
        // The character's box turned: its bounding box's centre is its
        // centre; its advance from that box and its known thickness (the
        // line's: the face's content area).
        const [cx, cy] = toLocal((r.left + r.right) / 2, (r.top + r.bottom) / 2);
        const a = Math.abs(L[0]), b = Math.abs(L[1]), c = Math.abs(L[2]), d = Math.abs(L[3]);
        const thick = box.height;
        let h: number;
        if (vertical) {
          w = thick;
          h = c >= d ? (r.width - a * w) / c : (r.height - b * w) / d;
        } else {
          h = thick;
          w = a >= b ? (r.width - c * h) / a : (r.height - d * h) / b;
        }
        left = cx - w / 2;
        top = cy - h / 2;
      }
      const baseline = top + box.baseline;
      const upright = vertical && uprightIn(own.textOrientation, codePoint);
      if (vertical && upright && verticalAlternate(codePoint)) continue;
      const resolved = resolveFace(families, weight, italic, codePoint === 0xa0 ? 0x20 : codePoint, true, false);
      if (resolved === "pending") {
        waiting = true;
        continue;
      }
      if (typeof resolved === "string") {
        // A face the overlay cannot read (a system or installed font): the
        // SVG picture draws the character itself, in that font, stroked
        // round (its contours as they stand).
        fallbacks.push({ text: segment, x: left, y: baseline, font: fontCss, fill: fillOf(parent, own) });
        continue;
      }
      if (resolved.face.wght !== undefined || resolved.metrics.variationTables) {
        fallbacks.push({ text: segment, x: left, y: baseline, font: fontCss, fill: fillOf(parent, own) });
        continue;
      }
      const served = drawsServedFile(resolved);
      if (served === "pending") {
        waiting = true;
        continue;
      }
      if (!served) installedDrawn = true;
      const url = resolved.face.url!;
      const path = glyphPath(url, resolved.glyph);
      if (path === "pending") {
        waiting = true;
        continue;
      }
      if (!path) {
        fallbacks.push({ text: segment, x: left, y: baseline, font: fontCss, fill: fillOf(parent, own) });
        continue;
      }
      if (!path.d) continue;
      const scale = (sizePx * resolved.face.sizeAdjust) / path.unitsPerEm;
      const synthesizedItalic = italic && resolved.face.style === "normal" && !own.getPropertyValue("font-synthesis-style").includes("none");
      // Skia's synthetic bold widens each glyph by its size over 24 (at 9 px)
      // to over 32 (at 36 px and up), half on each side.
      const ratio = sizePx <= 9 ? 1 / 24 : sizePx >= 36 ? 1 / 32 : 1 / 24 + ((sizePx - 9) / 27) * (1 / 32 - 1 / 24);
      const skew = synthesizedItalic ? SYNTHETIC_SKEW * scale : 0;
      let m: Glyph["m"];
      if (!vertical) {
        m = [scale, 0, skew, -scale, left, baseline];
      } else if (upright) {
        // Upright: its advance centred on its own face's central baseline
        // (the shared alphabetic baseline, moved by half its ascent less its
        // descent, as the browser rounds them), its vertical origin at the
        // top of its box.
        const alphabetic = left + w / 2 - (box.baseline - (box.height - box.baseline)) / 2;
        const unit = (sizePx * resolved.face.sizeAdjust) / resolved.metrics.unitsPerEm;
        const typo = resolved.metrics.os2?.useTypoMetrics ? resolved.metrics.os2 : null;
        const ascent = Math.round((typo ? typo.typoAscender : resolved.metrics.hhea.ascender) * unit);
        const descent = Math.round(-(typo ? typo.typoDescender : resolved.metrics.hhea.descender) * unit);
        m = [scale, 0, skew, -scale, alphabetic + (ascent - descent) / 2 - (path.advance * scale) / 2, top + path.vOrigin * scale];
      } else {
        // Sideways: turned a quarter clockwise, its em box centred on the line.
        const centre = left + w / 2;
        m = [0, scale, scale, skew, centre - (box.baseline - (box.height - box.baseline)) / 2, top];
      }
      glyphs.push({
        path,
        id: `${resolved.glyph}@${url}`,
        m,
        scale,
        bold: resolved.syntheticBold && !own.getPropertyValue("font-synthesis-weight").includes("none") ? sizePx * ratio : 0,
        fill: fillOf(parent, own),
      });
    }
    if (unsupported) break;
  }
  range.detach?.();
  if (unsupported || (!glyphs.length && !fallbacks.length)) {
    clear(element);
    if (waiting || unseen) tracked.set(element, { drawn: "", shape: shapeOf(element), waiting, positioned });
    return false;
  }
  const widthPx = widthPt * PT;
  const filled = installedDrawn && fillable;
  const svg = pictureOf(glyphs, fallbacks, color, widthPx, filled, stroke, clips);
  const shape = shapeOf(element);
  const drawn = `${svg.box.join(",")}|${svg.markup.length}|${svg.markup}`;
  const state = tracked.get(element);
  tracked.set(element, { drawn, shape, waiting, positioned });
  if (state?.drawn === drawn && element.style.getPropertyValue(OFFICE_TEXT_OUTLINE_IMAGE)) return filled;
  try {
    // Only what a quoted data URL must escape: the picture is read again at
    // each change, and percent-encoding every bracket tripled its size.
    element.style.setProperty(OFFICE_TEXT_OUTLINE_IMAGE, `url("data:image/svg+xml,${svg.markup.replace(/[%#"\\\n]/g, (ch) => encodeURIComponent(ch))}")`);
    element.style.setProperty(BOX[0], `${n2(svg.box[0])}px`);
    element.style.setProperty(BOX[1], `${n2(svg.box[1])}px`);
    element.style.setProperty(BOX[2], `${n2(svg.box[2])}px`);
    element.style.setProperty(BOX[3], `${n2(svg.box[3])}px`);
  } finally {
    wrote(element);
  }
  return filled;
}

/** A fill clipped to the words: the element's background (a gradient or a
 * pattern), its boxes (one per line it lies on, laid end to end as CSS lays a
 * broken inline's background: `offset` along `width`) in the picture's px. */
export interface ClipFill {
  readonly background: string;
  readonly width: number;
  readonly boxes: readonly { readonly x: number; readonly y: number; readonly width: number; readonly height: number; readonly offset: number }[];
}

const CLIP_FILL = "clip:";
const clipOf = (fill: string): number | null => (fill.startsWith(CLIP_FILL) ? Number(fill.slice(CLIP_FILL.length)) : null);

/** Whether an element's background is clipped to its words. */
function clipsToWords(style: CSSStyleDeclaration): boolean {
  return (style.backgroundClip === "text" || style.getPropertyValue("-webkit-background-clip") === "text")
    && style.backgroundImage !== "none" && style.backgroundImage !== "";
}

/** The background `clipped` lays under its words, in the element's own px
 * (`toLocal`; `L` its linear transform, axis aligned): each line's box of its
 * background positioning area. A textarea's is its mirror's padding box. */
function clipFillOf(
  clipped: Element, mirror: HTMLElement | null, toLocal: (x: number, y: number) => [number, number], L: readonly number[],
): ClipFill | null {
  const style = getComputedStyle(clipped);
  const layers = [
    `background-image:${style.backgroundImage}`,
    `background-size:${style.backgroundSize}`,
    `background-position:${style.backgroundPosition}`,
    `background-repeat:${style.backgroundRepeat}`,
    "background-color:transparent",
  ].join(";");
  if (mirror) {
    // The mirror's padding box is the textarea's (its borders copied).
    const origin = style.backgroundOrigin;
    const inset = (side: "Left" | "Top" | "Right" | "Bottom") =>
      (origin === "border-box" ? -(Number.parseFloat(style.getPropertyValue(`border-${side.toLowerCase()}-width`)) || 0) : 0)
      + (origin === "content-box" ? Number.parseFloat(style.getPropertyValue(`padding-${side.toLowerCase()}`)) || 0 : 0);
    const field = clipped as HTMLElement;
    const width = field.clientWidth - inset("Left") - inset("Right");
    const height = field.clientHeight - inset("Top") - inset("Bottom");
    if (!(width > 0 && height > 0)) return null;
    // The picture's origin is the mirror's padding box: the textarea's.
    return { background: layers, width, boxes: [{ x: inset("Left"), y: inset("Top"), width, height, offset: 0 }] };
  }
  const rects = clipped.getClientRects();
  if (!rects.length) return null;
  const boxes: { x: number; y: number; width: number; height: number; offset: number }[] = [];
  let offset = 0;
  for (const rect of rects) {
    const [x, y] = toLocal(rect.left, rect.top);
    const width = rect.width / L[0], height = rect.height / L[3];
    if (!(width > 0 && height > 0)) continue;
    boxes.push({ x, y, width, height, offset });
    offset += width;
  }
  return boxes.length ? { background: layers, width: offset, boxes } : null;
}

/** One picture of the band: each glyph's outline placed per character,
 * stroked `widthPx` wide, centred and round. */
export function pictureOf(
  glyphs: readonly Glyph[], fallbacks: readonly Fallback[], color: string, widthPx: number, filled = false,
  stroke: OfficeTextOutlineStroke = officeTextOutlineStroke(null),
  clips: readonly ClipFill[] = [],
): { markup: string; box: [number, number, number, number] } {
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  for (const glyph of glyphs) {
    const bbox = glyph.path.bbox;
    if (!bbox) continue;
    // A miter joins as far as its limit reaches past the corner.
    const reach = (widthPx / 2 + glyph.bold) * (stroke.join === "miter" ? Math.max(1, stroke.miterLimit) : 1) + 1;
    const [a, b, c, d, e, f] = glyph.m;
    for (const u of [bbox[0], bbox[2]]) {
      for (const v of [bbox[1], bbox[3]]) {
        const x = a * u + c * v + e, y = b * u + d * v + f;
        x0 = Math.min(x0, x - reach);
        x1 = Math.max(x1, x + reach);
        y0 = Math.min(y0, y - reach);
        y1 = Math.max(y1, y + reach);
      }
    }
  }
  for (const fallback of fallbacks) {
    const size = Number.parseFloat(fallback.font.split(" ").find((part) => part.endsWith("px")) ?? "16") || 16;
    x0 = Math.min(x0, fallback.x - size);
    x1 = Math.max(x1, fallback.x + size * 2);
    y0 = Math.min(y0, fallback.y - size * 1.2);
    y1 = Math.max(y1, fallback.y + size * 0.5);
  }
  if (!(x1 > x0 && y1 > y0)) return { markup: "", box: [0, 0, 0, 0] };
  x0 = Math.floor(x0);
  y0 = Math.floor(y0);
  x1 = Math.ceil(x1);
  y1 = Math.ceil(y1);
  const width = x1 - x0, height = y1 - y0;
  // Each glyph's outline written where it stands (no <use> of a shared
  // definition: the picture's own document clones every <use>, which cost
  // more to lay out than the repeated path data does to read).
  let uses = "";
  for (const glyph of glyphs) {
    const [a, b, c, d, e, f] = glyph.m;
    const unit = (widthPx + 2 * glyph.bold) / glyph.scale;
    const dash = stroke.dash ? ` stroke-dasharray="${stroke.dash.map((length) => p6(length * unit)).join(" ")}"` : "";
    uses += `<path d="${glyph.path.d}" transform="matrix(${p6(a)} ${p6(b)} ${p6(c)} ${p6(d)} ${n2(e - x0)} ${n2(f - y0)})" stroke-width="${p6(unit)}"${dash}/>`;
  }
  let texts = "";
  for (const fallback of fallbacks) {
    const dash = stroke.dash ? ` stroke-dasharray="${stroke.dash.map((length) => p6(length * widthPx)).join(" ")}"` : "";
    texts += `<text x="${n2(fallback.x - x0)}" y="${n2(fallback.y - y0)}" style="font:${svgEscape(fallback.font)}" stroke-width="${n2(widthPx)}"${dash}>${svgEscape(fallback.text)}</text>`;
  }
  // The words' own fill, from the same outlines, under the band (where an
  // installed font would fill them otherwise: drawsServedFile).
  let fill = "";
  if (filled) {
    // A gradient or pattern clipped to the words: its element's background,
    // laid where the element lays it, clipped to these outlines.
    const clipped = clips.map(() => "");
    for (const glyph of glyphs) {
      if (!glyph.fill || !glyph.path.d) continue;
      const [a, b, c, d, e, f] = glyph.m;
      const transform = `matrix(${p6(a)} ${p6(b)} ${p6(c)} ${p6(d)} ${n2(e - x0)} ${n2(f - y0)})`;
      const clip = clipOf(glyph.fill);
      if (clip !== null) {
        if (clip < clipped.length) clipped[clip] += `<path d="${glyph.path.d}" transform="${transform}"/>`;
        continue;
      }
      const bold = glyph.bold ? ` stroke="${svgEscape(glyph.fill)}" stroke-width="${p6((2 * glyph.bold) / glyph.scale)}"` : "";
      fill += `<path d="${glyph.path.d}" transform="${transform}" fill="${svgEscape(glyph.fill)}"${bold}/>`;
    }
    for (const fallback of fallbacks) {
      if (!fallback.fill) continue;
      const text = `<text x="${n2(fallback.x - x0)}" y="${n2(fallback.y - y0)}" style="font:${svgEscape(fallback.font)}"`;
      const clip = clipOf(fallback.fill);
      if (clip !== null) {
        if (clip < clipped.length) clipped[clip] += `${text}>${svgEscape(fallback.text)}</text>`;
        continue;
      }
      fill += `${text} fill="${svgEscape(fallback.fill)}">${svgEscape(fallback.text)}</text>`;
    }
    clips.forEach((clip, index) => {
      if (!clipped[index]) return;
      const boxes = clip.boxes.map((box) =>
        `<foreignObject x="${n2(box.x - x0)}" y="${n2(box.y - y0)}" width="${n2(box.width)}" height="${n2(box.height)}">` +
        `<div xmlns="http://www.w3.org/1999/xhtml" style="${svgEscape(`width:${n2(clip.width)}px;height:${n2(box.height)}px;margin-left:${n2(-box.offset)}px;${clip.background}`)}"/></foreignObject>`,
      ).join("");
      fill += `<clipPath id="w${index}" clipPathUnits="userSpaceOnUse">${clipped[index]}</clipPath><g clip-path="url(#w${index})">${boxes}</g>`;
    });
  }
  // A see-through outline (#rrggbbaa: Word's outline transparency) is the
  // whole band at that opacity, laid once: where neighbours' strokes meet
  // they do not darken, as Office strokes the words' outline as one path.
  const alpha = color.length === 9 ? parseInt(color.slice(7), 16) / 255 : 1;
  const opacity = alpha < 1 ? ` opacity="${p6(alpha)}"` : "";
  const markup =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" overflow="visible">` +
    (fill ? `<g stroke-linejoin="round">${fill}</g>` : "") +
    `<g fill="none" stroke="${color.slice(0, 7)}" stroke-linejoin="${stroke.join}"` +
    `${stroke.join === "miter" ? ` stroke-miterlimit="${p6(stroke.miterLimit)}"` : ""} stroke-linecap="${stroke.dash ? stroke.cap : "round"}"${opacity}>${uses}${texts}</g></svg>`;
  return { markup, box: [x0, y0, width, height] };
}

/** A title picture's glyphs laid out by the text engine (its advances and
 * Office's pair kerning) from the faces the page serves: null unless every
 * glyph's outline is kept already (the missing ones are asked for). */
function svgTitleGlyphs(
  text: string, family: string, weight: number, italic: boolean, sizePx: number,
): { glyphs: OfficeTextOutlineSvgGlyph[]; width: number } | null {
  if (SHAPED.test(text)) return null;
  const measured = measureText(text, { stack: family, weight, italic, sizePx, kerning: true });
  if (measured.status !== "ok") return null;
  const families = splitFamilyList(family);
  const glyphs: OfficeTextOutlineSvgGlyph[] = [];
  let pen = 0;
  let complete = true;
  for (const grapheme of measured.graphemes) {
    const codePoint = grapheme.text.codePointAt(0)!;
    if (grapheme.text.trim()) {
      const resolved = resolveFace(families, weight, italic, codePoint === 0xa0 ? 0x20 : codePoint, true, false);
      if (typeof resolved === "string" || !resolved.face.url) return null;
      const path = glyphPath(resolved.face.url, resolved.glyph);
      if (path === "pending") complete = false;
      else if (!path) return null;
      else if (path.d) {
        glyphs.push({
          key: `${resolved.glyph}@${resolved.face.url}`,
          d: path.d,
          x: pen,
          scale: (sizePx * resolved.face.sizeAdjust) / path.unitsPerEm,
          skew: italic && resolved.face.style === "normal" ? SYNTHETIC_SKEW : 0,
        });
      }
    }
    pen += grapheme.advance;
  }
  return complete ? { glyphs, width: measured.width } : null;
}

setOfficeTextOutlineSvgGlyphs(svgTitleGlyphs);

const canvasFonts = new Map<string, { family: string; weight: number; italic: boolean; sizePx: number } | null>();
const canvasPaths = new Map<string, Path2D>();

/** A canvas font string read as CSS reads it. */
function canvasFont(font: string): { family: string; weight: number; italic: boolean; sizePx: number } | null {
  if (canvasFonts.has(font)) return canvasFonts.get(font)!;
  let read: { family: string; weight: number; italic: boolean; sizePx: number } | null = null;
  if (typeof document !== "undefined") {
    const probe = document.createElement("span").style;
    probe.font = font;
    const sizePx = Number.parseFloat(probe.fontSize);
    if (probe.fontFamily && sizePx > 0 && probe.fontSize.endsWith("px")) {
      const weight = probe.fontWeight === "bold" ? 700 : probe.fontWeight === "normal" || !probe.fontWeight ? 400 : Number.parseInt(probe.fontWeight, 10) || 400;
      read = { family: probe.fontFamily, weight, italic: probe.fontStyle === "italic" || probe.fontStyle === "oblique", sizePx };
    }
  }
  if (canvasFonts.size > 256) canvasFonts.clear();
  canvasFonts.set(font, read);
  return read;
}

interface CanvasGlyph { path: Path2D; pen: number; scale: number; skew: number }

/** A canvas run's glyphs from the kept outlines, each at its pen (the
 * context's own measure of the text before it), and whether an installed
 * font draws any of them; null when they are not all to hand. */
function canvasGlyphs(context: CanvasRenderingContext2D, text: string): { placed: CanvasGlyph[]; installed: boolean } | null {
  if (typeof Path2D === "undefined" || SHAPED.test(text)) return null;
  if (!["start", "left"].includes(context.textAlign) || context.textBaseline !== "alphabetic" || context.direction === "rtl") return null;
  const font = canvasFont(context.font);
  if (!font) return null;
  const families = splitFamilyList(font.family);
  const placed: CanvasGlyph[] = [];
  let complete = true;
  let installed = false;
  let prefix = "";
  for (const { segment } of graphemes(text)) {
    const pen = prefix ? context.measureText(prefix).width : 0;
    prefix += segment;
    if (!segment.trim()) continue;
    const codePoint = segment.codePointAt(0)!;
    const resolved = resolveFace(families, font.weight, font.italic, codePoint === 0xa0 ? 0x20 : codePoint, true, false);
    if (typeof resolved === "string" || !resolved.face.url) return null;
    const served = drawsServedFile(resolved);
    if (served === "pending") return null;
    if (!served) installed = true;
    const path = glyphPath(resolved.face.url, resolved.glyph);
    if (path === "pending") {
      complete = false;
      continue;
    }
    if (!path) return null;
    if (!path.d) continue;
    const key = `${resolved.glyph}@${resolved.face.url}`;
    let shape = canvasPaths.get(key);
    if (!shape) {
      if (canvasPaths.size > 4096) canvasPaths.clear();
      shape = new Path2D(path.d);
      canvasPaths.set(key, shape);
    }
    placed.push({
      path: shape,
      pen,
      scale: (font.sizePx * resolved.face.sizeAdjust) / path.unitsPerEm,
      skew: font.italic && resolved.face.style === "normal" ? SYNTHETIC_SKEW : 0,
    });
  }
  return complete ? { placed, installed } : null;
}

/** A canvas run's band from the kept outlines (`officeTextOutlineCanvas`),
 * stroked centred and round in the context's stroke style. */
function canvasBand(context: CanvasRenderingContext2D, text: string, x: number, baseline: number, width: number): boolean {
  const glyphs = canvasGlyphs(context, text);
  if (!glyphs) return false;
  // The context's own join, cap and dash (the caller's: the outline's line,
  // round and solid where it states none), the dash in the glyph's units as
  // its width is.
  const dash = context.getLineDash();
  for (const glyph of glyphs.placed) {
    context.save();
    context.transform(glyph.scale, 0, glyph.skew * glyph.scale, -glyph.scale, x + glyph.pen, baseline);
    context.lineWidth = width / glyph.scale;
    if (dash.length) context.setLineDash(dash.map((length) => length / glyph.scale));
    context.stroke(glyph.path);
    context.restore();
  }
  return true;
}

/** An outlined canvas run's fill from the band's own outlines, where an
 * installed font draws the words (`officeTextOutlineCanvasFill`). */
function canvasFill(context: CanvasRenderingContext2D, text: string, x: number, baseline: number): boolean {
  const glyphs = canvasGlyphs(context, text);
  if (!glyphs?.installed) return false;
  for (const glyph of glyphs.placed) {
    context.save();
    context.transform(glyph.scale, 0, glyph.skew * glyph.scale, -glyph.scale, x + glyph.pen, baseline);
    context.fill(glyph.path);
    context.restore();
  }
  return true;
}

setOfficeTextOutlineCanvasBand(canvasBand, canvasFill);

setOfficeTextOutlineGlyphLookup((text, family, weight, italic, sizePx) => {
  if (!text.trim() || SHAPED.test(text)) return null;
  const codePoint = text.codePointAt(0)!;
  const resolved = resolveFace(splitFamilyList(family), weight, italic, codePoint === 0xa0 ? 0x20 : codePoint, true, false);
  if (resolved === "pending") return "pending";
  if (typeof resolved === "string" || !resolved.face.url || resolved.face.wght !== undefined || resolved.metrics.variationTables) return null;
  const served = drawsServedFile(resolved);
  if (served === "pending") return "pending";
  const path = glyphPath(resolved.face.url, resolved.glyph);
  if (path === "pending") return "pending";
  if (!path) return null;
  return {
    d: path.d,
    scale: (sizePx * resolved.face.sizeAdjust) / path.unitsPerEm,
    skew: italic && resolved.face.style === "normal" ? SYNTHETIC_SKEW : 0,
    ...(served ? {} : { installed: true }),
  };
});
