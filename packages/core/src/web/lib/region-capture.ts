/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Context } from "modern-screenshot";





export interface CircledSquare {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** The pointer's path as it circled, in viewport pixels. */
export type CircleInk = ReadonlyArray<{ x: number; y: number }>;

export interface CapturedRegion {
  /** What the box holds, for the model. */
  text: string;
  width: number;
  height: number;
  /** The answer (or message) the box was mostly taken from. */
  sourceMessageId: string | null;
}

/** The box around a circle drawn with the pointer, a little room around the
 * ink, at least a readable size, kept on screen. */
export function squareAround(
  points: CircleInk,
  viewport: { width: number; height: number },
  minimum = 48,
  room = 6,
): CircledSquare | null {
  if (points.length < 3) return null;
  const xs = points.map((point) => point.x);
  const ys = points.map((point) => point.y);
  const inkWidth = Math.max(...xs) - Math.min(...xs);
  const inkHeight = Math.max(...ys) - Math.min(...ys);
  if (inkWidth < 24 && inkHeight < 24) return null;
  const span = (low: number, high: number, extent: number) => {
    const size = Math.min(Math.max(high - low + 2 * room, minimum), extent);
    const start = Math.min(Math.max((low + high) / 2 - size / 2, 0), extent - size);
    return [start, size] as const;
  };
  const [left, width] = span(Math.min(...xs), Math.max(...xs), viewport.width);
  const [top, height] = span(Math.min(...ys), Math.max(...ys), viewport.height);
  return { left, top, width, height };
}

interface Box {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

function intersects(rect: DOMRect | Box, square: CircledSquare): boolean {
  const width = rect.right - rect.left;
  const height = rect.bottom - rect.top;
  return (
    rect.right > square.left &&
    rect.left < square.left + square.width &&
    rect.bottom > square.top &&
    rect.top < square.top + square.height &&
    width > 0 &&
    height > 0
  );
}

/** Never part of what was circled: the circling layer itself (still mounted
 * while the box is read), the dev tools' badge, and what draws nothing. */
const NEVER_DRAWN =
  ".region-ask-overlay, nextjs-portal, script, style, link, meta, template, noscript, head, title";




const TURN_SELECTOR = "li.dash-chat-msg, .ai-panel-messages .ai-panel-msg";

/** The turns inside `stream` the square crosses; a turn drawn inside another
 * (a message shown within a run's card) goes with the one around it. */
function circledTurns(stream: HTMLElement, square: CircledSquare): HTMLElement[] {
  return Array.from(stream.querySelectorAll<HTMLElement>(TURN_SELECTOR)).filter((turn) => {
    const outer = turn.parentElement?.closest(TURN_SELECTOR);
    if (outer && stream.contains(outer)) return false;
    return intersects(turn.getBoundingClientRect(), square);
  });
}

/**
 * The element whose drawing holds everything the screen shows inside the
 * square: from the page down, as long as one child alone covers the square
 * and nothing else (a floating plan, a window, a toast) shows over it. Never
 * below a transformed box: its children are measured on screen after the
 * transform, and drawn without it.
 */
export function regionRoot(square: CircledSquare): HTMLElement {
  let root: HTMLElement = document.body;
  for (let depth = 0; depth < 48; depth += 1) {
    let only: HTMLElement | null = null;
    let crowded = false;
    for (const child of Array.from(root.children)) {
      if (child.matches(NEVER_DRAWN)) continue;
      const rect = child.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0) {
        // A holder of no size (a portal's root, display: contents) may still
        // show its children anywhere.
        if (child.childElementCount && getComputedStyle(child).display !== "none") crowded = true;
      } else if (intersects(rect, square)) {
        if (only || !(child instanceof HTMLElement)) crowded = true;
        else only = child;
      }
      if (crowded) break;
    }
    if (crowded || !only) break;
    const rect = only.getBoundingClientRect();
    const covers =
      rect.left <= square.left + 0.5 &&
      rect.top <= square.top + 0.5 &&
      rect.right >= square.left + square.width - 0.5 &&
      rect.bottom >= square.top + square.height - 0.5;
    if (!covers) break;
    const style = getComputedStyle(only);
    if (
      style.transform !== "none" ||
      style.zoom !== "1" ||
      Math.abs(rect.width - only.offsetWidth) > 0.5 ||
      Math.abs(rect.height - only.offsetHeight) > 0.5
    ) {
      break;
    }
    root = only;
  }
  return root;
}

/** The words inside the square: formulas as TeX, pictures by caption. */
export function circledText(roots: Element[], square: CircledSquare): string {
  const lines: string[] = [];
  let line = "";
  const flush = () => {
    const text = line.replace(/[ \t]+/g, " ").trim();
    if (text) lines.push(text);
    line = "";
  };
  const range = document.createRange();
  const visit = (node: Node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const value = node.textContent || "";
      if (!value.trim()) return;
      range.selectNodeContents(node);
      if (Array.from(range.getClientRects()).some((rect) => intersects(rect as DOMRect, square))) {
        line += value;
      }
      return;
    }
    if (!(node instanceof Element)) return;
    if (node.matches(NEVER_DRAWN)) return;
    if (node.matches("button, [role=toolbar], .dash-chat-user-actions, textarea, input, select")) return;
    // A box wholly outside the square holds nothing circled: a long answer's
    
    const box = node.getBoundingClientRect();
    if (box.width > 0 && box.height > 0 && !intersects(box, square)) return;
    if (node.classList.contains("katex")) {
      if (intersects(node.getBoundingClientRect(), square)) {
        const tex = node.querySelector("annotation[encoding='application/x-tex']")?.textContent;
        const display = node.parentElement?.classList.contains("katex-display");
        if (tex) line += display ? `\n$$${tex}$$\n` : `$${tex}$`;
      }
      return;
    }
    if (node instanceof HTMLImageElement) {
      if (intersects(node.getBoundingClientRect(), square) && node.alt) line += `［图：${node.alt}］`;
      return;
    }
    const block = /^(P|LI|H[1-6]|TR|BLOCKQUOTE|PRE|DIV|FIGCAPTION|text)$/i.test(node.tagName);
    if (block) flush();
    for (const child of Array.from(node.childNodes)) visit(child);
    if (node.tagName === "TD" || node.tagName === "TH") line += " | ";
    if (block) flush();
  };
  for (const root of roots) visit(root);
  flush();
  return lines.join("\n").replace(/\n{3,}/g, "\n\n").slice(0, 20_000);
}

/**
 * What the square shows, as words for the model: everything on screen inside
 * it (a run's card, its plan, a made file's card), not only the
 * conversation's turns; and the answer it was mostly taken from.
 */
export function captureRegion(stream: HTMLElement, square: CircledSquare): CapturedRegion | null {
  const text = circledText([regionRoot(square)], square);
  const sourceMessageId =
    circledTurns(stream, square)
      .map((turn) => turn.querySelector<HTMLElement>("[data-message-id]")?.dataset.messageId ?? null)
      .find((id) => id && /^[0-9a-f-]{36}$/i.test(id)) ?? null;
  return { text, width: square.width, height: square.height, sourceMessageId };
}

/** The first message of a side ask: what was circled, then the question
 * (the server frames it so: routers/chat.py ``_with_side_ask_region``). */
export const SIDE_ASK_REGION_MARK = "【我圈选的内容】";
export const SIDE_ASK_QUESTION_MARK = "【我的问题】";

export function sideAskFirstMessage(region: string, question: string): string {
  return `${SIDE_ASK_REGION_MARK}\n${region}\n\n${SIDE_ASK_QUESTION_MARK}\n${question}`;
}

/** What the reader asked, without the circled words the message carries. */
export function sideAskQuestionOf(content: string): string {
  if (!content.startsWith(SIDE_ASK_REGION_MARK)) return content;
  const at = content.indexOf(SIDE_ASK_QUESTION_MARK);
  return at >= 0 ? content.slice(at + SIDE_ASK_QUESTION_MARK.length).trim() : content;
}

/** A side ask's title: the first line of what was circled that has words in
 * it, without formulas, table rules and Markdown marks (a circled table was
 * titled "| 目标天体 | 发射 | | |"). */
export function sideAskTitle(text: string, limit = 24): string {
  for (const raw of text.replace(/\$+[^$]*\$+/g, "").split("\n")) {
    const line = raw
      .replace(/[|｜]/g, " ")
      .replace(/^[\s#>*\-+•·]+/, "")
      .replace(/[*_`~]+/g, "")
      .replace(/\s+/g, " ")
      .trim();
    if (/[\p{L}\p{N}]{2,}/u.test(line)) return line.slice(0, limit);
  }
  return "";
}

// ---------------------------------------------------------------------------
// The picture
// ---------------------------------------------------------------------------

const isOpaqueColor = (color: string) =>
  Boolean(color) && color !== "transparent" && !/rgba?\([^)]*,\s*0\)$/.test(color);

/** The background the square is drawn on: the first opaque one behind the
 * drawn root (the page's own, when the root has none). */
function backdropOf(element: Element): string {
  for (let at: Element | null = element; at; at = at.parentElement) {
    const color = getComputedStyle(at).backgroundColor;
    if (isOpaqueColor(color)) return color;
  }
  return "#ffffff";
}

/** Something the page shows that a drawing of its markup cannot: a frame of
 * another origin, a canvas holding another site's pixels, a video. */
interface Unreadable {
  element: HTMLIFrameElement | HTMLCanvasElement | HTMLVideoElement;
  /** Its visible part, cut by the boxes that clip it, in viewport pixels. */
  shown: Box;
  label: string;
}

/** What is drawn of the square, decided at once when it is circled (the
 * page may change while the picture is drawn; a window opened meanwhile is
 * no part of it). */
interface RegionPlan {
  root: HTMLElement;
  rootRect: DOMRect;
  /** Elements drawn: their box at least. */
  drawn: Set<Element>;
  /** Of those, the ones drawn with their contents. */
  open: Set<Element>;
  /** The characters drawn, by the font-family list they are set in. */
  characters: Map<string, Set<number>>;
  /** The pictures drawn, by address: their size on screen and the element. */
  pictures: Map<string, { width: number; height: number; image: HTMLImageElement }>;
  unreadable: Unreadable[];
  /** Frames of this page's own origin, drawn on their own. */
  frames: Array<{ frame: HTMLIFrameElement; shown: Box }>;
}

/** Displays whose box keeps its size with its contents gone (a drawn element
 * carries its computed width and height); inline runs and table rows are
 * entered instead, their cells emptied. */
const ENTERED_OUTSIDE = /^(inline|contents|table-row|table-row-group|table-header-group|table-footer-group|table-column-group|table-column|ruby|ruby-text|ruby-base)$/;

/** The part of `element` the boxes around it (up to `root`) let show. */
function shownPart(element: Element, root: Element, square: CircledSquare): Box | null {
  const rect = element.getBoundingClientRect();
  let box: Box = { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom };
  for (let at = element.parentElement; at && at !== root.parentElement; at = at.parentElement) {
    const style = getComputedStyle(at);
    if (style.overflowX === "visible" && style.overflowY === "visible") continue;
    const clip = at.getBoundingClientRect();
    box = {
      left: Math.max(box.left, clip.left),
      top: Math.max(box.top, clip.top),
      right: Math.min(box.right, clip.right),
      bottom: Math.min(box.bottom, clip.bottom),
    };
  }
  box = {
    left: Math.max(box.left, square.left),
    top: Math.max(box.top, square.top),
    right: Math.min(box.right, square.left + square.width),
    bottom: Math.min(box.bottom, square.top + square.height),
  };
  return box.right > box.left && box.bottom > box.top ? box : null;
}

function canvasReadable(canvas: HTMLCanvasElement | HTMLVideoElement): boolean {
  try {
    const probe = document.createElement("canvas");
    probe.width = 1;
    probe.height = 1;
    const context = probe.getContext("2d");
    if (!context) return false;
    context.drawImage(canvas, 0, 0, 1, 1);
    context.getImageData(0, 0, 1, 1);
    return true;
  } catch {
    return false;
  }
}

/**
 * Decides what the picture of `square` draws: every element the screen shows
 * inside it, whatever it belongs to; the blocks beside it as empty boxes of
 * their size (so what is inside keeps its place), their insides never
 * cloned; nothing hidden. Only the path down through the square is measured.
 */
function planRegion(square: CircledSquare): RegionPlan {
  const root = regionRoot(square);
  const plan: RegionPlan = {
    root,
    rootRect: root.getBoundingClientRect(),
    drawn: new Set([root]),
    open: new Set([root]),
    characters: new Map(),
    pictures: new Map(),
    unreadable: [],
    frames: [],
  };
  const addCharacters = (families: string, text: string) => {
    if (!text) return;
    let set = plan.characters.get(families);
    if (!set) plan.characters.set(families, (set = new Set()));
    for (const character of text) set.add(character.codePointAt(0)!);
  };
  const inspect = (element: Element) => {
    if (element instanceof HTMLImageElement) {
      const source = element.currentSrc || element.src;
      const rect = element.getBoundingClientRect();
      if (source) plan.pictures.set(source, { width: rect.width, height: rect.height, image: element });
    } else if (element instanceof HTMLIFrameElement) {
      const shown = shownPart(element, root, square);
      if (!shown) return;
      let document_: Document | null = null;
      try {
        document_ = element.contentDocument;
      } catch {
        document_ = null;
      }
      if (document_?.documentElement) plan.frames.push({ frame: element, shown });
      else plan.unreadable.push({ element, shown, label: element.title || element.getAttribute("aria-label") || "" });
    } else if (element instanceof HTMLCanvasElement || element instanceof HTMLVideoElement) {
      if (canvasReadable(element) && element instanceof HTMLCanvasElement) return;
      const shown = shownPart(element, root, square);
      if (shown) plan.unreadable.push({ element, shown, label: element.getAttribute("aria-label") || element.title || "" });
    } else if (element instanceof HTMLInputElement || element instanceof HTMLTextAreaElement) {
      addCharacters(getComputedStyle(element).fontFamily, element.value || element.placeholder);
    }
  };
  const ownText = (element: Element) => {
    let text = "";
    for (const node of Array.from(element.childNodes)) {
      if (node.nodeType === Node.TEXT_NODE) text += node.textContent ?? "";
    }
    if (text.trim()) addCharacters(getComputedStyle(element).fontFamily, text);
  };
  // An inline <svg> is drawn whole: its parts are few, and its own
  // definitions (gradients, a sprite's symbols) live in boxes of no size.
  const drawWhole = (element: Element) => {
    plan.drawn.add(element);
    plan.open.add(element);
    ownText(element);
    for (const child of Array.from(element.children)) drawWhole(child);
  };
  const walk = (element: Element) => {
    ownText(element);
    for (const child of Array.from(element.children)) {
      if (child.matches(NEVER_DRAWN)) continue;
      const rect = child.getBoundingClientRect();
      const sized = rect.width > 0 && rect.height > 0;
      const inside = sized && intersects(rect, square);
      if (inside && !(child instanceof SVGSVGElement)) {
        plan.drawn.add(child);
        plan.open.add(child);
        inspect(child);
        walk(child);
        continue;
      }
      const display = getComputedStyle(child).display;
      if (display === "none") continue;
      if (child instanceof SVGSVGElement && (inside || !sized)) {
        drawWhole(child);
        continue;
      }
      plan.drawn.add(child);
      if (!sized || ENTERED_OUTSIDE.test(display)) {
        // No box of its own, or one that sizes to what it holds: its
        // children are what show (or keep the room).
        plan.open.add(child);
        walk(child);
      }
    }
  };
  walk(root);
  return plan;
}

// --- Faces ---------------------------------------------------------------

interface FaceRule {
  family: string;
  /** Code point ranges the face covers; ``null``: all. */
  ranges: Array<[number, number]> | null;
  /** The face's identity as document.fonts names it. */
  key: string;
  cssText: string;
  /** Its files, absolute; none for a face of the system's (local()). */
  urls: Array<{ written: string; absolute: string }>;
}

const unquoted = (value: string) => value.replace(/["']/g, "").trim().toLowerCase();

function parseUnicodeRange(value: string): Array<[number, number]> | null {
  if (!value.trim()) return null;
  const ranges: Array<[number, number]> = [];
  for (const raw of value.split(",")) {
    const token = raw.trim().replace(/^u\+/i, "");
    if (!token) continue;
    if (token.includes("?")) {
      ranges.push([parseInt(token.replace(/\?/g, "0"), 16), parseInt(token.replace(/\?/g, "F"), 16)]);
    } else if (token.includes("-")) {
      const [low, high] = token.split("-");
      ranges.push([parseInt(low, 16), parseInt(high, 16)]);
    } else {
      const point = parseInt(token, 16);
      ranges.push([point, point]);
    }
  }
  return ranges.filter(([low, high]) => Number.isFinite(low) && Number.isFinite(high));
}

const weightOf = (value: string) =>
  (value || "400").replace(/\bnormal\b/g, "400").replace(/\bbold\b/g, "700").replace(/\s+/g, " ").trim();

function faceKey(family: string, style: string, weight: string, range: string): string {
  const ranges = parseUnicodeRange(range);
  const rangeKey = ranges ? ranges.map(([low, high]) => `${low}-${high}`).join(",") : "all";
  return [unquoted(family), (style || "normal").trim(), weightOf(weight), rangeKey === "0-1114111" ? "all" : rangeKey].join("|");
}

/** Each stylesheet's faces, read once (sheets that come later are read when
 * they come: KaTeX's arrives with the first formula). */
const sheetFaces = new WeakMap<CSSStyleSheet, FaceRule[]>();

function facesOfRules(rules: CSSRuleList, base: string, out: FaceRule[]): void {
  for (const rule of Array.from(rules)) {
    if (rule instanceof CSSFontFaceRule) {
      const style = rule.style;
      const family = style.getPropertyValue("font-family");
      const urls = Array.from(style.getPropertyValue("src").matchAll(/url\(\s*(["']?)([^"')]+)\1\s*\)/g)).map(
        (match) => ({ written: match[0], absolute: new URL(match[2], base).href }),
      );
      out.push({
        family: unquoted(family),
        ranges: parseUnicodeRange(style.getPropertyValue("unicode-range")),
        key: faceKey(
          family,
          style.getPropertyValue("font-style"),
          style.getPropertyValue("font-weight"),
          style.getPropertyValue("unicode-range"),
        ),
        cssText: rule.cssText,
        urls,
      });
    } else if (rule instanceof CSSImportRule) {
      const sheet = rule.styleSheet;
      if (!sheet) continue;
      try {
        facesOfRules(sheet.cssRules, sheet.href || base, out);
      } catch {
        // another origin's sheet: its faces are not readable
      }
    } else if ("cssRules" in rule && (rule as CSSGroupingRule).cssRules) {
      facesOfRules((rule as CSSGroupingRule).cssRules, base, out);
    }
  }
}

function pageFaces(): FaceRule[] {
  const faces: FaceRule[] = [];
  for (const sheet of Array.from(document.styleSheets)) {
    let known = sheetFaces.get(sheet);
    if (!known) {
      known = [];
      try {
        facesOfRules(sheet.cssRules, sheet.href || location.href, known);
      } catch {
        // another origin's sheet
      }
      sheetFaces.set(sheet, known);
    }
    faces.push(...known);
  }
  return faces;
}

/** A face file as a data URL, read back from the browser's cache once a page. */
const faceData = new Map<string, Promise<string | null>>();

function blobToDataUrl(blob: Blob): Promise<string> {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

function faceFile(url: string): Promise<string | null> {
  let data = faceData.get(url);
  if (!data) {
    data = fetch(url, { cache: "force-cache" })
      .then((response) => (response.ok ? response.blob() : Promise.reject(new Error(String(response.status)))))
      .then(blobToDataUrl)
      .catch(() => {
        faceData.delete(url);
        return null;
      });
    faceData.set(url, data);
  }
  return data;
}

const covers = (ranges: Array<[number, number]> | null, points: number[]) =>
  !ranges || points.some((point) => ranges.some(([low, high]) => point >= low && point <= high));











async function faceCss(plan: RegionPlan, deadline: number): Promise<string> {
  return faceCssForCharacters(plan.characters, deadline);
}

/**
 * `@font-face` rules, their files embedded, for exactly the faces these
 * characters are drawn in on screen (``characters``: by the font-family list
 * they are set in, as code points): the faces of those families the page has
 * loaded whose unicode-range holds one of them. A picture of the page carries
 * these and no others -- not every unicode-range file of a Chinese family,
 * nor the faces of text no longer drawn. ``deadline`` (a `performance.now()`
 * time): a face not read by then is left out; ``Infinity`` waits for every
 * one (a recording, whose frames must not depend on the network's speed).
 */
export async function faceCssForCharacters(
  characters: ReadonlyMap<string, ReadonlySet<number>>,
  deadline: number,
): Promise<string> {
  const loaded = new Set<string>();
  for (const face of document.fonts ? Array.from(document.fonts) : []) {
    if (face.status === "loaded") loaded.add(faceKey(face.family, face.style, face.weight, face.unicodeRange));
  }
  const faces = pageFaces();
  const byFamily = new Map<string, FaceRule[]>();
  for (const face of faces) {
    const list = byFamily.get(face.family);
    if (list) list.push(face);
    else byFamily.set(face.family, [face]);
  }
  const chosen = new Set<FaceRule>();
  for (const [families, drawn] of characters) {
    const points = Array.from(drawn);
    for (const family of families.split(",").map(unquoted)) {
      for (const face of byFamily.get(family) ?? []) {
        if (chosen.has(face)) continue;
        // A face of the system's (local(), a metric fallback) costs nothing.
        if (!face.urls.length || (loaded.has(face.key) && covers(face.ranges, points))) chosen.add(face);
      }
    }
  }
  // One rule per file: a variable face is declared once per weight with the
  // same file (Noto Sans SC: 300–600), and each copy embedded it again.
  // Merged into a rule over the weights' range, which the screen draws alike.
  const groups = new Map<string, FaceRule[]>();
  for (const face of chosen) {
    const [family, style, weight, range] = face.key.split("|");
    const single = /^\d+$/.test(weight);
    const id = single && face.urls.length ? [family, style, range, face.urls[0].absolute].join("|") : face.cssText;
    const group = groups.get(id);
    if (group) group.push(face);
    else groups.set(id, [face]);
  }
  // Late faces are left out; the ones read by then are kept.
  const ready: Array<string | undefined> = [];
  const all = Promise.all(
    Array.from(groups.values()).map(async (group, index) => {
      const face = group[0];
      let text = face.cssText;
      if (group.length > 1) {
        const weights = group.map((member) => Number(member.key.split("|")[2]));
        text = text.replace(/font-weight:\s*[^;]+;/, `font-weight: ${Math.min(...weights)} ${Math.max(...weights)};`);
      }
      // The first file the rule names is the one the browser took (woff2
      // before woff and ttf); the others are not embedded.
      const [first, ...rest] = face.urls;
      for (const url of rest) {
        text = text.split(url.written).join("").replace(/,\s*format\([^)]*\)/g, "");
      }
      text = text.replace(/,(\s*,)+/g, ",").replace(/,\s*;/g, ";");
      if (first) {
        const data = await faceFile(first.absolute);
        if (!data) return;
        text = text.split(first.written).join(`url("${data}")`);
      }
      ready[index] = text;
    }),
  );
  if (Number.isFinite(deadline)) {
    const left = Math.max(0, deadline - performance.now());
    await Promise.race([all, new Promise<void>((resolve) => setTimeout(resolve, left))]);
  } else {
    await all;
  }
  // In the page's own order, whatever order the files came in: the same
  // characters give the same text.
  return ready.filter((text): text is string => text !== undefined).join("\n");
}

/**
 * Started when circling starts, while the reader draws: the page's faces
 * are read, and the files of the faces the screen has used are read back
 * from the cache, so the picture rarely waits for one.
 */
export function preparePictureFonts(): Promise<void> {
  try {
    const loaded = new Set<string>();
    for (const face of document.fonts ? Array.from(document.fonts) : []) {
      if (face.status === "loaded") loaded.add(faceKey(face.family, face.style, face.weight, face.unicodeRange));
    }
    for (const face of pageFaces()) {
      if (face.urls.length && loaded.has(face.key)) void faceFile(face.urls[0].absolute);
    }
  } catch {
    // read again when drawing
  }
  return Promise.resolve();
}

// --- Pictures inside the square ----------------------------------------

const TRANSPARENT_PIXEL =
  "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";

/** What stands for a picture the page shows but no script may read (another
 * site's, sent without CORS): a picture frame of its size, not a hole. */
function standInPicture(width: number, height: number, dark: boolean): string {
  const w = Math.max(8, Math.round(width));
  const h = Math.max(8, Math.round(height));
  const size = Math.max(10, Math.min(36, Math.min(w, h) * 0.4));
  const x = (w - size) / 2;
  const y = (h - size * 0.8) / 2;
  const fill = dark ? "#2b2f36" : "#eef0f3";
  const ink = dark ? "#8b93a1" : "#9aa3b2";
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">` +
    `<rect width="${w}" height="${h}" rx="6" fill="${fill}"/>` +
    `<g fill="none" stroke="${ink}" stroke-width="${Math.max(1.2, size / 16)}" stroke-linejoin="round">` +
    `<rect x="${x}" y="${y}" width="${size}" height="${size * 0.8}" rx="${size / 10}"/>` +
    `<path d="M${x + size * 0.12} ${y + size * 0.68} L${x + size * 0.38} ${y + size * 0.4} L${x + size * 0.58} ${y + size * 0.6} L${x + size * 0.72} ${y + size * 0.48} L${x + size * 0.88} ${y + size * 0.66}"/>` +
    `<circle cx="${x + size * 0.7}" cy="${y + size * 0.24}" r="${size * 0.07}"/></g></svg>`;
  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
}

async function fetchPicture(url: string, credentials: RequestCredentials, signal: AbortSignal): Promise<string> {
  const response = await fetch(url, { credentials, cache: "force-cache", signal });
  if (!response.ok) throw new Error(String(response.status));
  const blob = await response.blob();
  if (!blob.size) throw new Error("empty");
  return blobToDataUrl(blob);
}

/** A picture as the drawing embeds it: read as this page reads it (with its
 * sign-in, for the API's own pictures), else as anyone may, else from the
 * pixels already on screen; a stand-in of its size when none of these can. */
function pictureReader(plan: RegionPlan, dark: boolean, budgetMs: number) {
  return async (url: string): Promise<string> => {
    if (url.startsWith("data:")) return url;
    const shown = plan.pictures.get(url);
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), budgetMs);
    try {
      for (const credentials of ["include", "omit"] as const) {
        try {
          return await fetchPicture(url, credentials, controller.signal);
        } catch {
          if (controller.signal.aborted) break;
        }
      }
    } finally {
      clearTimeout(timer);
    }
    if (shown?.image.complete && shown.image.naturalWidth) {
      try {
        const canvas = document.createElement("canvas");
        canvas.width = shown.image.naturalWidth;
        canvas.height = shown.image.naturalHeight;
        canvas.getContext("2d")?.drawImage(shown.image, 0, 0);
        return canvas.toDataURL("image/png");
      } catch {
        // another site's pixels
      }
    }
    return shown ? standInPicture(shown.width, shown.height, dark) : TRANSPARENT_PIXEL;
  };
}

function isDark(color: string): boolean {
  const match = /rgba?\(\s*([\d.]+)[,\s]+([\d.]+)[,\s]+([\d.]+)/.exec(color);
  if (!match) return false;
  const [r, g, b] = match.slice(1, 4).map(Number);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b < 110;
}

/** A frame, a canvas or a video the drawing cannot read: a panel of its
 * size with a window mark and its title, so the picture says something was
 * there instead of showing a hole. */
function paintStandIn(
  context: CanvasRenderingContext2D,
  box: { x: number; y: number; width: number; height: number },
  label: string,
  scale: number,
  dark: boolean,
  fontFamily: string,
): void {
  context.save();
  context.beginPath();
  context.rect(box.x, box.y, box.width, box.height);
  context.clip();
  context.fillStyle = dark ? "#262a31" : "#f1f3f6";
  context.fillRect(box.x, box.y, box.width, box.height);
  context.strokeStyle = dark ? "#4a505b" : "#cfd5de";
  context.lineWidth = scale;
  context.setLineDash([4 * scale, 3 * scale]);
  context.strokeRect(box.x + scale / 2, box.y + scale / 2, box.width - scale, box.height - scale);
  context.setLineDash([]);
  const icon = Math.min(28 * scale, box.width * 0.4, box.height * 0.4);
  const centreX = box.x + box.width / 2;
  const hasLabel = Boolean(label) && box.height > icon + 24 * scale;
  const iconTop = box.y + box.height / 2 - (hasLabel ? icon * 0.9 : icon / 2);
  context.strokeStyle = dark ? "#8b93a1" : "#9aa3b2";
  context.lineWidth = Math.max(1, icon / 14);
  context.beginPath();
  context.roundRect(centreX - icon / 2, iconTop, icon, icon * 0.8, icon / 8);
  context.moveTo(centreX - icon / 2, iconTop + icon * 0.24);
  context.lineTo(centreX + icon / 2, iconTop + icon * 0.24);
  context.stroke();
  if (hasLabel) {
    context.fillStyle = dark ? "#b4bac5" : "#6b7280";
    context.font = `${12 * scale}px ${fontFamily}`;
    context.textAlign = "center";
    context.textBaseline = "top";
    let text = label;
    const room = box.width - 16 * scale;
    while (text.length > 1 && context.measureText(text).width > room) text = text.slice(0, -1);
    if (text !== label) text = `${text.slice(0, -1)}…`;
    context.fillText(text, centreX, iconTop + icon * 0.8 + 8 * scale);
  }
  context.restore();
}

/** The circle the reader drew, faint, over the picture: what they meant
 * inside the box, for them and for the model, without dimming the rest (it
 * is still read; the words sent are the whole box's). */
function paintInk(
  context: CanvasRenderingContext2D,
  ink: CircleInk,
  square: CircledSquare,
  scale: number,
): void {
  if (ink.length < 3) return;
  const points = ink.map((point) => ({
    x: (point.x - square.left) * scale,
    y: (point.y - square.top) * scale,
  }));
  context.save();
  context.strokeStyle = "rgba(99, 102, 241, 0.5)";
  context.lineWidth = 2 * scale;
  context.lineCap = "round";
  context.lineJoin = "round";
  context.beginPath();
  context.moveTo(points[0].x, points[0].y);
  // Through the midpoints: the pointer's jitter smoothed, the shape kept.
  for (let index = 1; index < points.length - 1; index += 1) {
    const here = points[index];
    const next = points[index + 1];
    context.quadraticCurveTo(here.x, here.y, (here.x + next.x) / 2, (here.y + next.y) / 2);
  }
  const last = points[points.length - 1];
  context.lineTo(last.x, last.y);
  // Closed when the reader came most of the way back.
  const gap = Math.hypot(last.x - points[0].x, last.y - points[0].y);
  if (gap < Math.max(square.width, square.height) * scale * 0.35) context.closePath();
  context.stroke();
  context.restore();
}

interface DrawOptions {
  /** Embed the faces the screen used (else the system's faces). */
  faces: boolean;
  /** Read the pictures inside the square (else stand-ins). */
  pictures: boolean;
  deadline: number;
}

/** The box a frame takes on the page, as an empty block's style. */
function frameBoxStyle(frame: HTMLIFrameElement): string {
  const style = getComputedStyle(frame);
  const rect = frame.getBoundingClientRect();
  const keep = [
    "display", "position", "top", "right", "bottom", "left", "float", "margin-top", "margin-right",
    "margin-bottom", "margin-left", "vertical-align", "flex-grow", "flex-shrink", "flex-basis",
    "align-self", "justify-self", "order", "grid-area", "z-index", "transform", "border-radius",
  ];
  const declarations = keep
    .map((name) => [name, style.getPropertyValue(name)] as const)
    .filter(([, value]) => value)
    .map(([name, value]) => `${name}:${value}`);
  // Its size as laid out (a transform's scale aside), border included.
  declarations.push(
    `width:${frame.offsetWidth || rect.width}px`,
    `height:${frame.offsetHeight || rect.height}px`,
    "box-sizing:border-box",
    style.display === "inline" ? "display:inline-block" : "",
  );
  return declarations.filter(Boolean).join(";");
}

/** The drawing of the plan's root: a canvas of the root's size. */
async function drawRoot(plan: RegionPlan, scale: number, dark: boolean, options: DrawOptions): Promise<HTMLCanvasElement> {
  const { createContext, destroyContext, domToCanvas } = await import("modern-screenshot");
  // The faces may take most of the budget (a busy server answering a cold
  // cache); the drawing itself takes well under a second of what is left.
  const fontCss = options.faces ? await faceCss(plan, Math.min(options.deadline - 1_500, performance.now() + 4_500)) : "";
  const readPicture = pictureReader(plan, dark, 3_000);
  const { root, rootRect, drawn, open } = plan;
  // A frame stops Chrome painting everything after it in a drawn picture
  // (and a frame of this origin is drawn as its whole page, in its place):
  // each is drawn as an empty box of its own size and place, and what it
  // shows is put over the picture afterwards (drawPicture).
  const frameBoxes = Array.from(root.querySelectorAll("iframe"))
    .filter((frame) => drawn.has(frame))
    .map(frameBoxStyle);
  // The context is made on an empty holder and pointed at the root: made on
  // the root it would first wait for every picture under it, the ones a
  // lazy list never loads off screen included (3 s each).
  const holder = document.createElement("div");
  const context: Context = await createContext(holder, {
    scale,
    width: root.offsetWidth || rootRect.width,
    height: root.offsetHeight || rootRect.height,
    backgroundColor: null,
    // Drawn where it stands: the root's own margin is dropped by the
    // drawing, and a <body> root would take the browser's 8 px default.
    style: { margin: "0" },
    timeout: 3_000,
    font: fontCss ? { cssText: fontCss } : false,
    // A scrolled list is drawn where it was scrolled to.
    features: { restoreScrollPosition: true },
    fetchFn: options.pictures
      ? readPicture
      : async (url: string) => {
          const shown = plan.pictures.get(url);
          return url.startsWith("data:") ? url : shown ? standInPicture(shown.width, shown.height, dark) : TRANSPARENT_PIXEL;
        },
    filter: (node: Node) => {
      if (node instanceof Element) return drawn.has(node);
      const parent = node.parentElement;
      return Boolean(parent && open.has(parent));
    },
    onCloneNode: (clone: Node) => {
      if (!(clone instanceof Element)) return;
      let index = 0;
      for (const frame of Array.from(clone.querySelectorAll("iframe, html"))) {
        if (!clone.contains(frame)) continue;
        const box = document.createElement("div");
        box.setAttribute("style", frameBoxes[index] ?? frame.getAttribute("style") ?? "");
        frame.replaceWith(box);
        index += 1;
      }
    },
    onCloneEachNode: (cloned: Node) => {
      const style = (cloned as HTMLElement).style;
      if (style) {
        // The drawing holds this instant's computed styles: nothing in it
        // may play again from its first keyframe (a turn fading in drew blank).
        style.setProperty("animation", "none", "important");
        style.setProperty("transition", "none", "important");
        // A backdrop blur stops Chrome painting what follows it in a drawn
        // picture: a window's frosted title bar left its whole body blank
        
        // translucent fill; only the blur behind it is not drawn.
        const blur = (name: string) => {
          const value = style.getPropertyValue(name);
          return Boolean(value) && value !== "none";
        };
        if (blur("backdrop-filter") || blur("-webkit-backdrop-filter")) {
          style.setProperty("backdrop-filter", "none", "important");
          style.setProperty("-webkit-backdrop-filter", "none", "important");
          // The blur also made the bar the box its fixed and absolute
          // children are placed in (a menu open in the AI panel's composer):
          // without it they were drawn against the page. An identity
          
          const set = (name: string) => {
            const value = style.getPropertyValue(name);
            return Boolean(value) && value !== "none" && value !== "normal" && value !== "auto";
          };
          const contain = style.getPropertyValue("contain");
          const anchored =
            set("transform") || set("perspective") || set("filter") ||
            /\b(layout|paint|strict|content)\b/.test(contain) ||
            /\b(transform|perspective|filter)\b/.test(style.getPropertyValue("will-change"));
          if (!anchored) style.setProperty("transform", "translateZ(0)");
        }
      }
    },
  });
  context.node = root;
  try {
    return await domToCanvas(context);
  } finally {
    destroyContext(context);
  }
}

/** A frame of this page's origin, drawn on its own. */
async function drawFrame(frame: HTMLIFrameElement, scale: number): Promise<HTMLCanvasElement | null> {
  const document_ = frame.contentDocument;
  if (!document_?.documentElement) return null;
  const { domToCanvas } = await import("modern-screenshot");
  return domToCanvas(document_.documentElement, {
    scale,
    width: frame.clientWidth,
    height: frame.clientHeight,
    timeout: 2_000,
    font: false,
    features: { restoreScrollPosition: true },
  });
}

async function drawPicture(
  plan: RegionPlan,
  square: CircledSquare,
  ink: CircleInk | undefined,
  options: DrawOptions,
  abandoned: () => boolean,
): Promise<Blob | null> {
  const backdrop = backdropOf(plan.root);
  const dark = isDark(backdrop);
  let scale = Math.min(2, Math.max(1, window.devicePixelRatio || 1));
  // The root is drawn whole, then cut to the square: kept within a sane size.
  const area = plan.rootRect.width * plan.rootRect.height;
  if (area * scale * scale > 24_000_000) scale = Math.max(1, Math.sqrt(24_000_000 / area));
  const [drawing, frames] = await Promise.all([
    drawRoot(plan, scale, dark, options),
    Promise.all(
      plan.frames.map(({ frame, shown }) =>
        drawFrame(frame, scale)
          .then((canvas) => ({ frame, shown, canvas }))
          .catch(() => ({ frame, shown, canvas: null })),
      ),
    ),
  ]);
  if (abandoned()) return null;
  const out = document.createElement("canvas");
  out.width = Math.round(square.width * scale);
  out.height = Math.round(square.height * scale);
  const context = out.getContext("2d");
  if (!context) return null;
  context.fillStyle = backdrop;
  context.fillRect(0, 0, out.width, out.height);
  context.drawImage(
    drawing,
    (square.left - plan.rootRect.left) * scale,
    (square.top - plan.rootRect.top) * scale,
    square.width * scale,
    square.height * scale,
    0,
    0,
    out.width,
    out.height,
  );
  const fontFamily = getComputedStyle(plan.root).fontFamily || "sans-serif";
  const place = (shown: Box) => ({
    x: (shown.left - square.left) * scale,
    y: (shown.top - square.top) * scale,
    width: (shown.right - shown.left) * scale,
    height: (shown.bottom - shown.top) * scale,
  });
  for (const { frame, shown, canvas } of frames) {
    if (!canvas) {
      paintStandIn(context, place(shown), frame.title, scale, dark, fontFamily);
      continue;
    }
    const rect = frame.getBoundingClientRect();
    const left = rect.left + frame.clientLeft;
    const top = rect.top + frame.clientTop;
    context.drawImage(
      canvas,
      (shown.left - left) * scale,
      (shown.top - top) * scale,
      (shown.right - shown.left) * scale,
      (shown.bottom - shown.top) * scale,
      (shown.left - square.left) * scale,
      (shown.top - square.top) * scale,
      (shown.right - shown.left) * scale,
      (shown.bottom - shown.top) * scale,
    );
  }
  for (const item of plan.unreadable) {
    const box = place(item.shown);
    if (item.element instanceof HTMLVideoElement && canvasReadable(item.element)) {
      const rect = item.element.getBoundingClientRect();
      context.save();
      context.beginPath();
      context.rect(box.x, box.y, box.width, box.height);
      context.clip();
      context.drawImage(
        item.element,
        (rect.left - square.left) * scale,
        (rect.top - square.top) * scale,
        rect.width * scale,
        rect.height * scale,
      );
      context.restore();
      continue;
    }
    paintStandIn(context, box, item.label, scale, dark, fontFamily);
  }
  if (ink) paintInk(context, ink, square, scale);
  return await new Promise<Blob | null>((resolve) => out.toBlob(resolve, "image/jpeg", 0.92));
}














export async function captureRegionPicture(
  square: CircledSquare,
  options: { ink?: CircleInk; timeoutMs?: number; plainTimeoutMs?: number } = {},
): Promise<Blob | null> {
  const { ink, timeoutMs = 6_000, plainTimeoutMs = 4_000 } = options;
  let plan: RegionPlan;
  try {
    plan = planRegion(square);
  } catch {
    return null;
  }
  const attempt = async (faithful: boolean, budget: number): Promise<Blob | null> => {
    let abandoned = false;
    const deadline = performance.now() + budget;
    try {
      return await Promise.race([
        drawPicture(plan, square, ink, { faces: faithful, pictures: faithful, deadline }, () => abandoned),
        new Promise<null>((resolve) =>
          setTimeout(() => {
            abandoned = true;
            resolve(null);
          }, budget),
        ),
      ]);
    } catch {
      return null;
    }
  };
  return (await attempt(true, timeoutMs)) ?? (await attempt(false, plainTimeoutMs));
}
