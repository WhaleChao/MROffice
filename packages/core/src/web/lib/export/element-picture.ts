/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { fontFaceDescriptors, fontFaceSource } from "../font-face-declarations";

export interface ElementPictureOptions {
  /** The picture's size in CSS px; the element's layout box by default. */
  readonly width?: number;
  readonly height?: number;
  /** Nodes of the element that are not part of the picture (editing controls). */
  readonly omit?: string;
  /**
   * Edit the styled clone before its resources are gathered (drop nodes,
   * place it); `false` means there is no picture to make.
   */
  readonly edit?: (clone: HTMLElement) => boolean | void;
  /** The node serialized; the clone itself by default (a wrapper that places it). */
  readonly frame?: (clone: HTMLElement, doc: Document) => Element;
  /**
   * @font-face rules the caller owns (a deck's own faces), or a function of
   * the families the picture's text names (lower case, unquoted).
   */
  readonly fontRules?:
    | readonly string[]
    | ((families: ReadonlySet<string>) => readonly string[]);
  /** Carry the faces the text is set in (default true). */
  readonly embedFonts?: boolean;
  readonly signal?: AbortSignal;
}

export interface ElementPicture {
  /** XHTML of the styled clone. */
  readonly markup: string;
  /** The rules the markup needs: its faces and its ::before/::after boxes. */
  readonly css: string;
  readonly width: number;
  readonly height: number;
}

export type ElementPictureFailure =
  | "unavailable"
  | "empty"
  | "too_large"
  | "timeout"
  | "aborted"
  | "unrenderable";

export class ElementPictureError extends Error {
  constructor(readonly code: ElementPictureFailure) {
    super(`element_picture_${code}`);
    this.name = "ElementPictureError";
  }
}

/** Inherited properties: written where they differ from the parent's. */
const INHERITED = [
  "color",
  "font-family",
  "font-size",
  "font-weight",
  "font-style",
  "font-stretch",
  "font-variant-caps",
  "font-variant-east-asian",
  "font-variant-ligatures",
  "font-variant-numeric",
  "font-kerning",
  "font-feature-settings",
  "font-variation-settings",
  "font-synthesis-weight",
  "font-synthesis-style",
  "line-height",
  "letter-spacing",
  "word-spacing",
  "text-align",
  "text-align-last",
  "text-indent",
  "text-transform",
  "text-shadow",
  "text-rendering",
  "text-underline-position",
  "text-emphasis-style",
  "text-emphasis-color",
  "text-emphasis-position",
  "white-space",
  "word-break",
  "overflow-wrap",
  "line-break",
  "hyphens",
  "tab-size",
  "direction",
  "writing-mode",
  "text-orientation",
  "text-combine-upright",
  "visibility",
  "list-style-type",
  "list-style-position",
  "list-style-image",
  "quotes",
  "-webkit-text-stroke-width",
  "-webkit-text-stroke-color",
  "-webkit-text-fill-color",
  "border-collapse",
  "border-spacing",
  "caption-side",
  "empty-cells",
  "fill",
  "fill-opacity",
  "fill-rule",
  "stroke",
  "stroke-width",
  "stroke-dasharray",
  "stroke-dashoffset",
  "stroke-linecap",
  "stroke-linejoin",
  "stroke-miterlimit",
  "stroke-opacity",
  "paint-order",
  "text-anchor",
  "clip-rule",
  "ruby-position",
  "image-rendering",
] as const;

/** `white-space` where an engine computes it only as its two longhands. */
const WHITE_SPACE_LONGHANDS = ["white-space-collapse", "text-wrap-mode"] as const;

/** The element's own properties, left out while they hold an initial value. */
const OWN: ReadonlyArray<readonly [string, readonly string[]]> = [
  ["display", []],
  ["position", ["static"]],
  ["top", ["auto"]],
  ["right", ["auto"]],
  ["bottom", ["auto"]],
  ["left", ["auto"]],
  ["z-index", ["auto"]],
  ["float", ["none"]],
  ["clear", ["none"]],
  ["box-sizing", []],
  ["width", ["auto"]],
  ["height", ["auto"]],
  ["min-width", ["auto", "0px"]],
  ["min-height", ["auto", "0px"]],
  ["max-width", ["none"]],
  ["max-height", ["none"]],
  ["aspect-ratio", ["auto"]],
  ["margin-top", ["0px"]],
  ["margin-right", ["0px"]],
  ["margin-bottom", ["0px"]],
  ["margin-left", ["0px"]],
  ["padding-top", ["0px"]],
  ["padding-right", ["0px"]],
  ["padding-bottom", ["0px"]],
  ["padding-left", ["0px"]],
  ["border-top-left-radius", ["0px"]],
  ["border-top-right-radius", ["0px"]],
  ["border-bottom-right-radius", ["0px"]],
  ["border-bottom-left-radius", ["0px"]],
  ["background-color", ["rgba(0, 0, 0, 0)", "transparent"]],
  ["background-image", ["none"]],
  ["background-clip", ["border-box"]],
  ["-webkit-background-clip", ["border-box"]],
  ["box-shadow", ["none"]],
  // A reflection (Picture, Shape and Text Effects, a table's, a chart
  // area's): left out, the presenter's slides and the HTML show drew none.
  ["-webkit-box-reflect", ["none"]],
  ["opacity", ["1"]],
  ["transform", ["none"]],
  ["rotate", ["none"]],
  ["scale", ["none"]],
  ["translate", ["none"]],
  ["transform-style", ["flat"]],
  ["backface-visibility", ["visible"]],
  ["perspective", ["none"]],
  ["filter", ["none"]],
  ["backdrop-filter", ["none"]],
  ["mix-blend-mode", ["normal"]],
  ["isolation", ["auto"]],
  ["clip-path", ["none"]],
  ["overflow-x", ["visible"]],
  ["overflow-y", ["visible"]],
  ["text-overflow", ["clip"]],
  ["vertical-align", ["baseline"]],
  ["unicode-bidi", ["normal"]],
  ["flex-direction", ["row"]],
  ["flex-wrap", ["nowrap"]],
  ["flex-grow", ["0"]],
  ["flex-shrink", ["1"]],
  ["flex-basis", ["auto"]],
  ["order", ["0"]],
  ["justify-content", ["normal"]],
  ["justify-items", ["normal", "legacy"]],
  ["align-items", ["normal"]],
  ["align-self", ["auto"]],
  ["align-content", ["normal"]],
  ["row-gap", ["normal"]],
  ["column-gap", ["normal"]],
  ["grid-template-columns", ["none"]],
  ["grid-template-rows", ["none"]],
  ["grid-template-areas", ["none"]],
  ["grid-auto-flow", ["row"]],
  ["grid-auto-columns", ["auto"]],
  ["grid-auto-rows", ["auto"]],
  ["grid-column-start", ["auto"]],
  ["grid-column-end", ["auto"]],
  ["grid-row-start", ["auto"]],
  ["grid-row-end", ["auto"]],
  ["table-layout", ["auto"]],
  ["object-fit", ["fill"]],
  ["object-position", ["50% 50%"]],
  ["column-count", ["auto"]],
  ["column-width", ["auto"]],
  ["counter-reset", ["none"]],
  ["counter-increment", ["none"]],
  ["counter-set", ["none"]],
  ["-webkit-line-clamp", ["none"]],
  ["-webkit-box-orient", ["horizontal"]],
  ["vector-effect", ["none"]],
];

/** Groups written only when their switch is on (a border side with a style). */
const SIDES = ["top", "right", "bottom", "left"] as const;
const BACKGROUND_LAYOUT = [
  "background-size",
  "background-position",
  "background-repeat",
  "background-origin",
] as const;
const MASK = [
  "-webkit-mask-image",
  "mask-image",
  "-webkit-mask-size",
  "-webkit-mask-position",
  "-webkit-mask-repeat",
  "-webkit-mask-composite",
  "mask-mode",
] as const;

/** Never part of a picture: script, its loaders, and what the caller marks. */
const ALWAYS_OMITTED =
  "script, noscript, template, link, style, meta, [data-office-picture-omit]";

/** A picture over this many CSS px a side is refused rather than rastered. */
const MAX_SIDE = 8192;
/** Faces carried by one picture, at most (bytes before encoding). */
const MAX_FONT_BYTES = 40 * 1024 * 1024;

function value(cs: CSSStyleDeclaration, property: string): string {
  return cs.getPropertyValue(property).trim();
}

const SAME_DOCUMENT_URL_RE = /url\((["']?)([^"')]*#[^"')]+)\1\)/g;

/**
 * A computed `url(...)` that names an element of this document (a gradient,
 * a clip path) reads as an absolute URL in some engines; in the picture it
 * must stay a reference into the picture itself.
 */
function localReferences(text: string, doc: Document): string {
  if (!text.includes("url(")) return text;
  const here = doc.URL.split("#")[0];
  return text.replace(SAME_DOCUMENT_URL_RE, (whole, _quote, url: string) => {
    const [base, fragment] = url.split("#");
    return !base || base === here ? `url("#${fragment}")` : whole;
  });
}

/** jsdom computes no pseudo-element boxes (and says so on every call). */
const PSEUDO_BOXES: readonly ("::before" | "::after")[] =
  typeof navigator !== "undefined" && /jsdom/i.test(navigator.userAgent)
    ? []
    : ["::before", "::after"];

/**
 * Elements the browser's own stylesheet styles: centred bold header cells,
 * cell padding and table spacing, paragraph and list margins, heading sizes,
 * bold / italic / link / code phrases, raised and lowered text, form fields.
 */
const BROWSER_STYLED_TAGS = new Set([
  "table", "caption", "th", "td", "p", "h1", "h2", "h3", "h4", "h5", "h6",
  "ul", "ol", "li", "dl", "dt", "dd", "blockquote", "figure", "figcaption",
  "pre", "code", "kbd", "samp", "tt", "b", "strong", "i", "em", "cite", "dfn",
  "var", "address", "a", "small", "big", "sub", "sup", "mark", "u", "ins", "s",
  "strike", "del", "q", "hr", "button", "input", "select", "textarea",
  "fieldset", "legend", "center", "details", "summary", "abbr", "label", "img",
  // An inner <svg> clips to its box by the browser's own rule: its stated
  // `overflow: visible` (a shape's outline, half of it outside the geometry)
  // left out as the initial value, a slide saved as PNG cut the outline's
  
  "svg",
]);

/** ::before/::after classes stay unique across pictures one file shows together. */
let pictureClassIndex = 0;

interface StyledClone {
  readonly clone: HTMLElement;
  readonly pseudoRules: string[];
  readonly families: Set<string>;
  readonly text: string[];
}

function familyNames(stack: string): string[] {
  return stack
    .split(",")
    .map((name) => name.trim().replace(/^["']|["']$/g, "").trim().toLowerCase())
    .filter(Boolean);
}

/**
 * Clone `root` with the page's computed styles written onto every element.
 * Synchronous: the styles are read from the live element in one pass.
 */
export function cloneWithComputedStyle(
  root: Element,
  omit?: string,
): StyledClone {
  const doc = root.ownerDocument;
  const view = doc.defaultView;
  const clone = root.cloneNode(true) as HTMLElement;
  const pseudoRules: string[] = [];
  const families = new Set<string>();
  const text: string[] = [];
  if (!view) return { clone, pseudoRules, families, text };
  const lives = [root, ...Array.from(root.querySelectorAll("*"))];
  const copies = [clone, ...Array.from(clone.querySelectorAll("*"))];
  if (lives.length !== copies.length) return { clone, pseudoRules, families, text };
  const inheritedOf = new Map<Element, string[]>();
  const removed: Element[] = [];

  // `full`: every value is written, for an element the browser's own
  // stylesheet styles (a header cell's centring, a paragraph's margins): left
  // out because it equals the parent's or the initial value, the picture
  // got the browser's instead (a left-aligned header drew centred).
  const declarations = (
    cs: CSSStyleDeclaration,
    parent: readonly string[] | undefined,
    full = false,
  ): { list: string[]; inherited: string[] } => {
    const list: string[] = [];
    const inherited = INHERITED.map((property) => value(cs, property));
    INHERITED.forEach((property, index) => {
      const own = inherited[index];
      if (own && (full || !parent || parent[index] !== own)) list.push(`${property}:${own}`);
    });
    if (!value(cs, "white-space")) {
      for (const property of WHITE_SPACE_LONGHANDS) {
        const own = value(cs, property);
        if (own) list.push(`${property}:${own}`);
      }
    }
    for (const [property, initial] of OWN) {
      const own = value(cs, property);
      if (own && (full || !initial.includes(own))) list.push(`${property}:${own}`);
    }
    for (const side of SIDES) {
      const style = value(cs, `border-${side}-style`);
      if (!style) continue;
      if (style === "none" || style === "hidden") {
        if (full) list.push(`border-${side}-style:none`);
        continue;
      }
      list.push(
        `border-${side}-style:${style}`,
        `border-${side}-width:${value(cs, `border-${side}-width`)}`,
        `border-${side}-color:${value(cs, `border-${side}-color`)}`,
      );
    }
    const outline = value(cs, "outline-style");
    if (outline && outline !== "none") {
      for (const property of ["outline-style", "outline-width", "outline-color", "outline-offset"]) {
        list.push(`${property}:${value(cs, property)}`);
      }
    }
    const decoration = value(cs, "text-decoration-line");
    if (decoration && decoration !== "none") {
      for (const property of [
        "text-decoration-line",
        "text-decoration-style",
        "text-decoration-color",
        "text-decoration-thickness",
        "text-underline-offset",
      ]) {
        const own = value(cs, property);
        if (own) list.push(`${property}:${own}`);
      }
    }
    if (
      ["transform", "rotate", "scale", "translate", "perspective"].some((property) => {
        const own = value(cs, property);
        return own && own !== "none";
      })
    ) {
      list.push(`transform-origin:${value(cs, "transform-origin")}`);
      const origin = value(cs, "perspective-origin");
      if (origin) list.push(`perspective-origin:${origin}`);
    }
    if (value(cs, "background-image") !== "none") {
      for (const property of BACKGROUND_LAYOUT) {
        const own = value(cs, property);
        if (own) list.push(`${property}:${own}`);
      }
    }
    const mask = value(cs, "-webkit-mask-image") || value(cs, "mask-image");
    if (mask && mask !== "none") {
      for (const property of MASK) {
        const own = value(cs, property);
        if (own) list.push(`${property}:${own}`);
      }
    }
    return {
      list: list
        .filter((declaration) => !declaration.endsWith(":"))
        .map((declaration) => localReferences(declaration, doc)),
      inherited,
    };
  };

  lives.forEach((live, index) => {
    const copy = copies[index];
    if (!(copy instanceof HTMLElement) && !(copy instanceof SVGElement)) return;
    if (removed.some((gone) => gone.contains(copy))) return;
    const cs = view.getComputedStyle(live);
    if (value(cs, "display") === "none") {
      removed.push(copy);
      return;
    }
    const parentInherited =
      live === root ? undefined : inheritedOf.get(live.parentElement as Element);
    const { list, inherited } = declarations(
      cs,
      parentInherited,
      BROWSER_STYLED_TAGS.has(live.localName),
    );
    inheritedOf.set(live, inherited);
    copy.removeAttribute("style");
    for (const declaration of list) {
      const colon = declaration.indexOf(":");
      copy.style.setProperty(declaration.slice(0, colon), declaration.slice(colon + 1));
    }
    for (const name of familyNames(value(cs, "font-family"))) families.add(name);
    for (const pseudo of PSEUDO_BOXES) {
      const box = view.getComputedStyle(live, pseudo);
      const content = value(box, "content");
      if (!content || content === "none" || content === "normal") continue;
      if (value(box, "display") === "none") continue;
      const rule = declarations(box, inherited).list;
      const name = `mrai-picture-${(pictureClassIndex += 1)}`;
      copy.classList.add(name);
      pseudoRules.push(`.${name}${pseudo}{content:${content};${rule.join(";")}}`);
      text.push(content);
    }
    // Attributes that load or run something are not the picture's.
    for (const attribute of Array.from(copy.attributes)) {
      const name = attribute.name.toLowerCase();
      if (name.startsWith("on") || name === "srcset" || name === "sizes" || name === "loading") {
        copy.removeAttribute(attribute.name);
      } else if (
        (name === "href" || name === "xlink:href") &&
        /^\s*javascript:/i.test(attribute.value)
      ) {
        copy.removeAttribute(attribute.name);
      }
    }
    if (live instanceof HTMLImageElement && live.currentSrc) {
      copy.setAttribute("src", live.currentSrc);
    }
  });
  for (const node of removed) node.remove();
  // What the element shows only while it is being edited, or never.
  const omitted = omit ? `${ALWAYS_OMITTED}, ${omit}` : ALWAYS_OMITTED;
  for (const node of Array.from(clone.querySelectorAll(omitted))) node.remove();
  // A text editor in progress shows its words with the editor's own look.
  for (const editor of Array.from(clone.querySelectorAll("textarea"))) {
    const words = doc.createElement("div");
    words.setAttribute(
      "style",
      `${editor.getAttribute("style") ?? ""};resize:none;white-space:pre-wrap`,
    );
    words.textContent = (editor as HTMLTextAreaElement).value || editor.textContent;
    editor.replaceWith(words);
  }
  clone.removeAttribute("tabindex");
  text.push(clone.textContent ?? "");
  return { clone, pseudoRules, families, text };
}

// ── resources ───────────────────────────────────────────────────────────────

const dataUrlCache = new Map<string, Promise<string | null>>();
const DATA_URL_CACHE_ENTRIES = 64;






export const ELEMENT_PICTURE_FETCH_DEADLINE_MS = 15_000;

function blobAsDataUrl(blob: Blob): Promise<string | null> {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onerror = () => resolve(null);
    reader.onload = () =>
      resolve(typeof reader.result === "string" ? reader.result : null);
    reader.readAsDataURL(blob);
  });
}

/** A URL's bytes as a data URL (cached across pictures), or null. */
function dataUrlOf(url: string, doc: Document, signal?: AbortSignal): Promise<string | null> {
  if (url.startsWith("data:")) return Promise.resolve(url);
  let absolute: URL;
  try {
    absolute = new URL(url, doc.baseURI);
  } catch {
    return Promise.resolve(null);
  }
  const key = absolute.href;
  const cached = dataUrlCache.get(key);
  if (cached) return cached;
  const sameOrigin =
    absolute.protocol === "blob:" || absolute.origin === new URL(doc.baseURI).origin;
  // The caller's signal or the deadline ends the request, its body included.
  const controller = typeof AbortController === "function" ? new AbortController() : null;
  const stop = () => controller?.abort();
  if (signal?.aborted) stop();
  signal?.addEventListener("abort", stop, { once: true });
  const timer = setTimeout(stop, ELEMENT_PICTURE_FETCH_DEADLINE_MS);
  const ended = new Promise<null>((resolve) => {
    if (!controller) return;
    if (controller.signal.aborted) resolve(null);
    else controller.signal.addEventListener("abort", () => resolve(null), { once: true });
  });
  const request = fetch(key, {
    credentials: sameOrigin ? "same-origin" : "omit",
    signal: controller?.signal ?? signal,
  })
    .then((response) => (response.ok ? response.blob() : null))
    .then((blob) => (blob ? blobAsDataUrl(blob) : null))
    .catch(() => null);
  // A fetch that ignores its signal still lets the picture go on.
  const pending = Promise.race([request, ended]).finally(() => {
    clearTimeout(timer);
    signal?.removeEventListener("abort", stop);
  });
  dataUrlCache.set(key, pending);
  void pending.then((result) => {
    if (result === null) dataUrlCache.delete(key);
  });
  while (dataUrlCache.size > DATA_URL_CACHE_ENTRIES) {
    const oldest = dataUrlCache.keys().next().value;
    if (oldest === undefined) break;
    dataUrlCache.delete(oldest);
  }
  return pending;
}

/** The pixels a live picture element shows, where the page may read them. */
function drawnPixels(source: CanvasImageSource, width: number, height: number): string | null {
  if (!(width > 0) || !(height > 0)) return null;
  try {
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d");
    if (!context) return null;
    context.drawImage(source, 0, 0, width, height);
    return canvas.toDataURL("image/png");
  } catch {
    // A picture from another origin taints the canvas; it stays as it was.
    return null;
  }
}

const CSS_URL_RE = /url\((["']?)([^"')]+)\1\)/g;

async function embedCssUrls(declarations: string, doc: Document, signal?: AbortSignal): Promise<string> {
  const urls = new Set<string>();
  for (const match of declarations.matchAll(CSS_URL_RE)) {
    if (!match[2].startsWith("data:") && !match[2].startsWith("#")) urls.add(match[2]);
  }
  if (!urls.size) return declarations;
  const embedded = new Map<string, string | null>();
  await Promise.all(
    [...urls].map(async (url) => embedded.set(url, await dataUrlOf(url, doc, signal))),
  );
  return declarations.replace(CSS_URL_RE, (whole, _quote, url: string) => {
    const data = embedded.get(url);
    return data ? `url("${data}")` : whole;
  });
}

/** Pictures, canvases and video frames of the clone, as data. */
async function embedPictures(
  live: Element,
  clone: HTMLElement,
  signal?: AbortSignal,
): Promise<void> {
  const doc = live.ownerDocument;
  const tasks: Promise<void>[] = [];
  // Canvases and videos first, matched to the live element they copy.
  const liveCanvases = Array.from(live.querySelectorAll("canvas"));
  const cloneCanvases = Array.from(clone.querySelectorAll("canvas"));
  if (liveCanvases.length === cloneCanvases.length) {
    cloneCanvases.forEach((copy, index) => {
      const source = liveCanvases[index];
      const pixels = drawnPixels(source, source.width, source.height);
      const picture = doc.createElement("img");
      picture.setAttribute("style", copy.getAttribute("style") ?? "");
      if (pixels) picture.setAttribute("src", pixels);
      copy.replaceWith(picture);
    });
  }
  const liveVideos = Array.from(live.querySelectorAll("video"));
  const cloneVideos = Array.from(clone.querySelectorAll("video"));
  if (liveVideos.length === cloneVideos.length) {
    cloneVideos.forEach((copy, index) => {
      const source = liveVideos[index];
      const picture = doc.createElement("img");
      picture.setAttribute("style", `${copy.getAttribute("style") ?? ""};object-fit:contain;background:#000`);
      const frame = source.readyState >= 2
        ? drawnPixels(source, source.videoWidth, source.videoHeight)
        : null;
      if (frame) picture.setAttribute("src", frame);
      else if (source.poster) picture.setAttribute("src", source.poster);
      copy.replaceWith(picture);
    });
  }
  for (const node of Array.from(clone.querySelectorAll("audio, iframe, object, embed"))) {
    const stand = doc.createElement("div");
    stand.setAttribute("style", node.getAttribute("style") ?? "");
    node.replaceWith(stand);
  }
  const liveImages = Array.from(live.querySelectorAll("img"));
  for (const picture of Array.from(clone.querySelectorAll("img"))) {
    const src = picture.getAttribute("src");
    if (!src || src.startsWith("data:")) continue;
    tasks.push(
      dataUrlOf(src, doc, signal).then((data) => {
        if (data) {
          picture.setAttribute("src", data);
          return;
        }
        const source = liveImages.find((image) => image.currentSrc === src || image.src === src);
        const pixels = source?.complete
          ? drawnPixels(source, source.naturalWidth, source.naturalHeight)
          : null;
        if (pixels) picture.setAttribute("src", pixels);
        else picture.removeAttribute("src");
      }),
    );
  }
  for (const image of Array.from(clone.querySelectorAll("image"))) {
    for (const attribute of ["href", "xlink:href"]) {
      const src = image.getAttribute(attribute);
      if (!src || src.startsWith("data:") || src.startsWith("#")) continue;
      tasks.push(
        dataUrlOf(src, doc, signal).then((data) => {
          if (data) image.setAttribute(attribute, data);
        }),
      );
    }
  }
  for (const node of [clone, ...Array.from(clone.querySelectorAll<HTMLElement>("[style*='url(']"))]) {
    const style = node.getAttribute("style");
    if (!style || !style.includes("url(")) continue;
    tasks.push(
      embedCssUrls(style, doc, signal).then((embedded) => node.setAttribute("style", embedded)),
    );
  }
  await Promise.all(tasks);
}

// ── faces ───────────────────────────────────────────────────────────────────

type Range = readonly [number, number];

function parseRanges(text: string | undefined): Range[] {
  if (!text) return [[0, 0x10ffff]];
  const ranges: Range[] = [];
  for (const part of text.split(",")) {
    const token = part.trim().replace(/^u\+/i, "");
    if (!token) continue;
    if (token.includes("?")) {
      ranges.push([
        Number.parseInt(token.replace(/\?/g, "0"), 16),
        Number.parseInt(token.replace(/\?/g, "f"), 16),
      ]);
      continue;
    }
    const [low, high] = token.split("-");
    const start = Number.parseInt(low, 16);
    const end = high === undefined ? start : Number.parseInt(high, 16);
    if (Number.isFinite(start) && Number.isFinite(end)) ranges.push([start, end]);
  }
  return ranges.length ? ranges : [[0, 0x10ffff]];
}

function rangesKey(ranges: readonly Range[]): string {
  return ranges.map(([low, high]) => `${low}-${high}`).join(",");
}

function covers(ranges: readonly Range[], codepoints: ReadonlySet<number>): boolean {
  for (const codepoint of codepoints) {
    for (const [low, high] of ranges) if (codepoint >= low && codepoint <= high) return true;
  }
  return false;
}

function weightKey(weight: string): string {
  const named: Record<string, string> = { normal: "400", bold: "700" };
  return weight
    .trim()
    .split(/\s+/)
    .map((part) => named[part] ?? part)
    .join(" ");
}

function faceKey(family: string, weight: string, style: string, ranges: readonly Range[]): string {
  return [
    familyNames(family)[0] ?? "",
    weightKey(weight || "400"),
    (style || "normal").trim(),
    rangesKey(ranges),
  ].join("|");
}

interface DeclaredFace {
  readonly src: string;
  readonly base: string;
}

/** The stylesheet faces of a document, by family / weight / style / range. */
function stylesheetFaces(doc: Document): Map<string, DeclaredFace> {
  const faces = new Map<string, DeclaredFace>();
  const visit = (rules: CSSRuleList | undefined, base: string) => {
    if (!rules) return;
    for (const rule of Array.from(rules)) {
      if (rule.type === CSSRule.FONT_FACE_RULE) {
        const style = (rule as CSSFontFaceRule).style;
        faces.set(
          faceKey(
            style.getPropertyValue("font-family"),
            style.getPropertyValue("font-weight"),
            style.getPropertyValue("font-style"),
            parseRanges(style.getPropertyValue("unicode-range") || undefined),
          ),
          { src: style.getPropertyValue("src"), base },
        );
      } else if (rule.type === CSSRule.IMPORT_RULE) {
        const sheet = (rule as CSSImportRule).styleSheet;
        try {
          visit(sheet?.cssRules, sheet?.href ?? base);
        } catch {
          // Another origin's sheet keeps its rules to itself.
        }
      } else if ("cssRules" in rule) {
        visit((rule as CSSGroupingRule).cssRules, base);
      }
    }
  };
  for (const sheet of Array.from(doc.styleSheets)) {
    try {
      visit(sheet.cssRules, sheet.href ?? doc.baseURI);
    } catch {
      // Another origin's sheet keeps its rules to itself.
    }
  }
  return faces;
}

/** Split a `src` list on its top-level commas. */
function sourceEntries(src: string): string[] {
  const entries: string[] = [];
  let depth = 0;
  let quote = "";
  let start = 0;
  for (let index = 0; index < src.length; index += 1) {
    const character = src[index];
    if (quote) {
      if (character === quote) quote = "";
    } else if (character === '"' || character === "'") {
      quote = character;
    } else if (character === "(") {
      depth += 1;
    } else if (character === ")") {
      depth -= 1;
    } else if (character === "," && depth === 0) {
      entries.push(src.slice(start, index).trim());
      start = index + 1;
    }
  }
  entries.push(src.slice(start).trim());
  return entries.filter(Boolean);
}

/** The face's `src` with its first fetched file as data; `local()` kept ahead of it. */
async function embeddedSource(
  declared: DeclaredFace,
  doc: Document,
  signal?: AbortSignal,
): Promise<{ src: string; bytes: number } | null> {
  const kept: string[] = [];
  for (const entry of sourceEntries(declared.src)) {
    if (/^local\(/i.test(entry)) {
      kept.push(entry);
      continue;
    }
    const match = /^url\((["']?)([^"')]+)\1\)\s*(.*)$/i.exec(entry);
    if (!match) continue;
    let absolute: string;
    try {
      absolute = new URL(match[2], declared.base).href;
    } catch {
      continue;
    }
    const data = await dataUrlOf(absolute, doc, signal);
    if (!data) continue;
    kept.push(`url("${data}")${match[3] ? ` ${match[3]}` : ""}`);
    return { src: kept.join(","), bytes: Math.round((data.length * 3) / 4) };
  }
  return null;
}

/** @font-face rules, as data, for the loaded faces the picture's text is set in. */
async function fontRulesFor(
  doc: Document,
  families: ReadonlySet<string>,
  text: string,
  signal?: AbortSignal,
): Promise<string[]> {
  const fonts = doc.fonts as FontFaceSet | undefined;
  if (!fonts || typeof fonts.forEach !== "function") return [];
  const codepoints = new Set<number>();
  for (const character of text) {
    const codepoint = character.codePointAt(0);
    if (codepoint !== undefined && codepoint > 0x20) codepoints.add(codepoint);
  }
  if (!codepoints.size) return [];
  const loaded: FontFace[] = [];
  fonts.forEach((face) => {
    if (face.status !== "loaded") return;
    const family = familyNames(face.family)[0];
    if (!family || !families.has(family)) return;
    if (!covers(parseRanges(face.unicodeRange), codepoints)) return;
    loaded.push(face);
  });
  if (!loaded.length) return [];
  const declared = stylesheetFaces(doc);
  const rules: string[] = [];
  let bytes = 0;
  const seen = new Set<string>();
  for (const face of loaded) {
    const ranges = parseRanges(face.unicodeRange);
    const key = faceKey(face.family, face.weight, face.style, ranges);
    if (seen.has(key)) continue;
    seen.add(key);
    const noted = fontFaceSource(face);
    const source = noted
      ? { src: noted, base: doc.baseURI }
      : declared.get(key);
    if (!source) continue;
    const embedded = await embeddedSource(source, doc, signal);
    if (!embedded) continue;
    if (bytes + embedded.bytes > MAX_FONT_BYTES) break;
    bytes += embedded.bytes;
    const family = (familyNames(face.family)[0] ? face.family : "").replace(/^["']|["']$/g, "");
    const descriptors = [
      `font-family:"${family.replace(/["\\]/g, "\\$&")}"`,
      `src:${embedded.src}`,
      `font-weight:${face.weight}`,
      `font-style:${face.style}`,
      `font-stretch:${face.stretch}`,
      `unicode-range:${face.unicodeRange}`,
      "font-display:block",
    ];
    const sizeAdjust = (face as FontFace & { sizeAdjust?: string }).sizeAdjust;
    if (sizeAdjust && sizeAdjust !== "100%") descriptors.push(`size-adjust:${sizeAdjust}`);
    // An alias drawn at its variable stand-in's own weight (office-fonts addFace).
    const variationSettings =
      fontFaceDescriptors(face)?.variationSettings ?? (face as FontFace & { variationSettings?: string }).variationSettings;
    if (variationSettings && variationSettings !== "normal") descriptors.push(`font-variation-settings:${variationSettings}`);
    rules.push(`@font-face{${descriptors.join(";")}}`);
  }
  return rules;
}

// ── the picture ─────────────────────────────────────────────────────────────

function aborted(signal?: AbortSignal): void {
  if (signal?.aborted) throw new ElementPictureError("aborted");
}

function layoutSize(root: Element): { width: number; height: number } {
  if (root instanceof HTMLElement && root.offsetWidth > 0) {
    return { width: root.offsetWidth, height: root.offsetHeight };
  }
  const rect = root.getBoundingClientRect();
  return { width: rect.width, height: rect.height };
}

/**
 * `root` as the page draws it: XHTML with every style written in, plus the
 * rules (faces, ::before/::after boxes) it needs. Throws
 * {@link ElementPictureError}.
 */
export async function captureElementPicture(
  root: Element,
  options: ElementPictureOptions = {},
): Promise<ElementPicture> {
  if (typeof document === "undefined" || typeof XMLSerializer === "undefined") {
    throw new ElementPictureError("unavailable");
  }
  const measured = layoutSize(root);
  const width = Math.ceil(options.width ?? measured.width);
  const height = Math.ceil(options.height ?? measured.height);
  if (!(width > 0) || !(height > 0)) throw new ElementPictureError("empty");
  if (width > MAX_SIDE || height > MAX_SIDE) throw new ElementPictureError("too_large");
  const { clone, pseudoRules, families, text } = cloneWithComputedStyle(root, options.omit);
  if (options.edit?.(clone) === false) throw new ElementPictureError("empty");
  await embedPictures(root, clone, options.signal);
  aborted(options.signal);
  const fontRules =
    options.embedFonts === false
      ? []
      : await fontRulesFor(root.ownerDocument, families, text.join(""), options.signal);
  aborted(options.signal);
  const framed = options.frame?.(clone, root.ownerDocument) ?? clone;
  const markup = new XMLSerializer().serializeToString(framed);
  const ownRules =
    typeof options.fontRules === "function"
      ? options.fontRules(families)
      : options.fontRules ?? [];
  const css = [
    ...ownRules,
    ...fontRules,
    ...pseudoRules,
    // The picture is a moment: nothing in it moves.
    "*,*::before,*::after{animation:none!important;transition:none!important}",
  ].join("\n");
  return { markup, css, width, height };
}

/** CSS text as the content of an XML `<style>` element. */
function xmlStyleText(css: string): string {
  return `<![CDATA[${css.replace(/]]>/g, "]]]]><![CDATA[>")}]]>`;
}

/** The picture as a standalone SVG document. */
export function elementPictureSvg(
  picture: ElementPicture,
  background?: string | null,
): string {
  const { width, height } = picture;
  const fill =
    background && /^#(?:[0-9a-fA-F]{3,8})$|^[a-zA-Z]+$/.test(background)
      ? `background:${background};`
      : "";
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">` +
    `<foreignObject x="0" y="0" width="${width}" height="${height}">` +
    `<div xmlns="http://www.w3.org/1999/xhtml" style="width:${width}px;height:${height}px;${fill}overflow:hidden;margin:0;padding:0">` +
    `<style>${xmlStyleText(picture.css)}</style>` +
    picture.markup +
    `</div></foreignObject></svg>`
  );
}







export function nextFrame(): Promise<void> {
  return new Promise((resolve) => {
    let done = false;
    const finish = () => {
      if (done) return;
      done = true;
      resolve();
    };
    if (typeof requestAnimationFrame === "function") requestAnimationFrame(finish);
    setTimeout(finish, typeof requestAnimationFrame === "function" ? 100 : 16);
  });
}

function withDeadline<T>(
  work: Promise<T>,
  timeoutMs: number,
  signal?: AbortSignal,
): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => reject(new ElementPictureError("timeout")), timeoutMs);
    const stop = () => {
      clearTimeout(timer);
      reject(new ElementPictureError("aborted"));
    };
    signal?.addEventListener("abort", stop, { once: true });
    work.then(
      (result) => {
        clearTimeout(timer);
        signal?.removeEventListener("abort", stop);
        resolve(result);
      },
      (error: unknown) => {
        clearTimeout(timer);
        signal?.removeEventListener("abort", stop);
        reject(error);
      },
    );
  });
}

/**
 * The picture as a PNG. `scale` is the device-pixel multiplier (clamped so
 * the canvas stays within what every engine allocates). Throws
 * {@link ElementPictureError}; never waits longer than `timeoutMs`.
 */
/** The picture files {@link rasterElementPicture} writes: a canvas encodes
 * PNG and JPEG; BMP and TIFF are written from its pixels (`raster-encoders`). */
export type ElementPictureFormat = "png" | "jpeg" | "bmp" | "tiff";

export function rasterElementPicture(
  picture: ElementPicture,
  options: {
    scale?: number;
    background?: string | null;
    timeoutMs?: number;
    signal?: AbortSignal;
    /** The file written; PNG when not given. JPEG, BMP and TIFF are opaque
     * (give a `background`). */
    format?: ElementPictureFormat;
    /** The resolution a BMP or TIFF states (dots per inch). */
    dpi?: number;
  } = {},
): Promise<Blob> {
  if (typeof document === "undefined" || typeof Image === "undefined") {
    return Promise.reject(new ElementPictureError("unavailable"));
  }
  const requested =
    typeof options.scale === "number" && Number.isFinite(options.scale) && options.scale > 0
      ? options.scale
      : 1;
  // 16.7 M pixels is WebKit's canvas ceiling.
  const ceiling = Math.sqrt(16_000_000 / (picture.width * picture.height));
  const scale = Math.max(0.25, Math.min(4, requested, ceiling));
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(picture.width * scale));
  canvas.height = Math.max(1, Math.round(picture.height * scale));
  let context: CanvasRenderingContext2D | null = null;
  try {
    context = canvas.getContext("2d");
  } catch {
    context = null;
  }
  if (!context) return Promise.reject(new ElementPictureError("unavailable"));
  const pen = context;
  const work = (async () => {
    const image = new Image();
    image.decoding = "sync";
    await new Promise<void>((resolve, reject) => {
      image.onload = () => resolve();
      image.onerror = () => reject(new ElementPictureError("unrenderable"));
      image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(
        elementPictureSvg(picture, options.background),
      )}`;
    });
    try {
      await image.decode();
    } catch {
      // Decoded on draw instead.
    }
    const paint = () => {
      pen.clearRect(0, 0, canvas.width, canvas.height);
      if (options.background) {
        pen.fillStyle = options.background;
        pen.fillRect(0, 0, canvas.width, canvas.height);
      }
      pen.drawImage(image, 0, 0, canvas.width, canvas.height);
    };
    // WebKit draws a foreignObject's faces and pictures only once the image
    // has been drawn before: draw, give it two frames, draw again.
    paint();
    await nextFrame();
    await nextFrame();
    paint();
    const format = options.format ?? "png";
    if (format === "bmp" || format === "tiff") {
      let pixels: ImageData;
      try {
        pixels = pen.getImageData(0, 0, canvas.width, canvas.height);
      } catch {
        throw new ElementPictureError("unrenderable");
      }
      const { encodeBmp, encodeTiff } = await import("./raster-encoders");
      const encoded = (format === "bmp" ? encodeBmp : encodeTiff)(pixels, options.dpi ?? 96);
      return new Blob([encoded as BlobPart], { type: format === "bmp" ? "image/bmp" : "image/tiff" });
    }
    return new Promise<Blob>((resolve, reject) => {
      try {
        canvas.toBlob(
          (blob) => {
            if (blob) resolve(blob);
            else reject(new ElementPictureError("unrenderable"));
          },
          format === "jpeg" ? "image/jpeg" : "image/png",
          // PowerPoint's JPEG of a slide is near lossless.
          format === "jpeg" ? 0.95 : undefined,
        );
      } catch {
        reject(new ElementPictureError("unrenderable"));
      }
    });
  })();
  return withDeadline(work, options.timeoutMs ?? 20_000, options.signal);
}
