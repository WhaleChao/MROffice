/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { EditorView } from "@tiptap/pm/view";

import type { FloatPaint } from "../../../../../lib/office-float-anchor-edit";

import { docsLayoutPerfLegacy } from "./layoutPerfLegacy";

/** Fired when a band layer has drawn (mounted or moved) its elements. */
export const BAND_LAYER_DRAWN_EVENT = "mrn-band-layer-drawn";

let announcePending = false;

/**
 * Once per task, after everything the task draws: each band box announces
 * from its own layout effect, and a commit that redraws a long document's
 * 126 boxes cut every band 126 times, each cut reading the boxes the last one
 * had just written (R286-D01: 29 s of a 276 s open). A microtask still runs
 * before the frame paints, so a strip is never shown uncut.
 */
export function announceBandLayerDrawn(): void {
  if (typeof window === "undefined" || announcePending) return;
  if (docsLayoutPerfLegacy()) {
    window.dispatchEvent(new Event(BAND_LAYER_DRAWN_EVENT));
    return;
  }
  announcePending = true;
  queueMicrotask(() => {
    announcePending = false;
    window.dispatchEvent(new Event(BAND_LAYER_DRAWN_EVENT));
  });
}

/** A body object over a band area: the element and the index of its page. */
export interface BandIntruder {
  readonly element: HTMLElement;
  readonly sheetIndex: number;
}

export interface BandHoles {
  intruders: BandIntruder[];
  /** The strips cut now: the clip each was given, and what the CSSOM made of it. */
  readonly cut: Map<HTMLElement, { asked: string; seen: string }>;
  /** The painted bands masked now, by the signature of their mask. */
  readonly masked: Map<HTMLElement, string>;
}

export function createBandHoles(): BandHoles {
  return { intruders: [], cut: new Map(), masked: new Map() };
}

/** The selection grips stand out of the object; the press on them is the object's too. */
const HOLE_PAD_PX = 2;

/**
 * The placed objects (the float pass's elements and readings) that reach
 * into their page's top or bottom margin.
 */
export function bodyObjectsInBandAreas(
  view: EditorView,
  elements: readonly HTMLElement[],
  paintOf: (element: HTMLElement) => FloatPaint | null,
): BandIntruder[] {
  const out: BandIntruder[] = [];
  let sheets: Element[] | null = null;
  for (const element of elements) {
    const paint = paintOf(element);
    if (!paint || element.closest("td, th")) continue;
    const { marginTopPx, marginBottomPx, pageHeightPx } = paint.geometry;
    const intoHeader = paint.topPx < marginTopPx;
    const intoFooter = paint.topPx + paint.heightPx > pageHeightPx - marginBottomPx;
    if (!intoHeader && !intoFooter) continue;
    const sheet = element.closest("article[data-page-sheet]");
    if (!sheet || sheet.parentElement !== view.dom) continue;
    sheets ??= Array.from(view.dom.children).filter((child) => child.matches("article[data-page-sheet]"));
    const sheetIndex = sheets.indexOf(sheet);
    if (sheetIndex >= 0) out.push({ element, sheetIndex });
  }
  return out;
}

/** What the object paints and the grips it shows, on screen. */
function objectRects(element: HTMLElement): DOMRect[] {
  const body =
    element.matches("figure[data-document-shape]")
      ? element.querySelector<HTMLElement>("[data-document-shape-canvas]")
      : element.matches("div[data-text-box]")
        ? element.querySelector<HTMLElement>(":scope > [data-text-box-frame]")
        : element;
  const rects: DOMRect[] = [];
  const add = (target: Element | null) => {
    const rect = target?.getBoundingClientRect();
    if (rect && rect.width > 0 && rect.height > 0) rects.push(rect);
  };
  add(body ?? element);
  for (const grip of element.querySelectorAll("[data-document-selection-chrome]")) add(grip);
  return rects;
}

function intersects(a: DOMRect, b: DOMRect): boolean {
  return a.left < b.right && b.left < a.right && a.top < b.bottom && b.top < a.bottom;
}

/** An even-odd clip of `element` with `holes` (screen rects) cut out of it. */
function clipWithHoles(element: HTMLElement, holes: readonly DOMRect[]): string {
  const rect = element.getBoundingClientRect();
  const scaleX = element.offsetWidth > 0 ? rect.width / element.offsetWidth : 1;
  const scaleY = element.offsetHeight > 0 ? rect.height / element.offsetHeight : 1;
  const px = (value: number) => Math.round(value * 100) / 100;
  // Far past the element's own box: a tall letterhead overflows its band box.
  let path = "M-100000 -100000H100000V100000H-100000Z";
  for (const hole of holes) {
    const left = (hole.left - rect.left) / (scaleX || 1) - HOLE_PAD_PX;
    const top = (hole.top - rect.top) / (scaleY || 1) - HOLE_PAD_PX;
    const width = hole.width / (scaleX || 1) + HOLE_PAD_PX * 2;
    const height = hole.height / (scaleY || 1) + HOLE_PAD_PX * 2;
    path += `M${px(left)} ${px(top)}h${px(width)}v${px(height)}h${px(-width)}Z`;
  }
  return `path(evenodd, "${path}")`;
}

/**
 * One piece of a painted band's mask: a drawing in the band's own px (SVG
 * markup, turned and mirrored as the object is drawn), or a picture the
 * mask reads by its own pixels. `key` says what it is and where, cheaply.
 */
type PaintLayer =
  | { readonly kind: "svg"; readonly markup: () => string; readonly key: string }
  | { readonly kind: "image"; readonly image: string; readonly rect: DOMRect; readonly key: string };

const SVG_NS = "http://www.w3.org/2000/svg";
const OPAQUE = "linear-gradient(#000, #000)";
/** A drawing's markup (its SVG's children, made self-contained), by the node and the SVG it draws. */
const drawingMarkups = new WeakMap<object, Map<number, string>>();
/** Short ids for nodes and pictures, so a mask is compared without its markup. */
const objectIds = new WeakMap<object, number>();
const pictureIds = new Map<string, number>();
let nextId = 1;

function objectId(object: object): number {
  let id = objectIds.get(object);
  if (id === undefined) {
    id = nextId++;
    objectIds.set(object, id);
  }
  return id;
}

function pictureId(src: string): number {
  let id = pictureIds.get(src);
  if (id === undefined) {
    if (pictureIds.size > 256) pictureIds.clear();
    id = nextId++;
    pictureIds.set(src, id);
  }
  return id;
}

/**
 * Whether the mask may read a picture's pixels: one of the page's own
 * (data:, blob:, this origin), or another host's that answers a CORS request
 * (our file host does). Asked once per picture with crossOrigin=anonymous;
 * until it answers -- and when it refuses -- the picture is an opaque box,
 * so it is drawn over the header's text, never under it.
 */
const pictureReadable = new Map<string, "yes" | "no" | "asking">();

function readablePicture(src: string, documentOrigin: string): boolean {
  if (/^(data|blob):/i.test(src)) return true;
  let origin: string;
  try {
    origin = new URL(src, documentOrigin).origin;
  } catch {
    return false;
  }
  if (origin === documentOrigin) return true;
  const known = pictureReadable.get(src);
  if (known) return known === "yes";
  if (pictureReadable.size > 256) pictureReadable.clear();
  pictureReadable.set(src, "asking");
  const probe = new Image();
  probe.crossOrigin = "anonymous";
  probe.onload = () => {
    pictureReadable.set(src, "yes");
    announceBandLayerDrawn();
  };
  probe.onerror = () => pictureReadable.set(src, "no");
  probe.src = src;
  return false;
}

function shapeNodeAt(view: EditorView, element: HTMLElement): object | null {
  try {
    const pos = view.posAtDOM(element, 0);
    for (const at of [pos, pos - 1]) {
      const node = at >= 0 ? view.state.doc.nodeAt(at) : null;
      if (node?.type.name === "documentShape" && view.nodeDOM(at) === element) return node;
    }
  } catch {
    // Not in the document any more.
  }
  return null;
}

/**
 * An SVG's children as markup a mask image can draw on its own: a picture
 * it shows from a URL (an SVG drawn as an image loads nothing) becomes the
 * box it fills, drawn opaque.
 */
function drawingMarkup(svg: SVGSVGElement): string {
  const clone = svg.cloneNode(true) as SVGSVGElement;
  for (const image of Array.from(clone.querySelectorAll("image"))) {
    const href = image.getAttribute("href") ?? image.getAttribute("xlink:href") ?? "";
    if (/^data:/i.test(href)) continue;
    const box = clone.ownerDocument.createElementNS(SVG_NS, "rect");
    for (const name of ["x", "y", "width", "height", "transform", "clip-path", "mask"]) {
      const value = image.getAttribute(name);
      if (value !== null) box.setAttribute(name, value);
    }
    box.setAttribute("fill", "#000");
    image.replaceWith(box);
  }
  const serializer = new XMLSerializer();
  return Array.from(clone.childNodes, (child) => serializer.serializeToString(child)).join("");
}

/** The band's px from screen px: the band's corner and its scale. */
interface BandFrame {
  readonly left: number;
  readonly top: number;
  readonly scaleX: number;
  readonly scaleY: number;
}

function bandFrameOf(band: HTMLElement): BandFrame {
  const box = band.getBoundingClientRect();
  return {
    left: box.left,
    top: box.top,
    scaleX: band.offsetWidth > 0 ? box.width / band.offsetWidth || 1 : 1,
    scaleY: band.offsetHeight > 0 ? box.height / band.offsetHeight || 1 : 1,
  };
}

const num = (value: number) => String(Math.round(value * 1000) / 1000);

/** An opaque polygon over `points` (screen px), in the band's px. */
function opaqueQuad(points: ReadonlyArray<{ x: number; y: number }>, frame: BandFrame, key: string): PaintLayer {
  const local = points.map((point) => `${num((point.x - frame.left) / frame.scaleX)},${num((point.y - frame.top) / frame.scaleY)}`);
  const markup = `<polygon points="${local.join(" ")}" fill="#000"/>`;
  return { kind: "svg", markup: () => markup, key: `${key}:${local.join(" ")}` };
}

function rectCorners(rect: DOMRect): Array<{ x: number; y: number }> {
  return [
    { x: rect.left, y: rect.top },
    { x: rect.right, y: rect.top },
    { x: rect.right, y: rect.bottom },
    { x: rect.left, y: rect.bottom },
  ];
}

/**
 * The four corners of `element` as drawn (screen px), when a CSS transform
 * between it and `host` turns or mirrors it; null when nothing does. Turned
 * about its centre, as the editor turns a picture; `scale` is the page's
 * zoom (screen px per CSS px).
 */
function turnedCorners(element: HTMLElement, host: HTMLElement, scale: number): Array<{ x: number; y: number }> | null {
  const win = element.ownerDocument.defaultView;
  if (!win) return null;
  let turned: DOMMatrix | null = null;
  for (let at: HTMLElement | null = element; at && at !== host.parentElement; at = at.parentElement) {
    const transform = win.getComputedStyle(at).transform;
    if (transform && transform !== "none") turned = new DOMMatrix(transform).multiply(turned ?? new DOMMatrix());
  }
  if (!turned) return null;
  const m = turned;
  const rect = element.getBoundingClientRect();
  const cx = rect.left + rect.width / 2;
  const cy = rect.top + rect.height / 2;
  const hw = element.offsetWidth / 2;
  const hh = element.offsetHeight / 2;
  return [
    [-hw, -hh],
    [hw, -hh],
    [hw, hh],
    [-hw, hh],
  ].map(([x = 0, y = 0]) => ({ x: cx + scale * (m.a * x + m.c * y), y: cy + scale * (m.b * x + m.d * y) }));
}

/** What `element` paints, as pieces of a mask of `band`. */
function objectPaintLayers(view: EditorView, element: HTMLElement, frame: BandFrame): PaintLayer[] {
  const win = element.ownerDocument.defaultView;
  if (!win) return [];
  if (element.matches("figure[data-document-shape]")) {
    const paint = element.querySelector<HTMLElement>("[data-document-shape-paint]");
    if (!paint) return [];
    const node = shapeNodeAt(view, element);
    const id = node ? objectId(node) : objectId(element);
    // Every drawing the shape shows -- its own, or each member's of a group
    // -- placed by its own screen matrix (the turn, the mirror, the group's
    // and the member's transforms, the zoom all in it).
    const svgs = Array.from(paint.querySelectorAll<SVGSVGElement>("svg")).filter(
      (svg) => !svg.parentElement?.closest("svg") && svg.getBoundingClientRect().width > 0,
    );
    const toBand = new DOMMatrix().scale(1 / frame.scaleX, 1 / frame.scaleY).translate(-frame.left, -frame.top);
    return svgs.flatMap((svg, index): PaintLayer[] => {
      const style = win.getComputedStyle(svg);
      if (style.display === "none" || style.visibility === "hidden") return [];
      const screen = svg.getScreenCTM();
      if (!screen) return [];
      const m = toBand.multiply(DOMMatrix.fromMatrix(screen));
      const matrix = [m.a, m.b, m.c, m.d, m.e, m.f].map(num).join(" ");
      return [
        {
          kind: "svg",
          key: `s${id}.${index}@${matrix}`,
          markup: () => {
            let byIndex = node ? drawingMarkups.get(node) : undefined;
            let inner = byIndex?.get(index);
            if (inner === undefined) {
              inner = drawingMarkup(svg);
              if (node) {
                if (!byIndex) drawingMarkups.set(node, (byIndex = new Map()));
                byIndex.set(index, inner);
              }
            }
            return `<g transform="matrix(${matrix})">${inner}</g>`;
          },
        },
      ];
    });
  }
  if (element.matches(".note-editor-image-resizable")) {
    const img = element.querySelector("img");
    const src = img?.currentSrc || img?.src;
    if (!img || !src || !img.complete || !(img.naturalWidth > 0)) return [];
    const rect = img.getBoundingClientRect();
    if (!(rect.width > 0 && rect.height > 0)) return [];
    const turned = turnedCorners(img, element, frame.scaleX);
    const id = pictureId(src);
    if (turned || !readablePicture(src, element.ownerDocument.location.origin)) {
      return [opaqueQuad(turned ?? rectCorners(rect), frame, `p${id}`)];
    }
    return [{ kind: "image", image: `url("${src.replace(/"/g, "%22")}")`, rect, key: `p${id}` }];
  }
  if (element.matches("div[data-text-box]")) {
    const frameEl = element.querySelector<HTMLElement>(":scope > [data-text-box-frame]");
    if (!frameEl) return [];
    const fill = win.getComputedStyle(frameEl).backgroundColor;
    const alpha = /rgba?\(([^)]+)\)/.exec(fill)?.[1]?.split(/[\s,/]+/).filter(Boolean)[3];
    const opaque = fill.startsWith("rgb(") || (alpha !== undefined && Number(alpha) >= 1);
    if (!opaque) return [];
    return [opaqueQuad(turnedCorners(frameEl, element, frame.scaleX) ?? rectCorners(frameEl.getBoundingClientRect()), frame, `t${objectId(element)}`)];
  }
  if (element.matches("[data-chart-float]")) {
    return [opaqueQuad(rectCorners(element.getBoundingClientRect()), frame, `c${objectId(element)}`)];
  }
  return [];
}

const MASK_PROPERTIES = ["mask-image", "mask-composite", "mask-position", "mask-size", "mask-repeat", "mask-mode"] as const;

/**
 * Mask `band` (a painted header and footer) where `layers` paint: the band
 * whole, minus the union of the drawings (one SVG image in the band's own
 * px) and of the pictures read by their pixels. None when there are none.
 */
function maskBand(band: HTMLElement, layers: readonly PaintLayer[], frame: BandFrame | null, masked: Map<HTMLElement, string>): void {
  if (layers.length === 0 || !frame) {
    if (masked.has(band)) {
      for (const property of MASK_PROPERTIES) band.style.removeProperty(property);
      masked.delete(band);
    }
    return;
  }
  const px = (value: number) => `${num(value)}px`;
  const images = layers.flatMap((layer) => (layer.kind === "image" ? [layer] : []));
  const drawings = layers.flatMap((layer) => (layer.kind === "svg" ? [layer] : []));
  const width = band.offsetWidth;
  const height = band.offsetHeight;
  const placed = images.map((layer) => ({
    image: layer.image,
    position: `${px((layer.rect.left - frame.left) / frame.scaleX)} ${px((layer.rect.top - frame.top) / frame.scaleY)}`,
    size: `${px(layer.rect.width / frame.scaleX)} ${px(layer.rect.height / frame.scaleY)}`,
  }));
  const signature = [
    `${width}x${height}`,
    ...drawings.map((layer) => layer.key),
    ...images.map((layer, index) => `${layer.key}@${placed[index]!.position}/${placed[index]!.size}`),
  ].join("|");
  if (masked.get(band) === signature && band.style.getPropertyValue("mask-image")) return;
  const all: Array<{ image: string; position: string; size: string }> = [];
  if (drawings.length > 0) {
    const svg =
      `<svg xmlns="${SVG_NS}" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" overflow="visible">` +
      drawings.map((layer) => layer.markup()).join("") +
      "</svg>";
    all.push({ image: `url("data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}")`, position: "0px 0px", size: `${width}px ${height}px` });
  }
  all.push(...placed);
  band.style.setProperty("mask-image", [OPAQUE, ...all.map((layer) => layer.image)].join(", "));
  band.style.setProperty("mask-composite", ["exclude", ...all.map(() => "add")].join(", "));
  band.style.setProperty("mask-position", ["0px 0px", ...all.map((layer) => layer.position)].join(", "));
  band.style.setProperty("mask-size", ["100% 100%", ...all.map((layer) => layer.size)].join(", "));
  band.style.setProperty("mask-repeat", "no-repeat");
  band.style.setProperty("mask-mode", "alpha");
  masked.set(band, signature);
}

interface SheetBands {
  readonly strips: HTMLElement[];
  painted: HTMLElement | null;
}

const NO_BANDS: SheetBands = { strips: [], painted: null };
const STRIP_TESTID = /^sheet-band-hit-(?:header|footer)-(\d+)$/;

/**
 * Every page's hit strips and painted band, read from the two band layers
 * (direct children of the content columns) only: a query of the columns
 * themselves walked the whole document's DOM once per page with an object
 * over its band.
 */
function bandsBySheetOf(columns: HTMLElement): Map<number, SheetBands> {
  const out = new Map<number, SheetBands>();
  const at = (index: number) => {
    let bands = out.get(index);
    if (!bands) out.set(index, (bands = { strips: [], painted: null }));
    return bands;
  };
  let editLayers: Element[] = [];
  let overlays: Element[] = [];
  for (const child of Array.from(columns.children)) {
    if (child.classList.contains("sheet-band-edit-layer")) editLayers.push(child);
    else if (child.classList.contains("sheet-band-overlay")) overlays.push(child);
  }
  // Mounted deeper by some surface: one search of the columns, not one per page.
  if (editLayers.length === 0) editLayers = Array.from(columns.querySelectorAll(".sheet-band-edit-layer"));
  if (overlays.length === 0) overlays = Array.from(columns.querySelectorAll(".sheet-band-overlay"));
  for (const layer of editLayers) {
    for (const strip of Array.from(layer.querySelectorAll<HTMLElement>("[data-testid^='sheet-band-hit-']"))) {
      const index = STRIP_TESTID.exec(strip.dataset.testid ?? "")?.[1];
      if (index !== undefined) at(Number(index)).strips.push(strip);
    }
  }
  for (const overlay of overlays) {
    for (const band of Array.from(overlay.children)) {
      const index = band instanceof HTMLElement ? band.dataset.bandSheetIndex : undefined;
      if (index === undefined || !/^\d+$/.test(index)) continue;
      const bands = at(Number(index));
      bands.painted ??= band as HTMLElement;
    }
  }
  return out;
}

/**
 * Cut the header and footer strips of `view`'s pages around the body objects
 * over them, and mask the painted header and footer where those objects
 * paint; undo both where no object is over them any more.
 */
export function cutBandsAroundBodyObjects(view: EditorView, holes: BandHoles): void {
  const columns = view.dom.closest<HTMLElement>(".note-editor-content-columns");
  const next = new Map<HTMLElement, DOMRect[]>();
  const paints = new Map<HTMLElement, PaintLayer[]>();
  const frames = new Map<HTMLElement, BandFrame>();
  if (columns && holes.intruders.length > 0) {
    const bandsBySheet = bandsBySheetOf(columns);
    const bandsOf = (index: number) => bandsBySheet.get(index) ?? NO_BANDS;
    for (const intruder of holes.intruders) {
      if (!intruder.element.isConnected) continue;
      const bands = bandsOf(intruder.sheetIndex);
      if (bands.strips.length > 0) {
        const rects = objectRects(intruder.element);
        for (const strip of bands.strips) {
          const stripRect = strip.getBoundingClientRect();
          const over = rects.filter((rect) => intersects(rect, stripRect));
          if (over.length > 0) next.set(strip, [...(next.get(strip) ?? []), ...over]);
        }
      }
      if (bands.painted) {
        let frame = frames.get(bands.painted);
        if (!frame) frames.set(bands.painted, (frame = bandFrameOf(bands.painted)));
        const layers = objectPaintLayers(view, intruder.element, frame);
        if (layers.length > 0) paints.set(bands.painted, [...(paints.get(bands.painted) ?? []), ...layers]);
      }
    }
  }
  for (const [band, clip] of holes.cut) {
    if (next.has(band)) continue;
    if (band.style.clipPath === clip.seen) band.style.clipPath = "";
    holes.cut.delete(band);
  }
  for (const [band, rects] of next) {
    const asked = clipWithHoles(band, rects);
    const held = holes.cut.get(band);
    if (held && held.asked === asked && band.style.clipPath === held.seen) continue;
    band.style.clipPath = asked;
    holes.cut.set(band, { asked, seen: band.style.clipPath });
  }
  for (const band of [...holes.masked.keys()]) if (!paints.has(band)) maskBand(band, [], null, holes.masked);
  for (const [band, layers] of paints) maskBand(band, layers, frames.get(band) ?? null, holes.masked);
}
