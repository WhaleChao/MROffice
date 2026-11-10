/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import {
  BUILTIN_LAYOUT_FRAMES,
  LAYOUT_DEFINITIONS,
  layoutBlocks,
  layoutTextStyleAfterTemplate,
  nativeLayoutCanApply,
  nativeLayoutRef,
  normalizeSlideLayoutRef,
  parseNativeLayoutCascadeReceipt,
  type LayoutFrameSource,
  type LayoutId,
} from "@mrnote/core/web/lib/presentation-layouts.ts";
import {
  officeMasterLayoutRef,
  parseOfficeMasterLayoutRef,
} from "@mrnote/core/web/lib/presentation-office-master-model.ts";
import {
  presentationSlideHasNativePptxSource,
  slideWithLayout,
  type PresentationButtonSlide,
} from "@mrnote/core/web/lib/presentation-slide-buttons.ts";
import { computeAutofitScale } from "@mrnote/core/web/lib/presentation-autofit.ts";
import { clampCropRect, type CropRect } from "@mrnote/core/web/lib/presentation-crop-group.ts";

export const DESIGNER_MIN_IDEAS = 3;
export const DESIGNER_MAX_IDEAS = 8;
/** Designer's photo limit (Microsoft Support, "Create professional slide layouts with Designer"). */
export const DESIGNER_MAX_PICTURES = 4;
/** A picture must be larger than 200 x 200 pixels for photo ideas. */
export const DESIGNER_MIN_PICTURE_PX = 200;

const EMU_PER_PX = 9_525;
const MIN_TEXT_FIT = 0.6;

export interface DesignerRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** What Designer knows about a picture after decoding it (all optional). */
export interface DesignerPictureFacts {
  /** Natural pixel size. */
  readonly width: number;
  readonly height: number;
  /** Where the subject sits, as 0..1 fractions of the whole source image. */
  readonly focus?: { readonly x: number; readonly y: number };
  /** The picture's dominant, accent-worthy colour (`#rrggbb`). */
  readonly accent?: string;
}

export interface DesignerContext {
  readonly bounds: { readonly width: number; readonly height: number };
  /** The deck's layout frames (built-in or an Office theme's). */
  readonly layoutSource?: LayoutFrameSource;
  /** Theme accents in order (Accent1, Accent2, ...), `#rrggbb`. */
  readonly accents?: readonly string[];
  /** Header & Footer fields drawn on this slide reserve the bottom band. */
  readonly footerBand?: boolean;
  /** Decoded pictures by block id. */
  readonly pictures?: ReadonlyMap<string, DesignerPictureFacts>;
}

export type DesignerEmptyReason =
  | "empty"
  | "drawn-objects"
  | "too-many-pictures"
  | "small-picture"
  | "fixed-objects"
  | "few-ideas";

export type DesignerIdeaKind =
  | "layout"
  | "theme"
  | "smartart"
  | "timeline"
  | "icons"
  | "title-band"
  | "side-panel"
  | "centered"
  | "two-column"
  | "title-column"
  | "statement"
  | "title-slide-band"
  | "title-slide-panel"
  | "title-slide-bottom"
  | "picture-left"
  | "picture-right"
  | "picture-top"
  | "picture-full"
  | "picture-framed"
  | "picture-color-panel"
  | "picture-grid"
  | "picture-mosaic"
  | "visual-focus"
  | "visual-right"
  | "visual-left"
  | "visual-pair"
  | "stacked"
  | "side-by-side"
  | "centered-stack";

/** What the ranking service is told about an idea: geometry only, no words. */
export interface DesignerIdeaSummary {
  id: string;
  kind: DesignerIdeaKind;
  layout: string | null;
  blocks: Array<{
    role: string;
    type: string;
    left: number;
    top: number;
    width: number;
    height: number;
    text_chars: number;
    filled: boolean;
    cropped: boolean;
  }>;
}

export interface DesignerIdea<S extends PresentationButtonSlide = PresentationButtonSlide> {
  id: string;
  kind: DesignerIdeaKind;
  /** The layout an idea of kind `layout` moves the slide to: the file's name, or a built-in layout's name key. */
  layoutName: string | null;
  layoutNameIsKey: boolean;
  slide: S;
  score: number;
  summary: DesignerIdeaSummary;
}

export interface DesignerResult<S extends PresentationButtonSlide = PresentationButtonSlide> {
  ideas: DesignerIdea<S>[];
  reason: DesignerEmptyReason | null;
}

type Role = "title" | "subtitle" | "body" | "picture" | "visual";

interface Item {
  id: string;
  index: number;
  block: Record<string, unknown>;
  role: Role;
  rect: DesignerRect;
  text: string;
  /** Pixel size of the font the text is set in (null: not text). */
  fontPx: number;
  cjk: boolean;
  attachments: Item[];
}

interface Analysis {
  title: Item | null;
  subtitle: Item | null;
  bodies: Item[];
  pictures: Item[];
  visuals: Item[];
  /** Objects the ideas keep clear of: footers, layout artwork, locked objects. */
  obstacles: DesignerRect[];
  /** Empty scratch placeholders an arrangement drops. */
  emptyPlaceholders: Set<string>;
  native: boolean;
  reason: DesignerEmptyReason | null;
}

function record(value: unknown): Record<string, unknown> | null {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function finite(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function rectOf(block: Record<string, unknown>): DesignerRect | null {
  const left = finite(block.left);
  const top = finite(block.top);
  const width = finite(block.width);
  const height = finite(block.height);
  if (left === null || top === null || width === null || height === null || width <= 0 || height <= 0) {
    return null;
  }
  return { left, top, width, height };
}

function area(rect: DesignerRect): number {
  return Math.max(0, rect.width) * Math.max(0, rect.height);
}

function intersection(a: DesignerRect, b: DesignerRect): number {
  const width = Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left);
  const height = Math.min(a.top + a.height, b.top + b.height) - Math.max(a.top, b.top);
  return width > 0 && height > 0 ? width * height : 0;
}

/** The words a block shows: its text, else its paragraphs'. */
export function designerBlockText(block: Record<string, unknown>): string {
  if (typeof block.text === "string" && block.text.trim()) return block.text;
  const paragraphs = Array.isArray(block.paragraphs) ? block.paragraphs : [];
  return paragraphs
    .map((paragraph) => {
      const text = record(paragraph)?.text;
      return typeof text === "string" ? text : "";
    })
    .join("\n");
}

const CJK = /[　-鿿가-힯＀-￯]/;

function placeholderType(block: Record<string, unknown>): string {
  const value = record(block.source)?.placeholderType;
  return typeof value === "string" ? value : "";
}

function placeholderKey(block: Record<string, unknown>): string {
  const role = block.role;
  if (typeof role === "string" && role) return role;
  const key = block.placeholder;
  return typeof key === "string" ? key : "";
}

const BAND_TYPES = new Set(["dt", "ftr", "hdr", "sldNum"]);
const BAND_KEYS = new Set(["footer", "slideNumber", "date", "dateTime", "header"]);

function isInherited(block: Record<string, unknown>): boolean {
  const from = record(block.source)?.inheritedFrom;
  return from === "layout" || from === "master";
}

function isGrouped(block: Record<string, unknown>): boolean {
  const source = record(block.source);
  if (source && ["groupId", "groupPath", "groupShapeId", "groupParentShapeId", "groupTransform"].some((key) => key in source)) {
    return true;
  }
  return typeof block.groupId === "string" && block.groupId !== "";
}

/** A block's font size in design px (a point-unit block's points x 4/3). */
function fontPxOf(block: Record<string, unknown>, fallbackPt: number): number {
  // Sizes are points on a point-unit block and on an imported one, design px otherwise.
  const points = block.textFontSizeUnit === "point" || typeof record(block.source)?.slideId === "string";
  const px = (size: number) => (points ? (size * 4) / 3 : size);
  const style = record(block.style);
  const size = finite(style?.fontSize);
  if (size !== null && size > 0) return px(size);
  const paragraphs = Array.isArray(block.paragraphs) ? block.paragraphs : [];
  for (const paragraph of paragraphs) {
    const runs = record(paragraph)?.runs;
    if (!Array.isArray(runs)) continue;
    for (const run of runs) {
      const runSize = finite(record(record(run)?.style)?.fontSize ?? record(run)?.fontSize);
      if (runSize !== null && runSize > 0) return px(runSize);
    }
  }
  return (fallbackPt * 4) / 3;
}

/** The runs of a text block with `color` stated on each (an idea's text on an accent fill). */
function recolorRuns(value: unknown, color: string): unknown {
  if (!Array.isArray(value)) return value;
  return value.map((run) => {
    const entry = record(run);
    if (!entry) return run;
    return { ...entry, style: { ...(record(entry.style) ?? {}), color } };
  });
}

/**
 * An idea's text look on a block, where the block keeps it: alignment on the
 * block itself when it states it there (an imported text box), else in its
 * style; a colour on the style and on every run that states its own.
 */
function withTextLook(next: Record<string, unknown>, look: Record<string, unknown>): void {
  const style = { ...(record(next.style) ?? {}) };
  for (const [key, value] of Object.entries(look)) {
    if ((key === "align" || key === "verticalAlign") && key in next) {
      if (value === undefined) delete next[key];
      else next[key] = value;
      continue;
    }
    if (value === undefined) delete style[key];
    else style[key] = value;
    if (key === "color" && typeof value === "string") {
      if (Array.isArray(next.runs)) next.runs = recolorRuns(next.runs, value);
      if (Array.isArray(next.paragraphs)) {
        next.paragraphs = (next.paragraphs as unknown[]).map((paragraph) => {
          const entry = record(paragraph);
          return entry && Array.isArray(entry.runs) ? { ...entry, runs: recolorRuns(entry.runs, value) } : paragraph;
        });
      }
    }
  }
  next.style = style;
}

/**
 * Whether a native object can take new geometry through the existing exact
 * slide/shape patch path (the same authority as dragging it): a direct object
 * of this slide with its own shape id and an untouched stacking order.
 */
function nativeMovable(block: Record<string, unknown>, nativeSlideId: string): boolean {
  if (!nativeSlideId) return true;
  const source = record(block.source);
  return (
    block.type !== "embed" &&
    source?.slideId === nativeSlideId &&
    typeof source.shapeId === "string" &&
    source.shapeId !== "" &&
    Number.isSafeInteger(block.sourceOrder) &&
    block.sourceOrder === block.zIndex
  );
}

function nativeSlideIdOf(slide: PresentationButtonSlide): string {
  const value = record(slide.source)?.slideId;
  return typeof value === "string" ? value.trim() : "";
}

function analyze(slide: PresentationButtonSlide, ctx: DesignerContext): Analysis {
  const { width: W, height: H } = ctx.bounds;
  const nativeSlideId = nativeSlideIdOf(slide);
  const native = presentationSlideHasNativePptxSource(slide);
  const result: Analysis = {
    title: null,
    subtitle: null,
    bodies: [],
    pictures: [],
    visuals: [],
    obstacles: [],
    emptyPlaceholders: new Set(),
    native,
    reason: null,
  };
  const drawn: Item[] = [];
  const make = (block: Record<string, unknown>, index: number, role: Role, rect: DesignerRect): Item => {
    const text = role === "picture" || role === "visual" ? "" : designerBlockText(block);
    return {
      id: String(block.id),
      index,
      block,
      role,
      rect,
      text,
      fontPx: fontPxOf(block, role === "title" ? 36 : role === "subtitle" ? 20 : 18),
      cjk: CJK.test(text),
      attachments: [],
    };
  };
  slide.blocks.forEach((raw, index) => {
    const block = raw as unknown as Record<string, unknown>;
    const rect = rectOf(block);
    if (!rect || typeof block.id !== "string") {
      // An object without geometry is not arranged.
      return;
    }
    if (block.hidden === true) return;
    const type = String(block.type ?? "");
    const ptype = placeholderType(block);
    const key = placeholderKey(block);
    if (isInherited(block)) {
      // The layout's own artwork. A full-slide background is not an obstacle.
      if (area(rect) < W * H * 0.6) result.obstacles.push(rect);
      return;
    }
    if (BAND_TYPES.has(ptype) || BAND_KEYS.has(key)) {
      result.obstacles.push(rect);
      return;
    }
    if (block.locked === true) {
      result.obstacles.push(rect);
      return;
    }
    const text = designerBlockText(block);
    const textBearing = type === "text" || (type === "shape" && text.trim() !== "");
    const isPlaceholder = Boolean(ptype) || key !== "" || typeof block.placeholder === "string";
    let role: Role | null = null;
    if (isPlaceholder && (ptype === "title" || ptype === "ctrTitle" || key === "title")) role = "title";
    else if (isPlaceholder && (ptype === "subTitle" || key === "subtitle")) role = "subtitle";
    else if (type === "image" && block.officeGraphicKind !== "svg") role = "picture";
    else if (type === "chart" || type === "table" || type === "embed") role = "visual";
    else if (isPlaceholder && textBearing) role = "body";
    if ((role === "subtitle" || role === "body") && !text.trim()) {
      // An empty placeholder draws nothing: a scratch one is dropped by an
      // arrangement, a native one stays where it is and is no obstacle. An
      // empty title still takes the title's place in the design, ready for
      // its words.
      if (!native) result.emptyPlaceholders.add(String(block.id));
      return;
    }
    if (role === "picture") {
      // A logo, an icon or a vector graphic is not a photo to arrange: it
      // stays where it is (PowerPoint's photo ideas want pictures larger
      // than 200 x 200 pixels).
      const facts = ctx.pictures?.get(String(block.id));
      const pixelWidth = facts?.width ?? (finite(block.pixelWidth) ?? Infinity);
      const pixelHeight = facts?.height ?? (finite(block.pixelHeight) ?? Infinity);
      const src = typeof block.src === "string" ? block.src : "";
      if (pixelWidth <= DESIGNER_MIN_PICTURE_PX || pixelHeight <= DESIGNER_MIN_PICTURE_PX || src.startsWith("data:image/svg")) {
        result.obstacles.push(rect);
        return;
      }
    }
    if (isGrouped(block) || !nativeMovable(block, nativeSlideId)) {
      // What cannot move stays put; Designer arranges around it.
      result.obstacles.push(rect);
      return;
    }
    if (!role) {
      drawn.push(make(block, index, "body", rect));
      return;
    }
    const item = make(block, index, role, rect);
    if (role === "title") {
      if (result.title) result.bodies.push({ ...item, role: "body" });
      else result.title = item;
    } else if (role === "subtitle") {
      if (result.subtitle) result.bodies.push({ ...item, role: "body" });
      else result.subtitle = item;
    } else if (role === "body") result.bodies.push(item);
    else if (role === "picture") result.pictures.push(item);
    else result.visuals.push(item);
  });
  // Drawn objects never stop Designer: a footnote along the bottom stays put;
  // a callout drawn over a chart or picture travels with it; a text box is
  // content like a placeholder's text (imported decks are mostly text
  // boxes); a drawn line or shape of its own stays where it is.
  const hosts = [...result.pictures, ...result.visuals];
  const textBoxes: Item[] = [];
  for (const item of drawn) {
    const rect = item.rect;
    const isText = item.text.trim() !== "";
    if (isText && rect.top >= H * 0.82 && rect.height <= H * 0.14) {
      result.obstacles.push(rect);
      continue;
    }
    let host: Item | null = null;
    let best = 0;
    for (const candidate of hosts) {
      const overlap = intersection(rect, candidate.rect) / Math.max(1, area(rect));
      if (overlap > best) {
        best = overlap;
        host = candidate;
      }
    }
    // A line has (nearly) no area: it is a callout when its box is inside.
    const inside = area(rect) < 4 ? (host && contains(host.rect, rect) ? 1 : 0) : best;
    if (host && inside >= 0.75) {
      host.attachments.push(item);
      continue;
    }
    if (isText) textBoxes.push(item);
    else result.obstacles.push(rect);
  }
  // Text sitting on the layout's own banner or artwork belongs to it.
  const inherited = [...result.obstacles];
  const onArtwork = (item: Item) =>
    inherited.some((rect) => intersection(item.rect, rect) >= area(item.rect) * 0.6 && area(rect) > area(item.rect));
  for (const item of [...textBoxes]) {
    if (onArtwork(item)) {
      textBoxes.splice(textBoxes.indexOf(item), 1);
      result.obstacles.push(item.rect);
    }
  }
  result.bodies = result.bodies.filter((item) => {
    if (!onArtwork(item)) return true;
    result.obstacles.push(item.rect);
    return false;
  });
  // Text boxes by prominence: with no title placeholder, the largest (then
  // topmost) text in the upper part of the slide is the title; a short line
  // just under the title is the subtitle when the slide has no other text;
  // the rest is body text, in reading order.
  textBoxes.sort((a, b) => a.rect.top - b.rect.top || a.rect.left - b.rect.left);
  if (!result.title?.text.trim() && textBoxes.length) {
    const candidates = textBoxes.filter((item) => item.rect.top < H * 0.6 && item.text.length <= 200);
    const pick = candidates.reduce<Item | null>(
      (best, item) => (!best || item.fontPx > best.fontPx * 1.15 || (item.fontPx >= best.fontPx / 1.15 && item.rect.top < best.rect.top) ? item : best),
      null,
    );
    if (pick) {
      if (result.title) {
        // An empty title placeholder stays where it is.
        result.obstacles.push(result.title.rect);
      }
      result.title = { ...pick, role: "title" };
      textBoxes.splice(textBoxes.indexOf(pick), 1);
    }
  }
  if (result.title && !result.subtitle && result.bodies.length === 0 && hosts.length === 0 && textBoxes.length) {
    const titleBottom = result.title.rect.top + result.title.rect.height;
    const below = textBoxes.find((item) => item.rect.top >= titleBottom - H * 0.05 && item.rect.top <= titleBottom + H * 0.25);
    if (below && below.fontPx <= result.title.fontPx) {
      result.subtitle = { ...below, role: "subtitle" };
      textBoxes.splice(textBoxes.indexOf(below), 1);
    }
  }
  result.bodies.push(...textBoxes.map((item) => ({ ...item, role: "body" as const })));
  result.bodies.sort((a, b) => a.rect.top - b.rect.top || a.rect.left - b.rect.left);
  if (!result.title?.text.trim() && !result.subtitle && result.bodies.length === 0 && hosts.length === 0) {
    result.reason = "empty";
  }
  return result;
}

function contains(outer: DesignerRect, inner: DesignerRect): boolean {
  return (
    inner.left >= outer.left - 1 &&
    inner.top >= outer.top - 1 &&
    inner.left + inner.width <= outer.left + outer.width + 1 &&
    inner.top + inner.height <= outer.top + outer.height + 1
  );
}

/** The largest part of `rect` clear of every obstacle (one cut per obstacle). */
function clearOf(rect: DesignerRect, obstacles: readonly DesignerRect[]): DesignerRect {
  let current = rect;
  for (const obstacle of obstacles) {
    if (intersection(current, obstacle) < area(current) * 0.01) continue;
    const right = current.left + current.width;
    const bottom = current.top + current.height;
    const pieces: DesignerRect[] = [
      { left: current.left, top: current.top, width: obstacle.left - current.left, height: current.height },
      { left: obstacle.left + obstacle.width, top: current.top, width: right - obstacle.left - obstacle.width, height: current.height },
      { left: current.left, top: current.top, width: current.width, height: obstacle.top - current.top },
      { left: current.left, top: obstacle.top + obstacle.height, width: current.width, height: bottom - obstacle.top - obstacle.height },
    ].filter((piece) => piece.width > 0 && piece.height > 0);
    if (pieces.length === 0) return current;
    current = pieces.reduce((best, piece) => (area(piece) > area(best) ? piece : best));
  }
  return current;
}

interface Frame {
  /** Where content may go, clear of footers and the layout's artwork. */
  safe: DesignerRect;
  /** Where a picture may bleed to the slide's edges. */
  bleed: DesignerRect;
  gap: number;
}

function frameFor(analysis: Analysis, ctx: DesignerContext): Frame {
  const { width: W, height: H } = ctx.bounds;
  const obstacles = [...analysis.obstacles];
  if (ctx.footerBand) obstacles.push({ left: 0, top: H * 0.9, width: W, height: H * 0.1 });
  const marginX = Math.round(W * 0.055);
  const marginY = Math.round(H * 0.07);
  const safe = clearOf({ left: marginX, top: marginY, width: W - 2 * marginX, height: H - 2 * marginY }, obstacles);
  // Bleeds reach the edges, stopping above a footer band along the bottom.
  const bottomBand = obstacles
    .filter((rect) => rect.top >= H * 0.75 && rect.width >= W * 0.15)
    .reduce((top, rect) => Math.min(top, rect.top), H);
  const topBand = obstacles
    .filter((rect) => rect.top <= H * 0.02 && rect.top + rect.height <= H * 0.25 && rect.width >= W * 0.5)
    .reduce((bottom, rect) => Math.max(bottom, rect.top + rect.height), 0);
  const bleed = { left: 0, top: topBand, width: W, height: Math.max(1, bottomBand - topBand) };
  return { safe, bleed, gap: Math.round(Math.min(W, H) * 0.035) };
}

interface Placement {
  rect: DesignerRect;
  /** Picture: crop to fill the frame (default) or keep the visible aspect. */
  fit?: "fill" | "contain";
  patch?: Record<string, unknown>;
  style?: Record<string, unknown>;
}

type Plan = Map<string, Placement>;

function inset(rect: DesignerRect, dx: number, dy = dx): DesignerRect {
  return {
    left: rect.left + dx,
    top: rect.top + dy,
    width: Math.max(1, rect.width - 2 * dx),
    height: Math.max(1, rect.height - 2 * dy),
  };
}

function splitRows(rect: DesignerRect, fractions: readonly number[], gap: number): DesignerRect[] {
  const total = fractions.reduce((sum, value) => sum + value, 0) || 1;
  const usable = rect.height - gap * (fractions.length - 1);
  let top = rect.top;
  return fractions.map((fraction) => {
    const height = Math.max(1, Math.round((usable * fraction) / total));
    const row = { left: rect.left, top, width: rect.width, height };
    top += height + gap;
    return row;
  });
}

function splitColumns(rect: DesignerRect, fractions: readonly number[], gap: number): DesignerRect[] {
  const total = fractions.reduce((sum, value) => sum + value, 0) || 1;
  const usable = rect.width - gap * (fractions.length - 1);
  let left = rect.left;
  return fractions.map((fraction) => {
    const width = Math.max(1, Math.round((usable * fraction) / total));
    const column = { left, top: rect.top, width, height: rect.height };
    left += width + gap;
    return column;
  });
}

function fitAspect(slot: DesignerRect, aspect: number): DesignerRect {
  if (!(aspect > 0)) return slot;
  let width = slot.width;
  let height = width / aspect;
  if (height > slot.height) {
    height = slot.height;
    width = height * aspect;
  }
  return {
    left: Math.round(slot.left + (slot.width - width) / 2),
    top: Math.round(slot.top + (slot.height - height) / 2),
    width: Math.max(1, Math.round(width)),
    height: Math.max(1, Math.round(height)),
  };
}

/** Text height in EMU at the item's font size when set in `width` EMU. */
function textHeight(item: Item, width: number): number {
  const fontEmu = item.fontPx * EMU_PER_PX;
  const charEm = item.cjk ? 1 : 0.52;
  const perLine = Math.max(1, Math.floor((width * 0.92) / (fontEmu * charEm)));
  const lines = item.text
    .split("\n")
    .reduce((sum, paragraph) => sum + Math.max(1, Math.ceil(paragraph.length / perLine)), 0);
  return lines * fontEmu * 1.2 + fontEmu * 0.6;
}

function textFit(item: Item, rect: DesignerRect): number {
  return computeAutofitScale({
    text: item.text,
    boxWidth: rect.width * 0.92,
    boxHeight: rect.height * 0.9,
    fontSize: item.fontPx * EMU_PER_PX,
    avgCharEm: item.cjk ? 1 : 0.52,
  });
}

/** A heading box for `item` in a column: tall enough for its lines, within bounds. */
function headingHeight(item: Item, column: DesignerRect, min: number, max: number): number {
  return Math.round(Math.min(column.height * max, Math.max(column.height * min, textHeight(item, column.width))));
}

function textItems(analysis: Analysis): Item[] {
  return [
    ...(analysis.subtitle ? [analysis.subtitle] : []),
    ...analysis.bodies,
  ];
}

/** Title over its text in one column; the text items share what is left. */
function placeTextColumn(
  plan: Plan,
  analysis: Analysis,
  column: DesignerRect,
  gap: number,
  options: { center?: boolean; withTitle?: boolean } = {},
): void {
  const items = textItems(analysis);
  let rest = column;
  if (analysis.title && options.withTitle !== false) {
    const height = items.length ? headingHeight(analysis.title, column, 0.18, 0.4) : column.height;
    plan.set(analysis.title.id, {
      rect: { ...column, height },
      style: options.center ? { align: "center" } : undefined,
    });
    rest = { ...column, top: column.top + height + gap, height: column.height - height - gap };
  }
  if (items.length === 0 || rest.height <= 0) return;
  const weights = items.map((item) => Math.max(1, textHeight(item, rest.width)));
  splitRows(rest, weights, gap).forEach((rect, index) => {
    plan.set(items[index].id, { rect, style: options.center ? { align: "center" } : undefined });
  });
}

function contrastText(hex: string): string {
  const match = /^#?([0-9a-f]{6})$/i.exec(hex);
  if (!match) return "#ffffff";
  const value = parseInt(match[1], 16);
  const channel = (shift: number) => {
    const c = ((value >> shift) & 255) / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  };
  const luminance = 0.2126 * channel(16) + 0.7152 * channel(8) + 0.0722 * channel(0);
  return luminance > 0.45 ? "#1f1f1f" : "#ffffff";
}

function accentAt(ctx: DesignerContext, index: number, fallback?: string): string {
  const accents = (ctx.accents ?? []).filter((value) => /^#[0-9a-f]{6}$/i.test(value));
  if (fallback && /^#[0-9a-f]{6}$/i.test(fallback)) return fallback.toLowerCase();
  return (accents[index % Math.max(1, accents.length)] ?? "#4472c4").toLowerCase();
}

/** The title on an accent fill (only a plain text box carries a fill of its own). */
function accentTitle(plan: Plan, item: Item | null, color: string, extra: Record<string, unknown> = {}): boolean {
  if (!item || item.block.type !== "text") return false;
  const placement = plan.get(item.id);
  if (!placement) return false;
  placement.patch = { ...placement.patch, fill: color };
  placement.style = { ...placement.style, color: contrastText(color), verticalAlign: "middle", ...extra };
  return true;
}

/* ------------------------------------------------------------------------ */
/* Families                                                                  */
/* ------------------------------------------------------------------------ */

interface FamilyInput {
  analysis: Analysis;
  frame: Frame;
  ctx: DesignerContext;
}

type Family = (input: FamilyInput) => Plan | null;

const titleBand: Family = ({ analysis, frame, ctx }) => {
  if (!analysis.title || analysis.title.block.type !== "text") return null;
  const plan: Plan = new Map();
  const { safe, bleed, gap } = frame;
  const band: DesignerRect = {
    left: bleed.left,
    top: bleed.top,
    width: bleed.width,
    height: Math.round(Math.max(safe.top - bleed.top, 0) + headingHeight(analysis.title, safe, 0.12, 0.3)),
  };
  plan.set(analysis.title.id, { rect: band });
  accentTitle(plan, analysis.title, accentAt(ctx, 0));
  const placement = plan.get(analysis.title.id);
  if (placement) placement.patch = { ...placement.patch, textInsets: { left: safe.left - bleed.left, right: safe.left - bleed.left, top: 45720, bottom: 45720 } };
  const below: DesignerRect = { ...safe, top: band.top + band.height + gap, height: safe.top + safe.height - band.top - band.height - gap };
  const items = textItems(analysis);
  const hosts = [...analysis.pictures, ...analysis.visuals];
  if (below.height <= 0) return null;
  if (hosts.length > 0 && items.length > 0) {
    const [text, visual] = splitColumns(below, [2, 3], gap);
    placeTextColumn(plan, { ...analysis, title: null }, text, gap);
    placeHosts(plan, hosts, visual, gap);
  } else if (hosts.length > 0) {
    placeHosts(plan, hosts, below, gap);
  } else {
    placeTextColumn(plan, { ...analysis, title: null }, below, gap);
  }
  return plan;
};

const sidePanel: Family = ({ analysis, frame, ctx }) => {
  if (!analysis.title || analysis.title.block.type !== "text") return null;
  const plan: Plan = new Map();
  const { safe, bleed, gap } = frame;
  const panelWidth = Math.round(bleed.width * 0.34);
  const panel: DesignerRect = { left: bleed.left, top: bleed.top, width: panelWidth, height: bleed.height };
  plan.set(analysis.title.id, { rect: panel });
  accentTitle(plan, analysis.title, accentAt(ctx, 0));
  const pad = safe.left - bleed.left;
  const placement = plan.get(analysis.title.id);
  if (placement) placement.patch = { ...placement.patch, textInsets: { left: pad, right: Math.round(pad / 2), top: pad, bottom: pad } };
  const right: DesignerRect = {
    left: panel.left + panel.width + gap * 1.5,
    top: safe.top,
    width: safe.left + safe.width - panel.left - panel.width - gap * 1.5,
    height: safe.height,
  };
  if (right.width <= 0) return null;
  const hosts = [...analysis.pictures, ...analysis.visuals];
  const items = textItems(analysis);
  if (hosts.length > 0 && items.length > 0) {
    const [text, visual] = splitRows(right, [1, 2], gap);
    placeTextColumn(plan, { ...analysis, title: null }, text, gap);
    placeHosts(plan, hosts, visual, gap);
  } else if (hosts.length > 0) placeHosts(plan, hosts, right, gap);
  else if (items.length > 0) placeTextColumn(plan, { ...analysis, title: null }, right, gap);
  else return null;
  return plan;
};

const centered: Family = ({ analysis, frame }) => {
  if (analysis.pictures.length || analysis.visuals.length) return null;
  const chars = textItems(analysis).reduce((sum, item) => sum + item.text.length, 0);
  if (chars > 320) return null;
  const plan: Plan = new Map();
  const { safe, gap } = frame;
  const column = inset(safe, Math.round(safe.width * 0.1), Math.round(safe.height * 0.06));
  placeTextColumn(plan, analysis, column, gap, { center: true });
  return plan;
};

const twoColumn: Family = ({ analysis, frame }) => {
  if (analysis.pictures.length || analysis.visuals.length) return null;
  const items = textItems(analysis);
  if (items.length < 2) return null;
  const plan: Plan = new Map();
  const { safe, gap } = frame;
  let body = safe;
  if (analysis.title) {
    const height = headingHeight(analysis.title, safe, 0.16, 0.3);
    plan.set(analysis.title.id, { rect: { ...safe, height } });
    body = { ...safe, top: safe.top + height + gap, height: safe.height - height - gap };
  }
  const columns = splitColumns(body, items.map(() => 1), gap * 1.5);
  items.forEach((item, index) => plan.set(item.id, { rect: columns[index] }));
  return plan;
};

const titleColumn: Family = ({ analysis, frame }) => {
  if (!analysis.title) return null;
  const items = textItems(analysis);
  const hosts = [...analysis.pictures, ...analysis.visuals];
  if (items.length === 0 && hosts.length === 0) return null;
  const plan: Plan = new Map();
  const { safe, gap } = frame;
  const [left, right] = splitColumns(safe, [1, 2.1], gap * 1.5);
  plan.set(analysis.title.id, { rect: left, style: { verticalAlign: "middle" } });
  if (hosts.length > 0 && items.length > 0) {
    const [text, visual] = splitRows(right, [1, 2], gap);
    placeTextColumn(plan, { ...analysis, title: null }, text, gap);
    placeHosts(plan, hosts, visual, gap);
  } else if (hosts.length > 0) placeHosts(plan, hosts, right, gap);
  else placeTextColumn(plan, { ...analysis, title: null }, right, gap);
  return plan;
};

const statement: Family = ({ analysis, frame, ctx }) => {
  if (!analysis.title || textItems(analysis).length || analysis.pictures.length || analysis.visuals.length) return null;
  const plan: Plan = new Map();
  const { safe, bleed } = frame;
  const height = headingHeight(analysis.title, safe, 0.3, 0.6);
  plan.set(analysis.title.id, {
    rect: { left: bleed.left, top: Math.round(safe.top + (safe.height - height) / 2), width: bleed.width, height },
    style: { align: "center" },
  });
  accentTitle(plan, analysis.title, accentAt(ctx, 0), { align: "center" });
  return plan;
};

/** Title slide: title and subtitle in one accent band across the middle. */
const titleSlideBand: Family = ({ analysis, frame, ctx }) => {
  if (!analysis.title || analysis.title.block.type !== "text" || analysis.pictures.length || analysis.visuals.length) return null;
  const plan: Plan = new Map();
  const { safe, bleed, gap } = frame;
  const sub = textItems(analysis);
  const titleHeight = headingHeight(analysis.title, safe, 0.24, 0.42);
  const subHeight = sub.length ? Math.round(Math.min(safe.height * 0.3, sub.reduce((sum, item) => sum + textHeight(item, safe.width * 0.8), 0))) : 0;
  const top = Math.round(safe.top + (safe.height - titleHeight - subHeight - gap) / 2);
  plan.set(analysis.title.id, { rect: { left: bleed.left, top, width: bleed.width, height: titleHeight } });
  accentTitle(plan, analysis.title, accentAt(ctx, 0), { align: "center" });
  if (sub.length) {
    const rest: DesignerRect = { left: safe.left + safe.width * 0.1, top: top + titleHeight + gap, width: safe.width * 0.8, height: subHeight };
    splitRows(rest, sub.map(() => 1), gap / 2).forEach((rect, index) => plan.set(sub[index].id, { rect, style: { align: "center" } }));
  }
  return plan;
};

/** Title slide: an accent panel on the left holding the title, the subtitle beside it. */
const titleSlidePanel: Family = ({ analysis, frame, ctx }) => {
  if (!analysis.title || analysis.title.block.type !== "text" || analysis.pictures.length || analysis.visuals.length) return null;
  const plan: Plan = new Map();
  const { safe, bleed, gap } = frame;
  const panel: DesignerRect = { left: bleed.left, top: bleed.top, width: Math.round(bleed.width * 0.55), height: bleed.height };
  plan.set(analysis.title.id, { rect: panel });
  accentTitle(plan, analysis.title, accentAt(ctx, 1));
  const pad = safe.left - bleed.left;
  const placement = plan.get(analysis.title.id);
  if (placement) placement.patch = { ...placement.patch, textInsets: { left: pad, right: pad, top: pad, bottom: pad } };
  const sub = textItems(analysis);
  if (sub.length) {
    const right: DesignerRect = {
      left: panel.left + panel.width + gap * 1.5,
      top: safe.top + safe.height * 0.25,
      width: safe.left + safe.width - panel.left - panel.width - gap * 1.5,
      height: safe.height * 0.5,
    };
    if (right.width <= 0) return null;
    splitRows(right, sub.map(() => 1), gap / 2).forEach((rect, index) => plan.set(sub[index].id, { rect, style: { verticalAlign: "middle" } }));
  }
  return plan;
};

/** Title slide: large title low on the slide over a rule of accent, subtitle under it. */
const titleSlideBottom: Family = ({ analysis, frame, ctx }) => {
  if (!analysis.title || analysis.pictures.length || analysis.visuals.length) return null;
  const plan: Plan = new Map();
  const { safe, gap } = frame;
  const sub = textItems(analysis);
  const subHeight = sub.length ? Math.round(Math.min(safe.height * 0.22, sub.reduce((sum, item) => sum + textHeight(item, safe.width * 0.7), 0))) : 0;
  const titleHeight = headingHeight(analysis.title, { ...safe, width: safe.width * 0.8 }, 0.25, 0.45);
  const subTop = safe.top + safe.height - subHeight;
  const titleTop = subTop - gap - titleHeight;
  plan.set(analysis.title.id, {
    rect: { left: safe.left, top: titleTop, width: Math.round(safe.width * 0.8), height: titleHeight },
    style: { verticalAlign: "bottom", color: accentAt(ctx, 0) },
  });
  if (sub.length) {
    splitRows({ left: safe.left, top: subTop, width: Math.round(safe.width * 0.7), height: subHeight }, sub.map(() => 1), gap / 2)
      .forEach((rect, index) => plan.set(sub[index].id, { rect }));
  }
  return plan;
};

/** Pictures and visuals arranged in `rect`: one fills it; several tile it. */
function placeHosts(plan: Plan, hosts: readonly Item[], rect: DesignerRect, gap: number): void {
  const count = hosts.length;
  if (count === 0) return;
  let cells: DesignerRect[];
  if (count === 1) cells = [rect];
  else if (count === 2) cells = rect.width >= rect.height ? splitColumns(rect, [1, 1], gap) : splitRows(rect, [1, 1], gap);
  else if (count === 3) {
    const [first, rest] = splitColumns(rect, [1.3, 1], gap);
    cells = [first, ...splitRows(rest, [1, 1], gap)];
  } else {
    // An even grid, as many columns as the square root asks.
    const columns = Math.ceil(Math.sqrt(count));
    const rows = Math.ceil(count / columns);
    cells = splitRows(rect, Array.from({ length: rows }, () => 1), gap).flatMap((row, rowIndex) => {
      const inRow = Math.min(columns, count - rowIndex * columns);
      return splitColumns(row, Array.from({ length: inRow }, () => 1), gap);
    });
  }
  hosts.forEach((item, index) => plan.set(item.id, { rect: cells[index] }));
}




/** Everything in reading order down the slide, the title on top. */
const stacked: Family = ({ analysis, frame }) => {
  const plan: Plan = new Map();
  const { safe, gap } = frame;
  const hosts = [...analysis.pictures, ...analysis.visuals];
  const items = textItems(analysis);
  let rest = safe;
  if (analysis.title) {
    const height = headingHeight(analysis.title, safe, 0.12, 0.25);
    plan.set(analysis.title.id, { rect: { ...safe, height } });
    rest = { ...safe, top: safe.top + height + gap, height: safe.height - height - gap };
  }
  const textWeight = items.reduce((sum, item) => sum + textHeight(item, rest.width), 0);
  if (hosts.length && items.length) {
    const share = Math.min(0.5, Math.max(0.2, textWeight / Math.max(1, rest.height)));
    const [text, visual] = splitRows(rest, [share, 1 - share], gap);
    placeTextColumn(plan, { ...analysis, title: null }, text, gap);
    placeHosts(plan, hosts, visual, gap);
  } else if (hosts.length) placeHosts(plan, hosts, rest, gap);
  else if (items.length) placeTextColumn(plan, { ...analysis, title: null }, rest, gap);
  return plan;
};

/** The text in a column on the left, the rest beside it. */
const sideBySide: Family = ({ analysis, frame }) => {
  const plan: Plan = new Map();
  const { safe, gap } = frame;
  const hosts = [...analysis.pictures, ...analysis.visuals];
  const items = textItems(analysis);
  if (!hosts.length && items.length < 2) return null;
  const [left, right] = splitColumns(safe, hosts.length ? [1, 1.4] : [1, 1], gap * 1.5);
  if (hosts.length) {
    placeTextColumn(plan, analysis, left, gap);
    placeHosts(plan, hosts, right, gap);
  } else {
    const half = Math.ceil(items.length / 2);
    placeTextColumn(plan, { ...analysis, bodies: items.slice(0, half), subtitle: null }, left, gap);
    placeTextColumn(plan, { ...analysis, title: null, subtitle: null, bodies: items.slice(half) }, right, gap);
  }
  return plan;
};

/** Everything centred in a calmer, narrower column. */
const centredStack: Family = (input) => {
  const { frame } = input;
  const narrow = inset(frame.safe, Math.round(frame.safe.width * 0.1), Math.round(frame.safe.height * 0.04));
  const plan = stacked({ ...input, frame: { ...frame, safe: narrow } });
  if (!plan) return null;
  for (const [id, placement] of plan) {
    const item = [input.analysis.title, ...textItems(input.analysis)].find((candidate) => candidate?.id === id);
    if (item) placement.style = { ...placement.style, align: "center" };
  }
  return plan;
};

function pictureSide(side: "left" | "right"): Family {
  return ({ analysis, frame }) => {
    if (analysis.pictures.length === 0 || analysis.visuals.length) return null;
    const plan: Plan = new Map();
    const { safe, bleed, gap } = frame;
    const width = Math.round(bleed.width * (analysis.pictures.length > 1 ? 0.55 : 0.47));
    const pictureRect: DesignerRect =
      side === "left"
        ? { left: bleed.left, top: bleed.top, width, height: bleed.height }
        : { left: bleed.left + bleed.width - width, top: bleed.top, width, height: bleed.height };
    placeHosts(plan, analysis.pictures, pictureRect, Math.round(gap / 3));
    const textRect: DesignerRect =
      side === "left"
        ? { left: pictureRect.left + width + gap * 1.5, top: safe.top, width: safe.left + safe.width - pictureRect.left - width - gap * 1.5, height: safe.height }
        : { left: safe.left, top: safe.top, width: pictureRect.left - gap * 1.5 - safe.left, height: safe.height };
    if (textRect.width <= 0) return null;
    if (!analysis.title && textItems(analysis).length === 0) return plan;
    placeTextColumn(plan, analysis, textRect, gap);
    return plan;
  };
}

/** The accent a picture gives (colour extraction, SPEC §1.3 item 2), else the theme's. */
function pictureAccent(analysis: Analysis, ctx: DesignerContext, index: number): string {
  const picture = analysis.pictures[0];
  return accentAt(ctx, index, picture ? ctx.pictures?.get(picture.id)?.accent : undefined);
}

const pictureTop: Family = ({ analysis, frame, ctx }) => {
  if (analysis.pictures.length === 0 || analysis.visuals.length) return null;
  const plan: Plan = new Map();
  const { safe, bleed, gap } = frame;
  const height = Math.round(bleed.height * 0.56);
  const strip: DesignerRect = { left: bleed.left, top: bleed.top, width: bleed.width, height };
  const cells = splitColumns(strip, analysis.pictures.map(() => 1), Math.round(gap / 3));
  analysis.pictures.forEach((item, index) => plan.set(item.id, { rect: cells[index] }));
  const below: DesignerRect = { left: safe.left, top: strip.top + height + gap, width: safe.width, height: safe.top + safe.height - strip.top - height - gap };
  if (below.height <= 0) return null;
  const items = textItems(analysis);
  if (analysis.title && items.length) {
    const [head, text] = splitColumns(below, [1, 1.6], gap * 1.5);
    plan.set(analysis.title.id, { rect: head });
    // The title on a block of the picture's own colour.
    accentTitle(plan, analysis.title, pictureAccent(analysis, ctx, 1));
    placeTextColumn(plan, { ...analysis, title: null }, text, gap);
  } else placeTextColumn(plan, analysis, below, gap);
  return plan;
};

/** One picture bleeding off the right; the words on a panel of the picture's colour. */
const pictureColorPanel: Family = ({ analysis, frame, ctx }) => {
  if (analysis.pictures.length !== 1 || analysis.visuals.length || !analysis.title) return null;
  const words = [analysis.title, ...textItems(analysis)];
  if (words.some((item) => item.block.type !== "text")) return null;
  const plan: Plan = new Map();
  const { safe, bleed } = frame;
  const panelWidth = Math.round(bleed.width * 0.44);
  const panel: DesignerRect = { left: bleed.left, top: bleed.top, width: panelWidth, height: bleed.height };
  plan.set(analysis.pictures[0].id, {
    rect: { left: panel.left + panelWidth, top: bleed.top, width: bleed.width - panelWidth, height: bleed.height },
  });
  const color = pictureAccent(analysis, ctx, 0);
  const pad = Math.max(0, safe.left - bleed.left);
  const insets = { left: pad, right: Math.round(pad / 2), top: Math.round(pad / 2), bottom: Math.round(pad / 2) };
  // The title and the text share one panel: their boxes meet with no gap.
  const items = textItems(analysis);
  const titleShare = items.length
    ? Math.min(0.5, Math.max(0.3, textHeight(analysis.title, panelWidth) / panel.height + 0.12))
    : 1;
  const rows = splitRows(panel, items.length ? [titleShare, 1 - titleShare] : [1], 0);
  plan.set(analysis.title.id, {
    rect: rows[0],
    patch: { fill: color, textInsets: insets },
    style: { color: contrastText(color), verticalAlign: items.length ? "bottom" : "middle" },
  });
  if (items.length) {
    splitRows(rows[1], items.map((item) => Math.max(1, textHeight(item, panelWidth))), 0).forEach((rect, index) =>
      plan.set(items[index].id, { rect, patch: { fill: color, textInsets: insets }, style: { color: contrastText(color) } }),
    );
  }
  return plan;
};

const pictureFull: Family = ({ analysis, frame }) => {
  if (analysis.pictures.length !== 1 || analysis.visuals.length || !analysis.title || analysis.title.block.type !== "text") return null;
  const picture = analysis.pictures[0];
  if (picture.attachments.length) return null;
  const items = textItems(analysis);
  const chars = items.reduce((sum, item) => sum + item.text.length, 0) + analysis.title.text.length;
  if (chars > 220) return null;
  // The picture must already be behind the words it will sit under.
  if ([analysis.title, ...items].some((item) => item.index < picture.index)) return null;
  if (items.some((item) => item.block.type !== "text")) return null;
  const plan: Plan = new Map();
  const { safe, bleed } = frame;
  plan.set(picture.id, { rect: bleed });
  const panel: DesignerRect = { left: safe.left, top: Math.round(safe.top + safe.height * 0.45), width: Math.round(safe.width * 0.6), height: Math.round(safe.height * 0.55) };
  const titleHeight = items.length ? Math.round(panel.height * 0.55) : panel.height;
  plan.set(analysis.title.id, {
    rect: { ...panel, height: titleHeight },
    patch: { fill: "#000000", fillTransparency: 35 },
    style: { color: "#ffffff", verticalAlign: items.length ? "bottom" : "middle" },
  });
  if (items.length) {
    const rest: DesignerRect = { ...panel, top: panel.top + titleHeight, height: panel.height - titleHeight };
    splitRows(rest, items.map(() => 1), 0).forEach((rect, index) =>
      plan.set(items[index].id, { rect, patch: { fill: "#000000", fillTransparency: 35 }, style: { color: "#ffffff" } }),
    );
  }
  return plan;
};

const pictureFramed: Family = ({ analysis, frame }) => {
  if (analysis.pictures.length !== 1 || analysis.visuals.length) return null;
  if (!analysis.title && textItems(analysis).length === 0) return null;
  const plan: Plan = new Map();
  const { safe, gap } = frame;
  const [text, picture] = splitColumns(safe, [1, 1.15], gap * 2);
  plan.set(analysis.pictures[0].id, { rect: picture });
  placeTextColumn(plan, analysis, { ...text, top: text.top + text.height * 0.08, height: text.height * 0.84 }, gap);
  return plan;
};

const pictureGrid: Family = ({ analysis, frame }) => {
  if (analysis.pictures.length < 2 || analysis.visuals.length) return null;
  const plan: Plan = new Map();
  const { safe, gap } = frame;
  let body = safe;
  if (analysis.title) {
    const height = headingHeight(analysis.title, safe, 0.14, 0.24);
    plan.set(analysis.title.id, { rect: { ...safe, height } });
    body = { ...safe, top: safe.top + height + gap, height: safe.height - height - gap };
  }
  const items = textItems(analysis);
  let grid = body;
  if (items.length) {
    const [text, pictures] = splitColumns(body, [1, 2.2], gap * 1.5);
    placeTextColumn(plan, { ...analysis, title: null }, text, gap);
    grid = pictures;
  }
  const count = analysis.pictures.length;
  const cells = count === 4
    ? [...splitColumns(splitRows(grid, [1, 1], gap)[0], [1, 1], gap), ...splitColumns(splitRows(grid, [1, 1], gap)[1], [1, 1], gap)]
    : splitColumns(grid, analysis.pictures.map(() => 1), gap);
  analysis.pictures.forEach((item, index) => plan.set(item.id, { rect: cells[index] }));
  return plan;
};

const pictureMosaic: Family = ({ analysis, frame }) => {
  if (analysis.pictures.length < 2 || analysis.visuals.length) return null;
  const plan: Plan = new Map();
  const { safe, bleed, gap } = frame;
  const width = Math.round(bleed.width * 0.62);
  const mosaic: DesignerRect = { left: bleed.left + bleed.width - width, top: bleed.top, width, height: bleed.height };
  const small = Math.round(gap / 4);
  const [big, rest] = splitColumns(mosaic, [1.4, 1], small);
  const others = analysis.pictures.slice(1);
  plan.set(analysis.pictures[0].id, { rect: big });
  splitRows(rest, others.map(() => 1), small).forEach((rect, index) => plan.set(others[index].id, { rect }));
  const text: DesignerRect = { left: safe.left, top: safe.top, width: mosaic.left - gap * 1.5 - safe.left, height: safe.height };
  if (text.width <= 0) return null;
  if (analysis.title || textItems(analysis).length) placeTextColumn(plan, analysis, text, gap);
  return plan;
};

const visualFocus: Family = ({ analysis, frame }) => {
  if (analysis.visuals.length === 0 && analysis.pictures.length === 0) return null;
  if (analysis.visuals.length === 0) return null;
  const plan: Plan = new Map();
  const { safe, gap } = frame;
  let body = safe;
  if (analysis.title) {
    const height = headingHeight(analysis.title, safe, 0.12, 0.2);
    plan.set(analysis.title.id, { rect: { ...safe, height }, style: { align: "center" } });
    body = { ...safe, top: safe.top + height + gap, height: safe.height - height - gap };
  }
  const items = textItems(analysis);
  if (items.length) {
    const height = Math.round(Math.min(body.height * 0.3, items.reduce((sum, item) => sum + textHeight(item, body.width), 0)));
    const [visual, text] = splitRows(body, [body.height - height - gap, height], gap);
    placeHosts(plan, [...analysis.visuals, ...analysis.pictures], visual, gap);
    placeTextColumn(plan, { ...analysis, title: null }, text, gap);
  } else placeHosts(plan, [...analysis.visuals, ...analysis.pictures], body, gap);
  return plan;
};

function visualSide(side: "left" | "right"): Family {
  return ({ analysis, frame }) => {
    if (analysis.visuals.length === 0) return null;
    if (!analysis.title && textItems(analysis).length === 0) return null;
    const plan: Plan = new Map();
    const { safe, gap } = frame;
    const [a, b] = splitColumns(safe, side === "right" ? [1, 2] : [2, 1], gap * 1.5);
    const text = side === "right" ? a : b;
    const visual = side === "right" ? b : a;
    placeHosts(plan, [...analysis.visuals, ...analysis.pictures], visual, gap);
    placeTextColumn(plan, analysis, { ...text, top: text.top + text.height * 0.1, height: text.height * 0.8 }, gap, {});
    return plan;
  };
}

const visualPair: Family = ({ analysis, frame }) => {
  const hosts = [...analysis.visuals, ...analysis.pictures];
  if (analysis.visuals.length === 0 || hosts.length !== 2) return null;
  const plan: Plan = new Map();
  const { safe, gap } = frame;
  let body = safe;
  if (analysis.title) {
    const height = headingHeight(analysis.title, safe, 0.12, 0.2);
    plan.set(analysis.title.id, { rect: { ...safe, height } });
    body = { ...safe, top: safe.top + height + gap, height: safe.height - height - gap };
  }
  const items = textItems(analysis);
  if (items.length) {
    const height = Math.round(Math.min(body.height * 0.25, items.reduce((sum, item) => sum + textHeight(item, body.width), 0)));
    const [visual, text] = splitRows(body, [body.height - height - gap, height], gap);
    const sideBySide = splitColumns(visual, [1, 1], gap);
    hosts.forEach((item, index) => plan.set(item.id, { rect: sideBySide[index] }));
    placeTextColumn(plan, { ...analysis, title: null }, text, gap);
  } else {
    // Side by side when they are wide enough, else one above the other.
    const cells = body.width / 2 >= body.height * 0.9 ? splitColumns(body, [1, 1], gap) : splitRows(body, [1, 1], gap);
    hosts.forEach((item, index) => plan.set(item.id, { rect: cells[index] }));
  }
  return plan;
};

const FAMILIES: ReadonlyArray<{ kind: DesignerIdeaKind; family: Family; weight: number }> = [
  { kind: "title-slide-band", family: titleSlideBand, weight: 0.95 },
  { kind: "title-slide-panel", family: titleSlidePanel, weight: 0.9 },
  { kind: "title-slide-bottom", family: titleSlideBottom, weight: 0.82 },
  { kind: "picture-left", family: pictureSide("left"), weight: 0.95 },
  { kind: "picture-right", family: pictureSide("right"), weight: 0.93 },
  { kind: "picture-full", family: pictureFull, weight: 0.9 },
  { kind: "picture-color-panel", family: pictureColorPanel, weight: 0.89 },
  { kind: "picture-top", family: pictureTop, weight: 0.86 },
  { kind: "picture-mosaic", family: pictureMosaic, weight: 0.88 },
  { kind: "picture-grid", family: pictureGrid, weight: 0.85 },
  { kind: "picture-framed", family: pictureFramed, weight: 0.8 },
  { kind: "visual-focus", family: visualFocus, weight: 0.9 },
  { kind: "visual-right", family: visualSide("right"), weight: 0.86 },
  { kind: "visual-left", family: visualSide("left"), weight: 0.82 },
  { kind: "visual-pair", family: visualPair, weight: 0.84 },
  { kind: "title-band", family: titleBand, weight: 0.88 },
  { kind: "side-panel", family: sidePanel, weight: 0.85 },
  { kind: "statement", family: statement, weight: 0.9 },
  { kind: "two-column", family: twoColumn, weight: 0.84 },
  { kind: "title-column", family: titleColumn, weight: 0.78 },
  { kind: "centered", family: centered, weight: 0.76 },
  { kind: "stacked", family: stacked, weight: 0.7 },
  { kind: "side-by-side", family: sideBySide, weight: 0.72 },
  { kind: "centered-stack", family: centredStack, weight: 0.66 },
];

function isTitleSlide(slide: PresentationButtonSlide, analysis: Analysis): boolean {
  if (analysis.pictures.length || analysis.visuals.length) return false;
  const layout = normalizeSlideLayoutRef(slide.layout);
  if (layout === "title" || /--title$/.test(String(slide.layout ?? ""))) return true;
  if (analysis.subtitle) return true;
  return false;
}

/* ------------------------------------------------------------------------ */
/* Materializing a plan                                                      */
/* ------------------------------------------------------------------------ */

function clampRect(rect: DesignerRect, bounds: { width: number; height: number }): DesignerRect {
  const width = Math.max(1, Math.min(bounds.width, Math.round(rect.width)));
  const height = Math.max(1, Math.min(bounds.height, Math.round(rect.height)));
  return {
    left: Math.max(0, Math.min(Math.round(rect.left), bounds.width - width)),
    top: Math.max(0, Math.min(Math.round(rect.top), bounds.height - height)),
    width,
    height,
  };
}

function currentCrop(block: Record<string, unknown>): CropRect {
  return clampCropRect(block.cropRect);
}

/** The visible picture's aspect (width / height) before the idea. */
function visibleAspect(item: Item, facts: DesignerPictureFacts | undefined): number {
  const crop = currentCrop(item.block);
  if (facts && facts.width > 0 && facts.height > 0) {
    return (facts.width * crop.w) / (facts.height * crop.h);
  }
  return item.rect.width / item.rect.height;
}

/**
 * Crop → Fill inside what the picture already shows: the kept region has the
 * frame's aspect, centred on the subject when one is known, never wider than
 * the author's own crop.
 */
export function designerFillCrop(
  current: CropRect,
  visible: number,
  frame: number,
  focus?: { x: number; y: number },
): CropRect {
  let w = 1;
  let h = 1;
  if (visible > frame) w = frame / visible;
  else if (visible < frame) h = visible / frame;
  // The focus in the visible region's own fractions.
  const fx = focus ? (focus.x - current.x) / current.w : 0.5;
  const fy = focus ? (focus.y - current.y) / current.h : 0.5;
  const x = Math.min(1 - w, Math.max(0, fx - w / 2));
  const y = Math.min(1 - h, Math.max(0, fy - h / 2));
  const round = (value: number) => Math.round(value * 100000) / 100000;
  return {
    x: round(current.x + x * current.w),
    y: round(current.y + y * current.h),
    w: round(w * current.w),
    h: round(h * current.h),
  };
}

function sameRect(a: DesignerRect, b: DesignerRect): boolean {
  return Math.abs(a.left - b.left) < 2 && Math.abs(a.top - b.top) < 2 && Math.abs(a.width - b.width) < 2 && Math.abs(a.height - b.height) < 2;
}

interface Materialized<S> {
  slide: S;
  fit: number;
  hostArea: number;
  changed: boolean;
}

function materialize<S extends PresentationButtonSlide>(
  slide: S,
  analysis: Analysis,
  plan: Plan,
  ctx: DesignerContext,
  lenient = false,
): Materialized<S> | null {
  const { bounds } = ctx;
  const items = [
    ...(analysis.title ? [analysis.title] : []),
    ...textItems(analysis),
    ...analysis.pictures,
    ...analysis.visuals,
  ];
  // Every piece of content gets a place.
  if (items.some((item) => !plan.has(item.id))) return null;
  const replacement = new Map<string, Record<string, unknown>>();
  let fit = 1;
  let hostArea = 0;
  let changed = false;
  for (const item of items) {
    const placement = plan.get(item.id)!;
    let rect = clampRect(placement.rect, bounds);
    const next: Record<string, unknown> = { ...item.block };
    if (item.role === "picture" || item.role === "visual") {
      const facts = item.role === "picture" ? ctx.pictures?.get(item.id) : undefined;
      if (item.role === "picture" && item.attachments.length === 0 && placement.fit !== "contain") {
        // A picture takes its frame and is cropped to fill it.
        const visible = visibleAspect(item, facts);
        const crop = designerFillCrop(currentCrop(item.block), visible, rect.width / rect.height, facts?.focus);
        next.cropRect = crop;
      } else {
        // Charts, tables and annotated pictures keep their proportions.
        rect = clampRect(fitAspect(rect, item.rect.width / item.rect.height), bounds);
      }
      hostArea += area(rect);
      // Callouts ride along with what they point at.
      const sx = rect.width / item.rect.width;
      const sy = rect.height / item.rect.height;
      for (const attachment of item.attachments) {
        const a = attachment.rect;
        const isText = attachment.text.trim() !== "";
        // A callout keeps its centre over the same point of its object; a
        // label keeps its size, a drawn line or arrow scales with the object.
        const width = isText ? a.width : Math.max(1, a.width * sx);
        const height = isText ? a.height : Math.max(1, a.height * sy);
        const centerX = rect.left + (a.left + a.width / 2 - item.rect.left) * sx;
        const centerY = rect.top + (a.top + a.height / 2 - item.rect.top) * sy;
        const moved = clampRect({ left: centerX - width / 2, top: centerY - height / 2, width, height }, bounds);
        replacement.set(attachment.id, { ...attachment.block, ...moved });
      }
    } else {
      const scale = textFit(item, rect);
      fit = Math.min(fit, scale);
      // Designer treats every placeholder as "Shrink text on overflow".
      next.autofit = "shrink";
      if (placement.style) withTextLook(next, placement.style);
    }
    if (placement.patch) Object.assign(next, placement.patch);
    Object.assign(next, rect);
    if (!sameRect(rect, item.rect) || placement.patch || placement.style || next.cropRect !== item.block.cropRect) changed = true;
    replacement.set(item.id, next);
  }
  const blocks = slide.blocks
    .filter((block) => !analysis.emptyPlaceholders.has(String((block as { id?: unknown }).id)))
    .map((block) => (replacement.get(String((block as { id?: unknown }).id)) ?? block) as PresentationBlock);
  // Placed content must not land on other placed content.
  const placed = items.map((item) => rectOf(replacement.get(item.id)!)!);
  const overlay = items.some((item) => {
    const patch = plan.get(item.id)?.patch;
    return patch && typeof patch.fillTransparency === "number";
  });
  if (!overlay && !lenient) {
    for (let i = 0; i < placed.length; i += 1) {
      for (let j = i + 1; j < placed.length; j += 1) {
        if (intersection(placed[i], placed[j]) > Math.min(area(placed[i]), area(placed[j])) * 0.02) return null;
      }
    }
  }
  return { slide: { ...slide, blocks }, fit, hostArea, changed };
}

/* ------------------------------------------------------------------------ */
/* Layout ideas                                                              */
/* ------------------------------------------------------------------------ */

const LAYOUT_NAME_KEY = new Map(LAYOUT_DEFINITIONS.map((layout) => [layout.id, layout.nameKey]));

interface LayoutIdea<S> {
  slide: S;
  name: string;
  nameIsKey: boolean;
  layout: string;
  fit: number;
  hostArea: number;
}

function builtinLayoutIdeas<S extends PresentationButtonSlide>(
  slide: S,
  analysis: Analysis,
  ctx: DesignerContext,
): LayoutIdea<S>[] {
  const source = ctx.layoutSource ?? BUILTIN_LAYOUT_FRAMES;
  const masterRef = parseOfficeMasterLayoutRef(slide.layout);
  const current = normalizeSlideLayoutRef(slide.layout);
  const currentId = (masterRef?.layoutId ?? current) as LayoutId;
  const previous = layoutBlocks(slide.layout, ctx.bounds, source);
  const hosts = [...analysis.pictures, ...analysis.visuals];
  const texts = textItems(analysis);
  const out: LayoutIdea<S>[] = [];
  for (const definition of LAYOUT_DEFINITIONS) {
    if (definition.id === currentId || definition.id === "blank") continue;
    if (definition.id === "title-vertical-text" || definition.id === "vertical-title-text") continue;
    const ref = masterRef ? officeMasterLayoutRef(masterRef.masterKey, definition.id) : definition.id;
    const templates = layoutBlocks(ref, ctx.bounds, source);
    const titles = templates.filter((template) => template.role === "title");
    const subtitles = templates.filter((template) => template.role === "subtitle");
    const bodies = templates.filter((template) => template.role === "body" || template.role === "content");
    const pictures = templates.filter((template) => template.role === "picture");
    // Every template is filled and every piece of content finds one.
    if (titles.length !== (analysis.title ? 1 : 0)) continue;
    const queueText = [...texts];
    const queueHosts = [...hosts];
    const assignments: Array<{ item: Item; template: (typeof templates)[number] }> = [];
    if (analysis.title) assignments.push({ item: analysis.title, template: titles[0] });
    let ok = true;
    for (const template of subtitles) {
      const item = queueText.shift();
      if (!item) ok = false;
      else assignments.push({ item, template });
    }
    for (const template of pictures) {
      const index = queueHosts.findIndex((item) => item.role === "picture");
      if (index < 0) ok = false;
      else assignments.push({ item: queueHosts.splice(index, 1)[0], template });
    }
    // Content places: text first when the layout keeps text and object
    // side by side (Two Content, Comparison), objects otherwise.
    for (const template of bodies) {
      const item = template.role === "body" ? queueText.shift() ?? queueHosts.shift() : queueHosts.shift() ?? queueText.shift();
      if (!item) ok = false;
      else assignments.push({ item, template });
    }
    if (!ok || queueText.length || queueHosts.length) continue;
    const replacement = new Map<string, Record<string, unknown>>();
    let fit = 1;
    let hostArea = 0;
    for (const { item, template } of assignments) {
      const rect = clampRect(template, ctx.bounds);
      const next: Record<string, unknown> = { ...item.block, placeholder: template.placeholder, role: template.role };
      if (item.role === "picture" || item.role === "visual") {
        let frame = rect;
        if (item.role === "picture" && item.attachments.length === 0) {
          next.cropRect = designerFillCrop(currentCrop(item.block), visibleAspect(item, ctx.pictures?.get(item.id)), rect.width / rect.height, ctx.pictures?.get(item.id)?.focus);
        } else frame = clampRect(fitAspect(rect, item.rect.width / item.rect.height), ctx.bounds);
        Object.assign(next, frame);
        hostArea += area(frame);
      } else {
        const before = previous.find((block) => block.role === template.role)?.style as Record<string, unknown> | undefined;
        const style = layoutTextStyleAfterTemplate(record(item.block.style) ?? undefined, before, template.style as Record<string, unknown> | undefined);
        if (style) next.style = style;
        else delete next.style;
        next.autofit = "shrink";
        Object.assign(next, rect);
        fit = Math.min(fit, textFit(item, rect));
      }
      replacement.set(item.id, next);
    }
    const blocks = slide.blocks
      .filter((block) => !analysis.emptyPlaceholders.has(String((block as { id?: unknown }).id)))
      .map((block) => (replacement.get(String((block as { id?: unknown }).id)) ?? block) as PresentationBlock);
    out.push({
      slide: { ...slide, blocks, layout: ref as S["layout"] },
      name: LAYOUT_NAME_KEY.get(definition.id) ?? "custom",
      nameIsKey: true,
      layout: ref,
      fit,
      hostArea,
    });
  }
  return out;
}

function nativeLayoutIdeas<S extends PresentationButtonSlide>(slide: S): LayoutIdea<S>[] {
  const receipt = parseNativeLayoutCascadeReceipt(record(slide.source)?.layoutCascade);
  if (!receipt) return [];
  const currentPart = typeof slide.layoutId === "string" ? slide.layoutId : receipt.layoutPart;
  const out: LayoutIdea<S>[] = [];
  const seen = new Set<string>();
  for (const entry of receipt.layouts) {
    if (entry.masterPart !== receipt.masterPart || entry.layoutPart === currentPart) continue;
    if (!nativeLayoutCanApply(receipt, entry)) continue;
    const next = slideWithLayout(slide, nativeLayoutRef(entry.layoutPart));
    if (!next || next === slide) continue;
    // Two layouts that place everything the same way are one idea.
    const key = JSON.stringify(next.blocks.map((block) => rectOf(block as unknown as Record<string, unknown>)));
    if (seen.has(key)) continue;
    seen.add(key);
    out.push({
      slide: next,
      name: entry.layoutName?.trim() || "custom",
      nameIsKey: !entry.layoutName?.trim(),
      layout: nativeLayoutRef(entry.layoutPart),
      fit: 1,
      hostArea: 0,
    });
  }
  return out;
}

/* ------------------------------------------------------------------------ */
/* Entry point                                                               */
/* ------------------------------------------------------------------------ */

function summaryOf<S extends PresentationButtonSlide>(
  id: string,
  kind: DesignerIdeaKind,
  layout: string | null,
  slide: S,
  analysis: Analysis,
): DesignerIdeaSummary {
  const roles = new Map<string, Item>();
  for (const item of [...(analysis.title ? [analysis.title] : []), ...textItems(analysis), ...analysis.pictures, ...analysis.visuals]) {
    roles.set(item.id, item);
  }
  const blocks: DesignerIdeaSummary["blocks"] = [];
  for (const raw of slide.blocks) {
    const block = raw as unknown as Record<string, unknown>;
    const item = roles.get(String(block.id));
    const rect = rectOf(block);
    if (!item || !rect) continue;
    blocks.push({
      role: item.role,
      type: String(block.type ?? ""),
      ...rect,
      text_chars: Math.min(20_000, item.text.length),
      filled: typeof block.fill === "string",
      cropped: block.cropRect !== undefined,
    });
  }
  return { id, kind, layout, blocks };
}

/**
 * The Design Ideas for one slide, best first, or none with the reason.
 * Every idea's `slide` is the whole slide as applying it leaves it.
 */
export function designerIdeasForSlide<S extends PresentationButtonSlide>(
  slide: S,
  ctx: DesignerContext,
  /** `minimum`: how few ideas still count (the editor adds graphics ideas before it judges). */
  options: { minimum?: number } = {},
): DesignerResult<S> {
  const analysis = analyze(slide, ctx);
  if (analysis.reason) return { ideas: [], reason: analysis.reason };
  const frame = frameFor(analysis, ctx);
  const ideas: DesignerIdea<S>[] = [];
  const seen = new Set<string>();
  const titleSlide = isTitleSlide(slide, analysis);
  const add = (kind: DesignerIdeaKind, layout: string | null, name: string | null, nameIsKey: boolean, built: { slide: S; fit: number; hostArea: number }, weight: number, lenient = false) => {
    if (!lenient && built.fit < MIN_TEXT_FIT) return;
    const key = JSON.stringify(built.slide.blocks.map((block) => {
      const rec = block as unknown as Record<string, unknown>;
      return [rec.id, rectOf(rec), rec.fill ?? null, rec.cropRect ?? null];
    }));
    if (seen.has(key)) return;
    seen.add(key);
    const slideArea = ctx.bounds.width * ctx.bounds.height;
    const score =
      weight * (0.75 + 0.25 * built.fit) +
      (analysis.pictures.length + analysis.visuals.length > 0 ? 0.25 * Math.min(1, built.hostArea / (slideArea * 0.55)) : 0);
    const id = layout ? `layout:${layout}` : kind;
    ideas.push({
      id,
      kind,
      layoutName: name,
      layoutNameIsKey: nameIsKey,
      slide: built.slide,
      score: Math.round(score * 1000) / 1000,
      summary: summaryOf(id, kind, layout, built.slide, analysis),
    });
  };
  // The current slide itself is not an idea.
  seen.add(JSON.stringify(slide.blocks.filter((block) => !analysis.emptyPlaceholders.has(String((block as { id?: unknown }).id))).map((block) => {
    const rec = block as unknown as Record<string, unknown>;
    return [rec.id, rectOf(rec), rec.fill ?? null, rec.cropRect ?? null];
  })));
  const runFamilies = (lenient: boolean) => {
    for (const { kind, family, weight } of FAMILIES) {
      if (kind.startsWith("title-slide-") && !titleSlide) continue;
      const plan = family({ analysis, frame, ctx });
      if (!plan) continue;
      const built = materialize(slide, analysis, plan, ctx, lenient);
      if (!built || !built.changed) continue;
      // A lenient idea (text that must shrink a lot, boxes that touch) ranks below the rest.
      add(kind, null, null, false, built, lenient ? weight * 0.7 : weight, lenient);
    }
  };
  runFamilies(false);
  const layoutIdeas = analysis.native ? nativeLayoutIdeas(slide) : builtinLayoutIdeas(slide, analysis, ctx);
  for (const idea of layoutIdeas.slice(0, 4)) {
    add("layout", idea.layout, idea.name, idea.nameIsKey, idea, analysis.native ? 0.92 : 0.8);
  }
  // A slide with content always gets ideas: when the strict pass makes too
  // few, the arrangements run again without the fit and overlap limits.
  if (ideas.length < DESIGNER_MIN_IDEAS) runFamilies(true);
  ideas.sort((a, b) => b.score - a.score || a.id.localeCompare(b.id));
  const capped = ideas.slice(0, DESIGNER_MAX_IDEAS);
  if (capped.length === 0 || capped.length < (options.minimum ?? 1)) return { ideas: [], reason: "few-ideas" };
  return { ideas: capped, reason: null };
}

/* ------------------------------------------------------------------------ */
/* Content transforms (SPEC §1.3 items 3–5, 7)                               */
/* ------------------------------------------------------------------------ */

export type DesignerTransformKind = "smartart" | "timeline" | "icons";

/**
 * A slide's text turned into a graphic: the list as a SmartArt diagram, a
 * list of dates as a timeline, key terms with icons. The engine says what to
 * build and where; the editor builds it with its own SmartArt engine (the
 * same path as Home > Convert to SmartArt), so the result is a real,
 * editable diagram that exports as a native one.
 */
export interface DesignerTransform {
  id: string;
  kind: DesignerTransformKind;
  /** The SmartArt catalog layout to build. */
  layoutId: string;
  /** The text block the diagram replaces. */
  targetId: string;
  /** Where the diagram goes (EMU). */
  rect: DesignerRect;
  /** For `icons`: the icon library id for each Level 1 item, in order. */
  icons: string[] | null;
  score: number;
}

const DATE_PATTERN =
  /\b(1[5-9]|20)\d{2}s?\b|\d{4}\s*年|\d{1,2}\s*月|\b(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?\s*\d{0,4}\b|\bq[1-4]\b|第[一二三四]季度/i;
const SEQUENCE_PATTERN =
  /^(step\s*\d+|first(ly)?|second(ly)?|third(ly)?|then|next|after(wards)?|finally|lastly|\d+[.)、]|第[一二三四五六七八九十\d]+[步阶段]|首先|其次|然后|接着|随后|最后|最终)/i;

/** One Level 1 line per outline row (the list Designer reads). */
function listRows(block: Record<string, unknown>): Array<{ text: string; level: number; numbered: boolean }> {
  const paragraphs = Array.isArray(block.paragraphs) ? (block.paragraphs as Array<Record<string, unknown>>) : null;
  if (paragraphs && paragraphs.length) {
    return paragraphs
      .map((paragraph) => ({
        text: String(paragraph.text ?? "").trim(),
        level: Number.isFinite(paragraph.level) ? Number(paragraph.level) : 0,
        numbered: paragraph.bullet === "number",
      }))
      .filter((row) => row.text);
  }
  const text = typeof block.text === "string" ? block.text : "";
  return text
    .split("\n")
    .map((line) => ({ text: line.trim(), level: 0, numbered: false }))
    .filter((row) => row.text);
}

const ICON_STOPWORDS = new Set([
  "the", "and", "for", "with", "our", "your", "their", "from", "into", "this", "that", "are", "was", "will", "can", "all", "new", "more", "most",
]);

function stem(word: string): string {
  return word.replace(/(ies)$/, "y").replace(/(ing|ed|es|s)$/, "");
}

/**
 * The icon from the Insert > Icons library a line of text is about: an icon
 * whose Chinese name appears in it, or whose English keywords include one of
 * its words. Null when nothing matches.
 */
export function designerIconForText(
  text: string,
  catalog: ReadonlyArray<{ id: string; zh: string; keywords: string }>,
): string | null {
  const lower = text.toLowerCase();
  const words = (lower.match(/[a-z]{3,}/g) ?? []).filter((word) => !ICON_STOPWORDS.has(word)).map(stem);
  let best: { id: string; score: number } | null = null;
  for (const entry of catalog) {
    let score = 0;
    for (const keyword of entry.keywords.split(/\s+/)) {
      if (!keyword) continue;
      if (CJK.test(keyword)) {
        if (keyword.length >= 2 && text.includes(keyword)) score = Math.max(score, keyword.length * 2);
      } else if (keyword.length >= 3) {
        const root = stem(keyword);
        // The same word, or one family ("secure" / "security", "global" / "globe").
        const related = words.some((word) => {
          if (word === root) return true;
          const length = Math.min(word.length, root.length, 4);
          return length >= 4 && word.slice(0, length) === root.slice(0, length);
        });
        if (related) score = Math.max(score, 3 + Math.min(4, keyword.length / 3));
      }
    }
    if (score > 0 && (!best || score > best.score)) best = { id: entry.id, score };
  }
  return best?.id ?? null;
}

/**
 * The graphics a slide's text can become, best first. Designer offers them
 * on a Title and Content slide whose one text box holds a list, with no
 * pictures, charts or drawn shapes on it (Microsoft's condition for
 * processes and lists).
 */
export function designerContentTransforms(
  slide: PresentationButtonSlide,
  ctx: DesignerContext,
  iconCatalog: ReadonlyArray<{ id: string; zh: string; keywords: string }> = [],
): DesignerTransform[] {
  const analysis = analyze(slide, ctx);
  if (analysis.reason || analysis.pictures.length || analysis.visuals.length) return [];
  const items = textItems(analysis);
  if (items.length !== 1 || analysis.subtitle) return [];
  const body = items[0];
  if (body.block.type !== "text") return [];
  const rows = listRows(body.block);
  const top = rows.filter((row) => row.level === Math.min(...rows.map((r) => r.level)));
  if (top.length < 2 || top.length > 8 || rows.length > 16) return [];
  if (rows.some((row) => row.text.length > 160)) return [];
  const frame = frameFor(analysis, ctx);
  const titleBottom = analysis.title ? analysis.title.rect.top + analysis.title.rect.height : frame.safe.top;
  const rectTop = Math.max(frame.safe.top, Math.min(body.rect.top, titleBottom + frame.gap));
  const rect: DesignerRect = {
    left: frame.safe.left,
    top: rectTop,
    width: frame.safe.width,
    height: Math.max(1, frame.safe.top + frame.safe.height - rectTop),
  };
  if (rect.height < ctx.bounds.height * 0.3) return [];
  const out: DesignerTransform[] = [];
  const add = (kind: DesignerTransformKind, layoutId: string, score: number, icons: string[] | null = null) =>
    out.push({ id: `${kind}:${layoutId}`, kind, layoutId, targetId: body.id, rect, icons, score });
  const dated = top.filter((row) => DATE_PATTERN.test(row.text)).length;
  const sequential = top.every((row) => row.numbered) || top.filter((row) => SEQUENCE_PATTERN.test(row.text)).length >= Math.max(2, top.length - 1);
  const short = top.every((row) => row.text.length <= 48);
  if (dated >= Math.max(2, Math.ceil(top.length * 0.6))) add("timeline", "timeline", 1.2);
  if (sequential) {
    add("smartart", "basic-process", 1.1);
    add("smartart", short ? "chevron-process" : "vertical-process", 1.05);
  }
  // Key terms with icons (Office's icon layouts, filled from Insert > Icons).
  if (iconCatalog.length) {
    const icons = top.map((row) => designerIconForText(row.text, iconCatalog));
    if (icons.filter(Boolean).length >= Math.max(2, Math.ceil(top.length * 0.75)) && icons.every(Boolean)) {
      add("icons", short && top.length <= 5 ? "icon-circle-label-list" : "icon-vertical-solid-list", 1.15, icons as string[]);
    }
  }
  add("smartart", "basic-list", 1.0);
  add("smartart", rows.length > top.length ? "vertical-bullet-list" : "vertical-block-list", 0.97);
  if (!sequential && top.length <= 6) add("smartart", "horizontal-bullet-list", 0.93);
  return out;
}

/** What the ranking service is told about a graphics idea: the title and the diagram's box. */
export function designerGraphicSummary(transform: DesignerTransform, slide: PresentationButtonSlide): DesignerIdeaSummary {
  const blocks: DesignerIdeaSummary["blocks"] = [];
  for (const raw of slide.blocks) {
    const block = raw as unknown as Record<string, unknown>;
    const rect = rectOf(block);
    if (rect && (block.placeholder === "title" || block.role === "title" || placeholderType(block) === "title")) {
      blocks.push({ role: "title", type: String(block.type), ...rect, text_chars: designerBlockText(block).length, filled: false, cropped: false });
    }
  }
  blocks.push({ role: "visual", type: transform.kind, ...transform.rect, text_chars: 0, filled: false, cropped: false });
  return { id: transform.id, kind: transform.kind, layout: transform.layoutId, blocks };
}

/** What the ranking service is told about a title-slide design: its title boxes as the theme places them. */
export function designerThemeSummary(id: string, slide: PresentationButtonSlide): DesignerIdeaSummary {
  const blocks: DesignerIdeaSummary["blocks"] = [];
  for (const raw of slide.blocks) {
    const block = raw as unknown as Record<string, unknown>;
    const rect = rectOf(block);
    const key = placeholderKey(block);
    if (!rect || (key !== "title" && key !== "subtitle")) continue;
    blocks.push({ role: key, type: String(block.type), ...rect, text_chars: designerBlockText(block).length, filled: false, cropped: false });
  }
  if (blocks.length === 0) {
    const { width, height } = { width: Number(slide.widthEmu) || 12_192_000, height: Number(slide.heightEmu) || 6_858_000 };
    blocks.push({ role: "visual", type: "theme", left: 0, top: 0, width, height, text_chars: 0, filled: true, cropped: false });
  }
  return { id, kind: "theme", layout: id, blocks };
}

/** The slide with `targetId` replaced, where it sat, by the built diagram. */
export function designerSlideWithGraphic<S extends PresentationButtonSlide>(
  slide: S,
  targetId: string,
  nodes: readonly PresentationBlock[],
): S {
  return {
    ...slide,
    blocks: slide.blocks.flatMap((block) => (block.id === targetId ? [...nodes] : [block])),
  };
}

/**
 * Order ideas as the ranking service chose: the ids it ranked first, in its
 * order; the rest after them by score. Ideas it left out drop only while at
 * least {@link DESIGNER_MIN_IDEAS} remain.
 */
export function designerRankedIdeas<T extends { id: string }>(ideas: readonly T[], ranking: readonly string[] | null): T[] {
  if (!ranking || ranking.length === 0) return [...ideas];
  const byId = new Map(ideas.map((idea) => [idea.id, idea]));
  const ranked: T[] = [];
  for (const id of ranking) {
    const idea = byId.get(id);
    if (idea && !ranked.includes(idea)) ranked.push(idea);
  }
  const rest = ideas.filter((idea) => !ranked.includes(idea));
  return ranked.length >= DESIGNER_MIN_IDEAS ? ranked : [...ranked, ...rest.slice(0, DESIGNER_MIN_IDEAS - ranked.length)];
}

/**
 * A cheap identity of what Designer reads on a slide (geometry, words,
 * picture crops and fills), to tell whether ideas still belong to it.
 */
export function designerSlideSignature(slide: { id: string; layout?: unknown; layoutId?: unknown; blocks: readonly unknown[] }): string {
  // Every field of every block counts (an idea is built from all of them),
  // a picture's bytes by their length and tail.
  return JSON.stringify([slide.id, slide.layout ?? null, slide.layoutId ?? null, slide.blocks], (key, value) =>
    (key === "src" || key === "graphicFallbackSrc") && typeof value === "string" && value.length > 96
      ? `${value.length}:${value.slice(0, 32)}:${value.slice(-48)}`
      : value,
  );
}
