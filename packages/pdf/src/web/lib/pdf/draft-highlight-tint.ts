/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { HIGHLIGHT_FILL_ALPHA } from "@mrnote/core/web/lib/pdf/editable-annotations.ts";
import { highlightQuadOutline, type HighlightMarkupStyle, type HighlightQuad } from "@mrnote/core/web/lib/pdf/highlight-markup.ts";
import { pdfMarkupColorToHex } from "@mrnote/core/web/lib/pdf/native-text-markup-writeback.ts";

/** On a highlight's target: the colour the wash multiplies by (CSS). */
export const PDF_HIGHLIGHT_WASH_ATTR = "data-pdf-highlight-wash";
/** On a highlight's target: the wash's opacity, 0..1. */
export const PDF_HIGHLIGHT_WASH_OPACITY_ATTR = "data-pdf-highlight-wash-opacity";
/** On a highlight's target: its outline, page percent, "x,y x,y ...". */
export const PDF_HIGHLIGHT_WASH_POINTS_ATTR = "data-pdf-highlight-wash-points";
/** On a copy of a draft's glyphs that carries a highlight over them. */
export const PDF_DRAFT_HIGHLIGHT_TINT_ATTR = "data-pdf-native-text-highlight-tint";
const PROBE_ATTR = "data-pdf-highlight-tint-probe";

export interface PdfHighlightWash {
  /** The colour the wash multiplies by (any CSS colour, alpha included). */
  readonly color: string;
  /** The wash's own opacity, applied over the colour's alpha. */
  readonly opacity: number;
  /** Its outline in page fractions (0..1), one [x, y] per vertex. */
  readonly points: readonly (readonly [number, number])[];
}

const round = (value: number) => Math.round(value * 10_000) / 10_000;

/**
 * What a reader highlight's target in the highlight layer says of the wash
 * the page shows for it; null for a markup that is no wash (underline,
 * strikethrough, squiggly). `drawnByPage`: the page's own drawing carries the
 * mark (its native appearance, the target itself transparent) -- written with
 * the colour's RGB at the highlight fill alpha under /BM /Multiply
 * (editable-annotations `markupAppearance`; the API's managed text markup
 * writes the same /CA). Otherwise the target paints the colour itself,
 * multiplied (`.pdf-viewer__highlight-rect`).
 */
export function pdfHighlightWashAttributes(input: Readonly<{
  color: string;
  markupStyle?: HighlightMarkupStyle;
  drawnByPage: boolean;
  rect: Readonly<{ x: number; y: number; width: number; height: number }>;
  quad?: HighlightQuad | null;
}>): Record<string, string> | null {
  if ((input.markupStyle ?? "highlight") !== "highlight") return null;
  const outline = input.quad
    ? highlightQuadOutline(input.quad).map(({ x, y }) => [x, y] as const)
    : [[input.rect.x, input.rect.y], [input.rect.x + input.rect.width, input.rect.y],
      [input.rect.x + input.rect.width, input.rect.y + input.rect.height], [input.rect.x, input.rect.y + input.rect.height]] as const;
  if (outline.length < 3 || !outline.every(([x, y]) => Number.isFinite(x) && Number.isFinite(y))) return null;
  return {
    [PDF_HIGHLIGHT_WASH_ATTR]: input.drawnByPage ? pdfMarkupColorToHex(input.color) : input.color,
    [PDF_HIGHLIGHT_WASH_OPACITY_ATTR]: String(input.drawnByPage ? HIGHLIGHT_FILL_ALPHA : 1),
    [PDF_HIGHLIGHT_WASH_POINTS_ATTR]: outline.map(([x, y]) => `${round(x)},${round(y)}`).join(" "),
  };
}

/** The washes the page shows (none while its comments are hidden). */
export function readPdfPageHighlightWashes(page: Element): PdfHighlightWash[] {
  if (page.closest(".pdf-viewer")?.getAttribute("data-pdf-comments-hidden") === "true") return [];
  const washes: PdfHighlightWash[] = [];
  for (const target of page.querySelectorAll<HTMLElement>(`[${PDF_HIGHLIGHT_WASH_ATTR}]`)) {
    const color = target.getAttribute(PDF_HIGHLIGHT_WASH_ATTR)?.trim();
    const opacity = Number(target.getAttribute(PDF_HIGHLIGHT_WASH_OPACITY_ATTR) ?? "1");
    const points = (target.getAttribute(PDF_HIGHLIGHT_WASH_POINTS_ATTR) ?? "").trim().split(/\s+/u)
      .map((pair) => pair.split(",").map(Number))
      .filter((pair) => pair.length === 2 && pair.every(Number.isFinite))
      .map(([x, y]) => [x / 100, y / 100] as const);
    if (!color || !(opacity > 0) || opacity > 1 || points.length < 3) continue;
    washes.push({ color, opacity, points });
  }
  return washes;
}

type Point = Readonly<{ x: number; y: number }>;
/** An element's own pixels (before its transforms) to the viewport. */
export type PdfLocalFrame = Readonly<{ origin: Point; ex: Point; ey: Point }>;

const PROBE_REACH = 1000;
/** Where `element`'s own (0, 0), (1000, 0) and (0, 1000) land on screen,
 * read from three empty probes: every transform on the way (the line's turn,
 * a synthetic slant, a zoom) is in it. Null when the element is not laid out. */
function measureLocalFrame(element: HTMLElement): PdfLocalFrame | null {
  const doc = element.ownerDocument;
  const probes = [[0, 0], [PROBE_REACH, 0], [0, PROBE_REACH]].map(([left, top]) => {
    const probe = doc.createElement("span");
    probe.setAttribute(PROBE_ATTR, "true");
    probe.style.cssText = `position:absolute;left:${left}px;top:${top}px;width:0;height:0;margin:0;padding:0;border:0`;
    element.appendChild(probe);
    return probe;
  });
  try {
    const [o, x, y] = probes.map((probe) => probe.getBoundingClientRect());
    const frame = {
      origin: { x: o.left, y: o.top },
      ex: { x: (x.left - o.left) / PROBE_REACH, y: (x.top - o.top) / PROBE_REACH },
      ey: { x: (y.left - o.left) / PROBE_REACH, y: (y.top - o.top) / PROBE_REACH },
    };
    return Math.abs(frame.ex.x * frame.ey.y - frame.ex.y * frame.ey.x) > 1e-9 ? frame : null;
  } finally {
    for (const probe of probes) probe.remove();
  }
}

/** A screen outline as a `clip-path` in the frame's own pixels. */
export function pdfTintClipPath(outline: readonly Point[], frame: PdfLocalFrame): string | null {
  const det = frame.ex.x * frame.ey.y - frame.ex.y * frame.ey.x;
  if (!(Math.abs(det) > 1e-9) || outline.length < 3) return null;
  const local = outline.map(({ x, y }) => {
    const dx = x - frame.origin.x, dy = y - frame.origin.y;
    return { u: (dx * frame.ey.y - dy * frame.ey.x) / det, v: (frame.ex.x * dy - frame.ex.y * dx) / det };
  });
  if (!local.every(({ u, v }) => Number.isFinite(u) && Number.isFinite(v))) return null;
  return `polygon(${local.map(({ u, v }) => `${Math.round(u * 1000) / 1000}px ${Math.round(v * 1000) / 1000}px`).join(", ")})`;
}

function tintsAfter(draft: Element): HTMLElement[] {
  const tints: HTMLElement[] = [];
  for (let next = draft.nextElementSibling; next instanceof HTMLElement && next.hasAttribute(PDF_DRAFT_HIGHLIGHT_TINT_ATTR);
    next = next.nextElementSibling) tints.push(next);
  return tints;
}

/** Shows or hides a draft and the highlights it carries (a static copy of
 * them, such as a finish ghost's, is not kept in step by its watcher). */
export function setPdfDraftVisibility(draft: HTMLElement, visibility: "visible" | "hidden"): void {
  draft.style.visibility = visibility;
  for (const tint of tintsAfter(draft)) tint.style.visibility = visibility;
}

/**
 * Lays the page's highlights over `draft`'s glyphs (see the module comment):
 * one copy per highlight its words reach, right after it, each in the draft's
 * own box, transforms and visibility; none when no highlight reaches them.
 */
export function syncPdfDraftHighlightTints(draft: HTMLElement): void {
  const old = tintsAfter(draft);
  const drop = (tints: readonly HTMLElement[]) => { for (const tint of tints) tint.remove(); };
  const parent = draft.parentElement;
  const page = draft.closest(".pdf-viewer__page");
  const washes = parent && page ? readPdfPageHighlightWashes(page) : [];
  if (!parent || !page || washes.length === 0) { drop(old); return; }
  const pageBox = page.getBoundingClientRect();
  const box = draft.getBoundingClientRect();
  if (!(pageBox.width > 0 && pageBox.height > 0 && box.width > 0 && box.height > 0)) { drop(old); return; }
  // Glyphs reach past the draft's box (overshoots, a superscript): its height around it.
  const reach = box.height;
  const hits = washes.map((wash) => ({
    wash,
    outline: wash.points.map(([x, y]) => ({ x: pageBox.left + x * pageBox.width, y: pageBox.top + y * pageBox.height })),
  })).filter(({ outline }) => {
    const xs = outline.map((point) => point.x), ys = outline.map((point) => point.y);
    return Math.min(...xs) < box.right + reach && Math.max(...xs) > box.left - reach &&
      Math.min(...ys) < box.bottom + reach && Math.max(...ys) > box.top - reach;
  });
  if (hits.length === 0) { drop(old); return; }
  drop(old.slice(hits.length));
  const doc = draft.ownerDocument;
  let frame: PdfLocalFrame | null = null;
  let anchor: Element = draft;
  for (const [index, { wash, outline }] of hits.entries()) {
    const tint = old[index] ?? doc.createElement("div");
    tint.setAttribute(PDF_DRAFT_HIGHLIGHT_TINT_ATTR, "true");
    tint.setAttribute("aria-hidden", "true");
    // The draft's own box, typography, transforms and visibility.
    tint.style.cssText = draft.style.cssText;
    tint.replaceChildren();
    if (anchor.nextSibling !== tint) parent.insertBefore(tint, anchor.nextSibling);
    frame ??= measureLocalFrame(tint);
    const clip = frame ? pdfTintClipPath(outline, frame) : null;
    if (!clip) { tint.remove(); continue; }
    anchor = tint;
    const copy = draft.cloneNode(true) as HTMLElement;
    for (const node of copy.querySelectorAll("[id], [data-testid]")) {
      node.removeAttribute("id");
      node.removeAttribute("data-testid");
    }
    tint.append(...copy.childNodes);
    tint.style.color = wash.color;
    tint.style.setProperty("-webkit-text-fill-color", wash.color);
    tint.style.background = "transparent";
    tint.style.opacity = String(wash.opacity);
    tint.style.mixBlendMode = "multiply";
    tint.style.pointerEvents = "none";
    tint.style.userSelect = "none";
    tint.style.clipPath = clip;
    tint.style.setProperty("-webkit-clip-path", clip);
  }
}

/**
 * Keeps `draft`'s highlights on it while it lives: laid again whenever its
 * words, look or place change (its own subtree, its frame's style). The
 * returned stop takes them away.
 */
export function watchPdfDraftHighlightTints(draft: HTMLElement): () => void {
  const sync = () => {
    try { syncPdfDraftHighlightTints(draft); } catch { /* a highlight never breaks the editor */ }
  };
  const view = draft.ownerDocument.defaultView;
  if (!view?.MutationObserver) return () => undefined;
  const own = new view.MutationObserver(sync);
  own.observe(draft, { attributes: true, childList: true, subtree: true, characterData: true });
  const parent = draft.parentElement;
  const moves = new view.MutationObserver(sync);
  if (parent) moves.observe(parent, { attributes: true, attributeFilter: ["style"] });
  sync();
  return () => {
    own.disconnect();
    moves.disconnect();
    for (const tint of tintsAfter(draft)) tint.remove();
  };
}
