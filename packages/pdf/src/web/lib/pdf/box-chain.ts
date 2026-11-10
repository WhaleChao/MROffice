/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Which text boxes are one paragraph flow.
 *
 * Extraction cuts a visual paragraph into separate boxes wherever the layout
 * does — at a column break, around a figure, across a page. Reflowing an edit
 * through its continuation needs to know which boxes those are, and the only
 * evidence is geometry.
 *
 * The inference is deliberately conservative. A wrong link is worse than a
 * missing one: reflow moves text into the linked box, and moving a caption's
 * text into an unrelated paragraph corrupts two things at once. Every link
 * carries the evidence it was made on, and anything ambiguous simply does not
 * link — the cost is only that reflow stops where it should have continued,
 * which the shrink-to-fit fallback already handles.
 */

export interface ChainableBox {
  id: string;
  /** 1-based page. */
  page: number;
  /** Page fractions (0..1), top-left origin. */
  x: number;
  y: number;
  width: number;
  height: number;
  /** Font size as a fraction of page height. */
  fontSize: number;
  text: string;
}

export type ChainLinkKind = "column" | "cross-column" | "cross-page";

export interface ChainLink {
  fromId: string;
  toId: string;
  kind: ChainLinkKind;
  /** 0..1; links below the threshold are never made. */
  confidence: number;
}

export interface BoxChain {
  id: string;
  /** Box ids in reading order. */
  boxIds: string[];
  links: ChainLink[];
}

export interface InferChainsOptions {
  /** Vertical gap up to this many line-heights still reads as one column. */
  maxLineGapRatio?: number;
  /** Minimum confidence for a link to be made at all. */
  minConfidence?: number;
}

/** Confidence below which a link is not worth making. */
export const DEFAULT_MIN_CHAIN_CONFIDENCE = 0.55;

/** A sentence that ended is weak evidence the paragraph continues. */
const TERMINAL_PUNCTUATION = /[.!?。！？:：]\s*["'”’」』)]?\s*$/;

/** A line that ends mid-word or unpunctuated strongly suggests continuation. */
function continuationLikelihood(text: string): number {
  const trimmed = text.trimEnd();
  if (!trimmed) return 0;
  if (/[-‐]$/.test(trimmed)) return 1; // explicit hyphen break
  if (TERMINAL_PUNCTUATION.test(trimmed)) return 0.25;
  if (/[,;，、;]$/.test(trimmed)) return 0.9;
  return 0.7; // ends mid-sentence, unpunctuated
}

function horizontalOverlap(a: ChainableBox, b: ChainableBox): number {
  const left = Math.max(a.x, b.x);
  const right = Math.min(a.x + a.width, b.x + b.width);
  const width = Math.min(a.width, b.width);
  if (width <= 0) return 0;
  return Math.max(0, right - left) / width;
}

function similarFont(a: ChainableBox, b: ChainableBox): number {
  const larger = Math.max(a.fontSize, b.fontSize);
  if (larger <= 0) return 0;
  const diff = Math.abs(a.fontSize - b.fontSize) / larger;
  return diff < 0.12 ? 1 : diff < 0.25 ? 0.5 : 0;
}

/**
 * Header/footer bands: boxes living in the outer 8% of the page never link
 * across anything. A running head looks exactly like a continuation of the
 * page's last line — same font, plausible geometry — and never is.
 */
function inMarginBand(box: ChainableBox): boolean {
  return box.y < 0.08 || box.y + box.height > 0.92;
}

function scoreColumnLink(a: ChainableBox, b: ChainableBox, maxGap: number): number {
  if (b.y <= a.y) return 0;
  const gap = b.y - (a.y + a.height);
  const lineHeight = Math.max(a.fontSize, b.fontSize, 0.008);
  if (gap > lineHeight * maxGap) return 0;
  const overlap = horizontalOverlap(a, b);
  if (overlap < 0.6) return 0;
  const font = similarFont(a, b);
  if (font === 0) return 0;
  const gapScore = gap <= lineHeight ? 1 : 1 - (gap / (lineHeight * maxGap)) * 0.5;
  return 0.5 * overlap + 0.3 * gapScore + 0.2 * font;
}

function scoreCrossColumnLink(a: ChainableBox, b: ChainableBox): number {
  // Next column: clearly to the right, back up near the top of the text area,
  // and `a` sits near the bottom — a paragraph only continues from where a
  // column RAN OUT.
  if (b.x < a.x + a.width * 0.8) return 0;
  if (a.y + a.height < 0.6) return 0;
  if (b.y > 0.45) return 0;
  const font = similarFont(a, b);
  if (font === 0) return 0;
  const widthRatio =
    Math.min(a.width, b.width) / Math.max(a.width, b.width, 1e-6);
  if (widthRatio < 0.7) return 0;
  // Continuation gates MULTIPLICATIVELY. Geometry alone scores high on every
  // adjacent column pair — that is what columns look like — and only the text
  // ending mid-thought says the paragraph actually crosses.
  const geometry = 0.45 * font + 0.55 * widthRatio;
  return geometry * (0.35 + 0.65 * continuationLikelihood(a.text));
}

function scoreCrossPageLink(a: ChainableBox, b: ChainableBox): number {
  if (b.page !== a.page + 1) return 0;
  if (a.y + a.height < 0.7) return 0;
  if (b.y > 0.35) return 0;
  const font = similarFont(a, b);
  if (font === 0) return 0;
  const overlap = horizontalOverlap(a, b);
  if (overlap < 0.5) return 0;
  // Same multiplicative gate as cross-column: every page bottom is adjacent
  // to every next-page top, and only the unfinished sentence is evidence.
  const geometry = 0.4 * font + 0.6 * overlap;
  return geometry * (0.35 + 0.65 * continuationLikelihood(a.text));
}

/**
 * Infer the paragraph chains among `boxes`.
 *
 * Each box takes at most one outgoing and one incoming link — a paragraph is
 * a path, not a tree — and when two candidates compete for the same box the
 * higher confidence wins outright.
 */
export function inferBoxChains(
  boxes: ReadonlyArray<ChainableBox>,
  options: InferChainsOptions = {},
): BoxChain[] {
  const maxGap = options.maxLineGapRatio ?? 1.8;
  const minConfidence = options.minConfidence ?? DEFAULT_MIN_CHAIN_CONFIDENCE;

  const usable = boxes.filter(
    (box) =>
      box.width > 0 &&
      box.height > 0 &&
      box.text.trim().length > 0 &&
      !inMarginBand(box),
  );

  const candidates: ChainLink[] = [];
  for (const a of usable) {
    for (const b of usable) {
      if (a.id === b.id) continue;
      if (b.page === a.page) {
        const column = scoreColumnLink(a, b, maxGap);
        if (column >= minConfidence) {
          candidates.push({
            fromId: a.id,
            toId: b.id,
            kind: "column",
            confidence: column,
          });
          continue;
        }
        const cross = scoreCrossColumnLink(a, b);
        if (cross >= minConfidence) {
          candidates.push({
            fromId: a.id,
            toId: b.id,
            kind: "cross-column",
            confidence: cross,
          });
        }
        continue;
      }
      const page = scoreCrossPageLink(a, b);
      if (page >= minConfidence) {
        candidates.push({
          fromId: a.id,
          toId: b.id,
          kind: "cross-page",
          confidence: page,
        });
      }
    }
  }

  // Greedy by confidence: each box one outgoing, one incoming.
  candidates.sort((x, y) => y.confidence - x.confidence);
  const hasOutgoing = new Set<string>();
  const hasIncoming = new Set<string>();
  const chosen: ChainLink[] = [];
  for (const link of candidates) {
    if (hasOutgoing.has(link.fromId) || hasIncoming.has(link.toId)) continue;
    hasOutgoing.add(link.fromId);
    hasIncoming.add(link.toId);
    chosen.push(link);
  }

  // Walk the paths.
  const next = new Map(chosen.map((link) => [link.fromId, link]));
  const chains: BoxChain[] = [];
  const inChain = new Set<string>();
  for (const box of usable) {
    if (hasIncoming.has(box.id) || inChain.has(box.id)) continue;
    if (!next.has(box.id)) continue;
    const ids: string[] = [box.id];
    const links: ChainLink[] = [];
    let cursor = box.id;
    while (next.has(cursor)) {
      const link = next.get(cursor)!;
      // A cycle cannot happen with one-in/one-out over a finite set that
      // always moves to an unvisited box, but a data bug must not hang us.
      if (ids.includes(link.toId)) break;
      ids.push(link.toId);
      links.push(link);
      cursor = link.toId;
    }
    ids.forEach((id) => inChain.add(id));
    chains.push({ id: `chain-${chains.length}`, boxIds: ids, links });
  }
  return chains;
}

/** The chain containing `boxId`, if any. */
export function chainForBox(
  chains: ReadonlyArray<BoxChain>,
  boxId: string,
): BoxChain | undefined {
  return chains.find((chain) => chain.boxIds.includes(boxId));
}
