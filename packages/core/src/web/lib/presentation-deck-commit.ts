/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type {
  PresentationBlock,
  PresentationTextBlock,
} from "../components/console/editor/native-document/documentCommands";
import { parseSlideLinkTarget } from "./presentation-block-fx";
import type { TextRun } from "./presentation-runs";

export interface PresentationCommitSlide {
  id: string;
  blocks: PresentationBlock[];
  commentAuthority?: unknown;
  comments?: unknown;
}

function asRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function withoutMatchingInternalRunLinks(
  runs: readonly TextRun[] | undefined,
  shouldClear: (slideId: string) => boolean,
): { runs: TextRun[] | undefined; changed: boolean } {
  if (!runs) return { runs: undefined, changed: false };
  let changed = false;
  const next = runs.map((run) => {
    const target = parseSlideLinkTarget(run.style?.href);
    if (!target || !shouldClear(target)) return run;
    changed = true;
    const style = { ...run.style };
    delete style.href;
    delete style.linkTooltip;
    delete style.linkColor;
    return Object.keys(style).length > 0
      ? { ...run, style }
      : { text: run.text };
  });
  return { runs: changed ? next : (runs as TextRun[]), changed };
}

/** The object without its Mouse Over link to a slide `shouldClear` names
 * (and the highlight that belongs to it); the same object when it has none. */
function withoutMatchingHoverSlideLink(
  block: PresentationBlock,
  shouldClear: (slideId: string) => boolean,
): PresentationBlock {
  const target = parseSlideLinkTarget((block as { hoverHref?: unknown }).hoverHref);
  if (!target || !shouldClear(target)) return block;
  const next = { ...block } as PresentationBlock & Record<string, unknown>;
  delete next.hoverHref;
  if (next.hoverAction === undefined) delete next.hoverHighlight;
  return next;
}

function withoutMatchingInternalRunLinksInBlock(
  block: PresentationBlock,
  shouldClear: (slideId: string) => boolean,
): PresentationBlock {
  const base = withoutMatchingHoverSlideLink(block, shouldClear);
  if (base.type !== "text" && base.type !== "shape") return base;
  const textBlock = base as PresentationTextBlock;
  const flat = withoutMatchingInternalRunLinks(textBlock.runs, shouldClear);
  let paragraphsChanged = false;
  const paragraphs = textBlock.paragraphs?.map((paragraph) => {
    const result = withoutMatchingInternalRunLinks(paragraph.runs, shouldClear);
    if (!result.changed) return paragraph;
    paragraphsChanged = true;
    const next = { ...paragraph };
    if (result.runs) next.runs = result.runs;
    else delete next.runs;
    return next;
  });
  if (!flat.changed && !paragraphsChanged) return base;
  const next: PresentationTextBlock = { ...textBlock };
  if (flat.runs) next.runs = flat.runs;
  else delete next.runs;
  if (paragraphs) next.paragraphs = paragraphs;
  return next;
}

/**
 * The slides with every text and Mouse Over link to a slide that is not in
 * the deck taken off (a slide the change deleted, or one a paste or a command
 * named from elsewhere): the page refuses a link to no slide of the deck, so
 * one left behind made the whole deck unsavable. A slide the change did not
 * touch was reconciled when it landed, so only a deletion re-reads it.
 */
export function reconcileDeletedSlideRunLinks<S extends PresentationCommitSlide>(
  previousSlides: readonly S[],
  nextSlides: readonly S[],
): S[] {
  const nextIds = new Set(nextSlides.map((slide) => slide.id));
  const deleted = previousSlides.some((slide) => !nextIds.has(slide.id));
  const previousById = new Map(previousSlides.map((slide) => [slide.id, slide]));
  const dangling = (slideId: string) => !nextIds.has(slideId);
  let changed = false;
  const reconciled = nextSlides.map((slide) => {
    if (!deleted && previousById.get(slide.id) === slide) return slide;
    const blocks = slide.blocks.map((block) =>
      withoutMatchingInternalRunLinksInBlock(block, dangling),
    );
    if (!blocks.some((block, index) => block !== slide.blocks[index])) return slide;
    changed = true;
    return { ...slide, blocks };
  });
  return changed ? reconciled : (nextSlides as S[]);
}








export function commentAuthorityText(value: unknown): string {
  const sort = (item: unknown): unknown => {
    if (Array.isArray(item)) return item.map(sort);
    const record = asRecord(item);
    if (!record) return item;
    return Object.fromEntries(Object.keys(record).sort().map((key) => [key, sort(record[key])]));
  };
  return JSON.stringify(sort(value));
}

export function reconcileSlideCommentAuthority<S extends PresentationCommitSlide>(
  previousSlides: readonly S[],
  nextSlides: readonly S[],
): S[] {
  const previousAuthorities = previousSlides
    .map((slide) => asRecord(slide.commentAuthority))
    .filter((value): value is Record<string, unknown> => value !== null);
  // The deck's receipt is the one every slide holds. Two spellings of the
  // same receipt (its keys in another order) are one receipt: read as two,
  // the commit took it off every slide and the server refused the save as a
  // comment change (presentation_comment_mutation_forbidden).
  const spelled = previousAuthorities[0] ? JSON.stringify(previousAuthorities[0]) : null;
  const canonical = previousAuthorities[0] ? commentAuthorityText(previousAuthorities[0]) : null;
  const sameReceipt = (value: Record<string, unknown>) =>
    JSON.stringify(value) === spelled || commentAuthorityText(value) === canonical;
  // The deck's receipt is the one the slides that state one all state. A
  // slide that states none (a stub read before its part, a slide another
  // write left without it) takes it; it is never the reason to take it off
  // every other slide -- that refused the part's save and its room's fold
  
  // receipt"). Slides stating different receipts are left as they are.
  const authority = canonical && previousAuthorities.every(sameReceipt) ? previousAuthorities[0] : null;
  const mixed = previousAuthorities.length > 0 && authority === null;
  // A slide that already carries exactly this authority (or, without one,
  // carries none) is kept as it is: every slide copied at every commit gave
  // every slide a new identity at every keystroke, and every slide's
  
  // 38-slide deck redrew all 38 miniatures at each letter typed).
  return nextSlides.map((slide) => {
    if (mixed) return slide;
    if (authority) {
      const own = asRecord(slide.commentAuthority);
      if (own !== null && JSON.stringify(own) === spelled) return slide;
    } else if (!Object.hasOwn(slide, "commentAuthority")) {
      return slide;
    }
    const next = { ...slide };
    if (authority) {
      next.commentAuthority = JSON.parse(JSON.stringify(authority));
    } else {
      delete next.commentAuthority;
    }
    return next;
  });
}












export function preservesOpaqueCommentStructure(
  previousSlides: readonly PresentationCommitSlide[],
  nextSlides: readonly PresentationCommitSlide[],
): boolean {
  const receipt = previousSlides
    .map((slide) => asRecord(slide.commentAuthority))
    .find((value): value is Record<string, unknown> => value !== null);
  if (receipt?.mode !== "blocked-preserve-only") {
    return true;
  }
  const nextById = new Map(nextSlides.map((slide) => [slide.id, slide]));
  for (const slide of previousSlides) {
    const next = nextById.get(slide.id);
    if (!next || !Array.isArray(slide.comments) || slide.comments.length === 0) {
      continue;
    }
    const nextBlockIds = new Set(next.blocks.map((block) => block.id));
    if (slide.blocks.some((block) => !nextBlockIds.has(block.id))) return false;
  }
  return true;
}
