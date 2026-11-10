/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Slide text after a collaborative merge.
 *
 * In the shared room a slide object's `text` is merged by character
 * (lib/collab/json-crdt.ts collaborative text), but its formatting layers —
 * `paragraphs` and block-level `runs` — are whole values: when two people type
 * into one box at once, the room keeps both people's characters in `text` and
 * ONE person's `paragraphs` / `runs`, which spell that person's text only.
 * The canvas draws the formatting layers, so the other person's characters
 * would be in the document but not on the slide.
 *
 * These pure helpers restore the invariant (`text` is what the layers spell)
 * by replaying the difference onto the layers the room kept, exactly as a
 * keystroke is reconciled, so untouched characters keep their formatting and
 * new ones take their neighbour's. They also carry a text caret across a
 * remote change.
 */

import type {
  PresentationBlock,
  PresentationTextBlock,
} from "../components/console/editor/native-document/documentCommands";
import { textSplice } from "./collab/json-crdt";
import { textFromParagraphs } from "./presentation-paragraphs";
import { normalizeRuns, textFromRuns } from "./presentation-runs";
import { reconcilePresentationTextBlock } from "./presentation-shape-buttons";
import { normalizeParagraphsWithRuns } from "./presentation-text-buttons";

type TextCarrier = { type?: unknown; text?: unknown; paragraphs?: unknown; runs?: unknown };

/** The text a block's formatting layers spell; null when it has none. */
function formattedText(block: TextCarrier): string | null {
  const paragraphs = normalizeParagraphsWithRuns(block.paragraphs);
  if (paragraphs) return textFromParagraphs(paragraphs);
  const runs = normalizeRuns(block.runs);
  return runs ? textFromRuns(runs) : null;
}

/**
 * `next` with its formatting layers rebuilt to spell its `text`, when a merge
 * left them behind. Only a change the merge introduced is repaired: the text
 * must have changed since `previous`, and `previous` must itself have been
 * consistent — a source that never satisfied the invariant (an import the
 * editor keeps as it is) is left alone, so a delivery never rewrites it.
 */
export function reconcileCollaborativeBlockText<B extends PresentationBlock>(
  previous: B | undefined,
  next: B,
): B {
  const carrier = next as TextCarrier;
  if ((carrier.type !== "text" && carrier.type !== "shape") || typeof carrier.text !== "string") return next;
  const prior = previous as TextCarrier | undefined;
  if (!prior || prior.text === carrier.text) return next;
  const spelled = formattedText(carrier);
  if (spelled === null || spelled === carrier.text) return next;
  const priorSpelled = formattedText(prior);
  if (priorSpelled !== null && priorSpelled !== prior.text) return next;
  return reconcilePresentationTextBlock(
    { ...(next as unknown as PresentationTextBlock), text: spelled },
    carrier.text,
  ) as unknown as B;
}









export function repairBlockTextLayers<B extends PresentationBlock>(block: B): B {
  const carrier = block as TextCarrier;
  if ((carrier.type !== "text" && carrier.type !== "shape") || typeof carrier.text !== "string") return block;
  const text = carrier.text;
  if (layersSpell(carrier, text)) return block;
  // Rebuilt by replaying the change onto what the layers spell, so untouched
  // characters keep their formatting.
  const spelled = formattedText(carrier) ?? rawLayerText(carrier);
  if (spelled !== null) {
    const rebuilt = reconcilePresentationTextBlock(
      { ...(block as unknown as PresentationTextBlock), text: spelled },
      text,
    ) as unknown as B;
    if (layersSpell(rebuilt as TextCarrier, text)) return rebuilt;
  }
  const { runs: _runs, paragraphs: _paragraphs, ...plain } = block as B & TextCarrier;
  void _runs;
  void _paragraphs;
  return plain as unknown as B;
}









export function repairContentTextLayers<C>(content: C): { content: C; repaired: boolean } {
  const record = content as { slides?: unknown } | null | undefined;
  if (!record || typeof record !== "object" || !Array.isArray(record.slides)) return { content, repaired: false };
  let repaired = false;
  const slides = record.slides.map((slide: unknown) => {
    const blocks = slide && typeof slide === "object" ? (slide as { blocks?: unknown }).blocks : undefined;
    if (!Array.isArray(blocks)) return slide;
    let changed = false;
    const next = blocks.map((block: unknown) => {
      if (!block || typeof block !== "object") return block;
      const fixed = repairBlockTextLayers(block as PresentationBlock);
      if (fixed !== block) changed = true;
      return fixed;
    });
    if (!changed) return slide;
    repaired = true;
    return { ...(slide as object), blocks: next };
  });
  return repaired ? { content: { ...(content as object), slides } as C, repaired } : { content, repaired };
}

/** What the stored layers spell, read as the server reads them. */
function rawLayerText(carrier: TextCarrier): string | null {
  const textOf = (entry: unknown) =>
    entry && typeof entry === "object" && typeof (entry as { text?: unknown }).text === "string"
      ? (entry as { text: string }).text
      : "";
  if (Array.isArray(carrier.paragraphs)) return carrier.paragraphs.map(textOf).join("\n");
  if (Array.isArray(carrier.runs)) return carrier.runs.map(textOf).join("");
  return null;
}

/** The server's rule: every layer present spells `text` exactly. */
function layersSpell(carrier: TextCarrier, text: string): boolean {
  const textOf = (entry: unknown) =>
    entry && typeof entry === "object" && typeof (entry as { text?: unknown }).text === "string"
      ? (entry as { text: string }).text
      : "";
  if (Array.isArray(carrier.runs) && carrier.runs.map(textOf).join("") !== text) return false;
  if (Array.isArray(carrier.paragraphs) && carrier.paragraphs.map(textOf).join("\n") !== text) return false;
  return true;
}

/** Every slide object of `next` through `reconcileCollaborativeBlockText`,
 *  against the same slide/object of `previous`. Identity-preserving. */
export function reconcileCollaborativeSlideText<
  S extends { id: string; blocks: PresentationBlock[] },
>(previous: readonly S[], next: S[]): S[] {
  const previousSlides = new Map(previous.map((slide) => [slide.id, slide]));
  let changed = false;
  const out = next.map((slide) => {
    const before = previousSlides.get(slide.id);
    if (!before) return slide;
    const beforeBlocks = new Map(before.blocks.map((block) => [block.id, block]));
    let slideChanged = false;
    const blocks = slide.blocks.map((block) => {
      const repaired = reconcileCollaborativeBlockText(beforeBlocks.get(block.id), block);
      if (repaired !== block) slideChanged = true;
      return repaired;
    });
    if (!slideChanged) return slide;
    changed = true;
    return { ...slide, blocks };
  });
  return changed ? out : next;
}

/**
 * Where a caret at `offset` in `before` belongs in `after`, for a change made
 * by someone else: before the change it stays, after it shifts by the change's
 * length, and inside a range they replaced it moves to the end of their text.
 */
export function mapCaretThroughTextChange(offset: number, before: string, after: string): number {
  if (before === after) return offset;
  const { start, end, insert } = textSplice(before, after);
  if (offset <= start) return Math.min(offset, after.length);
  if (offset >= end) return offset + insert.length - (end - start);
  return start + insert.length;
}
