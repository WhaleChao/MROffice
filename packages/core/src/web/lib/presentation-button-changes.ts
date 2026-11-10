/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { NativeDocumentOperation } from "../components/console/editor/native-document/documentCommands";
import { restoreRedactedReferences } from "./editor/redacted-references";

/** A slide or deck change the editor's buttons made (Copilot
 * run_editor_commands, pressed on the server against this very source). */
export type PresentationButtonChange = Extract<
  NativeDocumentOperation,
  { op: "apply_slide_changes" | "apply_deck_changes" }
>;

interface ButtonDeck {
  slides: readonly { id: string; blocks: readonly { id: string }[] }[];
}

const RESERVED_DECK_KEYS = new Set(["__proto__", "constructor", "prototype"]);
const plainDeckRecord = (value: unknown): value is Record<string, unknown> =>
  Boolean(value) &&
  typeof value === "object" &&
  !Array.isArray(value) &&
  Object.keys(value as object).every((key) => !RESERVED_DECK_KEYS.has(key));

/** One slide or deck change the editor's buttons made, applied to `content`,
 * with the ids of the blocks it changed. Null when it does not fit. The
 * Copilot's projection replays it for the preview, and the page applies it
 * the same way when the change is accepted. */
export function applyPresentationButtonChanges<T extends ButtonDeck>(
  content: T,
  operation: PresentationButtonChange,
): { content: T; touched: string[] } | null {
  if (operation.op === "apply_deck_changes") {
    if (
      (operation.set !== undefined && (!plainDeckRecord(operation.set) || "slides" in operation.set)) ||
      (operation.unset !== undefined && !Array.isArray(operation.unset))
    )
      return null;
    // The deck's own GUIDs come back redacted by the public frame: restored
    // from the deck (lib/editor/redacted-references, AI70H H36R3 C13).
    const deck: Record<string, unknown> = {
      ...content,
      ...restoreRedactedReferences(operation.set ?? {}, content),
    };
    for (const key of operation.unset ?? []) if (key !== "slides") delete deck[key];
    let slides = content.slides;
    if (operation.order !== undefined) {
      const byId = new Map(content.slides.map((slide) => [slide.id, slide]));
      if (
        !Array.isArray(operation.order) ||
        operation.order.length !== content.slides.length ||
        new Set(operation.order).size !== operation.order.length ||
        operation.order.some((id) => !byId.has(id))
      )
        return null;
      slides = operation.order.map((id) => byId.get(id)!);
    }
    return { content: { ...(deck as T), slides }, touched: [] };
  }
  const at = content.slides.findIndex((slide) => slide.id === operation.slide_id);
  if (operation.delete === true) {
    if (at < 0 || content.slides.length <= 1) return null;
    return { content: { ...content, slides: content.slides.filter((_, index) => index !== at) }, touched: [] };
  }
  const written = operation.slide;
  if (!plainDeckRecord(written) || written.id !== operation.slide_id || !Array.isArray(written.blocks)) return null;
  // The slide's own GUIDs (a table's tableStyleId) come back redacted by the
  // public frame: restored from the slide as the page holds it, or the deck
  // refused to save (table_loss … tableStyleId:invalid_guid, AI70H H36R3 C13).
  const next = restoreRedactedReferences(written, at >= 0 ? content.slides[at] : undefined);
  const slides = [...content.slides];
  const slide = next as unknown as T["slides"][number];
  if (at < 0) {
    const index = Number.isSafeInteger(operation.index)
      ? Math.min(Math.max(0, operation.index!), slides.length)
      : slides.length;
    slides.splice(index, 0, slide);
    return { content: { ...content, slides }, touched: slide.blocks.map((block) => block.id) };
  }
  const before = new Map(slides[at].blocks.map((block) => [block.id, JSON.stringify(block)]));
  slides[at] = slide;
  return {
    content: { ...content, slides },
    touched: slide.blocks
      .filter((block) => before.get(block.id) !== JSON.stringify(block))
      .map((block) => block.id),
  };
}

