/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What PowerPoint's Outline View shows of a slide (audit PPT-E012): its title
 * and the main text of its body placeholder, by paragraph level. Pictures,
 * shapes, charts, tables, media and free text boxes are graphical objects and
 * stay out of the outline; the slide icon only says they are there
 * (Microsoft Support, "Create and print a presentation in Outline view").
 *
 * Pure: derived from the live slide on every render, never stored.
 */

import type {
  PresentationBlock,
  PresentationTextBlock,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { isImportedPptxTextBlock } from "@mrnote/core/web/lib/presentation-shape-buttons.ts";

export type PresentationOutlineTextRole = "title" | "body" | "graphic";

const TITLE_PLACEHOLDERS: ReadonlySet<string> = new Set(["title", "ctrTitle"]);
const BODY_PLACEHOLDERS: ReadonlySet<string> = new Set(["body", "obj", "subTitle"]);

/** A text block's part in the outline. */
export function presentationOutlineTextRole(
  block: PresentationBlock,
): PresentationOutlineTextRole {
  if (block.type !== "text") return "graphic";
  const record = block as PresentationBlock & {
    kind?: unknown;
    placeholder?: unknown;
    placeholderType?: unknown;
    role?: unknown;
  };
  const placeholder = typeof record.placeholder === "string" ? record.placeholder.trim() : "";
  const role = typeof record.role === "string" ? record.role.trim() : "";
  const nativePlaceholder =
    typeof record.placeholderType === "string" ? record.placeholderType.trim() : "";
  if (
    placeholder === "title" ||
    placeholder === "sectionTitle" ||
    role === "title" ||
    TITLE_PLACEHOLDERS.has(nativePlaceholder)
  ) {
    return "title";
  }
  if (nativePlaceholder) {
    return BODY_PLACEHOLDERS.has(nativePlaceholder) ? "body" : "graphic";
  }
  if (placeholder || role) return "body";
  if (record.kind === "textbox") return "graphic";
  // A legacy imported text body with no placeholder authority is a graphic:
  // only a stated placeholder puts imported text in the outline.
  return isImportedPptxTextBlock(block) ? "graphic" : "body";
}

export interface PresentationOutlineProjection {
  readonly title: string;
  readonly titleBlockId: string | null;
  readonly bodyBlockId: string | null;
  /** One entry per body paragraph, empty ones included (they are lines). */
  readonly lines: readonly { readonly text: string; readonly level: number }[];
  /** The body keeps per-paragraph structure, so a line's level can change. */
  readonly bodyHasParagraphs: boolean;
  readonly hasGraphics: boolean;
}

export function presentationOutlineProjection(slide: {
  readonly title: string;
  readonly blocks: readonly PresentationBlock[];
}): PresentationOutlineProjection {
  const titleBlock = slide.blocks.find(
    (block): block is PresentationTextBlock =>
      block.type === "text" && presentationOutlineTextRole(block) === "title",
  );
  const bodyBlock = slide.blocks.find(
    (block): block is PresentationTextBlock =>
      block.type === "text" && presentationOutlineTextRole(block) === "body",
  );
  const paragraphs = Array.isArray(bodyBlock?.paragraphs) ? bodyBlock.paragraphs : null;
  const lines = bodyBlock
    ? paragraphs && paragraphs.length > 0
      ? paragraphs.map((paragraph) => ({
          text: paragraph.text,
          level: Math.min(4, Math.max(0, Math.trunc(Number(paragraph.level) || 0))),
        }))
      : bodyBlock.text.split(/\r?\n/).map((text) => ({ text, level: 0 }))
    : [];
  return {
    title: titleBlock ? titleBlock.text.replace(/\r?\n/g, " ") : slide.title,
    titleBlockId: titleBlock?.id ?? null,
    bodyBlockId: bodyBlock?.id ?? null,
    lines,
    bodyHasParagraphs: Boolean(paragraphs && paragraphs.length > 0),
    hasGraphics: slide.blocks.some(
      (block) => presentationOutlineTextRole(block) === "graphic",
    ),
  };
}







export function presentationSlideDisplayTitle(slide: {
  readonly title?: string;
  readonly blocks?: readonly PresentationBlock[];
}): string {
  const title = presentationOutlineProjection({
    title: typeof slide.title === "string" ? slide.title : "",
    blocks: slide.blocks ?? [],
  }).title;
  return title.replace(/\s+/g, " ").trim();
}
