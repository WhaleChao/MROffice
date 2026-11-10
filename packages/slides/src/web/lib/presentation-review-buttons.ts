/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The deck editor's Review > Comments buttons as pure transforms: who may
 * author comments on this deck (its comment identity mode), and the slide
 * a comment-list change writes. The ribbon (`PresentationWindow`) and the
 * headless editor commands call the same functions with the person's actor.
 */

import { normalizeComments, type SlideComment } from "@mrnote/core/web/lib/presentation-comments.ts";

/** The fields of a slide the review buttons read and write. */
export interface PresentationReviewSlide {
  id: string;
  comments?: SlideComment[];
  commentAuthority?: unknown;
  [key: string]: unknown;
}

export type PresentationDeckCommentMode =
  | "neutral"
  | "neutral-modern"
  | "modern-editable"
  | "legacy-preserve-only"
  | "blocked-preserve-only"
  | "historical-neutral-preserve-only";

function asRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

const PRESENTATION_COMMENT_UUID_RE =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;
const PRESENTATION_COMMENT_DATETIME_RE =
  /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?$/;
const PRESENTATION_COMMENT_XML_CONTROL_RE =
  /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/;

function hasStrictNeutralModernCommentGraph(
  slides: readonly PresentationReviewSlide[],
): boolean {
  const seen = new Set<string>();
  const authorNames = new Map<string, string>();
  let count = 0;
  const validRecord = (value: unknown): value is Record<string, unknown> => {
    const record = asRecord(value);
    if (!record) return false;
    const id = typeof record.id === "string" ? record.id : "";
    const body = typeof record.body === "string" ? record.body : "";
    const author = typeof record.author === "string" ? record.author : "";
    const authorId = typeof record.authorId === "string" ? record.authorId : "";
    const createdAt =
      typeof record.createdAt === "string" ? record.createdAt : "";
    if (
      !PRESENTATION_COMMENT_UUID_RE.test(id) ||
      seen.has(id) ||
      !body ||
      body.length > 4_000 ||
      PRESENTATION_COMMENT_XML_CONTROL_RE.test(body) ||
      !author ||
      author.length > 120 ||
      PRESENTATION_COMMENT_XML_CONTROL_RE.test(author) ||
      !authorId ||
      authorId.length > 256 ||
      PRESENTATION_COMMENT_XML_CONTROL_RE.test(authorId) ||
      !PRESENTATION_COMMENT_DATETIME_RE.test(createdAt) ||
      Number.isNaN(Date.parse(createdAt))
    ) {
      return false;
    }
    const priorName = authorNames.get(authorId);
    if (priorName !== undefined && priorName !== author) return false;
    authorNames.set(authorId, author);
    seen.add(id);
    count += 1;
    return true;
  };
  for (const slide of slides) {
    const comments = Array.isArray(slide.comments) ? slide.comments : [];
    for (const comment of comments) {
      if (!validRecord(comment)) return false;
      for (const reply of Array.isArray(comment.replies) ? comment.replies : []) {
        if (!validRecord(reply)) return false;
      }
    }
  }
  return count > 0;
}

/**
 * The deck's comment identity: the receipt every slide carries (receipts
 * that disagree block every write), else neutral, neutral-modern for a
 * well-formed modern comment graph, or a preserve-only historical graph.
 */
export function presentationDeckCommentIdentity(
  slides: readonly PresentationReviewSlide[],
): { hasComments: boolean; mode: PresentationDeckCommentMode } {
  let hasComments = false;
  const authorityValues = slides
    .map((slide) => asRecord(slide.commentAuthority))
    .filter((value): value is Record<string, unknown> => value !== null);
  const authorityMode =
    authorityValues.length === slides.length && authorityValues.length > 0
      ? authorityValues.every(
          (value) =>
            JSON.stringify(value) === JSON.stringify(authorityValues[0]),
        )
        ? authorityValues[0].mode
        : "blocked-preserve-only"
      : authorityValues.length > 0
        ? "blocked-preserve-only"
        : null;
  for (const slide of slides) {
    if (normalizeComments(slide.comments).length > 0) hasComments = true;
  }
  const mode: PresentationDeckCommentMode =
    authorityMode === "modern-editable" ||
    authorityMode === "legacy-preserve-only" ||
    authorityMode === "blocked-preserve-only"
      ? authorityMode
      : hasComments
        ? hasStrictNeutralModernCommentGraph(slides)
          ? "neutral-modern"
          : "historical-neutral-preserve-only"
        : "neutral";
  return { hasComments, mode };
}

/** Whether a person may author comments on a deck in this mode. */
export function presentationCommentModeIsAuthorable(mode: PresentationDeckCommentMode): boolean {
  return mode === "neutral" || mode === "neutral-modern" || mode === "modern-editable";
}











export type PresentationCommentAuthoringBlock =
  | "identity-loading"
  | "identity-unavailable"
  | "guest"
  | "legacy-comments"
  | "inconsistent-comments";

export function presentationCommentAuthoringBlock(input: {
  readonly mode: PresentationDeckCommentMode;
  readonly identityState: "guest" | "loading" | "ready" | "unavailable";
  readonly hasActor: boolean;
}): PresentationCommentAuthoringBlock | null {
  if (input.mode === "legacy-preserve-only") return "legacy-comments";
  if (!presentationCommentModeIsAuthorable(input.mode)) return "inconsistent-comments";
  if (input.hasActor) return null;
  if (input.identityState === "guest") return "guest";
  if (input.identityState === "unavailable") return "identity-unavailable";
  return "identity-loading";
}

/** The message key of each reason (`console-notebooks`). */
export const PRESENTATION_COMMENT_AUTHORING_BLOCK_MESSAGE = {
  "identity-loading": "presentation.comments.identityLoading",
  "identity-unavailable": "presentation.comments.identityUnavailable",
  guest: "presentation.comments.guestReadOnly",
  "legacy-comments": "presentation.comments.legacyReadOnly",
  "inconsistent-comments": "presentation.comments.inconsistentReadOnly",
} as const satisfies Readonly<Record<PresentationCommentAuthoringBlock, string>>;

/**
 * The slide with its comment list replaced: a non-empty list is written,
 * an empty one deletes the key so the slide stays byte-clean.
 */
export function slideWithComments<S extends PresentationReviewSlide>(
  slide: S,
  next: SlideComment[],
): S {
  if (next.length > 0) return { ...slide, comments: next };
  const cleared = { ...slide };
  delete cleared.comments;
  return cleared;
}
