/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for PDF annotation comment THREADS (Acrobat / WPS PDF "Reply"):
 * a flat list of replies hanging off a text annotation. Kept dependency-free and
 * impurity-free (the caller injects id / timestamp factories) so the parse +
 * append logic is unit-testable apart from the PDF window's autosave plumbing.
 */

/**
 * Author identity stamped onto an annotation / reply at creation time (the
 * current user). Persisted inside the payload (the annotations API stores raw
 * payload JSON), so a viewer can render "who said this" without a user lookup.
 */
export interface PdfAnnotationAuthor {
  userId: string;
  displayName: string;
}




export interface PdfAnnotationReviewState {
  model: "Review" | "Marked";
  state: string;
}

export interface PdfAnnotationReply {
  id: string;
  text: string;
  createdAt: string;
  /** Who wrote the reply; absent on legacy replies created before authors. */
  author?: PdfAnnotationAuthor;
  /** Present when this "reply" is the author's review state record. */
  reviewState?: PdfAnnotationReviewState;
}

/** Cap a single annotation's reply thread (guards persisted snapshot size). */
export const MAX_REPLIES_PER_ANNOTATION = 500;

/** Cap one reply's text (mirrors the server validator in datasets.py). */
export const MAX_REPLY_TEXT = 2000;

/** Caps for the persisted author identity (mirror the server validator). */
export const MAX_AUTHOR_USER_ID = 64;
export const MAX_AUTHOR_DISPLAY_NAME = 120;

/**
 * Parse a persisted author object from an untrusted payload: both fields must
 * be non-empty strings (after trimming the display name); over-cap values are
 * truncated rather than dropped. Tolerates the snake_case mirrors
 * (`user_id` / `display_name`) a non-client writer might use. Returns
 * undefined for anything else, so legacy payloads stay author-less.
 */
export function parseAnnotationAuthor(
  raw: unknown,
): PdfAnnotationAuthor | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const rec = raw as Record<string, unknown>;
  const userIdRaw = rec.userId ?? rec.user_id;
  const displayNameRaw = rec.displayName ?? rec.display_name;
  if (typeof userIdRaw !== "string" || typeof displayNameRaw !== "string") {
    return undefined;
  }
  const userId = userIdRaw.slice(0, MAX_AUTHOR_USER_ID);
  const displayName = displayNameRaw.trim().slice(0, MAX_AUTHOR_DISPLAY_NAME);
  if (!userId || !displayName) return undefined;
  return { userId, displayName };
}

/**
 * Parse persisted reply objects from an untrusted snapshot: keep only entries
 * with non-empty text (capped to {@link MAX_REPLY_TEXT}), backfill a missing
 * id / timestamp via the injected factories, and cap the count. Tolerates the
 * snake_case `created_at` mirror. A reply's author (if valid) is preserved.
 */
export function parseReplies(
  raw: unknown,
  fallbackId: (index: number) => string,
  fallbackCreatedAt: () => string,
): PdfAnnotationReply[] {
  if (!Array.isArray(raw)) return [];
  const out: PdfAnnotationReply[] = [];
  for (const item of raw) {
    if (!item || typeof item !== "object") continue;
    const rec = item as Record<string, unknown>;
    const text = String(rec.text ?? rec.comment ?? "")
      .trim()
      .slice(0, MAX_REPLY_TEXT);
    if (!text) continue;
    const createdAt =
      (typeof rec.createdAt === "string" && rec.createdAt) ||
      (typeof rec.created_at === "string" && rec.created_at) ||
      fallbackCreatedAt();
    const author = parseAnnotationAuthor(rec.author);
    out.push({
      id: typeof rec.id === "string" && rec.id ? rec.id : fallbackId(out.length),
      text,
      createdAt,
      // Key present only when the reply HAS an author, so a legacy reply
      // round-trips byte-identically.
      ...(author ? { author } : {}),
    });
    if (out.length >= MAX_REPLIES_PER_ANNOTATION) break;
  }
  return out;
}

/**
 * Return a new annotation list with `reply` appended to the thread of the
 * annotation whose id matches `annotationId`. Immutable (never mutates inputs);
 * a blank reply or an unknown id yields an unchanged (cloned) list. Only the
 * FIRST id match is touched, and the thread is capped at the most-recent
 * {@link MAX_REPLIES_PER_ANNOTATION}.
 */
export function appendReply<
  T extends { id: string; replies?: PdfAnnotationReply[] },
>(annotations: readonly T[], annotationId: string, reply: PdfAnnotationReply): T[] {
  if (!reply.text.trim()) return annotations.slice();
  let done = false;
  return annotations.map((annotation) => {
    if (done || annotation.id !== annotationId) return annotation;
    done = true;
    const replies = [...(annotation.replies ?? []), reply].slice(
      -MAX_REPLIES_PER_ANNOTATION,
    );
    return { ...annotation, replies };
  });
}

/**
 * What a reply shows in the comments pane. A reply that carries a review
 * state is that author's state record (ISO 32000-1 12.5.6.10, Table 172:
 * /State, /StateModel): it shows as the state, never as a chat message. Its
 * /Contents is the producer's own wording of the change -- Acrobat writes
 * "<State> set by <author>" (in its UI language) on every state record and
 * offers no way to add words to one -- so that wording is not repeated
 * under the state. The reply's data is untouched: it still round-trips as
 * the file wrote it, and the comment can still be answered.
 */
export function pdfReplyDisplayText(
  reply: Pick<PdfAnnotationReply, "text" | "reviewState">,
  stateLabel: (state: PdfAnnotationReviewState) => string,
): string {
  if (reply.reviewState) return stateLabel(reply.reviewState);
  return reply.text;
}
