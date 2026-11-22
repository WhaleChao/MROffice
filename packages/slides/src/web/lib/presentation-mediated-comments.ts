/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A deck's comments in an embed comment session (lib/embed/comment-session):
 * the session's account only reads the page, so the editor does not save a
 * comment gesture with the page. It sends the gesture to the session's
 * comment route, which writes that one change and answers with the thread as
 * the page now holds it; the editor then shows that thread in place, outside
 * the page's saves and its history.
 */

import type { EmbedCommentRequest, EmbedCommentResult } from "@mrnote/core/web/lib/embed/comment-session.ts";
import type { SlideComment } from "@mrnote/core/web/lib/presentation-comments.ts";
import { slideWithComments, type PresentationReviewSlide } from "./presentation-review-buttons";

/** One comment gesture of the deck's Comments pane or Review tab. */
export type PresentationCommentGesture =
  | { readonly kind: "add"; readonly slideId: string; readonly blockId?: string; readonly text: string }
  | { readonly kind: "reply"; readonly commentId: string; readonly text: string }
  | { readonly kind: "edit"; readonly commentId: string; readonly replyId?: string; readonly text: string }
  | { readonly kind: "resolve"; readonly commentId: string; readonly resolved: boolean }
  | { readonly kind: "delete"; readonly commentId: string; readonly replyId?: string };

/** The comment route's request for a gesture: a new comment names its slide
 * (and the object it is pinned to), everything else names its thread, and an
 * edit or delete of a reply names that reply too (omitted, the root). */
export function presentationEmbedCommentRequest(gesture: PresentationCommentGesture): EmbedCommentRequest {
  switch (gesture.kind) {
    case "add":
      return {
        op: "add",
        anchor: gesture.blockId ? { slide: gesture.slideId, block: gesture.blockId } : { slide: gesture.slideId },
        text: gesture.text,
      };
    case "reply":
      return { op: "reply", thread: gesture.commentId, text: gesture.text };
    case "edit":
      return gesture.replyId
        ? { op: "edit", thread: gesture.commentId, message: gesture.replyId, text: gesture.text }
        : { op: "edit", thread: gesture.commentId, text: gesture.text };
    case "resolve":
      return { op: gesture.resolved ? "resolve" : "reopen", thread: gesture.commentId };
    case "delete":
      return gesture.replyId
        ? { op: "delete", thread: gesture.commentId, message: gesture.replyId }
        : { op: "delete", thread: gesture.commentId };
  }
}

/**
 * The slides with the route's answer in place: the thread replaces the one
 * of the same id where it stands (a new one goes last, as the page appends
 * it), and a deleted one (`comment: null`) leaves its slide. The slide is
 * the one the answer names, else the one that holds the thread. The same
 * list comes back when nothing changes, so the caller knows not to redraw.
 */
export function slidesWithEmbedComment<S extends PresentationReviewSlide>(
  slides: S[],
  result: Pick<EmbedCommentResult<SlideComment>, "comment" | "slide" | "thread">,
): S[] {
  const threadId = result.comment?.id ?? result.thread;
  if (!threadId) return slides;
  const holds = (slide: S) => Array.isArray(slide.comments) && slide.comments.some((entry) => entry.id === threadId);
  let at = result.slide ? slides.findIndex((slide) => slide.id === result.slide) : -1;
  if (at < 0) at = slides.findIndex(holds);
  if (at < 0) return slides;
  const slide = slides[at];
  const current = Array.isArray(slide.comments) ? slide.comments : [];
  const index = current.findIndex((entry) => entry.id === threadId);
  let next: SlideComment[];
  if (result.comment === null) {
    if (index < 0) return slides;
    next = current.filter((_, position) => position !== index);
  } else if (index < 0) {
    next = [...current, result.comment];
  } else {
    next = current.slice();
    next[index] = result.comment;
  }
  const out = slides.slice();
  out[at] = slideWithComments(slide, next);
  return out;
}
