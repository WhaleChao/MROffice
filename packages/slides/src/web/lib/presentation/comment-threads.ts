/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The comment thread model of the Review pane: thread/message types, author
 * initials, and the flattened message list of a thread.
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */

import type { SlideComment, SlideCommentReply } from "@mrnote/core/web/lib/presentation-comments.ts";

export interface PresentationCommentMessage {
  author?: string;
  authorId?: string;
  body: string;
  commentId: string;
  createdAt: string;
  id: string;
  kind: "comment" | "reply";
}

export interface PresentationCommentThread {
  comment: SlideComment | null;
}

export function presentationCommentInitials(
  author: string | undefined,
  anonymous: string,
): string {
  const value = (author || anonymous).trim();
  if (!value) return "?";
  const words = value.split(/\s+/).filter(Boolean);
  if (words.length > 1) {
    return `${words[0]?.[0] ?? ""}${words.at(-1)?.[0] ?? ""}`.toUpperCase();
  }
  return Array.from(value).slice(0, 2).join("").toUpperCase();
}

export function presentationCommentMessages(
  comment: SlideComment | null,
): PresentationCommentMessage[] {
  if (!comment) return [];
  const parent: PresentationCommentMessage = {
    author: comment.author,
    authorId: comment.authorId,
    body: comment.body,
    commentId: comment.id,
    createdAt: comment.createdAt,
    id: comment.id,
    kind: "comment",
  };
  return [
    parent,
    ...(comment.replies ?? []).map(
      (reply: SlideCommentReply): PresentationCommentMessage => ({
        author: reply.author,
        authorId: reply.authorId,
        body: reply.body,
        commentId: comment.id,
        createdAt: reply.createdAt,
        id: reply.id,
        kind: "reply",
      }),
    ),
  ];
}
