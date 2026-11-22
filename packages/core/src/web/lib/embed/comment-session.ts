/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A comment session's writes to a workbook, a deck or a PDF (apps/api
 * app/services/embed/comments.py): the session's account only reads the
 * page, so each comment gesture -- add, reply, edit, resolve, reopen,
 * delete -- goes through the session's comment route, which applies that one
 * change to the page as it is (in the editor's own comment model) and writes
 * nothing else. The answer carries the thread as the page now holds it, for
 * the editor to show without saving anything itself. A refusal carries the
 * server's own words (another person's comment, a resolved thread, a comment
 * the file locks).
 */

import { ApiRequestError, apiPost } from "../api";

export type EmbedCommentOp = "add" | "reply" | "edit" | "resolve" | "reopen" | "delete";

/** Where a new comment goes, or (a workbook) which cell's thread. */
export type EmbedCommentAnchor =
  /** A PDF page (from 1), optionally the note's box in PDF points. */
  | { readonly page: number; readonly rect?: readonly [number, number, number, number] }
  /** A workbook cell: the sheet's id (or index) and `row:col` from 0. */
  | { readonly sheet: string | number; readonly cell: string }
  /** A deck slide (its id, or its place from 0), optionally one of its objects. */
  | { readonly slide: string | number; readonly block?: string };

export interface EmbedCommentRequest {
  readonly op: EmbedCommentOp;
  /** The thread: a PDF note's or a deck comment's id (a workbook names it by anchor). */
  readonly thread?: string;
  /** A message of the thread: a reply's id (a workbook's message id or its index). */
  readonly message?: string | number;
  readonly text?: string;
  readonly anchor?: EmbedCommentAnchor;
}

export interface EmbedCommentResult<Comment = unknown> {
  readonly op: EmbedCommentOp;
  /** The thread's id (a new comment's id after `add`). */
  readonly thread?: string;
  /** A new reply's id. */
  readonly message?: string;
  /** The thread as the page now holds it (the editor's model; null: deleted). */
  readonly comment: Comment | null;
  readonly content_version: number;
  /** A workbook's cell and sheet, a deck's slide. */
  readonly sheet?: string;
  readonly cell?: string;
  readonly slide?: string;
}

export type EmbedCommentWrite<Comment = unknown> =
  | { readonly ok: true; readonly result: EmbedCommentResult<Comment> }
  | { readonly ok: false; readonly code: string; readonly message: string };

/** The comment writes on their way, in order: one gesture after another,
 * as the editor's own saves go. */
let queue: Promise<unknown> = Promise.resolve();

export function writeEmbedComment<Comment = unknown>(
  request: EmbedCommentRequest,
): Promise<EmbedCommentWrite<Comment>> {
  const run = async (): Promise<EmbedCommentWrite<Comment>> => {
    try {
      const result = await apiPost<EmbedCommentResult<Comment>>("/api/v1/embed/session/comments", request);
      return { ok: true, result };
    } catch (error) {
      if (error instanceof ApiRequestError) {
        return { ok: false, code: error.code ?? "embed_comment_write_failed", message: error.message };
      }
      return { ok: false, code: "network", message: error instanceof Error ? error.message : "" };
    }
  };
  const next = queue.then(run, run);
  queue = next.catch(() => undefined);
  return next;
}
