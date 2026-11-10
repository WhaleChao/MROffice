/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Persistent REVIEW COMMENTS for presentation slides (PowerPoint / WPS "Review →
 * Comments": a threaded comment pinned to a slide or to a specific object, with
 * replies and a resolved state). These are PERSISTED on the deck (a `comments`
 * array on each slide), distinct from the ephemeral present-mode ink/laser in
 * `presentation-annotation.ts`, which is never saved.
 *
 * Dependency-free (no React / no DOM): the thread model + every mutation
 * (add / reply / resolve / reopen / delete) and the load-edge normalizer live
 * here so the editor component stays a thin render and the logic is unit-tested
 * in isolation. Every mutation returns a NEW array (immutable) so it slots
 * straight into React state + the deck's single commit path (undo/redo +
 * autosave). The model mirrors the persisted `PdfAnnotation` precedent
 * (`id` / `text` / `createdAt`) extended with replies + a resolved flag.
 *
 * BYTE-CLEAN: a slide that has no comments persists NO `comments` key (the
 * normalizer drops an empty array), so existing decks serialize byte-for-byte as
 * before — the field is purely additive.
 */

/** Hard caps so a malicious / runaway deck can't grow comment storage without bound. */
export const MAX_COMMENTS_PER_SLIDE = 200;
export const MAX_REPLIES_PER_COMMENT = 100;
/** Max stored length of a single comment / reply body (chars). */
export const MAX_COMMENT_BODY_LENGTH = 4000;
/** Max stored length of an author display name (chars). */
export const MAX_COMMENT_AUTHOR_LENGTH = 120;
/** Max stored length of the stable product/user id attached to a modern comment. */
export const MAX_COMMENT_AUTHOR_ID_LENGTH = 256;

/** A single reply within a comment thread. */
export interface SlideCommentReply {
  /** Stable id (React key + reply targeting). */
  id: string;
  /** The reply text. */
  body: string;
  /** Optional author display name. */
  author?: string;
  /** Stable product/user id. Present only for modern-comment identity. */
  authorId?: string;
  /** ISO timestamp the reply was created. */
  createdAt: string;
}

/**
 * A top-level review comment pinned to a slide (or, when `blockId` is set, to a
 * specific object on the slide). Carries an ordered list of `replies` and a
 * `resolved` flag (PowerPoint "Resolve"). Only meaningful / true fields are
 * persisted so a simple comment stays compact.
 */
export interface SlideComment {
  /** Stable id (React key + comment targeting). */
  id: string;
  /** The comment text. */
  body: string;
  /** Optional author display name. */
  author?: string;
  /** Stable product/user id. Present only for modern-comment identity. */
  authorId?: string;
  /** ISO timestamp the comment was created. */
  createdAt: string;
  /** Resolved (closed) state. Absent / false ⇒ open. */
  resolved?: boolean;
  /** Optional id of the slide OBJECT this comment is anchored to. */
  blockId?: string;
  /** Ordered replies (oldest first). Absent ⇒ no replies. */
  replies?: SlideCommentReply[];
}

/** A finite-string fallback. */
function asString(value: unknown): string {
  return typeof value === "string" ? value : "";
}

/**
 * Clamp + trim a body string to the stored max length. Trailing/leading
 * whitespace is trimmed; the result is capped at {@link MAX_COMMENT_BODY_LENGTH}.
 */
export function clampCommentBody(value: unknown): string {
  return asString(value).trim().slice(0, MAX_COMMENT_BODY_LENGTH);
}

/** Clamp + trim an author name; empty when absent. */
function clampAuthor(value: unknown): string {
  return asString(value).trim().slice(0, MAX_COMMENT_AUTHOR_LENGTH);
}

/** Clamp a stable author id; empty/invalid ids are omitted rather than invented. */
export function clampCommentAuthorId(value: unknown): string {
  return asString(value).trim().slice(0, MAX_COMMENT_AUTHOR_ID_LENGTH);
}

/**
 * Modern comments are author-owned. Classic comments have no stable author id,
 * so they retain the legacy editable-file policy instead of being silently
 * upgraded to a different comment model.
 */
export function canModifyCommentAuthor(
  storedAuthorId: unknown,
  actorId: unknown,
): boolean {
  const owner = clampCommentAuthorId(storedAuthorId);
  if (!owner) return true;
  const actor = clampCommentAuthorId(actorId);
  return Boolean(actor && actor === owner);
}

/** A fresh, reasonably-unique id. Uses `crypto.randomUUID` when available. */
function freshId(): string {
  try {
    if (
      typeof crypto !== "undefined" &&
      typeof crypto.randomUUID === "function"
    ) {
      return crypto.randomUUID();
    }
  } catch {
    // fall through to the manual id
  }
  return `c-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

/** An ISO timestamp; accepts an injected clock (tests) else `Date.now()`. */
function nowIso(now?: number | string): string {
  if (typeof now === "string") return now;
  const ms = typeof now === "number" && Number.isFinite(now) ? now : Date.now();
  return new Date(ms).toISOString();
}

/**
 * Build a new top-level {@link SlideComment} from a body string. Returns `null`
 * when the body is empty after trimming (an empty comment is never created). The
 * optional `author` / `blockId` anchor and an injectable `now`/`id` (for
 * deterministic tests) are honoured; only meaningful fields are set so the
 * serialized comment stays compact.
 */
export function newComment(
  body: string,
  options?: {
    author?: string;
    authorId?: string;
    blockId?: string;
    now?: number | string;
    id?: string;
  },
): SlideComment | null {
  const text = clampCommentBody(body);
  if (!text) return null;
  const author = clampAuthor(options?.author);
  const authorId = clampCommentAuthorId(options?.authorId);
  const blockId =
    typeof options?.blockId === "string" && options.blockId
      ? options.blockId
      : undefined;
  return {
    id: options?.id || freshId(),
    body: text,
    createdAt: nowIso(options?.now),
    ...(author ? { author } : {}),
    ...(authorId ? { authorId } : {}),
    ...(blockId ? { blockId } : {}),
  };
}

/**
 * Append a new comment built from `body` to `comments`, returning a NEW array.
 * An empty body is a no-op (returns the SAME array reference). Enforces the
 * per-slide cap (the oldest comment is dropped if the cap would be exceeded —
 * matches the "bounded storage" invariant; in practice the cap is never hit).
 */
export function addComment(
  comments: readonly SlideComment[],
  body: string,
  options?: {
    author?: string;
    authorId?: string;
    blockId?: string;
    now?: number | string;
    id?: string;
  },
): SlideComment[] {
  const comment = newComment(body, options);
  if (!comment) return comments as SlideComment[];
  const next = [...comments, comment];
  if (next.length > MAX_COMMENTS_PER_SLIDE) {
    next.splice(0, next.length - MAX_COMMENTS_PER_SLIDE);
  }
  return next;
}

/**
 * Append a reply built from `body` to the comment with `commentId`, returning a
 * NEW array. An empty body, an unknown `commentId`, or a comment already at the
 * reply cap is a no-op (returns the SAME array reference). The parent comment's
 * other fields are preserved; replies stay oldest-first.
 */
export function addReply(
  comments: readonly SlideComment[],
  commentId: string,
  body: string,
  options?: {
    author?: string;
    authorId?: string;
    now?: number | string;
    id?: string;
  },
): SlideComment[] {
  const text = clampCommentBody(body);
  if (!text) return comments as SlideComment[];
  let changed = false;
  const next = comments.map((comment) => {
    if (comment.id !== commentId) return comment;
    const existing = Array.isArray(comment.replies) ? comment.replies : [];
    if (existing.length >= MAX_REPLIES_PER_COMMENT) return comment;
    const author = clampAuthor(options?.author);
    const authorId = clampCommentAuthorId(options?.authorId);
    const reply: SlideCommentReply = {
      id: options?.id || freshId(),
      body: text,
      createdAt: nowIso(options?.now),
      ...(author ? { author } : {}),
      ...(authorId ? { authorId } : {}),
    };
    changed = true;
    return { ...comment, replies: [...existing, reply] };
  });
  return changed ? next : (comments as SlideComment[]);
}

/**
 * Replace one top-level comment body when the actor owns the modern comment (or
 * the target is a classic comment without stable identity). Empty text, unknown
 * ids and ownership mismatches are all no-ops, preserving the input reference.
 */
export function updateCommentBody(
  comments: readonly SlideComment[],
  commentId: string,
  body: string,
  actorId: string | null,
): SlideComment[] {
  const text = clampCommentBody(body);
  if (!text) return comments as SlideComment[];
  let changed = false;
  const next = comments.map((comment) => {
    if (
      comment.id !== commentId ||
      // A resolved thread is closed to edits until it is reopened, whatever
      // editor was still open when it was resolved (PPT-B417).
      comment.resolved === true ||
      !canModifyCommentAuthor(comment.authorId, actorId) ||
      comment.body === text
    ) {
      return comment;
    }
    changed = true;
    return { ...comment, body: text };
  });
  return changed ? next : (comments as SlideComment[]);
}

/** Author-gated equivalent of {@link deleteComment}. */
export function deleteCommentForActor(
  comments: readonly SlideComment[],
  commentId: string,
  actorId: string | null,
): SlideComment[] {
  const target = comments.find((comment) => comment.id === commentId);
  if (!target || !canModifyCommentAuthor(target.authorId, actorId)) {
    return comments as SlideComment[];
  }
  return deleteComment(comments, commentId);
}

/**
 * Replace one reply body with the same classic/modern author gate as the parent
 * helper. A reply remains independently owned by its own author id.
 */
export function updateReplyBody(
  comments: readonly SlideComment[],
  commentId: string,
  replyId: string,
  body: string,
  actorId: string | null,
): SlideComment[] {
  const text = clampCommentBody(body);
  if (!text) return comments as SlideComment[];
  let changed = false;
  const next = comments.map((comment) => {
    if (
      comment.id !== commentId ||
      comment.resolved === true ||
      !Array.isArray(comment.replies)
    ) {
      return comment;
    }
    let replyChanged = false;
    const replies = comment.replies.map((reply) => {
      if (
        reply.id !== replyId ||
        !canModifyCommentAuthor(reply.authorId, actorId) ||
        reply.body === text
      ) {
        return reply;
      }
      replyChanged = true;
      return { ...reply, body: text };
    });
    if (!replyChanged) return comment;
    changed = true;
    return { ...comment, replies };
  });
  return changed ? next : (comments as SlideComment[]);
}

/** Author-gated equivalent of {@link deleteReply}. */
export function deleteReplyForActor(
  comments: readonly SlideComment[],
  commentId: string,
  replyId: string,
  actorId: string | null,
): SlideComment[] {
  const parent = comments.find((comment) => comment.id === commentId);
  const reply = parent?.replies?.find((candidate) => candidate.id === replyId);
  if (!reply || !canModifyCommentAuthor(reply.authorId, actorId)) {
    return comments as SlideComment[];
  }
  return deleteReply(comments, commentId, replyId);
}

/** Whether a deck comment has opted into stable modern-comment identity. */
export function commentUsesModernIdentity(comment: SlideComment): boolean {
  if (clampCommentAuthorId(comment.authorId)) return true;
  return Boolean(
    comment.replies?.some((reply) => clampCommentAuthorId(reply.authorId)),
  );
}

/**
 * Set the `resolved` flag of the comment with `commentId`. Returns a NEW array
 * when it changed, else the SAME reference. Resolving sets `resolved: true`;
 * reopening DELETES the field (so an open comment serializes byte-clean — no
 * `resolved: false`). An unknown id is a no-op.
 */
export function setCommentResolved(
  comments: readonly SlideComment[],
  commentId: string,
  resolved: boolean,
): SlideComment[] {
  let changed = false;
  const next = comments.map((comment) => {
    if (comment.id !== commentId) return comment;
    const current = comment.resolved === true;
    if (current === resolved) return comment;
    changed = true;
    if (resolved) return { ...comment, resolved: true };
    const { resolved: _drop, ...rest } = comment;
    void _drop;
    return rest;
  });
  return changed ? next : (comments as SlideComment[]);
}

/** Resolve (close) a comment — convenience over {@link setCommentResolved}. */
export function resolveComment(
  comments: readonly SlideComment[],
  commentId: string,
): SlideComment[] {
  return setCommentResolved(comments, commentId, true);
}

/** Reopen a resolved comment — convenience over {@link setCommentResolved}. */
export function reopenComment(
  comments: readonly SlideComment[],
  commentId: string,
): SlideComment[] {
  return setCommentResolved(comments, commentId, false);
}

/**
 * Delete the comment with `commentId` (and its whole reply thread). Returns a
 * NEW array when something was removed, else the SAME reference (unknown id).
 */
export function deleteComment(
  comments: readonly SlideComment[],
  commentId: string,
): SlideComment[] {
  const next = comments.filter((comment) => comment.id !== commentId);
  return next.length === comments.length ? (comments as SlideComment[]) : next;
}

/**
 * Delete a single reply (`replyId`) from the comment with `commentId`. Returns a
 * NEW array when something was removed, else the SAME reference. The comment
 * itself is kept (deleting its last reply drops the `replies` field so it stays
 * byte-clean).
 */
export function deleteReply(
  comments: readonly SlideComment[],
  commentId: string,
  replyId: string,
): SlideComment[] {
  let changed = false;
  const next = comments.map((comment) => {
    if (comment.id !== commentId || !Array.isArray(comment.replies)) {
      return comment;
    }
    const keptReplies = comment.replies.filter((reply) => reply.id !== replyId);
    if (keptReplies.length === comment.replies.length) return comment;
    changed = true;
    if (keptReplies.length === 0) {
      const { replies: _drop, ...rest } = comment;
      void _drop;
      return rest;
    }
    return { ...comment, replies: keptReplies };
  });
  return changed ? next : (comments as SlideComment[]);
}

/** Count the OPEN (unresolved) comments in a list. */
export function countOpenComments(comments: readonly SlideComment[]): number {
  let open = 0;
  for (const comment of comments) {
    if (comment.resolved !== true) open += 1;
  }
  return open;
}

/** Total comment count (open + resolved). */
export function countComments(comments: readonly SlideComment[]): number {
  return Array.isArray(comments) ? comments.length : 0;
}

/**
 * Normalize a raw persisted value into a clean {@link SlideComment} array at the
 * deck LOAD edge: drops non-object / empty-body entries, trims + caps bodies and
 * authors, repairs missing ids / timestamps, keeps `resolved` only when strictly
 * `true`, keeps `blockId` only when a non-empty string, normalizes each reply the
 * same way (dropping empty ones + the whole `replies` field when none survive),
 * and enforces the per-slide / per-comment caps. A non-array input, or one that
 * normalizes to nothing, returns `[]` (the caller omits the key ⇒ byte-clean).
 * Pure + deterministic: a missing id/timestamp is filled deterministically from
 * the index so re-normalizing a clean array is idempotent in tests that inject
 * ids, and never throws.
 */
export function normalizeComments(value: unknown): SlideComment[] {
  if (!Array.isArray(value)) return [];
  const out: SlideComment[] = [];
  for (let i = 0; i < value.length && out.length < MAX_COMMENTS_PER_SLIDE; i += 1) {
    const raw = value[i];
    if (!raw || typeof raw !== "object") continue;
    const record = raw as Record<string, unknown>;
    const body = clampCommentBody(record.body ?? record.text);
    if (!body) continue;
    const id =
      typeof record.id === "string" && record.id ? record.id : `comment-${i + 1}`;
    const createdAt =
      typeof record.createdAt === "string" && record.createdAt
        ? record.createdAt
        : new Date(0).toISOString();
    const author = clampAuthor(record.author);
    const authorId = clampCommentAuthorId(record.authorId);
    const blockId =
      typeof record.blockId === "string" && record.blockId
        ? record.blockId
        : undefined;
    const replies = normalizeReplies(record.replies, id);
    out.push({
      id,
      body,
      createdAt,
      ...(record.resolved === true ? { resolved: true } : {}),
      ...(author ? { author } : {}),
      ...(authorId ? { authorId } : {}),
      ...(blockId ? { blockId } : {}),
      ...(replies.length > 0 ? { replies } : {}),
    });
  }
  return out;
}

/** Normalize the `replies` of one comment (see {@link normalizeComments}). */
function normalizeReplies(
  value: unknown,
  commentId: string,
): SlideCommentReply[] {
  if (!Array.isArray(value)) return [];
  const out: SlideCommentReply[] = [];
  for (
    let i = 0;
    i < value.length && out.length < MAX_REPLIES_PER_COMMENT;
    i += 1
  ) {
    const raw = value[i];
    if (!raw || typeof raw !== "object") continue;
    const record = raw as Record<string, unknown>;
    const body = clampCommentBody(record.body ?? record.text);
    if (!body) continue;
    const id =
      typeof record.id === "string" && record.id
        ? record.id
        : `${commentId}-reply-${i + 1}`;
    const createdAt =
      typeof record.createdAt === "string" && record.createdAt
        ? record.createdAt
        : new Date(0).toISOString();
    const author = clampAuthor(record.author);
    const authorId = clampCommentAuthorId(record.authorId);
    out.push({
      id,
      body,
      createdAt,
      ...(author ? { author } : {}),
      ...(authorId ? { authorId } : {}),
    });
  }
  return out;
}
