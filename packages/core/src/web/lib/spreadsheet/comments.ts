/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for THREADED cell comments (Excel / WPS threaded comments).
 *
 * A cell's comment used to be a single text note — `{ text, author? }` — that a
 * write simply overwrote. Excel/WPS model a comment as a CONVERSATION: an
 * initial message plus an ordered list of replies, with a sheet-wide
 * "resolved / reopen" toggle that hides the thread without deleting it. This
 * module owns the data model + the three state transitions (migrate-legacy,
 * append-reply, toggle-resolve) so they stay deterministic and unit-testable
 * away from the React surface.
 *
 * DETERMINISM — none of these helpers read the wall clock. `appendReply` takes
 * an optional `ts` argument; when omitted the appended message simply has no
 * timestamp. Tests pass an explicit `ts` (or none) and never have to mock
 * `Date.now`. The component is the only place that may supply `Date.now()`.
 *
 * BACK-COMPAT — `migrateLegacyComment` accepts EITHER the new threaded shape or
 * the old single-text shape and always yields a thread (or `undefined` for an
 * empty/absent comment), so a sheet saved before threading keeps its note as a
 * one-message thread and never loses text or throws.
 */

/**
 * The longest text a note or a threaded comment message holds: Excel's cell
 * text length, as in the API page validator. At 2,000 a workbook with one
 * longer note failed validation and its import was refused (live audit B315).
 */
export const MAX_SPREADSHEET_COMMENT_TEXT_LENGTH = 32_767;

/** Native package receipt retained only for an imported Office remark. */
export interface CommentNativeSource {
  partName?: string;
  relationshipId?: string;
  [key: string]: unknown;
}

/** One message in a comment thread. */
export interface CommentMessage {
  /** Stable collaboration identity. Additive; native Office ids stay separate. */
  id?: string;
  /** Display name of the author, if known. */
  author?: string;
  /** MRAI actor identity for comments authored in this editor. */
  actorId?: string;
  /** The message body (already trimmed of surrounding blanks on append). */
  text: string;
  /**
   * Optional epoch-millis timestamp. Deterministic by construction: the pure
   * helpers never fabricate one — the caller passes it (or leaves it off).
   */
  ts?: number;
  /** Stable native identities used by threadedComments/persons patchback. */
  nativeId?: string;
  personId?: string;
  /** Exact native dT receipt. `ts` is the editor-friendly epoch mirror. */
  dateTime?: string;
  [key: string]: unknown;
}

/** A classic Excel Note (`comments*.xml` + VML), never a conversation. */
export interface NoteComment {
  kind: "note";
  text: string;
  author?: string;
  /**
   * Review > Notes > Show/Hide Note: the note stays open on the sheet until it
   * is hidden again. Excel stores it on the note's VML shape
   * (`visibility:visible` plus `<x:Visible/>`), so it is workbook data and
   * prints under "As displayed on sheet". Only `true` is stored; absent means
   * hidden, so a note that was never shown stays byte-stable.
   */
  visible?: true;
  nativeSource?: CommentNativeSource;
  nativeBlocker?: string;
}

/** A threaded cell comment: an ordered conversation plus its native done flag. */
export interface ThreadComment {
  kind: "thread";
  /** Ordered messages, oldest first. Always at least one for a live comment. */
  thread: CommentMessage[];
  /** OOXML `done`: resolved threads remain present and may be reopened. */
  done?: boolean;
  /**
   * In-progress reply text the composer is editing but has NOT committed to the
   * thread yet. Persisted so a draft survives a reload (mirrors the old single
   * note's live-typing behaviour); a blank/absent draft is never stored.
   */
  draft?: string;
  /**
   * LEGACY MIRROR — the first message's text/author, denormalized onto the
   * comment object. These exist purely for BACK / FORWARD compatibility:
   *  - the persisted-cell Zod schema still declares `comment.text` as required,
   *    so a threaded comment must carry `text` to validate (and not be dropped);
   *  - an OLDER build of the app that only understands the single-note shape
   *    keeps rendering the first message via `comment.text`.
   * They are recomputed from `thread[0]` on every transition — never authored
   * directly. A thread-less comment (draft only) carries no `text`.
   */
  text?: string;
  author?: string;
  nativeSource?: CommentNativeSource;
  nativeBlocker?: string;
}

export type CellComment = NoteComment | ThreadComment;

/**
 * Stamp the legacy `text`/`author` mirror onto the persisted object so it stays
 * valid against the (unchanged) cell schema — which still declares `text` as a
 * REQUIRED string — and an older reader still sees a note. The mirror comes from
 * the first thread message; a thread-less comment (only a pending draft, i.e. an
 * un-committed FIRST comment) mirrors the DRAFT text instead, exactly like the
 * old single-note path showed in-progress typing live. A comment with neither a
 * message nor a draft gets no `text`/`author`.
 */
function withLegacyMirror(comment: ThreadComment): ThreadComment {
  const first = comment.thread[0];
  const next: ThreadComment = { ...comment };
  if (first) {
    next.text = first.text;
    if (first.author) next.author = first.author;
    else delete next.author;
  } else if (comment.draft && comment.draft.trim()) {
    next.text = comment.draft;
    delete next.author;
  } else {
    delete next.text;
    delete next.author;
  }
  return next;
}

function asString(value: unknown): string {
  return typeof value === "string" ? value : value == null ? "" : String(value);
}

function asAuthor(value: unknown): string | undefined {
  return typeof value === "string" && value ? value : undefined;
}

function asTs(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function boundedNativeId(value: unknown): string | undefined {
  return typeof value === "string" && value.length > 0 && value.length <= 200
    ? value
    : undefined;
}

function boundedActorId(value: unknown): string | undefined {
  return typeof value === "string" && value.trim().length > 0 && value.length <= 200
    ? value.trim()
    : undefined;
}

function exactDateTime(value: unknown): string | undefined {
  if (typeof value !== "string" || value.length > 64) return undefined;
  // Excel emits native threaded-comment dT values both with and without an
  // explicit zone. Preserve either exact lexical form instead of erasing the
  // receipt during browser normalization.
  return /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?(?:Z|[+-]\d{2}:\d{2})?$/.test(
    value,
  )
    ? value
    : undefined;
}

function nativeSource(value: unknown): CommentNativeSource | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const record = value as Record<string, unknown>;
  const source: CommentNativeSource = {};
  if (typeof record.partName === "string" && record.partName.length <= 512) {
    source.partName = record.partName;
  }
  if (
    typeof record.relationshipId === "string" &&
    record.relationshipId.length <= 200
  ) {
    source.relationshipId = record.relationshipId;
  }
  return Object.keys(source).length > 0 ? source : undefined;
}

/**
 * Coerce one raw message-like object into a {@link CommentMessage}. Tolerant of
 * missing/extra keys; returns `null` when there is no usable text so callers can
 * filter blank entries out of a thread.
 */
function normalizeMessage(raw: unknown, index: number): CommentMessage | null {
  if (!raw || typeof raw !== "object") return null;
  const record = raw as Record<string, unknown>;
  const text = asString(record.text);
  if (!text.trim()) return null;
  const nativeId = boundedNativeId(record.nativeId);
  const actorId = boundedActorId(record.actorId);
  const ts = asTs(record.ts);
  const messageId =
    boundedNativeId(record.id) ??
    (nativeId
      ? `native:${nativeId}`
      : actorId && ts !== undefined
        ? `actor:${actorId.slice(0, 128)}:${ts}`
        : `legacy:${index}`);
  const message: CommentMessage = { id: messageId, text };
  const author = asAuthor(record.author);
  if (author) message.author = author;
  if (actorId) message.actorId = actorId;
  if (ts !== undefined) message.ts = ts;
  if (nativeId) message.nativeId = nativeId;
  const personId = boundedNativeId(record.personId);
  if (personId) message.personId = personId;
  const dateTime = exactDateTime(record.dateTime);
  if (dateTime) message.dateTime = dateTime;
  return message;
}

/**
 * Migrate ANY persisted comment shape — the new threaded one OR the legacy
 * single-text `{ text, author? }` — into a normalized {@link ThreadComment}.
 *
 * Rules:
 *  - A new-shape comment (`{ thread: [...] }`) keeps its messages (each
 *    normalized; blank/garbage messages dropped, order preserved) and its
 *    `resolved` flag + any non-blank `draft`.
 *  - A legacy single-text comment loads as a ONE-message thread
 *    `{ thread: [{ author, text }] }` so the saved note survives untouched.
 *  - An absent comment, or one with no usable text anywhere, yields `undefined`
 *    (the caller then stores no `comment` key — byte-stable for note-less cells).
 */
export function migrateLegacyComment(raw: unknown): CellComment | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const record = raw as Record<string, unknown>;

  if (!Array.isArray(record.thread) && record.kind !== "thread") {
    const text = asString(record.text);
    if (!text.trim()) return undefined;
    const note: NoteComment = { kind: "note", text };
    const author = asAuthor(record.author);
    if (author) note.author = author;
    if (record.visible === true) note.visible = true;
    const source = nativeSource(record.nativeSource);
    if (source) note.nativeSource = source;
    if (typeof record.nativeBlocker === "string" && record.nativeBlocker) {
      note.nativeBlocker = record.nativeBlocker.slice(0, 200);
    }
    return note;
  }

  let messages: CommentMessage[];
  if (Array.isArray(record.thread)) {
    // Already threaded: normalize each message, dropping blanks but keeping order.
    messages = record.thread
      .map((entry, index) => normalizeMessage(entry, index))
      .filter((entry): entry is CommentMessage => entry !== null);
  } else messages = [];

  const draft = asString(record.draft);
  if (messages.length === 0) {
    if (typeof record.nativeBlocker === "string" && record.nativeBlocker) {
      return {
        kind: "thread",
        text: "",
        thread: [],
        nativeBlocker: record.nativeBlocker.slice(0, 200),
      };
    }
    // No real message: only keep the comment alive if there is a pending draft,
    // otherwise drop it entirely so note-less cells stay byte-stable.
    if (draft.trim()) {
      return withLegacyMirror({ kind: "thread", thread: [], draft });
    }
    return undefined;
  }

  const comment: ThreadComment = { kind: "thread", thread: messages };
  if (record.done === true || record.resolved === true) comment.done = true;
  if (draft.trim()) comment.draft = draft;
  const source = nativeSource(record.nativeSource);
  if (source) comment.nativeSource = source;
  if (typeof record.nativeBlocker === "string" && record.nativeBlocker) {
    comment.nativeBlocker = record.nativeBlocker.slice(0, 200);
  }
  return withLegacyMirror(comment);
}

export function isThreadComment(
  comment: CellComment | undefined,
): comment is ThreadComment {
  return comment?.kind === "thread";
}

export function isNoteComment(
  comment: CellComment | undefined,
): comment is NoteComment {
  return comment?.kind === "note";
}

/**
 * Append a reply message to a thread. A blank / whitespace-only `text` is a
 * NO-OP (returns the comment unchanged) so the thread never grows empty
 * messages. The new message is appended at the END (order preserved). `ts` is
 * optional and passed through verbatim — the helper never reads the clock.
 *
 * Appending also CLEARS any pending `draft` (the composer has just committed
 * its text). A previously-resolved thread is reopened by a new reply, matching
 * Excel/WPS (a fresh reply un-resolves the conversation).
 */
export function appendReply(
  comment: ThreadComment | undefined,
  text: string,
  author?: string,
  ts?: number,
  actorId?: string,
  messageId?: string,
): ThreadComment | undefined {
  const trimmed = (text ?? "").trim();
  if (!trimmed) return comment;
  const base = comment?.thread ?? [];
  const cleanActorId = boundedActorId(actorId);
  const cleanTimestamp = asTs(ts);
  const stableMessageId =
    boundedNativeId(messageId) ??
    `local:${(cleanActorId ?? "anonymous").slice(0, 64)}:${cleanTimestamp ?? "undated"}:${base.length}`;
  const message: CommentMessage = { id: stableMessageId, text: trimmed };
  const cleanAuthor = asAuthor(author);
  if (cleanAuthor) message.author = cleanAuthor;
  if (cleanActorId) message.actorId = cleanActorId;
  if (cleanTimestamp !== undefined) message.ts = cleanTimestamp;

  // A new reply reopens the conversation (drop the resolved flag); drop the
  // committed draft. Only `thread` is carried — `resolved`/`draft` intentionally
  // reset. The legacy `text`/`author` mirror is recomputed from thread[0].
  return withLegacyMirror({ kind: "thread", thread: [...base, message] });
}

/**
 * Replace one message body while preserving its native/person/author identity.
 * A stale index or blank body is a no-op; the caller owns authorisation.
 */
export function editCommentMessage(
  comment: ThreadComment | undefined,
  index: number,
  text: string,
): ThreadComment | undefined {
  if (!comment || !Number.isInteger(index) || index < 0 || index >= comment.thread.length) {
    return comment;
  }
  const trimmed = text.trim();
  if (!trimmed || trimmed === comment.thread[index]?.text) return comment;
  const thread = comment.thread.map((message, messageIndex) =>
    messageIndex === index ? { ...message, text: trimmed } : message,
  );
  return withLegacyMirror({
    ...comment,
    kind: "thread",
    thread,
  });
}

/**
 * Delete a reply while keeping the root message and thread identity intact.
 * Excel exposes root deletion as Delete thread, so index 0 deliberately fails
 * closed here.
 */
export function deleteCommentReply(
  comment: ThreadComment | undefined,
  index: number,
): ThreadComment | undefined {
  if (!comment || !Number.isInteger(index) || index <= 0 || index >= comment.thread.length) {
    return comment;
  }
  return withLegacyMirror({
    ...comment,
    kind: "thread",
    thread: comment.thread.filter((_message, messageIndex) => messageIndex !== index),
  });
}

/**
 * Toggle the thread's resolved flag. Resolving is a FLAG, never a delete — the
 * messages are kept verbatim. Returns the comment unchanged (no flag) when there
 * is nothing to resolve, i.e. an absent comment.
 */
export function toggleResolved(
  comment: ThreadComment | undefined,
): ThreadComment | undefined {
  if (!comment) return comment;
  const next: ThreadComment = {
    kind: "thread",
    thread: comment.thread,
    ...(comment.nativeSource ? { nativeSource: comment.nativeSource } : {}),
    ...(comment.nativeBlocker ? { nativeBlocker: comment.nativeBlocker } : {}),
  };
  if (comment.draft && comment.draft.trim()) next.draft = comment.draft;
  if (!comment.done) next.done = true;
  return withLegacyMirror(next);
}

/**
 * Set (or clear) the in-progress reply draft on a comment WITHOUT committing it
 * to the thread. A blank draft removes the `draft` key. When the comment has no
 * thread yet and the draft is blank, the whole comment collapses to `undefined`
 * (a note-less cell). The thread + resolved flag are otherwise untouched.
 */
export function setDraft(
  comment: ThreadComment | undefined,
  text: string,
): ThreadComment | undefined {
  const draft = text ?? "";
  const base = comment?.thread ?? [];
  if (!draft.trim()) {
    if (base.length === 0) return undefined;
    const next: ThreadComment = { kind: "thread", thread: base };
    if (comment?.done) next.done = true;
    return withLegacyMirror(next);
  }
  const next: ThreadComment = { kind: "thread", thread: base, draft };
  if (comment?.done && base.length > 0) next.done = true;
  return withLegacyMirror(next);
}

/** True when a comment carries at least one real message (a visible thread). */
export function hasThread(comment: CellComment | undefined): boolean {
  return Boolean(isThreadComment(comment) && comment.thread.length > 0);
}

export function hasNote(comment: CellComment | undefined): boolean {
  return Boolean(isNoteComment(comment) && comment.text.trim());
}

export function hasComment(comment: CellComment | undefined): boolean {
  return hasThread(comment) || hasNote(comment);
}

/** The first message's text, used for the cell tooltip / indicator title. */
export function commentSummary(comment: CellComment | undefined): string {
  return isThreadComment(comment)
    ? (comment.thread[0]?.text ?? "")
    : isNoteComment(comment)
      ? comment.text
      : "";
}

/** Create/update a classic Note without ever manufacturing a thread. */
export function setNote(
  text: string,
  author?: string,
  previous?: NoteComment,
): NoteComment | undefined {
  if (!text.trim()) return undefined;
  return {
    kind: "note",
    text,
    ...(asAuthor(author) ? { author: asAuthor(author) } : {}),
    // Editing a shown note's text leaves it shown, as in Excel.
    ...(previous?.visible === true ? { visible: true as const } : {}),
    ...(previous?.nativeSource ? { nativeSource: previous.nativeSource } : {}),
    ...(previous?.nativeBlocker ? { nativeBlocker: previous.nativeBlocker } : {}),
  };
}

/**
 * Review > Notes > Show/Hide Note: flip whether the note stays open on the
 * sheet. Showing stores `visible: true`; hiding removes the key, so a note
 * shown and hidden again is exactly the note it was. Anything that is not a
 * classic note with text, or a note held in native custody, is unchanged.
 */
export function toggleNoteVisible(
  comment: CellComment | undefined,
): CellComment | undefined {
  if (!isNoteComment(comment) || !comment.text.trim() || comment.nativeBlocker) {
    return comment;
  }
  if (comment.visible === true) {
    const hidden: NoteComment = { ...comment };
    delete hidden.visible;
    return hidden;
  }
  return { ...comment, visible: true };
}

/** Clipboard copies are new remarks: keep content/person, strip package IDs. */
export function cloneCellCommentForPaste(
  comment: CellComment | undefined,
): CellComment | undefined {
  if (isNoteComment(comment)) {
    return {
      kind: "note",
      text: comment.text,
      ...(comment.author ? { author: comment.author } : {}),
      // Excel's copy of a shown note is shown too (measured in Excel 16).
      ...(comment.visible === true ? { visible: true as const } : {}),
    };
  }
  if (!isThreadComment(comment)) return undefined;
  return withLegacyMirror({
    kind: "thread",
    thread: comment.thread.map((message) => {
      const clone = { ...message };
      delete clone.nativeId;
      return clone;
    }),
    ...(comment.done ? { done: true } : {}),
    ...(comment.draft ? { draft: comment.draft } : {}),
  });
}
