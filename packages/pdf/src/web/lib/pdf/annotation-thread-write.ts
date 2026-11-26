/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  MAX_REPLIES_PER_ANNOTATION,
  type PdfAnnotationReply,
} from "@mrnote/core/web/lib/pdf/annotation-replies.ts";

/**
 * The stored thread with this writer's replies it does not hold yet after
 * it, in the order they were written. Capped as every thread is.
 */
export function mergePdfAnnotationReplies(
  stored: readonly PdfAnnotationReply[],
  mine: readonly PdfAnnotationReply[],
  max: number = MAX_REPLIES_PER_ANNOTATION,
): PdfAnnotationReply[] {
  const held = new Set<string>();
  const thread: PdfAnnotationReply[] = [];
  for (const reply of [...stored, ...mine]) {
    if (held.has(reply.id)) continue;
    held.add(reply.id);
    thread.push(reply);
  }
  return thread.length > max ? thread.slice(-max) : thread;
}

type Payload = Record<string, unknown>;

export interface PdfThreadedAnnotationWrite {
  /** The payload the row now holds. */
  readonly payload: Payload;
  /** Its thread. */
  readonly replies: PdfAnnotationReply[];
  /** Whether the row held replies this writer had not sent. */
  readonly joined: boolean;
}

export interface PdfThreadedAnnotationWriter {
  /** A row as a list read gave it: its token and its thread. */
  saw(id: string, payloadSha256: unknown, replies: readonly PdfAnnotationReply[] | undefined): void;
  /**
   * Write `payload` to the row.
   *
   * `fields: "mine"` is the user's change to the annotation itself (its
   * text, colour, place or status): those fields are written as given and
   * only the thread is joined. `"theirs"` is a write that only adds replies:
   * over a row someone else changed, the other fields are the stored ones.
   * A `"theirs"` payload that is a thread alone (the reader does not show
   * the annotation) is never written as it is: the stored row is asked for
   * first and the thread is joined to it. `"place"` puts the row where the
   * page holds the annotation: over a row someone else changed, only its
   * `page`, `x` and `y` are this payload's, the other fields the stored ones.
   */
  write(id: string, payload: Payload, fields: "mine" | "theirs" | "place"): Promise<PdfThreadedAnnotationWrite>;
  /** The replies this writer knows the row to hold. */
  known(id: string): readonly PdfAnnotationReply[];
}

/** A stale write is made again this many times before it is a failed write. */
const MAX_STALE_WRITES = 5;

/**
 * The token of a row this writer has not read: it matches no stored payload,
 * so the first write is answered with the row as it stands instead of being
 * made blind.
 */
export const PDF_ANNOTATION_UNREAD_TOKEN = "unread";

function record(value: unknown): Payload | null {
  return value && typeof value === "object" && !Array.isArray(value) ? (value as Payload) : null;
}

/** The row a stale write was answered with, or null for any other failure. */
function storedRowOfStaleWrite(error: unknown): { payload: Payload; sha: string } | null {
  const failure = record(error);
  if (!failure || failure.status !== 409) return null;
  const row = record(record(failure.details)?.annotation);
  const payload = record(row?.payload_json);
  const sha = row?.payload_sha256;
  return payload && typeof sha === "string" && sha ? { payload, sha } : null;
}

function withThread(payload: Payload, replies: readonly PdfAnnotationReply[]): Payload {
  // A row that had no `replies` key keeps none while its thread is empty.
  if (replies.length === 0 && !("replies" in payload)) return payload;
  return { ...payload, replies };
}

export function createPdfThreadedAnnotationWriter(deps: {
  /** PATCH the row; rejects with the API's error (its `status` and `details`). */
  readonly patch: (
    id: string,
    body: { payload_json: Payload; expected_payload_sha256: string },
  ) => Promise<unknown>;
  /** The replies of a payload as the reader parses them. */
  readonly parseReplies: (id: string, raw: unknown) => PdfAnnotationReply[];
}): PdfThreadedAnnotationWriter {
  const tokens = new Map<string, string>();
  const threads = new Map<string, readonly PdfAnnotationReply[]>();

  const known = (id: string): readonly PdfAnnotationReply[] => threads.get(id) ?? [];
  const learn = (id: string, sha: unknown, replies: readonly PdfAnnotationReply[]): void => {
    if (typeof sha === "string" && sha) tokens.set(id, sha);
    else tokens.delete(id);
    // What this writer once saw in the thread stays in it: a list read
    // older than a reply already seen must not take that reply back out.
    threads.set(id, mergePdfAnnotationReplies(replies, known(id)));
  };

  return {
    known,
    saw(id, payloadSha256, replies) {
      learn(id, payloadSha256, replies ?? []);
    },
    async write(id, payload, fields) {
      // The caller's thread is the reader's own typed one; a stored thread
      // is read through the reader's parser.
      const mine = Array.isArray(payload.replies) ? (payload.replies as PdfAnnotationReply[]) : [];
      let thread = mergePdfAnnotationReplies(known(id), mine);
      let next = withThread(payload, thread);
      // A thread alone is not a row: its first write must be answered with
      // the stored one, whatever token this writer holds.
      if (fields === "theirs" && Object.keys(payload).every((key) => key === "replies")) {
        tokens.delete(id);
      }
      for (let stale = 0; ; stale += 1) {
        try {
          const answer = await deps.patch(id, {
            payload_json: next,
            expected_payload_sha256: tokens.get(id) ?? PDF_ANNOTATION_UNREAD_TOKEN,
          });
          learn(id, record(record(answer)?.annotation)?.payload_sha256, thread);
          const sent = new Set(mine.map((reply) => reply.id));
          return { payload: next, replies: thread, joined: thread.some((reply) => !sent.has(reply.id)) };
        } catch (error) {
          const stored = storedRowOfStaleWrite(error);
          if (!stored || stale >= MAX_STALE_WRITES) throw error;
          learn(id, stored.sha, deps.parseReplies(id, stored.payload.replies));
          thread = mergePdfAnnotationReplies(known(id), mine);
          // A stored payload the server could not read back comes as {}:
          // nothing to build on, so this writer's own fields are written.
          const kept = Object.keys(stored.payload).length > 0;
          const base = fields === "theirs" && kept
            ? stored.payload
            : fields === "place" && kept
              ? { ...stored.payload, page: payload.page, x: payload.x, y: payload.y }
              : next;
          next = withThread(base, thread);
        }
      }
    },
  };
}

/**
 * Replies a reader posted that their row does not hold yet, by file and
 * annotation. They are outside the reader so that one closed (or rebuilt)
 * over a failing write is not the end of them: the next reader of the file
 * in this tab shows them in their thread and writes them.
 */
const unsent = new Map<string, Map<string, PdfAnnotationReply[]>>();

export function keepPdfUnsentReply(file: string, annotationId: string, reply: PdfAnnotationReply): void {
  const threads = unsent.get(file) ?? new Map<string, PdfAnnotationReply[]>();
  unsent.set(file, threads);
  threads.set(annotationId, mergePdfAnnotationReplies(threads.get(annotationId) ?? [], [reply]));
}

export function pdfUnsentReplies(file: string, annotationId: string): readonly PdfAnnotationReply[] {
  return unsent.get(file)?.get(annotationId) ?? [];
}

/** The annotations of this file that have replies waiting to be written. */
export function pdfUnsentReplyThreads(file: string): string[] {
  return [...(unsent.get(file)?.keys() ?? [])];
}

/**
 * Take the replies that landed (or, with null, the whole waiting thread:
 * its annotation is gone) out of the waiting ones; what still waits comes back.
 */
export function settlePdfUnsentReplies(
  file: string,
  annotationId: string,
  landed: ReadonlySet<string> | null,
): readonly PdfAnnotationReply[] {
  const threads = unsent.get(file);
  const rest = landed ? (threads?.get(annotationId) ?? []).filter((reply) => !landed.has(reply.id)) : [];
  if (rest.length > 0) threads?.set(annotationId, rest);
  else {
    threads?.delete(annotationId);
    if (threads?.size === 0) unsent.delete(file);
  }
  return rest;
}

/** Tests only. */
export function resetPdfUnsentReplies(): void {
  unsent.clear();
}
