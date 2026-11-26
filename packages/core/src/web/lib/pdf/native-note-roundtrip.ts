/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  pdfMarkupColorToHex,
  sanitizeManagedPdfProofingAnnotation,
  sanitizeManagedPdfTextMarkup,
  type ManagedPdfProofingAnnotation,
  type ManagedPdfTextMarkup,
} from "./native-text-markup-writeback";
import {
  MAX_AUTHOR_DISPLAY_NAME,
  MAX_AUTHOR_USER_ID,
  MAX_REPLIES_PER_ANNOTATION,
  MAX_REPLY_TEXT,
  type PdfAnnotationReply,
} from "./annotation-replies";

const NATIVE_NOTE_PREFIX = "[[MRNOTE_NATIVE_NOTE_V1|";
const NATIVE_NOTE_PATTERN =
  /^\[\[MRNOTE_NATIVE_NOTE_V1\|(upsert|delete)\|([A-Za-z0-9._~%\-]{1,384})\]\](?:\r?\n([\s\S]*))?$/;
const IMPORTED_NATIVE_TEXT_PREFIX = "[[MRNOTE_NATIVE_TEXT_V1|";
const IMPORTED_NATIVE_TEXT_PATTERN =
  /^\[\[MRNOTE_NATIVE_TEXT_V1\|(upsert|delete)\|([1-9][0-9]{0,9}R(?:[0-9]{1,5})?)\]\](?:\r?\n([\s\S]*))?$/;
const IMPORTED_NATIVE_MARKUP_PREFIX = "[[MRNOTE_NATIVE_MARKUP_V1|";
const IMPORTED_NATIVE_MARKUP_PATTERN =
  /^\[\[MRNOTE_NATIVE_MARKUP_V1\|(upsert|delete)\|([1-9][0-9]{0,9}R(?:[0-9]{1,5})?)\|(Highlight|Underline|StrikeOut|Squiggly)\]\](?:\r?\n([\s\S]*))?$/;
const IMPORTED_NATIVE_MARKUP_COLOR_PREFIX = "[[MRNOTE_NATIVE_MARKUP_V2|";
const IMPORTED_NATIVE_MARKUP_COLOR_PATTERN =
  /^\[\[MRNOTE_NATIVE_MARKUP_V2\|(upsert)\|([1-9][0-9]{0,9}R(?:[0-9]{1,5})?)\|(Highlight|Underline|StrikeOut|Squiggly)\|([0-9A-Fa-f]{6})\]\](?:\r?\n([\s\S]*))?$/;
const MANAGED_NATIVE_MARKUP_PREFIX = "[[MRNOTE_NATIVE_MARKUP_CREATE_V1|";
const MANAGED_NATIVE_MARKUP_PATTERN =
  /^\[\[MRNOTE_NATIVE_MARKUP_CREATE_V1\|(upsert)\|([A-Za-z0-9._~%\-]{1,384})\|(Highlight|Underline|StrikeOut|Squiggly)\|([0-9A-Fa-f]{6})\|([A-Za-z0-9._~%\-]{0,1500})\|([-0-9.,]{1,16384})\]\](?:\r?\n([\s\S]*))?$/;
const MANAGED_NATIVE_PROOFING_PREFIX = "[[MRNOTE_NATIVE_PROOFING_V1|";
const MANAGED_NATIVE_PROOFING_UPSERT_PATTERN =
  /^\[\[MRNOTE_NATIVE_PROOFING_V1\|(upsert)\|([A-Za-z0-9._~%\-]{1,384})\|(replace|insert)\|([0-9A-Fa-f]{6})\|([A-Za-z0-9._~%\-]{0,1500})\|(none|[-0-9.,]{1,16384})\]\](?:\r?\n([\s\S]*))?$/;
const MANAGED_NATIVE_PROOFING_DELETE_PATTERN =
  /^\[\[MRNOTE_NATIVE_PROOFING_V1\|(delete)\|([A-Za-z0-9._~%\-]{1,384})\]\]$/;
const PDF_OBJECT_REF_PATTERN = /^[1-9][0-9]{0,9}R(?:[0-9]{1,5})?$/;
const PDF_MARKUP_COLOR_PATTERN = /^#[0-9A-Fa-f]{6}$/;
const CANONICAL_REPLY_TIMESTAMP_PATTERN =
  /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})\.(\d{3})Z$/;

export const MAX_NATIVE_NOTE_TOMBSTONES = 500;

function isCanonicalReplyTimestamp(value: string): boolean {
  const match = CANONICAL_REPLY_TIMESTAMP_PATTERN.exec(value);
  if (!match) return false;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = Number(match[4]);
  const minute = Number(match[5]);
  const second = Number(match[6]);
  const millisecond = Number(match[7]);
  if (year < 1) return false;
  const parsed = new Date(0);
  parsed.setUTCFullYear(year, month - 1, day);
  parsed.setUTCHours(hour, minute, second, millisecond);
  return (
    parsed.getUTCFullYear() === year &&
    parsed.getUTCMonth() === month - 1 &&
    parsed.getUTCDate() === day &&
    parsed.getUTCHours() === hour &&
    parsed.getUTCMinutes() === minute &&
    parsed.getUTCSeconds() === second &&
    parsed.getUTCMilliseconds() === millisecond
  );
}

export interface NativePdfTextAnnotationSource {
  kind: "pdf_text_annotation";
  objectRef: string;
  /** `text` is only the beginning of the source's `/Contents`. */
  textTruncated?: true;
}

export interface NativePdfTextMarkupAnnotationSource {
  kind: "pdf_text_markup_annotation";
  objectRef: string;
  subtype: "Highlight" | "Underline" | "StrikeOut" | "Squiggly";
  /** `text` is only the beginning of the source's `/Contents`. */
  textTruncated?: true;
}

export type NativePdfAnnotationSource =
  NativePdfTextAnnotationSource | NativePdfTextMarkupAnnotationSource;

export interface NativePdfPageAnnotation {
  id: string;
  text: string;
  color?: string;
  quote?: string;
  page?: number;
  rect?: [number, number, number, number];
  createdAt: string;
  author?: string;
  replies?: unknown[];
  /** Stored reply metadata was present but could not be preserved exactly. */
  nativeReplyRoundTripBlocked?: true;
  nativeSource?: NativePdfAnnotationSource;
  managedMarkup?: ManagedPdfTextMarkup;
  managedProofing?: ManagedPdfProofingAnnotation;
}

export type NativeNoteReplyAuthority =
  | { state: "absent" }
  | { state: "complete"; replies: PdfAnnotationReply[] }
  | { state: "malformed" };

export interface NativeNotePartition<T extends NativePdfPageAnnotation> {
  notes: T[];
  tombstones: T[];
}

export function nativeNoteIdToken(id: string): string | null {
  const trimmed = id.trim();
  if (!trimmed) return null;
  const token = encodeURIComponent(trimmed).replace(
    /[!'()*]/g,
    (character) =>
      `%${character.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")}`,
  );
  return token.length <= 384 ? token : null;
}

function nativeAnnotationObjectRef(
  annotation: Pick<NativePdfPageAnnotation, "nativeSource">,
): string | null {
  const source = annotation.nativeSource;
  if (
    (source?.kind !== "pdf_text_annotation" &&
      source?.kind !== "pdf_text_markup_annotation") ||
    !PDF_OBJECT_REF_PATTERN.test(source.objectRef)
  ) {
    return null;
  }
  return source.objectRef;
}

export function isNativeNoteRoundTripEligible(
  annotation: Pick<
    NativePdfPageAnnotation,
    | "id"
    | "nativeSource"
    | "managedMarkup"
    | "managedProofing"
    | "replies"
    | "nativeReplyRoundTripBlocked"
  >,
): boolean {
  if (
    annotation.nativeReplyRoundTripBlocked === true ||
    nativeNoteRepliesForRoundTrip(annotation.replies) === null
  ) {
    return false;
  }
  if (annotation.managedMarkup !== undefined) {
    return (
      sanitizeManagedPdfTextMarkup(annotation.managedMarkup) !== undefined &&
      nativeNoteIdToken(annotation.id) !== null
    );
  }
  if (annotation.managedProofing !== undefined) {
    return (
      sanitizeManagedPdfProofingAnnotation(annotation.managedProofing) !==
        undefined &&
      nativeNoteIdToken(annotation.id) !== null &&
      nativeNoteRepliesForRoundTrip(annotation.replies)?.length === 0
    );
  }
  return (
    nativeAnnotationObjectRef(annotation) !== null ||
    nativeNoteIdToken(annotation.id) !== null
  );
}

/**
 * Preserve the reply payload exactly or refuse the native save. The general UI
 * parser is intentionally tolerant for legacy snapshots; the original-package
 * writer cannot be tolerant because truncating or backfilling a reply would
 * silently change a discussion thread while reporting a successful save.
 */
export function nativeNoteRepliesForRoundTrip(
  raw: unknown,
): PdfAnnotationReply[] | null {
  if (raw === undefined) return [];
  if (!Array.isArray(raw) || raw.length > MAX_REPLIES_PER_ANNOTATION) {
    return null;
  }
  const replies: PdfAnnotationReply[] = [];
  for (const item of raw) {
    if (!item || typeof item !== "object") return null;
    const reply = item as Record<string, unknown>;
    // A review state record (Completed, None...: ISO 32000-1 table 172) is a
    // reply whose words may be empty; the writer gives it /State and
    // /StateModel (apps/api native_edit_primitives).
    const reviewState = reply.reviewState === undefined ? undefined : readPdfReplyReviewState(reply.reviewState);
    if (reviewState === null) return null;
    if (
      typeof reply.id !== "string" ||
      !reply.id.trim() ||
      reply.id.length > 384 ||
      typeof reply.text !== "string" ||
      (!reply.text.trim() && !reviewState) ||
      reply.text.length > MAX_REPLY_TEXT ||
      typeof reply.createdAt !== "string" ||
      !reply.createdAt ||
      reply.createdAt.length > 64 ||
      !isCanonicalReplyTimestamp(reply.createdAt)
    ) {
      return null;
    }
    let author: PdfAnnotationReply["author"];
    if (reply.author !== undefined) {
      if (!reply.author || typeof reply.author !== "object") return null;
      const rawAuthor = reply.author as Record<string, unknown>;
      if (
        typeof rawAuthor.userId !== "string" ||
        !rawAuthor.userId ||
        rawAuthor.userId.length > MAX_AUTHOR_USER_ID ||
        typeof rawAuthor.displayName !== "string" ||
        !rawAuthor.displayName.trim() ||
        rawAuthor.displayName.length > MAX_AUTHOR_DISPLAY_NAME
      ) {
        return null;
      }
      author = {
        userId: rawAuthor.userId,
        displayName: rawAuthor.displayName,
      };
    }
    replies.push({
      id: reply.id,
      text: reply.text,
      createdAt: reply.createdAt,
      ...(author ? { author } : {}),
      ...(reviewState ? { reviewState } : {}),
    });
  }
  return replies;
}

const PDF_REPLY_REVIEW_STATES: Readonly<Record<string, readonly string[]>> = {
  Review: ["Accepted", "Rejected", "Cancelled", "Completed", "None"],
  Marked: ["Marked", "Unmarked"],
};

/** A stored reply's review state, or null when it is not one the writer
 * takes (the save is refused rather than the state dropped). */
export function readPdfReplyReviewState(
  raw: unknown,
): PdfAnnotationReply["reviewState"] | null {
  if (!raw || typeof raw !== "object") return null;
  const value = raw as Record<string, unknown>;
  if (Object.keys(value).some((key) => key !== "model" && key !== "state")) return null;
  const { model, state } = value;
  if ((model !== "Review" && model !== "Marked") || typeof state !== "string") return null;
  if (!PDF_REPLY_REVIEW_STATES[model]?.includes(state)) return null;
  return { model, state };
}

/** The review states that close a PDF discussion (Acrobat's Accepted / Completed). */
export const PDF_RESOLVING_REVIEW_STATES: ReadonlySet<string> = new Set(["Accepted", "Completed"]);

/** Whether a note's latest review state record (a reply carrying one)
 * closes its discussion. */
export function pdfNoteRepliesResolved(replies: readonly PdfAnnotationReply[] | undefined): boolean {
  let latest: string | undefined;
  for (const reply of replies ?? []) if (reply.reviewState) latest = reply.reviewState.state;
  return latest !== undefined && PDF_RESOLVING_REVIEW_STATES.has(latest);
}

/**
 * Retain the three stored-thread states. In particular, an absent property is
 * not an authoritative empty thread, while a malformed present property must
 * block native round-trip instead of being normalized to `[]`.
 */
export function readNativeNoteReplyAuthority(annotation: {
  replies?: unknown;
}): NativeNoteReplyAuthority {
  if (!Object.prototype.hasOwnProperty.call(annotation, "replies")) {
    return { state: "absent" };
  }
  if (annotation.replies === undefined) {
    return { state: "malformed" };
  }
  const replies = nativeNoteRepliesForRoundTrip(annotation.replies);
  return replies === null
    ? { state: "malformed" }
    : { state: "complete", replies };
}

function nativeTextToken(value: string | undefined): string {
  if (!value?.trim()) return "";
  const token = encodeURIComponent(value.trim()).replace(
    /[!'()*]/g,
    (character) =>
      `%${character.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")}`,
  );
  return token.length <= 1500 ? token : "";
}

function decodeNativeTextToken(value: string): string | undefined {
  if (!value) return undefined;
  try {
    const decoded = decodeURIComponent(value).trim();
    return decoded ? decoded.slice(0, 500) : undefined;
  } catch {
    return undefined;
  }
}

function pdfCoordinateToken(value: number): string {
  return Number(value.toFixed(6)).toString();
}

export function encodeManagedNativeNote<T extends NativePdfPageAnnotation>(
  annotation: T,
): T {
  const managedProofing = sanitizeManagedPdfProofingAnnotation(
    annotation.managedProofing,
  );
  if (managedProofing) {
    const token = nativeNoteIdToken(annotation.id);
    if (!token || !annotation.text.trim()) return annotation;
    const color = pdfMarkupColorToHex(annotation.color, "#E85D75").slice(1);
    const author = nativeTextToken(annotation.author);
    const quadPoints = managedProofing.quadPoints
      ? managedProofing.quadPoints.map(pdfCoordinateToken).join(",")
      : "none";
    return {
      ...annotation,
      color: `#${color}`,
      quote: undefined,
      managedProofing: {
        ...managedProofing,
        quote: annotation.quote ?? managedProofing.quote,
      },
      text:
        `${MANAGED_NATIVE_PROOFING_PREFIX}upsert|${token}|` +
        `${managedProofing.kind}|${color}|${author}|${quadPoints}]]\n` +
        annotation.text,
    };
  }
  if (annotation.managedProofing !== undefined) return annotation;
  const managedMarkup = sanitizeManagedPdfTextMarkup(annotation.managedMarkup);
  if (managedMarkup) {
    const token = nativeNoteIdToken(annotation.id);
    if (!token) return annotation;
    const color = pdfMarkupColorToHex(annotation.color).slice(1);
    const author = nativeTextToken(annotation.author);
    const quadPoints = managedMarkup.quadPoints
      .map(pdfCoordinateToken)
      .join(",");
    return {
      ...annotation,
      color: `#${color}`,
      // `_pdf_annotation_contents` appends `quote` to ordinary notes. A native
      // text markup already carries exact QuadPoints, so keep the selected text
      // only inside the managed descriptor and write `/Contents` as comment-only.
      quote: undefined,
      managedMarkup: {
        ...managedMarkup,
        quote: annotation.quote ?? managedMarkup.quote,
      },
      text:
        `${MANAGED_NATIVE_MARKUP_PREFIX}upsert|${token}|` +
        `${managedMarkup.subtype}|${color}|${author}|${quadPoints}]]\n` +
        annotation.text,
    };
  }
  if (annotation.managedMarkup !== undefined) {
    // Invalid managed geometry must remain visibly unsupported; never turn it
    // into a superficially successful `/Text` note.
    return annotation;
  }
  const objectRef = nativeAnnotationObjectRef(annotation);
  if (objectRef) {
    if (annotation.nativeSource?.kind === "pdf_text_markup_annotation") {
      const color = PDF_MARKUP_COLOR_PATTERN.test(annotation.color ?? "")
        ? annotation.color!.slice(1).toUpperCase()
        : null;
      return {
        ...annotation,
        text: color
          ? `${IMPORTED_NATIVE_MARKUP_COLOR_PREFIX}upsert|${objectRef}|${annotation.nativeSource.subtype}|${color}]]\n${annotation.text}`
          : `${IMPORTED_NATIVE_MARKUP_PREFIX}upsert|${objectRef}|${annotation.nativeSource.subtype}]]\n${annotation.text}`,
      };
    }
    return {
      ...annotation,
      text: `${IMPORTED_NATIVE_TEXT_PREFIX}upsert|${objectRef}]]\n${annotation.text}`,
    };
  }
  const token = nativeNoteIdToken(annotation.id);
  if (!token) return annotation;
  return {
    ...annotation,
    text: `${NATIVE_NOTE_PREFIX}upsert|${token}]]\n${annotation.text}`,
  };
}

export function createManagedNativeNoteTombstone<
  T extends NativePdfPageAnnotation,
>(annotation: T): T | null {
  const objectRef = nativeAnnotationObjectRef(annotation);
  const token = annotation.managedProofing
    ? nativeNoteIdToken(annotation.id)
    : (objectRef ?? nativeNoteIdToken(annotation.id));
  if (!token) return null;
  const tombstone = {
    ...annotation,
    text:
      annotation.managedProofing
        ? `${MANAGED_NATIVE_PROOFING_PREFIX}delete|${token}]]`
        : annotation.nativeSource?.kind === "pdf_text_markup_annotation"
        ? `${IMPORTED_NATIVE_MARKUP_PREFIX}delete|${token}|${annotation.nativeSource.subtype}]]`
        : objectRef
          ? `${IMPORTED_NATIVE_TEXT_PREFIX}delete|${token}]]`
          : `${NATIVE_NOTE_PREFIX}delete|${token}]]`,
  };
  delete tombstone.quote;
  delete tombstone.replies;
  return tombstone;
}

export function partitionManagedNativeNotes<T extends NativePdfPageAnnotation>(
  annotations: readonly T[],
): NativeNotePartition<T> {
  const notes: T[] = [];
  const tombstones: T[] = [];
  for (const annotation of annotations) {
    const proofingDeleteMatch = MANAGED_NATIVE_PROOFING_DELETE_PATTERN.exec(
      annotation.text,
    );
    if (proofingDeleteMatch) {
      tombstones.push(annotation);
      continue;
    }
    const proofingMatch = MANAGED_NATIVE_PROOFING_UPSERT_PATTERN.exec(
      annotation.text,
    );
    if (proofingMatch) {
      const storedProofing = sanitizeManagedPdfProofingAnnotation(
        annotation.managedProofing,
      );
      const quadPoints =
        proofingMatch[6] === "none"
          ? undefined
          : proofingMatch[6].split(",").map(Number);
      const parsedProofing = sanitizeManagedPdfProofingAnnotation({
        kind: proofingMatch[3],
        rect: annotation.rect ?? storedProofing?.rect,
        ...(quadPoints ? { quadPoints } : {}),
        quote: storedProofing?.quote,
      });
      if (!parsedProofing) {
        notes.push(annotation);
        continue;
      }
      const reopened = {
        ...annotation,
        text: proofingMatch[7] ?? "",
        color: `#${proofingMatch[4].toUpperCase()}`,
        managedProofing: parsedProofing,
      };
      const author = decodeNativeTextToken(proofingMatch[5]);
      if (parsedProofing.quote) reopened.quote = parsedProofing.quote;
      else delete reopened.quote;
      if (author) reopened.author = author;
      notes.push(reopened);
      continue;
    }
    const managedMarkupMatch = MANAGED_NATIVE_MARKUP_PATTERN.exec(
      annotation.text,
    );
    if (managedMarkupMatch) {
      const storedMarkup = sanitizeManagedPdfTextMarkup(
        annotation.managedMarkup,
      );
      const quadPoints = managedMarkupMatch[6].split(",").map(Number);
      const parsedMarkup = sanitizeManagedPdfTextMarkup({
        subtype: managedMarkupMatch[3],
        rect: annotation.rect ?? storedMarkup?.rect,
        quadPoints,
        quote: storedMarkup?.quote,
        hasReaderOnlyFeatures: storedMarkup?.hasReaderOnlyFeatures,
      });
      if (!parsedMarkup) {
        notes.push(annotation);
        continue;
      }
      const reopened = {
        ...annotation,
        text: managedMarkupMatch[7] ?? "",
        color: `#${managedMarkupMatch[4].toUpperCase()}`,
        managedMarkup: parsedMarkup,
      };
      const author = decodeNativeTextToken(managedMarkupMatch[5]);
      if (parsedMarkup.quote) reopened.quote = parsedMarkup.quote;
      else delete reopened.quote;
      if (author) {
        reopened.author = author;
      }
      notes.push(reopened);
      continue;
    }
    const markupColorMatch = IMPORTED_NATIVE_MARKUP_COLOR_PATTERN.exec(
      annotation.text,
    );
    const markupMatch =
      markupColorMatch ?? IMPORTED_NATIVE_MARKUP_PATTERN.exec(annotation.text);
    const importedMatch = IMPORTED_NATIVE_TEXT_PATTERN.exec(annotation.text);
    const match =
      markupMatch ?? importedMatch ?? NATIVE_NOTE_PATTERN.exec(annotation.text);
    if (!match) {
      notes.push(annotation);
      continue;
    }
    if (match[1] === "delete") {
      tombstones.push(annotation);
      continue;
    }
    notes.push({
      ...annotation,
      text: markupColorMatch
        ? (match[5] ?? "")
        : markupMatch
          ? (match[4] ?? "")
          : (match[3] ?? ""),
      ...(markupMatch
        ? {
            ...(markupColorMatch
              ? { color: `#${match[4].toUpperCase()}` }
              : {}),
            nativeSource: sameSourceKept(annotation.nativeSource, {
              kind: "pdf_text_markup_annotation",
              objectRef: match[2],
              subtype: match[3] as
                "Highlight" | "Underline" | "StrikeOut" | "Squiggly",
            }),
          }
        : importedMatch
          ? {
              nativeSource: sameSourceKept(annotation.nativeSource, {
                kind: "pdf_text_annotation",
                objectRef: match[2],
              }),
            }
          : {}),
    });
  }
  return { notes, tombstones };
}













function sameSourceKept<S extends NativePdfAnnotationSource>(
  stored: NativePdfAnnotationSource | undefined,
  named: S,
): S {
  if (!stored || stored.kind !== named.kind || stored.objectRef !== named.objectRef) {
    return named;
  }
  const { textTruncated: _truncated, ...kept } = stored;
  void _truncated;
  return { ...kept, ...named } as S;
}

export function serializeManagedNativeNotes<T extends NativePdfPageAnnotation>(
  notes: readonly T[],
  tombstones: readonly T[],
): T[] {
  const latestTombstoneById = new Map<string, T>();
  for (const tombstone of tombstones) {
    latestTombstoneById.set(tombstone.id, tombstone);
  }
  for (const note of notes) {
    latestTombstoneById.delete(note.id);
  }
  return [
    // A source note whose text the model holds only the beginning of is not
    // written back: its upsert would replace the whole original with the cut
    // text, and it did, whenever any other note on the page was saved
    
    // deletion is a tombstone and still goes.
    ...notes
      .filter((note) => !note.nativeSource?.textTruncated)
      .map((note) => encodeManagedNativeNote(note)),
    ...latestTombstoneById.values(),
  ];
}

export function appendManagedNativeNoteTombstone<
  T extends NativePdfPageAnnotation,
>(tombstones: readonly T[], annotation: T): T[] | null {
  const tombstone = createManagedNativeNoteTombstone(annotation);
  if (!tombstone) return null;
  const retained = tombstones.filter((entry) => entry.id !== annotation.id);
  if (retained.length >= MAX_NATIVE_NOTE_TOMBSTONES) return null;
  return [...retained, tombstone];
}
