/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pdfNoteSourceKey, type PdfManagedNoteIdentityIndex } from "./managed-note-source-identity";
import {
  managedTextMarkupMatchesNative,
  type ManagedPdfTextMarkup,
} from "./native-text-markup-writeback";
import {
  MAX_AUTHOR_DISPLAY_NAME,
  MAX_AUTHOR_USER_ID,
  MAX_REPLIES_PER_ANNOTATION,
  MAX_REPLY_TEXT,
  type PdfAnnotationReply,
} from "./annotation-replies";

/**
 * Import native PDF `/Text` (sticky-note) annotations exposed by pdf.js.
 *
 * pdf.js gives each indirect annotation a stable object reference (`12R`,
 * `12R3`, ...). We keep that opaque reference in the editable page model so
 * the server can update the original annotation object in place instead of
 * appending a visually duplicated MRNote note.
 */

export const NATIVE_PDF_TEXT_SOURCE_KIND = "pdf_text_annotation";
export const NATIVE_PDF_TEXT_MARKUP_SOURCE_KIND = "pdf_text_markup_annotation";
export const MAX_NATIVE_PDF_TEXT_LENGTH = 32_000;
export const MAX_NATIVE_PDF_TEXT_ANNOTATIONS_PER_PAGE = 1_024;

const PDF_OBJECT_REF_RE = /^[1-9][0-9]{0,9}R(?:[0-9]{1,5})?$/;
const CONTROL_CHARS_RE = new RegExp(
  `[${String.fromCharCode(0)}-${String.fromCharCode(8)}` +
    `${String.fromCharCode(11)}-${String.fromCharCode(31)}` +
    `${String.fromCharCode(127)}]`,
  "g",
);
// PDF 32000-1 §7.9.4: D:YYYYMMDDHHmmSSOHH'mm' -- the trailing O is Z, + or -
// and gives the offset of local time from UT. The offset was ignored, so a
// note stamped 17:00+05'00' showed five hours after one stamped 12:00Z, the

const PDF_DATE_RE =
  /^D:(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?(?:([Zz+\-])(\d{2})?'?(\d{2})?'?)?/;

/**
 * The author's /F flags that decide what a reader may do with a source
 * annotation (PDF 32000-1 table 165). Carried on the source descriptor so
 * every surface (the paper, the list, the editor) reads one model.
 */
export interface NativePdfAnnotationFlags {
  /** Hidden (bit 2): not displayed, printed or interacted with. */
  hidden: boolean;
  /** NoView (bit 6): not displayed or interacted with on screen. */
  noView: boolean;
  /** ReadOnly (bit 7): no interaction with the user at all. */
  readOnly: boolean;
  /** Locked (bit 8): no move, resize or delete; contents stay editable. */
  locked: boolean;
  /** LockedContents (bit 10): the contents cannot be changed. */
  lockedContents: boolean;
  /**
   * NoZoom (bit 4): the appearance keeps its size on screen whatever the
   * reading zoom, anchored at its upper-left corner (ISO 32000-1 §12.5.3).
   */
  noZoom?: boolean;
  /** NoRotate (bit 5): the appearance does not follow the page rotation. */
  noRotate?: boolean;
}

/** The ISO date carried when the source annotation declares none. A fixed
 * sentinel, never "now": a note without a date must not look freshly
 * created on every open, nor sort differently from one open to the next. */
export const NATIVE_PDF_UNKNOWN_DATE_ISO = "1970-01-01T00:00:00.000Z";

export function isNativePdfUnknownDate(value: string | undefined): boolean {
  return value === NATIVE_PDF_UNKNOWN_DATE_ISO;
}

export function readNativePdfAnnotationFlags(
  value: unknown,
): NativePdfAnnotationFlags {
  const flags =
    typeof value === "number" && Number.isSafeInteger(value) && value >= 0
      ? value
      : 0;
  return {
    hidden: (flags & (1 << 1)) !== 0,
    noView: (flags & (1 << 5)) !== 0,
    readOnly: (flags & (1 << 6)) !== 0,
    locked: (flags & (1 << 7)) !== 0,
    lockedContents: (flags & (1 << 9)) !== 0,
    noZoom: (flags & (1 << 3)) !== 0,
    noRotate: (flags & (1 << 4)) !== 0,
  };
}

/** True when the annotation must not appear on screen. */
export function nativePdfAnnotationIsInvisible(
  flags: NativePdfAnnotationFlags | undefined,
): boolean {
  return Boolean(flags && (flags.hidden || flags.noView));
}

export interface NativePdfTextAnnotationSource {
  kind: typeof NATIVE_PDF_TEXT_SOURCE_KIND;
  objectRef: string;
  /** The author's protection and visibility flags. Absent on descriptors
   * written before they were read: treated as unprotected and visible. */
  flags?: NativePdfAnnotationFlags;
  /** The `/Name` icon of a Text annotation (Comment, Note, Help, ...). */
  icon?: string;
  /** The `/C` colour, as `#rrggbb`. */
  color?: string;
  /** False when the source carries no `/AP`: the reader draws the default
   * icon for `icon` itself. */
  hasAppearance?: boolean;
  /** True when the source declares no creation or modification date. */
  createdAtUnknown?: true;
  

  richText?: true;
  


  nestedReplies?: true;
  



  textTruncated?: true;
}

export type NativePdfTextMarkupSubtype =
  "Highlight" | "Underline" | "StrikeOut" | "Squiggly";

export interface NativePdfTextMarkupAnnotationSource {
  kind: typeof NATIVE_PDF_TEXT_MARKUP_SOURCE_KIND;
  objectRef: string;
  subtype: NativePdfTextMarkupSubtype;
  /** The author's protection and visibility flags (see the text source). */
  flags?: NativePdfAnnotationFlags;
  /** True when the source declares no creation or modification date. */
  createdAtUnknown?: true;
  /** True when the comment is longer than MAX_NATIVE_PDF_TEXT_LENGTH (see
   * the text source). */
  textTruncated?: true;
}

export type NativePdfAnnotationSource =
  NativePdfTextAnnotationSource | NativePdfTextMarkupAnnotationSource;

export interface NativePdfTextAnnotation {
  id: string;
  managedNoteId?: string;
  blockedManagedNoteId?: string;
  nativeIdentityRoundTripBlocked?: true;
  text: string;
  page: number;
  rect: [number, number, number, number];
  createdAt: string;
  author?: string;
  replies?: PdfAnnotationReply[];
  

  initiallyOpen?: true;
   
  reviewStates?: PdfAnnotationReply[];
  nativeSource: NativePdfTextAnnotationSource;
}

export interface NativePdfTextMarkupAnnotation {
  id: string;
  /** The mark MRAI made that this object is, by the name the file gives it
   * (`/NM`, lib/pdf/managed-note-source-identity): that mark's own id. Read
   * with the object, never kept on a row. */
  managedMarkupId?: string;
  text: string;
  page: number;
  rect: [number, number, number, number];
  createdAt: string;
  author?: string;
  color?: string;
  replies?: PdfAnnotationReply[];
  nativeSource: NativePdfTextMarkupAnnotationSource;
}

interface RawPdfString {
  str?: unknown;
}

interface RawNativeTextAnnotation {
  subtype?: unknown;
  id?: unknown;
  rect?: unknown;
  contentsObj?: unknown;
  titleObj?: unknown;
  creationDate?: unknown;
  modificationDate?: unknown;
  color?: unknown;
  inReplyTo?: unknown;
  replyType?: unknown;
  /** pdf.js `annotationFlags`: the annotation's /F. */
  annotationFlags?: unknown;
  /** pdf.js `name`: a Text annotation's /Name icon. */
  name?: unknown;
  /** pdf.js `hasAppearance`: whether the source carries an /AP. */
  hasAppearance?: unknown;
  /** pdf.js `richText`: the note's `/RC`, when it has one. */
  richText?: unknown;
  /** pdf.js Text `state` / `stateModel`: a review state record. */
  state?: unknown;
  stateModel?: unknown;
  /** pdf.js `popupRef`: the note's /Popup object, as "12R". */
  popupRef?: unknown;
  /** pdf.js (Popup annotations): the popup's /Open. */
  open?: unknown;
}

interface NativeTextMergeCandidate {
  id: string;
  nativeIdentityRoundTripBlocked?: true;
  text: string;
  page?: number;
  rect?: [number, number, number, number];
  createdAt: string;
  author?: string;
  replies?: PdfAnnotationReply[];
  nativeSource?: NativePdfAnnotationSource;
}

interface NativeMarkupMergeCandidate {
  id: string;
  text: string;
  page?: number;
  rect?: [number, number, number, number];
  createdAt: string;
  author?: string;
  color?: string;
  replies?: PdfAnnotationReply[];
  nativeSource?: NativePdfAnnotationSource;
  managedMarkup?: ManagedPdfTextMarkup;
}

export function isPdfObjectRef(value: unknown): value is string {
  return typeof value === "string" && PDF_OBJECT_REF_RE.test(value);
}

export function isNativePdfTextMarkupSubtype(
  value: unknown,
): value is NativePdfTextMarkupSubtype {
  return (
    value === "Highlight" ||
    value === "Underline" ||
    value === "StrikeOut" ||
    value === "Squiggly"
  );
}

/** Whether a note's text, read the way sanitizePdfText reads it, is longer
 * than the length the model keeps. */
function pdfTextExceeds(value: unknown, maxLength: number): boolean {
  return sanitizePdfText(value, maxLength + 1).length > maxLength;
}

export function sanitizePdfText(value: unknown, maxLength: number): string {
  const raw =
    typeof value === "string"
      ? value
      : value &&
          typeof value === "object" &&
          typeof (value as RawPdfString).str === "string"
        ? ((value as RawPdfString).str as string)
        : "";
  return raw
    .replace(/\r\n?/g, "\n")
    .replace(CONTROL_CHARS_RE, "")
    .slice(0, maxLength);
}

function normalizeRect(
  value: unknown,
): [number, number, number, number] | null {
  if (!Array.isArray(value) || value.length < 4) return null;
  const [a, b, c, d] = value;
  if (
    typeof a !== "number" ||
    !Number.isFinite(a) ||
    typeof b !== "number" ||
    !Number.isFinite(b) ||
    typeof c !== "number" ||
    !Number.isFinite(c) ||
    typeof d !== "number" ||
    !Number.isFinite(d)
  ) {
    return null;
  }
  const x1 = Math.min(a, c);
  const y1 = Math.min(b, d);
  const x2 = Math.max(a, c);
  const y2 = Math.max(b, d);
  if (x2 <= x1 || y2 <= y1) return null;
  return [x1, y1, x2, y2];
}

/** A PDF date string as an ISO instant, honouring its UT offset. */
export function parsePdfDateToIso(value: unknown, fallbackIso: string): string {
  return pdfDateToIso(value, fallbackIso);
}

function pdfDateToIso(value: unknown, fallbackIso: string): string {
  if (typeof value !== "string") return fallbackIso;
  const match = PDF_DATE_RE.exec(value);
  if (!match) {
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? fallbackIso : parsed.toISOString();
  }
  const year = Number(match[1]);
  const month = Number(match[2] ?? "01");
  const day = Number(match[3] ?? "01");
  const hour = Number(match[4] ?? "00");
  const minute = Number(match[5] ?? "00");
  const second = Number(match[6] ?? "00");
  const offsetSign = match[7];
  const offsetHours = Number(match[8] ?? "00");
  const offsetMinutes = Number(match[9] ?? "00");
  if (
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > 31 ||
    hour > 23 ||
    minute > 59 ||
    second > 59 ||
    offsetHours > 23 ||
    offsetMinutes > 59
  ) {
    return fallbackIso;
  }
  // Local time minus its offset from UT is UT. No offset given means the
  // relationship to UT is unknown (§7.9.4); it is read as UT, as before.
  const offsetSeconds =
    offsetSign === "+"
      ? (offsetHours * 60 + offsetMinutes) * 60
      : offsetSign === "-"
        ? -(offsetHours * 60 + offsetMinutes) * 60
        : 0;
  const parsed = new Date(
    Date.UTC(year, month - 1, day, hour, minute, second) - offsetSeconds * 1000,
  );
  return Number.isNaN(parsed.getTime()) ? fallbackIso : parsed.toISOString();
}

function normalizeNativePdfReplyThreads(
  rawAnnotations: readonly unknown[],
  fallbackIso: string,
): {
  replies: Map<string, PdfAnnotationReply[]>;
  

  states: Map<string, PdfAnnotationReply[]>;
  /** Roots whose discussion holds a reply to a reply. */
  nestedRoots: Set<string>;
} {
  const repliesByParent = new Map<string, PdfAnnotationReply[]>();
  const statesByParent = new Map<string, PdfAnnotationReply[]>();
  const nestedRoots = new Set<string>();
  // A reply to a reply (C -> B -> A) belongs to A's discussion: the card is
  // the root's, so every reply is attached to the root its /IRT chain leads
  // to. Grouping by direct parent alone dropped the second level and left
  
  const parentOf = new Map<string, string>();
  for (const raw of rawAnnotations) {
    if (!raw || typeof raw !== "object") continue;
    const annotation = raw as RawNativeTextAnnotation;
    if (
      annotation.replyType === "R" &&
      isPdfObjectRef(annotation.id) &&
      isPdfObjectRef(annotation.inReplyTo)
    ) {
      parentOf.set(annotation.id, annotation.inReplyTo);
    }
  }
  const rootOf = (id: string): string => {
    let current = id;
    for (let hops = 0; hops < 32; hops += 1) {
      const parent = parentOf.get(current);
      if (parent === undefined || parent === id) return current;
      current = parent;
    }
    return current;
  };
  for (const raw of rawAnnotations) {
    if (!raw || typeof raw !== "object") continue;
    const annotation = raw as RawNativeTextAnnotation;
    if (
      annotation.subtype !== "Text" ||
      annotation.replyType !== "R" ||
      !isPdfObjectRef(annotation.id) ||
      !isPdfObjectRef(annotation.inReplyTo)
    ) {
      continue;
    }
    const text = sanitizePdfText(annotation.contentsObj, MAX_REPLY_TEXT).trim();
    // A state record says who set which review state; it is kept as that,
    
    const stateModel: "Review" | "Marked" | null =
      annotation.stateModel === "Review" || annotation.stateModel === "Marked"
        ? annotation.stateModel
        : annotation.state != null
          ? "Review"
          : null;
    const reviewState =
      stateModel && typeof annotation.state === "string" && /^[A-Za-z]{1,32}$/u.test(annotation.state)
        ? { model: stateModel, state: annotation.state }
        : undefined;
    if (!text && !reviewState) continue;
    const root = rootOf(annotation.inReplyTo);
    if (root !== annotation.inReplyTo) nestedRoots.add(root);
    const target = text ? repliesByParent : statesByParent;
    const replies = target.get(root) ?? [];
    if (replies.length >= MAX_REPLIES_PER_ANNOTATION) continue;
    const displayName = sanitizePdfText(
      annotation.titleObj,
      MAX_AUTHOR_DISPLAY_NAME,
    ).trim();
    replies.push({
      // pdf.js exposes the indirect object identity used by /IRT. Keeping it
      // as the stable reply id lets patchback verify the existing object rather
      // than silently append a duplicate after reopen.
      id: annotation.id,
      text,
      createdAt: pdfDateToIso(
        annotation.creationDate ?? annotation.modificationDate,
        fallbackIso,
      ),
      ...(displayName
        ? {
            author: {
              userId: `pdf:${annotation.id}`.slice(0, MAX_AUTHOR_USER_ID),
              displayName,
            },
          }
        : {}),
      ...(reviewState ? { reviewState } : {}),
    });
    target.set(root, replies);
  }
  return { replies: repliesByParent, states: statesByParent, nestedRoots };
}

/** A reply's instant to the second: the page model keeps milliseconds, a
 * PDF date has none, so the same reply reads back a few hundred ms apart. */
function replySecond(iso: string): string {
  const time = Date.parse(iso);
  return Number.isFinite(time) ? String(Math.floor(time / 1000)) : iso;
}

function mergeNativePdfReplies(
  current: readonly PdfAnnotationReply[] | undefined,
  incoming: readonly PdfAnnotationReply[] | undefined,
): PdfAnnotationReply[] | undefined {
  if (!incoming?.length) return current as PdfAnnotationReply[] | undefined;
  const next = [...(current ?? [])];
  // A reply this window sent lives in the model under its own id and in the
  // PDF under an object number pdf.js cannot relate to it; read back, it is
  
  // twin -- same text, author and second -- so two real replies with the
  // same words stay two, and a twin is never invented across replies.
  const claimed = new Set<number>();
  let changed = false;
  for (const discovered of incoming) {
    const twin = next.findIndex(
      (reply, index) =>
        !claimed.has(index) &&
        (reply.id === discovered.id ||
          (reply.text === discovered.text &&
            replySecond(reply.createdAt) === replySecond(discovered.createdAt) &&
            reply.author?.displayName === discovered.author?.displayName)),
    );
    if (twin >= 0) {
      claimed.add(twin);
      continue;
    }
    if (next.length >= MAX_REPLIES_PER_ANNOTATION) break;
    next.push(discovered);
    changed = true;
  }
  return changed ? next : (current as PdfAnnotationReply[] | undefined);
}

export function sanitizeNativePdfTextAnnotationSource(
  value: unknown,
): NativePdfTextAnnotationSource | undefined {
  if (!value || typeof value !== "object") return undefined;
  const source = value as Record<string, unknown>;
  const objectRef = source.objectRef ?? source.object_ref;
  if (
    source.kind !== NATIVE_PDF_TEXT_SOURCE_KIND ||
    !isPdfObjectRef(objectRef)
  ) {
    return undefined;
  }
  return {
    kind: NATIVE_PDF_TEXT_SOURCE_KIND,
    objectRef,
  };
}

export function sanitizeNativePdfTextMarkupAnnotationSource(
  value: unknown,
): NativePdfTextMarkupAnnotationSource | undefined {
  if (!value || typeof value !== "object") return undefined;
  const source = value as Record<string, unknown>;
  const objectRef = source.objectRef ?? source.object_ref;
  const subtype = source.subtype;
  if (
    source.kind !== NATIVE_PDF_TEXT_MARKUP_SOURCE_KIND ||
    !isPdfObjectRef(objectRef) ||
    !isNativePdfTextMarkupSubtype(subtype)
  ) {
    return undefined;
  }
  return {
    kind: NATIVE_PDF_TEXT_MARKUP_SOURCE_KIND,
    objectRef,
    subtype,
  };
}

export function sanitizeNativePdfAnnotationSource(
  value: unknown,
): NativePdfAnnotationSource | undefined {
  return (
    sanitizeNativePdfTextAnnotationSource(value) ??
    sanitizeNativePdfTextMarkupAnnotationSource(value)
  );
}

export function sanitizeNativePdfMarkupColor(
  value: unknown,
): string | undefined {
  if (typeof value !== "string" || !/^#[0-9A-Fa-f]{6}$/.test(value)) {
    return undefined;
  }
  return value.toUpperCase();
}

function nativePdfColorToHex(value: unknown): string | undefined {
  if (
    !value ||
    typeof value !== "object" ||
    !("length" in value) ||
    typeof (value as { length?: unknown }).length !== "number"
  ) {
    return undefined;
  }
  const color = value as ArrayLike<unknown>;
  if (color.length < 3) return undefined;
  const channels = [color[0], color[1], color[2]];
  if (
    channels.some(
      (channel) =>
        typeof channel !== "number" ||
        !Number.isFinite(channel) ||
        channel < 0 ||
        channel > 255,
    )
  ) {
    return undefined;
  }
  return `#${channels
    .map((channel) =>
      Math.round(channel as number)
        .toString(16)
        .toUpperCase()
        .padStart(2, "0"),
    )
    .join("")}`;
}

/**
 * Normalize one page's raw pdf.js annotation list.
 *
 * Direct (non-indirect) annotations have synthetic `annot_*` ids and are kept
 * visible by pdf.js but deliberately omitted here: without a real PDF object
 * reference the server cannot safely address them for native writeback.
 */
export function normalizeNativePdfTextAnnotations(
  rawAnnotations: unknown,
  pageNumber: number,
  now: () => Date = () => new Date(),
  identities?: PdfManagedNoteIdentityIndex | null,
): NativePdfTextAnnotation[] {
  if (
    !Array.isArray(rawAnnotations) ||
    !Number.isInteger(pageNumber) ||
    pageNumber < 1
  ) {
    return [];
  }
  // `now` is kept for callers; the fallback is a fixed sentinel so a note
  // without a date neither looks freshly created nor re-sorts on each open.
  void now;
  const fallbackIso = NATIVE_PDF_UNKNOWN_DATE_ISO;
  const replyThreads = normalizeNativePdfReplyThreads(
    rawAnnotations,
    fallbackIso,
  );
  const normalized: NativePdfTextAnnotation[] = [];
  // Which popups the author left open, by their object reference.
  const openPopups = new Set<string>();
  for (const raw of rawAnnotations) {
    if (!raw || typeof raw !== "object") continue;
    const popup = raw as RawNativeTextAnnotation;
    if (popup.subtype === "Popup" && popup.open === true && isPdfObjectRef(popup.id)) {
      openPopups.add(popup.id);
    }
  }
  for (const raw of rawAnnotations) {
    if (normalized.length >= MAX_NATIVE_PDF_TEXT_ANNOTATIONS_PER_PAGE) break;
    if (!raw || typeof raw !== "object") continue;
    const annotation = raw as RawNativeTextAnnotation;
    if (
      annotation.subtype !== "Text" ||
      annotation.inReplyTo != null ||
      !isPdfObjectRef(annotation.id)
    ) {
      continue;
    }
    const text = sanitizePdfText(
      annotation.contentsObj,
      MAX_NATIVE_PDF_TEXT_LENGTH,
    ).trim();
    const rect = normalizeRect(annotation.rect);
    // An empty or blank note is still a note: it is drawn, and the user must
    
    // reference, never the text.
    if (!rect) continue;
    const author = sanitizePdfText(annotation.titleObj, 500).trim();
    const createdAt = pdfDateToIso(
      annotation.creationDate ?? annotation.modificationDate,
      fallbackIso,
    );
    const flags = readNativePdfAnnotationFlags(annotation.annotationFlags);
    const icon =
      typeof annotation.name === "string" && /^[A-Za-z]{1,32}$/.test(annotation.name)
        ? annotation.name
        : undefined;
    const noteColor = nativePdfColorToHex(annotation.color);
    const replies = replyThreads.replies.get(annotation.id);
    const reviewStates = replyThreads.states.get(annotation.id);
    const identity = identities?.get(pdfNoteSourceKey(pageNumber, annotation.id));
    const managedNoteId = identity?.kind === "managed" ? identity.id : null;
    normalized.push({
      id: managedNoteId ?? `pdf-native-text:${pageNumber}:${annotation.id}`,
      ...(managedNoteId ? { managedNoteId } : {}),
      ...(identity?.kind === "blocked" && identity.id ? { blockedManagedNoteId: identity.id } : {}),
      ...(identities === null || identity?.kind === "blocked" ? { nativeIdentityRoundTripBlocked: true as const } : {}),
      text,
      page: pageNumber,
      rect,
      createdAt,
      ...(author ? { author } : {}),
      ...(replies?.length ? { replies } : {}),
      ...(reviewStates?.length ? { reviewStates } : {}),
      ...(typeof annotation.popupRef === "string" &&
      openPopups.has(annotation.popupRef) &&
      !nativePdfAnnotationIsInvisible(flags)
        ? { initiallyOpen: true as const }
        : {}),
      nativeSource: {
        kind: NATIVE_PDF_TEXT_SOURCE_KIND,
        objectRef: annotation.id,
        flags,
        ...(icon ? { icon } : {}),
        ...(noteColor ? { color: noteColor } : {}),
        hasAppearance: annotation.hasAppearance !== false,
        ...(createdAt === fallbackIso ? { createdAtUnknown: true as const } : {}),
        ...(annotation.richText != null ? { richText: true as const } : {}),
        ...(replyThreads.nestedRoots.has(annotation.id)
          ? { nestedReplies: true as const }
          : {}),
        ...(pdfTextExceeds(annotation.contentsObj, MAX_NATIVE_PDF_TEXT_LENGTH)
          ? { textTruncated: true as const }
          : {}),
      },
    });
  }
  return normalized;
}

/**
 * Normalize native text-markup annotations without changing their geometry or
 * appearance. Empty `/Contents` values are intentionally retained: highlight
 * annotations often carry only `/QuadPoints`, and users must still be able to
 * identify, add a comment to, or delete those original PDF objects.
 */
export function normalizeNativePdfTextMarkupAnnotations(
  rawAnnotations: unknown,
  pageNumber: number,
  now: () => Date = () => new Date(),
): NativePdfTextMarkupAnnotation[] {
  if (
    !Array.isArray(rawAnnotations) ||
    !Number.isInteger(pageNumber) ||
    pageNumber < 1
  ) {
    return [];
  }
  // `now` is kept for callers; the fallback is a fixed sentinel so a note
  // without a date neither looks freshly created nor re-sorts on each open.
  void now;
  const fallbackIso = NATIVE_PDF_UNKNOWN_DATE_ISO;
  const replyThreads = normalizeNativePdfReplyThreads(
    rawAnnotations,
    fallbackIso,
  );
  const normalized: NativePdfTextMarkupAnnotation[] = [];
  for (const raw of rawAnnotations) {
    if (normalized.length >= MAX_NATIVE_PDF_TEXT_ANNOTATIONS_PER_PAGE) break;
    if (!raw || typeof raw !== "object") continue;
    const annotation = raw as RawNativeTextAnnotation;
    if (
      !isNativePdfTextMarkupSubtype(annotation.subtype) ||
      !isPdfObjectRef(annotation.id) ||
      // Acrobat-compatible replace-text proofing uses a grouped StrikeOut
      // child whose `/IRT` points at the Caret parent. It is one logical mark,
      // not an independent reader highlight. The managed parent descriptor is
      // reopened from the authenticated page model; importing this child would
      // duplicate the annotation and expose the internal group member.
      (annotation.replyType === "Group" &&
        isPdfObjectRef(annotation.inReplyTo))
    ) {
      continue;
    }
    const rect = normalizeRect(annotation.rect);
    if (!rect) continue;
    const text = sanitizePdfText(
      annotation.contentsObj,
      MAX_NATIVE_PDF_TEXT_LENGTH,
    ).trim();
    const author = sanitizePdfText(annotation.titleObj, 500).trim();
    const color = nativePdfColorToHex(annotation.color);
    const createdAt = pdfDateToIso(
      annotation.creationDate ?? annotation.modificationDate,
      fallbackIso,
    );
    const replies = replyThreads.replies.get(annotation.id);
    normalized.push({
      id: `pdf-native-markup:${pageNumber}:${annotation.id}`,
      text,
      page: pageNumber,
      rect,
      createdAt,
      ...(author ? { author } : {}),
      ...(color ? { color } : {}),
      ...(replies?.length ? { replies } : {}),
      nativeSource: {
        kind: NATIVE_PDF_TEXT_MARKUP_SOURCE_KIND,
        objectRef: annotation.id,
        subtype: annotation.subtype,
        flags: readNativePdfAnnotationFlags(annotation.annotationFlags),
        ...(createdAt === fallbackIso ? { createdAtUnknown: true as const } : {}),
        ...(pdfTextExceeds(annotation.contentsObj, MAX_NATIVE_PDF_TEXT_LENGTH)
          ? { textTruncated: true as const }
          : {}),
      },
    });
  }
  return normalized;
}

function sameNativeRect(
  left: readonly number[] | undefined,
  right: readonly number[],
): boolean {
  return (
    left !== undefined &&
    left.length === right.length &&
    left.every((value, index) => value === right[index])
  );
}

/**
 * Merge freshly discovered native notes into the host side-panel model.
 *
 * A text/page match upgrades an older MRNote V1 note with the real object
 * reference. This avoids showing the same already-synced note twice and
 * migrates later edits to exact in-place V2 targeting. If pikepdf renumbers an
 * object during save, the same match refreshes the stored reference.
 */
export function mergeNativePdfTextAnnotations<
  T extends NativeTextMergeCandidate,
>(current: readonly T[], incoming: readonly NativePdfTextAnnotation[], options?: {
  readonly managedIds?: ReadonlySet<string>;
  readonly deletedIds?: ReadonlySet<string>;
  readonly onIdentityCollision?: (id: string) => void;
}): T[] {
  if (incoming.length === 0) return current as T[];
  const next = [...current];
  let changed = false;
  const claimedIndexes = new Set<number>();

  for (const discovered of incoming) {
    if (options?.deletedIds?.has(discovered.id) || discovered.nativeIdentityRoundTripBlocked) continue;
    if (discovered.managedNoteId) {
      const match = next.findIndex(note => note.id === discovered.managedNoteId);
      const existing = match >= 0 ? next[match] : null;
      const conflictingSource = existing?.nativeSource && (
        existing.nativeSource.kind !== NATIVE_PDF_TEXT_SOURCE_KIND ||
        (existing.nativeSource.objectRef !== discovered.nativeSource.objectRef && incoming.some(other =>
          other !== discovered && other.nativeSource.objectRef === existing.nativeSource!.objectRef)));
      if (conflictingSource || (match >= 0 && claimedIndexes.has(match))) {
        // The PDF remains visible in the reader, but an ambiguous identity
        
        options?.onIdentityCollision?.(discovered.id);
        continue;
      }
      const managed: Omit<NativePdfTextAnnotation, "nativeSource"> & { nativeSource?: NativePdfTextAnnotationSource } = { ...discovered };
      delete managed.nativeSource;
      delete managed.managedNoteId;
      const replacement = existing ? { ...existing, rect: existing.rect ?? managed.rect, author: existing.author ?? managed.author,
        replies: existing.replies ?? managed.replies } : managed;
      // The PDF object reference belongs to this rendered version. A managed
      
      // file by a renumbered output object reference.
      delete (replacement as NativeTextMergeCandidate).nativeSource;
      if (existing) {
        claimedIndexes.add(match);
        if (existing.nativeSource || !existing.rect || (!existing.author && managed.author) || replacement.replies !== existing.replies) {
          next[match] = replacement as T;
          changed = true;
        }
      } else { next.push(replacement as T); claimedIndexes.add(next.length - 1); changed = true; }
      continue;
    }
    let matchIndex = next.findIndex(
      (annotation, index) =>
        !claimedIndexes.has(index) &&
        annotation.nativeSource?.kind === NATIVE_PDF_TEXT_SOURCE_KIND &&
        annotation.nativeSource.objectRef === discovered.nativeSource.objectRef,
    );
    // The text fallback upgrades an older MRNote note; an empty text never
    // identifies one (two empty notes on a page would take each other's row).
    if (
      matchIndex < 0 &&
      !discovered.nativeIdentityRoundTripBlocked &&
      discovered.text !== ""
    ) {
      matchIndex = next.findIndex(
        (annotation, index) =>
          !claimedIndexes.has(index) &&
          !options?.managedIds?.has(annotation.id) &&
          annotation.page === discovered.page &&
          (annotation.nativeSource == null ||
            annotation.nativeSource.kind === NATIVE_PDF_TEXT_SOURCE_KIND) &&
          annotation.text === discovered.text,
      );
    }
    if (matchIndex >= 0) {
      claimedIndexes.add(matchIndex);
      const existing = next[matchIndex];
      const replies = mergeNativePdfReplies(
        existing.replies,
        discovered.replies,
      );
      // The imported-note writer only ever rewrites /Contents, so the PDF's
      // own /Rect is the note's place: a page resize that moved it moves the
      // marker with it.
      const rect =
        existing.nativeSource?.kind === NATIVE_PDF_TEXT_SOURCE_KIND
          ? discovered.rect
          : (existing.rect ?? discovered.rect);
      if (
        existing.nativeSource?.objectRef !==
          discovered.nativeSource.objectRef ||
        !sameNativeRect(existing.rect, rect) ||
        (!existing.author && discovered.author) ||
        replies !== existing.replies ||
        (existing.nativeSource?.kind === NATIVE_PDF_TEXT_SOURCE_KIND &&
          (existing.nativeSource.nestedReplies !==
            discovered.nativeSource.nestedReplies ||
            // A row stored before the cut was recorded learns of it here,
            // before any save can send the cut text back.
            existing.nativeSource.textTruncated !==
              discovered.nativeSource.textTruncated))
      ) {
        next[matchIndex] = {
          ...existing,
          rect,
          author: existing.author ?? discovered.author,
          replies,
          nativeSource: discovered.nativeSource,
          ...(discovered.nativeIdentityRoundTripBlocked ? { nativeIdentityRoundTripBlocked: true as const } : {}),
        };
        changed = true;
      }
      continue;
    }
    next.push(discovered as T);
    claimedIndexes.add(next.length - 1);
    changed = true;
  }
  return changed ? next : (current as T[]);
}


















export function mergeNativePdfTextMarkupAnnotations<
  T extends NativeMarkupMergeCandidate,
>(
  current: readonly T[],
  incoming: readonly NativePdfTextMarkupAnnotation[],
  options?: { readonly ownManagedIds?: ReadonlySet<string> },
): T[] {
  if (incoming.length === 0) return current as T[];
  const next = [...current];
  let changed = false;
  const claimedIndexes = new Set<number>();

  for (const discovered of incoming) {
    const managedId = discovered.managedMarkupId;
    let matchIndex = managedId
      ? next.findIndex(
          (annotation, index) =>
            !claimedIndexes.has(index) &&
            annotation.id === managedId &&
            !!annotation.managedMarkup,
        )
      : -1;
    if (matchIndex < 0 && managedId && options?.ownManagedIds?.has(managedId)) {
      continue;
    }
    if (matchIndex < 0) {
      matchIndex = next.findIndex(
        (annotation, index) =>
          !claimedIndexes.has(index) &&
          annotation.nativeSource?.kind === NATIVE_PDF_TEXT_MARKUP_SOURCE_KIND &&
          annotation.nativeSource.objectRef === discovered.nativeSource.objectRef,
      );
    }
    if (matchIndex < 0) {
      matchIndex = next.findIndex(
        (annotation, index) =>
          !claimedIndexes.has(index) &&
          annotation.page === discovered.page &&
          annotation.nativeSource?.kind ===
            NATIVE_PDF_TEXT_MARKUP_SOURCE_KIND &&
          annotation.nativeSource.subtype === discovered.nativeSource.subtype &&
          annotation.rect?.every(
            (coordinate, coordinateIndex) =>
              Math.abs(coordinate - discovered.rect[coordinateIndex]) < 0.01,
          ),
      );
    }
    // (Only an object the file gives no such name is known by its looks.)
    if (matchIndex < 0 && !managedId) {
      matchIndex = next.findIndex(
        (annotation, index) =>
          !claimedIndexes.has(index) &&
          annotation.page === discovered.page &&
          !!annotation.managedMarkup &&
          managedTextMarkupMatchesNative(
            annotation.managedMarkup,
            {
              subtype: discovered.nativeSource.subtype,
              rect: discovered.rect,
              text: discovered.text,
              color: discovered.color,
            },
            {
              text: annotation.text,
              color: annotation.color,
            },
          ),
      );
    }
    if (matchIndex >= 0) {
      claimedIndexes.add(matchIndex);
      const existing = next[matchIndex];
      const replies = mergeNativePdfReplies(
        existing.replies,
        discovered.replies,
      );
      if (existing.managedMarkup) {
        // The reopened object was created from this managed descriptor. Keep
        // the stable MRNote id/geometry: original-package patchback always
        // replays from the immutable source, where the new object reference
        // does not exist.
        if (replies !== existing.replies) {
          next[matchIndex] = { ...existing, replies };
          changed = true;
        }
        continue;
      }
      // Imported markup edits keep the source geometry, so the PDF's /Rect
      // is authoritative (a page resize moves it).
      const rect =
        existing.nativeSource?.kind === NATIVE_PDF_TEXT_MARKUP_SOURCE_KIND
          ? discovered.rect
          : (existing.rect ?? discovered.rect);
      if (
        existing.nativeSource?.kind !== NATIVE_PDF_TEXT_MARKUP_SOURCE_KIND ||
        existing.nativeSource.objectRef !== discovered.nativeSource.objectRef ||
        !sameNativeRect(existing.rect, rect) ||
        (!existing.author && discovered.author) ||
        (!existing.color && discovered.color) ||
        replies !== existing.replies
      ) {
        next[matchIndex] = {
          ...existing,
          rect,
          author: existing.author ?? discovered.author,
          color: existing.color ?? discovered.color,
          replies,
          nativeSource: discovered.nativeSource,
        };
        changed = true;
      }
      continue;
    }
    const source: NativePdfTextMarkupAnnotation = { ...discovered };
    delete source.managedMarkupId;
    next.push(source as T);
    claimedIndexes.add(next.length - 1);
    changed = true;
  }
  return changed ? next : (current as T[]);
}
