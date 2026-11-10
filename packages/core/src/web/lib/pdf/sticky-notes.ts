/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for PDF sticky-note (text-comment) annotations.
 *
 * A sticky note is a single point on a page (page-relative percentage
 * coordinates) carrying a short, user-authored text comment. This module holds
 * everything that has no DOM / React dependency so it can be unit-tested in
 * isolation and reused by both the viewer and (potentially) the backend mirror:
 *
 *   - the persisted annotation shape + its sanitizer (mirrors the highlight
 *     sanitizer in PDFViewer so a malformed/legacy payload never crashes the
 *     reader),
 *   - position clamping (drop coordinates onto a page, keeping the marker fully
 *     on-page),
 *   - hit-testing (which note marker, if any, is under a click point),
 *   - text normalization (collapse runaway whitespace, cap length — the text is
 *     ALWAYS rendered as plain text, never HTML, so this is normalization not
 *     escaping).
 *
 * The persisted annotation type is `pdf_note` and is additive to the existing
 * `pdf_highlight` type: old PDFs with only highlights keep loading unchanged.
 */

import {
  parseAnnotationAuthor,
  parseReplies,
  type PdfAnnotationAuthor,
  type PdfAnnotationReply,
} from "./annotation-replies";
import { normalizeXfdfSourceId } from "./xfdf-import";
import { isSafeHexColor } from "./safe-color";

/** Default marker fill used when a payload omits / supplies an unknown color. */
export const DEFAULT_NOTE_COLOR = "#fbbf24";

/**
 * Allowed marker colors. Kept as an allow-list (like the highlight colors) so a
 * crafted payload can't inject arbitrary CSS into the inline `style`.
 */
export const PDF_NOTE_COLORS: readonly string[] = [
  "#fbbf24", // amber
  "#60a5fa", // blue
  "#34d399", // green
  "#f87171", // red
  "#a78bfa", // purple
];

const PDF_NOTE_COLOR_SET = new Set(PDF_NOTE_COLORS);

/** Max length of a note's text after normalization. */
export const PDF_NOTE_MAX_TEXT = 2000;

/**
 * Marker size as a fraction of the smaller page dimension is handled in CSS; the
 * hit-test radius below is expressed in page-percentage units so it is
 * resolution-independent (matches how rects are stored 0..100).
 */
export const PDF_NOTE_HIT_RADIUS = 2.4;

export interface StickyNoteAnnotation {
  id: string;
  /** 1-based page number. */
  page: number;
  /** Page-relative position in percent [0..100]. */
  x: number;
  y: number;
  /** User-authored comment, already normalized; render as PLAIN TEXT only. */
  text: string;
  /** Marker fill; always one of {@link PDF_NOTE_COLORS}. */
  color: string;
  /** Server creation time retained for exact native compensation. */
  createdAt?: string;
  /**
   * Review status (Acrobat/WPS comment workflow): a resolved note is rendered
   * muted + struck-through in the panel and can be filtered out. Only a strict
   * `true` is persisted (an open note stores nothing), so an existing note is
   * byte-for-byte unchanged.
   */
  resolved?: boolean;
  /** Who created the note; absent on legacy notes created before authors. */
  author?: PdfAnnotationAuthor;
  /** Comment-thread replies (Acrobat / WPS PDF "Reply"); absent for none. */
  replies?: PdfAnnotationReply[];
  xfdfSourceId?: string;
}

function isFinitePercent(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value) && value >= 0 && value <= 100;
}

/**
 * Clamp a raw point so the marker stays fully on the page. `inset` is the
 * page-percentage half-size of the marker (so the icon never spills past an
 * edge). Non-finite inputs collapse to the nearest valid bound.
 */
export function clampNotePosition(
  x: number,
  y: number,
  inset = PDF_NOTE_HIT_RADIUS,
): { x: number; y: number } {
  const safeInset = Number.isFinite(inset) ? Math.max(0, Math.min(50, inset)) : 0;
  const lo = safeInset;
  const hi = 100 - safeInset;
  const clamp = (value: number) => {
    if (!Number.isFinite(value)) return lo;
    if (value < lo) return lo;
    if (value > hi) return hi;
    return value;
  };
  return { x: clamp(x), y: clamp(y) };
}

/**
 * Normalize user note text: trim, collapse internal runs of whitespace (while
 * preserving single newlines as line breaks), and hard-cap the length. This is
 * NOT HTML escaping — the renderer must place this string into a text node, not
 * `innerHTML`. Returns "" for non-string / empty input.
 */
export function normalizeNoteText(value: unknown): string {
  if (typeof value !== "string") return "";
  // Normalize CRLF/CR to LF, collapse spaces/tabs (not newlines) within a line,
  // then collapse 3+ blank lines to a single blank line. Finally trim ends.
  const unified = value.replace(/\r\n?/g, "\n");
  const collapsed = unified
    .split("\n")
    .map((line) => line.replace(/[^\S\n]+/g, " ").trim())
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
  return collapsed.slice(0, PDF_NOTE_MAX_TEXT);
}

/** True when the value is a palette color or any strict hex color, as for ink
 * and shapes: a note read or imported with its own color (black, from an
 * XFDF) kept it, instead of turning into the default amber (user audit PDF
 * 0019). Still never arbitrary CSS. */
export function isNoteColor(value: unknown): value is string {
  return (
    typeof value === "string" &&
    (PDF_NOTE_COLOR_SET.has(value) || isSafeHexColor(value))
  );
}

/**
 * Validate + normalize a persisted `pdf_note` payload into a usable annotation,
 * or return null when it is structurally invalid (mirrors
 * `sanitizeHighlightPayload`). Unknown/missing color falls back to the default;
 * an out-of-range position is rejected (returns null) rather than silently
 * snapped, so corrupt data is dropped rather than mis-rendered.
 */
export function sanitizeNotePayload(
  id: string,
  payload: Record<string, unknown> | null | undefined,
): StickyNoteAnnotation | null {
  if (!id || typeof id !== "string") return null;
  if (!payload || typeof payload !== "object") return null;
  const page = Number((payload as Record<string, unknown>).page);
  if (!Number.isInteger(page) || page < 1 || page > 10000) return null;
  const x = (payload as Record<string, unknown>).x;
  const y = (payload as Record<string, unknown>).y;
  if (!isFinitePercent(x) || !isFinitePercent(y)) return null;
  const rawText = (payload as Record<string, unknown>).text;
  if (typeof rawText === "string" && rawText.length > PDF_NOTE_MAX_TEXT * 4) {
    // Reject grossly oversized blobs outright (defense before normalization).
    return null;
  }
  const text = normalizeNoteText(rawText);
  const colorRaw = (payload as Record<string, unknown>).color;
  const color = isNoteColor(colorRaw) ? colorRaw : DEFAULT_NOTE_COLOR;
  const resolved = (payload as Record<string, unknown>).resolved === true;
  const author = parseAnnotationAuthor(
    (payload as Record<string, unknown>).author,
  );
  // Replies persisted by the client / server validator always carry id +
  // createdAt, so the fallbacks below only fire on hand-crafted payloads (a
  // deterministic per-index id keeps React keys stable across re-sanitizes).
  const replies = parseReplies(
    (payload as Record<string, unknown>).replies,
    (index) => `${id}:reply:${index}`,
    () => new Date().toISOString(),
  );
  const xfdfSourceId = normalizeXfdfSourceId(
    (payload as Record<string, unknown>).xfdfSourceId,
  );
  return {
    id,
    page,
    x,
    y,
    text,
    color,
    ...(resolved ? { resolved: true } : {}),
    ...(author ? { author } : {}),
    ...(replies.length > 0 ? { replies } : {}),
    ...(xfdfSourceId ? { xfdfSourceId } : {}),
  };
}

/**
 * Build the payload to PUT/POST for a new or edited sticky note. Position is
 * clamped and text normalized so what we persist matches what we render. Color
 * is forced onto the allow-list.
 */
export function buildNotePayload(input: {
  page: number;
  x: number;
  y: number;
  text: string;
  color?: string;
  resolved?: boolean;
  author?: PdfAnnotationAuthor;
  replies?: PdfAnnotationReply[];
  xfdfSourceId?: string;
}): {
  page: number;
  x: number;
  y: number;
  text: string;
  color: string;
  resolved?: boolean;
  author?: PdfAnnotationAuthor;
  replies?: PdfAnnotationReply[];
  xfdfSourceId?: string;
} {
  const { x, y } = clampNotePosition(input.x, input.y);
  const xfdfSourceId = normalizeXfdfSourceId(input.xfdfSourceId);
  return {
    page: input.page,
    x,
    y,
    text: normalizeNoteText(input.text),
    color: isNoteColor(input.color) ? input.color : DEFAULT_NOTE_COLOR,
    // Byte-clean: only an explicitly resolved note persists the flag.
    ...(input.resolved ? { resolved: true as const } : {}),
    // Byte-clean: author / replies keys only when present, so a legacy note
    // (or one from a user without an identity) persists exactly as before.
    ...(input.author ? { author: input.author } : {}),
    ...(input.replies && input.replies.length > 0
      ? { replies: input.replies }
      : {}),
    ...(xfdfSourceId ? { xfdfSourceId } : {}),
  };
}

/**
 * Find the topmost note marker under a page-relative click point, within
 * {@link PDF_NOTE_HIT_RADIUS}. Notes are tested in reverse order so the most
 * recently added marker (rendered last / on top) wins overlaps. Returns the
 * note's id, or null when nothing is close enough.
 */
export function hitTestNotes(
  notes: ReadonlyArray<Pick<StickyNoteAnnotation, "id" | "x" | "y">>,
  point: { x: number; y: number },
  radius = PDF_NOTE_HIT_RADIUS,
): string | null {
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) return null;
  const r = Number.isFinite(radius) ? Math.max(0, radius) : PDF_NOTE_HIT_RADIUS;
  const r2 = r * r;
  let bestId: string | null = null;
  let bestDist = Infinity;
  for (let i = notes.length - 1; i >= 0; i -= 1) {
    const note = notes[i];
    const dx = note.x - point.x;
    const dy = note.y - point.y;
    const dist = dx * dx + dy * dy;
    if (dist <= r2 && dist < bestDist) {
      bestDist = dist;
      bestId = note.id;
    }
  }
  return bestId;
}

/** Short, single-line preview of a note for tooltips / aria labels. */
export function notePreview(text: string, max = 80): string {
  const flat = normalizeNoteText(text).replace(/\n+/g, " ");
  if (flat.length <= max) return flat;
  return `${flat.slice(0, max).trimEnd()}…`;
}
