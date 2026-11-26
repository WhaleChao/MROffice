/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for the PDF viewer's CONSOLIDATED annotations / comments list panel
 * (Acrobat's "Comments" pane, WPS's 批注 list).
 *
 * The viewer keeps each annotation kind in its own state array (highlights,
 * sticky notes, ink strokes, shapes, user bookmarks), each persisted per
 * `dataItemId`. This module folds those heterogeneous arrays into a SINGLE flat,
 * uniformly-shaped list ({@link AnnotationListItem}) so one panel can:
 *
 *   - list every annotation across every page,
 *   - filter by kind,
 *   - sort by page (then a stable secondary order),
 *   - export the whole set as JSON or a human-readable text summary.
 *
 * Everything here is pure (no DOM / React / pdf.js) so it can be unit-tested in
 * isolation and reused by the viewer (and a potential backend mirror). It mirrors
 * the style of `sticky-notes.ts` / `view-modes.ts`: small, defensive, total
 * functions that never throw on malformed input.
 *
 * Redactions are deliberately EXCLUDED — they are destructive content removal,
 * not a comment/markup, and Acrobat likewise keeps them out of the Comments
 * pane. Signatures are also excluded (they are a placed visual mark, surfaced via
 * their own affordance, not a comment).
 */

/** The annotation kinds the consolidated list can hold, in display order. */
export const ANNOTATION_LIST_KINDS = [
  "highlight",
  "note",
  "ink",
  "shape",
  "bookmark",
] as const;

export type AnnotationListKind = (typeof ANNOTATION_LIST_KINDS)[number];

/** Whether a value is a recognised consolidated-list annotation kind. */
export function isAnnotationListKind(
  value: unknown,
): value is AnnotationListKind {
  return (
    typeof value === "string" &&
    (ANNOTATION_LIST_KINDS as readonly string[]).includes(value)
  );
}

/**
 * A single row in the consolidated annotations list — a kind-tagged, uniformly
 * shaped projection of one underlying annotation.
 *
 *  - `id`      the underlying annotation id (used to jump / dedupe).
 *  - `kind`    which annotation kind this row came from.
 *  - `page`    1-based page the annotation lives on (the jump target).
 *  - `preview` a short human-readable label for the row (already plain text —
 *              callers render it into a text node, never innerHTML).
 *  - `color`   an optional swatch color (highlight / note / ink / shape) — a
 *              passthrough of the annotation's own already-sanitized color, or
 *              undefined when the kind has none (bookmark).
 *  - `order`   the item's position in the SOURCE array of its kind, so a stable
 *              secondary sort can preserve creation order within a page.
 */
/** Which text markup a "highlight" row is: the panel's Text Markup group
 * holds all four, and each row says its own (user audit PDF 0002: an
 * Underline was listed and exported as a Highlight). */
export type AnnotationListMarkup = "highlight" | "underline" | "strikeout" | "squiggly";

export interface AnnotationListItem {
  id: string;
  kind: AnnotationListKind;
  /** For a text markup row, the markup it is. */
  markup?: AnnotationListMarkup;
  page: number;
  /** On a page inserted in MR Office (`page` is then the page before it):
   * lib/pdf/inserted-page-identity. */
  insertedPageId?: string;
  preview: string;
  





  text?: string;
  color?: string;
  order: number;
  /** Creator display name (highlight / note), when the annotation has one. */
  authorName?: string;
  /** Stable creator id used for filtering when present. */
  authorId?: string;
  /** MRAI note workflow state; absent for non-note annotation families. */
  reviewState?: "active" | "resolved";
  /** Comment-thread replies (highlight / note), previewed + plain text. */
  replies?: AnnotationListReply[];
  





  origin?: "pdf";
  /** The source's own flags forbid changing it (ReadOnly / Locked...). */
  readOnly?: true;
}

export type AnnotationListBatchDeleteTarget = Pick<
  AnnotationListItem,
  "id" | "kind"
>;

/** One reply row shown under an expandable list item (already plain text). */
export interface AnnotationListReply {
  id: string;
  /** Capped, whitespace-collapsed label for the panel row. */
  text: string;
   
  fullText: string;
  authorName?: string;
  /** Present when the reply is its author's review state record: the row
   * shows the state, and `text` keeps the producer's own wording (empty, or
   * Acrobat's "<State> set by <author>"). */
  reviewState?: { model: "Review" | "Marked"; state: string };
}

/** Source-PDF rows sort after the reader's own on a page, in source order. */
const NATIVE_ORDER_OFFSET = 1_000_000;

/** Max characters kept for a row's preview label (longer text is ellipsised). */
export const ANNOTATION_PREVIEW_MAX = 80;

/**
 * Collapse whitespace, strip leading/trailing space, and hard-cap a preview
 * string to {@link ANNOTATION_PREVIEW_MAX} characters (appending an ellipsis when
 * truncated). NOT HTML escaping — the renderer must place the result into a text
 * node. Returns "" for non-string / empty input.
 */
export function annotationPreview(
  value: unknown,
  max = ANNOTATION_PREVIEW_MAX,
): string {
  if (typeof value !== "string") return "";
  const flat = value.replace(/\s+/g, " ").trim();
  const cap = Number.isInteger(max) && max > 0 ? max : ANNOTATION_PREVIEW_MAX;
  if (flat.length <= cap) return flat;
  return `${flat.slice(0, cap).trimEnd()}…`;
}

/**
 * The complete body of an annotation for export: control characters removed,
 * CR/CRLF folded to LF, surrounding whitespace trimmed, inner line breaks and
 * spacing kept. "" for non-string / empty input.
 */
export function annotationFullText(value: unknown): string {
  if (typeof value !== "string") return "";
  return value
    .replace(/\r\n?/g, "\n")
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, "")
    .trim();
}

// The minimal structural shapes each source array is expected to satisfy. Kept
// local + loose (only the fields the projection reads) so this module does not
// import the per-kind annotation interfaces and stays decoupled.
interface AuthorLike {
  userId?: string;
  displayName?: string;
}
interface ReplyLike {
  id?: string;
  text?: string;
  author?: AuthorLike;
  reviewState?: { model: "Review" | "Marked"; state: string };
}
interface HighlightLike {
  id: string;
  page: number;
  /** The reader's own markup style (highlight / underline / strikeout / squiggly). */
  markupStyle?: string;
  text?: string;
  note?: string;
  color?: string;
  author?: AuthorLike;
  replies?: readonly ReplyLike[];
}
interface NoteLike {
  id: string;
  page: number;
  text?: string;
  color?: string;
  author?: AuthorLike;
  replies?: readonly ReplyLike[];
  resolved?: boolean;
}
interface InkLike {
  id: string;
  page: number;
  color?: string;
  insertedPageId?: string;
}
interface ShapeLike {
  id: string;
  page: number;
  kind?: string;
  color?: string;
}
interface BookmarkLike {
  id: string;
  page: number;
  label?: string;
}

/** The source's protection and visibility flags, as read from /F. */
interface NativeFlagsLike {
  hidden?: boolean;
  noView?: boolean;
  readOnly?: boolean;
  locked?: boolean;
  lockedContents?: boolean;
}
/** A source-PDF annotation as the viewer's native readers project it. */
interface NativeSourceLike {
  objectRef?: string;
  subtype?: string;
  flags?: NativeFlagsLike;
}
interface NativeCommentLike {
  id: string;
  page: number;
  text?: string;
  color?: string;
  author?: string;
  replies?: readonly ReplyLike[];
   
  reviewStates?: readonly ReplyLike[];
  nativeSource?: NativeSourceLike;
}
interface NativeShapeLike extends ShapeLike {
  nativeSource?: NativeSourceLike;
  text?: string;
  author?: string;
}
interface NativeInkLike {
  id: string;
  /** Zero-based page index (the ink reader's convention). */
  page: number;
  color?: string;
  nativeSource?: NativeSourceLike;
}

/** The grouped source arrays the viewer hands in (each may be empty/missing). */
export interface AnnotationSources {
  highlights?: readonly HighlightLike[];
  notes?: readonly NoteLike[];
  inks?: readonly InkLike[];
  shapes?: readonly ShapeLike[];
  bookmarks?: readonly BookmarkLike[];
  /** The source PDF's own /Text notes. */
  nativeNotes?: readonly NativeCommentLike[];
  /** The source PDF's own text markups (Highlight, Underline, ...). */
  nativeMarkups?: readonly NativeCommentLike[];
  /** The source PDF's own drawing markups (Square, Circle, Line, ...). */
  nativeShapes?: readonly NativeShapeLike[];
  /** The source PDF's own /Ink paths (one row per annotation object). */
  nativeInks?: readonly NativeInkLike[];
   
  nativeFreeTexts?: readonly NativeCommentLike[];
}

function nativeInvisible(source: NativeSourceLike | undefined): boolean {
  return Boolean(source?.flags?.hidden || source?.flags?.noView);
}

function nativeFields(
  source: NativeSourceLike | undefined,
): Pick<AnnotationListItem, "origin" | "readOnly"> {
  const flags = source?.flags;
  return {
    origin: "pdf",
    ...(flags?.readOnly || flags?.locked || flags?.lockedContents
      ? { readOnly: true as const }
      : {}),
  };
}

function validPage(value: unknown): value is number {
  return Number.isInteger(value) && (value as number) >= 1;
}

function validId(value: unknown): value is string {
  return typeof value === "string" && value.length > 0;
}

/**
 * Project a highlight into its preview: prefer the quoted text, then any user
 * note; an empty highlight (e.g. an area highlight) yields "".
 */
function highlightPreview(h: HighlightLike): string {
  return annotationPreview(h.text) || annotationPreview(h.note);
}

/** The highlight's complete body: the quoted text, then any user note. */
function highlightFullText(h: HighlightLike): string {
  const quoted = annotationFullText(h.text);
  const note = annotationFullText(h.note);
  if (quoted && note) return `${quoted}\n${note}`;
  return quoted || note;
}

/** `{ text }` when the body is non-empty, nothing otherwise. */
function textField(value: unknown): Pick<AnnotationListItem, "text"> {
  const text = annotationFullText(value);
  return text ? { text } : {};
}

/** Author display name projected for a row / reply ("" → omitted). */
function authorName(author: AuthorLike | undefined): string {
  return typeof author?.displayName === "string"
    ? author.displayName.trim()
    : "";
}

/**
 * Project an annotation's reply thread into panel reply rows: text-bearing
 * entries only, previewed (whitespace-collapsed + capped) and tagged with the
 * reply author's display name when present. Defensive: never throws.
 */
function projectReplies(
  replies: readonly ReplyLike[] | undefined,
  parentId: string,
): AnnotationListReply[] {
  if (!Array.isArray(replies)) return [];
  const out: AnnotationListReply[] = [];
  replies.forEach((reply, index) => {
    if (!reply || typeof reply !== "object") return;
    const text = annotationPreview(reply.text);
    // A review state record is listed as that state, with or without words.
    const reviewState =
      reply.reviewState &&
      (reply.reviewState.model === "Review" || reply.reviewState.model === "Marked") &&
      typeof reply.reviewState.state === "string" &&
      reply.reviewState.state
        ? { model: reply.reviewState.model, state: reply.reviewState.state }
        : undefined;
    if (!text && !reviewState) return;
    const name = authorName(reply.author);
    out.push({
      id:
        typeof reply.id === "string" && reply.id
          ? reply.id
          : `${parentId}:reply:${index}`,
      text,
      fullText: annotationFullText(reply.text) || (reviewState ? reviewState.state : ""),
      ...(name ? { authorName: name } : {}),
      ...(reviewState ? { reviewState } : {}),
    });
  });
  return out;
}

/** Author / reply projection shared by the highlight + note rows. */
function threadFields(item: {
  id: string;
  author?: AuthorLike;
  replies?: readonly ReplyLike[];
}): Pick<AnnotationListItem, "authorName" | "authorId" | "replies"> {
  const name = authorName(item.author);
  const authenticatedId =
    typeof item.author?.userId === "string" ? item.author.userId.trim() : "";
  const filterId = authenticatedId || (name ? `legacy-name:${name}` : "");
  const replies = projectReplies(item.replies, item.id);
  return {
    ...(name ? { authorName: name } : {}),
    ...(filterId ? { authorId: filterId } : {}),
    ...(replies.length > 0 ? { replies } : {}),
  };
}

/**
 * Fold the per-kind source arrays into ONE flat list of {@link
 * AnnotationListItem}. Items with an invalid id / page are skipped (never
 * throws). `order` is the item's index within its own kind's array, so a later
 * stable sort can keep creation order within a page. The kinds are concatenated
 * in {@link ANNOTATION_LIST_KINDS} order; callers sort/filter afterwards.
 */
const MARKUP_BY_NAME: Readonly<Record<string, AnnotationListMarkup>> = {
  highlight: "highlight",
  underline: "underline",
  strikeout: "strikeout",
  strikethrough: "strikeout",
  squiggly: "squiggly",
};
/** The markup a reader style or a PDF /Subtype names, if it is one. */
function markupField(name: unknown): { markup?: AnnotationListMarkup } {
  const markup = typeof name === "string" ? MARKUP_BY_NAME[name.toLowerCase()] : undefined;
  return markup ? { markup } : {};
}

export function buildAnnotationList(
  sources: AnnotationSources,
): AnnotationListItem[] {
  const items: AnnotationListItem[] = [];
  // One entry per annotation. A highlight the reader made and the file now
  // carries as its own markup (the same id, synced into the PDF) came back in
  // both sources: listed and counted twice, and two list rows with one React
  // key. The reader's own copy is pushed first and is the editable one.
  const listed = new Set<string>();
  const push = (
    kind: AnnotationListKind,
    arr: readonly { id: string; page: number }[] | undefined,
    project: (item: never, index: number) => Omit<AnnotationListItem, "kind">,
  ) => {
    if (!Array.isArray(arr)) return;
    arr.forEach((raw, index) => {
      if (!raw || typeof raw !== "object") return;
      if (!validId(raw.id) || !validPage(raw.page)) return;
      const item = { kind, ...project(raw as never, index) };
      const key = `${item.kind}:${item.id}`;
      if (listed.has(key)) return;
      listed.add(key);
      items.push(item);
    });
  };

  push("highlight", sources.highlights, (h: HighlightLike, order) => ({
    id: h.id,
    page: h.page,
    ...markupField(h.markupStyle),
    preview: highlightPreview(h),
    ...textField(highlightFullText(h)),
    color: typeof h.color === "string" ? h.color : undefined,
    order,
    ...threadFields(h),
  }));
  push("note", sources.notes, (n: NoteLike, order) => ({
    id: n.id,
    page: n.page,
    preview: annotationPreview(n.text),
    ...textField(n.text),
    color: typeof n.color === "string" ? n.color : undefined,
    order,
    reviewState: n.resolved === true ? "resolved" : "active",
    ...threadFields(n),
  }));
  push("ink", sources.inks, (k: InkLike, order) => ({
    id: k.id,
    page: k.page,
    ...(typeof k.insertedPageId === "string" ? { insertedPageId: k.insertedPageId } : {}),
    preview: "",
    color: typeof k.color === "string" ? k.color : undefined,
    order,
  }));
  push("shape", sources.shapes, (s: ShapeLike, order) => ({
    id: s.id,
    page: s.page,
    // The shape kind (rect / line / arrow / ellipse) is the most useful label.
    preview: typeof s.kind === "string" ? s.kind : "",
    color: typeof s.color === "string" ? s.color : undefined,
    order,
  }));
  push("bookmark", sources.bookmarks, (b: BookmarkLike, order) => ({
    id: b.id,
    page: b.page,
    preview: annotationPreview(b.label),
    order,
  }));

  // The source PDF's own annotations, by object identity: the reader shows
  // them on the page and in the side panel, so the list counts them too
  
  const visibleNative = <T extends { nativeSource?: NativeSourceLike }>(
    arr: readonly T[] | undefined,
  ): T[] | undefined =>
    Array.isArray(arr)
      ? arr.filter((item) => item && !nativeInvisible(item.nativeSource))
      : undefined;
  const nativeThread = (item: NativeCommentLike) =>
    threadFields({
      id: item.id,
      author:
        typeof item.author === "string" && item.author.trim()
          ? { displayName: item.author }
          : undefined,
      replies:
        item.reviewStates && item.reviewStates.length > 0
          ? [...(item.replies ?? []), ...item.reviewStates]
          : item.replies,
    });
  push("note", visibleNative(sources.nativeNotes), (n: NativeCommentLike, order) => ({
    id: n.id,
    page: n.page,
    preview: annotationPreview(n.text),
    ...textField(n.text),
    color: typeof n.color === "string" ? n.color : undefined,
    order: order + NATIVE_ORDER_OFFSET,
    ...nativeThread(n),
    ...nativeFields(n.nativeSource),
  }));
  push(
    "highlight",
    visibleNative(sources.nativeMarkups),
    (h: NativeCommentLike, order) => ({
      id: h.id,
      page: h.page,
      ...markupField(h.nativeSource?.subtype),
      preview:
        annotationPreview(h.text) ||
        (typeof h.nativeSource?.subtype === "string"
          ? h.nativeSource.subtype
          : ""),
      ...textField(h.text),
      color: typeof h.color === "string" ? h.color : undefined,
      order: order + NATIVE_ORDER_OFFSET,
      ...nativeThread(h),
      ...nativeFields(h.nativeSource),
    }),
  );
  push(
    "shape",
    visibleNative(sources.nativeShapes),
    (sh: NativeShapeLike, order) => ({
      id: sh.id,
      page: sh.page,
      // A drawing with a comment is listed by it, with its author, like a
      
      preview:
        annotationPreview(sh.text) ||
        (typeof sh.kind === "string" ? sh.kind : ""),
      ...textField(sh.text),
      color: typeof sh.color === "string" ? sh.color : undefined,
      order: order + NATIVE_ORDER_OFFSET,
      ...nativeThread({ id: sh.id, page: sh.page, author: sh.author }),
      ...nativeFields(sh.nativeSource),
    }),
  );
  // A source text box or callout is a comment the file carries (WPS and
  // Acrobat list it): its words, its author, read-only where the page does
  // not offer it for editing.
  push("note", visibleNative(sources.nativeFreeTexts), (n: NativeCommentLike, order) => ({
    id: n.id,
    page: n.page,
    preview: annotationPreview(n.text),
    ...textField(n.text),
    color: typeof n.color === "string" ? n.color : undefined,
    order: order + NATIVE_ORDER_OFFSET * 2,
    ...nativeThread(n),
    ...nativeFields(n.nativeSource),
  }));
  const inkByObject = new Map<string, { id: string; page: number; color?: string; nativeSource?: NativeSourceLike }>();
  for (const stroke of visibleNative(sources.nativeInks) ?? []) {
    if (!stroke || typeof stroke !== "object" || !validId(stroke.id)) continue;
    const page = stroke.page + 1;
    const key = `${page}:${stroke.nativeSource?.objectRef ?? stroke.id}`;
    if (!inkByObject.has(key)) {
      inkByObject.set(key, {
        id: stroke.nativeSource?.objectRef
          ? `pdf-native-ink:${page}:${stroke.nativeSource.objectRef}`
          : stroke.id,
        page,
        color: stroke.color,
        nativeSource: stroke.nativeSource,
      });
    }
  }
  push("ink", [...inkByObject.values()], (k: NativeInkLike, order) => ({
    id: k.id,
    page: k.page,
    preview: "",
    color: typeof k.color === "string" ? k.color : undefined,
    order: order + NATIVE_ORDER_OFFSET,
    ...nativeFields(k.nativeSource),
  }));

  return items;
}

/**
 * Stable sort of the consolidated list by PAGE ascending, then by the source
 * `kind` order ({@link ANNOTATION_LIST_KINDS}), then by `order` (creation order
 * within a kind). Returns a NEW array; the input is untouched. Deterministic for
 * equal keys so the panel order does not jitter between renders.
 */
export function sortAnnotationList(
  items: readonly AnnotationListItem[],
): AnnotationListItem[] {
  const kindRank = (kind: AnnotationListKind) =>
    ANNOTATION_LIST_KINDS.indexOf(kind);
  return items.slice().sort((a, b) => {
    if (a.page !== b.page) return a.page - b.page;
    const ka = kindRank(a.kind);
    const kb = kindRank(b.kind);
    if (ka !== kb) return ka - kb;
    return a.order - b.order;
  });
}

/**
 * Filter the consolidated list to a set of kinds. `active` is the set of kinds
 * to KEEP; a null / undefined / empty set means "all kinds" (no filtering).
 * Returns a NEW array. Unknown kinds in the set are ignored.
 */
export function filterAnnotationList(
  items: readonly AnnotationListItem[],
  active: ReadonlySet<AnnotationListKind> | null | undefined,
): AnnotationListItem[] {
  if (!active || active.size === 0) return items.slice();
  return items.filter((item) => active.has(item.kind));
}

export interface AnnotationListFilter {
  kinds?: ReadonlySet<AnnotationListKind> | null;
  pages?: ReadonlySet<number> | null;
  authorIds?: ReadonlySet<string> | null;
  reviewStates?: ReadonlySet<"active" | "resolved"> | null;
}

/** Apply the Comments pane's type/page/author/review filters together. */
export function filterAnnotationListAdvanced(
  items: readonly AnnotationListItem[],
  filter: AnnotationListFilter,
): AnnotationListItem[] {
  return filterAnnotationList(items, filter.kinds).filter((item) => {
    if (filter.pages?.size && !filter.pages.has(item.page)) return false;
    if (
      filter.authorIds?.size &&
      (!item.authorId || !filter.authorIds.has(item.authorId))
    ) {
      return false;
    }
    if (
      filter.reviewStates?.size &&
      (!item.reviewState || !filter.reviewStates.has(item.reviewState))
    ) {
      return false;
    }
    return true;
  });
}

/** Per-kind counts across the (unfiltered) list — drives the filter chips. */
export function annotationCounts(
  items: readonly AnnotationListItem[],
): Record<AnnotationListKind, number> {
  const counts: Record<AnnotationListKind, number> = {
    highlight: 0,
    note: 0,
    ink: 0,
    shape: 0,
    bookmark: 0,
  };
  for (const item of items) {
    if (isAnnotationListKind(item.kind)) counts[item.kind] += 1;
  }
  return counts;
}

/**
 * Serialize the consolidated list to a STABLE JSON string for "export all
 * annotations". Sorted by page first (so the file is deterministic regardless of
 * the live panel filter), each entry carries the kind / page / preview / id /
 * color. `pretty` (default true) emits 2-space indentation.
 */
export function exportAnnotationsJson(
  items: readonly AnnotationListItem[],
  pretty = true,
): string {
  const sorted = sortAnnotationList(items);
  const payload = {
    version: 1 as const,
    count: sorted.length,
    annotations: sorted.map((item) => ({
      id: item.id,
      kind: item.kind,
      ...(item.markup ? { markup: item.markup } : {}),
      page: item.page,
      preview: item.preview,
      // The complete body and thread: "export all" is not a list of labels
      
      ...(item.text ? { text: item.text } : {}),
      ...(item.authorName ? { author: item.authorName } : {}),
      ...(item.reviewState ? { reviewState: item.reviewState } : {}),
      ...(item.replies && item.replies.length > 0
        ? {
            replies: item.replies.map((reply) => ({
              id: reply.id,
              text: reply.fullText || reply.text,
              ...(reply.authorName ? { author: reply.authorName } : {}),
            })),
          }
        : {}),
      ...(item.color ? { color: item.color } : {}),
      ...(item.origin ? { origin: item.origin } : {}),
    })),
  };
  return JSON.stringify(payload, null, pretty ? 2 : undefined);
}

/** The body an export writes for a row: the full text, else the label. */
function exportBody(item: AnnotationListItem): string {
  return item.text || item.preview;
}

/** Reply lines for an export, `Reply (Author): text` (author when known). */
function exportReplyLines(item: AnnotationListItem): string[] {
  return (item.replies ?? []).map((reply) => {
    const who = reply.authorName ? ` (${reply.authorName})` : "";
    return `Reply${who}: ${reply.fullText || reply.text}`;
  });
}

/** Human-readable label for a kind in the text summary (capitalized noun). */
const KIND_LABEL: Record<AnnotationListKind, string> = {
  highlight: "Highlight",
  note: "Note",
  ink: "Ink",
  shape: "Shape",
  bookmark: "Bookmark",
};
const MARKUP_LABEL: Record<AnnotationListMarkup, string> = {
  highlight: "Highlight",
  underline: "Underline",
  strikeout: "Strikeout",
  squiggly: "Squiggly",
};
/** A row's type in an export: a text markup by the markup it is. */
function exportKindLabel(item: AnnotationListItem): string {
  return (item.markup ? MARKUP_LABEL[item.markup] : KIND_LABEL[item.kind]) ?? item.kind;
}

/**
 * Render the consolidated list as a plain-text summary for "export all
 * annotations" (a copy-pasteable outline). Sorted by page; one line per
 * annotation: `Page N — Kind: preview` (the preview omitted when empty). A
 * header line with the total count precedes the body. Empty input yields just
 * the header. Pure text — safe to drop into a `.txt` download or clipboard.
 */
export function exportAnnotationsText(
  items: readonly AnnotationListItem[],
): string {
  const sorted = sortAnnotationList(items);
  const lines: string[] = [`PDF annotations (${sorted.length})`, ""];
  for (const item of sorted) {
    const label = exportKindLabel(item);
    const body = exportBody(item);
    const who = item.authorName ? ` (${item.authorName})` : "";
    // A multi-line body continues on indented lines; replies follow, indented.
    lines.push(
      `Page ${item.page} — ${label}${who}${body ? `: ${body.replace(/\n/g, "\n    ")}` : ""}`,
    );
    for (const reply of exportReplyLines(item)) {
      lines.push(`    ${reply.replace(/\n/g, "\n        ")}`);
    }
  }
  // Trailing newline so the file ends cleanly; the leading blank line after the
  // header is dropped when there are no items.
  if (sorted.length === 0) return `${lines[0]}\n`;
  return `${lines.join("\n")}\n`;
}

/**
 * Build a TipTap document node for "export annotations to Word" — a `Heading 1`
 * with the total count, then the annotations grouped under a `Heading 2` per
 * page, one paragraph per annotation (`Kind: preview`, preview omitted when
 * empty). Sorted by page (deterministic regardless of the live filter). The
 * returned plain object is sent to the server's Word writer by the caller
 * (`requestServerDocx`) — this module stays decoupled from the export layer.
 * The text is already plain (never HTML); the writer escapes it. PURE.
 */
export function annotationsToDocxDoc(items: readonly AnnotationListItem[]): {
  type: "doc";
  content: Array<Record<string, unknown>>;
} {
  const sorted = sortAnnotationList(items);
  const content: Array<Record<string, unknown>> = [
    {
      type: "heading",
      attrs: { level: 1 },
      content: [{ type: "text", text: `PDF Annotations (${sorted.length})` }],
    },
  ];
  let currentPage: number | null = null;
  for (const item of sorted) {
    if (item.page !== currentPage) {
      currentPage = item.page;
      content.push({
        type: "heading",
        attrs: { level: 2 },
        content: [{ type: "text", text: `Page ${item.page}` }],
      });
    }
    const label = exportKindLabel(item);
    const body = exportBody(item);
    const who = item.authorName ? ` (${item.authorName})` : "";
    content.push({
      type: "paragraph",
      content: paragraphText(body ? `${label}${who}: ${body}` : `${label}${who}`),
    });
    
    for (const reply of exportReplyLines(item)) {
      content.push({ type: "paragraph", content: paragraphText(reply) });
    }
  }
  return { type: "doc", content };
}

/** Text nodes for a paragraph, a hard break between the body's lines. */
function paragraphText(text: string): Array<Record<string, unknown>> {
  const out: Array<Record<string, unknown>> = [];
  text.split("\n").forEach((line, index) => {
    if (index > 0) out.push({ type: "hardBreak" });
    if (line) out.push({ type: "text", text: line });
  });
  return out.length > 0 ? out : [{ type: "text", text }];
}
