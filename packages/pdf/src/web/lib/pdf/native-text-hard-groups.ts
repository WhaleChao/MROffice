/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { GROUPING_ROLES, INLINE_ROLES } from "./tagged-reading-order";

/** The tag of the marked-content sequence the editor writes round a box. */
export const PDF_NATIVE_BOX_MARK_TAG = "MRAIBox";

const MAX_STRUCT_NODES = 10_000;
const MAX_STRUCT_DEPTH = 64;

type MarkedContentItem = Readonly<{ type?: unknown; tag?: unknown; id?: unknown }>;

export interface PdfJsMarkedTextItems {
  /** The text items, as getTextContent without marked content reports them. */
  readonly items: readonly unknown[];
  /** Per item: the ordinal of the innermost editor box around it. */
  readonly boxes: readonly (number | null)[];
  /** Per item: the innermost marked-content id around it (`p12R_mc3`). */
  readonly contentIds: readonly (string | null)[];
  /** Per item: drawn inside /Artifact marked content (pagination: a
   * watermark, a header or footer, a page number). */
  readonly artifacts: readonly boolean[];
  

  readonly ambiguousContentIds: ReadonlySet<string>;
}

/** Split getTextContent({ includeMarkedContent: true }) into its text items
 * and the marked content each one is drawn in. */
export function splitPdfJsMarkedContent(raw: readonly unknown[]): PdfJsMarkedTextItems {
  const items: unknown[] = [];
  const boxes: (number | null)[] = [];
  const contentIds: (string | null)[] = [];
  const artifacts: boolean[] = [];
  const frames: { box: number | null; id: string | null; artifact: boolean }[] = [];
  const opened = new Map<string, number>();
  let boxCount = 0;
  for (const value of raw) {
    const marker = value && typeof value === "object" ? (value as MarkedContentItem) : null;
    if (marker?.type === "beginMarkedContent" || marker?.type === "beginMarkedContentProps") {
      const tag = typeof marker.tag === "string" ? marker.tag : null;
      const id = marker.type === "beginMarkedContentProps" && typeof marker.id === "string" && marker.id
        ? marker.id
        : null;
      if (id) opened.set(id, (opened.get(id) ?? 0) + 1);
      frames.push({ box: tag === PDF_NATIVE_BOX_MARK_TAG ? boxCount++ : null, id, artifact: tag === "Artifact" });
      continue;
    }
    if (marker?.type === "endMarkedContent") {
      frames.pop();
      continue;
    }
    let box: number | null = null;
    let id: string | null = null;
    for (let index = frames.length - 1; index >= 0 && (box === null || id === null); index -= 1) {
      box ??= frames[index].box;
      id ??= frames[index].id;
    }
    items.push(value);
    boxes.push(box);
    contentIds.push(id);
    artifacts.push(frames.some((frame) => frame.artifact));
  }
  return Object.freeze({
    items: Object.freeze(items),
    boxes: Object.freeze(boxes),
    contentIds: Object.freeze(contentIds),
    artifacts: Object.freeze(artifacts),
    ambiguousContentIds: new Set([...opened].filter(([, count]) => count > 1).map(([id]) => id)),
  });
}

/**
 * The block each marked-content id of a page's structure tree belongs to:
 * the outermost element that is neither a grouping element (Document, Sect,
 * Table, TR, L ...) nor an inline one (Span, Link, Lbl, LBody ...) -- a P, a
 * heading, a list item, a table cell, a figure. A table of contents entry is
 * a block of its own. Content that only grouping and inline elements hold,
 * or that two elements name, belongs to none. Empty for a tree it cannot
 * read within its bounds.
 */
export function pdfStructBlockByContentId(structTree: unknown): ReadonlyMap<string, string> {
  const blocks = new Map<string, string>();
  const named = new Set<string>();
  let nodes = 0;
  let nextBlock = 0;
  const walk = (raw: unknown, depth: number, block: string | null): boolean => {
    if (!raw || typeof raw !== "object" || depth > MAX_STRUCT_DEPTH || ++nodes > MAX_STRUCT_NODES) return false;
    const node = raw as Readonly<{ type?: unknown; id?: unknown; role?: unknown; children?: unknown }>;
    if (node.type === "content") {
      const id = typeof node.id === "string" ? node.id : "";
      if (!id) return true;
      if (named.has(id)) blocks.delete(id);
      else if (block) blocks.set(id, block);
      named.add(id);
      return true;
    }
    if (node.type !== undefined) return true;
    const role = typeof node.role === "string" ? node.role.toLowerCase() : "";
    const opens = block === null && role !== "root" && !INLINE_ROLES.has(role) &&
      (!GROUPING_ROLES.has(role) || role === "toci");
    const own = opens ? `struct:${nextBlock++}` : block;
    if (node.children === undefined) return true;
    if (!Array.isArray(node.children)) return false;
    return node.children.every((child) => walk(child, depth + 1, own));
  };
  return walk(structTree, 0, null) ? blocks : new Map();
}

/**
 * The key of the box each text item is drawn in, by item index: `box:<n>`
 * for the editor's n-th box mark, `struct:<n>` for a structure block, null
 * for none. Null as a whole when no item has one (every untagged page the
 * editor never wrote), so the page is grouped as it always was.
 */
export function pdfNativeTextHardGroups(
  marked: PdfJsMarkedTextItems,
  structTree: unknown,
): readonly (string | null)[] | null {
  const blocks = marked.contentIds.some((id) => id !== null)
    ? pdfStructBlockByContentId(structTree)
    : new Map<string, string>();
  let any = false;
  const groups = marked.items.map((_, index) => {
    const box = marked.boxes[index];
    if (box !== null && box !== undefined) {
      any = true;
      return `box:${box}`;
    }
    const id = marked.contentIds[index];
    const block = id && !marked.ambiguousContentIds.has(id) ? blocks.get(id) : undefined;
    if (block) any = true;
    return block ?? null;
  });
  return any ? Object.freeze(groups) : null;
}

type PdfTextStyles = Record<
  string,
  { fontFamily?: string; ascent?: number; descent?: number; vertical?: boolean }
>;

export interface PdfMarkedTextContentPageLike {
  getTextContent(params?: { includeMarkedContent?: boolean; disableNormalization?: boolean }): Promise<{ items?: unknown; styles?: unknown }>;
  getStructTree?: () => Promise<unknown>;
}

/**
 * A page's text items with the box key of each (pdfNativeTextHardGroups).
 * The structure tree is read only when the page's text carries MCIDs.
 */
export async function readPdfTextContentWithHardGroups(
  page: PdfMarkedTextContentPageLike,
): Promise<{
  items: unknown[];
  styles: PdfTextStyles;
  hardGroups: readonly (string | null)[] | null;
  /** Per item: inside /Artifact marked content (splitPdfJsMarkedContent). */
  artifacts: readonly boolean[];
}> {
  const content = await page.getTextContent({ includeMarkedContent: true });
  const marked = splitPdfJsMarkedContent(Array.isArray(content?.items) ? (content.items as unknown[]) : []);
  const styles = (content?.styles && typeof content.styles === "object" ? content.styles : {}) as PdfTextStyles;
  let structTree: unknown = null;
  if (marked.contentIds.some((id) => id !== null) && typeof page.getStructTree === "function") {
    try {
      structTree = await page.getStructTree();
    } catch {
      // An unreadable tree names no blocks; the editor's own boxes still do.
      structTree = null;
    }
  }
  let items: unknown[] = [...marked.items];
  if (items.some((item) => (item as { dir?: unknown } | null)?.dir === "rtl")) {
    try {
      const raw = await page.getTextContent({ includeMarkedContent: true, disableNormalization: true });
      items = pdfJsRtlItemsInLogicalOrder(items,
        splitPdfJsMarkedContent(Array.isArray(raw?.items) ? (raw.items as unknown[]) : []).items);
    } catch {
      // Unread: the items as pdf.js normalized them.
    }
  }
  return {
    items,
    styles,
    hardGroups: pdfNativeTextHardGroups(marked, structTree),
    artifacts: marked.artifacts,
  };
}

/** A string's characters as a sorted list: the same letters in any order. */
export function samePdfTextCharacters(left: string, right: string): boolean {
  if (left === right) return true;
  if (left.length !== right.length) return false;
  return [...left].sort().join("") === [...right].sort().join("");
}










export function pdfJsRtlItemsInLogicalOrder(items: readonly unknown[], raw: readonly unknown[]): unknown[] {
  if (raw.length !== items.length) return [...items];
  return items.map((item, index) => {
    const read = item as { dir?: unknown; str?: unknown } | null;
    const unnormalized = (raw[index] as { str?: unknown } | null)?.str;
    if (!read || read.dir !== "rtl" || typeof read.str !== "string" || typeof unnormalized !== "string") return item;
    // Ligatures spelled out as pdf.js spells them; a letter's positional
    // form stays the form it is (pdf.js keeps those).
    const logical = [...unnormalized].map((character) => {
      const spelled = character.normalize("NFKC");
      return [...spelled].length > 1 ? spelled : character;
    }).join("");
    return logical !== read.str && samePdfTextCharacters(logical, read.str) ? { ...(item as object), str: logical } : item;
  });
}
