/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for the PDF outline / bookmarks panel.
 *
 * pdf.js exposes a document's bookmark tree via `getOutline()`. Each raw node
 * carries a `title`, a `dest` (either a named-destination string or an explicit
 * destination array whose first element is a page reference), and a nested
 * `items` array of children. This module holds everything with no DOM / pdf.js
 * runtime dependency so it can be unit-tested in isolation and reused by the
 * viewer:
 *
 *   - `normalizeOutlineTree` — turn the loosely-typed raw items into a typed,
 *     sanitized tree (collapsed whitespace titles, stable ids, depth cap),
 *   - `flattenOutline` — depth-first flatten for keyboard nav / search,
 *   - `resolveOutlineDestination` — resolve a raw `dest` to a 0-based page index
 *     via an INJECTED async resolver, so the resolution shape is unit-testable
 *     without a live pdf.js document.
 *
 * Titles are ALWAYS rendered as plain text (never HTML); the normalization here
 * collapses runaway whitespace and caps length — it is not HTML escaping (React
 * text nodes escape on render), but it keeps a hostile/garbage title from
 * blowing out the panel layout.
 */

/** A raw outline node as returned by pdf.js `getOutline()`. Loosely typed. */
export interface RawOutlineItem {
  title?: unknown;
  /**
   * Named destination (string) or explicit destination array (`dest[0]` is the
   * page reference). `null`/missing when the node is a non-navigating header or
   * an external link.
   */
  dest?: string | unknown[] | null;
  /** External URL — present on link bookmarks (not navigated by this panel). */
  url?: unknown;
  items?: unknown;
  /** The item's /F style flags, as pdf.js reads them. */
  bold?: unknown;
  italic?: unknown;
  /** The item's /Count: positive when the source shows its children, negative
   *  when it keeps them folded. */
  count?: unknown;
}

/** A normalized, sanitized outline node ready for rendering. */
export interface OutlineNode {
  /** Stable id derived from the node's position in the tree (e.g. "0-2-1"). */
  id: string;
  /** Sanitized plain-text title (whitespace-collapsed, length-capped). */
  title: string;
  /**
   * The raw destination carried over verbatim for later resolution, or `null`
   * when the node has no in-document destination (header / external link).
   */
  dest: string | unknown[] | null;
  /** Zero-based depth in the tree (root nodes are depth 0). */
  depth: number;
  /** Normalized children. */
  children: OutlineNode[];
   
  bold?: true;
  italic?: true;
  

  open?: boolean;
}

/** Max length of a sanitized title before truncation (with an ellipsis). */
export const OUTLINE_TITLE_MAX_LENGTH = 300;
export const MANAGED_OUTLINE_MAX_ABS_COORDINATE = 10_000_000;
export const MANAGED_OUTLINE_MAX_ZOOM = 64;

/**
 * Max nesting depth retained. Pathological / malicious PDFs can nest bookmarks
 * extremely deeply (or cyclically, once self-referencing); capping protects the
 * render + flatten from unbounded recursion. Nodes deeper than the cap are
 * dropped. Default mirrors a generous real-world TOC depth.
 */
export const OUTLINE_MAX_DEPTH = 12;

/** Fallback title shown when a node's title is empty after sanitization. */
export const OUTLINE_UNTITLED = "";

/** Native destination owned by MRAI. Page numbers are 1-based. */
export interface ManagedOutlineDestination {
  readonly pageNumber: number;
  readonly left: number | null;
  readonly top: number | null;
  readonly zoom: number | null;
}

/**
 * One node in the MRAI-owned outline subtree. Imported outline nodes never use
 * this shape: keeping the two trees separate is the primary mutation boundary.
 */
export interface ManagedOutlineNode {
  readonly id: string;
  readonly title: string;
  readonly destination: ManagedOutlineDestination;
  readonly children: readonly ManagedOutlineNode[];
}

export type ManagedOutlineMutation =
  | Readonly<{
      kind: "add";
      node: ManagedOutlineNode;
      parentId: string | null;
      index?: number;
    }>
  | Readonly<{ kind: "rename"; id: string; title: string }>
  | Readonly<{
      kind: "update-destination";
      id: string;
      destination: ManagedOutlineDestination;
    }>
  | Readonly<{
      kind: "move";
      id: string;
      parentId: string | null;
      index: number;
    }>
  | Readonly<{ kind: "delete"; id: string }>;

export type ManagedOutlineValidationCode =
  | "invalid-tree"
  | "invalid-id"
  | "duplicate-id"
  | "invalid-title"
  | "invalid-destination"
  | "depth-exceeded"
  | "missing-node"
  | "cycle";

export class ManagedOutlineValidationError extends Error {
  readonly code: ManagedOutlineValidationCode;

  constructor(code: ManagedOutlineValidationCode, message: string) {
    super(message);
    this.name = "ManagedOutlineValidationError";
    this.code = code;
  }
}

const MANAGED_OUTLINE_ID_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/;

export function hasUnpairedUtf16Surrogate(value: string): boolean {
  for (let index = 0; index < value.length; index += 1) {
    const unit = value.charCodeAt(index);
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return true;
      index += 1;
    } else if (unit >= 0xdc00 && unit <= 0xdfff) {
      return true;
    }
  }
  return false;
}

function managedDestination(value: unknown): ManagedOutlineDestination {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new ManagedOutlineValidationError(
      "invalid-destination",
      "A managed outline destination must be an object.",
    );
  }
  const candidate = value as Record<string, unknown>;
  const pageNumber = candidate.pageNumber;
  if (!Number.isSafeInteger(pageNumber) || Number(pageNumber) < 1) {
    throw new ManagedOutlineValidationError(
      "invalid-destination",
      "A managed outline destination requires a positive page number.",
    );
  }
  const coordinate = (key: "left" | "top" | "zoom") => {
    const entry = candidate[key];
    if (entry === null) return null;
    if (typeof entry !== "number" || !Number.isFinite(entry)) {
      throw new ManagedOutlineValidationError(
        "invalid-destination",
        `Managed outline ${key} must be finite or null.`,
      );
    }
    if (key !== "zoom" && Math.abs(entry) > MANAGED_OUTLINE_MAX_ABS_COORDINATE) {
      throw new ManagedOutlineValidationError(
        "invalid-destination",
        `Managed outline ${key} exceeds the supported coordinate range.`,
      );
    }
    if (key === "zoom" && (entry <= 0 || entry > MANAGED_OUTLINE_MAX_ZOOM)) {
      throw new ManagedOutlineValidationError(
        "invalid-destination",
        `Managed outline zoom must be greater than zero and at most ${MANAGED_OUTLINE_MAX_ZOOM}.`,
      );
    }
    return entry;
  };
  return Object.freeze({
    pageNumber: Number(pageNumber),
    left: coordinate("left"),
    top: coordinate("top"),
    zoom: coordinate("zoom"),
  });
}

/** Strictly clone and validate a complete managed tree. */
export function validateManagedOutlineTree(
  value: unknown,
): readonly ManagedOutlineNode[] {
  if (!Array.isArray(value)) {
    throw new ManagedOutlineValidationError(
      "invalid-tree",
      "The managed outline must be an array.",
    );
  }
  const ids = new Set<string>();
  const ancestors = new Set<object>();
  const walk = (items: unknown[], depth: number): readonly ManagedOutlineNode[] => {
    if (depth > OUTLINE_MAX_DEPTH) {
      throw new ManagedOutlineValidationError(
        "depth-exceeded",
        `Managed outlines may not exceed depth ${OUTLINE_MAX_DEPTH}.`,
      );
    }
    return Object.freeze(
      items.map((item) => {
        if (!item || typeof item !== "object" || Array.isArray(item)) {
          throw new ManagedOutlineValidationError(
            "invalid-tree",
            "Every managed outline entry must be an object.",
          );
        }
        if (ancestors.has(item)) {
          throw new ManagedOutlineValidationError(
            "cycle",
            "Managed outline trees may not contain cycles.",
          );
        }
        const candidate = item as Record<string, unknown>;
        if (
          typeof candidate.id !== "string" ||
          !MANAGED_OUTLINE_ID_PATTERN.test(candidate.id)
        ) {
          throw new ManagedOutlineValidationError(
            "invalid-id",
            "Managed outline ids must be stable bounded identifiers.",
          );
        }
        if (ids.has(candidate.id)) {
          throw new ManagedOutlineValidationError(
            "duplicate-id",
            `Duplicate managed outline id: ${candidate.id}`,
          );
        }
        ids.add(candidate.id);
        if (
          typeof candidate.title !== "string" ||
          hasUnpairedUtf16Surrogate(candidate.title)
        ) {
          throw new ManagedOutlineValidationError(
            "invalid-title",
            "Managed outline titles must contain well-formed UTF-16 text.",
          );
        }
        const title = sanitizeOutlineTitle(candidate.title);
        if (!title) {
          throw new ManagedOutlineValidationError(
            "invalid-title",
            "Managed outline titles may not be empty.",
          );
        }
        if (!Array.isArray(candidate.children)) {
          throw new ManagedOutlineValidationError(
            "invalid-tree",
            "Managed outline children must be an array.",
          );
        }
        ancestors.add(item);
        const children = walk(candidate.children, depth + 1);
        ancestors.delete(item);
        return Object.freeze({
          id: candidate.id,
          title,
          destination: managedDestination(candidate.destination),
          children,
        });
      }),
    );
  };
  return walk(value, 0);
}

function mutateManagedOutlineTree(
  tree: readonly ManagedOutlineNode[],
  id: string,
  replace: (node: ManagedOutlineNode) => ManagedOutlineNode | null,
): readonly ManagedOutlineNode[] | null {
  let found = false;
  const walk = (items: readonly ManagedOutlineNode[]): readonly ManagedOutlineNode[] =>
    items.flatMap((node) => {
      if (node.id === id) {
        found = true;
        const next = replace(node);
        return next ? [next] : [];
      }
      const children = walk(node.children);
      return children === node.children
        ? [node]
        : [{ ...node, children } satisfies ManagedOutlineNode];
    });
  const result = walk(tree);
  return found ? result : null;
}

function insertManagedOutlineNode(
  tree: readonly ManagedOutlineNode[],
  parentId: string | null,
  index: number | undefined,
  node: ManagedOutlineNode,
): readonly ManagedOutlineNode[] {
  if (index !== undefined && (!Number.isSafeInteger(index) || index < 0)) {
    throw new ManagedOutlineValidationError(
      "invalid-tree",
      "Managed outline insertion indexes must be non-negative safe integers.",
    );
  }
  const insert = (items: readonly ManagedOutlineNode[]) => {
    const at = Math.min(
      items.length,
      Math.max(0, Number.isSafeInteger(index) ? Number(index) : items.length),
    );
    return [...items.slice(0, at), node, ...items.slice(at)];
  };
  if (parentId === null) return insert(tree);
  const result = mutateManagedOutlineTree(tree, parentId, (parent) => ({
    ...parent,
    children: insert(parent.children),
  }));
  if (!result) {
    throw new ManagedOutlineValidationError(
      "missing-node",
      `Unknown managed outline parent: ${parentId}`,
    );
  }
  return result;
}

/** Apply one atomic managed-tree operation and revalidate the complete result. */
export function applyManagedOutlineMutation(
  tree: readonly ManagedOutlineNode[],
  mutation: ManagedOutlineMutation,
): readonly ManagedOutlineNode[] {
  const current = validateManagedOutlineTree(tree);
  if (mutation.kind === "add") {
    const node = validateManagedOutlineTree([mutation.node])[0];
    return validateManagedOutlineTree(
      insertManagedOutlineNode(current, mutation.parentId, mutation.index, node),
    );
  }
  if (mutation.kind === "rename") {
    const title = sanitizeOutlineTitle(mutation.title);
    if (!title) {
      throw new ManagedOutlineValidationError(
        "invalid-title",
        "Managed outline titles may not be empty.",
      );
    }
    const result = mutateManagedOutlineTree(current, mutation.id, (node) => ({
      ...node,
      title,
    }));
    if (!result) {
      throw new ManagedOutlineValidationError(
        "missing-node",
        `Unknown managed outline node: ${mutation.id}`,
      );
    }
    return validateManagedOutlineTree(result);
  }
  if (mutation.kind === "update-destination") {
    const destination = managedDestination(mutation.destination);
    const result = mutateManagedOutlineTree(current, mutation.id, (node) => ({
      ...node,
      destination,
    }));
    if (!result) {
      throw new ManagedOutlineValidationError(
        "missing-node",
        `Unknown managed outline node: ${mutation.id}`,
      );
    }
    return validateManagedOutlineTree(result);
  }
  if (mutation.kind === "delete") {
    const result = mutateManagedOutlineTree(current, mutation.id, () => null);
    if (!result) {
      throw new ManagedOutlineValidationError(
        "missing-node",
        `Unknown managed outline node: ${mutation.id}`,
      );
    }
    return validateManagedOutlineTree(result);
  }

  let moving: ManagedOutlineNode | null = null;
  const without = mutateManagedOutlineTree(current, mutation.id, (node) => {
    moving = node;
    return null;
  });
  if (!without || !moving) {
    throw new ManagedOutlineValidationError(
      "missing-node",
      `Unknown managed outline node: ${mutation.id}`,
    );
  }
  const containsParent = (node: ManagedOutlineNode): boolean =>
    node.id === mutation.parentId || node.children.some(containsParent);
  if (mutation.parentId !== null && containsParent(moving)) {
    throw new ManagedOutlineValidationError(
      "cycle",
      "A managed outline entry cannot be nested inside its own subtree.",
    );
  }
  return validateManagedOutlineTree(
    insertManagedOutlineNode(without, mutation.parentId, mutation.index, moving),
  );
}

/**
 * Collapse runaway whitespace (including control chars / newlines) into single
 * spaces, trim, and cap length. The title is rendered as a text node downstream
 * so this is normalization, NOT HTML escaping — it just keeps a garbage title
 * from breaking layout. A non-string input yields an empty string.
 */
export function sanitizeOutlineTitle(
  value: unknown,
  maxLength: number = OUTLINE_TITLE_MAX_LENGTH,
): string {
  if (typeof value !== "string") return OUTLINE_UNTITLED;
  // Replace ALL whitespace + ASCII/Unicode control characters with a single
  // space, then collapse runs. \s covers tabs/newlines; the explicit ranges
  // strip the C0 controls (\u0000-\u001f), DEL (\u007f), and the C1 range
  // (\u0080-\u009f) that \s misses, so a hostile title can't smuggle them in.
  const collapsed = value.replace(/[\s\u0000-\u001f\u007f-\u009f]+/g, " ").trim();
  const cap =
    Number.isFinite(maxLength) && maxLength > 0
      ? Math.trunc(maxLength)
      : OUTLINE_TITLE_MAX_LENGTH;
  if (collapsed.length <= cap) return collapsed;
  // Truncate and append an ellipsis so the cap is observable (ellipsis included
  // in the budget so the result never exceeds `cap`).
  let prefix = collapsed.slice(0, Math.max(0, cap - 1));
  const last = prefix.charCodeAt(prefix.length - 1);
  if (last >= 0xd800 && last <= 0xdbff) prefix = prefix.slice(0, -1);
  return `${prefix.trimEnd()}…`;
}

/**
 * Whether a raw `dest` is navigable (a non-empty named destination string or a
 * non-empty explicit destination array). External links / header nodes are not.
 */
export function hasNavigableDest(dest: unknown): dest is string | unknown[] {
  if (typeof dest === "string") return dest.length > 0;
  if (Array.isArray(dest)) return dest.length > 0;
  return false;
}

function normalizeDest(dest: unknown): string | unknown[] | null {
  if (typeof dest === "string") return dest.length > 0 ? dest : null;
  if (Array.isArray(dest)) return dest.length > 0 ? dest : null;
  return null;
}

/**
 * Normalize a raw pdf.js outline tree into typed, sanitized nodes.
 *
 *   - titles are whitespace-collapsed + length-capped (rendered as text),
 *   - each node gets a stable id from its path (`parentId-index`),
 *   - depth is recorded and nodes deeper than `maxDepth` are dropped,
 *   - non-array / malformed `items` are treated as no children,
 *   - the raw `dest` is carried over verbatim for later resolution.
 *
 * @param rawItems The array from `pdfDoc.getOutline()` (may be null/undefined).
 * @param maxDepth Max nesting depth to retain (root = depth 0).
 */
export function normalizeOutlineTree(
  rawItems: unknown,
  maxDepth: number = OUTLINE_MAX_DEPTH,
): OutlineNode[] {
  const cap =
    Number.isFinite(maxDepth) && maxDepth >= 0 ? Math.trunc(maxDepth) : OUTLINE_MAX_DEPTH;

  const walk = (items: unknown, depth: number, parentId: string): OutlineNode[] => {
    if (depth > cap) return [];
    if (!Array.isArray(items)) return [];
    const nodes: OutlineNode[] = [];
    items.forEach((raw, index) => {
      if (!raw || typeof raw !== "object") return;
      const item = raw as RawOutlineItem;
      const id = parentId ? `${parentId}-${index}` : String(index);
      const children =
        depth + 1 <= cap ? walk(item.items, depth + 1, id) : [];
      const count = typeof item.count === "number" && Number.isFinite(item.count) ? item.count : 0;
      nodes.push({
        id,
        title: sanitizeOutlineTitle(item.title),
        dest: normalizeDest(item.dest),
        depth,
        children,
        ...(item.bold === true ? { bold: true as const } : {}),
        ...(item.italic === true ? { italic: true as const } : {}),
        ...(count !== 0 && children.length > 0 ? { open: count > 0 } : {}),
      });
    });
    return nodes;
  };

  return walk(rawItems, 0, "");
}

/** A node from a flattened outline (depth-first, parents before children). */
export interface FlatOutlineNode {
  id: string;
  title: string;
  dest: string | unknown[] | null;
  depth: number;
}

/**
 * Depth-first flatten of a normalized outline tree, parents before their
 * children (document/reading order). Each entry keeps its id/title/dest/depth
 * but drops the nested `children` array — handy for keyboard nav / search.
 */
export function flattenOutline(nodes: OutlineNode[]): FlatOutlineNode[] {
  const out: FlatOutlineNode[] = [];
  const walk = (list: OutlineNode[]) => {
    for (const node of list) {
      out.push({ id: node.id, title: node.title, dest: node.dest, depth: node.depth });
      if (node.children.length > 0) walk(node.children);
    }
  };
  walk(nodes);
  return out;
}

/**
 * Resolver hooks the destination resolution needs from a live pdf.js document.
 * Injecting them keeps `resolveOutlineDestination` unit-testable without pdf.js:
 * a test supplies fakes, the viewer supplies the real `pdfDoc` methods.
 */
export interface OutlineDestResolver {
  /** pdf.js `getDestination(name)` — resolves a NAMED dest to a dest array. */
  getDestination: (name: string) => Promise<unknown[] | null>;
  /** pdf.js `getPageIndex(ref)` — resolves a page-ref object to a 0-based index. */
  getPageIndex: (ref: unknown) => Promise<number>;
}

/**
 * Resolve a raw outline `dest` to a ZERO-BASED page index using an injected
 * resolver, or `null` when it cannot be resolved (no dest, unknown named dest,
 * malformed explicit dest, or the resolver throws).
 *
 * pdf.js destinations come in two shapes:
 *   - a string — a NAMED destination; resolve it to an explicit array first via
 *     `getDestination(name)`,
 *   - an array — an EXPLICIT destination whose first element (`dest[0]`) is the
 *     page reference passed to `getPageIndex(ref)`.
 *
 * Errors from the resolver are swallowed into `null` so one broken bookmark
 * never breaks the whole panel.
 */
export async function resolveOutlineDestination(
  dest: string | unknown[] | null | undefined,
  resolver: OutlineDestResolver,
): Promise<number | null> {
  try {
    let explicit: unknown[] | null = null;
    if (typeof dest === "string") {
      if (dest.length === 0) return null;
      explicit = await resolver.getDestination(dest);
    } else if (Array.isArray(dest)) {
      explicit = dest;
    }
    if (!Array.isArray(explicit) || explicit.length === 0) return null;
    const ref = explicit[0];
    if (ref == null) return null;
    const index = await resolver.getPageIndex(ref);
    if (typeof index !== "number" || !Number.isInteger(index) || index < 0) {
      return null;
    }
    return index;
  } catch {
    return null;
  }
}

/**
 * Convenience wrapper turning a resolved zero-based page index into the
 * 1-based page NUMBER the viewer's `goToPage` expects, or `null` when
 * unresolved.
 */
export async function resolveOutlinePageNumber(
  dest: string | unknown[] | null | undefined,
  resolver: OutlineDestResolver,
): Promise<number | null> {
  const index = await resolveOutlineDestination(dest, resolver);
  return index == null ? null : index + 1;
}

// ---- current-section highlight ----------------------------------------------

/**
 * One outline entry with its destination RESOLVED to a 1-based page number,
 * used by the panel's current-section highlight. Built by the viewer from the
 * flattened outline + the async dest resolver, then kept sorted by page so
 * {@link currentOutlineSectionId} can binary-search it on every page change.
 */
export interface OutlineSectionEntry {
  /** Stable node id (matches {@link OutlineNode.id}). */
  id: string;
  /** 1-based page number the entry starts on. */
  page: number;
}

/**
 * Filter + sort resolved outline entries for {@link currentOutlineSectionId}:
 * entries whose page is not a positive integer are dropped, the rest are
 * sorted ascending by page. The sort is stable, so entries sharing a start
 * page keep their input (document/flatten) order — the LAST of a tie is the
 * deepest/most recent section starting on that page.
 */
export function sortOutlineSectionEntries(
  entries: ReadonlyArray<{ id: string; page: number | null | undefined }>,
): OutlineSectionEntry[] {
  if (!Array.isArray(entries)) return [];
  return entries
    .filter(
      (e): e is { id: string; page: number } =>
        Boolean(e) &&
        typeof e.id === "string" &&
        Number.isInteger(e.page) &&
        (e.page as number) >= 1,
    )
    .map((e) => ({ id: e.id, page: e.page }))
    .sort((a, b) => a.page - b.page);
}

/**
 * The id of the outline section the reader is currently inside: the LAST entry
 * (in the page-sorted list from {@link sortOutlineSectionEntries}) whose start
 * page is <= `currentPage`, or `null` when the reader is before the first
 * section / the list is empty / `currentPage` is invalid. Binary search — the
 * viewer calls this on every page change over potentially large outlines.
 * When several sections start on the same page, the last one wins (the
 * deepest / most recently started section, matching reader convention).
 */
export function currentOutlineSectionId(
  sorted: ReadonlyArray<OutlineSectionEntry>,
  currentPage: number,
): string | null {
  if (!Array.isArray(sorted) || sorted.length === 0) return null;
  if (!Number.isInteger(currentPage) || currentPage < 1) return null;
  // Binary search for the rightmost entry with page <= currentPage.
  let lo = 0;
  let hi = sorted.length - 1;
  let best = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (sorted[mid].page <= currentPage) {
      best = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return best === -1 ? null : sorted[best].id;
}

// ---- export (flatten) shape -----------------------------------------------

/**
 * One outline entry fully RESOLVED for a flatten export. Raw pdf.js `dest`
 * values (named strings / explicit arrays holding live page-ref objects) only
 * resolve against the open pdf.js document, while the export runs on pdf-lib —
 * so the viewer resolves every destination to a page NUMBER up front and hands
 * the flatten pipeline this plain serializable tree.
 */
export interface ExportOutlineNode {
  /** Sanitized plain-text title. */
  title: string;
  /**
   * 1-based ORIGINAL page number the entry navigates to, or `null` for a
   * non-navigating header (no dest / unresolvable dest / external link).
   */
  page: number | null;
  children: ExportOutlineNode[];
}

/**
 * Resolve a normalized outline tree ({@link normalizeOutlineTree}) into the
 * serializable {@link ExportOutlineNode} shape a flatten export consumes: each
 * node's raw `dest` is resolved to a 1-based ORIGINAL page number via the
 * injected resolver (the live pdf.js doc in the viewer; fakes in tests). An
 * unresolvable dest degrades the node to a header (`page: null`) rather than
 * dropping it, so its children survive. Depth stays bounded by the input tree
 * (already capped by {@link normalizeOutlineTree}).
 */
export async function outlineTreeForExport(
  nodes: ReadonlyArray<OutlineNode>,
  resolver: OutlineDestResolver,
): Promise<ExportOutlineNode[]> {
  const out: ExportOutlineNode[] = [];
  for (const node of nodes) {
    const index = await resolveOutlineDestination(node.dest, resolver);
    out.push({
      title: node.title,
      page: index == null ? null : index + 1,
      children: await outlineTreeForExport(node.children, resolver),
    });
  }
  return out;
}
