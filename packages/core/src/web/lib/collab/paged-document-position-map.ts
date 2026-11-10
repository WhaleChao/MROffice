/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { Fragment, Slice, type Node as ProseMirrorNode } from "@tiptap/pm/model";
import { AddMarkStep, AttrStep, RemoveMarkStep, ReplaceAroundStep, ReplaceStep, type Step } from "@tiptap/pm/transform";
import { canonicalDocumentJson, canonicalDocumentSchema, PAGINATION_RUNTIME_ATTRIBUTES } from "./paged-document-projection";

type JsonNode = {
  type: string;
  attrs?: Record<string, unknown>;
  text?: string;
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>;
  content?: JsonNode[];
  [TRACE_NODE]?: number;
};
const TRACE_NODE = Symbol("local-pagination-position");
const CONTINUES = ["pageSplitContinues", "pageSplitItemContinues", "pageTableContinues", "pageTableRowContinues", "pageTableCellContinues", "pageTableBlockContinues"];
const CONTINUATION = ["pageSplitContinuation", "pageSplitItemContinuation", "pageTableContinuation", "pageTableRowContinuation", "pageTableCellContinuation", "pageTableBlockContinuation"];
/** Continuation flags of a top-level split textblock: its head and tail are
 * joined by plain inline concatenation (a code block adds one "\n"). */
const TEXT_SPLIT_ATTRS: ReadonlySet<string> = new Set(["pageSplitContinues", "pageSplitContinuation"]);
/** What a piece of a split textblock may carry and still be joined by
 * concatenation: the identity it names is dropped with the flags. */
const JOINED_PIECE_ATTRS: ReadonlySet<string> = new Set([...TEXT_SPLIT_ATTRS, "pageSplitPieceOf"]);

interface SourceNode { node: ProseMirrorNode; from: number; replica: boolean }
interface Segment { localFrom: number; localTo: number; canonicalFrom: number; canonicalTo: number; replica: boolean }
interface Point { local: number; canonical: number; replica?: boolean }
/** A block boundary point: "open" is before/at the start of a node's
 * content, "close" at/after its end. Needed to shift a boundary that sits
 * exactly at an insertion into an empty textblock. */
interface BlockPoint extends Point { edge: "open" | "close" }
/** Everything a map answers from. `segments` stays in canonical order (the
 * order the projection walk emits text and leaves); `blockPoints` in the walk's
 * order. Both orders are part of the answer (tie-breaking), so a derived state
 * keeps them exactly as a fresh walk would produce them. */
interface MapState { local: ProseMirrorNode; canonical: ProseMirrorNode; segments: Segment[]; blockPoints: BlockPoint[] }

export interface PagedDocumentPositionMap {
  readonly canonical: ProseMirrorNode;
  toCanonical(position: number, association?: -1 | 1): number | null;
  toLocal(position: number, association?: -1 | 1, localHint?: number): number | null;
}

/** How many maps were built from scratch and how many were derived from a
 * previous map through a transaction (diagnostics and tests). */
export const pagedDocumentPositionMapBuilds = { full: 0, derived: 0 };

const states = new WeakMap<PagedDocumentPositionMap, MapState>();

/** Trace the existing, authoritative serializer instead of maintaining a
 * second paragraph/list/table merger. Temporary marks prevent adjacent text
 * fragments from coalescing before their source ranges have been measured.
 * They exist only in this detached JSON walk and never enter either editor. */
export function createPagedDocumentPositionMap(local: ProseMirrorNode): PagedDocumentPositionMap {
  pagedDocumentPositionMapBuilds.full += 1;
  return finalize(buildState(local));
}

/**
 * The map of `doc`, reached from `previous`'s document by `steps` (`docs[i]`
 * is the document before step i, as on a Transaction). A step that only
 * replaces inline content inside one textblock outside every page-fragment
 * structure (table fragments, repeated headers, split lists/containers, note
 * slices) changes the projection by the same inline replacement, so the
 * previous map is shifted instead of re-walking the whole document. Any other
 * step rebuilds from scratch. The answers are identical either way.
 */
export function derivePagedDocumentPositionMap(
  previous: PagedDocumentPositionMap,
  steps: readonly Step[],
  docs: readonly ProseMirrorNode[],
  doc: ProseMirrorNode,
): PagedDocumentPositionMap {
  let state = states.get(previous) ?? null;
  if (state && steps.length && docs[0] === state.local) {
    for (let index = 0; state && index < steps.length; index++) {
      const after = index + 1 < steps.length ? docs[index + 1] : doc;
      try {
        state = state.local === docs[index] && after
          ? deriveStep(state, steps[index], after) ?? deriveTopLevelRun(state, steps[index], after) ??
            derivePageBoundary(state, steps[index], after) ?? deriveLeafSwap(state, steps[index], after) ??
            deriveDroppedSheets(state, steps[index], after) ?? deriveDroppedCopies(state, steps[index], after) ??
            deriveAppendedPages(state, steps[index], after) ?? deriveSplitPieceMarkup(state, steps[index], after) ??
            deriveTextblockSplit(state, steps[index], after) ?? deriveJoinedPieces(state, steps[index], after)
          : null;
      } catch { state = null; }
    }
    if (state && state.local === doc) {
      pagedDocumentPositionMapBuilds.derived += 1;
      return finalize(state);
    }
    noteDerivationFallback(steps, docs);
  }
  return createPagedDocumentPositionMap(doc);
}

/** Development trace (a probe sets `window.__pagedPositionMapFallbacks` to
 * an array): what a step the map could not follow touched, so the edits
 * that rebuild the whole map on every key can be found. */
function noteDerivationFallback(steps: readonly Step[], docs: readonly ProseMirrorNode[]): void {
  const sink = (globalThis as { __pagedPositionMapFallbacks?: unknown[] }).__pagedPositionMapFallbacks;
  if (!Array.isArray(sink) || sink.length >= 200) return;
  try {
    const step = steps[0] as unknown as { from?: number; to?: number; slice?: { content: { forEach: (f: (n: ProseMirrorNode) => void) => void } } };
    const doc = docs[0];
    if (!doc || typeof step?.from !== "number") { sink.push({ kind: steps[0]?.constructor?.name ?? "?" }); return; }
    const $from = doc.resolve(step.from);
    const chain: string[] = [];
    for (let depth = $from.depth; depth > 0; depth -= 1) {
      const node = $from.node(depth);
      const runtime = Object.keys(node.attrs).filter((key) => PAGINATION_RUNTIME_ATTRIBUTES.has(key) && node.attrs[key] !== null && node.attrs[key] !== undefined && node.attrs[key] !== false && node.attrs[key] !== node.type.spec.attrs?.[key]?.default);
      chain.push(runtime.length ? `${node.type.name}[${runtime.join(",")}]` : node.type.name);
    }
    const inserted: string[] = [];
    step.slice?.content.forEach((node) => inserted.push(node.type.name));
    const inline: string[] = [];
    $from.parent.forEach((node) => { if (!node.isText) inline.push(node.isLeaf ? node.type.name : `${node.type.name}+`); });
    sink.push({ kind: steps[0]?.constructor?.name, steps: steps.length, chain, inserted, inline: [...new Set(inline)] });
  } catch {
    sink.push({ kind: "unreadable" });
  }
}

function buildState(local: ProseMirrorNode): MapState {
  const schema = canonicalDocumentSchema(local.type.schema);
  let traceMark = "__mrnote_pagination_position__";
  while (local.type.schema.marks[traceMark]) traceMark += "_";
  const sources = new Map<number, SourceNode[]>();
  const repeated = new Map<string, number>();
  const repeatedHeaders = new Set<string>();
  // A header row the page cut stands in two pieces beside its whole copy: the
  // copy is not any one of them, and tracing the three as one node joined the
  // pieces' text and refused the document at its first keystroke (corpus
  // R284-D15: "Depth", " (m)" and the repeated "Depth (m)").
  const cutRows = new Set<string>();
  local.descendants(node => {
    if (node.type.name !== "table" || typeof node.attrs.pageTableFragmentId !== "string") return;
    node.forEach(row => {
      const key = `table:${node.attrs.pageTableFragmentId}:row:${row.attrs.pageTableSourceRow}`;
      if (row.attrs.pageTableRepeatedHeader === true) repeatedHeaders.add(key);
      if (row.attrs.pageTableRowContinues === true || row.attrs.pageTableRowContinuation === true) cutRows.add(key);
    });
  });
  for (const key of cutRows) repeatedHeaders.delete(key);
  let nextId = 0;
  const trace = (node: ProseMirrorNode, from: number, alias?: string, replica = false): JsonNode => {
    replica ||= node.attrs.pageTableRepeatedHeader === true;
    const tableId = node.type.name === "table" && typeof node.attrs.pageTableFragmentId === "string"
      ? node.attrs.pageTableFragmentId : undefined;
    const nodeAlias = alias ?? (tableId ? `table-container:${tableId}` : node.type.name === "footnoteList" ? "footnote-list" : undefined);
    const id = nodeAlias && repeated.has(nodeAlias) ? repeated.get(nodeAlias)! : ++nextId;
    if (nodeAlias) repeated.set(nodeAlias, id);
    const records = sources.get(id) ?? [];
    records.push({ node, from, replica }); sources.set(id, records);
    const json = shallowJson(node);
    json[TRACE_NODE] = id;
    if (node.isText) {
      json.marks = [...(json.marks ?? []), { type: traceMark, attrs: { id } }];
      return json;
    }
    const children: JsonNode[] = [];
    node.forEach((child, offset, index) => {
      const headerAlias = `table:${tableId}:row:${child.attrs.pageTableSourceRow}`;
      const childAlias = tableId && repeatedHeaders.has(headerAlias)
        ? headerAlias
        : alias ? `${alias}/${index}` : undefined;
      children.push(trace(child, from + (node === local ? 0 : 1) + offset, childAlias, replica));
    });
    if (children.length) json.content = children;
    return json;
  };
  const projected = canonicalDocumentJson(trace(local, 0));
  const segments: Segment[] = [];
  const blockPoints: BlockPoint[] = [];
  const clean = (json: JsonNode, from: number, root = false): { json: JsonNode; size: number } => {
    const id = json.type === "text"
      ? Number(json.marks?.find(mark => mark.type === traceMark)?.attrs?.id)
      : json[TRACE_NODE];
    const result = { ...json };
    delete result[TRACE_NODE];
    if (result.marks) {
      result.marks = result.marks.filter(mark => mark.type !== traceMark);
      if (result.marks.length === 0) delete result.marks;
    }
    let contentSize = 0;
    if (json.content) result.content = json.content.map(child => {
      const next = clean(child, from + (root ? 0 : 1) + contentSize);
      contentSize += next.size;
      return next.json;
    });
    // The size the parsed node will have (the whole tree is parsed and
    // checked once below); parsing every subtree here re-parsed each node
    // once per ancestor.
    const size = root ? contentSize : projectedNodeSize(schema, result, contentSize);
    for (const source of sources.get(id ?? -1) ?? []) {
      const start = source.from;
      if (source.node.isText || source.node.isLeaf) {
        if (source.node.nodeSize !== size) throw new Error("A projected leaf changed semantic size");
        segments.push({ localFrom: start, localTo: start + size, canonicalFrom: from, canonicalTo: from + size, replica: source.replica });
      } else if (!root && source.node.type.name !== "pageBlock") {
        if (!CONTINUATION.some(key => source.node.attrs[key] === true)) {
          blockPoints.push({ local: start, canonical: from, replica: source.replica, edge: "open" }, { local: start + 1, canonical: from + 1, replica: source.replica, edge: "open" });
        }
        if (!CONTINUES.some(key => source.node.attrs[key] === true)) {
          blockPoints.push({ local: start + source.node.nodeSize - 1, canonical: from + size - 1, replica: source.replica, edge: "close" }, { local: start + source.node.nodeSize, canonical: from + size, replica: source.replica, edge: "close" });
        }
      }
    }
    return { json: result, size };
  };
  const cleaned = clean(projected, 0, true);
  const canonical = schema.nodeFromJSON(cleaned.json);
  canonical.check();
  if (canonical.content.size !== cleaned.size) throw new Error("The projected document changed semantic size");
  return { local, canonical, segments, blockPoints };
}

/** Node#toJSON without its content (the trace supplies traced children), with
 * the same key order. */
function shallowJson(node: ProseMirrorNode): JsonNode {
  const json: JsonNode = { type: node.type.name };
  if (Object.keys(node.attrs).length) json.attrs = node.attrs as Record<string, unknown>;
  if (node.content.size) json.content = [];
  if (node.marks.length) json.marks = node.marks.map(mark => mark.toJSON() as { type: string; attrs?: Record<string, unknown> });
  if (node.isText) json.text = node.text;
  return json;
}

function projectedNodeSize(schema: ProseMirrorNode["type"]["schema"], json: JsonNode, contentSize: number): number {
  if (json.type === "text") return typeof json.text === "string" ? json.text.length : 0;
  const type = schema.nodes[json.type];
  if (!type) throw new RangeError(`Unknown node type: ${json.type}`);
  return type.isLeaf ? 1 : contentSize + 2;
}

function finalize(state: MapState): PagedDocumentPositionMap {
  const { local, canonical, segments } = state;
  // Only a position outside every text and leaf reads them, and a map is made
  // for every key: three objects per paragraph of the document, nearly always
  // for nothing.
  let builtPoints: Point[] | null = null;
  const allPoints = (): Point[] => {
    if (builtPoints) return builtPoints;
    const points: Point[] = [...state.blockPoints, { local: 0, canonical: 0 }, { local: local.content.size, canonical: canonical.content.size }];
    for (const segment of segments) points.push(
      { local: segment.localFrom, canonical: segment.canonicalFrom, replica: segment.replica },
      { local: segment.localTo, canonical: segment.canonicalTo, replica: segment.replica },
    );
    return (builtPoints = points);
  };
  const select = (candidates: number[], association: -1 | 1, hint?: number) => {
    if (!candidates.length) return null;
    if (hint !== undefined) return candidates.reduce((best, value) => Math.abs(value - hint) < Math.abs(best - hint) ? value : best);
    return association < 0 ? Math.min(...candidates) : Math.max(...candidates);
  };
  const map: PagedDocumentPositionMap = {
    canonical,
    toCanonical(position, association = 1) {
      if (!Number.isInteger(position) || position < 0 || position > local.content.size) return null;
      const contained = segments.filter(s => s.localFrom <= position && position <= s.localTo);
      if (contained.length) return select(contained.map(s => s.canonicalFrom + position - s.localFrom), association);
      const points = allPoints();
      const exact = points.filter(point => point.local === position);
      if (exact.length) return select(exact.map(point => point.canonical), association);
      const before = points.filter(point => point.local < position).sort((a, b) => b.local - a.local)[0];
      const after = points.filter(point => point.local > position).sort((a, b) => a.local - b.local)[0];
      // Page wrappers and removed continuation boundaries collapse to seams.
      if (before && after && before.canonical === after.canonical) return before.canonical;
      const parent = local.resolve(position).parent;
      if (parent === local || parent.type.name === "pageBlock") return association < 0 ? before?.canonical ?? null : after?.canonical ?? null;
      return null;
    },
    toLocal(position, association = 1, hint) {
      if (!Number.isInteger(position) || position < 0 || position > canonical.content.size) return null;
      let contained = segments.filter(s => s.canonicalFrom <= position && position <= s.canonicalTo);
      if (hint === undefined && contained.some(s => !s.replica)) contained = contained.filter(s => !s.replica);
      if (contained.length) return select(contained.map(s => s.localFrom + position - s.canonicalFrom), association, hint);
      let exact = allPoints().filter(point => point.canonical === position);
      if (hint === undefined && exact.some(point => !point.replica)) exact = exact.filter(point => !point.replica);
      if (exact.length) return select(exact.map(point => point.local), association, hint);
      return null;
    },
  };
  states.set(map, state);
  return map;
}

/** A pagination runtime attribute that carries a value (a fragment id, a
 * continuation flag, a source row). */
function hasRuntimeState(node: ProseMirrorNode, ignore?: ReadonlySet<string>): boolean {
  for (const key in node.attrs) {
    if (!PAGINATION_RUNTIME_ATTRIBUTES.has(key) || ignore?.has(key)) continue;
    const value = node.attrs[key];
    if (value === null || value === undefined || value === false) continue;
    if (value === node.type.spec.attrs?.[key]?.default) continue;
    return true;
  }
  return false;
}

/** Inline content the projection copies verbatim and the map measures as
 * segments: text and leaf atoms without runtime state. */
function plainInline(fragment: Fragment): boolean {
  let plain = true;
  fragment.forEach(node => { if (plain && !(node.isText || (node.isInline && node.isLeaf && !hasRuntimeState(node)))) plain = false; });
  return plain;
}

function deriveStep(state: MapState, step: Step, after: ProseMirrorNode): MapState | null {
  if (!(step instanceof ReplaceStep)) return null;
  const { from, to, slice } = step;
  if (slice.openStart !== 0 || slice.openEnd !== 0) return null;
  const before = state.local;
  if (from < 0 || to > before.content.size || from > to) return null;
  const $from = before.resolve(from), $to = before.resolve(to);
  const block = $from.parent, depth = $from.depth;
  if (depth < 1 || !$from.sameParent($to) || !block.isTextblock) return null;
  // Every ancestor below the page must be copied verbatim by the projection:
  // no page-fragment structure (split containers/lists, table fragments,
  // repeated headers) and no aliased note list, whose slices share traces.
  // A list or container split across pages is the one exception: its pieces
  // are joined child by child, so a textblock inside one of its children that
  // carries no state of its own is still copied verbatim (the bibliography at
  // a book's end runs over three sheets: every key typed there rebuilt the
  
  for (let level = 1; level < depth; level++) {
    const ancestor = $from.node(level);
    if (ancestor.type.name === "pageBlock") { if (level !== 1) return null; continue; }
    if (ancestor.type.name === "footnoteList") return null;
    if (hasRuntimeState(ancestor, level === 2 && !ancestor.isTextblock && !ancestor.type.name.startsWith("table") ? TEXT_SPLIT_ATTRS : undefined)) return null;
  }
  if (hasRuntimeState(block, JOINED_PIECE_ATTRS)) return null;
  const split = block.attrs.pageSplitContinues === true || block.attrs.pageSplitContinuation === true;
  const container = $from.node(depth - 1);
  // Only the top-level block merger joins split textblocks, by plain inline
  // concatenation of non-empty pieces (an empty piece changes the join).
  if (split && (container !== before && container.type.name !== "pageBlock")) return null;
  const start = $from.before(depth), contentStart = start + 1, contentEnd = contentStart + block.content.size;
  const replaced = after.nodeAt(start);
  const delta = slice.size - (to - from);
  if (!replaced || replaced.type !== block.type || !replaced.sameMarkup(block) ||
    replaced.content.size !== block.content.size + delta || after.content.size !== before.content.size + delta) return null;
  if (split && (block.content.size === 0 || replaced.content.size === 0)) return null;
  if (!plainInline(block.content) || !plainInline(replaced.content) || !plainInline(slice.content)) return null;

  // Where this textblock's content starts in the projection. Its inline
  // children are one contiguous run of the canonical-ordered segments.
  const { segments, blockPoints } = state;
  let first = -1, count = 0;
  for (let index = 0; index < segments.length; index++) {
    const segment = segments[index];
    if (segment.localFrom < contentStart || segment.localTo > contentEnd) continue;
    if (first < 0) first = index;
    else if (index !== first + count) return null;
    count += 1;
  }
  if (count !== block.childCount) return null;
  let canonicalStart: number;
  if (count) {
    canonicalStart = segments[first].canonicalFrom;
    for (let index = first; index < first + count; index++) {
      const segment = segments[index];
      if (segment.replica || segment.canonicalFrom - canonicalStart !== segment.localFrom - contentStart) return null;
    }
  } else {
    const open = blockPoints.filter(point => point.local === contentStart && point.edge === "open");
    if (open.length !== 1 || open[0].replica) return null;
    canonicalStart = open[0].canonical;
    first = 0;
    while (first < segments.length && segments[first].canonicalFrom < canonicalStart) first += 1;
  }
  const canonicalFrom = canonicalStart + (from - contentStart), canonicalTo = canonicalStart + (to - contentStart);
  const $canonicalFrom = state.canonical.resolve(canonicalFrom), $canonicalTo = state.canonical.resolve(canonicalTo);
  if (!$canonicalFrom.sameParent($canonicalTo) || !$canonicalFrom.parent.isTextblock ||
    $canonicalFrom.parent.type.name !== block.type.name) return null;
  const schema = state.canonical.type.schema;
  const inserted: ProseMirrorNode[] = [];
  slice.content.forEach(node => inserted.push(schema.nodeFromJSON(node.toJSON())));
  const canonical = state.canonical.replace(canonicalFrom, canonicalTo, new Slice(Fragment.fromArray(inserted), 0, 0));
  canonical.resolve(canonicalFrom).parent.check();

  const shift = (value: number, low: number, high: number, edge: BlockPoint["edge"]): number | null => {
    if (value < low) return value;
    if (value > high) return value + delta;
    if (low === high) return edge === "close" ? value + delta : value;
    if (value === low) return value;
    if (value === high) return value + delta;
    return null;
  };
  const nextPoints: BlockPoint[] = [];
  for (const point of blockPoints) {
    const local = shift(point.local, from, to, point.edge);
    const canonicalPosition = shift(point.canonical, canonicalFrom, canonicalTo, point.edge);
    if (local === null || canonicalPosition === null) return null;
    nextPoints.push(local === point.local && canonicalPosition === point.canonical ? point : { ...point, local, canonical: canonicalPosition });
  }
  const moved = (segment: Segment): Segment | null => {
    let { localFrom, localTo, canonicalFrom: segmentFrom, canonicalTo: segmentTo } = segment;
    if (localFrom >= to) { localFrom += delta; localTo += delta; } else if (localTo > from) return null;
    if (segmentFrom >= canonicalTo) { segmentFrom += delta; segmentTo += delta; } else if (segmentTo > canonicalFrom) return null;
    if (localFrom === segment.localFrom && segmentFrom === segment.canonicalFrom) return segment;
    return { localFrom, localTo, canonicalFrom: segmentFrom, canonicalTo: segmentTo, replica: segment.replica };
  };
  const nextSegments: Segment[] = [];
  for (let index = 0; index < first; index++) {
    const segment = moved(segments[index]);
    if (!segment) return null;
    nextSegments.push(segment);
  }
  replaced.forEach((child, offset) => {
    nextSegments.push({ localFrom: contentStart + offset, localTo: contentStart + offset + child.nodeSize,
      canonicalFrom: canonicalStart + offset, canonicalTo: canonicalStart + offset + child.nodeSize, replica: false });
  });
  for (let index = first + count; index < segments.length; index++) {
    const segment = moved(segments[index]);
    if (!segment) return null;
    nextSegments.push(segment);
  }
  return { local: after, canonical, segments: nextSegments, blockPoints: nextPoints };
}

/** A top-level textblock the projection copies verbatim: a plain paragraph
 * or heading of a page, with no pagination state and only plain inline
 * content. */
/** A block the projection copies as it stands: no pagination runtime state
 * on it or anything inside it (not a page fragment, not a continuation, not
 * a repeated header), its text and inline leaves plain. A paragraph, and as
 * well a whole list, quote or table that no page boundary cuts. */
function verbatimBlock(node: ProseMirrorNode | null | undefined): node is ProseMirrorNode {
  if (!node || node.isInline || hasRuntimeState(node)) return false;
  if (node.isTextblock) return plainInline(node.content);
  let verbatim = true;
  node.forEach(child => { if (verbatim && !verbatimBlock(child)) verbatim = false; });
  return verbatim;
}

/** Same tree shape and sizes: positions inside the one are positions inside
 * the other. */
function sameShape(a: ProseMirrorNode, b: ProseMirrorNode): boolean {
  if (a.nodeSize !== b.nodeSize || a.isText !== b.isText || a.isLeaf !== b.isLeaf || a.childCount !== b.childCount) return false;
  for (let index = 0; index < a.childCount; index++) if (!sameShape(a.child(index), b.child(index))) return false;
  return true;
}

/** The map entries a fresh walk emits for a verbatim block at `local` /
 * `canonical`: its text and leaves as segments (document order), and for
 * every node below the page its four boundary points after its children's
 * (the walk's order). */
function emitVerbatim(node: ProseMirrorNode, local: number, canonical: number, segments: Segment[], points: BlockPoint[]) {
  if (node.isText || node.isLeaf) {
    segments.push({ localFrom: local, localTo: local + node.nodeSize, canonicalFrom: canonical, canonicalTo: canonical + node.nodeSize, replica: false });
    return;
  }
  node.forEach((child, offset) => emitVerbatim(child, local + 1 + offset, canonical + 1 + offset, segments, points));
  points.push(
    { local, canonical, replica: false, edge: "open" },
    { local: local + 1, canonical: canonical + 1, replica: false, edge: "open" },
    { local: local + node.nodeSize - 1, canonical: canonical + node.nodeSize - 1, replica: false, edge: "close" },
    { local: local + node.nodeSize, canonical: canonical + node.nodeSize, replica: false, edge: "close" },
  );
}

/** How many boundary points a verbatim block has in the walk. */
function verbatimPointCount(node: ProseMirrorNode): number {
  if (node.isText || node.isLeaf) return 0;
  let count = 4;
  node.forEach(child => { count += verbatimPointCount(child); });
  return count;
}

/**
 * Whole blocks put between two top-level blocks of one page, nothing
 * removed: an ink stroke landing (its shape goes in before the block it is
 * anchored to), a picture or a shape inserted. The run is the block beside
 * the gap -- the one after it, or before it at the page's end or beside a
 * page fragment -- replaced by itself plus the new blocks, so the run
 * derivation proves it like any other. Without it every stroke the pen
 * landed rebuilt the whole map: on an inked page that was the document
 * serialised, projected and parsed again for each landing, growing with
 * every stroke already on the page.
 */
function blocksInsertedAtGap(step: ReplaceStep, before: ProseMirrorNode): {
  page: ProseMirrorNode; firstIndex: number; lastIndex: number; runStart: number; runEnd: number;
} | null {
  const { from, to, slice } = step;
  if (from !== to || slice.openStart !== 0 || slice.openEnd !== 0 || slice.content.childCount === 0) return null;
  if (from < 0 || from > before.content.size) return null;
  const $gap = before.resolve(from);
  if ($gap.depth !== 1 || $gap.parent.type.name !== "pageBlock") return null;
  const page = $gap.parent, index = $gap.index(1);
  const next = index < page.childCount ? page.child(index) : null;
  if (next && verbatimBlock(next)) {
    return { page, firstIndex: index, lastIndex: index, runStart: from, runEnd: from + next.nodeSize };
  }
  const previous = index > 0 ? page.child(index - 1) : null;
  if (previous && verbatimBlock(previous)) {
    return { page, firstIndex: index - 1, lastIndex: index - 1, runStart: from - previous.nodeSize, runEnd: from };
  }
  return null;
}

/** The run of whole top-level blocks of one page a step rewrites: a replace
 * or a mark added or removed inside textblocks of one page (from the first
 * one's start to the last one's end), or a block's own markup or attribute
 * set around its unchanged content (the id the normaliser gives the
 * paragraph Enter made). */
function topLevelRunOf(step: Step, before: ProseMirrorNode): {
  page: ProseMirrorNode; firstIndex: number; lastIndex: number; runStart: number; runEnd: number; delta: number;
} | null {
  // Content replaced (or rewrapped around a kept gap) between two positions
  // inside, or at the edges of, top-level blocks of one page: everything
  // outside those blocks stands.
  const inner = (from: number, to: number) => {
    if (from < 0 || to > before.content.size || from > to) return null;
    const $from = before.resolve(from), $to = before.resolve(to);
    if ($from.depth < 1 || $to.depth < 1) return null;
    const page = $from.node(1);
    if (page !== $to.node(1) || page.type.name !== "pageBlock") return null;
    const firstIndex = $from.index(1), lastIndex = $to.depth >= 2 ? $to.index(1) : $to.index(1) - 1;
    if (lastIndex < firstIndex) return null;
    return { page, firstIndex, lastIndex, runStart: $from.depth >= 2 ? $from.before(2) : from, runEnd: $to.depth >= 2 ? $to.after(2) : to };
  };
  if (step instanceof ReplaceStep) {
    const run = inner(step.from, step.to) ?? blocksInsertedAtGap(step, before);
    return run && { ...run, delta: step.slice.size - (step.to - step.from) };
  }
  if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
    const { from, to } = step;
    if (from < 0 || to > before.content.size || from >= to) return null;
    const $from = before.resolve(from), $to = before.resolve(to);
    if ($from.depth < 2 || $to.depth < 2) return null;
    const page = $from.node(1);
    if (page !== $to.node(1) || page.type.name !== "pageBlock") return null;
    return { page, firstIndex: $from.index(1), lastIndex: $to.index(1), runStart: $from.before(2), runEnd: $to.after(2), delta: 0 };
  }
  if (step instanceof AttrStep) {
    if (step.pos < 0 || step.pos >= before.content.size) return null;
    const $at = before.resolve(step.pos);
    if ($at.depth < 1 || $at.node(1).type.name !== "pageBlock" || !$at.nodeAfter) return null;
    // The top-level block the attribute's node is or sits in.
    const runStart = $at.depth === 1 ? step.pos : $at.before(2);
    const block = before.nodeAt(runStart);
    if (!block) return null;
    return { page: $at.node(1), firstIndex: $at.index(1), lastIndex: $at.index(1), runStart, runEnd: runStart + block.nodeSize, delta: 0 };
  }
  if (step instanceof ReplaceAroundStep) {
    const { from, to, gapFrom, gapTo, insert, slice } = step;
    if (from < 0 || to > before.content.size) return null;
    const $from = before.resolve(from);
    const block = $from.nodeAfter;
    if ($from.depth === 1 && $from.parent.type.name === "pageBlock" && block && block.nodeSize === to - from &&
        gapFrom === from + 1 && gapTo === to - 1 && insert === 1 && !slice.openStart && !slice.openEnd &&
        slice.content.childCount === 1 && !slice.content.firstChild!.content.size) {
      // A top-level block's own markup set around its content.
      return { page: $from.parent, firstIndex: $from.index(1), lastIndex: $from.index(1), runStart: from, runEnd: to, delta: 0 };
    }
    const run = inner(from, to);
    return run && { ...run, delta: slice.size + (gapTo - gapFrom) - (to - from) };
  }
  return null;
}

/**
 * A step that replaces a run of whole top-level textblocks of one page by
 * another run of them: Enter splitting a paragraph, Backspace or Delete
 * joining two, a delete across paragraphs, a paste of plain paragraphs, a
 * paragraph's attributes set, bold on a selection. The projection copies such blocks verbatim (no
 * page fragment, no continuation), so the canonical document changes by the
 * same run, and the map changes only there: the run's own segments and block
 * points are built anew, everything after it moves by the size difference.
 * Anything else is null (built from scratch). Each Enter in a long document
 * rebuilt the whole map (the document serialised, projected and parsed
 * again), about a fifth of the key's time.
 */
function deriveTopLevelRun(state: MapState, step: Step, after: ProseMirrorNode): MapState | null {
  const before = state.local;
  const run = topLevelRunOf(step, before);
  if (!run) return null;
  const { page, firstIndex, lastIndex, runStart, runEnd, delta } = run;
  let runPointCount = 0;
  for (let index = firstIndex; index <= lastIndex; index++) {
    const block = page.child(index);
    if (!verbatimBlock(block)) return null;
    runPointCount += verbatimPointCount(block);
  }
  if (after.content.size !== before.content.size + delta) return null;
  // The run that replaced it: whole blocks of the same page, from the same start.
  const $start = after.resolve(runStart);
  if ($start.depth !== 1 || $start.parent.type.name !== "pageBlock") return null;
  const newBlocks: ProseMirrorNode[] = [];
  for (let position = runStart; position < runEnd + delta;) {
    const block = after.nodeAt(position);
    if (!verbatimBlock(block) || after.resolve(position).parent !== $start.parent) return null;
    newBlocks.push(block);
    position += block.nodeSize;
    if (position > runEnd + delta) return null;
  }
  if (newBlocks.length === 0) return null;
  // The page around the run is otherwise the same page.
  const newPage = $start.parent;
  if (newPage.childCount !== page.childCount - (lastIndex - firstIndex + 1) + newBlocks.length) return null;

  // Where the run sits in the projection: its first block's open point and
  // its last block's close point.
  const { segments, blockPoints } = state;
  const opening = blockPoints.filter(point => point.edge === "open" && point.local === runStart && !point.replica);
  const closing = blockPoints.filter(point => point.edge === "close" && point.local === runEnd && !point.replica);
  if (opening.length !== 1 || closing.length !== 1) return null;
  const canonicalStart = opening[0].canonical, canonicalEnd = closing[0].canonical;
  const $canonicalStart = state.canonical.resolve(canonicalStart), $canonicalEnd = state.canonical.resolve(canonicalEnd);
  if ($canonicalStart.depth !== 0 || $canonicalEnd.depth !== 0) return null;
  if ($canonicalEnd.index(0) - $canonicalStart.index(0) !== lastIndex - firstIndex + 1) return null;
  const schema = state.canonical.type.schema;
  const projected = canonicalDocumentJson({ type: "doc", content: newBlocks.map(block => block.toJSON()) }) as { content?: unknown[] };
  if (!Array.isArray(projected.content) || projected.content.length !== newBlocks.length) return null;
  const canonicalBlocks = projected.content.map(json => schema.nodeFromJSON(json));
  const canonical = state.canonical.replace(canonicalStart, canonicalEnd, new Slice(Fragment.fromArray(canonicalBlocks), 0, 0));
  if (canonical.content.size !== state.canonical.content.size + delta) return null;
  canonicalBlocks.forEach(block => block.check());

  // The run's own segments (canonical order) and block points (walk order),
  // as a fresh walk would emit them: per block its text, then its points.
  const runSegments: Segment[] = [];
  const runPoints: BlockPoint[] = [];
  let local = runStart, canonicalAt = canonicalStart;
  for (const [index, block] of newBlocks.entries()) {
    if (!sameShape(canonicalBlocks[index], block)) return null;
    emitVerbatim(block, local, canonicalAt, runSegments, runPoints);
    local += block.nodeSize;
    canonicalAt += block.nodeSize;
  }

  // The old run's points are one contiguous stretch of the walk; the points
  // of the blocks either side that touch it (a previous close, a next open)
  // stay where they are.
  const inRun = (point: BlockPoint) =>
    (point.edge === "open" && point.local >= runStart && point.local < runEnd) ||
    (point.edge === "close" && point.local > runStart && point.local <= runEnd);
  const firstPoint = blockPoints.findIndex(inRun);
  if (firstPoint < 0 || blockPoints.filter(inRun).length !== runPointCount) return null;
  for (let index = firstPoint; index < firstPoint + runPointCount; index++) if (!inRun(blockPoints[index])) return null;
  const movedPoint = (point: BlockPoint): BlockPoint | null => {
    const afterRun = point.local > runEnd || (point.local === runEnd && point.edge === "open");
    const afterCanonical = point.canonical > canonicalEnd || (point.canonical === canonicalEnd && point.edge === "open");
    if (afterRun !== afterCanonical) return null;
    // Every point past the run, per pen landing on a page of ink: written out,
    // a spread of each was most of what the map cost.
    return afterRun ? { local: point.local + delta, canonical: point.canonical + delta, replica: point.replica, edge: point.edge } : point;
  };
  const nextPoints: BlockPoint[] = [];
  for (let index = 0; index < blockPoints.length; index++) {
    if (index === firstPoint) { nextPoints.push(...runPoints); index += runPointCount - 1; continue; }
    const point = movedPoint(blockPoints[index]);
    if (!point) return null;
    nextPoints.push(point);
  }

  // The old run's segments are one contiguous stretch in canonical order.
  const inside = (segment: Segment) => segment.localFrom >= runStart && segment.localTo <= runEnd;
  const firstSegment = segments.findIndex(inside);
  const runSegmentCount = segments.filter(inside).length;
  const nextSegments: Segment[] = [];
  const movedSegment = (segment: Segment): Segment | null => {
    const afterRun = segment.localFrom >= runEnd;
    if (!afterRun && segment.localTo > runStart) return null;
    if ((segment.canonicalFrom >= canonicalEnd) !== afterRun) return null;
    return afterRun
      ? { localFrom: segment.localFrom + delta, localTo: segment.localTo + delta,
          canonicalFrom: segment.canonicalFrom + delta, canonicalTo: segment.canonicalTo + delta, replica: segment.replica }
      : segment;
  };
  if (firstSegment < 0) {
    // The run held no text: its segments go where a fresh walk puts them,
    // before the first segment past the run.
    const insertAt = segments.findIndex(segment => segment.canonicalFrom >= canonicalEnd);
    for (let index = 0; index < segments.length; index++) {
      if (index === insertAt) nextSegments.push(...runSegments);
      const segment = movedSegment(segments[index]);
      if (!segment) return null;
      nextSegments.push(segment);
    }
    if (insertAt < 0) nextSegments.push(...runSegments);
  } else {
    for (let index = firstSegment; index < firstSegment + runSegmentCount; index++) if (!inside(segments[index])) return null;
    for (let index = 0; index < segments.length; index++) {
      if (index === firstSegment) { nextSegments.push(...runSegments); index += runSegmentCount - 1; continue; }
      const segment = movedSegment(segments[index]);
      if (!segment) return null;
      nextSegments.push(segment);
    }
  }
  return { local: after, canonical, segments: nextSegments, blockPoints: nextPoints };
}

/** A top-level block a page boundary may stand next to without moving any
 * canonical position: not a piece of a block split across pages. */
function wholeBlock(node: ProseMirrorNode | null | undefined): boolean {
  if (!node) return false;
  for (const key of [...CONTINUES, ...CONTINUATION]) if (node.attrs[key] === true) return false;
  return true;
}







function boundaryNeighbours(before: ProseMirrorNode | null | undefined, after: ProseMirrorNode | null | undefined): boolean {
  if (wholeBlock(before) && wholeBlock(after)) return true;
  if (!before || !after || before.type !== after.type) return false;
  if (before.attrs.pageSplitContinues !== true || after.attrs.pageSplitContinuation !== true) return false;
  // Only the split-block flags: table fragments carry their own grouping.
  const flag = (node: ProseMirrorNode, keys: readonly string[]) => keys.some(key => !TEXT_SPLIT_ATTRS.has(key) && node.attrs[key] === true);
  return !flag(before, CONTINUES) && !flag(before, CONTINUATION) && !flag(after, CONTINUES) && !flag(after, CONTINUATION);
}











function derivePageBoundary(state: MapState, step: Step, after: ProseMirrorNode): MapState | null {
  if (!(step instanceof ReplaceStep)) return null;
  const { from, to, slice } = step;
  const before = state.local;
  if (from < 0 || to > before.content.size || from > to) return null;
  const page = before.type.schema.nodes.pageBlock;
  if (!page) return null;
  let delta: number;
  let boundary: number;
  if (from === to) {
    // [pageBlock close, pageBlock open] between two blocks of one page.
    if (slice.openStart !== 1 || slice.openEnd !== 1 || slice.content.childCount !== 2) return null;
    const ok = [0, 1].every(index => {
      const node = slice.content.child(index);
      return node.type === page && node.content.size === 0;
    });
    if (!ok) return null;
    const $at = before.resolve(from);
    if ($at.depth !== 1 || $at.parent.type !== page) return null;
    const index = $at.index();
    if (index === 0 || index >= $at.parent.childCount) return null;
    if (!boundaryNeighbours($at.nodeBefore, $at.nodeAfter)) return null;
    delta = 2;
    boundary = from;
  } else {
    // The boundary between two pages removed: from the end of one page's
    // content to the start of the next's.
    if (to !== from + 2 || slice.size !== 0) return null;
    const $from = before.resolve(from), $to = before.resolve(to);
    if ($from.depth !== 1 || $to.depth !== 1 || $from.parent.type !== page || $to.parent.type !== page) return null;
    if ($from.parentOffset !== $from.parent.content.size || $to.parentOffset !== 0) return null;
    if (!boundaryNeighbours($from.nodeBefore, $to.nodeAfter)) return null;
    delta = -2;
    boundary = from;
  }
  if (after.content.size !== before.content.size + delta) return null;
  const shiftLocal = (value: number, edge?: BlockPoint["edge"]): number | null => {
    if (delta > 0) {
      if (value < boundary) return value;
      if (value > boundary) return value + delta;
      // At the boundary: the end of the block before stays, the start of
      // the block after moves with it.
      return edge === "close" ? value : value + delta;
    }
    if (value <= from) return value;
    if (value >= to) return value + delta;
    return null;
  };
  const blockPoints: BlockPoint[] = [];
  for (const point of state.blockPoints) {
    const local = shiftLocal(point.local, point.edge);
    if (local === null) return null;
    blockPoints.push({ ...point, local });
  }
  const segments: Segment[] = [];
  for (const segment of state.segments) {
    const localFrom = shiftLocal(segment.localFrom, "open"), localTo = shiftLocal(segment.localTo, "close");
    if (localFrom === null || localTo === null || localTo - localFrom !== segment.localTo - segment.localFrom) return null;
    segments.push({ ...segment, localFrom, localTo });
  }
  return { local: after, canonical: state.canonical, segments, blockPoints };
}

/** The canonical form of a top-level leaf block: its runtime attributes gone, as canonicalDocumentJson strips them. */
function canonicalLeaf(schema: ProseMirrorNode["type"]["schema"], node: ProseMirrorNode): ProseMirrorNode {
  const json = node.toJSON() as JsonNode;
  if (json.attrs) json.attrs = Object.fromEntries(Object.entries(json.attrs).filter(([key]) => !PAGINATION_RUNTIME_ATTRIBUTES.has(key)));
  return schema.nodeFromJSON(json);
}

/** A copy of a row-sliced atom the serializer drops: the rows after the first copy's. */
function droppedRowSliceCopy(node: ProseMirrorNode): boolean {
  return node.isLeaf && Number(node.attrs.noteSliceStart) > 0;
}










function deriveLeafSwap(state: MapState, step: Step, after: ProseMirrorNode): MapState | null {
  if (!(step instanceof ReplaceStep)) return null;
  const { from, to, slice } = step;
  const before = state.local;
  if (to !== from + 1 || slice.openStart !== 0 || slice.openEnd !== 0 || slice.content.childCount !== 1) return null;
  if (from < 0 || to > before.content.size || after.content.size !== before.content.size) return null;
  const old = before.nodeAt(from), next = slice.content.child(0);
  if (!old || !old.isLeaf || old.nodeSize !== 1 || next.type !== old.type || !next.isLeaf || next.nodeSize !== 1) return null;
  const $at = before.resolve(from);
  if ($at.depth !== 1 || $at.parent.type.name !== "pageBlock") return null;
  if (after.nodeAt(from) !== next && !after.nodeAt(from)?.eq(next)) return null;
  const { segments, blockPoints } = state;
  const own = segments.filter(segment => segment.localFrom === from && segment.localTo === to);
  if (own.length === 0) {
    // Dropped by the walk: a continuation copy, and still one after the swap.
    if (!droppedRowSliceCopy(old) || !droppedRowSliceCopy(next)) return null;
    if (segments.some(segment => segment.localFrom < to && segment.localTo > from) ||
      blockPoints.some(point => point.local > from && point.local < to)) return null;
    return { local: after, canonical: state.canonical, segments, blockPoints };
  }
  if (own.length !== 1 || own[0].replica) return null;
  const canonicalFrom = own[0].canonicalFrom;
  const current = state.canonical.nodeAt(canonicalFrom);
  if (!current || current.type.name !== old.type.name || !current.isLeaf || current.nodeSize !== 1) return null;
  const replacement = canonicalLeaf(state.canonical.type.schema, next);
  if (replacement.nodeSize !== 1) return null;
  if (!canonicalLeaf(state.canonical.type.schema, old).eq(current)) return null;
  const canonical = state.canonical.replace(canonicalFrom, canonicalFrom + 1, new Slice(Fragment.from(replacement), 0, 0));
  canonical.resolve(canonicalFrom).parent.check();
  return { local: after, canonical, segments, blockPoints };
}

/**
 * Whole pages the walk emits nothing for, removed or put in at the top
 * level: the ledger's transitional sheet holding one continuation copy of a
 * row-sliced atom (a 3-token page), 5 of the 31 failures of the same book.
 * The canonical document is unchanged; local positions after the pages
 * move by their size.
 */
function deriveDroppedSheets(state: MapState, step: Step, after: ProseMirrorNode): MapState | null {
  if (!(step instanceof ReplaceStep)) return null;
  const { from, to, slice } = step;
  const before = state.local;
  if (from < 0 || to > before.content.size || from > to) return null;
  if (slice.openStart !== 0 || slice.openEnd !== 0) return null;
  const page = before.type.schema.nodes.pageBlock;
  if (!page) return null;
  const $from = before.resolve(from), $to = before.resolve(to);
  if ($from.depth !== 0 || $to.depth !== 0) return null;
  // Every page put in holds only copies the walk drops.
  let dropped = true;
  slice.content.forEach(node => {
    if (node.type !== page) { dropped = false; return; }
    node.forEach(child => { if (!droppedRowSliceCopy(child)) dropped = false; });
  });
  if (!dropped) return null;
  // Every page taken out had nothing of the map inside it.
  const { segments, blockPoints } = state;
  if (to > from && (segments.some(segment => segment.localFrom < to && segment.localTo > from) ||
    blockPoints.some(point => point.local > from && point.local < to))) return null;
  for (let index = $from.index(); index < $to.index(); index++) {
    const node = before.child(index);
    if (node.type !== page) return null;
    let kept = false;
    node.forEach(child => { if (!droppedRowSliceCopy(child)) kept = true; });
    if (kept) return null;
  }
  const delta = slice.size - (to - from);
  if (delta === 0 || after.content.size !== before.content.size + delta) return null;
  const shift = (value: number): number | null => {
    if (value <= from) return value;
    if (value >= to) return value + delta;
    return null;
  };
  const nextPoints: BlockPoint[] = [];
  for (const point of blockPoints) {
    const local = shift(point.local);
    if (local === null) return null;
    nextPoints.push({ ...point, local });
  }
  const nextSegments: Segment[] = [];
  for (const segment of segments) {
    const localFrom = shift(segment.localFrom), localTo = shift(segment.localTo);
    if (localFrom === null || localTo === null || localTo - localFrom !== segment.localTo - segment.localFrom) return null;
    nextSegments.push({ ...segment, localFrom, localTo });
  }
  return { local: after, canonical: state.canonical, segments: nextSegments, blockPoints: nextPoints };
}









function deriveDroppedCopies(state: MapState, step: Step, after: ProseMirrorNode): MapState | null {
  if (!(step instanceof ReplaceStep)) return null;
  const { from, to, slice } = step;
  const before = state.local;
  if (from < 0 || to > before.content.size || from > to) return null;
  if (slice.openStart !== 0 || slice.openEnd !== 0) return null;
  const $from = before.resolve(from), $to = before.resolve(to);
  if ($from.depth !== 1 || $to.depth !== 1 || $from.parent.type.name !== "pageBlock" || !$from.sameParent($to)) return null;
  let dropped = true;
  slice.content.forEach(node => { if (!droppedRowSliceCopy(node)) dropped = false; });
  if (!dropped) return null;
  for (let index = $from.index(); index < $to.index(); index++) {
    if (!droppedRowSliceCopy($from.parent.child(index))) return null;
  }
  const { segments, blockPoints } = state;
  if (to > from && (segments.some(segment => segment.localFrom < to && segment.localTo > from) ||
    blockPoints.some(point => point.local > from && point.local < to))) return null;
  const delta = slice.size - (to - from);
  if (delta === 0 || after.content.size !== before.content.size + delta) return null;
  // A copy put in or taken out between two blocks: the end of the block
  // before stays, the start of the block after moves.
  const shift = (value: number, edge: BlockPoint["edge"]): number | null => {
    if (value < from) return value;
    if (value > to) return value + delta;
    if (value === from && (to > from || edge === "close")) return value;
    if (value === to) return value + delta;
    return null;
  };
  const nextPoints: BlockPoint[] = [];
  for (const point of blockPoints) {
    const local = shift(point.local, point.edge);
    if (local === null) return null;
    nextPoints.push({ ...point, local });
  }
  const nextSegments: Segment[] = [];
  for (const segment of segments) {
    const localFrom = shift(segment.localFrom, "open"), localTo = shift(segment.localTo, "close");
    if (localFrom === null || localTo === null || localTo - localFrom !== segment.localTo - segment.localFrom) return null;
    nextSegments.push({ ...segment, localFrom, localTo });
  }
  return { local: after, canonical: state.canonical, segments: nextSegments, blockPoints: nextPoints };
}

/**
 * Whole pages appended at the end of the document: the chunked first fill
 * puts a long document's pages in a chunk at a time, and each chunk rebuilt
 * the map of everything before it (10 of the 40 remaining failures). The
 * new pages are walked on their own, as a document of their own, and their
 * entries appended behind the map's with the local and canonical offsets of
 * the end. Only when nothing joins across the seam: the last block before
 * does not continue, the first block after is no continuation, and the new
 * pages hold no aliased structure (notes, table fragments) the whole-document
 * walk would pair with earlier pages.
 */
function deriveAppendedPages(state: MapState, step: Step, after: ProseMirrorNode): MapState | null {
  if (!(step instanceof ReplaceStep)) return null;
  const { from, to, slice } = step;
  const before = state.local;
  if (from !== to || from !== before.content.size || slice.openStart !== 0 || slice.openEnd !== 0 || slice.content.childCount === 0) return null;
  const page = before.type.schema.nodes.pageBlock;
  if (!page) return null;
  let pages = true;
  slice.content.forEach(node => { if (node.type !== page || node.childCount === 0) pages = false; });
  if (!pages) return null;
  const last = before.lastChild?.lastChild ?? null;
  if (last && CONTINUES.some(key => last.attrs[key] === true)) return null;
  const first = slice.content.firstChild!.firstChild!;
  if (CONTINUATION.some(key => first.attrs[key] === true) || droppedRowSliceCopy(first)) return null;
  let aliased = false;
  slice.content.descendants(node => {
    if (aliased) return false;
    if (node.type.name === "footnoteList" || (node.type.name === "table" && typeof node.attrs.pageTableFragmentId === "string")) aliased = true;
    return !aliased;
  });
  if (aliased) return null;
  if (after.content.size !== before.content.size + slice.size) return null;
  const part = before.type.create(before.attrs, slice.content);
  const walked = buildState(part);
  const localBase = from, canonicalBase = state.canonical.content.size;
  const canonical = state.canonical.replace(canonicalBase, canonicalBase, new Slice(walked.canonical.content, 0, 0));
  if (canonical.content.size !== canonicalBase + walked.canonical.content.size) return null;
  canonical.check();
  return {
    local: after,
    canonical,
    segments: [...state.segments, ...walked.segments.map(segment => ({
      ...segment, localFrom: segment.localFrom + localBase, localTo: segment.localTo + localBase,
      canonicalFrom: segment.canonicalFrom + canonicalBase, canonicalTo: segment.canonicalTo + canonicalBase,
    }))],
    blockPoints: [...state.blockPoints, ...walked.blockPoints.map(point => ({ ...point, local: point.local + localBase, canonical: point.canonical + canonicalBase }))],
  };
}

function isContinuation(node: ProseMirrorNode): boolean { return CONTINUATION.some(key => node.attrs[key] === true); }
function isContinuing(node: ProseMirrorNode): boolean { return CONTINUES.some(key => node.attrs[key] === true); }










function deriveSplitPieceMarkup(state: MapState, step: Step, after: ProseMirrorNode): MapState | null {
  // The markup step: the node's content kept as the gap, one token put
  // around it on each side (setNodeMarkup; its `structure` flag is private).
  if (!(step instanceof ReplaceAroundStep) || step.insert !== 1) return null;
  const { from, to, gapFrom, gapTo, slice } = step;
  if (gapFrom !== from + 1 || gapTo !== to - 1 || slice.content.childCount !== 1 || slice.openStart !== 0 || slice.openEnd !== 0) return null;
  const before = state.local;
  if (from < 0 || to > before.content.size || after.content.size !== before.content.size) return null;
  const old = before.nodeAt(from), next = after.nodeAt(from);
  if (!old || !next || old.type !== next.type || old.nodeSize !== next.nodeSize || old.nodeSize !== to - from || !old.content.eq(next.content)) return null;
  const $at = before.resolve(from);
  if ($at.depth !== 1 || $at.parent.type.name !== "pageBlock" || !TEXT_SPLIT_ATTRS.size) return null;
  // Only the split flags and identity may differ; every other runtime state stays as it was.
  for (const key in { ...old.attrs, ...next.attrs }) {
    if (PAGINATION_RUNTIME_ATTRIBUTES.has(key) && !TEXT_SPLIT_ATTRS.has(key) && JSON.stringify(old.attrs[key] ?? null) !== JSON.stringify(next.attrs[key] ?? null)) return null;
  }
  const { segments, blockPoints } = state;
  if (isContinuation(old) && isContinuation(next)) {
    // A tail stays a tail: nothing of it is in the projection but its content.
    if (isContinuing(old) !== isContinuing(next)) return null;
    return { local: after, canonical: state.canonical, segments, blockPoints };
  }
  // A head written as it was (its flag again, before the seam after it goes).
  if (!isContinuation(old) && isContinuing(old) && old.sameMarkup(next)) {
    return { local: after, canonical: state.canonical, segments, blockPoints };
  }
  if (!isContinuation(old) && !isContinuation(next) && !isContinuing(old) && isContinuing(next)) {
    // A whole textblock becomes a head: its close points go.
    if (!old.isTextblock) return null;
    const closeAt = new Set([to - 1, to]);
    const nextPoints = blockPoints.filter(point => !(point.edge === "close" && closeAt.has(point.local)));
    if (nextPoints.length !== blockPoints.length - 2) return null;
    return { local: after, canonical: state.canonical, segments, blockPoints: nextPoints };
  }
  return null;
}

/**
 * A top-level textblock cut in two where the ledger's page ends, the second
 * piece marked as the continuation (2 of the 40 remaining failures; the cut
 * is made on the page before the piece moves on). The projection joins the
 * pieces back by plain concatenation, so the canonical document is
 * unchanged; local positions after the cut move by the two tokens that
 * close one piece and open the other, a text segment the cut runs through
 * is cut with it, and the pieces' boundary points are those a fresh walk
 * emits: a head that is not marked as continuing still closes at the cut,
 * the tail has none (the walk joins it into the head).
 */
function deriveTextblockSplit(state: MapState, step: Step, after: ProseMirrorNode): MapState | null {
  if (!(step instanceof ReplaceStep)) return null;
  const { from, to, slice } = step;
  if (from !== to || slice.openStart !== 1 || slice.openEnd !== 1 || slice.content.childCount !== 2 || slice.size !== 2) return null;
  const before = state.local;
  if (from < 1 || from > before.content.size) return null;
  const $cut = before.resolve(from);
  if ($cut.depth !== 2 || !$cut.parent.isTextblock || $cut.node(1).type.name !== "pageBlock") return null;
  const block = $cut.parent, start = $cut.before(2), end = $cut.after(2);
  if (after.content.size !== before.content.size + 2) return null;
  const head = after.nodeAt(start);
  const tailNode = after.nodeAt(from + 1);
  if (!head || !tailNode || head.type !== block.type || tailNode.type !== block.type) return null;
  if (!isContinuation(tailNode) || isContinuation(head) !== isContinuation(block)) return null;
  if (head.content.size !== from - (start + 1) || tailNode.content.size !== end - 1 - from) return null;
  if (block.type.name === "codeBlock") return null;
  // The block was one the projection copies by plain concatenation.
  if (hasRuntimeState(block, TEXT_SPLIT_ATTRS) || !plainInline(block.content)) return null;
  for (const key in { ...block.attrs, ...tailNode.attrs, ...head.attrs }) {
    if (!PAGINATION_RUNTIME_ATTRIBUTES.has(key) || TEXT_SPLIT_ATTRS.has(key)) continue;
    if (JSON.stringify(block.attrs[key] ?? null) !== JSON.stringify(head.attrs[key] ?? null)) return null;
  }
  const { segments, blockPoints } = state;
  const contentStart = start + 1;
  const own = segments.filter(segment => segment.localFrom >= contentStart && segment.localTo <= end - 1);
  if (own.some(segment => segment.replica)) return null;
  // The block's own close points carry the canonical end of the joined block.
  const closes = blockPoints.filter(point => point.edge === "close" && (point.local === end - 1 || point.local === end));
  const canonicalEnd = closes.find(point => point.local === end)?.canonical ?? null;
  if (!isContinuing(block) && (closes.length !== 2 || canonicalEnd === null)) return null;
  const nextSegments: Segment[] = [];
  for (const segment of segments) {
    if (segment.localTo <= from) { nextSegments.push(segment); continue; }
    if (segment.localFrom >= from) { nextSegments.push({ ...segment, localFrom: segment.localFrom + 2, localTo: segment.localTo + 2 }); continue; }
    // Cut through: a text segment only.
    if (!own.includes(segment) || segment.localTo - segment.localFrom !== segment.canonicalTo - segment.canonicalFrom) return null;
    const cutAt = from - segment.localFrom;
    nextSegments.push({ ...segment, localTo: from, canonicalTo: segment.canonicalFrom + cutAt });
    nextSegments.push({ ...segment, localFrom: from + 2, localTo: segment.localTo + 2, canonicalFrom: segment.canonicalFrom + cutAt });
  }
  const nextPoints: BlockPoint[] = [];
  for (const point of blockPoints) {
    if (closes.includes(point)) {
      // The head's close pair moves to the cut (when the head does not
      // continue), still at the joined block's canonical end. The tail is
      // joined into the head by the walk and has no points of its own.
      if (point.local === end - 1 && !isContinuing(head)) {
        nextPoints.push({ ...point, local: from }, { ...point, local: from + 1, canonical: canonicalEnd! });
      }
      continue;
    }
    // The head's own points (its open pair, at the cut when the head is
    // emptied) stay; everything after the cut moves by the two tokens.
    nextPoints.push(point.local <= from ? point : { ...point, local: point.local + 2 });
  }
  // Both pieces' close pairs must read as a fresh walk's: (end-1 → canonicalEnd-1, end → canonicalEnd).
  if (closes.length === 2) {
    const first = closes.find(point => point.local === end - 1)!;
    if (first.canonical !== canonicalEnd! - 1) return null;
  }
  return { local: after, canonical: state.canonical, segments: nextSegments, blockPoints: nextPoints };
}

/**
 * Two pieces of a split textblock joined across their sheets' boundary, the tail
 * opening the second sheet: the pieces of a paragraph as long as a book, joined
 * as the page ledger reaches them (ledgerGiantBlocks), each such pass of a long
 * book otherwise rebuilding the whole map. The projection had them joined
 * already, so the canonical document is unchanged and local positions after the
 * seam move back by its four tokens. A head that goes on continuing keeps no
 * close points and a tail had no open ones, so none is added or taken away; two
 * text runs the join makes one are one segment, as a fresh walk emits them.
 */
function deriveJoinedPieces(state: MapState, step: Step, after: ProseMirrorNode): MapState | null {
  if (!(step instanceof ReplaceStep)) return null;
  const { from, to, slice } = step;
  const before = state.local;
  if (to !== from + 4 || slice.size !== 0 || from < 2 || to > before.content.size) return null;
  if (after.content.size !== before.content.size - 4) return null;
  const $from = before.resolve(from), $to = before.resolve(to);
  if ($from.depth !== 2 || $to.depth !== 2 || $from.node(1).type.name !== "pageBlock" || $to.node(1).type.name !== "pageBlock") return null;
  const head = $from.parent, tail = $to.parent;
  if (!head.isTextblock || head.type !== tail.type || head.type.name === "codeBlock") return null;
  if ($from.parentOffset !== head.content.size || $to.parentOffset !== 0) return null;
  // Empty pieces change what the merge does.
  if (head.content.size === 0 || tail.content.size === 0) return null;
  if (head.attrs.pageSplitContinues !== true || tail.attrs.pageSplitContinuation !== true || tail.attrs.pageSplitContinues !== true) return null;
  if (hasRuntimeState(head, JOINED_PIECE_ATTRS) || hasRuntimeState(tail, JOINED_PIECE_ATTRS)) return null;
  if (!plainInline(head.content) || !plainInline(tail.content)) return null;
  const blockPoints: BlockPoint[] = [];
  for (const point of state.blockPoints) {
    if (point.local > from && point.local < to) return null;
    blockPoints.push(point.local >= to ? { ...point, local: point.local - 4 } : point);
  }
  const lastOfHead = head.lastChild, firstOfTail = tail.firstChild;
  const oneRun = Boolean(lastOfHead?.isText && firstOfTail?.isText && lastOfHead.sameMarkup(firstOfTail));
  const segments: Segment[] = [];
  let joined = false;
  for (const segment of state.segments) {
    if (segment.localFrom < to && segment.localTo > from) return null;
    if (segment.localFrom < to) { segments.push(segment); continue; }
    const moved = { ...segment, localFrom: segment.localFrom - 4, localTo: segment.localTo - 4 };
    const previous = segments[segments.length - 1];
    if (oneRun && !joined && segment.localFrom === to) {
      if (!previous || previous.localTo !== from || previous.canonicalTo !== segment.canonicalFrom || previous.replica || segment.replica) return null;
      segments[segments.length - 1] = { ...previous, localTo: moved.localTo, canonicalTo: segment.canonicalTo };
      joined = true;
      continue;
    }
    segments.push(moved);
  }
  if (oneRun && !joined) return null;
  return { local: after, canonical: state.canonical, segments, blockPoints };
}
