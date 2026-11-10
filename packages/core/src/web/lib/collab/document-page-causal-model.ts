/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { JSONContent } from "@tiptap/core";
import * as Y from "yjs";

import { DOCUMENT_ATTRIBUTE_CAPABILITIES } from "../document-feature-capabilities";
import {
  COLLABORATIVE_DOCUMENT_ATTRS_MAP,
  readCollaborativeDocumentAttrs,
} from "../../components/console/editor/native-document/collaborationDocumentAttrs";
import { serializeWhiteboardScenes } from "../../components/console/editor/extensions/whiteboardScene";
import { jsonEquals, type JsonValue } from "./json-crdt";
import { PAGINATION_RUNTIME_ATTRIBUTES } from "./paged-document-projection";
import nullDefaults from "./document-canonical-null-attributes.json";
import { yDocRevision, yTypeRevision } from "./y-doc-revision";
import { holdsForeignCanonicalXmlAttributes } from "./canonical-xml-attributes";

const PRIVATE_LAYOUT = new Set(["pageBlockSizes", "pageBlockLayoutKey", "pageRegionLayoutReceipt"]);
const ROOTS = new Set(["default", COLLABORATIVE_DOCUMENT_ATTRS_MAP]);
const nullSets = (entries: Record<string, string[]>) =>
  new Map(Object.entries(entries).map(([name, keys]) => [name, new Set(keys)]));
const NULL_DEFAULTS = nullSets(nullDefaults.nodes);
// A mark attribute is an attribute: ProseMirror fills a missing one with the
// schema default, so an explicit null and an absent key are the same mark. An
// importer that omits `characterStyle.name` and an editor that states it null
// must project identically, or the page can never mirror and never saves.
const NULL_MARK_DEFAULTS = nullSets(nullDefaults.marks);
const fail = (): never => { throw new Error("Invalid canonical document projection"); };
const record = (value: unknown): value is Record<string, unknown> =>
  Boolean(value && typeof value === "object" && !Array.isArray(value));
// The projection is read-only JSON: attribute values are carried by
// reference (a deep clone of every mark and node attribute -- an imported
// page's style catalog, an inlined picture -- on every projection made each
// save cost seconds on a long document); only the containers are new.
const markAttrs = (type: string, attrs: Record<string, unknown> | undefined) =>
  Object.fromEntries(Object.entries(attrs ?? {}).filter(([key, value]) =>
    !(value === null && NULL_MARK_DEFAULTS.get(type)?.has(key))));
// Optional undefined object fields disappear on the JSON wire and are absent
// to jsonEquals. They must not change mark order before recovery is encoded.
const markKey = (value: unknown): string => {
  if (Array.isArray(value)) return `[${value.map(markKey).join(",")}]`;
  if (record(value)) return `{${Object.keys(value).filter(key => value[key] !== undefined).sort().map(key => `${JSON.stringify(key)}:${markKey(value[key])}`).join(",")}}`;
  return JSON.stringify(value);
};

/** Storage and collaboration share semantic content. Private pagination
 * receipts are carried by the native editor and verified by its own layout
 * pipeline; neither a checkpoint nor this projection proves their geometry. */
export function documentCausalSemanticContent(value: unknown): Record<string, JsonValue> {
  // `importLosses` is the importer's record of what the file holds and the page
  // does not (lib/import-losses.ts). The server writes it once beside
  // type/attrs/content and no editor ever does, so it is no part of the
  // collaborative document. Read as an unknown root it failed this projection
  // for every imported page that recorded a loss, and the realtime session of
  // such a page never finished preparing. The API model leaves it out the same
  // way (document_page_causal_crdt.document_semantic_content).
  if (record(value) && Object.hasOwn(value, "importLosses")) {
    const { importLosses: _importLosses, ...authored } = value;
    value = authored;
  }
  requireDocumentRoot(value);
  return semanticNormalizer({ left: SEMANTIC_BUDGET })(value, 0) as Record<string, JsonValue>;
}






const SEMANTIC_BUDGET = 1_200_000;

function requireDocumentRoot(value: unknown): asserts value is Record<string, unknown> {
  if (!record(value) || value.type !== "doc" || !Array.isArray(value.content) ||
      Object.keys(value).some(key => !["type", "content", "attrs"].includes(key)) ||
      (value.attrs !== undefined && (!record(value.attrs) ||
        Object.keys(value.attrs).some(key => !Object.hasOwn(DOCUMENT_ATTRIBUTE_CAPABILITIES, key))))) fail();
}

/** The semantic form of one node and everything in it; each node visited
 * spends one of `budget` (the whole document's, shared by its blocks). */
function semanticNormalizer(budget: { left: number }) {
  const normalize = (raw: unknown, depth: number): JSONContent => {
    if (--budget.left < 0 || depth > 256 || !record(raw) || typeof raw.type !== "string" || raw.type === "pageBlock") return fail();
    const node = { ...raw } as JSONContent;
    if (node.type === "text" && (typeof node.text !== "string" || node.text.length === 0)) fail();
    if (node.attrs !== undefined) {
      if (!record(node.attrs)) fail();
      // An attribute the editor holds as `undefined` (a schema default the
      // importer never wrote, e.g. a shape's inkTool inside a table cell) is
      // no JSON value at all: the shared document cannot carry it, and a
      // comparison that counted the key refused every save of such a page.
      node.attrs = Object.fromEntries(Object.entries(node.attrs).filter(([key, value]) =>
        value !== undefined && !PAGINATION_RUNTIME_ATTRIBUTES.has(key) &&
        !(value === null && NULL_DEFAULTS.get(node.type!)?.has(key))));
      if (!Object.keys(node.attrs).length) delete node.attrs;
    }
    // The shared document carries a mark as a format of text. A mark the
    // editor leaves on an inline node (a hard break or an inline object inside
    // a selection that was underlined) has no place there, so the room never
    // says it; counting it refused every save of a page after Shift+Enter and
    // a format across the break.
    if (node.type !== "text" && Array.isArray(node.marks)) delete node.marks;
    if (node.marks !== undefined) {
      if (!Array.isArray(node.marks) || node.marks.some(mark => !record(mark) ||
          typeof mark.type !== "string" || (mark.attrs !== undefined && !record(mark.attrs)))) fail();
      // Each mark's key is built once, not once per comparison.
      node.marks = node.marks.map(mark => {
        const projected = { type: mark.type, attrs: markAttrs(mark.type, mark.attrs) };
        return { projected, key: markKey(projected) };
      }).sort((a, b) => a.key < b.key ? -1 : a.key > b.key ? 1 : 0).map(entry => entry.projected);
      if (!node.marks.length) delete node.marks;
    }
    if (node.content !== undefined) {
      if (!Array.isArray(node.content)) fail();
      const children: JSONContent[] = [];
      for (const rawChild of node.content) {
        const child = normalize(rawChild, depth + 1), previous = children.at(-1);
        if (child.type === "text" && previous?.type === "text" &&
            jsonEquals((child.marks ?? null) as JsonValue, (previous.marks ?? null) as JsonValue)) previous.text! += child.text!;
        else children.push(child);
      }
      node.content = children;
      if (!children.length && node.type !== "doc") delete node.content;
    }
    return node;
  };
  return normalize;
}

/** y-tiptap's yXmlFragmentToProsemirrorJSON, for one child of the fragment. */
function serializeYNode(item: unknown): JSONContent | JSONContent[] {
  if (item instanceof Y.XmlText) {
    return (item.toDelta() as Array<{ insert: unknown; attributes?: Record<string, Record<string, unknown>> }>).map((delta) => {
      const text: JSONContent = { type: "text", text: delta.insert as string };
      if (delta.attributes) {
        text.marks = Object.keys(delta.attributes).map((type) => ({
          type: hashedMarkName.exec(type)?.[1] ?? type, attrs: delta.attributes![type],
        }));
      }
      return text;
    });
  }
  if (item instanceof Y.XmlElement) {
    const response: JSONContent = { type: item.nodeName };
    const attrs = item.getAttributes();
    if (Object.keys(attrs).length) response.attrs = attrs;
    const children = item.toArray();
    if (children.length) response.content = children.map(serializeYNode).flat();
    return response;
  }
  return fail();
}
const hashedMarkName = /(.*)(--[a-zA-Z0-9+/=]{8})$/;

/** What the projection of a document holds for one top-level element. */
interface BlockProjection {
  /** The block's semantic JSON (frozen containers; attribute values are the
   * room's own, by reference, as in every projection). */
  readonly node: JSONContent;
  /** Nodes the structural check visited / the semantic form visited. */
  readonly checked: number;
  readonly normalized: number;
}

const STRUCTURE_BUDGET = SEMANTIC_BUDGET;

function projectBlock(child: unknown): BlockProjection {
  let checked = 0;
  const check = (node: unknown, depth: number): void => {
    if (++checked > STRUCTURE_BUDGET || depth > 256) fail();
    if (node instanceof Y.XmlText) {
      if (node.toDelta().some((run: { insert: unknown; attributes?: Record<string, unknown> }) =>
        typeof run.insert !== "string" || Object.values(run.attributes ?? {}).some(value => !record(value)))) fail();
      return;
    }
    // A layout flag left at its unset value (false / null) says nothing: a
    // room seeded from a body that stated `pageSplitPieceOf: null` on every
    
    // was refused here and the window stayed on "session changed". A set
    // flag -- a real page-split piece in the room -- is still refused.
    if (!(node instanceof Y.XmlElement) || node.nodeName === "pageBlock" ||
        Object.entries(node.getAttributes()).some(([key, value]) =>
          PAGINATION_RUNTIME_ATTRIBUTES.has(key) && value !== null && value !== false && value !== undefined)) fail();
    if (node instanceof Y.XmlElement) node.toArray().forEach(grandchild => check(grandchild, depth + 1));
  };
  check(child, 0);
  const budget = { left: SEMANTIC_BUDGET };
  const node = semanticNormalizer(budget)(serializeWhiteboardScenes(serializeYNode(child)), 1);
  return { node: freezeProjectionNode(node), checked, normalized: SEMANTIC_BUDGET - budget.left };
}

/** Freeze the containers a projection made (never the attribute values: those
 * are the room's own objects, which the editor holds too). */
function freezeProjectionNode(node: JSONContent): JSONContent {
  if (node.attrs) Object.freeze(node.attrs);
  if (node.marks) { for (const mark of node.marks) { if (mark.attrs) Object.freeze(mark.attrs); Object.freeze(mark); } Object.freeze(node.marks); }
  if (node.content) { for (const child of node.content) freezeProjectionNode(child); Object.freeze(node.content); }
  return Object.freeze(node);
}

/** A fresh copy of a projection's containers, for a consumer that owns it. */
function copyProjectionNode(node: JSONContent): JSONContent {
  const copy = { ...node };
  if (node.attrs) copy.attrs = { ...node.attrs };
  if (node.marks) copy.marks = node.marks.map((mark) => ({ ...mark, ...(mark.attrs ? { attrs: { ...mark.attrs } } : {}) }));
  if (node.content) copy.content = node.content.map(copyProjectionNode);
  return copy;
}

/**
 * A document's projection, kept per top-level element. Reading the whole
 * room on each save (its check, its JSON and its semantic form) cost tens of
 * milliseconds on an inked page, several times a save, and grew with every
 * stroke; an element nothing has changed since it was read is not read
 * again. Yjs reports every element a transaction changed, its ancestors
 * included (changedParentTypes); such an element is read anew. While any
 * transaction of the document is open, or its observers still run (they run
 * before the transaction is reported), nothing kept is used.
 */
interface ProjectionCache {
  readonly blocks: WeakMap<object, BlockProjection>;
  readonly open: Set<Y.Transaction>;
}
const projectionCaches = new WeakMap<Y.Doc, ProjectionCache>();

function projectionCache(doc: Y.Doc): ProjectionCache {
  const cache = projectionCaches.get(doc);
  if (cache) return cache;
  const created: ProjectionCache = { blocks: new WeakMap(), open: new Set() };
  doc.on("beforeTransaction", (transaction: Y.Transaction) => { created.open.add(transaction); });
  doc.on("afterTransaction", (transaction: Y.Transaction) => {
    for (const type of transaction.changedParentTypes.keys()) created.blocks.delete(type);
    for (const type of transaction.changed.keys()) created.blocks.delete(type);
    created.open.delete(transaction);
  });
  projectionCaches.set(doc, created);
  return created;
}

/** The document's attributes as read from its map, kept while the map does
 * not change (its values are then the same objects from read to read, and a
 * comparison of them is remembered: documentContentSaysProjection). */
const documentAttrsReads = new WeakMap<object, { revision: number; attrs: Readonly<Record<string, unknown>> }>();
function readDocumentAttrs(map: Y.Map<unknown>): Readonly<Record<string, unknown>> {
  const revision = yTypeRevision(map);
  const known = documentAttrsReads.get(map);
  if (revision !== null && known?.revision === revision) return known.attrs;
  const attrs = readCollaborativeDocumentAttrs(map);
  if (revision !== null && yTypeRevision(map) === revision) documentAttrsReads.set(map, { revision, attrs });
  return attrs;
}

function readProjectionBlocks(doc: Y.Doc): { root: Record<string, JsonValue>; blocks: JSONContent[] } {
  if ([...doc.share.keys()].some(key => !ROOTS.has(key))) fail();
  const attrsMap = doc.getMap(COLLABORATIVE_DOCUMENT_ATTRS_MAP);
  if ([...attrsMap.keys()].some(key => PRIVATE_LAYOUT.has(key) || !Object.hasOwn(DOCUMENT_ATTRIBUTE_CAPABILITIES, key))) fail();
  const cache = projectionCache(doc);
  const reuse = cache.open.size === 0 && !(doc as unknown as { _transaction: unknown })._transaction;
  const children = doc.getXmlFragment("default").toArray();
  const projected: BlockProjection[] = [];
  let checked = 0;
  for (const child of children) {
    let block = reuse ? cache.blocks.get(child) : undefined;
    if (!block) {
      block = projectBlock(child);
      if (reuse) cache.blocks.set(child, block);
    }
    checked += block.checked;
    if (checked > STRUCTURE_BUDGET) fail();
    projected.push(block);
  }
  const raw = { type: "doc", content: [] as JSONContent[], attrs: readDocumentAttrs(attrsMap) };
  requireDocumentRoot(raw);
  const budget = { left: SEMANTIC_BUDGET };
  const root = semanticNormalizer(budget)(raw, 0);
  let normalized = SEMANTIC_BUDGET - budget.left;
  for (const block of projected) {
    normalized += block.normalized;
    if (normalized > SEMANTIC_BUDGET) fail();
  }
  return { root: root as Record<string, JsonValue>, blocks: projected.map((block) => block.node) };
}

/**
 * Read a few more of the room's top-level blocks into the kept projection
 * (what readProjectionBlocks keeps per element), from block `start`, within
 * `budgetMs`; each block read is handed to `onBlock` (saveTextWarmup keeps
 * what the save's digest reads of it). The index reached, or -1 when the
 * room cannot be read ahead now (a transaction open, a block the read
 * refuses: the save reads it as before).
 */
export function warmDocumentProjectionBlocks(
  doc: Y.Doc, start: number, budgetMs: number, onBlock: (block: JSONContent, index: number) => void,
): number {
  const cache = projectionCache(doc);
  if (cache.open.size > 0 || (doc as unknown as { _transaction: unknown })._transaction) return -1;
  const children = doc.getXmlFragment("default").toArray();
  const end = performance.now() + budgetMs;
  let index = start;
  try {
    while (index < children.length) {
      const child = children[index++];
      let block = cache.blocks.get(child);
      if (!block) cache.blocks.set(child, (block = projectBlock(child)));
      onBlock(block.node, index - 1);
      // The time goes in what is read of a block as much as in its own read
      // (a projection another read already kept is read here by onBlock).
      if (performance.now() >= end) break;
    }
  } catch {
    return -1;
  }
  return index;
}

/** Read the actual flat XML and the existing synthetic attribute map. Never
 * flatten a legacy physical fragment and call it a coherent new original. */
export function readDocumentCausalProjection(doc: Y.Doc): Record<string, JsonValue> {
  const { root, blocks } = readProjectionBlocks(doc);
  root.content = blocks.map(copyProjectionNode) as unknown as JsonValue;
  // The document's attribute values are kept from read to read: a consumer
  // gets its own.
  if (record(root.attrs)) root.attrs = structuredClone(root.attrs);
  return root;
}

/**
 * The projection of a document at its current revision, read once per
 * revision for the save path's own comparisons (a capture compared the same
 * unchanged document three times). The result is shared: read it, never
 * mutate it or hand it to a consumer. Its blocks are kept from one revision
 * to the next while their elements do not change (frozen).
 */
const sharedProjections = new WeakMap<Y.Doc, { revision: string; projection: Record<string, JsonValue> }>();
export function readSharedDocumentCausalProjection(doc: Y.Doc): Record<string, JsonValue> {
  const revision = yDocRevision(doc);
  const read = () => {
    const { root, blocks } = readProjectionBlocks(doc);
    root.content = Object.freeze(blocks) as unknown as JsonValue;
    return root;
  };
  if (revision === null) return read();
  const known = sharedProjections.get(doc);
  if (known?.revision === revision) return known.projection;
  let projection = read();
  // A transaction that changed nothing (the save's own repair pass opens one
  // every time) leaves the same blocks and attributes: the same projection,
  // so what was proven against it still holds.
  if (known && sameProjection(known.projection, projection)) projection = known.projection;
  if (yDocRevision(doc) === revision) sharedProjections.set(doc, { revision, projection });
  return projection;
}

function sameProjection(a: Record<string, JsonValue>, b: Record<string, JsonValue>): boolean {
  const left = a.content as unknown as JSONContent[], right = b.content as unknown as JSONContent[];
  if (a.type !== b.type || left.length !== right.length || !left.every((block, index) => block === right[index])) return false;
  const x = (a.attrs ?? null) as Record<string, unknown> | null, y = (b.attrs ?? null) as Record<string, unknown> | null;
  if (x === y) return true;
  if (!x || !y) return false;
  const keys = Object.keys(x);
  return keys.length === Object.keys(y).length && keys.every(key => Object.hasOwn(y, key) && x[key] === y[key]);
}

/** The shared projection a content object was last found to mirror exactly
 * (`documentCausalContentMirrors` returned true for that very object). */
const mirroredBy = new WeakMap<object, Record<string, JsonValue>>();

/**
 * Whether `content` says what `projection` (a document's projection) says:
 * `jsonEquals(projection, documentCausalSemanticContent(content))`, block by
 * block. Each save compared the editor's whole body with the whole room,
 * twice a save, and grew with every stroke on the page; a block is now told
 * apart by its JSON text (what the save sends of it): its semantic form is
 * kept per text, and a semantic block already found equal to the very same
 * (unchanged, kept) block of the room is not compared again.
 */
export function documentContentSaysProjection(content: unknown, projection: unknown): boolean {
  if (record(content) && Object.hasOwn(content, "importLosses")) {
    const { importLosses: _importLosses, ...authored } = content;
    content = authored;
  }
  requireDocumentRoot(content);
  const blocks = content.content as unknown[];
  if (!record(projection) || !Array.isArray(projection.content) ||
      blocks.some(block => !record(block) || block.type === "text")) {
    // Not a document of blocks: compared whole, as ever.
    return jsonEquals(documentCausalSemanticContent(content), projection as JsonValue);
  }
  const room = projection.content as JsonValue[];
  const budget = { left: SEMANTIC_BUDGET };
  const root = semanticNormalizer(budget)({ ...content, content: [] }, 0);
  if (!sameDocumentRoot(root as Record<string, unknown>, projection) || room.length !== blocks.length) {
    // Still normalized whole: a body the semantic form refuses is refused.
    for (const block of blocks) semanticBlock(block, budget);
    return false;
  }
  rotateSemanticBlocks();
  let equal = true;
  for (let index = 0; index < blocks.length; index += 1) {
    const semantic = semanticBlock(blocks[index], budget);
    if (!equal) continue;
    if (provenEqual.get(semantic) === room[index]) continue;
    if (!jsonEquals(semantic as JsonValue, room[index])) { equal = false; continue; }
    if (Object.isFrozen(room[index])) provenEqual.set(semantic, room[index] as object);
  }
  return equal;
}

/** The two documents' own keys and attributes (not their blocks) are equal.
 * An attribute value of the editor's (ProseMirror's own, never changed) found
 * equal to the room's kept value of the same key is not compared again: the
 * document's style catalogues are tens of kilobytes. */
function sameDocumentRoot(semantic: Record<string, unknown>, projection: Record<string, unknown>): boolean {
  const keys = (value: Record<string, unknown>) => Object.keys(value).filter(key => key !== "content" && value[key] !== undefined);
  const left = keys(semantic), right = keys(projection);
  if (left.length !== right.length || left.some(key => !right.includes(key))) return false;
  for (const key of left) {
    if (key !== "attrs") {
      if (!jsonEquals(semantic[key] as JsonValue, projection[key] as JsonValue)) return false;
      continue;
    }
    const a = semantic.attrs, b = projection.attrs;
    if (!record(a) || !record(b)) { if (!jsonEquals(a as JsonValue, b as JsonValue)) return false; continue; }
    const names = Object.keys(a).filter(name => a[name] !== undefined);
    if (names.length !== Object.keys(b).filter(name => b[name] !== undefined).length) return false;
    for (const name of names) {
      const x = a[name], y = b[name];
      if (!Object.hasOwn(b, name) || y === undefined) return false;
      if (x === y) continue;
      const pair = x !== null && typeof x === "object" && y !== null && typeof y === "object";
      if (pair && provenEqualValue.get(x) === y) continue;
      if (!jsonEquals(x as JsonValue, y as JsonValue)) return false;
      if (pair) provenEqualValue.set(x, y);
    }
  }
  return true;
}
const provenEqualValue = new WeakMap<object, unknown>();

/** Semantic blocks by the JSON text of the editor's block, kept for the last
 * two comparisons (what one save cycle compares again). */
let semanticBlocks = new Map<string, { node: JSONContent; normalized: number }>();
let previousSemanticBlocks = new Map<string, { node: JSONContent; normalized: number }>();
function rotateSemanticBlocks() {
  previousSemanticBlocks = semanticBlocks;
  semanticBlocks = new Map();
}
/** A kept semantic block -> the frozen room block it was found equal to. */
const provenEqual = new WeakMap<object, object>();

/** Prove ahead of a save that the editor's frozen block says what a frozen
 * room block does (saveTextWarmup), keeping what documentContentSaysProjection
 * keeps when it finds them equal: the editor block's semantic form, and that
 * form found equal to the room block. Nothing is kept for blocks that differ. */
export function warmMirroredBlock(editorBlock: unknown, roomBlock: unknown): void {
  if (!record(editorBlock) || !Object.isFrozen(editorBlock) || !record(roomBlock) || !Object.isFrozen(roomBlock)) return;
  try {
    const semantic = semanticBlock(editorBlock, { left: SEMANTIC_BUDGET });
    if (provenEqual.get(semantic) === roomBlock) return;
    if (jsonEquals(semantic as JsonValue, roomBlock as JsonValue)) provenEqual.set(semantic, roomBlock);
  } catch {
    // The save compares it as before.
  }
}

/** Semantic blocks of frozen blocks (the editor's serialization keeps each
 * untouched block's JSON as the same frozen object), by the block itself. */
const semanticOfFrozen = new WeakMap<object, { node: JSONContent; normalized: number }>();

function semanticBlock(block: unknown, budget: { left: number }): JSONContent {
  if (record(block) && Object.isFrozen(block)) {
    let entry = semanticOfFrozen.get(block);
    if (!entry) {
      const own = { left: SEMANTIC_BUDGET };
      const node = semanticNormalizer(own)(block, 1);
      entry = { node: freezeProjectionNode(node), normalized: SEMANTIC_BUDGET - own.left };
      semanticOfFrozen.set(block, entry);
    }
    budget.left -= entry.normalized;
    if (budget.left < 0) fail();
    return entry.node;
  }
  const text = JSON.stringify(block);
  let entry = semanticBlocks.get(text);
  if (!entry) {
    entry = previousSemanticBlocks.get(text);
    if (!entry) {
      const own = { left: SEMANTIC_BUDGET };
      const node = semanticNormalizer(own)(block, 1);
      entry = { node: freezeProjectionNode(node), normalized: SEMANTIC_BUDGET - own.left };
    }
    semanticBlocks.set(text, entry);
  }
  budget.left -= entry.normalized;
  if (budget.left < 0) fail();
  return entry.node;
}

/**
 * ``content`` with the attributes the room holds on its nodes and the
 * editor's own JSON does not state at all: a newer client's, which this
 * client's schema does not declare and passes through (the room keeps them
 * through its edits, canonical-xml-attributes). The editor states every
 * attribute it declares, so an attribute it leaves out entirely is one it
 * cannot hold; the save body carries it as the room does, and the API's
 * write gate keeps it when the API declares it. Paired by position, element
 * by element, only where the two say the same types; copy on write.
 */
export function withRoomOnlyAttributes<T>(content: T, projection: unknown): T {
  const elements = (list: unknown) => Array.isArray(list) ? list.filter(child => record(child) && child.type !== "text") : [];
  const graft = (node: unknown, room: unknown): unknown => {
    if (!record(node) || !record(room) || node.type !== room.type) return node;
    let next: Record<string, unknown> = node;
    if (record(room.attrs)) {
      const own = record(node.attrs) ? node.attrs : {};
      const missing = Object.entries(room.attrs).filter(([key]) => !Object.hasOwn(own, key));
      if (missing.length) next = { ...next, attrs: { ...own, ...Object.fromEntries(missing) } };
    }
    if (Array.isArray(node.content)) {
      const mine = elements(node.content), theirs = elements(room.content);
      if (mine.length && mine.length === theirs.length) {
        const pairs = new Map(mine.map((child, index) => [child, graft(child, theirs[index])]));
        if ([...pairs].some(([from, to]) => from !== to)) {
          next = { ...next, content: (node.content as unknown[]).map(child => pairs.get(child) ?? child) };
        }
      }
    }
    return next;
  };
  if (!record(content) || !record(projection) || !Array.isArray(content.content)) return content;
  const blocks = content.content as unknown[], room = elements(projection.content);
  if (blocks.length !== room.length) return content;
  let changed = false;
  const grafted = blocks.map((block, index) => {
    const result = graft(block, room[index]);
    if (result !== block) changed = true;
    return result;
  });
  return (changed ? { ...content, content: grafted } : content) as T;
}

/** Projections of rooms holding passed-through attributes, and the body each
 * editor body proven to mirror one was grafted into. */
const foreignProjections = new WeakSet<object>();
const mirroredBodies = new WeakMap<object, unknown>();

/** What a save of ``requested`` sends against ``projection``: the body with
 * the room's passed-through attributes (see withRoomOnlyAttributes). */
export function documentSaveBodyFor<T>(requested: T, projection: unknown): T {
  if (record(requested) && mirroredBy.get(requested) === projection && mirroredBodies.has(requested)) {
    return mirroredBodies.get(requested) as T;
  }
  return record(projection) && foreignProjections.has(projection) ? withRoomOnlyAttributes(requested, projection) : requested;
}

/** The codec's mirror predicate: the editor's own JSON and the shared document
 * must already say the same thing before a local write is mirrored into Yjs. */
export function documentCausalContentMirrors(content: unknown, doc: Y.Doc): boolean {
  try {
    const projection = readSharedDocumentCausalProjection(doc);
    const foreign = holdsForeignCanonicalXmlAttributes(doc);
    if (foreign) foreignProjections.add(projection);
    const body = foreign ? withRoomOnlyAttributes(content, projection) : content;
    // A copy this module made of a body proven to mirror this very
    // projection (the same document revision) says what it says.
    const mirrors = (record(content) && copiedFrom.get(content) === projection) ||
      documentContentSaysProjection(body, projection);
    if (record(content)) {
      if (mirrors) mirroredBy.set(content, projection);
      else mirroredBy.delete(content);
      if (mirrors && body !== content) mirroredBodies.set(content, body);
      else mirroredBodies.delete(content);
    }
    lastMirrorMiss = mirrors ? null : documentMirrorMissPath(body, projection);
    return mirrors;
  } catch {
    // The semantic form refused the body or the room (an unknown document
    // attribute, an empty text node, the node budget): say so, rather than
    // leave the previous comparison's answer for the save diagnostics.
    lastMirrorMiss = "comparison_threw";
    return false;
  }
}

let lastMirrorMiss: string | null = null;

/** Where the last body refused by documentCausalContentMirrors first differs
 * from the room, as a content-free path for save diagnostics
 * (lib/page-save-diagnostics): block indices and attribute names only, never
 * a value, e.g. `content[29].attrs.source` or `content length 31/30`. */
export function lastDocumentMirrorMiss(): string | null {
  return lastMirrorMiss;
}

const MIRROR_MISS_DEPTH = 6;

function documentMirrorMissPath(content: unknown, projection: unknown): string | null {
  try {
    let authored = content;
    if (record(authored) && Object.hasOwn(authored, "importLosses")) {
      const { importLosses: _importLosses, ...rest } = authored;
      authored = rest;
    }
    const semantic = documentCausalSemanticContent(authored as JSONContent) as unknown;
    const walk = (a: unknown, b: unknown, path: string, depth: number): string | null => {
      if (a === b) return null;
      if (depth >= MIRROR_MISS_DEPTH) return jsonEquals(a as JsonValue, b as JsonValue) ? null : path;
      if (Array.isArray(a) && Array.isArray(b)) {
        if (a.length !== b.length) return `${path} length ${a.length}/${b.length}`;
        for (let index = 0; index < a.length; index += 1) {
          const found = walk(a[index], b[index], `${path}[${index}]`, depth + 1);
          if (found) return found;
        }
        return null;
      }
      if (record(a) && record(b)) {
        const keys = (value: Record<string, unknown>) => Object.keys(value).filter(key => value[key] !== undefined);
        const left = keys(a), right = keys(b);
        // Attribute names are the schema's; a map keyed by the document's
        // own names (style catalogues) is reported by its count only.
        const named = path.endsWith(".attrs") || path === "" || /\]$/.test(path);
        const onlyEditor = left.filter(key => !right.includes(key));
        const onlyRoom = right.filter(key => !left.includes(key));
        if (onlyEditor.length || onlyRoom.length) {
          return named
            ? `${path} keys editor+${onlyEditor.slice(0, 4).join("|")} room+${onlyRoom.slice(0, 4).join("|")}`
            : `${path} keys ${onlyEditor.length}/${onlyRoom.length}`;
        }
        for (const key of left) {
          const segment = named ? (path ? `${path}.${key}` : key) : `${path}.*`;
          const found = walk(a[key], b[key], segment, depth + 1);
          if (found) return found;
        }
        return null;
      }
      return jsonEquals(a as JsonValue, b as JsonValue) ? null : `${path || "$"} ${typeof a}/${typeof b}`;
    };
    return walk(semantic, projection, "", 0)?.slice(0, 200) ?? "no_semantic_difference";
  } catch {
    return "semantic_form_refused";
  }
}

/** A body's copy made for a save (contentForSave), and the projection the
 * body was proven to mirror when it was made. Nothing writes into a save
 * body; a copy of one is handed on as it is. */
const copiedFrom = new WeakMap<object, Record<string, JsonValue>>();
const ownCopies = new WeakSet<object>();










export function copySharingFrozen<T>(value: T): T {
  if (value === null || typeof value !== "object" || Object.isFrozen(value)) return value;
  if (Array.isArray(value)) return value.map(copySharingFrozen) as T;
  const copy: Record<string, unknown> = {};
  for (const key of Object.keys(value)) copy[key] = copySharingFrozen((value as Record<string, unknown>)[key]);
  return copy as T;
}

/** The save's own copy of `requested`, proven to mirror `projection`. */
export function documentSaveCopy(requested: Record<string, JsonValue>, projection: Record<string, JsonValue>): Record<string, JsonValue> {
  const copy = ownCopies.has(requested) ? requested
    : (globalThis as { __mraiCloneSaveCopy?: boolean }).__mraiCloneSaveCopy ? structuredClone(requested) : copySharingFrozen(requested);
  ownCopies.add(copy);
  copiedFrom.set(copy, projection);
  return copy;
}

/** `body` is a save body its maker hands over and never writes into again
 * (the editor's serialization of one save): it needs no copy of its own. */
export function markDocumentSaveBody(body: Record<string, unknown>): void {
  ownCopies.add(body);
}

/** Whether `content` is a save body's copy made here (its own, unshared). */
export function isDocumentSaveCopy(content: unknown): boolean {
  return record(content) && ownCopies.has(content);
}

/** Whether `content` was just proven to mirror exactly `projection` (the same
 * shared projection object, so the same document revision). The proof is
 * used once, by the save that follows the mirror check in the same call. */
export function documentCausalContentMirrored(content: unknown, projection: unknown): boolean {
  if (!record(content)) return false;
  const proven = mirroredBy.get(content) === projection;
  mirroredBy.delete(content);
  return proven;
}

/** A saved-source merge can keep a native layout receipt only when this
 * request's authored semantic body is unchanged. Native validation, not Yjs,
 * decides whether the carried receipt is valid for its source/fonts/layout. */
export function retainDocumentCausalSaveProjection(
  merged: Record<string, JsonValue>, authored: Record<string, JsonValue>,
  /** `owned`: `authored` is the caller's own copy and may be handed on as it is. */
  options: { owned?: boolean } = {},
): Record<string, JsonValue> {
  if (!documentContentSaysProjection(authored, merged)) return merged;
  return options.owned ? authored : structuredClone(authored);
}
