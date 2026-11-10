/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { Fragment, Node as ProseMirrorNode, Schema, type NodeSpec } from "@tiptap/pm/model";
import { reportClientError } from "../errors/report-beacon";
import { flattenPagedDocumentForStorage } from "../../components/console/editor/native-document/paged/serialization";
import { provisionalPageSizesForNodes } from "../../components/console/editor/native-document/paged/provisionalPages";
import { wholeBlockDistribution } from "../../components/console/editor/native-document/paged/pageBlockExtension";

/** Only physical pagination metadata is excluded. Native attrs and every
 * supported mark (including opaque font-face metadata) retain their schema. */
export { PAGINATION_RUNTIME_ATTRIBUTES } from "../../components/console/editor/native-document/paged/paginationProjectionKeys";
import { PAGINATION_RUNTIME_ATTRIBUTES } from "../../components/console/editor/native-document/paged/paginationProjectionKeys";

const schemas = new WeakMap<Schema, Schema>();

/** Whether a top-level room node is an inline node standing bare at block
 * level, which no document may hold (see withoutBareTopLevelInlineNodes). */
export function isBareTopLevelInlineNode(schema: Schema, typeName: string | undefined): boolean {
  const type = typeName ? schema.nodes[typeName] : undefined;
  return Boolean(type?.isInline);
}

/**
 * A room's document without the inline nodes stored bare at its top level.
 * The importer once wrote Word's cached page boundary before a chart that way
 * (`sourcePageBreak`, a `w:lastRenderedPageBreak` in the chart's own run;
 * R08-D12, R09-D13, R12-D16): an inline-only marker no sheet may hold, so the
 * whole document failed the schema and the page opened as one empty
 * paragraph. Such a node is cached layout, not content, and is dropped here;
 * the room binding skips the same elements (roomBinding), and the next write
 * from this document leaves them out of the room. The server heals stored
 * rooms the same way on load.
 */
export function withoutBareTopLevelInlineNodes<T>(schema: Schema, json: T): { json: T; dropped: number } {
  const root = json as { content?: Array<{ type?: string }> } | null;
  if (!root || !Array.isArray(root.content)) return { json, dropped: 0 };
  const kept = root.content.filter((node) => !isBareTopLevelInlineNode(schema, node?.type));
  const dropped = root.content.length - kept.length;
  if (!dropped) return { json, dropped: 0 };
  return { json: { ...root, content: kept.length ? kept : [{ type: "paragraph" }] } as T, dropped };
}

/**
 * Where a room's document first fails the schema: the path of the first
 * top-level block (and its first failing child) that cannot be read, for the
 * typed client error a page that cannot open reports.
 */
export function firstUnreadableNodePath(schema: Schema, json: unknown): string {
  const root = json as { content?: unknown[] } | null;
  if (!root || !Array.isArray(root.content)) return "doc";
  for (const [index, node] of root.content.entries()) {
    const typeName = (node as { type?: unknown } | null)?.type;
    if (typeof typeName !== "string" || !schema.nodes[typeName]) return `content[${index}](${String(typeName)}: unknown type)`;
    if (!schema.nodes[typeName].isBlock) return `content[${index}](${typeName}: not a block)`;
    try {
      schema.nodeFromJSON(node).check();
    } catch {
      const type = (node as { type?: unknown } | null)?.type;
      const children = (node as { content?: unknown[] } | null)?.content;
      if (Array.isArray(children)) {
        for (const [child, value] of children.entries()) {
          try { schema.nodeFromJSON(value).check(); } catch {
            return `content[${index}](${String(type)}).content[${child}](${String((value as { type?: unknown } | null)?.type)})`;
          }
        }
      }
      return `content[${index}](${String(type)})`;
    }
  }
  return "doc";
}

/** The native schema stays untouched. This private schema describes the same
 * semantic blocks without the local top-level page wrappers. */
export function canonicalDocumentSchema(localSchema: Schema): Schema {
  const cached = schemas.get(localSchema);
  if (cached) return cached;
  let nodes = localSchema.spec.nodes;
  nodes.forEach((name, original) => {
    const attrs = original.attrs && Object.fromEntries(
      Object.entries(original.attrs).filter(([key]) => !PAGINATION_RUNTIME_ATTRIBUTES.has(key)),
    );
    const spec: NodeSpec = {
      ...original,
      ...(attrs ? { attrs } : {}),
      ...(name === localSchema.topNodeType.name ? { content: "block+" } : {}),
    };
    nodes = nodes.update(name, spec);
  });
  const schema = new Schema({ ...localSchema.spec, nodes, marks: localSchema.spec.marks });
  schemas.set(localSchema, schema);
  return schema;
}









export function withoutUnsetPaginationRuntimeAttributes<T>(json: T): T {
  const strip = (value: unknown): unknown => {
    if (!value || typeof value !== "object" || Array.isArray(value)) return value;
    const node = value as Record<string, unknown>;
    let next: Record<string, unknown> = node;
    const attrs = node.attrs;
    if (attrs && typeof attrs === "object" && !Array.isArray(attrs)) {
      const entries = Object.entries(attrs as Record<string, unknown>);
      const kept = entries.filter(([key, item]) =>
        !(PAGINATION_RUNTIME_ATTRIBUTES.has(key) && (item === null || item === false || item === undefined)));
      if (kept.length !== entries.length) next = { ...node, attrs: Object.fromEntries(kept) };
    }
    if (Array.isArray(node.content)) {
      const children = node.content as unknown[];
      let changed = false;
      const stripped = children.map((child) => {
        const result = strip(child);
        if (result !== child) changed = true;
        return result;
      });
      if (changed) next = { ...next, content: stripped };
    }
    return next;
  };
  return strip(json) as T;
}

export function canonicalDocumentJson<T>(document: T): T {
  const flattened = flattenPagedDocumentForStorage(document, { paginationSettled: false });
  // The production serializer merges physical fragments, but some node
  // schemas still emit their default false/null layout flags afterward.
  // Remove only these known node-attribute keys before strict parsing;
  // opaque nested native metadata and all unknown author attrs stay intact.
  const strip = (value: unknown): unknown => {
    if (!value || typeof value !== "object" || Array.isArray(value)) return value;
    const node = value as Record<string, unknown>;
    const attrs = node.attrs && typeof node.attrs === "object" && !Array.isArray(node.attrs)
      ? Object.fromEntries(Object.entries(node.attrs).filter(([key]) => !PAGINATION_RUNTIME_ATTRIBUTES.has(key)))
      : node.attrs;
    return { ...node, ...(node.attrs !== undefined ? { attrs } : {}),
      ...(Array.isArray(node.content) ? { content: node.content.map(strip) } : {}) };
  };
  return strip(flattened) as T;
}









export const IDENTITY_FREE_TABLE_NODE_TYPES: ReadonlySet<string> = new Set(["tableRow", "tableCell", "tableHeader"]);

export function withoutUndeclaredTableIdentity<T>(schema: Schema, json: T): T {
  const undeclared = [...IDENTITY_FREE_TABLE_NODE_TYPES].filter((name) => {
    const type = schema.nodes[name];
    return type && !Object.prototype.hasOwnProperty.call(type.spec.attrs ?? {}, "block_id");
  });
  if (undeclared.length === 0) return json;
  const strip = (value: unknown): unknown => {
    if (!value || typeof value !== "object" || Array.isArray(value)) return value;
    const node = value as Record<string, unknown>;
    let next: Record<string, unknown> = node;
    const attrs = node.attrs;
    if (typeof node.type === "string" && undeclared.includes(node.type) && attrs && typeof attrs === "object" &&
        !Array.isArray(attrs) && Object.prototype.hasOwnProperty.call(attrs, "block_id")) {
      const { block_id: _identity, ...rest } = attrs as Record<string, unknown>;
      void _identity;
      next = { ...node, attrs: rest };
    }
    if (Array.isArray(node.content)) {
      const children = node.content as unknown[];
      let changed = false;
      const stripped = children.map((child) => {
        const result = strip(child);
        if (result !== child) changed = true;
        return result;
      });
      if (changed) next = { ...next, content: stripped };
    }
    return next;
  };
  return strip(json) as T;
}

/** What the tolerant read changed to open a body the strict read refused. */
export interface CanonicalReadRepair {
  /** Attributes (node "type.attr", mark "mark@attr") the schema does not declare. */
  undeclaredAttributes: string[];
  /** Node types the schema does not have, opened as their text. */
  unknownNodeTypes: string[];
  unknownMarkTypes: string[];
  emptyTextNodes: number;
  /** Containers given the empty first paragraph their content needs. */
  leadingParagraphs: number;
  /** Containers nothing repairs in place, opened as a paragraph of their words. */
  degradedNodes: string[];
  /** Blocks (an image, a table) found inside a paragraph or heading, lifted
   * out between the pieces of its text. */
  liftedBlocks: number;
}

/**
 * A body the strict read refuses, made one the editor holds by the same
 * rules the API's write gate repairs with (app/services/document_editor_schema.py;
 * both are held to fixtures/document-schema-verdicts.json): an attribute or
 * mark the schema does not declare is left out of the editing model, an
 * empty text node is dropped, a list item that opens with an equation gets
 * an empty first paragraph, and a node of a type this editor does not have
 * is opened as a paragraph of its text (an editable view, never a refusal).
 * Null when nothing here applies. Copy on write.
 */
/** Older writers' spellings of attributes the editor declares
 * (document_editor_schema._ATTRIBUTE_ALIASES on the API). */
const ATTRIBUTE_ALIASES: Readonly<Record<string, string>> = { attachmentId: "attachment_id" };

export function repairForCanonicalRead(schema: Schema, json: unknown): {
  json: unknown; repair: CanonicalReadRepair;
  /** Per top-level block of ``json``: how many blocks it was read as, and
   * whether it was read as it stood (null: the top level was not one for
   * one, an empty document given its paragraph). */
  blocks: CanonicalReadBlock[] | null;
} | null {
  const repair: CanonicalReadRepair = { undeclaredAttributes: [], unknownNodeTypes: [], unknownMarkTypes: [], emptyTextNodes: 0,
    leadingParagraphs: 0, degradedNodes: [], liftedBlocks: 0 };
  let blocks: CanonicalReadBlock[] | null = [];
  // Repairs other than leaving out an attribute that passes through (a newer
  // client's, kept in the room and in what is saved): those change what the
  // room's element must say, so its block is rewritten by the first write.
  let structural = 0;
  let changed = false;
  const note = (list: string[], value: string) => { if (!list.includes(value) && list.length < 40) list.push(value); changed = true; };
  const textOf = (value: unknown): string => {
    if (!value || typeof value !== "object") return "";
    const node = value as { text?: unknown; content?: unknown };
    return typeof node.text === "string" ? node.text : Array.isArray(node.content) ? node.content.map(textOf).join("") : "";
  };
  const fixMarks = (marks: unknown): unknown => {
    if (!Array.isArray(marks)) return marks;
    let touched = false;
    const kept = marks.flatMap((mark) => {
      const raw = mark as { type?: unknown; attrs?: unknown };
      const type = typeof raw?.type === "string" ? schema.marks[raw.type] : undefined;
      if (!type) { note(repair.unknownMarkTypes, String(raw?.type)); structural += 1; touched = true; return []; }
      if (raw.attrs && typeof raw.attrs === "object" && !Array.isArray(raw.attrs)) {
        const declared = type.spec.attrs ?? {};
        const entries = Object.entries(raw.attrs as Record<string, unknown>);
        const own = entries.filter(([key]) => Object.hasOwn(declared, key));
        if (own.length !== entries.length) {
          for (const [key] of entries) if (!Object.hasOwn(declared, key)) note(repair.undeclaredAttributes, `${type.name}@${key}`);
          structural += 1;
          touched = true;
          return [{ ...raw, attrs: Object.fromEntries(own) }];
        }
      }
      return [mark];
    });
    return touched ? kept : marks;
  };
  const fix = (value: unknown, top: boolean): unknown[] => {
    if (!value || typeof value !== "object" || Array.isArray(value)) return [value];
    const node = value as Record<string, unknown>;
    const type = typeof node.type === "string" ? schema.nodes[node.type] : undefined;
    if (!type) {
      note(repair.unknownNodeTypes, String(node.type));
      structural += 1;
      // As a paragraph of its words (an empty one keeps the room's element
      // count, which the room binding walks).
      const text = textOf(node);
      return [{ type: "paragraph", ...(text ? { content: [{ type: "text", text }] } : {}) }];
    }
    if (type.isText) {
      if (typeof node.text === "string" && node.text === "") { repair.emptyTextNodes += 1; structural += 1; changed = true; return []; }
      const marks = fixMarks(node.marks);
      return [marks === node.marks ? node : { ...node, marks }];
    }
    let next: Record<string, unknown> = node;
    const attrs = node.attrs;
    if (attrs && typeof attrs === "object" && !Array.isArray(attrs)) {
      const declared = type.spec.attrs ?? {};
      const entries = Object.entries(attrs as Record<string, unknown>);
      const own = entries.filter(([key]) => Object.hasOwn(declared, key) || PAGINATION_RUNTIME_ATTRIBUTES.has(key));
      if (own.length !== entries.length) {
        for (const [key] of entries) {
          if (!Object.hasOwn(declared, key) && !PAGINATION_RUNTIME_ATTRIBUTES.has(key)) note(repair.undeclaredAttributes, `${type.name}.${key}`);
        }
        const kept = Object.fromEntries(own);
        // An older writer's spelling of an attribute this editor declares
        // is that attribute (the API's write gate reads it the same way).
        for (const [alias, name] of Object.entries(ATTRIBUTE_ALIASES)) {
          const raw = attrs as Record<string, unknown>;
          if (Object.hasOwn(raw, alias) && !Object.hasOwn(declared, alias) && Object.hasOwn(declared, name)) {
            structural += 1;
            if (kept[name] === undefined || kept[name] === null || kept[name] === "") kept[name] = raw[alias];
          }
        }
        next = { ...next, attrs: kept };
      }
    }
    const marks = fixMarks(node.marks);
    if (marks !== node.marks) next = { ...next, marks };
    if (Array.isArray(node.content)) {
      let touched = false;
      const children = (node.content as unknown[]).flatMap((child) => {
        const before = structural;
        const fixed = fix(child, type === schema.topNodeType);
        const same = fixed.length === 1 && fixed[0] === child;
        if (!same) touched = true;
        if (type === schema.topNodeType) blocks?.push({ span: fixed.length, same: same || (fixed.length === 1 && structural === before) });
        return fixed;
      });
      let content = children;
      if (type.isTextblock && content.some((child) => isBlockJson(child))) {
        // A block written inside a paragraph (block_in_textblock, the same
        // repair the API's write gate makes): the paragraph is cut around
        // it, the block stands between the pieces, nothing is lost.
        changed = true;
        structural += 1;
        const pieces: unknown[] = [];
        let run: unknown[] = [];
        let identityGiven = false;
        const close = () => {
          if (!run.length) return;
          const pieceAttrs = next.attrs && typeof next.attrs === "object" && identityGiven && "block_id" in (next.attrs as object)
            ? Object.fromEntries(Object.entries(next.attrs as Record<string, unknown>).filter(([key]) => key !== "block_id"))
            : next.attrs;
          identityGiven = true;
          pieces.push({ ...next, ...(pieceAttrs !== undefined ? { attrs: pieceAttrs } : {}), content: run });
          run = [];
        };
        for (const child of content) {
          if (isBlockJson(child)) {
            close();
            const { marks: _marks, ...block } = child as Record<string, unknown>;
            void _marks;
            pieces.push(block);
            repair.liftedBlocks += 1;
          } else run.push(child);
        }
        close();
        return pieces;
      }
      try {
        const names = (list: unknown[]) => list.map((child) => schema.nodes[(child as { type: string }).type]);
        const matches = (list: unknown[]) => type.contentMatch.matchFragment(
          Fragment.fromArray(names(list).map((childType) => childType.isText ? schema.text("x") : childType.createAndFill() ?? childType.create())))?.validEnd;
        if (!matches(content) && schema.nodes.paragraph && matches([{ type: "paragraph" }, ...content])) {
          content = [{ type: "paragraph" }, ...content];
          repair.leadingParagraphs += 1;
          structural += 1;
          touched = true;
          if (type === schema.topNodeType) blocks = null;
        } else if (!matches(content) && type !== schema.topNodeType && !type.isInline && schema.nodes.paragraph) {
          // Nothing repairs this container in place (a row holding a
          // paragraph, a table without rows): it opens as a paragraph of its
          // words, editable, rather than the page not opening at all.
          note(repair.degradedNodes, type.name);
          structural += 1;
          const text = textOf(node);
          return [{ type: "paragraph", ...(text ? { content: [{ type: "text", text }] } : {}) }];
        }
      } catch {
        // The strict read names what is still wrong.
      }
      if (touched) { next = { ...next, content }; changed = true; }
    }
    return [next];
  };
  function isBlockJson(child: unknown): boolean {
    const childType = child && typeof child === "object" ? schema.nodes[(child as { type?: string }).type ?? ""] : undefined;
    return Boolean(childType && !childType.isInline);
  }
  const [result] = fix(json, false);
  return changed ? { json: result, repair, blocks } : null;
}

/** Whether an attribute ``type`` does not declare passes through a room this
 * client edits (a newer client's, kept for it), rather than being a known
 * leftover the read and the API's write gate both leave out: a layout flag,
 * a table row's or cell's block_id, an older writer's spelling. */
export function passesThroughCanonicalRead(type: string, key: string): boolean {
  return !PAGINATION_RUNTIME_ATTRIBUTES.has(key) && !Object.hasOwn(ATTRIBUTE_ALIASES, key) &&
    !(key === "block_id" && IDENTITY_FREE_TABLE_NODE_TYPES.has(type));
}

/** One top-level block of a body the tolerant read repaired. */
export interface CanonicalReadBlock {
  /** How many blocks of the read document it became (0: none, 2+: a block
   * lifted out of it). */
  span: number;
  /** Read as the room holds it (attributes that pass through aside): its
   * element binds to the block, nothing of it is rewritten. */
  same: boolean;
}

/** ProseMirror accepts unknown attributes by dropping them. At a shared
 * document boundary that would let an older client erase newer metadata on
 * its next edit. Verify every supplied attribute before creating a binding.
 *
 * A body that fails is not refused whole: it is read again repaired
 * (repairForCanonicalRead), and the repair is reported, so a page with a
 * stray attribute or an unknown node still opens and edits. */
export function parseCanonicalDocumentJson(schema: Schema, json: unknown): ProseMirrorNode {
  const stated = json;
  json = withoutUndeclaredTableIdentity(schema, json);
  // Top-level blocks the identity strip changed (one for one: copy on write).
  const statedBlocks = topLevelBlocks(stated), strippedBlocks = topLevelBlocks(json);
  const stripped = stated === json ? null : strippedBlocks?.map((block, index) => block !== statedBlocks?.[index]) ?? null;
  try {
    const parsed = parseCanonicalDocumentJsonStrict(schema, json);
    if (stated !== json) {
      readBlocks.set(parsed, stripped && strippedBlocks ? strippedBlocks.map((_, index) => ({ span: 1, same: !stripped[index] })) : null);
    }
    return parsed;
  } catch (error) {
    const repaired = repairForCanonicalRead(schema, json);
    if (!repaired) throw error;
    const parsed = parseCanonicalDocumentJsonStrict(schema, repaired.json);
    readRepairs.set(parsed, repaired.repair);
    readBlocks.set(parsed, repaired.blocks && (stated === json || (stripped && stripped.length === repaired.blocks.length))
      ? repaired.blocks.map((block, index) => ({ span: block.span, same: block.same && !stripped?.[index] })) : null);
    reportCanonicalReadRepair(repaired.repair, error);
    return parsed;
  }
}

function topLevelBlocks(json: unknown): unknown[] | null {
  const content = json && typeof json === "object" ? (json as { content?: unknown }).content : undefined;
  return Array.isArray(content) ? content : null;
}

const readRepairs = new WeakMap<ProseMirrorNode, CanonicalReadRepair>();
const readBlocks = new WeakMap<ProseMirrorNode, CanonicalReadBlock[] | null>();

/** For a document the read did not take exactly as stated (repaired, or a
 * table's undeclared identity left out), which top-level blocks it took as
 * they stand. ``undefined``: read exactly as stated; ``null``: the top level
 * is not one for one. A room's binding binds only the blocks taken as they
 * stand, so this client's first write puts the read form of the others in
 * the room (then room and editor say the same, and the page saves). */
export function canonicalReadBlocksOf(doc: ProseMirrorNode): CanonicalReadBlock[] | null | undefined {
  return readBlocks.has(doc) ? readBlocks.get(doc) ?? null : undefined;
}

/** What the tolerant read changed to open ``doc`` (null: it read as it stood). */
export function canonicalReadRepairOf(doc: ProseMirrorNode): CanonicalReadRepair | null {
  return readRepairs.get(doc) ?? null;
}

function reportCanonicalReadRepair(repair: CanonicalReadRepair, cause: unknown): void {
  try {
    // The message leads with the repairs' codes, comma separated, as the
    // Presentation/Spreadsheet read degradations do (alerts group by name).
    const codes = [
      repair.undeclaredAttributes.length && "undeclared_attribute",
      repair.unknownNodeTypes.length && "unknown_node_type",
      repair.unknownMarkTypes.length && "unknown_mark_type",
      repair.emptyTextNodes && "empty_text",
      repair.leadingParagraphs && "missing_leading_paragraph",
      repair.degradedNodes.length && "degraded_container",
      repair.liftedBlocks && "block_in_textblock",
    ].filter(Boolean).join(",");
    reportClientError(Object.assign(new Error(`${codes || "repaired"} ${JSON.stringify(repair).slice(0, 400)} (strict: ${
      cause instanceof Error ? cause.message.slice(0, 120) : "refused"})`), { name: "DocumentContentReadDegraded" }),
    { source: "window", windowType: "document" });
  } catch {
    // diagnostics never stop the read
  }
}

/** The strict read alone (tests compare it with the API's checker). */
export function parseCanonicalDocumentJsonStrict(schema: Schema, json: unknown): ProseMirrorNode {
  const parsed = schema.nodeFromJSON(json);
  parsed.check();
  const equal = (left: unknown, right: unknown) => Object.is(left, right) || JSON.stringify(left) === JSON.stringify(right);
  const verifyAttrs = (provided: unknown, actual: Readonly<Record<string, unknown>>) => {
    if (provided === undefined) return;
    if (provided === null || typeof provided !== "object" || Array.isArray(provided)) throw new Error("Invalid canonical attributes");
    for (const [key, value] of Object.entries(provided)) {
      if (!Object.prototype.hasOwnProperty.call(actual, key) || !equal(value, actual[key])) throw new Error(`Unsupported canonical attribute: ${key}`);
    }
  };
  const verify = (raw: Record<string, unknown>, node: ProseMirrorNode) => {
    verifyAttrs(raw.attrs, node.attrs);
    if (Array.isArray(raw.marks)) {
      for (const mark of raw.marks as Array<{ type: string; attrs?: Record<string, unknown> }>) {
        const compatible = node.marks.find(candidate => candidate.type.name === mark.type &&
          Object.entries(mark.attrs ?? {}).every(([key, value]) => Object.prototype.hasOwnProperty.call(candidate.attrs, key) && equal(value, candidate.attrs[key])));
        if (!compatible) throw new Error("Unsupported canonical mark metadata");
      }
    }
    const children = Array.isArray(raw.content) ? raw.content as Array<Record<string, unknown>> : [];
    // Adjacent equivalent text nodes can be coalesced by Fragment.fromArray.
    // Parse/verify each supplied child independently so that metadata is not
    // lost merely because the normalized child count differs. With an equal
    // count nothing was coalesced: the parsed children ARE those parses.
    if (children.length === node.childCount) children.forEach((child, index) => verify(child, node.child(index)));
    else for (const child of children) verify(child, schema.nodeFromJSON(child));
  };
  verify(json as Record<string, unknown>, parsed);
  return parsed;
}

export function projectPagedDocument(document: ProseMirrorNode): ProseMirrorNode {
  const schema = canonicalDocumentSchema(document.type.schema);
  const projected = parseCanonicalDocumentJson(schema, canonicalDocumentJson(document.toJSON()));
  projected.check();
  return projected;
}

/** A safe initial/local fallback projection. Actual page measurement and
 * continuation splitting remain owned by the unchanged native paginator. */
export function wrapCanonicalDocument(
  canonical: ProseMirrorNode,
  localSchema: Schema,
  localAttrs: Readonly<Record<string, unknown>> = canonical.attrs,
  /** A whole-block page distribution to open on (the pages this browser last
   * laid the document out on), used when it fits the document's blocks. */
  knownSizes?: readonly number[] | null,
): ProseMirrorNode {
  const page = localSchema.nodes.pageBlock;
  if (!page) throw new Error("A local paginated schema is required");
  const blocks: ProseMirrorNode[] = [];
  canonical.forEach(node => blocks.push(localSchema.nodeFromJSON(node.toJSON())));
  // A long document opens as provisional sheets of about a page each, which
  // the page ledger lays out progressively (provisionalPages); a short one as
  // one sheet, as before.
  const sizes = wholeBlockDistribution(knownSizes, blocks.length) ?? provisionalPageSizesForNodes(blocks, localAttrs);
  const pages: ProseMirrorNode[] = [];
  if (sizes) {
    let cursor = 0;
    for (const size of sizes) {
      pages.push(page.create(null, blocks.slice(cursor, cursor + size)));
      cursor += size;
    }
  } else {
    pages.push(page.create(null, blocks));
  }
  const document = localSchema.topNodeType.create(localAttrs, pages);
  document.check();
  return document;
}
