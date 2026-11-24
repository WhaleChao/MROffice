/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { Extension, type Editor, type JSONContent } from "@tiptap/core";
import type { Node as ProseMirrorNode, Slice } from "@tiptap/pm/model";
import { ReplaceStep } from "@tiptap/pm/transform";
import { Plugin, type Selection, type Transaction } from "@tiptap/pm/state";

import { liftHeadingsOutOfLists } from "./liftHeadingsOutOfLists";
import {
  imeSafeAppendTransaction,
} from "./imeCompositionGuard";
import { editsStayInsideTextblocks } from "./textblockOnlyEdits";
import { SPLIT_CONTINUATION_ATTR } from "./paged/paginationProjectionKeys";

export const MRNOTE_BLOCK_ID_ATTR = "block_id";
export const MRNOTE_SOURCE_ATTR = "source";
/**
 * Meta of the appended transaction that back-fills or re-mints identity.
 * Pagination copies a block's attrs onto every physical fragment it cuts,
 * so each reflow that splits a block is followed by one of these; the
 * serializer merges the fragments back, so the transaction never changes
 * the canonical document. Consumers that track derived pagination (the
 * export authority) recognize it by this meta.
 */
export const MRNOTE_BLOCK_IDENTITY_META = "mrnoteBlockIdentity";

export const MRNOTE_BLOCK_ID_NODE_TYPES = [
  "paragraph",
  "heading",
  "bulletList",
  "orderedList",
  "listItem",
  "taskList",
  "taskItem",
  "codeBlock",
  "blockquote",
  "table",
  "horizontalRule",
  "pageBreak",
  "image",
  "documentShape",
  "documentMedia",
  "mathBlock",
  "callout",
  "whiteboard",
  "file",
  "chart",
  "smartArt",
  "mermaid",
  "codeChart",
  "ai_output",
  "reference",
  "task",
  "flashcard",
  "textBox",
  // Every remaining top-level block type. The shared API read/write boundary
  // stamps `block_id` on every top-level node, and the collab canonical
  // boundary refuses attributes the schema does not declare, so the client
  // declaration must cover the whole document flow (a sectionBreak or
  // footnoteList without this declaration made the page unreadable).
  "imageRow",
  "documentSignatureLine",
  "documentAltChunkReadOnly",
  "documentVirtualTable",
  "columnBreak",
  "sectionBreak",
  "tableOfContents",
  "documentCoverPage",
  "frontmatter",
  "footnoteList",
  "bibliography",
  "caption",
  "listOfFigures",
  "blockContentControl",
];

/**
 * Node types whose own `source` attr is semantic content or node-defined
 * provenance (the mermaid DSL, the Python script, the signature line, the
 * alternate-content chunk, the bibliography field). A global provenance
 * attribute with the same name would clobber it, so they are excluded.
 */
const OWN_SOURCE_ATTR_NODE_TYPES = new Set([
  "mermaid",
  "codeChart",
  "documentSignatureLine",
  "documentAltChunkReadOnly",
  "bibliography",
]);

/**
 * Node types that additionally carry the technical provenance `source` attr.
 */
export const MRNOTE_SOURCE_NODE_TYPES = MRNOTE_BLOCK_ID_NODE_TYPES.filter(
  (type) => !OWN_SOURCE_ATTR_NODE_TYPES.has(type),
);

/**
 * Word "In Line with Text" pictures and charts among a paragraph's words.
 * They carry the same provenance receipt as the block picture/chart they
 * stand for (the save writer finds the package drawing by its
 * `sourceAnchor`), but they are characters of their line, never blocks, so
 * they get no block identity.
 */
export const MRNOTE_INLINE_SOURCE_NODE_TYPES = ["inlineImage", "inlineChart"];

export interface EnsureBlockIdsResult {
  document: JSONContent;
  changed: boolean;
}

export function generateNativeBlockId(): string {
  if (typeof crypto !== "undefined" && "randomUUID" in crypto) {
    return crypto.randomUUID();
  }
  return `block-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function cloneAttrs(attrs: Record<string, unknown> | undefined): Record<string, unknown> {
  return attrs ? { ...attrs } : {};
}










function splitTailKeepsSource(attrs: Record<string, unknown>): boolean {
  if (attrs[SPLIT_CONTINUATION_ATTR] !== true) return false;
  return typeof window === "undefined" || (window as Window & { __docsSplitTailKeepsSource?: boolean }).__docsSplitTailKeepsSource !== false;
}

function canonicalDocxParagraphSourceAnchor(attrs: Record<string, unknown>): string | null {
  const source = attrs[MRNOTE_SOURCE_ATTR];
  if (!source || typeof source !== "object" || Array.isArray(source)) {
    return null;
  }
  const record = source as Record<string, unknown>;
  if (record.adapter !== "docx" || record.nodeType !== "paragraph") {
    return null;
  }
  const rawAnchor = record.sourceAnchor ?? record.source_anchor;
  if (typeof rawAnchor !== "string") {
    return null;
  }
  const anchor = rawAnchor.trim();
  return /^word\/document\.xml#p\[\d+\](?:\/seg\[\d+\])?$/.test(anchor) ? anchor : null;
}

function shouldAssignBlockId(node: JSONContent): boolean {
  return Boolean(node.type && MRNOTE_BLOCK_ID_NODE_TYPES.includes(node.type));
}

function normalizeNodeBlockIds(
  node: JSONContent,
  createBlockId: () => string,
  seenDocxParagraphAnchors: Set<string>,
): { node: JSONContent; changed: boolean } {
  let changed = false;
  let content = node.content;
  let attrs = node.attrs as Record<string, unknown> | undefined;

  if (shouldAssignBlockId(node)) {
    // The attributes are copied only when something is written into them:
    // copying every block's on every save cost an inked page's save, under
    // the pen, and an unchanged node is handed back as it is anyway.
    let clonedAttrs: Record<string, unknown> | null = null;
    const writable = () => (clonedAttrs ??= cloneAttrs(attrs));
    const current = attrs ?? {};
    if (typeof current[MRNOTE_BLOCK_ID_ATTR] !== "string" || !current[MRNOTE_BLOCK_ID_ATTR]) {
      writable()[MRNOTE_BLOCK_ID_ATTR] = createBlockId();
      attrs = writable();
      changed = true;
    }
    if (node.type === "paragraph" || node.type === "heading") {
      const sourceAnchor = canonicalDocxParagraphSourceAnchor(clonedAttrs ?? current);
      if (sourceAnchor && splitTailKeepsSource(clonedAttrs ?? current)) {
        // The head's own provenance (see splitTailKeepsSource).
      } else if (sourceAnchor && seenDocxParagraphAnchors.has(sourceAnchor)) {
        // Provenance identifies one native w:p. A split sibling is a new
        // paragraph, not a second claimant for the same source paragraph.
        writable()[MRNOTE_SOURCE_ATTR] = null;
        attrs = writable();
        changed = true;
      } else if (sourceAnchor) {
        seenDocxParagraphAnchors.add(sourceAnchor);
      }
    }
  }

  if (Array.isArray(node.content)) {
    let contentChanged = false;
    const normalizedContent = node.content.map((child) => {
      const result = normalizeNodeBlockIds(child, createBlockId, seenDocxParagraphAnchors);
      contentChanged = contentChanged || result.changed;
      return result.node;
    });
    if (contentChanged) {
      content = normalizedContent;
      changed = true;
    }
  }

  if (!changed) {
    return { node, changed: false };
  }

  const nextNode = { ...node };
  if (attrs !== node.attrs) {
    nextNode.attrs = attrs;
  }
  if (content !== node.content) {
    nextNode.content = content;
  }
  return { node: nextNode, changed: true };
}

export function ensureNativeDocumentBlockIds(
  document: JSONContent,
  createBlockId: () => string = generateNativeBlockId,
): EnsureBlockIdsResult {
  // DEFENSIVE (belt & braces): legacy DOCX imports list-grouped Word numbered
  // outline headings, persisting `orderedList > listItem > heading` — content
  // the editor schema forbids, so the whole page failed to load. Lift such
  // headings to top level BEFORE the block-id backfill (split segments minted
  // without a block_id get fresh ids right here). New imports no longer
  // produce the shape; valid documents pass through untouched.
  const lifted = liftHeadingsOutOfLists(document);
  document = lifted.document;
  if (!Array.isArray(document.content)) {
    return { document, changed: lifted.changed };
  }

  let changed = lifted.changed;
  const seenDocxParagraphAnchors = new Set<string>();
  const content = document.content.map((node) => {
    const result = normalizeNodeBlockIds(node, createBlockId, seenDocxParagraphAnchors);
    changed = changed || result.changed;
    return result.node;
  });

  if (!changed) {
    return { document, changed: false };
  }
  return {
    document: {
      ...document,
      content,
    },
    changed: true,
  };
}

export type NativeTextBlockTypeName = "paragraph" | "heading";

function textBlocksInSelection(
  doc: ProseMirrorNode,
  selection: Selection,
): Array<{ pos: number; node: ProseMirrorNode }> {
  const positions = new Map<number, ProseMirrorNode>();
  doc.nodesBetween(selection.from, selection.to, (node, pos) => {
    if (node.type.name === "paragraph" || node.type.name === "heading") {
      positions.set(pos, node);
    }
  });
  if (selection.empty) {
    for (let depth = selection.$from.depth; depth > 0; depth -= 1) {
      const node = selection.$from.node(depth);
      if (node.type.name === "paragraph" || node.type.name === "heading") {
        positions.set(selection.$from.before(depth), node);
        break;
      }
    }
  }
  return [...positions.entries()]
    .map(([pos, node]) => ({ pos, node }))
    .sort((left, right) => left.pos - right.pos);
}

function retypedTextBlockAttrs(
  node: ProseMirrorNode,
  typeName: NativeTextBlockTypeName,
  attrs: Record<string, unknown>,
): Record<string, unknown> {
  const nextAttrs: Record<string, unknown> = {
    ...(node.attrs || {}),
    ...attrs,
  };
  if (typeName === "paragraph") {
    delete nextAttrs.level;
  }
  if (!nextAttrs[MRNOTE_BLOCK_ID_ATTR]) {
    delete nextAttrs[MRNOTE_BLOCK_ID_ATTR];
  }
  if (!nextAttrs[MRNOTE_SOURCE_ATTR]) {
    delete nextAttrs[MRNOTE_SOURCE_ATTR];
  }
  return nextAttrs;
}

export function setSelectedTextBlockTypePreservingAttrs(
  editor: Editor,
  typeName: NativeTextBlockTypeName,
  attrs: Record<string, unknown> = {},
): boolean {
  const nodeType = editor.state.schema.nodes[typeName];
  if (!nodeType) return false;
  const targets = textBlocksInSelection(editor.state.doc, editor.state.selection);
  if (targets.length === 0) return false;
  let transaction = editor.state.tr;
  for (const target of targets) {
    transaction = transaction.setNodeMarkup(
      target.pos,
      nodeType,
      retypedTextBlockAttrs(target.node, typeName, attrs),
    );
  }
  editor.view.dispatch(transaction.scrollIntoView());
  editor.commands.focus();
  return true;
}

/**
 * The same retyping inside a chain's own transaction (its selection, its
 * document), so it shares one undo step with the chain's other commands. A
 * block already of that type is left as it is.
 */
export function retypeSelectedTextBlocksInTransaction(
  tr: Transaction,
  typeName: NativeTextBlockTypeName,
): boolean {
  const nodeType = tr.doc.type.schema.nodes[typeName];
  if (!nodeType) return false;
  for (const target of textBlocksInSelection(tr.doc, tr.selection)) {
    if (target.node.type === nodeType) continue;
    tr.setNodeMarkup(
      target.pos,
      nodeType,
      retypedTextBlockAttrs(target.node, typeName, {}),
    );
  }
  return true;
}

/** The pen's landing (documentInk, DOCUMENT_INK_LANDING_META): strokes and
 * the paragraph they may bring, inserted with ids minted for them, nothing
 * removed or retyped. Every id stays unique without a look at the rest of the
 * document (a walk of every block after each stroke of a long note). */
const INK_LANDING_META = "documentInkLanding";
/** Ids minted for blocks the pen is about to land (mintFreshBlockId), not yet
 * seen in a document: only these let a landing skip the walk. */
const freshBlockIds = new Set<string>();
const FRESH_BLOCK_IDS_MAX = 4096;
/** A new block id that is in no document yet, for a block about to land. */
export function mintFreshBlockId(): string {
  if (freshBlockIds.size >= FRESH_BLOCK_IDS_MAX) freshBlockIds.clear();
  const id = generateNativeBlockId();
  freshBlockIds.add(id);
  return id;
}
export function landsWithFreshIds(transactions: readonly Transaction[]): boolean {
  const ids = new Set<string>();
  for (const tr of transactions) {
    if (tr.getMeta(INK_LANDING_META) !== true) return false;
    for (const step of tr.steps) {
      if (!(step instanceof ReplaceStep)) return false;
      const { from, to, slice } = step as unknown as { from: number; to: number; slice: Slice };
      if (from !== to || slice.openStart !== 0 || slice.openEnd !== 0) return false;
      let fresh = true;
      slice.content.descendants((node) => {
        if (!fresh) return false;
        if (!node.isBlock || !(MRNOTE_BLOCK_ID_ATTR in node.attrs)) return true;
        const id = node.attrs[MRNOTE_BLOCK_ID_ATTR];
        if (typeof id !== "string" || !freshBlockIds.has(id) || ids.has(id)) fresh = false;
        else ids.add(id);
        return fresh;
      });
      if (!fresh) return false;
    }
  }
  if (ids.size === 0) return false;
  // In a document now: never fresh again.
  for (const id of ids) freshBlockIds.delete(id);
  return true;
}

/** Documents every block of which holds a unique id (see the plugin). */
const completeBlockIdDocuments = new WeakSet<ProseMirrorNode>();

/**
 * The id of a page split's tail (the same paragraph as its head, drawn on the
 * next sheet), derived from the head's id and the tail's own words instead of
 * drawn at random. The ledger cuts a paragraph again on most passes: drawn
 * at random, every cut gave the tail a new block, whose line box and witness
 * a new node does not have (`setNodeMarkup` here re-creates it) and whose
 * measurements no memo keyed by id knew; measured bare, the tail cut two
 * characters further, and the next pass cut it back (R461, page 47: the tail
 * re-minted 896b90d3 -> 7d3524ac -> fe8a1ebd..., 36-55 reversals in 150 s at
 * rest). The same cut now gives the same id, a cut that moved a new one.
 * Never stored: a split merges back into one block at every serialization.
 * `window.__docsSplitTailDerivedId = false`: drawn at random, as before.
 */
function splitTailBlockId(node: ProseMirrorNode, headId: unknown, seen: ReadonlySet<string>): string | null {
  if (node.attrs[SPLIT_CONTINUATION_ATTR] !== true || typeof headId !== "string" || !headId) return null;
  if (typeof window !== "undefined" && (window as Window & { __docsSplitTailDerivedId?: boolean }).__docsSplitTailDerivedId === false) {
    return null;
  }
  for (let attempt = 0; attempt < 4; attempt += 1) {
    const id = derivedBlockId(`${headId}\u0000${attempt}\u0000${node.textContent}`);
    if (!seen.has(id)) return id;
  }
  return null;
}

/**
 * The id a page split gives its tail piece in the very step that cuts it:
 * derived from the head's id and where the cut falls in the head, so the same
 * cut made again gives the same id, and the identities' back-fill finds it
 * unique and has nothing to re-mark (re-marked with setNodeMarkup, the piece
 * was rebuilt and drew without its line box until the next pass: R461-D71's
 * page 47 cut back and forth at rest). `window.__docsSplitTailIdAtCut =
 * false`: the tail keeps the head's id and the back-fill gives it one, as
 * before.
 */
export function splitTailIdAttrs(headAttrs: Record<string, unknown> | null | undefined, offset: number): Record<string, unknown> {
  const headId = headAttrs?.[MRNOTE_BLOCK_ID_ATTR];
  if (typeof headId !== "string" || !headId) return {};
  if (typeof window !== "undefined" && (window as Window & { __docsSplitTailIdAtCut?: boolean }).__docsSplitTailIdAtCut === false) return {};
  return { [MRNOTE_BLOCK_ID_ATTR]: derivedBlockId(`${headId}\u0000cut\u0000${offset}`) };
}

/** A UUID-shaped id from 128 bits of a string's hash (cyrb128). */
function derivedBlockId(text: string): string {
  let h1 = 1779033703, h2 = 3144134277, h3 = 1013904242, h4 = 2773480762;
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    h1 = h2 ^ Math.imul(h1 ^ code, 597399067);
    h2 = h3 ^ Math.imul(h2 ^ code, 2869860233);
    h3 = h4 ^ Math.imul(h3 ^ code, 951274213);
    h4 = h1 ^ Math.imul(h4 ^ code, 2716044179);
  }
  h1 = Math.imul(h3 ^ (h1 >>> 18), 597399067);
  h2 = Math.imul(h4 ^ (h2 >>> 22), 2869860233);
  h3 = Math.imul(h1 ^ (h3 >>> 17), 951274213);
  h4 = Math.imul(h2 ^ (h4 >>> 19), 2716044179);
  h1 ^= h2 ^ h3 ^ h4; h2 ^= h1; h3 ^= h1; h4 ^= h1;
  const hex = [h1, h2, h3, h4].map((part) => (part >>> 0).toString(16).padStart(8, "0")).join("");
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-8${hex.slice(13, 16)}-${((parseInt(hex[16], 16) & 3) | 8).toString(16)}${hex.slice(17, 20)}-${hex.slice(20, 32)}`;
}

function assignMissingBlockIds(tr: Transaction, doc: ProseMirrorNode): boolean {
  let changed = false;
  // block_id is the identity key AI edits and traces target, so it must be
  // unique. Splitting a paragraph (Enter mid-text, inserting a table/image
  // mid-paragraph) copies attrs to BOTH halves, and pasting a copied block
  // round-trips data-block-id through parseHTML — both minted duplicates
  // that persisted to the backend. Reassign every id already seen earlier in
  // document order (the first occurrence keeps the original identity).
  const seen = new Set<string>();
  const seenDocxParagraphAnchors = new Set<string>();
  doc.descendants((node, pos) => {
    if (!node.isBlock || !(MRNOTE_BLOCK_ID_ATTR in node.attrs)) {
      return true;
    }
    let nextAttrs = node.attrs;
    let nodeChanged = false;
    const blockId = nextAttrs[MRNOTE_BLOCK_ID_ATTR];
    if (typeof blockId === "string" && blockId && !seen.has(blockId)) {
      seen.add(blockId);
    } else {
      let freshId = splitTailBlockId(node, blockId, seen) ?? generateNativeBlockId();
      if (seen.has(freshId)) freshId = generateNativeBlockId();
      nextAttrs = {
        ...nextAttrs,
        [MRNOTE_BLOCK_ID_ATTR]: freshId,
      };
      seen.add(freshId);
      nodeChanged = true;
    }

    if (node.type.name === "paragraph" || node.type.name === "heading") {
      const sourceAnchor = canonicalDocxParagraphSourceAnchor(nextAttrs);
      if (sourceAnchor && splitTailKeepsSource(nextAttrs)) {
        // The head's own provenance (see splitTailKeepsSource).
      } else if (sourceAnchor && seenDocxParagraphAnchors.has(sourceAnchor)) {
        nextAttrs = { ...nextAttrs, [MRNOTE_SOURCE_ATTR]: null };
        nodeChanged = true;
      } else if (sourceAnchor) {
        seenDocxParagraphAnchors.add(sourceAnchor);
      }
    }

    if (nodeChanged) {
      tr.setNodeMarkup(pos, undefined, nextAttrs);
      changed = true;
    }
    return true;
  });
  if (changed) {
    tr.setMeta("addToHistory", false);
  }
  return changed;
}

export const MrnoteBlockIdExtension = Extension.create({
  name: "mrnoteBlockIds",

  addGlobalAttributes() {
    return [
      {
        types: MRNOTE_BLOCK_ID_NODE_TYPES,
        attributes: {
          [MRNOTE_BLOCK_ID_ATTR]: {
            default: null,
            parseHTML: (element) => element.getAttribute("data-block-id"),
            renderHTML: (attributes) => {
              const blockId = attributes[MRNOTE_BLOCK_ID_ATTR];
              return typeof blockId === "string" && blockId
                ? { "data-block-id": blockId }
                : {};
            },
          },
        },
      },
      {
        types: MRNOTE_SOURCE_NODE_TYPES,
        attributes: {
          [MRNOTE_SOURCE_ATTR]: {
            default: null,
            keepOnSplit: false,
            parseHTML: () => null,
            renderHTML: (attributes) => {
              const source = attributes[MRNOTE_SOURCE_ATTR];
              // Inline Word tables are separated by their real paragraphs.
              // The generic Markdown table margin is not authored spacing.
              // Keep provenance model-only; project only this layout property.
              if (!source || typeof source !== "object" || Array.isArray(source)
                || source.adapter !== "docx" || source.nodeType !== "table") {
                return {};
              }
              // A Word table sizes its text like the document around it.
              return attributes.tablePositioning == null
                ? { "data-native-table": "true", "data-native-inline-table": "true", style: "margin-block: 0" }
                : { "data-native-table": "true" };
            },
          },
        },
      },
      {
        // Model-only provenance, like the block receipt above. Without it the
        // canonical collab parser rejects an imported page that holds an
        // in-line picture (an attribute the node type does not declare), and
        // a picture moved onto a text line would lose the anchor its save
        // needs.
        types: MRNOTE_INLINE_SOURCE_NODE_TYPES,
        attributes: {
          [MRNOTE_SOURCE_ATTR]: {
            default: null,
            keepOnSplit: false,
            parseHTML: () => null,
            renderHTML: () => ({}),
          },
        },
      },
    ];
  },

  onCreate({ editor }) {
    const tr = editor.state.tr;
    if (assignMissingBlockIds(tr, editor.state.doc)) {
      tr.setMeta(MRNOTE_BLOCK_IDENTITY_META, true);
      if (editor.state.storedMarks) tr.setStoredMarks(editor.state.storedMarks);
      editor.view.dispatch(tr);
    }
  },

  addProseMirrorPlugins() {
    return [
      new Plugin({
        // Never redraw a composing textblock by back-filling attrs. The
        // guard's settled meta transaction runs this scan after commit.
        appendTransaction: imeSafeAppendTransaction(
          (transactions, oldState, newState) => {
            // Every id is present and unique in a document found complete,
            // and an edit that stays inside text blocks keeps it so (no block
            // added, split or given attributes): a caret move or a keystroke
            // walked every block of the document to find that again.
            if (completeBlockIdDocuments.has(newState.doc)) return null;
            if (completeBlockIdDocuments.has(oldState.doc) &&
                (editsStayInsideTextblocks(transactions) || landsWithFreshIds(transactions))) {
              completeBlockIdDocuments.add(newState.doc);
              return null;
            }
            const tr = newState.tr;
            if (!assignMissingBlockIds(tr, newState.doc)) {
              completeBlockIdDocuments.add(newState.doc);
              return null;
            }
            tr.setMeta(MRNOTE_BLOCK_IDENTITY_META, true);
            // Back-filling identity is not a formatting change. Retain the
            // chosen typing style, including on a newly styled empty heading.
            if (newState.storedMarks) tr.setStoredMarks(newState.storedMarks);
            return tr;
          },
        ),
      }),
    ];
  },
});
