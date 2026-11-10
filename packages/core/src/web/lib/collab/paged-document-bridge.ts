/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { msSinceTyping, TYPING_IDLE_MS } from "../typing-session";
import { inkWritingActive } from "../ink-writing-session";
import { isDocumentSaveMaintenance } from "./document-save-maintenance";
import { PAGED_AWAITS_FOLLOW_UP_META, PAGED_FOLLOW_UP_META, PAGED_HISTORY_APPLIED_META } from "./paged-history-meta";
import { restoreStepTextByIdentity } from "./paged-history-text-fallback";
import { TYPING_RUN_GAP_MS } from "../editor/typing-run-time";
import { documentCausalSemanticContent, readDocumentCausalProjection } from "./document-page-causal-model";
import { bindUnchangedInSlices, noteForeignCanonicalXmlAttributes, updateCanonicalYFragment } from "./canonical-xml-attributes";
import { holdsEveryBlock, mergeCanonicalBlocks, type MergeBlock } from "./canonical-block-merge";
import { MappablePosition, getUpdatedPosition as mapCorePosition, type Editor, type GetUpdatedPositionResult } from "@tiptap/core";
import { Fragment, Slice, type Node as ProseMirrorNode, type Schema } from "@tiptap/pm/model";
import { AllSelection, NodeSelection, Plugin, Selection, TextSelection, type EditorState, type Transaction } from "@tiptap/pm/state";
import { CellSelection } from "@tiptap/pm/tables";
import { AttrStep, DocAttrStep, Mapping, ReplaceStep, type Step } from "@tiptap/pm/transform";
import type { EditorView } from "@tiptap/pm/view";
import { initProseMirrorDoc, relativePositionToAbsolutePosition, yCursorPluginKey, ySyncPluginKey, yUndoPluginKey, yXmlFragmentToProsemirrorJSON } from "@tiptap/y-tiptap";
import { COLLABORATIVE_DOCUMENT_ATTRS_MAP, COLLABORATIVE_DOCUMENT_ATTRS_REMOTE_META, publishCollaborativeDocumentAttrs, readCollaborativeDocumentAttrs } from "../../components/console/editor/native-document/collaborationDocumentAttrs";
import { INLINE_ATOM_RUN_STYLE_META } from "../../components/console/editor/native-document/inlineAtomRunStyle";
import * as Y from "yjs";
import { WHITEBOARD_SCENE_META, serializeWhiteboardScenes } from "../../components/console/editor/extensions/whiteboardScene";
import { pageDocumentJsonEquals } from "../page-save-acknowledgement";
import type { CollabSession } from "./collab-provider";
import { isLocalUsersOtherConnection, sanitizeCollabUser } from "./collab-state";
import { createCaretChangeFilter } from "./ink-presence";
import { MAX_DRAWN_CARETS, PeerCaretOverlay, type PeerCaret } from "./peer-caret-overlay";
import { canonicalDocumentJson, canonicalDocumentSchema, canonicalReadBlocksOf, canonicalReadRepairOf, firstUnreadableNodePath, passesThroughCanonicalRead, isBareTopLevelInlineNode, parseCanonicalDocumentJson, projectPagedDocument, withoutBareTopLevelInlineNodes, withoutUnsetPaginationRuntimeAttributes, wrapCanonicalDocument } from "./paged-document-projection";
import { reportClientError } from "../errors/report-beacon";
import { PROVISIONAL_PAGES_MIN_BLOCKS } from "../../components/console/editor/native-document/paged/provisionalPages";
import { markRestoredLayout } from "../../components/console/editor/native-document/paged/restoredLayoutMark";
import { forgetOpenedProjectionContent, paginationProjectionRefusal, rememberOpenedProjectionContent, restorePaginationProjection } from "../../components/console/editor/native-document/paged/paginationProjectionReceipt";
import { traceDocsLayoutReceipt, tracingDocsLayoutReceipt } from "../docs-layout-receipt-cache";
import type { JSONContent } from "@tiptap/core";
import { roomRelativePosition, VersionedMapping } from "./room-relative-position";
import { createPagedDocumentPositionMap, derivePagedDocumentPositionMap, type PagedDocumentPositionMap } from "./paged-document-position-map";
import { SAVED_WITHOUT_HISTORY_META, isTrustedReadOnlyDocumentMutation } from "../../components/console/editor/native-document/readOnlyTransactionAuthority";
import { isImeCompositionActive, isImeDeferredUserEdit, onImeCompositionSettled } from "../../components/console/editor/native-document/imeCompositionGuard";
import { msSinceLayoutStatusChange, PAGE_LEDGER_BACKGROUND_ATTR, PAGE_LEDGER_DEFERRED_ATTR } from "../../components/console/editor/native-document/documentLayoutStatus";
import { PAGED_BULK_FILL_ATTR, writeDocumentLayoutStatus } from "../../components/console/editor/native-document/documentLayoutStatus";
import { BULK_FILL_CHUNK_META } from "../../components/console/editor/native-document/bulkFillTransaction";
import { hollowSheetsEnabled, pageHollowSheetsKey, viewMayDrawHollow } from "../../components/console/editor/native-document/paged/pageHollowSheets";
import { carriesInlineAtomRunStyle } from "../../components/console/editor/native-document/inlineAtomText";
import { createBackgroundLane, type BackgroundLane } from "../background-work";

export const PAGED_CANONICAL_REMOTE_META = "mrnotePagedCanonicalRemote";

/** How long an awaiting step stays open for its follow-up (a frame, on a slow device). */
const PAGED_FOLLOW_UP_WAIT_MS = 1500;

/**
 * Whether `tr` is a step of the writer's own Undo history, as
 * prosemirror-history decides it: `addToHistory: false` keeps it out, and a
 * transaction a plugin appended follows the one it was appended to. What a
 * room update or a layout pass made a normalizer append (one that states no
 * meta of its own) is the view's own, like any normalization no gesture
 * authored -- not an undoable step of the writer's. (ProseMirror names the
 * transaction an appended one follows only after the plugins' filters, so
 * the fill's queue tells its normalizers apart by its own dispatch.)
 */
function addsToHistory(tr: Transaction): boolean {
  if (tr.getMeta("addToHistory") === false) return false;
  const root = tr.getMeta("appendedTransaction") as Transaction | undefined;
  return !(root && root !== tr && root.getMeta("addToHistory") === false);
}

/**
 * Equal as the room holds them. A field, cross-reference or phonetic guide
 * keeps its own format in `runStyle`; the marks it is drawn with are
 * derived from that when a document opens (DocumentInlineAtomRunStyle),
 * a sync the room never receives. Compared with them, every such block
 * read as changed by the room, at every change anywhere.
 */
function sameRoomNode(local: ProseMirrorNode, room: ProseMirrorNode): boolean {
  if (local === room || local.eq(room)) return true;
  const marks = carriesInlineAtomRunStyle(local) ? local.marks : room.marks;
  if (!local.hasMarkup(room.type, room.attrs, marks)) return false;
  if (local.isText) return local.text === room.text;
  return sameRoomContent(local.content, room.content);
}

function sameRoomContent(local: Fragment, room: Fragment): boolean {
  if (local.childCount !== room.childCount) return false;
  for (let index = 0; index < local.childCount; index += 1) {
    if (!sameRoomNode(local.child(index), room.child(index))) return false;
  }
  return true;
}
/** Awareness fields a view of the LEGACY (non-causal) page room publishes
 * when it binds the room's logical document through this bridge. */
export const LEGACY_ROOM_MODEL_FIELD = "docModel";
export const LEGACY_ROOM_MODEL_WRITER_FIELD = "docModelWriter";
export const LEGACY_ROOM_LOGICAL_MODEL = "logical-v1";
/** Quiet period (no room update, no membership change) before an older
 * physical room is upgraded in place. */
const LEGACY_ROOM_UPGRADE_SETTLE_MS = 2500;
// How long the keys rest before the room is read back after this view's own writes.
const CANONICAL_VERIFY_PAUSE_MS = 750;
const REBALANCE_META = "structuralRebalance";
const HISTORY_META = "mrnotePagedCanonicalHistory";
/** How long the hand has to be off the keys before the first edit is prepared for. */
const FIRST_EDIT_WARM_LOOK_MS = 2500;
/** How often a first fill held for its kept layout looks again. */
const FILL_RECEIPT_POLL_MS = 60;
/** How long the pages' layout has to have stood still for it too, and how
 * long it is waited for at most. */
const FIRST_EDIT_WARM_LAYOUT_QUIET_MS = 4000;
const FIRST_EDIT_WARM_LAYOUT_WAIT_MAX_MS = 60_000;
const TRUSTED_META = "mrnoteReadOnlyTrustedDocumentMutation";

/**
 * Whether the pages are still being laid out: shown before they settled, a
 * step of the background, a flag that moved a moment ago, faces still coming
 * in (each one lays its pages out again). The first edit's preparation (the
 * whole document's position map: 3-4 s in one task on R286-D01's 124 pages)
 * waits for that to end; run in the middle of it, it stood in the way of the
 * passes still to come.
 */
export function layoutStillMoving(root: HTMLElement): boolean {
  return root.hasAttribute(PAGE_LEDGER_DEFERRED_ATTR) || root.hasAttribute(PAGE_LEDGER_BACKGROUND_ATTR) ||
    root.hasAttribute(PAGED_BULK_FILL_ATTR) || msSinceLayoutStatusChange(root) < FIRST_EDIT_WARM_LAYOUT_QUIET_MS ||
    root.ownerDocument.fonts?.status === "loading";
}
export type PagedDocumentBridgeIssue = "legacy-room" | "model-not-negotiated" | "invalid-document" | "position-unavailable"
  // The room's document could not be read at all when the page opened: the
  // page says so instead of showing one empty paragraph (NoteEditor).
  | "unreadable-document";

export interface PagedDocumentBridgeOptions {
  session: CollabSession;
  schema: Schema;
  editor?: Editor;
  field?: string;
  /** Must eventually come from a server-verified model negotiation. Local
   * sessions need no network negotiation; a missing network verdict blocks. */
  canonicalNetworkModel?: () => boolean;
  mayWrite?: () => boolean;
  /**
   * The Y.Doc already holds the server's durable originals (an adopted causal
   * checkpoint) before any room replay. Local edits may then enter it while
   * the room is still connecting or has given up: a later replay merges into
   * the same lineage, and every save is a causal PATCH the server merges.
   * Without it, input waits for the room's first replay (a physical or
   * JSON-seeded room must not take edits on a copy the room may replace).
   */
  durableOriginals?: boolean;
  onIssue?: (issue: PagedDocumentBridgeIssue | null) => void;
  







  upgradeLegacyRoom?: boolean;
  /**
   * The layout receipt this browser last reached for the page
   * (lib/docs-layout-receipt-cache), for the first fill of an empty view: a
   * long document then opens on its finished pages instead of provisional
   * sheets. Applied only when it restores the room's content exactly.
   */
  initialLayoutReceipt?: (canonical?: JSONContent, root?: object) => unknown;
  /** The whole-block page distribution kept with it: the fallback when the
   * receipt does not restore (a different form of the same content). */
  initialLayoutSizes?: () => number[] | null;
  /** Whether the layout the first fill may open on is still being read (a
   * long document's kept layout, asked beside the page): the fill waits for
   * it, for as long as this says, instead of laying the pages out cold. */
  initialLayoutReceiptPending?: () => boolean;
}

interface RelativeSelection {
  /** "cell": anchor and head are the positions before a CellSelection's
   * anchor and head cells, so Undo selects the cells again. */
  kind: "text" | "node" | "all" | "cell";
  anchor: Y.RelativePosition;
  head: Y.RelativePosition;
  anchorHint: number;
  headHint: number;
  nodeType?: string;
}
interface PositionDerivation { base: ProseMirrorNode; steps: readonly Step[]; docs: readonly ProseMirrorNode[]; depth: number }
/** Longest chain of unmapped transactions a map is derived through. */
const MAX_POSITION_DERIVATION_DEPTH = 32;
/** A long document's first fill: blocks in its first dispatch (the sheets a
 * reader sees first), then about this many per later dispatch, one dispatch
 * per main-thread task. */
/**
 * Whether a long document's first fill goes in chunks: not when its later
 * sheets go in hollow, which is one cheap dispatch. A document docs-hollow
 * draws whole all the same (one with footnotes, a short one, an AI draft's
 * proof copy) is filled in chunks like any other: asked only whether the
 * surface was on, R286-D01 (footnotes) went in as one dispatch of its 124
 * pages, a 2.9 s task before its first page showed.
 * `window.__docsHollowFillWhole = true` fills such a document at once (A/B).
 */
export function fillsInChunks(hollowOn: boolean, drawsHollow: boolean): boolean {
  if (!hollowOn) return true;
  if ((globalThis as { __docsHollowFillWhole?: boolean }).__docsHollowFillWhole === true) return false;
  return !drawsHollow;
}

const FILL_FIRST_BLOCKS = 240;
const FILL_CHUNK_BLOCKS = 320;
/** How long layout-only updates must pause before their content check. */
const LAYOUT_ONLY_CHECK_PAUSE_MS = 600;
/** Layout-only runs kept for the check at the pause, at most. */
const MAX_CLOSED_LAYOUT_ONLY_RUNS = 4;
/** The same while the page ledger lays the document out in the background:
 * its passes close a run each, and proving the oldest at once mapped the
 * whole book in the middle of the layout (a 2.3 s task on the 298-page book).
 * They are proved after the pause instead; past this many, the oldest is
 * proved at once, so the documents kept stay bounded. */
const MAX_CLOSED_LAYOUT_ONLY_RUNS_WHILE_BUSY = 32;
interface HistorySelection {
  before: RelativeSelection | null;
  after: RelativeSelection | null;
  /** The room document before and after the step (see history): shared
   * nodes, so keeping them costs little. */
  beforeDoc?: ProseMirrorNode;
  afterDoc?: ProseMirrorNode | null;
}
/** An attribute publication's carets are read when its Undo step is made,
 * in its own transaction (publishDocumentAttrs reads them lazily). The same
 * object is kept: a gesture sets its `after` once it has published. */
function settledCarets(history: HistorySelection): HistorySelection {
  void history.before;
  return history;
}
interface CanonicalTextEdit { from: Y.RelativePosition; to: Y.RelativePosition; inserted: string; removed: string }

/** How far around the caret a composition held from the room's updates may
 * reach (characters before / after the caret when the hold begins). */
const COMPOSITION_HOLD_BEFORE = 256;
const COMPOSITION_HOLD_AFTER = 64;
/** How often a held composition looks whether its input method let go (a
 * view without the composition guard's settle signal). */
const COMPOSITION_HOLD_POLL_MS = 250;

/**
 * A view whose input method is composing while the room changes: the room's
 * update is not drawn under the composition (a redraw of the text node the
 * pinyin is in ends the composition, and the browser keeps the pinyin typed
 * so far as text: "ninih" before the Hanzi, two people typing in one
 * paragraph). The room keeps every update; the view takes them when the
 * composition settles, with what was composed meanwhile written into the room
 * at the place it was composed.
 */
interface CompositionHold {
  view: EditorView;
  /** The view's document when the hold began (the room's, then). */
  held: ProseMirrorNode;
  /** Room positions of the held document's positions near the caret: bound
   * to the character after (`right`) and before (`left`). */
  anchors: Map<number, { right: Y.RelativePosition; left: Y.RelativePosition }>;
  /** What the composition changed since: `[heldFrom, heldTo)` of the held
   * document is `[from, to)` of the view's. Null: nothing yet. */
  region: { heldFrom: number; heldTo: number; from: number; to: number } | null;
  /** Something other than text typed in one paragraph changed the view. */
  complex: boolean;
  release: () => void;
}
interface BridgeState {
  doc: Y.Doc;
  type: Y.XmlFragment;
  binding: { mapping: Map<Y.AbstractType<unknown>, ProseMirrorNode | ProseMirrorNode[]> };
  bridge: PagedDocumentBridge;
  changed: boolean;
  layoutOnly: boolean;
  historical: boolean;
  trustedHydration: boolean;
  trustedPublication: boolean;
  /** A change of the batch is kept out of the writer's Undo but is still the
   * document's to save (SAVED_WITHOUT_HISTORY_META): published outside
   * personal history, never left view-only. */
  savedWithoutHistory: boolean;
  textEdit: CanonicalTextEdit | null;
  sceneChange: boolean;
  /** A transaction of the batch closed the history group (closeHistory). */
  closesHistory: boolean;
  /** A transaction of the batch awaits a derived follow-up (PAGED_AWAITS_FOLLOW_UP_META). */
  awaitsFollowUp: boolean;
  /** Every authored change of the batch is a derived follow-up (PAGED_FOLLOW_UP_META). */
  followUp: boolean;
  /** A transaction of the batch replaced the whole document: content being
   * loaded (setContent, a template, an import), not a gesture inside it. */
  replacesDocument: boolean;
  /** When the batch's first and last authored transactions were made
   * (their `time`: the input event's, typingRunTime), null without one. */
  authoredFrom: number | null;
  authoredTo: number | null;
  /** Every change of the batch only inserts, none at the caret, and none set
   * a selection: the caret's room position (relative to the item it follows)
   * is the one it had before the batch. A landed stroke's batch, after which
   * capturing the caret again walked the room fragment to it. */
  keepsCaret: boolean;
}

/** `tr` inserts only, never at its selection's ends (in each step's own
 * document), and sets no selection. */
function insertsAwayFromCaret(tr: Transaction, selection: Selection): boolean {
  if (tr.selectionSet) return false;
  let anchor = selection.anchor, head = selection.head;
  for (let index = 0; index < tr.steps.length; index += 1) {
    const step = tr.steps[index];
    if (!(step instanceof ReplaceStep)) return false;
    const { from, to } = step as unknown as { from: number; to: number };
    if (from !== to || from === anchor || from === head) return false;
    const map = step.getMap();
    anchor = map.map(anchor, 1);
    head = map.map(head, 1);
  }
  return true;
}

/** prosemirror-history's closeHistory() meta (its PluginKey "closeHistory"). */
const CLOSE_HISTORY_META = "closeHistory$";

/** A 31-bit Yjs client identity from a string (FNV-1a), never 0. */
function legacyRoomUpgradeClientId(key: string): number {
  let hash = 0x811c9dc5;
  for (let index = 0; index < key.length; index += 1) {
    hash ^= key.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return ((hash >>> 1) || 1);
}

class PagedMappablePosition extends MappablePosition {
  constructor(position: number, readonly relative: Y.RelativePosition | null) { super(position); }
  override toJSON() { return { position: this.position, relative: this.relative }; }
}

/** y-tiptap keeps a caret with the character on its left, except in an empty
 * block or after a block's last inline node: there it records "the end of
 * the block", which follows whatever a peer types there. An idle
 * collaborator's caret then travelled with the other person's typing, and
 * their own next keystroke landed inside the other person's word. The caret
 * keeps to the content on its left there too. */
function leftBoundPosition(doc: Y.Doc, position: Y.RelativePosition): Y.RelativePosition {
  if (position.item !== null || position.assoc < 0) return position;
  const absolute = Y.createAbsolutePositionFromRelativePosition(position, doc);
  return absolute ? Y.createRelativePositionFromTypeIndex(absolute.type, absolute.index, -1) : position;
}

/** One room block as y-tiptap's reader writes it in a whole room's JSON. */
function roomBlockJson(block: Y.XmlElement): Record<string, unknown> {
  const inner = yXmlFragmentToProsemirrorJSON(block as unknown as Y.XmlFragment) as { content?: unknown[] };
  const attrs = block.getAttributes();
  return {
    type: block.nodeName,
    ...(Object.keys(attrs).length ? { attrs } : {}),
    ...(inner.content?.length ? { content: inner.content } : {}),
  };
}

/** The binding y-tiptap's reader builds (each XmlElement to its node), for a
 * document already read from the same fragment. */
function roomBinding(fragment: Y.XmlFragment, doc: ProseMirrorNode): ReturnType<typeof initProseMirrorDoc>["meta"] {
  const meta: ReturnType<typeof initProseMirrorDoc>["meta"] = { mapping: new VersionedMapping(), isOMark: new Map() };
  const elementsOf = (parent: Y.XmlFragment | Y.XmlElement) =>
    // A bare inline node at the room's top level is not in the document
    // (withoutBareTopLevelInlineNodes): it binds to nothing.
    parent.toArray().filter((child): child is Y.XmlElement => child instanceof Y.XmlElement &&
      !(parent === fragment && isBareTopLevelInlineNode(doc.type.schema, child.nodeName)));
  const modelsOf = (node: ProseMirrorNode) => {
    const models: ProseMirrorNode[] = [];
    node.forEach(child => { if (!child.isText) models.push(child); });
    return models;
  };
  // `named`: a repaired read binds an element only to a node of its own
  // type (a room read as stated is bound as y-tiptap's reader binds it).
  const bind = (parent: Y.XmlFragment | Y.XmlElement, node: ProseMirrorNode, into: Array<[Y.XmlElement, ProseMirrorNode]>, named: boolean) => {
    const elements = elementsOf(parent), models = modelsOf(node);
    if (elements.length !== models.length) throw new Error("Room binding mismatch");
    elements.forEach((element, index) => {
      if (named && element.nodeName !== models[index].type.name) throw new Error("Room binding mismatch");
      into.push([element, models[index]]);
      bind(element, models[index], into, named);
    });
  };
  const read = canonicalReadBlocksOf(doc);
  if (read === undefined) {
    const pairs: Array<[Y.XmlElement, ProseMirrorNode]> = [];
    bind(fragment, doc, pairs, false);
    for (const [element, model] of pairs) meta.mapping.set(element, model);
    return meta;
  }
  // A room the read repaired (an unknown node opened as its words, a block
  // lifted out of a paragraph, an undeclared attribute left out): only the
  // blocks read as they stand are bound. The others are no part of the
  // binding, so the first write of this client (y-tiptap rewrites what is
  // unbound and differs) puts their read form in the room, and from then on
  // the room says what the editor says -- what every save must prove. A
  // reader that never writes leaves the room as it is.
  const elements = elementsOf(fragment), models = modelsOf(doc);
  if (!read || read.length !== elements.length || read.reduce((sum, block) => sum + block.span, 0) !== models.length) return meta;
  let at = 0;
  read.forEach((block, index) => {
    const model = models[at];
    at += block.span;
    if (!block.same || block.span !== 1) return;
    const pairs: Array<[Y.XmlElement, ProseMirrorNode]> = [];
    try {
      if (elements[index].nodeName !== model.type.name) return;
      pairs.push([elements[index], model]);
      bind(elements[index], model, pairs, true);
    } catch { return; }
    for (const [element, bound] of pairs) meta.mapping.set(element, bound);
  });
  return meta;
}

/** A position in `from` at the same content in `to`, when the two documents
 * differ by one contiguous change: the view holds a keystroke before it is
 * published and an IME candidate the room has not heard of, or the room a
 * peer's edit the view has not drawn yet. A position inside the change keeps
 * its offset, clamped to the changed range. */
function acrossDifference(from: ProseMirrorNode, to: ProseMirrorNode, position: number): number {
  // Equal sizes (the usual case, or attributes only) shift nothing.
  if (from.content.size === to.content.size) return position;
  const start = from.content.findDiffStart(to.content);
  if (start === null || position <= start) return position;
  const end = from.content.findDiffEnd(to.content);
  if (!end) return position;
  const overlap = Math.max(0, start - Math.min(end.a, end.b));
  const fromEnd = end.a + overlap, toEnd = end.b + overlap;
  return position >= fromEnd ? position + toEnd - fromEnd : Math.min(position, toEnd);
}

/** A block this view just split off with Enter: the text it moved out of the
 * block before it, and what this client knew of every other client then. */
interface SplitRecord {
  element: Y.XmlElement;
  /** Where the moved text starts in the new block, kept through later typing. */
  start: Y.RelativePosition | null;
  text: string;
  known: Map<number, number>;
  at: number;
  settled: WeakSet<Y.XmlElement>;
}
const SPLIT_RECORD_TTL_MS = 60_000;
const SPLIT_RECORD_LIMIT = 16;

function plainText(type: Y.XmlElement | Y.XmlText): string {
  if (type instanceof Y.XmlText) {
    return (type.toDelta() as Array<{ insert: unknown }>).map(op => typeof op.insert === "string" ? op.insert : "").join("");
  }
  return type.toArray().map(child => child instanceof Y.XmlElement || child instanceof Y.XmlText ? plainText(child) : "").join("");
}

/** A semantic Yjs binding. Page wrappers and continuation edits are observed
 * only to update local position maps; they are never written into the Y.Doc. */
/** docs-hollow: the sheets after the first FILL_FIRST_BLOCKS blocks of a
 * long document's fill (at the start of `tr`'s document) drawn hollow; the
 * page ledger plans them from what it can read without their DOM and draws
 * whole the few it measures at a time. */
function fillHollowFrom(tr: Transaction, pages: Fragment): void {
  const hollow: number[] = [];
  let pos = 0;
  let blocks = 0;
  pages.forEach((page) => {
    if (blocks >= FILL_FIRST_BLOCKS && page.type.name === "pageBlock") hollow.push(pos);
    blocks += page.childCount;
    pos += page.nodeSize;
  });
  if (hollow.length) tr.setMeta(pageHollowSheetsKey, { hollow });
}

export class PagedDocumentBridge {
  readonly schema: Schema;
  readonly fragment: Y.XmlFragment;
  readonly undoManager: Y.UndoManager;
  readonly plugins: Plugin[];
  private readonly localOrigin = Symbol("canonical-user-edit");
  private readonly seedOrigin = Symbol("canonical-initial-seed");
  private readonly nonHistoryOrigin = Symbol("canonical-system-edit");
  private readonly positions = new WeakMap<ProseMirrorNode, PagedDocumentPositionMap>();
  /** How a document without a map yet was reached from an earlier one, so
   * its map can be derived from that one's instead of re-walking the whole
   * document on every keystroke. Chains are short and end at a mapped doc. */
  private readonly derivations = new WeakMap<ProseMirrorNode, PositionDerivation>();
  private readonly consumedTextEdits = new WeakSet<CanonicalTextEdit>();
  private view: EditorView | null = null;
  private lastViewState: EditorState | null = null;
  private viewBinding: object | null = null;
  private canonical: ProseMirrorNode;
  




  private confirmedPair: { doc: ProseMirrorNode; canonical: ProseMirrorNode } | null = null;
  /** The room document this view last held exactly: the common base when
   * the view holds something the room does not (canonical-block-merge). */
  private agreedCanonical: ProseMirrorNode | null = null;
  private recoverTimer: ReturnType<typeof setTimeout> | null = null;
  private recoverAttempts = 0;
  private meta: ReturnType<typeof initProseMirrorDoc>["meta"] = { mapping: new VersionedMapping(), isOMark: new Map() };
  private closed = false;
  private publishing = false;
  /** When the last authored change published into personal history was made. */
  private lastAuthoredTo: number | null = null;
  // A write by anything but this view's own publication reached the room
  // while it published: the room must be read back, not assumed.
  private foreignDuringPublish = false;
  // The room's document is this view's own last publication, not yet read
  // back; the read waits for a pause (adoptPublished).
  private adoptedUnverified = false;
  private canonicalVerifyTimer: ReturnType<typeof setTimeout> | null = null;
  /** A room update arrived inside this view's own publishing section. The
   * section reads the room back once when it has written (every section
   * does, or flushes this at its end), so the update is not read twice. */
  private canonicalReadPending = false;
  /** Each room block as last read back, by the node it read as (verifyAdopted
   * reads back only the blocks written or changed since: the whole room
   * after every pause in the writing was a full parse and a deep comparison
   * of every block, 80 ms on a Mac and several times that on an iPad for a
   * 3,000-stroke note). */
  private readonly verifiedElements = new WeakMap<Y.XmlElement, ProseMirrorNode>();
  /** Room blocks written or changed since they were last read back. */
  private readonly touchedElements = new Set<Y.XmlElement>();
  private originalNormalized = false;
  private applying = false;
  private refreshQueued = false;
  private caretChanged: ReturnType<typeof createCaretChangeFilter> | null = null;
  // Carets are redrawn when a member's caret or name moves, not when only
  // their live ink does (lib/collab/ink-presence: many times a second).
  private selectionPublishPending = false;
  private onPresenceChange = (changes: { added: number[]; updated: number[]; removed: number[] }) => {
    if (this.selectionPublishPending && changes.added.length > 0) this.deferSelectionPublish();
    // A caret moving is redrawn on its layer, without a transaction of
    // this view's: on a long document each one ran every plugin.
    if (!this.caretChanged || this.caretChanged(changes)) this.caretOverlay?.schedule();
  };
  private caretOverlay: PeerCaretOverlay | null = null;
  private splits: SplitRecord[] = [];
  private beforeRemote: RelativeSelection | null = null;
  /** The room positions near a composing view's caret, captured when a run of
   * Y transactions starts and the view still matches the room: the hold that
   * begins if the run is another writer's update (`beginCompositionHold`). */
  private compositionAnchors: {
    doc: ProseMirrorNode; anchors: CompositionHold["anchors"];
    /** The room's document then, and the caret's paragraph in it. */
    canonical: ProseMirrorNode; block: { from: number; to: number };
  } | null = null;
  private compositionHold: CompositionHold | null = null;
  /** Whether the reader has been in the editor yet: it had focus, or a
   * transaction of the view's own placed a caret. */
  private readerWasHere = false;
  private historyTarget: RelativeSelection | null | undefined;
  /** The carets of the step an Undo / Redo is applying: the stack item it
   * pushes on the other stack is the same step, so it keeps them. */
  private historyCarets: HistorySelection | undefined;
  private localHistory: HistorySelection | null = null;
  /** The Undo step a history-closing gesture opened, while that gesture's
   * dispatch may still write (its document attributes publish after this
   * view update, from the editor's transaction event). */
  private openStep: { readonly awaitsFollowUp: boolean; settled: boolean } | null = null;
  private issue: PagedDocumentBridgeIssue | null = null;
  private established = false;
  private explicitSaveRefusal: string | null = null;
  private recoveryConflicts: Array<{ kept: string | null; copy: string | null }> = [];
  private cursorKey: string | null = null;
  private unsubscribeConnection: (() => void) | null = null;
  private upgradeTimer: ReturnType<typeof setTimeout> | null = null;
  /** Layout-only updates (the page ledger moving blocks between sheets)
   * waiting for their check that the shared content did not change: the
   * document before the first of them, the latest, and the pause timer. */
  private layoutOnlyCheck: { base: ProseMirrorNode; latest: ProseMirrorNode; timer: ReturnType<typeof setTimeout> | null } | null = null;
  private selectionPublishTimer: ReturnType<typeof setTimeout> | null = null;
  /** Layout-only runs ended by a view-only batch, checked at the pause. */
  private closedLayoutOnlyRuns: Array<{ base: ProseMirrorNode; end: ProseMirrorNode }> = [];
  /** A long document's first fill still arriving (`fillInChunks`): the room
   * document it shows, its provisional sheets, and the next one to add. */
  private filling: { canonical: ProseMirrorNode; pages: readonly ProseMirrorNode[]; next: number; filled: number; lane: BackgroundLane } | null = null;
  /** Authored edits made while the fill was arriving (a command, never a
   * keystroke: the view is not editable): applied once it is complete,
   * mapped through everything the view took meanwhile (`flushFillEdits`). */
  private fillEdits: Array<{ steps: readonly Step[]; since: Mapping; authored: boolean }> = [];
  /** Waits for write authority to come back to apply edits the write gate
   * held (`fillEdits`, outside a fill). */
  private heldEditsTimer: ReturnType<typeof setInterval> | null = null;
  private heldEditsSince = 0;

  constructor(readonly options: PagedDocumentBridgeOptions) {
    this.schema = canonicalDocumentSchema(options.schema);
    this.fragment = options.session.ydoc.getXmlFragment(options.field ?? "default");
    this.canonical = this.schema.topNodeType.create(null, Fragment.empty);
    this.undoManager = new Y.UndoManager([this.fragment, options.session.ydoc.getMap(COLLABORATIVE_DOCUMENT_ATTRS_MAP)],
      { trackedOrigins: new Set([this.localOrigin]) });
    this.undoManager.on("stack-item-added", this.onHistoryAdded);
    this.undoManager.on("stack-item-updated", this.onHistoryUpdated);
    options.session.ydoc.on("beforeAllTransactions", this.onBeforeYTransactions);
    options.session.ydoc.on("afterAllTransactions", this.onAfterYTransactions);
    options.session.ydoc.on("update", this.onYUpdate);
    options.session.ydoc.on("afterTransaction", this.onAfterYTransaction);
    this.caretChanged = createCaretChangeFilter(options.session.awareness);
    options.session.awareness.on("change", this.onPresenceChange);
    if (options.upgradeLegacyRoom) {
      this.advertiseLegacyRoomModel();
      options.session.awareness.on("change", this.scheduleLegacyRoomUpgrade);
    }
    this.readCanonical();
    this.refreshNetworkModel();
    const readMapping = () => this.meta.mapping;
    const binding = { get mapping() { return readMapping(); } };
    this.plugins = [
      new Plugin<BridgeState>({
        key: ySyncPluginKey,
        state: {
          init: () => ({ doc: options.session.ydoc, type: this.fragment, binding, bridge: this, changed: false, layoutOnly: true, historical: false, trustedHydration: false, trustedPublication: false, savedWithoutHistory: false, textEdit: null, sceneChange: false, closesHistory: false, awaitsFollowUp: false, followUp: false, replacesDocument: false, authoredFrom: null, authoredTo: null, keepsCaret: true }),
          apply: (tr, previous, oldState) => {
            if (tr.docChanged) this.noteDerivation(oldState.doc, tr);
            if (tr.docChanged && this.compositionHold && !tr.getMeta(PAGED_CANONICAL_REMOTE_META)) this.noteHeldComposition(tr);
            const startsBatch = oldState === this.lastViewState;
            // The room's document attributes (page setup, styles) projected
            // onto the editor (NoteEditor's attrs binding) are the room's
            // already: an update that changes only them is applied like a
            // remote one. Taken for an authored step, it mapped the whole
            // book to the room and published the editor's content back as an
            // undoable edit: ~1.2 s in the one task that opens the 298-page
            // book, which sets the attributes right after the fill.
            const remote = Boolean(tr.getMeta(PAGED_CANONICAL_REMOTE_META)) ||
              (tr.getMeta(COLLABORATIVE_DOCUMENT_ATTRS_REMOTE_META) === true && tr.docChanged &&
                tr.steps.every((step) => step instanceof DocAttrStep));
            // The reader's own caret: one placed by a transaction of this view
            // that is not a room update, a layout pass or maintenance.
            if (!this.readerWasHere && tr.selectionSet && !remote && tr.getMeta(REBALANCE_META) !== true && addsToHistory(tr)) this.readerWasHere = true;
            const reason = tr.getMeta(TRUSTED_META);
            // An inline atom's marks drawn from its own attribute
            // (inlineAtomRunStyle) show what the document states already: the
            // room carries no mark on an inline atom. Not an authored step:
            // drawn on a 298-page book's atoms right after the fill, it
            // mapped the whole book to the room and published it back as an
            // undoable edit (~1.7 s of the task that opens the book). One
            // that sets the attribute (the user's own formatting) still is.
            const drawsAtomMarks = !remote && tr.getMeta(INLINE_ATOM_RUN_STYLE_META) === true &&
              tr.steps.every((step) => !(step instanceof AttrStep));
            let textEdit = startsBatch ? null : previous.textEdit;
            if (!remote && tr.getMeta(REBALANCE_META) !== true && !oldState.doc.content.eq(tr.doc.content)) {
              textEdit = (startsBatch || !previous.changed) && addsToHistory(tr)
                ? this.captureTextEdit(tr, oldState) : null;
            }
            const authored = tr.docChanged && !remote && !drawsAtomMarks && addsToHistory(tr);
            return {
              ...previous,
              authoredFrom: (startsBatch ? null : previous.authoredFrom) ?? (authored ? tr.time : null),
              authoredTo: authored ? tr.time : startsBatch ? null : previous.authoredTo,
              keepsCaret: (startsBatch || previous.keepsCaret) &&
                (!tr.docChanged || remote || insertsAwayFromCaret(tr, oldState.selection)),
              textEdit,
              sceneChange: (!startsBatch && previous.sceneChange) || (!remote && tr.getMeta(WHITEBOARD_SCENE_META) === true),
              closesHistory: (!startsBatch && previous.closesHistory) || (!remote && tr.getMeta(CLOSE_HISTORY_META) === true),
              awaitsFollowUp: (!startsBatch && previous.awaitsFollowUp) || (!remote && tr.getMeta(PAGED_AWAITS_FOLLOW_UP_META) === true),
              followUp: (startsBatch || previous.followUp) && (!tr.docChanged || remote || tr.getMeta(PAGED_FOLLOW_UP_META) === true),
              replacesDocument: (!startsBatch && previous.replacesDocument) || (!remote && tr.steps.some((step, index) => {
                const range = step as unknown as { from?: number; to?: number };
                return range.from === 0 && range.to === tr.docs[index].content.size;
              })),
              changed: (!startsBatch && previous.changed) || (tr.docChanged && !remote),
              layoutOnly: (startsBatch || previous.layoutOnly) && (!tr.docChanged || remote || tr.getMeta(REBALANCE_META) === true),
              // The settled wake-up of a typed IME composition carries the
              // work deferred from that edit: it is authored with it, as it
              // would have been in an ordinary keystroke's batch.
              historical: (!startsBatch && previous.historical) || (tr.docChanged && !remote && !drawsAtomMarks && addsToHistory(tr)) ||
                (!remote && isImeDeferredUserEdit([tr])),
              trustedHydration: (!startsBatch && previous.trustedHydration) || ((reason === "hydration-content" || reason === "content-restoration") && tr.getMeta("preventUpdate") === true && tr.getMeta("addToHistory") === false),
              trustedPublication: (!startsBatch && previous.trustedPublication) || (!remote && !drawsAtomMarks && addsToHistory(tr) && isTrustedReadOnlyDocumentMutation(tr, options.editor)),
              savedWithoutHistory: (!startsBatch && previous.savedWithoutHistory) ||
                (!remote && tr.docChanged && tr.getMeta(SAVED_WITHOUT_HISTORY_META) === true),
            };
          },
        },
        // Not editable while the room cannot take input: the write gate
        // below refuses the transaction, so an editable surface there
        // swallowed keystrokes without a trace (a member who started typing
        // in the second before their room was ready lost those characters).
        props: { editable: () => this.acceptsInput() && this.filling === null },
        filterTransaction: (tr, state) => {
          if (!tr.docChanged) return true;
          if (tr.getMeta(PAGED_CANONICAL_REMOTE_META)) {
            for (const edit of this.fillEdits) edit.since.appendMapping(tr.mapping);
            return true;
          }
          // While a long document is still being filled in, the view holds
          // only part of the room: an authored edit would publish a document
          // without the rest. Layout and view-only passes go on.
          if (this.filling !== null && (tr.getMeta("addToHistory") !== false || isImeDeferredUserEdit([tr]))) {
            // One that arrives inside the fill's own dispatch is a normalizer
            // appended to the arriving sheets: applied with the rest, but no
            // Undo step. Replayed as the writer's, the run style an inline
            // equation on a book's first sheets drew as marks was the first
            // Undo step of every opening: Ctrl+Z right after opening undid it
            
            this.fillEdits.push({ steps: tr.steps, since: new Mapping(), authored: !this.applying });
            return false;
          }
          // Verified remote attributes hydrate even before local admission.
          // This marker cannot grant a body mutation through the write gate.
          const passes = (tr.getMeta(COLLABORATIVE_DOCUMENT_ATTRS_REMOTE_META) && state.doc.content.eq(tr.doc.content)) ||
            isTrustedReadOnlyDocumentMutation(tr, options.editor) || this.mayPublish();
          if (passes) {
            if (this.fillEdits.length > 0) for (const edit of this.fillEdits) edit.since.appendMapping(tr.mapping);
            return true;
          }
          // The view took input the room cannot take this instant (its
          // write authority blinked while it reconnects): hold the writer's
          // edit and apply it once the room takes input again, instead of
          
          if (this.acceptsInput() && this.view?.editable !== false && tr.getMeta("addToHistory") !== false &&
              this.fillEdits.length < 10_000) {
            this.fillEdits.push({ steps: tr.steps, since: new Mapping(), authored: true });
            this.holdEditsUntilWritable();
          }
          return false;
        },
        view: view => {
          const previous = this.lastViewState;
          const binding = {};
          this.view = view; this.viewBinding = binding;
          if (!this.readerWasHere) view.dom.addEventListener("focusin", () => { this.readerWasHere = true; }, { once: true });
          // ProseMirror rebuilds every plugin view when an unrelated feature
          // registers/unregisters a plugin. The bridge and personal history
          // belong to the editor extension, not that temporary plugin view.
          if (previous && previous.doc !== view.state.doc) this.onLocalUpdate(view, previous);
          else {
            this.lastViewState = view.state;
            if (this.fragment.length > 0) this.applyCanonical(null);
          }
          return {
            update: (current, previous) => this.onLocalUpdate(current, previous),
            destroy: () => {
              if (this.viewBinding === binding) { this.view = null; this.viewBinding = null; }
            },
          };
        },
      }),
      new Plugin({
        key: yUndoPluginKey,
        state: {
          init: () => ({ undoManager: this.undoManager, hasUndoOps: false, hasRedoOps: false }),
          apply: () => ({ undoManager: this.undoManager, hasUndoOps: this.undoManager.undoStack.length > 0, hasRedoOps: this.undoManager.redoStack.length > 0 }),
        },
      }),
      new Plugin({
        key: yCursorPluginKey,
        view: (view) => {
          const overlay = new PeerCaretOverlay(view, () => this.peerCarets(view.state));
          this.caretOverlay = overlay;
          return {
            update: (current) => overlay.update(current),
            destroy: () => { overlay.destroy(); if (this.caretOverlay === overlay) this.caretOverlay = null; },
          };
        },
      }),
    ];
    this.unsubscribeConnection = options.session.network?.subscribeConnection(() => {
      this.refreshNetworkModel();
      if (options.upgradeLegacyRoom) {
        this.advertiseLegacyRoomModel();
        this.scheduleLegacyRoomUpgrade();
      }
      this.scheduleRefresh();
    }) ?? null;
  }

  private advertiseLegacyRoomModel() {
    if (this.closed) return;
    const awareness = this.options.session.awareness;
    const network = this.options.session.network;
    const writer = !network || (network.connected && network.writable);
    const local = awareness.getLocalState();
    if (local?.[LEGACY_ROOM_MODEL_FIELD] === LEGACY_ROOM_LOGICAL_MODEL &&
        local?.[LEGACY_ROOM_MODEL_WRITER_FIELD] === writer) return;
    awareness.setLocalStateField(LEGACY_ROOM_MODEL_FIELD, LEGACY_ROOM_LOGICAL_MODEL);
    awareness.setLocalStateField(LEGACY_ROOM_MODEL_WRITER_FIELD, writer);
  }

  private scheduleLegacyRoomUpgrade = () => {
    if (!this.options.upgradeLegacyRoom || this.closed || this.issue !== "legacy-room") return;
    if (this.upgradeTimer !== null) clearTimeout(this.upgradeTimer);
    this.upgradeTimer = setTimeout(() => {
      this.upgradeTimer = null;
      this.tryUpgradeLegacyRoom();
    }, LEGACY_ROOM_UPGRADE_SETTLE_MS);
  };

  /** One elected writer upgrades a quiet physical room once no member
   * present still binds the physical tree. Returns whether it upgraded. */
  tryUpgradeLegacyRoom(): boolean {
    if (!this.options.upgradeLegacyRoom || this.closed || this.issue !== "legacy-room") return false;
    const network = this.options.session.network;
    if (network && (network.failed || !network.ready || !network.connected || !network.writable)) return false;
    const awareness = this.options.session.awareness;
    const states = awareness.getStates();
    if (!states.has(awareness.clientID)) return false;
    let elected: number | null = null;
    for (const [clientId, state] of states) {
      // A member that does not advertise the logical model still binds the
      // physical tree: upgrading under it would hand it a document shape it
      // repairs by rewriting every paragraph.
      if (state?.[LEGACY_ROOM_MODEL_FIELD] !== LEGACY_ROOM_LOGICAL_MODEL) return false;
      if (state?.[LEGACY_ROOM_MODEL_WRITER_FIELD] === true && (elected === null || clientId < elected)) elected = clientId;
    }
    if (elected !== awareness.clientID) return false;
    return this.upgradeLegacyRoom();
  }

  /**
   * Replace the physical page tree with its logical document in ONE update
   * that any member starting from the same room state produces byte for
   * byte (a fixed client identity derived from that state). Two members that
   * both believe they were elected therefore merge into one copy, never two.
   */
  private upgradeLegacyRoom(): boolean {
    const room = this.options.session.ydoc;
    const field = this.options.field ?? "default";
    let staging: Y.Doc | null = null;
    try {
      const logical = parseCanonicalDocumentJson(this.schema,
        withoutBareTopLevelInlineNodes(this.schema, canonicalDocumentJson(yXmlFragmentToProsemirrorJSON(this.fragment))).json);
      const stateVector = Y.encodeStateVector(room);
      staging = new Y.Doc();
      staging.clientID = legacyRoomUpgradeClientId(`${field}\u0000${stateVector.join(",")}\u0000${
        JSON.stringify(logical.toJSON())}`);
      Y.applyUpdate(staging, Y.encodeStateAsUpdate(room));
      const target = staging;
      target.transact(() => {
        const fragment = target.getXmlFragment(field);
        fragment.delete(0, fragment.length);
        updateCanonicalYFragment(target, fragment, logical, { mapping: new Map(), isOMark: new Map() });
      });
      const update = Y.encodeStateAsUpdate(target, stateVector);
      Y.applyUpdate(room, update, this.nonHistoryOrigin);
      return true;
    } catch {
      return false;
    } finally {
      staging?.destroy();
    }
  }

  private refreshNetworkModel() {
    if (this.closed || !this.options.session.network) return;
    // A late handshake may clear only the negotiation hold, never a rejected
    // legacy/unknown document. Reconnect must re-check the server's verdict.
    if (this.issue !== null && this.issue !== "model-not-negotiated") return;
    this.setIssue(this.options.canonicalNetworkModel?.() === true ? null : "model-not-negotiated");
  }

  private setIssue(issue: PagedDocumentBridgeIssue | null, cause?: unknown) {
    if (this.issue === issue) return;
    this.issue = issue;
    // An issue stops every later publication of this editor's edits into the
    
    // editor for twenty minutes while its room stayed at the last save, and
    // nothing said why. Say which issue and what raised it (content-free).
    if (issue !== null) this.reportBridgeEvent("PagedBridgeIssue", issue, cause);
    // A failed publication or position mapping is no reason to stop for
    // good: try again shortly, with what the room took meanwhile merged in.
    if ((issue === "invalid-document" || issue === "position-unavailable") && this.established && !this.closed &&
        this.recoverAttempts < 10) {
      this.recoverAttempts += 1;
      if (this.recoverTimer !== null) clearTimeout(this.recoverTimer);
      this.recoverTimer = setTimeout(() => {
        this.recoverTimer = null;
        if (this.closed || this.issue !== issue) return;
        if (this.recoverFromIssue()) this.reconcileViewIntoRoom();
      }, 400);
    }
    queueMicrotask(() => { if (!this.closed && this.issue === issue) this.options.onIssue?.(issue); });
    this.scheduleRefresh();
  }

  private noteDerivation(base: ProseMirrorNode, tr: Transaction) {
    if (tr.doc === base || this.positions.has(tr.doc) || this.derivations.has(tr.doc)) return;
    const depth = this.positions.has(base) ? 1 : (this.derivations.get(base)?.depth ?? Infinity) + 1;
    if (depth <= MAX_POSITION_DERIVATION_DEPTH) this.derivations.set(tr.doc, { base, steps: tr.steps, docs: tr.docs, depth });
  }

  private positionMap(doc: ProseMirrorNode) {
    let mapping = this.positions.get(doc);
    if (mapping) return mapping;
    // Walk back to the nearest mapped document, then derive forward.
    const chain: Array<{ doc: ProseMirrorNode; derivation: PositionDerivation }> = [];
    let base: PagedDocumentPositionMap | undefined;
    for (let current = doc; ;) {
      const derivation = this.derivations.get(current);
      if (!derivation) break;
      chain.push({ doc: current, derivation });
      base = this.positions.get(derivation.base);
      if (base) break;
      current = derivation.base;
    }
    for (const { doc: derived } of chain) this.derivations.delete(derived);
    if (!base) {
      mapping = createPagedDocumentPositionMap(doc);
    } else {
      for (let index = chain.length - 1; index >= 0; index--) {
        const { doc: derived, derivation } = chain[index];
        base = derivePagedDocumentPositionMap(base, derivation.steps, derivation.docs, derived);
        this.positions.set(derived, base);
      }
      mapping = base;
    }
    this.positions.set(doc, mapping);
    return mapping;
  }

  private readCanonical(): boolean {
    const pending = this.canonicalReadPending;
    this.canonicalReadPending = false;
    this.adoptedUnverified = false;
    const read = this.readRoomCanonical();
    if (read) this.noteVerified(this.canonical);
    if (pending && read) this.scheduleRefresh();
    return read;
  }

  /**
   * After this view's own publication of `published`, with nothing else
   * written to the room meanwhile: take `published` as the room's document
   * instead of parsing the whole room back. Reading it back after every
   * write cost each key of a long document a full serialise-and-parse of the
   * room, and left the next key comparing every block with freshly parsed
   * nodes; the published document shares its unchanged nodes with the next
   * one, so the next comparison is by identity. The room is read back, and
   * that read is the authority, once the keys pause.
   */
  private adoptPublished(published: ProseMirrorNode): boolean {
    if (!this.established || this.issue !== null || this.foreignDuringPublish) return false;
    const pending = this.canonicalReadPending;
    this.canonicalReadPending = false;
    this.canonical = published;
    this.adoptedUnverified = true;
    if (pending) this.scheduleRefresh();
    this.scheduleCanonicalVerify();
    return true;
  }

  /** Read the room back once the keys pause (each publication restarts the
   * wait), unless something else read it meanwhile. */
  private scheduleCanonicalVerify() {
    if (this.canonicalVerifyTimer !== null) clearTimeout(this.canonicalVerifyTimer);
    this.canonicalVerifyTimer = setTimeout(() => {
      this.canonicalVerifyTimer = null;
      const run = () => {
        if (this.closed || !this.adoptedUnverified) return;
        // Nor under the pen: the read walks the room (135 ms between two
        // strokes on a 3,000-stroke note); it waits for the writing to pause.
        if (this.publishing || this.applying || inkWritingActive()) { this.scheduleCanonicalVerify(); return; }
        this.verifyAdopted();
      };
      const idle = (globalThis as { requestIdleCallback?: (callback: () => void, options?: { timeout: number }) => number })
        .requestIdleCallback;
      if (typeof idle === "function") idle(run, { timeout: 2_000 });
      else run();
    }, CANONICAL_VERIFY_PAUSE_MS);
  }

  private verifyAdopted() {
    const adopted = this.canonical, binding = this.meta;
    if (this.verifyTouchedElements(adopted)) {
      this.adoptedUnverified = false;
      return;
    }
    if (!this.readCanonical()) return;
    // The room says what this view published: keep the published nodes and
    // y-tiptap's binding to them, so the next key still compares by identity.
    // The blocks are compared: the document's own attributes are not in the
    // fragment (they have their own map), so a read has the schema's
    // defaults, and comparing the whole node refused every document with an
    // attribute set (a page size): each pause then left the next stroke a
    // deep comparison of every block and y-tiptap a binding to new nodes.
    if (this.issue === null && this.canonical !== adopted && this.canonical.content.eq(adopted.content)) {
      this.canonical = adopted;
      this.meta = binding;
      this.noteVerified(adopted);
    }
  }

  /** The room's top-level blocks, or null when it holds anything a read
   * repairs at the top (bare text, a bare inline node): read it whole. */
  private roomBlocks(): Y.XmlElement[] | null {
    const blocks: Y.XmlElement[] = [];
    for (const child of this.fragment.toArray()) {
      if (!(child instanceof Y.XmlElement) || isBareTopLevelInlineNode(this.schema, child.nodeName)) return null;
      blocks.push(child);
    }
    return blocks;
  }

  /** Every top-level block a room update changed (its own or another's). */
  private noteTouchedElements(transaction: Y.Transaction) {
    for (const type of transaction.changed.keys()) {
      let block: typeof type | null = type;
      while (block?._item && block._item.parent !== this.fragment) block = block._item.parent as typeof type | null;
      if (block instanceof Y.XmlElement && block._item?.parent === this.fragment) {
        this.touchedElements.add(block);
        this.originalBinding.delete(block);
      }
    }
  }

  /** The room's blocks, as just read into `doc`, are read back. */
  private noteVerified(doc: ProseMirrorNode) {
    this.touchedElements.clear();
    const blocks = this.roomBlocks();
    if (!blocks || blocks.length !== doc.childCount) return;
    blocks.forEach((block, index) => this.verifiedElements.set(block, doc.child(index)));
  }

  /**
   * The adopted document checked against the room by the blocks written or
   * changed since they were last read back, each read as a whole read reads
   * it; false (read the room whole) when anything else may differ: another
   * writer, a pending read, a block count that moved, a read that repairs.
   */
  private verifyTouchedElements(adopted: ProseMirrorNode): boolean {
    if (this.issue !== null || this.canonicalReadPending || !this.established) return false;
    const blocks = this.roomBlocks();
    if (!blocks || blocks.length !== adopted.childCount) return false;
    const reread: number[] = [];
    blocks.forEach((block, index) => {
      if (this.touchedElements.has(block) || this.verifiedElements.get(block) !== adopted.child(index)) reread.push(index);
    });
    if (reread.length * 2 > blocks.length) return false;
    if (reread.length > 0) {
      let read: ProseMirrorNode;
      try {
        const json = withoutUnsetPaginationRuntimeAttributes({ type: "doc", content: reread.map(index => roomBlockJson(blocks[index])) });
        read = parseCanonicalDocumentJson(this.schema, withoutBareTopLevelInlineNodes(this.schema, json).json);
      } catch {
        return false;
      }
      if (canonicalReadRepairOf(read) || read.childCount !== reread.length) return false;
      for (let at = 0; at < reread.length; at++) {
        if (!read.child(at).eq(adopted.child(reread[at]))) return false;
      }
    }
    for (const index of reread) this.verifiedElements.set(blocks[index], adopted.child(index));
    this.touchedElements.clear();
    return true;
  }

  private flushCanonicalRead() {
    if (this.canonicalReadPending && !this.closed) this.readCanonical();
  }

  private readRoomCanonical(): boolean {
    let readJson: unknown = null;
    try {
      if (this.fragment.length === 0) {
        if (this.established) { this.setIssue("invalid-document"); return false; }
        return true;
      }
      const json = withoutUnsetPaginationRuntimeAttributes(yXmlFragmentToProsemirrorJSON(this.fragment));
      readJson = json;
      const legacy = Array.isArray(json.content) && json.content.some((node: { type?: string }) => node.type === "pageBlock");
      if (legacy) {
        // Inspection/conversion is read-only. Never let y-tiptap's schema
        // repair silently delete legacy page wrappers in the shared document.
        this.canonical = parseCanonicalDocumentJson(this.schema, withoutBareTopLevelInlineNodes(this.schema, canonicalDocumentJson(json)).json);
        this.setIssue("legacy-room");
        this.scheduleLegacyRoomUpgrade();
        return true;
      }
      // Read without writing: y-tiptap's reader moves this client's text out
      // of an XmlText that follows another one, a write from inside a remote
      // update that left both views out of step with the room.
      this.canonical = parseCanonicalDocumentJson(this.schema, withoutBareTopLevelInlineNodes(this.schema, json).json);
      // Read repaired around a newer client's attributes: those stay in the
      // room through this client's edits (canonical-xml-attributes), and a
      // save carries them (withRoomOnlyAttributes). Known leftovers do not.
      if (canonicalReadRepairOf(this.canonical)?.undeclaredAttributes.length) {
        noteForeignCanonicalXmlAttributes(this.fragment, this.schema, passesThroughCanonicalRead);
      }
      this.meta = roomBinding(this.fragment, this.canonical);
      this.established = true;
      // The room was upgraded to its logical document (or re-seeded): the
      // physical-room hold is over. Negotiation holds still apply.
      if (this.issue === "legacy-room") { this.setIssue(null); this.refreshNetworkModel(); }
      this.armFirstEditWarm();
      return true;
    } catch (error) {
      // Typed, with the place it failed: a page that cannot be read never
      // opens as one empty paragraph without a word.
      const path = readJson === null ? "room" : firstUnreadableNodePath(this.schema, readJson);
      reportClientError(Object.assign(new Error(`canonical_document_unreadable at ${path}: ${error instanceof Error ? error.message.slice(0, 160) : "unreadable"}`), { name: "CanonicalDocumentUnreadable" }), { source: "window", windowType: "document" });
      this.setIssue(this.established ? "invalid-document" : "unreadable-document");
      return false;
    }
  }

  /** The room side of the write gate (the member's own permission is the
   * editor's `editable`): false while the room is still starting, or is
   * connected read-only. */
  private acceptsInput(): boolean {
    if (this.closed || this.issue !== null) return false;
    const network = this.options.session.network;
    if (network && !network.failed) {
      // A negotiated original can keep local edits while transport reconnects.
      // Initial seeding and connected read-only verdicts still deny writes;
      // durable originals need no seeding (see the option).
      if (network.connected ? !network.writable
        : (!network.ready && this.options.durableOriginals !== true) || !this.established ||
          this.options.canonicalNetworkModel?.() !== true) return false;
    }
    return true;
  }

  private mayPublish(trustedHydration = false, trustedPublication = false): boolean {
    if (!this.acceptsInput()) return false;
    // A view still being filled (fillInChunks) holds only the room's first
    // sheets: nothing it holds is the room's document, and published it
    // would be the room cut to them. Edits asked of it meanwhile are held and
    // applied once it is complete (fillEdits); no publication goes out
    
    // fill's first 212 blocks across a development reload in mid-open).
    if (this.filling !== null) return false;
    const network = this.options.session.network;
    if (trustedHydration && !this.established) return !network || (network.first === true && !network.hasState && network.writable && network.connected);
    if (trustedPublication) return true;
    return this.options.mayWrite?.() ?? (this.view?.editable !== false);
  }

  private captureTextEdit(transaction: Transaction, state: EditorState): CanonicalTextEdit | null {
    if (!this.established || this.issue || transaction.steps.length !== 1) return null;
    const step = transaction.steps[0];
    if (!(step instanceof ReplaceStep) || step.slice.openStart || step.slice.openEnd) return null;
    const from = state.doc.resolve(step.from), to = state.doc.resolve(step.to);
    if (!from.sameParent(to) || !from.parent.isTextblock) return null;
    let inserted = "", removed = "", onlyText = true;
    step.slice.content.forEach(node => { if (!node.isText) onlyText = false; else inserted += node.text; });
    state.doc.slice(step.from, step.to).content.forEach(node => { if (!node.isText) onlyText = false; else removed += node.text; });
    if (!onlyText) return null;
    try {
      const map = this.positionMap(state.doc);
      if (!map.canonical.content.eq(this.canonical.content)) return null;
      const start = map.toCanonical(step.from, 1), end = map.toCanonical(step.to, -1);
      if (start === null || end === null) return null;
      return {from: roomRelativePosition(start, this.fragment, this.meta.mapping),
        to: roomRelativePosition(end, this.fragment, this.meta.mapping), inserted, removed};
    } catch { return null; }
  }

  private applyTextEdit(edit: CanonicalTextEdit) {
    if (this.consumedTextEdits.has(edit)) return;
    this.consumedTextEdits.add(edit);
    const doc = this.options.session.ydoc;
    const from = Y.createAbsolutePositionFromRelativePosition(edit.from, doc);
    const to = Y.createAbsolutePositionFromRelativePosition(edit.to, doc);
    if (!from || !to || from.type !== to.type || !(from.type instanceof Y.XmlText) || to.index < from.index) return;
    const runs: Array<{insert: unknown}> = from.type.toDelta();
    if (runs.some(run => typeof run.insert !== "string")) return;
    const text = runs.map(run => run.insert).join("");
    if (text.slice(from.index, to.index) !== edit.removed) return;
    // Preserve the actual insertion boundary before reconciling marks/structure.
    // A whole-string prefix diff can borrow another author's matching prefix,
    // making personal Undo remove part of that peer's subsequently typed text.
    if (to.index > from.index) from.type.delete(from.index, to.index - from.index);
    if (edit.inserted) from.type.insert(from.index, edit.inserted);
  }

  private relativeRoomPosition(position: number): Y.RelativePosition {
    return leftBoundPosition(this.options.session.ydoc, roomRelativePosition(position, this.fragment, this.meta.mapping));
  }

  /** The last selection captured: a publication captures the caret before its
   * edit and after it, and the "before" of one is the "after" of the one
   * before (the same editor state against the same room document). Each
   * capture walked the room fragment to the caret, twice a stroke on a long
   * note. */
  private lastCapture: {
    doc: ProseMirrorNode; selection: Selection; cells: boolean; canonical: ProseMirrorNode; captured: RelativeSelection | null;
  } | null = null;

  private captureSelection(state: EditorState, cells = true): RelativeSelection | null {
    // By what a capture reads (the document and its selection), not by the
    // state object: the editor's own follow-ups between two publications
    // (a landed stroke placed, strokes put to rest) made new states of the
    // same document and caret.
    const last = this.lastCapture;
    if (last && last.doc === state.doc && last.cells === cells && last.canonical === this.canonical &&
        last.selection.eq(state.selection) && this.established && this.issue === null) return last.captured;
    const captured = this.captureSelectionFresh(state, cells);
    this.lastCapture = captured
      ? { doc: state.doc, selection: state.selection, cells, canonical: this.canonical, captured }
      : null;
    return captured;
  }

  private captureSelectionFresh(state: EditorState, cells: boolean): RelativeSelection | null {
    if (!this.established || this.issue !== null) return null;
    try {
      const map = this.positionMap(state.doc), selection = state.selection;
      // A cell selection's own anchor/head are its first cell's text; its
      // cells are kept by the positions before its anchor and head cells.
      const cell = cells && selection instanceof CellSelection;
      const anchorAt = cell ? selection.$anchorCell.pos : selection.anchor;
      const headAt = cell ? selection.$headCell.pos : selection.head;
      const anchor = selection instanceof AllSelection ? 0 : map.toCanonical(anchorAt, 1);
      const head = selection instanceof AllSelection ? this.canonical.content.size : map.toCanonical(headAt, cell ? 1 : -1);
      if (anchor === null || head === null) return null;
      const all = selection instanceof AllSelection;
      return {
        kind: all ? "all" : selection instanceof NodeSelection ? "node" : cell ? "cell" : "text",
        anchor: this.relativeRoomPosition(all ? anchor : acrossDifference(map.canonical, this.canonical, anchor)),
        head: this.relativeRoomPosition(all ? head : acrossDifference(map.canonical, this.canonical, head)),
        anchorHint: anchorAt, headHint: headAt,
        ...(selection instanceof NodeSelection ? { nodeType: selection.node.type.name } : {}),
      };
    } catch { return null; }
  }

  private selectionFor(doc: ProseMirrorNode, relative: RelativeSelection | null): Selection | null {
    if (!relative) return null;
    if (relative.kind === "all") return new AllSelection(doc);
    try {
      const a = relativePositionToAbsolutePosition(this.options.session.ydoc, this.fragment, relative.anchor, this.meta.mapping);
      const h = relativePositionToAbsolutePosition(this.options.session.ydoc, this.fragment, relative.head, this.meta.mapping);
      if (a === null || h === null) return null;
      const map = this.positionMap(doc);
      const anchor = map.toLocal(acrossDifference(this.canonical, map.canonical, a), 1, relative.anchorHint);
      const head = map.toLocal(acrossDifference(this.canonical, map.canonical, h), relative.kind === "cell" ? 1 : -1, relative.headHint);
      if (anchor === null || head === null) return null;
      if (relative.kind === "cell") {
        const $anchor = doc.resolve(anchor), $head = doc.resolve(head);
        const isCell = (role: unknown) => role === "cell" || role === "header_cell";
        if (isCell($anchor.nodeAfter?.type.spec.tableRole) && isCell($head.nodeAfter?.type.spec.tableRole) &&
          $anchor.node(-1) === $head.node(-1)) return CellSelection.create(doc, anchor, head);
      }
      const node = doc.nodeAt(anchor);
      if (relative.kind === "node" && node && node.type.name === relative.nodeType && NodeSelection.isSelectable(node)) return NodeSelection.create(doc, anchor);
      return TextSelection.between(doc.resolve(anchor), doc.resolve(head));
    } catch { return null; }
  }

  /** The caret is captured once, when a run of Y transactions starts and the
   * view still matches the room, and restored after each of them. Captured
   * per transaction, a write made while a remote update's observers ran (the
   * autosave recovery mirror writes the room there) re-read the old caret
   * position against the updated room: the caret attached to the other
   * writer's character, so two people typing at one spot interleaved their
   * text and each new character landed before its own previous one. */
  private onBeforeYTransactions = () => {
    // A view still being filled has no caret of the reader's to keep, and
    // capturing one maps the partial document. Nor does a view the reader has
    // not been in yet: its caret is where the editor put it, and mapping the
    // whole document to keep it was a 0.8 s task as the room's backlog
    // arrived just after the 298-page book was filled. Its selection maps
    // through the update as ProseMirror maps any selection.
    if (!this.readerWasHere && this.view?.hasFocus()) this.readerWasHere = true;
    if (!this.closed && !this.publishing && this.view && this.filling === null && this.readerWasHere) this.beforeRemote = this.captureSelection(this.view.state);
    if (!this.closed && !this.publishing && !this.compositionHold && this.view && isImeCompositionActive(this.view)) {
      this.compositionAnchors = this.captureCompositionAnchors(this.view);
    }
  };
  private onAfterYTransactions = () => { this.beforeRemote = null; this.compositionAnchors = null; };
  private onYUpdate = (_bytes: Uint8Array, _origin: unknown, _doc: Y.Doc, transaction: Y.Transaction) => {
    if (!this.publishing && [...transaction.changedParentTypes.keys()].some(type => Object.is(type, this.fragment))) this.originalNormalized = false;
    this.refreshNetworkModel();
    if (this.closed) return;
    this.noteTouchedElements(transaction);
    // A local publication reads the room back itself right after writing it.
    if (this.publishing) {
      const origin = transaction.origin;
      const foreign = origin !== this.localOrigin && origin !== this.nonHistoryOrigin && origin !== this.seedOrigin;
      if (foreign) this.foreignDuringPublish = true;
      // Its own write of the document's attributes alone (every gesture
      // publishes them) leaves the room's blocks as they were: reading all of
      // them back for it was a whole-room parse on a stroke of a long note.
      if (foreign || [...transaction.changedParentTypes.keys()].some(type => Object.is(type, this.fragment))) {
        this.canonicalReadPending = true;
      }
      return;
    }
    if (!this.readCanonical()) return;
    // An input method composing in the view: the update waits for it.
    if (this.compositionHold || (this.historyTarget === undefined && this.beginCompositionHold())) return;
    if (this.view) this.keepUnpublishedLocal(this.view);
    this.applyCanonical(this.historyTarget !== undefined ? this.historyTarget : this.beforeRemote);
    this.scheduleRefresh();
    if (!transaction.local) this.settleConcurrentSplits();
  };

  /* Enter moves a block's tail into a new block: the tail is deleted from the
   * old block and a COPY of it starts the new one. Two people who press Enter
   * in one paragraph at the same moment each make such a copy, so after the
   * merge the shared tail stood in the document twice ("abc" / "def" / "ef").
   *
   * Each view remembers the blocks it split off for a minute. When a peer's
   * block arrives that was inserted after the same block, concurrently, and
   * whose text shares its end with ours, the two are copies of one tail. The
   * one that comes first in the document gives the shared part up, so both
   * Enters keep their paragraph and every character stands once
   * ("abc" / "d" / "ef"). Both views reach the same verdict from the merged
   * order; only the owner of the first block writes. */
  private onAfterYTransaction = (transaction: Y.Transaction) => {
    if (this.closed || transaction.origin !== this.localOrigin) return;
    try { this.noteSplit(transaction); } catch { /* bookkeeping only */ }
  };

  private noteSplit(transaction: Y.Transaction) {
    const doc = this.options.session.ydoc, client = doc.clientID;
    const before = transaction.beforeState.get(client) ?? 0, after = transaction.afterState.get(client) ?? 0;
    if (after <= before) return;
    let removed = 0;
    Y.iterateDeletedStructs(transaction, transaction.deleteSet, struct => {
      if (struct instanceof Y.Item && struct.parent instanceof Y.XmlText && struct.content instanceof Y.ContentString) removed += struct.length;
    });
    if (removed === 0) return;
    const created: Y.XmlElement[] = [];
    for (const struct of doc.store.clients.get(client) ?? []) {
      if (struct.id.clock < before || !(struct instanceof Y.Item) || struct.deleted) continue;
      const type = struct.content instanceof Y.ContentType ? struct.content.type : null;
      if (!(type instanceof Y.XmlElement)) continue;
      const parentItem = (struct.parent as Y.AbstractType<unknown>)._item;
      // Top-most new elements only: their parent existed before this gesture.
      if (parentItem && parentItem.id.client === client && parentItem.id.clock >= before) continue;
      created.push(type);
    }
    if (created.length !== 1) return;
    const text = plainText(created[0]);
    if (text.length === 0 || text.length !== removed) return;
    const now = Date.now();
    this.splits = this.splits.filter(split => now - split.at < SPLIT_RECORD_TTL_MS).slice(-(SPLIT_RECORD_LIMIT - 1));
    const texts = created[0].toArray().filter((child): child is Y.XmlText => child instanceof Y.XmlText);
    this.splits.push({
      element: created[0], text, known: new Map(transaction.beforeState), at: now, settled: new WeakSet(),
      start: texts.length === 1 ? Y.createRelativePositionFromTypeIndex(texts[0], 0) : null,
    });
  }

  private settleConcurrentSplits() {
    if (this.splits.length === 0 || this.closed || this.issue !== null || !this.mayPublish()) return;
    const doc = this.options.session.ydoc, now = Date.now();
    this.splits = this.splits.filter(split => now - split.at < SPLIT_RECORD_TTL_MS &&
      split.element._item !== null && !split.element._item.deleted);
    const cuts: Array<{ text: Y.XmlText; from: number; length: number }> = [];
    for (const split of this.splits) {
      const mine = split.element._item!;
      const parent = split.element.parent as Y.XmlFragment | Y.XmlElement | null;
      if (!parent || mine.origin === null) continue;
      const siblings = parent.toArray();
      const myIndex = siblings.indexOf(split.element);
      siblings.forEach((sibling, index) => {
        if (!(sibling instanceof Y.XmlElement) || sibling === split.element || split.settled.has(sibling)) return;
        const item = sibling._item;
        if (!item || item.deleted || item.id.client === doc.clientID || sibling.nodeName !== split.element.nodeName) return;
        // Concurrent (this view had not seen it when it split) and placed after the same block.
        if (item.id.clock < (split.known.get(item.id.client) ?? 0)) return;
        if (item.origin === null || !Y.compareIDs(item.origin, mine.origin)) return;
        const theirs = plainText(sibling);
        if (theirs.length === 0) return;
        // The peer may have typed in its block before this view heard of it.
        const shared = split.text.endsWith(theirs) ? theirs : theirs.includes(split.text) ? split.text : null;
        if (shared === null) return;
        split.settled.add(sibling);
        if (myIndex > index) return; // the peer's block comes first: it gives the shared part up
        const at = split.start && Y.createAbsolutePositionFromRelativePosition(split.start, doc);
        if (!at || !(at.type instanceof Y.XmlText) || at.type.parent !== split.element) return;
        // Unchanged since the split, wherever later typing has pushed it.
        if (plainText(at.type).slice(at.index, at.index + split.text.length) !== split.text) return;
        cuts.push({ text: at.type, from: at.index + split.text.length - shared.length, length: shared.length });
        split.text = split.text.slice(0, split.text.length - shared.length);
      });
    }
    if (cuts.length === 0) return;
    doc.transact(() => { for (const cut of cuts) cut.text.delete(cut.from, cut.length); }, this.nonHistoryOrigin);
  }

  private onLocalUpdate(view: EditorView, previous: EditorState) {
    if (this.closed) return;
    this.lastViewState = view.state;
    if (this.applying) { this.checkLayoutOnlyUpdates(previous.doc); return; }
    // The view no longer shows the room's document: what is composed goes
    // into the room when the composition settles, at its place there.
    if (this.compositionHold) return;
    const state = ySyncPluginKey.getState(view.state) as BridgeState;
    // A transaction that moved no caret (a refresh, a plugin's meta) only
    // re-publishes in the background (`publishSelection`): the first caret
    // too, which once went out at once from the first such transaction after
    // the fill and so mapped the whole book (0.5 s in one task on a
    // 693k-character book) for a room with nobody else in it to see it. With
    // a peer in the room the background path publishes at once, as before.
    // An update that changed no document (a projection's decorations, a
    // caret move) leaves a pending layout-only run as it was: proving it now
    // mapped the whole paged document to the canonical one on every such
    // update, several seconds of a 298-page book's open in all.
    if (!state.changed) {
      if (view.state.doc !== previous.doc) this.checkLayoutOnlyUpdates(previous.doc);
      const caretStayed = view.state.selection.eq(previous.selection);
      // The ledger's own stamps after each page move, with a peer in the room:
      // a caret that did not move has the room position it had (layout never
      // moves one), and asking for it again mapped the whole paged document
      // after every background step (a fifth of the thread on a 2,747-page
      // document). It goes out when the layout is done.
      if (caretStayed && this.paginationBusy() && !this.positions.has(view.state.doc)) {
        this.deferSelectionPublish();
        return;
      }
      this.publishSelection(caretStayed);
      return;
    }
    // An edit of the author's own: the layout kept from here on is not one of
    // the content as opened (receiptForOpenedContent).
    if (state.historical || state.authoredFrom !== null) forgetOpenedProjectionContent(view.dom);
    if (state.layoutOnly) {
      // Pagination moved blocks between sheets: the shared content cannot
      // have changed, and proving it maps the whole paged document to the
      // canonical one (~1.5 s per pass on a 300-page book, hundreds of passes
      // while it is laid out in the background). The proof waits for a
      // pause and covers the whole run of layout-only updates at once.
      this.deferLayoutOnlyCheck(previous.doc, view.state.doc);
      return;
    }
    if (this.established && !state.historical && !state.trustedPublication && !state.sceneChange &&
        !state.savedWithoutHistory) {
      // A view-only batch: what the editor derived by itself (a page-ledger
      // pass with the identities and continuation attributes normalizers
      // appended to it). It is never published (below), so mapping the paged
      // document to the canonical one only to find that out cost a whole
      // document mapping per ledger pass. The caret's room position follows
      // when the run pauses.
      // A layout-only run before it is checked with the others at the pause.
      this.closeLayoutOnlyRun(previous.doc);
      this.deferSelectionPublish();
      return;
    }
    this.checkLayoutOnlyUpdates(previous.doc);
    try {
      const projected = this.positionMap(view.state.doc).canonical;
      const before = this.positionMap(previous.doc).canonical;
      const joinsStep = state.historical && this.joinsAwaitingStep(state.followUp);
      if (state.historical && !joinsStep) this.beginAuthoredStep(state.closesHistory && this.established, state.awaitsFollowUp);
      if (this.established && projected.content.eq(this.canonical.content)) {
        this.agreedCanonical = this.canonical;
        this.publishSelection();
        return;
      }
      // Rendering/normalization of an existing original is view-only. Only
      // an authored gesture, an explicit trusted publication or a change kept
      // out of Undo that is still the document's (a shape's `text` mirroring
      // its story, WUA 220) can replace it.
      if (this.established && !state.historical && !state.trustedPublication && !state.savedWithoutHistory) {
        // Whiteboard's explicit representation upgrade is semantically equal
        // and must precede the user's undoable gesture. Keep its baseline
        // objects outside personal history without authorizing other changes.
        const sceneRepresentationOnly = state.sceneChange && pageDocumentJsonEquals(
          serializeWhiteboardScenes(before.toJSON()), serializeWhiteboardScenes(projected.toJSON()));
        if (!sceneRepresentationOnly) { this.publishSelection(); return; }
      }
      if (!this.mayPublish(state.trustedHydration, state.trustedPublication)) {
        // An authored gesture the room did not take (it is published with
        // the next gesture the room takes, or never if the refusal lasts).
        if (state.historical) this.reportBridgeEvent("PagedBridgeGestureUnpublished", this.publishRefusal());
        return;
      }
      let seed = !this.established;
      const selectionBefore = this.captureSelection(previous);
      this.publishing = true;
      this.foreignDuringPublish = false;
      try {
        if (seed && state.historical && !state.replacesDocument && before.childCount > 0) {
          // The first edit of a room nobody has seeded yet (a new, empty
          // document) is a gesture like any other. Seeding the room WITH it
          // put the gesture outside personal history, so the first character
          // typed into a new document could never be undone. The document as
          // it stood before the gesture is the seed; the gesture then
          // publishes as the undoable step it is. Content that replaces the
          // whole document is the seed itself, as before.
          this.options.session.ydoc.transact(() =>
            updateCanonicalYFragment(this.options.session.ydoc, this.fragment, before, this.meta), this.seedOrigin);
          if (!this.readCanonical()) return;
          seed = !this.established;
        }
        if (!seed && state.historical && !this.materializeOriginal()) return;
        if (!seed && state.historical && this.established && !before.content.eq(this.canonical.content)) {
          // Local normalization that no gesture authored (identity back-fill
          // on nested list items and paragraphs, provenance clearing) is
          // view-only until now, so the pre-gesture document already differs
          // from the shared one. Publish that difference outside history
          // first: folded into the gesture it would be undone with it, the
          // normalizer would restore it at once, and the editor would then
          // differ from the shared document forever (every later save
          // refused as an unmirrorable capture).
          let wrote = false;
          const noteWrite = () => { wrote = true; };
          this.options.session.ydoc.on("update", noteWrite);
          try {
            this.options.session.ydoc.transact(() =>
              updateCanonicalYFragment(this.options.session.ydoc, this.fragment, before, this.meta), this.nonHistoryOrigin);
          } finally {
            this.options.session.ydoc.off("update", noteWrite);
          }
          if (wrote && !this.readCanonical()) return;
        }
        this.localHistory = { before: selectionBefore, after: null, beforeDoc: this.canonical, afterDoc: null };
        // A follow-up joins its step's Undo item however long its frame took
        // (the undo manager folds a change into the last item only within its
        // capture timeout); the step itself stopped capturing before it.
        const captureTimeout = this.undoManager.captureTimeout;
        if (joinsStep && !seed) this.undoManager.captureTimeout = Number.POSITIVE_INFINITY;
        else if (!seed && state.historical && state.authoredFrom !== null && this.lastAuthoredTo !== null) {
          // One typing run is one Undo step however long each key's own
          // publication took: the undo manager measures its timeout between
          // the ends of its transactions, and on a 3-million-character page
          // each key's took longer than that, so one Cmd+Z took back one
          // character (ab27 sac29). The gap is the one between the keys
          // (lib/editor/typing-run-time); a closed group stays closed
          // (stopCapturing above).
          const gap = state.authoredFrom - this.lastAuthoredTo;
          this.undoManager.captureTimeout = gap >= 0 && gap < TYPING_RUN_GAP_MS ? Number.POSITIVE_INFINITY : 0;
        }
        try {
          this.options.session.ydoc.transact(() => {
            if (state.textEdit) this.applyTextEdit(state.textEdit);
            updateCanonicalYFragment(this.options.session.ydoc, this.fragment, projected, this.meta);
          }, seed ? this.seedOrigin : state.historical ? this.localOrigin : this.nonHistoryOrigin);
        } finally {
          this.undoManager.captureTimeout = captureTimeout;
        }
        if (!seed && state.historical) this.lastAuthoredTo = state.authoredTo;
        if (seed || !this.adoptPublished(projected)) this.readCanonical();
        this.agreedCanonical = this.canonical;
        // The caret did not move in the room: the capture before is the one after.
        this.localHistory.after = state.keepsCaret && selectionBefore && !seed
          ? selectionBefore : this.captureSelection(view.state);
        this.localHistory.afterDoc = this.canonical;
        // The follow-up is in the step's Undo item: the step is complete.
        if (joinsStep) this.closeOpenStep();
      } finally { this.publishing = false; this.localHistory = null; this.flushCanonicalRead(); }
      this.publishSelection();
    } catch (error) { this.setIssue("invalid-document", error); }
  }

  private readonly reportedBridgeEvents = new Set<string>();

  /** One content-free client report per event and code for this bridge. */
  private reportBridgeEvent(name: string, code: string, cause?: unknown) {
    const key = `${name}:${code}`;
    if (this.reportedBridgeEvents.has(key)) return;
    this.reportedBridgeEvents.add(key);
    const detail = cause instanceof Error ? ` ${cause.name}: ${cause.message.slice(0, 160)}` : "";
    try {
      reportClientError(Object.assign(new Error(`${code}${detail}`), { name }), { source: "window", windowType: "document" });
    } catch { /* diagnostics never break the bridge */ }
  }

  private deferSelectionPublish() {
    if (this.selectionPublishTimer !== null) clearTimeout(this.selectionPublishTimer);
    const fire = () => {
      this.selectionPublishTimer = null;
      if (this.paginationBusy()) { this.deferSelectionPublish(); return; }
      // The runs kept through a background layout are proved one task each.
      const run = this.closedLayoutOnlyRuns.shift();
      if (run) this.checkLayoutOnlyRun(run);
      if (this.closedLayoutOnlyRuns.length > 0) { this.selectionPublishTimer = setTimeout(fire, 0); return; }
      this.publishSelection(true);
    };
    this.selectionPublishTimer = setTimeout(fire, LAYOUT_ONLY_CHECK_PAUSE_MS);
  }

  /**
   * The page ledger is still laying the document out (its first settle, or
   * the rest of a long document in the background). Its passes pause for a
   * frame or two between pages, long enough for the layout-only proof's
   * timer, and each proof maps the whole paged document to the canonical
   * one: a 298-page book's open ran it every few passes, seconds of the main
   * thread each time, while the reader tried to use the rest of the app. The
   * proof (and the caret publication that maps the document too) waits until
   * the layout is done; an authored step or a remote update still runs it
   * first, as before.
   */
  private paginationBusy(): boolean {
    const dom = this.view?.dom;
    // "data-pagination-initial-settling": pageRebalanceExtension's
    // INITIAL_PAGINATION_SETTLING_ATTR (a module this one does not load).
    return Boolean(dom && (dom.hasAttribute("data-pagination-initial-settling") ||
      dom.hasAttribute(PAGE_LEDGER_BACKGROUND_ATTR) || dom.hasAttribute(PAGE_LEDGER_DEFERRED_ATTR)));
  }

  /** End the pending layout-only run at `end` without checking it yet: a
   * view-only batch follows, and the run is checked with the others when the
   * updates pause (a long run of such batches checks the oldest at once, so
   * the documents kept for the check stay few). */
  private closeLayoutOnlyRun(end: ProseMirrorNode) {
    const pending = this.layoutOnlyCheck;
    if (!pending) return;
    this.layoutOnlyCheck = null;
    if (pending.timer !== null) clearTimeout(pending.timer);
    this.closedLayoutOnlyRuns.push({ base: pending.base, end: end === pending.base ? end : pending.latest });
    const kept = this.paginationBusy() ? MAX_CLOSED_LAYOUT_ONLY_RUNS_WHILE_BUSY : MAX_CLOSED_LAYOUT_ONLY_RUNS;
    while (this.closedLayoutOnlyRuns.length > kept) this.checkLayoutOnlyRun(this.closedLayoutOnlyRuns.shift()!);
  }

  private checkClosedLayoutOnlyRuns() {
    const runs = this.closedLayoutOnlyRuns.splice(0);
    for (const run of runs) this.checkLayoutOnlyRun(run);
  }

  private checkLayoutOnlyRun(run: { base: ProseMirrorNode; end: ProseMirrorNode }) {
    if (this.closed) return;
    try {
      if (!this.positionMap(run.end).canonical.content.eq(this.positionMap(run.base).canonical.content)) this.setIssue("invalid-document");
    } catch { this.setIssue("invalid-document"); }
  }

  private deferLayoutOnlyCheck(base: ProseMirrorNode, latest: ProseMirrorNode) {
    const pending = this.layoutOnlyCheck ?? { base, latest, timer: null };
    pending.latest = latest;
    if (pending.timer !== null) clearTimeout(pending.timer);
    const fire = () => {
      pending.timer = null;
      if (this.layoutOnlyCheck !== pending) return;
      if (this.paginationBusy()) { pending.timer = setTimeout(fire, LAYOUT_ONLY_CHECK_PAUSE_MS); return; }
      // Runs kept through a background layout first, one task each.
      const run = this.closedLayoutOnlyRuns.shift();
      if (run) {
        this.checkLayoutOnlyRun(run);
        if (this.closedLayoutOnlyRuns.length > 0) { pending.timer = setTimeout(fire, 0); return; }
      }
      this.checkLayoutOnlyUpdates(pending.latest);
      this.publishSelection(true);
    };
    pending.timer = setTimeout(fire, LAYOUT_ONLY_CHECK_PAUSE_MS);
    this.layoutOnlyCheck = pending;
  }

  /** Prove that the layout-only updates since the last check, ending with
   * `latest`, left the shared content as it was; before anything else
   * changes the document (an authored step, a remote update). */
  private checkLayoutOnlyUpdates(latest: ProseMirrorNode) {
    this.checkClosedLayoutOnlyRuns();
    const pending = this.layoutOnlyCheck;
    if (!pending) return;
    this.layoutOnlyCheck = null;
    if (pending.timer !== null) clearTimeout(pending.timer);
    // The run ends with the document the next change started from.
    this.checkLayoutOnlyRun({ base: pending.base, end: latest === pending.latest || latest === pending.base ? latest : pending.latest });
  }

  /** A command that closes the history group (a table's Split, a width or
   * height commit, a Section Break, an ink stroke) is one Undo step of its
   * own, as it is in the ProseMirror history: the undo manager otherwise
   * folds everything within its capture timeout (typing just before or
   * after) into it. Its step stays open to the rest of its own dispatch, so
   * the document attributes it changed (a Section Break's section layout)
   * are undone with it rather than by a Cmd+Z of their own. */
  private beginAuthoredStep(closesHistory: boolean, awaitsFollowUp = false) {
    if (this.openStep) { this.openStep = null; this.undoManager.stopCapturing(); }
    if (!closesHistory) return;
    this.undoManager.stopCapturing();
    const step = { awaitsFollowUp, settled: false };
    this.openStep = step;
    const close = () => {
      if (this.openStep !== step) return;
      this.openStep = null;
      if (!this.closed) this.undoManager.stopCapturing();
    };
    // A step awaiting its follow-up stays open for it after its dispatch;
    // any other authored step closes it first (above), so nothing else is
    // folded into it.
    if (awaitsFollowUp) {
      queueMicrotask(() => {
        step.settled = true;
      });
      setTimeout(close, PAGED_FOLLOW_UP_WAIT_MS);
    } else {
      queueMicrotask(close);
    }
  }

  /** A derived follow-up while its step is still open for it: published into that step's Undo item. */
  private joinsAwaitingStep(followUp: boolean): boolean {
    return followUp && this.established && this.openStep?.awaitsFollowUp === true;
  }

  private closeOpenStep() {
    if (!this.openStep) return;
    this.openStep = null;
    if (!this.closed) this.undoManager.stopCapturing();
  }

  private fillHeldForReceipt = false;

  /** The room positions of the composing view's positions near its caret,
   * while the view shows the room's document; null otherwise. */
  private captureCompositionAnchors(view: EditorView): PagedDocumentBridge["compositionAnchors"] {
    if (!this.established || this.issue !== null || this.filling !== null) return null;
    const { selection, doc } = view.state;
    if (!(selection instanceof TextSelection)) return null;
    const $head = selection.$head;
    if (!$head.parent.isTextblock) return null;
    try {
      const map = this.positionMap(doc);
      if (!map.canonical.content.eq(this.canonical.content)) return null;
      const ydoc = this.options.session.ydoc;
      const anchors: CompositionHold["anchors"] = new Map();
      const from = Math.max($head.start(), Math.min(selection.from, selection.head) - COMPOSITION_HOLD_BEFORE);
      const to = Math.min($head.end(), Math.max(selection.to, selection.head) + COMPOSITION_HOLD_AFTER);
      for (let position = from; position <= to; position += 1) {
        const canonical = map.toCanonical(position, 1);
        if (canonical === null) continue;
        const absolute = Y.createAbsolutePositionFromRelativePosition(
          roomRelativePosition(canonical, this.fragment, this.meta.mapping), ydoc);
        if (!absolute || !(absolute.type instanceof Y.XmlText)) continue;
        anchors.set(position, {
          right: Y.createRelativePositionFromTypeIndex(absolute.type, absolute.index, 0),
          left: Y.createRelativePositionFromTypeIndex(absolute.type, absolute.index, -1),
        });
      }
      const head = map.toCanonical(selection.head, 1);
      if (head === null || anchors.size === 0) return null;
      const $canonical = map.canonical.resolve(head);
      return { doc, anchors, canonical: map.canonical, block: { from: $canonical.start(), to: $canonical.end() } };
    } catch { return null; }
  }

  /** Begin holding the room's updates from the composing view (the update
   * just read is the first held); false when the view is not composing or
   * did not match the room when the update began. */
  private beginCompositionHold(): boolean {
    const view = this.view, captured = this.compositionAnchors;
    this.compositionAnchors = null;
    if (!view || !captured || captured.doc !== view.state.doc || !isImeCompositionActive(view)) return false;
    // Only an update of the paragraph being composed in: a redraw elsewhere
    // leaves the composition's text node alone (ProseMirror keeps it).
    const start = captured.canonical.content.findDiffStart(this.canonical.content);
    const end = captured.canonical.content.findDiffEnd(this.canonical.content);
    if (start === null || !end) return false;
    if (Math.max(start, end.a) < captured.block.from || Math.min(start, end.a) > captured.block.to) return false;
    let poll: ReturnType<typeof setInterval> | null = null;
    let unsubscribe: (() => void) | null = null;
    const hold: CompositionHold = {
      view, held: captured.doc, anchors: captured.anchors, region: null, complex: false,
      release: () => {
        unsubscribe?.(); unsubscribe = null;
        if (poll !== null) { clearInterval(poll); poll = null; }
      },
    };
    this.compositionHold = hold;
    const settle = () => {
      if (this.compositionHold !== hold || view.composing || isImeCompositionActive(view)) return;
      this.releaseCompositionHold(hold);
    };
    unsubscribe = onImeCompositionSettled(view, () => queueMicrotask(settle));
    // A view without the guard's settle signal, or one that never sends it.
    poll = setInterval(settle, COMPOSITION_HOLD_POLL_MS);
    return true;
  }

  /** A held view's own change: the composition's text, else marked complex. */
  private noteHeldComposition(tr: Transaction) {
    const hold = this.compositionHold;
    if (!hold || hold.complex) return;
    for (let index = 0; index < tr.steps.length; index += 1) {
      const step = tr.steps[index]!, doc = tr.docs[index]!;
      const range = step.getMap();
      let grows = 0;
      range.forEach((oldStart, oldEnd, newStart, newEnd) => { grows += (newEnd - newStart) - (oldEnd - oldStart); });
      const text = step instanceof ReplaceStep && !step.slice.openStart && !step.slice.openEnd && (() => {
        const $from = doc.resolve(step.from), $to = doc.resolve(step.to);
        if (!$from.sameParent($to) || !$from.parent.isTextblock) return false;
        let only = true;
        step.slice.content.forEach((node) => { if (!node.isText) only = false; });
        // Inside one paragraph, only characters are one position each.
        return only && doc.textBetween(step.from, step.to).length === step.to - step.from;
      })();
      if (!text) {
        // An attribute or a mark (an identity or an anchor a plugin derived,
        // the room's page setup) moves no position: derived again, or the
        // room's own, once the room's document is drawn.
        if (grows === 0) continue;
        hold.complex = true;
        return;
      }
      const { from, to } = step as ReplaceStep;
      const size = (step as ReplaceStep).slice.size;
      const region = hold.region;
      if (!region) {
        hold.region = { heldFrom: from, heldTo: to, from, to: from + size };
        continue;
      }
      // Each change of a composition is at the text it composed so far.
      if (from > region.to || to < region.from) { hold.complex = true; return; }
      const start = Math.min(region.from, from), end = Math.max(region.to, to);
      hold.region = {
        heldFrom: region.heldFrom - (region.from - start),
        heldTo: region.heldTo + (end - region.to),
        from: start,
        to: start + (end - start) - (to - from) + size,
      };
    }
  }

  /**
   * The composition settled: the view takes the room's document, and what
   * was composed meanwhile goes in at the place the room carried it to, as
   * this writer's own edit (published by the ordinary path, with its marks).
   */
  private releaseCompositionHold(hold: CompositionHold) {
    if (this.compositionHold !== hold) return;
    hold.release();
    this.compositionHold = null;
    const view = this.view;
    if (this.closed || !view || view !== hold.view) return;
    const composed = view.state.doc;
    const plan = hold.complex ? null : this.composedEdit(hold, composed, view.state.selection.head);
    if (!plan) {
      // Not one run of text in one paragraph: the view's document and the
      // room's are merged, both sides' blocks kept.
      if (!this.reconcileViewIntoRoom()) this.applyCanonical(null);
      return;
    }
    this.applyCanonical(null);
    if (plan === "nothing") return;
    try {
      const ydoc = this.options.session.ydoc;
      const a = relativePositionToAbsolutePosition(ydoc, this.fragment, plan.from, this.meta.mapping);
      const b = relativePositionToAbsolutePosition(ydoc, this.fragment, plan.to, this.meta.mapping);
      const map = this.positionMap(view.state.doc);
      let localFrom = a === null ? null : map.toLocal(acrossDifference(this.canonical, map.canonical, a), 1);
      const localTo = b === null ? null : map.toLocal(acrossDifference(this.canonical, map.canonical, b), -1);
      if (localFrom === null || localTo === null || localTo < localFrom) {
        this.reportBridgeEvent("PagedBridgeCompositionLost", "unplaced");
        return;
      }
      // What the composition replaced was changed by another writer
      // meanwhile: theirs stays, the composed text goes in after it.
      if (view.state.doc.textBetween(localFrom, localTo) !== plan.removed) localFrom = localTo;
      const tr = view.state.tr;
      if (plan.inserted.length > 0) tr.replaceWith(localFrom, localTo, Fragment.from(plan.inserted));
      else if (localTo > localFrom) tr.delete(localFrom, localTo);
      const end = localFrom + plan.inserted.reduce((sum, node) => sum + node.nodeSize, 0);
      tr.setSelection(TextSelection.near(tr.doc.resolve(Math.max(0, Math.min(tr.doc.content.size, end + plan.caret))), -1));
      if (tr.docChanged) view.dispatch(tr);
    } catch (error) {
      this.reportBridgeEvent("PagedBridgeCompositionLost", "exception", error);
    }
  }

  /** What a held composition changed, as a room edit: the room positions of
   * the text it replaced, that text, the text composed (with its marks), and
   * the caret from its end; "nothing" when it changed nothing, null when it
   * is not one run of text in one paragraph near the caret. */
  private composedEdit(hold: CompositionHold, composed: ProseMirrorNode, head: number): {
    from: Y.RelativePosition; to: Y.RelativePosition; removed: string; inserted: ProseMirrorNode[]; caret: number;
  } | "nothing" | null {
    const region = hold.region;
    if (!region) return "nothing";
    try {
      const $start = composed.resolve(region.from), $end = composed.resolve(region.to);
      const $heldStart = hold.held.resolve(region.heldFrom), $heldEnd = hold.held.resolve(region.heldTo);
      if (!$start.sameParent($end) || !$start.parent.isTextblock || !$heldStart.sameParent($heldEnd)) return null;
      const start = hold.anchors.get(region.heldFrom), end = hold.anchors.get(region.heldTo);
      if (!start || !end) return null;
      const replaces = region.heldTo > region.heldFrom;
      const inserted: ProseMirrorNode[] = [];
      composed.nodesBetween(region.from, region.to, (node, position) => {
        if (node.isText) {
          const piece = node.cut(Math.max(0, region.from - position), Math.min(node.nodeSize, region.to - position));
          if (piece.nodeSize > 0) inserted.push(piece);
        }
        return !node.isText;
      });
      return {
        // A replaced run is bound to its own first and last characters; an
        // insertion follows the character before it, this writer's own.
        from: replaces ? start.right : start.left,
        to: replaces ? end.left : start.left,
        removed: hold.held.textBetween(region.heldFrom, region.heldTo),
        inserted,
        caret: Math.min(0, head - region.to),
      };
    } catch { return null; }
  }

  private applyCanonical(selection: RelativeSelection | null) {
    const view = this.view;
    if (this.closed || !view || this.canonical.childCount === 0) return;
    // Drawn when the composition settles (`releaseCompositionHold`).
    if (this.compositionHold) return;
    if (this.filling !== null) {
      // The fill in progress shows this very room document.
      if (this.filling.canonical === this.canonical) return;
      // The room changed under it: the rest of the document it was filling
      // goes in at once, the change is applied as an ordinary difference
      // from that, then the edits made meanwhile.
      this.finishFillAtOnce(view);
      this.scheduleRefresh();
      queueMicrotask(() => this.flushFillEdits());
    }
    const confirmed = this.confirmedPair;
    if (confirmed && confirmed.doc === view.state.doc && confirmed.canonical === this.canonical) return;
    try {
      const map = this.positionMap(view.state.doc);
      if (map.canonical.content.eq(this.canonical.content)) {
        this.confirmedPair = { doc: view.state.doc, canonical: this.canonical };
        this.agreedCanonical = this.canonical;
        return;
      }
      const tr = view.state.tr;
      const start = map.canonical.content.findDiffStart(this.canonical.content);
      const end = map.canonical.content.findDiffEnd(this.canonical.content);
      let fitted = false;
      // Filling an empty view with a long document: one fitted step would put
      // every block on one sheet, laid out whole before its first page could
      // show. The wrapped document opens as provisional sheets instead, which
      // the page ledger composes progressively (provisionalPages).
      const bulkFill = map.canonical.childCount <= 1 && map.canonical.textContent === "" &&
        this.canonical.childCount >= PROVISIONAL_PAGES_MIN_BLOCKS;
      // The kept layout the fill would open on has not come in yet: the
      // fill waits for it (the page's read now answers before the rows of a
      // long document do, and the fill laid it out cold).
      if (bulkFill && this.options.initialLayoutReceiptPending?.()) {
        if (!this.fillHeldForReceipt) {
          this.fillHeldForReceipt = true;
          setTimeout(() => {
            this.fillHeldForReceipt = false;
            if (!this.closed) this.applyCanonical(selection);
          }, FILL_RECEIPT_POLL_MS);
        }
        return;
      }
      // The blocks that differ, each replaced inside its own sheet. One step
      // from the first difference to the last spanned every sheet between
      // them with unpaged canonical blocks: an Undo of one checkbox merged
      // three sheets, the page shrank and grew back (the reader's caret was
      // scrolled to on the shrunk page and the view jumped by a page).
      if (start !== null && end && !bulkFill) fitted = this.fitChangedBlocks(tr, map);
      if (!fitted && start !== null && end && !bulkFill) {
        const overlap = Math.max(0, start - Math.min(end.a, end.b));
        const oldEnd = end.a + overlap, newEnd = end.b + overlap;
        const from = map.toLocal(start, 1);
        const to = oldEnd === start ? from : map.toLocal(oldEnd, -1);
        if (from !== null && to !== null && from <= to) {
          const sourceSlice = this.canonical.slice(start, newEnd);
          const slice = Slice.fromJSON(this.options.schema, sourceSlice.toJSON());
          const step = new ReplaceStep(from, to, slice);
          const applied = step.apply(tr.doc);
          if (applied.doc && projectPagedDocument(applied.doc).content.eq(this.canonical.content)) { tr.step(step); fitted = true; }
        }
      }
      if (!fitted) {
        const restored = bulkFill ? this.restoredLayout(view) : null;
        if (bulkFill) markRestoredLayout(view.dom, restored !== null);
        const knownSizes = bulkFill && !restored ? this.options.initialLayoutSizes?.() ?? null : null;
        if (bulkFill && !restored) traceDocsLayoutReceipt(`bridge fill sizes=${knownSizes ? knownSizes.length : "none"}`);
        const roomAttrs = bulkFill ? this.roomDocumentAttrs(view.state.doc) : null;
        // The room's content as it opens, before the editor normalizes it: a
        // receipt kept later from these pages names it (the next opening of a
        // page nobody saved meets exactly it).
        if (bulkFill) rememberOpenedProjectionContent(view.dom, { ...(this.canonical.toJSON() as JSONContent), attrs: { ...view.state.doc.attrs, ...(roomAttrs ?? {}) } });
        const local = restored ?? wrapCanonicalDocument(this.canonical, this.options.schema,
          roomAttrs ? { ...view.state.doc.attrs, ...roomAttrs } : view.state.doc.attrs, knownSizes);
        if (bulkFill && this.fillInChunks(view, local, roomAttrs)) return;
        tr.replaceWith(0, tr.doc.content.size, local.content);
        if (bulkFill && viewMayDrawHollow(view, tr.doc)) fillHollowFrom(tr, local.content);
        if (roomAttrs) for (const [key, value] of Object.entries(roomAttrs)) tr.setDocAttribute(key, value);
      }
      const restored = this.selectionFor(tr.doc, selection);
      if (restored) tr.setSelection(restored);
      tr.setMeta(PAGED_CANONICAL_REMOTE_META, true).setMeta(ySyncPluginKey, { isChangeOrigin: true, isUndoRedoOperation: this.historyTarget !== undefined }).setMeta("addToHistory", false);
      this.applying = true;
      try { view.dispatch(tr); } finally { this.applying = false; }
      this.agreedCanonical = this.canonical;
      // A fitted step was proven to project to the canonical content above.
      if (fitted && view.state.doc === tr.doc) this.confirmedPair = { doc: tr.doc, canonical: this.canonical };
    } catch { this.setIssue("position-unavailable"); }
  }

  /** The room's document on the pages of the cached layout receipt, or null. */
  private restoredLayout(view: EditorView): ProseMirrorNode | null {
    try {
      // The room's content under the document's own settings, as the bulk
      // fill lays it out (the canonical node carries only their defaults; the
      // room's settings arrive beside it): the receipt was kept from pages
      // that held them, and its hash names them.
      const roomAttrs = this.roomDocumentAttrs(view.state.doc);
      const canonical = { ...(this.canonical.toJSON() as JSONContent), attrs: { ...view.state.doc.attrs, ...(roomAttrs ?? {}) } };
      const receipt = this.options.initialLayoutReceipt?.(canonical, view.dom);
      traceDocsLayoutReceipt(`bridge fill receipt=${Boolean(receipt)}`);
      if (!receipt) return null;
      const restored = restorePaginationProjection(canonical, receipt);
      traceDocsLayoutReceipt(`bridge restore ${restored ? "ok" : `refused ${tracingDocsLayoutReceipt() ? paginationProjectionRefusal(canonical, receipt) : ""}`}`);
      if (!restored) return null;
      const document = this.options.schema.nodeFromJSON({ ...restored, attrs: { ...view.state.doc.attrs, ...(roomAttrs ?? {}) } });
      document.check();
      return projectPagedDocument(document).content.eq(this.canonical.content) ? document : null;
    } catch {
      return null;
    }
  }

  /**
   * Replace, inside the sheets that hold them, only the top-level blocks the
   * room changed, when the room has the same number of blocks as this view
   * projects (an edit inside blocks: typing, formatting, an Undo of either).
   * Proven like the single step: the result must project to the room's
   * content, else the caller's step is used. False leaves `tr` untouched.
   */
  private fitChangedBlocks(tr: Transaction, map: PagedDocumentPositionMap): boolean {
    const local = map.canonical, room = this.canonical;
    if (local.childCount !== room.childCount) return false;
    const changes: { from: number; to: number; node: ProseMirrorNode }[] = [];
    let offset = 0;
    for (let index = 0; index < room.childCount; index += 1) {
      const before = local.child(index), after = room.child(index);
      if (!sameRoomNode(before, after)) {
        const from = map.toLocal(offset, 1);
        const to = map.toLocal(offset + before.nodeSize, -1);
        if (from === null || to === null || from > to) return false;
        // A block the view shows in pieces across sheets (a split table)
        // is left to the single step.
        if (!tr.doc.resolve(from).sameParent(tr.doc.resolve(to))) return false;
        changes.push({ from, to, node: Fragment.fromJSON(this.options.schema, [after.toJSON()]).firstChild! });
      }
      offset += before.nodeSize;
    }
    if (changes.length > 64) return false;
    // Nothing the room changed: what differs is drawn from it here.
    if (changes.length === 0) return true;
    // Last first, so each block's positions hold for the ones before it.
    let doc = tr.doc;
    const steps: ReplaceStep[] = [];
    for (let index = changes.length - 1; index >= 0; index -= 1) {
      const change = changes[index]!;
      const step = new ReplaceStep(change.from, change.to, new Slice(Fragment.from(change.node), 0, 0));
      const applied = step.apply(doc);
      if (!applied.doc) return false;
      doc = applied.doc;
      steps.push(step);
    }
    if (!sameRoomContent(projectPagedDocument(doc).content, room.content)) return false;
    for (const step of steps) tr.step(step);
    return true;
  }

  /**
   * Fill an empty view with a long document a few hundred blocks at a time.
   *
   * One dispatch of a whole book (7,390 blocks) built every block's DOM and
   * ran every plugin over it in a single main-thread task of 10+ seconds, and
   * every other window, the sidebar and the page library froze with it. The
   * first dispatch holds the first sheets, the reader's first pages; each
   * later one appends the next sheets in a task of its own. The view takes no
   * authored edit until the last sheet is in (`editable`, `filterTransaction`)
   * and nothing is published meanwhile; a room change under the fill turns
   * the rest into an ordinary difference (`applyCanonical`).
   */
  /**
   * The room's document attributes (page setup, styles: the shared attribute
   * map) that differ from the view's, for the fill of an empty view to carry.
   * Set by NoteEditor's attribute binding in a dispatch of its own right
   * after the fill, they ran every plugin and decoration over the whole
   * document a second time: about 1 s of the task that opens the 298-page
   * book. The binding then finds them set and dispatches nothing.
   */
  private roomDocumentAttrs(doc: ProseMirrorNode): Record<string, unknown> | null {
    try {
      const room = readCollaborativeDocumentAttrs(this.options.session.ydoc.getMap(COLLABORATIVE_DOCUMENT_ATTRS_MAP));
      const changed: Record<string, unknown> = {};
      for (const [key, value] of Object.entries(room)) {
        if (!Object.hasOwn(doc.attrs, key) || pageDocumentJsonEquals(doc.attrs[key], value)) continue;
        changed[key] = value;
      }
      return Object.keys(changed).length > 0 ? changed : null;
    } catch {
      return null;
    }
  }

  private fillInChunks(view: EditorView, local: ProseMirrorNode, roomAttrs: Record<string, unknown> | null = null): boolean {
    // A long document fills an empty view in chunks, a task each: the single
    // dispatch of a 298-page book's pages was one main-thread task of 5-6 s
    
    // cached layout) once went in at once, as their initial settle after a
    
    // pending-sheet convergence fix the chunked fill settles first, 130-136 s
    // against 257-259 s, with frames of at most 2.4 s against 4.7-6.0 s
    
    // window.__pagedBridgeChunkedFill = false.
    const forced = (globalThis as { __pagedBridgeChunkedFill?: boolean }).__pagedBridgeChunkedFill;
    if (forced === false) return false;
    // docs-hollow: the document goes in with one dispatch, its later sheets
    // hollow (fillHollowFrom): building their DOM was what the chunks spread
    // out, and the view is editable from the first frame instead of after
    // the last chunk (the book: about 8 s it could not be typed in).
    if (!fillsInChunks(hollowSheetsEnabled(), viewMayDrawHollow(view, local))) return false;
    const pages: ProseMirrorNode[] = [];
    local.forEach(page => pages.push(page));
    let next = 0, blocks = 0;
    while (next < pages.length && blocks < FILL_FIRST_BLOCKS) blocks += pages[next++].childCount;
    if (next >= pages.length) return false;
    // The chunks after the first are background work (lib/background-work):
    // input elsewhere in the app goes first, input in this window puts them
    // first, and they never wait longer than the deferral cap.
    const lane = createBackgroundLane({ label: "docs-fill", owner: () => view.dom.closest(".wm-window") ?? view.dom });
    this.filling = { canonical: this.canonical, pages, next, filled: blocks, lane };
    // The page ledger holds its passes while the sheets are still arriving
    // (the pages stay hidden under the initial settle meanwhile).
    writeDocumentLayoutStatus(view.dom, PAGED_BULK_FILL_ATTR, "true", view);
    const tr = view.state.tr.replaceWith(0, view.state.doc.content.size, pages.slice(0, next));
    if (roomAttrs) for (const [key, value] of Object.entries(roomAttrs)) tr.setDocAttribute(key, value);
    this.dispatchFill(view, tr.setMeta(BULK_FILL_CHUNK_META, true));
    this.scheduleFillChunk();
    return true;
  }

  private dispatchFill(view: EditorView, tr: Transaction) {
    tr.setMeta(PAGED_CANONICAL_REMOTE_META, true).setMeta(ySyncPluginKey, { isChangeOrigin: true, isUndoRedoOperation: false })
      .setMeta("addToHistory", false);
    this.applying = true;
    try { view.dispatch(tr); } finally { this.applying = false; }
  }

  private scheduleFillChunk() {
    const fill = this.filling;
    if (!fill) return;
    // A task of its own each time, after any input waiting for the thread.
    void fill.lane.turn().then(() => { if (this.filling === fill) this.fillChunk(); });
  }

  private fillChunk = () => {
    const fill = this.filling, view = this.view;
    if (fill === null) return;
    if (this.closed || !view) { this.endFill(view); return; }
    if (fill.canonical !== this.canonical) {
      this.finishFillAtOnce(view);
      this.applyCanonical(null);
      this.scheduleRefresh();
      queueMicrotask(() => this.flushFillEdits());
      return;
    }
    // Chunks grow with what is already in: the plugins and the view that
    // still run over the whole document on each dispatch run fewer times.
    let end = fill.next, blocks = 0;
    const budget = Math.max(FILL_CHUNK_BLOCKS, Math.round(fill.filled / 3));
    while (end < fill.pages.length && blocks < budget) blocks += fill.pages[end++].childCount;
    const last = end >= fill.pages.length;
    // The last chunk's dispatch is the one that starts the ledger's first
    // pass, and the one whole-document decorations are rebuilt on.
    if (last) writeDocumentLayoutStatus(view.dom, PAGED_BULK_FILL_ATTR, null, view);
    try {
      const tr = view.state.tr.insert(view.state.doc.content.size, fill.pages.slice(fill.next, end));
      this.dispatchFill(view, last ? tr : tr.setMeta(BULK_FILL_CHUNK_META, true));
    } catch {
      this.endFill(view);
      this.setIssue("position-unavailable");
      return;
    }
    if (this.filling !== fill) return;
    fill.next = end;
    fill.filled += blocks;
    if (end < fill.pages.length) { this.scheduleFillChunk(); return; }
    this.endFill(view);
    // The view is complete: it may take input now (`editable` is read on
    // the next update), what was asked of it meanwhile is applied, and its
    // caret is published at the pause.
    this.scheduleRefresh();
    // After that refresh: the view must be editable for its edit to pass.
    queueMicrotask(() => this.flushFillEdits());
    this.deferSelectionPublish();
  };

  /** End a fill early: the sheets still missing go in with one last dispatch,
   * which carries no chunk meta, so the whole-document decorations the
   * chunks only mapped are rebuilt. */
  private finishFillAtOnce(view: EditorView) {
    const fill = this.filling;
    this.endFill(view);
    if (!fill) return;
    try {
      this.dispatchFill(view, view.state.tr.insert(view.state.doc.content.size, fill.pages.slice(fill.next)));
    } catch { this.setIssue("position-unavailable"); }
  }

  private endFill(view: EditorView | null) {
    this.filling?.lane.dispose();
    this.filling = null;
    if (view) writeDocumentLayoutStatus(view.dom, PAGED_BULK_FILL_ATTR, null, view);
  }

  private holdEditsUntilWritable() {
    if (this.heldEditsTimer !== null) return;
    this.heldEditsSince = Date.now();
    this.reportBridgeEvent("PagedBridgeHeldEdits", "write_gate_closed");
    this.heldEditsTimer = setInterval(() => {
      if (this.closed || this.fillEdits.length === 0) { this.stopHoldingEdits(); return; }
      if (this.filling !== null) return; // the fill's end applies them
      if (this.mayPublish()) {
        this.stopHoldingEdits();
        this.flushFillEdits();
        return;
      }
      // Kept as long as the view is open: a long wait is reported once.
      if (this.heldEditsSince > 0 && Date.now() - this.heldEditsSince > 600_000) {
        this.heldEditsSince = 0;
        this.reportBridgeEvent("PagedBridgeHeldEditsLong", "write_gate_closed");
      }
    }, 250);
  }

  private stopHoldingEdits() {
    if (this.heldEditsTimer !== null) clearInterval(this.heldEditsTimer);
    this.heldEditsTimer = null;
  }

  /** Apply the authored edits made during the fill, each mapped through what
   * the view took after it (the rest of the sheets, any room change) and
   * through the edits applied before it, as a collaborative rebase does. */
  private flushFillEdits() {
    const edits = this.fillEdits.splice(0);
    const view = this.view;
    if (edits.length === 0 || this.closed || !view) return;
    const tr = view.state.tr;
    for (const edit of edits) {
      const mapping = new Mapping();
      const own = edit.steps.map(step => step.getMap());
      for (let index = own.length - 1; index >= 0; index -= 1) mapping.appendMap(own[index].invert());
      mapping.appendMapping(edit.since);
      mapping.appendMapping(tr.mapping);
      edit.steps.forEach((step, index) => {
        const mapped = step.map(mapping.slice(own.length - index));
        if (!mapped || tr.maybeStep(mapped).failed) return;
        // Mirrors the inverse of the same step taken out above.
        mapping.appendMap(mapped.getMap(), own.length - index - 1);
      });
    }
    // Only what the writer asked for is an Undo step; what normalizers drew
    // on the arriving sheets is not.
    const authored = edits.some(edit => edit.authored);
    if (tr.docChanged) view.dispatch(authored ? tr : tr.setMeta("addToHistory", false));
  }

  private onHistoryAdded = ({ stackItem }: { stackItem: { meta: Map<unknown, unknown> } }) => {
    if (this.localHistory) stackItem.meta.set(this, settledCarets(this.localHistory));
    // Undo pushes the step it took back onto the redo stack (and Redo onto
    // the undo stack) as a new stack item with empty meta: without the
    // step's carets, Redo restored no caret, and the one mapped through the
    // rebuilt content landed in the next paragraph -- typing after Redo of
    // an automatic list or of a row added by Tab went there (WUA 065).
    else if (this.historyCarets) stackItem.meta.set(this, this.historyCarets);
    this.scheduleRefresh();
  };
  private onHistoryUpdated = ({ stackItem }: { stackItem: { meta: Map<unknown, unknown> } }) => {
    const existing = stackItem.meta.get(this) as HistorySelection | undefined;
    const current = this.localHistory;
    if (current) settledCarets(current);
    if (current) stackItem.meta.set(this, { before: existing?.before ?? current.before, beforeDoc: existing?.beforeDoc ?? current.beforeDoc,
      get after() { return current.after; }, get afterDoc() { return current.afterDoc; } });
    this.scheduleRefresh();
  };

  /** Called only after explicit Save has drained IME/layout serialization.
   * Validate against the current room, then materialize view defaults and
   * generated identities in that same original outside personal history. */
  prepareExplicitSave(): boolean {
    if (this.view && this.established && (this.issue === "invalid-document" || this.issue === "position-unavailable")) {
      this.recoverFromIssue();
    }
    this.explicitSaveRefusal = !this.view ? "no_view" : !this.established ? "not_established"
      : !this.mayPublish() ? this.publishRefusal() : null;
    if (this.explicitSaveRefusal !== null) return false;
    try {
      const projected = projectPagedDocument(this.view!.state.doc);
      const original = readDocumentCausalProjection(this.options.session.ydoc);
      const requested = documentCausalSemanticContent(serializeWhiteboardScenes(projected.toJSON()));
      if (!isDocumentSaveMaintenance(original, requested, this.schema)) {
        // Edits the room never took (a publication that failed, a change
        // outside Undo): merged with what the room took meanwhile and
        
        // twenty minutes of ink were never saved).
        return this.reconcileViewIntoRoom();
      }
      if (pageDocumentJsonEquals(original, requested)) return true;
      this.publishing = true;
      try {
        this.options.session.ydoc.transact(() => {
          updateCanonicalYFragment(this.options.session.ydoc, this.fragment, projected, this.meta);
          publishCollaborativeDocumentAttrs(this.options.session.ydoc.getMap(COLLABORATIVE_DOCUMENT_ATTRS_MAP), projected.attrs);
        }, this.nonHistoryOrigin);
        const read = this.readCanonical();
        if (!read) this.explicitSaveRefusal = "room_read_back_failed";
        return read;
      } finally { this.publishing = false; this.flushCanonicalRead(); }
    } catch {
      this.explicitSaveRefusal = "exception";
      return false;
    }
  }

  private static blocksOf(node: ProseMirrorNode): MergeBlock[] {
    return ((node.toJSON() as JSONContent).content ?? []) as MergeBlock[];
  }

  /** Write `doc` into the room outside personal history and read it back. */
  private publishCanonicalDocument(doc: ProseMirrorNode): boolean {
    this.publishing = true;
    this.foreignDuringPublish = false;
    try {
      this.options.session.ydoc.transact(() =>
        updateCanonicalYFragment(this.options.session.ydoc, this.fragment, doc, this.meta), this.nonHistoryOrigin);
    } finally { this.publishing = false; }
    const read = this.readCanonical();
    this.flushCanonicalRead();
    return read;
  }

  






  private keepUnpublishedLocal(view: EditorView) {
    const base = this.agreedCanonical;
    if (!base || !this.established || this.filling !== null || this.issue !== null || this.closed) return;
    let local: ProseMirrorNode;
    try { local = this.positionMap(view.state.doc).canonical; } catch { return; }
    if (local.content.eq(base.content) || local.content.eq(this.canonical.content)) return;
    if (!this.mayPublish()) return;
    const blocks = PagedDocumentBridge.blocksOf;
    const merged = mergeCanonicalBlocks(blocks(base), blocks(local), blocks(this.canonical), { onlyAdditions: true });
    if (!merged.changedRemote) return;
    try {
      const doc = parseCanonicalDocumentJson(this.schema, { ...(this.canonical.toJSON() as JSONContent), content: merged.blocks as JSONContent[] });
      if (this.publishCanonicalDocument(doc)) this.reportBridgeEvent("PagedBridgeKeptUnpublished", "kept");
    } catch (error) {
      this.reportBridgeEvent("PagedBridgeKeepFailed", "merge_failed", error);
    }
  }

  /** Clear an issue a failed publication or mapping left, if the view and
   * the room read again; true when the bridge takes input again. */
  private recoverFromIssue(): boolean {
    const issue = this.issue;
    if (!this.view || (issue !== "invalid-document" && issue !== "position-unavailable")) return issue === null;
    this.issue = null;
    try {
      if (!this.readRoomCanonical()) throw new Error("room_unreadable");
      this.positionMap(this.view.state.doc);
      this.confirmedPair = null;
      this.recoverAttempts = 0;
      this.reportBridgeEvent("PagedBridgeRecovered", issue);
      queueMicrotask(() => { if (!this.closed && this.issue === null) this.options.onIssue?.(null); });
      this.scheduleRefresh();
      return true;
    } catch (error) {
      this.issue = issue;
      this.reportBridgeEvent("PagedBridgeRecoverFailed", issue, error);
      return false;
    }
  }

  





  private reconcileViewIntoRoom(): boolean {
    const view = this.view;
    if (!view || !this.established) { this.explicitSaveRefusal = !view ? "no_view" : "not_established"; return false; }
    if (!this.mayPublish()) { this.explicitSaveRefusal = this.publishRefusal(); return false; }
    try {
      const blocks = PagedDocumentBridge.blocksOf;
      const local = this.positionMap(view.state.doc).canonical;
      const base = this.agreedCanonical;
      const merged = mergeCanonicalBlocks(base ? blocks(base) : [], blocks(local), blocks(this.canonical));
      this.recoveryConflicts.push(...merged.conflictPairs);
      if (merged.changedRemote) {
        const doc = parseCanonicalDocumentJson(this.schema, { ...(this.canonical.toJSON() as JSONContent), content: merged.blocks as JSONContent[] });
        if (!this.publishCanonicalDocument(doc)) { this.explicitSaveRefusal = "room_read_back_failed"; return false; }
      }
      if (merged.changedLocal) this.applyCanonical(null);
      this.agreedCanonical = this.canonical;
      // Which way the views had parted, content-free: `room` means this view
      // held edits the room never took (its members never saw them until
      // now); `view` that the room held some this view never showed.
      if (merged.changedRemote || merged.changedLocal) {
        const parted = [merged.changedRemote ? "room" : "", merged.changedLocal ? "view" : ""].filter(Boolean).join("+");
        this.reportBridgeEvent("PagedBridgeReconciled",
          `${merged.conflicts > 0 ? "kept_both" : "merged"}:${parted}:base_${base ? "agreed" : "none"}`);
      }
      this.explicitSaveRefusal = null;
      return true;
    } catch (error) {
      this.explicitSaveRefusal = "reconcile_failed";
      this.reportBridgeEvent("PagedBridgeReconcileFailed", "exception", error);
      return false;
    }
  }

  






  adoptRecoveredContent(content: unknown, sameBase: boolean, baseContent?: unknown,
    options: { baseOnlyIfHeld?: boolean; fallbackBaseContent?: unknown } = {}): boolean {
    const view = this.view;
    if (!view || !this.established || this.filling !== null || this.closed) return false;
    if (this.issue === "invalid-document" || this.issue === "position-unavailable") this.recoverFromIssue();
    if (!this.mayPublish()) return false;
    try {
      const draft = parseCanonicalDocumentJson(this.schema,
        withoutBareTopLevelInlineNodes(this.schema, canonicalDocumentJson(content as JSONContent)).json);
      const blocks = PagedDocumentBridge.blocksOf;
      // The body the draft was edited from, when known, is the merge base:
      // only what the draft changed against it is taken, and only a block
      // both sides changed is kept twice.
      const blocksOfBody = (body: unknown) => blocks(parseCanonicalDocumentJson(this.schema,
        withoutBareTopLevelInlineNodes(this.schema, canonicalDocumentJson(body as JSONContent)).json));
      let base = sameBase ? blocks(this.canonical) : baseContent !== undefined ? blocksOfBody(baseContent) : [];
      // A base only offered as "the save that was on its way" (a reload while
      // it was in flight) is the base only if the page took that save. A page
      // that did not is merged against the body the copy was edited from;
      // with no base at all every block that reads differently is kept twice.
      if (options.baseOnlyIfHeld && !sameBase && !holdsEveryBlock(base, blocks(this.canonical))) {
        base = options.fallbackBaseContent !== undefined ? blocksOfBody(options.fallbackBaseContent) : [];
      }
      // A block the room and the copy both only typed into from the base, one
      // holding every word of the other (the room has the copy's earlier saves,
      // the copy those words and the keys typed since), is the fuller one, not
      // two: a whole view whose conflict was read again kept its paragraph
      
      // 11,003 blocks). Different words on each side are still both kept.
      const merged = mergeCanonicalBlocks(base, blocks(draft), blocks(this.canonical), { typedIntoBoth: true });
      this.recoveryConflicts.push(...merged.conflictPairs);
      if (merged.changedRemote) {
        const doc = parseCanonicalDocumentJson(this.schema, { ...(this.canonical.toJSON() as JSONContent), content: merged.blocks as JSONContent[] });
        if (!this.publishCanonicalDocument(doc)) return false;
      }
      this.applyCanonical(null);
      this.agreedCanonical = this.canonical;
      this.reportBridgeEvent("PagedBridgeAdoptedRecovery", merged.conflicts > 0 ? "kept_both" : sameBase ? "same_base" : "merged");
      return true;
    } catch (error) {
      this.reportBridgeEvent("PagedBridgeAdoptRecoveryFailed", "exception", error);
      return false;
    }
  }

  


  /** Whether the room's document has been read into this bridge. */
  isEstablished(): boolean { return this.established && !this.closed; }

  takeRecoveryConflicts(): Array<{ kept: string | null; copy: string | null }> {
    const taken = this.recoveryConflicts;
    this.recoveryConflicts = [];
    return taken;
  }

  /** Why the last explicit-save preparation refused, as a content-free code
   * (lib/page-save-diagnostics), or null after a successful one. */
  lastExplicitSaveRefusal(): string | null { return this.explicitSaveRefusal; }

  /** The first reason mayPublish() is false, as a code. */
  private publishRefusal(): string {
    if (this.closed) return "bridge_closed";
    if (this.issue !== null) return `issue_${this.issue}`;
    const network = this.options.session.network;
    if (network && !network.failed) {
      if (network.connected && !network.writable) return "room_read_only";
      if (!network.connected && !network.ready && this.options.durableOriginals !== true) return "room_not_ready";
      if (!network.connected && this.options.canonicalNetworkModel?.() !== true) return "model_not_negotiated";
    }
    return "write_authority_withheld";
  }

  /** Materialize the verified original's schema representation (the
   * default attributes and mark normalization the editor reads into every
   * untouched node) in the room before the user's first gesture, outside
   * personal history. Otherwise the first edit captures them in its Undo,
   * and an edit that writes no node (document attributes only) leaves the
   * room without them, so the page never mirrors and never saves. Runs
   * inside a publishing section; false when the room cannot be read back. */
  /**
   * What the first key would find out in its own task, found out at an idle
   * moment instead: whether the room already holds the normalized original (an
   * earlier session wrote it), and which of the room's blocks stand equal to the
   * view's. It writes nothing, and an edit in the meantime just ends it: the first
   * key then does what is left (a 124-page report: 0.8 s of that key).
   */
  private readonly firstEditWarm = { armed: false, done: false, since: 0 };
  /** The room's blocks found standing equal to `originalBindingOf`'s blocks
   * by the idle warm (bindUnchangedInSlices), kept across its slices and
   * attempts and handed to materializeOriginal, whose own scan then skips
   * them: the first stroke of a 3000-stroke note compared every block again
   * in its own task. A block a room update touches is dropped from it
   * (noteTouchedElements); a different canonical document starts it anew. */
  private originalBinding = new Map<Y.XmlElement, ProseMirrorNode>();
  private originalBindingOf: ProseMirrorNode | null = null;
  private originalBindingFor(canonical: ProseMirrorNode): Map<Y.XmlElement, ProseMirrorNode> {
    if (this.originalBindingOf !== canonical || (globalThis as { __mraiFreshOriginalBinding?: boolean }).__mraiFreshOriginalBinding) {
      this.originalBinding = new Map();
      this.originalBindingOf = canonical;
    }
    return this.originalBinding;
  }
  private armFirstEditWarm() {
    if (this.firstEditWarm.armed || this.firstEditWarm.done || this.closed) return;
    this.firstEditWarm.armed = true;
    this.firstEditWarm.since ||= performance.now();
    setTimeout(() => { this.firstEditWarm.armed = false; void this.warmFirstEdit(); }, FIRST_EDIT_WARM_LOOK_MS);
  }
  private async warmFirstEdit() {
    if (this.closed || this.firstEditWarm.done) return;
    const view = this.view;
    const busy = !view || !this.established || !view.editable || !this.mayPublish() || this.publishing || this.applying ||
      msSinceTyping() < FIRST_EDIT_WARM_LOOK_MS || view.dom.hasAttribute("data-pagination-initial-settling") ||
      (performance.now() - this.firstEditWarm.since < FIRST_EDIT_WARM_LAYOUT_WAIT_MAX_MS && layoutStillMoving(view.dom));
    if (busy) { this.armFirstEditWarm(); return; }
    this.firstEditWarm.done = true;
    const startDoc = view.state.doc;
    const stopped = () => this.closed || this.publishing || this.applying || this.view !== view || msSinceTyping() < TYPING_IDLE_MS;
    try {
      const canonical = this.canonical;
      if (!this.originalNormalized) {
        const held = await bindUnchangedInSlices(this.fragment, canonical,
          {mapping: this.originalBindingFor(canonical), isOMark: this.meta.isOMark}, stopped);
        if (held && !stopped() && this.canonical === canonical) this.originalNormalized = true;
        // Stopped part way (an edit, a key): what it bound is kept, and it
        // goes on at the next idle moment.
        else if (!held && stopped() && !this.closed) { this.firstEditWarm.done = false; this.armFirstEditWarm(); return; }
      }
      if (stopped() || view.state.doc !== startDoc) return;
      const projected = this.positionMap(startDoc).canonical;
      const bound = await bindUnchangedInSlices(this.fragment, projected, this.meta, stopped);
      if (bound && !stopped()) await this.adoptViewBlocksInSlices(canonical, projected, stopped);
    } catch { /* The first key does it. */ }
  }

  













  private async adoptViewBlocksInSlices(read: ProseMirrorNode, projected: ProseMirrorNode, stopped: () => boolean, sliceMs = 6) {
    if (!(globalThis as { __mraiAdoptViewBlocks?: boolean }).__mraiAdoptViewBlocks) return;
    if (read === projected || read.childCount !== projected.childCount) return;
    // A read the tolerant reader repaired keeps its own record (and binding).
    if (canonicalReadRepairOf(read) !== null || canonicalReadBlocksOf(read) !== undefined) return;
    for (let index = 0; index < read.childCount;) {
      if (stopped() || this.canonical !== read) return;
      const started = performance.now();
      for (; index < read.childCount && performance.now() - started < sliceMs; index += 1) {
        const mine = read.child(index), theirs = projected.child(index);
        if (mine !== theirs && !mine.eq(theirs)) return;
      }
      if (index < read.childCount) await new Promise((resolve) => setTimeout(resolve, 16));
    }
    if (stopped() || this.canonical !== read || !this.originalNormalized || this.issue !== null ||
        this.adoptedUnverified || this.canonicalReadPending || this.touchedElements.size > 0) return;
    const adopted = read.copy(projected.content);
    const blocks = this.roomBlocks();
    if (blocks && blocks.length === read.childCount) {
      blocks.forEach((block, index) => {
        if (this.verifiedElements.get(block) === read.child(index)) this.verifiedElements.set(block, adopted.child(index));
      });
    }
    if (this.agreedCanonical === read) this.agreedCanonical = adopted;
    this.canonical = adopted;
  }

  private materializeOriginal(): boolean {
    if (this.originalNormalized) return true;
    const ydoc = this.options.session.ydoc;
    // A room an earlier session normalized takes no write: nothing to read back,
    // and the whole document parsed again in the first key typed (~0.4 s of it
    // on a table-heavy 124 pages).
    let wrote = false;
    const noteWrite = () => { wrote = true; };
    ydoc.on("update", noteWrite);
    try {
      const binding = this.originalBindingFor(this.canonical);
      ydoc.transact(() =>
        updateCanonicalYFragment(ydoc, this.fragment, this.canonical,
          {mapping: binding, isOMark: this.meta.isOMark}), this.nonHistoryOrigin);
    } finally {
      ydoc.off("update", noteWrite);
    }
    if (wrote && !this.readCanonical()) return false;
    this.originalNormalized = true;
    this.originalBinding = new Map();
    this.originalBindingOf = null;
    return true;
  }

  /**
   * What the first edit of a session reads once (the whole document's position
   * map, the room's normalization) done ahead of it: the pen was picked up, so a
   * stroke is coming, and on a long handwritten note this held the first stroke's
   * landing for half a second. Nothing here is a change of the document; the
   * first edit does the same when this has not run.
   */
  warmForFirstEdit(): void {
    const view = this.view;
    if (this.closed || !view || !this.established || !this.mayPublish()) return;
    if (this.originalNormalized && this.positions.has(view.state.doc)) return;
    // The room compared with the original a few blocks a frame first (what it
    // proves equal is kept for materializeOriginal): done here in one task it
    // was 130-140 ms of a 3000-stroke note in WebKit as the pen was picked up.
    // A stroke landing meanwhile does what is left itself.
    if (!this.originalNormalized && !(globalThis as { __mraiEagerFirstEditWarm?: boolean }).__mraiEagerFirstEditWarm) {
      if (this.originalWarming) return;
      const canonical = this.canonical;
      const stopped = () => this.closed || this.publishing || this.applying || this.canonical !== canonical || this.originalNormalized;
      this.originalWarming = true;
      void bindUnchangedInSlices(this.fragment, canonical,
        {mapping: this.originalBindingFor(canonical), isOMark: this.meta.isOMark}, stopped)
        .then((held) => {
          if (held && !stopped()) this.originalNormalized = true;
        })
        .catch(() => undefined)
        .finally(() => {
          this.originalWarming = false;
          if (!this.closed && this.view === view) this.warmForFirstEditNow();
        });
      return;
    }
    this.warmForFirstEditNow();
  }

  private originalWarming = false;

  private warmForFirstEditNow(): void {
    const view = this.view;
    if (this.closed || !view || !this.established || !this.mayPublish()) return;
    if (this.originalNormalized && this.positions.has(view.state.doc)) return;
    this.publishing = true;
    this.foreignDuringPublish = false;
    try {
      if (this.materializeOriginal()) this.positionMap(view.state.doc);
    } catch {
      // The first edit reads the same things, and reports what it finds.
    } finally {
      this.publishing = false;
      this.flushCanonicalRead();
    }
  }

  isApplyingHistory(): boolean { return this.historyTarget !== undefined; }

  publishDocumentAttrs(publish: () => boolean, options: { historical: boolean }): boolean {
    if (this.closed || !this.mayPublish()) return false;
    // Its own dispatch's attributes join the step; a later attribute gesture
    // is not folded into a step still open for its follow-up.
    if (options.historical && this.openStep?.awaitsFollowUp && this.openStep.settled) this.closeOpenStep();
    // The caret is the Undo step's, so it is read only when the publication
    // makes or extends one (the history handlers below read it, in the
    // publication's own transaction, before the view can change). Every key
    // typed publishes the document attributes, and almost never changes
    // one: reading the caret as a room position walks the room from its
    // start, ~12 ms a key at the end of the 296-page book (dev).
    const view = this.view;
    let selection: RelativeSelection | null | undefined;
    const caret = () => selection === undefined ? (selection = view ? this.captureSelection(view.state) : null) : selection;
    this.localHistory = options.historical ? { get before() { return caret(); }, get after() { return caret(); } } : null;
    this.publishing = true;
    let changed = false;
    try {
      // A gesture that changes only document attributes (a header or footer
      // edit, the page colour) is the first edit of a freshly seeded room as
      // much as a typed character is: the save mirrors the whole document
      // against the room, so the original's schema representation is
      // materialized first here too, outside personal history.
      if (options.historical && this.established && !this.materializeOriginal()) return false;
      // The existing attribute publisher can retain its own nested transact;
      // this outer origin binds only this view's user operation to its Undo.
      this.options.session.ydoc.transact(() => { changed = publish(); },
        options.historical ? this.localOrigin : this.nonHistoryOrigin);
      return changed;
    } finally {
      this.publishing = false;
      this.localHistory = null;
      this.flushCanonicalRead();
      if (changed) this.scheduleRefresh();
    }
  }

  history(direction: "undo" | "redo", dispatch = true): boolean {
    const stack = direction === "undo" ? this.undoManager.undoStack : this.undoManager.redoStack;
    if (!this.mayPublish() || stack.length === 0) return false;
    if (!dispatch) return true;
    this.closeOpenStep();
    const metadata = stack[stack.length - 1].meta.get(this) as HistorySelection | undefined;
    this.historyTarget = (direction === "undo" ? metadata?.before : metadata?.after) ?? null;
    this.historyCarets = metadata;
    const contentBefore = this.canonical;
    try {
      let done = (direction === "undo" ? this.undoManager.undo() : this.undoManager.redo()) !== null;
      // Yjs drops a step whose every change a peer has since rewritten and
      // answers null: its words are then taken back by identity instead.
      if (direction === "undo" && this.undoTextOfRewrittenStep(metadata, contentBefore)) done = true;
      // The whole operation is in the view now (the content and the document
      // attributes arrive as separate updates, each its own dispatch): say so
      // once, still inside the operation, so a plugin can append what the
      // final state calls for (sectionChainHistoryRepair, WUA 168).
      const view = this.view;
      if (done && view && !this.closed) {
        view.dispatch(view.state.tr.setMeta(PAGED_HISTORY_APPLIED_META, direction).setMeta("addToHistory", false));
      }
      return done;
    }
    finally { this.historyTarget = undefined; this.historyCarets = undefined; this.publishSelection(); this.scheduleRefresh(); }
  }

  /**
   * An Undo that took nothing back because a peer rewrote what the step
   * touched (a list level change writes the item again: WUA 149) takes the
   * step's words back by block identity instead (restoreStepTextByIdentity),
   * outside personal history. Without it the step was spent and its words
   * stayed, with Undo and Redo both empty.
   */
  private undoTextOfRewrittenStep(metadata: HistorySelection | undefined, contentBefore: ProseMirrorNode): boolean {
    if (!metadata?.beforeDoc || !metadata.afterDoc) return false;
    this.flushCanonicalRead();
    if (!this.canonical.content.eq(contentBefore.content)) return false;
    const restored = restoreStepTextByIdentity(this.canonical, metadata.afterDoc, metadata.beforeDoc);
    if (!restored) return false;
    this.publishing = true;
    try {
      this.options.session.ydoc.transact(() =>
        updateCanonicalYFragment(this.options.session.ydoc, this.fragment, restored, this.meta), this.nonHistoryOrigin);
      this.readCanonical();
    } finally { this.publishing = false; this.flushCanonicalRead(); }
    this.applyCanonical(metadata.before ?? null);
    this.reportBridgeEvent("PagedBridgeUndoByIdentity", "restored");
    return true;
  }

  /** `background`: the caret did not move by the reader's hand (a pause
   * after pagination passes or after the first fill). With nobody else in
   * the room it is not mapped then -- mapping a long book's caret into the
   * room's coordinates maps the whole paged document, seconds of main
   * thread -- and is published when someone arrives (`onPresenceChange`). */
  private publishSelection(background = false) {
    if (!this.view || this.closed || this.applying) return;
    if (background && this.options.session.awareness.getStates().size <= 1) {
      this.selectionPublishPending = true;
      return;
    }
    this.selectionPublishPending = false;
    // A long document still being filled in: mapping the caret into the
    // room's coordinates would map the partial document after every sheet
    // added. It is published once, at the pause after the fill.
    if (this.filling !== null) return;
    // Peers draw a caret and a text range: a cell selection shows as its
    // first cell's text, as before.
    const selection = this.captureSelection(this.view.state, false);
    if (selection) {
      const cursor = { anchor: selection.anchor, head: selection.head };
      const key = JSON.stringify(cursor);
      if (key !== this.cursorKey) { this.cursorKey = key; this.options.session.awareness.setLocalStateField("cursor", cursor); }
    }
  }

  /**
   * The others' carets in `state`'s own positions, for the caret layer
   * (lib/collab/peer-caret-overlay): never this user's other connections,
   * and only the `MAX_DRAWN_CARETS` most recently updated in a crowded room.
   */
  peerCarets(state: EditorState): PeerCaret[] {
    if (this.closed || this.issue !== null || !this.established) return [];
    const carets: PeerCaret[] = [];
    // Built only when a peer's caret is to be drawn: alone in a document,
    // a keystroke must not map the whole document for nothing.
    let map: PagedDocumentPositionMap | null = null;
    const awareness = this.options.session.awareness;
    const states = awareness.getStates() as Map<number, Record<string, unknown>>;
    const peers = [...states].filter(([clientId, value]) =>
      clientId !== awareness.clientID && value?.cursor && !isLocalUsersOtherConnection(states, awareness.clientID, value));
    if (peers.length > MAX_DRAWN_CARETS) {
      const updatedAt = (clientId: number) => awareness.meta.get(clientId)?.lastUpdated ?? 0;
      peers.sort(([left], [right]) => updatedAt(right) - updatedAt(left)).length = MAX_DRAWN_CARETS;
    }
    for (const [clientId, value] of peers) {
      try { map ??= this.positionMap(state.doc); } catch { return []; }
      try {
        const cursor = value.cursor as Readonly<{ anchor: Parameters<typeof relativePositionToAbsolutePosition>[2];
          head: Parameters<typeof relativePositionToAbsolutePosition>[2] }>;
        const a = relativePositionToAbsolutePosition(this.options.session.ydoc, this.fragment, cursor.anchor, this.meta.mapping);
        const h = relativePositionToAbsolutePosition(this.options.session.ydoc, this.fragment, cursor.head, this.meta.mapping);
        if (a === null || h === null) continue;
        // The peer's caret is in the room's coordinates; this view can be a
        // keystroke or an IME candidate ahead of the room for a moment.
        const anchor = map.toLocal(acrossDifference(this.canonical, map.canonical, a), 1);
        const head = map.toLocal(acrossDifference(this.canonical, map.canonical, h), 1);
        if (anchor === null || head === null) continue;
        const user = sanitizeCollabUser((value.user ?? {}) as { name?: unknown; color?: unknown });
        carets.push({ clientId, name: user.name, color: user.color, anchor, head });
      } catch { /* An invalid peer cursor has no document authority. */ }
    }
    return carets;
  }

  private scheduleRefresh = () => {
    if (this.closed || this.refreshQueued || !this.view) return;
    this.refreshQueued = true;
    const refresh = () => {
      // The room's Undo stack moved with every key: nothing shows it while the
      // hand is on the keys, and each refresh asked every plugin of the editor.
      const sinceTyping = msSinceTyping();
      if (sinceTyping < TYPING_IDLE_MS) {
        setTimeout(refresh, TYPING_IDLE_MS - sinceTyping);
        return;
      }
      this.refreshQueued = false;
      if (!this.closed && this.view) this.view.dispatch(this.view.state.tr.setMeta(HISTORY_META, true).setMeta("addToHistory", false));
    };
    if (msSinceTyping() < TYPING_IDLE_MS) refresh();
    else queueMicrotask(refresh);
  };

  createMappablePosition(position: number): MappablePosition {
    try {
      const map = this.view && this.positionMap(this.view.state.doc);
      const mapped = map && map.toCanonical(position);
      return new PagedMappablePosition(position, map && typeof mapped === "number" && this.established
        ? this.relativeRoomPosition(acrossDifference(map.canonical, this.canonical, mapped)) : null);
    } catch { return new PagedMappablePosition(position, null); }
  }

  getUpdatedPosition(position: MappablePosition, tr: Transaction): GetUpdatedPositionResult {
    if (position instanceof PagedMappablePosition && position.relative && (tr.getMeta(PAGED_CANONICAL_REMOTE_META) || tr.getMeta(REBALANCE_META))) {
      const canonical = relativePositionToAbsolutePosition(this.options.session.ydoc, this.fragment, position.relative, this.meta.mapping);
      const map = this.positionMap(tr.doc);
      const local = canonical === null ? null : map.toLocal(acrossDifference(this.canonical, map.canonical, canonical), 1, position.position);
      if (local !== null) return { position: new PagedMappablePosition(local, position.relative), mapResult: null };
    }
    const result = mapCorePosition(position, tr);
    return { ...result, position: new PagedMappablePosition(result.position.position, position instanceof PagedMappablePosition ? position.relative : null) };
  }

  destroy() {
    if (this.closed) return;
    this.closed = true;
    this.compositionHold?.release(); this.compositionHold = null;
    this.unsubscribeConnection?.(); this.unsubscribeConnection = null;
    this.options.session.ydoc.off("beforeAllTransactions", this.onBeforeYTransactions);
    this.options.session.ydoc.off("afterAllTransactions", this.onAfterYTransactions);
    this.options.session.ydoc.off("update", this.onYUpdate);
    this.options.session.ydoc.off("afterTransaction", this.onAfterYTransaction);
    this.options.session.awareness.off("change", this.onPresenceChange);
    this.options.session.awareness.off("change", this.scheduleLegacyRoomUpgrade);
    if (this.upgradeTimer !== null) { clearTimeout(this.upgradeTimer); this.upgradeTimer = null; }
    if (this.canonicalVerifyTimer !== null) { clearTimeout(this.canonicalVerifyTimer); this.canonicalVerifyTimer = null; }
    if (this.recoverTimer !== null) { clearTimeout(this.recoverTimer); this.recoverTimer = null; }
    if (this.layoutOnlyCheck?.timer != null) clearTimeout(this.layoutOnlyCheck.timer);
    this.layoutOnlyCheck = null;
    if (this.selectionPublishTimer !== null) { clearTimeout(this.selectionPublishTimer); this.selectionPublishTimer = null; }
    this.closedLayoutOnlyRuns = [];
    if (this.filling !== null) this.endFill(this.view);
    this.fillEdits = [];
    this.stopHoldingEdits();
    this.options.session.awareness.setLocalStateField("cursor", null);
    this.undoManager.destroy(); this.undoManager.clear();
    this.view = null; this.viewBinding = null; this.lastViewState = null;
  }
}
