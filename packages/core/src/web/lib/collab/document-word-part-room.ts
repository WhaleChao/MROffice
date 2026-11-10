/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * One part of a Word document held as parts, edited live in its own room
 * (`apps/api/app/routers/document_part_collab.py`, protocol `document-parts/1`).
 *
 * The wire is every part room's (PartRoomConnection); the model is the page
 * room's for a document (`document-canonical-v1`): the part's flat semantic
 * blocks in the XmlFragment `default`, and for the body the document's
 * attributes in `document-attrs-v1`. The server seeds it from the stored part
 * and reads it back on a save (`DocumentPageCrdt`).
 *
 * The editor holds every loaded part as one ProseMirror document, so no part
 * is bound to a view as a page room is. The binding is by value:
 *
 * - `publish`: the part's blocks as the view holds them are written into the
 *   room as a diff (`updateCanonicalYFragment`), which makes the room say
 *   exactly that. Publishing a view that has not drawn a change the room
 *   already holds would take that change out again, so:
 * - before a remote change is applied, `beforeRemote` must publish whatever
 *   the view holds that the room has not heard of; it answers false while the
 *   view cannot be read (an IME composition) and the change waits.
 * - then the change is applied and handed over with what the room said before
 *   it (`onRemote`), which the caller draws in the view at once. Between the
 *   two the view is never behind the room when something is published.
 *
 * Frames that arrive together are applied together: one publish, one draw.
 */

import * as Y from "yjs";
import { Fragment, type Node as ProseMirrorNode, type Schema } from "@tiptap/pm/model";
import { yXmlFragmentToProsemirrorJSON } from "@tiptap/y-tiptap";

import {
  COLLABORATIVE_DOCUMENT_ATTRS_MAP,
  publishCollaborativeDocumentAttrs,
  readCollaborativeDocumentAttrs,
} from "../../components/console/editor/native-document/collaborationDocumentAttrs";
import { noteForeignCanonicalXmlAttributes, updateCanonicalYFragment } from "./canonical-xml-attributes";
import { PartRoomConnection, type PartRoomConnectionOptions } from "./document-part-room";
import {
  canonicalDocumentJson,
  parseCanonicalDocumentJson,
  passesThroughCanonicalRead,
  withoutBareTopLevelInlineNodes,
  withoutUnsetPaginationRuntimeAttributes,
} from "./paged-document-projection";

/** The page room's fragment (the server's `DocumentPageCrdt`). */
const FRAGMENT_KEY = "default";
/** How long frames wait for the ones arriving with them. */
const FRAME_BATCH_MS = 30;
/** How soon frames held for the view (an IME composition) are tried again. */
const HELD_FRAME_RETRY_MS = 150;

/** This member's letters a remote change deleted (read before the room lets go of them). */
interface DeletedOwnLetters {
  readonly clock: number;
  readonly length: number;
  readonly text: string;
  readonly block: string | null;
  readonly origin: Y.ID | null;
  readonly parentText: unknown;
}

/** A run of this member's letters a change took out unseen, and where it goes back. */
export interface LostLetters {
  readonly block: string;
  /** In the block's inline content; null: at its end (its text run went). */
  readonly offset: number | null;
  text: string;
  readonly items?: readonly DeletedOwnLetters[];
}

export interface WordPartRoomOptions extends PartRoomConnectionOptions {
  /** The flat schema of the room's blocks (`canonicalDocumentSchema`). */
  readonly schema: Schema;
  /** Publish what the view holds that the room has not heard of; false when
   * it cannot be read now (the frames wait). */
  readonly beforeRemote: () => boolean;
  /** Someone else changed the part: its blocks before and after, as the room
   * says them; and, by block id, this member's own letters the change took
   * with it without naming them (they were typed into a text run another
   * member's change took out whole, before it had seen them). */
  readonly onRemote: (
    before: readonly unknown[],
    after: readonly unknown[],
    lost?: readonly LostLetters[],
    /** Where this member's caret is after the change: right after the letter
     * it stood after, followed through the room itself. */
    caret?: { readonly block: string; readonly offset: number },
  ) => void;
  /** This member's caret in the view: the block it is in and its offset in
   * the block's inline content (null when it is not in one block's text). */
  readonly caret?: () => { readonly block: string; readonly offset: number } | null;
  /** Someone else changed the document's attributes (the body's room only). */
  readonly onRemoteAttrs?: (attrs: Readonly<Record<string, unknown>>) => void;
  /** The room has been read, for the first time or afresh after the server
   * seeded it again. `blocks()` says what it holds. */
  readonly onRead: () => void;
}

export class WordPartRoom extends PartRoomConnection {
  private fragment!: Y.XmlFragment;
  private attrsMap!: Y.Map<unknown>;
  private frames: Uint8Array[] = [];
  private frameTimer: ReturnType<typeof setTimeout> | null = null;

  constructor(private readonly word: WordPartRoomOptions) {
    super(word);
    this.openDoc();
  }

  protected bindModel(doc: Y.Doc): void {
    this.fragment = doc.getXmlFragment(FRAGMENT_KEY);
    this.attrsMap = doc.getMap<unknown>(COLLABORATIVE_DOCUMENT_ATTRS_MAP);
    // Which room element holds which published block, kept from publish to
    // publish: a block published again as the very same node is known unchanged
    // at once. What another member changes is forgotten as it comes in.
    const meta = { mapping: new Map(), isOMark: new Map() };
    this.publishMeta = meta;
    doc.on("afterTransaction", (transaction: Y.Transaction) => {
      if (transaction.origin === this.localOrigin) return;
      for (const type of transaction.changedParentTypes.keys()) meta.mapping.delete(type);
      for (const type of transaction.changed.keys()) meta.mapping.delete(type);
    });
  }

  private publishMeta: { mapping: Map<unknown, unknown>; isOMark: Map<unknown, unknown> } = { mapping: new Map(), isOMark: new Map() };
  /** Each published block (the very object) as the room's model reads it. */
  private readonly blockNodes = new WeakMap<object, ProseMirrorNode>();

  /** The part as the room's model reads it: each block read once for the very
   * object it is (a part's unchanged blocks are the same objects publish after
   * publish), the whole part read at once when a block does not read alone. */
  private partNode(blocks: readonly unknown[]): ProseMirrorNode {
    const nodes: ProseMirrorNode[] = [];
    for (const block of blocks) {
      let node = block && typeof block === "object" ? this.blockNodes.get(block) : undefined;
      if (!node) {
        const one = parseCanonicalDocumentJson(this.word.schema, canonicalDocumentJson({ type: "doc", content: [block] }));
        if (one.childCount !== 1) {
          return parseCanonicalDocumentJson(this.word.schema, canonicalDocumentJson({ type: "doc", content: [...blocks] }));
        }
        node = one.child(0);
        if (block && typeof block === "object") this.blockNodes.set(block, node);
      }
      nodes.push(node);
    }
    return this.word.schema.topNodeType.create(null, Fragment.fromArray(nodes));
  }

  protected dropModel(): void {
    this.frames = [];
    if (this.frameTimer !== null) clearTimeout(this.frameTimer);
    this.frameTimer = null;
  }

  protected firstRead(): void {
    // Attributes a newer client wrote stay in the room through this one's
    // writes (the page room does the same with what its read leaves out).
    noteForeignCanonicalXmlAttributes(this.fragment, this.word.schema, passesThroughCanonicalRead);
    this.dirty = false;
    this.word.onRead();
  }

  /** What the room holds: its top-level blocks. */
  blocks(): unknown[] {
    const json = withoutUnsetPaginationRuntimeAttributes(yXmlFragmentToProsemirrorJSON(this.fragment));
    const read = withoutBareTopLevelInlineNodes(this.word.schema, json).json as { content?: unknown };
    return Array.isArray(read.content) ? read.content : [];
  }

  attrs(): Readonly<Record<string, unknown>> {
    return readCollaborativeDocumentAttrs(this.attrsMap);
  }

  /**
   * Make the room say `blocks` (and, for the body, `attrs`), as one change of
   * this member's. False while the room is not read or not this member's to
   * write; true also when that changed nothing. Throws when the blocks are not
   * a document the schema takes: nothing is written then.
   */
  publish(blocks: readonly unknown[], attrs?: Readonly<Record<string, unknown>>): boolean {
    if (!this.accepting) return false;
    const node = this.partNode(blocks);
    this.doc.transact(() => {
      updateCanonicalYFragment(this.doc, this.fragment, node, this.publishMeta as never);
      if (attrs) publishCollaborativeDocumentAttrs(this.attrsMap, attrs);
    }, this.localOrigin);
    return true;
  }

  /** The top-level element of a block (by `block_id`). */
  private elementOf(block: string): Y.XmlElement | null {
    for (const child of this.fragment.toArray()) {
      if (child instanceof Y.XmlElement && child.getAttribute("block_id") === block) return child;
    }
    return null;
  }

  /** A caret in a block's text as a position in the room that stays right
   * after the letter before it (null when it is not inside a text run). */
  private relativeCaret(caret: { block: string; offset: number }): Y.RelativePosition | null {
    const element = this.elementOf(caret.block);
    if (!element) return null;
    let remain = caret.offset;
    for (const child of element.toArray()) {
      const length = child instanceof Y.XmlText ? child.length : 1;
      if (child instanceof Y.XmlText && remain <= length) return Y.createRelativePositionFromTypeIndex(child, remain, -1);
      if (remain < length) return null;
      remain -= length;
    }
    return null;
  }

  /** The caret after a change, its letters written back counted: after its
   * own letter when that letter is one written back, else where the room has
   * it, moved on by the letters written back before it in its block. */
  private caretAfterChange(relative: Y.RelativePosition, block: string, lost: readonly LostLetters[]): { block: string; offset: number } | null {
    const anchor = relative.item;
    if (anchor && anchor.client === this.doc.clientID) {
      for (const entry of lost) {
        if (entry.block !== block) continue;
        let before = 0;
        for (const item of entry.items ?? []) {
          if (anchor.clock >= item.clock && anchor.clock < item.clock + item.length) {
            // Its own letter is written back: right after it (a run going at the block's end: the hook puts it there).
            return entry.offset === null ? null : { block, offset: entry.offset + before + (anchor.clock - item.clock) + 1 };
          }
          before += item.length;
        }
      }
    }
    const at = this.absoluteCaret(relative, block);
    if (!at) return null;
    let offset = at.offset;
    for (const entry of lost) if (entry.block === block && entry.offset !== null && entry.offset <= at.offset) offset += entry.text.length;
    return { block, offset };
  }

  private absoluteCaret(relative: Y.RelativePosition, block: string): { block: string; offset: number } | null {
    const absolute = Y.createAbsolutePositionFromRelativePosition(relative, this.doc);
    if (!absolute || !(absolute.type instanceof Y.XmlText) || absolute.type._item?.deleted) return null;
    const element = this.elementOf(block);
    if (!element || absolute.type.parent !== element) return null;
    let offset = 0;
    for (const child of element.toArray()) {
      if (child === absolute.type) return { block, offset: offset + absolute.index };
      offset += child instanceof Y.XmlText ? child.length : 1;
    }
    return null;
  }

  /** The `block_id` of the top-level block an item stands in. */
  private blockOf(item: Y.Item): string | null {
    let type = item.parent as Y.AbstractType<unknown> | null;
    let element: Y.XmlElement | null = null;
    while (type && type !== this.fragment) {
      if (type instanceof Y.XmlElement) element = type;
      const holder = type._item;
      type = holder ? (holder.parent as Y.AbstractType<unknown>) : null;
    }
    if (type !== this.fragment || !element) return null;
    const id = element.getAttribute("block_id");
    return typeof id === "string" && id ? id : null;
  }

  /**
   * Of this member's letters the frames deleted, those they did not name:
   * typed into a text run another member's change took out (emptying a
   * paragraph that holds an inline object, joining two runs) before it had
   * seen them, so the room dropped them with it. Each run of them, in the
   * order typed, with where it stood: right after the letter it was typed
   * after (or the nearest one before that is still there), as an offset in
   * the block's inline content; null when its whole text run went (it goes
   * at the block's end). A deletion that names them (another member selected
   * them) is that member's and stands.
   */
  private ownLettersLost(frames: readonly Uint8Array[], deleted: readonly DeletedOwnLetters[]): LostLetters[] {
    const lost: LostLetters[] = [];
    if (deleted.length === 0) return lost;
    const own = this.doc.clientID;
    const named: Array<{ clock: number; len: number }> = [];
    for (const frame of frames) {
      try {
        const { ds } = Y.decodeUpdate(frame) as unknown as { ds: { clients: Map<number, Array<{ clock: number; len: number }>> } };
        for (const range of ds.clients.get(own) ?? []) named.push(range);
      } catch {
        return lost;
      }
    }
    let group: (LostLetters & { end: number; items: DeletedOwnLetters[] }) | null = null;
    const groups: Array<LostLetters & { end: number; items: DeletedOwnLetters[] }> = [];
    for (const item of [...deleted].sort((a, b) => a.clock - b.clock)) {
      if (!item.block) continue;
      if (named.some((range) => item.clock < range.clock + range.len && range.clock < item.clock + item.length)) continue;
      if (group && group.block === item.block && group.end === item.clock) {
        group.text += item.text;
        group.end = item.clock + item.length;
        group.items.push(item);
        continue;
      }
      group = { block: item.block, offset: this.offsetAfter(item), text: item.text, end: item.clock + item.length, items: [item] };
      groups.push(group);
    }
    for (const entry of groups) lost.push({ block: entry.block, offset: entry.offset, text: entry.text, items: entry.items });
    return lost;
  }

  /** Where a deleted letter of a still-standing text run goes back: right
   * after the letter it was typed after, or the nearest one before it still
   * there (a relative position on its left letter), in the block's inline
   * content; null when its text run itself went. */
  private offsetAfter(item: DeletedOwnLetters): number | null {
    const text = item.parentText;
    if (!(text instanceof Y.XmlText) || text._item?.deleted) return null;
    const element = text.parent;
    if (!(element instanceof Y.XmlElement)) return null;
    let index = 0;
    if (item.origin) {
      const absolute = Y.createAbsolutePositionFromRelativePosition(new Y.RelativePosition(null, null, item.origin, -1), this.doc);
      if (!absolute || absolute.type !== text) return null;
      index = absolute.index;
    }
    let offset = 0;
    for (const child of element.toArray()) {
      if (child === text) return offset + index;
      offset += child instanceof Y.XmlText ? child.length : 1;
    }
    return null;
  }

  protected applyFrame(payload: Uint8Array): void {
    if (!this.seeded) {
      // The room's reading: nothing is drawn until `synced` says it is whole.
      super.applyFrame(payload);
      return;
    }
    // This member's own update sent back: nothing to apply, publish or draw
    // (each key's publish came back and read the whole part twice for nothing).
    if (this.isOwnEcho(payload)) return;
    // The socket's buffer is reused by nobody, but a frame outlives this call.
    this.frames.push(payload.slice());
    if (this.frameTimer === null) this.frameTimer = setTimeout(() => this.drainFrames(), FRAME_BATCH_MS);
  }

  private drainFrames(): void {
    this.frameTimer = null;
    if (this.frames.length === 0 || this.isDestroyed) return;
    if (!this.seeded) return;
    let ready = false;
    try {
      ready = this.word.beforeRemote();
    } catch (error) {
      console.warn("[part-room] the view could not be published before a change arrived", error);
    }
    if (!ready) {
      this.frameTimer = setTimeout(() => this.drainFrames(), HELD_FRAME_RETRY_MS);
      return;
    }
    const before = this.blocks();
    const attrsBefore = JSON.stringify(this.attrs());
    const frames = this.frames;
    this.frames = [];
    // This member's letters the frames delete, read before the room lets go of
    // them (a deleted run's letters are collected as the transaction ends).
    const caretHere = this.word.caret?.() ?? null;
    const caretRelative = caretHere ? this.relativeCaret(caretHere) : null;
    const deletedOwn: DeletedOwnLetters[] = [];
    const collect = (transaction: Y.Transaction) => {
      if (transaction.origin !== this.remoteOrigin) return;
      const own = this.doc.clientID;
      Y.iterateDeletedStructs(transaction, transaction.deleteSet, (struct) => {
        if (struct.id.client !== own || !(struct instanceof Y.Item) || !(struct.content instanceof Y.ContentString)) return;
        deletedOwn.push({ clock: struct.id.clock, length: struct.length, text: struct.content.str, block: this.blockOf(struct),
          origin: struct.origin, parentText: struct.parent });
      });
    };
    this.doc.on("beforeObserverCalls", collect);
    try {
      this.doc.transact(() => {
        for (const frame of frames) Y.applyUpdate(this.doc, frame, this.remoteOrigin);
      }, this.remoteOrigin);
    } finally {
      this.doc.off("beforeObserverCalls", collect);
    }
    const after = this.blocks();
    const lost = this.ownLettersLost(frames, deletedOwn);
    const caretAfter = caretHere && caretRelative ? this.caretAfterChange(caretRelative, caretHere.block, lost) : null;
    if (JSON.stringify(before) !== JSON.stringify(after)) this.word.onRemote(before, after, lost.length > 0 ? lost : undefined, caretAfter ?? undefined);
    const attrsAfter = this.attrs();
    if (JSON.stringify(attrsAfter) !== attrsBefore) this.word.onRemoteAttrs?.(attrsAfter);
  }
}
