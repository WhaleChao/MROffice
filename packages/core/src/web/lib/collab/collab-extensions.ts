"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * TipTap extension wiring for the collaborative editing foundation.
 *
 * Builds the `Collaboration` + `CollaborationCaret` extension pair bound to a
 * {@link CollabSession}'s Y.Doc + awareness. This is the ONLY place that imports
 * the TipTap collaboration packages, and it is called by NoteEditor ONLY when
 * the `USE_COLLAB_EDITING` flag is ON — so the OFF path never loads these
 * extensions and the editor's existing history/content path is unchanged.
 *
 * IMPORTANT: TipTap's `Collaboration` extension replaces the default ProseMirror
 * history (undo/redo is driven by the Yjs undo manager). The caller therefore
 * disables StarterKit's `history` ONLY when collab is on (see NoteEditor); this
 * module assumes the history has already been removed from the base extension
 * list when these are appended.
 */

import { MappablePosition, getUpdatedPosition, type AnyExtension, type Content, type Editor } from "@tiptap/core";
import Collaboration, { type CollaborationOptions, type CollaborationStorage } from "@tiptap/extension-collaboration";
import CollaborationCaret from "@tiptap/extension-collaboration-caret";
import { initProseMirrorDoc, yCursorPlugin, yUndoPluginKey, yCursorPluginKey, ySyncPluginKey } from "@tiptap/y-tiptap";
import type { JSONContent } from "@tiptap/core";
import type { Awareness } from "y-protocols/awareness";
import { Plugin, type Transaction } from "@tiptap/pm/state";
import * as Y from "yjs";
import type { CollabSession } from "./collab-provider";
import { installBranchSafeHistory } from "./branch-safe-undo";
import { isLocalUsersOtherConnection, sanitizeCollabUser } from "./collab-state";
import { cachedSelfCollabIdentity } from "./self-identity";
import { PagedDocumentBridge, PAGED_CANONICAL_REMOTE_META, type PagedDocumentBridgeIssue } from "./paged-document-bridge";
import { MAX_DRAWN_CARETS } from "./peer-caret-overlay";
import { COLLABORATIVE_DOCUMENT_ATTRS_REMOTE_META } from "../../components/console/editor/native-document/collaborationDocumentAttrs";
import { REBALANCE_META } from "../../components/console/editor/native-document/paged/pageRebalanceExtension";
import { SAVED_WITHOUT_HISTORY_META } from "../../components/console/editor/native-document/readOnlyTransactionAuthority";
import { MRNOTE_BLOCK_IDENTITY_META } from "../../components/console/editor/native-document/blockIds";

const initialDocuments = new WeakMap<
  Editor,
  ReturnType<typeof initProseMirrorDoc>
>();

function whiteboardCreationHistoryGuard() {
  return new Plugin({
    view(view) {
      const manager = yUndoPluginKey.getState(view.state)?.undoManager as Y.UndoManager | undefined;
      if (!manager) return {};
      // An undo must take back the whole of what its step inserted, also the
      
      const restoreHistory = installBranchSafeHistory(manager);
      const doc = manager.doc;
      const originalFilter = manager.deleteFilter;
      let insertions: ReturnType<typeof Y.createDeleteSet> | null = null;
      const before = (transaction: Y.Transaction) => {
        if (transaction.origin !== manager) return;
        const stack = manager.undoing ? manager.undoStack : manager.redoStack;
        insertions = stack.at(-1)?.insertions ?? null;
      };
      const after = (transaction: Y.Transaction) => {
        if (transaction.origin === manager) insertions = null;
      };
      type HistoryType = Pick<Y.AbstractType<never>, "_map" | "_item">;
      const hasPeerContent = (type: HistoryType): boolean => {
        for (const item of type._map.values()) {
          if (!item.deleted && item.id.client !== doc.clientID) return true;
        }
        if (type instanceof Y.XmlElement) {
          for (const child of type.toArray()) {
            if (child._item && !child._item.deleted && child._item.id.client !== doc.clientID) return true;
            if (hasPeerContent(child)) return true;
          }
        }
        return false;
      };
      const filter = (item: Y.Item): boolean => {
        if (insertions) {
          let type: HistoryType | null = item.content instanceof Y.ContentType
            ? item.content.type : item.parent instanceof Y.AbstractType ? item.parent : null;
          while (type) {
            const creation = type._item;
            if (type instanceof Y.XmlElement && creation &&
                (type.nodeName === "whiteboardObject" || type.nodeName === "whiteboard") &&
                Y.isDeleted(insertions, creation.id) && hasPeerContent(type)) return false;
            type = creation?.parent instanceof Y.AbstractType ? creation.parent : null;
          }
        }
        return originalFilter(item);
      };
      doc.on("beforeTransaction", before);
      doc.on("afterTransaction", after);
      manager.deleteFilter = filter;
      return { destroy() {
        doc.off("beforeTransaction", before);
        doc.off("afterTransaction", after);
        if (manager.deleteFilter === filter) manager.deleteFilter = originalFilter;
        restoreHistory();
      } };
    },
  });
}

const InitializedCollaboration = Collaboration.extend({
  onBeforeCreate(event) {
    this.parent?.(event);
    const fragment =
      this.options.fragment ??
      this.options.document?.getXmlFragment(this.options.field ?? "default");
    if (!fragment?.length) return;
    // Start from the room itself, not a saved JSON seed or an empty editor.
    // The upstream fallback rebuilds the whole document then clamps the old
    // numeric caret into it; in a paged document that can point at pageBlock
    // instead of inline content. Its supported initial mapping avoids that
    // replacement entirely and preserves legal image/node selections too.
    const initial = initProseMirrorDoc(fragment, this.editor.schema);
    if (!initial.doc.type.validContent(initial.doc.content)) return;
    // TipTap's createNodeFromContent accepts a PM Node at runtime. Keep the
    // exact nodes used by the mapping; a JSON copy would discard that identity.
    this.editor.options.content = initial.doc as unknown as Content;
    initialDocuments.set(this.editor, initial);
  },
  addProseMirrorPlugins() {
    const initial = initialDocuments.get(this.editor);
    initialDocuments.delete(this.editor);
    // Options are freshly resolved for each TipTap lifecycle hook. Supply the
    // map in this plugin-construction context, shared by the parent factory.
    // A later remount must not reuse the original bootstrap receipt.
    if (initial && this.editor.state.doc === initial.doc) {
      this.options.ySyncOptions = {
        ...this.options.ySyncOptions,
        mapping: initial.mapping,
      };
    }
    return [...(this.parent?.() ?? []), whiteboardCreationHistoryGuard()];
  },
});

export interface CollabExtensionOptions {
  session: CollabSession;
  /** Local collaborator identity rendered on the caret label. */
  user: { name?: unknown; color?: unknown };
  /** Y.Doc fragment field name. Defaults to TipTap's `"default"`. */
  field?: string;
  /** Explicitly selected only after the native consumer negotiates canonical originals. */
  canonical?: boolean;
  mayWrite?: () => boolean;
  canonicalNetworkModel?: () => boolean;
  /** See PagedDocumentBridgeOptions.durableOriginals. */
  durableOriginals?: boolean;
  onPagedIssue?: (issue: PagedDocumentBridgeIssue | null) => void;
  /** With `canonical`: the legacy page room's logical binding (see
   * PagedDocumentBridgeOptions.upgradeLegacyRoom). */
  upgradeLegacyRoom?: boolean;
  /** See PagedDocumentBridgeOptions.initialLayoutReceipt. */
  initialLayoutReceipt?: (canonical?: JSONContent, root?: object) => unknown;
  /** See PagedDocumentBridgeOptions.initialLayoutSizes. */
  initialLayoutSizes?: () => number[] | null;
  /** See PagedDocumentBridgeOptions.initialLayoutReceiptPending. */
  initialLayoutReceiptPending?: () => boolean;
}

/** Route the existing native attribute publisher through the canonical
 * view's personal history. Hydration/repair callers pass historical:false. */
export function publishCanonicalDocumentAttrs(
  editor: Editor, publish: () => boolean, options: { historical: boolean },
): boolean {
  const bridge = (editor.storage.collaboration as (CollaborationStorage & {
    pagedBridge?: PagedDocumentBridge | null;
  }) | undefined)?.pagedBridge;
  return bridge ? bridge.publishDocumentAttrs(publish, options) : publish();
}

/** Explicit Save materializes only verified view maintenance, never author
 * content that failed to reach its original collaborative model. */
export function prepareCanonicalDocumentExplicitSave(editor: Editor): boolean {
  const bridge = (editor.storage.collaboration as (CollaborationStorage & {
    pagedBridge?: PagedDocumentBridge | null;
  }) | undefined)?.pagedBridge;
  return bridge?.prepareExplicitSave() ?? false;
}

/** Merge a draft this tab kept but could not save into the page's room and
 * view (PagedDocumentBridge.adoptRecoveredContent); false until the bridge
 * can take it. */
export function adoptRecoveredDocumentContent(editor: Editor, content: unknown, sameBase: boolean, baseContent?: unknown,
  options?: { baseOnlyIfHeld?: boolean; fallbackBaseContent?: unknown }): boolean {
  const bridge = (editor.storage.collaboration as (CollaborationStorage & {
    pagedBridge?: PagedDocumentBridge | null;
  }) | undefined)?.pagedBridge;
  return bridge?.adoptRecoveredContent(content, sameBase, baseContent, options) ?? false;
}

/** The pen was picked up: PagedDocumentBridge.warmForFirstEdit. */
export function warmDocumentRoomForFirstEdit(editor: Editor): void {
  (editor.storage?.collaboration as (CollaborationStorage & {
    pagedBridge?: PagedDocumentBridge | null;
  }) | undefined)?.pagedBridge?.warmForFirstEdit();
}

/** Whether the page's room has been read into the view (false without one). */
export function documentRoomEstablished(editor: Editor): boolean {
  const bridge = (editor.storage?.collaboration as (CollaborationStorage & {
    pagedBridge?: PagedDocumentBridge | null;
  }) | undefined)?.pagedBridge;
  return bridge?.isEstablished() ?? false;
}

/** The blocks merges into the room kept twice since last asked (scheme C). */
export function takeDocumentRecoveryConflicts(editor: Editor): Array<{ kept: string | null; copy: string | null }> {
  const bridge = (editor.storage?.collaboration as (CollaborationStorage & {
    pagedBridge?: PagedDocumentBridge | null;
  }) | undefined)?.pagedBridge;
  return bridge?.takeRecoveryConflicts() ?? [];
}

/** Why the last prepareCanonicalDocumentExplicitSave refused, as a
 * content-free code for save diagnostics; `no_bridge` without a canonical bridge. */
export function canonicalDocumentExplicitSaveRefusal(editor: Editor): string | null {
  const bridge = (editor.storage.collaboration as (CollaborationStorage & {
    pagedBridge?: PagedDocumentBridge | null;
  }) | undefined)?.pagedBridge;
  return bridge ? bridge.lastExplicitSaveRefusal() : "no_bridge";
}

/** Attribute observers run synchronously inside Undo/Redo. The consumer can
 * persist this view's history operation while skipping peer-only updates. */
export function isCanonicalDocumentHistoryTransaction(editor: Editor): boolean {
  return Boolean((editor.storage.collaboration as (CollaborationStorage & {
    pagedBridge?: PagedDocumentBridge | null;
  }) | undefined)?.pagedBridge?.isApplyingHistory());
}

/** TipTap emits one update for the complete appendTransaction batch. A
 * selection-only root may have triggered non-history identity maintenance;
 * persisting that view would bypass the canonical bridge's authoring rules. */
export function hasCanonicalDocumentAuthoringTransaction(
  transaction: Transaction | undefined,
  appendedTransactions: readonly Transaction[] = [],
): boolean {
  return [transaction, ...appendedTransactions].some(tr => tr?.docChanged &&
    (tr.getMeta("addToHistory") !== false || tr.getMeta(SAVED_WITHOUT_HISTORY_META) === true) &&
    !tr.getMeta(REBALANCE_META) &&
    !tr.getMeta(PAGED_CANONICAL_REMOTE_META) &&
    !tr.getMeta(COLLABORATIVE_DOCUMENT_ATTRS_REMOTE_META));
}

/**
 * The legacy page room delivered this update: a collaborator's change (or the
 * shared document-attribute map) projected into this view, followed only by
 * the closed, non-undoable maintenance plugins append to any document change.
 * It is not this member's authoring. Local typing, commands and this view's
 * own Yjs Undo/Redo (`isUndoRedoOperation`) are.
 */
export function isLegacyRoomProjectionUpdate(
  transaction: Transaction | undefined,
  appendedTransactions: readonly Transaction[] = [],
): boolean {
  if (!transaction?.docChanged) return false;
  const sync = transaction.getMeta(ySyncPluginKey) as
    | { isChangeOrigin?: unknown; isUndoRedoOperation?: unknown }
    | undefined;
  const delivered =
    (sync?.isChangeOrigin === true && sync.isUndoRedoOperation !== true) ||
    transaction.getMeta(COLLABORATIVE_DOCUMENT_ATTRS_REMOTE_META) === true;
  return delivered && appendedTransactions.every((tr) => !tr.docChanged ||
    tr.getMeta("addToHistory") === false || Boolean(tr.getMeta(REBALANCE_META)) ||
    tr.getMeta(MRNOTE_BLOCK_IDENTITY_META) === true);
}

/**
 * Render a remote collaborator's caret + name label. The user fields are
 * re-sanitized here (defense in depth: the awareness payload is peer-supplied),
 * so the color only ever reaches the DOM as a validated `#rrggbb` value and the
 * name is set via `textContent` (never `innerHTML`), preventing markup injection.
 */
function renderCollabCaret(rawUser: Record<string, unknown>): HTMLElement {
  const user = sanitizeCollabUser(
    rawUser as { name?: unknown; color?: unknown },
  );

  const caret = document.createElement("span");
  caret.classList.add("collab-caret");
  caret.setAttribute("style", `border-color: ${user.color}`);

  const label = document.createElement("div");
  label.classList.add("collab-caret__label");
  label.setAttribute("style", `background-color: ${user.color}`);
  // textContent (not innerHTML) — the name is untrusted peer input.
  label.textContent = user.name;

  caret.appendChild(label);
  return caret;
}

export { MAX_DRAWN_CARETS };

/** The caret filter for one editor: never its own, nor its user's other
 * connections, and only the `MAX_DRAWN_CARETS` most recently moved once the
 * room holds more. */
export function recentCaretFilter(awareness: Awareness, limit = MAX_DRAWN_CARETS): {
  filter: (currentClientId: number, userClientId: number, state: unknown) => boolean;
  dispose: () => void;
} {
  const movedAt = new Map<number, number>();
  let tick = 0;
  let drawn: Set<number> | null = null;
  let drawnAt = -1;
  const onChange = ({ added, updated, removed }: { added: number[]; updated: number[]; removed: number[] }) => {
    for (const id of [...added, ...updated]) movedAt.set(id, ++tick);
    for (const id of removed) movedAt.delete(id);
  };
  awareness.on("change", onChange);
  return {
    filter(currentClientId, userClientId, state) {
      if (currentClientId === userClientId) return false;
      if (isLocalUsersOtherConnection(
        awareness.getStates() as Map<number, Record<string, unknown>>, currentClientId, state,
      )) return false;
      if (awareness.getStates().size - 1 <= limit) return true;
      if (drawn === null || drawnAt !== tick) {
        drawn = new Set([...movedAt.entries()]
          .filter(([id]) => id !== currentClientId)
          .sort((left, right) => right[1] - left[1])
          .slice(0, limit)
          .map(([id]) => id));
        drawnAt = tick;
      }
      return drawn.has(userClientId);
    },
    dispose() { awareness.off("change", onChange); },
  };
}

/**
 * Build the collaboration extension pair for a live session. Returns an empty
 * array if (defensively) no session/doc is available, so a caller that flips the
 * flag on before a session exists simply gets the non-collab editor.
 */
export function buildCollabExtensions(
  options: CollabExtensionOptions | null | undefined,
): AnyExtension[] {
  if (!options?.session?.ydoc) return [];
  const { session, user, field } = options;
  // Configure the caret with the RESOLVED profile whenever it is already
  // known. This matters because CollaborationCaret re-stamps its CONFIGURED
  // user into awareness on every plugin-view init (every editor rebuild), so
  // a placeholder here becomes what peers see — the session's identity guard
  // repairs it, but starting correct avoids the visible flash entirely.
  // Falls back to awareness (already-published identity), then the caller's
  // placeholder.
  const cached = cachedSelfCollabIdentity();
  const live = session.awareness.getLocalState()?.user as
    { name?: unknown; color?: unknown } | undefined;
  const safeUser = sanitizeCollabUser(
    cached?.name
      ? { name: cached.name, color: cached.color }
      : typeof live?.name === "string" && live.name.trim()
        ? live
        : user,
  );

  const collaboration = options.canonical ? InitializedCollaboration.extend<CollaborationOptions, CollaborationStorage & { pagedBridge: PagedDocumentBridge | null }>({
    addStorage() { return { isDisabled: false, pagedBridge: null }; },
    onBeforeCreate(event) {
      if (!this.editor.schema.nodes.pageBlock) { this.parent?.(event); return; }
      this.editor.utils.createMappablePosition = position => this.storage.pagedBridge?.createMappablePosition(position) ?? new MappablePosition(position);
      this.editor.utils.getUpdatedPosition = (position, transaction) => this.storage.pagedBridge?.getUpdatedPosition(position, transaction) ?? getUpdatedPosition(position, transaction);
    },
    addProseMirrorPlugins() {
      if (!this.editor.schema.nodes.pageBlock) return this.parent?.() ?? [];
      const bridge = new PagedDocumentBridge({
        session, schema: this.editor.schema, editor: this.editor, field,
        mayWrite: options.mayWrite,
        canonicalNetworkModel: options.canonicalNetworkModel,
        durableOriginals: options.durableOriginals,
        onIssue: options.onPagedIssue,
        upgradeLegacyRoom: options.upgradeLegacyRoom,
        initialLayoutReceipt: options.initialLayoutReceipt,
        initialLayoutSizes: options.initialLayoutSizes,
        initialLayoutReceiptPending: options.initialLayoutReceiptPending,
      });
      this.storage.pagedBridge = bridge;
      return [...bridge.plugins, whiteboardCreationHistoryGuard()];
    },
    addCommands() {
      const inherited = this.parent?.();
      return {
        ...inherited,
        undo: () => props => {
          const bridge = this.storage.pagedBridge;
          if (!bridge) return inherited?.undo?.()(props) ?? false;
          props.tr.setMeta("preventDispatch", true);
          return bridge.history("undo", Boolean(props.dispatch));
        },
        redo: () => props => {
          const bridge = this.storage.pagedBridge;
          if (!bridge) return inherited?.redo?.()(props) ?? false;
          props.tr.setMeta("preventDispatch", true);
          return bridge.history("redo", Boolean(props.dispatch));
        },
      };
    },
    onDestroy() { this.storage.pagedBridge?.destroy(); this.storage.pagedBridge = null; },
  }) : InitializedCollaboration;
  const caret = (options.canonical ? CollaborationCaret.extend({
    addProseMirrorPlugins() {
      const plugins = this.parent?.() ?? [];
      return this.editor.schema.nodes.pageBlock
        ? plugins.filter(plugin => plugin.spec.key !== yCursorPluginKey)
        : plugins;
    },
  }) : CollaborationCaret).extend({
    addProseMirrorPlugins() {
      const plugins = this.parent?.() ?? [];
      const index = plugins.findIndex(plugin => plugin.spec.key === yCursorPluginKey);
      if (index < 0) return plugins;
      // The same caret plugin, drawing only the most recently moved carets.
      const recent = recentCaretFilter(session.awareness);
      const cursor = yCursorPlugin(session.awareness, {
        cursorBuilder: renderCollabCaret,
        selectionBuilder: this.options.selectionRender,
        awarenessStateFilter: recent.filter,
      });
      const release = new Plugin({ view: () => ({ destroy: recent.dispose }) });
      return [...plugins.slice(0, index), cursor, release, ...plugins.slice(index + 1)];
    },
  });
  return [
    collaboration.configure({
      document: session.ydoc,
      ...(field ? { field } : {}),
    }),
    caret.configure({
      provider: { awareness: session.awareness },
      user: { name: safeUser.name, color: safeUser.color },
      render: renderCollabCaret,
    }),
  ];
}
