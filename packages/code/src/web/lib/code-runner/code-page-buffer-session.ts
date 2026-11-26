/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * One in-browser text model and one autosave writer for every open standalone
 * Code page.
 *
 * A page can be shown by more than one `CodeWindow` (`force_new` duplicate
 * view). Two independent React buffers are unsafe: each autosave hook owns a
 * full-document draft and a content-version base, so two edits made inside the
 * debounce window can race even though both windows belong to one browser
 * session.
 *
 * This registry gives those windows the model/editor split used by mature code
 * editors:
 *
 * - one page-and-authority-keyed Y.Text is the live text model;
 * - every CodeMirror view binds to that same model;
 * - one shared Y.UndoManager owns text history across the views;
 * - the first live holder is the autosave writer, while every holder can submit
 *   a draft or invoke save/conflict actions through that writer;
 * - closing a holder is idempotent, and the final release destroys the local
 *   Yjs session.
 *
 * A known writer uses `createCollabSession({docId})`, leaving the app's
 * auto-wiring eligible so the Y.Text can join the live server room. Unknown
 * authority and the collaboration kill switch stay local-only. An ACL viewer
 * gets TWO fresh documents: a durable local projection for first paint and a
 * clean, never-hydrated room observer that replaces the editor binding only
 * after the room actually delivers state. Three consequences are load-bearing:
 *
 * - the REST-loaded page is NOT automatically canonical. A room that replays
 *   peers' unsaved work owns the text, and hydration must wait for the seeding
 *   election before deciding to seed or adopt (see {@link settleRoom});
 * - a peer's keystroke arriving over the room is not this browser's edit. It
 *   must reach every view without turning a clean buffer into an autosave
 *   writer. An already-dirty local draft must, however, include the merged
 *   text before its pending or in-flight save can finish.
 * - grant/revoke changes retire the old authority generation. A new Y.Doc and
 *   socket are mandatory; carrying a viewer's durable projection into a later
 *   writable room would merge two independent copies of the same source.
 */

import * as Y from "yjs";
import { diff } from "@codemirror/merge";

import {
  createCollabSession,
  type CollabSession,
} from "@mrnote/core/web/lib/collab/collab-provider.ts";
import {
  CODE_COLLAB_TEXT_KEY,
  getCodeYText,
} from "@mrnote/core/web/lib/collab/code-collab.ts";
import { jsonEquals, spliceYText, type JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import { MAX_CODE_SOURCE_CHARS } from "./code-file-import";
import { codeLanguageId } from "@mrnote/core/web/lib/code-runner/languages.ts";
import type {
  DiscardedChangesKeep,
  PageDocumentDraft,
  PageSaveConflictDetails,
  PageSaveStatus,
} from "@mrnote/core/web/hooks/useAutosavedPageDocument.ts";
import type { PageItem } from "@mrnote/core/web/lib/notebook-sdk.ts";
import type { PageContentDocument } from "@mrnote/core/web/hooks/usePageContentSync.ts";
import type { PageDocumentCollaborationAdapter } from "@/lib/collab/page-document-adapter";
import { createCodePageCausalAdapter } from "./code-page-causal-adapter";
import {
  applyCodeTextChanges,
  codeTextHunks,
  normalizeCodeTextChanges,
  writeCodeTextChanges,
  type CodeTextChange,
  type CodeTextChangesAgainst,
} from "./code-text-changes";

export interface CodePageBufferSaveState {
  saveStatus: PageSaveStatus;
  isDirty: boolean;
  isSaving: boolean;
  hasConflict: boolean;
  error: string | null;
  conflictDetails: PageSaveConflictDetails | null;
  contentVersion: number | null;
  /** The error is a save the writer retries on its own (a dropped
   * connection), not a refusal: views report it without taking focus. */
  retryableFailure?: boolean;
  /** Not yet written, but the page's live room holds every edit
   * (useAutosavedPageDocument syncedToServer): views say "Synced". */
  syncedToServer?: boolean;
}

export interface CodePageBufferSaveController {
  queueSave(draft: PageDocumentDraft): void;
  replacePendingDraft(draft: PageDocumentDraft | null): void;
  saveNow(draft?: PageDocumentDraft): Promise<PageItem | null>;
  retrySave(): Promise<PageItem | null>;
  overwriteRemote(): Promise<PageItem | null>;
  saveConflictAsCopy(): Promise<PageItem | null>;
  reload(): Promise<PageItem | null>;
  discardLocalAndReload(options?: { keepAside?: DiscardedChangesKeep }): Promise<PageItem | null>;
}

export interface CodePageBufferAuthority {
  /** False is the global collaboration kill switch, not a permission verdict. */
  collaborationEnabled: boolean;
  /** False until the page ACL/tier has been fetched at least once. */
  known: boolean;
  /** Page tier + notebook role + checkout verdict. */
  writable: boolean;
  reason: "read_only_tier" | "role" | "locked" | "missing" | null;
  /** Stable, last-verified editor binding; freshness must not change it. */
  binding?: CodePageBufferBinding;
  /** Current fail-closed write gate. False pauses writes in the same binding. */
  mayWriteNow?: boolean;
}

export type CodePageBufferBinding =
  | "local"
  | "participant"
  | "participant-locked"
  | "acl-viewer";

type CodePageBufferAuthorityMode = "local" | "participant" | "acl-viewer";

interface CodePageBufferHolder {
  token: string;
  buildDraft(content: string): PageDocumentDraft;
  applyText(content: string): void;
  /** A peer's edit replayed by the room. The editor view already shows it; a
   * holder may take it like its own typing (current at once in its refs, its
   * window rendered once a burst settles). Falls back to applyText. */
  applyPeerText?(content: string): void;
  applyDraft(draft: PageDocumentDraft): void;
  controller: CodePageBufferSaveController | null;
  saveState: CodePageBufferSaveState;
  editorOrigins: Set<object>;
  listeners: Set<() => void>;
}

interface CodePageBufferEntry {
  causal?: ReturnType<typeof createCodePageCausalAdapter>;
  pageId: string;
  disposed: boolean;
  registryKey: string;
  authorityGeneration: number;
  authorityBinding: CodePageBufferBinding;
  authorityMode: CodePageBufferAuthorityMode;
  authorityWritable: boolean;
  authorityDesiredWritable: boolean;
  authorizationIntentGeneration: number;
  authorizationRefresh: Promise<void> | null;
  transportWritable: boolean;
  bindingGeneration: number;
  session: CollabSession;
  text: Y.Text;
  /** Last text accepted from hydration, a peer, or an authorized local edit. */
  acceptedText: string;
  undoManager: Y.UndoManager;
  sessions: Set<CollabSession>;
  undoManagers: Set<Y.UndoManager>;
  observerSession: CollabSession | null;
  observerAdopted: boolean;
  observerUpdateListener: (() => void) | null;
  textObserver: ((event: Y.YTextEvent, transaction: Y.Transaction) => void) | null;
  transportUnsubscribe: (() => void) | null;
  connectionDiagnosticUnsubscribes: Set<() => void>;
  connectionFailureLogged: boolean;
  holders: Map<string, CodePageBufferHolder>;
  ownerToken: string | null;
  originTokens: WeakMap<object, string>;
  listeners: Set<() => void>;
  initialized: boolean;
  acceptNextHydration: boolean;
  serverVersion: number;
  revision: number;
  lastQueuedRevision: number;
  dirty: boolean;
  latestDraft: PageDocumentDraft | null;
  saveState: CodePageBufferSaveState;
  suppressTextObserver: number;
  restoreUnauthorizedTextScheduled: boolean;
  /**
   * True while a NETWORKED session's seeding election is still open. Until it
   * closes the Y.Text may be about to receive the room's replay, so nothing
   * may write the REST-loaded text over it.
   */
  roomPending: boolean;
  /**
   * The election closed with an empty room, so the EDITOR owns the seed — it
   * seeds from its live view, which includes anything typed while the room was
   * settling. The next local write is therefore that seed, not a user edit,
   * and must not be counted as an unsaved change.
   */
  awaitingSeed: boolean;
}

const DEFAULT_SAVE_STATE: CodePageBufferSaveState = {
  saveStatus: "saved",
  isDirty: false,
  isSaving: false,
  hasConflict: false,
  error: null,
  conflictDetails: null,
  contentVersion: null,
};

const _entries = new Map<string, CodePageBufferEntry>();
const _pageAuthorities = new Map<
  string,
  {
    binding: CodePageBufferBinding;
    generation: number;
  }
>();
let _holderSequence = 0;
let _authoritySequence = 0;

const CODE_PAGE_CONNECTION_STAGES = new Set([
  "start",
  "ticket",
  "snapshot",
  "socket",
  "socket-open",
  "hello",
  "replay",
  "ready",
]);
const CODE_PAGE_CONNECTION_FAILURE_LOG = "[code-collab] connection-failed";
const CODE_PAGE_EMPTY_SEED_ORIGIN = Symbol("code-page-empty-seed");

function authorityBinding(
  authority: CodePageBufferAuthority,
): CodePageBufferBinding {
  if (!authority.collaborationEnabled) return "local";
  if (authority.binding) return authority.binding;
  if (!authority.known) return "local";
  if (authority.reason === "locked") return "participant-locked";
  if (!authority.writable) {
    return "acl-viewer";
  }
  return "participant";
}

function authorityMode(
  authority: CodePageBufferAuthority,
): CodePageBufferAuthorityMode {
  const binding = authorityBinding(authority);
  return binding === "participant" || binding === "participant-locked"
    ? "participant"
    : binding;
}

function authorityMayWrite(authority: CodePageBufferAuthority): boolean {
  // The collaboration switch selects transport, never ACL. Missing authority
  // is not permission: local-only mode stays read-only until the page verdict
  // is known, exactly like the networked path.
  const snapshotAllowsWrite =
    authority.mayWriteNow ?? (authority.known && authority.writable);
  if (!authority.collaborationEnabled) return snapshotAllowsWrite;
  return authorityBinding(authority) === "participant" && snapshotAllowsWrite;
}

function entryMayWrite(entry: CodePageBufferEntry): boolean {
  if (entry.causal?.active()) return entry.causal.mayWrite();
  return entryAuthorityMayWrite(entry);
}

function entryAuthorityMayWrite(entry: CodePageBufferEntry): boolean {
  if (!entry.authorityWritable) return false;
  const network = entry.session.network;
  if (!network || network.failed) return true;
  // A dropped connection is not a refusal (the causal path's canWrite reads
  // it the same way): once the room has delivered its state, typing goes on
  // into the room's document and the provider sends it when the socket is
  // back, where the server judges it. A refusal arrives while connected.
  // Treating the drop as read-only reverted every key typed offline.
  if (!network.connected && !entry.roomPending) return true;
  return entry.transportWritable && network.writable;
}

function draftContent(draft: PageDocumentDraft): string {
  const content = draft.content_json?.content;
  return typeof content === "string" ? content : "";
}

function draftWithContent(
  draft: PageDocumentDraft,
  content: string,
): PageDocumentDraft {
  return {
    ...draft,
    content_json: {
      ...(draft.content_json ?? {}),
      content,
    },
  };
}

function ownerOf(
  entry: CodePageBufferEntry,
): CodePageBufferHolder | null {
  return entry.ownerToken
    ? (entry.holders.get(entry.ownerToken) ?? null)
    : null;
}

function emit(entry: CodePageBufferEntry): void {
  if (entry.disposed) return;
  for (const listener of [...entry.listeners]) listener();
}

function saveStateEquals(
  left: CodePageBufferSaveState,
  right: CodePageBufferSaveState,
): boolean {
  return (
    left.saveStatus === right.saveStatus &&
    left.isDirty === right.isDirty &&
    left.isSaving === right.isSaving &&
    left.hasConflict === right.hasConflict &&
    left.error === right.error &&
    left.conflictDetails === right.conflictDetails &&
    left.contentVersion === right.contentVersion &&
    Boolean(left.retryableFailure) === Boolean(right.retryableFailure)
  );
}

function setTextWithoutSave(
  entry: CodePageBufferEntry,
  content: string,
): void {
  entry.acceptedText = content;
  if (entry.text.toString() === content) return;
  entry.suppressTextObserver += 1;
  try {
    // Only the changed span: deleting and re-inserting the whole text gave
    // every character a new identity, so a collaborator's later insert
    // anchored to the old characters landed at the top of the file (an
    // offline member's typing, restored while it could not be written, put
    
    entry.session.ydoc.transact(() => {
      spliceYText(entry.text, content);
    }, entry);
  } finally {
    entry.suppressTextObserver -= 1;
  }
}

/**
 * The changes that write `next` over the Y.Text's `current` string as this
 * browser's own edit, or null when it must not be written.
 *
 * An edit that says where it was made is written at exactly those places
 * (they must describe `current` -> `next`; positions computed against an
 * older text are no positions). Without positions the edit is placed hunk by
 * hunk. A hunk whose letters meet the same letter at its edge ("(a" typed
 * before "(" reads as "a(" after it) is placed by a guess, and in a room a
 * collaborator's key typed at that moment then comes out inside the run
 * (TRIAGE 47, Code): such a write is refused. A local-only text has no one
 * typing at the same moment, so there the guess changes nothing anyone sees.
 */
function authoredTextChanges(
  entry: CodePageBufferEntry,
  current: string,
  next: string,
  positions?: CodeTextChangesAgainst,
): CodeTextChange[] | null {
  if (positions) {
    if (positions.expected !== current) return null;
    const changes = normalizeCodeTextChanges(current, positions.changes);
    return changes && applyCodeTextChanges(current, changes) === next
      ? changes
      : null;
  }
  const { hunks, ambiguous } = codeTextHunks(current, next);
  return ambiguous && entry.session.network ? null : hunks;
}

/** Write this browser's own edit (`authoredTextChanges`) into the Y.Text. */
function writeAuthoredText(
  entry: CodePageBufferEntry,
  content: string,
  changes: readonly CodeTextChange[],
): void {
  entry.acceptedText = content;
  entry.suppressTextObserver += 1;
  try {
    entry.session.ydoc.transact(() => {
      writeCodeTextChanges(entry.text, changes);
    }, entry);
  } finally {
    entry.suppressTextObserver -= 1;
  }
}

/**
 * A refused edit leaves the text as the room has it: every view (the one that
 * already shows the edit included) is handed that text back.
 */
function refuseAuthoredText(entry: CodePageBufferEntry): void {
  const current = entry.text.toString();
  for (const holder of entry.holders.values()) holder.applyText(current);
  emit(entry);
}

/**
 * An empty Y.Text has no Yjs structs, so the Page V2 provider has nothing to
 * snapshot after electing this browser as the first seeder. Materialize one
 * invisible Y.Text struct only after the durable page is known-empty and the
 * current room explicitly elected this writable client. A single transaction
 * inserts and removes the marker under a non-undo origin, while the local text
 * observer is suppressed. The editor and autosave therefore still see "".
 */
function materializeEmptyFirstSeed(entry: CodePageBufferEntry): void {
  const network = entry.session.network;
  if (
    entry.disposed ||
    entry.causal?.active() ||
    !entry.awaitingSeed ||
    !entry.initialized ||
    !entryMayWrite(entry) ||
    !network ||
    network.failed ||
    !network.connected ||
    network.first !== true ||
    network.hasState ||
    entry.text.length !== 0 ||
    draftContent(entry.latestDraft ?? {}).length !== 0 ||
    entry.session.ydoc.store.clients.size !== 0
  ) {
    return;
  }
  entry.suppressTextObserver += 1;
  try {
    entry.session.ydoc.transact(() => {
      entry.text.insert(0, "\u200b");
      entry.text.delete(0, 1);
    }, CODE_PAGE_EMPTY_SEED_ORIGIN);
  } finally {
    entry.suppressTextObserver -= 1;
  }
}

function restoreUnauthorizedText(entry: CodePageBufferEntry): void {
  if (entry.disposed || entry.restoreUnauthorizedTextScheduled) return;
  entry.restoreUnauthorizedTextScheduled = true;
  queueMicrotask(() => {
    entry.restoreUnauthorizedTextScheduled = false;
    if (entry.disposed) return;
    // Use the latest accepted snapshot: a peer update may have arrived after
    // the rejected local transaction but before this non-nested restoration.
    setTextWithoutSave(entry, entry.acceptedText);
  });
}

function queueLatestDraft(entry: CodePageBufferEntry): void {
  const owner = ownerOf(entry);
  if (
    !entryMayWrite(entry) ||
    !entry.dirty ||
    !owner?.controller ||
    !entry.latestDraft ||
    entry.lastQueuedRevision >= entry.revision
  ) {
    return;
  }
  entry.lastQueuedRevision = entry.revision;
  owner.controller.queueSave(entry.latestDraft);
}

function markDirty(entry: CodePageBufferEntry): void {
  entry.dirty = true;
  entry.saveState = {
    ...entry.saveState,
    saveStatus: entry.saveState.hasConflict ? "conflict" : "unsaved",
    isDirty: true,
  };
}

/**
 * Deliver text this browser did NOT type (a peer's edit replayed by the room)
 * to every view, without claiming it as an unsaved local draft.
 *
 * A clean observer does not become a save writer. If this browser already
 * owes a local save, rebase that draft onto the merge so its older snapshot
 * cannot overwrite a peer's edit. queueSave also retains the merge behind an
 * in-flight PATCH and in the existing recoverable draft.
 */
/** Whether two code page bodies differ at most in their text. */
function sameDraftMetadata(
  a: PageContentDocument,
  b: PageContentDocument,
): boolean {
  const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
  keys.delete("content");
  for (const key of keys) {
    if (!jsonEquals(
      (a as Record<string, JsonValue>)[key] ?? null,
      (b as Record<string, JsonValue>)[key] ?? null,
    )) return false;
  }
  return true;
}

function absorbRemoteText(
  entry: CodePageBufferEntry,
  content: string,
  peerEdit = false,
): void {
  if (content.length > MAX_CODE_SOURCE_CHARS) return;
  const source = ownerOf(entry) ?? entry.holders.values().next().value ?? null;
  entry.revision += 1;
  if (source) entry.latestDraft = source.buildDraft(content);
  for (const holder of entry.holders.values()) {
    if (peerEdit && holder.applyPeerText) holder.applyPeerText(content);
    else holder.applyText(content);
  }
  if (entry.dirty) {
    // Keep the revision pending when authority is suspended; the existing
    // grant/reconnect path will queue it only after writes are allowed again.
    queueLatestDraft(entry);
  } else {
    entry.lastQueuedRevision = entry.revision;
  }
  emit(entry);
}

function publishText(
  entry: CodePageBufferEntry,
  sourceToken: string | null,
  content: string,
): void {
  if (content.length > MAX_CODE_SOURCE_CHARS) return;
  const source =
    (sourceToken ? entry.holders.get(sourceToken) : null) ??
    ownerOf(entry) ??
    entry.holders.values().next().value ??
    null;
  if (!source) return;

  entry.revision += 1;
  entry.latestDraft = source.buildDraft(content);
  markDirty(entry);
  for (const holder of entry.holders.values()) {
    // Include the source holder. react-codemirror batches its public onChange,
    // but metadata actions (language/EOL/debug) may run before that callback.
    // Updating the host ref from the canonical Y.Text now prevents such an
    // action from serializing the pre-keystroke body back over the model.
    holder.applyText(content);
  }
  queueLatestDraft(entry);
  emit(entry);
}

function publishDraft(
  entry: CodePageBufferEntry,
  sourceToken: string,
  draft: PageDocumentDraft,
  replacePending: boolean,
  positions?: CodeTextChangesAgainst,
): boolean {
  if (!entry.holders.has(sourceToken)) return false;
  const content = draft.content_json?.content;
  if (typeof content !== "string") return false;
  if (content.length > MAX_CODE_SOURCE_CHARS) return false;
  const draftLanguage = draft.content_json?.language;
  if (codeLanguageId(draftLanguage) === null) return false;
  const current = entry.text.toString();
  if (current !== content) {
    const changes = authoredTextChanges(entry, current, content, positions);
    if (!changes) {
      refuseAuthoredText(entry);
      return false;
    }
    writeAuthoredText(entry, content, changes);
  } else {
    entry.acceptedText = content;
  }
  entry.revision += 1;
  entry.latestDraft = draft;
  markDirty(entry);
  for (const holder of entry.holders.values()) {
    if (holder.token !== sourceToken) holder.applyDraft(draft);
  }

  const owner = ownerOf(entry);
  if (replacePending) {
    entry.lastQueuedRevision = entry.revision;
    owner?.controller?.replacePendingDraft(draft);
  } else {
    queueLatestDraft(entry);
  }
  emit(entry);
  return true;
}

function acceptHydration(
  entry: CodePageBufferEntry,
  sourceToken: string,
  draft: PageDocumentDraft,
  contentVersion: number,
): PageDocumentDraft {
  if (entry.causal?.active()) {
    // The adapter installs verified originals. REST JSON must never rebuild
    // that same history, including during an own-save acknowledgement.
    return entry.latestDraft ?? draft;
  }
  // The clean ACL observer is the room's document, never a destination for the
  // durable REST projection. If it already had state when this entry was
  // acquired, preserve that state and only borrow the durable draft's metadata.
  // A settled networked room is the document the same way: its Y.Text holds
  // every member's live typing, a REST body is one member's past save. Writing
  // that body into the Y.Text (delete-all + insert-all) broadcast it over the
  // others' newer text, and with three people saving, three different
  // snapshots were installed at once and the file was mangled on every screen.
  const roomOwnsText =
    entry.observerAdopted ||
    (Boolean(entry.session.network) && !entry.roomPending);
  const canonicalDraft = roomOwnsText
    ? draftWithContent(draft, entry.text.toString())
    : draft;
  entry.revision += 1;
  entry.initialized = true;
  entry.acceptNextHydration = false;
  entry.serverVersion = Math.max(entry.serverVersion, contentVersion);
  entry.latestDraft = canonicalDraft;
  entry.dirty = false;
  entry.lastQueuedRevision = entry.revision;
  if (!entry.roomPending && !entry.observerAdopted) {
    if (!roomOwnsText) setTextWithoutSave(entry, draftContent(canonicalDraft));
    // An elected empty room still needs its shareable first struct.
    materializeEmptyFirstSeed(entry);
  }
  // While a room election is open the Y.Text is left alone: it may be about to
  // receive peers' unsaved work, and this REST-loaded body would either
  // destroy that or be destroyed by it. The holders still get the draft below,
  // so the views render immediately.
  for (const holder of entry.holders.values()) {
    if (holder.token !== sourceToken) holder.applyDraft(canonicalDraft);
  }
  emit(entry);
  return canonicalDraft;
}













function settleRoom(entry: CodePageBufferEntry, adopt: boolean): void {
  if (entry.disposed || entry.causal?.active() || !entry.roomPending) return;
  entry.roomPending = false;
  if (!adopt && entry.text.length === 0) {
    entry.awaitingSeed = true;
    materializeEmptyFirstSeed(entry);
    return;
  }
  const canonical = entry.text.toString();
  const known = entry.latestDraft ? draftContent(entry.latestDraft) : "";
  if (canonical !== known) absorbRemoteText(entry, canonical);
}

function detachTextObserver(entry: CodePageBufferEntry): void {
  if (!entry.textObserver) return;
  entry.text.unobserve(entry.textObserver);
  entry.textObserver = null;
}

function attachTextObserver(entry: CodePageBufferEntry): void {
  if (entry.disposed) return;
  const observedText = entry.text;
  const observer = (_event: Y.YTextEvent, transaction: Y.Transaction) => {
    if (
      entry.disposed ||
      entry.text !== observedText ||
      entry.causal?.applying() ||
      entry.suppressTextObserver > 0
    ) {
      return;
    }
    if (!transaction.local) {
      // Applied by Y.applyUpdate — a peer's edit arriving over the room, not
      // an edit made in this browser.
      entry.acceptedText = observedText.toString();
      absorbRemoteText(entry, observedText.toString(), true);
      return;
    }
    if (!entryMayWrite(entry)) {
      // `EditorState.readOnly` is an affordance, not a transaction firewall:
      // a retained host command or extension can still dispatch a programmatic
      // change. Leaving that change in the participant Y.Doc lets it hitchhike
      // when checkout is granted again. Restore the last authorized snapshot
      // on the next microtask (avoiding a nested Yjs transaction), without
      // publishing or saving either transaction.
      restoreUnauthorizedText(entry);
      return;
    }
    entry.acceptedText = observedText.toString();
    if (entry.awaitingSeed) {
      entry.awaitingSeed = false;
      const seeded = observedText.toString();
      const loaded = entry.latestDraft ? draftContent(entry.latestDraft) : "";
      if (seeded === loaded) {
        // Pure seed — the loaded body, unchanged. Share it with every view,
        // but never claim the page became unsaved just by being opened.
        absorbRemoteText(entry, seeded);
        return;
      }
      // The editor seeded from its LIVE view, which had already been typed in
      // while the election was open. That IS an unsaved change and falls
      // through to the normal publish path so autosave picks it up.
    }
    const origin =
      transaction.origin &&
      (typeof transaction.origin === "object" ||
        typeof transaction.origin === "function")
        ? (transaction.origin as object)
        : null;
    const sourceToken = origin
      ? (entry.originTokens.get(origin) ?? null)
      : null;
    publishText(entry, sourceToken, observedText.toString());
  };
  entry.textObserver = observer;
  observedText.observe(observer);
}

function adoptObserver(entry: CodePageBufferEntry): void {
  const observer = entry.observerSession;
  if (
    entry.disposed ||
    entry.causal?.active() ||
    !observer ||
    entry.observerAdopted ||
    !observer.network?.hasState ||
    !observer.ydoc.share.has(CODE_COLLAB_TEXT_KEY)
  ) {
    return;
  }
  detachTextObserver(entry);
  entry.session = observer;
  entry.text = getCodeYText(observer.ydoc);
  entry.acceptedText = entry.text.toString();
  entry.undoManager = new Y.UndoManager(entry.text);
  entry.undoManagers.add(entry.undoManager);
  entry.observerAdopted = true;
  entry.roomPending = false;
  entry.awaitingSeed = false;
  entry.bindingGeneration += 1;
  attachTextObserver(entry);
  // The update that made hasState true happened before the new Y.Text observer
  // was attached, so project its already-current value explicitly.
  absorbRemoteText(entry, entry.text.toString());
}

function watchTransportWritable(entry: CodePageBufferEntry): void {
  const provider = entry.session.provider;
  const subscribe = provider?.subscribeWritable;
  const unsubscribeWritable = subscribe?.((writable) => {
    if (entry.disposed) return;
    if (entry.transportWritable === writable) return;
    entry.transportWritable = writable;
    if (writable) queueLatestDraft(entry);
    emit(entry);
  });
  const unsubscribeConnection = entry.session.network?.subscribeConnection(
    (connected) => {
      if (!connected) return;
      // A freshness refresh resolves write authority at hello, before replay
      // completes. Retry the deferred empty seed only after this connection
      // is usable; a claimed-empty non-seeder never reaches connected=true.
      // Let the provider finish installing its initial-snapshot hold first.
      queueMicrotask(() => materializeEmptyFirstSeed(entry));
    },
  );
  entry.transportUnsubscribe = () => {
    unsubscribeWritable?.();
    unsubscribeConnection?.();
  };
}

function watchConnectionFailure(
  entry: CodePageBufferEntry,
  session: CollabSession,
): void {
  const subscribe =
    session.network?.subscribePageYjsConnectionFailureDiagnostic;
  if (!subscribe) return;
  const unsubscribe = subscribe((diagnostic) => {
    if (entry.disposed || entry.connectionFailureLogged) return;
    entry.connectionFailureLogged = true;
    const reached = [
      ...new Set(
        diagnostic.reached.filter((stage) =>
          CODE_PAGE_CONNECTION_STAGES.has(stage),
        ),
      ),
    ];
    try {
      console.warn(
        `${CODE_PAGE_CONNECTION_FAILURE_LOG} ${JSON.stringify({
          kind: "failure",
          reached,
        })}`,
      );
    } catch {
      // A diagnostic sink can never affect the shared editor buffer.
    }
  });
  entry.connectionDiagnosticUnsubscribes.add(unsubscribe);
}

function createEntry(
  pageId: string,
  binding: CodePageBufferBinding,
  mode: CodePageBufferAuthorityMode,
  generation: number,
  authority: CodePageBufferAuthority,
  causalEpoch: number,
): CodePageBufferEntry {
  // A page whose originals moved to a later epoch is always rebound by the
  // causal adapter to a fresh document bound to that epoch, so a room joined
  // here could only ask for the old realm: its ticket was refused (409
  // page_causal_epoch_stale) on every open of such a page.
  const projectionSession = createCollabSession({
    docId: pageId,
    pageJsonCausalModel: true,
    ...(mode === "participant" && causalEpoch <= 1 ? {} : { roomName: null }),
  });
  const text = getCodeYText(projectionSession.ydoc);
  const undoManager = new Y.UndoManager(text);
  const registryKey = `${pageId}\u0000${generation}`;
  const entry: CodePageBufferEntry = {
    pageId,
    disposed: false,
    registryKey,
    authorityGeneration: generation,
    authorityBinding: binding,
    authorityMode: mode,
    authorityWritable: authorityMayWrite(authority),
    authorityDesiredWritable: authorityMayWrite(authority),
    authorizationIntentGeneration: 0,
    authorizationRefresh: null,
    transportWritable: projectionSession.network?.writable ?? true,
    bindingGeneration: 0,
    session: projectionSession,
    text,
    acceptedText: text.toString(),
    undoManager,
    sessions: new Set([projectionSession]),
    undoManagers: new Set([undoManager]),
    observerSession: null,
    observerAdopted: false,
    observerUpdateListener: null,
    textObserver: null,
    transportUnsubscribe: null,
    connectionDiagnosticUnsubscribes: new Set(),
    connectionFailureLogged: false,
    holders: new Map(),
    ownerToken: null,
    originTokens: new WeakMap(),
    listeners: new Set(),
    initialized: false,
    acceptNextHydration: false,
    serverVersion: 0,
    revision: 0,
    lastQueuedRevision: 0,
    dirty: false,
    latestDraft: null,
    saveState: DEFAULT_SAVE_STATE,
    suppressTextObserver: 0,
    restoreUnauthorizedTextScheduled: false,
    roomPending: Boolean(projectionSession.network),
    awaitingSeed: false,
  };

  entry.causal = createCodePageCausalAdapter({
    pageId,
    session: () => entry.session,
    canWrite: () => entry.authorityWritable &&
      (!entry.session.network || !entry.session.network.connected ||
        entry.transportWritable && entry.session.network.writable),
    hasUnboundDraft: () => entry.dirty,
    // Durable originals go only into a document that holds nothing else. A
    // REST body hydrated before the adapter bound was written under a fresh
    // client, so merging the originals beside it showed the text twice until
    // the room session replaced this one. Nothing unsaved is dropped: a dirty
    // buffer fails preparation closed before this is asked.
    requiresNativeBinding: () =>
      entry.observerSession !== null || entry.session.ydoc.store.clients.size > 0,
    rebind: (epoch, joinLegacyRoom = false) => {
      detachTextObserver(entry);
      entry.transportUnsubscribe?.(); entry.transportUnsubscribe = null;
      if (entry.observerSession && entry.observerUpdateListener) entry.observerSession.ydoc.off("update", entry.observerUpdateListener);
      entry.observerUpdateListener = null; entry.observerSession = null;
      for (const session of entry.sessions) session.provider?.destroy();
      const session = createCollabSession({ docId: pageId, pageJsonCausalModel: true,
        ...(epoch ? { pageJsonCausalEpoch: epoch } : {}),
        ...((!epoch && !joinLegacyRoom) || entry.authorityMode === "local" ? { roomName: null } : {}) });
      entry.session = session; entry.sessions.add(session);
      entry.text = getCodeYText(session.ydoc); entry.acceptedText = entry.text.toString();
      entry.undoManager = new Y.UndoManager(entry.text); entry.undoManagers.add(entry.undoManager);
      entry.observerAdopted = false; entry.roomPending = false; entry.awaitingSeed = false;
      entry.transportWritable = session.network?.writable ?? true;
      entry.bindingGeneration++; attachTextObserver(entry); watchTransportWritable(entry); watchConnectionFailure(entry, session);
      emit(entry);
    },
    deliver: content => {
      const text = String(content.content ?? "");
      // A peer's keystroke changes the text and nothing else. Re-applying the
      // whole draft rebuilt every window's run, debug and test state from
      // fresh JSON for each character; the text alone goes the peer path.
      const previous = entry.initialized ? entry.latestDraft?.content_json : undefined;
      const textOnly = Boolean(previous) && typeof content.content === "string" &&
        sameDraftMetadata(previous as PageContentDocument, content);
      entry.acceptedText = text;
      entry.initialized = true; entry.revision++;
      entry.latestDraft = { ...entry.latestDraft, content_json: content };
      entry.lastQueuedRevision = entry.revision;
      for (const holder of entry.holders.values()) {
        if (textOnly && holder.applyPeerText) holder.applyPeerText(text);
        else holder.applyDraft(entry.latestDraft);
      }
    },
    changed: () => emit(entry),
  });

  attachTextObserver(entry);
  watchConnectionFailure(entry, projectionSession);

  const network = projectionSession.network;
  if (network) {
    watchTransportWritable(entry);
    network.whenReady
      // Adopt only when the room actually delivered state: a claimed-but-empty
      // room (its seeder died before contributing) must not blank the buffer.
      .then(({ first }) => {
        if (entry.disposed || entry.session !== projectionSession) return;
        entry.transportWritable = network.writable;
        settleRoom(entry, !first && network.hasState);
        if (entry.transportWritable) queueLatestDraft(entry);
        emit(entry);
      })
      .catch(() => settleRoom(entry, false)); // room dead → behave like local
  }

  if (mode === "acl-viewer") {
    const observerSession = createCollabSession({ docId: pageId });
    entry.observerSession = observerSession;
    entry.sessions.add(observerSession);
    watchConnectionFailure(entry, observerSession);
    const adoptIfRoomHasState = () => adoptObserver(entry);
    entry.observerUpdateListener = adoptIfRoomHasState;
    observerSession.ydoc.on("update", adoptIfRoomHasState);
    adoptIfRoomHasState();
    void observerSession.network?.whenReady
      .then(adoptIfRoomHasState)
      .catch(() => undefined);
  }

  return entry;
}

function destroyEntry(entry: CodePageBufferEntry): void {
  if (entry.disposed) return;
  entry.disposed = true;
  entry.causal?.dispose();
  entry.transportUnsubscribe?.();
  entry.transportUnsubscribe = null;
  for (const unsubscribe of entry.connectionDiagnosticUnsubscribes) {
    try {
      unsubscribe();
    } catch {
      // Diagnostic cleanup is best-effort and cannot block session teardown.
    }
  }
  entry.connectionDiagnosticUnsubscribes.clear();
  if (entry.observerSession && entry.observerUpdateListener) {
    entry.observerSession.ydoc.off("update", entry.observerUpdateListener);
  }
  entry.observerUpdateListener = null;
  detachTextObserver(entry);
  entry.listeners.clear();
  for (const undoManager of entry.undoManagers) undoManager.destroy();
  entry.undoManagers.clear();
  for (const session of entry.sessions) session.destroy();
  entry.sessions.clear();
}

function asyncNull(): Promise<PageItem | null> {
  return Promise.resolve(null);
}

function transitionEntryAuthority(
  entry: CodePageBufferEntry,
  nextAuthority: CodePageBufferAuthority,
): void {
  if (
    entry.disposed ||
    authorityBinding(nextAuthority) !== entry.authorityBinding
  ) {
    return;
  }
  const nextWritable = authorityMayWrite(nextAuthority);
  if (!nextWritable) {
    if (
      entry.authorityDesiredWritable ||
      entry.authorizationIntentGeneration === 0
    ) {
      entry.session.provider?.suspendLocalDocumentUpdates?.();
      entry.authorityDesiredWritable = false;
      entry.authorizationIntentGeneration += 1;
    }
    if (entry.authorityWritable) {
      entry.authorityWritable = false;
      emit(entry);
    }
    return;
  }
  if (entry.authorityDesiredWritable) return;
  entry.authorityDesiredWritable = true;
  const authorizationGeneration = ++entry.authorizationIntentGeneration;
  if (entry.authorityWritable) return;
  const provider = entry.session.provider;
  const refresh = provider?.refreshAuthorization;
  if (!refresh) {
    entry.authorityWritable = true;
    provider?.resumeLocalDocumentUpdates?.();
    entry.causal?.retryPreparation();
    materializeEmptyFirstSeed(entry);
    queueLatestDraft(entry);
    emit(entry);
    return;
  }
  const request = refresh()
    .then((transportWritable) => {
      if (
        entry.disposed ||
        entry.authorizationIntentGeneration !== authorizationGeneration ||
        !entry.authorityDesiredWritable
      ) {
        return;
      }
      entry.transportWritable = transportWritable;
      entry.authorityWritable = true;
      provider?.resumeLocalDocumentUpdates?.();
      entry.causal?.retryPreparation();
      materializeEmptyFirstSeed(entry);
      if (transportWritable) queueLatestDraft(entry);
      emit(entry);
    })
    .catch(() => undefined)
    .finally(() => {
      if (entry.authorizationRefresh === request) {
        entry.authorizationRefresh = null;
      }
    });
  entry.authorizationRefresh = request;
}

export interface CodePageCommandHistory {
  /** A failed apply must not leave its rejected edit in Redo. */
  rollback(): boolean;
  /** Explicit receipt Undo/failed-Undo recovery use the native history pair. */
  undo(): boolean;
  redo(): boolean;
}

export interface CodePageBufferLease {
  readonly collaborationAdapter: PageDocumentCollaborationAdapter;
  readonly pageId: string;
  readonly token: string;
  /** Stable for this grant/revoke generation; shared by duplicate windows. */
  readonly authorityIdentity: string;
  readonly session: CollabSession;
  readonly textKey: string;
  readonly undoManager: Y.UndoManager;
  matchesAuthority(authority: CodePageBufferAuthority): boolean;
  updateAuthority(authority: CodePageBufferAuthority): void;
  isWritable(): boolean;
  /** Changes whenever the concrete Y.Doc/editor binding must be rebuilt. */
  getBindingGeneration(): string;
  hydrate(
    draft: PageDocumentDraft,
    contentVersion: number,
  ): PageDocumentDraft;
  /**
   * This view's draft, as one local edit. `positions`, when the edit knows
   * where it was made (a Find-in-Files hit, a refactor's text changes), write
   * it at those places. False: not written (no authority, invalid draft, or a
   * text change whose place would be a guess; the views then have the room's
   * text back).
   */
  publishDraft(
    draft: PageDocumentDraft,
    positions?: CodeTextChangesAgainst,
  ): boolean;
  /** One explicit local text gesture; null means no edit was admitted. The
   * rollback only undoes this still-current gesture, never later edits. */
  applyCommandDraft(
    draft: PageDocumentDraft,
    expectedContent: string,
  ): CodePageCommandHistory | null;
  replaceDraft(draft: PageDocumentDraft): boolean;
  /** The page was renamed outside the editor (a collaborator, the titlebar,
   * a refused name handed back): the draft the buffer holds takes the name,
   * so a hydration or a room delivery never hands a window the old one. No
   * save: the name is already the server's. */
  followTitle(title: string): void;
  publishTextareaText(content: string): void;
  registerEditorOrigin(origin: unknown): void;
  updateHolder(options: {
    buildDraft(content: string): PageDocumentDraft;
    applyText(content: string): void;
    applyPeerText?(content: string): void;
    applyDraft(draft: PageDocumentDraft): void;
    controller: CodePageBufferSaveController;
    saveState: CodePageBufferSaveState;
  }): void;
  getSaveState(): CodePageBufferSaveState;
  /** Monotonic canonical document generation, including hydrations. */
  getDocumentRevision(): number | null;
  isCurrentText(content: string): boolean;
  isOwner(): boolean;
  subscribe(listener: () => void): () => void;
  saveNow(draft?: PageDocumentDraft): Promise<PageItem | null>;
  retrySave(): Promise<PageItem | null>;
  overwriteRemote(): Promise<PageItem | null>;
  saveConflictAsCopy(): Promise<PageItem | null>;
  reload(): Promise<PageItem | null>;
  discardLocalAndReload(options?: { keepAside?: DiscardedChangesKeep }): Promise<PageItem | null>;
  release(): void;
}

/**
 * Acquire one view of `pageId` under the current authority generation. The
 * first live view is the autosave owner; duplicate views with the same
 * authority share the entry.
 */
export function acquireCodePageBuffer(
  pageId: string,
  authority: CodePageBufferAuthority,
  /** The causal epoch of the page read that opens the buffer, when known. */
  causalEpoch = 0,
): CodePageBufferLease {
  const binding = authorityBinding(authority);
  const mode = authorityMode(authority);
  const previousAuthority = _pageAuthorities.get(pageId);
  const authorityState =
    previousAuthority?.binding === binding
      ? previousAuthority
      : { binding, generation: ++_authoritySequence };
  _pageAuthorities.set(pageId, authorityState);
  const registryKey = `${pageId}\u0000${authorityState.generation}`;
  const entry =
    _entries.get(registryKey) ??
    createEntry(pageId, binding, mode, authorityState.generation, authority, causalEpoch);
  if (!_entries.has(registryKey)) _entries.set(registryKey, entry);
  transitionEntryAuthority(entry, authority);

  _holderSequence += 1;
  const token = `code-buffer-${_holderSequence}`;
  const holder: CodePageBufferHolder = {
    token,
    buildDraft: (content) =>
      entry.latestDraft
        ? draftWithContent(entry.latestDraft, content)
        : { content_json: { content } },
    applyText: () => {},
    applyDraft: () => {},
    controller: null,
    saveState: DEFAULT_SAVE_STATE,
    editorOrigins: new Set(),
    listeners: new Set(),
  };
  entry.holders.set(token, holder);
  entry.ownerToken ??= token;

  let released = false;
  const ownerController = (): CodePageBufferSaveController | null =>
    ownerOf(entry)?.controller ?? null;
  const requireLive = (): boolean => !released && entry.holders.has(token);

  return {
    collaborationAdapter: entry.causal!.forLease(token),
    pageId,
    token,
    authorityIdentity: registryKey,
    get session() {
      return entry.session;
    },
    textKey: CODE_COLLAB_TEXT_KEY,
    get undoManager() {
      return entry.undoManager;
    },
    matchesAuthority(nextAuthority) {
      return authorityBinding(nextAuthority) === entry.authorityBinding;
    },
    updateAuthority(nextAuthority) {
      if (!requireLive()) return;
      transitionEntryAuthority(entry, nextAuthority);
    },
    isWritable() {
      return requireLive() && entryMayWrite(entry);
    },
    getBindingGeneration() {
      return `${entry.authorityGeneration}:${entry.bindingGeneration}${entry.causal?.active() ? entry.causal.ready() ? ":native-ready" : ":native-held" : ""}`;
    },
    hydrate(draft, contentVersion) {
      if (!requireLive()) return draft;
      const hydrationContent = draft.content_json?.content;
      const safeDraft =
        typeof hydrationContent === "string"
          ? draft
          : draftWithContent(draft, "");
      if (
        typeof hydrationContent === "string" &&
        hydrationContent.length > MAX_CODE_SOURCE_CHARS
      ) {
        // Never truncate canonical source. Keep the last accepted draft (or a
        // safe empty projection on first load) and leave the oversized server
        // value unaccepted so no later save can claim it was hydrated.
        return entry.latestDraft ?? draftWithContent(draft, "");
      }
      const validVersion =
        Number.isSafeInteger(contentVersion) && contentVersion > 0
          ? contentVersion
          : 0;
      if (
        !entry.initialized ||
        entry.acceptNextHydration ||
        (!entry.dirty && validVersion > entry.serverVersion)
      ) {
        return acceptHydration(entry, token, safeDraft, validVersion);
      }
      return entry.latestDraft ?? safeDraft;
    },
    publishDraft(draft, positions) {
      return (
        requireLive() &&
        entryMayWrite(entry) &&
        publishDraft(entry, token, draft, false, positions)
      );
    },
    applyCommandDraft(draft, expectedContent) {
      const content = draft.content_json?.content;
      if (
        !requireLive() ||
        !entryMayWrite(entry) ||
        typeof content !== "string" ||
        content.length > MAX_CODE_SOURCE_CHARS ||
        codeLanguageId(draft.content_json?.language) === null ||
        entry.text.toString() !== expectedContent
      )
        return null;
      const manager = entry.undoManager;
      const text = entry.text;
      const origin = {};
      // An AI command replaces the whole text (set_code_content): it has no
      // places of its own and no typed run to keep whole, so its hunks are
      // read off the two texts. Every edit that does know its places goes
      // through publishDraft's positions (code-text-changes, TRIAGE 47).
      const changes = diff(expectedContent, content, {
        scanLimit: 500,
        timeout: 20,
      });
      // Hydration/REST/rollback still use the untracked entry origin. Only this
      // admitted local gesture joins the editor's existing personal history.
      manager.stopCapturing();
      manager.addTrackedOrigin(origin);
      entry.suppressTextObserver++;
      try {
        entry.session.ydoc.transact(() => {
          for (let i = changes.length - 1; i >= 0; i--) {
            const change = changes[i];
            text.delete(change.fromA, change.toA - change.fromA);
            const insert = content.slice(change.fromB, change.toB);
            if (insert) text.insert(change.fromA, insert);
          }
        }, origin);
      } finally {
        entry.suppressTextObserver--;
        manager.removeTrackedOrigin(origin);
        manager.stopCapturing();
      }
      let undoItem = changes.length ? manager.undoStack.at(-1) : null;
      let redoItem: typeof undoItem = null;
      publishDraft(entry, token, draft, false);
      const moveHistory = (direction: "undo" | "redo", discardRedo = false) => {
        const item = direction === "undo" ? undoItem : redoItem;
        const stack =
          direction === "undo" ? manager.undoStack : manager.redoStack;
        const before = direction === "undo" ? content : expectedContent;
        const after = direction === "undo" ? expectedContent : content;
        if (
          !requireLive() ||
          !entryMayWrite(entry) ||
          entry.text !== text ||
          entry.undoManager !== manager ||
          text.toString() !== before ||
          (changes.length && (!item || stack.at(-1) !== item))
        )
          return false;
        if (!item) return true;
        // Both the ordinary editor buttons and receipt actions move the same
        // Yjs structs. A string restoration here would invalidate their history.
        entry.suppressTextObserver++;
        try {
          const moved = direction === "undo" ? manager.undo() : manager.redo();
          if (moved !== item) return false;
          undoItem = direction === "redo" ? manager.undoStack.at(-1) : null;
          redoItem = direction === "undo" ? manager.redoStack.at(-1) : null;
          if (discardRedo) {
            manager.clear(false, true);
            redoItem = null;
          }
          manager.stopCapturing();
          entry.acceptedText = text.toString();
          return text.toString() === after;
        } finally {
          entry.suppressTextObserver--;
        }
      };
      return {
        rollback: () => moveHistory("undo", true),
        undo: () => moveHistory("undo"),
        redo: () => moveHistory("redo"),
      };
    },
    followTitle(title) {
      if (!requireLive() || !entry.latestDraft || entry.latestDraft.title === title) return;
      entry.latestDraft = { ...entry.latestDraft, title };
    },
    replaceDraft(draft) {
      return (
        requireLive() &&
        entryMayWrite(entry) &&
        publishDraft(entry, token, draft, true)
      );
    },
    publishTextareaText(content) {
      if (
        !requireLive() ||
        !entryMayWrite(entry) ||
        content.length > MAX_CODE_SOURCE_CHARS ||
        entry.text.toString() === content
      ) {
        return;
      }
      const current = entry.text.toString();
      const changes = authoredTextChanges(entry, current, content);
      if (!changes) {
        refuseAuthoredText(entry);
        return;
      }
      writeAuthoredText(entry, content, changes);
      publishText(entry, token, content);
    },
    registerEditorOrigin(origin) {
      if (
        !requireLive() ||
        !origin ||
        (typeof origin !== "object" && typeof origin !== "function")
      ) {
        return;
      }
      const objectOrigin = origin as object;
      holder.editorOrigins.add(objectOrigin);
      entry.originTokens.set(objectOrigin, token);
    },
    updateHolder(options) {
      if (!requireLive()) return;
      holder.buildDraft = options.buildDraft;
      holder.applyText = options.applyText;
      holder.applyPeerText = options.applyPeerText;
      holder.applyDraft = options.applyDraft;
      holder.controller = options.controller;
      holder.saveState = options.saveState;
      if (entry.ownerToken === token) {
        const version = options.saveState.contentVersion;
        const savedVersionIsCurrent =
          (!entry.causal?.active() || entry.causal.saved()) &&
          options.saveState.saveStatus === "saved" &&
          !options.saveState.isDirty &&
          typeof version === "number" &&
          Number.isSafeInteger(version) &&
          version >= entry.serverVersion &&
          (!entry.dirty || version > entry.serverVersion);
        if (savedVersionIsCurrent) {
          entry.serverVersion = version;
          entry.dirty = false;
        }
        // queueSave schedules the hook's React update, but the registry emits
        // its dirty snapshot synchronously. If that render arrives first, the
        // holder may briefly report the pre-edit "saved" state at the same
        // version. Never let that stale state clear the canonical dirty bit.
        const nextSaveState =
          entry.dirty && options.saveState.saveStatus === "saved"
            ? {
                ...options.saveState,
                saveStatus: "unsaved" as const,
                isDirty: true,
              }
            : options.saveState;
        const saveStateChanged = !saveStateEquals(
          entry.saveState,
          nextSaveState,
        );
        entry.saveState = nextSaveState;
        queueLatestDraft(entry);
        if (saveStateChanged) emit(entry);
      }
    },
    getSaveState() {
      return entry.saveState;
    },
    getDocumentRevision() {
      return requireLive() ? entry.revision : null;
    },
    isCurrentText(content) {
      return entry.text.toString() === content;
    },
    isOwner() {
      return !released && entry.ownerToken === token;
    },
    subscribe(listener) {
      if (!requireLive()) return () => {};
      entry.listeners.add(listener);
      holder.listeners.add(listener);
      // Replay the current binding snapshot. Observer adoption can happen in a
      // microtask between acquire() and the parent's subscription effect; an
      // edge-only subscription would miss that generation and leave CodeMirror
      // bound to the durable projection.
      listener();
      return () => {
        holder.listeners.delete(listener);
        entry.listeners.delete(listener);
      };
    },
    saveNow(draft) {
      if (!requireLive() || !entryMayWrite(entry)) return asyncNull();
      if (draft) publishDraft(entry, token, draft, false);
      const controller = ownerController();
      if (!controller) return asyncNull();
      return draft || entry.dirty
        ? controller.saveNow(entry.latestDraft ?? draft)
        : controller.saveNow();
    },
    retrySave() {
      return requireLive() && entryMayWrite(entry)
        ? (ownerController()?.retrySave() ?? asyncNull())
        : asyncNull();
    },
    overwriteRemote() {
      return requireLive() && entryMayWrite(entry)
        ? (ownerController()?.overwriteRemote() ?? asyncNull())
        : asyncNull();
    },
    saveConflictAsCopy() {
      return requireLive() && entryMayWrite(entry)
        ? (ownerController()?.saveConflictAsCopy() ?? asyncNull())
        : asyncNull();
    },
    reload() {
      if (!requireLive()) return asyncNull();
      entry.acceptNextHydration = true;
      return (ownerController()?.reload() ?? asyncNull()).finally(() => {
        entry.acceptNextHydration = false;
      });
    },
    discardLocalAndReload(options) {
      if (!requireLive()) return asyncNull();
      entry.acceptNextHydration = true;
      return (
        ownerController()?.discardLocalAndReload(options) ?? asyncNull()
      ).finally(() => {
        entry.acceptNextHydration = false;
      });
    },
    release() {
      if (released) return;
      released = true;
      entry.causal?.release(token);
      const wasOwner = entry.ownerToken === token;
      for (const origin of holder.editorOrigins) {
        entry.originTokens.delete(origin);
      }
      for (const listener of holder.listeners) {
        entry.listeners.delete(listener);
      }
      holder.listeners.clear();
      entry.holders.delete(token);
      if (wasOwner) {
        entry.ownerToken = entry.holders.keys().next().value ?? null;
        const nextOwner = ownerOf(entry);
        if (nextOwner) {
          entry.saveState =
            entry.dirty && nextOwner.saveState.saveStatus === "saved"
              ? {
                  ...nextOwner.saveState,
                  saveStatus: "unsaved",
                  isDirty: true,
                }
              : nextOwner.saveState;
          entry.lastQueuedRevision = Math.min(
            entry.lastQueuedRevision,
            Math.max(0, entry.revision - 1),
          );
          queueLatestDraft(entry);
        }
      }
      emit(entry);
      if (entry.holders.size > 0) return;
      _entries.delete(entry.registryKey);
      destroyEntry(entry);
      if (![..._entries.values()].some((live) => live.pageId === pageId)) {
        _pageAuthorities.delete(pageId);
      }
    },
  };
}

/** Number of open CodeWindow holders for diagnostics/tests. */
export function codePageBufferHolderCount(pageId: string): number {
  let holders = 0;
  for (const entry of _entries.values()) {
    if (entry.pageId === pageId) holders += entry.holders.size;
  }
  return holders;
}

/** Test-only cleanup. */
export function _clearCodePageBuffersForTest(): void {
  for (const entry of _entries.values()) {
    destroyEntry(entry);
  }
  _entries.clear();
  _pageAuthorities.clear();
  _authoritySequence = 0;
}
