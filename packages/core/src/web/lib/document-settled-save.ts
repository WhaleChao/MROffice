/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pageDocumentJsonEquals } from "./page-save-acknowledgement";

/**
 * Whether a document page is already saved, so a caller that needs "the
 * server holds what I see, at version N" (the Copilot selection binding, a
 * server export, a version snapshot) can use N without sending the whole
 * document again.
 *
 * Every answer other than a proof is null, and null keeps the caller on its
 * ordinary save. The cheap checks come first (who owns the editor, pending
 * local work, the save generation, the autosave hook's own verdict); the
 * document comparison runs last, because not every document change is an
 * authored edit that marks the page dirty: the first touch of a list or table
 * back-fills block ids on its items through an appended transaction, and that
 * document still has to be saved before the server can bind to it.
 */
export interface SettledDocumentSaveInput {
  pageId: string;
  /** The editor is usable, hydrated and still bound to `pageId`. */
  editorOwnsPage: boolean;
  /** Anything between an editor transaction and the autosave queue: an IME
   *  composition or its deferred commit, a deferred large-document
   *  serialization, an open header/footer draft, an import. */
  localWorkPending: boolean;
  generation: {
    pageId: string;
    dirty: boolean;
    queuedDraft: unknown;
  };
  baseline: {
    pageId: string;
    title?: string;
    content: Record<string, unknown>;
    contentVersion: number;
  } | null;
  /** The title an explicit save would send. */
  title: string;
  /** The autosave hook's verdict: nothing queued, in flight, refused or
   *  unconfirmed by the collaboration model at this version, with the body
   *  the server returned for it (null otherwise). */
  settled: { contentVersion: number; content: Record<string, unknown> | null } | null;
  /** The live document normalized exactly as an explicit save would send it,
   *  or null when it cannot be. Read only after every other check passed. */
  readLiveContent: () => Record<string, unknown> | null;
  /** A page whose saves merge through the document causal model: the
   *  server's stored form of a body, i.e. the model's semantic projection
   *  (null attributes and other schema defaults dropped, text runs joined).
   *  The layout records it keeps beside that projection are compared
   *  separately, so a re-measured pagination still saves. */
  semanticProjection?: (content: Record<string, unknown>) => unknown;
  /** What has to match. "document" (the default): the whole body, root
   *  attributes and pagination records included, for a caller the server
   *  renders or paginates from (a server export, a version snapshot).
   *  "blocks": only the block content, for a caller that binds blocks at a
   *  version and nothing else (the Copilot's source; the server reads the
   *  selected blocks by id and treats the page context as read-only). A root
   *  attribute the user changes is an authored transaction, so it is already
   *  refused above as a dirty generation; what this skips is the root state an
   *  opened editor fills in by itself: its schema defaults (headingNumbering,
   *  listStyles, footnoteMode, ...) and the pagination it measures, which a
   *  page created through the API never holds. */
  compare?: "document" | "blocks";
  /** "blocks" only: the top-level blocks the caller binds by id (the
   *  Copilot's circled blocks). Inside them every block id must be the
   *  server's; elsewhere an id the opened editor back-filled on a nested item
   *  (a list's items, which the Word importer leaves without one) binds
   *  nothing at this version, so it waits for the ordinary autosave instead
   *  of a whole-document save before the question (AI70H H13 C06). Absent,
   *  every block id must match. */
  boundBlockIds?: readonly string[];
}

/** Kept beside the semantic projection by the server, as the editor sent them. */
const PRIVATE_LAYOUT_ATTRS = ["pageBlockSizes", "pageBlockLayoutKey", "pageRegionLayoutReceipt"] as const;

function rootAttrs(content: Record<string, unknown>): Record<string, unknown> {
  const attrs = content.attrs;
  return attrs && typeof attrs === "object" && !Array.isArray(attrs)
    ? attrs as Record<string, unknown>
    : {};
}

function sameStoredDocument(
  live: Record<string, unknown>,
  stored: Record<string, unknown>,
  project: (content: Record<string, unknown>) => unknown,
): boolean {
  const liveAttrs = rootAttrs(live);
  const storedAttrs = rootAttrs(stored);
  return PRIVATE_LAYOUT_ATTRS.every((key) =>
    pageDocumentJsonEquals(liveAttrs[key] ?? null, storedAttrs[key] ?? null),
  ) && pageDocumentJsonEquals(project(live), project(stored));
}

function blocksOf(projected: unknown): unknown {
  return projected && typeof projected === "object" && !Array.isArray(projected)
    ? (projected as Record<string, unknown>).content
    : undefined;
}

function plainRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? value as Record<string, unknown>
    : null;
}

/** Whether `live` holds exactly the server's blocks: the same nodes, text,
 *  marks and block ids, and every attribute the server stores. An attribute
 *  only the live side carries is what an opened editor fills in by itself (a
 *  schema default, a derived heading anchor); an authored one marks the
 *  generation dirty before this is asked. A block id is never such a fill-in
 *  where the server binds by it: in a bound block (or, with no `bound` set,
 *  anywhere), one the server lacks still needs a save. */
function sameServerBlocks(
  live: unknown,
  stored: unknown,
  bound: ReadonlySet<string> | null = null,
  nested = false,
  inBound = false,
): boolean {
  if (Array.isArray(live) || Array.isArray(stored)) {
    return Array.isArray(live) && Array.isArray(stored) && live.length === stored.length &&
      live.every((item, index) => sameServerBlocks(item, stored[index], bound, nested, inBound));
  }
  const a = plainRecord(live);
  const b = plainRecord(stored);
  if (!a || !b) return pageDocumentJsonEquals(live ?? null, stored ?? null);
  if (a.type !== b.type || (a.text ?? null) !== (b.text ?? null)) return false;
  const liveAttrs = plainRecord(a.attrs) ?? {};
  const storedAttrs = plainRecord(b.attrs) ?? {};
  const liveId = liveAttrs.block_id ?? null;
  const storedId = storedAttrs.block_id ?? null;
  const strict = !bound || inBound || (typeof liveId === "string" && bound.has(liveId));
  if (liveId !== storedId && (strict || !nested || storedId !== null)) return false;
  if (!Object.keys(storedAttrs).every((key) =>
    pageDocumentJsonEquals(liveAttrs[key] ?? null, storedAttrs[key] ?? null))) return false;
  return sameMarkSet(a.marks ?? [], b.marks ?? []) &&
    sameServerBlocks(a.content ?? [], b.content ?? [], bound, true, strict && bound !== null);
}

/** A run's marks are a set: the editor holds them in its schema's order, the
 *  server's importer in the order it read them from the file. Every imported
 *  Word document differed in that order alone, so the first Copilot press on
 *  it saved the whole book before asking (AI70H H13 C06: 45 s on a loaded
 *  machine, the 45 s source-save deadline then failed the press). */
function sameMarkSet(live: unknown, stored: unknown): boolean {
  if (!Array.isArray(live) || !Array.isArray(stored)) return sameServerBlocks(live, stored);
  if (live.length !== stored.length) return false;
  const unmatched = [...live];
  return stored.every((mark) => {
    const index = unmatched.findIndex((candidate) => sameServerBlocks(candidate, mark));
    if (index < 0) return false;
    unmatched.splice(index, 1);
    return true;
  });
}

export function settledDocumentContentVersion(
  input: SettledDocumentSaveInput,
): number | null {
  const { baseline, generation, settled } = input;
  if (
    !input.editorOwnsPage ||
    input.localWorkPending ||
    generation.pageId !== input.pageId ||
    generation.dirty ||
    generation.queuedDraft !== null ||
    !baseline ||
    baseline.pageId !== input.pageId ||
    !Number.isSafeInteger(baseline.contentVersion) ||
    baseline.contentVersion < 1 ||
    baseline.title !== input.title ||
    !settled ||
    settled.contentVersion !== baseline.contentVersion
  ) {
    return null;
  }
  try {
    const live = input.readLiveContent();
    if (live === null) return null;
    if (input.compare === "blocks") {
      if (!Array.isArray(live.content)) return null;
      // The blocks this editor last saved or loaded, byte for byte; else the
      // server's own blocks for that version, as the server reads them.
      if (pageDocumentJsonEquals(live.content, baseline.content.content)) return baseline.contentVersion;
      if (!settled.content) return null;
      const project = input.semanticProjection ?? ((content: Record<string, unknown>) => content);
      const bound = input.boundBlockIds ? new Set(input.boundBlockIds) : null;
      return sameServerBlocks(blocksOf(project(live)), blocksOf(project(settled.content)), bound)
        ? baseline.contentVersion
        : null;
    }
    // What this editor last saved or loaded, byte for byte.
    if (pageDocumentJsonEquals(live, baseline.content)) return baseline.contentVersion;
    // A page opened from the server holds the server's stored form, not the
    // editor's own (that form drops the editor's null defaults): compare the
    // way the causal model stores a body, against the server's own receipt.
    return input.semanticProjection && settled.content &&
      sameStoredDocument(live, settled.content, input.semanticProjection)
      ? baseline.contentVersion
      : null;
  } catch {
    return null;
  }
}
