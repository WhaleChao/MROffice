/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A deck held as parts (PPT progressive open, stage 3): the Presentation
 * editor's side of `apps/api/app/services/document_page_parts.py` for a deck,
 * as `lib/document-page-parts` is the Word editor's.
 *
 * The page body holds the deck-level fields (theme, masters, sections,
 * settings), the first slides (part 0) and the `deckParts` marker, which
 * names every slide of the deck in show order and the part it is in. Each
 * later part is a row of its own (`GET .../document-parts/{index}`), which a
 * deck's import may still be writing: `pending` until its slides are read,
 * `failed` when they could not be.
 *
 * The editor holds the whole deck as ONE list of slides, the way it holds an
 * ordinary deck: the body's slides, the slides of every later part it has
 * loaded, and a stub for each slide of a part it has not (`importPending`,
 * the server's own stub shape in `merged_deck_content`). A stub is the
 * original slide, stated by nothing: it can be selected, moved, hidden and
 * deleted (each only changes the order), and nothing else.
 *
 * Every save splits the editor's deck back by the order's parts
 * (`prepareBody`, the `splitDocumentParts` of a deck, the same rules as the
 * server's): the body is the deck-level fields, the slides of part 0 and the
 * order; a slide stays in the part the order put it in, whatever its place;
 * a new slide goes into the body (a later part's slides are exactly the
 * ones its order names, so a slide joins a later part only through the
 * server's own insert). A loaded later part whose slides changed is saved on
 * its own (`PUT .../document-parts/{index}`, compare-and-set on its
 * revision), kept in this tab's session storage until it lands, and merged
 * slide by slide when another save of the part came first -- a slide both
 * changed is kept twice, never dropped.
 *
 * A page without the marker is an ordinary deck: every function here hands
 * it back untouched and the session asks the server for nothing.
 */

import { commentAuthorityText } from "./presentation-deck-commit";
import { expandPresentationMedia } from "./presentation-page-media";
import {
  DECK_PARTS_KEY,
  DECK_PARTS_KIND,
  DOCUMENT_PARTS_SCHEMA,
  deckPartsMarker,
  type DeckOrderEntry,
  type DeckPartsMarker,
  type DocumentPartResponse,
} from "./document-page-parts";

export { deckPartsMarker } from "./document-page-parts";
export type { DeckOrderEntry, DeckPartsMarker } from "./document-page-parts";

export const DECK_PART_TYPE = "mrnote.presentation.part.v1";
/** The stub flag on a slide whose part the editor does not hold. */
export const DECK_STUB_FLAG = "importPending";

const MEDIA_KEY = "media";
const RECEIPTS_KEY = "deckReceipts";
const MEDIA_TOKEN_RE = /"mrai-media:([0-9a-f]{64})"/g;
const RECEIPT_TOKEN_RE = /"mrai-deck:([0-9A-Za-z_-]{1,64})"/g;
const RECOVERY_KEY_PREFIX = "mrnote:deck-part-recovery:v1:";

type Rec = Record<string, unknown>;

function asRec(value: unknown): Rec | null {
  return value && typeof value === "object" && !Array.isArray(value) ? (value as Rec) : null;
}

/** The least a slide is, to the parts: its id, whether it is hidden, whether it is a stub. */
export interface DeckSlideLike {
  readonly id: string;
  readonly hidden?: boolean;
  readonly importPending?: boolean;
}

export function isDeckStubSlide(slide: unknown): boolean {
  return asRec(slide)?.[DECK_STUB_FLAG] === true;
}

/** Whether a page body is a deck held as parts. */
export function isDeckPartsContent(content: unknown): boolean {
  return deckPartsMarker(content) !== null;
}

/**
 * A stub's stored form (what `normalizeSlides` reads): the slide the order
 * names, with what the order states of it and no blocks -- the server's
 * stub in `merged_deck_content`.
 */
export function deckStubSlideRecord(
  entry: DeckOrderEntry,
  deckSize?: { widthEmu?: unknown; heightEmu?: unknown },
  /** The comment authority every slide of the deck carries (a deck-wide
   * receipt, `reconcileSlideCommentAuthority`): a stub carries it too, or a
   * commit read the deck as one without it and took it off every slide. */
  commentAuthority?: unknown,
): Rec {
  const record: Rec = {
    id: entry.id,
    title: typeof entry.title === "string" ? entry.title : "",
    // A stated Blank keeps the stub empty (`blocksLeftOnSlide` gives any
    // other empty slide its body placeholder).
    layout: "blank",
    blocks: [],
    [DECK_STUB_FLAG]: true,
  };
  if (entry.slideId) record.source = { slideId: entry.slideId };
  if (entry.hidden === true) record.hidden = true;
  if (typeof entry.section === "string" && entry.section) record.section = entry.section;
  if (entry.masterId) record.masterId = entry.masterId;
  if (entry.layoutId) record.layoutId = entry.layoutId;
  if (entry.transition) record.transition = entry.transition;
  if (commentAuthority !== undefined) record.commentAuthority = JSON.parse(JSON.stringify(commentAuthority));
  const width = typeof entry.widthEmu === "number" ? entry.widthEmu : deckSize?.widthEmu;
  const height = typeof entry.heightEmu === "number" ? entry.heightEmu : deckSize?.heightEmu;
  if (typeof width === "number" && typeof height === "number") {
    record.widthEmu = width;
    record.heightEmu = height;
  }
  return record;
}

/** The media and receipt entries `value` names (in any string), and the
 * entries those receipts name in turn: what a page holding `value` keeps. */
function namedTableEntries(value: unknown, tables: { media: Rec | null; receipts: Rec | null }): {
  media: Rec | null;
  receipts: Rec | null;
} {
  const media: Rec = {};
  const receipts: Rec = {};
  let pending = [JSON.stringify(value)];
  while (pending.length > 0) {
    const next: string[] = [];
    for (const json of pending) {
      if (tables.media) {
        for (const match of json.matchAll(MEDIA_TOKEN_RE)) {
          const sha = match[1];
          if (Object.hasOwn(tables.media, sha)) media[sha] = tables.media[sha];
        }
      }
      if (tables.receipts) {
        for (const match of json.matchAll(RECEIPT_TOKEN_RE)) {
          const key = match[1];
          if (Object.hasOwn(receipts, key) || !Object.hasOwn(tables.receipts, key)) continue;
          receipts[key] = tables.receipts[key];
          next.push(JSON.stringify(tables.receipts[key]));
        }
      }
    }
    pending = next;
  }
  return {
    media: Object.keys(media).length > 0 ? media : null,
    receipts: Object.keys(receipts).length > 0 ? receipts : null,
  };
}

function contentTables(content: Rec): { media: Rec | null; receipts: Rec | null } {
  return { media: asRec(content[MEDIA_KEY]), receipts: asRec(content[RECEIPTS_KEY]) };
}

/**
 * `content` (a deck as `buildContent` writes it: store form) narrowed to
 * `slides`: the deck-level fields over them, with only the media and receipt
 * entries what remains names.
 */
function deckContentWithSlides(content: Rec, slides: readonly unknown[], extra: Rec): Rec {
  const tables = contentTables(content);
  const out: Rec = {};
  for (const [key, value] of Object.entries(content)) {
    if (key === "slides" || key === MEDIA_KEY || key === RECEIPTS_KEY || key === DECK_PARTS_KEY) continue;
    out[key] = value;
  }
  out.slides = [...slides];
  if (tables.media || tables.receipts) {
    const named = namedTableEntries(out, tables);
    if (named.media) out[MEDIA_KEY] = named.media;
    if (named.receipts) out[RECEIPTS_KEY] = named.receipts;
  }
  return { ...out, ...extra };
}

/** A later part's content for `slides` (store form, from `content`): the
 * slides with the media and receipt entries they name, nothing of the deck. */
function deckPartContentWithSlides(content: Rec, slides: readonly unknown[]): Rec {
  const part: Rec = { type: DECK_PART_TYPE, slides: [...slides] };
  const named = namedTableEntries(slides, contentTables(content));
  if (named.media) part[MEDIA_KEY] = named.media;
  if (named.receipts) part[RECEIPTS_KEY] = named.receipts;
  return part;
}

/** Why the editor does not change a slide. */
export type DeckSlideLock =
  /** Its part is still being read by the import. */
  | "pending"
  /** The import could not read its part (Retry reads it again). */
  | "failed"
  /** Its part is read but this editor has not loaded it yet. */
  | "loading"
  /** Its part is loaded but not saved part by part from here (the server
   * did not say it may be). */
  | "read-only";

export type DeckCommitRefusal =
  | { readonly reason: DeckSlideLock; readonly slideId: string }
  /** A copy of a slide whose content the editor does not have. */
  | { readonly reason: "stub-copy"; readonly slideId: string }
  /** The first part would hold no slide: the server keeps the body's. */
  | { readonly reason: "empty-part"; readonly part: number }
  /** A slide added or deleted while `count` slides are still being read
   * (the server's insert and delete wait for the import). */
  | { readonly reason: "structure-pending"; readonly count: number };

/** The server's insert and delete (`POST .../deck-parts/insert-slides`,
 * `.../delete-slides`) the editor's deck asks for. */
export interface DeckStructuralChange {
  readonly deletes: readonly string[];
  /** Each run of new slides after the slide it follows (null: first). */
  readonly inserts: ReadonlyArray<{
    readonly after: string | null;
    readonly ids: readonly string[];
    /** Whole slides, inline form. */
    readonly slides: readonly Rec[];
  }>;
}

export type DeckPartStatus = "unknown" | "pending" | "ready" | "failed";

interface HeldPart<S> {
  readonly index: number;
  readonly partId: string;
  contentSha256: string | null;
  revision: number;
  editable: boolean;
  /** The part's slides as the server holds them now, as the editor reads
   * them, by id. */
  byId: Map<string, S>;
  /** Each slide as last saved, as the editor showed it then: what a change
   * of it is measured against. */
  readonly saved: Map<string, S>;
  /** Each slide's stored form as last saved (when known). */
  readonly savedStore: Map<string, string>;
  /** The part's live room, when it has one: changes go there (and are saved
   * from it), not by the part's own save. */
  room?: DeckPartRoomPort;
  /** The part as it was installed, until its room is first read: what the
   * editor read it as, and what that reading was of (`readPartInput`). */
  installedRead?: { readonly key: string; readonly slides: S[] };
}

/** What the session needs of a part's room (lib/collab/document-part-room). */
export interface DeckPartRoomPort {
  /** The room is read and takes this member's input. */
  readonly accepting: boolean;
  /** Not read yet, and it will be soon (`roomIsReading`). */
  readonly reading?: boolean;
  /** This member's changes the part has not been saved with yet. */
  readonly unsaved?: boolean;
  /** The room's content (store form): what a change is written over. */
  content(): Record<string, unknown>;
  /** Write the part's content into the room; true when anything changed. */
  publish(content: Record<string, unknown>): boolean;
}

/** A later part's change on its way to the server (`PUT`). */
export interface DeckPartDraft<S = DeckSlideLike> {
  readonly index: number;
  /** The revision the change was made on (`expected_revision`). */
  readonly revision: number;
  /** Its slides, in the order the deck shows them. */
  readonly ids: readonly string[];
  /** `{type, slides, media?, deckReceipts?}`: what is sent. */
  readonly content: Rec;
  readonly json: string;
  /** The editor's slides it was made from (undefined: not known). */
  readonly live: ReadonlyMap<string, S | undefined>;
  readonly store: ReadonlyMap<string, string>;
}

/** What a part's arrival puts in the editor. */
export interface DeckPartArrival<S> {
  /** Its slides by id, each in place of its stub. */
  readonly slides: ReadonlyMap<string, S>;
  /** A slide's edit kept in this tab that the part was saved over meanwhile:
   * shown as a copy after the slide (both kept). */
  readonly copiesAfter: ReadonlyMap<string, S>;
}

/** How the editor brings a part another save moved on in line with it. */
export interface DeckPartMerge<S> {
  /** Slides this editor had not changed: the server's, in their place. */
  readonly replace: ReadonlyMap<string, S>;
  /** Slides both changed: the server's in their place, this editor's own
   * version kept as a copy after it. */
  readonly copiesAfter: ReadonlyMap<string, S>;
}

/** What the window draws from (one object per change: `version`). */
export interface DeckPartsSnapshot {
  readonly active: boolean;
  readonly count: number;
  /** Per part (0 is the body, always ready). */
  readonly statuses: readonly DeckPartStatus[];
  readonly held: readonly boolean[];
  /** Slides of the order whose part the editor does not hold. */
  readonly stubSlides: number;
  /** Of those, slides the import has not read yet. */
  readonly pendingSlides: number;
  readonly failedParts: readonly number[];
  readonly reimportNeeded: readonly number[];
  /** Later parts with changes not saved yet. */
  readonly unsavedParts: readonly number[];
  /** Later parts whose change the server refused (kept until changed). */
  readonly refusedParts: readonly number[];
  readonly version: number;
}

const INACTIVE_SNAPSHOT: DeckPartsSnapshot = {
  active: false,
  count: 1,
  statuses: ["ready"],
  held: [true],
  stubSlides: 0,
  pendingSlides: 0,
  failedParts: [],
  reimportNeeded: [],
  unsavedParts: [],
  refusedParts: [],
  version: 0,
};

/** The stored slides of a ready part answer for part `index`, or null. */
export function readyDeckPartSlides(answer: DocumentPartResponse, index: number, count: number): unknown[] | null {
  if (answer.partIndex !== index || answer.partCount !== count) return null;
  if ((answer.status ?? "ready") !== "ready") return null;
  const content = asRec(answer.content);
  if (!content || !Array.isArray(content.slides)) return null;
  if (content.type !== undefined && content.type !== DECK_PART_TYPE) return null;
  return content.slides as unknown[];
}

interface KeptPartEdit {
  /** The revision the edit was made on. */
  readonly rev: number;
  /** The part content it saves (`DeckPartDraft.content`). */
  readonly content: Rec;
  /** Its save was on its way. */
  readonly sending: boolean;
}

function readKeptEdits(pageId: string): Record<string, KeptPartEdit> {
  try {
    const raw = window.sessionStorage.getItem(`${RECOVERY_KEY_PREFIX}${pageId}`);
    const parsed: unknown = raw ? JSON.parse(raw) : null;
    return asRec(parsed) ? (parsed as Record<string, KeptPartEdit>) : {};
  } catch {
    return {};
  }
}

function writeKeptEdit(pageId: string, index: number, entry: KeptPartEdit | null): void {
  try {
    const all = readKeptEdits(pageId);
    if (entry) all[String(index)] = entry;
    else delete all[String(index)];
    const key = `${RECOVERY_KEY_PREFIX}${pageId}`;
    if (Object.keys(all).length === 0) window.sessionStorage.removeItem(key);
    else window.sessionStorage.setItem(key, JSON.stringify(all));
  } catch {
    // Storage full or unavailable: the edit is still saved as usual; it
    // only would not come back after a reload before that.
  }
}

/** A slide record as a save compares it. Hidden or shown is the order's alone:
 * no part save for it. */
function slideRecordJson(record: Rec): string {
  const { hidden: _hidden, ...stated } = record;
  void _hidden;
  return JSON.stringify(stated);
}

/**
 * Whether a live slide is what its part last saved, asked of every held slide
 * of a deck on each commit. A slide object does not change once made (a change
 * is a new object), so the answer is kept with the object: this was 23 ms of
 * every key on a 56-slide deck, each slide written out twice.
 */
const comparedWithSaved = new WeakMap<object, { saved: object; same: boolean }>();

function unchangedFromSaved(now: DeckSlideLike, saved: DeckSlideLike): boolean {
  const kept = comparedWithSaved.get(now);
  if (kept && kept.saved === saved) return kept.same;
  const same = sameSlideButVisibility(now, saved);
  comparedWithSaved.set(now, { saved, same });
  return same;
}

function roomKeepsWords(held: { room?: DeckPartRoomPort }): boolean {
  return held.room?.reading === true || held.room?.unsaved === true;
}





const writtenWithoutVisibility = new WeakMap<object, string>();

function slideWithoutVisibilityJson(slide: DeckSlideLike): string {
  const kept = writtenWithoutVisibility.get(slide);
  if (kept !== undefined) return kept;
  const { hidden: _hidden, ...rest } = slide as DeckSlideLike & Rec;
  void _hidden;
  const json = JSON.stringify(rest);
  writtenWithoutVisibility.set(slide, json);
  return json;
}

function sameSlideButVisibility(a: DeckSlideLike, b: DeckSlideLike): boolean {
  if (a === b) return true;
  return slideWithoutVisibilityJson(a) === slideWithoutVisibilityJson(b);
}

function sameIds(a: readonly string[], b: readonly string[]): boolean {
  if (a.length !== b.length) return false;
  const set = new Set(a);
  return b.every((id) => set.has(id));
}

/**
 * The parts of one deck page, as one editor holds them. Not React state: the
 * window reads it while it reads a page (`viewSlides`), commits an edit
 * (`checkCommit`) and saves (`prepareBody`); `useDeckPageParts` loads the
 * parts, saves the changed ones and tells the window when one arrived.
 */
export class DeckPartsSession<S extends DeckSlideLike> {
  private pageId: string | null = null;
  private marker: DeckPartsMarker | null = null;
  private bodyTables: { media: Rec | null; receipts: Rec | null; size: { widthEmu?: unknown; heightEmu?: unknown } } = {
    media: null,
    receipts: null,
    size: {},
  };
  private readonly parts = new Map<number, HeldPart<S>>();
  private readonly statuses = new Map<number, DeckPartStatus>();
  private readonly reasons = new Map<number, string>();
  /** Every slide's part as any order this session read named it. A slide
   * deleted from the order keeps its part here: Undo may bring it back, and
   * it is still that part's slide. */
  private readonly ownerOf = new Map<string, number>();
  /** What the latest order (or the last that named it) states of a slide. */
  private readonly entryOf = new Map<string, DeckOrderEntry>();
  private readonly stubs = new Map<string, { readonly key: string; readonly slide: S }>();
  private readonly hiddenVariants = new WeakMap<S, { shown?: S; hidden?: S }>();
  private readonly drafts = new Map<number, DeckPartDraft<S>>();
  /** A change the server refused as content (422): not sent again until it
   * changes. */
  private readonly refused = new Map<number, { readonly json: string; readonly message: string }>();
  private readonly listeners = new Set<() => void>();
  private snapshotValue: DeckPartsSnapshot = INACTIVE_SNAPSHOT;
  private versionValue = 0;

  constructor(private readonly normalize: (content: Rec) => S[]) {}

  /**
   * Slides another member's typing has reached this window in (a room's text
   * change names its slide). Read to tell whether an Undo would take someone's
   * words away; a slide's content is no answer, since the server's reading of a
   * new slide differs from the one the editor made without anyone typing.
   */
  private readonly remoteWords = new Set<string>();

  noteRemoteWords(slideId: string): void {
    this.remoteWords.add(slideId);
  }

  hasRemoteWords(slideId: string): boolean {
    return this.remoteWords.has(slideId);
  }

  /** Words the server holds that differ, block by block, from what this window
   * shows: this member's own input is in both (published again after a reseed,
   * or in the save that moved the part), so they are someone else's. A block the
   * server read otherwise has another id, and says nothing. */
  private noteWordsMoved(shownBefore: readonly S[], slides: readonly S[]): void {
    type Words = { blocks?: ReadonlyArray<{ id?: unknown; text?: unknown }> };
    const shown = new Map(shownBefore.map((slide) => [slide.id, slide as unknown as Words]));
    for (const slide of slides) {
      const blocks = shown.get(slide.id)?.blocks;
      if (!blocks) continue;
      const words = new Map(blocks.map((block) => [block?.id, block?.text]));
      if (((slide as unknown as Words).blocks ?? []).some((block) => {
        const was = words.get(block?.id);
        return typeof was === "string" && typeof block?.text === "string" && was !== block.text;
      })) this.noteRemoteWords(slide.id);
    }
  }

  /** A different page (or the same page opened again): nothing is held. */
  reset(pageId: string | null): void {
    if (this.pageId === pageId && this.marker === null && this.parts.size === 0 && this.drafts.size === 0) return;
    this.pageId = pageId;
    this.marker = null;
    this.remoteWords.clear();
    this.bodyTables = { media: null, receipts: null, size: {} };
    this.parts.clear();
    this.statuses.clear();
    this.reasons.clear();
    this.ownerOf.clear();
    this.entryOf.clear();
    this.stubs.clear();
    this.drafts.clear();
    this.refused.clear();
    this.stale.clear();
    this.rereading.clear();
    this.rebased(null);
    this.changed();
  }

  get page(): string | null {
    return this.pageId;
  }

  /** Whether the page read last is a deck held as parts. */
  get active(): boolean {
    return this.marker !== null;
  }

  get currentMarker(): DeckPartsMarker | null {
    return this.marker;
  }

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  snapshot(): DeckPartsSnapshot {
    return this.snapshotValue;
  }

  /** Readers of the snapshot told of everything installed quietly so far (see
   * installPart): one render for parts that arrive together. */
  flushNotices(): void {
    if (this.quietChange) this.changed();
  }

  private quietChange = false;

  private changed(): void {
    this.quietChange = false;
    this.versionValue += 1;
    const marker = this.marker;
    if (!marker) {
      this.snapshotValue = { ...INACTIVE_SNAPSHOT, version: this.versionValue };
    } else {
      const statuses = Array.from({ length: marker.count }, (_, index) =>
        index === 0 ? "ready" : (this.statuses.get(index) ?? "unknown")) as DeckPartStatus[];
      const held = Array.from({ length: marker.count }, (_, index) => index === 0 || this.parts.has(index));
      let stubSlides = 0;
      let pendingSlides = 0;
      for (const entry of marker.order) {
        if (entry.part === 0 || this.parts.has(entry.part)) continue;
        stubSlides += 1;
        if (statuses[entry.part] !== "ready") pendingSlides += 1;
      }
      this.snapshotValue = {
        active: true,
        count: marker.count,
        statuses,
        held,
        stubSlides,
        pendingSlides,
        failedParts: statuses.flatMap((status, index) => (status === "failed" ? [index] : [])),
        reimportNeeded: [...(marker.reimportNeeded ?? [])],
        unsavedParts: [...this.drafts.keys()].sort((a, b) => a - b),
        refusedParts: [...this.refused.keys()].sort((a, b) => a - b),
        version: this.versionValue,
      };
    }
    for (const listener of [...this.listeners]) listener();
  }

  /**
   * Read a page body the server sent (an open, a save's answer, a
   * collaborator's delivery). Returns whether it is a deck held as parts.
   * A body without the marker empties the session (the deck was made whole).
   */
  adoptBody(pageId: string | null, content: unknown): boolean {
    if (pageId !== this.pageId) this.reset(pageId);
    const marker = deckPartsMarker(content);
    if (!marker) {
      if (this.marker) {
        this.marker = null;
        this.parts.clear();
        this.statuses.clear();
        this.stubs.clear();
        this.drafts.clear();
        this.refused.clear();
        this.changed();
      }
      return false;
    }
    const record = asRec(content)!;
    const previous = this.marker;
    if (previous && previous.count !== marker.count) this.renumberParts(marker);
    // What the server listed or dropped since the editor's base was taken
    // (a body read here without a view of it: a structure change's answer,
    // another editor's insert or delete): never the editor's own.
    for (const entry of marker.order) this.listedSinceBase.add(entry.id);
    if (previous) {
      const now = new Set(marker.order.map((entry) => entry.id));
      for (const entry of previous.order) if (!now.has(entry.id)) this.droppedSinceBase.add(entry.id);
    }
    this.marker = marker;
    this.serverBodySlides = new Map(
      (Array.isArray(record.slides) ? record.slides : [])
        .flatMap((slide) => (asRec(slide) && typeof asRec(slide)!.id === "string" ? [[asRec(slide)!.id as string, asRec(slide)!]] : [])) as Array<[string, Rec]>,
    );
    const firstSlide = Array.isArray(record.slides) ? asRec(record.slides[0]) : null;
    this.bodyTables = {
      ...contentTables(record),
      size: { widthEmu: firstSlide?.widthEmu, heightEmu: firstSlide?.heightEmu },
    };
    for (const entry of marker.order) {
      this.ownerOf.set(entry.id, entry.part);
      this.entryOf.set(entry.id, entry);
    }
    if (previous === null || JSON.stringify(previous) !== JSON.stringify(marker)) this.changed();
    return true;
  }

  /**
   * The parts were numbered again (a delete emptied one and the later ones
   * moved up, or the deck was cut otherwise): each held part, its change and
   * its status follow its slides to the part the new order puts them in; a
   * part whose slides are now in more than one part is read again.
   */
  private renumberParts(marker: DeckPartsMarker): void {
    const partOf = new Map(marker.order.map((entry) => [entry.id, entry.part]));
    const moved = new Map<number, number>();
    for (const [index, held] of this.parts) {
      const targets = new Set<number>();
      for (const id of held.byId.keys()) {
        const part = partOf.get(id);
        if (part !== undefined) targets.add(part);
      }
      if (targets.size === 1) moved.set(index, [...targets][0]);
    }
    const parts = new Map(this.parts);
    const drafts = new Map(this.drafts);
    const refused = new Map(this.refused);
    this.parts.clear();
    this.drafts.clear();
    this.refused.clear();
    this.statuses.clear();
    this.stubs.clear();
    for (const [from, to] of moved) {
      if (to === 0 || this.parts.has(to)) continue;
      const held = parts.get(from)!;
      this.parts.set(to, { ...held, index: to });
      this.statuses.set(to, "ready");
      const draft = drafts.get(from);
      if (draft) this.drafts.set(to, { ...draft, index: to });
      const refusal = refused.get(from);
      if (refusal) this.refused.set(to, refusal);
    }
    this.ownerOf.clear();
    this.entryOf.clear();
    if (this.pageId) {
      // Kept edits are named by part: a renumbered part's are dropped (its
      // change is in the editor, and is queued again from there).
      for (const index of parts.keys()) writeKeptEdit(this.pageId, index, null);
    }
  }

  /** The part the order puts slide `id` in (undefined: a new slide). */
  partOf(id: string): number | undefined {
    return this.marker ? this.ownerOf.get(id) : undefined;
  }

  status(index: number): DeckPartStatus {
    return index === 0 ? "ready" : (this.statuses.get(index) ?? "unknown");
  }

  failureReason(index: number): string | null {
    return this.reasons.get(index) ?? null;
  }

  refusalOf(index: number): string | null {
    return this.refused.get(index)?.message ?? null;
  }

  isHeld(index: number): boolean {
    return index === 0 || this.parts.has(index);
  }

  /** Why slide `id` is not edited here, or null. */
  lockOf(id: string): DeckSlideLock | null {
    const part = this.partOf(id);
    if (part === undefined || part === 0) return null;
    const held = this.parts.get(part);
    if (held) return held.editable ? null : "read-only";
    // Only being read again: it is writable, and what is typed meanwhile stays
    // over the reading when it comes (`typedOverArrival`).
    if (this.rereading.has(part)) return null;
    const status = this.status(part);
    return status === "pending" ? "pending" : status === "failed" ? "failed" : "loading";
  }

  /** The parts holding these slides that the editor has not loaded. */
  partsToLoadFor(ids: Iterable<string>): number[] {
    const wanted = new Set<number>();
    for (const id of ids) {
      const part = this.partOf(id);
      if (part !== undefined && part > 0 && !this.parts.has(part)) wanted.add(part);
    }
    return [...wanted].sort((a, b) => a - b);
  }

  setStatus(index: number, status: DeckPartStatus, reason?: string | null): void {
    if (!this.marker || index < 1 || index >= this.marker.count) return;
    const before = this.statuses.get(index);
    const beforeReason = this.reasons.get(index) ?? null;
    this.statuses.set(index, status);
    if (status === "pending" || status === "failed") this.rereading.delete(index);
    if (status === "failed" && reason) this.reasons.set(index, reason);
    else this.reasons.delete(index);
    if (before !== status || beforeReason !== (this.reasons.get(index) ?? null)) this.changed();
  }

  private readPart(content: Rec): S[] {
    return this.withDeckCommentAuthority(this.normalize(this.readPartInput(content)));
  }

  /**
   * A part's slides holding the deck's comment authority in the body's own
   * spelling: the server keeps the same receipt on every slide but may write
   * its keys in another order, and a commit reads two spellings as two
   * authorities when their spellings are not one (`commentAuthorityText`),
   * and a slide re-spelled at a commit is a changed slide, its part saved
   * again for nothing. A slide that states none or another receipt is kept
   * as it is: the receipt is the deck's custody, never made up here (the
   * server refuses such a part, `slides[i].commentAuthority`).
   */
  private withDeckCommentAuthority(slides: S[]): S[] {
    const deck = this.commentAuthority;
    if (deck === undefined) return slides;
    const canonical = commentAuthorityText(deck);
    return slides.map((slide) => {
      const own = (slide as unknown as Rec).commentAuthority;
      if (own === undefined || JSON.stringify(own) === this.commentAuthorityKey) return slide;
      if (commentAuthorityText(own) !== canonical) return slide;
      return { ...slide, commentAuthority: JSON.parse(JSON.stringify(deck)) } as S;
    });
  }

  /** What a part's content is read as by `normalize`: its slides with the
   * part's pictures and receipts. */
  private readPartInput(content: Rec): Rec {
    // The part's pictures and receipts are its own; one it shares with the
    // deck is the body's (the part's table first, then the body's).
    const media = { ...(this.bodyTables.media ?? {}), ...(asRec(content[MEDIA_KEY]) ?? {}) };
    const receipts = { ...(this.bodyTables.receipts ?? {}), ...(asRec(content[RECEIPTS_KEY]) ?? {}) };
    return {
      type: "mrnote.presentation.v1",
      slides: Array.isArray(content.slides) ? content.slides : [],
      ...(Object.keys(media).length > 0 ? { [MEDIA_KEY]: media } : {}),
      ...(Object.keys(receipts).length > 0 ? { [RECEIPTS_KEY]: receipts } : {}),
    };
  }

  /** The held slide as the editor shows it: the visibility the order gives it. */
  private shown(slide: S): S {
    return this.withHidden(slide, this.entryOf.get(slide.id)?.hidden === true);
  }

  /**
   * A later part as the server sent it (`GET .../document-parts/{index}`):
   * held from now on. Returns what the editor puts in place of its stubs, or
   * null when the answer is not that part's ready content (its status is
   * kept). An edit of the part kept in this tab from before a reload comes
   * back with it.
   */
  installPart(
    index: number,
    answer: DocumentPartResponse,
    /** The slides this window shows: a part read again after its revision moved
     * can hold words no room told it of. */
    shownBefore?: readonly S[],
    /** Readers are told by flushNotices, with the other parts of this batch. */
    quiet = false,
  ): DeckPartArrival<S> | null {
    const marker = this.marker;
    if (!marker || index < 1 || index >= marker.count || this.parts.has(index)) return null;
    const status = answer.status ?? "ready";
    if (answer.partIndex !== index || answer.partCount !== marker.count) return null;
    if (status !== "ready") {
      this.setStatus(index, status, answer.reason ?? null);
      return null;
    }
    if (!readyDeckPartSlides(answer, index, marker.count)) return null;
    const input = this.readPartInput(asRec(answer.content)!);
    const read = this.normalize(input);
    const slides = this.withDeckCommentAuthority(read);
    // A slide changed since the part was let go is this member's own typing.
    if (shownBefore) {
      this.noteWordsMoved(shownBefore.filter((slide) => {
        const earlier = this.stale.get(slide.id);
        return !earlier || sameSlideButVisibility(earlier, slide);
      }), slides);
    }
    const revision = typeof answer.revision === "number" ? answer.revision : 0;
    const held: HeldPart<S> = {
      index,
      partId: answer.partId,
      contentSha256: typeof answer.contentSha256 === "string" ? answer.contentSha256 : null,
      revision,
      editable: answer.editable === true,
      byId: new Map(slides.map((slide) => [slide.id, slide])),
      saved: new Map(),
      savedStore: new Map(),
      // Its room's first reading is, nearly always, this same content: read
      // once, not twice (a part's schema check is most of the cost of
      // opening a long deck, and every part's room opens after it loads).
      ...(answer.editable === true ? { installedRead: { key: JSON.stringify(input), slides: read } } : {}),
    };
    for (const slide of slides) held.saved.set(slide.id, this.shown(slide));
    this.parts.set(index, held);
    this.rereading.delete(index);
    this.statuses.set(index, "ready");
    this.reasons.delete(index);
    for (const slide of slides) {
      if (!this.ownerOf.has(slide.id)) this.ownerOf.set(slide.id, index);
      this.stubs.delete(slide.id);
    }
    const shown = new Map<string, S>();
    for (const slide of slides) shown.set(slide.id, held.saved.get(slide.id)!);
    const copiesAfter = new Map<string, S>();
    const kept = this.pageId ? readKeptEdits(this.pageId)[String(index)] : undefined;
    if (kept && this.pageId) {
      const keptSlides = asRec(kept.content) ? this.readPart(kept.content) : [];
      if (kept.rev === revision && keptSlides.length > 0) {
        // Not saved before the reload: back in, and saved again.
        for (const slide of keptSlides) {
          if (shown.has(slide.id)) shown.set(slide.id, this.shown(slide));
        }
      } else if (!(kept.sending && kept.rev + 1 === revision)) {
        // The part was saved from elsewhere since: the server's slides in
        // place, each of this tab's that differs kept as a copy after it.
        for (const slide of keptSlides) {
          const now = shown.get(slide.id);
          if (now && !sameSlideButVisibility(now, this.shown(slide))) copiesAfter.set(slide.id, this.shown(slide));
        }
        writeKeptEdit(this.pageId, index, null);
      } else {
        writeKeptEdit(this.pageId, index, null);
      }
    }
    if (quiet) this.quietChange = true;
    else this.changed();
    return { slides: shown, copiesAfter };
  }

  /** The held slide with the visibility the order gives it. */
  private withHidden(slide: S, hidden: boolean): S {
    if ((slide.hidden === true) === hidden) return slide;
    const variants = this.hiddenVariants.get(slide) ?? {};
    const key = hidden ? "hidden" : "shown";
    let variant = variants[key];
    if (!variant) {
      const copy = { ...slide } as S & { hidden?: boolean };
      if (hidden) copy.hidden = true;
      else delete copy.hidden;
      variant = copy as S;
      variants[key] = variant;
      this.hiddenVariants.set(slide, variants);
    }
    return variant;
  }

  private stubFor(entry: DeckOrderEntry): S {
    const key = JSON.stringify([entry, this.commentAuthorityKey]);
    const kept = this.stubs.get(entry.id);
    if (kept && kept.key === key) return kept.slide;
    const [slide] = this.normalize({
      type: "mrnote.presentation.v1",
      slides: [deckStubSlideRecord(entry, this.bodyTables.size, this.commentAuthority)],
    });
    this.stubs.set(entry.id, { key, slide });
    return slide;
  }

  /**
   * The deck the editor shows for `content` of this session's page: an
   * ordinary page's slides as `normalizeBody` reads them; a deck held as
   * parts' every slide in the order's show order -- the body's, each held
   * part's (as the editor last had it, when `current` holds it), a stub for
   * the rest. Never drops a body slide (one the order does not name stays,
   * last).
   */
  viewSlides(content: unknown, normalizeBody: (content: Rec) => S[], current?: readonly S[]): S[] {
    const record = asRec(content);
    // Only a body with the marker says anything of the parts: content the
    // editor made itself (a slide read alone, a plan) is read as it is.
    const parts = deckPartsMarker(content) !== null && this.adoptBody(this.pageId, content);
    const body = normalizeBody(record ?? {});
    const marker = this.marker;
    if (!parts || !marker) return body;
    const bodyById = new Map(body.map((slide) => [slide.id, slide]));
    // The deck's comment authority, as the editor reads it: every slide of
    // the deck carries the same one, so its stubs do too.
    const authorities = body.map((slide) => (slide as unknown as Rec).commentAuthority);
    const authorityKey = authorities[0] === undefined ? null : JSON.stringify(authorities[0]);
    const shared = authorityKey !== null &&
      authorities.every((value) => value !== undefined && JSON.stringify(value) === authorityKey);
    this.commentAuthority = shared ? authorities[0] : undefined;
    this.commentAuthorityKey = shared ? authorityKey : null;
    // A later part's slide is the body's room's no more than the server's
    // part is: an edit of it not saved yet stays as the editor has it.
    const currentById = new Map<string, S>();
    for (const slide of current ?? []) {
      const part = this.ownerOf.get(slide.id);
      if (part !== undefined && part > 0 && this.drafts.has(part)) currentById.set(slide.id, slide);
    }
    // The editor's own inserts and deletes not landed yet stay as the editor
    // has them: the server's slides it has not deleted, in the server's
    // order, with each slide it added after the slide it follows. What the
    // server added or removed since the editor last read it comes or goes.
    const server = new Set(marker.order.map((entry) => entry.id));
    const base = this.baseIds ?? server;
    const deletedHere = new Set<string>();
    const addedAfter = new Map<string | null, S[]>();
    if (current) {
      const currentIds = new Set(current.map((slide) => slide.id));
      for (const id of base) {
        if (server.has(id) && !currentIds.has(id) && !this.droppedSinceBase.has(id)) deletedHere.add(id);
      }
      let anchor: string | null = null;
      for (const slide of current) {
        if (server.has(slide.id)) {
          anchor = slide.id;
          continue;
        }
        if (base.has(slide.id) || this.listedSinceBase.has(slide.id) || slide.importPending === true) continue;
        const list = addedAfter.get(anchor) ?? [];
        list.push(slide);
        addedAfter.set(anchor, list);
      }
    }
    this.rebased(server);
    const placed = new Set<string>();
    const out: S[] = [...(addedAfter.get(null) ?? [])];
    for (const slide of out) placed.add(slide.id);
    const follow = (id: string) => {
      for (const slide of addedAfter.get(id) ?? []) {
        if (placed.has(slide.id)) continue;
        out.push(slide);
        placed.add(slide.id);
      }
    };
    for (const entry of marker.order) {
      if (deletedHere.has(entry.id)) {
        follow(entry.id);
        continue;
      }
      if (entry.part === 0) {
        const slide = bodyById.get(entry.id);
        if (slide && !placed.has(slide.id)) {
          out.push(slide);
          placed.add(slide.id);
        }
        follow(entry.id);
        continue;
      }
      const editing = currentById.get(entry.id);
      const holding = this.parts.get(entry.part);
      const held = holding?.byId.get(entry.id);
      // A slide the server put in a part this editor holds (another
      // editor's insert or copy): the part is read again for it.
      if (holding && !held && !editing) this.incomplete.add(entry.part);
      out.push(
        editing
          ? this.withHidden(editing, entry.hidden === true)
          : held
            ? this.withHidden(held, entry.hidden === true)
            : this.stubFor(entry),
      );
      placed.add(entry.id);
      follow(entry.id);
    }
    for (const slide of body) if (!placed.has(slide.id)) out.push(slide);
    return out;
  }

  private commentAuthority: unknown = undefined;
  private commentAuthorityKey: string | null = null;

  /** The slides the editor last read the server's order with (what its own
   * inserts and deletes are measured from). */
  private baseIds: ReadonlySet<string> | null = null;
  






  private listedSinceBase = new Set<string>();
  private droppedSinceBase = new Set<string>();
  /** `baseIds` taken afresh: nothing listed or dropped since. */
  private rebased(ids: ReadonlySet<string> | null): void {
    this.baseIds = ids;
    this.listedSinceBase = new Set();
    this.droppedSinceBase = new Set();
  }

  /** The slide `id` as its part (or its stub) holds it, with the order's
   * visibility: what a locked slide is compared with. */
  private originalOf(id: string): S | undefined {
    const part = this.partOf(id);
    if (part === undefined || part === 0) return undefined;
    const entry = this.entryOf.get(id);
    const held = this.parts.get(part)?.byId.get(id);
    if (held) return this.withHidden(held, entry?.hidden === true);
    return entry ? this.stubFor(entry) : undefined;
  }

  /**
   * The rail after a body read for its order alone (an insert, a delete or a
   * copy somewhere): the slides the server now lists, in its order, each slide
   * on screen left as it is -- this reading says where slides stand, and a
   * slide's words come through its own room. `current` itself when nothing
   * moved, so nothing is drawn again.
   */
  followOrder(content: unknown, normalizeBody: (content: Rec) => S[], current: readonly S[]): readonly S[] {
    const viewed = this.viewSlides(content, normalizeBody, current);
    if (viewed.length === current.length && viewed.every((slide, at) => slide.id === current[at].id)) return current;
    const shown = new Map(current.map((slide) => [slide.id, slide]));
    return viewed.map((slide) => {
      const held = shown.get(slide.id);
      return held && held.importPending !== true ? held : slide;
    });
  }

  /**
   * Whether an edit that turns the deck `before` into `after` may be made:
   * a locked slide (`lockOf`) changes at most its visibility (it may also
   * move, or go); no copy of a stub is made; the first part keeps a slide
   * (the server keeps the body's first slides); and no slide is added or
   * deleted while a part is still being read (the server's insert and
   * delete wait for the import). Null when it may.
   */
  checkCommit(before: readonly S[], after: readonly S[]): DeckCommitRefusal | null {
    const marker = this.marker;
    if (!marker) return null;
    const beforeById = new Map(before.map((slide) => [slide.id, slide]));
    for (const slide of after) {
      const lock = this.lockOf(slide.id);
      if (!lock) {
        if (slide.importPending === true && this.partOf(slide.id) === undefined) {
          return { reason: "stub-copy", slideId: slide.id };
        }
        continue;
      }
      const was = beforeById.get(slide.id) ?? this.originalOf(slide.id);
      if (was === slide) continue;
      if (!was || !sameSlideButVisibility(was, slide)) return { reason: lock, slideId: slide.id };
    }
    const afterIds = new Set(after.map((slide) => slide.id));
    const added = after.some((slide) => !beforeById.has(slide.id));
    const removed = before.some((slide) => !afterIds.has(slide.id));
    if ((added || removed) && this.unreadSlides() > 0) {
      return { reason: "structure-pending", count: this.unreadSlides() };
    }
    if (!after.some((slide) => (this.partOf(slide.id) ?? 0) === 0)) return { reason: "empty-part", part: 0 };
    return null;
  }

  /** Slides of parts the import has not read (pending) or could not (failed). */
  unreadSlides(): number {
    const marker = this.marker;
    if (!marker) return 0;
    let count = 0;
    for (const entry of marker.order) {
      if (entry.part === 0 || this.parts.has(entry.part)) continue;
      const status = this.status(entry.part);
      if (status === "pending" || status === "failed") count += 1;
    }
    return count;
  }

  /** Held parts the server's order names slides of that they do not hold. */
  private readonly incomplete = new Set<number>();

  /** The held parts to read again for slides they do not hold yet (none
   * with a change of this editor's on its way: that one lands first). */
  takeIncompleteParts(): number[] {
    const out: number[] = [];
    for (const index of this.incomplete) {
      const held = this.parts.get(index);
      if (!held) {
        this.incomplete.delete(index);
        continue;
      }
      if (this.drafts.has(index)) continue;
      this.incomplete.delete(index);
      for (const [id, slide] of held.saved) this.stale.set(id, slide);
      this.parts.delete(index);
      this.statuses.set(index, "unknown");
      out.push(index);
    }
    if (out.length > 0) this.changed();
    return out;
  }

  /** The slides the server's order holds (its membership), by id. */
  private serverIds(): Set<string> {
    return new Set((this.marker?.order ?? []).map((entry) => entry.id));
  }

  /**
   * The body to save for the deck `content` the editor wrote (`buildContent`,
   * every slide it holds, stubs included). The body is saved against the
   * slides the server holds: the deck-level fields, the slides of part 0 and
   * the order of every slide it holds, each in its part, as the editor
   * orders and hides them. A slide the editor added or deleted is the
   * server's insert or delete (`structuralChange`, sent by useDeckPageParts);
   * until that lands the body keeps the server's slides as they are. Each
   * loaded later part whose slides changed is queued to be saved on its own
   * (`live`: the editor's slides `content` was made from). Content of an
   * ordinary page, or one already a body, comes back as it is.
   */
  prepareBody(content: Rec, live?: readonly S[]): Rec {
    const marker = this.marker;
    if (!marker || content[DECK_PARTS_KEY] !== undefined || !Array.isArray(content.slides)) return content;
    const server = this.serverIds();
    const records = new Map<string, Rec>();
    const editorIds: string[] = [];
    const loose: unknown[] = [];
    for (const item of content.slides) {
      const slide = asRec(item);
      const id = typeof slide?.id === "string" ? slide.id : null;
      if (!slide || !id) {
        loose.push(item);
        continue;
      }
      // A stub under an id the server does not hold is a copy the editor
      // refuses to make (`checkCommit`): nothing of it is saved.
      if (slide[DECK_STUB_FLAG] === true && !server.has(id)) continue;
      if (records.has(id)) continue;
      records.set(id, slide);
      editorIds.push(id);
    }
    // The server's slides in the editor's order; one the editor deleted
    // stays after the slide it followed until the delete lands.
    const deleted = marker.order.filter((entry) => !records.has(entry.id)).map((entry) => entry.id);
    const kept = new Set(editorIds.filter((id) => server.has(id)));
    const deletedAfter = new Map<string | null, string[]>();
    let previous: string | null = null;
    for (const entry of marker.order) {
      if (kept.has(entry.id)) {
        previous = entry.id;
        continue;
      }
      const list = deletedAfter.get(previous) ?? [];
      list.push(entry.id);
      deletedAfter.set(previous, list);
    }
    const orderedIds: string[] = [...(deletedAfter.get(null) ?? [])];
    for (const id of editorIds) {
      if (!kept.has(id)) continue;
      orderedIds.push(id);
      orderedIds.push(...(deletedAfter.get(id) ?? []));
    }
    const bodySlides: unknown[] = [...loose];
    const order: DeckOrderEntry[] = [];
    const partRecords = new Map<number, Array<{ id: string; record: Rec }>>();
    for (const id of orderedIds) {
      const part = this.ownerOf.get(id) ?? 0;
      const record = records.get(id) ?? null;
      if (part === 0) {
        const slide = record ?? this.serverBodySlides.get(id);
        if (!slide) continue;
        bodySlides.push(slide);
        order.push({ id, part: 0 });
        continue;
      }
      const stated = this.entryOf.get(id);
      const entry: Record<string, unknown> = { id, part };
      for (const [key, value] of Object.entries(stated ?? {})) {
        if (key !== "id" && key !== "part" && key !== "hidden" && key !== "section") entry[key] = value;
      }
      // Its visibility and its section are the order's to state (the
      // editor's, for a slide it has).
      const hidden = record ? record.hidden === true : stated?.hidden === true;
      if (hidden) entry.hidden = true;
      const section = record ? record.section : stated?.section;
      if (typeof section === "string" && section) entry.section = section.slice(0, 500);
      order.push(entry as unknown as DeckOrderEntry);
      if (record && record[DECK_STUB_FLAG] !== true && this.parts.has(part)) {
        const list = partRecords.get(part) ?? [];
        list.push({ id, record });
        partRecords.set(part, list);
      }
    }
    const base = this.baseIds ?? server;
    this.queueStructuralChange(
      content,
      editorIds,
      records,
      server,
      base,
      // Only what the editor deleted: a slide the server added since it last
      // read the order, and the editor has not shown yet, is not one.
      deleted.filter((id) => base.has(id) && !this.droppedSinceBase.has(id)),
    );
    this.queuePartDrafts(content, partRecords, live);
    const nextMarker: Rec = {
      schema: DOCUMENT_PARTS_SCHEMA,
      kind: DECK_PARTS_KIND,
      count: marker.count,
      order,
      ...(marker.reimportNeeded && marker.reimportNeeded.length > 0 ? { reimportNeeded: [...marker.reimportNeeded] } : {}),
    };
    // A deleted body slide is saved as the server holds it: its pictures
    // and receipts are the server's body's.
    const tables: Rec = { ...content };
    if (deleted.some((id) => (this.ownerOf.get(id) ?? 0) === 0)) {
      tables[MEDIA_KEY] = { ...(this.bodyTables.media ?? {}), ...(asRec(content[MEDIA_KEY]) ?? {}) };
      tables[RECEIPTS_KEY] = { ...(this.bodyTables.receipts ?? {}), ...(asRec(content[RECEIPTS_KEY]) ?? {}) };
    }
    return deckContentWithSlides(tables, bodySlides, { [DECK_PARTS_KEY]: nextMarker });
  }

  /** The server's insert and delete the editor's deck asks for. */
  private queueStructuralChange(
    content: Rec,
    editorIds: readonly string[],
    records: ReadonlyMap<string, Rec>,
    server: ReadonlySet<string>,
    base: ReadonlySet<string>,
    deleted: readonly string[],
  ): void {
    const inserts: Array<{ after: string | null; ids: string[] }> = [];
    let anchor: string | null = null;
    let run: { after: string | null; ids: string[] } | null = null;
    for (const id of editorIds) {
      if (server.has(id)) {
        anchor = id;
        run = null;
        continue;
      }
      // A slide the server removed since the editor read the order: it goes
      // with the next reading, not back to the server. So does one the
      // server listed meanwhile (another editor's, taken back since).
      if (base.has(id) || this.listedSinceBase.has(id)) continue;
      if (!run) {
        run = { after: anchor, ids: [] };
        inserts.push(run);
      }
      run.ids.push(id);
    }
    const tables = contentTables(content);
    const change: DeckStructuralChange | null = inserts.length === 0 && deleted.length === 0
      ? null
      : {
          deletes: [...deleted],
          inserts: inserts.map(({ after, ids }) => {
            // Sent whole, its pictures and receipts written out on it: the
            // part it joins holds its own tables.
            const inline = expandPresentationMedia({
              slides: ids.map((id) => records.get(id)),
              ...(tables.media ? { [MEDIA_KEY]: tables.media } : {}),
              ...(tables.receipts ? { [RECEIPTS_KEY]: tables.receipts } : {}),
            }) as { slides: Rec[] };
            return { after, ids, slides: inline.slides };
          }),
        };
    const json = change ? JSON.stringify(change) : null;
    if (json === this.structuralJson) return;
    this.structural = change;
    this.structuralJson = json;
    this.changed();
  }

  private structural: DeckStructuralChange | null = null;
  private structuralJson: string | null = null;
  /** The server's body slides as it last sent them (store form), by id. */
  private serverBodySlides = new Map<string, Rec>();

  /**
   * What an Undo of deleting `ids` sends back to the server, read before the
   * delete: a body slide as the server holds it now, whole (the server keeps
   * no copy of a body slide it deletes); a later part's slide by its id alone
   * (its part keeps it, and the server restores that copy).
   */
  restorableSlides(ids: readonly string[]): Rec[] {
    return ids.map((id) => {
      if ((this.ownerOf.get(id) ?? 0) > 0) return { id };
      const record = this.serverBodySlides.get(id);
      if (!record) return { id };
      const inline = expandPresentationMedia({
        slides: [record],
        ...(this.bodyTables.media ? { [MEDIA_KEY]: this.bodyTables.media } : {}),
        ...(this.bodyTables.receipts ? { [RECEIPTS_KEY]: this.bodyTables.receipts } : {}),
      }) as { slides: Rec[] };
      return inline.slides[0] ?? { id };
    });
  }

  /** The insert and delete still to send, or null. */
  structuralChange(): DeckStructuralChange | null {
    return this.structural;
  }

  /** The server took (or no longer needs) the structural change: the next
   * save works it out again from the order it now holds. */
  structuralLanded(): void {
    const change = this.structural;
    if (change && this.baseIds) {
      const base = new Set(this.baseIds);
      for (const id of change.deletes) base.delete(id);
      for (const insert of change.inserts) for (const id of insert.ids) base.add(id);
      this.rebased(base);
    }
    this.structural = null;
    this.structuralJson = null;
    this.changed();
  }

  /** The change was not sent as it stands (the server's order moved on):
   * the next save works it out again from the order it reads then. */
  structuralDiscarded(): void {
    this.structural = null;
    this.structuralJson = null;
    this.changed();
  }

  /** What the server now says each part's revision is (an insert's or a
   * delete's answer, the room's `deckSlidesChanged`): a held part whose
   * revision moved, with no change of this editor's on its way, is read
   * again. Returns the parts to read. A room that is being read afresh or
   * holds words not saved brings the part itself: closing it now would take
   * them with it. */
  applyPartRevisions(revisions: Readonly<Record<string, unknown>>): number[] {
    const stale: number[] = [];
    for (const [key, value] of Object.entries(revisions)) {
      const index = Number(key);
      const held = this.parts.get(index);
      if (!held || typeof value !== "number" || value === held.revision) continue;
      if (this.drafts.has(index)) continue;
      if (roomKeepsWords(held)) {
        held.revision = value;
        continue;
      }
      for (const [id, slide] of held.saved) this.stale.set(id, slide);
      this.parts.delete(index);
      this.rereading.add(index);
      this.statuses.set(index, "unknown");
      stale.push(index);
    }
    if (stale.length > 0) this.changed();
    return stale;
  }

  /** Each held part whose slides differ from what it last saved: a draft. */
  private queuePartDrafts(
    content: Rec,
    partRecords: ReadonlyMap<number, Array<{ id: string; record: Rec }>>,
    live: readonly S[] | undefined,
  ): void {
    const liveById = live ? new Map(live.map((slide) => [slide.id, slide])) : null;
    let touched = false;
    this.touchedRooms.clear();
    for (const [index, held] of this.parts) {
      if (!held.editable) continue;
      const records = partRecords.get(index) ?? [];
      const liveOf = new Map<string, S | undefined>();
      let changed = false;
      // A part with a live room: a slide that is another object than the one
      // the room last took is published, and the room itself finds a step
      // that changes nothing (DocumentPartRoom.publish compares what it holds
      // with what it is given) -- an Undo back to it writes no step and saves
      // nothing. The slide is not written out here to be compared first: the
      
      // slide, at a quarter CPU speed: 18 ms each, 276 ms over 41 keys).
      const byIdentity = held.room?.accepting === true;
      for (const { id, record } of records) {
        const now = liveById?.get(id);
        liveOf.set(id, now);
        const saved = held.saved.get(id);
        if (!saved) {
          changed = true;
        } else if (byIdentity) {
          if (!now || now !== saved) changed = true;
        } else if (now) {
          if (now === saved) {
            if (!held.savedStore.has(id)) held.savedStore.set(id, slideRecordJson(record));
          } else if (!unchangedFromSaved(now, saved)) {
            changed = true;
          }
        } else {
          const known = held.savedStore.get(id);
          if (known !== undefined && known !== slideRecordJson(record)) changed = true;
        }
      }
      if (held.room?.accepting) {
        // A part with a live room: what changed goes into the room, as this
        // member's step there; the room saves it into the part.
        if (changed && this.publishToRoom(held, content, records, liveOf)) this.touchedRooms.add(index);
        if (this.drafts.delete(index)) touched = true;
        if (this.pageId) writeKeptEdit(this.pageId, index, null);
        continue;
      }
      const existing = this.drafts.get(index);
      if (!changed) {
        // Back to what the server holds (an Undo): nothing to send.
        if (existing) {
          this.drafts.delete(index);
          if (this.pageId) writeKeptEdit(this.pageId, index, null);
          touched = true;
        }
        if (this.refused.delete(index)) touched = true;
        continue;
      }
      const partContent = deckPartContentWithSlides(content, records.map(({ record }) => record));
      const json = JSON.stringify(partContent);
      if (existing?.json === json && existing.revision === held.revision) continue;
      if (this.refused.get(index)?.json === json) continue;
      this.refused.delete(index);
      const draft: DeckPartDraft<S> = {
        index,
        revision: held.revision,
        ids: records.map(({ id }) => id),
        content: partContent,
        json,
        live: liveOf,
        store: new Map(records.map(({ id, record }) => [id, slideRecordJson(record)])),
      };
      this.drafts.set(index, draft);
      if (this.pageId) writeKeptEdit(this.pageId, index, { rev: held.revision, content: partContent, sending: false });
      touched = true;
    }
    if (touched) this.changed();
  }

  /** The parts whose rooms the last `prepareBody` wrote into. */
  private readonly touchedRooms = new Set<number>();

  roomsTouched(): ReadonlySet<number> {
    return this.touchedRooms;
  }

  /**
   * The slides of `held` that changed, written over its room's content: each
   * in its place there (a slide's visibility stays the room's: the order
   * holds it), with the pictures and receipts it names. The rest of the room
   * is left exactly as it is, so a step is the change and nothing else.
   */
  private publishToRoom(
    held: HeldPart<S>,
    content: Rec,
    records: ReadonlyArray<{ id: string; record: Rec }>,
    liveOf: ReadonlyMap<string, S | undefined>,
  ): boolean {
    const room = held.room!;
    const base = room.content();
    const baseSlides = Array.isArray(base.slides) ? (base.slides as unknown[]) : [];
    const changedRecords = new Map<string, Rec>();
    for (const { id, record } of records) {
      const now = liveOf.get(id);
      const saved = held.saved.get(id);
      // Another object than the room last took: the room compares (above).
      if (!saved || !now || now !== saved) changedRecords.set(id, record);
    }
    if (changedRecords.size === 0) return false;
    const nextSlides = baseSlides.map((slide) => {
      const id = asRec(slide)?.id;
      const record = typeof id === "string" ? changedRecords.get(id) : undefined;
      if (!record) return slide;
      const { hidden: _hidden, ...stated } = record;
      void _hidden;
      return asRec(slide)?.hidden === true ? { ...stated, hidden: true } : stated;
    });
    const named = namedTableEntries([...changedRecords.values()], contentTables(content));
    const next: Rec = { ...base, slides: nextSlides };
    if (named.media) next[MEDIA_KEY] = { ...(asRec(base[MEDIA_KEY]) ?? {}), ...named.media };
    if (named.receipts) next[RECEIPTS_KEY] = { ...(asRec(base[RECEIPTS_KEY]) ?? {}), ...named.receipts };
    const touched = room.publish(next);
    // The room holds these now: they are what the next change is measured
    // against, and what a delivery of the body shows.
    for (const id of changedRecords.keys()) {
      const now = liveOf.get(id);
      if (!now) continue;
      held.saved.set(id, now);
      held.byId.set(id, now);
    }
    return touched;
  }

  /** The part's room is about to be read: what is typed now goes into it
   * then, rather than round it by the part's save, which the room would meet
   * as a write from elsewhere. */
  roomIsReading(index: number): boolean {
    return this.parts.get(index)?.room?.reading === true;
  }

  /** Attach (or detach: null) the live room of held part `index`. */
  /** No part rooms in this editor: the installed readings kept for their
   * first read are let go. */
  dropInstalledReads(): void {
    for (const held of this.parts.values()) held.installedRead = undefined;
  }

  attachRoom(index: number, room: DeckPartRoomPort | null): void {
    const held = this.parts.get(index);
    if (!held) return;
    if (room) held.room = room;
    else delete held.room;
  }

  /** The held parts a room may be opened for: index, room key, editable. */
  heldParts(): Array<{ index: number; partId: string; editable: boolean }> {
    return [...this.parts.values()].map((held) => ({ index: held.index, partId: held.partId, editable: held.editable }));
  }

  /**
   * A part's room says the part is now `content` (another member's input, a
   * room read afresh, an Undo of this member's): the part holds it, and its
   * slides (as the editor shows them, the order's visibility on each) are
   * returned for the editor to put in place. Null when it is not that part.
   */
  applyRoomContent(
    index: number,
    content: Record<string, unknown>,
    /** The room read afresh (a reseed), with the slides this window shows: no
     * text change says what moved while it was away, so they are compared. */
    shownBefore?: readonly S[],
    /** The slide as this window holds it now, by id: a slide the room gives
     * exactly as the window holds it stays the window's object (below). */
    currentOf?: (id: string) => S | undefined,
  ): ReadonlyMap<string, S> | null {
    const held = this.parts.get(index);
    if (!held || !Array.isArray(content.slides)) return null;
    const input = this.readPartInput(content);
    const installed = held.installedRead;
    held.installedRead = undefined;
    const slides = this.withDeckCommentAuthority(
      installed && installed.key === JSON.stringify(input) ? installed.slides : this.normalize(input),
    );
    // Typed while the room could not take it (before its first read, or read
    // afresh after a reseed): those slides wait in the part's draft, unknown
    // to the room. Its content is not put over them, or the words are gone:
    // they go into the room now it can take them, and the draft with them.
    const draft = this.drafts.get(index);
    const typed = new Map<string, { now: S; saved: S }>();
    for (const [id, now] of draft?.live ?? []) {
      const saved = held.saved.get(id);
      if (now && saved && now !== saved && !unchangedFromSaved(now, saved)) typed.set(id, { now, saved });
    }
    if (shownBefore) this.noteWordsMoved(shownBefore, slides.filter((slide) => !typed.has(slide.id)));
    const shown = new Map<string, S>();
    // A slide the room gives exactly as the window holds it (its whole
    // normalized record, written out afresh on both sides: an object is never
    // taken for its content by a cache) stays the window's object, as the
    // part's saved slide too, and is not handed back: a room's first read, or
    // another member's change to one slide, made every slide of the part a
    // new object, held as saved and drawn again, and each was then compared
    
    
    // field the window holds and the room does not included -- is handed back
    // as before.
    const kept = new Map<string, S>();
    for (const slide of slides) {
      if (typed.has(slide.id)) continue;
      const current = currentOf?.(slide.id);
      if (!current) continue;
      const view = this.shown(slide);
      if (view === current || JSON.stringify(view) === JSON.stringify(current)) kept.set(slide.id, current);
    }
    held.byId = new Map(slides.map((slide) => [slide.id, typed.get(slide.id)?.now ?? kept.get(slide.id) ?? slide]));
    held.saved.clear();
    held.savedStore.clear();
    for (const slide of slides) {
      const mine = typed.get(slide.id);
      if (mine) {
        held.saved.set(slide.id, mine.saved);
        continue;
      }
      const same = kept.get(slide.id);
      if (same) {
        held.saved.set(slide.id, same);
        continue;
      }
      const view = this.shown(slide);
      held.saved.set(slide.id, view);
      shown.set(slide.id, view);
    }
    if (draft && typed.size > 0 && held.room?.accepting) {
      const records = (Array.isArray(draft.content.slides) ? (draft.content.slides as Rec[]) : [])
        .map((record) => ({ id: String(record.id), record }))
        .filter(({ id }) => typed.has(id));
      this.publishToRoom(held, draft.content as Rec, records, draft.live);
      this.drafts.delete(index);
      if (this.pageId) writeKeptEdit(this.pageId, index, null);
      this.changed();
    }
    return shown;
  }

  /** The next part change the server can take now: its slides are exactly
   * the ones the server's order puts in its part (a deletion or a move into
   * the body lands with the body first). */
  nextDraft(): DeckPartDraft<S> | null {
    const marker = this.marker;
    if (!marker) return null;
    for (const draft of [...this.drafts.values()].sort((a, b) => a.index - b.index)) {
      const held = this.parts.get(draft.index);
      if (!held || !held.editable || this.status(draft.index) !== "ready") continue;
      const ordered = marker.order.filter((entry) => entry.part === draft.index).map((entry) => entry.id);
      if (!sameIds(ordered, draft.ids)) continue;
      return draft.revision === held.revision ? draft : { ...draft, revision: held.revision };
    }
    return null;
  }

  hasDrafts(): boolean {
    return this.drafts.size > 0;
  }

  /** The change of part `draft.index` is on its way. */
  sending(draft: DeckPartDraft<S>): void {
    if (this.pageId) writeKeptEdit(this.pageId, draft.index, { rev: draft.revision, content: draft.content, sending: true });
  }

  /** The server stored `draft` as revision `revision`. */
  landed(draft: DeckPartDraft<S>, revision: number): void {
    const held = this.parts.get(draft.index);
    if (!held) return;
    held.revision = revision;
    for (const id of draft.ids) {
      const now = draft.live.get(id);
      if (now) {
        held.saved.set(id, now);
        held.byId.set(id, now);
      }
      const store = draft.store.get(id);
      if (store !== undefined) held.savedStore.set(id, store);
      else held.savedStore.delete(id);
    }
    for (const id of [...held.saved.keys()]) {
      if (!draft.ids.includes(id)) {
        held.saved.delete(id);
        held.byId.delete(id);
        held.savedStore.delete(id);
      }
    }
    const current = this.drafts.get(draft.index);
    if (current && current.json === draft.json) {
      this.drafts.delete(draft.index);
      if (this.pageId) writeKeptEdit(this.pageId, draft.index, null);
    } else if (current && this.pageId) {
      // A newer change of the part, kept against what is stored now.
      writeKeptEdit(this.pageId, draft.index, { rev: revision, content: current.content, sending: false });
    }
    this.changed();
  }

  /** The server refused `draft` as content: kept, not sent again until the
   * part changes. */
  refuse(draft: DeckPartDraft<S>, message: string): void {
    if (this.drafts.get(draft.index)?.json === draft.json) this.drafts.delete(draft.index);
    this.refused.set(draft.index, { json: draft.json, message });
    this.changed();
  }

  /**
   * Part `index` was saved elsewhere first (`fresh`): the server's slides
   * become the part's, and this editor's are merged with them slide by
   * slide -- a slide only it changed stays as it has it (and is saved again
   * on the new revision), one only the server changed takes the server's,
   * one both changed is kept twice. Null when `fresh` is not that part.
   */
  mergeConflict(index: number, fresh: DocumentPartResponse, live: readonly S[]): DeckPartMerge<S> | null {
    const marker = this.marker;
    const held = this.parts.get(index);
    if (!marker || !held || !readyDeckPartSlides(fresh, index, marker.count)) return null;
    const remote = this.readPart(asRec(fresh.content)!);
    const liveById = new Map(live.map((slide) => [slide.id, slide]));
    // What the editor and the server last agreed on, before it is replaced.
    const base = new Map(held.saved);
    const replace = new Map<string, S>();
    const copiesAfter = new Map<string, S>();
    held.saved.clear();
    held.savedStore.clear();
    held.byId = new Map(remote.map((slide) => [slide.id, slide]));
    held.revision = typeof fresh.revision === "number" ? fresh.revision : held.revision;
    held.contentSha256 = typeof fresh.contentSha256 === "string" ? fresh.contentSha256 : held.contentSha256;
    held.editable = fresh.editable === true;
    for (const slide of remote) {
      const theirs = this.shown(slide);
      held.saved.set(slide.id, theirs);
      const mine = liveById.get(slide.id);
      if (!mine) continue;
      const was = base.get(slide.id);
      if (!was) {
        // Not in the part when this editor read it: a slide the server put
        // there since (an insert of this editor's, as it was sent).
        if (!sameSlideButVisibility(mine, theirs)) replace.set(slide.id, theirs);
      } else if (sameSlideButVisibility(mine, was)) {
        if (!sameSlideButVisibility(mine, theirs)) replace.set(slide.id, theirs);
      } else if (was && sameSlideButVisibility(theirs, was)) {
        // Only this editor changed it: it stays, and is saved again.
      } else if (!sameSlideButVisibility(mine, theirs)) {
        replace.set(slide.id, theirs);
        copiesAfter.set(slide.id, mine);
      }
    }
    this.drafts.delete(index);
    this.refused.delete(index);
    if (this.pageId) writeKeptEdit(this.pageId, index, null);
    this.changed();
    return { replace, copiesAfter };
  }

  /** The revision of held part `index`, or null. */
  heldRevision(index: number): number | null {
    return this.parts.get(index)?.revision ?? null;
  }

  /**
   * The member chose to have part `index` read again from the file over
   * their edits (re-import): their change of it not saved yet goes, and
   * every slide of it the editor shows (`live`) takes the reading that
   * comes; the part is pending until then.
   */
  beginReimport(index: number, live: readonly S[]): void {
    const held = this.parts.get(index);
    for (const slide of live) {
      if (this.ownerOf.get(slide.id) === index) this.stale.set(slide.id, slide);
    }
    if (held) for (const [id, slide] of held.saved) if (!this.stale.has(id)) this.stale.set(id, slide);
    this.parts.delete(index);
    this.rereading.delete(index);
    this.drafts.delete(index);
    this.refused.delete(index);
    if (this.pageId) writeKeptEdit(this.pageId, index, null);
    this.statuses.set(index, "pending");
    this.changed();
  }

  /** Whether the server's slides of part `index` are the ones of `sha`. */
  heldSha(index: number): string | null {
    return this.parts.get(index)?.contentSha256 ?? null;
  }

  /** Drop a held part whose stored content is no longer what it holds (the
   * server says it now has `sha`), unless it has a change of this editor's
   * not saved yet: it is read again. */
  forgetStalePart(index: number, sha: string): boolean {
    const held = this.parts.get(index);
    if (!held || held.contentSha256 === sha || this.drafts.has(index) || roomKeepsWords(held)) return false;
    for (const [id, slide] of held.saved) this.stale.set(id, slide);
    this.parts.delete(index);
    this.rereading.add(index);
    this.statuses.set(index, "unknown");
    this.changed();
    return true;
  }

  /** Parts let go only to be read again: their slides stay in the editor. */
  private readonly rereading = new Set<number>();

  /** Whether the editor shows, of a part just read, a slide changed since the
   * part was let go: the reading did not replace it, and it is still to go
   * to the server. */
  typedOverArrival(index: number, live: readonly S[]): boolean {
    const held = this.parts.get(index);
    if (!held) return false;
    return live.some((slide) => {
      const saved = held.saved.get(slide.id);
      return saved !== undefined && slide !== saved && !unchangedFromSaved(slide, saved);
    });
  }

  /** Slides of a part the server rewrote (`forgetStalePart`), as the editor
   * showed them: one still shown so takes the new reading when it comes. */
  private readonly stale = new Map<string, S>();

  /** `slides` with each arrived slide in place of its stub, or of the
   * earlier reading of a rewritten part nobody changed since (the slide
   * keeps the visibility the editor gave it). Same array when none changed. */
  withArrivedSlides(slides: readonly S[], arrived: ReadonlyMap<string, S>): S[] {
    let out: S[] | null = null;
    slides.forEach((slide, index) => {
      const held = arrived.get(slide.id);
      if (!held) return;
      const earlier = this.stale.get(slide.id);
      if (slide.importPending !== true && !(earlier && sameSlideButVisibility(earlier, slide))) return;
      out ??= slides.slice();
      out[index] = this.withHidden(held, slide.hidden === true);
    });
    for (const id of arrived.keys()) this.stale.delete(id);
    return out ?? (slides as S[]);
  }
}

/** The deck-parts codes a body save may be refused with (`refuse_deck_body_save`). */
export const DECK_BODY_REFUSAL_CODES: readonly string[] = [
  "document_part_blocks_in_body",
  "deck_part_slide_moved",
  "deck_part_order_ahead",
  "deck_parts_order_mismatch",
];

/** How long a part refused or not found waits before it is asked again. */
export const DECK_PART_POLL_MS = 5_000;
/** How often the parts the import is still writing are asked after, when the
 * room has not said they are written (its notices are best effort: one sent
 * before this editor's socket joined is never seen). An import writes them
 * within seconds; one GET of a part not yet written is a few ms. */
export const DECK_PART_PENDING_POLL_MS = 2_000;
/** ...and while the connection to the server is down. */
export const DECK_PART_POLL_OFFLINE_MS = 2_000;
/** A part whose read or save failed (the network, not its answer) is tried again after. */
export const DECK_PART_RETRY_MS = 3_000;
/** A later part's change waits this long for the next keystroke before it is sent. */
export const DECK_PART_SAVE_DEBOUNCE_MS = 700;
/** How many slides on each side of the current one load with it. */
export const DECK_PART_NEIGHBOUR_SLIDES = 2;
/** A deck of up to this many parts loads every ready part once the ones
 * near the reader are in (the owner's target: every slide editable in 10 s);
 * a longer one loads the parts within `DECK_PART_PREFETCH_RADIUS` of the
 * current slide's. */
export const DECK_PART_PREFETCH_ALL_MAX_PARTS = 48;
export const DECK_PART_PREFETCH_RADIUS = 3;

/**
 * The later parts to load next, nearest first: the current slide's part and
 * its neighbours' (`neighbours` slides each side), then the parts of slides
 * the rail shows (`visibleIds`), then -- background -- every other part (a
 * short deck) or those within `DECK_PART_PREFETCH_RADIUS` parts of the
 * current one. Parts held, or known not to be ready, are left out.
 */
export function deckPartsToLoad(input: {
  readonly count: number;
  /** The editor's slide ids in show order. */
  readonly slideIds: readonly string[];
  readonly activeIndex: number;
  readonly visibleIds: Iterable<string>;
  readonly partOf: (id: string) => number | undefined;
  readonly loadable: (part: number) => boolean;
  readonly neighbours?: number;
}): { readonly urgent: number[]; readonly background: number[] } {
  const urgent: number[] = [];
  const seen = new Set<number>();
  const want = (id: string | undefined, list: number[]) => {
    if (!id) return;
    const part = input.partOf(id);
    if (part === undefined || part === 0 || seen.has(part) || !input.loadable(part)) return;
    seen.add(part);
    list.push(part);
  };
  const reach = input.neighbours ?? DECK_PART_NEIGHBOUR_SLIDES;
  want(input.slideIds[input.activeIndex], urgent);
  for (let step = 1; step <= reach; step += 1) {
    want(input.slideIds[input.activeIndex + step], urgent);
    want(input.slideIds[input.activeIndex - step], urgent);
  }
  for (const id of input.visibleIds) want(id, urgent);
  const background: number[] = [];
  const activePart = input.partOf(input.slideIds[input.activeIndex] ?? "") ?? 0;
  const all = input.count <= DECK_PART_PREFETCH_ALL_MAX_PARTS;
  const candidates: number[] = [];
  for (let part = 1; part < input.count; part += 1) {
    if (all || Math.abs(part - activePart) <= DECK_PART_PREFETCH_RADIUS) candidates.push(part);
  }
  candidates.sort((a, b) => Math.abs(a - activePart) - Math.abs(b - activePart) || a - b);
  for (const part of candidates) {
    if (seen.has(part) || !input.loadable(part)) continue;
    seen.add(part);
    background.push(part);
  }
  return { urgent, background };
}
