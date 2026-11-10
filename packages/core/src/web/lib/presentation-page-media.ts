/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  inlinePresentationDeckReceipts,
  sharePresentationDeckReceipts,
} from "./presentation-deck-receipts";
import {
  compactPictureObjectRecords,
  expandPictureObjectRecords,
  pictureObjectTokenForUrl,
  pictureObjectUrlForToken,
  pictureObjectUrlSha,
} from "./page-picture-objects";

export const PRESENTATION_MEDIA_KEY = "media";
export const PRESENTATION_MEDIA_TOKEN_PREFIX = "mrai-media:";
/** The in-memory memo: which stored picture each field's bytes came from. */
export const PRESENTATION_MEDIA_REFS_KEY = "mediaRefs";
const MEDIA_FIELDS = ["src", "graphicFallbackSrc"] as const;
type MediaField = (typeof MEDIA_FIELDS)[number];
const TOKEN_RE = /^mrai-media:([0-9a-f]{64})$/;

type Rec = Record<string, unknown>;

/**
 * Bytes this module expanded from a repeated-picture token, by value → digest.
 * Bounded: the oldest entries go first (a string that is no longer known is
 * simply sent as bytes, which the server stores once again).
 */
const REPEATED_MEDIA_LIMIT = 2_000;
const repeatedMediaDigests = new Map<string, string>();
/** Changes whenever `repeatedMediaDigests` does (see `compactedUnits`). */
let repeatedMediaGeneration = 0;

function rememberRepeatedMedia(value: string, sha: string): void {
  if (repeatedMediaDigests.has(value)) return;
  if (repeatedMediaDigests.size >= REPEATED_MEDIA_LIMIT) {
    const oldest = repeatedMediaDigests.keys().next().value;
    if (oldest !== undefined) repeatedMediaDigests.delete(oldest);
  }
  repeatedMediaDigests.set(value, sha);
  repeatedMediaGeneration += 1;
}










const SESSION_MEDIA_LIMIT = 2_000;
const sessionMediaBytes = new Map<string, string>();
/** How many tokens were put back from `sessionMediaBytes` (see below). */
let restoredFromSession = 0;

function rememberSessionMedia(sha: string, bytes: string): void {
  if (sessionMediaBytes.get(sha) === bytes) return;
  sessionMediaBytes.delete(sha);
  if (sessionMediaBytes.size >= SESSION_MEDIA_LIMIT) {
    const oldest = sessionMediaBytes.keys().next().value;
    if (oldest !== undefined) sessionMediaBytes.delete(oldest);
  }
  sessionMediaBytes.set(sha, bytes);
}

/** The bytes stored under `sha`: the page's own entry, else this session's. */
function mediaBytes(media: Rec, sha: string): string | undefined {
  const own = media[sha];
  if (typeof own === "string") {
    rememberSessionMedia(sha, own);
    return own;
  }
  const known = sessionMediaBytes.get(sha);
  if (known !== undefined) restoredFromSession += 1;
  return known;
}

/**
 * Changes whenever an expansion put back a picture whose entry the page had
 * lost: the editor then writes the page again, which stores the entry anew.
 */
export function presentationMediaRestoredCount(): number {
  return restoredFromSession;
}

const SHARED_SUBTREES = ["slides", "masters"] as const;

/**
 * `node` with `visit` applied to every string, copy-on-write. The `src` /
 * `graphicFallbackSrc` of an inherited image block are left to the memo path.
 */
function walkStrings(node: unknown, visit: (value: string) => string): unknown {
  if (typeof node === "string") return visit(node);
  if (Array.isArray(node)) {
    let copy: unknown[] | null = null;
    node.forEach((item, index) => {
      const next = walkStrings(item, visit);
      if (next !== item) {
        copy ??= node.slice();
        copy[index] = next;
      }
    });
    return copy ?? node;
  }
  const record = asRec(node);
  if (!record) return node;
  const inherited = presentationBlockMayShareMedia(record);
  let copy: Rec | null = null;
  for (const [key, item] of Object.entries(record)) {
    if (inherited && (MEDIA_FIELDS as readonly string[]).includes(key)) continue;
    const next = walkStrings(item, visit);
    if (next !== item) {
      copy ??= { ...record };
      copy[key] = next;
    }
  }
  return copy ?? record;
}

function rewriteDrawnStrings(content: Rec, visit: (value: string) => string): Rec {
  let out = content;
  for (const key of SHARED_SUBTREES) {
    if (!(key in content)) continue;
    const next = walkStrings(content[key], visit);
    if (next !== content[key]) {
      if (out === content) out = { ...content };
      out[key] = next;
    }
  }
  return out;
}

function asRec(value: unknown): Rec | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Rec)
    : null;
}

export function presentationMediaTokenSha(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const match = TOKEN_RE.exec(value);
  return match ? match[1] : null;
}

/** An image block the projection stamped as a layout/master graphic. */
export function presentationBlockMayShareMedia(block: unknown): boolean {
  const record = asRec(block);
  if (!record || record.type !== "image") return false;
  const origin = asRec(record.source)?.inheritedFrom;
  return origin === "layout" || origin === "master";
}

/** `content` with every block list passed through `rewrite` (shallow copies). */
function rebuilt(content: Rec, rewrite: (blocks: unknown[]) => unknown[]): Rec {
  const out: Rec = { ...content };
  if (Array.isArray(content.slides)) {
    out.slides = content.slides.map((slide) => {
      const record = asRec(slide);
      return record && Array.isArray(record.blocks)
        ? { ...record, blocks: rewrite(record.blocks) }
        : slide;
    });
  }
  if (Array.isArray(content.masters)) {
    out.masters = content.masters.map((master) => {
      const record = asRec(master);
      if (!record) return master;
      const next: Rec = { ...record };
      if (Array.isArray(record.blocks)) next.blocks = rewrite(record.blocks);
      if (Array.isArray(record.layouts)) {
        next.layouts = record.layouts.map((layout) => {
          const layoutRecord = asRec(layout);
          if (!layoutRecord) return layout;
          const nextLayout: Rec = { ...layoutRecord };
          for (const key of ["blocks", "masterGraphics"] as const) {
            if (Array.isArray(layoutRecord[key])) {
              nextLayout[key] = rewrite(layoutRecord[key] as unknown[]);
            }
          }
          return nextLayout;
        });
      }
      return next;
    });
  }
  return out;
}

function hasToken(content: Rec): boolean {
  let found = false;
  rebuilt(content, (blocks) => {
    if (!found) {
      found = blocks.some((block) => {
        const record = asRec(block);
        return (
          record?.type === "image" &&
          (MEDIA_FIELDS.some((field) => presentationMediaTokenSha(record[field])) ||
            pictureObjectUrlForToken(record.src) !== null)
        );
      });
    }
    return blocks;
  });
  return found;
}

/**
 * A picture stored beside the page (`page-picture-objects`) drawn from its
 * URL: the token becomes the URL the page's picture route answers, when a
 * page read so far holds it.
 */
function withPictureObjectUrl(block: unknown): unknown {
  const image = asRec(block);
  if (!image || image.type !== "image") return block;
  const url = pictureObjectUrlForToken(image.src);
  return url ? { ...image, src: url } : block;
}

/** A picture drawn from its URL saved as its token again. */
function withPictureObjectToken(block: unknown): unknown {
  const image = asRec(block);
  if (!image || image.type !== "image" || typeof image.src !== "string") return block;
  if (pictureObjectUrlSha(image.src) === null) return block;
  return { ...image, src: pictureObjectTokenForUrl(image.src) };
}









export function expandPresentationMedia<T>(content: T): T {
  // The receipts every slide repeats are stored once too
  // (`presentation-deck-receipts`); they go back on each slide after the
  // walks below, which then meet each of them once.
  return inlinePresentationDeckReceipts(expandMediaOnly(content));
}

function expandMediaOnly<T>(content: T): T {
  const original = asRec(content);
  if (!original) return content;
  // A background or picture fill stored beside the page is drawn from its URL.
  const record = rewriteStoredPictureRecords(original, expandPictureObjectRecords);
  const media = asRec(record[PRESENTATION_MEDIA_KEY]) ?? {};
  if (record[PRESENTATION_MEDIA_KEY] === undefined && !hasToken(record)) {
    return record as T;
  }
  const out = rebuilt(record, (blocks) =>
    blocks.map((stored) => {
      // A picture stored beside the page is drawn from its URL.
      const block = withPictureObjectUrl(stored);
      const image = asRec(block);
      // A slide's own picture is a repeated one: the pass below expands it.
      if (!image || !presentationBlockMayShareMedia(image)) return block;
      let next: Rec | null = null;
      const refs: Partial<Record<MediaField, string>> = {};
      for (const field of MEDIA_FIELDS) {
        const sha = presentationMediaTokenSha(image[field]);
        if (!sha) continue;
        const value = mediaBytes(media, sha);
        if (typeof value !== "string") continue;
        next ??= { ...image };
        next[field] = value;
        refs[field] = sha;
      }
      if (!next) return block;
      next[PRESENTATION_MEDIA_REFS_KEY] = refs;
      return next;
    }),
  );
  // Any other token is a picture the page repeats.
  const expanded = rewriteDrawnStrings(out, (value) => {
    const sha = presentationMediaTokenSha(value);
    const bytes = sha ? mediaBytes(media, sha) : undefined;
    if (!sha || typeof bytes !== "string") return value;
    rememberRepeatedMedia(bytes, sha);
    return bytes;
  });
  delete expanded[PRESENTATION_MEDIA_KEY];
  return expanded as T;
}

/**
 * The stored form: each remembered picture back in `media`, its blocks
 * holding the token. Only an inherited image block whose field still holds
 * bytes under a remembered digest is compacted; a second, different value
 * under the same digest stays inline. Every `mediaRefs` memo is dropped, so
 * nothing the server does not know leaves the editor.
 */
/** A block as it is stored: its media memo compacted, and a picture stored
 * beside the page named by its token instead of the URL it is drawn from. */
function compactBlockMediaRefs(block: unknown, media: Record<string, string>): unknown {
  return withPictureObjectToken(compactBlockMediaRefsOnly(block, media));
}

/** An image block with its `mediaRefs` memo dropped and each field still
 * holding a remembered picture's bytes written as its token (into `media`);
 * `block` itself when it carries no memo. */
function compactBlockMediaRefsOnly(block: unknown, media: Record<string, string>): unknown {
  const image = asRec(block);
  if (!image || image.type !== "image" || !(PRESENTATION_MEDIA_REFS_KEY in image)) {
    return block;
  }
  const refs = asRec(image[PRESENTATION_MEDIA_REFS_KEY]) ?? {};
  const next: Rec = { ...image };
  delete next[PRESENTATION_MEDIA_REFS_KEY];
  if (!presentationBlockMayShareMedia(image)) return next;
  for (const field of MEDIA_FIELDS) {
    const sha = refs[field];
    const value = image[field];
    if (
      typeof sha !== "string" ||
      !/^[0-9a-f]{64}$/.test(sha) ||
      typeof value !== "string" ||
      !value.startsWith("data:")
    ) {
      continue;
    }
    if (media[sha] !== undefined && media[sha] !== value) continue;
    media[sha] = value;
    next[field] = `${PRESENTATION_MEDIA_TOKEN_PREFIX}${sha}`;
  }
  return next;
}

/** A repeated picture's bytes as its token while they are still those bytes. */
function repeatedMediaToken(value: string, media: Record<string, string>): string {
  if (!value.startsWith("data:")) return value;
  const sha = repeatedMediaDigests.get(value);
  if (!sha || (media[sha] !== undefined && media[sha] !== value)) {
    return value;
  }
  media[sha] = value;
  return `${PRESENTATION_MEDIA_TOKEN_PREFIX}${sha}`;
}

/** `content` with its slides' and masters' picture records rewritten. */
function rewriteStoredPictureRecords(content: Rec, rewrite: <N>(node: N) => N): Rec {
  let out = content;
  for (const key of SHARED_SUBTREES) {
    if (!(key in content)) continue;
    const next = rewrite(content[key]);
    if (next !== content[key]) {
      if (out === content) out = { ...content };
      out[key] = next;
    }
  }
  return out;
}

/** The whole page compacted in one pass (see `compactPresentationMedia`). */
function compactPresentationMediaWhole<T>(content: T, record: Rec): T {
  const media: Record<string, string> = {};
  let touched = false;
  const out = rebuilt(rewriteStoredPictureRecords(record, compactPictureObjectRecords), (blocks) =>
    blocks.map((block) => {
      const next = compactBlockMediaRefs(block, media);
      if (next !== block) touched = true;
      return next;
    }),
  );
  // A repeated picture goes back as its token while its bytes are unchanged.
  const shared =
    repeatedMediaDigests.size > 0
      ? rewriteDrawnStrings(out, (value) => repeatedMediaToken(value, media))
      : out;
  if (!touched && shared === out && out === record) return content;
  if (Object.keys(media).length > 0) {
    shared[PRESENTATION_MEDIA_KEY] = media;
  } else {
    delete shared[PRESENTATION_MEDIA_KEY];
  }
  return shared as T;
}

/** One slide or master compacted, and the media entries it wrote. */
interface CompactedUnit {
  readonly generation: number;
  readonly out: unknown;
  readonly media: readonly (readonly [string, string])[];
}

/**
 * Each slide and master as last compacted, by the object itself: a keystroke
 * re-serializes the deck, and every slide it did not touch is the same object
 * as before, so it is not walked again. Valid while the remembered digests
 * are those it was compacted under.
 */
const compactedUnits = new WeakMap<object, CompactedUnit>();

function compactedBlocks(blocks: unknown[], media: Record<string, string>): unknown[] {
  let copy: unknown[] | null = null;
  blocks.forEach((block, index) => {
    const next = compactBlockMediaRefs(block, media);
    if (next !== block) {
      copy ??= blocks.slice();
      copy[index] = next;
    }
  });
  return copy ?? blocks;
}

function compactUnit(unit: Rec, master: boolean): CompactedUnit {
  const cached = compactedUnits.get(unit);
  if (cached && cached.generation === repeatedMediaGeneration) return cached;
  const media: Record<string, string> = {};
  // A background or picture fill drawn from a stored picture's URL is saved
  // as its token again.
  const source: Rec = compactPictureObjectRecords(unit);
  let out: Rec = source;
  const set = (key: string, value: unknown) => {
    if (value === out[key]) return;
    if (out === unit) out = { ...unit };
    out[key] = value;
  };
  if (Array.isArray(source.blocks)) set("blocks", compactedBlocks(source.blocks, media));
  if (master && Array.isArray(source.layouts)) {
    let layouts: unknown[] | null = null;
    (source.layouts as unknown[]).forEach((layout, index) => {
      const layoutRecord = asRec(layout);
      if (!layoutRecord) return;
      let nextLayout: Rec = layoutRecord;
      for (const key of ["blocks", "masterGraphics"] as const) {
        if (!Array.isArray(layoutRecord[key])) continue;
        const next = compactedBlocks(layoutRecord[key] as unknown[], media);
        if (next === layoutRecord[key]) continue;
        if (nextLayout === layoutRecord) nextLayout = { ...layoutRecord };
        nextLayout[key] = next;
      }
      if (nextLayout !== layoutRecord) {
        layouts ??= (source.layouts as unknown[]).slice();
        layouts[index] = nextLayout;
      }
    });
    if (layouts) set("layouts", layouts);
  }
  const drawn = repeatedMediaDigests.size > 0
    ? walkStrings(out, (value) => repeatedMediaToken(value, media))
    : out;
  const compacted: CompactedUnit = {
    generation: repeatedMediaGeneration,
    out: drawn,
    media: Object.entries(media),
  };
  compactedUnits.set(unit, compacted);
  return compacted;
}

export function compactPresentationMedia<T>(content: T): T {
  // Each receipt every slide repeats is named by its key first
  // (`presentation-deck-receipts`), so the walks below meet it once.
  const record = asRec(sharePresentationDeckReceipts(content));
  if (!record || !Array.isArray(record.slides)) return content;
  content = record as T;
  const media: Record<string, string> = {};
  let conflict = false;
  let changed = false;
  const compactList = (list: unknown[], master: boolean): unknown[] => {
    let copy: unknown[] | null = null;
    list.forEach((unit, index) => {
      const unitRecord = asRec(unit);
      const compacted: Pick<CompactedUnit, "out" | "media"> = unitRecord
        ? compactUnit(unitRecord, master)
        : (() => {
            const own: Record<string, string> = {};
            const out = repeatedMediaDigests.size > 0
              ? walkStrings(unit, (value) => repeatedMediaToken(value, own))
              : unit;
            return { out, media: Object.entries(own) };
          })();
      for (const [sha, value] of compacted.media) {
        if (media[sha] !== undefined && media[sha] !== value) conflict = true;
        media[sha] = value;
      }
      if (compacted.out !== unit) {
        copy ??= list.slice();
        copy[index] = compacted.out;
      }
    });
    if (copy) changed = true;
    return copy ?? list;
  };
  const slides = compactList(record.slides, false);
  const masters = Array.isArray(record.masters) ? compactList(record.masters, true) : null;
  // Two different pictures under one digest: which of them is written as the
  // token depends on the page's order, so the page is compacted as a whole.
  if (conflict) return compactPresentationMediaWhole(content, record);
  if (!changed) return content;
  const shared: Rec = { ...record, slides };
  if (masters) shared.masters = masters;
  if (Object.keys(media).length > 0) {
    shared[PRESENTATION_MEDIA_KEY] = media;
  } else {
    delete shared[PRESENTATION_MEDIA_KEY];
  }
  return shared as T;
}
