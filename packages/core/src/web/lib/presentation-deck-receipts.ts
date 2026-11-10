/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const PRESENTATION_DECK_RECEIPTS_KEY = "deckReceipts";
export const PRESENTATION_DECK_RECEIPT_TOKEN_PREFIX = "mrai-deck:";
const REF_KEY = "deckReceipt";
const SCOPED_SLIDE_SIZE_VERSION = "pptx-slide-size-cascade-v2";
const KEY_RE = /^mrai-deck:([0-9A-Za-z_-]{1,64})$/;
const REF_KEYS = [REF_KEY, "scopeSlideId", "slidePart", "slidePartIndex"];
/** The block fields a stored block may name (API `SHARED_BLOCK_FIELDS`). */
const SHARED_BLOCK_FIELDS: ReadonlyArray<readonly string[]> = [
  ["customPath"],
  ["source", "nativeShapeDrawing"],
  ["source", "nativeGroupAuthority"],
  ["source", "nativeGroupMove"],
];

type Rec = Record<string, unknown>;

function asRec(value: unknown): Rec | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Rec)
    : null;
}

export function presentationDeckReceiptTokenKey(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const match = KEY_RE.exec(value);
  return match ? match[1] : null;
}

/** A stored Slide Size reference (`{ deckReceipt, … }`). */
export function isPresentationDeckReceiptRef(value: unknown): boolean {
  const record = asRec(value);
  return Boolean(record && REF_KEY in record);
}






const SESSION_LIMIT = 8192;
const sessionValues = new Map<string, unknown>();
/** The key of each entry object, and of each receipt the loader joined. */
const keyOfObject = new WeakMap<object, string>();
/** The key of each known entry by its JSON, filled lazily. */
const keyOfJson = new Map<string, string>();
/** Entries not yet in `keyOfJson`. */
let unindexed: Array<[string, object]> = [];
/** Each receipt's JSON, computed once per object. */
const jsonOfObject = new WeakMap<object, string>();
/** Changes whenever the entries this session knows do (see `shareSlide`). */
let generation = 0;

function rememberEntry(key: string, value: unknown): void {
  if (!value || typeof value !== "object") return;
  if (sessionValues.get(key) !== value) {
    sessionValues.delete(key);
    if (sessionValues.size >= SESSION_LIMIT) {
      const oldest = sessionValues.keys().next().value;
      if (oldest !== undefined) sessionValues.delete(oldest);
    }
    sessionValues.set(key, value);
    unindexed.push([key, value as object]);
    generation += 1;
  }
  keyOfObject.set(value as object, key);
}

/** JSON with every object's keys sorted: equal values, whatever order the
 * editor built them in, read the same. */
function sortedJson(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map((item) => sortedJson(item ?? null)).join(",")}]`;
  if (value && typeof value === "object") {
    const record = value as Rec;
    return `{${Object.keys(record)
      .filter((key) => record[key] !== undefined)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${sortedJson(record[key])}`)
      .join(",")}}`;
  }
  return JSON.stringify(value) ?? "null";
}

function jsonOf(value: object): string {
  let json = jsonOfObject.get(value);
  if (json === undefined) {
    json = sortedJson(value);
    jsonOfObject.set(value, json);
  }
  return json;
}

function fieldOf(block: Rec, path: readonly string[]): unknown {
  let node: unknown = block;
  for (const name of path) node = asRec(node)?.[name];
  return node;
}

/** `block` with each path set, copy-on-write, key order kept. */
function withFields(block: Rec, values: ReadonlyArray<[readonly string[], unknown]>): Rec {
  const out: Rec = { ...block };
  for (const [path, value] of values) {
    if (path.length === 1) {
      out[path[0]] = value;
    } else {
      const parent = asRec(out[path[0]]);
      out[path[0]] = { ...(parent ?? {}), [path[1]]: value };
    }
  }
  return out;
}

/** The key of an entry this session knows equal to `value`, or `null`. */
function knownKey(value: object): string | null {
  const direct = keyOfObject.get(value);
  if (direct !== undefined && sessionValues.get(direct) === value) return direct;
  if (sessionValues.size === 0) return null;
  for (const [key, entry] of unindexed) {
    if (sessionValues.get(key) === entry) keyOfJson.set(jsonOf(entry), key);
  }
  unindexed = [];
  const key = keyOfJson.get(jsonOf(value));
  const entry = key === undefined ? undefined : sessionValues.get(key);
  if (key === undefined || !entry || jsonOf(entry as object) !== jsonOf(value)) return null;
  keyOfObject.set(value, key);
  return key;
}

/** Test seam: forget every remembered entry. */
export function resetPresentationDeckReceiptsForTests(): void {
  sessionValues.clear();
  keyOfJson.clear();
  unindexed = [];
  generation += 1;
}

interface SlideSizeSplit {
  template: Rec;
  ref: { scopeSlideId: string; slidePart: Rec; slidePartIndex: number };
}

/** `{ template, ref }` of a scoped Slide Size receipt, or `null`. */
function splitSlideSizeAuthority(receipt: unknown): SlideSizeSplit | null {
  const record = asRec(receipt);
  if (
    !record ||
    record.schemaVersion !== SCOPED_SLIDE_SIZE_VERSION ||
    typeof record.scopeSlideId !== "string" ||
    !Array.isArray(record.parts)
  ) {
    return null;
  }
  const parts = record.parts as unknown[];
  let at = -1;
  for (let index = 0; index < parts.length; index += 1) {
    if (asRec(parts[index])?.kind !== "slide") continue;
    if (at !== -1) return null;
    at = index;
  }
  const own = at === -1 ? null : asRec(parts[at]);
  if (!own || own.ownerId !== record.scopeSlideId) return null;
  const template: Rec = {};
  for (const [key, value] of Object.entries(record)) {
    template[key] = key === "scopeSlideId" ? null : value;
  }
  template.parts = [...parts.slice(0, at), ...parts.slice(at + 1)];
  return {
    template,
    ref: { scopeSlideId: record.scopeSlideId, slidePart: own, slidePartIndex: at },
  };
}

function joinSlideSizeAuthority(template: unknown, ref: Rec): Rec | null {
  const entry = asRec(template);
  const index = ref.slidePartIndex;
  if (
    !entry ||
    entry.schemaVersion !== SCOPED_SLIDE_SIZE_VERSION ||
    !("scopeSlideId" in entry) ||
    entry.scopeSlideId !== null ||
    !Array.isArray(entry.parts) ||
    Object.keys(ref).length !== REF_KEYS.length ||
    REF_KEYS.some((key) => !(key in ref)) ||
    typeof ref.scopeSlideId !== "string" ||
    !asRec(ref.slidePart) ||
    typeof index !== "number" ||
    !Number.isInteger(index) ||
    index < 0 ||
    index > (entry.parts as unknown[]).length
  ) {
    return null;
  }
  const parts = (entry.parts as unknown[]).slice();
  parts.splice(index, 0, ref.slidePart);
  return { ...entry, scopeSlideId: ref.scopeSlideId, parts };
}

/** Which entry (and which split) each joined Slide Size receipt came from. */
const joinedSlideSize = new WeakMap<object, { key: string; split: SlideSizeSplit }>();

/**
 * The form the editor works on: every receipt on its slide (each slide
 * naming one entry holds that entry itself), no `deckReceipts`. A name whose
 * entry the page lacks takes the one this session knows, else is left as it
 * is. Returns `content` itself when it holds neither a table nor a name.
 */
export function inlinePresentationDeckReceipts<T>(content: T): T {
  const record = asRec(content);
  if (!record || !Array.isArray(record.slides)) return content;
  const table = asRec(record[PRESENTATION_DECK_RECEIPTS_KEY]) ?? {};
  for (const [key, value] of Object.entries(table)) rememberEntry(key, value);
  const entry = (key: string | null): unknown =>
    key === null ? undefined : (table[key] ?? sessionValues.get(key));
  let slides: unknown[] | null = null;
  (record.slides as unknown[]).forEach((slide, index) => {
    const slideRecord = asRec(slide);
    if (!slideRecord) return;
    const source = asRec(slideRecord.source) ?? {};
    let nextSource: Rec | null = null;
    const cascade = asRec(source.layoutCascade);
    const layoutsKey = presentationDeckReceiptTokenKey(cascade?.layouts);
    const layouts = entry(layoutsKey);
    if (cascade && Array.isArray(layouts)) {
      nextSource = { ...source, layoutCascade: { ...cascade, layouts } };
    }
    const ref = asRec(source.slideSizeAuthority);
    if (ref && REF_KEY in ref) {
      const key = presentationDeckReceiptTokenKey(ref[REF_KEY]);
      const template = entry(key);
      const joined = joinSlideSizeAuthority(template, ref);
      if (joined && key !== null) {
        joinedSlideSize.set(joined, {
          key,
          split: {
            template: template as Rec,
            ref: {
              scopeSlideId: ref.scopeSlideId as string,
              slidePart: ref.slidePart as Rec,
              slidePartIndex: ref.slidePartIndex as number,
            },
          },
        });
        nextSource = { ...(nextSource ?? source), slideSizeAuthority: joined };
      }
    }
    let blocks: unknown[] | null = null;
    if (Array.isArray(slideRecord.blocks)) {
      (slideRecord.blocks as unknown[]).forEach((block, blockIndex) => {
        const blockRecord = asRec(block);
        if (!blockRecord) return;
        const values: Array<[readonly string[], unknown]> = [];
        for (const path of SHARED_BLOCK_FIELDS) {
          const named = entry(presentationDeckReceiptTokenKey(fieldOf(blockRecord, path)));
          if (asRec(named)) values.push([path, named]);
        }
        if (values.length === 0) return;
        blocks ??= (slideRecord.blocks as unknown[]).slice();
        blocks[blockIndex] = withFields(blockRecord, values);
      });
    }
    if (!nextSource && !blocks) return;
    slides ??= (record.slides as unknown[]).slice();
    slides[index] = {
      ...slideRecord,
      ...(nextSource ? { source: nextSource } : {}),
      ...(blocks ? { blocks } : {}),
    };
  });
  if (!slides && !(PRESENTATION_DECK_RECEIPTS_KEY in record)) return content;
  const out: Rec = { ...record };
  delete out[PRESENTATION_DECK_RECEIPTS_KEY];
  if (slides) out.slides = slides;
  return out as T;
}

/** One slide as last stored, by the slide object itself (a keystroke
 * re-serializes the deck; every slide it did not touch is the same object). */
interface SharedSlide {
  readonly generation: number;
  readonly out: unknown;
  readonly entries: readonly string[];
}
const sharedSlides = new WeakMap<object, SharedSlide>();

const sharedForms = new WeakMap<object, { generation: number; out: Rec }>();

/** The stored form of `from` made in this generation, else `made`, kept. */
function sameShared(from: Rec, made: Rec): Rec {
  const cached = sharedForms.get(from);
  if (cached && cached.generation === generation) return cached.out;
  sharedForms.set(from, { generation, out: made });
  return made;
}

function shareSlide(slide: unknown): SharedSlide {
  const slideRecord = asRec(slide);
  if (!slideRecord) return { generation, out: slide, entries: [] };
  const source = asRec(slideRecord.source) ?? {};
  const cached = sharedSlides.get(slideRecord);
  if (cached && cached.generation === generation) return cached;
  const entries: string[] = [];
  let nextSource: Rec | null = null;
  const cascade = asRec(source.layoutCascade);
  const namedLayouts = presentationDeckReceiptTokenKey(cascade?.layouts);
  if (namedLayouts !== null) {
    // Already stored by name (a stored body serialized again).
    entries.push(namedLayouts);
  } else if (cascade && Array.isArray(cascade.layouts)) {
    const key = knownKey(cascade.layouts);
    if (key !== null) {
      entries.push(key);
      nextSource = {
        ...source,
        layoutCascade: { ...cascade, layouts: `${PRESENTATION_DECK_RECEIPT_TOKEN_PREFIX}${key}` },
      };
    }
  }
  const authority = asRec(source.slideSizeAuthority);
  const namedSlideSize = authority && REF_KEY in authority
    ? presentationDeckReceiptTokenKey(authority[REF_KEY])
    : null;
  if (namedSlideSize !== null) {
    entries.push(namedSlideSize);
  } else if (authority && !(REF_KEY in authority)) {
    const joined = joinedSlideSize.get(authority);
    let key: string | null = null;
    let split: SlideSizeSplit | null = null;
    if (joined && sessionValues.get(joined.key) === joined.split.template) {
      ({ key, split } = joined);
    } else {
      split = splitSlideSizeAuthority(authority);
      key = split ? knownKey(split.template) : null;
    }
    if (key !== null && split) {
      entries.push(key);
      nextSource = {
        ...(nextSource ?? source),
        slideSizeAuthority: {
          [REF_KEY]: `${PRESENTATION_DECK_RECEIPT_TOKEN_PREFIX}${key}`,
          ...split.ref,
        },
      };
    }
  }
  let blocks: unknown[] | null = null;
  if (Array.isArray(slideRecord.blocks)) {
    (slideRecord.blocks as unknown[]).forEach((block, blockIndex) => {
      const blockRecord = asRec(block);
      if (!blockRecord) return;
      const values: Array<[readonly string[], unknown]> = [];
      for (const path of SHARED_BLOCK_FIELDS) {
        const value = fieldOf(blockRecord, path);
        const named = presentationDeckReceiptTokenKey(value);
        if (named !== null) {
          entries.push(named);
          continue;
        }
        const record = asRec(value);
        const key = record ? knownKey(record) : null;
        if (key === null) continue;
        entries.push(key);
        values.push([path, `${PRESENTATION_DECK_RECEIPT_TOKEN_PREFIX}${key}`]);
      }
      if (values.length === 0) return;
      blocks ??= (slideRecord.blocks as unknown[]).slice();
      blocks[blockIndex] = withFields(blockRecord, values);
    });
  }
  // An edited slide is a new object, but its receipts and its other objects
  // are the ones it had: their stored forms are the ones made last time, the
  // same objects, so what compares the deck skips them at once.
  if (nextSource) nextSource = sameShared(source, nextSource);
  if (blocks) {
    const shared = blocks as unknown[];
    shared.forEach((block, index) => {
      const original = (slideRecord.blocks as unknown[])[index];
      if (block !== original && asRec(original)) shared[index] = sameShared(original as Rec, block as Rec);
    });
  }
  const shared: SharedSlide = {
    generation,
    out: nextSource || blocks
      ? {
          ...slideRecord,
          ...(nextSource ? { source: nextSource } : {}),
          ...(blocks ? { blocks } : {}),
        }
      : slide,
    entries,
  };
  sharedSlides.set(slideRecord, shared);
  return shared;
}

/**
 * The stored form: each receipt this session knows named by its key, its
 * entry in `deckReceipts`. A receipt it does not know stays written out.
 * Returns `content` itself when nothing is named.
 */






let lastTable: Rec | null = null;

function sameTableAsLast(table: Rec): Rec {
  const last = lastTable;
  if (last) {
    const keys = Object.keys(table);
    if (keys.length === Object.keys(last).length && keys.every((key) => Object.hasOwn(last, key) && last[key] === table[key])) {
      return last;
    }
  }
  lastTable = table;
  return table;
}

export function sharePresentationDeckReceipts<T>(content: T): T {
  const record = asRec(content);
  if (!record || !Array.isArray(record.slides)) return content;
  const own = asRec(record[PRESENTATION_DECK_RECEIPTS_KEY]) ?? {};
  for (const [key, value] of Object.entries(own)) rememberEntry(key, value);
  const table: Rec = {};
  let slides: unknown[] | null = null;
  (record.slides as unknown[]).forEach((slide, index) => {
    const shared = shareSlide(slide);
    for (const key of shared.entries) {
      const value = own[key] ?? sessionValues.get(key);
      if (value !== undefined) table[key] = value;
    }
    if (shared.out !== slide) {
      slides ??= (record.slides as unknown[]).slice();
      slides[index] = shared.out;
    }
  });
  if (!slides && !(PRESENTATION_DECK_RECEIPTS_KEY in record)) return content;
  const out: Rec = { ...record };
  if (slides) out.slides = slides;
  if (Object.keys(table).length > 0) out[PRESENTATION_DECK_RECEIPTS_KEY] = sameTableAsLast(table);
  else delete out[PRESENTATION_DECK_RECEIPTS_KEY];
  return out as T;
}
