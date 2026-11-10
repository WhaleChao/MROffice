/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
interface FaceReach {
  /** Code point ranges; null: every character. */
  ranges: Array<[number, number]> | null;
  /** Whether the ranges hold the space (U+0020). */
  space: boolean;
}

export interface FontLoadReach {
  /** The families the load finished faces of, lower-cased. */
  readonly families: ReadonlySet<string>;
  /** Whether it can change `text` drawn in `fontFamily` (a CSS family list,
   * or the families of a stack in order). */
  touches(fontFamily: string | readonly string[] | undefined, text: string): boolean;
  /** The same for a CSS font shorthand (`italic 700 15px "A", B`). */
  touchesFont(font: string, text: string): boolean;
}

/** Each range list read once: every load walks every registered face (some
 * 4,000 on a CJK workbook, one family's lists 21,000 ranges long). */
const rangesRead = new Map<string, FaceReach>();
const RANGES_KEPT = 20_000;

function unicodeRanges(value: unknown): FaceReach {
  if (typeof value !== "string" || !value.trim()) return { ranges: null, space: true };
  const known = rangesRead.get(value);
  if (known) return known;
  const read = readUnicodeRanges(value);
  if (rangesRead.size >= RANGES_KEPT) rangesRead.clear();
  rangesRead.set(value, read);
  return read;
}

function readUnicodeRanges(value: string): FaceReach {
  const out: Array<[number, number]> = [];
  let space = false;
  for (const raw of value.split(",")) {
    const match = /^U\+([0-9A-F?]{1,6})(?:-([0-9A-F]{1,6}))?$/.exec(raw.trim().toUpperCase());
    if (!match) return { ranges: null, space: true };
    const [, first, last] = match;
    const low = parseInt(first.replace(/\?/g, "0"), 16);
    const high = last ? parseInt(last, 16) : parseInt(first.replace(/\?/g, "F"), 16);
    if (low === 0 && high >= 0x10ffff) return { ranges: null, space: true };
    if (low <= 0x20 && high >= 0x20) space = true;
    out.push([low, high]);
  }
  return { ranges: out, space };
}

function familyName(raw: unknown): string {
  return String(raw ?? "").replace(/^["']|["']$/g, "").trim().toLowerCase();
}

/** Family lists read once: a load asks of every cached text's font. */
const familiesRead = new Map<string, string[]>();
const FAMILIES_KEPT = 4096;

/** The families of a CSS family list, lower-cased, in order. */
export function cssFontFamilies(value: string | readonly string[] | undefined): string[] {
  if (value === undefined) return [];
  if (typeof value === "string") {
    const known = familiesRead.get(value);
    if (known) return [...known];
    const read = readFamilies(value.split(","));
    if (familiesRead.size >= FAMILIES_KEPT) familiesRead.clear();
    familiesRead.set(value, read);
    return [...read];
  }
  return readFamilies(value);
}

function readFamilies(list: readonly string[]): string[] {
  const out: string[] = [];
  for (const part of list) {
    const family = familyName(part.trim());
    if (family && !out.includes(family)) out.push(family);
  }
  return out;
}

/** The family list of a CSS font shorthand: what follows its size. */
export function fontShorthandFamilies(font: string): string[] {
  const match = /\d(?:\.\d+)?(?:px|pt|em|rem|%)(?:\/\S+)?\s+(.+)$/.exec(font);
  return cssFontFamilies(match ? match[1] : font);
}

/** Characters every face before the fallbacks draws, once one is loaded. */
const STACK_DRAWS_BELOW = 0x100;

function drawsIn(text: string, ranges: ReadonlyArray<[number, number]>, from: number): boolean {
  for (const char of text) {
    const point = char.codePointAt(0)!;
    if (point < from) continue;
    for (const [low, high] of ranges) if (point >= low && point <= high) return true;
  }
  return false;
}

interface ReadyFamilies { ready: Set<string>; registered: Set<string> }

/** The families a line's first available font can be in now: one with a
 * loaded face holding the space, or one with no face registered here (an
 * installed font). One whose faces all failed or have not loaded is not. */
function readyFamilies(): ReadyFamilies {
  const ready = new Set<string>();
  const registered = new Set<string>();
  const fonts = typeof document !== "undefined" ? (document.fonts as unknown as Iterable<FontFace> | undefined) : undefined;
  if (!fonts || typeof (fonts as { [Symbol.iterator]?: unknown })[Symbol.iterator] !== "function") return { ready, registered };
  for (const face of fonts) {
    const family = familyName(face.family);
    if (!family) continue;
    registered.add(family);
    if (face.status === "loaded" && unicodeRanges((face as { unicodeRange?: unknown }).unicodeRange).space) ready.add(family);
  }
  return { ready, registered };
}

/** One reading per event: each cache's listener asks of the same load. */
const reachOfEvent = new WeakMap<Event, FontLoadReach | null>();

/** What a `loadingdone` can change; null for one that names no faces (the
 * text engine's own: callers keep what they did for it). */
export function fontLoadReach(event: Event): FontLoadReach | null {
  if (reachOfEvent.has(event)) return reachOfEvent.get(event) ?? null;
  const reach = readFontLoadReach(event);
  reachOfEvent.set(event, reach);
  return reach;
}

function readFontLoadReach(event: Event): FontLoadReach | null {
  const faces = (event as Event & { fontfaces?: readonly FontFace[] }).fontfaces;
  if (!Array.isArray(faces) && !(faces && typeof (faces as { length?: unknown }).length === "number")) return null;
  const loaded = new Map<string, FaceReach>();
  for (const face of Array.from(faces as ArrayLike<FontFace>)) {
    const family = familyName(face.family);
    if (!family) continue;
    const reach = unicodeRanges((face as { unicodeRange?: unknown }).unicodeRange);
    const held = loaded.get(family);
    loaded.set(family, !held ? reach : {
      ranges: reach.ranges === null || held.ranges === null ? null : [...held.ranges, ...reach.ranges],
      space: reach.space || held.space,
    });
  }
  let ready: ReadyFamilies | null = null;
  const touchesStack = (families: readonly string[], text: string): boolean => {
    if (!families.some((family) => loaded.has(family))) return false;
    ready ??= readyFamilies();
    let covered = false;
    for (const family of families) {
      const reach = loaded.get(family);
      if (reach) {
        if (reach.ranges === null) return true;
        if (reach.space && !covered) return true;
        if (drawsIn(text, reach.ranges, covered ? STACK_DRAWS_BELOW : 0)) return true;
      }
      if (ready.ready.has(family) || !ready.registered.has(family)) covered = true;
    }
    return false;
  };
  return {
    families: new Set(loaded.keys()),
    touches: (fontFamily, text) => touchesStack(cssFontFamilies(fontFamily), text),
    touchesFont: (font, text) => touchesStack(fontShorthandFamilies(font), text),
  };
}
