/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The edits a visitor made on the homepage, placed on the page their file
 * became after sign-up.
 *
 * The homepage editor opened the file through the guest reader; the page the
 * signed-in import makes is the same file read by the same reader, but held
 * in its persisted form: pictures are references into the stored package
 * instead of inline data, fonts are page assets, blocks carry server ids.
 * Saving the guest's JSON over it would throw all of that away, so the guest's
 * edit is replayed instead, as a three-way merge:
 *
 *   result = realBase + (guestFinal - guestBase)
 *
 * - a subtree the visitor did not touch keeps the imported page's own bytes;
 * - list items pair by their id (an element `id`, or `attrs.block_id` /
 *   `attrs.id`), else by an LCS over their serialized form, a changed item
 *   pairing with the base item of the same type in its gap;
 * - a value the visitor wrote passes through the guest→page value map built
 *   by walking the two bases side by side (ids, picture sources, assets: the
 *   only places the two readings are expected to differ).
 *
 * Fail closed: an edit landing where the two bases differ in some other way,
 * or a guest-only inline picture left in what the visitor wrote, refuses the
 * whole rebase. The caller then keeps the imported page untouched and says so.
 */

type Json = null | boolean | number | string | Json[] | { [key: string]: Json };
type JsonObject = { [key: string]: Json };

export type ClassicRebaseResult =
  | { readonly ok: true; readonly changed: boolean; readonly content: Record<string, unknown> }
  | { readonly ok: false; readonly reason: ClassicRebaseRefusal; readonly at: string };

export type ClassicRebaseRefusal =
  | "invalid_input"
  | "edit_on_divergent_content"
  | "guest_inline_picture";

class Refusal extends Error {
  constructor(
    readonly reason: ClassicRebaseRefusal,
    readonly at: string,
  ) {
    super(reason);
  }
}

/** Largest LCS table one list gap may take before it pairs by position. */
const LCS_MAX_CELLS = 4_000_000;

function isObject(value: unknown): value is JsonObject {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

const serialized = new WeakMap<object, string>();

function keyOf(value: Json | undefined): string {
  if (value === undefined) return "\u0000undefined";
  if (value === null || typeof value !== "object") return JSON.stringify(value);
  const cached = serialized.get(value);
  if (cached !== undefined) return cached;
  const text = JSON.stringify(value);
  serialized.set(value, text);
  return text;
}

function same(a: Json | undefined, b: Json | undefined): boolean {
  if (a === b) return true;
  if (a === undefined || b === undefined) return false;
  if (typeof a !== "object" || typeof b !== "object" || a === null || b === null) return false;
  return keyOf(a) === keyOf(b);
}

/** Keys whose values are identities or storage references, not content. */
function isIdentityKey(key: string): boolean {
  return (
    key === "id" ||
    key === "src" ||
    key === "asset" ||
    key === "assetId" ||
    /(^|_)id$/i.test(key) ||
    /[a-z]Id$/.test(key) ||
    /(^|_)ids?$/i.test(key)
  );
}

function isInlineData(value: unknown): value is string {
  return typeof value === "string" && value.startsWith("data:");
}

function mapKey(key: string, value: Json): string {
  return `${key}\u0000${keyOf(value)}`;
}

function idOf(value: Json | undefined): string | null {
  if (!isObject(value)) return null;
  if (typeof value.id === "string" && value.id) return value.id;
  const attrs = value.attrs;
  if (isObject(attrs)) {
    if (typeof attrs.block_id === "string" && attrs.block_id) return attrs.block_id;
    if (typeof attrs.id === "string" && attrs.id) return attrs.id;
  }
  return null;
}

function typeOf(value: Json | undefined): string {
  if (isObject(value) && typeof value.type === "string") return value.type;
  if (Array.isArray(value)) return "[]";
  return value === null ? "null" : typeof value;
}

/** guest value → page value, for identity keys and inline pictures. */
class ValueMap {
  private readonly map = new Map<string, Json>();
  private readonly ambiguous = new Set<string>();
  /** Inline data the guest base held, by its text: it may be copied. */
  readonly guestInline = new Set<string>();

  record(key: string, guest: Json, real: Json): void {
    if (same(guest, real)) return;
    const k = mapKey(key, guest);
    if (this.ambiguous.has(k)) return;
    const known = this.map.get(k);
    if (known !== undefined && !same(known, real)) {
      this.map.delete(k);
      this.ambiguous.add(k);
      return;
    }
    this.map.set(k, real);
    if (isInlineData(guest)) this.map.set(mapKey("\u0000inline", guest), real);
  }

  lookup(key: string, guest: Json): Json | undefined {
    const k = mapKey(key, guest);
    if (this.map.has(k)) return this.map.get(k);
    // the same picture under another key (a copy keeps its source)
    if (isInlineData(guest)) return this.map.get(mapKey("\u0000inline", guest));
    return undefined;
  }
}

/** Walk the two bases side by side, learning how the guest reading names things. */
function learn(guest: Json | undefined, real: Json | undefined, key: string, map: ValueMap): void {
  if (guest === undefined || real === undefined) return;
  if (isInlineData(guest)) map.guestInline.add(guest);
  if (Array.isArray(guest)) {
    if (!Array.isArray(real)) return;
    if (guest.length === real.length) {
      guest.forEach((item, index) => learn(item, real[index], key, map));
      return;
    }
    // unequal lists still align by their ids where both carry them
    const realById = new Map<string, Json>();
    for (const item of real) {
      const id = idOf(item);
      if (id) realById.set(id, item);
    }
    for (const item of guest) {
      const id = idOf(item);
      const match = id ? realById.get(id) : undefined;
      if (match !== undefined) learn(item, match, key, map);
    }
    return;
  }
  if (isObject(guest)) {
    if (!isObject(real)) return;
    for (const [childKey, childValue] of Object.entries(guest)) {
      learn(childValue, real[childKey], childKey, map);
    }
    return;
  }
  if (isObject(real) || Array.isArray(real)) return;
  if (guest !== real && (isIdentityKey(key) || isInlineData(guest))) {
    map.record(key, guest, real);
  }
}

/** `real` reads as `base` once the guest's names are translated. */
function alignedEqual(real: Json | undefined, base: Json | undefined, key: string, map: ValueMap): boolean {
  if (same(real, base)) return true;
  if (real === undefined || base === undefined) return false;
  if (Array.isArray(base)) {
    return (
      Array.isArray(real) &&
      real.length === base.length &&
      base.every((item, index) => alignedEqual(real[index], item, key, map))
    );
  }
  if (isObject(base)) {
    if (!isObject(real)) return false;
    // keys only the page holds (stored references, import receipts) are its own
    return Object.entries(base).every(([childKey, childValue]) =>
      alignedEqual(real[childKey], childValue, childKey, map),
    );
  }
  if (isObject(real) || Array.isArray(real)) return false;
  const mapped = map.lookup(key, base);
  return mapped !== undefined && same(mapped, real);
}

/** What the visitor wrote, in the page's names. */
function translate(value: Json, key: string, map: ValueMap, path: string): Json {
  if (Array.isArray(value)) {
    return value.map((item, index) => translate(item, key, map, `${path}[${index}]`));
  }
  if (isObject(value)) {
    const out: JsonObject = {};
    for (const [childKey, childValue] of Object.entries(value)) {
      out[childKey] = translate(childValue, childKey, map, `${path}.${childKey}`);
    }
    return out;
  }
  const mapped = map.lookup(key, value);
  if (mapped !== undefined) return mapped;
  if (isInlineData(value) && map.guestInline.has(value)) {
    // a picture the guest reading inlined that the page names otherwise,
    // and the walk could not tell how: refuse rather than inline it
    throw new Refusal("guest_inline_picture", path);
  }
  return value;
}

function merge(
  real: Json | undefined,
  base: Json | undefined,
  final: Json | undefined,
  key: string,
  map: ValueMap,
  path: string,
): Json | undefined {
  if (same(base, final)) return real;
  // both readings already agree on the visitor's value (a stamped formula
  // result, a default both editors fill in): the page's own bytes stand
  // (a value, not a structure: a structure converges leaf by leaf below,
  // where a key the visitor removed is not mistaken for one only the page has)
  if (
    final !== undefined &&
    real !== undefined &&
    !isObject(final) &&
    !Array.isArray(final) &&
    alignedEqual(real, final, key, map)
  ) {
    return real;
  }
  if (final === undefined) {
    // removed by the visitor
    return undefined;
  }
  if (base === undefined) {
    // added by the visitor (the page may hold its own value: theirs wins,
    // as the save of their editor would write it)
    return translate(final, key, map, path);
  }
  if (Array.isArray(base) && Array.isArray(final) && Array.isArray(real)) {
    return mergeList(real, base, final, key, map, path);
  }
  if (isObject(base) && isObject(final) && isObject(real)) {
    return mergeObject(real, base, final, map, path);
  }
  if (alignedEqual(real, base, key, map)) return translate(final, key, map, path);
  throw new Refusal("edit_on_divergent_content", path);
}

function mergeObject(
  real: JsonObject,
  base: JsonObject,
  final: JsonObject,
  map: ValueMap,
  path: string,
): JsonObject {
  const out: JsonObject = { ...real };
  const keys = new Set([...Object.keys(base), ...Object.keys(final)]);
  for (const childKey of keys) {
    const inReal = Object.prototype.hasOwnProperty.call(real, childKey);
    const inBase = Object.prototype.hasOwnProperty.call(base, childKey);
    const childPath = `${path}.${childKey}`;
    if (inBase && !inReal) {
      // the guest reading has a key the page does not: only an untouched one is fine
      if (same(base[childKey], final[childKey])) continue;
      if (!Object.prototype.hasOwnProperty.call(final, childKey)) continue;
      throw new Refusal("edit_on_divergent_content", childPath);
    }
    const merged = merge(real[childKey], base[childKey], final[childKey], childKey, map, childPath);
    if (merged === undefined) delete out[childKey];
    else out[childKey] = merged;
  }
  return out;
}

/** Pairs of (base index, final index); every other index stands alone.
 *  Items that carry an id pair by it; the rest pair among themselves, in
 *  order, by an LCS over their serialized form and, in each gap, by type. */
function pairList(base: Json[], final: Json[]): Array<[number, number]> {
  const baseIds = base.map(idOf);
  const finalIds = final.map(idOf);
  const counts = new Map<string, number>();
  for (const id of baseIds) if (id) counts.set(id, (counts.get(id) ?? 0) + 1);
  const byId = new Map<string, number>();
  baseIds.forEach((id, index) => {
    if (id && counts.get(id) === 1) byId.set(id, index);
  });
  const pairs: Array<[number, number]> = [];
  const pairedBase = new Set<number>();
  const pairedFinal = new Set<number>();
  finalIds.forEach((id, index) => {
    const match = id ? byId.get(id) : undefined;
    if (match !== undefined && !pairedBase.has(match)) {
      pairedBase.add(match);
      pairedFinal.add(index);
      pairs.push([match, index]);
    }
  });
  // what is left (no id, an id seen twice, an id the other side lacks)
  const restBase = base.map((_, i) => i).filter((i) => !pairedBase.has(i) && !(baseIds[i] && byId.has(baseIds[i]!) && finalIds.includes(baseIds[i])));
  const restFinal = final.map((_, i) => i).filter((i) => !pairedFinal.has(i) && !(finalIds[i] && byId.has(finalIds[i]!)));
  for (const [b, f] of pairByOrder(restBase.map((i) => base[i]!), restFinal.map((i) => final[i]!))) {
    pairs.push([restBase[b]!, restFinal[f]!]);
  }
  return pairs;
}

/** Pairs of two id-less sequences, by LCS and then by type within gaps. */
function pairByOrder(base: Json[], final: Json[]): Array<[number, number]> {
  const bk = base.map(keyOf);
  const fk = final.map(keyOf);
  let start = 0;
  while (start < bk.length && start < fk.length && bk[start] === fk[start]) start += 1;
  let endB = bk.length;
  let endF = fk.length;
  while (endB > start && endF > start && bk[endB - 1] === fk[endF - 1]) {
    endB -= 1;
    endF -= 1;
  }
  const pairs: Array<[number, number]> = [];
  for (let i = 0; i < start; i += 1) pairs.push([i, i]);
  const anchors: Array<[number, number]> = [];
  const m = endB - start;
  const n = endF - start;
  if (m > 0 && n > 0 && m * n <= LCS_MAX_CELLS) {
    const width = n + 1;
    const table = new Uint32Array((m + 1) * width);
    for (let i = m - 1; i >= 0; i -= 1) {
      for (let j = n - 1; j >= 0; j -= 1) {
        table[i * width + j] =
          bk[start + i] === fk[start + j]
            ? table[(i + 1) * width + j + 1] + 1
            : Math.max(table[(i + 1) * width + j], table[i * width + j + 1]);
      }
    }
    let i = 0;
    let j = 0;
    while (i < m && j < n) {
      if (bk[start + i] === fk[start + j]) {
        anchors.push([start + i, start + j]);
        i += 1;
        j += 1;
      } else if (table[(i + 1) * width + j] >= table[i * width + j + 1]) {
        i += 1;
      } else {
        j += 1;
      }
    }
  }
  let prevB = start;
  let prevF = start;
  const gaps = [...anchors, [endB, endF] as [number, number]];
  for (const [ab, af] of gaps) {
    let bi = prevB;
    let fi = prevF;
    while (bi < ab && fi < af) {
      if (typeOf(base[bi]) === typeOf(final[fi])) {
        pairs.push([bi, fi]);
        bi += 1;
        fi += 1;
      } else if (ab - bi > af - fi) {
        bi += 1;
      } else {
        fi += 1;
      }
    }
    if (ab < endB || af < endF) pairs.push([ab, af]);
    prevB = ab + 1;
    prevF = af + 1;
  }
  for (let k = 0; endB + k < bk.length; k += 1) pairs.push([endB + k, endF + k]);
  return pairs;
}

function mergeList(
  real: Json[],
  base: Json[],
  final: Json[],
  key: string,
  map: ValueMap,
  path: string,
): Json[] {
  if (real.length !== base.length) {
    // the two readings hold different lists here; the visitor changed it
    throw new Refusal("edit_on_divergent_content", path);
  }
  const pairs = pairList(base, final);
  const baseOfFinal = new Map<number, number>();
  for (const [b, f] of pairs) baseOfFinal.set(f, b);
  const out: Json[] = [];
  final.forEach((item, index) => {
    const b = baseOfFinal.get(index);
    const itemPath = `${path}[${index}]`;
    if (b === undefined) {
      out.push(translate(item, key, map, itemPath));
      return;
    }
    const merged = merge(real[b], base[b], item, key, map, itemPath);
    if (merged !== undefined) out.push(merged);
  });
  return out;
}

export function classicEditsChanged(
  guestBase: Record<string, unknown>,
  guestFinal: Record<string, unknown>,
): boolean {
  return !same(guestBase as Json, guestFinal as Json);
}

/** Root attributes an editor restamps from its own layout on every save
 *  (page distribution, layout keys, pagination receipts): not the visitor's
 *  edit, and the page's editor stamps its own. */
const LAYOUT_ROOT_KEYS = ["pageBlockSizes", "pageBlockLayoutKey", "pageRegionLayoutReceipt", "paginationProjection"];

function withoutLayoutStamps(document: Record<string, unknown>): Record<string, unknown> {
  const attrs = document.attrs;
  if (!isObject(attrs) || !LAYOUT_ROOT_KEYS.some((key) => key in attrs)) return document;
  const rest: JsonObject = { ...attrs };
  for (const key of LAYOUT_ROOT_KEYS) delete rest[key];
  return { ...document, attrs: rest };
}

export function rebaseClassicEdits(input: {
  readonly realBase: Record<string, unknown>;
  readonly guestBase: Record<string, unknown>;
  readonly guestFinal: Record<string, unknown>;
}): ClassicRebaseResult {
  const { realBase } = input;
  if (!isObject(input.guestBase) || !isObject(input.guestFinal)) {
    return { ok: false, reason: "invalid_input", at: "$" };
  }
  const guestBase = withoutLayoutStamps(input.guestBase);
  const guestFinal = withoutLayoutStamps(input.guestFinal);
  if (!isObject(realBase) || !isObject(guestBase) || !isObject(guestFinal)) {
    return { ok: false, reason: "invalid_input", at: "$" };
  }
  if (same(guestBase as Json, guestFinal as Json)) {
    return { ok: true, changed: false, content: realBase };
  }
  const map = new ValueMap();
  learn(guestBase as Json, realBase as Json, "", map);
  try {
    const merged = merge(realBase as Json, guestBase as Json, guestFinal as Json, "", map, "$");
    if (!isObject(merged)) return { ok: false, reason: "invalid_input", at: "$" };
    return { ok: true, changed: true, content: merged };
  } catch (error) {
    if (error instanceof Refusal) return { ok: false, reason: error.reason, at: error.at };
    throw error;
  }
}
