/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import * as Y from "yjs";

const plain = (value: unknown): value is Record<string, unknown> => {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
};

/** Whether `value` holds an `undefined` anywhere a JSON value cannot. */
export function holdsUndefined(value: unknown, depth = 0): boolean {
  if (value === undefined) return true;
  if (depth > 512) return false;
  if (Array.isArray(value)) return value.some(item => holdsUndefined(item, depth + 1));
  if (plain(value)) {
    for (const key in value) if (holdsUndefined(value[key], depth + 1)) return true;
  }
  return false;
}

/**
 * `value` as JSON says it: object keys holding `undefined` left out, array
 * items that are `undefined` as null. The same reference when it holds none
 * (the usual write), so nothing is copied for it. Yjs types, bytes and other
 * non-plain objects are returned as they are.
 */
export function jsonCleanYValue<T>(value: T): T {
  if (!holdsUndefined(value)) return value;
  return clean(value, 0) as T;
}

function clean(value: unknown, depth: number): unknown {
  if (value === undefined) return null;
  if (depth > 512) return value;
  if (Array.isArray(value)) return value.map(item => clean(item, depth + 1));
  if (plain(value)) {
    const out: Record<string, unknown> = {};
    for (const key in value) {
      const item = value[key];
      if (item !== undefined) out[key] = clean(item, depth + 1);
    }
    return out;
  }
  return value;
}

const GUARDED = Symbol.for("mrai.yJsonWriteGuard");
let bypass = 0;

/**
 * Make every write of a JSON value into a Yjs document JSON-clean. Idempotent
 * and global to this realm (the prototypes are shared by every Y.Doc and by
 * y-tiptap's binding). Yjs's own integration code does not go through these
 * methods, so applying remote updates is unaffected.
 */
export function installYJsonWriteGuard(): void {
  const marker = Y.XmlElement.prototype as unknown as Record<symbol, boolean>;
  if (marker[GUARDED]) return;
  marker[GUARDED] = true;

  const setAttribute = Y.XmlElement.prototype.setAttribute;
  Y.XmlElement.prototype.setAttribute = function guardedSetAttribute(this: Y.XmlElement, name: string, value: unknown) {
    if (bypass) return setAttribute.call(this, name, value as never);
    if (value === undefined) {
      // y-tiptap writes a schema default of `undefined` for a node it makes;
      // absent says the same to every reader.
      if (this.doc === null || Object.hasOwn(this.getAttributes(), name)) this.removeAttribute(name);
      return;
    }
    return setAttribute.call(this, name, jsonCleanYValue(value) as never);
  } as typeof setAttribute;

  const mapSet = Y.Map.prototype.set;
  Y.Map.prototype.set = function guardedMapSet(this: Y.Map<unknown>, key: string, value: unknown) {
    if (bypass) return mapSet.call(this, key, value as never);
    if (value === undefined) {
      if (this.doc === null || this.has(key)) this.delete(key);
      return value as never;
    }
    return mapSet.call(this, key, jsonCleanYValue(value) as never);
  } as typeof mapSet;

  const arrayInsert = Y.Array.prototype.insert;
  Y.Array.prototype.insert = function guardedArrayInsert(this: Y.Array<unknown>, index: number, content: unknown[]) {
    return arrayInsert.call(this, index, !bypass && holdsUndefined(content) ? content.map(item => jsonCleanYValue(item === undefined ? null : item)) : content);
  } as typeof arrayInsert;
  const arrayPush = Y.Array.prototype.push;
  Y.Array.prototype.push = function guardedArrayPush(this: Y.Array<unknown>, content: unknown[]) {
    return arrayPush.call(this, !bypass && holdsUndefined(content) ? content.map(item => jsonCleanYValue(item === undefined ? null : item)) : content);
  } as typeof arrayPush;
  const arrayUnshift = Y.Array.prototype.unshift;
  Y.Array.prototype.unshift = function guardedArrayUnshift(this: Y.Array<unknown>, content: unknown[]) {
    return arrayUnshift.call(this, !bypass && holdsUndefined(content) ? content.map(item => jsonCleanYValue(item === undefined ? null : item)) : content);
  } as typeof arrayUnshift;
}

/**
 * Write a map's values that hold an `undefined` (written before the guard,
 * or by a browser without it) as JSON says them. Returns whether anything
 * was written. The caller owns the transaction.
 */
export function omitUndefinedMapValues(map: Y.Map<unknown>): boolean {
  let changed = false;
  for (const [key, value] of [...map.entries()]) {
    if (value instanceof Y.AbstractType) continue;
    if (value === undefined) { map.delete(key); changed = true; }
    else if (holdsUndefined(value)) { map.set(key, clean(value, 0)); changed = true; }
  }
  return changed;
}

/** `value` with every `undefined` in it as null: the API's (pycrdt's) read. */
function undefinedAsNull(value: unknown, depth: number): unknown {
  if (value === undefined) return null;
  if (depth > 512) return value;
  if (Array.isArray(value)) return value.map(item => undefinedAsNull(item, depth + 1));
  if (plain(value)) {
    const out: Record<string, unknown> = {};
    for (const key in value) out[key] = undefinedAsNull(value[key], depth + 1);
    return out;
  }
  return value;
}











export function writeUndefinedAsNull(roots: ReadonlyArray<Y.AbstractType<unknown>>): boolean {
  let changed = false;
  const visit = (type: Y.AbstractType<unknown>, depth: number): void => {
    if (depth > 512) return;
    if (type instanceof Y.XmlElement) {
      for (const [key, value] of Object.entries(type.getAttributes() as Record<string, unknown>)) {
        if (value instanceof Y.AbstractType) { visit(value, depth + 1); continue; }
        if (value === undefined || holdsUndefined(value)) {
          type.setAttribute(key, undefinedAsNull(value, 0) as never);
          changed = true;
        }
      }
    }
    if (type instanceof Y.XmlFragment || type instanceof Y.XmlElement) {
      for (const child of type.toArray()) if (child instanceof Y.AbstractType) visit(child as Y.AbstractType<unknown>, depth + 1);
      return;
    }
    if (type instanceof Y.Map) {
      for (const [key, value] of [...type.entries()]) {
        if (value instanceof Y.AbstractType) { visit(value, depth + 1); continue; }
        if (value === undefined || holdsUndefined(value)) {
          type.set(key, undefinedAsNull(value, 0));
          changed = true;
        }
      }
      return;
    }
    if (type instanceof Y.Array) {
      const items = type.toArray();
      for (let index = 0; index < items.length; index += 1) {
        const value = items[index];
        if (value instanceof Y.AbstractType) { visit(value, depth + 1); continue; }
        if (value === undefined || holdsUndefined(value)) {
          type.delete(index, 1);
          type.insert(index, [undefinedAsNull(value, 0)]);
          changed = true;
        }
      }
    }
  };
  // The guard would write a key holding undefined as absent: null is meant.
  bypass += 1;
  try {
    for (const root of roots) {
      const doc = root.doc;
      if (doc) doc.transact(() => visit(root, 0));
      else visit(root, 0);
    }
  } finally { bypass -= 1; }
  return changed;
}

/** Runs `write` with the guard off: only for tests that must build the
 * state an older browser left in a room. */
export function withoutYJsonWriteGuardForTests<T>(write: () => T): T {
  bypass += 1;
  try { return write(); } finally { bypass -= 1; }
}
