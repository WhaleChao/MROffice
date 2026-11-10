/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The size of a value's compact JSON text in UTF-8 bytes -- exactly
 * `new Blob([JSON.stringify(value)]).size` -- without writing the text of
 * what was measured before.
 *
 * An object or array frozen all the way down never changes, so its size is
 * kept with it. A
 * handwritten page's body is thousands of frozen strokes under a few objects
 * that change with each stroke: measuring it again reads only those few (the
 * page-full gauge read the whole body as text each minute, 70 ms of an iPad's
 * frame on 3,000 strokes).
 */

const frozenSizes = new WeakMap<object, number>();

/** UTF-8 bytes of a string with no lone surrogate (JSON.stringify escapes
 * those to ASCII), as a Blob encodes it. */
function utf8Bytes(text: string): number {
  let bytes = text.length;
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    if (code < 0x80) continue;
    if (code < 0x800) bytes += 1;
    else if (code >= 0xd800 && code <= 0xdbff && index + 1 < text.length) {
      const next = text.charCodeAt(index + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        // A pair: two UTF-16 units, four bytes.
        bytes += 2;
        index += 1;
      } else bytes += 2;
    } else bytes += 2;
  }
  return bytes;
}

/** Whether JSON.stringify leaves this property out of an object. */
function omitted(value: unknown): boolean {
  return value === undefined || typeof value === "function" || typeof value === "symbol";
}

/** Set while measuring: something under the object being measured may still
 * change (an object not frozen, or one with a form of its own). */
let mayChange = false;

function measure(value: unknown): number {
  if (value === null) return 4;
  switch (typeof value) {
    case "string":
      return utf8Bytes(JSON.stringify(value));
    case "number":
    case "boolean":
      return JSON.stringify(value).length;
    case "bigint":
      // JSON.stringify throws on it: so does this.
      return JSON.stringify(value).length;
    case "object":
      break;
    default:
      return 4;
  }
  const object = value as object;
  if (typeof (object as { toJSON?: unknown }).toJSON === "function" ||
      object instanceof Number || object instanceof String || object instanceof Boolean) {
    // Its own form (a date, which changes even frozen): measured from its
    // text each time.
    mayChange = true;
    return utf8Bytes(JSON.stringify(object) ?? "null");
  }
  const frozen = Object.isFrozen(object);
  if (frozen) {
    const known = frozenSizes.get(object);
    if (known !== undefined) return known;
  }
  const outer = mayChange;
  mayChange = false;
  let size: number;
  if (Array.isArray(object)) {
    size = 2 + Math.max(0, object.length - 1);
    for (let index = 0; index < object.length; index += 1) {
      const item = object[index];
      size += omitted(item) ? 4 : measure(item);
    }
  } else {
    size = 2;
    let first = true;
    for (const key of Object.keys(object)) {
      const item = (object as Record<string, unknown>)[key];
      if (omitted(item)) continue;
      size += (first ? 0 : 1) + utf8Bytes(JSON.stringify(key)) + 1 + measure(item);
      first = false;
    }
  }
  // Kept only for an object frozen all the way down: `Object.freeze` is one
  // level deep, and a frozen object holding one that is not can still change.
  const settled = frozen && !mayChange;
  if (settled) frozenSizes.set(object, size);
  mayChange = outer || !settled;
  return size;
}

/** `new Blob([JSON.stringify(value)]).size`, the frozen parts measured once.
 * Values JSON.stringify does not write (undefined, a function) measure 0. */
export function jsonUtf8Size(value: unknown): number {
  if (omitted(value)) return 0;
  mayChange = false;
  return measure(value);
}
