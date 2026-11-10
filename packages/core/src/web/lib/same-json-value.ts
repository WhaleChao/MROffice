/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Whether two document attribute values are the same value: the same object,
 * or equal JSON. A document read back from the collaboration room carries its
 * attributes as newly parsed objects, equal to the ones before; compared by
 * reference, every such read re-derived what React builds from them (each
 * page's header and footer bands of a long document, on every edit).
 *
 * Document attributes are plain JSON and never mutated once made, so each
 * object's JSON is computed once.
 */
const serialized = new WeakMap<object, string>();

function jsonOf(value: object): string | null {
  const known = serialized.get(value);
  if (known !== undefined) return known;
  let json: string;
  try {
    json = JSON.stringify(value) ?? "";
  } catch {
    return null;
  }
  serialized.set(value, json);
  return json;
}

export function sameJsonValue(previous: unknown, next: unknown): boolean {
  if (previous === next) return true;
  if (previous === null || next === null || typeof previous !== "object" || typeof next !== "object") return false;
  const a = jsonOf(previous);
  return a !== null && a === jsonOf(next);
}
