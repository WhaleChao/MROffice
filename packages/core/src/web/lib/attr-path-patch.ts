/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Attribute paths in a Copilot patch.
 *
 * A patch key with a dot or an index names one leaf of a nested attribute
 * ("rows[0][1].fill", "chartV2.format.legend", "style.fontSize") so a model
 * can change one field of a structure without replacing the structure. The
 * API checks the name and the value against the editor vocabulary
 * (apps/api/app/services/editor_vocabulary.py); this module writes the leaf
 * on the client without touching anything beside it.
 */

export type AttrPathSegment = string | number;

const HEAD = /^[A-Za-z_][A-Za-z0-9_]*/;
const STEP = /\.([A-Za-z_][A-Za-z0-9_]*)|\.(\d{1,4})|\[(\d{1,4})\]/y;
export const MAX_ATTR_PATH_SEGMENTS = 8;
const MAX_INDEX = 999;

/** True for a key that names a path rather than a plain attribute. */
export function isAttrPathKey(key: string): boolean {
  return key.includes(".") || key.includes("[");
}

/** The segments of a path key, or null when the key is malformed. */
export function parseAttrPath(key: string): AttrPathSegment[] | null {
  const head = HEAD.exec(key);
  if (!head) return null;
  const segments: AttrPathSegment[] = [head[0]];
  STEP.lastIndex = head[0].length;
  while (STEP.lastIndex < key.length) {
    const start = STEP.lastIndex;
    const step = STEP.exec(key);
    if (!step || step.index !== start) return null;
    if (step[1] !== undefined) segments.push(step[1]);
    else {
      const index = Number(step[2] ?? step[3]);
      if (index > MAX_INDEX) return null;
      segments.push(index);
    }
  }
  if (segments.length > MAX_ATTR_PATH_SEGMENTS) return null;
  return segments;
}

const isRecord = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === "object" && !Array.isArray(value);

/** A path `setAtAttrPath` cannot write (distinct from a null leaf value). */
const UNWRITABLE = Symbol("unwritable attribute path");

/**
 * `root` with the leaf at `path` set to `value`, cloning only the containers
 * on the path. Null removes a key; written into a list it is the item's value
 * (a chart's blank point, which the chart's own validator admits or refuses).
 * Returns null when the path cannot be written: a name step into a list, an
 * index step into an object, or an index past the end of a list.
 */
export function setAtAttrPath(
  root: unknown,
  path: readonly AttrPathSegment[],
  value: unknown,
): unknown | null {
  const written = writeAtAttrPath(root, path, value);
  return written === UNWRITABLE ? null : written;
}

function writeAtAttrPath(
  root: unknown,
  path: readonly AttrPathSegment[],
  value: unknown,
): unknown {
  if (!path.length) return value;
  const [segment, ...rest] = path;
  if (typeof segment === "number") {
    if (root !== undefined && !Array.isArray(root)) return UNWRITABLE;
    const list = root === undefined ? [] : [...(root as unknown[])];
    if (segment > list.length) return UNWRITABLE;
    if (!rest.length && value === undefined) return UNWRITABLE;
    const child = writeAtAttrPath(list[segment], rest, value);
    if (child === UNWRITABLE) return UNWRITABLE;
    list[segment] = child;
    return list;
  }
  if (root !== undefined && !isRecord(root)) return UNWRITABLE;
  const record = root === undefined ? {} : { ...(root as Record<string, unknown>) };
  if (!rest.length) {
    if (value === null || value === undefined) delete record[segment];
    else record[segment] = value;
    return record;
  }
  const child = writeAtAttrPath(record[segment], rest, value);
  if (child === UNWRITABLE) return UNWRITABLE;
  record[segment] = child;
  return record;
}

/**
 * A patch split into its plain keys and its path keys, or null when a path
 * key is malformed.
 */
export function splitAttrPatch(
  patch: Record<string, unknown>,
): { plain: Record<string, unknown>; paths: [AttrPathSegment[], unknown][] } | null {
  const plain: Record<string, unknown> = {};
  const paths: [AttrPathSegment[], unknown][] = [];
  for (const [key, value] of Object.entries(patch)) {
    if (!isAttrPathKey(key)) {
      plain[key] = value;
      continue;
    }
    const path = parseAttrPath(key);
    if (!path) return null;
    paths.push([path, value]);
  }
  return { plain, paths };
}

/**
 * `attrs` with a patch applied: plain keys replace (null removes), path keys
 * write one leaf each. Null when a path cannot be written.
 */
export function applyAttrPatch(
  attrs: Record<string, unknown>,
  patch: Record<string, unknown>,
): Record<string, unknown> | null {
  const split = splitAttrPatch(patch);
  if (!split) return null;
  let next: Record<string, unknown> = { ...attrs };
  for (const [key, value] of Object.entries(split.plain)) {
    if (value === null || value === undefined) delete next[key];
    else next[key] = value;
  }
  for (const [path, value] of split.paths) {
    const written = setAtAttrPath(next, path, value);
    if (!isRecord(written)) return null;
    next = written;
  }
  return next;
}
