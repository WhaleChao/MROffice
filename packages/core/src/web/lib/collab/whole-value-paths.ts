/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { WHOLE_VALUE_PATHS } from "./whole-value-paths.registry";

const patterns = new Set<string>();
const byLastSegment = new Map<string, string[][]>();
/** Whether a pattern could end at an array's index (`*` or a number last): only then is an array's path built. */
let endsAtIndex = false;

/** Whether every path `inner` names starts with a path `outer` names (`outer` shorter). */
const covers = (outer: readonly string[], inner: readonly string[]) =>
  outer.every((segment, index) => segment === "*" || segment === inner[index]);

/**
 * The patterns, checked: each the segments of a value's path, and none that could never take effect (it throws,
 * naming the pattern). Exported for the registry's own test.
 */
export function checkedWholeValuePatterns(list: readonly string[]): string[][] {
  const out: string[][] = [];
  for (const pattern of list) {
    const segments = pattern.split(".");
    if (segments.some((segment) => segment === "")) throw new Error(`whole-value path "${pattern}" has an empty segment`);
    const odd = segments.find((segment) => /[\[\]\s]/.test(segment));
    if (odd !== undefined) throw new Error(`whole-value path "${pattern}": segment "${odd}" can name no JSON key (write a list's entries as .*, not [*])`);
    if (segments.every((segment) => segment === "*")) throw new Error(`whole-value path "${pattern}" is only *: it names no value`);
    out.push(segments);
  }
  for (const inner of out) {
    const outer = out.find((other) => other.length < inner.length && covers(other, inner));
    if (outer) throw new Error(`whole-value path "${inner.join(".")}" is inside "${outer.join(".")}", written whole with it: it can never take effect`);
  }
  return out;
}

function registerWholeValuePath(segments: readonly string[]): void {
  if (patterns.has(segments.join("."))) return;
  patterns.add(segments.join("."));
  const last = segments[segments.length - 1];
  byLastSegment.set(last, [...(byLastSegment.get(last) ?? []), [...segments]]);
  if (last === "*" || /^\d+$/.test(last)) endsAtIndex = true;
}

/** Whether any pattern could end at this key (cheap: no path is built unless one could). */
export function mayBeWholeValueKey(key: string): boolean {
  return byLastSegment.has(key) || byLastSegment.has("*");
}

/** Whether any pattern could end at an entry of a plain array (cheap: false unless one ends in `*` or a number). */
export function mayBeWholeValueIndex(): boolean {
  return endsAtIndex;
}

export function isWholeValuePath(path: readonly string[]): boolean {
  const candidates = [...(byLastSegment.get(path[path.length - 1] ?? "") ?? []), ...(byLastSegment.get("*") ?? [])];
  return candidates.some((segments) => segments.length === path.length && segments.every((segment, index) => segment === "*" || segment === path[index]));
}

export function wholeValuePaths(): readonly string[] {
  return [...patterns];
}

for (const segments of checkedWholeValuePatterns(WHOLE_VALUE_PATHS.map((entry) => entry.pattern))) registerWholeValuePath(segments);





export function registerWholeValuePathForTest(pattern: string): () => void {
  const [segments] = checkedWholeValuePatterns([pattern]);
  const had = { patterns: new Set(patterns), byLastSegment: new Map(byLastSegment), endsAtIndex };
  registerWholeValuePath(segments);
  return () => {
    patterns.clear();
    for (const value of had.patterns) patterns.add(value);
    byLastSegment.clear();
    for (const [key, value] of had.byLastSegment) byLastSegment.set(key, value);
    endsAtIndex = had.endsAtIndex;
  };
}
