/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Watched names in the debugger — the one standard debug panel this editor did
 * not have.
 *
 * Breakpoints, stepping, the call stack and per-frame locals were all here, but
 * following one variable meant re-reading the locals list at every stop and
 * finding it again. Every debugger solves that by letting you pin a few names.
 *
 * **Names, not expressions, and that is a statement about this debugger rather
 * than a shortcut.** Python here is a RECORDING debugger: the run finishes, and
 * the panel replays what the tracer captured. There is no live interpreter at a
 * pause to evaluate `len(items) > 2` against, and a watch list that silently
 * failed on anything but a bare name would be worse than one that says up front
 * what it takes. What the recording does have is each frame's locals, and a
 * pinned NAME is answered exactly and at every stop.
 *
 * A name that is not in the frame being viewed is reported as out of scope, not
 * as empty — "this frame does not have it" and "it is the empty string" are
 * different answers, and a debugger that confuses them is worse than useless.
 */

/** A pinned name. `id` is stable so a list can be reordered or removed by key. */
export interface DebugWatch {
  readonly id: string;
  readonly name: string;
}

/** What a watch resolves to at the currently viewed stop. */
export interface ResolvedDebugWatch extends DebugWatch {
  /** The recorded repr, or undefined when this frame does not have the name. */
  readonly value?: string;
  /** True when the frame's locals were captured and simply lack this name. */
  readonly outOfScope: boolean;
}

export const MAX_DEBUG_WATCHES = 20;
export const MAX_DEBUG_WATCH_NAME_CHARS = 128;

/**
 * Identifiers this can actually answer: one bare name.
 *
 * Deliberately permissive about the alphabet (Python, JS, Go and friends all
 * allow `_` and digits after the first character) and deliberately strict about
 * the SHAPE — no dots, no calls, no indexing. The locals snapshot maps a name
 * to a string repr, so `obj.field` could never be resolved from it, and
 * accepting it would promise an answer that cannot exist.
 */
const WATCH_NAME =
  /^[$_\p{ID_Start}][$\u200c\u200d\p{ID_Continue}]*$/u;

export function isWatchableName(name: string): boolean {
  return (
    typeof name === "string" &&
    name.length <= MAX_DEBUG_WATCH_NAME_CHARS &&
    WATCH_NAME.test(name)
  );
}

/**
 * Add a name to the list.
 *
 * Returns the SAME list when the name is unusable or already pinned, so a
 * caller can treat an unchanged reference as "nothing to do" and a duplicate
 * never quietly doubles the panel.
 */
export function addDebugWatch(
  watches: readonly DebugWatch[],
  rawName: string,
  makeId: () => string,
): readonly DebugWatch[] {
  const name = typeof rawName === "string" ? rawName.trim() : "";
  if (!isWatchableName(name)) return watches;
  if (watches.length >= MAX_DEBUG_WATCHES) return watches;
  if (watches.some((watch) => watch.name === name)) return watches;
  return [...watches, { id: makeId(), name }];
}

export function removeDebugWatch(
  watches: readonly DebugWatch[],
  id: string,
): readonly DebugWatch[] {
  const next = watches.filter((watch) => watch.id !== id);
  return next.length === watches.length ? watches : next;
}

/**
 * Rename one pinned watch without changing its stable id or list position.
 *
 * Invalid names, duplicate names, missing ids and no-op renames return the SAME
 * list, allowing the panel to keep the draft visible and avoid publishing a
 * misleading configuration change.
 */
export function renameDebugWatch(
  watches: readonly DebugWatch[],
  id: string,
  rawName: string,
): readonly DebugWatch[] {
  const name = typeof rawName === "string" ? rawName.trim() : "";
  const index = watches.findIndex((watch) => watch.id === id);
  if (
    index < 0 ||
    !isWatchableName(name) ||
    watches[index]!.name === name ||
    watches.some((watch, candidateIndex) =>
      candidateIndex !== index && watch.name === name
    )
  ) {
    return watches;
  }
  return watches.map((watch, candidateIndex) =>
    candidateIndex === index ? { ...watch, name } : watch,
  );
}

/**
 * Resolve every watch against the locals of the frame currently being viewed.
 *
 * `locals` is null when the frame's locals were never captured — an outer frame
 * the tracer did not record. That is NOT the same as the name being absent, so
 * every watch reports `outOfScope: false` and no value: the honest answer is
 * "nothing was recorded here", and the panel says that once for the frame
 * rather than implying each name is missing.
 */
export function resolveDebugWatches(
  watches: readonly DebugWatch[],
  locals: Readonly<Record<string, string>> | null | undefined,
): ResolvedDebugWatch[] {
  return watches.map((watch) => {
    if (!locals) return { ...watch, outOfScope: false };
    return Object.prototype.hasOwnProperty.call(locals, watch.name)
      ? { ...watch, value: locals[watch.name], outOfScope: false }
      : { ...watch, outOfScope: true };
  });
}

/**
 * Normalize a persisted watch list from untrusted storage.
 *
 * Same shape of caution as the breakpoint list it will be stored beside: a
 * bounded scan, unusable entries dropped rather than repaired, and duplicates
 * collapsed so a hostile payload cannot make the panel unbounded.
 */
export function normalizeDebugWatches(value: unknown): DebugWatch[] {
  if (!Array.isArray(value)) return [];
  const seen = new Set<string>();
  const watches: DebugWatch[] = [];
  const scanLimit = Math.min(value.length, MAX_DEBUG_WATCHES * 4);
  for (let index = 0; index < scanLimit; index += 1) {
    if (watches.length >= MAX_DEBUG_WATCHES) break;
    const candidate = value[index];
    if (!candidate || typeof candidate !== "object") continue;
    const record = candidate as Record<string, unknown>;
    const name = typeof record.name === "string" ? record.name.trim() : "";
    if (!isWatchableName(name) || seen.has(name)) continue;
    const id =
      typeof record.id === "string" && record.id.length > 0 && record.id.length <= 64
        ? record.id
        : `watch-${name}`;
    seen.add(name);
    watches.push({ id, name });
  }
  return watches;
}
