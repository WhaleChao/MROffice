/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useMemo } from "react";

type AnyTranslate = ((key: string, ...rest: unknown[]) => string) &
  Record<string, unknown>;

/**
 * Wraps a next-intl translator so a message asked for without values is
 * formatted once, not on every call. A large editor render reads thousands of
 * static labels (every ribbon caption, tooltip and menu item) and next-intl
 * resolves and formats each one again on each call; the argument-free result
 * of a translator is fixed for that translator's locale and messages, so it
 * can be kept for the translator's lifetime.
 *
 * Calls with values or formats go straight through. `rich`, `markup`, `raw`
 * and `has` are the underlying translator's own. A new translator (another
 * locale or message set) gets a new, empty cache.
 */
export function memoizedTranslator<T>(translate: T): T {
  const base = translate as unknown as AnyTranslate;
  const cache = new Map<string, string>();
  const memoized = (key: string, ...rest: unknown[]): string => {
    if (rest.some((argument) => argument !== undefined)) return base(key, ...rest);
    const hit = cache.get(key);
    if (hit !== undefined) return hit;
    const value = base(key);
    cache.set(key, value);
    return value;
  };
  for (const name of Object.keys(base)) {
    (memoized as unknown as Record<string, unknown>)[name] = base[name];
  }
  return memoized as unknown as T;
}

/** `memoizedTranslator` for a hook's translator, kept per translator identity. */
export function useMemoizedTranslator<T>(translate: T): T {
  return useMemo(() => memoizedTranslator(translate), [translate]);
}
