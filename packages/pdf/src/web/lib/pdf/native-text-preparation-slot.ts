/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** One speculative result per window. Taking transfers exclusive ownership;
 * eviction, expiry and teardown dispose even a result that finishes late. */
export function createPdfNativeTextPreparationSlot<T>(dispose: (value: T) => void, ttl = 15_000) {
  type Entry = { key: string; target: object; context?: object; controller: AbortController; promise: Promise<T>;
    timer: ReturnType<typeof setTimeout> };
  let current: Entry | null = null;
  const clear = () => {
    const entry = current;
    current = null;
    if (!entry) return;
    clearTimeout(entry.timer);
    entry.controller.abort();
    void entry.promise.then(dispose, () => undefined);
  };
  const warm = (key: string, target: object, prepare: (signal: AbortSignal) => Promise<T>, context?: object) => {
    if (current?.key === key && current.target === target && current.context === context) return current;
    clear();
    const controller = new AbortController();
    const promise = Promise.resolve().then(() => prepare(controller.signal));
    const entry: Entry = { key, target, context, controller, promise, timer: setTimeout(() => {
      if (current === entry) clear();
    }, ttl) };
    current = entry;
    void promise.catch(() => { if (current === entry) clear(); });
    return entry;
  };
  return {
    warm,
    take(key: string, target: object, prepare: (signal: AbortSignal) => Promise<T>, context?: object) {
      const entry = warm(key, target, prepare, context);
      current = null;
      clearTimeout(entry.timer);
      return entry;
    },
    clear,
  };
}
