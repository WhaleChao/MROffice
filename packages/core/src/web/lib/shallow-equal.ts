/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The two records hold the same values, key by key (`Object.is`): what a
 * memoized component compares its props or its `ctl` bag by, so a bag
 * rebuilt every render with the same contents does not draw it again.
 */
export function shallowEqualRecords(
  a: Readonly<Record<string, unknown>> | null | undefined,
  b: Readonly<Record<string, unknown>> | null | undefined,
): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  const keys = Object.keys(a);
  if (keys.length !== Object.keys(b).length) return false;
  for (const key of keys) {
    if (!Object.prototype.hasOwnProperty.call(b, key) || !Object.is(a[key], b[key])) return false;
  }
  return true;
}
