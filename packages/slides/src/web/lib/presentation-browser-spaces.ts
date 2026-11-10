/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export function withBrowserSpacesRestored(before: string, after: string): string {
  if (!after.includes(" ")) return after;
  const same = (a: string | undefined, b: string | undefined) => a === b || (a === " " && b === " ");
  const max = Math.min(before.length, after.length);
  let prefix = 0;
  while (prefix < max && same(before[prefix], after[prefix])) prefix += 1;
  let suffix = 0;
  while (
    suffix < max - prefix &&
    same(before[before.length - 1 - suffix], after[after.length - 1 - suffix])
  ) {
    suffix += 1;
  }
  let out = "";
  for (let index = 0; index < after.length; index += 1) {
    const char = after[index];
    const fromBefore =
      index < prefix ? before[index] : index >= after.length - suffix ? before[before.length - (after.length - index)] : undefined;
    out += char === " " && fromBefore === " " ? " " : char;
  }
  return out;
}
