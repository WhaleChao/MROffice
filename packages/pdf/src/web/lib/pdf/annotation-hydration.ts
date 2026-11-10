/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Apply one server read of a reader annotation list without losing what the
 * user did while that read was in flight.
 *
 * `base` is the list as it stood when the read began, `current` the list now
 * and `server` what the read returned. When nothing changed locally the server
 * list simply replaces the state. Otherwise, per id:
 *
 * - added or edited locally during the read → the local item wins;
 * - deleted locally during the read → it stays deleted even if the server
 *   (read before the deletion landed) still lists it;
 * - untouched locally and absent from the server → the server removed it;
 * - otherwise the server item is taken.
 *
 * Items keep the server order; items only the local side knows follow in
 * their local order.
 */
export function mergeHydratedAnnotationList<T extends { readonly id: string }>(
  base: readonly T[],
  current: readonly T[],
  server: readonly T[],
): T[] {
  if (current === base) return [...server];
  const baseById = new Map(base.map((item) => [item.id, item]));
  const currentById = new Map(current.map((item) => [item.id, item]));
  const merged: T[] = [];
  const seen = new Set<string>();
  for (const item of server) {
    if (seen.has(item.id)) continue;
    seen.add(item.id);
    const local = currentById.get(item.id);
    const before = baseById.get(item.id);
    if (!local) {
      if (before) continue;
      merged.push(item);
      continue;
    }
    merged.push(local !== before ? local : item);
  }
  for (const local of current) {
    if (seen.has(local.id)) continue;
    seen.add(local.id);
    if (baseById.get(local.id) === local) continue;
    merged.push(local);
  }
  return merged;
}
