/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Derived numeric formula result, bound to the exact canonical evaluator input. */
export interface ChartValueCache {
  kind: "number";
  source: string;
  value: number;
}

export function matchingChartValueCache(
  value: unknown,
  source: string,
): value is ChartValueCache {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const cache = value as Partial<ChartValueCache>;
  return (
    Object.keys(value).length === 3 &&
    source.trimStart().startsWith("=") &&
    cache.kind === "number" &&
    cache.source === source &&
    typeof cache.value === "number" &&
    Number.isFinite(cache.value)
  );
}
