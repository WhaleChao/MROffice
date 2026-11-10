/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const KEY = "mrai-page-costs-v1";
const KEPT = 200;
/** A new measurement is written only when it differs this much (fraction). */
const WORTH_WRITING = 0.25;

export type RememberedPageCostKind = "save" | "recover-copy";

let costs: Record<string, number> | null = null;

function load(): Record<string, number> {
  if (costs) return costs;
  costs = {};
  try {
    const parsed = JSON.parse(window.localStorage.getItem(KEY) ?? "{}");
    if (parsed && typeof parsed === "object") {
      for (const [key, value] of Object.entries(parsed)) {
        if (typeof value === "number" && Number.isFinite(value) && value >= 0) costs[key] = value;
      }
    }
  } catch {
    // nothing remembered
  }
  return costs;
}

/** The cost measured in an earlier visit, ms; 0 when none is known. */
export function rememberedPageCost(kind: RememberedPageCostKind, pageId: string): number {
  if (typeof window === "undefined" || !pageId) return 0;
  return load()[`${kind}:${pageId}`] ?? 0;
}

export function rememberPageCost(kind: RememberedPageCostKind, pageId: string, ms: number): void {
  if (typeof window === "undefined" || !pageId || !Number.isFinite(ms) || ms < 0) return;
  const all = load();
  const key = `${kind}:${pageId}`;
  const before = all[key];
  if (before !== undefined && Math.abs(ms - before) <= Math.max(before, 1) * WORTH_WRITING) return;
  delete all[key];
  all[key] = Math.round(ms);
  const keys = Object.keys(all);
  for (const old of keys.slice(0, Math.max(0, keys.length - KEPT))) delete all[old];
  try {
    window.localStorage.setItem(KEY, JSON.stringify(all));
  } catch {
    // kept for this visit only
  }
}

export function resetRememberedPageCostsForTests(): void {
  costs = null;
}
