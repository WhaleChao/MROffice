/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiGet } from "@mrnote/core/web/lib/api.ts";

/**
 * Sheets a workbook of many sheets was imported without (SHEET-DEFER-DESIGN.md):
 * each stands on the page as a placeholder (`sheetPending`: its tab, nothing
 * else) and is filled in for the window from the reading the page's saves
 * compare against. What is stored stays the placeholder, as the import made
 * it, until the sheet itself changes: the server fills a placeholder in from
 * that same reading, and refuses a save that brings one back changed or bare.
 */
type SheetLike = { id: string; sheetPending?: unknown };

/** The stored placeholder each sheet model stands for, while it is unchanged. */
const placeholders = new WeakMap<object, Readonly<Record<string, unknown>>>();

/** `model` is `placeholder` as stored: still pending, or filled in and not changed since. */
export function rememberPlaceholderSheet(model: object, placeholder: Readonly<Record<string, unknown>>): void {
  placeholders.set(model, placeholder);
}

/** The placeholder to store for `model`, when it is one (unchanged since its import or its fill). */
export function storedPlaceholderSheet(model: object): Readonly<Record<string, unknown>> | undefined {
  return placeholders.get(model);
}

export function pendingSheetIds(sheets: ReadonlyArray<SheetLike>): string[] {
  return sheets.filter((sheet) => sheet?.sheetPending === true).map((sheet) => sheet.id);
}

/** Whether a sheet of the workbook is still a placeholder: what reads or moves
 * across sheets waits for it, as for a chart still loading. */
export function sheetsStillLoading(sheets: ReadonlyArray<{ sheetPending?: unknown }>): boolean {
  return sheets.some((sheet) => sheet?.sheetPending === true);
}

const NO_PENDING_SHEETS: ReadonlySet<string> = new Set<string>();

/** The lowercased names of the sheets still placeholders: what a formula on
 *  another sheet reads of them is the file's last value until they are in
 *  (kept-formula-results). One shared empty set when there is none. */
export function pendingSheetNames(sheets: ReadonlyArray<{ id: string; name?: unknown; sheetPending?: unknown }>): ReadonlySet<string> {
  let names: Set<string> | null = null;
  for (const sheet of sheets) {
    if (sheet?.sheetPending !== true) continue;
    (names ??= new Set()).add((typeof sheet.name === "string" && sheet.name ? sheet.name : sheet.id).toLowerCase());
  }
  return names ?? NO_PENDING_SHEETS;
}

export type DeferredSheetBodies = {
  sheets: Record<string, Record<string, unknown>>;
  /** What the whole reading could not import, of every sheet. */
  importLosses?: unknown;
};

export async function fetchDeferredSheetBodies(
  pageId: string,
  ids: readonly string[],
  signal?: AbortSignal,
): Promise<DeferredSheetBodies> {
  const answer = await apiGet<DeferredSheetBodies>(
    `/api/v1/pages/${encodeURIComponent(pageId)}/sheet-bodies?ids=${encodeURIComponent(ids.join(","))}`,
    signal ? { signal } : undefined,
  );
  return {
    sheets: answer && typeof answer.sheets === "object" && answer.sheets ? answer.sheets : {},
    ...(answer?.importLosses ? { importLosses: answer.importLosses } : {}),
  };
}

/**
 * The stored content with its placeholders filled in (`filled`, by sheet id):
 * read as a whole, so a filled sheet is normalized as the import would have
 * stored it, beside the sheets it shares a workbook with.
 */
export function storedContentWithSheetsFilled(
  stored: Record<string, unknown>,
  filled: Record<string, Record<string, unknown>>,
): Record<string, unknown> {
  const sheets = Array.isArray(stored.sheets) ? (stored.sheets as Array<Record<string, unknown>>) : [];
  return {
    ...stored,
    sheets: sheets.map((sheet) =>
      sheet?.sheetPending === true && typeof sheet.id === "string" && filled[sheet.id] ? filled[sheet.id] : sheet),
  };
}
