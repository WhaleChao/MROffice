/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A chart imported with its model deferred is filled in for the window alone
 * (lib/spreadsheet/deferred-charts): what is stored stays its placeholder until
 * the chart itself changes, since the server reads a placeholder as the chart
 * the workbook's package holds. Storing 72 charts the person had not touched
 * took the server 40 s after every such import, and an export of the untouched
 * file could no longer hand back the upload.
 *
 * Any change to a chart (its format, its anchor, its references following an
 * insert) makes a new object, which has no placeholder here and is stored whole.
 */
const placeholders = new WeakMap<object, Readonly<Record<string, unknown>>>();

export function rememberFilledChart(whole: object, placeholder: Readonly<Record<string, unknown>>): void {
  placeholders.set(whole, placeholder);
}

type SheetWithObjects = { objects?: ReadonlyArray<unknown> };

/** The sheets as stored: each chart filled in and unchanged since, its placeholder. The same array when none is. */
export function withFilledChartsAsPlaceholders<S extends object>(sheets: readonly S[]): S[] {
  let changed = false;
  const stored = sheets.map((sheet) => {
    const objects = (sheet as SheetWithObjects).objects;
    if (!objects?.some((object) => typeof object === "object" && object !== null && placeholders.has(object))) return sheet;
    changed = true;
    return {
      ...sheet,
      objects: objects.map((object) => (typeof object === "object" && object !== null && placeholders.get(object)) || object),
    };
  });
  return changed ? stored : (sheets as S[]);
}
