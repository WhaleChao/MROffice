/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// The workbook's custom number formats (XLS-A056, Format Cells > Number >
// Custom > Delete).
//
// Excel's Custom list holds the built-in codes and then every number format
// the workbook defines; Delete removes a workbook format (a built-in one
// cannot be deleted), and every cell that wore it goes back to General
// (Microsoft Support, "Delete a custom number format": "Any cells in the
// workbook that were formatted with the custom format that you deleted are
// displayed in the default General format"). An xlsx keeps those codes in
// styles.xml; MRAI's workbook keeps a number format on the cells (and row /
// column formats) that wear it, so the workbook's list is read from them.
// ---------------------------------------------------------------------------

import { customTypeList } from "@mrnote/core/web/lib/spreadsheet/format-code-builder.ts";

const BUILT_IN = new Set([...customTypeList("en"), ...customTypeList("zh")]);

/** A code of Excel's built-in set (ids 0-49), which Delete leaves alone. */
export function isBuiltInNumberFormat(code: string): boolean {
  return BUILT_IN.has(code.trim());
}

interface FormatHolder {
  style?: { numberFormat?: unknown } | null;
}
interface FormatSheet {
  cells: Record<string, FormatHolder | undefined>;
  rowStyles?: Record<string, { numberFormat?: unknown } | undefined>;
  columnStyles?: Record<string, { numberFormat?: unknown } | undefined>;
}

/** Every custom code the workbook's cells and lines wear, in first-seen order. */
export function workbookCustomNumberFormats(sheets: readonly FormatSheet[]): string[] {
  const seen = new Set<string>();
  const add = (value: unknown) => {
    if (typeof value !== "string") return;
    const code = value.trim();
    if (code === "" || isBuiltInNumberFormat(code)) return;
    seen.add(code);
  };
  for (const sheet of sheets) {
    for (const cell of Object.values(sheet.cells)) add(cell?.style?.numberFormat);
    for (const style of Object.values(sheet.rowStyles ?? {})) add(style?.numberFormat);
    for (const style of Object.values(sheet.columnStyles ?? {})) add(style?.numberFormat);
  }
  return [...seen];
}

function withoutFormat<T extends { numberFormat?: unknown }>(style: T): T | undefined {
  const next = { ...style };
  delete next.numberFormat;
  return Object.keys(next).length > 0 ? next : undefined;
}

/** The sheet with `code` taken off every cell and line that wore it (General). */
export function sheetWithoutNumberFormat<S extends FormatSheet>(sheet: S, code: string): S {
  const target = code.trim();
  const wears = (value: unknown) => typeof value === "string" && value.trim() === target;
  let cells: S["cells"] | null = null;
  for (const [key, cell] of Object.entries(sheet.cells)) {
    if (!cell?.style || !wears(cell.style.numberFormat)) continue;
    cells ??= { ...sheet.cells };
    const style = withoutFormat(cell.style);
    const next = { ...cell } as FormatHolder;
    if (style) next.style = style;
    else delete next.style;
    (cells as Record<string, FormatHolder>)[key] = next;
  }
  const lines = (map: FormatSheet["rowStyles"]) => {
    if (!map || !Object.values(map).some((style) => style && wears(style.numberFormat))) return map;
    const out: Record<string, { numberFormat?: unknown }> = {};
    for (const [key, style] of Object.entries(map)) {
      if (!style) continue;
      const next = wears(style.numberFormat) ? withoutFormat(style) : style;
      if (next) out[key] = next;
    }
    return out;
  };
  const rowStyles = lines(sheet.rowStyles);
  const columnStyles = lines(sheet.columnStyles);
  if (!cells && rowStyles === sheet.rowStyles && columnStyles === sheet.columnStyles) return sheet;
  return {
    ...sheet,
    ...(cells ? { cells } : {}),
    ...(rowStyles !== sheet.rowStyles ? { rowStyles } : {}),
    ...(columnStyles !== sheet.columnStyles ? { columnStyles } : {}),
  };
}
