/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where Windows Excel's screen wrap turns "spaced" in a workbook's load
 * (officeHintedWrapLines' `spaced`). Excel fits rows as it opens the file,
 * the active sheet first and then the other shown sheets in order, each from
 * the top; once it has fitted a row through a wrapped cell of complex-script
 * text (Bengali, Arabic, Hebrew, Thai, Devanagari..., or a combining
 * diacritical mark: OFFICE_SPACED_WRAP_TRIGGER_TEXT;
 * a short word that does not wrap counts, unwrapped text does not, a
 * custom-height row is not fitted so does not, a hidden sheet does not), each
 * later wrap of the load keeps room for one space after a line it breaks at a
 * space (lookwin p115b/p117b/p117d/p118/p118b/p118c-lw: R379-X18's formula
 * notes stand 9 and 19 lines, 8 and 18 in a workbook without such a cell).
 * The API's spaced_wrap_after (notebooks.py) is the same rule.
 */
import { cellTextWraps } from "./cell-indexes";
import { OFFICE_SPACED_WRAP_TRIGGER_TEXT } from "@mrnote/core/web/lib/spreadsheet/office-hinted-wrap.ts";

interface TriggerCell {
  readonly value?: unknown;
  readonly runs?: unknown;
  readonly style?: { readonly wrapText?: boolean; readonly horizontalJustify?: unknown };
}

interface TriggerSheet {
  readonly id: string;
  readonly hidden?: boolean;
  readonly veryHidden?: boolean;
  readonly sheetState?: string;
  readonly cells: Readonly<Record<string, TriggerCell | undefined>>;
  readonly nativeGrid?: { readonly rowMetrics?: { readonly rows?: Readonly<Record<string, { readonly customHeight?: boolean } | undefined>> } };
}

/** A sheet's spaced wraps: every one ("all"), those after a (row, column) of
 * its own, or none (absent from the map). */
export type OfficeSpacedWrap = "all" | readonly [number, number];

function cellText(cell: TriggerCell): string {
  const value = typeof cell.value === "string" ? cell.value : "";
  if (!Array.isArray(cell.runs)) return value;
  return value + cell.runs.map((run) => (run && typeof run === "object" ? String((run as { text?: unknown }).text ?? "") : "")).join("");
}

/** Each sheet's trigger by its cells and its rows' metrics: an edit replaces
 * one sheet's cells, and the other sheets are not read again. */
const triggers = new WeakMap<object, { rows: unknown; trigger: readonly [number, number] | undefined }>();

/** The first (row, column) of a sheet whose wrapped complex-script cell Excel fits, or undefined. */
function sheetTrigger(sheet: TriggerSheet): readonly [number, number] | undefined {
  const rows = sheet.nativeGrid?.rowMetrics?.rows;
  const known = triggers.get(sheet.cells);
  if (known && known.rows === rows) return known.trigger;
  const trigger = readSheetTrigger(sheet.cells, rows);
  triggers.set(sheet.cells, { rows, trigger });
  return trigger;
}

function readSheetTrigger(cells: TriggerSheet["cells"], rows: NonNullable<NonNullable<TriggerSheet["nativeGrid"]>["rowMetrics"]>["rows"]): readonly [number, number] | undefined {
  let first: [number, number] | undefined;
  for (const key in cells) {
    const cell = cells[key];
    if (!cell) continue;
    const text = cellText(cell);
    if (!text || !OFFICE_SPACED_WRAP_TRIGGER_TEXT.test(text) || !cellTextWraps(cell.style)) continue;
    const separator = key.indexOf(":");
    const row = Number(key.slice(0, separator)), col = Number(key.slice(separator + 1));
    if (!Number.isInteger(row) || !Number.isInteger(col)) continue;
    if (rows?.[String(row)]?.customHeight === true) continue;
    if (!first || row < first[0] || (row === first[0] && col < first[1])) first = [row, col];
  }
  return first;
}

/** Each shown sheet's spaced wraps, `activeId` the sheet the workbook opens on. */
export function officeSpacedWraps(sheets: readonly TriggerSheet[], activeId: string | undefined): Map<string, OfficeSpacedWrap> {
  const out = new Map<string, OfficeSpacedWrap>();
  const shown = sheets.filter((sheet) => !sheet.hidden && !sheet.veryHidden && !sheet.sheetState);
  const ordered = [...shown.filter((sheet) => sheet.id === activeId), ...shown.filter((sheet) => sheet.id !== activeId)];
  let fromStart = false;
  for (const sheet of ordered) {
    if (fromStart) {
      out.set(sheet.id, "all");
      continue;
    }
    const trigger = sheetTrigger(sheet);
    if (trigger) {
      out.set(sheet.id, trigger);
      fromStart = true;
    }
  }
  return out;
}

/** Whether the cell at (row, col) wraps spaced under `spaced` (strictly after its trigger cell). */
export function officeWrapSpacedAt(spaced: OfficeSpacedWrap | undefined, row: number, col: number): boolean {
  if (!spaced) return false;
  if (spaced === "all") return true;
  return row > spaced[0] || (row === spaced[0] && col > spaced[1]);
}
