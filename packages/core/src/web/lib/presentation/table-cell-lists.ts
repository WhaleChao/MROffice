/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Home > Bullets / Numbering on the words of a table cell (FIXQUEUE R28).
 *
 * PowerPoint formats a table cell's paragraphs as it formats a text box's
 * (Microsoft Support, "Add bullets or numbers to text"; ECMA-376 §21.1.2.2.7
 * `a:pPr` with `a:buChar` / `a:buAutoNum`, `a:buClr`, `a:buSzPct`, `@marL`,
 * `@indent`): the paragraphs the selection touches in the cell being typed
 * in, or every paragraph of every selected cell. Each paragraph's state lives
 * in the cell's `textLayout` (the shared text-body receipt, `office_text_body`),
 * which the importer reads from the package and the canvas draws
 * (`table-cell-paragraphs.ts`); the deck's writers state it back on the
 * cell's paragraphs (`table_projection.table_cell_list_rows`).
 *
 * A list PowerPoint adds in a cell states PowerPoint's own hanging indent
 * (the corpus' cells: `marL="285750" indent="-285750"` with Arial's `•`,
 * `marL="342900" indent="-342900"` for `a:buAutoNum`); a list taken off
 * returns the paragraph to the cell's edge. A new marker keeps the old one's
 * colour and size, as PowerPoint keeps them.
 */

import type {
  PresentationTableBlock,
  PresentationTableCell,
} from "../../components/console/editor/native-document/documentCommands";
import { POWERPOINT_AUTONUM_BY_NUMBER_FORMAT } from "../office-list-catalog";
import {
  normalizeOfficeTextLayout,
  officeTextLayoutFollowingText,
  type OfficeTextLayout,
  type OfficeTextParagraph,
} from "../office-text-layout";
import type {
  BulletsNumberingResult,
  PresentationListLibraryItem,
} from "../presentation-bullet-library";
import type { NativeBulletMarker } from "../presentation-paragraphs";
import {
  tableCellIntersectsSelection,
  type TableSelectionRectangle,
} from "../presentation-table-selection";

/** PowerPoint's indents for a list it adds in a table cell, in EMU. */
export const TABLE_CELL_BULLET_INDENT_EMU = 285_750;
export const TABLE_CELL_NUMBER_INDENT_EMU = 342_900;

export type TableCellListCommand =
  | { readonly kind: "toggle"; readonly list: "bullet" | "number" }
  | { readonly kind: "none" }
  | { readonly kind: "item"; readonly item: PresentationListLibraryItem }
  /** The Bullets and Numbering dialog's OK: its marker, Size (% of text), Color and Start at. */
  | { readonly kind: "dialog"; readonly result: BulletsNumberingResult };

/** The paragraphs a command reaches in one cell: a span, or all of them. */
export type TableCellParagraphSpan = readonly [number, number] | "all";

type Positions = ReadonlyArray<ReadonlyArray<{ row: number; col: number } | null | undefined>>;

function textOf(cell: PresentationTableCell): string {
  return typeof cell.text === "string" ? cell.text : "";
}

function textOfRuns(value: unknown): string | null {
  if (!value || typeof value !== "object") return null;
  const runs = (value as { runs?: unknown }).runs;
  if (!Array.isArray(runs)) return null;
  let text = "";
  for (const run of runs) {
    if (!run || typeof run !== "object" || typeof (run as { text?: unknown }).text !== "string") return null;
    text += (run as { text: string }).text;
  }
  return text;
}

function autoNumOf(format: string | undefined): string {
  return (
    POWERPOINT_AUTONUM_BY_NUMBER_FORMAT[format as keyof typeof POWERPOINT_AUTONUM_BY_NUMBER_FORMAT] ??
    "arabicPeriod"
  );
}

/**
 * Whether the cell's paragraphs take a list here: a cell whose display runs
 * draw their own markers (`marker`, a body the run writer cannot rewrite)
 * keeps the list it has.
 */
export function tableCellTakesLists(cell: PresentationTableCell): boolean {
  const runs = cell.runs;
  return !(
    Array.isArray(runs) &&
    runs.some((run) => Boolean(run && typeof run === "object" && (run as { marker?: unknown }).marker))
  );
}

/** The cell's paragraphs' rows, one per paragraph of its words as they are now. */
export function tableCellListRows(cell: PresentationTableCell): OfficeTextParagraph[] {
  const text = textOf(cell);
  const count = text.split("\n").length;
  const layout = normalizeOfficeTextLayout(cell.textLayout);
  const readWith = textOfRuns(cell.nativeText) ?? text;
  const followed = layout ? officeTextLayoutFollowingText(layout, readWith, text) : undefined;
  const rows = followed?.paragraphs;
  if (rows && rows.length === count) return rows.map((row) => ({ ...row }));
  return Array.from({ length: count }, () => ({ level: 0, bullet: "none" as const }));
}

/** The list kind the rows share, "mixed", or null for no rows. */
export function tableCellListKind(
  rows: readonly OfficeTextParagraph[],
): OfficeTextParagraph["bullet"] | "mixed" | null {
  const kinds = new Set(rows.map((row) => row.bullet));
  if (kinds.size === 0) return null;
  return kinds.size === 1 ? [...kinds][0]! : "mixed";
}

/** The gallery item every row draws, or null (none, several, or one the gallery lacks). */
export function tableCellListItemId(
  rows: readonly OfficeTextParagraph[],
  library: readonly PresentationListLibraryItem[],
): string | null {
  const ids = new Set(
    rows.map((row) => {
      const native = row.nativeBullet;
      const item = library.find((candidate) => {
        if (row.bullet === "bullet" && candidate.kind === "bullet") {
          if (candidate.char !== undefined)
            return native?.char === candidate.char && native?.typeface === candidate.typeface;
          const glyph = native?.char ?? row.bulletChar;
          return glyph === candidate.glyph;
        }
        if (row.bullet === "number" && candidate.kind === "number") {
          const scheme = native?.autoNum ?? autoNumOf(row.numberFormat);
          return scheme === (candidate.autoNum ?? autoNumOf(candidate.format));
        }
        return false;
      });
      return item?.id;
    }),
  );
  return ids.size === 1 ? ([...ids][0] ?? null) : null;
}

/** The marker's own colour and size a new marker keeps. */
function markerLook(row: OfficeTextParagraph): Pick<NativeBulletMarker, "color" | "sizePct"> {
  const look: Pick<NativeBulletMarker, "color" | "sizePct"> = {};
  if (row.bullet !== "none" && row.nativeBullet?.color) look.color = row.nativeBullet.color;
  if (row.bullet !== "none" && row.nativeBullet?.sizePct !== undefined) look.sizePct = row.nativeBullet.sizePct;
  return look;
}

function hasLook(look: Pick<NativeBulletMarker, "color" | "sizePct">): boolean {
  return Boolean(look.color) || look.sizePct !== undefined;
}

function bulletRow(
  row: OfficeTextParagraph,
  char: { glyph?: string; char?: string; typeface?: string },
): OfficeTextParagraph {
  const look = markerLook(row);
  const next: OfficeTextParagraph = {
    ...row,
    bullet: "bullet",
    marL: TABLE_CELL_BULLET_INDENT_EMU,
    indent: -TABLE_CELL_BULLET_INDENT_EMU,
  };
  delete next.numberFormat;
  delete next.nativeBullet;
  delete next.bulletChar;
  if (char.char !== undefined) {
    next.nativeBullet = { char: char.char, ...(char.typeface ? { typeface: char.typeface } : {}), ...look };
  } else if (char.glyph === undefined || char.glyph === "•") {
    // PowerPoint's own bullet in a cell: Arial's round bullet.
    next.nativeBullet = { char: "•", typeface: "Arial", ...look };
  } else {
    next.bulletChar = char.glyph;
    if (hasLook(look)) next.nativeBullet = { char: char.glyph, ...look };
  }
  return next;
}

function numberRow(
  row: OfficeTextParagraph,
  format: { format?: string; autoNum?: string },
): OfficeTextParagraph {
  const look = markerLook(row);
  const next: OfficeTextParagraph = {
    ...row,
    bullet: "number",
    marL: TABLE_CELL_NUMBER_INDENT_EMU,
    indent: -TABLE_CELL_NUMBER_INDENT_EMU,
  };
  delete next.bulletChar;
  delete next.nativeBullet;
  delete next.numberFormat;
  if (format.autoNum !== undefined) {
    next.nativeBullet = { autoNum: format.autoNum, ...look };
  } else {
    next.numberFormat = format.format ?? "decimal";
    if (hasLook(look)) next.nativeBullet = { autoNum: autoNumOf(next.numberFormat), ...look };
  }
  return next;
}

function plainRow(row: OfficeTextParagraph): OfficeTextParagraph {
  const next: OfficeTextParagraph = { ...row, bullet: "none" };
  delete next.bulletChar;
  delete next.numberFormat;
  delete next.nativeBullet;
  delete next.marL;
  delete next.indent;
  return next;
}

/** The dialog's marker with its Size, Color and Start at (Automatic and 100 % state none). */
function dialogRow(row: OfficeTextParagraph, result: BulletsNumberingResult): OfficeTextParagraph {
  const item = result.item;
  if (item === null) return row.bullet === "none" ? row : plainRow(row);
  const base =
    item.kind === "bullet"
      ? bulletRow(row, { glyph: item.glyph, char: item.char, typeface: item.typeface })
      : numberRow(row, { format: item.format, autoNum: item.autoNum });
  const marker: NativeBulletMarker =
    base.nativeBullet ??
    (item.kind === "bullet" ? { char: base.bulletChar ?? "•" } : { autoNum: autoNumOf(base.numberFormat) });
  // The marker's identity; the dialog states its own colour, size and start.
  const identity: NativeBulletMarker = { ...marker };
  delete identity.color;
  delete identity.sizePct;
  delete identity.startAt;
  const stated: NativeBulletMarker = {
    ...identity,
    ...(result.color ? { color: result.color.toLowerCase() } : {}),
    ...(result.sizePercent !== 100 ? { sizePct: result.sizePercent * 1000 } : {}),
    ...(item.kind === "number" && result.startAt !== 1 ? { startAt: result.startAt } : {}),
  };
  const next: OfficeTextParagraph = { ...base };
  if (base.nativeBullet || Object.keys(stated).length > Object.keys(identity).length) next.nativeBullet = stated;
  else delete next.nativeBullet;
  return next;
}

function rowsWithCommand(
  rows: readonly OfficeTextParagraph[],
  span: TableCellParagraphSpan,
  command: TableCellListCommand,
  /** Whether every paragraph the whole command reaches already has the list (a toggle takes it off). */
  allListed: boolean,
): OfficeTextParagraph[] {
  const [first, last] = span === "all" ? [0, rows.length - 1] : span;
  return rows.map((row, index) => {
    if (index < first || index > last) return row;
    if (command.kind === "none") return row.bullet === "none" ? row : plainRow(row);
    if (command.kind === "toggle") {
      if (allListed) return plainRow(row);
      if (row.bullet === command.list) return row;
      return command.list === "bullet" ? bulletRow(row, {}) : numberRow(row, {});
    }
    if (command.kind === "dialog") return dialogRow(row, command.result);
    const item = command.item;
    return item.kind === "bullet"
      ? bulletRow(row, { glyph: item.glyph, char: item.char, typeface: item.typeface })
      : numberRow(row, { format: item.format, autoNum: item.autoNum });
  });
}

function layoutWithRows(cell: PresentationTableCell, rows: OfficeTextParagraph[]): OfficeTextLayout | undefined {
  const own = normalizeOfficeTextLayout(cell.textLayout);
  const stated = rows.some(
    (row) => row.bullet !== "none" || row.marL || row.indent || row.spaceBefore || row.spaceAfter,
  );
  if (!stated) return undefined;
  return { ...(own?.bodyPr ? { bodyPr: own.bodyPr } : {}), paragraphs: rows };
}

/** One cell with the command applied to the paragraphs of `span`; the cell itself when nothing changes. */
export function tableCellWithList(
  cell: PresentationTableCell,
  span: TableCellParagraphSpan,
  command: TableCellListCommand,
  allListed?: boolean,
): PresentationTableCell {
  if (!tableCellTakesLists(cell)) return cell;
  const rows = tableCellListRows(cell);
  const reached = span === "all" ? rows : rows.slice(span[0], span[1] + 1);
  const listed =
    allListed ??
    (command.kind === "toggle" && reached.length > 0 && reached.every((row) => row.bullet === command.list));
  const layout = layoutWithRows(cell, rowsWithCommand(rows, span, command, listed));
  const before = normalizeOfficeTextLayout(cell.textLayout);
  if (JSON.stringify(layout ?? null) === JSON.stringify(before ?? null)) return cell;
  const next: PresentationTableCell = { ...cell };
  if (layout) next.textLayout = layout;
  else delete next.textLayout;
  return next;
}

/**
 * The paragraphs a character range of a cell's words reaches, as PowerPoint
 * formats the paragraphs a selection or caret is in.
 */
export function tableCellParagraphSpan(text: string, start: number, end: number): readonly [number, number] {
  const from = Math.max(0, Math.min(start, end, text.length));
  const to = Math.max(0, Math.min(Math.max(start, end), text.length));
  const first = text.slice(0, from).split("\n").length - 1;
  const last = text.slice(0, to).split("\n").length - 1;
  return [first, Math.max(first, last)];
}

function reachedCells(
  block: PresentationTableBlock,
  positions: Positions,
  selection: TableSelectionRectangle,
): Array<[number, number]> {
  const reached: Array<[number, number]> = [];
  block.rows.forEach((row, rowIndex) =>
    row.forEach((cell, cellIndex) => {
      const position = positions[rowIndex]?.[cellIndex] ?? null;
      if (tableCellIntersectsSelection(cell, position, selection) && tableCellTakesLists(cell))
        reached.push([rowIndex, cellIndex]);
    }),
  );
  return reached;
}

/**
 * The table with the command applied to every paragraph of every cell the
 * selection reaches. A toggle takes the list off only when every paragraph
 * it reaches already has it, as PowerPoint's button does over several cells.
 */
export function tableWithCellLists(
  block: PresentationTableBlock,
  positions: Positions,
  selection: TableSelectionRectangle,
  command: TableCellListCommand,
): PresentationTableBlock {
  const reached = reachedCells(block, positions, selection);
  if (reached.length === 0) return block;
  const allListed =
    command.kind === "toggle" &&
    reached.every(([row, col]) =>
      tableCellListRows(block.rows[row]![col]!).every((paragraph) => paragraph.bullet === command.list),
    );
  const keys = new Set(reached.map(([row, col]) => `${row}:${col}`));
  let changed = false;
  const rows = block.rows.map((row, rowIndex) =>
    row.map((cell, cellIndex) => {
      if (!keys.has(`${rowIndex}:${cellIndex}`)) return cell;
      const next = tableCellWithList(cell, "all", command, allListed);
      if (next !== cell) changed = true;
      return next;
    }),
  );
  return changed ? { ...block, rows } : block;
}

/** The table with one cell's paragraphs of `span` given the command. */
export function tableWithCellParagraphList(
  block: PresentationTableBlock,
  row: number,
  col: number,
  span: TableCellParagraphSpan,
  command: TableCellListCommand,
): PresentationTableBlock {
  const cell = block.rows[row]?.[col];
  if (!cell) return block;
  const next = tableCellWithList(cell, span, command);
  if (next === cell) return block;
  return {
    ...block,
    rows: block.rows.map((cells, rowIndex) =>
      rowIndex === row ? cells.map((item, cellIndex) => (cellIndex === col ? next : item)) : cells,
    ),
  };
}

/** The rows of every paragraph of the cells a selection reaches. */
export function tableSelectionListRows(
  block: PresentationTableBlock,
  positions: Positions,
  selection: TableSelectionRectangle,
): OfficeTextParagraph[] {
  return reachedCells(block, positions, selection).flatMap(([row, col]) =>
    tableCellListRows(block.rows[row]![col]!),
  );
}

