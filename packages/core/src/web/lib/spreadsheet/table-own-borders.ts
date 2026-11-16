/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { BorderLineStyle, CellBorderEdges } from "./cell-format";
import { BORDER_LINE_STYLES } from "./cell-format";
import { normalizeThemeColor, type SpreadsheetThemeColor } from "./theme-colors";

export interface TableOwnEdge {
  style: BorderLineStyle;
  /** The colour drawn ("#rrggbb"); none where the file's is automatic
   *  (reading xlsxTableBorderAutoColour, xlsx-rounds R431-X41), drawn as a
   *  cell's own automatic edge is. */
  color?: string;
  /** The theme colour the file names, kept for the save. */
  theme?: SpreadsheetThemeColor;
}

type Side = "top" | "right" | "bottom" | "left";

export interface TableOwnBorders {
  /** tableBorderDxfId: the outline of the whole table, its totals row included. */
  table?: Partial<Record<Side, TableOwnEdge>>;
  /** headerRowBorderDxfId: the header row's top and bottom lines. */
  header?: Partial<Record<"top" | "bottom", TableOwnEdge>>;
  /** totalsRowBorderDxfId: the totals row's top and bottom lines. */
  totals?: Partial<Record<"top" | "bottom", TableOwnEdge>>;
}

/** The table geometry the borders are laid on (a DataTable's). */
export interface TableOwnBordersTable {
  range: { r1: number; c1: number; r2: number; c2: number };
  headerRow: boolean;
  totalsRow?: boolean;
  ownBorders?: TableOwnBorders;
}

function normalizeEdge(raw: unknown): TableOwnEdge | null {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const { style, color, theme } = raw as Record<string, unknown>;
  if (typeof style !== "string" || !(BORDER_LINE_STYLES as readonly string[]).includes(style)) return null;
  if (color !== undefined && (typeof color !== "string" || !/^#[0-9a-f]{6}$/i.test(color))) return null;
  const ref = theme === undefined ? null : normalizeThemeColor(theme);
  return {
    style: style as BorderLineStyle,
    ...(typeof color === "string" ? { color: color.toLowerCase() } : {}),
    ...(ref ? { theme: ref } : {}),
  };
}

function normalizeSides<S extends string>(raw: unknown, sides: readonly S[]): Partial<Record<S, TableOwnEdge>> | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const out: Partial<Record<S, TableOwnEdge>> = {};
  for (const side of sides) {
    const edge = normalizeEdge((raw as Record<string, unknown>)[side]);
    if (edge) out[side] = edge;
  }
  return Object.keys(out).length ? out : undefined;
}

/** A stored `ownBorders`, or undefined when it states no edge. */
export function normalizeTableOwnBorders(raw: unknown): TableOwnBorders | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const r = raw as Record<string, unknown>;
  const table = normalizeSides(r.table, ["top", "right", "bottom", "left"] as const);
  const header = normalizeSides(r.header, ["top", "bottom"] as const);
  const totals = normalizeSides(r.totals, ["top", "bottom"] as const);
  if (!table && !header && !totals) return undefined;
  return { ...(table ? { table } : {}), ...(header ? { header } : {}), ...(totals ? { totals } : {}) };
}

/**
 * The table's own edges at (row, col), as Excel lays them: the header row's
 * lines on the header row, the totals row's on the totals row, then the
 * outline over them (top of the first row, bottom of the last -- the totals
 * row when shown -- left of the first column, right of the last). Undefined
 * where the table draws none.
 */
export function tableOwnEdgesAt(table: TableOwnBordersTable, row: number, col: number): CellBorderEdges | undefined {
  const own = table.ownBorders;
  if (!own) return undefined;
  const { r1, c1, r2, c2 } = table.range;
  const last = table.totalsRow ? r2 + 1 : r2;
  if (row < r1 || row > last || col < c1 || col > c2) return undefined;
  const out: CellBorderEdges = {};
  const lay = (side: Side, edge: TableOwnEdge | undefined) => {
    if (edge) out[side] = { style: edge.style, ...(edge.color ? { color: edge.color } : {}) };
  };
  if (table.headerRow && row === r1 && own.header) {
    lay("top", own.header.top);
    lay("bottom", own.header.bottom);
  }
  if (table.totalsRow && row === last && own.totals) {
    lay("top", own.totals.top);
    lay("bottom", own.totals.bottom);
  }
  if (own.table) {
    if (row === r1) lay("top", own.table.top);
    if (row === last) lay("bottom", own.table.bottom);
    if (col === c1) lay("left", own.table.left);
    if (col === c2) lay("right", own.table.right);
  }
  return Object.keys(out).length ? out : undefined;
}

type OwnEdgeCell = {
  value?: unknown;
  kind?: unknown;
  style?: {
    borderEdges?: CellBorderEdges;
    border?: unknown;
    themeColors?: Record<string, unknown>;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

const SIDE_FLAG: Record<Side, string> = { top: "borderTop", right: "borderRight", bottom: "borderBottom", left: "borderLeft" };

/** The cell's own edge kept under each side a table lays (`tableEdgesUnder`,
 *  the import's reading xlsxTableEdgeOwnUnder): null where the cell had none. */
type OwnUnder = Partial<Record<Side, { style: string; color?: string; theme?: Record<string, unknown> } | null>>;

function ownUnder(style: OwnEdgeCell["style"] | undefined): OwnUnder {
  const raw = style?.tableEdgesUnder;
  return raw && typeof raw === "object" && !Array.isArray(raw) ? { ...(raw as OwnUnder) } : {};
}

function markedSides(style: OwnEdgeCell["style"] | undefined): Side[] {
  const raw = style?.tableEdges;
  return Array.isArray(raw) ? raw.filter((side): side is Side => side === "top" || side === "right" || side === "bottom" || side === "left") : [];
}












export function styleWithOwnEdgeUnderTable<T extends object | undefined>(
  style: T,
  side: string,
  edge: { style: string; color?: string | null; theme?: object | null } | null,
): T {
  if (!style || (side !== "top" && side !== "right" && side !== "bottom" && side !== "left")) return style;
  if (!markedSides(style as OwnEdgeCell["style"]).includes(side as Side)) return style;
  const under = ownUnder(style as OwnEdgeCell["style"]);
  under[side as Side] = edge
    ? {
        style: edge.style,
        ...(edge.color ? { color: edge.color.toLowerCase() } : {}),
        ...(edge.theme ? { theme: { ...(edge.theme as Record<string, unknown>) } } : {}),
      }
    : null;
  return { ...style, tableEdgesUnder: under } as T;
}

/**
 * The cells a structural edit (rows or columns inserted, deleted or moved)
 * may have left a table's own edges wrong on: each table's outline, header
 * and totals rows, as "row:col" keys -- or null past `limit` cells.
 */
export function tableOwnEdgeKeys(tables: readonly TableOwnBordersTable[], limit = 200_000): string[] | null {
  const keys = new Set<string>();
  for (const table of tables) {
    const { r1, c1, r2, c2 } = table.range;
    const last = table.totalsRow ? r2 + 1 : r2;
    const rows = new Set([r1, r1 + 1, last, last - 1]);
    for (const row of rows) {
      if (row < r1 || row > last) continue;
      for (let col = c1; col <= c2; col += 1) keys.add(`${row}:${col}`);
    }
    for (let row = r1; row <= last; row += 1) {
      keys.add(`${row}:${c1}`);
      keys.add(`${row}:${c2}`);
      if (keys.size > limit) return null;
    }
  }
  return [...keys];
}

/** The tables of a sheet that have own borders. */
export function tablesWithOwnBorders(raw: unknown): TableOwnBordersTable[] {
  if (!Array.isArray(raw)) return [];
  const out: TableOwnBordersTable[] = [];
  for (const table of raw) {
    if (!table || typeof table !== "object") continue;
    const t = table as Record<string, unknown>;
    const ownBorders = normalizeTableOwnBorders(t.ownBorders);
    const range = t.range as TableOwnBordersTable["range"] | undefined;
    if (!ownBorders || !range || ![range.r1, range.c1, range.r2, range.c2].every(Number.isInteger)) continue;
    out.push({ range, headerRow: t.headerRow !== false, totalsRow: t.totalsRow === true, ownBorders });
  }
  return out;
}

/** One cell with `edges` laid over its own: its theme receipt for each side
 *  is the table edge's (or none). `drop`: sides a table no longer draws,
 *  taken off the cell. The same cell when nothing changes. */
function cellWithEdges(
  cell: OwnEdgeCell | undefined,
  edges: Partial<Record<Side, TableOwnEdge>>,
  drop: readonly Side[] = [],
): OwnEdgeCell | undefined {
  const style = cell?.style ?? {};
  const before = style.borderEdges ?? {};
  let changed = false;
  const next: CellBorderEdges = { ...before };
  const receipts: Record<string, unknown> = { ...(style.themeColors ?? {}) };
  const flags: Record<string, unknown> = {};
  // The sides an earlier lay marked as the table's (`tableEdges`): copied or
  // moved with the cell, they are taken off where no table draws them now.
  const marked = markedSides(style);
  // The cell's own edges under the table's (R427-X50): a side the table no
  // longer draws gets the cell's own edge back, and a side it starts drawing
  // keeps the cell's own underneath, for the save and for later.
  const under = ownUnder(style);
  const underBefore = JSON.stringify(under);
  // A cell boxed all round (the `border` shorthand) that loses a side is
  // boxed no longer: its other sides become their own flags, or the box
  // would still draw -- and export -- the side taken off (R404-X10 N5).
  let boxed = Boolean(style.border);
  for (const side of new Set([...drop, ...marked])) {
    const own = marked.includes(side) && !edges[side] ? under[side] : undefined;
    if (!edges[side]) delete under[side];
    if (own && typeof own === "object" && typeof own.style === "string" && (BORDER_LINE_STYLES as readonly string[]).includes(own.style)) {
      next[side] = { style: own.style as BorderLineStyle, ...(own.color ? { color: own.color } : {}) };
      if (own.theme) receipts[`borderEdges.${side}.color`] = { ...own.theme };
      else delete receipts[`borderEdges.${side}.color`];
      if (!boxed) flags[SIDE_FLAG[side]] = true;
      changed = true;
      continue;
    }
    if (next[side] && !edges[side]) {
      delete next[side];
      delete receipts[`borderEdges.${side}.color`];
      if (boxed) {
        boxed = false;
        flags.border = undefined;
        for (const flag of Object.values(SIDE_FLAG)) flags[flag] = true;
      }
      flags[SIDE_FLAG[side]] = undefined;
      changed = true;
    }
  }
  for (const [side, edge] of Object.entries(edges) as Array<[Side, TableOwnEdge]>) {
    const had = before[side];
    if (!marked.includes(side) && !(side in under)) {
      // The table starts drawing this side: the cell's own edge stays under.
      const receipt = (style.themeColors ?? {})[`borderEdges.${side}.color`];
      under[side] = had
        ? { style: had.style, ...(had.color ? { color: had.color } : {}), ...(receipt && typeof receipt === "object" ? { theme: { ...(receipt as Record<string, unknown>) } } : {}) }
        : boxed || style[SIDE_FLAG[side]] ? { style: "thin" } : null;
    }
    const receipt = receipts[`borderEdges.${side}.color`] as { theme?: number; tint?: number } | undefined;
    const sameTheme = edge.theme
      ? receipt?.theme === edge.theme.theme && (receipt?.tint ?? 0) === (edge.theme.tint ?? 0)
      : receipt === undefined;
    if (had?.style === edge.style && (had.color ?? "").toLowerCase() === (edge.color ?? "") && sameTheme && (boxed || style[SIDE_FLAG[side]] || flags[SIDE_FLAG[side]])) continue;
    next[side] = { style: edge.style, ...(edge.color ? { color: edge.color } : {}) };
    if (edge.theme) receipts[`borderEdges.${side}.color`] = { ...edge.theme };
    else delete receipts[`borderEdges.${side}.color`];
    if (!boxed) flags[SIDE_FLAG[side]] = true;
    changed = true;
  }
  const marks = (Object.keys(edges) as Side[]).sort();
  if (marks.join() !== [...marked].sort().join()) changed = true;
  for (const side of Object.keys(under) as Side[]) if (!marks.includes(side)) delete under[side];
  if (JSON.stringify(under) !== underBefore) changed = true;
  if (!changed) return cell;
  const nextStyle: Record<string, unknown> = { ...style, ...flags, borderEdges: next };
  for (const [flag, value] of Object.entries(flags)) if (value === undefined) delete nextStyle[flag];
  if (marks.length) nextStyle.tableEdges = marks;
  else delete nextStyle.tableEdges;
  if (marks.length && Object.keys(under).length) nextStyle.tableEdgesUnder = under;
  else delete nextStyle.tableEdgesUnder;
  if (Object.keys(next).length === 0) delete nextStyle.borderEdges;
  if (Object.keys(receipts).length) nextStyle.themeColors = receipts;
  else delete nextStyle.themeColors;
  return { ...(cell ?? { value: "", kind: "text" }), style: nextStyle as OwnEdgeCell["style"] };
}

/**
 * After an edit wrote the cells at `keys`, the table's own edges laid back
 * over those of them a table draws on: Excel keeps drawing a table's own
 * borders over whatever the cell's own become (probe G). The same map when
 * no cell changes.
 */
export function cellsWithTableOwnEdges<M extends object>(
  map: M,
  keys: readonly string[],
  tables: readonly TableOwnBordersTable[],
): M {
  // No table left with own borders still takes marked sides off the cells
  // named (a table deleted or resized away from them).
  if (!keys.length) return map;
  const cells = map as unknown as Readonly<Record<string, OwnEdgeCell | undefined>>;
  let out: Record<string, OwnEdgeCell | undefined> | null = null;
  for (const key of keys) {
    const colon = key.indexOf(":");
    const row = Number(key.slice(0, colon));
    const col = Number(key.slice(colon + 1));
    let edges: Partial<Record<Side, TableOwnEdge>> | null = null;
    for (const table of tables) {
      const own = tableOwnEdgeSpecsAt(table, row, col);
      if (own) edges = { ...(edges ?? {}), ...own };
    }
    const cell = (out ?? cells)[key];
    if (!edges && !markedSides(cell?.style).length) continue;
    const laid = cellWithEdges(cell, edges ?? {});
    if (laid === cell) continue;
    out ??= { ...cells };
    out[key] = laid;
  }
  return (out ?? map) as M;
}

/** As {@link tableOwnEdgesAt}, with each edge's theme colour. */
function tableOwnEdgeSpecsAt(table: TableOwnBordersTable, row: number, col: number): Partial<Record<Side, TableOwnEdge>> | null {
  const own = table.ownBorders;
  if (!own) return null;
  const { r1, c1, r2, c2 } = table.range;
  const last = table.totalsRow ? r2 + 1 : r2;
  if (row < r1 || row > last || col < c1 || col > c2) return null;
  const out: Partial<Record<Side, TableOwnEdge>> = {};
  const lay = (side: Side, edge: TableOwnEdge | undefined) => { if (edge) out[side] = edge; };
  if (table.headerRow && row === r1 && own.header) { lay("top", own.header.top); lay("bottom", own.header.bottom); }
  if (table.totalsRow && row === last && own.totals) { lay("top", own.totals.top); lay("bottom", own.totals.bottom); }
  if (own.table) {
    if (row === r1) lay("top", own.table.top);
    if (row === last) lay("bottom", own.table.bottom);
    if (col === c1) lay("left", own.table.left);
    if (col === c2) lay("right", own.table.right);
  }
  return Object.keys(out).length ? out : null;
}

/**
 * Excel's No Border over exactly the header row of a table that has a
 * header row border (Home > Borders > No Border): the header row's border
 * format goes and the table's outline stays (probe F; Excel removes
 * headerRowBorderDxfId). The sheet with that table changed and the header
 * row's cells without the lines it laid, or null when `rect` is no such row.
 * The outline is laid back by the commit (`cellsWithTableOwnEdges`).
 */
export function sheetWithoutHeaderOwnBorder<S extends { cells: object; dataTables?: unknown }>(
  sheet: S,
  rect: { r1: number; c1: number; r2: number; c2: number },
): S | null {
  if (!Array.isArray(sheet.dataTables)) return null;
  const index = sheet.dataTables.findIndex((raw) => {
    const t = raw as Record<string, unknown> & { range?: TableOwnBordersTable["range"] };
    const own = normalizeTableOwnBorders(t.ownBorders);
    return Boolean(own?.header) && t.headerRow !== false && t.range !== undefined &&
      rect.r1 === t.range.r1 && rect.r2 === t.range.r1 && rect.c1 === t.range.c1 && rect.c2 === t.range.c2;
  });
  if (index < 0) return null;
  const raw = sheet.dataTables[index] as Record<string, unknown> & { range: TableOwnBordersTable["range"] };
  const own = normalizeTableOwnBorders(raw.ownBorders)!;
  const rest: TableOwnBorders = { ...own };
  delete rest.header;
  const table = { ...raw };
  if (Object.keys(rest).length) table.ownBorders = rest;
  else delete table.ownBorders;
  const tables = [...sheet.dataTables];
  tables[index] = table;
  const cells: Record<string, OwnEdgeCell | undefined> = { ...(sheet.cells as Record<string, OwnEdgeCell | undefined>) };
  for (let col = raw.range.c1; col <= raw.range.c2; col += 1) {
    const key = `${raw.range.r1}:${col}`;
    const stripped = cellWithEdges(cells[key], {}, ["top", "bottom"]);
    if (stripped) cells[key] = stripped;
  }
  return { ...sheet, dataTables: tables, cells: cells as unknown as S["cells"] };
}
