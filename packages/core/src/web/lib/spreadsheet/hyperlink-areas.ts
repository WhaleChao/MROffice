/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
type Rect = { r1: number; c1: number; r2: number; c2: number };

export interface SpreadsheetHyperlinkArea {
  range: Rect;
  hyperlink: string;
  hyperlinkTooltip?: string;
  hyperlinkDisplay?: string;
  /** Its place in the file as the cells' `hyperlinkRank` counts it: the
   *  cells' elements the file lists before it (reading 981). */
  hyperlinkRank?: number;
}

/** The link fields a place's cell may state. */
interface LinkedCell {
  hyperlink?: string;
  hyperlinkTooltip?: string;
  hyperlinkDisplay?: string;
  /** The cell's own link is listed after the area's over it (reading 950). */
  hyperlinkUnderArea?: boolean;
}

/** The link a place follows: its address and ScreenTip, and the area when it is an area's. */
export interface PlaceLink {
  hyperlink: string;
  hyperlinkTooltip?: string;
  area?: SpreadsheetHyperlinkArea;
}

function usable(area: SpreadsheetHyperlinkArea | null | undefined): area is SpreadsheetHyperlinkArea {
  const range = area?.range;
  return Boolean(
    area &&
      range &&
      Number.isInteger(range.r1) &&
      Number.isInteger(range.c1) &&
      Number.isInteger(range.r2) &&
      Number.isInteger(range.c2) &&
      typeof area.hyperlink === "string" &&
      area.hyperlink,
  );
}

function listed(areas: unknown): (SpreadsheetHyperlinkArea | null | undefined)[] {
  return Array.isArray(areas) ? (areas as (SpreadsheetHyperlinkArea | null | undefined)[]) : [];
}

/**
 * The first of a sheet's `hyperlinkAreas` over a place: the link the place
 * follows when its cell states none of its own (Excel follows the first
 * `<hyperlink>` over a cell).
 */
export function hyperlinkAreaAt(
  areas: unknown,
  row: number,
  col: number,
): SpreadsheetHyperlinkArea | undefined {
  for (const area of listed(areas)) {
    if (!usable(area)) continue;
    const range = area.range;
    if (row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2) return area;
  }
  return undefined;
}

/**
 * The link a place follows: the first the file lists over it. A cell's own
 * link is that one unless the file lists it after the area's
 * (`hyperlinkUnderArea`); a cell with none follows the area's.
 */
export function linkOverPlace(
  place: object | null | undefined,
  areas: unknown,
  row: number,
  col: number,
): PlaceLink | undefined {
  const cell = (place ?? undefined) as LinkedCell | undefined;
  const area = hyperlinkAreaAt(areas, row, col);
  if (cell?.hyperlink && !(area && cell.hyperlinkUnderArea === true)) {
    return {
      hyperlink: cell.hyperlink,
      ...(cell.hyperlinkTooltip ? { hyperlinkTooltip: cell.hyperlinkTooltip } : {}),
    };
  }
  if (!area) return undefined;
  return {
    hyperlink: area.hyperlink,
    ...(area.hyperlinkTooltip ? { hyperlinkTooltip: area.hyperlinkTooltip } : {}),
    area,
  };
}

/** The link a copied cell follows when it is an area's, and the one beneath it. */
export interface AreaLinkSource {
  /** The sheet the area is on. */
  sheetId: string;
  /** A large area's range (`hyperlinkAreas`). */
  range?: Rect;
  /** A small area's place in the file (laid on its cells, `hyperlinkRank`). */
  rank?: number;
  /** The link the place follows once that area's link is gone: its own, if any. */
  beneath?: { hyperlink: string; hyperlinkTooltip?: string; hyperlinkDisplay?: string };
}

/**
 * Where a copied place's link comes from when it is an area's (large or laid),
 * with the cell's own link beneath it. A cell pasted inside that same area
 * deletes the area's link first, and then takes what its source follows
 * without it: its own, or none (Windows Excel, xlsx-rounds probe/n39-linkmore
 * x6, CRITERIA XC-206 C 11: B5 -> C5 over A1:C8 pasted own-after, not the
 * area's). Undefined when the place follows its own link or none.
 */
export function areaLinkSourceOf(
  sheetId: string,
  place: object | null | undefined,
  areas: unknown,
  row: number,
  col: number,
): AreaLinkSource | undefined {
  const cell = (place ?? undefined) as LaidLinkCell | undefined;
  const layer = (link: { hyperlink?: string; hyperlinkTooltip?: string; hyperlinkDisplay?: string } | undefined) =>
    link && typeof link.hyperlink === "string" && link.hyperlink
      ? {
          beneath: {
            hyperlink: link.hyperlink,
            ...(link.hyperlinkTooltip ? { hyperlinkTooltip: link.hyperlinkTooltip } : {}),
            ...(link.hyperlinkDisplay ? { hyperlinkDisplay: link.hyperlinkDisplay } : {}),
          },
        }
      : {};
  if (cell?.hyperlinkInArea === true && typeof cell.hyperlinkRank === "number" && Number.isInteger(cell.hyperlinkRank)) {
    return { sheetId, rank: cell.hyperlinkRank, ...layer(cell.hyperlinkOwnUnder) };
  }
  const link = linkOverPlace(place, areas, row, col);
  if (!link?.area) return undefined;
  return { sheetId, range: { ...link.area.range }, ...layer(cell) };
}

/**
 * The link a place prints with (xlsx-rounds CRITERIA XC-204): Windows
 * Excel's Save As PDF writes a link annotation over a linked cell's text, to
 * the LAST link the file lists over the cell -- not the first, which the
 * window follows (`linkOverPlace`). The base workbook of probe/n35-p948/a
 * lists B3's own link, then A1:C8's, then B5's own: its PDF has B3 at the
 * area's address and B5 at its own (pa2, pa7); a link added last over B2
 * prints (pa5). So: the layer beneath the one the cell shows (the small
 * area's under its own link, or its own under the area's), a large area's
 * unless the file lists the cell's own after it (`hyperlinkUnderArea`), and
 * of several large areas the last.
 */
export function printedLinkOverPlace(
  place: object | null | undefined,
  areas: unknown,
  row: number,
  col: number,
): { hyperlink: string; hyperlinkTooltip?: string } | undefined {
  const cell = (place ?? undefined) as LaidLinkCell | undefined;
  let own: { hyperlink: string; hyperlinkTooltip?: string } | undefined;
  if (cell?.hyperlink) {
    own = { hyperlink: cell.hyperlink, ...(cell.hyperlinkTooltip ? { hyperlinkTooltip: cell.hyperlinkTooltip } : {}) };
    const under = cell.hyperlinkInArea === true ? cell.hyperlinkOwnUnder : cell.hyperlinkArea;
    const listedFirst =
      typeof cell.hyperlinkRank === "number" && typeof under?.hyperlinkRank === "number" && under.hyperlinkRank < cell.hyperlinkRank;
    if (under?.hyperlink && !listedFirst) {
      own = { hyperlink: under.hyperlink, ...(under.hyperlinkTooltip ? { hyperlinkTooltip: under.hyperlinkTooltip } : {}) };
    }
  }
  let over: SpreadsheetHyperlinkArea | undefined;
  for (const area of listed(areas)) {
    if (!usable(area)) continue;
    const range = area.range;
    if (row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2) over = area;
  }
  if (own && (!over || cell?.hyperlinkUnderArea === true)) return own;
  if (over) return { hyperlink: over.hyperlink, ...(over.hyperlinkTooltip ? { hyperlinkTooltip: over.hyperlinkTooltip } : {}) };
  return own;
}

function withAreas<S extends object>(sheet: S, kept: SpreadsheetHyperlinkArea[]): S {
  const next: S = { ...sheet };
  const held = next as { hyperlinkAreas?: unknown };
  if (kept.length > 0) held.hyperlinkAreas = kept;
  else delete held.hyperlinkAreas;
  return next;
}

function without<S extends object>(sheet: S, gone: (range: Rect) => boolean): S {
  const areas = listed((sheet as { hyperlinkAreas?: unknown }).hyperlinkAreas);
  if (areas.length === 0) return sheet;
  const kept = areas.filter((area): area is SpreadsheetHyperlinkArea => usable(area) && !gone(area.range));
  return kept.length === areas.length ? sheet : withAreas(sheet, kept);
}

/**
 * The sheet without the area links that touch `range`: Remove Hyperlink and
 * Clear Hyperlinks delete every link that meets the selection, an area's
 * whole. The same sheet when none does.
 */
export function sheetWithoutHyperlinkAreasOver<S extends object>(sheet: S, range: Rect): S {
  return without(
    sheet,
    (area) => area.r1 <= range.r2 && area.r2 >= range.r1 && area.c1 <= range.c2 && area.c2 >= range.c1,
  );
}

/**
 * The sheet without the area links that hold all of `range`: one cell pasted
 * inside an area's link, or a merge made inside it, deletes that whole link;
 * a paste or a merge reaching past the area leaves it. The same sheet when
 * none does.
 */
export function sheetWithoutHyperlinkAreasAround<S extends object>(sheet: S, range: Rect): S {
  return without(
    sheet,
    (area) => area.r1 <= range.r1 && area.r2 >= range.r2 && area.c1 <= range.c1 && area.c2 >= range.c2,
  );
}

/**
 * The sheet without the area links Clear All and Clear Contents delete
 * whole: those holding all of `range` and those inside it
 * (`laidAreaLinksClearedBy`). The same sheet when none is.
 */
export function sheetWithoutHyperlinkAreasClearedBy<S extends object>(sheet: S, range: Rect): S {
  return without(
    sheet,
    (area) =>
      (area.r1 <= range.r1 && area.r2 >= range.r2 && area.c1 <= range.c1 && area.c2 >= range.c2) ||
      (area.r1 >= range.r1 && area.r2 <= range.r2 && area.c1 >= range.c1 && area.c2 <= range.c2),
  );
}

/**
 * The sheet with the first area link over a place given a new address and
 * ScreenTip (Edit Hyperlink on a cell the area's link is first over): the
 * whole area's, its range and its place in the list unchanged. A new address
 * drops the display text that described the old one. The same sheet when no
 * area is over the place.
 */
export function sheetWithHyperlinkAreaEdited<S extends object>(
  sheet: S,
  row: number,
  col: number,
  href: string,
  tooltip: string | undefined,
): S {
  const areas = listed((sheet as { hyperlinkAreas?: unknown }).hyperlinkAreas);
  const target = hyperlinkAreaAt(areas, row, col);
  if (!target) return sheet;
  const edited: SpreadsheetHyperlinkArea = { ...target, hyperlink: href };
  if (target.hyperlink !== href) {
    delete edited.hyperlinkTooltip;
    delete edited.hyperlinkDisplay;
  }
  if (tooltip !== undefined) {
    if (tooltip) edited.hyperlinkTooltip = tooltip;
    else delete edited.hyperlinkTooltip;
  }
  return withAreas(
    sheet,
    areas.filter(usable).map((area) => (area === target ? edited : area)),
  );
}

/** The link fields of a cell under a link laid over a small area. */
interface LaidLinkCell extends LinkedCell {
  /** The link the cell shows is one over an area, not its own. */
  hyperlinkInArea?: boolean;
  /** The place in the file of the link the cell shows. */
  hyperlinkRank?: number;
  /** The area's link beneath the cell's own, which the file lists first. */
  hyperlinkArea?: { hyperlink?: string; hyperlinkTooltip?: string; hyperlinkDisplay?: string; hyperlinkRank?: number };
  /** The cell's own link beneath the area's, which the file lists first. */
  hyperlinkOwnUnder?: { hyperlink?: string; hyperlinkTooltip?: string; hyperlinkDisplay?: string; hyperlinkRank?: number };
}

type LaidCells = Readonly<Record<string, object | undefined>>;

/** The file place of the laid area link over a cell: the one it shows, or the one beneath its own. */
function laidAreaRank(cell: object | null | undefined): number | undefined {
  const linked = (cell ?? undefined) as LaidLinkCell | undefined;
  if (!linked) return undefined;
  if (linked.hyperlinkInArea === true) {
    return typeof linked.hyperlinkRank === "number" && Number.isInteger(linked.hyperlinkRank) ? linked.hyperlinkRank : undefined;
  }
  const beneath = linked.hyperlinkArea?.hyperlinkRank;
  return linked.hyperlink && typeof beneath === "number" && Number.isInteger(beneath) ? beneath : undefined;
}

function placeOf(key: string): { row: number; col: number } {
  const separator = key.indexOf(":");
  return { row: Number(key.slice(0, separator)), col: Number(key.slice(separator + 1)) };
}

/**
 * The laid area links that touch `range`, by their places in the file:
 * every one a cell of the range shows or holds beneath its own.
 */
export function laidAreaLinksOver(cells: LaidCells | null | undefined, range: Rect): ReadonlySet<number> {
  const ranks = new Set<number>();
  if (!cells) return ranks;
  const places = (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1);
  if (places <= 4_096) {
    for (let row = range.r1; row <= range.r2; row += 1) {
      for (let col = range.c1; col <= range.c2; col += 1) {
        const rank = laidAreaRank(cells[`${row}:${col}`]);
        if (rank !== undefined) ranks.add(rank);
      }
    }
    return ranks;
  }
  for (const key of Object.keys(cells)) {
    const rank = laidAreaRank(cells[key]);
    if (rank === undefined || ranks.has(rank)) continue;
    const { row, col } = placeOf(key);
    if (row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2) ranks.add(rank);
  }
  return ranks;
}

/** The box of the cells that hold each laid area link `touching` names. */
function laidAreaBoxes(cells: LaidCells, touching: ReadonlySet<number>): Map<number, Rect> {
  const boxes = new Map<number, Rect>();
  for (const key of Object.keys(cells)) {
    const rank = laidAreaRank(cells[key]);
    if (rank === undefined || !touching.has(rank)) continue;
    const { row, col } = placeOf(key);
    const box = boxes.get(rank);
    if (!box) boxes.set(rank, { r1: row, c1: col, r2: row, c2: col });
    else {
      box.r1 = Math.min(box.r1, row);
      box.c1 = Math.min(box.c1, col);
      box.r2 = Math.max(box.r2, row);
      box.c2 = Math.max(box.c2, col);
    }
  }
  return boxes;
}

/**
 * The laid area links that hold all of `range`: those touching it whose
 * cells reach every side of it.
 */
export function laidAreaLinksAround(cells: LaidCells | null | undefined, range: Rect): ReadonlySet<number> {
  const touching = laidAreaLinksOver(cells, range);
  if (!cells || touching.size === 0) return touching;
  const around = new Set<number>();
  for (const [rank, box] of laidAreaBoxes(cells, touching)) {
    if (box.r1 <= range.r1 && box.r2 >= range.r2 && box.c1 <= range.c1 && box.c2 >= range.c2) around.add(rank);
  }
  return around;
}

/**
 * The laid area links Clear All and Clear Contents delete whole: those that
 * hold all of `range` and those `range` holds all of; one reaching past the
 * selection on some side and not holding it stays (Windows Excel,
 * xlsx-rounds probe/n48-clearlinks: Clear Contents A1:D9 over A1:C8 deletes
 * it, B7:B10 and B3:D4 across its edge leave it; CRITERIA XC-209).
 */
export function laidAreaLinksClearedBy(cells: LaidCells | null | undefined, range: Rect): ReadonlySet<number> {
  const touching = laidAreaLinksOver(cells, range);
  if (!cells || touching.size === 0) return touching;
  const cleared = new Set<number>();
  for (const [rank, box] of laidAreaBoxes(cells, touching)) {
    const around = box.r1 <= range.r1 && box.r2 >= range.r2 && box.c1 <= range.c1 && box.c2 >= range.c2;
    const within = box.r1 >= range.r1 && box.r2 <= range.r2 && box.c1 >= range.c1 && box.c2 <= range.c2;
    if (around || within) cleared.add(rank);
  }
  return cleared;
}

const LAID_LINK_KEYS = [
  "hyperlink",
  "hyperlinkTooltip",
  "hyperlinkDisplay",
  "hyperlinkInArea",
  "hyperlinkRank",
  "hyperlinkOwnUnder",
  "hyperlinkArea",
] as const;

/**
 * `cells` without the laid area links `ranks` names, each link whole: a
 * cell that showed one shows its own link the file lists after it, or none
 * (an empty cell that held only the link leaves the sheet); a cell whose own
 * link stood before one keeps its own alone. Inside `within`, the cells the
 * command acted on: "unlinked" leaves them no link at all (Remove Hyperlink
 * and the Clears delete every link that meets them), "own" leaves what they
 * hold as their own (a paste's cells). The same map when no cell changes.
 */
export function cellsWithoutLaidAreaLinks<C extends LaidCells>(
  cells: C,
  ranks: ReadonlySet<number>,
  within?: Rect,
  inside: "unlinked" | "own" = "unlinked",
): C {
  if (ranks.size === 0) return cells;
  let next: Record<string, object | undefined> | null = null;
  for (const key of Object.keys(cells)) {
    const cell = cells[key] as LaidLinkCell | undefined;
    if (!cell) continue;
    const { row, col } = placeOf(key);
    const acted = Boolean(within && row >= within.r1 && row <= within.r2 && col >= within.c1 && col <= within.c2);
    const shows = cell.hyperlinkInArea === true && ranks.has(cell.hyperlinkRank ?? Number.NaN);
    const beneath = !shows && cell.hyperlinkArea !== undefined && ranks.has(cell.hyperlinkArea.hyperlinkRank ?? Number.NaN);
    const stale = acted && inside === "unlinked" && LAID_LINK_KEYS.some((name) => name in cell);
    if (!shows && !beneath && !stale) continue;
    const plain: Record<string, unknown> = { ...cell };
    if (stale) {
      for (const name of LAID_LINK_KEYS) delete plain[name];
    } else if (beneath) {
      delete plain.hyperlinkArea;
    } else if (acted) {
      // The command's own cell: the link it holds is its own now.
      delete plain.hyperlinkInArea;
      delete plain.hyperlinkRank;
      delete plain.hyperlinkOwnUnder;
      delete plain.hyperlinkArea;
    } else {
      const own = cell.hyperlinkOwnUnder;
      for (const name of LAID_LINK_KEYS) delete plain[name];
      if (own && typeof own.hyperlink === "string" && own.hyperlink) {
        plain.hyperlink = own.hyperlink;
        if (own.hyperlinkTooltip) plain.hyperlinkTooltip = own.hyperlinkTooltip;
        if (own.hyperlinkDisplay) plain.hyperlinkDisplay = own.hyperlinkDisplay;
        if (typeof own.hyperlinkRank === "number") plain.hyperlinkRank = own.hyperlinkRank;
      }
    }
    next ??= { ...cells };
    const left = Object.keys(plain);
    const empty =
      !plain.hyperlink &&
      (plain.value === "" || plain.value === undefined || plain.value === null) &&
      left.every((name) => name === "value" || name === "kind");
    if (empty) delete next[key];
    else next[key] = plain;
  }
  return (next ?? cells) as C;
}

/**
 * `cells` (after a Clear All or Clear Contents over `range`, the links of
 * `ranks` deleted by `cellsWithoutLaidAreaLinks`) with the range's cells
 * holding what Windows Excel leaves them: their own links go with their
 * contents, and a laid area link the clear did not delete (one reaching past
 * the selection) stays on them -- shown where it was shown, and shown now
 * where it lay beneath the cell's own link (xlsx-rounds probe/n48-clearlinks
 * pc3: Clear Contents B3:D4 deletes B3's own link, A1:C8 stays and B3
 * follows it). A cleared cell that showed a deleted link holds none: the
 * one under it was its own (R462-X28 WALK1, Table List A32:B35 over
 * A32:B32 and A32's own link kept hyperlinkOwnUnder and the link's place,
 * and the save wrote the deleted area back). `before` is the sheet's cells
 * before the clear.
 */
export function clearedCellsWithAreaLinksKept<C extends LaidCells>(
  before: LaidCells,
  cells: C,
  ranks: ReadonlySet<number>,
  range: Rect,
): C {
  if (laidAreaLinksOver(before, range).size === 0) return cells;
  let next: Record<string, object | undefined> | null = null;
  for (const key of Object.keys(before)) {
    const was = before[key] as LaidLinkCell | undefined;
    if (!was) continue;
    const { row, col } = placeOf(key);
    if (row < range.r1 || row > range.r2 || col < range.c1 || col > range.c2) continue;
    if (!LAID_LINK_KEYS.some((name) => name in was)) continue;
    const plain: Record<string, unknown> = { ...((cells[key] as object | undefined) ?? { value: "", kind: "text" }) };
    for (const name of LAID_LINK_KEYS) delete plain[name];
    const shown = was.hyperlinkInArea === true ? was.hyperlinkRank : undefined;
    const beneath = was.hyperlink && was.hyperlinkInArea !== true ? was.hyperlinkArea : undefined;
    let kept: { hyperlink?: string; hyperlinkTooltip?: string; hyperlinkDisplay?: string; hyperlinkRank?: number } | undefined;
    if (typeof shown === "number" && Number.isInteger(shown) && !ranks.has(shown)) {
      kept = { hyperlink: was.hyperlink, hyperlinkTooltip: was.hyperlinkTooltip, hyperlinkDisplay: was.hyperlinkDisplay, hyperlinkRank: shown };
    } else if (
      beneath &&
      typeof beneath.hyperlinkRank === "number" &&
      Number.isInteger(beneath.hyperlinkRank) &&
      !ranks.has(beneath.hyperlinkRank)
    ) {
      kept = beneath;
    }
    if (kept && typeof kept.hyperlink === "string" && kept.hyperlink) {
      plain.hyperlink = kept.hyperlink;
      if (kept.hyperlinkTooltip) plain.hyperlinkTooltip = kept.hyperlinkTooltip;
      if (kept.hyperlinkDisplay) plain.hyperlinkDisplay = kept.hyperlinkDisplay;
      plain.hyperlinkInArea = true;
      plain.hyperlinkRank = kept.hyperlinkRank;
    }
    const current = cells[key] as Record<string, unknown> | undefined;
    const left = Object.keys(plain);
    const empty =
      !plain.hyperlink &&
      (plain.value === "" || plain.value === undefined || plain.value === null) &&
      left.every((name) => name === "value" || name === "kind");
    if (empty ? current === undefined : current !== undefined && sameFields(current, plain)) continue;
    next ??= { ...cells };
    if (empty) delete next[key];
    else next[key] = plain;
  }
  return (next ?? cells) as C;
}

function sameFields(a: Record<string, unknown>, b: Record<string, unknown>): boolean {
  const keys = Object.keys(a);
  return keys.length === Object.keys(b).length && keys.every((key) => Object.is(a[key], b[key]));
}

/**
 * `after` without the laid area links that touched `range` on `before`
 * (Remove Hyperlink, Clear Hyperlinks): each whole, the range's cells left
 * with no link. `after` itself when `before` held none there.
 */
export function sheetWithoutLaidAreaLinksOver<S extends object>(before: S, after: S, range: Rect): S {
  const ranks = laidAreaLinksOver((before as { cells?: LaidCells }).cells, range);
  if (ranks.size === 0) return after;
  const cells = (after as { cells?: LaidCells }).cells ?? {};
  const next = cellsWithoutLaidAreaLinks(cells, ranks, range, "unlinked");
  return next === cells ? after : ({ ...after, cells: next } as S);
}

/**
 * `after` without the laid area links that held all of `range` on `before`
 * (a Clear, a paste or a merge inside one): each whole. The range's cells
 * are left with no link ("unlinked", a Clear) or with what the command gave
 * them as their own ("own", a paste). `after` itself when `before` held
 * none around the range.
 */
export function sheetWithoutLaidAreaLinksAround<S extends object>(
  before: S,
  after: S,
  range: Rect,
  inside: "unlinked" | "own" = "unlinked",
): S {
  const ranks = laidAreaLinksAround((before as { cells?: LaidCells }).cells, range);
  if (ranks.size === 0) return after;
  const cells = (after as { cells?: LaidCells }).cells ?? {};
  const next = cellsWithoutLaidAreaLinks(cells, ranks, range, inside);
  return next === cells ? after : ({ ...after, cells: next } as S);
}

/**
 * The sheet with the laid area link a cell shows given a new address and
 * ScreenTip on every cell that shows it (Edit Hyperlink changes the first
 * link over the cell, the whole area's when that is first), and beneath the
 * cells whose own link stands before it. A new address drops the ScreenTip
 * and display text that described the old one. The same sheet when the cell
 * shows no laid area link.
 */
export function sheetWithLaidAreaLinkEdited<S extends object>(
  sheet: S,
  row: number,
  col: number,
  href: string,
  tooltip: string | undefined,
): S {
  const cells = (sheet as { cells?: LaidCells }).cells;
  const target = cells?.[`${row}:${col}`] as LaidLinkCell | undefined;
  if (!cells || !target || target.hyperlinkInArea !== true) return sheet;
  const rank = laidAreaRank(target);
  if (rank === undefined) return sheet;
  const next: Record<string, object | undefined> = { ...cells };
  const edit = (link: Record<string, unknown>): void => {
    if (link.hyperlink !== href) {
      delete link.hyperlinkTooltip;
      delete link.hyperlinkDisplay;
    }
    link.hyperlink = href;
    if (tooltip !== undefined) {
      if (tooltip) link.hyperlinkTooltip = tooltip;
      else delete link.hyperlinkTooltip;
    }
  };
  for (const key of Object.keys(cells)) {
    const cell = cells[key] as LaidLinkCell | undefined;
    if (!cell) continue;
    if (cell.hyperlinkInArea === true && cell.hyperlinkRank === rank) {
      const edited: Record<string, unknown> = { ...cell };
      edit(edited);
      next[key] = edited;
    } else if (cell.hyperlinkArea?.hyperlinkRank === rank) {
      const beneath: Record<string, unknown> = { ...cell.hyperlinkArea };
      edit(beneath);
      next[key] = { ...cell, hyperlinkArea: beneath };
    }
  }
  return { ...sheet, cells: next } as S;
}

const CELL_LINK_KEYS = [
  "hyperlink",
  "hyperlinkTooltip",
  "hyperlinkDisplay",
  "hyperlinkUnderArea",
  "hyperlinkInArea",
  "hyperlinkRank",
  "hyperlinkOwnUnder",
  "hyperlinkArea",
] as const;

/** The most filled cells that take an area's link along, each as its own. */
const FILLED_LINK_CELLS_MAX = 4_096;

/**
 * Whether a cell of `target` is filled from a place under a link laid over a
 * small area: its edge filled into cells that all lie outside it meets no
 * link on what it writes, and each still takes the link its source follows
 * (xlsx-rounds R456-X01 LINKFILLOUT: `A41:E41`'s A41 filled down into A42).
 * A fill reads each cell from its own row or column, so the target's first
 * row and first column name every source.
 */
function filledFromUnderLaidLink(
  read: Record<string, Record<string, unknown> | undefined>,
  target: Rect,
  sourceOf: (row: number, col: number) => { row: number; col: number },
): boolean {
  const laid = (row: number, col: number) => {
    const source = sourceOf(row, col);
    return laidAreaRank(read[`${source.row}:${source.col}`]) !== undefined;
  };
  const rows = Math.min(target.r2, target.r1 + FILLED_LINK_CELLS_MAX - 1);
  const cols = Math.min(target.c2, target.c1 + FILLED_LINK_CELLS_MAX - 1);
  for (let col = target.c1; col <= cols; col += 1) if (laid(target.r1, col)) return true;
  for (let row = target.r1 + 1; row <= rows; row += 1) if (laid(row, target.c1)) return true;
  return false;
}

/**
 * A fill's links where the sheet holds links over large areas, as Windows
 * Excel leaves them (xlsx-rounds probe/n35-p948/a pb1-pb16). `before` is the
 * sheet the fill read, `filled` what it wrote, `target` the cells it filled
 * (the source not among them), `sourceOf` the cell each one was filled from.
 * Filled cells all inside an area's link: that whole link goes and the
 * cells hold no link. Otherwise each cell filled from a place that follows
 * an area's link takes that link as its own; the area stays.
 */
export function sheetWithLinksFilled<S extends object>(
  before: S,
  filled: S,
  target: Rect,
  sourceOf: (row: number, col: number) => { row: number; col: number },
): S {
  const areas = listed((before as { hyperlinkAreas?: unknown }).hyperlinkAreas);
  type Cells = Record<string, Record<string, unknown> | undefined>;
  const read = (before as { cells?: Cells }).cells ?? {};
  // A link laid over a small area that holds every filled cell goes whole
  // too, off the cells outside the fill as well.
  const laidAround = laidAreaLinksAround(read, target);
  const laidOver = laidAround.size > 0 ? laidAround : laidAreaLinksOver(read, target);
  if (areas.length === 0 && laidOver.size === 0 && !filledFromUnderLaidLink(read, target, sourceOf)) return filled;
  const written =
    laidAround.size > 0
      ? cellsWithoutLaidAreaLinks((filled as { cells?: Cells }).cells ?? {}, laidAround, target, "unlinked")
      : ((filled as { cells?: Cells }).cells ?? {});
  const next: Cells = { ...written };
  let changed = laidAround.size > 0;
  const unlinked = (cell: Record<string, unknown>): Record<string, unknown> => {
    const plain = { ...cell };
    for (const key of CELL_LINK_KEYS) delete plain[key];
    return plain;
  };
  const inside = sheetWithoutHyperlinkAreasAround(filled, target);
  if (inside !== filled || laidAround.size > 0) {
    for (const key of Object.keys(next)) {
      const cell = next[key];
      if (!cell?.hyperlink) continue;
      const separator = key.indexOf(":");
      const row = Number(key.slice(0, separator));
      const col = Number(key.slice(separator + 1));
      if (row < target.r1 || row > target.r2 || col < target.c1 || col > target.c2) continue;
      next[key] = unlinked(cell);
      changed = true;
    }
    return changed ? ({ ...inside, cells: next } as S) : inside;
  }
  if ((target.r2 - target.r1 + 1) * (target.c2 - target.c1 + 1) > FILLED_LINK_CELLS_MAX) return filled;
  for (let row = target.r1; row <= target.r2; row += 1) {
    for (let col = target.c1; col <= target.c2; col += 1) {
      const source = sourceOf(row, col);
      const from = read[`${source.row}:${source.col}`];
      const over = linkOverPlace(from, areas, source.row, source.col);
      const key = `${row}:${col}`;
      if (!over?.area) {
        // Filled from under a link laid over a small area, or over one: the
        // filled cell holds the link its source follows as its own, no part
        // of the area the source lies in.
        const copy = next[key];
        if (copy && from?.hyperlink && (laidAreaRank(from) !== undefined || laidAreaRank(copy) !== undefined)) {
          const own = unlinked(copy);
          own.hyperlink = from.hyperlink;
          if (from.hyperlinkTooltip) own.hyperlinkTooltip = from.hyperlinkTooltip;
          next[key] = own;
          changed = true;
        }
        continue;
      }
      const cell = unlinked(next[key] ?? { value: "", kind: "text" });
      cell.hyperlink = over.hyperlink;
      if (over.hyperlinkTooltip) cell.hyperlinkTooltip = over.hyperlinkTooltip;
      next[key] = cell;
      changed = true;
    }
  }
  return changed ? ({ ...filled, cells: next } as S) : filled;
}
