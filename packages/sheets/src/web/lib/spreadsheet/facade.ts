/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The spreadsheet facade — slice 1: a read-only query surface.
 *
 * Everything that wants to READ the grid today — tests, AI tooling, a future
 * scripting console — reaches into the component: test ids, DOM scraping, or
 * imports of private helpers. Each of those couples the consumer to the
 * component's internals, which is exactly what kept breaking them. The facade
 * is the stable seam: a small object of queries whose answers come from the
 * SAME resolvers the grid renders with, so what it reports is what the user
 * sees. The idea — a thin API object decoupled from the implementation — is
 * Univer's facade design (Apache-2.0, dream-num/univer); the code is written
 * against this component's own accessors and reads none of theirs.
 *
 * Slice 2 wires mutation THROUGH the command registry: execute(id) runs the
 * same handler the keyboard runs, so a scripted fillDown and a Ctrl+D are one
 * code path — same undo step, same protection guard. The facade still never
 * touches state itself; it forwards to the registry and reports whether a
 * handler ran. Commands whose handlers need an input event (paste) simply
 * no-op safely without one.
 *
 * The host attaches the facade to its ROOT ELEMENT as a non-enumerable
 * `__mrai_spreadsheet_facade` property. That placement is the point: a
 * consumer holds a DOM node it already has (each window is one element), never
 * a global — two spreadsheet windows have two facades, and a facade dies with
 * its element.
 */

export interface SpreadsheetSelectionRect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export interface SpreadsheetFacade {
  /** Schema version of this surface; bump on breaking changes only. */
  readonly version: 1;
  /** The active sheet's display name. */
  sheetName(): string;
  /** How many sheets the workbook holds (hidden ones included). */
  sheetCount(): number;
  /** The active cell (0-based). */
  activeCell(): { row: number; col: number };
  /** The selection rectangle, normalized (r1<=r2, c1<=c2; 0-based). */
  selection(): SpreadsheetSelectionRect;
  /** A cell's raw stored text ("" for blank): formula text for formulas. */
  cellRaw(row: number, col: number): string;
  /** A cell's DISPLAY string — computed, formatted, exactly as rendered. */
  cellDisplay(row: number, col: number): string;
  /** Display strings for a rectangle, row-major. Bounded by the caller. */
  rangeDisplay(rect: SpreadsheetSelectionRect): string[][];
  /** The command ids currently executable through this window. */
  commands(): readonly string[];
  /**
   * Run a registry command against the CURRENT selection — the same handler,
   * undo step and protection guard the keyboard path uses. Returns whether a
   * handler ran. No payload: event-carrying commands no-op safely.
   */
  execute(command: string): boolean;
  /**
   * Write one cell through the SAME guarded commit path typing uses — sheet
   * protection refuses it with the user-visible notice, it lands as one undo
   * step, and autosave picks it up. Returns whether the command ran (false
   * for out-of-range coordinates or a non-string value, checked here so a
   * scripting mistake cannot reach the handler malformed).
   */
  setCell(row: number, col: number, value: string): boolean;
  /**
   * Write a rectangular block anchored at (row, col) through the shared
   * paste-grid commit path — ONE undo step, protection guard, autosave.
   * `values` is row-major; every entry must be a string; the block is bounded
   * by FACADE_MAX_RANGE_CELLS. Malformed input reports false here, before it
   * can reach the handler.
   */
  setRange(row: number, col: number, values: readonly (readonly string[])[]): boolean;
}

/** What the host must provide; every function must read CURRENT state. */
export interface SpreadsheetFacadeAccessors {
  sheetName(): string;
  sheetCount(): number;
  activeCell(): { row: number; col: number };
  selection(): SpreadsheetSelectionRect;
  cellRaw(row: number, col: number): string;
  cellDisplay(row: number, col: number): string;
  commands(): readonly string[];
  execute(command: string): boolean;
  setCell(row: number, col: number, value: string): boolean;
  setRange(
    row: number,
    col: number,
    values: readonly (readonly string[])[],
  ): boolean;
}

/** A range wider than this refuses rather than assembling a giant array. */
export const FACADE_MAX_RANGE_CELLS = 100_000;

export function createSpreadsheetFacade(
  accessors: SpreadsheetFacadeAccessors,
): SpreadsheetFacade {
  return {
    version: 1,
    sheetName: () => accessors.sheetName(),
    sheetCount: () => accessors.sheetCount(),
    activeCell: () => accessors.activeCell(),
    selection: () => {
      const raw = accessors.selection();
      return {
        r1: Math.min(raw.r1, raw.r2),
        r2: Math.max(raw.r1, raw.r2),
        c1: Math.min(raw.c1, raw.c2),
        c2: Math.max(raw.c1, raw.c2),
      };
    },
    cellRaw: (row, col) => accessors.cellRaw(row, col),
    cellDisplay: (row, col) => accessors.cellDisplay(row, col),
    commands: () => accessors.commands(),
    execute: (command) => accessors.execute(command),
    setCell: (row, col, value) => {
      if (
        !Number.isInteger(row) ||
        !Number.isInteger(col) ||
        row < 0 ||
        col < 0 ||
        typeof value !== "string"
      ) {
        return false;
      }
      return accessors.setCell(row, col, value);
    },
    setRange: (row, col, values) => {
      if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0)
        return false;
      if (!Array.isArray(values) || values.length === 0) return false;
      let cells = 0;
      for (const line of values) {
        if (!Array.isArray(line) || line.length === 0) return false;
        cells += line.length;
        if (cells > FACADE_MAX_RANGE_CELLS) return false;
        for (const value of line) {
          if (typeof value !== "string") return false;
        }
      }
      return accessors.setRange(row, col, values);
    },
    rangeDisplay: (rect) => {
      const r1 = Math.min(rect.r1, rect.r2);
      const r2 = Math.max(rect.r1, rect.r2);
      const c1 = Math.min(rect.c1, rect.c2);
      const c2 = Math.max(rect.c1, rect.c2);
      if ((r2 - r1 + 1) * (c2 - c1 + 1) > FACADE_MAX_RANGE_CELLS) {
        throw new Error(
          `range exceeds ${FACADE_MAX_RANGE_CELLS} cells; narrow it`,
        );
      }
      const rows: string[][] = [];
      for (let r = r1; r <= r2; r += 1) {
        const cells: string[] = [];
        for (let c = c1; c <= c2; c += 1) {
          cells.push(accessors.cellDisplay(r, c));
        }
        rows.push(cells);
      }
      return rows;
    },
  };
}

export const SPREADSHEET_FACADE_PROPERTY = "__mrai_spreadsheet_facade";

/** Attach non-enumerably so serializers and for-in consumers never see it.
 * The detach it returns takes the facade off again: an element outlives its
 * window wherever something still holds it, and the facade's accessors are
 * the whole window. */
export function attachSpreadsheetFacade(
  element: Element,
  facade: SpreadsheetFacade,
): () => void {
  Object.defineProperty(element, SPREADSHEET_FACADE_PROPERTY, {
    value: facade,
    configurable: true,
    enumerable: false,
    writable: false,
  });
  return () => {
    const own = Object.getOwnPropertyDescriptor(element, SPREADSHEET_FACADE_PROPERTY);
    if (own?.value === facade) delete (element as unknown as Record<string, unknown>)[SPREADSHEET_FACADE_PROPERTY];
  };
}
