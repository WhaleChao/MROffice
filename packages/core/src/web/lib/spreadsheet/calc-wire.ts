/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The spreadsheet calculation session's wire format (engine `sheet-calc`,
 * DESIGN-WAVE2 section 5). Browser-safe: no evaluator, only how the values
 * and coordinates the Web Worker used to pass by structured clone travel as
 * JSON.
 *
 * Values: a formula result is a plain JSON-like object except where JSON
 * loses something -- a `Map` (a LAMBDA's closure, a seed's snapshot), a
 * `Set` (a seed's Manual-mode entries), a non-finite number and negative
 * zero. Those travel tagged (`{"$mrai":"map","entries":[...]}`,
 * `{"$mrai":"set","items":[...]}`, `{"$mrai":"num","v":"NaN"}`) and come
 * back as they were. No engine value carries a `$mrai` key of its own.
 *
 * Coordinates: one string per sheet, `"row:col row:col ..."`, so a request
 * naming 100,000 formula cells is one short string per sheet rather than
 * 100,000 JSON objects; `{"all": true}` names every formula of the session.
 */

import type { WorksheetFormulaCoordinate } from "./calculation-mode";

const TAG = "$mrai";

type Wire = null | boolean | number | string | Wire[] | { [key: string]: Wire };

export function encodeCalcValue(value: unknown): Wire {
  if (value === null || value === undefined) return null;
  if (typeof value === "number") {
    if (Number.isNaN(value)) return { [TAG]: "num", v: "NaN" };
    if (value === Infinity) return { [TAG]: "num", v: "Infinity" };
    if (value === -Infinity) return { [TAG]: "num", v: "-Infinity" };
    if (Object.is(value, -0)) return { [TAG]: "num", v: "-0" };
    return value;
  }
  if (typeof value === "string" || typeof value === "boolean") return value;
  if (Array.isArray(value)) return value.map((item) => encodeCalcValue(item));
  if (value instanceof Map) {
    return { [TAG]: "map", entries: [...value].map(([key, item]) => [encodeCalcValue(key), encodeCalcValue(item)]) };
  }
  if (value instanceof Set) return { [TAG]: "set", items: [...value].map((item) => encodeCalcValue(item)) };
  if (typeof value === "object") {
    const out: { [key: string]: Wire } = {};
    for (const [key, item] of Object.entries(value as Record<string, unknown>)) {
      if (item === undefined || typeof item === "function") continue;
      out[key] = encodeCalcValue(item);
    }
    return out;
  }
  return null;
}

export function decodeCalcValue(value: unknown): unknown {
  if (value === null || typeof value !== "object") return value;
  if (Array.isArray(value)) return value.map((item) => decodeCalcValue(item));
  const record = value as Record<string, unknown>;
  if (record[TAG] === "num") {
    const v = record.v;
    return v === "NaN" ? NaN : v === "Infinity" ? Infinity : v === "-Infinity" ? -Infinity : -0;
  }
  if (record[TAG] === "map" && Array.isArray(record.entries)) {
    return new Map((record.entries as unknown[][]).map(([key, item]) => [decodeCalcValue(key), decodeCalcValue(item)]));
  }
  if (record[TAG] === "set" && Array.isArray(record.items)) return new Set(record.items.map((item) => decodeCalcValue(item)));
  const out: Record<string, unknown> = {};
  for (const [key, item] of Object.entries(record)) out[key] = decodeCalcValue(item);
  return out;
}

/** A stable spelling of a value, for "did it change since the client last saw it". */
export function calcValueFingerprint(value: unknown): string {
  return JSON.stringify(encodeCalcValue(value));
}

export type CalcCoordinates = { readonly all: true } | { readonly sheets: ReadonlyArray<readonly [string, string]> };

/** `{sheets: [[sheetId, "r:c r:c"]]}`, in the request's own order per sheet. */
export function encodeCalcCoordinates(coordinates: readonly WorksheetFormulaCoordinate[]): CalcCoordinates {
  const bySheet = new Map<string, string[]>();
  for (const cell of coordinates) {
    let list = bySheet.get(cell.sheetId);
    if (!list) bySheet.set(cell.sheetId, (list = []));
    list.push(`${cell.row}:${cell.col}`);
  }
  return { sheets: [...bySheet].map(([sheetId, cells]) => [sheetId, cells.join(" ")] as const) };
}

/** The coordinates back, given each sheet id's name; null when malformed. */
export function decodeCalcCoordinates(
  encoded: CalcCoordinates,
  sheetName: (sheetId: string) => string | undefined,
  max: number,
): WorksheetFormulaCoordinate[] | null {
  if (!("sheets" in encoded) || !Array.isArray(encoded.sheets)) return null;
  const out: WorksheetFormulaCoordinate[] = [];
  for (const entry of encoded.sheets) {
    if (!Array.isArray(entry) || typeof entry[0] !== "string" || typeof entry[1] !== "string") return null;
    const name = sheetName(entry[0]);
    if (name === undefined) return null;
    if (!entry[1]) continue;
    for (const pair of entry[1].split(" ")) {
      const match = /^(\d{1,7}):(\d{1,5})$/.exec(pair);
      if (!match) return null;
      out.push({ sheetId: entry[0], sheetName: name, row: Number(match[1]), col: Number(match[2]) });
      if (out.length > max) return null;
    }
  }
  return out;
}
