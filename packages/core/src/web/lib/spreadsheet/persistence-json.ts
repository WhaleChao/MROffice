/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Cell objects a save has walked, so the next save of 30,000 cells need not (about 20 ms per entry). */
const cellsFoundSound = new WeakSet<object>();

/** Optional editor properties may be undefined in memory. Omit those object
 * properties before either REST serialization or in-memory Workbook V2 sees
 * the document. Preserve every defined value; malformed values must not be
 * silently converted to null (as JSON.stringify would do for NaN/arrays).
 * Copy only branches that need omission, keeping immutable model sharing.
 */
export function spreadsheetPersistenceJson<T>(value: T): T {
  const ancestors = new WeakSet<object>();
  // The path being visited, so a refusal says which value it met.
  const path: (string | number)[] = [];
  const visit = (input: unknown): unknown => {
    if (input === null || typeof input === "string" || typeof input === "boolean") return input;
    if (typeof input === "number" && Number.isFinite(input)) return input;
    if (typeof input !== "object" || input === null) {
      throw new Error(`Spreadsheet content contains a non-JSON value at ${path.join(".") || "(root)"}`);
    }
    if (ancestors.has(input)) throw new Error("Spreadsheet content contains a cycle");
    const array = Array.isArray(input);
    const prototype = Object.getPrototypeOf(input);
    if (!array && prototype !== Object.prototype && prototype !== null) throw new Error("Spreadsheet content contains a non-JSON object");
    ancestors.add(input);
    try {
      if (array) {
        let result: unknown[] | undefined;
        for (let index = 0; index < input.length; index += 1) {
          path.push(index);
          const child = visit(input[index]);
          path.pop();
          if (child !== input[index]) {
            result ??= input.slice();
            result[index] = child;
          }
        }
        return result ?? input;
      }
      const object = input as Record<string, unknown>;
      let result: Record<string, unknown> | undefined;
      const aMapOfCells = path[path.length - 1] === "cells";
      for (const key of Object.keys(object)) {
        const child = object[key];
        if (child === undefined) {
          result ??= { ...object };
          delete result[key];
        } else if (aMapOfCells && cellsFoundSound.has(child as object)) {
          continue;
        } else {
          path.push(key);
          const next = visit(child);
          path.pop();
          // A cell is never changed once made: found sound, it stays so.
          if (aMapOfCells && next === child && typeof child === "object") cellsFoundSound.add(child as object);
          if (next !== child) {
            result ??= { ...object };
            result[key] = next;
          }
        }
      }
      return result ?? input;
    } finally {
      ancestors.delete(input);
    }
  };
  return visit(value) as T;
}
