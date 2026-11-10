/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { WorkbookSheetInput } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import type { BackgroundCalculationHostBinding, BackgroundCalculationHostSheet } from "@mrnote/core/web/lib/spreadsheet/background-calculation-host.ts";
import type { BackgroundCalculationSheetContext } from "@mrnote/core/web/lib/spreadsheet/background-calculation-protocol.ts";

type InputEntry = { role: string; keys: readonly unknown[]; input: WorkbookSheetInput };
type SourceSheet = BackgroundCalculationHostSheet & { readonly merges?: unknown };
const sameKeys = (a: readonly unknown[], b: readonly unknown[]) => a.length === b.length && a.every((item, index) => Object.is(item, b[index]));

const NO_MERGES = {};
const semanticKeys = new WeakMap<object, WeakMap<object, object>>();
/**
 * What, beside the cells, decides how a sheet's raw cells read (a binding's
 * `rawSemanticsKey`): its merge mask, and its What-If Data Tables, whose
 * results read as =TABLE(...). One stable identity per pair, so a change of
 * either resets the calculation session.
 */
export function calculationRawSemanticsKey(sheet: { readonly merges?: unknown; readonly whatIfTables?: unknown }): unknown {
  const tables = sheet.whatIfTables;
  if (!tables || typeof tables !== "object") return sheet.merges;
  const merges = sheet.merges && typeof sheet.merges === "object" ? sheet.merges : NO_MERGES;
  let byTables = semanticKeys.get(merges);
  if (!byTables) semanticKeys.set(merges, (byTables = new WeakMap()));
  let key = byTables.get(tables);
  if (!key) byTables.set(tables, (key = {}));
  return key;
}

/** Per-editor exact immutable inputs. Two active/nonactive variants per sheet,
 * two ordered workbook contexts; no formula values or cross-session authority. */
export class SpreadsheetFormulaInputCache {
  #inputs = new Map<string, InputEntry[]>();
  #inputVectors: WorkbookSheetInput[][] = [];
  #vectors: BackgroundCalculationHostBinding[][] = [];
  input(sheetId: string, role: string, keys: readonly unknown[], build: () => WorkbookSheetInput): WorkbookSheetInput {
    const entries = this.#inputs.get(sheetId) ?? [];
    const found = entries.find(entry => entry.role === role && sameKeys(entry.keys, keys));
    if (found) return found.input;
    const input = build();
    this.#inputs.set(sheetId, [{ role, keys: [...keys], input }, ...entries.filter(entry => entry.role !== role)].slice(0, 2));
    return input;
  }
  orderedInputs(inputs: WorkbookSheetInput[]): WorkbookSheetInput[] {
    const found = this.#inputVectors.find(vector => sameKeys(vector, inputs));
    if (found) return found;
    this.#inputVectors = [inputs, ...this.#inputVectors].slice(0, 2);
    return inputs;
  }
  bindings(sheets: readonly SourceSheet[], inputs: readonly WorkbookSheetInput[], context: (input: WorkbookSheetInput) => BackgroundCalculationSheetContext): BackgroundCalculationHostBinding[] {
    const same = (entry: BackgroundCalculationHostBinding, sheet: SourceSheet, input: WorkbookSheetInput) =>
      entry.input === input && entry.rawSemanticsKey === calculationRawSemanticsKey(sheet) && entry.sheet.cells === sheet.cells &&
      entry.sheet.id === sheet.id && entry.sheet.name === sheet.name && entry.sheet.rowCount === sheet.rowCount &&
      entry.sheet.columnCount === sheet.columnCount && entry.sheet.pivot === sheet.pivot;
    const found = this.#vectors.find(vector => vector.length === sheets.length && vector.every((entry, index) => same(entry, sheets[index], inputs[index])));
    if (found) return found;
    const priorById = new Map<string, BackgroundCalculationHostBinding[]>();
    for (const vector of this.#vectors) for (const entry of vector) {
      const previous = priorById.get(entry.sheet.id) ?? [];
      if (!previous.includes(entry)) previous.push(entry);
      priorById.set(entry.sheet.id, previous);
    }
    const vector = inputs.map((input, index) => {
      const sheet = sheets[index];
      return priorById.get(sheet.id)?.find(entry => same(entry, sheet, input)) ?? {
        sheet, input, rawSemanticsKey: calculationRawSemanticsKey(sheet),
        resolverContext: { options: input.options, serialized: context(input) },
      };
    });
    this.#vectors = [vector, ...this.#vectors].slice(0, 2);
    const ids = new Set(sheets.map(sheet => sheet.id));
    for (const id of this.#inputs.keys()) if (!ids.has(id)) this.#inputs.delete(id);
    return vector;
  }
}
