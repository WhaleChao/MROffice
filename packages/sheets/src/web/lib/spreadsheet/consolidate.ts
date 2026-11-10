/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ─── Types ──────────────────────────────────────────────────────────────────

/** A 2-D grid of display-value strings. rows[r][c] is row r, column c. */
export type ValueGrid = readonly (readonly string[])[];

/**
 * A source cell's value as the functions read it: a number, an error (its
 * code), other content (text, a logical) or blank.
 */
export type ConsolidateValue = number | string | { readonly error: string } | null;

/** Which aggregate function to apply across the matching source values. */
export type ConsolidateFunction =
  | "Sum"
  | "Count"
  | "Average"
  | "Max"
  | "Min"
  | "Product"
  | "CountNumbers"
  | "StdDev"
  | "StdDevp"
  | "Var"
  | "Varp";

/** By-position mode: corresponding cell coordinates are aggregated directly. */
export interface ByPositionMode {
  readonly kind: "position";
}

/**
 * By-category mode: top-row labels and/or left-column labels form the union
 * key set; matching intersections across source grids are aggregated.
 * At least one of `useTopRow` / `useLeftColumn` must be true.
 */
export interface ByCategoryMode {
  readonly kind: "category";
  /** Whether the first row of each grid contains column-label headers. */
  readonly useTopRow: boolean;
  /** Whether the first column of each grid contains row-label headers. */
  readonly useLeftColumn: boolean;
}

export type ConsolidateMode = ByPositionMode | ByCategoryMode;

export interface ConsolidateSpec {
  /** Two or more source grids to aggregate. Grids may differ in shape. */
  readonly sources: readonly ValueGrid[];
  /** The aggregate function applied to each bucket of matching values. */
  readonly fn: ConsolidateFunction;
  /** How matching cells are determined across the source grids. */
  readonly mode: ConsolidateMode;
  /**
   * The sources' values, in the shape of `sources`, as formulas read them
   * (a formatted number is its number, a quoted "7" is text). Absent, each
   * display string is read: a number, an error code, or text.
   */
  readonly values?: readonly (readonly (readonly ConsolidateValue[])[])[];
}

/** The consolidated output grid (rows × columns of result strings). */
export interface ConsolidateResult {
  /**
   * The output grid as display-value strings. Numeric results are serialised
   * with up to 15 significant digits, an error as its code; a bucket with no
   * number and no error is "".
   */
  readonly grid: readonly (readonly string[])[];
  /** Number of output rows. */
  readonly rows: number;
  /** Number of output columns. */
  readonly cols: number;
  /** Exact source coordinates contributing to each output cell. */
  readonly contributors: ReadonlyArray<
    ReadonlyArray<ReadonlyArray<ConsolidateContributor>>
  >;
}

/** One source cell participating in a consolidated output bucket. */
export interface ConsolidateContributor {
  readonly sourceIndex: number;
  readonly row: number;
  readonly col: number;
}

const CONSOLIDATE_FORMULA_FUNCTIONS: Readonly<
  Record<ConsolidateFunction, string>
> = {
  Sum: "SUM",
  Count: "COUNTA",
  Average: "AVERAGE",
  Max: "MAX",
  Min: "MIN",
  Product: "PRODUCT",
  CountNumbers: "COUNT",
  StdDev: "STDEV.S",
  StdDevp: "STDEV.P",
  Var: "VAR.S",
  Varp: "VAR.P",
};

/** Build a live linked-result formula from authenticated source references. */
export function consolidateLinkFormula(
  fn: ConsolidateFunction,
  refs: readonly string[],
): string | null {
  if (refs.length === 0) return null;
  const args = refs.join(",");
  const aggregateFormula = `${CONSOLIDATE_FORMULA_FUNCTIONS[fn]}(${args})`;
  // A bucket with no number stays blank -- unless a source holds an error:
  // SUM over the sources carries their first error out, as the source's own
  // SUM does, instead of the blank that hid #DIV/0! and #N/A (S-070).
  const formula =
    fn === "Count" || fn === "CountNumbers"
      ? `=${aggregateFormula}`
      : `=IF(COUNT(${args})=0,IF(ISERROR(SUM(${args})),SUM(${args}),""),${aggregateFormula})`;
  return formula.length <= 8_192 ? formula : null;
}

// ─── Aggregate helpers ───────────────────────────────────────────────────────

const ERROR_CODES = new Set([
  "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
  "#GETTING_DATA", "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!",
  "#BUSY!", "#UNKNOWN!", "#EXTERNAL!", "#PYTHON!",
]);

/** A display string read as a value: blank, an error code, a number or text. */
export function consolidateValueOfDisplay(display: string): ConsolidateValue {
  const trimmed = display.trim();
  if (trimmed === "") return null;
  if (ERROR_CODES.has(trimmed.toUpperCase())) return { error: trimmed.toUpperCase() };
  const n = Number(trimmed);
  return Number.isFinite(n) ? n : display;
}

/**
 * A numeric result to Excel's 15 significant digits (Excel writes the full
 * value: StdDev of 5 and 3 is 1.4142135623731, which ten digits cut short).
 * parseFloat strips the trailing zeros, so "1.50000000000000" → "1.5".
 */
function fmtNum(n: number): string {
  return String(parseFloat(n.toPrecision(15)));
}

/** The squared deviations of `nums` from their mean, summed. */
function sumOfSquares(nums: readonly number[]): number {
  const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
  return nums.reduce((a, n) => a + (n - mean) ** 2, 0);
}

/** Apply the aggregate function to a bucket of source values, in source order. */
function aggregate(values: readonly ConsolidateValue[], fn: ConsolidateFunction): string {
  // Count is Excel's COUNTA: every value that is not blank, an error included.
  if (fn === "Count") return String(values.filter((value) => value !== null).length);
  const nums: number[] = [];
  let error: string | null = null;
  for (const value of values) {
    if (typeof value === "number") nums.push(value);
    else if (value !== null && typeof value === "object") error ??= value.error;
  }
  if (fn === "CountNumbers") return String(nums.length);
  // The sources' first error comes out, as Excel's consolidation gives it.
  if (error !== null) return error;
  if (nums.length === 0) return "";

  switch (fn) {
    case "Sum":
      return fmtNum(nums.reduce((a, b) => a + b, 0));
    case "Average":
      return fmtNum(nums.reduce((a, b) => a + b, 0) / nums.length);
    case "Max":
      return fmtNum(nums.reduce((a, b) => (b > a ? b : a), -Infinity));
    case "Min":
      return fmtNum(nums.reduce((a, b) => (b < a ? b : a), Infinity));
    case "Product":
      return fmtNum(nums.reduce((a, b) => a * b, 1));
    case "StdDev":
      // A sample (STDEV.S): one number has no deviation to estimate.
      return nums.length < 2 ? "#DIV/0!" : fmtNum(Math.sqrt(sumOfSquares(nums) / (nums.length - 1)));
    case "StdDevp":
      return fmtNum(Math.sqrt(sumOfSquares(nums) / nums.length));
    case "Var":
      return nums.length < 2 ? "#DIV/0!" : fmtNum(sumOfSquares(nums) / (nums.length - 1));
    case "Varp":
      return fmtNum(sumOfSquares(nums) / nums.length);
  }
}

/** The value of source `index`'s cell (row, col): from `values`, else its display. */
type ConsolidateReader = (index: number, row: number, col: number) => ConsolidateValue;

// ─── By-position consolidation ───────────────────────────────────────────────

/**
 * Consolidate by position: each (r, c) coordinate collects values from the
 * same position across all source grids. The output size equals the LARGEST
 * source grid (shorter grids contribute "" for their missing cells).
 */
function consolidateByPosition(
  sources: readonly ValueGrid[],
  fn: ConsolidateFunction,
  read: ConsolidateReader,
): ConsolidateResult {
  if (sources.length === 0)
    return { grid: [], rows: 0, cols: 0, contributors: [] };

  // Determine output dimensions as the max across all sources.
  let maxRows = 0;
  let maxCols = 0;
  for (const grid of sources) {
    if (grid.length > maxRows) maxRows = grid.length;
    for (const row of grid) {
      if (row.length > maxCols) maxCols = row.length;
    }
  }

  const outGrid: string[][] = [];
  const contributors: ConsolidateContributor[][][] = [];
  for (let r = 0; r < maxRows; r += 1) {
    const row: string[] = [];
    const contributorRow: ConsolidateContributor[][] = [];
    for (let c = 0; c < maxCols; c += 1) {
      const bucket: ConsolidateValue[] = [];
      const sourceCells: ConsolidateContributor[] = [];
      for (let sourceIndex = 0; sourceIndex < sources.length; sourceIndex += 1) {
        const grid = sources[sourceIndex];
        bucket.push(read(sourceIndex, r, c));
        if (grid[r] && c < grid[r].length) {
          sourceCells.push({ sourceIndex, row: r, col: c });
        }
      }
      row.push(aggregate(bucket, fn));
      contributorRow.push(sourceCells);
    }
    outGrid.push(row);
    contributors.push(contributorRow);
  }

  return { grid: outGrid, rows: maxRows, cols: maxCols, contributors };
}

// ─── By-category consolidation ───────────────────────────────────────────────

/**
 * Collect the ordered-union of label strings from a set of label arrays,
 * preserving first-seen order (Excel keeps source order for labels).
 */
function labelUnion(labelArrays: readonly (readonly string[])[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const arr of labelArrays) {
    for (const label of arr) {
      if (!seen.has(label)) {
        seen.add(label);
        result.push(label);
      }
    }
  }
  return result;
}

/**
 * Consolidate by category: top-row labels, left-column labels, or both form
 * the alignment keys. The output grid contains the union of all labels in the
 * chosen dimension(s), with aggregated values at intersections.
 */
function consolidateByCategory(
  sources: readonly ValueGrid[],
  fn: ConsolidateFunction,
  useTopRow: boolean,
  useLeftColumn: boolean,
  read: ConsolidateReader,
): ConsolidateResult {
  if (sources.length === 0)
    return { grid: [], rows: 0, cols: 0, contributors: [] };

  // Fallback: if neither label dimension is active, treat as by-position.
  if (!useTopRow && !useLeftColumn) {
    return consolidateByPosition(sources, fn, read);
  }

  // For each source extract the label row, label column, and data sub-grid.
  const dataR0 = useTopRow ? 1 : 0; // first data row index in source
  const dataC0 = useLeftColumn ? 1 : 0; // first data col index in source

  // Collect per-source row-labels and col-labels (blank label treated as "").
  const sourceRowLabels: (readonly string[])[] = sources.map((grid) => {
    if (!useLeftColumn) return [];
    const labels: string[] = [];
    for (let r = dataR0; r < grid.length; r += 1) {
      labels.push(grid[r]?.[0] ?? "");
    }
    return labels;
  });

  const sourceColLabels: (readonly string[])[] = sources.map((grid) => {
    if (!useTopRow) return [];
    const labelRow = grid[0] ?? [];
    return labelRow.slice(dataC0);
  });

  // Build the union of row-labels and col-labels across all sources.
  const rowLabels = useLeftColumn ? labelUnion(sourceRowLabels) : [];
  const colLabels = useTopRow ? labelUnion(sourceColLabels) : [];

  // Build lookup maps: label → index into the union arrays.
  const rowLabelIndex = new Map<string, number>();
  rowLabels.forEach((lbl, i) => rowLabelIndex.set(lbl, i));
  const colLabelIndex = new Map<string, number>();
  colLabels.forEach((lbl, i) => colLabelIndex.set(lbl, i));

  // Determine data dimensions in the output.
  const dataRows = useLeftColumn ? rowLabels.length : 0;
  const dataCols = useTopRow ? colLabels.length : 0;

  // Determine the full output grid size.
  // When both label dimensions are used we need to handle the corner cell.
  // Output rows = (header row if useTopRow) + (one per rowLabel if useLeftColumn)
  // Output cols = (header col if useLeftColumn) + (one per colLabel if useTopRow)
  // When only one dimension is used:
  //   useTopRow only → 1 header row + N data rows (matching by col-label only)
  //   useLeftColumn only → 1 header col + M data cols (matching by row-label only)

  // Determine the number of data rows in the output when only useTopRow is active
  // (no left-column labels): we align by column only, so row count = max data rows
  // across sources.
  let outDataRowsWhenColOnly = 0;
  if (useTopRow && !useLeftColumn) {
    for (const grid of sources) {
      const n = Math.max(0, grid.length - dataR0);
      if (n > outDataRowsWhenColOnly) outDataRowsWhenColOnly = n;
    }
  }

  let outDataColsWhenRowOnly = 0;
  if (useLeftColumn && !useTopRow) {
    for (const grid of sources) {
      const row0 = grid[0] ?? [];
      const n = Math.max(0, row0.length - dataC0);
      if (n > outDataColsWhenRowOnly) outDataColsWhenRowOnly = n;
    }
  }

  // Build accumulation buckets: outR → outC → the values, in source order.
  // We use two index systems:
  //   outR: 0-based row in the OUTPUT DATA block (excludes any header row)
  //   outC: 0-based col in the OUTPUT DATA block (excludes any header col)
  const totalDataRows =
    useLeftColumn ? dataRows : outDataRowsWhenColOnly;
  const totalDataCols =
    useTopRow ? dataCols : outDataColsWhenRowOnly;

  const buckets: ConsolidateValue[][][] = Array.from({ length: totalDataRows }, () =>
    Array.from({ length: totalDataCols }, () => []),
  );
  const contributorBuckets: ConsolidateContributor[][][] = Array.from(
    { length: totalDataRows },
    () => Array.from({ length: totalDataCols }, () => []),
  );

  // Accumulate values from each source grid.
  for (let si = 0; si < sources.length; si += 1) {
    const grid = sources[si];
    const srcRowLabels = sourceRowLabels[si];
    const srcColLabels = sourceColLabels[si];

    const srcDataRowCount = Math.max(0, grid.length - dataR0);
    const srcDataColCount = Math.max(
      0,
      (grid[0]?.length ?? 0) - dataC0,
    );

    for (let dr = 0; dr < srcDataRowCount; dr += 1) {
      // Which output row?
      let outR: number;
      if (useLeftColumn) {
        const lbl = srcRowLabels[dr] ?? "";
        const idx = rowLabelIndex.get(lbl);
        if (idx === undefined) continue; // label not in union (shouldn't happen)
        outR = idx;
      } else {
        outR = dr; // by-position within data rows (useTopRow only)
      }
      if (outR >= totalDataRows) continue;

      for (let dc = 0; dc < srcDataColCount; dc += 1) {
        // Which output col?
        let outC: number;
        if (useTopRow) {
          const lbl = srcColLabels[dc] ?? "";
          const idx = colLabelIndex.get(lbl);
          if (idx === undefined) continue;
          outC = idx;
        } else {
          outC = dc; // by-position within data cols (useLeftColumn only)
        }
        if (outC >= totalDataCols) continue;

        buckets[outR][outC].push(read(si, dataR0 + dr, dataC0 + dc));
        contributorBuckets[outR][outC].push({
          sourceIndex: si,
          row: dataR0 + dr,
          col: dataC0 + dc,
        });
      }
    }
  }

  // Assemble the output grid (with label rows/cols prepended).
  const headerRows = useTopRow ? 1 : 0;
  const headerCols = useLeftColumn ? 1 : 0;
  const outRows = headerRows + totalDataRows;
  const outCols = headerCols + totalDataCols;

  const outGrid: string[][] = Array.from({ length: outRows }, () =>
    Array.from({ length: outCols }, () => ""),
  );
  const contributors: ConsolidateContributor[][][] = Array.from(
    { length: outRows },
    () => Array.from({ length: outCols }, () => []),
  );

  // Fill the corner cell when both label dimensions are active.
  // (Excel leaves it blank — we do the same.)

  // Fill top-row label cells.
  if (useTopRow) {
    for (let ci = 0; ci < colLabels.length; ci += 1) {
      outGrid[0][headerCols + ci] = colLabels[ci];
    }
  }

  // Fill left-column label cells.
  if (useLeftColumn) {
    for (let ri = 0; ri < rowLabels.length; ri += 1) {
      outGrid[headerRows + ri][0] = rowLabels[ri];
    }
  }

  // Fill data cells.
  for (let r = 0; r < totalDataRows; r += 1) {
    for (let c = 0; c < totalDataCols; c += 1) {
      outGrid[headerRows + r][headerCols + c] = aggregate(buckets[r][c], fn);
      contributors[headerRows + r][headerCols + c] =
        contributorBuckets[r][c];
    }
  }

  return { grid: outGrid, rows: outRows, cols: outCols, contributors };
}

// ─── Public entry point ─────────────────────────────────────────────────────

/**
 * Consolidate multiple source grids into a single summary grid.
 *
 * @param spec the consolidation parameters
 * @returns the output grid and its dimensions
 */
export function consolidate(spec: ConsolidateSpec): ConsolidateResult {
  const { sources, fn, mode, values } = spec;
  const read: ConsolidateReader = (index, row, col) => {
    const typed = values?.[index]?.[row]?.[col];
    return typed !== undefined ? typed : consolidateValueOfDisplay(sources[index]?.[row]?.[col] ?? "");
  };

  if (mode.kind === "position") {
    return consolidateByPosition(sources, fn, read);
  }

  return consolidateByCategory(sources, fn, mode.useTopRow, mode.useLeftColumn, read);
}

/** The ordered list of aggregate function names for the dialog select. */
export const CONSOLIDATE_FUNCTIONS: readonly ConsolidateFunction[] = [
  "Sum",
  "Count",
  "Average",
  "Max",
  "Min",
  "Product",
  "CountNumbers",
  "StdDev",
  "StdDevp",
  "Var",
  "Varp",
] as const;

/** Each function's name in the catalog, shared with Subtotal's list (spreadsheet.subtotalFn.*). */
export const CONSOLIDATE_FUNCTION_LABEL_KEYS: Readonly<Record<ConsolidateFunction, string>> = {
  Sum: "spreadsheet.subtotalFn.sum",
  Count: "spreadsheet.subtotalFn.count",
  Average: "spreadsheet.subtotalFn.average",
  Max: "spreadsheet.subtotalFn.max",
  Min: "spreadsheet.subtotalFn.min",
  Product: "spreadsheet.subtotalFn.product",
  CountNumbers: "spreadsheet.subtotalFn.countNumbers",
  StdDev: "spreadsheet.subtotalFn.stdDev",
  StdDevp: "spreadsheet.subtotalFn.stdDevp",
  Var: "spreadsheet.subtotalFn.var",
  Varp: "spreadsheet.subtotalFn.varp",
};
