/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  applyPivotFilters,
  computePivot,
  headerLabel,
  matchDataFieldIndex,
  rangeContains,
  normalizePivotRange,
  pivotBucketLabel,
  pivotGroupItemOrder,
  pivotRangeToA1,
  type PivotConfig,
  type PivotSubtotalPosition,
  type PivotDataCriterion,
  type PivotDataResult,
  type PivotMatrix,
  type PivotPlacement,
  type PivotRange,
  type PivotResult,
  PIVOT_SHOW_AS_DIVIDE_BY_ZERO,
  PIVOT_SHOW_AS_NO_FIELD,
  isPivotAggDivideByZero,
  showAsIndex,
  showAsShareOfTotal,
  showValuesAsAlongBase,
  showValuesAsNeedsBaseField,
  showValuesAsNeedsBaseItem,
  valueFieldLabel,
  expandPivotGroupLevels,
  pivotItemCaption,
  pivotRowsAfterFieldFilters,
  pivotValueFieldAggregate,
  pivotFieldCaption,
  pivotSubtotalFunctionLabel,
  pivotCalculatedItemRecords,
  aggregate as aggregateValues,
  type PivotAgg,
} from "./pivot";
import { applyNumberFormat } from "./number-format";

export type PivotOutputCellRole =
  | "corner"
  | "column-header"
  | "row-header"
  | "value"
  | "subtotal-label"
  | "subtotal-value"
  | "blank"
  | "grand-total-label"
  | "grand-total-value"
  | "page-field"
  | "page-item";

export interface PivotOutputCell {
  row: number;
  col: number;
  value: string | number | null;
  role: PivotOutputCellRole;
  valueIndex?: number;
  /** Compact form indents an inner item's label by its level. */
  indent?: number;
  /** Compact form's Row Labels column: Excel draws every label in it from
   * the left, a number included (S01_pivottable_error_item_filter's 0.5). */
  compactLabel?: true;
  /** The source field a caption or report-filter cell stands for (its
   * header dropdown filters and sorts that field). */
  field?: number;
  /** An outer item's label carries Excel's expand/collapse button; the path
   * names the item to the planner's `collapsed` option. */
  drill?: "expanded" | "collapsed";
  drillPath?: string;
  /** The row field and item the button shows or hides the details of
   * (Excel's show-details flag is the item's own, `PivotConfig.collapsed`). */
  drillField?: number;
  drillItem?: string;
  /** A column item's label: the column field whose item it is (a format on
   * that field's item labels paints it, X-294a). */
  itemField?: number;
  /** A data cell's row and column items, outer to inner (raw), as the
   * report's formats on chosen items' values find it (X-294a). */
  rowPath?: readonly string[];
  columnPath?: readonly string[];
}

/** The text a PivotTable writes itself, in the viewer's language. */
export interface PivotOutputLabels {
  grandTotal: string;
  /** A group's subtotal ("East Total"). */
  subtotal: (item: string) => string;
  /** A value field's Grand Total column ("Total Sum of Sales"). */
  valueTotal: (caption: string) => string;
  /** A column group's subtotal of one value field ("East Sum of Sales"). */
  itemValue: (item: string, caption: string) => string;
  rowLabels: string;
  columnLabels: string;
  values: string;
  allItems: string;
  multipleItems: string;
  blank: string;
  /** A PivotChart's one series (or category) without a field ("Total"). */
  total: string;
  /** A Custom subtotal's function, as the subtotal row names it ("Sum",
   * 求和); English when the host names none. */
  subtotalFunction?: (agg: PivotAgg) => string;
}

export interface PivotOutputPlan {
  status: "ready" | "missing-anchor" | "out-of-bounds";
  /** The report (the PivotTable's location), header rows first. */
  output: PivotRange | null;
  /** The report filter rows above the report: a field beside its selection. */
  pageRange: PivotRange | null;
  /** Rows of field captions and column items at the top of the report. */
  headerRows: number;
  /** Columns of row labels (or the value caption) at the left of the report. */
  firstDataCol: number;
  cells: ReadonlyMap<string, PivotOutputCell>;
  collisions: readonly string[];
  placement: PivotPlacement | null;
  resolvePivotData: (
    anchor: { row: number; col: number },
    dataField: string,
    criteria: readonly PivotDataCriterion[],
  ) => PivotDataResult;
  /**
   * What a PivotChart bound to this report plots, whether or not the report
   * can be placed: its categories and series as the report shows them.
   */
  chart?: PivotChartReport;
}

/** One PivotChart series: a value column of the report. */
export interface PivotChartReportSeries {
  /** The legend text: the column items level by level, then the value
   * caption when there are several values, joined " - "; "Total" alone. */
  name: string;
  valueIndex: number;
  /** The column items (raw) the series stands for; [] without column fields. */
  columnPath: string[];
  /** One value per category; null is a gap (no records, or an error). */
  values: (number | null)[];
}

/** An axis, legend or filter field of the report, for the chart's buttons. */
export interface PivotChartReportField {
  colIndex: number;
  caption: string;
}

/**
 * A PivotChart's reading of its PivotTable (Excel): the categories are the
 * report's row items at the innermost shown level (a collapsed item stands
 * for its group), without subtotal and grand-total rows; the series are its
 * value columns without subtotal and grand-total columns. A report without
 * row fields has the one category "Total".
 */
export interface PivotChartReport {
  /** The innermost label of each category. */
  categories: string[];
  /** Each category's items outer to inner (raw); shorter when collapsed. */
  categoryPaths: string[][];
  /** Each category's labels outer to inner, as the report shows them. */
  categoryLabels: string[][];
  series: PivotChartReportSeries[];
  rowFields: PivotChartReportField[];
  columnFields: PivotChartReportField[];
  filterFields: PivotChartReportField[];
  valueCaptions: string[];
}

export interface PlanPivotOutputOptions {
  rowCount: number;
  columnCount: number;
  isOccupied?: (row: number, col: number) => boolean;
  labels?: Partial<PivotOutputLabels>;
  /** Outer items shown collapsed (their `drillPath`s): one row each, carrying
   * the item's totals, as Excel's minus button leaves them. View state only. */
  collapsed?: ReadonlySet<string>;
}

const DEFAULT_LABELS: PivotOutputLabels = {
  grandTotal: "Grand Total",
  subtotal: (item) => `${item} Total`,
  valueTotal: (caption) => `Total ${caption}`,
  itemValue: (item, caption) => `${item} ${caption}`,
  rowLabels: "Row Labels",
  columnLabels: "Column Labels",
  values: "Values",
  allItems: "(All)",
  multipleItems: "(Multiple Items)",
  blank: "(blank)",
  total: "Total",
};

/**
 * Pick a deterministic initial destination beside the source, falling back
 * below it when the source already reaches the sheet's right edge.
 *
 * `footprint.above` counts the rows report filters take above the report
 * (the filters and the blank row under them): a report beside the source
 * moves down to make room, one below the source starts that much lower.
 */
export function suggestPivotOutputAnchor(
  source: unknown,
  bounds: { rowCount: number; columnCount: number },
  footprint: { height: number; width: number; above?: number } = {
    height: 1,
    width: 1,
  },
): string | null {
  const range = normalizePivotRange(source);
  if (!range || bounds.rowCount < 1 || bounds.columnCount < 1) return null;
  const height = Math.max(1, Math.trunc(footprint.height));
  const width = Math.max(1, Math.trunc(footprint.width));
  const above = Math.max(0, Math.trunc(footprint.above ?? 0));
  const rightCol = range.c2 + 2;
  if (
    rightCol + width <= bounds.columnCount &&
    Math.max(range.r1, above) + height <= bounds.rowCount
  ) {
    return pivotRangeToA1({
      r1: range.r1,
      c1: rightCol,
      r2: range.r1,
      c2: rightCol,
    });
  }
  const belowRow = range.r2 + 2 + above;
  if (
    belowRow + height <= bounds.rowCount &&
    range.c1 + width <= bounds.columnCount
  ) {
    return pivotRangeToA1({
      r1: belowRow,
      c1: range.c1,
      r2: belowRow,
      c2: range.c1,
    });
  }
  return null;
}

/** Whether any of the ranges holds the cell. */
export function pivotRangesContain(
  ranges: readonly (PivotRange | null | undefined)[],
  row: number,
  col: number,
): boolean {
  return ranges.some(
    (range) =>
      Boolean(range) &&
      row >= range!.r1 &&
      row <= range!.r2 &&
      col >= range!.c1 &&
      col <= range!.c2,
  );
}

/**
 * The report filter rows an imported PivotTable left above its report, while
 * they still hold only what a PivotTable writes there: a field name beside its
 * selection ("(All)" in Excel's own language). The export treats the same
 * block as the PivotTable's.
 */
export function nativePivotPageBlock(
  report: PivotRange,
  filterCount: number,
  fieldNames: ReadonlySet<string>,
  cellAt: (row: number, col: number) => { value?: unknown; formula?: unknown } | undefined,
): PivotRange | null {
  const top = report.r1 - filterCount - 1;
  if (filterCount < 1 || top < 0) return null;
  const block = { r1: top, c1: report.c1, r2: report.r1 - 2, c2: report.c1 + 1 };
  for (let row = block.r1; row <= block.r2; row += 1) {
    for (const col of [block.c1, block.c2]) {
      const cell = cellAt(row, col);
      if (!cell) continue;
      if (String(cell.formula ?? "").trim()) return null;
      const value = cell.value;
      if (value === undefined || value === null || value === "") continue;
      if (col === block.c1 && !fieldNames.has(String(value))) return null;
    }
  }
  return block;
}

function cellKey(row: number, col: number): string {
  return `${row}:${col}`;
}

function emptyResolver(): PivotOutputPlan["resolvePivotData"] {
  return () => ({ kind: "not-pivot" });
}

const KEY_SEP = String.fromCharCode(0);
const pathKey = (path: readonly string[]) => path.join(KEY_SEP);

interface AxisNode {
  item: string;
  path: string[];
  rows: string[][];
  children: AxisNode[];
}

/** A report column: its column group (or none) and the value it shows. */
type ReportColumn =
  | { kind: "value"; valueIndex: number | null }
  | { kind: "item"; path: string[]; valueIndex: number | null }
  | { kind: "subtotal"; path: string[]; valueIndex: number | null }
  | { kind: "grand"; valueIndex: number | null };

interface PlannedCell {
  value: string | number | null;
  role: PivotOutputCellRole;
  valueIndex?: number;
  indent?: number;
  field?: number;
  drill?: "expanded" | "collapsed";
  drillPath?: string;
  drillField?: number;
  drillItem?: string;
  itemField?: number;
  rowPath?: readonly string[];
  columnPath?: readonly string[];
}

/**
 * Lay a PivotTable out as Excel does (checked against Excel-made files; the
 * XLSX writer's `_pivot_output_grid` is the same layout).
 *
 * Tabular form gives each row field a column and starts a group on the row of
 * its first inner item; outline form gives an outer item its own row; compact
 * form does too, in one indented column under "Row Labels". Subtotals sit on
 * the outer item's row (top) or below the group; tabular form has only the
 * latter. Column fields and several values add header rows: the field
 * captions, the column items level by level, then the value captions. A
 * combination without records stays blank. Report filters take a row each
 * above the report, with a blank row between; an anchor too near the top moves
 * the report down, as Excel moves a PivotTable when a filter is added.
 *
 * The plan never mutates or serializes worksheet cells. It gives the React host
 * one bounded map for rendering plus the exact placement required by
 * GETPIVOTDATA. An occupied destination is reported rather than overwritten,
 * and an output that would exceed the logical sheet is rejected atomically.
 */
/**
 * Where a report draws its group subtotals: the stored position, else the
 * legacy on/off flag's "bottom", else none. The Field List's Subtotals box
 * reads the same rule so it toggles what the report shows.
 */
export function pivotSubtotalPosition(
  config: Pick<PivotConfig, "subtotalPosition" | "rowSubtotals">,
): PivotSubtotalPosition {
  return config.subtotalPosition ?? (config.rowSubtotals ? "bottom" : "none");
}

export function planPivotOutput(
  sourceConfig: PivotConfig,
  sourceMatrix: PivotMatrix,
  options: PlanPivotOutputOptions,
): PivotOutputPlan {
  // A date field grouped by several units is laid out as several axis
  // fields (Years, Quarters, then the field), as Excel lays it out.
  const { matrix, config } = expandPivotGroupLevels(sourceMatrix, sourceConfig);
  const anchor = normalizePivotRange(config.output);
  const labels: PivotOutputLabels = {
    ...DEFAULT_LABELS,
    ...options.labels,
    // The report's own Values caption, as its file states it (X-249).
    ...(config.native?.dataCaption ? { values: config.native.dataCaption } : {}),
  };
  const rowFields = config.rows;
  const columnFields = config.columns;
  const values = config.values;
  const valueCount = values.length;
  // Σ Values in the Rows area (Excel's dataOnRows), as the innermost row
  // field: each innermost row item gets a row per value field, and the
  // columns no longer split by value (X-294c, measured in Excel for Mac).
  const valuesOnRows = config.valuesOnRows === true && valueCount > 1;
  const layout = config.reportLayout ?? "tabular";
  const compact = layout === "compact" && rowFields.length > 0;
  const tabular = layout === "tabular";
  const subtotalPosition = pivotSubtotalPosition(config);
  const subtotals = subtotalPosition !== "none";
  const subtotalsOnTop = subtotalPosition === "top" && !tabular;
  const repeatLabels = config.repeatItemLabels === true && !compact;
  // "On for Rows Only" totals each row (the Grand Total column); "On for
  // Columns Only" is the Grand Total row.
  const grandTotals = config.grandTotals ?? "both";
  const grandColumn = grandTotals === "both" || grandTotals === "rows";
  const grandRow = grandTotals === "both" || grandTotals === "columns";
  const captions = values.map((field) =>
    valueFieldLabel(field, headerLabel(matrix, field.colIndex)),
  );
  // An item typed over in the report shows its own caption (the item's `n`).
  const itemCaptionsByField = new Map<number, Map<string, string>>();
  for (const entry of config.itemCaptions ?? []) {
    let byItem = itemCaptionsByField.get(entry.colIndex);
    if (!byItem) itemCaptionsByField.set(entry.colIndex, (byItem = new Map()));
    byItem.set(entry.item, entry.caption);
  }
  const itemLabel = (item: string, field?: number) =>
    (field === undefined ? undefined : itemCaptionsByField.get(field)?.get(item)) ??
    (item === "" ? labels.blank : pivotItemCaption(item));
  const fieldLabel = (field: number) => pivotFieldCaption(config, matrix, field);

  const sourceRows = Array.isArray(matrix) ? matrix.slice(1) : [];
  const groupingByCol = new Map(
    (config.grouping ?? []).map((grouping) => [grouping.colIndex, grouping]),
  );
  const rawValue = (row: readonly string[], col: number) =>
    row[col] == null ? "" : String(row[col]);
  // Excel compares a PivotTable's text items without regard to case: "Scone",
  // "SconE" and "SCONE" are one item, spelled as the source first spells it,
  // as computePivot folds them (CX093; X-294b).
  const spellings = new Map<number, Map<string, string>>();
  const spelledAs = (col: number, raw: string): string => {
    let spelled = spellings.get(col);
    if (!spelled) {
      spelled = new Map();
      for (const row of matrix.slice(1)) {
        const value = rawValue(row, col);
        const folded = value.toLowerCase();
        if (!spelled.has(folded)) spelled.set(folded, value);
      }
      spellings.set(col, spelled);
    }
    return spelled.get(raw.toLowerCase()) ?? raw;
  };
  const itemOf = (row: readonly string[], col: number) => {
    const grouping = groupingByCol.get(col);
    const raw = rawValue(row, col);
    return grouping ? pivotBucketLabel(raw, grouping) : spelledAs(col, raw);
  };
  // Items keep the order they first appear in the source (the cache order the
  // XLSX writer lists them in), level by level — unless the field is sorted,
  // or lists its items in a manual order of its own (Excel's PivotField item
  // list), which then comes first; items it does not name follow first-seen.
  const itemOrders = new Map<number, Map<string, number>>();
  const compareItems = (a: string, b: string) => {
    const x = Number(a);
    const y = Number(b);
    const numericA = a.trim() !== "" && Number.isFinite(x);
    const numericB = b.trim() !== "" && Number.isFinite(y);
    if (numericA && numericB) return x - y;
    if (numericA !== numericB) return numericA ? -1 : 1;
    return a.localeCompare(b);
  };
  const itemOrder = (col: number) => {
    let order = itemOrders.get(col);
    if (!order) {
      const seen: string[] = [];
      for (const row of sourceRows) {
        const item = itemOf(row, col);
        if (!seen.includes(item)) seen.push(item);
      }
      // A calculated item follows the field's own items, as Excel lists it.
      for (const item of config.calculatedItems ?? []) {
        if (item.colIndex === col && !seen.includes(item.name)) seen.push(item.name);
      }
      const sortEntry = config.fieldSorts?.find((entry) => entry.colIndex === col);
      // A sort by a value field orders items per outer item (buildAxis).
      const sort = sortEntry?.byValue === undefined ? sortEntry?.order : undefined;
      const manual = config.fieldItemOrders?.find((entry) => entry.colIndex === col);
      let ordered = seen;
      const grouping = groupingByCol.get(col);
      // A date or number group lists its items in the group's own order
      // (years ascending, months Jan to Dec, bins ascending), as Excel's
      // groupItems do and as computePivot sorts them, not by first sighting:
      // a year whose February rows come after March still shows Jan, Feb, Mar.
      const groupOrder =
        grouping && grouping.kind !== "discrete" ? pivotGroupItemOrder(grouping, seen) : null;
      if (groupOrder) {
        ordered = [...seen].sort((a, b) => (groupOrder.get(a) ?? 0) - (groupOrder.get(b) ?? 0));
      } else if (sort && !groupingByCol.has(col)) {
        ordered = [...seen].sort((a, b) => (sort === "descending" ? -compareItems(a, b) : compareItems(a, b)));
      } else if (manual && !groupingByCol.has(col)) {
        const position = new Map(manual.items.map((item, index) => [item, index] as const));
        ordered = [...seen].sort((a, b) => {
          const pa = position.get(a);
          const pb = position.get(b);
          if (pa === undefined && pb === undefined) return seen.indexOf(a) - seen.indexOf(b);
          if (pa === undefined) return 1;
          if (pb === undefined) return -1;
          return pa - pb;
        });
      }
      order = new Map(ordered.map((item, index) => [item, index] as const));
      itemOrders.set(col, order);
    }
    return order;
  };
  const reportFiltered = applyPivotFilters(
    sourceRows,
    [...config.filters, ...(config.itemFilters ?? [])],
    config.dateFilters,
    config.grouping,
  );
  // Calculated items add their own records (pivotCalculatedItemRecords).
  const calculatedRecords = pivotCalculatedItemRecords(reportFiltered, matrix[0] ?? [], config);
  const filteredRows = pivotRowsAfterFieldFilters(
    [...reportFiltered, ...calculatedRecords],
    matrix[0] ?? [],
    config,
  );
  // A field's Value filter judges its items within the whole layout, so the
  // per-depth aggregates below read the records the full report kept rather
  // than re-filtering a report cut to fewer fields.
  const fieldFiltered = (config.fieldFilters?.length ?? 0) > 0 || calculatedRecords.length > 0;
  const aggregateMatrix = fieldFiltered ? [matrix[0] ?? [], ...filteredRows] : matrix;
  const buildAxis = (
    rows: string[][],
    fields: readonly number[],
    depth: number,
    prefix: string[],
  ): AxisNode[] => {
    if (depth >= fields.length) return [];
    const groups = new Map<string, string[][]>();
    for (const row of rows) {
      const item = itemOf(row, fields[depth]);
      const group = groups.get(item);
      if (group) group.push(row);
      else groups.set(item, [row]);
    }
    const order = itemOrder(fields[depth]);
    // A field sorted by a value field orders each outer item's items by that
    // value over every column (Excel's autoSortScope); items without a value
    // follow in their own order.
    const valueSort = config.fieldSorts?.find(
      (entry) => entry.colIndex === fields[depth] && entry.byValue !== undefined && entry.byValue < valueCount,
    );
    const measured = valueSort
      ? new Map(
          [...groups].map(([item, group]) => [
            item,
            pivotValueFieldAggregate(config, matrix[0] ?? [], group, valueSort.byValue as number),
          ]),
        )
      : null;
    const byValue = (a: string, b: string): number => {
      const x = measured!.get(a);
      const y = measured!.get(b);
      const missingX = x === null || x === undefined || Number.isNaN(x);
      const missingY = y === null || y === undefined || Number.isNaN(y);
      if (missingX !== missingY) return missingX ? 1 : -1;
      if (!missingX && x !== y) return valueSort!.order === "descending" ? (y as number) - (x as number) : (x as number) - (y as number);
      return (order.get(a) ?? 0) - (order.get(b) ?? 0);
    };
    return [...groups.keys()]
      .sort(measured ? byValue : (a, b) => (order.get(a) ?? 0) - (order.get(b) ?? 0))
      .map((item) => {
        const path = [...prefix, item];
        const group = groups.get(item)!;
        return {
          item,
          path,
          rows: group,
          children: buildAxis(group, fields, depth + 1, path),
        };
      });
  };
  const rowTree = buildAxis(filteredRows, rowFields, 0, []);
  const columnTree = buildAxis(filteredRows, columnFields, 0, []);
  // An item whose details are hidden (Excel's show-details flag, `sd="0"`,
  // an item's own and so the same under every parent), or a path the
  // report's own minus button collapsed.
  const collapsedItems = new Map<number, Set<string>>();
  for (const entry of config.collapsed ?? []) {
    collapsedItems.set(entry.colIndex, new Set(entry.items));
  }
  const isCollapsed = (node: AxisNode, depth: number) =>
    depth < rowFields.length - 1 &&
    (options.collapsed?.has(pathKey(node.path)) === true ||
      collapsedItems.get(rowFields[depth])?.has(node.item) === true);
  // A column item whose details are hidden is one column (its total), as
  // Excel shows a year collapsed over its quarters and months; a PivotChart
  // then has one series per collapsed item (a workbook whose years are
  // collapsed charts 2015, 2016, 2017, not every month).
  const isColumnCollapsed = (node: AxisNode, depth: number) =>
    depth < columnFields.length - 1 &&
    collapsedItems.get(columnFields[depth])?.has(node.item) === true;

  // Which (row group, column group) combinations have records, by depth.
  const present = new Set<string>();
  for (const row of filteredRows) {
    const rowItems = rowFields.map((col) => itemOf(row, col));
    const columnItems = columnFields.map((col) => itemOf(row, col));
    for (let rowDepth = 0; rowDepth <= rowFields.length; rowDepth += 1) {
      const rowKey = pathKey(rowItems.slice(0, rowDepth));
      for (let colDepth = 0; colDepth <= columnFields.length; colDepth += 1) {
        present.add(
          `${rowDepth}|${rowKey}|${colDepth}|${pathKey(columnItems.slice(0, colDepth))}`,
        );
      }
    }
  }
  const hasRecords = (rowPath: readonly string[], columnPath: readonly string[]) =>
    present.has(
      `${rowPath.length}|${pathKey(rowPath)}|${columnPath.length}|${pathKey(columnPath)}`,
    );

  // Plain aggregates of a row group over a column group, per value field. The
  // compute engine owns aggregation (calculated fields included); each depth
  // pair is computed once.
  const plainValues = values.map((field) => ({ ...field, showAs: undefined }));
  const aggregates = new Map<
    string,
    { result: PivotResult; rows: Map<string, number>; columns: Map<string, number> }
  >();
  const aggregate = (
    rowPath: readonly string[],
    columnPath: readonly string[],
    valueIndex: number,
  ): number | null => {
    if (!hasRecords(rowPath, columnPath)) return null;
    const depthKey = `${rowPath.length}:${columnPath.length}`;
    let table = aggregates.get(depthKey);
    if (!table) {
      const result = computePivot(aggregateMatrix, {
        ...config,
        ...(fieldFiltered
          ? {
              filters: [],
              itemFilters: [],
              dateFilters: undefined,
              fieldFilters: undefined,
              calculatedItems: undefined,
            }
          : {}),
        rows: rowFields.slice(0, rowPath.length),
        columns: columnFields.slice(0, columnPath.length),
        values: plainValues,
        rowSubtotals: false,
        subtotalPosition: undefined,
      });
      table = {
        result,
        rows: new Map(result.rowHeaders.map((tuple, index) => [pathKey(tuple), index])),
        columns: new Map(
          (result.columnTuples ?? [[]]).map((tuple, index) => [pathKey(tuple), index]),
        ),
      };
      aggregates.set(depthKey, table);
    }
    const rowIndex = table.rows.get(pathKey(rowPath));
    const columnIndex = table.columns.get(pathKey(columnPath));
    if (rowIndex === undefined || columnIndex === undefined) return null;
    return table.result.cells[rowIndex]?.[columnIndex * valueCount + valueIndex] ?? null;
  };

  // A field's own Subtotals (Field Settings): None, or Custom functions.
  const fieldSubtotalsByCol = new Map(
    (config.fieldSubtotals ?? []).map((entry) => [entry.colIndex, entry.functions] as const),
  );
  // A field's one Custom subtotal shown at the top of each group, outside
  // tabular form (`native.subtotalTopFields`, X-311): drawn on the group's own
  // row, with no row of its own below the group.
  const subtotalTopFields = new Set(config.native?.subtotalTopFields ?? []);
  const ownSubtotalOnTop = (field: number, own: readonly PivotAgg[] | undefined): boolean =>
    !tabular && own?.length === 1 && subtotalTopFields.has(field);
  // A custom subtotal: the value field's records under the group, by the
  // subtotal's own function (a calculated field keeps its formula).
  const customSubtotal = (
    rowPath: readonly string[],
    column: ReportColumn,
    valueIndex: number,
    agg: PivotAgg,
  ): number | string | null => {
    const columnPath = column.kind === "item" || column.kind === "subtotal" ? column.path : [];
    if (!hasRecords(rowPath, columnPath)) return null;
    const records = filteredRows.filter(
      (row) =>
        rowPath.every((item, depth) => itemOf(row, rowFields[depth]) === item) &&
        columnPath.every((item, depth) => itemOf(row, columnFields[depth]) === item),
    );
    const field = values[valueIndex];
    const value = field.calcRef
      ? pivotValueFieldAggregate(config, matrix[0] ?? [], records, valueIndex)
      : aggregateValues(
          agg,
          records.map((row) => (row[field.colIndex] == null ? "" : String(row[field.colIndex]))),
        );
    return typeof value === "number" && isPivotAggDivideByZero(value) ? PIVOT_SHOW_AS_DIVIDE_BY_ZERO : value;
  };
  // The report's columns: one per value field under each column item, or
  // (Σ Values on rows) one per column item, the value given by the row.
  const reportColumns = (splitByValue: boolean): ReportColumn[] => {
    const out: ReportColumn[] = [];
    const pushColumns = (make: (valueIndex: number | null) => ReportColumn) => {
      if (!splitByValue) out.push(make(null));
      else if (valueCount <= 1) out.push(make(valueCount ? 0 : null));
      else for (let index = 0; index < valueCount; index += 1) out.push(make(index));
    };
    if (columnFields.length === 0) {
      if (!splitByValue) out.push({ kind: "value", valueIndex: null });
      else for (let index = 0; index < valueCount; index += 1) out.push({ kind: "value", valueIndex: index });
      return out;
    }
    const leafDepth = columnFields.length - 1;
    const emitColumns = (nodes: AxisNode[], depth: number) => {
      for (const node of nodes) {
        if (depth === leafDepth || isColumnCollapsed(node, depth)) {
          pushColumns((valueIndex) => ({ kind: "item", path: node.path, valueIndex }));
          continue;
        }
        emitColumns(node.children, depth + 1);
        // A field's own Subtotals: None drops its column; Custom subtotals
        // on a column field show as one subtotal column here.
        const own = fieldSubtotalsByCol.get(columnFields[depth]);
        if (own ? own.length > 0 : subtotals) {
          pushColumns((valueIndex) => ({ kind: "subtotal", path: node.path, valueIndex }));
        }
      }
    };
    emitColumns(columnTree, 0);
    if (grandColumn) pushColumns((valueIndex) => ({ kind: "grand", valueIndex }));
    return out;
  };
  // A PivotChart plots value columns either way.
  const chartColumns = reportColumns(true);
  const columns = valuesOnRows ? reportColumns(false) : chartColumns;

  // Show Values As. A share of a total divides every cell, the subtotals and
  // grand totals included. Every other mode runs ALONG ITS BASE FIELD: the
  // cells it measures against are the ones differing only in that field's
  // item, so a cell naming no item of it stays blank (the axis' own grand
  // total always does), the other axis takes part in full, and a base field
  // inside a group restarts at every group. Measured against Excel on
  
  const siblingItems = (
    tree: AxisNode[],
    path: readonly string[],
    pos: number,
  ): string[] => {
    let nodes = tree;
    for (let depth = 0; depth < pos; depth += 1) {
      const next = nodes.find((node) => node.item === path[depth]);
      if (!next) return [];
      nodes = next.children;
    }
    return nodes.map((node) => node.item);
  };
  const looseItem = (a: string, b: string) =>
    a.trim().toLowerCase() === b.trim().toLowerCase();
  // A sample StdDev / Var over one number aggregates to NaN; the report, the
  // chart and GETPIVOTDATA read it as Excel's #DIV/0! (XLS-A071).
  const shown = (
    rowPath: readonly string[],
    column: ReportColumn,
    valueIndex: number,
  ): number | string | null => {
    const value = shownValue(rowPath, column, valueIndex);
    return typeof value === "number" && isPivotAggDivideByZero(value)
      ? PIVOT_SHOW_AS_DIVIDE_BY_ZERO
      : value;
  };
  function shownValue(
    rowPath: readonly string[],
    column: ReportColumn,
    valueIndex: number,
  ): number | string | null {
    const columnPath = column.kind === "item" || column.kind === "subtotal" ? column.path : [];
    const showAs = values[valueIndex].showAs;
    const mode = showAs?.mode ?? "normal";
    if (mode === "normal") return aggregate(rowPath, columnPath, valueIndex);
    if (!showValuesAsNeedsBaseField(mode)) {
      const raw = aggregate(rowPath, columnPath, valueIndex);
      if (mode === "index") {
        // How much of the cell is its own: (cell x grand) / (row x column).
        return showAsIndex(
          raw,
          aggregate(rowPath, [], valueIndex),
          aggregate([], columnPath, valueIndex),
          aggregate([], [], valueIndex),
        );
      }
      const total =
        mode === "percentOfGrandTotal"
          ? aggregate([], [], valueIndex)
          : mode === "percentOfRowTotal"
            ? aggregate(rowPath, [], valueIndex)
            : mode === "percentOfParentRowTotal"
              ? // The row one level out, which for a top-level row (and for
                // the Grand Total row itself) is the report's own total.
                aggregate(rowPath.slice(0, -1), columnPath, valueIndex)
              : mode === "percentOfParentColumnTotal"
                ? aggregate(rowPath, columnPath.slice(0, -1), valueIndex)
                : aggregate([], columnPath, valueIndex);
      return showAsShareOfTotal(raw, total);
    }
    const base = showAs?.base;
    const axis = base?.axis ?? (mode === "runningTotalInColumns" ? "col" : "row");
    if (showValuesAsNeedsBaseItem(mode) && (!base || base.axis !== axis)) return null;
    const axisFields = axis === "row" ? rowFields : columnFields;
    // A base field that is no longer on its axis leaves the whole value field
    // reading #N/A, the way Excel fills one whose base field is not in the
    // layout (S092 in the open-document corpus).
    if (base && base.pos >= axisFields.length) return PIVOT_SHOW_AS_NO_FIELD;
    // Without a named field the mode runs along the innermost field of its axis.
    const pos =
      base && base.pos < axisFields.length ? base.pos : axisFields.length - 1;
    if (pos < 0) return null;
    const path = axis === "row" ? rowPath : columnPath;
    if (path.length <= pos) return null; // this cell names no item of the field
    if (mode === "percentOfParentTotal") {
      // The parent of the base field: the total of the item this cell sits
      // under at that field's own level.
      const parent =
        axis === "row"
          ? aggregate(path.slice(0, pos + 1), columnPath, valueIndex)
          : aggregate(rowPath, path.slice(0, pos + 1), valueIndex);
      return showAsShareOfTotal(aggregate(rowPath, columnPath, valueIndex), parent);
    }
    const items = siblingItems(axis === "row" ? rowTree : columnTree, path, pos);
    const members = items.map((item) => {
      const memberPath = [...path.slice(0, pos), item, ...path.slice(pos + 1)];
      return axis === "row"
        ? aggregate(memberPath, columnPath, valueIndex)
        : aggregate(rowPath, memberPath, valueIndex);
    });
    return showValuesAsAlongBase(
      mode,
      members,
      items.indexOf(path[pos]),
      base?.item === undefined
        ? -1
        : items.findIndex((item) => looseItem(item, base.item ?? "")),
    );
  }

  const chartCategoryPaths: string[][] = [];
  if (!rowFields.length) {
    if (valueCount) chartCategoryPaths.push([]);
  } else {
    const walk = (nodes: AxisNode[], depth: number) => {
      for (const node of nodes) {
        if (depth === rowFields.length - 1 || isCollapsed(node, depth)) {
          chartCategoryPaths.push(node.path);
        } else {
          walk(node.children, depth + 1);
        }
      }
    };
    walk(rowTree, 0);
  }
  // A chart names an item as the table shows it: a field with a number
  // format (Excel's field settings, "KW "00 over a week number) formats the
  // item's label, on the category axis and in a series name alike.
  const fieldFormats = new Map(
    (config.fieldNumberFormats ?? []).map((entry) => [entry.colIndex, entry.numberFormat]),
  );
  const chartItemLabel = (item: string, field: number | undefined): string => {
    const label = itemLabel(item, field);
    const format = field === undefined ? undefined : fieldFormats.get(field);
    if (!format) return label;
    const numeric = Number(label);
    if (label.trim() === "" || !Number.isFinite(numeric)) return label;
    return applyNumberFormat(numeric, format);
  };
  const chartCategoryLabels = chartCategoryPaths.map((path) =>
    path.length ? path.map((item, depth) => chartItemLabel(item, rowFields[depth])) : [labels.total],
  );
  const fieldOf = (colIndex: number): PivotChartReportField => ({
    colIndex,
    caption: fieldLabel(colIndex),
  });
  const chart: PivotChartReport = {
    categories: chartCategoryLabels.map((path) => path[path.length - 1] ?? ""),
    categoryPaths: chartCategoryPaths,
    categoryLabels: chartCategoryLabels,
    series: chartColumns.flatMap((column): PivotChartReportSeries[] => {
      if ((column.kind !== "item" && column.kind !== "value") || column.valueIndex === null) {
        return [];
      }
      const valueIndex = column.valueIndex;
      const parts =
        column.kind === "item"
          ? column.path.map((item, depth) => chartItemLabel(item, columnFields[depth]))
          : [];
      if (valueCount > 1) parts.push(captions[valueIndex]);
      if (parts.length === 0) parts.push(labels.total);
      return [
        {
          name: parts.join(" - "),
          valueIndex,
          columnPath: column.kind === "item" ? [...column.path] : [],
          values: chartCategoryPaths.map((path) => {
            const value = shown(path, column, valueIndex);
            return typeof value === "number" && Number.isFinite(value) ? value : null;
          }),
        },
      ];
    }),
    rowFields: rowFields.map(fieldOf),
    columnFields: columnFields.map(fieldOf),
    filterFields: config.filters.map((filter) => fieldOf(filter.colIndex)),
    valueCaptions: captions,
  };
  if (!anchor) {
    return {
      status: "missing-anchor",
      output: null,
      pageRange: null,
      headerRows: 0,
      firstDataCol: 0,
      cells: new Map(),
      collisions: [],
      placement: null,
      resolvePivotData: emptyResolver(),
      chart,
    };
  }

  // Classic PivotTable layout keeps the column area's own header row with no
  // column field: the value field's caption heads the report and its column
  // reads "Total" over the data (Excel, S26_pivottable_date_field_filter;
  // the writer's _pivot_output_grid lays it out the same; X-007).
  const classicColumnRow =
    config.native?.classicLayout === true &&
    layout !== "compact" &&
    rowFields.length > 0 &&
    columnFields.length === 0 &&
    valueCount === 1;
  // Show the Values row cleared: with no column field the value fields'
  // captions head their columns in the one header row (X-266).
  const valuesRow =
    valueCount > 1 && !valuesOnRows && !(config.native?.hideValuesRow === true && columnFields.length === 0);
  const levels = columnFields.length + (valuesRow ? 1 : 0) + (classicColumnRow ? 1 : 0);
  const headerRows = levels + 1;
  // Without row fields a single value field's caption keeps a label column
  // beside its column items; several value fields head their own columns
  // (Excel: firstDataCol="0").
  // Σ Values on rows takes a column of its own outside compact form (its
  // captions), and the label column without row fields.
  const firstDataCol = rowFields.length
    ? compact
      ? 1
      : rowFields.length + (valuesOnRows ? 1 : 0)
    : valuesOnRows || (columnFields.length && valueCount === 1)
      ? 1
      : 0;
  const width = Math.max(1, firstDataCol + columns.length);
  const grid: PlannedCell[][] = [];
  const newRow = (
    labelRole: PivotOutputCellRole,
    dataRole: PivotOutputCellRole,
  ): PlannedCell[] =>
    Array.from({ length: width }, (_unused, col) => ({
      value: null,
      role: col < firstDataCol ? labelRole : dataRole,
    }));

  // Header rows: captions, the column items level by level, then the value
  // captions (or, without column fields or values, a single caption row).
  const header = Array.from({ length: headerRows }, () =>
    newRow("corner", "column-header"),
  );
  const last = headerRows - 1;
  const setHeader = (row: number, col: number, value: string, field?: number) => {
    if (header[row]?.[col]) {
      header[row][col].value = value;
      if (field !== undefined) header[row][col].field = field;
    }
  };
  if (rowFields.length) {
    if (compact) setHeader(last, 0, labels.rowLabels, rowFields[0]);
    else rowFields.forEach((field, index) => setHeader(last, index, fieldLabel(field), field));
  }
  if (valuesOnRows && !compact) {
    setHeader(last, rowFields.length, layout === "compact" ? labels.rowLabels : labels.values);
  }
  if (levels) {
    if (rowFields.length && valueCount === 1) setHeader(0, 0, captions[0]);
    // Compact form heads the column area "Column Labels" with or without
    // row fields (Excel-refreshed ClosedXML corpus K06/K24).
    const levelCaptions = compact || (layout === "compact" && columnFields.length > 0)
      ? [columnFields.length ? labels.columnLabels : labels.values]
      : [
          ...columnFields.map((field) => fieldLabel(field)),
          ...(valueCount > 1 && !valuesOnRows ? [labels.values] : []),
        ];
    levelCaptions.forEach((caption, index) =>
      setHeader(0, firstDataCol + index, caption, compact ? columnFields[0] : columnFields[index]),
    );
  }
  const openGroups: string[] = [];
  columns.forEach((column, index) => {
    const col = firstDataCol + index;
    const caption = column.valueIndex === null ? null : captions[column.valueIndex];
    if (column.kind === "value") {
      setHeader(last, col, classicColumnRow ? labels.total : caption ?? "");
      return;
    }
    if (column.kind === "item") {
      column.path.forEach((item, depth) => {
        const key = pathKey(column.path.slice(0, depth + 1));
        if (openGroups[depth] === key) return;
        openGroups.length = depth;
        openGroups[depth] = key;
        setHeader(depth + 1, col, itemLabel(item, columnFields[depth]));
        if (header[depth + 1]?.[col]) header[depth + 1][col].itemField = columnFields[depth];
      });
      if (valueCount > 1 && caption !== null) setHeader(last, col, caption);
      return;
    }
    openGroups.length = 0;
    if (column.kind === "subtotal") {
      const item = itemLabel(column.path[column.path.length - 1], columnFields[column.path.length - 1]);
      setHeader(
        column.path.length,
        col,
        valueCount > 1 && caption !== null
          ? labels.itemValue(item, caption)
          : labels.subtotal(item),
      );
      return;
    }
    setHeader(
      1,
      col,
      valueCount > 1 && caption !== null ? labels.valueTotal(caption) : labels.grandTotal,
    );
  });
  grid.push(...header);

  const dataRole = (
    column: ReportColumn,
    rowRole: "value" | "subtotal-value" | "grand-total-value",
  ): PivotOutputCellRole =>
    rowRole !== "value"
      ? rowRole
      : column.kind === "grand"
        ? "grand-total-value"
        : column.kind === "subtotal"
          ? "subtotal-value"
          : rowFields.length === 0 && column.kind === "value"
            ? "grand-total-value"
            : "value";
  // PivotTable Options: For error values show / For empty cells show.
  const reportValue = (value: number | string | null): number | string | null => {
    if (value === null) {
      if (config.emptyCaption === undefined) return null;
      // Excel reads the caption as a typed entry: "0" and "-" are the number 0.
      const text = config.emptyCaption.trim();
      if (text === "-") return 0;
      const number = text === "" ? Number.NaN : Number(text);
      return Number.isFinite(number) ? number : config.emptyCaption;
    }
    if (typeof value === "string" && config.errorCaption !== undefined) {
      return config.errorCaption === "" ? null : config.errorCaption;
    }
    return value;
  };
  const fillData = (
    cells: PlannedCell[],
    rowPath: readonly string[] | null,
    rowRole: "value" | "subtotal-value" | "grand-total-value",
    agg?: PivotAgg,
    // The row's value field, when Σ Values is on rows.
    rowValueIndex?: number,
  ) => {
    columns.forEach((column, index) => {
      const cell = cells[firstDataCol + index];
      cell.role = dataRole(column, rowRole);
      if (column.kind === "item" || column.kind === "subtotal") cell.columnPath = column.path;
      if (rowPath) cell.rowPath = rowPath;
      const valueIndex = column.valueIndex ?? rowValueIndex;
      if (valueIndex === undefined || valueIndex === null) return;
      cell.valueIndex = valueIndex;
      if (rowPath) {
        cell.value = reportValue(
          agg
            ? customSubtotal(rowPath, column, valueIndex, agg)
            : shown(rowPath, column, valueIndex),
        );
      }
    });
  };

  
  // X-294c): every real row item has a label row of its own (none in tabular
  // form, whose first value row carries the labels), each innermost item a
  // row per value field under it. Only the fields outside the innermost one
  // subtotal, below their group, a row per value ("East Price") whatever the
  // subtotal position; the grand total is a row per value ("Total Price").
  function emitValuesOnRows() {
    const leafDepth = rowFields.length - 1;
    const valuesCol = compact ? 0 : rowFields.length;
    let previous: string[] = [];
    const sharedWithPrevious = (members: readonly string[]) => {
      let shared = 0;
      while (shared < previous.length && shared < members.length && previous[shared] === members[shared]) {
        shared += 1;
      }
      return Math.min(shared, members.length - 1);
    };
    const valueMember = (index: number) => `${KEY_SEP}${index}`;
    const valueRows = (path: readonly string[], role: "value" | "subtotal-value") => {
      for (let index = 0; index < valueCount; index += 1) {
        const members = [...path, valueMember(index)];
        const shared = sharedWithPrevious(members);
        const cells = newRow("row-header", role);
        if (!compact && (tabular || repeatLabels)) {
          // Tabular form starts a block on its first value row; Repeat All
          // Item Labels writes the items on every row.
          const first = repeatLabels ? 0 : shared;
          for (let position = first; position < path.length; position += 1) {
            cells[position].value = itemLabel(path[position], rowFields[position]);
            if (position < leafDepth && position >= shared) {
              cells[position].drill = "expanded";
              cells[position].drillPath = pathKey(path.slice(0, position + 1));
              cells[position].drillField = rowFields[position];
              cells[position].drillItem = path[position];
            }
          }
        }
        cells[valuesCol].value = captions[index];
        if (compact && path.length) cells[valuesCol].indent = path.length;
        fillData(cells, path, role, undefined, index);
        grid.push(cells);
        previous = members;
      }
    };
    const emitNodes = (nodes: AxisNode[], depth: number) => {
      for (const node of nodes) {
        const leaf = depth === leafDepth;
        const collapsed = !leaf && isCollapsed(node, depth);
        if (!tabular || collapsed) {
          const shared = sharedWithPrevious(node.path);
          const cells = newRow("subtotal-label", "subtotal-value");
          const labelCell = cells[compact ? 0 : depth];
          if (compact) {
            labelCell.value = itemLabel(node.item, rowFields[depth]);
            if (depth) labelCell.indent = depth;
          } else {
            const first = repeatLabels ? 0 : tabular ? shared : depth;
            for (let position = first; position <= depth; position += 1) {
              cells[position].value = itemLabel(node.path[position], rowFields[position]);
            }
          }
          if (!leaf) {
            labelCell.drill = collapsed ? "collapsed" : "expanded";
            labelCell.drillPath = pathKey(node.path);
            labelCell.drillField = rowFields[depth];
            labelCell.drillItem = node.item;
          }
          fillData(cells, null, "subtotal-value");
          grid.push(cells);
          previous = node.path;
        }
        if (leaf || collapsed) {
          valueRows(node.path, leaf ? "value" : "subtotal-value");
          if (collapsed && config.blankRows === true) grid.push(newRow("blank", "blank"));
          continue;
        }
        emitNodes(node.children, depth + 1);
        const own = fieldSubtotalsByCol.get(rowFields[depth]);
        const subtotalRows = (agg: PivotAgg | undefined, name: string) => {
          for (let index = 0; index < valueCount; index += 1) {
            const cells = newRow("subtotal-label", "subtotal-value");
            const labelCell = cells[compact ? 0 : depth];
            labelCell.value = labels.itemValue(name, captions[index]);
            if (compact && depth) labelCell.indent = depth;
            fillData(cells, node.path, "subtotal-value", agg, index);
            grid.push(cells);
          }
          previous = node.path;
        };
        const item = itemLabel(node.item, rowFields[depth]);
        if (own && own.length > 0) {
          for (const agg of own) {
            subtotalRows(agg, `${item} ${(labels.subtotalFunction ?? pivotSubtotalFunctionLabel)(agg)}`);
          }
        } else if (own === undefined && subtotals) {
          subtotalRows(undefined, item);
        }
        if (config.blankRows === true) {
          grid.push(newRow("blank", "blank"));
          previous = node.path;
        }
      }
    };
    if (rowFields.length) emitNodes(rowTree, 0);
    else valueRows([], "value");
    if (grandRow && rowFields.length) {
      for (let index = 0; index < valueCount; index += 1) {
        const cells = newRow("grand-total-label", "grand-total-value");
        cells[0].value = labels.valueTotal(captions[index]);
        fillData(cells, [], "grand-total-value", undefined, index);
        grid.push(cells);
      }
    }
  }

  if (valuesOnRows) {
    emitValuesOnRows();
  } else if (!rowFields.length) {
    // One data row; a report without values has none.
    if (valueCount) {
      const cells = newRow("row-header", "value");
      if (firstDataCol) cells[0].value = captions[0];
      fillData(cells, [], "value");
      grid.push(cells);
    }
  } else {
    const leafDepth = rowFields.length - 1;
    let previous: string[] = [];
    const sharedWithPrevious = (path: readonly string[]) => {
      let shared = 0;
      while (
        shared < previous.length &&
        shared < path.length &&
        previous[shared] === path[shared]
      ) {
        shared += 1;
      }
      return Math.min(shared, path.length - 1);
    };
    const emitRows = (nodes: AxisNode[], depth: number) => {
      for (const node of nodes) {
        const leaf = depth === leafDepth;
        // Excel's minus button: a collapsed outer item is one row carrying
        // its totals, its inner rows and its subtotal row gone.
        const collapsed = !leaf && isCollapsed(node, depth);
        if (collapsed) {
          const shared = sharedWithPrevious(node.path);
          const cells = newRow("subtotal-label", "subtotal-value");
          const labelCell = cells[compact ? 0 : depth];
          if (compact) {
            labelCell.value = itemLabel(node.item, rowFields[depth]);
            if (depth) labelCell.indent = depth;
          } else {
            const first = repeatLabels ? 0 : tabular ? shared : depth;
            for (let position = first; position <= depth; position += 1) {
              cells[position].value = itemLabel(node.path[position], rowFields[position]);
            }
          }
          labelCell.drill = "collapsed";
          labelCell.drillPath = pathKey(node.path);
          labelCell.drillField = rowFields[depth];
          labelCell.drillItem = node.item;
          fillData(cells, node.path, "subtotal-value");
          grid.push(cells);
          previous = node.path;
          // Blank Rows > Insert Blank Line after Each Item follows every
          // outer item, a collapsed one included: Excel keeps the blank line
          // under an item whose details are hidden.
          if (config.blankRows === true) {
            grid.push(newRow("blank", "blank"));
          }
          continue;
        }
        if (leaf || !tabular) {
          // A row of its own: a leaf, or an outer item outside tabular form.
          const shared = sharedWithPrevious(node.path);
          const cells = leaf
            ? newRow("row-header", "value")
            : newRow("subtotal-label", "subtotal-value");
          if (compact) {
            cells[0].value = itemLabel(node.item, rowFields[depth]);
            if (depth) cells[0].indent = depth;
            if (!leaf) {
              cells[0].drill = "expanded";
              cells[0].drillPath = pathKey(node.path);
              cells[0].drillField = rowFields[depth];
              cells[0].drillItem = node.item;
            }
          } else {
            const first = repeatLabels ? 0 : tabular ? shared : depth;
            for (let position = first; position <= depth; position += 1) {
              cells[position].value = itemLabel(node.path[position], rowFields[position]);
              // An outer item's label (the first row of its block in tabular
              // form, its own row otherwise) carries the expand button.
              if (position < leafDepth && position >= shared) {
                cells[position].drill = "expanded";
                cells[position].drillPath = pathKey(node.path.slice(0, position + 1));
                cells[position].drillField = rowFields[position];
                cells[position].drillItem = node.path[position];
              }
            }
          }
          // An outer item's row carries its subtotal when subtotals are
          // shown at the top of the group: Automatic ones by the report's
          // setting, one Custom function by the field's own (`subtotalTop`,
          // X-311: CX100's Cake 9 / 3 / 12 on the Cake row).
          const ownSubtotals = leaf ? undefined : fieldSubtotalsByCol.get(rowFields[depth]);
          const customOnTop = ownSubtotalOnTop(rowFields[depth], ownSubtotals);
          fillData(
            cells,
            leaf || (subtotalsOnTop && ownSubtotals === undefined) || customOnTop ? node.path : null,
            leaf ? "value" : "subtotal-value",
            customOnTop ? ownSubtotals![0] : undefined,
          );
          grid.push(cells);
          previous = node.path;
        }
        if (leaf) continue;
        emitRows(node.children, depth + 1);
        const own = fieldSubtotalsByCol.get(rowFields[depth]);
        if (own && own.length > 0 && !ownSubtotalOnTop(rowFields[depth], own)) {
          // Custom subtotals: one row per function, below the group.
          for (const agg of own) {
            const cells = newRow("subtotal-label", "subtotal-value");
            const labelCell = cells[compact ? 0 : depth];
            labelCell.value = `${itemLabel(node.item, rowFields[depth])} ${(labels.subtotalFunction ?? pivotSubtotalFunctionLabel)(agg)}`;
            if (compact && depth) labelCell.indent = depth;
            fillData(cells, node.path, "subtotal-value", agg);
            grid.push(cells);
            previous = node.path;
          }
        } else if (own === undefined && subtotals && !subtotalsOnTop) {
          const cells = newRow("subtotal-label", "subtotal-value");
          const labelCell = cells[compact ? 0 : depth];
          labelCell.value = labels.subtotal(itemLabel(node.item, rowFields[depth]));
          if (compact && depth) labelCell.indent = depth;
          fillData(cells, node.path, "subtotal-value");
          grid.push(cells);
          previous = node.path;
        }
        if (config.blankRows === true) {
          grid.push(newRow("blank", "blank"));
          previous = node.path;
        }
      }
    };
    emitRows(rowTree, 0);
    if (grandRow) {
      const cells = newRow("grand-total-label", "grand-total-value");
      cells[0].value = labels.grandTotal;
      fillData(cells, [], "grand-total-value");
      grid.push(cells);
    }
  }

  const pageCount = config.filters.length;
  const reportRow = pageCount ? Math.max(anchor.r1, pageCount + 1) : anchor.r1;
  const output: PivotRange = {
    r1: reportRow,
    c1: anchor.c1,
    r2: reportRow + grid.length - 1,
    c2: anchor.c1 + width - 1,
  };
  const pageRange: PivotRange | null = pageCount
    ? { r1: reportRow - pageCount - 1, c1: anchor.c1, r2: reportRow - 2, c2: anchor.c1 + 1 }
    : null;
  const shape = { output, pageRange, headerRows, firstDataCol };
  if (
    options.rowCount < 1 ||
    options.columnCount < 1 ||
    output.r2 >= options.rowCount ||
    Math.max(output.c2, pageRange?.c2 ?? 0) >= options.columnCount
  ) {
    return {
      status: "out-of-bounds",
      ...shape,
      cells: new Map(),
      collisions: [],
      placement: null,
      resolvePivotData: emptyResolver(),
      chart,
    };
  }

  const cells = new Map<string, PivotOutputCell>();
  const collisions = new Set<string>();
  const place = (row: number, col: number, cell: PlannedCell) => {
    const key = cellKey(row, col);
    cells.set(key, {
      row,
      col,
      value: cell.value,
      role: cell.role,
      ...(cell.valueIndex === undefined ? {} : { valueIndex: cell.valueIndex }),
      ...(cell.indent ? { indent: cell.indent } : {}),
      ...(cell.field === undefined ? {} : { field: cell.field }),
      ...(cell.itemField === undefined ? {} : { itemField: cell.itemField }),
      ...(cell.rowPath ? { rowPath: cell.rowPath } : {}),
      ...(cell.columnPath ? { columnPath: cell.columnPath } : {}),
      ...(cell.drill
        ? {
            drill: cell.drill,
            drillPath: cell.drillPath,
            ...(cell.drillField === undefined
              ? {}
              : { drillField: cell.drillField, drillItem: cell.drillItem ?? "" }),
          }
        : {}),
    });
    if (options.isOccupied?.(row, col)) collisions.add(key);
  };
  config.filters.forEach((filter, index) => {
    const row = reportRow - pageCount - 1 + index;
    const selected = new Set(filter.selected);
    let shownItem: string;
    if (selected.size === 1) {
      shownItem = itemLabel([...selected][0], filter.colIndex);
    } else {
      const everyItemShown =
        selected.size === 0 ||
        sourceRows.every((source) => selected.has(rawValue(source, filter.colIndex)));
      shownItem = everyItemShown ? labels.allItems : labels.multipleItems;
    }
    place(row, anchor.c1, {
      value: fieldLabel(filter.colIndex),
      role: "page-field",
      field: filter.colIndex,
    });
    place(row, anchor.c1 + 1, { value: shownItem, role: "page-item", field: filter.colIndex });
  });
  grid.forEach((cellsInRow, rowOffset) =>
    cellsInRow.forEach((cell, colOffset) => {
      place(output.r1 + rowOffset, output.c1 + colOffset, cell);
      if (
        compact &&
        colOffset === 0 &&
        (cell.role === "row-header" || cell.role === "subtotal-label" || cell.role === "grand-total-label")
      ) {
        const placed = cells.get(cellKey(output.r1 + rowOffset, output.c1));
        if (placed) placed.compactLabel = true;
      }
    }),
  );
  const placement: PivotPlacement = { output, config, matrix };
  // GETPIVOTDATA reads the report AS SHOWN: the same aggregates, Show Values
  // As percentages and error cells the grid carries, a parent item's subtotal
  // included (live audit B128-B130). A criterion names a row or column field
  // by its caption; the named items must form a prefix of the axis (an inner
  // item without its outer one addresses no cell of the report).
  const resolvePivotData: PivotOutputPlan["resolvePivotData"] = (
    anchorCell,
    dataField,
    criteria,
  ) => {
    if (!rangeContains(output, anchorCell.row, anchorCell.col)) return { kind: "not-pivot" };
    const valueIndex = matchDataFieldIndex(matrix, config, dataField);
    if (valueIndex === null) return { kind: "no-match" };
    const rowItems: (string | undefined)[] = rowFields.map(() => undefined);
    const colItems: (string | undefined)[] = columnFields.map(() => undefined);
    for (const crit of criteria) {
      const rowPos = rowFields.findIndex((ci) => looseItem(headerLabel(matrix, ci), crit.field));
      if (rowPos >= 0) {
        rowItems[rowPos] = crit.item;
        continue;
      }
      const colPos = columnFields.findIndex((ci) => looseItem(headerLabel(matrix, ci), crit.field));
      if (colPos >= 0) {
        colItems[colPos] = crit.item;
        continue;
      }
      return { kind: "no-match" };
    }
    const pathOf = (items: readonly (string | undefined)[], tree: AxisNode[]): string[] | null => {
      const path: string[] = [];
      let nodes = tree;
      for (let depth = 0; depth < items.length; depth += 1) {
        const want = items[depth];
        if (want === undefined) {
          return items.slice(depth).some((item) => item !== undefined) ? null : path;
        }
        const node = nodes.find(
          (candidate) => looseItem(candidate.item, want) || looseItem(itemLabel(candidate.item), want),
        );
        if (!node) return null;
        path.push(node.item);
        nodes = node.children;
      }
      return path;
    };
    const rowPath = pathOf(rowItems, rowTree);
    const columnPath = pathOf(colItems, columnTree);
    if (!rowPath || !columnPath) return { kind: "no-match" };
    const column: ReportColumn =
      columnFields.length === 0
        ? { kind: "value", valueIndex }
        : columnPath.length === 0
          ? { kind: "grand", valueIndex }
          : columnPath.length === columnFields.length
            ? { kind: "item", path: columnPath, valueIndex }
            : { kind: "subtotal", path: columnPath, valueIndex };
    return { kind: "value", value: shown(rowPath, column, valueIndex) };
  };
  return {
    status: "ready",
    ...shape,
    cells,
    collisions: [...collisions],
    placement,
    resolvePivotData,
    chart,
  };
}
