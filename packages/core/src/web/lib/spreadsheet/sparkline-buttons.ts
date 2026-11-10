/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Insert > Sparklines and the Sparkline contextual tab, as pure functions over
 * a sheet's sparkline map (`sheet.sparklines`, keyed "row:col"). Moved out of
 * SpreadsheetWindow so the window's buttons and the spreadsheet editor
 * commands (lib/editor-commands/spreadsheet) run the same code; behaviour is
 * unchanged. Protection notices, error messages and the active cell stay with
 * the caller.
 */

import { chartRangeToA1, normalizeChartRange, type ChartRange } from "./chart-config";
import { shiftFormula } from "./formula-engine";
import { MAX_SPARKLINES, normalizeSparkline, type Sparkline, type SparklineColorField, type SparklineType } from "./sparkline";

export type SparklineMap = Readonly<Record<string, Sparkline>>;

interface SparklineSourceAssignment {
  key: string;
  range: string;
}

/**
 * Map one source rectangle onto the selected destination cells using Excel's
 * row/column dimensional contract. A single destination accepts one 1-D
 * series; a vertical destination receives one source row per cell; a
 * horizontal destination receives one source column per cell. A 2-D
 * destination or mismatched dimensions is rejected before any workbook write.
 */
export function sparklineSourceAssignments(
  source: ChartRange,
  location: ChartRange,
): SparklineSourceAssignment[] | null {
  const sourceRows = source.r2 - source.r1 + 1;
  const sourceColumns = source.c2 - source.c1 + 1;
  const locationRows = location.r2 - location.r1 + 1;
  const locationColumns = location.c2 - location.c1 + 1;

  if (locationRows === 1 && locationColumns === 1) {
    if (sourceRows > 1 && sourceColumns > 1) return null;
    return [
      {
        key: `${location.r1}:${location.c1}`,
        range: chartRangeToA1(source),
      },
    ];
  }

  if (locationColumns === 1 && locationRows === sourceRows) {
    return Array.from({ length: locationRows }, (_, index) => {
      const row = source.r1 + index;
      return {
        key: `${location.r1 + index}:${location.c1}`,
        range: chartRangeToA1({
          r1: row,
          c1: source.c1,
          r2: row,
          c2: source.c2,
        }),
      };
    });
  }

  if (locationRows === 1 && locationColumns === sourceColumns) {
    return Array.from({ length: locationColumns }, (_, index) => {
      const column = source.c1 + index;
      return {
        key: `${location.r1}:${location.c1 + index}`,
        range: chartRangeToA1({
          r1: source.r1,
          c1: column,
          r2: source.r2,
          c2: column,
        }),
      };
    });
  }

  return null;
}

/** Why Insert Sparklines wrote nothing. */
export type SparklineInsertRefusal = "range-invalid" | "dimensions" | "limit";

/**
 * Insert Sparklines: one sparkline of `type` per location cell, each bound to
 * its own row or column of the data range, sharing one new group. The whole
 * assignment is validated first, so a mismatch or the sheet's cap writes
 * nothing.
 */
export function sparklinesInserted(
  sparklines: SparklineMap,
  type: SparklineType,
  dataRange: string,
  location: ChartRange,
  /** The new group's id, given how many sparklines it holds. */
  groupIdFor: (count: number) => string,
  /** The colours a new group starts with: Excel gives it the default
   *  sparkline style (lib/spreadsheet/sparkline-styles). */
  colors?: Partial<Sparkline>,
): { sparklines: Record<string, Sparkline> } | SparklineInsertRefusal {
  const candidate = normalizeSparkline({ ...colors, type, range: dataRange });
  const source = candidate ? normalizeChartRange(candidate.range) : null;
  if (!candidate || !source) return "range-invalid";
  const assignments = sparklineSourceAssignments(source, location);
  if (!assignments) return "dimensions";
  const assignmentKeys = new Set(assignments.map((assignment) => assignment.key));
  const retainedCount = Object.keys(sparklines).filter((key) => !assignmentKeys.has(key)).length;
  if (retainedCount + assignments.length > MAX_SPARKLINES) return "limit";
  const next: Record<string, Sparkline> = { ...sparklines };
  const groupId = groupIdFor(assignments.length);
  for (const assignment of assignments) {
    next[assignment.key] = {
      ...candidate,
      groupId,
      range: assignment.range,
    };
  }
  return { sparklines: next };
}

/**
 * Clear Sparklines: every sparkline whose cell lies in the rectangle goes.
 * Null when there was none there.
 */
export function sparklinesCleared(sparklines: SparklineMap, rect: ChartRange): Record<string, Sparkline> | null {
  const next: Record<string, Sparkline> = {};
  let removed = 0;
  for (const [key, value] of Object.entries(sparklines)) {
    const [r, c] = key.split(":").map(Number);
    const inSel = r >= rect.r1 && r <= rect.r2 && c >= rect.c1 && c <= rect.c2;
    if (inSel) {
      removed += 1;
      continue;
    }
    next[key] = value;
  }
  return removed === 0 ? null : next;
}

/**
 * Why one sparkline - a single cell - cannot read `range`, or null when it
 * can. CT_Sparkline's data is one area from a single row or a single column
 * ([MS-XLSX] 2.6.9), the contract Insert Sparklines applies to a one-cell
 * location (sparklineSourceAssignments): a 2-D range is refused, never read
 * as its rows run together into one series (live audit B468). Only a group's
 * location range splits a 2-D source, one row or column per member.
 */
export function sparklineMemberRangeRefusal(
  range: string,
): Exclude<SparklineInsertRefusal, "limit"> | null {
  const source = normalizeChartRange(range);
  if (!source) return "range-invalid";
  return sparklineSourceAssignments(source, { r1: 0, c1: 0, r2: 0, c2: 0 })
    ? null
    : "dimensions";
}

/**
 * A Sparkline-tab edit of the sparkline at `key` alone (Edit Data > Single
 * Sparkline's Data). Null when the result is not a valid sparkline - a new
 * data range must be one a single sparkline can read
 * (sparklineMemberRangeRefusal) - and the same map when nothing changes.
 */
export function sparklinePatched(
  sparklines: SparklineMap,
  key: string,
  patch: Partial<Sparkline>,
): SparklineMap | null {
  const current = sparklines[key];
  if (!current) return null;
  const nextSparkline = normalizeSparkline({ ...current, ...patch });
  if (!nextSparkline || (nextSparkline.range !== "" && !normalizeChartRange(nextSparkline.range))) {
    return null;
  }
  if (
    nextSparkline.range !== "" &&
    nextSparkline.range !== current.range &&
    sparklineMemberRangeRefusal(nextSparkline.range) !== null
  ) {
    return null;
  }
  if (JSON.stringify(nextSparkline) === JSON.stringify(current)) return sparklines;
  return { ...sparklines, [key]: nextSparkline };
}

/**
 * A Sparkline-tab edit of the whole group the sparkline at `key` belongs to
 * (Type, Show, Style, Axis, Hidden & Empty Cells). Null when a member would
 * not be a valid sparkline; the same map when nothing changes.
 */
export function sparklineGroupPatched(
  sparklines: SparklineMap,
  key: string,
  patch: Partial<Sparkline>,
): SparklineMap | null {
  const current = sparklines[key];
  if (!current) return null;
  const memberKeys = current.groupId
    ? Object.entries(sparklines)
        .filter(([, candidate]) => candidate.groupId === current.groupId)
        .map(([memberKey]) => memberKey)
    : [key];
  const next = { ...sparklines };
  let changed = false;
  for (const memberKey of memberKeys) {
    const normalized = normalizeSparkline({ ...next[memberKey], ...patch });
    if (!normalized || (normalized.range !== "" && !normalizeChartRange(normalized.range))) {
      return null;
    }
    if (JSON.stringify(normalized) !== JSON.stringify(next[memberKey])) {
      next[memberKey] = normalized;
      changed = true;
    }
  }
  return changed ? next : sparklines;
}

/**
 * The colours of a sparkline group the Style group sets: Sparkline Color,
 * and Marker Color's Negative Points, Markers, High / Low / First / Last Point.
 */
export const SPARKLINE_COLOR_TARGET_FIELDS = {
  primary: "color",
  negative: "negativeColor",
  markers: "markerColor",
  high: "highColor",
  low: "lowColor",
  first: "firstColor",
  last: "lastColor",
} as const satisfies Record<string, SparklineColorField>;
export type SparklineColorTarget = keyof typeof SPARKLINE_COLOR_TARGET_FIELDS;

/** A sparkline's group: its groupId, or itself when it has none. */
function sparklineGroupKey(sparklines: SparklineMap, key: string): string {
  return sparklines[key]?.groupId ?? `cell:${key}`;
}

/**
 * Excel's Sparkline > Group: the selected sparklines join the active one's
 * group and take its settings (type, colours, points, axis, empty cells),
 * each keeping its own data range. Null when the selection holds fewer than
 * two groups, which Excel leaves the command off for.
 */
export function sparklinesGrouped(
  sparklines: SparklineMap,
  selectedKeys: readonly string[],
  activeKey: string,
  newGroupId: () => string,
): SparklineMap | null {
  const active = sparklines[activeKey];
  if (!active) return null;
  const keys = [...new Set([activeKey, ...selectedKeys])].filter((key) => sparklines[key]);
  if (new Set(keys.map((key) => sparklineGroupKey(sparklines, key))).size < 2) return null;
  const groupId = active.groupId ?? newGroupId();
  const next = { ...sparklines };
  for (const key of keys) {
    const joined = normalizeSparkline({ ...active, range: sparklines[key].range, groupId });
    if (!joined) return null;
    next[key] = joined;
  }
  return next;
}

/**
 * Excel's Sparkline > Ungroup: each group the selection touches splits into
 * one group per sparkline, every member keeping the settings it shows. Null
 * when no selected sparkline shares its group.
 */
export function sparklinesUngrouped(
  sparklines: SparklineMap,
  selectedKeys: readonly string[],
  newGroupId: () => string,
): SparklineMap | null {
  const members = new Map<string, string[]>();
  for (const [key, sparkline] of Object.entries(sparklines)) {
    if (!sparkline.groupId) continue;
    members.set(sparkline.groupId, [...(members.get(sparkline.groupId) ?? []), key]);
  }
  const split = new Set(
    selectedKeys
      .map((key) => sparklines[key]?.groupId)
      .filter((groupId): groupId is string => !!groupId && (members.get(groupId)?.length ?? 0) > 1),
  );
  if (split.size === 0) return null;
  const next = { ...sparklines };
  for (const groupId of split) {
    for (const key of members.get(groupId) ?? []) {
      next[key] = { ...sparklines[key], groupId: newGroupId() };
    }
  }
  return next;
}

/**
 * The sparkline a copied cell brings `rowDelta` rows and `colDelta` columns
 * away (Paste, Fill Down / Right, the fill handle). Excel copies a sparkline
 * "to other cells in a column or row ... by dragging or using Fill Down": its
 * data range moves like a relative reference in a copied formula ($ parts
 * stay), and the copy keeps every group setting, group identity included, so
 * it joins the source's group. A sparkline without a data source copies as
 * it is. Null when the moved range would leave the grid.
 */
export function sparklineCopiedBy(
  sparkline: Sparkline,
  rowDelta: number,
  colDelta: number,
): Sparkline | null {
  if (sparkline.range === "" || (rowDelta === 0 && colDelta === 0)) {
    return normalizeSparkline(JSON.parse(JSON.stringify(sparkline)));
  }
  const range = shiftFormula(sparkline.range, rowDelta, colDelta);
  if (range.includes("#REF!")) return null;
  const copied = normalizeSparkline({ ...JSON.parse(JSON.stringify(sparkline)), range });
  return copied && normalizeChartRange(copied.range) ? copied : null;
}

/** The sheet carrying `sparklines`; an empty map drops the key. */
export function sheetWithSparklines<S extends { sparklines?: unknown }>(sheet: S, sparklines: Record<string, Sparkline>): S {
  const nextSheet: S = { ...sheet };
  if (Object.keys(sparklines).length > 0) nextSheet.sparklines = sparklines;
  else delete nextSheet.sparklines;
  return nextSheet;
}
