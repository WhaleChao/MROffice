/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The one Office chart catalog shared by Docs, Presentation and Spreadsheet.
 *
 * Office ships a single Insert > Chart catalog for Word, Excel and PowerPoint:
 * the same 15 families and the same 61 leaves, in the same menu order, with the
 * same 2D/3D sectioning inside each family flyout. This module is that catalog
 * for MRAI. It is React-free on purpose so every host — ribbon component,
 * insert handler, validator, importer — reads the identical list instead of
 * keeping a private kind table that silently drifts.
 *
 * The chart MODEL (families, subtypes, data kinds, validation, defaults) stays
 * in `chartBlockShared`; this module re-exports the public surface and adds the
 * host-neutral helpers that the model itself has no opinion about: menu order,
 * per-host support, and the legacy `chartType` label PPT stores alongside V2.
 */
import {
  CHART_FAMILIES,
  CHART_SUBTYPES_BY_FAMILY,
  CHART_SUBTYPE_CATALOG,
  chartSubtypeDefinition,
  createChartBlockAttrsV2,
  defaultChartData,
  isChartSubtypeForFamily,
  validateChartBlockAttrsV2,
  type ChartBlockAttrsV2,
  type ChartData,
  type ChartDataKind,
  type ChartFamily,
  type ChartSubtype,
  type ChartSubtypeDefinition,
} from "../../components/console/editor/extensions/chartBlockShared";

export {
  CHART_FAMILIES,
  CHART_SUBTYPES_BY_FAMILY,
  CHART_SUBTYPE_CATALOG,
  chartSubtypeDefinition,
  createChartBlockAttrsV2,
  defaultChartData,
  isChartSubtypeForFamily,
  validateChartBlockAttrsV2,
};
export type {
  ChartBlockAttrsV2,
  ChartData,
  ChartDataKind,
  ChartFamily,
  ChartSubtype,
  ChartSubtypeDefinition,
};

/**
 * The families Office lists first in Insert > Chart, in Office's own order.
 * Word/Excel/PowerPoint all show this block above the divider.
 */
export const CHART_INSERT_OFFICE_FAMILIES = [
  "column",
  "line",
  "pie",
  "hierarchy",
  "statistical",
  "xy",
  "waterfall",
  "combo",
  "map",
] as const satisfies readonly ChartFamily[];

/** The remaining catalog families, below the divider. */
export const CHART_INSERT_ENHANCEMENT_FAMILIES = [
  "bar",
  "area",
  "stock",
  "radar",
  "funnel",
  "surface",
] as const satisfies readonly ChartFamily[];

/**
 * PowerPoint (and Word) for Mac's Insert > Chart menu: nine entries in
 * Office's order, each a flyout of the families Office files under it --
 * Bar under Column, Area under Line, and Funnel, Stock, Surface and Radar
 * under Waterfall (PowerPoint for Mac 16.86, the same nine as its Change
 * Chart Type cascade, PPT-A062). Every catalog family appears exactly once.
 */
export const OFFICE_MAC_CHART_MENU_GROUPS: readonly {
  readonly family: ChartFamily;
  readonly families: readonly ChartFamily[];
}[] = [
  { family: "column", families: ["column", "bar"] },
  { family: "line", families: ["line", "area"] },
  { family: "pie", families: ["pie"] },
  { family: "hierarchy", families: ["hierarchy"] },
  { family: "statistical", families: ["statistical"] },
  { family: "xy", families: ["xy"] },
  { family: "waterfall", families: ["waterfall", "funnel", "stock", "surface", "radar"] },
  { family: "combo", families: ["combo"] },
  { family: "map", families: ["map"] },
];

/** Every family in menu order. Exactly the 15 catalog families, no aliases. */
export const OFFICE_CHART_FAMILY_ORDER: readonly ChartFamily[] = [
  ...CHART_INSERT_OFFICE_FAMILIES,
  ...CHART_INSERT_ENHANCEMENT_FAMILIES,
];

export type ChartInsertSectionKind =
  | "twoDimensional"
  | "threeDimensional"
  | "specialized";

/**
 * Families Office presents as one undivided card list: they have no 3-D
 * counterpart, so a "2D" heading would be noise.
 */
const SPECIALIZED = new Set<ChartFamily>([
  "stock",
  "radar",
  "combo",
  "hierarchy",
  "statistical",
  "waterfall",
  "funnel",
  "map",
]);

/** The leaves Office files under the "3-D" heading of each family flyout. */
const THREE_DIMENSIONAL: Partial<Record<ChartFamily, readonly ChartSubtype[]>> =
  {
    column: ["clustered3d", "stacked3d", "percentStacked3d", "column3d"],
    bar: ["clustered3d", "stacked3d", "percentStacked3d"],
    line: ["line3d"],
    area: ["standard3d", "stacked3d", "percentStacked3d"],
    pie: ["pie3d"],
    xy: ["bubble3d"],
    surface: ["surface3d", "wireframe3d"],
  };

/** The sections of one family flyout, in the order Office renders them. */
export function chartInsertSections(family: ChartFamily): Array<{
  kind: ChartInsertSectionKind;
  subtypes: ChartSubtype[];
}> {
  const all = [...CHART_SUBTYPES_BY_FAMILY[family]] as ChartSubtype[];
  if (SPECIALIZED.has(family)) return [{ kind: "specialized", subtypes: all }];
  const three = [...(THREE_DIMENSIONAL[family] ?? [])];
  const two = all.filter((subtype) => !three.includes(subtype));
  return [
    ...(two.length ? [{ kind: "twoDimensional" as const, subtypes: two }] : []),
    ...(three.length
      ? [{ kind: "threeDimensional" as const, subtypes: three }]
      : []),
  ];
}










const WEBGL_LEAVES: Partial<Record<ChartFamily, readonly ChartSubtype[]>> = {};

export function chartLeafNeedsWebGL(family: ChartFamily, subtype: ChartSubtype): boolean {
  return (WEBGL_LEAVES[family] as readonly ChartSubtype[] | undefined)?.includes(subtype) ?? false;
}

/** Whether Office files this leaf under its family's "3-D" heading. */
export function isThreeDimensionalChartLeaf(
  family: ChartFamily,
  subtype: ChartSubtype,
): boolean {
  return (
    THREE_DIMENSIONAL[family] as readonly ChartSubtype[] | undefined
  )?.includes(subtype) ?? false;
}

export interface OfficeChartLeaf {
  family: ChartFamily;
  subtype: ChartSubtype;
}

export type OfficeChartHost = "docs" | "presentation" | "spreadsheet";

/**
 * The single per-host support table. A family is listed for a host only when
 * that host can insert it AND write it back losslessly; nothing degrades
 * silently into a different chart.
 *
 * - `docs` / `presentation`: the whole catalog. Classic leaves export through
 *   the shared `c:chartSpace` builder; hierarchy / statistical / waterfall /
 *   funnel / map export as chartEx with the classic fallback.
 * - `spreadsheet`: the whole catalog. A surface is plotted over the selected
 *   range's series (a series per row of the height matrix, Excel's own
 *   reading) and the XLSX writer states it as `c:surface3DChart` /
 *   `c:surfaceChart` bound to the sheet's formulas.
 */
const HOST_UNSUPPORTED_FAMILIES: Record<
  OfficeChartHost,
  ReadonlySet<ChartFamily>
> = {
  docs: new Set(),
  presentation: new Set(),
  spreadsheet: new Set<ChartFamily>(),
};

/**
 * Leaves a host cannot write back even though it offers their family, keyed
 * `family/subtype`. Support is per-LEAF and not only per-family because a host
 * writer can carry a family's default grammar and still have no vocabulary for
 * one of its variants:
 *
 * - `spreadsheet` binds combo series to sheet formulas. Stacked Area -
 *   Clustered Column is written by its subtype (every series but the last as
 *   stacked area, the last as a column); Custom Combination by each series'
 *   chart type and axis, asked in the shared Custom Combination pane before
 *   the chart is inserted.
 *
 * No host leaves a leaf out today; the table stays the one place to do so.
 */
const HOST_UNSUPPORTED_LEAVES: Record<OfficeChartHost, ReadonlySet<string>> = {
  docs: new Set(),
  presentation: new Set(),
  spreadsheet: new Set<string>(),
};

/**
 * Whether `host` can insert this family (and, when a subtype is given, this
 * exact leaf). Every menu, insert handler and coverage test reads this one
 * function, so a host never re-states its own support list.
 */
export function officeChartHostSupport(
  host: OfficeChartHost,
  family: ChartFamily,
  subtype?: ChartSubtype,
): boolean {
  if (
    !(CHART_FAMILIES as readonly ChartFamily[]).includes(family) ||
    HOST_UNSUPPORTED_FAMILIES[host].has(family)
  ) {
    return false;
  }
  if (subtype === undefined) return true;
  return (
    isChartSubtypeForFamily(family, subtype) &&
    !HOST_UNSUPPORTED_LEAVES[host].has(`${family}/${subtype}`)
  );
}

/** The families a host offers, in menu order. */
export function officeChartInsertFamilies(
  host?: OfficeChartHost,
): readonly ChartFamily[] {
  return host
    ? OFFICE_CHART_FAMILY_ORDER.filter((family) =>
        officeChartHostSupport(host, family),
      )
    : OFFICE_CHART_FAMILY_ORDER;
}

/**
 * Every insertable family/subtype pair in menu order: family order first, then
 * each family's flyout sections (2D before 3D). Callers that need "the whole
 * catalog" — insert galleries, parametrized tests, coverage audits — read this
 * instead of flattening `CHART_SUBTYPES_BY_FAMILY` by hand, so their order
 * matches what the user actually sees. With a `host`, leaves that host cannot
 * write back are dropped, so a coverage test reads exactly what it can prove.
 */
export function officeChartInsertLeaves(
  host?: OfficeChartHost,
): readonly OfficeChartLeaf[] {
  return officeChartInsertFamilies(host).flatMap((family) =>
    chartInsertSections(family).flatMap((section) =>
      section.subtypes
        .filter((subtype) => !host || officeChartHostSupport(host, family, subtype))
        .map((subtype) => ({ family, subtype })),
    ),
  );
}

/**
 * The ten leaves that carry a legacy PowerPoint kind name. These names predate
 * Chart V2 and are still what a PPT chart block stores in `chartType`, what
 * imported decks are matched against, and what the native chart fixtures are
 * filed under.
 *
 * MUST stay byte-identical to `KINDS` in
 * `apps/api/app/services/file_adapters/pptx/chart_v2.py`: the browser writes
 * `chartType` and the backend recomputes it from the same pair, so a divergence
 * is an invisible contract break rather than a test failure.
 */
export const OFFICE_CHART_LEGACY_KINDS = {
  column: ["column", "clustered"],
  bar: ["bar", "clustered"],
  line: ["line", "standardMarkers"],
  area: ["area", "standard"],
  pie: ["pie", "pie2d"],
  doughnut: ["pie", "doughnut"],
  stacked: ["column", "stacked"],
  radar: ["radar", "standard"],
  scatter: ["xy", "markers"],
  bubble: ["xy", "bubble"],
} as const satisfies Record<string, readonly [ChartFamily, ChartSubtype]>;

export type OfficeChartLegacyKind = keyof typeof OFFICE_CHART_LEGACY_KINDS;

/**
 * The label a PPT chart block stores in `chartType`. Mirrors the backend's
 * `legacy_kind(attrs)`: the legacy name when the pair has one, else the family
 * name. Every one of the 61 leaves therefore has a stable label.
 */
export function officeChartLegacyKind(
  family: ChartFamily,
  subtype: ChartSubtype,
): string {
  return (
    Object.entries(OFFICE_CHART_LEGACY_KINDS).find(
      ([, pair]) => pair[0] === family && pair[1] === subtype,
    )?.[0] ?? family
  );
}

/** The leaf a legacy kind name refers to, or null when the name is unknown. */
export function officeChartLeafForLegacyKind(
  kind: string,
): OfficeChartLeaf | null {
  const pair =
    OFFICE_CHART_LEGACY_KINDS[kind as OfficeChartLegacyKind] ?? undefined;
  return pair ? { family: pair[0], subtype: pair[1] } : null;
}

/** The default insert of the split button: Office's clustered column. */
export const OFFICE_CHART_DEFAULT_LEAF: OfficeChartLeaf = {
  family: "column",
  subtype: "clustered",
};
