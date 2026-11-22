/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for spreadsheet conditional formatting + cell hyperlinks
 * (Excel/WPS parity).
 *
 * Conditional formatting recolors a cell when its resolved value satisfies a
 * comparison rule (e.g. "> 100 → red background"). The sheet model stores the
 * rules as a bounded list; this module keeps the rule math — does a rule cover a
 * cell? does a value match a rule's comparison? how do we coerce a persisted
 * rule into a safe shape, and which style does a cell get? — out of the React
 * component so the renderer, the rule-builder UI and the persistence layer all
 * agree on the model and so each piece can be unit-tested in isolation.
 *
 * Rule model: each rule carries an inclusive 0-based range rectangle
 * ({ r1, c1, r2, c2 }, normalized so r1<=r2 and c1<=c2) plus a `kind`
 * discriminator (Excel parity). The kinds are:
 *   - "comparison" (DEFAULT): a numeric `operator` (>, <, >=, <=, =, !=) + a
 *     numeric `value` + its differential style (possibly none). A cell is recolored when the rule covers it and
 *     the cell's resolved numeric value matches. This is the ORIGINAL rule; a
 *     rule that OMITS `kind` is treated as comparison so old persisted data keeps
 *     evaluating identically, and the normalizer emits no `kind` for it so the
 *     stored shape is byte-for-byte unchanged.
 *   - "formula": a spreadsheet `formula` evaluated per covered cell (by the
 *     caller's formula engine, threaded in as a predicate); a truthy result
 *     applies its style.
 *   - "colorScale": a 2-/3-point gradient mapping the range's min/mid/max value
 *     to an interpolated background color.
 *   - "dataBar": a proportional in-cell bar for numeric values in the range.
 *   - "iconSet": a thresholded icon (traffic lights / arrows) per value bucket.
 *   - "textContains": a text predicate (contains / does-not-contain / begins- /
 *     ends-with a substring) on the cell's displayed text → color.
 *   - "duplicateValues": highlight values that are duplicate (or, inverted,
 *     unique) within the rule's range → color.
 *   - "topBottom": the top/bottom N items — or top/bottom N percent — of the
 *     numeric values in the rule's range → color.
 *   - "aboveBelowAverage": values above / below the range's average, optionally
 *     offset by k standard deviations → color.
 * Non-numeric cell values never match comparison / colorScale / dataBar /
 * iconSet / topBottom / aboveBelowAverage rules (like Excel's cell-value rules);
 * textContains acts on text and duplicateValues on the raw value; formula rules
 * decide via their own predicate.
 *
 * DEFERRED (not yet implemented): Excel's "A Date Occurring…" (today / this week /
 * last month / …) and the blanks / no-blanks / errors / no-errors highlight
 * rules. Date-occurring needs a serial-date model + a "today" clock threaded in;
 * the blanks/errors variants need a blank-vs-error cell distinction the renderer
 * does not yet surface here. Both can be added later as further `kind`s.
 *
 * Conflict policy: array order is native priority. Matching lower-priority
 * rules may fill visual channels not already supplied by a higher rule, but do
 * not overwrite higher-priority channels. A matching `stopIfTrue` rule ends
 * evaluation, matching Excel's Manage Rules contract.
 *
 * Colors are ALWAYS funneled through normalizeHexColor so a persisted/dialog
 * color can never inject arbitrary CSS — only a literal #rrggbb survives.
 */

import { excelMaxDateSerial, excelSystemSerialToParts } from "./excel-1900-calendar";
import pageEnvelope from "./spreadsheet-page-envelope.shared.json";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "./grid-sizing";
import {
  BORDER_EDGES,
  BORDER_LINE_STYLES,
  type BorderLineStyle,
  type CellBorderEdges,
  type RichBorderEdge,
} from "./cell-format";
import {
  isSpreadsheetFillPatternType,
  normalizeSpreadsheetGradientFill,
  spreadsheetFillPatternCss,
  spreadsheetGradientFillCss,
  type SpreadsheetFillPatternType,
  type SpreadsheetGradientFill,
} from "./fill-pattern";

function isBorderLineStyle(value: unknown): value is BorderLineStyle {
  return (
    typeof value === "string" &&
    (BORDER_LINE_STYLES as readonly string[]).includes(value)
  );
}
import { excelWildcardToRegexSource, lettersToColumn, shiftFormula } from "./formula-engine";
import { formulaReferenceSpans } from "./formula-reference-highlight";
import { type ExcelDateSystemArg } from "./excel-1900-calendar";
import { excelMinDateSerial } from "./excel-1900-calendar";

/** The comparison operators a conditional-format rule may use. */
export type ConditionalOperator =
  | ">"
  | "<"
  | ">="
  | "<="
  | "="
  | "!="
  | "between"
  | "notBetween";

const OPERATORS: readonly ConditionalOperator[] = [
  ">",
  "<",
  ">=",
  "<=",
  "=",
  "!=",
  "between",
  "notBetween",
];

/**
 * The differential members a rule carries beyond the two colours the editing
 * panel has controls for; the panel keeps them on the rule across an edit.
 */
/**
 * Whether a rule's fill replaces the cell's own pattern fill. A rule that
 * only colours the background (a dxf patternFill without a patternType, as
 * Excel writes its fills) changes the colour alone and the cell's hatch stays
 * over it (PH122_issue.4248's darkDown "Non" cells on pink, X-241); a rule's
 * own pattern or gradient (its background image) replaces it.
 */
export function conditionalFillReplacesCellPattern(
  ruleStyle: { backgroundColor?: unknown; backgroundImage?: unknown },
  cellHasPattern: boolean,
): boolean {
  if (!ruleStyle.backgroundColor) return false;
  return ruleStyle.backgroundImage !== undefined || !cellHasPattern;
}

export function conditionalStyleExtras(
  rule: ConditionalFormatRule,
): Partial<ConditionalDifferentialStyle> {
  const source = rule as Partial<ConditionalDifferentialStyle>;
  const extras: Partial<ConditionalDifferentialStyle> = {};
  for (const key of [
    "textThemeColor",
    "backgroundThemeColor",
    "bold",
    "italic",
    "strikethrough",
    "underline",
    "fontSize",
    "fontName",
    "fontExtras",
    "numberFormat",
    "borderEdges",
    "gradientFill",
    "fillPattern",
  ] as const) {
    const value = source[key];
    if (value !== undefined) (extras as Record<string, unknown>)[key] = value;
  }
  return extras;
}

/** Excel's two-threshold operators: inclusive between, and its negation. */
export function isRangeConditionalOperator(
  operator: ConditionalOperator,
): operator is "between" | "notBetween" {
  return operator === "between" || operator === "notBetween";
}

/** Excel's dxf underline: single, double, or an explicit "none" that turns
 * an inherited underline off. */
export type ConditionalUnderline = "single" | "double" | "none";

/**
 * The differential style a style-kind rule paints when it matches, mirroring
 * Excel's Format Cells dialog for a rule. Font emphasis flags are explicit:
 * `true` turns the emphasis on, `false` (`<b val="0"/>`) turns an inherited
 * one off, absent leaves the cell's own. Font size / name and the remaining
 * dxf font members Excel accepts but does not expose are carried so a
 * workbook from another producer keeps them across an edit.
 */
export interface ConditionalDifferentialStyle {
  textColor?: string;
  /** The workbook theme slot `textColor` came from; written back as one
   * while its rgb still equals the carried colour. */
  textThemeColor?: ConditionalThemeColor;
  backgroundColor?: string;
  backgroundThemeColor?: ConditionalThemeColor;
  bold?: boolean;
  italic?: boolean;
  strikethrough?: boolean;
  underline?: ConditionalUnderline;
  /** Points, 1..409. */
  fontSize?: number;
  fontName?: string;
  fontExtras?: Record<string, string>;
  /** An Excel number format code applied while the rule matches. */
  numberFormat?: string;
  borderEdges?: CellBorderEdges;
  gradientFill?: SpreadsheetGradientFill;
  /** A hatched / dotted fill (Format Cells > Fill > Pattern Style): its lines
   * in `foregroundColor` (absent: Excel's automatic black) over the rule's
   * `backgroundColor`, or over the cell's own fill when the rule states none
   * (P07_FillWithoutColor, X-081). */
  fillPattern?: { type: SpreadsheetFillPatternType; foregroundColor?: string };
  /** Format Cells > Fill > No Color: matched cells show no fill at all. */
  noFill?: boolean;
}

export const CONDITIONAL_FONT_EXTRA_MEMBERS = [
  "charset",
  "family",
  "scheme",
  "condense",
  "extend",
  "outline",
  "shadow",
  "vertAlign",
] as const;

/** A rectangular range in inclusive, 0-based grid coordinates. */
export interface ConditionalRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/**
 * Discriminator for the conditional-format rule kinds (Excel parity):
 *  - "comparison": numeric cell-value comparison (the ORIGINAL rule; default).
 *  - "formula":    a spreadsheet formula evaluated per covered cell; a truthy
 *                  result applies the rule's color(s).
 *  - "colorScale": a 2-/3-point gradient mapping the range's min/mid/max value
 *                  to interpolated background colors.
 *  - "dataBar":    an in-cell proportional bar for numeric values in the range.
 *  - "iconSet":    a thresholded icon (traffic lights / arrows) per value bucket.
 *  - "textContains":     a per-cell text predicate (contains / not-contains /
 *                  begins-/ends-with) → color (Excel "Highlight Cells Rules").
 *  - "duplicateValues":  highlight duplicate (or unique) values within the range.
 *  - "topBottom":        top/bottom N items — or N percent — of the range's
 *                  numeric values (Excel "Top/Bottom Rules").
 *  - "aboveBelowAverage": above/below the range's average, ±k std-devs.
 *
 * A rule that omits `kind` is treated as "comparison" so OLD persisted rules
 * (which never carried a kind) keep evaluating identically.
 */
export type ConditionalRuleKind =
  | "comparison"
  | "formula"
  | "colorScale"
  | "dataBar"
  | "iconSet"
  | "textContains"
  | "duplicateValues"
  | "topBottom"
  | "aboveBelowAverage"
  | "dateOccurring";

const RULE_KINDS: readonly ConditionalRuleKind[] = [
  "comparison",
  "formula",
  "colorScale",
  "dataBar",
  "iconSet",
  "textContains",
  "duplicateValues",
  "topBottom",
  "aboveBelowAverage",
  "dateOccurring",
];

/**
 * A numeric cell-value comparison rule — the ORIGINAL conditional-format rule.
 * The `kind` field is OPTIONAL and, when present, is always "comparison"; it is
 * omitted by the normalizer for comparison rules so the persisted shape is
 * byte-for-byte identical to pre-union data (back-compat). A raw rule with no
 * `kind` normalizes to exactly this shape.
 */
export interface ConditionalComparisonRule extends ConditionalDifferentialStyle {
  kind?: "comparison";
  range: ConditionalRange;
  operator: ConditionalOperator;
  value: number;
  /** Upper threshold of a between / notBetween rule (inclusive). */
  value2?: number;
  /** Validated #rrggbb text color applied when a covered cell matches. */
  textColor?: string;
  /** Validated #rrggbb background color applied when a covered cell matches. */
  backgroundColor?: string;
}

/**
 * A formula rule: the `formula` is evaluated per covered cell (by the caller's
 * formula engine, threaded in as a predicate) and the color(s) apply when the
 * result is truthy. References in the formula are relative to the range's
 * top-left anchor (Excel parity) — the renderer shifts them per cell. At least
 * one of textColor / backgroundColor is required (same as comparison).
 */
export interface ConditionalFormulaRule extends ConditionalDifferentialStyle {
  kind: "formula";
  range: ConditionalRange;
  /** A spreadsheet formula (with or without a leading "="); truthy → apply. */
  formula: string;
  textColor?: string;
  backgroundColor?: string;
  /**
   * Import-only: the workbook holds this rule as Excel's "Cell Value" rule
   * (cellIs) against text, a reference or a name, read as the expression it
   * evaluates. The xlsx writer puts it back as cellIs while the formula keeps
   * that shape.
   */
  nativeCellIs?: boolean;
  /** Read from Excel's containsErrors / notContainsErrors rule: the writer
   *  puts the predicate back in that family while it still is one. */
  nativeErrorRule?: boolean;
}

/**
 * One stop of a color-scale gradient: a validated #rrggbb color plus its
 * threshold family (SpreadsheetML CT_Cfvo). An untyped first stop is the range
 * minimum, an untyped last stop the maximum, and an untyped midpoint the 50th
 * percentile. `formula` stops keep their expression; without an evaluator the
 * renderer treats them as the endpoint they replace.
 */
export interface ColorScaleStop {
  color: string;
  /** The workbook theme slot this colour came from; written back as one. */
  themeColor?: ConditionalThemeColor;
  type?: "min" | "max" | "percent" | "percentile" | "number" | "formula";
  /** Threshold value for percent / percentile / number stops. */
  value?: number;
  /** Threshold expression for formula stops (no leading "="). */
  formula?: string;
}

/**
 * A color-scale rule: a 2- or 3-point gradient. The range's MIN value maps to
 * `min.color`, its MAX to `max.color`, and (for a 3-point scale) its MID — the
 * midpoint between min and max — to `mid.color`. Each covered numeric cell gets
 * an interpolated background color. Non-numeric cells are skipped.
 */
export interface ConditionalColorScaleRule {
  kind: "colorScale";
  range: ConditionalRange;
  min: ColorScaleStop;
  /** Present for a 3-point scale; omitted for a 2-point scale. */
  mid?: ColorScaleStop;
  max: ColorScaleStop;
}

/**
 * A data-bar rule: each covered numeric cell renders an in-cell horizontal bar
 * whose width is proportional to where the value sits between the range's min
 * and max (min → 0%, max → 100%). `color` is the validated #rrggbb bar fill.
 */
export interface ConditionalDataBarRule {
  kind: "dataBar";
  range: ConditionalRange;
  /** Validated #rrggbb bar fill color. */
  color: string;
  /** Classic SpreadsheetML defaults are 10 and 90 percent. */
  minLength?: number;
  maxLength?: number;
  /** false mirrors classic dataBar showValue="0". */
  showValue?: boolean;
  /** x14 data-bar fill treatment. Omitted for classic-only rules. */
  gradient?: boolean;
  /** Excel 2010+ look, as the projection read it: a negative value's bar
   * colour (absent when it keeps the positive colour), the zero axis colour
   * (absent when the axis is off), a bordered bar's border, the axis placement. */
  negativeColor?: string;
  axisColor?: string;
  borderColor?: string;
  axisPosition?: "automatic" | "middle" | "none";
  /** Authenticated x14 rule identity retained for original-package patchback. */
  nativeDataBarId?: string;
  /** The workbook theme slot the bar colour came from; written back as one. */
  themeColor?: ConditionalThemeColor;
  /**
   * Shortest / longest bar thresholds (CT_Cfvo). Omitted endpoints are the
   * range minimum and maximum; the same families as colour-scale stops.
   */
  min?: DataBarThreshold;
  max?: DataBarThreshold;
}

/**
 * A data-bar endpoint: the family and value of one cfvo, without a colour.
 * `automatic` is Excel 2010's default (x14 autoMin / autoMax): the shortest
 * bar stands for the smaller of zero and the lowest value, the longest for
 * the larger of zero and the highest, read from the data every time it is
 * drawn, so a value that crosses zero later still gets its axis and negative
 * bar (live audit B247). A pinned `number` 0 is the user's own choice and
 * stays what it says.
 */
export interface DataBarThreshold {
  type?: "min" | "max" | "percent" | "percentile" | "number" | "formula" | "automatic";
  value?: number;
  formula?: string;
}

/** A theme colour reference: slot 0-11, optional tint, and the resolved rgb. */
export interface ConditionalThemeColor {
  theme: number;
  tint?: number;
  rgb: string;
}

function normalizeConditionalThemeColor(raw: unknown): ConditionalThemeColor | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const r = raw as Record<string, unknown>;
  const theme = Number(r.theme);
  const rgb = normalizeHexColor(r.rgb);
  if (!Number.isInteger(theme) || theme < 0 || theme > 11 || !rgb) return undefined;
  const tint = Number(r.tint);
  return {
    theme,
    ...(Number.isFinite(tint) && tint !== 0 && Math.abs(tint) <= 1 ? { tint } : {}),
    rgb,
  };
}

/**
 * The built-in icon-set families (Excel parity), rendered as inline SVG (no
 * deps). Each family has a FIXED icon count of 3, 4, or 5; that count drives how
 * many buckets the value→icon mapping produces (an N-icon set uses N-1 thresholds
 * → N buckets). The two original ids ("3-traffic-lights", "3-arrows") are kept as
 * the first two 3-icon families so OLD persisted rules and presets are unchanged.
 *
 * 3-icon: 3 Arrows (colored / gray), 3 Triangles, 3 Traffic Lights (unrimmed /
 *   rimmed), 3 Signs, 3 Symbols (circled / uncircled), 3 Flags, 3 Stars.
 * 4-icon: 4 Arrows (colored / gray), 4 Red-to-Black, 4 Traffic Lights, 4 Ratings.
 * 5-icon: 5 Arrows (colored / gray), 5 Ratings, 5 Quarters, 5 Boxes.
 */
export type IconSetName =
  // 3-icon families (the first two are the ORIGINAL ids, unchanged).
  | "3-traffic-lights"
  | "3-arrows"
  | "3-arrows-gray"
  | "3-triangles"
  | "3-traffic-lights-rimmed"
  | "3-signs"
  | "3-symbols-circled"
  | "3-symbols"
  | "3-flags"
  | "3-stars"
  // 4-icon families.
  | "4-arrows"
  | "4-arrows-gray"
  | "4-red-to-black"
  | "4-traffic-lights"
  | "4-ratings"
  // 5-icon families.
  | "5-arrows"
  | "5-arrows-gray"
  | "5-ratings"
  | "5-quarters"
  | "5-boxes";

/**
 * Every icon-set family with its fixed icon count (3, 4 or 5). The map is the
 * single source of truth for which families exist and how many buckets each
 * produces; the renderer, normalizer, default-threshold helper and bucket math
 * all read it. Declaration order is the gallery order shown in the rule dialog.
 */
export const ICON_SET_FAMILIES: ReadonlyArray<{
  name: IconSetName;
  count: 3 | 4 | 5;
}> = [
  { name: "3-traffic-lights", count: 3 },
  { name: "3-arrows", count: 3 },
  { name: "3-arrows-gray", count: 3 },
  { name: "3-triangles", count: 3 },
  { name: "3-traffic-lights-rimmed", count: 3 },
  { name: "3-signs", count: 3 },
  { name: "3-symbols-circled", count: 3 },
  { name: "3-symbols", count: 3 },
  { name: "3-flags", count: 3 },
  { name: "3-stars", count: 3 },
  { name: "4-arrows", count: 4 },
  { name: "4-arrows-gray", count: 4 },
  { name: "4-red-to-black", count: 4 },
  { name: "4-traffic-lights", count: 4 },
  { name: "4-ratings", count: 4 },
  { name: "5-arrows", count: 5 },
  { name: "5-arrows-gray", count: 5 },
  { name: "5-ratings", count: 5 },
  { name: "5-quarters", count: 5 },
  { name: "5-boxes", count: 5 },
];

const ICON_SET_COUNT_BY_NAME: ReadonlyMap<IconSetName, 3 | 4 | 5> = new Map(
  ICON_SET_FAMILIES.map((f) => [f.name, f.count] as const),
);

const ICON_SET_NAMES: readonly IconSetName[] = ICON_SET_FAMILIES.map(
  (f) => f.name,
);

/** True when a candidate string names a built-in icon-set family. */
export function isIconSetName(value: unknown): value is IconSetName {
  return (
    typeof value === "string" &&
    (ICON_SET_NAMES as readonly string[]).includes(value)
  );
}

/**
 * How many icons (= buckets) a family has: 3, 4 or 5. An unknown name falls back
 * to 3 (the original count) so a forward-incompatible persisted name still maps
 * to a sensible bucket count rather than throwing.
 */
export function iconSetCount(name: IconSetName): 3 | 4 | 5 {
  return ICON_SET_COUNT_BY_NAME.get(name) ?? 3;
}

/**
 * The Excel default percent thresholds for an N-icon set as N-1 sorted fractions
 * in (0,1): a 3-icon set splits at 1/3 & 2/3, a 4-icon set at 1/4, 2/4 & 3/4, a
 * 5-icon set at 1/5 … 4/5. Returned fresh each call so callers can mutate freely.
 */
export function defaultIconThresholds(count: 3 | 4 | 5): number[] {
  const out: number[] = [];
  for (let i = 1; i < count; i += 1) out.push(i / count);
  return out;
}

/**
 * An icon-set rule: each covered numeric cell shows one icon from the chosen
 * `iconSet` family, picked by value bucket. The value's fraction within the
 * range's [min, max] span is compared against the family's N-1 ascending
 * `thresholds` (fractions in (0,1)): below threshold 0 → icon 0, below threshold
 * 1 → icon 1, …, otherwise → the top icon. `thresholds` always has exactly
 * (iconSetCount(iconSet) - 1) entries (the normalizer enforces this and fills
 * Excel defaults when missing/wrong-length). `reverse` flips which icon a bucket
 * gets (Excel "Reverse Icon Order"): bucket b → icon (count-1-b). `showIconOnly`
 * hides the cell's value text, showing only the icon (Excel "Show Icon Only").
 */
/** One slot of an icon set whose pictures were picked per slot (Excel's
 * "Custom" icon style): any family's icon, or no icon at all. */
export interface ConditionalCustomIcon {
  iconSet: IconSetName | "no-icons";
  /** Index within that family; 0 for "no-icons". */
  iconId: number;
}

/** The families a typed icon boundary may belong to (Excel's Type column). */
export type IconThresholdType = "number" | "percent" | "percentile" | "formula";

/**
 * One typed bucket boundary of an icon set, Excel's Value / Type / ">=" columns
 * for every icon but the lowest. `value` carries a number, a percent (0..100)
 * or a percentile (0..100); `formula` carries a formula without its "=".
 * `gte: false` starts the bucket strictly above the boundary (Excel ">");
 * omitted, the boundary itself belongs to the bucket (">=", Excel's default).
 */
export interface IconThreshold {
  type: IconThresholdType;
  value?: number;
  formula?: string;
  gte?: false;
}

export interface ConditionalIconSetRule {
  kind: "iconSet";
  range: ConditionalRange;
  iconSet: IconSetName;
  /**
   * The (count-1) bucket boundaries as fractions of the range span (0..1),
   * sorted ascending. Defaults to {@link defaultIconThresholds} for the family's
   * icon count when omitted/invalid/wrong-length.
   */
  thresholds: number[];
  /**
   * The (count-1) typed boundaries, ascending by bucket, when any of them is
   * not a plain ">=" percent (a number, a percentile, a formula, or a ">"
   * boundary). Authoritative over `thresholds`, which then mirrors the
   * percent boundaries and holds the family default for the others.
   */
  stops?: IconThreshold[];
  /** true → reverse icon order (highest value gets icon 0). Omitted = false. */
  reverse?: boolean;
  /** true → show only the icon, hiding the cell value. Omitted = false. */
  showIconOnly?: boolean;
  /**
   * Per-slot custom icons (Excel 2010+). The rule then lives in the
   * workbook's x14 extension, as do the 3 Triangles, 3 Stars and 5 Boxes
   * families; `nativeX14Id` is its extension identity from the package.
   */
  customIcons?: ConditionalCustomIcon[];
  nativeX14Id?: string;
}

/** The text-match operators a textContains rule may use. */
export type TextMatchOperator =
  | "contains"
  | "notContains"
  | "beginsWith"
  | "endsWith";

const TEXT_MATCH_OPERATORS: readonly TextMatchOperator[] = [
  "contains",
  "notContains",
  "beginsWith",
  "endsWith",
];

/**
 * A text-match rule (Excel "Highlight Cells Rules" → Text that Contains): the
 * cell's displayed text is tested against `text` with `operator` (case-
 * insensitive, Excel parity) and the color(s) apply on a match. `notContains`
 * matches when the substring is ABSENT (every covered cell, including blanks,
 * since "" does not contain a non-empty needle). At least one of textColor /
 * backgroundColor is required (same as comparison).
 */
export interface ConditionalTextContainsRule extends ConditionalDifferentialStyle {
  kind: "textContains";
  range: ConditionalRange;
  /** The (non-empty) substring to test the cell text against. */
  text: string;
  /** How to test: contains / does-not-contain / begins-with / ends-with. */
  operator: TextMatchOperator;
  








  formula?: string;
  textColor?: string;
  backgroundColor?: string;
}

/** The formula a rule is evaluated by and that moves with it: a formula
 *  rule's, or a text rule's own (`ConditionalTextContainsRule.formula`). */
export function conditionalRuleFormula(rule: ConditionalFormatRule): string | undefined {
  if (rule.kind === "formula") return rule.formula;
  if (rule.kind === "textContains") return rule.formula;
  return undefined;
}

/**
 * A duplicate/unique-values rule (Excel "Highlight Cells Rules" → Duplicate
 * Values): every covered cell whose value occurs MORE than once in the range is
 * highlighted (`unique: false`, the default). When `unique` is true the polarity
 * inverts — cells whose value occurs EXACTLY once are highlighted. Values are
 * compared by their trimmed, case-insensitive text (Excel treats "A"/"a" and
 * 1/"1" as equal); blank cells are ignored (never duplicates). At least one
 * color is required.
 */
export interface ConditionalDuplicateValuesRule extends ConditionalDifferentialStyle {
  kind: "duplicateValues";
  range: ConditionalRange;
  /** true → highlight UNIQUE values; false/omitted → highlight DUPLICATES. */
  unique?: boolean;
  textColor?: string;
  backgroundColor?: string;
}

/**
 * A top/bottom rule (Excel "Top/Bottom Rules"): highlight the N highest
 * (`bottom: false`, the default) or N lowest (`bottom: true`) numeric values in
 * the range. When `percent` is true, N is read as a PERCENTAGE of the numeric
 * cell count (Excel's "Top 10%"), rounded down to at least one cell. Ties are
 * inclusive — every cell at or beyond the cutoff value is highlighted, so a
 * "top 2" over [9, 9, 1] highlights both 9s. Non-numeric cells never match. At
 * least one color is required.
 */
export interface ConditionalTopBottomRule extends ConditionalDifferentialStyle {
  kind: "topBottom";
  range: ConditionalRange;
  /** How many items (or percent) to highlight; clamped to ≥1. */
  rank: number;
  /** true → BOTTOM N (lowest); false/omitted → TOP N (highest). */
  bottom?: boolean;
  /** true → `rank` is a percentage of the numeric cell count. */
  percent?: boolean;
  textColor?: string;
  backgroundColor?: string;
}

/**
 * An above/below-average rule (Excel "Top/Bottom Rules" → Above/Below Average):
 * highlight numeric cells above (`below: false`, default) or below
 * (`below: true`) the range's mean. An optional `stdDevs` (k) shifts the
 * threshold by k standard deviations of the range — above-average uses
 * mean + k·σ, below-average uses mean − k·σ (Excel's "1 std dev above"). The
 * comparison is strict (a value equal to the threshold does not match). Non-
 * numeric cells never match. At least one color is required.
 */
export interface ConditionalAboveBelowAverageRule extends ConditionalDifferentialStyle {
  kind: "aboveBelowAverage";
  range: ConditionalRange;
  /** true → BELOW average; false/omitted → ABOVE average. */
  below?: boolean;
  /** Excel's "equal or above / below average": the mean itself matches. */
  equalAverage?: boolean;
  /** Standard-deviation offset k (≥0); omitted/0 → the plain mean. */
  stdDevs?: number;
  textColor?: string;
  backgroundColor?: string;
}

/** The relative periods an "A Date Occurring" rule may test (Excel parity). */
export const DATE_OCCURRING_PERIODS = [
  "yesterday",
  "today",
  "tomorrow",
  "last7Days",
  "lastWeek",
  "thisWeek",
  "nextWeek",
  "lastMonth",
  "thisMonth",
  "nextMonth",
] as const;
export type DateOccurringPeriod = (typeof DATE_OCCURRING_PERIODS)[number];

export function isDateOccurringPeriod(
  value: unknown,
): value is DateOccurringPeriod {
  return (
    typeof value === "string" &&
    (DATE_OCCURRING_PERIODS as readonly string[]).includes(value)
  );
}

/**
 * A date-occurring rule (Excel "Highlight Cells Rules" → A Date Occurring…): the
 * cell's date is highlighted when it falls in `period` RELATIVE to today (today /
 * yesterday / tomorrow / last 7 days / this-last-next week / this-last-next
 * month). "Today" is supplied at resolve time via {@link ConditionalDisplayOptions.now}
 * (default: the current date) so the rule re-evaluates as the clock advances, like
 * Excel. A cell whose value is not a recognized date never matches. At least one
 * color is required.
 */
export interface ConditionalDateOccurringRule extends ConditionalDifferentialStyle {
  kind: "dateOccurring";
  range: ConditionalRange;
  period: DateOccurringPeriod;
  textColor?: string;
  backgroundColor?: string;
}

/**
 * A single conditional-format rule. Discriminated on `kind`; a rule WITHOUT a
 * `kind` is a comparison rule (back-compat). Existing callers that only need the
 * comparison fields can still narrow via `rule.kind === undefined ||
 * rule.kind === "comparison"`.
 */
export type ConditionalFormatRule = (
  | ConditionalComparisonRule
  | ConditionalFormulaRule
  | ConditionalColorScaleRule
  | ConditionalDataBarRule
  | ConditionalIconSetRule
  | ConditionalTextContainsRule
  | ConditionalDuplicateValuesRule
  | ConditionalTopBottomRule
  | ConditionalAboveBelowAverageRule
  | ConditionalDateOccurringRule
) & {
  /** Differential font emphasis applied when the rule matches (style-kind
   * rules carry the full differential style; see ConditionalDifferentialStyle). */
  bold?: boolean;
  /**
   * Document-order ordinal among the source rules that share one native
   * priority (some producers restart priorities per range). Absent means 1.
   */
  nativeSequence?: number;
  /**
   * Every rectangle the rule applies to when there is more than one (Excel's
   * "Applies to" `=$A$1:$B$2,$D$1:$E$2`, a multi-area sqref). `range` is
   * always the first area, which is also the anchor relative formulas are
   * written against. Absent for a single-rectangle rule.
   */
  areas?: ConditionalRange[];
  /**
   * Stable editor identity for Manage Rules. This is deliberately opaque and
   * never rendered to the user; imported classic rules derive it from their
   * authenticated native priority while newly-authored rules receive a UUID.
   */
  id?: string;
  /**
   * Import-only identity of the native cfRule priority slot. It is preserved
   * across manager reordering so original-package patchback can update only
   * priority/stopIfTrue without rebuilding the rule body or foreign rules.
   */
  nativePriority?: number;
  /** Stop evaluating lower-priority rules after this rule matches. */
  stopIfTrue?: boolean;
  /**
   * Per-rule native custody receipt. Imported rules carrying this marker stay
   * visible, but Manage Rules must preserve their original OOXML body.
   */
  nativeBlocker?: string;
};

const CONDITIONAL_RULE_ID_RE = /^[A-Za-z0-9][A-Za-z0-9:._-]{0,127}$/;

/**
 * The bounded classic SpreadsheetML subset whose complete body MRAI can
 * rebuild without guessing. Other projected rule kinds remain fully usable in
 * from-scratch workbooks, but an imported package keeps them preserve-only.
 */
export function isClassicNativeConditionalRuleEditable(
  rule: ConditionalFormatRule,
): boolean {
  const kind = rule.kind ?? "comparison";
  return (
    kind === "comparison" ||
    kind === "formula" ||
    kind === "colorScale" ||
    kind === "dataBar" ||
    kind === "iconSet" ||
    kind === "textContains" ||
    kind === "duplicateValues" ||
    kind === "topBottom" ||
    kind === "aboveBelowAverage" ||
    kind === "dateOccurring"
  );
}

/** Visual rules do not support Excel's Stop If True switch. */
export function conditionalRuleSupportsStopIfTrue(
  rule: ConditionalFormatRule,
): boolean {
  const kind = rule.kind ?? "comparison";
  return !["colorScale", "dataBar", "iconSet"].includes(kind);
}

/** The CSS-ready style a matching rule contributes to a cell. */
export interface ConditionalStyle {
  color?: string;
  backgroundColor?: string;
  fontWeight?: number;
  fontStyle?: "italic" | "normal";
  textDecoration?: string;
  /** CSS pixels (the rule's points at 96 dpi). */
  fontSize?: number;
  fontFamily?: string;
  /** CSS gradient image for a gradient-fill or pattern-fill rule. */
  backgroundImage?: string;
  /** The tile of a pattern-fill rule's image (the dotted grays). */
  backgroundSize?: string;
  /** Applied by the grid on top of the cell's own borders (not a CSS key). */
  borderEdges?: CellBorderEdges;
  /** Overrides the cell's number format while the rule matches (not CSS). */
  numberFormat?: string;
}

/**
 * A rule's CSS as the grid lays it over a cell's own: the cell states its
 * underline and strikethrough as `textDecorationLine` (with a double rule's
 * style and thickness), and the display layer draws exactly that key, so a
 * rule's `textDecoration` shorthand is that line -- else a matching rule's
 * underline or strikethrough was never drawn (X-349). A rule that drops the
 * underline drops the cell's double style with it.
 */
export function conditionalStyleOverCellCss<S extends ConditionalStyle>(style: S): Omit<S, "textDecoration"> & {
  textDecorationLine?: string;
  textDecorationStyle?: undefined;
  textDecorationThickness?: undefined;
} {
  if (style.textDecoration === undefined) return style;
  const { textDecoration, ...rest } = style;
  return {
    ...rest,
    textDecorationLine: textDecoration,
    ...(textDecoration.includes("underline") ? {} : { textDecorationStyle: undefined, textDecorationThickness: undefined }),
  };
}

/**
 * The CSS-facing part of a rule's differential style. `borderEdges` and
 * `numberFormat` are not CSS and are applied by the grid separately.
 */
export function applyDifferentialCss(
  style: ConditionalStyle,
  rule: ConditionalDifferentialStyle,
): ConditionalStyle {
  if (rule.textColor) style.color = rule.textColor;
  if (rule.backgroundColor) style.backgroundColor = rule.backgroundColor;
  else if (rule.noFill) style.backgroundColor = "transparent";
  if (rule.bold !== undefined) style.fontWeight = rule.bold ? 700 : 400;
  if (rule.italic !== undefined) style.fontStyle = rule.italic ? "italic" : "normal";
  // Underline and strikethrough are two font properties, so a rule that sets
  // one leaves the other as an earlier matching rule set it; only an explicit
  // "none"/false turns its own property off (live audit B301).
  const decorations = new Set(
    style.textDecoration && style.textDecoration !== "none"
      ? style.textDecoration.split(" ")
      : [],
  );
  if (rule.underline === "single" || rule.underline === "double") {
    decorations.add("underline");
  } else if (rule.underline === "none") {
    decorations.delete("underline");
  }
  if (rule.strikethrough === true) decorations.add("line-through");
  else if (rule.strikethrough === false) decorations.delete("line-through");
  if (decorations.size > 0) style.textDecoration = [...decorations].join(" ");
  else if (rule.underline === "none" || rule.strikethrough === false) {
    style.textDecoration = "none";
  }
  if (rule.fontSize !== undefined) style.fontSize = (rule.fontSize * 96) / 72;
  if (rule.fontName !== undefined) style.fontFamily = rule.fontName;
  if (rule.gradientFill) {
    const gradient = spreadsheetGradientFillCss(rule.gradientFill);
    if (gradient) {
      style.backgroundImage = gradient.backgroundImage;
      style.backgroundColor = gradient.backgroundColor;
    }
  } else if (rule.fillPattern) {
    const pattern = spreadsheetFillPatternCss({
      type: rule.fillPattern.type,
      foregroundColor: rule.fillPattern.foregroundColor ?? "#000000",
      backgroundColor: rule.backgroundColor ?? "#ffffff",
    });
    if (pattern) {
      // The lines alone: the rule's background (above) or the cell's own
      // fill shows between them.
      style.backgroundImage = pattern.backgroundImage;
      if (pattern.backgroundSize) style.backgroundSize = pattern.backgroundSize;
    }
  }
  if (rule.borderEdges) style.borderEdges = rule.borderEdges;
  if (rule.numberFormat !== undefined) style.numberFormat = rule.numberFormat;
  return style;
}

/** A proportional in-cell data bar a dataBar rule contributes to a cell. */
export interface ConditionalDataBar {
  /** Bar width as a percentage of the cell, 0..100. */
  widthPercent: number;
  /** Automatic zero axis in cell coordinates (0 left, 100 right). */
  axisPercent: number;
  /** Bar's left edge in cell coordinates. */
  startPercent: number;
  /** Negative values extend left from the automatic zero axis. */
  direction: "positive" | "negative";
  /** Whether the underlying cell value remains visible over the bar. */
  showValue: boolean;
  /** Validated #rrggbb bar fill color (already the negative colour for a
   * negative bar). */
  color: string;
  /** Excel 2010+ look: a gradient fades to white away from the axis, a
   * bordered bar draws its border, and the zero axis is drawn in its colour
   * where the range crosses zero. */
  gradient?: boolean;
  borderColor?: string;
  axisColor?: string;
  /** An Excel 2007 bar (no x14 extension): square, opaque, from the left. */
  classic?: boolean;
}

/** The icon an iconSet rule contributes to a cell. */
export interface ConditionalIcon {
  iconSet: IconSetName;
  /** Which icon in the set: 0 (lowest bucket) .. count-1 (highest bucket). */
  index: number;
  /** Total icons in the family (3, 4 or 5) so the renderer can size correctly. */
  count: 3 | 4 | 5;
  /** true → only the icon shows; the cell's value text is hidden (Excel parity). */
  iconOnly?: boolean;
}

/**
 * Where a cell's value starts when an icon-set icon is drawn beside it: the
 * icon's 2px inset, its 16px box and a 3px gap (the grid's
 * `.spreadsheet-window__cond-icon` and the editor's padding).
 */
export const CONDITIONAL_ICON_TEXT_INSET_PX = 21;

/**
 * How much further in than the cell's text Windows Excel prints the value of a
 * cell with an icon-set icon: the icon's own 16 px, at the page's scale
 * (lookwin R12-X02: 65-80 dot icons, each value exactly that far past the
 * column's other text, where the grid's 21 px less its padding set it 3 px
 * further; the API's draw_cell_icon).
 */
export const CONDITIONAL_ICON_PRINTED_TEXT_SHIFT_PX = 16;

/**
 * The room an icon-set icon takes from the value beside it, over the cell's
 * own text padding: AutoFit Column Width adds it, and a number has that much
 * less width before it shows #. Excel does both (measured: 0-100 with 3
 * Arrows fits 11pt wider than without, and shows ## at a width where the
 * plain number still shows). 0 without an icon, or with Show Icon Only.
 */
export function conditionalIconTextAllowancePx(
  icon: ConditionalIcon | undefined,
  cellPaddingPx: number,
): number {
  return icon && !icon.iconOnly ? Math.max(0, CONDITIONAL_ICON_TEXT_INSET_PX - cellPaddingPx) : 0;
}

/**
 * The full conditional-format display a cell receives: an optional color style
 * (comparison / formula / colorScale), an optional data bar, and an optional
 * icon. All are independently optional so the renderer can layer whichever the
 * first matching rule supplies.
 */
export interface ConditionalDisplay {
  style: ConditionalStyle;
  dataBar?: ConditionalDataBar;
  icon?: ConditionalIcon;
}





export const MAX_CONDITIONAL_FORMATS = pageEnvelope.maxConditionalFormatsPerSheet;

/**
 * A one-click conditional-format preset (Excel/WPS "quick styles"): a named rule
 * template the UI applies to the current selection. `params` is a rule body sans
 * `range` — the caller passes it straight to {@link buildRule}(selectionRange,
 * params), so it is re-validated by the normalizer (colors funneled through
 * normalizeHexColor, operators/kinds checked). A malformed preset therefore
 * degrades to a null rule rather than injecting unverified state. `id` doubles
 * as the i18n key suffix (`spreadsheet.cfPreset.<id>`).
 */
export interface ConditionalFormatPreset {
  id: string;
  params: Record<string, unknown>;
}

// Canonical Excel conditional-format highlight colors (fill + matching text).
const CF_GREEN_FILL = /* own data: CF_GREEN_FILL */ "#d2f4e0";
const CF_GREEN_TEXT = /* own data: CF_GREEN_TEXT */ "#13723a";
const CF_RED_FILL = /* own data: CF_RED_FILL */ "#f4d2d6";
const CF_RED_TEXT = /* own data: CF_RED_TEXT */ "#72131f";
const CF_YELLOW_FILL = /* own data: CF_YELLOW_FILL */ "#f4ecd2";
const CF_YELLOW_TEXT = /* own data: CF_YELLOW_TEXT */ "#724513";
// Excel's default 3-color-scale stops and data-bar fills.
const CF_SCALE_RED = /* own data: CF_SCALE_RED */ "#e56173";
const CF_SCALE_YELLOW = /* own data: CF_SCALE_YELLOW */ "#ead07b";
const CF_SCALE_GREEN = /* own data: CF_SCALE_GREEN */ "#4fba7c";
const CF_SCALE_WHITE = "#ffffff";
const CF_BAR_BLUE = /* own data: CF_BAR_BLUE */ "#5c8ccc";

/**
 * Excel's Data Bars gallery: six colours, offered once under "Gradient Fill"
 * and once under "Solid Fill", in this order. The preset id is
 * `dataBar<Id>` for the gradient bar and `dataBarSolid<Id>` for the solid one.
 */
export const CONDITIONAL_DATA_BAR_GALLERY: readonly {
  id: "Blue" | "Green" | "Red" | "Orange" | "LightBlue" | "Purple";
  color: string;
}[] = /* own data: CONDITIONAL_DATA_BAR_GALLERY */ [{"id":"Blue","color":"#5c8ccc"},{"id":"Green","color":"#47d777"},{"id":"Red","color":"#d7474c"},{"id":"Orange","color":"#d78f47"},{"id":"LightBlue","color":"#13abec"},{"id":"Purple","color":"#d747d7"}];

// The look Excel gives a gallery bar: negative values in red on the far side
// of a black zero axis; a gradient bar carries a border in its own colour, a
// solid bar has none.
const dataBarGalleryPreset = (
  entry: (typeof CONDITIONAL_DATA_BAR_GALLERY)[number],
  gradient: boolean,
): ConditionalFormatPreset => ({
  id: `dataBar${gradient ? "" : "Solid"}${entry.id}`,
  params: {
    kind: "dataBar",
    color: entry.color,
    gradient,
    negativeColor: "#ff0000",
    axisColor: "#000000",
    ...(gradient ? { borderColor: entry.color } : {}),
  },
});

/**
 * The built-in quick-apply presets, grouped by rule kind. Comparison presets
 * default their threshold to 0 (highlight positives / negatives / zeros) — a
 * meaningful one-click default the user can refine afterwards. Color scales,
 * data bars and icon sets need no threshold. Every entry's `params` is a valid
 * {@link buildRule} body; the presets-build test asserts each one normalizes to
 * the expected kind with its colors intact.
 */
export const CONDITIONAL_FORMAT_PRESETS: readonly ConditionalFormatPreset[] = [
  {
    id: "positiveGreen",
    params: {
      operator: ">",
      value: 0,
      textColor: CF_GREEN_TEXT,
      backgroundColor: CF_GREEN_FILL,
    },
  },
  {
    id: "negativeRed",
    params: {
      operator: "<",
      value: 0,
      textColor: CF_RED_TEXT,
      backgroundColor: CF_RED_FILL,
    },
  },
  {
    id: "zeroYellow",
    params: {
      operator: "=",
      value: 0,
      textColor: CF_YELLOW_TEXT,
      backgroundColor: CF_YELLOW_FILL,
    },
  },
  // Excel names a colour scale from its HIGHEST value down: in "Green -
  // Yellow - Red" the higher values are green and the lower ones red.
  {
    id: "scaleRedYellowGreen",
    params: {
      kind: "colorScale",
      // Excel's three-colour presets: lowest value, 50th percentile, highest
      // value, written out so the rule, its rendering and its editor agree.
      min: { color: CF_SCALE_GREEN },
      mid: { color: CF_SCALE_YELLOW, type: "percentile", value: 50 },
      max: { color: CF_SCALE_RED },
    },
  },
  {
    id: "scaleGreenYellowRed",
    params: {
      kind: "colorScale",
      min: { color: CF_SCALE_RED },
      mid: { color: CF_SCALE_YELLOW, type: "percentile", value: 50 },
      max: { color: CF_SCALE_GREEN },
    },
  },
  {
    id: "scaleWhiteBlue",
    params: {
      kind: "colorScale",
      min: { color: CF_SCALE_WHITE },
      max: { color: CF_BAR_BLUE },
    },
  },
  ...CONDITIONAL_DATA_BAR_GALLERY.map((entry) => dataBarGalleryPreset(entry, true)),
  ...CONDITIONAL_DATA_BAR_GALLERY.map((entry) => dataBarGalleryPreset(entry, false)),
  {
    id: "iconTrafficLights",
    params: {
      kind: "iconSet",
      iconSet: "3-traffic-lights",
      thresholds: [0.33, 0.67],
    },
  },
  {
    id: "iconArrows",
    params: { kind: "iconSet", iconSet: "3-arrows", thresholds: [0.33, 0.67] },
  },
  {
    id: "iconStars",
    params: { kind: "iconSet", iconSet: "3-stars" },
  },
  {
    id: "iconArrows4",
    params: { kind: "iconSet", iconSet: "4-arrows" },
  },
  {
    id: "iconRatings5",
    params: { kind: "iconSet", iconSet: "5-ratings" },
  },
  {
    id: "iconQuarters5",
    params: { kind: "iconSet", iconSet: "5-quarters" },
  },
  {
    id: "duplicateValues",
    params: {
      kind: "duplicateValues",
      textColor: CF_RED_TEXT,
      backgroundColor: CF_RED_FILL,
    },
  },
  {
    id: "top10",
    params: {
      kind: "topBottom",
      rank: 10,
      percent: true,
      textColor: CF_GREEN_TEXT,
      backgroundColor: CF_GREEN_FILL,
    },
  },
  {
    id: "aboveAverage",
    params: {
      kind: "aboveBelowAverage",
      textColor: CF_GREEN_TEXT,
      backgroundColor: CF_GREEN_FILL,
    },
  },
];

const HEX_COLOR_RE = /^#[0-9a-f]{6}$/;

/** True when a candidate string is one of the supported operators. */
export function isConditionalOperator(
  value: unknown,
): value is ConditionalOperator {
  return (
    typeof value === "string" &&
    (OPERATORS as readonly string[]).includes(value)
  );
}

/** True when a candidate string is one of the supported rule kinds. */
export function isConditionalRuleKind(
  value: unknown,
): value is ConditionalRuleKind {
  return (
    typeof value === "string" && (RULE_KINDS as readonly string[]).includes(value)
  );
}

/**
 * The kind a raw persisted rule declares. A missing/blank/unknown `kind` is
 * treated as "comparison" so OLD rules (which never carried a kind) and any
 * forward-incompatible value both fall back to the original comparison
 * behavior rather than being dropped.
 */
export function ruleKindOf(raw: unknown): ConditionalRuleKind {
  if (!raw || typeof raw !== "object") return "comparison";
  const k = (raw as Record<string, unknown>).kind;
  return isConditionalRuleKind(k) ? k : "comparison";
}

/**
 * Coerce a raw color into a canonical lowercase #rrggbb string, or null when it
 * is not a literal hex color. A 3-digit shorthand (#abc) is expanded to 6
 * digits. Anything else — named colors, rgb()/hsl(), `url(...)`, an expression,
 * whitespace tricks — is rejected so a persisted color can never become a CSS
 * injection vector when spread into an inline style object.
 */
export function normalizeHexColor(raw: unknown): string | null {
  if (typeof raw !== "string") return null;
  const trimmed = raw.trim().toLowerCase();
  if (/^#[0-9a-f]{3}$/.test(trimmed)) {
    const [, r, g, b] = trimmed;
    return `#${r}${r}${g}${g}${b}${b}`;
  }
  if (HEX_COLOR_RE.test(trimmed)) return trimmed;
  return null;
}

/**
 * The largest valid 0-based row / column index a conditional-format range may
 * address: the grid is MAX_GRID_ROWS rows × MAX_GRID_COLUMNS columns, so the
 * last addressable cell is (MAX_GRID_ROWS - 1, MAX_GRID_COLUMNS - 1). A range
 * coordinate is always clamped to this inclusive maximum (never MAX_GRID_ROWS /
 * MAX_GRID_COLUMNS themselves, which would point one cell past the grid edge).
 */
export const MAX_RANGE_ROW = MAX_GRID_ROWS - 1;
export const MAX_RANGE_COL = MAX_GRID_COLUMNS - 1;

/**
 * Clamp a single range coordinate into the inclusive grid bounds, optionally
 * tightening to a caller-supplied maximum (e.g. the real sheet's last row/col).
 * The value is assumed already non-negative; the upper bound is the smaller of
 * the grid maximum and any supplied `limit`.
 */
function clampCoord(value: number, gridMax: number, limit?: number): number {
  const upper = limit === undefined ? gridMax : Math.min(gridMax, limit);
  return Math.max(0, Math.min(value, upper));
}

/**
 * Normalize a raw range into a safe rectangle (r1<=r2, c1<=c2), clamped to the
 * grid bounds, or null when any coordinate is non-finite/negative. Coordinates
 * are floored so a fractional persisted value cannot slip through. The clamp
 * ceiling is the LAST addressable cell (MAX_RANGE_ROW / MAX_RANGE_COL), never
 * one cell past the grid edge. Optional `bounds` tightens the clamp to a real
 * sheet's last row/column index (maxRow / maxCol) so a range persisted against a
 * larger sheet can never address a cell outside the current grid.
 */
export function normalizeRange(
  raw: unknown,
  bounds?: { maxRow?: number; maxCol?: number },
): ConditionalRange | null {
  if (!raw || typeof raw !== "object") return null;
  const r = raw as Record<string, unknown>;
  const r1 = Math.floor(Number(r.r1));
  const c1 = Math.floor(Number(r.c1));
  const r2 = Math.floor(Number(r.r2));
  const c2 = Math.floor(Number(r.c2));
  if (
    !Number.isFinite(r1) ||
    !Number.isFinite(c1) ||
    !Number.isFinite(r2) ||
    !Number.isFinite(c2)
  ) {
    return null;
  }
  if (r1 < 0 || c1 < 0 || r2 < 0 || c2 < 0) return null;
  const lowR = Math.min(r1, r2);
  const highR = Math.max(r1, r2);
  const lowC = Math.min(c1, c2);
  const highC = Math.max(c1, c2);
  const maxRow = bounds?.maxRow;
  const maxCol = bounds?.maxCol;
  return {
    r1: clampCoord(lowR, MAX_RANGE_ROW, maxRow),
    c1: clampCoord(lowC, MAX_RANGE_COL, maxCol),
    r2: clampCoord(highR, MAX_RANGE_ROW, maxRow),
    c2: clampCoord(highC, MAX_RANGE_COL, maxCol),
  };
}

/** True when (row, col) lies inside the (inclusive, normalized) range. */
export function rangeContains(
  range: ConditionalRange,
  row: number,
  col: number,
): boolean {
  return (
    row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2
  );
}

/** True when two ranges share at least one cell. */
export function rangesOverlap(
  a: ConditionalRange,
  b: ConditionalRange,
): boolean {
  return a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
}

function rangesEqual(a: ConditionalRange, b: ConditionalRange): boolean {
  return a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2;
}

/** Most areas one rule may apply to. Excel sets no limit; this only keeps a
 * malformed page out, as the save's writer does (65,536). At 64, a rule read
 * from a file with more lost all its areas but the first on any command over
 * the sheet: R179-X04, 121 areas, a paste wrote the rule as A1:AE8 alone. */
export const MAX_CONDITIONAL_AREAS = 65_536;

/** The rectangles a rule applies to: its areas, or just its range. */
export function ruleAreas(
  rule: { range: ConditionalRange; areas?: ConditionalRange[] },
): readonly ConditionalRange[] {
  return rule.areas && rule.areas.length > 1 ? rule.areas : [rule.range];
}

/**
 * `next` (a sheet's rules as an edit leaves them, read through the sheet's own
 * bounds) with each rule the edit did not move given back the areas `stored`
 * holds: a rule over whole columns reaches the grid's last row (Excel's
 * A5:AK1048576), which the sheet's bounds clamp to its last used row for the
 * grid to draw. Written back clamped, a new rule on the sheet cut it to A5:AK79
 * (xlsx-rounds R316-X13 RULESHIFT).
 */
export function rulesKeepingStoredAreas(
  stored: unknown,
  next: readonly ConditionalFormatRule[],
  bounds: { maxRow?: number; maxCol?: number },
): ConditionalFormatRule[] {
  const full = new Map(normalizeRules(stored).flatMap((rule) => (rule.id ? [[rule.id, rule] as const] : [])));
  if (full.size === 0) return [...next];
  const shown = new Map(normalizeRules(stored, bounds).flatMap((rule) => (rule.id ? [[rule.id, rule] as const] : [])));
  return next.map((rule) => {
    const own = rule.id ? full.get(rule.id) : undefined;
    const clamped = rule.id ? shown.get(rule.id) : undefined;
    if (!own || !clamped) return rule;
    const areas = JSON.stringify(ruleAreas(rule));
    if (areas !== JSON.stringify(ruleAreas(clamped)) || areas === JSON.stringify(ruleAreas(own))) return rule;
    const restored: ConditionalFormatRule = { ...rule, range: own.range };
    if (own.areas && own.areas.length > 1) restored.areas = own.areas;
    return restored;
  });
}

/** Whether a cell lies in any of the rule's areas. */
export function ruleContains(
  rule: { range: ConditionalRange; areas?: ConditionalRange[] },
  row: number,
  col: number,
): boolean {
  return ruleAreas(rule).some((area) => rangeContains(area, row, col));
}

interface ConditionalCoverage {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  /** Every area of every rule, flattened as r1, c1, r2, c2. */
  areas: Float64Array;
}

const conditionalCoverageCache = new WeakMap<
  readonly { range: ConditionalRange; areas?: ConditionalRange[] }[],
  ConditionalCoverage
>();

/**
 * Whether any rule's areas contain the cell, without resolving anything. The
 * grid asks this before `resolveCellConditionalDisplay`, so a cell outside
 * every rule skips the value read, the options object and the per-rule walk.
 * The flattened areas are built once per rules array.
 */
export function conditionalRulesCover(
  rules: readonly { range: ConditionalRange; areas?: ConditionalRange[] }[],
  row: number,
  col: number,
): boolean {
  let coverage = conditionalCoverageCache.get(rules);
  if (!coverage) {
    const flat: number[] = [];
    let r1 = Infinity;
    let c1 = Infinity;
    let r2 = -Infinity;
    let c2 = -Infinity;
    for (const rule of rules) {
      for (const area of ruleAreas(rule)) {
        flat.push(area.r1, area.c1, area.r2, area.c2);
        r1 = Math.min(r1, area.r1);
        c1 = Math.min(c1, area.c1);
        r2 = Math.max(r2, area.r2);
        c2 = Math.max(c2, area.c2);
      }
    }
    coverage = { r1, c1, r2, c2, areas: Float64Array.from(flat) };
    conditionalCoverageCache.set(rules, coverage);
  }
  if (row < coverage.r1 || row > coverage.r2 || col < coverage.c1 || col > coverage.c2) return false;
  const { areas } = coverage;
  for (let i = 0; i < areas.length; i += 4) {
    if (row >= areas[i] && row <= areas[i + 2] && col >= areas[i + 1] && col <= areas[i + 3]) return true;
  }
  return false;
}

/** Whether any of the rule's areas overlaps `range`. */
export function ruleOverlaps(
  rule: { range: ConditionalRange; areas?: ConditionalRange[] },
  range: ConditionalRange,
): boolean {
  return ruleAreas(rule).some((area) => rangesOverlap(area, range));
}

/**
 * A multi-area rule as one single-rectangle rule per area, with relative
 * formulas rebased from the first area's anchor to each area's own. This is
 * what copying or clearing part of such a rule operates on.
 */
export function splitRuleAreas(
  rule: ConditionalFormatRule,
): ConditionalFormatRule[] {
  const areas = ruleAreas(rule);
  if (areas.length === 1 && !rule.areas) return [rule];
  return areas.map((area, index) => {
    const { areas: _areas, id: _id, ...rest } = rule;
    void _areas;
    void _id;
    const candidate: Record<string, unknown> = { ...rest, range: area };
    const formula = conditionalRuleFormula(rule);
    if (formula !== undefined && index > 0) {
      candidate.formula = shiftFormula(
        formula,
        area.r1 - rule.range.r1,
        area.c1 - rule.range.c1,
      );
    }
    const normalized = normalizeRule(candidate);
    return normalized ?? { ...rule, range: area };
  });
}

/** Longest formula string a formula rule may carry (DoS / paste-bomb bound). */
export const MAX_CONDITIONAL_FORMULA_LENGTH = 1_000;

/**
 * Normalize a comparison rule (the ORIGINAL shape). Emits NO `kind` field so the
 * persisted output is byte-for-byte identical to pre-union data. Returns null
 * when the operator/value is invalid. A rule with no style is kept (B469).
 */
function normalizeComparisonRule(
  r: Record<string, unknown>,
  range: ConditionalRange,
): ConditionalComparisonRule | null {
  if (!isConditionalOperator(r.operator)) return null;
  const value = Number(r.value);
  if (!Number.isFinite(value)) return null;
  const colors = extractColors(r);
  const rule: ConditionalComparisonRule = { range, operator: r.operator, value };
  if (isRangeConditionalOperator(r.operator)) {
    const value2 = Number(r.value2);
    if (!Number.isFinite(value2)) return null;
    rule.value2 = value2;
  }
  applyColors(rule, colors);
  return rule;
}

/** Normalize a formula rule: a non-empty bounded formula + its style. */
function normalizeFormulaRule(
  r: Record<string, unknown>,
  range: ConditionalRange,
): ConditionalFormulaRule | null {
  if (typeof r.formula !== "string") return null;
  const formula = r.formula.trim();
  if (formula === "" || formula.length > MAX_CONDITIONAL_FORMULA_LENGTH) {
    return null;
  }
  const colors = extractColors(r);
  const rule: ConditionalFormulaRule = { kind: "formula", range, formula };
  applyColors(rule, colors);
  if (r.nativeCellIs === true) rule.nativeCellIs = true;
  if (r.nativeErrorRule === true) rule.nativeErrorRule = true;
  return rule;
}

/** Normalize a color-scale rule: valid min/max stops, optional mid (3-point). */
function normalizeColorScaleRule(
  r: Record<string, unknown>,
  range: ConditionalRange,
): ConditionalColorScaleRule | null {
  const minColor = normalizeHexColor(
    (r.min as Record<string, unknown> | undefined)?.color,
  );
  const maxColor = normalizeHexColor(
    (r.max as Record<string, unknown> | undefined)?.color,
  );
  if (!minColor || !maxColor) return null;
  const midColor = normalizeHexColor(
    (r.mid as Record<string, unknown> | undefined)?.color,
  );
  const rule: ConditionalColorScaleRule = {
    kind: "colorScale",
    range,
    min: normalizeColorScaleStop(r.min as Record<string, unknown>, minColor, "min"),
    max: normalizeColorScaleStop(r.max as Record<string, unknown>, maxColor, "max"),
  };
  if (midColor) {
    rule.mid = normalizeColorScaleStop(
      r.mid as Record<string, unknown>,
      midColor,
      "mid",
    );
  }
  return rule;
}

/**
 * Coerce one persisted stop. A family the renderer cannot represent, or a
 * family without its value, falls back to the position's default so a stale
 * or foreign payload never yields a rule that throws while painting.
 */
function normalizeColorScaleStop(
  raw: Record<string, unknown>,
  color: string,
  position: "min" | "mid" | "max",
): ColorScaleStop {
  const themeColor = normalizeConditionalThemeColor(raw.themeColor);
  const base: ColorScaleStop = { color, ...(themeColor ? { themeColor } : {}) };
  const rawType = raw.type;
  if (rawType === "formula") {
    const formula = typeof raw.formula === "string" ? raw.formula.trim() : "";
    if (formula && formula.length <= MAX_CONDITIONAL_FORMULA_LENGTH) {
      return { ...base, type: "formula", formula };
    }
    return base;
  }
  if (rawType === "min" || rawType === "max") {
    // An endpoint of its own family is the default shape; a midpoint pinned
    // to min or max keeps the family explicitly.
    return position === "mid" ? { ...base, type: rawType } : base;
  }
  if (rawType === "percent" || rawType === "percentile" || rawType === "number") {
    const value = Number(raw.value);
    if (Number.isFinite(value)) return { ...base, type: rawType, value };
  }
  return base;
}

/** Normalize a data-bar rule: a single valid bar color. */
function normalizeDataBarRule(
  r: Record<string, unknown>,
  range: ConditionalRange,
): ConditionalDataBarRule | null {
  const color = normalizeHexColor(r.color);
  if (!color) return null;
  const rawMin = Number(r.minLength);
  const rawMax = Number(r.maxLength);
  const minLength = Number.isSafeInteger(rawMin) ? rawMin : 10;
  const maxLength = Number.isSafeInteger(rawMax) ? rawMax : 90;
  if (
    minLength < 0 ||
    maxLength > 100 ||
    minLength > maxLength
  ) {
    return null;
  }
  return {
    kind: "dataBar",
    range,
    color,
    ...(minLength !== 10 ? { minLength } : {}),
    ...(maxLength !== 90 ? { maxLength } : {}),
    ...(r.showValue === false ? { showValue: false } : {}),
    ...(typeof r.gradient === "boolean" ? { gradient: r.gradient } : {}),
    ...(normalizeHexColor(r.negativeColor) ? { negativeColor: normalizeHexColor(r.negativeColor)! } : {}),
    ...(normalizeHexColor(r.axisColor) ? { axisColor: normalizeHexColor(r.axisColor)! } : {}),
    ...(normalizeHexColor(r.borderColor) ? { borderColor: normalizeHexColor(r.borderColor)! } : {}),
    ...(r.axisPosition === "automatic" || r.axisPosition === "middle" || r.axisPosition === "none"
      ? { axisPosition: r.axisPosition }
      : {}),
    ...(normalizeConditionalThemeColor(r.themeColor)
      ? { themeColor: normalizeConditionalThemeColor(r.themeColor) }
      : {}),
    ...(typeof r.nativeDataBarId === "string" &&
      /^\{[0-9A-Fa-f-]{36}\}$/.test(r.nativeDataBarId)
      ? { nativeDataBarId: r.nativeDataBarId.toUpperCase() }
      : {}),
    ...dataBarThresholdEntries(r),
  };
}

/** The typed min/max thresholds of a persisted data bar, defaults omitted. */
function dataBarThresholdEntries(
  r: Record<string, unknown>,
): { min?: DataBarThreshold; max?: DataBarThreshold } {
  const out: { min?: DataBarThreshold; max?: DataBarThreshold } = {};
  for (const position of ["min", "max"] as const) {
    const raw = r[position];
    if (!raw || typeof raw !== "object") continue;
    if ((raw as Record<string, unknown>).type === "automatic") {
      out[position] = { type: "automatic" };
      continue;
    }
    const { color: _color, ...threshold } = normalizeColorScaleStop(
      raw as Record<string, unknown>,
      "#000000",
      position,
    );
    if (threshold.type) out[position] = threshold;
  }
  return out;
}

/**
 * The value axis a data bar maps onto: the range extent unless the rule pins
 * an endpoint to a number, percent, percentile or formula (resolved like a
 * colour-scale stop; a formula stop keeps the range endpoint here).
 */
export function dataBarExtent(
  rule: ConditionalDataBarRule,
  extent: RangeExtent,
  sortedNumbers?: readonly number[],
): RangeExtent {
  if (!rule.min?.type && !rule.max?.type) return extent;
  const autoMin = rule.min?.type === "automatic";
  const autoMax = rule.max?.type === "automatic";
  const thresholds = colorScaleThresholds(
    {
      kind: "colorScale",
      range: rule.range,
      min: { color: "#000000", ...(autoMin ? {} : (rule.min ?? {})) } as ColorScaleStop,
      max: { color: "#000000", ...(autoMax ? {} : (rule.max ?? {})) } as ColorScaleStop,
    },
    extent,
    sortedNumbers,
  );
  // Two untyped ends read the range extent; an automatic end then sets its
  // side. Excel's automatic endpoints: min(0, lowest) and max(0, highest).
  let [min, max] = thresholds ?? [extent.min, extent.max];
  if (autoMin) min = Math.min(0, extent.min);
  if (autoMax) max = Math.max(0, extent.max);
  return max > min ? { min, max } : extent;
}

/**
 * Coerce a raw thresholds value into exactly (count-1) sorted fractions strictly
 * inside (0,1) for the given icon count. A missing / non-array / wrong-length /
 * non-finite input falls back to the Excel defaults for that count
 * ({@link defaultIconThresholds}). When the right number of finite values is
 * present they are clamped into (0,1) and sorted ascending so every bucket stays
 * reachable and the buckets are monotonic regardless of input order.
 */
function normalizeThresholds(raw: unknown, count: 3 | 4 | 5): number[] {
  const need = count - 1;
  if (!Array.isArray(raw) || raw.length !== need) {
    return defaultIconThresholds(count);
  }
  const nums = raw.map((v) => Number(v));
  if (nums.some((n) => !Number.isFinite(n))) return defaultIconThresholds(count);
  const clamp = (v: number) => Math.max(0.001, Math.min(0.999, v));
  return nums.map(clamp).sort((a, b) => a - b);
}

/** A ">=" percent boundary: the shape `thresholds` alone can carry. */
function isPlainPercentStop(stop: IconThreshold): boolean {
  return stop.type === "percent" && stop.gte !== false;
}

/**
 * Coerce raw typed icon boundaries into exactly (count-1) valid stops, or
 * null when absent or unusable (the rule then keeps its `thresholds`).
 */
function normalizeIconStops(raw: unknown, count: 3 | 4 | 5): IconThreshold[] | null {
  if (!Array.isArray(raw) || raw.length !== count - 1) return null;
  const stops: IconThreshold[] = [];
  for (const item of raw) {
    if (!item || typeof item !== "object") return null;
    const entry = item as Record<string, unknown>;
    const type = entry.type;
    let stop: IconThreshold;
    if (type === "formula") {
      const formula = typeof entry.formula === "string" ? entry.formula.trim() : "";
      if (!formula || formula.length > MAX_CONDITIONAL_FORMULA_LENGTH) return null;
      stop = { type, formula };
    } else if (type === "number" || type === "percent" || type === "percentile") {
      const value = Number(entry.value);
      if (!Number.isFinite(value)) return null;
      if (type !== "number" && (value < 0 || value > 100)) return null;
      stop = { type, value };
    } else {
      return null;
    }
    if (entry.gte === false) stop.gte = false;
    stops.push(stop);
  }
  return stops;
}

/**
 * Normalize an icon-set rule: a known family + exactly (count-1) sorted
 * thresholds for that family + optional reverse / show-icon-only flags. An
 * unknown family name is rejected (the rule is dropped). The flags are emitted
 * ONLY when true so a default rule's persisted shape stays minimal.
 */
function normalizeIconSetRule(
  r: Record<string, unknown>,
  range: ConditionalRange,
): ConditionalIconSetRule | null {
  if (!isIconSetName(r.iconSet)) return null;
  const iconSet = r.iconSet;
  const count = iconSetCount(iconSet);
  const rule: ConditionalIconSetRule = {
    kind: "iconSet",
    range,
    iconSet,
    thresholds: normalizeThresholds(r.thresholds, count),
  };
  const stops = normalizeIconStops(r.stops, count);
  if (stops) {
    if (stops.every(isPlainPercentStop)) {
      rule.thresholds = normalizeThresholds(
        stops.map((stop) => (stop.value ?? 0) / 100),
        count,
      );
    } else {
      rule.stops = stops;
      const defaults = defaultIconThresholds(count);
      rule.thresholds = normalizeThresholds(
        stops.map((stop, index) =>
          stop.type === "percent" ? (stop.value ?? 0) / 100 : defaults[index],
        ),
        count,
      );
    }
  }
  if (r.reverse === true) rule.reverse = true;
  if (r.showIconOnly === true) rule.showIconOnly = true;
  if (Array.isArray(r.customIcons) && r.customIcons.length === count) {
    const slots: ConditionalCustomIcon[] = [];
    for (const raw of r.customIcons) {
      if (!raw || typeof raw !== "object") break;
      const slot = raw as Record<string, unknown>;
      const family = slot.iconSet === "no-icons" || isIconSetName(slot.iconSet) ? slot.iconSet : null;
      const iconId = typeof slot.iconId === "number" && Number.isInteger(slot.iconId) ? slot.iconId : -1;
      if (
        !family ||
        iconId < 0 ||
        (family !== "no-icons" && iconId >= iconSetCount(family))
      ) {
        break;
      }
      slots.push({ iconSet: family, iconId });
    }
    if (slots.length === count) rule.customIcons = slots;
  }
  if (typeof r.nativeX14Id === "string" && /^\{[0-9A-Fa-f-]{36}\}$/.test(r.nativeX14Id)) {
    rule.nativeX14Id = r.nativeX14Id.toUpperCase();
  }
  return rule;
}

/** Longest substring a textContains rule may carry (DoS / paste-bomb bound). */
export const MAX_CONDITIONAL_TEXT_LENGTH = 1_000;

/** True when a candidate string is one of the supported text-match operators. */
export function isTextMatchOperator(value: unknown): value is TextMatchOperator {
  return (
    typeof value === "string" &&
    (TEXT_MATCH_OPERATORS as readonly string[]).includes(value)
  );
}

/**
 * Normalize every differential-style field a raw style-kind rule carries.
 * `hasStyle` is false when the rule would paint nothing at all, which is the
 * one shape Excel itself refuses.
 */
export function normalizeDifferentialStyle(r: Record<string, unknown>): {
  style: ConditionalDifferentialStyle;
  hasStyle: boolean;
} {
  const style: ConditionalDifferentialStyle = {};
  const textColor = normalizeHexColor(r.textColor);
  const backgroundColor = normalizeHexColor(r.backgroundColor);
  if (textColor) style.textColor = textColor;
  if (backgroundColor) style.backgroundColor = backgroundColor;
  const textThemeColor = normalizeConditionalThemeColor(r.textThemeColor);
  if (textThemeColor && textThemeColor.rgb === textColor) {
    style.textThemeColor = textThemeColor;
  }
  const backgroundThemeColor = normalizeConditionalThemeColor(r.backgroundThemeColor);
  if (backgroundThemeColor && backgroundThemeColor.rgb === backgroundColor) {
    style.backgroundThemeColor = backgroundThemeColor;
  }
  for (const key of ["bold", "italic", "strikethrough"] as const) {
    if (typeof r[key] === "boolean") style[key] = r[key];
  }
  if (r.underline === "single" || r.underline === "double" || r.underline === "none") {
    style.underline = r.underline;
  }
  const fontSize = typeof r.fontSize === "number" ? r.fontSize : Number.NaN;
  if (Number.isFinite(fontSize) && fontSize >= 1 && fontSize <= 409) {
    style.fontSize = fontSize;
  }
  if (typeof r.fontName === "string" && r.fontName !== "" && r.fontName.length <= 64) {
    style.fontName = r.fontName;
  }
  if (r.fontExtras && typeof r.fontExtras === "object" && !Array.isArray(r.fontExtras)) {
    const extras: Record<string, string> = {};
    for (const key of CONDITIONAL_FONT_EXTRA_MEMBERS) {
      const value = (r.fontExtras as Record<string, unknown>)[key];
      if (typeof value === "string" && value !== "" && value.length <= 64) extras[key] = value;
    }
    if (Object.keys(extras).length > 0) style.fontExtras = extras;
  }
  if (
    typeof r.numberFormat === "string" &&
    r.numberFormat !== "" &&
    r.numberFormat.length <= 255
  ) {
    style.numberFormat = r.numberFormat;
  }
  if (r.borderEdges && typeof r.borderEdges === "object" && !Array.isArray(r.borderEdges)) {
    const edges: CellBorderEdges = {};
    for (const edge of BORDER_EDGES) {
      const raw = (r.borderEdges as Record<string, unknown>)[edge];
      if (!raw || typeof raw !== "object") continue;
      const entry = raw as Record<string, unknown>;
      // A side the rule clears (the dxf's empty `<left/>`; Format Cells >
      // Border > None): Excel removes the cell's own edge there (X-304).
      if (entry.style === "none" && edge !== "diagonal") {
        edges[edge] = { style: "none" } as unknown as RichBorderEdge;
        continue;
      }
      if (!isBorderLineStyle(entry.style)) continue;
      const color = normalizeHexColor(entry.color);
      const normalized: RichBorderEdge = { style: entry.style };
      if (color) normalized.color = color;
      if (edge === "diagonal") {
        if (entry.up === true) normalized.up = true;
        if (entry.down === true) normalized.down = true;
      }
      edges[edge] = normalized;
    }
    if (Object.keys(edges).length > 0) style.borderEdges = edges;
  }
  const gradient = normalizeSpreadsheetGradientFill(r.gradientFill);
  if (gradient) {
    style.gradientFill = gradient;
    style.backgroundColor = gradient.stops[0].color;
  } else if (r.fillPattern && typeof r.fillPattern === "object" && isSpreadsheetFillPatternType((r.fillPattern as { type?: unknown }).type)) {
    const raw = r.fillPattern as { type: SpreadsheetFillPatternType; foregroundColor?: unknown };
    const foreground = normalizeHexColor(raw.foregroundColor);
    style.fillPattern = { type: raw.type, ...(foreground ? { foregroundColor: foreground } : {}) };
  }
  if (r.noFill === true && !style.backgroundColor) style.noFill = true;
  return { style, hasStyle: Object.keys(style).length > 0 };
}

/**
 * The differential style a rule paints, which may be empty: a cfRule's dxfId
 * is optional (CT_CfRule), and a rule with no format still decides, takes its
 * priority slot and, with Stop If True, keeps lower rules off the cells it
 * matches (Excel lists it as "No Format Set"). So no rule kind is dropped for
 * want of a style (live audit B469).
 */
function extractColors(r: Record<string, unknown>): {
  style: ConditionalDifferentialStyle;
} {
  return { style: normalizeDifferentialStyle(r).style };
}

/** Attach the surviving differential style to a partially-built rule. */
function applyColors(
  rule: ConditionalDifferentialStyle,
  colors: { style: ConditionalDifferentialStyle },
): void {
  Object.assign(rule, colors.style);
}

/**
 * Normalize a textContains rule: a known text operator + a non-empty bounded
 * needle + its style. The needle is kept EXACTLY as typed: a leading or trailing
 * space is part of what SEARCH looks for ("cat " does not match "catx"), so
 * trimming it would change the rule (live audit B037).
 */
function normalizeTextContainsRule(
  r: Record<string, unknown>,
  range: ConditionalRange,
): ConditionalTextContainsRule | null {
  if (!isTextMatchOperator(r.operator)) return null;
  if (typeof r.text !== "string") return null;
  const text = r.text;
  if (text === "" || text.length > MAX_CONDITIONAL_TEXT_LENGTH) return null;
  const colors = extractColors(r);
  const rule: ConditionalTextContainsRule = {
    kind: "textContains",
    range,
    text,
    operator: r.operator,
  };
  if (typeof r.formula === "string") {
    const formula = r.formula.trim();
    if (formula !== "" && formula.length <= MAX_CONDITIONAL_FORMULA_LENGTH) rule.formula = formula;
  }
  applyColors(rule, colors);
  return rule;
}

/**
 * Normalize a date-occurring rule: a recognized relative `period` + its style.
 */
function normalizeDateOccurringRule(
  r: Record<string, unknown>,
  range: ConditionalRange,
): ConditionalDateOccurringRule | null {
  if (!isDateOccurringPeriod(r.period)) return null;
  const colors = extractColors(r);
  const rule: ConditionalDateOccurringRule = {
    kind: "dateOccurring",
    range,
    period: r.period,
  };
  applyColors(rule, colors);
  return rule;
}

/** Normalize a duplicate/unique-values rule: optional `unique` flag + its style. */
function normalizeDuplicateValuesRule(
  r: Record<string, unknown>,
  range: ConditionalRange,
): ConditionalDuplicateValuesRule | null {
  const colors = extractColors(r);
  const rule: ConditionalDuplicateValuesRule = {
    kind: "duplicateValues",
    range,
  };
  if (r.unique === true) rule.unique = true;
  applyColors(rule, colors);
  return rule;
}

/**
 * Normalize a top/bottom rule: a finite rank ≥1 (floored), optional bottom /
 * percent flags, its style. A percent rank is additionally capped at 100. A
 * non-finite or sub-1 rank is rejected (an empty highlight is never useful).
 */
function normalizeTopBottomRule(
  r: Record<string, unknown>,
  range: ConditionalRange,
): ConditionalTopBottomRule | null {
  const rawRank = Math.floor(Number(r.rank));
  if (!Number.isFinite(rawRank) || rawRank < 1) return null;
  const colors = extractColors(r);
  const percent = r.percent === true;
  const rank = percent ? Math.min(rawRank, 100) : rawRank;
  const rule: ConditionalTopBottomRule = { kind: "topBottom", range, rank };
  if (r.bottom === true) rule.bottom = true;
  if (percent) rule.percent = true;
  applyColors(rule, colors);
  return rule;
}

/**
 * Normalize an above/below-average rule: optional below flag, optional finite
 * stdDevs ≥0 (omitted when 0), its style. A negative/non-finite stdDevs is
 * dropped (treated as the plain mean) rather than failing the rule.
 */
function normalizeAboveBelowAverageRule(
  r: Record<string, unknown>,
  range: ConditionalRange,
): ConditionalAboveBelowAverageRule | null {
  const colors = extractColors(r);
  const rule: ConditionalAboveBelowAverageRule = {
    kind: "aboveBelowAverage",
    range,
  };
  if (r.below === true) rule.below = true;
  if (r.equalAverage === true) rule.equalAverage = true;
  const k = Number(r.stdDevs);
  if (Number.isFinite(k) && k > 0) rule.stdDevs = k;
  applyColors(rule, colors);
  return rule;
}

/**
 * Coerce a single raw entry into a normalized ConditionalFormatRule, or null
 * when it is malformed. Dispatches on `kind` (a missing/unknown kind is treated
 * as "comparison" so OLD data keeps working):
 *  - comparison: bad range/operator/value → dropped; emits
 *    NO `kind` field so persisted output is byte-for-byte identical to pre-union
 *    data.
 *  - formula:    needs a non-empty bounded formula. A style-kind rule with
 *    no differential style is kept: it is Excel's "No Format Set" rule.
 *  - colorScale: needs valid min + max colors (mid optional → 3-point).
 *  - dataBar:    needs a valid bar color.
 *  - iconSet:    needs a known icon set.
 * Colors always run through normalizeHexColor (no CSS injection). Optional
 * `bounds` tightens the range clamp to a real sheet's last row/column index so a
 * rule persisted against a larger sheet is clamped into the current grid.
 */
export function normalizeRule(
  raw: unknown,
  bounds?: { maxRow?: number; maxCol?: number },
): ConditionalFormatRule | null {
  if (!raw || typeof raw !== "object") return null;
  const r = raw as Record<string, unknown>;
  const range = normalizeRange(r.range, bounds);
  if (!range) return null;
  let rule: ConditionalFormatRule | null;
  switch (ruleKindOf(raw)) {
    case "formula":
      rule = normalizeFormulaRule(r, range);
      break;
    case "colorScale":
      rule = normalizeColorScaleRule(r, range);
      break;
    case "dataBar":
      rule = normalizeDataBarRule(r, range);
      break;
    case "iconSet":
      rule = normalizeIconSetRule(r, range);
      break;
    case "textContains":
      rule = normalizeTextContainsRule(r, range);
      break;
    case "duplicateValues":
      rule = normalizeDuplicateValuesRule(r, range);
      break;
    case "topBottom":
      rule = normalizeTopBottomRule(r, range);
      break;
    case "aboveBelowAverage":
      rule = normalizeAboveBelowAverageRule(r, range);
      break;
    case "dateOccurring":
      rule = normalizeDateOccurringRule(r, range);
      break;
    case "comparison":
    default:
      rule = normalizeComparisonRule(r, range);
      break;
  }
  if (
    rule &&
    typeof r.nativePriority === "number" &&
    Number.isSafeInteger(r.nativePriority) &&
    r.nativePriority > 0 &&
    r.nativePriority <= 2_147_483_647
  ) {
    rule.nativePriority = r.nativePriority;
  }
  if (
    rule &&
    typeof r.id === "string" &&
    CONDITIONAL_RULE_ID_RE.test(r.id)
  ) {
    rule.id = r.id;
  } else if (rule?.nativePriority !== undefined) {
    rule.id = `cf-native-${rule.nativePriority}`;
  }
  if (rule && r.stopIfTrue === true) rule.stopIfTrue = true;
  if (rule && Array.isArray(r.areas) && r.areas.length >= 2 && r.areas.length <= MAX_CONDITIONAL_AREAS) {
    const areas = r.areas.map((area) => normalizeRange(area, bounds));
    if (
      areas.every((area): area is ConditionalRange => area !== null) &&
      rangesEqual(areas[0], rule.range)
    ) {
      rule.areas = areas;
    }
  }
  if (
    rule &&
    typeof r.nativeSequence === "number" &&
    Number.isSafeInteger(r.nativeSequence) &&
    r.nativeSequence > 1
  ) {
    rule.nativeSequence = r.nativeSequence;
  }
  if (
    rule &&
    typeof r.nativeBlocker === "string" &&
    r.nativeBlocker.length > 0 &&
    r.nativeBlocker.length <= 220
  ) {
    rule.nativeBlocker = r.nativeBlocker;
  }
  return rule;
}

/**
 * Normalize a raw conditionalFormats array off a persisted sheet: drop
 * malformed entries and cap the list length (DoS bound). Returns a fresh array;
 * array order is preserved so first-match-wins lookup is deterministic. Optional
 * `bounds` tightens each rule's range clamp to a real sheet's last row/column
 * index (maxRow / maxCol) so rules persisted against a larger sheet are clamped
 * into the current grid on load.
 */
export function normalizeRules(
  raw: unknown,
  bounds?: { maxRow?: number; maxCol?: number },
): ConditionalFormatRule[] {
  if (!Array.isArray(raw)) return [];
  const out: ConditionalFormatRule[] = [];
  for (const item of raw) {
    const rule = normalizeRule(item, bounds);
    if (!rule) continue;
    out.push(rule);
    if (out.length >= MAX_CONDITIONAL_FORMATS) break;
  }
  return out;
}

/** The rules past the page one sheet may add to a render
 * (`conditionalFormatsPastPage`; the API's MAX_PAST_PAGE_RULES_PER_SHEET). */
export const MAX_PAST_PAGE_CONDITIONAL_FORMATS = 20_000;

/**
 * A sheet's rules past the ones its page lists, as a render draws them after
 * its own: `conditionalFormatsPastPage`, which the API reads from the
 * workbook's saved file for a print it makes (spreadsheet_cf_overflow;
 * R463-X31's grey survey rows past the page's 1,000), never saved or edited.
 * Normalized as the page's own are, bounded by
 * MAX_PAST_PAGE_CONDITIONAL_FORMATS.
 */
export function normalizePastPageRules(
  raw: unknown,
  bounds?: { maxRow?: number; maxCol?: number },
): ConditionalFormatRule[] {
  if (!Array.isArray(raw)) return [];
  const out: ConditionalFormatRule[] = [];
  for (const item of raw) {
    const rule = normalizeRule(item, bounds);
    if (!rule) continue;
    out.push(rule);
    if (out.length >= MAX_PAST_PAGE_CONDITIONAL_FORMATS) break;
  }
  return out;
}

/**
 * Coerce a cell value to a finite number, or null. Numbers pass through;
 * strings are trimmed then parsed; blank/non-numeric values yield null so they
 * never match a comparison (matching Excel's cell-value rules, which only fire
 * on numeric cells).
 */
function toNumber(value: unknown): number | null {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed === "") return null;
    const n = Number(trimmed);
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

/**
 * The value a numeric rule (colour scale, data bar, icon set, top/bottom,
 * above/below average) acts on: a NUMBER cell only. Text that looks like a
 * number ('100, a Text-formatted entry) is text to Excel and gets no bar,
 * colour or icon, and does not stretch the range's min/max (live audit B035).
 */
function numericCellValue(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

/**
 * The text a text rule searches: the cell's value as text, untrimmed, so a
 * needle with spaces meets the cell's own spaces exactly as SEARCH would.
 */
function cellSearchText(value: unknown): string {
  if (typeof value === "string") return value;
  return cellText(value);
}

/**
 * Coerce a cell value to the trimmed display text used by text/duplicate rules.
 * Numbers/booleans stringify; strings trim; blank/nullish → "". This is the same
 * text the user sees, so a textContains needle and duplicate-equality both match
 * Excel's "by displayed value" behavior.
 */
function cellText(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value.trim();
  if (typeof value === "number") return Number.isFinite(value) ? String(value) : "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return "";
}

/**
 * The case-insensitive equality key for duplicate/unique comparison: the
 * trimmed lowercased text, or null for a blank cell (blanks never participate in
 * duplicate detection, Excel parity). 1 and "1" share a key (both → "1"), as do
 * "A" and "a".
 */
function valueKey(value: unknown): string | null {
  const text = cellText(value);
  if (text === "") return null;
  return text.toLowerCase();
}

/** True when a rule is (or defaults to) the comparison kind. */
export function isComparisonRule(
  rule: ConditionalFormatRule,
): rule is ConditionalComparisonRule {
  return rule.kind === undefined || rule.kind === "comparison";
}

/**
 * Evaluate a COMPARISON rule against a cell value. Returns true only when the
 * value coerces to a finite number AND satisfies the operator. A BLANK cell
 * (null) reads as 0, as Excel's Cell Value rule compares it (`=A1=0` holds
 * for an empty A1: P18_48539's "equal to 0" fills its empty Pass/Fail cells,
 * X-060); a stored empty string is text and, like any non-numeric value,
 * never matches. A non-comparison rule never matches here
 * (formula/scale/bar/icon kinds are resolved via resolveCellConditionalDisplay).
 * The parameter accepts the full union for back-compat; the comparison math is
 * byte-for-byte unchanged.
 */
export function evaluateRule(
  rule: ConditionalFormatRule,
  value: unknown,
): boolean {
  if (!isComparisonRule(rule)) return false;
  // Text never matches, a number written as text included: a formula's "1"
  // is not equal to 1 in Excel (lookwin R139-X01: IF(..., "1") quartiles
  // under an "equal to 1" fill print unfilled). The calculation hands a
  // number cell's value over as a number.
  if (typeof value === "string") return false;
  const n = value === null || value === undefined ? 0 : toNumber(value);
  if (n === null) return false;
  switch (rule.operator) {
    case ">":
      return n > rule.value;
    case "<":
      return n < rule.value;
    case ">=":
      return n >= rule.value;
    case "<=":
      return n <= rule.value;
    case "=":
      return n === rule.value;
    case "!=":
      return n !== rule.value;
    case "between":
      return rule.value2 !== undefined && n >= rule.value && n <= rule.value2;
    case "notBetween":
      return rule.value2 !== undefined && (n < rule.value || n > rule.value2);
    default:
      return false;
  }
}

/**
 * Evaluate a textContains rule against a cell value (case-insensitive, Excel
 * parity). Excel stores `contains` as NOT(ISERROR(SEARCH(text,A1))) and
 * `notContains` as ISERROR(SEARCH(text,A1)), so both follow SEARCH: `?` is any
 * one character, `*` any run, `~` escapes a literal ?, * or ~ (live audit
 * B036). `beginsWith`/`endsWith` are Excel's LEFT/RIGHT equality — literal.
 * `notContains` matches a blank cell too. The needle is used exactly as typed.
 */
export function evaluateTextContains(
  rule: ConditionalTextContainsRule,
  value: unknown,
): boolean {
  const hay = cellSearchText(value).toLowerCase();
  const needle = rule.text.toLowerCase();
  switch (rule.operator) {
    case "contains":
      return searchPattern(rule.text).test(hay);
    case "notContains":
      return !searchPattern(rule.text).test(hay);
    case "beginsWith":
      return hay.startsWith(needle);
    case "endsWith":
      return hay.endsWith(needle);
    default:
      return false;
  }
}

const searchPatternCache = new Map<string, RegExp>();

/** SEARCH's wildcard needle as a case-insensitive regex (memoized per needle). */
function searchPattern(needle: string): RegExp {
  let re = searchPatternCache.get(needle);
  if (!re) {
    if (searchPatternCache.size > 256) searchPatternCache.clear();
    re = new RegExp(excelWildcardToRegexSource(needle), "is");
    searchPatternCache.set(needle, re);
  }
  return re;
}

/**
 * Parse a cell's display text into calendar parts for a date-occurring rule, or
 * null when it is not a recognized date. Deterministic (no `Date` locale parsing):
 * ISO `YYYY-MM-DD` / `YYYY-MM` / `YYYY` and US `M/D/YYYY`; month/day range-checked.
 */
function parseCellDate(
  text: string,
): { year: number; month: number; day: number } | null {
  const s = text.trim();
  let m: RegExpMatchArray | null;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/)))
    return inRange(+m[1], +m[2], +m[3]);
  if ((m = s.match(/^(\d{4})-(\d{1,2})$/))) return inRange(+m[1], +m[2], 1);
  if ((m = s.match(/^(\d{4})$/))) return inRange(+m[1], 1, 1);
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/)))
    return inRange(+m[3], +m[1], +m[2]); // US M/D/YYYY
  return null;
}

function inRange(
  year: number,
  month: number,
  day: number,
): { year: number; month: number; day: number } | null {
  if (month < 1 || month > 12 || day < 1) return null;
  
  // as a day of September (and so as "last month") just because 31 is a day
  // some months have (live audit B364).
  const leap = (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
  const days = month === 2 ? (leap ? 29 : 28) : [4, 6, 9, 11].includes(month) ? 30 : 31;
  if (day > days) return null;
  return { year, month, day };
}

/** The UTC day index of a calendar date — a stable integer for day arithmetic. */
function dayIndexUTC(year: number, month: number, day: number): number {
  return Math.floor(Date.UTC(year, month - 1, day) / 86_400_000);
}

/** Shift a (year, 1-based month) by `delta` months, normalizing the year. */
function shiftMonth(
  year: number,
  month: number,
  delta: number,
): { year: number; month: number } {
  const idx = year * 12 + (month - 1) + delta;
  return { year: Math.floor(idx / 12), month: (((idx % 12) + 12) % 12) + 1 };
}

/**
 * Whether a parsed cell date falls in `period` relative to `today`. Day-of-week
 * uses Sunday-start weeks (Excel default). Pure — `today` is supplied, never read
 * from the clock here — so the period math is fully unit-testable.
 */
export function dateInPeriod(
  cell: { year: number; month: number; day: number },
  period: DateOccurringPeriod,
  today: { year: number; month: number; day: number },
): boolean {
  const ci = dayIndexUTC(cell.year, cell.month, cell.day);
  const ti = dayIndexUTC(today.year, today.month, today.day);
  const weekday = new Date(
    Date.UTC(today.year, today.month - 1, today.day),
  ).getUTCDay();
  const weekStart = ti - weekday; // Sunday of today's week
  switch (period) {
    case "today":
      return ci === ti;
    case "yesterday":
      return ci === ti - 1;
    case "tomorrow":
      return ci === ti + 1;
    case "last7Days":
      return ci <= ti && ci >= ti - 6;
    case "thisWeek":
      return ci >= weekStart && ci <= weekStart + 6;
    case "lastWeek":
      return ci >= weekStart - 7 && ci <= weekStart - 1;
    case "nextWeek":
      return ci >= weekStart + 7 && ci <= weekStart + 13;
    case "thisMonth":
      return cell.year === today.year && cell.month === today.month;
    case "lastMonth": {
      const p = shiftMonth(today.year, today.month, -1);
      return cell.year === p.year && cell.month === p.month;
    }
    case "nextMonth": {
      const n = shiftMonth(today.year, today.month, 1);
      return cell.year === n.year && cell.month === n.month;
    }
    default:
      return false;
  }
}

/**
 * The calendar day a cell value names for a date-occurring rule. A number is
 * a date serial in the workbook's date system (Excel's 1900 system, or 1904
 * where serial 0 is 1904-01-01), which is what a typed date, a date formula
 * (TODAY(), DATE(...)) or a date-time holds; the time of day is dropped, as
 * Excel's own `FLOOR(A1,1)=TODAY()` reading does. Only a text value goes
 * through the explicit text shapes of {@link parseCellDate}.
 */
function cellDateParts(
  value: unknown,
  date1904: ExcelDateSystemArg = false,
): { year: number; month: number; day: number } | null {
  if (typeof value === "number") {
    if (
      !Number.isFinite(value) ||
      value < excelMinDateSerial(date1904) ||
      value >= excelMaxDateSerial(date1904) + 1
    ) {
      return null;
    }
    const { year, month, day } = excelSystemSerialToParts(value, date1904);
    return { year, month, day };
  }
  return parseCellDate(cellText(value));
}

/**
 * Whether a cell value matches a date-occurring rule, with `now` as "today"
 * (default: the current date, so the rule advances with the clock). A non-date
 * cell never matches.
 */
export function evaluateDateOccurring(
  rule: ConditionalDateOccurringRule,
  value: unknown,
  now?: Date,
  date1904: ExcelDateSystemArg = false,
): boolean {
  const cell = cellDateParts(value, date1904);
  if (!cell) return false;
  const t = now ?? new Date();
  return dateInPeriod(
    cell,
    rule.period,
    { year: t.getFullYear(), month: t.getMonth() + 1, day: t.getDate() },
  );
}

/**
 * Resolve the conditional COLOR style for a cell at (row, col) with the given
 * value, considering only comparison rules. The FIRST comparison rule (by array
 * order) that both covers the cell and matches the value supplies the whole
 * style (its textColor + backgroundColor). Returns an empty object when none
 * applies, so the caller can spread it harmlessly.
 *
 * Kept for back-compat with callers that only handle comparison colors. The
 * richer entry point resolveCellConditionalDisplay also handles formula /
 * colorScale colors plus data bars and icons; new code should prefer it. Colors
 * stored on a normalized rule are already validated #rrggbb, so the returned
 * style is always safe to spread into an inline CSS object.
 */
export function resolveCellConditionalStyle(
  rules: readonly ConditionalFormatRule[] | undefined,
  value: unknown,
  row: number,
  col: number,
): ConditionalStyle {
  if (!rules || rules.length === 0) return {};
  for (const rule of rules) {
    if (!ruleContains(rule, row, col)) continue;
    if (!isComparisonRule(rule)) continue;
    if (!evaluateRule(rule, value)) continue;
    return applyDifferentialCss({}, rule);
  }
  return {};
}

// ---------------------------------------------------------------------------
// Range statistics + per-kind math for scale / bar / icon rules.
// ---------------------------------------------------------------------------

/** The numeric min/max over a range; null when the range has no numeric cell. */
export interface RangeExtent {
  min: number;
  max: number;
}

/**
 * Compute the numeric min/max of the values in a range. `cellValueAt(row, col)`
 * returns a cell's resolved value (number, string, blank, …); non-numeric cells
 * are skipped (Excel parity: scales/bars/icons act on numbers only). Returns
 * null when the range contains no numeric value. The range is already clamped to
 * the grid, so the scan is bounded.
 */
export function computeRangeExtent(
  range: ConditionalRange,
  cellValueAt: (row: number, col: number) => unknown,
): RangeExtent | null {
  let min = Infinity;
  let max = -Infinity;
  let seen = false;
  for (let r = range.r1; r <= range.r2; r += 1) {
    for (let c = range.c1; c <= range.c2; c += 1) {
      const n = numericCellValue(cellValueAt(r, c));
      if (n === null) continue;
      seen = true;
      if (n < min) min = n;
      if (n > max) max = n;
    }
  }
  return seen ? { min, max } : null;
}

/**
 * Range-wide statistics shared by the duplicate / topBottom / aboveBelowAverage
 * rules: how many times each value-key occurs (for duplicate/unique detection),
 * the sorted numeric values (for top/bottom cutoffs), and the numeric mean +
 * population standard deviation (for above/below average). Computed ONCE per
 * range (memoized in `statsCache`) so a whole-range scan never runs per cell.
 */
/** The numeric min/max over several rectangles (a multi-area rule). */
export function computeRangesExtent(
  ranges: readonly ConditionalRange[],
  cellValueAt: (row: number, col: number) => unknown,
): RangeExtent | null {
  let out: RangeExtent | null = null;
  for (const range of ranges) {
    const extent = computeRangeExtent(range, cellValueAt);
    if (!extent) continue;
    out = out
      ? { min: Math.min(out.min, extent.min), max: Math.max(out.max, extent.max) }
      : extent;
  }
  return out;
}

export interface RangeStats {
  /** value-key (trimmed, lowercased text) → occurrence count over the range. */
  counts: Map<string, number>;
  /** Every numeric value in the range, ASCENDING (ties preserved). */
  sortedNumbers: number[];
  /** Arithmetic mean of the numeric values, or null when there are none. */
  mean: number | null;
  /** Population standard deviation of the numeric values (0 when <2 numbers). */
  stdDev: number;
}

/**
 * Scan a range once and compute its {@link RangeStats}. Each cell contributes to
 * the value-key count (blanks skipped, Excel parity) and, when numeric, to the
 * sorted-number list / mean / std-dev. The range is already grid-clamped so the
 * scan is bounded.
 */
export function computeRangeStats(
  range: ConditionalRange,
  cellValueAt: (row: number, col: number) => unknown,
): RangeStats {
  return computeRangesStats([range], cellValueAt);
}

/**
 * {@link computeRangeStats} over several rectangles (a multi-area rule). The
 * areas are one "Applies to" union, as in Excel: a cell listed by two
 * overlapping areas (`A1:A2,A2:A3`) is still one cell, so it counts once
 * toward duplicates, ranks, the average and percentiles.
 */
export function computeRangesStats(
  ranges: readonly ConditionalRange[],
  cellValueAt: (row: number, col: number) => unknown,
): RangeStats {
  const counts = new Map<string, number>();
  const numbers: number[] = [];
  let sum = 0;
  for (let index = 0; index < ranges.length; index += 1) {
    const range = ranges[index];
    const earlier = ranges
      .slice(0, index)
      .filter((area) => rangesOverlap(area, range));
    for (let r = range.r1; r <= range.r2; r += 1) {
      for (let c = range.c1; c <= range.c2; c += 1) {
        if (earlier.some((area) => rangeContains(area, r, c))) continue;
        const v = cellValueAt(r, c);
        const key = valueKey(v);
        if (key !== null) counts.set(key, (counts.get(key) ?? 0) + 1);
        const n = numericCellValue(v);
        if (n !== null) {
          numbers.push(n);
          sum += n;
        }
      }
    }
  }
  numbers.sort((a, b) => a - b);
  const mean = numbers.length > 0 ? sum / numbers.length : null;
  let variance = 0;
  if (numbers.length > 0 && mean !== null) {
    for (const n of numbers) variance += (n - mean) * (n - mean);
    variance /= numbers.length;
  }
  return { counts, sortedNumbers: numbers, mean, stdDev: Math.sqrt(variance) };
}

/**
 * Whether a cell's value is a DUPLICATE (occurs >1×) — or, when `unique` is set,
 * a UNIQUE value (occurs exactly 1×) — within the range, given the precomputed
 * value-key counts. A blank cell (null key) is never highlighted either way.
 */
export function evaluateDuplicateValues(
  rule: ConditionalDuplicateValuesRule,
  value: unknown,
  stats: RangeStats,
): boolean {
  const key = valueKey(value);
  if (key === null) return false;
  const count = stats.counts.get(key) ?? 0;
  return rule.unique ? count === 1 : count > 1;
}

/**
 * Whether a numeric value falls in the rule's top/bottom band, given the range's
 * sorted numbers. `rank` is the item count (or, when `percent`, a percentage of
 * the numeric cell count, rounded down to ≥1). Ties are INCLUSIVE: the cutoff is
 * the value at the rank boundary and any value at-or-beyond it matches, so "top
 * 2" of [9, 9, 1] highlights both 9s. A non-numeric value never matches.
 */
export function isInTopBottom(
  rule: ConditionalTopBottomRule,
  value: unknown,
  stats: RangeStats,
): boolean {
  const n = numericCellValue(value);
  if (n === null) return false;
  const sorted = stats.sortedNumbers;
  const total = sorted.length;
  if (total === 0) return false;
  let count = rule.rank;
  if (rule.percent) count = Math.floor((total * rule.rank) / 100);
  if (count < 1) count = 1;
  if (count >= total) return true; // whole range highlighted
  if (rule.bottom) {
    // The N smallest: cutoff is the Nth value from the low end (inclusive).
    const cutoff = sorted[count - 1];
    return n <= cutoff;
  }
  // The N largest: cutoff is the Nth value from the high end (inclusive).
  const cutoff = sorted[total - count];
  return n >= cutoff;
}

/**
 * Whether a numeric value is above (default) or below the range's average, given
 * the range stats. An optional `stdDevs` (k) shifts the threshold by k·σ — above
 * uses mean + k·σ, below uses mean − k·σ. The comparison is STRICT. Returns
 * false for a non-numeric value or an all-non-numeric range (no mean).
 */
export function isAboveBelowAverage(
  rule: ConditionalAboveBelowAverageRule,
  value: unknown,
  stats: RangeStats,
): boolean {
  const n = numericCellValue(value);
  if (n === null || stats.mean === null) return false;
  const offset = (rule.stdDevs ?? 0) * stats.stdDev;
  // Excel's "equal or above / equal or below": the threshold itself matches.
  if (rule.below) {
    return rule.equalAverage ? n <= stats.mean - offset : n < stats.mean - offset;
  }
  return rule.equalAverage ? n >= stats.mean + offset : n > stats.mean + offset;
}

/**
 * Where a value sits in the range span as a fraction in [0, 1]: min → 0, max →
 * 1, clamped. A degenerate span (min === max) yields 1 so every cell renders a
 * full bar / the top icon bucket, matching Excel's behavior for a flat range.
 */
function fractionInExtent(n: number, extent: RangeExtent): number {
  const span = extent.max - extent.min;
  if (span <= 0) return 1;
  const f = (n - extent.min) / span;
  return f < 0 ? 0 : f > 1 ? 1 : f;
}

/** Parse a validated #rrggbb (already normalized) into [r, g, b] 0..255. */
function hexToRgb(hex: string): [number, number, number] {
  return [
    parseInt(hex.slice(1, 3), 16),
    parseInt(hex.slice(3, 5), 16),
    parseInt(hex.slice(5, 7), 16),
  ];
}

/** Serialize [r, g, b] (rounded, clamped 0..255) back to a #rrggbb string. */
function rgbToHex(r: number, g: number, b: number): string {
  const ch = (v: number) =>
    Math.max(0, Math.min(255, Math.round(v)))
      .toString(16)
      .padStart(2, "0");
  return `#${ch(r)}${ch(g)}${ch(b)}`;
}

/** Linear-interpolate two #rrggbb colors at fraction t (0 → a, 1 → b). */
function lerpHex(a: string, b: string, t: number): string {
  const [ar, ag, ab] = hexToRgb(a);
  const [br, bg, bb] = hexToRgb(b);
  return rgbToHex(ar + (br - ar) * t, ag + (bg - ag) * t, ab + (bb - ab) * t);
}

/**
 * Interpolate a color-scale rule's background color for a numeric value. For a
 * 2-point scale the value's fraction interpolates min.color → max.color. For a
 * 3-point scale the lower half (fraction ≤ 0.5) interpolates min → mid and the
 * upper half mid → max, with the midpoint mapping to mid.color exactly.
 * Endpoints are exact: the range min returns min.color and the range max returns
 * max.color. Returns the validated #rrggbb result.
 */
export function interpolateColorScale(
  rule: ConditionalColorScaleRule,
  value: number,
  extent: RangeExtent,
  sortedNumbers?: readonly number[],
): string {
  const thresholds = colorScaleThresholds(rule, extent, sortedNumbers);
  if (!thresholds) {
    // Default endpoints, default midpoint: the historical fraction geometry.
    const f = fractionInExtent(value, extent);
    if (!rule.mid) return lerpHex(rule.min.color, rule.max.color, f);
    if (f <= 0.5) return lerpHex(rule.min.color, rule.mid.color, f / 0.5);
    return lerpHex(rule.mid.color, rule.max.color, (f - 0.5) / 0.5);
  }
  const stops = rule.mid ? [rule.min, rule.mid, rule.max] : [rule.min, rule.max];
  // At or past the top threshold first: a value on thresholds that all
  // coincide (one number in the range, min = percentile 50 = max) takes the
  // top colour, as Excel paints it (S30_tdf156028, X-204) and as the default
  // geometry already did (fractionInExtent of an empty span is 1).
  const last = thresholds.length - 1;
  if (value >= thresholds[last]) return stops[last].color;
  if (value <= thresholds[0]) return stops[0].color;
  for (let index = 1; index <= last; index += 1) {
    const lower = thresholds[index - 1];
    const upper = thresholds[index];
    if (value <= upper) {
      const span = upper - lower;
      const t = span > 0 ? (value - lower) / span : 1;
      return lerpHex(stops[index - 1].color, stops[index].color, t);
    }
  }
  return stops[last].color;
}

/**
 * Resolve every stop to a number on the value axis, ascending, or null when
 * every stop is the position default (min / 50th-percentile-as-half / max),
 * which keeps the exact historical geometry. Percentile stops need the
 * range's sorted numbers; a formula stop has no evaluator here and stands in
 * for the endpoint it replaces (its midpoint counterpart is the half-way
 * value), which is the honest reading without the sheet's engine.
 */
function colorScaleThresholds(
  rule: ConditionalColorScaleRule,
  extent: RangeExtent,
  sortedNumbers: readonly number[] | undefined,
): number[] | null {
  const stops = rule.mid ? [rule.min, rule.mid, rule.max] : [rule.min, rule.max];
  // An untyped midpoint is Excel's default one, the 50th PERCENTILE of the
  // range (XLS-A047, the xlsx projection stores `percentile 50` that way):
  // with -60…100 it is 37.5, not the half-way value 20.
  const medianMid =
    rule.mid !== undefined && !rule.mid.type && sortedNumbers && sortedNumbers.length > 0
      ? percentileOf(sortedNumbers, 50)
      : undefined;
  if (medianMid === undefined && stops.every((stop) => !stop.type)) return null;
  const span = extent.max - extent.min;
  const resolve = (stop: ColorScaleStop, position: "min" | "mid" | "max"): number => {
    const fallback =
      position === "min"
        ? extent.min
        : position === "max"
          ? extent.max
          : medianMid ?? extent.min + span / 2;
    switch (stop.type) {
      case "min":
        return extent.min;
      case "max":
        return extent.max;
      case "number":
        return stop.value ?? fallback;
      case "percent":
        return stop.value === undefined ? fallback : extent.min + (span * stop.value) / 100;
      case "percentile": {
        if (stop.value === undefined || !sortedNumbers || sortedNumbers.length === 0) {
          return fallback;
        }
        return percentileOf(sortedNumbers, stop.value);
      }
      case "formula":
      default:
        return fallback;
    }
  };
  const thresholds = stops.map((stop, index) =>
    resolve(stop, index === 0 ? "min" : index === stops.length - 1 ? "max" : "mid"),
  );
  // Out-of-order values are put in ascending order while the colours keep
  // their places: minimum 0 / midpoint 4 / maximum 3 is green at 3 and red at
  // 4, and a 2-point scale from 5 to 2 runs from 2 (MEASURED in Excel for
  
  return thresholds.sort((a, b) => a - b);
}

/** Excel's PERCENTILE.INC over an ascending list (linear interpolation). */
function percentileOf(sorted: readonly number[], percent: number): number {
  const k = Math.max(0, Math.min(100, percent)) / 100;
  const position = (sorted.length - 1) * k;
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  if (lower === upper) return sorted[lower];
  return sorted[lower] + (sorted[upper] - sorted[lower]) * (position - lower);
}

/**
 * Classic SpreadsheetML data-bar geometry. A one-sided range maps its minimum
 * and maximum to minLength/maxLength (10%/90% by default); a mixed-sign range
 * places an automatic zero axis and sends negative/positive bars to opposite
 * sides without drawing a bar for zero itself.
 */
export function dataBarGeometry(
  value: number,
  extent: RangeExtent,
  minLength = 10,
  maxLength = 90,
  axisPosition: "automatic" | "middle" | "none" = "automatic",
): Pick<
  ConditionalDataBar,
  "widthPercent" | "axisPercent" | "startPercent" | "direction"
> {
  const min = extent.min;
  const max = extent.max;
  const span = max - min;
  // Excel's axis placements: automatic puts the axis where zero falls,
  // cell midpoint always in the middle, none draws every bar from the left.
  const crossesZero = axisPosition === "middle" ? min < 0 : axisPosition !== "none" && min < 0 && max > 0;
  const allNegative = axisPosition !== "none" && max <= 0;
  const axisPercent = axisPosition === "middle" && min < 0
    ? 50
    : crossesZero
      ? (-min / span) * 100
      : allNegative
        ? 100
        : 0;
  // With no axis every bar runs rightward from the left edge, a negative too.
  const direction = value < 0 && axisPosition !== "none" ? "negative" : "positive";
  const boundedMinLength = Math.max(0, Math.min(100, minLength));
  const boundedMaxLength = Math.max(
    boundedMinLength,
    Math.min(100, maxLength),
  );
  let fraction: number;
  const available = direction === "negative" ? axisPercent : 100 - axisPercent;
  if (value === 0 && crossesZero) {
    fraction = 0;
  } else if (span === 0) {
    fraction = 1;
  } else if (crossesZero) {
    const sideMax = direction === "negative" ? Math.abs(min) : max;
    fraction = sideMax === 0 ? 0 : Math.abs(value) / sideMax;
  } else if (allNegative) {
    fraction = (max - value) / span;
  } else {
    fraction = (value - min) / span;
  }
  const zeroAxisValue = crossesZero && value === 0;
  // The minimum and maximum lengths are shares of the room the bar has: the
  // whole cell, or the side of the axis it grows into. Scaled on the whole
  // cell and then cut at the axis, -5 and -10 (and 5 and 10) of a -10..10
  // range all came out half a cell long (user audit S-028).
  let widthPercent = zeroAxisValue
    ? 0
    : (available / 100) *
      (boundedMinLength +
        Math.max(0, Math.min(1, fraction)) *
          (boundedMaxLength - boundedMinLength));
  widthPercent = Math.max(0, Math.min(available, widthPercent));
  const startPercent =
    direction === "negative"
      ? Math.max(0, axisPercent - widthPercent)
      : axisPercent;
  return { widthPercent, axisPercent, startPercent, direction };
}

/** Backwards-compatible normalized fraction helper used by model callers. */
export function dataBarWidthPercent(value: number, extent: RangeExtent): number {
  return fractionInExtent(value, extent) * 100;
}

/**
 * The icon index (0 .. count-1) for a numeric value given the rule's (count-1)
 * ascending thresholds (fractions of the range span). The value's fraction is
 * compared against each threshold in order: below threshold i and no earlier
 * threshold crossed → bucket i; at or above every threshold → the top bucket
 * (count-1). With the 3-icon defaults [1/3, 2/3] this splits the range into low /
 * mid / high thirds; a 5-icon set's [1/5 … 4/5] splits into fifths.
 *
 * When the family's `reverse` flag is set the bucket→icon mapping is flipped so
 * the HIGHEST values get icon 0 (Excel "Reverse Icon Order"): the returned index
 * becomes (count - 1 - bucket). The icon count is derived from the family so the
 * top bucket is always count-1 even if `thresholds` is a stale length.
 */
export function iconBucketIndex(
  rule: ConditionalIconSetRule,
  value: number,
  extent: RangeExtent,
  sortedNumbers?: readonly number[],
): number {
  const count = iconSetCount(rule.iconSet);
  let bucket = count - 1;
  if (rule.stops) {
    // Excel decides from the top icon down: the first boundary the value
    // satisfies (">=" or, for a ">" boundary, strictly above) names the
    // bucket; below every boundary is the lowest icon.
    const span = extent.max - extent.min;
    const defaults = defaultIconThresholds(count);
    bucket = 0;
    for (let i = Math.min(rule.stops.length, count - 1) - 1; i >= 0; i -= 1) {
      const stop = rule.stops[i];
      let boundary: number;
      switch (stop.type) {
        case "number":
          boundary = stop.value ?? extent.min + span * defaults[i];
          break;
        case "percent":
          boundary = extent.min + (span * (stop.value ?? defaults[i] * 100)) / 100;
          break;
        case "percentile":
          boundary =
            stop.value !== undefined && sortedNumbers && sortedNumbers.length > 0
              ? percentileOf(sortedNumbers, stop.value)
              : extent.min + span * defaults[i];
          break;
        default:
          // A formula boundary is not evaluated here; the family's default
          // split stands in for it on screen.
          boundary = extent.min + span * defaults[i];
      }
      if (stop.gte === false ? value > boundary : value >= boundary) {
        bucket = i + 1;
        break;
      }
    }
    return rule.reverse ? count - 1 - bucket : bucket;
  }
  const f = fractionInExtent(value, extent);
  for (let i = 0; i < rule.thresholds.length && i < count - 1; i += 1) {
    if (f < rule.thresholds[i]) {
      bucket = i;
      break;
    }
  }
  return rule.reverse ? count - 1 - bucket : bucket;
}

/**
 * Options for resolveCellConditionalDisplay. Scale / data-bar / icon rules need
 * the range's numeric min/max, so the caller supplies `cellValueAt` to read the
 * range's cells; the resolver memoizes each range's extent in `extentCache` (one
 * Map per render pass) so the scan runs once per rule, not once per cell. Formula
 * rules are evaluated through `evalFormula` (the spreadsheet formula engine,
 * threaded in so this module never imports it).
 */
export interface ConditionalDisplayOptions {
  /** Resolve any cell's value for range-extent scans (scale/bar/icon). */
  cellValueAt?: (row: number, col: number) => unknown;
  /**
   * Evaluate a formula rule for the cell at (row, col). The anchor (the range's
   * top-left) is passed so the caller can shift relative references Excel-style.
   * A truthy return applies the rule. When omitted, formula rules never match.
   */
  evalFormula?: (
    formula: string,
    row: number,
    col: number,
    anchorRow: number,
    anchorCol: number,
  ) => unknown;
  /** Per-render memo of each rule's range extent (avoids re-scanning per cell). */
  extentCache?: Map<ConditionalFormatRule, RangeExtent | null>;
  /**
   * Per-render memo of each rule's range statistics (counts / sorted numbers /
   * mean / std-dev) for duplicate / topBottom / aboveBelowAverage rules, so the
   * whole-range scan runs once per rule, not once per cell. Like `extentCache`,
   * one Map per render pass.
   */
  statsCache?: Map<ConditionalFormatRule, RangeStats | null>;
  /**
   * "Today" for date-occurring rules (Excel "A Date Occurring…"), so the relative
   * period (today / this week / last month / …) advances with the clock. Defaults
   * to the current date; injected in tests for determinism. Other rule kinds
   * ignore it.
   */
  now?: Date;
  /** The workbook reads serials in the 1904 date system (date-occurring). */
  date1904?: ExcelDateSystemArg;
}

/**
 * Interpret a formula engine result as the rule's condition, the way Excel
 * reads a value as a logical: TRUE and a non-zero number match; FALSE, 0,
 * a blank and an error do not. Text is read as IF reads it: "TRUE"/"FALSE"
 * are logicals, a numeric text is its number, and any other text ("invalid")
 * is no logical at all, so it never matches.
 */
export function isFormulaTruthy(result: unknown): boolean {
  if (result === null || result === undefined) return false;
  if (typeof result === "boolean") return result;
  if (typeof result === "number") return Number.isFinite(result) && result !== 0;
  if (typeof result === "string") {
    const s = result.trim();
    if (s === "") return false;
    if (s.toUpperCase() === "FALSE") return false;
    if (s.toUpperCase() === "TRUE") return true;
    const number = Number(s);
    return Number.isFinite(number) && number !== 0;
  }
  // A FormulaError object (or any other shape) does not match.
  return false;
}

/** Look up (and memoize) a rule's range extent via the options' cellValueAt. */
/**
 * A formula-typed stop (`<cfvo type="formula">`) is one value for the whole
 * range. Excel's dialog refuses a relative reference there; in a file, a
 * reference with a relative part reads as blank while an absolute one reads
 * its cell (measured in Excel: `A1*2+8` gives 8 with A1 = 11, `$A$1*2+8`
 * gives 30). The stop then acts as that number; a formula that yields no
 * number keeps the default.
 */
type FormulaStopLike = { type?: string; formula?: string; value?: number };

function columnLetters(index: number): string {
  let letters = "";
  for (let rest = index + 1; rest > 0; rest = Math.floor((rest - 1) / 26)) {
    letters = String.fromCharCode(65 + ((rest - 1) % 26)) + letters;
  }
  return letters;
}

/** `formula` with every reference that has a relative part moved to a blank
 * area of the same size in the sheet's far corner. */
export function blankRelativeStopReferences(formula: string): string {
  const entry = `=${formula}`;
  let result = entry;
  for (const span of formulaReferenceSpans(entry).reverse()) {
    const text = entry.slice(span.start, span.end);
    const local = text.slice(text.lastIndexOf("!") + 1);
    const absolute = local.split(":").every((part) => /^\$[A-Za-z]+\$\d+$|^\$[A-Za-z]+$|^\$\d+$/.test(part));
    if (absolute) continue;
    let corner: string;
    if (span.r2 === Number.POSITIVE_INFINITY) {
      const width = span.c2 - span.c1;
      corner = `$${columnLetters(MAX_GRID_COLUMNS - 1 - width)}:$${columnLetters(MAX_GRID_COLUMNS - 1)}`;
    } else if (span.c2 === Number.POSITIVE_INFINITY) {
      const height = span.r2 - span.r1;
      corner = `$${MAX_GRID_ROWS - height}:$${MAX_GRID_ROWS}`;
    } else {
      const top = MAX_GRID_ROWS - (span.r2 - span.r1);
      const left = MAX_GRID_COLUMNS - 1 - (span.c2 - span.c1);
      corner = `$${columnLetters(left)}$${top}:$${columnLetters(MAX_GRID_COLUMNS - 1)}$${MAX_GRID_ROWS}`;
    }
    result = result.slice(0, span.start) + corner + result.slice(span.end);
  }
  return result.slice(1);
}

function withFormulaStopsResolved<
  T extends { range: ConditionalRange; min?: FormulaStopLike; mid?: FormulaStopLike; max?: FormulaStopLike },
>(rule: T, row: number, col: number, opts: ConditionalDisplayOptions): T {
  const evalFormula = opts.evalFormula;
  if (!evalFormula) return rule;
  const resolve = <S extends FormulaStopLike>(stop: S | undefined): S | undefined => {
    if (!stop || stop.type !== "formula" || !stop.formula) return stop;
    const formula = blankRelativeStopReferences(stop.formula);
    const evaluated = toNumber(evalFormula(formula, rule.range.r1, rule.range.c1, rule.range.r1, rule.range.c1));
    return evaluated === null ? stop : { ...stop, type: "number", value: evaluated };
  };
  const min = resolve(rule.min);
  const mid = resolve(rule.mid);
  const max = resolve(rule.max);
  if (min === rule.min && mid === rule.mid && max === rule.max) return rule;
  return { ...rule, ...(min ? { min } : {}), ...(mid ? { mid } : {}), ...(max ? { max } : {}) } as T;
}

/** An icon rule with every formula boundary evaluated for this cell, like
 * {@link withFormulaStopsResolved} does for a colour scale's stops. A
 * threshold formula that does not produce a number (=NA(), text) leaves the
 * rule without a usable boundary: it paints nothing rather than the family's
 * default split (live audit B121). Null means "no icon for this cell". */
function withIconFormulaStopsResolved(
  rule: ConditionalIconSetRule,
  row: number,
  col: number,
  opts: ConditionalDisplayOptions,
): ConditionalIconSetRule | null {
  const evalFormula = opts.evalFormula;
  if (!evalFormula || !rule.stops?.some((stop) => stop.type === "formula")) return rule;
  let failed = false;
  const stops = rule.stops.map((stop) => {
    if (stop.type !== "formula" || !stop.formula) return stop;
    const result = evalFormula(stop.formula, row, col, rule.range.r1, rule.range.c1);
    // No answer yet (the calculation has not evaluated the threshold: the
    // host reads a pending formula as null, while a blank reference answers
    // 0) is no boundary: the cell shows no icon until it lands. Read as 0,
    // every value passed both thresholds and PH052's 3Flags column showed
    // the green flag on every row (X-277).
    if (result === null || result === undefined) {
      failed = true;
      return stop;
    }
    // A blank result reads as 0, as a blank cell does in a comparison.
    const evaluated = result === "" ? 0 : toNumber(result);
    if (evaluated === null) {
      failed = true;
      return stop;
    }
    const resolved: IconThreshold = { type: "number", value: evaluated };
    if (stop.gte === false) resolved.gte = false;
    return resolved;
  });
  if (failed) return null;
  return { ...rule, stops };
}

function extentFor(
  rule: { range: ConditionalRange; areas?: ConditionalRange[] },
  opts: ConditionalDisplayOptions,
): RangeExtent | null {
  const cache = opts.extentCache;
  if (cache && cache.has(rule as ConditionalFormatRule)) {
    return cache.get(rule as ConditionalFormatRule) ?? null;
  }
  const cellValueAt = opts.cellValueAt;
  const extent = cellValueAt
    ? computeRangesExtent(ruleAreas(rule), cellValueAt)
    : null;
  if (cache) cache.set(rule as ConditionalFormatRule, extent);
  return extent;
}

/**
 * Look up (and memoize) a rule's range statistics via the options' cellValueAt.
 * Returns null when no cellValueAt is wired (duplicate/topBottom/average rules
 * then never match), mirroring extentFor.
 */
function statsFor(
  rule: { range: ConditionalRange; areas?: ConditionalRange[] },
  opts: ConditionalDisplayOptions,
): RangeStats | null {
  const cache = opts.statsCache;
  if (cache && cache.has(rule as ConditionalFormatRule)) {
    return cache.get(rule as ConditionalFormatRule) ?? null;
  }
  const cellValueAt = opts.cellValueAt;
  const stats = cellValueAt ? computeRangesStats(ruleAreas(rule), cellValueAt) : null;
  if (cache) cache.set(rule as ConditionalFormatRule, stats);
  return stats;
}

const CONDITIONAL_STYLE_SINGLE_CHANNELS = [
  "color",
  "fontWeight",
  "fontStyle",
  "textDecoration",
  "fontSize",
  "fontFamily",
  "numberFormat",
] as const;

/**
 * Fold one matching rule's differential style into the cell's resolved
 * style. Excel applies every matching rule whose formats do not conflict and
 * lets the rule higher in precedence win where they do (a bold rule and a red
 * text rule both apply; of two font colours the higher rule's shows), so each
 * channel belongs to the first rule that sets it: font colour, the fill (a
 * solid colour and a gradient are one fill), each font member, the number
 * format, and each border edge on its own.
 */
function mergeConditionalStyle(
  resolved: ConditionalStyle,
  style: ConditionalStyle,
): void {
  for (const key of CONDITIONAL_STYLE_SINGLE_CHANNELS) {
    if (style[key] !== undefined && resolved[key] === undefined) {
      (resolved as Record<string, unknown>)[key] = style[key];
    }
  }
  if (
    (style.backgroundColor || style.backgroundImage) &&
    !resolved.backgroundColor &&
    !resolved.backgroundImage
  ) {
    if (style.backgroundColor) resolved.backgroundColor = style.backgroundColor;
    if (style.backgroundImage) resolved.backgroundImage = style.backgroundImage;
    if (style.backgroundSize) resolved.backgroundSize = style.backgroundSize;
  }
  if (style.borderEdges) {
    const edges: CellBorderEdges = { ...(resolved.borderEdges ?? {}) };
    let added = false;
    for (const edge of BORDER_EDGES) {
      const line = style.borderEdges[edge];
      if (!line || edges[edge]) continue;
      edges[edge] = line;
      added = true;
    }
    if (added) resolved.borderEdges = edges;
  }
}

/**
 * Resolve the FULL conditional-format display for a cell at (row, col), in
 * array/native-priority order. Matching rules may contribute independent
 * visual channels until a matching Stop If True rule ends evaluation; a lower
 * priority rule never overwrites a channel already supplied by a higher one.
 * Dispatches on rule kind:
 *  - comparison: numeric comparison on `value` → text/background color.
 *  - formula:    evalFormula(...) truthy → text/background color.
 *  - colorScale: interpolated background color from the range's min/mid/max.
 *  - dataBar:    a proportional in-cell bar (width %) for a numeric value.
 *  - iconSet:    a thresholded icon for a numeric value.
 * A blank/non-numeric cell never matches a comparison / colorScale / dataBar /
 * iconSet rule (Excel parity); formula rules decide via their own predicate.
 * Returns an empty `{ style: {} }` when nothing applies.
 */
export function resolveCellConditionalDisplay(
  rules: readonly ConditionalFormatRule[] | undefined,
  value: unknown,
  row: number,
  col: number,
  opts: ConditionalDisplayOptions = {},
): ConditionalDisplay {
  const resolved: ConditionalDisplay = { style: {} };
  if (!rules || rules.length === 0) return resolved;
  for (const rule of rules) {
    if (!ruleContains(rule, row, col)) continue;
    let effect: ConditionalDisplay | null = null;
    switch (rule.kind) {
      case "formula": {
        const evalFormula = opts.evalFormula;
        if (!evalFormula) continue;
        const result = evalFormula(
          rule.formula,
          row,
          col,
          rule.range.r1,
          rule.range.c1,
        );
        if (!isFormulaTruthy(result)) continue;
        const style: ConditionalStyle = {};
        applyDifferentialCss(style, rule);
        effect = { style };
        break;
      }
      case "colorScale": {
        const n = numericCellValue(value);
        if (n === null) continue;
        const extent = extentFor(rule, opts);
        if (!extent) continue;
        const cellRule = withFormulaStopsResolved(rule, row, col, opts);
        effect = {
          style: {
            backgroundColor: interpolateColorScale(
              cellRule,
              n,
              extent,
              [cellRule.min, cellRule.mid, cellRule.max].some((stop) => stop?.type === "percentile") ||
                (cellRule.mid !== undefined && !cellRule.mid.type)
                ? statsFor(rule, opts)?.sortedNumbers
                : undefined,
            ),
          },
        };
        break;
      }
      case "dataBar": {
        const n = numericCellValue(value);
        if (n === null) continue;
        const extent = extentFor(rule, opts);
        if (!extent) continue;
        const barRule = withFormulaStopsResolved(rule, row, col, opts);
        const axis = dataBarExtent(
          barRule,
          extent,
          [barRule.min, barRule.max].some((stop) => stop?.type === "percentile")
            ? statsFor(rule, opts)?.sortedNumbers
            : undefined,
        );
        // A bar with no x14 fill setting is an Excel 2007 data bar (the
        // export writes it without the x14 twin): no negative axis, every
        // bar from the left edge, (value - min) / (max - min) between the
        // min and max lengths, a colour-to-white gradient and square ends
        // (XLS-A043, measured in Excel).
        const classic = rule.gradient === undefined;
        const geometry = dataBarGeometry(
          Math.max(axis.min, Math.min(axis.max, n)),
          axis,
          rule.minLength ?? 10,
          rule.maxLength ?? 90,
          classic ? "none" : rule.axisPosition,
        );
        effect = {
          style: {},
          dataBar: {
            ...geometry,
            ...(classic ? { classic: true } : {}),
            color: !classic && geometry.direction === "negative" && rule.negativeColor ? rule.negativeColor : rule.color,
            showValue: rule.showValue !== false,
            ...(rule.gradient || classic ? { gradient: true } : {}),
            ...(rule.borderColor ? { borderColor: rule.borderColor } : {}),
            ...(rule.axisColor && geometry.axisPercent > 0 && geometry.axisPercent < 100
              ? { axisColor: rule.axisColor }
              : {}),
          },
        };
        break;
      }
      case "iconSet": {
        // A preserve-only custom mapping the model could not read keeps its
        // per-slot icons in the package; painting the base family's icons
        // would show the wrong pictures, so it is listed but paints nothing.
        if (rule.nativeBlocker === "conditional-format-x14-custom-icons-preserve-only") continue;
        const n = numericCellValue(value);
        if (n === null) continue;
        const extent = extentFor(rule, opts);
        if (!extent) continue;
        const iconRule = withIconFormulaStopsResolved(rule, row, col, opts);
        if (!iconRule) continue;
        const bucket = iconBucketIndex(
          iconRule,
          n,
          extent,
          iconRule.stops?.some((stop) => stop.type === "percentile")
            ? statsFor(rule, opts)?.sortedNumbers
            : undefined,
        );
        const custom = rule.customIcons?.[bucket];
        if (custom) {
          // Excel's custom icon style: this slot shows another family's
          // icon, or none at all.
          if (custom.iconSet === "no-icons") {
            effect = { style: {} };
            break;
          }
          const icon: ConditionalIcon = {
            iconSet: custom.iconSet,
            index: custom.iconId,
            count: iconSetCount(custom.iconSet),
          };
          if (rule.showIconOnly) icon.iconOnly = true;
          effect = { style: {}, icon };
          break;
        }
        const icon: ConditionalIcon = {
          iconSet: rule.iconSet,
          index: bucket,
          count: iconSetCount(rule.iconSet),
        };
        if (rule.showIconOnly) icon.iconOnly = true;
        effect = { style: {}, icon };
        break;
      }
      case "textContains": {
        if (rule.formula !== undefined) {
          
          if (!opts.evalFormula) continue;
          if (!isFormulaTruthy(opts.evalFormula(rule.formula, row, col, rule.range.r1, rule.range.c1))) continue;
        } else if (!evaluateTextContains(rule, value)) continue;
        const style: ConditionalStyle = {};
        applyDifferentialCss(style, rule);
        effect = { style };
        break;
      }
      case "dateOccurring": {
        if (!evaluateDateOccurring(rule, value, opts.now, opts.date1904)) continue;
        const style: ConditionalStyle = {};
        applyDifferentialCss(style, rule);
        effect = { style };
        break;
      }
      case "duplicateValues": {
        const stats = statsFor(rule, opts);
        if (!stats || !evaluateDuplicateValues(rule, value, stats)) continue;
        const style: ConditionalStyle = {};
        applyDifferentialCss(style, rule);
        effect = { style };
        break;
      }
      case "topBottom": {
        const stats = statsFor(rule, opts);
        if (!stats || !isInTopBottom(rule, value, stats)) continue;
        const style: ConditionalStyle = {};
        applyDifferentialCss(style, rule);
        effect = { style };
        break;
      }
      case "aboveBelowAverage": {
        const stats = statsFor(rule, opts);
        if (!stats || !isAboveBelowAverage(rule, value, stats)) continue;
        const style: ConditionalStyle = {};
        applyDifferentialCss(style, rule);
        effect = { style };
        break;
      }
      case "comparison":
      default: {
        // kind === undefined falls here too (back-compat comparison).
        if (!evaluateRule(rule, value)) continue;
        const style: ConditionalStyle = {};
        applyDifferentialCss(style, rule);
        effect = { style };
        break;
      }
    }
    // Array order is native priority. A lower-priority rule may fill a visual
    // channel the higher rule did not set, but it never overwrites a channel
    // already owned by the higher rule. Stop If True halts only after a match.
    if (!effect) continue;
    if (rule.bold) effect.style.fontWeight = 700;
    mergeConditionalStyle(resolved.style, effect.style);
    if (effect.dataBar && !resolved.dataBar) resolved.dataBar = effect.dataBar;
    if (effect.icon && !resolved.icon) resolved.icon = effect.icon;
    if (rule.stopIfTrue) break;
  }
  return resolved;
}

/**
 * Build a normalized rule for a selection rectangle from raw dialog params.
 * Used by the rule-builder UI to turn the user's choices into a persistable
 * rule. The params are a partial rule body (any kind); the range is attached and
 * the whole thing is normalized. A comparison rule omits `kind` (back-compat);
 * the default — params with no `kind` — builds a comparison rule exactly as
 * before. Returns null when the params do not yield a valid rule (bad value, no
 * valid color, …), so the caller surfaces a validation error.
 */
export function buildRule(
  range: ConditionalRange,
  params: Record<string, unknown>,
): ConditionalFormatRule | null {
  return normalizeRule({ ...params, range });
}

/**
 * A short, kind-aware summary of a rule for the rules-list chip UI: a label and
 * the swatch colors to preview it with. Comparison rules read "{op}{value}" and
 * preview with their own colors (unchanged from before the union); the other
 * kinds read a friendly label and preview with a representative color (the bar /
 * first scale stop / a neutral chip for icons). Returned colors are validated
 * #rrggbb (or undefined), safe to drop into an inline style.
 */
export function describeRule(rule: ConditionalFormatRule): {
  label: string;
  backgroundColor?: string;
  textColor?: string;
} {
  switch (rule.kind) {
    case "formula":
      return {
        label: `f(x) ${rule.formula}`,
        backgroundColor: rule.backgroundColor,
        textColor: rule.textColor,
      };
    case "colorScale":
      return { label: "color scale", backgroundColor: rule.min.color };
    case "dataBar":
      return { label: "data bar", backgroundColor: rule.color };
    case "iconSet":
      // A friendly label derived from the family id (e.g. "4-traffic-lights" →
      // "4 traffic lights") plus a reverse/icon-only suffix so the chip is
      // self-describing without an i18n round-trip.
      return {
        label:
          rule.iconSet.replace(/-/g, " ") +
          (rule.reverse ? " ↺" : "") +
          (rule.showIconOnly ? " •" : ""),
      };
    case "textContains": {
      const verb =
        rule.operator === "notContains"
          ? "not ⊇"
          : rule.operator === "beginsWith"
            ? "starts"
            : rule.operator === "endsWith"
              ? "ends"
              : "⊇";
      return {
        label: `${verb} "${rule.text}"`,
        backgroundColor: rule.backgroundColor,
        textColor: rule.textColor,
      };
    }
    case "duplicateValues":
      return {
        label: rule.unique ? "unique" : "duplicate",
        backgroundColor: rule.backgroundColor,
        textColor: rule.textColor,
      };
    case "topBottom":
      return {
        label: `${rule.bottom ? "bottom" : "top"} ${rule.rank}${rule.percent ? "%" : ""}`,
        backgroundColor: rule.backgroundColor,
        textColor: rule.textColor,
      };
    case "aboveBelowAverage":
      return {
        label: `${rule.below ? "below" : "above"} avg${rule.stdDevs ? ` ${rule.stdDevs}σ` : ""}`,
        backgroundColor: rule.backgroundColor,
        textColor: rule.textColor,
      };
    case "dateOccurring":
      return {
        label: `date: ${rule.period}`,
        backgroundColor: rule.backgroundColor,
        textColor: rule.textColor,
      };
    case "comparison":
    default:
      return {
        label: `${rule.operator}${rule.value}`,
        backgroundColor: rule.backgroundColor,
        textColor: rule.textColor,
      };
  }
}

/** `area` without the cells of `cut`: at most four rectangles (the rows
 * above, the columns left and right of it, the rows below). */
function subtractConditionalRange(
  area: ConditionalRange,
  cut: ConditionalRange,
): ConditionalRange[] {
  if (!rangesOverlap(area, cut)) return [area];
  const bandR1 = Math.max(area.r1, cut.r1);
  const bandR2 = Math.min(area.r2, cut.r2);
  const pieces: ConditionalRange[] = [];
  if (area.r1 < cut.r1) {
    pieces.push({ r1: area.r1, c1: area.c1, r2: cut.r1 - 1, c2: area.c2 });
  }
  if (area.c1 < cut.c1) {
    pieces.push({ r1: bandR1, c1: area.c1, r2: bandR2, c2: cut.c1 - 1 });
  }
  if (area.c2 > cut.c2) {
    pieces.push({ r1: bandR1, c1: cut.c2 + 1, r2: bandR2, c2: area.c2 });
  }
  if (area.r2 > cut.r2) {
    pieces.push({ r1: cut.r2 + 1, c1: area.c1, r2: area.r2, c2: area.c2 });
  }
  return pieces;
}

/** A formula stop (no leading "=") moved by the same delta as a rule anchor. */
function shiftStopFormula<S extends { type?: string; formula?: string }>(
  stop: S | undefined,
  rowDelta: number,
  colDelta: number,
): S | undefined {
  if (!stop || stop.type !== "formula" || !stop.formula) return stop;
  const shifted = shiftFormula(`=${stop.formula}`, rowDelta, colDelta);
  return { ...stop, formula: shifted.replace(/^=/, "") };
}

/**
 * The same rule applied to `areas` instead of its own. Relative references in
 * a formula rule, and in formula thresholds, are written against the rule's
 * top-left cell; when that anchor moves they are shifted by the same delta, so
 * every cell that keeps the rule still reads the cells it read before.
 */
function ruleOnAreas(
  rule: ConditionalFormatRule,
  areas: readonly ConditionalRange[],
  keepIdentity: boolean,
): ConditionalFormatRule | null {
  const anchor = areas[0];
  const rowDelta = anchor.r1 - rule.range.r1;
  const colDelta = anchor.c1 - rule.range.c1;
  const { areas: _areas, ...rest } = rule;
  void _areas;
  const candidate: Record<string, unknown> = { ...rest, range: anchor };
  if (areas.length > 1) candidate.areas = areas;
  if (!keepIdentity) {
    delete candidate.id;
    for (const key of NATIVE_RULE_IDENTITY) delete candidate[key];
  }
  if (rowDelta !== 0 || colDelta !== 0) {
    const formula = conditionalRuleFormula(rule);
    if (formula !== undefined) {
      candidate.formula = shiftFormula(formula, rowDelta, colDelta);
    } else if (rule.kind === "colorScale" || rule.kind === "dataBar") {
      for (const key of ["min", "mid", "max"] as const) {
        const stop = (rule as unknown as Record<string, FormulaStopLike | undefined>)[key];
        if (stop) candidate[key] = shiftStopFormula(stop, rowDelta, colDelta);
      }
    } else if (rule.kind === "iconSet" && rule.stops) {
      candidate.stops = rule.stops.map(
        (stop) => shiftStopFormula(stop, rowDelta, colDelta) ?? stop,
      );
    }
  }
  return normalizeRule(candidate);
}

/**
 * Take the given range out of every rule that covers part of it (Excel's
 * "Clear Rules from Selected Cells", and the rule half of Clear Formats).
 * Excel removes conditional formatting from the selected cells and from
 * nothing else: a rule wholly inside the selection goes, and a rule reaching
 * past it keeps its other cells as the same rule (one priority, one Applies to
 * union, so a duplicate or top-10 rule still ranks over all of them). The
 * remainder becomes the rule's areas; relative formulas stay anchored to what
 * they read. Rules clear of the range are returned as the same objects in the
 * same order. Returns a fresh array.
 */
export function clearRulesInRange(
  rules: readonly ConditionalFormatRule[],
  range: ConditionalRange,
): ConditionalFormatRule[] {
  const cut = normalizeRange(range);
  if (!cut) return [...rules];
  const out: ConditionalFormatRule[] = [];
  for (const rule of rules) {
    if (!ruleOverlaps(rule, cut)) {
      out.push(rule);
      continue;
    }
    const pieces = ruleAreas(rule).flatMap((area) =>
      subtractConditionalRange(area, cut),
    );
    // Past the per-rule area bound the remainder carries on as further rules
    // with the same format, each its own union; the first keeps the identity.
    for (let start = 0; start < pieces.length; start += MAX_CONDITIONAL_AREAS) {
      const kept = ruleOnAreas(
        rule,
        pieces.slice(start, start + MAX_CONDITIONAL_AREAS),
        start === 0,
      );
      if (kept) out.push(kept);
    }
  }
  return out;
}

/**
 * Plan the conditional-format half of Home > Clear Formats / Clear All.
 * Conditional formats are formats in Excel, but they live on the sheet rather
 * than inside each cell's style bag. Returning an explicit `changed` bit lets
 * the component combine this result with its cell-facet clear and commit both
 * through one worksheet replacement (one save / Undo entry), including the
 * important case where the selected cells have no static style to clear.
 */
export function clearRulesForFormatClear(
  rules: readonly ConditionalFormatRule[],
  range: ConditionalRange,
): { rules: ConditionalFormatRule[]; changed: boolean } {
  const next = clearRulesInRange(rules, range);
  return {
    rules: next,
    changed:
      next.length !== rules.length ||
      next.some((rule, index) => rule !== rules[index]),
  };
}

/**
 * Rectangles that share a whole edge merged into one, repeatedly; the result
 * is ordered top to bottom, left to right. The cells covered never change.
 */
export function coalesceRanges<R extends ConditionalRange>(ranges: readonly R[]): ConditionalRange[] {
  const out: ConditionalRange[] = ranges.map(({ r1, c1, r2, c2 }) => ({ r1, c1, r2, c2 }));
  let merged = true;
  while (merged) {
    merged = false;
    for (let i = 0; i < out.length && !merged; i += 1) {
      for (let j = i + 1; j < out.length && !merged; j += 1) {
        const a = out[i];
        const b = out[j];
        const stacked = a.c1 === b.c1 && a.c2 === b.c2 && (a.r2 + 1 === b.r1 || b.r2 + 1 === a.r1);
        const sideBySide = a.r1 === b.r1 && a.r2 === b.r2 && (a.c2 + 1 === b.c1 || b.c2 + 1 === a.c1);
        if (!stacked && !sideBySide) continue;
        out[i] = { r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1), r2: Math.max(a.r2, b.r2), c2: Math.max(a.c2, b.c2) };
        out.splice(j, 1);
        merged = true;
      }
    }
  }
  return out.sort((a, b) => a.r1 - b.r1 || a.c1 - b.c1);
}

/**
 * The copies of `source` a fill lays over `filled` (the source and the cells
 * it fills, one rectangle reaching past it on one axis), as offsets of whole
 * source lengths along that axis -- the repeat Fill Down, Fill Right and the
 * fill handle make. Null when `filled` is not such a rectangle.
 */
export function fillTileOffsets(
  source: ConditionalRange,
  filled: ConditionalRange,
): Array<{ rowDelta: number; colDelta: number }> | null {
  const vertical = filled.c1 === source.c1 && filled.c2 === source.c2;
  const horizontal = filled.r1 === source.r1 && filled.r2 === source.r2;
  if (
    (!vertical && !horizontal) ||
    filled.r1 > source.r1 ||
    filled.c1 > source.c1 ||
    filled.r2 < source.r2 ||
    filled.c2 < source.c2
  ) {
    return null;
  }
  const length = vertical ? source.r2 - source.r1 + 1 : source.c2 - source.c1 + 1;
  const before = vertical ? source.r1 - filled.r1 : source.c1 - filled.c1;
  const after = vertical ? filled.r2 - source.r2 : filled.c2 - source.c2;
  const offsets: Array<{ rowDelta: number; colDelta: number }> = [];
  for (let k = -Math.ceil(before / length); k <= Math.ceil(after / length); k += 1) {
    offsets.push(vertical ? { rowDelta: k * length, colDelta: 0 } : { rowDelta: 0, colDelta: k * length });
  }
  return offsets;
}

/**
 * Plan the conditional formats of a paste that repeats the copy over a
 * selection a whole number of copies wide and high (Excel fills the
 * selection with the copy). Each copied rule lands once, on every copy at
 * once: its areas repeat tile by tile, touching rectangles merged, written
 * against the first copy's corner as a single paste would write them. The
 * selection's own rules lose its cells (or are kept beside, with `merge`).
 * `null` on a provenance, bounds or cap failure, as planRulesForFormatPaste.
 */






function pastedRuleId(): string {
  const uuid = globalThis.crypto?.randomUUID?.();
  return uuid
    ? `cf-${uuid}`
    : `cf-${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
}

/** What makes a rule the workbook's own rule, beyond its editor id. */
const NATIVE_RULE_IDENTITY = [
  "nativePriority",
  "nativeSequence",
  "nativeBlocker",
  "nativeX14Id",
  "nativeDataBarId",
] as const;

/**
 * A rule copied as a NEW rule (a paste, Format Painter, Duplicate Rule), as a
 * record for the caller to place and normalize: a fresh id and none of the
 * source's native identity (its priority and sequence in the workbook, its
 * x14 and data bar ids, its import blocker). Two rules never share them: the
 * manager edits a rule by its id and the writer keys a rule by its priority,
 * so a pasted copy that kept them was edited, saved and reopened as its
 * source, and editing the copy moved the source's rule onto the copy's cells
 * (live audit S-142).
 */
export function newConditionalRuleCopy(
  rule: ConditionalFormatRule,
  id: string = pastedRuleId(),
): Record<string, unknown> {
  const copy: Record<string, unknown> = { ...rule, id };
  for (const key of NATIVE_RULE_IDENTITY) delete copy[key];
  return copy;
}

export function planRulesForTiledFormatPaste(
  destinationRules: readonly ConditionalFormatRule[],
  copiedRules: readonly ConditionalFormatRule[],
  sourceRange: ConditionalRange,
  selection: ConditionalRange,
  bounds?: { maxRow?: number; maxCol?: number },
  options: { merge?: boolean } = {},
): { rules: ConditionalFormatRule[]; changed: boolean } | null {
  const rows = sourceRange.r2 - sourceRange.r1 + 1;
  const cols = sourceRange.c2 - sourceRange.c1 + 1;
  const tileRows = (selection.r2 - selection.r1 + 1) / rows;
  const tileCols = (selection.c2 - selection.c1 + 1) / cols;
  if (!Number.isInteger(tileRows) || !Number.isInteger(tileCols) || tileRows < 1 || tileCols < 1) return null;
  const base = options.merge ? [...destinationRules] : clearRulesInRange(destinationRules, selection);
  const first = planRulesForFormatPaste(
    base,
    copiedRules,
    sourceRange,
    { r1: selection.r1, c1: selection.c1, r2: selection.r1 + rows - 1, c2: selection.c1 + cols - 1 },
    bounds,
    { merge: true },
  );
  if (!first) return null;
  const translated = first.rules.slice(0, first.rules.length - base.length);
  const repeated: ConditionalFormatRule[] = [];
  for (const rule of translated) {
    const areas = ruleAreas(rule);
    const copies: ConditionalRange[] = [];
    for (let tileRow = 0; tileRow < tileRows; tileRow += 1) {
      for (let tileCol = 0; tileCol < tileCols; tileCol += 1) {
        for (const area of areas) {
          copies.push({
            r1: area.r1 + tileRow * rows,
            c1: area.c1 + tileCol * cols,
            r2: area.r2 + tileRow * rows,
            c2: area.c2 + tileCol * cols,
          });
        }
      }
    }
    const merged = coalesceRanges(copies);
    for (let start = 0; start < merged.length; start += MAX_CONDITIONAL_AREAS) {
      const piece = ruleOnAreas(rule, merged.slice(start, start + MAX_CONDITIONAL_AREAS), start === 0);
      if (!piece) return null;
      repeated.push(piece);
    }
  }
  const rules = [...repeated, ...base];
  if (rules.length > MAX_CONDITIONAL_FORMATS) return null;
  return { rules, changed: !rulesSayTheSame(rules, destinationRules) };
}

/** Whether two rule lists differ only in their ids: a pasted rule gets a new
 *  one, so pasting the same rules again changed nothing yet read as a change,
 *  an empty save and an Undo step (a paste repeated at the same place). */
function rulesSayTheSame(
  left: readonly ConditionalFormatRule[],
  right: readonly ConditionalFormatRule[],
): boolean {
  const spelled = (rules: readonly ConditionalFormatRule[]) =>
    JSON.stringify(rules.map(({ id: _id, ...rule }) => rule));
  return left.length === right.length && spelled(left) === spelled(right);
}

/**
 * Plan the conditional formats of a fill (Home > Fill > Down / Right / Up /
 * Left and the fill handle; live audit S-034). Excel 16 measured: Fill Down
 * from R164, whose rule is =$S164>0, leaves ONE rule applying to R164:R166 --
 * the cells filled take the source cells' rules, each read from its own row
 * as a copy would be, and the rule they came from simply reaches further.
 * So a rule's areas inside the source repeat over the filled cells, joined to
 * its areas outside them (its identity and rank kept; relative formulas stay
 * written against the first area, see ruleOnAreas). A rule that reaches the
 * filled cells but not the source loses them, as a paste replaces them.
 * `null` when `filled` is not a fill of `source`, or a rule outgrows the
 * bounds or the per-sheet cap.
 */
export function planRulesForFill(
  rules: readonly ConditionalFormatRule[],
  source: ConditionalRange,
  filled: ConditionalRange,
  bounds?: { maxRow?: number; maxCol?: number },
): { rules: ConditionalFormatRule[]; changed: boolean } | null {
  const offsets = fillTileOffsets(source, filled);
  if (!offsets) return null;
  if (
    (bounds?.maxRow !== undefined && filled.r2 > bounds.maxRow) ||
    (bounds?.maxCol !== undefined && filled.c2 > bounds.maxCol)
  ) {
    return null;
  }
  const clip = (area: ConditionalRange, to: ConditionalRange): ConditionalRange | null => {
    const clipped = {
      r1: Math.max(area.r1, to.r1),
      c1: Math.max(area.c1, to.c1),
      r2: Math.min(area.r2, to.r2),
      c2: Math.min(area.c2, to.c2),
    };
    return clipped.r1 <= clipped.r2 && clipped.c1 <= clipped.c2 ? clipped : null;
  };
  const out: ConditionalFormatRule[] = [];
  for (const rule of rules) {
    if (!ruleOverlaps(rule, filled)) {
      out.push(rule);
      continue;
    }
    const areas = ruleAreas(rule);
    const kept = areas.flatMap((area) => subtractConditionalRange(area, filled));
    const inSource = areas
      .map((area) => clip(area, source))
      .filter((area): area is ConditionalRange => area !== null);
    const copies = inSource.flatMap((area) =>
      offsets
        .map(({ rowDelta, colDelta }) =>
          clip({ r1: area.r1 + rowDelta, c1: area.c1 + colDelta, r2: area.r2 + rowDelta, c2: area.c2 + colDelta }, filled),
        )
        .filter((area): area is ConditionalRange => area !== null),
    );
    const next = coalesceRanges([...kept, ...copies]);
    for (let start = 0; start < next.length; start += MAX_CONDITIONAL_AREAS) {
      const piece = ruleOnAreas(rule, next.slice(start, start + MAX_CONDITIONAL_AREAS), start === 0);
      if (piece) out.push(piece);
    }
  }
  if (out.length > MAX_CONDITIONAL_FORMATS) return null;
  return {
    rules: out,
    changed: JSON.stringify(out) !== JSON.stringify(rules),
  };
}

/**
 * The rule with every area moved by the given offset and its formulas left as
 * written: what a moved cell carries before its references are re-pointed.
 * `null` when an area leaves `bounds`.
 */
export function ruleMovedBy(
  rule: ConditionalFormatRule,
  rowDelta: number,
  colDelta: number,
  bounds?: { maxRow?: number; maxCol?: number },
): ConditionalFormatRule | null {
  const areas = ruleAreas(rule).map((area) => ({
    r1: area.r1 + rowDelta,
    c1: area.c1 + colDelta,
    r2: area.r2 + rowDelta,
    c2: area.c2 + colDelta,
  }));
  const { areas: _areas, ...rest } = rule;
  void _areas;
  const moved = normalizeRule(
    { ...rest, range: areas[0], ...(areas.length > 1 ? { areas } : {}) },
    bounds,
  );
  const movedAreas = moved ? ruleAreas(moved) : [];
  return moved && movedAreas.length === areas.length && movedAreas.every((area, index) => rangesEqual(area, areas[index]))
    ? moved
    : null;
}

/** Every formula a rule holds, as ruleOnAreas shifts them. */
function ruleFormulaTexts(rule: ConditionalFormatRule): string[] {
  const formula = conditionalRuleFormula(rule);
  if (formula !== undefined) return [formula];
  const stops: (FormulaStopLike | undefined)[] = [];
  if (rule.kind === "colorScale" || rule.kind === "dataBar") {
    const record = rule as unknown as Record<string, FormulaStopLike | undefined>;
    stops.push(record.min, record.mid, record.max);
  } else if (rule.kind === "iconSet" && rule.stops) {
    stops.push(...rule.stops);
  }
  return stops.flatMap((stop) => (stop?.type === "formula" && stop.formula ? [stop.formula] : []));
}

const RELATIVE_REF_RE = /(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})/g;
const MAX_MOVED_RULE_PIECES = 4_096;












function movedShareFollowed(
  moved: ConditionalFormatRule,
  follow: (rule: ConditionalFormatRule) => ConditionalFormatRule,
  cut: ConditionalRange,
  rowDelta: number,
  colDelta: number,
): ConditionalFormatRule[] {
  const whole = [follow(moved)];
  const anchor = moved.range;
  const blocks = [cut, { r1: cut.r1 + rowDelta, c1: cut.c1 + colDelta, r2: cut.r2 + rowDelta, c2: cut.c2 + colDelta }];
  // Where a relative reference starts or stops naming a moved (or overwritten)
  // cell as the reading cell walks the share.
  const rowCuts = new Set<number>();
  const colCuts = new Set<number>();
  for (const text of ruleFormulaTexts(moved)) {
    for (const match of text.matchAll(RELATIVE_REF_RE)) {
      const row = Number(match[4]) - 1;
      const col = lettersToColumn(match[2].toUpperCase());
      for (const block of blocks) {
        if (!match[3]) {
          rowCuts.add(block.r1 - row + anchor.r1);
          rowCuts.add(block.r2 - row + anchor.r1 + 1);
        }
        if (!match[1]) {
          colCuts.add(block.c1 - col + anchor.c1);
          colCuts.add(block.c2 - col + anchor.c1 + 1);
        }
      }
    }
  }
  const segments = (from: number, to: number, cuts: Set<number>): Array<[number, number]> => {
    const points = [...cuts].filter((point) => point > from && point <= to).sort((a, b) => a - b);
    const out: Array<[number, number]> = [];
    let start = from;
    for (const point of points) {
      out.push([start, point - 1]);
      start = point;
    }
    out.push([start, to]);
    return out;
  };
  const pieces: ConditionalRange[] = [];
  for (const area of ruleAreas(moved)) {
    for (const [r1, r2] of segments(area.r1, area.r2, rowCuts)) {
      for (const [c1, c2] of segments(area.c1, area.c2, colCuts)) pieces.push({ r1, c1, r2, c2 });
    }
    if (pieces.length > MAX_MOVED_RULE_PIECES) return whole;
  }
  if (pieces.length <= ruleAreas(moved).length) return whole;
  const body = (rule: ConditionalFormatRule) => {
    const record: Record<string, unknown> = { ...rule };
    delete record.id;
    for (const key of NATIVE_RULE_IDENTITY) delete record[key];
    return JSON.stringify(record);
  };
  const groups: Array<{ read: ConditionalFormatRule; areas: ConditionalRange[] }> = [];
  for (const piece of pieces) {
    const own = ruleOnAreas(moved, [piece], true);
    if (!own) return whole;
    const read = follow(own);
    const spelled = body(read);
    const group = groups.find((candidate) => {
      const there = ruleOnAreas(candidate.read, [piece], true);
      return there !== null && body(there) === spelled;
    });
    if (group) group.areas.push(piece);
    else groups.push({ read, areas: [piece] });
  }
  if (groups.length === 1) return whole;
  const out: ConditionalFormatRule[] = [];
  for (const [index, group] of groups.entries()) {
    const areas = coalesceRanges(group.areas);
    for (let start = 0; start < areas.length; start += MAX_CONDITIONAL_AREAS) {
      const rule = ruleOnAreas(group.read, areas.slice(start, start + MAX_CONDITIONAL_AREAS), index === 0 && start === 0);
      if (!rule) return whole;
      out.push(rule);
    }
  }
  return out;
}





























export function planRulesForCellMove(
  rules: readonly ConditionalFormatRule[],
  from: ConditionalRange,
  rowDelta: number,
  colDelta: number,
  follow: (rule: ConditionalFormatRule) => ConditionalFormatRule,
  bounds?: { maxRow?: number; maxCol?: number },
  overflow?: "demote",
): { rules: ConditionalFormatRule[]; changed: boolean; demoted?: ConditionalFormatRule[] } | null {
  const cut = normalizeRange(from);
  if (!cut) return null;
  const out: ConditionalFormatRule[] = [];
  // Pieces of a rule the move split, ranked after every rule (Excel).
  const trailing: ConditionalFormatRule[] = [];
  for (const rule of rules) {
    if (!ruleOverlaps(rule, cut)) {
      out.push(follow(rule));
      continue;
    }
    const areas = ruleAreas(rule);
    const keptAreas = areas.flatMap((area) => subtractConditionalRange(area, cut));
    const movedSourceAreas = areas
      .filter((area) => rangesOverlap(area, cut))
      .map((area) => ({
        r1: Math.max(area.r1, cut.r1),
        c1: Math.max(area.c1, cut.c1),
        r2: Math.min(area.r2, cut.r2),
        c2: Math.min(area.c2, cut.c2),
      }));
    // The moved cells' share of the rule, written against their own anchor
    // and then carried over unchanged: only `follow` re-points references.
    const movedAtSource = ruleOnAreas(rule, movedSourceAreas, keptAreas.length === 0);
    const movedCandidate = movedAtSource
      ? ruleMovedBy(movedAtSource, rowDelta, colDelta, bounds)
      : null;
    if (!movedCandidate) return null;
    const movedShare = movedShareFollowed(movedCandidate, follow, cut, rowDelta, colDelta);
    if (keptAreas.length === 0) {
      out.push(...movedShare);
      continue;
    }
    const kept: ConditionalFormatRule[] = [];
    for (let start = 0; start < keptAreas.length; start += MAX_CONDITIONAL_AREAS) {
      const piece = ruleOnAreas(
        rule,
        keptAreas.slice(start, start + MAX_CONDITIONAL_AREAS),
        start === 0,
      );
      if (piece) kept.push(follow(piece));
    }
    // One rule again when a moved share only differs by where it is written
    // (Excel keeps P121:P122 and R120 as one Applies to for =P120>15; A4:A6
    // and D11:D12 for B2>0 over A1:A6 cut to D11, its D13 a rule of its own,
    // backlog/xq829 s2), so a top-10, duplicate or average rule keeps a
    // single population.
    let first = kept[0];
    const apart: ConditionalFormatRule[] = [];
    for (const moved of movedShare) {
      const movedAreas = ruleAreas(moved);
      if (
        first &&
        kept.length === 1 &&
        ruleAreas(first).length + movedAreas.length <= MAX_CONDITIONAL_AREAS &&
        !movedAreas.some((area) => keptAreas.some((keptArea) => rangesOverlap(area, keptArea)))
      ) {
        const joined = ruleOnAreas(first, [...ruleAreas(first), ...movedAreas], true);
        const rewritten = joined ? ruleOnAreas(joined, movedAreas, false) : null;
        const { id: _id, nativePriority: _priority, nativeSequence: _sequence, ...movedBody } = moved;
        void _id;
        void _priority;
        void _sequence;
        if (joined && rewritten && JSON.stringify(rewritten) === JSON.stringify(movedBody)) {
          first = joined;
          continue;
        }
      }
      apart.push(moved);
    }
    if (first && kept.length === 1) out.push(first);
    else out.push(...kept);
    trailing.push(...apart);
  }
  out.push(...trailing);
  const changed = out.length !== rules.length || out.some((rule, index) => rule !== rules[index]);
  if (out.length <= MAX_CONDITIONAL_FORMATS) return { rules: out, changed };
  const demoted = out.slice(MAX_CONDITIONAL_FORMATS);
  if (
    overflow !== "demote" ||
    demoted.length > MAX_DEMOTED_CONDITIONAL_FORMATS ||
    demoted.some((rule) => !trailing.includes(rule) || ruleHasX14Part(rule))
  ) {
    return null;
  }
  return { rules: out.slice(0, MAX_CONDITIONAL_FORMATS), changed: true, demoted };
}

 
export const MAX_DEMOTED_CONDITIONAL_FORMATS = 64;

/** Whether Excel writes a rule with an x14 part (a twin, or the rule itself):
 *  a data bar, an icon set of custom icons, or a rule naming one. */
function ruleHasX14Part(rule: ConditionalFormatRule): boolean {
  const record = rule as unknown as Record<string, unknown>;
  return (
    rule.kind === "dataBar" ||
    (rule.kind === "iconSet" && Array.isArray(record.customIcons) && record.customIcons.length > 0) ||
    record.nativeX14Id !== undefined ||
    record.nativeDataBarId !== undefined
  );
}

/**
 * Detach the conditional-format portion of a copied range for Paste Formats.
 *
 * A rule may cover more cells than the copied rectangle. The clipboard owns
 * only their intersection, so clip the range and rebase formula rules from the
 * original rule anchor to the clipped top-left. Re-normalizing each result both
 * deep-clones nested payloads and keeps malformed persisted rules out of the
 * session clipboard.
 */
export function snapshotRulesForFormatPaste(
  rules: readonly ConditionalFormatRule[],
  sourceRange: ConditionalRange,
): ConditionalFormatRule[] {
  const source = normalizeRange(sourceRange);
  if (!source) return [];
  const snapshot: ConditionalFormatRule[] = [];
  for (const rawRule of rules) {
    const rule = normalizeRule(rawRule);
    if (!rule) continue;
    // A multi-area rule stays ONE rule: its areas share one ranking / average
    // / duplicate population, which splitting them would turn into per-area
    // statistics (live audit B305). Each area is clipped to the copy.
    const clippedAreas: ConditionalRange[] = [];
    for (const area of ruleAreas(rule)) {
      if (!rangesOverlap(area, source)) continue;
      clippedAreas.push({
        r1: Math.max(area.r1, source.r1),
        c1: Math.max(area.c1, source.c1),
        r2: Math.min(area.r2, source.r2),
        c2: Math.min(area.c2, source.c2),
      });
    }
    if (clippedAreas.length === 0) continue;
    // `areas` lists every rectangle, the first being the rule's own range.
    const clipped = clippedAreas[0];
    const { areas: _areas, ...detachedRule } = rule;
    void _areas;
    const candidate = {
      ...detachedRule,
      range: clipped,
      ...(clippedAreas.length > 1 ? { areas: clippedAreas } : {}),
    } as Record<string, unknown>;
    const formula = conditionalRuleFormula(rule);
    if (formula !== undefined) {
      candidate.formula = shiftFormula(
        formula,
        clipped.r1 - rule.range.r1,
        clipped.c1 - rule.range.c1,
      );
    }
    const detached = normalizeRule(candidate);
    if (detached) snapshot.push(detached);
  }
  return snapshot;
}

/**
 * Plan the worksheet-level conditional-format half of Format Painter.
 *
 * Excel creates NEW destination rules rather than replacing rules already on
 * the target. The copied snapshot is clipped to `sourceRange` at arm time; a
 * one-cell target has already been expanded by the component, while a larger
 * target is an exact positive-integer tiling of that source rectangle. For
 * every source rule, create one translated rule per row-major tile, preserving
 * source priority before putting the new rules at the top of the worksheet's
 * global rule list. Existing rules keep their exact relative order.
 *
 * An empty source snapshot is authoritative only for "nothing to add" — unlike
 * Paste Formats, Format Painter does NOT clear destination conditional rules.
 * `null` is the all-or-nothing fail-closed result for malformed provenance,
 * range/bounds mismatch, or exceeding the per-sheet rule cap.
 */
export function planRulesForFormatPainter(
  destinationRules: readonly ConditionalFormatRule[],
  copiedRules: readonly ConditionalFormatRule[],
  sourceRange: ConditionalRange,
  destinationRange: ConditionalRange,
  bounds?: { maxRow?: number; maxCol?: number },
): { rules: ConditionalFormatRule[]; changed: boolean } | null {
  const source = normalizeRange(sourceRange);
  const destination = normalizeRange(destinationRange, bounds);
  if (!source || !destination) return null;
  const rangeIsExact = (
    normalized: ConditionalRange,
    raw: ConditionalRange,
  ) =>
    normalized.r1 === raw.r1 &&
    normalized.c1 === raw.c1 &&
    normalized.r2 === raw.r2 &&
    normalized.c2 === raw.c2;
  if (
    !rangeIsExact(source, sourceRange) ||
    !rangeIsExact(destination, destinationRange)
  ) {
    return null;
  }

  const sourceRows = source.r2 - source.r1 + 1;
  const sourceColumns = source.c2 - source.c1 + 1;
  const destinationRows = destination.r2 - destination.r1 + 1;
  const destinationColumns = destination.c2 - destination.c1 + 1;
  if (
    destinationRows % sourceRows !== 0 ||
    destinationColumns % sourceColumns !== 0
  ) {
    return null;
  }
  const tileRows = destinationRows / sourceRows;
  const tileColumns = destinationColumns / sourceColumns;
  if (
    !Number.isSafeInteger(tileRows) ||
    !Number.isSafeInteger(tileColumns) ||
    tileRows < 1 ||
    tileColumns < 1 ||
    destinationRules.length > MAX_CONDITIONAL_FORMATS
  ) {
    return null;
  }

  // The caller supplies its live normalized rule list. Revalidate exact ranges
  // against the current destination bounds so a stale/clamped worksheet rule
  // cannot be silently rewritten as a side effect of painting static styles.
  for (const rawRule of destinationRules) {
    const normalized = normalizeRule(rawRule, bounds);
    if (!normalized || !rangeIsExact(normalized.range, rawRule.range)) {
      return null;
    }
  }

  const tileCount = tileRows * tileColumns;
  const availableRuleSlots =
    MAX_CONDITIONAL_FORMATS - destinationRules.length;
  if (
    copiedRules.length > 0 &&
    tileCount > Math.floor(availableRuleSlots / copiedRules.length)
  ) {
    return null;
  }

  const translated: ConditionalFormatRule[] = [];
  for (const rawRule of copiedRules.flatMap(splitRuleAreas)) {
    const rule = normalizeRule(rawRule);
    if (
      !rule ||
      rule.range.r1 < source.r1 ||
      rule.range.c1 < source.c1 ||
      rule.range.r2 > source.r2 ||
      rule.range.c2 > source.c2
    ) {
      return null;
    }
    for (let tileRow = 0; tileRow < tileRows; tileRow += 1) {
      for (let tileColumn = 0; tileColumn < tileColumns; tileColumn += 1) {
        const rowDelta =
          destination.r1 + tileRow * sourceRows - source.r1;
        const columnDelta =
          destination.c1 + tileColumn * sourceColumns - source.c1;
        const nextRange: ConditionalRange = {
          r1: rule.range.r1 + rowDelta,
          c1: rule.range.c1 + columnDelta,
          r2: rule.range.r2 + rowDelta,
          c2: rule.range.c2 + columnDelta,
        };
        // Each tile is a new rule of its own (Excel adds rules; S-142).
        const candidate: Record<string, unknown> = {
          ...newConditionalRuleCopy(rule),
          range: nextRange,
        };
        const formula = conditionalRuleFormula(rule);
        if (formula !== undefined) {
          candidate.formula = shiftFormula(
            formula,
            rowDelta,
            columnDelta,
          );
        }
        const nextRule = normalizeRule(candidate, bounds);
        if (!nextRule || !rangeIsExact(nextRule.range, nextRange)) return null;
        translated.push(nextRule);
      }
    }
  }

  return {
    rules: [...translated, ...destinationRules],
    changed: translated.length > 0,
  };
}

/**
 * The conditional formats of a Copy that leaves rows out (the visible rows
 * of a filtered table), laid out as that Copy pastes: the kept rows close up
 * under the first one and each run of them takes its rules along, as a paste
 * moves them (live audit S-045). A rule across several runs becomes one rule
 * per run.
 */
export function snapshotRulesForKeptRowsFormatPaste(
  rules: readonly ConditionalFormatRule[],
  sourceRange: ConditionalRange,
  runs: ReadonlyArray<{ from: number; to: number; at: number }>,
): ConditionalFormatRule[] | null {
  const out: ConditionalFormatRule[] = [];
  for (const run of runs) {
    const from = { r1: run.from, c1: sourceRange.c1, r2: run.to, c2: sourceRange.c2 };
    const placed = planRulesForFormatPaste(
      [],
      snapshotRulesForFormatPaste(rules, from),
      from,
      { ...from, r1: run.at, r2: run.at + run.to - run.from },
    );
    if (!placed) return null;
    out.push(...placed.rules);
  }
  return out.length > MAX_CONDITIONAL_FORMATS ? null : out;
}

/**
 * Plan the sheet-level conditional-format half of Paste Formats.
 *
 * The copied snapshot is already clipped to `sourceRange`. Every copied rule is
 * translated to the same relative position inside `destinationRange`; formula
 * rules move through the shared formula shifter so relative/mixed/absolute and
 * sheet-qualified references follow the same copy semantics as cell formulas.
 * Existing destination-intersecting rules are replaced, while every unrelated
 * destination rule keeps its relative order. `null` is the fail-closed result
 * for malformed provenance, bounds overflow, or the per-sheet rule cap.
 *
 * `merge` is Paste Special's "All merging conditional formats" (XLS-A063):
 * the destination's own rules stay, intersecting ones included, and the
 * copied rules join them ahead of them.
 */
export function planRulesForFormatPaste(
  destinationRules: readonly ConditionalFormatRule[],
  copiedRules: readonly ConditionalFormatRule[],
  sourceRange: ConditionalRange,
  destinationRange: ConditionalRange,
  bounds?: { maxRow?: number; maxCol?: number },
  options: { merge?: boolean } = {},
): { rules: ConditionalFormatRule[]; changed: boolean } | null {
  const source = normalizeRange(sourceRange);
  const destination = normalizeRange(destinationRange);
  if (!source || !destination) return null;
  if (
    source.r1 !== sourceRange.r1 ||
    source.c1 !== sourceRange.c1 ||
    source.r2 !== sourceRange.r2 ||
    source.c2 !== sourceRange.c2 ||
    destination.r1 !== destinationRange.r1 ||
    destination.c1 !== destinationRange.c1 ||
    destination.r2 !== destinationRange.r2 ||
    destination.c2 !== destinationRange.c2 ||
    source.r2 - source.r1 !== destination.r2 - destination.r1 ||
    source.c2 - source.c1 !== destination.c2 - destination.c1 ||
    (bounds?.maxRow !== undefined && destination.r2 > bounds.maxRow) ||
    (bounds?.maxCol !== undefined && destination.c2 > bounds.maxCol)
  ) {
    return null;
  }

  const rowDelta = destination.r1 - source.r1;
  const colDelta = destination.c1 - source.c1;
  const translated: ConditionalFormatRule[] = [];
  // A copied multi-area rule is pasted as one rule with every area moved by
  // the same offset, so its areas keep sharing one population (B305).
  const translate = (range: ConditionalRange): ConditionalRange | null =>
    range.r1 < source.r1 || range.c1 < source.c1 || range.r2 > source.r2 || range.c2 > source.c2
      ? null
      : { r1: range.r1 + rowDelta, c1: range.c1 + colDelta, r2: range.r2 + rowDelta, c2: range.c2 + colDelta };
  for (const rawRule of copiedRules) {
    const rule = normalizeRule(rawRule);
    if (!rule) return null;
    const nextRange = translate(rule.range);
    if (!nextRange) return null;
    const nextAreas: ConditionalRange[] = [];
    for (const area of rule.areas ?? []) {
      const moved = translate(area);
      if (!moved) return null;
      nextAreas.push(moved);
    }
    const candidate: Record<string, unknown> = {
      ...newConditionalRuleCopy(rule),
      range: nextRange,
      ...(nextAreas.length > 1 ? { areas: nextAreas } : {}),
    };
    if (nextAreas.length <= 1) delete candidate.areas;
    const formula = conditionalRuleFormula(rule);
    if (formula !== undefined) {
      candidate.formula = shiftFormula(formula, rowDelta, colDelta);
    }
    const nextRule = normalizeRule(candidate, bounds);
    if (
      !nextRule ||
      nextRule.range.r1 !== nextRange.r1 ||
      nextRule.range.c1 !== nextRange.c1 ||
      nextRule.range.r2 !== nextRange.r2 ||
      nextRule.range.c2 !== nextRange.c2 ||
      (nextRule.areas?.length ?? 0) !== (nextAreas.length > 1 ? nextAreas.length : 0)
    ) {
      return null;
    }
    translated.push(nextRule);
  }

  const retained = options.merge
    ? [...destinationRules]
    : clearRulesInRange(destinationRules, destination);
  if (translated.length + retained.length > MAX_CONDITIONAL_FORMATS) {
    return null;
  }
  const rules = [...translated, ...retained];
  return {
    rules,
    changed: !rulesSayTheSame(rules, destinationRules),
  };
}







export function formatPasteOverRuleLimit(
  destinationRules: readonly ConditionalFormatRule[],
  copiedRules: readonly ConditionalFormatRule[],
  sourceRange: ConditionalRange,
  destinationRange: ConditionalRange,
  bounds?: { maxRow?: number; maxCol?: number },
): boolean {
  const destination = normalizeRange(destinationRange);
  if (!destination) return false;
  const retained = clearRulesInRange(destinationRules, destination);
  if (copiedRules.length + retained.length <= MAX_CONDITIONAL_FORMATS) return false;
  // The same plan with room for every copied rule: it holds when nothing but
  // the bound refused it.
  return planRulesForFormatPaste([], copiedRules, sourceRange, destinationRange, bounds) !== null;
}










export function planRulesForFormatPasteWithFile(
  destinationRules: readonly ConditionalFormatRule[],
  copiedRules: readonly ConditionalFormatRule[],
  sourceRange: ConditionalRange,
  destinationRange: ConditionalRange,
  sourceRuleCount: number,
  bounds?: { maxRow?: number; maxCol?: number },
  options: { merge?: boolean } = {},
): { plan: { rules: ConditionalFormatRule[]; changed: boolean }; fileCopy: { page: boolean } | null } | null {
  return planRulesWithFile(
    planRulesForFormatPaste,
    destinationRules,
    copiedRules,
    sourceRange,
    destinationRange,
    sourceRuleCount,
    bounds,
    options,
  );
}

type FormatPastePlanner = (
  destinationRules: readonly ConditionalFormatRule[],
  copiedRules: readonly ConditionalFormatRule[],
  sourceRange: ConditionalRange,
  destinationRange: ConditionalRange,
  bounds?: { maxRow?: number; maxCol?: number },
  options?: { merge?: boolean },
) => { rules: ConditionalFormatRule[]; changed: boolean } | null;








export function planRulesWithFile(
  planner: FormatPastePlanner,
  destinationRules: readonly ConditionalFormatRule[],
  copiedRules: readonly ConditionalFormatRule[],
  sourceRange: ConditionalRange,
  destinationRange: ConditionalRange,
  sourceRuleCount: number,
  bounds?: { maxRow?: number; maxCol?: number },
  options: { merge?: boolean } = {},
): { plan: { rules: ConditionalFormatRule[]; changed: boolean }; fileCopy: { page: boolean } | null } | null {
  const plan = planner(destinationRules, copiedRules, sourceRange, destinationRange, bounds, options);
  if (plan) return { plan, fileCopy: sourceRuleCount >= MAX_CONDITIONAL_FORMATS ? { page: true } : null };
  if (options.merge) return null;
  const destination = normalizeRange(destinationRange);
  if (!destination) return null;
  const retained = clearRulesInRange(destinationRules, destination);
  // Only the bound reads the destination's own rules: the same plan on a
  // sheet with none holds when nothing but the bound refused it.
  if (!planner([], copiedRules, sourceRange, destinationRange, bounds)) return null;
  return {
    plan: { rules: retained, changed: !rulesSayTheSame(retained, destinationRules) },
    fileCopy: { page: false },
  };
}

/**
 * Plan conditional formats for Home > Paste > Transpose.
 *
 * The copied rules are already clipped to the source rectangle. Transposition
 * swaps each rule range's source-relative row/column offsets and dimensions;
 * a formula remains relative to that rule's top-left, so it is shifted from the
 * clipped source anchor to the mapped destination anchor. Existing rules that
 * intersect the complete transposed destination are replaced atomically.
 */
export function planRulesForTransposedFormatPaste(
  destinationRules: readonly ConditionalFormatRule[],
  copiedRules: readonly ConditionalFormatRule[],
  sourceRange: ConditionalRange,
  destinationRange: ConditionalRange,
  bounds?: { maxRow?: number; maxCol?: number },
): { rules: ConditionalFormatRule[]; changed: boolean } | null {
  const source = normalizeRange(sourceRange);
  const destination = normalizeRange(destinationRange);
  if (!source || !destination) return null;
  if (
    source.r1 !== sourceRange.r1 ||
    source.c1 !== sourceRange.c1 ||
    source.r2 !== sourceRange.r2 ||
    source.c2 !== sourceRange.c2 ||
    destination.r1 !== destinationRange.r1 ||
    destination.c1 !== destinationRange.c1 ||
    destination.r2 !== destinationRange.r2 ||
    destination.c2 !== destinationRange.c2 ||
    source.r2 - source.r1 !== destination.c2 - destination.c1 ||
    source.c2 - source.c1 !== destination.r2 - destination.r1 ||
    (bounds?.maxRow !== undefined && destination.r2 > bounds.maxRow) ||
    (bounds?.maxCol !== undefined && destination.c2 > bounds.maxCol)
  ) {
    return null;
  }

  const translated: ConditionalFormatRule[] = [];
  for (const rawRule of copiedRules.flatMap(splitRuleAreas)) {
    const rule = normalizeRule(rawRule);
    if (
      !rule ||
      rule.range.r1 < source.r1 ||
      rule.range.c1 < source.c1 ||
      rule.range.r2 > source.r2 ||
      rule.range.c2 > source.c2
    ) {
      return null;
    }
    const nextRange: ConditionalRange = {
      r1: destination.r1 + rule.range.c1 - source.c1,
      c1: destination.c1 + rule.range.r1 - source.r1,
      r2: destination.r1 + rule.range.c2 - source.c1,
      c2: destination.c1 + rule.range.r2 - source.r1,
    };
    const rowDelta = nextRange.r1 - rule.range.r1;
    const colDelta = nextRange.c1 - rule.range.c1;
    const candidate: Record<string, unknown> = { ...newConditionalRuleCopy(rule), range: nextRange };
    const formula = conditionalRuleFormula(rule);
    if (formula !== undefined) {
      candidate.formula = shiftFormula(formula, rowDelta, colDelta);
    }
    const nextRule = normalizeRule(candidate, bounds);
    if (
      !nextRule ||
      nextRule.range.r1 !== nextRange.r1 ||
      nextRule.range.c1 !== nextRange.c1 ||
      nextRule.range.r2 !== nextRange.r2 ||
      nextRule.range.c2 !== nextRange.c2
    ) {
      return null;
    }
    translated.push(nextRule);
  }

  const retained = clearRulesInRange(destinationRules, destination);
  if (translated.length + retained.length > MAX_CONDITIONAL_FORMATS) {
    return null;
  }
  const rules = [...translated, ...retained];
  return {
    rules,
    changed: JSON.stringify(rules) !== JSON.stringify(destinationRules),
  };
}

/** True when any rule's range intersects the given range. */
export function rangeHasRule(
  rules: readonly ConditionalFormatRule[],
  range: ConditionalRange,
): boolean {
  return rules.some((rule) => ruleOverlaps(rule, range));
}

/**
 * Reorder a conditional-format rule by one position (Excel "Manage Rules" Move
 * Up / Move Down). Rule PRIORITY is array order — the FIRST rule that matches a
 * cell supplies its style (resolveCellConditionalStyle / Display iterate front
 * to back), so index 0 is the HIGHEST priority. "up" moves the rule toward index
 * 0 (raises its priority by swapping with its predecessor); "down" moves it
 * toward the end (lowers it). A move at the boundary (first rule up / last rule
 * down) or an out-of-range index is a no-op and returns the SAME array reference
 * so the caller can skip a needless re-render / sheet write. Otherwise returns a
 * NEW array with the two neighbours swapped.
 */
export function moveConditionalRule(
  rules: readonly ConditionalFormatRule[],
  index: number,
  direction: "up" | "down",
): ConditionalFormatRule[] | readonly ConditionalFormatRule[] {
  if (!Number.isInteger(index) || index < 0 || index >= rules.length) {
    return rules;
  }
  const target = direction === "up" ? index - 1 : index + 1;
  if (target < 0 || target >= rules.length) return rules;
  const next = [...rules];
  [next[index], next[target]] = [next[target], next[index]];
  return next;
}

const HYPERLINK_ALLOWED_SCHEMES = new Set(["http:", "https:", "mailto:"]);

/**
 * Strip ASCII control characters and spaces that browsers ignore inside a URL
 * but that an attacker can use to smuggle a scheme past a naive check
 * (e.g. "java\tscript:"). Mirrors the defense in lib/security.
 */
function stripUrlInsignificantChars(value: string): string {
  let out = "";
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    // Drop C0 controls + space (<=0x20) and DEL (0x7f).
    if (code <= 0x20 || code === 0x7f) continue;
    out += value[index];
  }
  return out;
}

/**
 * Sanitize a cell hyperlink to a safe href, or null when it is unsafe/empty.
 * ONLY http(s) and mailto: links are allowed — javascript:, data:, vbscript:,
 * file:, tel:, relative, and protocol-relative ("//evil") URLs are all rejected
 * — so a persisted/typed hyperlink can never become a script-execution or
 * data-exfil vector when rendered as a clickable <a>. Control characters and
 * surrounding whitespace are stripped before the scheme is checked to defeat
 * obfuscation. The returned href is the parsed, canonical URL string.
 */
export function sanitizeHyperlink(raw: unknown): string | null {
  if (typeof raw !== "string") return null;
  const stripped = stripUrlInsignificantChars(raw);
  if (stripped === "") return null;
  // Protocol-relative URLs resolve to the page origin under whatever scheme the
  // page uses; reject outright so a cell can't silently become same-origin.
  if (stripped.startsWith("//")) return null;
  // A bare scheme check first: reject anything that is not http/https/mailto so
  // we never even hand a `javascript:`/`data:` string to the URL parser.
  const schemeMatch = /^([a-z][a-z0-9+\-.]*):/i.exec(stripped);
  if (!schemeMatch) return null;
  const scheme = `${schemeMatch[1].toLowerCase()}:`;
  if (!HYPERLINK_ALLOWED_SCHEMES.has(scheme)) return null;
  try {
    const url = new URL(stripped);
    if (!HYPERLINK_ALLOWED_SCHEMES.has(url.protocol)) return null;
    return url.toString();
  } catch {
    return null;
  }
}

/**
 * Excel's data bar box, px at 100% (lookwin p158-lw, Windows Excel's printed
 * bars): 2 px in from the cell's left and right lines and its top line, 3.2 px
 * (2.40 pt) above the next row's line, whatever the row's height or the
 * column's width. A bar's place, its length and its zero axis are shares of
 * the box's width, not of the cell's. The API's draw_cell_visuals is the same.
 */
export const DATA_BAR_BOX = { side: 2, top: 2, bottom: 3.2 } as const;
/** The box printed, px of the paper at 96 dpi: 1.44 pt in, 2.40 pt above the foot. */
export const DATA_BAR_PRINTED_BOX = { side: 1.92, top: 1.92, bottom: 3.2 } as const;
type DataBarBox = { side: number; top: number; bottom: number };

/** A data bar's gradient end: its colour 85% of the way to white, each
 * channel truncated (lookwin p158: 156082 fades to DBE7EC), not white. */
export function dataBarFadedColor(color: string): string {
  const hex = /^#([0-9a-f]{6})$/i.exec(color)?.[1];
  if (!hex) return "#ffffff";
  return `#${[0, 2, 4]
    .map((at) => {
      const part = parseInt(hex.slice(at, at + 2), 16);
      return Math.floor(part + (255 - part) * 0.85 + 1e-9).toString(16).padStart(2, "0");
    })
    .join("")}`;
}

function boxPx(value: number): string {
  return `${Number(value.toFixed(3))}px`;
}

/** A share (0..100) of the data bar box's width as a CSS left, in its cell:
 * the box's left plus that share of the cell less both sides. */
export function dataBarBoxLeft(percent: number, scale = 1, box: DataBarBox = DATA_BAR_BOX): string {
  const offset = box.side * scale * (1 - percent / 50);
  return offset === 0 ? `${percent}%` : `calc(${percent}% ${offset < 0 ? "-" : "+"} ${boxPx(Math.abs(offset))})`;
}

/**
 * The CSS a painted data bar draws with: its place and length in Excel's box
 * (DATA_BAR_BOX; a printed page's DATA_BAR_PRINTED_BOX at its scale), its fill (a
 * gradient fades to its colour 85% of the way to white away from the axis, as
 * Excel 2010+ draws it) and, for a bordered bar, its border.
 */
export function dataBarPaintCss(bar: ConditionalDataBar, scale = 1, box: DataBarBox = DATA_BAR_BOX): {
  left: string;
  width: string;
  top: string;
  bottom: string;
  backgroundColor: string;
  backgroundImage?: string;
  border?: string;
  borderRadius?: number;
  opacity?: number;
} {
  const narrower = (2 * box.side * scale * bar.widthPercent) / 100;
  return {
    left: dataBarBoxLeft(bar.startPercent, scale, box),
    width: narrower === 0 ? `${bar.widthPercent}%` : `calc(${bar.widthPercent}% - ${boxPx(narrower)})`,
    top: boxPx(box.top * scale),
    bottom: boxPx(box.bottom * scale),
    backgroundColor: bar.color,
    ...(bar.gradient
      ? {
          backgroundImage: `linear-gradient(to ${bar.direction === "negative" ? "left" : "right"}, ${bar.color}, ${dataBarFadedColor(bar.color)})`,
        }
      : {}),
    ...(bar.borderColor ? { border: `1px solid ${bar.borderColor}` } : {}),
    ...(bar.classic ? { borderRadius: 0, opacity: 1 } : {}),
  };
}
