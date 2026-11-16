/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The New / Edit Formatting Rule dialog's draft, and the two-way mapping
 * between it and the persisted conditional-format rule model.
 *
 * Excel's dialog is organised by "Select a Rule Type" (six types), each with
 * its own "Edit the Rule Description" pane; the model is organised by rule
 * kind. The mapping lives here, free of React, so the dialog stays a view and
 * every rule still goes through {@link buildRule} (the model's normaliser).
 *
 * Where Excel has a rule the model does not carry as its own kind, the rule is
 * written the way the workbook importer already reads that Excel rule:
 * "Blanks" / "No Blanks" / "Errors" / "No Errors" become formula rules
 * (`=LEN(TRIM(A1))=0`, `=ISERROR(A1)` …) and a Cell Value rule whose operand
 * is text or a reference becomes `=A1>"text"`. Reading a rule back recognises
 * those shapes, so the dialog reopens on the type the user chose.
 */

import {
  buildRule,
  conditionalRuleSupportsStopIfTrue,
  DATE_OCCURRING_PERIODS,
  defaultIconThresholds,
  iconSetCount,
  isRangeConditionalOperator,
  MAX_CONDITIONAL_AREAS,
  ruleAreas,
  type ColorScaleStop,
  type ConditionalDifferentialStyle,
  type ConditionalFormatRule,
  type ConditionalOperator,
  type ConditionalRange,
  type ConditionalRuleKind,
  type DataBarThreshold,
  type DateOccurringPeriod,
  type IconSetName,
  type IconThreshold,
  type IconThresholdType,
  type TextMatchOperator,
} from "./conditional-format";
import { columnToLetters, Parser, tokenize } from "./formula-engine";
import { formulaArityProblem } from "./formula-entry-arity";
import { splitSheetQualifier } from "./name-syntax";
import { parseA1Range } from "./print-settings";

/** Excel's "Select a Rule Type" list, in its order. */
export const CONDITIONAL_RULE_TYPES = [
  "allValues",
  "cellsThatContain",
  "topBottom",
  "aboveBelowAverage",
  "uniqueDuplicate",
  "formula",
] as const;
export type ConditionalRuleTypeId = (typeof CONDITIONAL_RULE_TYPES)[number];

/** "Format all cells based on their values" ▸ Format Style. */
export const CONDITIONAL_VALUE_FORMAT_STYLES = [
  "twoColorScale",
  "threeColorScale",
  "dataBar",
  "iconSet",
] as const;
export type ConditionalValueFormatStyle =
  (typeof CONDITIONAL_VALUE_FORMAT_STYLES)[number];

/** "Format only cells with" ▸ first list, in Excel's order. */
export const CONDITIONAL_CONTAIN_SUBJECTS = [
  "cellValue",
  "specificText",
  "datesOccurring",
  "blanks",
  "noBlanks",
  "errors",
  "noErrors",
] as const;
export type ConditionalContainSubject =
  (typeof CONDITIONAL_CONTAIN_SUBJECTS)[number];

/** Cell Value operators, in the order Excel lists them. */
export const CONDITIONAL_CELL_VALUE_OPERATORS: readonly ConditionalOperator[] = [
  "between",
  "notBetween",
  "=",
  "!=",
  ">",
  "<",
  ">=",
  "<=",
];

/** A threshold family (Excel's Type list): Lowest / Highest Value, Number,
 * Percent, Formula, Percentile; a data bar's endpoints offer Automatic too. */
export type ConditionalThresholdType =
  | "automatic"
  | "min"
  | "max"
  | "number"
  | "percent"
  | "formula"
  | "percentile";

/** The Type list a threshold column offers: a midpoint has no endpoints, and
 * only a data bar's endpoints can be Automatic (Excel 2010's default). */
export function conditionalThresholdTypes(
  position: "min" | "mid" | "max",
  family: "scale" | "bar" = "scale",
): readonly ConditionalThresholdType[] {
  const automatic: ConditionalThresholdType[] = family === "bar" ? ["automatic"] : [];
  if (position === "min") return [...automatic, "min", "number", "percent", "formula", "percentile"];
  if (position === "max") return [...automatic, "max", "number", "percent", "formula", "percentile"];
  return ["number", "percent", "formula", "percentile"];
}

/** "Format values that are: ___ the average for the selected range". */
export const CONDITIONAL_AVERAGE_OPTIONS = [
  "above",
  "below",
  "equalAbove",
  "equalBelow",
  "stdDev1Above",
  "stdDev1Below",
  "stdDev2Above",
  "stdDev2Below",
  "stdDev3Above",
  "stdDev3Below",
] as const;
export type ConditionalAverageOption =
  (typeof CONDITIONAL_AVERAGE_OPTIONS)[number];

export interface ConditionalThresholdDraft {
  type: ConditionalThresholdType;
  value: string;
}

export interface ConditionalScaleStopDraft extends ConditionalThresholdDraft {
  color: string;
}

export interface ConditionalRuleDraft {
  type: ConditionalRuleTypeId;
  formatStyle: ConditionalValueFormatStyle;
  scale: {
    min: ConditionalScaleStopDraft;
    mid: ConditionalScaleStopDraft;
    max: ConditionalScaleStopDraft;
  };
  bar: {
    min: ConditionalThresholdDraft;
    max: ConditionalThresholdDraft;
    color: string;
    gradient: boolean;
    barOnly: boolean;
  };
  icon: {
    set: IconSetName;
    /** Ascending bucket boundaries, one fewer than icons: a percent (0–100),
     * a number, a percentile (0–100) or a formula, per `types`. */
    thresholds: string[];
    /** The family of each boundary (Excel's Type column). */
    types: IconThresholdType[];
    /** true → ">=" (Excel's default), false → ">" for that boundary. */
    gte: boolean[];
    reverse: boolean;
    iconOnly: boolean;
  };
  subject: ConditionalContainSubject;
  operator: ConditionalOperator;
  value1: string;
  value2: string;
  textOperator: TextMatchOperator;
  text: string;
  period: DateOccurringPeriod;
  bottom: boolean;
  rank: string;
  percent: boolean;
  average: ConditionalAverageOption;
  unique: boolean;
  formula: string;
  /** The differential style (Excel's Format… button); empty = No Format Set. */
  format: ConditionalDifferentialStyle;
}

// Excel's New Formatting Rule defaults for a scale: a 2-colour scale runs
// orange to yellow, a 3-colour scale red, yellow, green with the midpoint at
// the 50th percentile; a data bar is blue.
const TWO_SCALE_MIN = /* own data: TWO_SCALE_MIN */ "#f58d47";
const TWO_SCALE_MAX = /* own data: TWO_SCALE_MAX */ "#fbeda7";
const THREE_SCALE_MIN = /* own data: THREE_SCALE_MIN */ "#e56173";
const THREE_SCALE_MID = /* own data: THREE_SCALE_MID */ "#ead07b";
const THREE_SCALE_MAX = /* own data: THREE_SCALE_MAX */ "#4fba7c";
const DEFAULT_BAR = /* own data: DEFAULT_BAR */ "#5c8ccc";

/** Excel's "Light Red Fill with Dark Red Text", the Highlight rules' default. */
export const CONDITIONAL_LIGHT_RED_FORMAT: ConditionalDifferentialStyle = /* own data: CONDITIONAL_LIGHT_RED_FORMAT */ {"textColor":"#72131f","backgroundColor":"#f4d2d6"};

function iconThresholdStrings(set: IconSetName): string[] {
  return defaultIconThresholds(iconSetCount(set)).map((value) =>
    String(Math.round(value * 100)),
  );
}

/** A fresh draft: Excel opens New Formatting Rule on "Format all cells based
 * on their values", 2-Color Scale, with no format set. */
export function defaultConditionalRuleDraft(): ConditionalRuleDraft {
  return {
    type: "allValues",
    formatStyle: "twoColorScale",
    scale: {
      min: { type: "min", value: "", color: TWO_SCALE_MIN },
      mid: { type: "percentile", value: "50", color: THREE_SCALE_MID },
      max: { type: "max", value: "", color: TWO_SCALE_MAX },
    },
    bar: {
      min: { type: "min", value: "" },
      max: { type: "max", value: "" },
      color: DEFAULT_BAR,
      gradient: true,
      barOnly: false,
    },
    icon: {
      set: "3-traffic-lights",
      thresholds: iconThresholdStrings("3-traffic-lights"),
      types: iconThresholdStrings("3-traffic-lights").map(() => "percent" as const),
      gte: iconThresholdStrings("3-traffic-lights").map(() => true),
      reverse: false,
      iconOnly: false,
    },
    subject: "cellValue",
    operator: "between",
    value1: "",
    value2: "",
    textOperator: "contains",
    text: "",
    period: "yesterday",
    bottom: false,
    rank: "10",
    percent: false,
    average: "above",
    unique: false,
    formula: "",
    format: {},
  };
}

/** Switching Format Style re-seeds the scale colours Excel shows for it. */
export function withConditionalFormatStyle(
  draft: ConditionalRuleDraft,
  formatStyle: ConditionalValueFormatStyle,
): ConditionalRuleDraft {
  if (draft.formatStyle === formatStyle) return draft;
  const next = { ...draft, formatStyle };
  if (formatStyle === "twoColorScale") {
    next.scale = {
      ...draft.scale,
      min: { ...draft.scale.min, color: TWO_SCALE_MIN },
      max: { ...draft.scale.max, color: TWO_SCALE_MAX },
    };
  } else if (formatStyle === "threeColorScale") {
    next.scale = {
      min: { ...draft.scale.min, color: THREE_SCALE_MIN },
      mid: { ...draft.scale.mid, color: THREE_SCALE_MID },
      max: { ...draft.scale.max, color: THREE_SCALE_MAX },
    };
  }
  return next;
}

/** Choosing an icon style resets its thresholds to Excel's even split. */
export function withConditionalIconSet(
  draft: ConditionalRuleDraft,
  set: IconSetName,
): ConditionalRuleDraft {
  const thresholds = iconThresholdStrings(set);
  return {
    ...draft,
    icon: {
      ...draft.icon,
      set,
      thresholds,
      types: thresholds.map(() => "percent" as const),
      gte: thresholds.map(() => true),
    },
  };
}

/** The Home ▸ Conditional Formatting entries that open a prefilled rule. */
export type ConditionalRibbonRuleDraft =
  | "greaterThan"
  | "lessThan"
  | "between"
  | "equals"
  | "textContains"
  | "dateOccurring"
  | "duplicateValues"
  | "topItems"
  | "topPercent"
  | "bottomItems"
  | "bottomPercent"
  | "aboveAverage"
  | "belowAverage";

/** The rule a Highlight Cells / Top-Bottom entry starts from, formatted with
 * Excel's default Light Red Fill with Dark Red Text. */
export function conditionalRuleDraftForRibbon(
  entry: ConditionalRibbonRuleDraft,
): ConditionalRuleDraft {
  const draft: ConditionalRuleDraft = {
    ...defaultConditionalRuleDraft(),
    format: { ...CONDITIONAL_LIGHT_RED_FORMAT },
  };
  switch (entry) {
    case "greaterThan":
    case "lessThan":
    case "equals":
    case "between":
      return {
        ...draft,
        type: "cellsThatContain",
        subject: "cellValue",
        operator:
          entry === "greaterThan"
            ? ">"
            : entry === "lessThan"
              ? "<"
              : entry === "equals"
                ? "="
                : "between",
      };
    case "textContains":
      return { ...draft, type: "cellsThatContain", subject: "specificText" };
    case "dateOccurring":
      return {
        ...draft,
        type: "cellsThatContain",
        subject: "datesOccurring",
        period: "yesterday",
      };
    case "duplicateValues":
      return { ...draft, type: "uniqueDuplicate", unique: false };
    case "topItems":
    case "topPercent":
    case "bottomItems":
    case "bottomPercent":
      return {
        ...draft,
        type: "topBottom",
        bottom: entry === "bottomItems" || entry === "bottomPercent",
        percent: entry === "topPercent" || entry === "bottomPercent",
      };
    case "aboveAverage":
    case "belowAverage":
      return {
        ...draft,
        type: "aboveBelowAverage",
        average: entry === "belowAverage" ? "below" : "above",
      };
  }
}

// ── Applies to ───────────────────────────────────────────────────────────────

function absoluteArea(area: ConditionalRange): string {
  const start = `$${columnToLetters(area.c1)}$${area.r1 + 1}`;
  const end = `$${columnToLetters(area.c2)}$${area.r2 + 1}`;
  return start === end ? start : `${start}:${end}`;
}

/** Excel's Applies to text: `=$C$4:$C$10`, areas joined by commas. */
export function formatConditionalAppliesTo(
  areas: readonly ConditionalRange[],
): string {
  return `=${areas.map(absoluteArea).join(",")}`;
}

/** The rule's own areas as Applies to text. */
export function conditionalRuleAppliesTo(rule: ConditionalFormatRule): string {
  return formatConditionalAppliesTo(ruleAreas(rule));
}

/**
 * The areas of a reference list, split at the `,` / `;` union operators that
 * stand OUTSIDE a quoted sheet name: in `'CF,Audit'!$A$1:$A$2` the comma is
 * part of the name. A doubled `''` inside a quoted name toggles twice, so it
 * stays inside the name.
 */
function splitReferenceUnion(body: string): string[] {
  const parts: string[] = [];
  let quoted = false;
  let start = 0;
  for (let i = 0; i < body.length; i += 1) {
    const character = body[i];
    if (character === "'") quoted = !quoted;
    else if (!quoted && (character === "," || character === ";")) {
      parts.push(body.slice(start, i));
      start = i + 1;
    }
  }
  parts.push(body.slice(start));
  return parts.map((part) => part.trim());
}

/**
 * Read an Applies to reference: optional `=`, `$` anchors, one or more
 * comma-separated areas, each optionally qualified with this sheet's name.
 * Returns null for anything outside the sheet (Excel refuses a reference to
 * another worksheet here). `bounds` is the worksheet's address space, not its
 * used range: a rule may be set ahead on rows not yet filled in.
 */
export function parseConditionalAppliesTo(
  text: string,
  bounds: { rowCount: number; columnCount: number; sheetName?: string },
): ConditionalRange[] | null {
  const body = text.trim().replace(/^=/, "");
  if (body === "") return null;
  const parts = splitReferenceUnion(body);
  if (parts.length > MAX_CONDITIONAL_AREAS || parts.some((part) => part === "")) {
    return null;
  }
  const areas: ConditionalRange[] = [];
  for (const part of parts) {
    const { sheet, ref: reference } = splitSheetQualifier(part);
    if (sheet !== undefined && (!bounds.sheetName || sheet !== bounds.sheetName)) {
      return null;
    }
    // A quote left open, or a qualifier spelled so it needs quotes, is not a
    // reference at all.
    if (sheet === undefined && part.includes("!")) return null;
    const area = parseA1Range(reference.replace(/\$/g, ""));
    if (
      !area ||
      area.r1 < 0 ||
      area.c1 < 0 ||
      area.r2 >= bounds.rowCount ||
      area.c2 >= bounds.columnCount
    ) {
      return null;
    }
    areas.push(area);
  }
  return areas;
}

/** The rule moved onto new areas; a formula stays relative to the new
 * top-left cell, as Excel keeps it when Applies to changes. */
export function withConditionalRuleAreas(
  rule: ConditionalFormatRule,
  areas: readonly ConditionalRange[],
): ConditionalFormatRule {
  const next = { ...rule, range: { ...areas[0] } } as ConditionalFormatRule;
  if (areas.length > 1) next.areas = areas.map((area) => ({ ...area }));
  else delete next.areas;
  return next;
}

// ── Rule → draft ─────────────────────────────────────────────────────────────

function anchorOf(range: ConditionalRange): string {
  return `${columnToLetters(range.c1)}${range.r1 + 1}`;
}

const CELL_REF = String.raw`\$?[A-Z]{1,3}\$?\d{1,7}`;
const OPERAND = String.raw`(?:"(?:[^"]|"")*"|${CELL_REF}|-?\d+(?:\.\d+)?(?:E[+-]?\d+)?)`;

function sameCell(reference: string, anchor: string): boolean {
  return reference.replace(/\$/g, "").toUpperCase() === anchor;
}

/** A formula operand as the Value box shows it: text unquoted, a reference
 * with its `=`, a number as typed. */
function operandText(operand: string): string {
  if (operand.startsWith('"')) return operand.slice(1, -1).replace(/""/g, '"');
  if (/^-?\d/.test(operand)) return operand;
  return `=${operand}`;
}

const FORMULA_COMPARATORS: Record<string, ConditionalOperator> = {
  ">": ">",
  "<": "<",
  ">=": ">=",
  "<=": "<=",
  "=": "=",
  "<>": "!=",
};

type RecognisedFormula =
  | { subject: "blanks" | "noBlanks" | "errors" | "noErrors" }
  | {
      subject: "cellValue";
      operator: ConditionalOperator;
      value1: string;
      value2: string;
    };

/** Only an operand the model's number-only Cell Value rule cannot hold
 * (text, a reference) is written as a formula; `=A1>1` is a formula rule of
 * its own (it also matches text, which Cell Value > 1 does not). */
function nonNumericOperand(operand: string): boolean {
  return !/^-?\d/.test(operand);
}

/** The Excel rule a formula rule was written for, when it is one of the
 * shapes this module (or the importer) writes. */
function recogniseFormula(
  formula: string,
  range: ConditionalRange,
): RecognisedFormula | null {
  const anchor = anchorOf(range);
  const text = formula.trim().replace(/^=/, "").trim();
  let match = new RegExp(String.raw`^LEN\(TRIM\((${CELL_REF})\)\)(=0|>0)$`, "i").exec(text);
  if (match && sameCell(match[1], anchor)) {
    return { subject: match[2] === "=0" ? "blanks" : "noBlanks" };
  }
  match = new RegExp(String.raw`^(NOT\()?ISERROR\((${CELL_REF})\)(\))?$`, "i").exec(text);
  if (match && sameCell(match[2], anchor) && Boolean(match[1]) === Boolean(match[3])) {
    return { subject: match[1] ? "noErrors" : "errors" };
  }
  match = new RegExp(
    String.raw`^AND\((${CELL_REF})>=(${OPERAND}),(${CELL_REF})<=(${OPERAND})\)$`,
    "i",
  ).exec(text);
  if (
    match &&
    sameCell(match[1], anchor) &&
    sameCell(match[3], anchor) &&
    (nonNumericOperand(match[2]) || nonNumericOperand(match[4]))
  ) {
    return {
      subject: "cellValue",
      operator: "between",
      value1: operandText(match[2]),
      value2: operandText(match[4]),
    };
  }
  match = new RegExp(
    String.raw`^OR\((${CELL_REF})<(${OPERAND}),(${CELL_REF})>(${OPERAND})\)$`,
    "i",
  ).exec(text);
  if (
    match &&
    sameCell(match[1], anchor) &&
    sameCell(match[3], anchor) &&
    (nonNumericOperand(match[2]) || nonNumericOperand(match[4]))
  ) {
    return {
      subject: "cellValue",
      operator: "notBetween",
      value1: operandText(match[2]),
      value2: operandText(match[4]),
    };
  }
  match = new RegExp(String.raw`^(${CELL_REF})(<>|>=|<=|>|<|=)(${OPERAND})$`, "i").exec(text);
  if (match && sameCell(match[1], anchor) && nonNumericOperand(match[3])) {
    return {
      subject: "cellValue",
      operator: FORMULA_COMPARATORS[match[2]],
      value1: operandText(match[3]),
      value2: "",
    };
  }
  return null;
}

function thresholdDraft(
  stop: Partial<ColorScaleStop> | DataBarThreshold | undefined,
  position: "min" | "mid" | "max",
): ConditionalThresholdDraft {
  const type = stop?.type;
  if (type === "automatic") return { type: "automatic", value: "" };
  if (type === "formula") {
    return { type: "formula", value: `=${stop?.formula ?? ""}` };
  }
  if (type === "number" || type === "percent" || type === "percentile") {
    return { type, value: String(stop?.value ?? "") };
  }
  if (position === "mid") {
    // An untyped midpoint is Excel's default one, the 50th PERCENTILE: the
    // xlsx projection reads `percentile 50` that way, both writers write it
    // back so, and the grid draws it so (XLS-A047). The draft says the same;
    // opening and confirming a rule must not move its midpoint (B328).
    return { type: "percentile", value: "50" };
  }
  return { type: position, value: "" };
}

function averageOption(
  below: boolean,
  equalAverage: boolean,
  stdDevs: number | undefined,
): ConditionalAverageOption {
  if (stdDevs && stdDevs > 0) {
    const k = Math.min(3, Math.max(1, Math.round(stdDevs)));
    return `stdDev${k}${below ? "Below" : "Above"}` as ConditionalAverageOption;
  }
  if (equalAverage) return below ? "equalBelow" : "equalAbove";
  return below ? "below" : "above";
}

/** The dialog draft that shows an existing rule. */
export function conditionalRuleDraftFromRule(
  rule: ConditionalFormatRule,
): ConditionalRuleDraft {
  const base = defaultConditionalRuleDraft();
  const format = conditionalRuleFormat(rule);
  switch (rule.kind) {
    case "colorScale": {
      const three = Boolean(rule.mid);
      return {
        ...base,
        type: "allValues",
        formatStyle: three ? "threeColorScale" : "twoColorScale",
        scale: {
          min: { ...thresholdDraft(rule.min, "min"), color: rule.min.color },
          mid: {
            ...thresholdDraft(rule.mid, "mid"),
            color: rule.mid?.color ?? THREE_SCALE_MID,
          },
          max: { ...thresholdDraft(rule.max, "max"), color: rule.max.color },
        },
      };
    }
    case "dataBar":
      return {
        ...base,
        type: "allValues",
        formatStyle: "dataBar",
        bar: {
          min: thresholdDraft(rule.min, "min"),
          max: thresholdDraft(rule.max, "max"),
          color: rule.color,
          gradient: rule.gradient !== false,
          barOnly: rule.showValue === false,
        },
      };
    case "iconSet":
      return {
        ...base,
        type: "allValues",
        formatStyle: "iconSet",
        icon: {
          set: rule.iconSet,
          thresholds: rule.stops
            ? rule.stops.map((stop) =>
                stop.type === "formula" ? `=${stop.formula ?? ""}` : String(stop.value ?? ""),
              )
            : rule.thresholds.map((value) => String(Number((value * 100).toFixed(2)))),
          types: rule.stops
            ? rule.stops.map((stop) => stop.type)
            : rule.thresholds.map(() => "percent" as const),
          gte: rule.stops
            ? rule.stops.map((stop) => stop.gte !== false)
            : rule.thresholds.map(() => true),
          reverse: rule.reverse === true,
          iconOnly: rule.showIconOnly === true,
        },
      };
    case "formula": {
      const recognised = recogniseFormula(rule.formula, rule.range);
      if (recognised?.subject === "cellValue") {
        return {
          ...base,
          type: "cellsThatContain",
          subject: "cellValue",
          operator: recognised.operator,
          value1: recognised.value1,
          value2: recognised.value2,
          format,
        };
      }
      if (recognised) {
        return { ...base, type: "cellsThatContain", subject: recognised.subject, format };
      }
      return {
        ...base,
        type: "formula",
        formula: rule.formula.startsWith("=") ? rule.formula : `=${rule.formula}`,
        format,
      };
    }
    case "textContains":
      return {
        ...base,
        type: "cellsThatContain",
        subject: "specificText",
        textOperator: rule.operator,
        text: rule.text,
        format,
      };
    case "dateOccurring":
      return {
        ...base,
        type: "cellsThatContain",
        subject: "datesOccurring",
        period: rule.period,
        format,
      };
    case "duplicateValues":
      return { ...base, type: "uniqueDuplicate", unique: rule.unique === true, format };
    case "topBottom":
      return {
        ...base,
        type: "topBottom",
        bottom: rule.bottom === true,
        percent: rule.percent === true,
        rank: String(rule.rank),
        format,
      };
    case "aboveBelowAverage":
      return {
        ...base,
        type: "aboveBelowAverage",
        average: averageOption(
          rule.below === true,
          rule.equalAverage === true,
          rule.stdDevs,
        ),
        format,
      };
    case "comparison":
    default:
      return {
        ...base,
        type: "cellsThatContain",
        subject: "cellValue",
        operator: rule.operator,
        value1: String(rule.value),
        value2: rule.value2 === undefined ? "" : String(rule.value2),
        format,
      };
  }
}

const DIFFERENTIAL_KEYS = [
  "textColor",
  "textThemeColor",
  "backgroundColor",
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
] as const;

/** The differential style a style rule paints; empty for the visual kinds. */
export function conditionalRuleFormat(
  rule: ConditionalFormatRule,
): ConditionalDifferentialStyle {
  const kind = rule.kind ?? "comparison";
  if (kind === "colorScale" || kind === "dataBar" || kind === "iconSet") return {};
  const source = rule as Partial<ConditionalDifferentialStyle>;
  const out: ConditionalDifferentialStyle = {};
  for (const key of DIFFERENTIAL_KEYS) {
    const value = source[key];
    if (value !== undefined) (out as Record<string, unknown>)[key] = value;
  }
  return out;
}

/** Whether the draft has any format at all (Excel's "No Format Set"). */
export function conditionalFormatIsSet(format: ConditionalDifferentialStyle): boolean {
  return DIFFERENTIAL_KEYS.some((key) => format[key] !== undefined);
}

// ── Draft → rule ─────────────────────────────────────────────────────────────

/** The model kind a draft writes (the native-write gate asks about it). */
export function conditionalDraftKind(draft: ConditionalRuleDraft): ConditionalRuleKind {
  switch (draft.type) {
    case "allValues":
      return draft.formatStyle === "dataBar"
        ? "dataBar"
        : draft.formatStyle === "iconSet"
          ? "iconSet"
          : "colorScale";
    case "cellsThatContain":
      if (draft.subject === "specificText") return "textContains";
      if (draft.subject === "datesOccurring") return "dateOccurring";
      if (draft.subject === "cellValue") {
        const numeric =
          isNumericOperand(draft.value1) &&
          (!isRangeConditionalOperator(draft.operator) || isNumericOperand(draft.value2));
        return numeric ? "comparison" : "formula";
      }
      return "formula";
    case "topBottom":
      return "topBottom";
    case "aboveBelowAverage":
      return "aboveBelowAverage";
    case "uniqueDuplicate":
      return "duplicateValues";
    case "formula":
    default:
      return "formula";
  }
}

/**
 * Whether Excel takes `body` (a formula without its "=") as a rule's formula:
 * it parses and calls no function with a refused argument count. MEASURED in
 * Excel for Mac 16: adding an expression rule "=1+" fails ("parameter error")
 * and leaves the range's rules as they were, where "=A1>=20" is added.
 */
function formulaBodyAccepted(body: string): boolean {
  const tokens = tokenize(body);
  if (!tokens) return false;
  try {
    new Parser(tokens).parse();
  } catch {
    return false;
  }
  return formulaArityProblem(`=${body}`) === null;
}

function isNumericOperand(text: string): boolean {
  const trimmed = text.trim();
  return trimmed !== "" && Number.isFinite(Number(trimmed));
}

/** A Value box entry as a formula operand: `=B1` → `B1`, a number as is,
 * anything else as quoted text. */
function formulaOperand(text: string): string | null {
  const trimmed = text.trim();
  if (trimmed === "") return null;
  if (trimmed.startsWith("=")) {
    const body = trimmed.slice(1).trim();
    return body === "" || !formulaBodyAccepted(body) ? null : body;
  }
  if (Number.isFinite(Number(trimmed))) return trimmed;
  return `"${trimmed.replace(/"/g, '""')}"`;
}

export type ConditionalDraftProblem =
  | "value"
  | "text"
  | "formula"
  | "formulaSyntax"
  | "rank"
  | "threshold"
  | "format"
  | "invalid";

export type ConditionalDraftBuild =
  | { ok: true; rule: ConditionalFormatRule }
  | { ok: false; problem: ConditionalDraftProblem };

function isColorScaleThreshold(
  threshold: DataBarThreshold,
): threshold is Omit<ColorScaleStop, "color" | "themeColor"> {
  return threshold.type !== "automatic";
}

function thresholdStop(
  draft: ConditionalThresholdDraft,
  position: "min" | "mid" | "max",
): DataBarThreshold | null {
  if (draft.type === "min" || draft.type === "max") {
    return position === "mid" ? null : {};
  }
  if (draft.type === "automatic") {
    return position === "mid" ? null : { type: "automatic" };
  }
  if (draft.type === "formula") {
    const formula = draft.value.trim().replace(/^=/, "").trim();
    return formula === "" ? null : { type: "formula", formula };
  }
  const value = Number(draft.value.trim());
  if (draft.value.trim() === "" || !Number.isFinite(value)) return null;
  if ((draft.type === "percent" || draft.type === "percentile") && (value < 0 || value > 100)) {
    return null;
  }
  return { type: draft.type, value };
}

/**
 * Build the rule a draft describes over `range`. `previous` is the rule being
 * edited: its identity (id, native priority, Stop If True) and the members the
 * dialog has no control for (data-bar lengths and axis, custom icons, native
 * extension ids, theme slots of unchanged colours) ride along unchanged.
 */
export function conditionalRuleFromDraft(
  draft: ConditionalRuleDraft,
  areas: readonly ConditionalRange[],
  options: {
    previous?: ConditionalFormatRule | null;
    /** An imported workbook written back into its own package. */
    patchback?: boolean;
  } = {},
): ConditionalDraftBuild {
  const range = areas[0];
  if (!range) return { ok: false, problem: "invalid" };
  const previous = options.previous ?? null;
  const format = draft.format;
  const needsFormat = draft.type !== "allValues";
  // A formula that cannot be read is the first thing wrong with the rule.
  const typedFormula = draft.type === "formula" ? draft.formula.trim() : "";
  if (typedFormula !== "" && typedFormula !== "=" &&
      !formulaBodyAccepted(typedFormula.startsWith("=") ? typedFormula.slice(1) : typedFormula)) {
    return { ok: false, problem: "formulaSyntax" };
  }
  if (needsFormat && !conditionalFormatIsSet(format)) {
    return { ok: false, problem: "format" };
  }
  let params: Record<string, unknown>;
  switch (draft.type) {
    case "allValues": {
      if (draft.formatStyle === "dataBar") {
        const min = thresholdStop(draft.bar.min, "min");
        const max = thresholdStop(draft.bar.max, "max");
        if (!min || !max) return { ok: false, problem: "threshold" };
        const bar = previous?.kind === "dataBar" ? previous : null;
        params = {
          kind: "dataBar",
          color: draft.bar.color,
          // A new bar gets Excel 2010's full-width look; an edited one keeps
          // its own lengths and axis.
          minLength: bar ? bar.minLength : 0,
          maxLength: bar ? bar.maxLength : 100,
          ...(draft.bar.barOnly ? { showValue: false } : {}),
          // Solid or gradient fill is Excel 2010's x14 setting; the original-
          // package writer gives a new bar the x14 twin Excel writes, so the
          // choice is kept on either save route.
          gradient: draft.bar.gradient,
          ...(bar
            ? {
                negativeColor: bar.negativeColor,
                axisColor: bar.axisColor,
                borderColor: bar.borderColor,
                axisPosition: bar.axisPosition,
                nativeDataBarId: bar.nativeDataBarId,
                themeColor: bar.color === draft.bar.color ? bar.themeColor : undefined,
              }
            : {}),
          ...(min.type ? { min } : {}),
          ...(max.type ? { max } : {}),
        };
        break;
      }
      if (draft.formatStyle === "iconSet") {
        const plainPercent = draft.icon.thresholds.every(
          (_value, index) =>
            (draft.icon.types[index] ?? "percent") === "percent" &&
            draft.icon.gte[index] !== false,
        );
        const thresholds = draft.icon.thresholds.map((value) => Number(value.trim()));
        if (
          plainPercent &&
          (draft.icon.thresholds.some((value) => value.trim() === "") ||
            thresholds.some((value) => !Number.isFinite(value) || value <= 0 || value >= 100) ||
            thresholds.some((value, index) => index > 0 && value <= thresholds[index - 1]))
        ) {
          return { ok: false, problem: "threshold" };
        }
        // Excel's Value / Type / ">=" columns: any boundary that is not a
        // plain ">=" percent makes the rule carry typed stops.
        const stops: IconThreshold[] = [];
        if (!plainPercent) {
          for (let index = 0; index < draft.icon.thresholds.length; index += 1) {
            const type = draft.icon.types[index] ?? "percent";
            const text = draft.icon.thresholds[index].trim();
            let stop: IconThreshold;
            if (type === "formula") {
              const formula = text.replace(/^=/, "").trim();
              if (formula === "") return { ok: false, problem: "threshold" };
              stop = { type, formula };
            } else {
              const value = Number(text);
              if (text === "" || !Number.isFinite(value)) return { ok: false, problem: "threshold" };
              if (type !== "number" && (value < 0 || value > 100)) {
                return { ok: false, problem: "threshold" };
              }
              stop = { type, value };
            }
            if (draft.icon.gte[index] === false) stop.gte = false;
            stops.push(stop);
          }
        }
        const icons = previous?.kind === "iconSet" ? previous : null;
        params = {
          kind: "iconSet",
          iconSet: draft.icon.set,
          thresholds: plainPercent
            ? thresholds.map((value) => value / 100)
            : defaultIconThresholds(iconSetCount(draft.icon.set)),
          ...(plainPercent ? {} : { stops }),
          reverse: draft.icon.reverse,
          showIconOnly: draft.icon.iconOnly,
          ...(icons?.customIcons && icons.iconSet === draft.icon.set
            ? { customIcons: icons.customIcons }
            : {}),
          ...(icons?.nativeX14Id ? { nativeX14Id: icons.nativeX14Id } : {}),
        };
        break;
      }
      const three = draft.formatStyle === "threeColorScale";
      const scale = previous?.kind === "colorScale" ? previous : null;
      const stop = (
        position: "min" | "mid" | "max",
      ): (ColorScaleStop & DataBarThreshold) | null => {
        const source = draft.scale[position];
        const threshold = thresholdStop(source, position);
        // A colour scale's Type list has no Automatic.
        if (!threshold || !isColorScaleThreshold(threshold)) return null;
        const before = scale?.[position];
        return {
          ...threshold,
          color: source.color,
          ...(before?.themeColor && before.color === source.color
            ? { themeColor: before.themeColor }
            : {}),
        };
      };
      const min = stop("min");
      const max = stop("max");
      const mid = three ? stop("mid") : undefined;
      if (!min || !max || mid === null) return { ok: false, problem: "threshold" };
      params = { kind: "colorScale", min, max, ...(mid ? { mid } : {}) };
      break;
    }
    case "cellsThatContain": {
      const anchor = anchorOf(range);
      switch (draft.subject) {
        case "specificText": {
          // The needle is kept as typed: its spaces are part of the search.
          const text = draft.text;
          if (text === "") return { ok: false, problem: "text" };
          params = { kind: "textContains", operator: draft.textOperator, text };
          break;
        }
        case "datesOccurring":
          params = { kind: "dateOccurring", period: draft.period };
          break;
        case "blanks":
        case "noBlanks":
          params = {
            kind: "formula",
            formula: `=LEN(TRIM(${anchor}))${draft.subject === "blanks" ? "=0" : ">0"}`,
          };
          break;
        case "errors":
          params = { kind: "formula", formula: `=ISERROR(${anchor})` };
          break;
        case "noErrors":
          params = { kind: "formula", formula: `=NOT(ISERROR(${anchor}))` };
          break;
        case "cellValue":
        default: {
          const between = isRangeConditionalOperator(draft.operator);
          if (
            isNumericOperand(draft.value1) &&
            (!between || isNumericOperand(draft.value2))
          ) {
            params = {
              operator: draft.operator,
              value: Number(draft.value1.trim()),
              ...(between ? { value2: Number(draft.value2.trim()) } : {}),
            };
            break;
          }
          const first = formulaOperand(draft.value1);
          const second = between ? formulaOperand(draft.value2) : "";
          if (first === null || second === null) return { ok: false, problem: "value" };
          const formula =
            draft.operator === "between"
              ? `=AND(${anchor}>=${first},${anchor}<=${second})`
              : draft.operator === "notBetween"
                ? `=OR(${anchor}<${first},${anchor}>${second})`
                : `=${anchor}${draft.operator === "!=" ? "<>" : draft.operator}${first}`;
          params = { kind: "formula", formula };
          break;
        }
      }
      break;
    }
    case "topBottom": {
      const rank = Number(draft.rank.trim());
      if (
        !Number.isInteger(rank) ||
        rank < 1 ||
        rank > (draft.percent ? 100 : 1000)
      ) {
        return { ok: false, problem: "rank" };
      }
      params = {
        kind: "topBottom",
        rank,
        ...(draft.bottom ? { bottom: true } : {}),
        ...(draft.percent ? { percent: true } : {}),
      };
      break;
    }
    case "aboveBelowAverage": {
      const option = draft.average;
      const below = option.includes("elow");
      const std = /^stdDev(\d)/.exec(option);
      params = {
        kind: "aboveBelowAverage",
        ...(below ? { below: true } : {}),
        ...(option === "equalAbove" || option === "equalBelow" ? { equalAverage: true } : {}),
        ...(std ? { stdDevs: Number(std[1]) } : {}),
      };
      break;
    }
    case "uniqueDuplicate":
      params = { kind: "duplicateValues", ...(draft.unique ? { unique: true } : {}) };
      break;
    case "formula":
    default: {
      const formula = draft.formula.trim();
      if (formula === "" || formula === "=") return { ok: false, problem: "formula" };
      // A formula Excel cannot read is refused, not saved as a rule that
      // never applies (live audit P-CF-FORMULA-SYNTAX-FEEDBACK: "=1+").
      if (!formulaBodyAccepted(formula.startsWith("=") ? formula.slice(1) : formula)) {
        return { ok: false, problem: "formulaSyntax" };
      }
      params = { kind: "formula", formula: formula.startsWith("=") ? formula : `=${formula}` };
      break;
    }
  }
  const body: Record<string, unknown> = needsFormat ? { ...format, ...params } : params;
  const rule = buildRule(range, body);
  if (!rule) return { ok: false, problem: "invalid" };
  if (areas.length > 1) rule.areas = areas.map((area) => ({ ...area }));
  if (previous?.id) rule.id = previous.id;
  if (previous?.nativePriority !== undefined) rule.nativePriority = previous.nativePriority;
  if (previous?.nativeSequence !== undefined) rule.nativeSequence = previous.nativeSequence;
  if (
    previous?.kind === "textContains" &&
    previous.formula !== undefined &&
    rule.kind === "textContains" &&
    rule.text === previous.text &&
    rule.operator === previous.operator &&
    rule.range.r1 === previous.range.r1 &&
    rule.range.c1 === previous.range.c1
  ) {
    
    // a changed text is the predicate Excel writes for it.
    rule.formula = previous.formula;
  }
  if (previous?.stopIfTrue && conditionalRuleSupportsStopIfTrue(rule)) {
    rule.stopIfTrue = true;
  }
  return { ok: true, rule };
}

/** A fresh opaque editor identity for a newly authored rule. */
export function createConditionalRuleId(): string {
  const uuid = globalThis.crypto?.randomUUID?.();
  return uuid
    ? `cf-${uuid}`
    : `cf-${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
}

/** The ten Dates Occurring periods, in Excel's order. */
export const CONDITIONAL_DATE_PERIODS: readonly DateOccurringPeriod[] =
  DATE_OCCURRING_PERIODS;
