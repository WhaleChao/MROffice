/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { OfficeRibbonCommandGate } from "../components/office-ribbon/officeRibbonCommandGate";
import {
  gatedByTestId,
  selectionGateFromEntries,
  selectionGateReason,
} from "../components/office-ribbon/officeRibbonSelectionGate";
import {
  type OfficeChartText,
  type OfficeChartTextCommand,
  type OfficeChartTextUnderline,
} from "./office-charts/chart-text";
import {
  stepOfficeFontSizePt,
  type SmartArtTextAlign,
  type SmartArtTextFormatCommand,
  type SmartArtTextFormatSummary,
  type SmartArtTextUnderline,
} from "./office-smartart/text-format";

/** The kinds of object a Home target can be, as the three editors tell them apart. */
export type OfficeHomeObjectKind =
  | "shape"
  | "textBox"
  | "smartArt"
  | "group"
  | "chart"
  | "picture"
  | "ink"
  | "slicer"
  | "table"
  | "object";

/** Home's own vocabulary for one command over an object's text. */
export type OfficeHomeTextCommand = SmartArtTextFormatCommand;

/**
 * What Home shows for the target. A `"mixed"` reads as Office's blank size
 * box / unpressed-but-indeterminate toggle when the objects (or the runs
 * inside them) disagree.
 */
export interface OfficeHomeObjectTextState {
  readonly fontFamily: string | undefined;
  readonly fontFamilyMixed: boolean;
  readonly fontSizePt: number | undefined;
  readonly fontSizeMixed: boolean;
  readonly bold: boolean | "mixed";
  readonly italic: boolean | "mixed";
  /** A chart's text may carry any DrawingML underline style. */
  readonly underline: "none" | SmartArtTextUnderline | OfficeChartTextUnderline | "mixed";
  readonly strike: boolean | "mixed";
  readonly color: string | undefined;
  readonly colorMixed: boolean;
  readonly align: SmartArtTextAlign | undefined | "mixed";
}

export interface OfficeHomeObjectTextTarget {
  readonly kind: OfficeHomeObjectKind;
  /** The ids of the objects the target formats (one, or a multi-selection). */
  readonly ids: readonly string[];
  /** Home's state for the target's text. */
  readonly state: OfficeHomeObjectTextState;
  /** One command over every object's text; `false` when nothing changed. */
  readonly apply: (command: OfficeHomeTextCommand) => boolean;
  /** The reason a Home control cannot act on the target, by test id. */
  readonly gate: OfficeRibbonCommandGate | null;
}

const EMPTY_STATE: OfficeHomeObjectTextState = Object.freeze({
  fontFamily: undefined,
  fontFamilyMixed: false,
  fontSizePt: undefined,
  fontSizeMixed: false,
  bold: false,
  italic: false,
  underline: "none",
  strike: false,
  color: undefined,
  colorMixed: false,
  align: undefined,
});

/** Home's state from the shared engine's summary of the target's objects. */
export function officeHomeObjectTextState(
  summary: SmartArtTextFormatSummary | null,
): OfficeHomeObjectTextState {
  if (!summary) return EMPTY_STATE;
  const toggle = (agreement: { value: boolean; mixed: boolean }) =>
    agreement.mixed ? ("mixed" as const) : agreement.value;
  return Object.freeze({
    fontFamily: summary.fontFamily.mixed ? undefined : summary.fontFamily.value,
    fontFamilyMixed: summary.fontFamily.mixed,
    fontSizePt: summary.fontSizePt.mixed ? undefined : summary.fontSizePt.value,
    fontSizeMixed: summary.fontSizePt.mixed,
    bold: toggle(summary.bold),
    italic: toggle(summary.italic),
    underline: summary.underline.mixed ? "mixed" : summary.underline.value,
    strike: toggle(summary.strike),
    color: summary.color.mixed ? undefined : summary.color.value,
    colorMixed: summary.color.mixed,
    align: summary.align.mixed ? "mixed" : summary.align.value,
  });
}

/** Home's state for a selected chart: the chart's one text, nothing mixed. */
export function officeHomeChartTextState(
  text: OfficeChartText | undefined,
): OfficeHomeObjectTextState {
  return Object.freeze({
    ...EMPTY_STATE,
    fontFamily: text?.fontFamily,
    fontSizePt: text?.fontSizePt,
    bold: text?.bold === true,
    italic: text?.italic === true,
    underline: text?.underline ?? "none",
    color: text?.color,
  });
}

/**
 * Whether `testId` is named by `entries`: an exact test id, or a prefix
 * entry ending in "-" that covers a menu's every item. The one matching
 * rule every editor's greying table uses.
 */
export function officeHomeCommandGatedBy(
  testId: string,
  entries: readonly string[],
): boolean {
  return gatedByTestId(testId, entries);
}

/**
 * A gate that greys `entries` with one reason ("Not available for the
 * selected chart" / "选中的图表不支持此命令"), as `OfficeRibbonCommandGateProvider`
 * consumes it. `null` when nothing is greyed.
 */
export function officeHomeObjectGate(
  entries: readonly string[],
  label: { readonly zh: string; readonly en: string },
  chinese: boolean,
): OfficeRibbonCommandGate | null {
  // One table-to-gate rule for the three editors: the shared helper the
  // measured greying tables already use.
  return selectionGateFromEntries(entries, selectionGateReason(chinese, label));
}

/** Office's chart text default size (the style part's `sz="1000"`). */
export const OFFICE_CHART_TEXT_DEFAULT_SIZE_PT = 10;

/**
 * The chart-text command for a Home command, or `null` for one a chart does
 * not take (strikethrough and alignment have no chart-level home). Grow /
 * Shrink Font steps the chart's one size along Office's ladder.
 */
export function officeChartTextCommandOf(
  command: OfficeHomeTextCommand,
  current: OfficeChartText | undefined,
): OfficeChartTextCommand | null {
  switch (command.kind) {
    case "family":
      return { kind: "family", family: command.family };
    case "variant":
      return { kind: "family", family: command.family };
    case "size":
      return { kind: "size", pt: command.pt };
    case "toggle":
      return command.field === "strike"
        ? null
        : { kind: "toggle", field: command.field };
    case "underline":
      return { kind: "underline", variant: command.variant };
    case "color":
      return { kind: "color", color: command.color };
    case "step": {
      const pt = stepOfficeFontSizePt(
        current?.fontSizePt ?? OFFICE_CHART_TEXT_DEFAULT_SIZE_PT,
        command.direction,
      );
      return pt === null ? null : { kind: "size", pt };
    }
    case "align":
      return null;
  }
}
