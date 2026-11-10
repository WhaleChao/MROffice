/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { OfficeRibbonCommandGate } from "@mrnote/core/web/components/office-ribbon/officeRibbonCommandGate.tsx";
import { selectionGateFromEntries } from "@mrnote/core/web/components/office-ribbon/officeRibbonSelectionGate.ts";

/** The kinds of drawing object a worksheet selection can be. */
export type SpreadsheetSelectedStructure =
  | "chart"
  | "picture"
  | "shape"
  | "textBox"
  | "smartArtShape"
  | "smartArtGraphic"
  | "slicer";

/** What the window knows about its selection, as the gate reads it. */
export interface SpreadsheetSelectionFacts {
  /** The selected drawing object, if one is (its `type` and shape kind). */
  readonly selectedObject: {
    readonly type: string;
    readonly shape?: { readonly kind?: string } | null;
    readonly groupId?: string;
    readonly smartArtRole?: string;
  } | null;
  /** The whole SmartArt graphic is selected (its frame, or Esc from a shape). */
  readonly smartArtGraphicSelected: boolean;
  /** A native (workbook) chart is selected; charts are not drawing objects. */
  readonly selectedChartId: string | null;
  /** A slicer is selected. */
  readonly selectedSlicerId: string | null;
  /** The user is typing in a cell (or a shape's text): nothing is greyed. */
  readonly editingText: boolean;
}

/** The object the selection is, or null for cells (and while typing). */
export function spreadsheetSelectedStructure(
  facts: SpreadsheetSelectionFacts,
): SpreadsheetSelectedStructure | null {
  if (facts.editingText) return null;
  if (facts.selectedChartId) return "chart";
  if (facts.selectedSlicerId) return "slicer";
  const object = facts.selectedObject;
  if (!object) return null;
  if (facts.smartArtGraphicSelected) return "smartArtGraphic";
  if (object.type === "smartart") return "smartArtGraphic";
  if (typeof object.smartArtRole === "string" && object.groupId) {
    return "smartArtShape";
  }
  if (object.type === "chart") return "chart";
  if (object.type === "image") return "picture";
  if (object.type === "shape") {
    return object.shape?.kind === "text-box" ? "textBox" : "shape";
  }
  return null;
}

/** The catalogue key of the reason line each selected object shows. */
const STRUCTURE_REASON_KEY: Record<SpreadsheetSelectedStructure, string> = {
  chart: "spreadsheet.selectionGate.chart",
  picture: "spreadsheet.selectionGate.picture",
  shape: "spreadsheet.selectionGate.shape",
  textBox: "spreadsheet.selectionGate.textBox",
  smartArtShape: "spreadsheet.selectionGate.smartArtShape",
  smartArtGraphic: "spreadsheet.selectionGate.smartArtGraphic",
  slicer: "spreadsheet.selectionGate.slicer",
};

const HOME = "spreadsheet-ribbon-home-";
const RIBBON = "spreadsheet-ribbon-";

/**
 * The Number group: a drawing object holds no cell value to format. Excel
 * greys the whole group with a chart or a shape selected; measured here, the
 * format menu, the accounting menu, the four quick formats and Increase
 * Decimal changed nothing with a chart or a text box selected while they
 * changed a cell (Decrease Decimal is the same command in the other
 * direction; the sample cell had no decimals to take).
 */
const NUMBER_GROUP = [
  `${HOME}number-format`,
  `${RIBBON}number-format-`,
  `${HOME}accounting-options`,
  `${HOME}accounting-split`,
  `${HOME}accountingCny`,
  `${HOME}accountingEur`,
  `${HOME}accountingGbp`,
  `${HOME}accountingUsd`,
  `${HOME}comma`,
  `${HOME}currency`,
  `${HOME}percent`,
  `${HOME}decimals-more`,
  `${HOME}decimals-less`,
];
/**
 * Conditional Formatting's one-click presets (data bars, colour scales, icon
 * sets) format a cell range; measured dead with a chart and a text box
 * selected. The rule dialogs (New Rule, Manage Rules, the comparison
 * rules) open a dialog and stay available.
 */
const CONDITIONAL_FORMAT_PRESETS = [
  // The three preset submenus (raw Radix sub-triggers, gated by the window
  // itself) and every preset in them.
  `${RIBBON}cf-data-bars`,
  `${RIBBON}cf-color-scales`,
  `${RIBBON}cf-icon-sets`,
  `${RIBBON}cf-dataBarBlue`,
  `${RIBBON}cf-dataBarGreen`,
  `${RIBBON}cf-dataBarLightBlue`,
  `${RIBBON}cf-dataBarOrange`,
  `${RIBBON}cf-dataBarPurple`,
  `${RIBBON}cf-dataBarRed`,
  `${RIBBON}cf-dataBarSolidBlue`,
  `${RIBBON}cf-dataBarSolidGreen`,
  `${RIBBON}cf-dataBarSolidLightBlue`,
  `${RIBBON}cf-dataBarSolidOrange`,
  `${RIBBON}cf-dataBarSolidRed`,
  `${RIBBON}cf-iconArrows`,
  `${RIBBON}cf-iconArrows4`,
  `${RIBBON}cf-iconQuarters5`,
  `${RIBBON}cf-iconRatings5`,
  `${RIBBON}cf-iconStars`,
  `${RIBBON}cf-iconTrafficLights`,
  `${RIBBON}cf-scaleGreenYellowRed`,
  `${RIBBON}cf-scaleRedYellowGreen`,
  `${RIBBON}cf-scaleWhiteBlue`,
];
const CELL_VALUE_COMMANDS = [...NUMBER_GROUP, ...CONDITIONAL_FORMAT_PRESETS];

/**
 * Commands that rewrite the selected cells' content or order: Insert and
 * Delete Cells that shift the cells, Fill Down/Right/Up/Left, Clear All /
 * Formats / Contents, Sort A to Z / Z to A and the AutoFilter toggle (Home
 * and Data), the in-cell form controls (Insert > More: option button, spin
 * button, slider) and Unlock Cells. With a chart, picture, shape, text box,
 * SmartArt, group, several shapes, ink or slicer selected every one of them
 * wrote the hidden range behind the object (measured). Excel greys them all
 * while an object is selected, Clear aside, which clears the chart's own
 * formatting there. The row and column commands (insert, delete, hide,
 * group, autofit) stay: they act on the sheet around the object and move it.
 */
const CELL_EDIT_COMMANDS = [
  // Insert Cells… / Delete Cells… (XLS-A009: Excel's dialog holds the shifts).
  `${HOME}insert-cells`,
  `${HOME}delete-cells`,
  `${HOME}fill-down`,
  `${HOME}fill-right`,
  `${HOME}fill-up`,
  `${HOME}fill-left`,
  `${HOME}clear-all`,
  `${HOME}clear-formats`,
  `${HOME}clear-contents`,
  `${HOME}sort-asc`,
  `${HOME}sort-desc`,
  `${HOME}sort-filter-column`,
  `${RIBBON}sort-asc`,
  `${RIBBON}sort-desc`,
  `${RIBBON}insert-radio`,
  `${RIBBON}insert-spinner`,
  `${RIBBON}insert-slider`,
  "spreadsheet-review-unlock-selection",
];
/**
 * AutoSum's functions (Home and Formulas): with a picture, a SmartArt shape
 * or a slicer selected they wrote the formula into the hidden cells and the
 * object stayed selected. With a chart, a SmartArt graphic or several shapes
 * selected they left the object for the cell, so the result is seen; a
 * single shape or text box shares its structure with a multi-selection here,
 * so those keep AutoSum too.
 */
const AUTOSUM_FUNCTIONS = ["sum", "average", "count", "max", "min"].flatMap(
  (fn) => [`${HOME}autosum-${fn}`, `${RIBBON}formulas-autosum-${fn}`],
);
/**
 * Home's cell formatting for a slicer: the object scope does not cover a
 * selected slicer, so the Font, Alignment, Fill, Border and Cell Style
 * commands still wrote the cells behind it (measured).
 */
const SLICER_CELL_FORMATTING = [
  `${HOME}font-family`,
  `${HOME}font-size`,
  `${HOME}font-size-options`,
  `${HOME}font-size-option-`,
  `${HOME}grow-font`,
  `${HOME}shrink-font`,
  `${HOME}bold`,
  `${HOME}italic`,
  `${HOME}underline`,
  `${HOME}underline-double`,
  `${HOME}strikethrough`,
  `${HOME}text-color`,
  `${HOME}text-color-options`,
  `${RIBBON}text-color-`,
  `${HOME}fill-color`,
  `${HOME}fill-color-options`,
  `${RIBBON}fill-color-`,
  `${HOME}borders-apply`,
  ...[
    "bottom",
    "top",
    "left",
    "right",
    "none",
    "all",
    "outside",
    "thickBox",
    "bottomDouble",
    "thickBottom",
    "topAndBottom",
    "topAndThickBottom",
    "topAndDoubleBottom",
  ].map((preset) => `${HOME}borders-${preset}`),
  `${HOME}align-top`,
  `${HOME}align-middle`,
  `${HOME}align-bottom`,
  `${HOME}align-left`,
  `${HOME}align-center`,
  `${HOME}align-right`,
  `${HOME}indent-decrease`,
  `${HOME}indent-increase`,
  `${HOME}wrap`,
  `${HOME}orientation`,
  `${HOME}orientation-`,
  `${HOME}cell-style-`,
];

/**
 * Test ids of the commands that have no meaning for each object, as
 * measured. A prefix entry ending in "-" covers a menu's every item.
 *
 * Not in this table on purpose: the Font, Alignment, Border and Fill
 * commands for a drawing object. With one selected they act on its text or
 * the object scope greys them (visible-command-authority), which decides
 * each one. A chart's text takes no strikethrough, so that one is greyed
 * for a chart (Excel strikes chart text; MRAI's chart text has no strike).
 */
const GATED: Record<SpreadsheetSelectedStructure, readonly string[]> = {
  chart: [...CELL_VALUE_COMMANDS, ...CELL_EDIT_COMMANDS, `${HOME}strikethrough`],
  picture: [...CELL_VALUE_COMMANDS, ...CELL_EDIT_COMMANDS, ...AUTOSUM_FUNCTIONS],
  slicer: [
    ...CELL_VALUE_COMMANDS,
    ...SLICER_CELL_FORMATTING,
    ...CELL_EDIT_COMMANDS,
    ...AUTOSUM_FUNCTIONS,
  ],
  shape: [...CELL_VALUE_COMMANDS, ...CELL_EDIT_COMMANDS],
  textBox: [...CELL_VALUE_COMMANDS, ...CELL_EDIT_COMMANDS],
  smartArtShape: [
    ...CELL_VALUE_COMMANDS,
    ...CELL_EDIT_COMMANDS,
    ...AUTOSUM_FUNCTIONS,
  ],
  smartArtGraphic: [...CELL_VALUE_COMMANDS, ...CELL_EDIT_COMMANDS],
};

/** The gate for `structure`; null for cells, where every command applies. */
export function spreadsheetRibbonSelectionGate(
  structure: SpreadsheetSelectedStructure | null,
  translate: (key: string) => string,
): OfficeRibbonCommandGate | null {
  if (!structure) return null;
  return selectionGateFromEntries(
    GATED[structure],
    translate(STRUCTURE_REASON_KEY[structure]),
  );
}

export { HOME as SPREADSHEET_HOME_TEST_ID_PREFIX };
