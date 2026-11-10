/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// XLS-A106 — Excel's enhanced ScreenTips for every spreadsheet ribbon command.
//
// Excel shows, after a short hover, the command's name and shortcut in bold,
// then a sentence on what it does. The sentences are many (one per command,
// eighteen languages), so they are not message-catalog keys: each language is
// a separate data module loaded on demand, the way the function help is
// (`function-docs-locale.ts`). A session downloads only the language it shows;
// a language without a module, or a key its module lacks, reads English.
//
// Gallery tiles (table / chart / picture / shape / WordArt styles, the shape
// quick-insert tiles) keep their plain name, as Excel's gallery items do.
// ---------------------------------------------------------------------------

import type { SCREEN_TIPS } from "./ribbon-screen-tips/en";
import {
  OFFICE_SCREEN_TIP_LOCALES,
  officeScreenTipLocale,
  type OfficeScreenTipLocale,
} from "@mrnote/core/web/lib/office-screen-tips/locale.ts";

export type SpreadsheetScreenTipKey = keyof typeof SCREEN_TIPS;
/** A language module: one sentence per ScreenTip key. */
export type SpreadsheetScreenTipTable = Readonly<Record<SpreadsheetScreenTipKey, string>>;

export const SPREADSHEET_SCREEN_TIP_LOCALES = OFFICE_SCREEN_TIP_LOCALES;
type ScreenTipLocale = OfficeScreenTipLocale;

const LOADERS: Record<ScreenTipLocale, () => Promise<{ SCREEN_TIPS: SpreadsheetScreenTipTable }>> = {
  da: () => import("./ribbon-screen-tips/da"),
  de: () => import("./ribbon-screen-tips/de"),
  en: () => import("./ribbon-screen-tips/en"),
  es: () => import("./ribbon-screen-tips/es"),
  fi: () => import("./ribbon-screen-tips/fi"),
  fr: () => import("./ribbon-screen-tips/fr"),
  id: () => import("./ribbon-screen-tips/id"),
  it: () => import("./ribbon-screen-tips/it"),
  ja: () => import("./ribbon-screen-tips/ja"),
  ko: () => import("./ribbon-screen-tips/ko"),
  nl: () => import("./ribbon-screen-tips/nl"),
  no: () => import("./ribbon-screen-tips/no"),
  "pt-BR": () => import("./ribbon-screen-tips/pt-BR"),
  sv: () => import("./ribbon-screen-tips/sv"),
  th: () => import("./ribbon-screen-tips/th"),
  vi: () => import("./ribbon-screen-tips/vi"),
  zh: () => import("./ribbon-screen-tips/zh"),
  "zh-TW": () => import("./ribbon-screen-tips/zh-TW"),
};

/** The module that serves a UI locale (en-GB → en, es-419 → es, zh-CN → zh). */
export const spreadsheetScreenTipLocale = officeScreenTipLocale;

const cache = new Map<ScreenTipLocale, Promise<SpreadsheetScreenTipTable>>();

/** Load (once per language) the ScreenTip sentences; English fills any gap. */
export function loadSpreadsheetScreenTips(locale: string): Promise<SpreadsheetScreenTipTable> {
  const key = spreadsheetScreenTipLocale(locale);
  let pending = cache.get(key);
  if (!pending) {
    const english = key === "en" ? null : LOADERS.en().then((module) => module.SCREEN_TIPS);
    pending = LOADERS[key]()
      .then(async (module) => (english ? { ...(await english), ...module.SCREEN_TIPS } : module.SCREEN_TIPS))
      .catch(() => (english ?? LOADERS.en().then((module) => module.SCREEN_TIPS)));
    cache.set(key, pending);
  }
  return pending;
}

/** Ribbon controls by their test id. Split parts and duplicates share a key. */
const BY_TEST_ID: Readonly<Record<string, SpreadsheetScreenTipKey>> = {
  "spreadsheet-undo": "undo",
  "spreadsheet-redo": "redo",
  "spreadsheet-ribbon-home-paste": "paste",
  "spreadsheet-ribbon-home-cut": "cut",
  "spreadsheet-ribbon-home-copy": "copy",
  "spreadsheet-ribbon-home-format-painter": "formatPainter",
  "spreadsheet-ribbon-home-font-family": "fontFamily",
  "spreadsheet-ribbon-home-font-size": "fontSize",
  "spreadsheet-ribbon-home-grow-font": "growFont",
  "spreadsheet-ribbon-home-shrink-font": "shrinkFont",
  "spreadsheet-ribbon-home-bold": "bold",
  "spreadsheet-ribbon-home-italic": "italic",
  "spreadsheet-ribbon-home-underline": "underline",
  "spreadsheet-ribbon-home-underline-double": "underlineDouble",
  "spreadsheet-ribbon-home-strikethrough": "strikethrough",
  "spreadsheet-ribbon-home-borders-apply": "borders",
  "spreadsheet-ribbon-home-borders": "borders",
  "spreadsheet-ribbon-home-fill-color": "fillColor",
  "spreadsheet-ribbon-home-text-color": "fontColor",
  "spreadsheet-ribbon-home-phonetic": "phonetic",
  "spreadsheet-ribbon-home-font-launcher": "formatCellsFont",
  "spreadsheet-ribbon-home-align-top": "alignTop",
  "spreadsheet-ribbon-home-align-middle": "alignMiddle",
  "spreadsheet-ribbon-home-align-bottom": "alignBottom",
  "spreadsheet-ribbon-home-align-left": "alignLeft",
  "spreadsheet-ribbon-home-align-center": "alignCenter",
  "spreadsheet-ribbon-home-align-right": "alignRight",
  "spreadsheet-ribbon-home-indent-decrease": "indentDecrease",
  "spreadsheet-ribbon-home-indent-increase": "indentIncrease",
  "spreadsheet-ribbon-home-orientation": "orientation",
  "spreadsheet-ribbon-home-wrap": "wrapText",
  "spreadsheet-ribbon-home-merge": "mergeCenter",
  "spreadsheet-ribbon-home-alignment-launcher": "formatCellsAlignment",
  "spreadsheet-ribbon-home-number-format": "numberFormat",
  "spreadsheet-ribbon-home-currency": "accounting",
  "spreadsheet-ribbon-home-accounting-options": "accounting",
  "spreadsheet-ribbon-home-percent": "percent",
  "spreadsheet-ribbon-home-comma": "comma",
  "spreadsheet-ribbon-home-decimals-more": "decimalsMore",
  "spreadsheet-ribbon-home-decimals-less": "decimalsLess",
  "spreadsheet-ribbon-home-number-launcher": "formatCellsNumber",
  "spreadsheet-ribbon-conditional-format": "conditionalFormatting",
  "spreadsheet-ribbon-home-format-table": "formatAsTable",
  "spreadsheet-ribbon-home-cell-styles": "cellStyles",
  "spreadsheet-ribbon-home-insert-menu": "insertCells",
  "spreadsheet-ribbon-home-delete-menu": "deleteCells",
  "spreadsheet-ribbon-home-cells-format": "formatRowsColumns",
  "spreadsheet-ribbon-home-autosum": "autoSum",
  "spreadsheet-ribbon-home-fill": "fill",
  "spreadsheet-ribbon-home-clear": "clear",
  "spreadsheet-ribbon-home-sort": "sortFilter",
  "spreadsheet-ribbon-home-find": "findSelect",
  // Insert
  "spreadsheet-ribbon-insert-pivot": "pivotTable",
  "spreadsheet-ribbon-insert-table": "insertTable",
  "spreadsheet-ribbon-insert-picture": "pictures",
  "spreadsheet-ribbon-insert-smartart": "smartArt",
  "spreadsheet-ribbon-insert-shape": "shapes",
  "spreadsheet-ribbon-insert-screenshot": "screenshot",
  "spreadsheet-ribbon-insert-checkbox": "checkbox",
  "spreadsheet-ribbon-insert-recommended-chart": "recommendedCharts",
  "spreadsheet-ribbon-insert-chart-column": "chartColumn",
  "spreadsheet-ribbon-insert-chart-line": "chartLine",
  "spreadsheet-ribbon-insert-chart-scatter": "chartScatter",
  "spreadsheet-ribbon-insert-chart-pie": "chartPie",
  "spreadsheet-ribbon-insert-chart-hierarchy": "chartHierarchy",
  "spreadsheet-ribbon-insert-chart-statistical": "chartStatistical",
  "spreadsheet-ribbon-insert-chart-combo": "chartCombo",
  "spreadsheet-ribbon-insert-chart-other": "chartOther",
  "spreadsheet-ribbon-insert-chart-map": "chartMap",
  "spreadsheet-ribbon-insert-pivot-chart": "pivotChart",
  "spreadsheet-ribbon-insert-sparkline-line": "sparklineLine",
  "spreadsheet-ribbon-insert-sparkline-column": "sparklineColumn",
  "spreadsheet-ribbon-insert-sparkline-winloss": "sparklineWinLoss",
  "spreadsheet-ribbon-insert-slicer": "slicer",
  "spreadsheet-ribbon-insert-timeline": "timeline",
  "spreadsheet-ribbon-insert-hyperlink": "hyperlink",
  "spreadsheet-ribbon-insert-comment": "newComment",
  "spreadsheet-ribbon-insert-text-box": "textBox",
  "spreadsheet-ribbon-insert-header-footer": "headerFooter",
  "spreadsheet-ribbon-insert-wordart": "wordArt",
  "spreadsheet-ribbon-insert-more": "insertMore",
  "spreadsheet-ribbon-insert-equation": "equation",
  "spreadsheet-ribbon-insert-symbol": "symbol",
  // Draw
  "spreadsheet-ribbon-draw-mode": "drawMode",
  "spreadsheet-ribbon-draw-eraser": "eraser",
  "spreadsheet-ribbon-draw-lasso": "lasso",
  "spreadsheet-ribbon-draw-add-pen": "addPen",
  "spreadsheet-ribbon-draw-trackpad": "trackpad",
  "spreadsheet-ribbon-draw-ink-to-shape": "inkToShape",
  "spreadsheet-ribbon-draw-ink-to-math": "inkToMath",
  "spreadsheet-ribbon-draw-ink-replay": "inkReplay",
  // Page Layout
  "spreadsheet-ribbon-page-layout-themes": "themes",
  "spreadsheet-ribbon-page-layout-colors": "themeColors",
  "spreadsheet-ribbon-page-layout-fonts": "themeFonts",
  "spreadsheet-ribbon-page-layout-effects": "themeEffects",
  "spreadsheet-ribbon-page-layout-margins": "margins",
  "spreadsheet-ribbon-page-layout-orientation": "pageOrientation",
  "spreadsheet-ribbon-page-layout-size": "pageSize",
  "spreadsheet-ribbon-page-layout-print-area": "printArea",
  "spreadsheet-ribbon-page-breaks": "breaks",
  "spreadsheet-ribbon-page-layout-background": "background",
  "spreadsheet-ribbon-page-layout-print-titles": "printTitles",
  "spreadsheet-ribbon-page-setup": "pageSetup",
  "spreadsheet-ribbon-page-layout-scale-width": "scaleWidth",
  "spreadsheet-ribbon-page-layout-scale-height": "scaleHeight",
  "spreadsheet-ribbon-page-layout-gridlines-view": "gridlinesView",
  "spreadsheet-ribbon-page-layout-headings-view": "headingsView",
  "spreadsheet-ribbon-page-layout-gridlines-print": "gridlinesPrint",
  "spreadsheet-ribbon-page-layout-headings-print": "headingsPrint",
  "spreadsheet-ribbon-arrange-front": "bringToFront",
  "spreadsheet-ribbon-arrange-back": "sendToBack",
  "spreadsheet-ribbon-arrange-forward": "bringForward",
  "spreadsheet-ribbon-arrange-backward": "sendBackward",
  "spreadsheet-ribbon-page-arrange-selection-pane": "selectionPane",
  "spreadsheet-ribbon-page-arrange-align": "align",
  "spreadsheet-ribbon-page-arrange-group": "group",
  "spreadsheet-ribbon-page-arrange-rotate": "rotate",
  "spreadsheet-ribbon-picture-arrange-selection-pane": "selectionPane",
  "spreadsheet-ribbon-picture-arrange-align": "align",
  "spreadsheet-ribbon-picture-arrange-group": "group",
  "spreadsheet-ribbon-picture-arrange-rotate": "rotate",
  // Formulas
  "spreadsheet-ribbon-insert-function": "insertFunction",
  "spreadsheet-ribbon-autosum": "autoSum",
  "spreadsheet-ribbon-function-recent": "recentlyUsed",
  "spreadsheet-ribbon-function-financial": "financial",
  "spreadsheet-ribbon-function-logical": "logical",
  "spreadsheet-ribbon-function-text": "textFunctions",
  "spreadsheet-ribbon-function-dateTime": "dateTime",
  "spreadsheet-ribbon-function-lookup": "lookupReference",
  "spreadsheet-ribbon-function-math": "mathTrig",
  "spreadsheet-ribbon-more-functions": "moreFunctions",
  "spreadsheet-ribbon-name-manager": "nameManager",
  "spreadsheet-ribbon-formulas-names": "defineName",
  "spreadsheet-ribbon-use-in-formula": "useInFormula",
  "spreadsheet-ribbon-create-from-selection": "createFromSelection",
  "spreadsheet-ribbon-show-formulas": "showFormulas",
  "spreadsheet-ribbon-trace-precedents": "tracePrecedents",
  "spreadsheet-ribbon-trace-dependents": "traceDependents",
  "spreadsheet-ribbon-remove-arrows": "removeArrows",
  "spreadsheet-ribbon-error-check": "errorChecking",
  "spreadsheet-ribbon-evaluate-formula": "evaluateFormula",
  "spreadsheet-ribbon-formulas-watch": "watchWindow",
  "spreadsheet-ribbon-calculation-options": "calculationOptions",
  "spreadsheet-ribbon-calculate-workbook": "calculateNow",
  "spreadsheet-ribbon-calculate-sheet": "calculateSheet",
  // Data
  "spreadsheet-ribbon-data-get-data": "getData",
  "spreadsheet-ribbon-data-from-table-range": "fromTableRange",
  "spreadsheet-ribbon-data-from-picture": "fromPicture",
  "spreadsheet-ribbon-data-refresh": "refresh",
  "spreadsheet-ribbon-data-refresh-all": "refreshAll",
  "spreadsheet-ribbon-data-queries-connections": "queriesConnections",
  "spreadsheet-ribbon-data-connection-properties": "connectionProperties",
  "spreadsheet-ribbon-data-workbook-links": "workbookLinks",
  "spreadsheet-ribbon-sort-asc": "sortAscending",
  "spreadsheet-ribbon-sort-desc": "sortDescending",
  "spreadsheet-ribbon-sort-dialog": "customSort",
  "spreadsheet-ribbon-filter-column": "filter",
  "spreadsheet-ribbon-clear-filters": "clearFilter",
  "spreadsheet-ribbon-reapply-filters": "reapply",
  "spreadsheet-ribbon-advanced-filter": "advancedFilter",
  "spreadsheet-ribbon-data-text-to-columns": "textToColumns",
  "spreadsheet-ribbon-data-flash-fill": "flashFill",
  "spreadsheet-ribbon-data-remove-duplicates": "removeDuplicates",
  "spreadsheet-ribbon-data-consolidate": "consolidate",
  "spreadsheet-ribbon-validation": "dataValidation",
  "spreadsheet-ribbon-data-analyze": "analyzeData",
  "spreadsheet-ribbon-data-what-if": "whatIf",
  "spreadsheet-ribbon-data-forecast-sheet": "forecastSheet",
  "spreadsheet-ribbon-data-group": "outlineGroup",
  "spreadsheet-ribbon-data-ungroup": "outlineUngroup",
  "spreadsheet-ribbon-data-subtotal": "subtotal",
  "spreadsheet-ribbon-data-show-detail": "showDetail",
  "spreadsheet-ribbon-data-hide-detail": "hideDetail",
  // Review
  "spreadsheet-ribbon-review-spelling": "spelling",
  "spreadsheet-ribbon-review-performance": "checkPerformance",
  "spreadsheet-ribbon-review-statistics": "workbookStatistics",
  "spreadsheet-ribbon-review-accessibility": "accessibility",
  "spreadsheet-ribbon-comment-new": "newComment",
  "spreadsheet-ribbon-comment-show": "showComments",
  "spreadsheet-ribbon-comment-previous": "previousComment",
  "spreadsheet-ribbon-comment-next": "nextComment",
  "spreadsheet-ribbon-comment-delete": "deleteComment",
  "spreadsheet-ribbon-review-notes": "notes",
  "spreadsheet-ribbon-review-protect-sheet": "protectSheet",
  "spreadsheet-ribbon-review-protect-workbook": "protectWorkbook",
  "spreadsheet-ribbon-review-allow-edit-ranges": "allowEditRanges",
  "spreadsheet-ribbon-review-manage-protection": "manageProtection",
  "spreadsheet-ribbon-review-pause-protection": "pauseProtection",
  "spreadsheet-ribbon-review-show-changes": "showChanges",
  "spreadsheet-ribbon-review-version-history": "versionHistory",
  // View
  "spreadsheet-ribbon-view-sheet-view": "sheetView",
  "spreadsheet-view-sheet-view-keep": "keepSheetView",
  "spreadsheet-view-sheet-view-new": "newSheetView",
  "spreadsheet-view-sheet-view-options": "sheetViewOptions",
  "spreadsheet-view-sheet-view-exit": "exitSheetView",
  "spreadsheet-ribbon-view-normal": "normalView",
  "spreadsheet-ribbon-view-page-break": "pageBreakPreview",
  "spreadsheet-ribbon-view-page-layout": "pageLayoutView",
  "spreadsheet-ribbon-view-custom-views": "customViews",
  "spreadsheet-ribbon-view-navigation": "navigation",
  "spreadsheet-ribbon-headings": "headingsView",
  "spreadsheet-ribbon-gridlines": "gridlinesView",
  "spreadsheet-ribbon-formula-bar": "formulaBar",
  "spreadsheet-ribbon-view-ruler": "ruler",
  "spreadsheet-ribbon-view-focus-cell": "focusCell",
  "spreadsheet-ribbon-view-zoom": "zoom",
  "spreadsheet-ribbon-view-zoom-100": "zoom100",
  "spreadsheet-ribbon-view-zoom-selection": "zoomSelection",
  "spreadsheet-ribbon-view-new-window": "newWindow",
  "spreadsheet-ribbon-view-arrange-all": "arrangeAll",
  "spreadsheet-ribbon-freeze-panes": "freezePanes",
  "spreadsheet-ribbon-view-split": "split",
  "spreadsheet-ribbon-view-hide-window": "hideWindow",
  "spreadsheet-ribbon-view-unhide-window": "unhideWindow",
  "spreadsheet-ribbon-view-side-by-side": "sideBySide",
  "spreadsheet-ribbon-view-sync-scrolling": "syncScrolling",
  "spreadsheet-ribbon-view-reset-window-position": "resetWindowPosition",
  "spreadsheet-ribbon-view-switch-windows": "switchWindows",
  "spreadsheet-ribbon-view-full-screen": "fullScreen",
  "spreadsheet-ribbon-view-immersive-reader": "immersiveReader",
  "spreadsheet-ribbon-view-macros": "macros",
  // Automate / Share
  "spreadsheet-ribbon-automate-new-script": "newScript",
  "spreadsheet-ribbon-automate-view-scripts": "viewScripts",
  "spreadsheet-ribbon-automate-record-actions": "recordActions",
  "spreadsheet-ribbon-automate-unhide-all": "unhideAllRowsColumns",
  "spreadsheet-ribbon-automate-freeze-selection": "freezeSelection",
  "spreadsheet-ribbon-automate-clear-hyperlinks": "clearHyperlinks",
  "spreadsheet-ribbon-automate-count-blank-rows": "countBlankRows",
  "spreadsheet-ribbon-share-workbook": "shareWorkbook",
  "spreadsheet-ribbon-share-copy-link": "copyLink",
  "spreadsheet-ribbon-share-worksheet-link": "worksheetLink",
  // Table Design
  "spreadsheet-ribbon-table-name": "tableName",
  "spreadsheet-ribbon-table-alt-text": "tableAltText",
  "spreadsheet-ribbon-table-resize": "resizeTable",
  "spreadsheet-ribbon-table-insert-slicer": "slicer",
  "spreadsheet-ribbon-table-summarize-pivot": "summarizePivot",
  "spreadsheet-ribbon-table-remove-duplicates": "removeDuplicates",
  "spreadsheet-ribbon-table-convert-range": "convertToRange",
  "spreadsheet-ribbon-table-refresh-all": "refreshAll",
  "spreadsheet-ribbon-table-header-row": "headerRow",
  "spreadsheet-ribbon-table-total-row": "totalRow",
  "spreadsheet-ribbon-table-first-column": "firstColumn",
  "spreadsheet-ribbon-table-last-column": "lastColumn",
  "spreadsheet-ribbon-table-banded-rows": "bandedRows",
  "spreadsheet-ribbon-table-banded-columns": "bandedColumns",
  "spreadsheet-ribbon-table-filter-button": "filterButton",
  "spreadsheet-ribbon-table-column-agg": "columnTotal",
  // Chart Design
  "spreadsheet-ribbon-chart-design-add-element": "addChartElement",
  "spreadsheet-chart-quick-layout-trigger": "quickLayout",
  "spreadsheet-ribbon-chart-change-colors": "changeColors",
  "spreadsheet-ribbon-chart-switch-row-column": "switchRowColumn",
  "spreadsheet-ribbon-chart-select-data": "selectData",
  "spreadsheet-ribbon-chart-hidden-empty-cells": "hiddenEmptyCells",
  "spreadsheet-ribbon-chart-change-type": "changeChartType",
  "spreadsheet-ribbon-chart-move": "moveChart",
  // Chart Format
  "spreadsheet-chart-format-ribbon-selection": "chartElements",
  "spreadsheet-chart-format-ribbon-reset-style": "resetToMatchStyle",
  "spreadsheet-chart-format-ribbon-insert-shapes": "shapes",
  "spreadsheet-chart-format-ribbon-change-shape": "editShape",
  "spreadsheet-chart-format-ribbon-fill": "shapeFill",
  "spreadsheet-chart-format-ribbon-outline": "shapeOutline",
  "spreadsheet-chart-format-ribbon-effects": "shapeEffects",
  "spreadsheet-chart-format-ribbon-text-fill": "textFill",
  "spreadsheet-chart-format-ribbon-text-outline": "textOutline",
  "spreadsheet-chart-format-ribbon-text-effects": "textEffects",
  "spreadsheet-chart-arrange-front": "bringForward",
  "spreadsheet-chart-arrange-back": "sendBackward",
  "spreadsheet-chart-selection-pane": "selectionPane",
  "spreadsheet-chart-align": "align",
  "spreadsheet-chart-group": "group",
  "spreadsheet-chart-rotate": "rotate",
  "spreadsheet-chart-format-ribbon-height": "objectHeight",
  "spreadsheet-chart-format-ribbon-width": "objectWidth",
  "spreadsheet-chart-format-ribbon-pane": "formatPane",
  "spreadsheet-ribbon-chart-series-options": "formatSelection",
  // Picture Format
  "spreadsheet-ribbon-picture-artistic-effects": "artisticEffects",
  "spreadsheet-ribbon-picture-alt-command": "altText",
  "spreadsheet-ribbon-picture-crop": "crop",
  "spreadsheet-ribbon-picture-crop-fill": "cropFill",
  "spreadsheet-ribbon-picture-crop-fit": "cropFit",
  "spreadsheet-ribbon-picture-size-height": "objectHeight",
  "spreadsheet-ribbon-picture-size-width": "objectWidth",
  // Shape Format
  "spreadsheet-ribbon-shape-insert-more": "shapes",
  "spreadsheet-ribbon-shape-change": "editShape",
  "spreadsheet-ribbon-shape-text-box": "textBox",
  "spreadsheet-ribbon-object-fill-color": "shapeFill",
  "spreadsheet-ribbon-object-stroke-color": "shapeOutline",
  "spreadsheet-ribbon-shape-effects": "shapeEffects",
  "spreadsheet-ribbon-shape-text-fill": "textFill",
  "spreadsheet-ribbon-shape-text-outline": "textOutline",
  "spreadsheet-ribbon-shape-text-effects": "textEffects",
  "spreadsheet-ribbon-shape-alt": "altText",
  "spreadsheet-ribbon-shape-selection-pane": "selectionPane",
  "spreadsheet-ribbon-shape-align": "align",
  "spreadsheet-ribbon-shape-group": "group",
  "spreadsheet-ribbon-shape-rotate": "rotate",
  "spreadsheet-ribbon-shape-size-width": "objectWidth",
  "spreadsheet-ribbon-shape-size-height": "objectHeight",
  "spreadsheet-ribbon-shape-lock-aspect": "lockAspect",
  "spreadsheet-ribbon-shape-format-pane": "formatPane",
  // Sparkline
  "spreadsheet-ribbon-sparkline-type-line": "sparklineTypeLine",
  "spreadsheet-ribbon-sparkline-type-column": "sparklineTypeColumn",
  "spreadsheet-ribbon-sparkline-type-winloss": "sparklineTypeWinLoss",
  "spreadsheet-ribbon-sparkline-markers": "markers",
  "spreadsheet-ribbon-sparkline-high": "highPoint",
  "spreadsheet-ribbon-sparkline-low": "lowPoint",
  "spreadsheet-ribbon-sparkline-first": "firstPoint",
  "spreadsheet-ribbon-sparkline-last": "lastPoint",
  "spreadsheet-ribbon-sparkline-negative": "negativePoints",
  "spreadsheet-ribbon-sparkline-show-axis": "sparklineAxis",
  "spreadsheet-ribbon-sparkline-min-axis-type": "sparklineMinAxis",
  "spreadsheet-ribbon-sparkline-max-axis-type": "sparklineMaxAxis",
  "spreadsheet-ribbon-sparkline-range": "sparklineData",
  "spreadsheet-ribbon-sparkline-clear": "sparklineClear",
  "spreadsheet-ribbon-sparkline-empty-cells": "sparklineEmptyCells",
  "spreadsheet-ribbon-sparkline-display-hidden": "sparklineHiddenData",
  "spreadsheet-ribbon-sparkline-primary-color": "sparklineColor",
  "spreadsheet-ribbon-sparkline-negative-color": "sparklineNegativeColor",
  "spreadsheet-ribbon-sparkline-marker-color": "sparklineMarkerColor",
  "spreadsheet-ribbon-sparkline-group": "sparklineGroup",
  "spreadsheet-ribbon-sparkline-ungroup": "sparklineUngroup",
  // Slicer
  "spreadsheet-ribbon-slicer-title": "slicerCaption",
  "spreadsheet-ribbon-slicer-show-header": "slicerHeader",
  "spreadsheet-ribbon-slicer-columns": "slicerColumns",
  "spreadsheet-ribbon-slicer-buttonHeight": "slicerButtonHeight",
  "spreadsheet-ribbon-slicer-width": "objectWidth",
  "spreadsheet-ribbon-slicer-height": "objectHeight",
  "spreadsheet-ribbon-slicer-clear": "slicerClearFilter",
  "spreadsheet-ribbon-slicer-remove": "slicerRemove",
  // PivotTable / PivotChart
  "spreadsheet-ribbon-pivot-refresh": "pivotRefresh",
  "spreadsheet-ribbon-pivot-source": "changeDataSource",
  "spreadsheet-ribbon-pivot-select": "pivotSelect",
  "spreadsheet-ribbon-pivot-clear": "pivotClear",
  "spreadsheet-ribbon-pivot-move": "movePivot",
  "spreadsheet-ribbon-pivot-chart": "pivotChartPreview",
  "spreadsheet-ribbon-pivot-chart-create": "pivotChart",
  "spreadsheet-ribbon-pivot-insert-slicer": "slicer",
  "spreadsheet-ribbon-pivot-fields": "fieldList",
  "spreadsheet-ribbon-pivot-expand-collapse-buttons": "expandCollapseButtons",
  "spreadsheet-ribbon-pivot-design-subtotals": "subtotals",
  "spreadsheet-ribbon-pivot-design-grand-totals": "grandTotals",
  "spreadsheet-ribbon-pivot-design-report-layout": "reportLayout",
  "spreadsheet-ribbon-pivot-design-blank-rows": "blankRows",
  "spreadsheet-ribbon-pivot-design-showRowHeaders": "rowHeaders",
  "spreadsheet-ribbon-pivot-design-showColumnHeaders": "columnHeaders",
  "spreadsheet-ribbon-pivot-design-bandedRows": "bandedRows",
  "spreadsheet-ribbon-pivot-design-bandedColumns": "bandedColumns",
  "spreadsheet-ribbon-pivotchart-expand-field": "expandField",
  "spreadsheet-ribbon-pivotchart-collapse-field": "collapseField",
  "spreadsheet-ribbon-pivotchart-fields": "fieldList",
  "spreadsheet-ribbon-pivotchart-field-buttons": "fieldButtons",
  // SmartArt Design (shared OfficeSmartArtDesignTab, "spreadsheet" prefix)
  "spreadsheet-ribbon-smartart-textPane": "smartArtTextPane",
  "spreadsheet-ribbon-smartart-addShape": "smartArtAddShape",
  "spreadsheet-ribbon-smartart-add-shape": "smartArtAddShape",
  "spreadsheet-ribbon-smartart-add-shape-split": "smartArtAddShape",
  "spreadsheet-ribbon-smartart-promote": "smartArtPromote",
  "spreadsheet-ribbon-smartart-demote": "smartArtDemote",
  "spreadsheet-ribbon-smartart-right-to-left": "smartArtRightToLeft",
  "spreadsheet-ribbon-smartart-add-bullet": "smartArtAddBullet",
  "spreadsheet-ribbon-smartart-move-up": "smartArtMoveUp",
  "spreadsheet-ribbon-smartart-move-down": "smartArtMoveDown",
  "spreadsheet-ribbon-smartart-picture": "smartArtPicture",
  "spreadsheet-ribbon-smartart-colors": "smartArtColors",
  "spreadsheet-ribbon-smartart-larger": "smartArtLarger",
  "spreadsheet-ribbon-smartart-smaller": "smartArtSmaller",
  "spreadsheet-ribbon-smartart-reset": "smartArtReset",
  "spreadsheet-ribbon-smartart-convert": "smartArtConvert",
  "spreadsheet-smartart-convert-native": "smartArtConvert",
  "spreadsheet-ribbon-smartart-alt-text": "altText",
};

/** Gallery tiles keep their plain name, as Excel's gallery items do. */
const PLAIN_GALLERY_TILE = /-(gallery|styles|text-styles|wordart)-inline-|-insert-shapes-quick-|-shape-insert-quick-/;

/**
 * The ScreenTip key of a ribbon control's test id, or undefined when the
 * control keeps a plain tooltip. A split button's drop-down part (`-options`
 * / `-menu`) shares its main part's tip.
 */
export function spreadsheetScreenTipKey(testId: string): SpreadsheetScreenTipKey | undefined {
  if (PLAIN_GALLERY_TILE.test(testId)) return undefined;
  const exact = BY_TEST_ID[testId];
  if (exact) return exact;
  if (/-(gallery|styles|wordart|text-styles)-more$/.test(testId)) return "galleryMore";
  if (testId.startsWith("spreadsheet-ribbon-draw-ink-pen-")) return "pen";
  if (testId.startsWith("spreadsheet-ribbon-draw-ink-pencil-")) return "pencil";
  if (testId.startsWith("spreadsheet-ribbon-draw-ink-highlighter-")) return "highlighter";
  const part = /^(.*)-(options|menu)$/.exec(testId);
  return part ? BY_TEST_ID[part[1]!] : undefined;
}

/**
 * Keyboard shortcuts MRAI really binds (lib/spreadsheet/command-registry.ts,
 * calculation-mode.ts, the window's F7 handler), by ScreenTip key. `Mod` is
 * Ctrl, or ⌘ on a Mac. A control whose own title already names its shortcut
 * (Undo, AutoSum, Show Formulas, Flash Fill) is left as it is.
 */
const SHORTCUTS: Partial<Record<SpreadsheetScreenTipKey, string>> = {
  paste: "Mod+V",
  cut: "Mod+X",
  copy: "Mod+C",
  bold: "Mod+B",
  italic: "Mod+I",
  underline: "Mod+U",
  strikethrough: "Mod+5",
  formatCellsFont: "Mod+1",
  formatCellsAlignment: "Mod+1",
  formatCellsNumber: "Mod+1",
  percent: "Mod+Shift+%",
  insertTable: "Mod+T",
  formatAsTable: "Mod+T",
  filter: "Mod+Shift+L",
  hyperlink: "Mod+K",
  insertCells: "Mod+Shift+=",
  deleteCells: "Mod+-",
  calculateNow: "F9",
  calculateSheet: "Shift+F9",
  spelling: "F7",
};

/** Whether a test id has a ScreenTip of its own (not a split button's
 * drop-down part borrowing its main part's). */
export function spreadsheetScreenTipIsOwnControl(testId: string): boolean {
  return Object.prototype.hasOwnProperty.call(BY_TEST_ID, testId);
}

/** The shortcut of a ScreenTip key for this platform, e.g. "Ctrl+B" / "⌘B". */
export function spreadsheetScreenTipShortcut(
  key: SpreadsheetScreenTipKey,
  mac: boolean,
): string | undefined {
  const raw = SHORTCUTS[key];
  if (!raw) return undefined;
  if (!mac) return raw.replace(/^Mod\+/, "Ctrl+");
  return raw.replace(/^Mod\+Shift\+/, "⇧⌘").replace(/^Mod\+/, "⌘").replace(/^Shift\+/, "⇧");
}
