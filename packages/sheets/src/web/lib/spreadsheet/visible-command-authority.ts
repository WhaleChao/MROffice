/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  checkEnabledFeatureCommandRegistry,
  createFeatureCommandBindingContract,
  createFeatureCommandRegistry,
  createFeatureCommandRegistryLineage,
  type FeatureCommandHandler,
  type FeatureCommandRegistryLineage,
} from "@mrnote/core/web/components/office-ribbon/featureCommandRegistry.ts";
import {
  isSpreadsheetFeatureExecutable,
  readSpreadsheetFeatureCapabilityCatalog,
  resolveSpreadsheetFeatureCapability,
  SPREADSHEET_CAPABILITY_CATALOG_SCHEMA_VERSION,
  spreadsheetFeatureCommandSemanticContract,
  verifySpreadsheetFeatureCapabilityCatalogBindings,
  type SpreadsheetCapabilityControlKind,
  type SpreadsheetFeatureCapabilityContract,
  type SpreadsheetNativeWriteRoute,
} from "./feature-capabilities";
import type { SheetProtectionPermissions } from "@mrnote/core/web/lib/spreadsheet/protection.ts";

export type SpreadsheetVisibleCommandReason =
  | "spreadsheet.capability.activeCommentRequired"
  | "spreadsheet.capability.activeNoteRequired"
  | "spreadsheet.capability.activeFilterRequired"
  | "spreadsheet.capability.actionPenUnavailable"
  | "spreadsheet.capability.analyzeDataUnavailable"
  | "spreadsheet.capability.clipboardUnavailable"
  | "spreadsheet.capability.pasteSourceRequired"
  | "spreadsheet.capability.cellShiftUnsafe"
  | "spreadsheet.capability.commentAuthorUnavailable"
  | "spreadsheet.capability.commentKindConflict"
  | "spreadsheet.capability.commandUnavailable"
  | "spreadsheet.capability.objectSelected"
  | "spreadsheet.capability.mergeInTable"
  | "spreadsheet.capability.printAreaAlreadySelection"
  | "spreadsheet.capability.printAreaNone"
  | "spreadsheet.capability.pageBreakInsertUnavailable"
  | "spreadsheet.capability.printAreaAddOutsideRequired"
  | "spreadsheet.capability.pageBreakNoneAtSelection"
  | "spreadsheet.capability.pageBreaksNone"
  | "spreadsheet.capability.dataTypesUnavailable"
  | "spreadsheet.capability.dataFormListRequired"
  | "spreadsheet.capability.fullscreenUnavailable"
  | "spreadsheet.capability.immersiveReaderUnavailable"
  | "spreadsheet.capability.importInProgress"
  | "spreadsheet.capability.multiColumnSelectionRequired"
  | "spreadsheet.capability.multiRowSelectionRequired"
  | "spreadsheet.capability.multiSelectionRequired"
  | "spreadsheet.capability.noComments"
  | "spreadsheet.capability.noNotes"
  | "spreadsheet.capability.manualCalculationUnavailable"
  | "spreadsheet.capability.noQueries"
  | "spreadsheet.capability.noWorkbookLinks"
  | "spreadsheet.capability.notesUnavailable"
  | "spreadsheet.capability.namedViewsUnavailable"
  | "spreadsheet.capability.navigationUnavailable"
  | "spreadsheet.capability.outlineGroupUnavailable"
  | "spreadsheet.capability.outlineNativeBlocked"
  | "spreadsheet.capability.outlineUngroupUnavailable"
  | "spreadsheet.capability.pageWriteUnavailable"
  | "spreadsheet.capability.changeHistoryUnavailable"
  | "spreadsheet.capability.formsUnavailable"
  | "spreadsheet.capability.powerBiUnavailable"
  | "spreadsheet.capability.pictureInCellUnavailable"
  | "spreadsheet.capability.pictureLibraryUnavailable"
  | "spreadsheet.capability.proofingServiceUnavailable"
  | "spreadsheet.capability.queryTableSelectionRequired"
  | "spreadsheet.capability.providerUnavailable"
  | "spreadsheet.capability.publicPublishUnavailable"
  | "spreadsheet.capability.reapplyFilterOrSortRequired"
  | "spreadsheet.capability.savedWorkbookRequired"
  | "spreadsheet.capability.shareOperationInProgress"
  | "spreadsheet.capability.shareUnavailable"
  | "spreadsheet.capability.shareWorkbookRequired"
  | "spreadsheet.capability.sheetBackgroundWebUnavailable"
  | "spreadsheet.capability.sheetViewRequired"
  | "spreadsheet.capability.sheetViewSavedWorkbookRequired"
  | "spreadsheet.capability.pivotSourceIsScenarios"
  | "spreadsheet.capability.selectionAlreadyLocked"
  | "spreadsheet.capability.selectionAlreadyUnlocked"
  | "spreadsheet.capability.selectionRequired"
  | "spreadsheet.capability.sheetProtected"
  | "spreadsheet.capability.splitPaneLayersUnsupported"
  | "spreadsheet.capability.splitPaneSelectionOutsideGrid"
  | "spreadsheet.capability.tableAlreadyActive"
  | "spreadsheet.capability.tableLimit"
  | "spreadsheet.capability.temporarySheetViewRequired"
  | "spreadsheet.capability.truePauseProtectionUnavailable"
  | "spreadsheet.capability.unsupported"
  | "spreadsheet.capability.writerUnavailable"
  | "spreadsheet.capability.workbookStructureProtected"
  | "spreadsheet.capability.workbookProtectionUnsupported"
  | "spreadsheet.capability.workbookProtectionUnavailable"
  | "spreadsheet.capability.workspaceAiEgressPolicyUnavailable"
  | "spreadsheet.capability.workspaceAiProviderUnavailable";

type SpreadsheetVisiblePredicate =
  | "change-history"
  | "copilot-apply"
  | "data-analysis"
  | "data-query-create"
  | "data-query-refresh"
  | "data-query-refresh-all"
  | "data-query-properties"
  | "active-comment"
  | "active-comment-editable"
  | "active-remark-editable"
  | "manual-calculation"
  | "active-note-editable"
  | "active-filter"
  | "active-native-chart"
  | "active-native-chart-color-editable"
  | "active-native-chart-editable"
  | "active-native-chart-line-editable"
  | "active-native-chart-style-editable"
  | "active-native-chart-series-orientation-editable"
  | "active-pivot"
  | "active-pivot-clear-all-editable"
  | "active-pivot-chart-create-editable"
  | "active-pivot-chart"
  | "active-pivot-chart-editable"
  | "active-pivot-chart-options-editable"
  | "active-pivot-editable"
  | "active-pivot-source-editable"
  | "active-pivot-output"
  | "active-picture-effects-editable"
  | "active-picture-metadata-editable"
  | "active-picture-transform-editable"
  | "active-shape-editable"
  | "active-slicer-editable"
  | "active-slicer-filter-editable"
  | "active-sparkline-editable"
  | "active-table-alt-text-editable"
  | "active-table-name-editable"
  | "active-table-resize-editable"
  | "active-table-style-editable"
  | "active-table-pivot-source"
  | "active-table-totals-editable"
  | "alignment-editable"
  | "cell-alignment-editable"
  | "reapplicable-filter-or-sort"
  | "always"
  | "comments"
  | "comment-post-editable"
  | "clipboard-copy-picture"
  | "clipboard-paste-available"
  | "clipboard-paste-editable"
  | "cell-style-editable"
  | "cell-shift-down-editable"
  | "cell-shift-right-editable"
  | "cell-shift-up-editable"
  | "cell-shift-left-editable"
  | "conditional-format-selection-clearable"
  | "conditional-format-sheet-clearable"
  | "conditional-format-manager-apply"
  | "conditional-format-editable"
  | "data-source-inventory"
  | "data-from-picture"
  | "data-from-picture-clipboard"
  | "data-form-delete"
  | "data-form-delete-apply"
  | "data-form-launchable"
  | "data-form-open"
  | "data-form-restore"
  | "data-validation-launcher"
  | "data-validation-apply"
  | "data-validation-clear-draft"
  | "fullscreen"
  | "font-editable"
  | "cell-font-editable"
  | "formula-evaluation"
  | "formula-trace-clear"
  | "format-cells-apply"
  | "filter-editable"
  | "find-replace-match"
  | "find-replace-current-editable"
  | "find-replace-all-editable"
  | "import-ready"
  | "import-ready-editable"
  | "menu-children"
  | "merge-across-editable"
  | "merge-selection-editable"
  | "multi-column-editable"
  | "multi-row-editable"
  | "multi-selection-editable"
  | "remove-duplicates-editable"
  | "new-comment"
  | "new-note"
  | "no-verified-authority"
  | "number-format-editable"
  | "notes"
  | "outline-group-columns"
  | "outline-group-rows"
  | "outline-ungroup-columns"
  | "outline-ungroup-rows"
  | "page-break-insertable"
  | "page-break-removable"
  | "page-break-resettable"
  | "page-margin-applicable"
  | "page-setup-clearable"
  | "page-setup-editable"
  | "page-setup-launchable"
  | "paste-special-apply"
  | "picture-rotation-apply"
  | "print-area-addable"
  | "print-area-clearable"
  | "print-area-settable"
  | "saved-window"
  | "selection"
  | "cell-selection"
  | "cell-selection-editable"
  | "selection-clipboard"
  | "selection-display-text"
  | "selection-editable"
  | "selection-has-merge-editable"
  | "selection-can-lock"
  | "selection-can-unlock"
  | "sheet-view-exit"
  | "sheet-view-keep"
  | "sheet-view-new"
  | "sheet-view-options"
  | "sheet-view-switch"
  | "sort-authorized"
  | "split-pane-toggle"
  | "sheet-editable"
  | "sheet-protection-configured"
  | "sheet-protected"
  | "share-clipboard"
  | "share-web"
  | "shape-rotation-apply"
  | "table-create"
  | "active-table-slicer-insert"
  | "table-slicer-insert-apply"
  /** Insert ▸ Slicer: a Table or a PivotTable under the selection. */
  | "active-slicer-insert"
  | "active-pivot-slicer-insert"
  | "pivot-slicer-insert-apply"
  | "table-resize-apply"
  | "table-alt-text-apply"
  | "version-history"
  | "workbook-protection-launchable"
  | "workbook-protection-applicable"
  | "workbook-link-inventory"
  | "workbook-sheet-structure-editable"
  | "worksheet-structure-editable"
  /** Whole-line inserts/deletes: Excel's Insert/Delete rows/columns
   *  permissions on a protected sheet (`structuralLineEdits`). */
  | "worksheet-insert-rows-editable"
  | "worksheet-insert-columns-editable"
  | "worksheet-delete-rows-editable"
  | "worksheet-delete-columns-editable"
  /** Insert/Edit Link: Excel's Insert hyperlinks permission + unlocked cell. */
  | "hyperlink-insertable"
  /** Insert shapes/pictures/text boxes/charts: Excel's Edit objects. */
  | "object-insertable"
  /** Insert PivotTable/PivotChart: Excel's Use PivotTable & PivotChart. */
  | "pivot-insertable";

export interface SpreadsheetVisibleCommandContext {
  /** True only for this editor's currently validated private Copilot receipt. */
  readonly copilotApplyAvailable?: boolean;
  readonly activeComment: boolean;
  /** Authenticated actor currently owns live page write/checkout authority. */
  readonly commentAuthoringAvailable?: boolean;
  /** The active threaded-comment composer owns a live, writable launch receipt. */
  readonly commentPostAvailable?: boolean;
  /** Current active cell has a non-empty formatted display value. */
  readonly activeCellHasDisplayText?: boolean;
  /** True only while the ordinary cell grid owns the current selection. */
  readonly activeCellSelectionSupported?: boolean;
  readonly activeNote?: boolean;
  readonly activeRemarkBlocked?: boolean;
  readonly activeFilter: boolean;
  /** A real MRAI-native chart is selected on the active worksheet. */
  readonly activeNativeChart?: boolean;
  /** The exact selected chart generation can be changed by the native writer. */
  readonly activeNativeChartEditable?: boolean;
  /** The selected classic chart has re-authenticated series-colour authority. */
  readonly activeNativeChartColorEditable?: boolean;
  /** The selected classic chart has re-authenticated native style authority. */
  readonly activeNativeChartStyleEditable?: boolean;
  /** The selected chart can switch rows and columns (a PivotChart pivots its PivotTable's fields). */
  readonly activeNativeChartSeriesOrientationEditable?: boolean;
  /** The selected chart's line/marker formatting is the writer's to change. */
  readonly activeNativeChartLineEditable?: boolean;
  /** The active cell belongs to one ready, collision-free PivotTable output. */
  readonly activePivot?: boolean;
  /** A PivotChart whose PivotTable exists is the selected chart. */
  readonly activePivotChart?: boolean;
  /** That PivotChart's PivotTable takes a recipe edit (Expand/Collapse Field). */
  readonly activePivotChartEditable?: boolean;
  /** The PivotChart record itself takes an edit (its field buttons). */
  readonly activePivotChartOptionsEditable?: boolean;
  /** Manual calculation can capture this workbook now (coordinates ready). */
  readonly manualCalculationAvailable?: boolean;
  /** The selected PivotTable generation can use the native recipe writer. */
  readonly activePivotEditable?: boolean;
  /**
   * The selected PivotTable reads a worksheet range or table (a Scenario
   * PivotTable report's cache is its scenarios: there is no range to change).
   */
  readonly activePivotHasSourceRange?: boolean;
  /** Clear All can preserve source/anchor/cache while resetting this layout. */
  readonly activePivotClearAllAvailable?: boolean;
  /** The selected PivotTable owns a selectable derived output rectangle. */
  readonly activePivotOutputAvailable?: boolean;
  /** One worksheet picture owns the current contextual selection. */
  readonly activePicture?: boolean;
  /** Picture effects can currently be authored only by the scratch writer. */
  readonly activePictureEffectsEditable?: boolean;
  /** Media replacement and Alt Text have verified native part identity. */
  readonly activePictureMetadataEditable?: boolean;
  /** Its DrawingML anchor/delete/order identity is authenticated for writing. */
  readonly activePictureTransformEditable?: boolean;
  /** More Rotation Options owns a live Picture receipt and valid degree draft. */
  readonly pictureRotationApplyAvailable?: boolean;
  /** One drawing shape, rather than a cell/chart/picture, owns selection. */
  readonly activeShape?: boolean;
  /** The selected shape has an authenticated native writer for this workbook. */
  readonly activeShapeEditable?: boolean;
  /** More Rotation Options owns a live Shape receipt and valid degree draft. */
  readonly shapeRotationApplyAvailable?: boolean;
  /** One Slicer Design surface owns a stable selected slicer. */
  readonly activeSlicer?: boolean;
  /** Slicer view/cache parts can be regenerated by the active writer. */
  readonly activeSlicerEditable?: boolean;
  /** The selected slicer's source filter can accept a value-mode mutation. */
  readonly activeSlicerFilterEditable?: boolean;
  /** The active cell owns one normalized Sparkline group member. */
  readonly activeSparkline?: boolean;
  /** The selected Sparkline group can use the native worksheet writer. */
  readonly activeSparklineEditable?: boolean;
  readonly reapplicableFilterOrSort?: boolean;
  readonly activeTable: boolean;
  /** The selected table generation can edit its native x14 alternate text. */
  readonly activeTableAltTextEditable?: boolean;
  /** The selected table generation can be renamed without native loss. */
  readonly activeTableNameEditable?: boolean;
  /** The selected table generation can open the Web Resize Table workflow. */
  readonly activeTableResizeEditable?: boolean;
  /** The selected native table and its full occupied range can change style. */
  readonly activeTableStyleEditable?: boolean;
  /** At least one stable Table field can create a native slicer graph. */
  readonly activeTableSlicerInsertAvailable?: boolean;
  /** The selected PivotTable has a PivotCache field without a slicer. */
  readonly activePivotSlicerInsertAvailable?: boolean;
  /** The open PivotTable slicer dialog still matches the live workbook. */
  readonly pivotSlicerInsertApplyAvailable?: boolean;
  /** The selected Table owns a header+body range accepted by Pivot creation. */
  readonly activeTablePivotSourceAvailable?: boolean;
  /** Its physical Totals Row and native recipe can be rewritten exactly. */
  readonly activeTableTotalsEditable?: boolean;
  /** The open Resize Table dialog owns a current launch receipt and draft. */
  readonly tableResizeApplyAvailable?: boolean;
  /** The open Insert Slicers dialog has current authority and selected fields. */
  readonly tableSlicerInsertApplyAvailable?: boolean;
  /** The open Alt Text dialog still owns its launch generation and draft. */
  readonly tableAltTextApplyAvailable?: boolean;
  readonly canOpenWorkbookWindow: boolean;
  readonly clipboardSupported?: boolean;
  /** Current internal rich-cell clipboard can satisfy a Home Paste command. */
  readonly richClipboardAvailable?: boolean;
  /** One copied chart is the clipboard: Paste Special lists what it pastes as. */
  readonly chartClipboardAvailable?: boolean;
  /** Open Paste Special owns its exact clipboard, sheet, selection, and mode. */
  readonly pasteSpecialApplyAvailable?: boolean;
  readonly commentedCellCount: number;
  readonly notedCellCount?: number;
  /** True when a native/imported source can be inspected without executing it. */
  readonly dataSourceInventoryInspectable?: boolean;
  /** Exact workspace policy/provider/page authority permits transient OCR. */
  readonly dataFromPictureAvailable?: boolean;
  /** Browser Clipboard API can supply an image without a hidden fallback. */
  readonly dataFromPictureClipboardAvailable?: boolean;
  /** Precise fail-closed reason resolved from the live server capability. */
  readonly dataFromPictureReason?: SpreadsheetVisibleCommandReason;
  /** Current selection is a header plus 1..32 fields and at least one row. */
  readonly dataFormLaunchAvailable?: boolean;
  /** The modal still owns its exact page/sheet/list authority. */
  readonly dataFormOpenAvailable?: boolean;
  /** A dirty current record can be restored without touching the worksheet. */
  readonly dataFormRestoreAvailable?: boolean;
  /** A safe from-scratch record deletion can enter confirmation. */
  readonly dataFormDeleteAvailable?: boolean;
  /** The pending confirmation still owns that exact safe record generation. */
  readonly dataFormDeleteApplyAvailable?: boolean;
  /** The open Data Validation dialog owns a current, writable selection receipt. */
  readonly dataValidationApplyAvailable?: boolean;
  /** The current validation selection can enter the non-mutating Clear draft. */
  readonly dataValidationClearDraftAvailable?: boolean;
  readonly fullscreenSupported: boolean;
  /** Current worksheet layers can be rendered faithfully in every split pane. */
  readonly splitPaneRenderingSupported?: boolean;
  /** Existing native split state can always be removed, even outside selection. */
  readonly splitPanePresent?: boolean;
  /** Active cell can seed a new split inside the persisted worksheet grid. */
  readonly splitPaneSelectionFitsGrid?: boolean;
  /** Live Home font projection is valid for the current selection. */
  readonly homeFontControlsAvailable?: boolean;
  /**
   * A SHAPE (or a SmartArt shape / graphic) is selected, so Home ▸ Font
   * formats its label instead of the cells behind it (Excel's own rule).
   */
  readonly homeFontShapeScope?: boolean;
  /** That shape's label may be written: unprotected, writable route, text. */
  readonly homeFontShapeEditable?: boolean;
  /**
   * ANY drawing object or chart is selected (a shape, a picture, a chart, a
   * slicer, an ink stroke, a preserved diagram): Home never writes the cells
   * under it. What the object cannot take is greyed with the reason; what it
   * can take (a shape's label, a chart's text) is routed to the object.
   */
  readonly homeObjectSelected?: boolean;
  /** A chart is selected: Home ▸ Font formats the chart's own text. */
  readonly homeChartTextScope?: boolean;
  readonly homeChartTextEditable?: boolean;
  /** Live Home alignment projection is valid for the current selection. */
  readonly homeAlignmentControlsAvailable?: boolean;
  /** Live Home number-format projection is valid for the current selection. */
  readonly homeNumberFormatControlsAvailable?: boolean;
  /** Directional cell shifts are lossless for every formula in the sheet. */
  readonly cellShiftDownAvailable?: boolean;
  readonly cellShiftRightAvailable?: boolean;
  readonly cellShiftUpAvailable?: boolean;
  readonly cellShiftLeftAvailable?: boolean;
  /** The active formula can be evaluated by the bounded audit dialog. */
  readonly formulaEvaluationAvailable?: boolean;
  /** The current audit session owns trace arrows that can be removed. */
  readonly formulaTraceAvailable?: boolean;
  /** The open Find/Replace surface owns at least one live match. */
  readonly findReplaceMatchAvailable?: boolean;
  /** The currently focused/first match can use the native cell writer. */
  readonly findReplaceCurrentEditable?: boolean;
  /** Every current-scope match can be replaced in one transaction. */
  readonly findReplaceAllEditable?: boolean;
  /** Live named-cell-style projection can apply to the current selection. */
  readonly homeCellStyleControlsAvailable?: boolean;
  /** Current selection intersects at least one conditional-format rule. */
  readonly homeSelectionHasConditionalRule?: boolean;
  /** Selection-only clear would truncate a preserve-only native Applies To. */
  readonly conditionalFormatSelectionClearBlocked?: boolean;
  /** Active sheet owns at least one editable conditional-format rule. */
  readonly homeSheetHasConditionalRules?: boolean;
  /** A current Manage Rules intent can be committed by the live native writer. */
  readonly conditionalFormatManagerApplyAvailable?: boolean;
  /** Effective protection and Sheet View projection permit grid remapping. */
  readonly homeStructureControlsAvailable?: boolean;
  readonly hasMultiColumnSelection: boolean;
  readonly hasMultiRowSelection: boolean;
  readonly hasMultiSelection: boolean;
  /**
   * The selection as the merge commands see it: expanded to the merges it
   * touches. Selecting a merged cell selects its whole area (Microsoft
   * Support, "Merge and unmerge cells"), so Merge & Center stays live on one,
   * where it is the unmerge toggle. Falls back to the plain selection.
   */
  readonly mergeSelectionMultiCell?: boolean;
  /** The cells Merge would take reach into a table, where Excel greys it. */
  readonly mergeSelectionInTable?: boolean;
  readonly mergeSelectionMultiColumn?: boolean;
  readonly removeDuplicatesAvailable?: boolean;
  readonly removeDuplicatesUnavailableReason?: SpreadsheetVisibleCommandReason;
  readonly hasSelection: boolean;
  readonly dataAnalysisAvailable?: boolean;
  readonly dataQueryCreateAvailable?: boolean;
  readonly dataQueryRefreshAvailable?: boolean;
  readonly dataQueryRefreshAllAvailable?: boolean;
  /** The active cell is in a query's output table (Data > Properties). */
  readonly dataQueryPropertiesAvailable?: boolean;
  readonly importInProgress: boolean;
  readonly nativeWriteRoute: SpreadsheetNativeWriteRoute;
  readonly outlineGroupColumnsAvailable?: boolean;
  readonly outlineGroupRowsAvailable?: boolean;
  /**
   * The active sheet's outline could not be read exactly as its file holds
   * it (`outlineGroupsNativeBlocker` on the patchback route), so the native
   * writer would refuse any outline change there.
   */
  readonly outlineNativeBlocked?: boolean;
  readonly outlineUngroupColumnsAvailable?: boolean;
  readonly outlineUngroupRowsAvailable?: boolean;
  /** Live active-sheet identity can safely host Page Setup inspection UI. */
  readonly pageSetupLaunchable?: boolean;
  /** Effective protection and the native writer permit Page Setup changes. */
  readonly pageSetupEditable?: boolean;
  /** The active worksheet owns normalized settings that Clear can remove. */
  readonly pageSetupClearAvailable?: boolean;
  /** Current pending/custom margin receipt can be applied losslessly. */
  readonly pageMarginApplyAvailable?: boolean;
  /** Open Format Cells owns one current editable selection and draft. */
  readonly formatCellsApplyAvailable?: boolean;
  /** Current selection can become the worksheet print area. */
  readonly printAreaSetAvailable?: boolean;
  /** The active worksheet currently owns a removable print area. */
  readonly printAreaClearAvailable?: boolean;
  /** A print area exists and the selection lies outside it (Excel's Add). */
  readonly printAreaAddAvailable?: boolean;
  /** Current active cell supports inserting a manual page break. */
  readonly pageBreakInsertAvailable?: boolean;
  /** Current active cell intersects a removable manual page break. */
  readonly pageBreakRemoveAvailable?: boolean;
  /** The active worksheet owns at least one resettable manual page break. */
  readonly pageBreakResetAvailable?: boolean;
  /** Exact lock state for the current selection; mixed enables either action. */
  readonly selectionLockState?: boolean | "mixed" | null;
  /** Backward-compatible selection receipt for older catalog consumers. */
  readonly selectionLockStateAvailable: boolean;
  readonly selectionEditable: boolean;
  /** Current Sheet View/selection permits copying the selected cell range. */
  readonly selectionClipboardAvailable?: boolean;
  /** Current selection intersects at least one merged range. */
  readonly selectionHasMerge?: boolean;
  readonly shareAuthorityAvailable?: boolean;
  readonly shareOperationBusy?: boolean;
  readonly webShareSupported?: boolean;
  /** Protected-sheet Allow Sort + current selection lock authority. */
  readonly sortAuthorized?: boolean;
  /** Durable worksheet protection remains configured, even while paused. */
  readonly sheetProtectionConfigured?: boolean;
  readonly sheetProtected: boolean;
  /** Protected sheet: the Format cells permission was granted. */
  readonly formatCellsAllowed?: boolean;
  /**
   * Excel's Protect Sheet permissions resolved for the active sheet (every
   * one true when it is unprotected or paused). Absent: the older reading,
   * where a protected sheet allows none of them.
   */
  readonly protectionAllows?: Readonly<SheetProtectionPermissions>;
  /**
   * Whole-line structure edits the selection may make now: the Insert/Delete
   * rows/columns permissions, and for deletes every cell of the selected
   * lines unlocked. Absent: protection blocks them all.
   */
  readonly structuralLineEdits?: Readonly<{
    insertRows: boolean;
    insertColumns: boolean;
    deleteRows: boolean;
    deleteColumns: boolean;
  }>;
  /** Named Sheet View availability, resolved from the live active sheet. */
  readonly namedSheetViewCount?: number;
  readonly namedSheetViewSwitchableCount?: number;
  readonly sheetViewMode?: "default" | "temporary" | "named";
  /**
   * The Sheet View box holds an edited, unsent name. Keep commits that edit
   * with no dialog; in a temporary view it needs none and uses a default name.
   */
  readonly sheetViewNameDraftPending?: boolean;
  /** A stored "just me / everyone" default that the Sheet View menu can reset. */
  readonly sheetViewScopeRemembered?: boolean;
  readonly sheetViewAuthoringAvailable?: boolean;
  readonly sheetViewCapacityAvailable?: boolean;
  readonly tableCapacityAvailable: boolean;
  readonly versionHistoryAvailable: boolean;
  /**
   * Review > Show Changes reads the workbook's server change log, which only
   * a saved workbook page has (Excel: a workbook saved to the cloud).
   */
  readonly changeHistoryAvailable?: boolean;
  /** True only after the bounded package inventory found external-link parts. */
  readonly workbookLinksAvailable?: boolean;
  /** Current workbook topology lock, independent from worksheet protection. */
  readonly workbookStructureLocked?: boolean;
  /** A live dialog receipt and valid password intent can be committed. */
  readonly workbookProtectionApplyAvailable?: boolean;
  /** Imported strong/unknown workbookProtection authority is preserve-only. */
  readonly workbookProtectionNativeBlocked?: boolean;
}

type SpreadsheetVisibleIoProfile =
  | "worksheet-copilot"
  | "session-only"
  | "unsupported"
  | "workbook-calculation-mode"
  | "workbook-date-system"
  | "workbook-import"
  | "workbook-defined-names"
  | "workbook-structure-protection"
  | "workbook-view-preference"
  | "worksheet-advanced-filter"
  | "worksheet-cells"
  | "worksheet-cell-controls"
  | "worksheet-cell-pictures"
  | "worksheet-background-picture"
  | "worksheet-queries"
  | "worksheet-comments"
  | "worksheet-conditional-formats"
  | "worksheet-data-form-delete"
  | "worksheet-drawing-objects"
  | "worksheet-shape-format"
  | "worksheet-chart-metadata"
  | "worksheet-charts"
  | "worksheet-filters"
  | "worksheet-freeze"
  | "worksheet-outline"
  | "worksheet-paste-special"
  | "worksheet-print-settings"
  | "worksheet-pivots"
  | "worksheet-picture-metadata"
  | "worksheet-picture-transforms"
  | "worksheet-protection"
  | "worksheet-structure"
  | "worksheet-sheet-views"
  | "worksheet-slicers"
  | "worksheet-split"
  | "worksheet-table"
  | "worksheet-sparklines"
  | "worksheet-validation";

/**
 * Catalog grouping only. It is not a completeness denominator: physical Web
 * evidence, the MRAI renderer and executable handler authority are reconciled
 * independently in `SPREADSHEET_WEB_ATLAS_COVERAGE_TRUTH` (`./visible-command-coverage`).
 */
export const SPREADSHEET_VISIBLE_CONTROL_CATALOG_GROUPS = Object.freeze([
  "Home",
  "Insert",
  "Share",
  "Page Layout",
  "Formulas",
  "Data",
  "Review",
  "View",
  "Draw",
] as const);

type SpreadsheetVisibleControlGroup =
  | (typeof SPREADSHEET_VISIBLE_CONTROL_CATALOG_GROUPS)[number]
  | "Table Design"
  | "Shape Format"
  | "Slicer Design"
  | "Picture Format"
  | "Sparkline Design"
  | "Chart Design"
  | "PivotTable Analyze"
  | "PivotTable Design"
  | "PivotChart Analyze";

interface SpreadsheetVisibleControlDefinition {
  readonly controlId: string;
  readonly featureId: string;
  readonly canonicalId: string;
  readonly group: SpreadsheetVisibleControlGroup;
  readonly order: number;
  readonly controlKind: SpreadsheetCapabilityControlKind;
  readonly predicate: SpreadsheetVisiblePredicate;
  readonly reasonKey: SpreadsheetVisibleCommandReason | null;
  readonly command: "bound" | "disclosure" | "unsupported";
  readonly undoPolicy: "single-transaction" | "content-none";
  readonly ioProfile: SpreadsheetVisibleIoProfile;
  readonly children?: readonly string[];
}

const RIBBON_PATH_BY_GROUP = Object.freeze({
  Home: "Ribbon.Home",
  Insert: "Ribbon.Insert",
  Share: "Ribbon.Share",
  "Page Layout": "Ribbon.PageLayout",
  Formulas: "Ribbon.Formulas",
  Data: "Ribbon.Data",
  Review: "Ribbon.Review",
  View: "Ribbon.View",
  Draw: "Ribbon.Draw",
  "Table Design": "Ribbon.Context.TableDesign",
  "Shape Format": "Ribbon.Context.ShapeFormat",
  "Slicer Design": "Ribbon.Context.SlicerDesign",
  "Picture Format": "Ribbon.Context.PictureFormat",
  "Sparkline Design": "Ribbon.Context.SparklineDesign",
  "Chart Design": "Ribbon.Context.ChartDesign",
  "PivotTable Analyze": "Ribbon.Context.PivotTableAnalyze",
  "PivotTable Design": "Ribbon.Context.PivotTableDesign",
  "PivotChart Analyze": "Ribbon.Context.PivotChartAnalyze",
} satisfies Record<SpreadsheetVisibleControlGroup, string>);

const define = (
  definition: SpreadsheetVisibleControlDefinition,
): SpreadsheetVisibleControlDefinition => Object.freeze(definition);

/**
 * Bounded Home catalog authority. Every bound leaf below is a physical
 * alias of an existing Data command semantic; it must retain the same feature
 * id, predicate, handler identity, Undo policy and writer contract. Analyze
 * Data belongs physically to Home but remains provider-blocked and omitted.
 */
export const SPREADSHEET_HOME_GROUNDWORK_CONTROL_DEFINITIONS = Object.freeze([
  define({
    controlId: "XLSCTL.HOME.EDIT.SORT_FILTER_MENU",
    featureId: "spreadsheet.home.edit.sortFilterMenu",
    canonicalId: "XLS-HOM-EDIT-04",
    group: "Home",
    order: 34,
    controlKind: "menu",
    predicate: "menu-children",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    command: "disclosure",
    undoPolicy: "content-none",
    ioProfile: "session-only",
    children: [
      "XLSCTL.HOME.EDIT.SORT_FILTER.ASCENDING",
      "XLSCTL.HOME.EDIT.SORT_FILTER.DESCENDING",
      "XLSCTL.HOME.EDIT.SORT_FILTER.CUSTOM_SORT",
      "XLSCTL.HOME.EDIT.SORT_FILTER.FILTER",
      "XLSCTL.HOME.EDIT.SORT_FILTER.CLEAR",
      "XLSCTL.HOME.EDIT.SORT_FILTER.REAPPLY",
    ],
  }),
  ...[
    [
      "XLSCTL.HOME.EDIT.SORT_FILTER.ASCENDING",
      "spreadsheet.data.sortAscending",
      35,
      "button",
    ],
    [
      "XLSCTL.HOME.EDIT.SORT_FILTER.DESCENDING",
      "spreadsheet.data.sortDescending",
      36,
      "button",
    ],
    [
      "XLSCTL.HOME.EDIT.SORT_FILTER.CUSTOM_SORT",
      "spreadsheet.data.openSortDialog",
      37,
      "dialog-launcher",
    ],
  ].map(([controlId, featureId, order, controlKind]) =>
    define({
      controlId: String(controlId),
      featureId: String(featureId),
      canonicalId: "XLS-DAT-SORT-01",
      group: "Home",
      order: Number(order),
      controlKind: controlKind as SpreadsheetCapabilityControlKind,
      predicate: "sort-authorized",
      reasonKey: "spreadsheet.capability.sheetProtected",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-cells",
    }),
  ),
  define({
    controlId: "XLSCTL.HOME.EDIT.SORT_FILTER.FILTER",
    featureId: "spreadsheet.data.toggleAutoFilter",
    canonicalId: "XLS-DAT-FILTER-01",
    group: "Home",
    order: 38,
    controlKind: "toggle",
    predicate: "filter-editable",
    reasonKey: "spreadsheet.capability.selectionRequired",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-filters",
  }),
  define({
    controlId: "XLSCTL.HOME.EDIT.SORT_FILTER.CLEAR",
    featureId: "spreadsheet.data.clearFilters",
    canonicalId: "XLS-DAT-FILTER-01",
    group: "Home",
    order: 39,
    controlKind: "button",
    predicate: "active-filter",
    reasonKey: "spreadsheet.capability.activeFilterRequired",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-filters",
  }),
  define({
    controlId: "XLSCTL.HOME.EDIT.SORT_FILTER.REAPPLY",
    featureId: "spreadsheet.data.reapplyFilters",
    canonicalId: "XLS-DAT-FILTER-01",
    group: "Home",
    order: 40,
    controlKind: "button",
    predicate: "reapplicable-filter-or-sort",
    reasonKey: "spreadsheet.capability.reapplyFilterOrSortRequired",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-filters",
  }),
  define({
    controlId: "XLSCTL.HOME.EDIT.ANALYZE_DATA",
    featureId: "spreadsheet.data.analyzeData",
    canonicalId: "XLS-DAT-TOOLS-01",
    group: "Home",
    order: 41,
    controlKind: "dialog-launcher",
    predicate: "data-analysis",
    reasonKey: "spreadsheet.capability.analyzeDataUnavailable",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

const defineHomeCoreBound = (config: {
  readonly controlId: string;
  readonly featureId: string;
  readonly canonicalId: string;
  readonly order: number;
  readonly controlKind: SpreadsheetCapabilityControlKind;
  readonly predicate: SpreadsheetVisiblePredicate;
  readonly reasonKey: SpreadsheetVisibleCommandReason | null;
  readonly undoPolicy: "single-transaction" | "content-none";
  readonly ioProfile: SpreadsheetVisibleIoProfile;
}): SpreadsheetVisibleControlDefinition =>
  define({
    ...config,
    group: "Home",
    command: "bound",
  });

const defineHomeCoreDisclosure = (config: {
  readonly controlId: string;
  readonly featureId: string;
  readonly canonicalId: string;
  readonly order: number;
  readonly children: readonly string[];
  /**
   * A disclosure follows its children by default (`menu-children`: open while
   * any item is live). Excel keeps some split-button arrows openable on a
   * writable sheet even when every item inside is greyed, so the user can read
   * each item's own reason; those declare `sheet-editable` and answer for
   * themselves.
   */
  readonly predicate?:
    | "menu-children"
    | "sheet-editable"
    | "workbook-sheet-structure-editable";
}): SpreadsheetVisibleControlDefinition =>
  define({
    ...config,
    group: "Home",
    controlKind: "menu",
    predicate: config.predicate ?? "menu-children",
    reasonKey:
      config.predicate === "sheet-editable"
        ? "spreadsheet.capability.sheetProtected"
        : config.predicate === "workbook-sheet-structure-editable"
          ? "spreadsheet.capability.workbookStructureProtected"
          : "spreadsheet.capability.commandUnavailable",
    command: "disclosure",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  });

/**
 * First production Home vertical slice. This list deliberately contains only
 * fixed controls and bounded aggregate value controls whose existing
 * Spreadsheet callbacks can be routed through the authenticated command
 * registry. Parameterized controls use Atlas occurrence receipts and a
 * consume-once host value receipt; their dynamic values are not invented as
 * fixed catalog leaves. Table-style and conditional-rule galleries remain
 * outside this slice because they need separate semantic value registries.
 */
/**
 * Internal commit authority for the MRAI Manage Rules dialog. It is not an
 * additional Excel ribbon occurrence, so keep it out of the frozen Home
 * physical-control denominator while still binding every persisted action to
 * one exact live writer.
 */
export const SPREADSHEET_CONDITIONAL_FORMAT_MANAGER_CONTROL_DEFINITIONS =
  Object.freeze([
    defineHomeCoreBound({
      controlId: "XLSCTL.HOME.STYLE.CONDITIONAL.MANAGE_RULES.APPLY",
      featureId: "spreadsheet.home.styles.applyConditionalManager",
      canonicalId: "XLS-HOM-STYLE-01",
      order: 48,
      controlKind: "button",
      predicate: "conditional-format-manager-apply",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-conditional-formats",
    }),
  ] as const);

export const SPREADSHEET_HOME_CORE_VERTICAL_CONTROL_DEFINITIONS = Object.freeze([
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.CLIPBOARD.PASTE_MENU",
    featureId: "spreadsheet.home.clipboard.paste",
    canonicalId: "XLS-HOM-CLIP-01",
    order: 1,
    controlKind: "button",
    predicate: "clipboard-paste-editable",
    reasonKey: "spreadsheet.capability.clipboardUnavailable",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-cells",
  }),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.CLIPBOARD.PASTE_MENU.DISCLOSURE",
    featureId: "spreadsheet.home.clipboard.pasteMenuDisclosure",
    canonicalId: "XLS-HOM-CLIP-01",
    order: 2,
    children: [
      "XLSCTL.HOME.CLIPBOARD.PASTE.VALUES",
      "XLSCTL.HOME.CLIPBOARD.PASTE.FORMATS",
      "XLSCTL.HOME.CLIPBOARD.PASTE.NO_BORDERS",
      "XLSCTL.HOME.CLIPBOARD.PASTE.TRANSPOSE",
    ],
  }),
  ...[
    ["XLSCTL.HOME.CLIPBOARD.PASTE.VALUES", "pasteValues", 3],
    ["XLSCTL.HOME.CLIPBOARD.PASTE.FORMATS", "pasteFormats", 4],
    ["XLSCTL.HOME.CLIPBOARD.PASTE.NO_BORDERS", "pasteNoBorders", 5],
    ["XLSCTL.HOME.CLIPBOARD.PASTE.TRANSPOSE", "pasteTranspose", 6],
  ].map(([controlId, suffix, order]) =>
    defineHomeCoreBound({
      controlId: String(controlId),
      featureId: `spreadsheet.home.clipboard.${String(suffix)}`,
      canonicalId: "XLS-HOM-CLIP-01",
      order: Number(order),
      controlKind: "button",
      predicate: "clipboard-paste-editable",
      reasonKey: "spreadsheet.capability.clipboardUnavailable",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-cells",
    }),
  ),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.PASTE_SPECIAL.DIALOG",
    featureId: "spreadsheet.home.clipboard.openPasteSpecial",
    canonicalId: "XLS-HOM-CLIP-01",
    order: 7,
    controlKind: "dialog-launcher",
    predicate: "clipboard-paste-available",
    reasonKey: "spreadsheet.capability.clipboardUnavailable",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.PASTE_SPECIAL.MRAI.APPLY",
    featureId: "spreadsheet.home.clipboard.applyPasteSpecial",
    canonicalId: "XLS-HOM-CLIP-01",
    order: 8,
    controlKind: "button",
    predicate: "paste-special-apply",
    reasonKey: "spreadsheet.capability.writerUnavailable",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-paste-special",
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.PASTE_SPECIAL.MRAI.DISMISS",
    featureId: "spreadsheet.home.clipboard.closePasteSpecial",
    canonicalId: "XLS-HOM-CLIP-01",
    order: 9,
    controlKind: "button",
    predicate: "always",
    reasonKey: null,
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.FONT.COLOR_MENU.DISCLOSURE",
    featureId: "spreadsheet.home.font.textColorDisclosure",
    canonicalId: "XLS-HOM-FONT-01",
    order: 23,
    children: ["XLSCTL.HOME.FONT.COLOR_MENU"],
  }),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.FONT.FILL_MENU.DISCLOSURE",
    featureId: "spreadsheet.home.font.fillColorDisclosure",
    canonicalId: "XLS-HOM-FONT-01",
    order: 24,
    children: ["XLSCTL.HOME.FONT.FILL_MENU"],
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.CLIPBOARD.CUT",
    featureId: "spreadsheet.home.clipboard.cut",
    canonicalId: "XLS-HOM-CLIP-02",
    order: 7,
    controlKind: "button",
    predicate: "selection-clipboard",
    reasonKey: "spreadsheet.capability.clipboardUnavailable",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.CLIPBOARD.COPY_MENU",
    featureId: "spreadsheet.home.clipboard.copy",
    canonicalId: "XLS-HOM-CLIP-02",
    order: 8,
    controlKind: "button",
    predicate: "selection-clipboard",
    reasonKey: "spreadsheet.capability.clipboardUnavailable",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.CLIPBOARD.COPY_MENU.DISCLOSURE",
    featureId: "spreadsheet.home.clipboard.copyMenuDisclosure",
    canonicalId: "XLS-HOM-CLIP-02",
    order: 9,
    children: [
      "XLSCTL.HOME.CLIPBOARD.COPY",
      "XLSCTL.HOME.CLIPBOARD.COPY_AS_PICTURE",
    ],
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.CLIPBOARD.COPY",
    featureId: "spreadsheet.home.clipboard.copy",
    canonicalId: "XLS-HOM-CLIP-02",
    order: 10,
    controlKind: "button",
    predicate: "selection-clipboard",
    reasonKey: "spreadsheet.capability.clipboardUnavailable",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.CLIPBOARD.COPY_AS_PICTURE",
    featureId: "spreadsheet.home.clipboard.copyAsPicture",
    canonicalId: "XLS-HOM-CLIP-02",
    order: 11,
    controlKind: "button",
    predicate: "clipboard-copy-picture",
    reasonKey: "spreadsheet.capability.clipboardUnavailable",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.CLIPBOARD.FORMAT_PAINTER",
    featureId: "spreadsheet.home.clipboard.formatPainter",
    canonicalId: "XLS-HOM-CLIP-02",
    order: 12,
    controlKind: "toggle",
    predicate: "selection-clipboard",
    reasonKey: "spreadsheet.capability.clipboardUnavailable",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  ...[
    ["XLSCTL.HOME.FONT.FAMILY", "family", 13],
    ["XLSCTL.HOME.FONT.SIZE", "size", 14],
    ["XLSCTL.HOME.FONT.COLOR_MENU", "textColor", 20],
    ["XLSCTL.HOME.FONT.FILL_MENU", "fillColor", 21],
    ["XLSCTL.HOME.FONT.BORDER_APPLY", "borderPreset", 22],
  ].map(([controlId, suffix, order]) =>
    defineHomeCoreBound({
      controlId: String(controlId),
      featureId: `spreadsheet.home.font.${String(suffix)}`,
      canonicalId: "XLS-HOM-FONT-01",
      order: Number(order),
      controlKind:
        controlId === "XLSCTL.HOME.FONT.FAMILY" ||
        controlId === "XLSCTL.HOME.FONT.SIZE"
          ? "combobox"
          : "menu",
      predicate:
        controlId === "XLSCTL.HOME.FONT.FILL_MENU" ||
        controlId === "XLSCTL.HOME.FONT.BORDER_APPLY"
          ? "cell-font-editable"
          : "font-editable",
      reasonKey: "spreadsheet.capability.sheetProtected",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-cells",
    }),
  ),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.FONT.FAMILY.DISCLOSURE",
    featureId: "spreadsheet.home.font.familyDisclosure",
    canonicalId: "XLS-HOM-FONT-01",
    order: 13,
    children: ["XLSCTL.HOME.FONT.FAMILY"],
  }),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.FONT.BORDER_MENU",
    featureId: "spreadsheet.home.font.borderMenu",
    canonicalId: "XLS-HOM-FONT-01",
    order: 22,
    children: [
      "XLSCTL.HOME.FONT.BORDER_APPLY",
      "XLSCTL.HOME.FONT.BORDER_MORE",
    ],
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.FONT.BORDER_MORE",
    featureId: "spreadsheet.home.font.openFormatCells",
    canonicalId: "XLS-HOM-FONT-01",
    order: 23,
    controlKind: "dialog-launcher",
    // More Borders opens the CELLS' dialog: with an object selected it would
    // reach the cells under it.
    predicate: "cell-font-editable",
    reasonKey: "spreadsheet.capability.selectionRequired",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  ...[
    ["XLSCTL.HOME.FONT.GROW", "grow", 13],
    ["XLSCTL.HOME.FONT.SHRINK", "shrink", 14],
    ["XLSCTL.HOME.FONT.BOLD", "bold", 15],
    ["XLSCTL.HOME.FONT.ITALIC", "italic", 16],
    ["XLSCTL.HOME.FONT.STRIKETHROUGH", "strikethrough", 17],
  ].map(([controlId, suffix, order]) =>
    defineHomeCoreBound({
      controlId: String(controlId),
      featureId: `spreadsheet.home.font.${String(suffix)}`,
      canonicalId: "XLS-HOM-FONT-01",
      order: Number(order),
      controlKind: "toggle",
      predicate: "font-editable",
      reasonKey: "spreadsheet.capability.sheetProtected",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-cells",
    }),
  ),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.FONT.UNDERLINE_MENU",
    featureId: "spreadsheet.home.font.underlineSingle",
    canonicalId: "XLS-HOM-FONT-01",
    order: 18,
    controlKind: "toggle",
    predicate: "font-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-cells",
  }),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.FONT.UNDERLINE_MENU.DISCLOSURE",
    featureId: "spreadsheet.home.font.underlineMenuDisclosure",
    canonicalId: "XLS-HOM-FONT-01",
    order: 19,
    children: [
      "XLSCTL.HOME.FONT.UNDERLINE.SINGLE",
      "XLSCTL.HOME.FONT.UNDERLINE.DOUBLE",
    ],
  }),
  ...[
    ["XLSCTL.HOME.FONT.UNDERLINE.SINGLE", "underlineSingle", 20],
    ["XLSCTL.HOME.FONT.UNDERLINE.DOUBLE", "underlineDouble", 21],
  ].map(([controlId, suffix, order]) =>
    defineHomeCoreBound({
      controlId: String(controlId),
      featureId: `spreadsheet.home.font.${String(suffix)}`,
      canonicalId: "XLS-HOM-FONT-01",
      order: Number(order),
      controlKind: "toggle",
      predicate: "font-editable",
      reasonKey: "spreadsheet.capability.sheetProtected",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-cells",
    }),
  ),
  // Home > Font > Phonetic Guide (Excel's 拼音指南 / ふりがな): Show
  // Phonetic Field toggles the selection's c@ph, and Phonetic Settings
  // writes the type, alignment and font of the cells' furigana. Both are
  // cell formatting, so an object selection stands them down.
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.FONT.PHONETIC_MENU",
    featureId: "spreadsheet.home.font.phoneticGuideMenu",
    canonicalId: "XLS-HOM-FONT-01",
    order: 25,
    children: [
      "XLSCTL.HOME.FONT.PHONETIC.SHOW",
      "XLSCTL.HOME.FONT.PHONETIC.SETTINGS",
    ],
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.FONT.PHONETIC.SHOW",
    featureId: "spreadsheet.home.font.showPhoneticField",
    canonicalId: "XLS-HOM-FONT-01",
    order: 26,
    controlKind: "toggle",
    predicate: "cell-font-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-cells",
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.FONT.PHONETIC.SETTINGS",
    featureId: "spreadsheet.home.font.openPhoneticSettings",
    canonicalId: "XLS-HOM-FONT-01",
    order: 27,
    controlKind: "dialog-launcher",
    predicate: "cell-font-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.FONT.PHONETIC_SETTINGS.OK",
    featureId: "spreadsheet.home.font.applyPhoneticSettings",
    canonicalId: "XLS-HOM-FONT-01",
    order: 28,
    controlKind: "button",
    predicate: "cell-font-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-cells",
  }),
  ...[
    ["XLSCTL.HOME.ALIGN.TOP", "top", 22],
    ["XLSCTL.HOME.ALIGN.MIDDLE", "middle", 23],
    ["XLSCTL.HOME.ALIGN.BOTTOM", "bottom", 24],
    ["XLSCTL.HOME.ALIGN.LEFT", "left", 25],
    ["XLSCTL.HOME.ALIGN.CENTER", "center", 26],
    ["XLSCTL.HOME.ALIGN.RIGHT", "right", 27],
    ["XLSCTL.HOME.ALIGN.INDENT_DECREASE", "indentDecrease", 28],
    ["XLSCTL.HOME.ALIGN.INDENT_INCREASE", "indentIncrease", 29],
  ].map(([controlId, suffix, order]) =>
    defineHomeCoreBound({
      controlId: String(controlId),
      featureId: `spreadsheet.home.alignment.${String(suffix)}`,
      canonicalId: "XLS-HOM-ALIGN-01",
      order: Number(order),
      controlKind: "toggle",
      // The indents are cell formatting only; the six alignments reach a
      // selected shape's label.
      predicate: String(controlId).includes("INDENT")
        ? "cell-alignment-editable"
        : "alignment-editable",
      reasonKey: "spreadsheet.capability.sheetProtected",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-cells",
    }),
  ),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.ALIGN.WRAP_MENU",
    featureId: "spreadsheet.home.alignment.wrap",
    canonicalId: "XLS-HOM-ALIGN-01",
    order: 30,
    controlKind: "toggle",
    predicate: "cell-alignment-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-cells",
  }),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.ALIGN.WRAP_MENU.DISCLOSURE",
    featureId: "spreadsheet.home.alignment.wrapMenuDisclosure",
    canonicalId: "XLS-HOM-ALIGN-01",
    order: 31,
    children: [
      "XLSCTL.HOME.ALIGN.WRAP.TOGGLE",
      "XLSCTL.HOME.ALIGN.WRAP.FORMAT_CELLS",
    ],
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.ALIGN.WRAP.TOGGLE",
    featureId: "spreadsheet.home.alignment.wrap",
    canonicalId: "XLS-HOM-ALIGN-01",
    order: 32,
    controlKind: "toggle",
    predicate: "cell-alignment-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-cells",
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.ALIGN.WRAP.FORMAT_CELLS",
    featureId: "spreadsheet.home.alignment.openFormatCells",
    canonicalId: "XLS-HOM-ALIGN-01",
    order: 33,
    controlKind: "dialog-launcher",
    predicate: "cell-selection",
    reasonKey: "spreadsheet.capability.selectionRequired",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.ALIGN.MERGE_MENU",
    featureId: "spreadsheet.home.alignment.mergeCenter",
    canonicalId: "XLS-HOM-ALIGN-01",
    order: 34,
    controlKind: "button",
    predicate: "merge-selection-editable",
    reasonKey: "spreadsheet.capability.multiSelectionRequired",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-structure",
  }),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.ALIGN.MERGE_MENU.DISCLOSURE",
    featureId: "spreadsheet.home.alignment.mergeMenuDisclosure",
    canonicalId: "XLS-HOM-ALIGN-01",
    order: 35,
    children: [
      "XLSCTL.HOME.ALIGN.MERGE.CENTER",
      "XLSCTL.HOME.ALIGN.MERGE.ACROSS",
      "XLSCTL.HOME.ALIGN.MERGE.CELLS",
      "XLSCTL.HOME.ALIGN.MERGE.UNMERGE",
    ],
    // Excel greys Merge & Center only for a protected sheet, a shared
    // workbook, a table or tracked changes; a single-cell selection greys the
    // items, not the arrow, so Unmerge can still be read with its reason.
    predicate: "sheet-editable",
  }),
  ...[
    ["XLSCTL.HOME.ALIGN.MERGE.CENTER", "mergeCenter", 36, "merge-selection-editable"],
    ["XLSCTL.HOME.ALIGN.MERGE.ACROSS", "mergeAcross", 37, "merge-across-editable"],
    ["XLSCTL.HOME.ALIGN.MERGE.CELLS", "mergeCells", 38, "merge-selection-editable"],
    ["XLSCTL.HOME.ALIGN.MERGE.UNMERGE", "unmerge", 39, "selection-has-merge-editable"],
  ].map(([controlId, suffix, order, predicate]) =>
    defineHomeCoreBound({
      controlId: String(controlId),
      featureId: `spreadsheet.home.alignment.${String(suffix)}`,
      canonicalId: "XLS-HOM-ALIGN-01",
      order: Number(order),
      controlKind: "button",
      predicate: predicate as SpreadsheetVisiblePredicate,
      reasonKey:
        predicate === "merge-across-editable"
          ? "spreadsheet.capability.multiColumnSelectionRequired"
          : predicate === "selection-has-merge-editable"
            ? "spreadsheet.capability.selectionRequired"
            : "spreadsheet.capability.multiSelectionRequired",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-structure",
    }),
  ),
  ...[
    ["XLSCTL.HOME.NUMBER.PERCENT", "percent", 40],
    ["XLSCTL.HOME.NUMBER.COMMA", "comma", 41],
    ["XLSCTL.HOME.NUMBER.DECIMAL_INCREASE", "decimalIncrease", 42],
    ["XLSCTL.HOME.NUMBER.DECIMAL_DECREASE", "decimalDecrease", 43],
  ].map(([controlId, suffix, order]) =>
    defineHomeCoreBound({
      controlId: String(controlId),
      featureId: `spreadsheet.home.number.${String(suffix)}`,
      canonicalId: "XLS-HOM-NUM-01",
      order: Number(order),
      controlKind: "button",
      predicate: "number-format-editable",
      reasonKey: "spreadsheet.capability.sheetProtected",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-cells",
    }),
  ),
  ...[
    ["XLSCTL.HOME.NUMBER.FORMAT_APPLY", "format", 44, "combobox"],
    ["XLSCTL.HOME.NUMBER.ACCOUNTING_MENU", "accounting", 45, "menu"],
  ].map(([controlId, suffix, order, controlKind]) =>
    defineHomeCoreBound({
      controlId: String(controlId),
      featureId: `spreadsheet.home.number.${String(suffix)}`,
      canonicalId: "XLS-HOM-NUM-01",
      order: Number(order),
      controlKind: controlKind as SpreadsheetCapabilityControlKind,
      predicate: "number-format-editable",
      reasonKey: "spreadsheet.capability.sheetProtected",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-cells",
    }),
  ),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.NUMBER.FORMAT_COMBO",
    featureId: "spreadsheet.home.number.formatMenu",
    canonicalId: "XLS-HOM-NUM-01",
    order: 44,
    children: [
      "XLSCTL.HOME.NUMBER.FORMAT_APPLY",
      "XLSCTL.HOME.NUMBER.FORMAT.MORE",
    ],
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.NUMBER.FORMAT.MORE",
    featureId: "spreadsheet.home.number.openFormatCells",
    canonicalId: "XLS-HOM-NUM-01",
    order: 45,
    controlKind: "dialog-launcher",
    predicate: "cell-selection",
    reasonKey: "spreadsheet.capability.selectionRequired",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.NUMBER.ACCOUNTING_MENU.DISCLOSURE",
    featureId: "spreadsheet.home.number.accountingDisclosure",
    canonicalId: "XLS-HOM-NUM-01",
    order: 46,
    children: ["XLSCTL.HOME.NUMBER.ACCOUNTING_MENU"],
  }),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.STYLE.CONDITIONAL_MENU",
    featureId: "spreadsheet.home.styles.conditionalMenu",
    canonicalId: "XLS-HOM-STYLE-01",
    order: 44,
    children: [
      "XLSCTL.HOME.STYLE.CONDITIONAL.NEW_RULE",
      "XLSCTL.HOME.STYLE.CONDITIONAL.CLEAR_SELECTION",
      "XLSCTL.HOME.STYLE.CONDITIONAL.CLEAR_SHEET",
      "XLSCTL.HOME.STYLE.CONDITIONAL.MANAGE_RULES",
    ],
  }),
  ...[
    ["XLSCTL.HOME.STYLE.CONDITIONAL.NEW_RULE", "openConditionalNewRule", 45],
    ["XLSCTL.HOME.STYLE.CONDITIONAL.MANAGE_RULES", "openConditionalManager", 46],
  ].map(([controlId, suffix, order]) =>
    defineHomeCoreBound({
      controlId: String(controlId),
      featureId: `spreadsheet.home.styles.${String(suffix)}`,
      canonicalId: "XLS-HOM-STYLE-01",
      order: Number(order),
      controlKind: "dialog-launcher",
      predicate: "conditional-format-editable",
      reasonKey: "spreadsheet.capability.selectionRequired",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
  ),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.STYLE.CELL_STYLE_APPLY",
    featureId: "spreadsheet.home.styles.cellStyle",
    canonicalId: "XLS-HOM-STYLE-03",
    order: 49,
    controlKind: "menu",
    predicate: "cell-style-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-cells",
  }),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.STYLE.CELL_STYLE_MENU",
    featureId: "spreadsheet.home.styles.cellStyleMenu",
    canonicalId: "XLS-HOM-STYLE-03",
    order: 49,
    children: ["XLSCTL.HOME.STYLE.CELL_STYLE_APPLY"],
  }),
  ...[
    [
      "XLSCTL.HOME.STYLE.CONDITIONAL.CLEAR_SELECTION",
      "clearConditionalSelection",
      47,
    ],
    [
      "XLSCTL.HOME.STYLE.CONDITIONAL.CLEAR_SHEET",
      "clearConditionalSheet",
      48,
    ],
  ].map(([controlId, suffix, order]) =>
    defineHomeCoreBound({
      controlId: String(controlId),
      featureId: `spreadsheet.home.styles.${String(suffix)}`,
      canonicalId: "XLS-HOM-STYLE-01",
      order: Number(order),
      controlKind: "button",
      predicate:
        controlId === "XLSCTL.HOME.STYLE.CONDITIONAL.CLEAR_SELECTION"
          ? "conditional-format-selection-clearable"
          : "conditional-format-sheet-clearable",
      reasonKey: "spreadsheet.capability.sheetProtected",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-conditional-formats",
    }),
  ),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.CELLS.INSERT_MENU",
    featureId: "spreadsheet.home.cells.insertMenu",
    canonicalId: "XLS-HOM-CELL-01",
    order: 47,
    // The menu also carries Insert Sheet, which worksheet protection does not
    // block, so it cannot follow its cell/row/column children: on a protected
    // sheet Excel opens the menu, greys the cell items per the protection
    // options and leaves Insert Sheet live. Only Protect Workbook ▸ Structure
    // closes it (Microsoft Support, "Protect a worksheet").
    predicate: "workbook-sheet-structure-editable",
    children: [
      "XLSCTL.HOME.CELLS.INSERT.SHIFT_DOWN",
      "XLSCTL.HOME.CELLS.INSERT.SHIFT_RIGHT",
      "XLSCTL.HOME.CELLS.INSERT.ROW",
      "XLSCTL.HOME.CELLS.INSERT.COLUMN",
    ],
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.CELLS.INSERT.SHIFT_DOWN",
    featureId: "spreadsheet.home.cells.insertShiftDown",
    canonicalId: "XLS-HOM-CELL-01",
    order: 48,
    controlKind: "button",
    predicate: "cell-shift-down-editable",
    reasonKey: "spreadsheet.capability.cellShiftUnsafe",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-structure",
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.CELLS.INSERT.SHIFT_RIGHT",
    featureId: "spreadsheet.home.cells.insertShiftRight",
    canonicalId: "XLS-HOM-CELL-01",
    order: 49,
    controlKind: "button",
    predicate: "cell-shift-right-editable",
    reasonKey: "spreadsheet.capability.cellShiftUnsafe",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-structure",
  }),
  ...[
    ["XLSCTL.HOME.CELLS.INSERT.ROW", "insertRow", 48, "worksheet-insert-rows-editable"],
    ["XLSCTL.HOME.CELLS.INSERT.COLUMN", "insertColumn", 49, "worksheet-insert-columns-editable"],
  ].map(([controlId, suffix, order, predicate]) =>
    defineHomeCoreBound({
      controlId: String(controlId),
      featureId: `spreadsheet.home.cells.${String(suffix)}`,
      canonicalId: "XLS-HOM-CELL-01",
      order: Number(order),
      controlKind: "button",
      predicate: predicate as SpreadsheetVisiblePredicate,
      reasonKey: "spreadsheet.capability.sheetProtected",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-structure",
    }),
  ),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.CELLS.DELETE_MENU",
    featureId: "spreadsheet.home.cells.deleteMenu",
    canonicalId: "XLS-HOM-CELL-01",
    order: 50,
    // Delete Sheet, like Insert Sheet above, answers to workbook structure
    // protection rather than to worksheet protection.
    predicate: "workbook-sheet-structure-editable",
    children: [
      "XLSCTL.HOME.CELLS.DELETE.SHIFT_UP",
      "XLSCTL.HOME.CELLS.DELETE.SHIFT_LEFT",
      "XLSCTL.HOME.CELLS.DELETE.ROW",
      "XLSCTL.HOME.CELLS.DELETE.COLUMN",
    ],
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.CELLS.DELETE.SHIFT_UP",
    featureId: "spreadsheet.home.cells.deleteShiftUp",
    canonicalId: "XLS-HOM-CELL-01",
    order: 51,
    controlKind: "button",
    predicate: "cell-shift-up-editable",
    reasonKey: "spreadsheet.capability.cellShiftUnsafe",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-structure",
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.CELLS.DELETE.SHIFT_LEFT",
    featureId: "spreadsheet.home.cells.deleteShiftLeft",
    canonicalId: "XLS-HOM-CELL-01",
    order: 52,
    controlKind: "button",
    predicate: "cell-shift-left-editable",
    reasonKey: "spreadsheet.capability.cellShiftUnsafe",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-structure",
  }),
  ...[
    ["XLSCTL.HOME.CELLS.DELETE.ROW", "deleteRow", 51, "worksheet-delete-rows-editable"],
    ["XLSCTL.HOME.CELLS.DELETE.COLUMN", "deleteColumn", 52, "worksheet-delete-columns-editable"],
  ].map(([controlId, suffix, order, predicate]) =>
    defineHomeCoreBound({
      controlId: String(controlId),
      featureId: `spreadsheet.home.cells.${String(suffix)}`,
      canonicalId: "XLS-HOM-CELL-01",
      order: Number(order),
      controlKind: "button",
      predicate: predicate as SpreadsheetVisiblePredicate,
      reasonKey: "spreadsheet.capability.sheetProtected",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-structure",
    }),
  ),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.CELLS.FORMAT_MENU",
    featureId: "spreadsheet.home.cells.formatMenu",
    canonicalId: "XLS-HOM-CELL-02",
    order: 53,
    children: ["XLSCTL.HOME.CELLS.FORMAT.FORMAT_CELLS"],
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.CELLS.FORMAT.FORMAT_CELLS",
    featureId: "spreadsheet.home.cells.openFormatCells",
    canonicalId: "XLS-HOM-CELL-02",
    order: 54,
    controlKind: "dialog-launcher",
    predicate: "cell-selection",
    reasonKey: "spreadsheet.capability.selectionRequired",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.FORMAT_CELLS.DIALOG",
    featureId: "spreadsheet.home.cells.formatCellsDialog",
    canonicalId: "XLS-HOM-CELL-02",
    order: 55,
    children: [
      "XLSCTL.HOME.FORMAT_CELLS.OK",
      "XLSCTL.HOME.FORMAT_CELLS.CANCEL",
    ],
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.FORMAT_CELLS.OK",
    featureId: "spreadsheet.home.cells.applyFormatCells",
    canonicalId: "XLS-HOM-CELL-02",
    order: 56,
    controlKind: "button",
    predicate: "format-cells-apply",
    reasonKey: "spreadsheet.capability.writerUnavailable",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-cells",
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.FORMAT_CELLS.CANCEL",
    featureId: "spreadsheet.home.cells.closeFormatCells",
    canonicalId: "XLS-HOM-CELL-02",
    order: 57,
    controlKind: "button",
    predicate: "always",
    reasonKey: null,
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.EDIT.FILL_MENU",
    featureId: "spreadsheet.home.fill.menu",
    canonicalId: "XLS-HOM-EDIT-02",
    order: 55,
    children: [
      "XLSCTL.HOME.EDIT.FILL.DOWN",
      "XLSCTL.HOME.EDIT.FILL.RIGHT",
      "XLSCTL.HOME.EDIT.FILL.UP",
      "XLSCTL.HOME.EDIT.FILL.LEFT",
      "XLSCTL.HOME.EDIT.FILL.SERIES",
      "XLSCTL.HOME.EDIT.FILL.FLASH_FILL",
    ],
  }),
  ...[
    ["XLSCTL.HOME.EDIT.FILL.DOWN", "down", 56],
    ["XLSCTL.HOME.EDIT.FILL.RIGHT", "right", 57],
    ["XLSCTL.HOME.EDIT.FILL.UP", "up", 58],
    ["XLSCTL.HOME.EDIT.FILL.LEFT", "left", 59],
  ].map(([controlId, suffix, order]) =>
    defineHomeCoreBound({
      controlId: String(controlId),
      featureId: `spreadsheet.home.fill.${String(suffix)}`,
      canonicalId: "XLS-HOM-EDIT-02",
      order: Number(order),
      controlKind: "button",
      predicate: "selection-editable",
      reasonKey: "spreadsheet.capability.sheetProtected",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-cells",
    }),
  ),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.EDIT.FILL.FLASH_FILL",
    featureId: "spreadsheet.data.flashFill",
    canonicalId: "XLS-DAT-TOOLS-01",
    order: 62,
    controlKind: "button",
    predicate: "selection-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-cells",
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.EDIT.FILL.SERIES",
    featureId: "spreadsheet.home.fill.openSeries",
    canonicalId: "XLS-HOM-EDIT-02",
    order: 60,
    controlKind: "dialog-launcher",
    predicate: "selection-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.EDIT.FIND_SELECT_MENU",
    featureId: "spreadsheet.home.find.menu",
    canonicalId: "XLS-HOM-EDIT-05",
    order: 63,
    children: [
      "XLSCTL.HOME.EDIT.FIND_SELECT.FIND",
      "XLSCTL.HOME.EDIT.FIND_SELECT.REPLACE",
      "XLSCTL.HOME.EDIT.FIND_SELECT.GO_TO",
      "XLSCTL.HOME.EDIT.FIND_SELECT.GO_TO_SPECIAL",
    ],
  }),
  ...[
    ["XLSCTL.HOME.EDIT.FIND_SELECT.FIND", "openFind", 64],
    ["XLSCTL.HOME.EDIT.FIND_SELECT.REPLACE", "openReplace", 65],
    ["XLSCTL.HOME.EDIT.FIND_SELECT.GO_TO", "openGoTo", 66],
    ["XLSCTL.HOME.EDIT.FIND_SELECT.GO_TO_SPECIAL", "openGoToSpecial", 67],
  ].map(([controlId, suffix, order]) =>
    defineHomeCoreBound({
      controlId: String(controlId),
      featureId: `spreadsheet.home.find.${String(suffix)}`,
      canonicalId: "XLS-HOM-EDIT-05",
      order: Number(order),
      controlKind: "dialog-launcher",
      predicate: "always",
      reasonKey: null,
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
  ),
  defineHomeCoreDisclosure({
    controlId: "XLSCTL.HOME.FIND_REPLACE.DIALOG",
    featureId: "spreadsheet.home.find.dialog",
    canonicalId: "XLS-HOM-EDIT-05",
    order: 68,
    children: [
      "XLSCTL.HOME.FIND_REPLACE.FIND_PREVIOUS",
      "XLSCTL.HOME.FIND_REPLACE.FIND_NEXT",
      "XLSCTL.HOME.FIND_REPLACE.REPLACE",
      "XLSCTL.HOME.FIND_REPLACE.REPLACE_ALL",
      "XLSCTL.HOME.FIND_REPLACE.MRAI.DISMISS",
      "XLSCTL.HOME.FIND_REPLACE.FIND_ALL",
    ],
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.FIND_REPLACE.FIND_PREVIOUS",
    featureId: "spreadsheet.home.find.findPrevious",
    canonicalId: "XLS-HOM-EDIT-05",
    order: 69,
    controlKind: "button",
    predicate: "find-replace-match",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.FIND_REPLACE.FIND_NEXT",
    featureId: "spreadsheet.home.find.findNext",
    canonicalId: "XLS-HOM-EDIT-05",
    order: 70,
    controlKind: "button",
    predicate: "find-replace-match",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.FIND_REPLACE.REPLACE",
    featureId: "spreadsheet.home.find.replaceCurrent",
    canonicalId: "XLS-HOM-EDIT-05",
    order: 71,
    controlKind: "button",
    predicate: "find-replace-current-editable",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-cells",
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.FIND_REPLACE.REPLACE_ALL",
    featureId: "spreadsheet.home.find.replaceAll",
    canonicalId: "XLS-HOM-EDIT-05",
    order: 72,
    controlKind: "button",
    predicate: "find-replace-all-editable",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-cells",
  }),
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.FIND_REPLACE.MRAI.DISMISS",
    featureId: "spreadsheet.home.find.closeDialog",
    canonicalId: "XLS-HOM-EDIT-05",
    order: 73,
    controlKind: "button",
    predicate: "always",
    reasonKey: null,
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  // Excel's Find All: lists every hit under the dialog; reads, never writes.
  defineHomeCoreBound({
    controlId: "XLSCTL.HOME.FIND_REPLACE.FIND_ALL",
    featureId: "spreadsheet.home.find.findAll",
    canonicalId: "XLS-HOM-EDIT-05",
    order: 74,
    controlKind: "button",
    predicate: "find-replace-match",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

const definePageSetupPrintBound = (config: {
  readonly controlId: string;
  readonly featureId: string;
  readonly canonicalId?: "XLS-PGL-SETUP-01" | "XLS-PGL-BREAK-01";
  readonly order: number;
  readonly controlKind: SpreadsheetCapabilityControlKind;
  readonly predicate: SpreadsheetVisiblePredicate;
  readonly undoPolicy: "single-transaction" | "content-none";
  readonly ioProfile: "session-only" | "worksheet-print-settings";
}): SpreadsheetVisibleControlDefinition =>
  define({
    ...config,
    canonicalId: config.canonicalId ?? "XLS-PGL-SETUP-01",
    group: "Page Layout",
    reasonKey:
      config.predicate === "always"
        ? null
        : config.predicate === "page-setup-editable"
          ? "spreadsheet.capability.sheetProtected"
          : "spreadsheet.capability.commandUnavailable",
    command: "bound",
  });

const definePageSetupPrintDisclosure = (config: {
  readonly controlId: string;
  readonly featureId: string;
  readonly canonicalId?: "XLS-PGL-SETUP-01" | "XLS-PGL-BREAK-01";
  readonly order: number;
  readonly children: readonly string[];
}): SpreadsheetVisibleControlDefinition =>
  define({
    ...config,
    canonicalId: config.canonicalId ?? "XLS-PGL-SETUP-01",
    group: "Page Layout",
    controlKind: "menu",
    predicate: "menu-children",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    command: "disclosure",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  });

/**
 * Bounded Page Setup / Print authority. Menu roots and field launchers are
 * session-only; the nine persistent leaves are the only commands that own a
 * `printSettings` write transaction. `GRIDLINES.PRINT` joins them: Excel for
 * the web cannot print gridlines, but the product target is desktop Excel,
 * whose Sheet Options group prints them (`printOptions@gridLines`).
 */
export const SPREADSHEET_PAGE_SETUP_PRINT_VERTICAL_CONTROL_DEFINITIONS =
  Object.freeze([
    definePageSetupPrintBound({
      controlId: "XLSCTL.PAGE_LAYOUT.PAGE_SETUP.DIALOG",
      featureId: "spreadsheet.pageLayout.openPageSetup",
      order: 1,
      controlKind: "dialog-launcher",
      predicate: "page-setup-launchable",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    definePageSetupPrintDisclosure({
      controlId: "XLSCTL.PAGE_LAYOUT.MARGINS",
      featureId: "spreadsheet.pageLayout.marginsMenu",
      order: 2,
      children: [
        "XLSCTL.PAGE_LAYOUT.MARGINS.APPLY",
        "XLSCTL.PAGE_LAYOUT.MARGINS.CUSTOM",
      ],
    }),
    definePageSetupPrintBound({
      controlId: "XLSCTL.PAGE_LAYOUT.MARGINS.APPLY",
      featureId: "spreadsheet.pageLayout.applyMarginPreset",
      order: 3,
      controlKind: "gallery",
      predicate: "page-margin-applicable",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-print-settings",
    }),
    definePageSetupPrintBound({
      controlId: "XLSCTL.PAGE_LAYOUT.MARGINS.CUSTOM",
      featureId: "spreadsheet.pageLayout.openCustomMargins",
      order: 4,
      controlKind: "dialog-launcher",
      predicate: "page-setup-editable",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    ...[
      ["XLSCTL.PAGE_LAYOUT.ORIENTATION", "openOrientation", 5],
      ["XLSCTL.PAGE_LAYOUT.SIZE", "openPaperSize", 6],
      ["XLSCTL.PAGE_LAYOUT.PRINT_AREA", "openPrintArea", 7],
    ].map(([controlId, suffix, order]) =>
      definePageSetupPrintBound({
        controlId: String(controlId),
        featureId: `spreadsheet.pageLayout.${String(suffix)}`,
        order: Number(order),
        controlKind: "dialog-launcher",
        predicate: "page-setup-launchable",
        undoPolicy: "content-none",
        ioProfile: "session-only",
      }),
    ),
    definePageSetupPrintDisclosure({
      controlId: "XLSCTL.PAGE_LAYOUT.PRINT_AREA.DISCLOSURE",
      featureId: "spreadsheet.pageLayout.printAreaMenu",
      order: 8,
      children: [
        "XLSCTL.PAGE_LAYOUT.PRINT_AREA.SET",
        "XLSCTL.PAGE_LAYOUT.PRINT_AREA.CLEAR",
        "XLSCTL.PAGE_LAYOUT.PRINT_AREA.ADD",
        "XLSCTL.PAGE_LAYOUT.PRINT_AREA.EDIT",
      ],
    }),
    definePageSetupPrintBound({
      controlId: "XLSCTL.PAGE_LAYOUT.PRINT_AREA.SET",
      featureId: "spreadsheet.pageLayout.setPrintArea",
      order: 9,
      controlKind: "button",
      predicate: "print-area-settable",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-print-settings",
    }),
    definePageSetupPrintBound({
      controlId: "XLSCTL.PAGE_LAYOUT.PRINT_AREA.CLEAR",
      featureId: "spreadsheet.pageLayout.clearPrintArea",
      order: 10,
      controlKind: "button",
      predicate: "print-area-clearable",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-print-settings",
    }),
    // Excel "Set or clear a print area on a worksheet": Add to Print Area
    // appends the selection as a further area (each area prints on its own
    // page); it is offered once an area exists and the selection is outside.
    definePageSetupPrintBound({
      controlId: "XLSCTL.PAGE_LAYOUT.PRINT_AREA.ADD",
      featureId: "spreadsheet.pageLayout.addToPrintArea",
      order: 11,
      controlKind: "button",
      predicate: "print-area-addable",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-print-settings",
    }),
    definePageSetupPrintBound({
      controlId: "XLSCTL.PAGE_LAYOUT.PRINT_AREA.EDIT",
      featureId: "spreadsheet.pageLayout.openPrintArea",
      order: 12,
      controlKind: "dialog-launcher",
      predicate: "page-setup-launchable",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    definePageSetupPrintDisclosure({
      controlId: "XLSCTL.PAGE_LAYOUT.BREAKS",
      featureId: "spreadsheet.pageLayout.breaksMenu",
      canonicalId: "XLS-PGL-BREAK-01",
      order: 12,
      children: [
        "XLSCTL.PAGE_LAYOUT.BREAKS.INSERT",
        "XLSCTL.PAGE_LAYOUT.BREAKS.REMOVE",
        "XLSCTL.PAGE_LAYOUT.BREAKS.RESET",
        "XLSCTL.PAGE_LAYOUT.BREAKS.PREVIEW",
      ],
    }),
    ...[
      ["XLSCTL.PAGE_LAYOUT.BREAKS.INSERT", "insertPageBreak", 13, "page-break-insertable"],
      ["XLSCTL.PAGE_LAYOUT.BREAKS.REMOVE", "removePageBreak", 14, "page-break-removable"],
      ["XLSCTL.PAGE_LAYOUT.BREAKS.RESET", "resetPageBreaks", 15, "page-break-resettable"],
    ].map(([controlId, suffix, order, predicate]) =>
      definePageSetupPrintBound({
        controlId: String(controlId),
        featureId: `spreadsheet.pageLayout.${String(suffix)}`,
        canonicalId: "XLS-PGL-BREAK-01",
        order: Number(order),
        controlKind: "button",
        predicate: predicate as SpreadsheetVisiblePredicate,
        undoPolicy: "single-transaction",
        ioProfile: "worksheet-print-settings",
      }),
    ),
    definePageSetupPrintBound({
      controlId: "XLSCTL.PAGE_LAYOUT.BREAKS.PREVIEW",
      featureId: "spreadsheet.pageLayout.togglePageBreakPreview",
      canonicalId: "XLS-PGL-BREAK-01",
      order: 16,
      controlKind: "toggle",
      predicate: "always",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    ...[
      ["XLSCTL.PAGE_LAYOUT.PRINT_TITLES", "openPrintTitles", 17],
      ["XLSCTL.PAGE_LAYOUT.SCALE.WIDTH", "openScaleWidth", 18],
      ["XLSCTL.PAGE_LAYOUT.SCALE.HEIGHT", "openScaleHeight", 19],
      ["XLSCTL.PAGE_LAYOUT.SCALE.DIALOG", "openScaleSettings", 20],
    ].map(([controlId, suffix, order]) =>
      definePageSetupPrintBound({
        controlId: String(controlId),
        featureId: `spreadsheet.pageLayout.${String(suffix)}`,
        order: Number(order),
        controlKind: "dialog-launcher",
        predicate: "page-setup-launchable",
        undoPolicy: "content-none",
        ioProfile: "session-only",
      }),
    ),
    definePageSetupPrintBound({
      controlId: "XLSCTL.PAGE_LAYOUT.HEADINGS.PRINT",
      featureId: "spreadsheet.pageLayout.togglePrintHeadings",
      order: 21,
      controlKind: "toggle",
      predicate: "page-setup-editable",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-print-settings",
    }),
    definePageSetupPrintBound({
      controlId: "XLSCTL.PAGE_LAYOUT.GRIDLINES.PRINT",
      featureId: "spreadsheet.pageLayout.togglePrintGridlines",
      order: 22,
      controlKind: "toggle",
      predicate: "page-setup-editable",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-print-settings",
    }),
    definePageSetupPrintBound({
      controlId: "XLSCTL.PAGE_LAYOUT.PAGE_SETUP.APPLY",
      featureId: "spreadsheet.pageLayout.applyPageSetup",
      order: 22,
      controlKind: "button",
      predicate: "page-setup-editable",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-print-settings",
    }),
    definePageSetupPrintBound({
      controlId: "XLSCTL.PAGE_LAYOUT.PAGE_SETUP.CLEAR",
      featureId: "spreadsheet.pageLayout.clearPageSetup",
      order: 23,
      controlKind: "button",
      predicate: "page-setup-clearable",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-print-settings",
    }),
  ] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

const defineFormulasBound = (config: {
  readonly controlId: string;
  readonly featureId: string;
  readonly canonicalId:
    | "XLS-FRM-INSERT-01"
    | "XLS-FRM-FUNC-01"
    | "XLS-FRM-NAME-01"
    | "XLS-FRM-AUDIT-01"
    | "XLS-FRM-CALC-01";
  readonly order: number;
  readonly controlKind: SpreadsheetCapabilityControlKind;
  readonly predicate: SpreadsheetVisiblePredicate;
  readonly reasonKey?: SpreadsheetVisibleCommandReason | null;
  readonly undoPolicy: "single-transaction" | "content-none";
  readonly ioProfile: SpreadsheetVisibleIoProfile;
}): SpreadsheetVisibleControlDefinition =>
  define({
    ...config,
    group: "Formulas",
    reasonKey: config.reasonKey ?? null,
    command: "bound",
  });

const defineFormulasDisclosure = (config: {
  readonly controlId: string;
  readonly featureId: string;
  readonly canonicalId:
    | "XLS-FRM-INSERT-01"
    | "XLS-FRM-FUNC-01"
    | "XLS-FRM-NAME-01"
    | "XLS-FRM-CALC-01";
  readonly order: number;
  readonly children: readonly string[];
}): SpreadsheetVisibleControlDefinition =>
  define({
    ...config,
    group: "Formulas",
    controlKind: "menu",
    predicate: "menu-children",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    command: "disclosure",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  });

/**
 * High-frequency Formulas vertical. Gallery/category controls and audit tools
 * own session state only; AutoSum, Function Arguments and Defined Names own
 * the exact confirmed write. Parameterized leaves consume a host receipt and
 * are not expanded into invented per-value control IDs.
 */
export const SPREADSHEET_FORMULAS_VERTICAL_CONTROL_DEFINITIONS = Object.freeze([
  defineFormulasBound({
    controlId: "XLSCTL.FORMULAS.INSERT_FUNCTION",
    featureId: "spreadsheet.formulas.openInsertFunction",
    canonicalId: "XLS-FRM-INSERT-01",
    order: 1,
    controlKind: "dialog-launcher",
    predicate: "selection",
    reasonKey: "spreadsheet.capability.selectionRequired",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  defineFormulasBound({
    controlId: "XLSCTL.FORMULAS.FUNCTION_ARGUMENTS.APPLY",
    featureId: "spreadsheet.formulas.applyFunctionArguments",
    canonicalId: "XLS-FRM-INSERT-01",
    order: 2,
    controlKind: "button",
    predicate: "selection-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-cells",
  }),
  defineFormulasDisclosure({
    controlId: "XLSCTL.FORMULAS.AUTOSUM_MENU",
    featureId: "spreadsheet.formulas.autoSumMenu",
    canonicalId: "XLS-FRM-INSERT-01",
    order: 3,
    children: [
      "XLSCTL.FORMULAS.AUTOSUM.APPLY",
      "XLSCTL.FORMULAS.INSERT_FUNCTION",
    ],
  }),
  defineFormulasBound({
    controlId: "XLSCTL.FORMULAS.AUTOSUM.APPLY",
    featureId: "spreadsheet.formulas.applyAutoSum",
    canonicalId: "XLS-FRM-INSERT-01",
    order: 4,
    controlKind: "gallery",
    predicate: "cell-selection-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-cells",
  }),
  ...[
    ["XLSCTL.FORMULAS.LIBRARY.FINANCIAL", "financial", 5],
    ["XLSCTL.FORMULAS.LIBRARY.LOGICAL", "logical", 6],
    ["XLSCTL.FORMULAS.LIBRARY.TEXT", "text", 7],
    ["XLSCTL.FORMULAS.LIBRARY.DATE_TIME", "dateTime", 8],
    ["XLSCTL.FORMULAS.LIBRARY.LOOKUP_REFERENCE", "lookupReference", 9],
    ["XLSCTL.FORMULAS.LIBRARY.MATH_TRIG", "mathTrig", 10],
  ].map(([controlId, suffix, order]) =>
    defineFormulasBound({
      controlId: String(controlId),
      featureId: `spreadsheet.formulas.openFunctionCategory.${String(suffix)}`,
      canonicalId: "XLS-FRM-FUNC-01",
      order: Number(order),
      controlKind: "gallery",
      predicate: "selection",
      reasonKey: "spreadsheet.capability.selectionRequired",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
  ),
  defineFormulasDisclosure({
    controlId: "XLSCTL.FORMULAS.LIBRARY.MORE",
    featureId: "spreadsheet.formulas.moreFunctionsMenu",
    canonicalId: "XLS-FRM-FUNC-01",
    order: 11,
    children: [
      "XLSCTL.FORMULAS.LIBRARY.MORE.CATEGORY",
      "XLSCTL.FORMULAS.LIBRARY.ALL_FUNCTIONS",
    ],
  }),
  defineFormulasBound({
    controlId: "XLSCTL.FORMULAS.LIBRARY.MORE.CATEGORY",
    featureId: "spreadsheet.formulas.openMoreFunctionCategory",
    canonicalId: "XLS-FRM-FUNC-01",
    order: 12,
    controlKind: "gallery",
    predicate: "selection",
    reasonKey: "spreadsheet.capability.selectionRequired",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  defineFormulasBound({
    controlId: "XLSCTL.FORMULAS.LIBRARY.ALL_FUNCTIONS",
    featureId: "spreadsheet.formulas.openAllFunctions",
    canonicalId: "XLS-FRM-FUNC-01",
    order: 13,
    controlKind: "gallery",
    predicate: "selection",
    reasonKey: "spreadsheet.capability.selectionRequired",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  defineFormulasBound({
    controlId: "XLSCTL.FORMULAS.NAME.MANAGER",
    featureId: "spreadsheet.formulas.openNameManager",
    canonicalId: "XLS-FRM-NAME-01",
    order: 14,
    controlKind: "button",
    predicate: "always",
    reasonKey: null,
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  defineFormulasDisclosure({
    controlId: "XLSCTL.FORMULAS.NAME.DEFINE_MENU",
    featureId: "spreadsheet.formulas.nameMenu",
    canonicalId: "XLS-FRM-NAME-01",
    order: 14,
    children: [
      "XLSCTL.FORMULAS.NAME.DEFINE.APPLY",
      "XLSCTL.FORMULAS.NAME.DELETE.APPLY",
    ],
  }),
  defineFormulasBound({
    controlId: "XLSCTL.FORMULAS.NAME.DEFINE.APPLY",
    featureId: "spreadsheet.formulas.defineName",
    canonicalId: "XLS-FRM-NAME-01",
    order: 15,
    controlKind: "button",
    predicate: "selection-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    undoPolicy: "single-transaction",
    ioProfile: "workbook-defined-names",
  }),
  defineFormulasBound({
    controlId: "XLSCTL.FORMULAS.NAME.DELETE.APPLY",
    featureId: "spreadsheet.formulas.deleteName",
    canonicalId: "XLS-FRM-NAME-01",
    order: 16,
    controlKind: "button",
    predicate: "selection-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    undoPolicy: "single-transaction",
    ioProfile: "workbook-defined-names",
  }),
  defineFormulasBound({
    controlId: "XLSCTL.FORMULAS.NAME.MANAGER.APPLY",
    featureId: "spreadsheet.formulas.applyNameManager",
    canonicalId: "XLS-FRM-NAME-01",
    order: 140,
    controlKind: "button",
    predicate: "selection-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    undoPolicy: "single-transaction",
    ioProfile: "workbook-defined-names",
  }),
  defineFormulasBound({
    controlId: "XLSCTL.FORMULAS.NAME.MANAGER.DELETE",
    featureId: "spreadsheet.formulas.deleteManagedName",
    canonicalId: "XLS-FRM-NAME-01",
    order: 141,
    controlKind: "button",
    predicate: "selection-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    undoPolicy: "single-transaction",
    ioProfile: "workbook-defined-names",
  }),
  ...[
    [
      "XLSCTL.FORMULAS.AUDIT.SHOW_FORMULAS",
      "toggleShowFormulas",
      17,
      "always",
    ],
    [
      "XLSCTL.FORMULAS.AUDIT.PRECEDENTS",
      "tracePrecedents",
      18,
      "selection",
    ],
    [
      "XLSCTL.FORMULAS.AUDIT.DEPENDENTS",
      "traceDependents",
      19,
      "selection",
    ],
    [
      "XLSCTL.FORMULAS.AUDIT.REMOVE_ARROWS_MENU",
      "removeTraceArrows",
      20,
      "formula-trace-clear",
    ],
    [
      "XLSCTL.FORMULAS.AUDIT.ERROR_CHECK_MENU",
      "toggleErrorCheck",
      21,
      "always",
    ],
    [
      "XLSCTL.FORMULAS.EVALUATE.DIALOG",
      "toggleEvaluateFormula",
      22,
      "formula-evaluation",
    ],
    [
      "XLSCTL.FORMULAS.REFERENCE_STYLE",
      "toggleReferenceStyle",
      23,
      "always",
    ],
  ].map(([controlId, suffix, order, predicate]) =>
    defineFormulasBound({
      controlId: String(controlId),
      featureId: `spreadsheet.formulas.${String(suffix)}`,
      canonicalId: "XLS-FRM-AUDIT-01",
      order: Number(order),
      controlKind:
        String(controlId).includes("SHOW_FORMULAS") ||
        String(controlId).includes("REFERENCE_STYLE")
          ? "toggle"
          : "button",
      predicate: predicate as SpreadsheetVisiblePredicate,
      reasonKey: "spreadsheet.capability.commandUnavailable",
      undoPolicy: "content-none",
      ioProfile:
        String(controlId).includes("SHOW_FORMULAS") ||
        String(controlId).includes("REFERENCE_STYLE")
          ? "workbook-view-preference"
          : "session-only",
    }),
  ),
  defineFormulasDisclosure({
    controlId: "XLSCTL.FORMULAS.CALC.OPTIONS_MENU",
    featureId: "spreadsheet.formulas.calculationOptionsMenu",
    canonicalId: "XLS-FRM-CALC-01",
    order: 24,
    children: [
      "XLSCTL.FORMULAS.CALC.OPTIONS.AUTOMATIC",
      "XLSCTL.FORMULAS.CALC.OPTIONS.MANUAL",
      "XLSCTL.FORMULAS.CALC.OPTIONS.DATE1904",
    ],
  }),
  ...[
    ["XLSCTL.FORMULAS.CALC.OPTIONS.AUTOMATIC", "setCalculationAutomatic", 25],
    ["XLSCTL.FORMULAS.CALC.OPTIONS.MANUAL", "setCalculationManual", 26],
  ].map(([controlId, suffix, order]) =>
    defineFormulasBound({
      controlId: String(controlId),
      featureId: `spreadsheet.formulas.${String(suffix)}`,
      canonicalId: "XLS-FRM-CALC-01",
      order: Number(order),
      controlKind: "toggle",
      // Manual mode captures the workbook's calculation first (PivotTables
      // included); until the formula coordinates that capture needs exist,
      // the menu item says so rather than silently staying on Automatic.
      predicate: suffix === "setCalculationManual" ? "manual-calculation" : "always",
      reasonKey:
        suffix === "setCalculationManual"
          ? "spreadsheet.capability.manualCalculationUnavailable"
          : null,
      undoPolicy: "content-none",
      ioProfile: "workbook-calculation-mode",
    }),
  ),
  // Excel's "Use 1904 date system" (Options > Advanced > When calculating this
  // workbook; Excel for Mac: Preferences > Calculation, beside Automatic and
  // Manual). A workbook setting that changes how serials read as dates, never
  // the serials, so it is a toggle in Calculation Options.
  defineFormulasBound({
    controlId: "XLSCTL.FORMULAS.CALC.OPTIONS.DATE1904",
    featureId: "spreadsheet.formulas.toggleDate1904",
    canonicalId: "XLS-FRM-CALC-01",
    order: 29,
    controlKind: "toggle",
    predicate: "always",
    undoPolicy: "content-none",
    ioProfile: "workbook-date-system",
  }),
  ...[
    ["XLSCTL.FORMULAS.CALC.CALCULATE_NOW", "calculateNow", 27],
    ["XLSCTL.FORMULAS.CALC.CALCULATE_SHEET", "calculateSheet", 28],
  ].map(([controlId, suffix, order]) =>
    defineFormulasBound({
      controlId: String(controlId),
      featureId: `spreadsheet.formulas.${String(suffix)}`,
      canonicalId: "XLS-FRM-CALC-01",
      order: Number(order),
      controlKind: "button",
      predicate: "always",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
  ),
] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

/**
 * Insert is intentionally split between session-only launchers/disclosures and
 * the exact leaves that create workbook content.  Object/chart/pivot/sparkline
 * writers are separately guarded by their native route; an imported workbook
 * must never fall back to a from-scratch rewrite merely because an Insert
 * control was reachable.
 */
export const SPREADSHEET_INSERT_VERTICAL_CONTROL_DEFINITIONS = Object.freeze([
  define({
    controlId: "XLSCTL.INSERT.PIVOT",
    featureId: "spreadsheet.insert.createPivot",
    canonicalId: "XLS-INS-PIVOT-01",
    group: "Insert",
    order: 1,
    controlKind: "dialog-launcher",
    predicate: "selection",
    reasonKey: "spreadsheet.capability.selectionRequired",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.INSERT.PIVOT.APPLY",
    featureId: "spreadsheet.insert.applyPivot",
    canonicalId: "XLS-INS-PIVOT-01",
    group: "Insert",
    order: 2,
    controlKind: "button",
    predicate: "pivot-insertable",
    reasonKey: "spreadsheet.capability.selectionRequired",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-pivots",
  }),
  define({
    controlId: "XLSCTL.INSERT.PIVOT.DISCLOSURE",
    featureId: "spreadsheet.insert.pivotMenu",
    canonicalId: "XLS-INS-PIVOT-01",
    group: "Insert",
    order: 3,
    controlKind: "menu",
    predicate: "menu-children",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    command: "disclosure",
    children: [
      "XLSCTL.INSERT.PIVOT",
      "XLSCTL.INSERT.RECOMMENDED_PIVOT",
    ],
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.INSERT.RECOMMENDED_PIVOT",
    featureId: "spreadsheet.insert.openRecommendedPivot",
    canonicalId: "XLS-INS-PIVOT-02",
    group: "Insert",
    order: 4,
    controlKind: "dialog-launcher",
    predicate: "selection",
    reasonKey: "spreadsheet.capability.selectionRequired",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.INSERT.RECOMMENDED_PIVOT.APPLY",
    featureId: "spreadsheet.insert.applyRecommendedPivot",
    canonicalId: "XLS-INS-PIVOT-02",
    group: "Insert",
    order: 5,
    controlKind: "button",
    predicate: "pivot-insertable",
    reasonKey: "spreadsheet.capability.selectionRequired",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-pivots",
  }),
  define({
    controlId: "XLSCTL.INSERT.PICTURES.MENU",
    featureId: "spreadsheet.insert.insertPicture",
    canonicalId: "XLS-INS-IMAGE-01",
    group: "Insert",
    order: 5,
    controlKind: "button",
    predicate: "object-insertable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-drawing-objects",
  }),
  define({
    controlId: "XLSCTL.INSERT.PICTURES.MENU.DISCLOSURE",
    featureId: "spreadsheet.insert.pictureMenu",
    canonicalId: "XLS-INS-IMAGE-01",
    group: "Insert",
    order: 6,
    controlKind: "menu",
    predicate: "menu-children",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    command: "disclosure",
    children: [
      "XLSCTL.INSERT.PICTURES.PLACE_IN_CELLS",
      "XLSCTL.INSERT.PICTURES.OVER_CELLS.MENU",
    ],
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    // Excel's Place in Cell: the chosen pictures become the values of the
    // active cell and the cells below it, one worksheet transaction.
    controlId: "XLSCTL.INSERT.PICTURES.PLACE_IN_CELLS",
    featureId: "spreadsheet.insert.pictureInCell",
    canonicalId: "XLS-INS-IMAGE-01",
    group: "Insert",
    order: 7,
    controlKind: "button",
    predicate: "selection-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-cell-pictures",
  }),
  define({
    controlId: "XLSCTL.INSERT.PICTURES.OVER_CELLS.MENU",
    featureId: "spreadsheet.insert.pictureOverCellsMenu",
    canonicalId: "XLS-INS-IMAGE-01",
    group: "Insert",
    order: 8,
    controlKind: "menu",
    predicate: "menu-children",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    command: "disclosure",
    children: [
      "XLSCTL.INSERT.PICTURES.OVER_CELLS.DEVICE",
      "XLSCTL.INSERT.PICTURES.OVER_CELLS.AI_IMAGE",
      "XLSCTL.INSERT.PICTURES.OVER_CELLS.ONLINE",
    ],
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.INSERT.PICTURES.OVER_CELLS.DEVICE",
    featureId: "spreadsheet.insert.insertPictureFromDevice",
    canonicalId: "XLS-INS-IMAGE-01",
    group: "Insert",
    order: 9,
    controlKind: "button",
    predicate: "object-insertable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-drawing-objects",
  }),
  define({
    // Excel for the web starts a form from Insert ("Select Insert > Forms >
    // New Form"), which is why this is catalogued at all: MRAI has no Forms
    // service behind it, so it is unsupported, but the command exists on the
    // web edition's own ribbon and the entry says so rather than vanishing.
    controlId: "XLSCTL.INSERT.FORMS",
    featureId: "spreadsheet.insert.forms",
    canonicalId: "XLS-INS-TABLE-01",
    group: "Insert",
    order: 4,
    controlKind: "button",
    predicate: "no-verified-authority",
    reasonKey: "spreadsheet.capability.formsUnavailable",
    command: "unsupported",
    undoPolicy: "content-none",
    ioProfile: "unsupported",
  }),
  define({
    controlId: "XLSCTL.INSERT.PICTURES.OVER_CELLS.AI_IMAGE",
    featureId: "spreadsheet.insert.insertPictureFromAi",
    canonicalId: "XLS-INS-IMAGE-01",
    group: "Insert",
    order: 10,
    controlKind: "dialog-launcher",
    predicate: "object-insertable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-drawing-objects",
  }),
  define({
    controlId: "XLSCTL.INSERT.PICTURES.OVER_CELLS.ONLINE",
    featureId: "spreadsheet.insert.pictureOnlineImages",
    canonicalId: "XLS-INS-IMAGE-01",
    group: "Insert",
    order: 11,
    controlKind: "dialog-launcher",
    predicate: "no-verified-authority",
    reasonKey: "spreadsheet.capability.pictureLibraryUnavailable",
    command: "unsupported",
    undoPolicy: "content-none",
    ioProfile: "unsupported",
  }),
  define({
    controlId: "XLSCTL.INSERT.SHAPES.MENU",
    featureId: "spreadsheet.insert.insertShape",
    canonicalId: "XLS-INS-SHAPE-01",
    group: "Insert",
    order: 6,
    controlKind: "button",
    predicate: "object-insertable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-drawing-objects",
  }),
  define({
    controlId: "XLSCTL.INSERT.SHAPES.MENU.DISCLOSURE",
    featureId: "spreadsheet.insert.shapesMenu",
    canonicalId: "XLS-INS-SHAPE-01",
    group: "Insert",
    order: 7,
    controlKind: "menu",
    predicate: "menu-children",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    command: "disclosure",
    children: ["XLSCTL.INSERT.SHAPES.MENU"],
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  // Insert ▸ SmartArt writes the SHARED Office SmartArt catalog's layouts as
  // ordinary worksheet drawing shapes (one group per diagram), so it takes the
  // same authority shape as Insert ▸ Shapes: one bound writer plus its menu
  // disclosure. Excel for the web has no SmartArt entry, which is why this
  // control was a rendered boundary until the shared layer could write it.
  define({
    controlId: "XLSCTL.INSERT.SMARTART.MENU",
    featureId: "spreadsheet.insert.insertSmartArt",
    canonicalId: "XLS-INS-SMART-01",
    group: "Insert",
    order: 6,
    controlKind: "button",
    predicate: "object-insertable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-drawing-objects",
  }),
  define({
    controlId: "XLSCTL.INSERT.SMARTART.MENU.DISCLOSURE",
    featureId: "spreadsheet.insert.smartArtMenu",
    canonicalId: "XLS-INS-SMART-01",
    group: "Insert",
    order: 7,
    controlKind: "menu",
    predicate: "menu-children",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    command: "disclosure",
    children: ["XLSCTL.INSERT.SMARTART.MENU"],
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.INSERT.CONTROL.CHECKBOX",
    featureId: "spreadsheet.insert.insertCheckbox",
    canonicalId: "XLS-INS-CONTROL-01",
    group: "Insert",
    order: 8,
    controlKind: "button",
    predicate: "selection-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-cell-controls",
  }),
  define({
    controlId: "XLSCTL.INSERT.FILTER.SLICER",
    featureId: "spreadsheet.insert.openSlicerDialog",
    canonicalId: "XLS-INS-FILTER-01",
    group: "Insert",
    order: 9,
    controlKind: "dialog-launcher",
    predicate: "active-slicer-insert",
    reasonKey: "spreadsheet.capability.writerUnavailable",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.INSERT.CHART.RECOMMENDED_MENU",
    featureId: "spreadsheet.insert.openRecommendedChart",
    canonicalId: "XLS-INS-CHART-02",
    group: "Insert",
    order: 9,
    controlKind: "dialog-launcher",
    predicate: "selection",
    reasonKey: "spreadsheet.capability.selectionRequired",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.INSERT.CHART.RECOMMENDED_MENU.DISCLOSURE",
    featureId: "spreadsheet.insert.recommendedChartMenu",
    canonicalId: "XLS-INS-CHART-02",
    group: "Insert",
    order: 10,
    controlKind: "menu",
    predicate: "menu-children",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    command: "disclosure",
    children: ["XLSCTL.INSERT.CHART.RECOMMENDED_MENU"],
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.INSERT.CHART.RECOMMENDED_MENU.APPLY",
    featureId: "spreadsheet.insert.applyRecommendedChart",
    canonicalId: "XLS-INS-CHART-02",
    group: "Insert",
    order: 11,
    controlKind: "button",
    predicate: "object-insertable",
    reasonKey: "spreadsheet.capability.selectionRequired",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-charts",
  }),
  ...[
    // Excel's Insert ▸ Charts buttons, in Excel's order (Column or Bar, Line
    // or Area, Pie or Doughnut, Hierarchy, Statistic, Scatter or Bubble,
    // Waterfall / Funnel / Stock / Surface / Radar, Combo, Maps).
    ["XLSCTL.INSERT.CHART.COLUMN_MENU", "column", "XLS-INS-CHART-01", ["XLSCTL.INSERT.CHART.TYPE.CLUSTERED_COLUMN", "XLSCTL.INSERT.CHART.TYPE.CLUSTERED_BAR"]],
    ["XLSCTL.INSERT.CHART.LINE_MENU", "line", "XLS-INS-CHART-01", ["XLSCTL.INSERT.CHART.TYPE.LINE", "XLSCTL.INSERT.CHART.TYPE.AREA"]],
    ["XLSCTL.INSERT.CHART.PIE_MENU", "pie", "XLS-INS-CHART-01", ["XLSCTL.INSERT.CHART.TYPE.PIE", "XLSCTL.INSERT.CHART.TYPE.DOUGHNUT"]],
    ["XLSCTL.INSERT.CHART.HIERARCHY_MENU", "hierarchy", "XLS-INS-CHART-03", ["XLSCTL.INSERT.CHART.HIERARCHY_MENU"]],
    ["XLSCTL.INSERT.CHART.STATISTICAL_MENU", "statistical", "XLS-INS-CHART-03", ["XLSCTL.INSERT.CHART.STATISTICAL_MENU"]],
    ["XLSCTL.INSERT.CHART.SCATTER_MENU", "scatter", "XLS-INS-CHART-01", ["XLSCTL.INSERT.CHART.TYPE.XY_SCATTER", "XLSCTL.INSERT.CHART.TYPE.BUBBLE"]],
    ["XLSCTL.INSERT.CHART.OTHER_MENU", "radar", "XLS-INS-CHART-03", ["XLSCTL.INSERT.CHART.TYPE.RADAR"]],
    ["XLSCTL.INSERT.CHART.COMBO_MENU", "combo", "XLS-INS-CHART-01", ["XLSCTL.INSERT.CHART.TYPE.COMBO_CLUSTERED_COLUMN_LINE"]],
    ["XLSCTL.INSERT.CHART.MAP", "map", "XLS-INS-CHART-03", ["XLSCTL.INSERT.CHART.MAP"]],
  ].flatMap(([controlId, suffix, canonicalId, children], index) => [
    define({
      controlId: String(controlId),
      featureId: `spreadsheet.insert.chart.${String(suffix)}`,
      canonicalId: String(canonicalId),
      group: "Insert",
      order: 11 + index * 2,
      controlKind: "button",
      predicate: "object-insertable",
      reasonKey: "spreadsheet.capability.selectionRequired",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-charts",
    }),
    define({
      controlId: `${String(controlId)}.DISCLOSURE`,
      featureId: `spreadsheet.insert.chart.${String(suffix)}Menu`,
      canonicalId: String(canonicalId),
      group: "Insert",
      order: 12 + index * 2,
      controlKind: "menu",
      predicate: "menu-children",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "disclosure",
      children: children as readonly string[],
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
  ]),
  ...[
    ["XLSCTL.INSERT.CHART.TYPE.CLUSTERED_COLUMN", "clusteredColumn", "XLS-INS-CHART-01"],
    ["XLSCTL.INSERT.CHART.TYPE.LINE", "line", "XLS-INS-CHART-01"],
    ["XLSCTL.INSERT.CHART.TYPE.AREA", "area", "XLS-INS-CHART-01"],
    ["XLSCTL.INSERT.CHART.TYPE.XY_SCATTER", "scatter", "XLS-INS-CHART-01"],
    ["XLSCTL.INSERT.CHART.TYPE.BUBBLE", "bubble", "XLS-INS-CHART-01"],
    ["XLSCTL.INSERT.CHART.TYPE.PIE", "pie", "XLS-INS-CHART-01"],
    ["XLSCTL.INSERT.CHART.TYPE.DOUGHNUT", "doughnut", "XLS-INS-CHART-01"],
    ["XLSCTL.INSERT.CHART.TYPE.CLUSTERED_BAR", "clusteredBar", "XLS-INS-CHART-01"],
    ["XLSCTL.INSERT.CHART.TYPE.COMBO_CLUSTERED_COLUMN_LINE", "combo", "XLS-INS-CHART-01"],
    ["XLSCTL.INSERT.CHART.TYPE.RADAR", "radar", "XLS-INS-CHART-03"],
  ].map(([controlId, suffix, canonicalId], index) =>
    define({
      controlId: String(controlId),
      featureId: `spreadsheet.insert.chartType.${String(suffix)}`,
      canonicalId: String(canonicalId),
      group: "Insert",
      order: 30 + index,
      controlKind: "button",
      predicate: "object-insertable",
      reasonKey: "spreadsheet.capability.selectionRequired",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-charts",
    }),
  ),
  // Insert > PivotChart (Microsoft Support, "Create a PivotChart"): in a
  // PivotTable it adds that table's PivotChart; on worksheet data it creates
  // the PivotTable and its PivotChart together (PivotChart & PivotTable).
  define({
    controlId: "XLSCTL.INSERT.CHART.PIVOT_CHART",
    featureId: "spreadsheet.insert.pivotChart",
    canonicalId: "XLS-INS-PIVOTCHART-01",
    group: "Insert",
    order: 27,
    controlKind: "button",
    predicate: "pivot-insertable",
    reasonKey: "spreadsheet.capability.selectionRequired",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-charts",
  }),
  define({
    controlId: "XLSCTL.INSERT.CHART.PIVOT_CHART.DISCLOSURE",
    featureId: "spreadsheet.insert.pivotChartMenu",
    canonicalId: "XLS-INS-PIVOTCHART-01",
    group: "Insert",
    order: 28,
    controlKind: "menu",
    predicate: "menu-children",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    command: "disclosure",
    children: [
      "XLSCTL.INSERT.CHART.PIVOT_CHART",
      "XLSCTL.INSERT.CHART.PIVOT_CHART_AND_TABLE",
    ],
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.INSERT.CHART.PIVOT_CHART_AND_TABLE",
    featureId: "spreadsheet.insert.pivotChartAndTable",
    canonicalId: "XLS-INS-PIVOTCHART-01",
    group: "Insert",
    order: 29,
    controlKind: "button",
    predicate: "pivot-insertable",
    reasonKey: "spreadsheet.capability.selectionRequired",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-charts",
  }),
  define({
    controlId: "XLSCTL.INSERT.HYPERLINK.DIALOG",
    featureId: "spreadsheet.insert.openHyperlink",
    canonicalId: "XLS-INS-LINK-01",
    group: "Insert",
    order: 17,
    controlKind: "dialog-launcher",
    predicate: "hyperlink-insertable",
    reasonKey: "spreadsheet.capability.selectionRequired",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.INSERT.NEW_COMMENT",
    featureId: "spreadsheet.insert.openComment",
    canonicalId: "XLS-INS-COMMENT-01",
    group: "Insert",
    order: 18,
    controlKind: "dialog-launcher",
    predicate: "new-comment",
    // The predicate states its own reason (protection, writer, author,
    // note/comment conflict); this is the fallback it falls back to.
    reasonKey: "spreadsheet.capability.commentAuthorUnavailable",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.INSERT.TEXTBOX.MENU",
    featureId: "spreadsheet.insert.insertTextBox",
    canonicalId: "XLS-INS-TEXT-01",
    group: "Insert",
    order: 19,
    controlKind: "button",
    predicate: "object-insertable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-drawing-objects",
  }),
  // Excel's Insert > Text > WordArt: one fill-less text box carrying a SHARED
  // WordArt recipe (the catalog Docs, the deck and the Shape Format tab use).
  define({
    controlId: "XLSCTL.INSERT.WORDART.MENU",
    featureId: "spreadsheet.insert.insertWordArt",
    canonicalId: "XLS-INS-TEXT-01",
    group: "Insert",
    order: 21,
    controlKind: "button",
    predicate: "object-insertable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-drawing-objects",
  }),
  define({
    controlId: "XLSCTL.INSERT.SYMBOL",
    featureId: "spreadsheet.insert.openSymbolPicker",
    canonicalId: "XLS-INS-SYMBOL-01",
    group: "Insert",
    order: 20,
    controlKind: "dialog-launcher",
    predicate: "selection-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

/**
 * Excel-for-Web Data Validation vertical. The launcher and Clear All draft are
 * session-only; Apply is the only worksheet-validation writer and consumes an
 * exact workbook/sheet/selection receipt in the host.
 */
export const SPREADSHEET_DATA_VALIDATION_VERTICAL_CONTROL_DEFINITIONS =
  Object.freeze([
    define({
      controlId: "XLSCTL.DATA.TOOLS.VALIDATION",
      featureId: "spreadsheet.data.openDataValidation",
      canonicalId: "XLS-DAT-VALID-01",
      group: "Data",
      order: 22,
      controlKind: "dialog-launcher",
      predicate: "data-validation-launcher",
      reasonKey: "spreadsheet.capability.sheetProtected",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.DATA.VALIDATION.MRAI.CLEAR",
      featureId: "spreadsheet.data.clearValidationDraft",
      canonicalId: "XLS-DAT-VALID-01",
      group: "Data",
      order: 23,
      controlKind: "button",
      predicate: "data-validation-clear-draft",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.DATA.VALIDATION.MRAI.APPLY",
      featureId: "spreadsheet.data.applyDataValidation",
      canonicalId: "XLS-DAT-VALID-01",
      group: "Data",
      order: 24,
      controlKind: "button",
      predicate: "data-validation-apply",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-validation",
    }),
  ] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

/**
 * First production Table Design slice. The contextual tab is selection-owned;
 * this one aggregate gallery is the only persistent command exposed inside it.
 * Dynamic style members are parameters of the same authenticated writer rather
 * than invented fixed leaf commands.
 */
export const SPREADSHEET_TABLE_DESIGN_STYLE_VERTICAL_CONTROL_DEFINITIONS =
  Object.freeze([
    define({
      controlId: "XLSCTL.CONTEXT.TABLE.STYLE_GALLERY",
      featureId: "spreadsheet.context.table.applyStyle",
      canonicalId: "XLS-TBL-STYLE-01",
      group: "Table Design",
      order: 1,
      controlKind: "gallery",
      predicate: "active-table-style-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-table",
    }),
  ] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

/**
 * Native Table Style Options. Each checkbox is a parameterized writer over the
 * selected table's `tableStyleInfo`; the host binds the exact table and next
 * boolean value before dispatch so a retained control cannot toggle a newer
 * table generation back to the wrong state.
 */
export const SPREADSHEET_TABLE_STYLE_OPTIONS_VERTICAL_CONTROL_DEFINITIONS =
  Object.freeze([
    define({
      controlId: "XLSCTL.CONTEXT.TABLE.BANDED_ROWS",
      featureId: "spreadsheet.context.table.toggleBandedRows",
      canonicalId: "XLS-TBL-STYLE-01",
      group: "Table Design",
      order: 2,
      controlKind: "button",
      predicate: "active-table-style-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-table",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.TABLE.BANDED_COLUMNS",
      featureId: "spreadsheet.context.table.toggleBandedColumns",
      canonicalId: "XLS-TBL-STYLE-01",
      group: "Table Design",
      order: 3,
      controlKind: "button",
      predicate: "active-table-style-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-table",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.TABLE.FIRST_COLUMN",
      featureId: "spreadsheet.context.table.toggleFirstColumn",
      canonicalId: "XLS-TBL-STYLE-01",
      group: "Table Design",
      order: 4,
      controlKind: "button",
      predicate: "active-table-style-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-table",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.TABLE.LAST_COLUMN",
      featureId: "spreadsheet.context.table.toggleLastColumn",
      canonicalId: "XLS-TBL-STYLE-01",
      group: "Table Design",
      order: 5,
      controlKind: "button",
      predicate: "active-table-style-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-table",
    }),
  ] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

/** One authenticated workbook-wide Table Name writer. */
export const SPREADSHEET_TABLE_NAME_VERTICAL_CONTROL_DEFINITIONS = Object.freeze([
  define({
    controlId: "XLSCTL.CONTEXT.TABLE.NAME",
    featureId: "spreadsheet.context.table.rename",
    canonicalId: "XLS-TBL-PROP-01",
    group: "Table Design",
    order: 1,
    controlKind: "field",
    predicate: "active-table-name-editable",
    reasonKey: "spreadsheet.capability.writerUnavailable",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-table",
  }),
] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

/** One truthful Web Resize Table launcher and its authenticated commit. */
export const SPREADSHEET_TABLE_RESIZE_VERTICAL_CONTROL_DEFINITIONS =
  Object.freeze([
    define({
      controlId: "XLSCTL.CONTEXT.TABLE.RESIZE",
      featureId: "spreadsheet.context.table.openResize",
      canonicalId: "XLS-TBL-PROP-01",
      group: "Table Design",
      order: 2,
      controlKind: "dialog-launcher",
      predicate: "active-table-resize-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.TABLE.RESIZE.MRAI.APPLY",
      featureId: "spreadsheet.context.table.applyResize",
      canonicalId: "XLS-TBL-PROP-01",
      group: "Table Design",
      order: 3,
      controlKind: "button",
      predicate: "table-resize-apply",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-table",
    }),
  ] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

/**
 * Office Web Table Design commands that already have a native Spreadsheet
 * writer, plus truthful fail-closed leaves for capabilities whose provider is
 * not yet available.  These reuse the same table selection authority as the
 * existing style controls; no parallel ribbon component or editor model is
 * introduced.
 */
export const SPREADSHEET_TABLE_DESIGN_OFFICE_PARITY_CONTROL_DEFINITIONS =
  Object.freeze([
    define({
      controlId: "XLSCTL.CONTEXT.TABLE.ALT_TEXT",
      featureId: "spreadsheet.context.table.altText",
      canonicalId: "XLS-TBL-PROP-01",
      group: "Table Design",
      order: 4,
      controlKind: "dialog-launcher",
      predicate: "active-table-alt-text-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.TABLE.ALT_TEXT.MRAI.APPLY",
      featureId: "spreadsheet.context.table.applyAltText",
      canonicalId: "XLS-TBL-PROP-01",
      group: "Table Design",
      order: 4,
      controlKind: "button",
      predicate: "table-alt-text-apply",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-table",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.TABLE.INSERT_SLICER",
      featureId: "spreadsheet.context.table.insertSlicer",
      canonicalId: "XLS-TBL-ACTION-01",
      group: "Table Design",
      order: 5,
      controlKind: "dialog-launcher",
      predicate: "active-table-slicer-insert",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.TABLE.INSERT_SLICER.MRAI.APPLY",
      featureId: "spreadsheet.context.table.applySlicers",
      canonicalId: "XLS-TBL-ACTION-01",
      group: "Table Design",
      order: 5,
      controlKind: "button",
      predicate: "table-slicer-insert-apply",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-slicers",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.TABLE.SUMMARIZE_PIVOT",
      featureId: "spreadsheet.context.table.summarizePivot",
      canonicalId: "XLS-INS-PIVOT-01",
      group: "Table Design",
      order: 6,
      controlKind: "dialog-launcher",
      predicate: "active-table-pivot-source",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.TABLE.REMOVE_DUPLICATES",
      featureId: "spreadsheet.context.table.openRemoveDuplicates",
      canonicalId: "XLS-TBL-ACTION-01",
      group: "Table Design",
      order: 7,
      controlKind: "dialog-launcher",
      predicate: "active-table-style-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.TABLE.CONVERT_TO_RANGE",
      featureId: "spreadsheet.context.table.convertToRange",
      canonicalId: "XLS-TBL-ACTION-01",
      group: "Table Design",
      order: 8,
      controlKind: "button",
      predicate: "active-table-style-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-table",
    }),
    ...[
      [
        "XLSCTL.CONTEXT.TABLE.HEADER_ROW",
        "toggleHeaderRow",
        9,
        "active-table-style-editable",
      ],
      [
        "XLSCTL.CONTEXT.TABLE.TOTAL_ROW",
        "toggleTotalRow",
        10,
        "active-table-totals-editable",
      ],
      [
        "XLSCTL.CONTEXT.TABLE.FILTER_BUTTON",
        "toggleFilterButton",
        11,
        "active-table-style-editable",
      ],
    ].map(([controlId, suffix, order, predicate]) =>
      define({
        controlId: String(controlId),
        featureId: `spreadsheet.context.table.${String(suffix)}`,
        canonicalId: "XLS-TBL-STYLE-01",
        group: "Table Design",
        order: Number(order),
        controlKind: "toggle",
        predicate: predicate as SpreadsheetVisiblePredicate,
        reasonKey: "spreadsheet.capability.writerUnavailable",
        command: "bound",
        undoPolicy: "single-transaction",
        ioProfile: "worksheet-table",
      }),
    ),
  ] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

/**
 * Selection-owned Shape Format writer leaves. Dynamic gallery/menu values are
 * carried by one consume-once Spreadsheet receipt; imported shapes without a
 * safe native rewrite route remain visible but fail closed.
 */
export const SPREADSHEET_SHAPE_FORMAT_NATIVE_CORE_CONTROL_DEFINITIONS =
  Object.freeze([
    define({
      controlId: "XLSCTL.CONTEXT.SHAPE.STYLE_GALLERY.MRAI.APPLY",
      featureId: "spreadsheet.context.shape.applyQuickStyle",
      canonicalId: "XLS-INS-SHAPE-01",
      group: "Shape Format",
      order: 1,
      controlKind: "gallery",
      predicate: "active-shape-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-shape-format",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.SHAPE.CHANGE_SHAPE.MRAI.APPLY",
      featureId: "spreadsheet.context.shape.changeShape",
      canonicalId: "XLS-INS-SHAPE-01",
      group: "Shape Format",
      order: 2,
      controlKind: "menu",
      predicate: "active-shape-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-shape-format",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.SHAPE.OUTLINE_WEIGHT.MRAI.APPLY",
      featureId: "spreadsheet.context.shape.setOutlineWeight",
      canonicalId: "XLS-INS-SHAPE-01",
      group: "Shape Format",
      order: 3,
      controlKind: "menu",
      predicate: "active-shape-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-shape-format",
    }),
    ...[
      [
        "XLSCTL.CONTEXT.SHAPE.FILL_COLOR.MRAI.APPLY",
        "setFillColor",
        4,
      ],
      [
        "XLSCTL.CONTEXT.SHAPE.OUTLINE_COLOR.MRAI.APPLY",
        "setOutlineColor",
        5,
      ],
      [
        "XLSCTL.CONTEXT.SHAPE.SHADOW_PRESET.MRAI.APPLY",
        "setShadowPreset",
        6,
      ],
      [
        "XLSCTL.CONTEXT.SHAPE.SHADOW_COLOR.MRAI.APPLY",
        "setShadowColor",
        7,
      ],
    ].map(([controlId, suffix, order]) =>
      define({
        controlId: String(controlId),
        featureId: `spreadsheet.context.shape.${String(suffix)}`,
        canonicalId: "XLS-INS-SHAPE-01",
        group: "Shape Format",
        order: Number(order),
        controlKind: "menu",
        predicate: "active-shape-editable",
        reasonKey: "spreadsheet.capability.writerUnavailable",
        command: "bound",
        undoPolicy: "single-transaction",
        ioProfile: ["setShadowPreset", "setShadowColor"].includes(
          String(suffix),
        )
          ? "worksheet-drawing-objects"
          : "worksheet-shape-format",
      }),
    ),
    ...[
      ["XLSCTL.CONTEXT.SHAPE.TEXT.MRAI.APPLY", "setText", 8, "field"],
      [
        "XLSCTL.CONTEXT.SHAPE.TEXT_SIZE.MRAI.APPLY",
        "setTextSize",
        9,
        "field",
      ],
      [
        "XLSCTL.CONTEXT.SHAPE.TEXT_COLOR.MRAI.APPLY",
        "setTextColor",
        10,
        "menu",
      ],
      [
        "XLSCTL.CONTEXT.SHAPE.ALT_TEXT.MRAI.APPLY",
        "setAltText",
        11,
        "field",
      ],
    ].map(([controlId, suffix, order, controlKind]) =>
      define({
        controlId: String(controlId),
        featureId: `spreadsheet.context.shape.${String(suffix)}`,
        canonicalId: "XLS-INS-SHAPE-01",
        group: "Shape Format",
        order: Number(order),
        controlKind: controlKind as "field" | "menu",
        predicate: "active-shape-editable",
        reasonKey: "spreadsheet.capability.writerUnavailable",
        command: "bound",
        undoPolicy: "single-transaction",
        ioProfile: ["setText", "setTextSize"].includes(String(suffix))
          ? "worksheet-drawing-objects"
          : "worksheet-shape-format",
      }),
    ),
    ...[
      ["XLSCTL.CONTEXT.SHAPE.WIDTH.MRAI.APPLY", "setWidth", 12, "field"],
      [
        "XLSCTL.CONTEXT.SHAPE.HEIGHT.MRAI.APPLY",
        "setHeight",
        13,
        "field",
      ],
      ["XLSCTL.CONTEXT.SHAPE.DELETE.MRAI.APPLY", "delete", 14, "button"],
      // Lock aspect ratio: Excel's Size group checkbox (Format Shape >
      // Size & Properties > Lock aspect ratio), persisted on the object as
      // `a:spLocks noChangeAspect` is in the file.
      ["XLSCTL.CONTEXT.SHAPE.LOCK_ASPECT.MRAI.APPLY", "toggleLockAspect", 15, "toggle"],
      // WordArt Styles (text fill, outline, effects) on the label's runs.
      ["XLSCTL.CONTEXT.SHAPE.TEXT_PAINT.MRAI.APPLY", "setTextPaint", 16, "button"],
      // The yellow adjustment handles: the dragged values land in the
      // preset's `a:prstGeom/a:avLst` (ECMA-376 §20.1.9.3).
      ["XLSCTL.CONTEXT.SHAPE.ADJUST.MRAI.APPLY", "setAdjustments", 17, "button"],
      // Edit Shape > Edit Points: the outline becomes the shape's own
      // `a:custGeom` path, the one custom-geometry model every host stores.
      ["XLSCTL.CONTEXT.SHAPE.EDIT_POINTS.MRAI.APPLY", "setCustomPath", 18, "button"],
      // Shape Format > Shape Effects: the shared preset menu (none, shadow and
      // the five `office_shape_effects` presets) as one command.
      ["XLSCTL.CONTEXT.SHAPE.EFFECT_PRESET.MRAI.APPLY", "setEffectPreset", 19, "button"],
      // Format Shape > Size & Properties > Properties: Move and size with
      // cells / Move but don't size / Don't move or size (`editAs`).
      ["XLSCTL.CONTEXT.SHAPE.PLACEMENT.MRAI.APPLY", "setPlacement", 20, "field"],
      // Format Shape > Size & Properties > Properties > Print object
      // (`xdr:clientData@fPrintsWithSheet`).
      ["XLSCTL.CONTEXT.SHAPE.PRINT_OBJECT.MRAI.APPLY", "togglePrintObject", 21, "toggle"],
    ].map(([controlId, suffix, order, controlKind]) =>
      define({
        controlId: String(controlId),
        featureId: `spreadsheet.context.shape.${String(suffix)}`,
        canonicalId: "XLS-INS-SHAPE-01",
        group: "Shape Format",
        order: Number(order),
        controlKind: controlKind as "button" | "field" | "toggle",
        predicate: "active-shape-editable",
        reasonKey: "spreadsheet.capability.writerUnavailable",
        command: "bound",
        undoPolicy: "single-transaction",
        ioProfile: "worksheet-shape-format",
      }),
    ),
    ...[
      [
        "XLSCTL.CONTEXT.SHAPE.BRING_TO_FRONT.MRAI.APPLY",
        "bringToFront",
        15,
      ],
      [
        "XLSCTL.CONTEXT.SHAPE.BRING_FORWARD.MRAI.APPLY",
        "bringForward",
        16,
      ],
      [
        "XLSCTL.CONTEXT.SHAPE.SEND_BACKWARD.MRAI.APPLY",
        "sendBackward",
        17,
      ],
      [
        "XLSCTL.CONTEXT.SHAPE.SEND_TO_BACK.MRAI.APPLY",
        "sendToBack",
        18,
      ],
    ].map(([controlId, suffix, order]) =>
      define({
        controlId: String(controlId),
        featureId: `spreadsheet.context.shape.${String(suffix)}`,
        canonicalId: "XLS-INS-SHAPE-01",
        group: "Shape Format",
        order: Number(order),
        controlKind: "button",
        predicate: "active-shape-editable",
        reasonKey: "spreadsheet.capability.writerUnavailable",
        command: "bound",
        undoPolicy: "single-transaction",
        ioProfile: "worksheet-shape-format",
      }),
    ),
    define({
      controlId: "XLSCTL.CONTEXT.SHAPE.ROTATE_MENU",
      featureId: "spreadsheet.context.shape.rotateMenu",
      canonicalId: "XLS-INS-SHAPE-01",
      group: "Shape Format",
      order: 19,
      controlKind: "menu",
      predicate: "menu-children",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "disclosure",
      children: [
        "XLSCTL.CONTEXT.SHAPE.ROTATE_LEFT.MRAI.APPLY",
        "XLSCTL.CONTEXT.SHAPE.ROTATE_RIGHT.MRAI.APPLY",
        "XLSCTL.CONTEXT.SHAPE.FLIP_VERTICAL.MRAI.APPLY",
        "XLSCTL.CONTEXT.SHAPE.FLIP_HORIZONTAL.MRAI.APPLY",
        "XLSCTL.CONTEXT.SHAPE.ROTATION",
      ],
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    ...[
      [
        "XLSCTL.CONTEXT.SHAPE.ROTATE_LEFT.MRAI.APPLY",
        "rotateLeft",
        20,
      ],
      [
        "XLSCTL.CONTEXT.SHAPE.ROTATE_RIGHT.MRAI.APPLY",
        "rotateRight",
        21,
      ],
    ].map(([controlId, suffix, order]) =>
      define({
        controlId: String(controlId),
        featureId: `spreadsheet.context.shape.${String(suffix)}`,
        canonicalId: "XLS-INS-SHAPE-01",
        group: "Shape Format",
        order: Number(order),
        controlKind: "button",
        predicate: "active-shape-editable",
        reasonKey: "spreadsheet.capability.writerUnavailable",
        command: "bound",
        undoPolicy: "single-transaction",
        ioProfile: "worksheet-shape-format",
      }),
    ),
    ...[
      [
        "XLSCTL.CONTEXT.SHAPE.FLIP_VERTICAL.MRAI.APPLY",
        "flipVertical",
        22,
      ],
      [
        "XLSCTL.CONTEXT.SHAPE.FLIP_HORIZONTAL.MRAI.APPLY",
        "flipHorizontal",
        23,
      ],
    ].map(([controlId, suffix, order]) =>
      define({
        controlId: String(controlId),
        featureId: `spreadsheet.context.shape.${String(suffix)}`,
        canonicalId: "XLS-INS-SHAPE-01",
        group: "Shape Format",
        order: Number(order),
        controlKind: "toggle",
        predicate: "active-shape-editable",
        reasonKey: "spreadsheet.capability.writerUnavailable",
        command: "bound",
        undoPolicy: "single-transaction",
        ioProfile: "worksheet-shape-format",
      }),
    ),
    define({
      controlId: "XLSCTL.CONTEXT.SHAPE.ROTATION",
      featureId: "spreadsheet.context.shape.openRotation",
      canonicalId: "XLS-INS-SHAPE-01",
      group: "Shape Format",
      order: 24,
      controlKind: "dialog-launcher",
      predicate: "active-shape-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.SHAPE.ROTATION.MRAI.APPLY",
      featureId: "spreadsheet.context.shape.setRotation",
      canonicalId: "XLS-INS-SHAPE-01",
      group: "Shape Format",
      order: 25,
      controlKind: "button",
      predicate: "shape-rotation-apply",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-shape-format",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.SHAPE.ALIGN_MENU",
      featureId: "spreadsheet.context.shape.alignMenu",
      canonicalId: "XLS-INS-SHAPE-01",
      group: "Shape Format",
      order: 26,
      controlKind: "menu",
      predicate: "menu-children",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "disclosure",
      children: [
        "XLSCTL.CONTEXT.SHAPE.SNAP_TO_GRID",
        "XLSCTL.CONTEXT.SHAPE.SNAP_TO_SHAPE",
      ],
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    ...[
      ["XLSCTL.CONTEXT.SHAPE.SNAP_TO_GRID", "toggleSnapGrid", 27],
      ["XLSCTL.CONTEXT.SHAPE.SNAP_TO_SHAPE", "toggleSnapShape", 28],
    ].map(([controlId, suffix, order]) =>
      define({
        controlId: String(controlId),
        featureId: `spreadsheet.context.shape.${String(suffix)}`,
        canonicalId: "XLS-INS-SHAPE-01",
        group: "Shape Format",
        order: Number(order),
        controlKind: "toggle",
        predicate: "active-shape-editable",
        reasonKey: "spreadsheet.capability.writerUnavailable",
        command: "bound",
        undoPolicy: "content-none",
        ioProfile: "session-only",
      }),
    ),
    // Shape Outline > Arrows. Appended so the orders above keep their places.
    define({
      controlId: "XLSCTL.CONTEXT.SHAPE.LINE_ENDS.MRAI.APPLY",
      featureId: "spreadsheet.context.shape.setLineEnds",
      canonicalId: "XLS-INS-SHAPE-01",
      group: "Shape Format",
      order: 29,
      controlKind: "gallery",
      predicate: "active-shape-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-shape-format",
    }),
    // Shape Outline > Dashes (Office's eight) and the full Shape Effects menu
    // (Preset, Shadow, Reflection, Glow, Soft Edges, Bevel, 3-D Rotation, the
    // object's `shapeEffects`). Appended, as Arrows was.
    define({
      controlId: "XLSCTL.CONTEXT.SHAPE.OUTLINE_DASH.MRAI.APPLY",
      featureId: "spreadsheet.context.shape.setOutlineDash",
      canonicalId: "XLS-INS-SHAPE-01",
      group: "Shape Format",
      order: 30,
      controlKind: "menu",
      predicate: "active-shape-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-shape-format",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.SHAPE.OUTLINE_SKETCH.MRAI.APPLY",
      featureId: "spreadsheet.context.shape.setOutlineSketch",
      canonicalId: "XLS-INS-SHAPE-01",
      group: "Shape Format",
      order: 32,
      controlKind: "menu",
      predicate: "active-shape-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-shape-format",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.SHAPE.TEXT_TRANSFORM.MRAI.APPLY",
      featureId: "spreadsheet.context.shape.setTextTransform",
      canonicalId: "XLS-INS-SHAPE-01",
      group: "Shape Format",
      order: 33,
      controlKind: "gallery",
      predicate: "active-shape-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-shape-format",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.SHAPE.LINE_DETAIL.MRAI.APPLY",
      featureId: "spreadsheet.context.shape.setLineDetail",
      canonicalId: "XLS-INS-SHAPE-01",
      group: "Shape Format",
      order: 34,
      controlKind: "menu",
      predicate: "active-shape-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-shape-format",
    }),
    define({
      // Format Shape > Fill & Line (the shared page): pattern fill, the fill's
      // and line's Transparency, gradient fill and line, weight.
      controlId: "XLSCTL.CONTEXT.SHAPE.FILL_LINE.MRAI.APPLY",
      featureId: "spreadsheet.context.shape.setFillLine",
      canonicalId: "XLS-INS-SHAPE-01",
      group: "Shape Format",
      order: 35,
      controlKind: "menu",
      predicate: "active-shape-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-shape-format",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.SHAPE.EFFECTS.MRAI.APPLY",
      featureId: "spreadsheet.context.shape.setEffects",
      canonicalId: "XLS-INS-SHAPE-01",
      group: "Shape Format",
      order: 31,
      controlKind: "gallery",
      predicate: "active-shape-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-shape-format",
    }),
  ] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

export const SPREADSHEET_SHAPE_FORMAT_VISIBLE_CONTROL_IDS = Object.freeze([
  "XLSCTL.CONTEXT.SHAPE.STYLE_GALLERY.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.CHANGE_SHAPE.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.FILL_COLOR.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.OUTLINE_COLOR.MRAI.APPLY",
  // Shape Outline > WEIGHT. Excel offers it beside the outline colour ("Go to
  // Format > Shape Outline, point to Weight"), and the writer, the ladder and
  // this row all existed -- but the control was listed INTERNAL, so the nine
  // weights were reachable only by a caller holding a receipt and the ribbon
  // showed nothing at all (audit 64.7, closed in 68).
  "XLSCTL.CONTEXT.SHAPE.OUTLINE_WEIGHT.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.LINE_ENDS.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.OUTLINE_DASH.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.OUTLINE_SKETCH.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.TEXT_TRANSFORM.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.LINE_DETAIL.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.FILL_LINE.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.EFFECTS.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.ALT_TEXT.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.BRING_TO_FRONT.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.BRING_FORWARD.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.SEND_BACKWARD.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.SEND_TO_BACK.MRAI.APPLY",
  // The ONE Shape Format tab (components/office-ribbon/OfficeShapeFormatTab)
  // draws Office's anatomy on the sheet too: the Size fields, the Rotate menu
  // and its exact box, the Text Fill bucket and Shape Effects are on the
  // ribbon now, and a shape's text is typed in place (Excel's own way). A
  // writer whose control is on screen must be a VISIBLE row, or the
  // authority refuses the write it has just offered (that was the silent
  // "nothing happens" behind every one of these until now).
  "XLSCTL.CONTEXT.SHAPE.WIDTH.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.HEIGHT.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.ROTATE_LEFT.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.ROTATE_RIGHT.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.FLIP_HORIZONTAL.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.FLIP_VERTICAL.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.TEXT_COLOR.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.SHADOW_PRESET.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.TEXT.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.LOCK_ASPECT.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.TEXT_PAINT.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.ADJUST.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.EDIT_POINTS.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.EFFECT_PRESET.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.PLACEMENT.MRAI.APPLY",
  "XLSCTL.CONTEXT.SHAPE.PRINT_OBJECT.MRAI.APPLY",
] as const);

const SPREADSHEET_SHAPE_FORMAT_VISIBLE_CONTROL_ID_SET = new Set<string>(
  SPREADSHEET_SHAPE_FORMAT_VISIBLE_CONTROL_IDS,
);

export const SPREADSHEET_SHAPE_FORMAT_INTERNAL_CONTROL_IDS = Object.freeze(
  SPREADSHEET_SHAPE_FORMAT_NATIVE_CORE_CONTROL_DEFINITIONS.map(
    ({ controlId }) => controlId,
  ).filter(
    (controlId) =>
      !SPREADSHEET_SHAPE_FORMAT_VISIBLE_CONTROL_ID_SET.has(controlId),
  ),
);

const SPREADSHEET_SHAPE_FORMAT_INTERNAL_CONTROL_ID_SET = new Set<string>(
  SPREADSHEET_SHAPE_FORMAT_INTERNAL_CONTROL_IDS,
);

/** Picture Format transforms that have exact DrawingML anchor/order authority. */
export const SPREADSHEET_PICTURE_FORMAT_TRANSFORM_CONTROL_DEFINITIONS =
  Object.freeze([
    ...[
      ["XLSCTL.CONTEXT.PICTURE.ROTATE_LEFT.MRAI.APPLY", "rotateLeft", 1],
      ["XLSCTL.CONTEXT.PICTURE.ROTATE_RIGHT.MRAI.APPLY", "rotateRight", 2],
      ["XLSCTL.CONTEXT.PICTURE.FLIP_VERTICAL.MRAI.APPLY", "flipVertical", 3],
      ["XLSCTL.CONTEXT.PICTURE.FLIP_HORIZONTAL.MRAI.APPLY", "flipHorizontal", 4],
      ["XLSCTL.CONTEXT.PICTURE.WIDTH.MRAI.APPLY", "setWidth", 5],
      ["XLSCTL.CONTEXT.PICTURE.HEIGHT.MRAI.APPLY", "setHeight", 6],
      ["XLSCTL.CONTEXT.PICTURE.DELETE.MRAI.APPLY", "delete", 7],
      [
        "XLSCTL.CONTEXT.PICTURE.BRING_TO_FRONT.MRAI.APPLY",
        "bringToFront",
        8,
      ],
      [
        "XLSCTL.CONTEXT.PICTURE.BRING_FORWARD.MRAI.APPLY",
        "bringForward",
        9,
      ],
      [
        "XLSCTL.CONTEXT.PICTURE.SEND_BACKWARD.MRAI.APPLY",
        "sendBackward",
        10,
      ],
      [
        "XLSCTL.CONTEXT.PICTURE.SEND_TO_BACK.MRAI.APPLY",
        "sendToBack",
        11,
      ],
    ].map(([controlId, suffix, order]) =>
      define({
        controlId: String(controlId),
        featureId: `spreadsheet.context.picture.${String(suffix)}`,
        canonicalId: "XLS-INS-IMAGE-01",
        group: "Picture Format",
        order: Number(order),
        controlKind: String(suffix).startsWith("flip") ? "toggle" : "button",
        predicate: "active-picture-transform-editable",
        reasonKey: "spreadsheet.capability.writerUnavailable",
        command: "bound",
        undoPolicy: "single-transaction",
        ioProfile: "worksheet-picture-transforms",
      }),
    ),
    define({
      controlId: "XLSCTL.CONTEXT.PICTURE.CROP",
      featureId: "spreadsheet.context.picture.toggleCrop",
      canonicalId: "XLS-INS-IMAGE-01",
      group: "Picture Format",
      order: 12,
      controlKind: "toggle",
      predicate: "active-picture-transform-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PICTURE.CROP.MRAI.APPLY",
      featureId: "spreadsheet.context.picture.setCrop",
      canonicalId: "XLS-INS-IMAGE-01",
      group: "Picture Format",
      order: 13,
      controlKind: "field",
      predicate: "active-picture-transform-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-picture-transforms",
    }),
    // Picture Format > Picture Styles: Office's 28 built-in styles, written
    // into the picture's own `xdr:pic/xdr:spPr` through the drawing identity
    // a crop uses.
    define({
      controlId: "XLSCTL.CONTEXT.PICTURE.STYLE_GALLERY.MRAI.APPLY",
      featureId: "spreadsheet.context.picture.setPictureStyle",
      canonicalId: "XLS-INS-IMAGE-01",
      group: "Picture Format",
      order: 14,
      controlKind: "gallery",
      predicate: "active-picture-transform-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-picture-transforms",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PICTURE.ROTATE_MENU",
      featureId: "spreadsheet.context.picture.rotateMenu",
      canonicalId: "XLS-INS-IMAGE-01",
      group: "Picture Format",
      order: 20,
      controlKind: "menu",
      predicate: "menu-children",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "disclosure",
      children: [
        "XLSCTL.CONTEXT.PICTURE.ROTATE_RIGHT.MRAI.APPLY",
        "XLSCTL.CONTEXT.PICTURE.ROTATE_LEFT.MRAI.APPLY",
        "XLSCTL.CONTEXT.PICTURE.FLIP_VERTICAL.MRAI.APPLY",
        "XLSCTL.CONTEXT.PICTURE.FLIP_HORIZONTAL.MRAI.APPLY",
        "XLSCTL.CONTEXT.PICTURE.ROTATION",
      ],
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PICTURE.ROTATION",
      featureId: "spreadsheet.context.picture.openRotation",
      canonicalId: "XLS-INS-IMAGE-01",
      group: "Picture Format",
      order: 21,
      controlKind: "dialog-launcher",
      predicate: "active-picture-transform-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PICTURE.ROTATION.MRAI.APPLY",
      featureId: "spreadsheet.context.picture.setRotation",
      canonicalId: "XLS-INS-IMAGE-01",
      group: "Picture Format",
      order: 22,
      controlKind: "button",
      predicate: "picture-rotation-apply",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-picture-transforms",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PICTURE.ALIGN_MENU",
      featureId: "spreadsheet.context.picture.alignMenu",
      canonicalId: "XLS-INS-IMAGE-01",
      group: "Picture Format",
      order: 23,
      controlKind: "menu",
      predicate: "menu-children",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "disclosure",
      children: [
        "XLSCTL.CONTEXT.PICTURE.SNAP_TO_GRID",
        "XLSCTL.CONTEXT.PICTURE.SNAP_TO_SHAPE",
      ],
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    ...[
      ["XLSCTL.CONTEXT.PICTURE.SNAP_TO_GRID", "toggleSnapGrid", 24],
      ["XLSCTL.CONTEXT.PICTURE.SNAP_TO_SHAPE", "toggleSnapShape", 25],
    ].map(([controlId, suffix, order]) =>
      define({
        controlId: String(controlId),
        featureId: `spreadsheet.context.picture.${String(suffix)}`,
        canonicalId: "XLS-INS-IMAGE-01",
        group: "Picture Format",
        order: Number(order),
        controlKind: "toggle",
        predicate: "active-picture-transform-editable",
        reasonKey: "spreadsheet.capability.writerUnavailable",
        command: "bound",
        undoPolicy: "content-none",
        ioProfile: "session-only",
      }),
    ),
    // Format Picture > Size > Lock aspect ratio: a Height or Width typed then
    // scales the other side, as Excel scales a locked picture; persisted on
    // the object as the shape's lock is.
    define({
      controlId: "XLSCTL.CONTEXT.PICTURE.LOCK_ASPECT.MRAI.APPLY",
      featureId: "spreadsheet.context.picture.toggleLockAspect",
      canonicalId: "XLS-INS-IMAGE-01",
      group: "Picture Format",
      order: 26,
      controlKind: "toggle",
      predicate: "active-picture-transform-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-picture-transforms",
    }),
    // Format Picture > Size & Properties > Properties: how the picture
    // follows the cells under it (`editAs`).
    define({
      controlId: "XLSCTL.CONTEXT.PICTURE.PLACEMENT.MRAI.APPLY",
      featureId: "spreadsheet.context.picture.setPlacement",
      canonicalId: "XLS-INS-IMAGE-01",
      group: "Picture Format",
      order: 27,
      controlKind: "field",
      predicate: "active-picture-transform-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-picture-transforms",
    }),
    // Format Picture > Properties > Print object (`fPrintsWithSheet`).
    define({
      controlId: "XLSCTL.CONTEXT.PICTURE.PRINT_OBJECT.MRAI.APPLY",
      featureId: "spreadsheet.context.picture.togglePrintObject",
      canonicalId: "XLS-INS-IMAGE-01",
      group: "Picture Format",
      order: 28,
      controlKind: "toggle",
      predicate: "active-picture-transform-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-picture-transforms",
    }),
  ] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

export const SPREADSHEET_PICTURE_FORMAT_EFFECT_CONTROL_DEFINITIONS =
  Object.freeze(
    [
      [
        "XLSCTL.CONTEXT.PICTURE.SHADOW_PRESET.MRAI.APPLY",
        "setShadowPreset",
        10,
      ],
      [
        "XLSCTL.CONTEXT.PICTURE.SHADOW_COLOR.MRAI.APPLY",
        "setShadowColor",
        11,
      ],
      // Picture Effects beyond the Shadow (Reflection, Glow, Soft Edges,
      // Bevel, 3-D Format, 3-D Rotation): one pick, `{kind, value}` as JSON
      // (lib/spreadsheet/picture-effects), from the menu or the pane.
      [
        "XLSCTL.CONTEXT.PICTURE.EFFECTS.MRAI.APPLY",
        "setPictureEffect",
        14,
      ],
    ].map(([controlId, suffix, order]) =>
      define({
        controlId: String(controlId),
        featureId: `spreadsheet.context.picture.${String(suffix)}`,
        canonicalId: "XLS-INS-IMAGE-01",
        group: "Picture Format",
        order: Number(order),
        controlKind: "menu",
        predicate: "active-picture-effects-editable",
        reasonKey: "spreadsheet.capability.writerUnavailable",
        command: "bound",
        undoPolicy: "single-transaction",
        ioProfile: "worksheet-drawing-objects",
      }),
    ),
  );

// Picture Format > Adjust > Artistic Effects (the shared gallery). The
// from-scratch writer renders the effect into the picture; the in-place
// writer adds the rendered image and names the effect beside the original
// (`a14:imgProps`), so it takes the effects predicate on either route.
export const SPREADSHEET_PICTURE_FORMAT_ARTISTIC_EFFECT_CONTROL_DEFINITIONS =
  Object.freeze([
    define({
      controlId: "XLSCTL.CONTEXT.PICTURE.ARTISTIC_EFFECTS.MRAI.APPLY",
      featureId: "spreadsheet.context.picture.setArtisticEffect",
      canonicalId: "XLS-INS-IMAGE-01",
      group: "Picture Format",
      order: 15,
      controlKind: "gallery",
      predicate: "active-picture-effects-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-drawing-objects",
    }),
    // Picture Format > Adjust > Corrections, Color and Transparency: the
    // shared galleries, written on the picture's blip (`office_blip_adjust`).
    ...(
      [
        ["XLSCTL.CONTEXT.PICTURE.CORRECTIONS.MRAI.APPLY", "setCorrections", 12],
        ["XLSCTL.CONTEXT.PICTURE.COLOR.MRAI.APPLY", "setColor", 13],
        ["XLSCTL.CONTEXT.PICTURE.TRANSPARENCY.MRAI.APPLY", "setTransparency", 16],
      ] as const
    ).map(([controlId, suffix, order]) =>
      define({
        controlId,
        featureId: `spreadsheet.context.picture.${suffix}`,
        canonicalId: "XLS-INS-IMAGE-01",
        group: "Picture Format",
        order,
        controlKind: "gallery",
        predicate: "active-picture-effects-editable",
        reasonKey: "spreadsheet.capability.writerUnavailable",
        command: "bound",
        undoPolicy: "single-transaction",
        ioProfile: "worksheet-drawing-objects",
      }),
    ),
  ] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

export const SPREADSHEET_PICTURE_FORMAT_METADATA_CONTROL_DEFINITIONS =
  Object.freeze([
    define({
      controlId: "XLSCTL.CONTEXT.PICTURE.CHANGE",
      featureId: "spreadsheet.context.picture.openChange",
      canonicalId: "XLS-INS-IMAGE-01",
      group: "Picture Format",
      order: 12,
      controlKind: "dialog-launcher",
      predicate: "active-picture-metadata-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PICTURE.CHANGE.MRAI.APPLY",
      featureId: "spreadsheet.context.picture.applyChange",
      canonicalId: "XLS-INS-IMAGE-01",
      group: "Picture Format",
      order: 12,
      controlKind: "field",
      predicate: "active-picture-metadata-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-picture-metadata",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PICTURE.ALT_TEXT",
      featureId: "spreadsheet.context.picture.openAltText",
      canonicalId: "XLS-INS-IMAGE-01",
      group: "Picture Format",
      order: 13,
      controlKind: "dialog-launcher",
      predicate: "active-picture-metadata-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PICTURE.ALT_TEXT.MRAI.APPLY",
      featureId: "spreadsheet.context.picture.applyAltText",
      canonicalId: "XLS-INS-IMAGE-01",
      group: "Picture Format",
      order: 14,
      controlKind: "field",
      predicate: "active-picture-metadata-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-picture-metadata",
    }),
  ] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

export const SPREADSHEET_PICTURE_FORMAT_VISIBLE_CONTROL_IDS = Object.freeze([
  "XLSCTL.CONTEXT.PICTURE.ALT_TEXT",
  "XLSCTL.CONTEXT.PICTURE.ALT_TEXT.MRAI.APPLY",
  "XLSCTL.CONTEXT.PICTURE.BRING_TO_FRONT.MRAI.APPLY",
  "XLSCTL.CONTEXT.PICTURE.BRING_FORWARD.MRAI.APPLY",
  "XLSCTL.CONTEXT.PICTURE.SEND_BACKWARD.MRAI.APPLY",
  "XLSCTL.CONTEXT.PICTURE.SEND_TO_BACK.MRAI.APPLY",
  "XLSCTL.CONTEXT.PICTURE.CROP",
  "XLSCTL.CONTEXT.PICTURE.CROP.MRAI.APPLY",
  "XLSCTL.CONTEXT.PICTURE.STYLE_GALLERY.MRAI.APPLY",
  // Excel's Picture Format > Size > Height / Width (the shared size fields).
  "XLSCTL.CONTEXT.PICTURE.WIDTH.MRAI.APPLY",
  "XLSCTL.CONTEXT.PICTURE.HEIGHT.MRAI.APPLY",
  // Format Picture > Size > Lock aspect ratio.
  "XLSCTL.CONTEXT.PICTURE.LOCK_ASPECT.MRAI.APPLY",
  // Format Picture > Properties > Object positioning.
  "XLSCTL.CONTEXT.PICTURE.PLACEMENT.MRAI.APPLY",
  // Format Picture > Properties > Print object.
  "XLSCTL.CONTEXT.PICTURE.PRINT_OBJECT.MRAI.APPLY",
  "XLSCTL.CONTEXT.PICTURE.ARTISTIC_EFFECTS.MRAI.APPLY",
  "XLSCTL.CONTEXT.PICTURE.CORRECTIONS.MRAI.APPLY",
  "XLSCTL.CONTEXT.PICTURE.COLOR.MRAI.APPLY",
  "XLSCTL.CONTEXT.PICTURE.TRANSPARENCY.MRAI.APPLY",
] as const);

const SPREADSHEET_PICTURE_FORMAT_VISIBLE_CONTROL_ID_SET = new Set<string>(
  SPREADSHEET_PICTURE_FORMAT_VISIBLE_CONTROL_IDS,
);

export const SPREADSHEET_PICTURE_FORMAT_INTERNAL_CONTROL_IDS = Object.freeze(
  [
    ...SPREADSHEET_PICTURE_FORMAT_TRANSFORM_CONTROL_DEFINITIONS,
    ...SPREADSHEET_PICTURE_FORMAT_EFFECT_CONTROL_DEFINITIONS,
    ...SPREADSHEET_PICTURE_FORMAT_ARTISTIC_EFFECT_CONTROL_DEFINITIONS,
    ...SPREADSHEET_PICTURE_FORMAT_METADATA_CONTROL_DEFINITIONS,
  ]
    .map(({ controlId }) => controlId)
    .filter(
      (controlId) =>
        !SPREADSHEET_PICTURE_FORMAT_VISIBLE_CONTROL_ID_SET.has(controlId),
    ),
);

const SPREADSHEET_PICTURE_FORMAT_INTERNAL_CONTROL_ID_SET = new Set<string>(
  SPREADSHEET_PICTURE_FORMAT_INTERNAL_CONTROL_IDS,
);

// The Shape and Picture Format writers Excel for the web has no ribbon row for
// (Delete, the exact rotation, Shadow Color, Text Size, Change Picture) are
// never drawn, so the visible executor refuses them as hidden. Their own
// surfaces -- the object's context menu, the Format pane, the keyboard -- file
// a receipt and run the bound feature directly; the receipt writer still
// revalidates the object, its native write route and sheet protection.
const SPREADSHEET_OBJECT_FORMAT_INTERNAL_FEATURE_BY_CONTROL = new Map<
  string,
  string
>(
  [
    ...SPREADSHEET_SHAPE_FORMAT_NATIVE_CORE_CONTROL_DEFINITIONS,
    ...SPREADSHEET_PICTURE_FORMAT_TRANSFORM_CONTROL_DEFINITIONS,
    ...SPREADSHEET_PICTURE_FORMAT_EFFECT_CONTROL_DEFINITIONS,
    ...SPREADSHEET_PICTURE_FORMAT_ARTISTIC_EFFECT_CONTROL_DEFINITIONS,
    ...SPREADSHEET_PICTURE_FORMAT_METADATA_CONTROL_DEFINITIONS,
  ]
    .filter(
      ({ controlId, command }) =>
        command === "bound" &&
        (SPREADSHEET_SHAPE_FORMAT_INTERNAL_CONTROL_ID_SET.has(controlId) ||
          SPREADSHEET_PICTURE_FORMAT_INTERNAL_CONTROL_ID_SET.has(controlId)),
    )
    .map(({ controlId, featureId }) => [controlId, featureId] as const),
);

/** The feature an internal-only Shape / Picture Format writer runs, else null. */
export function spreadsheetObjectFormatInternalFeatureId(
  controlId: string,
): string | null {
  return SPREADSHEET_OBJECT_FORMAT_INTERNAL_FEATURE_BY_CONTROL.get(controlId) ?? null;
}

export const SPREADSHEET_SLICER_DESIGN_FORMAT_CONTROL_DEFINITIONS =
  Object.freeze([
    ...[
      ["XLSCTL.CONTEXT.SLICER.TITLE.MRAI.APPLY", "setTitle", 1, "field"],
      [
        "XLSCTL.CONTEXT.SLICER.SHOW_HEADER.MRAI.APPLY",
        "toggleHeader",
        2,
        "toggle",
      ],
      ["XLSCTL.CONTEXT.SLICER.STYLE.MRAI.APPLY", "setStyle", 3, "gallery"],
      ["XLSCTL.CONTEXT.SLICER.COLUMNS.MRAI.APPLY", "setColumns", 4, "field"],
      [
        "XLSCTL.CONTEXT.SLICER.BUTTON_HEIGHT.MRAI.APPLY",
        "setButtonHeight",
        5,
        "field",
      ],
      ["XLSCTL.CONTEXT.SLICER.WIDTH.MRAI.APPLY", "setWidth", 6, "field"],
      ["XLSCTL.CONTEXT.SLICER.HEIGHT.MRAI.APPLY", "setHeight", 7, "field"],
    ].map(([controlId, suffix, order, controlKind]) =>
      define({
        controlId: String(controlId),
        featureId: `spreadsheet.context.slicer.${String(suffix)}`,
        canonicalId: "XLS-INS-FILTER-01",
        group: "Slicer Design",
        order: Number(order),
        controlKind: controlKind as "field" | "gallery" | "toggle",
        predicate: "active-slicer-editable",
        reasonKey: "spreadsheet.capability.writerUnavailable",
        command: "bound",
        undoPolicy: "single-transaction",
        ioProfile: "worksheet-slicers",
      }),
    ),
    define({
      controlId: "XLSCTL.CONTEXT.SLICER.CLEAR_FILTER.MRAI.APPLY",
      featureId: "spreadsheet.context.slicer.clearFilter",
      canonicalId: "XLS-INS-FILTER-01",
      group: "Slicer Design",
      order: 8,
      controlKind: "button",
      predicate: "active-slicer-filter-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-table",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.SLICER.REMOVE.MRAI.APPLY",
      featureId: "spreadsheet.context.slicer.remove",
      canonicalId: "XLS-INS-FILTER-01",
      group: "Slicer Design",
      order: 9,
      controlKind: "button",
      predicate: "active-slicer-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-slicers",
    }),
  ] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

/** High-frequency Sparkline Design leaves backed by x14 sparkline groups. */
export const SPREADSHEET_SPARKLINE_DESIGN_NATIVE_CORE_CONTROL_DEFINITIONS =
  Object.freeze([
    ...[
      ["XLSCTL.CONTEXT.SPARKLINE.TYPE.LINE", "setTypeLine", 1],
      ["XLSCTL.CONTEXT.SPARKLINE.TYPE.COLUMN", "setTypeColumn", 2],
      ["XLSCTL.CONTEXT.SPARKLINE.TYPE.WIN_LOSS", "setTypeWinLoss", 3],
    ].map(([controlId, suffix, order]) =>
      define({
        controlId: String(controlId),
        featureId: `spreadsheet.context.sparkline.${String(suffix)}`,
        canonicalId: "XLS-INS-SPARK-01",
        group: "Sparkline Design",
        order: Number(order),
        controlKind: "button",
        predicate: "active-sparkline-editable",
        reasonKey: "spreadsheet.capability.writerUnavailable",
        command: "bound",
        undoPolicy: "single-transaction",
        ioProfile: "worksheet-sparklines",
      }),
    ),
    define({
      controlId: "XLSCTL.CONTEXT.SPARKLINE.COLOR.MRAI.APPLY",
      featureId: "spreadsheet.context.sparkline.applyColor",
      canonicalId: "XLS-INS-SPARK-01",
      group: "Sparkline Design",
      order: 4,
      controlKind: "menu",
      predicate: "active-sparkline-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-sparklines",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.SPARKLINE.CLEAR",
      featureId: "spreadsheet.context.sparkline.clear",
      canonicalId: "XLS-INS-SPARK-01",
      group: "Sparkline Design",
      order: 5,
      controlKind: "button",
      predicate: "active-sparkline-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-sparklines",
    }),
    ...[
      ["XLSCTL.CONTEXT.SPARKLINE.MARKERS", "toggleMarkers", 6, "toggle"],
      ["XLSCTL.CONTEXT.SPARKLINE.HIGH_POINT", "toggleHigh", 7, "toggle"],
      ["XLSCTL.CONTEXT.SPARKLINE.LOW_POINT", "toggleLow", 8, "toggle"],
      ["XLSCTL.CONTEXT.SPARKLINE.FIRST_POINT", "toggleFirst", 9, "toggle"],
      ["XLSCTL.CONTEXT.SPARKLINE.LAST_POINT", "toggleLast", 10, "toggle"],
      ["XLSCTL.CONTEXT.SPARKLINE.NEGATIVE_POINTS", "toggleNegative", 11, "toggle"],
      ["XLSCTL.CONTEXT.SPARKLINE.SHOW_AXIS", "toggleAxis", 12, "toggle"],
      ["XLSCTL.CONTEXT.SPARKLINE.MIN_AXIS_TYPE", "setMinAxisType", 13, "field"],
      ["XLSCTL.CONTEXT.SPARKLINE.MAX_AXIS_TYPE", "setMaxAxisType", 14, "field"],
      ["XLSCTL.CONTEXT.SPARKLINE.MIN_AXIS_VALUE", "setMinAxisValue", 15, "field"],
      ["XLSCTL.CONTEXT.SPARKLINE.MAX_AXIS_VALUE", "setMaxAxisValue", 16, "field"],
      ["XLSCTL.CONTEXT.SPARKLINE.DATA_RANGE", "setDataRange", 17, "field"],
      ["XLSCTL.CONTEXT.SPARKLINE.EMPTY_CELLS", "setEmptyCells", 18, "field"],
      ["XLSCTL.CONTEXT.SPARKLINE.DISPLAY_HIDDEN", "toggleDisplayHidden", 19, "toggle"],
      ["XLSCTL.CONTEXT.SPARKLINE.GROUP", "group", 20, "button"],
      ["XLSCTL.CONTEXT.SPARKLINE.UNGROUP", "ungroup", 21, "button"],
      // The Style gallery: Excel's built-in sparkline styles.
      ["XLSCTL.CONTEXT.SPARKLINE.STYLE.MRAI.APPLY", "applyStyle", 22, "menu"],
    ].map(([controlId, suffix, order, controlKind]) =>
      define({
        controlId: String(controlId),
        featureId: `spreadsheet.context.sparkline.${String(suffix)}`,
        canonicalId: "XLS-INS-SPARK-01",
        group: "Sparkline Design",
        order: Number(order),
        controlKind: controlKind as SpreadsheetCapabilityControlKind,
        predicate: "active-sparkline-editable",
        reasonKey: "spreadsheet.capability.writerUnavailable",
        command: "bound",
        undoPolicy: "single-transaction",
        ioProfile: "worksheet-sparklines",
      }),
    ),
  ] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

/**
 * Selection-owned Chart Design surface. The same command IDs authorize both
 * locally-created charts and authenticated classic charts imported from XLSX;
 * foreign/chartEx parts remain preserve-only. Dynamic values travel in strict
 * consume-once receipts rather than being invented as fixed catalog leaves.
 */
export const SPREADSHEET_CHART_DESIGN_NATIVE_CORE_CONTROL_DEFINITIONS =
  Object.freeze([
    define({
      controlId: "XLSCTL.CONTEXT.CHART.SWITCH_ROW_COLUMN",
      featureId: "spreadsheet.context.chart.switchRowColumn",
      canonicalId: "XLS-INS-CHART-01",
      group: "Chart Design",
      order: 1,
      controlKind: "button",
      predicate: "active-native-chart-series-orientation-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-charts",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.CHART.ADD_ELEMENT_MENU",
      featureId: "spreadsheet.context.chart.addElementMenu",
      canonicalId: "XLS-INS-CHART-01",
      group: "Chart Design",
      order: 2,
      controlKind: "menu",
      predicate: "menu-children",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "disclosure",
      undoPolicy: "content-none",
      ioProfile: "session-only",
      children: [
        "XLSCTL.CONTEXT.CHART.TITLE.MRAI.APPLY",
        "XLSCTL.CONTEXT.CHART.AXIS_TITLE_X.MRAI.APPLY",
        "XLSCTL.CONTEXT.CHART.AXIS_TITLE_Y.MRAI.APPLY",
        "XLSCTL.CONTEXT.CHART.AXIS_TITLE_DEPTH.MRAI.APPLY",
        "XLSCTL.CONTEXT.CHART.LEGEND",
        "XLSCTL.CONTEXT.CHART.LEGEND_POSITION.MRAI.APPLY",
        "XLSCTL.CONTEXT.CHART.DATA_LABELS",
        "XLSCTL.CONTEXT.CHART.DATA_LABEL_POSITION.MRAI.APPLY",
        "XLSCTL.CONTEXT.CHART.DATA_TABLE",
        "XLSCTL.CONTEXT.CHART.AXES",
        "XLSCTL.CONTEXT.CHART.GRIDLINES",
        "XLSCTL.CONTEXT.CHART.TRENDLINE.MRAI.APPLY",
        "XLSCTL.CONTEXT.CHART.ERROR_BARS.MRAI.APPLY",
        "XLSCTL.CONTEXT.CHART.UP_DOWN_BARS.MRAI.APPLY",
        "XLSCTL.CONTEXT.CHART.LINES_MENU",
        "XLSCTL.CONTEXT.CHART.SECONDARY_AXIS.MRAI.APPLY",
      ],
    }),
    define({
      controlId: "XLSCTL.CONTEXT.CHART.QUICK_LAYOUT.MRAI.APPLY",
      featureId: "spreadsheet.context.chart.applyQuickLayout",
      canonicalId: "XLS-INS-CHART-01",
      group: "Chart Design",
      order: 3,
      controlKind: "gallery",
      predicate: "active-native-chart-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-chart-metadata",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.CHART.CHANGE_COLORS.MRAI.APPLY",
      featureId: "spreadsheet.context.chart.applyColorPalette",
      canonicalId: "XLS-INS-CHART-01",
      group: "Chart Design",
      order: 4,
      controlKind: "gallery",
      predicate: "active-native-chart-color-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-chart-metadata",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.CHART.STYLE_GALLERY.MRAI.APPLY",
      featureId: "spreadsheet.context.chart.applyStyle",
      canonicalId: "XLS-INS-CHART-01",
      group: "Chart Design",
      order: 5,
      controlKind: "gallery",
      predicate: "active-native-chart-style-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-chart-metadata",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.CHART.CHANGE_TYPE.MRAI.APPLY",
      featureId: "spreadsheet.context.chart.changeType",
      canonicalId: "XLS-INS-CHART-01",
      group: "Chart Design",
      order: 6,
      controlKind: "combobox",
      predicate: "active-native-chart-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-chart-metadata",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.CHART.TITLE.MRAI.APPLY",
      featureId: "spreadsheet.context.chart.setTitle",
      canonicalId: "XLS-INS-CHART-01",
      group: "Chart Design",
      order: 6,
      controlKind: "field",
      predicate: "active-native-chart-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-chart-metadata",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.CHART.DATA_LABELS",
      featureId: "spreadsheet.context.chart.toggleDataLabels",
      canonicalId: "XLS-INS-CHART-01",
      group: "Chart Design",
      order: 7,
      controlKind: "toggle",
      predicate: "active-native-chart-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-chart-metadata",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.CHART.AXIS_TITLE_X.MRAI.APPLY",
      featureId: "spreadsheet.context.chart.setAxisTitleX",
      canonicalId: "XLS-INS-CHART-01",
      group: "Chart Design",
      order: 8,
      controlKind: "field",
      predicate: "active-native-chart-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-chart-metadata",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.CHART.AXIS_TITLE_Y.MRAI.APPLY",
      featureId: "spreadsheet.context.chart.setAxisTitleY",
      canonicalId: "XLS-INS-CHART-01",
      group: "Chart Design",
      order: 9,
      controlKind: "field",
      predicate: "active-native-chart-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-chart-metadata",
    }),
    // Axis Titles > Depth: a 3-D chart's depth (series) axis title.
    define({
      controlId: "XLSCTL.CONTEXT.CHART.AXIS_TITLE_DEPTH.MRAI.APPLY",
      featureId: "spreadsheet.context.chart.setAxisTitleDepth",
      canonicalId: "XLS-INS-CHART-01",
      group: "Chart Design",
      order: 9,
      controlKind: "field",
      predicate: "active-native-chart-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-chart-metadata",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.CHART.VALUE_AXIS_NUMBER_FORMAT.MRAI.APPLY",
      featureId: "spreadsheet.context.chart.setValueAxisNumberFormat",
      canonicalId: "XLS-INS-CHART-01",
      group: "Chart Design",
      order: 10,
      controlKind: "field",
      predicate: "active-native-chart-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-chart-metadata",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.CHART.LINE_MARKER_SYMBOL.MRAI.APPLY",
      featureId: "spreadsheet.context.chart.setLineMarkerSymbol",
      canonicalId: "XLS-INS-CHART-01",
      group: "Chart Design",
      order: 10,
      controlKind: "combobox",
      predicate: "active-native-chart-line-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-chart-metadata",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.CHART.LINE_MARKER_SIZE.MRAI.APPLY",
      featureId: "spreadsheet.context.chart.setLineMarkerSize",
      canonicalId: "XLS-INS-CHART-01",
      group: "Chart Design",
      order: 10,
      controlKind: "field",
      predicate: "active-native-chart-line-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-chart-metadata",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.CHART.SELECT_DATA",
      featureId: "spreadsheet.context.chart.openSelectData",
      canonicalId: "XLS-INS-CHART-01",
      group: "Chart Design",
      order: 10,
      controlKind: "dialog-launcher",
      predicate: "active-native-chart-series-orientation-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.CHART.SELECT_DATA.MRAI.APPLY",
      featureId: "spreadsheet.context.chart.selectData",
      canonicalId: "XLS-INS-CHART-01",
      group: "Chart Design",
      order: 11,
      controlKind: "button",
      predicate: "active-native-chart-series-orientation-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-charts",
    }),
    ...[
      ["XLSCTL.CONTEXT.CHART.LEGEND", "toggleLegend", 12, "toggle"],
      ["XLSCTL.CONTEXT.CHART.LEGEND_POSITION.MRAI.APPLY", "setLegendPosition", 13, "menu"],
      // Excel's Chart Elements > Data Labels > position submenu (and a filled
      // map's Map labels tri-state): the same placement pick the deck and
      // Docs make through `chart-elements.chartWithDataLabels`.
      ["XLSCTL.CONTEXT.CHART.DATA_LABEL_POSITION.MRAI.APPLY", "setDataLabelPosition", 13, "menu"],
      ["XLSCTL.CONTEXT.CHART.AXES", "toggleAxes", 14, "toggle"],
      ["XLSCTL.CONTEXT.CHART.GRIDLINES", "toggleGridlines", 15, "toggle"],
      ["XLSCTL.CONTEXT.CHART.DATA_TABLE", "toggleDataTable", 16, "toggle"],
    ].map(([controlId, suffix, order, controlKind]) =>
      define({
        controlId: String(controlId),
        featureId: `spreadsheet.context.chart.${String(suffix)}`,
        canonicalId: "XLS-INS-CHART-01",
        group: "Chart Design",
        order: Number(order),
        controlKind: controlKind as SpreadsheetCapabilityControlKind,
        predicate: "active-native-chart-editable",
        reasonKey: "spreadsheet.capability.writerUnavailable",
        command: "bound",
        undoPolicy: "single-transaction",
        ioProfile: "worksheet-chart-metadata",
      }),
    ),
    define({
      controlId: "XLSCTL.CONTEXT.CHART.LINES_MENU",
      featureId: "spreadsheet.context.chart.linesMenu",
      canonicalId: "XLS-INS-CHART-01",
      group: "Chart Design",
      order: 19,
      controlKind: "menu",
      predicate: "menu-children",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "disclosure",
      undoPolicy: "content-none",
      ioProfile: "session-only",
      children: ["XLSCTL.CONTEXT.CHART.LINES.MRAI.APPLY"],
    }),
    ...[
      ["XLSCTL.CONTEXT.CHART.TRENDLINE.MRAI.APPLY", "setTrendline", 17, "menu"],
      ["XLSCTL.CONTEXT.CHART.ERROR_BARS.MRAI.APPLY", "setErrorBars", 18, "menu"],
      ["XLSCTL.CONTEXT.CHART.UP_DOWN_BARS.MRAI.APPLY", "toggleUpDownBars", 20, "menu"],
      ["XLSCTL.CONTEXT.CHART.LINES.MRAI.APPLY", "setLines", 21, "menu"],
      ["XLSCTL.CONTEXT.CHART.SECONDARY_AXIS.MRAI.APPLY", "setSecondaryAxis", 22, "field"],
      // Format Data Labels on one series or one label (the Format pane's
      // series / data point page): the same dataLabelOverride intent the
      // deck and Docs apply, as a JSON parameter {target, patch}.
      ["XLSCTL.CONTEXT.CHART.DATA_LABEL_OVERRIDE.MRAI.APPLY", "setDataLabelOverride", 23, "field"],
      // Hidden and Empty Cells ▸ Show empty cells as (c:dispBlanksAs).
      ["XLSCTL.CONTEXT.CHART.DISPLAY_BLANKS.MRAI.APPLY", "setDisplayBlanks", 24, "menu"],
      // Format Axis ▸ Bins of a histogram or pareto (cx:binning).
      ["XLSCTL.CONTEXT.CHART.HISTOGRAM_BINS.MRAI.APPLY", "setHistogramBins", 25, "field"],
    ].map(([controlId, suffix, order, controlKind]) =>
      define({
        controlId: String(controlId),
        featureId: `spreadsheet.context.chart.${String(suffix)}`,
        canonicalId: "XLS-INS-CHART-01",
        group: "Chart Design",
        order: Number(order),
        controlKind: controlKind as SpreadsheetCapabilityControlKind,
        predicate: "active-native-chart-editable",
        reasonKey: "spreadsheet.capability.writerUnavailable",
        command: "bound",
        undoPolicy: "single-transaction",
        ioProfile: "worksheet-chart-metadata",
      }),
    ),
    define({
      controlId: "XLSCTL.CONTEXT.CHART.FORMAT_PANE",
      featureId: "spreadsheet.context.chart.toggleFormatPane",
      canonicalId: "XLS-INS-CHART-01",
      group: "Chart Design",
      order: 23,
      controlKind: "toggle",
      predicate: "active-native-chart",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
  ] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

/**
 * Selection-owned PivotTable Analyze surface. Session commands inspect or
 * navigate the current derived PivotTable without dirtying the workbook. The
 * MRAI commit leaves are the only persistent recipe writers and therefore share
 * the fail-closed `worksheet-pivots` route.
 */
/**
 * PivotChart Analyze (the selected PivotChart's own tab): Field Buttons
 * (Show Report Filter / Legend / Axis / Value / Expand-Collapse Entire Field
 * Buttons, Hide All) and the Active Field's Expand/Collapse Field.
 */
export const SPREADSHEET_PIVOT_CHART_ANALYZE_CONTROL_DEFINITIONS = Object.freeze([
  define({
    controlId: "XLSCTL.CONTEXT.PIVOTCHART.FIELD_BUTTONS_MENU",
    featureId: "spreadsheet.context.pivotChart.fieldButtonsMenu",
    canonicalId: "XLS-PVT-CHART-01",
    group: "PivotChart Analyze",
    order: 1,
    controlKind: "menu",
    predicate: "menu-children",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    command: "disclosure",
    undoPolicy: "content-none",
    ioProfile: "session-only",
    children: ["XLSCTL.CONTEXT.PIVOTCHART.FIELD_BUTTONS.MRAI.APPLY"],
  }),
  define({
    controlId: "XLSCTL.CONTEXT.PIVOTCHART.FIELD_BUTTONS.MRAI.APPLY",
    featureId: "spreadsheet.context.pivotChart.fieldButtons",
    canonicalId: "XLS-PVT-CHART-01",
    group: "PivotChart Analyze",
    order: 2,
    controlKind: "button",
    predicate: "active-pivot-chart-options-editable",
    reasonKey: "spreadsheet.capability.writerUnavailable",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-charts",
  }),
  define({
    controlId: "XLSCTL.CONTEXT.PIVOTCHART.EXPAND_FIELD.MRAI.APPLY",
    featureId: "spreadsheet.context.pivotChart.expandField",
    canonicalId: "XLS-PVT-CHART-01",
    group: "PivotChart Analyze",
    order: 3,
    controlKind: "button",
    predicate: "active-pivot-chart-editable",
    reasonKey: "spreadsheet.capability.writerUnavailable",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-pivots",
  }),
  define({
    controlId: "XLSCTL.CONTEXT.PIVOTCHART.COLLAPSE_FIELD.MRAI.APPLY",
    featureId: "spreadsheet.context.pivotChart.collapseField",
    canonicalId: "XLS-PVT-CHART-01",
    group: "PivotChart Analyze",
    order: 4,
    controlKind: "button",
    predicate: "active-pivot-chart-editable",
    reasonKey: "spreadsheet.capability.writerUnavailable",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-pivots",
  }),
]);

export const SPREADSHEET_PIVOT_ANALYZE_VERTICAL_CONTROL_DEFINITIONS =
  Object.freeze([
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.REFRESH",
      featureId: "spreadsheet.context.pivot.refreshSession",
      canonicalId: "XLS-PVT-SOURCE-01",
      group: "PivotTable Analyze",
      order: 1,
      controlKind: "button",
      predicate: "active-pivot",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.REFRESH.MRAI.APPLY",
      featureId: "spreadsheet.context.pivot.refresh",
      canonicalId: "XLS-PVT-SOURCE-01",
      group: "PivotTable Analyze",
      order: 2,
      controlKind: "button",
      predicate: "active-pivot-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-pivots",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.CHANGE_SOURCE",
      featureId: "spreadsheet.context.pivot.changeSource",
      canonicalId: "XLS-PVT-SOURCE-01",
      group: "PivotTable Analyze",
      order: 3,
      controlKind: "dialog-launcher",
      predicate: "active-pivot-source-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.CHANGE_SOURCE.MRAI.APPLY",
      featureId: "spreadsheet.context.pivot.applySource",
      canonicalId: "XLS-PVT-SOURCE-01",
      group: "PivotTable Analyze",
      order: 4,
      controlKind: "button",
      predicate: "active-pivot-source-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-pivots",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.FIELD_LIST",
      featureId: "spreadsheet.context.pivot.toggleFieldList",
      canonicalId: "XLS-PVT-FIELD-01",
      group: "PivotTable Analyze",
      order: 5,
      controlKind: "toggle",
      predicate: "active-pivot",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.SELECT_MENU",
      featureId: "spreadsheet.context.pivot.selectMenu",
      canonicalId: "XLS-PVT-ACTION-01",
      group: "PivotTable Analyze",
      order: 6,
      controlKind: "menu",
      predicate: "menu-children",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "disclosure",
      undoPolicy: "content-none",
      ioProfile: "session-only",
      // Excel's order: Labels and Values, Values, Labels, Entire PivotTable.
      children: [
        "XLSCTL.CONTEXT.PIVOT.SELECT_LABELS_VALUES.MRAI",
        "XLSCTL.CONTEXT.PIVOT.SELECT_VALUES.MRAI",
        "XLSCTL.CONTEXT.PIVOT.SELECT_LABELS.MRAI",
        "XLSCTL.CONTEXT.PIVOT.SELECT_ENTIRE.MRAI",
      ],
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.SELECT_LABELS_VALUES.MRAI",
      featureId: "spreadsheet.context.pivot.selectLabelsValues",
      canonicalId: "XLS-PVT-ACTION-01",
      group: "PivotTable Analyze",
      order: 7,
      controlKind: "button",
      predicate: "active-pivot-output",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.SELECT_ENTIRE.MRAI",
      featureId: "spreadsheet.context.pivot.selectEntire",
      canonicalId: "XLS-PVT-ACTION-01",
      group: "PivotTable Analyze",
      order: 7,
      controlKind: "button",
      predicate: "active-pivot-output",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.SELECT_LABELS.MRAI",
      featureId: "spreadsheet.context.pivot.selectLabels",
      canonicalId: "XLS-PVT-ACTION-01",
      group: "PivotTable Analyze",
      order: 8,
      controlKind: "button",
      predicate: "active-pivot-output",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.SELECT_VALUES.MRAI",
      featureId: "spreadsheet.context.pivot.selectValues",
      canonicalId: "XLS-PVT-ACTION-01",
      group: "PivotTable Analyze",
      order: 9,
      controlKind: "button",
      predicate: "active-pivot-output",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.CLEAR_MENU",
      featureId: "spreadsheet.context.pivot.clearMenu",
      canonicalId: "XLS-PVT-ACTION-01",
      group: "PivotTable Analyze",
      order: 10,
      controlKind: "menu",
      predicate: "menu-children",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "disclosure",
      undoPolicy: "content-none",
      ioProfile: "session-only",
      children: [
        "XLSCTL.CONTEXT.PIVOT.CLEAR_FILTERS.MRAI.APPLY",
        "XLSCTL.CONTEXT.PIVOT.CLEAR_ALL.MRAI.APPLY",
        "XLSCTL.CONTEXT.PIVOT.CLEAR.MRAI.APPLY",
      ],
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.CLEAR_FILTERS.MRAI.APPLY",
      featureId: "spreadsheet.context.pivot.clearFilters",
      canonicalId: "XLS-PVT-ACTION-01",
      group: "PivotTable Analyze",
      order: 11,
      controlKind: "button",
      predicate: "active-pivot-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-pivots",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.CLEAR_ALL.MRAI.APPLY",
      featureId: "spreadsheet.context.pivot.clearAll",
      canonicalId: "XLS-PVT-ACTION-01",
      group: "PivotTable Analyze",
      order: 12,
      controlKind: "button",
      predicate: "active-pivot-clear-all-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-pivots",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.MOVE",
      featureId: "spreadsheet.context.pivot.move",
      canonicalId: "XLS-PVT-ACTION-01",
      group: "PivotTable Analyze",
      order: 13,
      controlKind: "dialog-launcher",
      predicate: "active-pivot-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.MOVE.MRAI.APPLY",
      featureId: "spreadsheet.context.pivot.applyMove",
      canonicalId: "XLS-PVT-ACTION-01",
      group: "PivotTable Analyze",
      order: 14,
      controlKind: "button",
      predicate: "active-pivot-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-pivots",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.CHART_MENU",
      featureId: "spreadsheet.context.pivot.chartMenu",
      canonicalId: "XLS-PVT-CHART-01",
      group: "PivotTable Analyze",
      order: 15,
      controlKind: "menu",
      predicate: "menu-children",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "disclosure",
      undoPolicy: "content-none",
      ioProfile: "session-only",
      children: ["XLSCTL.CONTEXT.PIVOT.CHART.MRAI.APPLY"],
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.CHART_PREVIEW.MRAI",
      featureId: "spreadsheet.context.pivot.toggleChartPreview",
      canonicalId: "XLS-PVT-CHART-01",
      group: "PivotTable Analyze",
      order: 16,
      controlKind: "toggle",
      predicate: "active-pivot-output",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.CHART.MRAI.APPLY",
      featureId: "spreadsheet.context.pivot.createChart",
      canonicalId: "XLS-PVT-CHART-01",
      group: "PivotTable Analyze",
      order: 17,
      controlKind: "button",
      predicate: "active-pivot-chart-create-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-charts",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.FIELD_LIST.MRAI.APPLY",
      featureId: "spreadsheet.context.pivot.applyFieldList",
      canonicalId: "XLS-PVT-FIELD-01",
      group: "PivotTable Analyze",
      order: 18,
      controlKind: "button",
      predicate: "active-pivot-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-pivots",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.VALUE_FIELD_SETTINGS",
      featureId: "spreadsheet.context.pivot.openValueFieldSettings",
      canonicalId: "XLS-DAT-PIVOT-01",
      group: "PivotTable Analyze",
      order: 19,
      controlKind: "dialog-launcher",
      predicate: "active-pivot-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.VALUE_FIELD_SETTINGS.NUMBER_FORMAT",
      featureId: "spreadsheet.context.pivot.openValueNumberFormat",
      canonicalId: "XLS-DAT-PIVOT-01",
      group: "PivotTable Analyze",
      order: 20,
      controlKind: "dialog-launcher",
      predicate: "active-pivot-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.VALUE_FIELD_SETTINGS.MRAI.APPLY",
      featureId: "spreadsheet.context.pivot.applyValueFieldSettings",
      canonicalId: "XLS-DAT-PIVOT-01",
      group: "PivotTable Analyze",
      order: 21,
      controlKind: "button",
      predicate: "active-pivot-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-pivots",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.CLEAR.MRAI.APPLY",
      featureId: "spreadsheet.context.pivot.clear",
      canonicalId: "XLS-PVT-ACTION-01",
      group: "PivotTable Analyze",
      order: 22,
      controlKind: "button",
      predicate: "active-pivot-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-pivots",
    }),
    define({
      controlId:
        "XLSCTL.CONTEXT.PIVOT.SHOW_EXPAND_COLLAPSE_BUTTONS.MRAI.APPLY",
      featureId: "spreadsheet.context.pivot.toggleExpandCollapseButtons",
      canonicalId: "XLS-DAT-PIVOT-01",
      group: "PivotTable Analyze",
      order: 23,
      controlKind: "toggle",
      predicate: "active-pivot-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-pivots",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.DESIGN.SUBTOTALS_MENU",
      featureId: "spreadsheet.context.pivot.designSubtotalsMenu",
      canonicalId: "XLS-DAT-PIVOT-01",
      group: "PivotTable Design",
      order: 16,
      controlKind: "menu",
      predicate: "menu-children",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "disclosure",
      undoPolicy: "content-none",
      ioProfile: "session-only",
      children: ["XLSCTL.CONTEXT.PIVOT.DESIGN.MRAI.APPLY"],
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.DESIGN.GRAND_TOTALS_MENU",
      featureId: "spreadsheet.context.pivot.designGrandTotalsMenu",
      canonicalId: "XLS-DAT-PIVOT-01",
      group: "PivotTable Design",
      order: 17,
      controlKind: "menu",
      predicate: "menu-children",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "disclosure",
      undoPolicy: "content-none",
      ioProfile: "session-only",
      children: ["XLSCTL.CONTEXT.PIVOT.DESIGN.MRAI.APPLY"],
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.DESIGN.REPORT_LAYOUT_MENU",
      featureId: "spreadsheet.context.pivot.designReportLayoutMenu",
      canonicalId: "XLS-DAT-PIVOT-01",
      group: "PivotTable Design",
      order: 18,
      controlKind: "menu",
      predicate: "menu-children",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "disclosure",
      undoPolicy: "content-none",
      ioProfile: "session-only",
      children: ["XLSCTL.CONTEXT.PIVOT.DESIGN.MRAI.APPLY"],
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.DESIGN.BLANK_ROWS_MENU",
      featureId: "spreadsheet.context.pivot.designBlankRowsMenu",
      canonicalId: "XLS-DAT-PIVOT-01",
      group: "PivotTable Design",
      order: 19,
      controlKind: "menu",
      predicate: "menu-children",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "disclosure",
      undoPolicy: "content-none",
      ioProfile: "session-only",
      children: ["XLSCTL.CONTEXT.PIVOT.DESIGN.MRAI.APPLY"],
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.DESIGN.MRAI.APPLY",
      featureId: "spreadsheet.context.pivot.applyDesign",
      canonicalId: "XLS-DAT-PIVOT-01",
      group: "PivotTable Design",
      order: 20,
      controlKind: "gallery",
      predicate: "active-pivot-editable",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-pivots",
    }),
    // Excel for the web inserts PivotTable slicers from PivotTable Analyze.
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.INSERT_SLICER",
      featureId: "spreadsheet.context.pivot.insertSlicer",
      canonicalId: "XLS-INS-FILTER-01",
      group: "PivotTable Analyze",
      order: 24,
      controlKind: "dialog-launcher",
      predicate: "active-pivot-slicer-insert",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
    define({
      controlId: "XLSCTL.CONTEXT.PIVOT.INSERT_SLICER.MRAI.APPLY",
      featureId: "spreadsheet.context.pivot.applySlicers",
      canonicalId: "XLS-INS-FILTER-01",
      group: "PivotTable Analyze",
      order: 25,
      controlKind: "button",
      predicate: "pivot-slicer-insert-apply",
      reasonKey: "spreadsheet.capability.writerUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-slicers",
    }),
  ] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

/**
 * One exact physical row per verified visible command across the requested Web
 * tabs, plus the explicitly separated Home groundwork above. Menu roots own
 * only disclosure; their enabled state is derived from these exact child rows,
 * so an empty or all-inert menu cannot present as available.
 */
export const SPREADSHEET_VISIBLE_CONTROL_DEFINITIONS = Object.freeze([
  ...SPREADSHEET_HOME_GROUNDWORK_CONTROL_DEFINITIONS,
  ...SPREADSHEET_HOME_CORE_VERTICAL_CONTROL_DEFINITIONS,
  ...SPREADSHEET_CONDITIONAL_FORMAT_MANAGER_CONTROL_DEFINITIONS,
  ...SPREADSHEET_PAGE_SETUP_PRINT_VERTICAL_CONTROL_DEFINITIONS,
  ...SPREADSHEET_FORMULAS_VERTICAL_CONTROL_DEFINITIONS,
  ...SPREADSHEET_INSERT_VERTICAL_CONTROL_DEFINITIONS,
  ...SPREADSHEET_DATA_VALIDATION_VERTICAL_CONTROL_DEFINITIONS,
  define({
    controlId: "XLSCTL.INSERT.TABLE",
    featureId: "spreadsheet.insert.createTable",
    canonicalId: "XLS-INS-TABLE-01",
    group: "Insert",
    order: 3,
    controlKind: "dialog-launcher",
    predicate: "table-create",
    reasonKey: "spreadsheet.capability.tableAlreadyActive",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-table",
  }),
  define({
    controlId: "XLSCTL.SHARE.WORKBOOK",
    featureId: "spreadsheet.share.workbook",
    canonicalId: "XLS-SHL-SHARE-01",
    group: "Share",
    order: 1,
    controlKind: "button",
    predicate: "share-web",
    reasonKey: "spreadsheet.capability.shareUnavailable",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.SHARE.LINK.COPY",
    featureId: "spreadsheet.share.copyLink",
    canonicalId: "XLS-SHL-SHARE-01",
    group: "Share",
    order: 2,
    controlKind: "button",
    predicate: "share-clipboard",
    reasonKey: "spreadsheet.capability.clipboardUnavailable",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.SHARE.LINK.WORKSHEET",
    featureId: "spreadsheet.share.copyWorksheetLink",
    canonicalId: "XLS-SHL-SHARE-01",
    group: "Share",
    order: 3,
    controlKind: "button",
    predicate: "share-clipboard",
    reasonKey: "spreadsheet.capability.clipboardUnavailable",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.SHARE.PUBLISH.WEB",
    featureId: "spreadsheet.share.publishWeb",
    canonicalId: "XLS-SHL-SHARE-01",
    group: "Share",
    order: 4,
    controlKind: "button",
    predicate: "no-verified-authority",
    reasonKey: "spreadsheet.capability.publicPublishUnavailable",
    command: "unsupported",
    undoPolicy: "content-none",
    ioProfile: "unsupported",
  }),
  define({
    // Desktop Excel's Background / Delete Background (Excel for the web has
    // none): one picture tiled behind the sheet, never printed.
    controlId: "XLSCTL.PAGE_LAYOUT.BACKGROUND",
    featureId: "spreadsheet.pageLayout.sheetBackground",
    canonicalId: "XLS-PGL-SETUP-01",
    group: "Page Layout",
    order: 12,
    controlKind: "button",
    predicate: "sheet-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-background-picture",
  }),
  define({
    controlId: "XLSCTL.PAGE_LAYOUT.GRIDLINES.VIEW",
    featureId: "spreadsheet.view.toggleGridlines",
    canonicalId: "XLS-VIE-SHOW-01",
    group: "Page Layout",
    order: 13,
    controlKind: "toggle",
    predicate: "always",
    reasonKey: null,
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "workbook-view-preference",
  }),
  define({
    controlId: "XLSCTL.PAGE_LAYOUT.HEADINGS.VIEW",
    featureId: "spreadsheet.view.toggleHeadings",
    canonicalId: "XLS-VIE-SHOW-01",
    group: "Page Layout",
    order: 15,
    controlKind: "toggle",
    predicate: "always",
    reasonKey: null,
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "workbook-view-preference",
  }),
  define({
    controlId: "XLSCTL.DRAW.ACTION_PEN",
    featureId: "spreadsheet.draw.actionPen",
    canonicalId: "XLS-INS-SHAPE-01",
    group: "Draw",
    order: 10,
    controlKind: "button",
    predicate: "no-verified-authority",
    reasonKey: "spreadsheet.capability.actionPenUnavailable",
    command: "unsupported",
    undoPolicy: "content-none",
    ioProfile: "unsupported",
  }),
  define({
    controlId: "XLSCTL.DATA.QUERY.GET_DATA_MENU",
    featureId: "spreadsheet.data.getDataMenu",
    canonicalId: "XLS-DAT-CONN-01",
    group: "Data",
    order: 1,
    controlKind: "menu",
    predicate: "menu-children",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    command: "disclosure",
    undoPolicy: "content-none",
    ioProfile: "session-only",
    children: [
      "XLSCTL.DATA.QUERY.GET_DATA.FROM_TEXT_CSV",
      "XLSCTL.DATA.QUERY.GET_DATA.FROM_WORKBOOK",
      "XLSCTL.DATA.QUERY.FROM_PICTURE_MENU",
      "XLSCTL.DATA.QUERY.GET_DATA.FROM_POWER_BI",
    ],
  }),
  define({
    controlId: "XLSCTL.DATA.QUERY.GET_DATA.FROM_TEXT_CSV",
    featureId: "spreadsheet.data.importTextCsv",
    canonicalId: "XLS-DAT-IMPORT-01",
    group: "Data",
    order: 2,
    controlKind: "button",
    predicate: "import-ready-editable",
    reasonKey: "spreadsheet.capability.importInProgress",
    command: "bound",
    undoPolicy: "single-transaction",
    // From Text/CSV places the file's values in the active sheet at the
    // active cell and never replaces the workbook (the input refuses an
    // .xlsx), so it is a cell edit the in-place patch writes (audit S7).
    ioProfile: "worksheet-cells",
  }),
  define({
    controlId: "XLSCTL.DATA.QUERY.GET_DATA.FROM_WORKBOOK",
    featureId: "spreadsheet.data.importWorkbook",
    canonicalId: "XLS-DAT-IMPORT-01",
    group: "Data",
    order: 3,
    controlKind: "button",
    predicate: "import-ready",
    reasonKey: "spreadsheet.capability.importInProgress",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "workbook-import",
  }),
  define({
    controlId: "XLSCTL.DATA.QUERY.FROM_PICTURE_MENU",
    featureId: "spreadsheet.data.pictureMenu",
    canonicalId: "XLS-INS-DATA-01",
    group: "Data",
    order: 4,
    controlKind: "menu",
    predicate: "menu-children",
    reasonKey: "spreadsheet.capability.unsupported",
    command: "disclosure",
    undoPolicy: "content-none",
    ioProfile: "session-only",
    children: [
      "XLSCTL.DATA.QUERY.FROM_PICTURE.FROM_FILE",
      "XLSCTL.DATA.QUERY.FROM_PICTURE.FROM_CLIPBOARD",
    ],
  }),
  define({
    controlId: "XLSCTL.DATA.QUERY.FROM_PICTURE.FROM_FILE",
    featureId: "spreadsheet.data.pictureFromFile",
    canonicalId: "XLS-INS-DATA-01",
    group: "Data",
    order: 4,
    controlKind: "button",
    predicate: "data-from-picture",
    reasonKey: "spreadsheet.capability.providerUnavailable",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.DATA.QUERY.FROM_PICTURE.FROM_CLIPBOARD",
    featureId: "spreadsheet.data.pictureFromClipboard",
    canonicalId: "XLS-INS-DATA-01",
    group: "Data",
    order: 5,
    controlKind: "button",
    predicate: "data-from-picture-clipboard",
    reasonKey: "spreadsheet.capability.clipboardUnavailable",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.DATA.QUERY.GET_DATA.FROM_POWER_BI",
    featureId: "spreadsheet.data.fromPowerBi",
    canonicalId: "XLS-DAT-CONN-01",
    group: "Data",
    order: 6,
    controlKind: "button",
    predicate: "no-verified-authority",
    reasonKey: "spreadsheet.capability.powerBiUnavailable",
    command: "unsupported",
    undoPolicy: "content-none",
    ioProfile: "unsupported",
  }),
  define({
    controlId: "XLSCTL.DATA.QUERY.FROM_TABLE_RANGE",
    featureId: "spreadsheet.data.fromTableRange",
    canonicalId: "XLS-DAT-CONN-01",
    group: "Data",
    order: 7,
    controlKind: "button",
    predicate: "data-query-create",
    reasonKey: "spreadsheet.capability.selectionRequired",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "worksheet-queries",
  }),
  define({
    controlId: "XLSCTL.DATA.QUERY.REFRESH_ALL_MENU",
    featureId: "spreadsheet.data.refreshAll",
    canonicalId: "XLS-DAT-CONN-01",
    group: "Data",
    order: 8,
    controlKind: "button",
    predicate: "data-query-refresh-all",
    reasonKey: "spreadsheet.capability.noQueries",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-queries",
  }),
  define({
    controlId: "XLSCTL.DATA.QUERY.QUERIES_CONNECTIONS",
    featureId: "spreadsheet.data.queriesConnections",
    canonicalId: "XLS-DAT-CONN-01",
    group: "Data",
    order: 9,
    controlKind: "dialog-launcher",
    // Live with no query at all, as Excel's (measured): the pane lists none.
    predicate: "always",
    reasonKey: "spreadsheet.capability.noQueries",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.DATA.QUERY.PROPERTIES",
    featureId: "spreadsheet.data.connectionProperties",
    canonicalId: "XLS-DAT-CONN-01",
    group: "Data",
    order: 10,
    controlKind: "dialog-launcher",
    // External Data Properties of the query table the active cell is in,
    // with its Query Properties; each OK is one edit of the query's
    // connection / queryTable settings (a rename loads the query again).
    predicate: "data-query-properties",
    reasonKey: "spreadsheet.capability.queryTableSelectionRequired",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-queries",
  }),
  define({
    controlId: "XLSCTL.DATA.QUERY.WORKBOOK_LINKS",
    featureId: "spreadsheet.data.workbookLinks",
    canonicalId: "XLS-DAT-CONN-01",
    group: "Data",
    order: 11,
    controlKind: "dialog-launcher",
    predicate: "data-source-inventory",
    reasonKey: "spreadsheet.capability.noWorkbookLinks",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.DATA.QUERY.REFRESH",
    featureId: "spreadsheet.data.refresh",
    canonicalId: "XLS-DAT-CONN-01",
    group: "Data",
    order: 12,
    controlKind: "button",
    predicate: "data-query-refresh",
    reasonKey: "spreadsheet.capability.noQueries",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-queries",
  }),
  define({
    controlId: "XLSCTL.DATA.FORM.MRAI.DIALOG",
    featureId: "spreadsheet.data.openDataForm",
    canonicalId: "XLS-DAT-FORM-01",
    group: "Data",
    order: 45,
    controlKind: "dialog-launcher",
    predicate: "data-form-launchable",
    reasonKey: "spreadsheet.capability.dataFormListRequired",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  ...[
    ["XLSCTL.DATA.FORM.MRAI.PREVIOUS", "previousDataFormRecord", 46],
    ["XLSCTL.DATA.FORM.MRAI.NEXT", "nextDataFormRecord", 47],
    ["XLSCTL.DATA.FORM.MRAI.NEW", "newDataFormRecord", 48],
  ].map(([controlId, suffix, order]) =>
    define({
      controlId: String(controlId),
      featureId: `spreadsheet.data.${String(suffix)}`,
      canonicalId: "XLS-DAT-FORM-01",
      group: "Data",
      order: Number(order),
      controlKind: "button",
      predicate: "data-form-open",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-cells",
    }),
  ),
  define({
    controlId: "XLSCTL.DATA.FORM.MRAI.DELETE",
    featureId: "spreadsheet.data.requestDeleteDataFormRecord",
    canonicalId: "XLS-DAT-FORM-01",
    group: "Data",
    order: 49,
    controlKind: "button",
    predicate: "data-form-delete",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.DATA.FORM.MRAI.DELETE.APPLY",
    featureId: "spreadsheet.data.applyDeleteDataFormRecord",
    canonicalId: "XLS-DAT-FORM-01",
    group: "Data",
    order: 50,
    controlKind: "button",
    predicate: "data-form-delete-apply",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-data-form-delete",
  }),
  ...[
    ["XLSCTL.DATA.FORM.MRAI.CRITERIA", "toggleDataFormCriteria", 51],
    ["XLSCTL.DATA.FORM.MRAI.RESTORE", "restoreDataFormRecord", 52],
    ["XLSCTL.DATA.FORM.MRAI.DISMISS", "closeDataForm", 53],
  ].map(([controlId, suffix, order]) =>
    define({
      controlId: String(controlId),
      featureId: `spreadsheet.data.${String(suffix)}`,
      canonicalId: "XLS-DAT-FORM-01",
      group: "Data",
      order: Number(order),
      controlKind: "button",
      predicate:
        suffix === "restoreDataFormRecord"
          ? "data-form-restore"
          : "data-form-open",
      reasonKey: "spreadsheet.capability.commandUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
  ),
  ...[
    ["XLSCTL.DATA.TYPE.STOCKS", "stocks", 13],
    ["XLSCTL.DATA.TYPE.CURRENCIES", "currencies", 14],
  ].map(([controlId, suffix, order]) =>
    define({
      controlId: String(controlId),
      featureId: `spreadsheet.data.${String(suffix)}`,
      canonicalId: "XLS-DAT-TYPE-01",
      group: "Data",
      order: Number(order),
      controlKind: "button",
      predicate: "no-verified-authority",
      reasonKey: "spreadsheet.capability.dataTypesUnavailable",
      command: "unsupported",
      undoPolicy: "content-none",
      ioProfile: "unsupported",
    }),
  ),
  ...[
    ["XLSCTL.DATA.SORT.ASC", "sortAscending", 12],
    ["XLSCTL.DATA.SORT.DESC", "sortDescending", 13],
    ["XLSCTL.DATA.SORT.DIALOG", "openSortDialog", 14],
  ].map(([controlId, suffix, order]) =>
    define({
      controlId: String(controlId),
      featureId: `spreadsheet.data.${String(suffix)}`,
      canonicalId: "XLS-DAT-SORT-01",
      group: "Data",
      order: Number(order),
      controlKind: suffix === "openSortDialog" ? "dialog-launcher" : "button",
      predicate: "sort-authorized",
      reasonKey: "spreadsheet.capability.sheetProtected",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-cells",
    }),
  ),
  define({
    controlId: "XLSCTL.DATA.SORT.FILTER",
    featureId: "spreadsheet.data.toggleAutoFilter",
    canonicalId: "XLS-DAT-FILTER-01",
    group: "Data",
    order: 15,
    controlKind: "toggle",
    predicate: "filter-editable",
    reasonKey: "spreadsheet.capability.selectionRequired",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-filters",
  }),
  define({
    controlId: "XLSCTL.DATA.SORT.CLEAR",
    featureId: "spreadsheet.data.clearFilters",
    canonicalId: "XLS-DAT-FILTER-01",
    group: "Data",
    order: 16,
    controlKind: "button",
    predicate: "active-filter",
    reasonKey: "spreadsheet.capability.activeFilterRequired",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-filters",
  }),
  define({
    controlId: "XLSCTL.DATA.SORT.REAPPLY",
    featureId: "spreadsheet.data.reapplyFilters",
    canonicalId: "XLS-DAT-FILTER-01",
    group: "Data",
    order: 17,
    controlKind: "button",
    predicate: "reapplicable-filter-or-sort",
    reasonKey: "spreadsheet.capability.reapplyFilterOrSortRequired",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-filters",
  }),
  define({
    controlId: "XLSCTL.DATA.SORT.ADVANCED",
    featureId: "spreadsheet.data.openAdvancedFilter",
    canonicalId: "XLS-DAT-FILTER-01",
    group: "Data",
    order: 18,
    controlKind: "dialog-launcher",
    predicate: "selection-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    command: "bound",
    // Opening the dialog is non-mutating. The confirmed in-place/copy mode is
    // resolved separately because the two modes have different native writers.
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  ...[
    [
      "XLSCTL.DATA.TOOLS.TEXT_TO_COLUMNS",
      "openTextToColumns",
      19,
      "selection-editable",
    ],
    ["XLSCTL.DATA.TOOLS.FLASH_FILL", "flashFill", 20, "selection-editable"],
    [
      "XLSCTL.DATA.TOOLS.REMOVE_DUPLICATES",
      "openRemoveDuplicates",
      21,
      "remove-duplicates-editable",
    ],
  ].map(([controlId, suffix, order, predicate]) =>
    define({
      controlId: String(controlId),
      featureId: `spreadsheet.data.${String(suffix)}`,
      canonicalId: "XLS-DAT-TOOLS-01",
      group: "Data",
      order: Number(order),
      controlKind: String(suffix).startsWith("open")
        ? "dialog-launcher"
        : "button",
      predicate: predicate as SpreadsheetVisiblePredicate,
      reasonKey:
        predicate === "multi-selection-editable" || predicate === "remove-duplicates-editable"
          ? "spreadsheet.capability.multiSelectionRequired"
          : "spreadsheet.capability.sheetProtected",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-cells",
    }),
  ),
  define({
    controlId: "XLSCTL.DATA.OUTLINE.GROUP_MENU",
    featureId: "spreadsheet.data.groupMenu",
    canonicalId: "XLS-DAT-OUTLINE-01",
    group: "Data",
    order: 24,
    controlKind: "menu",
    predicate: "menu-children",
    reasonKey: "spreadsheet.capability.selectionRequired",
    command: "disclosure",
    undoPolicy: "content-none",
    ioProfile: "session-only",
    children: [
      "XLSCTL.DATA.OUTLINE.GROUP.ROWS",
      "XLSCTL.DATA.OUTLINE.GROUP.COLUMNS",
    ],
  }),
  define({
    controlId: "XLSCTL.DATA.OUTLINE.UNGROUP_MENU",
    featureId: "spreadsheet.data.ungroupMenu",
    canonicalId: "XLS-DAT-OUTLINE-01",
    group: "Data",
    order: 25,
    controlKind: "menu",
    // Excel's Ungroup button is live with no outline at all (measured); on a
    // writable sheet it opens and each entry says what it needs.
    predicate: "sheet-editable",
    reasonKey: "spreadsheet.capability.sheetProtected",
    command: "disclosure",
    undoPolicy: "content-none",
    ioProfile: "session-only",
    children: [
      "XLSCTL.DATA.OUTLINE.UNGROUP.ROWS",
      "XLSCTL.DATA.OUTLINE.UNGROUP.COLUMNS",
    ],
  }),
  ...[
    ["XLSCTL.DATA.OUTLINE.GROUP.ROWS", "groupRows", 26, "outline-group-rows"],
    [
      "XLSCTL.DATA.OUTLINE.GROUP.COLUMNS",
      "groupColumns",
      27,
      "outline-group-columns",
    ],
    [
      "XLSCTL.DATA.OUTLINE.UNGROUP.ROWS",
      "ungroupRows",
      28,
      "outline-ungroup-rows",
    ],
    [
      "XLSCTL.DATA.OUTLINE.UNGROUP.COLUMNS",
      "ungroupColumns",
      29,
      "outline-ungroup-columns",
    ],
  ].map(([controlId, suffix, order, predicate]) =>
    define({
      controlId: String(controlId),
      featureId: `spreadsheet.data.${String(suffix)}`,
      canonicalId: "XLS-DAT-OUTLINE-01",
      group: "Data",
      order: Number(order),
      controlKind: "button",
      predicate: predicate as SpreadsheetVisiblePredicate,
      reasonKey: String(controlId).includes(".UNGROUP.")
        ? "spreadsheet.capability.outlineUngroupUnavailable"
        : "spreadsheet.capability.outlineGroupUnavailable",
      command: "bound",
      undoPolicy: "single-transaction",
      ioProfile: "worksheet-outline",
    }),
  ),
  define({
    controlId: "XLSCTL.REVIEW.CHANGES.SHOW",
    featureId: "spreadsheet.review.showChanges",
    canonicalId: "XLS-REV-COMMENT-01",
    group: "Review",
    order: 6,
    controlKind: "dialog-launcher",
    // Excel's Changes pane: the workbook's cell edits, newest first, from the
    // server change log every editor of the page appends to.
    predicate: "change-history",
    reasonKey: "spreadsheet.capability.savedWorkbookRequired",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.REVIEW.COMMENT.NEW",
    featureId: "spreadsheet.review.newComment",
    canonicalId: "XLS-REV-COMMENT-01",
    group: "Review",
    order: 1,
    controlKind: "dialog-launcher",
    predicate: "new-comment",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.REVIEW.COMMENT.POST",
    featureId: "spreadsheet.review.postComment",
    canonicalId: "XLS-REV-COMMENT-01",
    group: "Review",
    order: 1,
    controlKind: "button",
    predicate: "comment-post-editable",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-comments",
  }),
  define({
    controlId: "XLSCTL.REVIEW.COMMENT.RESOLVE_TOGGLE",
    featureId: "spreadsheet.review.toggleCommentResolved",
    canonicalId: "XLS-REV-COMMENT-01",
    group: "Review",
    order: 1,
    controlKind: "toggle",
    predicate: "active-comment-editable",
    reasonKey: "spreadsheet.capability.activeCommentRequired",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-comments",
  }),
  ...[
    // Excel's Show Comments opens its pane on a workbook without comments
    // too (New and a hint); only Previous/Next need a comment to go to.
    ["XLSCTL.REVIEW.COMMENT.SHOW", "showComments", 2, "always"],
    ["XLSCTL.REVIEW.COMMENT.PREVIOUS", "previousComment", 3, "comments"],
    ["XLSCTL.REVIEW.COMMENT.NEXT", "nextComment", 4, "comments"],
    [
      // Excel's Delete removes the active cell's threaded comment OR note.
      "XLSCTL.REVIEW.COMMENT.DELETE",
      "deleteComment",
      5,
      "active-remark-editable",
    ],
    [
      "XLSCTL.REVIEW.CHANGES.VERSION_HISTORY",
      "versionHistory",
      7,
      "version-history",
    ],
    // Excel greys Spelling on a protected sheet.
    ["XLSCTL.REVIEW.PROOF.SPELL", "spelling", 8, "sheet-editable"],
    ["XLSCTL.REVIEW.PROOF.WORKBOOK_STATS", "statistics", 10, "always"],
    ["XLSCTL.REVIEW.PROOF.ACCESSIBILITY_MENU", "accessibility", 11, "always"],
    ["XLSCTL.REVIEW.PROOF.PERFORMANCE", "performance", 12, "always"],
    [
      "XLSCTL.REVIEW.PROTECT.SELECTION_UNLOCK",
      "unlockSelection",
      15,
      "selection-can-unlock",
    ],
    [
      "XLSCTL.REVIEW.PROTECT.SELECTION_LOCK",
      "lockSelection",
      16,
      "selection-can-lock",
    ],
    ["XLSCTL.REVIEW.PROTECT.SHEET", "toggleSheetProtection", 17, "always"],
  ].map(([controlId, suffix, order, predicate]) => {
    return define({
      controlId: String(controlId),
      featureId: `spreadsheet.review.${String(suffix)}`,
      canonicalId:
        String(suffix).includes("Protection") ||
        String(suffix).includes("Selection")
          ? "XLS-REV-PROTECT-01"
          : String(suffix) === "spelling" ||
              String(suffix) === "accessibility" ||
              String(suffix) === "performance" ||
              String(suffix) === "statistics"
            ? "XLS-REV-PROOF-01"
            : "XLS-REV-COMMENT-01",
      group: "Review",
      order: Number(order),
      controlKind:
        String(suffix).startsWith("show") || String(suffix) === "versionHistory"
          ? "dialog-launcher"
          : "button",
      predicate: predicate as SpreadsheetVisiblePredicate,
      reasonKey:
        predicate === "comments"
          ? "spreadsheet.capability.noComments"
          : predicate === "active-comment" ||
              predicate === "active-comment-editable" ||
              predicate === "active-remark-editable"
            ? "spreadsheet.capability.activeCommentRequired"
            : predicate === "version-history"
              ? "spreadsheet.capability.savedWorkbookRequired"
              : predicate === "selection-can-lock" ||
                  predicate === "selection-can-unlock"
                ? "spreadsheet.capability.selectionRequired"
                : predicate === "selection-editable" ||
                    predicate === "sheet-editable"
                  ? "spreadsheet.capability.sheetProtected"
                  : null,
      command: "bound",
      undoPolicy:
        String(suffix).includes("Selection") ||
        String(suffix).includes("Protection") ||
        String(suffix) === "deleteComment"
          ? "single-transaction"
          : "content-none",
      ioProfile:
        String(suffix).includes("Selection") ||
        String(suffix).includes("Protection")
          ? "worksheet-protection"
          : String(suffix) === "deleteComment"
            ? "worksheet-comments"
            : "session-only",
    });
  }),
  ...[
    ["XLSCTL.REVIEW.PROOF.THESAURUS", "thesaurus", 9],
    ["XLSCTL.REVIEW.PROOF.SMART_LOOKUP", "smartLookup", 13],
    ["XLSCTL.REVIEW.PROOF.TRANSLATE", "translate", 14],
  ].map(([controlId, suffix, order]) =>
    define({
      controlId: String(controlId),
      featureId: `spreadsheet.review.${String(suffix)}`,
      canonicalId: "XLS-REV-PROOF-01",
      group: "Review",
      order: Number(order),
      controlKind: "dialog-launcher",
      predicate: "no-verified-authority",
      reasonKey: "spreadsheet.capability.proofingServiceUnavailable",
      command: "unsupported",
      undoPolicy: "content-none",
      ioProfile: "unsupported",
    }),
  ),
  define({
    controlId: "XLSCTL.REVIEW.PROTECT.MANAGE_MENU",
    featureId: "spreadsheet.review.manageProtectionMenu",
    canonicalId: "XLS-REV-PROTECT-01",
    group: "Review",
    order: 18,
    controlKind: "menu",
    predicate: "menu-children",
    reasonKey: "spreadsheet.capability.commandUnavailable",
    command: "disclosure",
    undoPolicy: "content-none",
    ioProfile: "session-only",
    children: [
      "XLSCTL.REVIEW.PROTECT.SELECTION_UNLOCK",
      "XLSCTL.REVIEW.PROTECT.SELECTION_LOCK",
      "XLSCTL.REVIEW.PROTECT.SHEET",
      "XLSCTL.REVIEW.PROTECT.WORKBOOK",
    ],
  }),
  define({
    controlId: "XLSCTL.REVIEW.PROTECT.WORKBOOK",
    featureId: "spreadsheet.review.protectWorkbook",
    canonicalId: "XLS-REV-PROTECT-01",
    group: "Review",
    order: 19,
    controlKind: "dialog-launcher",
    predicate: "workbook-protection-launchable",
    reasonKey: "spreadsheet.capability.workbookProtectionUnavailable",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.REVIEW.PROTECT.WORKBOOK.MRAI.APPLY",
    featureId: "spreadsheet.review.applyWorkbookProtection",
    canonicalId: "XLS-REV-PROTECT-01",
    group: "Review",
    order: 19,
    controlKind: "button",
    predicate: "workbook-protection-applicable",
    reasonKey: "spreadsheet.capability.workbookProtectionUnsupported",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "workbook-structure-protection",
  }),
  define({
    controlId: "XLSCTL.REVIEW.PROTECT.PAUSE",
    featureId: "spreadsheet.review.pauseProtection",
    canonicalId: "XLS-REV-PROTECT-01",
    group: "Review",
    order: 20,
    controlKind: "button",
    predicate: "sheet-protection-configured",
    reasonKey: "spreadsheet.capability.truePauseProtectionUnavailable",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.REVIEW.NOTE.MENU",
    featureId: "spreadsheet.review.notesMenu",
    canonicalId: "XLS-REV-COMMENT-01",
    group: "Review",
    order: 21,
    controlKind: "menu",
    predicate: "menu-children",
    reasonKey: "spreadsheet.capability.notesUnavailable",
    command: "disclosure",
    undoPolicy: "content-none",
    ioProfile: "session-only",
    children: [
      "XLSCTL.REVIEW.NOTE.NEW",
      "XLSCTL.REVIEW.NOTE.SHOW_HIDE",
      "XLSCTL.REVIEW.NOTE.SHOW_ALL",
    ],
  }),
  // Microsoft Support "Print comments and notes in Excel": right-click a cell
  // > Show/Hide Note displays that one note, Review > Notes > Show All Notes
  // displays all of them. Show/Hide Note is the note's own state (its VML
  // `<x:Visible/>`), so it writes the workbook as one undoable edit; Show All
  // Notes is Excel's display setting and stays in the session.
  define({
    controlId: "XLSCTL.REVIEW.NOTE.SHOW_HIDE",
    featureId: "spreadsheet.review.showHideNote",
    canonicalId: "XLS-REV-COMMENT-01",
    group: "Review",
    order: 23,
    controlKind: "button",
    predicate: "active-note-editable",
    reasonKey: "spreadsheet.capability.activeNoteRequired",
    command: "bound",
    undoPolicy: "single-transaction",
    ioProfile: "worksheet-comments",
  }),
  ...[
    ["XLSCTL.REVIEW.NOTE.NEW", "newNote", 22],
    ["XLSCTL.REVIEW.NOTE.SHOW_ALL", "showAllNotes", 24],
  ].map(([controlId, suffix, order]) =>
    define({
      controlId: String(controlId),
      featureId: `spreadsheet.review.${String(suffix)}`,
      canonicalId: "XLS-REV-COMMENT-01",
      group: "Review",
      order: Number(order),
      controlKind: "button",
      predicate: suffix === "newNote" ? "new-note" : "notes",
      // Notes are not comments: the disabled Show All Notes said "no
      // comments" over a workbook whose notes were simply absent.
      reasonKey:
        suffix === "newNote"
          ? "spreadsheet.capability.commandUnavailable"
          : "spreadsheet.capability.noNotes",
      command: "bound",
      undoPolicy: suffix === "newNote" ? "single-transaction" : "content-none",
      ioProfile: suffix === "newNote" ? "worksheet-comments" : "session-only",
    }),
  ),
  ...[
    ["XLSCTL.VIEW.MODE.NORMAL", "normalView", 1],
    ["XLSCTL.VIEW.MODE.PAGE_BREAK", "pageBreakView", 2],
    ["XLSCTL.VIEW.MODE.PAGE_LAYOUT", "pageLayoutView", 3],
  ].map(([controlId, suffix, order]) =>
    define({
      controlId: String(controlId),
      featureId: `spreadsheet.view.${String(suffix)}`,
      canonicalId: "XLS-VIE-MODE-01",
      group: "View",
      order: Number(order),
      controlKind: "toggle",
      predicate: "always",
      reasonKey: null,
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "session-only",
    }),
  ),
  ...[
    [
      "XLSCTL.VIEW.SHOW.HEADINGS",
      "toggleHeadings",
      "XLS-VIE-SHOW-01",
      2,
      "always",
    ],
    [
      "XLSCTL.VIEW.SHOW.GRIDLINES",
      "toggleGridlines",
      "XLS-VIE-SHOW-01",
      3,
      "always",
    ],
    [
      "XLSCTL.VIEW.SHOW.FORMULA_BAR",
      "toggleFormulaBar",
      "XLS-VIE-SHOW-01",
      4,
      "always",
    ],
    [
      "XLSCTL.VIEW.SHOW.FOCUS_CELL",
      "toggleFocusCell",
      "XLS-VIE-SHOW-01",
      5,
      "always",
    ],
    ["XLSCTL.VIEW.ZOOM.VALUE", "commitZoom", "XLS-VIE-ZOOM-01", 6, "always"],
    ["XLSCTL.VIEW.ZOOM.RESET_100", "zoomReset", "XLS-VIE-ZOOM-01", 7, "always"],
    [
      "XLSCTL.VIEW.WINDOW.NEW",
      "newWindow",
      "XLS-VIE-WINDOW-01",
      8,
      "saved-window",
    ],
    [
      "XLSCTL.VIEW.WINDOW.FREEZE",
      "freezePanes",
      "XLS-VIE-WINDOW-01",
      9,
      "always",
    ],
    [
      "XLSCTL.VIEW.WINDOW.FREEZE_TOP_ROW",
      "freezeTopRow",
      "XLS-VIE-WINDOW-01",
      10,
      "always",
    ],
    // Excel for the web lists "Up to row N" / "Up to column X" as their own
    // commands beneath the fixed "Top row" / "First column" entries, so they
    // are their own authority leaves rather than a parameter on those two.
    [
      "XLSCTL.VIEW.WINDOW.FREEZE_ROWS_TO_ACTIVE",
      "freezeRowsToActive",
      "XLS-VIE-WINDOW-01",
      10,
      "always",
    ],
    [
      "XLSCTL.VIEW.WINDOW.FREEZE_FIRST_COLUMN",
      "freezeFirstColumn",
      "XLS-VIE-WINDOW-01",
      11,
      "always",
    ],
    [
      "XLSCTL.VIEW.WINDOW.FREEZE_COLUMNS_TO_ACTIVE",
      "freezeColumnsToActive",
      "XLS-VIE-WINDOW-01",
      11,
      "always",
    ],
    [
      "XLSCTL.VIEW.WINDOW.SPLIT",
      "splitPanes",
      "XLS-VIE-WINDOW-01",
      12,
      "split-pane-toggle",
    ],
    [
      "XLSCTL.VIEW.WINDOW.FULLSCREEN",
      "fullscreen",
      "XLS-VIE-WINDOW-01",
      13,
      "fullscreen",
    ],
  ].map(([controlId, suffix, canonicalId, order, predicate]) =>
    define({
      controlId: String(controlId),
      featureId: `spreadsheet.view.${String(suffix)}`,
      canonicalId: String(canonicalId),
      group: "View",
      order: Number(order),
      controlKind:
        suffix === "commitZoom"
          ? "combobox"
          : suffix === "newWindow"
            ? "button"
            : "toggle",
      predicate: predicate as SpreadsheetVisiblePredicate,
      reasonKey:
        predicate === "saved-window"
          ? "spreadsheet.capability.savedWorkbookRequired"
          : predicate === "fullscreen"
            ? "spreadsheet.capability.fullscreenUnavailable"
            : null,
      command: "bound",
      undoPolicy: String(suffix).startsWith("freeze")
        ? "single-transaction"
        : "content-none",
      ioProfile: String(suffix).startsWith("freeze")
        ? "worksheet-freeze"
        : String(suffix) === "splitPanes"
          ? "worksheet-split"
        : String(suffix).startsWith("toggle")
          ? "workbook-view-preference"
          : "session-only",
    }),
  ),
  ...[
    ["XLSCTL.VIEW.SHEET_VIEW.SWITCH", "defaultSheetView", 1, "sheet-view-switch"],
    ["XLSCTL.VIEW.SHEET_VIEW.EXIT", "exitSheetView", 12, "sheet-view-exit"],
    ["XLSCTL.VIEW.SHEET_VIEW.KEEP", "keepSheetView", 13, "sheet-view-keep"],
    ["XLSCTL.VIEW.SHEET_VIEW.NEW", "newSheetView", 14, "sheet-view-new"],
    ["XLSCTL.VIEW.SHEET_VIEW.OPTIONS", "sheetViewOptions", 15, "sheet-view-options"],
  ].map(([controlId, suffix, order, predicate]) =>
    define({
      controlId: String(controlId),
      featureId: `spreadsheet.view.${String(suffix)}`,
      canonicalId: "XLS-VIE-MODE-01",
      group: "View",
      order: Number(order),
      controlKind: String(suffix) === "defaultSheetView" ? "menu" : "button",
      predicate: predicate as SpreadsheetVisiblePredicate,
      reasonKey: "spreadsheet.capability.namedViewsUnavailable",
      command: "bound",
      undoPolicy: "content-none",
      ioProfile: "worksheet-sheet-views",
    }),
  ),
  define({
    controlId: "XLSCTL.VIEW.SHOW.NAVIGATION",
    featureId: "spreadsheet.view.navigation",
    canonicalId: "XLS-VIE-SHOW-01",
    group: "View",
    order: 16,
    controlKind: "button",
    predicate: "always",
    reasonKey: null,
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
  define({
    controlId: "XLSCTL.VIEW.SHOW.DATA_TYPE_ICONS",
    featureId: "spreadsheet.view.dataTypeIcons",
    canonicalId: "XLS-VIE-SHOW-01",
    group: "View",
    order: 17,
    controlKind: "button",
    predicate: "no-verified-authority",
    reasonKey: "spreadsheet.capability.dataTypesUnavailable",
    command: "unsupported",
    undoPolicy: "content-none",
    ioProfile: "unsupported",
  }),
  define({
    controlId: "XLSCTL.VIEW.IMMERSIVE_READER",
    featureId: "spreadsheet.view.immersiveReader",
    canonicalId: "XLS-VIE-SHOW-01",
    group: "View",
    order: 18,
    controlKind: "button",
    predicate: "selection-display-text",
    reasonKey: "spreadsheet.capability.immersiveReaderUnavailable",
    command: "bound",
    undoPolicy: "content-none",
    ioProfile: "session-only",
  }),
] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

/** Executable registry = immutable Web base inventory + contextual slices. */
export const SPREADSHEET_COPILOT_APPLY_CONTROL_ID = "XLSCTL.COPILOT.MRAI.APPLY";
const SPREADSHEET_COPILOT_CONTROL = define({
  controlId: SPREADSHEET_COPILOT_APPLY_CONTROL_ID,
  featureId: "spreadsheet.copilot.apply",
  canonicalId: "XLS-MRAI-COPILOT-APPLY-01",
  group: "Home", order: 0, controlKind: "button", predicate: "copilot-apply",
  reasonKey: "spreadsheet.capability.commandUnavailable", command: "bound",
  undoPolicy: "single-transaction", ioProfile: "worksheet-copilot",
});
export const SPREADSHEET_VISIBLE_COMMAND_DEFINITIONS = Object.freeze([
  SPREADSHEET_COPILOT_CONTROL,
  ...SPREADSHEET_VISIBLE_CONTROL_DEFINITIONS,
  ...SPREADSHEET_TABLE_DESIGN_STYLE_VERTICAL_CONTROL_DEFINITIONS,
  ...SPREADSHEET_TABLE_STYLE_OPTIONS_VERTICAL_CONTROL_DEFINITIONS,
  ...SPREADSHEET_TABLE_NAME_VERTICAL_CONTROL_DEFINITIONS,
  ...SPREADSHEET_TABLE_RESIZE_VERTICAL_CONTROL_DEFINITIONS,
  ...SPREADSHEET_TABLE_DESIGN_OFFICE_PARITY_CONTROL_DEFINITIONS,
  ...SPREADSHEET_SHAPE_FORMAT_NATIVE_CORE_CONTROL_DEFINITIONS,
  ...SPREADSHEET_PICTURE_FORMAT_TRANSFORM_CONTROL_DEFINITIONS,
  ...SPREADSHEET_PICTURE_FORMAT_EFFECT_CONTROL_DEFINITIONS,
  ...SPREADSHEET_PICTURE_FORMAT_ARTISTIC_EFFECT_CONTROL_DEFINITIONS,
  ...SPREADSHEET_PICTURE_FORMAT_METADATA_CONTROL_DEFINITIONS,
  ...SPREADSHEET_SLICER_DESIGN_FORMAT_CONTROL_DEFINITIONS,
  ...SPREADSHEET_SPARKLINE_DESIGN_NATIVE_CORE_CONTROL_DEFINITIONS,
  ...SPREADSHEET_CHART_DESIGN_NATIVE_CORE_CONTROL_DEFINITIONS,
  ...SPREADSHEET_PIVOT_ANALYZE_VERTICAL_CONTROL_DEFINITIONS,
  ...SPREADSHEET_PIVOT_CHART_ANALYZE_CONTROL_DEFINITIONS,
] as const satisfies readonly SpreadsheetVisibleControlDefinition[]);

export type SpreadsheetVisibleControlId =
  (typeof SPREADSHEET_VISIBLE_COMMAND_DEFINITIONS)[number]["controlId"];

// The Excel-for-the-Web atlas coverage truth (canonical parents, official web
// evidence, rendered/boundary control partitions) is audit data the ribbon
// never reads at runtime: it lives in `./visible-command-coverage`, outside the
// browser bundle (`engine-host/server-only-modules.json` -> auditOnly).


const definitionByControl = new Map(
  SPREADSHEET_VISIBLE_COMMAND_DEFINITIONS.map((definition) => [
    definition.controlId,
    definition,
  ]),
);

function commandKey(featureId: string): string {
  return featureId.replace("spreadsheet.", "spreadsheet.command.");
}

function handlerRef(featureId: string): string {
  return `SpreadsheetWindow#${featureId.slice("spreadsheet.".length)}`;
}

type SpreadsheetVisibleIoContract = Pick<
  SpreadsheetFeatureCapabilityContract,
  | "domain"
  | "import"
  | "fromScratchExport"
  | "patchbackExport"
  | "officeOutcome"
>;

/**
 * Writer truth for every verified visible command. A bound mutator may execute
 * only when its current native-write route resolves to `native` or to an
 * explicit declared-loss writer. `reject` is a real fail-closed boundary,
 * never an optimistic placeholder.
 */
const SPREADSHEET_VISIBLE_IO_CONTRACTS = {
  "worksheet-copilot": {
    domain: { disposition: "typed", schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.cells", "native-page-schemas.ts#SpreadsheetSheet.objects", "native-page-schemas.ts#SpreadsheetSheet.charts", "native-page-schemas.ts#SpreadsheetSheet.dataTables"], version: "1", persistence: "workbook" },
    import: { disposition: "native", projectionRefs: ["editable_page.py#_xlsx_content"], preservedUnknownParts: true, lossCodes: [], flags: ["per-target-native-authority-required"] },
    fromScratchExport: { disposition: "native", serializerRefs: ["notebooks.py#_spreadsheet_to_xlsx"], lossCodes: [], flags: ["per-target-native-authority-required"] },
    patchbackExport: { disposition: "native", serializerRefs: ["xlsx/patchback.py#patchback_xlsx"], preservePolicy: "owned-parts-only", lossCodes: [], flags: ["per-target-native-authority-required"] },
    officeOutcome: { excel: "unknown", wps: "unknown" },
  },
  "worksheet-queries": {
    domain: { disposition: "typed", schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.dataQueries"], version: "1", persistence: "worksheet" },
    import: { disposition: "native", projectionRefs: ["editable_page.py#_xlsx_content", "xlsx/data_queries.py#project_native_queries"], preservedUnknownParts: true, lossCodes: [], flags: ["source-custody-only", "closed-query-grammar", "unknown-queries-preserve-only"] },
    fromScratchExport: { disposition: "native", serializerRefs: ["notebooks.py#_spreadsheet_to_xlsx", "xlsx/data_queries.py#attach_native_queries"], lossCodes: [], flags: ["closed-query-grammar", "external-interop-unverified"] },
    patchbackExport: { disposition: "native", serializerRefs: ["xlsx/patchback.py#patchback_xlsx", "xlsx/data_queries.py#attach_native_queries"], preservePolicy: "owned-parts-only", lossCodes: [], flags: ["closed-query-grammar", "external-interop-unverified"] },
    officeOutcome: { excel: "unknown", wps: "unknown" },
  },
  "session-only": {
    domain: {
      disposition: "view-only",
      schemaPaths: [],
      version: null,
      persistence: "session",
    },
    import: {
      disposition: "not-applicable",
      projectionRefs: [],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "not-applicable",
      serializerRefs: [],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "not-applicable",
      serializerRefs: [],
      preservePolicy: null,
      lossCodes: [],
      flags: [],
    },
    officeOutcome: { excel: "visible", wps: "visible" },
  },
  unsupported: {
    domain: {
      disposition: "none",
      schemaPaths: [],
      version: null,
      persistence: "none",
    },
    import: {
      disposition: "unsupported",
      projectionRefs: [],
      preservedUnknownParts: true,
      lossCodes: ["XLS_VISIBLE_COMMAND_UNSUPPORTED"],
      flags: ["fail-closed"],
    },
    fromScratchExport: {
      disposition: "reject",
      serializerRefs: [],
      lossCodes: ["XLS_VISIBLE_COMMAND_UNSUPPORTED"],
      flags: ["fail-closed"],
    },
    patchbackExport: {
      disposition: "reject",
      serializerRefs: [],
      preservePolicy: null,
      lossCodes: ["XLS_VISIBLE_COMMAND_UNSUPPORTED"],
      flags: ["fail-closed"],
    },
    officeOutcome: { excel: "unsupported", wps: "unsupported" },
  },
  "workbook-import": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetContent"],
      version: "1",
      persistence: "workbook",
    },
    import: {
      disposition: "native",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: ["explicit-loss-confirmation"],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#_spreadsheet_to_xlsx"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "reject",
      serializerRefs: [],
      preservePolicy: null,
      lossCodes: ["XLS_PATCHBACK_WORKBOOK_REPLACEMENT_UNSUPPORTED"],
      flags: ["fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "workbook-defined-names": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.namedRanges"],
      version: "1",
      persistence: "workbook",
    },
    import: {
      disposition: "native",
      projectionRefs: [
        "editable_page.py#_xlsx_defined_names/_xlsx_sheet_defined_names",
        "editable_page.py#_xlsx_content",
      ],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: ["per-name-preserve-only"],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#_spreadsheet_to_xlsx"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_defined_name_patches",
        "xlsx/patchback.py#_apply_defined_name_patch",
      ],
      preservePolicy:
        "Hidden, built-in, external and extension-bearing defined names remain byte-owned by the immutable source package; only source-authenticated classic user names are patched.",
      lossCodes: [],
      flags: ["per-name-preserve-only", "source-identity-cas"],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "workbook-structure-protection": {
    domain: {
      disposition: "typed",
      schemaPaths: [
        "native-page-schemas.ts#SpreadsheetContent.workbookStructureProtection",
        "protection.ts#WorkbookStructureProtection",
      ],
      version: "1",
      persistence: "workbook",
    },
    import: {
      disposition: "native",
      projectionRefs: [
        "editable_page.py#_xlsx_workbook_structure_protection",
      ],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: ["unknown-hash-preserve-only"],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#_spreadsheet_to_xlsx"],
      lossCodes: [],
      flags: ["password-plaintext-never-persisted"],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_workbook_structure_protection_patch",
        "xlsx/patchback.py#_apply_workbook_structure_protection_patch",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: ["unknown-hash-preserve-only", "fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "workbook-calculation-mode": {
    domain: {
      disposition: "typed",
      schemaPaths: [
        "native-page-schemas.ts#SpreadsheetContent.calculationMode",
        "calculation-mode.ts#CalculationMode",
      ],
      version: "1",
      persistence: "workbook",
    },
    import: {
      disposition: "native",
      projectionRefs: ["editable_page.py#workbook.calculation.calcMode"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#_spreadsheet_to_xlsx/calculationMode"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_calculation_mode_patch",
        "xlsx/patchback.py#_apply_calculation_mode_patch",
      ],
      preservePolicy: "copy-original-package-and-patch-calcPr",
      lossCodes: [],
      flags: [],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "workbook-date-system": {
    domain: {
      disposition: "typed",
      schemaPaths: [
        "native-page-schemas.ts#SpreadsheetContent.dateSystem",
        "excel-1900-calendar.ts#SpreadsheetDateSystem",
      ],
      version: "1",
      persistence: "workbook",
    },
    import: {
      disposition: "native",
      projectionRefs: [
        "editable_page.py#_xlsx_content",
        "editable_page.py#workbook.epoch",
      ],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#_spreadsheet_to_xlsx/dateSystem"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_date_system_patch",
        "xlsx/patchback.py#_apply_date_system_patch",
      ],
      preservePolicy: "copy-original-package-and-patch-workbookPr",
      lossCodes: [],
      flags: [],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "workbook-view-preference": {
    domain: {
      disposition: "typed",
      schemaPaths: ["view-options.ts#SpreadsheetViewOptions"],
      version: "1",
      persistence: "workbook",
    },
    import: {
      disposition: "not-applicable",
      projectionRefs: [],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "not-applicable",
      serializerRefs: [],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "not-applicable",
      serializerRefs: [],
      preservePolicy: null,
      lossCodes: [],
      flags: [],
    },
    officeOutcome: { excel: "visible", wps: "visible" },
  },
  "worksheet-advanced-filter": {
    domain: {
      disposition: "typed",
      schemaPaths: [
        "native-page-schemas.ts#SpreadsheetSheet.cells",
        "native-page-schemas.ts#SpreadsheetSheet.advancedFilterState",
      ],
      version: "2",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#_spreadsheet_to_xlsx"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_advanced_filter_state_patches",
        "xlsx/patchback.py#patchback_xlsx",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: [],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-cells": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.cells"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#_spreadsheet_to_xlsx"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_cell_patches",
        "xlsx/patchback.py#patchback_xlsx",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: [],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  // A cell control the editor adds (Insert > Checkbox; More > Option button,
  // Spin button, Scroll bar). A new workbook writes the check box as Excel's
  // checkbox cell format (`cell_controls.inject_checkbox_controls`); the
  // in-place patch of an imported workbook has no writer for a new cell
  // control and declines the sheet's whole control facet
  // (`_xlsx_blocked_cell_control_edit_error`), so on that route the command
  // is unavailable instead of dropped after the save (import audit S1/S2).
  "worksheet-cell-controls": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.cells.style"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: ["xlsx/cell_controls.py#checkbox_style_indexes"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["xlsx/cell_controls.py#inject_checkbox_controls"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "reject",
      serializerRefs: [],
      preservePolicy: null,
      lossCodes: ["XLS_PATCHBACK_NEW_CELL_CONTROL_UNSUPPORTED"],
      flags: ["fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "visible" },
  },
  // Excel's Place in Cell: a picture that is a cell's value, read from and
  // written to the workbook's rich value parts (xl/metadata.xml valueMetadata
  // and futureMetadata, xl/richData/*, the cell's vm attribute). An imported
  // workbook's own package is appended to (its rich values are never
  // renumbered), so a new, moved, retitled or removed picture saves in place.
  "worksheet-cell-pictures": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.cells.picture"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: [
        "editable_page.py#_xlsx_content",
        "xlsx/cell_pictures.py#read_cell_pictures",
      ],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: [
        "notebooks.py#_spreadsheet_to_xlsx",
        "xlsx/cell_pictures.py#inject_cell_pictures_and_backgrounds",
      ],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_cell_picture_sheets",
        "xlsx/cell_pictures.py#patch_cell_pictures_and_backgrounds",
      ],
      preservePolicy:
        "Rich values, structures and relationships already in the package keep their positions; new ones are appended, and a picture the projection could not read stays in its cell or follows it.",
      lossCodes: [],
      flags: ["fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "unknown" },
  },
  // Excel desktop's Page Layout > Background: the worksheet's <picture r:id>
  // image part, shown tiled behind the cells and never printed.
  "worksheet-background-picture": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.backgroundPicture"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: [
        "editable_page.py#_xlsx_content",
        "xlsx/cell_pictures.py#read_sheet_backgrounds",
      ],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: [
        "notebooks.py#_spreadsheet_to_xlsx",
        "xlsx/cell_pictures.py#inject_cell_pictures_and_backgrounds",
      ],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_cell_picture_sheets",
        "xlsx/cell_pictures.py#patch_cell_pictures_and_backgrounds",
      ],
      preservePolicy:
        "Only the worksheet's <picture r:id>, its relationship and an image part nothing else names are replaced or removed.",
      lossCodes: [],
      flags: ["fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "unknown" },
  },
  "worksheet-paste-special": {
    domain: {
      disposition: "typed",
      schemaPaths: [
        "native-page-schemas.ts#SpreadsheetSheet.cells",
        "native-page-schemas.ts#SpreadsheetSheet.validations",
        "native-page-schemas.ts#SpreadsheetSheet.columnWidths",
      ],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: [
        "notebooks.py#_spreadsheet_to_xlsx",
        "notebooks.py#apply_data_validations",
        "notebooks.py#apply_sheet_chrome",
      ],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_cell_patches",
        "editable_page_patchback.py#_xlsx_cell_comment_patch",
        "editable_page_patchback.py#_xlsx_data_validation_patches",
        "editable_page_patchback.py#_xlsx_sheet_dimension_patches",
        "xlsx/patchback.py#patchback_xlsx",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: ["validation-patchback-fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-data-form-delete": {
    domain: {
      disposition: "typed",
      schemaPaths: [
        "native-page-schemas.ts#SpreadsheetSheet.cells",
        "native-page-schemas.ts#SpreadsheetSheet.dataTables",
      ],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: [
        "notebooks.py#_spreadsheet_to_xlsx",
        "notebooks.py#add_sheet_table",
      ],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      // The moved records are cell moves: each cell carries its style,
      // hyperlink and note (comment + VML box) to its new address.
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_cell_patches",
        "editable_page_patchback.py#_xlsx_structured_table_patches",
        "xlsx/patchback.py#_relocate_moved_cell_notes",
        "xlsx/patchback.py#patchback_xlsx",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: [],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-comments": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetCell.comment"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#_spreadsheet_to_xlsx"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_cell_comment_patch",
        "xlsx/patchback.py#patchback_xlsx",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: [],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-conditional-formats": {
    domain: {
      disposition: "typed",
      schemaPaths: [
        "native-page-schemas.ts#SpreadsheetSheet.conditionalFormats",
      ],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#apply_conditional_formats"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_conditional_format_patches",
        "xlsx/patchback.py#XlsxConditionalFormatCollectionPatch",
        "xlsx/patchback.py#patchback_xlsx",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: ["fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-drawing-objects": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.objects"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "collapsed",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: ["spreadsheet_objects_omitted"],
      flags: ["explicit-loss-confirmation"],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: [
        "notebooks.py#_spreadsheet_to_xlsx",
        "object_export.py#inject_spreadsheet_shape_objects",
      ],
      lossCodes: [],
      flags: [],
    },
    // An imported workbook takes a NEW picture (an icon is an SVG picture
    // with its PNG), shape, text box, WordArt and SmartArt in place: each is
    // written whole on every save into the original package, a drawing part
    // made when the sheet has none (`XlsxInsertedObjectPatch`,
    // `XlsxInsertedSmartArtPatch`); an imported object's own edits are the
    // anchor/format writer's.
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_inserted_object_patches",
        "xlsx/patchback.py#XlsxInsertedObjectPatch",
        "xlsx/object_export.py#append_spreadsheet_pictures",
        "xlsx/object_export.py#inject_spreadsheet_shape_objects",
        "xlsx/patchback.py#XlsxInsertedSmartArtPatch",
        "xlsx/patchback.py#XlsxObjectAnchorPatch",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: ["fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-shape-format": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.objects"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: [
        "editable_page.py#_xlsx_sheet_objects_by_index",
        "editable_page.py#_xlsx_shape_object",
        "editable_page.py#_xlsx_shape_transform",
      ],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: ["stable-nonvisual-shape-identity"],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["object_export.py#inject_spreadsheet_shape_objects"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_object_anchor_patches",
        "xlsx/patchback.py#XlsxObjectAnchorPatch",
        "xlsx/patchback.py#patchback_xlsx",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: ["stable-nonvisual-shape-identity", "fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-picture-transforms": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.objects"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: ["editable_page.py#_xlsx_sheet_objects_by_index"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: ["stable-drawing-identity"],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: [
        "notebooks.py#_spreadsheet_to_xlsx",
        "object_export.py#inject_spreadsheet_shape_objects",
      ],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "xlsx/patchback.py#XlsxObjectAnchorPatch",
        "xlsx/patchback.py#XlsxDrawingOrderPatch",
        "xlsx/patchback.py#XlsxDeleteObjectPatch",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: ["stable-drawing-identity", "fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-picture-metadata": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.objects"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: ["editable_page.py#_xlsx_sheet_objects_by_index"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: ["stable-media-identity"],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#_spreadsheet_to_xlsx"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "xlsx/patchback.py#XlsxImagePatch",
        "xlsx/patchback.py#XlsxImageMetadataPatch",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: ["stable-media-identity", "fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-chart-metadata": {
    domain: {
      disposition: "typed",
      schemaPaths: [
        "native-page-schemas.ts#SpreadsheetSheet.charts",
        "native-page-schemas.ts#SpreadsheetSheet.objects",
      ],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "collapsed",
      projectionRefs: ["editable_page.py#_xlsx_chart_metadata"],
      preservedUnknownParts: true,
      lossCodes: ["xlsx_chart_not_projected"],
      flags: ["explicit-loss-confirmation"],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#_spreadsheet_to_xlsx/charts"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_chart_patches",
        "xlsx/patchback.py#XlsxChartPatch",
        "xlsx/patchback.py#patchback_xlsx",
      ],
      preservePolicy:
        "authenticated classic chart metadata is patched surgically; chartEx and unsupported chart families remain preserve-only",
      lossCodes: [],
      flags: ["native-identity-required", "fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-charts": {
    domain: {
      disposition: "typed",
      schemaPaths: [
        "native-page-schemas.ts#SpreadsheetSheet.charts",
        "native-page-schemas.ts#SpreadsheetSheet.objects",
      ],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "collapsed",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: ["xlsx_chart_not_projected"],
      flags: ["explicit-loss-confirmation"],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#_spreadsheet_to_xlsx/charts"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: ["xlsx/chart_append.py#reconcile_native_chart_recipes"],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: ["verified-source-recipe-bindings", "fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-pivots": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.pivotTables"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "collapsed",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: ["xlsx_pivot_table_not_projected"],
      flags: ["explicit-loss-confirmation"],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#_inject_xlsx_pivot_tables"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_native_pivot_patches",
        "xlsx/patchback.py#_plan_pivot_table_patch",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: ["fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-sparklines": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.sparklines"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "declared-loss",
      serializerRefs: ["notebooks.py#_inject_xlsx_sparklines"],
      lossCodes: ["XLS_SPARKLINE_INJECTION_PRESERVED_ON_FAILURE"],
      flags: ["loss-manifest-required"],
    },
    // An imported workbook's x14 sparkline groups are rewritten in place;
    // a graph the writer cannot own stays byte-for-byte
    // (`nativeSparklinesBlocker`) and Sparkline Design stays closed on it.
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_sparkline_patches",
        "xlsx/patchback.py#XlsxSparklineCollectionPatch",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: ["fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-slicers": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.slicers"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: ["table_slicer.py#read_table_slicers"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: ["stable-table-field-identity"],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["table_slicer.py#inject_table_slicers"],
      lossCodes: [],
      flags: ["fail-closed"],
    },
    // Import usability audit S6: an imported workbook's save reconciles the
    // table and PivotTable slicers the editor added, changed or removed
    // after every other patch (sync_editable_page_after_save).
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "table_slicer.py#reconcile_table_slicers",
        "pivot_slicer.py#reconcile_pivot_slicers",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: ["fail-closed", "preserve-native-parts"],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-filters": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.filters"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#apply_auto_filter"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_auto_filter_patches",
        "xlsx/patchback.py#patchback_xlsx",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: [],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-freeze": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.freezeRows"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#apply_freeze_panes"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_freeze_patches",
        "xlsx/patchback.py#patchback_xlsx",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: [],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-split": {
    domain: {
      disposition: "typed",
      schemaPaths: [
        "native-page-schemas.ts#SpreadsheetSheet.splitRow",
        "native-page-schemas.ts#SpreadsheetSheet.splitCol",
        "native-page-schemas.ts#SpreadsheetSheet.splitPaneState",
      ],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#apply_split_panes"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_split_patches",
        "xlsx/patchback.py#patchback_xlsx",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: [],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-outline": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.outlineGroups"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "collapsed",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: ["xlsx_outline_group_not_projected"],
      flags: ["explicit-loss-confirmation"],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#apply_sheet_outline_groups"],
      lossCodes: [],
      flags: [],
    },
    // The writer replaces a sheet's outline only when the file still holds
    // the outline the page started from (fail-closed); a sheet whose outline
    // the import could not read exactly carries outlineGroupsNativeBlocker
    // and keeps outline edits off.
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_outline_group_patches",
        "xlsx/patchback.py#XlsxOutlineGroupsPatch",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: ["fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-print-settings": {
    domain: {
      disposition: "typed",
      schemaPaths: ["print-settings.ts#PrintSettings/normalizePrintSettings"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "collapsed",
      projectionRefs: [
        "editable_page.py#_xlsx_content",
        "page_projection.py#_worksheet_print_settings",
        "xlsx/header_footer_pictures.py#project_header_footer_pictures",
      ],
      preservedUnknownParts: true,
      lossCodes: [
        "xlsx_print_area_not_projected",
        "xlsx_print_titles_not_projected",
        "xlsx_header_footer_code_not_projected",
        "xlsx_header_footer_page_offset_not_projected",
        "xlsx_header_footer_picture_not_projected",
        "xlsx_header_footer_truncated",
        "xlsx_manual_page_breaks_not_projected",
        "xlsx_page_margins_not_projected",
        "xlsx_page_setup_option_not_projected",
        "xlsx_paper_size_not_projected",
        "xlsx_print_options_not_projected",
        "xlsx_print_orientation_not_projected",
        "xlsx_print_scaling_not_projected",
        "xlsx_printer_resolution_not_projected",
        "xlsx_printer_settings_not_projected",
        "xlsx_theme_color_unresolved",
      ],
      flags: ["explicit-loss-confirmation"],
    },
    fromScratchExport: {
      disposition: "declared-loss",
      serializerRefs: [
        "notebooks.py#apply_print_settings",
        "xlsx/header_footer_pictures.py#inject_header_footer_pictures",
      ],
      lossCodes: ["XLS_PRINT_HEADER_FOOTER_RICH_RUNS_FLATTENED"],
      flags: ["rich-header-footer-edit-fail-closed"],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_print_settings_patches",
        "xlsx/patchback.py#_apply_print_settings_to_worksheet",
        "xlsx/header_footer_pictures.py#patch_header_footer_pictures",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: ["rich-header-footer-edit-fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-protection": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.protected"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#apply_sheet_chrome"],
      lossCodes: [],
      flags: [],
    },
    
    // Protect/Unprotect Sheet and a cell's Locked flag (N06/N07 tests in
    // test_xlsx_native_cell_edit_fidelity.py); this row said it did not.
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_sheet_protection_patches",
        "xlsx/patchback.py#XlsxSheetProtectionPatch",
        "editable_page_patchback.py#_xlsx_cell_patches",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: ["fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-structure": {
    domain: {
      disposition: "typed",
      schemaPaths: [
        "native-page-schemas.ts#SpreadsheetSheet.cells",
        "native-page-schemas.ts#SpreadsheetSheet.merges",
      ],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#_spreadsheet_to_xlsx"],
      lossCodes: [],
      flags: [],
    },
    // Import usability audit S3/S4: merges (`_xlsx_merge_patches`) and
    // inserted/deleted cells, rows and columns (`_xlsx_moved_cell_sources`
    // with their structural followers) are written in place -- the row
    // header's Insert, gated elsewhere, saved them all along.
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_merge_patches",
        "editable_page_patchback.py#_xlsx_moved_cell_sources",
        "xlsx/patchback.py#patchback_xlsx",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: ["fail-closed"],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-sheet-views": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.namedSheetViews"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: ["xlsx_named_sheet_view_not_projected"],
      flags: ["explicit-loss-confirmation"],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#_inject_xlsx_named_sheet_views"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_named_sheet_view_patches",
        "xlsx/patchback.py#patchback_xlsx",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: [],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-table": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.dataTables"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "native",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: [],
      flags: [],
    },
    fromScratchExport: {
      disposition: "native",
      serializerRefs: ["notebooks.py#add_sheet_table"],
      lossCodes: [],
      flags: [],
    },
    patchbackExport: {
      disposition: "native",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_structured_table_patches",
        "xlsx/patchback.py#patchback_xlsx",
      ],
      preservePolicy: "owned-parts-only",
      lossCodes: [],
      flags: [],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
  "worksheet-validation": {
    domain: {
      disposition: "typed",
      schemaPaths: ["native-page-schemas.ts#SpreadsheetSheet.validations"],
      version: "1",
      persistence: "worksheet",
    },
    import: {
      disposition: "collapsed",
      projectionRefs: ["editable_page.py#_xlsx_content"],
      preservedUnknownParts: true,
      lossCodes: ["xlsx_data_validation_rule_not_projected"],
      flags: ["explicit-loss-confirmation"],
    },
    fromScratchExport: {
      disposition: "declared-loss",
      serializerRefs: ["notebooks.py#apply_data_validations"],
      lossCodes: ["XLS_VALIDATION_DEPENDENT_WEB_ONLY"],
      flags: ["loss-manifest-required"],
    },
    patchbackExport: {
      disposition: "declared-loss",
      serializerRefs: [
        "editable_page_patchback.py#_xlsx_data_validation_patches",
        "file_adapters/xlsx/patchback.py#XlsxDataValidationPatch",
        "file_adapters/xlsx/patchback.py#_apply_data_validation_patch_to_worksheet",
      ],
      preservePolicy:
        "supported classic single-sqref list/number/text/date/time/custom subtypes apply; dependent, unknown, duplicate, invalid or over-budget list edits remain source-preserved and are not applied; unchanged receipt XML is restored separately; the low-level writer refuses overlap, multi-area and x14 graphs",
      lossCodes: ["data_validation_not_patchable"],
      flags: [
        "native-subtype-guard",
        "collector-preserve-only-unsupported-subtype",
        "low-level-fail-closed",
      ],
    },
    officeOutcome: { excel: "editable", wps: "editable" },
  },
} as const satisfies Record<
  SpreadsheetVisibleIoProfile,
  SpreadsheetVisibleIoContract
>;

/**
 * Whether the save route can write a new cell control (check box, option
 * button, spin button, scroll bar): the one decision the ribbon's Checkbox
 * (through its catalog row) and the More menu's form controls share.
 */
export function spreadsheetCellControlWritable(route: SpreadsheetNativeWriteRoute): boolean {
  return spreadsheetVisibleIoWriterReady(SPREADSHEET_VISIBLE_IO_CONTRACTS["worksheet-cell-controls"], route);
}

function spreadsheetVisibleIoWriterReady(
  contract: Pick<
    SpreadsheetFeatureCapabilityContract,
    "fromScratchExport" | "patchbackExport"
  >,
  route: SpreadsheetNativeWriteRoute,
): boolean {
  const writer =
    route === "from-scratch"
      ? contract.fromScratchExport
      : contract.patchbackExport;
  return (
    writer.disposition === "native" ||
    (writer.disposition === "declared-loss" &&
      writer.serializerRefs.length > 0 &&
      writer.lossCodes.length > 0 &&
      (route !== "patchback" ||
        contract.patchbackExport.preservePolicy !== null))
  );
}

/**
 * Exact source-generation receipt used by every persistent Page Setup command.
 * This pure guard is exported so the focused authority gate can prove the nine
 * writer IDs share one fail-closed contract without duplicating UI journeys.
 */
export function spreadsheetPageSetupWriterReceiptMatches(input: {
  readonly expectedControlId: string;
  readonly receiptControlId: string | null | undefined;
  readonly sourceWorkbook: unknown;
  readonly sourceActiveSheetIndex: number;
  readonly sourceSheet: unknown;
  readonly sourceSheetId: string;
  readonly liveWorkbook: unknown;
  readonly liveActiveSheetIndex: number;
  readonly liveSheet: unknown;
  readonly liveSheetId: string;
  readonly receiptRow?: number;
  readonly receiptCol?: number;
  readonly liveRow?: number;
  readonly liveCol?: number;
  readonly receiptSelectionEpoch?: number;
  readonly liveSelectionEpoch?: number;
  readonly receiptRange?: string;
  readonly liveRange?: string | null;
}): boolean {
  if (
    input.receiptControlId !== input.expectedControlId ||
    input.sourceWorkbook !== input.liveWorkbook ||
    input.sourceActiveSheetIndex !== input.liveActiveSheetIndex ||
    input.sourceSheet !== input.liveSheet ||
    input.sourceSheetId !== input.liveSheetId
  ) {
    return false;
  }
  const hasReceiptCoordinate =
    input.receiptRow !== undefined || input.receiptCol !== undefined;
  if (
    hasReceiptCoordinate &&
    (input.receiptRow !== input.liveRow ||
      input.receiptCol !== input.liveCol)
  ) {
    return false;
  }
  const hasSelectionEpoch =
    input.receiptSelectionEpoch !== undefined ||
    input.liveSelectionEpoch !== undefined;
  if (
    hasSelectionEpoch &&
    input.receiptSelectionEpoch !== input.liveSelectionEpoch
  ) {
    return false;
  }
  const hasRange =
    input.receiptRange !== undefined || input.liveRange !== undefined;
  return !hasRange || input.receiptRange === input.liveRange;
}

/** Exact source/selection receipt shared by Data Validation Apply and Clear. */
export function spreadsheetDataValidationWriterReceiptMatches(input: {
  readonly expectedControlId: string;
  readonly receiptControlId: string | null | undefined;
  readonly sourceWorkbook: unknown;
  readonly sourceActiveSheetIndex: number;
  readonly sourceSheet: unknown;
  readonly sourceSheetId: string;
  readonly sourceSelectionEpoch: number;
  readonly sourceRange: Readonly<{
    r1: number;
    r2: number;
    c1: number;
    c2: number;
  }>;
  readonly liveWorkbook: unknown;
  readonly liveActiveSheetIndex: number;
  readonly liveSheet: unknown;
  readonly liveSheetId: string;
  readonly liveSelectionEpoch: number;
  readonly liveRange: Readonly<{
    r1: number;
    r2: number;
    c1: number;
    c2: number;
  }>;
}): boolean {
  return spreadsheetInsertWriterReceiptMatches(input);
}

/**
 * Exact workbook/sheet/selection/rule-generation guard for Manage Rules.
 * `parameter` is a canonical action payload (rule id, target range, direction,
 * or Stop If True value); a retained DOM callback cannot substitute another
 * rule or reuse authority after any live generation changes.
 */
export function spreadsheetConditionalFormatWriterReceiptMatches(input: {
  readonly expectedControlId: string;
  readonly receiptControlId: string | null | undefined;
  readonly sourceWorkbook: unknown;
  readonly sourceActiveSheetIndex: number;
  readonly sourceSheet: unknown;
  readonly sourceSheetId: string;
  readonly sourceRuleCollection: unknown;
  readonly sourceSelectionEpoch: number;
  readonly sourceRange: Readonly<{
    r1: number;
    r2: number;
    c1: number;
    c2: number;
  }>;
  readonly sourceAction: string;
  readonly sourceParameter: string;
  readonly liveWorkbook: unknown;
  readonly liveActiveSheetIndex: number;
  readonly liveSheet: unknown;
  readonly liveSheetId: string;
  readonly liveRuleCollection: unknown;
  readonly liveSelectionEpoch: number;
  readonly liveRange: Readonly<{
    r1: number;
    r2: number;
    c1: number;
    c2: number;
  }>;
  readonly liveAction: string;
  readonly liveParameter: string;
}): boolean {
  return (
    input.receiptControlId === input.expectedControlId &&
    input.sourceWorkbook === input.liveWorkbook &&
    input.sourceActiveSheetIndex === input.liveActiveSheetIndex &&
    input.sourceSheet === input.liveSheet &&
    input.sourceSheetId === input.liveSheetId &&
    input.sourceRuleCollection === input.liveRuleCollection &&
    input.sourceSelectionEpoch === input.liveSelectionEpoch &&
    input.sourceRange.r1 === input.liveRange.r1 &&
    input.sourceRange.r2 === input.liveRange.r2 &&
    input.sourceRange.c1 === input.liveRange.c1 &&
    input.sourceRange.c2 === input.liveRange.c2 &&
    input.sourceAction === input.liveAction &&
    input.sourceParameter === input.liveParameter
  );
}

/**
 * Exact generation/parameter guard for the Table Design style gallery. A
 * retained tile may not mutate a replaced workbook, sheet, table collection,
 * table entity or selection, even when the public ids still compare equal.
 */
export function spreadsheetTableStyleWriterReceiptMatches(input: {
  readonly expectedControlId: string;
  readonly receiptControlId: string | null | undefined;
  readonly sourceWorkbook: unknown;
  readonly sourceActiveSheetIndex: number;
  readonly sourceSheet: unknown;
  readonly sourceSheetId: string;
  readonly sourceTableCollection: unknown;
  readonly sourceTable: unknown;
  readonly sourceTableId: string;
  readonly sourceSelectionEpoch: number;
  readonly sourceRange: Readonly<{
    r1: number;
    r2: number;
    c1: number;
    c2: number;
  }>;
  readonly sourceStyleId: string;
  readonly liveWorkbook: unknown;
  readonly liveActiveSheetIndex: number;
  readonly liveSheet: unknown;
  readonly liveSheetId: string;
  readonly liveTableCollection: unknown;
  readonly liveTable: unknown;
  readonly liveTableId: string;
  readonly liveSelectionEpoch: number;
  readonly liveRange: Readonly<{
    r1: number;
    r2: number;
    c1: number;
    c2: number;
  }>;
  readonly liveStyleId: string;
}): boolean {
  return (
    input.receiptControlId === input.expectedControlId &&
    input.sourceWorkbook === input.liveWorkbook &&
    input.sourceActiveSheetIndex === input.liveActiveSheetIndex &&
    input.sourceSheet === input.liveSheet &&
    input.sourceSheetId === input.liveSheetId &&
    input.sourceTableCollection === input.liveTableCollection &&
    input.sourceTable === input.liveTable &&
    input.sourceTableId === input.liveTableId &&
    input.sourceSelectionEpoch === input.liveSelectionEpoch &&
    input.sourceRange.r1 === input.liveRange.r1 &&
    input.sourceRange.r2 === input.liveRange.r2 &&
    input.sourceRange.c1 === input.liveRange.c1 &&
    input.sourceRange.c2 === input.liveRange.c2 &&
    input.sourceStyleId === input.liveStyleId
  );
}

/**
 * Exact workbook/table generation and old/new alternate-text guard. The
 * dialog is session state; Apply may mutate only the table captured when the
 * dialog opened and only while the same selection generation is still live.
 */
export function spreadsheetTableAltTextWriterReceiptMatches(input: {
  readonly expectedControlId: string;
  readonly receiptControlId: string | null | undefined;
  readonly sourceWorkbook: unknown;
  readonly sourceActiveSheetIndex: number;
  readonly sourceSheet: unknown;
  readonly sourceSheetId: string;
  readonly sourceTableCollection: unknown;
  readonly sourceTable: unknown;
  readonly sourceTableId: string;
  readonly sourceSelectionEpoch: number;
  readonly sourceRange: Readonly<{
    r1: number;
    r2: number;
    c1: number;
    c2: number;
  }>;
  readonly sourceAltText: string;
  readonly sourceAltTextSummary: string;
  readonly nextAltText: string;
  readonly nextAltTextSummary: string;
  readonly liveWorkbook: unknown;
  readonly liveActiveSheetIndex: number;
  readonly liveSheet: unknown;
  readonly liveSheetId: string;
  readonly liveTableCollection: unknown;
  readonly liveTable: unknown;
  readonly liveTableId: string;
  readonly liveSelectionEpoch: number;
  readonly liveRange: Readonly<{
    r1: number;
    r2: number;
    c1: number;
    c2: number;
  }>;
  readonly liveAltText: string;
  readonly liveAltTextSummary: string;
  readonly liveNextAltText: string;
  readonly liveNextAltTextSummary: string;
}): boolean {
  return (
    input.receiptControlId === input.expectedControlId &&
    input.sourceWorkbook === input.liveWorkbook &&
    input.sourceActiveSheetIndex === input.liveActiveSheetIndex &&
    input.sourceSheet === input.liveSheet &&
    input.sourceSheetId === input.liveSheetId &&
    input.sourceTableCollection === input.liveTableCollection &&
    input.sourceTable === input.liveTable &&
    input.sourceTableId === input.liveTableId &&
    input.sourceSelectionEpoch === input.liveSelectionEpoch &&
    input.sourceRange.r1 === input.liveRange.r1 &&
    input.sourceRange.r2 === input.liveRange.r2 &&
    input.sourceRange.c1 === input.liveRange.c1 &&
    input.sourceRange.c2 === input.liveRange.c2 &&
    input.sourceAltText === input.liveAltText &&
    input.sourceAltTextSummary === input.liveAltTextSummary &&
    input.nextAltText === input.liveNextAltText &&
    input.nextAltTextSummary === input.liveNextAltTextSummary
  );
}

/**
 * Exact workbook/table generation and old/new-name guard for Table Name. The
 * writer updates structured references across the workbook, so both workbook
 * identity and the selected table entity must still be the captured objects.
 */
export function spreadsheetTableNameWriterReceiptMatches(input: {
  readonly expectedControlId: string;
  readonly receiptControlId: string | null | undefined;
  readonly sourceWorkbook: unknown;
  readonly sourceActiveSheetIndex: number;
  readonly sourceSheet: unknown;
  readonly sourceSheetId: string;
  readonly sourceTableCollection: unknown;
  readonly sourceTable: unknown;
  readonly sourceTableId: string;
  readonly sourceSelectionEpoch: number;
  readonly sourceRange: Readonly<{
    r1: number;
    r2: number;
    c1: number;
    c2: number;
  }>;
  readonly sourcePreviousName: string;
  readonly sourceNextName: string;
  readonly liveWorkbook: unknown;
  readonly liveActiveSheetIndex: number;
  readonly liveSheet: unknown;
  readonly liveSheetId: string;
  readonly liveTableCollection: unknown;
  readonly liveTable: unknown;
  readonly liveTableId: string;
  readonly liveSelectionEpoch: number;
  readonly liveRange: Readonly<{
    r1: number;
    r2: number;
    c1: number;
    c2: number;
  }>;
  readonly livePreviousName: string;
  readonly liveNextName: string;
}): boolean {
  return (
    input.receiptControlId === input.expectedControlId &&
    input.sourceWorkbook === input.liveWorkbook &&
    input.sourceActiveSheetIndex === input.liveActiveSheetIndex &&
    input.sourceSheet === input.liveSheet &&
    input.sourceSheetId === input.liveSheetId &&
    input.sourceTableCollection === input.liveTableCollection &&
    input.sourceTable === input.liveTable &&
    input.sourceTableId === input.liveTableId &&
    input.sourceSelectionEpoch === input.liveSelectionEpoch &&
    input.sourceRange.r1 === input.liveRange.r1 &&
    input.sourceRange.r2 === input.liveRange.r2 &&
    input.sourceRange.c1 === input.liveRange.c1 &&
    input.sourceRange.c2 === input.liveRange.c2 &&
    input.sourcePreviousName === input.livePreviousName &&
    input.sourceNextName === input.liveNextName
  );
}

/** Strict launch-generation and target-range guard for Resize Table. */
export function spreadsheetTableResizeWriterReceiptMatches(input: {
  readonly expectedControlId: string;
  readonly receiptControlId: string | null | undefined;
  readonly sourceWorkbook: unknown;
  readonly sourceActiveSheetIndex: number;
  readonly sourceSheet: unknown;
  readonly sourceSheetId: string;
  readonly sourceTableCollection: unknown;
  readonly sourceTable: unknown;
  readonly sourceTableId: string;
  readonly sourceSelectionEpoch: number;
  readonly sourceSelectionRange: Readonly<{
    r1: number;
    r2: number;
    c1: number;
    c2: number;
  }>;
  readonly sourceTableRange: string;
  readonly sourceTargetRange: string;
  readonly liveWorkbook: unknown;
  readonly liveActiveSheetIndex: number;
  readonly liveSheet: unknown;
  readonly liveSheetId: string;
  readonly liveTableCollection: unknown;
  readonly liveTable: unknown;
  readonly liveTableId: string;
  readonly liveSelectionEpoch: number;
  readonly liveSelectionRange: Readonly<{
    r1: number;
    r2: number;
    c1: number;
    c2: number;
  }>;
  readonly liveTableRange: string;
  readonly liveTargetRange: string;
}): boolean {
  return (
    input.receiptControlId === input.expectedControlId &&
    input.sourceWorkbook === input.liveWorkbook &&
    input.sourceActiveSheetIndex === input.liveActiveSheetIndex &&
    input.sourceSheet === input.liveSheet &&
    input.sourceSheetId === input.liveSheetId &&
    input.sourceTableCollection === input.liveTableCollection &&
    input.sourceTable === input.liveTable &&
    input.sourceTableId === input.liveTableId &&
    input.sourceSelectionEpoch === input.liveSelectionEpoch &&
    input.sourceSelectionRange.r1 === input.liveSelectionRange.r1 &&
    input.sourceSelectionRange.r2 === input.liveSelectionRange.r2 &&
    input.sourceSelectionRange.c1 === input.liveSelectionRange.c1 &&
    input.sourceSelectionRange.c2 === input.liveSelectionRange.c2 &&
    input.sourceTableRange === input.liveTableRange &&
    input.sourceTargetRange === input.liveTargetRange
  );
}

/** Strict identity and parameter guard for contextual native chart writers. */
export function spreadsheetChartDesignWriterReceiptMatches(input: {
  readonly expectedControlId: string;
  readonly receiptControlId: string | null | undefined;
  readonly sourceWorkbook: unknown;
  readonly sourceActiveSheetIndex: number;
  readonly sourceSheet: unknown;
  readonly sourceSheetId: string;
  readonly sourceSelection: unknown;
  readonly sourceChartCollection: unknown;
  readonly sourceChart: unknown;
  readonly sourceChartId: string;
  readonly sourceParameter: string;
  readonly liveWorkbook: unknown;
  readonly liveActiveSheetIndex: number;
  readonly liveSheet: unknown;
  readonly liveSheetId: string;
  readonly liveSelection: unknown;
  readonly liveChartCollection: unknown;
  readonly liveChart: unknown;
  readonly liveChartId: string;
  readonly liveParameter: string;
}): boolean {
  return (
    input.receiptControlId === input.expectedControlId &&
    input.sourceWorkbook === input.liveWorkbook &&
    input.sourceActiveSheetIndex === input.liveActiveSheetIndex &&
    input.sourceSheet === input.liveSheet &&
    input.sourceSheetId === input.liveSheetId &&
    input.sourceSelection === input.liveSelection &&
    input.sourceChartCollection === input.liveChartCollection &&
    input.sourceChart === input.liveChart &&
    input.sourceChartId === input.liveChartId &&
    input.sourceParameter === input.liveParameter
  );
}

/** Strict live workbook, PivotTable generation, selection and draft guard. */
export function spreadsheetPivotWriterReceiptMatches(input: {
  readonly expectedControlId: string;
  readonly receiptControlId: string | null | undefined;
  readonly sourceWorkbook: unknown;
  readonly sourceActiveSheetIndex: number;
  readonly sourceSheet: unknown;
  readonly sourceSheetId: string;
  readonly sourcePivot: unknown;
  readonly sourceSelectionEpoch: number;
  readonly sourceActiveRow: number;
  readonly sourceActiveCol: number;
  readonly sourceRange: Readonly<{
    r1: number;
    r2: number;
    c1: number;
    c2: number;
  }>;
  readonly sourceParameter: string;
  readonly liveWorkbook: unknown;
  readonly liveActiveSheetIndex: number;
  readonly liveSheet: unknown;
  readonly liveSheetId: string;
  readonly livePivot: unknown;
  readonly liveSelectionEpoch: number;
  readonly liveActiveRow: number;
  readonly liveActiveCol: number;
  readonly liveRange: Readonly<{
    r1: number;
    r2: number;
    c1: number;
    c2: number;
  }>;
  readonly liveParameter: string;
}): boolean {
  return (
    input.receiptControlId === input.expectedControlId &&
    input.sourceWorkbook === input.liveWorkbook &&
    input.sourceActiveSheetIndex === input.liveActiveSheetIndex &&
    input.sourceSheet === input.liveSheet &&
    input.sourceSheetId === input.liveSheetId &&
    input.sourcePivot === input.livePivot &&
    input.sourceSelectionEpoch === input.liveSelectionEpoch &&
    input.sourceActiveRow === input.liveActiveRow &&
    input.sourceActiveCol === input.liveActiveCol &&
    input.sourceRange.r1 === input.liveRange.r1 &&
    input.sourceRange.r2 === input.liveRange.r2 &&
    input.sourceRange.c1 === input.liveRange.c1 &&
    input.sourceRange.c2 === input.liveRange.c2 &&
    input.sourceParameter === input.liveParameter
  );
}

/** Strict live authority for creating one PivotTable-bound native chart. */
export function spreadsheetPivotChartCreateReceiptMatches(input: {
  readonly expectedControlId: string;
  readonly receiptControlId: string | null | undefined;
  readonly sourceWorkbook: unknown;
  readonly sourceActiveSheetIndex: number;
  readonly sourceSheet: unknown;
  readonly sourceSheetId: string;
  readonly sourcePivot: unknown;
  readonly sourceChartCollection: unknown;
  readonly sourceSelectionEpoch: number;
  readonly sourceActiveRow: number;
  readonly sourceActiveCol: number;
  readonly sourceRange: Readonly<{
    r1: number;
    r2: number;
    c1: number;
    c2: number;
  }>;
  readonly sourceOutputRange: string;
  readonly sourceChartType: string;
  readonly liveWorkbook: unknown;
  readonly liveActiveSheetIndex: number;
  readonly liveSheet: unknown;
  readonly liveSheetId: string;
  readonly livePivot: unknown;
  readonly liveChartCollection: unknown;
  readonly liveSelectionEpoch: number;
  readonly liveActiveRow: number;
  readonly liveActiveCol: number;
  readonly liveRange: Readonly<{
    r1: number;
    r2: number;
    c1: number;
    c2: number;
  }>;
  readonly liveOutputRange: string;
  readonly liveChartType: string;
}): boolean {
  return (
    input.receiptControlId === input.expectedControlId &&
    input.sourceWorkbook === input.liveWorkbook &&
    input.sourceActiveSheetIndex === input.liveActiveSheetIndex &&
    input.sourceSheet === input.liveSheet &&
    input.sourceSheetId === input.liveSheetId &&
    input.sourcePivot === input.livePivot &&
    input.sourceChartCollection === input.liveChartCollection &&
    input.sourceSelectionEpoch === input.liveSelectionEpoch &&
    input.sourceActiveRow === input.liveActiveRow &&
    input.sourceActiveCol === input.liveActiveCol &&
    input.sourceRange.r1 === input.liveRange.r1 &&
    input.sourceRange.r2 === input.liveRange.r2 &&
    input.sourceRange.c1 === input.liveRange.c1 &&
    input.sourceRange.c2 === input.liveRange.c2 &&
    input.sourceOutputRange === input.liveOutputRange &&
    input.sourceChartType === input.liveChartType
  );
}

/**
 * Strict launch-generation, cell and parameter guard for threaded-comment
 * Post, Resolve/Reopen and Delete writers.
 */
export function spreadsheetReviewCommentWriterReceiptMatches(input: {
  readonly expectedControlId: string;
  readonly receiptControlId: string | null | undefined;
  readonly sourceWorkbook: unknown;
  readonly sourceActiveSheetIndex: number;
  readonly sourceSheet: unknown;
  readonly sourceSheetId: string;
  readonly sourceSelectionEpoch: number;
  readonly sourceRow: number;
  readonly sourceCol: number;
  readonly sourceComment: unknown;
  readonly sourceParameter: string;
  readonly sourcePageId: string;
  readonly sourcePageAuthorityGeneration: number;
  readonly sourceActorId: string;
  readonly sourceActorName: string;
  readonly liveWorkbook: unknown;
  readonly liveActiveSheetIndex: number;
  readonly liveSheet: unknown;
  readonly liveSheetId: string;
  readonly liveSelectionEpoch: number;
  readonly liveRow: number;
  readonly liveCol: number;
  readonly liveComment: unknown;
  readonly liveParameter: string;
  readonly livePageId: string;
  readonly livePageAuthorityGeneration: number;
  readonly liveActorId: string;
  readonly liveActorName: string;
}): boolean {
  return (
    input.sourcePageId.trim().length > 0 &&
    Number.isSafeInteger(input.sourcePageAuthorityGeneration) &&
    input.sourcePageAuthorityGeneration > 0 &&
    input.sourceActorId.trim().length > 0 &&
    input.sourceActorName.trim().length > 0 &&
    input.receiptControlId === input.expectedControlId &&
    input.sourceWorkbook === input.liveWorkbook &&
    input.sourceActiveSheetIndex === input.liveActiveSheetIndex &&
    input.sourceSheet === input.liveSheet &&
    input.sourceSheetId === input.liveSheetId &&
    input.sourceSelectionEpoch === input.liveSelectionEpoch &&
    input.sourceRow === input.liveRow &&
    input.sourceCol === input.liveCol &&
    input.sourceComment === input.liveComment &&
    input.sourceParameter === input.liveParameter &&
    input.sourcePageId === input.livePageId &&
    input.sourcePageAuthorityGeneration ===
      input.livePageAuthorityGeneration &&
    input.sourceActorId === input.liveActorId &&
    input.sourceActorName === input.liveActorName
  );
}

const validationTextWithin = (
  value: unknown,
  limit: number,
  required = false,
): boolean => {
  if (value === undefined) return !required;
  if (typeof value !== "string") return false;
  const normalized = value
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "")
    .trim();
  return normalized.length <= limit && (!required || normalized.length > 0);
};

const validationComparisonSupported = (
  rule: Readonly<Record<string, unknown>>,
  firstKey: string,
  secondKey: string,
): boolean => {
  const operator = typeof rule.operator === "string" ? rule.operator : "";
  const first = rule[firstKey];
  const second = rule[secondKey];
  const firstPresent = first !== undefined;
  const secondPresent = second !== undefined;
  if (operator === "between" || operator === "notBetween") {
    return firstPresent && secondPresent;
  }
  if (
    [
      "equal",
      "notEqual",
      "greaterThan",
      "lessThan",
      "greaterThanOrEqual",
      "lessThanOrEqual",
    ].includes(operator)
  ) {
    return firstPresent;
  }
  return firstPresent || secondPresent;
};

const validationSourceRangeSupported = (value: string): boolean => {
  // A source on another worksheet is written to the x14 extension, the way
  // Excel stores it; the sheet prefix is plain or quoted ('Q1 Lists'!).
  const normalized = value
    .trim()
    .replace(/^=/, "")
    .replace(/^(?:'(?:[^']|'')+'|[^'!:]+)!/, "");
  if (normalized.includes("!")) return false;
  const match = /^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i.exec(
    normalized,
  );
  if (!match) return false;
  const columnNumber = (letters: string) =>
    [...letters.toUpperCase()].reduce(
      (total, letter) => total * 26 + letter.charCodeAt(0) - 64,
      0,
    );
  const startColumn = columnNumber(match[1]);
  const startRow = Number(match[2]);
  const endColumn = columnNumber(match[3] ?? match[1]);
  const endRow = Number(match[4] ?? match[2]);
  return (
    startColumn >= 1 &&
    startColumn <= 16_384 &&
    endColumn >= startColumn &&
    endColumn <= 16_384 &&
    startRow >= 1 &&
    startRow <= 1_048_576 &&
    endRow >= startRow &&
    endRow <= 1_048_576
  );
};

/**
 * Why a validation rule cannot be written into an imported workbook, or null
 * when it can. Every rule Excel itself can hold is writable (the writer puts
 * a rule whose formulas reach another worksheet in the x14 extension, as
 * Excel does); what remains are MRAI-only rule families and Excel's own
 * 255-character formula limit:
 *  - "dependent-list": MRAI's dependent dropdown has no Excel rule type;
 *  - "date-limits": Excel's Date rule has no weekday or "today" limit;
 *  - "formula-length": a list, source or formula past 255 characters;
 *  - "invalid": a malformed rule (the dialog never builds one).
 */
export type SpreadsheetDataValidationPatchbackBlocker =
  | "dependent-list"
  | "date-limits"
  | "formula-length"
  | "invalid";

export function spreadsheetDataValidationPatchbackRuleBlocker(
  value: unknown,
): SpreadsheetDataValidationPatchbackBlocker | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return "invalid";
  const rule = value as Readonly<Record<string, unknown>>;
  const range = rule.range;
  if (!range || typeof range !== "object" || Array.isArray(range)) return "invalid";
  const bounds = range as Readonly<Record<string, unknown>>;
  if (
    ![bounds.r1, bounds.r2, bounds.c1, bounds.c2].every(
      (coordinate) => Number.isSafeInteger(coordinate) && Number(coordinate) >= 0,
    ) ||
    Number(bounds.r2) < Number(bounds.r1) ||
    Number(bounds.c2) < Number(bounds.c1) ||
    Number(bounds.r2) >= 1_048_576 ||
    Number(bounds.c2) >= 16_384
  ) {
    return "invalid";
  }
  for (const [field, limit] of [
    ["inputTitle", 255],
    ["inputMessage", 2_000],
    ["errorTitle", 255],
    ["errorMessage", 2_000],
  ] as const) {
    if (!validationTextWithin(rule[field], limit)) return "invalid";
  }
  const boundFormulaFits = (field: string): boolean => {
    const formula = rule[field];
    if (formula === undefined) return true;
    return (
      typeof formula === "string" &&
      formula.trim().replace(/^=/, "").trim().length <= 255
    );
  };
  switch (rule.kind) {
    case "any":
      return null;
    case "dependent":
      return "dependent-list";
    case "list": {
      const sourceRange = rule.sourceRange;
      if (sourceRange !== undefined) {
        if (typeof sourceRange !== "string" || !sourceRange.trim()) return "invalid";
        if (sourceRange.trim().replace(/^=/, "").length > 255) return "formula-length";
        return validationSourceRangeSupported(sourceRange) ? null : "invalid";
      }
      if (!Array.isArray(rule.values) || rule.values.length === 0) return "invalid";
      if (
        rule.values.some(
          (entry) =>
            typeof entry !== "string" ||
            entry.trim().length === 0 ||
            entry.includes('"') ||
            entry.includes(","),
        )
      ) {
        return "invalid";
      }
      return `"${rule.values.join(",")}"`.length <= 255 ? null : "formula-length";
    }
    case "number":
      if (
        (rule.min !== undefined && !Number.isFinite(rule.min)) ||
        (rule.max !== undefined && !Number.isFinite(rule.max))
      ) {
        return "invalid";
      }
      if (!boundFormulaFits("minFormula") || !boundFormulaFits("maxFormula")) {
        return "formula-length";
      }
      return validationComparisonSupported(rule, "min", "max") ||
        (rule.integer === true &&
          rule.min === undefined &&
          rule.max === undefined) ||
        rule.minFormula !== undefined ||
        rule.maxFormula !== undefined
        ? null
        : "invalid";
    case "text":
      if (
        (rule.minLength !== undefined && !Number.isFinite(rule.minLength)) ||
        (rule.maxLength !== undefined && !Number.isFinite(rule.maxLength))
      ) {
        return "invalid";
      }
      return validationComparisonSupported(rule, "minLength", "maxLength")
        ? null
        : "invalid";
    case "formula":
      if (!validationTextWithin(rule.formula, 1_000, true)) return "invalid";
      return String(rule.formula).trim().replace(/^([^=])/, "=$1").length <= 255
        ? null
        : "formula-length";
    case "date": {
      if (
        rule.weekdays !== undefined ||
        rule.notBefore !== undefined ||
        rule.notAfter !== undefined
      ) {
        return "date-limits";
      }
      const validDate = (candidate: unknown) => {
        if (candidate === undefined) return true;
        if (typeof candidate !== "string") return false;
        const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(candidate.trim());
        if (!match) return false;
        const year = Number(match[1]);
        const month = Number(match[2]);
        const day = Number(match[3]);
        const parsed = new Date(Date.UTC(year, month - 1, day));
        return (
          parsed.getUTCFullYear() === year &&
          parsed.getUTCMonth() === month - 1 &&
          parsed.getUTCDate() === day
        );
      };
      if (!validDate(rule.min) || !validDate(rule.max)) return "invalid";
      return validationComparisonSupported(rule, "min", "max") ? null : "invalid";
    }
    case "time": {
      const validTime = (candidate: unknown) =>
        candidate === undefined ||
        /^(?:[01]\d|2[0-3]):[0-5]\d(?::[0-5]\d)?$/.test(String(candidate));
      return validTime(rule.min) &&
        validTime(rule.max) &&
        validationComparisonSupported(rule, "min", "max")
        ? null
        : "invalid";
    }
    default:
      return "invalid";
  }
}

/** Whether {@link spreadsheetDataValidationPatchbackRuleBlocker} finds nothing. */
export function spreadsheetDataValidationPatchbackRuleSupported(
  value: unknown,
): boolean {
  return spreadsheetDataValidationPatchbackRuleBlocker(value) === null;
}

/**
 * Whether the rules a validation edit touches have one owner per cell. The
 * save re-scopes every rule the selection reaches, as Excel does (a rule
 * applied or cleared over part of another rule takes only those cells), so a
 * partial intersection is ordinary. Rules that overlap one another are not a
 * shape Excel writes: their shared cells have no single rule to re-scope, and
 * they stay as the file has them.
 */
export function spreadsheetDataValidationPatchbackSelectionSupported(input: {
  readonly selectionRange: Readonly<{
    r1: number;
    r2: number;
    c1: number;
    c2: number;
  }>;
  readonly rules: readonly unknown[];
}): boolean {
  const selection = input.selectionRange;
  type Bounds = { r1: number; r2: number; c1: number; c2: number };
  const boundsOf = (value: unknown): Bounds | null => {
    if (!value || typeof value !== "object" || Array.isArray(value)) return null;
    const range = (value as Readonly<Record<string, unknown>>).range;
    if (!range || typeof range !== "object" || Array.isArray(range)) return null;
    const bounds = range as Readonly<Record<string, unknown>>;
    return {
      r1: Number(bounds.r1),
      r2: Number(bounds.r2),
      c1: Number(bounds.c1),
      c2: Number(bounds.c2),
    };
  };
  const overlaps = (left: Bounds, right: Bounds) =>
    left.r1 <= right.r2 &&
    right.r1 <= left.r2 &&
    left.c1 <= right.c2 &&
    right.c1 <= left.c2;
  // A rule over several areas (data-validation.ts ValidationAreas) is each
  // of its areas; its own areas are one rule, never two over a cell.
  const intersecting: Array<{ bounds: Bounds; rule: number }> = [];
  input.rules.forEach((rule, position) => {
    const areas =
      rule && typeof rule === "object" && Array.isArray((rule as { areas?: unknown }).areas)
        ? ((rule as { areas: unknown[] }).areas.map((area) => boundsOf({ range: area })))
        : [boundsOf(rule)];
    for (const bounds of areas) {
      if (bounds !== null && overlaps(bounds, selection)) intersecting.push({ bounds, rule: position });
    }
  });
  // By first row: a rectangle is asked only of those starting within its rows
  // (thousands of areas, R443-X20, are not asked of each other pairwise).
  intersecting.sort((left, right) => left.bounds.r1 - right.bounds.r1);
  for (let index = 0; index < intersecting.length; index += 1) {
    const current = intersecting[index]!;
    for (
      let other = index + 1;
      other < intersecting.length && intersecting[other]!.bounds.r1 <= current.bounds.r2;
      other += 1
    ) {
      if (current.rule !== intersecting[other]!.rule && overlaps(current.bounds, intersecting[other]!.bounds)) {
        return false;
      }
    }
  }
  return true;
}

/** Shared fail-closed identity guard for persistent Insert commands. */
export function spreadsheetInsertWriterReceiptMatches(input: {
  readonly expectedControlId: string;
  readonly receiptControlId: string | null | undefined;
  readonly sourceWorkbook: unknown;
  readonly sourceActiveSheetIndex: number;
  readonly sourceSheet: unknown;
  readonly sourceSheetId: string;
  readonly sourceSelectionEpoch: number;
  readonly sourceRange: Readonly<{ r1: number; r2: number; c1: number; c2: number }>;
  readonly liveWorkbook: unknown;
  readonly liveActiveSheetIndex: number;
  readonly liveSheet: unknown;
  readonly liveSheetId: string;
  readonly liveSelectionEpoch: number;
  readonly liveRange: Readonly<{ r1: number; r2: number; c1: number; c2: number }>;
}): boolean {
  return (
    input.receiptControlId === input.expectedControlId &&
    input.sourceWorkbook === input.liveWorkbook &&
    input.sourceActiveSheetIndex === input.liveActiveSheetIndex &&
    input.sourceSheet === input.liveSheet &&
    input.sourceSheetId === input.liveSheetId &&
    input.sourceSelectionEpoch === input.liveSelectionEpoch &&
    input.sourceRange.r1 === input.liveRange.r1 &&
    input.sourceRange.r2 === input.liveRange.r2 &&
    input.sourceRange.c1 === input.liveRange.c1 &&
    input.sourceRange.c2 === input.liveRange.c2
  );
}

export interface SpreadsheetAdvancedFilterModeAuthority {
  readonly enabled: boolean;
  readonly reasonKey: SpreadsheetVisibleCommandReason | null;
}

/**
 * Resolve the confirmed Advanced Filter mode against the live writer route.
 * Copy mode owns only destination cells. In-place mode owns the dedicated
 * advancedFilterState projection and cannot coexist with worksheet AutoFilter:
 * both grammars own the same persisted row-hidden surface.
 */
export function resolveSpreadsheetAdvancedFilterModeAuthority(
  mode: "in-place" | "copy",
  context: Pick<
    SpreadsheetVisibleCommandContext,
    "hasSelection" | "nativeWriteRoute" | "selectionEditable" | "sheetProtected"
  > & {
    worksheetAutoFilterActive?: boolean;
  },
): SpreadsheetAdvancedFilterModeAuthority {
  if (!context.hasSelection) {
    return {
      enabled: false,
      reasonKey: "spreadsheet.capability.selectionRequired",
    };
  }
  if (
    !context.selectionEditable ||
    (mode === "in-place" && context.sheetProtected)
  ) {
    return {
      enabled: false,
      reasonKey: "spreadsheet.capability.sheetProtected",
    };
  }
  if (mode === "in-place" && context.worksheetAutoFilterActive === true) {
    return {
      enabled: false,
      reasonKey: "spreadsheet.capability.writerUnavailable",
    };
  }
  // An imported workbook takes a new in-place filter as well: the writer
  // adds _FilterDatabase, Criteria, filterMode and the hidden rows.
  const contract =
    SPREADSHEET_VISIBLE_IO_CONTRACTS[
      mode === "copy" ? "worksheet-cells" : "worksheet-advanced-filter"
    ];
  return spreadsheetVisibleIoWriterReady(contract, context.nativeWriteRoute)
    ? { enabled: true, reasonKey: null }
    : {
        enabled: false,
        reasonKey: "spreadsheet.capability.writerUnavailable",
      };
}

function capabilityRow(
  definition: SpreadsheetVisibleControlDefinition,
): SpreadsheetFeatureCapabilityContract {
  const isBound = definition.command === "bound";
  const isDisclosure = definition.command === "disclosure";
  const isSessionOnly = definition.ioProfile === "session-only";
  const internalShape = SPREADSHEET_SHAPE_FORMAT_INTERNAL_CONTROL_ID_SET.has(
    definition.controlId,
  );
  const internalPicture =
    SPREADSHEET_PICTURE_FORMAT_INTERNAL_CONTROL_ID_SET.has(
      definition.controlId,
    );
  const internalOnly = internalShape || internalPicture;
  const ioContract = SPREADSHEET_VISIBLE_IO_CONTRACTS[definition.ioProfile];
  return {
    featureId: definition.featureId,
    canonicalId: definition.canonicalId,
    controlId: definition.controlId,
    sourceVersion: SPREADSHEET_CAPABILITY_CATALOG_SCHEMA_VERSION,
    officeBaseline: {
      platform: "Excel for the web",
      version: null,
      evidenceRefs: definition.controlId === SPREADSHEET_COPILOT_APPLY_CONTROL_ID
        ? ["https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel"]
        : ["EXCEL_SPREADSHEET_FEATURE_REPLICATION_TREE.md#15"],
    },
    surface: {
      disposition: "cataloged",
      physicalPath: [
        definition.controlId === SPREADSHEET_COPILOT_APPLY_CONTROL_ID ? "Internal.Copilot" : internalOnly
          ? internalShape
            ? "Internal.ShapeFormat"
            : "Internal.PictureFormat"
          : RIBBON_PATH_BY_GROUP[definition.group],
        definition.controlId,
      ],
      group: definition.group,
      order: definition.order,
      controlKind: definition.controlKind,
      availabilityPredicate: definition.predicate,
    },
    command: {
      disposition: isBound
        ? "bound"
        : isDisclosure
          ? "external"
          : "unsupported",
      key: isBound ? commandKey(definition.featureId) : null,
      handlerRef: isBound ? handlerRef(definition.featureId) : null,
      undoPolicy: isBound
        ? definition.undoPolicy
        : isDisclosure
          ? "external"
          : "unknown",
      availabilityPredicate: definition.predicate,
    },
    authoring: {
      disposition: isBound
        ? isSessionOnly
          ? "projection"
          : "native"
        : isDisclosure
          ? "external"
          : "unsupported",
    },
    semantics: {
      defaults: {},
      limits: {},
      interactions: isDisclosure ? ["open", "Escape"] : ["click", "Enter"],
      shortcutIds: [],
    },
    domain: ioContract.domain,
    import: ioContract.import,
    fromScratchExport: ioContract.fromScratchExport,
    patchbackExport: ioContract.patchbackExport,
    officeOutcome: ioContract.officeOutcome,
    acceptanceContractIds: [`XLSACT.${definition.featureId}`],
    owner: "Spreadsheet",
  };
}

export const SPREADSHEET_VISIBLE_COMMAND_CAPABILITY_CATALOG =
  readSpreadsheetFeatureCapabilityCatalog({
    version: SPREADSHEET_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: SPREADSHEET_VISIBLE_COMMAND_DEFINITIONS.map(capabilityRow),
  });

interface SpreadsheetVisibleSemanticCapability {
  readonly id: string;
  readonly reasonKey?: SpreadsheetVisibleCommandReason;
  readonly capability: SpreadsheetFeatureCapabilityContract;
}

const boundCapabilities =
  SPREADSHEET_VISIBLE_COMMAND_CAPABILITY_CATALOG.features.filter(
    (capability) =>
      capability.command.disposition === "bound" ||
      capability.command.disposition === "parameterized",
  );
const semanticCatalog = Object.freeze(
  [
    ...new Map(
      boundCapabilities.map((capability) => [capability.featureId, capability]),
    ).values(),
  ]
    .sort((left, right) => left.featureId.localeCompare(right.featureId))
    .map((capability) =>
      Object.freeze({
        id: capability.featureId,
        capability,
      }),
    ),
);
const semanticByFeature = new Map(
  semanticCatalog.map((capability) => [capability.id, capability]),
);

export type SpreadsheetVisibleBoundFeatureId =
  (typeof semanticCatalog)[number]["id"];

export type SpreadsheetVisibleCommandHandlers = Readonly<
  Record<string, FeatureCommandHandler>
>;

// Sheet Views are kept with the stored workbook (Microsoft Support, "Create and
// manage Sheet Views in Excel": only documents stored in OneDrive or SharePoint
// can use them). While there is no saved workbook that is the reason a blocked
// Sheet View command gives, in its own words rather than New Window's, because
// saving is the one thing that unblocks it.
function sheetViewBlockedReason(
  context: SpreadsheetVisibleCommandContext,
  reasonKey: SpreadsheetVisibleCommandReason,
): SpreadsheetVisibleCommandReason {
  return context.canOpenWorkbookWindow
    ? reasonKey
    : "spreadsheet.capability.sheetViewSavedWorkbookRequired";
}

function predicateDecision(
  predicate: SpreadsheetVisiblePredicate,
  context: SpreadsheetVisibleCommandContext,
  fallbackReason: SpreadsheetVisibleCommandReason | null,
): { enabled: boolean; reasonKey: SpreadsheetVisibleCommandReason | null } {
  const blocked = (reasonKey: SpreadsheetVisibleCommandReason) => ({
    enabled: false,
    reasonKey,
  });
  // One Protect Sheet permission on the active sheet. Without the resolved
  // map a protected sheet allows none (the reading before permissions).
  const allows = (permission: keyof SheetProtectionPermissions): boolean =>
    !context.sheetProtected ||
    (context.protectionAllows
      ? context.protectionAllows[permission]
      : permission === "formatCells" && context.formatCellsAllowed === true);
  const editableFamily = (available: boolean | undefined) => {
    // Formatting is its own permission on a protected sheet: an unlocked
    // cell may be entered, not restyled (live audit B138), and with Format
    // cells any cell may be restyled, locked or not ("format any cell",
    // Worksheet.Protect AllowFormattingCells).
    if (context.sheetProtected) {
      if (!allows("formatCells"))
        return blocked("spreadsheet.capability.sheetProtected");
    } else if (!context.selectionEditable) {
      return blocked("spreadsheet.capability.sheetProtected");
    }
    if (!context.hasSelection)
      return blocked("spreadsheet.capability.selectionRequired");
    return available === true
      ? { enabled: true, reasonKey: null }
      : blocked("spreadsheet.capability.commandUnavailable");
  };
  switch (predicate) {
    case "copilot-apply":
      return context.copilotApplyAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.commandUnavailable");
    case "always":
      return { enabled: true, reasonKey: null };
    case "manual-calculation":
      return context.manualCalculationAvailable !== false
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.manualCalculationUnavailable");
    case "active-comment":
      return context.activeComment
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.activeCommentRequired");
    case "active-remark-editable":
      // The active cell's threaded comment or its legacy note. Excel's Edit
      // objects opens notes on a protected sheet; threaded comments stay shut.
      if (
        context.sheetProtected &&
        !(context.activeNote && !context.activeComment && allows("editObjects"))
      )
        return blocked("spreadsheet.capability.sheetProtected");
      if (!context.activeComment && !context.activeNote)
        return blocked(
          fallbackReason ?? "spreadsheet.capability.activeCommentRequired",
        );
      if (context.activeRemarkBlocked)
        return blocked("spreadsheet.capability.writerUnavailable");
      if (context.activeNote && !context.activeComment)
        return { enabled: true, reasonKey: null };
      if (context.commentAuthoringAvailable !== true)
        return blocked("spreadsheet.capability.commentAuthorUnavailable");
      return { enabled: true, reasonKey: null };
    case "active-comment-editable":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (!context.activeComment)
        return blocked(
          fallbackReason ?? "spreadsheet.capability.activeCommentRequired",
        );
      if (context.activeRemarkBlocked)
        return blocked("spreadsheet.capability.writerUnavailable");
      if (context.commentAuthoringAvailable !== true)
        return blocked("spreadsheet.capability.commentAuthorUnavailable");
      return { enabled: true, reasonKey: null };
    case "active-native-chart-editable":
      if (!context.activeNativeChart)
        return blocked("spreadsheet.capability.commandUnavailable");
      if (!allows("editObjects"))
        return blocked("spreadsheet.capability.sheetProtected");
      return context.activeNativeChartEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.writerUnavailable");
    case "active-native-chart-color-editable":
      if (!context.activeNativeChart)
        return blocked("spreadsheet.capability.commandUnavailable");
      if (!allows("editObjects"))
        return blocked("spreadsheet.capability.sheetProtected");
      return context.activeNativeChartColorEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.writerUnavailable");
    case "active-native-chart-line-editable":
      if (!context.activeNativeChart)
        return blocked("spreadsheet.capability.commandUnavailable");
      if (!allows("editObjects"))
        return blocked("spreadsheet.capability.sheetProtected");
      return context.activeNativeChartLineEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.writerUnavailable");
    case "active-native-chart-style-editable":
      if (!context.activeNativeChart)
        return blocked("spreadsheet.capability.commandUnavailable");
      if (!allows("editObjects"))
        return blocked("spreadsheet.capability.sheetProtected");
      return context.activeNativeChartStyleEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.writerUnavailable");
    case "active-native-chart":
      return context.activeNativeChart
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.commandUnavailable");
    case "active-native-chart-series-orientation-editable":
      if (!context.activeNativeChart)
        return blocked("spreadsheet.capability.commandUnavailable");
      if (!allows("editObjects"))
        return blocked("spreadsheet.capability.sheetProtected");
      return context.activeNativeChartSeriesOrientationEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.commandUnavailable");
    case "active-pivot":
      return context.activePivot === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.commandUnavailable",
          );
    case "active-pivot-output":
      return context.activePivot === true &&
        context.activePivotOutputAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.commandUnavailable",
          );
    case "active-pivot-chart-create-editable":
      if (context.activePivot !== true)
        return blocked("spreadsheet.capability.commandUnavailable");
      if (!allows("pivotTables"))
        return blocked("spreadsheet.capability.sheetProtected");
      // An imported workbook appends the PivotChart as a native part bound
      // to its (editable) PivotTable, as a new workbook writes it.
      return context.activePivotEditable === true &&
        (context.nativeWriteRoute === "from-scratch" ||
          context.nativeWriteRoute === "patchback")
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.writerUnavailable");
    case "active-pivot-chart":
      return context.activePivotChart === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.commandUnavailable");
    case "active-pivot-chart-editable":
      if (context.activePivotChart !== true)
        return blocked("spreadsheet.capability.commandUnavailable");
      if (!allows("pivotTables"))
        return blocked("spreadsheet.capability.sheetProtected");
      return context.activePivotChartEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.writerUnavailable");
    case "active-pivot-chart-options-editable":
      if (context.activePivotChart !== true)
        return blocked("spreadsheet.capability.commandUnavailable");
      if (!allows("pivotTables"))
        return blocked("spreadsheet.capability.sheetProtected");
      return context.activePivotChartOptionsEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.writerUnavailable");
    case "active-pivot-editable":
      if (context.activePivot !== true)
        return blocked(
          fallbackReason ?? "spreadsheet.capability.commandUnavailable",
        );
      if (!allows("pivotTables"))
        return blocked("spreadsheet.capability.sheetProtected");
      return context.activePivotEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "active-pivot-source-editable":
      if (context.activePivot !== true)
        return blocked(
          fallbackReason ?? "spreadsheet.capability.commandUnavailable",
        );
      if (!allows("pivotTables"))
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.activePivotHasSourceRange === false)
        return blocked("spreadsheet.capability.pivotSourceIsScenarios");
      return context.activePivotEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "active-pivot-clear-all-editable":
      if (!allows("pivotTables"))
        return blocked("spreadsheet.capability.sheetProtected");
      if (!context.activePivot)
        return blocked("spreadsheet.capability.selectionRequired");
      return context.activePivotClearAllAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "active-shape-editable":
      if (context.activeShape !== true)
        return blocked(
          fallbackReason ?? "spreadsheet.capability.commandUnavailable",
        );
      if (!allows("editObjects"))
        return blocked("spreadsheet.capability.sheetProtected");
      return context.activeShapeEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "shape-rotation-apply":
      if (context.activeShape !== true)
        return blocked(
          fallbackReason ?? "spreadsheet.capability.commandUnavailable",
        );
      if (!allows("editObjects"))
        return blocked("spreadsheet.capability.sheetProtected");
      return context.shapeRotationApplyAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "active-picture-effects-editable":
      if (context.activePicture !== true)
        return blocked(
          fallbackReason ?? "spreadsheet.capability.commandUnavailable",
        );
      if (!allows("editObjects"))
        return blocked("spreadsheet.capability.sheetProtected");
      return context.activePictureEffectsEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "active-picture-metadata-editable":
      if (context.activePicture !== true)
        return blocked(
          fallbackReason ?? "spreadsheet.capability.commandUnavailable",
        );
      if (!allows("editObjects"))
        return blocked("spreadsheet.capability.sheetProtected");
      return context.activePictureMetadataEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "active-picture-transform-editable":
      if (context.activePicture !== true)
        return blocked(
          fallbackReason ?? "spreadsheet.capability.commandUnavailable",
        );
      if (!allows("editObjects"))
        return blocked("spreadsheet.capability.sheetProtected");
      return context.activePictureTransformEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "picture-rotation-apply":
      if (context.activePicture !== true)
        return blocked(
          fallbackReason ?? "spreadsheet.capability.commandUnavailable",
        );
      if (!allows("editObjects"))
        return blocked("spreadsheet.capability.sheetProtected");
      return context.pictureRotationApplyAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "active-slicer-editable":
      if (context.activeSlicer !== true)
        return blocked(
          fallbackReason ?? "spreadsheet.capability.commandUnavailable",
        );
      if (!allows("editObjects"))
        return blocked("spreadsheet.capability.sheetProtected");
      return context.activeSlicerEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "active-slicer-filter-editable":
      if (context.activeSlicer !== true)
        return blocked(
          fallbackReason ?? "spreadsheet.capability.commandUnavailable",
        );
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.activeSlicerFilterEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "active-sparkline-editable":
      if (context.activeSparkline !== true)
        return blocked(
          fallbackReason ?? "spreadsheet.capability.commandUnavailable",
        );
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.activeSparklineEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "active-filter":
      // Excel: with protection on, filters change only under Use AutoFilter;
      // Clear changes criteria, never the AutoFilter itself.
      if (!allows("autoFilter"))
        return blocked("spreadsheet.capability.sheetProtected");
      return context.activeFilter
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.activeFilterRequired");
    case "active-table-style-editable":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (!context.activeTable)
        return blocked("spreadsheet.capability.selectionRequired");
      return context.activeTableStyleEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
          fallbackReason ?? "spreadsheet.capability.writerUnavailable",
        );
    case "active-slicer-insert":
      if (!allows("editObjects"))
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.activePivotSlicerInsertAvailable === true)
        return { enabled: true, reasonKey: null };
      if (!context.activeTable)
        return blocked("spreadsheet.capability.selectionRequired");
      return context.activeTableSlicerInsertAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "active-pivot-slicer-insert":
      if (!allows("editObjects"))
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.activePivot !== true)
        return blocked("spreadsheet.capability.selectionRequired");
      return context.activePivotSlicerInsertAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "pivot-slicer-insert-apply":
      return context.pivotSlicerInsertApplyAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "active-table-slicer-insert":
      if (!allows("editObjects"))
        return blocked("spreadsheet.capability.sheetProtected");
      if (!context.activeTable)
        return blocked("spreadsheet.capability.selectionRequired");
      return context.activeTableSlicerInsertAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "active-table-pivot-source":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (!context.activeTable)
        return blocked("spreadsheet.capability.selectionRequired");
      return context.activeTablePivotSourceAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "active-table-totals-editable":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (!context.activeTable)
        return blocked("spreadsheet.capability.selectionRequired");
      return context.activeTableTotalsEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "active-table-alt-text-editable":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (!context.activeTable)
        return blocked("spreadsheet.capability.selectionRequired");
      return context.activeTableAltTextEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked(fallbackReason ?? "spreadsheet.capability.writerUnavailable");
    case "active-table-name-editable":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (!context.activeTable)
        return blocked("spreadsheet.capability.selectionRequired");
      return context.activeTableNameEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "active-table-resize-editable":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (!context.activeTable)
        return blocked("spreadsheet.capability.selectionRequired");
      return context.activeTableResizeEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "table-resize-apply":
      return context.tableResizeApplyAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
          fallbackReason ?? "spreadsheet.capability.writerUnavailable",
        );
    case "table-slicer-insert-apply":
      return context.tableSlicerInsertApplyAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "table-alt-text-apply":
      return context.tableAltTextApplyAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(fallbackReason ?? "spreadsheet.capability.writerUnavailable");
    case "font-editable":
      // Home ▸ Font on a selected shape formats the shape's text, never the
      // cells behind it, and is gated like every other shape write.
      if (context.homeFontShapeScope === true) {
        if (!allows("editObjects"))
          return blocked("spreadsheet.capability.sheetProtected");
        return context.homeFontShapeEditable === true
          ? { enabled: true, reasonKey: null }
          : blocked("spreadsheet.capability.writerUnavailable");
      }
      // A selected chart: the chart's own text (`c:chartSpace/c:txPr`, as
      // Excel's Home ▸ Font formats a selected chart).
      if (context.homeChartTextScope === true) {
        if (!allows("editObjects"))
          return blocked("spreadsheet.capability.sheetProtected");
        return context.homeChartTextEditable === true
          ? { enabled: true, reasonKey: null }
          : blocked("spreadsheet.capability.writerUnavailable");
      }
      // A picture, a slicer, an ink stroke or a preserved diagram has no
      // text Home can format, and the cells under it are not the selection.
      if (context.homeObjectSelected === true)
        return blocked("spreadsheet.capability.objectSelected");
      return editableFamily(context.homeFontControlsAvailable);
    case "cell-font-editable":
      // Fill Color and Borders are CELL formatting; with an object selected
      // they would reach cells the user is not looking at, so they stand down.
      if (context.homeFontShapeScope === true || context.homeObjectSelected === true)
        return blocked("spreadsheet.capability.objectSelected");
      return editableFamily(context.homeFontControlsAvailable);
    case "format-cells-apply":
      if (context.homeObjectSelected === true)
        return blocked("spreadsheet.capability.objectSelected");
      // Format cells formats any cell of a protected sheet, locked or not.
      if (context.sheetProtected ? !allows("formatCells") : !context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.formatCellsApplyAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(fallbackReason ?? "spreadsheet.capability.writerUnavailable");
    case "formula-evaluation":
      return context.formulaEvaluationAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.commandUnavailable",
          );
    case "formula-trace-clear":
      return context.formulaTraceAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.commandUnavailable",
          );
    case "alignment-editable":
      // Left / Center / Right and Top / Middle / Bottom on a selected shape
      // align the shape's label (Excel); on any other object they stand down.
      if (context.homeFontShapeScope === true) {
        if (!allows("editObjects"))
          return blocked("spreadsheet.capability.sheetProtected");
        return context.homeFontShapeEditable === true
          ? { enabled: true, reasonKey: null }
          : blocked("spreadsheet.capability.writerUnavailable");
      }
      if (context.homeObjectSelected === true)
        return blocked("spreadsheet.capability.objectSelected");
      return editableFamily(context.homeAlignmentControlsAvailable);
    case "cell-alignment-editable":
      // Wrap Text and the indents are CELL formatting only.
      if (context.homeObjectSelected === true)
        return blocked("spreadsheet.capability.objectSelected");
      return editableFamily(context.homeAlignmentControlsAvailable);
    case "number-format-editable":
      if (context.homeObjectSelected === true)
        return blocked("spreadsheet.capability.objectSelected");
      return editableFamily(context.homeNumberFormatControlsAvailable);
    case "cell-style-editable":
      if (context.homeObjectSelected === true)
        return blocked("spreadsheet.capability.objectSelected");
      return editableFamily(context.homeCellStyleControlsAvailable);
    case "cell-shift-down-editable":
    case "cell-shift-right-editable":
    case "cell-shift-up-editable":
    case "cell-shift-left-editable": {
      // Insert/Delete Cells (shifting) is never open on a protected sheet,
      // whatever its Insert/Delete rows/columns permissions say.
      if (!context.selectionEditable || context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (!context.hasSelection)
        return blocked("spreadsheet.capability.selectionRequired");
      if (context.homeStructureControlsAvailable !== true)
        return blocked("spreadsheet.capability.commandUnavailable");
      const available =
        predicate === "cell-shift-down-editable"
          ? context.cellShiftDownAvailable
          : predicate === "cell-shift-right-editable"
            ? context.cellShiftRightAvailable
            : predicate === "cell-shift-up-editable"
              ? context.cellShiftUpAvailable
              : context.cellShiftLeftAvailable;
      return available === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ??
              "spreadsheet.capability.cellShiftUnsafe",
          );
    }
    case "conditional-format-selection-clearable":
      // Every conditional-format writer answers to Excel's Format cells on a
      // protected sheet (any cell, locked or not); the menu says so instead
      // of answering with a toast.
      if (context.sheetProtected ? !allows("formatCells") : !context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      if (!context.hasSelection)
        return blocked("spreadsheet.capability.selectionRequired");
      if (context.conditionalFormatSelectionClearBlocked)
        return blocked("spreadsheet.capability.writerUnavailable");
      return context.homeSelectionHasConditionalRule === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.commandUnavailable");
    case "conditional-format-sheet-clearable":
      if (context.sheetProtected ? !allows("formatCells") : !context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.homeSheetHasConditionalRules === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.commandUnavailable");
    case "conditional-format-editable":
      if (!allows("formatCells"))
        return blocked("spreadsheet.capability.sheetProtected");
      return context.hasSelection
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.selectionRequired");
    case "conditional-format-manager-apply":
      if (context.sheetProtected ? !allows("formatCells") : !context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      if (!context.hasSelection)
        return blocked("spreadsheet.capability.selectionRequired");
      return context.conditionalFormatManagerApplyAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.writerUnavailable",
          );
    case "clipboard-paste-editable":
      if (!context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      if (!context.hasSelection)
        return blocked("spreadsheet.capability.selectionRequired");
      // A sorted or filtered sheet view refuses a multi-row range; otherwise
      // Paste needs a range copied or cut in this workbook (a paste from
      // another app arrives through Ctrl+V).
      if (context.selectionClipboardAvailable !== true)
        return blocked("spreadsheet.capability.clipboardUnavailable");
      return context.richClipboardAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.pasteSourceRequired");
    case "clipboard-paste-available":
      // A copied chart opens Paste Special on what a chart pastes as (the
      // chart or its picture), whatever cells are selected.
      if (context.chartClipboardAvailable === true) return { enabled: true, reasonKey: null };
      // A sorted or filtered sheet view refuses a multi-row range; otherwise
      // Paste needs a range copied or cut in this workbook (a paste from
      // another app arrives through Ctrl+V).
      if (context.selectionClipboardAvailable !== true)
        return blocked("spreadsheet.capability.clipboardUnavailable");
      return context.richClipboardAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.pasteSourceRequired");
    case "paste-special-apply":
      if (!context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      if (!context.richClipboardAvailable)
        return blocked("spreadsheet.capability.clipboardUnavailable");
      return context.pasteSpecialApplyAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(fallbackReason ?? "spreadsheet.capability.writerUnavailable");
    case "selection-clipboard":
    case "clipboard-copy-picture":
      if (!context.hasSelection)
        return blocked("spreadsheet.capability.selectionRequired");
      return context.selectionClipboardAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.clipboardUnavailable");
    case "reapplicable-filter-or-sort":
      // Reapply re-runs the AutoFilter's criteria: Use AutoFilter.
      if (!allows("autoFilter"))
        return blocked("spreadsheet.capability.sheetProtected");
      return (context.reapplicableFilterOrSort ?? context.activeFilter)
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ??
              "spreadsheet.capability.reapplyFilterOrSortRequired",
          );
    case "comments":
      return context.commentedCellCount > 0
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.noComments");
    case "comment-post-editable":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.activeRemarkBlocked)
        return blocked("spreadsheet.capability.commandUnavailable");
      return context.commentPostAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.commandUnavailable");
    case "notes":
      return (context.notedCellCount ?? 0) > 0
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.noNotes");
    case "new-comment":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.activeRemarkBlocked)
        return blocked("spreadsheet.capability.writerUnavailable");
      if (context.commentAuthoringAvailable !== true)
        return blocked("spreadsheet.capability.commentAuthorUnavailable");
      return context.activeNote
        ? blocked("spreadsheet.capability.commentKindConflict")
        : { enabled: true, reasonKey: null };
    case "active-note-editable":
      // Excel's Edit objects covers adding and editing notes.
      if (!allows("editObjects"))
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.activeRemarkBlocked)
        return blocked("spreadsheet.capability.commandUnavailable");
      return context.activeNote
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.activeNoteRequired");
    case "new-note":
      if (!allows("editObjects"))
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.activeRemarkBlocked)
        return blocked("spreadsheet.capability.commandUnavailable");
      return context.activeComment
        ? blocked("spreadsheet.capability.commandUnavailable")
        : { enabled: true, reasonKey: null };
    case "data-source-inventory":
      return context.dataSourceInventoryInspectable
        ? { enabled: true, reasonKey: null }
        : blocked(fallbackReason ?? "spreadsheet.capability.noQueries");
    case "data-form-launchable":
      if (context.sheetProtected || !context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.dataFormLaunchAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.dataFormListRequired",
          );
    case "data-form-open":
      return context.dataFormOpenAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.commandUnavailable");
    case "data-form-restore":
      return context.dataFormOpenAvailable === true &&
        context.dataFormRestoreAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.commandUnavailable");
    case "data-form-delete":
      if (context.sheetProtected || !context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.dataFormDeleteAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(fallbackReason ?? "spreadsheet.capability.commandUnavailable");
    case "data-form-delete-apply":
      if (context.sheetProtected || !context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.dataFormDeleteApplyAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(fallbackReason ?? "spreadsheet.capability.commandUnavailable");
    case "data-from-picture":
      if (!context.selectionEditable || context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.dataFromPictureAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            context.dataFromPictureReason ??
              fallbackReason ??
              "spreadsheet.capability.workspaceAiProviderUnavailable",
          );
    case "data-from-picture-clipboard":
      if (!context.selectionEditable || context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.dataFromPictureAvailable !== true)
        return blocked(
          context.dataFromPictureReason ??
            "spreadsheet.capability.workspaceAiProviderUnavailable",
        );
      return context.dataFromPictureClipboardAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.clipboardUnavailable");
    case "data-validation-launcher":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.hasSelection
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.selectionRequired");
    case "data-validation-apply":
      if (context.sheetProtected || !context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      if (!context.hasSelection)
        return blocked("spreadsheet.capability.selectionRequired");
      return context.dataValidationApplyAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(fallbackReason ?? "spreadsheet.capability.writerUnavailable");
    case "data-validation-clear-draft":
      if (context.sheetProtected || !context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.dataValidationClearDraftAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(fallbackReason ?? "spreadsheet.capability.commandUnavailable");
    case "find-replace-match":
      return context.findReplaceMatchAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(fallbackReason ?? "spreadsheet.capability.commandUnavailable");
    case "find-replace-current-editable":
      if (context.findReplaceMatchAvailable !== true)
        return blocked(fallbackReason ?? "spreadsheet.capability.commandUnavailable");
      return context.findReplaceCurrentEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.sheetProtected");
    case "find-replace-all-editable":
      if (context.findReplaceMatchAvailable !== true)
        return blocked(fallbackReason ?? "spreadsheet.capability.commandUnavailable");
      return context.findReplaceAllEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.sheetProtected");
    case "fullscreen":
      return context.fullscreenSupported
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.fullscreenUnavailable");
    case "split-pane-toggle":
      // Removing a split never needs the panes drawn: Excel's Split toggle
      // always removes one. Creating one needs a view that has panes.
      if (context.splitPanePresent === true) {
        return { enabled: true, reasonKey: null };
      }
      if (context.splitPaneRenderingSupported !== true) {
        return blocked("spreadsheet.capability.splitPaneLayersUnsupported");
      }
      if (context.splitPaneSelectionFitsGrid !== true) {
        return blocked(
          "spreadsheet.capability.splitPaneSelectionOutsideGrid",
        );
      }
      return { enabled: true, reasonKey: null };
    case "import-ready":
      if (context.importInProgress) {
        return blocked("spreadsheet.capability.importInProgress");
      }
      return { enabled: true, reasonKey: null };
    // Importing values writes them into the selection, so it answers to sheet
    // protection as well: pasteGridIntoCells refuses a locked destination, and
    // without this the entry walks the user through a file picker and a preview
    // before the commit turns them away. Replacing the whole workbook
    // (FROM_WORKBOOK) is a different act and keeps the plain import-ready.
    case "import-ready-editable":
      if (context.importInProgress) {
        return blocked("spreadsheet.capability.importInProgress");
      }
      if (!context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.hasSelection
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.selectionRequired");
    case "multi-column-editable":
      if (!context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.hasMultiColumnSelection
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.multiColumnSelectionRequired");
    // The merge commands read the selection Excel gives them: a merged cell
    // selects its whole area, so Merge & Center stays live on one (there it is
    // the unmerge toggle) instead of greying out on a 1x1 anchor selection.
    case "merge-selection-editable":
      if (context.homeObjectSelected === true)
        return blocked("spreadsheet.capability.objectSelected");
      if (!context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.mergeSelectionInTable === true)
        return blocked("spreadsheet.capability.mergeInTable");
      return (context.mergeSelectionMultiCell ?? context.hasMultiSelection)
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.multiSelectionRequired");
    case "merge-across-editable":
      if (context.homeObjectSelected === true)
        return blocked("spreadsheet.capability.objectSelected");
      if (!context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.mergeSelectionInTable === true)
        return blocked("spreadsheet.capability.mergeInTable");
      return (
        context.mergeSelectionMultiColumn ?? context.hasMultiColumnSelection
      )
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.multiColumnSelectionRequired");
    case "multi-row-editable":
      if (!context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.hasMultiRowSelection
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.multiRowSelectionRequired");
    case "data-analysis":
      return context.dataAnalysisAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.analyzeDataUnavailable");
    case "data-query-properties":
      // Opening only reads; the dialog itself says when nothing can change.
      return context.dataQueryPropertiesAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.queryTableSelectionRequired");
    case "data-query-create":
    case "data-query-refresh":
    case "data-query-refresh-all": {
      if (!context.selectionEditable || context.sheetProtected) return blocked("spreadsheet.capability.sheetProtected");
      // Refresh All stays live in a workbook with no connection, as Excel's
      // (measured, Excel for Mac 16); a press there refreshes nothing.
      if (predicate === "data-query-refresh-all") return { enabled: true, reasonKey: null };
      const available = predicate === "data-query-create" ? context.dataQueryCreateAvailable
        : predicate === "data-query-refresh" ? context.dataQueryRefreshAvailable : context.dataQueryRefreshAllAvailable;
      return available === true ? { enabled: true, reasonKey: null }
        : blocked(predicate === "data-query-create" ? "spreadsheet.capability.commandUnavailable" : "spreadsheet.capability.noQueries");
    }
    case "remove-duplicates-editable":
      if (!context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.activeCellSelectionSupported !== false &&
        (context.removeDuplicatesAvailable ?? context.hasMultiSelection)
        ? { enabled: true, reasonKey: null }
        : blocked(context.removeDuplicatesUnavailableReason ?? "spreadsheet.capability.multiSelectionRequired");
    case "multi-selection-editable":
      if (!context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.hasMultiSelection
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.multiSelectionRequired");
    case "no-verified-authority":
      return blocked(fallbackReason ?? "spreadsheet.capability.unsupported");
    // Excel groups and ungroups the selection's own lines, one cell included
    // (Group asks Rows or Columns); availability comes from the level math.
    // A sheet whose outline the import could not read exactly keeps every
    // outline edit off: the native writer would refuse the whole save.
    case "outline-group-columns":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.outlineNativeBlocked === true)
        return blocked("spreadsheet.capability.outlineNativeBlocked");
      return context.outlineGroupColumnsAvailable === false
        ? blocked(
            fallbackReason ?? "spreadsheet.capability.outlineGroupUnavailable",
          )
        : { enabled: true, reasonKey: null };
    case "outline-group-rows":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.outlineNativeBlocked === true)
        return blocked("spreadsheet.capability.outlineNativeBlocked");
      return context.outlineGroupRowsAvailable === false
        ? blocked(
            fallbackReason ?? "spreadsheet.capability.outlineGroupUnavailable",
          )
        : { enabled: true, reasonKey: null };
    case "outline-ungroup-columns":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.outlineNativeBlocked === true)
        return blocked("spreadsheet.capability.outlineNativeBlocked");
      return context.outlineUngroupColumnsAvailable === false
        ? blocked(
            fallbackReason ??
              "spreadsheet.capability.outlineUngroupUnavailable",
          )
        : { enabled: true, reasonKey: null };
    case "outline-ungroup-rows":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.outlineNativeBlocked === true)
        return blocked("spreadsheet.capability.outlineNativeBlocked");
      return context.outlineUngroupRowsAvailable === false
        ? blocked(
            fallbackReason ??
              "spreadsheet.capability.outlineUngroupUnavailable",
          )
        : { enabled: true, reasonKey: null };
    case "page-setup-launchable":
      return context.pageSetupLaunchable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.commandUnavailable",
          );
    case "page-setup-editable":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.pageSetupEditable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.commandUnavailable",
          );
    case "page-setup-clearable":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.pageSetupEditable === true &&
        context.pageSetupClearAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.commandUnavailable",
          );
    case "page-margin-applicable":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.pageSetupEditable === true &&
        context.pageMarginApplyAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.commandUnavailable",
          );
    case "print-area-settable":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.pageSetupEditable !== true)
        return blocked(
          fallbackReason ?? "spreadsheet.capability.commandUnavailable",
        );
      return context.printAreaSetAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.printAreaAlreadySelection");
    case "print-area-clearable":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.pageSetupEditable !== true)
        return blocked(
          fallbackReason ?? "spreadsheet.capability.commandUnavailable",
        );
      // Excel greys Clear Print Area while the sheet has none; the reason
      // names that state instead of the menu-level "nothing to run" line.
      return context.printAreaClearAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.printAreaNone");
    case "print-area-addable":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.pageSetupEditable !== true)
        return blocked(
          fallbackReason ?? "spreadsheet.capability.commandUnavailable",
        );
      return context.printAreaAddAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.printAreaAddOutsideRequired");
    case "page-break-insertable":
    case "page-break-removable":
    case "page-break-resettable": {
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.pageSetupEditable !== true)
        return blocked(
          fallbackReason ?? "spreadsheet.capability.commandUnavailable",
        );
      const available =
        predicate === "page-break-insertable"
          ? context.pageBreakInsertAvailable
          : predicate === "page-break-removable"
            ? context.pageBreakRemoveAvailable
            : context.pageBreakResetAvailable;
      if (available === true) return { enabled: true, reasonKey: null };
      // Remove / Reset are greyed for want of a break, not for want of a
      // command, so each names the missing break rather than the generic line.
      return blocked(
        predicate === "page-break-removable"
          ? "spreadsheet.capability.pageBreakNoneAtSelection"
          : predicate === "page-break-resettable"
            ? "spreadsheet.capability.pageBreaksNone"
            : "spreadsheet.capability.pageBreakInsertUnavailable",
      );
    }
    case "saved-window":
      return context.canOpenWorkbookWindow
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.savedWorkbookRequired");
    case "sheet-view-switch":
      // A remembered "just me / everyone" choice is reset from this menu, so
      // the menu has something to offer even before any view exists.
      return (context.namedSheetViewSwitchableCount ?? 0) > 0 ||
        context.sheetViewMode !== "default" ||
        context.sheetViewScopeRemembered === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            sheetViewBlockedReason(
              context,
              "spreadsheet.capability.namedViewsUnavailable",
            ),
          );
    case "sheet-view-exit":
      return context.sheetViewMode && context.sheetViewMode !== "default"
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.sheetViewRequired");
    case "sheet-view-keep":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      // Excel's Keep is live whenever a temporary view is active: it keeps the
      // view at once under the next default name ("View1"), and the name is
      // changed in the Sheet View box afterwards. It is also the box's own
      // commit, so an edited, unsent name keeps it live in a named view.
      return context.sheetViewMode === "temporary" ||
        context.sheetViewNameDraftPending === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            sheetViewBlockedReason(
              context,
              "spreadsheet.capability.temporarySheetViewRequired",
            ),
          );
    case "sheet-view-new":
      if (!context.canOpenWorkbookWindow)
        return blocked("spreadsheet.capability.sheetViewSavedWorkbookRequired");
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.sheetViewCapacityAvailable === false)
        return blocked("spreadsheet.capability.namedViewsUnavailable");
      return context.sheetViewAuthoringAvailable
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.namedViewsUnavailable");
    case "sheet-view-options":
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      return (context.namedSheetViewCount ?? 0) > 0
        ? { enabled: true, reasonKey: null }
        : blocked(
            sheetViewBlockedReason(
              context,
              "spreadsheet.capability.namedViewsUnavailable",
            ),
          );
    case "version-history":
      return context.versionHistoryAvailable
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.savedWorkbookRequired");
    // Show Changes only reads: a protected sheet or a read-only member still
    // sees who changed what, as in Excel.
    case "change-history":
      return context.changeHistoryAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.savedWorkbookRequired");
    case "workbook-protection-launchable":
      return { enabled: true, reasonKey: null };
    case "workbook-protection-applicable":
      if (context.workbookProtectionNativeBlocked) {
        return blocked("spreadsheet.capability.workbookProtectionUnsupported");
      }
      return context.workbookProtectionApplyAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ??
              "spreadsheet.capability.workbookProtectionUnavailable",
          );
    case "workbook-link-inventory":
      return context.workbookLinksAvailable
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.noWorkbookLinks",
          );
    case "worksheet-structure-editable": {
      // A menu of cell/row/column structure items: live while any item is.
      const lines = context.structuralLineEdits;
      const anyItem = context.sheetProtected
        ? Boolean(
            lines &&
              (lines.insertRows ||
                lines.insertColumns ||
                lines.deleteRows ||
                lines.deleteColumns),
          )
        : context.selectionEditable;
      if (!anyItem) return blocked("spreadsheet.capability.sheetProtected");
      if (!context.hasSelection)
        return blocked("spreadsheet.capability.selectionRequired");
      return context.homeStructureControlsAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.commandUnavailable");
    }
    case "hyperlink-insertable":
      if (!context.hasSelection)
        return blocked("spreadsheet.capability.selectionRequired");
      return context.sheetProtected &&
        (!allows("insertHyperlinks") || !context.selectionEditable)
        ? blocked("spreadsheet.capability.sheetProtected")
        : { enabled: true, reasonKey: null };
    case "object-insertable":
    case "pivot-insertable":
      // Excel's Edit objects / Use PivotTable & PivotChart, never the active
      // cell's lock: objects float over the grid and a report creates its own
      // output. Unprotected, the established cell-based reading.
      if (context.sheetProtected) {
        if (!allows(predicate === "pivot-insertable" ? "pivotTables" : "editObjects"))
          return blocked("spreadsheet.capability.sheetProtected");
      } else if (!context.selectionEditable) {
        return blocked("spreadsheet.capability.sheetProtected");
      }
      return context.hasSelection
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.selectionRequired");
    case "worksheet-insert-rows-editable":
    case "worksheet-insert-columns-editable":
    case "worksheet-delete-rows-editable":
    case "worksheet-delete-columns-editable": {
      // Excel's Insert/Delete rows/columns permissions: inserting needs only
      // the permission; deleting also needs every cell of the lines unlocked.
      const key =
        predicate === "worksheet-insert-rows-editable"
          ? "insertRows"
          : predicate === "worksheet-insert-columns-editable"
            ? "insertColumns"
            : predicate === "worksheet-delete-rows-editable"
              ? "deleteRows"
              : "deleteColumns";
      const permitted = context.sheetProtected
        ? context.structuralLineEdits?.[key] === true
        : context.selectionEditable;
      if (!permitted) return blocked("spreadsheet.capability.sheetProtected");
      if (!context.hasSelection)
        return blocked("spreadsheet.capability.selectionRequired");
      return context.homeStructureControlsAvailable === true
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.commandUnavailable");
    }
    case "selection":
      return context.hasSelection
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.selectionRequired");
    // A command on the CELLS (Format Cells, AutoSum): with a picture or a
    // shape selected the cells are not what the user is looking at, so it
    // stands down like the Home font and number commands instead of acting
    // on the cell that was active before (live audit B424).
    case "cell-selection":
      if (context.homeObjectSelected === true)
        return blocked("spreadsheet.capability.objectSelected");
      return context.hasSelection
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.selectionRequired");
    case "cell-selection-editable":
      if (context.homeObjectSelected === true)
        return blocked("spreadsheet.capability.objectSelected");
      if (!context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.hasSelection
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.selectionRequired");
    case "filter-editable":
      // Excel "Protect a worksheet": an AutoFilter can never be applied or
      // removed on a protected sheet, whatever Use AutoFilter says (that
      // permission only opens the criteria of an existing one).
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.hasSelection
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.selectionRequired");
    case "selection-display-text":
      return context.hasSelection &&
        context.activeCellSelectionSupported &&
        context.activeCellHasDisplayText
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ??
              "spreadsheet.capability.immersiveReaderUnavailable",
          );
    case "selection-editable":
      if (!context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.hasSelection
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.selectionRequired");
    case "selection-has-merge-editable":
      if (context.homeObjectSelected === true)
        return blocked("spreadsheet.capability.objectSelected");
      if (!context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      return context.hasSelection && context.selectionHasMerge === true
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ?? "spreadsheet.capability.selectionRequired",
          );
    case "selection-can-lock":
    case "selection-can-unlock": {
      if (context.sheetProtected)
        return blocked("spreadsheet.capability.sheetProtected");
      const lockState =
        context.selectionLockState === undefined
          ? context.selectionLockStateAvailable
            ? "mixed"
            : null
          : context.selectionLockState;
      if (lockState === null)
        return blocked("spreadsheet.capability.selectionRequired");
      if (lockState === "mixed") return { enabled: true, reasonKey: null };
      if (predicate === "selection-can-lock") {
        return lockState
          ? blocked("spreadsheet.capability.selectionAlreadyLocked")
          : { enabled: true, reasonKey: null };
      }
      return lockState
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.selectionAlreadyUnlocked");
    }
    case "sort-authorized":
      if (!(context.sortAuthorized ?? context.selectionEditable))
        return blocked("spreadsheet.capability.sheetProtected");
      return context.hasSelection
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.selectionRequired");
    case "share-web":
    case "share-clipboard":
      if (context.shareOperationBusy)
        return blocked("spreadsheet.capability.shareOperationInProgress");
      if (!context.shareAuthorityAvailable)
        return blocked("spreadsheet.capability.shareWorkbookRequired");
      if (predicate === "share-web" && !context.webShareSupported)
        return blocked("spreadsheet.capability.shareUnavailable");
      if (predicate === "share-clipboard" && !context.clipboardSupported)
        return blocked("spreadsheet.capability.clipboardUnavailable");
      return { enabled: true, reasonKey: null };
    case "sheet-editable":
      return context.sheetProtected
        ? blocked("spreadsheet.capability.sheetProtected")
        : { enabled: true, reasonKey: null };
    case "sheet-protection-configured":
      return context.sheetProtectionConfigured
        ? { enabled: true, reasonKey: null }
        : blocked(
            fallbackReason ??
              "spreadsheet.capability.truePauseProtectionUnavailable",
          );
    case "sheet-protected":
      return context.sheetProtected
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.commandUnavailable");
    case "table-create":
      if (!context.selectionEditable)
        return blocked("spreadsheet.capability.sheetProtected");
      if (context.activeTable)
        return blocked("spreadsheet.capability.tableAlreadyActive");
      return context.tableCapacityAvailable
        ? { enabled: true, reasonKey: null }
        : blocked("spreadsheet.capability.tableLimit");
    case "workbook-sheet-structure-editable":
      // Insert Sheet / Delete Sheet. Worksheet protection does not reach them:
      // "Protect a worksheet" lists what a protected sheet stops, and adding or
      // removing a worksheet is not on it — only Protect Workbook ▸ Structure
      // stops that (Microsoft Support, "Protect a worksheet" and "Lock or
      // unlock specific areas of a protected worksheet"). The two protections
      // are independent, so a menu that carries both kinds of item stays open
      // while either kind is still live, and each item gives its own reason.
      if (context.workbookStructureLocked !== true) {
        return { enabled: true, reasonKey: null };
      }
      return predicateDecision(
        "worksheet-structure-editable",
        context,
        "spreadsheet.capability.workbookStructureProtected",
      );
    case "menu-children":
      return blocked("spreadsheet.capability.commandUnavailable");
  }
}

/**
 * The catalog's own answer for a parameterized gallery leaf.
 *
 * Conditional-format presets and table styles are deliberately NOT fixed
 * catalog leaves (see the Home vertical-slice note above): their entries are
 * values, not controls, so they carry no controlId and cannot be resolved
 * through `definitionByControl`. They still write into the selection, and the
 * writers behind them already refuse a protected one — so without this the
 * gallery is the one place where the same question gets a second, hand-written
 * answer, and an entry sits enabled next to greyed siblings only to answer a
 * click with a toast.
 *
 * `selection-editable` is the predicate its registered neighbours use
 * (Clear Rules from selection, Cell Styles), so the reason strings match too.
 */
export function spreadsheetVisibleGalleryDecision(
  context: SpreadsheetVisibleCommandContext,
): {
  enabled: boolean;
  reasonKey: SpreadsheetVisibleCommandReason | null;
} {
  return predicateDecision("selection-editable", context, null);
}

/**
 * The same, for the AutoFilter header menu's sort aliases. They are the Data
 * tab's own Sort A to Z / Z to A acting on the clicked column, but they live in
 * a popover with its own `XLSCTL.SHEET.AUTOFILTER.*` ids that the ribbon catalog
 * does not own. Excel gates sorting on a protected sheet behind the Allow Sort
 * permission, which is exactly what `sort-authorized` reads.
 */
export function spreadsheetVisibleSortDecision(
  context: SpreadsheetVisibleCommandContext,
): {
  enabled: boolean;
  reasonKey: SpreadsheetVisibleCommandReason | null;
} {
  return predicateDecision("sort-authorized", context, null);
}








export function spreadsheetVisibleControlDecision(
  controlId: SpreadsheetVisibleControlId,
  context: SpreadsheetVisibleCommandContext,
): {
  enabled: boolean;
  reasonKey: SpreadsheetVisibleCommandReason | null;
} {
  const definition = definitionByControl.get(controlId);
  if (!definition || definition.command === "unsupported") {
    return { enabled: false, reasonKey: "spreadsheet.capability.commandUnavailable" };
  }
  const predicate = predicateDecision(definition.predicate, context, definition.reasonKey);
  // The ribbon's writer question too (resolveSemanticCapability): a headless
  // run (the AI's commands, the Copilot) on an imported workbook must not
  // make a change the in-place save cannot write (import audit S4, §6).
  if (
    !predicate.enabled ||
    definition.undoPolicy !== "single-transaction" ||
    spreadsheetVisibleIoWriterReady(
      SPREADSHEET_VISIBLE_IO_CONTRACTS[definition.ioProfile],
      context.nativeWriteRoute,
    )
  ) {
    return predicate;
  }
  return { enabled: false, reasonKey: "spreadsheet.capability.writerUnavailable" };
}

function resolveSemanticCapability(
  capability: SpreadsheetVisibleSemanticCapability,
  context: SpreadsheetVisibleCommandContext,
): { enabled: boolean; reasonKey: string | null } {
  const definition = definitionByControl.get(capability.capability.controlId);
  if (!definition) {
    return {
      enabled: false,
      reasonKey: "spreadsheet.capability.commandUnavailable",
    };
  }
  const predicate = predicateDecision(
    definition.predicate,
    context,
    definition.reasonKey,
  );
  if (!predicate.enabled) return predicate;
  if (capability.capability.command.undoPolicy !== "single-transaction") {
    return predicate;
  }
  return spreadsheetVisibleIoWriterReady(
    capability.capability,
    context.nativeWriteRoute,
  )
    ? predicate
    : {
        enabled: false,
        reasonKey: "spreadsheet.capability.writerUnavailable",
      };
}

export interface ResolvedSpreadsheetVisibleControl {
  readonly visible: boolean;
  readonly enabled: boolean;
  readonly reasonKey: SpreadsheetVisibleCommandReason | null;
  readonly featureId: string;
  readonly controlId: string;
}

export interface SpreadsheetVisibleCommandAuthority {
  readonly capabilityCatalog: typeof SPREADSHEET_VISIBLE_COMMAND_CAPABILITY_CATALOG;
  resolve(controlId: string): ResolvedSpreadsheetVisibleControl;
  /** Read-only UI projection; never publishes execution authority. */
  resolveForRender(controlId: string, context: SpreadsheetVisibleCommandContext): ResolvedSpreadsheetVisibleControl;
  execute(controlId: string): Promise<boolean>;
}

export function createSpreadsheetVisibleCommandLineage(
  instanceId: string,
): FeatureCommandRegistryLineage {
  return createFeatureCommandRegistryLineage(
    "spreadsheet",
    `spreadsheet.visible.${instanceId}`,
  );
}

/**
 * Build one authenticated command authority for a live Spreadsheet window.
 * The same predicate is consulted for rendering and again immediately before
 * executing, so stale enabled UI cannot bypass a changed protection/selection
 * state. Exact menu roots derive from exact child decisions.
 */
export function createSpreadsheetVisibleCommandAuthority(config: {
  readonly lineage: FeatureCommandRegistryLineage;
  readonly getContext: () => SpreadsheetVisibleCommandContext;
  readonly getHandlers: () => SpreadsheetVisibleCommandHandlers;
}): SpreadsheetVisibleCommandAuthority {
  const expectedFeatureIds = semanticCatalog.map((capability) => capability.id);
  const suppliedFeatureIds = Object.keys(config.getHandlers()).sort();
  if (
    suppliedFeatureIds.length !== expectedFeatureIds.length ||
    suppliedFeatureIds.some(
      (featureId, index) => featureId !== expectedFeatureIds[index],
    )
  ) {
    throw new Error("Incomplete spreadsheet visible command handlers");
  }

  const bindings = Object.freeze(
    semanticCatalog.map(({ id, capability }) =>
      Object.freeze({
        featureId: id,
        action: capability.command.key ?? commandKey(id),
        handlerRef: capability.command.handlerRef ?? handlerRef(id),
        execute: () => {
          const handlers = config.getHandlers();
          const currentFeatureIds = Object.keys(handlers).sort();
          if (
            currentFeatureIds.length !== expectedFeatureIds.length ||
            currentFeatureIds.some(
              (featureId, index) => featureId !== expectedFeatureIds[index],
            )
          ) {
            return false;
          }
          const handler = handlers[id];
          return typeof handler === "function" ? handler() : false;
        },
      }),
    ),
  );
  const catalogEntries = Object.freeze(
    semanticCatalog.map(({ id, capability }) =>
      Object.freeze({
        featureId: id,
        semanticContract: spreadsheetFeatureCommandSemanticContract(capability),
        action: capability.command.key,
        handlerRef: capability.command.handlerRef,
      }),
    ),
  );
  const bindingContract = createFeatureCommandBindingContract({
    lineage: config.lineage,
    catalogEntries,
    bindings,
    resolveCapability: resolveSemanticCapability,
  });
  const registry = createFeatureCommandRegistry<
    SpreadsheetVisibleSemanticCapability,
    SpreadsheetVisibleCommandContext,
    Record<string, never>
  >({
    editor: "spreadsheet",
    catalog: semanticCatalog,
    lineage: config.lineage,
    bindingContract,
    describeBindingContract: ({ capability }) =>
      spreadsheetFeatureCommandSemanticContract(capability),
    find: (featureId) => semanticByFeature.get(featureId) ?? null,
    resolveCapability: resolveSemanticCapability,
    bindings,
  });
  const capabilityCatalog = verifySpreadsheetFeatureCapabilityCatalogBindings(
    SPREADSHEET_VISIBLE_COMMAND_CAPABILITY_CATALOG,
    registry.bindingManifest,
    {
      expectedLineage: config.lineage,
      expectedRegistry: registry,
      expectedBindingContract: bindingContract,
    },
  );

  const resolveForRender = (
    controlId: string,
    context: SpreadsheetVisibleCommandContext,
  ): ResolvedSpreadsheetVisibleControl => {
    const definition = definitionByControl.get(controlId);
    if (!definition) {
      return {
        visible: false,
        enabled: false,
        reasonKey: "spreadsheet.capability.commandUnavailable",
        featureId: controlId,
        controlId,
      };
    }
    if (
      SPREADSHEET_SHAPE_FORMAT_INTERNAL_CONTROL_ID_SET.has(controlId) ||
      SPREADSHEET_PICTURE_FORMAT_INTERNAL_CONTROL_ID_SET.has(controlId)
    ) {
      return {
        visible: false,
        enabled: false,
        reasonKey: "spreadsheet.capability.commandUnavailable",
        featureId: definition.featureId,
        controlId,
      };
    }
    if (definition.command === "disclosure") {
      const children = (definition.children ?? []).map((child) => resolveForRender(child, context));
      const firstVisible = children.find((child) => child.visible);
      const firstEnabled = children.find((child) => child.enabled);
      // A disclosure that declares its own predicate answers for itself (the
      // Merge arrow stays openable on a writable sheet so its greyed items
      // can be read with their reasons); the default follows its children.
      const own =
        definition.predicate === "menu-children"
          ? null
          : predicateDecision(
              definition.predicate,
              context,
              definition.reasonKey,
            );
      return {
        visible: children.some((child) => child.visible),
        enabled: own ? own.enabled : Boolean(firstEnabled),
        reasonKey: own
          ? own.reasonKey
          : firstEnabled
            ? null
            : (firstVisible?.reasonKey ?? definition.reasonKey),
        featureId: definition.featureId,
        controlId,
      };
    }
    if (definition.command === "unsupported") {
      const decision = predicateDecision(
        definition.predicate,
        context,
        definition.reasonKey,
      );
      return {
        // Omitted, not drawn inert. Most of what lands here runs only inside
        // Microsoft's own environment — Forms, Data from Power BI, Publish to
        // Web and the linked data types behind Data Type Icons are all tenant
        // services, not spreadsheet behaviour — and MRAI is its own editor
        // rather than a copy of that ribbon, so it does not carry a permanent
        // dead row advertising somebody else's service. The Atlas still
        // records that Excel for the web offers them, and each keeps a reason
        // that says plainly what we have not built, so the day one is built or
        // wanted on screen it is one decision away. Context-disabled
        // *supported* commands remain visible below with their reason.
        visible: false,
        enabled: false,
        reasonKey: decision.reasonKey,
        featureId: definition.featureId,
        controlId,
      };
    }

    const capability = resolveSpreadsheetFeatureCapability(
      controlId,
      capabilityCatalog,
    );
    const decision = registry.resolve(
      definition.featureId,
      context,
      Object.freeze({}),
    );
    const executable = isSpreadsheetFeatureExecutable(
      capability,
      capabilityCatalog,
    );
    if (executable && decision.enabled) {
      // Development only: an enabled control drawn from a registry its
      // lineage has already replaced would be silent (b24a8ad5c2).
      checkEnabledFeatureCommandRegistry(
        registry,
        definition.featureId,
        "rendered",
      );
    }
    return {
      visible: true,
      enabled: executable && decision.enabled,
      reasonKey: executable
        ? (decision.reasonKey as SpreadsheetVisibleCommandReason | null)
        : "spreadsheet.capability.commandUnavailable",
      featureId: definition.featureId,
      controlId,
    };
  };

  const resolve = (controlId: string) => resolveForRender(controlId, config.getContext());

  return Object.freeze({
    capabilityCatalog,
    resolve,
    resolveForRender,
    async execute(controlId: string): Promise<boolean> {
      const state = resolve(controlId);
      if (!state.visible || !state.enabled) return false;
      checkEnabledFeatureCommandRegistry(registry, state.featureId, "invoked");
      const context = config.getContext();
      return (
        (await registry.execute(
          state.featureId,
          context,
          Object.freeze({}),
        )) === "executed"
      );
    },
  });
}
