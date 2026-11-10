"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The table Layout contextual tab of the presentation Ribbon (Table, Rows and
 * Columns, Merge, Cell Size, Alignment, Table Size, Arrange), drawn from the
 * window bindings it reads.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the tab
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import {
  OfficeHomeRibbonGroup,
  OfficeHomeCommandRow,
  OfficeHomeMenuTrigger,
  OfficeHomeCommand,
  OfficeHomeCommandStack,
  OfficeHomeSplitLayout,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import {
  DropdownMenu,
  DropdownMenuTrigger,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuCheckboxItem,
  DropdownMenuSeparator,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import { tableLogicalWidth, MAX_TABLE_SPLIT_ROWS, MAX_TABLE_SPLIT_COLUMNS } from "@mrnote/core/web/lib/presentation-table-merge.ts";
import { OfficeDialogShell } from "@mrnote/core/web/components/office-ribbon/OfficeDialogShell.tsx";
import { officeDialogParts as dialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import {
  TABLE_LAYOUT_CELL_ALIGNMENTS,
  tableCellMarginPresetPatch,
  type CellStyleInput,
  type TableCellFontFamily,
} from "@mrnote/core/web/lib/presentation-table-style.ts";
import {
  IconParagraphAlign,
  IconBringForward,
  IconBringToFront,
  IconSendBackward,
  IconSendToBack,
  IconAlignObjects,
  IconDistributeObjects,
  IconRotateObjects,
  IconFlipVertical,
  IconFlipHorizontal,
} from "@mrnote/core/web/components/notebook/contents/PresentationRibbonIcons.tsx";
import { VERTICAL_ANCHORS, persistedVerticalAnchor } from "@mrnote/core/web/lib/presentation-vertical-anchor.ts";
import {
  AlignVerticalJustifyStart,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  Trash2,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { PresentationTableCellChoiceMenu, PresentationTableCellMarginMenu, TABLE_CELL_FONT_SIZE_OPTIONS } from "./PresentationTableCellMenus";
import { RibbonFontFamilyPicker } from "@mrnote/core/web/components/office-ribbon/RibbonFontFamilyPicker.tsx";
import { RibbonValueCombobox } from "@mrnote/core/web/components/office-ribbon/RibbonValueCombobox.tsx";
import { parseOfficeFontSizeInput } from "@mrnote/core/web/components/office-ribbon/officeFontSizeCatalog.ts";
import {
  TABLE_COLUMN_AUTOFIT_MODES,
  tableAxisPxToCm,
  MIN_TABLE_AXIS_PX,
  MAX_TABLE_AXIS_PX,
  tableFrameSizeCm,
  MIN_TABLE_FRAME_CM,
  MAX_TABLE_FRAME_CM,
  type TableAxisSizeSelectionState,
  type TableColumnAutoFitMode,
  type TableFrameAxis,
} from "@mrnote/core/web/lib/presentation-table-size.ts";
import { OfficeRibbonNumberField } from "@mrnote/core/web/components/office-ribbon/OfficeRibbonNumberField.tsx";
import { presentationMeasurementCmDraft, type PresentationMeasurementUnit } from "@mrnote/core/web/lib/presentation-measurement.ts";
import { hasLayoutGeometry } from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import { PRESENTATION_RIBBON_LABELED_BUTTON_STYLE } from "../../../../lib/presentation/window-chrome-styles";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { TableAxisSelectionRange, TableSelectionScope } from "@mrnote/core/web/lib/presentation-table-selection.ts";
import type { Dispatch, RefObject, SetStateAction } from "react";
import type { ContextualTableAxis, TableCellMarginMenuValue } from "./PresentationTableCellMenus";
import type { PresentationCellTextElement } from "@mrnote/core/web/components/notebook/contents/PresentationCellRichTextEditor.tsx";
import type { PresentationBlock, PresentationTableBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { AlignEdge, DistributeAxis } from "@mrnote/core/web/lib/presentation-align.ts";
import type { QuarterTurnDirection } from "@mrnote/core/web/lib/presentation-rotate-quarter.ts";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationTableLayoutRibbonTabCtl {
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  sizeUnit: PresentationMeasurementUnit;
  sizeUnitLabel: string;
  sizeStep: number;
  sizeBound: (centimetres: number) => number;
  sizeDisplayFromEmu: (emu: unknown, cm: string) => string;
  tableCellFocus: { blockId: string; row: number; col: number; } | null;
  tableSelectionScope: TableSelectionScope;
  tableCellFontSizeDraft: string;
  setTableCellFontSizeDraft: Dispatch<SetStateAction<string>>;
  setTableCellFontSizeFocused: Dispatch<SetStateAction<boolean>>;
  setTableCellFontSizeDirty: Dispatch<SetStateAction<boolean>>;
  setTableCellMarginsDialogOpen: Dispatch<SetStateAction<boolean>>;
  tableCellMarginsTriggerRef: RefObject<HTMLButtonElement | null>;
  tableAxisSizeDrafts: Record<ContextualTableAxis, string>;
  tableSplitPanelOpen: boolean;
  setTableSplitPanelOpen: Dispatch<SetStateAction<boolean>>;
  tableSplitRowsDraft: string;
  setTableSplitRowsDraft: Dispatch<SetStateAction<string>>;
  tableSplitColumnsDraft: string;
  setTableSplitColumnsDraft: Dispatch<SetStateAction<string>>;
  tableSplitTriggerRef: RefObject<HTMLButtonElement | null>;
  tableSplitRowsInputRef: RefObject<HTMLInputElement | null>;
  tableFrameAspectLocked: boolean;
  setTableFrameAspectLocked: Dispatch<SetStateAction<boolean>>;
  tableGridlinesVisible: boolean;
  setTableGridlinesVisible: Dispatch<SetStateAction<boolean>>;
  selectWholeTableCellRef: RefObject<boolean>;
  tableScopeChosenFromMenuRef: RefObject<boolean>;
  tableCellInputRefs: RefObject<Map<string, PresentationCellTextElement>>;
  restoringTableScopeFocusRef: RefObject<boolean>;
  alignToSlide: boolean;
  writeAlignToSlide: (next: boolean) => void;
  contextualTableBlock: PresentationTableBlock;
  contextualTableSelectionIsWritable: boolean;
  /** The table's frame may change (Lock fixes it): rows and columns in or out, the table's and its cells' sizes. */
  contextualTableFrameIsWritable: boolean;
  contextualTableFocus: { blockId: string; row: number; col: number; } | null;
  contextualTableCellFormatState: { fill: Readonly<{ mixed: boolean; value: string | null; }> | null; bevel: Readonly<{ mixed: boolean; value: "none" | "relaxedInset" | "circle" | "slope" | "cross" | "angle" | "softRound" | "convex" | "coolSlant" | "divot" | "riblet" | "hardEdge" | "artDeco"; }> | null; color: Readonly<{ mixed: boolean; value: string | null; }> | null; highlight: Readonly<{ mixed: boolean; value: string | null; }> | null; fontFamily: Readonly<{ mixed: boolean; value: TableCellFontFamily | null; }> | null; fontSize: Readonly<{ mixed: boolean; value: number | null; }> | null; bold: Readonly<{ mixed: boolean; value: boolean; }> | null; italic: Readonly<{ mixed: boolean; value: boolean; }> | null; underline: Readonly<{ mixed: boolean; value: boolean; }> | null; strike: Readonly<{ mixed: boolean; value: boolean; }> | null; superscript: Readonly<{ mixed: boolean; value: boolean; }> | null; subscript: Readonly<{ mixed: boolean; value: boolean; }> | null; align: Readonly<{ mixed: boolean; value: "left" | "center" | "right" | "justify" | "distributed" | "thaiDistributed" | "justifyLow"; }> | null; verticalAlign: Readonly<{ mixed: boolean; value: "top" | "middle" | "bottom"; }> | null; textDirection: Readonly<{ mixed: boolean; value: "horizontal" | "rotate-90" | "rotate-270" | "stacked" | "eastAsianVertical" | "mongolianVertical" | "stackedRightToLeft"; }> | null; marginLeft: Readonly<{ mixed: boolean; value: number | null; }> | null; marginRight: Readonly<{ mixed: boolean; value: number | null; }> | null; marginTop: Readonly<{ mixed: boolean; value: number | null; }> | null; marginBottom: Readonly<{ mixed: boolean; value: number | null; }> | null; };
  contextualTableCellFontFamilyValue: TableCellFontFamily | null;
  contextualTableCellFontFamilyLabel: string;
  contextualTableCellFontSections: { id: string; options: ({ id: "Aptos" | "Calibri" | "Arial" | "Cambria" | "Times New Roman" | "Georgia" | "Courier New" | "Consolas" | "DengXian" | "Microsoft YaHei" | "SimSun" | "PingFang SC" | "Microsoft JhengHei" | "Yu Gothic" | "Malgun Gothic"; label: "Aptos" | "Calibri" | "Arial" | "Cambria" | "Times New Roman" | "Georgia" | "Courier New" | "Consolas" | "DengXian (等线)" | "Microsoft YaHei (微软雅黑)" | "SimSun (宋体)" | "PingFang SC (苹方)" | "Microsoft JhengHei (微軟正黑體)" | "Yu Gothic (游ゴシック)" | "Malgun Gothic (맑은 고딕)"; value: "Aptos" | "Calibri" | "Arial" | "Cambria" | "Times New Roman" | "Georgia" | "Courier New" | "Consolas" | "DengXian" | "Microsoft YaHei" | "SimSun" | "PingFang SC" | "Microsoft JhengHei" | "Yu Gothic" | "Malgun Gothic"; previewFamily: "Aptos, Calibri, 'Segoe UI', sans-serif" | "Calibri, 'Segoe UI', Candara, sans-serif" | "Arial, 'Helvetica Neue', Helvetica, sans-serif" | "Cambria, Georgia, 'Times New Roman', serif" | "'Times New Roman', Times, serif" | "Georgia, Cambria, 'Times New Roman', serif" | "'Courier New', Courier, monospace" | "Consolas, 'SFMono-Regular', Menlo, monospace" | "DengXian, 'Microsoft YaHei', 'PingFang SC', sans-serif" | "'Microsoft YaHei', 'PingFang SC', sans-serif" | "SimSun, 'Songti SC', serif" | "'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', sans-serif" | "'Microsoft JhengHei', 'PingFang TC', sans-serif" | "'Yu Gothic', 'Hiragino Kaku Gothic ProN', sans-serif" | "'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif"; selected: boolean; testId: string; } | { id: string; label: string; value: null; selected: boolean; testId: string; })[]; }[];
  contextualTableCellMarginMenuValue: TableCellMarginMenuValue;
  tableCellMarginMenuLabels: Readonly<Record<TableCellMarginMenuValue, string> & { title: string; }>;
  contextualTableRowStructureRange: TableAxisSelectionRange | null;
  contextualTableColumnStructureRange: TableAxisSelectionRange | null;
  contextualTableRowSpanRange: TableAxisSelectionRange | null;
  contextualTableColumnSpanRange: TableAxisSelectionRange | null;
  contextualTableColumnSizeState: TableAxisSizeSelectionState | null;
  contextualTableRowSizeState: TableAxisSizeSelectionState | null;
  formatPaneShowing: boolean;
  toggleFormatPane: () => void;
  formatPaneInFront: boolean;
  deleteBlock: (blockId: string) => void;
  selectedGeometryBlocks: PresentationBlock[];
  alignSelectedBlocks: (edge: AlignEdge) => void;
  distributeSelectedBlocks: (axis: DistributeAxis) => void;
  flipSelectedBlocks: (axis: "flipH" | "flipV") => void;
  canFlipSelection: boolean;
  quarterTurnSelectedBlocks: (direction: QuarterTurnDirection) => void;
  ungroupSelectedBlocks: () => void;
  contextualTableCanUngroup: boolean;
  contextualTableLayerCapability: Readonly<{ canMoveBackward: boolean; canMoveForward: boolean; }>;
  bringBlockToFront: (explicitId?: string, explicitOnly?: boolean) => void;
  sendBlockToBack: (explicitId?: string, explicitOnly?: boolean) => void;
  bringBlockForward: (explicitId?: string, explicitOnly?: boolean, sourceFreeOnly?: boolean, explicitIds?: readonly string[]) => void;
  sendBlockBackward: (explicitId?: string, explicitOnly?: boolean, sourceFreeOnly?: boolean, explicitIds?: readonly string[]) => void;
  insertTableRowAtFocus: (blockId: string, placement: "before" | "after") => void;
  insertTableColumnAtFocus: (blockId: string, placement: "before" | "after") => void;
  deleteTableRowAtFocus: (blockId: string) => void;
  deleteTableColumnAtFocus: (blockId: string) => void;
  deleteTableAtFocus: (blockId: string) => void;
  selectTableScope: (blockId: string, scope: TableSelectionScope, requestedFocus?: { row: number; col: number; }, extendSelection?: boolean) => void;
  mergeTableSelection: () => void;
  splitTableCellIntoParts: (rowParts: number, columnParts: number) => void;
  tableSelectionCanMerge: boolean;
  tableFocusCanUseSplitCommand: boolean;
  requestTableCellSplit: () => void;
  updateContextualTableCellStyle: (patch: Partial<CellStyleInput>) => void;
  selectContextualTableCellFontFamily: (fontFamily: string | null) => void;
  commitContextualTableCellFontSize: (rawValue: string, options: { allowClear: boolean; restoreFocus: boolean; }) => boolean;
  commitContextualTableAxisSize: (axis: ContextualTableAxis, rawValue: string, options: { allowClear: boolean; restoreFocus: boolean; }) => boolean;
  commitContextualTableFrameSize: (axis: TableFrameAxis, rawValue: string) => boolean;
  distributeTableAxis: (blockId: string, axis: "col" | "row", selectionRange?: TableAxisSelectionRange | null) => void;
  autoFitTableColumns: (blockId: string, mode: TableColumnAutoFitMode) => void;
  tableSplitMaxRows: number;
  tableSplitMaxColumns: number;
  tableSplitRowsValue: number;
  tableSplitColumnsValue: number;
  tableSplitDraftIsValid: boolean;
  tableCellTextDirectionOptions: { value: "horizontal" | "rotate-90" | "rotate-270" | "stacked"; label: string; }[];
}

export function PresentationTableLayoutRibbonTab({ ctl }: { ctl: PresentationTableLayoutRibbonTabCtl }) {
  const {
    t,
    sizeUnit,
    sizeUnitLabel,
    sizeStep,
    sizeBound,
    sizeDisplayFromEmu,
    tableCellFocus,
    tableSelectionScope,
    tableCellFontSizeDraft,
    setTableCellFontSizeDraft,
    setTableCellFontSizeFocused,
    setTableCellFontSizeDirty,
    setTableCellMarginsDialogOpen,
    tableCellMarginsTriggerRef,
    tableAxisSizeDrafts,
    tableSplitPanelOpen,
    setTableSplitPanelOpen,
    tableSplitRowsDraft,
    setTableSplitRowsDraft,
    tableSplitColumnsDraft,
    setTableSplitColumnsDraft,
    tableSplitTriggerRef,
    tableSplitRowsInputRef,
    tableFrameAspectLocked,
    setTableFrameAspectLocked,
    tableGridlinesVisible,
    setTableGridlinesVisible,
    selectWholeTableCellRef,
    tableScopeChosenFromMenuRef,
    tableCellInputRefs,
    restoringTableScopeFocusRef,
    alignToSlide,
    writeAlignToSlide,
    contextualTableBlock,
    contextualTableSelectionIsWritable,
    contextualTableFrameIsWritable,
    contextualTableFocus,
    contextualTableCellFormatState,
    contextualTableCellFontFamilyValue,
    contextualTableCellFontFamilyLabel,
    contextualTableCellFontSections,
    contextualTableCellMarginMenuValue,
    tableCellMarginMenuLabels,
    contextualTableRowStructureRange,
    contextualTableColumnStructureRange,
    contextualTableRowSpanRange,
    contextualTableColumnSpanRange,
    contextualTableColumnSizeState,
    contextualTableRowSizeState,
    formatPaneShowing,
    toggleFormatPane,
    formatPaneInFront,
    deleteBlock,
    selectedGeometryBlocks,
    alignSelectedBlocks,
    distributeSelectedBlocks,
    flipSelectedBlocks,
    canFlipSelection,
    quarterTurnSelectedBlocks,
    ungroupSelectedBlocks,
    contextualTableCanUngroup,
    contextualTableLayerCapability,
    bringBlockToFront,
    sendBlockToBack,
    bringBlockForward,
    sendBlockBackward,
    insertTableRowAtFocus,
    insertTableColumnAtFocus,
    deleteTableRowAtFocus,
    deleteTableColumnAtFocus,
    deleteTableAtFocus,
    selectTableScope,
    mergeTableSelection,
    splitTableCellIntoParts,
    tableSelectionCanMerge,
    tableFocusCanUseSplitCommand,
    requestTableCellSplit,
    updateContextualTableCellStyle,
    selectContextualTableCellFontFamily,
    commitContextualTableCellFontSize,
    commitContextualTableAxisSize,
    commitContextualTableFrameSize,
    distributeTableAxis,
    autoFitTableColumns,
    tableSplitMaxRows,
    tableSplitMaxColumns,
    tableSplitRowsValue,
    tableSplitColumnsValue,
    tableSplitDraftIsValid,
    tableCellTextDirectionOptions,
  } = ctl;
  return (
                  <>
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.tableSelectionGroup}
                      label={t("presentation.tableRibbon.layoutTools")}
                      size="standard"
                      testId="presentation-ribbon-table-selection-group"
                    >
                      <OfficeHomeCommandRow>
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              active={tableSelectionScope !== "cells"}
                              data-control-id="PPTCTL.CONTEXT.TABLE.LAYOUT.SELECT"
                              hasMenuIndicator
                              icon={<OfficeSemanticIcon semanticId="table.select" size={24} />}
                              label={t("toolbar.select")}
                              testId="presentation-ribbon-table-select-menu"
                              variant="large"
                            />
                          </DropdownMenuTrigger>
                          <RibbonDropdownMenuContent
                            align="start"
                            data-testid="presentation-ribbon-table-select-menu-content"
                            // Choosing a scope ends with the caret back in the
                            // table, which is what makes Ctrl+C / Ctrl+V act on
                            // the range just selected. Closing the menu any
                            // other way (Escape, a click outside) keeps Radix's
                            // ordinary return to the trigger.
                            onCloseAutoFocus={(event) => {
                              if (!tableScopeChosenFromMenuRef.current) return;
                              tableScopeChosenFromMenuRef.current = false;
                              event.preventDefault();
                              // Select Cell: the menu has let go of focus now,
                              // so the cell can take it with its whole text
                              // selected (PPT-B123).
                              if (!selectWholeTableCellRef.current) return;
                              selectWholeTableCellRef.current = false;
                              const focus = tableCellFocus;
                              const cell = focus
                                ? tableCellInputRefs.current.get(
                                    `${focus.blockId}:${focus.row}:${focus.col}`,
                                  )
                                : undefined;
                              if (!cell || typeof cell.setSelectionRange !== "function") return;
                              // After the menu's focus scope has unmounted.
                              window.requestAnimationFrame(() => {
                                if (!cell.isConnected) return;
                                if (document.activeElement !== cell) {
                                  restoringTableScopeFocusRef.current = true;
                                  cell.focus({ preventScroll: true });
                                  if (document.activeElement !== cell) {
                                    restoringTableScopeFocusRef.current = false;
                                    return;
                                  }
                                }
                                cell.setSelectionRange(0, cell.value.length);
                              });
                            }}
                          >
                            <DropdownMenuRadioGroup value={tableSelectionScope}>
                              <DropdownMenuRadioItem
                                data-control-id="PPTCTL.CONTEXT.TABLE.LAYOUT.SELECT.CELL"
                                value="cells"
                                onSelect={() => {
                                  tableScopeChosenFromMenuRef.current = true;
                                  selectWholeTableCellRef.current = true;
                                  selectTableScope(
                                    contextualTableBlock.id,
                                    "cells",
                                  );
                                }}
                                data-testid="presentation-ribbon-table-select-cell"
                              >
                                {t("presentation.tableSelectCell")}
                              </DropdownMenuRadioItem>
                              <DropdownMenuRadioItem
                                data-control-id="PPTCTL.CONTEXT.TABLE.LAYOUT.SELECT.ROW"
                                value="row"
                                onSelect={() => {
                                  tableScopeChosenFromMenuRef.current = true;
                                  selectTableScope(
                                    contextualTableBlock.id,
                                    "row",
                                  );
                                }}
                                data-testid="presentation-ribbon-table-select-row"
                              >
                                {t("presentation.tableSelectRow")}
                              </DropdownMenuRadioItem>
                              <DropdownMenuRadioItem
                                data-control-id="PPTCTL.CONTEXT.TABLE.LAYOUT.SELECT.COLUMN"
                                value="column"
                                onSelect={() => {
                                  tableScopeChosenFromMenuRef.current = true;
                                  selectTableScope(
                                    contextualTableBlock.id,
                                    "column",
                                  );
                                }}
                                data-testid="presentation-ribbon-table-select-column"
                              >
                                {t("presentation.tableSelectColumn")}
                              </DropdownMenuRadioItem>
                              <DropdownMenuRadioItem
                                data-control-id="PPTCTL.CONTEXT.TABLE.LAYOUT.SELECT.TABLE"
                                value="table"
                                onSelect={() => {
                                  tableScopeChosenFromMenuRef.current = true;
                                  selectTableScope(
                                    contextualTableBlock.id,
                                    "table",
                                  );
                                }}
                                data-testid="presentation-ribbon-table-select-table"
                              >
                                {t("presentation.tableSelectTable")}
                              </DropdownMenuRadioItem>
                            </DropdownMenuRadioGroup>
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                        {/* PowerPoint's Layout > Table group: Select, then View
                        Gridlines as its own toggle beside it (PPT-A140). An
                        editing aid only: it never dirties the slide. */}
                        <OfficeHomeCommand
                          variant="large"
                          active={tableGridlinesVisible}
                          controlId="PPTCTL.CONTEXT.TABLE.LAYOUT.VIEW_GRIDLINES"
                          icon={<OfficeSemanticIcon semanticId="table.gridlines" size={24} />}
                          label={t("toolbar.tableViewGridlines")}
                          onClick={() => setTableGridlinesVisible((visible) => !visible)}
                          testId="presentation-ribbon-table-gridlines"
                        />
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.tableStructureGroup}
                      label={t("presentation.tableRibbon.rowsAndColumns")}
                      size="wide"
                      testId="presentation-ribbon-table-structure-group"
                    >
                      <OfficeHomeCommandRow>
                        {/* PowerPoint keeps Delete in Rows & Columns, before the inserts. */}
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              data-control-id="PPTCTL.CONTEXT.TABLE.LAYOUT.DELETE"
                              disabled={!contextualTableSelectionIsWritable}
                              hasMenuIndicator
                              icon={<OfficeSemanticIcon semanticId="table.delete" size={24} />}
                              label={t("presentation.contextMenu.delete")}
                              testId="presentation-ribbon-table-delete-menu"
                              variant="large"
                            />
                          </DropdownMenuTrigger>
                          <RibbonDropdownMenuContent align="start" data-testid="presentation-ribbon-table-delete-menu-content">
                            <DropdownMenuItem disabled={!contextualTableFrameIsWritable || contextualTableBlock.rows.length <= (contextualTableRowStructureRange?.count ?? 1)} onSelect={() => deleteTableRowAtFocus(contextualTableBlock.id)} data-control-id="PPTCTL.CONTEXT.TABLE.LAYOUT.DELETE.ROWS" data-testid="presentation-ribbon-table-delete-row">
                              <OfficeSemanticIcon semanticId="table.delete" size={16} />{t("toolbar.deleteRow")}
                            </DropdownMenuItem>
                            <DropdownMenuItem disabled={!contextualTableFrameIsWritable || tableLogicalWidth(contextualTableBlock.rows) <= (contextualTableColumnStructureRange?.count ?? 1)} onSelect={() => deleteTableColumnAtFocus(contextualTableBlock.id)} data-control-id="PPTCTL.CONTEXT.TABLE.LAYOUT.DELETE.COLUMNS" data-testid="presentation-ribbon-table-delete-column">
                              <OfficeSemanticIcon semanticId="table.delete" size={16} />{t("toolbar.deleteColumn")}
                            </DropdownMenuItem>
                            <DropdownMenuItem disabled={!contextualTableSelectionIsWritable} onSelect={() => deleteTableAtFocus(contextualTableBlock.id)} data-control-id="PPTCTL.CONTEXT.TABLE.LAYOUT.DELETE.TABLE" data-testid="presentation-ribbon-table-delete-table">
                              <OfficeSemanticIcon semanticId="table.delete" size={16} />{t("toolbar.deleteTable")}
                            </DropdownMenuItem>
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                          <OfficeHomeCommand
                            variant="large"
                            controlId="PPTCTL.CONTEXT.TABLE.LAYOUT.INSERT_ABOVE"
                            disabled={
                              !contextualTableFrameIsWritable ||
                              contextualTableBlock.rows.length >=
                                MAX_TABLE_SPLIT_ROWS
                            }
                            icon={<OfficeSemanticIcon semanticId="table.insertAbove" size={20} />}
                            displayLabel={t(
                              "presentation.tableRibbon.insertAboveShort",
                            )}
                            label={t("presentation.tableRibbon.insertAboveShort")}
                            onClick={() =>
                              insertTableRowAtFocus(
                                contextualTableBlock.id,
                                "before",
                              )
                            }
                            testId="presentation-ribbon-table-insert-row-above"
                          />
                          <OfficeHomeCommand
                            variant="large"
                            controlId="PPTCTL.CONTEXT.TABLE.LAYOUT.INSERT_BELOW"
                            disabled={
                              !contextualTableFrameIsWritable ||
                              contextualTableBlock.rows.length >=
                                MAX_TABLE_SPLIT_ROWS
                            }
                            icon={<OfficeSemanticIcon semanticId="table.insertBelow" size={20} />}
                            displayLabel={t(
                              "presentation.tableRibbon.insertBelowShort",
                            )}
                            label={t("presentation.tableRibbon.insertBelowShort")}
                            onClick={() =>
                              insertTableRowAtFocus(
                                contextualTableBlock.id,
                                "after",
                              )
                            }
                            testId="presentation-ribbon-table-add-row"
                          />
                          <OfficeHomeCommand
                            variant="large"
                            controlId="PPTCTL.CONTEXT.TABLE.LAYOUT.INSERT_LEFT"
                            disabled={
                              !contextualTableFrameIsWritable ||
                              tableLogicalWidth(contextualTableBlock.rows) >=
                                MAX_TABLE_SPLIT_COLUMNS
                            }
                            icon={<OfficeSemanticIcon semanticId="table.insertLeft" size={20} />}
                            displayLabel={t(
                              "presentation.tableRibbon.insertLeftShort",
                            )}
                            label={t("presentation.tableRibbon.insertLeftShort")}
                            onClick={() =>
                              insertTableColumnAtFocus(
                                contextualTableBlock.id,
                                "before",
                              )
                            }
                            testId="presentation-ribbon-table-insert-column-left"
                          />
                          <OfficeHomeCommand
                            variant="large"
                            controlId="PPTCTL.CONTEXT.TABLE.LAYOUT.INSERT_RIGHT"
                            disabled={
                              !contextualTableFrameIsWritable ||
                              tableLogicalWidth(contextualTableBlock.rows) >=
                                MAX_TABLE_SPLIT_COLUMNS
                            }
                            icon={<OfficeSemanticIcon semanticId="table.insertRight" size={20} />}
                            displayLabel={t(
                              "presentation.tableRibbon.insertRightShort",
                            )}
                            label={t("presentation.tableRibbon.insertRightShort")}
                            onClick={() =>
                              insertTableColumnAtFocus(
                                contextualTableBlock.id,
                                "after",
                              )
                            }
                            testId="presentation-ribbon-table-add-column"
                          />
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.tableMergeGroup}
                      label={t("presentation.tableRibbon.mergeCells")}
                      size="compact"
                      testId="presentation-ribbon-table-merge-group"
                    >
                      <OfficeHomeCommandRow>
                        <OfficeHomeCommand
                          variant="large"
                          controlId="PPTCTL.CONTEXT.TABLE.LAYOUT.MERGE"
                          disabled={!tableSelectionCanMerge}
                          displayLabel={t(
                            "presentation.tableRibbon.mergeShort",
                          )}
                          icon={<OfficeSemanticIcon semanticId="table.merge" size={20} />}
                          // PowerPoint greys Merge Cells without a hint of
                          // how to enable it (PPT-A142): the name only.
                          label={t("presentation.tableMerge")}
                          onClick={mergeTableSelection}
                          testId="presentation-ribbon-table-merge"
                        />
                        <OfficeHomeCommand
                          variant="large"
                          activation="click"
                          buttonRef={tableSplitTriggerRef}
                          controlId="PPTCTL.CONTEXT.TABLE.LAYOUT.SPLIT"
                          disabled={!tableFocusCanUseSplitCommand}
                          displayLabel={t(
                            "presentation.tableRibbon.splitShort",
                          )}
                          icon={<OfficeSemanticIcon semanticId="table.split" size={20} />}
                          label={
                            tableFocusCanUseSplitCommand
                              ? t("presentation.tableSplit")
                              : `${t("presentation.tableSplit")}: ${t("presentation.tableSplitHint")}`
                          }
                          onClick={requestTableCellSplit}
                          testId="presentation-ribbon-table-split"
                        />
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    <OfficeDialogShell
                      open={
                        tableSplitPanelOpen && contextualTableFocus !== null
                      }
                      title={t("presentation.tableSplit")}
                      description=""
                      defaultActionLabel={
                        <span
                          data-control-id="PPTCTL.DIALOG.TABLE_SPLIT.OK"
                          data-testid="presentation-ribbon-table-split-apply"
                        >
                          {t("toolbar.dialogOk")}
                        </span>
                      }
                      cancelActionLabel={
                        <span
                          data-control-id="PPTCTL.DIALOG.TABLE_SPLIT.CANCEL"
                          data-testid="presentation-ribbon-table-split-cancel"
                        >
                          {t("toolbar.cancel")}
                        </span>
                      }
                      onDefaultAction={() => {
                        if (
                          !contextualTableSelectionIsWritable ||
                          !tableSplitDraftIsValid
                        ) {
                          return;
                        }
                        splitTableCellIntoParts(
                          tableSplitRowsValue,
                          tableSplitColumnsValue,
                        );
                      }}
                      onCancel={() => setTableSplitPanelOpen(false)}
                      defaultActionDisabled={
                        !contextualTableSelectionIsWritable ||
                        !tableSplitDraftIsValid
                      }
                      initialFocusRef={tableSplitRowsInputRef}
                      returnFocusRef={tableSplitTriggerRef}
                      testIdPrefix="presentation-ribbon-table-split"
                    >
                      <div
                        data-control-id="PPTCTL.DIALOG.TABLE_SPLIT.ROOT"
                        data-testid="presentation-ribbon-table-split-panel"
                        className="grid grid-cols-2 gap-3"
                      >
                        {/* PowerPoint: Number of columns, then Number of rows (PPT-B017). */}
                        <label className={`${dialogParts.field} grid gap-1`}>
                          {t("presentation.tableSplitColumns")}
                          <input
                            ref={tableSplitRowsInputRef}
                            type="number"
                            inputMode="numeric"
                            min={1}
                            max={tableSplitMaxColumns}
                            step={1}
                            disabled={!contextualTableSelectionIsWritable}
                            value={tableSplitColumnsDraft}
                            onChange={(event) =>
                              setTableSplitColumnsDraft(event.target.value)
                            }
                            data-control-id="PPTCTL.DIALOG.TABLE_SPLIT.COLUMNS.FIELD"
                            data-testid="presentation-ribbon-table-split-columns"
                          />
                        </label>
                        <label className={`${dialogParts.field} grid gap-1`}>
                          {t("presentation.tableSplitRows")}
                          <input
                            type="number"
                            inputMode="numeric"
                            min={1}
                            max={tableSplitMaxRows}
                            step={1}
                            disabled={!contextualTableSelectionIsWritable}
                            value={tableSplitRowsDraft}
                            onChange={(event) =>
                              setTableSplitRowsDraft(event.target.value)
                            }
                            data-control-id="PPTCTL.DIALOG.TABLE_SPLIT.ROWS.FIELD"
                            data-testid="presentation-ribbon-table-split-rows"
                          />
                        </label>
                      </div>
                    </OfficeDialogShell>
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.tableCellGroup}
                      label={t("presentation.tableRibbon.alignment")}
                      size="standard"
                      testId="presentation-ribbon-table-cell-group"
                    >
                      {contextualTableFocus ? (
                        <>
                          <OfficeHomeCommandRow>
                            <div className={presentationChrome.tableCellAlignmentGrid} data-testid="presentation-ribbon-table-cell-alignment-grid" role="group" aria-label={t("presentation.tableCellAlign")}>
                              {/* PowerPoint's Layout > Alignment has no Justify:
                              it stays on Home > Paragraph (PPT-A145). */}
                              {TABLE_LAYOUT_CELL_ALIGNMENTS.map((align) => (
                                <OfficeHomeCommand
                                  key={align}
                                  active={
                                    contextualTableCellFormatState.align
                                      ?.mixed
                                      ? "mixed"
                                      : contextualTableCellFormatState.align
                                          ?.value === align
                                  }
                                  variant="icon"
                                  disabled={!contextualTableSelectionIsWritable}
                                  icon={
                                    <IconParagraphAlign
                                      home
                                      size={18}
                                      alignment={align}
                                    />
                                  }
                                  label={t(
                                    `presentation.tableAlign.${align}` as "presentation.tableAlign.left",
                                  )}
                                  onClick={() =>
                                    updateContextualTableCellStyle({ align })
                                  }
                                  testId={`presentation-ribbon-table-cell-align-${align}-layout`}
                                />
                              ))}
                              {VERTICAL_ANCHORS.map((verticalAlign) => (
                                <OfficeHomeCommand
                                  key={verticalAlign}
                                  active={
                                    contextualTableCellFormatState.verticalAlign
                                      ?.mixed
                                      ? "mixed"
                                      : contextualTableCellFormatState
                                          .verticalAlign?.value === verticalAlign
                                  }
                                  variant="icon"
                                  disabled={!contextualTableSelectionIsWritable}
                                  icon={
                                    verticalAlign === "top" ? (
                                      <AlignVerticalJustifyStart size={18} />
                                    ) : verticalAlign === "middle" ? (
                                      <AlignVerticalJustifyCenter size={18} />
                                    ) : (
                                      <AlignVerticalJustifyEnd size={18} />
                                    )
                                  }
                                  label={t(
                                    `presentation.tableVAlign.${verticalAlign}` as "presentation.tableVAlign.top",
                                  )}
                                  onClick={() =>
                                    updateContextualTableCellStyle({
                                      verticalAlign:
                                        persistedVerticalAnchor(verticalAlign),
                                    })
                                  }
                                  testId={`presentation-ribbon-table-cell-valign-${verticalAlign}-layout`}
                                />
                              ))}
                            </div>
                            <PresentationTableCellChoiceMenu
                              defaultValue="horizontal"
                              disabled={!contextualTableSelectionIsWritable}
                              displayLabel={t("presentation.writingModeLabel")}
                              icon={<OfficeSemanticIcon semanticId="table.cell.textDirection" size={24} />}
                              label={t("presentation.writingModeLabel")}
                              mixed={
                                contextualTableCellFormatState.textDirection
                                  ?.mixed ?? false
                              }
                              mixedLabel={t("toolbar.fontPickerMixed")}
                              value={
                                contextualTableCellFormatState.textDirection
                                  ?.value ?? "horizontal"
                              }
                              options={tableCellTextDirectionOptions}
                              onSelect={(textDirection) =>
                                updateContextualTableCellStyle({
                                  textDirection,
                                })
                              }
                              menuControlId="PPTCTL.CONTEXT.TABLE.LAYOUT.TEXT_DIRECTION"
                              testId="presentation-ribbon-table-cell-text-direction-layout"
                              triggerVariant="large"
                            />
                            <PresentationTableCellMarginMenu
                              disabled={!contextualTableSelectionIsWritable}
                              displayLabel={tableCellMarginMenuLabels.title}
                              labels={tableCellMarginMenuLabels}
                              value={
                                contextualTableCellMarginMenuValue
                              }
                              onCustom={() =>
                                setTableCellMarginsDialogOpen(true)
                              }
                              onSelect={(preset) =>
                                updateContextualTableCellStyle({
                                  ...tableCellMarginPresetPatch(preset),
                                })
                              }
                              testId="presentation-ribbon-table-cell-margin-preset"
                              triggerRef={tableCellMarginsTriggerRef}
                              triggerVariant="large"
                            />
                          </OfficeHomeCommandRow>
                          <OfficeHomeCommandRow className={presentationChrome.tableCellTypographyRow}>
                            <RibbonFontFamilyPicker
                              label={t("toolbar.fontFamily")}
                              activeLabel={contextualTableCellFontFamilyLabel}
                              activeFontFamily={
                                contextualTableCellFormatState.fontFamily
                                  ?.mixed
                                  ? undefined
                                  : (contextualTableCellFontFamilyValue ??
                                    undefined)
                              }
                              sections={contextualTableCellFontSections}
                              disabled={!contextualTableSelectionIsWritable}
                              onSelect={selectContextualTableCellFontFamily}
                              mixed={
                                contextualTableCellFormatState.fontFamily
                                  ?.mixed ?? false
                              }
                              testId="presentation-ribbon-table-cell-font-family-layout"
                            />
                            <RibbonValueCombobox
                              testId="presentation-ribbon-table-cell-font-size-layout"
                              label={t("toolbar.fontSize")}
                              placeholder={t("toolbar.fontSizeDefault")}
                              value={tableCellFontSizeDraft}
                              inputMode="decimal"
                              options={TABLE_CELL_FONT_SIZE_OPTIONS}
                              formatPreview={(value) => {
                                const fontSizePt = parseOfficeFontSizeInput(value, "powerpoint");
                                return fontSizePt === undefined ? null : { fontSizePt };
                              }}
                              disabled={!contextualTableSelectionIsWritable}
                              mixed={
                                contextualTableCellFormatState.fontSize
                                  ?.mixed ?? false
                              }
                              mixedLabel={t("toolbar.fontPickerMixed")}
                              selectOnFocus
                              restoreMenuInvokerOnSelect
                              onDraftChange={(value) => {
                                setTableCellFontSizeDraft(value);
                                setTableCellFontSizeDirty(true);
                              }}
                              onFocusChange={(focused) => {
                                setTableCellFontSizeFocused(focused);
                                if (focused) setTableCellFontSizeDirty(false);
                              }}
                              onRememberSelection={() => undefined}
                              onCommit={commitContextualTableCellFontSize}
                            />
                            <OfficeHomeCommand
                              active={
                                contextualTableCellFormatState.bold?.mixed
                                  ? "mixed"
                                  : (contextualTableCellFormatState.bold
                                      ?.value ?? false)
                              }
                              icon={
                                <OfficeSemanticIcon
                                  semanticId="text.bold"
                                  size={16}
                                />
                              }
                              disabled={!contextualTableSelectionIsWritable}
                              label={t("toolbar.bold")}
                              onClick={() =>
                                updateContextualTableCellStyle({
                                  bold:
                                    contextualTableCellFormatState.bold
                                      ?.mixed ||
                                    !contextualTableCellFormatState.bold?.value
                                      ? true
                                      : undefined,
                                })
                              }
                              testId="presentation-ribbon-table-cell-bold-layout"
                              variant="icon"
                            />
                            <OfficeHomeCommand
                              active={
                                contextualTableCellFormatState.italic?.mixed
                                  ? "mixed"
                                  : (contextualTableCellFormatState.italic
                                      ?.value ?? false)
                              }
                              icon={
                                <OfficeSemanticIcon
                                  semanticId="text.italic"
                                  size={16}
                                />
                              }
                              disabled={!contextualTableSelectionIsWritable}
                              label={t("toolbar.italic")}
                              onClick={() =>
                                updateContextualTableCellStyle({
                                  italic:
                                    contextualTableCellFormatState.italic
                                      ?.mixed ||
                                    !contextualTableCellFormatState.italic
                                      ?.value
                                      ? true
                                      : undefined,
                                })
                              }
                              testId="presentation-ribbon-table-cell-italic-layout"
                              variant="icon"
                            />
                            <OfficeHomeCommand
                              active={
                                contextualTableCellFormatState.underline?.mixed
                                  ? "mixed"
                                  : (contextualTableCellFormatState.underline
                                      ?.value ?? false)
                              }
                              icon={
                                <OfficeSemanticIcon
                                  semanticId="text.underline"
                                  size={16}
                                />
                              }
                              disabled={!contextualTableSelectionIsWritable}
                              label={t("toolbar.underline")}
                              onClick={() =>
                                updateContextualTableCellStyle({
                                  underline:
                                    contextualTableCellFormatState.underline
                                      ?.mixed ||
                                    !contextualTableCellFormatState.underline
                                      ?.value
                                      ? true
                                      : undefined,
                                })
                              }
                              testId="presentation-ribbon-table-cell-underline-layout"
                              variant="icon"
                            />
                          </OfficeHomeCommandRow>
                        </>
                      ) : (
                        <span className="note-editor-office-toolbar__field-label">
                          {t("presentation.tableCellHint")}
                        </span>
                      )}
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.tableAxisGroup}
                      label={t("presentation.tableRibbon.cellSize")}
                      size="wide"
                      testId="presentation-ribbon-table-axis-group"
                    >
                      {contextualTableFocus ? (
                        <>
                          <OfficeHomeCommandRow>
                            <div
                              hidden
                              data-testid="presentation-ribbon-table-autofit-columns-hidden"
                            >
                            <DropdownMenu>
                              <DropdownMenuTrigger asChild>
                                <OfficeHomeMenuTrigger
                                  data-control-id="PPTCTL.CONTEXT.TABLE.LAYOUT.AUTOFIT"
                                  disabled={!contextualTableFrameIsWritable}
                                  displayLabel={t(
                                    "spreadsheet.autoFitColumnWidth",
                                  )}
                                  hasMenuIndicator
                                  icon={<OfficeSemanticIcon semanticId="table.resize" size={20} />}
                                  label={t(
                                    "spreadsheet.autoFitColumnWidth",
                                  )}
                                  testId="presentation-ribbon-table-autofit-columns"
                                />
                              </DropdownMenuTrigger>
                              <RibbonDropdownMenuContent
                                align="start"
                                data-testid="presentation-ribbon-table-autofit-columns-menu"
                              >
                                <DropdownMenuLabel>
                                  {t("spreadsheet.autoFitColumnWidth")}
                                </DropdownMenuLabel>
                                {TABLE_COLUMN_AUTOFIT_MODES.map((mode) => {
                                  const label =
                                    mode === "contents"
                                      ? t("toolbar.autoFitContents")
                                      : mode === "window"
                                        ? t("toolbar.autoFitWindow")
                                        : t(
                                            "spreadsheet.textToColumnsModeFixedWidth",
                                          );
                                  return (
                                    <DropdownMenuItem
                                      key={mode}
                                      data-control-id={
                                        mode === "contents"
                                          ? "PPTCTL.CONTEXT.TABLE.LAYOUT.AUTOFIT.CONTENTS"
                                          : mode === "window"
                                            ? "PPTCTL.CONTEXT.TABLE.LAYOUT.AUTOFIT.WINDOW"
                                            : "PPTCTL.CONTEXT.TABLE.LAYOUT.AUTOFIT.FIXED"
                                      }
                                      disabled={
                                        !contextualTableFrameIsWritable
                                      }
                                      onSelect={() =>
                                        autoFitTableColumns(
                                          contextualTableBlock.id,
                                          mode,
                                        )
                                      }
                                      data-testid={`presentation-ribbon-table-autofit-columns-${mode}`}
                                    >
                                      {label}
                                    </DropdownMenuItem>
                                  );
                                })}
                              </RibbonDropdownMenuContent>
                            </DropdownMenu>
                            </div>
                            <OfficeHomeCommandStack>
                              {(["row", "col"] as const).map((axis) => {
                                const label = t(
                                  axis === "row"
                                    ? "presentation.tableRibbon.heightShort"
                                    : "presentation.tableRibbon.widthShort",
                                );
                                const mixed = axis === "row" ? (contextualTableRowSizeState?.mixed ?? false) : (contextualTableColumnSizeState?.mixed ?? false);
                                return (
                                  <OfficeRibbonNumberField
                                    key={axis}
                                    icon={
                                      <OfficeSemanticIcon
                                        semanticId={
                                          axis === "row"
                                            ? "table.rowHeight"
                                            : "table.columnWidth"
                                        }
                                        size={18}
                                      />
                                    }
                                    visibleLabel={t(
                                      axis === "row"
                                        ? "presentation.tableRibbon.heightShort"
                                        : "presentation.tableRibbon.widthShort",
                                    )}
                                    testId={
                                      axis === "row"
                                        ? "presentation-ribbon-table-row-height"
                                        : "presentation-ribbon-table-col-width"
                                    }
                                    label={label}
                                    unit={sizeUnitLabel}
                                    value={tableAxisSizeDrafts[axis]}
                                    mixed={mixed}
                                    mixedLabel={t("toolbar.fontPickerMixed")}
                                    min={sizeBound(tableAxisPxToCm(MIN_TABLE_AXIS_PX))}
                                    max={sizeBound(tableAxisPxToCm(MAX_TABLE_AXIS_PX))}
                                    step={sizeStep}
                                    disabled={!contextualTableFrameIsWritable}
                                    onCommit={(rawValue) =>
                                      commitContextualTableAxisSize(
                                        axis,
                                        rawValue,
                                        {
                                          allowClear: true,
                                          restoreFocus: false,
                                        },
                                      )
                                    }
                                  />
                                );
                              })}
                            </OfficeHomeCommandStack>
                            <OfficeHomeCommandStack>
                            <OfficeHomeCommand
                              variant="compact"
                              controlId="PPTCTL.CONTEXT.TABLE.LAYOUT.DISTRIBUTE_ROWS"
                              disabled={
                                !contextualTableSelectionIsWritable ||
                                (contextualTableRowSpanRange?.count ??
                                  0) < 2
                              }
                              icon={<OfficeSemanticIcon semanticId="table.distributeRows" size={20} />}
                              displayLabel={t(
                                "presentation.tableRibbon.distributeRowsShort",
                              )}
                              label={t("presentation.tableRibbon.distributeRowsShort")}
                              onClick={() =>
                                distributeTableAxis(
                                  contextualTableBlock.id,
                                  "row",
                                  contextualTableRowSpanRange,
                                )
                              }
                              testId="presentation-ribbon-table-distribute-rows"
                            />
                            <OfficeHomeCommand
                              variant="compact"
                              controlId="PPTCTL.CONTEXT.TABLE.LAYOUT.DISTRIBUTE_COLUMNS"
                              disabled={
                                !contextualTableSelectionIsWritable ||
                                (contextualTableColumnSpanRange?.count ??
                                  0) < 2
                              }
                              icon={<OfficeSemanticIcon semanticId="table.distributeColumns" size={20} />}
                              displayLabel={t(
                                "presentation.tableRibbon.distributeColumnsShort",
                              )}
                              label={t(
                                "presentation.tableRibbon.distributeColumnsShort",
                              )}
                              onClick={() =>
                                distributeTableAxis(
                                  contextualTableBlock.id,
                                  "col",
                                  contextualTableColumnSpanRange,
                                )
                              }
                              testId="presentation-ribbon-table-distribute-columns"
                            />
                            </OfficeHomeCommandStack>
                          </OfficeHomeCommandRow>
                        </>
                      ) : (
                        <span className="note-editor-office-toolbar__field-label">
                          {t("presentation.tableCellHint")}
                        </span>
                      )}
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.tableFrameGroup}
                      label={t("presentation.tableRibbon.tableSize")}
                      size="standard"
                      testId="presentation-ribbon-table-frame-size"
                    >
                      <OfficeHomeCommandRow
                        className={presentationChrome.tableFrameSizeRow}
                      >
                        <OfficeHomeCommandStack>
                          {(["height", "width"] as const).map((axis) => {
                            const currentSize = tableFrameSizeCm(
                              contextualTableBlock[axis],
                            );
                            const label = t(
                              axis === "height"
                                ? "presentation.tableRibbon.heightShort"
                                : "presentation.tableRibbon.widthShort",
                            );
                            return (
                              <OfficeRibbonNumberField
                                key={axis}
                                icon={
                                  <OfficeSemanticIcon
                                    semanticId={
                                      axis === "height"
                                        ? "object.height"
                                        : "object.width"
                                    }
                                    size={18}
                                  />
                                }
                                visibleLabel={t(
                                  axis === "height"
                                    ? "presentation.tableRibbon.heightShort"
                                    : "presentation.tableRibbon.widthShort",
                                )}
                                testId={`presentation-ribbon-table-frame-${axis}`}
                                label={label}
                                unit={sizeUnitLabel}
                                value={
                                  currentSize === null
                                    ? ""
                                    : sizeDisplayFromEmu(
                                        contextualTableBlock[axis],
                                        String(currentSize),
                                      )
                                }
                                min={sizeBound(MIN_TABLE_FRAME_CM)}
                                max={sizeBound(MAX_TABLE_FRAME_CM)}
                                step={sizeStep}
                                disabled={!contextualTableFrameIsWritable}
                                onCommit={(rawValue) =>
                                  commitContextualTableFrameSize(
                                    axis,
                                    presentationMeasurementCmDraft(rawValue, sizeUnit),
                                  )
                                }
                              />
                            );
                          })}
                        </OfficeHomeCommandStack>
                        <label className={presentationChrome.tableFrameAspectLock}>
                          <input
                            type="checkbox"
                            checked={tableFrameAspectLocked}
                            data-control-id="PPTCTL.CONTEXT.TABLE.LAYOUT.LOCK_ASPECT"
                            disabled={!contextualTableSelectionIsWritable}
                            onChange={(event) =>
                              setTableFrameAspectLocked(event.target.checked)
                            }
                            data-testid="presentation-ribbon-table-frame-aspect-lock"
                          />
                          {t("presentation.tableRibbon.lockAspectRatio")}
                        </label>
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    {/* PowerPoint's Layout tab ends with Arrange (Bring
                    Forward, Send Backward, Align, Group, Rotate) and Format
                    Pane (PPT-A148); PowerPoint for the web hides them. */}
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.tableArrangeGroup}
                      label={t("presentation.home.arrange")}
                      size="standard"
                      testId="presentation-ribbon-table-arrange-group"
                    >
                      {


}
                      <OfficeHomeCommandRow className={presentationChrome.pictureArrangeRow}>
                        <DropdownMenu>
                          <OfficeHomeSplitLayout
                            testId="presentation-ribbon-table-bring-forward-split"
                            primary={
                              <OfficeHomeCommand
                                disabled={
                                  !hasLayoutGeometry(contextualTableBlock) ||
                                  !contextualTableLayerCapability.canMoveForward
                                }
                                icon={<IconBringForward size={20} />}
                                label={t(
                                  "presentation.contextMenu.bringForward",
                                )}
                                onClick={() =>
                                  bringBlockForward(
                                    contextualTableBlock.id,
                                    true,
                                  )
                                }
                                testId="presentation-ribbon-table-bring-forward"
                              />
                            }
                            trigger={
                              <DropdownMenuTrigger asChild>
                                <OfficeHomeMenuTrigger
                                  variant="icon"
                                  disabled={
                                    !hasLayoutGeometry(contextualTableBlock) ||
                                    !contextualTableLayerCapability.canMoveForward
                                  }
                                  icon={
                                    <OfficeSemanticIcon
                                      semanticId="menu.chevron.down"
                                      size={12}
                                    />
                                  }
                                  label={t(
                                    "presentation.contextMenu.bringToFront",
                                  )}
                                  testId="presentation-ribbon-table-bring-forward-menu"
                                />
                              </DropdownMenuTrigger>
                            }
                          />
                          <RibbonDropdownMenuContent
                            align="start"
                            data-testid="presentation-ribbon-table-bring-forward-menu-content"
                          >
                            <DropdownMenuItem
                              disabled={
                                !contextualTableLayerCapability.canMoveForward
                              }
                              onSelect={() =>
                                bringBlockToFront(
                                  contextualTableBlock.id,
                                  true,
                                )
                              }
                              data-testid="presentation-ribbon-table-bring-to-front"
                            >
                              <IconBringToFront size={14} />
                              {t("presentation.contextMenu.bringToFront")}
                            </DropdownMenuItem>
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                        <DropdownMenu>
                          <OfficeHomeSplitLayout
                            testId="presentation-ribbon-table-send-backward-split"
                            primary={
                              <OfficeHomeCommand
                                disabled={
                                  !hasLayoutGeometry(contextualTableBlock) ||
                                  !contextualTableLayerCapability.canMoveBackward
                                }
                                icon={<IconSendBackward size={20} />}
                                label={t(
                                  "presentation.contextMenu.sendBackward",
                                )}
                                onClick={() =>
                                  sendBlockBackward(
                                    contextualTableBlock.id,
                                    true,
                                  )
                                }
                                testId="presentation-ribbon-table-send-backward"
                              />
                            }
                            trigger={
                              <DropdownMenuTrigger asChild>
                                <OfficeHomeMenuTrigger
                                  variant="icon"
                                  disabled={
                                    !hasLayoutGeometry(contextualTableBlock) ||
                                    !contextualTableLayerCapability.canMoveBackward
                                  }
                                  icon={
                                    <OfficeSemanticIcon
                                      semanticId="menu.chevron.down"
                                      size={12}
                                    />
                                  }
                                  label={t(
                                    "presentation.contextMenu.sendToBack",
                                  )}
                                  testId="presentation-ribbon-table-send-backward-menu"
                                />
                              </DropdownMenuTrigger>
                            }
                          />
                          <RibbonDropdownMenuContent
                            align="start"
                            data-testid="presentation-ribbon-table-send-backward-menu-content"
                          >
                            <DropdownMenuItem
                              disabled={
                                !contextualTableLayerCapability.canMoveBackward
                              }
                              onSelect={() =>
                                sendBlockToBack(
                                  contextualTableBlock.id,
                                  true,
                                )
                              }
                              data-testid="presentation-ribbon-table-send-to-back"
                            >
                              <IconSendToBack size={14} />
                              {t("presentation.contextMenu.sendToBack")}
                            </DropdownMenuItem>
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              variant="icon"
                              active={alignToSlide}
                              displayLabel={t(
                                "presentation.ribbon.align",
                              )}
                              hasMenuIndicator
                              icon={<IconAlignObjects size={20} />}
                              label={t("presentation.ribbon.align")}
                              disabled={
                                !hasLayoutGeometry(contextualTableBlock)
                              }
                              testId="presentation-ribbon-table-object-align"
                            />
                          </DropdownMenuTrigger>
                          <RibbonDropdownMenuContent
                            align="start"
                            data-testid="presentation-ribbon-table-object-align-menu"
                          >
                            <DropdownMenuLabel>
                              {t("presentation.ribbon.align")}
                            </DropdownMenuLabel>
                            <DropdownMenuCheckboxItem
                              checked={alignToSlide}
                              onCheckedChange={(checked) =>
                                writeAlignToSlide(checked === true)
                              }
                              data-testid="presentation-ribbon-table-align-to-slide"
                            >
                              {t("presentation.align.toSlideLabel")}
                            </DropdownMenuCheckboxItem>
                            <DropdownMenuSeparator />
                            {(
                              [
                                "left",
                                "hcenter",
                                "right",
                                "top",
                                "vcenter",
                                "bottom",
                              ] as const
                            ).map((edge) => (
                              <DropdownMenuItem
                                key={edge}
                                disabled={
                                  !canFlipSelection ||
                                  selectedGeometryBlocks.length <
                                    (alignToSlide ? 1 : 2)
                                }
                                onSelect={() => alignSelectedBlocks(edge)}
                                data-testid={`presentation-ribbon-table-align-${edge}`}
                              >
                                <IconAlignObjects size={14} />
                                {t(
                                  `presentation.align.${edge}` as "presentation.align.left",
                                )}
                              </DropdownMenuItem>
                            ))}
                            <DropdownMenuSeparator />
                            <DropdownMenuItem
                              disabled={
                                !canFlipSelection ||
                                selectedGeometryBlocks.length <
                                  (alignToSlide ? 1 : 3)
                              }
                              onSelect={() =>
                                distributeSelectedBlocks("horizontal")
                              }
                              data-testid="presentation-ribbon-table-distribute-horizontal"
                            >
                              <IconDistributeObjects size={14} />
                              {t(
                                "presentation.home.distributeHorizontally",
                              )}
                            </DropdownMenuItem>
                            <DropdownMenuItem
                              disabled={
                                !canFlipSelection ||
                                selectedGeometryBlocks.length <
                                  (alignToSlide ? 1 : 3)
                              }
                              onSelect={() =>
                                distributeSelectedBlocks("vertical")
                              }
                              data-testid="presentation-ribbon-table-distribute-vertical"
                            >
                              <IconDistributeObjects size={14} />
                              {t(
                                "presentation.home.distributeVertically",
                              )}
                            </DropdownMenuItem>
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              variant="icon"
                              displayLabel={t("presentation.group.group")}
                              hasMenuIndicator
                              icon={<OfficeSemanticIcon semanticId="object.group" size={20} />}
                              label={t("presentation.group.group")}
                              disabled={!contextualTableCanUngroup}
                              data-control-id="PPTCTL.CONTEXT.TABLE.ARRANGE.GROUP"
                              testId="presentation-ribbon-table-group-menu"
                            />
                          </DropdownMenuTrigger>
                          <RibbonDropdownMenuContent
                            align="start"
                            data-testid="presentation-ribbon-table-group-menu-content"
                          >
                            <DropdownMenuLabel>
                              {t("presentation.group.group")}
                            </DropdownMenuLabel>
                            {/* PowerPoint never groups a table: Group stays
                            grey; Ungroup frees a table already in a group. */}
                            <DropdownMenuItem
                              disabled
                              data-testid="presentation-ribbon-table-group"
                            >
                              <OfficeSemanticIcon semanticId="object.group" size={14} />
                              {t("presentation.group.group")}
                            </DropdownMenuItem>
                            <DropdownMenuItem
                              disabled={!contextualTableCanUngroup}
                              onSelect={ungroupSelectedBlocks}
                              data-testid="presentation-ribbon-table-ungroup"
                            >
                              <OfficeSemanticIcon semanticId="object.ungroup" size={14} />
                              {t("presentation.group.ungroup")}
                            </DropdownMenuItem>
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <OfficeHomeMenuTrigger
                              variant="icon"
                              displayLabel={t(
                                "presentation.geometry.rotation",
                              )}
                              hasMenuIndicator
                              icon={<IconRotateObjects size={20} />}
                              label={t("presentation.geometry.rotation")}
                              disabled={
                                !hasLayoutGeometry(contextualTableBlock) ||
                                !canFlipSelection
                              }
                              testId="presentation-ribbon-table-rotate"
                            />
                          </DropdownMenuTrigger>
                          <RibbonDropdownMenuContent
                            align="start"
                            data-testid="presentation-ribbon-table-rotate-menu"
                          >
                            <DropdownMenuLabel>
                              {t("presentation.geometry.rotation")}
                            </DropdownMenuLabel>
                            <DropdownMenuItem
                              disabled={!canFlipSelection}
                              onSelect={() =>
                                quarterTurnSelectedBlocks("right")
                              }
                              data-testid="presentation-ribbon-table-rotate-right"
                            >
                              <IconRotateObjects size={14} />
                              {t("presentation.rotate.right")}
                            </DropdownMenuItem>
                            <DropdownMenuItem
                              disabled={!canFlipSelection}
                              onSelect={() =>
                                quarterTurnSelectedBlocks("left")
                              }
                              data-testid="presentation-ribbon-table-rotate-left"
                            >
                              <IconRotateObjects size={14} />
                              {t("presentation.rotate.left")}
                            </DropdownMenuItem>
                            <DropdownMenuSeparator />
                            <DropdownMenuItem
                              disabled={!canFlipSelection}
                              onSelect={() => flipSelectedBlocks("flipV")}
                              data-testid="presentation-ribbon-table-flip-vertical"
                            >
                              <IconFlipVertical size={14} />
                              {t("presentation.flip.vertical")}
                            </DropdownMenuItem>
                            <DropdownMenuItem
                              disabled={!canFlipSelection}
                              onSelect={() => flipSelectedBlocks("flipH")}
                              data-testid="presentation-ribbon-table-flip-horizontal"
                            >
                              <IconFlipHorizontal size={14} />
                              {t("presentation.flip.horizontal")}
                            </DropdownMenuItem>
                          </RibbonDropdownMenuContent>
                        </DropdownMenu>
                      </OfficeHomeCommandRow>
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.tableFormatPaneGroup}
                      label={t("presentation.ui.shapeFormat.formatPane")}
                      size="narrow"
                      testId="presentation-ribbon-table-format-pane-group"
                    >
                      <OfficeHomeCommand
                        active={formatPaneShowing && formatPaneInFront}
                        variant="large"
                        controlId="PPTCTL.CONTEXT.TABLE.LAYOUT.FORMAT_PANE"
                        icon={<OfficeSemanticIcon semanticId="image.format" size={24} />}
                        label={t("presentation.ui.shapeFormat.formatPane")}
                        onClick={toggleFormatPane}
                        testId="presentation-ribbon-table-format-pane"
                      />
                    </OfficeHomeRibbonGroup>
                    <div
                      className={`note-editor-office-toolbar__group ${presentationChrome.tableDeleteGroup}`}
                      aria-hidden="true"
                      data-office-web-hidden="true"
                      hidden
                      data-ribbon-label={t("presentation.deleteBlock")}
                    >
                      <div className="note-editor-ribbon-row">
                        <button
                          type="button"
                          className="office-ribbon-control note-editor-office-toolbar__button" data-office-control-part="button"
                          onClick={() => deleteBlock(contextualTableBlock.id)}
                          data-testid="presentation-ribbon-table-delete"
                          style={PRESENTATION_RIBBON_LABELED_BUTTON_STYLE}
                        >
                          <Trash2 size={15} />
                          {t("presentation.deleteBlock")}
                        </button>
                      </div>
                    </div>
                  </>
  );
}
