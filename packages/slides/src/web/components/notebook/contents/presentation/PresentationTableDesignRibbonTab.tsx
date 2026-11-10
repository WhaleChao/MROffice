"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Table Design contextual tab of the presentation Ribbon (Table Style
 * Options, Table Styles, Shading, Borders, WordArt Styles, Draw Borders),
 * drawn from the window bindings it reads.
 *
 * Extracted verbatim from the render of
 * `components/notebook/contents/PresentationWindow.tsx`: the JSX is unchanged;
 * the window passes the bindings it used to close over as `ctl`, and the tab
 * condition and key stay at the call site. No hooks run here, so the window's
 * hook order is unchanged.
 */

import { lineWeightDesignPxToPoints, lineWeightPointsToDesignPx } from "@mrnote/core/web/lib/presentation/text-units.ts";
import {
  OfficeHomeRibbonGroup,
  OfficeHomeCommandRow,
  OfficeHomeCommand,
  OfficeHomeMenuTrigger,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { PresentationTableStyleGallery } from "@mrnote/core/web/components/notebook/contents/ribbon/PresentationTableStyleGallery.tsx";
import {
  PresentationTableCellColorMenu,
  TABLE_CELL_BEVEL_LABELS,
  TABLE_CELL_FONT_SIZE_OPTIONS,
  PresentationTableCellChoiceMenu,
  PresentationTableBordersMenu,
  TABLE_CELL_BORDER_WIDTH_OPTIONS,
  PresentationTableCellMarginMenu,
  TABLE_CELL_MARGIN_VALUE_OPTIONS,
  TABLE_CELL_BORDER_POINT_WEIGHTS,
} from "./PresentationTableCellMenus";
import { tableWithBackground, tableWithEffect, type PresentationTableStyleSelection } from "@mrnote/core/web/lib/presentation-table-buttons.ts";
import { PresentationTableEffectsMenu } from "./PresentationTableEffectsMenu";
import {
  type TableCellBevelPreset,
  CELL_ALIGNMENTS,
  TABLE_CELL_BORDER_STYLES,
  type TableCellBorderStyle,
  normalizeCellColor,
  TABLE_CELL_MARGIN_PRESETS,
  tableCellMarginPresetPatch,
  type CellStyleInput,
  type ResolvedCellStyle,
  type TableBorderSelectionPosition,
  type TableBorderTarget,
  type TableCellBorder,
  type TableCellFontFamily,
  type TableStyleLook,
} from "@mrnote/core/web/lib/presentation-table-style.ts";
import { DocsHomeIcon } from "@mrnote/core/web/components/console/editor/DocumentHomeRibbonIcons.tsx";
import { RibbonFontFamilyPicker } from "@mrnote/core/web/components/office-ribbon/RibbonFontFamilyPicker.tsx";
import { RibbonValueCombobox } from "@mrnote/core/web/components/office-ribbon/RibbonValueCombobox.tsx";
import { parseOfficeFontSizeInput } from "@mrnote/core/web/components/office-ribbon/officeFontSizeCatalog.ts";
import { IconParagraphAlign, IconAlignTextMenu, IconTableDraw } from "@mrnote/core/web/components/notebook/contents/PresentationRibbonIcons.tsx";
import { tableCellAlignIcon } from "../../../../lib/presentation/table-canvas-layout";
import {
  AlignLeft,
  AlignHorizontalJustifyCenter,
  AlignRight,
  AlignVerticalJustifyStart,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  PanelLeft,
  PanelRight,
  PanelTop,
  PanelBottom,
  ChevronsUpDown,
} from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { VERTICAL_ANCHORS, persistedVerticalAnchor } from "@mrnote/core/web/lib/presentation-vertical-anchor.ts";
import { PresentationTableWordArtGroup } from "./PresentationTableWordArtGroup";
import { presentationWordArtStyleOf, type PresentationWordArtStyle } from "@mrnote/core/web/lib/presentation-wordart-gallery.ts";
import { DocsInsertIcon } from "@mrnote/core/web/components/console/editor/DocumentInsertRibbonIcons.tsx";
import { patchOfficeTextPaint, type OfficeTextPaint } from "@mrnote/core/web/lib/office-text-paint.ts";
import { OfficeBorderLineStyleMenu, OfficeBorderLinePreview, OfficeBorderWeightMenu } from "@mrnote/core/web/components/office-ribbon/OfficeBordersMenu.tsx";
import type {
  CSSProperties,
  Dispatch,
  ReactNode,
  RefObject,
  SetStateAction,
} from "react";
import type { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { embeddedPresentationAutosave } from "../../../../lib/presentation-embedded-session";
import type { DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";
import type { PresentationTableBlock, PresentationTableCell } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { TableCellMarginMenuValue } from "./PresentationTableCellMenus";
import type { OfficeColorPalettePanelLabels } from "@mrnote/core/web/components/office-ribbon/OfficeColorPalettePanel.tsx";
import type { TableCellMarginField } from "../../../../lib/presentation/TableMarginsDialog";
import type { TableBorderPenMode } from "../../../../lib/presentation-table-border-pen";

/** The window bindings this tab reads, passed through unchanged. */
export interface PresentationTableDesignRibbonTabCtl {
  pageId: string;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  tableCellPaletteTarget: "fill" | "tableBackground" | "color" | "highlight" | "border" | null;
  setTableCellPaletteTarget: Dispatch<SetStateAction<"fill" | "tableBackground" | "color" | "highlight" | "border" | null>>;
  tableCellRecentColors: string[];
  setTableCellRecentColors: Dispatch<SetStateAction<string[]>>;
  tableCellFontSizeDraft: string;
  setTableCellFontSizeDraft: Dispatch<SetStateAction<string>>;
  setTableCellFontSizeFocused: Dispatch<SetStateAction<boolean>>;
  setTableCellFontSizeDirty: Dispatch<SetStateAction<boolean>>;
  tableCellMarginDrafts: Record<"marginTop" | "marginBottom" | "marginLeft" | "marginRight", string>;
  setTableCellMarginDrafts: Dispatch<SetStateAction<Record<"marginTop" | "marginBottom" | "marginLeft" | "marginRight", string>>>;
  setTableCellMarginFocused: Dispatch<SetStateAction<"marginTop" | "marginBottom" | "marginLeft" | "marginRight" | null>>;
  setTableCellMarginDirty: Dispatch<SetStateAction<boolean>>;
  setTableCellMarginsDialogOpen: Dispatch<SetStateAction<boolean>>;
  tableCellMarginsTriggerRef: RefObject<HTMLButtonElement | null>;
  tableBorderWidthDraft: string;
  setTableBorderWidthDraft: Dispatch<SetStateAction<string>>;
  setTableBorderWidthFocused: Dispatch<SetStateAction<boolean>>;
  setTableBorderWidthDirty: Dispatch<SetStateAction<boolean>>;
  tableGridlinesVisible: boolean;
  setTableGridlinesVisible: Dispatch<SetStateAction<boolean>>;
  tableBorderTarget: "all" | "outside" | "inside" | "inside-horizontal" | "inside-vertical" | "top" | "bottom" | "left" | "right" | "diagonal-down" | "diagonal-up" | "none";
  tableBorderPenMode: "draw" | "erase" | null;
  tableBorderPen: TableCellBorder;
  autosave: ReturnType<typeof embeddedPresentationAutosave>;
  tableDrawAuthorityRef: RefObject<{ pageId: string; readOnly: boolean; }>;
  activeEffectiveTheme: DeckTheme | undefined;
  deckWordArtStyles: readonly PresentationWordArtStyle[];
  wordArtStyleLabel: (id: string) => string;
  contextualTableBlock: PresentationTableBlock;
  contextualTableSelectionIsWritable: boolean;
  contextualTableFocus: { blockId: string; row: number; col: number; } | null;
  contextualTableCell: PresentationTableCell | null;
  contextualTableCellStyle: ResolvedCellStyle | null;
  contextualTableCellFormatState: { fill: Readonly<{ mixed: boolean; value: string | null; }> | null; bevel: Readonly<{ mixed: boolean; value: "none" | "circle" | "relaxedInset" | "slope" | "cross" | "angle" | "softRound" | "convex" | "coolSlant" | "divot" | "riblet" | "hardEdge" | "artDeco"; }> | null; color: Readonly<{ mixed: boolean; value: string | null; }> | null; highlight: Readonly<{ mixed: boolean; value: string | null; }> | null; fontFamily: Readonly<{ mixed: boolean; value: TableCellFontFamily | null; }> | null; fontSize: Readonly<{ mixed: boolean; value: number | null; }> | null; bold: Readonly<{ mixed: boolean; value: boolean; }> | null; italic: Readonly<{ mixed: boolean; value: boolean; }> | null; underline: Readonly<{ mixed: boolean; value: boolean; }> | null; strike: Readonly<{ mixed: boolean; value: boolean; }> | null; superscript: Readonly<{ mixed: boolean; value: boolean; }> | null; subscript: Readonly<{ mixed: boolean; value: boolean; }> | null; align: Readonly<{ mixed: boolean; value: "left" | "right" | "center" | "justify" | "distributed" | "thaiDistributed" | "justifyLow"; }> | null; verticalAlign: Readonly<{ mixed: boolean; value: "top" | "bottom" | "middle"; }> | null; textDirection: Readonly<{ mixed: boolean; value: "horizontal" | "rotate-90" | "rotate-270" | "stacked" | "eastAsianVertical" | "mongolianVertical" | "stackedRightToLeft"; }> | null; marginLeft: Readonly<{ mixed: boolean; value: number | null; }> | null; marginRight: Readonly<{ mixed: boolean; value: number | null; }> | null; marginTop: Readonly<{ mixed: boolean; value: number | null; }> | null; marginBottom: Readonly<{ mixed: boolean; value: number | null; }> | null; };
  contextualTableBorderState: { style: { value: "none" | "solid" | "dash" | "dot" | "double"; mixed: boolean; }; width: { value: number; mixed: boolean; }; } | null;
  contextualTableBorderTargetOptions: Readonly<{ disabled?: boolean; icon?: ReactNode; label: string; value: "all" | "outside" | "inside" | "inside-horizontal" | "inside-vertical" | "top" | "bottom" | "left" | "right" | "diagonal-down" | "diagonal-up" | "none"; }>[];
  contextualTableCellFontFamilyValue: TableCellFontFamily | null;
  contextualTableCellFontFamilyLabel: string;
  contextualTableCellFontSections: { id: string; options: ({ id: "Aptos" | "Calibri" | "Arial" | "Cambria" | "Times New Roman" | "Georgia" | "Courier New" | "Consolas" | "DengXian" | "Microsoft YaHei" | "SimSun" | "PingFang SC" | "Microsoft JhengHei" | "Yu Gothic" | "Malgun Gothic"; label: "Aptos" | "Calibri" | "Arial" | "Cambria" | "Times New Roman" | "Georgia" | "Courier New" | "Consolas" | "DengXian (等线)" | "Microsoft YaHei (微软雅黑)" | "SimSun (宋体)" | "PingFang SC (苹方)" | "Microsoft JhengHei (微軟正黑體)" | "Yu Gothic (游ゴシック)" | "Malgun Gothic (맑은 고딕)"; value: "Aptos" | "Calibri" | "Arial" | "Cambria" | "Times New Roman" | "Georgia" | "Courier New" | "Consolas" | "DengXian" | "Microsoft YaHei" | "SimSun" | "PingFang SC" | "Microsoft JhengHei" | "Yu Gothic" | "Malgun Gothic"; previewFamily: "Aptos, Calibri, 'Segoe UI', sans-serif" | "Calibri, 'Segoe UI', Candara, sans-serif" | "Arial, 'Helvetica Neue', Helvetica, sans-serif" | "Cambria, Georgia, 'Times New Roman', serif" | "'Times New Roman', Times, serif" | "Georgia, Cambria, 'Times New Roman', serif" | "'Courier New', Courier, monospace" | "Consolas, 'SFMono-Regular', Menlo, monospace" | "DengXian, 'Microsoft YaHei', 'PingFang SC', sans-serif" | "'Microsoft YaHei', 'PingFang SC', sans-serif" | "SimSun, 'Songti SC', serif" | "'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', sans-serif" | "'Microsoft JhengHei', 'PingFang TC', sans-serif" | "'Yu Gothic', 'Hiragino Kaku Gothic ProN', sans-serif" | "'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif"; selected: boolean; testId: string; } | { id: string; label: string; value: null; selected: boolean; testId: string; })[]; }[];
  contextualTableCellMarginMenuValue: TableCellMarginMenuValue;
  tableCellMarginMenuLabels: Readonly<Record<TableCellMarginMenuValue, string> & { title: string; }>;
  tableCellColorPaletteLabels: OfficeColorPalettePanelLabels;
  updateTableCellStyle: (blockId: string, rowIndex: number, cellIndex: number, patch: Partial<CellStyleInput> | ((position: TableBorderSelectionPosition, style: ResolvedCellStyle) => Partial<CellStyleInput>)) => void;
  updateContextualTableCellStyle: (patch: Partial<CellStyleInput>) => void;
  writeContextualTableTextPaint: (write: (paint: OfficeTextPaint, style: ResolvedCellStyle) => Partial<CellStyleInput>) => void;
  contextualTableTextPaint: OfficeTextPaint | null;
  contextualTableTextPaintMixed: boolean;
  pickContextualTableCellColor: (field: "fill" | "color" | "highlight", color: string | null) => void;
  selectContextualTableCellFontFamily: (fontFamily: string | null) => void;
  commitContextualTableCellFontSize: (rawValue: string, options: { allowClear: boolean; restoreFocus: boolean; }) => boolean;
  commitContextualTableCellMargin: (field: TableCellMarginField, rawValue: string, options: { allowClear: boolean; restoreFocus: boolean; }) => boolean;
  commitContextualTableBorderWidth: (rawValue: string) => boolean;
  selectContextualTableBorderStyle: (style: TableCellBorderStyle) => void;
  toggleTableBorderPen: (mode: TableBorderPenMode) => void;
  updateDrawingPen: (patch: Partial<TableCellBorder>) => void;
  armDrawingPen: () => void;
  applyDrawingPenToBorders: (target: TableBorderTarget) => void;
  updateTableFlags: (blockId: string, patch: Partial<Pick<PresentationTableBlock, "borders" | "headerRow" | "bandedRows" | "totalRow" | "firstColumn" | "lastColumn" | "bandedColumns">>) => void;
  updateTableFrame: (blockId: string, change: (block: PresentationTableBlock) => PresentationTableBlock) => void;
  previewTableStyle: (blockId: string, selection: PresentationTableStyleSelection | null) => void;
  commitTableStyle: (blockId: string, selection: PresentationTableStyleSelection) => void;
  deckTableStyleLooks: TableStyleLook[];
  tableCellTextDirectionOptions: { value: "horizontal" | "rotate-90" | "rotate-270" | "stacked"; label: string; }[];
}

export function PresentationTableDesignRibbonTab({ ctl }: { ctl: PresentationTableDesignRibbonTabCtl }) {
  const {
    pageId,
    t,
    tableCellPaletteTarget,
    setTableCellPaletteTarget,
    tableCellRecentColors,
    setTableCellRecentColors,
    tableCellFontSizeDraft,
    setTableCellFontSizeDraft,
    setTableCellFontSizeFocused,
    setTableCellFontSizeDirty,
    tableCellMarginDrafts,
    setTableCellMarginDrafts,
    setTableCellMarginFocused,
    setTableCellMarginDirty,
    setTableCellMarginsDialogOpen,
    tableCellMarginsTriggerRef,
    tableBorderWidthDraft,
    setTableBorderWidthDraft,
    setTableBorderWidthFocused,
    setTableBorderWidthDirty,
    tableGridlinesVisible,
    setTableGridlinesVisible,
    tableBorderTarget,
    tableBorderPenMode,
    tableBorderPen,
    autosave,
    tableDrawAuthorityRef,
    activeEffectiveTheme,
    deckWordArtStyles,
    wordArtStyleLabel,
    contextualTableBlock,
    contextualTableSelectionIsWritable,
    contextualTableFocus,
    contextualTableCell,
    contextualTableCellStyle,
    contextualTableCellFormatState,
    contextualTableBorderState,
    contextualTableBorderTargetOptions,
    contextualTableCellFontFamilyValue,
    contextualTableCellFontFamilyLabel,
    contextualTableCellFontSections,
    contextualTableCellMarginMenuValue,
    tableCellMarginMenuLabels,
    tableCellColorPaletteLabels,
    updateTableCellStyle,
    updateContextualTableCellStyle,
    writeContextualTableTextPaint,
    contextualTableTextPaint,
    contextualTableTextPaintMixed,
    pickContextualTableCellColor,
    selectContextualTableCellFontFamily,
    commitContextualTableCellFontSize,
    commitContextualTableCellMargin,
    commitContextualTableBorderWidth,
    selectContextualTableBorderStyle,
    toggleTableBorderPen,
    updateDrawingPen,
    armDrawingPen,
    applyDrawingPenToBorders,
    updateTableFlags,
    updateTableFrame,
    previewTableStyle,
    commitTableStyle,
    deckTableStyleLooks,
    tableCellTextDirectionOptions,
  } = ctl;
  return (
                  <>
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.tableStyleOptionsGroup}
                      label={t("presentation.tableRibbon.styleOptions")}
                      size="wide"
                      testId="presentation-ribbon-table-style-options-group"
                    >
                      {([
                        [["headerRow", t("presentation.tableHeaderRow"), "PPTCTL.CONTEXT.TABLE.DESIGN.OPTION.HEADER_ROW", "table.style.headerRow"], ["firstColumn", t("presentation.tableFirstColumn"), "PPTCTL.CONTEXT.TABLE.DESIGN.OPTION.FIRST_COLUMN", "table.style.firstColumn"]],
                        [["totalRow", t("presentation.tableTotalRow"), "PPTCTL.CONTEXT.TABLE.DESIGN.OPTION.TOTAL_ROW", "table.style.totalRow"], ["lastColumn", t("presentation.tableLastColumn"), "PPTCTL.CONTEXT.TABLE.DESIGN.OPTION.LAST_COLUMN", "table.style.lastColumn"]],
                        [["bandedRows", t("presentation.tableBandedRows"), "PPTCTL.CONTEXT.TABLE.DESIGN.OPTION.BANDED_ROWS", "table.style.bandedRows"], ["bandedColumns", t("presentation.tableBandedColumns"), "PPTCTL.CONTEXT.TABLE.DESIGN.OPTION.BANDED_COLUMNS", "table.style.bandedColumns"]],
                      ] as const).map((row, rowIndex) => (
                        <OfficeHomeCommandRow key={rowIndex}>
                          {row.map(([key, label, controlId, semanticId]) => (
                            <OfficeHomeCommand
                              key={key}
                              active={contextualTableBlock[key] === true}
                              controlId={controlId}
                              disabled={!contextualTableSelectionIsWritable}
                              icon={<OfficeSemanticIcon semanticId={semanticId} size={18} />}
                              label={label}
                              onClick={() => updateTableFlags(contextualTableBlock.id, { [key]: contextualTableBlock[key] ? undefined : true })}
                              testId={`presentation-ribbon-table-${key}`}
                              variant="compact"
                            />
                          ))}
                        </OfficeHomeCommandRow>
                      ))}
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.tableStyleGalleryGroup}
                      controlId="PPTCTL.CONTEXT.TABLE.DESIGN.STYLES"
                      label={t("presentation.tableRibbon.tableStyles")}
                      size="gallery"
                      testId="presentation-ribbon-table-style-gallery-group"
                    >
                      <PresentationTableStyleGallery
                        accentColors={activeEffectiveTheme?.accentColors}
                        deckLooks={deckTableStyleLooks}
                        disabled={!contextualTableSelectionIsWritable}
                        disabledReason={t(
                          "presentation.tableRibbon.styleReadOnly",
                        )}
                        labels={{
                          ariaLabel: t(
                            "presentation.tableRibbon.tableStyles",
                          ),
                          more: t("toolbar.quickOverflow"),
                          builtIn: t("toolbar.stylesBuiltinSection"),
                          authored: t(
                            "presentation.tableRibbon.accentStyles",
                          ),
                          deck: t(
                            "presentation.tableRibbon.documentStyles",
                          ),
                          clear: t(
                            "presentation.tableRibbon.clearStyle",
                          ),
                          unavailable: t(
                            "presentation.tableRibbon.currentNativeStyle",
                          ),
                        }}
                        onCommit={(selection) =>
                          commitTableStyle(contextualTableBlock.id, selection)
                        }
                        onPreviewChange={(selection) =>
                          previewTableStyle(
                            contextualTableBlock.id,
                            selection,
                          )
                        }
                        table={contextualTableBlock}
                      />
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.tableCellFormatGroup}
                      label={t("presentation.tableRibbon.shading")}
                      size="narrow"
                      testId="presentation-ribbon-table-cell-format-group"
                    >
                      {contextualTableFocus &&
                      contextualTableCell &&
                      contextualTableCellStyle ? (
                        <>
                          <div className="note-editor-ribbon-row">
                            <PresentationTableCellColorMenu
                              anchorClassName={
                                presentationChrome.backgroundPaletteAnchor
                              }
                              displayLabel={t("toolbar.cellShading")}
                              icon={
                                <span
                                  className="note-editor-office-toolbar__color-glyph"
                                  aria-hidden="true"
                                >
                                  <OfficeSemanticIcon semanticId="table.shading" size={24} />
                                  <span
                                    className="note-editor-office-toolbar__color-underline"
                                    style={{
                                      backgroundColor:
                                        contextualTableCellFormatState.fill
                                          ?.mixed
                                          ? "transparent"
                                          : (contextualTableCellFormatState.fill
                                              ?.value ?? "transparent"),
                                    }}
                                  />
                                </span>
                              }
                              label={t("toolbar.cellShading")}
                              labels={tableCellColorPaletteLabels}
                              menuControlId="PPTCTL.CONTEXT.TABLE.DESIGN.SHADING.MENU"
                              mixed={
                                contextualTableCellFormatState.fill?.mixed ??
                                false
                              }
                              mixedLabel={t("toolbar.fontPickerMixed")}
                              noFillLabel={t("presentation.home.noFill")}
                              onDismiss={() =>
                                setTableCellPaletteTarget(null)
                              }
                              onPick={(color) =>
                                pickContextualTableCellColor("fill", color)
                              }
                              onToggle={() =>
                                setTableCellPaletteTarget((current) =>
                                  current === "fill" ? null : "fill",
                                )
                              }
                              open={tableCellPaletteTarget === "fill"}
                              recents={tableCellRecentColors}
                              disabled={!contextualTableSelectionIsWritable}
                              secondary={{
                                label: t("presentation.tableRibbon.tableBackground"),
                                noFillLabel: t("presentation.home.noFill"),
                                open: tableCellPaletteTarget === "tableBackground",
                                onOpen: () => setTableCellPaletteTarget("tableBackground"),
                                onPick: (color) => {
                                  setTableCellPaletteTarget(null);
                                  updateTableFrame(contextualTableBlock.id, (block) =>
                                    tableWithBackground(block, color),
                                  );
                                },
                                testId: "presentation-ribbon-table-background",
                                value:
                                  typeof contextualTableBlock.tableBackground === "string"
                                    ? contextualTableBlock.tableBackground
                                    : null,
                              }}
                              testId="presentation-ribbon-table-cell-fill"
                              triggerVariant="large"
                              value={
                                contextualTableCellFormatState.fill?.value ===
                                "none"
                                  ? null
                                  : (contextualTableCellFormatState.fill
                                      ?.value ?? null)
                              }
                            />
                            {/* Table Design > Effects (PPT-A137): Cell Bevel on
                            the selected cells -- DrawingML's own preset, the
                            canvas showing the shading a bevel produces -- and
                            Shadow / Reflection on the whole table. */}
                            <PresentationTableEffectsMenu
                              bevel={
                                (contextualTableCellFormatState.bevel?.value ??
                                  "none") as TableCellBevelPreset | "none"
                              }
                              bevelMixed={
                                contextualTableCellFormatState.bevel?.mixed ??
                                false
                              }
                              controlId="PPTCTL.CONTEXT.TABLE.DESIGN.EFFECTS"
                              disabled={!contextualTableSelectionIsWritable}
                              effects={contextualTableBlock.tableEffects}
                              labels={{
                                button: t("presentation.tableRibbon.effects"),
                                cellBevel: t("presentation.ui.tableCellBevel.label"),
                                noCellBevel: t("presentation.ui.tableCellBevel.none"),
                                bevel: (preset) => t(TABLE_CELL_BEVEL_LABELS[preset]),
                                shadow: t("toolbar.shapeFormatOption.effect.shadow"),
                                noShadow: t("office.shapeEffects.none.shadow"),
                                shadowGroup: (group) =>
                                  t(`office.shadowEffects.group.${group}` as "office.shadowEffects.group.outer"),
                                shadowVariation: (group, id) =>
                                  t(`office.shadowEffects.${group}.${id}` as "office.shadowEffects.outer.bottomRight"),
                                reflection: t("toolbar.shapeFormatOption.effect.reflection"),
                                noReflection: t("office.shapeEffects.none.reflection"),
                                reflectionVariations: t("office.pictureEffects.variations.reflection"),
                                reflectionVariation: (id) =>
                                  t(
                                    `office.pictureEffects.reflection.${id}` as "office.pictureEffects.reflection.tightTouching",
                                  ),
                              }}
                              onBevel={(preset) =>
                                updateContextualTableCellStyle({
                                  bevel: preset === null ? undefined : { preset },
                                })
                              }
                              onEffect={(kind, value) =>
                                updateTableFrame(contextualTableBlock.id, (block) =>
                                  tableWithEffect(block, kind, value),
                                )
                              }
                              testId="presentation-ribbon-table-effects"
                            />
                            <PresentationTableCellColorMenu
                              allowNoFill={false}
                              automaticLabel={t("toolbar.textColorAutomatic")}
                              anchorClassName={
                                presentationChrome.backgroundPaletteAnchor
                              }
                              icon={
                                <span
                                  className="note-editor-office-toolbar__color-glyph"
                                  aria-hidden="true"
                                >
                                  <DocsHomeIcon name="fontColor" size={17} />
                                  <span
                                    className="note-editor-office-toolbar__color-underline"
                                    style={{
                                      backgroundColor:
                                        contextualTableCellFormatState.color
                                          ?.mixed
                                          ? "transparent"
                                          : (contextualTableCellFormatState
                                              .color?.value ?? "transparent"),
                                    }}
                                  />
                                </span>
                              }
                              label={t("presentation.tableCellColor")}
                              labels={tableCellColorPaletteLabels}
                              mixed={
                                contextualTableCellFormatState.color?.mixed ??
                                false
                              }
                              mixedLabel={t("toolbar.fontPickerMixed")}
                              noFillLabel={t("toolbar.textColorClear")}
                              onDismiss={() =>
                                setTableCellPaletteTarget(null)
                              }
                              onPick={(color) =>
                                pickContextualTableCellColor("color", color)
                              }
                              onAutomatic={() =>
                                pickContextualTableCellColor("color", null)
                              }
                              onToggle={() =>
                                setTableCellPaletteTarget((current) =>
                                  current === "color" ? null : "color",
                                )
                              }
                              open={tableCellPaletteTarget === "color"}
                              recents={tableCellRecentColors}
                              disabled={!contextualTableSelectionIsWritable}
                              testId="presentation-ribbon-table-cell-color"
                              formatPreview="color"
                              value={
                                contextualTableCellFormatState.color?.value ??
                                null
                              }
                            />
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
                              testId="presentation-ribbon-table-cell-font-family"
                            />
                            <RibbonValueCombobox
                              testId="presentation-ribbon-table-cell-font-size"
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
                          </div>
                          <div className="note-editor-ribbon-row">
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
                              testId="presentation-ribbon-table-cell-bold"
                              formatPreview={{
                                bold: !(
                                  contextualTableCellFormatState.bold?.value === true &&
                                  !contextualTableCellFormatState.bold?.mixed
                                ),
                              }}
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
                              testId="presentation-ribbon-table-cell-italic"
                              formatPreview={{
                                italic: !(
                                  contextualTableCellFormatState.italic?.value === true &&
                                  !contextualTableCellFormatState.italic?.mixed
                                ),
                              }}
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
                              testId="presentation-ribbon-table-cell-underline"
                              formatPreview={{
                                underline: !(
                                  contextualTableCellFormatState.underline?.value === true &&
                                  !contextualTableCellFormatState.underline?.mixed
                                ),
                              }}
                              variant="icon"
                            />
                            <OfficeHomeCommand
                              active={
                                contextualTableCellFormatState.strike?.mixed
                                  ? "mixed"
                                  : (contextualTableCellFormatState.strike
                                      ?.value ?? false)
                              }
                              icon={
                                <OfficeSemanticIcon
                                  semanticId="text.strikethrough"
                                  size={16}
                                />
                              }
                              disabled={!contextualTableSelectionIsWritable}
                              label={t("toolbar.strikethrough")}
                              onClick={() =>
                                updateContextualTableCellStyle({
                                  strike:
                                    contextualTableCellFormatState.strike
                                      ?.mixed ||
                                    !contextualTableCellFormatState.strike
                                      ?.value
                                      ? true
                                      : undefined,
                                })
                              }
                              testId="presentation-ribbon-table-cell-strike"
                              formatPreview={{
                                strike: !(
                                  contextualTableCellFormatState.strike?.value === true &&
                                  !contextualTableCellFormatState.strike?.mixed
                                ),
                              }}
                              variant="icon"
                            />
                            <OfficeHomeCommand
                              active={
                                contextualTableCellFormatState.superscript
                                  ?.mixed
                                  ? "mixed"
                                  : (contextualTableCellFormatState.superscript
                                      ?.value ?? false)
                              }
                              icon={
                                <OfficeSemanticIcon
                                  semanticId="text.superscript"
                                  size={16}
                                />
                              }
                              disabled={!contextualTableSelectionIsWritable}
                              label={t("toolbar.superscript")}
                              onClick={() =>
                                updateContextualTableCellStyle({
                                  superscript:
                                    contextualTableCellFormatState.superscript
                                      ?.mixed ||
                                    !contextualTableCellFormatState.superscript
                                      ?.value
                                      ? true
                                      : undefined,
                                  subscript: undefined,
                                })
                              }
                              testId="presentation-ribbon-table-cell-superscript"
                              variant="icon"
                            />
                            <OfficeHomeCommand
                              active={
                                contextualTableCellFormatState.subscript?.mixed
                                  ? "mixed"
                                  : (contextualTableCellFormatState.subscript
                                      ?.value ?? false)
                              }
                              icon={
                                <OfficeSemanticIcon
                                  semanticId="text.subscript"
                                  size={16}
                                />
                              }
                              disabled={!contextualTableSelectionIsWritable}
                              label={t("toolbar.subscript")}
                              onClick={() =>
                                updateContextualTableCellStyle({
                                  superscript: undefined,
                                  subscript:
                                    contextualTableCellFormatState.subscript
                                      ?.mixed ||
                                    !contextualTableCellFormatState.subscript
                                      ?.value
                                      ? true
                                      : undefined,
                                })
                              }
                              testId="presentation-ribbon-table-cell-subscript"
                              variant="icon"
                            />
                            <PresentationTableCellColorMenu
                              anchorClassName={
                                presentationChrome.backgroundPaletteAnchor
                              }
                              icon={
                                <span
                                  className="note-editor-office-toolbar__color-glyph"
                                  aria-hidden="true"
                                >
                                  <DocsHomeIcon name="highlight" size={16} />
                                  <span
                                    className="note-editor-office-toolbar__color-underline"
                                    style={{
                                      backgroundColor:
                                        contextualTableCellFormatState
                                          .highlight?.mixed
                                          ? "transparent"
                                          : (contextualTableCellFormatState
                                              .highlight?.value ??
                                            "transparent"),
                                    }}
                                  />
                                </span>
                              }
                              label={t("presentation.runHighlight")}
                              labels={tableCellColorPaletteLabels}
                              mixed={
                                contextualTableCellFormatState.highlight
                                  ?.mixed ?? false
                              }
                              mixedLabel={t("toolbar.fontPickerMixed")}
                              noFillLabel={t("toolbar.highlightClear")}
                              onDismiss={() =>
                                setTableCellPaletteTarget(null)
                              }
                              onPick={(color) =>
                                pickContextualTableCellColor(
                                  "highlight",
                                  color,
                                )
                              }
                              onToggle={() =>
                                setTableCellPaletteTarget((current) =>
                                  current === "highlight" ? null : "highlight",
                                )
                              }
                              open={tableCellPaletteTarget === "highlight"}
                              recents={tableCellRecentColors}
                              disabled={!contextualTableSelectionIsWritable}
                              testId="presentation-ribbon-table-cell-highlight"
                              value={
                                contextualTableCellFormatState.highlight
                                  ?.value ?? null
                              }
                            />
                            <PresentationTableCellChoiceMenu
                              defaultValue="left"
                              disabled={!contextualTableSelectionIsWritable}
                              icon={
                                <IconParagraphAlign
                                  home
                                  size={20}
                                  alignment={tableCellAlignIcon(
                                    contextualTableCellFormatState.align
                                      ?.value,
                                  )}
                                />
                              }
                              label={t("presentation.tableCellAlign")}
                              mixed={
                                contextualTableCellFormatState.align?.mixed ??
                                false
                              }
                              mixedLabel={t("toolbar.fontPickerMixed")}
                              value={
                                contextualTableCellFormatState.align?.value ??
                                "left"
                              }
                              options={CELL_ALIGNMENTS.map((value) => ({
                                value,
                                label: t(
                                  `presentation.tableAlign.${value}` as "presentation.tableAlign.left",
                                ),
                                icon:
                                  value === "left" ? (
                                    <AlignLeft size={16} />
                                  ) : value === "center" ? (
                                    <AlignHorizontalJustifyCenter size={16} />
                                  ) : (
                                    <AlignRight size={16} />
                                  ),
                              }))}
                              onSelect={(align) =>
                                updateContextualTableCellStyle({ align })
                              }
                              testId="presentation-ribbon-table-cell-align"
                            />
                            <PresentationTableCellChoiceMenu
                              defaultValue="top"
                              disabled={!contextualTableSelectionIsWritable}
                              icon={<IconAlignTextMenu home size={20} />}
                              label={t("presentation.tableCellVAlign")}
                              mixed={
                                contextualTableCellFormatState.verticalAlign
                                  ?.mixed ?? false
                              }
                              mixedLabel={t("toolbar.fontPickerMixed")}
                              value={
                                contextualTableCellFormatState.verticalAlign
                                  ?.value ?? "top"
                              }
                              options={VERTICAL_ANCHORS.map((value) => ({
                                value,
                                label: t(
                                  `presentation.tableVAlign.${value}` as "presentation.tableVAlign.top",
                                ),
                                icon:
                                  value === "top" ? (
                                    <AlignVerticalJustifyStart size={16} />
                                  ) : value === "middle" ? (
                                    <AlignVerticalJustifyCenter size={16} />
                                  ) : (
                                    <AlignVerticalJustifyEnd size={16} />
                                  ),
                              }))}
                              onSelect={(verticalAlign) =>
                                updateContextualTableCellStyle({
                                  verticalAlign:
                                    persistedVerticalAnchor(verticalAlign),
                                })
                              }
                              testId="presentation-ribbon-table-cell-valign"
                            />
                          </div>
                        </>
                      ) : (
                        <span
                          className="note-editor-office-toolbar__field-label"
                          data-testid="presentation-ribbon-table-cell-hint"
                        >
                          {t("presentation.tableCellHint")}
                        </span>
                      )}
                    </OfficeHomeRibbonGroup>
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.tableBorderGroup}
                      label={t("presentation.tableRibbon.borders")}
                      size="standard"
                      testId="presentation-ribbon-table-border-group"
                    >
                      {contextualTableFocus && contextualTableCellStyle ? (
                        <>
                          <div className="note-editor-ribbon-row">
                            <PresentationTableBordersMenu
                              disabled={!contextualTableSelectionIsWritable}
                              icon={<OfficeSemanticIcon semanticId="table.borders" size={24} />}
                              label={t("presentation.tableRibbon.borders")}
                              menuControlId="PPTCTL.CONTEXT.TABLE.DESIGN.BORDERS.MENU"
                              value={tableBorderTarget}
                              options={contextualTableBorderTargetOptions}
                              onSelect={applyDrawingPenToBorders}
                              testId="presentation-ribbon-table-cell-border-target"
                              styleSection={{
                                title: t("spreadsheet.borderLineStyle"),
                                value: tableBorderPen.style,
                                options: TABLE_CELL_BORDER_STYLES.map((style) => ({
                                  value: style,
                                  label:
                                    style === "none"
                                      ? t("toolbar.cellBorderNone")
                                      : style === "solid"
                                        ? t("presentation.shapeOutlineStyle.solid")
                                        : style === "dash"
                                          ? t("presentation.shapeOutlineStyle.dashed")
                                          : style === "dot"
                                            ? t("presentation.shapeOutlineStyle.dotted")
                                            : t("spreadsheet.borderStyle.double"),
                                })),
                                onSelect: (style) => {
                                  updateDrawingPen({ style: style as TableCellBorderStyle });
                                  armDrawingPen();
                                },
                              }}
                              colorSection={{
                                title: t("presentation.tableRibbon.penColor"),
                                value: tableBorderPen.color,
                                labels: tableCellColorPaletteLabels,
                                recents: tableCellRecentColors,
                                onPick: (color) => {
                                  const normalized = normalizeCellColor(color, undefined);
                                  if (!normalized || tableDrawAuthorityRef.current.readOnly) return;
                                  updateDrawingPen({ color: normalized });
                                  armDrawingPen();
                                  setTableCellRecentColors((current) =>
                                    [normalized, ...current.filter((entry) => entry !== normalized)].slice(0, 8),
                                  );
                                },
                              }}
                              marginSection={{
                                title: tableCellMarginMenuLabels.title,
                                value:
                                  contextualTableCellMarginMenuValue === "mixed" ||
                                  contextualTableCellMarginMenuValue === "custom"
                                    ? ""
                                    : contextualTableCellMarginMenuValue,
                                options: TABLE_CELL_MARGIN_PRESETS.map((preset) => ({
                                  value: preset,
                                  label: tableCellMarginMenuLabels[preset],
                                })),
                                customLabel: tableCellMarginMenuLabels.custom,
                                onSelect: (preset) =>
                                  updateTableCellStyle(
                                    contextualTableBlock.id,
                                    contextualTableFocus.row,
                                    contextualTableFocus.col,
                                    tableCellMarginPresetPatch(preset),
                                  ),
                                onCustom: () => setTableCellMarginsDialogOpen(true),
                              }}
                              drawTable={{
                                active: tableBorderPenMode === "draw",
                                label: t("presentation.tableRibbon.drawTable"),
                                onToggle: () => toggleTableBorderPen("draw"),
                              }}
                              gridlines={{
                                checked: tableGridlinesVisible,
                                label: t("toolbar.tableViewGridlines"),
                                onChange: (checked) => setTableGridlinesVisible(checked),
                              }}
                            />
                            {/* While a border pen is held these three controls
                            describe the PEN (WPS retargets 笔样式/笔颜色/笔画粗细
                            the moment Draw Table is armed); otherwise they keep
                            editing the current cell selection directly. */}
                            <PresentationTableCellChoiceMenu
                              defaultValue="solid"
                              disabled={!contextualTableSelectionIsWritable}
                              icon={
                                <OfficeSemanticIcon
                                  semanticId="format.lineStyle"
                                  size={20}
                                />
                              }
                              label={t("spreadsheet.borderLineStyle")}
                              mixed={
                                !tableBorderPenMode &&
                                (contextualTableBorderState?.style.mixed ??
                                  false)
                              }
                              mixedLabel={t("toolbar.fontPickerMixed")}
                              value={
                                tableBorderPenMode
                                  ? tableBorderPen.style
                                  : (contextualTableBorderState?.style.value ??
                                    "none")
                              }
                              options={TABLE_CELL_BORDER_STYLES.map(
                                (style) => ({
                                  value: style,
                                  label:
                                    style === "none"
                                      ? t("toolbar.cellBorderNone")
                                      : style === "solid"
                                        ? t(
                                            "presentation.shapeOutlineStyle.solid",
                                          )
                                        : style === "dash"
                                          ? t(
                                              "presentation.shapeOutlineStyle.dashed",
                                            )
                                          : style === "dot"
                                            ? t(
                                                "presentation.shapeOutlineStyle.dotted",
                                              )
                                            : t(
                                                "spreadsheet.borderStyle.double",
                                              ),
                                }),
                              )}
                              onSelect={selectContextualTableBorderStyle}
                              testId="presentation-ribbon-table-cell-border-style"
                            />
                            {



}
                            <RibbonValueCombobox
                              testId="presentation-ribbon-table-cell-border-width"
                              label={t("presentation.home.lineWeight")}
                              placeholder={t("toolbar.fontPickerMixed")}
                              value={tableBorderWidthDraft}
                              inputMode="decimal"
                              options={TABLE_CELL_BORDER_WIDTH_OPTIONS}
                              disabled={!contextualTableSelectionIsWritable}
                              mixed={
                                !tableBorderPenMode &&
                                (contextualTableBorderState?.width.mixed ??
                                  false)
                              }
                              mixedLabel={t("toolbar.fontPickerMixed")}
                              selectOnFocus
                              restoreMenuInvokerOnSelect
                              onDraftChange={(value) => {
                                setTableBorderWidthDraft(value);
                                setTableBorderWidthDirty(true);
                              }}
                              onFocusChange={(focused) => {
                                setTableBorderWidthFocused(focused);
                                if (focused) setTableBorderWidthDirty(false);
                              }}
                              onRememberSelection={() => undefined}
                              onCommit={commitContextualTableBorderWidth}
                            />
                          </div>
                          <div className="note-editor-ribbon-row">
                            <PresentationTableCellMarginMenu
                              disabled={!contextualTableSelectionIsWritable}
                              labels={tableCellMarginMenuLabels}
                              value={contextualTableCellMarginMenuValue}
                              onCustom={() =>
                                setTableCellMarginsDialogOpen(true)
                              }
                              onSelect={(preset) =>
                                updateTableCellStyle(
                                  contextualTableBlock.id,
                                  contextualTableFocus.row,
                                  contextualTableFocus.col,
                                  tableCellMarginPresetPatch(preset),
                                )
                              }
                              testId="presentation-ribbon-table-cell-margin"
                              triggerRef={tableCellMarginsTriggerRef}
                            />
                            {(
                              [
                                [
                                  "marginLeft",
                                  <PanelLeft key="margin-left" size={14} />,
                                  t("spreadsheet.borderEdge.left"),
                                ],
                                [
                                  "marginRight",
                                  <PanelRight key="margin-right" size={14} />,
                                  t("spreadsheet.borderEdge.right"),
                                ],
                                [
                                  "marginTop",
                                  <PanelTop key="margin-top" size={14} />,
                                  t("spreadsheet.borderEdge.top"),
                                ],
                                [
                                  "marginBottom",
                                  <PanelBottom key="margin-bottom" size={14} />,
                                  t("spreadsheet.borderEdge.bottom"),
                                ],
                              ] as const
                            ).map(([key, icon, directionLabel]) => {
                              const label = `${t("toolbar.cellMargin")}: ${directionLabel}`;
                              return (
                                <span
                                  key={key}
                                  className="note-editor-office-toolbar__control"
                                  title={label}
                                >
                                  <span aria-hidden="true">{icon}</span>
                                  <RibbonValueCombobox
                                    testId={`presentation-ribbon-table-cell-${key}`}
                                    label={label}
                                    placeholder={t("toolbar.cellMarginNone")}
                                    value={tableCellMarginDrafts[key]}
                                    inputMode="decimal"
                                    options={TABLE_CELL_MARGIN_VALUE_OPTIONS}
                                    className="note-editor-ribbon-value-combobox--font-size"
                                    disabled={
                                      !contextualTableSelectionIsWritable
                                    }
                                    mixed={
                                      contextualTableCellFormatState[key]
                                        ?.mixed ?? false
                                    }
                                    mixedLabel={t("toolbar.fontPickerMixed")}
                                    selectOnFocus
                                    restoreMenuInvokerOnSelect
                                    onDraftChange={(value) => {
                                      setTableCellMarginDrafts((drafts) => ({
                                        ...drafts,
                                        [key]: value,
                                      }));
                                      setTableCellMarginDirty(true);
                                    }}
                                    onFocusChange={(focused) => {
                                      setTableCellMarginFocused(
                                        focused ? key : null,
                                      );
                                      if (focused) {
                                        setTableCellMarginDirty(false);
                                      }
                                    }}
                                    onRememberSelection={() => undefined}
                                    onCommit={(value, options) =>
                                      commitContextualTableCellMargin(
                                        key,
                                        value,
                                        options,
                                      )
                                    }
                                  />
                                </span>
                              );
                            })}
                            <PresentationTableCellChoiceMenu
                              defaultValue="horizontal"
                              disabled={!contextualTableSelectionIsWritable}
                              icon={<OfficeSemanticIcon semanticId="table.cell.textDirection" size={20} />}
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
                              testId="presentation-ribbon-table-cell-text-direction"
                            />
                          </div>
                        </>
                      ) : (
                        <span className="note-editor-office-toolbar__field-label">
                          {t("presentation.tableCellHint")}
                        </span>
                      )}
                    </OfficeHomeRibbonGroup>
                    <PresentationTableWordArtGroup
                      className={presentationChrome.tableCellFormatGroup}
                      disabled={!contextualTableSelectionIsWritable || !contextualTableFocus}
                      labels={{
                        group: t("presentation.ui.shapeFormat.wordArtStyles"),
                        quickStyles: t("presentation.home.quickStyles"),
                        clearWordArt: t("presentation.tableRibbon.clearWordArt"),
                        textFill: t("presentation.ui.shapeFormat.textFill"),
                        textOutline: t("presentation.ui.shapeFormat.textOutline"),
                        textEffects: t("presentation.ui.shapeFormat.textEffects"),
                        noFill: t("presentation.home.noFill"),
                        noOutline: t("presentation.home.noOutline"),
                        mixed: t("toolbar.fontPickerMixed"),
                        effect: (effect) =>
                          t(`toolbar.shapeFormatOption.effect.${effect}` as "toolbar.shapeFormatOption.effect.none"),
                        palette: tableCellColorPaletteLabels,
                      }}
                      styles={deckWordArtStyles.map((style) => ({
                        id: style.id,
                        label: wordArtStyleLabel(style.id),
                        paint: style.paint,
                      }))}
                      styleId={(() => {
                        const item = contextualTableTextPaintMixed
                          ? null
                          : presentationWordArtStyleOf(contextualTableTextPaint);
                        return item ? `ppt-wordart-${item}` : null;
                      })()}
                      paint={contextualTableTextPaintMixed ? null : contextualTableTextPaint}
                      mixed={contextualTableTextPaintMixed}
                      recents={tableCellRecentColors}
                      quickStylesIcon={<DocsInsertIcon name="wordArt" size={24} />}
                      onStyle={(styleId) => {
                        const style = deckWordArtStyles.find((candidate) => candidate.id === styleId);
                        if (!style) return;
                        // The style's bold is part of it (PowerPoint reads the
                        // style back only with its `b`, L26).
                        writeContextualTableTextPaint(() => ({
                          textPaint: style.paint,
                          ...(style.bold ? { bold: true } : {}),
                        }));
                      }}
                      onClear={() => writeContextualTableTextPaint(() => ({ textPaint: undefined }))}
                      onFill={(color) =>
                        writeContextualTableTextPaint((paint) => ({
                          textPaint: patchOfficeTextPaint(paint, { fill: color ?? "none" }),
                        }))
                      }
                      onOutline={(color) =>
                        writeContextualTableTextPaint((paint) => ({
                          textPaint: patchOfficeTextPaint(paint, { outline: color ?? "none" }),
                        }))
                      }
                      onEffect={(effect) =>
                        writeContextualTableTextPaint((paint) => ({
                          textPaint: patchOfficeTextPaint(paint, { effect }),
                        }))
                      }
                    />
                    <OfficeHomeRibbonGroup
                      className={presentationChrome.tableBorderToolsGroup}
                      label={t("presentation.tableRibbon.drawBorders")}
                      size="standard"
                      testId="presentation-ribbon-table-border-tools-group"
                    >
                      <div className={presentationChrome.tableBorderToolsRow}>
                        <div
                          className={presentationChrome.tableDrawingSettings}
                          data-testid="presentation-ribbon-table-border-menu-row"
                        >
                          <OfficeBorderLineStyleMenu
                            key={`style:${pageId}:${contextualTableBlock.id}`}
                            disabled={autosave.readOnly || !contextualTableSelectionIsWritable}
                            testId="presentation-ribbon-table-border-pen-style-menu"
                            section={{
                              kind: "line-styles",
                              id: "style",
                              label: t("presentation.tableRibbon.penStyle"),
                              triggerTestId: "presentation-ribbon-table-border-pen-style",
                              value: (() => {
                                const value = tableBorderPen.style;
                                return value === "dash"
                                  ? "dashed"
                                  : value === "dot"
                                    ? "dotted"
                                    : value === "double"
                                      ? "double"
                                      : "thin";
                              })(),
                              options: (["thin", "dashed", "dotted", "double"] as const).map(
                                (id) => ({
                                  id,
                                  testId: `presentation-ribbon-table-border-pen-style-${id}`,
                                  label:
                                    id === "double"
                                      ? t("spreadsheet.borderStyle.double")
                                      : t(
                                          `presentation.shapeOutlineStyle.${id === "thin" ? "solid" : id}`,
                                        ),
                                }),
                              ),
                              onSelect: (value) =>
                                updateDrawingPen({
                                  style:
                                    value === "dashed"
                                      ? "dash"
                                      : value === "dotted"
                                        ? "dot"
                                        : value === "double"
                                          ? "double"
                                          : "solid",
                                }),
                            }}
                            renderTrigger={(open) => (
                              <OfficeHomeMenuTrigger
                                active={open}
                                disabled={
                                  autosave.readOnly || !contextualTableSelectionIsWritable
                                }
                                className={presentationChrome.tableDrawingValue}
                                icon={
                                  <span className={presentationChrome.tableDrawingLineValue}>
                                    <OfficeBorderLinePreview
                                      lineStyle={
                                        tableBorderPen.style === "dash"
                                          ? "dashed"
                                          : tableBorderPen.style === "dot"
                                            ? "dotted"
                                            : tableBorderPen.style === "double"
                                              ? "double"
                                              : "thin"
                                      }
                                    />
                                    <ChevronsUpDown size={12} />
                                  </span>
                                }
                                label={t("presentation.tableRibbon.penStyle")}
                                testId="presentation-ribbon-table-border-pen-style"
                                variant="icon"
                              />
                            )}
                          />
                          <OfficeBorderWeightMenu
                            key={`${pageId}:${contextualTableBlock.id}`}
                            disabled={autosave.readOnly || !contextualTableSelectionIsWritable}
                            label={t("presentation.tableRibbon.penWeight")}
                            unit={t("presentation.ui.table.pointsUnit")}
                            points={TABLE_CELL_BORDER_POINT_WEIGHTS}
                            valuePt={lineWeightDesignPxToPoints(tableBorderPen.width)}
                            onSelect={(points) => updateDrawingPen({ width: lineWeightPointsToDesignPx(points) })}
                            testIdPrefix="presentation-ribbon-table-border-pen-width"
                            renderTrigger={(open) => (
                              <OfficeHomeMenuTrigger
                                active={open}
                                disabled={
                                  autosave.readOnly || !contextualTableSelectionIsWritable
                                }
                                className={presentationChrome.tableDrawingValue}
                                icon={
                                  <span className={presentationChrome.tableDrawingLineValue}>
                                    <span>
                                      {Number(lineWeightDesignPxToPoints(tableBorderPen.width).toFixed(3))}{" "}
                                      {t("presentation.ui.table.pointsUnit")}
                                    </span>
                                    <OfficeBorderLinePreview
                                      sizePt={lineWeightDesignPxToPoints(tableBorderPen.width)}
                                    />
                                    <ChevronsUpDown size={12} />
                                  </span>
                                }
                                label={t("presentation.tableRibbon.penWeight")}
                                controlId="PPTCTL.CONTEXT.TABLE.DESIGN.DRAW_BORDERS.WEIGHT"
                                testId="presentation-ribbon-table-border-pen-width"
                                variant="icon"
                              />
                            )}
                          />
                          <PresentationTableCellColorMenu
                            allowNoFill={false}
                            anchorClassName={presentationChrome.backgroundPaletteAnchor}
                            displayLabel={t("presentation.tableRibbon.penColor")}
                            icon={
                              <span
                                aria-hidden="true"
                                style={
                                  {
                                    display: "inline-flex",
                                    "--mrai-office-icon-accent": tableBorderPen.color,
                                  } as CSSProperties
                                }
                              >
                                <OfficeSemanticIcon semanticId="format.lineColor" size={20} />
                              </span>
                            }
                            label={t("presentation.tableRibbon.penColor")}
                            labels={tableCellColorPaletteLabels}
                            menuControlId="PPTCTL.CONTEXT.TABLE.DESIGN.DRAW_BORDERS.PEN_COLOR"
                            mixed={false}
                            mixedLabel={t("toolbar.fontPickerMixed")}
                            noFillLabel={t("toolbar.cellBorderNone")}
                            onDismiss={() => setTableCellPaletteTarget(null)}
                            onPick={(color) => {
                              const normalized = normalizeCellColor(color, undefined);
                              if (!normalized || tableDrawAuthorityRef.current.readOnly) return;
                              updateDrawingPen({ color: normalized });
                              setTableCellRecentColors((current) =>
                                [
                                  normalized,
                                  ...current.filter((entry) => entry !== normalized),
                                ].slice(0, 8),
                              );
                              setTableCellPaletteTarget(null);
                            }}
                            onToggle={() =>
                              setTableCellPaletteTarget((current) =>
                                current === "border" ? null : "border",
                              )
                            }
                            open={tableCellPaletteTarget === "border"}
                            recents={tableCellRecentColors}
                            disabled={autosave.readOnly || !contextualTableSelectionIsWritable}
                            testId="presentation-ribbon-table-border-pen-color"
                            triggerVariant="compact"
                            value={tableBorderPen.color}
                          />
                        </div>
                        {/* Owner-approved desktop exception, native table geometry. */}
                        <OfficeHomeCommandRow testId="presentation-ribbon-table-border-pointer-tools">
                          <OfficeHomeCommand
                            active={tableBorderPenMode === "draw"}
                            variant="large"
                            disabled={autosave.readOnly}
                            icon={<IconTableDraw />}
                            onClick={() => toggleTableBorderPen("draw")}
                            controlId="PPTCTL.CONTEXT.TABLE.DESIGN.DRAW_TABLE"
                            label={t("presentation.tableRibbon.drawTable")}
                            testId="presentation-ribbon-table-border-pen"
                          />
                          <OfficeHomeCommand
                            active={tableBorderPenMode === "erase"}
                            variant="large"
                            disabled={
                              !contextualTableSelectionIsWritable || !contextualTableFocus
                            }
                            icon={<IconTableDraw erase />}
                            onClick={() => toggleTableBorderPen("erase")}
                            controlId="PPTCTL.CONTEXT.TABLE.DESIGN.ERASER"
                            label={t("presentation.tableRibbon.eraser")}
                            testId="presentation-ribbon-table-border-pen-eraser"
                          />
                        </OfficeHomeCommandRow>
                      </div>
                    </OfficeHomeRibbonGroup>
                  </>
  );
}
