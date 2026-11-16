"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { type RefObject } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import type { OfficePictureFillSource } from "@mrnote/core/web/lib/office-picture-fill.ts";
import { officeShapeEffectMenuValue } from "@mrnote/core/web/lib/office-shape-effects.ts";
import { officeLineWeightPtToPx } from "@mrnote/core/web/lib/office-line-weight.ts";
import { type SpreadsheetVisibleCommandReason } from "../../../../lib/spreadsheet/visible-command-authority";
import { SPREADSHEET_SHAPE_RECENT_STORAGE_KEY, spreadsheetShapeKindFromOfficePreset, officeShapePresetForSpreadsheetKind, MAX_SHAPE_STROKE_WIDTH, type SpreadsheetPresetShapeKind } from "@mrnote/core/web/lib/spreadsheet/shapes.ts";
import { spreadsheetShapeLineEnds } from "../../../../lib/spreadsheet/shape-line-ends";
import { type OfficeShapeFormatAdapter, OFFICE_SHAPE_EFFECTS, type OfficeShapeAlignAction } from "@mrnote/core/web/components/office-ribbon/OfficeShapeFormatTab.tsx";
import { sameOfficeTextPaint, patchOfficeTextPaint, type OfficeTextPaint } from "@mrnote/core/web/lib/office-text-paint.ts";
import { officeTextPaintEffect } from "@mrnote/core/web/lib/document-wordart-style-adapter.ts";
import { OFFICE_WORDART_GALLERY_RECIPES } from "@mrnote/core/web/lib/office-wordart-gallery.ts";
import { officeLineWeightPxToPt } from "@mrnote/core/web/lib/office-line-weight.ts";
import { pushRecentColor } from "@mrnote/core/web/lib/spreadsheet/color-palette.ts";
import { themeColorGrid } from "@mrnote/core/web/lib/spreadsheet/theme-colors.ts";
import { type SheetModel, type SpreadsheetObject } from "./spreadsheetSheetModel";
import { type ShapeFormatWriterControlId, type ShapeRotationDialogAuthority } from "./spreadsheetMutationReceipts";
import { type SpreadsheetShapeStyle } from "@mrnote/core/web/lib/spreadsheet/shapes.ts";
import { type Dispatch, type SetStateAction } from "react";
import { type ResolvedSpreadsheetVisibleControl } from "../../../../lib/spreadsheet/visible-command-authority";
import { type OfficeShapeStyleRecipe } from "@mrnote/core/web/lib/office-shape-style-catalog.ts";
import { OfficePictureEffectsMenu } from "@mrnote/core/web/components/office-ribbon/OfficePictureEffectsMenu.tsx";
import { presentationEffectsMenuLabels } from "@mrnote/core/web/components/notebook/contents/presentation/presentationEffectsLabels.ts";
import { OFFICE_MENU_DASHES, type OfficeMenuDash } from "@mrnote/core/web/lib/office-stroke-dash.ts";
import { spreadsheetShapeOfficeEffects } from "../../../../lib/spreadsheet/shape-office-effects";
import { OFFICE_GLOW_ACCENTS, OFFICE_GLOW_SIZES_PT, officeGlowVariation } from "@mrnote/core/web/lib/office-picture-effects.ts";
import { OFFICE_TEXT_EFFECT_KINDS, officeTextPaintOfficeEffects, withOfficeTextEffect } from "@mrnote/core/web/lib/office-text-effects.ts";
import { OFFICE_MERGE_SHAPES_OPERATIONS, type OfficeMergeShapesOperation } from "@mrnote/core/web/lib/office-shape-merge-client.ts";

/** The Dashes entry a sheet shape's line shows checked (null: one the menu
 * does not list). */
function spreadsheetShapeMenuDash(style: SpreadsheetShapeStyle | null): OfficeMenuDash | null {
  const dash = style?.outlineDash;
  if (!dash) return "solid";
  return (OFFICE_MENU_DASHES as readonly string[]).includes(dash) ? (dash as OfficeMenuDash) : null;
}

/**
 * PowerPoint's Shape Format commands on the sheet's drawings, computed by
 * SpreadsheetWindow: Edit Shape ▸ Reroute Connectors, Merge Shapes, Reorder
 * Objects, Group ▸ Regroup and the selection count beside Arrange.
 */
/**
 * Where the worksheet's Format Shape / Format Picture pane opens: its page
 * (`tab`), and the section a command brings into view there (an Effects
 * submenu's "... Options": `effects-glow`, ...), `serial` telling a repeated
 * request from the one already shown.
 */
export interface SpreadsheetFormatPaneOpening {
  readonly tab: string;
  readonly section?: string;
  readonly serial?: number;
}

/** The pane at `tab` with `section` in view, a new request after `current`. */
export function spreadsheetFormatPaneAt(
  current: SpreadsheetFormatPaneOpening | null,
  tab: string,
  section: string,
): SpreadsheetFormatPaneOpening {
  return { tab, section, serial: (current?.serial ?? 0) + 1 };
}

export interface SpreadsheetShapeFormatPptParity {
  readonly reroute: { readonly enabled: boolean; readonly run: () => boolean };
  readonly merge: {
    readonly enabled: boolean;
    readonly twoSelected: boolean;
    readonly run: (op: OfficeMergeShapesOperation) => void;
  };
  readonly reorder: {
    readonly enabled: boolean;
    readonly overlappingEnabled: boolean;
    readonly open: (overlappingOnly: boolean) => void;
  };
  readonly regroup: { readonly enabled: boolean; readonly run: () => boolean };
  /** Ungroup that remembers the group for Regroup. */
  readonly ungroup: () => boolean;
  readonly selectionCount: number;
  /** Why no drawing on this sheet can change (read-only, protected, chart sheet). */
  readonly lockedReason: string | undefined;
}

/** What spreadsheetShapeFormatAdapter reads from SpreadsheetWindow's render scope. */
export interface SpreadsheetShapeFormatAdapterInputs {
  pptParity: SpreadsheetShapeFormatPptParity;
  /** Shape Fill > Picture...: the file picker for the selected shape's fill
   * picture, and the gate Fill & Line's command reports. */
  beginShapeFillPicture: (source?: OfficePictureFillSource) => void;
  shapeFillLineDisabledReason: string | undefined;
  activeShapeQuickStyleId: string | null;
  alignDrawingSelection: (action: OfficeShapeAlignAction, dryRun?: boolean) => boolean;
  canGroupSelection: boolean;
  canUngroupSelection: boolean;
  cmToPx: (cm: string) => number;
  executeShapeFormatMutation: (controlId: ShapeFormatWriterControlId, parameter: string | number, surface?: "visible" | "keyboard") => Promise<boolean | void>;
  groupDisabledReason: string | undefined;
  groupDrawingSelection: () => boolean;
  insertShape: (kind: SpreadsheetPresetShapeKind) => void;
  liveActiveSheetIndexRef: RefObject<number>;
  liveSheetsRef: RefObject<SheetModel[]>;
  objectSnapToGrid: boolean;
  objectSnapToShape: boolean;
  openShapeAltTextPane: () => boolean;
  pxToCm: (px: number) => string;
  recentColors: string[];
  /** Excel: the layer commands grey only when the sheet holds one drawing. */
  selectedObjectArrangeDisabled: boolean;
  selectedShapeObject: SpreadsheetObject | null;
  selectedShapeStyle: SpreadsheetShapeStyle | null;
  selectionIsGroup: boolean;
  setObjectSnapToGrid: Dispatch<SetStateAction<boolean>>;
  setObjectSnapToShape: Dispatch<SetStateAction<boolean>>;
  setRecentColors: Dispatch<SetStateAction<string[]>>;
  setSelectionPaneOpen: Dispatch<SetStateAction<boolean>>;
  setShapeFormatPane: Dispatch<SetStateAction<SpreadsheetFormatPaneOpening | null>>;
  setShapeLastFillColor: Dispatch<SetStateAction<string>>;
  setShapeLastKind: Dispatch<SetStateAction<SpreadsheetPresetShapeKind>>;
  setShapeLastStrokeColor: Dispatch<SetStateAction<string>>;
  setShapeRotationDraft: Dispatch<SetStateAction<{ authority: ShapeRotationDialogAuthority; value: string; error: string | null; silent?: boolean; } | null>>;
  setSheetFreeformTool: Dispatch<SetStateAction<"curve" | "freeform" | "scribble" | null>>;
  shapeAltTextCommandState: ResolvedSpreadsheetVisibleControl;
  shapeArrangeControlIdByMove: { readonly front: "XLSCTL.CONTEXT.SHAPE.BRING_TO_FRONT.MRAI.APPLY"; readonly forward: "XLSCTL.CONTEXT.SHAPE.BRING_FORWARD.MRAI.APPLY"; readonly backward: "XLSCTL.CONTEXT.SHAPE.SEND_BACKWARD.MRAI.APPLY"; readonly back: "XLSCTL.CONTEXT.SHAPE.SEND_TO_BACK.MRAI.APPLY"; };
  shapeChangeCommandState: ResolvedSpreadsheetVisibleControl;
  shapeEditPointsCommand: { onSelect: () => void; disabled?: undefined; disabledReason?: undefined; } | { onSelect: () => undefined; disabled: boolean; disabledReason: string | undefined; } | undefined;
  shapeFillColorCommandState: ResolvedSpreadsheetVisibleControl;
  shapeFormatDisabledReason: string | undefined;
  /** Why Rotate / Flip are greyed beyond Shape Format's own reason: a whole
   * SmartArt graphic, which Excel does not rotate. */
  shapeRotateDisabledReason: string | undefined;
  shapeLastFillColor: string;
  shapeLastStrokeColor: string;
  shapeLineEndsDisabledReason: string | undefined;
  shapeOutlineColorCommandState: ResolvedSpreadsheetVisibleControl;
  shapeOutlineWeightDisabledReason: string | undefined;
  shapeOutlineWeightPx: number;
  shapeQuickStyleCommandState: ResolvedSpreadsheetVisibleControl;
  shapeQuickStyleRecipes: readonly OfficeShapeStyleRecipe[];
  shapeShadowCommandState: ResolvedSpreadsheetVisibleControl;
  shapeSizePx: { width: number; height: number; } | null;
  shapeTextPaintBase: OfficeTextPaint;
  shapeTextPaintDisabledReason: string | undefined;
  shapeWriterUnavailable: string | undefined;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  ungroupDrawingSelection: () => boolean;
  visibleCommandReason: (reasonKey: SpreadsheetVisibleCommandReason | null) => string | undefined;
  workbookThemePaletteId: string;
  writeShapeTextPaint: (paint: OfficeTextPaint | null) => Promise<boolean | void>;
  /** WordArt on the words: the typed characters while they are typed (a
   * caret: its word), else every character of the selected shape, each
   * paint edited by `edit` (`style`: a whole WordArt style). */
  editShapeTextPaint: (edit: (paint: OfficeTextPaint) => OfficeTextPaint, style?: boolean) => Promise<boolean | void>;
  /** Text Fill on the words, as {@link editShapeTextPaint} picks them. */
  fillShapeText: (color: string | null) => Promise<boolean | void>;
  /** Shape Outline > Dashes and the full Shape Effects menu: their gates,
   * and the workbook theme's six accents (Glow's colours). */
  shapeOutlineDashDisabledReason: string | undefined;
  shapeOutlineSketchDisabledReason: string | undefined;
  shapeTextTransformDisabledReason: string | undefined;
  shapeEffectsDisabledReason: string | undefined;
  workbookThemeAccentColors: readonly string[];
}

/** Moved verbatim from SpreadsheetWindow (called at the same position, so hook order is unchanged). */
export function spreadsheetShapeFormatAdapter({
  pptParity,
  beginShapeFillPicture,
  shapeFillLineDisabledReason,
  shapeOutlineDashDisabledReason,
  shapeOutlineSketchDisabledReason,
  shapeTextTransformDisabledReason,
  shapeEffectsDisabledReason,
  workbookThemeAccentColors,
  activeShapeQuickStyleId,
  alignDrawingSelection,
  canGroupSelection,
  canUngroupSelection,
  cmToPx,
  executeShapeFormatMutation,
  groupDisabledReason,
  groupDrawingSelection,
  insertShape,
  liveActiveSheetIndexRef,
  liveSheetsRef,
  objectSnapToGrid,
  objectSnapToShape,
  openShapeAltTextPane,
  pxToCm,
  recentColors,
  selectedObjectArrangeDisabled,
  selectedShapeObject,
  selectedShapeStyle,
  selectionIsGroup,
  setObjectSnapToGrid,
  setObjectSnapToShape,
  setRecentColors,
  setSelectionPaneOpen,
  setShapeFormatPane,
  setShapeLastFillColor,
  setShapeLastKind,
  setShapeLastStrokeColor,
  setShapeRotationDraft,
  setSheetFreeformTool,
  shapeAltTextCommandState,
  shapeArrangeControlIdByMove,
  shapeChangeCommandState,
  shapeEditPointsCommand,
  shapeFillColorCommandState,
  shapeFormatDisabledReason,
  shapeRotateDisabledReason,
  shapeLastFillColor,
  shapeLastStrokeColor,
  shapeLineEndsDisabledReason,
  shapeOutlineColorCommandState,
  shapeOutlineWeightDisabledReason,
  shapeOutlineWeightPx,
  shapeQuickStyleCommandState,
  shapeQuickStyleRecipes,
  shapeShadowCommandState,
  shapeSizePx,
  shapeTextPaintDisabledReason,
  shapeWriterUnavailable,
  t,
  ungroupDrawingSelection,
  visibleCommandReason,
  workbookThemePaletteId,
  editShapeTextPaint,
  fillShapeText,
}: SpreadsheetShapeFormatAdapterInputs) {
  // The palettes show the workbook's own Theme Colors. The shared palette
  // reads a palette id only from its own short list and fell back to Office
  // 2013's accents under the workbook theme's names (S-125).
  const themeColumns = selectedShapeObject ? themeColorGrid(workbookThemePaletteId) : undefined;
  const shapeFormatAdapter: OfficeShapeFormatAdapter | null =
    selectedShapeObject && selectedShapeStyle
      ? {
          application: "spreadsheet",
          labels: {
            insertShapes: t("spreadsheet.chartFormatCopy.insert.insertShapes"),
            shapes: t("spreadsheet.insertShape"),
            editShape: t("spreadsheet.uiCopy.editShape"),
            changeShape: t("presentation.changeShape"),
            editPoints: t("spreadsheet.shapeFormat.editPoints"),
            editPointsReason: t("spreadsheet.shapeFormat.noSingleOutline"),
            rerouteConnectors: t("presentation.ui.shapeFormat.rerouteConnectors"),
            textBox: t("toolbar.textBox.horizontal"),
            textBoxHorizontal: t("toolbar.textBox.horizontal"),
            textBoxVertical: t("toolbar.textBox.vertical"),
            shapeStyles: t("presentation.home.shapeStyles"),
            moreShapeStyles: t("spreadsheet.uiCopy.moreShapeStyles"),
            previousStyles: t(
              "spreadsheet.chartFormatCopy.shape.previousStyles",
            ),
            nextStyles: t("spreadsheet.chartFormatCopy.shape.nextStyles"),
            shapeFill: t("presentation.shapeFormatFill"),
            fill: t("spreadsheet.shapeFill"),
            shapeOutline: t("presentation.shapeFormatOutline"),
            outline: t("spreadsheet.shapeStroke"),
            shapeEffects: t("spreadsheet.chartFormatCopy.shape.shapeEffects"),
            effects: t("spreadsheet.shapeFormat.effects"),
            effect: (effect) =>
              t(
                `toolbar.shapeFormatOption.effect.${effect}` as "toolbar.shapeFormatOption.effect.none",
              ),
            outlineLine: t("spreadsheet.shapeFormat.outlineLine"),
            arrows: t("officeArrows.arrows"),
            weight: t("presentation.home.lineWeight"),
            lineStyle: t("presentation.home.dashes"),
            lineStyleOption: (style) =>
              t(
                style === "solid"
                  ? "presentation.shapeOutlineStyle.solid"
                  : style === "dashed"
                    ? "presentation.shapeOutlineStyle.dashed"
                    : "presentation.shapeOutlineStyle.dotted",
              ),
            wordArtStyles: t(
              "spreadsheet.chartFormatCopy.wordArt.wordArtStyles",
            ),
            moreWordArtStyles: t(
              "spreadsheet.chartFormatCopy.wordArtGallery.more",
            ),
            textFill: t("spreadsheet.chartFormatCopy.wordArt.textFill"),
            textOutline: t("spreadsheet.chartFormatCopy.wordArt.textOutline"),
            textEffects: t("spreadsheet.chartFormatCopy.wordArt.textEffects"),
            text: t("spreadsheet.shapeFormat.text"),
            textDirection: t("spreadsheet.shapeFormat.textDirection"),
            textDirectionOption: (direction) =>
              t(
                `toolbar.shapeFormatOption.textDirection.${direction}` as "toolbar.shapeFormatOption.textDirection.horizontal",
              ),
            alignText: t("presentation.home.alignText"),
            alignTextOption: (align) =>
              t(
                `toolbar.shapeFormatOption.verticalAlign.${align}` as "toolbar.shapeFormatOption.verticalAlign.middle",
              ),
            createLink: t("spreadsheet.shapeFormat.createLink"),
            accessibility: t("spreadsheet.review.accessibility"),
            altText: t("spreadsheet.objectAlt"),
            arrange: t("spreadsheet.arrangeMenu"),
            position: t("spreadsheet.chartFormatCopy.arrange.position"),
            wrapText: t("spreadsheet.chartFormatCopy.arrange.wrapText"),
            bringForward: t("spreadsheet.bringForward"),
            bringToFront: t("spreadsheet.bringToFront"),
            sendBackward: t("spreadsheet.sendBackward"),
            sendToBack: t("spreadsheet.sendToBack"),
            selectionPane: t("spreadsheet.chartFormatCopy.arrange.selectionPane"),
            align: t("presentation.ribbon.align"),
            alignOption: (action) =>
              ({
                left: t("spreadsheet.dynamicCopy.align.left"),
                center: t("spreadsheet.dynamicCopy.align.center"),
                right: t("spreadsheet.dynamicCopy.align.right"),
                top: t("spreadsheet.dynamicCopy.align.top"),
                middle: t("spreadsheet.dynamicCopy.align.middle"),
                bottom: t("spreadsheet.dynamicCopy.align.bottom"),
                distributeHorizontal: t(
                  "spreadsheet.shapeFormat.distributeHorizontally",
                ),
                distributeVertical: t(
                  "spreadsheet.shapeFormat.distributeVertically",
                ),
              })[action],
            group: t("presentation.group.group"),
            ungroup: t("presentation.group.ungroup"),
            rotate: t("presentation.ui.shapeFormat.rotate"),
            rotateOption: (action) =>
              ({
                rotateRight90: t("spreadsheet.dynamicCopy.rotateRight90"),
                rotateLeft90: t("spreadsheet.dynamicCopy.rotateLeft90"),
                flipVertical: t("spreadsheet.shapeFormat.flipVertical"),
                flipHorizontal: t("spreadsheet.shapeFormat.flipHorizontal"),
              })[action],
            moreRotationOptions: t(
              "spreadsheet.shapeFormat.moreRotationOptions",
            ),
            alreadyAtFront: t("presentation.bringForwardAtFront"),
            alreadyAtBack: t("presentation.sendBackwardAtBack"),
            size: t("presentation.smartartContext.sizeGroup"),
            width: t("presentation.geometry.width"),
            height: t("presentation.geometry.height"),
            lockAspectRatio: t("presentation.tableRibbon.lockAspectRatio"),
            mixed: t("toolbar.fontPickerMixed"),
            formatPane: t("spreadsheet.shapeFormat.formatShape"),
            themeColors: t("spreadsheet.themeColors"),
            standardColors: t("spreadsheet.standardColors"),
            recentColors: t("spreadsheet.recentColors"),
            moreFillColors: t("spreadsheet.moreColors"),
            moreOutlineColors: t("spreadsheet.moreColors"),
            moreTextColors: t("spreadsheet.moreColors"),
            noFill: t("spreadsheet.noFill"),
            noOutline: t("spreadsheet.uiCopy.noOutline"),
          },
          insert: {
            disabled: Boolean(shapeFormatDisabledReason),
            disabledReason: shapeFormatDisabledReason ?? undefined,
            recentStorageKey: SPREADSHEET_SHAPE_RECENT_STORAGE_KEY,
            onInsert: (preset) => {
              insertShape(spreadsheetShapeKindFromOfficePreset(preset));
              return true;
            },
            onInsertTool: setSheetFreeformTool,
          },
          changeShape: {
            controlId: "XLSCTL.CONTEXT.SHAPE.CHANGE_SHAPE.MRAI.APPLY",
            disabled: !shapeChangeCommandState.enabled,
            disabledReason: visibleCommandReason(shapeChangeCommandState.reasonKey),
            recentStorageKey: SPREADSHEET_SHAPE_RECENT_STORAGE_KEY,
            selectedPreset: officeShapePresetForSpreadsheetKind(selectedShapeStyle.kind),
            tileIdentity: spreadsheetShapeKindFromOfficePreset,
            onChange: (preset) => {
              const kind = spreadsheetShapeKindFromOfficePreset(preset);
              setShapeLastKind(kind);
              void executeShapeFormatMutation(
                "XLSCTL.CONTEXT.SHAPE.CHANGE_SHAPE.MRAI.APPLY",
                kind,
              );
              return true;
            },
          },
          editPoints: shapeEditPointsCommand,
          // Excel's Edit Shape is a menu, as PowerPoint's: Change Shape ▸,
          // Edit Points, Reroute Connectors (the gallery used to fill the
          // menu and left Edit Points below a long scroll).
          changeShapeSubmenu: true,
          rerouteConnectors: pptParity.reroute.enabled
            ? { onSelect: () => void pptParity.reroute.run() }
            : {
                onSelect: () => undefined,
                disabled: true,
                disabledReason:
                  pptParity.lockedReason ??
                  t("presentation.ui.shapeFormat.rerouteConnectorsUnavailable"),
              },
          textBox: {
            disabled: Boolean(shapeFormatDisabledReason),
            disabledReason: shapeFormatDisabledReason ?? undefined,
            onInsert: () => insertShape("text-box"),
          },
          mergeShapes: {
            label: t("presentation.contextMenu.mergeShapes"),
            disabled: !pptParity.merge.enabled,
            disabledReason:
              pptParity.lockedReason ??
              (pptParity.merge.twoSelected
                ? t("presentation.ui.contextMenu.shapesCannotMerge")
                : t("presentation.ui.contextMenu.selectTwoShapes")),
            items: OFFICE_MERGE_SHAPES_OPERATIONS.map((op) => ({
              key: op,
              label: t(`presentation.mergeShapes.${op}`),
              onSelect: () => pptParity.merge.run(op),
            })),
          },
          styles: {
            quickCount: 7,
            controlId: "XLSCTL.CONTEXT.SHAPE.STYLE_GALLERY.MRAI.APPLY",
            recipes: shapeQuickStyleRecipes,
            value: activeShapeQuickStyleId,
            disabled: !shapeQuickStyleCommandState.enabled,
            disabledReason: visibleCommandReason(shapeQuickStyleCommandState.reasonKey),
            labelFor: (recipe) =>
                t("spreadsheet.sharedStyleCopy.shape.label", {
                  style: t(`spreadsheet.sharedStyleCopy.shape.rows.${recipe.row}`),
                  color: t(`spreadsheet.sharedStyleCopy.shape.columns.${recipe.column}`),
                }),
            onCommit: (r) =>
              void executeShapeFormatMutation(
                "XLSCTL.CONTEXT.SHAPE.STYLE_GALLERY.MRAI.APPLY",
                r.id,
              ),
          },
          fill: {
            controlId: "XLSCTL.CONTEXT.SHAPE.FILL_COLOR.MRAI.APPLY",
            value: selectedShapeStyle.fill,
            lastColor: shapeLastFillColor,
            paletteId: workbookThemePaletteId,
            themeColumns,
            recents: recentColors,
            disabled: Boolean(
              shapeFormatDisabledReason ?? !shapeFillColorCommandState.enabled,
            ),
            disabledReason:
              shapeFormatDisabledReason ??
              visibleCommandReason(shapeFillColorCommandState.reasonKey),
            onPick: (color) => {
              const next = color ?? "none";
              setShapeLastFillColor(next);
              void executeShapeFormatMutation(
                "XLSCTL.CONTEXT.SHAPE.FILL_COLOR.MRAI.APPLY",
                next,
              ).then((applied) => {
                if (applied && color)
                  setRecentColors((current) => pushRecentColor(current, color));
              });
            },
            // Office's row under the colours: Picture... fills the shape with a
            // picture stretched over it, as Format Shape's Picture fill does.
            extras: [
              {
                id: "picture",
                label: t("office.shapeFill.picture"),
                disabled: Boolean(shapeFillLineDisabledReason),
                disabledReason: shapeFillLineDisabledReason,
                onSelect: () => beginShapeFillPicture(),
              },
            ],
          },
          outline: {
            controlId: "XLSCTL.CONTEXT.SHAPE.OUTLINE_COLOR.MRAI.APPLY",
            value: selectedShapeStyle.stroke,
            lastColor: shapeLastStrokeColor,
            paletteId: workbookThemePaletteId,
            themeColumns,
            recents: recentColors,
            disabled: Boolean(
              shapeFormatDisabledReason ?? !shapeOutlineColorCommandState.enabled,
            ),
            disabledReason:
              shapeFormatDisabledReason ??
              visibleCommandReason(shapeOutlineColorCommandState.reasonKey),
            onPick: (color) => {
              const next = color ?? "none";
              setShapeLastStrokeColor(next);
              void executeShapeFormatMutation(
                "XLSCTL.CONTEXT.SHAPE.OUTLINE_COLOR.MRAI.APPLY",
                next,
              ).then((applied) => {
                if (applied && color)
                  setRecentColors((current) => pushRecentColor(current, color));
              });
            },
            weightPt: officeLineWeightPxToPt(shapeOutlineWeightPx),
            weight: {
              controlId: "XLSCTL.CONTEXT.SHAPE.OUTLINE_WEIGHT.MRAI.APPLY",
              disabled: Boolean(shapeOutlineWeightDisabledReason),
              disabledReason: shapeOutlineWeightDisabledReason,
            },
            onWeight: (pt) =>
              void executeShapeFormatMutation(
                "XLSCTL.CONTEXT.SHAPE.OUTLINE_WEIGHT.MRAI.APPLY",
                String(Math.min(officeLineWeightPtToPx(pt), MAX_SHAPE_STROKE_WIDTH)),
              ),
            lineStyle: "solid",
            lineStyleCapability: {
              controlId: "XLSCTL.CONTEXT.SHAPE.OUTLINE_DASH.MRAI.APPLY",
              disabled: Boolean(shapeOutlineDashDisabledReason),
              disabledReason: shapeOutlineDashDisabledReason,
            },
            onLineStyle: () => undefined,
            // Office's eight, as Excel lists them.
            dashes: {
              value: spreadsheetShapeMenuDash(selectedShapeStyle),
              onPick: (dash) =>
                void executeShapeFormatMutation("XLSCTL.CONTEXT.SHAPE.OUTLINE_DASH.MRAI.APPLY", dash),
              optionLabel: (dash) =>
                t(`presentation.officeGraphic.dash.${dash}` as "presentation.officeGraphic.dash.solid"),
            },
            // Shape Outline > Sketched, as Excel lists it.
            sketched: {
              label: t("office.outline.sketched"),
              value: selectedShapeStyle?.outlineSketch ?? "none",
              controlId: "XLSCTL.CONTEXT.SHAPE.OUTLINE_SKETCH.MRAI.APPLY",
              disabled: Boolean(shapeOutlineSketchDisabledReason),
              disabledReason: shapeOutlineSketchDisabledReason,
              onPick: (sketch) =>
                void executeShapeFormatMutation("XLSCTL.CONTEXT.SHAPE.OUTLINE_SKETCH.MRAI.APPLY", sketch),
              optionLabel: (sketch) => t(`office.outline.sketch.${sketch}` as "office.outline.sketch.none"),
            },
            lineEnds: spreadsheetShapeLineEnds(selectedShapeObject.shape),
            lineEndsCapability: {
              controlId: "XLSCTL.CONTEXT.SHAPE.LINE_ENDS.MRAI.APPLY",
              disabled: Boolean(shapeLineEndsDisabledReason),
              disabledReason: shapeLineEndsDisabledReason,
            },
            onLineEnds: (next) =>
              void executeShapeFormatMutation(
                "XLSCTL.CONTEXT.SHAPE.LINE_ENDS.MRAI.APPLY",
                JSON.stringify(next),
              ),
          },
          effects: {
            controlId: "XLSCTL.CONTEXT.SHAPE.EFFECT_PRESET.MRAI.APPLY",
            value: officeShapeEffectMenuValue(
              selectedShapeObject.effect,
              Boolean(selectedShapeObject.shadow),
            ),
            available: OFFICE_SHAPE_EFFECTS,
            unavailableReason: shapeWriterUnavailable,
            disabled: Boolean(shapeFormatDisabledReason) || !shapeShadowCommandState.enabled,
            disabledReason:
              shapeFormatDisabledReason ??
              (shapeShadowCommandState.enabled
                ? undefined
                : visibleCommandReason(shapeShadowCommandState.reasonKey)),
            onPick: (effect) =>
              void executeShapeFormatMutation(
                "XLSCTL.CONTEXT.SHAPE.EFFECT_PRESET.MRAI.APPLY",
                effect,
              ),
            // Office's whole Shape Effects menu (Preset, Shadow, Reflection,
            // Glow, Soft Edges, Bevel, 3-D Rotation), combinable, on the
            // object's `shapeEffects` -- the menu slides and pictures use.
            render: (trigger) => {
              const accents = {
                accent1: workbookThemeAccentColors[0] ?? "#4472c4",
                accent2: workbookThemeAccentColors[1] ?? "#ed7d31",
                accent3: workbookThemeAccentColors[2] ?? "#a5a5a5",
                accent4: workbookThemeAccentColors[3] ?? "#ffc000",
                accent5: workbookThemeAccentColors[4] ?? "#5b9bd5",
                accent6: workbookThemeAccentColors[5] ?? "#70ad47",
              };
              return (
                <OfficePictureEffectsMenu
                  effects={spreadsheetShapeOfficeEffects(selectedShapeObject ?? {}, accents.accent1)}
                  shadow={false}
                  accentColors={accents}
                  disabled={Boolean(shapeEffectsDisabledReason)}
                  disabledReason={shapeEffectsDisabledReason}
                  controlId="XLSCTL.CONTEXT.SHAPE.EFFECTS.MRAI.APPLY"
                  testId={trigger.testId}
                  displayLabel={trigger.displayLabel}
                  labels={presentationEffectsMenuLabels(t, trigger.label)}
                  onPick={(kind, value) =>
                    void executeShapeFormatMutation(
                      "XLSCTL.CONTEXT.SHAPE.EFFECTS.MRAI.APPLY",
                      JSON.stringify({ kind, value, accent1: accents.accent1 }),
                    )
                  }
                  // Each submenu's "... Options": Format Shape at its
                  // Effects page, that effect's section in view.
                  onEffectOptions={(section) =>
                    setShapeFormatPane((current) => spreadsheetFormatPaneAt(current, "effects", section))
                  }
                />
              );
            },
          },
          // WordArt Styles: the SHARED OfficeTextPaint on the label, the
          // site's gallery (the deck's twenty, as Excel offers them), written
          // through one receipt.
          wordArt: {
            value:
              OFFICE_WORDART_GALLERY_RECIPES.find((recipe) =>
                selectedShapeStyle.textPaint && sameOfficeTextPaint(selectedShapeStyle.textPaint, recipe.paint),
              )?.id ?? null,
            disabled: Boolean(shapeTextPaintDisabledReason),
            disabledReason: shapeTextPaintDisabledReason,
            onCommit: (recipe) => void editShapeTextPaint(() => recipe.paint, true),
          },
          textFill: {
            controlId: "XLSCTL.CONTEXT.SHAPE.TEXT_COLOR.MRAI.APPLY",
            value: selectedShapeStyle.textPaint?.fill ?? selectedShapeStyle.textColor ?? undefined,
            paletteId: workbookThemePaletteId,
            themeColumns,
            recents: recentColors,
            disabled: Boolean(shapeFormatDisabledReason),
            disabledReason: shapeFormatDisabledReason ?? undefined,
            onPick: (color) => void fillShapeText(color),
          },
          textOutline: {
            value: selectedShapeStyle.textPaint?.outline === "none" ? undefined : selectedShapeStyle.textPaint?.outline,
            paletteId: workbookThemePaletteId,
            themeColumns,
            recents: recentColors,
            disabled: Boolean(shapeTextPaintDisabledReason),
            disabledReason: shapeTextPaintDisabledReason,
            onPick: (color) =>
              void editShapeTextPaint((paint) =>
                patchOfficeTextPaint(paint, {
                  outline: color ?? "none",
                  outlineWidthPt: color ? 1 : 0,
                }),
              ),
          },
          // Text Effects > Transform: the label's preset warp, as Excel
          // offers it (a shape with no words takes none).
          textTransform: {
            value: selectedShapeStyle.textWarp?.preset ?? null,
            controlId: "XLSCTL.CONTEXT.SHAPE.TEXT_TRANSFORM.MRAI.APPLY",
            disabled: Boolean(shapeTextTransformDisabledReason) || !selectedShapeStyle.text,
            disabledReason: shapeTextTransformDisabledReason ?? t("spreadsheet.shapeFormat.typeTextFirst"),
            labels: {
              transform: t("office.textTransform.transform"),
              noTransform: t("office.textTransform.none"),
              followPath: t("office.textTransform.followPath"),
              warp: t("office.textTransform.warp"),
              name: (token) => t(`office.textTransform.name.${token}` as "office.textTransform.name.textPlain"),
            },
            onPick: (token) =>
              void executeShapeFormatMutation("XLSCTL.CONTEXT.SHAPE.TEXT_TRANSFORM.MRAI.APPLY", token ?? "none"),
          },
          textEffects: {
            value: officeTextPaintEffect(selectedShapeStyle.textPaint),
            available: ["none", "shadow", "reflection", "glow", "bevel"],
            unavailableReason: t("spreadsheet.shapeFormat.effectHasNoWriter"),
            disabled: Boolean(shapeTextPaintDisabledReason),
            disabledReason: shapeTextPaintDisabledReason,
            onPick: (effect) =>
              void editShapeTextPaint((paint) => patchOfficeTextPaint(paint, { effect })),
            // Excel's Text Effects galleries on the words' WordArt paint
            // (Shadow, Reflection, Glow, Bevel, 3-D Rotation, then
            // Transform), combinable.
            render: (trigger) => {
              const accents = {
                accent1: workbookThemeAccentColors[0] ?? "#4472c4",
                accent2: workbookThemeAccentColors[1] ?? "#ed7d31",
                accent3: workbookThemeAccentColors[2] ?? "#a5a5a5",
                accent4: workbookThemeAccentColors[3] ?? "#ffc000",
                accent5: workbookThemeAccentColors[4] ?? "#5b9bd5",
                accent6: workbookThemeAccentColors[5] ?? "#70ad47",
              };
              return (
                <OfficePictureEffectsMenu
                  effects={officeTextPaintOfficeEffects(
                    selectedShapeStyle.textPaint,
                    OFFICE_GLOW_SIZES_PT.flatMap((size) =>
                      OFFICE_GLOW_ACCENTS.map((accent) => officeGlowVariation(size, accent, accents[accent])),
                    ),
                  )}
                  shadow={false}
                  kinds={OFFICE_TEXT_EFFECT_KINDS}
                  extra={trigger.extra}
                  accentColors={accents}
                  disabled={Boolean(shapeTextPaintDisabledReason)}
                  disabledReason={shapeTextPaintDisabledReason}
                  testId={trigger.testId}
                  displayLabel={trigger.displayLabel}
                  labels={presentationEffectsMenuLabels(t, trigger.label)}
                  onPick={(kind, value) =>
                    void editShapeTextPaint((paint) => withOfficeTextEffect(paint, kind, value))
                  }
                  // Each submenu's "... Options": Format Shape > Text
                  // Options, where the words' effects and whole 3-D are
                  // edited, that effect's section in view.
                  onEffectOptions={(section) =>
                    setShapeFormatPane((current) => spreadsheetFormatPaneAt(current, "text", section))
                  }
                />
              );
            },
          },
          altText: {
            controlId: "XLSCTL.CONTEXT.SHAPE.ALT_TEXT.MRAI.APPLY",
            active: Boolean(selectedShapeObject.altText),
            disabled: Boolean(shapeFormatDisabledReason) || !shapeAltTextCommandState.enabled,
            disabledReason:
              shapeFormatDisabledReason ??
              visibleCommandReason(shapeAltTextCommandState.reasonKey),
            onOpen: () => void openShapeAltTextPane(),
          },
          zOrder: {
            disabled: Boolean(shapeFormatDisabledReason),
            disabledReason: shapeFormatDisabledReason ?? undefined,
            controlIdFor: (action) =>
              action === "bringForward"
                ? shapeArrangeControlIdByMove.forward
                : action === "bringToFront"
                  ? shapeArrangeControlIdByMove.front
                  : action === "sendBackward"
                    ? shapeArrangeControlIdByMove.backward
                    : shapeArrangeControlIdByMove.back,
            atBoundary: () => selectedObjectArrangeDisabled,
            onAction: (action) =>
              void executeShapeFormatMutation(
                action === "bringForward"
                  ? shapeArrangeControlIdByMove.forward
                  : action === "bringToFront"
                    ? shapeArrangeControlIdByMove.front
                    : action === "sendBackward"
                      ? shapeArrangeControlIdByMove.backward
                      : shapeArrangeControlIdByMove.back,
                action === "bringForward"
                  ? "forward"
                  : action === "bringToFront"
                    ? "front"
                    : action === "sendBackward"
                      ? "backward"
                      : "back",
              ),
          },
          selectionPane: {
            onSelect: () => setSelectionPaneOpen((open) => !open),
          },
          selectionCount: {
            count: pptParity.selectionCount,
            label: t("presentation.align.selectionCount", { count: pptParity.selectionCount }),
          },
          reorder: {
            label: t("presentation.arrange.reorderObjects"),
            overlappingLabel: t("presentation.arrange.reorderOverlapping"),
            disabled: !pptParity.reorder.enabled,
            disabledReason: pptParity.lockedReason ?? t("office.shapeFormat.reorderNeedsTwoOnSheet"),
            overlappingDisabled: !pptParity.reorder.overlappingEnabled,
            onReorder: () => pptParity.reorder.open(false),
            onReorderOverlapping: () => pptParity.reorder.open(true),
          },
          // Excel's Align menu: the edges and centres of two or more selected
          // objects, Distribute for three or more, then Snap to Grid / Snap
          // to Shape (the drag's snapping, and the arrows' step to the grid).
          alignObjects: {
            disabled: Boolean(shapeFormatDisabledReason),
            disabledReason: shapeFormatDisabledReason ?? undefined,
            isActionDisabled: (action) =>
              !alignDrawingSelection(action, true),
            onAction: (action) => void alignDrawingSelection(action),
            toggles: [
              {
                key: "snap-to-grid",
                checked: objectSnapToGrid,
                label: t("spreadsheet.shapeFormat.snapToGrid"),
                onChange: (checked) => setObjectSnapToGrid(checked),
                testId: "spreadsheet-shape-format-snap-to-grid",
              },
              {
                key: "snap-to-shape",
                checked: objectSnapToShape,
                label: t("spreadsheet.shapeFormat.snapToShape"),
                onChange: (checked) => setObjectSnapToShape(checked),
                testId: "spreadsheet-shape-format-snap-to-shape",
              },
            ],
          },
          group: {
            grouped: selectionIsGroup,
            canGroup: canGroupSelection,
            canUngroup: canUngroupSelection,
            disabled: !canGroupSelection && !canUngroupSelection,
            disabledReason: !canGroupSelection && !canUngroupSelection ? groupDisabledReason : undefined,
            onGroup: () => void groupDrawingSelection(),
            onUngroup: () => void pptParity.ungroup(),
            regroup: {
              label: t("office.shapeContextMenu.regroup"),
              disabled: !pptParity.regroup.enabled,
              disabledReason: pptParity.lockedReason ?? t("presentation.arrange.regroupUnavailable"),
              onSelect: () => void pptParity.regroup.run(),
            },
          },
          rotate: {
            disabled: Boolean(shapeFormatDisabledReason ?? shapeRotateDisabledReason),
            disabledReason: shapeFormatDisabledReason ?? shapeRotateDisabledReason ?? undefined,
            // The writers check the parameter: a quarter turn is "90", a
            // flip is "toggle" (the action name was refused silently).
            onAction: (action) =>
              void executeShapeFormatMutation(
                action === "rotateRight90"
                  ? "XLSCTL.CONTEXT.SHAPE.ROTATE_RIGHT.MRAI.APPLY"
                  : action === "rotateLeft90"
                    ? "XLSCTL.CONTEXT.SHAPE.ROTATE_LEFT.MRAI.APPLY"
                    : action === "flipVertical"
                      ? "XLSCTL.CONTEXT.SHAPE.FLIP_VERTICAL.MRAI.APPLY"
                      : "XLSCTL.CONTEXT.SHAPE.FLIP_HORIZONTAL.MRAI.APPLY",
                action === "rotateRight90" ? "90" : action === "rotateLeft90" ? "-90" : "toggle",
              ),
            exact: {
              disabled: Boolean(shapeFormatDisabledReason ?? shapeRotateDisabledReason),
              disabledReason: shapeFormatDisabledReason ?? shapeRotateDisabledReason ?? undefined,
              onSelect: () => {
                const liveWorkbook = liveSheetsRef.current;
                const liveIndex = liveActiveSheetIndexRef.current;
                const liveSheet = liveWorkbook[liveIndex];
                const liveObjects = liveSheet?.objects;
                const liveObject = liveObjects?.find(
                  (object) => object.id === selectedShapeObject.id,
                );
                if (!liveSheet || !liveObjects || !liveObject) return;
                setShapeRotationDraft({
                  authority: {
                    workbook: liveWorkbook,
                    activeSheetIndex: liveIndex,
                    sheet: liveSheet,
                    sheetId: liveSheet.id,
                    objectCollection: liveObjects,
                    object: liveObject,
                    objectId: liveObject.id,
                    selection: liveObject.id,
                  },
                  value: String(
                    Math.round(Number((liveObject as { rotation?: number }).rotation ?? 0)),
                  ),
                  error: null,
                });
              },
            },
          },
          formatPane: {
            onSelect: () =>
              setShapeFormatPane((current) => (current ? null : { tab: "shape" })),
          },
          size: {
            // Excel's Size group lists Shape Height above Shape Width, as
            // PowerPoint's does.
            heightFirst: true,
            controlId: "XLSCTL.CONTEXT.SHAPE.WIDTH.MRAI.APPLY",
            width: shapeSizePx ? pxToCm(shapeSizePx.width) : "",
            height: shapeSizePx ? pxToCm(shapeSizePx.height) : "",
            disabled: Boolean(shapeFormatDisabledReason),
            disabledReason: shapeFormatDisabledReason ?? undefined,
            onCommit: (axis, value) =>
              void executeShapeFormatMutation(
                axis === "width"
                  ? "XLSCTL.CONTEXT.SHAPE.WIDTH.MRAI.APPLY"
                  : "XLSCTL.CONTEXT.SHAPE.HEIGHT.MRAI.APPLY",
                String(Math.max(1, cmToPx(value))),
              ),
            lockAspect: {
              value: selectedShapeObject.lockAspectRatio === true,
              disabled: Boolean(shapeFormatDisabledReason),
              disabledReason: shapeFormatDisabledReason ?? undefined,
              onToggle: () =>
                void executeShapeFormatMutation(
                  "XLSCTL.CONTEXT.SHAPE.LOCK_ASPECT.MRAI.APPLY",
                  selectedShapeObject.lockAspectRatio === true ? "unlock" : "lock",
                ),
            },
          },
        }
      : null;
  return {
    shapeFormatAdapter,
  };
}
