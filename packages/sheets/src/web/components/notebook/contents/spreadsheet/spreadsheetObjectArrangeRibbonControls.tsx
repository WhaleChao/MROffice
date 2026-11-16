"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { type RefObject } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { OfficeHomeCommandStack } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import spreadsheetHomeStyles from "./SpreadsheetHomeRibbon.module.css";
import { SpreadsheetPictureIcon } from "./SpreadsheetPictureRibbonIcons";
import { type SheetProtectionPermission } from "@mrnote/core/web/lib/spreadsheet/protection.ts";
import { type SpreadsheetObject, type SheetModel } from "./spreadsheetSheetModel";
import { SpreadsheetRibbonAction } from "./SpreadsheetRibbonControls";
import { type DrawingStackEntry } from "../../../../lib/spreadsheet/drawing-stack";

/** What spreadsheetObjectArrangeRibbonControls reads from SpreadsheetWindow's render scope. */
export interface SpreadsheetObjectArrangeRibbonControlsInputs {
  drawingStack: DrawingStackEntry[];
  /**
   * Excel greys the four layer commands only when the sheet holds a single
   * drawing; at the top or bottom of the stack they stay on and do nothing.
   */
  drawingsArrangeable: boolean;
  reorderSheetObject: (objectId: string, move: "front" | "back" | "forward" | "backward") => void;
  /** A selected native chart moves through the same stack (Page Layout > Arrange). */
  reorderSheetChart: (chartId: string, move: "front" | "back" | "forward" | "backward") => void;
  selectedNativeChartId: string | null;
  selectedObject: SpreadsheetObject | null;
  selectedObjectStackIndex: number;
  sheet: SheetModel;
  sheetProtectionAllowsRef: RefObject<(candidate: SheetModel | null | undefined, permission: SheetProtectionPermission) => boolean>;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
}

/** Moved verbatim from SpreadsheetWindow (called at the same position, so hook order is unchanged). */
export function spreadsheetObjectArrangeRibbonControls({
  drawingStack,
  drawingsArrangeable,
  reorderSheetObject,
  reorderSheetChart,
  selectedNativeChartId,
  selectedObject,
  selectedObjectStackIndex,
  sheet,
  sheetProtectionAllowsRef,
  t,
}: SpreadsheetObjectArrangeRibbonControlsInputs) {
  const nativeChartStackIndex = selectedNativeChartId
    ? drawingStack.findIndex((entry) => entry.kind === "chart" && entry.id === selectedNativeChartId)
    : -1;
  const targetStackIndex = selectedObject ? selectedObjectStackIndex : nativeChartStackIndex;
  const arrangeSelected = (move: "front" | "back" | "forward" | "backward") => {
    if (selectedObject) reorderSheetObject(selectedObject.id, move);
    else if (selectedNativeChartId && nativeChartStackIndex >= 0) reorderSheetChart(selectedNativeChartId, move);
  };
  const stackDisabled = targetStackIndex < 0 || !drawingsArrangeable;
  const objectArrangeRibbonControls = (
    compact = true,
    options?: {
      disabledReason?: string;
      controlIds?: Partial<
        Record<"front" | "forward" | "backward" | "back", string>
      >;
      onArrange?: (
        move: "front" | "forward" | "backward" | "back",
      ) => void;
    },
  ) => {
    const commands = {
      front: (
          <SpreadsheetRibbonAction
            homeVariant={compact ? "compact" : "large"}
            controlId={options?.controlIds?.front}
            disabled={
              Boolean(options?.disabledReason) ||
              !sheetProtectionAllowsRef.current(sheet, "editObjects") ||
              stackDisabled
            }
            disabledReason={
              !sheetProtectionAllowsRef.current(sheet, "editObjects")
                ? t("spreadsheet.protectionBlocked")
                : options?.disabledReason
            }
            icon={<SpreadsheetPictureIcon name="bringToFront" />}
            label={t("spreadsheet.bringToFront")}
            onClick={() =>
              options?.onArrange
                ? options.onArrange("front")
                : arrangeSelected("front")
            }
            testId="spreadsheet-ribbon-arrange-front"
          />
      ),
      forward: (
          <SpreadsheetRibbonAction
            homeVariant={compact ? "compact" : "large"}
            controlId={options?.controlIds?.forward}
            disabled={
              Boolean(options?.disabledReason) ||
              !sheetProtectionAllowsRef.current(sheet, "editObjects") ||
              stackDisabled
            }
            disabledReason={
              !sheetProtectionAllowsRef.current(sheet, "editObjects")
                ? t("spreadsheet.protectionBlocked")
                : options?.disabledReason
            }
            icon={<SpreadsheetPictureIcon name="bringForward" />}
            label={t("spreadsheet.bringForward")}
            onClick={() =>
              options?.onArrange
                ? options.onArrange("forward")
                : arrangeSelected("forward")
            }
            testId="spreadsheet-ribbon-arrange-forward"
          />
      ),
      backward: (
          <SpreadsheetRibbonAction
            homeVariant={compact ? "compact" : "large"}
            controlId={options?.controlIds?.backward}
            disabled={
              Boolean(options?.disabledReason) ||
              !sheetProtectionAllowsRef.current(sheet, "editObjects") ||
              stackDisabled
            }
            disabledReason={
              !sheetProtectionAllowsRef.current(sheet, "editObjects")
                ? t("spreadsheet.protectionBlocked")
                : options?.disabledReason
            }
            icon={<SpreadsheetPictureIcon name="sendBackward" />}
            label={t("spreadsheet.sendBackward")}
            onClick={() =>
              options?.onArrange
                ? options.onArrange("backward")
                : arrangeSelected("backward")
            }
            testId="spreadsheet-ribbon-arrange-backward"
          />
      ),
      back: (
          <SpreadsheetRibbonAction
            homeVariant={compact ? "compact" : "large"}
            controlId={options?.controlIds?.back}
            disabled={
              Boolean(options?.disabledReason) ||
              !sheetProtectionAllowsRef.current(sheet, "editObjects") ||
              stackDisabled
            }
            disabledReason={
              !sheetProtectionAllowsRef.current(sheet, "editObjects")
                ? t("spreadsheet.protectionBlocked")
                : options?.disabledReason
            }
            icon={<SpreadsheetPictureIcon name="sendToBack" />}
            label={t("spreadsheet.sendToBack")}
            onClick={() =>
              options?.onArrange
                ? options.onArrange("back")
                : arrangeSelected("back")
            }
            testId="spreadsheet-ribbon-arrange-back"
          />
      ),
    };
    // Excel's Arrange commands, two to a column: the jumps to the top and
    // bottom of the stack, then the one-step moves. A single row of four
    // compact commands floated mid-band and read as a loose strip.
    return compact ? (
      <>
        <OfficeHomeCommandStack className={spreadsheetHomeStyles.viewSmallCommands}>
          {commands.front}
          {commands.back}
        </OfficeHomeCommandStack>
        <OfficeHomeCommandStack className={spreadsheetHomeStyles.viewSmallCommands}>
          {commands.forward}
          {commands.backward}
        </OfficeHomeCommandStack>
      </>
    ) : (
      <>
        {commands.front}
        {commands.forward}
        {commands.backward}
        {commands.back}
      </>
    );
  };
  return {
    objectArrangeRibbonControls,
  };
}
