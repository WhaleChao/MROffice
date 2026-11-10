"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { MAX_PICTURE_NAME_LENGTH } from "../../../../lib/spreadsheet/drawing-buttons";
import { type RefObject } from "react";
import { ArrowUpToLine, ChevronUp, ChevronDown, ArrowDownToLine } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { officeDialogParts as dialogParts } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { OfficeTaskPane } from "@mrnote/core/web/components/office-ribbon/OfficeTaskPane.tsx";
import { OfficeDockedPane } from "@mrnote/core/web/components/office-ribbon/OfficeTaskPaneDockHost.tsx";
import { OfficeSelectionPaneList, OfficeSelectionPaneRow } from "@mrnote/core/web/components/office-ribbon/OfficeSelectionPaneList.tsx";
import spreadsheetHomeStyles from "./SpreadsheetHomeRibbon.module.css";
import { drawingStackPosition } from "../../../../lib/spreadsheet/drawing-stack";
import { type SheetProtectionPermission } from "@mrnote/core/web/lib/spreadsheet/protection.ts";
import { type SpreadsheetObject, type SheetModel } from "./spreadsheetSheetModel";
import { spreadsheetDrawingDisplayName, type SpreadsheetDrawing } from "./spreadsheetObjectContext";
import { type SpreadsheetChart } from "@mrnote/core/web/lib/spreadsheet/chart.ts";
import { spreadsheetSelectionPaneIcon } from "./spreadsheetWindowHelpers";
import { SPREADSHEET_TASK_PANE_EDGES } from "./spreadsheetWindowConstants";
import { type DrawingStackEntry } from "../../../../lib/spreadsheet/drawing-stack";
import { type Dispatch, type SetStateAction } from "react";
import { type OfficeTaskPaneDockRegistry } from "@mrnote/core/web/components/office-ribbon/officeTaskPaneDock.ts";

/** What renderSpreadsheetSelectionPane reads from SpreadsheetWindow's render scope. */
export interface RenderSpreadsheetSelectionPaneInputs {
  /** The sheet's own charts: listed, shown and hidden with the objects (user
   * audit S-011). */
  charts: readonly SpreadsheetChart[];
  commitSelectionPaneRename: () => void;
  /** Excel's default name of a drawing without one ("矩形 1", "图表 1"). */
  drawingDefaultName: (drawing: SpreadsheetDrawing) => string;
  drawingStack: DrawingStackEntry[];
  reorderSheetChart: (chartId: string, move: "front" | "back" | "forward" | "backward") => void;
  reorderSheetObject: (objectId: string, move: "front" | "back" | "forward" | "backward") => void;
  selectedChartId: string | null;
  selectedObjectId: string | null;
  selectSessionChart: (chartId: string) => void;
  selectionPaneRef: RefObject<HTMLElement | null>;
  selectionPaneRename: { id: string; draft: string; } | null;
  selectSessionObject: (objectId: string) => boolean;
  setSelectionPaneOpen: Dispatch<SetStateAction<boolean>>;
  setSelectionPaneRename: Dispatch<SetStateAction<{ id: string; draft: string; } | null>>;
  setSheetObjectsHidden: (ids: ReadonlySet<string>, hidden: boolean, chartIds?: ReadonlySet<string>) => void;
  sheet: SheetModel;
  sheetObjectPaneWritable: (object: SpreadsheetObject, field: "hidden" | "name") => boolean;
  sheetObjects: SpreadsheetObject[];
  sheetProtectionAllowsRef: RefObject<(candidate: SheetModel | null | undefined, permission: SheetProtectionPermission) => boolean>;
  sheetTaskPaneDocks: { right: OfficeTaskPaneDockRegistry; };
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  uniqueSheetObjectIndex: (objectId: string) => number | null;
}

/** Moved verbatim from SpreadsheetWindow's JSX: a render function (not a component), so the element tree React reconciles is unchanged. */
export function renderSpreadsheetSelectionPane({
  charts,
  commitSelectionPaneRename,
  drawingDefaultName,
  drawingStack,
  reorderSheetChart,
  reorderSheetObject,
  selectedChartId,
  selectedObjectId,
  selectSessionChart,
  selectionPaneRef,
  selectionPaneRename,
  selectSessionObject,
  setSelectionPaneOpen,
  setSelectionPaneRename,
  setSheetObjectsHidden,
  sheet,
  sheetObjectPaneWritable,
  sheetObjects,
  sheetProtectionAllowsRef,
  sheetTaskPaneDocks,
  t,
  uniqueSheetObjectIndex,
}: RenderSpreadsheetSelectionPaneInputs) {
  // The sheet's own charts take the pane's writes while the sheet's objects
  // may be edited (protection's Edit objects).
  const chartsWritable = sheetProtectionAllowsRef.current(sheet, "editObjects");
  return (
        <OfficeDockedPane
          docks={sheetTaskPaneDocks}
          edges={SPREADSHEET_TASK_PANE_EDGES}
          pane="selection"
        >
          {/* Excel's Selection Pane (Home > Find & Select > Selection Pane):
              the shared Office task pane and object list Docs uses, docked
              with the workbook's other panes. The list reads front-most
              first; the rail moves an object through the stack, which is the
              sheet's object order. */}
          <OfficeTaskPane
            title={t("toolbar.selectionPane")}
            closeLabel={t("common.close")}
            onClose={() => setSelectionPaneOpen(false)}
            testIdPrefix="spreadsheet-selection"
            testId="spreadsheet-selection-pane"
            closeTestId="spreadsheet-selection-pane-close"
            paneRef={selectionPaneRef}
            dockRight={false}
            className={spreadsheetHomeStyles.chartFormatPaneFill}
          >
            {(() => {
              const paneIds = new Set(
                sheetObjects
                  .filter((object) => sheetObjectPaneWritable(object, "hidden"))
                  .map((object) => object.id),
              );
              const chartIds = chartsWritable ? new Set(charts.map((chart) => chart.id)) : new Set<string>();
              if (paneIds.size === 0 && chartIds.size === 0) return null;
              const paneHidden = [
                ...sheetObjects.filter((object) => paneIds.has(object.id)).map((object) => object.hidden === true),
                ...charts.filter((chart) => chartIds.has(chart.id)).map((chart) => chart.hidden === true),
              ];
              return (
                <span className={dialogParts.row}>
                  <button
                    type="button"
                    className={dialogParts.compactButton}
                    disabled={!paneHidden.some(Boolean)}
                    onClick={() => setSheetObjectsHidden(paneIds, false, chartIds)}
                    data-testid="spreadsheet-selection-pane-show-all"
                  >
                    {t("presentation.selectionPane.showAll")}
                  </button>
                  <button
                    type="button"
                    className={dialogParts.compactButton}
                    disabled={paneHidden.every(Boolean)}
                    onClick={() => setSheetObjectsHidden(paneIds, true, chartIds)}
                    data-testid="spreadsheet-selection-pane-hide-all"
                  >
                    {t("presentation.selectionPane.hideAll")}
                  </button>
                </span>
              );
            })()}
            <OfficeSelectionPaneList
              isEmpty={sheetObjects.length === 0 && charts.length === 0}
              emptyLabel={t("goTo.empty")}
              emptyTestId="spreadsheet-selection-pane-empty"
              listLabel={t("toolbar.selectionPane")}
            >
              {[
                ...sheetObjects.flatMap((object, index) =>
                  // A SmartArt item its layout does not draw is in the
                  // outline (the Text Pane), not on the sheet.
                  (object as { smartArtHidden?: unknown }).smartArtHidden === true
                    ? []
                    : [{
                        kind: "object" as const,
                        object,
                        index,
                        level: drawingStackPosition(drawingStack, "object", object.id),
                      }],
                ),
                ...charts.map((chart, index) => ({
                  kind: "chart" as const,
                  chart,
                  index,
                  level: drawingStackPosition(drawingStack, "chart", chart.id),
                })),
              ]
                // Front-most first, across pictures, shapes and charts: the
                // one stack Arrange moves them through.
                .sort((left, right) => right.level - left.level || right.index - left.index)
                .map((row) => {
                  const reorderActions = (
                    top: boolean,
                    bottom: boolean,
                    testSuffix: string,
                    onMove: (move: "front" | "back" | "forward" | "backward") => void,
                  ) => ({
                    label: t("toolbar.objectLayer"),
                    actions: (
                      [
                        ["front", "bringToFront", ArrowUpToLine, top],
                        ["forward", "bringForward", ChevronUp, top],
                        ["backward", "sendBackward", ChevronDown, bottom],
                        ["back", "sendToBack", ArrowDownToLine, bottom],
                      ] as const
                    ).map(([move, labelKey, Icon, disabled]) => ({
                      key: move,
                      label: t(`spreadsheet.${labelKey}`),
                      icon: <Icon size={13} aria-hidden />,
                      disabled,
                      testId: `spreadsheet-selection-pane-${move}-${testSuffix}`,
                      onClick: () => onMove(move),
                    })),
                  });
                  const drawingCount = sheetObjects.length + charts.length;
                  if (row.kind === "chart") {
                    const { chart, index } = row;
                    const chartLabel = spreadsheetDrawingDisplayName(chart, drawingDefaultName);
                    const hidden = chart.hidden === true;
                    const visibilityLabel = hidden
                      ? t("presentation.selectionPane.showObject")
                      : t("presentation.selectionPane.hideObject");
                    const renameId = `chart:${chart.id}`;
                    const top = row.level === drawingStack.length - 1;
                    const bottom = row.level <= 0;
                    return (
                      <OfficeSelectionPaneRow
                        key={renameId}
                        icon={spreadsheetSelectionPaneIcon({ type: "chart" })}
                        label={chartLabel}
                        selected={chart.id === selectedChartId}
                        hidden={hidden}
                        reorderable={chartsWritable && drawingCount > 1}
                        testId={`spreadsheet-selection-pane-chart-${index}`}
                        rowData={{ "data-chart-id": chart.id, "data-object-type": "chart" }}
                        onClick={() => {
                          if (!hidden) selectSessionChart(chart.id);
                        }}
                        onKeyDown={(event) => {
                          if (event.key === "F2" && chartsWritable) {
                            event.preventDefault();
                            setSelectionPaneRename({ id: renameId, draft: chartLabel });
                          }
                        }}
                        onLabelDoubleClick={() => {
                          if (chartsWritable) setSelectionPaneRename({ id: renameId, draft: chartLabel });
                        }}
                        rename={
                          selectionPaneRename?.id === renameId
                            ? {
                                draft: selectionPaneRename.draft,
                                inputLabel: t("presentation.selectionPane.objectName"),
                                testId: "spreadsheet-selection-pane-rename-input",
                                maxLength: MAX_PICTURE_NAME_LENGTH,
                                onChange: (draft) => setSelectionPaneRename({ id: renameId, draft }),
                                onCommit: commitSelectionPaneRename,
                                onCancel: () => setSelectionPaneRename(null),
                              }
                            : null
                        }
                        visibility={{
                          hidden,
                          disabled: !chartsWritable,
                          title: visibilityLabel,
                          ariaLabel: visibilityLabel,
                          testId: `spreadsheet-selection-pane-chart-visibility-${index}`,
                          onToggle: () => setSheetObjectsHidden(new Set(), !hidden, new Set([chart.id])),
                        }}
                        reorder={
                          chartsWritable && drawingCount > 1
                            ? reorderActions(top, bottom, `chart-${index}`, (move) => reorderSheetChart(chart.id, move))
                            : null
                        }
                      />
                    );
                  }
                  const { object, index } = row;
                  const unique = uniqueSheetObjectIndex(object.id) !== null;
                  // A picture says its caption or "图片 N"; any other drawing
                  // its name, its words, or Excel's default name ("矩形 1"),
                  // never its internal id (user audit S-010).
                  const objectLabel = spreadsheetDrawingDisplayName(object, drawingDefaultName);
                  const renameable =
                    unique && sheetObjectPaneWritable(object, "name");
                  const visibilityWritable =
                    unique && sheetObjectPaneWritable(object, "hidden");
                  const hidden = object.hidden === true;
                  const visibilityLabel = hidden
                    ? t("presentation.selectionPane.showObject")
                    : t("presentation.selectionPane.hideObject");
                  const reorderable =
                    unique &&
                    drawingCount > 1 &&
                    sheetProtectionAllowsRef.current(sheet, "editObjects");
                  const top = row.level === drawingStack.length - 1;
                  const bottom = row.level <= 0;
                  return (
                    <OfficeSelectionPaneRow
                      key={`${object.id}:${index}`}
                      icon={spreadsheetSelectionPaneIcon(object)}
                      label={objectLabel}
                      selected={unique && object.id === selectedObjectId}
                      hidden={hidden}
                      reorderable={reorderable}
                      testId={`spreadsheet-selection-pane-object-${index}`}
                      rowData={{
                        "data-object-id": unique ? object.id : undefined,
                        "data-object-type": object.type,
                      }}
                      onClick={() => {
                        // A hidden object stays listed but is not drawn, so
                        // there is nothing on the grid to select.
                        if (unique && !hidden) selectSessionObject(object.id);
                      }}
                      onKeyDown={(event) => {
                        if (event.key === "F2" && renameable) {
                          event.preventDefault();
                          setSelectionPaneRename({ id: object.id, draft: objectLabel });
                        }
                      }}
                      onLabelDoubleClick={() => {
                        if (renameable) {
                          setSelectionPaneRename({ id: object.id, draft: objectLabel });
                        }
                      }}
                      rename={
                        selectionPaneRename?.id === object.id
                          ? {
                              draft: selectionPaneRename.draft,
                              inputLabel: t("presentation.selectionPane.objectName"),
                              testId: "spreadsheet-selection-pane-rename-input",
                              maxLength: MAX_PICTURE_NAME_LENGTH,
                              onChange: (draft) =>
                                setSelectionPaneRename({ id: object.id, draft }),
                              onCommit: commitSelectionPaneRename,
                              onCancel: () => setSelectionPaneRename(null),
                            }
                          : null
                      }
                      visibility={{
                        hidden,
                        disabled: !visibilityWritable,
                        title: visibilityLabel,
                        ariaLabel: visibilityLabel,
                        testId: `spreadsheet-selection-pane-visibility-${index}`,
                        onToggle: () =>
                          setSheetObjectsHidden(new Set([object.id]), !hidden),
                      }}
                      reorder={
                        reorderable
                          ? reorderActions(top, bottom, String(index), (move) => reorderSheetObject(object.id, move))
                          : null
                      }
                    />
                  );
                })}
            </OfficeSelectionPaneList>
          </OfficeTaskPane>
        </OfficeDockedPane>
  );
}
