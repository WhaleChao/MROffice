"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { spreadsheetObjectHasPatchbackTransformIdentity } from "../../../../lib/spreadsheet/drawing-buttons";
import { useMemo, type RefObject } from "react";
import { type SpreadsheetNavigationSheetSnapshot } from "./SpreadsheetNavigationPane";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { normalizeNamedRanges, isListedNamedRange } from "@mrnote/core/web/lib/spreadsheet/named-ranges.ts";
import { normalizeDataTables } from "@mrnote/core/web/lib/spreadsheet/data-table.ts";
import { chartSourceRangeToA1 } from "@mrnote/core/web/lib/spreadsheet/chart-config.ts";
import { normalizeCharts } from "@mrnote/core/web/lib/spreadsheet/chart.ts";
import { type PivotConfig, pivotPreservedReasonKey } from "@mrnote/core/web/lib/spreadsheet/pivot.ts";
import { type SheetModel } from "./spreadsheetSheetModel";
import { columnLabel } from "./spreadsheetRangeA1";
import { spreadsheetPictureOrdinalMap, spreadsheetObjectHasPatchbackVisibilityIdentity, spreadsheetPictureDisplayName, isSpreadsheetInkObject, spreadsheetSheetDrawingNamer, spreadsheetDrawingKindLabel } from "./spreadsheetObjectContext";

/** What useSpreadsheetNavigationSheets reads from SpreadsheetWindow's render scope. */
export interface SpreadsheetNavigationSheetsInputs {
  homeZh: boolean;
  isSheetProtectionEffectiveRef: RefObject<(candidate: SheetModel | null | undefined) => boolean>;
  sheets: SheetModel[];
  spreadsheetNativeWriteRoute: "patchback" | "from-scratch";
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
}

/** Moved verbatim from SpreadsheetWindow (called at the same position, so hook order is unchanged). */
export function useSpreadsheetNavigationSheets({
  homeZh,
  isSheetProtectionEffectiveRef,
  sheets,
  spreadsheetNativeWriteRoute,
  t,
}: SpreadsheetNavigationSheetsInputs) {
  const navigationSheets = useMemo<
    readonly SpreadsheetNavigationSheetSnapshot[]
  >(
    () =>
      sheets
        .filter((candidate) => !candidate.hidden)
        .map((candidate) => {
          const objects = candidate.objects ?? [];
          const pictureOrdinals = spreadsheetPictureOrdinalMap(objects);
          const sheetCharts = normalizeCharts(candidate.charts);
          // The names the Selection Pane gives (user audit S-010): never an id.
          const defaultName = spreadsheetSheetDrawingNamer(objects, sheetCharts, t);
          const objectIds = new Set(objects.map((object) => object.id));
          const namedRanges = normalizeNamedRanges(candidate.namedRanges)
            .filter(isListedNamedRange)
            .map((range) => ({
              id: `range:${range.name.toLowerCase()}`,
              name: range.name,
              detail: range.ref,
            }));
          const tables = normalizeDataTables(candidate.dataTables, {
            rowCount: candidate.rowCount,
            columnCount: candidate.columnCount,
          }).map((table) => ({
            id: table.id,
            name: table.name,
            detail: `${columnLabel(table.range.c1)}${table.range.r1 + 1}:${columnLabel(table.range.c2)}${table.range.r2 + 1}`,
          }));
          const protectedReason =
            isSheetProtectionEffectiveRef.current(candidate)
              ? t("spreadsheet.protectionBlocked")
              : null;
          // A PivotTable kept as it came in says why it cannot be renamed.
          const pivotRenameAction = (pivot: PivotConfig) => ({
            id: "rename" as const,
            label: t("code.rename.submit"),
            disabledReason:
              protectedReason ??
              (pivot.native?.disposition === "preserve-only" || pivot.native?.blocker
                ? t(pivotPreservedReasonKey(pivot.native?.blocker))
                : null),
          });
          const visibilityAction = (
            hidden: boolean | undefined,
            disabledReason: string | null,
          ) => ({
            id: hidden ? ("show" as const) : ("hide" as const),
            label: hidden
              ? t("spreadsheet.view.show")
              : t("spreadsheet.uiCopy.hide"),
            disabledReason,
          });
          const chartActions = (
            disabledReason: string | null,
            deleteReason = disabledReason,
            visibilityDisabledReason = disabledReason,
            hidden?: boolean,
          ) => [
            {
              id: "rename" as const,
              label: t("code.rename.submit"),
              disabledReason,
            },
            visibilityAction(hidden, visibilityDisabledReason),
            {
              id: "delete" as const,
              label: t("spreadsheet.deleteObject"),
              destructive: true,
              disabledReason: deleteReason,
            },
          ];
          const chartSnapshots = [
            ...objects
              .filter((object) => object.type === "chart")
              .map((object) => {
                const nativeIdentity =
                  spreadsheetObjectHasPatchbackTransformIdentity(object);
                const disabledReason =
                  protectedReason ??
                  (spreadsheetNativeWriteRoute === "patchback" &&
                  !nativeIdentity
                    ? t("spreadsheet.uiCopy.thisImportedChartHasNoVerifiableNativeObjectName")
                      : null);
                const visibilityDisabledReason =
                  protectedReason ??
                  object.nativeVisibilityBlocker ??
                  (spreadsheetNativeWriteRoute === "patchback" &&
                  !spreadsheetObjectHasPatchbackVisibilityIdentity(object)
                    ? t("spreadsheet.uiCopy.thisImportedChartHasNoVerifiableNativeVisibilityIdentity")
                    : null);
                return {
                  id: object.id,
                  name:
                    object.name ||
                    object.title ||
                    object.caption ||
                    object.alt ||
                    defaultName(object),
                  detail: object.range || object.ranges?.join(", ") || null,
                  hidden: object.hidden,
                  actions: chartActions(
                    disabledReason,
                    disabledReason,
                    visibilityDisabledReason,
                    object.hidden,
                  ),
                };
              }),
            ...sheetCharts
              .filter((chart) => !objectIds.has(chart.id))
              .map((chart) => {
                const disabledReason =
                  protectedReason ??
                  (spreadsheetNativeWriteRoute === "patchback"
                    ? t("spreadsheet.uiCopy.thisNewlyAuthoredChartHasNoVerifiableOriginalPackage")
                    : null);
                return {
                  id: chart.id,
                  name: chart.name || chart.title || defaultName(chart),
                  detail: chartSourceRangeToA1(chart),
                  hidden: chart.hidden,
                  actions: chartActions(
                    disabledReason,
                    disabledReason,
                    disabledReason,
                    chart.hidden,
                  ),
                };
              }),
          ];
          const images = objects
            .filter((object) => object.type === "image")
            .map((object) => {
              const source =
                object.source &&
                typeof object.source === "object" &&
                !Array.isArray(object.source)
                  ? (object.source as Record<string, unknown>)
                  : null;
              const nativePart = [
                source?.part,
                source?.partName,
                source?.part_name,
              ].some(
                (value) =>
                  typeof value === "string" && value.trim().length > 0,
              );
              const protectedReason =
                isSheetProtectionEffectiveRef.current(candidate)
                  ? t("spreadsheet.protectionBlocked")
                  : null;
              const metadataReason =
                protectedReason ??
                (spreadsheetNativeWriteRoute === "patchback" && !nativePart
                  ? t("spreadsheet.uiCopy.thisImportedPictureHasNoVerifiableNativeNameIdentity")
                  : null);
              const deleteReason =
                protectedReason ??
                (spreadsheetNativeWriteRoute === "patchback" &&
                !spreadsheetObjectHasPatchbackTransformIdentity(object)
                  ? t("spreadsheet.uiCopy.thisImportedPictureHasNoVerifiableNativeObjectIdentity")
                    : null);
              const visibilityReason =
                protectedReason ??
                object.nativeVisibilityBlocker ??
                (spreadsheetNativeWriteRoute === "patchback" &&
                !spreadsheetObjectHasPatchbackVisibilityIdentity(object)
                  ? t("spreadsheet.uiCopy.thisImportedPictureHasNoVerifiableNativeVisibilityIdentity")
                  : null);
              return {
                id: object.id,
                name:
                  object.caption ||
                  spreadsheetPictureDisplayName(
                    pictureOrdinals.get(object) ?? 0,
                    t("contentControl.type.picture"),
                  ),
                hidden: object.hidden,
                actions: [
                  {
                    id: "rename" as const,
                    label: t("code.rename.submit"),
                    disabledReason: metadataReason,
                  },
                  visibilityAction(object.hidden, visibilityReason),
                  {
                    id: "delete" as const,
                    label: t("spreadsheet.deleteObject"),
                    destructive: true,
                    disabledReason: deleteReason,
                  },
                ],
              };
            });
          const shapes = objects
            .filter(
              (object) =>
                object.type === "shape" && !isSpreadsheetInkObject(object),
            )
            .map((object) => {
              const visibilityReason =
                protectedReason ??
                object.nativeVisibilityBlocker ??
                (spreadsheetNativeWriteRoute === "patchback" &&
                !spreadsheetObjectHasPatchbackVisibilityIdentity(object)
                  ? t("spreadsheet.uiCopy.thisImportedShapeHasNoVerifiableNativeVisibilityIdentity")
                  : null);
              return {
                id: object.id,
                name:
                  object.name ||
                  object.shape?.text ||
                  defaultName(object),
                detail: spreadsheetDrawingKindLabel(object, t),
                hidden: object.hidden,
                actions: [visibilityAction(object.hidden, visibilityReason)],
              };
            });
          return {
            id: candidate.id,
            name: candidate.name || candidate.id,
            namedRanges,
            tables,
            charts: chartSnapshots,
            pivotTables: candidate.pivot
              ? [
                  {
                    id: `pivot:${candidate.id}`,
                    name:
                      candidate.pivot.name ||
                      candidate.pivot.native?.pivotName ||
                      t("spreadsheet.pivotTable"),
                    detail: candidate.pivot.output || candidate.pivot.source,
                    actions: [pivotRenameAction(candidate.pivot)],
                  },
                  // The sheet's other PivotTables: listed, navigable and
                  // renamed the same way.
                  ...(candidate.morePivots ?? []).map((pivot, index) => ({
                    id: `pivot:${candidate.id}:${index + 1}`,
                    name:
                      pivot.name ||
                      pivot.native?.pivotName ||
                      t("spreadsheet.pivotTable"),
                    detail: pivot.native?.outputRef || pivot.output || pivot.source,
                    actions: [pivotRenameAction(pivot)],
                  })),
                ]
              : [],
            images,
            shapes,
          };
        }),
    [homeZh, sheets, spreadsheetNativeWriteRoute, t],
  );
  return {
    navigationSheets,
  };
}
