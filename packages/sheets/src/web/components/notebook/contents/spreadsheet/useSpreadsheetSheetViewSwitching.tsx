"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useCallback } from "react";
import { type MergeSpan } from "@mrnote/core/web/lib/spreadsheet/merge-cells.ts";
import { sheetViewAuthoringBlocker, sheetViewStateOfSheet, rebuildNamedSheetViewState, applySheetViewFilters, wholeSheetViewState, snapshotSheetViewFilters, sheetViewSessionTables, createNamedSheetViewId, sheetViewNameFor, type SpreadsheetNamedSheetView, ensureSheetViewSessionTables, sheetViewStateHasFilters, defaultNamedSheetViewRef, sheetViewsRenamed } from "../../../../lib/spreadsheet/sheet-views";
import { type SheetModel } from "./spreadsheetSheetModel";
import { type SheetViewDialogState, type SheetViewSession } from "./spreadsheetSheetViews";
import { type Dispatch, type SetStateAction } from "react";

/** What useSpreadsheetSheetViewSwitching reads from SpreadsheetWindow's render scope. */
export interface SpreadsheetSheetViewSwitchingInputs {
  captureSheetViewReturnFocus: () => void;
  commitSheetViewRuntime: (nextSheet: SheetModel, nextSession: SheetViewSession | null, persistMetadata: boolean) => boolean;
  mergeSpans: MergeSpan[];
  setSheetViewDialog: Dispatch<SetStateAction<SheetViewDialogState | null>>;
  sheet: SheetModel;
  sheetViewSession: SheetViewSession | null;
  sheetViewSynthesizedRef: () => string;
  sheetViewUnavailable: (reason: string) => void;
}

/** Moved verbatim from SpreadsheetWindow (called at the same position, so hook order is unchanged). */
export function useSpreadsheetSheetViewSwitching({
  captureSheetViewReturnFocus,
  commitSheetViewRuntime,
  mergeSpans,
  setSheetViewDialog,
  sheet,
  sheetViewSession,
  sheetViewSynthesizedRef,
  sheetViewUnavailable,
}: SpreadsheetSheetViewSwitchingInputs) {
  const switchSheetView = useCallback(
    (viewId: string | null) => {
      const currentSession =
        sheetViewSession?.sheetId === sheet.id ? sheetViewSession : null;
      if (viewId && currentSession?.activeViewId === viewId) return true;
      const target = viewId
        ? (sheet.namedSheetViews ?? []).find((view) => view.id === viewId)
        : null;
      if (
        viewId &&
        (!target ||
          target.nativeBlocker ||
          sheetViewAuthoringBlocker(sheet) ||
          ((target.sortState || target.tables?.some((table) => table.sortState)) &&
            mergeSpans.length > 0))
      ) {
        sheetViewUnavailable("unsupported native view");
        return false;
      }
      const defaultState =
        currentSession?.defaultState ?? sheetViewStateOfSheet(sheet, null);
      let nextSheet = sheet;
      const persistMetadata = false;
      if (
        currentSession &&
        currentSession.activeViewId !== "temporary"
      ) {
        const state = sheetViewStateOfSheet(sheet, currentSession);
        const namedSheetViews = (sheet.namedSheetViews ?? []).map((view) =>
          view.id === currentSession.activeViewId
            ? rebuildNamedSheetViewState(
                view,
                state,
                currentSession.filterId,
                currentSession.ref,
              )
            : view,
        );
        nextSheet = { ...sheet, namedSheetViews };
      }
      nextSheet = applySheetViewFilters(
        nextSheet,
        target ? wholeSheetViewState(target) : defaultState,
      );
      if (!target && currentSession) {
        const restored: SheetModel = { ...nextSheet };
        delete restored.autoFilterId;
        delete restored.autoFilterRef;
        delete restored.namedSheetViewsOwnsAutoFilter;
        if (currentSession.defaultAutoFilterId) {
          restored.autoFilterId = currentSession.defaultAutoFilterId;
        }
        if (currentSession.defaultAutoFilterRef) {
          restored.autoFilterRef = currentSession.defaultAutoFilterRef;
        }
        if (currentSession.defaultOwnsAutoFilter) {
          restored.namedSheetViewsOwnsAutoFilter = true;
        }
        nextSheet = restored;
      }
      const defaultAutoFilterId = currentSession
        ? currentSession.defaultAutoFilterId
        : sheet.autoFilterId;
      const defaultAutoFilterRef = currentSession
        ? currentSession.defaultAutoFilterRef
        : sheet.autoFilterRef;
      const defaultOwnsAutoFilter = currentSession
        ? currentSession.defaultOwnsAutoFilter
        : sheet.namedSheetViewsOwnsAutoFilter;
      const ownsAutoFilter = currentSession
        ? currentSession.ownsAutoFilter
        : sheet.namedSheetViewsOwnsAutoFilter === true ||
          (!sheet.autoFilterId && !sheet.autoFilterRef)
          ? true
          : undefined;
      return commitSheetViewRuntime(
        nextSheet,
        target
          ? {
              sheetId: sheet.id,
              activeViewId: target.id,
              sharedState: wholeSheetViewState(snapshotSheetViewFilters(target)),
              defaultState,
              ...(defaultAutoFilterId ? { defaultAutoFilterId } : {}),
              ...(defaultAutoFilterRef ? { defaultAutoFilterRef } : {}),
              ...(defaultOwnsAutoFilter
                ? { defaultOwnsAutoFilter: true }
                : {}),
              ...(target.sortState ? { sortState: target.sortState } : {}),
              ...(target.tables ? { tables: sheetViewSessionTables(target) } : {}),
              ...(ownsAutoFilter ? { ownsAutoFilter: true } : {}),
              filterId:
                target.filterId ??
                sheet.autoFilterId ??
                createNamedSheetViewId(),
              ref:
                target.ref ??
                sheet.autoFilterRef ??
                sheetViewSynthesizedRef(),
            }
          : null,
        persistMetadata,
      );
    },
    [
      sheetViewSynthesizedRef,
      commitSheetViewRuntime,
      mergeSpans.length,
      sheet,
      sheetViewSession,
      sheetViewUnavailable,
    ],
  );

  const openKeepSheetViewDialog = useCallback(() => {
    const session =
      sheetViewSession?.sheetId === sheet.id ? sheetViewSession : null;
    if (!session || session.activeViewId !== "temporary") return false;
    captureSheetViewReturnFocus();
    const names = new Set(
      (sheet.namedSheetViews ?? []).map((view) => view.name.toLocaleLowerCase()),
    );
    let index = (sheet.namedSheetViews?.length ?? 0) + 1;
    let name = `View ${index}`;
    while (names.has(name.toLocaleLowerCase())) {
      index += 1;
      name = `View ${index}`;
    }
    setSheetViewDialog({ mode: "keep", name });
    return true;
  }, [captureSheetViewReturnFocus, sheet, sheetViewSession]);

  const openSheetViewOptions = useCallback(() => {
    const views = sheet.namedSheetViews ?? [];
    const selected =
      views.find(
        (view) =>
          sheetViewSession?.sheetId === sheet.id &&
          sheetViewSession.activeViewId === view.id,
      ) ?? views[0];
    if (!selected) return false;
    captureSheetViewReturnFocus();
    setSheetViewDialog({
      mode: "options",
      selectedId: selected.id,
      name: selected.name,
    });
    return true;
  }, [captureSheetViewReturnFocus, sheet, sheetViewSession]);

  // One name commit for both the dialog and the ribbon's name box: keep a
  // temporary view under `name`, or rename the selected named one. Excel for
  // the web accepts the name straight from the Sheet View box on Enter, so the
  // rule lives here once and the callers only decide how to show a refusal.
  const commitSheetViewName = useCallback(
    (
      request: SheetViewDialogState,
    ): { ok: true } | { ok: false; error: "invalid" | "duplicate" | "blocked" } => {
    const views = sheet.namedSheetViews ?? [];
    const named = sheetViewNameFor(
      views,
      request.name,
      request.mode === "keep" ? undefined : request.selectedId,
    );
    if ("error" in named) return { ok: false, error: named.error };
    const { name } = named;
    if (request.mode === "keep") {
      const session =
        sheetViewSession?.sheetId === sheet.id ? sheetViewSession : null;
      if (!session || session.activeViewId !== "temporary")
        return { ok: false, error: "blocked" };
      const view: SpreadsheetNamedSheetView = {
        id: createNamedSheetViewId(),
        name,
      };
      const { tables: tableStates, ...state } = sheetViewStateOfSheet(
        sheet,
        ensureSheetViewSessionTables(session, sheet),
      );
      // A Table the view filters or sorts keeps its own nsvFilter, with or
      // without a worksheet AutoFilter association.
      const keptTables = (tableStates ?? []).filter((table) => table.filterId);
      if (keptTables.length > 0) view.tables = keptTables;
      const hasRuntimeState = Boolean(
        state.sortState || sheetViewStateHasFilters(state),
      );
      const hasPersistedAssociation = Boolean(
        session.filterId &&
          (sheet.autoFilterId === session.filterId ||
            (sheet.namedSheetViews ?? []).some(
              (candidate) => candidate.filterId === session.filterId,
            )),
      );
      const attachFilterAssociation = Boolean(
        session.filterId && (hasRuntimeState || hasPersistedAssociation),
      );
      if (attachFilterAssociation && session.filterId) {
        view.filterId = session.filterId;
        view.ref = session.ref ?? sheet.autoFilterRef ?? defaultNamedSheetViewRef(sheet);
        Object.assign(view, state);
      }
      const nextSheet = {
        ...sheet,
        ...(attachFilterAssociation && session.filterId
          ? {
              autoFilterId: session.filterId,
              autoFilterRef:
                session.ref ??
                sheet.autoFilterRef ??
                defaultNamedSheetViewRef(sheet),
              ...(session.ownsAutoFilter
                ? { namedSheetViewsOwnsAutoFilter: true as const }
                : {}),
            }
          : {}),
        namedSheetViews: [...views, view],
      };
      if (
        commitSheetViewRuntime(
          nextSheet,
          {
            sheetId: session.sheetId,
            activeViewId: view.id,
            sharedState: wholeSheetViewState(snapshotSheetViewFilters(view)),
            defaultState: session.defaultState,
            ...(view.tables ? { tables: sheetViewSessionTables(view) } : {}),
            ...(session.defaultAutoFilterId
              ? { defaultAutoFilterId: session.defaultAutoFilterId }
              : {}),
            ...(session.defaultAutoFilterRef
              ? { defaultAutoFilterRef: session.defaultAutoFilterRef }
              : {}),
            ...(session.defaultOwnsAutoFilter
              ? { defaultOwnsAutoFilter: true as const }
              : {}),
            ...(attachFilterAssociation && session.filterId
              ? {
                  filterId: session.filterId,
                  ref:
                    session.ref ??
                    sheet.autoFilterRef ??
                    defaultNamedSheetViewRef(sheet),
                  ...(session.sortState
                    ? { sortState: session.sortState }
                    : {}),
                  ...(session.ownsAutoFilter
                    ? { ownsAutoFilter: true as const }
                    : {}),
                  defaultAutoFilterId:
                    session.defaultAutoFilterId ?? session.filterId,
                  defaultAutoFilterRef:
                    session.defaultAutoFilterRef ??
                    session.ref ??
                    sheet.autoFilterRef ??
                    defaultNamedSheetViewRef(sheet),
                  ...(session.ownsAutoFilter
                    ? { defaultOwnsAutoFilter: true as const }
                    : {}),
                }
              : {}),
          },
          true,
        )
      ) {
        return { ok: true };
      }
      return { ok: false, error: "blocked" };
    }
    const nextViews = sheetViewsRenamed(views, request.selectedId, name);
    if (!nextViews) return { ok: false, error: "blocked" };
    return commitSheetViewRuntime(
      { ...sheet, namedSheetViews: nextViews },
      sheetViewSession?.sheetId === sheet.id ? sheetViewSession : null,
      true,
    )
      ? { ok: true }
      : { ok: false, error: "blocked" };
    },
    [commitSheetViewRuntime, sheet, sheetViewSession],
  );
  return {
    commitSheetViewName,
    openSheetViewOptions,
    switchSheetView,
  };
}
