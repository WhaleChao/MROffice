"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useMemo, type RefObject } from "react";
import { SquarePlus, SquareMinus, ListTree, ListFilter } from "@mrnote/core/web/components/office-ribbon/icons/houseIcons.tsx";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { type OfficeStyleGalleryItem, type OfficeStyleGalleryGroup } from "@mrnote/core/web/components/office-ribbon/OfficeStyleGallery.tsx";
import { type SpreadsheetVisibleCommandReason } from "../../../../lib/spreadsheet/visible-command-authority";
import { normalizePivotChartOptions, type SpreadsheetPivotChartOptions } from "@mrnote/core/web/lib/spreadsheet/pivot-chart.ts";
import { type SpreadsheetChart, type SpreadsheetPivotChartSource } from "@mrnote/core/web/lib/spreadsheet/chart.ts";
import { PIVOT_STYLE_NAMES, type PivotConfig } from "@mrnote/core/web/lib/spreadsheet/pivot.ts";
import { builtInPivotCellPaint } from "../../../../lib/spreadsheet/pivot-style-recipes";
import { DropdownMenuSeparator } from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import { type SpreadsheetObject } from "./spreadsheetSheetModel";
import { SpreadsheetRibbonAction, SpreadsheetRibbonMenuAction, DropdownMenuCheckboxItem } from "./SpreadsheetRibbonControls";
import { spreadsheetPivotDesignRibbonGroups } from "./spreadsheetPivotDesignRibbonGroups";
import { type JSX } from "react";
import { type ResolvedSpreadsheetVisibleControl } from "../../../../lib/spreadsheet/visible-command-authority";

/** What useSpreadsheetPivotDesignState reads from SpreadsheetWindow's render scope. */
export interface SpreadsheetPivotDesignStateInputs {
  executePivotDesignMutation: (patch: Partial<PivotConfig>) => Promise<boolean>;
  executeVisibleCommand: (controlId: string) => Promise<boolean>;
  pivotAnalyzeRibbonGroups: { id: string; label: string; content: JSX.Element; }[];
  pivotChartContextActive: boolean;
  pivotChartFieldButtonsCommandRef: RefObject<SpreadsheetPivotChartOptions | null>;
  pivotConfig: PivotConfig | null;
  pivotContextActive: boolean;
  pivotPanelOpen: boolean;
  selectedChartDesignTarget: { readonly kind: "drawing"; readonly chart: SpreadsheetObject; } | { readonly kind: "native"; readonly chart: SpreadsheetChart; } | null;
  selectedPivotChartSource: SpreadsheetPivotChartSource | undefined;
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  visibleCommandActionProps: (controlId: string) => { controlId: string; disabled: boolean; disabledReason: string | undefined; hidden: boolean; onClick: () => Promise<boolean>; };
  visibleCommandDisclosureProps: (controlId: string) => { controlId: string; disabled: boolean; disabledReason: string | undefined; hidden: boolean; };
  visibleCommandReason: (reasonKey: SpreadsheetVisibleCommandReason | null) => string | undefined;
  visibleCommandState: (controlId: string) => ResolvedSpreadsheetVisibleControl;
  workbookThemeAccentColors: readonly string[];
}

/** Moved verbatim from SpreadsheetWindow (called at the same position, so hook order is unchanged). */
export function useSpreadsheetPivotDesignState({
  executePivotDesignMutation,
  executeVisibleCommand,
  pivotAnalyzeRibbonGroups,
  pivotChartContextActive,
  pivotChartFieldButtonsCommandRef,
  pivotConfig,
  pivotContextActive,
  pivotPanelOpen,
  selectedChartDesignTarget,
  selectedPivotChartSource,
  t,
  visibleCommandActionProps,
  visibleCommandDisclosureProps,
  visibleCommandReason,
  visibleCommandState,
  workbookThemeAccentColors,
}: SpreadsheetPivotDesignStateInputs) {
  // PivotChart Analyze: the selected PivotChart's Active Field, its
  // PivotTable's Filter, Data and Actions (while that table is on this
  // sheet), and Show/Hide with the Field List and the chart's Field Buttons.
  const selectedPivotChartOptions = normalizePivotChartOptions(
    (selectedChartDesignTarget?.chart as { pivotOptions?: unknown } | undefined)?.pivotOptions,
  );
  const runPivotChartFieldButtons = (patch: SpreadsheetPivotChartOptions) => {
    pivotChartFieldButtonsCommandRef.current = patch;
    void executeVisibleCommand("XLSCTL.CONTEXT.PIVOTCHART.FIELD_BUTTONS.MRAI.APPLY").finally(() => {
      pivotChartFieldButtonsCommandRef.current = null;
    });
  };
  const pivotChartFieldButtonsState = visibleCommandState(
    "XLSCTL.CONTEXT.PIVOTCHART.FIELD_BUTTONS.MRAI.APPLY",
  );
  const pivotChartAnalyzeRibbonGroups = selectedPivotChartSource
    ? [
        {
          id: "pivotchart-analyze-active-field",
          label: t("spreadsheet.pivotChartActiveField"),
          content: (
            <>
              <SpreadsheetRibbonAction
                {...visibleCommandActionProps("XLSCTL.CONTEXT.PIVOTCHART.EXPAND_FIELD.MRAI.APPLY")}
                controlOccurrenceId="expand-field"
                controlSurface="pivotchart-analyze"
                homeVariant="large"
                icon={<SquarePlus size={24} />}
                label={t("spreadsheet.pivotChartExpandField")}
                onClick={() => void executeVisibleCommand("XLSCTL.CONTEXT.PIVOTCHART.EXPAND_FIELD.MRAI.APPLY")}
                testId="spreadsheet-ribbon-pivotchart-expand-field"
              />
              <SpreadsheetRibbonAction
                {...visibleCommandActionProps("XLSCTL.CONTEXT.PIVOTCHART.COLLAPSE_FIELD.MRAI.APPLY")}
                controlOccurrenceId="collapse-field"
                controlSurface="pivotchart-analyze"
                homeVariant="large"
                icon={<SquareMinus size={24} />}
                label={t("spreadsheet.pivotChartCollapseField")}
                onClick={() => void executeVisibleCommand("XLSCTL.CONTEXT.PIVOTCHART.COLLAPSE_FIELD.MRAI.APPLY")}
                testId="spreadsheet-ribbon-pivotchart-collapse-field"
              />
            </>
          ),
        },
        ...pivotAnalyzeRibbonGroups.filter((group) =>
          ["pivot-analyze-filter", "pivot-analyze-data", "pivot-analyze-actions"].includes(group.id),
        ),
        {
          id: "pivotchart-analyze-show",
          label: t("spreadsheet.view.show"),
          content: (
            <>
              {pivotChartContextActive ? (
                <SpreadsheetRibbonAction
                  {...visibleCommandActionProps("XLSCTL.CONTEXT.PIVOT.FIELD_LIST")}
                  active={pivotPanelOpen}
                  controlOccurrenceId="pivotchart-field-list"
                  controlSurface="pivot-analyze"
                  homeVariant="large"
                  icon={<ListTree size={24} />}
                  label={t("spreadsheet.uiCopy.fieldList")}
                  onClick={() => executeVisibleCommand("XLSCTL.CONTEXT.PIVOT.FIELD_LIST")}
                  testId="spreadsheet-ribbon-pivotchart-fields"
                />
              ) : null}
              <SpreadsheetRibbonMenuAction
                {...visibleCommandDisclosureProps("XLSCTL.CONTEXT.PIVOTCHART.FIELD_BUTTONS_MENU")}
                controlOccurrenceId="field-buttons"
                controlSurface="pivotchart-analyze"
                homeVariant="large"
                icon={<ListFilter size={24} />}
                label={t("spreadsheet.pivotChartFieldButtons")}
                testId="spreadsheet-ribbon-pivotchart-field-buttons"
                menuClassName="docs-cupertino-menu w-max max-w-[300px]"
              >
                {(
                  [
                    ["filter", "spreadsheet.pivotChartShowButtons_filter"],
                    ["legend", "spreadsheet.pivotChartShowButtons_legend"],
                    ["axis", "spreadsheet.pivotChartShowButtons_axis"],
                    ["values", "spreadsheet.pivotChartShowButtons_values"],
                    ["expandCollapse", "spreadsheet.pivotChartShowButtons_expandCollapse"],
                  ] as const
                ).map(([kind, label]) => (
                  <DropdownMenuCheckboxItem
                    key={kind}
                    checked={selectedPivotChartOptions?.[kind] !== false}
                    disabled={!pivotChartFieldButtonsState.enabled || selectedPivotChartOptions?.visible === false}
                    data-control-id="XLSCTL.CONTEXT.PIVOTCHART.FIELD_BUTTONS.MRAI.APPLY"
                    data-control-occurrence-id={`show-${kind}`}
                    data-testid={`spreadsheet-ribbon-pivotchart-field-buttons-${kind}`}
                    onSelect={(event) => {
                      event.preventDefault();
                      runPivotChartFieldButtons({ [kind]: selectedPivotChartOptions?.[kind] === false });
                    }}
                  >
                    {t(label)}
                  </DropdownMenuCheckboxItem>
                ))}
                <DropdownMenuSeparator />
                <DropdownMenuCheckboxItem
                  checked={selectedPivotChartOptions?.visible === false}
                  disabled={!pivotChartFieldButtonsState.enabled}
                  data-control-id="XLSCTL.CONTEXT.PIVOTCHART.FIELD_BUTTONS.MRAI.APPLY"
                  data-control-occurrence-id="hide-all"
                  data-testid="spreadsheet-ribbon-pivotchart-field-buttons-hide-all"
                  onSelect={(event) => {
                    event.preventDefault();
                    runPivotChartFieldButtons({ visible: selectedPivotChartOptions?.visible === false });
                  }}
                >
                  {t("spreadsheet.pivotChartHideAllButtons")}
                </DropdownMenuCheckboxItem>
              </SpreadsheetRibbonMenuAction>
            </>
          ),
        },
      ]
    : [];
  const pivotDesignCommandState = visibleCommandState(
    "XLSCTL.CONTEXT.PIVOT.DESIGN.MRAI.APPLY",
  );
  const pivotDesignDisabledReason = visibleCommandReason(
    pivotDesignCommandState.reasonKey,
  );
  // The 84 swatch names are formatted per translator, not per render: the
  // gallery's items only change with its enabled state and reason.
  const pivotDesignEnabled = pivotDesignCommandState.enabled;
  const pivotStyleItems: readonly OfficeStyleGalleryItem<string>[] = useMemo(() =>
    PIVOT_STYLE_NAMES.map((styleName) => {
      const tone = styleName.includes("Light")
        ? "light"
        : styleName.includes("Medium")
          ? "medium"
          : "dark";
      return {
      id: styleName,
      // Excel names the swatch "PivotTable Style Light 16", not the OOXML id.
      label: t("spreadsheet.dynamicCopy.pivotStyle", {
        tone: t(`spreadsheet.dynamicCopy.tableStyleTone.${tone}`),
        index: styleName.replace(/^PivotStyle(?:Light|Medium|Dark)/, ""),
      }),
      groupId: tone,
      ...(!pivotDesignEnabled && pivotDesignDisabledReason
        ? { disabledReason: pivotDesignDisabledReason }
        : {}),
      };
    }), [pivotDesignDisabledReason, pivotDesignEnabled, t]);
  const pivotStyleGroups: readonly OfficeStyleGalleryGroup<string>[] = [
    { id: "light", label: t("spreadsheet.tableStylesLight"), itemIds: PIVOT_STYLE_NAMES.slice(0, 28) },
    { id: "medium", label: t("spreadsheet.uiCopy.tableStyleToneMedium"), itemIds: PIVOT_STYLE_NAMES.slice(28, 56) },
    { id: "dark", label: t("spreadsheet.tableStylesDark"), itemIds: PIVOT_STYLE_NAMES.slice(56, 84) },
  ];
  const pivotStyleQuickIds = [
    "PivotStyleLight16",
    "PivotStyleLight21",
    "PivotStyleMedium2",
    "PivotStyleMedium9",
    "PivotStyleMedium16",
    "PivotStyleDark2",
    "PivotStyleDark9",
  ];
  const renderPivotStylePreview = (item: OfficeStyleGalleryItem<string>) => {
    // The swatch is the style's own recipe: header row, a striped body row
    // over a plain one, the grand total row.
    const swatch = (element: Parameters<typeof builtInPivotCellPaint>[2]["elements"][number]["element"]) =>
      builtInPivotCellPaint(item.id, workbookThemeAccentColors, {
        elements: [{ element: "wholeTable" }, { element, top: true, bottom: true, left: true, right: true }],
      });
    const header = swatch("headerRow");
    const stripe = swatch("firstRowStripe");
    const plain = swatch("wholeTable");
    const total = swatch("totalRow");
    const band = (paint: ReturnType<typeof swatch>) => ({
      backgroundColor: paint.backgroundColor ?? "#ffffff",
      borderTop: paint.borderEdges?.top ? `1px solid ${paint.borderEdges.top.color}` : undefined,
      borderBottom: paint.borderEdges?.bottom ? `1px solid ${paint.borderEdges.bottom.color}` : undefined,
    });
    return (
      <span className="grid h-full w-full grid-rows-4 overflow-hidden rounded-[2px] border border-slate-300 bg-white">
        <span style={band(header)} />
        <span style={band(stripe)} />
        <span style={band(plain)} />
        <span style={band(total)} />
      </span>
    );
  };
  const {
    pivotDesignRibbonGroups,
  } = spreadsheetPivotDesignRibbonGroups({
    executePivotDesignMutation,
    pivotConfig,
    pivotContextActive,
    pivotDesignCommandState,
    pivotStyleGroups,
    pivotStyleItems,
    pivotStyleQuickIds,
    renderPivotStylePreview,
    t,
    visibleCommandDisclosureProps,
  });
  return {
    pivotChartAnalyzeRibbonGroups,
    pivotDesignRibbonGroups,
  };
}
