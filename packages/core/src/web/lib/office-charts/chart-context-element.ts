/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A chart's right-click menu follows the element it lands on, in Word, Excel
 * and PowerPoint alike (PPT-C010 / PPT-C011): on a series Reset to Match
 * Style resets that series, Add Data Labels labels it and Format Data
 * Series... opens the Format pane on it; on the legend or the chart title
 * Delete takes that element away, Reset to Match Style clears its own
 * formatting, Font... formats its text ("right-click the chart element ...
 * and select Font", Microsoft Support "Change the chart text font") and
 * Format Legend... / Format Chart Title... opens its page. An element menu has
 * no Cut / Copy / Paste, Group or z-order rows (they belong to the chart as an
 * object), nor Format Chart Area; a series' menu has no Font (Office's
 * neither).
 *
 * The element is what the picture reports: a series by the series it drew
 * (`OfficeChartPicture`'s right-click hit), the legend and title by the
 * parts the picture marks (`officeChartPicturePartAt`). An axis title is not
 * offered: which model axis a drawn axis title stands for depends on the
 * chart's orientation, which only the chart engine knows.
 *
 * One rule for the three editors; each host gives its own writer (a command
 * over the chart's current model) and its own Format pane opener.
 */
import type { ChartBlockAttrsV2 } from "../../components/console/editor/extensions/chartBlockShared";
import type { DocumentChartAreaStyleKey } from "../../components/console/editor/native-document/DocumentChartFormatRibbon";
import type {
  OfficeChartContextCommand,
  OfficeChartContextCommandId,
} from "../../components/office-ribbon/OfficeChartContextMenu";
import {
  chartDataLabelStatementOf,
  chartWithDataLabelOverride,
  chartWithElementIntent,
} from "./chart-elements";
import type { OfficeChartTextElement } from "./chart-element-text";
import {
  officeChartHitSeries,
  officeChartPicturePartAt,
  type OfficeChartHit,
  type OfficeChartPicturePart,
} from "./chart-hit";
import {
  presentationChartAreaStyle,
  setPresentationChartAreaStyle,
} from "../presentation-chart-desktop";

export type OfficeChartContextElement =
  | { readonly kind: "series"; readonly hit: OfficeChartHit }
  | { readonly kind: "part"; readonly part: OfficeChartPicturePart };

/**
 * The element a right click on a drawn chart landed on: the legend or the
 * title the picture marks under the pointer, else the series its picture
 * reported for this right click (`seriesHit`); undefined on anything else
 * (the chart area, the plot, an axis), whose menu is the chart's own.
 */
export function officeChartContextElementAt(
  root: Element | null | undefined,
  clientX: number,
  clientY: number,
  seriesHit: OfficeChartHit | null | undefined,
): OfficeChartContextElement | undefined {
  const part = root ? officeChartPicturePartAt(root, clientX, clientY) : null;
  if (part === "legend" || part === "title") return { kind: "part", part };
  return seriesHit && !part ? { kind: "series", hit: seriesHit } : undefined;
}

/** The rows of the chart-as-object menu an element's menu leaves out. */
const OBJECT_ROWS: readonly OfficeChartContextCommandId[] = [
  "cut",
  "copy",
  "paste",
  "group",
  "ungroup",
  "bringToFront",
  "bringForward",
  "sendToBack",
  "sendBackward",
  "formatChartArea",
];

export interface OfficeChartElementMenu {
  readonly commands: Partial<Record<OfficeChartContextCommandId, OfficeChartContextCommand>>;
  readonly omit: readonly OfficeChartContextCommandId[];
}

/** A change over the chart's current model; null leaves it as it is. */
export type OfficeChartModelCommand = (attrs: ChartBlockAttrsV2) => ChartBlockAttrsV2 | null;

/**
 * The element's own rows, or null when the right click named nothing the
 * menu addresses (the chart area, the plot, an axis): the chart's own menu.
 * A host merges `commands` over its chart menu and drops `omit`
 * (`officeChartMenuWithElement`).
 */
export function officeChartElementMenu(
  attrs: ChartBlockAttrsV2,
  element: OfficeChartContextElement | null | undefined,
  deps: {
    /** Why the chart cannot be written now (null: it can). */
    readonly refused: string | null;
    /** Why its title cannot change (a linked chart's cell-bound title). */
    readonly titleRefused: string | null;
    /** Reset to Match Style with nothing custom on the element. */
    readonly nothingCustom: string;
    /** Writes the command over the chart's current model, as one Undo step. */
    readonly commit: (command: OfficeChartModelCommand) => void;
    readonly openFormat: (area: DocumentChartAreaStyleKey) => void;
    /** Font... on the element's text; a host without it lists no Font row. */
    readonly openFont?: (element: OfficeChartTextElement) => void;
  },
): OfficeChartElementMenu | null {
  if (!element) return null;
  const write = (command: OfficeChartModelCommand | null) => () => {
    if (command) deps.commit(command);
  };
  // Reset to Match Style on one element: its own formatting goes.
  const resetOf = (area: DocumentChartAreaStyleKey): OfficeChartContextCommand => {
    const custom = Boolean(presentationChartAreaStyle(attrs, area));
    return {
      run: write(custom ? (current) => setPresentationChartAreaStyle(current, area, null) : null),
      disabledReason: deps.refused ?? (custom ? null : deps.nothingCustom),
    };
  };
  const fontOf = (
    text: OfficeChartTextElement,
  ): Partial<Record<OfficeChartContextCommandId, OfficeChartContextCommand>> =>
    deps.openFont
      ? { font: { run: () => deps.openFont?.(text), disabledReason: deps.refused } }
      : {};
  // The chart's own Font... is the chart area's text; on an element it is
  // that element's, or none.
  const omit = (font: boolean): readonly OfficeChartContextCommandId[] =>
    font && deps.openFont ? OBJECT_ROWS : [...OBJECT_ROWS, "font"];
  if (element.kind === "series") {
    const seriesIndex = officeChartHitSeries(attrs, element.hit.seriesName);
    if (seriesIndex === null) return null;
    const area = `series:${seriesIndex}` as const;
    const labelled = chartDataLabelStatementOf(attrs, { seriesIndex }).dataLabels;
    return {
      commands: {
        resetToMatchStyle: resetOf(area),
        ...(labelled
          ? {}
          : {
              addDataLabels: {
                run: write((current) =>
                  chartDataLabelStatementOf(current, { seriesIndex }).dataLabels
                    ? null
                    : chartWithDataLabelOverride(
                        current,
                        { seriesIndex },
                        { dataLabels: true, dataLabelValue: true },
                      ),
                ),
                disabledReason: deps.refused,
              },
            }),
        formatSeries: {
          run: () => deps.openFormat(area),
          disabledReason: deps.refused,
        },
      },
      omit: omit(false),
    };
  }
  if (element.part === "legend") {
    const removed = chartWithElementIntent(attrs, { element: "legend", position: "none" });
    return {
      commands: {
        ...(removed
          ? {
              deleteElement: {
                run: write((current) =>
                  chartWithElementIntent(current, { element: "legend", position: "none" }),
                ),
                disabledReason: deps.refused,
              },
            }
          : {}),
        resetToMatchStyle: resetOf("legendStyle"),
        ...fontOf("legend"),
        formatLegend: {
          run: () => deps.openFormat("legendStyle"),
          disabledReason: deps.refused,
        },
      },
      omit: omit(true),
    };
  }
  if (element.part === "title") {
    const removed = chartWithElementIntent(attrs, {
      element: "chartTitle",
      on: false,
      text: attrs.title ?? "",
    });
    return {
      commands: {
        ...(removed
          ? {
              deleteElement: {
                run: write((current) =>
                  chartWithElementIntent(current, {
                    element: "chartTitle",
                    on: false,
                    text: current.title ?? "",
                  }),
                ),
                disabledReason: deps.refused ?? deps.titleRefused,
              },
            }
          : {}),
        resetToMatchStyle: resetOf("chartTitleStyle"),
        ...fontOf("title"),
        formatTitle: {
          run: () => deps.openFormat("chartTitleStyle"),
          disabledReason: deps.refused,
        },
      },
      omit: omit(true),
    };
  }
  return null;
}

/** The chart menu with the element's rows over it (none: the chart's own). */
export function officeChartMenuWithElement(
  commands: Partial<Record<OfficeChartContextCommandId, OfficeChartContextCommand>>,
  element: OfficeChartElementMenu | null,
): Partial<Record<OfficeChartContextCommandId, OfficeChartContextCommand>> {
  if (!element) return commands;
  const merged = { ...commands, ...element.commands };
  for (const id of element.omit) delete merged[id];
  return merged;
}
