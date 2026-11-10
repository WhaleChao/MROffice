"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * 图表设计 and 格式 for a chart made in MRAI on a PDF page (managed-chart.ts):
 * the deck's chart ribbon itself (PresentationChartContextRibbon), on the
 * chart's model -- Change Chart Type, Quick Layout, Change Colors, the style
 * gallery, Chart Elements, Switch Row/Column, Select Data and Edit Data,
 * and the Format tab's pane. A change is the chart with its new model; the
 * page asks the chart engine for its new picture (PdfChartLayer).
 *
 * The chart keeps no workbook: Select Data opens the ribbon's own data
 * dialog and Edit Data the data window (PDFViewer), as a Docs chart without
 * a workbook does.
 */
import { useMemo } from "react";
import { useLocale, useTranslations } from "@mrnote/core/host/next-intl.tsx";

import type { ChartBlockAttrsV2 } from "@mrnote/core/web/components/console/editor/extensions/chartBlockShared.ts";
import { normalizeChartBlockAttrs } from "@mrnote/core/web/components/console/editor/extensions/chartBlockShared.ts";
import { OFFICE_WEB_SHAPE_ACCENTS } from "@mrnote/core/web/lib/office-shape-style-catalog.ts";
import type { PdfChart } from "../../../../lib/pdf/managed-chart";
import type { OfficeTaskPaneDockRegistry } from "@mrnote/core/web/components/office-ribbon/officeTaskPaneDock.ts";

import { PresentationChartContextRibbon } from "@mrnote/slides/web/components/notebook/contents/presentation/PresentationChartContextRibbon.tsx";

/** EMU per PDF point (12 700, as the deck's geometry). */
const EMU_PER_POINT = 12_700;

export interface PdfChartContextRibbonProps {
  readonly tab: "chart-design" | "chart-format";
  readonly chart: PdfChart;
  /** The chart's page in points (its frame is a share of it). */
  readonly pageWidth: number;
  readonly pageHeight: number;
  readonly writable: boolean;
  readonly onChange: (chart: PdfChart) => void;
  /** A gallery's hovered pick drawn on the chart until the menu closes. */
  readonly onPreview: (attrs: ChartBlockAttrsV2 | null) => void;
  readonly onEditData: () => void;
  /** Format > More Options: the Format tab. */
  readonly onOpenFormat?: () => void;
  /** Where the Format pane docks (the reader's task pane dock). */
  readonly paneDock?: OfficeTaskPaneDockRegistry | null;
}

/** The chart's model as the chart engine reads it (a stored model checked whole). */
export function pdfChartAttrs(chart: PdfChart): ChartBlockAttrsV2 | null {
  const normalized = normalizeChartBlockAttrs(chart.attrs);
  return normalized.source === "v2" && normalized.validation.valid ? normalized.attrs : null;
}

export function PdfChartContextRibbon({
  tab, chart, pageWidth, pageHeight, writable, onChange, onPreview, onEditData, onOpenFormat, paneDock,
}: PdfChartContextRibbonProps) {
  const t = useTranslations("console-notebooks");
  const locale = useLocale();
  const attrs = useMemo(() => pdfChartAttrs(chart), [chart]);
  // What the Format tab remembers its Current Selection by: one object for
  // as long as the same chart is selected.
  const identity = useMemo(() => ({ chart: chart.id }), [chart.id]);
  if (!attrs) return null;
  const size = {
    widthEmu: Math.round(chart.width * pageWidth * EMU_PER_POINT),
    heightEmu: Math.round(chart.height * pageHeight * EMU_PER_POINT),
  };
  return (
    <div data-pdf-ribbon-tab={tab} data-testid={`pdf-${tab}-ribbon`} style={{ display: "contents" }}>
      <PresentationChartContextRibbon
        tab={tab}
        attrs={attrs}
        identity={identity}
        authority={chart}
        disabledReason={writable ? null : t("pdf.editNotReady")}
        locale={locale}
        accentColors={OFFICE_WEB_SHAPE_ACCENTS}
        size={size}
        onSize={(axis, emu) => {
          if (!writable || !Number.isFinite(emu) || emu <= 0) return false;
          const points = emu / EMU_PER_POINT;
          const next = axis === "widthEmu"
            ? { width: Math.min(1 - chart.x, points / pageWidth) }
            : { height: Math.min(1 - chart.y, points / pageHeight) };
          if ((next.width ?? 1) < 0.02 || (next.height ?? 1) < 0.02) return false;
          onChange({ ...chart, ...next, svg: null });
          return true;
        }}
        onCommit={(next) => {
          if (!writable) return false;
          onPreview(null);
          onChange({ ...chart, attrs: next as unknown as Record<string, unknown>, svg: null });
          return true;
        }}
        onPreview={onPreview}
        onEditData={onEditData}
        onOpenFormat={onOpenFormat ? () => onOpenFormat() : undefined}
        arrange={null}
        insertShapes={null}
        paneContainer={null}
        paneDock={paneDock}
      />
    </div>
  );
}
