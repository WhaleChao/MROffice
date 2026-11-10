"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Table settings for a table on a PDF page (ONLYOFFICE's Table Settings at
 * the right; PowerPoint's Format Table): the shared OfficeShapeFormatPane on
 * the table's model (managed-table.ts), every field one the Table Layout tab
 * (PdfTableContextRibbon) already writes -- the style options, the grid's
 * colour and weight, and the selected cells' alignment, shading, text colour,
 * bold and italic (the whole table when no cell is selected).
 */
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeShapeFormatPane } from "@mrnote/core/web/components/office-ribbon/OfficeShapeFormatPane.tsx";
import {
  pdfTableHasBandedRows,
  pdfTableHasHeaderRow,
  setPdfTableBandedRows,
  setPdfTableHeaderRow,
  updatePdfTableCells,
  type PdfTable,
  type PdfTableAlign,
  type PdfTableCell,
  type PdfTableRange,
  type PdfTableVerticalAlign,
} from "../../../../lib/pdf/managed-table";

import { pdfTaskPaneClassName } from "./pdfTaskPaneDock";

const DEFAULT_BORDER = "#2F5597";
const ALIGNS: readonly PdfTableAlign[] = ["left", "center", "right"];
const VERTICAL_ALIGNS: readonly PdfTableVerticalAlign[] = ["top", "middle", "bottom"];

export function PdfTableSettingsPane({
  table,
  range,
  writable,
  onChange,
  onClose,
}: {
  readonly table: PdfTable;
  /** The cells the cell settings apply to (the whole table when none is selected). */
  readonly range: PdfTableRange;
  readonly writable: boolean;
  readonly onChange: (table: PdfTable) => void;
  readonly onClose: () => void;
}) {
  const t = useTranslations("console-notebooks");
  const owner = table.rows[range.row]?.cells[range.column] ?? null;
  const cellsIn = table.rows.slice(range.row, range.endRow + 1).flatMap((row) =>
    row.cells.slice(range.column, range.endColumn + 1).filter((cell): cell is PdfTableCell => cell !== null));
  const every = (test: (cell: PdfTableCell) => boolean) => cellsIn.length > 0 && cellsIn.every(test);
  const cells = (change: Partial<PdfTableCell>) => {
    const next = updatePdfTableCells(table, range, change);
    if (next !== table) onChange(next);
  };
  const alignLabels: Readonly<Record<PdfTableAlign, string>> = {
    left: t("toolbar.alignLeft"), center: t("toolbar.alignCenter"), right: t("toolbar.alignRight"),
  };
  const verticalLabels: Readonly<Record<PdfTableVerticalAlign, string>> = {
    top: t("presentation.valign.top"), middle: t("presentation.valign.middle"), bottom: t("presentation.valign.bottom"),
  };

  return (
    <OfficeShapeFormatPane
      testIdPrefix="pdf-table-settings-pane"
      labels={{
        title: t("toolbar.tableProperties"),
        close: t("common.close"),
        apply: t("document.shapeFormatPane.apply"),
        cancel: t("document.shapeFormatPane.cancel"),
      }}
      tabs={[
        {
          id: "table",
          label: t("presentation.tableInsert.tableNoun"),
          sections: [
            {
              id: "style-options",
              title: t("presentation.tableRibbon.styleOptions"),
              fields: [
                {
                  kind: "checkbox", id: "header-row", label: t("presentation.tableHeaderRow"),
                  checked: pdfTableHasHeaderRow(table),
                  onChange: (checked) => onChange(setPdfTableHeaderRow(table, checked)),
                },
                {
                  kind: "checkbox", id: "banded-rows", label: t("presentation.tableBandedRows"),
                  checked: pdfTableHasBandedRows(table),
                  onChange: (checked) => onChange(setPdfTableBandedRows(table, checked)),
                },
              ],
            },
            {
              id: "borders",
              title: t("presentation.tableRibbon.borders"),
              fields: [
                {
                  kind: "color", id: "border-color", label: t("presentation.tableRibbon.penColor"),
                  value: table.borderColor ?? "none", fallback: DEFAULT_BORDER,
                  none: { label: t("toolbar.noBorder"), restore: DEFAULT_BORDER },
                  onChange: (value) => onChange({
                    ...table,
                    borderColor: value === "none" ? null : value.toUpperCase(),
                    borderWidth: table.borderWidth || 0.75,
                  }),
                },
                {
                  kind: "number", id: "border-width", label: t("presentation.tableRibbon.penWeight"),
                  value: table.borderWidth, min: 0.25, max: 6, step: 0.25, unit: "pt",
                  disabled: table.borderColor === null,
                  onCommit: (raw) => {
                    const width = Number.parseFloat(raw);
                    if (!Number.isFinite(width) || width < 0.25 || width > 6) return false;
                    onChange({ ...table, borderWidth: width });
                    return true;
                  },
                },
              ],
            },
          ],
        },
        {
          id: "cells",
          label: t("spreadsheet.watchCell"),
          sections: [
            {
              id: "alignment",
              title: t("presentation.tableRibbon.alignment"),
              fields: [
                {
                  kind: "select", id: "align", label: t("spreadsheet.fmtAlignHorizontal"),
                  value: owner?.align ?? "left",
                  options: ALIGNS.map((align) => ({ value: align, label: alignLabels[align] })),
                  onChange: (value) => cells({ align: value === "left" ? undefined : value as PdfTableAlign }),
                },
                {
                  kind: "select", id: "vertical-align", label: t("presentation.tableCellVAlign"),
                  value: owner?.verticalAlign ?? "top",
                  options: VERTICAL_ALIGNS.map((vertical) => ({ value: vertical, label: verticalLabels[vertical] })),
                  onChange: (value) => cells({ verticalAlign: value === "top" ? undefined : value as PdfTableVerticalAlign }),
                },
              ],
            },
            {
              id: "look",
              title: t("presentation.tableRibbon.shading"),
              fields: [
                {
                  kind: "color", id: "shading", label: t("presentation.tableRibbon.shading"),
                  value: owner?.fill ?? "none", fallback: "#FFFFFF",
                  none: { label: t("toolbar.noFill"), restore: "#DEEAF6" },
                  onChange: (value) => cells({ fill: value === "none" ? undefined : value.toUpperCase() }),
                },
                {
                  kind: "color", id: "text-color", label: t("toolbar.textColor"),
                  value: owner?.color ?? "#000000",
                  onChange: (value) => { if (value !== "none") cells({ color: value.toUpperCase() }); },
                },
                {
                  kind: "checkbox", id: "bold", label: t("toolbar.bold"), checked: every((cell) => Boolean(cell.bold)),
                  onChange: (checked) => cells({ bold: checked ? true : undefined }),
                },
                {
                  kind: "checkbox", id: "italic", label: t("toolbar.italic"), checked: every((cell) => Boolean(cell.italic)),
                  onChange: (checked) => cells({ italic: checked ? true : undefined }),
                },
              ],
            },
          ],
        },
      ]}
      writable={writable}
      onClose={onClose}
      dockRight={false}
      className={pdfTaskPaneClassName("pdf-viewer__shape-format-pane")}
    />
  );
}
