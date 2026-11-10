"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * 表格布局: the contextual tab of a table selected on a PDF page
 * (PdfTableLayer), the presentation editor's Table Layout and Design
 * commands on the table's model (managed-table.ts): rows and columns,
 * merge and split, the cells' alignment, weight and slant, shading,
 * borders, and the header and banded rows.
 */
import { useRef, useState } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import {
  OfficeHomeCommand,
  OfficeHomeCommandRow,
  OfficeHomeMenuTrigger,
  OfficeHomeRibbonGroup,
} from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import type { OfficeSemanticIconId } from "@mrnote/core/web/components/office-ribbon/icons/officeIconRegistry.ts";
import { OfficeColorPalettePopover } from "@mrnote/core/web/components/office-ribbon/OfficeColorPalettePopover.tsx";
import { RibbonDropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import { DropdownMenu, DropdownMenuItem, DropdownMenuTrigger } from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import {
  deletePdfTableColumn,
  deletePdfTableRow,
  insertPdfTableColumn,
  insertPdfTableRow,
  mergePdfTableCells,
  pdfTableHasBandedRows,
  pdfTableHasHeaderRow,
  pdfTableRange,
  setPdfTableBandedRows,
  setPdfTableHeaderRow,
  splitPdfTableCell,
  updatePdfTableCells,
  type PdfTable,
  type PdfTableAlign,
  type PdfTableVerticalAlign,
} from "../../../../lib/pdf/managed-table";

import type { PdfTableCellAddress } from "./PdfTableLayer";
import { PdfTableSettingsPane } from "./PdfTableSettingsPane";
import { OfficeTaskPaneDockSlot } from "@mrnote/core/web/components/office-ribbon/OfficeTaskPaneDockSlot.tsx";
import type { OfficeTaskPaneDockRegistry } from "@mrnote/core/web/components/office-ribbon/officeTaskPaneDock.ts";

export interface PdfTableContextRibbonProps {
  readonly table: PdfTable;
  /** The cell (or range) selected; null with the table as a whole. */
  readonly selection: PdfTableCellAddress | null;
  readonly writable: boolean;
  readonly onChange: (table: PdfTable) => void;
  readonly onDelete: (table: PdfTable) => void;
  readonly onSelectionChange: (selection: PdfTableCellAddress | { readonly tableId: string } | null) => void;
  /** The reader's right-edge dock, where Table Properties opens. */
  readonly paneDock?: OfficeTaskPaneDockRegistry | null;
}

const DEFAULT_BORDER = "#2F5597";

export function PdfTableContextRibbon({ table, selection, writable, onChange, onDelete, onSelectionChange, paneDock = null }: PdfTableContextRibbonProps) {
  const t = useTranslations("console-notebooks");
  const shadingRef = useRef<HTMLButtonElement | null>(null);
  const [shadingOpen, setShadingOpen] = useState(false);
  const [paneOpen, setPaneOpen] = useState(false);
  // The whole table when no cell is selected (PowerPoint applies a cell
  // format to the table selected as a whole).
  const range = selection
    ? pdfTableRange(table, selection, { row: selection.endRow ?? selection.row, column: selection.endColumn ?? selection.column })
    : { row: 0, column: 0, endRow: table.rows.length - 1, endColumn: table.columns.length - 1 };
  const owner = table.rows[range.row]?.cells[range.column] ?? null;
  const multiple = range.endRow > range.row || range.endColumn > range.column;
  const mergedHere = owner !== null && ((owner.colSpan ?? 1) > 1 || (owner.rowSpan ?? 1) > 1);
  const cellsIn = table.rows.slice(range.row, range.endRow + 1).flatMap((row) =>
    row.cells.slice(range.column, range.endColumn + 1).filter((cell) => cell !== null));
  const allBold = cellsIn.length > 0 && cellsIn.every((cell) => cell?.bold);
  const allItalic = cellsIn.length > 0 && cellsIn.every((cell) => cell?.italic);
  const disabled = !writable;
  const apply = (next: PdfTable | null) => {
    if (!next) {
      onDelete(table);
      onSelectionChange(null);
      return;
    }
    if (next !== table) onChange(next);
  };
  const keep = (row: number, column: number) => onSelectionChange({ tableId: table.id, row, column });

  const alignCommand = (align: PdfTableAlign, label: string, icon: OfficeSemanticIconId) => (
    <OfficeHomeCommand
      key={align}
      active={owner?.align === align || (align === "left" && !owner?.align)}
      controlId={`PDFCTL.CONTEXT.TABLE.ALIGN.${align.toUpperCase()}`}
      disabled={disabled}
      icon={<OfficeSemanticIcon semanticId={icon} size={16} />}
      label={label}
      testId={`pdf-table-align-${align}`}
      variant="icon"
      onClick={() => apply(updatePdfTableCells(table, range, { align: align === "left" ? undefined : align }))}
    />
  );
  const verticalCommand = (vertical: PdfTableVerticalAlign, label: string, icon: OfficeSemanticIconId) => (
    <OfficeHomeCommand
      key={vertical}
      active={owner?.verticalAlign === vertical || (vertical === "top" && !owner?.verticalAlign)}
      controlId={`PDFCTL.CONTEXT.TABLE.VALIGN.${vertical.toUpperCase()}`}
      disabled={disabled}
      icon={<OfficeSemanticIcon semanticId={icon} size={16} />}
      label={label}
      testId={`pdf-table-valign-${vertical}`}
      variant="icon"
      onClick={() => apply(updatePdfTableCells(table, range, { verticalAlign: vertical === "top" ? undefined : vertical }))}
    />
  );

  return (
    <>
      <OfficeHomeRibbonGroup label={t("presentation.tableRibbon.rowsAndColumns")} ribbon="table-format"
        controlId="PDFCTL.CONTEXT.TABLE.ROWS_COLUMNS" testId="pdf-table-rows-group">
        <OfficeHomeCommandRow>
          <DropdownMenu>
            <DropdownMenuTrigger asChild>
              <OfficeHomeMenuTrigger
                data-control-id="PDFCTL.CONTEXT.TABLE.DELETE"
                disabled={disabled}
                hasMenuIndicator
                icon={<OfficeSemanticIcon semanticId="table.delete" size={24} />}
                label={t("presentation.contextMenu.delete")}
                testId="pdf-table-delete-menu"
                variant="large"
              />
            </DropdownMenuTrigger>
            <RibbonDropdownMenuContent align="start" data-testid="pdf-table-delete-menu-content">
              <DropdownMenuItem data-testid="pdf-table-delete-row" onSelect={() => {
                let next: PdfTable | null = table;
                for (let row = range.endRow; row >= range.row && next; row -= 1) next = deletePdfTableRow(next, row);
                apply(next);
                if (next) keep(Math.min(range.row, next.rows.length - 1), range.column);
              }}>
                <OfficeSemanticIcon semanticId="table.delete" size={16} />{t("toolbar.deleteRow")}
              </DropdownMenuItem>
              <DropdownMenuItem data-testid="pdf-table-delete-column" onSelect={() => {
                let next: PdfTable | null = table;
                for (let column = range.endColumn; column >= range.column && next; column -= 1) next = deletePdfTableColumn(next, column);
                apply(next);
                if (next) keep(range.row, Math.min(range.column, next.columns.length - 1));
              }}>
                <OfficeSemanticIcon semanticId="table.delete" size={16} />{t("toolbar.deleteColumn")}
              </DropdownMenuItem>
              <DropdownMenuItem data-testid="pdf-table-delete-table" onSelect={() => apply(null)}>
                <OfficeSemanticIcon semanticId="table.delete" size={16} />{t("toolbar.deleteTable")}
              </DropdownMenuItem>
            </RibbonDropdownMenuContent>
          </DropdownMenu>
          <OfficeHomeCommand variant="large" controlId="PDFCTL.CONTEXT.TABLE.INSERT_ABOVE" disabled={disabled}
            icon={<OfficeSemanticIcon semanticId="table.insertAbove" size={20} />}
            label={t("presentation.tableRibbon.insertAboveShort")} testId="pdf-table-insert-above"
            onClick={() => { apply(insertPdfTableRow(table, range.row, false)); keep(range.row, range.column); }} />
          <OfficeHomeCommand variant="large" controlId="PDFCTL.CONTEXT.TABLE.INSERT_BELOW" disabled={disabled}
            icon={<OfficeSemanticIcon semanticId="table.insertBelow" size={20} />}
            label={t("presentation.tableRibbon.insertBelowShort")} testId="pdf-table-insert-below"
            onClick={() => { apply(insertPdfTableRow(table, range.endRow, true)); keep(range.endRow + 1, range.column); }} />
          <OfficeHomeCommand variant="large" controlId="PDFCTL.CONTEXT.TABLE.INSERT_LEFT" disabled={disabled}
            icon={<OfficeSemanticIcon semanticId="table.insertLeft" size={20} />}
            label={t("presentation.tableRibbon.insertLeftShort")} testId="pdf-table-insert-left"
            onClick={() => { apply(insertPdfTableColumn(table, range.column, false)); keep(range.row, range.column); }} />
          <OfficeHomeCommand variant="large" controlId="PDFCTL.CONTEXT.TABLE.INSERT_RIGHT" disabled={disabled}
            icon={<OfficeSemanticIcon semanticId="table.insertRight" size={20} />}
            label={t("presentation.tableRibbon.insertRightShort")} testId="pdf-table-insert-right"
            onClick={() => { apply(insertPdfTableColumn(table, range.endColumn, true)); keep(range.row, range.endColumn + 1); }} />
        </OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>
      <OfficeHomeRibbonGroup label={t("presentation.tableRibbon.mergeCells")} ribbon="table-format"
        controlId="PDFCTL.CONTEXT.TABLE.MERGE" testId="pdf-table-merge-group">
        <OfficeHomeCommandRow>
          <OfficeHomeCommand variant="large" controlId="PDFCTL.CONTEXT.TABLE.MERGE_CELLS" disabled={disabled || !multiple}
            icon={<OfficeSemanticIcon semanticId="table.merge" size={20} />}
            label={t("presentation.tableMerge")} testId="pdf-table-merge"
            onClick={() => { apply(mergePdfTableCells(table, range)); keep(range.row, range.column); }} />
          <OfficeHomeCommand variant="large" controlId="PDFCTL.CONTEXT.TABLE.SPLIT_CELLS" disabled={disabled || !mergedHere || multiple}
            icon={<OfficeSemanticIcon semanticId="table.split" size={20} />}
            label={t("presentation.tableSplit")} testId="pdf-table-split"
            onClick={() => { apply(splitPdfTableCell(table, range.row, range.column)); keep(range.row, range.column); }} />
        </OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>
      <OfficeHomeRibbonGroup label={t("presentation.tableRibbon.alignment")} ribbon="table-format"
        controlId="PDFCTL.CONTEXT.TABLE.ALIGNMENT" testId="pdf-table-alignment-group">
        <OfficeHomeCommandRow>
          {alignCommand("left", t("toolbar.alignLeft"), "paragraph.align.left")}
          {alignCommand("center", t("toolbar.alignCenter"), "paragraph.align.center")}
          {alignCommand("right", t("toolbar.alignRight"), "paragraph.align.right")}
        </OfficeHomeCommandRow>
        <OfficeHomeCommandRow>
          {verticalCommand("top", t("presentation.valign.top"), "table.cell.alignTop")}
          {verticalCommand("middle", t("presentation.valign.middle"), "table.cell.alignMiddle")}
          {verticalCommand("bottom", t("presentation.valign.bottom"), "table.cell.alignBottom")}
        </OfficeHomeCommandRow>
        <OfficeHomeCommandRow>
          <OfficeHomeCommand active={allBold} controlId="PDFCTL.CONTEXT.TABLE.BOLD" disabled={disabled}
            icon={<OfficeSemanticIcon semanticId="text.bold" size={16} />} label={t("toolbar.bold")}
            testId="pdf-table-bold" variant="icon"
            onClick={() => apply(updatePdfTableCells(table, range, { bold: allBold ? undefined : true }))} />
          <OfficeHomeCommand active={allItalic} controlId="PDFCTL.CONTEXT.TABLE.ITALIC" disabled={disabled}
            icon={<OfficeSemanticIcon semanticId="text.italic" size={16} />} label={t("toolbar.italic")}
            testId="pdf-table-italic" variant="icon"
            onClick={() => apply(updatePdfTableCells(table, range, { italic: allItalic ? undefined : true }))} />
        </OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>
      <OfficeHomeRibbonGroup label={t("presentation.tableRibbon.tableStyles")} ribbon="table-format"
        controlId="PDFCTL.CONTEXT.TABLE.STYLES" testId="pdf-table-style-group"
        launcher={paneDock ? {
          label: t("toolbar.tableProperties"),
          onClick: () => setPaneOpen((open) => !open),
          active: paneOpen,
          testId: "pdf-table-properties-launcher",
        } : undefined}>
        <OfficeHomeCommandRow>
          <OfficeHomeCommand buttonRef={shadingRef} activation="click" ariaExpanded={shadingOpen}
            controlId="PDFCTL.CONTEXT.TABLE.SHADING" disabled={disabled}
            icon={<OfficeSemanticIcon semanticId="table.shading" size={20} />}
            label={t("presentation.tableRibbon.shading")} testId="pdf-table-shading" variant="large"
            onClick={() => setShadingOpen((open) => !open)} />
          <OfficeHomeCommand active={table.borderColor !== null} controlId="PDFCTL.CONTEXT.TABLE.BORDERS_ALL" disabled={disabled}
            icon={<OfficeSemanticIcon semanticId="table.borders" size={20} />}
            label={table.borderColor !== null ? t("toolbar.noBorder") : t("toolbar.tableBordersAll")}
            testId="pdf-table-borders" variant="large"
            onClick={() => apply({ ...table, borderColor: table.borderColor !== null ? null : DEFAULT_BORDER, borderWidth: table.borderWidth || 0.75 })} />
        </OfficeHomeCommandRow>
        <OfficeHomeCommandRow>
          <OfficeHomeCommand active={pdfTableHasHeaderRow(table)} controlId="PDFCTL.CONTEXT.TABLE.HEADER_ROW" disabled={disabled}
            icon={<OfficeSemanticIcon semanticId="table.style.headerRow" size={16} />} label={t("presentation.tableHeaderRow")}
            displayLabel={t("presentation.tableHeaderRow")} testId="pdf-table-header-row" variant="compact"
            onClick={() => apply(setPdfTableHeaderRow(table, !pdfTableHasHeaderRow(table)))} />
          <OfficeHomeCommand active={pdfTableHasBandedRows(table)} controlId="PDFCTL.CONTEXT.TABLE.BANDED_ROWS" disabled={disabled}
            icon={<OfficeSemanticIcon semanticId="table.style.bandedRows" size={16} />} label={t("presentation.tableBandedRows")}
            displayLabel={t("presentation.tableBandedRows")} testId="pdf-table-banded-rows" variant="compact"
            onClick={() => apply(setPdfTableBandedRows(table, !pdfTableHasBandedRows(table)))} />
        </OfficeHomeCommandRow>
      </OfficeHomeRibbonGroup>
      {paneOpen && paneDock ? (
        <OfficeTaskPaneDockSlot dock={paneDock} id="table-properties">
          <PdfTableSettingsPane table={table} range={range} writable={writable}
            onChange={(next) => apply(next)} onClose={() => setPaneOpen(false)} />
        </OfficeTaskPaneDockSlot>
      ) : null}
      {shadingOpen ? (
        <OfficeColorPalettePopover
          anchorRef={shadingRef}
          value={owner?.fill}
          labels={{
            themeColors: t("pdf.formField.themeColors"),
            standardColors: t("pdf.formField.standardColors"),
            recentColors: t("pdf.formField.recentColors"),
            moreColors: t("pdf.formField.moreColors"),
            noFill: t("toolbar.noFill"),
          }}
          allowNoFill
          noFillSelected={!owner?.fill}
          onPick={(color) => {
            apply(updatePdfTableCells(table, range, { fill: color ? color.toUpperCase() : undefined }));
            setShadingOpen(false);
          }}
          onDismiss={() => setShadingOpen(false)}
          testIdPrefix="pdf-table-shading"
        />
      ) : null}
    </>
  );
}
