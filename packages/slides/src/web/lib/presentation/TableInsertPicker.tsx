"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useRef, useState, type KeyboardEvent } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";

import { OfficeSemanticIcon } from "@mrnote/core/web/components/office-ribbon/icons/OfficeSemanticIcon.tsx";
import { OfficeHomeMenuTrigger } from "@mrnote/core/web/components/office-ribbon/OfficeHomeRibbon.tsx";
import { RibbonDropdownMenuContent as DropdownMenuContent } from "@mrnote/core/web/components/office-ribbon/CupertinoRibbonMenu.tsx";
import {
  DropdownMenu,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "@mrnote/core/web/components/ui/dropdown-menu.tsx";
import {
  INSERT_TABLE_DIALOG_DEFAULT_COLUMNS,
  INSERT_TABLE_DIALOG_DEFAULT_ROWS,
  QUICK_INSERT_TABLE_COLUMNS,
  QUICK_INSERT_TABLE_ROWS,
  type TableInsertGridPosition,
  tableInsertGridNavigationTarget,
  validateTableInsertDimensions,
} from "../presentation-table-insert";

import { PresentationTableInsertDialog } from "./TableInsertDialog";

interface PresentationTableInsertPickerProps {
  label: string;
  disabled?: boolean;
  onInsert: (rows: number, columns: number) => boolean;
  onDrawTable?: () => void;
  /** Another editor's own names for the launcher (the PDF's Insert tab). */
  testId?: string;
  controlId?: string;
}

const FIRST_GRID_CELL: TableInsertGridPosition = { rows: 1, columns: 1 };

/**
 * PowerPoint-style Table launcher: a keyboard-navigable 10×8 quick gallery and
 * a proper Insert Table dialog for larger dimensions. Domain mutation remains
 * entirely caller-owned through the one live `onInsert` authority.
 */
export default function PresentationTableInsertPicker({
  label,
  disabled = false,
  onInsert,
  onDrawTable,
  testId = "presentation-ribbon-add-table",
  controlId = "PPTCTL.RIBBON.INSERT.TABLE.MENU",
}: PresentationTableInsertPickerProps) {
  const t = useTranslations("console-notebooks");
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  const gridCellRefs = useRef(new Map<string, HTMLButtonElement>());
  const suppressMenuReturnFocusRef = useRef(false);
  const [open, setOpen] = useState(false);
  const [dialogOpen, setDialogOpen] = useState(false);
  const [dialogSession, setDialogSession] = useState(0);
  const [focused, setFocused] =
    useState<TableInsertGridPosition>(FIRST_GRID_CELL);
  const [preview, setPreview] =
    useState<TableInsertGridPosition>(FIRST_GRID_CELL);
  const [commandError, setCommandError] = useState<string | null>(null);

  const focusGridCell = (position: TableInsertGridPosition) => {
    setFocused(position);
    setPreview(position);
    gridCellRefs.current.get(`${position.rows}x${position.columns}`)?.focus();
  };

  const insert = (requestedRows: unknown, requestedColumns: unknown) => {
    const validation = validateTableInsertDimensions(
      requestedRows,
      requestedColumns,
    );
    if (
      !validation.ok ||
      validation.rows === null ||
      validation.columns === null
    ) {
      setCommandError(t("presentation.tableInsert.invalidDimensions"));
      return false;
    }
    if (!onInsert(validation.rows, validation.columns)) {
      setCommandError(t("presentation.tableInsert.commandError"));
      return false;
    }
    setCommandError(null);
    suppressMenuReturnFocusRef.current = true;
    setOpen(false);
    return true;
  };

  const handleGridKeyDown = (
    event: KeyboardEvent<HTMLButtonElement>,
    position: TableInsertGridPosition,
  ) => {
    if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      suppressMenuReturnFocusRef.current = false;
      setOpen(false);
      return;
    }
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      event.stopPropagation();
      insert(position.rows, position.columns);
      return;
    }
    const target = tableInsertGridNavigationTarget(
      position,
      event.key,
      event.metaKey || event.ctrlKey,
    );
    if (!target) return;
    event.preventDefault();
    event.stopPropagation();
    focusGridCell(target);
  };

  return (
    <>
      <DropdownMenu
        open={open}
        onOpenChange={(nextOpen) => {
          if (disabled && nextOpen) return;
          setOpen(nextOpen);
          setCommandError(null);
          if (nextOpen) {
            suppressMenuReturnFocusRef.current = false;
            setFocused(FIRST_GRID_CELL);
            setPreview(FIRST_GRID_CELL);
            window.requestAnimationFrame(() =>
              gridCellRefs.current.get("1x1")?.focus(),
            );
          }
        }}
      >
        <DropdownMenuTrigger asChild>
          <OfficeHomeMenuTrigger
            ref={triggerRef}
            variant="large"
            hasMenuIndicator
            icon={<OfficeSemanticIcon semanticId="table.insert" size={24} />}
            label={label}
            testId={testId}
            data-control-id={controlId}
            aria-haspopup="menu"
            aria-expanded={open}
            disabled={disabled}
          />
        </DropdownMenuTrigger>
        <DropdownMenuContent
          align="start"
          sideOffset={4}
          data-testid="presentation-table-insert-menu"
          onCloseAutoFocus={(event) => {
            if (!suppressMenuReturnFocusRef.current) return;
            event.preventDefault();
            suppressMenuReturnFocusRef.current = false;
          }}
        >
          {/* PowerPoint names a table columns first: "3x2 Table" over the
              grid, "Table with 3 columns, 2 rows" on a cell (PPT-A060). */}
          <DropdownMenuLabel
            data-testid="presentation-table-insert-preview"
            data-columns={preview.columns}
            data-rows={preview.rows}
            aria-live="polite"
            aria-atomic="true"
          >
            {t("presentation.tableInsert.gridHeader", {
              columns: preview.columns,
              rows: preview.rows,
            })}
          </DropdownMenuLabel>
          <div
            className="note-editor-table-grid"
            role="grid"
            aria-label={t("presentation.tableInsert.gridAria", {
              rows: preview.rows,
              columns: preview.columns,
            })}
            aria-rowcount={QUICK_INSERT_TABLE_ROWS}
            aria-colcount={QUICK_INSERT_TABLE_COLUMNS}
            data-testid="presentation-table-insert-grid"
            onMouseLeave={() => setPreview(focused)}
          >
            {Array.from({ length: QUICK_INSERT_TABLE_ROWS }, (_, rowIndex) => {
              const rows = rowIndex + 1;
              return (
                <div
                  key={rows}
                  className="note-editor-table-grid__row"
                  role="row"
                  aria-rowindex={rows}
                >
                  {Array.from(
                    { length: QUICK_INSERT_TABLE_COLUMNS },
                    (_, columnIndex) => {
                      const columns = columnIndex + 1;
                      const position = { rows, columns };
                      const highlighted =
                        rows <= preview.rows && columns <= preview.columns;
                      return (
                        <button
                          key={columns}
                          ref={(cell) => {
                            const key = `${rows}x${columns}`;
                            if (cell) gridCellRefs.current.set(key, cell);
                            else gridCellRefs.current.delete(key);
                          }}
                          type="button"
                          role="gridcell"
                          tabIndex={
                            focused.rows === rows && focused.columns === columns
                              ? 0
                              : -1
                          }
                          aria-rowindex={rows}
                          aria-colindex={columns}
                          aria-selected={
                            preview.rows === rows && preview.columns === columns
                          }
                          aria-label={t("presentation.tableInsert.gridAria", {
                            rows,
                            columns,
                          })}
                          data-testid={`presentation-table-insert-${rows}x${columns}`}
                          data-control-id={`PPTCTL.RIBBON.INSERT.TABLE.PICKER.R${String(rows).padStart(2, "0")}C${String(columns).padStart(2, "0")}`}
                          className={
                            highlighted
                              ? "note-editor-table-grid__cell note-editor-table-grid__cell--active"
                              : "note-editor-table-grid__cell"
                          }
                          onMouseEnter={() => setPreview(position)}
                          onFocus={() => {
                            setFocused(position);
                            setPreview(position);
                          }}
                          onKeyDown={(event) =>
                            handleGridKeyDown(event, position)
                          }
                          onClick={() => insert(rows, columns)}
                        />
                      );
                    },
                  )}
                </div>
              );
            })}
          </div>
          {commandError ? (
            <p
              className="mx-2 rounded-md bg-destructive/10 px-2 py-1.5 text-xs text-destructive"
              role="alert"
              data-testid="presentation-table-insert-menu-error"
            >
              {commandError}
            </p>
          ) : null}
          <DropdownMenuSeparator />
          <DropdownMenuItem
            data-testid="presentation-table-insert-dialog-entry"
            data-control-id="PPTCTL.RIBBON.INSERT.TABLE.DIALOG_ENTRY"
            onSelect={() => {
              suppressMenuReturnFocusRef.current = true;
              setOpen(false);
              setDialogSession((session) => session + 1);
              setDialogOpen(true);
            }}
          >
            <OfficeSemanticIcon semanticId="table.insert" size={16} />
            {t("toolbar.insertTableSize")}
          </DropdownMenuItem>
          {onDrawTable ? (
            <DropdownMenuItem
              disabled={disabled}
              data-testid="presentation-table-draw-entry"
              onSelect={() => {
                if (disabled) return;
                suppressMenuReturnFocusRef.current = true;
                setOpen(false);
                onDrawTable();
              }}
            >
              <OfficeSemanticIcon semanticId="draw.pen" size={16} />
              {t("presentation.tableRibbon.drawTable")}
            </DropdownMenuItem>
          ) : null}
        </DropdownMenuContent>
      </DropdownMenu>
      <PresentationTableInsertDialog
        key={dialogSession}
        open={dialogOpen}
        initialRows={INSERT_TABLE_DIALOG_DEFAULT_ROWS}
        initialColumns={INSERT_TABLE_DIALOG_DEFAULT_COLUMNS}
        returnFocusRef={triggerRef}
        onInsert={insert}
        onClose={() => setDialogOpen(false)}
      />
    </>
  );
}
