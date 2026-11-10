"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Error Checking panel (公式 ▸ 错误检查).
 *
 * Excel and WPS walk you through suspect formulas one at a time rather than
 * dumping a list: Previous/Next move the selection to the cell, and the panel
 * says what it found there. The scan itself lives in `error-checking.ts`; this
 * is only the walk and the wording.
 *
 * Display copy comes from the caller's active message catalogue.
 */

import { useEffect, useMemo, useState } from "react";

import {
  officeDialogClassName,
  officeDialogParts as dialogParts,
} from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";

import type { ErrorCheckFinding, ErrorCheckRule } from "./error-checking";

export interface ErrorCheckPanelProps {
  open: boolean;
  translate: (key: string, values?: Record<string, string | number>) => string;
  findings: readonly ErrorCheckFinding[];
  /** Column label for a 0-based index ("A", "AB"), so the panel names cells. */
  columnLabel: (col: number) => string;
  /** Move the grid selection to a finding. */
  onSelect: (row: number, col: number) => void;
  /** Open the existing formula walkthrough for this exact finding. */
  onShowCalculationSteps: (finding: ErrorCheckFinding) => void;
  /** Transfer focus to the existing formula bar for this exact finding. */
  onEditInFormulaBar: (finding: ErrorCheckFinding) => void;
  onClose: () => void;
}

export function ErrorCheckPanel({
  open,
  translate,
  findings,
  columnLabel,
  onSelect,
  onShowCalculationSteps,
  onEditInFormulaBar,
  onClose,
}: ErrorCheckPanelProps) {
  const text = {
    title: translate("spreadsheet.helperCopy.errorCheck.title"),
    none: translate("spreadsheet.helperCopy.errorCheck.none"),
    position: (index: number, total: number) => translate("spreadsheet.helperCopy.errorCheck.position", { index, total }),
    cell: translate("spreadsheet.helperCopy.errorCheck.cell"),
    found: translate("spreadsheet.helperCopy.errorCheck.found"),
    previous: translate("spreadsheet.helperCopy.errorCheck.previous"),
    next: translate("spreadsheet.helperCopy.errorCheck.next"),
    calculationSteps: translate("spreadsheet.helperCopy.errorCheck.calculationSteps"),
    editFormula: translate("spreadsheet.helperCopy.errorCheck.editFormula"),
    close: translate("spreadsheet.helperCopy.errorCheck.close"),
  };
  const ruleText: Record<ErrorCheckRule, string> = {
    "error-value": translate("spreadsheet.helperCopy.errorCheck.rule.error-value"),
    "inconsistent-calculated-column": translate("spreadsheet.helperCopy.errorCheck.rule.inconsistent-calculated-column"),
    "inconsistent-formula": translate("spreadsheet.helperCopy.errorCheck.rule.inconsistent-formula"),
    "omits-adjacent-cells": translate("spreadsheet.helperCopy.errorCheck.rule.omits-adjacent-cells"),
    "number-stored-as-text": translate("spreadsheet.helperCopy.errorCheck.rule.number-stored-as-text"),
    "formula-references-empty": translate("spreadsheet.helperCopy.errorCheck.rule.formula-references-empty"),
    "unlocked-formula": translate("spreadsheet.helperCopy.errorCheck.rule.unlocked-formula"),
    "two-digit-year": translate("spreadsheet.helperCopy.errorCheck.rule.two-digit-year"),
    "invalid-table-data": translate("spreadsheet.helperCopy.errorCheck.rule.invalid-table-data"),
  };
  const [index, setIndex] = useState(0);

  useEffect(() => {
    if (open) setIndex(0);
  }, [open, findings]);

  const current = findings[index] ?? null;
  const label = useMemo(
    () => (current ? `${columnLabel(current.col)}${current.row + 1}` : ""),
    [current, columnLabel],
  );

  // Selecting on arrival is what makes Previous/Next feel like Excel's walk.
  useEffect(() => {
    if (open && current) onSelect(current.row, current.col);
    // Only the pointer moving should re-select, not an unrelated re-render.
  }, [open, current, onSelect]);

  if (!open) return null;

  return (
    <div
      className={officeDialogClassName({ tone: "spreadsheet", size: "md" })}
      role="dialog"
      aria-label={text.title}
      data-testid="spreadsheet-error-check-panel"
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault();
          onClose();
        }
      }}
      style={{ position: "absolute", top: "calc(100% + 4px)", left: 0, zIndex: 30 }}
    >
      <span className={dialogParts.title}>{text.title}</span>
      {current ? (
        <>
          <span
            data-testid="spreadsheet-error-check-position"
            className={dialogParts.hint}
          >
            {text.position(index + 1, findings.length)}
          </span>
          <span
            data-testid="spreadsheet-error-check-cell"
            className={dialogParts.sectionTitle}
          >
            {text.cell}: {label}
          </span>
          <span
            data-testid="spreadsheet-error-check-reason"
            className={dialogParts.hint}
          >
            {ruleText[current.rule]}
          </span>
          <code
            data-testid="spreadsheet-error-check-detail"
            className={dialogParts.list}
            style={{
              display: "block",
              padding: "4px 6px",
              whiteSpace: "pre-wrap",
              wordBreak: "break-all",
            }}
          >
            {current.detail}
          </code>
        </>
      ) : (
        <span
          data-testid="spreadsheet-error-check-empty"
          className={dialogParts.hint}
        >
          {text.none}
        </span>
      )}
      {current ? (
        <div className={dialogParts.row} style={{ gap: 6 }}>
          <button
            type="button"
            data-testid="spreadsheet-error-check-calculation-steps"
            disabled={!current.detail.startsWith("=")}
            onClick={() => onShowCalculationSteps(current)}
            className={dialogParts.compactButton}
          >
            {text.calculationSteps}
          </button>
          <button
            type="button"
            data-testid="spreadsheet-error-check-edit-formula"
            onClick={() => onEditInFormulaBar(current)}
            className={dialogParts.compactButton}
          >
            {text.editFormula}
          </button>
        </div>
      ) : null}
      <div className={dialogParts.footer}>
        <button
          type="button"
          data-testid="spreadsheet-error-check-prev"
          disabled={index <= 0}
          onClick={() => setIndex((value) => Math.max(0, value - 1))}
          className={dialogParts.button}
        >
          {text.previous}
        </button>
        <button
          type="button"
          data-testid="spreadsheet-error-check-next"
          disabled={index >= findings.length - 1}
          onClick={() =>
            setIndex((value) => Math.min(findings.length - 1, value + 1))
          }
          className={dialogParts.button}
        >
          {text.next}
        </button>
        <button
          type="button"
          data-testid="spreadsheet-error-check-close"
          onClick={onClose}
          className={dialogParts.button}
        >
          {text.close}
        </button>
      </div>
    </div>
  );
}

export default ErrorCheckPanel;
