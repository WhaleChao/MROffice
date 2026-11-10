"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * "Evaluate Formula" / 公式求值 — the Formulas-tab walkthrough Excel and WPS
 * both ship.
 *
 * The dialog owns only the WALK: which span is underlined, what has been
 * substituted so far, and the step-in stack. Every value comes from the host's
 * `evaluate` callback, which runs the same engine the grid uses, so a step can
 * never report a number the cell would not show.
 *
 * Display copy comes from the caller's active message catalogue.
 */

import { useCallback, useEffect, useMemo, useRef, useState } from "react";

import {
  officeDialogParts as dialogParts,
  spreadsheetDialogClassName,
} from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";

import {
  MAX_EVALUATION_STEPS,
  applyEvaluationStep,
  formatEvaluatedLiteral,
  nextEvaluationTarget,
  stepIntoTarget,
  type EvaluationTarget,
} from "./formula-evaluation-steps";

export interface FormulaEvaluationLocation {
  /** Stable worksheet identity; names may change while the dialog is open. */
  sheetId: string;
  /** Exact cell whose formula supplies this evaluation frame. */
  row: number;
  col: number;
}

export interface FormulaEvaluationSource extends FormulaEvaluationLocation {
  /** Display label of the cell being explained, e.g. "Sheet1!B4". */
  label: string;
  /** The formula body, without the leading "=". */
  expression: string;
}

export interface FormulaEvaluationReference
  extends FormulaEvaluationLocation {
  label: string;
  expression: string;
}

export interface FormulaEvaluationDialogProps {
  open: boolean;
  translate: (key: string, values?: Record<string, string | number>) => string;
  /** The cell the user asked about. */
  source: FormulaEvaluationSource | null;
  /**
   * Evaluate one sub-expression in its source cell's context. Returning
   * undefined means the engine refused it; the dialog then stops rather than
   * inventing a value. A promise (the calculation session answers it) holds
   * the step buttons until it settles.
   */
  evaluate: (
    expression: string,
    location: FormulaEvaluationLocation,
  ) => unknown;
  /**
   * The formula body at a referenced cell, or null when that cell holds a
   * constant (Step In is then unavailable, exactly as in Excel).
   */
  getFormulaAt?: (
    from: FormulaEvaluationLocation,
    sheet: string | null,
    ref: string,
  ) => FormulaEvaluationReference | null;
  onClose: () => void;
}

interface Frame extends FormulaEvaluationLocation {
  label: string;
  expression: string;
  /** Exact span in the parent expression that opened this child frame. */
  returnTarget?: EvaluationTarget;
}

function hasUnsupportedReferenceScope(expression: string): boolean {
  // Do not inspect worksheet-looking text inside Excel string literals. Keep
  // offsets intact while masking them so quoted sheet names remain visible.
  const executable = [...expression];
  let inString = false;
  for (let index = 0; index < executable.length; index += 1) {
    if (executable[index] === '"') {
      executable[index] = " ";
      if (inString && executable[index + 1] === '"') {
        executable[index + 1] = " ";
        index += 1;
      } else {
        inString = !inString;
      }
    } else if (inString) {
      executable[index] = " ";
    }
  }

  const formula = executable.join("");
  for (let bang = formula.indexOf("!"); bang >= 0; bang = formula.indexOf("!", bang + 1)) {
    if (!/^\s*\$?[A-Za-z]{1,3}\$?[0-9]{1,7}/.test(formula.slice(bang + 1))) {
      continue;
    }
    const beforeBang = formula.slice(0, bang);
    const quoted = /'((?:[^']|'')*)'$/.exec(beforeBang);
    const unquoted = /([^\s+\-*/^(),=<>;&]+)$/.exec(beforeBang);
    const qualifier = quoted?.[1] ?? unquoted?.[1] ?? "";
    if (
      (qualifier.includes("[") && qualifier.includes("]")) ||
      qualifier.includes(":")
    ) {
      return true;
    }
  }
  return false;
}

function isSameEvaluationTarget(
  current: EvaluationTarget,
  expected: EvaluationTarget,
): boolean {
  return (
    current.start === expected.start &&
    current.end === expected.end &&
    current.kind === expected.kind &&
    current.text === expected.text
  );
}

function isSameEvaluationReference(
  current: FormulaEvaluationReference,
  expected: FormulaEvaluationReference,
): boolean {
  return (
    current.sheetId === expected.sheetId &&
    current.row === expected.row &&
    current.col === expected.col &&
    current.label === expected.label &&
    current.expression === expected.expression
  );
}

export function FormulaEvaluationDialog({
  open,
  translate,
  source,
  evaluate,
  getFormulaAt,
  onClose,
}: FormulaEvaluationDialogProps) {
  const text = {
    title: translate("spreadsheet.helperCopy.formulaEvaluation.title"),
    reference: translate("spreadsheet.helperCopy.formulaEvaluation.reference"),
    evaluation: translate("spreadsheet.helperCopy.formulaEvaluation.evaluation"),
    evaluate: translate("spreadsheet.helperCopy.formulaEvaluation.evaluate"),
    stepIn: translate("spreadsheet.helperCopy.formulaEvaluation.stepIn"),
    stepOut: translate("spreadsheet.helperCopy.formulaEvaluation.stepOut"),
    restart: translate("spreadsheet.helperCopy.formulaEvaluation.restart"),
    close: translate("spreadsheet.helperCopy.formulaEvaluation.close"),
    done: translate("spreadsheet.helperCopy.formulaEvaluation.done"),
    failed: translate("spreadsheet.helperCopy.formulaEvaluation.failed"),
    hint: translate("spreadsheet.helperCopy.formulaEvaluation.hint"),
    limit: translate("spreadsheet.helperCopy.formulaEvaluation.limit"),
  };
  const [stack, setStack] = useState<Frame[]>([]);
  const [status, setStatus] = useState<"idle" | "done" | "failed" | "limit">(
    "idle",
  );
  const [steps, setSteps] = useState(0);
  // A step the calculation session is answering; any other change of the
  // stack (Restart, a new source) makes its answer stale.
  const [busy, setBusy] = useState(false);
  const generationRef = useRef(0);

  const reset = useCallback(() => {
    generationRef.current += 1;
    setBusy(false);
    setStack(source ? [source] : []);
    setStatus("idle");
    setSteps(0);
  }, [source]);

  /** Apply `evaluate`'s answer now, or when its promise settles. */
  const settle = useCallback((answer: unknown, apply: (value: unknown) => void) => {
    if (!(answer instanceof Promise)) { apply(answer); return; }
    const generation = ++generationRef.current;
    setBusy(true);
    void answer.then(
      (value) => { if (generation === generationRef.current) { setBusy(false); apply(value); } },
      () => { if (generation === generationRef.current) { setBusy(false); setStatus("failed"); } },
    );
  }, []);

  useEffect(() => {
    if (open) reset();
  }, [open, reset]);

  const frame = stack.at(-1) ?? null;
  const target = useMemo<EvaluationTarget | null>(
    () =>
      frame &&
      status === "idle" &&
      !hasUnsupportedReferenceScope(frame.expression)
        ? nextEvaluationTarget(frame.expression)
        : null,
    [frame, status],
  );

  const stepIn = useMemo(() => {
    if (!frame || !target || !getFormulaAt) return null;
    const reference = stepIntoTarget(target);
    if (!reference) return null;
    return getFormulaAt(frame, reference.sheet, reference.ref);
  }, [frame, target, getFormulaAt]);

  const handleEvaluate = useCallback(() => {
    if (!frame || !target) return;
    if (steps >= MAX_EVALUATION_STEPS) {
      setStatus("limit");
      return;
    }
    settle(evaluate(target.text, frame), (value) => {
      if (value === undefined) {
        setStatus("failed");
        return;
      }
      const next = applyEvaluationStep(
        frame.expression,
        target,
        formatEvaluatedLiteral(value),
      );
      setSteps((count) => count + 1);
      setStack((current) => [
        ...current.slice(0, -1),
        { ...frame, expression: next },
      ]);
      if (!nextEvaluationTarget(next)) setStatus("done");
    });
  }, [evaluate, frame, settle, steps, target]);

  const handleStepIn = useCallback(() => {
    if (!frame || !stepIn || !target || !getFormulaAt) return;
    const currentTarget = nextEvaluationTarget(frame.expression);
    if (!currentTarget || !isSameEvaluationTarget(currentTarget, target)) {
      return;
    }
    const reference = stepIntoTarget(currentTarget);
    if (!reference) return;
    const currentStepIn = getFormulaAt(
      frame,
      reference.sheet,
      reference.ref,
    );
    if (
      !currentStepIn ||
      !isSameEvaluationReference(currentStepIn, stepIn)
    ) {
      return;
    }
    setStack((current) => [
      ...current,
      {
        ...currentStepIn,
        returnTarget: currentTarget,
      },
    ]);
    setStatus("idle");
  }, [frame, getFormulaAt, stepIn, target]);

  const handleStepOut = useCallback(() => {
    if (!frame || stack.length < 2 || !frame.returnTarget) {
      return;
    }
    if (steps >= MAX_EVALUATION_STEPS) {
      setStatus("limit");
      return;
    }
    const parent = stack[stack.length - 2];
    const returnTarget = frame.returnTarget;
    if (
      parent.expression.slice(returnTarget.start, returnTarget.end) !==
      returnTarget.text
    ) {
      setStatus("failed");
      return;
    }
    settle(evaluate(frame.expression, frame), (value) => {
      if (value === undefined) {
        setStatus("failed");
        return;
      }
      const next = applyEvaluationStep(
        parent.expression,
        returnTarget,
        formatEvaluatedLiteral(value),
      );
      setSteps((count) => count + 1);
      setStack([
        ...stack.slice(0, -2),
        { ...parent, expression: next },
      ]);
      setStatus(nextEvaluationTarget(next) ? "idle" : "done");
    });
  }, [evaluate, frame, settle, stack, steps]);

  if (!open || !frame) return null;

  const before = target
    ? frame.expression.slice(0, target.start)
    : frame.expression;
  const highlighted = target
    ? frame.expression.slice(target.start, target.end)
    : "";
  const after = target ? frame.expression.slice(target.end) : "";

  return (
    <div
      className={spreadsheetDialogClassName("lg")}
      role="dialog"
      aria-modal="true"
      aria-label={text.title}
      data-testid="spreadsheet-evaluate-formula-dialog"
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault();
          onClose();
        }
      }}
    >
      <span className={dialogParts.title}>{text.title}</span>
      <span
        data-testid="spreadsheet-evaluate-formula-reference"
        className={dialogParts.hint}
      >
        {text.reference}: {frame.label}
      </span>
      <code
        data-testid="spreadsheet-evaluate-formula-expression"
        className={dialogParts.list}
        style={{
          display: "block",
          padding: "6px 8px",
          maxHeight: 140,
          lineHeight: 1.5,
          whiteSpace: "pre-wrap",
          wordBreak: "break-all",
        }}
      >
        {"="}
        {before}
        {target ? (
          <span
            data-testid="spreadsheet-evaluate-formula-next"
            style={{
              textDecoration: "underline",
              textUnderlineOffset: 2,
              fontWeight: 600,
            }}
          >
            {highlighted}
          </span>
        ) : null}
        {after}
      </code>
      <span
        data-testid="spreadsheet-evaluate-formula-status"
        className={dialogParts.hint}
      >
        {status === "done"
          ? text.done
          : status === "failed"
            ? text.failed
            : status === "limit"
              ? text.limit
              : text.hint}
      </span>
      <div className={dialogParts.footer}>
        <button
          type="button"
          data-testid="spreadsheet-evaluate-formula-step"
          disabled={!target || busy}
          aria-busy={busy || undefined}
          onClick={handleEvaluate}
          className={dialogParts.primaryButton}
        >
          {text.evaluate}
        </button>
        <button
          type="button"
          data-testid="spreadsheet-evaluate-formula-step-in"
          disabled={!stepIn || busy}
          onClick={handleStepIn}
          className={dialogParts.button}
        >
          {text.stepIn}
        </button>
        <button
          type="button"
          data-testid="spreadsheet-evaluate-formula-step-out"
          disabled={stack.length < 2 || !frame.returnTarget || busy}
          onClick={handleStepOut}
          className={dialogParts.button}
        >
          {text.stepOut}
        </button>
        <button
          type="button"
          data-testid="spreadsheet-evaluate-formula-restart"
          onClick={reset}
          className={dialogParts.button}
        >
          {text.restart}
        </button>
        <button
          type="button"
          data-testid="spreadsheet-evaluate-formula-close"
          onClick={onClose}
          className={dialogParts.button}
        >
          {text.close}
        </button>
      </div>
    </div>
  );
}

export default FormulaEvaluationDialog;
