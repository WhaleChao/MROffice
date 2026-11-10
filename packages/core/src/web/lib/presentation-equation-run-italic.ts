/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  equationLatexWithLetterFormat,
  equationLettersFormatState,
  type EquationLetterFormatKind,
} from "./office-equation-letter-format";
import type { TextRun } from "./presentation-runs";

type RunContainer = {
  runs?: unknown;
  paragraphs?: unknown;
  style?: unknown;
};

/** The block's equation runs, each with a key that names where it sits. */
function mapEquationRuns<T>(block: T, map: (run: TextRun, key: string) => TextRun): T {
  const record = block as unknown as RunContainer;
  let changed = false;
  const mapList = (runs: unknown, prefix: string): unknown =>
    Array.isArray(runs)
      ? runs.map((run: TextRun, index) => {
          if (!run || typeof run !== "object" || !run.math) return run;
          const next = map(run, `${prefix}${index}`);
          if (next !== run) changed = true;
          return next;
        })
      : runs;
  const runs = mapList(record.runs, "r");
  const paragraphs = Array.isArray(record.paragraphs)
    ? record.paragraphs.map((paragraph: unknown, at) =>
        paragraph && typeof paragraph === "object" && Array.isArray((paragraph as RunContainer).runs)
          ? { ...(paragraph as object), runs: mapList((paragraph as RunContainer).runs, `p${at}:`) }
          : paragraph,
      )
    : record.paragraphs;
  if (!changed) return block;
  return {
    ...(block as object),
    ...(Array.isArray(record.runs) ? { runs } : {}),
    ...(Array.isArray(record.paragraphs) ? { paragraphs } : {}),
  } as T;
}

/** Whether the block holds an equation run. */
export function blockHoldsEquationRun(block: unknown): boolean {
  let found = false;
  mapEquationRuns(block, (run) => {
    found = true;
    return run;
  });
  return found;
}

export type EquationItalicStates = ReadonlyMap<string, boolean>;

/** The formats an equation run states in its letters rather than as a run flag. */
export type EquationRunFormat = Extract<EquationLetterFormatKind, "bold" | "italic" | "underline">;

/**
 * An equation run's `kind` as the button reads it. Italic is its letters'
 * (an equation without letters reads italic, so it never holds a toggle
 * off); Bold and Underline are on when its letters say so or the run's own
 * flag does (a flag written before the letters carried them).
 */
function equationRunFormatState(run: TextRun, kind: EquationRunFormat): boolean {
  const letters = equationLettersFormatState(run.math!.latex, kind);
  if (kind === "italic") return letters ?? true;
  const flag = run.style?.[kind] === true;
  return flag || (letters ?? false);
}

/**
 * The block with each equation run's `kind` stated as its run flag, and
 * those states.
 */
export function withEquationFormatMaterialized<T>(
  block: T,
  kind: EquationRunFormat,
): { readonly block: T; readonly states: EquationItalicStates } {
  const states = new Map<string, boolean>();
  const next = mapEquationRuns(block, (run, key) => {
    const on = equationRunFormatState(run, kind);
    states.set(key, on);
    return { ...run, style: { ...(run.style ?? {}), [kind]: on } };
  });
  return { block: next, states };
}

/**
 * The block as the buttons read it: every equation run's Bold, Italic and
 * Underline stated as its run flags (`withEquationFormatMaterialized`).
 */
export function withEquationFormatsRead<T>(block: T): T {
  if (!blockHoldsEquationRun(block)) return block;
  let read = block;
  for (const kind of ["bold", "italic", "underline"] as const) {
    read = withEquationFormatMaterialized(read, kind).block;
  }
  return read;
}

/**
 * The toggled block with what the toggle wrote on each equation run put into
 * its letters, and the run flags taken off again.
 */
export function withEquationFormatSettled<T>(
  block: T,
  states: EquationItalicStates,
  kind: EquationRunFormat,
  /** The object's own format was written (the box selected): every run takes it. */
  wholeObject = false,
): T {
  const boxOn = ((block as unknown as RunContainer).style as Record<string, unknown> | undefined)?.[kind] === true;
  return mapEquationRuns(block, (run, key) => {
    const before = states.get(key);
    const own = run.style?.[kind];
    const on = !wholeObject && typeof own === "boolean" ? own : boxOn;
    const style = { ...(run.style ?? {}) };
    delete style[kind];
    if (kind === "underline") delete style.underlineStyle;
    // The letters take the new state; a flag the run carried before the
    // letters did is gone with it (Office Math keeps format in the math).
    const latex =
      before === undefined || on === before
        ? run.math!.latex
        : equationLatexWithLetterFormat(run.math!.latex, kind, on);
    const next: TextRun = { ...run, math: { ...run.math!, latex } };
    if (Object.keys(style).length) next.style = style;
    else delete next.style;
    return next;
  });
}

/** `toggle` run over the block with its equations' `kind` taken as Office does. */
export function withEquationFormatToggle<T>(
  block: T,
  kind: EquationRunFormat,
  toggle: (materialized: T) => T | null,
): T | null {
  if (!blockHoldsEquationRun(block)) return toggle(block);
  const { block: materialized, states } = withEquationFormatMaterialized(block, kind);
  const toggled = toggle(materialized);
  return toggled ? withEquationFormatSettled(toggled, states, kind) : null;
}

/** Italic's own names for the above. */
export const withEquationItalicMaterialized = <T,>(block: T) => withEquationFormatMaterialized(block, "italic");
export const withEquationItalicSettled = <T,>(block: T, states: EquationItalicStates, wholeObject = false): T =>
  withEquationFormatSettled(block, states, "italic", wholeObject);
export const withEquationItalicToggle = <T,>(block: T, toggle: (materialized: T) => T | null): T | null =>
  withEquationFormatToggle(block, "italic", toggle);
