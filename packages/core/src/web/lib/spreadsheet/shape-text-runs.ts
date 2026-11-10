/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Character formatting of PART of a sheet shape's words.
 *
 * Excel formats the characters selected inside a shape's text, not the whole
 * shape: with a caret in the words, Home ▸ Font and Shape Format ▸ WordArt
 * Styles act on the selected characters (a caret: its word), and with the
 * shape itself selected on all of them (Microsoft Support, "Change the font
 * size" / "Add or remove WordArt"; DrawingML states each run's format on its
 * own `a:rPr`, ECMA-376 §21.1.2.3.9). A sheet shape kept one whole-shape
 * format (`textBold`, `textSize`, `textPaint`, …), so Bold on a typed word
 * bolded the whole label.
 *
 * The words are now also `textRuns`, in the SAME run model, normaliser and
 * field a Docs shape uses (`lib/document-shape-text-runs`, edited in the
 * shared `OfficeShapeTextField`), restricted to what the workbook's `a:rPr`
 * writers state: `b`, `i`, `u`, `strike`, `sz`, the solid fill, the
 * typefaces and the WordArt paint (`a:ln`, fill, effects). A run's field
 * overrides the shape's own; a field a run leaves out is the shape's. Runs
 * are kept only while they spell exactly the shape's `text`. Sizes are CSS
 * px, as the shape's `textSize` is.
 */

import { normalizeDocumentShapeRunStyle, normalizeDocumentShapeTextRuns } from "../document-shape-text-runs";
import { patchOfficeTextPaint } from "../office-text-paint";
import { cellCharacterStyleAt, cellTextRunsInRange } from "../presentation-cell-rich-text";
import type { RunMath, RunStyle, TextRun } from "../presentation-runs";

/** The most LaTeX one equation among a shape's words keeps: the equation
 * text box's own bound (`MAX_SHAPE_EQUATION_LATEX_LENGTH`, the server's). */
const MAX_SHAPE_RUN_MATH_LATEX_LENGTH = 8_000;

/** The run fields a workbook shape's `a:rPr` writers state; `underlineStyle`
 * is a double underline (`a:rPr@u="dbl"`) only. */
export const SPREADSHEET_SHAPE_RUN_STYLE_KEYS = [
  "bold",
  "italic",
  "underline",
  "underlineStyle",
  "strike",
  "color",
  "fontSize",
  "fontFamily",
  "textPaint",
] as const satisfies readonly (keyof RunStyle)[];

export type SpreadsheetShapeRunStyleKey = (typeof SPREADSHEET_SHAPE_RUN_STYLE_KEYS)[number];

const KEPT = new Set<string>(SPREADSHEET_SHAPE_RUN_STYLE_KEYS);

/** Which characters of `value`'s runs are underlined with a double line. */
function doubleUnderlined(value: unknown): boolean[] {
  if (!Array.isArray(value)) return [];
  return value.flatMap((run) => {
    const record = run && typeof run === "object" ? (run as { text?: unknown; style?: RunStyle }) : {};
    const text = typeof record.text === "string" ? record.text : "";
    const double = record.style?.underline === true && record.style.underlineStyle === "double";
    return Array.from({ length: text.length }, () => double);
  });
}









export function spreadsheetShapeRunMath(value: unknown): RunMath | null {
  if (!value || typeof value !== "object") return null;
  const { latex, display } = value as { latex?: unknown; display?: unknown };
  if (typeof latex !== "string" || !latex.trim() || latex.length > MAX_SHAPE_RUN_MATH_LATEX_LENGTH) return null;
  return display === true ? { latex, display: true } : { latex };
}

/** A run's style held to what a workbook run states. */
function keptRunStyle(style: RunStyle | undefined): Record<string, unknown> | undefined {
  if (!style) return undefined;
  const kept: Record<string, unknown> = {};
  for (const [key, field] of Object.entries(style)) {
    if (KEPT.has(key)) kept[key] = field;
  }
  return kept;
}

/**
 * Runs that hold equations: each equation run stays ONE atomic run (it never
 * merges with the words beside it, as a zone is one element in the file),
 * and the words between equations are normalized as any shape's words are.
 * Every run is kept, formatted or not, since the equations alone are worth
 * keeping. `undefined` when the runs do not spell `text` exactly.
 */
function normalizeRunsWithEquations(value: readonly unknown[], text: string): TextRun[] | undefined {
  const out: TextRun[] = [];
  let words: unknown[] = [];
  // The words since the last equation, normalized as one stretch (formatted
  // or not). False when one of them is not a run.
  const flushWords = (): boolean => {
    const stretch = words;
    words = [];
    if (!stretch.length) return true;
    const spelled = stretch.map((run) =>
      run && typeof run === "object" && typeof (run as { text?: unknown }).text === "string"
        ? ((run as { text: string }).text)
        : null,
    );
    if (spelled.some((part) => part === null)) return false;
    const segment = spelled.join("");
    if (segment) out.push(...(normalizeFormattedRuns(stretch, segment) ?? [{ text: segment }]));
    return true;
  };
  for (const raw of value) {
    const record = raw && typeof raw === "object" ? (raw as { text?: unknown; style?: unknown; math?: unknown }) : null;
    const math = record ? spreadsheetShapeRunMath(record.math) : null;
    if (!record || !math) {
      words.push(raw);
      continue;
    }
    if (!flushWords()) return undefined;
    // An equation's characters are one line of the zone (`a14:m` holds no
    // line break).
    if (typeof record.text !== "string" || !record.text || /[\n\r\v]/.test(record.text)) return undefined;
    const style = normalizeDocumentShapeRunStyle(keptRunStyle(record.style as RunStyle | undefined));
    out.push(style ? { text: record.text, style, math } : { text: record.text, math });
  }
  if (!flushWords()) return undefined;
  return out.map((run) => run.text).join("") === text ? out : undefined;
}

/**
 * The shape's runs, or `undefined` when they carry no formatting of the kinds
 * a workbook states per run or do not spell `text` exactly (a stale reading
 * is dropped, never applied to other words). The server twin is
 * `normalize_spreadsheet_shape_text_runs` (document_shape_text_runs.py).
 * Runs that hold an equation (`math`) are kept whole and atomic.
 */
export function normalizeSpreadsheetShapeTextRuns(
  value: unknown,
  text: string,
): TextRun[] | undefined {
  if (
    Array.isArray(value) &&
    value.some((run) => run && typeof run === "object" && spreadsheetShapeRunMath((run as { math?: unknown }).math))
  ) {
    return normalizeRunsWithEquations(value, text);
  }
  return normalizeFormattedRuns(value, text);
}

/** `normalizeSpreadsheetShapeTextRuns` for runs that hold no equation. */
function normalizeFormattedRuns(value: unknown, text: string): TextRun[] | undefined {
  const runs = normalizeDocumentShapeTextRuns(value, text);
  if (!runs) return undefined;
  const kept = normalizeDocumentShapeTextRuns(
    runs.map((run) => {
      if (!run.style) return { text: run.text };
      const style: Record<string, unknown> = {};
      for (const [key, field] of Object.entries(run.style)) {
        if (KEPT.has(key)) style[key] = field;
      }
      return { text: run.text, style };
    }),
    text,
  );
  const doubled = doubleUnderlined(value);
  if (!kept || !doubled.some(Boolean)) return kept;
  // The Docs model states no line style: a run's double line goes back on
  // the characters that had it, as their own run.
  const out: TextRun[] = [];
  let at = 0;
  for (const run of kept) {
    for (const char of run.text) {
      const style = doubled[at] && run.style?.underline === true
        ? ({ ...run.style, underlineStyle: "double" } as RunStyle)
        : run.style;
      at += char.length;
      const last = out.at(-1);
      if (last && JSON.stringify(last.style) === JSON.stringify(style)) last.text += char;
      else out.push(style ? { text: char, style } : { text: char });
    }
  }
  return out;
}

/** `piece` still the equation `run` is, when `run` is one (formatting a shape
 * reformats an equation's characters and keeps the equation). */
function withEquation(piece: TextRun, run: TextRun): TextRun {
  return run.math ? { ...piece, math: run.math } : piece;
}

/**
 * The runs without the given fields: what a command on the SELECTED shape
 * leaves, since it formats all of the words and a run's own value must no
 * longer override it (Excel rewrites every run of the label).
 */
export function spreadsheetShapeTextRunsWithout(
  runs: readonly TextRun[] | null | undefined,
  text: string,
  keys: readonly SpreadsheetShapeRunStyleKey[],
): TextRun[] | undefined {
  if (!runs?.length) return undefined;
  const dropped = new Set<string>(keys.includes("underline") ? [...keys, "underlineStyle"] : keys);
  return normalizeSpreadsheetShapeTextRuns(
    runs.map((run) => {
      if (!run.style) return withEquation({ text: run.text }, run);
      const style = Object.fromEntries(Object.entries(run.style).filter(([key]) => !dropped.has(key)));
      return withEquation({ text: run.text, style }, run);
    }),
    text,
  );
}

/**
 * Font Color / Text Fill on the selected shape: every run's own colour goes
 * (the label's is the one), and a run that wears WordArt of its own takes the
 * colour as that paint's fill, keeping its outline and effects.
 */
export function spreadsheetShapeTextRunsFilled(
  runs: readonly TextRun[] | null | undefined,
  text: string,
  fill: string,
): TextRun[] | undefined {
  if (!runs?.length) return undefined;
  return normalizeSpreadsheetShapeTextRuns(
    runs.map((run) => {
      if (!run.style) return withEquation({ text: run.text }, run);
      const style = { ...run.style } as RunStyle;
      delete style.color;
      if (style.textPaint) {
        try {
          style.textPaint = patchOfficeTextPaint(style.textPaint, { fill });
        } catch {
          delete style.textPaint;
        }
      }
      return withEquation({ text: run.text, style }, run);
    }),
    text,
  );
}

/**
 * The shape's runs carried into new words set as a whole (Set Text from a
 * pane, an older editor, an assistant): the characters before and after the
 * change keep their runs, and the words put in take the format at the
 * change, as text typed there would. Undefined when the runs do not spell
 * `before` (nothing to follow) or nothing formatted is left.
 */
export function spreadsheetShapeTextRunsFollowing(
  runs: readonly TextRun[] | null | undefined,
  before: string,
  after: string,
): TextRun[] | undefined {
  if (!runs?.length || runs.map((run) => run.text).join("") !== before) return undefined;
  if (before === after) return normalizeSpreadsheetShapeTextRuns(runs, after);
  let prefix = 0;
  while (prefix < before.length && prefix < after.length && before[prefix] === after[prefix]) prefix += 1;
  let suffix = 0;
  while (
    suffix < before.length - prefix &&
    suffix < after.length - prefix &&
    before[before.length - 1 - suffix] === after[after.length - 1 - suffix]
  )
    suffix += 1;
  const inserted = after.slice(prefix, after.length - suffix);
  const style = cellCharacterStyleAt(runs, prefix);
  return normalizeSpreadsheetShapeTextRuns(
    [
      ...cellTextRunsInRange(runs, { start: 0, end: prefix }),
      ...(inserted ? [style ? { text: inserted, style } : { text: inserted }] : []),
      ...cellTextRunsInRange(runs, { start: before.length - suffix, end: before.length }),
    ],
    after,
  );
}

/** The first `max` characters of `runs`. */
export function spreadsheetShapeTextRunsClipped(runs: readonly TextRun[], max: number): TextRun[] {
  const out: TextRun[] = [];
  let left = max;
  for (const run of runs) {
    if (left <= 0) break;
    const text = run.text.slice(0, left);
    left -= text.length;
    // An equation cut short is no equation: it is dropped whole.
    if (run.math && text.length < run.text.length) break;
    out.push(withEquation(run.style ? { text, style: run.style } : { text }, run));
  }
  return out;
}
