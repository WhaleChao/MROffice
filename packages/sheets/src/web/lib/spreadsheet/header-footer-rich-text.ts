/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Editing the formatted text of one header/footer section (Excel's Custom
 * Header / Custom Footer boxes; Microsoft Support, "Headers and footers in a
 * worksheet"): typing keeps the format of the characters it does not touch,
 * typed characters take the format of the character before them, and Format
 * Text formats a character range.
 *
 * The persisted model is the flat {@link PrintHeaderFooterRun} (one record per
 * run, sizes in points, single or double underline). Editing works on the
 * `{ text, style }` shape the cell rich-text helpers already use, so the
 * edit-preserving rule is the very one cells use
 * ({@link editedRunsKeepingFormat}); only the style vocabulary differs.
 *
 * Pure (no React, no DOM) so the rules are unit-testable on their own.
 */
import type { CSSProperties } from "react";

import { editedRunsKeepingFormat } from "@mrnote/core/web/lib/spreadsheet/cell-rich-text.ts";
import {
  headerFooterRunsForSegment,
  normalizeHeaderFooter,
  type PrintHeaderFooter,
  type PrintHeaderFooterRun,
} from "@mrnote/core/web/lib/spreadsheet/print-settings.ts";

/** A run's formatting: everything a {@link PrintHeaderFooterRun} carries but its text. */
export type HeaderFooterRunStyle = Omit<PrintHeaderFooterRun, "text">;

/** One run in the editing shape the cell rich-text helpers share. */
export interface HeaderFooterStyledRun {
  text: string;
  style?: HeaderFooterRunStyle;
}

/** The formatting fields the Font dialog reads and writes. */
export const HEADER_FOOTER_FORMAT_FIELDS = [
  "bold",
  "italic",
  "underline",
  "doubleUnderline",
  "strikethrough",
  "superscript",
  "subscript",
  "fontFamily",
  "fontSize",
  "color",
] as const satisfies ReadonlyArray<keyof HeaderFooterRunStyle>;

export type HeaderFooterFormatField = (typeof HEADER_FOOTER_FORMAT_FIELDS)[number];

/**
 * A format change: a field set to a value turns it on, a field present with
 * `undefined` / `false` takes it back to the default.
 */
export type HeaderFooterFormatPatch = {
  [K in HeaderFooterFormatField]?: HeaderFooterRunStyle[K] | false | undefined;
};

/** A persisted run's formatting, without its text; undefined when it has none. */
export function headerFooterRunStyle(
  run: PrintHeaderFooterRun,
): HeaderFooterRunStyle | undefined {
  const style: Record<string, unknown> = {};
  for (const key of HEADER_FOOTER_FORMAT_FIELDS) {
    if (run[key] !== undefined) style[key] = run[key];
  }
  return Object.keys(style).length > 0 ? (style as HeaderFooterRunStyle) : undefined;
}

function sameStyle(
  a: HeaderFooterRunStyle | undefined,
  b: HeaderFooterRunStyle | undefined,
): boolean {
  const left = a ?? {};
  const right = b ?? {};
  return HEADER_FOOTER_FORMAT_FIELDS.every((key) => left[key] === right[key]);
}

/**
 * A section's runs in the editing shape. Runs that do not spell `text` (or no
 * runs) read as one unformatted run, so the first format on a plain section
 * works.
 */
export function headerFooterStyledRuns(
  text: string,
  runs: readonly PrintHeaderFooterRun[] | undefined,
): HeaderFooterStyledRun[] {
  if (runs && runs.length > 0 && runs.map((run) => run.text).join("") === text) {
    return runs
      .filter((run) => run.text !== "")
      .map((run) => {
        const style = headerFooterRunStyle(run);
        return style ? { text: run.text, style } : { text: run.text };
      });
  }
  return text ? [{ text }] : [];
}

/**
 * The persisted runs for the editing shape: neighbours with the same format
 * merged, empty runs dropped, and undefined when nothing carries formatting
 * (the section is then its plain text alone).
 */
export function headerFooterRunsFromStyled(
  runs: readonly HeaderFooterStyledRun[],
): PrintHeaderFooterRun[] | undefined {
  const out: PrintHeaderFooterRun[] = [];
  let previous: HeaderFooterRunStyle | undefined;
  for (const run of runs) {
    if (!run.text) continue;
    const last = out.at(-1);
    if (last && sameStyle(previous, run.style)) {
      last.text += run.text;
      continue;
    }
    out.push({ text: run.text, ...(run.style ?? {}) });
    previous = run.style;
  }
  return out.some((run) => Object.keys(run).length > 1) ? out : undefined;
}

/**
 * The format text typed at `offset` takes: the character before it, or the
 * first character at the very start (Excel's rich edit box).
 */
export function headerFooterStyleAt(
  runs: readonly HeaderFooterStyledRun[],
  offset: number,
): HeaderFooterRunStyle | undefined {
  let end = 0;
  for (const run of runs) {
    end += run.text.length;
    if (offset <= end && (offset > end - run.text.length || offset === 0)) {
      return run.style;
    }
  }
  return runs.at(-1)?.style;
}

function sliceRuns(
  runs: readonly HeaderFooterStyledRun[],
  from: number,
  to: number,
): HeaderFooterStyledRun[] {
  const out: HeaderFooterStyledRun[] = [];
  let offset = 0;
  for (const run of runs) {
    const low = Math.max(from - offset, 0);
    const high = Math.min(to - offset, run.text.length);
    if (low < high) {
      const text = run.text.slice(low, high);
      out.push(run.style ? { text, style: run.style } : { text });
    }
    offset += run.text.length;
  }
  return out;
}

function boundedRange(
  runs: readonly HeaderFooterStyledRun[],
  start: number,
  end: number,
): [number, number] {
  const length = runs.reduce((sum, run) => sum + run.text.length, 0);
  const a = Math.min(length, Math.max(0, Math.round(start)));
  const b = Math.min(length, Math.max(0, Math.round(end)));
  return a <= b ? [a, b] : [b, a];
}

/**
 * Replace the characters `[start, end)` with `inserted`. Typed at a caret the
 * new text takes the format before it; typed over a selection, the format of
 * the first character it replaces; `typingStyle` (a format chosen at the
 * caret with Format Text) overrides both.
 */
export function replaceHeaderFooterRange(
  runs: readonly HeaderFooterStyledRun[],
  start: number,
  end: number,
  inserted: string,
  typingStyle?: HeaderFooterRunStyle | null,
): HeaderFooterStyledRun[] {
  const [from, to] = boundedRange(runs, start, end);
  const total = runs.reduce((sum, run) => sum + run.text.length, 0);
  const style =
    typingStyle !== undefined && typingStyle !== null
      ? Object.keys(typingStyle).length > 0
        ? typingStyle
        : undefined
      : headerFooterStyleAt(runs, to > from ? from + 1 : from);
  return [
    ...sliceRuns(runs, 0, from),
    ...(inserted ? [style ? { text: inserted, style } : { text: inserted }] : []),
    ...sliceRuns(runs, to, total),
  ];
}

/** One run's format with `patch` applied (Excel's exclusive pairs kept apart). */
export function patchHeaderFooterStyle(
  style: HeaderFooterRunStyle | undefined,
  patch: HeaderFooterFormatPatch,
): HeaderFooterRunStyle | undefined {
  const next: Record<string, unknown> = { ...(style ?? {}) };
  for (const key of HEADER_FOOTER_FORMAT_FIELDS) {
    if (!Object.prototype.hasOwnProperty.call(patch, key)) continue;
    const value = patch[key];
    if (value === undefined || value === false || value === "") delete next[key];
    else next[key] = value;
  }
  if (patch.doubleUnderline) delete next.underline;
  if (patch.underline) delete next.doubleUnderline;
  if (patch.superscript) delete next.subscript;
  if (patch.subscript) delete next.superscript;
  return Object.keys(next).length > 0 ? (next as HeaderFooterRunStyle) : undefined;
}

/** Format the characters `[start, end)`; an empty range changes nothing. */
export function formatHeaderFooterRange(
  runs: readonly HeaderFooterStyledRun[],
  start: number,
  end: number,
  patch: HeaderFooterFormatPatch,
): HeaderFooterStyledRun[] {
  const [from, to] = boundedRange(runs, start, end);
  if (from === to) return [...runs];
  const total = runs.reduce((sum, run) => sum + run.text.length, 0);
  return [
    ...sliceRuns(runs, 0, from),
    ...sliceRuns(runs, from, to).map((run) => {
      const style = patchHeaderFooterStyle(run.style, patch);
      return style ? { text: run.text, style } : { text: run.text };
    }),
    ...sliceRuns(runs, to, total),
  ];
}

/**
 * The format a range shows, as the Font dialog opens on it: each field the
 * range's characters agree on, and the fields they disagree on (`mixed`). A
 * collapsed range reads the format typing there would take.
 */
export function headerFooterRangeFormat(
  runs: readonly HeaderFooterStyledRun[],
  start: number,
  end: number,
): { style: HeaderFooterRunStyle; mixed: HeaderFooterFormatField[] } {
  const [from, to] = boundedRange(runs, start, end);
  if (from === to) {
    return { style: { ...(headerFooterStyleAt(runs, from) ?? {}) }, mixed: [] };
  }
  const covered = sliceRuns(runs, from, to);
  const style: Record<string, unknown> = {};
  const mixed: HeaderFooterFormatField[] = [];
  for (const key of HEADER_FOOTER_FORMAT_FIELDS) {
    const values = covered.map((run) => run.style?.[key]);
    if (values.every((value) => value === values[0])) {
      if (values[0] !== undefined) style[key] = values[0];
    } else {
      mixed.push(key);
    }
  }
  return { style: style as HeaderFooterRunStyle, mixed };
}

/**
 * The runs a section keeps after its plain text was edited in place (the
 * Page Layout view's band box): the characters the edit kept keep their
 * format, typed ones take the format before them ({@link editedRunsKeepingFormat}).
 * Undefined when nothing formatted survives.
 */
export function headerFooterRunsForEditedText(
  previousText: string,
  previousRuns: readonly PrintHeaderFooterRun[] | undefined,
  nextText: string,
): PrintHeaderFooterRun[] | undefined {
  const styled = headerFooterStyledRuns(previousText, previousRuns);
  const edited = editedRunsKeepingFormat(styled, nextText);
  return edited ? headerFooterRunsFromStyled(edited) : undefined;
}

/**
 * A band with one section's text (and runs) replaced, normalized the way the
 * page stores bands: trimmed, its runs aligned with the text, blank sections
 * dropped; undefined when nothing is left.
 */
export function headerFooterWithSection(
  band: PrintHeaderFooter | undefined,
  side: "left" | "center" | "right",
  text: string,
  runs: readonly PrintHeaderFooterRun[] | undefined,
): PrintHeaderFooter | undefined {
  const runsKey = `${side}Runs` as const;
  const next: Record<string, unknown> = { ...(band ?? {}), [side]: text };
  if (runs && runs.length > 0) next[runsKey] = [...runs];
  else delete next[runsKey];
  return normalizeHeaderFooter(next);
}

/** Whether a section keeps these runs (they spell it and format it). Its
 * spaces are text like any other (B386). */
export function headerFooterSectionKeepsRuns(
  text: string,
  runs: readonly PrintHeaderFooterRun[] | undefined,
): boolean {
  return headerFooterRunsForSegment(text, runs) !== undefined;
}

/**
 * The CSS one run is drawn with. `pointToPx` turns its size in points into
 * pixels at the surface's scale (the page view scales with the paper, the
 * dialog draws actual size); `defaultPoints` is the size of unformatted text.
 */
export function headerFooterRunCss(
  style: HeaderFooterRunStyle | undefined,
  options: { pointToPx: (points: number) => number; defaultPoints?: number },
): CSSProperties {
  const run = style ?? {};
  const decorations: string[] = [];
  if (run.underline || run.doubleUnderline) decorations.push("underline");
  if (run.strikethrough) decorations.push("line-through");
  const basePointSize = run.fontSize ?? options.defaultPoints ?? 11;
  const scriptScale = run.superscript || run.subscript ? 0.75 : 1;
  const safeFamily = run.fontFamily?.replace(/["'\\;{}()]/g, "").trim();
  return {
    fontWeight: run.bold ? 700 : undefined,
    fontStyle: run.italic ? "italic" : undefined,
    textDecorationLine: decorations.length ? decorations.join(" ") : undefined,
    textDecorationStyle: run.doubleUnderline ? "double" : undefined,
    verticalAlign: run.superscript ? "super" : run.subscript ? "sub" : undefined,
    fontFamily: safeFamily
      ? `"${safeFamily}", Aptos, Calibri, sans-serif`
      : undefined,
    fontSize: options.pointToPx(basePointSize * scriptScale),
    color: run.color,
  };
}
