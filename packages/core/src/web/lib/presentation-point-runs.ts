/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** PPT's DrawingML sz is measured in hundredths of a point. Keep this opt-in
 * at the PPT boundary; shared run helpers retain their existing default. */
import * as shared from "./presentation-runs";
export * from "./presentation-runs";
const POINTS = Object.freeze({ fontSizePrecision: "hundredth" as const });
export const normalizeRunStyle = (value: unknown) => shared.normalizeRunStyle(value, POINTS);
export const normalizeRuns = (value: unknown) => shared.normalizeRuns(value, POINTS);
export const mergeRuns = (runs: readonly shared.TextRun[]) => shared.mergeRuns(runs, POINTS);
export const isFlatEquivalentRuns = (runs: readonly shared.TextRun[] | undefined) => shared.isFlatEquivalentRuns(runs, POINTS);
export const runCollectionsEqual = (a: readonly shared.TextRun[] | undefined, b: readonly shared.TextRun[] | undefined) => shared.runCollectionsEqual(a, b, POINTS);
export const paragraphRunSelectionRanges = (paragraphs: readonly shared.ParagraphRunsLike[], start: number, end: number) => shared.paragraphRunSelectionRanges(paragraphs, start, end, POINTS);
export const applyRunFormat = (
  runs: readonly shared.TextRun[],
  start: number,
  end: number,
  patch: Partial<shared.RunStyle> | ((style: shared.RunStyle | undefined) => Partial<shared.RunStyle>),
) => shared.applyRunFormat(runs, start, end, patch, POINTS);
export const applyRunField = (runs: readonly shared.TextRun[], start: number, end: number, field: shared.RunField) =>
  shared.applyRunField(runs, start, end, field, POINTS);
export const applyStylePatch = (base: shared.RunStyle | undefined, patch: Partial<shared.RunStyle>) => shared.applyStylePatch(base, patch, POINTS);
export function runFieldAtSelection<K extends keyof shared.RunStyle>(runs: readonly shared.TextRun[], start: number, end: number, pick: K) {
  return shared.runFieldAtSelection(runs, start, end, pick, POINTS);
}
export const parseRunFontSizeInput = (raw: string) => shared.parseRunFontSizeInput(raw, POINTS);
export const clampRunFontSize = (value: number) => shared.clampRunFontSize(value, POINTS);
export const stepRunFontSize = (current: number | undefined, delta: number) => shared.stepRunFontSize(current, delta, POINTS);

export type TextEditDelta = Readonly<{
  from: number;
  to: number;
  inserted: string;
}>;

export type TextEditSelection = Readonly<{
  start: number;
  end: number;
}>;

/** The single minimal replacement that changes ``previous`` into ``next``. */
export function textEditDelta(
  previous: string,
  next: string,
  selection?: TextEditSelection,
): TextEditDelta {
  if (selection) {
    const rawStart = Math.min(
      previous.length,
      Math.max(0, Math.trunc(selection.start)),
    );
    const rawEnd = Math.min(
      previous.length,
      Math.max(0, Math.trunc(selection.end)),
    );
    const from = Math.min(rawStart, rawEnd);
    const to = Math.max(rawStart, rawEnd);
    const prefix = previous.slice(0, from);
    const suffix = previous.slice(to);
    if (
      next.length >= prefix.length + suffix.length &&
      next.startsWith(prefix) &&
      next.endsWith(suffix)
    ) {
      return {
        from,
        to,
        inserted: next.slice(from, next.length - suffix.length),
      };
    }
  }
  let from = 0;
  const shared = Math.min(previous.length, next.length);
  while (from < shared && previous[from] === next[from]) from += 1;
  let suffix = 0;
  while (
    suffix < previous.length - from &&
    suffix < next.length - from &&
    previous[previous.length - 1 - suffix] === next[next.length - 1 - suffix]
  ) {
    suffix += 1;
  }
  return {
    from,
    to: previous.length - suffix,
    inserted: next.slice(from, next.length - suffix),
  };
}

/** Copy the run fragments intersecting ``[from, to)``. */
export function textRunSlice(
  runs: readonly shared.TextRun[],
  from: number,
  to: number,
): shared.TextRun[] {
  if (from >= to) return [];
  const pieces: shared.TextRun[] = [];
  let cursor = 0;
  for (const run of runs) {
    const end = cursor + run.text.length;
    const localFrom = Math.max(from, cursor);
    const localTo = Math.min(to, end);
    if (localFrom < localTo) {
      const text = run.text.slice(localFrom - cursor, localTo - cursor);
      const style = normalizeRunStyle(run.style);
      const piece: shared.TextRun = style ? { text, style: { ...style } } : { text };
      // A field or an equation taken whole is still that field or equation;
      // a cut through it is the plain text PowerPoint leaves of a broken one.
      if (text.length === run.text.length && shared.runIsAtomic(run)) {
        if (run.field) piece.field = run.field;
        if (run.math) piece.math = run.math;
      }
      pieces.push(piece);
    }
    cursor = end;
    if (cursor >= to) break;
  }
  return pieces;
}

/** Office-style typing affinity: replaced glyph first, otherwise glyph left. */
export function textInsertionRunStyle(
  runs: readonly shared.TextRun[],
  delta: TextEditDelta,
): shared.RunStyle | undefined {
  const length = shared.textFromRuns(runs).length;
  if (length === 0) return undefined;
  const offset =
    delta.from < delta.to
      ? Math.min(delta.from, length - 1)
      : delta.from > 0
        ? Math.min(delta.from - 1, length - 1)
        : 0;
  let cursor = 0;
  for (const run of runs) {
    const end = cursor + run.text.length;
    if (offset < end) {
      const style = normalizeRunStyle(run.style);
      return style ? { ...style } : undefined;
    }
    cursor = end;
  }
  return undefined;
}

function styleWithoutLink(style: shared.RunStyle): shared.RunStyle | undefined {
  const next: shared.RunStyle = { ...style };
  delete next.href;
  delete next.linkTooltip;
  delete next.linkColor;
  return Object.keys(next).length ? next : undefined;
}

/**
 * Reconcile a plain-text DOM edit into existing authenticated runs.
 *
 * The old implementation deleted the complete run array after every keypress,
 * so selecting an imported Office text box instantly removed mixed bold/font
 * boundaries.  A contenteditable input is one contiguous replacement: keep the
 * untouched prefix/suffix runs and give inserted text the caret's surrounding
 * run style.  Mismatched/stale runs still fail closed to flat text.
 */
export function reconcileTextRunsAfterEdit(
  previousText: string,
  nextText: string,
  rawRuns: readonly shared.TextRun[] | undefined,
  editDelta?: TextEditDelta,
  options?: { newParagraphEndsLink?: boolean },
): shared.TextRun[] | undefined {
  const runs = normalizeRuns(rawRuns);
  if (!runs || shared.textFromRuns(runs) !== previousText) return undefined;
  if (previousText === nextText) return runs;
  const delta = editDelta ?? textEditDelta(previousText, nextText);
  const insertedStyle = textInsertionRunStyle(runs, delta);
  // A new paragraph takes its formatting from the paragraph end
  // (`a:endParaRPr`), which holds no hyperlink: text typed after Enter at the
  // end of a link is not part of it, and the save writes it unlinked
  
  const breakAt = options?.newParagraphEndsLink ? delta.inserted.indexOf("\n") : -1;
  const unlinkedStyle =
    insertedStyle && breakAt >= 0 && insertedStyle.href !== undefined
      ? styleWithoutLink(insertedStyle)
      : insertedStyle;
  const piece = (text: string, style: shared.RunStyle | undefined) =>
    text ? [style ? { text, style: { ...style } } : { text }] : [];
  const inserted = delta.inserted
    ? breakAt >= 0 && unlinkedStyle !== insertedStyle
      ? [
          ...piece(delta.inserted.slice(0, breakAt), insertedStyle),
          ...piece(delta.inserted.slice(breakAt), unlinkedStyle),
        ]
      : [
          insertedStyle
            ? { text: delta.inserted, style: insertedStyle }
            : { text: delta.inserted },
        ]
    : [];
  return normalizeRuns(
    mergeRuns([
      ...textRunSlice(runs, 0, delta.from),
      ...inserted,
      ...textRunSlice(runs, delta.to, previousText.length),
    ]),
  );
}
