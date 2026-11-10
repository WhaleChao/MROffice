/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure model + ops for MANUALLY hidden rows / columns (Excel/WPS "Hide" and
 * "Unhide" on a row or column header). This is DISTINCT from outline grouping
 * (outline-groups.ts): grouping hides a contiguous body behind a [+]/[-] toggle
 * and is driven by ranges, whereas a manual hide is an arbitrary SET of absolute
 * line indices the user has collapsed one (or a span) at a time. The two hide
 * mechanisms are independent and UNION at the render layer — a line is hidden if
 * EITHER a collapsed group covers it OR it is in the manual-hidden set.
 *
 * The model is a plain SET of 0-based line indices per axis, persisted on the
 * sheet as `manuallyHiddenRows` / `manuallyHiddenCols` number arrays so it
 * round-trips trivially. Everything here is pure and data-only (no DOM, no
 * formula engine) so it can be unit-tested in isolation and reused by the grid:
 * the grid normalizes the persisted arrays on load, derives the hidden sets,
 * unions them with the outline-group hidden sets when filtering rendered lines,
 * and calls hideLines / unhideLines from the header context menu.
 *
 * UNHIDE semantics match Excel: selecting a span that CROSSES a hidden line
 * reveals every hidden line strictly inside the span (so selecting rows 4–8 with
 * rows 5–7 hidden unhides 5,6,7). Selecting two adjacent visible lines that
 * bracket a run of hidden lines (Excel's "select the rows on both sides, then
 * Unhide") reveals the bracketed run too — handled by treating the inclusive
 * selection span [start, end] as the reveal window.
 */

/** Which axis a manual hide lives on. */
export type ManualHideAxis = "row" | "col";

/**
 * Coerce a persisted hidden-line array into a normalized SET of in-bounds 0-based
 * indices: drops non-integer / negative entries, clamps out-of-range entries out
 * (a line index >= lineCount cannot be hidden — there is no such line), and
 * de-duplicates. `lineCount` is the axis dimension (rowCount / columnCount);
 * `limit` mirrors it so a malicious file cannot build an unbounded set (a sheet
 * can hide at most one entry per addressable line). Returns a NEW set; empty when
 * nothing valid survives so the no-hidden render path is byte-stable.
 */
export function normalizeManualHidden(
  raw: unknown,
  lineCount: number,
): Set<number> {
  const out = new Set<number>();
  if (!Array.isArray(raw)) return out;
  const max = Math.max(0, Math.trunc(lineCount) - 1);
  // One entry per addressable line bounds the set; slice the input first so a
  // pathological array never drives an unbounded loop.
  for (const entry of raw.slice(0, max + 1)) {
    const idx = Number(entry);
    if (!Number.isInteger(idx) || idx < 0 || idx > max) continue;
    out.add(idx);
  }
  return out;
}

/**
 * Hide every line in the inclusive span [start, end] (a header right-click →
 * Hide, where the selection may be one line or a span). Returns a NEW set with
 * the span added to `hidden`; entries already hidden are kept. A malformed /
 * inverted span is normalized (min/max), and only non-negative indices are
 * added.
 */
export function hideLines(
  hidden: ReadonlySet<number>,
  start: number,
  end: number,
): Set<number> {
  const next = new Set(hidden);
  const a = Math.trunc(Number(start));
  const b = Math.trunc(Number(end));
  if (!Number.isFinite(a) || !Number.isFinite(b)) return next;
  const lo = Math.max(0, Math.min(a, b));
  const hi = Math.max(a, b);
  for (let i = lo; i <= hi; i += 1) next.add(i);
  return next;
}

/**
 * Unhide every hidden line inside the inclusive span [start, end] (a header
 * right-click → Unhide over a selection that crosses or brackets hidden lines).
 * Returns a NEW set with those entries removed; lines outside the span stay
 * hidden. A malformed / inverted span is normalized.
 */
export function unhideLines(
  hidden: ReadonlySet<number>,
  start: number,
  end: number,
): Set<number> {
  const next = new Set(hidden);
  const a = Math.trunc(Number(start));
  const b = Math.trunc(Number(end));
  if (!Number.isFinite(a) || !Number.isFinite(b)) return next;
  const lo = Math.min(a, b);
  const hi = Math.max(a, b);
  for (const idx of hidden) {
    if (idx >= lo && idx <= hi) next.delete(idx);
  }
  return next;
}

/**
 * Hide or show exactly these lines, however scattered (0 typed in Column
 * Width for a selection hides it; a width typed for a selection that holds
 * hidden columns shows them again, as in Excel). Returns a NEW set.
 */
export function setLinesHidden(
  hidden: ReadonlySet<number>,
  lines: Iterable<number>,
  hide: boolean,
): Set<number> {
  const next = new Set(hidden);
  for (const line of lines) {
    if (!Number.isSafeInteger(line) || line < 0) continue;
    if (hide) next.add(line);
    else next.delete(line);
  }
  return next;
}

/** True when any line in the inclusive span [start, end] is hidden. */
export function spanHasHidden(
  hidden: ReadonlySet<number>,
  start: number,
  end: number,
): boolean {
  const lo = Math.min(start, end);
  const hi = Math.max(start, end);
  for (const idx of hidden) {
    if (idx >= lo && idx <= hi) return true;
  }
  return false;
}

/**
 * Serialize a hidden-line set back to a SORTED number array for persistence
 * (stable ordering keeps the saved model byte-stable across re-saves), or
 * undefined when empty so the field is omitted from the sheet entirely (the
 * no-hidden case writes nothing — matching how outlineGroups / conditionalFormats
 * are dropped when empty).
 */
export function serializeManualHidden(
  hidden: ReadonlySet<number>,
): number[] | undefined {
  if (hidden.size === 0) return undefined;
  return Array.from(hidden).sort((p, q) => p - q);
}
