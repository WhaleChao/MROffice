/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Persisted per-document CodeMirror view state.
 *
 * This is intentionally separate from editor-prefs.ts: preferences are global
 * authoring choices (word wrap, minimap, font scale), while this snapshot is
 * scoped to one code page and restores where the author was working after the
 * page/window is closed or the browser reloads.
 *
 * The payload is bounded and treated as untrusted storage:
 * - a snapshot only restores against the exact document length it was saved for;
 * - offsets and scroll coordinates are finite, non-negative, and clamped;
 * - at most 2,048 selections and 2,048 folded regions are persisted, keeping
 *   pathological sessions well below localStorage's usual quota and cheap to
 *   parse;
 * - malformed, stale, unavailable, or quota-exhausted storage is a safe no-op.
 */

export interface CodeEditorViewRange {
  anchor: number;
  head: number;
}

export interface CodeEditorFoldRange {
  from: number;
  to: number;
}

export interface CodeEditorViewStateSnapshot {
  version: 1;
  documentLength: number;
  ranges: CodeEditorViewRange[];
  mainIndex: number;
  scrollTop: number;
  scrollLeft: number;
  /** Optional for backward compatibility with snapshots saved before folds. */
  folds?: CodeEditorFoldRange[];
}

export const MAX_CODE_EDITOR_VIEW_STATE_RANGES = 2_048;
export const MAX_CODE_EDITOR_VIEW_STATE_FOLDS = 2_048;

const VIEW_STATE_STORAGE_PREFIX = "mrnote.code-editor.view.v1:";
const MAX_STORED_SCROLL_OFFSET = 1_000_000_000;
const MAX_STORED_PAYLOAD_BYTES = 256_000;

/**
 * Capture the live editor's folded ranges without scanning a pathological
 * DecorationSet past the same bound enforced for persisted snapshots.
 *
 * The callback mirrors CodeMirror RangeSet's `between` contract: returning
 * `false` stops iteration; returning `void` continues.
 */
export function collectCodeEditorFoldRanges(
  visitRanges: (
    visit: (from: number, to: number) => false | void,
  ) => void,
): CodeEditorFoldRange[] {
  const folds: CodeEditorFoldRange[] = [];
  visitRanges((from, to) => {
    folds.push({ from, to });
    if (folds.length >= MAX_CODE_EDITOR_VIEW_STATE_FOLDS) return false;
  });
  return folds;
}

/** Stable, per-code-page localStorage key. */
export function codeEditorViewStateStorageKey(scope: string): string {
  return `${VIEW_STATE_STORAGE_PREFIX}${encodeURIComponent(scope)}`;
}

function nonNegativeInteger(value: unknown, max: number): number | null {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    !Number.isInteger(value) ||
    value < 0
  ) {
    return null;
  }
  return Math.min(value, max);
}

function boundedInteger(value: unknown, max: number): number | null {
  const normalized = nonNegativeInteger(value, max);
  return normalized === value ? normalized : null;
}

function scrollOffset(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
    return 0;
  }
  return Math.min(value, MAX_STORED_SCROLL_OFFSET);
}

function normalizeFoldRanges(
  value: unknown,
  documentLength: number,
): CodeEditorFoldRange[] | undefined {
  if (value === undefined) return undefined;
  if (!Array.isArray(value)) return [];

  const folds: CodeEditorFoldRange[] = [];
  for (
    let index = 0;
    index < value.length &&
    folds.length < MAX_CODE_EDITOR_VIEW_STATE_FOLDS;
    index += 1
  ) {
    const candidate = value[index];
    if (!candidate || typeof candidate !== "object") continue;
    const record = candidate as Record<string, unknown>;
    const from = boundedInteger(record.from, documentLength);
    const to = boundedInteger(record.to, documentLength);
    if (from === null || to === null || from >= to) continue;
    folds.push({ from, to });
  }

  folds.sort((left, right) => left.from - right.from || right.to - left.to);
  return folds.filter(
    (fold, index) =>
      index === 0 ||
      fold.from !== folds[index - 1].from ||
      fold.to !== folds[index - 1].to,
  );
}

/**
 * Validate and bound a candidate snapshot. The returned ranges are ordered and
 * non-overlapping, which makes them safe to pass to EditorSelection.create.
 */
export function normalizeCodeEditorViewState(
  candidate: unknown,
  expectedDocumentLength?: number,
): CodeEditorViewStateSnapshot | null {
  if (!candidate || typeof candidate !== "object") return null;
  const record = candidate as Record<string, unknown>;
  if (record.version !== 1) return null;

  const documentLength = nonNegativeInteger(
    record.documentLength,
    Number.MAX_SAFE_INTEGER,
  );
  if (documentLength === null) return null;
  if (
    expectedDocumentLength !== undefined &&
    documentLength !== expectedDocumentLength
  ) {
    return null;
  }
  if (!Array.isArray(record.ranges) || record.ranges.length === 0) return null;

  const sourceMainIndex = nonNegativeInteger(
    record.mainIndex,
    Math.max(0, record.ranges.length - 1),
  );
  const ranges: CodeEditorViewRange[] = [];
  let mainIndex = 0;
  let previousTo = -1;

  for (
    let sourceIndex = 0;
    sourceIndex < record.ranges.length &&
    ranges.length < MAX_CODE_EDITOR_VIEW_STATE_RANGES;
    sourceIndex += 1
  ) {
    const sourceRange = record.ranges[sourceIndex];
    if (!sourceRange || typeof sourceRange !== "object") continue;
    const rangeRecord = sourceRange as Record<string, unknown>;
    const anchor = nonNegativeInteger(rangeRecord.anchor, documentLength);
    const head = nonNegativeInteger(rangeRecord.head, documentLength);
    if (anchor === null || head === null) continue;
    const from = Math.min(anchor, head);
    const to = Math.max(anchor, head);
    if (from < previousTo) continue;
    if (sourceIndex === sourceMainIndex) mainIndex = ranges.length;
    ranges.push({ anchor, head });
    previousTo = to;
  }

  if (ranges.length === 0) return null;
  if (mainIndex >= ranges.length) mainIndex = 0;
  const folds = normalizeFoldRanges(record.folds, documentLength);

  return {
    version: 1,
    documentLength,
    ranges,
    mainIndex,
    scrollTop: scrollOffset(record.scrollTop),
    scrollLeft: scrollOffset(record.scrollLeft),
    ...(folds === undefined ? {} : { folds }),
  };
}

/** Load a current, schema-valid snapshot for one code page. */
export function loadCodeEditorViewState(
  scope: string | undefined,
  documentLength: number,
): CodeEditorViewStateSnapshot | null {
  if (!scope) return null;
  try {
    const raw = globalThis.localStorage?.getItem(
      codeEditorViewStateStorageKey(scope),
    );
    if (!raw || raw.length > MAX_STORED_PAYLOAD_BYTES) return null;
    return normalizeCodeEditorViewState(JSON.parse(raw), documentLength);
  } catch {
    return null;
  }
}

/**
 * Persist one bounded snapshot. Returns the normalized value actually written,
 * or null when the candidate is invalid/storage is unavailable.
 */
export function saveCodeEditorViewState(
  scope: string | undefined,
  snapshot: CodeEditorViewStateSnapshot,
): CodeEditorViewStateSnapshot | null {
  if (!scope) return null;
  const normalized = normalizeCodeEditorViewState(snapshot);
  if (!normalized) return null;
  try {
    globalThis.localStorage?.setItem(
      codeEditorViewStateStorageKey(scope),
      JSON.stringify(normalized),
    );
    return normalized;
  } catch {
    return null;
  }
}
