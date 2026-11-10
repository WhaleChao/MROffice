/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Persisted, per-editor display PREFERENCES for the CodeMirror 6 editor
 * (VSCode-parity toggles): word wrap, render-whitespace, and bracket-pair
 * colorization, plus the editor's font-size zoom (`fontScale`). Each is a value
 * the user changes from the editor chrome and we remember across reloads in
 * `localStorage`, the same way VSCode keeps its "editor.wordWrap" /
 * "editor.fontSize" view state.
 *
 * Framework-free + unit-tested: `CodeMirrorEditor` seeds its toggle `useState`
 * from {@link loadEditorPref} and writes back via {@link saveEditorPref}. All
 * helpers are TOTAL — they never throw, so a private-mode / disabled-storage
 * browser (where `localStorage` access raises) degrades to the supplied default
 * rather than breaking the editor.
 */

import { pinchZoomLevel, wheelDeltaPixels } from "@mrnote/core/web/lib/editor-zoom.ts";

/** The display preferences we persist. Stable string keys = the storage keys. */
export type EditorPrefKey =
  // VSCode `editor.minimap.enabled`: default ON, but a user's explicit choice
  // survives editor/page remounts and automatic large-file mode.
  | "minimap"
  | "wordWrap"
  | "renderWhitespace"
  | "bracketColors"
  | "indentGuides"
  | "bracketPairGuides"
  | "stickyScroll"
  | "inlayHints"
  | "codeLens"
  // VSCode "editor.formatOnSave": when ON, an explicit Ctrl/Cmd+S pretty-prints
  // the buffer before persisting. Default OFF (see CodeMirrorEditor), matching
  // VSCode's own default and the other opt-in view toggles.
  | "formatOnSave"
  // VSCode "files.trimTrailingWhitespace": when ON, an explicit Ctrl/Cmd+S strips
  // trailing whitespace from every line before persisting. Independent of
  // formatOnSave (it matters for languages with no Prettier formatter — Python /
  // Shell / SQL — whose save path never trims otherwise). Default OFF.
  | "trimTrailingOnSave"
  // VSCode "files.insertFinalNewline": when ON, an explicit Ctrl/Cmd+S ensures the
  // file ends with exactly one trailing newline before persisting. Default OFF.
  | "insertFinalNewline"
  // VSCode "files.trimFinalNewlines": when ON, an explicit Ctrl/Cmd+S removes all
  // trailing blank lines (no final newline) before persisting. insertFinalNewline
  // takes priority when both are somehow enabled (matching VSCode). Default OFF.
  | "trimFinalNewlines"
  // Overleaf's "Auto compile": a LaTeX page compiles once typing pauses
  // (lib/code-runner/latex-auto-compile). Default OFF: every compile is a run.
  | "latexAutoCompile";

const STORAGE_PREFIX = "mrnote.code-editor.";

/** The localStorage key a preference is stored under. */
export function editorPrefStorageKey(key: EditorPrefKey): string {
  return `${STORAGE_PREFIX}${key}`;
}

/**
 * Read a persisted boolean preference, falling back to `defaultValue` when it
 * was never set, holds an unrecognised value, or `localStorage` is unavailable.
 * Only the exact strings `"true"` / `"false"` are honoured; anything else (a
 * legacy/corrupt value) yields the default.
 */
export function loadEditorPref(key: EditorPrefKey, defaultValue: boolean): boolean {
  try {
    const raw = globalThis.localStorage?.getItem(editorPrefStorageKey(key));
    if (raw === "true") return true;
    if (raw === "false") return false;
    return defaultValue;
  } catch {
    return defaultValue;
  }
}

/**
 * Persist a boolean preference. Total: a storage failure (quota, private mode)
 * is swallowed so toggling never throws into the React tree.
 */
export function saveEditorPref(key: EditorPrefKey, value: boolean): void {
  try {
    globalThis.localStorage?.setItem(editorPrefStorageKey(key), value ? "true" : "false");
  } catch {
    /* storage unavailable — the in-memory toggle still works this session. */
  }
}

/**
 * Editor font-size zoom (VSCode parity, "editor.fontSize" view state) expressed
 * as a SCALE multiplier of the editor's base font size — 1.0 = 100% = the
 * default size, so a brand-new editor that never zoomed is byte-for-byte the
 * pre-feature surface. Bounded to [{@link FONT_SCALE_MIN}, {@link FONT_SCALE_MAX}]
 * (80%–200%) and stepped by {@link FONT_SCALE_STEP} on each zoom in/out.
 */
export const FONT_SCALE_MIN = 0.8;
export const FONT_SCALE_MAX = 2;
export const FONT_SCALE_STEP = 0.1;
export const FONT_SCALE_DEFAULT = 1;

/** Storage key for the font-scale pref; kept under the same editor namespace. */
const FONT_SCALE_STORAGE_KEY = `${STORAGE_PREFIX}fontScale`;

/**
 * Clamp an arbitrary number to the legal font-scale range, snapping to one
 * decimal place so repeated step arithmetic can't accumulate float drift
 * (0.1 + 0.2 …). A non-finite input (NaN / ±Infinity, e.g. a corrupt stored
 * value) collapses to {@link FONT_SCALE_DEFAULT} rather than propagating.
 */
export function clampFontScale(value: number): number {
  if (!Number.isFinite(value)) return FONT_SCALE_DEFAULT;
  const bounded = Math.min(FONT_SCALE_MAX, Math.max(FONT_SCALE_MIN, value));
  // Snap to 0.01 precision so 0.8999999 reads back as 0.9.
  return Math.round(bounded * 100) / 100;
}

/**
 * Ctrl/Cmd+wheel and trackpad pinch over the code editor: the next font scale
 * from a wheel delta, on the Docs / workbook / deck gesture curve
 * (lib/editor-zoom `pinchZoomLevel`: one mouse notch is one 10% step, growth
 * is multiplicative, a trackpad's small deltas move proportionally less),
 * bounded to the code editor's own range. The result is NOT snapped: the
 * caller keeps it as the live gesture level (snapping each sub-pixel nudge
 * would undo it) and lets {@link clampFontScale} snap what it shows and
 * stores.
 */
export function wheelZoomFontScale(
  scale: number,
  deltaY: number,
  deltaMode = 0,
): number {
  const base = Number.isFinite(scale) ? scale : FONT_SCALE_DEFAULT;
  const next = pinchZoomLevel(base, wheelDeltaPixels(deltaY, deltaMode));
  return Math.min(FONT_SCALE_MAX, Math.max(FONT_SCALE_MIN, next));
}

/**
 * Read the persisted font scale, falling back to {@link FONT_SCALE_DEFAULT} when
 * it was never set, holds a non-numeric value, or `localStorage` is unavailable.
 * The stored value is always re-clamped on load so a hand-edited / legacy
 * out-of-range entry can never push the editor past the bounds.
 */
export function loadFontScale(): number {
  try {
    const raw = globalThis.localStorage?.getItem(FONT_SCALE_STORAGE_KEY);
    if (raw == null || raw === "") return FONT_SCALE_DEFAULT;
    const parsed = Number.parseFloat(raw);
    if (!Number.isFinite(parsed)) return FONT_SCALE_DEFAULT;
    return clampFontScale(parsed);
  } catch {
    return FONT_SCALE_DEFAULT;
  }
}

/**
 * Persist the font scale (clamped first, so an out-of-range caller value is
 * normalised before it lands). Total: a storage failure is swallowed so zooming
 * never throws into the React tree. Returns the clamped value actually stored so
 * the caller can keep its in-memory state in lockstep with what was persisted.
 */
export function saveFontScale(value: number): number {
  const clamped = clampFontScale(value);
  try {
    globalThis.localStorage?.setItem(FONT_SCALE_STORAGE_KEY, String(clamped));
  } catch {
    /* storage unavailable — the in-memory zoom still works this session. */
  }
  return clamped;
}

/**
 * Vertical column rulers (VSCode `editor.rulers`) — the 1-based columns at which
 * the editor draws a faint full-height guide line (e.g. 80 / 120 to mark line-
 * length limits). Persisted as a comma-separated list so a user's chosen columns
 * survive reloads, defaulting to EMPTY (no rulers) so a brand-new editor that
 * never set them is byte-for-byte the pre-feature surface.
 */
const RULER_COLUMNS_STORAGE_KEY = `${STORAGE_PREFIX}rulerColumns`;

/** A ruler column must be a positive integer; anything beyond is unreasonable. */
const RULER_COLUMN_MAX = 1000;

/**
 * Parse a user-entered ruler-columns string (e.g. "80, 120") into a normalized,
 * de-duplicated, ascending list of positive integer columns. Total + pure:
 *  - separators are commas and/or whitespace,
 *  - non-numeric / non-positive / non-finite tokens are dropped,
 *  - each value is floored and clamped to (0, {@link RULER_COLUMN_MAX}],
 *  - duplicates are removed and the result is sorted ascending,
 *  - an empty / all-garbage input yields `[]` (rulers off).
 * Used both to parse the settings input and to sanitize a stored value on load.
 */
export function parseRulerColumns(input: string): number[] {
  const out = new Set<number>();
  for (const token of input.split(/[\s,]+/)) {
    if (token === "") continue;
    const n = Number.parseInt(token, 10);
    if (!Number.isFinite(n) || n <= 0) continue;
    out.add(Math.min(RULER_COLUMN_MAX, Math.floor(n)));
  }
  return Array.from(out).sort((a, b) => a - b);
}

/** Serialize ruler columns back to the stored / displayed comma-separated form. */
export function formatRulerColumns(columns: number[]): string {
  return parseRulerColumns(columns.join(",")).join(", ");
}

/**
 * Read the persisted ruler columns, falling back to `[]` (no rulers) when never
 * set, empty, or `localStorage` is unavailable. The stored value is re-parsed +
 * re-sanitized on load (via {@link parseRulerColumns}) so a hand-edited / legacy
 * entry can never feed bad columns into the editor.
 */
export function loadRulerColumns(): number[] {
  try {
    const raw = globalThis.localStorage?.getItem(RULER_COLUMNS_STORAGE_KEY);
    if (raw == null || raw === "") return [];
    return parseRulerColumns(raw);
  } catch {
    return [];
  }
}

/**
 * Persist ruler columns (sanitized first). Total: a storage failure is swallowed
 * so editing the setting never throws into the React tree. Returns the
 * normalized columns actually stored so the caller keeps its in-memory state in
 * lockstep with what was persisted.
 */
export function saveRulerColumns(columns: number[]): number[] {
  const normalized = parseRulerColumns(columns.join(","));
  try {
    globalThis.localStorage?.setItem(
      RULER_COLUMNS_STORAGE_KEY,
      normalized.join(","),
    );
  } catch {
    /* storage unavailable — the in-memory rulers still work this session. */
  }
  return normalized;
}
