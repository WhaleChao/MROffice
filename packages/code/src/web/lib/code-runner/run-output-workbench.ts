/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { stripAnsi } from "@mrnote/core/web/lib/code-runner/ansi.ts";
import {
  streamForLogIndex,
  type LogStream,
} from "@mrnote/core/web/lib/code-runner/cell-output-format.ts";
import type { RunResult } from "@mrnote/core/web/lib/code-runner/js-runner.ts";

/** Keep clipboard/download payloads bounded even for a legacy runaway runner. */
export const MAX_RUN_OUTPUT_TEXT_CHARS = 1_000_000;

const OUTPUT_WRAP_STORAGE_KEY = "mrnote.code-output.wordWrap";
const OUTPUT_PANEL_VISIBLE_STORAGE_KEY =
  "mrnote.code-output.panelVisible";
const OUTPUT_PANEL_PERCENT_STORAGE_KEY =
  "mrnote.code-output.panelHeightPercent";
const TRUNCATION_NOTICE = "\n… output truncated …";
const FILENAME_INVALID =
  /[\\/:*?"<>|\x00-\x1f\x7f-\x9f\u200e\u200f\u202a-\u202e\u2066-\u2069]/g;

export interface RunOutputTextSnapshot {
  /** ANSI/terminal controls removed; safe for clipboard and text export. */
  readonly text: string;
  /** True when the caller-visible payload hit its character budget. */
  readonly truncated: boolean;
}

export type RunOutputStreamFilter = "all" | LogStream;

export interface RunOutputLogLine {
  readonly index: number;
  readonly text: string;
  readonly stream: LogStream;
}

export interface RunOutputLogWindow {
  /** Offset inside the filtered sequence, not the unfiltered logs array. */
  readonly start: number;
  readonly end: number;
  readonly total: number;
  readonly hiddenBefore: number;
  readonly hiddenAfter: number;
  readonly stdoutCount: number;
  readonly stderrCount: number;
  readonly lines: readonly RunOutputLogLine[];
}

/** Hard DOM ceiling for standalone run-output log rows. */
export const RUN_OUTPUT_LOG_PAGE_SIZE = 500;
/** Default and hard bounds for the resizable bottom output panel. */
export const RUN_OUTPUT_PANEL_DEFAULT_PERCENT = 40;
export const RUN_OUTPUT_PANEL_MIN_PERCENT = 18;
export const RUN_OUTPUT_PANEL_MAX_PERCENT = 72;
export const RUN_OUTPUT_PANEL_KEYBOARD_STEP = 5;
/** Keep a pasted/automated query from becoming its own memory-pressure source. */
export const MAX_RUN_OUTPUT_SEARCH_QUERY_CHARS = 256;
/** Search the newest bounded tail when a legacy producer bypasses runner caps. */
export const MAX_RUN_OUTPUT_SEARCH_LINES = 100_000;
/** Visible-text budget across one search pass (ANSI bytes do not count). */
export const MAX_RUN_OUTPUT_SEARCH_TEXT_CHARS = 2_000_000;
/** Result coordinates stay bounded even when every output line is identical. */
export const MAX_RUN_OUTPUT_SEARCH_MATCHES = 5_000;
const MAX_RUN_OUTPUT_SEARCH_LINE_CHARS = 64_000;

export interface RunOutputSearchMatch {
  /** Offset in the unfiltered `logs` array; stable while streaming appends. */
  readonly index: number;
  /** Offset inside the selected stdout/stderr sequence. */
  readonly filteredIndex: number;
}

export interface RunOutputSearchResult {
  readonly query: string;
  readonly matches: readonly RunOutputSearchMatch[];
  readonly inspected: number;
  /** Earlier rows, over-budget text, query text, or matches were omitted. */
  readonly truncated: boolean;
}

export interface RunOutputPanelShortcutGesture {
  readonly key: string;
  readonly metaKey?: boolean;
  readonly ctrlKey?: boolean;
  readonly shiftKey?: boolean;
  readonly altKey?: boolean;
  readonly repeat?: boolean;
  readonly isComposing?: boolean;
  readonly keyCode?: number;
}

export type RunOutputPanelShortcutAction = "toggle" | "consume";

/**
 * Resolve VS Code's Ctrl/Cmd+J panel gesture without touching editor state.
 * Held-key repeats are consumed so one press cannot flicker the panel.
 */
export function resolveRunOutputPanelShortcut(
  gesture: RunOutputPanelShortcutGesture,
): RunOutputPanelShortcutAction | null {
  if (gesture.isComposing || gesture.keyCode === 229) return null;
  if (
    !(gesture.metaKey || gesture.ctrlKey) ||
    gesture.shiftKey ||
    gesture.altKey ||
    gesture.key.toLowerCase() !== "j"
  ) {
    return null;
  }
  return gesture.repeat ? "consume" : "toggle";
}

/**
 * Select one stable page of stream rows without ever growing the rendered DOM.
 * `requestedStart=null` means follow the newest page. A numeric start anchors
 * an older page while new streaming rows arrive.
 */
export function buildRunOutputLogWindow(
  output: Pick<RunResult, "logs" | "logStreams">,
  filter: RunOutputStreamFilter,
  requestedStart: number | null,
  requestedPageSize = RUN_OUTPUT_LOG_PAGE_SIZE,
): RunOutputLogWindow {
  const pageSize = Math.min(
    RUN_OUTPUT_LOG_PAGE_SIZE,
    Math.max(
      1,
      Number.isFinite(requestedPageSize)
        ? Math.floor(requestedPageSize)
        : RUN_OUTPUT_LOG_PAGE_SIZE,
    ),
  );
  const matchingIndices: number[] = [];
  let stdoutCount = 0;
  let stderrCount = 0;
  for (let index = 0; index < output.logs.length; index += 1) {
    const stream = streamForLogIndex(output.logStreams, index);
    if (stream === "stderr") stderrCount += 1;
    else stdoutCount += 1;
    if (filter === "all" || filter === stream) matchingIndices.push(index);
  }

  const total = matchingIndices.length;
  const latestStart = Math.max(0, total - pageSize);
  const start =
    requestedStart === null || !Number.isFinite(requestedStart)
      ? latestStart
      : Math.min(latestStart, Math.max(0, Math.floor(requestedStart)));
  const end = Math.min(total, start + pageSize);
  const lines = matchingIndices.slice(start, end).map((index) => ({
    index,
    text: output.logs[index],
    stream: streamForLogIndex(output.logStreams, index),
  }));
  return {
    start,
    end,
    total,
    hiddenBefore: start,
    hiddenAfter: total - end,
    stdoutCount,
    stderrCount,
    lines,
  };
}

/**
 * Search terminal-visible log text without materializing an unbounded plain
 * transcript. Matches retain both raw and filtered coordinates so the UI can
 * jump directly to a fixed-size page even with stdout/stderr filtering active.
 *
 * The newest tail wins if a legacy producer exceeds the normal runner limit.
 * That keeps live debugging useful while making the incomplete count explicit.
 */
export function searchRunOutputLogs(
  output: Pick<RunResult, "logs" | "logStreams">,
  filter: RunOutputStreamFilter,
  rawQuery: string,
): RunOutputSearchResult {
  const query = rawQuery.slice(0, MAX_RUN_OUTPUT_SEARCH_QUERY_CHARS);
  let truncated = rawQuery.length > query.length;
  if (!query) {
    return { query, matches: [], inspected: 0, truncated };
  }

  const normalizedQuery = query.toLowerCase();
  const firstInspectedIndex = Math.max(
    0,
    output.logs.length - MAX_RUN_OUTPUT_SEARCH_LINES,
  );
  truncated ||= firstInspectedIndex > 0;

  // Count only stream metadata before the bounded text tail. This preserves a
  // match's exact filtered offset without touching or allocating old log text.
  let filteredIndex = 0;
  for (let index = 0; index < firstInspectedIndex; index += 1) {
    const stream = streamForLogIndex(output.logStreams, index);
    if (filter === "all" || filter === stream) filteredIndex += 1;
  }

  const matches: RunOutputSearchMatch[] = [];
  let inspected = 0;
  let visibleChars = 0;
  for (
    let index = firstInspectedIndex;
    index < output.logs.length;
    index += 1
  ) {
    const stream = streamForLogIndex(output.logStreams, index);
    if (filter !== "all" && filter !== stream) continue;

    const currentFilteredIndex = filteredIndex;
    filteredIndex += 1;
    if (inspected >= MAX_RUN_OUTPUT_SEARCH_LINES) {
      truncated = true;
      break;
    }
    const remainingChars = MAX_RUN_OUTPUT_SEARCH_TEXT_CHARS - visibleChars;
    if (remainingChars <= 0) {
      truncated = true;
      break;
    }

    inspected += 1;
    const visible = visibleTextWithin(
      output.logs[index] ?? "",
      Math.min(MAX_RUN_OUTPUT_SEARCH_LINE_CHARS, remainingChars),
    );
    visibleChars += visible.text.length;
    if (visible.text.toLowerCase().includes(normalizedQuery)) {
      matches.push({ index, filteredIndex: currentFilteredIndex });
      if (matches.length >= MAX_RUN_OUTPUT_SEARCH_MATCHES) {
        truncated = true;
        break;
      }
    }
    if (visible.truncated) {
      truncated = true;
      break;
    }
  }

  return { query, matches, inspected, truncated };
}

/** Center a filtered match inside the same bounded window used by the DOM. */
export function runOutputLogStartForMatch(
  total: number,
  filteredIndex: number,
  requestedPageSize = RUN_OUTPUT_LOG_PAGE_SIZE,
): number {
  const pageSize = Math.min(
    RUN_OUTPUT_LOG_PAGE_SIZE,
    Math.max(
      1,
      Number.isFinite(requestedPageSize)
        ? Math.floor(requestedPageSize)
        : RUN_OUTPUT_LOG_PAGE_SIZE,
    ),
  );
  const safeTotal = Number.isFinite(total) ? Math.max(0, Math.floor(total)) : 0;
  const latestStart = Math.max(0, safeTotal - pageSize);
  const safeIndex = Number.isFinite(filteredIndex)
    ? Math.max(0, Math.floor(filteredIndex))
    : 0;
  return Math.min(
    latestStart,
    Math.max(0, safeIndex - Math.floor(pageSize / 2)),
  );
}

/** O(1) visibility check used by output toolbar controls during streaming. */
export function hasRunOutputContent(
  output: RunResult | null | undefined,
): boolean {
  if (output?.notebookOutputs !== undefined) {
    return output.notebookOutputs.length > 0;
  }
  return Boolean(
    output &&
      (output.error ||
        output.logs.length > 0 ||
        output.result !== undefined ||
        output.html ||
        output.rich ||
        (output.images?.length ?? 0) > 0),
  );
}

/**
 * Whether Copy Output has a faithful plain-text projection. Imported ordered
 * notebook outputs remain downloadable as complete JSON, but copying them as
 * one flattened transcript would discard MIME bundles and stream ordering.
 */
export function hasRunOutputTextProjection(
  output: RunResult | null | undefined,
): boolean {
  return Boolean(
    output &&
      output.notebookOutputs === undefined &&
      hasRunOutputContent(output),
  );
}

/** Whether the workbench's wrap preference affects a rendered text row. */
export function hasRunOutputWrapTarget(
  output: RunResult | null | undefined,
): boolean {
  return Boolean(
    output &&
      output.notebookOutputs === undefined &&
      (output.logs.length > 0 ||
        (output.result !== undefined && output.rich === undefined)),
  );
}

/** Number shown in the output header; counts rich items without scanning rows. */
export function runOutputItemCount(
  output: RunResult | null | undefined,
): number {
  if (!output) return 0;
  if (output.notebookOutputs !== undefined) {
    return output.notebookOutputs.length;
  }
  return (
    output.logs.length +
    (output.error ? 1 : 0) +
    (output.html ? 1 : 0) +
    (output.rich ? 1 : output.result !== undefined ? 1 : 0) +
    (output.images?.length ?? 0)
  );
}

function normalizeBudget(maxChars: number): number {
  if (!Number.isFinite(maxChars)) return MAX_RUN_OUTPUT_TEXT_CHARS;
  return Math.max(0, Math.floor(maxChars));
}

/**
 * Resolve terminal overwrite semantics, strip ANSI controls, and cap the scan.
 * The raw scan allowance is larger than the visible allowance so ordinary SGR
 * colour sequences do not consume the user's text budget. A hostile escape-only
 * payload still cannot make clipboard/export parse an unbounded string.
 */
function visibleTextWithin(
  raw: string,
  visibleBudget: number,
): RunOutputTextSnapshot {
  if (visibleBudget <= 0) {
    return { text: "", truncated: raw.length > 0 };
  }
  const frameStart = raw.lastIndexOf("\r") + 1;
  const frameLength = raw.length - frameStart;
  const rawScanBudget = Math.min(
    frameLength,
    Math.max(visibleBudget * 4, visibleBudget + 4_096),
  );
  const scanned = raw.slice(frameStart, frameStart + rawScanBudget);
  const visible = stripAnsi(scanned);
  const text = visible.slice(0, visibleBudget);
  return {
    text,
    truncated:
      rawScanBudget < frameLength || visible.length > visibleBudget,
  };
}

function safeTableCell(value: unknown): string {
  if (value === null) return "null";
  if (value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "object") {
    try {
      return JSON.stringify(value);
    } catch {
      // Fall through to a guarded string conversion.
    }
  }
  try {
    return String(value);
  } catch {
    return "";
  }
}

/**
 * Build the plain-text representation behind Copy Output / Download Output.
 * It follows the visible panel order (error, streams, final result), removes
 * ANSI/terminal controls, and preserves the newest terminal frame after `\r`.
 */
export function buildRunOutputText(
  output: RunResult,
  maxChars = MAX_RUN_OUTPUT_TEXT_CHARS,
): RunOutputTextSnapshot {
  const budget = normalizeBudget(maxChars);
  if (budget === 0) {
    const hasText =
      Boolean(output.error) ||
      output.logs.some((line) => line.length > 0) ||
      output.result !== undefined;
    return { text: "", truncated: hasText };
  }

  const parts: string[] = [];
  let used = 0;
  let truncated = false;

  const append = (raw: string) => {
    if (truncated) return;
    const separator = parts.length > 0 ? "\n" : "";
    if (used + separator.length >= budget) {
      truncated = raw.length > 0 || separator.length > 0;
      return;
    }
    const remaining = budget - used - separator.length;
    const visible = visibleTextWithin(raw, remaining);
    if (separator) {
      parts.push(separator);
      used += separator.length;
    }
    if (visible.text) {
      parts.push(visible.text);
      used += visible.text.length;
    }
    truncated = visible.truncated;
  };

  if (output.error) append(output.error);
  for (let index = 0; index < output.logs.length; index += 1) {
    if (truncated) break;
    append(output.logs[index]);
  }
  if (!truncated && output.html && output.logs.length === 0) {
    append("[HTML output]");
  }
  if (!truncated && output.images) {
    for (let index = 0; index < output.images.length; index += 1) {
      if (truncated) break;
      append(`[Image output: ${output.imageMimes?.[index] ?? "image/png"}]`);
    }
  }
  if (!truncated && output.rich?.kind === "table") {
    append(output.rich.columns.map(safeTableCell).join("\t"));
    for (const row of output.rich.rows) {
      if (truncated) break;
      append(row.map(safeTableCell).join("\t"));
    }
  } else if (!truncated && output.rich?.kind === "image") {
    append(`[Image output: ${output.rich.mime}]`);
  } else if (!truncated && output.rich?.kind === "html") {
    append("[HTML output]");
  } else if (!truncated && output.result !== undefined) {
    append(`⇒ ${output.result}`);
  }

  let text = parts.join("");
  if (truncated) {
    const notice =
      budget >= TRUNCATION_NOTICE.length
        ? TRUNCATION_NOTICE
        : TRUNCATION_NOTICE.slice(0, budget);
    text = `${text.slice(0, Math.max(0, budget - notice.length))}${notice}`;
  }
  return { text, truncated };
}

/** Safe, stable filename for an exported standalone-code output transcript. */
export function runOutputFilename(title: string | null | undefined): string {
  let safe = (title ?? "").replace(FILENAME_INVALID, "-").trim();
  if (safe === "." || safe === "..") safe = "";
  return `${safe || "code"}-output.txt`;
}

/** Complete JSON artifact name for output that cannot be represented as text. */
export function runOutputArtifactFilename(
  title: string | null | undefined,
): string {
  return runOutputFilename(title).replace(/\.txt$/, ".json");
}

/**
 * A compiled document's name, for its window and its download: the page's
 * name with the document's extension ("paper.tex" -> "paper.pdf"), or the
 * name it was compiled under ("main.pdf") for a page without a usable name.
 */
export function compiledDocumentFilename(
  title: string | null | undefined,
  compiledName: string,
): string {
  const base = (title ?? "")
    .replace(FILENAME_INVALID, "-")
    .trim()
    .replace(/\.[A-Za-z0-9]{1,8}$/, "");
  const dot = compiledName.lastIndexOf(".");
  if (!base || /^\.+$/.test(base) || dot <= 0) return compiledName;
  return `${base}${compiledName.slice(dot)}`;
}

/** Whether transcript export would discard structured output bytes. */
export function runOutputNeedsArtifact(output: RunResult): boolean {
  return Boolean(
    output.notebookOutputs !== undefined ||
      output.html ||
      output.rich ||
      (output.images?.length ?? 0) > 0,
  );
}

function loadBooleanPreference(
  storageKey: string,
  defaultValue: boolean,
): boolean {
  try {
    const raw = globalThis.localStorage?.getItem(storageKey);
    if (raw === "true") return true;
    if (raw === "false") return false;
    return defaultValue;
  } catch {
    return defaultValue;
  }
}

function saveBooleanPreference(storageKey: string, value: boolean): void {
  try {
    globalThis.localStorage?.setItem(
      storageKey,
      value ? "true" : "false",
    );
  } catch {
    /* The in-memory preference still works for this window. */
  }
}

/** Output wrapping defaults ON to preserve the pre-workbench rendering. */
export function loadRunOutputWrap(defaultValue = true): boolean {
  return loadBooleanPreference(OUTPUT_WRAP_STORAGE_KEY, defaultValue);
}

/** Persist output wrapping without letting disabled/private storage break UI. */
export function saveRunOutputWrap(value: boolean): void {
  saveBooleanPreference(OUTPUT_WRAP_STORAGE_KEY, value);
}

/** A restored result follows the user's last explicit panel visibility. */
export function loadRunOutputPanelVisible(defaultValue = true): boolean {
  return loadBooleanPreference(
    OUTPUT_PANEL_VISIBLE_STORAGE_KEY,
    defaultValue,
  );
}

/** Persist only presentation state; run output remains in the page snapshot. */
export function saveRunOutputPanelVisible(value: boolean): void {
  saveBooleanPreference(OUTPUT_PANEL_VISIBLE_STORAGE_KEY, value);
}

/** Clamp panel geometry from storage, keyboard input, or a pointer drag. */
export function normalizeRunOutputPanelPercent(
  value: unknown,
  fallback = RUN_OUTPUT_PANEL_DEFAULT_PERCENT,
): number {
  const safeFallback =
    Number.isFinite(fallback) && typeof fallback === "number"
      ? Math.min(
          RUN_OUTPUT_PANEL_MAX_PERCENT,
          Math.max(RUN_OUTPUT_PANEL_MIN_PERCENT, fallback),
        )
      : RUN_OUTPUT_PANEL_DEFAULT_PERCENT;
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return safeFallback;
  }
  const clamped = Math.min(
    RUN_OUTPUT_PANEL_MAX_PERCENT,
    Math.max(RUN_OUTPUT_PANEL_MIN_PERCENT, value),
  );
  return Math.round(clamped * 10) / 10;
}

/**
 * Convert a vertical pointer coordinate into bottom-panel height percentage.
 * The function is allocation-free because pointermove can call it every frame.
 */
export function runOutputPanelPercentFromPointer(
  containerTop: number,
  containerHeight: number,
  clientY: number,
  fallback = RUN_OUTPUT_PANEL_DEFAULT_PERCENT,
): number {
  if (
    !Number.isFinite(containerTop) ||
    !Number.isFinite(containerHeight) ||
    containerHeight <= 0 ||
    !Number.isFinite(clientY)
  ) {
    return normalizeRunOutputPanelPercent(fallback);
  }
  const percent =
    ((containerTop + containerHeight - clientY) / containerHeight) * 100;
  return normalizeRunOutputPanelPercent(percent, fallback);
}

/**
 * Keyboard contract for the horizontal separator. Arrow/Page keys move the
 * split, while Home/End expose deterministic minimum/maximum positions.
 */
export function runOutputPanelPercentForKey(
  current: number,
  key: string,
): number | null {
  const value = normalizeRunOutputPanelPercent(current);
  if (key === "ArrowUp") {
    return normalizeRunOutputPanelPercent(
      value + RUN_OUTPUT_PANEL_KEYBOARD_STEP,
    );
  }
  if (key === "ArrowDown") {
    return normalizeRunOutputPanelPercent(
      value - RUN_OUTPUT_PANEL_KEYBOARD_STEP,
    );
  }
  if (key === "PageUp") {
    return normalizeRunOutputPanelPercent(
      value + RUN_OUTPUT_PANEL_KEYBOARD_STEP * 2,
    );
  }
  if (key === "PageDown") {
    return normalizeRunOutputPanelPercent(
      value - RUN_OUTPUT_PANEL_KEYBOARD_STEP * 2,
    );
  }
  if (key === "Home") return RUN_OUTPUT_PANEL_MIN_PERCENT;
  if (key === "End") return RUN_OUTPUT_PANEL_MAX_PERCENT;
  return null;
}

/** Restore one global workbench preference without trusting hand-edited data. */
export function loadRunOutputPanelPercent(
  fallback = RUN_OUTPUT_PANEL_DEFAULT_PERCENT,
): number {
  const safeFallback = normalizeRunOutputPanelPercent(fallback);
  try {
    const raw = globalThis.localStorage?.getItem(
      OUTPUT_PANEL_PERCENT_STORAGE_KEY,
    );
    if (raw === null || raw.trim() === "") return safeFallback;
    const parsed = Number(raw);
    return Number.isFinite(parsed)
      ? normalizeRunOutputPanelPercent(parsed, safeFallback)
      : safeFallback;
  } catch {
    return safeFallback;
  }
}

/** Persist the normalized split without allowing storage policy to break UI. */
export function saveRunOutputPanelPercent(value: number): number {
  const normalized = normalizeRunOutputPanelPercent(value);
  try {
    globalThis.localStorage?.setItem(
      OUTPUT_PANEL_PERCENT_STORAGE_KEY,
      String(normalized),
    );
  } catch {
    /* The in-memory split remains usable for this window. */
  }
  return normalized;
}

/**
 * User-gesture clipboard write with a legacy DOM fallback. Returns false rather
 * than throwing when the browser or permission policy rejects both paths.
 */
export async function copyRunOutputText(text: string): Promise<boolean> {
  if (!text) return false;
  try {
    if (globalThis.navigator?.clipboard?.writeText) {
      await globalThis.navigator.clipboard.writeText(text);
      return true;
    }
  } catch {
    // Fall through to the synchronous selection-based copy path.
  }

  if (typeof document === "undefined") return false;
  const textarea = document.createElement("textarea");
  textarea.value = text;
  textarea.readOnly = true;
  textarea.style.position = "fixed";
  textarea.style.opacity = "0";
  document.body.appendChild(textarea);
  try {
    textarea.select();
    return document.execCommand?.("copy") === true;
  } catch {
    return false;
  } finally {
    textarea.remove();
  }
}
