/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Runtime output budgets for code execution.
 *
 * Persistence already has independent size limits, but applying those only when
 * a cell is saved is too late: a runaway program can first grow an unbounded
 * worker array, flood postMessage, and allocate the same payload on the main
 * thread. These helpers enforce the boundary while output is produced and once
 * more when an untrusted worker/iframe response crosses into the application.
 */

import type { LogStream } from "./cell-output-format";
import type { RichOutput, RunResult } from "./js-runner";
import {
  isLanguageExecutionErrorCode,
  type LanguageExecutionErrorCode,
} from "./execution-error-code";

export const MAX_EXECUTION_LOG_LINE_CHARS = 20_000;
export const MAX_EXECUTION_LOG_LINES = 2_000;
export const MAX_EXECUTION_LOG_CHARS = 300_000;
export const MAX_EXECUTION_TEXT_CHARS = 50_000;
export const MAX_EXECUTION_HTML_CHARS = 200_000;
export const MAX_EXECUTION_IMAGE_CHARS = 1_000_000;
export const MAX_EXECUTION_IMAGE_TOTAL_CHARS = 1_000_000;
export const MAX_EXECUTION_IMAGES = 20;

const MAX_EXECUTION_TABLE_ROWS = 1_000;
const MAX_EXECUTION_TABLE_COLS = 100;
const MAX_EXECUTION_TABLE_CELL_CHARS = 2_000;
const MAX_EXECUTION_TABLE_JSON_CHARS = 500_000;

const RUN_RESULT_KEYS = [
  "logs",
  "logStreams",
  "result",
  "error",
  "executionErrorCode",
  "timedOut",
  "interrupted",
  "notStarted",
  "images",
  "imageMimes",
  "html",
  "rich",
] as const;
const RICH_OUTPUT_KEYS = [
  "kind",
  "html",
  "mime",
  "data",
  "columns",
  "rows",
] as const;

export const EXECUTION_OUTPUT_TRUNCATION_MARKER =
  "… [output truncated: execution budget reached]";

const ALLOWED_RASTER_MIMES = new Set([
  "image/png",
  "image/jpeg",
  "image/gif",
  "image/webp",
  "image/bmp",
]);

/**
 * Copy only own data properties from an untrusted result envelope. Accessor
 * properties are rejected without reading them, so a worker/version-skewed
 * payload cannot execute a getter on the application thread.
 */
function ownDataFields(
  value: unknown,
  keys: readonly string[],
): Record<string, unknown> | null {
  if (!value || typeof value !== "object") return null;
  try {
    if (Array.isArray(value)) return null;
  } catch {
    return null;
  }
  const fields: Record<string, unknown> = Object.create(null);
  for (const key of keys) {
    let descriptor: PropertyDescriptor | undefined;
    try {
      descriptor = Object.getOwnPropertyDescriptor(value, key);
    } catch {
      return null;
    }
    if (!descriptor) continue;
    if (!("value" in descriptor)) return null;
    fields[key] = descriptor.value;
  }
  return fields;
}

/**
 * Copy a bounded array prefix by property descriptor. Sparse entries stay
 * undefined and indexed accessors reject the array without being invoked.
 */
function ownArrayPrefix(value: unknown, maxItems: number): unknown[] | null {
  let isArray = false;
  try {
    isArray = Array.isArray(value);
  } catch {
    return null;
  }
  if (!isArray) return null;
  let lengthDescriptor: PropertyDescriptor | undefined;
  try {
    lengthDescriptor = Object.getOwnPropertyDescriptor(value, "length");
  } catch {
    return null;
  }
  if (
    !lengthDescriptor ||
    !("value" in lengthDescriptor) ||
    !Number.isSafeInteger(lengthDescriptor.value) ||
    lengthDescriptor.value < 0
  ) {
    return null;
  }
  const length = Math.min(lengthDescriptor.value, Math.max(0, maxItems));
  const items: unknown[] = [];
  for (let index = 0; index < length; index += 1) {
    let descriptor: PropertyDescriptor | undefined;
    try {
      descriptor = Object.getOwnPropertyDescriptor(value, String(index));
    } catch {
      return null;
    }
    if (!descriptor) {
      items.push(undefined);
      continue;
    }
    if (!("value" in descriptor)) return null;
    items.push(descriptor.value);
  }
  return items;
}

function truncateWithMarker(value: string, maxChars: number): string {
  if (value.length <= maxChars) return value;
  if (maxChars <= EXECUTION_OUTPUT_TRUNCATION_MARKER.length) {
    return EXECUTION_OUTPUT_TRUNCATION_MARKER.slice(0, maxChars);
  }
  return (
    value.slice(0, maxChars - EXECUTION_OUTPUT_TRUNCATION_MARKER.length) +
    EXECUTION_OUTPUT_TRUNCATION_MARKER
  );
}

export interface ExecutionLogCapture {
  readonly logs: string[];
  readonly logStreams: LogStream[];
  readonly truncated: boolean;
  /** Returns the retained text (for progressive streaming), or null if sealed. */
  push(text: unknown, stream: LogStream): string | null;
}

/** Bounded aggregate capture used inside synchronous/in-process runtimes. */
export function createExecutionLogCapture(): ExecutionLogCapture {
  const logs: string[] = [];
  const logStreams: LogStream[] = [];
  let totalChars = 0;
  let sealed = false;

  const appendMarker = (): string | null => {
    if (sealed) return null;
    sealed = true;
    const remaining = MAX_EXECUTION_LOG_CHARS - totalChars;
    if (remaining <= 0 || logs.length >= MAX_EXECUTION_LOG_LINES) return null;
    const marker = EXECUTION_OUTPUT_TRUNCATION_MARKER.slice(0, remaining);
    logs.push(marker);
    logStreams.push("stderr");
    totalChars += marker.length;
    return marker;
  };

  return {
    logs,
    logStreams,
    get truncated() {
      return sealed;
    },
    push(text, stream) {
      if (sealed) return null;
      // Reserve the last entry for a visible marker. This makes the retained
      // aggregate and the progressively streamed sequence agree exactly.
      if (logs.length >= MAX_EXECUTION_LOG_LINES - 1) return appendMarker();
      let raw: string;
      try {
        raw = typeof text === "string" ? text : String(text);
      } catch {
        raw = "<unprintable>";
      }
      let retained = truncateWithMarker(raw, MAX_EXECUTION_LOG_LINE_CHARS);
      const remaining =
        MAX_EXECUTION_LOG_CHARS -
        totalChars -
        EXECUTION_OUTPUT_TRUNCATION_MARKER.length;
      if (retained.length > remaining) {
        if (remaining <= 0) return appendMarker();
        retained = truncateWithMarker(retained, remaining);
        logs.push(retained);
        logStreams.push(stream === "stderr" ? "stderr" : "stdout");
        totalChars += retained.length;
        sealed = true;
        return retained;
      }
      logs.push(retained);
      logStreams.push(stream === "stderr" ? "stderr" : "stdout");
      totalChars += retained.length;
      return retained;
    },
  };
}

export interface ExecutionStreamLimiter {
  /** Returns a bounded chunk, one marker, or null after the budget is sealed. */
  push(chunk: unknown): string | null;
}

/** Host-side rate/size guard for incremental worker and iframe messages. */
export function createExecutionStreamLimiter(): ExecutionStreamLimiter {
  let chunks = 0;
  let chars = 0;
  let sealed = false;
  return {
    push(chunk) {
      if (sealed || typeof chunk !== "string" || chunk === "") return null;
      if (chunks >= MAX_EXECUTION_LOG_LINES - 1) {
        sealed = true;
        chunks += 1;
        return EXECUTION_OUTPUT_TRUNCATION_MARKER;
      }
      let retained = truncateWithMarker(chunk, MAX_EXECUTION_LOG_LINE_CHARS);
      const remaining =
        MAX_EXECUTION_LOG_CHARS -
        chars -
        EXECUTION_OUTPUT_TRUNCATION_MARKER.length;
      if (retained.length > remaining) {
        sealed = true;
        if (remaining <= 0) return EXECUTION_OUTPUT_TRUNCATION_MARKER;
        retained = truncateWithMarker(retained, remaining);
      }
      chunks += 1;
      chars += retained.length;
      return retained;
    },
  };
}

function boundText(value: unknown, maxChars: number): string | undefined {
  return typeof value === "string"
    ? truncateWithMarker(value, maxChars)
    : undefined;
}

function boundTableCell(value: unknown): unknown {
  if (value === undefined) return undefined;
  if (value === null || typeof value === "boolean") return value;
  if (typeof value === "number")
    return Number.isFinite(value) ? value : String(value);
  if (typeof value === "string")
    return truncateWithMarker(value, MAX_EXECUTION_TABLE_CELL_CHARS);
  // Current workers serialize complex rich-table cells before transport. Treat
  // an object/function arriving here as an untrusted or version-skewed payload;
  // do not JSON.stringify it on the UI thread (that would recreate the very
  // unbounded traversal this boundary exists to prevent).
  return "<complex output omitted>";
}

function boundRichOutput(value: unknown): RichOutput | undefined {
  const raw = ownDataFields(value, RICH_OUTPUT_KEYS);
  if (!raw) return undefined;
  if (raw.kind === "html") {
    const html = boundText(raw.html, MAX_EXECUTION_HTML_CHARS);
    return html === undefined ? undefined : { kind: "html", html };
  }
  if (raw.kind === "image") {
    if (
      typeof raw.mime !== "string" ||
      !ALLOWED_RASTER_MIMES.has(raw.mime) ||
      typeof raw.data !== "string" ||
      raw.data.length > MAX_EXECUTION_IMAGE_CHARS
    ) {
      return undefined;
    }
    return { kind: "image", mime: raw.mime, data: raw.data };
  }
  if (raw.kind !== "table") return undefined;
  const rawColumns = ownArrayPrefix(raw.columns, MAX_EXECUTION_TABLE_COLS);
  const columns = (rawColumns ?? []).map((column) =>
    truncateWithMarker(
      typeof column === "string" ? column : "<invalid column>",
      1_000,
    ),
  );
  const rawRows = ownArrayPrefix(raw.rows, MAX_EXECUTION_TABLE_ROWS);
  if (!rawRows) return undefined;
  const rows: unknown[][] = [];
  let used = JSON.stringify({ kind: "table", columns, rows }).length;
  for (const candidate of rawRows) {
    const rawRow = ownArrayPrefix(candidate, MAX_EXECUTION_TABLE_COLS);
    if (!rawRow) continue;
    const row = rawRow.map(boundTableCell);
    const rowChars = JSON.stringify(row).length + 1;
    if (used + rowChars > MAX_EXECUTION_TABLE_JSON_CHARS) break;
    rows.push(row);
    used += rowChars;
  }
  return { kind: "table", columns, rows };
}

/**
 * Validate and bound an untrusted final worker/iframe result. The transport may
 * be buggy, version-skewed, or directly reachable from authored code, so do not
 * cast and expose its arrays/strings without this last boundary.
 */
export function boundExecutionResult(value: unknown): RunResult {
  const raw = ownDataFields(value, RUN_RESULT_KEYS) ?? {};
  const capture = createExecutionLogCapture();
  const rawLogs = ownArrayPrefix(raw.logs, MAX_EXECUTION_LOG_LINES) ?? [];
  const rawStreams = ownArrayPrefix(raw.logStreams, MAX_EXECUTION_LOG_LINES);
  for (
    let index = 0;
    index < rawLogs.length && !capture.truncated;
    index += 1
  ) {
    const log = rawLogs[index];
    // This is a cross-runtime trust boundary.  Worker transports should only
    // send already-formatted strings; coercing an object here would execute an
    // attacker-controlled toString hook on the UI thread.
    if (typeof log !== "string") continue;
    capture.push(log, rawStreams?.[index] === "stderr" ? "stderr" : "stdout");
  }

  const result: {
    logs: string[];
    logStreams?: LogStream[];
    result?: string;
    error?: string;
    executionErrorCode?: LanguageExecutionErrorCode;
    timedOut?: boolean;
    interrupted?: boolean;
    notStarted?: boolean;
    images?: string[];
    imageMimes?: string[];
    html?: string;
    rich?: RichOutput;
  } = { logs: capture.logs };
  if (rawStreams) result.logStreams = capture.logStreams;
  const resultText = boundText(raw.result, MAX_EXECUTION_TEXT_CHARS);
  if (resultText !== undefined) result.result = resultText;
  const error = boundText(raw.error, MAX_EXECUTION_TEXT_CHARS);
  if (error !== undefined) result.error = error;
  if (isLanguageExecutionErrorCode(raw.executionErrorCode)) {
    result.executionErrorCode = raw.executionErrorCode;
  }
  if (raw.timedOut === true) result.timedOut = true;
  if (raw.interrupted === true) result.interrupted = true;
  if (raw.notStarted === true) result.notStarted = true;
  const html = boundText(raw.html, MAX_EXECUTION_HTML_CHARS);
  if (html !== undefined) result.html = html;

  const rawImages = ownArrayPrefix(raw.images, MAX_EXECUTION_IMAGES);
  if (rawImages) {
    const rawMimes = ownArrayPrefix(raw.imageMimes, MAX_EXECUTION_IMAGES);
    const images: string[] = [];
    const mimes: string[] = [];
    let imageChars = 0;
    for (
      let index = 0;
      index < rawImages.length && images.length < MAX_EXECUTION_IMAGES;
      index += 1
    ) {
      const image = rawImages[index];
      if (
        typeof image !== "string" ||
        image.length > MAX_EXECUTION_IMAGE_CHARS ||
        imageChars + image.length > MAX_EXECUTION_IMAGE_TOTAL_CHARS
      ) {
        continue;
      }
      images.push(image);
      imageChars += image.length;
      const mime = rawMimes?.[index];
      mimes.push(
        typeof mime === "string" && ALLOWED_RASTER_MIMES.has(mime)
          ? mime
          : "image/png",
      );
    }
    if (images.length > 0) {
      result.images = images;
      if (mimes.some((mime) => mime !== "image/png")) {
        result.imageMimes = mimes;
      }
    }
  }

  const rich = boundRichOutput(raw.rich);
  if (rich) result.rich = rich;
  return result;
}
