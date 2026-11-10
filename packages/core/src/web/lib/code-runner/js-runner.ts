/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Client-side JavaScript execution for the code editor.
 *
 * User code runs in a dedicated Web Worker (isolated from the page DOM and
 * terminable), so an infinite loop or a throw cannot freeze the app. Console
 * output is captured and returned. A timeout terminates runaway code.
 *
 * `executeCodeSync` is the pure capture-and-run core (also mirrored inside the
 * worker script); it is unit-tested directly. `runJavaScript` is the worker
 * orchestration with an injectable worker factory for testing.
 */

import type { LogStream } from "./cell-output-format";
import type { LanguageExecutionErrorCode } from "./execution-error-code";
import {
  MAX_DEBUG_SNAPSHOT_EXPRESSION_CHARS,
  MAX_DEBUG_SNAPSHOT_VALUE_CHARS,
  normalizeDebugSnapshotExpression,
  type DebugSnapshotEvaluationReceipt,
} from "./debug-snapshot-evaluate";
import {
  DEBUG_HOOK_NAME,
  instrumentForDebug,
  type InstrumentResult,
} from "./js-debug";
import type { OnOutput } from "./output-coalescer";
import {
  EXECUTION_OUTPUT_TRUNCATION_MARKER,
  MAX_EXECUTION_LOG_CHARS,
  MAX_EXECUTION_LOG_LINE_CHARS,
  MAX_EXECUTION_LOG_LINES,
  MAX_EXECUTION_TEXT_CHARS,
  boundExecutionResult,
  createExecutionLogCapture,
  createExecutionStreamLimiter,
} from "./execution-output-budget";
import {
  endSandboxSession,
  createSandboxDebugWorker,
  getBrowserChannel,
  isSandboxExecutionEnabled,
  runInSandbox,
  type SandboxPromptMode,
} from "./sandbox";

const MAX_DEBUG_PAUSED_LOCALS = 24;
const MAX_DEBUG_PAUSED_LOCAL_NAME_CHARS = 200;
const MAX_DEBUG_PAUSED_LOCAL_VALUE_CHARS = 200;

/**
 * A single structured rich output (Jupyter-style), kept as data — never as raw
 * HTML — so the renderer can display it without ever interpreting untrusted
 * markup. `table` renders as a native React `<table>` (React escapes every
 * cell); `image` as an `<img>` from a `data:` URI; `html` still goes through the
 * locked-down sandbox iframe in CellOutput (no allow-scripts).
 */
export type RichOutput =
  | {
      readonly kind: "table";
      /** Column headers, already stringified for display. */
      readonly columns: readonly string[];
      /** Row-major cell values; rendered as text (React-escaped). */
      readonly rows: readonly (readonly unknown[])[];
    }
  | {
      readonly kind: "image";
      /** Image MIME type, e.g. "image/png". */
      readonly mime: string;
      /** Base64-encoded image bytes (no data: prefix). */
      readonly data: string;
    }
  | {
      readonly kind: "html";
      /** Raw HTML; rendered ONLY inside a script-disabled sandbox iframe. */
      readonly html: string;
    };

export interface RunResult {
  readonly logs: string[];
  /**
   * An authoritative, ordered nbformat output snapshot for an imported Jupyter
   * cell. Native executions use the aggregate fields below; imported notebooks
   * keep this sequence so interleaved streams, repeated display bundles and
   * multiple errors are rendered in their original order instead of being
   * collapsed into a lossy projection.
   */
  readonly notebookOutputs?: readonly Record<string, unknown>[];
  /**
   * Per-line stream tag for `logs`, aligned by index ("stdout" | "stderr").
   * Optional and additive: when absent (legacy output) every line renders as
   * stdout. A shorter array than `logs` also defaults the untagged tail to
   * stdout (see `streamForLogIndex`). This keeps `logs` a plain `string[]`.
   */
  readonly logStreams?: readonly LogStream[];
  readonly result?: string;
  readonly error?: string;
  /** Trusted executor classification; never inferred from authored stderr. */
  readonly executionErrorCode?: LanguageExecutionErrorCode;
  readonly timedOut?: boolean;
  /** The run was aborted by the user (Stop / interrupt) before it finished. */
  readonly interrupted?: boolean;
  /**
   * The execution service declined the run before it started (busy, quota,
   * outage): `error` is its status message, not the program's failure.
   */
  readonly notStarted?: boolean;
  /** Rich outputs (Jupyter-style): base64 raster figures, e.g. from matplotlib. */
  readonly images?: readonly string[];
  /** Per-image MIME types aligned to `images`; absent ⇒ all `image/png`. */
  readonly imageMimes?: readonly string[];
  /** Rich HTML output (e.g. a pandas DataFrame _repr_html_); rendered sandboxed. */
  readonly html?: string;
  /**
   * A structured rich output (table / image / html) for the final value, when
   * the runner can recognise its shape. Optional and additive: existing outputs
   * are unaffected when this is absent.
   */
  readonly rich?: RichOutput;
}

/** Largest tabular result we promote to a `rich` table (anti-freeze guard). */
const RICH_TABLE_ROW_LIMIT = 1000;
const RICH_TABLE_COL_LIMIT = 100;
const RICH_TABLE_CELL_CHAR_LIMIT = 2_000;
const RICH_TABLE_COLUMN_CHAR_LIMIT = 1_000;
const RICH_TABLE_JSON_CHAR_LIMIT = 500_000;

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return (
    typeof value === "object" &&
    value !== null &&
    !Array.isArray(value) &&
    (Object.getPrototypeOf(value) === Object.prototype ||
      Object.getPrototypeOf(value) === null)
  );
}

function boundedRichTableCell(value: unknown): unknown {
  if (value === undefined) return undefined;
  if (value === null || typeof value === "boolean") return value;
  if (typeof value === "number")
    return Number.isFinite(value) ? value : String(value);
  const rendered = typeof value === "string" ? value : display(value);
  return rendered.length <= RICH_TABLE_CELL_CHAR_LIMIT
    ? rendered
    : `${rendered.slice(
        0,
        RICH_TABLE_CELL_CHAR_LIMIT - "[output truncated]".length,
      )}[output truncated]`;
}

function boundedRichRows(
  source: readonly unknown[],
  columns: readonly string[],
  rowFor: (value: unknown) => unknown[],
): unknown[][] {
  const rows: unknown[][] = [];
  let used = JSON.stringify({ kind: "table", columns, rows }).length;
  for (const value of source) {
    const row = rowFor(value).map(boundedRichTableCell);
    const rowChars = JSON.stringify(row).length + 1;
    if (used + rowChars > RICH_TABLE_JSON_CHAR_LIMIT) break;
    rows.push(row);
    used += rowChars;
  }
  return rows;
}

/**
 * Recognise a JS value that should render as a Jupyter-style table:
 *   - an array of plain objects (`[{a:1,b:2}, …]`) → columns = union of keys;
 *   - a 2D array (`[[1,2],[3,4]]`) → positional columns "0","1",….
 * Anything else returns `undefined` so callers fall back to the text result.
 * Pure and capped so a giant array can't build a huge payload.
 */
export function detectRichTable(value: unknown): RichOutput | undefined {
  if (!Array.isArray(value) || value.length === 0) return undefined;
  const rowCount = Math.min(value.length, RICH_TABLE_ROW_LIMIT);
  // Never scan rows that cannot be retained. Besides avoiding wasted work,
  // this keeps a huge/proxied authored value from turning table detection into
  // an unbounded main/worker-thread traversal.
  const retainedRows = value.slice(0, rowCount);

  // Array-of-objects: every element a plain object (no arrays).
  if (retainedRows.every((row) => isPlainObject(row))) {
    const columns: string[] = [];
    const accessKeys: string[] = [];
    const seen = new Set<string>();
    for (const row of retainedRows) {
      for (const key in row as Record<string, unknown>) {
        if (!Object.prototype.hasOwnProperty.call(row, key)) continue;
        if (!seen.has(key)) {
          seen.add(key);
          accessKeys.push(key);
          columns.push(key.slice(0, RICH_TABLE_COLUMN_CHAR_LIMIT));
          if (columns.length >= RICH_TABLE_COL_LIMIT) break;
        }
      }
      if (columns.length >= RICH_TABLE_COL_LIMIT) break;
    }
    if (columns.length === 0) return undefined;
    const rows = boundedRichRows(retainedRows, columns, (row) => {
      const record = row as Record<string, unknown>;
      return accessKeys.map((key) =>
        Object.prototype.hasOwnProperty.call(record, key)
          ? record[key]
          : undefined,
      );
    });
    return { kind: "table", columns, rows };
  }

  // 2D array: every element an array (positional columns).
  if (retainedRows.every((row) => Array.isArray(row))) {
    const colCount = Math.min(
      Math.max(0, ...retainedRows.map((row) => (row as unknown[]).length)),
      RICH_TABLE_COL_LIMIT,
    );
    if (colCount === 0) return undefined;
    const columns = Array.from({ length: colCount }, (_, i) => String(i));
    const rows = boundedRichRows(retainedRows, columns, (row) =>
      columns.map((_, i) => (row as unknown[])[i]),
    );
    return { kind: "table", columns, rows };
  }

  return undefined;
}

/**
 * Console-grade value formatter, as embeddable plain-JS SOURCE so the TS test
 * core and BOTH inlined worker scripts share one implementation. The old
 * `JSON.stringify` lost data: a Map/Set/typed-array serialized to "{}", and an
 * object key whose value was `undefined`/a function vanished — silently, because
 * the string flowed straight into logs, the persisted snapshot, and .ipynb
 * export. This renders Map/Set/typed-arrays/cycles/undefined faithfully.
 */
const CONSOLE_FORMAT_SRC = `
  var FORMAT_MAX_NODES = 512;
  var FORMAT_MAX_ENTRIES = 100;
  var FORMAT_MAX_STRING_CHARS = 2000;
  var FORMAT_TRUNCATION = "[output truncated]";
  function formatTruncate(value, maxChars) {
    value = String(value);
    return value.length <= maxChars
      ? value
      : value.slice(0, Math.max(0, maxChars - FORMAT_TRUNCATION.length)) + FORMAT_TRUNCATION;
  }
  function formatValue(value, seen, depth, budget) {
    seen = seen || new Set();
    depth = depth || 0;
    budget = budget || { nodes: 0 };
    budget.nodes += 1;
    if (budget.nodes > FORMAT_MAX_NODES) return FORMAT_TRUNCATION;
    var t = typeof value;
    if (t === "string") return depth === 0 ? value : JSON.stringify(formatTruncate(value, FORMAT_MAX_STRING_CHARS));
    if (value === undefined) return "undefined";
    if (value === null) return "null";
    if (t === "number" || t === "boolean") return String(value);
    if (t === "bigint") return String(value) + "n";
    if (t === "symbol") return String(value);
    if (t === "function") {
      return value.name ? "[Function: " + value.name + "]" : "[Function (anonymous)]";
    }
    if (seen.has(value)) return "[Circular]";
    if (depth > 6) return Array.isArray(value) ? "[Array]" : "[Object]";
    seen.add(value);
    try {
      if (value instanceof Map) {
        var me = [], mi = value.entries(), ms, mc = 0;
        while (mc < FORMAT_MAX_ENTRIES && !(ms = mi.next()).done) {
          me.push(formatValue(ms.value[0], seen, depth + 1, budget) + " => " + formatValue(ms.value[1], seen, depth + 1, budget));
          mc += 1;
        }
        if (value.size > mc) me.push(FORMAT_TRUNCATION);
        return "Map(" + value.size + ") {" + me.join(", ") + "}";
      }
      if (value instanceof Set) {
        var se = [], si = value.values(), ss, sc = 0;
        while (sc < FORMAT_MAX_ENTRIES && !(ss = si.next()).done) {
          se.push(formatValue(ss.value, seen, depth + 1, budget));
          sc += 1;
        }
        if (value.size > sc) se.push(FORMAT_TRUNCATION);
        return "Set(" + value.size + ") {" + se.join(", ") + "}";
      }
      if (Array.isArray(value)) {
        var ae = [], al = Math.min(value.length, FORMAT_MAX_ENTRIES);
        for (var ai = 0; ai < al; ai += 1) ae.push(formatValue(value[ai], seen, depth + 1, budget));
        if (value.length > al) ae.push(FORMAT_TRUNCATION);
        return "[" + ae.join(", ") + "]";
      }
      if (typeof ArrayBuffer !== "undefined" && ArrayBuffer.isView(value) && !(value instanceof DataView)) {
        var tl = Math.min(value.length, FORMAT_MAX_ENTRIES), ta = [];
        for (var ti = 0; ti < tl; ti += 1) ta.push(String(value[ti]));
        if (value.length > tl) ta.push(FORMAT_TRUNCATION);
        var taName = (value.constructor && value.constructor.name) || "TypedArray";
        return taName + "(" + value.length + ") [" + ta.join(", ") + "]";
      }
      if (value instanceof Date) {
        return isNaN(value.getTime()) ? "Invalid Date" : value.toISOString();
      }
      if (value instanceof RegExp) return value.toString();
      if (value instanceof Error) return (value.name || "Error") + ": " + value.message;
      var parts = [], keyCount = 0, hasMore = false;
      for (var key in value) {
        if (!Object.prototype.hasOwnProperty.call(value, key)) continue;
        if (keyCount >= FORMAT_MAX_ENTRIES) { hasMore = true; break; }
        parts.push(key + ": " + formatValue(value[key], seen, depth + 1, budget));
        keyCount += 1;
      }
      if (hasMore) parts.push(FORMAT_TRUNCATION);
      var body = parts.join(", ");
      var ctor = value.constructor && value.constructor.name;
      var tag = ctor && ctor !== "Object" ? ctor + " " : "";
      return tag + "{" + body + "}";
    } catch (e) {
      try { return Object.prototype.toString.call(value); }
      catch (ignored) { return "<unprintable>"; }
    } finally {
      seen.delete(value);
    }
  }
`;

let formatValueFn: ((value: unknown) => string) | null = null;
/** Lazily build the formatter (uses new Function, same as the run path). */
function getFormatValue(): (value: unknown) => string {
  if (!formatValueFn) {
    formatValueFn = new Function(
      `${CONSOLE_FORMAT_SRC}\nreturn formatValue;`,
    )() as (value: unknown) => string;
  }
  return formatValueFn;
}

/** Serialize a console argument for display (console-grade, lossless for Map/Set). */
function display(value: unknown): string {
  if (typeof value === "string") return value;
  try {
    return getFormatValue()(value);
  } catch {
    try {
      return Object.prototype.toString.call(value);
    } catch {
      return "<unprintable>";
    }
  }
}

/**
 * The rest of a browser's / Node's console for code run in the sandbox:
 * table (a box-drawn table, as Node prints it), dir, trace, assert,
 * count/countReset, time/timeLog/timeEnd and group/groupCollapsed/groupEnd
 * (indenting what follows). Plain JS, embedded in both worker blobs and
 * evaluated once for the synchronous path, so all three behave alike.
 * `sink(stream)` is the console line writer, `display` its value printer.
 */
const CONSOLE_EXTRAS_SRC = `
  function extendSandboxConsole(target, sink, display) {
    var indent = "";
    var counts = Object.create(null);
    var timers = Object.create(null);
    function write(stream) {
      var put = sink(stream);
      return function () {
        var args = Array.prototype.slice.call(arguments);
        if (!indent) return put.apply(null, args);
        var text = args.map(display).join(" ");
        put(indent + text.split("\\n").join("\\n" + indent));
      };
    }
    ["log", "info", "debug"].forEach(function (name) { target[name] = write("stdout"); });
    ["warn", "error"].forEach(function (name) { target[name] = write("stderr"); });
    target.dir = function (value) { target.log(value); };
    target.dirxml = target.log;
    target.trace = function () {
      var args = Array.prototype.slice.call(arguments);
      target.error.apply(null, ["Trace:"].concat(args));
    };
    target.assert = function (condition) {
      if (condition) return;
      var args = Array.prototype.slice.call(arguments, 1);
      target.error.apply(null, ["Assertion failed" + (args.length ? ":" : "")].concat(args));
    };
    target.count = function (label) {
      var key = label === undefined ? "default" : String(label);
      counts[key] = (counts[key] || 0) + 1;
      target.log(key + ": " + counts[key]);
    };
    target.countReset = function (label) { counts[label === undefined ? "default" : String(label)] = 0; };
    function now() { return typeof performance !== "undefined" ? performance.now() : Date.now(); }
    function elapsed(key) { return Math.round((now() - timers[key]) * 1000) / 1000 + "ms"; }
    target.time = function (label) { timers[label === undefined ? "default" : String(label)] = now(); };
    target.timeLog = function (label) {
      var key = label === undefined ? "default" : String(label);
      if (!(key in timers)) return target.warn("Timer '" + key + "' does not exist");
      target.log.apply(null, [key + ": " + elapsed(key)].concat(Array.prototype.slice.call(arguments, 1)));
    };
    target.timeEnd = function (label) {
      var key = label === undefined ? "default" : String(label);
      if (!(key in timers)) return target.warn("Timer '" + key + "' does not exist");
      target.log(key + ": " + elapsed(key));
      delete timers[key];
    };
    target.group = function () {
      if (arguments.length) target.log.apply(null, arguments);
      indent += "  ";
    };
    target.groupCollapsed = target.group;
    target.groupEnd = function () { indent = indent.slice(0, -2); };
    target.table = function (data, properties) {
      if (data === null || typeof data !== "object") return target.log(data);
      var rows = [], indexKey = "(index)", valueKey = "Values";
      var columns = [], seen = Object.create(null), hasValues = false;
      var keys = Object.keys(data);
      for (var i = 0; i < keys.length && i < 1000; i += 1) {
        var key = keys[i], item = data[key], row = { index: key, cells: Object.create(null), value: undefined };
        if (item !== null && typeof item === "object") {
          var itemKeys = Object.keys(item);
          for (var j = 0; j < itemKeys.length; j += 1) {
            var column = itemKeys[j];
            if (properties && properties.indexOf(column) < 0) continue;
            if (!seen[column]) { seen[column] = true; columns.push(column); }
            row.cells[column] = display(item[column]);
          }
        } else {
          hasValues = true;
          row.value = display(item);
        }
        rows.push(row);
      }
      if (properties) columns = properties.filter(function (c) { return seen[c]; });
      var header = [indexKey].concat(columns, hasValues ? [valueKey] : []);
      var body = rows.map(function (row) {
        return [String(row.index)].concat(
          columns.map(function (c) { return c in row.cells ? row.cells[c] : ""; }),
          hasValues ? [row.value === undefined ? "" : row.value] : []
        );
      });
      var widths = header.map(function (h, c) {
        return body.reduce(function (w, r) { return Math.max(w, String(r[c]).length); }, String(h).length) + 2;
      });
      function center(text, width) {
        text = String(text);
        var left = Math.floor((width - text.length) / 2);
        return new Array(left + 1).join(" ") + text + new Array(width - text.length - left + 1).join(" ");
      }
      function rule(l, m, r) { return l + widths.map(function (w) { return new Array(w + 1).join("─"); }).join(m) + r; }
      var lines = [rule("┌", "┬", "┐"), "│" + header.map(function (h, c) { return center(h, widths[c]); }).join("│") + "│", rule("├", "┼", "┤")];
      body.forEach(function (r) { lines.push("│" + r.map(function (v, c) { return center(v, widths[c]); }).join("│") + "│"); });
      lines.push(rule("└", "┴", "┘"));
      target.log(lines.join("\\n"));
    };
    return target;
  }
`;

type ExtendSandboxConsole = (
  target: Record<string, unknown>,
  sink: (stream: LogStream) => (...args: unknown[]) => void,
  display: (value: unknown) => string,
) => Record<string, (...args: unknown[]) => void>;
let extendSandboxConsoleSync: ExtendSandboxConsole | null = null;

/**
 * The synchronous path's copy of the same extras (one source for all three).
 * Built on first use, never at import: the app origin's CSP has no
 * 'unsafe-eval', and only a realm that already runs code with new Function
 * (the sandbox, a test) calls this.
 */
function extendSandboxConsole(
  ...args: Parameters<ExtendSandboxConsole>
): ReturnType<ExtendSandboxConsole> {
  extendSandboxConsoleSync ??= new Function(
    `${CONSOLE_EXTRAS_SRC}\nreturn extendSandboxConsole;`,
  )() as ExtendSandboxConsole;
  return extendSandboxConsoleSync(...args);
}

/**
 * Run a JS source string with console capture, synchronously. Pure: no DOM,
 * no network setup. This is the core used by both the worker and tests.
 */
export function executeCodeSync(code: string): RunResult {
  const capture = createExecutionLogCapture();
  const sink =
    (stream: LogStream) =>
    (...args: unknown[]) => {
      capture.push(args.map(display).join(" "), stream);
    };
  const sandboxConsole = extendSandboxConsole({}, sink, display);
  let result: string | undefined;
  let error: string | undefined;
  let rich: RichOutput | undefined;
  try {
    const fn = new Function("console", `"use strict";\n${code}`);
    const ret = fn(sandboxConsole);
    if (ret !== undefined) {
      result = display(ret);
      rich = detectRichTable(ret);
    }
  } catch (err) {
    error = err instanceof Error ? `${err.name}: ${err.message}` : String(err);
  }
  return boundExecutionResult({
    logs: capture.logs,
    logStreams: capture.logStreams,
    result,
    error,
    rich,
  });
}

// Plain-JS capture helper embedded into both worker blobs. Keep the numeric
// values sourced from the shared host constants so worker and host budgets
// cannot drift during refactors.
const WORKER_OUTPUT_BUDGET_SRC = `
  var MAX_OUTPUT_LINES = ${MAX_EXECUTION_LOG_LINES};
  var MAX_OUTPUT_LINE_CHARS = ${MAX_EXECUTION_LOG_LINE_CHARS};
  var MAX_OUTPUT_LOG_CHARS = ${MAX_EXECUTION_LOG_CHARS};
  var MAX_OUTPUT_TEXT_CHARS = ${MAX_EXECUTION_TEXT_CHARS};
  var OUTPUT_TRUNCATION_MARKER = ${JSON.stringify(EXECUTION_OUTPUT_TRUNCATION_MARKER)};
  function truncateOutput(value, maxChars) {
    value = String(value);
    if (value.length <= maxChars) return value;
    if (maxChars <= OUTPUT_TRUNCATION_MARKER.length) return OUTPUT_TRUNCATION_MARKER.slice(0, maxChars);
    return value.slice(0, maxChars - OUTPUT_TRUNCATION_MARKER.length) + OUTPUT_TRUNCATION_MARKER;
  }
  function createOutputCapture() {
    var logs = [], logStreams = [], chars = 0, sealed = false;
    function marker() {
      if (sealed) return null;
      sealed = true;
      var remaining = MAX_OUTPUT_LOG_CHARS - chars;
      if (remaining <= 0 || logs.length >= MAX_OUTPUT_LINES) return null;
      var kept = OUTPUT_TRUNCATION_MARKER.slice(0, remaining);
      logs.push(kept); logStreams.push("stderr"); chars += kept.length;
      return { text: kept, stream: "stderr" };
    }
    return {
      logs: logs,
      logStreams: logStreams,
      push: function (value, streamName) {
        if (sealed) return null;
        if (logs.length >= MAX_OUTPUT_LINES - 1) return marker();
        var kept;
        try { kept = truncateOutput(value, MAX_OUTPUT_LINE_CHARS); }
        catch (e) { kept = "<unprintable>"; }
        var remaining = MAX_OUTPUT_LOG_CHARS - chars - OUTPUT_TRUNCATION_MARKER.length;
        if (kept.length > remaining) {
          if (remaining <= 0) return marker();
          kept = truncateOutput(kept, remaining);
          sealed = true;
        }
        var tag = streamName === "stderr" ? "stderr" : "stdout";
        logs.push(kept); logStreams.push(tag); chars += kept.length;
        return { text: kept, stream: tag };
      }
    };
  }
`;

const WORKER_RICH_TABLE_SRC = `
  var ROW_LIMIT = ${RICH_TABLE_ROW_LIMIT}, COL_LIMIT = ${RICH_TABLE_COL_LIMIT};
  var TABLE_CELL_LIMIT = ${RICH_TABLE_CELL_CHAR_LIMIT};
  var TABLE_COLUMN_LIMIT = ${RICH_TABLE_COLUMN_CHAR_LIMIT};
  var TABLE_JSON_LIMIT = ${RICH_TABLE_JSON_CHAR_LIMIT};
  function isPlainObject(v) {
    if (typeof v !== "object" || v === null || Array.isArray(v)) return false;
    var p = Object.getPrototypeOf(v);
    return p === Object.prototype || p === null;
  }
  function tableCell(v) {
    if (v === undefined) return undefined;
    if (v === null || typeof v === "boolean") return v;
    if (typeof v === "number") return Number.isFinite(v) ? v : String(v);
    var rendered = typeof v === "string" ? v : display(v);
    return rendered.length <= TABLE_CELL_LIMIT
      ? rendered
      : rendered.slice(0, TABLE_CELL_LIMIT - "[output truncated]".length) + "[output truncated]";
  }
  function tableRows(source, columns, rowFor) {
    var rows = [];
    var used = JSON.stringify({ kind: "table", columns: columns, rows: rows }).length;
    for (var i = 0; i < source.length; i++) {
      var row = rowFor(source[i]).map(tableCell);
      var rowChars = JSON.stringify(row).length + 1;
      if (used + rowChars > TABLE_JSON_LIMIT) break;
      rows.push(row); used += rowChars;
    }
    return rows;
  }
  function detectRichTable(value) {
    if (!Array.isArray(value) || value.length === 0) return undefined;
    var rowCount = Math.min(value.length, ROW_LIMIT);
    var retainedRows = value.slice(0, rowCount);
    if (retainedRows.every(isPlainObject)) {
      var columns = [], accessKeys = [], seen = Object.create(null);
      for (var i = 0; i < retainedRows.length && columns.length < COL_LIMIT; i++) {
        for (var key in retainedRows[i]) {
          if (!Object.prototype.hasOwnProperty.call(retainedRows[i], key) || seen[key]) continue;
          seen[key] = true; accessKeys.push(key); columns.push(key.slice(0, TABLE_COLUMN_LIMIT));
          if (columns.length >= COL_LIMIT) break;
        }
      }
      if (columns.length === 0) return undefined;
      var rows = tableRows(retainedRows, columns, function (row) {
        return accessKeys.map(function (key) {
          return Object.prototype.hasOwnProperty.call(row, key) ? row[key] : undefined;
        });
      });
      return { kind: "table", columns: columns, rows: rows };
    }
    if (retainedRows.every(function (row) { return Array.isArray(row); })) {
      var colCount = Math.min(
        Math.max.apply(null, [0].concat(retainedRows.map(function (row) { return row.length; }))),
        COL_LIMIT
      );
      if (colCount === 0) return undefined;
      var cols = [];
      for (var c = 0; c < colCount; c++) cols.push(String(c));
      var rows2 = tableRows(retainedRows, cols, function (row) {
        return cols.map(function (_, idx) { return row[idx]; });
      });
      return { kind: "table", columns: cols, rows: rows2 };
    }
    return undefined;
  }
`;

// The worker mirrors executeCodeSync as a self-contained string (workers can't
// import module code without bundling).
const WORKER_SCRIPT = `
self.onmessage = function (event) {
  ${WORKER_OUTPUT_BUDGET_SRC}
  var output = createOutputCapture();
  var logs = output.logs;
  var logStreams = output.logStreams;
  // When the caller opts into progressive output, post each console line to the
  // main thread immediately (Jupyter streaming parity). For synchronous code the
  // worker thread is busy until the run returns, so these queue and deliver when
  // it yields; async code (awaits) pumps them mid-run.
  var stream = event.data && event.data.stream === true;
  ${CONSOLE_FORMAT_SRC}
  function display(value) {
    if (typeof value === "string") return value;
    try { return formatValue(value); }
    catch (e) {
      try { return Object.prototype.toString.call(value); }
      catch (ignored) { return "<unprintable>"; }
    }
  }
  function sink(streamName) { return function () {
    var args = Array.prototype.slice.call(arguments);
    var line = args.map(display).join(" ");
    var kept = output.push(line, streamName);
    if (stream && kept) self.postMessage({ type: "stream", stream: kept.stream, chunk: kept.text + "\\n" });
  }; }
  ${CONSOLE_EXTRAS_SRC}
  var sandboxConsole = extendSandboxConsole({}, sink, display);
  ${WORKER_RICH_TABLE_SRC}
  var result, error, rich;
  // Where it was thrown in the person's code, as the sandbox page reports it.
  function where(e){var out=[];try{var lines=String(e&&e.stack||'').split('\\n');for(var i=0;i<lines.length&&out.length<10;i+=1){var m=/(?:<anonymous>|> Function):(\\d+):(\\d+)\\)?\\s*$/.exec(lines[i]);if(!m)continue;var line=Number(m[1])-3;if(!(line>=1))continue;var f=/^\\s*at ([^\\s(]+) \\(/.exec(lines[i])||/^([^@\\s]+)@/.exec(lines[i]);var fn=f&&f[1]!=='eval'&&f[1]!=='anonymous'?f[1]:'';out.push('    at '+(fn?fn+' (<anonymous>:'+line+':'+m[2]+')':'<anonymous>:'+line+':'+m[2]))}if(!out.length&&e&&typeof e.line==='number'&&e.line-3>=1)out.push('    at <anonymous>:'+(e.line-3)+':'+(typeof e.column==='number'?e.column:0))}catch(ignored){}return out.length?'\\n'+out.join('\\n'):''}
  try {
    var fn = new Function("console", '"use strict";\\n' + event.data.code);
    var ret = fn(sandboxConsole);
    if (ret !== undefined) { result = truncateOutput(display(ret), MAX_OUTPUT_TEXT_CHARS); rich = detectRichTable(ret); }
  } catch (err) {
    error = truncateOutput(((err && err.name && err.message) ? (err.name + ": " + err.message) : String(err)) + where(err), MAX_OUTPUT_TEXT_CHARS);
  }
  self.postMessage({ logs: logs, logStreams: logStreams, result: result, error: error, rich: rich });
};
`;

/**
 * An incremental output message the worker posts before the final result (when
 * streaming is requested): a single stdout/stderr chunk. Distinguished from the
 * final {@link RunResult} message by `type: "stream"`.
 */
export interface WorkerStreamMessage {
  type: "stream";
  stream: LogStream;
  chunk: string;
}

/**
 * A debug-pause message the DEBUG worker posts when instrumented code reaches a
 * breakpoint (or a pending single-step). The worker then awaits a
 * {@link WorkerDebugResumeMessage} before continuing. Only ever emitted on the
 * flag-gated debug run path; the normal run path never produces it.
 */
export interface WorkerDebugPausedMessage {
  type: "debug-paused";
  /** 1-based source line the run paused at. */
  line: number;
  /** Best-effort `{ name: displayString }` of simple in-scope locals. */
  locals: Record<string, string>;
  /** 1-based ordinal of this pause within the run. */
  step: number;
  /** Uncaught-exception pauses are terminal: Continue publishes the final error. */
  reason?: "breakpoint" | "exception";
  /** Bounded exception display text, present only when `reason === "exception"`. */
  exception?: string;
}

/** The reply the main thread posts to release a paused debug run. */
export interface WorkerDebugResumeMessage {
  type: "debug-resume";
  /** "continue" = run to the next breakpoint; "step" = pause at the next hooked line. */
  mode: "continue" | "step";
}

/** One exact paused-frame expression result from the isolated debug worker. */
export interface WorkerDebugEvaluatedMessage
  extends DebugSnapshotEvaluationReceipt {
  type: "debug-evaluated";
  /** Unguessable request owner; late/foreign replies are ignored. */
  requestId: string;
  /** The expression created asynchronous work, so this worker must be retired. */
  retireSession?: true;
}

/** Any message the debug worker may post (superset of the normal worker's). */
export type WorkerDebugOutboundMessage =
  | RunResult
  | WorkerStreamMessage
  | WorkerDebugPausedMessage
  | WorkerDebugEvaluatedMessage;

/**
 * Fail closed at the worker boundary before a paused receipt reaches any UI
 * callback. The worker already emits this shape, but an injected or drifted
 * transport must not be able to hand raw objects or unbounded strings to the
 * host under the trusted `debug-paused` discriminator.
 */
function normalizeWorkerDebugPausedMessage(
  value: unknown,
  sourceLineCount: number,
): WorkerDebugPausedMessage | null {
  if (!isPlainObject(value) || value.type !== "debug-paused") return null;

  const line = value.line;
  const step = value.step;
  if (
    typeof line !== "number" ||
    !Number.isSafeInteger(line) ||
    line < 1 ||
    line > sourceLineCount ||
    typeof step !== "number" ||
    !Number.isSafeInteger(step) ||
    step < 1
  ) {
    return null;
  }

  const reason = value.reason;
  if (
    reason !== undefined &&
    reason !== "breakpoint" &&
    reason !== "exception"
  ) {
    return null;
  }
  const exception = value.exception;
  if (
    (exception !== undefined &&
      (typeof exception !== "string" ||
        exception.length > MAX_EXECUTION_TEXT_CHARS)) ||
    (reason === "exception" && typeof exception !== "string") ||
    (reason !== "exception" && exception !== undefined)
  ) {
    return null;
  }

  if (!isPlainObject(value.locals)) return null;
  let entries: [string, unknown][];
  try {
    entries = Object.entries(value.locals);
  } catch {
    return null;
  }
  if (entries.length > MAX_DEBUG_PAUSED_LOCALS) return null;

  const locals = Object.create(null) as Record<string, string>;
  for (const [name, localValue] of entries) {
    if (
      name.length === 0 ||
      name.length > MAX_DEBUG_PAUSED_LOCAL_NAME_CHARS ||
      typeof localValue !== "string" ||
      localValue.length > MAX_DEBUG_PAUSED_LOCAL_VALUE_CHARS
    ) {
      return null;
    }
    locals[name] = localValue;
  }

  return {
    type: "debug-paused",
    line,
    locals,
    step,
    ...(reason === undefined ? {} : { reason }),
    ...(exception === undefined ? {} : { exception }),
  };
}

export interface WorkerLike {
  postMessage(data: unknown): void;
  terminate(): void;
  onmessage:
    | ((event: { data: RunResult | WorkerStreamMessage }) => void)
    | null;
  onerror: ((event: { message?: string }) => void) | null;
}

function defaultCreateWorker(): WorkerLike {
  const blob = new Blob([WORKER_SCRIPT], { type: "application/javascript" });
  const url = URL.createObjectURL(blob);
  const worker = new Worker(url) as unknown as WorkerLike;
  // The worker keeps its script once constructed, so the object URL can be
  // released immediately to avoid leaking a blob per run.
  URL.revokeObjectURL(url);
  return worker;
}

// ---------------------------------------------------------------------------
// Debug worker (flag-gated breakpoints + step) — see lib/code-runner/js-debug.ts
// ---------------------------------------------------------------------------

// The DEBUG worker mirrors the capture core but (a) runs the (already-
// instrumented) user code inside an async wrapper so the breakpoint hook can be
// awaited, (b) installs `__mrnoteDebugHook(line, localsThunk)` which pauses at a
// breakpoint (or pending step) by posting a "debug-paused" message and awaiting a
// "debug-resume" reply, and (c) installs `__dbgFmt` for the locals snapshot. The
// hook is a NO-OP fast path when the line is not a breakpoint and no step is
// pending, so an instrumented program with no active breakpoints runs to
// completion unchanged. This script is ONLY used by `runJavaScriptDebug`; the
// normal run path's WORKER_SCRIPT is byte-for-byte untouched.
const DEBUG_WORKER_SCRIPT = `
self.onmessage = function (event) {
  var data = event.data || {};
  if (data.type === "debug-resume") return; // handled via the resume queue below
  // Every protocol message this harness sends is wrapped in a token envelope,
  // exactly as the normal run worker does, and BOTH consumers drop anything
  // without the token. The user's own program runs in this same worker and can
  // call postMessage; without the envelope it could post {logs:[...]} to end
  // its own debug run with a fabricated result, or {type:"debug-paused"} to
  // clear the host's runaway timer and never time out again. The token is a
  // closure variable the user's scope is never handed, and nativePost is bound
  // before user code runs so overwriting self.postMessage cannot redirect it.
  var protocolToken = data.protocolToken;
  var nativePost = self.postMessage.bind(self);
  // Capture the few intrinsics used by the trusted pause protocol before the
  // authored program can replace their global bindings. In particular, locals
  // formatting must never call a helper exposed on the worker global: user code could
  // replace that helper and make a raw object cross the worker boundary.
  var nativeObjectCreate = Object.create;
  var nativeObjectKeys = Object.keys;
  var nativeObjectGetOwnPropertyDescriptor = Object.getOwnPropertyDescriptor;
  var nativeObjectGetPrototypeOf = Object.getPrototypeOf;
  var nativeString = String;
  var nativeReflectApply = Reflect.apply;
  var nativeStringSlice = String.prototype.slice;
  var NativeError = Error;
  var NativeFunction = Function;
  var NativeProxy = Proxy;
  var nativeAddEventListener = self.addEventListener.bind(self);
  var nativeClose = typeof self.close === "function" ? self.close.bind(self) : null;
  function trustedStringSlice(value, start, end) {
    return nativeReflectApply(nativeStringSlice, value, [start, end]);
  }
  function trustedPost(payload) {
    nativePost({ protocolToken: protocolToken, payload: payload });
  }
  ${WORKER_OUTPUT_BUDGET_SRC}
  var output = createOutputCapture();
  var logs = output.logs;
  var logStreams = output.logStreams;
  var stream = data.stream === true;
  ${CONSOLE_FORMAT_SRC}
  function display(value) {
    if (typeof value === "string") return value;
    try { return formatValue(value); }
    catch (e) {
      try { return Object.prototype.toString.call(value); }
      catch (ignored) { return "<unprintable>"; }
    }
  }
  function sink(streamName) { return function () {
    var args = Array.prototype.slice.call(arguments);
    var line = args.map(display).join(" ");
    var kept = output.push(line, streamName);
    if (stream && kept) trustedPost({ type: "stream", stream: kept.stream, chunk: kept.text + "\\n" });
  }; }
  ${CONSOLE_EXTRAS_SRC}
  var sandboxConsole = extendSandboxConsole({}, sink, display);

  // Active breakpoint set + step-pending flag, mutable across pauses.
  var breakpoints = {};
  (data.breakpoints || []).forEach(function (n) { breakpoints[n] = true; });
  var stepPending = false;
  var pauseCount = 0;
  // Every injected hook updates this location even when it does not pause. If
  // the program later rejects, this is the last trustworthy ORIGINAL source
  // line (instrumentation changes generated line numbers, so stack parsing
  // would point at the wrong line).
  var currentLine = null;
  var currentLocalsThunk = null;
  // Raw values never cross the worker boundary. They live only for the exact
  // active pause so the trusted harness can evaluate one bounded expression in
  // that lexical snapshot; resume/finish/termination drops the snapshot.
  var currentPauseLocals = null;

  // Resume coordination: while paused we await a promise resolved by the next
  // "debug-resume" message. A single pending resolver at a time (debugging is
  // strictly sequential — one pause outstanding).
  var resumeResolver = null;
  function expressionIsValid(value) {
    return typeof value === "string" && value.length > 0 &&
      value.length <= ${MAX_DEBUG_SNAPSHOT_EXPRESSION_CHARS} &&
      !/[\\0\\r\\n\\u2028\\u2029]/u.test(value);
  }
  function boundedEvaluationValue(value) {
    var text = display(value);
    if (typeof text !== "string") text = nativeString(text);
    return text.length <= ${MAX_DEBUG_SNAPSHOT_VALUE_CHARS}
      ? text
      : trustedStringSlice(text, 0, ${MAX_DEBUG_SNAPSHOT_VALUE_CHARS - 1}) + "…";
  }
  function boundedEvaluationError(error) {
    var text = "Evaluation failed.";
    try {
      text = error && error.name && error.message
        ? error.name + ": " + error.message
        : nativeString(error);
    } catch (ignored) {}
    if (typeof text !== "string") text = "Evaluation failed.";
    return text.length <= ${MAX_DEBUG_SNAPSHOT_VALUE_CHARS}
      ? text
      : trustedStringSlice(text, 0, ${MAX_DEBUG_SNAPSHOT_VALUE_CHARS - 1}) + "…";
  }
  function boundedDebugLocal(value) {
    if (typeof value !== "string") return "<unserializable>";
    return value.length <= ${MAX_DEBUG_PAUSED_LOCAL_VALUE_CHARS}
      ? value
      : trustedStringSlice(value, 0, ${MAX_DEBUG_PAUSED_LOCAL_VALUE_CHARS - 1}) + "…";
  }
  function displayLocals(raw) {
    var visible = nativeObjectCreate(null);
    if (!raw || typeof raw !== "object") return visible;
    var names = [];
    try { names = nativeObjectKeys(raw); }
    catch (ignored) { return visible; }
    for (var index = 0; index < names.length && index < ${MAX_DEBUG_PAUSED_LOCALS}; index += 1) {
      var name = names[index];
      if (typeof name !== "string" || name.length === 0 || name.length > ${MAX_DEBUG_PAUSED_LOCAL_NAME_CHARS}) continue;
      try {
        visible[name] = boundedDebugLocal(formatDebugLocal(raw[name]));
      } catch (ignored) { visible[name] = "<unserializable>"; }
    }
    return visible;
  }
  function evaluationHasThenDescriptor(value) {
    if (value === null || (typeof value !== "object" && typeof value !== "function")) {
      return false;
    }
    var owner = value;
    // Never read the value's then property: an accessor is authored code and can schedule an
    // unbounded microtask chain before returning or throwing. Inspect captured
    // descriptors instead. A hostile Proxy that throws from either reflective
    // operation is also unsafe and retires the generation.
    for (var depth = 0; owner !== null && depth < 32; depth += 1) {
      var descriptor = nativeObjectGetOwnPropertyDescriptor(owner, "then");
      if (descriptor !== undefined) return true;
      owner = nativeObjectGetPrototypeOf(owner);
    }
    return owner !== null;
  }
  function retireEvaluation(requestId, error) {
    trustedPost({
      type: "debug-evaluated",
      requestId: requestId,
      status: "error",
      value: boundedEvaluationError(error),
      retireSession: true
    });
    // The host also terminates after consuming this exact receipt. Closing here
    // prevents authored microtasks queued by the expression from receiving a
    // second turn while that cross-thread receipt is in flight.
    try { if (nativeClose) nativeClose(); } catch (ignored) {}
  }
  function evaluateAtPause(requestId, expression) {
    if (
      !currentPauseLocals ||
      !resumeResolver ||
      typeof requestId !== "string" ||
      requestId.length !== 32 ||
      !/^[0-9a-f]+$/.test(requestId) ||
      !expressionIsValid(expression)
    ) {
      return;
    }
    try {
      var safeGlobals = {
        Math: Math, JSON: JSON, Number: Number, String: String,
        Boolean: Boolean, BigInt: BigInt, Array: Array, Object: Object,
        Date: Date, RegExp: RegExp
      };
      var scope = new NativeProxy(currentPauseLocals, {
        has: function () { return true; },
        get: function (target, key) {
          if (key === Symbol.unscopables) return undefined;
          if (Object.prototype.hasOwnProperty.call(target, key)) return target[key];
          if (Object.prototype.hasOwnProperty.call(safeGlobals, key)) return safeGlobals[key];
          return undefined;
        },
        set: function () { return false; }
      });
      // Deliberately expression-only. It runs in the same credential-free,
      // network-denied worker as the authored program and cannot access the UI.
      var value = new NativeFunction("scope", "with(scope){return (" + expression + ");}")(scope);
      try {
        if (evaluationHasThenDescriptor(value)) {
          retireEvaluation(
            requestId,
            new NativeError("Async evaluation is unavailable while paused."),
          );
          return;
        }
      } catch (error) {
        retireEvaluation(requestId, error);
        return;
      }
      trustedPost({
        type: "debug-evaluated",
        requestId: requestId,
        status: "success",
        value: boundedEvaluationValue(value)
      });
    } catch (error) {
      trustedPost({
        type: "debug-evaluated",
        requestId: requestId,
        status: "error",
        value: boundedEvaluationError(error)
      });
    }
  }
  // The authored program can read, call, or overwrite self.onmessage. Keep
  // the debugger's control authority in an unreferenced event listener instead
  // and accept browser-delivered messages only. Synthetic dispatch from the
  // expression has isTrusted false; replacing the public property cannot
  // remove this listener.
  self.onmessage = null;
  nativeAddEventListener("message", function (ev) {
    if (!ev || ev.isTrusted !== true) return;
    var d = ev.data || {};
    if (d.type === "debug-resume") {
      var r = resumeResolver;
      resumeResolver = null;
      currentPauseLocals = null;
      if (r) r(d.mode === "step" ? "step" : "continue");
      return;
    }
    if (d.type === "debug-evaluate") {
      evaluateAtPause(d.requestId, d.expression);
      return;
    }
    // Ignore any further run requests while a debug run is active.
  });

  // Best-effort value formatter for the locals snapshot (short, never throws).
  // It remains a closure-local trusted helper; the authored AsyncFunction has
  // no binding through which it can replace it.
  function formatDebugLocal(v) {
    try {
      if (typeof v === "function") return "[function]";
      if (typeof v === "string") return v;
      if (typeof v === "object" && v !== null) {
        var s = display(v);
        return typeof s === "string" ? s : "<unserializable>";
      }
      return nativeString(v);
    } catch (e) { return "<unserializable>"; }
  }

  // The breakpoint hook. Returns a promise; resolves immediately (no-op) unless
  // the line is a breakpoint or a step is pending. The optional third argument
  // is a CONDITION verdict precomputed in the user scope (conditional
  // breakpoints): false skips the pause — but never a pending STEP, which must
  // stop on the next hooked line unconditionally.
  self.${DEBUG_HOOK_NAME} = function (line, localsThunk, condHit) {
    currentLine = line;
    currentLocalsThunk = localsThunk || null;
    if (!breakpoints[line] && !stepPending) return Promise.resolve();
    if (!stepPending && condHit === false) return Promise.resolve();
    stepPending = false;
    pauseCount += 1;
    var locals = {};
    try { locals = localsThunk ? localsThunk() : {}; } catch (e) { locals = {}; }
    currentPauseLocals = locals;
    trustedPost({ type: "debug-paused", reason: "breakpoint", line: line, locals: displayLocals(locals), step: pauseCount });
    return new Promise(function (resolve) {
      resumeResolver = function (mode) {
        if (mode === "step") stepPending = true;
        resolve();
      };
    });
  };

  function exceptionText(err) {
    var name = "";
    var message = "";
    var fallback = "Uncaught exception";
    try { if (err && typeof err.name === "string") name = err.name; } catch (ignored) {}
    try { if (err && typeof err.message === "string") message = err.message; } catch (ignored) {}
    if (name && message) fallback = name + ": " + message;
    else if (message) fallback = message;
    else {
      try { fallback = String(err); } catch (ignored) {}
    }
    return truncateOutput(fallback, MAX_OUTPUT_TEXT_CHARS);
  }

  function exceptionLine(err) {
    if (typeof currentLine === "number" && currentLine >= 1) return currentLine;
    var sourceLineCount =
      typeof data.sourceLineCount === "number" && data.sourceLineCount >= 1
        ? Math.floor(data.sourceLineCount)
        : 1;
    // An instrumented stack refers to generated hook lines. Without a hook hit
    // there is no exact reverse map, so fail closed to the first source line.
    if (data.codeWasInstrumented === true) return 1;
    try {
      var stack = err && typeof err.stack === "string" ? err.stack : "";
      var match = /<anonymous>:(\d+):(\d+)/.exec(stack);
      if (match) {
        // AsyncFunction adds two wrapper lines and this worker prepends the
        // explicit strict-mode line, so authored line N appears as N + 3.
        var authoredLine = Number(match[1]) - 3;
        if (Number.isFinite(authoredLine) && authoredLine >= 1) {
          return Math.min(Math.floor(authoredLine), sourceLineCount);
        }
      }
    } catch (ignored) {}
    return 1;
  }

  // An uncaught exception has already unwound user code, but the most recent
  // injected locals thunk retains the lexical snapshot boundary. Pause before
  // publishing the terminal receipt; Continue then exposes the exact same
  // bounded error the non-pausing debug path would have returned.
  function pauseOnUncaughtException(err) {
    var text = exceptionText(err);
    if (data.pauseOnUncaughtExceptions !== true) return Promise.resolve(text);
    pauseCount += 1;
    var locals = {};
    try { locals = currentLocalsThunk ? currentLocalsThunk() : {}; } catch (ignored) { locals = {}; }
    currentPauseLocals = locals;
    trustedPost({
      type: "debug-paused",
      reason: "exception",
      line: exceptionLine(err),
      locals: displayLocals(locals),
      step: pauseCount,
      exception: text
    });
    return new Promise(function (resolve) {
      resumeResolver = function () {
        currentPauseLocals = null;
        resolve(text);
      };
    });
  }

  ${WORKER_RICH_TABLE_SRC}

  // Run the instrumented code inside an async wrapper so the awaited hook can
  // pause. We build an async function whose body is the user's (instrumented)
  // code; its completion value (the user's own return) resolves the promise,
  // mirroring new Function semantics with await support.
  var error;
  try {
    // eslint-disable-next-line no-new-func
    var AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
    var fn = new AsyncFunction("console", '"use strict";\\n' + data.code);
    fn(sandboxConsole).then(function (ret) {
      var result, rich;
      if (ret !== undefined) { result = truncateOutput(display(ret), MAX_OUTPUT_TEXT_CHARS); rich = detectRichTable(ret); }
      trustedPost({ logs: logs, logStreams: logStreams, result: result, error: undefined, rich: rich });
    }, function (err) {
      pauseOnUncaughtException(err).then(function (e) {
        trustedPost({ logs: logs, logStreams: logStreams, result: undefined, error: e });
      });
    });
  } catch (err) {
    error = truncateOutput((err && err.name && err.message) ? (err.name + ": " + err.message) : String(err), MAX_OUTPUT_TEXT_CHARS);
    trustedPost({ logs: logs, logStreams: logStreams, error: error });
  }
};
`;

/** Test-only visibility for executing the exact generated worker sources. */
export function __executionWorkerScriptsForTests(): {
  normal: string;
  debug: string;
} {
  return { normal: WORKER_SCRIPT, debug: DEBUG_WORKER_SCRIPT };
}

function defaultCreateDebugWorker(sandboxOwner?: string): WorkerLike {
  // Debug runs where every other execution runs. A blob Worker built here would
  // be on the APP's origin and inherits the app's CSP: in development that put
  // user JavaScript beside the session cookie, and in production it could not
  // evaluate anything at all, because 'unsafe-eval' was deliberately removed
  // from the app once execution moved into the sandbox page. Debug was the one
  // path that never followed, and no single environment showed both halves.
  if (isSandboxExecutionEnabled()) {
    return createSandboxDebugWorker(
      DEBUG_WORKER_SCRIPT,
      sandboxOwner,
    ) as WorkerLike;
  }
  // No document (SSR, and the in-process test path): there is no iframe to host
  // the session, so fall back to a local worker.
  const blob = new Blob([DEBUG_WORKER_SCRIPT], {
    type: "application/javascript",
  });
  const url = URL.createObjectURL(blob);
  const worker = new Worker(url) as unknown as WorkerLike;
  URL.revokeObjectURL(url);
  return withDebugProtocolEnvelope(worker);
}

/**
 * Give a debug worker the same token envelope the sandbox host applies, so both
 * transports check the same thing and the caller sees the same messages either
 * way.
 *
 * Without it the user's own program — which runs inside this worker — can post
 * a protocol message and be believed: `{logs:[...]}` ends its own debug run with
 * a result nothing computed, and `{type:"debug-paused"}` clears the runaway
 * timer so a later infinite loop never times out. The normal run path has always
 * had this envelope; the debug path was relayed verbatim, on the reasoning that
 * a host which does not interpret the protocol cannot drift from it. That was
 * right about drift and wrong about trust: the envelope is exactly what
 * separates the harness's messages from the program's.
 */
export function withDebugProtocolEnvelope(worker: WorkerLike): WorkerLike {
  const token = randomProtocolToken();
  let started = false;
  const wrapper: WorkerLike = {
    postMessage(data: unknown) {
      // Only the bootstrap message carries the outbound-envelope secret. The
      // private worker listener authenticates later controls with the browser's
      // unforgeable MessageEvent.isTrusted bit; exposing the token on Resume or
      // Evaluate would hand it to an authored `self.onmessage` observer.
      const message =
        !started && data && typeof data === "object" && !Array.isArray(data)
          ? { ...(data as Record<string, unknown>), protocolToken: token }
          : data;
      started = true;
      worker.postMessage(message);
    },
    terminate() {
      worker.terminate();
    },
    onmessage: null,
    onerror: null,
  };
  worker.onmessage = (event: { data: unknown }) => {
    const envelope = event.data as
      | { protocolToken?: unknown; payload?: unknown }
      | null;
    if (
      !envelope ||
      typeof envelope !== "object" ||
      envelope.protocolToken !== token ||
      !envelope.payload ||
      typeof envelope.payload !== "object"
    ) {
      return;
    }
    wrapper.onmessage?.({
      data: envelope.payload as RunResult | WorkerStreamMessage,
    });
  };
  worker.onerror = (event: { message?: string }) => wrapper.onerror?.(event);
  return wrapper;
}

/** An unguessable per-worker protocol token. */
function randomProtocolToken(): string {
  const bytes = new Uint8Array(16);
  if (typeof crypto !== "undefined" && crypto.getRandomValues) {
    crypto.getRandomValues(bytes);
  } else {
    for (let i = 0; i < bytes.length; i += 1) {
      // Only reachable where Web Crypto is absent; the token still has to be
      // present for the envelope check, and this path has no untrusted code.
      bytes[i] = (i * 31 + 7) % 256;
    }
  }
  return Array.from(bytes, (b) => b.toString(16).padStart(2, "0")).join("");
}

export interface RunOptions {
  timeoutMs?: number;
  createWorker?: () => WorkerLike;
  /** Page-scoped iframe owner; duplicate views pass the same stable key. */
  sandboxOwner?: string;
  /**
   * Abort the run (user Stop / interrupt): terminates the worker and resolves
   * with an `interrupted` result. A file run leaves nothing behind (the next run
   * starts fresh); a console `session` loses its realm, which is unavoidable —
   * terminating the worker is the only thing that stops code that will not yield.
   */
  signal?: AbortSignal;
  /**
   * Console session id: submissions sharing one also share a realm, so what the
   * prompt defines survives into the next submission. Only the isolated-sandbox
   * path keeps sessions; the in-page worker fallback (tests, SSR) runs each
   * submission fresh and simply ignores this.
   */
  session?: string;
  /**
   * Console prompt reporting (sandbox only): see SandboxRequest.prompt. An
   * "await" submission is an async function the sandbox waits for.
   */
  prompt?: SandboxPromptMode;
  /**
   * Progressive-output sink: called with each `(chunk, stream)` of console
   * output the worker produces, so the cell can render live (Jupyter streaming
   * parity). OPTIONAL/additive — when omitted behaviour is unchanged (final
   * batched result only). NOTE: a fully SYNCHRONOUS snippet holds the worker
   * thread until it returns, so its chunks arrive together right before the
   * final result; async code (with awaits) pumps them genuinely mid-run. The
   * final RunResult always carries the complete logs and is the source of truth.
   */
  onOutput?: OnOutput;
}

export const AWAIT_COMPLETION_SOURCE_PREFIX =
  "/*__MRNOTE_CODE_TEST_AWAIT_COMPLETION_V1__*/";

/**
 * Internal Test Explorer execution mode. The marker carries no authorization;
 * it selects the existing isolated async worker so an authored async test can
 * settle before the trusted provider publishes its final receipt.
 */
export function isAwaitCompletionJavaScript(source: string): boolean {
  return source.startsWith(AWAIT_COMPLETION_SOURCE_PREFIX);
}

function runJavaScriptAwaited(
  code: string,
  options: RunOptions,
): Promise<RunResult> {
  const timeoutMs = options.timeoutMs ?? 5000;
  const createWorker =
    options.createWorker ??
    (() => defaultCreateDebugWorker(options.sandboxOwner));
  const signal = options.signal;
  return new Promise<RunResult>((resolve) => {
    let settled = false;
    let worker: WorkerLike | null = null;
    const streamLimiter = createExecutionStreamLimiter();
    const finish = (result: RunResult) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      signal?.removeEventListener("abort", onAbort);
      try {
        worker?.terminate();
      } catch {
        /* already retired */
      }
      resolve(boundExecutionResult(result));
    };
    const onAbort = () =>
      finish({ logs: [], error: "Execution interrupted.", interrupted: true });
    const timer = setTimeout(
      () =>
        finish({
          logs: [],
          error: `Execution timed out after ${timeoutMs}ms`,
          timedOut: true,
        }),
      timeoutMs,
    );
    if (signal) {
      if (signal.aborted) {
        onAbort();
        return;
      }
      signal.addEventListener("abort", onAbort);
    }
    try {
      worker = createWorker();
      worker.onmessage = (event) => {
        const data = event.data as
          | RunResult
          | WorkerStreamMessage
          | WorkerDebugPausedMessage;
        if (data && (data as WorkerStreamMessage).type === "stream") {
          const message = data as WorkerStreamMessage;
          const chunk = streamLimiter.push(message.chunk);
          if (!settled && options.onOutput && chunk !== null) {
            options.onOutput(
              chunk,
              message.stream === "stderr" ? "stderr" : "stdout",
            );
          }
          return;
        }
        if (data && (data as WorkerDebugPausedMessage).type === "debug-paused") {
          finish({ logs: [], error: "Async execution protocol rejected." });
          return;
        }
        finish(data as RunResult);
      };
      worker.onerror = (event) =>
        finish({ logs: [], error: event.message || "Worker error" });
      worker.postMessage({
        type: "debug",
        code,
        codeWasInstrumented: false,
        sourceLineCount: Math.max(1, code.split(/\r\n?|\n/).length),
        breakpoints: [],
        pauseOnUncaughtExceptions: false,
        stream: Boolean(options.onOutput),
      });
    } catch (error) {
      finish({
        logs: [],
        error:
          error instanceof Error ? error.message : "Failed to start worker",
      });
    }
  });
}

/**
 * End a console session's realm (its globals, timers and intervals go with it).
 * A no-op where sessions are not kept — the in-page worker fallback never had
 * one to end.
 */
export function endJavaScriptSession(
  session: string,
  sandboxOwner?: string,
): void {
  if (!isSandboxExecutionEnabled()) return;
  endSandboxSession(session, getBrowserChannel(sandboxOwner));
}

/** Run JS in an isolated worker with a timeout. Always resolves. */
export function runJavaScript(
  code: string,
  options: RunOptions = {},
): Promise<RunResult> {
  if (isAwaitCompletionJavaScript(code)) {
    return runJavaScriptAwaited(code, options);
  }
  // When enabled, route through the isolated-origin sandbox (no app-origin
  // eval). Tests that inject a worker keep using the in-line worker path.
  if (!options.createWorker && isSandboxExecutionEnabled()) {
    return runInSandbox("javascript", code, {
      timeoutMs: options.timeoutMs,
      signal: options.signal,
      session: options.session,
      ...(options.prompt ? { prompt: options.prompt } : {}),
      sandboxOwner: options.sandboxOwner,
      onOutput: options.onOutput,
    });
  }
  const timeoutMs = options.timeoutMs ?? 5000;
  const createWorker = options.createWorker ?? defaultCreateWorker;
  const signal = options.signal;
  return new Promise<RunResult>((resolve) => {
    let settled = false;
    let worker: WorkerLike;
    const streamLimiter = createExecutionStreamLimiter();
    const finish = (res: RunResult) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (signal) signal.removeEventListener("abort", onAbort);
      try {
        worker?.terminate();
      } catch {
        /* ignore */
      }
      resolve(boundExecutionResult(res));
    };
    const onAbort = () =>
      finish({ logs: [], error: "Execution interrupted.", interrupted: true });
    const timer = setTimeout(
      () =>
        finish({
          logs: [],
          error: `Execution timed out after ${timeoutMs}ms`,
          timedOut: true,
        }),
      timeoutMs,
    );
    if (signal) {
      if (signal.aborted) {
        finish({
          logs: [],
          error: "Execution interrupted.",
          interrupted: true,
        });
        return;
      }
      signal.addEventListener("abort", onAbort);
    }
    try {
      worker = createWorker();
    } catch (err) {
      finish({
        logs: [],
        error: err instanceof Error ? err.message : "Failed to start worker",
      });
      return;
    }
    worker.onmessage = (event) => {
      const data = event.data;
      // An incremental stream chunk (progressive output): route it to onOutput
      // WITHOUT settling — the final result message (no `type`) settles the run.
      if (data && (data as WorkerStreamMessage).type === "stream") {
        const msg = data as WorkerStreamMessage;
        const chunk = streamLimiter.push(msg.chunk);
        if (!settled && options.onOutput && chunk !== null) {
          options.onOutput(
            chunk,
            msg.stream === "stderr" ? "stderr" : "stdout",
          );
        }
        return;
      }
      finish(data as RunResult);
    };
    worker.onerror = (event) =>
      finish({ logs: [], error: event.message || "Worker error" });
    // `stream: true` asks the worker to also post incremental chunks; older
    // workers ignore the flag and degrade to final-batch behaviour.
    worker.postMessage({ code, stream: Boolean(options.onOutput) });
  });
}

// ---------------------------------------------------------------------------
// Debug run orchestration (flag-gated; NEVER on the normal run path)
// ---------------------------------------------------------------------------

/** A live debug session handle: drive the run, then resume each pause. */
export interface DebugRunOptions {
  /** 1-based breakpoint lines (only these pause). */
  breakpoints?: readonly number[];
  /**
   * Per-line breakpoint conditions (VSCode 条件断点). The expression is baked
   * into that line's hook and evaluated in the USER scope right before the
   * pause; falsy skips it, an evaluation error counts as a hit. Entries must
   * pass isValidBreakpointCondition — invalid ones are ignored.
   */
  conditions?: Readonly<Record<number, string>>;
  /** Pause once before publishing an exception that escapes user code. */
  pauseOnUncaughtExceptions?: boolean;
  /** Called each time the run pauses; the UI shows line + locals here. */
  onPause: (pause: WorkerDebugPausedMessage) => void;
  timeoutMs?: number;
  /** Injectable debug worker (tests). */
  createWorker?: () => WorkerLike;
  /** Page-scoped iframe owner; normal run/debug/console share this boundary. */
  sandboxOwner?: string;
  /** Abort the run (Stop): terminates the worker, resolves as interrupted. */
  signal?: AbortSignal;
  onOutput?: OnOutput;
}

/** Controls a paused debug run; the resolved {@link RunResult} arrives via the promise. */
export interface DebugRunController {
  /** Resume to the next breakpoint. No-op if not currently paused. */
  continue(): void;
  /** Advance one hooked line, then pause again. No-op if not currently paused. */
  step(): void;
  /** Evaluate one bounded expression in the exact currently paused JS/TS frame. */
  evaluate(expression: string): Promise<DebugSnapshotEvaluationReceipt>;
  /** The eventual run result (always resolves; mirrors runJavaScript). */
  result: Promise<RunResult>;
}

/**
 * Run JS in DEBUG mode: instrument the source (best-effort, safe — see
 * {@link instrumentForDebug}), execute it in the debug worker, and pause at each
 * breakpoint, invoking `onPause`. The returned controller exposes `continue` /
 * `step` to release a pause and a `result` promise for the final output.
 *
 * This path is ONLY reached when the FE flag is on AND the caller explicitly asks
 * for a debug run; the normal {@link runJavaScript} path is untouched. If the
 * source cannot be safely instrumented (syntax/construct the lexer doesn't model)
 * the original source is run verbatim with NO hooks — so a debug run can never
 * miscompile the user's code, it just won't pause.
 *
 * NOTE: TypeScript debugging works by type-stripping first (the caller passes the
 * already-transpiled JS). PYTHON / Pyodide debugging is OUT OF SCOPE here.
 */
export function runJavaScriptDebug(
  code: string,
  options: DebugRunOptions,
): DebugRunController {
  const timeoutMs = options.timeoutMs ?? 5000;
  const createWorker =
    options.createWorker ??
    (() => defaultCreateDebugWorker(options.sandboxOwner));
  const signal = options.signal;

  // Instrument ALL safe statement lines (not just breakpoint lines): every hook
  // is a no-op fast path unless its line is an ACTIVE breakpoint OR a single-step
  // is pending — so hooking everywhere is what makes Step able to halt at the next
  // statement, while keeping zero pause overhead where no breakpoint is set. The
  // ACTIVE breakpoints are exactly the user's `breakpoints` (∅ when none),
  // intersected with the lines that actually received a hook.
  const instrumentation: InstrumentResult = instrumentForDebug(
    code,
    undefined,
    options.conditions,
  );
  const requested = options.breakpoints ?? [];
  const activeBreakpoints = instrumentation.hookedLines.filter((line) =>
    requested.includes(line),
  );
  const sourceLineCount = Math.max(1, code.split(/\r\n?|\n/).length);

  let worker: WorkerLike | null = null;
  let settled = false;
  let paused = false;
  let pauseGeneration = 0;
  let evaluationSequence = 0;
  let finishRun: (result: RunResult) => void = () => {};
  const pendingEvaluations = new Map<
    string,
    {
      readonly pauseGeneration: number;
      readonly resolve: (receipt: DebugSnapshotEvaluationReceipt) => void;
      readonly timer: ReturnType<typeof setTimeout>;
    }
  >();
  const retireEvaluations = (value = "Evaluation unavailable.") => {
    for (const pending of pendingEvaluations.values()) {
      clearTimeout(pending.timer);
      pending.resolve({ status: "error", value });
    }
    pendingEvaluations.clear();
  };
  // Timer controls the (paused-aware) promise body installs; declared here so
  // the synchronously-run executor can assign them without a TDZ hazard.
  let rearmTimer: () => void = () => {};
  let clearRunTimer: () => void = () => {};

  const result = new Promise<RunResult>((resolve) => {
    const streamLimiter = createExecutionStreamLimiter();
    const finish = (res: RunResult) => {
      if (settled) return;
      settled = true;
      paused = false;
      pauseGeneration += 1;
      retireEvaluations();
      clearTimeout(timer);
      if (signal) signal.removeEventListener("abort", onAbort);
      try {
        worker?.terminate();
      } catch {
        /* ignore */
      }
      resolve(boundExecutionResult(res));
    };
    finishRun = finish;
    const onAbort = () =>
      finish({ logs: [], error: "Execution interrupted.", interrupted: true });
    // The timeout is paused-aware: while halted at a breakpoint the clock is the
    // user's, so we only arm the runaway-timer while actually executing. It is
    // (re)armed on start and after each resume, and cleared on each pause.
    let timer = setTimeout(onTimeout, timeoutMs);
    function onTimeout() {
      finish({
        logs: [],
        error: `Execution timed out after ${timeoutMs}ms`,
        timedOut: true,
      });
    }
    // Expose a re-arm/clear pair to the controller via closure variables.
    rearmTimer = () => {
      clearTimeout(timer);
      timer = setTimeout(onTimeout, timeoutMs);
    };
    clearRunTimer = () => clearTimeout(timer);

    if (signal) {
      if (signal.aborted) {
        finish({
          logs: [],
          error: "Execution interrupted.",
          interrupted: true,
        });
        return;
      }
      signal.addEventListener("abort", onAbort);
    }
    try {
      worker = createWorker();
    } catch (err) {
      finish({
        logs: [],
        error: err instanceof Error ? err.message : "Failed to start worker",
      });
      return;
    }
    // The debug worker also posts WorkerDebugPausedMessage; WorkerLike's
    // `onmessage` data type is the narrower run union, so we cast `event.data`
    // up to the debug superset inside the handler (the assignment keeps the
    // WorkerLike-compatible signature).
    (worker as WorkerLike).onmessage = (event: {
      data: RunResult | WorkerStreamMessage;
    }) => {
      const data = event.data as
        | RunResult
        | WorkerStreamMessage
        | WorkerDebugPausedMessage
        | WorkerDebugEvaluatedMessage;
      if (data && (data as WorkerStreamMessage).type === "stream") {
        const msg = data as WorkerStreamMessage;
        const chunk = streamLimiter.push(msg.chunk);
        if (!settled && options.onOutput && chunk !== null) {
          options.onOutput(
            chunk,
            msg.stream === "stderr" ? "stderr" : "stdout",
          );
        }
        return;
      }
      if (data && (data as WorkerDebugPausedMessage).type === "debug-paused") {
        if (settled) return;
        const pause = normalizeWorkerDebugPausedMessage(data, sourceLineCount);
        if (!pause) {
          finish({
            logs: [],
            error: "Debug worker returned malformed paused state.",
          });
          return;
        }
        retireEvaluations();
        pauseGeneration += 1;
        paused = true;
        clearRunTimer(); // don't time out while halted at a breakpoint
        options.onPause(pause);
        return;
      }
      if (
        data &&
        (data as WorkerDebugEvaluatedMessage).type === "debug-evaluated"
      ) {
        const evaluated = data as WorkerDebugEvaluatedMessage;
        const pending = pendingEvaluations.get(evaluated.requestId);
        if (!pending) return;
        pendingEvaluations.delete(evaluated.requestId);
        clearTimeout(pending.timer);
        if (
          settled ||
          !paused ||
          pending.pauseGeneration !== pauseGeneration ||
          (evaluated.status !== "success" && evaluated.status !== "error") ||
          typeof evaluated.value !== "string" ||
          evaluated.value.length > MAX_DEBUG_SNAPSHOT_VALUE_CHARS ||
          (evaluated.retireSession !== undefined &&
            evaluated.retireSession !== true) ||
          (evaluated.retireSession === true && evaluated.status !== "error")
        ) {
          pending.resolve({
            status: "error",
            value: "Evaluation unavailable.",
          });
          return;
        }
        pending.resolve({
          status: evaluated.status,
          value: evaluated.value,
        });
        if (evaluated.retireSession === true) {
          finishRun({
            logs: [],
            error: "Paused evaluation retired the debug session.",
            interrupted: true,
          });
        }
        return;
      }
      finish(data as RunResult);
    };
    (worker as WorkerLike).onerror = (event: { message?: string }) =>
      finish({ logs: [], error: event.message || "Worker error" });
    (worker as WorkerLike).postMessage({
      type: "debug",
      code: instrumentation.code,
      codeWasInstrumented: instrumentation.instrumented,
      sourceLineCount,
      breakpoints: activeBreakpoints,
      pauseOnUncaughtExceptions: options.pauseOnUncaughtExceptions === true,
      stream: Boolean(options.onOutput),
    });
  });

  const resume = (mode: "continue" | "step") => {
    if (settled || !paused || !worker) return;
    retireEvaluations();
    pauseGeneration += 1;
    paused = false;
    rearmTimer(); // resume the runaway clock now that execution continues
    worker.postMessage({ type: "debug-resume", mode });
  };

  const evaluate = (
    rawExpression: string,
  ): Promise<DebugSnapshotEvaluationReceipt> => {
    const expression = normalizeDebugSnapshotExpression(rawExpression);
    if (!expression || settled || !paused || !worker) {
      return Promise.resolve({
        status: "error",
        value: "Evaluation unavailable.",
      });
    }
    evaluationSequence += 1;
    const requestId = `${randomProtocolToken().slice(0, 24)}${evaluationSequence
      .toString(16)
      .padStart(8, "0")
      .slice(-8)}`;
    const ownerPauseGeneration = pauseGeneration;
    return new Promise<DebugSnapshotEvaluationReceipt>((resolve) => {
      const timer = setTimeout(() => {
        const pending = pendingEvaluations.get(requestId);
        if (!pending) return;
        pendingEvaluations.delete(requestId);
        pending.resolve({
          status: "error",
          value: "Evaluation timed out.",
        });
        finishRun({
          logs: [],
          error: "Evaluation timed out.",
          timedOut: true,
        });
      }, Math.min(timeoutMs, 1_500));
      pendingEvaluations.set(requestId, {
        pauseGeneration: ownerPauseGeneration,
        resolve,
        timer,
      });
      try {
        worker?.postMessage({
          type: "debug-evaluate",
          requestId,
          expression,
        });
      } catch (error) {
        const pending = pendingEvaluations.get(requestId);
        if (pending) {
          pendingEvaluations.delete(requestId);
          clearTimeout(pending.timer);
          pending.resolve({
            status: "error",
            value: "Evaluation unavailable.",
          });
        }
        finishRun({
          logs: [],
          error: error instanceof Error ? error.message : "Worker error",
        });
      }
    });
  };

  return {
    continue: () => resume("continue"),
    step: () => resume("step"),
    evaluate,
    result,
  };
}
