/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Office-Scripts-style user scripting for the spreadsheet — the modern,
 * sandbox-safe equivalent of Excel's VBA macros.
 *
 * Microsoft's own VBA successor (Office Scripts) is sandboxed JavaScript driving
 * a workbook object model. We do the same WITHOUT touching the security-critical
 * sandbox runtime: a "snapshot-in / writes-out" model. We harness-wrap the user's
 * script with (a) an inline, frozen `sheet` object model backed by a JSON
 * SNAPSHOT of the active sheet (a run reads its own earlier writes, as Office
 * Scripts' getValue after setValue does), and (b) a trailing emit that
 * serialises the collected writes onto a sentinel-delimited console line. The
 * existing, unmodified {@link runInSandbox} runs that string in its opaque-origin iframe
 * Worker (no app cookies/storage/DOM); the host parses the writes back out of the
 * captured logs and applies them through the normal cell-write path (one undo
 * step, protection-aware, bounded). The script only ever sees the snapshot we
 * hand it — there is no live bridge and no new escape surface.
 *
 * This module is the pure, dependency-free, DOM-free core: build the harness,
 * parse the emitted writes, and validate/bound them. Exhaustively unit-testable.
 */

/** A populated cell in the snapshot handed to the script: `[row, col, value]`. */
export type SnapshotCell = readonly [number, number, string | number];

/** The active sheet snapshot passed into a user script. */
export interface SheetSnapshot {
  rowCount: number;
  columnCount: number;
  /** Only the populated cells, as `[row, col, value]` (numbers stay numbers). */
  cells: readonly SnapshotCell[];
}

/** A write the script asked for, post-validation: a string value at (row, col). */
export interface ScriptWrite {
  row: number;
  col: number;
  value: string;
}

/** The raw, untrusted write shape the sandboxed script emits. */
interface RawWrite {
  r: unknown;
  c: unknown;
  value: unknown;
}

/** The default ceiling on how many cells one script run may write. */
export const DEFAULT_MAX_SCRIPT_WRITES = 100_000;

/**
 * Sentinel that brackets the JSON writes payload on the script's final console
 * line. A private-use code point keeps it clear of ordinary script output, and
 * we take the LAST matching line, so a user `console.log` can't spoof the result.
 */
const WRITES_SENTINEL = "⁣MRAI_SCRIPT_WRITES⁣";

/**
 * Build the self-contained JS source to run in the sandbox: an inline `sheet`
 * object model over `snapshot`, then the user's `code`, then the writes emit.
 *
 * The harness is synchronous by design — the snapshot/writes model has no live
 * channel, so async has nothing to await; keeping it sync means a script that
 * never returns is killed cleanly by the sandbox timeout. A user syntax error or
 * thrown exception aborts before the emit line, so the host sees no writes and
 * applies nothing (atomic: a failed script never half-mutates the sheet).
 */
/**
 * Takes a worker global's network reach away: fetch, XMLHttpRequest,
 * WebSocket, EventSource, importScripts, nested workers, BroadcastChannel,
 * WebTransport and navigator.sendBeacon become undefined and cannot be put
 * back (non-writable, non-configurable). Source text, because it runs inside
 * the sandbox worker ahead of the user's script.
 */
export const SCRIPT_GLOBAL_SEAL_SOURCE = `(function (g) {
  if (!g) return;
  var blocked = ["fetch", "XMLHttpRequest", "WebSocket", "EventSource", "importScripts", "Worker", "SharedWorker", "BroadcastChannel", "WebTransport"];
  for (var i = 0; i < blocked.length; i++) {
    try { Object.defineProperty(g, blocked[i], { value: undefined, writable: false, configurable: false }); } catch (e) { try { g[blocked[i]] = undefined; } catch (e2) {} }
  }
  try { if (g.navigator) Object.defineProperty(g.navigator, "sendBeacon", { value: undefined, writable: false, configurable: false }); } catch (e) {}
})`;

export function buildScriptHarness(code: string, snapshot: SheetSnapshot): string {
  const data: Record<string, string | number> = {};
  for (const [r, c, value] of snapshot.cells) {
    if (Number.isInteger(r) && Number.isInteger(c)) data[`${r}_${c}`] = value;
  }
  const dataJson = JSON.stringify(data);
  const rows = Math.max(0, Math.trunc(snapshot.rowCount));
  const cols = Math.max(0, Math.trunc(snapshot.columnCount));
  const sentinelJson = JSON.stringify(WRITES_SENTINEL);
  // The user code is embedded as raw source (NOT a string) so it runs with normal
  // syntax. It sees only `sheet`; the snapshot + write buffer are closure-private.
  return `"use strict";
(function () {
  // The script works on the sheet it was handed and nothing else
  // (SCRIPT_GLOBAL_SEAL_SOURCE): the worker's network reach is taken away
  // before the script runs, so a worksheet's data cannot be sent anywhere.
  ${SCRIPT_GLOBAL_SEAL_SOURCE}(typeof self !== "undefined" ? self : globalThis);
  var __data = ${dataJson};
  var __rows = ${rows};
  var __cols = ${cols};
  var __writes = [];
  var __sentinel = ${sentinelJson};
  function __key(r, c) { return r + "_" + c; }
  // The run reads its own writes: each accepted setCell also updates the
  // private read view, in the shape the NEXT run's snapshot would show it
  // (numeric text as a number, TRUE/FALSE for booleans, null/"" clears).
  // Writes the host would drop (outside the grid, non-writable values) leave
  // the view alone, so reads never show a value that will not be committed.
  function __viewValue(value) {
    if (value === null || value === undefined) return null;
    if (typeof value === "number") return isFinite(value) ? value : undefined;
    if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
    if (typeof value !== "string") return undefined;
    if (value === "") return null;
    var trimmed = value.trim();
    var num = Number(trimmed);
    return trimmed !== "" && isFinite(num) && String(num) === trimmed ? num : value;
  }
  var sheet = Object.freeze({
    rowCount: __rows,
    columnCount: __cols,
    getCell: function (row, col) {
      var r = Math.trunc(row), c = Math.trunc(col);
      var v = __data[__key(r, c)];
      return v === undefined ? null : v;
    },
    setCell: function (row, col, value) {
      var r = Math.trunc(row), c = Math.trunc(col);
      __writes.push({ r: r, c: c, value: value });
      if (r >= 0 && c >= 0 && r < __rows && c < __cols) {
        var v = __viewValue(value);
        if (v === null) delete __data[__key(r, c)];
        else if (v !== undefined) __data[__key(r, c)] = v;
      }
      return value;
    },
    forEachCell: function (callback) {
      // Walk the cells as they stand when the walk starts; a cell the callback
      // clears is skipped, one it rewrites is read at its new value.
      var keys = Object.keys(__data);
      for (var i = 0; i < keys.length; i++) {
        var k = keys[i];
        if (!Object.prototype.hasOwnProperty.call(__data, k)) continue;
        var p = k.split("_");
        callback(__data[k], Number(p[0]), Number(p[1]));
      }
    },
  });
  // ----- user script -----
${code}
  // ----- end user script -----
  console.log(__sentinel + JSON.stringify(__writes) + __sentinel);
})();`;
}

/**
 * Recover the writes the script emitted from the sandbox's captured `logs`.
 * Returns `null` when no sentinel line is present — i.e. the script threw, timed
 * out, or never reached the emit — so the caller applies nothing.
 */
export function parseScriptWrites(logs: readonly string[]): RawWrite[] | null {
  // Scan from the end: the emit is the last line, and taking the last match means
  // a user console.log of the sentinel can't pre-empt the real payload.
  for (let i = logs.length - 1; i >= 0; i -= 1) {
    const line = logs[i];
    if (typeof line !== "string") continue;
    const first = line.indexOf(WRITES_SENTINEL);
    if (first === -1) continue;
    const start = first + WRITES_SENTINEL.length;
    const end = line.indexOf(WRITES_SENTINEL, start);
    if (end === -1) continue;
    const json = line.slice(start, end);
    try {
      const parsed = JSON.parse(json);
      return Array.isArray(parsed) ? (parsed as RawWrite[]) : [];
    } catch {
      return null;
    }
  }
  return null;
}

/** Coerce an emitted JS value to the string the cell-write path stores. */
function coerceValue(value: unknown): string | null {
  if (value === null || value === undefined) return ""; // clears the cell
  if (typeof value === "string") return value;
  if (typeof value === "number") return Number.isFinite(value) ? String(value) : null;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return null; // objects / arrays / functions are not writable values
}

/** The outcome of validating + bounding a script's raw writes. */
export interface NormalizedWrites {
  writes: ScriptWrite[];
  /** Raw writes rejected as out-of-bounds or non-coercible. */
  dropped: number;
  /** True when the write count exceeded `maxWrites` and was truncated. */
  truncated: boolean;
  /** The extent that holds every kept write (the stored one, grown). */
  rowCount: number;
  columnCount: number;
}

/**
 * Validate + bound the raw writes against the live grid: integer coordinates in
 * `[0, rowCount) × [0, columnCount)`, values coerced to a writable string, and
 * the total capped at `maxWrites`. Out-of-bounds / non-coercible writes are
 * dropped (counted), never clamped — a script can't reach outside the sheet.
 */
export function normalizeScriptWrites(
  raw: readonly RawWrite[],
  bounds: {
    rowCount: number;
    columnCount: number;
    maxWrites?: number;
    /**
     * The worksheet's own limits. Given, a write past the stored extent but
     * within them is kept and the extent grows to hold it, as a macro replay
     * grows it (Office Scripts write anywhere on the sheet); omitted, the
     * stored extent is the bound.
     */
    maxRows?: number;
    maxColumns?: number;
  },
): NormalizedWrites {
  const maxWrites = Math.max(0, bounds.maxWrites ?? DEFAULT_MAX_SCRIPT_WRITES);
  let rowCount = Math.max(0, Math.trunc(bounds.rowCount));
  let columnCount = Math.max(0, Math.trunc(bounds.columnCount));
  const rowLimit = Math.max(rowCount, Math.trunc(bounds.maxRows ?? rowCount));
  const columnLimit = Math.max(columnCount, Math.trunc(bounds.maxColumns ?? columnCount));
  const writes: ScriptWrite[] = [];
  let dropped = 0;
  let truncated = false;
  for (const w of raw) {
    if (writes.length >= maxWrites) {
      truncated = true;
      break;
    }
    const row = typeof w?.r === "number" ? w.r : NaN;
    const col = typeof w?.c === "number" ? w.c : NaN;
    if (
      !Number.isInteger(row) ||
      !Number.isInteger(col) ||
      row < 0 ||
      col < 0 ||
      row >= rowLimit ||
      col >= columnLimit
    ) {
      dropped += 1;
      continue;
    }
    const value = coerceValue(w.value);
    if (value === null) {
      dropped += 1;
      continue;
    }
    writes.push({ row, col, value });
    rowCount = Math.max(rowCount, row + 1);
    columnCount = Math.max(columnCount, col + 1);
  }
  return { writes, dropped, truncated, rowCount, columnCount };
}
