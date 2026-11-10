/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for multi-buffer host cell-output fidelity (Jupyter parity):
 *   - stream classification, so stdout and stderr can render differently;
 *   - traceback parsing, so a raised error renders as a structured, styled block
 *     (a red header with the error type/message + the preserved traceback body)
 *     instead of one flat string.
 *
 * Everything here is synchronous and side-effect free so it can be unit-tested
 * without a real runtime and mirrored verbatim inside the runners' worker
 * scripts (JS Web Worker, Pyodide worker).
 *
 * Backward-compatibility contract:
 *   - `logs` stays a `string[]`; a parallel, OPTIONAL `logStreams` array tags
 *     each line's stream by index. A missing tag (legacy output, or fewer tags
 *     than lines) renders as stdout. `streamForLogIndex` encodes that default.
 */

import { stripAnsi } from "./ansi";

/** Which standard stream a captured log line came from. */
export type LogStream = "stdout" | "stderr";

/** JS `console` methods that map to stderr (Jupyter shows warnings as stderr). */
const STDERR_CONSOLE_METHODS = new Set(["error", "warn"]);

/**
 * Map a JS `console` method name to its stream. `error`/`warn` are stderr;
 * `log`/`info`/`debug` (and anything unknown) are stdout. Pure.
 */
export function classifyConsoleStream(method: string): LogStream {
  return STDERR_CONSOLE_METHODS.has(method) ? "stderr" : "stdout";
}

/**
 * Resolve the stream for the log line at `index`, defaulting to "stdout" when no
 * tag exists (legacy output, or the tag array is shorter than the log array).
 * The single place the "untagged ⇒ stdout" rule lives so it stays consistent
 * between the runners and the renderer. Pure.
 */
export function streamForLogIndex(
  streams: readonly LogStream[] | undefined,
  index: number,
): LogStream {
  const tag = streams?.[index];
  return tag === "stderr" ? "stderr" : "stdout";
}

/**
 * A single source location parsed out of a traceback (Jupyter/Colab "click a
 * frame to jump to the line" parity). `cell` is true when the frame points at
 * the executed cell itself (a synthetic Pyodide filename like `<exec>` or a JS
 * stack frame with no real file) rather than a separate file on disk — only
 * those are navigable in a single-cell notebook. Everything is data, never
 * markup; the renderer escapes it.
 */
export interface TracebackFrame {
  /**
   * The originating file as written in the traceback (e.g. `<exec>`,
   * `/path/to/mod.py`, `<anonymous>`). Undefined when the line carried only a
   * location (e.g. a bare JS `at <anonymous>:3:5`).
   */
  readonly file?: string;
  /** 1-based line number (always ≥ 1; parsed numerically and clamped). */
  readonly line: number;
  /**
   * Column when present (JS stacks), else undefined. JS stack frames report a
   * 0-based column, and column 0 is a real location (kept, not dropped).
   */
  readonly column?: number;
  /** Enclosing function/scope name when the frame names one (Python `in foo`). */
  readonly func?: string;
  /**
   * True when the frame refers to the executed cell itself (so the editor can
   * jump its caret there). False for a distinct on-disk file, which is shown but
   * not actionable in the common single-cell case.
   */
  readonly cell: boolean;
  /** The raw traceback line the frame was parsed from, for display. */
  readonly raw: string;
  /**
   * True for the runner's own machinery (Pyodide's `_pyodide/_base.py` that
   * evaluates the cell), not code the user wrote or imported.
   */
  readonly runtime?: boolean;
}

/** A parsed traceback: a one-line header plus the preserved multi-line body. */
export interface ParsedTraceback {
  /**
   * The error's headline — `"<Type>: <message>"` when both are present, else the
   * type or the trimmed first/last meaningful line. Suitable for the red header.
   */
  readonly header: string;
  /** The error type (class) name when detected, e.g. "NameError", "TypeError". */
  readonly errorType?: string;
  /** The error message (without the type prefix) when detected. */
  readonly message?: string;
  /**
   * Lines the runtime printed after the `Type: message` line (Pyodide's "The
   * module 'tkinter' is removed … See https://…"), when there are any.
   */
  readonly note?: string;
  /**
   * The full visible traceback text, with newlines intact and terminal control
   * sequences removed so ANSI styling bytes never leak into headers or frame
   * parsing. Equal to the visible input (trimmed of leading/trailing blank
   * lines) when there is no separable header.
   */
  readonly body: string;
  /**
   * Source locations extracted from the body, in traceback order (Python is
   * outermost→innermost; JS innermost→outermost — both preserved as printed).
   * OPTIONAL/additive: existing callers that read only `{header, body}` are
   * unaffected. Empty when no frame is recognised.
   */
  readonly frames: readonly TracebackFrame[];
}

/** An exception class name, as opposed to any `Word:` prefix. */
const ERROR_CLASS_RE = /(?:Error|Exception|Warning|Exit|Interrupt)$/;

/** At most this many plain lines may follow the `Type: message` summary. */
const TRAILING_NOTE_LINE_LIMIT = 6;

/** `Type: message` head, where Type is a CamelCase/dotted identifier. */
const ERROR_LINE_RE = /^([A-Za-z_][\w.]*(?:Error|Exception|Warning)?)\s*:\s*([\s\S]*)$/;

/**
 * Python frame: `File "<exec>", line 12, in <module>` (the `, in <func>` suffix
 * is optional; Pyodide/CPython quote the filename). The filename group is lazy
 * so an embedded quote in a path doesn't swallow the rest of the line.
 */
const PY_FRAME_RE =
  /^\s*File "(.*?)", line (\d+)(?:, in (.+))?\s*$/;

/**
 * JS stack frame: `at func (file:line:col)` or a bare `at file:line:col`. The
 * trailing `:line:col` is the location; an optional parenthesised group holds it
 * when a function name is present. Matches V8/Node/Chromium stacks.
 */
const JS_FRAME_RE =
  /^\s*at (?:(.*?) \()?(.*?):(\d+):(\d+)\)?\s*$/;

/**
 * Synthetic "this is the executed cell, not a real file" filenames produced by
 * the in-browser runtimes (Pyodide uses `<exec>`/`<string>`; notebooks use
 * `<cell>`/`<ipython-input-…>`; a JS Worker stack has `<anonymous>`/`eval`).
 * A frame whose file matches is navigable to the current cell's line.
 */
/** Pyodide's own evaluation frames (`/lib/python312.zip/_pyodide/_base.py`). */
const RUNTIME_FILE_RE = /(?:^\/lib\/python\d+\.zip\/|\/_pyodide\/)/;

const CELL_FILE_RE =
  /^<(?:exec|string|stdin|module|cell|ipython-input[^>]*|anonymous)>$/;

/**
 * Maximum source frames promoted to interactive traceback chips. The complete
 * traceback body is still preserved as one text node, but a pathological
 * recursive stack can never create thousands of buttons and freeze the output
 * workbench.
 */
export const MAX_TRACEBACK_FRAMES = 100;
const TRACEBACK_HEAD_FRAME_LIMIT = MAX_TRACEBACK_FRAMES / 2;
const TRACEBACK_TAIL_FRAME_LIMIT =
  MAX_TRACEBACK_FRAMES - TRACEBACK_HEAD_FRAME_LIMIT;

/** True when a traceback filename denotes the executed cell, not an on-disk file. */
function isCellFile(file: string | undefined): boolean {
  if (file === undefined) return true; // a bare `at :line:col` is in-cell.
  return CELL_FILE_RE.test(file.trim());
}

/**
 * Parse a 1-based LINE number, returning null for anything non-finite / ≤ 0.
 * Lines are always ≥ 1, so a `:0` line is junk and rejected.
 */
function parseLineNumber(raw: string): number | null {
  const n = Number.parseInt(raw, 10);
  return Number.isFinite(n) && n >= 1 ? n : null;
}

/**
 * Parse a COLUMN number, returning null for anything non-finite / < 0. Distinct
 * from {@link parseLineNumber}: JS stack frames legitimately report column 0
 * (e.g. `at <anonymous>:3:0`), so a 0 column is a real location, not junk, and
 * must be accepted rather than dropped.
 */
function parseColumnNumber(raw: string): number | null {
  const n = Number.parseInt(raw, 10);
  return Number.isFinite(n) && n >= 0 ? n : null;
}

/**
 * Extract source-location FRAMES from a traceback body, in printed order. Pure +
 * never throws: a malformed line that doesn't match a known frame shape is
 * skipped (not a frame), and a non-numeric / non-positive line number is
 * rejected so the renderer can't be handed a junk caret target.
 *
 * Recognises Python `File "…", line N, in func` and JS `at func (file:line:col)`
 * (and the bare `at file:line:col`) shapes. Exported for direct unit testing.
 */
export function parseTracebackFrames(body: string): TracebackFrame[] {
  // Keep both ends of an oversized traceback: V8 prints the most relevant
  // frame first, while Python prints the innermost/raised frame last. A bounded
  // rolling tail preserves both without allocating one object per recursive
  // call.
  const headFrames: TracebackFrame[] = [];
  const tailFrames: TracebackFrame[] = [];
  const collect = (frame: TracebackFrame) => {
    if (headFrames.length < TRACEBACK_HEAD_FRAME_LIMIT) {
      headFrames.push(frame);
      return;
    }
    tailFrames.push(frame);
    if (tailFrames.length > TRACEBACK_TAIL_FRAME_LIMIT) tailFrames.shift();
  };
  for (const raw of body.split("\n")) {
    const py = PY_FRAME_RE.exec(raw);
    if (py) {
      const line = parseLineNumber(py[2]);
      if (line === null) continue;
      const file = py[1];
      const func = py[3]?.trim() || undefined;
      collect({
        file,
        line,
        func,
        cell: isCellFile(file),
        raw: raw.trim(),
        ...(RUNTIME_FILE_RE.test(file) ? { runtime: true } : {}),
      });
      continue;
    }
    const js = JS_FRAME_RE.exec(raw);
    if (js) {
      const line = parseLineNumber(js[3]);
      if (line === null) continue;
      const column = parseColumnNumber(js[4]) ?? undefined;
      const func = js[1]?.trim() || undefined;
      // A function-less stack writes the location into the file slot
      // (`at file:line:col`); a named one parenthesises it. `<anonymous>` /
      // empty file ⇒ the cell.
      const fileRaw = js[2]?.trim();
      const file = fileRaw ? fileRaw : undefined;
      collect({
        file,
        line,
        column,
        func,
        cell: isCellFile(file),
        raw: raw.trim(),
      });
    }
  }
  return headFrames.concat(tailFrames);
}

function stripBlankEdges(lines: string[]): string[] {
  let start = 0;
  let end = lines.length;
  while (start < end && lines[start].trim() === "") start += 1;
  while (end > start && lines[end - 1].trim() === "") end -= 1;
  return lines.slice(start, end);
}

/**
 * Parse a captured error / traceback string into a `{header, errorType?,
 * message?, body}` shape for the renderer.
 *
 * Handles the common shapes without being a full parser (pure + cheap):
 *   - Python Pyodide tracebacks: a "Traceback (most recent call last):" block
 *     whose LAST non-empty line is "ErrorType: message" → that becomes the
 *     header; the whole block is the body.
 *   - A bare "ErrorType: message" (JS `name: message`, or a one-line Python
 *     error) → header is the line; body is the line.
 *   - Anything else → header is the first non-empty line; body is the text.
 *
 * Never throws; an empty/whitespace input yields an empty header and body. Pure.
 */
export function parseTraceback(text: string): ParsedTraceback {
  // Normalize line endings BEFORE terminal-control handling: ANSI's carriage
  // return semantics intentionally keep only the final progress update, while
  // the `\r` in a normal CRLF pair is just part of the newline.
  const normalized = stripAnsi((text ?? "").replace(/\r\n/g, "\n"));
  const lines = stripBlankEdges(normalized.split("\n"));
  if (lines.length === 0) {
    return { header: "", body: "", frames: [] };
  }
  const body = lines.join("\n");
  const frames = parseTracebackFrames(body);

  // A multi-line traceback: the summary "Type: message" is the last meaningful
  // line (Python prints it after the frames), except that a runtime may add a
  // few plain lines after it (Pyodide: "The module 'tkinter' is removed …").
  // Scan from the bottom across those and the indented frame block.
  let sawIndented = false;
  let trailing = 0;
  for (let i = lines.length - 1; i >= 0; i -= 1) {
    const candidate = lines[i].trim();
    if (candidate === "") continue;
    const indented = /^\s/.test(lines[i]);
    const match = indented ? null : ERROR_LINE_RE.exec(candidate);
    // Under trailing lines, only a real error class is the summary: a "See:
    // …" or "Note: …" line is part of the note.
    if (match && (trailing === 0 || sawIndented || ERROR_CLASS_RE.test(match[1]))) {
      const errorType = match[1];
      const message = match[2].trim();
      const header = message ? `${errorType}: ${message}` : errorType;
      const note = lines
        .slice(i + 1)
        .map((line) => line.trim())
        .filter(Boolean)
        .join("\n");
      return {
        header,
        errorType,
        message: message || undefined,
        ...(note ? { note } : {}),
        body,
        frames,
      };
    }
    if (indented) {
      sawIndented = true;
      continue;
    }
    // A plain line above the frames ends the search; below them it is one of
    // the runtime's trailing lines.
    if (sawIndented || ++trailing > TRAILING_NOTE_LINE_LIMIT) break;
  }

  // No "Type: message" summary found: use the first meaningful line as header.
  const header = lines.find((line) => line.trim() !== "")?.trim() ?? "";
  return { header, body, frames };
}
