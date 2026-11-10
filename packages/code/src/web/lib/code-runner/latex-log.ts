/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * TeX's log (main.log, returned with every LaTeX receipt by
 * lib/code-runner/output-files) read into the entries a member acts on, the
 * way Overleaf's "Logs and output files" lists them: errors, warnings, and
 * the overfull and underfull boxes of typesetting, each with the file and
 * line it came from.
 *
 * The log names a file when TeX opens it, as "(" followed by its path, and
 * closes it with the matching ")". Which file a message belongs to is the
 * innermost file still open when it is printed, so the reader keeps that
 * stack. Text TeX echoes from the document (an error's context lines, a bad
 * box's contents) can hold parentheses of its own, so the lines that belong to
 * a recognised message are never read for file changes.
 *
 * The compile passes -file-line-error (errors read `./main.tex:12: ...`) and
 * max_print_line=1000 (lines are rarely wrapped); LaTeX's own errors still
 * print as `! LaTeX Error: ...` with the line in a later `l.12` line.
 */

import { compiledProjectPath } from "./latex-project";

export type LatexLogLevel = "error" | "warning" | "typesetting";

export interface LatexLogEntry {
  readonly level: LatexLogLevel;
  readonly message: string;
  /**
   * The file it came from: a project path ("chapters/intro.tex", "main.tex")
   * when `inProject`, else the path as the log names it (a TeX Live file);
   * null when the log does not say.
   */
  readonly file: string | null;
  readonly inProject: boolean;
  readonly line: number | null;
  /** What TeX printed after an error (its `l.12 ...` context); else "". */
  readonly context: string;
}

export interface LatexLogCounts {
  readonly error: number;
  readonly warning: number;
  readonly typesetting: number;
}

export const MAX_LATEX_LOG_ENTRIES = 500;
const MAX_MESSAGE_CHARS = 2_000;
/** How far after an error its `l.N` line is looked for. */
const MAX_CONTEXT_SEARCH_LINES = 16;
const MAX_CONTINUATION_LINES = 12;
/**
 * Lines after an error or a bad box whose parentheses are TeX echoing the
 * document; the echo ends at the next blank line.
 */
const MAX_ECHO_LINES = 40;

const FILE_LINE_ERROR =
  /^((?:\.{1,2}\/|\/)(?:[^:\s]*\/)?[^:\s/]+|[A-Za-z0-9_][^:\s/]*\.[A-Za-z0-9]+):(\d+): (.*)$/;
const WARNING = /^((?:LaTeX\w*|Package|Class|Module)(?: [^\s:]+)?) Warning: (.*)$/;
const ENGINE_WARNING = /^(?:(?:pdf|Lua|Xe|e-)?TeX|x?dvipdfmx)(?: \([^)]*\))?:? ?warning\b/i;
const BAD_BOX =
  /^(?:Over|Under)full \\[hv]box \([^)]*\)(?: (?:in paragraph|in alignment) at lines (\d+)--\d+| detected at line (\d+)| has occurred while \\output is active)?/;
const CONTEXT_LINE = /^l\.(\d+)/;
const INPUT_LINE = /\s*on input line (\d+)\.?\s*$/;
const FATAL = /^==> Fatal error occurred/;
const EMERGENCY_STOP = /^Emergency stop\.?$/;
// A path TeX opens: relative to the compile's folder or absolute.
const OPENED_PATH = /^(?:\.{1,2}\/|\/)[^\s(){}<>[\]"]*/;

/** The log's text (TeX writes bytes; a cut log may start mid-character). */
export function decodeLatexLog(bytes: Uint8Array): string {
  return new TextDecoder("utf-8", { fatal: false }).decode(bytes);
}

function bounded(message: string): string {
  const visible = message
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return visible.length <= MAX_MESSAGE_CHARS
    ? visible
    : `${visible.slice(0, MAX_MESSAGE_CHARS - 1)}…`;
}

function located(path: string | null): Pick<LatexLogEntry, "file" | "inProject"> {
  if (!path) return { file: null, inProject: false };
  const projectPath = compiledProjectPath(path);
  return projectPath
    ? { file: projectPath, inProject: true }
    : { file: path, inProject: false };
}

/** The log's entries: errors first, then warnings, then typesetting. */
export function parseLatexLog(text: string): LatexLogEntry[] {
  const lines = text.split(/\r?\n/);
  // The files TeX has open, innermost last; null for a parenthesis that is
  // not a file, so its ")" closes nothing.
  const open: (string | null)[] = [];
  const currentFile = (): string | null => {
    for (let index = open.length - 1; index >= 0; index -= 1) {
      const file = open[index];
      if (file) return file;
    }
    return null;
  };
  const errors: LatexLogEntry[] = [];
  const warnings: LatexLogEntry[] = [];
  const boxes: LatexLogEntry[] = [];
  const stops: LatexLogEntry[] = [];
  const add = (list: LatexLogEntry[], entry: LatexLogEntry) => {
    if (errors.length + warnings.length + boxes.length < MAX_LATEX_LOG_ENTRIES) {
      list.push(entry);
    }
  };
  let echo = 0;

  const readFileChanges = (line: string) => {
    for (let index = 0; index < line.length; index += 1) {
      const char = line[index];
      if (char === "(") {
        const opened = OPENED_PATH.exec(line.slice(index + 1));
        if (opened && /[^./]/.test(opened[0])) {
          open.push(opened[0]);
          index += opened[0].length;
        } else {
          open.push(null);
        }
      } else if (char === ")") {
        open.pop();
      }
    }
  };

  // An error's `l.N` line and the one after it (the rest of the source line).
  // LaTeX's own errors print "Emergency stop" before it; that line is passed.
  const errorContext = (start: number): { line: number | null; text: string } => {
    const last = Math.min(lines.length, start + MAX_CONTEXT_SEARCH_LINES);
    for (let index = start; index < last; index += 1) {
      const line = lines[index] ?? "";
      if (line.startsWith("! ")) break;
      const stop = FILE_LINE_ERROR.exec(line);
      if (stop && !EMERGENCY_STOP.test((stop[3] ?? "").trim())) break;
      const at = CONTEXT_LINE.exec(line);
      if (!at) continue;
      const context = [line.trimEnd()];
      const next = lines[index + 1];
      if (next !== undefined && next.trim() !== "" && !next.startsWith("! ")) {
        context.push(next.trimEnd());
      }
      return { line: Number(at[1]), text: context.join("\n") };
    }
    return { line: null, text: "" };
  };

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index] ?? "";
    if (line.trim() === "") {
      echo = 0;
      continue;
    }

    const fileLine = line.startsWith("! ") ? null : FILE_LINE_ERROR.exec(line);
    if (fileLine || line.startsWith("! ")) {
      const message = fileLine ? (fileLine[3] ?? "").trim() : line.slice(2).trim();
      echo = MAX_ECHO_LINES;
      if (FATAL.test(message) || EMERGENCY_STOP.test(message)) {
        stops.push({
          level: "error",
          message: bounded(message.replace(/^==>\s*/, "")),
          ...located(fileLine ? (fileLine[1] ?? null) : currentFile()),
          line: fileLine ? Number(fileLine[2]) : null,
          context: "",
        });
        continue;
      }
      const context = errorContext(index + 1);
      add(errors, {
        level: "error",
        message: bounded(message),
        ...located(fileLine ? (fileLine[1] ?? null) : currentFile()),
        line: fileLine ? Number(fileLine[2]) : context.line,
        context: context.text,
      });
      continue;
    }

    if (echo > 0) {
      echo -= 1;
      continue;
    }

    const warning = WARNING.exec(line);
    if (warning) {
      const source = warning[1] ?? "";
      // Continuation lines start with "(hyperref)", "(Font)", "(LaTeX)".
      const prefix = `(${source.split(" ").at(-1) ?? source})`;
      const parts = [warning[2] ?? ""];
      while (parts.length <= MAX_CONTINUATION_LINES) {
        const next = lines[index + 1];
        if (next === undefined || next.trim() === "") break;
        if (next.startsWith(prefix)) parts.push(next.slice(prefix.length));
        else if (/^\s/.test(next) && !/^\s*\(/.test(next)) parts.push(next);
        else break;
        index += 1;
      }
      const joined = bounded(parts.join(" "));
      const inputLine = INPUT_LINE.exec(joined);
      const text = inputLine ? joined.slice(0, inputLine.index).trim() : joined;
      add(warnings, {
        level: "warning",
        // LaTeX's own warnings read as Overleaf lists them; a package or
        // class keeps its name.
        message: source === "LaTeX" ? text : `${source} Warning: ${text}`,
        ...located(currentFile()),
        line: inputLine ? Number(inputLine[1]) : null,
        context: "",
      });
      continue;
    }

    if (ENGINE_WARNING.test(line)) {
      add(warnings, {
        level: "warning",
        message: bounded(line),
        ...located(currentFile()),
        line: null,
        context: "",
      });
      continue;
    }

    const box = BAD_BOX.exec(line);
    if (box) {
      const at = box[1] ?? box[2];
      add(boxes, {
        level: "typesetting",
        message: bounded(line.replace(/\s*\[\]\s*$/, "")),
        ...located(currentFile()),
        line: at ? Number(at) : null,
        context: "",
      });
      // The box's contents follow, up to a blank line.
      echo = MAX_ECHO_LINES;
      continue;
    }

    readFileChanges(line);
  }

  // The stop TeX reports after an error adds nothing to it; alone it is the
  // only word on why the compile ended.
  return [...errors, ...(errors.length > 0 ? [] : stops.slice(0, 1)), ...warnings, ...boxes];
}

export function countLatexLogEntries(entries: readonly LatexLogEntry[]): LatexLogCounts {
  let error = 0;
  let warning = 0;
  let typesetting = 0;
  for (const entry of entries) {
    if (entry.level === "error") error += 1;
    else if (entry.level === "warning") warning += 1;
    else typesetting += 1;
  }
  return { error, warning, typesetting };
}
