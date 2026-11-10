/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeLanguage } from "./language-metadata";
import type { LatexLogEntry } from "./latex-log";

export const MAX_COMPILER_PROBLEMS = 100;
export const MAX_COMPILER_PROBLEM_MESSAGE_CHARS = 2_048;
const MAX_COMPILER_ERROR_TEXT_CHARS = 300_000;

export interface CompilerProblem {
  from: number;
  to: number;
  severity: "error";
  message: string;
  line: number;
  col: number;
  navigable: true;
}

const SOURCE_FILE_BY_LANGUAGE: Readonly<Record<string, string>> = {
  c: "main.c",
  cpp: "main.cpp",
  go: "main.go",
  java: "Main.java",
  latex: "main.tex",
  perl: "main.pl",
  php: "main.php",
  r: "main.R",
  rust: "main.rs",
  swift: "main.swift",
};

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function workspacePathPattern(filename: string): string {
  return `(?:file:\\/\\/)?\\/?workspace\\/${escapeRegExp(filename)}`;
}

function boundedMessage(value: string): string {
  const visible = value
    .replace(/\u001b\[[0-?]*[ -/]*[@-~]/g, "")
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, " ")
    .trim();
  const message = visible || "Compilation failed";
  return message.length <= MAX_COMPILER_PROBLEM_MESSAGE_CHARS
    ? message
    : `${message.slice(0, MAX_COMPILER_PROBLEM_MESSAGE_CHARS - 1)}…`;
}

interface LocatedCompilerError {
  line: number;
  col: number | null;
  message: string;
}

function directLocations(
  language: string,
  filename: string,
  lines: readonly string[],
): LocatedCompilerError[] {
  const location = new RegExp(
    `^\\s*(?:-->|note:)?\\s*${workspacePathPattern(filename)}:(\\d+):(\\d+)(?::\\s*(.*))?\\s*$`,
  );
  const found: LocatedCompilerError[] = [];
  for (let index = 0; index < lines.length; index += 1) {
    const match = location.exec(lines[index] ?? "");
    if (!match) continue;
    let message = match[3] ?? "";
    if (language === "rust") {
      for (let previous = index - 1; previous >= 0; previous -= 1) {
        const header = /^\s*error(?:\[[^\]]+\])?:\s*(.+)$/.exec(
          lines[previous] ?? "",
        );
        if (header) {
          message = header[1] ?? message;
          break;
        }
        if ((lines[previous] ?? "").trim() === "") break;
      }
    }
    message = message.replace(/^(?:fatal\s+)?error(?:\[[^\]]+\])?:\s*/i, "");
    found.push({
      line: Number(match[1]),
      col: Number(match[2]),
      message,
    });
  }
  return found;
}

function javaLocations(
  filename: string,
  lines: readonly string[],
): LocatedCompilerError[] {
  const location = new RegExp(
    `^\\s*${workspacePathPattern(filename)}:(\\d+):\\s*error:\\s*(.*)$`,
  );
  const found: LocatedCompilerError[] = [];
  for (let index = 0; index < lines.length; index += 1) {
    const match = location.exec(lines[index] ?? "");
    if (!match) continue;
    const caret = (lines[index + 2] ?? "").indexOf("^");
    found.push({
      line: Number(match[1]),
      col: caret >= 0 ? caret + 1 : null,
      message: match[2] ?? "",
    });
  }
  return found;
}

function lineOnlyLocations(
  language: string,
  filename: string,
  lines: readonly string[],
): LocatedCompilerError[] {
  const path = workspacePathPattern(filename);
  const perl = new RegExp(`^\\s*(.+?)\\s+at\\s+${path}\\s+line\\s+(\\d+)(?:[,.].*)?$`);
  const php = new RegExp(
    `^\\s*(?:(?:PHP\\s+)?(?:Parse|Fatal) error):\\s*(.+?)\\s+in\\s+${path}\\s+on\\s+line\\s+(\\d+)\\s*$`,
    "i",
  );
  const found: LocatedCompilerError[] = [];
  for (const line of lines) {
    const match = (language === "perl" ? perl : php).exec(line);
    if (!match) continue;
    found.push({ line: Number(match[2]), col: null, message: match[1] ?? "" });
  }
  return found;
}

/**
 * TeX's `-file-line-error` form, `./main.tex:12: Undefined control sequence.`:
 * a line and no column. The context lines TeX prints after it (`l.12 ...`)
 * stay Output text.
 */
function texLocations(
  filename: string,
  lines: readonly string[],
): LocatedCompilerError[] {
  const location = new RegExp(
    `^\\s*(?:\\.\\/|(?:file:\\/\\/)?\\/?workspace\\/)?${escapeRegExp(filename)}:(\\d+):\\s*(.*)$`,
  );
  const found: LocatedCompilerError[] = [];
  for (const line of lines) {
    const match = location.exec(line);
    if (!match) continue;
    found.push({ line: Number(match[1]), col: null, message: match[2] ?? "" });
  }
  return found;
}

interface SourceLineRange {
  line: number;
  start: number;
  end: number;
  textLength: number;
}

function sourceLineRange(source: string, requestedLine: number): SourceLineRange {
  const starts = [0];
  for (let index = 0; index < source.length; index += 1) {
    if (source.charCodeAt(index) === 10) starts.push(index + 1);
  }
  const safeLine = Math.min(
    starts.length,
    Math.max(1, Number.isSafeInteger(requestedLine) ? requestedLine : 1),
  );
  const start = starts[safeLine - 1] ?? 0;
  const rawEnd =
    safeLine < starts.length ? (starts[safeLine] ?? source.length) - 1 : source.length;
  const end = rawEnd > start && source.charCodeAt(rawEnd - 1) === 13
    ? rawEnd - 1
    : rawEnd;
  return { line: safeLine, start, end, textLength: Math.max(0, end - start) };
}

function projectLocation(
  source: string,
  location: LocatedCompilerError,
): CompilerProblem {
  const range = sourceLineRange(source, location.line);
  if (location.col === null || !Number.isSafeInteger(location.col)) {
    return {
      from: range.start,
      to: range.end,
      severity: "error",
      message: boundedMessage(location.message),
      line: range.line,
      col: 1,
      navigable: true,
    };
  }
  const columnOffset = Math.min(
    range.textLength,
    Math.max(0, location.col - 1),
  );
  const from = range.start + columnOffset;
  return {
    from,
    to: from < range.end ? from + 1 : from,
    severity: "error",
    message: boundedMessage(location.message),
    line: range.line,
    col: columnOffset + 1,
    navigable: true,
  };
}

function locatedProblems(
  source: string,
  located: readonly LocatedCompilerError[],
): CompilerProblem[] {
  const problems: CompilerProblem[] = [];
  const seen = new Set<string>();
  for (const location of located) {
    if (problems.length >= MAX_COMPILER_PROBLEMS) break;
    if (!Number.isSafeInteger(location.line) || location.line < 1) continue;
    const problem = projectLocation(source, location);
    const key = `${problem.line}:${problem.col}:${problem.message}`;
    if (seen.has(key)) continue;
    seen.add(key);
    problems.push(problem);
  }
  return problems.sort((left, right) => left.line - right.line || left.col - right.col);
}

/**
 * The errors TeX's log (lib/code-runner/latex-log) places in one file of a
 * LaTeX project, as problems of that file's `source`. The log knows every
 * error with its file, where the compile's stdout lists only main.tex's.
 */
export function compilerProblemsFromLatexLog(
  source: string,
  entries: readonly LatexLogEntry[],
  projectPath: string,
): CompilerProblem[] {
  if (!source) return [];
  return locatedProblems(
    source,
    entries
      .filter(
        (entry) =>
          entry.level === "error" &&
          entry.inProject &&
          entry.file === projectPath &&
          entry.line !== null,
      )
      .map((entry) => ({ line: entry.line!, col: null, message: entry.message })),
  );
}

/**
 * Project only diagnostics emitted for the closed, pinned single-file compiler
 * plans. Foreign paths and unknown languages stay ordinary Output text.
 */
export function compilerProblemsForExecution(
  language: string,
  source: string,
  compilerError: string,
): CompilerProblem[] {
  const normalized = normalizeLanguage(language);
  const filename = SOURCE_FILE_BY_LANGUAGE[normalized];
  if (!filename || !source || !compilerError) return [];
  const lines = compilerError.slice(0, MAX_COMPILER_ERROR_TEXT_CHARS).split(/\r?\n/);
  const located =
    normalized === "java"
      ? javaLocations(filename, lines)
      : normalized === "perl" || normalized === "php"
        ? lineOnlyLocations(normalized, filename, lines)
        : normalized === "latex"
          ? texLocations(filename, lines)
          : directLocations(normalized, filename, lines);
  return locatedProblems(source, located);
}
