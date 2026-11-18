/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type CodeRuntimeGuardKind = "gui" | "terminal" | "process";

export interface CodeRuntimeGuardFinding {
  /** The module as the source names it (`tkinter`, `PyQt5.QtWidgets` ...). */
  readonly module: string;
  readonly kind: CodeRuntimeGuardKind;
  /** 1-based line of the import. */
  readonly line: number;
}

/** Top-level modules that can never run in the code page's Python. */
const PYTHON_BLOCKED_MODULES: ReadonlyMap<string, CodeRuntimeGuardKind> = new Map([
  ["tkinter", "gui"],
  ["Tkinter", "gui"],
  ["_tkinter", "gui"],
  ["PyQt4", "gui"],
  ["PyQt5", "gui"],
  ["PyQt6", "gui"],
  ["PySide", "gui"],
  ["PySide2", "gui"],
  ["PySide6", "gui"],
  ["wx", "gui"],
  ["kivy", "gui"],
  ["pyglet", "gui"],
  ["arcade", "gui"],
  ["pygame_gui", "gui"],
  ["curses", "terminal"],
  ["_curses", "terminal"],
  ["subprocess", "process"],
]);

/**
 * ``source`` with every comment and string literal blanked to spaces (line
 * breaks kept, so line numbers stay true). Handles single, double and
 * triple quotes, string prefixes (r, b, f, u and their pairs), backslash
 * escapes and line continuations.
 */
export function pythonCodeOnly(source: string): string {
  let out = "";
  let index = 0;
  const length = source.length;
  const blank = (text: string) => text.replace(/[^\n]/g, " ");
  while (index < length) {
    const char = source[index];
    if (char === "#") {
      const end = source.indexOf("\n", index);
      const stop = end === -1 ? length : end;
      out += blank(source.slice(index, stop));
      index = stop;
      continue;
    }
    if (char === "'" || char === '"') {
      // A prefix (r, b, f, u ...) was copied as identifier letters, which is
      // harmless. A backslash escapes the next character in every string,
      // raw ones included (r"\"" does not end at the escaped quote).
      const triple = source.startsWith(char.repeat(3), index);
      const quote = triple ? char.repeat(3) : char;
      let cursor = index + quote.length;
      while (cursor < length) {
        const at = source[cursor];
        if (at === "\\") {
          cursor += 2;
          continue;
        }
        if (!triple && at === "\n") break;
        if (source.startsWith(quote, cursor)) {
          cursor += quote.length;
          break;
        }
        cursor += 1;
      }
      const stop = Math.min(cursor, length);
      out += blank(source.slice(index, stop));
      index = stop;
      continue;
    }
    out += char;
    index += 1;
  }
  return out;
}

/** The modules each import statement of ``code`` (comments and strings
 * already blanked) names, with their 1-based lines. */
function pythonImports(code: string): Array<{ module: string; line: number }> {
  const found: Array<{ module: string; line: number }> = [];
  // Logical lines: a backslash continuation or an open parenthesis joins
  // physical lines (`from x import (\n a,\n b)` names x once).
  const physical = code.split("\n");
  let depth = 0;
  let logical = "";
  let startLine = 1;
  const flush = () => {
    for (const statement of logical.split(";")) {
      const text = statement.trim();
      const from = /^from\s+([A-Za-z_][\w.]*)\s+import\b/.exec(text);
      if (from) {
        found.push({ module: from[1], line: startLine });
        continue;
      }
      const plain = /^import\s+(.+)$/.exec(text);
      if (plain) {
        for (const part of plain[1].split(",")) {
          const name = /^\s*([A-Za-z_][\w.]*)/.exec(part);
          if (name) found.push({ module: name[1], line: startLine });
        }
      }
    }
    logical = "";
  };
  physical.forEach((rawLine, position) => {
    if (logical === "") startLine = position + 1;
    const continued = /\\\s*$/.test(rawLine);
    const line = continued ? rawLine.replace(/\\\s*$/, " ") : rawLine;
    for (const char of line) {
      if (char === "(" || char === "[" || char === "{") depth += 1;
      else if ((char === ")" || char === "]" || char === "}") && depth > 0) depth -= 1;
    }
    logical += `${line} `;
    if (!continued && depth === 0) flush();
  });
  if (logical.trim()) flush();
  // Statements nested in blocks (`if x:\n    import tkinter`) and after a
  // colon on one line (`try: import tkinter`) count too.
  return found.concat(
    physical.flatMap((line, position) => {
      const inline = /:\s*(?:import\s+([A-Za-z_][\w.]*)|from\s+([A-Za-z_][\w.]*)\s+import\b)/.exec(line);
      const named = inline?.[1] ?? inline?.[2];
      return named ? [{ module: named, line: position + 1 }] : [];
    }),
  );
}

/**
 * Why ``source`` cannot run on a code page of ``language``: each import that
 * needs a window, a terminal or a process the browser run does not have.
 * Empty when nothing in it is certain to fail (only Python is checked).
 */
export function codeRuntimeGuardFindings(
  language: string,
  source: string,
): CodeRuntimeGuardFinding[] {
  if (language !== "python" || !source) return [];
  const code = pythonCodeOnly(source);
  const findings: CodeRuntimeGuardFinding[] = [];
  const seen = new Set<string>();
  for (const { module, line } of pythonImports(code)) {
    const top = module.split(".")[0];
    const kind = PYTHON_BLOCKED_MODULES.get(top);
    if (!kind || seen.has(top)) continue;
    seen.add(top);
    findings.push({ module, kind, line });
  }
  return findings.sort((a, b) => a.line - b.line);
}
