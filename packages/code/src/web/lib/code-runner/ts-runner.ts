/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * TypeScript execution: strip the types with sucrase (fast, no type-checking)
 * and run the resulting JavaScript through the existing isolated JS runner.
 * Type errors are not reported (this is "run my code", not "tsc"); a syntax
 * error surfaces as a compile error.
 */

import { transform } from "sucrase";

import {
  AWAIT_COMPLETION_SOURCE_PREFIX,
  isAwaitCompletionJavaScript,
  runJavaScript,
  type RunOptions,
  type RunResult,
} from "@mrnote/core/web/lib/code-runner/js-runner.ts";

export interface TranspileResult {
  code?: string;
  error?: string;
  /** Characters the transpile put before the person's first line. */
  firstLineOffset?: number;
}

/**
 * Sucrase's `imports` transform lowers TypeScript ESM exports to CommonJS.
 * The execution worker is deliberately not a module loader, so provide only
 * the inert export objects that single-file snippets need. `require` always
 * fails closed: authored code cannot resolve browser or host modules through
 * this compatibility layer.
 *
 * ONE LINE, and joined to the source WITHOUT a newline, because everything
 * downstream counts lines: a stack trace names the line that threw, and the
 * debugger instruments the line a breakpoint is on. Sucrase preserves line
 * numbers when it strips types, so a multi-line prelude was the only thing
 * standing between the user's line 3 and the runtime's line 3 — it reported
 * errors five lines off, and a TypeScript breakpoint paused inside this
 * prelude instead of the line the user marked. It ends in `;` so the source
 * that follows on the same line can never be glued onto it.
 */
const TYPESCRIPT_MODULE_PRELUDE =
  'var exports = Object.create(null); var module = { exports: exports }; ' +
  'var require = function (specifier) { throw new Error("Module imports are unavailable in the MRNote TypeScript sandbox: " + String(specifier)); }; ';

const STRICT_DIRECTIVE = '"use strict";';

/**
 * The prelude shares line 1 with the person's code, so a frame on line 1
 * names a column that far to the right ("<anonymous>:1:245" for a call in
 * column 13). Give line-1 frames back the person's column.
 */
function placeFirstLineColumns(result: RunResult, offset: number): RunResult {
  if (typeof result.error !== "string" || offset <= 0) return result;
  const error = result.error.replace(
    /<anonymous>:1:(\d+)/g,
    (frame, column: string) => {
      const authored = Number(column) - offset;
      return authored >= 1 ? `<anonymous>:1:${authored}` : frame;
    },
  );
  return error === result.error ? result : { ...result, error };
}

/** Strip TypeScript types to runnable JavaScript. Pure + unit-testable. */
export function transpileTypeScript(source: string): TranspileResult {
  try {
    // `imports` removes ESM export syntax that classic Worker/new Function
    // cannot parse. JSX stays intentionally disabled (there is no React in the
    // sandbox), so angle-bracket casts (<T>x) keep working.
    const { code } = transform(source, {
      transforms: ["typescript", "imports"],
      disableESTransforms: true,
    });
    // The `imports` transform also opens line 1 with its own directive.
    const directive =
      code.startsWith(STRICT_DIRECTIVE) && !source.startsWith(STRICT_DIRECTIVE)
        ? STRICT_DIRECTIVE.length
        : 0;
    return {
      code: `${TYPESCRIPT_MODULE_PRELUDE}${code}`,
      firstLineOffset: TYPESCRIPT_MODULE_PRELUDE.length + directive,
    };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return { error: `TypeScript compile error: ${message}` };
  }
}

export async function runTypeScript(
  code: string,
  options: RunOptions = {},
): Promise<RunResult> {
  const awaitCompletion = isAwaitCompletionJavaScript(code);
  const transpiled = transpileTypeScript(code);
  if (transpiled.error || transpiled.code === undefined) {
    return { logs: [], error: transpiled.error ?? "TypeScript compile error" };
  }
  const marked = awaitCompletion && !isAwaitCompletionJavaScript(transpiled.code);
  const executable = marked
    ? `${AWAIT_COMPLETION_SOURCE_PREFIX}${transpiled.code}`
    : transpiled.code;
  return placeFirstLineColumns(
    await runJavaScript(executable, options),
    (transpiled.firstLineOffset ?? 0) +
      (marked ? AWAIT_COMPLETION_SOURCE_PREFIX.length : 0),
  );
}
