/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Which breakpoints this source can actually pause at.
 *
 * The JavaScript debugger pauses by inserting an AWAITED hook at safe statement
 * boundaries, so it cannot pause inside a non-async function body — `await` is
 * not legal there. A breakpoint on such a line is not an error; it simply never
 * fires. Saying nothing about it is the problem: the user presses Debug and
 * watches the program run straight past the line they marked.
 *
 * One place answers the question so the breakpoint list and the editor gutter
 * can never disagree about which dots are armed.
 */

import { instrumentForDebug } from "@mrnote/core/web/lib/code-runner/js-debug.ts";
import { transpileTypeScript } from "./ts-runner";

/**
 * Sources longer than this are not judged. Claiming a breakpoint is dead when
 * it is not would be worse than saying nothing, and this runs on every edit.
 */
export const MAX_BINDING_ANALYSIS_CHARACTERS = 256 * 1024;

/**
 * The subset of `breakpointLines` the debugger cannot pause at.
 *
 * Returns an empty list whenever the answer is unknown — an unsupported
 * language (Python traces real frames and can stop anywhere), a source too
 * large to analyse, or one the instrumenter could not read. An empty result
 * always means "nothing to report", never "all of them are fine".
 */
export function unbindableBreakpointLines(
  source: string,
  breakpointLines: readonly number[],
  language: string,
): number[] {
  if (breakpointLines.length === 0) return [];
  if (language !== "javascript" && language !== "typescript") return [];
  if (source.length > MAX_BINDING_ANALYSIS_CHARACTERS) return [];
  try {
    // TypeScript is judged on the JavaScript that will actually run. The
    // transpile preserves line numbers (its module prelude is one line, joined
    // without a newline), so a line number means the same thing on both sides.
    const javascript =
      language === "typescript"
        ? (transpileTypeScript(source).code ?? null)
        : source;
    if (javascript === null) return [];
    // Ask where hooks COULD go, not where they went for this breakpoint set:
    // instrumenting for a single unbindable breakpoint inserts nothing at all,
    // which is indistinguishable from "could not read the source".
    const { instrumented, hookedLines } = instrumentForDebug(javascript);
    if (!instrumented) return [];
    const bindable = new Set(hookedLines);
    return breakpointLines.filter((line) => !bindable.has(line));
  } catch {
    return [];
  }
}
