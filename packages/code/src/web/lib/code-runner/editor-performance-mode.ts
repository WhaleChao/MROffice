/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Bounded workload classification for the local CodeMirror surface.
 *
 * The source-file import ceiling is 2 MiB, but semantic diagnostics and
 * whole-document chrome can still turn a valid import into a sluggish editor.
 * This classifier deliberately inspects at most the large-source threshold:
 * source length is O(1), and sources below that bound are scanned once for
 * pathological line counts / very long lines.
 */

export const LARGE_SOURCE_CHARACTER_THRESHOLD = 512 * 1024;
export const LARGE_SOURCE_LINE_THRESHOLD = 20_000;
export const VERY_LONG_LINE_CHARACTER_THRESHOLD = 20_000;

export type EditorPerformanceModeReason =
  "source-size" | "line-count" | "long-line";

export interface EditorPerformanceProfile {
  mode: "full" | "lightweight";
  reason: EditorPerformanceModeReason | null;
  characterCount: number;
  /** Characters inspected after the O(1) source-length check. */
  scannedCharacterCount: number;
  /** Exact only for sources that stayed below every early-exit threshold. */
  lineCount: number | null;
  /** Exact only for sources that stayed below every early-exit threshold. */
  longestLineCharacterCount: number | null;
}

/**
 * Select the editor workload profile without ever scanning an unbounded source.
 *
 * Lightweight mode is automatic and reversible. It is intentionally a chrome /
 * background-work policy rather than a read-only mode: editing, search/replace,
 * undo/redo, multi-cursor, folding, save and sandboxed execution remain active.
 */
export function editorPerformanceProfile(
  source: string,
): EditorPerformanceProfile {
  const characterCount = source.length;
  if (characterCount >= LARGE_SOURCE_CHARACTER_THRESHOLD) {
    return {
      mode: "lightweight",
      reason: "source-size",
      characterCount,
      scannedCharacterCount: 0,
      lineCount: null,
      longestLineCharacterCount: null,
    };
  }

  let lineCount = 1;
  let currentLineCharacterCount = 0;
  let longestLineCharacterCount = 0;

  for (let index = 0; index < characterCount; index += 1) {
    if (source.charCodeAt(index) === 10) {
      longestLineCharacterCount = Math.max(
        longestLineCharacterCount,
        currentLineCharacterCount,
      );
      currentLineCharacterCount = 0;
      lineCount += 1;
      if (lineCount >= LARGE_SOURCE_LINE_THRESHOLD) {
        return {
          mode: "lightweight",
          reason: "line-count",
          characterCount,
          scannedCharacterCount: index + 1,
          lineCount,
          longestLineCharacterCount,
        };
      }
      continue;
    }

    currentLineCharacterCount += 1;
    if (currentLineCharacterCount >= VERY_LONG_LINE_CHARACTER_THRESHOLD) {
      return {
        mode: "lightweight",
        reason: "long-line",
        characterCount,
        scannedCharacterCount: index + 1,
        lineCount,
        longestLineCharacterCount: currentLineCharacterCount,
      };
    }
  }

  longestLineCharacterCount = Math.max(
    longestLineCharacterCount,
    currentLineCharacterCount,
  );
  return {
    mode: "full",
    reason: null,
    characterCount,
    scannedCharacterCount: characterCount,
    lineCount,
    longestLineCharacterCount,
  };
}
