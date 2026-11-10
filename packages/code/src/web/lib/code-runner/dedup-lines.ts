/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure "Remove Duplicate Lines" transform (VSCode parity — the
 * "editor.action.removeDuplicateLines" / "...removeAllExceptFirst" commands).
 *
 * Two modes:
 *  - `"all"`        — keep the FIRST occurrence of each distinct line and drop
 *                     every later duplicate, anywhere in the text (VSCode "Delete
 *                     Duplicate Lines"). Order of the surviving lines is the order
 *                     of their first appearance.
 *  - `"consecutive"`— collapse only RUNS of identical adjacent lines to a single
 *                     line (Unix `uniq`), leaving non-adjacent duplicates intact.
 *
 * Framework-free string→string so it is unit-testable in node and reused by
 * CodeMirrorEditor over the selected line block (or the whole document when the
 * selection is empty). The line COUNT shrinks by exactly the number of removed
 * duplicates; the relative order of kept lines is preserved.
 *
 * CRLF handling: the transform splits on "\n" only, so a "\r" from a literal
 * CRLF buffer rides along as the LAST char of each segment. Two CRLF lines with
 * identical content therefore compare equal (both carry the trailing "\r") and
 * dedupe correctly, and a kept line retains its "\r" so the document's line
 * endings survive. A document with no duplicates (or an editor that never runs
 * this) is returned byte-for-byte unchanged.
 */

/** Which duplicates to remove. */
export type DedupMode = "all" | "consecutive";

/**
 * Remove duplicate lines from `text` per {@link DedupMode}.
 *
 * Pure + total: empty input maps to empty output, a single-line buffer (no
 * "\n") is returned unchanged, and it never throws. Comparison is exact
 * (case-sensitive, whitespace-significant) — only byte-identical lines are
 * treated as duplicates, matching VSCode.
 *
 * @param text The text (a selected line block, or the whole document).
 * @param mode `"all"` (default) drops every non-first duplicate; `"consecutive"`
 *             collapses only adjacent runs.
 * @returns The de-duplicated text, lines re-joined with "\n".
 */
export function removeDuplicateLines(text: string, mode: DedupMode = "all"): string {
  if (!text) return text;
  const lines = text.split("\n");
  if (lines.length < 2) return text;

  if (mode === "consecutive") {
    const out: string[] = [];
    for (const line of lines) {
      if (out.length === 0 || out[out.length - 1] !== line) {
        out.push(line);
      }
    }
    return out.join("\n");
  }

  // "all": keep first occurrence, drop later exact duplicates.
  const seen = new Set<string>();
  const out: string[] = [];
  for (const line of lines) {
    if (seen.has(line)) continue;
    seen.add(line);
    out.push(line);
  }
  return out.join("\n");
}
