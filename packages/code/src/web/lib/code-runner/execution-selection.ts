/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface ExecutionRange {
  anchor: number;
  head: number;
}

export interface ExecutionLineMapping {
  /** Inclusive 1-based line bounds in the joined source sent to the runner. */
  executionLineFrom: number;
  executionLineTo: number;
  /** The original document line corresponding to `executionLineFrom`. */
  documentLineFrom: number;
}

export interface ExecutionSelection {
  /** A literal selection, or the caret's whole current line. */
  kind: "selection" | "line";
  /** The exact source sent through the existing sandbox runner. */
  source: string;
  /** Outer document bounds covered by the executable source. */
  from: number;
  to: number;
  /** One-based source line bounds for compact run-scope UI. */
  lineFrom: number;
  lineTo: number;
  /** Document characters selected, excluding separators added between ranges. */
  selectedChars: number;
  /** Number of disjoint executable ranges joined into `source`. */
  rangeCount: number;
  /**
   * Compact, sorted mappings for every selected range. Keeping one entry per
   * range (rather than one per source line) makes million-line selections
   * practical while still allowing logarithmic traceback navigation.
   *
   * Optional for compatibility with persisted/caller-created selections from
   * before multi-range traceback navigation shipped.
   */
  lineMappings?: readonly ExecutionLineMapping[];
}

/**
 * Translate a 1-based line from the exact source sent to a runner back to the
 * original editor document. A full-document run passes `null`. New selections
 * carry a compact per-range map, including disjoint multi-cursor selections.
 * Older one-range callers without that metadata retain the original offset
 * fallback. Synthetic separator lines deliberately return null.
 */
export function executionLineToDocumentLine(
  selection: ExecutionSelection | null,
  executionLine: number,
): number | null {
  if (!Number.isFinite(executionLine) || executionLine < 1) return null;
  const safeLine = Math.floor(executionLine);
  if (!selection) return safeLine;

  const mappings = selection.lineMappings;
  if (mappings && mappings.length > 0) {
    let low = 0;
    let high = mappings.length - 1;
    let candidate: ExecutionLineMapping | null = null;
    while (low <= high) {
      const middle = (low + high) >>> 1;
      const mapping = mappings[middle];
      if (mapping.executionLineFrom <= safeLine) {
        candidate = mapping;
        low = middle + 1;
      } else {
        high = middle - 1;
      }
    }
    if (!candidate || safeLine > candidate.executionLineTo) return null;
    return candidate.documentLineFrom + safeLine - candidate.executionLineFrom;
  }

  if (selection.rangeCount !== 1) return null;
  const documentLine = selection.lineFrom + safeLine - 1;
  return documentLine <= selection.lineTo ? documentLine : null;
}

interface NormalizedRange {
  from: number;
  to: number;
}

function clampOffset(value: number, length: number): number {
  if (!Number.isFinite(value)) return 0;
  return Math.min(Math.max(Math.trunc(value), 0), length);
}

function lineNumberAt(text: string, offset: number): number {
  let line = 1;
  for (let index = 0; index < offset; index += 1) {
    if (text.charCodeAt(index) === 10) line += 1;
  }
  return line;
}

function countNewlines(text: string, from: number, to: number): number {
  let count = 0;
  for (let index = from; index < to; index += 1) {
    if (text.charCodeAt(index) === 10) count += 1;
  }
  return count;
}

function normalizeRanges(
  textLength: number,
  ranges: readonly ExecutionRange[],
): NormalizedRange[] {
  const normalized = ranges
    .map(({ anchor, head }) => {
      const safeAnchor = clampOffset(anchor, textLength);
      const safeHead = clampOffset(head, textLength);
      return {
        from: Math.min(safeAnchor, safeHead),
        to: Math.max(safeAnchor, safeHead),
      };
    })
    .filter((range) => range.to > range.from)
    .sort((a, b) => (a.from === b.from ? a.to - b.to : a.from - b.from));

  const merged: NormalizedRange[] = [];
  for (const range of normalized) {
    const previous = merged.at(-1);
    if (previous && range.from <= previous.to) {
      previous.to = Math.max(previous.to, range.to);
    } else {
      merged.push({ ...range });
    }
  }
  return merged;
}

/**
 * Resolve the executable source behind VS Code/Jupyter-style "Run Selection /
 * Line" without mutating the editor document.
 *
 * Any non-empty ranges win and are executed in document order. Disjoint
 * multi-cursor selections are separated by one newline so tokens from adjacent
 * selections can never fuse accidentally. With only collapsed carets, the
 * first caret's complete current line is returned. Empty/whitespace-only input
 * is a no-op. All offsets are clamped so stale DOM selections cannot escape the
 * current buffer.
 */
export function buildExecutionSelection(
  text: string,
  ranges: readonly ExecutionRange[],
): ExecutionSelection | null {
  if (ranges.length === 0) return null;

  const normalized = normalizeRanges(text.length, ranges);
  if (normalized.length > 0) {
    const parts: string[] = [];
    const lineMappings: ExecutionLineMapping[] = [];
    let scanOffset = 0;
    let documentLine = 1;
    let executionLineFrom = 1;
    const advanceDocumentLineTo = (offset: number) => {
      while (scanOffset < offset) {
        if (text.charCodeAt(scanOffset) === 10) documentLine += 1;
        scanOffset += 1;
      }
      return documentLine;
    };

    for (const range of normalized) {
      const documentLineFrom = advanceDocumentLineTo(range.from);
      const documentLineTo = advanceDocumentLineTo(
        Math.max(range.from, range.to - 1),
      );
      const selectedLineCount = documentLineTo - documentLineFrom + 1;
      lineMappings.push({
        executionLineFrom,
        executionLineTo: executionLineFrom + selectedLineCount - 1,
        documentLineFrom,
      });
      parts.push(text.slice(range.from, range.to));

      // A range ending in "\n" leaves a synthetic empty execution line before
      // the joiner's separator. It is intentionally absent from the mapping.
      const sourceLineCount = countNewlines(text, range.from, range.to) + 1;
      executionLineFrom += sourceLineCount;
    }

    const source = parts.join("\n");
    if (source.trim().length === 0) return null;
    const first = normalized[0];
    const last = normalized[normalized.length - 1];
    const selectedChars = normalized.reduce(
      (total, range) => total + range.to - range.from,
      0,
    );
    return {
      kind: "selection",
      source,
      from: first.from,
      to: last.to,
      lineFrom: lineNumberAt(text, first.from),
      lineTo: lineNumberAt(text, Math.max(last.from, last.to - 1)),
      selectedChars,
      rangeCount: normalized.length,
      lineMappings,
    };
  }

  const caret = clampOffset(ranges[0].head, text.length);
  const from = text.lastIndexOf("\n", caret - 1) + 1;
  const nextBreak = text.indexOf("\n", caret);
  const to = nextBreak === -1 ? text.length : nextBreak;
  const source = text.slice(from, to);
  if (source.trim().length === 0) return null;
  const line = lineNumberAt(text, from);
  return {
    kind: "line",
    source,
    from,
    to,
    lineFrom: line,
    lineTo: line,
    selectedChars: source.length,
    rangeCount: 1,
    lineMappings: [
      {
        executionLineFrom: 1,
        executionLineTo: 1,
        documentLineFrom: line,
      },
    ],
  };
}
