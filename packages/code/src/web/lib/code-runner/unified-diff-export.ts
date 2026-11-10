/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Bounded unified-diff export for the Code Review comparison surface.
 *
 * The existing CodeMirror Myers implementation remains the diff authority. To
 * make it line-aware without adding another dependency, each distinct logical
 * line is encoded as one safe BMP code unit before calling `diff`. The returned
 * ranges therefore map directly back to source lines and can be emitted as
 * ordinary context hunks.
 */

import { diff, type Change } from "@codemirror/merge";

export const UNIFIED_DIFF_EXPORT_FILENAME = "code-comparison.patch";
export const MAX_UNIFIED_DIFF_SOURCE_BYTES = 2 * 1024 * 1024;
export const MAX_UNIFIED_DIFF_OUTPUT_BYTES = 4 * 1024 * 1024;
export const MAX_UNIFIED_DIFF_LINES = 100_000;
export const MAX_UNIFIED_DIFF_UNIQUE_LINES = 63_000;
export const MAX_UNIFIED_DIFF_CHANGES = 50_000;

const DEFAULT_CONTEXT_LINES = 3;
const MAX_CONTEXT_LINES = 10;
const DIFF_SCAN_LIMIT = 50_000;
const DIFF_TIMEOUT_MS = 250;
const FIRST_LINE_TOKEN = 0x0100;
const FIRST_TOKEN_RANGE_LENGTH = 0xd800 - FIRST_LINE_TOKEN;

export type UnifiedDiffExportFailureReason =
  | "identical"
  | "invalid_source"
  | "source_too_large"
  | "too_many_lines"
  | "too_many_changes"
  | "output_too_large"
  | "diff_failed";

export type UnifiedDiffExportReceipt =
  | {
      ok: true;
      patch: string;
      byteLength: number;
      hunkCount: number;
    }
  | {
      ok: false;
      reason: UnifiedDiffExportFailureReason;
      limit?: number;
    };

export interface UnifiedDiffExportInput {
  before: string;
  after: string;
  beforeLabel?: string;
  afterLabel?: string;
  contextLines?: number;
}

interface LogicalLine {
  body: string;
  terminated: boolean;
}

interface DiffHunk {
  fromA: number;
  toA: number;
  fromB: number;
  toB: number;
  changes: Change[];
}

function utf8ByteLength(value: string): number | null {
  try {
    if (typeof TextEncoder === "undefined") return null;
    return new TextEncoder().encode(value).byteLength;
  } catch {
    return null;
  }
}

function safeLabel(value: string | undefined, fallback: string): string {
  if (typeof value !== "string") return fallback;
  const leaf = value.trim().replaceAll("\\", "/").split("/").at(-1) ?? "";
  const bounded = leaf
    .replace(/[\u0000-\u001f\u007f]/g, "_")
    .replace(/[^A-Za-z0-9._-]+/g, "_")
    .replace(/^\.+/, "")
    .slice(0, 120);
  return bounded || fallback;
}

function contextLineCount(value: number | undefined): number {
  if (value === undefined || !Number.isFinite(value)) {
    return DEFAULT_CONTEXT_LINES;
  }
  return Math.max(0, Math.min(MAX_CONTEXT_LINES, Math.trunc(value)));
}

function splitLogicalLines(source: string): LogicalLine[] {
  if (source === "") return [];
  const lines: LogicalLine[] = [];
  let start = 0;
  while (start < source.length) {
    const newline = source.indexOf("\n", start);
    if (newline < 0) {
      lines.push({ body: source.slice(start), terminated: false });
      break;
    }
    lines.push({ body: source.slice(start, newline), terminated: true });
    start = newline + 1;
  }
  return lines;
}

function lineKey(line: LogicalLine): string {
  return `${line.terminated ? "1" : "0"}\u0000${line.body}`;
}

function lineToken(index: number): string | null {
  if (index < FIRST_TOKEN_RANGE_LENGTH) {
    return String.fromCharCode(FIRST_LINE_TOKEN + index);
  }
  const secondIndex = index - FIRST_TOKEN_RANGE_LENGTH;
  if (secondIndex >= 0x10000 - 0xe000) return null;
  return String.fromCharCode(0xe000 + secondIndex);
}

function encodeLines(
  before: readonly LogicalLine[],
  after: readonly LogicalLine[],
): { before: string; after: string } | null {
  const tokens = new Map<string, string>();
  const encode = (lines: readonly LogicalLine[]): string | null => {
    const encoded: string[] = [];
    for (const line of lines) {
      const key = lineKey(line);
      let token = tokens.get(key);
      if (token === undefined) {
        if (tokens.size >= MAX_UNIFIED_DIFF_UNIQUE_LINES) return null;
        token = lineToken(tokens.size) ?? undefined;
        if (token === undefined) return null;
        tokens.set(key, token);
      }
      encoded.push(token);
    }
    return encoded.join("");
  };
  const encodedBefore = encode(before);
  if (encodedBefore === null) return null;
  const encodedAfter = encode(after);
  return encodedAfter === null
    ? null
    : { before: encodedBefore, after: encodedAfter };
}

function validChange(
  change: Change,
  previous: Change | undefined,
  beforeLineCount: number,
  afterLineCount: number,
): boolean {
  return (
    Number.isSafeInteger(change.fromA) &&
    Number.isSafeInteger(change.toA) &&
    Number.isSafeInteger(change.fromB) &&
    Number.isSafeInteger(change.toB) &&
    change.fromA >= 0 &&
    change.fromA <= change.toA &&
    change.toA <= beforeLineCount &&
    change.fromB >= 0 &&
    change.fromB <= change.toB &&
    change.toB <= afterLineCount &&
    (!previous ||
      (change.fromA >= previous.toA && change.fromB >= previous.toB))
  );
}

function groupHunks(
  changes: readonly Change[],
  beforeLineCount: number,
  afterLineCount: number,
  contextLines: number,
): DiffHunk[] {
  const hunks: DiffHunk[] = [];
  for (const change of changes) {
    const fromA = Math.max(0, change.fromA - contextLines);
    const toA = Math.min(beforeLineCount, change.toA + contextLines);
    const fromB = Math.max(0, change.fromB - contextLines);
    const toB = Math.min(afterLineCount, change.toB + contextLines);
    const current = hunks.at(-1);
    if (current && fromA <= current.toA && fromB <= current.toB) {
      current.toA = Math.max(current.toA, toA);
      current.toB = Math.max(current.toB, toB);
      current.changes.push(change);
    } else {
      hunks.push({ fromA, toA, fromB, toB, changes: [change] });
    }
  }
  return hunks;
}

function rangeLabel(start: number, count: number): string {
  return count === 0 ? `${start},0` : `${start + 1},${count}`;
}

function appendPatchLine(
  parts: string[],
  prefix: " " | "+" | "-",
  line: LogicalLine,
): void {
  parts.push(prefix, line.body, "\n");
  if (!line.terminated) {
    parts.push("\\ No newline at end of file\n");
  }
}

function sameLine(a: LogicalLine, b: LogicalLine): boolean {
  return a.body === b.body && a.terminated === b.terminated;
}

function appendHunk(
  parts: string[],
  hunk: DiffHunk,
  before: readonly LogicalLine[],
  after: readonly LogicalLine[],
): boolean {
  parts.push(
    `@@ -${rangeLabel(hunk.fromA, hunk.toA - hunk.fromA)} +${rangeLabel(
      hunk.fromB,
      hunk.toB - hunk.fromB,
    )} @@\n`,
  );
  let cursorA = hunk.fromA;
  let cursorB = hunk.fromB;
  for (const change of hunk.changes) {
    const contextA = change.fromA - cursorA;
    const contextB = change.fromB - cursorB;
    if (contextA !== contextB || contextA < 0) return false;
    for (let index = 0; index < contextA; index += 1) {
      const lineA = before[cursorA + index];
      const lineB = after[cursorB + index];
      if (!lineA || !lineB || !sameLine(lineA, lineB)) return false;
      appendPatchLine(parts, " ", lineA);
    }
    for (let index = change.fromA; index < change.toA; index += 1) {
      const line = before[index];
      if (!line) return false;
      appendPatchLine(parts, "-", line);
    }
    for (let index = change.fromB; index < change.toB; index += 1) {
      const line = after[index];
      if (!line) return false;
      appendPatchLine(parts, "+", line);
    }
    cursorA = change.toA;
    cursorB = change.toB;
  }
  const trailingA = hunk.toA - cursorA;
  const trailingB = hunk.toB - cursorB;
  if (trailingA !== trailingB || trailingA < 0) return false;
  for (let index = 0; index < trailingA; index += 1) {
    const lineA = before[cursorA + index];
    const lineB = after[cursorB + index];
    if (!lineA || !lineB || !sameLine(lineA, lineB)) return false;
    appendPatchLine(parts, " ", lineA);
  }
  return true;
}

/** Create a complete, portable patch without mutating either compared source. */
export function createUnifiedDiffExport(
  input: UnifiedDiffExportInput,
): UnifiedDiffExportReceipt {
  try {
    if (
      !input ||
      typeof input.before !== "string" ||
      typeof input.after !== "string" ||
      input.before.includes("\u0000") ||
      input.after.includes("\u0000")
    ) {
      return { ok: false, reason: "invalid_source" };
    }
    if (input.before === input.after) {
      return { ok: false, reason: "identical" };
    }
    const beforeBytes = utf8ByteLength(input.before);
    const afterBytes = utf8ByteLength(input.after);
    if (
      beforeBytes === null ||
      afterBytes === null ||
      beforeBytes > MAX_UNIFIED_DIFF_SOURCE_BYTES ||
      afterBytes > MAX_UNIFIED_DIFF_SOURCE_BYTES
    ) {
      return {
        ok: false,
        reason: "source_too_large",
        limit: MAX_UNIFIED_DIFF_SOURCE_BYTES,
      };
    }

    const beforeLines = splitLogicalLines(input.before);
    const afterLines = splitLogicalLines(input.after);
    if (
      beforeLines.length > MAX_UNIFIED_DIFF_LINES ||
      afterLines.length > MAX_UNIFIED_DIFF_LINES
    ) {
      return {
        ok: false,
        reason: "too_many_lines",
        limit: MAX_UNIFIED_DIFF_LINES,
      };
    }
    const encoded = encodeLines(beforeLines, afterLines);
    if (!encoded) {
      return {
        ok: false,
        reason: "too_many_lines",
        limit: MAX_UNIFIED_DIFF_UNIQUE_LINES,
      };
    }
    const changes = diff(encoded.before, encoded.after, {
      scanLimit: DIFF_SCAN_LIMIT,
      timeout: DIFF_TIMEOUT_MS,
    });
    if (changes.length === 0) return { ok: false, reason: "identical" };
    if (changes.length > MAX_UNIFIED_DIFF_CHANGES) {
      return {
        ok: false,
        reason: "too_many_changes",
        limit: MAX_UNIFIED_DIFF_CHANGES,
      };
    }
    for (let index = 0; index < changes.length; index += 1) {
      if (
        !validChange(
          changes[index]!,
          changes[index - 1],
          beforeLines.length,
          afterLines.length,
        )
      ) {
        return { ok: false, reason: "diff_failed" };
      }
    }

    const hunks = groupHunks(
      changes,
      beforeLines.length,
      afterLines.length,
      contextLineCount(input.contextLines),
    );
    const parts = [
      `--- a/${safeLabel(input.beforeLabel, "current.code")}\n`,
      `+++ b/${safeLabel(input.afterLabel, "incoming.code")}\n`,
    ];
    for (const hunk of hunks) {
      if (!appendHunk(parts, hunk, beforeLines, afterLines)) {
        return { ok: false, reason: "diff_failed" };
      }
    }
    const patch = parts.join("");
    const byteLength = utf8ByteLength(patch);
    if (byteLength === null || byteLength > MAX_UNIFIED_DIFF_OUTPUT_BYTES) {
      return {
        ok: false,
        reason: "output_too_large",
        limit: MAX_UNIFIED_DIFF_OUTPUT_BYTES,
      };
    }
    return { ok: true, patch, byteLength, hunkCount: hunks.length };
  } catch {
    return { ok: false, reason: "diff_failed" };
  }
}
