/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const MAX_DEBUG_SNAPSHOT_EXPRESSION_CHARS = 500;
export const MAX_DEBUG_SNAPSHOT_VALUE_CHARS = 5_000;
export const MAX_DEBUG_SNAPSHOT_HISTORY_ENTRIES = 20;
export const MAX_DEBUG_SNAPSHOT_HISTORY_BYTES = 32 * 1024;

export type DebugSnapshotEvaluationStatus = "success" | "error";

export interface DebugSnapshotEvaluationReceipt {
  readonly status: DebugSnapshotEvaluationStatus;
  readonly value: string;
}

export interface DebugSnapshotEvaluationEntry
  extends DebugSnapshotEvaluationReceipt {
  readonly id: number;
  readonly expression: string;
}

export function normalizeDebugSnapshotExpression(
  value: unknown,
): string | null {
  if (typeof value !== "string") return null;
  const expression = value.trim();
  if (
    expression.length === 0 ||
    expression.length > MAX_DEBUG_SNAPSHOT_EXPRESSION_CHARS ||
    /[\0\r\n\u2028\u2029]/u.test(expression)
  ) {
    return null;
  }
  return expression;
}

function boundedEntry(
  entry: DebugSnapshotEvaluationEntry,
): DebugSnapshotEvaluationEntry | null {
  const expression = normalizeDebugSnapshotExpression(entry.expression);
  if (
    !expression ||
    !Number.isSafeInteger(entry.id) ||
    entry.id < 1 ||
    (entry.status !== "success" && entry.status !== "error") ||
    typeof entry.value !== "string"
  ) {
    return null;
  }
  return {
    id: entry.id,
    expression,
    status: entry.status,
    value:
      entry.value.length <= MAX_DEBUG_SNAPSHOT_VALUE_CHARS
        ? entry.value
        : `${entry.value.slice(0, MAX_DEBUG_SNAPSHOT_VALUE_CHARS - 1)}…`,
  };
}

function entryBytes(entry: DebugSnapshotEvaluationEntry): number {
  return new TextEncoder().encode(
    `${entry.id}\0${entry.expression}\0${entry.status}\0${entry.value}`,
  ).byteLength;
}

/**
 * Append one trusted paused-frame evaluation to the newest-first, session-only
 * console. History is deliberately not persisted: expressions and local values
 * can contain secrets even when the durable debug timeline is safe to reopen.
 */
export function appendDebugSnapshotEvaluation(
  current: readonly DebugSnapshotEvaluationEntry[],
  candidate: DebugSnapshotEvaluationEntry,
): readonly DebugSnapshotEvaluationEntry[] {
  const next = boundedEntry(candidate);
  if (!next) return current;

  const retained: DebugSnapshotEvaluationEntry[] = [];
  let bytes = 0;
  for (const entry of [next, ...current]) {
    const bounded = boundedEntry(entry);
    if (!bounded) continue;
    const size = entryBytes(bounded);
    if (size > MAX_DEBUG_SNAPSHOT_HISTORY_BYTES) continue;
    if (
      retained.length >= MAX_DEBUG_SNAPSHOT_HISTORY_ENTRIES ||
      bytes + size > MAX_DEBUG_SNAPSHOT_HISTORY_BYTES
    ) {
      break;
    }
    retained.push(bounded);
    bytes += size;
  }
  return retained;
}
