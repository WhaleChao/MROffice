/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { jsonEquals, type JsonValue } from "./json-crdt";

const JOURNAL = "pdfNativeTransactions";

export function pdfNativeJournal(content: Record<string, JsonValue>): Record<string, JsonValue>[] {
  const journal = content[JOURNAL] === undefined ? [] : content[JOURNAL];
  if (!Array.isArray(journal) || journal.some(record => !record || typeof record !== "object" ||
      Array.isArray(record) || typeof record.id !== "string" || !record.id)) throw new Error("Invalid PDF native journal");
  const records = journal as Record<string, JsonValue>[];
  if (new Set(records.map(record => record.id)).size !== records.length) throw new Error("Duplicate PDF native journal id");
  return records;
}

/** Native commands do not belong in the shared notes model. */
export function pdfSharedContent(content: Record<string, JsonValue>): Record<string, JsonValue> {
  pdfNativeJournal(content);
  return Object.fromEntries(Object.entries(content).filter(([key]) => key !== JOURNAL));
}

/** Ordinary notes carry the current server journal, never an obsolete local
 * command. Explicit native submission has its own single-writer envelope. */
export function withPdfDurableJournal(
  shared: Record<string, JsonValue>, durable: Record<string, JsonValue>,
): Record<string, JsonValue> {
  const content = pdfSharedContent(shared);
  pdfNativeJournal(durable);
  return Object.hasOwn(durable, JOURNAL)
    ? { ...content, [JOURNAL]: structuredClone(durable[JOURNAL]) } : content;
}

export function appendPdfCausalNativeCommand(
  source: Record<string, JsonValue>, requested: Record<string, JsonValue>, operationId: string,
): Record<string, JsonValue> {
  const before = pdfNativeJournal(source), after = pdfNativeJournal(requested);
  if (!operationId || after.length !== before.length + 1 || after.at(-1)?.id !== operationId ||
      !jsonEquals(before, after.slice(0, -1)) || !jsonEquals(pdfSharedContent(source), pdfSharedContent(requested))) {
    throw new Error("Native PDF command must append once to its exact saved source");
  }
  return structuredClone(requested);
}
