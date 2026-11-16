/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import { isLosslessDurableNativeSaveReceipt, readNativeSaveReceiptSnapshot } from "@mrnote/core/web/lib/native-save-receipt.ts";
import type { PageItem } from "@mrnote/core/web/lib/notebook-sdk.ts";
import { pageStoredJsonEquals } from "@mrnote/core/web/lib/page-save-acknowledgement.ts";
import { isMraiNativeBytesPublishRecord } from "./engine/native/bytes-publish-record";
import type { PdfSharedStateRecord } from "./engine/shared-state-bridge";

/** The most revisions one chain carries (the server's bound). */
export const PDF_NATIVE_CHAIN_MAX = 32;

/** The failure of a chain that wrote nothing and may go as ordinary saves:
 * the queue sends its edits one at a time instead. */
export const PDF_NATIVE_CHAIN_UNAVAILABLE = "pdf-native-chain-unavailable";






export const PDF_NATIVE_CHAIN_SAVE_SWITCH = "mrai.pdf.chainSave";
export function pdfNativeChainSaveEnabled(): boolean {
  try {
    return typeof window === "undefined" || window.localStorage.getItem(PDF_NATIVE_CHAIN_SAVE_SWITCH) !== "off";
  } catch {
    return true;
  }
}

export function isPdfNativeChainUnavailable(error: unknown): boolean {
  return error instanceof Error && error.message === PDF_NATIVE_CHAIN_UNAVAILABLE;
}

type Row = Record<string, unknown>;
const row = (value: unknown): Row | null =>
  value && typeof value === "object" && !Array.isArray(value) ? value as Row : null;
const journalOf = (value: unknown): unknown[] | null =>
  value === undefined || value === null ? [] : Array.isArray(value) ? value : null;

/**
 * How many of the records a chain appended the saved journal holds: the
 * journal before it entry for entry, then the first j records sent, the
 * first exactly as sent and each later one as sent but for the stored
 * version its source names (the server's; it must name one) -- and each
 * standing on the bytes the one before it made. 0 when the saved journal is
 * anything else.
 */
export function pdfNativeChainAcknowledged(saved: unknown, before: unknown, submitted: unknown): number {
  const prior = journalOf(before), sent = journalOf(submitted), held = journalOf(saved);
  if (!prior || !sent || !held || sent.length < prior.length + 2 ||
    held.length <= prior.length || held.length > sent.length) return 0;
  try {
    if (!pageStoredJsonEquals(sent.slice(0, prior.length), prior) ||
      !pageStoredJsonEquals(held.slice(0, prior.length), prior)) return 0;
    for (let index = prior.length; index < held.length; index += 1) {
      const stored = row(held[index]), mine = row(sent[index]);
      if (!stored || !mine || !isMraiNativeBytesPublishRecord(stored) || !isMraiNativeBytesPublishRecord(mine)) return 0;
      if (index === prior.length) {
        if (!pageStoredJsonEquals(stored, mine)) return 0;
        continue;
      }
      const storedSource = row(stored.source), mineSource = row(mine.source);
      const previous = held[index - 1] as { derivedCandidate?: { sha256?: unknown } };
      if (!storedSource || !mineSource || typeof storedSource.resourceVersionId !== "string" ||
        !storedSource.resourceVersionId.trim() || storedSource.sha256 !== previous.derivedCandidate?.sha256 ||
        !pageStoredJsonEquals({ ...stored, source: { ...storedSource, resourceVersionId: mineSource.resourceVersionId } },
          mine)) return 0;
    }
  } catch {
    return 0;
  }
  return held.length - prior.length;
}

export interface PdfNativeChainStep {
  /** The record as the server stored it. */
  readonly record: PdfSharedStateRecord;
  /** The page as it stood once this record was saved. */
  readonly page: PageItem;
}

type Receipt = Record<string, unknown>;

/**
 * A chain's answer read as the saves it was: for each record the saved page
 * holds, that record and the page as it stood after it (its journal ending
 * there, under that save's own receipt). `page` is the save's answer, or the
 * page read again when the answer was lost.
 *
 * Every receipt must be a lossless durable one naming that record's own
 * bytes. The last is the page's own. An earlier one is the one the server
 * reported for it (`pdf_native_chain.steps`), which must agree with the
 * journal -- the next record's source is the version it made -- or, with no
 * report (the page read again), the receipt that pin states. Null when the
 * page is not that.
 */
export function readPdfNativeChainAnswer(input: Readonly<{
  page: PageItem | null | undefined; pageId: string; before: unknown; submitted: unknown;
}>): readonly PdfNativeChainStep[] | null {
  const page = input.page;
  const content = row(page?.content_json);
  if (!page || page.id !== input.pageId || !content) return null;
  const journal = journalOf(content.pdfNativeTransactions);
  const prior = journalOf(input.before);
  const written = pdfNativeChainAcknowledged(journal, input.before, input.submitted);
  if (!journal || !prior || written < 1) return null;
  const records = journal.slice(prior.length) as PdfSharedStateRecord[];
  const last = row((page as unknown as Row).editable_page_patchback);
  const version = Number(page.content_version);
  if (!last || !Number.isSafeInteger(version) || version < written) return null;
  const report = row((page as unknown as Row).pdf_native_chain);
  const reported = report && Array.isArray(report.steps) ? report.steps.map(row) : null;
  if (report && (!reported || reported.length !== written || report.written !== written)) return null;
  const steps: PdfNativeChainStep[] = [];
  let previousVersionId: string | null = null;
  for (let index = 0; index < written; index += 1) {
    const record = records[index] as unknown as {
      id: string; source: { resourceId?: unknown; resourceVersionId?: unknown };
      derivedCandidate: { sha256: string };
    };
    const next = records[index + 1] as unknown as { source?: { resourceVersionId?: unknown } } | undefined;
    const step = reported?.[index] ?? null;
    if (reported && (!step || step.id !== record.id)) return null;
    let receipt: Receipt | null;
    if (index === written - 1) {
      receipt = last;
      const mine = row(step?.editable_page_patchback);
      if (step && (!mine || mine.resource_version_id !== last.resource_version_id ||
        mine.output_sha256 !== last.output_sha256)) return null;
    } else if (step) {
      receipt = row(step.editable_page_patchback);
    } else {
      receipt = { status: "succeeded", adapter: "pdf", resource_id: last.resource_id,
        resource_version_id: next?.source?.resourceVersionId, output_sha256: record.derivedCandidate.sha256,
        patched_anchor_count: 1, loss_entry_count: 0 };
    }
    // The version this save made is the one the next record was saved on.
    if (!receipt || (next && receipt.resource_version_id !== next.source?.resourceVersionId)) return null;
    const stepVersion = step && Number.isSafeInteger(step.content_version)
      ? Number(step.content_version) : version - (written - 1 - index);
    const { pdf_native_chain: _report, ...rest } = page as unknown as Row;
    void _report;
    const stepPage = {
      ...rest,
      content_version: stepVersion,
      content_json: { ...content, pdfNativeTransactions: journal.slice(0, prior.length + index + 1) as JsonValue },
      editable_page_patchback: receipt,
      ...(index === written - 1 ? {} : { editable_page_patchback_download: null, editable_page_patchback_source: null }),
    } as unknown as PageItem;
    if (!isLosslessDurableNativeSaveReceipt(readNativeSaveReceiptSnapshot(stepPage), {
      expectedAdapter: "pdf",
      ...(typeof record.source.resourceId === "string" ? { expectedResourceId: record.source.resourceId } : {}),
      expectedOutputSha256: record.derivedCandidate.sha256,
      ...(previousVersionId ? { previousResourceVersionId: previousVersionId } : {}),
    })) return null;
    previousVersionId = String(receipt.resource_version_id);
    steps.push(Object.freeze({ record: records[index], page: stepPage }));
  }
  return Object.freeze(steps);
}

/**
 * The queue's leading edits the server's page already holds: a chain goes on
 * being saved after the window that sent it has gone (a reload mid-save), so
 * the page read later may hold several of the queue's records where a single
 * save in flight left at most one. For each of the queue's ids found in the
 * journal, in order from the first and each standing on the bytes the one
 * before it made, the record and the page as it stood after it: its receipt
 * is the one the next record's source states (the version it was saved on is
 * the version this one made), or the page's own when it is the journal's
 * last. Null when the journal does not hold the queue's first edit so.
 */
export function readPdfNativeLandedRecords(input: Readonly<{
  page: PageItem | null | undefined; pageId: string; operationIds: readonly string[];
}>): readonly PdfNativeChainStep[] | null {
  const page = input.page;
  const content = row(page?.content_json);
  const journal = content ? journalOf(content.pdfNativeTransactions) : null;
  if (!page || page.id !== input.pageId || !content || !journal || !input.operationIds.length) return null;
  const start = journal.findIndex((entry) => row(entry)?.id === input.operationIds[0]);
  const last = row((page as unknown as Row).editable_page_patchback);
  const version = Number(page.content_version);
  if (start < 0 || !last) return null;
  const { pdf_native_chain: _report, ...rest } = page as unknown as Row;
  void _report;
  const steps: PdfNativeChainStep[] = [];
  let previousVersionId: string | null = null;
  for (let index = 0; index < input.operationIds.length && start + index < journal.length; index += 1) {
    const at = start + index;
    const record = row(journal[at]);
    if (!record || record.id !== input.operationIds[index] || !isMraiNativeBytesPublishRecord(record)) break;
    const source = row(record.source);
    const before = index ? journal[at - 1] as { derivedCandidate?: { sha256?: unknown } } : null;
    if (!source || (before && source.sha256 !== before.derivedCandidate?.sha256)) break;
    const output = record.derivedCandidate.sha256;
    let receipt: Receipt;
    if (at === journal.length - 1) receipt = last;
    else {
      const nextSource = row(row(journal[at + 1])?.source);
      if (!nextSource || nextSource.sha256 !== output || typeof nextSource.resourceVersionId !== "string") break;
      receipt = { status: "succeeded", adapter: "pdf", resource_id: last.resource_id,
        resource_version_id: nextSource.resourceVersionId, output_sha256: output,
        patched_anchor_count: 1, loss_entry_count: 0 };
    }
    const stepPage = {
      ...rest,
      ...(Number.isSafeInteger(version) ? { content_version: Math.max(1, version - (journal.length - 1 - at)) } : {}),
      content_json: { ...content, pdfNativeTransactions: journal.slice(0, at + 1) as JsonValue },
      editable_page_patchback: receipt,
      ...(at === journal.length - 1 ? {} : { editable_page_patchback_download: null, editable_page_patchback_source: null }),
    } as unknown as PageItem;
    if (!isLosslessDurableNativeSaveReceipt(readNativeSaveReceiptSnapshot(stepPage), {
      expectedAdapter: "pdf",
      ...(typeof source.resourceId === "string" ? { expectedResourceId: source.resourceId } : {}),
      expectedOutputSha256: output,
      ...(previousVersionId ? { previousResourceVersionId: previousVersionId } : {}),
    })) break;
    previousVersionId = String(receipt.resource_version_id);
    steps.push(Object.freeze({ record: record as unknown as PdfSharedStateRecord, page: stepPage }));
  }
  return steps.length ? Object.freeze(steps) : null;
}
