/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Finish & Merge pipeline.
 *
 * Every requested recipient becomes one ordinary editable DOCX. The merge
 * itself (MERGEFIELD atoms and the recipient store materialized/removed in
 * fresh copies; the live template is never edited) happens here; each merged
 * copy's Word file is written by the server's Word writer, the records sent
 * in bounded batches (`requestServerDocxBatch`: at most 25 documents and
 * about 8 MB of JSON per request, one request after another). The files are
 * bundled into one ZIP so browsers do not block a burst of downloads.
 *
 * A merged copy keeps the main document's tracked changes, as Word's Finish &
 * Merge does (every merged record holds the revisions the main document still
 * has; Word settles none of them): the writer keeps them (its batch default),
 * and the manifest says so.
 */

import {
  MAX_MAIL_MERGE_FINISH_RECORDS,
  finalizeMergedDocument,
  normalizeMailMergeData,
  resolveMailMergeSettings,
  selectMailMergeRecords,
  type MailMergeData,
  type MailMergeRecord,
} from "./mailMerge";
import { canonicalDocumentJson } from "../collab/paged-document-projection";
import {
  docxFilename,
  isDocxExportNeedsConnectionError,
  requestServerDocxBatch,
} from "../document-export/docx-server-export";
import type { TipTapNode } from "@/lib/document-export/tiptap-json";

/** Asks the server's Word writer for the merged copies' files, in order. */
export type MailMergeDocxWriter = (
  documents: ReadonlyArray<{ document: unknown; title?: string }>,
  options: { signal?: AbortSignal },
) => Promise<Blob[]>;

export interface MailMergeBatchFile {
  recordNumber: number;
  filename: string;
}

export interface MailMergeBatchResult {
  /** Null means the exact request resolved to zero output; do not download. */
  blob: Blob | null;
  filename: string;
  files: MailMergeBatchFile[];
}

export const MAX_MAIL_MERGE_BATCH_FILES = MAX_MAIL_MERGE_FINISH_RECORDS;

export class MailMergeBatchExportError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "MailMergeBatchExportError";
  }
}

export function trustedMailMergeDownloadErrorMessage(
  error: unknown,
  fallback: string,
  downloadUnavailable: string,
  needsConnection: string = fallback,
): string {
  // The Word files are written by the server: without a connection there
  // are none, and the person is told to reconnect.
  if (isDocxExportNeedsConnectionError(error)) {
    return needsConnection;
  }
  return error instanceof Error && error.message === downloadUnavailable
    ? downloadUnavailable
    : fallback;
}

function safeStem(title: string | undefined): string {
  return docxFilename(title, "mail-merge").replace(/\.docx$/i, "");
}

export function mailMergeRecordFilename(
  title: string | undefined,
  record: MailMergeRecord,
  total: number,
): string {
  const width = Math.max(3, String(Math.max(total, record.number)).length);
  return `${safeStem(title)}-${String(record.number).padStart(width, "0")}.docx`;
}

export function mailMergeBatchFilename(title: string | undefined): string {
  return `${safeStem(title)}-merged-docx.zip`;
}

/**
 * Resolve an explicit list exactly. A malformed, duplicate, filtered-out or
 * out-of-range record number produces zero files instead of silently widening
 * or partially fulfilling the request.
 */
function selectRequestedRecords(
  data: MailMergeData,
  recordNumbers?: readonly number[],
): MailMergeRecord[] {
  const selected = selectMailMergeRecords(data);
  if (recordNumbers === undefined) return selected;
  if (
    recordNumbers.length === 0 ||
    recordNumbers.some(
      (recordNumber) =>
        !Number.isInteger(recordNumber) || recordNumber <= 0,
    ) ||
    new Set(recordNumbers).size !== recordNumbers.length
  ) {
    return [];
  }
  const recordsByNumber = new Map(
    selected.map((record) => [record.number, record] as const),
  );
  const records = recordNumbers.map((recordNumber) =>
    recordsByNumber.get(recordNumber),
  );
  return records.every(
    (record): record is MailMergeRecord => record !== undefined,
  )
    ? records
    : [];
}

async function blobBytes(blob: Blob): Promise<Uint8Array> {
  if (typeof blob.arrayBuffer === "function") {
    return new Uint8Array(await blob.arrayBuffer());
  }
  if (typeof FileReader !== "undefined") {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () =>
        resolve(new Uint8Array(reader.result as ArrayBuffer));
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(blob);
    });
  }
  throw new MailMergeBatchExportError(
    "DOCX files can't be packaged here.",
  );
}

function emptyBatch(title: string | undefined): MailMergeBatchResult {
  return {
    blob: null,
    filename: mailMergeBatchFilename(title),
    files: [],
  };
}

export async function buildMailMergeDocxBatch(
  doc: TipTapNode,
  dataCandidate: MailMergeData,
  title?: string,
  recordNumbers?: readonly number[],
  options: { writeDocx?: MailMergeDocxWriter; signal?: AbortSignal } = {},
): Promise<MailMergeBatchResult> {
  const data = normalizeMailMergeData(dataCandidate);
  const records = selectRequestedRecords(data, recordNumbers);
  if (records.length === 0) return emptyBatch(title);
  if (records.length > MAX_MAIL_MERGE_BATCH_FILES) {
    throw new MailMergeBatchExportError(
      `A local batch can contain at most ${MAX_MAIL_MERGE_BATCH_FILES} editable DOCX files. Choose a smaller From/To range and retry.`,
    );
  }

  const { default: JSZip } = await import("jszip");
  const zip = new JSZip();
  const files: MailMergeBatchFile[] = [];
  // The storage-shaped body a save would send, merged once per record.
  const template = canonicalDocumentJson(doc);
  const merged = records.map((record) => ({
    document: finalizeMergedDocument(template, record) as TipTapNode,
    title,
  }));
  const writeDocx = options.writeDocx ?? requestServerDocxBatch;
  const blobs = await writeDocx(merged, { signal: options.signal });
  if (blobs.length !== records.length) {
    throw new MailMergeBatchExportError(
      "The Word writer answered for a different number of records.",
    );
  }
  for (const [index, record] of records.entries()) {
    const filename = mailMergeRecordFilename(title, record, data.rows.length);
    zip.file(filename, await blobBytes(blobs[index]));
    files.push({ recordNumber: record.number, filename });
  }

  const settings = resolveMailMergeSettings(data);
  zip.file(
    "mrnote-mail-merge-manifest.json",
    JSON.stringify(
      {
        schemaVersion: 1,
        output: "editable-docx-batch",
        mergeType: settings.mergeType,
        selectedRecordNumbers: files.map((file) => file.recordNumber),
        files,
        fidelity: {
          finalFields: "materialized-as-editable-text",
          sourceTemplateMutated: false,
          recipientTableEmbedded: false,
          filtersAndSortApplied: true,
          trackedChanges: "kept-as-in-main-document",
        },
        losses:
          settings.mergeType === "email"
            ? [
                {
                  feature: "external-email-send",
                  disposition: "not-performed",
                  reason:
                    "A configured email connector and final user confirmation are required. No message was sent.",
                },
              ]
            : [],
      },
      null,
      2,
    ),
  );

  return {
    blob: await zip.generateAsync({
      type: "blob",
      mimeType: "application/zip",
      compression: "DEFLATE",
    }),
    filename: mailMergeBatchFilename(title),
    files,
  };
}
