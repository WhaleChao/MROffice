/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Browser-safe part of the native page split writer (WPS 分割页面): the
 * command the editor sends and the payload the adapter reads back. The
 * planner (`planMraiNativePageSplit`) is server-only; it checks every
 * geometry again against the PDF.
 */
import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import { readPdfCommandOwnershipManifest } from "../command-ownership";
import type { PdfEngineCommand, PdfEngineSourceIdentity } from "../contracts";
import type {
  PdfNativePageBoxEntries,
  PdfNativePageSplitMarkers,
  PdfNativePageSplitRequest,
} from "./content-stream-transaction";
import {
  MRAI_NATIVE_PAGE_SPLIT_ACTION,
  MRAI_NATIVE_PAGE_SPLIT_SCHEMA_VERSION,
  PDF_NATIVE_PAGE_SPLIT_CONTROL_ID,
  PDF_NATIVE_PAGE_SPLIT_FEATURE_ID,
  PDF_NATIVE_PAGE_SPLIT_MAX_LINES,
  PDF_NATIVE_PAGE_SPLIT_MAX_PAGES,
  PDF_NATIVE_PAGE_SPLIT_MAX_PARTS,
} from "./vocabulary";

export interface PdfNativePageSplitPayload {
  readonly schemaVersion: typeof MRAI_NATIVE_PAGE_SPLIT_SCHEMA_VERSION;
  readonly sourceSha256: string;
  readonly request: PdfNativePageSplitRequest;
}

const SHA256 = /^[0-9a-f]{64}$/;
const SPLIT_ID = /^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$/;
const BOX_KEYS = ["MediaBox", "CropBox", "TrimBox", "BleedBox", "ArtBox"] as const;
const MARKER_KEYS = ["MRAISplitId", "MRAISplitPart", "MRAISplitOf"] as const;

function invalid(message: string): never {
  throw new Error(message);
}

function plain(value: unknown): Readonly<Record<string, unknown>> | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) return null;
  return value as Readonly<Record<string, unknown>>;
}

function objectRef(value: unknown): Readonly<{ objectNumber: number; generation: number }> {
  const record = plain(value);
  if (
    !record ||
    !Number.isSafeInteger(record.objectNumber) ||
    (record.objectNumber as number) < 1 ||
    !Number.isSafeInteger(record.generation) ||
    (record.generation as number) < 0 ||
    (record.generation as number) > 65_535
  ) {
    invalid("Page split object reference is invalid.");
  }
  return Object.freeze({
    objectNumber: record.objectNumber as number,
    generation: record.generation as number,
  });
}

function cutLines(value: unknown): readonly number[] {
  if (
    !Array.isArray(value) ||
    value.length > PDF_NATIVE_PAGE_SPLIT_MAX_LINES ||
    value.some((line) => typeof line !== "number" || !Number.isFinite(line) || line <= 0 || line >= 1)
  ) {
    invalid("Page split lines are invalid.");
  }
  return Object.freeze([...(value as number[])]);
}

function rawEntries<K extends string>(
  value: unknown,
  keys: readonly K[],
): Readonly<Record<K, string | null>> {
  const record = plain(value);
  if (!record || Object.keys(record).some((key) => !(keys as readonly string[]).includes(key))) {
    invalid("Page split entries are invalid.");
  }
  const entries = {} as Record<K, string | null>;
  for (const key of keys) {
    const entry = record[key] ?? null;
    if (entry !== null && (typeof entry !== "string" || entry.length < 1 || entry.length > 200)) {
      invalid("Page split entries are invalid.");
    }
    entries[key] = entry as string | null;
  }
  return Object.freeze(entries);
}

/** Read a page split request as the editor may send it; the planner then
 * holds every value to the PDF. */
export function readPdfNativePageSplitRequest(value: unknown): PdfNativePageSplitRequest {
  const record = plain(value);
  if (
    !record ||
    (record.operation !== "split" && record.operation !== "merge") ||
    typeof record.splitId !== "string" ||
    !SPLIT_ID.test(record.splitId) ||
    typeof record.expectedSourceSha256 !== "string" ||
    !SHA256.test(record.expectedSourceSha256) ||
    typeof record.keepOriginal !== "boolean"
  ) {
    invalid("Page split request is invalid.");
  }
  if (record.operation === "split") {
    const pages = record.pages;
    if (!Array.isArray(pages) || pages.length < 1 || pages.length > PDF_NATIVE_PAGE_SPLIT_MAX_PAGES) {
      invalid("Page split pages are invalid.");
    }
    return Object.freeze({
      operation: "split" as const,
      splitId: record.splitId,
      expectedSourceSha256: record.expectedSourceSha256,
      keepOriginal: record.keepOriginal,
      pages: Object.freeze(
        pages.map((item) => {
          const page = plain(item);
          const grid = page ? plain(page.grid) : null;
          if (
            !page ||
            !grid ||
            !Number.isSafeInteger(page.pageNumber) ||
            (page.pageNumber as number) < 1
          ) {
            invalid("Page split page is invalid.");
          }
          const rows = cutLines(grid.rows);
          const columns = cutLines(grid.columns);
          const parts = (rows.length + 1) * (columns.length + 1);
          if (parts < 2 || parts > PDF_NATIVE_PAGE_SPLIT_MAX_PARTS) {
            invalid("Page split lines are invalid.");
          }
          return Object.freeze({
            pageNumber: page.pageNumber as number,
            grid: Object.freeze({ rows, columns }),
          });
        }),
      ),
    });
  }
  const groups = record.groups;
  if (!Array.isArray(groups) || groups.length < 1 || groups.length > PDF_NATIVE_PAGE_SPLIT_MAX_PAGES) {
    invalid("Page merge groups are invalid.");
  }
  return Object.freeze({
    operation: "merge" as const,
    splitId: record.splitId,
    expectedSourceSha256: record.expectedSourceSha256,
    keepOriginal: record.keepOriginal,
    groups: Object.freeze(
      groups.map((item) => {
        const group = plain(item);
        if (
          !group ||
          !Array.isArray(group.parts) ||
          group.parts.length < 1 ||
          group.parts.length > PDF_NATIVE_PAGE_SPLIT_MAX_PARTS
        ) {
          invalid("Page merge group is invalid.");
        }
        return Object.freeze({
          page: objectRef(group.page),
          parts: Object.freeze(group.parts.map(objectRef)),
          restoreBoxes: rawEntries(group.restoreBoxes, BOX_KEYS) as PdfNativePageBoxEntries,
          restoreMarkers: rawEntries(group.restoreMarkers, MARKER_KEYS) as PdfNativePageSplitMarkers,
        });
      }),
    ),
  });
}

export function readPdfNativePageSplitPayload(value: unknown): PdfNativePageSplitPayload {
  const record = plain(value);
  if (
    !record ||
    record.schemaVersion !== MRAI_NATIVE_PAGE_SPLIT_SCHEMA_VERSION ||
    typeof record.sourceSha256 !== "string" ||
    !SHA256.test(record.sourceSha256)
  ) {
    invalid("Page split payload is invalid.");
  }
  const request = readPdfNativePageSplitRequest(record.request);
  if (request.expectedSourceSha256 !== record.sourceSha256) {
    invalid("Page split payload names another source.");
  }
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_PAGE_SPLIT_SCHEMA_VERSION,
    sourceSha256: record.sourceSha256,
    request,
  });
}

export function createPdfNativePageSplitOwnershipManifest() {
  return readPdfCommandOwnershipManifest([
    {
      featureId: PDF_NATIVE_PAGE_SPLIT_FEATURE_ID,
      action: MRAI_NATIVE_PAGE_SPLIT_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
  ]);
}

export function buildPdfNativePageSplitCommand(input: {
  readonly source: PdfEngineSourceIdentity;
  readonly payload: PdfNativePageSplitPayload;
  readonly transactionId: string;
  readonly sessionRevision: number;
  readonly selectionRevision: number;
}): PdfEngineCommand {
  const payload = readPdfNativePageSplitPayload(input.payload);
  if (
    payload.sourceSha256 !== input.source.sourceSha256 ||
    !input.transactionId.trim() ||
    !Number.isSafeInteger(input.sessionRevision) ||
    input.sessionRevision < 0 ||
    !Number.isSafeInteger(input.selectionRevision) ||
    input.selectionRevision < 1
  ) {
    invalid("Page split command authority is stale.");
  }
  return Object.freeze({
    ...input.source,
    featureId: PDF_NATIVE_PAGE_SPLIT_FEATURE_ID,
    controlId: PDF_NATIVE_PAGE_SPLIT_CONTROL_ID,
    action: MRAI_NATIVE_PAGE_SPLIT_ACTION,
    transactionId: input.transactionId,
    sessionRevision: input.sessionRevision,
    selectionRevision: input.selectionRevision,
    output: "current-pdf",
    payload: payload as unknown as JsonValue,
  });
}
