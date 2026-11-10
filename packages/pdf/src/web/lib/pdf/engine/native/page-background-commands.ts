/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Browser-safe part of the native page background writer (WPS 背景): the
 * command the editor (or the AI's pdf.page.background) sends and the payload
 * the adapter reads back. The planner (`planMraiNativePageBackground`) runs
 * on the engine host and checks every page against the PDF again.
 */
import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import { readPdfCommandOwnershipManifest } from "../command-ownership";
import type { PdfEngineCommand, PdfEngineSourceIdentity } from "../contracts";
import type {
  PdfNativePageBackgroundFillRequest,
  PdfNativePageBackgroundRequest,
} from "./content-stream-transaction";
import {
  MRAI_NATIVE_PAGE_BACKGROUND_ACTION,
  MRAI_NATIVE_PAGE_BACKGROUND_SCHEMA_VERSION,
  PDF_NATIVE_PAGE_BACKGROUND_CONTROL_ID,
  PDF_NATIVE_PAGE_BACKGROUND_FEATURE_ID,
  PDF_NATIVE_PAGE_BACKGROUND_MAX_PAGES,
} from "./vocabulary";

export interface PdfNativePageBackgroundPayload {
  readonly schemaVersion: typeof MRAI_NATIVE_PAGE_BACKGROUND_SCHEMA_VERSION;
  readonly sourceSha256: string;
  readonly request: PdfNativePageBackgroundRequest;
}

const SHA256 = /^[0-9a-f]{64}$/;
const BASE64 = /^[A-Za-z0-9+/]+={0,2}$/;
/** A background picture's bytes, base64 (the native image writer's limit). */
const MAX_IMAGE_BASE64 = 2_796_208;

function invalid(message: string): never {
  throw new Error(message);
}

function plain(value: unknown): Readonly<Record<string, unknown>> | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) return null;
  return value as Readonly<Record<string, unknown>>;
}

function pageNumber(value: unknown): number {
  if (!Number.isSafeInteger(value) || (value as number) < 1 || (value as number) > 100_000) {
    invalid("Page background page is invalid.");
  }
  return value as number;
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
    invalid("Page background object reference is invalid.");
  }
  return Object.freeze({ objectNumber: record.objectNumber as number, generation: record.generation as number });
}

function readFill(value: unknown): PdfNativePageBackgroundFillRequest {
  const fill = plain(value);
  if (fill?.kind === "color") {
    const color = fill.color;
    if (
      !Array.isArray(color) ||
      color.length !== 3 ||
      color.some((item) => typeof item !== "number" || !Number.isFinite(item) || item < 0 || item > 1)
    ) {
      invalid("Page background colour is invalid.");
    }
    return Object.freeze({ kind: "color" as const, color: Object.freeze([color[0], color[1], color[2]]) as readonly [number, number, number] });
  }
  if (fill?.kind === "image") {
    const opacity = fill.opacity;
    if (
      (fill.mimeType !== "image/png" && fill.mimeType !== "image/jpeg") ||
      typeof fill.imageBase64 !== "string" ||
      fill.imageBase64.length < 4 ||
      fill.imageBase64.length > MAX_IMAGE_BASE64 ||
      !BASE64.test(fill.imageBase64) ||
      (opacity !== undefined && (typeof opacity !== "number" || !Number.isFinite(opacity) || opacity < 0.05 || opacity > 1)) ||
      (fill.fit !== undefined && fill.fit !== "stretch" && fill.fit !== "fit") ||
      (fill.color !== undefined &&
        (!Array.isArray(fill.color) || fill.color.length !== 3 ||
          fill.color.some((item) => typeof item !== "number" || !Number.isFinite(item) || item < 0 || item > 1)))
    ) {
      invalid("Page background picture is invalid.");
    }
    const under = fill.color as readonly number[] | undefined;
    return Object.freeze({
      kind: "image" as const,
      mimeType: fill.mimeType,
      imageBase64: fill.imageBase64,
      ...(opacity === undefined ? {} : { opacity: opacity as number }),
      ...(fill.fit === undefined ? {} : { fit: fill.fit as "stretch" | "fit" }),
      ...(under ? { color: Object.freeze([under[0], under[1], under[2]]) as readonly [number, number, number] } : {}),
    });
  }
  invalid("Page background fill is invalid.");
}

/** Read a page background request as the editor may send it; the planner
 * then holds every page to the PDF. */
export function readPdfNativePageBackgroundRequest(value: unknown): PdfNativePageBackgroundRequest {
  const record = plain(value);
  if (
    !record ||
    (record.operation !== "apply" && record.operation !== "restore") ||
    typeof record.expectedSourceSha256 !== "string" ||
    !SHA256.test(record.expectedSourceSha256)
  ) {
    invalid("Page background request is invalid.");
  }
  if (record.operation === "apply") {
    const pages = record.pageNumbers;
    if (!Array.isArray(pages) || pages.length < 1 || pages.length > PDF_NATIVE_PAGE_BACKGROUND_MAX_PAGES) {
      invalid("Page background pages are invalid.");
    }
    return Object.freeze({
      operation: "apply" as const,
      expectedSourceSha256: record.expectedSourceSha256,
      pageNumbers: Object.freeze(pages.map(pageNumber)),
      fill: readFill(record.fill),
    });
  }
  const pages = record.pages;
  if (!Array.isArray(pages) || pages.length < 1 || pages.length > PDF_NATIVE_PAGE_BACKGROUND_MAX_PAGES) {
    invalid("Page background pages are invalid.");
  }
  return Object.freeze({
    operation: "restore" as const,
    expectedSourceSha256: record.expectedSourceSha256,
    pages: Object.freeze(
      pages.map((item) => {
        const page = plain(item);
        if (!page) invalid("Page background page is invalid.");
        return Object.freeze({
          pageNumber: pageNumber(page.pageNumber),
          background: page.background === null || page.background === undefined ? null : objectRef(page.background),
        });
      }),
    ),
  });
}

export function readPdfNativePageBackgroundPayload(value: unknown): PdfNativePageBackgroundPayload {
  const record = plain(value);
  if (
    !record ||
    record.schemaVersion !== MRAI_NATIVE_PAGE_BACKGROUND_SCHEMA_VERSION ||
    typeof record.sourceSha256 !== "string" ||
    !SHA256.test(record.sourceSha256)
  ) {
    invalid("Page background payload is invalid.");
  }
  const request = readPdfNativePageBackgroundRequest(record.request);
  if (request.expectedSourceSha256 !== record.sourceSha256) {
    invalid("Page background payload names another source.");
  }
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_PAGE_BACKGROUND_SCHEMA_VERSION,
    sourceSha256: record.sourceSha256,
    request,
  });
}

export function createPdfNativePageBackgroundOwnershipManifest() {
  return readPdfCommandOwnershipManifest([
    {
      featureId: PDF_NATIVE_PAGE_BACKGROUND_FEATURE_ID,
      action: MRAI_NATIVE_PAGE_BACKGROUND_ACTION,
      authority: "mrai-native",
      mutation: "native-bytes",
      output: "current-pdf",
    },
  ]);
}

export function buildPdfNativePageBackgroundCommand(input: {
  readonly source: PdfEngineSourceIdentity;
  readonly payload: PdfNativePageBackgroundPayload;
  readonly transactionId: string;
  readonly sessionRevision: number;
  readonly selectionRevision: number;
}): PdfEngineCommand {
  const payload = readPdfNativePageBackgroundPayload(input.payload);
  if (
    payload.sourceSha256 !== input.source.sourceSha256 ||
    !input.transactionId.trim() ||
    !Number.isSafeInteger(input.sessionRevision) ||
    input.sessionRevision < 0 ||
    !Number.isSafeInteger(input.selectionRevision) ||
    input.selectionRevision < 1
  ) {
    invalid("Page background command authority is stale.");
  }
  return Object.freeze({
    ...input.source,
    featureId: PDF_NATIVE_PAGE_BACKGROUND_FEATURE_ID,
    controlId: PDF_NATIVE_PAGE_BACKGROUND_CONTROL_ID,
    action: MRAI_NATIVE_PAGE_BACKGROUND_ACTION,
    transactionId: input.transactionId,
    sessionRevision: input.sessionRevision,
    selectionRevision: input.selectionRevision,
    output: "current-pdf",
    payload: payload as unknown as JsonValue,
  });
}

const NATIVE_BACKGROUND_MARKER = new TextEncoder().encode("/MRAIBackground true");

/** Whether the PDF holds a background stream this writer made (its
 * dictionary's marker): the Background dialog then offers Remove. */
export function pdfHasNativePageBackgroundMarker(bytes: Uint8Array): boolean {
  const marker = NATIVE_BACKGROUND_MARKER;
  const first = marker[0];
  for (
    let index = bytes.indexOf(first);
    index >= 0 && index <= bytes.length - marker.length;
    index = bytes.indexOf(first, index + 1)
  ) {
    let matched = true;
    for (let offset = 1; offset < marker.length; offset += 1) {
      if (bytes[index + offset] !== marker[offset]) {
        matched = false;
        break;
      }
    }
    if (matched) return true;
  }
  return false;
}
