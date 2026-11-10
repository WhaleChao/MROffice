/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type * as CompressLossless from "../compress-lossless";
import type * as DocumentWriters from "../document-writers-server";
import { callEngine, EngineRefusedError } from "@mrnote/core/web/lib/engine-host/client.ts";
import { encodePdfNativeWire } from "@mrnote/core/web/lib/engine-host/pdf-native-wire.ts";
import { PdfNativeTransactionError } from "../engine/native/transaction-error";
import { readablePdfEngineError, UNKNOWN_SOURCE } from "../pdf-native-remote";
import { rethrow, workingHead } from "./pdf-native-planners";

// eslint-disable-next-line @typescript-eslint/no-explicit-any -- any writer signature
type Writer = (...args: any[]) => unknown;
type Remote<F extends Writer> = (...args: Parameters<F>) => Promise<Awaited<ReturnType<F>>>;

/** ArrayBuffer arguments as the Uint8Array the head registry knows. */
function asBytes(args: readonly unknown[]): unknown[] {
  return args.map((arg) => (arg instanceof ArrayBuffer ? new Uint8Array(arg) : arg));
}

/** Shorter byte arrays travel inline (base64); anything this size or more
 * that is not the working head is an input of its own (a picture). */
const MIN_INPUT_BYTES = 64;

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", bytes as Uint8Array<ArrayBuffer>);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

function startsWith(bytes: Uint8Array, prefix: Uint8Array): boolean {
  if (bytes.byteLength < prefix.byteLength) return false;
  for (let index = 0; index < prefix.byteLength; index += 1) if (bytes[index] !== prefix[index]) return false;
  return true;
}

/** The writer's other inputs (a background picture, replacement images) out
 * of its arguments: each becomes `{"$blob": sha256}` and travels as an
 * uploaded input of this call (checked by content on the server, never kept);
 * the working head and anything appended to it stay for the wire codec. */
async function withInputs(
  value: unknown,
  head: Uint8Array,
  inputs: Map<string, Uint8Array>,
  depth = 0,
): Promise<unknown> {
  if (depth > 16) return value;
  if (value instanceof Uint8Array) {
    if (value === head || value.byteLength < MIN_INPUT_BYTES || startsWith(value, head)) return value;
    const sha256 = await sha256Hex(value);
    inputs.set(sha256, value);
    return { $blob: sha256 };
  }
  if (Array.isArray(value)) return Promise.all(value.map((item) => withInputs(item, head, inputs, depth + 1)));
  if (value instanceof Map) {
    const entries = await Promise.all([...value].map(async ([key, item]) =>
      [key, await withInputs(item, head, inputs, depth + 1)] as const));
    return new Map(entries);
  }
  if (!value || typeof value !== "object" || ArrayBuffer.isView(value) || value instanceof Set) return value;
  const entries = await Promise.all(Object.entries(value as Record<string, unknown>).map(async ([key, item]) =>
    [key, await withInputs(item, head, inputs, depth + 1)] as const));
  return Object.fromEntries(entries);
}

export async function callPdfDocumentWriter(
  /** The page whose working head the writer rewrites (its writer set's). */
  pageId: string | null,
  fn: string,
  rawArgs: readonly unknown[],
  signal?: AbortSignal,
): Promise<unknown> {
  const head = await workingHead(asBytes(rawArgs), pageId);
  if (!head) {
    throw new PdfNativeTransactionError("source-mismatch", "The PDF source is not open for editing here.");
  }
  const inputs = new Map<string, Uint8Array>();
  const args = (await withInputs(asBytes(rawArgs), head.bytes, inputs)) as unknown[];
  const input = {
    source: { ...head.ref, sourceSha256: head.sha256 },
    fn,
    args: (await encodePdfNativeWire([...args], { sha256: head.sha256, bytes: head.bytes })) as unknown[],
  };
  const inputParts = [...inputs].map(([sha256, bytes]) => ({ sha256, bytes, role: "input" as const }));
  let answer;
  try {
    try {
      answer = await callEngine("pdf-document", "call", input, { signal, uploads: inputParts });
    } catch (error) {
      if (!(error instanceof EngineRefusedError) || !UNKNOWN_SOURCE.has(error.reason)) throw error;
      // A head the server never saw (an unsaved working chain): sent once,
      // with the inputs again (they live only for one call).
      answer = await callEngine("pdf-document", "call", input, {
        signal, uploads: [{ sha256: head.sha256, bytes: head.bytes, role: "source" }, ...inputParts],
      });
    }
  } catch (error) {
    throw readablePdfEngineError(error);
  }
  if (!answer.ok) rethrow(answer.error);
  return answer.value;
}

function remote<F extends Writer>(pageId: string | null, fn: string): Remote<F> {
  return ((...args: Parameters<F>) => callPdfDocumentWriter(pageId, fn, args)) as Remote<F>;
}

/** Every writer, each call finding its head under `pageId`. */
function bindPdfDocumentWriters(pageId: string | null) {
  const compressPdfWithStats = remote<typeof CompressLossless.compressPdfWithStats>(pageId, "compressPdfWithStats");
  return {
    /** The page this set writes for (`null`: none, see `workingHead`). */
    pageId,
    compressPdfWithStats,
    /** The bytes only (the convenience shape). */
    compressPdf: async (...args: Parameters<typeof CompressLossless.compressPdf>): Promise<Uint8Array> =>
      (await compressPdfWithStats(...args)).bytes,
    // -- B4.1: page background and picture apply ---------------------------------
    planPageBackgroundMutation:
      remote<typeof DocumentWriters.planPageBackgroundMutation>(pageId, "planPageBackgroundMutation"),
    writePageBackground: remote<typeof DocumentWriters.writePageBackground>(pageId, "writePageBackground"),
    pdfHasPageBackground: remote<typeof DocumentWriters.pdfHasPageBackground>(pageId, "pdfHasPageBackground"),
    planEmbeddedImageOverlayMutation:
      remote<typeof DocumentWriters.planEmbeddedImageOverlayMutation>(pageId, "planEmbeddedImageOverlayMutation"),
    writeEmbeddedImageOverlay:
      remote<typeof DocumentWriters.writeEmbeddedImageOverlay>(pageId, "writeEmbeddedImageOverlay"),
    /** The signal stays here: it cancels the wait, the server read runs out. */
    inspectDirectEmbeddedImagePlacements: (
      source: Uint8Array,
      signal?: AbortSignal,
    ): ReturnType<typeof DocumentWriters.inspectDirectEmbeddedImagePlacements> =>
      callPdfDocumentWriter(pageId, "inspectDirectEmbeddedImagePlacements", [source], signal) as
        ReturnType<typeof DocumentWriters.inspectDirectEmbeddedImagePlacements>,
  };
}

/** The document writers one page's window calls. */
export type PdfDocumentWriters = ReturnType<typeof bindPdfDocumentWriters>;

const boundSets = new Map<string | null, PdfDocumentWriters>();
const BOUND_SET_LIMIT = 64;

/**
 * The writers bound to one page, as `createPdfNativePlanners`: a call finds
 * its head under that page's registration only, never a twin page's that
 * holds the same bytes; `null` binds none and is refused for twins.
 */
export function createPdfDocumentWriters(pageId: string | null): PdfDocumentWriters {
  const key = pageId || null;
  let set = boundSets.get(key);
  if (set) return set;
  set = bindPdfDocumentWriters(key);
  boundSets.set(key, set);
  while (boundSets.size > BOUND_SET_LIMIT) boundSets.delete(boundSets.keys().next().value as string | null);
  return set;
}
