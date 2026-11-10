/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type * as Server from "../flatten-export-server";
import type { FlattenImageBytes } from "../flatten-apply";
import type { ImageOp, SignatureImageOp, TextOp } from "@mrnote/core/web/lib/pdf/flatten.ts";
import { callEngine, type EngineBlobUploadPart } from "@mrnote/core/web/lib/engine-host/client.ts";
import type { PdfDocumentCallError } from "@/lib/engine-host/pdf-document-types";
import { encodePdfNativeWire } from "@mrnote/core/web/lib/engine-host/pdf-native-wire.ts";
import { readablePdfEngineError } from "../pdf-native-remote";
import { rethrow } from "./pdf-native-planners";

type ExportMethod = "export.plan" | "export" | "export.redact";

/** Measurements written after the bake (and after any redacted-page swap). */
export type FlattenExportMeasurementsInput = Server.FlattenExportMeasurements;

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", bytes as Uint8Array<ArrayBuffer>);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

/** Every byte array of the arguments as `{"$blob": sha}`, collected. */
async function withUploads(value: unknown, uploads: Map<string, Uint8Array>, depth = 0): Promise<unknown> {
  if (depth > 24) return value;
  if (value instanceof ArrayBuffer) value = new Uint8Array(value);
  if (value instanceof Uint8Array) {
    const sha256 = await sha256Hex(value);
    uploads.set(sha256, value);
    return { $blob: sha256 };
  }
  if (Array.isArray(value)) return Promise.all(value.map((item) => withUploads(item, uploads, depth + 1)));
  if (value instanceof Map) {
    return new Map(await Promise.all([...value].map(async ([key, item]) =>
      [key, await withUploads(item, uploads, depth + 1)] as const)));
  }
  if (!value || typeof value !== "object" || ArrayBuffer.isView(value) || value instanceof Set || value instanceof Date) {
    return value;
  }
  return Object.fromEntries(await Promise.all(Object.entries(value as Record<string, unknown>).map(
    async ([key, item]) => [key, await withUploads(item, uploads, depth + 1)] as const)));
}

/**
 * One sourceless `pdf-document` call (the export, the form commands): every
 * byte array of the arguments travels as an upload, `source` (the reader's
 * PDF) as the call's source. `rebuild` may turn a refusal back into the
 * caller's own error class.
 */
export async function callPdfDocumentSourceless(
  method: ExportMethod | "form.plan" | "form",
  rawArgs: readonly unknown[],
  options: {
    readonly source?: Uint8Array;
    readonly fn?: string;
    readonly rebuild?: (error: PdfDocumentCallError) => Error | null;
  } = {},
): Promise<unknown> {
  const uploads = new Map<string, Uint8Array>();
  const args = (await withUploads([...rawArgs], uploads)) as unknown[];
  const sourceSha = options.source ? await sha256Hex(options.source) : null;
  const parts: EngineBlobUploadPart[] = [...uploads].map(([sha256, bytes]) => ({
    sha256, bytes, role: sha256 === sourceSha ? "source" : "input",
  }));
  const input = {
    ...(options.fn ? { fn: options.fn } : {}),
    args: (await encodePdfNativeWire(args, null)) as unknown[],
  };
  let answer;
  try {
    answer = await callEngine("pdf-document", method as "form", input as never,
      parts.length ? { uploads: parts } : {});
  } catch (error) {
    throw readablePdfEngineError(error);
  }
  if (!answer.ok) {
    const rebuilt = options.rebuild?.(answer.error);
    if (rebuilt) throw rebuilt;
    rethrow(answer.error);
  }
  return answer.value;
}

function callPdfExport(method: ExportMethod, rawArgs: readonly unknown[], source?: Uint8Array): Promise<unknown> {
  return callPdfDocumentSourceless(method, rawArgs, source ? { source } : {});
}

export function planFlattenResources(
  ...args: Parameters<typeof Server.planFlattenResources>
): Promise<ReturnType<typeof Server.planFlattenResources>> {
  return callPdfExport("export.plan", args) as Promise<ReturnType<typeof Server.planFlattenResources>>;
}

export function exportAnnotatedPdf(
  ...args: Parameters<typeof Server.exportAnnotatedPdf>
): ReturnType<typeof Server.exportAnnotatedPdf> {
  return callPdfExport("export", args, args[0]) as ReturnType<typeof Server.exportAnnotatedPdf>;
}

export function replaceRedactedPages(
  ...args: Parameters<typeof Server.replaceRedactedPages>
): ReturnType<typeof Server.replaceRedactedPages> {
  return callPdfExport("export.redact", args, args[0]) as ReturnType<typeof Server.replaceRedactedPages>;
}

type Picture = FlattenImageBytes | Uint8Array | ArrayBuffer | null | undefined;

/** The browser's renderers, by kind (each the bake's former resolver). */
export interface FlattenResourceRenderers {
  readonly image?: (op: ImageOp) => Promise<Picture>;
  readonly formula?: (op: TextOp) => Promise<Picture>;
  readonly text?: (op: TextOp) => Promise<Picture>;
  readonly signature?: (op: SignatureImageOp) => Promise<Picture>;
}

function asPicture(value: Picture): FlattenImageBytes | null {
  if (!value) return null;
  if (value instanceof Uint8Array || value instanceof ArrayBuffer) return { bytes: new Uint8Array(value) };
  return {
    bytes: value.bytes instanceof Uint8Array ? value.bytes : new Uint8Array(value.bytes),
    ...(value.mimeType ? { mimeType: value.mimeType } : {}),
    ...(typeof value.displayWidth === "number" ? { displayWidth: value.displayWidth } : {}),
    ...(typeof value.displayHeight === "number" ? { displayHeight: value.displayHeight } : {}),
  };
}

/**
 * Draw what the export's plan asked for. As the bake's resolvers did: an
 * image or signature that fails to draw fails the export; a formula or a
 * source-run text box that fails is sent as "no picture" (the bake then
 * applies its own refusal or plain-text fallback).
 */
export async function drawFlattenResources(
  requests: readonly Server.FlattenResourceRequest[],
  renderers: FlattenResourceRenderers,
): Promise<Record<string, FlattenImageBytes | null>> {
  const resolved: Record<string, FlattenImageBytes | null> = {};
  for (const request of requests) {
    if (request.kind === "image") {
      resolved[request.key] = asPicture(renderers.image ? await renderers.image(request.op as ImageOp) : null);
    } else if (request.kind === "signature") {
      resolved[request.key] = asPicture(
        renderers.signature ? await renderers.signature(request.op as SignatureImageOp) : null);
    } else {
      const render = request.kind === "formula" ? renderers.formula : renderers.text;
      try {
        resolved[request.key] = asPicture(render ? await render(request.op as TextOp) : null);
      } catch {
        resolved[request.key] = null;
      }
    }
  }
  return resolved;
}

/** The catalog key of an export failure the reader can act on, if any. */
export function pdfExportFailureMessageKey(
  error: unknown,
): "engineHost.exportFontMissing" | "pdf.viewer.pageOperationsSavingDownload" | null {
  const code = (error as { code?: unknown } | null)?.code;
  if (code === "export_font_missing") return "engineHost.exportFontMissing";
  // A page operation still saving: the export waits for the file to hold it.
  if (code === "page_structure_saving") return "pdf.viewer.pageOperationsSavingDownload";
  return null;
}
