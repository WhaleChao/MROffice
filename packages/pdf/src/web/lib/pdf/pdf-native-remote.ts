/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { callEngine, EngineRefusedError, EngineUnavailableError, engineErrorMessageKey, type EngineUnavailableReason } from "@mrnote/core/web/lib/engine-host/client.ts";
import { notePdfNativeHeadHeld, pdfNativeHeadBase, pdfNativeHeadUnheld } from "./remote/pdf-native-sources";
import type { PdfNativeEngineMethods, PdfNativeEngineSource } from "@/lib/engine-host/pdf-native-types";
import {
  decodePdfNativeWire,
  PDF_WIRE_BLOB,
  PDF_WIRE_CANDIDATE,
  PDF_WIRE_OFFLOAD_BYTES,
  encodePdfNativeWire,
} from "@mrnote/core/web/lib/engine-host/pdf-native-wire.ts";

/** The working head's identity minus its digest (the bytes give that). */
export type PdfNativeEngineSourceRef = Omit<PdfNativeEngineSource, "sourceSha256">;

export const UNKNOWN_SOURCE: ReadonlySet<string> = new Set(["blob_unknown", "blob_digest_mismatch"]);

/** What the editor shows when the engine could not answer: the editors show
 * a planner's own error text, so the transport's ("engine offline: ...")
 * would otherwise reach the user. Same class and `reason`, readable text. */
const UNAVAILABLE_TEXT: Readonly<Record<EngineUnavailableReason, string>> = Object.freeze({
  offline: "PDF editing needs a connection. Nothing was changed.",
  signed_out: "Sign in again to edit this PDF. Nothing was changed.",
  busy: "The PDF editor is busy. Try again in a moment; nothing was changed.",
  rate_limited: "The PDF editor is busy. Try again in a moment; nothing was changed.",
  timeout: "The PDF edit took too long and was stopped. Nothing was changed.",
  failed: "The PDF editor could not answer. Nothing was changed.",
});

const REFUSED_TEXT = /^The PDF editor refused this edit \([^)]*\)\. Nothing was changed\.$/;

/** The catalog key of an engine failure that reaches the editor only as its
 * text: the PDF engine session turns an adapter's failure into a refusal
 * whose `detail` is the sentence above, and callers rethrow that detail. */
export function pdfEngineFailureMessageKey(text: string): ReturnType<typeof engineErrorMessageKey> | null {
  for (const [reason, sentence] of Object.entries(UNAVAILABLE_TEXT)) {
    if (sentence === text) return engineErrorMessageKey(new EngineUnavailableError(reason as EngineUnavailableReason));
  }
  return REFUSED_TEXT.test(text) ? "engineHost.refused" : null;
}

export function readablePdfEngineError(error: unknown): unknown {
  if (error instanceof EngineUnavailableError) {
    const readable = new EngineUnavailableError(error.reason, undefined, error.retryAfterMs);
    readable.message = UNAVAILABLE_TEXT[error.reason] ?? UNAVAILABLE_TEXT.failed;
    if (error.status !== undefined) readable.status = error.status;
    if (error.code !== undefined) readable.code = error.code;
    return readable;
  }
  if (error instanceof EngineRefusedError && !UNKNOWN_SOURCE.has(error.reason)) {
    const readable = new EngineRefusedError(error.reason, error.details);
    readable.message = `The PDF editor refused this edit (${error.reason}). Nothing was changed.`;
    return readable;
  }
  return error;
}

/** The methods that read a PDF (the sourceless `pure` goes through `callEngine`). */
type PdfNativeSourceMethod = Exclude<keyof PdfNativeEngineMethods, "pure">;

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const owned = new Uint8Array(bytes.byteLength);
  owned.set(bytes);
  const digest = await crypto.subtle.digest("SHA-256", owned.buffer);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

/**
 * The large byte arrays of one call's arguments, kept out of its JSON (the
 * request is bounded at 1 MiB; a candidate carrying an installed font or a
 * replacement picture is larger): `offload` for `encodePdfNativeWire`. A
 * candidate built on the head becomes `{"$candidate": sha}` -- the server
 * usually holds it (it built it) and only after `blob_unknown` is it sent;
 * anything else becomes `{"$blob": sha}` and travels with the call.
 */
export function createPdfNativeOffload() {
  const candidates = new Map<string, Uint8Array>();
  const inputs = new Map<string, Uint8Array>();
  const offload = async (bytes: Uint8Array, appended: boolean): Promise<unknown | null> => {
    if (bytes.byteLength < PDF_WIRE_OFFLOAD_BYTES) return null;
    const sha256 = await sha256Hex(bytes);
    if (appended) {
      candidates.set(sha256, bytes);
      return { [PDF_WIRE_CANDIDATE]: sha256 };
    }
    inputs.set(sha256, bytes);
    return { [PDF_WIRE_BLOB]: sha256 };
  };
  return { offload, candidates, inputs };
}

export type PdfNativeOffload = ReturnType<typeof createPdfNativeOffload>;

function isWireRecord(value: unknown): value is Readonly<Record<string, unknown>> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

/** A plan's command in its wire form: a long text it carries (a stamp's or
 * a picture's base64) travels beside the JSON as `$text`, the route bounding
 * the JSON at 1 MiB; the engine reads it back before the command. */
export async function encodePdfNativeCommand(
  command: unknown,
  offloaded: PdfNativeOffload,
): Promise<Readonly<Record<string, unknown>>> {
  const encoded = await encodePdfNativeWire(JSON.parse(JSON.stringify(command)), null, undefined, offloaded.offload);
  if (!isWireRecord(encoded)) throw new Error("The PDF command did not encode as an object.");
  return encoded;
}








const BUSY_RETRY_BUDGET_MS = 45_000;





const RELOAD_RETRY_BUDGET_MS = 20_000;
const RELOAD_NOT_FOUND_BUDGET_MS = 8_000;

/** An answer an API reload gives a call for a moment: the dev proxy's bare
 * 500 while the server is down, a gateway 502/504, the engine pool not up
 * yet (503 engine_unavailable), a route or source not found yet (404). The
 * engine's own failure (502 engine_failed) and a planner's refusal are
 * never one of these. */
function reloadTransient(error: EngineUnavailableError, elapsedMs: number): boolean {
  if (error.reason !== "failed" || error.status === undefined) return false;
  const code = error.code;
  if (error.status === 404) {
    return (code === undefined || code === "not_found" || code === "engine_not_found") &&
      elapsedMs < RELOAD_NOT_FOUND_BUDGET_MS;
  }
  if (elapsedMs >= RELOAD_RETRY_BUDGET_MS) return false;
  if (error.status === 503) return code === undefined || code === "engine_unavailable";
  return (error.status === 500 || error.status === 502 || error.status === 504) && code === undefined;
}

function waitOrAbort(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) { reject(new DOMException("Cancelled", "AbortError")); return; }
    const timer = setTimeout(() => { signal?.removeEventListener("abort", onAbort); resolve(); }, ms);
    const onAbort = () => { clearTimeout(timer); reject(new DOMException("Cancelled", "AbortError")); };
    signal?.addEventListener("abort", onAbort, { once: true });
  });
}

export async function callPdfNativeEngine<M extends PdfNativeSourceMethod>(
  method: M,
  input: PdfNativeEngineMethods[M]["input"],
  options: Readonly<{ signal?: AbortSignal; workingBytes: () => Uint8Array | null; offloaded?: PdfNativeOffload }>,
): Promise<PdfNativeEngineMethods[M]["output"]> {
  // A prewarm (or a warmed proof) is a hint: a busy engine is its answer.
  if (method === "prewarm" || (method === "prepare" && (input as { warm?: boolean }).warm === true)) {
    return callPdfNativeEngineOnce(method, input, options);
  }
  const started = Date.now();
  for (;;) {
    try {
      return await callPdfNativeEngineOnce(method, input, options);
    } catch (error) {
      // A dropped request while the browser is online (a reload of the
      // server, a connection reset under load) is asked again like a busy
      // engine: a save's plan failed once as "needs a connection" with every
      
      const transient = error instanceof EngineUnavailableError &&
        (error.reason === "busy" ||
          (error.reason === "offline" && (typeof navigator === "undefined" || navigator.onLine !== false)) ||
          reloadTransient(error, Date.now() - started));
      if (!transient) throw error;
      const delay = error.reason === "busy" ? Math.min(2_000, Math.max(250, error.retryAfterMs ?? 1_000))
        : error.reason === "failed" ? 750 : 500;
      if (Date.now() - started + delay > BUSY_RETRY_BUDGET_MS) throw error;
      await waitOrAbort(delay, options.signal);
    }
  }
}

async function callPdfNativeEngineOnce<M extends PdfNativeSourceMethod>(
  method: M,
  input: PdfNativeEngineMethods[M]["input"],
  options: Readonly<{ signal?: AbortSignal; workingBytes: () => Uint8Array | null; offloaded?: PdfNativeOffload }>,
): Promise<PdfNativeEngineMethods[M]["output"]> {
  const inputs = [...(options.offloaded?.inputs ?? [])].map(([sha256, bytes]) => ({ sha256, bytes, role: "input" as const }));
  const candidates = [...(options.offloaded?.candidates ?? [])]
    .map(([sha256, bytes]) => ({ sha256, bytes, role: "input" as const }));
  const sourceSha256 = input.source.sourceSha256;
  try {
    // A head assembled here that the server was never sent goes with this
    // first call instead of after its `blob_unknown` (pdf-native-sources).
    const firstBytes = !candidates.length && pdfNativeHeadUnheld(sourceSha256) ? options.workingBytes() : null;
    const base = firstBytes && !inputs.length ? pdfNativeHeadBase(sourceSha256) : null;
    if (firstBytes && base && firstBytes.byteLength > base.length) {
      // Only what this head appended to the one before (which the server
      // usually holds); a base it does not hold answers blob_unknown and the
      // whole head goes below.
      try {
        const answer = (await callEngine("pdf-native", method, input as never, {
          signal: options.signal,
          upload: { sha256: sourceSha256, bytes: firstBytes.subarray(base.length),
            base: { sha256: base.sha256, size: base.length } },
        })) as PdfNativeEngineMethods[M]["output"];
        notePdfNativeHeadHeld(sourceSha256);
        return answer;
      } catch (error) {
        if (!(error instanceof EngineRefusedError) || !UNKNOWN_SOURCE.has(error.reason)) throw error;
      }
    }
    if (firstBytes) {
      const answer = (await callEngine("pdf-native", method, input as never, {
        signal: options.signal,
        ...(inputs.length
          ? { uploads: [{ sha256: sourceSha256, bytes: firstBytes, role: "source" as const }, ...inputs] }
          : { upload: { sha256: sourceSha256, bytes: firstBytes } }),
      })) as PdfNativeEngineMethods[M]["output"];
      notePdfNativeHeadHeld(sourceSha256);
      return answer;
    }
    try {
      return (await callEngine("pdf-native", method, input as never, {
        signal: options.signal,
        ...(inputs.length ? { uploads: inputs } : {}),
      })) as PdfNativeEngineMethods[M]["output"];
    } catch (error) {
      if (!(error instanceof EngineRefusedError) || !UNKNOWN_SOURCE.has(error.reason)) throw error;
      const bytes = options.workingBytes();
      if (!bytes && !candidates.length) throw error;
      if (!inputs.length && !candidates.length) {
        const answer = (await callEngine("pdf-native", method, input as never, {
          signal: options.signal,
          upload: { sha256: sourceSha256, bytes: bytes! },
        })) as PdfNativeEngineMethods[M]["output"];
        notePdfNativeHeadHeld(sourceSha256);
        return answer;
      }
      // Whatever the server did not hold goes along once: the head and the
      // candidates this call names, with its other inputs.
      const answer = (await callEngine("pdf-native", method, input as never, {
        signal: options.signal,
        uploads: [
          ...(bytes ? [{ sha256: sourceSha256, bytes, role: "source" as const }] : []),
          ...candidates, ...inputs,
        ],
      })) as PdfNativeEngineMethods[M]["output"];
      if (bytes) notePdfNativeHeadHeld(sourceSha256);
      return answer;
    }
  } catch (error) {
    throw readablePdfEngineError(error);
  }
}

/** Rebuild an answer's tagged bytes against the head this editor holds. */
export function decodePdfNativeAnswer<T>(answer: unknown, source: Readonly<{ sha256: string; bytes: Uint8Array }>): Promise<T> {
  return decodePdfNativeWire<T>(answer, (base) => (base === source.sha256 ? source.bytes : null));
}
