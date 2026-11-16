/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { markPdfTextLatency } from "./text-edit-latency";
import { readPdfNativeTextRowStarts } from "./native-text-row-starts";
import { type PdfNativeTransactionErrorCode, type PdfNativePreparedContentStreamTransaction } from "./engine/native/content-stream-transaction";
import { MAX_MRAI_NATIVE_SOURCE_BYTES, PdfNativeTransactionError } from "./engine/native/transaction-error";
import type {
  PdfNativeListTextSelectionEvidence,
  PdfNativeListTextSelectionProof,
  PdfNativeSingleLineFitHint,
  PdfNativeListTextReplacementPlan,
} from "./engine/native/list-text-fidelity";
import { MRAI_NATIVE_LIST_TEXT_MISSING_GLYPH_MESSAGE } from "./engine/native/vocabulary";
import type { PdfSourceTextBoxBinding } from "./source-text-box";
import { callPdfNativeEngine, decodePdfNativeAnswer, type PdfNativeEngineSourceRef } from "./pdf-native-remote";
import { PdfWireMismatchError } from "@mrnote/core/web/lib/engine-host/pdf-native-wire.ts";
import { EngineUnavailableError } from "@mrnote/core/web/lib/engine-host/client.ts";
import {
  PdfNativeBrowserPrepareUnavailable,
  pdfNativeBrowserPrepareEnabled,
  preparePdfNativeTextInBrowser,
} from "./native-text-browser-prepare";
import { notePdfNativeHeadUnheld } from "./remote/pdf-native-sources";
import { issuePdfCompiledTextPreview } from "./native-text-compiled-preview";
import { pdfNativeResourceIssueMessage, readPdfNativeResourceIssue, type PdfNativeResourceIssue } from "./engine/native/resource-issue";

export {
  readPdfCompiledTextPreview,
  releasePdfCompiledTextPreview,
  retainPdfCompiledTextPreview,
  samePdfNativeTextCommandSelection,
  type PdfCompiledTextPreview,
} from "./native-text-compiled-preview";

export interface PdfNativeTextPreparationRequest {
  readonly requestId: string;
  readonly sourceBytes: ArrayBuffer;
  readonly selection: PdfNativeListTextSelectionEvidence;
  readonly managedBoxId: string | null;
  readonly sourceText: PdfSourceTextBoxBinding | null;
  /** Build the preview candidate of the typed text; with a resized
   * paragraph frame (its right edge, an x of the page), of the paragraph set
   * again in that frame. */
  readonly compile?: Readonly<{ replacementText: string; paragraphFrameRight?: number; paragraphLineStep?: number;
    paragraphSizeScale?: number;
    /** Where the editor's rows start in the text (native-text-row-starts). */
    rowStarts?: readonly number[] }>;
}

export interface PdfNativeTextPreparation {
  readonly compiled?: Readonly<{ plan: PdfNativeListTextReplacementPlan; prepared: PdfNativePreparedContentStreamTransaction;
    sourceGeometry?: import("./source-text-appearance").PdfTextSourceAppearance }>;
  readonly singleLineFit?: PdfNativeSingleLineFitHint;
  readonly proof: PdfNativeListTextSelectionProof;
  readonly blankedSource: Uint8Array | null;
  readonly hiddenAnnotationId: string | null;
}

export type PdfNativeTextPreparationResponse =
  | Readonly<{
      requestId: string;
      sourceSha256: string;
      result: PdfNativeTextPreparation;
    }>
  | Readonly<{
      requestId: string;
      failure: PdfNativeTransactionErrorCode | "worker-failed" | "worker-cancelled";
      missingGlyph?: true;
      resourceIssue?: PdfNativeResourceIssue;
      /** Engine boundary message for the refusal, when one exists. */
      detail?: string;
    }>;

/** The engine's refusal of a preparation, as the editor's own error. */
function preparationFailure(response: Extract<PdfNativeTextPreparationResponse, { failure: unknown }>): Error {
  const issue = readPdfNativeResourceIssue(response.resourceIssue);
  return response.failure === "worker-cancelled"
    ? new DOMException("PDF preparation superseded", "AbortError")
    : response.failure === "worker-failed"
    ? new Error("PDF text preparation failed.")
    : new PdfNativeTransactionError(
        response.failure,
        // The engine's own detail names every boundary the selection hit, and
        // the missing glyph is one of them; replacing it with the bare
        // sentence threw away what the user needs to read when the font
        // extension then cannot take the run either.
        response.missingGlyph
          ? typeof response.detail === "string" && response.detail
            ? response.detail
            : MRAI_NATIVE_LIST_TEXT_MISSING_GLYPH_MESSAGE
          : issue ? pdfNativeResourceIssueMessage(issue)
          : typeof response.detail === "string" && response.detail
            ? response.detail
            : "The selected PDF text structure could not be verified.",
        issue ?? undefined,
      );
}







const PREPARE_RETRY_BUDGET_MS = 150_000;

/** The pause before asking again, or null when the failure is the engine's
 * own answer (a refusal, a source it does not hold) and asking again would
 * get the same. */
export function pdfNativePrepareRetryDelayMs(error: unknown, attempt: number): number | null {
  if (!(error instanceof EngineUnavailableError)) return null;
  const transient = error.reason === "busy" ||
    (error.reason === "offline" && (typeof navigator === "undefined" || navigator.onLine !== false)) ||
    (error.reason === "failed" && (error.status === undefined || [500, 502, 503, 504].includes(error.status)));
  if (!transient) return null;
  const base = Math.min(8_000, 1_000 * 2 ** Math.min(attempt, 3));
  return error.reason === "busy" && error.retryAfterMs ? Math.max(base, Math.min(8_000, error.retryAfterMs)) : base;
}

function pauseOrAbort(ms: number, signal: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal.aborted) { reject(new DOMException("Cancelled", "AbortError")); return; }
    const timer = setTimeout(() => { signal.removeEventListener("abort", onAbort); resolve(); }, ms);
    const onAbort = () => { clearTimeout(timer); reject(new DOMException("Cancelled", "AbortError")); };
    signal.addEventListener("abort", onAbort, { once: true });
  });
}

/**
 * Open a selection for editing (its proof and the page with the selection
 * blanked), or with `compile` build the preview candidate of the typed text.
 * Runs on the server's engine host (`pdf-native/prepare`); the answer's
 * candidates are rebuilt against `sourceBytes` and checked by SHA-256. A
 * newer call aborts an older one through its signal.
 */
export async function preparePdfNativeTextOffThread(
  input: Omit<PdfNativeTextPreparationRequest, "requestId" | "sourceBytes"> & {
    readonly sourceBytes: Uint8Array;
    readonly signal: AbortSignal;
    /** The page and stored version the working head belongs to. */
    readonly engineSource: PdfNativeEngineSourceRef;
    /** The server's engine host only (a caller that needs its answer kept there). */
    readonly serverOnly?: boolean;
    /** Nobody waits on it: the browser's background Worker. */
    readonly background?: boolean;
  },
): Promise<PdfNativeTextPreparation> {
  if (input.signal.aborted) throw new DOMException("Cancelled", "AbortError");
  if (!input.sourceBytes.length || input.sourceBytes.length > MAX_MRAI_NATIVE_SOURCE_BYTES) {
    throw new PdfNativeTransactionError("unsupported-document", "PDF text preparation exceeds the source budget.");
  }
  const sourceSha256 = input.selection.sourceSha256;
  const ownedBytes = Uint8Array.from(input.sourceBytes);
  try {
    
    // and the candidate come from the reader's machine, nothing waits on the
    // server. What the browser cannot do (no Worker, a Worker that failed
    // or timed out) the server's engine host does as before.
    // A candidate compiled here is handed to the server in the background
    // (below), and again, awaited, by the save that names it
    // (ensurePdfNativeCandidateHeld): a hand-over that failed no longer
    // sends the compile to the engine host -- with the server busy that
    
    
    if (pdfNativeBrowserPrepareEnabled() && !input.serverOnly) {
      const browserStartedAt = performance.now();
      try {
        const result = await preparePdfNativeTextInBrowser({
          selection: input.selection, managedBoxId: input.managedBoxId, sourceText: input.sourceText,
          ...(input.compile ? { compile: input.compile } : {}),
          sourceBytes: ownedBytes, signal: input.signal, ...(input.background ? { background: true } : {}),
        });
        markPdfTextLatency("browser.prepare", { ms: Math.round(performance.now() - browserStartedAt),
          compile: Boolean(input.compile), bytes: ownedBytes.byteLength });
        if (result.compiled) {
          holdPdfNativeCandidateOnServer({
            engineSource: input.engineSource, baseSha256: sourceSha256, baseLength: ownedBytes.byteLength,
            candidateSha256: result.compiled.prepared.candidateSha256,
            candidateBytes: result.compiled.prepared.candidateBytes,
          });
        }
        return result;
      } catch (error) {
        if (!(error instanceof PdfNativeBrowserPrepareUnavailable)) throw error;
        markPdfTextLatency("browser.prepare:fallback", { reason: error.message,
          ms: Math.round(performance.now() - browserStartedAt) });
      }
    }
    const engineStartedAt = performance.now();
    const request = {
      source: { ...input.engineSource, sourceSha256 },
      selection: input.selection,
      managedBoxId: input.managedBoxId,
      sourceText: input.sourceText,
      ...(input.compile ? { compile: input.compile } : {}),
    };
    let answer: Awaited<ReturnType<typeof callPdfNativeEngine<"prepare">>>;
    for (let attempt = 0; ; attempt++) {
      try {
        answer = await callPdfNativeEngine("prepare", request, { signal: input.signal, workingBytes: () => ownedBytes });
        break;
      } catch (error) {
        const delay = pdfNativePrepareRetryDelayMs(error, attempt);
        if (delay === null || input.signal.aborted ||
          performance.now() - engineStartedAt + delay > PREPARE_RETRY_BUDGET_MS) throw error;
        markPdfTextLatency("engine.prepare:retry", { attempt: attempt + 1, delay, compile: Boolean(input.compile),
          reason: (error as EngineUnavailableError).reason, status: (error as EngineUnavailableError).status ?? null });
        await pauseOrAbort(delay, input.signal);
      }
    }
    // The engine round trip alone (the server's prepare): what a line's
    // proof or compile waits on, apart from this reader's own work.
    markPdfTextLatency("engine.prepare", { ms: Math.round(performance.now() - engineStartedAt),
      compile: Boolean(input.compile), bytes: ownedBytes.byteLength });
    if (input.signal.aborted) throw new DOMException("Cancelled", "AbortError");
    const response = await decodePdfNativeAnswer<PdfNativeTextPreparationResponse>(
      { requestId: "engine", ...answer }, { sha256: sourceSha256, bytes: ownedBytes });
    if ("failure" in response) throw preparationFailure(response);
    if (response.sourceSha256 !== sourceSha256 || !response.result?.proof) {
      response.result?.blankedSource?.fill(0);
      response.result?.compiled?.prepared.candidateBytes.fill(0);
      throw new Error("The PDF text preparation source changed.");
    }
    return response.result;
  } catch (error) {
    if (error instanceof PdfWireMismatchError) throw new Error("The PDF text preparation source changed.");
    throw error;
  } finally {
    ownedBytes.fill(0);
  }
}


/** Candidates this browser made and the server already holds (or is being
 * handed), by digest: one hand-over per candidate. */
const heldCandidates = new Set<string>();
const HELD_CANDIDATES_LIMIT = 64;
/** Until when the hand-over is known not to work (a server that does not
 * take `hold` yet, a dropped connection): candidates are compiled on the
 * engine host meanwhile, which holds what it compiles, so a save finds its
 * candidate either way. The proof stays in the browser. */
let holdUnavailableUntil = 0;
const HOLD_UNAVAILABLE_MS = 60_000;

/** Whether a candidate made here can be handed to the server now. */
export function pdfNativeCandidateHoldAvailable(): boolean {
  return Date.now() >= holdUnavailableUntil;
}

/** A candidate made in this browser's Worker, handed to the server before
 * a save names it (`hold`): the server stores the bytes a call carries in
 * the caller's candidate store on the way in. Sent as the suffix appended
 * to the head it was made from (the server holds the saved version and the
 * heads earlier calls sent). Runs in the background; a hand-over that
 * failed is tried again by the save's own wait for the candidate, and the
 * save says so if none comes. */
export function holdPdfNativeCandidateOnServer(input: Readonly<{
  engineSource: PdfNativeEngineSourceRef;
  baseSha256: string;
  baseLength: number;
  candidateSha256: string;
  candidateBytes: Uint8Array;
}>): void {
  void ensurePdfNativeCandidateHeld(input).catch(() => undefined);
}

/** A result the server wrote and keeps itself: a page's rows planned by
 * pdf-row-plan are put in the engine host's candidate store under this
 * caller and page, the store the save checks. The save that names it hands
 * nothing over (5.3 MB, 2.5-6 s on the dev server, after every batch of
 * marks); were it let go meanwhile, the save's bytes-unavailable refusal
 * hands it over anew (`again`). Never for bytes this browser planned. */
export function notePdfNativeCandidateHeldByServer(candidateSha256: string): void {
  const key = candidateSha256.toLowerCase();
  if (holding.has(key)) return;
  heldCandidates.delete(key);
  heldCandidates.add(key);
  while (heldCandidates.size > HELD_CANDIDATES_LIMIT) heldCandidates.delete(heldCandidates.values().next().value as string);
}

/** The hand-overs on their way, by candidate: a second ask for the same
 * candidate waits for the first. */
const holding = new Map<string, Promise<void>>();









export function ensurePdfNativeCandidateHeld(input: Readonly<{
  engineSource: PdfNativeEngineSourceRef;
  baseSha256: string;
  baseLength: number;
  candidateSha256: string;
  candidateBytes: Uint8Array;
  again?: boolean;
}>): Promise<void> {
  const key = input.candidateSha256.toLowerCase();
  if (input.again) {
    heldCandidates.delete(key);
    holding.delete(key);
  }
  const inFlight = holding.get(key);
  if (inFlight) return inFlight;
  if (heldCandidates.has(key)) return Promise.resolve();
  heldCandidates.add(key);
  while (heldCandidates.size > HELD_CANDIDATES_LIMIT) heldCandidates.delete(heldCandidates.values().next().value as string);
  // The planners zero the candidate once the editor is done with it: the
  // hand-over keeps its own copy for as long as it takes.
  const bytes = Uint8Array.from(input.candidateBytes);
  notePdfNativeHeadUnheld(key, { sha256: input.baseSha256, length: input.baseLength });
  const startedAt = performance.now();
  const run = callPdfNativeEngine("hold", { source: { ...input.engineSource, sourceSha256: key } },
    { workingBytes: () => bytes })
    .then(() => { markPdfTextLatency("browser.hold", { ms: Math.round(performance.now() - startedAt), bytes: bytes.byteLength }); },
      (error: unknown) => {
        heldCandidates.delete(key);
        holdUnavailableUntil = Date.now() + HOLD_UNAVAILABLE_MS;
        markPdfTextLatency("browser.hold:failed", { reason: String((error as Error)?.message ?? error).slice(0, 120) });
        throw error;
      })
    .finally(() => { bytes.fill(0); if (holding.get(key) === run) holding.delete(key); });
  holding.set(key, run);
  return run;
}








export async function warmPdfNativeTextPreparation(input: Readonly<{
  selection: PdfNativeTextPreparationRequest["selection"];
  managedBoxId: string | null;
  sourceText: PdfNativeTextPreparationRequest["sourceText"];
  engineSource: PdfNativeEngineSourceRef;
  signal?: AbortSignal;
  /** The working head's bytes, when the caller has them: the line is then
   * proved in this browser's own Worker (its parsed document and page
   * appearance stay for the click), and the server is not asked. */
  sourceBytes?: Uint8Array | null;
}>): Promise<boolean> {
  if (input.signal?.aborted) return false;
  if (pdfNativeBrowserPrepareEnabled() && input.sourceBytes?.length) {
    const startedAt = performance.now();
    const controller = new AbortController();
    const onAbort = () => controller.abort();
    input.signal?.addEventListener("abort", onAbort, { once: true });
    try {
      const result = await preparePdfNativeTextInBrowser({
        selection: input.selection, managedBoxId: input.managedBoxId, sourceText: input.sourceText,
        sourceBytes: input.sourceBytes, signal: controller.signal,
      });
      result.blankedSource?.fill(0);
      result.compiled?.prepared.candidateBytes.fill(0);
      markPdfTextLatency("browser.warm", { ms: Math.round(performance.now() - startedAt) });
      return true;
    } catch (error) {
      if (!(error instanceof PdfNativeBrowserPrepareUnavailable)) return false;
      // The browser could not: the server's engine host is warmed as before.
    } finally {
      input.signal?.removeEventListener("abort", onAbort);
    }
  }
  try {
    await callPdfNativeEngine("prepare", {
      source: { ...input.engineSource, sourceSha256: input.selection.sourceSha256 },
      selection: input.selection,
      managedBoxId: input.managedBoxId,
      sourceText: input.sourceText,
      warm: true,
    }, { signal: input.signal, workingBytes: () => null });
    return true;
  } catch {
    return false;
  }
}

export async function compilePdfNativeTextOffThread(input: Omit<Parameters<typeof preparePdfNativeTextOffThread>[0], "managedBoxId" | "sourceText" | "compile"> & { replacementText: string; paragraphFrameRight?: number; paragraphLineStep?: number;
  paragraphSizeScale?: number }) {
  const selection = structuredClone(input.selection);
  // The rows the editor showed this very edit in, where it showed one: every
  // plain edit is compiled from here (its preview, its Leave, a queued one),
  // so each is planned with them. A paragraph set in a frame asked sets
  // every line anew and takes none.
  const framed = input.paragraphFrameRight !== undefined || input.paragraphLineStep !== undefined ||
    input.paragraphSizeScale !== undefined;
  const rowStarts = framed ? null : readPdfNativeTextRowStarts(selection.item.str, input.replacementText);
  const result = await preparePdfNativeTextOffThread({ ...input, selection, managedBoxId: null, sourceText: null,
    compile: { replacementText: input.replacementText,
      ...(input.paragraphFrameRight !== undefined ? { paragraphFrameRight: input.paragraphFrameRight } : {}),
      ...(input.paragraphLineStep !== undefined ? { paragraphLineStep: input.paragraphLineStep } : {}),
      ...(input.paragraphSizeScale !== undefined ? { paragraphSizeScale: input.paragraphSizeScale } : {}),
      ...(rowStarts ? { rowStarts: [...rowStarts] } : {}) } });
  if (!result.compiled) throw new Error("The PDF compiler did not return an exact native candidate.");
  issuePdfCompiledTextPreview(result.compiled, selection);
  return result.compiled;
}
