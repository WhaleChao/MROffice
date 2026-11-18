/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { MAX_MRAI_NATIVE_SOURCE_BYTES, PdfNativeTransactionError } from "./engine/native/transaction-error";
import { MRAI_NATIVE_LIST_TEXT_MISSING_GLYPH_MESSAGE } from "./engine/native/vocabulary";
import { pdfNativeResourceIssueMessage, readPdfNativeResourceIssue } from "./engine/native/resource-issue";
import type {
  PdfNativeTextPreparation,
  PdfNativeTextPreparationRequest,
  PdfNativeTextPreparationResponse,
} from "./native-text-edit-preparation";
import {
  askPdfNativeTextWorker,
  onPdfNativeTextWorkerDiscard,
  PdfNativeTextWorkerUnavailable,
  prewarmPdfNativeTextWorker,
  retirePdfNativeTextWorker,
} from "./native-text-worker";

/** The browser could not do the work (no Worker, the Worker failed or
 * answered nothing in time): the server's engine host is asked instead.
 * Never a refusal of the selection itself, which is final either way. */
export class PdfNativeBrowserPrepareUnavailable extends Error {
  constructor(reason: string) {
    super(reason);
    this.name = "PdfNativeBrowserPrepareUnavailable";
  }
}

/** One proof of a long document on a slow machine; past it the server
 * answers (its engine is warm for the file by then). */
const BROWSER_PREPARE_TIMEOUT_MS = 45_000;
/** Workers kept warm between requests (their parsed document stays). */

export function pdfNativeBrowserPrepareEnabled(): boolean {
  const flag = process.env.NEXT_PUBLIC_MRAI_PDF_BROWSER_PREPARE;
  
  // asks for this one: jsdom's Worker answers nothing, and a suite that
  // stages text waited out the browser's timeout before falling back.
  if (process.env.NODE_ENV === "test" && flag !== "1") return false;
  return typeof Worker === "function" &&
    typeof crypto !== "undefined" && typeof crypto.randomUUID === "function" &&
    flag !== "0";
}

/** An answer that came after its request was given up: its bytes cleared. */
onPdfNativeTextWorkerDiscard("prepare", (payload) => {
  const result = (payload as { result?: PdfNativeTextPreparation } | null)?.result;
  result?.blankedSource?.fill(0);
  result?.compiled?.prepared.candidateBytes.fill(0);
});

/** Start the page's PDF text Worker before the first line is clicked (its
 * modules load meanwhile). */
export function warmPdfNativeBrowserPrepare(): void {
  if (!pdfNativeBrowserPrepareEnabled()) return;
  prewarmPdfNativeTextWorker();
}

/** Read the document into the Worker's read session ahead of the first line
 * opened on it, so that line is proved on it (1.3 s cold on the Attention
 * paper's page). */
export function retainPdfNativeBrowserPrepareSource(sourceBytes: Uint8Array,
  /** The Workers that read it: the proofs' (a press is proved there) and the
   * compiles' (a leave writes there), by default both. */
  lanes: readonly ("interactive" | "compile")[] = ["interactive", "compile"]): void {
  if (!pdfNativeBrowserPrepareEnabled() || !sourceBytes.length || sourceBytes.length > MAX_MRAI_NATIVE_SOURCE_BYTES) return;
  for (const lane of lanes) {
    const copy = Uint8Array.from(sourceBytes);
    void askPdfNativeTextWorker("prepare", { requestId: crypto.randomUUID(), kind: "retain", sourceBytes: copy.buffer } as never,
      { timeoutMs: BROWSER_PREPARE_TIMEOUT_MS, transfer: [copy.buffer], lane }).catch(() => undefined);
  }
}

export async function preparePdfNativeTextInBrowser(
  input: Omit<PdfNativeTextPreparationRequest, "requestId" | "sourceBytes"> & {
    readonly sourceBytes: Uint8Array;
    readonly signal: AbortSignal;
    /** A proof nobody waits on (an open editor's line proved again on the
     * head) goes to the background Worker, never ahead of a press. */
    readonly background?: boolean;
    /** A proof of the line a compile on these bytes has just read: asked of
     * the compiles' Worker, whose read session keeps that line's shows as
     * the compile read them (0.4 s for a paragraph of two thousand shows
     * that the proofs' Worker, reading it first, takes 4 s and more for). */
    readonly afterCompile?: boolean;
  },
): Promise<PdfNativeTextPreparation> {
  if (input.signal.aborted) throw new DOMException("Cancelled", "AbortError");
  if (!input.sourceBytes.length || input.sourceBytes.length > MAX_MRAI_NATIVE_SOURCE_BYTES) {
    throw new PdfNativeTransactionError("unsupported-document",
      "PDF text preparation exceeds the source budget.");
  }
  const ownedBytes = Uint8Array.from(input.sourceBytes);
  const request = {
    requestId: crypto.randomUUID(),
    sourceBytes: ownedBytes.buffer,
    selection: input.selection,
    managedBoxId: input.managedBoxId,
    sourceText: input.sourceText,
    ...(input.compile ? { compile: input.compile } : {}),
  } satisfies PdfNativeTextPreparationRequest;
  let response: PdfNativeTextPreparationResponse;
  try {
    // The page's one PDF text Worker (native-text-worker.ts), its read
    // session with the document it has read; requests answered in turn.
    // A compile in its own Worker: the proof of the line pressed next never
    // waits behind it (1.4 s on the IPCC report).
    response = await askPdfNativeTextWorker<PdfNativeTextPreparationResponse>("prepare", request,
      { timeoutMs: BROWSER_PREPARE_TIMEOUT_MS, signal: input.signal, transfer: [ownedBytes.buffer],
        lane: input.compile || input.afterCompile ? "compile" : input.background ? "background" : "interactive" });
  } catch (error) {
    if (ownedBytes.byteLength) ownedBytes.fill(0);
    if (error instanceof PdfNativeTextWorkerUnavailable) throw new PdfNativeBrowserPrepareUnavailable(error.message);
    throw error;
  }
  if ("failure" in response) {
    if (response.failure === "worker-cancelled") throw new DOMException("PDF preparation superseded", "AbortError");
    if (response.failure === "worker-failed") {
      // A preparation that failed the Worker's own way: replaced for the next.
      retirePdfNativeTextWorker("worker-failed");
      throw new PdfNativeBrowserPrepareUnavailable("worker-failed");
    }
    const issue = readPdfNativeResourceIssue(response.resourceIssue);
    throw new PdfNativeTransactionError(
      response.failure,
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
  if (response.sourceSha256 !== input.selection.sourceSha256 || !response.result?.proof) {
    response.result?.blankedSource?.fill(0);
    response.result?.compiled?.prepared.candidateBytes.fill(0);
    throw new PdfNativeBrowserPrepareUnavailable("source-changed");
  }
  return response.result;
}
