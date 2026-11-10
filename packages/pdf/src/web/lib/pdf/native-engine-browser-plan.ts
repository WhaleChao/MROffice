/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfEngineCommand, PdfEngineSourceIdentity } from "./engine/contracts";
import { pdfNativeBrowserPrepareEnabled } from "./native-text-browser-prepare";
import type { PdfNativeEnginePlan } from "./remote/pdf-native-plan";
import type { PdfNativeEnginePlanWorkerRequest, PdfNativeEnginePlanWorkerResponse } from "./native-engine-plan.worker";
import type { PdfNativeListTextSelectionEvidence } from "./engine/native/list-text-fidelity";
import { askPdfNativeTextWorker, prewarmPdfNativeTextWorker, PdfNativeTextWorkerUnavailable } from "./native-text-worker";

/** The browser could not plan (no Worker, a Worker that failed or timed
 * out): the caller asks the engine host, as before. */
export class PdfNativeBrowserPlanUnavailable extends Error {
  constructor(reason: string) {
    super(reason);
    this.name = "PdfNativeBrowserPlanUnavailable";
  }
}

const BROWSER_PLAN_TIMEOUT_MS = 45_000;

/** Ask the page's one PDF text Worker (native-text-worker.ts) on its plan
 * channel: what it has read of a head (the paragraph inspection, the
 * stream of the file read) serves every later plan on the same bytes, and
 * every planner's too. A Worker that times out or fails is replaced, its
 * requests failed. */
async function ask(request: PdfNativeEnginePlanWorkerRequest, signal?: AbortSignal): Promise<PdfNativeEnginePlanWorkerResponse> {
  try {
    return await askPdfNativeTextWorker<PdfNativeEnginePlanWorkerResponse>("plan", request,
      { timeoutMs: BROWSER_PLAN_TIMEOUT_MS, signal, transfer: [request.sourceBytes] });
  } catch (error) {
    if (error instanceof PdfNativeTextWorkerUnavailable) throw new PdfNativeBrowserPlanUnavailable(error.message);
    throw error;
  }
}





export function prewarmPdfNativeEnginePlanWorker(): void {
  if (!pdfNativeBrowserPlanEnabled()) return;
  prewarmPdfNativeTextWorker();
}

/** Read the document a box was opened on into the plan Worker now (its
 * paragraph inspection), so the first plan of that box's paragraph -- a
 * centred line typed on, a box's formats -- does not read it first. */
export function prewarmPdfNativeParagraphPlans(sourceBytes: Uint8Array,
  selection?: PdfNativeListTextSelectionEvidence | null): void {
  if (!pdfNativeBrowserPlanEnabled() || typeof Worker === "undefined") return;
  const copy = Uint8Array.from(sourceBytes);
  void ask({ requestId: crypto.randomUUID(), kind: "warm-paragraph", sourceBytes: copy.buffer as ArrayBuffer,
    ...(selection ? { selection: JSON.parse(JSON.stringify(selection)) as PdfNativeListTextSelectionEvidence } : {}),
    source: { resourceId: "", resourceVersionId: "", sourceSha256: "" }, command: null, options: {} }).catch(() => undefined);
}

/** Whether a command is planned in this browser first (the same switch as
 * the preparation's: NEXT_PUBLIC_MRAI_PDF_BROWSER_PREPARE). */
export function pdfNativeBrowserPlanEnabled(): boolean {
  return pdfNativeBrowserPrepareEnabled();
}

/** Plan `command` on `sourceBytes` (the head, whose digest is
 * `source.sourceSha256`) in this browser's Worker. Resolves as the engine
 * host's plan does; rejects with PdfNativeBrowserPlanUnavailable when the
 * browser could not plan at all. */
export async function planPdfNativeEngineCommandInBrowser(input: Readonly<{
  source: PdfEngineSourceIdentity;
  sourceBytes: Uint8Array;
  command: PdfEngineCommand;
  options?: Readonly<Record<string, unknown>>;
  signal?: AbortSignal;
}>): Promise<PdfNativeEnginePlan> {
  if (!pdfNativeBrowserPlanEnabled()) throw new PdfNativeBrowserPlanUnavailable("disabled");
  const copy = Uint8Array.from(input.sourceBytes);
  const response = await ask({
    requestId: crypto.randomUUID(), sourceBytes: copy.buffer as ArrayBuffer, source: input.source,
    command: JSON.parse(JSON.stringify(input.command)) as unknown,
    options: { ...(input.options ?? {}) },
  }, input.signal);
  if (response.status === "failed") throw new PdfNativeBrowserPlanUnavailable(response.reason);
  if (response.status === "refused") return Object.freeze({ status: "refused" as const, result: response.result });
  return Object.freeze({ status: "planned" as const, record: response.record,
    expectedCandidateSha256: response.expectedCandidateSha256, candidateBytes: response.candidateBytes,
    ...(response.phases ? { phases: response.phases } : {}) });
}
