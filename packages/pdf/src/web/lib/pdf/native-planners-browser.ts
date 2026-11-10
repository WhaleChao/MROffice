/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The editor's side of native-planners.worker.ts: one Worker for the page,
 * several calls in flight, each answered by its request id. Unavailable
 * (no Worker, a crash, a timeout, an argument structured clone cannot
 * carry) throws PdfNativeBrowserPlanUnavailable and the caller asks the
 * engine host instead; a planner's own refusal or answer is final.
 */
import { pdfNativeBrowserPrepareEnabled } from "./native-text-browser-prepare";
import type { PdfNativePlannerWorkerRequest, PdfNativePlannerWorkerResponse } from "./native-planners.worker";
import { askPdfNativeTextWorker, PdfNativeTextWorkerUnavailable } from "./native-text-worker";

export class PdfNativeBrowserPlannerUnavailable extends Error {
  constructor(reason: string) {
    super(reason);
    this.name = "PdfNativeBrowserPlannerUnavailable";
  }
}

/** The answer a planner threw, as it threw it (rethrown by the caller). */
export class PdfNativeBrowserPlannerThrew extends Error {
  constructor(readonly thrown: Readonly<{ name: string; message: string; code?: string; resourceIssue?: unknown }>) {
    super(thrown.message);
    this.name = "PdfNativeBrowserPlannerThrew";
  }
}

const BROWSER_PLANNER_TIMEOUT_MS = 60_000;
/** Planners the Worker does not hold (engine-host/engines/pdf-native-functions.ts
 * minus native-planners.worker.ts): asked of the engine host straight away. */
const SERVER_ONLY_PLANNERS = new Set([
  "planMraiNativePageBackground", "preparePdfNativeAuthorTextRecovery", "inspectPdfTaggedTextTargets", "createPdfTaggedTextOperation",
  "verifyPdfTaggedTextOutput",
]);

/** Reads nobody waits on: answered by the background Worker, so they never
 * stand in front of an edit's plan (native-text-worker.ts). */
const BACKGROUND_PLANNERS = new Set(["readMraiNativeManagedOutline", "inspectPdfNativeFontReadingRepair"]);

export function pdfNativeBrowserPlannerEnabled(fn: string): boolean {
  return pdfNativeBrowserPrepareEnabled() && !SERVER_ONLY_PLANNERS.has(fn) && typeof Worker !== "undefined";
}

export async function callPdfNativePlannerInBrowser(fn: string, args: readonly unknown[], signal?: AbortSignal): Promise<unknown> {
  if (!pdfNativeBrowserPlannerEnabled(fn)) throw new PdfNativeBrowserPlannerUnavailable("disabled");
  const request: PdfNativePlannerWorkerRequest = { requestId: crypto.randomUUID(), fn, args };
  let response: PdfNativePlannerWorkerResponse;
  try {
    // The page's one PDF text Worker (native-text-worker.ts), with what it
    // has read of the head; a planner this slow replaces it for the next.
    response = await askPdfNativeTextWorker<PdfNativePlannerWorkerResponse>("planner", request,
      { timeoutMs: BROWSER_PLANNER_TIMEOUT_MS, signal,
        lane: BACKGROUND_PLANNERS.has(fn) ? "background" : "interactive" });
  } catch (error) {
    if (error instanceof PdfNativeTextWorkerUnavailable) throw new PdfNativeBrowserPlannerUnavailable(error.message);
    throw error;
  }
  if (response.status === "failed") throw new PdfNativeBrowserPlannerUnavailable(response.reason);
  if (response.status === "thrown") throw new PdfNativeBrowserPlannerThrew(response.error);
  return response.value;
}
