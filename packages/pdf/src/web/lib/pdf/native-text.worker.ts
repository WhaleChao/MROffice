/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The one PDF text Worker: a line's proof and its typing compiled
 * (preparation), the engine's plans of a box's edits (plan), and the named
 * planners (a move, a decoration, a delete, the vector tools) -- one thread,
 * one copy of what each has read of the document (the inspection kept for a
 * head, docs/pdf/text-edit-unified-model.md). Three Workers each read the
 * same head again, and one could be cold where another was warm: a box
 * dragged planned its move on a cold read (4 s on the Attention paper's first
 * page) while the plan Worker held the head warm.
 *
 * A message is `{ channel, payload }`; its answer is `{ channel, payload }`.
 */
import { createPdfNativeTextPreparationLane } from "./native-text-edit-preparation.worker";
import { planAndZeroizePdfNativeEngineCommand, type PdfNativeEnginePlanWorkerRequest } from "./native-engine-plan.worker";
import { answerPdfNativePlannerRequest, isPdfNativePlannerWorkerRequest } from "./native-planners.worker";
import type { PdfNativeTextPreparationRequest } from "./native-text-edit-preparation";
import type { PdfNativeTextWorkerChannel } from "./native-text-worker";

type Envelope = Readonly<{ channel: PdfNativeTextWorkerChannel; payload: unknown }>;

const scope = globalThis as unknown as {
  document?: unknown;
  onmessage: ((event: MessageEvent<Envelope>) => void) | null;
  postMessage?: (message: Envelope, transfer?: Transferable[]) => void;
};

if (scope.document === undefined && typeof scope.postMessage === "function") {
  const post = (channel: PdfNativeTextWorkerChannel, payload: unknown, transfer: Transferable[] = []) => {
    try {
      scope.postMessage!({ channel, payload }, transfer);
    } catch (error) {
      // An answer structured clone cannot carry (a planner's function, a
      // cycle): the caller asks the engine host instead.
      const requestId = (payload as { requestId?: unknown } | null)?.requestId;
      if (typeof requestId === "string") scope.postMessage!({ channel, payload: { requestId, status: "failed",
        reason: `uncloneable-answer:${error instanceof Error ? error.message : String(error)}` } });
    }
  };
  const preparation = createPdfNativeTextPreparationLane();
  // The engine's plans one at a time, an edit's own ahead of the previews
  // and warm reads still waiting: interleaved on this one thread they only
  // shared it, and a box's commit finished with the two previews before it
  
  // not started.
  const plans: { request: PdfNativeEnginePlanWorkerRequest; urgent: boolean }[] = [];
  let planning = false;
  const urgentPlan = (request: PdfNativeEnginePlanWorkerRequest) => {
    const id = (request.command as { transactionId?: unknown } | null)?.transactionId;
    return request.kind !== "warm-paragraph" && !(typeof id === "string" && id.startsWith("preview:"));
  };
  const nextPlan = () => {
    if (planning) return;
    const next = plans.shift();
    if (!next) return;
    planning = true;
    void planAndZeroizePdfNativeEngineCommand(next.request).then((response) =>
      post("plan", response, response.status === "planned" ? [response.candidateBytes.buffer as ArrayBuffer] : []))
      .finally(() => { planning = false; nextPlan(); });
  };
  scope.onmessage = (event) => {
    const { channel, payload } = event.data ?? ({} as Envelope);
    const cancel = payload as { cancel?: unknown; requestId?: unknown } | null;
    if (cancel?.cancel === true) {
      if (channel === "plan") {
        const at = plans.findIndex((entry) => entry.request.requestId === cancel.requestId);
        if (at >= 0) plans.splice(at, 1);
      }
      return;
    }
    if (channel === "prepare") {
      const retain = payload as { kind?: string; requestId?: string; sourceBytes?: ArrayBuffer };
      if (retain?.kind === "retain" && retain.sourceBytes instanceof ArrayBuffer) {
        void preparation.retain(retain.sourceBytes).then((kept) =>
          post(channel, { requestId: retain.requestId ?? "", failure: kept ? "worker-cancelled" : "worker-failed" }));
        return;
      }
      void preparation.prepare(payload as PdfNativeTextPreparationRequest)
        .then(({ response, transfer }) => post(channel, response, transfer),
          () => post(channel, { requestId: (payload as { requestId?: string }).requestId ?? "", failure: "worker-failed" }));
      return;
    }
    if (channel === "plan") {
      const request = payload as PdfNativeEnginePlanWorkerRequest;
      const urgent = urgentPlan(request);
      const at = urgent ? plans.findIndex((entry) => !entry.urgent) : -1;
      plans.splice(at < 0 ? plans.length : at, 0, { request, urgent });
      nextPlan();
      return;
    }
    if (channel === "planner" && isPdfNativePlannerWorkerRequest(payload)) {
      void answerPdfNativePlannerRequest(payload).then((response) => post(channel, response), (error) =>
        post(channel, { requestId: payload.requestId, status: "failed",
          reason: error instanceof Error ? error.message : String(error) }));
    }
  };
}
