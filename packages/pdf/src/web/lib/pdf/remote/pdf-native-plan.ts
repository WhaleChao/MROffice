/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * One engine command planned without writing anything: the record the
 * server would replay and the exact candidate bytes it produces. The working
 * document stages the candidate at once (the edit shows on its bytes) and
 * plans the command again against the saved version when it publishes it
 * (lib/pdf/native-working-session.ts, engine operations).
 */
import { encodePdfNativeWire } from "@mrnote/core/web/lib/engine-host/pdf-native-wire.ts";
import type { PdfEngineCommand, PdfEngineResult, PdfEngineSourceIdentity } from "../engine/contracts";
import type { PdfSharedStateRecord } from "../engine/shared-state-bridge";
import { sha256PdfBytes } from "../engine/native/transaction-error";
import { callPdfNativeEngine, createPdfNativeOffload, decodePdfNativeAnswer, encodePdfNativeCommand } from "../pdf-native-remote";

export type PdfNativeEnginePlan =
  | Readonly<{ status: "planned"; record: PdfSharedStateRecord; expectedCandidateSha256: string;
      /** Present when the head's bytes were given (the candidate is rebuilt on them). */
      candidateBytes: Uint8Array | null;
      /** Planned in this browser's Worker: where its time went there. */
      phases?: Readonly<{ digest: number; checkpoint: number; plan: number; shared: boolean; prepared: boolean }> }>
  | Readonly<{ status: "refused"; result: PdfEngineResult }>;

export async function planPdfNativeEngineCommand(input: {
  readonly pageId: string;
  /** The head the command is planned on: the saved version's ids with the
   * bytes' own digest (a working head is registered with the same ids). */
  readonly source: PdfEngineSourceIdentity;
  /** The head's bytes, when the caller holds them: uploaded if the server
   * does not, and the candidate is rebuilt on them. Without them the server
   * reads its stored version and only the record comes back. */
  readonly sourceBytes?: Uint8Array | null;
  readonly command: PdfEngineCommand;
  readonly options?: Readonly<Record<string, unknown>>;
  readonly signal?: AbortSignal;
}): Promise<PdfNativeEnginePlan> {
  const held = input.sourceBytes ? Uint8Array.from(input.sourceBytes) : null;
  try {
    if (held && await sha256PdfBytes(held) !== input.source.sourceSha256) throw new Error("The PDF head does not match its digest.");
    const offloaded = createPdfNativeOffload();
    const options = input.options
      ? (await encodePdfNativeWire({ ...input.options }, null, undefined, offloaded.offload)) as Record<string, unknown>
      : {};
    const planned = await callPdfNativeEngine("plan", {
      source: { pageId: input.pageId, ...input.source },
      // A long text the command carries travels beside the JSON (1 MiB route bound).
      command: await encodePdfNativeCommand({ ...input.command, resourceId: input.source.resourceId,
        resourceVersionId: input.source.resourceVersionId, sourceSha256: input.source.sourceSha256 }, offloaded),
      options,
    }, { signal: input.signal, workingBytes: () => held, offloaded });
    if (planned.status !== "planned") return Object.freeze({ status: "refused" as const, result: planned.result });
    let candidateBytes: Uint8Array | null = null;
    if (held) {
      if (!planned.candidate) throw new Error("The PDF engine returned no candidate for this edit.");
      candidateBytes = await decodePdfNativeAnswer<Uint8Array>(planned.candidate, { sha256: input.source.sourceSha256, bytes: held });
      if (await sha256PdfBytes(candidateBytes) !== planned.expectedCandidateSha256) {
        candidateBytes.fill(0);
        throw new Error("The PDF engine's candidate does not match its digest.");
      }
    }
    return Object.freeze({ status: "planned" as const, record: planned.record,
      expectedCandidateSha256: planned.expectedCandidateSha256, candidateBytes });
  } finally {
    held?.fill(0);
  }
}
