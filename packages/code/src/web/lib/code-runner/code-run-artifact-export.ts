/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  normalizeCodePageRunTimeline,
  type PersistedCodeRunTimelineEntry,
} from "./code-page-output-persistence";
import { storedOutputsToNbformat } from "./output-persistence";

export const CODE_RUN_ARTIFACT_FORMAT = "mrnote.code-run";
export const CODE_RUN_ARTIFACT_VERSION = 1;
export const MAX_CODE_RUN_ARTIFACT_BYTES = 2 * 1024 * 1024;

export type CodeRunArtifactFailure = "empty" | "invalid" | "too_large";
export type CodeRunArtifactReceipt =
  | { readonly ok: true; readonly json: string; readonly outputCount: number }
  | { readonly ok: false; readonly reason: CodeRunArtifactFailure };

/**
 * Export one accepted run without copying source text, stdin text, credentials,
 * or runtime state. Only their stable hashes and bounded selection metadata are
 * carried alongside the normalized nbformat output objects.
 */
export function createCodeRunArtifactExport(
  entry: PersistedCodeRunTimelineEntry,
): CodeRunArtifactReceipt {
  try {
    const normalized = normalizeCodePageRunTimeline([entry])[0];
    if (!normalized) return { ok: false, reason: "invalid" };
    const outputs = storedOutputsToNbformat(normalized.output);
    if (outputs.length === 0) return { ok: false, reason: "empty" };
    const json = `${JSON.stringify(
      {
        format: CODE_RUN_ARTIFACT_FORMAT,
        version: CODE_RUN_ARTIFACT_VERSION,
        run: {
          completedAt: normalized.completedAt,
          completedAtIso: new Date(normalized.completedAt).toISOString(),
          language: normalized.language,
          providerId: normalized.providerId,
          sourceSha256: normalized.sourceSha256,
          outcome: normalized.outcome,
          exitCode: normalized.exitCode ?? null,
          input: normalized.input,
          navigation: normalized.navigation,
        },
        outputs,
      },
      null,
      2,
    )}\n`;
    const bytes = new TextEncoder().encode(json).byteLength;
    if (bytes > MAX_CODE_RUN_ARTIFACT_BYTES) {
      return { ok: false, reason: "too_large" };
    }
    return { ok: true, json, outputCount: outputs.length };
  } catch {
    return { ok: false, reason: "invalid" };
  }
}
