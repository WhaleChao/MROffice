/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Portable, bounded export receipt for one persisted Code execution output.
 *
 * The payload deliberately contains nbformat output objects rather than the
 * renderer's session shape. Existing persistence adapters remain the single
 * authority for truncation, MIME validation and ordered imported outputs.
 */

import type { RunResult } from "@mrnote/core/web/lib/code-runner/js-runner.ts";
import {
  normalizeNbformatOutputsSnapshot,
  runResultToStoredOutputs,
  storedOutputsToNbformat,
} from "./output-persistence";

export const CODE_OUTPUT_EXPORT_FORMAT = "mrnote.code-output";
export const CODE_OUTPUT_EXPORT_VERSION = 1;
export const CODE_OUTPUT_EXPORT_FILENAME = "cell-output.json";
export const MAX_CODE_OUTPUT_EXPORT_BYTES = 2 * 1024 * 1024;

export type CellOutputExportFailure = "empty" | "invalid" | "too_large";

export type CellOutputExportReceipt =
  | { ok: true; json: string; outputCount: number }
  | { ok: false; reason: CellOutputExportFailure };

const INVALID_EXPORT_FIELD = Symbol("invalid-export-field");

function ownExportField(
  record: object,
  key: string,
): unknown | typeof INVALID_EXPORT_FIELD {
  try {
    const descriptor = Object.getOwnPropertyDescriptor(record, key);
    if (!descriptor) return undefined;
    return "value" in descriptor ? descriptor.value : INVALID_EXPORT_FIELD;
  } catch {
    return INVALID_EXPORT_FIELD;
  }
}

function utf8ByteLength(value: string): number | null {
  try {
    if (typeof TextEncoder === "undefined") return null;
    return new TextEncoder().encode(value).byteLength;
  } catch {
    return null;
  }
}

function supportsNativeRichOutput(value: unknown): boolean {
  if (value === undefined) return true;
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const kind = ownExportField(value, "kind");
  if (kind === INVALID_EXPORT_FIELD) return false;
  return kind === "table" || kind === "image" || kind === "html";
}

/**
 * Produce one complete output artifact without trusting live renderer data.
 * Imported ordered outputs are normalized directly; native results take the
 * existing RunResult -> StoredCellOutput -> nbformat path. Any unsupported rich
 * renderer kind fails closed rather than being silently mislabeled or dropped.
 */
export function createCellOutputExport(
  output: unknown,
): CellOutputExportReceipt {
  try {
    if (!output || typeof output !== "object" || Array.isArray(output)) {
      return { ok: false, reason: "invalid" };
    }
    const record = output as Record<string, unknown>;
    const notebookOutputs = ownExportField(record, "notebookOutputs");
    const rich = ownExportField(record, "rich");
    if (
      notebookOutputs === INVALID_EXPORT_FIELD ||
      rich === INVALID_EXPORT_FIELD
    ) {
      return { ok: false, reason: "invalid" };
    }
    let outputs: readonly Record<string, unknown>[];

    if (notebookOutputs !== undefined) {
      const normalized = normalizeNbformatOutputsSnapshot(
        notebookOutputs,
      );
      if (!normalized) return { ok: false, reason: "invalid" };
      outputs = normalized;
    } else {
      if (!supportsNativeRichOutput(rich)) {
        return { ok: false, reason: "invalid" };
      }
      const stored = runResultToStoredOutputs(output as RunResult);
      if (!stored) return { ok: false, reason: "empty" };
      outputs = storedOutputsToNbformat(stored);
    }

    if (outputs.length === 0) return { ok: false, reason: "empty" };
    const json = `${JSON.stringify(
      {
        format: CODE_OUTPUT_EXPORT_FORMAT,
        version: CODE_OUTPUT_EXPORT_VERSION,
        outputs,
      },
      null,
      2,
    )}\n`;
    const bytes = utf8ByteLength(json);
    if (bytes === null || bytes > MAX_CODE_OUTPUT_EXPORT_BYTES) {
      return { ok: false, reason: "too_large" };
    }
    return { ok: true, json, outputCount: outputs.length };
  } catch {
    return { ok: false, reason: "invalid" };
  }
}
