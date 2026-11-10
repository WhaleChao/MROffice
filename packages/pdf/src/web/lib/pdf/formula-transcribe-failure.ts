/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { ApiRequestError } from "@mrnote/core/web/lib/api.ts";

/**
 * Why a formula transcription (`/api/v1/ai/notebook/pdf-formula-transcribe`)
 * produced no LaTeX, as the two callers word it for the user:
 *
 * - `unreadable`: the model answered and found no readable formula in the
 *   crop (`formula_unreadable`). Retrying may help; typing the source always
 *   does. An empty result is the same answer.
 * - `unavailable`: the AI service gave no usable answer (`model_api_unavailable`,
 *   a gateway status, or no connection). Retrying later helps.
 * - `failed`: anything else (a local crop error, a refusal, a server fault).
 *
 * Reads the refusal's `code`/`status` fields rather than `instanceof`, so a
 * second module instance of the API client classifies the same.
 */
export type FormulaTranscribeFailure = "unreadable" | "unavailable" | "failed";

export function formulaTranscribeFailure(
  error: unknown,
): FormulaTranscribeFailure {
  if (!error || typeof error !== "object") return "failed";
  const { code, status } = error as Partial<
    Pick<ApiRequestError, "code" | "status">
  >;
  if (code === "formula_unreadable") return "unreadable";
  if (
    code === "model_api_unavailable" ||
    code === "network_unreachable" ||
    status === 502 ||
    status === 503 ||
    status === 504
  ) {
    return "unavailable";
  }
  return "failed";
}
