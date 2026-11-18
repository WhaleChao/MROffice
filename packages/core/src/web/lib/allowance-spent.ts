/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** How recently work must have stopped for the dialog's choice to cover it. */
export const ALLOWANCE_STOP_FRESH_MS = 15 * 60_000;
const ALLOWANCE_CODE = "ai_allowance_exhausted";

/**
 * Open the allowance dialog for a refusal that arrived some other way than a
 * 402 (lib/api.ts and lib/api-stream.ts announce those): a stream's error
 * frame, an Agent run that stopped part way. It also marks the held
 * allowance as spent and reads it again (AiAllowancePrimer).
 */
export function announceAllowanceSpent(details: Record<string, unknown> = {}): void {
  if (typeof window === "undefined") return;
  window.dispatchEvent(
    new CustomEvent("mrnote:plan-required", {
      detail: { code: ALLOWANCE_CODE, details, continuable: true },
    }),
  );
}

/**
 * A streamed turn's `error` frame that says the allowance ran out opens the
 * dialog, once, while it is news: a frame replayed from an old turn
 * (`at` more than a few minutes ago) does not.
 */
export function noticeAllowanceStreamError(
  event: { event: string; data: unknown; at?: number },
  now: number = Date.now(),
): boolean {
  if (event.event !== "error") return false;
  const data = event.data;
  if (!data || typeof data !== "object") return false;
  const record = data as Record<string, unknown>;
  const code = record.error_code ?? record.code;
  if (code !== ALLOWANCE_CODE) return false;
  if (typeof event.at === "number" && now - event.at > ALLOWANCE_STOP_FRESH_MS) return false;
  const details =
    record.details && typeof record.details === "object" && !Array.isArray(record.details)
      ? (record.details as Record<string, unknown>)
      : {};
  announceAllowanceSpent(details);
  return true;
}

