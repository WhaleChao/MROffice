/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export function pdfDevErrorText(error: unknown, limit = 600): string {
  try {
    if (error instanceof Error) return `${error.name}: ${String(error.message).slice(0, limit)}`;
    if (typeof error === "string") return error.slice(0, limit);
    return String(JSON.stringify(error) ?? error).slice(0, limit);
  } catch {
    return "unprintable error";
  }
}
