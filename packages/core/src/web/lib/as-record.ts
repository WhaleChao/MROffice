/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A JSON object's fields, or null for anything else (null, arrays,
 * primitives). Kept in a module of its own so a reader that needs only this
 * (the spreadsheet's shape text, through presentation/text-layout-css) does
 * not load the deck model and everything it imports.
 */
export function asRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}
