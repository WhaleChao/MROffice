/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { sanitizeHyperlink } from "@mrnote/core/web/lib/spreadsheet/conditional-format.ts";
import { parseInternalLink, parseNamedLink } from "@mrnote/core/web/lib/spreadsheet/sheet-interactions.ts";

/** Excel's hyperlink address limit (the server's `MAX_HYPERLINK_CHARS`). */
export const MAX_SPREADSHEET_OBJECT_HYPERLINK_CHARS = 2_083;

export function normalizeSpreadsheetObjectHyperlink(raw: string): string | null {
  const link = raw.trim();
  if (!link) return null;
  const stored =
    parseInternalLink(link) || parseNamedLink(link) !== null ? link : sanitizeHyperlink(link);
  if (!stored || stored.length > MAX_SPREADSHEET_OBJECT_HYPERLINK_CHARS) return null;
  if (!stored.startsWith("#")) {
    try {
      const url = new URL(stored);
      if (url.protocol !== "mailto:" && !url.host) return null;
    } catch {
      return null;
    }
  }
  return stored;
}
