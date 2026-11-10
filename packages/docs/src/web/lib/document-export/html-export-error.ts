/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** A fail-closed boundary error for standalone HTML serialization. */
export class OfflineHtmlExportError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "OfflineHtmlExportError";
  }
}
