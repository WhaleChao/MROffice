/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const SUPPORTED_UNICODE_CMAP_NAME = /^Uni(?:GB|CNS|JIS|KS)-UCS2(?:-HW)?-[HV]$/;

export function isMraiNativePredefinedUnicodeCMapName(name: string): boolean {
  return SUPPORTED_UNICODE_CMAP_NAME.test(name);
}
