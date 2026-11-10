/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Whether this browser runs on Apple's platforms (macOS, iOS, iPadOS), for
 * shortcut labels and Option-key behaviour. Its own module, free of any
 * editor engine: the document editor asks only this, and importing it from
 * the code editor's chord helpers shipped CodeMirror with every document.
 */
export function isMacPlatform(): boolean {
  if (typeof navigator === "undefined") return false;
  const platform =
    (navigator as Navigator & { userAgentData?: { platform?: string } })
      .userAgentData?.platform ||
    navigator.platform ||
    navigator.userAgent ||
    "";
  return /mac|iphone|ipad|ipod/i.test(platform);
}
