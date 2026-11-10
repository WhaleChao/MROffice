/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeListTextSelectionEvidence } from "./engine/native/list-text-fidelity";

/** Item indexes are reader-local positions, not identities across revisions.
 * Keep the target the user actually hit. A source transition still requires
 * normal preflight against current bytes; this function grants no authority. */
export function rebindPdfNativeTextActivationSelection(
  selection: PdfNativeListTextSelectionEvidence | null,
  currentSourceSha256: string | null,
): PdfNativeListTextSelectionEvidence | null {
  if (!selection || !currentSourceSha256 || selection.sourceSha256 === currentSourceSha256) return selection;
  return Object.freeze({ ...selection, sourceSha256: currentSourceSha256,
    selectionRevision: selection.selectionRevision + 1 });
}
