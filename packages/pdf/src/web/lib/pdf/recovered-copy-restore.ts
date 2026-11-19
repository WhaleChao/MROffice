/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Restoring a PDF page's recovered copy (TRIAGE 41): the page body a
 * conflict's Reload set aside (lib/autorecover-store), put back in place --
 * asked first, and only over the same PDF file it was made on.
 *
 * A PDF page body describes edits OF ONE FILE: its `dataItemId` names the
 * file and its `source` record names the exact bytes the edits were made
 * against (resourceId, resourceVersionId, sha256; the window's own insert and
 * replace authority compares the same two, PdfPageWindow
 * pdfInsertPageIdentityRef). A save that wrote edits into the file
 * (patchback) leaves a new source: the copy's edits then describe bytes the
 * page no longer has, and putting them back would draw them over the wrong
 * text or drop them. So the copy is restored only while the page's file is
 * the one it was made on; otherwise it is refused, said why, and kept.
 */
import { PDF_CONTENT_TYPE } from "@mrnote/core/web/lib/native-page-content-types.ts";
import type { AutoRecoverCopy } from "@mrnote/core/web/lib/autorecover-store.ts";

type Body = Record<string, unknown>;

function record(value: unknown): Body | null {
  return value && typeof value === "object" && !Array.isArray(value) ? (value as Body) : null;
}

/** The file a PDF page body is of, as one comparable string; null for a body that is not a PDF page's. */
export function pdfBodyFileIdentity(body: unknown): string | null {
  const content = record(body);
  if (!content || content.type !== PDF_CONTENT_TYPE) return null;
  const source = record(content.source);
  return JSON.stringify({
    dataItemId: content.dataItemId ?? content.data_item_id ?? null,
    resourceId: source?.resourceId ?? null,
    resourceVersionId: source?.resourceVersionId ?? null,
    sha256: source?.sha256 ?? null,
  });
}

export type PdfRecoveredCopyRefusal = "not-pdf" | "file-changed" | "unsaved";

/** Why a copy cannot be put back over the page now, or null when it can. */
export function pdfRecoveredCopyRefusal(
  copy: Pick<AutoRecoverCopy, "payload">,
  pageBody: unknown,
  windowSaved: boolean,
): PdfRecoveredCopyRefusal | null {
  const kept = pdfBodyFileIdentity(copy.payload.content);
  const now = pdfBodyFileIdentity(pageBody);
  if (kept === null || now === null) return "not-pdf";
  if (kept !== now) return "file-changed";
  // The window's own unsaved edits would be replaced unasked.
  if (!windowSaved) return "unsaved";
  return null;
}

export type PdfRecoveredCopyOutcome =
  | { readonly outcome: "restored" }
  | { readonly outcome: "refused"; readonly reason: PdfRecoveredCopyRefusal }
  | { readonly outcome: "failed" };

/**
 * Put a copy back (after the person confirmed): refused copies are kept, a
 * restored one is deleted, a write that failed keeps it too.
 */
export async function restorePdfRecoveredCopy({
  copy, pageBody, windowSaved, write, deleteCopy,
}: {
  copy: AutoRecoverCopy;
  pageBody: unknown;
  windowSaved: boolean;
  /** The window's own write of the body (its save, then a reload): true when the page holds it. */
  write: (body: Body) => Promise<boolean>;
  deleteCopy: (copy: AutoRecoverCopy) => void;
}): Promise<PdfRecoveredCopyOutcome> {
  const reason = pdfRecoveredCopyRefusal(copy, pageBody, windowSaved);
  if (reason) return { outcome: "refused", reason };
  let written = false;
  try {
    written = await write(structuredClone(copy.payload.content));
  } catch {
    written = false;
  }
  if (!written) return { outcome: "failed" };
  deleteCopy(copy);
  return { outcome: "restored" };
}
