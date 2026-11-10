/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The native planner's public vocabulary that the editor shares with it: the
 * refusal class (its closed codes), the byte budgets, and the SHA-256 every
 * source and candidate is named by. No PDF knowledge: the planners (server-
 * only, `content-stream-transaction.ts` and its kin) re-export these, and the
 * browser imports them from here.
 */

import type { PdfNativeResourceIssue } from "./resource-issue";
import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";
import { sha256HexSync } from "./sha256-sync";

export const MAX_MRAI_NATIVE_SOURCE_BYTES = PDF_EDITABLE_MAX_BYTES;
export const MAX_MRAI_NATIVE_EMBEDDED_IMAGE_BYTES = 2 * 1024 * 1024;

export type PdfNativeTransactionErrorCode =
  | "invalid-request"
  | "source-mismatch"
  | "precondition-mismatch"
  | "unsupported-document"
  | "verification-failed";

export class PdfNativeTransactionError extends Error {
  constructor(
    readonly code: PdfNativeTransactionErrorCode,
    message: string,
    readonly resourceIssue?: PdfNativeResourceIssue,
  ) {
    super(message);
    this.name = "PdfNativeTransactionError";
  }
}

/** Up to this size a digest is computed in this thread (sha256-sync): the
 * planners hash an object's bytes, an operator's span, by the thousand, and
 * each answer of `crypto.subtle` comes back from another thread. A whole
 * PDF still goes to `crypto.subtle`. */
const SYNC_DIGEST_MAX_BYTES = 256 * 1024;

export async function sha256PdfBytes(bytes: Uint8Array): Promise<string> {
  if (bytes.length <= SYNC_DIGEST_MAX_BYTES) return sha256HexSync(bytes);
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) throw new PdfNativeTransactionError("unsupported-document", "SHA-256 is unavailable.");
  const owned = Uint8Array.from(bytes);
  const digest = await subtle.digest("SHA-256", owned);
  return Array.from(new Uint8Array(digest), (value) =>
    value.toString(16).padStart(2, "0"),
  ).join("");
}
