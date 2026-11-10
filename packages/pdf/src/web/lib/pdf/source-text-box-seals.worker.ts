/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// Reads every managed text box's appearance seal of one PDF off the main
// thread (lib/pdf/source-text-box.ts): loading a 4 MB PDF in pdf-lib held
// the reader's thread ~0.3 s at every text-box save, the editor waiting on it.
import { scanPdfSourceTextBoxAppearanceSeals, type PdfSourceTextBoxSealIndex } from "./source-text-box-seal-scan";

type Request = { id: number; sha256: string; bytes: ArrayBuffer };
type Response = { id: number; index?: PdfSourceTextBoxSealIndex; error?: string };

const toHex = (buffer: ArrayBuffer) => [...new Uint8Array(buffer)].map((byte) => byte.toString(16).padStart(2, "0")).join("");

self.onmessage = async (event: MessageEvent<Request>) => {
  const { id, sha256, bytes } = event.data;
  const post = (response: Response) => (self as unknown as Worker).postMessage(response);
  try {
    // Only bytes of exactly the claimed source are read for it.
    if (toHex(await crypto.subtle.digest("SHA-256", bytes)) !== sha256) {
      post({ id, error: "source-mismatch" });
      return;
    }
    post({ id, index: await scanPdfSourceTextBoxAppearanceSeals(new Uint8Array(bytes)) });
  } catch (error) {
    post({ id, error: error instanceof Error ? error.message : String(error) });
  }
};
