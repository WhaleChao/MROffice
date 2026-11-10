/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// Inspects a PDF's source risk profile (lib/pdf/save-policy.ts) off the main
// thread: pdf-lib parses the whole file and walks its object graph there,
// seconds of the reader's thread on a long book (the owner's 132-page

import { inspectPdfRiskProfileBytes, type PdfSourceRiskProfile } from "./save-policy";

type Request = { id: number; bytes: ArrayBuffer };
type Response = { ready: true } | { id: number; profile?: PdfSourceRiskProfile; error?: string };

const scope = self as unknown as { postMessage(value: Response): void; onmessage: ((event: MessageEvent<Request>) => void) | null };

scope.postMessage({ ready: true });
scope.onmessage = (event) => {
  const request = event.data;
  if (!request || typeof request.id !== "number") return;
  const bytes = new Uint8Array(request.bytes);
  inspectPdfRiskProfileBytes(bytes).then(
    (profile) => scope.postMessage({ id: request.id, profile }),
    (error) => scope.postMessage({ id: request.id, error: error instanceof Error ? error.message : String(error) }),
  );
};
