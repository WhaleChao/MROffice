/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The text outline worker (office-text-outline.worker.ts), started. Its own
 * module so a host that must not carry PathKit swaps it whole: the document
 * verification host aliases it to outline-worker-start.verification.ts
 * (scripts/build-document-verification-host.mjs), whose source graph cannot
 * place PathKit's script and wasm, which the worker loads as plain assets.
 */
export function startOutlineWorker(): Worker {
  // Classic: it loads PathKit's script with importScripts.
  return new Worker(new URL("./office-text-outline.worker.ts", import.meta.url), { name: "mrnote-text-outline" });
}
