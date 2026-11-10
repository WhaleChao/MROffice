/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** How many revisions a chain may stack on its head before they are
 * squashed into one: well under the parser's 32 sections over any head the
 * server lets stand (it consolidates at 24). */
export const PDF_NATIVE_CHAIN_SQUASH_DEPTH = 4;

export interface PdfNativeChainStep<T> {
  readonly plan: T;
  /** The bytes this step was planned on (its command names them). */
  readonly sourceSha256: string;
}

/** The engine reads a chain needs (the planners' own, or their remote
 * callers in the browser). */
export interface PdfNativeChainEngine {
  readonly sectionCount: (bytes: Uint8Array) => number | Promise<number>;
  readonly squash: (base: Uint8Array, candidate: Uint8Array) => Uint8Array | null | Promise<Uint8Array | null>;
  readonly sha256: (bytes: Uint8Array) => Promise<string>;
  /** Called for every set of bytes the chain plans on, before it plans on
   * them: the browser's planners take only bytes registered as a working
   * head of the open file. */
  readonly adopt?: (sha256: string) => void;
}












export async function planPdfNativeRevisionChain<
  T extends { readonly prepared: { readonly candidateBytes: Uint8Array; readonly candidateSha256: string } },
>(
  engine: PdfNativeChainEngine,
  headBytes: Uint8Array,
  count: number,
  planStep: (index: number, sourceBytes: Uint8Array, sourceSha256: string) => Promise<T>,
): Promise<{ readonly steps: readonly PdfNativeChainStep<T>[]; readonly candidateBytes: Uint8Array;
  readonly candidateSha256: string }> {
  if (!Number.isSafeInteger(count) || count < 1) throw new Error("A PDF edit chain needs at least one step.");
  const headDepth = await engine.sectionCount(headBytes);
  const steps: PdfNativeChainStep<T>[] = [];
  let bytes = headBytes;
  let sha256 = await engine.sha256(headBytes);
  const squash = async () => {
    if ((await engine.sectionCount(bytes)) - headDepth < 2) return;
    const squashed = await engine.squash(headBytes, bytes);
    if (!squashed) throw new Error("The PDF edit's revisions could not be written as one.");
    bytes = squashed;
    sha256 = await engine.sha256(squashed);
    engine.adopt?.(sha256);
  };
  for (let index = 0; index < count; index += 1) {
    const plan = await planStep(index, bytes, sha256);
    steps.push(Object.freeze({ plan, sourceSha256: sha256 }));
    bytes = plan.prepared.candidateBytes;
    sha256 = plan.prepared.candidateSha256;
    engine.adopt?.(sha256);
    if (index < count - 1 && (await engine.sectionCount(bytes)) - headDepth >= PDF_NATIVE_CHAIN_SQUASH_DEPTH) {
      await squash();
    }
  }
  await squash();
  return Object.freeze({ steps: Object.freeze(steps), candidateBytes: bytes, candidateSha256: sha256 });
}
