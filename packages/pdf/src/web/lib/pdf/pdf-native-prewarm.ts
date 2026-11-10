/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Get a PDF head ready on the engine host before the reader's first click
 * (`pdf-native/prewarm`): a worker holding the bytes, pdf.js and its assets
 * loaded, the document's shared reading built and the named pages read. No
 * side effect; a hint only -- it never throws and never blocks an edit.
 *
 *   prewarmPdfNativeSource({ engineSource, sourceSha256, pages: [3, 4], signal })
 *
 * - Idempotent: a page already warmed for this head (in this tab, within
 *   WARM_TTL_MS) is not asked again; the same request in flight is shared.
 * - Cancellable: `signal` stops waiting (the host finishes its read anyway;
 *   a later prewarm or prepare finds it done).
 * - Per page or per document: `pages` (at most 8, 1-based) names what is on
 *   screen; none = the document only. A caller batching visible lines asks
 *   for their pages as the viewport settles.
 * - `workingBytes` (optional) lets an unsaved head the server never saw be
 *   sent once, as every pdf-native call does; without it such a head is
 *   simply not warmed.
 */
import type { PdfNativeEngineSourceRef } from "./pdf-native-remote";
import { callPdfNativeEngine } from "./pdf-native-remote";

/** A worker that read a head keeps it until recycled; past this a page is
 * asked again (a cheap answer when it is still warm). */
export const WARM_TTL_MS = 10 * 60_000;
const MAX_PAGES = 8;
const MAX_HEADS = 32;

export type PdfNativePrewarmStatus = "warm" | "already-warm" | "aborted" | "failed";

export interface PdfNativePrewarmResult {
  readonly status: PdfNativePrewarmStatus;
  readonly pages: readonly number[];
  readonly ms: number;
}

const warmed = new Map<string, { at: number; pages: Set<number> }>();
const inFlight = new Map<string, Promise<PdfNativePrewarmResult>>();

function remember(sha256: string, pages: readonly number[]): void {
  const entry = warmed.get(sha256);
  const now = Date.now();
  const fresh = entry && now - entry.at < WARM_TTL_MS ? entry : { at: now, pages: new Set<number>() };
  for (const page of pages) fresh.pages.add(page);
  fresh.at = now;
  warmed.delete(sha256);
  warmed.set(sha256, fresh);
  while (warmed.size > MAX_HEADS) warmed.delete(warmed.keys().next().value as string);
}

export function prewarmPdfNativeSource(input: Readonly<{
  engineSource: PdfNativeEngineSourceRef;
  sourceSha256: string;
  pages?: readonly number[];
  signal?: AbortSignal;
  workingBytes?: () => Uint8Array | null;
}>): Promise<PdfNativePrewarmResult> {
  const started = typeof performance !== "undefined" ? performance.now() : Date.now();
  const elapsed = () => Math.round((typeof performance !== "undefined" ? performance.now() : Date.now()) - started);
  const sha256 = input.sourceSha256.toLowerCase();
  const requested = [...new Set((input.pages ?? []).filter((page) => Number.isSafeInteger(page) && page >= 1))]
    .sort((a, b) => a - b).slice(0, MAX_PAGES);
  const entry = warmed.get(sha256);
  const current = entry && Date.now() - entry.at < WARM_TTL_MS ? entry : null;
  const missing = requested.filter((page) => !current?.pages.has(page));
  if (current && missing.length === 0) {
    return Promise.resolve({ status: "already-warm", pages: requested, ms: 0 });
  }
  if (input.signal?.aborted) return Promise.resolve({ status: "aborted", pages: [], ms: 0 });
  const key = `${sha256}|${missing.join(",")}`;
  let shared = inFlight.get(key);
  if (!shared) {
    shared = callPdfNativeEngine("prewarm", {
      source: { ...input.engineSource, sourceSha256: sha256 },
      ...(missing.length ? { pages: missing } : {}),
    }, { workingBytes: input.workingBytes ?? (() => null) }).then(
      (answer) => {
        remember(sha256, answer.pages);
        return { status: "warm" as const, pages: answer.pages, ms: answer.ms };
      },
      () => ({ status: "failed" as const, pages: [], ms: 0 }),
    ).finally(() => inFlight.delete(key));
    inFlight.set(key, shared);
  }
  if (!input.signal) return shared.then((result) => ({ ...result, ms: elapsed() }));
  const signal = input.signal;
  return new Promise<PdfNativePrewarmResult>((resolve) => {
    const onAbort = () => resolve({ status: "aborted", pages: [], ms: elapsed() });
    signal.addEventListener("abort", onAbort, { once: true });
    void shared!.then((result) => {
      signal.removeEventListener("abort", onAbort);
      resolve({ ...result, ms: elapsed() });
    });
  });
}

/** Tests. */
export function clearPdfNativePrewarmState(): void {
  warmed.clear();
  inFlight.clear();
}
