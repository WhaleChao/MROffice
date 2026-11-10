/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Development-only latency marks for the PDF text edit path (click, proof,
 * keystroke, exact preview, finish, save). Each mark is a `performance.mark`
 * named `pdf-latency:<event>` plus one `console.debug("[pdf-latency] …")`
 * line, so a walk-through can read the timeline from the console or the
 * Performance panel. Production builds compile it to nothing; a diagnostic
 * never throws into the edit it measures.
 */
// A production build made for measuring sets NEXT_PUBLIC_MRAI_PDF_LATENCY=1
// at build time; any other production build compiles the marks away.
const enabled = process.env.NODE_ENV === "development" || process.env.NEXT_PUBLIC_MRAI_PDF_LATENCY === "1";

export function markPdfTextLatency(event: string, detail?: Readonly<Record<string, unknown>>): void {
  if (!enabled) return;
  try {
    performance.mark(`pdf-latency:${event}`);
    console.debug("[pdf-latency]", event, Math.round(performance.now()), detail ?? "");
  } catch {
    /* A diagnostic must not break the edit it measures. */
  }
}

/** Marks `<event>:start` now; the returned function marks `<event>:end` with
 * the elapsed milliseconds (and an outcome, when one is given). */
export function startPdfTextLatency(event: string, detail?: Readonly<Record<string, unknown>>): (outcome?: string) => void {
  if (!enabled) return () => undefined;
  let start = 0;
  try { start = performance.now(); } catch { /* see markPdfTextLatency */ }
  markPdfTextLatency(`${event}:start`, detail);
  let ended = false;
  return (outcome) => {
    if (ended) return;
    ended = true;
    let elapsed = 0;
    try { elapsed = Math.round(performance.now() - start); } catch { /* see markPdfTextLatency */ }
    markPdfTextLatency(`${event}:end`, outcome ? { ms: elapsed, outcome } : { ms: elapsed });
  };
}
