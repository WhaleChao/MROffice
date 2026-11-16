/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// (No import of pdf-lib here: some callers load it only when they need it.)
/** pdf-lib's own `ParseSpeeds.Fast` and `ParseSpeeds.Fastest`. */
const PARSE_FAST = 1_500;
const PARSE_IN_ONE_GO = Number.POSITIVE_INFINITY;

function onPageThread(): boolean {
  return typeof window !== "undefined" && typeof document !== "undefined";
}

/** pdf-lib's `parseSpeed` (the objects it parses between two waits). */
export function pdfLibParseSpeed(): number {
  return onPageThread() ? PARSE_FAST : PARSE_IN_ONE_GO;
}

/** pdf-lib's `objectsPerTick` (the objects it writes between two waits). */
export function pdfLibObjectsPerTick(): number {
  return onPageThread() ? 1_000 : Number.POSITIVE_INFINITY;
}
