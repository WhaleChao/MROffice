/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** One wet stroke's save: its row's id, a refusal no retry changes, or a failure. */
export type PdfWetInkWrite =
  | { readonly id: string }
  /** Too short to be a stroke once its points are cleaned. */
  | "not-a-stroke"
  /** The page already holds the most strokes a page can. */
  | "page-full"
  /** The write failed (or could not start): kept wet, tried again. */
  | "failed";

export interface PdfWetInkSettled<S> {
  /** The wet strokes after this save: the failed ones stay, in their order. */
  readonly wet: S[];
  /** The batch's strokes that stay wet for another try. */
  readonly failed: S[];
  /** The batch's strokes the page had no room for (said to the reader). */
  readonly pageFull: S[];
}

/**
 * The wet strokes after one save of `batch`. `taken` are strokes the reader
 * took back (the Eraser) while they were being saved: one of those whose
 * write failed goes, as asked, instead of being kept for a retry.
 */
export function settlePdfWetInkBatch<S extends { readonly key: number }>(
  wet: readonly S[],
  batch: readonly S[],
  writes: readonly (PdfWetInkWrite | undefined)[],
  taken: ReadonlySet<number> = new Set(),
): PdfWetInkSettled<S> {
  const failed: S[] = [];
  const pageFull: S[] = [];
  const leaves = new Set<number>();
  batch.forEach((stroke, index) => {
    // A write with no answer is a failed one: nothing is dropped on a doubt.
    const write = writes[index] ?? "failed";
    if (write === "failed") {
      if (taken.has(stroke.key)) leaves.add(stroke.key);
      else failed.push(stroke);
      return;
    }
    leaves.add(stroke.key);
    if (write === "page-full") pageFull.push(stroke);
  });
  return { wet: wet.filter((stroke) => !leaves.has(stroke.key)), failed, pageFull };
}

/** The landed row's id, or null. */
export function pdfWetInkLandedId(write: PdfWetInkWrite | undefined): string | null {
  return typeof write === "object" ? write.id : null;
}

/**
 * Wet strokes a reader was closed over while their save kept failing, by
 * file: the next reader of that file in this tab takes them back, draws
 * them and saves them, so closing a page (or the reader being rebuilt) is
 * not the end of a stroke the file never got.
 */
const kept = new Map<string, unknown[]>();

export function keepPdfWetInk<S>(file: string, strokes: readonly S[]): void {
  if (strokes.length === 0) return;
  kept.set(file, [...(kept.get(file) ?? []), ...strokes]);
}

export function takePdfKeptWetInk<S>(file: string): S[] {
  const strokes = kept.get(file) ?? [];
  kept.delete(file);
  return strokes as S[];
}

/** Tests only. */
export function resetPdfKeptWetInk(): void {
  kept.clear();
}
