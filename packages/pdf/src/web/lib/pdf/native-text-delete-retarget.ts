/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A queued delete of a run of the page's own text (Delete, the eraser),
 * stood on other bytes than it was made on: the version another window
 * saved first, or the same content the server rewrote as another file.
 *
 * The delete's command names its run by a span of the content stream --
 * two offsets, held to no digest -- beside the reader's selection the run
 * was proven from. On other bytes those offsets name whatever lies there
 * now, so the command is never planned by them. The run is found again by
 * its words and place (`find`, as the reader reads those bytes), proven
 * there as it was proven when the edit was made (`prove`), and the span that
 * proof names is the one deleted -- only when it is as long as the span the
 * edit took, in the stream that drew it: the same run, not a longer or a
 * shorter cut of it. Anything less is not aimed at.
 */
import { pdfNativeListTextSourceRunSpan } from "./engine/native/list-text-model";
import type { PdfNativeListTextSelectionEvidence, PdfNativeListTextSelectionProof } from "./engine/native/list-text-fidelity";

type ReaderItem = Readonly<{ str: string; fontName?: unknown; transform: readonly unknown[]; width?: unknown;
  height?: unknown; hasEOL?: unknown }>;

/**
 * One show of the page as the reader's own item on other bytes: the one item
 * with the same words whose pen starts at the same place, to a hundredth of
 * a point (`items`: the page's text items in the reader's order, the ones
 * that hold a string). A line inside a paragraph is found so; no box of the
 * page stands for it on its own. Null when no item, or more than one, reads
 * so, and for a selection of several items or of part of one (the item
 * found is selected whole, which a part of it is not).
 */
export function pdfNativeReaderItemSelectionOn(items: readonly ReaderItem[], sourceSha256: string,
  selection: PdfNativeListTextSelectionEvidence): PdfNativeListTextSelectionEvidence | null {
  if (selection.items?.length || selection.startOffset !== 0 || selection.endOffset !== selection.item.str.length) return null;
  const found = items.flatMap((item, itemIndex) => item.str === selection.item.str &&
    Math.abs(Number(item.transform[4]) - Number(selection.item.transform[4])) < 0.01 &&
    Math.abs(Number(item.transform[5]) - Number(selection.item.transform[5])) < 0.01 ? [{ item, itemIndex }] : []);
  if (found.length !== 1) return null;
  const { item, itemIndex } = found[0]!;
  return Object.freeze({ sourceSha256, selectionRevision: selection.selectionRevision, pageNumber: selection.pageNumber,
    startOffset: 0, endOffset: item.str.length,
    item: Object.freeze({ itemIndex, str: item.str, fontName: String(item.fontName),
      transform: Object.freeze(item.transform.map(Number)) as unknown as PdfNativeListTextSelectionEvidence["item"]["transform"],
      width: Number(item.width), height: Number(item.height), hasEOL: Boolean(item.hasEOL) }) });
}

export type PdfNativeTextDeleteStanding =
  /** Not a delete that names other bytes (or one this cannot stand: a
   * delete taken back, one with no selection): the command goes as it is. */
  | Readonly<{ status: "as-it-is" }>
  | Readonly<{ status: "stood"; payload: Record<string, unknown> }>
  | Readonly<{ status: "not-found" }>;

export async function standPdfNativeTextObjectDeleteOn(input: Readonly<{
  /** The delete command's payload. */
  payload: unknown;
  /** The bytes it is stood on. */
  sourceSha256: string;
  /** Those bytes are the content the edit was made on, as another file
   * (its streams numbered anew). */
  sameContent?: boolean;
  /** The run's selection on those bytes, by its words and place; null when
   * the page does not show it there once. */
  find: (selection: PdfNativeListTextSelectionEvidence) => Promise<PdfNativeListTextSelectionEvidence | null>;
  /** The planner's proof of a selection on those bytes. */
  prove: (selection: PdfNativeListTextSelectionEvidence) =>
    Promise<Pick<PdfNativeListTextSelectionProof, "evidence" | "sourceRun" | "target"> | null>;
}>): Promise<PdfNativeTextDeleteStanding> {
  const payload = input.payload && typeof input.payload === "object" && !Array.isArray(input.payload)
    ? input.payload as { direction?: unknown; selection?: { runSpan?: { start?: unknown; end?: unknown };
        contentStream?: { objectNumber?: unknown; generation?: unknown }; textSelection?: PdfNativeListTextSelectionEvidence } }
    : null;
  const made = payload?.direction === "delete" ? payload.selection?.textSelection : undefined;
  if (!payload?.selection || !made || typeof made.sourceSha256 !== "string" || !made.item ||
    made.sourceSha256.toLowerCase() === input.sourceSha256.toLowerCase()) return Object.freeze({ status: "as-it-is" as const });
  const notFound = Object.freeze({ status: "not-found" as const });
  const span = payload.selection.runSpan;
  if (!Number.isSafeInteger(span?.start) || !Number.isSafeInteger(span?.end)) return notFound;
  const found = await input.find(made);
  const proven = found ? await input.prove(found).catch(() => null) : null;
  const runSpan = pdfNativeListTextSourceRunSpan(proven);
  if (!proven || !runSpan || !proven.evidence.selection) return notFound;
  if (runSpan.end - runSpan.start !== (span!.end as number) - (span!.start as number)) return notFound;
  const stream = proven.sourceRun?.contentStream ?? proven.target ?? null;
  const held = payload.selection.contentStream;
  if (held && !stream) return notFound;
  // (On a version another window saved, the stream that draws the run is the
  // one that drew it.)
  if (held && stream && !input.sameContent &&
    (stream.objectNumber !== held.objectNumber || stream.generation !== held.generation)) return notFound;
  return Object.freeze({ status: "stood" as const, payload: { ...(payload as Record<string, unknown>),
    selection: { ...(payload.selection as Record<string, unknown>),
      runSpan: { start: runSpan.start, end: runSpan.end },
      ...(held && stream ? { contentStream: { objectNumber: stream.objectNumber, generation: stream.generation } } : {}),
      textSelection: proven.evidence.selection } } });
}
