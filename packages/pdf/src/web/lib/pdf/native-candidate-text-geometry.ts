/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where a candidate preview draws the edited box's characters: read from the
 * candidate page itself, in the boxes the page shows once the edit is
 * written (nativeParagraphEditBoxesOfTextContent), so the caret, the
 * selection and the frame the editor draws over a preview are placed by the
 * pixels under them and by nothing else. A preview the editor cannot read
 * the box back from carries no geometry: the editor then draws no caret over
 * it rather than one placed by its own draft's layout (a title set larger
 * and wrapped onto two rows had its selection painted where the word stood
 * before, owner 10-27/10-28).
 */
import type { PDFPageProxy } from "pdfjs-dist";
import { leadingRunSelection, movedRunSelection } from "./moved-run-selection";
import { nativeParagraphEditBoxesOfTextContent } from "./native-paragraph-edit-boxes";
import { readPdfTextContentWithHardGroups } from "./native-text-hard-groups";
import { readPdfLoadedTextSourceGeometry } from "./native-text-loaded-surface";
import type { PdfTextSourceAppearance } from "./source-text-appearance";
import type { PdfNativeKeptTextBlock } from "./native-kept-text-blocks";

export interface PdfCandidateTextGeometryRequest {
  /** The editor's key for the box (its layout's sourceInputKey). */
  readonly key: string;
  /** The box's text as the editor holds it. */
  readonly text: string;
  /** The baseline (PDF user space) the box's first line stood on. */
  readonly baselineY: number;
  /** The blocks the editor kept (pdfNativeKeptTextBlocks), as the viewer reads the page with. */
  readonly keptBlocks?: readonly PdfNativeKeptTextBlock[];
}

/** The box `request` names on the candidate `page`, as character geometry
 * in the page's own coordinates; null when the page does not show one box
 * with that text there. */
export async function readPdfCandidateTextGeometry(input: Readonly<{
  page: PDFPageProxy;
  candidateSha256: string;
  request: PdfCandidateTextGeometryRequest;
  ops: Readonly<Record<string, number>>;
  signal?: AbortSignal;
}>): Promise<PdfTextSourceAppearance | null> {
  const { page, request } = input;
  let content: Awaited<ReturnType<typeof readPdfTextContentWithHardGroups>>;
  try {
    content = await readPdfTextContentWithHardGroups(page as unknown as Parameters<typeof readPdfTextContentWithHardGroups>[0]);
  } catch {
    return null;
  }
  if (input.signal?.aborted) return null;
  const { boxes } = nativeParagraphEditBoxesOfTextContent({
    pageNumber: page.pageNumber,
    content,
    viewport: page.getViewport({ scale: 1 }) as never,
    ...(request.keptBlocks ? { keptBlocks: request.keptBlocks } : {}),
  });
  // The box with the text, on the line it opened on (a box set again keeps
  // its first line where it stood); anywhere on the page when it moved; or
  // the leading items of the box on its line, where the reader joins it to
  // the text after it until the write's kept block parts them.
  const found = movedRunSelection(boxes, request.text, request.baselineY) ??
    movedRunSelection(boxes, request.text, request.baselineY, { anywhere: true }) ??
    leadingRunSelection(boxes, request.text, request.baselineY);
  if (!found) return null;
  const selection = Object.freeze({
    sourceSha256: input.candidateSha256,
    selectionRevision: 1,
    pageNumber: page.pageNumber,
    startOffset: 0,
    endOffset: found.endOffset,
    item: found.item,
    ...(found.items ? { items: found.items } : {}),
  });
  const geometry = await readPdfLoadedTextSourceGeometry({ page, sourceSha256: input.candidateSha256,
    selection, text: request.text, ops: input.ops, ...(input.signal ? { signal: input.signal } : {}) });
  return geometry && geometry.text === request.text ? geometry : null;
}

const candidateInk = new WeakMap<HTMLCanvasElement, PdfCandidateTextInk>();

/** Where a text stands on a page, in shares of the page. */
export type PdfCandidateTextInk = Readonly<{ left: number; top: number; right: number; bottom: number }>;

export function rememberPdfCandidateTextInk(canvas: HTMLCanvasElement, ink: PdfCandidateTextInk): void {
  candidateInk.set(canvas, ink);
}

export function readPdfCandidateTextInk(canvas: HTMLCanvasElement): PdfCandidateTextInk | null {
  return candidateInk.get(canvas) ?? null;
}

/**
 * Where the words of `request` stand on the candidate `page`, from the items
 * that spell them in reading order. A size set larger takes a word out of its
 * line and the reader out of one box, so no box reads as the text and there is
 * no character geometry; the words are still where the items say.
 */
export async function measurePdfCandidateTextInk(input: Readonly<{
  page: PDFPageProxy;
  request: PdfCandidateTextGeometryRequest;
}>): Promise<PdfCandidateTextInk | null> {
  const squeeze = (text: string) => text.replace(/\s+/gu, "");
  const wanted = squeeze(input.request.text);
  if (!wanted) return null;
  let content: Awaited<ReturnType<PDFPageProxy["getTextContent"]>>;
  try {
    content = await input.page.getTextContent();
  } catch {
    return null;
  }
  const items = content.items.flatMap((item) => {
    if (!("str" in item) || !item.transform || !squeeze(item.str)) return [];
    return [{ text: squeeze(item.str), x: item.transform[4] as number, y: item.transform[5] as number,
      width: item.width, height: item.height || Math.abs(item.transform[3] as number) }];
  });
  let best: (typeof items) | null = null;
  let bestDistance = Infinity;
  for (let start = 0; start < items.length; start += 1) {
    if (!wanted.startsWith(items[start].text)) continue;
    let at = 0, end = start;
    while (end < items.length && wanted.startsWith(items[end].text, at)) {
      at += items[end].text.length;
      end += 1;
      if (at === wanted.length) break;
    }
    if (at !== wanted.length) continue;
    const distance = Math.abs(items[start].y - input.request.baselineY);
    if (distance < bestDistance) {
      bestDistance = distance;
      best = items.slice(start, end);
    }
  }
  if (!best) return null;
  const viewport = input.page.getViewport({ scale: 1 });
  const points = best.flatMap((item) => [
    viewport.convertToViewportPoint(item.x, item.y - item.height * 0.2),
    viewport.convertToViewportPoint(item.x + item.width, item.y + item.height * 0.8),
  ]);
  const xs = points.map((point) => point[0] / viewport.width), ys = points.map((point) => point[1] / viewport.height);
  return Object.freeze({ left: Math.min(...xs), right: Math.max(...xs), top: Math.min(...ys), bottom: Math.max(...ys) });
}
