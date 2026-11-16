/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  inspectMraiNativeContentStream,
  inspectMraiNativeDocumentPageContentGraphs,
  sha256PdfBytes,
} from "./engine/native/content-stream-transaction";
import { isMraiNativeSourceTextShowWindow, type PdfNativeListTextSelectionProof } from "./engine/native/list-text-fidelity";
import {
  durableSourceTextProof,
  normalizedText,
  pdfSourceTextBoxOperator,
  readPdfSourceTextBoxBinding,
  spanBindingProof,
  type PdfSourceTextBoxBinding,
} from "./source-text-box";

export async function bindPdfSourceTextBox(input: {
  sourceBytes: Uint8Array;
  page: number;
  resourceId: string;
  pageId: string;
  boxId: string;
  proof: PdfNativeListTextSelectionProof;
}): Promise<PdfSourceTextBoxBinding> {
  
  // which bytes draw every glyph the reader shows of the box -- both passes
  // of a fake-bold title, the codes of a reader chunk, each line of a
  // paragraph set line by line -- and the server hides exactly those. A
  // clip, a paint state or a paragraph proof the server could not re-prove
  // for one show window (version 1) is no reason to keep the reader's text
  // out (docs/pdf/edit-gates.md). Version 1 remains only for a box the
  // span reading cannot find, when its native proof names one show window.
  const selection = input.proof.evidence.selection;
  const wholeBox = !selection || normalizedText(input.proof.evidence.originalDisplayText) ===
    normalizedText(selection.item.str);
  let data = input.proof.fallbackText ?? null;
  if (!data && selection) {
    const { locatePdfFallbackText } = await import("./source-text-fallback");
    try { data = await locatePdfFallbackText({ sourceBytes: input.sourceBytes, selection }); }
    catch (error) {
      // A refusal that is about what hiding the text would lose is not a
      // reading the one-show proof can stand in for: that proof hides the
      // same glyphs.
      const code = (error as { code?: string }).code;
      // (Nor is a line that is another box's hidden source a line to bind.)
      if (!wholeBox || code === "type3" || code === "glyphs-read-as-blanks" || code === "text-box-source") throw error;
      return bindProvedSourceTextBox(input);
    }
  }
  if (!data) return bindProvedSourceTextBox(input);
  const first = data.pieces[0];
  return Object.freeze({
    version: 3,
    pieces: data.pieces,
    resourceId: input.resourceId,
    pageId: input.pageId,
    boxId: input.boxId,
    expectedAppearance: null,
    page: input.page,
    object: first.object,
    generation: first.generation,
    operator: data.firstPieceHead,
    sha256: first.sha256,
    originalProof: spanBindingProof(input.proof, data),
  });
}

async function bindProvedSourceTextBox(input: {
  sourceBytes: Uint8Array;
  page: number;
  resourceId: string;
  pageId: string;
  boxId: string;
  proof: PdfNativeListTextSelectionProof;
}): Promise<PdfSourceTextBoxBinding> {
  const stream = await inspectMraiNativeContentStream(
    input.sourceBytes,
    input.proof.target,
  );
  // A box that is one reader item of a longer show (a chapter title and its
  // page number set as one TJ, cut by the reader at the wide gap) owns only
  // its own codes: it is hidden by span (bindPdfSourceTextBox).
  if (input.proof.evidence.readerChunk) {
    throw new Error("This text is part of a longer run drawn as one.");
  }
  const span = input.proof.evidence.textOperatorSpan;
  const raw = stream.streamBytes.slice(span.start, span.end);
  if (raw.length > 24000)
    throw new Error("Source text operator exceeds the bounded editor size.");
  const source = String.fromCharCode(...raw);
  // The same exact per-show evidence gates first promotion and reentry.
  // A broad string regex cannot distinguish one show from a font-switched paragraph.
  if (!(await isMraiNativeSourceTextShowWindow(raw, input.proof))) {
    throw new Error(
      "This text cannot yet be moved into a transparent text box safely.",
    );
  }
  const sha256 = await sha256PdfBytes(raw);
  if (sha256 !== span.sha256)
    throw new Error("The original text operator changed.");
  // The server proves a one-show binding only on a page drawn by one
  // content stream, the one bound (source_text_box.py: "source text stream
  // not isolated"). A show in a form the page draws, or in one stream of
  // several, bound here was stored as a box the server never writes, and it
  // refused every later write of the page with it (walk 0515: the PDF left
  // unwritten, the window unsaved). Such a line is bound by span or not at
  // all. The page is the one the proof read; one this reader cannot find
  // or walk is left to the server, as before.
  const proved = input.proof.evidence.resources?.page;
  const drawn = !proved ? null : await inspectMraiNativeDocumentPageContentGraphs(input.sourceBytes).then(
    (pages) => pages.find((page) => page.page.objectNumber === proved.objectNumber &&
      page.page.generation === proved.generation)?.contentStreams ?? null, () => null);
  if (drawn && (drawn.length !== 1 || drawn[0].objectNumber !== input.proof.target.objectNumber ||
    drawn[0].generation !== input.proof.target.generation)) {
    throw new Error(
      "This text cannot yet be moved into a transparent text box safely.",
    );
  }
  return Object.freeze({
    version: 1,
    resourceId: input.resourceId,
    pageId: input.pageId,
    boxId: input.boxId,
    expectedAppearance: null,
    page: input.page,
    object: input.proof.target.objectNumber,
    generation: input.proof.target.generation,
    operator: btoa(source),
    sha256,
    originalProof: durableSourceTextProof(input.proof),
  });
}

export async function verifyPdfSourceTextBox(input: {
  sourceBytes: Uint8Array;
  binding: PdfSourceTextBoxBinding;
}): Promise<boolean> {
  const b = readPdfSourceTextBoxBinding(input.binding);
  if (b?.version === 3 && b.pieces) {
    const { isPdfFallbackTextHidden } = await import("./source-text-fallback");
    return isPdfFallbackTextHidden(input.sourceBytes, { pageNumber: b.page, pieces: b.pieces });
  }
  if (!b || (await sha256PdfBytes(pdfSourceTextBoxOperator(b))) !== b.sha256)
    return false;
  const stream = await inspectMraiNativeContentStream(input.sourceBytes, {
    objectNumber: b.object,
    generation: b.generation,
  });
  const { isMraiNativeHiddenTextShow } =
    await import("./engine/native/list-text-fidelity");
  return isMraiNativeHiddenTextShow(
    stream.streamBytes,
    pdfSourceTextBoxOperator(b),
    b.originalProof,
  );
}
