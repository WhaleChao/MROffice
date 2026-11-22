/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  MAX_MRAI_NATIVE_SOURCE_BYTES,
  PdfNativeTransactionError,
  prepareMraiNativeContentStreamTransaction,
  sha256PdfBytes,
  verifyMraiNativeContentStreamCandidate,
} from "./engine/native/content-stream-transaction";
import {
  prepareMraiNativeTextEditingSelection,
  createMraiNativeTextReadSession,
  readMraiNativeTextStandingLook,
  mraiNativeTextTakesFormatsWhereItStands,
  planMraiNativeListTextSelectionReplacement,
  proofFromNativeListTextPlan,
  readMraiNativeListTextStandingLook,
  isMraiNativeListTextMissingGlyphError,
  MRAI_NATIVE_BLANK_TYPE3_GLYPHS_MESSAGE,
  type PdfNativeListTextSelectionProof,
  type PdfNativeSingleLineFitHint,
} from "./engine/native/list-text-fidelity";
import { pdfTextHasRightToLeftRun } from "./engine/native/reader-order";
import { PDF_NATIVE_MAX_FIXED_FRAME_ITEMS } from "./engine/native/vocabulary";
import { PDF_NATIVE_TEXT_NOT_A_FORMATS_BLOCK_MESSAGE } from "./text-edit/block-only-opening";
import { inheritPdfSourceAppearance } from "./native-source-appearance";
import { matchPdfTextSourceAppearance, pdfNativeCandidateAppearanceSelection } from "./source-text-appearance";
import { resolveManagedTextBoxPreviewAnnotationId } from "./native-text-box-roundtrip";
import { readPdfNativeResourceIssue } from "./engine/native/resource-issue";
import { pdfSourceTextBindingNamesItem, readPdfSourceTextBoxBinding } from "./source-text-box";
import { verifyPdfSourceTextBox } from "./source-text-box-proof";
import type {
  PdfNativeTextPreparationRequest,
  PdfNativeTextPreparationResponse,
} from "./native-text-edit-preparation";







export async function openPdfNativeTextInFallbackLane(
  bytes: Uint8Array,
  selection: PdfNativeTextPreparationRequest["selection"],
  refusal: unknown,
  blank = true,
): Promise<{ proof: PdfNativeListTextSelectionProof; blankedSource: Uint8Array | null }> {
  if (!(refusal instanceof PdfNativeTransactionError) || refusal.code !== "unsupported-document" ||
    /encrypt|certified|protected PDF state/iu.test(refusal.message)) throw refusal;
  const { PDF_FALLBACK_TYPE3_MESSAGE, PdfFallbackTextRefusal, preparePdfFallbackTextOpen } =
    await import("./source-text-fallback");
  if (refusal.message.includes(MRAI_NATIVE_BLANK_TYPE3_GLYPHS_MESSAGE) ||
    /drawn in a Type ?3 font|paragraph drawn in a Type ?3/iu.test(refusal.message)) {
    throw new PdfNativeTransactionError("unsupported-document", PDF_FALLBACK_TYPE3_MESSAGE);
  }
  try {
    return await preparePdfFallbackTextOpen({ sourceBytes: bytes, selection, blank });
  } catch (error) {
    if (error instanceof PdfFallbackTextRefusal) {
      throw new PdfNativeTransactionError("unsupported-document", error.message);
    }
    throw refusal;
  }
}

















export async function proofOfParagraphAsFormatsBlock(
  bytes: Uint8Array,
  selection: PdfNativeTextPreparationRequest["selection"],
): Promise<PdfNativeListTextSelectionProof | null> {
  if ((selection.items?.length ?? 0) < 2) return null;
  try {
    const read = await mraiNativeTextTakesFormatsWhereItStands({ sourceBytes: bytes, selection });
    if (!read || read.shows <= PDF_NATIVE_MAX_FIXED_FRAME_ITEMS) return null;
    const { pdfWordsFixedTextProof } = await import("./source-text-fallback");
    
    // them: the window then asks the weight of it as of any box -- a line of
    // it with an emoji on it takes Bold the same way, `type3Pictures`.)
    return Object.freeze({ ...(await pdfWordsFixedTextProof({ sourceBytes: bytes, selection }, "too-large-to-set-again")),
      ...read.look, ...(read.type3Pictures ? { type3Pictures: true as const } : {}),
      ...(read.type3Weight ? { type3Weight: true as const } : {}) });
  } catch {
    return null;
  }
}











export async function proofOfBoxBoldByType3Passes(
  bytes: Uint8Array,
  selection: PdfNativeTextPreparationRequest["selection"],
): Promise<PdfNativeListTextSelectionProof | null> {
  try {
    const read = await mraiNativeTextTakesFormatsWhereItStands({ sourceBytes: bytes, selection });
    if (!read?.type3Passes) return null;
    const { pdfWordsFixedTextProof } = await import("./source-text-fallback");
    return Object.freeze({ ...(await pdfWordsFixedTextProof({ sourceBytes: bytes, selection }, "weight-drawn-again")),
      ...read.look, type3Text: true as const, ...(read.type3Weight ? { type3Weight: true as const } : {}) });
  } catch {
    return null;
  }
}

/** The proof of a line of right-to-left text whose words are fixed
 * (pdfRightToLeftLineWordsFixedReason, pdfWordsFixedTextProof); null for any
 * other line, and where the page cannot be read that way. */
async function wordsFixedProofOfRightToLeftLine(
  bytes: Uint8Array,
  selection: PdfNativeTextPreparationRequest["selection"],
): Promise<PdfNativeListTextSelectionProof | null> {
  if (!pdfTextHasRightToLeftRun(selection.item.str)) return null;
  try {
    const { pdfRightToLeftLineWordsFixedReason, pdfWordsFixedTextProof } = await import("./source-text-fallback");
    const reason = await pdfRightToLeftLineWordsFixedReason({ sourceBytes: bytes, selection });
    return reason ? await pdfWordsFixedTextProof({ sourceBytes: bytes, selection }, reason) : null;
  } catch {
    return null;
  }
}

/** Testable source-bound read preparation; no API, persistence or writer dispatch. */
export async function prepareAndZeroizePdfNativeText(
  input: PdfNativeTextPreparationRequest,
  readSession?: ReturnType<typeof createMraiNativeTextReadSession>,
  appearanceReader?: (source:Uint8Array,sha:string,page:number)=>Promise<import("./source-page-appearance").PdfPageSourceAppearance>,
): Promise<PdfNativeTextPreparationResponse> {
  const bytes = new Uint8Array(input.sourceBytes);
  let blankedSource: Uint8Array | null = null;
  
  // the rules, the stroked weight and the lean a format set on the box where
  // it stands left on all of its text are read as that write reads them, so
  // its buttons read pressed and the next press takes them off.
  const withStandingLook = async (opened: PdfNativeListTextSelectionProof): Promise<PdfNativeListTextSelectionProof> => {
    const standing = await (readSession?.standingLook ?? readMraiNativeTextStandingLook)({
      sourceBytes: bytes, selection: input.selection }).catch(() => null);
    let proof = standing ? Object.freeze({ ...opened, ...standing }) : opened;
    // A box of right-to-left text: its formats are written where the
    // text stands by its shows alone, and how it is drawn is read the
    // same way (the reading above does not read text the reader
    // reorders): its rules, stroked weight and lean, as the proof's own
    // textRules and drawnStyle.
    const look = await readMraiNativeListTextStandingLook({ sourceBytes: bytes, selection: input.selection })
      .catch(() => null);
    if (look) {
      // (A box drawn in a bold face reads bold by that face, and says the
      // weight is the face's own: `boldFaces`, which no press takes off.)
      const { underline, strikethrough, syntheticBold, slanted, drawnBold, boldFaces, regularFace } = look;
      proof = Object.freeze({
        ...proof,
        ...(underline || strikethrough
          ? { textRules: Object.freeze({ ...proof.textRules, ...(underline ? { underline } : {}), ...(strikethrough ? { strikethrough } : {}) }) } : {}),
        ...(syntheticBold || slanted || drawnBold || boldFaces
          ? { drawnStyle: Object.freeze({ ...proof.drawnStyle, ...(syntheticBold || drawnBold ? { bold: true as const } : {}),
            ...(slanted ? { italic: true as const } : {}), ...(boldFaces ? { boldFaces } : {}),
            ...(regularFace ? { regularFace: true as const } : {}) }) } : {}),
      });
    }
    return proof;
  };
  // A box of right-to-left text that opens on a run's own proof (one show:
  // a single letter between two full stops, a word a producer drew whole) and
  // is drawn in a bold face: the proof says so the way a box of several shows
  // does -- `boldFaces`, and `regularFace` where the engine sets the box's
  // words again in the family's regular face -- so that its editor asks the
  
  // was offered there, its press set the run in a face of the proof by the
  // run's own writer -- which places no glyph of text the reader reorders --
  // and was refused when the box was left, while the engine sets that box
  // regular where it stands (pdf-walk-1000 0979 pp.1-6: 15 single-letter
  
  const withBoldFaceLook = async (opened: PdfNativeListTextSelectionProof): Promise<PdfNativeListTextSelectionProof> => {
    if (!opened.properties || opened.drawnStyle) return opened;
    const look = await readMraiNativeListTextStandingLook({ sourceBytes: bytes, selection: input.selection }).catch(() => null);
    if (!look?.boldFaces) return opened;
    return Object.freeze({ ...opened, drawnStyle: Object.freeze({
      ...(look.syntheticBold || look.drawnBold ? { bold: true as const } : {}), boldFaces: look.boldFaces,
      ...(look.regularFace ? { regularFace: true as const } : {}) }) });
  };
  try {
    if (
      !bytes.length ||
      bytes.length > MAX_MRAI_NATIVE_SOURCE_BYTES ||
      (await sha256PdfBytes(bytes)) !== input.selection.sourceSha256
    ) {
      throw new PdfNativeTransactionError("source-mismatch", "Source mismatch");
    }
    if (input.compile) {
      if (input.sourceText || input.managedBoxId) throw new PdfNativeTransactionError("invalid-request", "Managed text uses its native owner compiler.");
      const plan = await (readSession?.planReplacement ?? planMraiNativeListTextSelectionReplacement)({
        sourceBytes: bytes, selection: input.selection, replacementText: input.compile.replacementText,
        ...(input.compile.paragraphFrameRight !== undefined
          ? { paragraphFrameRight: input.compile.paragraphFrameRight } : {}),
        ...(input.compile.paragraphLineStep !== undefined
          ? { paragraphLineStep: input.compile.paragraphLineStep } : {}),
        ...(input.compile.paragraphSizeScale !== undefined
          ? { paragraphSizeScale: input.compile.paragraphSizeScale } : {}),
        ...(input.compile.rowStarts?.length ? { rowStarts: input.compile.rowStarts } : {}),
      });
      const prepared = await prepareMraiNativeContentStreamTransaction(bytes, plan.request);
      try { await verifyMraiNativeContentStreamCandidate(bytes, prepared); }
      catch (error) { prepared.candidateBytes.fill(0); throw error; }
      let sourceGeometry;
      if (appearanceReader && plan.evidence.selection) try {
        const scene=await appearanceReader(prepared.candidateBytes,prepared.candidateSha256,input.selection.pageNumber);
        const selection=pdfNativeCandidateAppearanceSelection(plan.evidence,prepared.candidateSha256);
        sourceGeometry=selection?matchPdfTextSourceAppearance(scene,selection)??undefined:undefined;
      } catch { /* A caret-geometry miss cannot invalidate an independently verified text candidate. */ }
      return { requestId: input.requestId, sourceSha256: input.selection.sourceSha256,
        result: { proof: proofFromNativeListTextPlan(plan), blankedSource: null, hiddenAnnotationId: null, compiled: { plan, prepared, ...(sourceGeometry?{sourceGeometry}:{}) } } };
    }
    if (input.formatsWhole) {
      // The box as a block that takes formats where its text stands and no
      
      // of a paragraph they have refused).
      if (input.sourceText || input.managedBoxId) throw new PdfNativeTransactionError("invalid-request", "A text box is not opened as a block of the page's own text.");
      const block = await proofOfParagraphAsFormatsBlock(bytes, input.selection);
      if (!block) throw new PdfNativeTransactionError("unsupported-document", PDF_NATIVE_TEXT_NOT_A_FORMATS_BLOCK_MESSAGE);
      return { requestId: input.requestId, sourceSha256: input.selection.sourceSha256,
        result: { proof: block, blankedSource: null, hiddenAnnotationId: null } };
    }
    const binding = input.sourceText
      ? readPdfSourceTextBoxBinding(input.sourceText)
      : null;
    const hiddenAnnotationId = input.managedBoxId
      ? await resolveManagedTextBoxPreviewAnnotationId(
          bytes,
          input.selection.pageNumber,
          input.managedBoxId,
        )
      : null;
    if (input.managedBoxId && !hiddenAnnotationId)
      throw new PdfNativeTransactionError(
        "precondition-mismatch",
        "Appearance mismatch",
      );
    let proof!: PdfNativeListTextSelectionProof;
    let singleLineFit: PdfNativeSingleLineFitHint | undefined;
    if (input.sourceText) {
      // The line the box was bound to: by its reader item, or by its text
      // where it stands once an edit above moved that item's index.
      if (
        !binding ||
        binding.boxId !== input.managedBoxId ||
        binding.page !== input.selection.pageNumber ||
        !pdfSourceTextBindingNamesItem(binding, input.selection.pageNumber, input.selection.item) ||
        !(await verifyPdfSourceTextBox({ sourceBytes: bytes, binding }))
      ) {
        throw new PdfNativeTransactionError(
          "precondition-mismatch",
          "Source text owner mismatch",
        );
      }
      proof = binding.originalProof;
      blankedSource = Uint8Array.from(bytes);
    } else {
      let result: Awaited<ReturnType<typeof prepareMraiNativeTextEditingSelection>> | null = null;
      try {
        result = await (readSession?.prepare ?? prepareMraiNativeTextEditingSelection)({
          sourceBytes: bytes,
          selection: input.selection,
        });
      } catch (error) {
        
        // is edited as a text box over its hidden source run. Only a Type3
        // face (WPS refuses these too) and a document that may not be
        // written stay refused (docs/pdf/edit-gates.md).
        
        
        if (process.env.NODE_ENV !== "production") {
          try {
            console.debug("[pdf-text-prepare] fallback lane " + JSON.stringify({
              page: input.selection.pageNumber, text: input.selection.item.str.slice(0, 24),
              reason: error instanceof Error ? error.message.slice(0, 600) : String(error).slice(0, 600) }));
          } catch { /* diagnostics never throw */ }
        }
        // (A box bold by a Type3 show drawn again opens for its weight: the
        
        const bold = await proofOfBoxBoldByType3Passes(bytes, input.selection);
        if (bold) proof = bold;
        else {
          const opened = await openPdfNativeTextInFallbackLane(bytes, input.selection, error);
          proof = await withStandingLook(opened.proof);
          blankedSource = opened.blankedSource;
        }
      }
      
      // lines the page draws glyphs their font only calls white space (the
      // dots an Urdu face stands over its letters, each a show of its own):
      
      
      // format set on it left the dots beside it as they were -- and no
      // writer takes what is typed into it: the native ones place no word
      
      // not take such a line. Its formats are written where the text stands,
      
      // proves of it: there the writers keep the glyphs they do not retype.)
      // So are the words of a line the reader reads as other words than the
      // page shows (a Nastaleeq face's ligatures, read turned round).
      const fixed = result ? await wordsFixedProofOfRightToLeftLine(bytes, input.selection) : null;
      if (fixed) {
        proof = await withStandingLook(fixed);
        result = null;
      }
      if (result) {
        proof = await withBoldFaceLook(result.proof);
        singleLineFit = result.singleLineFit;
        if (result.blankingRequest) {
          const prepared = await prepareMraiNativeContentStreamTransaction(
            bytes,
            result.blankingRequest,
          );
          blankedSource = prepared.candidateBytes;
          await verifyMraiNativeContentStreamCandidate(bytes, prepared);
        }
      }
    }
    if (!proof.fallbackText && proof.evidence.selection?.sourceSha256===input.selection.sourceSha256 &&
      (!proof.sourceGeometry || proof.sourceAppearance?.some(run=>run.text.trim()&&!run.displayColor))) {
      const owned: { current: Awaited<ReturnType<typeof import("./source-appearance-session")["createPdfSourceAppearanceSession"]>> | null } = { current: null };
      try {
        proof=await inheritPdfSourceAppearance(proof,async page=>{
          if(appearanceReader)return appearanceReader(bytes,input.selection.sourceSha256,page);
          const {createPdfSourceAppearanceSession}=await import("./source-appearance-session");
          owned.current=await createPdfSourceAppearanceSession(bytes,input.selection.sourceSha256);
          return owned.current.appearance(page);
        });
      } finally { owned.current?.dispose(); }
    }
    return {
      requestId: input.requestId,
      sourceSha256: input.selection.sourceSha256,
      result: { proof, blankedSource, hiddenAnnotationId, ...(singleLineFit ? { singleLineFit } : {}) },
    };
  } catch (error) {
    blankedSource?.fill(0);
    return {
      requestId: input.requestId,
      failure:
        error instanceof PdfNativeTransactionError
          ? error.code
          : "worker-failed",
      ...(isMraiNativeListTextMissingGlyphError(error) ? { missingGlyph: true as const } : {}),
      ...(error instanceof PdfNativeTransactionError && readPdfNativeResourceIssue(error.resourceIssue)
        ? { resourceIssue: readPdfNativeResourceIssue(error.resourceIssue)! } : {}),
      // The engine's own boundary message: it names the unsupported structure
      // (font, operator, frame) instead of a generic verification failure.
      // Source failures stay redacted: their message describes the caller's
      // bytes, not a structure of the document.
      ...(error instanceof PdfNativeTransactionError && error.message && error.code !== "source-mismatch"
        ? { detail: error.message.slice(0, 720) } : {}),
    };
  } finally {
    bytes.fill(0);
  }
}







export function createPdfNativeTextPreparationLane() {
  const readSession = createMraiNativeTextReadSession();
  let appearanceSource: {sha:string;session:Promise<Awaited<ReturnType<typeof import("./source-appearance-session")["createPdfSourceAppearanceSession"]>>>} | null = null;
  const readAppearance = async(bytes:Uint8Array,sha:string,page:number)=>{
    if(appearanceSource?.sha!==sha){
      void appearanceSource?.session.then(session=>session.dispose(),()=>undefined);
      const source=Uint8Array.from(bytes);
      const session=import("./source-appearance-session").then(async({createPdfSourceAppearanceSession})=>{
        try{return await createPdfSourceAppearanceSession(source,sha);}finally{source.fill(0);}
      });
      appearanceSource={sha,session};
    }
    return (await appearanceSource.session).appearance(page);
  };
  let queue: Promise<unknown> = Promise.resolve();
  return Object.freeze({
    /** The document read ahead of the first line opened on it (Edit mode
     * shown): the read session keeps it for that line's proof. */
    retain: (sourceBytes: ArrayBuffer): Promise<boolean> => {
      const bytes = new Uint8Array(sourceBytes);
      return readSession.retain(bytes).catch(() => false).finally(() => bytes.fill(0));
    },
    prepare: (input: PdfNativeTextPreparationRequest): Promise<Readonly<{
      response: PdfNativeTextPreparationResponse; transfer: Transferable[];
    }>> => {
      const answered = queue.then(async () => {
        const response = await prepareAndZeroizePdfNativeText(input, readSession, readAppearance);
        const result = "result" in response ? response.result : null;
        const buffer = result?.compiled?.prepared.candidateBytes.buffer ?? result?.blankedSource?.buffer;
        return Object.freeze({ response, transfer: buffer ? [buffer as ArrayBuffer] : [] });
      });
      queue = answered.catch(() => undefined);
      return answered;
    },
  });
}
