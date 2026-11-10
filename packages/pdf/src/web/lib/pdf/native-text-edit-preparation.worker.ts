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
  planMraiNativeListTextSelectionReplacement,
  proofFromNativeListTextPlan,
  isMraiNativeListTextMissingGlyphError,
  MRAI_NATIVE_BLANK_TYPE3_GLYPHS_MESSAGE,
  type PdfNativeListTextSelectionProof,
  type PdfNativeSingleLineFitHint,
} from "./engine/native/list-text-fidelity";
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

/** Testable source-bound read preparation; no API, persistence or writer dispatch. */
export async function prepareAndZeroizePdfNativeText(
  input: PdfNativeTextPreparationRequest,
  readSession?: ReturnType<typeof createMraiNativeTextReadSession>,
  appearanceReader?: (source:Uint8Array,sha:string,page:number)=>Promise<import("./source-page-appearance").PdfPageSourceAppearance>,
): Promise<PdfNativeTextPreparationResponse> {
  const bytes = new Uint8Array(input.sourceBytes);
  let blankedSource: Uint8Array | null = null;
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
        const opened = await openPdfNativeTextInFallbackLane(bytes, input.selection, error);
        proof = opened.proof;
        blankedSource = opened.blankedSource;
      }
      if (result) {
        proof = result.proof;
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
