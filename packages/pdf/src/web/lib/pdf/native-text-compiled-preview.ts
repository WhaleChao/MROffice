/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A compiled text preview the editor may hand its commit (the preview's
 * plan and candidate, privately cloned): only compilations the preparation
 * path issued can seed it. Kept apart from the preparation transport so the
 * planner (which reads a retained preview) does not import the browser's
 * engine client.
 */
import type { PdfNativeListTextSelectionEvidence } from "./engine/native/list-text-fidelity";
import type { PdfNativeTextPreparation } from "./native-text-edit-preparation";

type CompiledText = NonNullable<PdfNativeTextPreparation["compiled"]>;
type RetainedCompilation = {compiled:CompiledText; selectionKey:string};
const issuedCompilations = new WeakMap<object,RetainedCompilation>();
declare const compiledTextPreviewBrand: unique symbol;
export type PdfCompiledTextPreview = Readonly<{[compiledTextPreviewBrand]:true}>;
const retainedCompilations = new WeakMap<PdfCompiledTextPreview, RetainedCompilation>();
function textCommandSelectionKey(selection:PdfNativeListTextSelectionEvidence):string {
  // Only fields admitted by the native command schema, retaining every item
  // and geometric witness. PDF.js-only display metadata is not an identity.
  const item = (value:PdfNativeListTextSelectionEvidence["item"]) => value ? [value.itemIndex,value.str,value.fontName,
    value.transform,value.width,value.height,value.hasEOL] : null;
  return JSON.stringify([selection.sourceSha256,selection.selectionRevision,selection.pageNumber,
    selection.startOffset,selection.endOffset,item(selection.item),selection.items?.map(item)??null]);
}

/** Compare every admitted command witness, excluding renderer-only grouping
 * metadata that the native proof schema deliberately does not retain. */
export function samePdfNativeTextCommandSelection(left:PdfNativeListTextSelectionEvidence,
  right:PdfNativeListTextSelectionEvidence):boolean {
  return textCommandSelectionKey(left)===textCommandSelectionKey(right);
}

/** Only results returned by this source worker may seed the save cache. All
 * bytes/evidence are privately cloned; later UI/canvas mutations confer no
 * authority. The adapter still runs its normal fresh source verification. */
export function retainPdfCompiledTextPreview(compiled: CompiledText): PdfCompiledTextPreview | null {
  const issued = issuedCompilations.get(compiled);
  if (!issued) return null;
  const token = Object.freeze(Object.create(null)) as PdfCompiledTextPreview;
  issuedCompilations.delete(compiled);
  retainedCompilations.set(token,issued);
  return token;
}
export function releasePdfCompiledTextPreview(token: PdfCompiledTextPreview | null): void {
  if (!token) return;
  retainedCompilations.get(token)?.compiled.prepared.candidateBytes.fill(0);
  retainedCompilations.delete(token);
}
/** The compiled display text keeps the trailing space the source item ends
 * in ("131 "); the box holds only the typed text. */
function sameTypedDisplayText(compiled: string, typed: string, originalDisplay: string): boolean {
  if (compiled === typed) return true;
  const trailing = /\s+$/u.exec(originalDisplay)?.[0];
  return Boolean(trailing) && !/\s$/u.test(typed) && compiled === typed + trailing;
}
export function readPdfCompiledTextPreview(token: PdfCompiledTextPreview | null | undefined,
  selection: PdfNativeListTextSelectionEvidence, replacementText: string): CompiledText | null {
  const cached = token ? retainedCompilations.get(token) : null;
  if (!cached || cached.compiled.prepared.request.identity.sourceSha256 !== selection.sourceSha256 ||
    cached.selectionKey !== textCommandSelectionKey(selection) ||
    !sameTypedDisplayText(cached.compiled.plan.evidence.replacementDisplayText, replacementText,
      cached.compiled.plan.evidence.originalDisplayText) ||
    cached.compiled.plan.evidence.propertyMutation) return null;
  return structuredClone(cached.compiled);
}

/** Mark a compilation the preparation path returned as one a later commit
 * may reuse (`retainPdfCompiledTextPreview`). */
export function issuePdfCompiledTextPreview(compiled: CompiledText, selection: PdfNativeListTextSelectionEvidence): void {
  issuedCompilations.set(compiled, { compiled: structuredClone(compiled), selectionKey: textCommandSelectionKey(selection) });
}
