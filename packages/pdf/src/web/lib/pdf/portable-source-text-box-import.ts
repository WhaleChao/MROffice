/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  PDFArray, PDFBool, PDFDict, PDFDocument, PDFHexString, PDFName, PDFRef, PDFString,
} from "pdf-lib";
import { PDF_JS_DOCUMENT_ASSETS, loadCompatiblePdfJs } from "@mrnote/core/web/lib/pdf/pdfjs-runtime.ts";
import { createPdfSessionDestroyer, releasePdfCanvasBackingStore } from "./render-lifecycle";
import { rotatedBoundingBox } from "@mrnote/core/web/lib/pdf/rotated-box.ts";
import { readPdfPortableTextLayout, type PdfPortableTextLayout } from "./portable-text-layout";
import { readPortablePdfSourceTextBoxBinding, type PdfSourceTextBoxBinding } from "./source-text-box";
import type { PdfNativePlanners } from "./remote/pdf-native-planners";
import { managedNativeTextBoxAnnotationName, managedNativeTextBoxStyleKey, sanitizeManagedNativeTextBoxImage,
  MAX_MANAGED_NATIVE_TEXT_BOX_IMAGE_TOTAL_CHARS, type ManagedNativeTextBoxAnnotation } from "./native-text-box-roundtrip";
import type { TextBoxEditLike } from "./text-box-flatten";
import { decodePdfTextBoxRunsField, pdfTextStyleRunsFromDrawnRuns } from "./text-box-runs";
import { sha256PdfBytes } from "./engine/native/transaction-error";
import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";
import { pdfLibObjectsPerTick, pdfLibParseSpeed } from "@mrnote/core/web/lib/pdf/pdf-lib-pacing.ts";

const MAX_BYTES = PDF_EDITABLE_MAX_BYTES;
const MAX_BOXES = 400;
const stringValue = (value: unknown): string | null =>
  value instanceof PDFString || value instanceof PDFHexString ? value.decodeText() : null;
const abort = () => new DOMException("PDF source import cancelled", "AbortError");

export interface PortableSourceTextBoxImport {
  readonly edit: TextBoxEditLike & { id: string; page: number; text: string };
  readonly record: ManagedNativeTextBoxAnnotation;
  readonly annotationRef: { objectNumber: number; generation: number };
}

/** Reconstruct source-owned text from the current PDF, without old notebook
 * JSON or fetching the resource named by a previous authoring session. The
 * returned records remain read projections until an existing native command
 * proves and commits an actual edit against the current resource receipt. */
export async function importPortableSourceTextBoxes(input: {
  /** The planners of `pageId`, the page the bytes are shown on. */
  planners: PdfNativePlanners;
  sourceBytes: Uint8Array;
  sourceSha256: string;
  resourceId: string;
  pageId: string;
  pageNumber: number;
  signal: AbortSignal;
  knownBoxIds?: readonly string[];
}): Promise<readonly PortableSourceTextBoxImport[]> {
  if (input.signal.aborted) throw abort();
  if (!input.sourceBytes.length || input.sourceBytes.length > MAX_BYTES ||
    await sha256PdfBytes(input.sourceBytes) !== input.sourceSha256) throw Error("PDF import source changed");
  const pdf = await PDFDocument.load(input.sourceBytes, { parseSpeed: pdfLibParseSpeed(), updateMetadata: false });
  if (pdf.isEncrypted || pdf.getPageCount() > 10_000 ||
    !Number.isSafeInteger(input.pageNumber) || input.pageNumber < 1 ||
    input.pageNumber > pdf.getPageCount()) throw Error("PDF import source is unavailable");
  const sourcePage = pdf.getPage(input.pageNumber - 1);
  const annots = sourcePage.node.lookupMaybe(PDFName.of("Annots"), PDFArray);
  if (!annots) return [];
  const candidates: { index: number; ref: PDFRef; dict: PDFDict }[] = [];
  const knownNames = new Set((input.knownBoxIds ?? []).map(managedNativeTextBoxAnnotationName));
  for (let index = 0; index < annots.size(); index++) {
    const ref = annots.get(index), dict = pdf.context.lookup(ref);
    if (!(dict instanceof PDFDict) || !dict.has(PDFName.of("MRNoteSourceTextBinding"))) continue;
    if (knownNames.has(stringValue(dict.get(PDFName.of("NM"))))) continue;
    if (!(ref instanceof PDFRef) || dict.get(PDFName.of("Subtype")) !== PDFName.of("FreeText") ||
      dict.get(PDFName.of("MRNoteTextBox")) !== PDFBool.True ||
      dict.get(PDFName.of("MRNoteManaged")) !== PDFBool.True ||
      ["A", "AA", "RC", "IRT", "Popup", "RD"].some(key => dict.has(PDFName.of(key))))
      throw Error("PDF portable text owner is unsupported");
    candidates.push({ index, ref, dict });
    if (candidates.length > MAX_BOXES) throw Error("PDF portable text import budget exceeded");
  }
  if (!candidates.length) return [];
  // Object identity is global, not a rectangle or a text match. A duplicate
  
  const ownerCounts = new Map<string, number>();
  const sourceOwnerCounts = new Map<string, number>();
  let scannedAnnotations = 0;
  for (const page of pdf.getPages()) {
    const rows = page.node.lookupMaybe(PDFName.of("Annots"), PDFArray);
    for (let i = 0; rows && i < rows.size(); i++) {
      if (++scannedAnnotations > MAX_BOXES * 10) throw Error("PDF annotation import budget exceeded");
      const ref = rows.get(i);
      if (ref instanceof PDFRef) ownerCounts.set(ref.toString(), (ownerCounts.get(ref.toString()) ?? 0) + 1);
      const dict = pdf.context.lookup(ref);
      const identity = dict instanceof PDFDict ? stringValue(dict.get(PDFName.of("MRNoteSourceTextIdentity"))) : null;
      if (identity) sourceOwnerCounts.set(identity, (sourceOwnerCounts.get(identity) ?? 0) + 1);
    }
  }
  const api = await loadCompatiblePdfJs();
  const sourceTask = api.getDocument({ data: Uint8Array.from(input.sourceBytes), ...PDF_JS_DOCUMENT_ASSETS });
  const destroySource = createPdfSessionDestroyer(() => ({ loadingTask: sourceTask }));
  const onAbort = () => { void destroySource().catch(() => undefined); };
  input.signal.addEventListener("abort", onAbort, { once: true });
  const imports: PortableSourceTextBoxImport[] = [];
  const ids = new Set<string>();
  let imageCharacters = 0;
  try {
    const sourceDocument = await sourceTask.promise;
    const pageProxy = await sourceDocument.getPage(input.pageNumber);
    const viewport = pageProxy.getViewport({ scale: 1 });
    for (const { index, ref, dict } of candidates) {
      if (input.signal.aborted) throw abort();
      if (ownerCounts.get(ref.toString()) !== 1) throw Error("PDF portable text owner is ambiguous");
      const bindingJson = stringValue(dict.get(PDFName.of("MRNoteSourceTextBinding")));
      const specJson = stringValue(dict.get(PDFName.of("MRNoteSourceTextSpec")));
      const seal = stringValue(dict.get(PDFName.of("MRNoteSourceTextSeal")));
      if (!bindingJson || bindingJson.length > 50_000 || !specJson || specJson.length > 100_000 || !seal)
        throw Error("PDF portable text metadata is unavailable");
      const binding = readPortablePdfSourceTextBoxBinding(JSON.parse(bindingJson), input, seal);
      const spec = JSON.parse(specJson) as Record<string, unknown>;
      const text = stringValue(dict.get(PDFName.of("Contents")));
      const styleKey = stringValue(dict.get(PDFName.of("MRNoteTextBoxStyleKey")));
      if (!binding || binding.page !== input.pageNumber || ids.has(binding.boxId) ||
        stringValue(dict.get(PDFName.of("NM"))) !== managedNativeTextBoxAnnotationName(binding.boxId) ||
        stringValue(dict.get(PDFName.of("MRNoteSourceTextIdentity"))) !== binding.sourceIdentity ||
        sourceOwnerCounts.get(binding.sourceIdentity!) !== 1 ||
        text === null || text !== spec.text || !styleKey || !/^[a-fA-F0-9]{16}$/.test(styleKey) ||
        !Array.isArray(spec.box) || spec.box.length !== 4 ||
        !spec.box.every(value => typeof value === "number" && Number.isFinite(value)) ||
        !["sans", "serif", "mono"].includes(String(spec.font)) ||
        !["left", "center", "right"].includes(String(spec.align)) ||
        typeof spec.fontSize !== "number" || !Number.isFinite(spec.fontSize) || spec.fontSize <= 0 ||
        typeof spec.angle !== "number" || !Number.isFinite(spec.angle) ||
        typeof spec.color !== "string" || !/^[a-fA-F0-9]{6}$/.test(spec.color) || spec.background !== null ||
        [spec.bold, spec.italic, spec.underline].some(value => typeof value !== "boolean") ||
        !(await input.planners.verifyPdfSourceTextBox({ sourceBytes: input.sourceBytes, binding })))
        throw Error("PDF portable text metadata does not match its source");
      // One show's binding names the page and its one content stream. A
      // span binding (version 3) names the spans it hides, which may stand
      // in any stream the page draws (one of several, a form of its own):
      // that they are hidden on this page is what was just verified.
      const pageIdentity = binding.pieces ? null : binding.originalProof.evidence.resources.page;
      const contents = sourcePage.node.get(PDFName.of("Contents"));
      if ((pageIdentity !== null && (pageIdentity.objectNumber !== sourcePage.ref.objectNumber ||
        pageIdentity.generation !== sourcePage.ref.generationNumber ||
        !(contents instanceof PDFRef) || contents.objectNumber !== binding.object ||
        contents.generationNumber !== binding.generation)) ||
        (dict.has(PDFName.of("P")) && dict.get(PDFName.of("P"))?.toString() !== sourcePage.ref.toString()))
        throw Error("PDF portable text page identity changed");
      ids.add(binding.boxId);
      const [x, y, width, height] = spec.box as number[];
      if (x < 0 || y < 0 || width <= 0 || height <= 0 || x + width > 1.0001 || y + height > 1.0001)
        throw Error("PDF portable text geometry is invalid");
      // A box set in runs comes back in its runs (the writer's `^` field).
      const runs = spec.runs === undefined ? null
        : typeof spec.runs === "string" ? decodePdfTextBoxRunsField(`^${spec.runs}`, text) : null;
      if (spec.runs !== undefined && !runs) throw Error("PDF portable text runs are invalid");
      const layout = readPdfPortableTextLayout(spec.portableLayout, { text,
        textSha256: await sha256PdfBytes(new TextEncoder().encode(text)),
        width: width * viewport.width, height: height * viewport.height,
        fontSize: spec.fontSize * viewport.height,
        ...(runs ? { minimumBaseline: Math.min(...runs.map((run) => run.fontSize)) * viewport.height } : {}) });
      if (!layout) throw Error("PDF portable text layout is invalid");
      const visibleBounds = rotatedBoundingBox({ x: x * viewport.width, y: y * viewport.height,
        width: width * viewport.width, height: height * viewport.height }, spec.angle);
      const corners = [viewport.convertToPdfPoint(visibleBounds.x, visibleBounds.y),
        viewport.convertToPdfPoint(visibleBounds.x + visibleBounds.width, visibleBounds.y + visibleBounds.height)];
      const expectedRect = [Math.min(corners[0][0], corners[1][0]), Math.min(corners[0][1], corners[1][1]),
        Math.max(corners[0][0], corners[1][0]), Math.max(corners[0][1], corners[1][1])];
      const rect = dict.lookup(PDFName.of("Rect"), PDFArray).asArray().map(value => Number(value.toString()));
      if (rect.length !== 4 || rect.some((value, i) => !Number.isFinite(value) || Math.abs(value - expectedRect[i]) > 0.01))
        throw Error("PDF portable text rectangle changed");
      const image = await renderOwnedAppearance(pdf, input.pageNumber, index, rect, input.signal);
      imageCharacters += image.length;
      if (!sanitizeManagedNativeTextBoxImage(image) || imageCharacters > MAX_MANAGED_NATIVE_TEXT_BOX_IMAGE_TOTAL_CHARS)
        throw Error("PDF portable appearance image budget exceeded");
      const { edit, record } = portableSourceTextBoxProjection({ binding, pageNumber: input.pageNumber,
        box: [x, y, width, height], text, image, runs, layout: layout as PdfPortableTextLayout,
        font: spec.font as "sans" | "serif" | "mono", fontSize: spec.fontSize, color: spec.color, angle: spec.angle,
        align: spec.align as "left" | "center" | "right", bold: spec.bold as boolean, italic: spec.italic as boolean,
        underline: spec.underline as boolean });
      imports.push({ edit, record, annotationRef: { objectNumber: ref.objectNumber, generation: ref.generationNumber } });
    }
    return imports;
  } finally {
    input.signal.removeEventListener("abort", onAbort);
    await destroySource();
  }
}














export function portableSourceTextBoxProjection(input: {
  binding: PdfSourceTextBoxBinding;
  pageNumber: number;
  box: readonly [number, number, number, number];
  text: string;
  image: string;
  runs: ReturnType<typeof decodePdfTextBoxRunsField>;
  layout: PdfPortableTextLayout;
  font: "sans" | "serif" | "mono";
  fontSize: number;
  color: string;
  angle: number;
  align: "left" | "center" | "right";
  bold: boolean;
  italic: boolean;
  underline: boolean;
}): Pick<PortableSourceTextBoxImport, "edit" | "record"> {
  const { binding, text, image, runs, font } = input;
  const [x, y, width, height] = input.box;
  const edit = { id: binding.boxId, page: input.pageNumber - 1, x, y, width, height, text,
    fontSize: input.fontSize, color: `#${input.color}`, angle: input.angle,
    align: input.align, bold: input.bold, italic: input.italic, underline: input.underline,
    sourceFontName: font === "mono" ? "Courier" : font === "serif" ? "Times-Roman" : "Helvetica",
    ...(runs ? { runs: pdfTextStyleRunsFromDrawnRuns(runs) } : {}),
    created: true, sourceText: binding, nativeAppearanceImage: image };
  const record: ManagedNativeTextBoxAnnotation = { id: binding.boxId, page: input.pageNumber,
    x, y, width, height, text, fontSize: input.fontSize, color: `#${input.color}`, angle: input.angle,
    align: edit.align, font, bold: edit.bold, italic: edit.italic, underline: edit.underline,
    styleKey: managedNativeTextBoxStyleKey(edit), nativeAppearanceImage: image, portableLayout: input.layout,
    sourceText: binding, ...(runs ? { runs } : {}) };
  return { edit, record };
}

/** Render the actual retained AP into a cache image. No fabricated preview or
 * text reconstruction is used as evidence for the imported object's paint. */
async function renderOwnedAppearance(pdf: PDFDocument, pageNumber: number, annotationIndex: number,
  rect: number[], signal: AbortSignal): Promise<string> {
  const scratch = await PDFDocument.create();
  const [page] = await scratch.copyPages(pdf, [pageNumber - 1]);
  scratch.addPage(page);
  const annotations = page.node.lookup(PDFName.of("Annots"), PDFArray);
  page.node.set(PDFName.of("Annots"), scratch.context.obj([annotations.get(annotationIndex)]));
  page.node.set(PDFName.of("Contents"), scratch.context.register(scratch.context.flateStream(new Uint8Array())));
  page.setCropBox(rect[0], rect[1], rect[2] - rect[0], rect[3] - rect[1]);
  const bytes = await scratch.save({ objectsPerTick: pdfLibObjectsPerTick(), useObjectStreams: false, addDefaultPage: false });
  if (signal.aborted) throw abort();
  const api = await loadCompatiblePdfJs();
  const task = api.getDocument({ data: bytes, ...PDF_JS_DOCUMENT_ASSETS });
  const destroy = createPdfSessionDestroyer(() => ({ loadingTask: task }));
  const canvas = document.createElement("canvas");
  const onAbort = () => { void destroy().catch(() => undefined); };
  signal.addEventListener("abort", onAbort, { once: true });
  try {
    const documentProxy = await task.promise;
    const pageProxy = await documentProxy.getPage(1);
    const viewport = pageProxy.getViewport({ scale: 2 });
    if (viewport.width <= 0 || viewport.height <= 0 || viewport.width * viewport.height > 16_000_000)
      throw Error("PDF portable appearance pixel budget exceeded");
    canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
    const context = canvas.getContext("2d");
    if (!context) throw Error("PDF portable appearance canvas is unavailable");
    await pageProxy.render({ canvas, canvasContext: context, viewport, background: "rgba(0,0,0,0)" }).promise;
    if (signal.aborted) throw abort();
    return canvas.toDataURL("image/png");
  } finally {
    signal.removeEventListener("abort", onAbort);
    await destroy();
    releasePdfCanvasBackingStore(canvas);
  }
}
