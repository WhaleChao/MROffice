/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The picture a Picture Format write drew from (planMraiNativePictureOrigin
 * names it in the page's /PieceInfo /MRAI data), drawn back in a later
 * session: Reset Picture returns to it and later formats draw from it again,
 * as Word's do.
 */

import { PDFDict, PDFDocument, PDFName, PDFObjectCopier, PDFRef, type PDFObject } from "pdf-lib";

import { loadCompatiblePdfJs, PDF_JS_DOCUMENT_ASSETS } from "@mrnote/core/web/lib/pdf/pdfjs-runtime.ts";
import { pdfLibObjectsPerTick, pdfLibParseSpeed } from "@mrnote/core/web/lib/pdf/pdf-lib-pacing.ts";

type ObjectRef = Readonly<{ objectNumber: number; generation: number }>;

/** The longest side an original is drawn back at. */
const ORIGINAL_MAX_SIDE = 6000;

function lookup(document: PDFDocument, value: PDFObject | undefined): PDFObject | undefined {
  return value instanceof PDFRef ? document.context.lookup(value) : value;
}

async function load(bytes: Uint8Array): Promise<PDFDocument> {
  return PDFDocument.load(bytes, { parseSpeed: pdfLibParseSpeed(), ignoreEncryption: true, updateMetadata: false, throwOnInvalidObject: false });
}

/**
 * The image object `ref` drawn at its own pixel size, as a page reader draws
 * it (its filters, colour space and mask): copied alone into a one-page file
 * the size of its pixels and rendered there. Null when it is not an image.
 */
export async function renderPdfImageObject(bytes: Uint8Array, ref: ObjectRef): Promise<HTMLCanvasElement | null> {
  const source = await load(bytes);
  const imageRef = PDFRef.of(ref.objectNumber, ref.generation);
  const image = source.context.lookup(imageRef);
  const dict = image && "dict" in image ? (image as unknown as { dict: PDFDict }).dict : null;
  if (!dict || lookup(source, dict.get(PDFName.of("Subtype")))?.toString() !== "/Image") return null;
  const widthValue = Number(lookup(source, dict.get(PDFName.of("Width")))?.toString());
  const heightValue = Number(lookup(source, dict.get(PDFName.of("Height")))?.toString());
  if (!(widthValue >= 1) || !(heightValue >= 1)) return null;
  const scale = Math.min(1, ORIGINAL_MAX_SIDE / Math.max(widthValue, heightValue));
  const width = Math.max(1, Math.round(widthValue * scale));
  const height = Math.max(1, Math.round(heightValue * scale));
  const single = await PDFDocument.create({ updateMetadata: false });
  const copied = PDFObjectCopier.for(source.context, single.context).copy(imageRef);
  const page = single.addPage([width, height]);
  page.node.setXObject(PDFName.of("Im0"), copied as PDFRef);
  const content = single.context.flateStream(`q ${width} 0 0 ${height} 0 0 cm /Im0 Do Q`);
  page.node.set(PDFName.of("Contents"), single.context.register(content));
  const singleBytes = await single.save({ objectsPerTick: pdfLibObjectsPerTick(), useObjectStreams: false });
  const pdfjs = await loadCompatiblePdfJs();
  const task = pdfjs.getDocument({ data: singleBytes, ...PDF_JS_DOCUMENT_ASSETS });
  const rendered = await task.promise;
  try {
    const shown = await rendered.getPage(1);
    const viewport = shown.getViewport({ scale: 1 });
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d");
    if (!context) return null;
    await shown.render({ canvas, canvasContext: context, viewport, background: "rgba(0,0,0,0)" }).promise;
    return canvas;
  } finally {
    await rendered.destroy().catch(() => undefined);
  }
}
