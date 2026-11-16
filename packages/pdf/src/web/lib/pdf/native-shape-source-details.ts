/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pdfLibParseSpeed } from "@mrnote/core/web/lib/pdf/pdf-lib-pacing.ts";
/**
 * The parts of a third-party drawing markup that pdf.js does not hand to the
 * page: a Line's own `/L` endpoints in their stored order (pdf.js normalizes
 * them into a min/max box, so the direction and which end carries `/LE` are
 * gone) and a closed shape's `/IC` interior colour (pdf.js paints it into a
 * generated appearance but never reports it).
 *
 * Read once per pdf.js document from its own bytes. These are display/copy
 * hints for the exact loaded revision, never write permission: the native
 * writers re-read the source object themselves.
 */

export interface NativePdfShapeSourceDetails {
  /** `/L` in PDF user space, start point first. */
  line?: readonly [number, number, number, number];
  /** `/IC` as RGB components in 0..1 (gray and CMYK converted). */
  interiorColor?: readonly [number, number, number];
}

export type NativePdfShapeSourceDetailsIndex = ReadonlyMap<
  string,
  NativePdfShapeSourceDetails
>;

const SHAPE_SUBTYPES = new Set(["Square", "Circle", "Line", "Polygon", "PolyLine"]);
const MAX_SOURCE_BYTES = 134_217_728;
const MAX_ANNOTATIONS = 20_000;

function clamp01(value: number): number {
  return Math.min(1, Math.max(0, value));
}

/** /IC as RGB: 1 component gray, 3 RGB, 4 CMYK (naive), else none. */
export function nativeShapeInteriorRgb(
  values: readonly number[],
): readonly [number, number, number] | undefined {
  if (!values.every((value) => Number.isFinite(value))) return undefined;
  if (values.length === 1) {
    const gray = clamp01(values[0]);
    return [gray, gray, gray];
  }
  if (values.length === 3) {
    return [clamp01(values[0]), clamp01(values[1]), clamp01(values[2])];
  }
  if (values.length === 4) {
    const [c, m, y, k] = values.map(clamp01);
    return [(1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k)];
  }
  return undefined;
}

export async function readNativePdfShapeSourceDetails(
  sourceBytes: Uint8Array,
): Promise<NativePdfShapeSourceDetailsIndex | null> {
  if (!sourceBytes.byteLength || sourceBytes.byteLength > MAX_SOURCE_BYTES) {
    return null;
  }
  try {
    const { PDFDocument, PDFName, PDFDict, PDFArray, PDFRef, PDFNumber } =
      await import("pdf-lib");
    // Numbers are never encrypted, so an encrypted file still yields /L and /IC.
    const document = await PDFDocument.load(sourceBytes, {
      parseSpeed: pdfLibParseSpeed(),
      updateMetadata: false,
      ignoreEncryption: true,
    });
    const numbers = (value: unknown): number[] | null => {
      const array = document.context.lookup(value as never);
      if (!(array instanceof PDFArray)) return null;
      const result: number[] = [];
      for (let index = 0; index < array.size(); index += 1) {
        const item = document.context.lookup(array.get(index));
        if (!(item instanceof PDFNumber)) return null;
        result.push(item.asNumber());
      }
      return result;
    };
    const index = new Map<string, NativePdfShapeSourceDetails>();
    let count = 0;
    for (const page of document.getPages()) {
      const annotations = page.node.lookupMaybe(PDFName.of("Annots"), PDFArray);
      if (!annotations) continue;
      count += annotations.size();
      if (count > MAX_ANNOTATIONS) return null;
      for (let position = 0; position < annotations.size(); position += 1) {
        const reference = annotations.get(position);
        if (!(reference instanceof PDFRef)) continue;
        const annotation = document.context.lookup(reference);
        if (!(annotation instanceof PDFDict)) continue;
        const subtype = annotation.get(PDFName.of("Subtype"));
        const subtypeName =
          subtype instanceof PDFName ? subtype.decodeText() : null;
        if (!subtypeName || !SHAPE_SUBTYPES.has(subtypeName)) continue;
        const details: {
          line?: [number, number, number, number];
          interiorColor?: readonly [number, number, number];
        } = {};
        if (subtypeName === "Line") {
          const line = numbers(annotation.get(PDFName.of("L")));
          if (line && line.length === 4 && line.every(Number.isFinite)) {
            details.line = [line[0], line[1], line[2], line[3]];
          }
        }
        const interior = numbers(annotation.get(PDFName.of("IC")));
        const rgb = interior ? nativeShapeInteriorRgb(interior) : undefined;
        if (rgb) details.interiorColor = rgb;
        if (!details.line && !details.interiorColor) continue;
        const ref = `${reference.objectNumber}R${reference.generationNumber || ""}`;
        index.set(ref, details);
      }
    }
    return index;
  } catch {
    return null;
  }
}

interface PdfDataSource {
  getData(): Promise<Uint8Array>;
}

const detailsByDocument = new WeakMap<
  PdfDataSource,
  Promise<NativePdfShapeSourceDetailsIndex | null>
>();

/**
 * One bounded read per pdf.js document incarnation. The caller must still
 * check its active document/render identity after awaiting this result.
 */
export function nativePdfShapeSourceDetailsForDocument(
  document: PdfDataSource,
): Promise<NativePdfShapeSourceDetailsIndex | null> {
  let pending = detailsByDocument.get(document);
  if (!pending) {
    pending = document
      .getData()
      .then(readNativePdfShapeSourceDetails)
      .catch(() => null);
    detailsByDocument.set(document, pending);
  }
  return pending;
}
