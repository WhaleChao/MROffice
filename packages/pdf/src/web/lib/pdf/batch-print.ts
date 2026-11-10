/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFName,
  PDFNumber,
  PDFRef,
  PDFStream,
  concatTransformationMatrix,
  drawObject,
  popGraphicsState,
  pushGraphicsState,
} from "pdf-lib";

import { loadPdfForMutation } from "./encrypted-guard";
import {
  inspectPdfPagePackageSource,
  type PdfPagePackageSourceSnapshot,
  type PdfSavePolicyRevision,
} from "./save-policy";
import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";

export const PDF_BATCH_PRINT_MAX_FILES = 16;
export const PDF_BATCH_PRINT_MAX_SOURCE_BYTES = PDF_EDITABLE_MAX_BYTES;
export const PDF_BATCH_PRINT_MAX_AGGREGATE_BYTES = 256 * 1024 * 1024;
export const PDF_BATCH_PRINT_MAX_PAGES = 10_000;
export const PDF_BATCH_PRINT_MAX_RANGE_LENGTH = 512;

export interface PdfBatchPrintItem {
  readonly id: string;
  readonly name: string;
  readonly bytes: Uint8Array;
  readonly source: PdfPagePackageSourceSnapshot;
  readonly range: string;
}

export interface PdfBatchPrintBuildResult {
  readonly bytes: Uint8Array;
  readonly pageCount: number;
  readonly sourceCount: number;
}

function copyBytes(bytes: Uint8Array | ArrayBuffer): Uint8Array {
  const view = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const copy = new Uint8Array(view.byteLength);
  copy.set(view);
  return copy;
}

function normalizeRange(value: string): string {
  return value
    .replace(/[‒–—―−]/gu, "-")
    .replace(/\s+/gu, "")
    .slice(0, PDF_BATCH_PRINT_MAX_RANGE_LENGTH);
}

export function parsePdfBatchPrintRange(
  value: string,
  pageCount: number,
): readonly number[] {
  if (!Number.isSafeInteger(pageCount) || pageCount < 1) {
    throw new Error("Batch print source has no printable pages");
  }
  const normalized = normalizeRange(value);
  if (normalized === "") {
    return Object.freeze(
      Array.from({ length: pageCount }, (_, index) => index + 1),
    );
  }
  if (value.length > PDF_BATCH_PRINT_MAX_RANGE_LENGTH) {
    throw new Error("Batch print page range is too long");
  }
  const tokens = normalized.split(",");
  if (tokens.some((token) => token === "")) {
    throw new Error("Batch print page range is invalid");
  }
  const pages = new Set<number>();
  for (const token of tokens) {
    const match = /^(\d+)(?:-(\d+))?$/u.exec(token);
    if (!match) throw new Error("Batch print page range is invalid");
    const start = Number(match[1]);
    const end = match[2] === undefined ? start : Number(match[2]);
    if (
      !Number.isSafeInteger(start) ||
      !Number.isSafeInteger(end) ||
      start < 1 ||
      end < start ||
      end > pageCount
    ) {
      throw new Error("Batch print page range is outside the document");
    }
    for (let page = start; page <= end; page += 1) {
      pages.add(page);
      if (pages.size > PDF_BATCH_PRINT_MAX_PAGES) {
        throw new Error("Batch print page limit exceeded");
      }
    }
  }
  if (pages.size === 0) {
    throw new Error("Batch print page range is empty");
  }
  return Object.freeze([...pages].sort((left, right) => left - right));
}

export async function inspectPdfBatchPrintItem(options: {
  readonly id: string;
  readonly name: string;
  readonly bytes: Uint8Array | ArrayBuffer;
  readonly dataItemId: string;
  readonly byteGenerationKey: string;
  readonly documentRevision: PdfSavePolicyRevision;
  readonly range?: string;
}): Promise<PdfBatchPrintItem> {
  const name = options.name.trim();
  const bytes = copyBytes(options.bytes);
  if (
    !name ||
    bytes.byteLength < 1 ||
    bytes.byteLength > PDF_BATCH_PRINT_MAX_SOURCE_BYTES
  ) {
    bytes.fill(0);
    throw new Error("Batch print PDF is empty or exceeds the editable PDF limit");
  }
  try {
    const source = await inspectPdfPagePackageSource({
      dataItemId: options.dataItemId,
      byteGenerationKey: options.byteGenerationKey,
      bytes,
      documentRevision: options.documentRevision,
    });
    parsePdfBatchPrintRange(options.range ?? "", source.pageCount);
    return Object.freeze({
      id: options.id,
      name,
      bytes,
      source,
      range: options.range ?? "",
    });
  } catch (error) {
    bytes.fill(0);
    throw error;
  }
}

export function releasePdfBatchPrintItems(
  items: readonly PdfBatchPrintItem[],
): void {
  for (const item of items) item.bytes.fill(0);
}

function exactPageGeometry(page: {
  getWidth(): number;
  getHeight(): number;
  getRotation(): { angle: number };
}): string {
  return [
    page.getWidth().toFixed(6),
    page.getHeight().toFixed(6),
    String(page.getRotation().angle),
  ].join(":");
}

const ANNOTATION_HIDDEN = 1 << 1;
const ANNOTATION_PRINT = 1 << 2;

function numbersOf(value: unknown, length: number): number[] | null {
  if (!(value instanceof PDFArray) || value.size() !== length) return null;
  const numbers = value
    .asArray()
    .map((item) => (item instanceof PDFNumber ? item.asNumber() : Number.NaN));
  return numbers.every(Number.isFinite) ? numbers : null;
}

/**
 * Draw every printable widget's appearance into its page and take the widget
 * off the page, for a print copy. The appearance is placed the way a reader
 * places it (ISO 32000-1 §12.5.5: its /BBox, through its /Matrix, onto
 * /Rect). A widget the reader would not print (Hidden, or no Print flag) is
 * left out; a field that holds a value but has no appearance cannot be drawn
 * faithfully, so the batch is refused rather than printed blank.
 */
export function flattenWidgetsForPrint(doc: PDFDocument): void {
  for (const page of doc.getPages()) {
    const annots = page.node.lookupMaybe(PDFName.of("Annots"), PDFArray);
    if (!annots) continue;
    const kept: unknown[] = [];
    for (let index = 0; index < annots.size(); index += 1) {
      const raw = annots.get(index);
      const widget = doc.context.lookup(raw);
      if (
        !(widget instanceof PDFDict) ||
        widget.get(PDFName.of("Subtype")) !== PDFName.of("Widget")
      ) {
        kept.push(raw);
        continue;
      }
      const flagsValue = widget.lookup(PDFName.of("F"));
      const flags = flagsValue instanceof PDFNumber ? flagsValue.asNumber() : 0;
      if ((flags & ANNOTATION_HIDDEN) !== 0 || (flags & ANNOTATION_PRINT) === 0) {
        continue;
      }
      const rect = numbersOf(widget.lookup(PDFName.of("Rect")), 4);
      const appearances = widget.lookupMaybe(PDFName.of("AP"), PDFDict);
      let normal = appearances?.get(PDFName.of("N"));
      const normalResolved = normal ? doc.context.lookup(normal) : undefined;
      if (normalResolved instanceof PDFDict && !(normalResolved instanceof PDFStream)) {
        const state = widget.get(PDFName.of("AS"));
        normal = state instanceof PDFName ? normalResolved.get(state) : undefined;
      }
      const stream = normal ? doc.context.lookup(normal) : undefined;
      if (!rect || !(normal instanceof PDFRef) || !(stream instanceof PDFStream)) {
        const value = widget.lookup(PDFName.of("V"));
        const parent = widget.lookupMaybe(PDFName.of("Parent"), PDFDict);
        if (value !== undefined || parent?.lookup(PDFName.of("V")) !== undefined) {
          throw new Error("Batch print cannot draw a filled form field without an appearance");
        }
        continue;
      }
      const box = numbersOf(stream.dict.lookup(PDFName.of("BBox")), 4);
      const matrix = numbersOf(stream.dict.lookup(PDFName.of("Matrix")), 6) ?? [1, 0, 0, 1, 0, 0];
      if (!box) throw new Error("Batch print form appearance is malformed");
      const corners = [
        [box[0], box[1]],
        [box[2], box[1]],
        [box[0], box[3]],
        [box[2], box[3]],
      ].map(([x, y]) => [
        matrix[0] * x + matrix[2] * y + matrix[4],
        matrix[1] * x + matrix[3] * y + matrix[5],
      ]);
      const xs = corners.map((corner) => corner[0]);
      const ys = corners.map((corner) => corner[1]);
      const left = Math.min(...xs);
      const bottom = Math.min(...ys);
      const width = Math.max(...xs) - left;
      const height = Math.max(...ys) - bottom;
      const rectLeft = Math.min(rect[0], rect[2]);
      const rectBottom = Math.min(rect[1], rect[3]);
      const rectWidth = Math.abs(rect[2] - rect[0]);
      const rectHeight = Math.abs(rect[3] - rect[1]);
      if (width <= 0 || height <= 0 || rectWidth <= 0 || rectHeight <= 0) continue;
      const scaleX = rectWidth / width;
      const scaleY = rectHeight / height;
      const name = page.node.newXObject("PrintWidget", normal);
      page.pushOperators(
        pushGraphicsState(),
        concatTransformationMatrix(
          scaleX,
          0,
          0,
          scaleY,
          rectLeft - left * scaleX,
          rectBottom - bottom * scaleY,
        ),
        drawObject(name),
        popGraphicsState(),
      );
    }
    page.node.set(PDFName.of("Annots"), doc.context.obj(kept as never));
  }
}

export async function buildPdfBatchPrintDocument(
  items: readonly PdfBatchPrintItem[],
): Promise<PdfBatchPrintBuildResult> {
  if (items.length < 1 || items.length > PDF_BATCH_PRINT_MAX_FILES) {
    throw new Error("Batch print accepts 1 to 16 PDF files");
  }
  const output = await PDFDocument.create();
  const expectedGeometry: string[] = [];
  let aggregateBytes = 0;
  let pageCount = 0;
  const ids = new Set<string>();

  for (const item of items) {
    if (ids.has(item.id)) throw new Error("Duplicate batch print source");
    ids.add(item.id);
    aggregateBytes += item.bytes.byteLength;
    if (
      !Number.isSafeInteger(aggregateBytes) ||
      aggregateBytes > PDF_BATCH_PRINT_MAX_AGGREGATE_BYTES
    ) {
      throw new Error("Batch print sources exceed 256 MiB");
    }
    const reinspection = await inspectPdfPagePackageSource({
      dataItemId: item.source.source.identity.dataItemId,
      byteGenerationKey: item.source.source.identity.byteGenerationKey,
      bytes: item.bytes,
      documentRevision: item.source.source.identity.documentRevision,
    });
    if (
      reinspection.source.identity.sha256 !==
        item.source.source.identity.sha256 ||
      reinspection.source.profileDigest !== item.source.source.profileDigest ||
      reinspection.pageCount !== item.source.pageCount ||
      reinspection.byteLength !== item.source.byteLength
    ) {
      throw new Error("Batch print source identity changed");
    }
    const pages = parsePdfBatchPrintRange(item.range, item.source.pageCount);
    pageCount += pages.length;
    if (pageCount > PDF_BATCH_PRINT_MAX_PAGES) {
      throw new Error("Batch print page limit exceeded");
    }
    const source = await loadPdfForMutation(item.bytes);
    // Form fields print from their appearances, which a merged document
    // loses with the AcroForm it cannot carry: PDFium printed both filled
    
    flattenWidgetsForPrint(source);
    const indices = pages.map((page) => page - 1);
    for (const index of indices) {
      expectedGeometry.push(exactPageGeometry(source.getPage(index)));
    }
    const copied = await output.copyPages(source, indices);
    for (const page of copied) output.addPage(page);
  }

  if (pageCount < 1 || output.getPageCount() !== pageCount) {
    throw new Error("Batch print produced no pages");
  }
  const bytes = copyBytes(await output.save());
  try {
    const reopened = await loadPdfForMutation(bytes);
    if (
      reopened.getPageCount() !== pageCount ||
      reopened
        .getPages()
        .some(
          (page, index) => exactPageGeometry(page) !== expectedGeometry[index],
        )
    ) {
      throw new Error("Batch print output failed page verification");
    }
    return Object.freeze({
      bytes,
      pageCount,
      sourceCount: items.length,
    });
  } catch (error) {
    bytes.fill(0);
    throw error;
  }
}
