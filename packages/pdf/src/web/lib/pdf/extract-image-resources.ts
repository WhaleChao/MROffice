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
  PDFNull,
  PDFNumber,
  PDFRawStream,
  PDFRef,
  decodePDFRawStream,
  type PDFObject,
} from "pdf-lib";

import { loadPdfForInspection } from "./encrypted-guard";
import { encodePng } from "./png-encode";
import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";

export const PDF_IMAGE_EXTRACTION_MAX_SOURCE_BYTES = PDF_EDITABLE_MAX_BYTES;
export const PDF_IMAGE_EXTRACTION_MAX_IMAGES = 500;
export const PDF_IMAGE_EXTRACTION_MAX_TOTAL_BYTES = PDF_EDITABLE_MAX_BYTES;
export const PDF_IMAGE_EXTRACTION_MAX_ZIP_BYTES = 96 * 1024 * 1024;

type ExtractableImageFormat = "jpeg" | "jpeg2000" | "png";

export interface PdfExtractableImageResource {
  readonly id: string;
  readonly objectRef: string;
  readonly filename: string;
  readonly format: ExtractableImageFormat;
  readonly mimeType: "image/jpeg" | "image/jp2" | "image/png";
  readonly width: number;
  readonly height: number;
  readonly bitsPerComponent: number;
  readonly colorSpace: string;
  readonly encodedByteLength: number;
  readonly encodedSha256: string;
  readonly pageNumbers: readonly number[];
  readonly resourceNames: readonly string[];
  readonly resourceReferenceCount: number;
  readonly hasSoftMask: boolean;
  readonly bytes: Uint8Array;
}

export type PdfSkippedImageReason =
  | "aggregate-limit"
  | "direct-resource"
  | "image-mask"
  | "invalid-dimensions"
  | "unsupported-filter";

export interface PdfSkippedImageResource {
  readonly objectRef: string | null;
  readonly filter: string;
  readonly pageNumbers: readonly number[];
  readonly resourceNames: readonly string[];
  readonly resourceReferenceCount: number;
  readonly reason: PdfSkippedImageReason;
}

export interface PdfImageExtractionPlan {
  readonly version: "mrai.pdf-image-extraction.v1";
  readonly sourceSha256: string;
  readonly sourceByteLength: number;
  readonly pageCount: number;
  readonly images: readonly PdfExtractableImageResource[];
  readonly skipped: readonly PdfSkippedImageResource[];
  readonly caveats: readonly [
    "inline-images-not-enumerated",
    "soft-masks-not-composited",
    "original-encoded-streams-only",
  ];
}

export type SelectedOriginalImageUnavailableReason =
  | "source-mismatch"
  | "selection-out-of-bounds"
  | "resource-not-found"
  | "resource-ambiguous"
  | "unsupported-format"
  | "soft-mask"
  | "ambiguous-color";

export class SelectedOriginalImageUnavailableError extends Error {
  constructor(readonly reason: SelectedOriginalImageUnavailableReason) {
    super(reason);
    this.name = "SelectedOriginalImageUnavailableError";
  }
}

export interface SelectedOriginalImageResource {
  readonly objectRef: string;
  readonly encodedSha256: string;
  readonly mimeType: "image/jpeg" | "image/jp2";
  readonly filename: string;
  readonly bytes: Uint8Array;
}

interface MutableImageRecord {
  readonly ref: PDFRef;
  readonly stream: PDFRawStream;
  readonly pages: Set<number>;
  readonly resourceNames: Set<string>;
  resourceReferenceCount: number;
}

interface DirectSkipRecord {
  readonly page: number;
  readonly resourceName: string;
}

const ZIP_DATE = new Date("1980-01-01T00:00:00.000Z");
const MAX_FORM_DEPTH = 16;
const MAX_RESOURCE_EDGES = 4_000;

// A dangling reference resolves to undefined; every caller narrows the result
// with instanceof, so a missing object reads as "not that kind".
function resolveObject(
  doc: PDFDocument,
  value: PDFObject,
): PDFObject | undefined {
  return value instanceof PDFRef ? doc.context.lookup(value) : value;
}

function resolveDict(
  doc: PDFDocument,
  value: PDFObject | undefined,
): PDFDict | null {
  if (!value) return null;
  const resolved = resolveObject(doc, value);
  return resolved instanceof PDFDict ? resolved : null;
}

function numberEntry(
  doc: PDFDocument,
  dictionary: PDFDict,
  key: string,
): number | null {
  const value = dictionary.get(PDFName.of(key));
  if (!value) return null;
  const resolved = resolveObject(doc, value);
  return resolved instanceof PDFNumber ? resolved.asNumber() : null;
}

function nameEntry(
  doc: PDFDocument,
  dictionary: PDFDict,
  key: string,
): string | null {
  const value = dictionary.get(PDFName.of(key));
  if (!value) return null;
  const resolved = resolveObject(doc, value);
  if (resolved instanceof PDFName) return resolved.asString();
  if (resolved instanceof PDFArray && resolved.size() === 1) {
    const member = resolveObject(doc, resolved.get(0));
    return member instanceof PDFName ? member.asString() : null;
  }
  return null;
}

function displayEntry(dictionary: PDFDict, key: string): string {
  const value = dictionary.get(PDFName.of(key));
  return value ? String(value) : "none";
}

function normalizedFilter(value: string | null): string {
  return (value ?? "none").replace(/^\//, "");
}

function exactFormat(filter: string): {
  format: ExtractableImageFormat;
  extension: "jpg" | "jp2";
  mimeType: "image/jpeg" | "image/jp2";
} | null {
  if (filter === "DCTDecode" || filter === "DCT") {
    return { format: "jpeg", extension: "jpg", mimeType: "image/jpeg" };
  }
  if (filter === "JPXDecode") {
    return {
      format: "jpeg2000",
      extension: "jp2",
      mimeType: "image/jp2",
    };
  }
  return null;
}

const MAX_DECODED_IMAGE_BYTES = 64 * 1024 * 1024;

function resolveMaybe(
  doc: PDFDocument,
  value: PDFObject | undefined,
): PDFObject | undefined {
  return value === undefined ? undefined : resolveObject(doc, value);
}

function singleFilter(dictionary: PDFDict, doc: PDFDocument): string | null | false {
  const filter = resolveMaybe(doc, dictionary.get(PDFName.of("Filter")));
  if (filter === undefined) return null;
  if (filter instanceof PDFName) return filter.asString().replace(/^\//u, "");
  if (filter instanceof PDFArray && filter.size() === 1) {
    const only = resolveObject(doc, filter.get(0));
    return only instanceof PDFName ? only.asString().replace(/^\//u, "") : false;
  }
  return false;
}

/** The number of colour components a supported colour space has, or null. */
function colorComponents(
  doc: PDFDocument,
  value: PDFObject | undefined,
): 1 | 3 | null {
  const space = resolveMaybe(doc, value);
  if (space instanceof PDFName) {
    const name = space.asString();
    return name === "/DeviceGray" ? 1 : name === "/DeviceRGB" ? 3 : null;
  }
  if (space instanceof PDFArray && space.size() === 2) {
    const kind = resolveObject(doc, space.get(0));
    const profile = resolveObject(doc, space.get(1));
    if (kind instanceof PDFName && kind.asString() === "/ICCBased" && profile instanceof PDFRawStream) {
      const n = resolveMaybe(doc, profile.dict.get(PDFName.of("N")));
      const count = n instanceof PDFNumber ? n.asNumber() : 0;
      return count === 1 ? 1 : count === 3 ? 3 : null;
    }
  }
  return null;
}









export function unpredictPdfImageSamples(
  data: Uint8Array,
  parameters: Readonly<{ predictor: number; colors: number; bitsPerComponent: number; columns: number }>,
): Uint8Array | null {
  const { predictor, colors, bitsPerComponent, columns } = parameters;
  if (predictor === 1) return data;
  if (bitsPerComponent !== 8 || colors < 1 || colors > 4 || columns < 1) return null;
  const bpp = colors;
  const rowBytes = columns * colors;
  if (predictor === 2) {
    const out = Uint8Array.from(data);
    for (let row = 0; row + rowBytes <= out.length; row += rowBytes) {
      for (let index = bpp; index < rowBytes; index += 1) {
        out[row + index] = (out[row + index]! + out[row + index - bpp]!) & 0xff;
      }
    }
    return out;
  }
  if (predictor < 10 || predictor > 15) return null;
  const rows = Math.floor(data.length / (rowBytes + 1));
  const out = new Uint8Array(rows * rowBytes);
  for (let row = 0; row < rows; row += 1) {
    const type = data[row * (rowBytes + 1)]!;
    const source = row * (rowBytes + 1) + 1;
    const target = row * rowBytes;
    for (let index = 0; index < rowBytes; index += 1) {
      const raw = data[source + index]!;
      const left = index >= bpp ? out[target + index - bpp]! : 0;
      const up = row > 0 ? out[target - rowBytes + index]! : 0;
      const upLeft = row > 0 && index >= bpp ? out[target - rowBytes + index - bpp]! : 0;
      let value: number;
      if (type === 0) value = raw;
      else if (type === 1) value = raw + left;
      else if (type === 2) value = raw + up;
      else if (type === 3) value = raw + ((left + up) >> 1);
      else if (type === 4) {
        const estimate = left + up - upLeft;
        const toLeft = Math.abs(estimate - left);
        const toUp = Math.abs(estimate - up);
        const toUpLeft = Math.abs(estimate - upLeft);
        value = raw + (toLeft <= toUp && toLeft <= toUpLeft ? left : toUp <= toUpLeft ? up : upLeft);
      } else return null;
      out[target + index] = value & 0xff;
    }
  }
  return out;
}

function decodedSamples(
  doc: PDFDocument,
  stream: PDFRawStream,
  expected: number,
  layout?: Readonly<{ colors: number; columns: number }>,
): Uint8Array | null {
  const filter = singleFilter(stream.dict, doc);
  if (filter === false) return null;
  if (filter !== null && filter !== "FlateDecode" && filter !== "Fl") return null;
  if (stream.dict.get(PDFName.of("Decode")) !== undefined) return null;
  let decoded: Uint8Array;
  try {
    decoded = filter === null ? Uint8Array.from(stream.getContents()) : decodePDFRawStream(stream).decode();
  } catch {
    return null;
  }
  const parameters = resolveMaybe(doc, stream.dict.get(PDFName.of("DecodeParms")));
  const onlyParameters =
    parameters instanceof PDFArray && parameters.size() === 1
      ? resolveObject(doc, parameters.get(0))
      : parameters;
  if (filter !== null && onlyParameters instanceof PDFDict) {
    const entry = (key: string, fallback: number) => numberEntry(doc, onlyParameters, key) ?? fallback;
    const unpredicted = unpredictPdfImageSamples(decoded, {
      predictor: entry("Predictor", 1),
      colors: entry("Colors", layout?.colors ?? 1),
      bitsPerComponent: entry("BitsPerComponent", 8),
      columns: entry("Columns", layout?.columns ?? 1),
    });
    if (!unpredicted) return null;
    decoded = unpredicted;
  } else if (filter !== null && onlyParameters !== undefined && onlyParameters !== PDFNull) {
    return null;
  }
  return decoded.length >= expected ? decoded.subarray(0, expected) : null;
}

/** A lossless 8-bit RGB or gray image (with an optional same-size 8-bit
 * gray soft mask) as PNG bytes, or null when it is anything else. */
function losslessImageAsPng(
  doc: PDFDocument,
  stream: PDFRawStream,
  width: number,
  height: number,
  bits: number,
): Uint8Array | null {
  if (bits !== 8) return null;
  const components = colorComponents(doc, stream.dict.get(PDFName.of("ColorSpace")));
  if (!components) return null;
  const pixelCount = width * height;
  if (pixelCount * (components + 1) > MAX_DECODED_IMAGE_BYTES) return null;
  const color = decodedSamples(doc, stream, pixelCount * components, { colors: components, columns: width });
  if (!color) return null;
  const maskValue = resolveMaybe(doc, stream.dict.get(PDFName.of("SMask")));
  if (maskValue === undefined) {
    return encodePng(width, height, components, color);
  }
  if (!(maskValue instanceof PDFRawStream)) return null;
  const maskWidth = numberEntry(doc, maskValue.dict, "Width");
  const maskHeight = numberEntry(doc, maskValue.dict, "Height");
  const maskBits = numberEntry(doc, maskValue.dict, "BitsPerComponent");
  if (maskWidth !== width || maskHeight !== height || maskBits !== 8) return null;
  const alpha = decodedSamples(doc, maskValue, pixelCount, { colors: 1, columns: width });
  if (!alpha) return null;
  const channels = components + 1;
  const pixels = new Uint8Array(pixelCount * channels);
  for (let index = 0; index < pixelCount; index += 1) {
    for (let channel = 0; channel < components; channel += 1) {
      pixels[index * channels + channel] = color[index * components + channel]!;
    }
    pixels[index * channels + components] = alpha[index]!;
  }
  return encodePng(width, height, channels as 2 | 4, pixels);
}

async function sha256(bytes: Uint8Array): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) throw new Error("SHA-256 is unavailable");
  const digest = await subtle.digest("SHA-256", Uint8Array.from(bytes));
  return Array.from(new Uint8Array(digest), (value) =>
    value.toString(16).padStart(2, "0"),
  ).join("");
}

function sameBytes(left: Uint8Array, right: Uint8Array): boolean {
  return (
    left.byteLength === right.byteLength &&
    left.every((value, index) => value === right[index])
  );
}

export async function resolveSelectedOriginalImageResource(
  plan: PdfImageExtractionPlan,
  selection: Readonly<{
    page: number;
    order: number;
    objectRef: string;
    expectedSourceSha256: string;
  }>,
): Promise<SelectedOriginalImageResource> {
  if (plan.sourceSha256 !== selection.expectedSourceSha256.toLowerCase()) {
    throw new SelectedOriginalImageUnavailableError("source-mismatch");
  }
  if (
    !Number.isSafeInteger(selection.page) ||
    selection.page < 1 ||
    selection.page > plan.pageCount ||
    !Number.isSafeInteger(selection.order) ||
    selection.order < 0
  ) {
    throw new SelectedOriginalImageUnavailableError("selection-out-of-bounds");
  }
  const refMatch = /^(\d+) (\d+) R$/.exec(selection.objectRef);
  if (!refMatch) {
    throw new SelectedOriginalImageUnavailableError("resource-not-found");
  }
  const matches = plan.images.filter(
    (image) => image.objectRef === selection.objectRef,
  );
  if (matches.length === 0) {
    throw new SelectedOriginalImageUnavailableError("resource-not-found");
  }
  const canonical = matches[0];
  if (
    matches.some(
      (image) =>
        image.encodedSha256 !== canonical.encodedSha256 ||
        image.mimeType !== canonical.mimeType ||
        image.encodedByteLength !== canonical.encodedByteLength ||
        !sameBytes(image.bytes, canonical.bytes),
    )
  ) {
    throw new SelectedOriginalImageUnavailableError("resource-ambiguous");
  }
  if (
    (canonical.format !== "jpeg" && canonical.format !== "jpeg2000") ||
    (canonical.mimeType !== "image/jpeg" && canonical.mimeType !== "image/jp2")
  ) {
    throw new SelectedOriginalImageUnavailableError("unsupported-format");
  }
  if (canonical.hasSoftMask) {
    throw new SelectedOriginalImageUnavailableError("soft-mask");
  }
  if (
    !["/DeviceGray", "/DeviceRGB", "/DeviceCMYK"].includes(canonical.colorSpace)
  ) {
    throw new SelectedOriginalImageUnavailableError("ambiguous-color");
  }
  const bytes = canonical.bytes.slice();
  if (
    bytes.byteLength !== canonical.encodedByteLength ||
    (await sha256(bytes)) !== canonical.encodedSha256
  ) {
    bytes.fill(0);
    throw new SelectedOriginalImageUnavailableError("resource-ambiguous");
  }
  return Object.freeze({
    objectRef: canonical.objectRef,
    encodedSha256: canonical.encodedSha256,
    mimeType: canonical.mimeType,
    filename: `original-page-${selection.page}-image-${selection.order + 1}-${refMatch[1]}.${canonical.format === "jpeg" ? "jpg" : "jp2"}`,
    bytes,
  });
}

function imageMask(dictionary: PDFDict): boolean {
  return String(dictionary.get(PDFName.of("ImageMask")) ?? "false") === "true";
}

function softMask(dictionary: PDFDict): boolean {
  const value = dictionary.get(PDFName.of("SMask"));
  return Boolean(value && String(value) !== "/None");
}

function refOrder(left: MutableImageRecord, right: MutableImageRecord): number {
  return (
    left.ref.objectNumber - right.ref.objectNumber ||
    left.ref.generationNumber - right.ref.generationNumber
  );
}

export async function inspectPdfImageResources(
  input: Uint8Array | ArrayBuffer,
): Promise<PdfImageExtractionPlan> {
  const source = input instanceof Uint8Array ? input : new Uint8Array(input);
  if (
    source.byteLength < 1 ||
    source.byteLength > PDF_IMAGE_EXTRACTION_MAX_SOURCE_BYTES
  ) {
    throw new Error("PDF image extraction source exceeds the editable PDF limit");
  }
  const owned = Uint8Array.from(source);
  try {
    const doc = await loadPdfForInspection(owned, { updateMetadata: false });
    if (!doc) throw new Error("Encrypted PDF image extraction is unavailable");
    const records = new Map<string, MutableImageRecord>();
    const directSkips = new Map<string, DirectSkipRecord>();
    let resourceEdges = 0;

    const walkResources = (
      resources: PDFDict | null,
      page: number,
      scope: string,
      activeForms: ReadonlySet<string>,
      depth: number,
    ): void => {
      if (!resources) return;
      if (depth > MAX_FORM_DEPTH) {
        throw new Error("PDF image resource nesting exceeds the safety limit");
      }
      const xObjects = resources.lookupMaybe(PDFName.of("XObject"), PDFDict);
      if (!xObjects) return;
      for (const [name, value] of xObjects.entries()) {
        resourceEdges += 1;
        if (resourceEdges > MAX_RESOURCE_EDGES) {
          throw new Error("PDF image resource graph exceeds the safety limit");
        }
        const resourceName = `${scope}/${name.asString()}`;
        if (!(value instanceof PDFRef)) {
          const direct = resolveObject(doc, value);
          if (
            direct instanceof PDFRawStream &&
            String(direct.dict.get(PDFName.of("Subtype"))) === "/Image"
          ) {
            directSkips.set(`${page}:${resourceName}`, {
              page,
              resourceName,
            });
          } else if (
            direct instanceof PDFRawStream &&
            String(direct.dict.get(PDFName.of("Subtype"))) === "/Form"
          ) {
            walkResources(
              resolveDict(doc, direct.dict.get(PDFName.of("Resources"))),
              page,
              resourceName,
              activeForms,
              depth + 1,
            );
          }
          continue;
        }
        const object = doc.context.lookup(value);
        if (!(object instanceof PDFRawStream)) continue;
        const subtype = String(object.dict.get(PDFName.of("Subtype")) ?? "");
        if (subtype === "/Image") {
          const key = value.toString();
          let record = records.get(key);
          if (!record) {
            if (records.size >= PDF_IMAGE_EXTRACTION_MAX_IMAGES) {
              throw new Error("PDF image resource count exceeds 500");
            }
            record = {
              ref: value,
              stream: object,
              pages: new Set<number>(),
              resourceNames: new Set<string>(),
              resourceReferenceCount: 0,
            };
            records.set(key, record);
          }
          record.pages.add(page);
          record.resourceNames.add(resourceName);
          record.resourceReferenceCount += 1;
          continue;
        }
        if (subtype !== "/Form") continue;
        const formKey = value.toString();
        if (activeForms.has(formKey)) continue;
        const nextForms = new Set(activeForms);
        nextForms.add(formKey);
        walkResources(
          resolveDict(doc, object.dict.get(PDFName.of("Resources"))),
          page,
          resourceName,
          nextForms,
          depth + 1,
        );
      }
    };

    for (const [index, page] of doc.getPages().entries()) {
      walkResources(
        resolveDict(
          doc,
          page.node.getInheritableAttribute(PDFName.of("Resources")),
        ),
        index + 1,
        `page-${index + 1}`,
        new Set<string>(),
        0,
      );
    }

    const images: PdfExtractableImageResource[] = [];
    const skipped: PdfSkippedImageResource[] = Array.from(
      directSkips.values(),
      (entry) => ({
        objectRef: null,
        filter: "unknown",
        pageNumbers: [entry.page],
        resourceNames: [entry.resourceName],
        resourceReferenceCount: 1,
        reason: "direct-resource" as const,
      }),
    );
    let aggregateBytes = 0;
    for (const record of Array.from(records.values()).sort(refOrder)) {
      const pages = Array.from(record.pages).sort(
        (left, right) => left - right,
      );
      const names = Array.from(record.resourceNames).sort();
      const filter = normalizedFilter(
        nameEntry(doc, record.stream.dict, "Filter") ??
          displayEntry(record.stream.dict, "Filter"),
      );
      const format = exactFormat(filter);
      const width = numberEntry(doc, record.stream.dict, "Width");
      const height = numberEntry(doc, record.stream.dict, "Height");
      const bits = numberEntry(doc, record.stream.dict, "BitsPerComponent");
      const skippedBase = {
        objectRef: record.ref.toString(),
        filter,
        pageNumbers: pages,
        resourceNames: names,
        resourceReferenceCount: record.resourceReferenceCount,
      };
      if (imageMask(record.stream.dict)) {
        skipped.push({ ...skippedBase, reason: "image-mask" });
        continue;
      }
      if (
        !Number.isSafeInteger(width) ||
        !Number.isSafeInteger(height) ||
        !Number.isSafeInteger(bits) ||
        (width ?? 0) < 1 ||
        (height ?? 0) < 1 ||
        (bits ?? 0) < 1 ||
        (bits ?? 0) > 16 ||
        (width ?? 0) * (height ?? 0) > 100_000_000
      ) {
        skipped.push({ ...skippedBase, reason: "invalid-dimensions" });
        continue;
      }
      // A lossless (Flate or unfiltered) RGB/gray image is decoded and saved
      // as PNG, its soft mask as the alpha channel: an icon is not a JPEG,
      
      const losslessPng = format
        ? null
        : losslessImageAsPng(doc, record.stream, width!, height!, bits!);
      if (!format && !losslessPng) {
        skipped.push({ ...skippedBase, reason: "unsupported-filter" });
        continue;
      }
      const bytes = losslessPng ?? Uint8Array.from(record.stream.getContents());
      if (
        bytes.byteLength < 1 ||
        aggregateBytes + bytes.byteLength > PDF_IMAGE_EXTRACTION_MAX_TOTAL_BYTES
      ) {
        bytes.fill(0);
        skipped.push({ ...skippedBase, reason: "aggregate-limit" });
        continue;
      }
      aggregateBytes += bytes.byteLength;
      const index = images.length + 1;
      images.push({
        id: `image-${String(index).padStart(3, "0")}`,
        objectRef: record.ref.toString(),
        filename: `images/image-${String(index).padStart(3, "0")}-p${String(pages[0] ?? 0).padStart(3, "0")}-${record.ref.objectNumber}-${record.ref.generationNumber}.${format ? format.extension : "png"}`,
        format: format ? format.format : "png",
        mimeType: format ? format.mimeType : "image/png",
        width: width!,
        height: height!,
        bitsPerComponent: bits!,
        colorSpace: displayEntry(record.stream.dict, "ColorSpace"),
        encodedByteLength: bytes.byteLength,
        encodedSha256: await sha256(bytes),
        pageNumbers: pages,
        resourceNames: names,
        resourceReferenceCount: record.resourceReferenceCount,
        hasSoftMask: softMask(record.stream.dict),
        bytes,
      });
    }

    return {
      version: "mrai.pdf-image-extraction.v1",
      sourceSha256: await sha256(owned),
      sourceByteLength: source.byteLength,
      pageCount: doc.getPageCount(),
      images: Object.freeze(images),
      skipped: Object.freeze(
        skipped.sort((left, right) =>
          `${left.objectRef ?? ""}:${left.resourceNames.join("\u0000")}`.localeCompare(
            `${right.objectRef ?? ""}:${right.resourceNames.join("\u0000")}`,
          ),
        ),
      ),
      caveats: [
        "inline-images-not-enumerated",
        "soft-masks-not-composited",
        "original-encoded-streams-only",
      ],
    };
  } finally {
    owned.fill(0);
  }
}

function manifestFor(plan: PdfImageExtractionPlan) {
  return {
    version: plan.version,
    sourceSha256: plan.sourceSha256,
    sourceByteLength: plan.sourceByteLength,
    pageCount: plan.pageCount,
    imageCount: plan.images.length,
    skippedCount: plan.skipped.length,
    caveats: plan.caveats,
    images: plan.images.map((image) => ({
      id: image.id,
      objectRef: image.objectRef,
      filename: image.filename,
      format: image.format,
      mimeType: image.mimeType,
      width: image.width,
      height: image.height,
      bitsPerComponent: image.bitsPerComponent,
      colorSpace: image.colorSpace,
      encodedByteLength: image.encodedByteLength,
      encodedSha256: image.encodedSha256,
      pageNumbers: image.pageNumbers,
      resourceNames: image.resourceNames,
      resourceReferenceCount: image.resourceReferenceCount,
      hasSoftMask: image.hasSoftMask,
    })),
    skipped: plan.skipped,
  };
}

export async function buildPdfImageExtractionZip(
  plan: PdfImageExtractionPlan,
): Promise<Blob> {
  if (plan.version !== "mrai.pdf-image-extraction.v1") {
    throw new Error("Unsupported PDF image extraction plan");
  }
  if (
    plan.images.length < 1 ||
    plan.images.length > PDF_IMAGE_EXTRACTION_MAX_IMAGES
  ) {
    throw new Error("PDF image extraction has no downloadable images");
  }
  const { default: JSZip } = await import("jszip");
  const zip = new JSZip();
  let aggregate = 0;
  const filenames = new Set<string>();
  for (const image of plan.images) {
    if (filenames.has(image.filename)) {
      throw new Error("Duplicate PDF image extraction filename");
    }
    filenames.add(image.filename);
    if (
      image.bytes.byteLength !== image.encodedByteLength ||
      (await sha256(image.bytes)) !== image.encodedSha256
    ) {
      throw new Error("PDF image extraction bytes changed after review");
    }
    aggregate += image.bytes.byteLength;
    if (aggregate > PDF_IMAGE_EXTRACTION_MAX_TOTAL_BYTES) {
      throw new Error("PDF image extraction exceeds the aggregate limit");
    }
    zip.file(image.filename, image.bytes, {
      binary: true,
      compression: "STORE",
      date: ZIP_DATE,
    });
  }
  zip.file("manifest.json", JSON.stringify(manifestFor(plan), null, 2), {
    compression: "DEFLATE",
    date: ZIP_DATE,
  });
  const bytes = await zip.generateAsync({
    type: "uint8array",
    mimeType: "application/zip",
    compression: "STORE",
    platform: "UNIX",
  });
  if (bytes.byteLength > PDF_IMAGE_EXTRACTION_MAX_ZIP_BYTES) {
    bytes.fill(0);
    throw new Error("PDF image extraction ZIP exceeds the 96 MiB limit");
  }
  return new Blob([Uint8Array.from(bytes)], { type: "application/zip" });
}

export function releasePdfImageExtractionPlan(
  plan: PdfImageExtractionPlan | null,
): void {
  if (!plan) return;
  for (const image of plan.images) image.bytes.fill(0);
}
