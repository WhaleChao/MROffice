/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Lossless PDF size optimization ("Reduce file size"), the client-side analogue
 * of Acrobat's "Save As Optimized" — but ONLY the structural / lossless half.
 *
 * What this does
 * --------------
 * Re-loads the PDF with pdf-lib and re-serializes it with object-stream
 * compression (`useObjectStreams: true`). Re-writing the document through
 * pdf-lib's writer:
 *   - Packs indirect objects into compressed object streams + a cross-reference
 *     stream (PDF 1.5 "cross-reference streams"). Many PDFs produced by other
 *     tools still ship a classic uncompressed xref TABLE + loose objects; the
 *     rebuild folds those into compressed streams, which is where the savings
 *     come from.
 *   - Drops objects that are no longer reachable from the trailer (orphaned
 *     by prior incremental edits). The writer itself emits every object it
 *     loaded, so the pass walks the object graph and removes the rest first
 *     (`pruneUnreachablePdfObjects`).
 *   - Collapses incremental-update layers into a single linear write (no
 *     appended "previous version" bodies).
 *
 * What this does NOT do (documented honestly)
 * -------------------------------------------
 *   - It does NOT re-encode, downsample, or recompress IMAGES. pdf-lib has no
 *     image transcoder, so a scan-heavy PDF (its bytes dominated by DCT/JPEG or
 *     Flate image XObjects) will see little to no shrink — image downsampling is
 *     explicitly OUT OF SCOPE. That is the lossy half Acrobat does and we don't.
 *   - It does NOT subset / strip fonts, flatten transparency, or touch content
 *     streams.
 *   - On a PDF that is ALREADY written with object streams and has no orphaned
 *     objects (e.g. a file we just produced via "Download with annotations",
 *     which goes through this same pdf-lib writer), the re-save can come out the
 *     SAME size or even a few bytes LARGER. To never make a file bigger, the
 *     result is the smaller of {re-saved bytes, original bytes}; when the
 *     original wins, the returned bytes are the original (so `savedBytes === 0`).
 *     Callers must therefore report the real delta and never claim a saving that
 *     didn't happen.
 *
 * The lossless pass is a pure-ish async function (no DOM, no network) so it is
 * unit-testable with an in-memory PDF, mirroring lib/pdf/flatten-apply.ts (the
 * other module that drives pdf-lib).
 *
 * The LOSSY half lives server-side: `compressPdfServerSideWithStats` posts to
 * `/api/v1/ai/notebook/pdf-compress` (pikepdf + PIL image downsampling + JPEG
 * re-encode at quality presets) and adapts the response into the same
 * {@link CompressPdfResult} shape so the UI's honest stats display is shared.
 * The endpoint carries its own never-larger guarantee (unshrinkable input
 * returns the original bytes with equal size headers).
 */

export {
  PDF_COMPRESS_LOSSLESS_PRESET,
  PDF_COMPRESS_SERVER_PRESETS,
  type CompressPdfOptions,
  type CompressPdfResult,
  type PdfCompressLevel,
  type PdfCompressLosslessPreset,
  type PdfCompressServerPreset,
} from "./compress-contract";
import type { CompressPdfResult, PdfCompressServerPreset } from "./compress-contract";
import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";

// shapes, asynchronous, the page's working head named by digest under the
// page a writer set is bound to (`createPdfDocumentWriters(pageId)`).
export { createPdfDocumentWriters, type PdfDocumentWriters } from "./remote/pdf-document-writers";

/**
 * Lossy server-side compression (image downsampling + JPEG re-encode) at a
 * quality preset, adapted into the same {@link CompressPdfResult} shape as
 * the local lossless pass so callers reuse one stats display.
 *
 * The ORIGINAL size comes from the `X-Original-Size` header (the server read
 * the stored bytes — the client never downloads them for this path), and the
 * server's never-larger guarantee means `savedBytes === 0` is an honest
 * "could not shrink" outcome carrying the original bytes.
 *
 * @throws `ApiRequestError` from `apiFetchRaw` on any non-2xx response
 *   (e.g. 400 `pdf_encrypted`, 413, 501 `feature_disabled`). Callers wiring
 *   this into UI MUST catch, mirroring the local pass's contract.
 */
export async function compressPdfServerSideWithStats(
  dataItemId: string,
  preset: PdfCompressServerPreset,
): Promise<CompressPdfResult> {
  // Lazy import keeps this module loadable in pure pdf-lib unit tests that
  // never touch the network path.
  const { apiFetchRaw } = await import("@mrnote/core/web/lib/api.ts");
  const res = await apiFetchRaw(
    "/api/v1/ai/notebook/pdf-compress",
    {
      method: "POST",
      body: JSON.stringify({ data_item_id: dataItemId, preset }),
    },
    { contentType: "application/json" },
  );
  const bytes = new Uint8Array(await res.arrayBuffer());
  const compressedSize = bytes.byteLength;
  // Trust the payload over the headers if a proxy strips them: fall back to
  // "no saving" (originalSize = compressedSize) rather than a fabricated one.
  const originalHeader = Number(res.headers.get("X-Original-Size"));
  const originalSize =
    Number.isFinite(originalHeader) && originalHeader >= compressedSize
      ? originalHeader
      : compressedSize;
  const savedBytes = originalSize - compressedSize;
  return {
    bytes,
    originalSize,
    compressedSize,
    savedBytes,
    savedRatio: originalSize > 0 ? savedBytes / originalSize : 0,
  };
}

const PDF_COMPRESS_PLAN_MAX_BYTES = 64 * 1024;
const PDF_COMPRESS_OUTPUT_MAX_BYTES = PDF_EDITABLE_MAX_BYTES;
const PDF_COMPRESS_MAX_PAGES = 1_000;
const PDF_COMPRESS_MAX_OBJECTS = 1_000_000;
const PDF_COMPRESS_MAX_IMAGES = 128;
const PDF_COMPRESS_MAX_IMAGE_PIXELS = 64_000_000;
const PDF_COMPRESS_MAX_AGGREGATE_PIXELS = 128_000_000;
const SHA256_RE = /^[a-f0-9]{64}$/;
const OBJECT_REF_RE = /^[1-9][0-9]*:[0-9]+$/;

const PDF_COMPRESS_PRESET_CONTRACT = Object.freeze({
  screen: Object.freeze({ targetDpi: 72, jpegQuality: 60 }),
  ebook: Object.freeze({ targetDpi: 150, jpegQuality: 75 }),
  print: Object.freeze({ targetDpi: 300, jpegQuality: 85 }),
}) satisfies Readonly<
  Record<
    PdfCompressServerPreset,
    Readonly<{ targetDpi: number; jpegQuality: number }>
  >
>;

export type PdfCompressionEvidenceSource =
  | Readonly<{
      kind: "data_item";
      datasetId: string;
      dataItemId: string;
      expectedSha256: string;
    }>
  | Readonly<{
      kind: "resource_version";
      resourceId: string;
      resourceVersionId: string;
      expectedPageId?: string | null;
      expectedSha256: string;
    }>;

export interface PdfCompressionLockedSourceIdentity {
  readonly kind: "data_item" | "resource_version";
  readonly workspaceId: string;
  readonly datasetId: string | null;
  readonly dataItemId: string | null;
  readonly resourceId: string | null;
  readonly resourceVersionId: string | null;
  readonly versionKind: string | null;
  readonly expectedPageId: string | null;
  readonly storageCustodyDigest: string;
  readonly sha256: string;
  readonly size: number;
}

export interface PdfLossyCompressionEvidencePlan {
  readonly sourceRequest: PdfCompressionEvidenceSource;
  readonly source: PdfCompressionLockedSourceIdentity;
  readonly preset: PdfCompressServerPreset;
  readonly targetDpi: 72 | 150 | 300;
  readonly jpegQuality: 60 | 75 | 85;
  readonly pageCount: number;
  readonly objectCount: number;
  readonly imageCount: number;
  readonly transformedObjectCount: number;
  readonly pageGeometryDigest: string;
  readonly nonTargetDigest: string;
  readonly sourceTrailerDigest: string;
  readonly nativePlanDigest: string;
  readonly evidenceDigest: string;
}

export interface PdfLossyCompressionEvidenceResult extends CompressPdfResult {
  readonly outputSha256: string;
  readonly outputTargetDigest: string;
  readonly outputTrailerDigest: string;
  readonly imagesRecompressed: number;
}

function exactRecord(
  value: unknown,
  keys: readonly string[],
  label: string,
): Record<string, unknown> {
  if (
    value === null ||
    typeof value !== "object" ||
    Array.isArray(value) ||
    Object.getPrototypeOf(value) !== Object.prototype
  ) {
    throw new Error(`Invalid ${label}`);
  }
  const actual = Object.keys(value).sort();
  const expected = [...keys].sort();
  if (
    actual.length !== expected.length ||
    actual.some((key, index) => key !== expected[index])
  ) {
    throw new Error(`Invalid ${label} schema`);
  }
  return value as Record<string, unknown>;
}

function exactString(value: unknown, label: string, maxLength = 512): string {
  if (
    typeof value !== "string" ||
    value.length < 1 ||
    value.length > maxLength ||
    value !== value.trim() ||
    value.includes("\0")
  ) {
    throw new Error(`Invalid ${label}`);
  }
  return value;
}

function exactNullableString(
  value: unknown,
  label: string,
  maxLength = 512,
): string | null {
  return value === null ? null : exactString(value, label, maxLength);
}

function exactSha256(value: unknown, label: string): string {
  const digest = exactString(value, label, 64).toLowerCase();
  if (!SHA256_RE.test(digest)) throw new Error(`Invalid ${label}`);
  return digest;
}

function exactInteger(
  value: unknown,
  label: string,
  minimum: number,
  maximum: number,
): number {
  if (
    typeof value !== "number" ||
    !Number.isSafeInteger(value) ||
    value < minimum ||
    value > maximum
  ) {
    throw new Error(`Invalid ${label}`);
  }
  return value;
}

function exactPositiveDecimal(value: unknown, label: string): void {
  if (typeof value !== "string" || value.length > 64) {
    throw new Error(`Invalid ${label}`);
  }
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    throw new Error(`Invalid ${label}`);
  }
}

function canonicalJson(value: unknown): string {
  if (value === null) return "null";
  if (typeof value === "string" || typeof value === "boolean") {
    return JSON.stringify(value);
  }
  if (typeof value === "number") {
    if (!Number.isSafeInteger(value)) {
      throw new Error("Invalid compression evidence number");
    }
    return JSON.stringify(value);
  }
  if (Array.isArray(value)) {
    return `[${value.map((entry) => canonicalJson(entry)).join(",")}]`;
  }
  if (
    typeof value !== "object" ||
    Object.getPrototypeOf(value) !== Object.prototype
  ) {
    throw new Error("Invalid compression evidence value");
  }
  return `{${Object.keys(value)
    .sort()
    .map(
      (key) =>
        `${JSON.stringify(key)}:${canonicalJson((value as Record<string, unknown>)[key])}`,
    )
    .join(",")}}`;
}

async function sha256Hex(value: Uint8Array | string): Promise<string> {
  const bytes =
    typeof value === "string" ? new TextEncoder().encode(value) : value;
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  try {
    const digest = await crypto.subtle.digest("SHA-256", copy.buffer);
    return Array.from(new Uint8Array(digest), (byte) =>
      byte.toString(16).padStart(2, "0"),
    ).join("");
  } finally {
    copy.fill(0);
  }
}

function normalizeStreamChunk(value: unknown): Uint8Array {
  if (!ArrayBuffer.isView(value)) {
    throw new Error("Invalid compression response chunk");
  }
  const view = value as ArrayBufferView;
  return new Uint8Array(view.buffer, view.byteOffset, view.byteLength);
}

async function readBoundedResponseBytes(
  response: Response,
  maximumBytes: number,
  label: string,
): Promise<Uint8Array> {
  const lengthHeader = response.headers.get("content-length");
  let declaredLength: number | null = null;
  if (lengthHeader !== null) {
    declaredLength = Number(lengthHeader);
    if (
      !Number.isSafeInteger(declaredLength) ||
      declaredLength < 1 ||
      declaredLength > maximumBytes
    ) {
      throw new Error(`${label} exceeds byte limit`);
    }
  }
  const reader = response.body?.getReader();
  if (!reader) throw new Error(`${label} body is not stream-readable`);
  const chunks: Uint8Array[] = [];
  let total = 0;
  try {
    while (true) {
      const result = await reader.read();
      if (result.done) break;
      const view = normalizeStreamChunk(result.value);
      if (view.byteLength === 0) continue;
      if (view.byteLength > maximumBytes - total) {
        await reader.cancel(`${label} exceeds byte limit`);
        throw new Error(`${label} exceeds byte limit`);
      }
      const copy = new Uint8Array(view.byteLength);
      copy.set(view);
      chunks.push(copy);
      total += copy.byteLength;
    }
    if (total < 1 || (declaredLength !== null && declaredLength !== total)) {
      throw new Error(`${label} length mismatch`);
    }
    const output = new Uint8Array(total);
    let offset = 0;
    for (const chunk of chunks) {
      output.set(chunk, offset);
      offset += chunk.byteLength;
      chunk.fill(0);
    }
    return output;
  } catch (error) {
    for (const chunk of chunks) chunk.fill(0);
    throw error;
  } finally {
    reader.releaseLock();
  }
}

function sourceRequestWire(source: PdfCompressionEvidenceSource) {
  const expectedSha256 = exactSha256(
    source.expectedSha256,
    "compression source SHA-256",
  );
  if (source.kind === "data_item") {
    return Object.freeze({
      kind: "data_item" as const,
      expected_sha256: expectedSha256,
      dataset_id: exactString(source.datasetId, "compression dataset ID", 64),
      data_item_id: exactString(
        source.dataItemId,
        "compression data item ID",
        64,
      ),
    });
  }
  return Object.freeze({
    kind: "resource_version" as const,
    expected_sha256: expectedSha256,
    resource_id: exactString(source.resourceId, "compression resource ID", 64),
    resource_version_id: exactString(
      source.resourceVersionId,
      "compression resource version ID",
      64,
    ),
    expected_page_id:
      source.expectedPageId == null
        ? null
        : exactString(source.expectedPageId, "compression page ID", 64),
  });
}

function normalizeSourceRequest(
  source: PdfCompressionEvidenceSource,
): PdfCompressionEvidenceSource {
  const expectedSha256 = exactSha256(
    source.expectedSha256,
    "compression source SHA-256",
  );
  if (source.kind === "data_item") {
    return Object.freeze({
      kind: "data_item",
      datasetId: exactString(source.datasetId, "compression dataset ID", 64),
      dataItemId: exactString(
        source.dataItemId,
        "compression data item ID",
        64,
      ),
      expectedSha256,
    });
  }
  return Object.freeze({
    kind: "resource_version",
    resourceId: exactString(source.resourceId, "compression resource ID", 64),
    resourceVersionId: exactString(
      source.resourceVersionId,
      "compression resource version ID",
      64,
    ),
    expectedPageId:
      source.expectedPageId == null
        ? null
        : exactString(source.expectedPageId, "compression page ID", 64),
    expectedSha256,
  });
}

function normalizeLockedSource(
  value: unknown,
  expected: PdfCompressionEvidenceSource,
): PdfCompressionLockedSourceIdentity {
  const dataItemKeys = [
    "kind",
    "workspaceId",
    "datasetId",
    "dataItemId",
    "storageCustodyDigest",
    "sha256",
    "size",
  ] as const;
  const resourceKeys = [
    "kind",
    "workspaceId",
    "resourceId",
    "resourceVersionId",
    "versionKind",
    "expectedPageId",
    "storageCustodyDigest",
    "sha256",
    "size",
  ] as const;
  const initial = exactRecord(
    value,
    expected.kind === "data_item" ? dataItemKeys : resourceKeys,
    "compression locked source",
  );
  if (initial.kind !== expected.kind) {
    throw new Error("Compression source kind mismatch");
  }
  const common = {
    kind: expected.kind,
    workspaceId: exactString(
      initial.workspaceId,
      "compression workspace ID",
      64,
    ),
    storageCustodyDigest: exactSha256(
      initial.storageCustodyDigest,
      "compression storage custody digest",
    ),
    sha256: exactSha256(initial.sha256, "compression locked SHA-256"),
    size: exactInteger(
      initial.size,
      "compression source size",
      1,
      PDF_COMPRESS_OUTPUT_MAX_BYTES,
    ),
  } as const;
  if (common.sha256 !== expected.expectedSha256) {
    throw new Error("Compression source SHA-256 mismatch");
  }
  if (expected.kind === "data_item") {
    const datasetId = exactString(
      initial.datasetId,
      "compression dataset ID",
      64,
    );
    const dataItemId = exactString(
      initial.dataItemId,
      "compression data item ID",
      64,
    );
    if (
      datasetId !== expected.datasetId ||
      dataItemId !== expected.dataItemId
    ) {
      throw new Error("Compression data item lineage mismatch");
    }
    return Object.freeze({
      ...common,
      kind: "data_item",
      datasetId,
      dataItemId,
      resourceId: null,
      resourceVersionId: null,
      versionKind: null,
      expectedPageId: null,
    });
  }
  const resourceId = exactString(
    initial.resourceId,
    "compression resource ID",
    64,
  );
  const resourceVersionId = exactString(
    initial.resourceVersionId,
    "compression resource version ID",
    64,
  );
  const expectedPageId = exactNullableString(
    initial.expectedPageId,
    "compression expected page ID",
    64,
  );
  if (
    resourceId !== expected.resourceId ||
    resourceVersionId !== expected.resourceVersionId ||
    expectedPageId !== (expected.expectedPageId ?? null)
  ) {
    throw new Error("Compression resource version lineage mismatch");
  }
  return Object.freeze({
    ...common,
    kind: "resource_version",
    datasetId: null,
    dataItemId: null,
    resourceId,
    resourceVersionId,
    versionKind: exactString(
      initial.versionKind,
      "compression resource version kind",
      64,
    ),
    expectedPageId,
  });
}

function lockedSourceWire(source: PdfCompressionLockedSourceIdentity) {
  return source.kind === "data_item"
    ? {
        kind: source.kind,
        workspaceId: source.workspaceId,
        datasetId: source.datasetId,
        dataItemId: source.dataItemId,
        storageCustodyDigest: source.storageCustodyDigest,
        sha256: source.sha256,
        size: source.size,
      }
    : {
        kind: source.kind,
        workspaceId: source.workspaceId,
        resourceId: source.resourceId,
        resourceVersionId: source.resourceVersionId,
        versionKind: source.versionKind,
        expectedPageId: source.expectedPageId,
        storageCustodyDigest: source.storageCustodyDigest,
        sha256: source.sha256,
        size: source.size,
      };
}

function validateTransformedRecord(value: unknown, sourceSize: number): void {
  const record = exactRecord(
    value,
    [
      "ref",
      "before",
      "placements",
      "placementsDigest",
      "effectiveDpi",
      "action",
      "after",
    ],
    "compression transformed image",
  );
  const ref = exactString(record.ref, "compression image ref", 32);
  if (!OBJECT_REF_RE.test(ref))
    throw new Error("Invalid compression image ref");
  if (record.action !== "recompress") {
    throw new Error("Invalid compression image action");
  }
  exactInteger(
    record.placements,
    "compression image placement count",
    1,
    PDF_COMPRESS_MAX_OBJECTS,
  );
  exactSha256(record.placementsDigest, "compression placements digest");
  exactPositiveDecimal(record.effectiveDpi, "compression effective DPI");
  const before = exactRecord(
    record.before,
    ["width", "height", "color", "bits", "filters", "size", "sha256"],
    "compression source image",
  );
  const beforeWidth = exactInteger(
    before.width,
    "compression source image width",
    1,
    PDF_COMPRESS_MAX_IMAGE_PIXELS,
  );
  const beforeHeight = exactInteger(
    before.height,
    "compression source image height",
    1,
    PDF_COMPRESS_MAX_IMAGE_PIXELS,
  );
  if (beforeWidth * beforeHeight > PDF_COMPRESS_MAX_IMAGE_PIXELS) {
    throw new Error("Compression source image exceeds pixel limit");
  }
  if (before.color !== "/DeviceRGB" && before.color !== "/DeviceGray") {
    throw new Error("Invalid compression source image color");
  }
  if (before.bits !== 8 || !Array.isArray(before.filters)) {
    throw new Error("Invalid compression source image encoding");
  }
  exactInteger(before.size, "compression source image size", 1, sourceSize);
  exactSha256(before.sha256, "compression source image SHA-256");
  const after = exactRecord(
    record.after,
    ["width", "height", "color", "bits", "filter", "size", "sha256"],
    "compression output image",
  );
  exactInteger(after.width, "compression output image width", 1, beforeWidth);
  exactInteger(
    after.height,
    "compression output image height",
    1,
    beforeHeight,
  );
  if (
    (after.color !== "/DeviceRGB" && after.color !== "/DeviceGray") ||
    after.bits !== 8 ||
    after.filter !== "/DCTDecode"
  ) {
    throw new Error("Invalid compression output image encoding");
  }
  const beforeSize = before.size as number;
  exactInteger(after.size, "compression output image size", 1, beforeSize - 1);
  exactSha256(after.sha256, "compression output image SHA-256");
}

async function normalizeNativePlan(
  value: unknown,
  source: PdfCompressionLockedSourceIdentity,
  preset: PdfCompressServerPreset,
): Promise<{
  pageCount: number;
  objectCount: number;
  imageCount: number;
  transformedObjectCount: number;
  pageGeometryDigest: string;
  nonTargetDigest: string;
  sourceTrailerDigest: string;
  nativePlanDigest: string;
}> {
  const plan = exactRecord(
    value,
    [
      "version",
      "engine",
      "sourceSha256",
      "sourceSize",
      "preset",
      "targetDpi",
      "jpegQuality",
      "pageCount",
      "objectCount",
      "sourceObjectRefs",
      "sourceObjectRefsDigest",
      "pageGeometryDigest",
      "imageCount",
      "aggregateImagePixels",
      "imageGraphDigest",
      "nonTargetDigest",
      "sourceTrailer",
      "trailerPolicy",
      "skipped",
      "transformed",
      "planDigest",
    ],
    "compression native plan",
  );
  if (
    plan.version !== 1 ||
    plan.sourceSha256 !== source.sha256 ||
    plan.sourceSize !== source.size ||
    plan.preset !== preset ||
    plan.skipped !== 0
  ) {
    throw new Error("Compression native plan identity mismatch");
  }
  const engine = exactRecord(
    plan.engine,
    ["pikepdf", "qpdf", "pillow"],
    "compression engine identity",
  );
  if (
    engine.pikepdf !== "9.4.0" ||
    engine.qpdf !== "11.9.1" ||
    engine.pillow !== "12.3.0"
  ) {
    throw new Error("Unexpected compression engine identity");
  }
  const presetContract = PDF_COMPRESS_PRESET_CONTRACT[preset];
  if (
    plan.targetDpi !== presetContract.targetDpi ||
    plan.jpegQuality !== presetContract.jpegQuality
  ) {
    throw new Error("Compression preset semantics mismatch");
  }
  const pageCount = exactInteger(
    plan.pageCount,
    "compression page count",
    1,
    PDF_COMPRESS_MAX_PAGES,
  );
  const objectCount = exactInteger(
    plan.objectCount,
    "compression object count",
    1,
    PDF_COMPRESS_MAX_OBJECTS,
  );
  // A document with no picture, or none the preset may shrink, is a plan
  // too: it tells the reader there is nothing to shrink (user audit PDF
  // 0029: a text-only file failed with 「无法压缩该 PDF」).
  const imageCount = exactInteger(
    plan.imageCount,
    "compression image count",
    0,
    PDF_COMPRESS_MAX_IMAGES,
  );
  exactInteger(
    plan.aggregateImagePixels,
    "compression aggregate pixels",
    0,
    PDF_COMPRESS_MAX_AGGREGATE_PIXELS,
  );
  if (
    !Array.isArray(plan.sourceObjectRefs) ||
    plan.sourceObjectRefs.length !== objectCount
  ) {
    throw new Error("Invalid compression source object set");
  }
  const refs = plan.sourceObjectRefs.map((ref) =>
    exactString(ref, "compression source object ref", 32),
  );
  if (
    refs.some((ref) => !OBJECT_REF_RE.test(ref)) ||
    refs.some((ref, index) => index > 0 && ref <= refs[index - 1])
  ) {
    throw new Error("Invalid compression source object set");
  }
  exactSha256(plan.sourceObjectRefsDigest, "compression source ref digest");
  const transformed = plan.transformed;
  if (
    !Array.isArray(transformed) ||
    transformed.length > imageCount
  ) {
    throw new Error("Invalid compression transformed image count");
  }
  transformed.forEach((record) =>
    validateTransformedRecord(record, source.size),
  );
  const trailer = exactRecord(
    plan.sourceTrailer,
    [
      "keys",
      "size",
      "maxObjectNumber",
      "objectCount",
      "rootDigest",
      "infoDigest",
      "ids",
      "digest",
    ],
    "compression source trailer",
  );
  if (
    trailer.objectCount !== objectCount ||
    !Array.isArray(trailer.keys) ||
    !Array.isArray(trailer.ids)
  ) {
    throw new Error("Invalid compression source trailer");
  }
  exactInteger(
    trailer.maxObjectNumber,
    "compression maximum object number",
    1,
    PDF_COMPRESS_MAX_OBJECTS,
  );
  exactInteger(
    trailer.size,
    "compression trailer size",
    2,
    PDF_COMPRESS_MAX_OBJECTS + 1,
  );
  const sourceTrailerDigest = exactSha256(
    trailer.digest,
    "compression source trailer digest",
  );
  const pageGeometryDigest = exactSha256(
    plan.pageGeometryDigest,
    "compression page geometry digest",
  );
  const nonTargetDigest = exactSha256(
    plan.nonTargetDigest,
    "compression non-target digest",
  );
  exactSha256(plan.imageGraphDigest, "compression image graph digest");
  const nativePlanDigest = exactSha256(
    plan.planDigest,
    "compression native plan digest",
  );
  const core = { ...plan };
  delete core.planDigest;
  if ((await sha256Hex(canonicalJson(core))) !== nativePlanDigest) {
    throw new Error("Compression native plan digest mismatch");
  }
  return {
    pageCount,
    objectCount,
    imageCount,
    transformedObjectCount: transformed.length,
    pageGeometryDigest,
    nonTargetDigest,
    sourceTrailerDigest,
    nativePlanDigest,
  };
}

export async function requestPdfLossyCompressionEvidencePlan(
  sourceValue: PdfCompressionEvidenceSource,
  preset: PdfCompressServerPreset,
): Promise<PdfLossyCompressionEvidencePlan> {
  if (!(preset in PDF_COMPRESS_PRESET_CONTRACT)) {
    throw new Error("Unsupported lossy PDF compression preset");
  }
  const sourceRequest = normalizeSourceRequest(sourceValue);
  const { apiFetchRaw } = await import("@mrnote/core/web/lib/api.ts");
  const response = await apiFetchRaw(
    "/api/v1/ai/notebook/pdf-compress/plan",
    {
      method: "POST",
      body: JSON.stringify({
        source: sourceRequestWire(sourceRequest),
        preset,
      }),
    },
    { contentType: "application/json" },
  );
  const contentType = response.headers.get("content-type") ?? "";
  if (!contentType.toLowerCase().startsWith("application/json")) {
    throw new Error("Invalid compression plan content type");
  }
  const bytes = await readBoundedResponseBytes(
    response,
    PDF_COMPRESS_PLAN_MAX_BYTES,
    "Compression plan response",
  );
  try {
    const text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    const payload = exactRecord(
      JSON.parse(text),
      ["source", "plan", "evidence"],
      "compression evidence response",
    );
    const source = normalizeLockedSource(payload.source, sourceRequest);
    const native = await normalizeNativePlan(payload.plan, source, preset);
    const evidence = exactRecord(
      payload.evidence,
      ["version", "source", "preset", "nativePlanDigest", "evidenceDigest"],
      "compression evidence envelope",
    );
    if (
      evidence.version !== 1 ||
      evidence.preset !== preset ||
      evidence.nativePlanDigest !== native.nativePlanDigest ||
      canonicalJson(evidence.source) !== canonicalJson(lockedSourceWire(source))
    ) {
      throw new Error("Compression evidence envelope mismatch");
    }
    const evidenceDigest = exactSha256(
      evidence.evidenceDigest,
      "compression evidence digest",
    );
    const evidenceCore = {
      version: 1,
      source: lockedSourceWire(source),
      preset,
      nativePlanDigest: native.nativePlanDigest,
    };
    if ((await sha256Hex(canonicalJson(evidenceCore))) !== evidenceDigest) {
      throw new Error("Compression evidence digest mismatch");
    }
    const presetContract = PDF_COMPRESS_PRESET_CONTRACT[preset];
    return Object.freeze({
      sourceRequest,
      source,
      preset,
      targetDpi: presetContract.targetDpi,
      jpegQuality: presetContract.jpegQuality,
      pageCount: native.pageCount,
      objectCount: native.objectCount,
      imageCount: native.imageCount,
      transformedObjectCount: native.transformedObjectCount,
      pageGeometryDigest: native.pageGeometryDigest,
      nonTargetDigest: native.nonTargetDigest,
      sourceTrailerDigest: native.sourceTrailerDigest,
      nativePlanDigest: native.nativePlanDigest,
      evidenceDigest,
    });
  } finally {
    bytes.fill(0);
  }
}

function decodeBase64UrlJson(value: string): unknown {
  if (value.length < 1 || value.length > PDF_COMPRESS_PLAN_MAX_BYTES * 2) {
    throw new Error("Invalid compression result proof header");
  }
  const normalized = value.replace(/-/g, "+").replace(/_/g, "/");
  if (!/^[A-Za-z0-9+/]*={0,2}$/.test(normalized)) {
    throw new Error("Invalid compression result proof encoding");
  }
  const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=");
  const binary = globalThis.atob(padded);
  const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
  try {
    return JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(bytes));
  } finally {
    bytes.fill(0);
  }
}

function validateOutputTrailerIds(value: unknown): void {
  if (
    !Array.isArray(value) ||
    value.length !== 2 ||
    value.some(
      (entry) =>
        typeof entry !== "string" ||
        entry.length < 2 ||
        entry.length > 128 ||
        entry.length % 2 !== 0 ||
        !/^[a-f0-9]+$/.test(entry),
    )
  ) {
    throw new Error("Invalid compression output trailer IDs");
  }
}

export async function executePdfLossyCompressionEvidencePlan(
  plan: PdfLossyCompressionEvidencePlan,
  isCurrentAuthority: () => boolean = () => true,
): Promise<PdfLossyCompressionEvidenceResult> {
  const executePath = "/api/v1/ai/notebook/pdf-compress/execute";
  const { apiFetchRaw, buildStreamPostHeaders } = await import("@mrnote/core/web/lib/api.ts");
  const headers = await buildStreamPostHeaders(executePath);
  // Module loading and authenticated-header/CSRF acquisition are both async
  // boundaries. Supplying those completed headers with requireCsrf=false makes
  // apiFetchRaw initiate fetch synchronously at the call below, so no hidden
  // await remains between this live predicate and the execute writer request.
  if (!isCurrentAuthority()) {
    throw new Error("Compression execute authority became stale");
  }
  const responsePromise = apiFetchRaw(
    executePath,
    {
      method: "POST",
      headers,
      body: JSON.stringify({
        source: sourceRequestWire(plan.sourceRequest),
        preset: plan.preset,
        expected_plan_digest: plan.nativePlanDigest,
        expected_evidence_digest: plan.evidenceDigest,
      }),
    },
    { requireCsrf: false, contentType: "application/json" },
  );
  const response = await responsePromise;
  const contentType = response.headers.get("content-type") ?? "";
  if (!contentType.toLowerCase().startsWith("application/pdf")) {
    throw new Error("Invalid compression output content type");
  }
  const headerPlanDigest = exactSha256(
    response.headers.get("X-PDF-Compress-Plan-Digest"),
    "compression result plan digest",
  );
  const headerEvidenceDigest = exactSha256(
    response.headers.get("X-PDF-Compress-Evidence-Digest"),
    "compression result evidence digest",
  );
  const headerOutputSha256 = exactSha256(
    response.headers.get("X-PDF-Compress-Output-SHA256"),
    "compression result SHA-256",
  );
  if (
    headerPlanDigest !== plan.nativePlanDigest ||
    headerEvidenceDigest !== plan.evidenceDigest
  ) {
    throw new Error("Compression result authority mismatch");
  }
  const proof = exactRecord(
    decodeBase64UrlJson(
      exactString(
        response.headers.get("X-PDF-Compress-Result-Proof"),
        "compression result proof",
        PDF_COMPRESS_PLAN_MAX_BYTES * 2,
      ),
    ),
    ["version", "source", "evidenceDigest", "native"],
    "compression result proof",
  );
  if (
    proof.version !== 1 ||
    proof.evidenceDigest !== plan.evidenceDigest ||
    canonicalJson(proof.source) !== canonicalJson(lockedSourceWire(plan.source))
  ) {
    throw new Error("Compression result lineage mismatch");
  }
  const native = exactRecord(
    proof.native,
    [
      "version",
      "sourceSha256",
      "sourceSize",
      "planDigest",
      "preset",
      "outputSha256",
      "outputSize",
      "imagesRecompressed",
      "pageCount",
      "pageGeometryDigest",
      "nonTargetDigest",
      "transformedObjectCount",
      "outputTargetDigest",
      "sourceTrailerDigest",
      "outputTrailerDigest",
      "outputTrailerIds",
    ],
    "compression native result proof",
  );
  const outputSize = exactInteger(
    native.outputSize,
    "compression output size",
    1,
    Math.min(plan.source.size, PDF_COMPRESS_OUTPUT_MAX_BYTES),
  );
  const imagesRecompressed = exactInteger(
    native.imagesRecompressed,
    "compression output image count",
    1,
    PDF_COMPRESS_MAX_IMAGES,
  );
  if (
    native.version !== 1 ||
    native.sourceSha256 !== plan.source.sha256 ||
    native.sourceSize !== plan.source.size ||
    native.planDigest !== plan.nativePlanDigest ||
    native.preset !== plan.preset ||
    native.outputSha256 !== headerOutputSha256 ||
    native.pageCount !== plan.pageCount ||
    native.pageGeometryDigest !== plan.pageGeometryDigest ||
    native.nonTargetDigest !== plan.nonTargetDigest ||
    native.sourceTrailerDigest !== plan.sourceTrailerDigest ||
    native.transformedObjectCount !== plan.transformedObjectCount ||
    imagesRecompressed !== plan.transformedObjectCount
  ) {
    throw new Error("Compression native result proof mismatch");
  }
  const outputTargetDigest = exactSha256(
    native.outputTargetDigest,
    "compression output target digest",
  );
  const outputTrailerDigest = exactSha256(
    native.outputTrailerDigest,
    "compression output trailer digest",
  );
  validateOutputTrailerIds(native.outputTrailerIds);
  const bytes = await readBoundedResponseBytes(
    response,
    Math.min(plan.source.size, PDF_COMPRESS_OUTPUT_MAX_BYTES),
    "Compression output",
  );
  if (bytes.byteLength !== outputSize) {
    bytes.fill(0);
    throw new Error("Compression output size mismatch");
  }
  const outputSha256 = await sha256Hex(bytes);
  if (outputSha256 !== headerOutputSha256) {
    bytes.fill(0);
    throw new Error("Compression output SHA-256 mismatch");
  }
  const savedBytes = plan.source.size - bytes.byteLength;
  return Object.freeze({
    bytes,
    originalSize: plan.source.size,
    compressedSize: bytes.byteLength,
    savedBytes,
    savedRatio: savedBytes / plan.source.size,
    outputSha256,
    outputTargetDigest,
    outputTrailerDigest,
    imagesRecompressed,
  });
}
