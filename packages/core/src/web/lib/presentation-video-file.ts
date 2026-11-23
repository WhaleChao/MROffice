/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  PRESENTATION_MEDIA_MAX_BYTES,
  requireWithinDocumentTotal,
} from "./document-size-limits";
import {
  MAX_PRESENTATION_VIDEO_FILE_BYTES,
  MAX_PRESENTATION_MEDIA_TIME_MS,
  assetBackedPresentationVideoSource,
  defaultEmbedGeometry,
  presentationStoredMediaBytes,
} from "./presentation-embed";

export const MAX_PRESENTATION_VIDEO_ASSETS_PER_DECK = 16;
export const NATIVE_PRESENTATION_VIDEO_SCHEMA_VERSION =
  "pptx-native-video-v1" as const;
export const NATIVE_PRESENTATION_AUDIO_SCHEMA_VERSION =
  "pptx-native-audio-v1" as const;

const NATIVE_PRESENTATION_VIDEO_KEYS = [
  "schemaVersion",
  "disposition",
  "slideId",
  "slidePart",
  "slideFingerprint",
  "relationshipsPart",
  "relationshipsFingerprint",
  "shapeId",
  "shapeFingerprint",
  "videoRelationshipId",
  "videoRelationshipType",
  "mediaRelationshipId",
  "mediaRelationshipType",
  "mediaPart",
  "mediaFingerprint",
  "mediaContentType",
  "mediaSizeBytes",
  "posterRelationshipId",
  "posterRelationshipType",
  "posterPart",
  "posterFingerprint",
] as const;
const NATIVE_PRESENTATION_VIDEO_OPTIONAL_KEYS = ["mediaDurationMs"] as const;

const NATIVE_PRESENTATION_VIDEO_RELATIONSHIP_TYPES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/video",
]);
const NATIVE_PRESENTATION_AUDIO_RELATIONSHIP_TYPES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/audio",
  ...NATIVE_PRESENTATION_VIDEO_RELATIONSHIP_TYPES,
]);
const NATIVE_PRESENTATION_MEDIA_RELATIONSHIP_TYPE =
  "http://schemas.microsoft.com/office/2007/relationships/media";
const NATIVE_PRESENTATION_IMAGE_RELATIONSHIP_TYPES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/image",
]);
const SHA256_HEX = /^[0-9a-f]{64}$/;
const RELATIONSHIP_ID = /^rId[1-9][0-9]*$/;
const SLIDE_ID = /^sld[1-9][0-9]*$/;
const SHAPE_ID = /^shp[1-9][0-9]*$/;
const SLIDE_PART = /^ppt\/slides\/slide[1-9][0-9]*\.xml$/;
const MEDIA_PART = /^ppt\/media\/[A-Za-z0-9][A-Za-z0-9._-]{0,199}$/;

export interface NativePresentationVideoAuthority {
  readonly schemaVersion: typeof NATIVE_PRESENTATION_VIDEO_SCHEMA_VERSION;
  readonly disposition: "native";
  readonly slideId: string;
  readonly slidePart: string;
  readonly slideFingerprint: string;
  readonly relationshipsPart: string;
  readonly relationshipsFingerprint: string;
  readonly shapeId: string;
  readonly shapeFingerprint: string;
  readonly videoRelationshipId: string;
  readonly videoRelationshipType: string;
  readonly mediaRelationshipId: string;
  readonly mediaRelationshipType: string;
  readonly mediaPart: string;
  readonly mediaFingerprint: string;
  readonly mediaContentType: "video/mp4";
  readonly mediaSizeBytes: number;
  readonly mediaDurationMs?: number;
  readonly posterRelationshipId: string;
  readonly posterRelationshipType: string;
  readonly posterPart: string;
  readonly posterFingerprint: string;
}

const NATIVE_PRESENTATION_AUDIO_KEYS = [
  "schemaVersion",
  "disposition",
  "slideId",
  "slidePart",
  "slideFingerprint",
  "relationshipsPart",
  "relationshipsFingerprint",
  "shapeId",
  "shapeFingerprint",
  "audioRelationshipId",
  "audioRelationshipType",
  "mediaRelationshipId",
  "mediaRelationshipType",
  "mediaPart",
  "mediaFingerprint",
  "mediaContentType",
  "mediaSizeBytes",
  "posterRelationshipId",
  "posterRelationshipType",
  "posterPart",
  "posterFingerprint",
] as const;
const NATIVE_PRESENTATION_AUDIO_OPTIONAL_KEYS = ["mediaDurationMs"] as const;

export interface NativePresentationAudioAuthority {
  readonly schemaVersion: typeof NATIVE_PRESENTATION_AUDIO_SCHEMA_VERSION;
  readonly disposition: "native";
  readonly slideId: string;
  readonly slidePart: string;
  readonly slideFingerprint: string;
  readonly relationshipsPart: string;
  readonly relationshipsFingerprint: string;
  readonly shapeId: string;
  readonly shapeFingerprint: string;
  readonly audioRelationshipId: string;
  readonly audioRelationshipType: string;
  readonly mediaRelationshipId: string;
  readonly mediaRelationshipType: string;
  readonly mediaPart: string;
  readonly mediaFingerprint: string;
  readonly mediaContentType: string;
  readonly mediaSizeBytes: number;
  readonly mediaDurationMs?: number;
  readonly posterRelationshipId: string;
  readonly posterRelationshipType: string;
  readonly posterPart: string;
  readonly posterFingerprint: string;
}

export interface PresentationNativeVideoReplacementTarget {
  readonly slideId: string;
  readonly blockId: string;
  readonly authority: NativePresentationVideoAuthority;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export function parseNativePresentationVideoAuthority(
  value: unknown,
): NativePresentationVideoAuthority | null {
  if (!isRecord(value)) return null;
  const keys = Object.keys(value).sort();
  const requiredKeys = keys.filter(
    (key) =>
      !(NATIVE_PRESENTATION_VIDEO_OPTIONAL_KEYS as readonly string[]).includes(
        key,
      ),
  );
  if (
    requiredKeys.length !== NATIVE_PRESENTATION_VIDEO_KEYS.length ||
    requiredKeys.some(
      (key) =>
        !(NATIVE_PRESENTATION_VIDEO_KEYS as readonly string[]).includes(key),
    ) ||
    keys.length !== requiredKeys.length +
      ("mediaDurationMs" in value ? 1 : 0)
  ) {
    return null;
  }
  const receipt = value as unknown as NativePresentationVideoAuthority;
  const expectedRelationshipsPart = receipt.slidePart.replace(
    /^ppt\/slides\//,
    "ppt/slides/_rels/",
  ) + ".rels";
  if (
    receipt.schemaVersion !== NATIVE_PRESENTATION_VIDEO_SCHEMA_VERSION ||
    receipt.disposition !== "native" ||
    !SLIDE_ID.test(receipt.slideId) ||
    !SLIDE_PART.test(receipt.slidePart) ||
    receipt.relationshipsPart !== expectedRelationshipsPart ||
    !SHAPE_ID.test(receipt.shapeId) ||
    !SHA256_HEX.test(receipt.slideFingerprint) ||
    !SHA256_HEX.test(receipt.relationshipsFingerprint) ||
    !SHA256_HEX.test(receipt.shapeFingerprint) ||
    !SHA256_HEX.test(receipt.mediaFingerprint) ||
    !SHA256_HEX.test(receipt.posterFingerprint) ||
    !RELATIONSHIP_ID.test(receipt.videoRelationshipId) ||
    !RELATIONSHIP_ID.test(receipt.mediaRelationshipId) ||
    !RELATIONSHIP_ID.test(receipt.posterRelationshipId) ||
    new Set([
      receipt.videoRelationshipId,
      receipt.mediaRelationshipId,
      receipt.posterRelationshipId,
    ]).size !== 3 ||
    !NATIVE_PRESENTATION_VIDEO_RELATIONSHIP_TYPES.has(
      receipt.videoRelationshipType,
    ) ||
    receipt.mediaRelationshipType !==
      NATIVE_PRESENTATION_MEDIA_RELATIONSHIP_TYPE ||
    !NATIVE_PRESENTATION_IMAGE_RELATIONSHIP_TYPES.has(
      receipt.posterRelationshipType,
    ) ||
    !MEDIA_PART.test(receipt.mediaPart) ||
    !MEDIA_PART.test(receipt.posterPart) ||
    receipt.mediaPart === receipt.posterPart ||
    receipt.mediaContentType !== "video/mp4" ||
    !Number.isInteger(receipt.mediaSizeBytes) ||
    receipt.mediaSizeBytes <= 0 ||
    receipt.mediaSizeBytes > MAX_PRESENTATION_VIDEO_FILE_BYTES ||
    (receipt.mediaDurationMs !== undefined &&
      (!Number.isInteger(receipt.mediaDurationMs) ||
        receipt.mediaDurationMs <= 0 ||
        receipt.mediaDurationMs > MAX_PRESENTATION_MEDIA_TIME_MS))
  ) {
    return null;
  }
  return receipt;
}

export function nativePresentationVideoAuthority(
  block: Readonly<Record<string, unknown>>,
): NativePresentationVideoAuthority | null {
  if (
    block.type !== "embed" ||
    block.kind !== "video" ||
    block.url !== "" ||
    typeof block.id !== "string" ||
    !block.id
  ) {
    return null;
  }
  const source = isRecord(block.source) ? block.source : null;
  const receipt = parseNativePresentationVideoAuthority(
    source?.nativeVideoAuthority,
  );
  return receipt &&
    source?.slideId === receipt.slideId &&
    source?.shapeId === receipt.shapeId
    ? receipt
    : null;
}

export function parseNativePresentationAudioAuthority(
  value: unknown,
): NativePresentationAudioAuthority | null {
  if (!isRecord(value)) return null;
  const keys = Object.keys(value).sort();
  const requiredKeys = keys.filter(
    (key) =>
      !(NATIVE_PRESENTATION_AUDIO_OPTIONAL_KEYS as readonly string[]).includes(
        key,
      ),
  );
  if (
    requiredKeys.length !== NATIVE_PRESENTATION_AUDIO_KEYS.length ||
    requiredKeys.some(
      (key) =>
        !(NATIVE_PRESENTATION_AUDIO_KEYS as readonly string[]).includes(key),
    ) ||
    keys.length !== requiredKeys.length +
      ("mediaDurationMs" in value ? 1 : 0)
  ) {
    return null;
  }
  const receipt = value as unknown as NativePresentationAudioAuthority;
  const expectedRelationshipsPart = receipt.slidePart.replace(
    /^ppt\/slides\//,
    "ppt/slides/_rels/",
  ) + ".rels";
  if (
    receipt.schemaVersion !== NATIVE_PRESENTATION_AUDIO_SCHEMA_VERSION ||
    receipt.disposition !== "native" ||
    !SLIDE_ID.test(receipt.slideId) ||
    !SLIDE_PART.test(receipt.slidePart) ||
    receipt.relationshipsPart !== expectedRelationshipsPart ||
    !SHAPE_ID.test(receipt.shapeId) ||
    !SHA256_HEX.test(receipt.slideFingerprint) ||
    !SHA256_HEX.test(receipt.relationshipsFingerprint) ||
    !SHA256_HEX.test(receipt.shapeFingerprint) ||
    !SHA256_HEX.test(receipt.mediaFingerprint) ||
    !SHA256_HEX.test(receipt.posterFingerprint) ||
    !RELATIONSHIP_ID.test(receipt.audioRelationshipId) ||
    !RELATIONSHIP_ID.test(receipt.mediaRelationshipId) ||
    !RELATIONSHIP_ID.test(receipt.posterRelationshipId) ||
    new Set([
      receipt.audioRelationshipId,
      receipt.mediaRelationshipId,
      receipt.posterRelationshipId,
    ]).size !== 3 ||
    !NATIVE_PRESENTATION_AUDIO_RELATIONSHIP_TYPES.has(
      receipt.audioRelationshipType,
    ) ||
    receipt.mediaRelationshipType !==
      NATIVE_PRESENTATION_MEDIA_RELATIONSHIP_TYPE ||
    !NATIVE_PRESENTATION_IMAGE_RELATIONSHIP_TYPES.has(
      receipt.posterRelationshipType,
    ) ||
    !MEDIA_PART.test(receipt.mediaPart) ||
    !MEDIA_PART.test(receipt.posterPart) ||
    receipt.mediaPart === receipt.posterPart ||
    !/^audio\/[A-Za-z0-9.+-]{1,100}$/.test(receipt.mediaContentType) ||
    !Number.isInteger(receipt.mediaSizeBytes) ||
    receipt.mediaSizeBytes <= 0 ||
    receipt.mediaSizeBytes > MAX_PRESENTATION_VIDEO_FILE_BYTES ||
    (receipt.mediaDurationMs !== undefined &&
      (!Number.isInteger(receipt.mediaDurationMs) ||
        receipt.mediaDurationMs <= 0 ||
        receipt.mediaDurationMs > MAX_PRESENTATION_MEDIA_TIME_MS))
  ) {
    return null;
  }
  return receipt;
}

/**
 * A clip's length as its package states it (an imported video or audio
 * clip's `mediaDurationMs`), or null: the duration a Play media call states.
 */
export function presentationClipDurationMs(
  block: Readonly<Record<string, unknown>>,
): number | null {
  const authority =
    nativePresentationVideoAuthority(block) ?? nativePresentationAudioAuthority(block);
  const duration = authority?.mediaDurationMs;
  return typeof duration === "number" && Number.isInteger(duration) && duration > 0
    ? duration
    : null;
}

export function nativePresentationAudioAuthority(
  block: Readonly<Record<string, unknown>>,
): NativePresentationAudioAuthority | null {
  if (
    block.type !== "embed" ||
    block.kind !== "audio" ||
    block.url !== "" ||
    typeof block.id !== "string" ||
    !block.id
  ) {
    return null;
  }
  const source = isRecord(block.source) ? block.source : null;
  const receipt = parseNativePresentationAudioAuthority(
    source?.nativeAudioAuthority,
  );
  return receipt &&
    source?.slideId === receipt.slideId &&
    source?.shapeId === receipt.shapeId
    ? receipt
    : null;
}

function nativeReplacementTargetMatches(
  current: PresentationNativeVideoReplacementTarget | null,
  expected: PresentationNativeVideoReplacementTarget,
): boolean {
  if (
    current === null ||
    current.slideId !== expected.slideId ||
    current.blockId !== expected.blockId
  ) {
    return false;
  }
  const currentAuthority = current.authority as unknown as Record<
    string,
    unknown
  >;
  const expectedAuthority = expected.authority as unknown as Record<
    string,
    unknown
  >;
  return NATIVE_PRESENTATION_VIDEO_KEYS.every(
    (key) => currentAuthority[key] === expectedAuthority[key],
  );
}

export type PresentationVideoFileFailureCode =
  | "empty"
  | "unsupported-type"
  | "too-large"
  | "bad-magic";

export class PresentationVideoFileError extends Error {
  readonly code: PresentationVideoFileFailureCode;

  constructor(code: PresentationVideoFileFailureCode) {
    super(`presentation_video_file_${code}`);
    this.name = "PresentationVideoFileError";
    this.code = code;
  }
}

export interface PreparedPresentationVideoFile {
  readonly mimeType: "video/mp4";
  readonly sizeBytes: number;
  readonly title: string;
}

export type PresentationVideoInsertFailureCode =
  | "busy"
  | "no-target"
  | "native-deck"
  | "stale-target"
  | "slide-full"
  | "deck-budget"
  | "invalid-asset-id"
  | "identity-collision"
  | "commit-failed"
  | "compensation-failed";

export class PresentationVideoInsertError extends Error {
  readonly code: PresentationVideoInsertFailureCode;
  readonly cause?: unknown;

  constructor(code: PresentationVideoInsertFailureCode, cause?: unknown) {
    super(`presentation_video_insert_${code}`);
    this.name = "PresentationVideoInsertError";
    this.code = code;
    this.cause = cause;
  }
}

export interface PresentationVideoInsertSlide {
  readonly id: string;
  readonly blocks: readonly Record<string, unknown>[];
}

export interface PresentationVideoFileCommandAccessors<
  TSlide extends PresentationVideoInsertSlide,
> {
  getSlides: () => readonly TSlide[];
  getActiveIndex: () => number;
  getSlideBounds: (slide: TSlide) => { width: number; height: number };
  deckHasNativeSource: (slides: readonly TSlide[]) => boolean;
  getNativeReplacementTarget: () =>
    | PresentationNativeVideoReplacementTarget
    | null;
  openFilePicker: () => boolean;
  prepareFile: (file: File) => Promise<PreparedPresentationVideoFile>;
  uploadFile: (file: File) => Promise<string>;
  attestUploadedAsset: (
    assetId: string,
    prepared: PreparedPresentationVideoFile,
  ) => Promise<void>;
  deleteUploadedAsset: (assetId: string) => Promise<void>;
  commitSlides: (slides: TSlide[]) => void;
  selectInsertedVideo: (blockId: string, slideIndex: number) => void;
  nextId: () => string;
  maxBlocksPerSlide?: number;
}

function fail(code: PresentationVideoFileFailureCode): never {
  throw new PresentationVideoFileError(code);
}

const MAX_PRESENTATION_MP4_MOOV_BYTES = 8 * 1024 * 1024;
const MAX_PRESENTATION_MP4_FTYP_BYTES = 4 * 1024;
const MAX_PRESENTATION_MP4_BOXES = 4_096;
const MAX_PRESENTATION_MP4_SAMPLES = 1_000_000;
const MAX_PRESENTATION_MP4_CHUNKS = 262_144;
const MAX_PRESENTATION_MP4_SAMPLE_TABLE_ENTRIES = 65_536;
const MAX_PRESENTATION_MP4_FIRST_SAMPLE_BYTES = 16 * 1024 * 1024;
const PRESENTATION_MP4_BRANDS = new Set([
  "isom",
  "iso2",
  "mp41",
  "mp42",
  "avc1",
  "M4V ",
]);
const PRESENTATION_MP4_VIDEO_SAMPLE_ENTRIES = new Set(["avc1", "avc3"]);

interface PresentationMp4Box {
  readonly kind: string;
  readonly payloadStart: number;
  readonly end: number;
}

interface PresentationMp4Budget {
  remaining: number;
}

interface PresentationMp4VideoSamplePlan {
  readonly firstSampleStart: number;
  readonly firstSampleSize: number;
  readonly nalLengthSize: number;
}

function mp4Uint32(bytes: Uint8Array, offset: number): number | null {
  if (offset < 0 || offset + 4 > bytes.length) return null;
  return new DataView(
    bytes.buffer,
    bytes.byteOffset + offset,
    4,
  ).getUint32(0, false);
}

function mp4Kind(bytes: Uint8Array, offset: number): string | null {
  if (offset < 0 || offset + 4 > bytes.length) return null;
  return String.fromCharCode(
    bytes[offset],
    bytes[offset + 1],
    bytes[offset + 2],
    bytes[offset + 3],
  );
}

function parsePresentationMp4Boxes(
  bytes: Uint8Array,
  start: number,
  end: number,
  budget: PresentationMp4Budget,
): readonly PresentationMp4Box[] | null {
  const boxes: PresentationMp4Box[] = [];
  let cursor = start;
  while (cursor < end) {
    budget.remaining -= 1;
    if (budget.remaining < 0 || end - cursor < 8) return null;
    let size = mp4Uint32(bytes, cursor);
    const kind = mp4Kind(bytes, cursor + 4);
    if (size === null || kind === null) return null;
    let headerSize = 8;
    if (size === 1) {
      const high = mp4Uint32(bytes, cursor + 8);
      const low = mp4Uint32(bytes, cursor + 12);
      if (high === null || low === null || high > 0x1f_ffff) return null;
      size = high * 0x1_0000_0000 + low;
      headerSize = 16;
    } else if (size === 0) {
      size = end - cursor;
    }
    if (
      !Number.isSafeInteger(size) ||
      size < headerSize ||
      size > end - cursor
    ) {
      return null;
    }
    boxes.push({
      kind,
      payloadStart: cursor + headerSize,
      end: cursor + size,
    });
    cursor += size;
  }
  return cursor === end ? boxes : null;
}

async function readPresentationMp4Range(
  file: File,
  start: number,
  end: number,
): Promise<Uint8Array | null> {
  if (
    !Number.isSafeInteger(start) ||
    !Number.isSafeInteger(end) ||
    start < 0 ||
    end < start ||
    end > file.size
  ) {
    return null;
  }
  const blob = file.slice(start, end);
  const buffer =
    typeof blob.arrayBuffer === "function"
      ? await blob.arrayBuffer()
      : await new Promise<ArrayBuffer>((resolve, reject) => {
          const reader = new FileReader();
          reader.onerror = () => reject(reader.error ?? new Error("file-read"));
          reader.onload = () =>
            reader.result instanceof ArrayBuffer
              ? resolve(reader.result)
              : reject(new Error("file-read"));
          reader.readAsArrayBuffer(blob);
        });
  const bytes = new Uint8Array(buffer);
  return bytes.length === end - start ? bytes : null;
}

async function readPresentationMp4TopLevel(
  file: File,
  budget: PresentationMp4Budget,
): Promise<readonly PresentationMp4Box[] | null> {
  const boxes: PresentationMp4Box[] = [];
  let cursor = 0;
  while (cursor < file.size) {
    budget.remaining -= 1;
    if (budget.remaining < 0 || file.size - cursor < 8) return null;
    const header = await readPresentationMp4Range(
      file,
      cursor,
      Math.min(file.size, cursor + 16),
    );
    if (!header || header.length < 8) return null;
    let size = mp4Uint32(header, 0);
    const kind = mp4Kind(header, 4);
    if (size === null || kind === null) return null;
    let headerSize = 8;
    if (size === 1) {
      const high = mp4Uint32(header, 8);
      const low = mp4Uint32(header, 12);
      if (high === null || low === null || high > 0x1f_ffff) return null;
      size = high * 0x1_0000_0000 + low;
      headerSize = 16;
    } else if (size === 0) {
      size = file.size - cursor;
    }
    if (
      !Number.isSafeInteger(size) ||
      size < headerSize ||
      size > file.size - cursor
    ) {
      return null;
    }
    boxes.push({
      kind,
      payloadStart: cursor + headerSize,
      end: cursor + size,
    });
    cursor += size;
  }
  return cursor === file.size ? boxes : null;
}

function presentationAvcConfigNalLengthSize(bytes: Uint8Array): number | null {
  if (
    bytes.length < 8 ||
    bytes[0] !== 1 ||
    (bytes[4] & 0xfc) !== 0xfc ||
    (bytes[5] & 0xe0) !== 0xe0
  ) {
    return null;
  }
  const spsCount = bytes[5] & 0x1f;
  if (spsCount === 0) return null;
  let cursor = 6;
  for (let index = 0; index < spsCount; index += 1) {
    if (cursor + 2 > bytes.length) return null;
    const length = (bytes[cursor] << 8) | bytes[cursor + 1];
    cursor += 2;
    if (length === 0 || cursor + length > bytes.length) return null;
    cursor += length;
  }
  if (cursor >= bytes.length) return null;
  const ppsCount = bytes[cursor];
  cursor += 1;
  if (ppsCount === 0) return null;
  for (let index = 0; index < ppsCount; index += 1) {
    if (cursor + 2 > bytes.length) return null;
    const length = (bytes[cursor] << 8) | bytes[cursor + 1];
    cursor += 2;
    if (length === 0 || cursor + length > bytes.length) return null;
    cursor += length;
  }
  const lengthSizeMinusOne = bytes[4] & 0x03;
  return cursor <= bytes.length && lengthSizeMinusOne !== 2
    ? lengthSizeMinusOne + 1
    : null;
}

function presentationStsdSupportedVideo(
  bytes: Uint8Array,
  stsd: PresentationMp4Box,
  budget: PresentationMp4Budget,
): { descriptionIndex: number; nalLengthSize: number } | null {
  if (stsd.end - stsd.payloadStart < 8 || bytes[stsd.payloadStart] !== 0) {
    return null;
  }
  const count = mp4Uint32(bytes, stsd.payloadStart + 4);
  if (count === null || count === 0 || count > 64) return null;
  const entries = parsePresentationMp4Boxes(
    bytes,
    stsd.payloadStart + 8,
    stsd.end,
    budget,
  );
  if (!entries || entries.length !== count) return null;
  const supported: { descriptionIndex: number; nalLengthSize: number }[] = [];
  for (const [entryOffset, entry] of entries.entries()) {
    if (!PRESENTATION_MP4_VIDEO_SAMPLE_ENTRIES.has(entry.kind)) continue;
    if (entry.end - entry.payloadStart < 78) return null;
    const width =
      (bytes[entry.payloadStart + 24] << 8) |
      bytes[entry.payloadStart + 25];
    const height =
      (bytes[entry.payloadStart + 26] << 8) |
      bytes[entry.payloadStart + 27];
    if (width === 0 || height === 0) return null;
    const children = parsePresentationMp4Boxes(
      bytes,
      entry.payloadStart + 78,
      entry.end,
      budget,
    );
    if (!children) return null;
    const configs = children.filter((box) => box.kind === "avcC");
    if (configs.length !== 1) return null;
    const config = configs[0];
    const nalLengthSize = presentationAvcConfigNalLengthSize(
      bytes.subarray(config.payloadStart, config.end),
    );
    if (nalLengthSize === null) return null;
    supported.push({ descriptionIndex: entryOffset + 1, nalLengthSize });
  }
  return supported.length === 1 ? supported[0] : null;
}

function presentationFullBoxPayloadStart(
  bytes: Uint8Array,
  box: PresentationMp4Box,
): number | null {
  return box.end - box.payloadStart >= 4 &&
    bytes[box.payloadStart] === 0 &&
    bytes[box.payloadStart + 1] === 0 &&
    bytes[box.payloadStart + 2] === 0 &&
    bytes[box.payloadStart + 3] === 0
    ? box.payloadStart + 4
    : null;
}

function presentationSttsSampleCount(
  bytes: Uint8Array,
  box: PresentationMp4Box,
): number | null {
  const start = presentationFullBoxPayloadStart(bytes, box);
  if (start === null || box.end - start < 4) return null;
  const entryCount = mp4Uint32(bytes, start);
  if (
    entryCount === null ||
    entryCount < 1 ||
    entryCount > MAX_PRESENTATION_MP4_SAMPLE_TABLE_ENTRIES ||
    box.end - start !== 4 + entryCount * 8
  ) {
    return null;
  }
  let total = 0;
  for (let offset = start + 4; offset < box.end; offset += 8) {
    const count = mp4Uint32(bytes, offset);
    const delta = mp4Uint32(bytes, offset + 4);
    if (
      count === null ||
      delta === null ||
      count === 0 ||
      delta === 0 ||
      total > MAX_PRESENTATION_MP4_SAMPLES - count
    ) {
      return null;
    }
    total += count;
  }
  return total > 0 ? total : null;
}

function presentationSampleSizes(
  bytes: Uint8Array,
  box: PresentationMp4Box,
): readonly number[] | null {
  const start = presentationFullBoxPayloadStart(bytes, box);
  if (start === null || box.end - start < 8) return null;
  const defaultSize = mp4Uint32(bytes, start);
  const sampleCount = mp4Uint32(bytes, start + 4);
  if (
    defaultSize === null ||
    sampleCount === null ||
    sampleCount < 1 ||
    sampleCount > MAX_PRESENTATION_MP4_SAMPLES
  ) {
    return null;
  }
  if (defaultSize > 0) {
    return box.end - start === 8
      ? Array.from({ length: sampleCount }, () => defaultSize)
      : null;
  }
  if (box.end - start !== 8 + sampleCount * 4) return null;
  const sizes: number[] = [];
  for (let offset = start + 8; offset < box.end; offset += 4) {
    const size = mp4Uint32(bytes, offset);
    if (size === null || size === 0) return null;
    sizes.push(size);
  }
  return sizes.length === sampleCount ? sizes : null;
}

function presentationChunkOffsets(
  bytes: Uint8Array,
  box: PresentationMp4Box,
): readonly number[] | null {
  const start = presentationFullBoxPayloadStart(bytes, box);
  if (start === null || box.end - start < 4) return null;
  const count = mp4Uint32(bytes, start);
  const width = box.kind === "stco" ? 4 : 8;
  if (
    count === null ||
    count < 1 ||
    count > MAX_PRESENTATION_MP4_CHUNKS ||
    box.end - start !== 4 + count * width
  ) {
    return null;
  }
  const offsets: number[] = [];
  for (let offset = start + 4; offset < box.end; offset += width) {
    if (width === 4) {
      const value = mp4Uint32(bytes, offset);
      if (value === null) return null;
      offsets.push(value);
    } else {
      const high = mp4Uint32(bytes, offset);
      const low = mp4Uint32(bytes, offset + 4);
      if (high === null || low === null || high > 0x1f_ffff) return null;
      offsets.push(high * 0x1_0000_0000 + low);
    }
  }
  return offsets.length === count ? offsets : null;
}

function presentationSampleToChunk(
  bytes: Uint8Array,
  box: PresentationMp4Box,
): readonly (readonly [number, number, number])[] | null {
  const start = presentationFullBoxPayloadStart(bytes, box);
  if (start === null || box.end - start < 4) return null;
  const count = mp4Uint32(bytes, start);
  if (
    count === null ||
    count < 1 ||
    count > MAX_PRESENTATION_MP4_SAMPLE_TABLE_ENTRIES ||
    box.end - start !== 4 + count * 12
  ) {
    return null;
  }
  const entries: [number, number, number][] = [];
  for (let offset = start + 4; offset < box.end; offset += 12) {
    const firstChunk = mp4Uint32(bytes, offset);
    const samplesPerChunk = mp4Uint32(bytes, offset + 4);
    const descriptionIndex = mp4Uint32(bytes, offset + 8);
    if (
      firstChunk === null ||
      samplesPerChunk === null ||
      descriptionIndex === null ||
      firstChunk === 0 ||
      samplesPerChunk === 0 ||
      descriptionIndex === 0 ||
      (entries.length > 0 && entries[entries.length - 1][0] >= firstChunk)
    ) {
      return null;
    }
    entries.push([firstChunk, samplesPerChunk, descriptionIndex]);
  }
  return entries[0]?.[0] === 1 ? entries : null;
}

function presentationAvcSampleHasVcl(
  sample: Uint8Array,
  nalLengthSize: number,
): boolean {
  if (nalLengthSize < 1 || nalLengthSize > 4 || sample.length < nalLengthSize + 1) {
    return false;
  }
  let cursor = 0;
  let foundVcl = false;
  while (cursor < sample.length) {
    if (cursor + nalLengthSize > sample.length) return false;
    let nalSize = 0;
    for (let index = 0; index < nalLengthSize; index += 1) {
      nalSize = nalSize * 256 + sample[cursor + index];
    }
    cursor += nalLengthSize;
    if (nalSize === 0 || nalSize > sample.length - cursor) return false;
    const nalType = sample[cursor] & 0x1f;
    if (nalType >= 1 && nalType <= 5) foundVcl = true;
    cursor += nalSize;
  }
  return cursor === sample.length && foundVcl;
}

function presentationVideoSamplePlan(
  bytes: Uint8Array,
  children: readonly PresentationMp4Box[],
  supported: { descriptionIndex: number; nalLengthSize: number },
  mediaData: readonly PresentationMp4Box[],
): PresentationMp4VideoSamplePlan | null {
  const stts = children.filter((box) => box.kind === "stts");
  const stsc = children.filter((box) => box.kind === "stsc");
  const stsz = children.filter((box) => box.kind === "stsz");
  const offsets = children.filter(
    (box) => box.kind === "stco" || box.kind === "co64",
  );
  if (stts.length !== 1 || stsc.length !== 1 || stsz.length !== 1 || offsets.length !== 1) {
    return null;
  }
  const timingCount = presentationSttsSampleCount(bytes, stts[0]);
  const sizes = presentationSampleSizes(bytes, stsz[0]);
  const chunkOffsets = presentationChunkOffsets(bytes, offsets[0]);
  const chunkMap = presentationSampleToChunk(bytes, stsc[0]);
  if (
    timingCount === null ||
    !sizes ||
    !chunkOffsets ||
    !chunkMap ||
    timingCount !== sizes.length ||
    chunkMap[chunkMap.length - 1][0] > chunkOffsets.length ||
    chunkMap.some((entry) => entry[2] !== supported.descriptionIndex)
  ) {
    return null;
  }
  let sampleIndex = 0;
  let mapIndex = 0;
  let previousEnd = -1;
  let firstSampleStart = -1;
  let firstSampleSize = 0;
  let mediaIndex = 0;
  if (
    mediaData.some(
      (box, index) =>
        index > 0 && mediaData[index - 1].end > box.payloadStart,
    )
  ) {
    return null;
  }
  for (let chunkIndex = 0; chunkIndex < chunkOffsets.length; chunkIndex += 1) {
    const chunkNumber = chunkIndex + 1;
    while (
      mapIndex + 1 < chunkMap.length &&
      chunkMap[mapIndex + 1][0] <= chunkNumber
    ) {
      mapIndex += 1;
    }
    const samplesPerChunk = chunkMap[mapIndex][1];
    const chunkStart = chunkOffsets[chunkIndex];
    if (sampleIndex > sizes.length - samplesPerChunk || chunkStart < previousEnd) {
      return null;
    }
    let cursor = chunkStart;
    for (let index = 0; index < samplesPerChunk; index += 1) {
      const size = sizes[sampleIndex];
      if (!Number.isSafeInteger(cursor) || cursor > Number.MAX_SAFE_INTEGER - size) {
        return null;
      }
      const end = cursor + size;
      while (
        mediaIndex < mediaData.length &&
        cursor >= mediaData[mediaIndex].end
      ) {
        mediaIndex += 1;
      }
      if (
        mediaIndex >= mediaData.length ||
        cursor < mediaData[mediaIndex].payloadStart ||
        end > mediaData[mediaIndex].end
      ) {
        return null;
      }
      if (sampleIndex === 0) {
        firstSampleStart = cursor;
        firstSampleSize = size;
      }
      cursor = end;
      sampleIndex += 1;
    }
    previousEnd = cursor;
  }
  if (
    sampleIndex !== sizes.length ||
    firstSampleSize < 1 ||
    firstSampleSize > MAX_PRESENTATION_MP4_FIRST_SAMPLE_BYTES
  ) {
    return null;
  }
  return {
    firstSampleStart,
    firstSampleSize,
    nalLengthSize: supported.nalLengthSize,
  };
}

function presentationTrakVideoDisposition(
  bytes: Uint8Array,
  trak: PresentationMp4Box,
  budget: PresentationMp4Budget,
  mediaData: readonly PresentationMp4Box[],
): { readonly kind: "non-video" } | {
  readonly kind: "video";
  readonly plan: PresentationMp4VideoSamplePlan | null;
} | null {
  const children = parsePresentationMp4Boxes(
    bytes,
    trak.payloadStart,
    trak.end,
    budget,
  );
  const mdia = children?.filter((box) => box.kind === "mdia") ?? [];
  if (mdia.length !== 1) return null;
  const mdiaChildren = parsePresentationMp4Boxes(
    bytes,
    mdia[0].payloadStart,
    mdia[0].end,
    budget,
  );
  if (!mdiaChildren) return null;
  const handlers = mdiaChildren.filter((box) => box.kind === "hdlr");
  if (handlers.length !== 1) return null;
  const handler = handlers[0];
  if (
    handler.end - handler.payloadStart < 12 ||
    bytes[handler.payloadStart] !== 0
  ) {
    return null;
  }
  if (mp4Kind(bytes, handler.payloadStart + 8) !== "vide") {
    return { kind: "non-video" };
  }
  const mediaInfo = mdiaChildren.filter((box) => box.kind === "minf");
  if (mediaInfo.length !== 1) return null;
  const minfChildren = parsePresentationMp4Boxes(
    bytes,
    mediaInfo[0].payloadStart,
    mediaInfo[0].end,
    budget,
  );
  const sampleTables = minfChildren?.filter((box) => box.kind === "stbl") ?? [];
  if (sampleTables.length !== 1) return null;
  const sampleTableChildren = parsePresentationMp4Boxes(
    bytes,
    sampleTables[0].payloadStart,
    sampleTables[0].end,
    budget,
  );
  if (!sampleTableChildren) return null;
  const descriptions = sampleTableChildren.filter((box) => box.kind === "stsd");
  if (descriptions.length !== 1) return null;
  const supported = presentationStsdSupportedVideo(
    bytes,
    descriptions[0],
    budget,
  );
  return {
    kind: "video",
    plan: supported
      ? presentationVideoSamplePlan(bytes, sampleTableChildren, supported, mediaData)
      : null,
  };
}

async function presentationMp4HasSupportedVideo(file: File): Promise<boolean> {
  const budget: PresentationMp4Budget = {
    remaining: MAX_PRESENTATION_MP4_BOXES,
  };
  const topLevel = await readPresentationMp4TopLevel(file, budget);
  if (!topLevel) return false;
  const ftyp = topLevel.filter((box) => box.kind === "ftyp");
  const moov = topLevel.filter((box) => box.kind === "moov");
  const mdat = topLevel.filter((box) => box.kind === "mdat");
  if (ftyp.length !== 1 || moov.length !== 1 || mdat.length === 0) {
    return false;
  }
  if (moov[0].end - moov[0].payloadStart > MAX_PRESENTATION_MP4_MOOV_BYTES) {
    return false;
  }
  if (!mdat.every((box) => box.end > box.payloadStart)) return false;
  if (ftyp[0].end - ftyp[0].payloadStart > MAX_PRESENTATION_MP4_FTYP_BYTES) {
    return false;
  }
  const ftypBytes = await readPresentationMp4Range(
    file,
    ftyp[0].payloadStart,
    ftyp[0].end,
  );
  if (!ftypBytes || ftypBytes.length < 8 || ftypBytes.length % 4 !== 0) {
    return false;
  }
  const brands = new Set<string>();
  for (let offset = 0; offset < ftypBytes.length; offset += 4) {
    if (offset === 4) continue;
    const brand = mp4Kind(ftypBytes, offset);
    if (brand) brands.add(brand);
  }
  if (![...brands].some((brand) => PRESENTATION_MP4_BRANDS.has(brand))) {
    return false;
  }
  const moovBytes = await readPresentationMp4Range(
    file,
    moov[0].payloadStart,
    moov[0].end,
  );
  if (!moovBytes) return false;
  const moovChildren = parsePresentationMp4Boxes(
    moovBytes,
    0,
    moovBytes.length,
    budget,
  );
  if (!moovChildren) return false;
  const rebasedMediaData = mdat.map((box) => ({
    ...box,
    payloadStart: box.payloadStart,
    end: box.end,
  }));
  const dispositions = moovChildren
    .filter((box) => box.kind === "trak")
    .map((box) =>
      presentationTrakVideoDisposition(
        moovBytes,
        box,
        budget,
        rebasedMediaData,
      ),
    );
  if (dispositions.some((disposition) => disposition === null)) return false;
  const videoDispositions = dispositions.filter(
    (disposition): disposition is {
      readonly kind: "video";
      readonly plan: PresentationMp4VideoSamplePlan | null;
    } => disposition?.kind === "video",
  );
  if (videoDispositions.length !== 1 || !videoDispositions[0].plan) return false;
  const plan = videoDispositions[0].plan;
  const firstSample = await readPresentationMp4Range(
    file,
    plan.firstSampleStart,
    plan.firstSampleStart + plan.firstSampleSize,
  );
  return Boolean(
    firstSample && presentationAvcSampleHasVcl(firstSample, plan.nalLengthSize),
  );
}

/**
 * Validate the bounded PowerPoint-for-the-web local-file slice without reading
 * the whole movie into the browser heap. Top-level headers skip over `mdat`;
 * only the bounded `moov` graph is loaded. A real H.264 video track is required,
 * so corrupt, truncated, audio-only, and unsupported-codec MP4s never upload.
 */
export async function preparePresentationVideoFile(
  file: File,
): Promise<PreparedPresentationVideoFile> {
  if (!Number.isInteger(file.size) || file.size <= 0) fail("empty");
  if (file.size > MAX_PRESENTATION_VIDEO_FILE_BYTES) fail("too-large");
  const name = file.name.trim();
  const declaredType = file.type.trim().toLowerCase();
  if (
    !name.toLowerCase().endsWith(".mp4") ||
    !["", "application/octet-stream", "video/mp4"].includes(declaredType)
  ) {
    fail("unsupported-type");
  }
  try {
    if (!(await presentationMp4HasSupportedVideo(file))) fail("bad-magic");
  } catch (error) {
    if (error instanceof PresentationVideoFileError) throw error;
    fail("bad-magic");
  }
  return Object.freeze({
    mimeType: "video/mp4" as const,
    sizeBytes: file.size,
    // PowerPoint names a clip it inserts by its file's name without the
    // extension (office-truth probe media-names: "My Clip.mp4" -> "My Clip"),
    // the name the saved file states and a reopen reads.
    title: name.replace(/\.mp4$/i, "").trim().slice(0, 160) || "Video",
  });
}

async function compensateOrThrow(
  deleteUploadedAsset: (assetId: string) => Promise<void>,
  assetId: string,
  original: PresentationVideoInsertError,
): Promise<never> {
  try {
    await deleteUploadedAsset(assetId);
  } catch (error) {
    throw new PresentationVideoInsertError("compensation-failed", error);
  }
  throw original;
}

function targetSlideIndex<TSlide extends PresentationVideoInsertSlide>(
  slides: readonly TSlide[],
  activeIndex: number,
  targetSlideId: string,
): number {
  const active = slides[activeIndex];
  if (active?.id === targetSlideId) return activeIndex;
  const matches = slides.reduce<number[]>((indices, slide, index) => {
    if (slide.id === targetSlideId) indices.push(index);
    return indices;
  }, []);
  return matches.length === 1 ? matches[0] : -1;
}

function existingVideoAssetBudget<TSlide extends PresentationVideoInsertSlide>(
  slides: readonly TSlide[],
  replacing?: Pick<PresentationNativeVideoReplacementTarget, "slideId" | "blockId">,
): { count: number; sizeBytes: number } | null {
  const assets = new Map<string, number>();
  for (const slide of slides) {
    for (const block of slide.blocks) {
      if (slide.id === replacing?.slideId && block.id === replacing.blockId) {
        continue;
      }
      if (
        block.type !== "embed" ||
        block.kind !== "video" ||
        block.sourceKind !== "asset"
      ) {
        continue;
      }
      const source = assetBackedPresentationVideoSource(block);
      if (!source) return null;
      const existing = assets.get(source.assetId);
      if (existing !== undefined && existing !== source.sizeBytes) return null;
      assets.set(source.assetId, source.sizeBytes);
    }
  }
  return {
    count: assets.size,
    sizeBytes: [...assets.values()].reduce((total, size) => total + size, 0),
  };
}

/**
 * One local-video transaction: picker capture, bounded preflight, DataItem
 * upload, live-target revalidation, one editor commit, and compensating delete
 * on every post-upload rejection. Scratch and imported decks can append a
 * movie; an exactly selected imported movie is replaced instead.
 */
export function createPresentationVideoFileCommandAdapter<
  TSlide extends PresentationVideoInsertSlide,
>(accessors: PresentationVideoFileCommandAccessors<TSlide>) {
  const maxBlocks = accessors.maxBlocksPerSlide ?? 200;
  let pendingTarget:
    | {
        readonly kind: "insert";
        readonly slideId: string;
        readonly nativeDeck: boolean;
        /** An empty content placeholder the video takes the place of. */
        readonly placeholderId?: string;
      }
    | ({ readonly kind: "replace" } & PresentationNativeVideoReplacementTarget)
    | null = null;
  let busy = false;

  const canStart = (): boolean => {
    if (busy || pendingTarget !== null) return false;
    const slides = accessors.getSlides();
    const active = slides[accessors.getActiveIndex()];
    const nativeTarget = accessors.getNativeReplacementTarget();
    const nativeDeck = accessors.deckHasNativeSource(slides);
    const budget = existingVideoAssetBudget(
      slides,
      nativeDeck && nativeTarget ? nativeTarget : undefined,
    );
    return Boolean(
      active &&
        budget &&
        budget.count < MAX_PRESENTATION_VIDEO_ASSETS_PER_DECK &&
        budget.sizeBytes + presentationStoredMediaBytes(slides, { audioOnly: true }) <
          PRESENTATION_MEDIA_MAX_BYTES &&
        (nativeDeck
          ? nativeTarget?.slideId === active.id || active.blocks.length < maxBlocks
          : active.blocks.length < maxBlocks),
    );
  };

  /**
   * Opens the file picker. With `placeholderId` (a content placeholder's
   * Insert Video icon), the chosen video takes that empty placeholder's place
   * and geometry, as PowerPoint fills it; a cancelled picker changes nothing.
   */
  const execute = (options?: { readonly placeholderId?: string }): boolean => {
    if (!canStart()) return false;
    const slides = accessors.getSlides();
    const target = slides[accessors.getActiveIndex()];
    if (!target) return false;
    const nativeTarget = accessors.getNativeReplacementTarget();
    const nativeDeck = accessors.deckHasNativeSource(slides);
    const placeholderId = options?.placeholderId;
    pendingTarget =
      placeholderId && target.blocks.some((block) => block.id === placeholderId)
        ? { kind: "insert", slideId: target.id, nativeDeck, placeholderId }
        : nativeDeck && nativeTarget
          ? { kind: "replace", ...nativeTarget }
          : { kind: "insert", slideId: target.id, nativeDeck };
    if (accessors.openFilePicker()) return true;
    pendingTarget = null;
    return false;
  };

  const cancelPicker = (): void => {
    if (!busy) pendingTarget = null;
  };

  const insertFile = async (file: File): Promise<boolean> => {
    if (busy) throw new PresentationVideoInsertError("busy");
    const targetIntent = pendingTarget;
    pendingTarget = null;
    if (!targetIntent) throw new PresentationVideoInsertError("no-target");
    busy = true;
    let assetId: string | null = null;
    try {
      const prepared = await accessors.prepareFile(file);
      let slides = accessors.getSlides();
      let targetIndex = targetSlideIndex(
        slides,
        accessors.getActiveIndex(),
        targetIntent.slideId,
      );
      if (targetIndex < 0) throw new PresentationVideoInsertError("stale-target");
      if (
        targetIntent.kind === "insert" &&
        !targetIntent.placeholderId &&
        slides[targetIndex].blocks.length >= maxBlocks
      ) {
        throw new PresentationVideoInsertError("slide-full");
      }
      if (
        targetIntent.kind === "replace" &&
        !nativeReplacementTargetMatches(
          accessors.getNativeReplacementTarget(),
          targetIntent,
        )
      ) {
        throw new PresentationVideoInsertError("stale-target");
      }
      const beforeUploadBudget = existingVideoAssetBudget(
        slides,
        targetIntent.kind === "replace" ? targetIntent : undefined,
      );
      if (
        !beforeUploadBudget ||
        beforeUploadBudget.count >= MAX_PRESENTATION_VIDEO_ASSETS_PER_DECK
      ) {
        throw new PresentationVideoInsertError("deck-budget");
      }
      // The deck's media total: its videos, stored clips and sounds together
      // (lib/document-size-limits), said as such before the upload.
      requireWithinDocumentTotal(
        "presentation",
        beforeUploadBudget.sizeBytes +
          presentationStoredMediaBytes(slides, { audioOnly: true }),
        prepared.sizeBytes,
      );

      const uploadedId = await accessors.uploadFile(file);
      assetId = typeof uploadedId === "string" ? uploadedId : "";
      if (
        !assetId ||
        assetId !== assetId.trim() ||
        assetId.length > 200
      ) {
        if (assetId) {
          await compensateOrThrow(
            accessors.deleteUploadedAsset,
            assetId,
            new PresentationVideoInsertError("invalid-asset-id"),
          );
        }
        throw new PresentationVideoInsertError("invalid-asset-id");
      }

      try {
        await accessors.attestUploadedAsset(assetId, prepared);
      } catch (error) {
        return compensateOrThrow(
          accessors.deleteUploadedAsset,
          assetId,
          new PresentationVideoInsertError("commit-failed", error),
        );
      }

      slides = accessors.getSlides();
      const nativeDeck = accessors.deckHasNativeSource(slides);
      if (
        (targetIntent.kind === "insert" &&
          nativeDeck !== targetIntent.nativeDeck) ||
        (targetIntent.kind === "replace" &&
          (!nativeDeck ||
            !nativeReplacementTargetMatches(
              accessors.getNativeReplacementTarget(),
              targetIntent,
            )))
      ) {
        return compensateOrThrow(
          accessors.deleteUploadedAsset,
          assetId,
          new PresentationVideoInsertError("stale-target"),
        );
      }
      targetIndex = targetSlideIndex(
        slides,
        accessors.getActiveIndex(),
        targetIntent.slideId,
      );
      if (targetIndex < 0) {
        return compensateOrThrow(
          accessors.deleteUploadedAsset,
          assetId,
          new PresentationVideoInsertError("stale-target"),
        );
      }
      const target = slides[targetIndex];
      if (targetIntent.kind === "insert" && target.blocks.length >= maxBlocks) {
        return compensateOrThrow(
          accessors.deleteUploadedAsset,
          assetId,
          new PresentationVideoInsertError("slide-full"),
        );
      }
      let blockId = targetIntent.kind === "replace" ? targetIntent.blockId : "";
      let block: Readonly<Record<string, unknown>>;
      let placeholderFilled: string | null = null;
      if (targetIntent.kind === "replace") {
        const live = target.blocks.find((candidate) => candidate.id === blockId);
        if (!live || nativePresentationVideoAuthority(live) === null) {
          return compensateOrThrow(
            accessors.deleteUploadedAsset,
            assetId,
            new PresentationVideoInsertError("stale-target"),
          );
        }
        block = Object.freeze({
          ...live,
          url: "",
          sourceKind: "asset" as const,
          assetId,
          mimeType: prepared.mimeType,
          sizeBytes: prepared.sizeBytes,
        });
      } else {
        try {
          blockId = accessors.nextId();
        } catch (error) {
          return compensateOrThrow(
            accessors.deleteUploadedAsset,
            assetId,
            new PresentationVideoInsertError("commit-failed", error),
          );
        }
        if (
          !blockId ||
          slides.some(
            (slide) =>
              slide.id === blockId ||
              slide.blocks.some((candidate) => candidate.id === blockId),
          )
        ) {
          return compensateOrThrow(
            accessors.deleteUploadedAsset,
            assetId,
            new PresentationVideoInsertError("identity-collision"),
          );
        }
        const holder =
          targetIntent.kind === "insert" && targetIntent.placeholderId
            ? target.blocks.find((candidate) => candidate.id === targetIntent.placeholderId)
            : undefined;
        const number = (value: unknown) =>
          typeof value === "number" && Number.isFinite(value) ? value : null;
        const holderGeometry =
          holder &&
          number(holder.width) !== null &&
          number(holder.height) !== null
            ? {
                left: number(holder.left) ?? 0,
                top: number(holder.top) ?? 0,
                width: number(holder.width) as number,
                height: number(holder.height) as number,
              }
            : null;
        block = Object.freeze({
          id: blockId,
          type: "embed" as const,
          kind: "video" as const,
          url: "",
          sourceKind: "asset" as const,
          assetId,
          mimeType: prepared.mimeType,
          sizeBytes: prepared.sizeBytes,
          title: prepared.title,
          ...(holderGeometry ?? defaultEmbedGeometry("video", accessors.getSlideBounds(target))),
          // The video stays its placeholder's content (PowerPoint writes it
          // on the placeholder), so a layout reset keeps it in that place.
          ...(holder && typeof holder.placeholder === "string" ? { placeholder: holder.placeholder } : {}),
          ...(holder && typeof holder.role === "string" ? { role: holder.role } : {}),
        });
        if (holder) placeholderFilled = holder.id as string;
      }
      const nextSlides = slides.map((slide, index) =>
        index === targetIndex
          ? ({
              ...slide,
              blocks:
                targetIntent.kind === "replace"
                  ? slide.blocks.map((candidate) =>
                      candidate.id === blockId ? block : candidate,
                    )
                  : placeholderFilled
                    ? slide.blocks.map((candidate) =>
                        candidate.id === placeholderFilled ? block : candidate,
                      )
                    : [...slide.blocks, block],
            } as TSlide)
          : slide,
      );
      try {
        accessors.commitSlides(nextSlides);
      } catch (error) {
        return compensateOrThrow(
          accessors.deleteUploadedAsset,
          assetId,
          new PresentationVideoInsertError("commit-failed", error),
        );
      }
      accessors.selectInsertedVideo(blockId, targetIndex);
      return true;
    } finally {
      busy = false;
    }
  };

  return Object.freeze({
    canStart,
    execute,
    cancelPicker,
    insertFile,
    isBusy: () => busy,
  });
}
