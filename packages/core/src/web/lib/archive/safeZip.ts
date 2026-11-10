/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { Inflate } from "fflate";
import JSZip from "jszip";

const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50;
const CENTRAL_DIRECTORY_FILE_SIGNATURE = 0x02014b50;
const CENTRAL_DIRECTORY_DIGITAL_SIGNATURE = 0x05054b50;
const LOCAL_FILE_HEADER_SIGNATURE = 0x04034b50;
const MAX_ZIP_COMMENT_BYTES = 0xffff;
const END_OF_CENTRAL_DIRECTORY_BYTES = 22;
const CENTRAL_DIRECTORY_FILE_HEADER_BYTES = 46;
const LOCAL_FILE_HEADER_BYTES = 30;
const MAX_SAFE_ZIP_PATH_CHARS = 1024;
const UTF8_FILENAME_FLAG = 1 << 11;
const ENCRYPTED_FLAG = 1;
const DATA_DESCRIPTOR_FLAG = 1 << 3;
const DATA_DESCRIPTOR_SIGNATURE = 0x08074b50;
const DATA_DESCRIPTOR_BYTES = 12;
const SIGNED_DATA_DESCRIPTOR_BYTES = 16;
const ZIP64_U16_SENTINEL = 0xffff;
const ZIP64_U32_SENTINEL = 0xffffffff;
const STORE_METHOD = 0;
const DEFLATE_METHOD = 8;
// A small compressed input slice keeps a forged low-size DEFLATE stream from
// making the inflater allocate the complete output before we can enforce the
// real byte count. The largest temporary output is bounded by one slice rather
// than by the archive or entry size.
const INFLATE_INPUT_CHUNK_BYTES = 1024;

export interface SafeZipLimits {
  /** Compressed input bytes allowed before a Blob is materialized in memory. */
  maxCompressedBytes: number;
  /** Every central-directory member counts, including directories and ignored files. */
  maxEntries: number;
  /** Declared uncompressed bytes allowed for one member. */
  maxEntryUncompressedBytes: number;
  /** Declared uncompressed bytes allowed across the complete archive. */
  maxTotalUncompressedBytes: number;
  /** Maximum declared-uncompressed / compressed ratio for large DEFLATE members. */
  maxCompressionRatio: number;
  /** Avoid ratio false positives on tiny metadata parts. */
  minRatioCheckUncompressedBytes: number;
}

export type SafeZipInput = Blob | ArrayBuffer | ArrayBufferView;

export class SafeZipError extends Error {
  constructor(
    readonly code: string,
    detail?: string,
  ) {
    super(detail ? `${code}: ${detail}` : code);
    this.name = "SafeZipError";
  }
}

interface CentralDirectoryEntry {
  compressedSize: number;
  compressionMethod: number;
  crc32: number;
  dataEnd: number;
  dataOffset: number;
  isDirectory: boolean;
  localRecordEnd: number;
  localHeaderOffset: number;
  normalizedPath: string;
  rawPath: string;
  uncompressedSize: number;
}

function malformed(detail: string): never {
  throw new SafeZipError("zip_central_directory_invalid", detail);
}

function readUint16(view: DataView, offset: number, limit: number): number {
  if (offset < 0 || offset + 2 > limit) malformed("truncated uint16");
  return view.getUint16(offset, true);
}

function readUint32(view: DataView, offset: number, limit: number): number {
  if (offset < 0 || offset + 4 > limit) malformed("truncated uint32");
  return view.getUint32(offset, true);
}

function findEndOfCentralDirectory(bytes: Uint8Array, view: DataView): number {
  const firstPossible = Math.max(
    0,
    bytes.byteLength - END_OF_CENTRAL_DIRECTORY_BYTES - MAX_ZIP_COMMENT_BYTES,
  );
  for (
    let offset = bytes.byteLength - END_OF_CENTRAL_DIRECTORY_BYTES;
    offset >= firstPossible;
    offset -= 1
  ) {
    if (
      readUint32(view, offset, bytes.byteLength) !==
      END_OF_CENTRAL_DIRECTORY_SIGNATURE
    ) {
      continue;
    }
    const commentLength = readUint16(view, offset + 20, bytes.byteLength);
    if (
      offset + END_OF_CENTRAL_DIRECTORY_BYTES + commentLength ===
      bytes.byteLength
    ) {
      return offset;
    }
  }
  malformed("end-of-central-directory record is missing");
}

function decodeFilename(bytes: Uint8Array, utf8Flag: boolean): string {
  const asciiOnly = bytes.every((byte) => byte <= 0x7f);
  if (!asciiOnly && !utf8Flag) {
    throw new SafeZipError(
      "zip_filename_encoding_unsupported",
      "non-ASCII filenames must declare UTF-8",
    );
  }
  if (asciiOnly) {
    let result = "";
    for (const byte of bytes) result += String.fromCharCode(byte);
    return result;
  }
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    throw new SafeZipError("zip_filename_encoding_invalid");
  }
}

function normalizeArchivePath(path: string): string {
  if (!path || path.length > MAX_SAFE_ZIP_PATH_CHARS) {
    throw new SafeZipError("zip_path_unsafe", "empty or overlong path");
  }
  if (
    path.startsWith("/") ||
    /^[A-Za-z]:/.test(path) ||
    path.includes("\\") ||
    /[\u0000-\u001f\u007f]/.test(path)
  ) {
    throw new SafeZipError("zip_path_unsafe", path);
  }

  const directory = path.endsWith("/");
  const body = directory ? path.slice(0, -1) : path;
  const parts = body.split("/");
  if (!body || parts.some((part) => !part || part === "." || part === "..")) {
    throw new SafeZipError("zip_path_unsafe", path);
  }

  // NFC gives collision checks one stable representation while preserving the
  // original JSZip entry names for the caller. It also prevents visually
  // identical decomposed/composed names from bypassing duplicate detection.
  return `${parts.join("/").normalize("NFC")}${directory ? "/" : ""}`;
}

function equalBytes(left: Uint8Array, right: Uint8Array): boolean {
  if (left.byteLength !== right.byteLength) return false;
  for (let index = 0; index < left.byteLength; index += 1) {
    if (left[index] !== right[index]) return false;
  }
  return true;
}

function dataDescriptorEnd(
  view: DataView,
  offset: number,
  limit: number,
  expectedCrc32: number,
  expectedCompressedSize: number,
  expectedUncompressedSize: number,
): number {
  const matches = (recordOffset: number): boolean =>
    recordOffset + DATA_DESCRIPTOR_BYTES <= limit &&
    readUint32(view, recordOffset, limit) === expectedCrc32 &&
    readUint32(view, recordOffset + 4, limit) === expectedCompressedSize &&
    readUint32(view, recordOffset + 8, limit) === expectedUncompressedSize;

  if (
    offset + SIGNED_DATA_DESCRIPTOR_BYTES <= limit &&
    readUint32(view, offset, limit) === DATA_DESCRIPTOR_SIGNATURE &&
    matches(offset + 4)
  ) {
    return offset + SIGNED_DATA_DESCRIPTOR_BYTES;
  }
  // The signature is optional. Trying the unsigned form second also handles
  // the legal edge case where the CRC itself equals 0x08074b50.
  if (matches(offset)) return offset + DATA_DESCRIPTOR_BYTES;
  malformed("data descriptor does not match the central directory");
}

function centralDirectoryEntries(
  bytes: Uint8Array,
  limits: SafeZipLimits,
): CentralDirectoryEntry[] {
  if (bytes.byteLength < END_OF_CENTRAL_DIRECTORY_BYTES) {
    malformed("archive is too short");
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocdOffset = findEndOfCentralDirectory(bytes, view);
  const diskNumber = readUint16(view, eocdOffset + 4, bytes.byteLength);
  const centralDirectoryDisk = readUint16(
    view,
    eocdOffset + 6,
    bytes.byteLength,
  );
  const entriesOnDisk = readUint16(view, eocdOffset + 8, bytes.byteLength);
  const entryCount = readUint16(view, eocdOffset + 10, bytes.byteLength);
  const centralDirectorySize = readUint32(
    view,
    eocdOffset + 12,
    bytes.byteLength,
  );
  const centralDirectoryOffset = readUint32(
    view,
    eocdOffset + 16,
    bytes.byteLength,
  );

  if (
    diskNumber !== 0 ||
    centralDirectoryDisk !== 0 ||
    entriesOnDisk !== entryCount
  ) {
    throw new SafeZipError("zip_multivolume_unsupported");
  }
  if (
    entryCount === ZIP64_U16_SENTINEL ||
    centralDirectorySize === ZIP64_U32_SENTINEL ||
    centralDirectoryOffset === ZIP64_U32_SENTINEL
  ) {
    throw new SafeZipError("zip64_unsupported");
  }
  if (entryCount > limits.maxEntries) {
    throw new SafeZipError(
      "zip_entry_count_exceeded",
      `${entryCount} > ${limits.maxEntries}`,
    );
  }

  const centralDirectoryEnd = centralDirectoryOffset + centralDirectorySize;
  if (
    centralDirectoryOffset < 0 ||
    centralDirectoryEnd > eocdOffset ||
    centralDirectoryEnd < centralDirectoryOffset
  ) {
    malformed("central-directory bounds are invalid");
  }

  const entries: CentralDirectoryEntry[] = [];
  const normalizedPaths = new Set<string>();
  let cursor = centralDirectoryOffset;
  for (let index = 0; index < entryCount; index += 1) {
    if (
      cursor + CENTRAL_DIRECTORY_FILE_HEADER_BYTES > centralDirectoryEnd ||
      readUint32(view, cursor, centralDirectoryEnd) !==
        CENTRAL_DIRECTORY_FILE_SIGNATURE
    ) {
      malformed(`entry ${index} header is missing`);
    }

    const flags = readUint16(view, cursor + 8, centralDirectoryEnd);
    const compressionMethod = readUint16(
      view,
      cursor + 10,
      centralDirectoryEnd,
    );
    const crc32 = readUint32(view, cursor + 16, centralDirectoryEnd);
    const compressedSize = readUint32(view, cursor + 20, centralDirectoryEnd);
    const uncompressedSize = readUint32(view, cursor + 24, centralDirectoryEnd);
    const filenameLength = readUint16(view, cursor + 28, centralDirectoryEnd);
    const extraLength = readUint16(view, cursor + 30, centralDirectoryEnd);
    const commentLength = readUint16(view, cursor + 32, centralDirectoryEnd);
    const diskStart = readUint16(view, cursor + 34, centralDirectoryEnd);
    const localHeaderOffset = readUint32(
      view,
      cursor + 42,
      centralDirectoryEnd,
    );
    const next =
      cursor +
      CENTRAL_DIRECTORY_FILE_HEADER_BYTES +
      filenameLength +
      extraLength +
      commentLength;
    if (next > centralDirectoryEnd) malformed(`entry ${index} is truncated`);
    if (
      compressedSize === ZIP64_U32_SENTINEL ||
      uncompressedSize === ZIP64_U32_SENTINEL ||
      localHeaderOffset === ZIP64_U32_SENTINEL ||
      diskStart === ZIP64_U16_SENTINEL
    ) {
      throw new SafeZipError("zip64_unsupported");
    }
    if (diskStart !== 0) throw new SafeZipError("zip_multivolume_unsupported");
    if ((flags & ENCRYPTED_FLAG) !== 0) {
      throw new SafeZipError("zip_encryption_unsupported");
    }
    if (
      compressionMethod !== STORE_METHOD &&
      compressionMethod !== DEFLATE_METHOD
    ) {
      throw new SafeZipError(
        "zip_compression_method_unsupported",
        String(compressionMethod),
      );
    }
    if (
      localHeaderOffset + LOCAL_FILE_HEADER_BYTES > centralDirectoryOffset ||
      readUint32(view, localHeaderOffset, centralDirectoryOffset) !==
        LOCAL_FILE_HEADER_SIGNATURE
    ) {
      malformed(`entry ${index} local header is invalid`);
    }

    const filenameStart = cursor + CENTRAL_DIRECTORY_FILE_HEADER_BYTES;
    const filenameBytes = bytes.subarray(
      filenameStart,
      filenameStart + filenameLength,
    );
    const rawPath = decodeFilename(
      filenameBytes,
      (flags & UTF8_FILENAME_FLAG) !== 0,
    );
    const normalizedPath = normalizeArchivePath(rawPath);
    if (normalizedPaths.has(normalizedPath)) {
      throw new SafeZipError("zip_duplicate_path", normalizedPath);
    }
    normalizedPaths.add(normalizedPath);

    const localFlags = readUint16(
      view,
      localHeaderOffset + 6,
      centralDirectoryOffset,
    );
    const localCompressionMethod = readUint16(
      view,
      localHeaderOffset + 8,
      centralDirectoryOffset,
    );
    const localCrc32 = readUint32(
      view,
      localHeaderOffset + 14,
      centralDirectoryOffset,
    );
    const localCompressedSize = readUint32(
      view,
      localHeaderOffset + 18,
      centralDirectoryOffset,
    );
    const localUncompressedSize = readUint32(
      view,
      localHeaderOffset + 22,
      centralDirectoryOffset,
    );
    const localFilenameLength = readUint16(
      view,
      localHeaderOffset + 26,
      centralDirectoryOffset,
    );
    const localExtraLength = readUint16(
      view,
      localHeaderOffset + 28,
      centralDirectoryOffset,
    );
    const localFilenameStart = localHeaderOffset + LOCAL_FILE_HEADER_BYTES;
    const dataOffset =
      localFilenameStart + localFilenameLength + localExtraLength;
    const dataEnd = dataOffset + compressedSize;
    if (
      dataOffset < localFilenameStart ||
      dataOffset > centralDirectoryOffset ||
      dataEnd < dataOffset ||
      dataEnd > centralDirectoryOffset
    ) {
      malformed(`entry ${index} local payload bounds are invalid`);
    }
    if (localFlags !== flags || localCompressionMethod !== compressionMethod) {
      malformed(`entry ${index} local flags or compression method disagree`);
    }
    const localFilenameBytes = bytes.subarray(
      localFilenameStart,
      localFilenameStart + localFilenameLength,
    );
    if (!equalBytes(localFilenameBytes, filenameBytes)) {
      malformed(`entry ${index} local filename disagrees`);
    }
    // Decode through the same strict policy as the central name. This makes
    // the local UTF-8 flag and filename byte sequence part of one identity.
    const localPath = decodeFilename(
      localFilenameBytes,
      (localFlags & UTF8_FILENAME_FLAG) !== 0,
    );
    if (localPath !== rawPath) {
      malformed(`entry ${index} local filename identity disagrees`);
    }

    const usesDataDescriptor = (flags & DATA_DESCRIPTOR_FLAG) !== 0;
    if (!usesDataDescriptor) {
      if (
        localCrc32 !== crc32 ||
        localCompressedSize !== compressedSize ||
        localUncompressedSize !== uncompressedSize
      ) {
        malformed(`entry ${index} local size or CRC disagrees`);
      }
    } else {
      const localValuesAreEmpty =
        localCrc32 === 0 &&
        localCompressedSize === 0 &&
        localUncompressedSize === 0;
      const localValuesMatch =
        localCrc32 === crc32 &&
        localCompressedSize === compressedSize &&
        localUncompressedSize === uncompressedSize;
      if (!localValuesAreEmpty && !localValuesMatch) {
        malformed(`entry ${index} local descriptor placeholders disagree`);
      }
    }
    const localRecordEnd = usesDataDescriptor
      ? dataDescriptorEnd(
          view,
          dataEnd,
          centralDirectoryOffset,
          crc32,
          compressedSize,
          uncompressedSize,
        )
      : dataEnd;
    entries.push({
      compressedSize,
      compressionMethod,
      crc32,
      dataEnd,
      dataOffset,
      isDirectory: normalizedPath.endsWith("/"),
      localHeaderOffset,
      localRecordEnd,
      normalizedPath,
      rawPath,
      uncompressedSize,
    });
    cursor = next;
  }

  // A central-directory digital signature is legal but carries no file entry.
  // Accept that one bounded record; reject every other trailing structure.
  if (
    cursor + 6 <= centralDirectoryEnd &&
    readUint32(view, cursor, centralDirectoryEnd) ===
      CENTRAL_DIRECTORY_DIGITAL_SIGNATURE
  ) {
    const signatureLength = readUint16(view, cursor + 4, centralDirectoryEnd);
    cursor += 6 + signatureLength;
  }
  if (cursor !== centralDirectoryEnd) {
    malformed("central-directory size does not match its entries");
  }

  const localRecords = [...entries].sort(
    (left, right) => left.localHeaderOffset - right.localHeaderOffset,
  );
  for (let index = 1; index < localRecords.length; index += 1) {
    if (
      localRecords[index].localHeaderOffset <
      localRecords[index - 1].localRecordEnd
    ) {
      malformed("local file records overlap");
    }
  }
  return entries;
}

function enforceInflateLimits(
  entries: CentralDirectoryEntry[],
  limits: SafeZipLimits,
): void {
  let totalUncompressedBytes = 0;
  for (const entry of entries) {
    if (
      entry.isDirectory &&
      (entry.compressedSize !== 0 || entry.uncompressedSize !== 0)
    ) {
      throw new SafeZipError(
        "zip_directory_payload_unsupported",
        entry.normalizedPath,
      );
    }
    if (entry.uncompressedSize > limits.maxEntryUncompressedBytes) {
      throw new SafeZipError(
        "zip_entry_uncompressed_size_exceeded",
        `${entry.normalizedPath}: ${entry.uncompressedSize} > ${limits.maxEntryUncompressedBytes}`,
      );
    }
    totalUncompressedBytes += entry.uncompressedSize;
    if (totalUncompressedBytes > limits.maxTotalUncompressedBytes) {
      throw new SafeZipError(
        "zip_total_uncompressed_size_exceeded",
        `${totalUncompressedBytes} > ${limits.maxTotalUncompressedBytes}`,
      );
    }
    if (
      entry.compressionMethod === DEFLATE_METHOD &&
      entry.uncompressedSize >= limits.minRatioCheckUncompressedBytes
    ) {
      const ratio =
        entry.compressedSize === 0
          ? Number.POSITIVE_INFINITY
          : entry.uncompressedSize / entry.compressedSize;
      if (ratio > limits.maxCompressionRatio) {
        throw new SafeZipError(
          "zip_compression_ratio_exceeded",
          `${entry.normalizedPath}: ${ratio.toFixed(1)} > ${limits.maxCompressionRatio}`,
        );
      }
    }
  }
}

let crcTable: Uint32Array | undefined;

function getCrcTable(): Uint32Array {
  if (crcTable) return crcTable;
  const table = new Uint32Array(256);
  for (let value = 0; value < table.length; value += 1) {
    let crc = value;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = (crc & 1) !== 0 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1;
    }
    table[value] = crc >>> 0;
  }
  crcTable = table;
  return table;
}

function updateCrc32(crc: number, chunk: Uint8Array): number {
  const table = getCrcTable();
  let next = crc;
  for (const byte of chunk) {
    next = table[(next ^ byte) & 0xff] ^ (next >>> 8);
  }
  return next >>> 0;
}

function validateInflatedEntries(
  bytes: Uint8Array,
  entries: CentralDirectoryEntry[],
  limits: SafeZipLimits,
): void {
  let totalActualBytes = 0;

  for (const entry of entries) {
    const compressed = bytes.subarray(entry.dataOffset, entry.dataEnd);
    let actualBytes = 0;
    let crc = 0xffffffff;
    let outputError: SafeZipError | undefined;

    const account = (chunk: Uint8Array): void => {
      if (outputError) return;
      actualBytes += chunk.byteLength;
      if (actualBytes > limits.maxEntryUncompressedBytes) {
        outputError = new SafeZipError(
          "zip_entry_uncompressed_size_exceeded",
          `${entry.normalizedPath}: actual output exceeds ${limits.maxEntryUncompressedBytes}`,
        );
        return;
      }
      totalActualBytes += chunk.byteLength;
      if (totalActualBytes > limits.maxTotalUncompressedBytes) {
        outputError = new SafeZipError(
          "zip_total_uncompressed_size_exceeded",
          `actual output exceeds ${limits.maxTotalUncompressedBytes}`,
        );
        return;
      }
      crc = updateCrc32(crc, chunk);
    };

    if (entry.compressionMethod === STORE_METHOD) {
      for (
        let offset = 0;
        offset < compressed.byteLength && !outputError;
        offset += INFLATE_INPUT_CHUNK_BYTES
      ) {
        account(
          compressed.subarray(
            offset,
            Math.min(offset + INFLATE_INPUT_CHUNK_BYTES, compressed.byteLength),
          ),
        );
      }
    } else {
      const inflate = new Inflate((chunk) => account(chunk));
      try {
        if (compressed.byteLength === 0) {
          inflate.push(compressed, true);
        } else {
          for (
            let offset = 0;
            offset < compressed.byteLength && !outputError;
            offset += INFLATE_INPUT_CHUNK_BYTES
          ) {
            const end = Math.min(
              offset + INFLATE_INPUT_CHUNK_BYTES,
              compressed.byteLength,
            );
            inflate.push(
              compressed.subarray(offset, end),
              end === compressed.byteLength,
            );
          }
        }
      } catch (error) {
        throw new SafeZipError(
          "zip_deflate_invalid",
          error instanceof Error ? error.message : String(error),
        );
      }
    }

    if (outputError) throw outputError;
    if (actualBytes !== entry.uncompressedSize) {
      throw new SafeZipError(
        "zip_uncompressed_size_mismatch",
        `${entry.normalizedPath}: ${actualBytes} != ${entry.uncompressedSize}`,
      );
    }
    if (
      entry.compressionMethod === DEFLATE_METHOD &&
      actualBytes >= limits.minRatioCheckUncompressedBytes
    ) {
      const actualRatio =
        entry.compressedSize === 0
          ? Number.POSITIVE_INFINITY
          : actualBytes / entry.compressedSize;
      if (actualRatio > limits.maxCompressionRatio) {
        throw new SafeZipError(
          "zip_compression_ratio_exceeded",
          `${entry.normalizedPath}: actual ${actualRatio.toFixed(1)} > ${limits.maxCompressionRatio}`,
        );
      }
    }
    const actualCrc32 = (crc ^ 0xffffffff) >>> 0;
    if (actualCrc32 !== entry.crc32) {
      throw new SafeZipError("zip_crc32_mismatch", entry.normalizedPath);
    }
  }
}

function validateLoadedIdentities(
  zip: JSZip,
  entries: CentralDirectoryEntry[],
): void {
  const loadedNames = Object.keys(zip.files);
  if (loadedNames.length !== entries.length) {
    malformed("JSZip changed the number of archive identities");
  }
  const expected = new Map(entries.map((entry) => [entry.rawPath, entry]));
  for (const name of loadedNames) {
    const entry = expected.get(name);
    const loaded = zip.files[name];
    if (!entry || loaded.dir !== entry.isDirectory) {
      malformed("JSZip changed an archive path identity");
    }
  }
}

async function blobBytes(blob: Blob): Promise<ArrayBuffer> {
  if (typeof blob.arrayBuffer === "function") return blob.arrayBuffer();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () =>
      reject(reader.error || new Error("Failed to read ZIP input."));
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.readAsArrayBuffer(blob);
  });
}

async function inputBytes(input: SafeZipInput): Promise<Uint8Array> {
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (ArrayBuffer.isView(input)) {
    return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  }
  return new Uint8Array(await blobBytes(input));
}

function inputByteLength(input: SafeZipInput): number {
  if (input instanceof ArrayBuffer) return input.byteLength;
  if (ArrayBuffer.isView(input)) return input.byteLength;
  return input.size;
}

/**
 * Parse central and local records, enforce declared budgets, then validate
 * every member with bounded streaming inflate and incremental CRC accounting.
 * Only after that boundary may JSZip index the already-verified archive; its
 * eager `checkCRC32` mode is deliberately disabled because it has no output
 * budget and can allocate a forged member in full.
 */
/** Each entry's unpacked size as the archive's central directory states it,
 * by name, read without inflating anything (null: the archive is past `limits`
 * or not readable, which `loadZipSafely` reports properly). */
export async function zipEntrySizes(
  input: SafeZipInput,
  limits: SafeZipLimits,
): Promise<ReadonlyMap<string, number> | null> {
  try {
    if (inputByteLength(input) > limits.maxCompressedBytes) return null;
    const entries = centralDirectoryEntries(await inputBytes(input), limits);
    return new Map(entries.map((entry) => [entry.normalizedPath, entry.uncompressedSize]));
  } catch {
    return null;
  }
}

export async function loadZipSafely(
  input: SafeZipInput,
  limits: SafeZipLimits,
): Promise<JSZip> {
  const compressedBytes = inputByteLength(input);
  if (compressedBytes > limits.maxCompressedBytes) {
    throw new SafeZipError(
      "zip_compressed_size_exceeded",
      `${compressedBytes} > ${limits.maxCompressedBytes}`,
    );
  }
  const bytes = await inputBytes(input);
  const entries = centralDirectoryEntries(bytes, limits);
  enforceInflateLimits(entries, limits);
  validateInflatedEntries(bytes, entries, limits);
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes, { checkCRC32: false });
  } catch (error) {
    malformed(
      error instanceof Error
        ? `JSZip rejected archive: ${error.message}`
        : "JSZip rejected archive",
    );
  }
  validateLoadedIdentities(zip, entries);
  return zip;
}
