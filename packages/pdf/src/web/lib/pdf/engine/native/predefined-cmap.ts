/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { PDF_JS_CMAP_BASE_URL } from "@mrnote/core/web/lib/pdf/pdfjs-asset-urls.ts";
import { isMraiNativePredefinedUnicodeCMapName } from "./predefined-cmap-names";
export { isMraiNativePredefinedUnicodeCMapName };
/**
 * Predefined CMap authority for native text editing.
 *
 * Type0 fonts that encode text through one of Adobe's Unicode UCS-2 CMaps
 * (`UniGB-UCS2-H`, `UniJIS-UCS2-H`, ...) carry their character codes as
 * UTF-16 code units. The same packed binary CMaps pdf.js renders with
 * (`public/pdfjs-cmaps/*.bcmap`) are the code → CID authority here, so the
 * editor never invents a mapping the renderer does not use. Only CMaps whose
 * every codespace range is two bytes wide are accepted; mixed-width legacy
 * encodings (GBK-EUC-H, 90ms-RKSJ-H, ...) stay preserve-only.
 */

/** The URL pdf.js itself reads them from, so one kept copy serves both. */
export const PDF_NATIVE_PREDEFINED_CMAP_BASE_URL = PDF_JS_CMAP_BASE_URL;

const MAX_BCMAP_BYTES = 4 * 1024 * 1024;
const MAX_USE_CMAP_DEPTH = 4;

export interface PdfNativePredefinedCMap {
  readonly name: string;
  readonly vertical: boolean;
  /** Two-byte character code → CID for every mapped code. */
  readonly codeToCid: ReadonlyMap<number, number>;
  /** An Identity font's ordering: what the reader reads each CID as. */
  readonly cidReading?: ReadonlyMap<number, string>;
}

export type PdfNativePredefinedCMapLoader = (name: string) => Promise<Uint8Array>;


let activeLoader: PdfNativePredefinedCMapLoader | null = null;
const cache = new Map<string, Promise<PdfNativePredefinedCMap>>();

/** Tests and probes read the packed CMaps from disk instead of the app origin. */
export function setMraiNativePredefinedCMapLoader(
  loader: PdfNativePredefinedCMapLoader | null,
): void {
  activeLoader = loader;
  cache.clear();
  orderingReadings.clear();
}

async function defaultLoader(name: string): Promise<Uint8Array> {
  const response = await fetch(`${PDF_NATIVE_PREDEFINED_CMAP_BASE_URL}${name}.bcmap`);
  if (!response.ok) throw new Error(`Predefined CMap ${name} is unavailable.`);
  return new Uint8Array(await response.arrayBuffer());
}

export async function loadMraiNativePredefinedCMap(
  name: string,
): Promise<PdfNativePredefinedCMap> {
  if (!isMraiNativePredefinedUnicodeCMapName(name)) {
    throw new Error(`Predefined CMap ${name} is not a supported Unicode CMap.`);
  }
  let pending = cache.get(name);
  if (!pending) {
    pending = resolveCMap(name, 0);
    cache.set(name, pending);
    pending.catch(() => cache.delete(name));
  }
  return pending;
}

/** The code space of any predefined CMap, mixed widths included (one
 * usecmap level): what a reader needs to cut a show into its codes. It is
 * read-only authority for cutting codes, never for mapping them. */
export async function readMraiPredefinedCMapCodespace(
  name: string,
): Promise<readonly ParsedCodespaceRange[]> {
  if (!/^[A-Za-z0-9-]{1,64}$/.test(name)) throw new Error("Predefined CMap name is invalid.");
  const load = activeLoader ?? defaultLoader;
  const parsed = parseMraiBinaryCMap(await load(name));
  if (parsed.codespace.length) return parsed.codespace;
  if (!parsed.useCMap || !/^[A-Za-z0-9-]{1,64}$/.test(parsed.useCMap)) return [];
  return parseMraiBinaryCMap(await load(parsed.useCMap)).codespace;
}

const orderingReadings = new Map<string, Promise<ReadonlyMap<number, string>>>();

/**
 * What a reader reads each CID of an ordering as, for a font with no ToUnicode:
 * its `Adobe-<ordering>-UCS2` CMap (pdf.js buildToUnicode). The `Uni*-UCS2` CMaps
 * map characters to CIDs and several share one (文 and the Kangxi radical ⽂), so
 * inverted they cannot say which of them the reader reads.
 */
export async function loadMraiNativeOrderingReading(ordering: string): Promise<ReadonlyMap<number, string>> {
  if (!/^(?:GB1|CNS1|Japan1|Korea1)$/.test(ordering)) throw new Error(`Ordering ${ordering} has no reading.`);
  let pending = orderingReadings.get(ordering);
  if (!pending) {
    pending = (async () => {
      const data = await (activeLoader ?? defaultLoader)(`Adobe-${ordering}-UCS2`);
      if (!data.length || data.length > MAX_BCMAP_BYTES) throw new Error("Predefined CMap size is unsupported.");
      const readings = new Map<number, string>();
      for (const { low, high, unit } of parseMraiBinaryCMap(data).cidReadings) {
        if (high < low || high - low > 0xffff) throw new Error("Predefined CMap range is invalid.");
        for (let cid = low; cid <= high; cid += 1) {
          const value = unit + (cid - low);
          if (value > 0 && value <= 0xffff && (value < 0xd800 || value > 0xdfff)) readings.set(cid, String.fromCharCode(value));
        }
      }
      if (readings.size === 0) throw new Error(`Ordering ${ordering} reads no CID.`);
      return readings;
    })();
    orderingReadings.set(ordering, pending);
    pending.catch(() => orderingReadings.delete(ordering));
  }
  return pending;
}

async function resolveCMap(name: string, depth: number): Promise<PdfNativePredefinedCMap> {
  if (depth > MAX_USE_CMAP_DEPTH) throw new Error("Predefined CMap usecmap chain is too deep.");
  if (!/^[A-Za-z0-9-]{1,64}$/.test(name)) throw new Error("Predefined CMap name is invalid.");
  const data = await (activeLoader ?? defaultLoader)(name);
  if (!data.length || data.length > MAX_BCMAP_BYTES) throw new Error("Predefined CMap size is unsupported.");
  const parsed = parseMraiBinaryCMap(data);
  const codeToCid = new Map<number, number>();
  if (parsed.useCMap) {
    if (!/^[A-Za-z0-9-]{1,64}$/.test(parsed.useCMap)) throw new Error("Predefined CMap usecmap name is invalid.");
    const parentData = await (activeLoader ?? defaultLoader)(parsed.useCMap);
    const parent = parseMraiBinaryCMap(parentData);
    if (parent.useCMap) throw new Error("Nested predefined CMap usecmap is unsupported.");
    assertTwoByteCodespace(parent.codespace, parsed.useCMap);
    fillCodeToCid(parent, codeToCid);
  }
  // A vertical CMap usually inherits its whole codespace from the horizontal
  // CMap it uses; it must not widen that codespace with its own ranges.
  if (parsed.codespace.length || !parsed.useCMap) assertTwoByteCodespace(parsed.codespace, name);
  fillCodeToCid(parsed, codeToCid);
  if (codeToCid.size === 0) throw new Error(`Predefined CMap ${name} maps no character codes.`);
  return Object.freeze({ name, vertical: parsed.vertical, codeToCid });
}

function assertTwoByteCodespace(
  codespace: readonly ParsedCodespaceRange[],
  name: string,
): void {
  if (!codespace.length || codespace.some((range) => range.byteLength !== 2)) {
    throw new Error(`Predefined CMap ${name} is not a uniform two-byte encoding.`);
  }
}

function fillCodeToCid(parsed: ParsedBinaryCMap, into: Map<number, number>): void {
  for (const range of parsed.cidRanges) {
    if (range.byteLength !== 2) throw new Error("Predefined CMap range width is unsupported.");
    for (let code = range.low; code <= range.high; code += 1) {
      into.set(code, range.cid + (code - range.low));
    }
  }
  for (const single of parsed.cidSingles) {
    if (single.byteLength !== 2) throw new Error("Predefined CMap entry width is unsupported.");
    into.set(single.code, single.cid);
  }
}

export interface ParsedCodespaceRange {
  readonly byteLength: number;
  readonly low: number;
  readonly high: number;
}

export interface ParsedBinaryCMap {
  readonly vertical: boolean;
  readonly useCMap: string | null;
  readonly codespace: readonly ParsedCodespaceRange[];
  readonly cidSingles: readonly Readonly<{ byteLength: number; code: number; cid: number }>[];
  readonly cidRanges: readonly Readonly<{ byteLength: number; low: number; high: number; cid: number }>[];
  /** What an `Adobe-<ordering>-UCS2` CMap reads a CID as: the first UTF-16 unit
   * of its destination, as pdf.js takes it. Empty for a code-to-CID CMap. */
  readonly cidReadings: readonly Readonly<{ low: number; high: number; unit: number }>[];
}

const MAX_NUM_SIZE = 16;
const MAX_ENCODED_NUM_SIZE = 19;

class BinaryCMapStream {
  private pos = 0;
  private readonly tmp = new Uint8Array(MAX_ENCODED_NUM_SIZE);

  constructor(private readonly data: Uint8Array) {}

  getByte(): number {
    return this.pos >= this.data.length ? -1 : this.data[this.pos++];
  }

  readNumber(): number {
    let n = 0;
    let last = false;
    do {
      const b = this.getByte();
      if (b < 0) throw new Error("Unexpected end of predefined CMap.");
      last = !(b & 0x80);
      n = ((n << 7) | (b & 0x7f)) >>> 0;
    } while (!last);
    return n;
  }

  readSigned(): number {
    const n = this.readNumber();
    return n & 1 ? ~(n >>> 1) : n >>> 1;
  }

  readHex(num: Uint8Array, size: number): void {
    const end = this.pos + size + 1;
    if (end > this.data.length) throw new Error("Unexpected end of predefined CMap.");
    num.set(this.data.subarray(this.pos, end));
    this.pos = end;
  }

  readHexNumber(num: Uint8Array, size: number): void {
    let last = false;
    const stack = this.tmp;
    let sp = 0;
    do {
      const b = this.getByte();
      if (b < 0) throw new Error("Unexpected end of predefined CMap.");
      if (sp >= stack.length) throw new Error("Predefined CMap number is too long.");
      last = !(b & 0x80);
      stack[sp++] = b & 0x7f;
    } while (!last);
    let i = size;
    let buffer = 0;
    let bufferSize = 0;
    while (i >= 0) {
      while (bufferSize < 8 && sp > 0) {
        buffer |= stack[--sp] << bufferSize;
        bufferSize += 7;
      }
      num[i] = buffer & 255;
      i -= 1;
      buffer >>= 8;
      bufferSize -= 8;
    }
  }

  readHexSigned(num: Uint8Array, size: number): void {
    this.readHexNumber(num, size);
    const sign = num[size] & 1 ? 255 : 0;
    let c = 0;
    for (let i = 0; i <= size; i += 1) {
      c = ((c & 1) << 8) | num[i];
      num[i] = (c >> 1) ^ sign;
    }
  }

  readString(): string {
    const length = this.readNumber();
    if (length > 256) throw new Error("Predefined CMap string is too long.");
    let out = "";
    for (let i = 0; i < length; i += 1) out += String.fromCharCode(this.readNumber());
    return out;
  }
}

function hexToInt(a: Uint8Array, size: number): number {
  let n = 0;
  for (let i = 0; i <= size; i += 1) n = ((n << 8) | a[i]) >>> 0;
  return n >>> 0;
}

function addHex(a: Uint8Array, b: Uint8Array, size: number): void {
  let c = 0;
  for (let i = size; i >= 0; i -= 1) {
    c += a[i] + b[i];
    a[i] = c & 255;
    c >>= 8;
  }
}

function incHex(a: Uint8Array, size: number): void {
  let c = 1;
  for (let i = size; i >= 0 && c > 0; i -= 1) {
    c += a[i];
    a[i] = c & 255;
    c >>= 8;
  }
}

/** Port of pdf.js BinaryCMapReader restricted to CID (not bf) records. */
export function parseMraiBinaryCMap(data: Uint8Array): ParsedBinaryCMap {
  const stream = new BinaryCMapStream(data);
  const header = stream.getByte();
  if (header < 0) throw new Error("Predefined CMap is empty.");
  const vertical = !!(header & 1);
  let useCMap: string | null = null;
  const codespace: ParsedCodespaceRange[] = [];
  const cidSingles: Array<{ byteLength: number; code: number; cid: number }> = [];
  const cidRanges: Array<{ byteLength: number; low: number; high: number; cid: number }> = [];
  const cidReadings: Array<{ low: number; high: number; unit: number }> = [];
  const firstUnit = (bytes: Uint8Array, size: number) => size === 0 ? bytes[0] : ((bytes[0] << 8) | bytes[1]) >>> 0;
  const start = new Uint8Array(MAX_NUM_SIZE);
  const end = new Uint8Array(MAX_NUM_SIZE);
  const char = new Uint8Array(MAX_NUM_SIZE);
  const charCode = new Uint8Array(MAX_NUM_SIZE);
  const tmp = new Uint8Array(MAX_NUM_SIZE);
  let b: number;
  let entries = 0;
  while ((b = stream.getByte()) >= 0) {
    const type = b >> 5;
    if (type === 7) {
      switch (b & 0x1f) {
        case 0:
          stream.readString();
          break;
        case 1:
          useCMap = stream.readString();
          break;
        default:
          break;
      }
      continue;
    }
    const sequence = !!(b & 0x10);
    const dataSize = b & 15;
    if (dataSize + 1 > MAX_NUM_SIZE) throw new Error("Predefined CMap data size is invalid.");
    const ucs2DataSize = 1;
    const subitemsCount = stream.readNumber();
    if ((entries += subitemsCount) > 1_000_000) throw new Error("Predefined CMap is too large.");
    switch (type) {
      case 0:
        stream.readHex(start, dataSize);
        stream.readHexNumber(end, dataSize);
        addHex(end, start, dataSize);
        codespace.push({ byteLength: dataSize + 1, low: hexToInt(start, dataSize), high: hexToInt(end, dataSize) });
        for (let i = 1; i < subitemsCount; i += 1) {
          incHex(end, dataSize);
          stream.readHexNumber(start, dataSize);
          addHex(start, end, dataSize);
          stream.readHexNumber(end, dataSize);
          addHex(end, start, dataSize);
          codespace.push({ byteLength: dataSize + 1, low: hexToInt(start, dataSize), high: hexToInt(end, dataSize) });
        }
        break;
      case 1:
        // notdef ranges carry no editable mapping; consume and ignore them.
        stream.readHex(start, dataSize);
        stream.readHexNumber(end, dataSize);
        addHex(end, start, dataSize);
        stream.readNumber();
        for (let i = 1; i < subitemsCount; i += 1) {
          incHex(end, dataSize);
          stream.readHexNumber(start, dataSize);
          addHex(start, end, dataSize);
          stream.readHexNumber(end, dataSize);
          addHex(end, start, dataSize);
          stream.readNumber();
        }
        break;
      case 2: {
        stream.readHex(char, dataSize);
        let code = stream.readNumber();
        cidSingles.push({ byteLength: dataSize + 1, code: hexToInt(char, dataSize), cid: code });
        for (let i = 1; i < subitemsCount; i += 1) {
          incHex(char, dataSize);
          if (!sequence) {
            stream.readHexNumber(tmp, dataSize);
            addHex(char, tmp, dataSize);
          }
          code = stream.readSigned() + (code + 1);
          cidSingles.push({ byteLength: dataSize + 1, code: hexToInt(char, dataSize), cid: code });
        }
        break;
      }
      case 3: {
        stream.readHex(start, dataSize);
        stream.readHexNumber(end, dataSize);
        addHex(end, start, dataSize);
        let code = stream.readNumber();
        cidRanges.push({ byteLength: dataSize + 1, low: hexToInt(start, dataSize), high: hexToInt(end, dataSize), cid: code });
        for (let i = 1; i < subitemsCount; i += 1) {
          incHex(end, dataSize);
          if (!sequence) {
            stream.readHexNumber(start, dataSize);
            addHex(start, end, dataSize);
          } else {
            start.set(end);
          }
          stream.readHexNumber(end, dataSize);
          addHex(end, start, dataSize);
          code = stream.readNumber();
          cidRanges.push({ byteLength: dataSize + 1, low: hexToInt(start, dataSize), high: hexToInt(end, dataSize), cid: code });
        }
        break;
      }
      case 4:
        // bfchar records only occur in ToUnicode-style CMaps; they never become
        // CID authority, only what a reader reads a CID as.
        stream.readHex(char, ucs2DataSize);
        stream.readHex(charCode, dataSize);
        cidReadings.push({ low: hexToInt(char, ucs2DataSize), high: hexToInt(char, ucs2DataSize), unit: firstUnit(charCode, dataSize) });
        for (let i = 1; i < subitemsCount; i += 1) {
          incHex(char, ucs2DataSize);
          if (!sequence) {
            stream.readHexNumber(tmp, ucs2DataSize);
            addHex(char, tmp, ucs2DataSize);
          }
          incHex(charCode, dataSize);
          stream.readHexSigned(tmp, dataSize);
          addHex(charCode, tmp, dataSize);
          cidReadings.push({ low: hexToInt(char, ucs2DataSize), high: hexToInt(char, ucs2DataSize), unit: firstUnit(charCode, dataSize) });
        }
        break;
      case 5:
        stream.readHex(start, ucs2DataSize);
        stream.readHexNumber(end, ucs2DataSize);
        addHex(end, start, ucs2DataSize);
        stream.readHex(charCode, dataSize);
        cidReadings.push({ low: hexToInt(start, ucs2DataSize), high: hexToInt(end, ucs2DataSize), unit: firstUnit(charCode, dataSize) });
        for (let i = 1; i < subitemsCount; i += 1) {
          incHex(end, ucs2DataSize);
          if (!sequence) {
            stream.readHexNumber(start, ucs2DataSize);
            addHex(start, end, ucs2DataSize);
          } else {
            start.set(end);
          }
          stream.readHexNumber(end, ucs2DataSize);
          addHex(end, start, ucs2DataSize);
          stream.readHex(charCode, dataSize);
          cidReadings.push({ low: hexToInt(start, ucs2DataSize), high: hexToInt(end, ucs2DataSize), unit: firstUnit(charCode, dataSize) });
        }
        break;
      default:
        throw new Error(`Predefined CMap record type ${type} is unknown.`);
    }
  }
  for (const range of cidRanges) {
    if (range.high < range.low || range.high - range.low > 0xffff) {
      throw new Error("Predefined CMap range is invalid.");
    }
  }
  return Object.freeze({ vertical, useCMap, codespace, cidSingles, cidRanges, cidReadings });
}
