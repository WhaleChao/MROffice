/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  PDFArray,
  PDFDict,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFObjectParser,
  PDFRawStream,
  PDFRef,
  PDFString,
  decodePDFRawStream,
  type PDFDocument,
  type PDFPage,
} from "pdf-lib";

/** Largest content stream inspected per page (bytes). */
const MAX_CONTENT_BYTES = 32 * 1024 * 1024;



const MAX_ACTUAL_TEXT = 1_000_000;
const WHITESPACE = new Set([0x00, 0x09, 0x0a, 0x0c, 0x0d, 0x20]);
const DELIMITERS = new Set([
  0x28, 0x29, 0x3c, 0x3e, 0x5b, 0x5d, 0x7b, 0x7d, 0x2f, 0x25,
]);

type Token =
  | { kind: "name"; value: string }
  | { kind: "dict"; start: number; end: number }
  | { kind: "op"; value: string }
  | { kind: "other" };

function pageContentBytes(doc: PDFDocument, page: PDFPage): Uint8Array | null {
  const contents = page.node.Contents();
  if (!contents) return null;
  const streams: PDFRawStream[] = [];
  const push = (ref: unknown) => {
    const stream = ref instanceof PDFRef ? doc.context.lookup(ref) : ref;
    if (stream instanceof PDFRawStream) streams.push(stream);
  };
  if (contents instanceof PDFArray) {
    for (let i = 0; i < contents.size(); i += 1) push(contents.get(i));
  } else {
    push(contents);
  }
  const parts: Uint8Array[] = [];
  let total = 0;
  for (const stream of streams) {
    const bytes = decodePDFRawStream(stream).decode();
    total += bytes.length + 1;
    if (total > MAX_CONTENT_BYTES) return null;
    parts.push(bytes);
  }
  if (parts.length === 0) return null;
  const out = new Uint8Array(total);
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
    out[offset] = 0x0a;
    offset += 1;
  }
  return out;
}

/** Tokenize a content stream far enough to find `BDC` and its two operands. */
function* tokenize(bytes: Uint8Array): Generator<Token> {
  const n = bytes.length;
  let i = 0;
  const isRegular = (c: number) => !WHITESPACE.has(c) && !DELIMITERS.has(c);
  while (i < n) {
    const c = bytes[i]!;
    if (WHITESPACE.has(c)) {
      i += 1;
      continue;
    }
    if (c === 0x25) {
      // % comment to end of line
      while (i < n && bytes[i] !== 0x0a && bytes[i] !== 0x0d) i += 1;
      continue;
    }
    if (c === 0x28) {
      // ( literal string ) with nesting and escapes
      let depth = 0;
      while (i < n) {
        const d = bytes[i]!;
        if (d === 0x5c) {
          i += 2;
          continue;
        }
        if (d === 0x28) depth += 1;
        else if (d === 0x29) {
          depth -= 1;
          if (depth === 0) {
            i += 1;
            break;
          }
        }
        i += 1;
      }
      yield { kind: "other" };
      continue;
    }
    if (c === 0x3c) {
      if (bytes[i + 1] === 0x3c) {
        // << dictionary >> with nesting; strings inside are skipped
        const start = i;
        let depth = 0;
        while (i < n) {
          const d = bytes[i]!;
          if (d === 0x3c && bytes[i + 1] === 0x3c) {
            depth += 1;
            i += 2;
            continue;
          }
          if (d === 0x3e && bytes[i + 1] === 0x3e) {
            depth -= 1;
            i += 2;
            if (depth === 0) break;
            continue;
          }
          if (d === 0x28) {
            let inner = 0;
            while (i < n) {
              const e = bytes[i]!;
              if (e === 0x5c) {
                i += 2;
                continue;
              }
              if (e === 0x28) inner += 1;
              else if (e === 0x29) {
                inner -= 1;
                if (inner === 0) {
                  i += 1;
                  break;
                }
              }
              i += 1;
            }
            continue;
          }
          i += 1;
        }
        yield { kind: "dict", start, end: i };
        continue;
      }
      // < hex string >
      while (i < n && bytes[i] !== 0x3e) i += 1;
      i += 1;
      yield { kind: "other" };
      continue;
    }
    if (c === 0x2f) {
      let j = i + 1;
      while (j < n && isRegular(bytes[j]!)) j += 1;
      yield { kind: "name", value: decodeName(bytes.subarray(i + 1, j)) };
      i = j;
      continue;
    }
    if (DELIMITERS.has(c)) {
      // [ ] { } ) > on their own
      i += 1;
      yield { kind: "other" };
      continue;
    }
    let j = i;
    while (j < n && isRegular(bytes[j]!)) j += 1;
    const word = String.fromCharCode(...bytes.subarray(i, j));
    i = j;
    if (/^[+-.\d]/.test(word)) {
      yield { kind: "other" };
      continue;
    }
    if (word === "ID") {
      // Inline image data runs until whitespace EI whitespace.
      while (i < n) {
        if (
          bytes[i] === 0x45 &&
          bytes[i + 1] === 0x49 &&
          WHITESPACE.has(bytes[i - 1] ?? 0x20) &&
          (i + 2 >= n || WHITESPACE.has(bytes[i + 2]!))
        ) {
          i += 2;
          break;
        }
        i += 1;
      }
      yield { kind: "op", value: "EI" };
      continue;
    }
    yield { kind: "op", value: word };
  }
}

function decodeName(bytes: Uint8Array): string {
  let out = "";
  for (let i = 0; i < bytes.length; i += 1) {
    const c = bytes[i]!;
    if (c === 0x23 && i + 2 < bytes.length) {
      const hex = String.fromCharCode(bytes[i + 1]!, bytes[i + 2]!);
      if (/^[0-9a-fA-F]{2}$/.test(hex)) {
        out += String.fromCharCode(parseInt(hex, 16));
        i += 2;
        continue;
      }
    }
    out += String.fromCharCode(c);
  }
  return out;
}

function readActualTextEntry(dict: PDFDict): [number, string] | null {
  const mcid = dict.get(PDFName.of("MCID"));
  const actual = dict.get(PDFName.of("ActualText"));
  if (!(mcid instanceof PDFNumber)) return null;
  if (!(actual instanceof PDFString || actual instanceof PDFHexString)) {
    return null;
  }
  const id = mcid.asNumber();
  if (!Number.isInteger(id) || id < 0) return null;
  let text: string;
  try {
    text = decodeActualTextString(actual);
  } catch {
    return null;
  }
  if (text.length > MAX_ACTUAL_TEXT) return null;
  return [id, text];
}

const UTF8 = new TextDecoder("utf-8", { fatal: true });

/**
 * A text string as a reader reads it: UTF-16BE behind its BOM, UTF-8 behind
 * its BOM (PDF 2.0), otherwise PDFDocEncoding -- unless its bytes are UTF-8
 * written without the BOM, which producers do: `A\xc2\xa0B` read byte by
 * byte came out as "AÂ€B", a no-break space turned into two letters in the
 * extracted text (user audit PDF 0040). Bytes above 0x7F that make valid
 * UTF-8 are read as UTF-8 (natural PDFDocEncoding text, "café" as 63 61 66
 * E9, is not valid UTF-8 and keeps its meaning).
 */
export function decodeActualTextString(value: PDFString | PDFHexString): string {
  const bytes = value.asBytes();
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return UTF8.decode(bytes.subarray(3));
  }
  const utf16 =
    (bytes[0] === 0xfe && bytes[1] === 0xff) || (bytes[0] === 0xff && bytes[1] === 0xfe);
  if (!utf16 && bytes.some((byte) => byte > 0x7f)) {
    try {
      return UTF8.decode(bytes);
    } catch {
      // Not UTF-8: PDFDocEncoding, as written.
    }
  }
  return value.decodeText();
}

/** MCID → ActualText for one page (0-based index). Never throws. */
export function readPageActualTextByMcid(
  doc: PDFDocument,
  pageIndex: number,
): Map<number, string> {
  const out = new Map<number, string>();
  try {
    const page = doc.getPages()[pageIndex];
    if (!page) return out;
    const bytes = pageContentBytes(doc, page);
    if (!bytes) return out;
    let properties: PDFDict | undefined;
    const namedProperties = (name: string): PDFDict | undefined => {
      if (properties === undefined) {
        properties =
          page.node.Resources()?.lookupMaybe(PDFName.of("Properties"), PDFDict) ??
          undefined;
      }
      return properties?.lookupMaybe(PDFName.of(name), PDFDict);
    };
    const operands: Token[] = [];
    for (const token of tokenize(bytes)) {
      if (token.kind !== "op") {
        operands.push(token);
        if (operands.length > 64) operands.shift();
        continue;
      }
      if (token.value === "BDC" && operands.length >= 2) {
        const props = operands[operands.length - 1]!;
        let dict: PDFDict | undefined;
        if (props.kind === "dict") {
          const parsed = PDFObjectParser.forBytes(
            bytes.subarray(props.start, props.end),
            doc.context,
          ).parseObject();
          if (parsed instanceof PDFDict) dict = parsed;
        } else if (props.kind === "name") {
          dict = namedProperties(props.value);
        }
        if (dict) {
          const entry = readActualTextEntry(dict);
          if (entry && !out.has(entry[0])) out.set(entry[0], entry[1]);
        }
      }
      operands.length = 0;
    }
  } catch {
    /* an unreadable content stream substitutes nothing */
  }
  return out;
}

/**
 * Read every page's ActualText map from raw PDF bytes: 1-based page number →
 * (MCID → text), pages without any entry absent. Empty for anything
 * unreadable.
 */
export async function readActualTextFromBytes(
  bytes: ArrayBuffer | Uint8Array,
): Promise<Map<number, Map<number, string>>> {
  const out = new Map<number, Map<number, string>>();
  try {
    const { loadPdfForInspection } = await import("./encrypted-guard");
    const doc = await loadPdfForInspection(bytes, { updateMetadata: false });
    if (!doc) return out;
    const count = doc.getPageCount();
    for (let index = 0; index < count; index += 1) {
      const page = readPageActualTextByMcid(doc, index);
      if (page.size > 0) out.set(index + 1, page);
    }
  } catch {
    return out;
  }
  return out;
}

/** The MCID inside a pdf.js marked-content id (`p0_mc12` → 12), or null. */
export function mcidFromPdfJsContentId(id: string): number | null {
  const match = /_mc(\d+)$/.exec(id);
  if (!match) return null;
  const value = Number(match[1]);
  return Number.isSafeInteger(value) ? value : null;
}
