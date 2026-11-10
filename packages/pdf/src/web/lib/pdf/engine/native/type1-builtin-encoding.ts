/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** A Type 1 font program's own encoding, for symbolic fonts that carry no
 * ToUnicode and no base encoding.
 *
 * Such a font's codes mean whatever its program says they mean, and PDF 32000-1
 * 9.10.2 names exactly this fallback: read the built-in encoding, then resolve
 * each glyph name through the Adobe Glyph List. The encoding lives in the
 * cleartext portion of the program, before `eexec`, as an ordinary PostScript
 * array of `dup <code> /<name> put` entries -- nothing has to be decrypted to
 * read it. A name the list does not know (TeX's `summationdisplay` and the
 * other math extensions) simply leaves that one code unmapped, which costs
 * those characters and not the page.
 */

import { readCffBuiltinEncodingNames } from "./cff-builtin-encoding";
import { readTrueTypeSymbolicBuiltinEncoding } from "./truetype-builtin-encoding";
import {
  PDF_SIMPLE_FONT_BASE_ENCODINGS,
  pdfSimpleFontGlyphNameToCharacter,
} from "./simple-font-encodings";

const MAX_CLEARTEXT_BYTES = 1 << 20;
const MAX_ENCODING_ENTRIES = 256;
const ENTRY = /dup\s+(\d{1,3})\s*\/([A-Za-z0-9._]{1,127})\s+put/g;

/** Codes a Type 1 program's own encoding array assigns, as glyph names. */
export function readType1BuiltinEncodingNames(
  program: Uint8Array,
): ReadonlyMap<number, string> | "standard" | null {
  if (!program.length) return null;
  const limit = Math.min(program.length, MAX_CLEARTEXT_BYTES);
  let cleartext = "";
  for (let offset = 0; offset < limit; offset += 0x8000) {
    cleartext += String.fromCharCode(
      ...program.subarray(offset, Math.min(offset + 0x8000, limit)),
    );
  }
  // `eexec` begins the encrypted portion; the encoding is always before it.
  const encrypted = cleartext.indexOf("eexec");
  const header = encrypted >= 0 ? cleartext.slice(0, encrypted) : cleartext;
  if (!/\/Encoding\b/.test(header)) return null;
  // Naming the standard encoding is as definite as spelling the codes out.
  if (/\/Encoding\s+StandardEncoding\s+def/.test(header)) return "standard";
  const names = new Map<number, string>();
  ENTRY.lastIndex = 0;
  for (let match = ENTRY.exec(header); match; match = ENTRY.exec(header)) {
    const code = Number(match[1]);
    if (code > 255) continue;
    if (names.size >= MAX_ENCODING_ENTRIES) return null;
    // A program that assigns one code twice does not say what it means.
    if (names.has(code) && names.get(code) !== match[2]) return null;
    if (match[2] !== ".notdef") names.set(code, match[2]);
  }
  // An array that leaves every code `.notdef` is still the program's
  // answer: the PDF's own /Differences then name the codes it uses.
  return names;
}

/** The code table a symbolic font's own embedded program defines.
 *
 * A Type 1 program states its encoding as PostScript text, a bare CFF one
 * states it as binary tables and a TrueType one through its Mac Roman `cmap`
 * subtable; all answer the same question, so the caller does not have to know
 * which kind it holds. */
const builtinEncodings = new WeakMap<
  Uint8Array,
  Promise<ReadonlyMap<number, string> | null>
>();

export function readType1BuiltinEncoding(
  program: Uint8Array,
): Promise<ReadonlyMap<number, string> | null> {
  // Every candidate route asks the same font the same question; reading a
  // program once per proof instead of once per attempt keeps a page with many
  // fonts from paying for it over and over.
  const cached = builtinEncodings.get(program);
  if (cached) return cached;
  const pending = decodeBuiltinEncoding(program);
  builtinEncodings.set(program, pending);
  return pending;
}

async function decodeBuiltinEncoding(
  program: Uint8Array,
): Promise<ReadonlyMap<number, string> | null> {
  const trueType = readTrueTypeSymbolicBuiltinEncoding(program);
  if (trueType) return trueType;
  const names =
    readCffBuiltinEncodingNames(program) ?? readType1BuiltinEncodingNames(program);
  if (!names) return null;
  if (names === "standard") {
    return new Map(
      Object.entries(PDF_SIMPLE_FONT_BASE_ENCODINGS.StandardEncoding).map(
        ([code, character]) => [Number(code), character],
      ),
    );
  }
  // The names pdf.js reads: the Adobe Glyph List and its TeX names (a
  // Computer Modern extension face calls its display integral
  // `integraldisplay`).
  const { PDFJS_GLYPH_LIST } = await import("./adobe-glyph-list");
  const decode = new Map<number, string>();
  for (const [code, name] of names) {
    const character = pdfSimpleFontGlyphNameToCharacter(name, PDFJS_GLYPH_LIST);
    if (character !== null) decode.set(code, character);
  }
  // A program whose names the glyph list does not know still has an
  // encoding; the codes it leaves unmapped cost those characters, not the
  // font (a /Differences array may still name them).
  return decode;
}
