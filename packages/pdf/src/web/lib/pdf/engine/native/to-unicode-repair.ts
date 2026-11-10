/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pdfTrueTypeGlyphName, readPdfTrueTypeGlyphCharacters } from "./cid-glyph-map";
import { LATIN1_EXACT } from "./latin1";
import { readMraiNativeToUnicodeReading } from "./list-text-fidelity";

export interface PdfToUnicodeRepair {
  /** The CMap written again, as the stream's decoded bytes. */
  readonly stream: Uint8Array;
  /** Each code named again, with the character it now reads as. */
  readonly repaired: ReadonlyMap<number, string>;
  /** The codes still called white space that draw (the program names them
   * nothing a reader shows). */
  readonly left: readonly number[];
}

/** The most entries one `beginbfchar` / `beginbfrange` block lists (the CMap
 * operators take up to 100 each). */
const BLOCK_ENTRIES = 100;

const hex4 = (value: number) => value.toString(16).toUpperCase().padStart(4, "0");

/** A string as the UTF-16BE hex a ToUnicode destination is written in. */
function utf16Hex(text: string): string {
  let out = "";
  for (let index = 0; index < text.length; index += 1) out += hex4(text.charCodeAt(index));
  return out;
}

const sameCharacter = (left: string, right: string) =>
  left === right || left.normalize("NFKC") === right.normalize("NFKC");

/**
 * The CMap `source` with each code of `repairs` named its character, every
 * other entry as written. Null where an entry of a block that holds one of
 * the codes is not one this reads (the stream is then left alone).
 */
function rewriteToUnicode(source: string, repairs: ReadonlyMap<number, string>): string | null {
  let failed = false;
  const blocks = /(\d+)(\s+)begin(bfchar|bfrange)([\s\S]*?)end\3/g;
  const rewritten = source.replace(blocks, (whole, _count: string, _gap: string, kind: string, body: string) => {
    if (failed) return whole;
    const entries: string[] = [];
    let touched = false;
    if (kind === "bfchar") {
      const entry = /<([0-9a-f]{4})>(\s*)<((?:\s*[0-9a-f]{4}){0,16}\s*)>/gi;
      if (body.replace(entry, "").trim()) { failed = true; return whole; }
      for (const match of body.matchAll(entry)) {
        const repair = repairs.get(Number.parseInt(match[1], 16));
        if (repair === undefined) { entries.push(match[0]); continue; }
        touched = true;
        entries.push(`<${match[1]}> <${utf16Hex(repair)}>`);
      }
    } else {
      const entry =
        /<([0-9a-f]{4})>\s*<([0-9a-f]{4})>\s*(?:\[((?:\s*<(?:\s*[0-9a-f]{4}){0,16}\s*>)+)\s*\]|<((?:\s*[0-9a-f]{4}){1,16}\s*)>)/gi;
      if (body.replace(entry, "").trim()) { failed = true; return whole; }
      for (const match of body.matchAll(entry)) {
        const start = Number.parseInt(match[1], 16), end = Number.parseInt(match[2], 16);
        if (end < start) { failed = true; return whole; }
        const inside: number[] = [];
        for (const code of repairs.keys()) if (code >= start && code <= end) inside.push(code);
        if (!inside.length) { entries.push(match[0]); continue; }
        touched = true;
        inside.sort((a, b) => a - b);
        if (match[3] !== undefined) {
          // Each code's destination named in turn: the named ones replaced.
          const destinations = [...match[3].matchAll(/<((?:\s*[0-9a-f]{4}){0,16}\s*)>/gi)].map((item) => item[1].replace(/\s+/g, ""));
          if (destinations.length !== end - start + 1) { failed = true; return whole; }
          for (const code of inside) destinations[code - start] = utf16Hex(repairs.get(code)!);
          entries.push(`<${match[1]}> <${match[2]}> [${destinations.map((destination) => `<${destination}>`).join(" ")}]`);
          continue;
        }
        // One destination counted up along the range: the range is written
        // as the stretches between the named codes, each counted from where
        // the range had it, and each named code by itself.
        const destination = match[4].replace(/\s+/g, "");
        const first = Number.parseInt(destination, 16);
        // (A destination past one UTF-16 unit is counted by readers in ways
        // that differ past its last unit's end; such a range is left.)
        if (destination.length > 8 || !Number.isSafeInteger(first) ||
          (destination.length > 4 && ((first & 0xffff) + (end - start)) > 0xffff) ||
          (destination.length === 4 && first + (end - start) > 0xffff)) { failed = true; return whole; }
        const at = (code: number) => (first + code - start).toString(16).toUpperCase().padStart(destination.length, "0");
        let from = start;
        for (const code of inside) {
          if (code > from) entries.push(`<${hex4(from)}> <${hex4(code - 1)}> <${at(from)}>`);
          entries.push(`<${hex4(code)}> <${hex4(code)}> <${utf16Hex(repairs.get(code)!)}>`);
          from = code + 1;
        }
        if (from <= end) entries.push(`<${hex4(from)}> <${hex4(end)}> <${at(from)}>`);
      }
    }
    if (!touched) return whole;
    const out: string[] = [];
    for (let index = 0; index < entries.length; index += BLOCK_ENTRIES) {
      const part = entries.slice(index, index + BLOCK_ENTRIES);
      out.push(`${part.length} begin${kind}\n${part.join("\n")}\nend${kind}`);
    }
    return out.join("\n");
  });
  return failed ? null : rewritten;
}

/**
 * The repair of one Identity-encoded CIDFontType2's ToUnicode (Identity
 * CIDToGIDMap: the code is the glyph id), from its decoded stream and its
 * TrueType program. Null where nothing is to be named: the table calls no
 * drawn glyph a space, the program's names are not shown to be the table's,
 * or it names none of those glyphs.
 */
export function planPdfToUnicodeRepair(toUnicodeBytes: Uint8Array, trueTypeProgram: Uint8Array): PdfToUnicodeRepair | null {
  const reading = readMraiNativeToUnicodeReading(toUnicodeBytes, trueTypeProgram);
  if (!reading || reading.inkedBlankCodes.size === 0) return null;
  const names = readPdfTrueTypeGlyphCharacters(trueTypeProgram);
  if (!names) return null;
  // The program's names stand for this table only where the two agree on
  // every glyph both name, and on one at least.
  let agreed = 0;
  for (const [code, text] of reading.decode) {
    if (reading.inkedBlankCodes.has(code) || [...text].length !== 1 || /^[\s\p{C}]$/u.test(text)) continue;
    const points = names.get(code);
    if (!points?.length) continue;
    if (!points.some((point) => sameCharacter(String.fromCodePoint(point), text))) return null;
    agreed += 1;
  }
  if (agreed === 0) return null;
  const repaired = new Map<number, string>();
  for (const code of reading.inkedBlankCodes) {
    const name = pdfTrueTypeGlyphName(names.get(code));
    if (name !== null) repaired.set(code, name);
  }
  if (repaired.size === 0) return null;
  const source = LATIN1_EXACT.decode(toUnicodeBytes);
  const rewritten = rewriteToUnicode(source, repaired);
  if (rewritten === null || rewritten === source) return null;
  const stream = new Uint8Array(rewritten.length);
  for (let index = 0; index < rewritten.length; index += 1) {
    const unit = rewritten.charCodeAt(index);
    if (unit > 0xff) return null;
    stream[index] = unit;
  }
  // Read again the way the font will be read: the named codes read as
  // their characters, the others as they did, and only the codes the
  // program names nothing for still draw as blanks.
  const check = readMraiNativeToUnicodeReading(stream, trueTypeProgram);
  if (!check || check.decode.size !== reading.decode.size) return null;
  for (const [code, text] of reading.decode) {
    if (check.decode.get(code) !== (repaired.get(code) ?? text)) return null;
  }
  const left = [...reading.inkedBlankCodes].filter((code) => !repaired.has(code)).sort((a, b) => a - b);
  if (check.inkedBlankCodes.size !== left.length || left.some((code) => !check.inkedBlankCodes.has(code))) return null;
  return Object.freeze({ stream, repaired, left: Object.freeze(left) });
}
