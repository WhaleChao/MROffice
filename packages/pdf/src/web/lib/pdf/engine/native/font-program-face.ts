/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The family, weight and slant an embedded font program names itself by:
 * its sfnt `name` table (the typographic family, else the family) and its
 * `OS/2` weight class and italic bit. A subset's PostScript name is often a
 * tag and a hash (ReportLab's per-character subsets of an AI poster are
 * "MRTBOffice-<slice>-700"), while the program still says "Noto Serif SC" at
 * weight 700: what a writer needs to set the same text again in a served
 * face of that family. Read-only, for the editor's drafts; never a write
 * authority.
 */

export interface PdfFontProgramFace {
  readonly family: string;
  readonly weight: number;
  readonly italic: boolean;
}

const STYLE_WORDS =
  /(?:\s+|[-_])(?:thin|hairline|extra\s*light|ultra\s*light|extralight|ultralight|light|book|regular|normal|roman|medium|semi\s*bold|demi\s*bold|semibold|demibold|bold|extra\s*bold|ultra\s*bold|extrabold|ultrabold|black|heavy|italic|oblique)$/i;

/** A family name without the style words a static instance appends. */
function familyOf(name: string): string {
  let family = name.trim();
  for (let guard = 0; guard < 4; guard += 1) {
    const next = family.replace(STYLE_WORDS, "").trim();
    if (next === family || !next) break;
    family = next;
  }
  return family;
}

function utf16be(bytes: Uint8Array, start: number, length: number): string {
  let text = "";
  for (let at = start; at + 1 < start + length; at += 2) text += String.fromCharCode((bytes[at] << 8) | bytes[at + 1]);
  return text;
}

export function readPdfFontProgramFace(program: Uint8Array): PdfFontProgramFace | null {
  try {
    if (program.length < 12) return null;
    const view = new DataView(program.buffer, program.byteOffset, program.byteLength);
    const version = view.getUint32(0);
    if (version !== 0x00010000 && version !== 0x74727565 && version !== 0x4f54544f) return null;
    const count = view.getUint16(4);
    if (!count || count > 4096 || 12 + count * 16 > program.length) return null;
    const tables = new Map<string, { offset: number; length: number }>();
    for (let index = 0; index < count; index += 1) {
      const entry = 12 + index * 16;
      const tag = String.fromCharCode(program[entry], program[entry + 1], program[entry + 2], program[entry + 3]);
      const offset = view.getUint32(entry + 8), length = view.getUint32(entry + 12);
      if (offset + length <= program.length) tables.set(tag, { offset, length });
    }
    const name = tables.get("name");
    if (!name || name.length < 6) return null;
    const records = view.getUint16(name.offset + 2);
    const strings = name.offset + view.getUint16(name.offset + 4);
    const found = new Map<string, string>();
    for (let index = 0; index < records; index += 1) {
      const record = name.offset + 6 + index * 12;
      if (record + 12 > name.offset + name.length) break;
      const platform = view.getUint16(record), encoding = view.getUint16(record + 2);
      const language = view.getUint16(record + 4), nameId = view.getUint16(record + 6);
      const length = view.getUint16(record + 8), offset = strings + view.getUint16(record + 10);
      if ((nameId !== 1 && nameId !== 16) || offset + length > program.length) continue;
      // Windows Unicode names, the English one first; else the Mac Roman one.
      if (platform === 3 && (encoding === 1 || encoding === 10)) {
        if (language === 0x409 || !found.has(`w${nameId}`)) found.set(`w${nameId}`, utf16be(program, offset, length));
      } else if (platform === 1 && encoding === 0 && !found.has(`m${nameId}`)) {
        found.set(`m${nameId}`, String.fromCharCode(...program.subarray(offset, offset + length)));
      }
    }
    const raw = found.get("w16") ?? found.get("w1") ?? found.get("m16") ?? found.get("m1");
    const family = raw ? familyOf(raw) : "";
    if (!family || family.length > 96) return null;
    const os2 = tables.get("OS/2");
    const weight = os2 && os2.length >= 6 ? view.getUint16(os2.offset + 4) : 400;
    const italic = os2 && os2.length >= 64 ? (view.getUint16(os2.offset + 62) & 1) === 1 : false;
    return Object.freeze({ family, weight: weight >= 1 && weight <= 1000 ? weight : 400, italic });
  } catch {
    return null;
  }
}
