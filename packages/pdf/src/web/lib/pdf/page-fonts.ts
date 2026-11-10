/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Fonts from a reader's own computer that a PDF page keeps (WPS sets PDF
 * text in any installed font and embeds it).
 *
 * The browser reads the face (Local Font Access, page-fonts-browser.ts), the
 * page keeps the face the server prepared from it as a font attachment
 * (app/services/pdf_page_fonts.py: the PDF licence rule, TrueType outlines),
 * and a write names that program by digest: a native run's embedded-font
 * key in the family (native-embedded-fonts.ts installs it as a served face
 * is installed, its package marked `page`), a text-box run's `pageFont`.
 * The server proves each against the page's copy; a page opened on another
 * machine lists the same fonts (GET /pages/{id}/pdf-fonts) and draws and
 * writes in them without the font installed.
 *
 * Pure: no browser API, so the engine host can plan with it.
 */

export interface PdfPageFont {
  /** The family as the server named the prepared face (its name table). */
  readonly family: string;
  /** 700 for a bold face, else 400 (the face's role). */
  readonly weight: number;
  readonly italic: boolean;
  /** The prepared program's digest and size: what writes name and prove. */
  readonly sha256: string;
  readonly size: number;
  readonly attachmentId: string;
}

import { tryClassifyFontName } from "./font-resolve";

const SHA256 = /^[0-9a-f]{64}$/u;
const FAMILY_TOKEN = /^(.{1,60})#p=([0-9a-f]{64})$/u;

/**
 * The name a page font goes by wherever the editors carry a family (a run's
 * Office family, a native draft's embedded-font key, the font list's
 * value): its family and the digest of the one face it is, so every write
 * and every reopening names exactly that program. Shown as the family.
 */
export function pdfPageFontFamilyToken(font: Pick<PdfPageFont, "family" | "sha256">): string {
  return `${font.family}#p=${font.sha256}`;
}

/** The family and face digest a page-font name stands for, else null. */
export function readPdfPageFontFamilyToken(
  value: string | null | undefined,
): Readonly<{ family: string; sha256: string }> | null {
  const match = typeof value === "string" ? FAMILY_TOKEN.exec(value) : null;
  return match && match[1].trim() === match[1] ? Object.freeze({ family: match[1], sha256: match[2] }) : null;
}

/** The family a font name shows as: a page font's family, else the name. */
export function pdfPageFontDisplayFamily(value: string): string {
  // (xelatex names a face with its CMap: FandolSong-Bold-Identity-H.)
  return readPdfPageFontFamilyToken(value)?.family ?? value.replace(/-Identity-[HV]$/u, "");
}
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/iu;

/** A page font as the server lists it, or null when any field is off. */
export function readPdfPageFont(value: unknown): PdfPageFont | null {
  if (!value || typeof value !== "object") return null;
  const font = value as Record<string, unknown>;
  if (
    typeof font.family !== "string" || !font.family || font.family.length > 200 ||
    (font.weight !== 400 && font.weight !== 700) ||
    typeof font.italic !== "boolean" ||
    typeof font.sha256 !== "string" || !SHA256.test(font.sha256) ||
    typeof font.size !== "number" || !Number.isSafeInteger(font.size) || font.size <= 0 ||
    typeof font.attachmentId !== "string" || !UUID.test(font.attachmentId)
  ) {
    return null;
  }
  return Object.freeze({
    family: font.family, weight: font.weight, italic: font.italic,
    sha256: font.sha256, size: font.size, attachmentId: font.attachmentId,
  });
}

const sameFamily = (left: string, right: string) => left.trim().toLowerCase() === right.trim().toLowerCase();

/** The page's faces of one family (by its name, any case). */
export function pdfPageFontFaces(fonts: readonly PdfPageFont[], family: string): readonly PdfPageFont[] {
  return fonts.filter((font) => sameFamily(font.family, family));
}

/**
 * The face a style is set in among a family's page faces (CSS matching: the
 * slant asked, then the nearest weight), with what the face cannot draw
 * synthesized by the writer as for any face. Null when the page keeps no
 * face of the family.
 */
export function choosePdfPageFont(
  fonts: readonly PdfPageFont[],
  family: string,
  style: Readonly<{ bold: boolean; italic: boolean }>,
): Readonly<{ font: PdfPageFont; syntheticBold: boolean; syntheticItalic: boolean }> | null {
  // A page-font name is that one face.
  const token = readPdfPageFontFamilyToken(family);
  if (token) {
    const font = fonts.find((candidate) => candidate.sha256 === token.sha256);
    return font
      ? Object.freeze({ font, syntheticBold: style.bold && font.weight < 600, syntheticItalic: style.italic && !font.italic })
      : null;
  }
  const faces = pdfPageFontFaces(fonts, family);
  if (!faces.length) return null;
  const slanted = faces.filter((face) => face.italic === style.italic);
  const pool = slanted.length ? slanted : faces;
  const wanted = style.bold ? 700 : 400;
  const font = pool.find((face) => face.weight === wanted) ??
    [...pool].sort((left, right) => Math.abs(left.weight - wanted) - Math.abs(right.weight - wanted))[0];
  return Object.freeze({
    font,
    syntheticBold: style.bold && font.weight < 600,
    syntheticItalic: style.italic && !font.italic,
  });
}

/** The CSS family a page font is registered under in the document. */
export function pdfPageFontCssFamily(font: Pick<PdfPageFont, "sha256">): string {
  return `MRAIPageFont-${font.sha256.slice(0, 16)}`;
}

/**
 * One face of a font collection (`ttcf`) as a font file of its own: its
 * table directory and the tables it names, copied (offsets rebuilt, every
 * table on a 4-byte boundary). The face is the one whose PostScript name is
 * `postscriptName`; a file that is one face already is returned as is.
 * Chrome hands Local Font Access a collection whole (PingFang.ttc is 74 MB,
 * 24 faces); the server prepares one face.
 */
export function pdfFontCollectionFace(bytes: Uint8Array, postscriptName: string): Uint8Array | null {
  try {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    if (view.getUint32(0) !== 0x74746366) return bytes;
    const count = view.getUint32(8);
    if (count < 1 || count > 128 || 12 + count * 4 > bytes.length) return null;
    for (let index = 0; index < count; index += 1) {
      const offset = view.getUint32(12 + index * 4);
      const face = sfntFace(bytes, view, offset);
      if (face && face.postscriptName === postscriptName) return face.build();
    }
    return null;
  } catch {
    return null;
  }
}

function sfntFace(bytes: Uint8Array, view: DataView, offset: number) {
  if (offset + 12 > bytes.length) return null;
  const version = view.getUint32(offset);
  if (version !== 0x00010000 && version !== 0x4f54544f && version !== 0x74727565) return null;
  const count = view.getUint16(offset + 4);
  if (!count || count > 256 || offset + 12 + count * 16 > bytes.length) return null;
  const tables: { tag: number; checksum: number; offset: number; length: number }[] = [];
  for (let index = 0; index < count; index += 1) {
    const at = offset + 12 + index * 16;
    const table = {
      tag: view.getUint32(at), checksum: view.getUint32(at + 4),
      offset: view.getUint32(at + 8), length: view.getUint32(at + 12),
    };
    if (table.offset + table.length > bytes.length) return null;
    tables.push(table);
  }
  const name = tables.find((table) => table.tag === 0x6e616d65);
  const postscriptName = name ? sfntPostscriptName(bytes, view, name.offset, name.length) : null;
  return {
    postscriptName,
    build(): Uint8Array {
      const sorted = [...tables].sort((left, right) => left.tag - right.tag);
      let size = 12 + sorted.length * 16;
      const placed = sorted.map((table) => {
        const at = size;
        size += (table.length + 3) & ~3;
        return { ...table, at };
      });
      const out = new Uint8Array(size);
      const write = new DataView(out.buffer);
      let power = 1;
      while (power * 2 <= sorted.length) power *= 2;
      write.setUint32(0, version);
      write.setUint16(4, sorted.length);
      write.setUint16(6, power * 16);
      write.setUint16(8, Math.log2(power));
      write.setUint16(10, sorted.length * 16 - power * 16);
      placed.forEach((table, index) => {
        const at = 12 + index * 16;
        write.setUint32(at, table.tag);
        write.setUint32(at + 4, table.checksum);
        write.setUint32(at + 8, table.at);
        write.setUint32(at + 12, table.length);
        out.set(bytes.subarray(table.offset, table.offset + table.length), table.at);
      });
      return out;
    },
  };
}

function sfntPostscriptName(bytes: Uint8Array, view: DataView, offset: number, length: number): string | null {
  if (length < 6) return null;
  const records = view.getUint16(offset + 2);
  const strings = offset + view.getUint16(offset + 4);
  for (let index = 0; index < records; index += 1) {
    const at = offset + 6 + index * 12;
    if (at + 12 > offset + length) break;
    const platform = view.getUint16(at), nameId = view.getUint16(at + 6);
    const size = view.getUint16(at + 8), start = strings + view.getUint16(at + 10);
    if (nameId !== 6 || start + size > bytes.length) continue;
    if (platform === 3 || platform === 0) {
      let text = "";
      for (let byte = start; byte + 1 < start + size; byte += 2) text += String.fromCharCode((bytes[byte] << 8) | bytes[byte + 1]);
      return text;
    }
    if (platform === 1) return String.fromCharCode(...bytes.subarray(start, start + size));
  }
  return null;
}

const CJK_NAME = /[㐀-鿿豈-﫿]|pingfang|hiragino|heiti|songti|kaiti|fangsong|yahei|jhenghei|dengxian|simsun|simhei|mincho|gothic|ming|noto (?:sans|serif) (?:cjk|sc|tc|jp|kr)|source han|\b(?:sc|tc)\b/iu;

/**
 * The served family nearest a family the PDF cannot have (a browser that
 * cannot read the reader's fonts, a licence that forbids embedding): a
 * Chinese face of the same kind (楷 → 楷体, 仿宋 → 仿宋, 宋/明 → 宋体, else
 * 微软雅黑), a Latin one of the same class (Times New Roman, Courier New,
 * Arial).
 */
export function nearestServedPdfFamily(name: string): string {
  if (CJK_NAME.test(name)) {
    if (/kai|楷/iu.test(name)) return "KaiTi";
    if (/fang ?song|仿宋/iu.test(name)) return "FangSong";
    if (/song|ming|mincho|宋|明|serif/iu.test(name)) return "SimSun";
    return "Microsoft YaHei";
  }
  if (/mono|courier|consol|menlo|monaco|\bcode\b/iu.test(name)) return "Courier New";
  const kind = tryClassifyFontName(name);
  return kind === "serif" ? "Times New Roman" : kind === "mono" ? "Courier New" : "Arial";
}

/** Fonts a computer commonly has that MRAI does not serve: what Browse
 * looks for by drawing them where the browser cannot read installed fonts
 * (Safari). macOS's Chinese and Latin faces, and common Windows and Adobe
 * ones. */
export const PDF_SYSTEM_FAMILY_CANDIDATES: readonly string[] = Object.freeze([
  "PingFang SC", "PingFang TC", "Hiragino Sans GB", "Heiti SC", "STHeiti", "Songti SC", "STSong", "Kaiti SC",
  "STKaiti", "Baoli SC", "Libian SC", "Yuanti SC", "Wawati SC", "Xingkai SC", "Hannotate SC", "HanziPen SC",
  "Lantinghei SC", "Weibei SC", "Yuppy SC", "Source Han Sans SC", "Source Han Serif SC", "Noto Sans CJK SC",
  "Noto Serif CJK SC", "Helvetica Neue", "Avenir", "Avenir Next", "Futura", "Gill Sans", "Optima", "Didot",
  "Baskerville", "Hoefler Text", "American Typewriter", "Menlo", "Monaco", "Chalkduster", "Marker Felt",
  "Noteworthy", "Papyrus", "Zapfino", "Snell Roundhand", "Copperplate", "Palatino", "SF Pro", "Segoe UI",
]);
