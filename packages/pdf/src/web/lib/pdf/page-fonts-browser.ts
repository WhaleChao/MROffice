/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The browser half of a PDF page's fonts from readers' computers
 * (page-fonts.ts): the page's list, a face's verified program, its CSS face
 * for drawing, and -- in Chrome and Edge, which read installed fonts (Local
 * Font Access) -- a face of the reader's own font prepared for the page.
 */

import { apiFetchRaw } from "@mrnote/core/web/lib/api.ts";
import { verifiedFontAssetBytes } from "@mrnote/core/web/components/office-ribbon/officeFontAssets.ts";
import { resolveOfficeLocalFontVariant } from "@mrnote/core/web/components/office-ribbon/officeFontMetadata.ts";
import type { OfficeLocalFontRecord } from "@mrnote/core/web/components/office-ribbon/officeTypographyCatalog.ts";
import { readCffSfntProgramAsTrueType, readSfntProgram } from "@mrnote/core/web/lib/pdf/engine/native/sfnt/sfnt-reader.ts";
import { buildPdfNativeSfntSubset } from "./engine/native/sfnt/sfnt-subset";
import {
  pdfFontCollectionFace,
  pdfPageFontCssFamily,
  readPdfPageFont,
  type PdfPageFont,
} from "./page-fonts";

/** Why a reader's font cannot be set: the browser cannot read installed
 * fonts, the reader declined, the font is gone, its licence forbids
 * embedding, or the upload failed. */
export type PdfLocalFontRefusal = "unsupported" | "denied" | "missing" | "restricted" | "failed";

export class PdfLocalFontError extends Error {
  constructor(readonly code: PdfLocalFontRefusal, readonly family: string) {
    super(`The font ${family} cannot be set in this PDF (${code}).`);
    this.name = "PdfLocalFontError";
  }
}

const listed = new Map<string, Promise<readonly PdfPageFont[]>>();

/** The page's fonts from readers' computers (asked once per page; a font
 * the reader adds is put in by rememberPdfPageFont). */
export function listPdfPageFonts(pageId: string, refresh = false): Promise<readonly PdfPageFont[]> {
  const cached = listed.get(pageId);
  if (cached && !refresh) return cached;
  const loading = (async () => {
    const response = await apiFetchRaw(`/api/v1/pages/${encodeURIComponent(pageId)}/pdf-fonts`);
    if (!response.ok) return [];
    const body = (await response.json()) as { fonts?: unknown };
    return Array.isArray(body.fonts) ? body.fonts.flatMap((font) => readPdfPageFont(font) ?? []) : [];
  })().catch(() => [] as readonly PdfPageFont[]);
  listed.set(pageId, loading);
  return loading;
}

function rememberPdfPageFont(pageId: string, font: PdfPageFont): void {
  const known = listed.get(pageId) ?? Promise.resolve([] as readonly PdfPageFont[]);
  listed.set(pageId, known.then((fonts) =>
    fonts.some((entry) => entry.sha256 === font.sha256) ? fonts : [...fonts, font]));
}

/** A page font's program, read from the page and checked against its
 * digest and size (kept for the session). */
export async function loadPdfPageFontProgram(font: PdfPageFont, signal?: AbortSignal): Promise<Uint8Array> {
  const bytes = await verifiedFontAssetBytes({
    attachmentId: font.attachmentId,
    sha256: font.sha256,
    sizeBytes: font.size,
    postscriptName: "",
    role: font.weight >= 600 ? (font.italic ? "boldItalic" : "bold") : font.italic ? "italic" : "regular",
  }, signal);
  return new Uint8Array(bytes);
}

const registered = new Map<string, Promise<void>>();

/** The page font drawn by the document under pdfPageFontCssFamily. */
export function registerPdfPageFontFace(font: PdfPageFont): Promise<void> {
  const key = font.sha256;
  const known = registered.get(key);
  if (known) return known;
  const loading = (async () => {
    if (typeof FontFace === "undefined" || typeof document === "undefined") return;
    const bytes = await loadPdfPageFontProgram(font);
    const face = new FontFace(pdfPageFontCssFamily(font), Uint8Array.from(bytes).buffer);
    await face.load();
    document.fonts.add(face);
  })();
  registered.set(key, loading);
  loading.catch(() => registered.delete(key));
  return loading;
}

interface LocalFontAccess {
  queryLocalFonts?: (options?: { postscriptNames?: string[] }) => Promise<OfficeLocalFontRecord[]>;
}

const STYLE_WEIGHTS: ReadonlyArray<readonly [RegExp, number]> = [
  [/thin|hairline/i, 100], [/extra\s*light|ultra\s*light/i, 200], [/light/i, 300],
  [/medium/i, 500], [/semi\s*bold|demi\s*bold/i, 600], [/extra\s*bold|ultra\s*bold/i, 800],
  [/black|heavy/i, 900], [/bold/i, 700],
];

function styleWeight(style: string): number {
  for (const [pattern, weight] of STYLE_WEIGHTS) if (pattern.test(style)) return weight;
  return 400;
}

/** Whether this browser can read the fonts installed on the computer. */
export function pdfLocalFontsReadable(): boolean {
  return typeof window !== "undefined" && typeof (window as Window & LocalFontAccess).queryLocalFonts === "function";
}

/**
 * The face of an installed family a style is set in, prepared for the page:
 * read from the computer (the reader grants the browser this once), one
 * face of a collection taken out, sent to the server (which checks the
 * licence and prepares it), and the prepared program read back. Refused with
 * PdfLocalFontError.
 */
/** Characters a page keeps of a font beside the text it is chosen for (the
 * ones typing reaches for first): printable ASCII and Latin-1, the general
 * and CJK punctuation and full-width forms, and -- for a Chinese face -- the
 * 3,755 common characters of GB 2312 level 1. Only those the face has. */
function commonCharacters(): string {
  let text = "";
  const range = (from: number, to: number) => { for (let code = from; code <= to; code += 1) text += String.fromCodePoint(code); };
  range(0x20, 0x7e);
  range(0xa0, 0xff);
  range(0x2010, 0x2027);
  range(0x3000, 0x303f);
  range(0xff01, 0xff5e);
  try {
    const gb = new TextDecoder("gbk");
    for (let row = 0xb0; row <= 0xd7; row += 1) {
      const bytes: number[] = [];
      for (let cell = 0xa1; cell <= (row === 0xd7 ? 0xf9 : 0xfe); cell += 1) bytes.push(row, cell);
      text += gb.decode(Uint8Array.from(bytes));
    }
  } catch {
    // No GB decoder: the text chosen and the punctuation only.
  }
  return text;
}

/**
 * The part of a face the page keeps: the text it is chosen for and the
 * common characters, as a TrueType program with the face's own names and
 * licence bits (a CFF face's glyphs turned quadratic here). WPS embeds the
 * characters a PDF uses; the server prepares this in moments instead of
 * converting a whole 50,000-glyph face.
 */
async function pageFontSubset(face: Uint8Array, text: string): Promise<Uint8Array> {
  const program = new DataView(face.buffer, face.byteOffset, face.byteLength).getUint32(0) === 0x4f54544f
    ? await readCffSfntProgramAsTrueType(face)
    : await readSfntProgram(face);
  const wanted = [...new Set([...text, ...commonCharacters()])]
    .filter((character) => (program.cmap.get(character.codePointAt(0)!) ?? 0) !== 0)
    .slice(0, 8_000);
  if (!wanted.length) throw new Error("The face has none of these characters.");
  return buildPdfNativeSfntSubset({ program, text: wanted.join(""), maxCharacters: 8_000 }).bytes;
}

export async function preparePdfPageFont(
  pageId: string,
  family: string,
  style: Readonly<{ bold: boolean; italic: boolean }>,
  text = "",
  signal?: AbortSignal,
): Promise<PdfPageFont> {
  const query = (window as Window & LocalFontAccess).queryLocalFonts;
  if (typeof query !== "function") throw new PdfLocalFontError("unsupported", family);
  let records: OfficeLocalFontRecord[];
  try {
    records = (await query()).filter((record) => record.family.toLowerCase() === family.toLowerCase());
  } catch (error) {
    throw new PdfLocalFontError(
      error instanceof DOMException && (error.name === "NotAllowedError" || error.name === "SecurityError") ? "denied" : "failed",
      family,
    );
  }
  if (!records.length) throw new PdfLocalFontError("missing", family);
  // CSS matching over the family's faces: the slant asked, then the weight.
  const italic = (record: OfficeLocalFontRecord) => /italic|oblique/i.test(record.style);
  const slanted = records.filter((record) => italic(record) === style.italic);
  const pool = slanted.length ? slanted : records;
  const wanted = style.bold ? 700 : 400;
  const record = [...pool].sort((left, right) =>
    Math.abs(styleWeight(left.style) - wanted) - Math.abs(styleWeight(right.style) - wanted))[0];
  signal?.throwIfAborted();
  let source: Uint8Array;
  try {
    const blob = await record.blob!();
    source = new Uint8Array(await blob.arrayBuffer());
  } catch {
    throw new PdfLocalFontError("failed", family);
  }
  const face = pdfFontCollectionFace(source, record.postscriptName);
  if (!face) throw new PdfLocalFontError("failed", family);
  let subset: Uint8Array;
  try {
    subset = await pageFontSubset(face, text);
  } catch {
    throw new PdfLocalFontError("failed", family);
  }
  const faceBlob = new Blob([subset.slice().buffer], { type: "font/ttf" });
  const variant = await resolveOfficeLocalFontVariant({
    family: record.family, style: record.style, fullName: record.fullName,
    postscriptName: record.postscriptName, blob: async () => faceBlob,
  }, { inspectSource: true });
  if (!variant?.fontFace) throw new PdfLocalFontError("failed", family);
  const form = new FormData();
  form.set("file", faceBlob, "font.ttf");
  form.set("font_face", JSON.stringify({ face: variant.fontFace, standalone: false }));
  let response: Response;
  try {
    response = await apiFetchRaw(`/api/v1/pages/${encodeURIComponent(pageId)}/attachments/upload`, {
      method: "POST", body: form, signal,
    });
  } catch (error) {
    const code = (error as { code?: unknown } | null)?.code;
    throw new PdfLocalFontError(code === "document_font_embedding_restricted" ? "restricted" : "failed", family);
  }
  const result = (await response.json().catch(() => null)) as {
    attachment_id?: string;
    font?: { family?: string; file?: { attachmentId?: string; sha256?: string; sizeBytes?: number; role?: string } };
  } | null;
  const file = result?.font?.file;
  const prepared = readPdfPageFont({
    family: result?.font?.family,
    weight: file?.role === "bold" || file?.role === "boldItalic" ? 700 : 400,
    italic: file?.role === "italic" || file?.role === "boldItalic",
    sha256: file?.sha256,
    size: file?.sizeBytes,
    attachmentId: file?.attachmentId,
  });
  if (!prepared) throw new PdfLocalFontError("failed", family);
  await loadPdfPageFontProgram(prepared, signal);
  rememberPdfPageFont(pageId, prepared);
  // Drawn (and measured) in its own face from the first row.
  await registerPdfPageFontFace(prepared).catch(() => undefined);
  return prepared;
}
