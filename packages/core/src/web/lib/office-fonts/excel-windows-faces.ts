/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import cellPitchFaces from "./excel-windows-cell-pitch-faces.json";
import data from "./excel-windows-faces.json";

/**
 * The faces Windows Excel (Microsoft 365, cloud fonts on) draws a font record
 * in, read from its own PDFs of one row per named face
 * (xlsx-rounds/lookwin/faces, `excel-windows-faces.json` names the probe).
 *
 * Excel draws through GDI: a name Windows has is drawn in that face, or in the
 * face Windows aliases it to (Helvetica is Arial, Times is Times New Roman);
 * a name it lacks -- Carlito, Helvetica Neue, "Calibri (Body)", a name with a
 * trailing space or in quotes -- is drawn in the face GDI maps the record's
 * family class and charset to: Arial for a record that says nothing (or
 * swiss), Times New Roman for roman, Courier New for modern, MS PGothic /
 * Gulim / SimSun / PMingLiU for the East Asian charsets. Han and kana the
 * Latin face has no glyphs for come from its link (Arial: Microsoft YaHei,
 * Times New Roman: SimSun), Hangul from Malgun Gothic or Batang.
 *
 * Excel sizes the grid in that same face: a workbook whose Normal font
 * Windows lacks has Arial's digit width and row height, not Calibri's.
 * Names the probe did not reach are left to the caller.
 */
export interface ExcelWindowsFaces {
  /** The face its Latin text is drawn in. */
  latin: string;
  /** The face its Han and kana are drawn in. */
  han: string;
  /** The face its Hangul is drawn in. */
  hangul: string;
}

/** The record's own family class (`<family val>`) and charset (`<charset val>`). */
export interface ExcelFontRecordClass {
  family?: number | null;
  charset?: number | null;
}

/**
 * The class a page carries for a font record whose face Windows lacks
 * (`nativeGrid.font.record`, a cell style's `fontRecord`; the API's
 * font_record_class, reading xlsxFontRecordClass): `{}` when the record states
 * none, which GDI draws in Arial. Undefined for anything else -- a page
 * imported before carries none and keeps drawing the name by its corpus
 * majority.
 */
export function excelFontRecordOf(value: unknown): ExcelFontRecordClass | undefined {
  if (value === null || typeof value !== "object" || Array.isArray(value)) return undefined;
  const out: ExcelFontRecordClass = {};
  for (const key of ["family", "charset"] as const) {
    const entry = (value as Record<string, unknown>)[key];
    if (entry === undefined) continue;
    if (typeof entry !== "number" || !Number.isInteger(entry) || entry < 0 || entry > 255) return undefined;
    out[key] = entry;
  }
  return out;
}

type Table = Record<string, Record<string, ExcelWindowsFaces>>;
const DRAWN: ReadonlyMap<string, string> = new Map(Object.entries((data as { drawn: Record<string, string> }).drawn));
const LACKING: ReadonlySet<string> = new Set((data as { lacking: string[] }).lacking);
const RECORDLESS: ReadonlyMap<string, string> = new Map(Object.entries((data as { recordless?: Record<string, string> }).recordless ?? {}));
const SUBSTITUTES = (data as unknown as { substitutes: { family: Record<string, ExcelWindowsFaces>; charset: Table } }).substitutes;

/** The faces GDI maps a font record Windows lacks to, by its charset and family class. */
export function excelWindowsSubstitute(record: ExcelFontRecordClass = {}): ExcelWindowsFaces {
  const charset = record.charset == null ? undefined : SUBSTITUTES.charset[String(record.charset)];
  // A charset's own pick for the family, else its pick for any family; a
  // charset that names only some families (161 Greek: script 4 and
  // decorative 5) leaves the rest to the family table (probe fresh-co).
  const faces = charset ? charset[String(record.family)] ?? charset["*"] : undefined;
  return faces ?? SUBSTITUTES.family[String(record.family)] ?? SUBSTITUTES.family["*"];
}

/**
 * The faces Windows Excel draws `name` in, where that is not `name` itself;
 * undefined for a face it has (or one the probe did not reach). `name` as the
 * file states it (a trailing space is part of it).
 */
export function excelWindowsDrawnFaces(
  name: string | null | undefined,
  record?: ExcelFontRecordClass,
): ExcelWindowsFaces | undefined {
  if (!name) return undefined;
  const key = name.toLowerCase();
  // A face it lacks, with the record's own class: GDI's pick for that class
  // (a record that states none, `{}`, is Arial's).
  if (LACKING.has(key) && record) {
    // A Mac PostScript name stating no class: GDI's pick by the name (lookwin
    // p234-lw: TimesNewRomanPSMT alone is Times New Roman; with family 2, as
    // Excel for Mac writes it, Arial, R462-X17).
    const recordless = record.family == null && record.charset == null ? RECORDLESS.get(key) : undefined;
    return recordless ? facesOf(recordless) : excelWindowsSubstitute(record);
  }
  const drawn = DRAWN.get(key);
  if (!drawn) return undefined;
  // Else the face its records in real files land on (明朝: MS PMincho), or
  // the face Windows aliases the name to (Helvetica, Times), with its links.
  return facesOf(drawn);
}

const FACES_BY_LATIN: ReadonlyMap<string, ExcelWindowsFaces> = (() => {
  const out = new Map<string, ExcelWindowsFaces>();
  const add = (faces: ExcelWindowsFaces) => {
    if (!out.has(faces.latin)) out.set(faces.latin, faces);
  };
  for (const faces of Object.values(SUBSTITUTES.family)) add(faces);
  for (const byFamily of Object.values(SUBSTITUTES.charset)) for (const faces of Object.values(byFamily)) add(faces);
  return out;
})();

function facesOf(latin: string): ExcelWindowsFaces {
  return FACES_BY_LATIN.get(latin) ?? { latin, han: SUBSTITUTES.family["*"].han, hangul: SUBSTITUTES.family["*"].hangul };
}

/**
 * The lower-case key Office metric tables (digit widths, row heights, print
 * and wrap advances) are read under for `family`: the face Windows Excel
 * draws it in. A record that says nothing of its class is the Arial a name
 * Windows lacks is drawn in.
 */
export function excelWindowsFaceKey(family: string | null | undefined, record?: ExcelFontRecordClass | null): string {
  const own = (family ?? "").trim();
  const drawn = excelWindowsDrawnFaces(own, record ?? undefined);
  return (drawn?.latin ?? own).toLowerCase();
}

const CELL_PITCH_FACES: ReadonlySet<string> = new Set((cellPitchFaces as { faces: string[] }).faces);

/**
 * Whether `family` is a fixed-pitch East Asian face of GDI's (ＭＳ ゴシック,
 * ＭＳ 明朝, 黑体, 新宋体, 細明體, 楷体, 仿宋, 굴림체, 돋움체, 바탕체, 궁서체):
 * Windows Excel prints each of its full-width characters twice the half-width
 * cell wide, 2 x floor(ppem / 2 + 0.5) printer dots, not at the em as it sets
 * 宋体's or ＭＳ Ｐゴシック's (lookwin p186-lw; the API's
 * excel_windows_cell_pitch, the same table).
 */
export function excelWindowsCellPitch(family: string | null | undefined): boolean {
  return !!family && CELL_PITCH_FACES.has(family.trim().replace(/^["']|["']$/g, "").trim().toLowerCase());
}
