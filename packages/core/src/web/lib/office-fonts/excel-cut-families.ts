/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// The sheet's own family for a face Windows Excel prints lighter than the cut
// the font catalog names for its stand-in.
//
// The catalog is Word's and PowerPoint's too and stays as it is: 游ゴシック
// Medium is Noto Sans JP's 500 there, half as heavy again as the ink Excel
// puts down (excel-windows-stand-in-weights.json, measured off the font
// programs in Excel's own PDFs; the API's writer prints the same table's
// cuts). A sheet names, in front of such a face, a family of its own --
// "MRAI Sheet 游ゴシック Medium" -- declared with the face's very faces (its
// local() names, its line overrides, its slices and their ranges) at the
// measured weight (the served slices are variable: their wght axis is set
// to the weight of equal ink, 328 for 游ゴシック Medium's Han where the
// catalog draws 500), and the face's Latin from its recut clone
// (cjk-latin-faces.json `recuts`: the Office face's advances, as before).
// Han and kana are a full em at every weight and the clone keeps the
// advances, so a line is as wide as it was; only its ink changes. Nothing but a sheet's stack names the family
// (lib/spreadsheet/native-typography `spreadsheetFontStack`).
//
// A reader of a stack's first family -- row heights, printed line metrics,
// digit widths are all kept by face -- asks `stackPrimaryFamily`
// (hangul-em.ts), which passes over the family as it does a Hangul clone.
//
// `window.__officeSheetCutOff = true` leaves every stack as it was.
import cjkLatinFaces from "./cjk-latin-faces.json";
import standInWeights from "./excel-windows-stand-in-weights.json";

/** The weights (wght) of the stand-in one style of a face is drawn at. */
export interface ExcelCutStyle {
  /** For a slice holding Han, kana or other East Asian characters. */
  readonly cjk: number;
  /** For any other slice. */
  readonly latin: number;
}

export interface ExcelCutFamily {
  /** The family a sheet's stack names in front of the face. */
  readonly alias: string;
  /** The face, as the stack names it. */
  readonly face: string;
  readonly regular: ExcelCutStyle;
  /** Absent for a face Windows has no bold of: Excel prints its regular
   * emboldened, and so the family has no bold face either. */
  readonly bold?: ExcelCutStyle;
  /** The recut Latin clone's files by weight ("400", "700"), under /fonts/cjk-latin/. */
  readonly latin: Readonly<Record<string, string>>;
  /** The characters the clone has. */
  readonly latinRange: string | undefined;
}

export const EXCEL_CUT_FAMILY_PREFIX = "MRAI Sheet ";
const PREFIX_KEY = EXCEL_CUT_FAMILY_PREFIX.toLowerCase();

interface WeightsTable {
  readonly faces: Readonly<Record<string, { readonly regular?: Partial<ExcelCutStyle>; readonly bold?: Partial<ExcelCutStyle> }>>;
  readonly names: Readonly<Record<string, string>>;
}
interface ClonesTable {
  readonly faces: Readonly<Record<string, Readonly<Record<string, string>>>>;
  readonly names: Readonly<Record<string, string>>;
  readonly ranges?: Readonly<Record<string, string>>;
  readonly recuts?: Readonly<Record<string, string>>;
}
const WEIGHTS = standInWeights as unknown as WeightsTable;
const CLONES = cjkLatinFaces as unknown as ClonesTable;

const unquote = (name: string) => name.trim().replace(/^["']|["']$/g, "").trim();

function cutsOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __officeSheetCutOff?: boolean }).__officeSheetCutOff === true;
}

function styleOf(entry: Partial<ExcelCutStyle> | undefined): ExcelCutStyle | undefined {
  return entry && typeof entry.cjk === "number" && typeof entry.latin === "number" &&
    entry.cjk >= 100 && entry.cjk <= 900 && entry.latin >= 100 && entry.latin <= 900
    ? { cjk: entry.cjk, latin: entry.latin }
    : undefined;
}

const known = new Map<string, ExcelCutFamily | null>();

/**
 * The sheet's family for `face` (one family name, as a stack names it), or
 * null: a face the table measures whose Latin has a recut clone -- without
 * one the stand-in's own Latin would be drawn in another cut, at that cut's
 * advances (the Yu Gothic UI faces, which stay the catalog's).
 */
export function excelCutFamilyOf(face: string | null | undefined): ExcelCutFamily | null {
  if (!face || cutsOff()) return null;
  const name = unquote(face);
  const cached = known.get(name);
  if (cached !== undefined) return cached;
  const key = name.toLowerCase();
  let found: ExcelCutFamily | null = null;
  const entry = key.startsWith(PREFIX_KEY) ? undefined : WEIGHTS.faces[WEIGHTS.names[key] ?? key];
  const regular = styleOf(entry?.regular);
  const slug = CLONES.names[key];
  const recut = slug ? CLONES.recuts?.[slug] : undefined;
  const latin = recut ? CLONES.faces[recut] : undefined;
  if (regular && recut && latin?.["400"]) {
    const bold = styleOf(entry?.bold);
    found = {
      alias: `${EXCEL_CUT_FAMILY_PREFIX}${name}`,
      face: name,
      regular,
      ...(bold && latin["700"] ? { bold } : {}),
      latin,
      latinRange: CLONES.ranges?.[recut] ?? (slug ? CLONES.ranges?.[slug] : undefined),
    };
  }
  if (known.size >= 512) known.clear();
  known.set(name, found);
  return found;
}

/** Whether a stack token is one of the sheet's cut families. */
export function isExcelCutFamily(family: string): boolean {
  return unquote(family).toLowerCase().startsWith(PREFIX_KEY);
}

/** The face a family name stands for: a cut family's face, any other name itself (unquoted). */
export function excelCutFamilyFace(family: string): string {
  const name = unquote(family);
  return name.toLowerCase().startsWith(PREFIX_KEY) ? name.slice(EXCEL_CUT_FAMILY_PREFIX.length).trim() : name;
}

/** `stack` less the cut families at its head: a stack a sheet built, handed
 * back as a family to build from. */
export function withoutLeadingExcelCutFamilies(stack: string): string {
  if (!stack.toLowerCase().includes(PREFIX_KEY)) return stack;
  const tokens = stack.split(",");
  let at = 0;
  while (at < tokens.length - 1 && isExcelCutFamily(tokens[at]!)) at += 1;
  return at ? tokens.slice(at).join(",").trimStart() : stack;
}

/**
 * Whether bold text in `fontFamily` (a CSS family list) is drawn by a cut
 * family that has no bold face: Windows has no bold of the face and Excel
 * prints its regular emboldened (游ゴシック Medium, 游ゴシック Light), as a
 * sheet draws an Office family without a bold.
 */
export function excelCutFamilyEmboldens(fontFamily: string | null | undefined): boolean {
  if (typeof fontFamily !== "string" || !fontFamily) return false;
  const first = fontFamily.split(",")[0] ?? "";
  if (!isExcelCutFamily(first)) return false;
  const cut = excelCutFamilyOf(excelCutFamilyFace(first));
  return cut !== null && cut.bold === undefined;
}

/** East Asian blocks: a slice of the stand-in holding any of them is drawn at the `cjk` weight. */
const EAST_ASIAN_BLOCKS: ReadonlyArray<readonly [number, number]> = [
  [0x2e80, 0xd7ff], [0xf900, 0xfaff], [0xfe30, 0xfe4f], [0xff00, 0xffef], [0x20000, 0x3ffff],
];

/** Whether a unicode-range (as code point pairs; none: every character) holds East Asian characters. */
export function excelCutRangeIsEastAsian(ranges: ReadonlyArray<readonly [number, number]> | null | undefined): boolean {
  if (!ranges) return true;
  return ranges.some(([start, end]) => EAST_ASIAN_BLOCKS.some(([low, high]) => start <= high && end >= low));
}
