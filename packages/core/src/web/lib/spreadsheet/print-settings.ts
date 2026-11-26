/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { TranslationFn } from "../errors/user-message";
import { sanitizeImageSource } from "./sheet-image";

/**
 * Print / Page Setup (打印 / 页面布局) — Excel/WPS parity.
 *
 * A pure, framework-free module that defines the PrintSettings model,
 * normalisation helpers, and a pagination helper consumed by the BE PDF exporter
 * (via content_json) and validated by the FE schema.
 *
 * Key concepts
 * ─────────────
 * printArea      : optional cell range "A1:D20" restricting what is printed;
 *                  Excel allows several, comma separated ("A1:D20,F1:G5"),
 *                  each printed on its own page.
 * repeatTitleRows: optional row range "1:3" repeated at the top of every page.
 * repeatTitleCols: optional col range "A:B" repeated at the left of every page.
 * orientation    : "portrait" | "landscape" (default landscape for wide sheets).
 * margins        : preset name OR {top,right,bottom,left} in points.
 * scaling        : percent 10–400, OR fit-to-width N pages, OR fit-to-N-pages M.
 *
 * None of the fields are mandatory; a sheet without printSettings uses the
 * same defaults the existing _spreadsheet_to_pdf already applies (landscape,
 * 40pt margins).  That default path is byte-for-byte unchanged.
 */

// ── Types ────────────────────────────────────────────────────────────────────

export type PrintOrientation = "portrait" | "landscape";

export type MarginPreset = "normal" | "narrow" | "wide";

/** Common Excel/WPS paper sizes with exact OOXML identifiers. */
export type PrintPaperSize = (typeof SPREADSHEET_PAPER_SIZES)[number]["id"];

export interface PrinterResolution {
  horizontalDpi?: number;
  verticalDpi?: number;
}

/**
 * ST_UnsignedInt: Excel writes real resolutions (300, 600) and also the
 * Windows DEVMODE constants as unsigned 32-bit values (4294967292 ...
 * 4294967295 for draft/low/medium/high), which must round-trip as written.
 */
export const MAX_PRINTER_DPI = 4_294_967_295;
/** A device printer configuration part is a few KB; 64 KB bounds custody. */
export const MAX_PRINTER_SETTINGS_PART_BYTES = 64 * 1024;
/** base64 of MAX_PRINTER_SETTINGS_PART_BYTES. */
export const MAX_PRINTER_SETTINGS_PART_CHARS =
  Math.ceil(MAX_PRINTER_SETTINGS_PART_BYTES / 3) * 4;
const PRINTER_SETTINGS_PART_BASE64 = /^[A-Za-z0-9+/]+={0,2}$/;

export interface PrintMargins {
  top: number;    // points
  right: number;
  bottom: number;
  left: number;
}

/** Percent scaling: 10–400 (matches Excel). */
export interface ScalePercent {
  kind: "percent";
  percent: number; // 10 – 400
}

/** Fit sheet to N pages wide (height is automatic). */
export interface ScaleFitToWidth {
  kind: "fitToWidth";
  pages: number; // 1 – 100
}

/** Legacy editor policy: fit the sheet to N pages total (both axes). */
export interface ScaleFitToPages {
  kind: "fitToPages";
  pages: number; // 1 – 100
}

/** Native Excel policy: constrain width and height independently (0 = Automatic). */
export interface ScaleFitToDimensions {
  kind: "fitToPages";
  widthPages: number; // 0 – 100
  heightPages: number; // 0 – 100
}

export type PrintScaling =
  | ScalePercent
  | ScaleFitToWidth
  | ScaleFitToPages
  | ScaleFitToDimensions;

/**
 * A print header or footer (Excel/WPS "Page Setup ▸ Header/Footer"): three
 * independently-aligned text segments — left / center / right — each of which
 * may embed Excel-style field CODES that the renderer substitutes per page:
 *   - `&[Page]`  → the current page number
 *   - `&[Pages]` → the total page count
 *   - `&[Date]`  → the print date
 *   - `&[Time]`  → the print time
 *   - `&[Tab]`   → the sheet (tab) name
 *   - `&[File]`  → the file (workbook) name
 * A segment is plain text otherwise. Empty / whitespace-only segments are
 * dropped on normalize so an all-blank header serializes to nothing.
 */
export interface PrintHeaderFooter {
  left?: string;
  center?: string;
  right?: string;
  /**
   * OPTIONAL formatting overlay (Excel's `&"font,style"` / `&B` / `&K…` codes):
   * when present, each segment's runs concatenate to EXACTLY the plain-text
   * segment above (field codes included), so renderers that ignore runs still
   * print the right text and renderers that honor them add the formatting.
   * The run style vocabulary mirrors the cell style model.
   */
  leftRuns?: PrintHeaderFooterRun[];
  centerRuns?: PrintHeaderFooterRun[];
  rightRuns?: PrintHeaderFooterRun[];
  /**
   * The picture of a section (Excel's Insert Picture in the Custom Header /
   * Footer box; `&G`, shown as `&[Picture]`). Excel allows one picture per
   * section and draws it where the section's `&[Picture]` code sits; a
   * section without the code holds no picture, so deleting the code deletes
   * the picture ({@link normalizeHeaderFooter}).
   */
  leftPicture?: PrintHeaderFooterPicture;
  centerPicture?: PrintHeaderFooterPicture;
  rightPicture?: PrintHeaderFooterPicture;
}

/**
 * A header/footer section's picture with Excel's Format Picture settings
 * (VBA `Graphic`: Height, Width, Crop*, ColorType, Brightness, Contrast,
 * LockAspectRatio; written as the worksheet's `legacyDrawingHF` VML shape).
 */
export interface PrintHeaderFooterPicture {
  /** An embedded raster data URL (lib/spreadsheet/sheet-image). */
  src: string;
  /** Printed width in points, after cropping and scaling. */
  width: number;
  /** Printed height in points. */
  height: number;
  /**
   * The share of the ORIGINAL picture cropped off each side, 0..1 (VML
   * `v:imagedata` cropleft/croptop/...; Excel's Crop in points is measured
   * against the original size). Absent sides are 0.
   */
  crop?: { left?: number; top?: number; right?: number; bottom?: number };
  /** Format Picture > Picture > Color; absent = Automatic. */
  colorType?: "grayscale" | "blackAndWhite" | "washout";
  /** Brightness 0..1 (Excel's default 0.5 is stored as absent). */
  brightness?: number;
  /** Contrast 0..1 (default 0.5, stored as absent). */
  contrast?: number;
  /** Lock aspect ratio; Excel's default (on) is stored as absent. */
  lockAspectRatio?: false;
  /** Alt Text (the VML shape's title). */
  alt?: string;
  





  metafileRendered?: true;
  sourcePart?: string;
}

/** The section keys of a header/footer band. */
export type PrintHeaderFooterSide = "left" | "center" | "right";

/** Largest printed header/footer picture side, points (22 inches, Excel's page bound). */
export const MAX_HEADER_FOOTER_PICTURE_POINTS = 1584;
/** Alt Text bound, as for a picture in a cell. */
export const MAX_HEADER_FOOTER_PICTURE_ALT = 2_000;
/**
 * The character a renderer gets in place of a section's `&[Picture]` code
 * when it draws the picture there (U+FFFC OBJECT REPLACEMENT CHARACTER).
 */
export const HEADER_FOOTER_PICTURE_MARK = "￼";

const HEADER_FOOTER_PICTURE_TOKEN_RE = /&\[picture\]/i;

/** Whether a section's text holds the picture code. */
export function headerFooterTextHasPicture(text: string | undefined): boolean {
  return typeof text === "string" && HEADER_FOOTER_PICTURE_TOKEN_RE.test(text);
}

function unitInterval(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value)
    ? Math.max(0, Math.min(1, value))
    : undefined;
}

// A sheet's print settings are normalized on every load and commit; the
// picture object a sheet stores is stable, so its (decoding) validation runs
// once per object.
const normalizedHeaderFooterPictures = new WeakMap<
  object,
  PrintHeaderFooterPicture | null
>();

/** Validate one header/footer picture; null for anything unusable. */
export function normalizeHeaderFooterPicture(
  raw: unknown,
): PrintHeaderFooterPicture | null {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const cached = normalizedHeaderFooterPictures.get(raw);
  if (cached !== undefined) return cached;
  const result = normalizeHeaderFooterPictureUncached(raw as Record<string, unknown>);
  normalizedHeaderFooterPictures.set(raw, result);
  return result;
}

function normalizeHeaderFooterPictureUncached(
  r: Record<string, unknown>,
): PrintHeaderFooterPicture | null {
  const src = sanitizeImageSource(r.src);
  if (src === null) return null;
  const size = (value: unknown) =>
    typeof value === "number" && Number.isFinite(value) && value > 0
      ? Math.min(MAX_HEADER_FOOTER_PICTURE_POINTS, Math.round(value * 100) / 100)
      : null;
  const width = size(r.width);
  const height = size(r.height);
  if (width === null || height === null) return null;
  const out: PrintHeaderFooterPicture = { src, width, height };
  if (r.crop && typeof r.crop === "object" && !Array.isArray(r.crop)) {
    const crop = r.crop as Record<string, unknown>;
    const sides: NonNullable<PrintHeaderFooterPicture["crop"]> = {};
    for (const side of ["left", "top", "right", "bottom"] as const) {
      const value = unitInterval(crop[side]);
      if (value !== undefined && value > 0) sides[side] = value;
    }
    // Something of the picture has to remain on each axis.
    const horizontal = (sides.left ?? 0) + (sides.right ?? 0);
    const vertical = (sides.top ?? 0) + (sides.bottom ?? 0);
    if (horizontal < 1 && vertical < 1 && Object.keys(sides).length > 0) {
      out.crop = sides;
    }
  }
  if (
    r.colorType === "grayscale" ||
    r.colorType === "blackAndWhite" ||
    r.colorType === "washout"
  ) {
    out.colorType = r.colorType;
  }
  for (const key of ["brightness", "contrast"] as const) {
    const value = unitInterval(r[key]);
    if (value !== undefined && Math.abs(value - 0.5) > 1e-6) {
      out[key] = Math.round(value * 1000) / 1000;
    }
  }
  if (r.lockAspectRatio === false) out.lockAspectRatio = false;
  if (typeof r.alt === "string") {
    const alt = r.alt
      .replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]/g, "")
      .slice(0, MAX_HEADER_FOOTER_PICTURE_ALT);
    if (alt.trim() !== "") out.alt = alt;
  }
  if (
    r.metafileRendered === true &&
    typeof r.sourcePart === "string" &&
    r.sourcePart.length > 0 &&
    r.sourcePart.length <= 255
  ) {
    out.metafileRendered = true;
    out.sourcePart = r.sourcePart;
  }
  return out;
}

/** Every picture a sheet's print settings hold, in band order. */
export function printSettingsHeaderFooterPictures(
  settings: unknown,
): PrintHeaderFooterPicture[] {
  if (!settings || typeof settings !== "object") return [];
  const record = settings as Record<string, unknown>;
  const out: PrintHeaderFooterPicture[] = [];
  for (const band of [
    "header",
    "footer",
    "firstHeader",
    "firstFooter",
    "evenHeader",
    "evenFooter",
  ]) {
    const value = record[band];
    if (!value || typeof value !== "object") continue;
    for (const side of ["left", "center", "right"] as const) {
      const picture = (value as Record<string, unknown>)[`${side}Picture`];
      if (
        picture &&
        typeof picture === "object" &&
        typeof (picture as { src?: unknown }).src === "string"
      ) {
        out.push(picture as PrintHeaderFooterPicture);
      }
    }
  }
  return out;
}
/**
 * One formatted run inside a header/footer segment. `text` may embed the same
 * `&[…]` field codes as the plain segments. Style fields are additive — absent
 * means "inherit the default header/footer font" (Excel's `&"-,Regular"`).
 */
export interface PrintHeaderFooterRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  doubleUnderline?: boolean;
  strikethrough?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  fontFamily?: string;
  /** Font size in points (Excel bounds: 1–409). */
  fontSize?: number;
  /** Resolved #rrggbb text color (theme `&KTT±nnn` is resolved at import). */
  color?: string;
}

/** The recognized header/footer field codes (Excel "&[…]" tokens). */
export const HEADER_FOOTER_TOKENS = {
  page: "&[Page]",
  pages: "&[Pages]",
  date: "&[Date]",
  time: "&[Time]",
  sheetName: "&[Tab]",
  fileName: "&[File]",
  filePath: "&[Path]",
  /** Excel's `&G`: the section's picture is drawn here. */
  picture: "&[Picture]",
} as const;

/**
 * Excel's "Characters in a header or footer" limit (Excel specifications and
 * limits): 255, counted on the NATIVE codes a section is written as
 * ({@link headerFooterNativeLength}: `&[Page]` is the 2-character `&P`, a
 * literal `&` the 2-character `&&`), never on the longer tokens the page
 * keeps. Editors validate against it; nothing is cut to it on read.
 */
export const MAX_HEADER_FOOTER_SEGMENT = 255;

/**
 * Max characters of one section's text in the page's notation (DoS bound).
 * Tokens outgrow their codes (`&[Pages]` is 8 characters for `&N`'s 2), so a
 * legal 255-character band reads as at most 4 × 255 characters; this bound
 * only cuts text that was over Excel's limit to begin with.
 */
export const MAX_HEADER_FOOTER_SECTION_TEXT = 1024;

/** A header/footer field, by the name its `&[…]` token spells (lower case). */
export type HeaderFooterFieldName =
  | "page"
  | "pages"
  | "date"
  | "time"
  | "tab"
  | "file"
  | "path"
  | "picture";

/** One piece of a section's text: literal characters, or a field. */
export type HeaderFooterPiece =
  | { kind: "text"; text: string }
  | {
      kind: "field";
      field: HeaderFooterFieldName;
      /** `&[Page]+n` / `&[Page]-n` (Excel `&P+n` / `&P-n`): "+2", "-1"; "" for none. */
      offset: string;
    };

const HEADER_FOOTER_FIELD_TOKEN_RE =
  /&\[(page|pages|date|time|tab|file|path|picture)\]/iy;
const HEADER_FOOTER_PAGE_OFFSET_RE = /[+-][0-9]+/y;

/**
 * A section's text read as literal text and fields, left to right — the one
 * reader of the page's header/footer notation, which is what Excel's Custom
 * Header boxes show ("Headers and footers in a worksheet"): `&&` is a literal
 * ampersand (so a literal "&[Page]" is `&&[Page]`), `&[Page]` … `&[Picture]`
 * (any case) are fields, the page field taking a directly following `+n` /
 * `-n` as its offset, and any other `&` is a literal ampersand (text saved
 * before the escape existed). The server reads the same notation
 * (apps/api/app/services/file_adapters/xlsx/header_footer_codes.py).
 */
export function headerFooterPieces(text: string): HeaderFooterPiece[] {
  const pieces: HeaderFooterPiece[] = [];
  let literal = "";
  let index = 0;
  while (index < text.length) {
    const char = text[index];
    if (char !== "&") {
      literal += char;
      index += 1;
      continue;
    }
    if (text[index + 1] === "&") {
      literal += "&";
      index += 2;
      continue;
    }
    HEADER_FOOTER_FIELD_TOKEN_RE.lastIndex = index;
    const match = HEADER_FOOTER_FIELD_TOKEN_RE.exec(text);
    if (!match) {
      literal += "&";
      index += 1;
      continue;
    }
    if (literal) pieces.push({ kind: "text", text: literal });
    literal = "";
    const field = match[1].toLowerCase() as HeaderFooterFieldName;
    index += match[0].length;
    let offset = "";
    if (field === "page") {
      HEADER_FOOTER_PAGE_OFFSET_RE.lastIndex = index;
      const offsetMatch = HEADER_FOOTER_PAGE_OFFSET_RE.exec(text);
      if (offsetMatch) {
        offset = offsetMatch[0];
        index += offset.length;
      }
    }
    pieces.push({ kind: "field", field, offset });
  }
  if (literal) pieces.push({ kind: "text", text: literal });
  return pieces;
}

/**
 * How many characters a section's text takes in the native header/footer
 * string: a field is its 2-character code (plus a page offset as written), a
 * literal `&` is written `&&`, every other character counts once (a line
 * break included, as Excel counts it).
 */
export function headerFooterNativeLength(text: string): number {
  let total = 0;
  for (const piece of headerFooterPieces(text)) {
    if (piece.kind === "field") {
      total += 2 + piece.offset.length;
    } else {
      total += piece.text.length;
      for (const char of piece.text) if (char === "&") total += 1;
    }
  }
  return total;
}

/**
 * A band's length in native characters as Excel counts it against
 * {@link MAX_HEADER_FOOTER_SEGMENT}: each non-blank section's `&L`/`&C`/`&R`
 * marker plus its text (formatting codes aside, which the writers check).
 */
export function headerFooterBandNativeLength(
  band: Partial<Record<"left" | "center" | "right", string>> | undefined,
): number {
  let total = 0;
  for (const side of ["left", "center", "right"] as const) {
    const text = band?.[side]?.trim();
    if (text) total += 2 + headerFooterNativeLength(text);
  }
  return total;
}

/**
 * The native characters a band leaves for one section's text: Excel's 255
 * less the section's own marker and the other sections.
 */
export function headerFooterSectionRoom(
  band: Partial<Record<"left" | "center" | "right", string>> | undefined,
  side: "left" | "center" | "right",
): number {
  return (
    MAX_HEADER_FOOTER_SEGMENT - 2 - headerFooterBandNativeLength({ ...band, [side]: "" })
  );
}

/** Max formatted runs honored per header/footer segment (DoS bound). Overflow
 * runs keep their TEXT (folded into the last run) and lose only formatting
 * granularity, so the run cap can never drop characters. */
export const MAX_HEADER_FOOTER_RUNS = 24;

/** Max characters honored for a header/footer run font name. */
export const MAX_HEADER_FOOTER_FONT_NAME = 64;

export interface PrintSettings {
  /** A1-notation range, e.g. "A1:F40", or a comma-separated list of ranges
   * ("A1:F40,H1:J5") each printed as its own page(s). Absent/empty = whole
   * sheet. */
  printArea?: string;
  /** Row range to repeat at top of every page, e.g. "1:2".  Absent = none. */
  repeatTitleRows?: string;
  /** Column range to repeat at left of every page, e.g. "A:B".  Absent = none. */
  repeatTitleCols?: string;
  orientation?: PrintOrientation;
  /**
   * Physical paper stock. Absent = US Letter, preserving the editor's original
   * default while allowing imported A-series/JIS sheets to retain pagination.
   */
  paperSize?: PrintPaperSize;
  /** Exact native stock identity, present on new imports or explicit Size edits.
   * Absent on older cached pages whose Letter/A4 Small identities were folded. */
  paperSizeNativeId?: number;
  /**
   * Either a preset name or explicit point values.
   * Absent = "normal" preset (36pt all sides).
   */
  margins?: MarginPreset | PrintMargins;
  /**
   * Scaling rule.  Absent = 100 % (no scaling).
   */
  scaling?: PrintScaling;
  /**
   * Device print resolution (Excel Page Setup ▸ Print quality), carried as
   * custody so a workbook keeps its `horizontalDpi`/`verticalDpi` on export.
   * Never used to lay out anything.
   */
  printerResolution?: PrinterResolution;
  /**
   * The worksheet's device printer configuration part (Excel's
   * `printerSettingsN.bin`, an opaque DEVMODE blob for a configured printer),
   * base64. Carried as custody so export writes the same part back; never
   * interpreted.
   */
  printerSettingsPart?: string;
  /**
   * Manual page breaks (Excel/WPS "Insert Page Break", driven from Page Break
   * Preview): 0-based row indices BEFORE which a new page starts. A break at row
   * 10 means rows …9 end a page and row 10 begins the next. Absent/empty = no
   * manual row breaks (automatic pagination only). De-duplicated + sorted on
   * normalize. View/print concern only — never touches cell data.
   */
  manualRowBreaks?: number[];
  /**
   * Manual page breaks on the COLUMN axis: 0-based column indices BEFORE which a
   * new page starts. Mirror of {@link manualRowBreaks} for horizontal pagination.
   */
  manualColBreaks?: number[];
  /**
   * Print header drawn in the top margin of every page (Excel/WPS "Header").
   * Absent = no header. See {@link PrintHeaderFooter} for the field codes.
   * When {@link differentOddEven} is set this is the ODD-page header (OOXML
   * `oddHeader`), matching Excel.
   */
  header?: PrintHeaderFooter;
  /**
   * Print footer drawn in the bottom margin of every page (Excel/WPS "Footer").
   * Absent = no footer.
   */
  footer?: PrintHeaderFooter;
  /**
   * Excel "Different first page": page 1 uses {@link firstHeader}/
   * {@link firstFooter} (possibly empty = blank) instead of the normal pair.
   * Non-default (true) only — absent/false = same header on every page.
   */
  differentFirstPage?: boolean;
  /**
   * Excel "Different odd and even pages": even pages use {@link evenHeader}/
   * {@link evenFooter}; {@link header}/{@link footer} become the odd-page pair.
   * Non-default (true) only.
   */
  differentOddEven?: boolean;
  /** First-page header, honored when {@link differentFirstPage} is set. */
  firstHeader?: PrintHeaderFooter;
  /** First-page footer, honored when {@link differentFirstPage} is set. */
  firstFooter?: PrintHeaderFooter;
  /** Even-page header, honored when {@link differentOddEven} is set. */
  evenHeader?: PrintHeaderFooter;
  /** Even-page footer, honored when {@link differentOddEven} is set. */
  evenFooter?: PrintHeaderFooter;
  /**
   * Distance from the paper's top edge to the header, in points (OOXML
   * `pageMargins@header`, Excel "Header margin"). Absent = renderer default.
   */
  headerMargin?: number;
  /**
   * Distance from the paper's bottom edge to the footer, in points (OOXML
   * `pageMargins@footer`).
   */
  footerMargin?: number;
  /**
   * Excel "Page Setup ▸ Header/Footer ▸ Scale with document" — OOXML
   * `headerFooter@scaleWithDoc`. Excel's default is ON, so only the explicit
   * non-default `false` is stored; absent = scale with the document.
   */
  headerFooterScaleWithDoc?: boolean;
  /**
   * Excel "Align with page margins" — OOXML `headerFooter@alignWithMargins`.
   * Excel's default is ON; only the explicit non-default `false` is stored.
   */
  headerFooterAlignWithMargins?: boolean;
  /**
   * Print the cell gridlines (Excel/WPS "Sheet ▸ Gridlines"). Excel's PRINT
   * default is OFF (gridlines show on screen but not in print), so this is the
   * NON-default value only — absent/false = gridlines NOT printed.
   */
  showGridlines?: boolean;
  /**
   * Print the row numbers and column letters (Excel/WPS "Sheet ▸ Row and column
   * headings"). Excel's print default is OFF, so this is the non-default value
   * only — absent/false = headings NOT printed.
   */
  showHeadings?: boolean;
  /**
   * Center the printed area horizontally on the page (Excel "Margins ▸ Center
   * on page ▸ Horizontally", OOXML `printOptions@horizontalCentered`).
   * Non-default (true) only.
   */
  centerHorizontally?: boolean;
  /** Center the printed area vertically (`printOptions@verticalCentered`). */
  centerVertically?: boolean;
  /**
   * Explicit starting page number (Excel "Page ▸ First page number", honored
   * when OOXML `useFirstPageNumber` is set; "Auto" = absent).
   */
  firstPageNumber?: number;
  /**
   * Multi-page ordering (Excel "Sheet ▸ Page order"). Excel's default is
   * down-then-over, so only the non-default value is stored.
   */
  pageOrder?: "overThenDown";
  /** Excel "Sheet ▸ Black and white". Non-default (true) only. */
  blackAndWhite?: boolean;
  /** Excel "Sheet ▸ Draft quality" (skips graphics). Non-default only. */
  draft?: boolean;
  /**
   * How printed cell ERRORS render (Excel "Sheet ▸ Cell errors as"). Excel's
   * default prints them as displayed, so only non-default values are stored.
   */
  printErrorsAs?: "blank" | "dash" | "na";
  /**
   * Whether cell comments print (Excel "Sheet ▸ Comments"). Default is not to
   * print them; only the non-default placements are stored.
   */
  printCellComments?: "asDisplayed" | "atEnd";
  /** Print copy count (Excel "Copies"). Stored only when above 1. */
  copies?: number;
}

/** Max manual page breaks honored per axis (DoS bound on the preview overlay). */
export const MAX_MANUAL_BREAKS = 1000;
export const MAX_PARSE_ROW = 1_048_575;
export const MAX_PARSE_COL = 16_383;

// ── Margin presets (in PDF points; 1pt ≈ 0.353 mm) ──────────────────────────

// Excel's Page Layout ▸ Margins gallery (Microsoft Support, "Set page margins
// before printing a worksheet"): Normal is 0.75" top/bottom and 0.7" left/right,
// Narrow keeps 0.75" top/bottom over 0.25" sides, Wide is 1" all round. Only
// the four page sides live here; header/footer distances are separate fields.
export const MARGIN_PRESETS: Record<MarginPreset, PrintMargins> = {
  normal: { top: 54, right: 50.4, bottom: 54, left: 50.4 }, // 0.75" / 0.7"
  narrow: { top: 54, right: 18, bottom: 54, left: 18 },     // 0.75" / 0.25"
  wide:   { top: 72, right: 72, bottom: 72, left: 72 },     // 1"
};

// Applying a gallery preset also sets the header and footer distances
// (desktop Excel, read back after each preset): Normal and Narrow 0.3",
// Wide 0.5". 0.3" is also the distance a sheet has when the file names none.
export const MARGIN_PRESET_HEADER_FOOTER: Record<MarginPreset, number> = {
  normal: 21.6,
  narrow: 21.6,
  wide: 36,
};

/**
 * One paper authority for the Size gallery, Page Setup, import and pagination.
 * Dimensions are PDF points (72/in), in the stock's intrinsic orientation:
 * Ledger and the double postcard are wide even before a landscape transform.
 * Keep the existing B4/B5 JIS dimensions; do not silently relabel them as ISO.
 * Native IDs: SpreadsheetML pageSetup + Windows PaperKind (43/69/70).
 */
export const SPREADSHEET_PAPER_SIZES = [
  { id: "letter", nativeId: 1, width: 612, height: 792, unit: "in", labelKey: "spreadsheet.helperCopy.print.paper.letter" },
  { id: "letter-small", nativeId: 2, width: 612, height: 792, unit: "in", labelKey: "spreadsheet.helperCopy.print.paper.letter-small" },
  { id: "tabloid", nativeId: 3, width: 792, height: 1224, unit: "in", labelKey: "spreadsheet.helperCopy.print.paper.tabloid" },
  { id: "ledger", nativeId: 4, width: 1224, height: 792, unit: "in", labelKey: "spreadsheet.helperCopy.print.paper.ledger" },
  { id: "legal", nativeId: 5, width: 612, height: 1008, unit: "in", labelKey: "spreadsheet.helperCopy.print.paper.legal" },
  { id: "statement", nativeId: 6, width: 396, height: 612, unit: "in", labelKey: "spreadsheet.helperCopy.print.paper.statement" },
  { id: "executive", nativeId: 7, width: 522, height: 756, unit: "in", labelKey: "spreadsheet.helperCopy.print.paper.executive" },
  { id: "a3", nativeId: 8, width: 841.89, height: 1190.551, unit: "mm", labelKey: "spreadsheet.helperCopy.print.paper.a3" },
  { id: "a4", nativeId: 9, width: 595.276, height: 841.89, unit: "mm", labelKey: "spreadsheet.helperCopy.print.paper.a4" },
  { id: "a4-small", nativeId: 10, width: 595.276, height: 841.89, unit: "mm", labelKey: "spreadsheet.helperCopy.print.paper.a4-small" },
  { id: "a5", nativeId: 11, width: 419.528, height: 595.276, unit: "mm", labelKey: "spreadsheet.helperCopy.print.paper.a5" },
  { id: "a6", nativeId: 70, width: 297.638, height: 419.528, unit: "mm", labelKey: "spreadsheet.helperCopy.print.paper.a6" },
  { id: "b4", nativeId: 12, width: 728.504, height: 1031.811, unit: "mm", labelKey: "spreadsheet.helperCopy.print.paper.b4" },
  { id: "b5", nativeId: 13, width: 515.906, height: 728.504, unit: "mm", labelKey: "spreadsheet.helperCopy.print.paper.b5" },
  { id: "folio", nativeId: 14, width: 612, height: 936, unit: "in", labelKey: "spreadsheet.helperCopy.print.paper.folio" },
  { id: "envelope-10", nativeId: 20, width: 297, height: 684, unit: "in", labelKey: "spreadsheet.helperCopy.print.paper.envelope-10" },
  { id: "envelope-dl", nativeId: 27, width: 311.811, height: 623.622, unit: "mm", labelKey: "spreadsheet.helperCopy.print.paper.envelope-dl" },
  { id: "envelope-c5", nativeId: 28, width: 459.213, height: 649.134, unit: "mm", labelKey: "spreadsheet.helperCopy.print.paper.envelope-c5" },
  { id: "envelope-b5", nativeId: 34, width: 498.898, height: 708.661, unit: "mm", labelKey: "spreadsheet.helperCopy.print.paper.envelope-b5" },
  { id: "envelope-c3", nativeId: 29, width: 918.425, height: 1298.268, unit: "mm", labelKey: "spreadsheet.helperCopy.print.paper.envelope-c3" },
  { id: "envelope-c4", nativeId: 30, width: 649.134, height: 918.425, unit: "mm", labelKey: "spreadsheet.helperCopy.print.paper.envelope-c4" },
  { id: "envelope-c6", nativeId: 31, width: 323.15, height: 459.213, unit: "mm", labelKey: "spreadsheet.helperCopy.print.paper.envelope-c6" },
  { id: "envelope-monarch", nativeId: 37, width: 279, height: 540, unit: "in", labelKey: "spreadsheet.helperCopy.print.paper.envelope-monarch" },
  { id: "japanese-postcard", nativeId: 43, width: 283.465, height: 419.528, unit: "mm", labelKey: "spreadsheet.helperCopy.print.paper.japanese-postcard" },
  { id: "double-japanese-postcard", nativeId: 69, width: 566.929, height: 419.528, unit: "mm", labelKey: "spreadsheet.helperCopy.print.paper.double-japanese-postcard" },
  { id: "quarto", nativeId: 15, width: 609.449, height: 779.528, unit: "mm", labelKey: "spreadsheet.helperCopy.print.paper.quarto" },
] as const;

export const PAPER_SIZE_POINTS = Object.fromEntries(
  SPREADSHEET_PAPER_SIZES.map(({ id, width, height }) => [id, { width, height }]),
) as Record<PrintPaperSize, { width: number; height: number }>;

export const PAPER_SIZE_NATIVE_IDS = Object.fromEntries(
  SPREADSHEET_PAPER_SIZES.map(({ id, nativeId }) => [id, nativeId]),
) as Record<PrintPaperSize, number>;








export function newSheetPrintSettings(): Pick<PrintSettings, "paperSize" | "paperSizeNativeId"> {
  return { paperSize: "a4", paperSizeNativeId: PAPER_SIZE_NATIVE_IDS.a4 };
}

export const OOXML_PAPER_SIZES: Readonly<Record<number, PrintPaperSize>> = {
  ...Object.fromEntries(SPREADSHEET_PAPER_SIZES.map(({ id, nativeId }) => [nativeId, id])),
  17: "tabloid", // Existing 11x17 alias.
  18: "letter", // Existing Note alias.
};

export function spreadsheetPaperSizeOptions(selected?: PrintPaperSize) {
  // Quarto remains supported and visible when imported/selected, but is not
  // added to the screenshot's 25-choice authoring list.
  return SPREADSHEET_PAPER_SIZES.filter(({ id }) => id !== "quarto" || id === selected);
}

export function spreadsheetPaperSizeLabels(
  paperSize: PrintPaperSize,
  translate: TranslationFn,
): { title: string; inches: string } {
  const paper = SPREADSHEET_PAPER_SIZES.find(({ id }) => id === paperSize)!;
  const inches = (points: number) => Number((points / 72).toFixed(2));
  const primary = (points: number) => {
    if (paper.unit === "mm") return String(Math.round(points * 25.4 / 72));
    const value = points / 72;
    const fractions: Record<number, string> = {
      0.125: "⅛", 0.25: "¼", 0.375: "⅜", 0.5: "½",
      0.625: "⅝", 0.75: "¾", 0.875: "⅞",
    };
    const fraction = fractions[value % 1];
    return fraction ? `${Math.floor(value) || ""}${fraction}` : String(value);
  };
  const unit = translate(paper.unit === "mm" ? "spreadsheet.helperCopy.print.millimeters" : "spreadsheet.helperCopy.print.inches");
  return {
    title: translate("spreadsheet.helperCopy.print.paperTitle", { name: translate(paper.labelKey), width: primary(paper.width), height: primary(paper.height), unit }),
    inches: `${inches(paper.width)}″ × ${inches(paper.height)}″`,
  };
}

/** The regions whose default paper is Letter; everywhere else it is A4. */
const LETTER_REGIONS = new Set(["US", "CA", "MX", "PH", "CL", "CO", "VE", "PR", "GT", "CR", "PA", "DO", "SV", "HN", "NI", "BO"]);

/**
 * The paper a sheet that names none prints on: the region's default, as
 * Excel takes it from the system / printer (XLS-A069: a Chinese Excel opens
 * Page Setup on A4, a US one on Letter). The browser's locale is this
 * runtime's region; a language with no region is read as en → US, others as
 * their main country.
 */
export function regionalDefaultPaperSize(
  locale: string | undefined = typeof navigator !== "undefined" ? navigator.language : undefined,
): "letter" | "a4" {
  if (!locale) return "letter";
  const [language, region] = locale.replace(/_/g, "-").split("-");
  if (region && /^[A-Za-z]{2}$/.test(region)) return LETTER_REGIONS.has(region.toUpperCase()) ? "letter" : "a4";
  return (language ?? "").toLowerCase() === "en" ? "letter" : "a4";
}

/** Resolve absent/unknown persisted stock to the region's default paper. */
export function resolvePaperSizePoints(
  paperSize: PrintSettings["paperSize"],
): { width: number; height: number } {
  return PAPER_SIZE_POINTS[paperSize ?? regionalDefaultPaperSize()] ?? PAPER_SIZE_POINTS.letter;
}

// ── Normalisation ─────────────────────────────────────────────────────────────

function boundedCell(
  column: string,
  row: string,
): { row: number; col: number } | null {
  const col = colLetterToIndex(column);
  const parsedRow = Number(row) - 1;
  if (
    col < 0 ||
    col > MAX_PARSE_COL ||
    !Number.isSafeInteger(parsedRow) ||
    parsedRow < 0 ||
    parsedRow > MAX_PARSE_ROW
  ) {
    return null;
  }
  return { row: parsedRow, col };
}

/** Validate one bounded cell/cell-range print area block such as "A1:D20". */
/**
 * Excel writes its print area and print titles anchored ("$A$1:$B$3",
 * "$1:$2"), and its Page Setup sheet shows them so; a typed range may carry
 * the anchors or not. The stored form drops them.
 */
export function stripRangeAnchors(v: string): string {
  return v.replace(/\$/g, "").trim();
}

export function isValidPrintAreaRangeString(v: unknown): v is string {
  if (typeof v !== "string") return false;
  const match = /^([A-Z]{1,3})([1-9][0-9]*)(?::([A-Z]{1,3})([1-9][0-9]*))?$/i.exec(
    stripRangeAnchors(v),
  );
  if (!match) return false;
  const start = boundedCell(match[1].toUpperCase(), match[2]);
  const end = match[3]
    ? boundedCell(match[3].toUpperCase(), match[4])
    : start;
  return Boolean(
    start &&
      end &&
      start.row <= end.row &&
      start.col <= end.col,
  );
}

/**
 * Excel lets one worksheet own several print areas ("Set or clear a print
 * area on a worksheet": select several ranges with Ctrl, or Add to Print
 * Area; each area prints on its own page). They are stored as one
 * `_xlnm.Print_Area` name whose refs are comma separated. The ceiling bounds
 * every pager that walks the list; Excel itself has no documented limit.
 */
export const MAX_PRINT_AREA_RANGES = 64;

/** Split a persisted print area into its comma-separated blocks. */
export function printAreaRanges(v: unknown): string[] {
  if (typeof v !== "string") return [];
  return v
    .split(",")
    .map((range) => stripRangeAnchors(range).toUpperCase())
    .filter((range) => range.length > 0);
}

/**
 * Validate a print area: one bounded block, or a comma-separated list of
 * them. Excel keeps a repeated or overlapping block (each still prints as its
 * own page), so only the block grammar and the list ceiling are enforced.
 */
export function isValidPrintAreaString(v: unknown): v is string {
  if (typeof v !== "string" || v.trim() === "") return false;
  const ranges = v.split(",");
  return (
    ranges.length <= MAX_PRINT_AREA_RANGES &&
    ranges.every((range) => isValidPrintAreaRangeString(range))
  );
}

/**
 * Canonical persisted form: upper case, no blanks, duplicates dropped, order
 * of first appearance kept (Excel prints the areas in the order they were
 * added). `undefined` when nothing valid remains.
 */
export function normalizePrintAreaString(v: unknown): string | undefined {
  if (!isValidPrintAreaString(v)) return undefined;
  const seen = new Set<string>();
  const ranges: string[] = [];
  for (const range of printAreaRanges(v)) {
    if (seen.has(range)) continue;
    seen.add(range);
    ranges.push(range);
  }
  return ranges.length > 0 ? ranges.join(",") : undefined;
}

/** Column index (0-based) back to its letters. */
function colIndexToLetter(col: number): string {
  let n = col + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function formatA1Range(b: { r1: number; c1: number; r2: number; c2: number }): string {
  return `${colIndexToLetter(b.c1)}${b.r1 + 1}:${colIndexToLetter(b.c2)}${b.r2 + 1}`;
}

/**
 * Excel's Add to Print Area: a selection adjacent to an existing block
 * (sharing its full width or height) extends that block, so A4:B5 added to
 * A1:B3 prints as the one area A1:B5 (Microsoft: "if the cells you want to
 * add are adjacent to the existing print area, the print area is expanded";
 * live audit B164). A non-adjacent selection "creates an additional print
 * area". A block the area already lists is not repeated; a selection lying
 * wholly inside one existing block adds nothing, which is also why the
 * command is offered only then.
 */
export function addPrintAreaRange(
  existing: unknown,
  range: string,
): string | undefined {
  const block = range.trim().toUpperCase();
  if (!isValidPrintAreaRangeString(block)) return normalizePrintAreaString(existing);
  const current = normalizePrintAreaString(existing);
  if (!current) return block;
  if (isRangeInsidePrintArea(block, current)) return current;
  const ranges = printAreaRanges(current);
  const target = parseA1Range(block);
  if (target) {
    for (let i = 0; i < ranges.length; i += 1) {
      const b = parseA1Range(ranges[i]);
      if (!b) continue;
      const contained = b.r1 >= target.r1 && b.r2 <= target.r2 && b.c1 >= target.c1 && b.c2 <= target.c2;
      const sameCols = b.c1 === target.c1 && b.c2 === target.c2;
      const sameRows = b.r1 === target.r1 && b.r2 === target.r2;
      const rowsTouch = target.r1 === b.r2 + 1 || target.r2 === b.r1 - 1;
      const colsTouch = target.c1 === b.c2 + 1 || target.c2 === b.c1 - 1;
      if (contained || (sameCols && rowsTouch) || (sameRows && colsTouch)) {
        const merged = formatA1Range({
          r1: Math.min(b.r1, target.r1),
          r2: Math.max(b.r2, target.r2),
          c1: Math.min(b.c1, target.c1),
          c2: Math.max(b.c2, target.c2),
        });
        return [...ranges.slice(0, i), merged, ...ranges.slice(i + 1)].join(",");
      }
    }
  }
  if (ranges.length >= MAX_PRINT_AREA_RANGES) return current;
  return [...ranges, block].join(",");
}

/** True when `range` lies wholly inside one block of the print area. */
export function isRangeInsidePrintArea(range: string, printArea: unknown): boolean {
  const target = parseA1Range(range);
  if (!target) return false;
  return printAreaRanges(printArea).some((block) => {
    const bounds = parseA1Range(block);
    return (
      bounds !== null &&
      target.r1 >= bounds.r1 &&
      target.r2 <= bounds.r2 &&
      target.c1 >= bounds.c1 &&
      target.c2 <= bounds.c2
    );
  });
}

/** Validate a bounded, forward row range such as "1:3". */
export function isValidTitleRowRangeString(v: unknown): v is string {
  if (typeof v !== "string") return false;
  const match = /^([1-9][0-9]*):([1-9][0-9]*)$/.exec(stripRangeAnchors(v));
  if (!match) return false;
  const start = Number(match[1]) - 1;
  const end = Number(match[2]) - 1;
  return (
    Number.isSafeInteger(start) &&
    Number.isSafeInteger(end) &&
    start <= end &&
    end <= MAX_PARSE_ROW
  );
}

/** Validate a bounded, forward column range such as "A:C". */
export function isValidTitleColRangeString(v: unknown): v is string {
  if (typeof v !== "string") return false;
  const match = /^([A-Z]{1,3}):([A-Z]{1,3})$/i.exec(stripRangeAnchors(v));
  if (!match) return false;
  const start = colLetterToIndex(match[1].toUpperCase());
  const end = colLetterToIndex(match[2].toUpperCase());
  return start >= 0 && start <= end && end <= MAX_PARSE_COL;
}

/** Validate any supported bounded A1 range string. */
export function isValidRangeString(v: unknown): v is string {
  return (
    isValidPrintAreaRangeString(v) ||
    isValidTitleRowRangeString(v) ||
    isValidTitleColRangeString(v)
  );
}

/** Resolve the margin preset/object to concrete point values. */
export function resolveMargins(
  margins: PrintSettings["margins"],
): PrintMargins {
  if (!margins) return MARGIN_PRESETS.normal;
  if (typeof margins === "string") return MARGIN_PRESETS[margins] ?? MARGIN_PRESETS.normal;
  // Explicit object — clamp values to [0, 288] (4 inches).
  const clamp = (n: unknown) =>
    typeof n === "number" && Number.isFinite(n) ? Math.max(0, Math.min(288, n)) : 54;
  return {
    top: clamp(margins.top),
    right: clamp(margins.right),
    bottom: clamp(margins.bottom),
    left: clamp(margins.left),
  };
}

/** Normalise a raw PrintSettings object from persisted JSON. */
export function normalizePrintSettings(raw: unknown): PrintSettings | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const r = raw as Record<string, unknown>;
  const out: PrintSettings = {};

  // printArea: one block, or Excel's comma-separated list of blocks
  const printArea = normalizePrintAreaString(r.printArea);
  if (printArea) out.printArea = printArea;

  // repeatTitleRows / repeatTitleCols
  if (isValidTitleRowRangeString(r.repeatTitleRows))
    out.repeatTitleRows = stripRangeAnchors(r.repeatTitleRows).toUpperCase();
  if (isValidTitleColRangeString(r.repeatTitleCols))
    out.repeatTitleCols = stripRangeAnchors(r.repeatTitleCols).toUpperCase();

  // orientation
  if (r.orientation === "portrait" || r.orientation === "landscape")
    out.orientation = r.orientation;

  // paper size (absent remains the historic Letter default)
  if (
    typeof r.paperSize === "string" &&
    Object.prototype.hasOwnProperty.call(PAPER_SIZE_POINTS, r.paperSize)
  ) {
    out.paperSize = r.paperSize as PrintPaperSize;
    if (
      typeof r.paperSizeNativeId === "number" &&
      Number.isInteger(r.paperSizeNativeId) &&
      OOXML_PAPER_SIZES[r.paperSizeNativeId] === out.paperSize
    ) {
      out.paperSizeNativeId = r.paperSizeNativeId;
    }
  }

  // margins
  if (r.margins !== undefined) {
    if (r.margins === "normal" || r.margins === "narrow" || r.margins === "wide") {
      out.margins = r.margins as MarginPreset;
    } else if (typeof r.margins === "object" && !Array.isArray(r.margins)) {
      out.margins = resolveMargins(r.margins as PrintMargins);
    }
  }

  // scaling
  if (r.scaling !== null && typeof r.scaling === "object" && !Array.isArray(r.scaling)) {
    const sc = r.scaling as Record<string, unknown>;
    if (sc.kind === "percent" && typeof sc.percent === "number") {
      out.scaling = { kind: "percent", percent: Math.max(10, Math.min(400, sc.percent)) };
    } else if (sc.kind === "fitToWidth" && typeof sc.pages === "number") {
      out.scaling = { kind: "fitToWidth", pages: Math.max(1, Math.min(100, sc.pages)) };
    } else if (
      sc.kind === "fitToPages" &&
      typeof sc.widthPages === "number" &&
      typeof sc.heightPages === "number"
    ) {
      const widthPages = Math.max(0, Math.min(100, Math.trunc(sc.widthPages)));
      const heightPages = Math.max(0, Math.min(100, Math.trunc(sc.heightPages)));
      if (widthPages > 0 || heightPages > 0) {
        out.scaling = { kind: "fitToPages", widthPages, heightPages };
      }
    } else if (sc.kind === "fitToPages" && typeof sc.pages === "number") {
      out.scaling = {
        kind: "fitToPages",
        widthPages: Math.max(1, Math.min(100, Math.trunc(sc.pages))),
        heightPages: 1,
      };
    }
  }

  // Manual page breaks (Page Break Preview): keep only positive integer indices,
  // de-dupe + sort, and cap the count. A break at index 0 is dropped (a page
  // cannot start "before row 0" — that is just the start of the sheet).
  const rowBreaks = normalizeBreakArray(r.manualRowBreaks);
  if (rowBreaks.length > 0) out.manualRowBreaks = rowBreaks;
  const colBreaks = normalizeBreakArray(r.manualColBreaks);
  if (colBreaks.length > 0) out.manualColBreaks = colBreaks;

  // Header / footer: each is three optional aligned segments; drop entirely when
  // every segment is blank so an all-empty header serializes to nothing.
  const header = normalizeHeaderFooter(r.header);
  if (header) out.header = header;
  const footer = normalizeHeaderFooter(r.footer);
  if (footer) out.footer = footer;

  // First/even-page variants (Excel "Different first page" / "Different odd and
  // even pages"). The flags are non-default (true) only; a variant header kept
  // WITHOUT its flag stays stored (additive) but renderers key off the flags.
  if (r.differentFirstPage === true) out.differentFirstPage = true;
  if (r.differentOddEven === true) out.differentOddEven = true;
  for (const field of [
    "firstHeader",
    "firstFooter",
    "evenHeader",
    "evenFooter",
  ] as const) {
    const variant = normalizeHeaderFooter(r[field]);
    if (variant) out[field] = variant;
  }

  // Header/footer distance from the paper edge, points, same [0, 288] bound as
  // the page margins.
  for (const field of ["headerMargin", "footerMargin"] as const) {
    const value = r[field];
    if (typeof value === "number" && Number.isFinite(value)) {
      out[field] = Math.max(0, Math.min(288, value));
    }
  }

  // Header/footer layout switches: Excel defaults both ON, so only the explicit
  // non-default false is persisted (mirroring the non-default-only booleans
  // above — a stray truthy value cannot flip a default).
  if (r.headerFooterScaleWithDoc === false) out.headerFooterScaleWithDoc = false;
  if (r.headerFooterAlignWithMargins === false)
    out.headerFooterAlignWithMargins = false;

  // Print gridlines / headings: non-default (true) only, like the view-options
  // booleans — a stray truthy value can't silently flip the default OFF.
  if (r.showGridlines === true) out.showGridlines = true;
  if (r.showHeadings === true) out.showHeadings = true;

  // Center-on-page, sheet print quality, and page-order switches: all Excel
  // non-defaults, stored as such.
  if (r.centerHorizontally === true) out.centerHorizontally = true;
  if (r.centerVertically === true) out.centerVertically = true;
  if (r.pageOrder === "overThenDown") out.pageOrder = "overThenDown";
  if (r.blackAndWhite === true) out.blackAndWhite = true;
  if (r.draft === true) out.draft = true;
  if (
    r.printErrorsAs === "blank" ||
    r.printErrorsAs === "dash" ||
    r.printErrorsAs === "na"
  ) {
    out.printErrorsAs = r.printErrorsAs;
  }
  if (r.printCellComments === "asDisplayed" || r.printCellComments === "atEnd")
    out.printCellComments = r.printCellComments;
  if (
    typeof r.firstPageNumber === "number" &&
    Number.isInteger(r.firstPageNumber) &&
    r.firstPageNumber >= 0 &&
    r.firstPageNumber <= 32_767
  ) {
    out.firstPageNumber = r.firstPageNumber;
  }
  if (
    typeof r.copies === "number" &&
    Number.isInteger(r.copies) &&
    r.copies > 1 &&
    r.copies <= 9_999
  ) {
    out.copies = r.copies;
  }

  if (r.printerResolution && typeof r.printerResolution === "object") {
    const raw = r.printerResolution as Record<string, unknown>;
    const resolution: PrinterResolution = {};
    for (const key of ["horizontalDpi", "verticalDpi"] as const) {
      const dpi = raw[key];
      if (
        typeof dpi === "number" &&
        Number.isInteger(dpi) &&
        dpi >= 1 &&
        dpi <= MAX_PRINTER_DPI
      ) {
        resolution[key] = dpi;
      }
    }
    if (Object.keys(resolution).length > 0) out.printerResolution = resolution;
  }
  if (
    typeof r.printerSettingsPart === "string" &&
    r.printerSettingsPart.length > 0 &&
    r.printerSettingsPart.length <= MAX_PRINTER_SETTINGS_PART_CHARS &&
    PRINTER_SETTINGS_PART_BASE64.test(r.printerSettingsPart)
  ) {
    out.printerSettingsPart = r.printerSettingsPart;
  }

  if (Object.keys(out).length === 0) return undefined;
  return out;
}

/**
 * Normalize a raw manual-page-break array: keep only positive integers (a break
 * "before index 0" is meaningless and dropped), de-duplicate, sort ascending,
 * and cap to {@link MAX_MANUAL_BREAKS}. Returns a fresh array (empty when nothing
 * valid remains, so the field can be omitted).
 */
export function normalizeBreakArray(raw: unknown): number[] {
  if (!Array.isArray(raw)) return [];
  const seen = new Set<number>();
  for (const entry of raw) {
    const n = Number(entry);
    if (!Number.isInteger(n) || n <= 0) continue;
    seen.add(n);
  }
  return Array.from(seen)
    .sort((a, b) => a - b)
    .slice(0, MAX_MANUAL_BREAKS);
}

// ── Header / footer (pure, testable) ──────────────────────────────────────────

/** DoS-bound one header/footer segment; "" / blank → undefined. A segment
 * with text keeps it exactly, leading and trailing spaces included: Excel
 * prints them, and a section indented with spaces is laid out by them (live
 * audit B386). Excel's 255-character limit counts native codes and is checked
 * where text is edited ({@link headerFooterNativeLength}), never by cutting
 * it here. */
function normalizeSegment(raw: unknown): string | undefined {
  if (typeof raw !== "string") return undefined;
  if (raw.trim() === "") return undefined;
  return raw.slice(0, MAX_HEADER_FOOTER_SECTION_TEXT);
}

// C0/C1 control characters other than the line break (a section may run over
// several lines, as Excel's Custom Header boxes do; the plain segment keeps
// its "\n" and so do the runs that spell it), mirroring the docs header/footer
// sanitizer: bytes that could break the band DOM or exported XML are removed,
// never merely escaped.
const HEADER_FOOTER_CONTROL_CHARS_RE = /[\x00-\x09\x0b-\x1f\x7f-\x9f]/g;
// A font name is written inside Excel's &"name,style" code, which cannot hold
// a quote or an ampersand.
const HEADER_FOOTER_FONT_NAME_UNSAFE_RE = /[\x00-\x1f\x7f-\x9f"&]/g;

/** Validate + clamp one formatted run; null when nothing usable remains. */
function normalizeHeaderFooterRun(raw: unknown): PrintHeaderFooterRun | null {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const r = raw as Record<string, unknown>;
  if (typeof r.text !== "string") return null;
  const text = r.text
    .replace(HEADER_FOOTER_CONTROL_CHARS_RE, " ")
    .slice(0, MAX_HEADER_FOOTER_SECTION_TEXT);
  if (text === "") return null;
  const out: PrintHeaderFooterRun = { text };
  for (const flag of [
    "bold",
    "italic",
    "underline",
    "doubleUnderline",
    "strikethrough",
    "superscript",
    "subscript",
  ] as const) {
    if (r[flag] === true) out[flag] = true;
  }
  // One underline and one baseline at a time: Excel's Font list offers one
  // underline style, and of superscript and subscript the first one wins.
  if (out.doubleUnderline) delete out.underline;
  if (out.superscript) delete out.subscript;
  if (typeof r.fontFamily === "string") {
    const fontFamily = r.fontFamily
      .replace(HEADER_FOOTER_FONT_NAME_UNSAFE_RE, "")
      .trim()
      .slice(0, MAX_HEADER_FOOTER_FONT_NAME);
    if (fontFamily !== "") out.fontFamily = fontFamily;
  }
  // Excel's &nn size code is a whole number of points.
  if (typeof r.fontSize === "number" && Number.isFinite(r.fontSize)) {
    const fontSize = Math.round(r.fontSize);
    if (fontSize >= 1 && fontSize <= 409) out.fontSize = fontSize;
  }
  if (typeof r.color === "string" && /^#[0-9a-fA-F]{6}$/.test(r.color)) {
    out.color = r.color.toLowerCase();
  }
  return out;
}

/** Whether a run carries any formatting beyond its text. */
export function headerFooterRunHasFormat(run: PrintHeaderFooterRun): boolean {
  return Object.keys(run).some((key) => key !== "text");
}

/** Whether two runs carry the same formatting (their text aside). */
export function sameHeaderFooterRunFormat(
  a: PrintHeaderFooterRun,
  b: PrintHeaderFooterRun,
): boolean {
  const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
  keys.delete("text");
  for (const key of keys) {
    if (
      a[key as keyof PrintHeaderFooterRun] !==
      b[key as keyof PrintHeaderFooterRun]
    ) {
      return false;
    }
  }
  return true;
}

/**
 * The runs a section keeps beside its plain text `plain`: neighbours with the
 * same format merged, and undefined when they no longer spell `plain` exactly
 * or carry no formatting (the plain text then stands alone, as an unformatted
 * section does). The text is kept as written, spaces at either end included,
 * exactly as the plain segment is (live audit B386).
 */
export function headerFooterRunsForSegment(
  plain: string,
  runs: readonly PrintHeaderFooterRun[] | undefined,
): PrintHeaderFooterRun[] | undefined {
  if (!runs || runs.length === 0) return undefined;
  const out: PrintHeaderFooterRun[] = [];
  for (const run of runs) {
    const text = run.text;
    if (text === "") continue;
    const last = out.at(-1);
    if (last && sameHeaderFooterRunFormat(last, run)) {
      out[out.length - 1] = { ...last, text: last.text + text };
    } else {
      out.push({ ...run, text });
    }
  }
  const spelled = out.map((run) => run.text).join("");
  if (spelled !== plain) {
    // The runs spell the section but for the control characters a run's
    // text is cleaned of and the plain text keeps (a carriage return a file
    // ends its line with: " OFFICIAL-SENSITIVE#\r"): each run takes its own
    // stretch of the plain text, so the two say the same and the fonts stay.
    // Dropped, every sheet of such a workbook lost its header's fonts at the
    
    if (spelled.length !== plain.length || spelled !== plain.replace(HEADER_FOOTER_CONTROL_CHARS_RE, " ")) return undefined;
    let at = 0;
    for (let index = 0; index < out.length; index += 1) {
      const run = out[index]!;
      out[index] = { ...run, text: plain.slice(at, at + run.text.length) };
      at += run.text.length;
    }
  }
  return out.some(headerFooterRunHasFormat) ? out : undefined;
}

/**
 * Normalize one segment's run array: per-run validation, then a run-count cap
 * that FOLDS overflow text into the final run (dropping only formatting
 * granularity, never characters), then a total-length cap aligned with the
 * plain segment bound. Returns undefined when no usable run remains.
 */
function normalizeHeaderFooterRuns(
  raw: unknown,
): PrintHeaderFooterRun[] | undefined {
  if (!Array.isArray(raw)) return undefined;
  const runs: PrintHeaderFooterRun[] = [];
  for (const entry of raw) {
    const run = normalizeHeaderFooterRun(entry);
    if (run) runs.push(run);
  }
  if (runs.length === 0) return undefined;
  if (runs.length > MAX_HEADER_FOOTER_RUNS) {
    const kept = runs.slice(0, MAX_HEADER_FOOTER_RUNS);
    kept[kept.length - 1] = {
      ...kept[kept.length - 1],
      text:
        kept[kept.length - 1].text +
        runs
          .slice(MAX_HEADER_FOOTER_RUNS)
          .map((run) => run.text)
          .join(""),
    };
    runs.length = 0;
    runs.push(...kept);
  }
  let total = 0;
  const bounded: PrintHeaderFooterRun[] = [];
  for (const run of runs) {
    if (total >= MAX_HEADER_FOOTER_SECTION_TEXT) break;
    const room = MAX_HEADER_FOOTER_SECTION_TEXT - total;
    const text = run.text.slice(0, room);
    total += text.length;
    bounded.push(text === run.text ? run : { ...run, text });
  }
  return bounded.length > 0 ? bounded : undefined;
}

/**
 * Normalize a raw header/footer blob into a {@link PrintHeaderFooter}, keeping
 * only the non-blank left/center/right segments. Returns undefined when the blob
 * is not an object or every segment is blank (so the field can be omitted and an
 * all-empty header never persists). A segment's formatted runs are kept only
 * when the segment itself survives and they spell its text exactly
 * ({@link headerFooterRunsForSegment}); a runs array whose plain segment is
 * missing derives the plain text from the runs, so the two views can never
 * disagree.
 */
export function normalizeHeaderFooter(
  raw: unknown,
): PrintHeaderFooter | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const r = raw as Record<string, unknown>;
  const out: PrintHeaderFooter = {};
  for (const [field, runsField, pictureField] of [
    ["left", "leftRuns", "leftPicture"],
    ["center", "centerRuns", "centerPicture"],
    ["right", "rightRuns", "rightPicture"],
  ] as const) {
    const runs = normalizeHeaderFooterRuns(r[runsField]);
    const plain = normalizeSegment(
      typeof r[field] === "string"
        ? r[field]
        : runs
          ? runs.map((run) => run.text).join("")
          : undefined,
    );
    if (plain === undefined) continue;
    out[field] = plain;
    const kept = headerFooterRunsForSegment(plain, runs);
    if (kept) out[runsField] = kept;
    // A picture lives where its section's &[Picture] code is: a section
    // whose code was deleted no longer holds it.
    if (headerFooterTextHasPicture(plain)) {
      const picture = normalizeHeaderFooterPicture(r[pictureField]);
      if (picture) out[pictureField] = picture;
    }
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

/** The per-page values substituted into header/footer field codes. */
export interface HeaderFooterContext {
  /** 1-based current page number. */
  pageNumber: number;
  /** How a page number field prints: a whole-workbook print numbers its
   * pages through the job ({@link markJobPageNumber}), the window's
   * numbers being the sheet's own. Plain digits when absent. */
  pageNumberText?: (pageNumber: number) => string;
  /** Total page count. */
  totalPages: number;
  /** How the page count field prints: a whole-workbook print counts the
   * job's pages ({@link markJobPageCount}), the window's count being the
   * sheet's own. Plain digits when absent. */
  pagesText?: (totalPages: number) => string;
  /** Pre-formatted print date string (caller's locale). */
  date: string;
  /** Pre-formatted print time string. */
  time: string;
  /** The sheet (tab) name. */
  sheetName: string;
  /** The file (workbook / page) name. */
  fileName: string;
  /**
   * The folder the file lives in (Excel `&Z` → `&[Path]`, which Excel's File
   * Path button pairs with `&[File]` as `&[Path]&[File]`). A page lives in no
   * folder, so absent prints nothing and the pair prints the file name once —
   * the same text the PDF export prints.
   */
  filePath?: string;
}

/**
 * Substitute every Excel field code in a header/footer segment with its value
 * from {@link HeaderFooterContext}. Codes are matched case-insensitively (Excel
 * accepts `&[page]`/`&[PAGE]`) and ALL occurrences of each are replaced; an
 * escaped `&&` prints one literal `&` and `&[Page]±n` prints the offset page
 * number ({@link headerFooterPieces}). Unknown `&[…]` sequences are left
 * verbatim. Pure — no Date.now()/locale lookups happen
 * here; the caller supplies the already-formatted date/time so the result is
 * deterministic and unit-testable.
 */
export function substituteHeaderFooterTokens(
  text: string,
  ctx: HeaderFooterContext,
  options: {
    /**
     * What `&[Picture]` becomes: nothing (a picture is not text), or
     * {@link HEADER_FOOTER_PICTURE_MARK} for a renderer that draws the
     * section's picture at that place.
     */
    pictureMark?: string;
  } = {},
): string {
  if (typeof text !== "string" || text === "") return "";
  const map: Record<HeaderFooterFieldName, string> = {
    page: ctx.pageNumberText ? ctx.pageNumberText(ctx.pageNumber) : String(ctx.pageNumber),
    pages: ctx.pagesText ? ctx.pagesText(ctx.totalPages) : String(ctx.totalPages),
    date: ctx.date,
    time: ctx.time,
    tab: ctx.sheetName,
    file: ctx.fileName,
    path: ctx.filePath ?? "",
    picture: options.pictureMark ?? "",
  };
  // Read as Excel reads the section: `&&[Page]` is the literal "&[Page]",
  // `&[Page]+2` the page number plus 2.
  return headerFooterPieces(text)
    .map((piece) => {
      if (piece.kind === "text") return piece.text;
      if (piece.field === "page" && piece.offset) {
        const number = ctx.pageNumber + Number(piece.offset);
        return ctx.pageNumberText ? ctx.pageNumberText(number) : String(number);
      }
      return map[piece.field];
    })
    .join("");
}

/**
 * A page number of a sheet printed in a whole-workbook job, as the sheet's
 * own window numbers it: marked so the job renumbers it once the pages of
 * the sheets before it are known ({@link numberJobPages}). Windows Excel
 * numbers a job's pages straight through: a sheet whose First page number
 * is Auto goes on from the sheet before (R492-X10's second sheet prints
 * "Seite 11" after ten pages), one that states a number starts there; the
 * server's writer does the same.
 */
export function markJobPageNumber(pageNumber: number): string {
  return `\uE000${pageNumber}\uE001`;
}

/** Rewrites the marked page numbers ({@link markJobPageNumber}) under
 * `root` to their place in the job: the sheet's own number plus `before`
 * (the number its first page follows). */
export function numberJobPages(root: Node, before: number): void {
  const doc = root.ownerDocument ?? (root as Document);
  const walker = doc.createTreeWalker(root, 4 /* NodeFilter.SHOW_TEXT */);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node.nodeValue ?? "";
    if (!text.includes("\uE000")) continue;
    node.nodeValue = text.replace(/\uE000(-?\d+)\uE001/g, (_match, number: string) => String(Number(number) + before));
  }
}

/**
 * The page count of a sheet printed in a whole-workbook job, as the sheet's
 * own window counts it: marked so the job prints its own count once every
 * sheet is stacked ({@link countJobPages}). Windows Excel's &N counts the
 * job's pages, a sheet that states its first page number too (lookwin
 * probe pn: four sheets of 2 + 3 + 1 + 2 pages, the third from page 7,
 * print "N8" on every page); the server's writer does the same.
 */
export function markJobPageCount(totalPages: number): string {
  return `\uE002${totalPages}\uE003`;
}

/** Rewrites the marked page counts ({@link markJobPageCount}) under `root`
 * to the job's page count, or, `null`, to the sheet's own count (a streamed
 * print hands its pages on before the job's count is known). */
export function countJobPages(root: Node, total: number | null): void {
  const doc = root.ownerDocument ?? (root as Document);
  const walker = doc.createTreeWalker(root, 4 /* NodeFilter.SHOW_TEXT */);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node.nodeValue ?? "";
    if (!text.includes("\uE002")) continue;
    node.nodeValue = text.replace(/\uE002(\d+)\uE003/g, (_match, own: string) => String(total ?? Number(own)));
  }
}

// ── Range parsing helpers ─────────────────────────────────────────────────────

/**
 * Convert a column letter string ("A", "AB", "AAA") to a 0-based column index.
 * Returns -1 for invalid input.
 */
export function colLetterToIndex(col: string): number {
  if (!col || !/^[A-Z]+$/.test(col)) return -1;
  let idx = 0;
  for (let i = 0; i < col.length; i++) {
    idx = idx * 26 + (col.charCodeAt(i) - 64);
  }
  return idx - 1;
}

/**
 * Parse an A1-notation cell address ("B3") into { row, col } (0-based).
 * Returns null for invalid input.
 */
export function parseA1Cell(addr: string): { row: number; col: number } | null {
  const m = /^([A-Z]+)([0-9]+)$/.exec(addr.toUpperCase().trim());
  if (!m) return null;
  return boundedCell(m[1], m[2]);
}

/**
 * Parse an A1-notation range "A1:D20" → { r1, c1, r2, c2 } (0-based, inclusive).
 * Also handles pure row ranges "1:3" → { r1:0, c1:0, r2:2, c2:MAX_COL }
 * and pure col ranges "A:C" → { r1:0, c1:0, r2:MAX_ROW, c2:2 }.
 * Returns null for invalid input.
 */
export function parseA1Range(
  range: string,
): { r1: number; c1: number; r2: number; c2: number } | null {
  if (!range) return null;
  const s = range.trim().toUpperCase();
  const parts = s.split(":");
  if (parts.length !== 2) {
    // Single cell
    const cell = parseA1Cell(s);
    if (!cell) return null;
    return { r1: cell.row, c1: cell.col, r2: cell.row, c2: cell.col };
  }
  const [a, b] = parts;

  // Row range "1:3"
  if (/^[0-9]+$/.test(a) && /^[0-9]+$/.test(b)) {
    const rawR1 = Number(a) - 1;
    const rawR2 = Number(b) - 1;
    if (
      !Number.isSafeInteger(rawR1) ||
      !Number.isSafeInteger(rawR2) ||
      rawR1 < 0 ||
      rawR2 < 0 ||
      rawR1 > MAX_PARSE_ROW ||
      rawR2 > MAX_PARSE_ROW
    ) {
      return null;
    }
    const r1 = Math.min(rawR1, rawR2);
    const r2 = Math.max(rawR1, rawR2);
    return { r1, c1: 0, r2, c2: MAX_PARSE_COL };
  }

  // Col range "A:C"
  if (/^[A-Z]+$/.test(a) && /^[A-Z]+$/.test(b)) {
    const rawC1 = colLetterToIndex(a);
    const rawC2 = colLetterToIndex(b);
    if (
      rawC1 < 0 ||
      rawC2 < 0 ||
      rawC1 > MAX_PARSE_COL ||
      rawC2 > MAX_PARSE_COL
    ) {
      return null;
    }
    const c1 = Math.min(rawC1, rawC2);
    const c2 = Math.max(rawC1, rawC2);
    return { r1: 0, c1, r2: MAX_PARSE_ROW, c2 };
  }

  // Full A1 range "A1:D20"
  const from = parseA1Cell(a);
  const to = parseA1Cell(b);
  if (!from || !to) return null;
  return {
    r1: Math.min(from.row, to.row),
    c1: Math.min(from.col, to.col),
    r2: Math.max(from.row, to.row),
    c2: Math.max(from.col, to.col),
  };
}

// ── Pagination helper (pure, testable) ───────────────────────────────────────

export interface PageBreakResult {
  /** 0-based row index ranges for each page [startRow, endRow] inclusive. */
  rowRanges: Array<[number, number]>;
  /** 0-based col index ranges for each page (column pagination). */
  colRanges: Array<[number, number]>;
  /** Number of repeat-title rows prepended to every page (0 = none). */
  repeatRowCount: number;
  /** Number of repeat-title cols prepended to every page (0 = none). */
  repeatColCount: number;
}

export interface PaginateInput {
  /** Total data row count (0-based). */
  totalRows: number;
  /** Total data col count (0-based). */
  totalCols: number;
  /** Page height in points available for data rows. */
  dataAreaHeight: number;
  /** Page width in points available for data columns. */
  dataAreaWidth: number;
  /** Height per data row in points. */
  rowHeight: number;
  /** Width per data column in points. */
  colWidth: number;
  rowSizes?: readonly number[];
  colSizes?: readonly number[];
  /** The column positions hidden source columns stand right before, within
   * the printed range (paginateAxis ``hiddenLed``; the API's hidden_led_cols). */
  hiddenLedCols?: readonly number[];
  /** Hidden source columns stand past the last printed position, within the
   * range (a wide column's text spilling on over them): after a band of one
   * column wider than the page they print a blank page of their own
   * (paginateAxis ``hiddenTail``; the API's hidden_tail_cols). */
  hiddenTailCols?: boolean;
  /** Page Setup > Sheet > Row and column headings: the column-letter band's
   * height and the row-number column's width, in points at the same scale as
   * the rows and columns. They print inside the margins, so they take their
   * room from every page (0 or absent when the sheet does not print them). */
  headingHeight?: number;
  headingWidth?: number;
  /** Parsed printSettings for this sheet. */
  settings: PrintSettings;
}

function paginateAxis(
  total: number,
  repeatCount: number,
  capacity: number,
  fallbackSize: number,
  sizes: readonly number[] | undefined,
  manualBreaks: readonly number[] | undefined,
  hiddenLed?: ReadonlySet<number>,
  hiddenTail = false,
): Array<[number, number]> {
  const sizeAt = (index: number) => {
    const candidate = sizes?.[index];
    return typeof candidate === "number" && Number.isFinite(candidate) && candidate > 0
      ? candidate
      : fallbackSize;
  };
  let titleSize = 0;
  for (let index = 0; index < repeatCount; index += 1) titleSize += sizeAt(index);
  const available = Math.max(capacity - titleSize, fallbackSize);
  const forced = new Set(
    (manualBreaks ?? []).filter(
      (value) => Number.isSafeInteger(value) && value > repeatCount && value < total,
    ),
  );
  if (total <= repeatCount) return [[repeatCount, Math.max(repeatCount, total - 1)]];
  const ranges: Array<[number, number]> = [];
  let start = repeatCount;
  let overflowed = true;
  while (start < total) {
    // Hidden lines are zero-width members of Windows Excel's bands: ones that
    // open a band (the range's first, or after a band of one line wider than
    // the page) before a line wider than the page stand a band of their own,
    // its pages blank (lookwin R466-X78: A:B hidden before a 120-character C,
    // 8 blank pages). The API's paginate_axis ``hidden_led``.
    if (hiddenLed?.has(start) && overflowed && sizeAt(start) > available) ranges.push([start, start - 1]);
    let end = start;
    let used = 0;
    while (end < total) {
      if (end > start && forced.has(end)) break;
      const size = sizeAt(end);
      if (end > start && used + size > available) break;
      used += size;
      end += 1;
    }
    ranges.push([start, Math.max(start, end - 1)]);
    overflowed = used > available;
    start = Math.max(start + 1, end);
  }
  // Hidden lines the range runs on into past its last shown one, after a band
  // of one line wider than the page, stand a band of their own too, its page
  // blank (lookwin p223-lw w6h, R460-X20's Readme: A 742 pt wide, its text
  // past the 511 pt page, B:XFD hidden). The API's paginate_axis ``hidden_tail``.
  if (hiddenTail && overflowed && ranges.length) ranges.push([total, total - 1]);
  return ranges;
}
/**
 * Compute row/column page-break ranges for a sheet given its dimensions
 * and PrintSettings.  Pure and deterministic — suitable for unit testing
 * without a PDF canvas.
 *
 * The result lets the PDF renderer iterate pages:
 *   for each colRange, for each rowRange → emit a page.
 *
 * When repeatTitleRows is set, the first `repeatRowCount` rows are
 * re-emitted at the top of every page (reducing the available height
 * per page accordingly).  Same for repeatTitleCols on every column page.
 */
export function paginateSheet(input: PaginateInput): PageBreakResult {
  const {
    totalRows,
    totalCols,
    dataAreaHeight,
    dataAreaWidth,
    rowHeight,
    colWidth,
    rowSizes,
    colSizes,
    settings,
  } = input;
  const headingHeight = Number.isFinite(input.headingHeight) ? Math.max(0, input.headingHeight!) : 0;
  const headingWidth = Number.isFinite(input.headingWidth) ? Math.max(0, input.headingWidth!) : 0;

  // 1. Repeat-title counts
  let repeatRowCount = 0;
  let repeatColCount = 0;

  if (settings.repeatTitleRows) {
    const r = parseA1Range(settings.repeatTitleRows);
    if (r) repeatRowCount = Math.min(r.r2 - r.r1 + 1, totalRows);
  }
  if (settings.repeatTitleCols) {
    const r = parseA1Range(settings.repeatTitleCols);
    if (r) repeatColCount = Math.min(r.c2 - r.c1 + 1, totalCols);
  }

  const spare = printPageSparePoints(settings.scaling);
  const rowRanges = paginateAxis(
    totalRows,
    repeatRowCount,
    dataAreaHeight - headingHeight - printPageRowSparePoints(settings.scaling),
    rowHeight,
    rowSizes,
    printManualBreaksApply(settings.scaling, "row") ? settings.manualRowBreaks : undefined,
  );
  const colRanges = paginateAxis(
    totalCols,
    repeatColCount,
    dataAreaWidth - headingWidth - spare,
    colWidth,
    colSizes,
    printManualBreaksApply(settings.scaling, "col") ? settings.manualColBreaks : undefined,
    input.hiddenLedCols?.length ? new Set(input.hiddenLedCols) : undefined,
    input.hiddenTailCols === true,
  );

  return { rowRanges, colRanges, repeatRowCount, repeatColCount };
}

/** The room Windows Excel leaves free at a page's right and bottom edges
 * before a column or row goes on to the next page: 7 printer dots at 600 dpi
 * (0.84 pt), at any percent scale (lookwin probes p34/p34b: a column that
 * leaves 7 dots fits, one that leaves 6.5 starts a page, at 100% and 200%,
 * bordered or not; rows the same within half a dot). Fit To keeps a dot
 * more (printPageSparePoints). */
export const PRINT_PAGE_SPARE_POINTS = (7 * 72) / 600;

/** Fit To's room: its scale is the largest whole percent whose pages leave
 * more than 8 printer dots free (lookwin probes p50-lw/p50b-lw, Letter,
 * Calibri 11: rows at 40/60/80% and columns at 40/60% print one percent
 * smaller with 8 dots to spare and at the percent with 8.5; R101-X17's 48%
 * left 4.36 and Excel chose 47%). Between the two: 8.25 dots. */
export const PRINT_FIT_SPARE_POINTS = (8.25 * 72) / 600;

export function printPageSparePoints(scaling: PrintScaling | undefined): number {
  return !scaling || scaling.kind === "percent" ? PRINT_PAGE_SPARE_POINTS : PRINT_FIT_SPARE_POINTS;
}

/** A row wants more than PRINT_PAGE_SPARE_POINTS free, a column that much
 * will do (lookwin probe p34b: a row leaving exactly 7 dots starts a page, one
 * leaving 7.5 fits; a column leaving 7 fits. R467-X11's first page, 15.24 +
 * 46 x 14.52 = 683.16 of 684 pt, breaks a row sooner, as Excel): the sliver
 * past 7 dots a row's room loses at a percent scale; the API's
 * PRINT_ROW_SPARE_EXCESS_POINTS. */
export const PRINT_ROW_SPARE_EXCESS_POINTS = 72 / 600 / 64;

export function printPageRowSparePoints(scaling: PrintScaling | undefined): number {
  const spare = printPageSparePoints(scaling);
  return !scaling || scaling.kind === "percent" ? spare + PRINT_ROW_SPARE_EXCESS_POINTS : spare;
}

/** A row's printed height in points: Windows Excel prints every row a whole
 * number of printer dots (600 dpi). At 100% that is its screen pixels times
 * the Normal font's printed over screen default row, plus an eighth of a dot,
 * floored (lookwin probe p72-lw: 1,625 rows of 10-70 px over 28 Normal fonts,
 * every one; Calibri 11's 14 px rows print 84 dots, 17 px 102, 18 px 109 --
 * rounding gave 85 and 103, and R79-X23's 72-row pages held 71); a scale
 * then takes its 100% dots, rounded (lookwin probe p41: 123-dot Arial 12 rows
 * print 58 dots at 47%, 121-dot Calibri 11 rows 57). The API's
 * printed_row_points (office_rows.py) is the same rule; office-row-rules.cases.json
 * holds cases both answer. */
export function printedRowPoints(size: number, scale: number): number {
  const dots = Math.floor((size * 600) / 72 + PRINTED_ROW_DOT_OFFSET + 1e-9);
  return dots > 0 ? (Math.max(1, Math.floor(dots * scale + 0.5)) * 72) / 600 : 0;
}
export const PRINTED_ROW_DOT_OFFSET = 0.125;

/** A column's printed width in points: whole printer dots too, its 100% dots
 * (rounded) times the scale, rounded (lookwin probe p50b-lw: 423-dot columns
 * print 343 dots at 81%, 254 at 60%, 169 at 40%), so a Fit To scale is judged
 * on the widths the page actually takes. The API's printed_col_points is the same. */
export function printedColumnPoints(size: number, scale: number): number {
  const dots = Math.floor((size * 600) / 72 + 0.5);
  return dots > 0 ? (Math.max(1, Math.floor(dots * scale + 0.5)) * 72) / 600 : 0;
}

/**
 * Whether an axis's manual page breaks print: Fit To ignores the breaks of an
 * axis it fits and keeps those of an axis it leaves automatic (Windows Excel,
 * lookwin probe p32: Fit To 0 wide x 1 tall breaks at the manual column breaks
 * and ignores the row ones; 1 wide x 0 tall the other way round; 2 x 2 ignores
 * both). A scale in percent keeps both.
 */
export function printManualBreaksApply(scaling: PrintScaling | undefined, axis: "row" | "col"): boolean {
  if (!scaling || scaling.kind === "percent") return true;
  if (scaling.kind === "fitToWidth") return axis === "row";
  if ("widthPages" in scaling) return (axis === "row" ? scaling.heightPages : scaling.widthPages) === 0;
  return false;
}

// ── Page Break Preview (pure, testable) ──────────────────────────────────────

export interface PageBreakLines {
  /**
   * 0-based row indices at which a page break is drawn (a horizontal line drawn
   * ABOVE this row). Sorted ascending, de-duplicated. Each entry < totalRows.
   */
  rowBreaks: number[];
  /**
   * 0-based column indices at which a page break is drawn (a vertical line drawn
   * to the LEFT of this column). Sorted ascending, de-duplicated. Each < totalCols.
   */
  colBreaks: number[];
  /** Which row breaks are MANUAL (user-set) vs automatic — same index basis. */
  manualRows: Set<number>;
  /** Which col breaks are MANUAL (user-set) vs automatic. */
  manualCols: Set<number>;
}

/**
 * Compute the page-break lines to draw in Page Break Preview, combining the
 * AUTOMATIC pagination (from {@link paginateSheet}) with the sheet's MANUAL
 * breaks. A manual break overrides the automatic flow downstream of it the way
 * Excel does — once a page is forced to start at a manual break, the subsequent
 * automatic breaks are recomputed from there — but for a faithful-enough preview
 * we draw the UNION of automatic break starts and manual break indices, marking
 * which are manual so the renderer can style them (manual = solid, auto = dashed)
 * and only let manual ones be dragged/removed.
 *
 * Pure + deterministic: it derives the automatic break starts from the same
 * `paginateSheet` ranges the PDF exporter uses, so the preview lines match where
 * the export actually paginates (absent manual overrides). A break at index 0 is
 * never emitted (the sheet's top/left edge is not a "break").
 */
export function computePageBreakLines(input: PaginateInput): PageBreakLines {
  const pagination = paginateSheet(input);
  const totalRows = Math.max(0, Math.floor(input.totalRows));
  const totalCols = Math.max(0, Math.floor(input.totalCols));

  // Automatic break = the START row/col of every page after the first.
  const autoRows = new Set<number>();
  for (let i = 1; i < pagination.rowRanges.length; i += 1) {
    const start = pagination.rowRanges[i][0];
    if (start > 0 && start < totalRows) autoRows.add(start);
  }
  const autoCols = new Set<number>();
  for (let i = 1; i < pagination.colRanges.length; i += 1) {
    const start = pagination.colRanges[i][0];
    if (start > 0 && start < totalCols) autoCols.add(start);
  }

  const manualRows = new Set<number>();
  for (const b of (printManualBreaksApply(input.settings.scaling, "row") ? input.settings.manualRowBreaks : []) ?? []) {
    if (b > 0 && b < totalRows) manualRows.add(b);
  }
  const manualCols = new Set<number>();
  for (const b of (printManualBreaksApply(input.settings.scaling, "col") ? input.settings.manualColBreaks : []) ?? []) {
    if (b > 0 && b < totalCols) manualCols.add(b);
  }

  const rowBreaks = Array.from(new Set([...autoRows, ...manualRows])).sort(
    (a, b) => a - b,
  );
  const colBreaks = Array.from(new Set([...autoCols, ...manualCols])).sort(
    (a, b) => a - b,
  );
  return { rowBreaks, colBreaks, manualRows, manualCols };
}

/**
 * Toggle a manual page break at `index` in a break array (Excel "Insert/Remove
 * Page Break"): if the index is already present it is removed, otherwise it is
 * added. The result is de-duplicated, sorted, and capped — i.e. run through
 * {@link normalizeBreakArray} — so a break "before index 0" can never be added.
 * Returns a fresh array; the input is not mutated.
 */
export function toggleManualBreak(breaks: readonly number[], index: number): number[] {
  const has = breaks.includes(index);
  const next = has ? breaks.filter((b) => b !== index) : [...breaks, index];
  return normalizeBreakArray(next);
}

/**
 * Move a manual page break from `fromIndex` to `toIndex` (dragging a break line
 * in Page Break Preview). The old index is removed and the new one added; if the
 * break wasn't present at `fromIndex` it is still added at `toIndex` (a drag that
 * started on an AUTO break "materializes" it as a manual break there, matching
 * Excel — dragging an automatic break converts it to a manual one). Normalized
 * (deduped/sorted/capped, ≥1). Returns a fresh array.
 */
export function moveManualBreak(
  breaks: readonly number[],
  fromIndex: number,
  toIndex: number,
): number[] {
  const without = breaks.filter((b) => b !== fromIndex);
  return normalizeBreakArray([...without, toIndex]);
}
