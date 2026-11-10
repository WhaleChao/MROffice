/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The characters of a text box as the native text-box writer draws them:
 * each run of its text with its own face, size, colour, bold, italic,
 * underline, strikethrough and superscript/subscript (WPS and Word format
 * part of a text box's text this way).
 *
 * One model end to end. The editors keep `PdfTextStyleRun`s on the box
 * (text-style-runs.ts: a run states only what differs from the box); the
 * writer is sent "drawn runs", every field resolved against the box, in the
 * `^` field of the box's directive (native-text-box-roundtrip.ts). The API
 * reads the same field (`_decode_managed_text_box_runs` in
 * native_edit_primitives.py) and draws each run in its face:
 *
 * - a portable class (`f`: sans / serif / mono): the base-14 face for the
 *   characters Windows-1252 carries, the embedded Noto Sans SC subset for
 *   the East Asian ones (as a box without runs);
 * - an Office family (`o`, a lib/office-fonts catalog name): the served open
 *   substitute's own font program (the woff2 slice files the browser draws
 *   it with), embedded as a subset; East Asian characters the substitute
 *   has no glyph for fall back to the embedded Noto Sans SC, as Word and WPS
 *   set them in the East Asian face. A character neither has fails closed.
 *
 * A box with runs is always written with its row plan (the V5 portable
 * layout): the window measures each row in the same font programs the writer
 * embeds (Office substitutes by their own served faces; the portable classes
 * by Arimo / Tinos / Cousine, metric-compatible with Helvetica / Times /
 * Courier) with kerning, ligatures and CJK punctuation trimming off, so a row
 * measures as the sum of its glyphs' advances on both sides; the writer
 * re-measures every row and refuses one that disagrees. Rows are as tall as
 * their largest run (Word's single spacing): `pdfTextBoxRunRowGeometry`.
 *
 * A run-less box, or one whose runs all draw as the box itself does, sends
 * no `^` field: its directive is byte-for-byte the legacy one.
 */
import { isPdfBundledFaceExtraCharacter } from "./bundled-font-coverage";
import { officeFontSubstitution, officeFontSubstitutes, ensureOfficeFontFaces } from "@mrnote/core/web/lib/office-fonts/substitutes.ts";
import { FONT_STACK, needsEastAsianFace, tryClassifyFontName } from "./font-resolve";
import { ensurePdfEastAsianFace, PDF_EAST_ASIAN_CSS_FAMILY } from "./east-asian-face";
import {
  PDF_TEXT_SCRIPT_SIZE_RATIO,
  PDF_TEXT_SUBSCRIPT_DROP,
  PDF_TEXT_SUPERSCRIPT_RISE,
} from "@mrnote/core/web/lib/pdf/text-box-typography.ts";
import { normalizePdfTextStyleRuns, type PdfTextStyleRun } from "@mrnote/core/web/lib/pdf/text-style-runs.ts";
import { pdfPageFontCssFamily, readPdfPageFontFamilyToken } from "./page-fonts";
import { pdfDraftItalic, type PdfDraftItalicPresentation } from "./draft-italic";

export type PdfTextBoxFontClass = "serif" | "sans" | "mono";
export type PdfTextBoxScript = "superscript" | "subscript";

/** One run as the writer draws it: every field resolved against the box. */
export interface PdfTextBoxDrawnRun {
  readonly text: string;
  /** The portable class: the base-14 face (and the class an export falls back to). */
  readonly font: PdfTextBoxFontClass;
  /** An Office family whose served substitute the writer embeds. */
  readonly officeFamily?: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: boolean;
  readonly strikethrough: boolean;
  readonly script?: PdfTextBoxScript;
  /** Font size, a fraction of the page height (as the box's `fontSize`). */
  readonly fontSize: number;
  /** #RRGGBB, upper case. */
  readonly color: string;
}

/** The box fields a drawn run inherits from. */
export interface PdfTextBoxRunsSource {
  text?: string;
  fontSize?: number;
  color?: string;
  fontFamily?: string;
  sourceFontName?: string;
  officeFamily?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  textScript?: PdfTextBoxScript;
  runs?: readonly PdfTextStyleRun[] | unknown;
}

/** `DEFAULT_NEW_BOX_FONT_SIZE` (text-box-create.ts), kept a leaf import. */
const DEFAULT_BOX_FONT_SIZE = 0.018;
export const MAX_PDF_TEXT_BOX_RUNS = 5_000;
/** The `^` field's bound (its base64url JSON), well inside the directive's. */
export const MAX_PDF_TEXT_BOX_RUNS_FIELD_CHARS = 262_144;
const HEX = /^#[0-9A-Fa-f]{6}$/u;

const WIN_ANSI_EXTRA_CODE_POINTS = new Set([
  0x0152, 0x0153, 0x0160, 0x0161, 0x0178, 0x017d, 0x017e, 0x0192, 0x02c6,
  0x02dc, 0x2013, 0x2014, 0x2018, 0x2019, 0x201a, 0x201c, 0x201d, 0x201e,
  0x2020, 0x2021, 0x2022, 0x2026, 0x2030, 0x2039, 0x203a, 0x20ac, 0x2122,
]);

/** A character a base-14 face draws (Windows-1252). */
export function isPdfWinAnsiCharacter(character: string): boolean {
  const codePoint = character.codePointAt(0);
  return (
    codePoint !== undefined &&
    (codePoint <= 0x7f ||
      (codePoint >= 0xa0 && codePoint <= 0xff) ||
      WIN_ANSI_EXTRA_CODE_POINTS.has(codePoint))
  );
}

/** A character of the ranges the embedded Noto Sans SC subset covers. */
export function isPdfBundledCjkCharacter(character: string): boolean {
  const codePoint = character.codePointAt(0);
  if (codePoint === undefined || codePoint > 0xffff) return false;
  return (
    (codePoint >= 0x2e80 && codePoint <= 0x2fff) ||
    (codePoint >= 0x3000 && codePoint <= 0x318f) ||
    (codePoint >= 0x31a0 && codePoint <= 0x31ff) ||
    (codePoint >= 0x3400 && codePoint <= 0x4dbf) ||
    (codePoint >= 0x4e00 && codePoint <= 0x9fff) ||
    (codePoint >= 0xac00 && codePoint <= 0xd7af) ||
    (codePoint >= 0xf900 && codePoint <= 0xfaff) ||
    (codePoint >= 0xff00 && codePoint <= 0xffef) ||
    // The same face's Greek, Cyrillic, math and arrows: the writer sets
    // every character it has a glyph for (a formula line went to page data
    
    isPdfBundledFaceExtraCharacter(codePoint)
  );
}

/**
 * The catalog's own spelling of an Office family the writer can embed (one
 * with a served substitute), or null. "calibri" and "Calibri" are one name;
 * a local font found by Browse is not an Office family and has no program
 * the server holds.
 */
export function pdfTextBoxOfficeFamily(name: string | null | undefined): string | null {
  if (typeof name !== "string" || !name.trim() || name.length > 128) return null;
  // A font from the reader's computer the page keeps (page-fonts.ts): its
  // name pins the face, which the writer draws from the page's copy.
  if (readPdfPageFontFamilyToken(name)) return name;
  const entry = officeFontSubstitution(name);
  return entry && entry.substitutes.length > 0 ? entry.office : null;
}

function upperColor(value: unknown, fallback: string): string | null {
  if (value === undefined || value === null || value === "") return fallback;
  return typeof value === "string" && HEX.test(value) ? value.toUpperCase() : null;
}

function sameDrawnStyle(a: PdfTextBoxDrawnRun, b: PdfTextBoxDrawnRun): boolean {
  return (
    a.font === b.font &&
    a.officeFamily === b.officeFamily &&
    a.bold === b.bold &&
    a.italic === b.italic &&
    a.underline === b.underline &&
    a.strikethrough === b.strikethrough &&
    a.script === b.script &&
    Math.abs(a.fontSize - b.fontSize) < 1e-9 &&
    a.color === b.color
  );
}

/**
 * The box's text as drawn runs, adjacent runs of one style merged; null when
 * the box cannot be drawn so (runs that do not spell its text, an Office
 * family the catalog has no substitute for, an invalid size or colour).
 */
export function resolvePdfTextBoxDrawnRuns(
  box: PdfTextBoxRunsSource,
): PdfTextBoxDrawnRun[] | null {
  const text = typeof box.text === "string" ? box.text : "";
  if (!text) return null;
  const boxSize = box.fontSize === undefined ? DEFAULT_BOX_FONT_SIZE : Number(box.fontSize);
  const boxColor = upperColor(box.color, "#000000");
  if (!Number.isFinite(boxSize) || boxSize <= 0 || boxSize > 1 || !boxColor) return null;
  const boxClass =
    tryClassifyFontName(box.sourceFontName) ?? tryClassifyFontName(box.fontFamily) ?? "sans";
  const boxOffice = box.officeFamily === undefined ? undefined : pdfTextBoxOfficeFamily(box.officeFamily);
  if (boxOffice === null) return null;
  const boxScript =
    box.textScript === "superscript" || box.textScript === "subscript" ? box.textScript : undefined;
  const runs = normalizePdfTextStyleRuns(box.runs, { expectedText: text });
  if (Array.isArray(box.runs) && box.runs.length > 0 && !runs) return null;
  const out: PdfTextBoxDrawnRun[] = [];
  for (const run of runs ?? [{ text }]) {
    // A run that names a face of its own is set in it; one that names none
    // is set in the box's (its Office family, else its class).
    const ownFace =
      run.officeFamily !== undefined || run.fontFamily !== undefined || run.sourceFontName !== undefined;
    const office = ownFace
      ? run.officeFamily === undefined ? undefined : pdfTextBoxOfficeFamily(run.officeFamily)
      : boxOffice;
    if (office === null) return null;
    const size = run.fontSize === undefined ? boxSize : Number(run.fontSize);
    const color = run.color === undefined ? boxColor : upperColor(run.color, boxColor);
    if (!Number.isFinite(size) || size <= 0 || size > 1 || !color) return null;
    const script = run.script ?? boxScript;
    const drawn: PdfTextBoxDrawnRun = {
      text: run.text,
      font: ownFace
        ? tryClassifyFontName(run.sourceFontName) ?? tryClassifyFontName(run.fontFamily) ??
          tryClassifyFontName(office) ?? boxClass
        : boxClass,
      ...(office ? { officeFamily: office } : {}),
      bold: run.bold ?? Boolean(box.bold),
      italic: run.italic ?? Boolean(box.italic),
      underline: run.underline ?? Boolean(box.underline),
      strikethrough: run.strikethrough ?? Boolean(box.strikethrough),
      ...(script ? { script } : {}),
      fontSize: size,
      color,
    };
    const previous = out.at(-1);
    if (previous && sameDrawnStyle(previous, drawn)) {
      out[out.length - 1] = { ...previous, text: previous.text + drawn.text };
    } else {
      out.push(drawn);
    }
  }
  return out.length > 0 && out.length <= MAX_PDF_TEXT_BOX_RUNS ? out : null;
}

/**
 * True when the box needs the run writer: its text is drawn in more than
 * one style, or in an Office family. A box drawn wholly as its own legacy
 * fields describe (class face, bold, italic, decorations, size, colour, the
 * box's own script) keeps the legacy directive, byte for byte.
 */
export function pdfTextBoxDrawnRunsNeedRunWriter(
  box: PdfTextBoxRunsSource,
  drawn: readonly PdfTextBoxDrawnRun[],
): boolean {
  if (drawn.length !== 1) return true;
  const only = drawn[0];
  const boxScript =
    box.textScript === "superscript" || box.textScript === "subscript" ? box.textScript : undefined;
  const legacy: PdfTextBoxDrawnRun = {
    text: only.text,
    font: tryClassifyFontName(box.sourceFontName) ?? tryClassifyFontName(box.fontFamily) ?? "sans",
    bold: Boolean(box.bold),
    italic: Boolean(box.italic),
    underline: Boolean(box.underline),
    strikethrough: Boolean(box.strikethrough),
    ...(boxScript ? { script: boxScript } : {}),
    fontSize: box.fontSize === undefined ? DEFAULT_BOX_FONT_SIZE : Number(box.fontSize),
    color: upperColor(box.color, "#000000") ?? "",
  };
  return !sameDrawnStyle(only, legacy);
}

/** Whether the writer draws every character of the runs: a class run only
 * Windows-1252 and the bundled CJK ranges (the box rule without runs); an
 * Office run any character but an astral one (its substitute's glyphs, else
 * the bundled CJK face, are checked by the writer against the programs). */
export function pdfTextBoxDrawnRunsPortable(runs: readonly PdfTextBoxDrawnRun[]): boolean {
  return runs.every((run) =>
    [...run.text].every((character) =>
      run.officeFamily
        ? character.length === 1
        : isPdfWinAnsiCharacter(character) || isPdfBundledCjkCharacter(character)));
}

function numberToken(value: number): number {
  return Number(value.toFixed(6));
}

function base64Url(text: string): string {
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return globalThis.btoa(binary).replaceAll("+", "-").replaceAll("/", "_").replace(/=+$/u, "");
}

/**
 * The directive's `^` field: `^` + base64url of
 * `{"v":1,"r":[{"n":5,"f":"sans","b":1,"i":0,"d":0,"z":0.018,"c":"FF0000"},…]}`.
 * `n` is the run's length in UTF-16 units (they spell the text token's text
 * in order), `d` a 2-bit decoration code (1 underline, 2 strikethrough), `z`
 * the size as a fraction of the page height, `o` an Office family and `x`
 * "sup" / "sub". `^` is outside base64url and `~`, so the field is lifted
 * out of a directive unambiguously. Null when it cannot be encoded.
 */
export function encodePdfTextBoxRunsField(runs: readonly PdfTextBoxDrawnRun[]): string | null {
  if (runs.length === 0 || runs.length > MAX_PDF_TEXT_BOX_RUNS) return null;
  const json = {
    v: 1,
    r: runs.map((run) => ({
      n: run.text.length,
      f: run.font,
      b: run.bold ? 1 : 0,
      i: run.italic ? 1 : 0,
      d: (run.underline ? 1 : 0) | (run.strikethrough ? 2 : 0),
      z: numberToken(run.fontSize),
      c: run.color.slice(1).toUpperCase(),
      ...(run.officeFamily ? officeFamilyFields(run.officeFamily) : {}),
      ...(run.script ? { x: run.script === "superscript" ? "sup" : "sub" } : {}),
    })),
  };
  if (json.r.some((run) => !(run.n > 0) || !(run.z > 0))) return null;
  const field = `^${base64Url(JSON.stringify(json))}`;
  return field.length <= MAX_PDF_TEXT_BOX_RUNS_FIELD_CHARS ? field : null;
}

/** A run's family as written: a page font's family and its face's digest
 * (`p`), which the server draws the run in (app/services/pdf_page_fonts.py). */
function officeFamilyFields(officeFamily: string): { o: string; p?: string } {
  const page = readPdfPageFontFamilyToken(officeFamily);
  return page ? { o: page.family, p: page.sha256 } : { o: officeFamily };
}

const RUN_KEYS = new Set(["n", "f", "b", "i", "d", "z", "c", "o", "x", "p"]);
const REQUIRED_RUN_KEYS = ["n", "f", "b", "i", "d", "z", "c"] as const;

/** The drawn runs of a `^` field for `text`, or null when it is not one. */
export function decodePdfTextBoxRunsField(field: string, text: string): PdfTextBoxDrawnRun[] | null {
  if (
    !/^\^[A-Za-z0-9_-]+$/u.test(field) ||
    field.length > MAX_PDF_TEXT_BOX_RUNS_FIELD_CHARS
  ) {
    return null;
  }
  try {
    const padded = field.slice(1).replaceAll("-", "+").replaceAll("_", "/");
    const binary = globalThis.atob(padded + "=".repeat((4 - (padded.length % 4)) % 4));
    const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
    if (base64Url(new TextDecoder("utf-8", { fatal: true }).decode(bytes)) !== field.slice(1)) return null;
    const json = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(bytes)) as unknown;
    if (!json || typeof json !== "object" || Array.isArray(json)) return null;
    const record = json as Record<string, unknown>;
    if (Object.keys(record).some((key) => key !== "v" && key !== "r") || record.v !== 1) return null;
    const rows = record.r;
    if (!Array.isArray(rows) || rows.length === 0 || rows.length > MAX_PDF_TEXT_BOX_RUNS) return null;
    const out: PdfTextBoxDrawnRun[] = [];
    let offset = 0;
    for (const row of rows) {
      if (!row || typeof row !== "object" || Array.isArray(row)) return null;
      const run = row as Record<string, unknown>;
      if (Object.keys(run).some((key) => !RUN_KEYS.has(key)) ||
        REQUIRED_RUN_KEYS.some((key) => !(key in run))) return null;
      const length = run.n;
      if (typeof length !== "number" || !Number.isSafeInteger(length) || length < 1) return null;
      const end = offset + length;
      if (end > text.length || splitsSurrogate(text, end)) return null;
      if (run.f !== "sans" && run.f !== "serif" && run.f !== "mono") return null;
      if ((run.b !== 0 && run.b !== 1) || (run.i !== 0 && run.i !== 1)) return null;
      if (typeof run.d !== "number" || !Number.isInteger(run.d) || run.d < 0 || run.d > 3) return null;
      if (typeof run.z !== "number" || !Number.isFinite(run.z) || run.z <= 0 || run.z > 1) return null;
      if (typeof run.c !== "string" || !/^[0-9A-F]{6}$/u.test(run.c)) return null;
      let officeFamily: string | undefined;
      if (run.p !== undefined) {
        const token = typeof run.o === "string" && typeof run.p === "string" ? `${run.o}#p=${run.p}` : "";
        if (!readPdfPageFontFamilyToken(token)) return null;
        officeFamily = token;
      } else if (run.o !== undefined) {
        if (typeof run.o !== "string" || pdfTextBoxOfficeFamily(run.o) !== run.o) return null;
        officeFamily = run.o;
      }
      if (run.x !== undefined && run.x !== "sup" && run.x !== "sub") return null;
      out.push({
        text: text.slice(offset, end),
        font: run.f,
        ...(officeFamily ? { officeFamily } : {}),
        bold: run.b === 1,
        italic: run.i === 1,
        underline: (run.d & 1) === 1,
        strikethrough: (run.d & 2) === 2,
        ...(run.x ? { script: run.x === "sup" ? "superscript" : "subscript" } as const : {}),
        fontSize: run.z,
        color: `#${run.c}`,
      });
      offset = end;
    }
    return offset === text.length ? out : null;
  } catch {
    return null;
  }
}

function splitsSurrogate(text: string, offset: number): boolean {
  if (offset <= 0 || offset >= text.length) return false;
  const before = text.charCodeAt(offset - 1);
  const after = text.charCodeAt(offset);
  return before >= 0xd800 && before <= 0xdbff && after >= 0xdc00 && after <= 0xdfff;
}

/** The editors' runs that reopen drawn runs (a box read back from its PDF). */
export function pdfTextStyleRunsFromDrawnRuns(runs: readonly PdfTextBoxDrawnRun[]): PdfTextStyleRun[] {
  return runs.map((run) => ({
    text: run.text,
    fontSize: run.fontSize,
    color: run.color.toLowerCase(),
    bold: run.bold,
    italic: run.italic,
    underline: run.underline,
    strikethrough: run.strikethrough,
    ...(run.script ? { script: run.script } : {}),
    ...(run.officeFamily
      ? { officeFamily: run.officeFamily, sourceFontName: run.officeFamily }
      : { sourceFontName: PDF_TEXT_BOX_CLASS_SOURCE_FONT[run.font] }),
    fontFamily: pdfTextBoxRunCssFamily(run),
  }));
}

/** A portable class's base-14 face name, as the editors name it. */
export const PDF_TEXT_BOX_CLASS_SOURCE_FONT: Record<PdfTextBoxFontClass, string> = {
  sans: "Helvetica",
  serif: "Times-Roman",
  mono: "Courier",
};

/**
 * The served faces the window measures and draws a portable class in:
 * Arimo, Tinos and Cousine have Helvetica's, Times's and Courier's advance
 * widths (the Liberation metrics), and are the same files on every machine,
 * so a row measured here is the row the base-14 face draws.
 */
export const PDF_TEXT_BOX_CLASS_MEASURE_FAMILY: Record<PdfTextBoxFontClass, string> = {
  sans: "Arimo",
  serif: "Tinos",
  mono: "Cousine",
};

/**
 * The CSS font-family a drawn run is measured and shown in: an Office run's
 * served substitutes by their own names (never the alias, which tries a
 * local Office font first and would measure another program), a class run's
 * metric twin; then the East Asian face the writer falls back to.
 */
export function pdfTextBoxRunCssFamily(
  run: Pick<PdfTextBoxDrawnRun, "font" | "officeFamily">,
  text?: string,
): string {
  const generic = run.font === "serif" ? "serif" : run.font === "mono" ? "monospace" : "sans-serif";
  let faces: string[];
  const page = readPdfPageFontFamilyToken(run.officeFamily);
  if (page) {
    // The page's copy of the face (registered by the window), else the
    // reader's own installed family of that name.
    faces = [`"${pdfPageFontCssFamily(page)}"`, `"${page.family.replace(/["\\]/g, "")}"`];
  } else if (run.officeFamily) {
    const substitutes = officeFontSubstitutes(run.officeFamily);
    for (const substitute of substitutes) ensureOfficeFontFaces(substitute);
    faces = substitutes.map((name) => `"${name.replace(/["\\]/g, "")}"`);
  } else {
    const twin = PDF_TEXT_BOX_CLASS_MEASURE_FAMILY[run.font];
    ensureOfficeFontFaces(twin);
    faces = [`"${twin}"`, ...FONT_STACK[run.font].split(",").map((name) => name.trim()).filter((name) => name !== generic)];
  }
  // What none of them has falls back to the face the writer falls back to
  // (the embedded Noto Sans SC), always named: a substitute that is itself a
  // Noto CJK family still lacks characters the embedded face has.
  if (text === undefined || needsEastAsianFace(text)) void ensurePdfEastAsianFace();
  return [...faces, `"${PDF_EAST_ASIAN_CSS_FAMILY}"`, generic].join(", ");
}

/** The CSS weight a drawn run asks its family for: an Office style-named
 * family ("Aptos Light") its own weight, bold 700. */
export function pdfTextBoxRunCssWeight(run: Pick<PdfTextBoxDrawnRun, "bold" | "officeFamily">): number {
  if (run.bold) return 700;
  return (run.officeFamily ? officeFontSubstitution(run.officeFamily)?.weight : undefined) ?? 400;
}

/** The size a run's glyphs are set at: a script run's is two thirds. */
export function pdfTextBoxRunGlyphSize(run: Pick<PdfTextBoxDrawnRun, "script">, fontSize: number): number {
  return run.script ? fontSize * PDF_TEXT_SCRIPT_SIZE_RATIO : fontSize;
}

/** How far a run's baseline is raised (negative: lowered), in its size. */
export function pdfTextBoxRunRise(run: Pick<PdfTextBoxDrawnRun, "script">, fontSize: number): number {
  return run.script === "superscript"
    ? PDF_TEXT_SUPERSCRIPT_RISE * fontSize
    : run.script === "subscript"
      ? -PDF_TEXT_SUBSCRIPT_DROP * fontSize
      : 0;
}

/** How a run's italic is drawn over its family: asked of a face that has
 * one, else upright glyphs sheared as the writer slants them. */
export function pdfTextBoxRunItalicPresentation(
  run: Pick<PdfTextBoxDrawnRun, "font" | "officeFamily" | "italic">,
  text?: string,
): PdfDraftItalicPresentation {
  return pdfDraftItalic(pdfTextBoxRunCssFamily(run, text), Boolean(run.italic));
}

/** The CSS `font` shorthand of a run at `fontSize` px. */
export function pdfTextBoxRunCssFont(
  run: Pick<PdfTextBoxDrawnRun, "font" | "officeFamily" | "bold" | "italic" | "script">,
  fontSize: number,
  text?: string,
): string {
  const family = pdfTextBoxRunCssFamily(run, text);
  const italic = pdfDraftItalic(family, Boolean(run.italic));
  return `${italic.fontStyle === "italic" ? "italic " : ""}${pdfTextBoxRunCssWeight(run)} ${pdfTextBoxRunGlyphSize(run, fontSize)}px ${family}`;
}

/**
 * The browser's own adjustments of glyph advances turned off, for a box set
 * in runs: kerning, ligatures and CJK punctuation trimming. The writer sets
 * each glyph at its font's advance; with these off the browser does too, so
 * a row's width is the same sum on both sides.
 */
export const PDF_TEXT_BOX_EXACT_ADVANCE_STYLE: Readonly<Record<string, string>> = Object.freeze({
  fontKerning: "none",
  fontVariantLigatures: "none",
  fontFeatureSettings: '"kern" 0, "liga" 0, "clig" 0, "calt" 0',
  textSpacingTrim: "space-all",
});

/** Apply `PDF_TEXT_BOX_EXACT_ADVANCE_STYLE` to an element. */
export function applyPdfTextBoxExactAdvanceStyle(element: HTMLElement): void {
  element.style.fontKerning = "none";
  element.style.fontVariantLigatures = "none";
  element.style.fontFeatureSettings = PDF_TEXT_BOX_EXACT_ADVANCE_STYLE.fontFeatureSettings;
  element.style.setProperty("text-spacing-trim", "space-all");
}

/** One row of a run box: its text range and its runs' sizes. */
export interface PdfTextBoxRunRow {
  readonly start: number;
  readonly end: number;
  /** The largest nominal size (px or pt) of the runs on the row. */
  readonly size: number;
}

/**
 * Word's single spacing for a box set in runs: each row is 1.2 times its
 * largest run tall (a superscript counts at its run's size), a blank line
 * left by two breaks as tall as the run it is in; the first baseline is one
 * row size below the top. Returns each row's top and baseline. The rows are
 * the browser's line breaks (`start`/`end` offsets in the text); `sizeAt`
 * gives the nominal size of the run holding an offset.
 */
export function pdfTextBoxRunRowGeometry(
  text: string,
  rows: readonly PdfTextBoxRunRow[],
  sizeAt: (offset: number) => number,
  lineHeightRatio = 1.2,
): { top: number; baseline: number }[] {
  const out: { top: number; baseline: number }[] = [];
  let top = 0;
  let cursor = 0;
  for (const row of rows) {
    // Every break before this row but the one ending the previous row (or
    // all of them before the first) stands for a blank line.
    let breaks = 0;
    const blanks: number[] = [];
    for (let offset = cursor; offset < row.start; offset += 1) {
      const character = text[offset];
      if (character === "\r" && text[offset + 1] === "\n") continue;
      if (character === "\n" || character === "\r") {
        if (breaks > 0 || out.length === 0) blanks.push(offset);
        breaks += 1;
      }
    }
    for (const offset of blanks) top += sizeAt(offset) * lineHeightRatio;
    out.push({ top, baseline: top + row.size });
    top += row.size * lineHeightRatio;
    cursor = row.end;
  }
  return out;
}

/** Why the PDF writer refused a box set in runs, read from the save's
 * patchback reason (`pdf_text_box_*` details of native_edit_primitives.py
 * and pdf_text_box_office_fonts.py); null for any other refusal. */
export type PdfTextBoxRunsRefusal =
  | { readonly kind: "glyph"; readonly family: string; readonly character: string }
  | { readonly kind: "font"; readonly family: string }
  | { readonly kind: "rows" }
  | { readonly kind: "plan" };

export function pdfTextBoxRunsRefusal(reason: unknown): PdfTextBoxRunsRefusal | null {
  if (typeof reason !== "string") return null;
  const glyph = /pdf_text_box_glyph_missing:(.+?):U\+([0-9A-F]{4,6})/u.exec(reason);
  if (glyph) {
    const code = Number.parseInt(glyph[2], 16);
    return { kind: "glyph", family: glyph[1], character: code <= 0x10ffff ? String.fromCodePoint(code) : "" };
  }
  const font = /pdf_text_box_font_unavailable:([^:\]]+)/u.exec(reason);
  if (font) return { kind: "font", family: font[1].trim() };
  if (reason.includes("pdf_text_box_rows_disagree")) return { kind: "rows" };
  if (reason.includes("pdf_text_box_runs_need_row_plan")) return { kind: "plan" };
  return null;
}
