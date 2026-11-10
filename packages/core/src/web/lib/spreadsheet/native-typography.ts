/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { PX_PER_PT } from "../office-units";
import type { CSSProperties } from "react";
import { officePrimaryFontFamily } from "../../components/office-ribbon/officeTypographyCatalog";
import { officeFaceName } from "../office-fonts/face-names";
import { ensureOfficeFauxBoldProbe, officeFamilyHasNoBold } from "../office-fonts/faux-bold";
import { ensureOfficeFontFaces, loadOfficeFamilySpace, officeFontDeclared, officeFontSubstitutes } from "../office-fonts/substitutes";
import { columnBaseWidthToPixels, columnFileWidthToPixels } from "./grid-sizing";
import { isSpreadsheetFauxBoldStroke, spreadsheetEngineIdealAdvances, spreadsheetFauxBoldCss, spreadsheetTextEngineOn } from "./spreadsheet-text-engine";
import { withScriptFallbacks } from "../office-fonts/script-fallbacks";
import { GENERIC_FAMILIES } from "../text-engine/face-index";
import { fontLoadReach } from "./font-load-reach";

/** SpreadsheetML widths are font-relative units, never CSS pixels.
 * Keep the source unit beside its legacy projection so merely opening a file
 * on a different display/font environment cannot rewrite the native package.
 * ISO 29500 CT_Col: round(width * MDW) on the target device's pixel grid.
 */
export interface NativeGridFont {
  family: string;
  sizePt: number;
  bold?: boolean;
  italic?: boolean;
  headingFamily?: string;
  bodyFamily?: string;
  themeRole?: "major" | "minor";
}
export interface NativeEditProfile { dpi: 72 | 96; digitWidth: number }
export interface NativeColumnWidth { width: number; pixels: number; profile?: NativeEditProfile }
export interface NativeRowHeight { heightPt: number; pixels: number; customHeight: boolean }
export interface NativeRowMetrics { version: 1; default: NativeRowHeight; rows: Record<string, NativeRowHeight> }
export interface NativeGrid {
  version: 1;
  font: NativeGridFont;
  metricsFont?: NativeGridFont;
  columns: Record<string, NativeColumnWidth>;
  defaultWidth?: number;
  baseColumnWidth: number;
  defaultPixels: number;
  sourceThemeFontId?: string;
  defaultProfile?: NativeEditProfile;
  rowMetrics?: NativeRowMetrics;
}
const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === "object" && !Array.isArray(value);
const finite = (value: unknown, min: number, max: number): value is number =>
  typeof value === "number" && Number.isFinite(value) && value >= min && value <= max;
/** Excel's widest column, 255 characters, stored with the cell padding. */
const MAX_NATIVE_COLUMN_WIDTH = 256;
export function normalizeNativeGrid(value: unknown): NativeGrid | undefined {
  if (!record(value) || value.version !== 1 || !record(value.font) || !record(value.columns) ||
      typeof value.font.family !== "string" || !value.font.family.trim() || value.font.family.length > 200 ||
      /[\u0000-\u001f<>;{}"'\\]/u.test(value.font.family) || !finite(value.font.sizePt, 1, 409) ||
      !finite(value.baseColumnWidth, 0, 255) || !finite(value.defaultPixels, 1, 4096) ||
      (value.defaultWidth !== undefined && !finite(value.defaultWidth, 0, 255))) return undefined;
  const metricsFont = value.metricsFont === undefined ? undefined : normalizeNativeGrid({...value, font: value.metricsFont, metricsFont: undefined})?.font;
  if (value.metricsFont !== undefined && !metricsFont) return undefined;
  const profile = value.defaultProfile;
  const validProfile = (profile: unknown): profile is NativeEditProfile => record(profile) &&
    (profile.dpi === 72 || profile.dpi === 96) && finite(profile.digitWidth,1,512) && Number.isInteger(profile.digitWidth);
  if (profile !== undefined && !validProfile(profile)) return undefined;
  const entries = Object.entries(value.columns);
  if (entries.length > 16384) return undefined;
  const columns: Record<string, NativeColumnWidth> = {};
  for (const key of ["headingFamily", "bodyFamily"] as const) {
    const family = value.font[key];
    if (family !== undefined && (typeof family !== "string" || !family.trim() || family.length > 200 || /[\u0000-\u001f<>;{}"'\\]/u.test(family))) return undefined;
  }
  for (const [index, entry] of entries) {
    if (!/^(0|[1-9]\d*)$/.test(index) || Number(index) > 16383 || !record(entry) ||
        !finite(entry.width, 0, MAX_NATIVE_COLUMN_WIDTH) || !finite(entry.pixels, 1, 4096)) return undefined;
    if (entry.profile !== undefined && !validProfile(entry.profile)) return undefined;
    columns[index] = { width: entry.width, pixels: entry.pixels, ...(validProfile(entry.profile) ? { profile: entry.profile } : {}) };
  }
  const rowMetrics = normalizeNativeRowMetrics(value.rowMetrics);
  if (value.rowMetrics !== undefined && !rowMetrics) return undefined;
  return { version: 1, font: { family: value.font.family, sizePt: value.font.sizePt,
    ...(typeof value.font.headingFamily === "string" ? { headingFamily: value.font.headingFamily } : {}),
    ...(typeof value.font.bodyFamily === "string" ? { bodyFamily: value.font.bodyFamily } : {}),
    ...(value.font.themeRole === "major" || value.font.themeRole === "minor" ? { themeRole: value.font.themeRole } : {}),
    ...(value.font.bold === true ? { bold: true } : {}), ...(value.font.italic === true ? { italic: true } : {}) },
    ...(metricsFont ? { metricsFont } : {}),
    columns, baseColumnWidth: value.baseColumnWidth, defaultPixels: value.defaultPixels,
    ...(typeof value.sourceThemeFontId === "string" && value.sourceThemeFontId.length <= 100 ? { sourceThemeFontId: value.sourceThemeFontId } : {}),
    ...(value.defaultWidth !== undefined ? { defaultWidth: value.defaultWidth } : {}),
    ...(validProfile(profile) ? { defaultProfile: profile } : {}),
    ...(rowMetrics ? {rowMetrics} : {}) };
}

export function normalizeNativeRowMetrics(value: unknown): NativeRowMetrics | undefined {
  const valid = (row: unknown): row is NativeRowHeight => record(row) && finite(row.heightPt,0.05,409.5) &&
    finite(row.pixels,1,4096) && typeof row.customHeight === "boolean";
  if (!record(value) || value.version !== 1 || !valid(value.default) || !record(value.rows) || Object.keys(value.rows).length > 200000) return undefined;
  const rows: Record<string,NativeRowHeight> = {};
  const copy = (row: NativeRowHeight) => ({heightPt:row.heightPt,pixels:row.pixels,customHeight:row.customHeight});
  for (const [key,row] of Object.entries(value.rows)) {
    if (!/^(0|[1-9]\d*)$/.test(key) || Number(key)>1048575 || !valid(row)) return undefined;
    rows[key]=copy(row);
  }
  return {version:1,default:copy(value.default),rows};
}

/** Pairing validates dimensions only; native write authority remains server-owned. */
export function nativeRowPixelPairMatches(metrics: NativeRowMetrics | undefined, index: string, pixels: unknown): boolean {
  const row = metrics?.rows[index];
  return Boolean(row && typeof pixels === "number" && pixels === row.pixels &&
    pixels === Math.max(1, Math.round(row.heightPt / .75)));
}

/** Pure view projection. Retain legacy saved pixels; only a matching pair may
 * use the original point height. Never feed this projection into autosave. */
export function resolveNativeRowHeights(metrics: NativeRowMetrics | undefined, heights: Record<string,number> | undefined, defaultHeight: number | undefined, fontRowHeight?: number) {
  if (!metrics) return undefined;
  const resolved = {...heights};
  for (const [key,row] of Object.entries(metrics.rows)) if (heights?.[key] === row.pixels) resolved[key]=row.heightPt / .75;
  const pairedDefault = defaultHeight === metrics.default.pixels;
  const automatic = pairedDefault && !metrics.default.customHeight && finite(fontRowHeight, .05/.75, 4096);
  return {heights:resolved,defaultHeight:automatic ? clampNativeAutomaticRowHeight(fontRowHeight) : pairedDefault ? metrics.default.heightPt/.75 : defaultHeight};
}

/** Excel's row limit applies to computed heights as well as authored heights. */
export function clampNativeAutomaticRowHeight(height: number): number {
  return Math.min(409.5 / .75, height);
}

/** Synchronize only explicit canonical dimension changes. Structural writers
 * remap the original pairs first; display/font measurement never calls this. */
export function syncNativeRowMetrics(grid: NativeGrid, heights: Record<string,number> | undefined, defaultHeight: number | undefined): NativeGrid {
  const source=grid.rowMetrics;
  if (!source) return grid;
  const rows:Record<string,NativeRowHeight>={};
  for (const [key,pixels] of Object.entries(heights??{})) {
    rows[key]=source.rows[key]?.pixels===pixels ? source.rows[key] : {heightPt:pixels*.75,pixels,customHeight:true};
  }
  const rowDefault=defaultHeight===undefined || defaultHeight===source.default.pixels ? source.default : {heightPt:defaultHeight*.75,pixels:defaultHeight,customHeight:true};
  return {...grid,rowMetrics:{version:1,default:rowDefault,rows}};
}

export function editNativeRowSizes(grid: NativeGrid, sizes: readonly {index:number;size:number}[]): NativeGrid {
  if (!grid.rowMetrics) return grid;
  const rows={...grid.rowMetrics.rows};
  for (const {index,size} of sizes) {
    const bounded=clampNativeRowHeight(size,20);
    rows[index]={heightPt:bounded*.75,pixels:Math.round(bounded),customHeight:true};
  }
  return {...grid,rowMetrics:{...grid.rowMetrics,rows}};
}

export function clampNativeRowHeight(value: unknown, fallback: number): number {
  const size=typeof value==="number" && Number.isFinite(value) ? value : Number.isFinite(fallback) ? fallback : 20;
  return Math.max(18,Math.min(512,size));
}

export function nativeRowCssClamp(value: unknown, fallback: number): number {
  return finite(value,.05/.75,4096) ? value : finite(fallback,.05/.75,4096) ? fallback : 20;
}
export function nativeWidthToPixels(width: number, digitWidth: number, dpi: number): number {
  return Math.max(1, columnFileWidthToPixels(width, digitWidth)) * 96 / dpi;
}
/** CT_Col reserves two device pixels of text padding on each side. Keep
 * that physical inset on the same device grid as native glyph advances. */
export function nativeCellPaddingPx(dpi: number): number {
  return 2 * 96 / dpi;
}
export function pixelsToNativeWidth(pixels: number, digitWidth: number, dpi: number): number {
  // Choose the smallest 1/256 native unit that roundtrips to this device pixel.
  const device = Math.max(1, Math.round(pixels * dpi / 96));
  return Math.max(0, Math.min(255, Math.ceil((device * 256 / digitWidth) - Math.floor(128 / digitWidth)) / 256));
}
export function nativeDefaultWidth(grid: NativeGrid, digitWidth: number, dpi: number) {
  return grid.defaultWidth === undefined
    ? columnBaseWidthToPixels(grid.baseColumnWidth, digitWidth, dpi) * 96 / dpi
    : nativeWidthToPixels(grid.defaultWidth, digitWidth, dpi);
}
export function nativeColumnCssClamp(value: unknown, fallback: number): number {
  return finite(value, 0, 4096) ? Math.max(1, value) : fallback;
}

export function editNativeColumns(grid: NativeGrid, sizes: readonly { index: number; size: number }[], digitWidth: number, dpi: number): NativeGrid {
  const columns = { ...grid.columns };
  for (const { index, size } of sizes) {
    const width = pixelsToNativeWidth(size, digitWidth, dpi);
    columns[index] = { width, pixels: nativeWidthToPixels(width,digitWidth,dpi), profile: { dpi: dpi as 72 | 96, digitWidth } };
  }
  return { ...grid, columns };
}

const faceAvailability = new Map<string, boolean>();
/** Bumped whenever a font finishes loading (the face and advance caches are
 * dropped then): text drawn before it may be measured anew. */
let typographyGeneration = 0;
export function nativeTypographyGeneration(): number {
  return typographyGeneration;
}
const typographyListeners = new Set<() => void>();
let fontLoadsWatched = false;
function watchFontLoads(): void {
  if (fontLoadsWatched || typeof document === "undefined" || !document.fonts) return;
  fontLoadsWatched = true;
  document.fonts.addEventListener("loadingdone", (event) => {
    // Only the families the load finished faces of can have changed; one
    
    // fallback slice moved the generation, and every measure keyed by it).
    const reach = fontLoadReach(event);
    let changed = !reach;
    if (!reach) faceAvailability.clear();
    else {
      for (const [family, available] of [...faceAvailability]) {
        if (!reach.families.has(family.replace(/^["']|["']$/g, "").trim().toLowerCase())) continue;
        faceAvailability.delete(family);
        if (!available) changed = true;
      }
    }
    if (!changed) return;
    typographyGeneration += 1;
    for (const listener of [...typographyListeners]) listener();
  });
}




export function subscribeNativeTypography(listener: () => void): () => void {
  watchFontLoads();
  typographyListeners.add(listener);
  return () => {
    typographyListeners.delete(listener);
  };
}
let availabilityContext: CanvasRenderingContext2D | null | undefined;
export function spreadsheetFontAvailable(family: string): boolean | undefined {
  if (typeof document === "undefined") return undefined;
  if (availabilityContext === undefined) {
    availabilityContext = document.createElement("canvas").getContext("2d");
    watchFontLoads();
  }
  if (!availabilityContext) return undefined;
  if (faceAvailability.has(family)) return faceAvailability.get(family);
  // A face the product serves itself (styles/office-fonts*.css alias Office's
  // families to open faces behind local()) is available wherever the page
  // is, even before its file has arrived: the width probe below would read
  // the interim fallback and pick the substitution stack for a font that is
  // on its way. Ask for it now so the swap lands with the first paint it
  // can. A family whose stand-in the stack names itself counts the same.
  if (officeFontDeclared(family)) {
    void document.fonts.load(`12px "${family}"`).catch(() => undefined);
    faceAvailability.set(family, true);
    return true;
  }
  if (officeFontSubstitutes(family).length) {
    faceAvailability.set(family, true);
    return true;
  }
  // FontFaceSet.check returns true for unknown local families. Independent
  // fallback probes distinguish a real selected face from silent substitution.
  const widths = ["serif", "sans-serif", "monospace"].map(fallback => {
    availabilityContext!.font = `100px "${family}", ${fallback}`;
    return availabilityContext!.measureText("0123456789MWil").width;
  });
  const available = Math.max(...widths) - Math.min(...widths) < 0.01;
  if (faceAvailability.size > 512) faceAvailability.clear();
  faceAvailability.set(family, available);
  return available;
}

const LATIN_SERIF_FACE = /^(times( new roman)?|georgia|cambria|garamond|book antiqua|palatino( linotype)?|century( schoolbook)?|constantia|bookman old style|baskerville( old face)?|perpetua|goudy old style|bodoni mt|didot|hoefler text|charter|iowan old style)$/i;
const LATIN_SERIF_STAND_IN = /^(tinos|caladea|gelasio|literata|bodoni moda|rokkitt|cormorant garamond|eb garamond|tex gyre pagella|cardo|alegreya|liberation serif|crimson (text|pro)|libre baskerville)$/i;

/** A Latin face drawn with serifs, by name or by the open face standing in for it. */
function latinSerifFace(face: string, standIns: readonly string[]): boolean {
  return LATIN_SERIF_FACE.test(face) || (standIns.length > 0 && LATIN_SERIF_STAND_IN.test(standIns[0]));
}

/** The same declared face/fallback stack is used by canvas, DOM and AutoFit.
 * Keep Office Korean coverage after the established Han/Kana fallbacks, so
 * missing Hangul glyphs do not fall through to a different system font.
 */
// Every cell of every render asks for its stack, and building one walks
// substitutes, script fallbacks and the served faces (in Page Layout, about
// 0.6 s of each scroll's re-render). A stack depends only on the family, the
// traits, the text engine switch and what the browser has loaded, which
// advances `typographyGeneration`, so it is kept per generation.
const fontStackCache = new Map<string, string>();
let fontStackCacheGeneration = -1;

export function spreadsheetFontStack(family: string | undefined, traits: Pick<NativeGridFont, "bold" | "italic"> = {}): string {
  const engine = spreadsheetTextEngineOn();
  if (fontStackCacheGeneration !== typographyGeneration) {
    fontStackCache.clear();
    fontStackCacheGeneration = typographyGeneration;
  }
  const mac = typeof navigator !== "undefined" && /Mac/i.test(navigator.platform);
  const key = `${engine ? 1 : 0}${mac ? 1 : 0}${traits.bold ? 1 : 0}${traits.italic ? 1 : 0}|${family ?? ""}`;
  const cached = fontStackCache.get(key);
  if (cached !== undefined) return cached;
  const stack = buildSpreadsheetFontStack(family, traits);
  if (typeof document !== "undefined") {
    if (fontStackCache.size >= 4096) fontStackCache.clear();
    fontStackCache.set(key, stack);
  }
  return stack;
}

function buildSpreadsheetFontStack(family: string | undefined, traits: Pick<NativeGridFont, "bold" | "italic">): string {
  // A cell stores one face, but the shared font picker hands a CSS stack
  // ("Arial, Helvetica, sans-serif"): its first family is the face. Read
  // whole, the stack was one unknown name drawn in the serif fallback (user
  // audit S-006).
  const named = family?.split(",")[0]?.replace(/["'\\;{}()<>]/g, "").trim().slice(0, 200);
  // A face named by its full or PostScript name ("Lucida Sans Regular") is
  // that face of its family, as Excel for Mac opens it (X-028).
  const face = named && spreadsheetFontAvailable(named) === false ? officeFaceName(named)?.family ?? named : named;
  // The face stands first in the stack: its slice holding the space is the
  // line's first available font once loaded (loadOfficeFamilySpace).
  loadOfficeFamilySpace(face);
  if (face && spreadsheetFontAvailable(face) === false && /Mac/i.test(navigator.platform)) {
    // Native Mac Excel uses a serif substitute for regular missing faces,
    // but a sans-serif substitute for bold/italic variants. Keep the requested
    // family first so an available or explicitly bound face always wins.
    const latin = traits.bold || traits.italic ? '"Helvetica", "Times New Roman"' : '"Times New Roman"';
    return servedStack(`"${face}", ${latin}, "MS Mincho", "Malgun Gothic", "Mangal", "Tahoma", serif`);
  }
  ensureOfficeFontFaces(face);
  const standIns = face ? officeFontSubstitutes(face) : [];
  const substitutes = standIns.map((name) => `"${name}", `).join("");
  // Han text in a Latin face has no glyphs of its own and takes the system's
  // East Asian fallback, which keeps the face's class: Excel draws the Han in
  // a Times New Roman cell as Song (serif) and in Calibri or Arial as a sans
  // (measured, Excel for Mac). Only a serif Latin face gets the serif Han.
  const hanSerif = face && latinSerifFace(face, standIns) ? (ensureOfficeFontFaces("SimSun"), '"SimSun", "Songti SC", ') : "";
  // A serif face's Cyrillic or Greek its stand-in lacks (Caladea, Cambria's,
  // has none) is drawn by a proportional serif with those scripts, not by
  // the Han face after it, whose Cyrillic is wide and spaced (PH057's
  // Cambria "Ключевые характеристики", X-246).
  const scriptSerif =
    hanSerif && !/^times new roman$/i.test(face ?? "")
      ? (ensureOfficeFontFaces("Times New Roman"), '"Times New Roman", ')
      : "";
  return servedStack(`${face ? `"${face}", ${substitutes}${scriptSerif}${hanSerif}` : ""}"Calibri","Microsoft YaHei", "Malgun Gothic", "Mangal", "Tahoma", "PingFang SC", sans-serif`);
}

/**
 * With the text engine on, a cell's stack ends in faces the page serves
 * rather than in whatever the visitor's system has: every Office family in
 * it is declared, the system's PingFang gives way to the served script
 * fallback family, and only then the generic family. Text the engine
 * measures is drawn in exactly the faces it measured (lib/text-engine).
 */
function servedStack(stack: string): string {
  if (!spreadsheetTextEngineOn()) return stack;
  const parts = stack.split(",").map((part) => part.trim()).filter((part) => part && !/^["']?PingFang SC["']?$/i.test(part));
  for (const part of parts) {
    const name = part.replace(/^["']|["']$/g, "");
    if (!GENERIC_FAMILIES.has(name.toLowerCase())) ensureOfficeFontFaces(name);
  }
  return withScriptFallbacks(parts.join(", "));
}

/** Resolve after the style cascade, including conditional/table and rich-run
 * traits. Measurement and paint must select the same actual fallback face. */
export function resolveSpreadsheetFontStyle(style: CSSProperties): CSSProperties {
  return withOfficeFauxBold(resolveSpreadsheetFontFamily(style));
}

/** Bold in an Office family with no bold face, decided on the cascaded
 * weight (a table style's or a rule's bold too, and a rule that turns a bold
 * cell regular): the faux-bold stroke where the browser synthesizes none
 * (lib/office-fonts/faux-bold; S11_totalsRowFunction's table-style totals
 * row, X-099). */
function withOfficeFauxBold(style: CSSProperties): CSSProperties {
  const bold = style.fontWeight === "bold" || Number(style.fontWeight) >= 600;
  const family = typeof style.fontFamily === "string" ? style.fontFamily : undefined;
  if (bold && officeFamilyHasNoBold(family)) {
    const faux = spreadsheetFauxBoldCss();
    if (style.WebkitTextStrokeWidth === faux.WebkitTextStrokeWidth && style.fontSynthesis === faux.fontSynthesis) return style;
    ensureOfficeFauxBoldProbe();
    return { ...style, ...faux };
  }
  if (!isSpreadsheetFauxBoldStroke(style.WebkitTextStrokeWidth)) return style;
  const rest = { ...style };
  delete rest.WebkitTextStrokeWidth;
  delete rest.fontSynthesis;
  return rest;
}

function resolveSpreadsheetFontFamily(style: CSSProperties): CSSProperties {
  if (typeof style.fontFamily !== "string") return style;
  const family = officePrimaryFontFamily(style.fontFamily);
  if (!family || spreadsheetFontAvailable(family) !== false || !/Mac/i.test(navigator.platform)) return style;
  const fontFamily = spreadsheetFontStack(family, {
    bold: style.fontWeight === "bold" || Number(style.fontWeight) >= 600,
    italic: style.fontStyle === "italic" || style.fontStyle?.startsWith("oblique") === true,
  });
  return fontFamily === style.fontFamily ? style : { ...style, fontFamily };
}

/** Native Mac substitution follows the leading script of a source font run.
 * Paired Han/Latin controls with two absent families retain italic in the
 * source, but paint Han-led italic runs with regular fallback faces. Resolve
 * before wrapping and carry the original leading text onto every fragment.
 * Available/bound fonts and other native device profiles keep their traits.
 */
export function resolveSpreadsheetTextStyle(style: CSSProperties, text: string, dpi?: number): CSSProperties {
  let resolved = resolveSpreadsheetFontStyle(style);
  // Mac screen-copy vectors quantize the drawing size before positioning
  // glyphs, for both available faces and substitutes. Keep this a disposable
  // CSS projection: imported point sizes and their native pairs stay exact.
  if (dpi === 72 && typeof navigator !== "undefined" && /Mac/i.test(navigator.platform) &&
      typeof resolved.fontSize === "number" && Number.isFinite(resolved.fontSize)) {
    const fontSize = Math.round(resolved.fontSize * .75) / .75;
    if (fontSize > 0 && fontSize !== resolved.fontSize) resolved = { ...resolved, fontSize };
  }
  if (dpi !== 72 || typeof navigator === "undefined" || !/Mac/i.test(navigator.platform) ||
      style.fontStyle !== "italic" || style.fontWeight === "bold" || Number(style.fontWeight) >= 600 ||
      !/^\p{Script=Han}/u.test(text) || typeof style.fontFamily !== "string") return resolved;
  const family = officePrimaryFontFamily(style.fontFamily);
  if (!family || spreadsheetFontAvailable(family) !== false) return resolved;
  return { ...resolved, fontFamily: spreadsheetFontStack(family), fontStyle: "normal" };
}

let textLayoutQuantum: number | undefined;
/** Detect the layout engine's fractional CSS grid instead of assuming the
 * same subpixel denominator across browsers. Independent of display DPR. */
export function nativeTextLayoutQuantum(): number {
  if (textLayoutQuantum !== undefined) return textLayoutQuantum;
  if (typeof document === "undefined" || !document.body) return 0;
  const probe = document.createElement("span");
  probe.style.cssText = "position:fixed;visibility:hidden;pointer-events:none;display:block;box-sizing:content-box;border:0;padding:0;margin:0;height:0;contain:strict;";
  document.body.appendChild(probe);
  let smallest = 0;
  let cssScale = 1;
  for (let power = 0; power <= 16; power += 1) {
    probe.style.width = `${1 / 2 ** power}px`;
    const width = probe.getBoundingClientRect().width;
    if (power === 0 && width > 0) cssScale = width;
    if (width > 0) smallest = width / cssScale;
    else break;
  }
  probe.remove();
  textLayoutQuantum = smallest;
  return smallest;
}

let textContext: CanvasRenderingContext2D | null | undefined;
const advanceCache = new Map<string, number>();
const idealAdvanceCache = new Map<string, number>();
/** Whole answers by text, font, grid, scale and offset: a scrolled band asks
 * again for what the rows it reuses were given (the same values down a
 * column, a row rendered again for its interaction). Same lifetime as the
 * per-glyph caches. */
const advancesCache = new Map<string, ReadonlyArray<{ text: string; spacing: number; advance?: number }>>();
let graphemes: Intl.Segmenter | undefined;
/** Excel device-pixel advances, independently measured for each grapheme.
 * Joining scripts retain browser shaping as a run; no code-unit splitting.
 */
/** A font load drops the advances it can change: those of a font and text
 * it reaches (font-load-reach), every one when it names no faces. True when
 * any went. */
export function dropAdvancesTouched(event: Event): boolean {
  const reach = fontLoadReach(event);
  if (!reach) {
    advanceCache.clear(); idealAdvanceCache.clear(); advancesCache.clear();
    return true;
  }
  let dropped = false;
  for (const cache of [advanceCache, idealAdvanceCache]) {
    for (const key of [...cache.keys()]) {
      const split = key.indexOf(":");
      if (split < 0 || reach.touchesFont(key.slice(0, split), key.slice(split + 1))) {
        cache.delete(key);
        dropped = true;
      }
    }
  }
  for (const key of [...advancesCache.keys()]) {
    const [text, font] = key.split("\u0000");
    if (font === undefined || reach.touchesFont(font, text)) {
      advancesCache.delete(key);
      dropped = true;
    }
  }
  return dropped;
}

export function nativeTextAdvances(text: string, fontCss: string, dpi: number, scale = 1, _purpose: "paint" | "measure" = "paint", advanceOffset = 0): Array<{ text: string; spacing: number; advance?: number }> {
  // Keep the existing call contract; both screen operations share advances.
  void _purpose;
  if (typeof document === "undefined" || /[\u0590-\u0fff]/u.test(text)) return [{ text, spacing: 0 }];
  if (textContext === undefined) {
    textContext = document.createElement("canvas").getContext("2d");
    document.fonts?.addEventListener("loadingdone", (event) => {
      if (!dropAdvancesTouched(event)) return;
      // Texts measured with what went are measured again (SpreadsheetNativeText).
      typographyGeneration += 1;
      for (const listener of [...typographyListeners]) listener();
    });
  }
  if (!textContext) return [{ text, spacing: 0 }];
  const answerKey = `${text}\u0000${fontCss}\u0000${dpi}\u0000${scale}\u0000${advanceOffset}`;
  const known = advancesCache.get(answerKey);
  // A copy: callers own what they are given.
  if (known) return known.map((run) => ({ ...run }));
  const answer = measureNativeTextAdvances(textContext, text, fontCss, dpi, scale, advanceOffset);
  if (advancesCache.size > 20000) advancesCache.clear();
  advancesCache.set(answerKey, answer.map((run) => ({ ...run })));
  return answer;
}

/** A measuring context for each font a sheet draws in, kept with that font
 * set. A measurement just after a canvas font changes resolves the face
 * again (a family list with CJK fallbacks: about 30 µs), and cells in
 * different fonts or sizes interleave on every row, so one shared context
 * switching font paid that for each new text (a cold Japanese workbook:
 * about a second of its first renders). Fonts beyond the bound go oldest
 * first. */
const measureContexts = new Map<string, CanvasRenderingContext2D>();
const MAX_MEASURE_CONTEXTS = 48;
export function fontMeasureContext(font: string, kerning: CanvasFontKerning = "auto"): CanvasRenderingContext2D | null {
  const key = `${kerning}\u0000${font}`;
  const known = measureContexts.get(key);
  if (known) return known;
  if (typeof document === "undefined") return null;
  const context = document.createElement("canvas").getContext("2d");
  if (!context) return null;
  context.font = font;
  context.fontKerning = kerning;
  if (measureContexts.size >= MAX_MEASURE_CONTEXTS) {
    const oldest = measureContexts.keys().next().value;
    if (oldest !== undefined) measureContexts.delete(oldest);
  }
  measureContexts.set(key, context);
  return context;
}

function measureNativeTextAdvances(context: CanvasRenderingContext2D, text: string, fontCss: string, dpi: number, scale: number, advanceOffset: number): Array<{ text: string; spacing: number; advance?: number }> {
  const textContext = context;
  const fontSize = fontCss.match(/([0-9]+(?:\.[0-9]+)?)(px|pt)/);
  if (!fontSize) return [{text,spacing:0}];
  // Screen-copy vectors position visible glyphs on the native device grid
  // even when the declared font is absent. Measure the resolved fallback
  // stack below; print-PDF text advances are not a screen positioning oracle.
  const pixels = Number(fontSize[1]) * (fontSize[2] === "pt" ? PX_PER_PT : 1);
  const safeScale = Number.isFinite(scale) && scale > 0 ? scale : 1;
  const paintedFont = fontCss.replace(fontSize[0], `${pixels * safeScale}px`);
  const quantum = nativeTextLayoutQuantum();
  let totalDesired = advanceOffset;
  let previousSnapped = quantum ? Math.round(advanceOffset / quantum) * quantum : advanceOffset;
  const idealFont = fontCss.replace(fontSize[0], "1000px");
  graphemes ??= new Intl.Segmenter(undefined, { granularity: "grapheme" });
  const segments = Array.from(graphemes.segment(text), ({ segment }) => segment);
  // The design widths from the served font files (lib/text-engine) where the
  // engine knows every grapheme's face: the same in every browser. Else the
  // browser's own, as before.
  const engineIdeals = spreadsheetEngineIdealAdvances(segments, fontCss);
  // The design and painted widths come from contexts kept in their fonts
  // (fontMeasureContext). The painted widths are measured, not scaled from
  // the design ones: hinted CJK faces are not linear (0.33 px at 12 px).
  const ideals = segments.map((segment, index) => {
    if (segment === "\n" || segment === "\r") return 0;
    // The glyph's design width, read once per face: every cell on screen
    // asks for the same few dozen characters on every render.
    const idealKey = `${idealFont}:${segment}`;
    let ideal = engineIdeals?.[index] ?? idealAdvanceCache.get(idealKey);
    if (ideal === undefined) {
      ideal = (fontMeasureContext(idealFont, "none") ?? textContext).measureText(segment).width;
      if (idealAdvanceCache.size > 12000) idealAdvanceCache.clear();
      idealAdvanceCache.set(idealKey, ideal);
    }
    return ideal;
  });
  const painted = segments.map((segment) => {
    if (segment === "\n" || segment === "\r") return 0;
    const key = `${paintedFont}:${segment}`;
    let width = advanceCache.get(key);
    if (width === undefined) {
      width = (fontMeasureContext(paintedFont, "none") ?? textContext).measureText(segment).width;
      if (advanceCache.size > 12000) advanceCache.clear();
      advanceCache.set(key, width);
    }
    return width;
  });
  return segments.map((segment, index) => {
    if (segment === "\n" || segment === "\r") {
      totalDesired = 0;
      previousSnapped = 0;
      return { text: segment, spacing: 0 };
    }
    const idealWidth = ideals[index] * pixels / 1000;
    const desired = Math.round(idealWidth * dpi / 96) * 96 / dpi * safeScale;
    totalDesired += desired;
    const snapped = quantum ? Math.round(totalDesired / quantum) * quantum : totalDesired;
    const targetAdvance = snapped - previousSnapped;
    previousSnapped = snapped;
    const spacing = (targetAdvance - painted[index]) / safeScale;
    return { text: segment, spacing, advance: desired };
  });
}
