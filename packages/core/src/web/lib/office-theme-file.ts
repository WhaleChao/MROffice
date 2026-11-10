/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Office theme files, as Word's Design > Themes > Browse for Themes reads
 * them (Excel's Page Layout and PowerPoint's Design tabs read the same): an
 * Office theme (`.thmx`) or any themed Office document (`.docx`, `.dotx`,
 * `.xlsx`, `.pptx` and their template / macro kinds), whose theme part is
 * the theme. MRAI's own earlier theme file (JSON) still reads.
 *
 * The theme comes back as the shared catalog sees it: per scheme (Colors,
 * Fonts, Effects) Office's id where one is the same, else the file's own
 * scheme as it is (the `native*` shapes the spreadsheet already keeps), and
 * the Office theme all three are, if any. The theme part is found the way
 * Office finds it, by relationships: the package's main part (a theme
 * file's theme manager, a document, a workbook) relates the theme; a
 * presentation's comes from its first slide master.
 */

import { loadZipSafely, type SafeZipLimits } from "./archive/safeZip";
import {
  OFFICE_COLOR_SCHEMES,
  OFFICE_EFFECT_SCHEMES,
  OFFICE_FONT_SCHEMES,
  OFFICE_THEMES,
  OFFICE_THEME_COLOR_SLOTS,
  OFFICE_THEME_FONT_SCRIPTS,
  officeSchemeSignature,
  officeSignatureHash,
} from "./office-theme-catalog";
import type { SpreadsheetNativeThemeColors } from "./spreadsheet/color-palette";
import {
  normalizeSpreadsheetNativeThemeColors,
  normalizeSpreadsheetNativeThemeEffects,
  normalizeSpreadsheetNativeThemeFonts,
  parseSpreadsheetWorkbookThemeFile,
  type SpreadsheetNativeThemeEffects,
  type SpreadsheetNativeThemeFonts,
} from "./spreadsheet/workbook-theme";

export interface OfficeThemeFileTheme {
  /** The theme's own name (`a:theme/@name`), or null for MRAI's JSON file. */
  readonly name: string | null;
  readonly colors: { readonly id: string } | { readonly native: SpreadsheetNativeThemeColors };
  readonly fonts: { readonly id: string } | { readonly native: SpreadsheetNativeThemeFonts };
  readonly effects: { readonly id: string } | { readonly native: SpreadsheetNativeThemeEffects };
  /** The Office theme whose three schemes these are, if any. */
  readonly theme: string | null;
}

/** The files Browse for Themes offers ("Office Themes and Themed Documents"). */
export const OFFICE_THEME_FILE_ACCEPT = [
  ".thmx",
  ".docx",
  ".docm",
  ".dotx",
  ".dotm",
  ".xlsx",
  ".xlsm",
  ".xltx",
  ".xltm",
  ".pptx",
  ".pptm",
  ".potx",
  ".potm",
  ".ppsx",
  ".json",
].join(",");

const LIMITS: Readonly<SafeZipLimits> = Object.freeze({
  maxCompressedBytes: 100 * 1024 * 1024,
  maxEntries: 5_000,
  maxEntryUncompressedBytes: 64 * 1024 * 1024,
  maxTotalUncompressedBytes: 256 * 1024 * 1024,
  maxCompressionRatio: 1_000,
  minRatioCheckUncompressedBytes: 1024 * 1024,
});

/** A theme part is a few kilobytes; anything past this is not one. */
const MAX_THEME_PART_CHARS = 2 * 1024 * 1024;

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const RELS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const THEME_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const SLIDE_MASTER_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";

function parseXml(text: string): Document | null {
  if (typeof DOMParser === "undefined" || text.length > MAX_THEME_PART_CHARS) return null;
  const parsed = new DOMParser().parseFromString(text, "application/xml");
  return parsed.getElementsByTagName("parsererror").length ? null : parsed;
}

function directChildren(parent: Element, local: string, namespace = A): Element[] {
  return Array.from(parent.children).filter(
    (child) => child.localName === local && child.namespaceURI === namespace,
  );
}

function only(parent: Element | null | undefined, local: string): Element | null {
  if (!parent) return null;
  const found = directChildren(parent, local);
  return found.length === 1 ? found[0] : null;
}

/* ------------------------------------------------------------ the package */

function partDirectory(part: string): string {
  const slash = part.lastIndexOf("/");
  return slash < 0 ? "" : part.slice(0, slash + 1);
}

function relsPartOf(part: string): string {
  const directory = partDirectory(part);
  return `${directory}_rels/${part.slice(directory.length)}.rels`;
}

/** A relationship target inside the package, or null (external, escaping). */
function resolveTarget(source: string, target: string): string | null {
  if (!target || /^[a-z][a-z0-9+.-]*:/i.test(target)) return null;
  const segments = (target.startsWith("/") ? target.slice(1) : partDirectory(source) + target).split("/");
  const out: string[] = [];
  for (const segment of segments) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") {
      if (!out.length) return null;
      out.pop();
    } else out.push(segment);
  }
  return out.join("/");
}

type ZipLike = { file(name: string): { async(type: "string"): Promise<string> } | null };

async function read(zip: ZipLike, part: string): Promise<string | null> {
  const entry = zip.file(part);
  return entry ? entry.async("string") : null;
}

/** The parts a part relates by type (internal targets only). */
async function related(zip: ZipLike, part: string, type: string): Promise<string[]> {
  const text = await read(zip, part === "" ? "_rels/.rels" : relsPartOf(part));
  const rels = text ? parseXml(text) : null;
  if (!rels) return [];
  return Array.from(rels.getElementsByTagNameNS(RELS, "Relationship"))
    .filter((rel) => rel.getAttribute("Type") === type && !/^external$/i.test(rel.getAttribute("TargetMode") ?? ""))
    .map((rel) => resolveTarget(part, rel.getAttribute("Target") ?? ""))
    .filter((target): target is string => target !== null);
}

async function themePartOf(zip: ZipLike): Promise<string | null> {
  const [main] = await related(zip, "", OFFICE_DOCUMENT_REL);
  if (!main) return null;
  const [theme] = await related(zip, main, THEME_REL);
  if (theme) return read(zip, theme);
  // A presentation's theme is its first slide master's.
  const [master] = await related(zip, main, SLIDE_MASTER_REL);
  const [masterTheme] = master ? await related(zip, master, THEME_REL) : [];
  return masterTheme ? read(zip, masterTheme) : null;
}

/* ---------------------------------------------------------- the schemes */

function colorValue(slot: Element | null): string | null {
  const children = slot ? Array.from(slot.children) : [];
  if (children.length !== 1) return null;
  const color = children[0];
  const value =
    color.localName === "srgbClr"
      ? color.getAttribute("val")
      : color.localName === "sysClr"
        ? color.getAttribute("lastClr")
        : null;
  return value && /^[0-9a-f]{6}$/i.test(value) ? `#${value.toLowerCase()}` : null;
}

function readColors(scheme: Element): SpreadsheetNativeThemeColors | null {
  const colors: Record<string, string> = {};
  for (const slot of OFFICE_THEME_COLOR_SLOTS) {
    const value = colorValue(only(scheme, slot));
    if (!value) return null;
    colors[slot] = value;
  }
  return normalizeSpreadsheetNativeThemeColors({ name: scheme.getAttribute("name") ?? "", colors });
}

function readFontCollection(container: Element | null) {
  if (!container) return null;
  const face = (local: string) => only(container, local)?.getAttribute("typeface") ?? null;
  const latin = face("latin");
  const ea = face("ea");
  const cs = face("cs");
  if (latin === null || ea === null || cs === null) return null;
  const known = new Set<string>(OFFICE_THEME_FONT_SCRIPTS);
  const scripts: Record<string, string> = {};
  for (const font of directChildren(container, "font")) {
    const script = font.getAttribute("script") ?? "";
    if (known.has(script) && !(script in scripts)) scripts[script] = font.getAttribute("typeface") ?? "";
  }
  return { latin, ea, cs, scripts };
}

function readFonts(scheme: Element): SpreadsheetNativeThemeFonts | null {
  return normalizeSpreadsheetNativeThemeFonts({
    name: scheme.getAttribute("name") ?? "",
    major: readFontCollection(only(scheme, "majorFont")),
    minor: readFontCollection(only(scheme, "minorFont")),
  });
}

function sameColors(left: Readonly<Record<string, string>>, right: Readonly<Record<string, string>>): boolean {
  return OFFICE_THEME_COLOR_SLOTS.every((slot) => left[slot]?.toLowerCase() === right[slot]?.toLowerCase());
}

/** Office's Colors scheme with these twelve colours: the one named the same first, the gallery's next. */
function matchColors(native: SpreadsheetNativeThemeColors): string | null {
  const matches = OFFICE_COLOR_SCHEMES.filter((scheme) => sameColors(scheme.colors, native.colors));
  return (
    matches.find((scheme) => scheme.schemeName === native.name)?.id ??
    matches.find((scheme) => scheme.gallery)?.id ??
    matches[0]?.id ??
    null
  );
}

/** Word's Office theme fonts (Aptos, or 等线 in a Simplified Chinese Normal template). */
const WORD_OFFICE_FONT_KEYS = new Set(["aptos display\u0000\u0000aptos\u0000", "等线 light\u0000\u0000等线\u0000"]);

function fontKey(major: { latin: string; ea: string }, minor: { latin: string; ea: string }): string {
  return [major.latin.trim().toLowerCase(), major.ea.trim(), minor.latin.trim().toLowerCase(), minor.ea.trim()].join("\u0000");
}

function matchFonts(native: SpreadsheetNativeThemeFonts): string | null {
  const key = fontKey(native.major, native.minor);
  if (WORD_OFFICE_FONT_KEYS.has(key)) return "office";
  const matches = OFFICE_FONT_SCHEMES.filter((scheme) => fontKey(scheme.major, scheme.minor) === key);
  return (
    matches.find((scheme) => scheme.schemeName === native.name)?.id ??
    matches.find((scheme) => scheme.gallery)?.id ??
    matches[0]?.id ??
    null
  );
}

function matchEffects(scheme: Element): string | null {
  const signature = officeSignatureHash(officeSchemeSignature(scheme));
  return OFFICE_EFFECT_SCHEMES.find((effect) => effect.signature === signature)?.id ?? null;
}

/** An effect scheme is kept verbatim (the size the spreadsheet allows). */
function readEffects(scheme: Element): SpreadsheetNativeThemeEffects | null {
  const name = scheme.getAttribute("name") ?? "";
  const xml = new XMLSerializer().serializeToString(scheme);
  return (
    normalizeSpreadsheetNativeThemeEffects({ name, xml }) ??
    normalizeSpreadsheetNativeThemeEffects({ name })
  );
}

function themeOf(
  colors: string | null | undefined,
  fonts: string | null | undefined,
  effects: string | null | undefined,
): string | null {
  if (!colors || !fonts || !effects) return null;
  return OFFICE_THEMES.find((theme) => theme.colors === colors && theme.fonts === fonts && theme.effects === effects)?.id ?? null;
}

/** A theme part (`a:theme`) read against the catalog, or null if it is not one. */
export function officeThemeFromThemeXml(text: string): OfficeThemeFileTheme | null {
  const parsed = parseXml(text);
  const root = parsed?.documentElement;
  if (!root || root.localName !== "theme" || root.namespaceURI !== A) return null;
  const elements = only(root, "themeElements");
  const colorScheme = only(elements, "clrScheme");
  const fontScheme = only(elements, "fontScheme");
  const formatScheme = only(elements, "fmtScheme");
  if (!colorScheme || !fontScheme || !formatScheme) return null;
  const nativeColors = readColors(colorScheme);
  const nativeFonts = readFonts(fontScheme);
  if (!nativeColors || !nativeFonts) return null;
  const colorsId = matchColors(nativeColors);
  const fontsId = matchFonts(nativeFonts);
  const effectsId = matchEffects(formatScheme);
  const nativeEffects = effectsId ? null : readEffects(formatScheme);
  if (!effectsId && !nativeEffects) return null;
  const name = (root.getAttribute("name") ?? "").trim();
  return {
    name: name || null,
    colors: colorsId ? { id: colorsId } : { native: nativeColors },
    fonts: fontsId ? { id: fontsId } : { native: nativeFonts },
    effects: effectsId ? { id: effectsId } : { native: nativeEffects! },
    theme: themeOf(colorsId, fontsId, effectsId),
  };
}

/** A small text file's text (the browser's Blob.text, else FileReader). */
async function blobText(blob: Blob): Promise<string> {
  if (typeof blob.text === "function") return blob.text();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error || new Error("Failed to read the theme file."));
    reader.onload = () => resolve(String(reader.result ?? ""));
    reader.readAsText(blob);
  });
}

/** MRAI's earlier JSON theme file is a few hundred bytes. */
const MAX_JSON_THEME_BYTES = 256 * 1024;

/**
 * The theme a chosen file carries, or null when it carries none this reads
 * (not an Office package, no theme part, an unreadable scheme).
 */
export async function readOfficeThemeFile(file: Blob & { name?: string }): Promise<OfficeThemeFileTheme | null> {
  if (/\.json$/i.test(file.name ?? "")) {
    // MRAI's earlier theme file: a selection of Office schemes.
    if (file.size > MAX_JSON_THEME_BYTES) return null;
    const selection = parseSpreadsheetWorkbookThemeFile(await blobText(file));
    if (!selection) return null;
    const colors = selection.nativeColors ? { native: selection.nativeColors } : { id: selection.colorId };
    const fonts = selection.nativeFonts ? { native: selection.nativeFonts } : { id: selection.fontId };
    const effects = selection.nativeEffects ? { native: selection.nativeEffects } : { id: selection.effectId };
    return {
      name: null,
      colors,
      fonts,
      effects,
      theme: themeOf(
        "id" in colors ? colors.id : null,
        "id" in fonts ? fonts.id : null,
        "id" in effects ? effects.id : null,
      ),
    };
  }
  let zip: ZipLike;
  try {
    zip = await loadZipSafely(file, LIMITS);
  } catch {
    return null;
  }
  const theme = await themePartOf(zip).catch(() => null);
  return theme ? officeThemeFromThemeXml(theme) : null;
}
