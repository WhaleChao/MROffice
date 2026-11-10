/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The faces the page declares, read from the same place the browser reads
 * them: every `@font-face` rule of the document's stylesheets
 * (styles/office-fonts.css and the rest) and every face added through the
 * FontFace API whose builder noted its source (lib/font-face-declarations:
 * the Office registry faces, the script fallback family, workbook faces).
 * Selection follows CSS Fonts 4 §5.2 (style, then weight, then the
 * composite face's unicode-range in reverse declaration order), so the face
 * the engine measures a character in is the face the browser draws it in.
 *
 * A face whose `src` tries a `local()` font first may draw in a font the
 * visitor has installed rather than the served file; such faces carry their
 * local names, and the measurer checks them once against the browser
 * (verify.ts) before trusting the served file's widths for them.
 */
import { fontFaceDescriptors, fontFaceSource, notedFontFaceCount, notedFontFaces, type NotedFontFace } from "../font-face-declarations";

export interface FaceDeclaration {
  /** Normalized (lower-case, unquoted) family name. */
  readonly family: string;
  readonly weightMin: number;
  readonly weightMax: number;
  /** "normal" | "italic" | "oblique". */
  readonly style: "normal" | "italic" | "oblique";
  /** Inclusive code point ranges; null = everything. */
  readonly ranges: readonly (readonly [number, number])[] | null;
  /** Absolute URL of the first url() source, or null when there is none. */
  readonly url: string | null;
  /** local() names tried ahead of the URL. */
  readonly localNames: readonly string[];
  /** size-adjust as a factor (1 = none). */
  readonly sizeAdjust: number;
  /** ascent/descent/line-gap overrides as fractions of the em, when declared. */
  readonly ascentOverride?: number;
  readonly descentOverride?: number;
  readonly lineGapOverride?: number;
  /** The wght the face states for itself (`font-variation-settings`): an
   * alias drawn from a variable stand-in at the stand-in's own weight. */
  readonly wght?: number;
  /** Declaration order across the whole index. */
  readonly order: number;
  /** Added through the FontFace API (not an @font-face rule). */
  readonly api?: boolean;
}

export const normalizeFamily = (family: string): string =>
  family.trim().replace(/^["']|["']$/g, "").trim().toLowerCase();

/** A CSS font-family list, split outside quotes, normalized. */
export function splitFamilyList(stack: string): string[] {
  const out: string[] = [];
  let current = "";
  let quote: string | null = null;
  for (const char of stack) {
    if (quote) {
      if (char === quote) quote = null;
      else current += char;
    } else if (char === '"' || char === "'") quote = char;
    else if (char === ",") {
      if (current.trim()) out.push(normalizeFamily(current));
      current = "";
    } else current += char;
  }
  if (current.trim()) out.push(normalizeFamily(current));
  return out;
}

export const GENERIC_FAMILIES: ReadonlySet<string> = new Set([
  "serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui", "ui-sans-serif", "ui-serif",
  "ui-monospace", "ui-rounded", "math", "emoji", "fangsong", "-apple-system", "blinkmacsystemfont",
]);

// One parse per list: a family's slices repeat the same unicode-range lists
// across its weights and styles, and the office faces' stylesheet and the
// faces declared by script were thousands of parses of a few hundred lists

const parsedUnicodeRanges = new Map<string, [number, number][] | null>();
const PARSED_UNICODE_RANGES_KEPT = 4096;

export function parseUnicodeRange(value: string | undefined | null): [number, number][] | null {
  if (!value) return null;
  const known = parsedUnicodeRanges.get(value);
  if (known !== undefined) return known;
  const parsed = parseUnicodeRangeList(value);
  if (parsedUnicodeRanges.size >= PARSED_UNICODE_RANGES_KEPT) parsedUnicodeRanges.clear();
  parsedUnicodeRanges.set(value, parsed);
  return parsed;
}

function parseUnicodeRangeList(value: string): [number, number][] | null {
  if (!value.trim()) return null;
  const out: [number, number][] = [];
  for (const token of value.split(",")) {
    const match = /^\s*u\+([0-9a-f?]{1,6})(?:-([0-9a-f]{1,6}))?\s*$/i.exec(token);
    if (!match) continue;
    if (match[1].includes("?")) {
      out.push([parseInt(match[1].replace(/\?/g, "0"), 16), parseInt(match[1].replace(/\?/g, "f"), 16)]);
    } else {
      const start = parseInt(match[1], 16);
      out.push([start, match[2] ? parseInt(match[2], 16) : start]);
    }
  }
  return out.length ? out : null;
}

export function parseWeight(value: string | undefined | null): [number, number] {
  const words = (value ?? "").trim().toLowerCase().split(/\s+/).filter(Boolean);
  const one = (word: string | undefined) => word === "bold" ? 700 : word === "normal" || word === undefined ? 400 : Number(word);
  const min = one(words[0]);
  const max = words.length > 1 ? one(words[1]) : min;
  return [Number.isFinite(min) ? min : 400, Number.isFinite(max) ? max : 400];
}

function parseStyle(value: string | undefined | null): FaceDeclaration["style"] {
  const word = (value ?? "").trim().toLowerCase();
  return word.startsWith("italic") ? "italic" : word.startsWith("oblique") ? "oblique" : "normal";
}

function parsePercent(value: string | undefined | null): number | undefined {
  const match = /^\s*(-?[0-9.]+)%\s*$/.exec(value ?? "");
  return match ? Number(match[1]) / 100 : undefined;
}

/** url() and local() sources of a CSS `src` list, in order. */
export function parseSources(src: string, base: string | undefined): { url: string | null; localNames: string[] } {
  const localNames: string[] = [];
  let url: string | null = null;
  const pattern = /(url|local)\(\s*(?:"((?:[^"\\]|\\.)*)"|'((?:[^'\\]|\\.)*)'|([^)]*?))\s*\)/gi;
  for (let match = pattern.exec(src); match; match = pattern.exec(src)) {
    const value = (match[2] ?? match[3] ?? match[4] ?? "").replace(/\\(.)/g, "$1").trim();
    if (match[1].toLowerCase() === "local") {
      if (url === null) localNames.push(value);
    } else if (url === null && value) {
      try {
        url = base ? new URL(value, base).href : value;
      } catch {
        url = value;
      }
    }
  }
  return { url, localNames };
}

interface Index {
  /** What the document's sheets and noted faces were when it was read. */
  readonly signature: string;
  /** Changes exactly when the declared faces do (a sheet without a face
   * coming or going leaves it). */
  readonly generation: number;
  readonly sources: readonly Omit<FaceDeclaration, "order">[];
  readonly ordered: readonly FaceDeclaration[];
  readonly families: ReadonlyMap<string, readonly FaceDeclaration[]>;
}
let index: Index | null = null;
let testFaces: FaceDeclaration[] | null = null;

/** Tests: a fixed declaration list instead of the document's. */
export function setFaceDeclarationsForTests(faces: readonly Omit<FaceDeclaration, "order">[] | null): void {
  testFaces = faces ? faces.map((face, order) => ({ ...face, order })) : null;
  index = null;
}

// Not `document.fonts.size`: in WebKit reading the set's size brings the
// document's style up to date first (every version; and its faces' getters
// do too in Safari 16.0-16.3), and this is read on every measurement. The
// index is the stylesheets' faces and the faces noted with a source, so
// those are what can change it.
function signatureOf(doc: Document): string {
  return `${doc.styleSheets.length}:${notedFontFaceCount()}`;
}

function collectRules(rules: CSSRuleList, base: string | undefined, out: Omit<FaceDeclaration, "order">[]): void {
  for (const rule of Array.from(rules)) {
    if (typeof CSSFontFaceRule !== "undefined" && rule instanceof CSSFontFaceRule) {
      const style = rule.style;
      const family = style.getPropertyValue("font-family");
      const src = style.getPropertyValue("src");
      if (!family || !src) continue;
      const [weightMin, weightMax] = parseWeight(style.getPropertyValue("font-weight"));
      const { url, localNames } = parseSources(src, base);
      out.push({
        family: normalizeFamily(family),
        weightMin, weightMax,
        style: parseStyle(style.getPropertyValue("font-style")),
        ranges: parseUnicodeRange(style.getPropertyValue("unicode-range")),
        url, localNames,
        sizeAdjust: parsePercent(style.getPropertyValue("size-adjust")) ?? 1,
        ascentOverride: parsePercent(style.getPropertyValue("ascent-override")),
        descentOverride: parsePercent(style.getPropertyValue("descent-override")),
        lineGapOverride: parsePercent(style.getPropertyValue("line-gap-override")),
        ...statedWght(style.getPropertyValue("font-variation-settings")),
      });
    } else if ("cssRules" in rule) {
      try {
        collectRules((rule as CSSGroupingRule).cssRules, base, out);
      } catch { /* unreadable */ }
    }
  }
}

// A stylesheet's faces are read once: reading every @font-face rule is the
// expensive part (styles/office-fonts.css alone is thousands of them), and
// sheets come and go far more often than their rules change (each part of
// the editor loaded after the deck opens brings its own sheet). A sheet is
// read again when its base or its number of top-level rules changes, which
// is how a script adds rules to it (insertRule, appendRule).
const sheetFaces = new WeakMap<CSSStyleSheet, {
  readonly base: string | undefined;
  readonly rules: number;
  readonly faces: readonly Omit<FaceDeclaration, "order">[];
}>();

function facesOfSheet(sheet: CSSStyleSheet, base: string | undefined): readonly Omit<FaceDeclaration, "order">[] | null {
  let rules: CSSRuleList;
  try {
    rules = sheet.cssRules;
  } catch {
    return null;
  }
  const known = sheetFaces.get(sheet);
  if (known && known.base === base && known.rules === rules.length) return known.faces;
  const faces: Omit<FaceDeclaration, "order">[] = [];
  collectRules(rules, base, faces);
  sheetFaces.set(sheet, { base, rules: rules.length, faces });
  return faces;
}

function stylesheetFaces(doc: Document): Omit<FaceDeclaration, "order">[] {
  const faces: Omit<FaceDeclaration, "order">[] = [];
  for (const sheet of Array.from(doc.styleSheets)) {
    const own = facesOfSheet(sheet, sheet.href ?? doc.baseURI);
    if (own) for (const face of own) faces.push(face);
  }
  return faces;
}

/** The wght a `font-variation-settings` value states, as a declaration field. */
function statedWght(settings: string | undefined): { wght?: number } {
  const match = settings ? /["']wght["']\s+(-?[\d.]+)/.exec(settings) : null;
  const value = match ? Number(match[1]) : NaN;
  return Number.isFinite(value) ? { wght: value } : {};
}

function apiFace(face: FontFace, src: string, base: string | undefined, entry?: NotedFontFace): Omit<FaceDeclaration, "order"> {
  // What the builder declared, when it said: in Safari 16.0-16.3 each
  // FontFace getter restyles the document first (font-face-declarations).
  const declared = entry?.declared;
  const family = declared ? declared.family : face.family;
  const weight = declared ? declared.weight ?? "normal" : face.weight;
  const style = declared ? declared.style ?? "normal" : face.style;
  const unicodeRange = declared ? declared.unicodeRange ?? "U+0-10FFFF" : face.unicodeRange;
  const [weightMin, weightMax] = parseWeight(weight);
  const { url, localNames } = parseSources(src, base);
  // The builder's own descriptors first: WebKit does not expose the
  // overrides on the FontFace object, and the engine must read the same
  // line box in every browser.
  const raw = face as FontFace & { sizeAdjust?: string; ascentOverride?: string; descentOverride?: string; lineGapOverride?: string; variationSettings?: string };
  const noted = entry ? entry.built ?? {} : fontFaceDescriptors(face) ?? {};
  const read = (name: keyof typeof raw & string) => (entry?.declared ? undefined : (raw[name] as string | undefined));
  const extra = {
    sizeAdjust: noted.sizeAdjust ?? read("sizeAdjust"),
    ascentOverride: noted.ascentOverride ?? read("ascentOverride"),
    descentOverride: noted.descentOverride ?? read("descentOverride"),
    lineGapOverride: noted.lineGapOverride ?? read("lineGapOverride"),
    variationSettings: noted.variationSettings ?? read("variationSettings"),
  };
  return {
    api: true,
    family: normalizeFamily(family),
    weightMin, weightMax,
    style: parseStyle(style),
    ranges: parseUnicodeRange(unicodeRange),
    url, localNames,
    sizeAdjust: parsePercent(extra.sizeAdjust) ?? 1,
    ascentOverride: parsePercent(extra.ascentOverride),
    descentOverride: parsePercent(extra.descentOverride),
    lineGapOverride: parsePercent(extra.lineGapOverride),
    ...statedWght(extra.variationSettings),
  };
}

// A noted face's declaration is read once: what it was declared with does
// not change (lib/font-face-declarations), and a family of CJK slices is
// hundreds of unicode-range lists to parse at every face added.
const notedDeclarations = new WeakMap<NotedFontFace, { readonly base: string | undefined; readonly face: Omit<FaceDeclaration, "order"> }>();

function notedDeclaration(entry: NotedFontFace, base: string | undefined): Omit<FaceDeclaration, "order"> {
  const known = notedDeclarations.get(entry);
  if (known && known.base === base) return known.face;
  const face = apiFace(entry.face, entry.src, base, entry);
  notedDeclarations.set(entry, { base, face });
  return face;
}

/** Every face the document declares, in declaration order. */
function declarationsOf(doc: Document): Omit<FaceDeclaration, "order">[] {
  const faces = stylesheetFaces(doc);
  // FontFace API faces follow the stylesheet faces (CSS Fonts 4 §2:
  // CSS-connected faces first, then the rest in insertion order). Only the
  // faces noted with a source are indexed, so the noted list is read rather
  // than the document's set: walking the set restyles the document in
  // WebKit, and reading its faces did too in Safari 16.0-16.3.
  const noted = notedFontFaces(doc);
  if (noted) {
    for (const entry of noted) faces.push(notedDeclaration(entry, doc.baseURI));
  } else {
    doc.fonts?.forEach((face) => {
      const src = fontFaceSource(face);
      if (src) faces.push(apiFace(face, src, doc.baseURI));
    });
  }
  return faces;
}

/** The index of `declarations`, keeping each face of `previous` that is
 * the same declaration at the same place (an unchanged index is
 * `previous` itself, its generation kept). */
function indexOf(signature: string, declarations: readonly Omit<FaceDeclaration, "order">[], previous: Index | null): Index {
  const ordered: FaceDeclaration[] = new Array(declarations.length);
  let same = previous !== null && previous.sources.length === declarations.length;
  for (let order = 0; order < declarations.length; order += 1) {
    const source = declarations[order];
    if (previous && previous.sources[order] === source) ordered[order] = previous.ordered[order];
    else {
      same = false;
      ordered[order] = { ...source, order };
    }
  }
  if (same && previous) return { ...previous, signature };
  return {
    signature,
    generation: (previous?.generation ?? 0) + 1,
    sources: declarations,
    ordered,
    families: group(ordered),
  };
}

/** The index a full read of the document makes, sharing nothing with an
 * earlier one (tests: the incremental index must equal it). */
export function fullFaceIndexForTests(doc: Document): ReadonlyMap<string, readonly FaceDeclaration[]> {
  const faces: Omit<FaceDeclaration, "order">[] = [];
  for (const sheet of Array.from(doc.styleSheets)) {
    let rules: CSSRuleList;
    try {
      rules = sheet.cssRules;
    } catch {
      continue;
    }
    collectRules(rules, sheet.href ?? doc.baseURI, faces);
  }
  const noted = notedFontFaces(doc);
  if (noted) for (const entry of noted) faces.push(apiFace(entry.face, entry.src, doc.baseURI, entry));
  else doc.fonts?.forEach((face) => {
    const src = fontFaceSource(face);
    if (src) faces.push(apiFace(face, src, doc.baseURI));
  });
  return group(faces.map((face, order) => ({ ...face, order })));
}

function group(faces: readonly FaceDeclaration[]): Map<string, FaceDeclaration[]> {
  const families = new Map<string, FaceDeclaration[]>();
  for (const face of faces) {
    const list = families.get(face.family);
    if (list) list.push(face);
    else families.set(face.family, [face]);
  }
  return families;
}

/** Every declared face of a family, in declaration order; empty when the
 * page declares none (a system font, or nothing). */
export function declaredFaces(family: string): readonly FaceDeclaration[] {
  if (testFaces) {
    if (!index) index = { signature: "test", generation: 0, sources: testFaces, ordered: testFaces, families: group(testFaces) };
    return index.families.get(family) ?? [];
  }
  if (typeof document === "undefined") return [];
  return currentIndex(document).families.get(family) ?? [];
}

function currentIndex(doc: Document): Index {
  const signature = signatureOf(doc);
  if (!index || index.signature !== signature) index = indexOf(signature, declarationsOf(doc), index);
  return index;
}

/** The index's identity: changes when faces are declared or removed, and
 * not when a stylesheet without a face comes or goes. */
export function faceIndexSignature(): string {
  if (testFaces) return "test";
  return typeof document === "undefined" ? "" : `faces:${currentIndex(document).generation}`;
}

/**
 * CSS Fonts 4 §5.2 step 4: of a family's faces, the ones forming the
 * composite face a (weight, italic) request selects -- every face sharing
 * the chosen face's style and weight descriptors, in declaration order.
 */
export function selectComposite(
  faces: readonly FaceDeclaration[], weight: number, italic: boolean, apiOrder: "last-wins" | "first-wins" = "last-wins",
): readonly FaceDeclaration[] {
  if (!faces.length) return faces;
  const styleOrder: FaceDeclaration["style"][] = italic ? ["italic", "oblique", "normal"] : ["normal", "oblique", "italic"];
  let pool: readonly FaceDeclaration[] = [];
  for (const style of styleOrder) {
    pool = faces.filter((face) => face.style === style);
    if (pool.length) break;
  }
  const covers = (face: FaceDeclaration) => face.weightMin <= weight && weight <= face.weightMax;
  let chosen = pool.find(covers);
  if (!chosen) {
    const below = pool.filter((face) => face.weightMax < weight).sort((a, b) => b.weightMax - a.weightMax);
    const above = pool.filter((face) => face.weightMin > weight).sort((a, b) => a.weightMin - b.weightMin);
    if (weight >= 400 && weight <= 500) {
      chosen = above.find((face) => face.weightMin <= 500) ?? below[0] ?? above[0];
    } else if (weight < 400) {
      chosen = below[0] ?? above[0];
    } else {
      chosen = above[0] ?? below[0];
    }
  }
  if (!chosen) return [];
  const pick = chosen;
  const composite = pool.filter((face) => face.style === pick.style && face.weightMin === pick.weightMin && face.weightMax === pick.weightMax);
  // Priority runs from the end of the array. Rule faces: the last defined
  // first (CSS Fonts 4 §4.5). API faces: as this browser orders them
  // (face-order.ts), after the rule faces.
  if (apiOrder !== "first-wins" || !composite.some((face) => face.api)) return composite;
  return [...composite.filter((face) => !face.api), ...composite.filter((face) => face.api).reverse()];
}

export function faceCovers(face: FaceDeclaration, codePoint: number): boolean {
  if (!face.ranges) return true;
  for (const [start, end] of face.ranges) if (codePoint >= start && codePoint <= end) return true;
  return false;
}
