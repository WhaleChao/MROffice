/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// Which open face stands in for an Office font family, wherever the page runs.
//
// styles/office-fonts.css declares Office's default families as @font-face
// aliases (local() first, the served substitute behind it), so plain CSS
// resolves "Calibri" on a machine without Office. Every other Office family
// is registered through the FontFace API the first time code names it, from
// registry.generated.json: the same alias-with-local()-first faces, plus the
// substitute under its own name. This module is that table for code that
// builds font stacks or judges availability.
//
// The registry is served by the API one family at a time (`lib/office-data/
// store.ts`, "font-registry"), not bundled: to a signed-in page and, as its
// mapping (the faces and their descriptors), to a guest's.
import generated from "./substitutes.generated.json";
import cjkLatinFaces from "./cjk-latin-faces.json";
import { excelCutFamilyOf, excelCutRangeIsEastAsian, isExcelCutFamily } from "./excel-cut-families";
import { HANGUL_EM_PREFIX, hangulEmFamilyOf, stackPrimaryFamily } from "./hangul-em";
import powerPointTable from "./powerpoint-substitutes.json";
import { notedFontFaces, noteFamilyRegistration, noteFontFaceSource, type DeclaredFontFace } from "../font-face-declarations";
import { loadOfficeDataParts, officeDataObtainable, officeDataPartKnown } from "../office-data/store";
import { parseUnicodeRange } from "../text-engine/face-index";
import { officeFamilyHasNoBold } from "./faux-bold";

export type OfficeFontMetric = "exact" | "cloned" | "near" | "style" | "fallback" | "none";

/**
 * Whether the served face sets the Office family at its own advance widths:
 * a metric-compatible open face (`exact`: Carlito for Calibri) or a metric
 * clone (`cloned`: an open face redrawn to the Office face's advances and line
 * box, lib/office-fonts/metric-clones, scripts/build-metric-clones.py).
 */
export function officeFontMetricMatchesWidths(metric: string | undefined): boolean {
  return metric === "exact" || metric === "cloned";
}

export interface OfficeFontSubstitution {
  /** The Office family the name belongs to (aliases point at their family). */
  readonly office: string;
  /** Open families that stand in, in stack order; empty when nothing can. */
  readonly substitutes: readonly string[];
  /** How faithfully the substitute reproduces the original's widths. */
  readonly metric: OfficeFontMetric;
  /** The substitute weight that renders this name's regular style. */
  readonly weight?: number;
  /** A face served for files that name it, not one the font list offers. */
  readonly unlisted?: boolean;
  /** A legacy encoding the family's codes stand for (legacy-encodings/README.md):
   * set where MRAI deliberately draws what the author wrote, not the Latin
   * gibberish Word for Mac shows without the font. */
  readonly encoding?: string;
  /** The stand-in is drawn with Word's own line box, and sets a line as wide
   * as Word's font (near, East Asian, or scaled to Word's widths). */
  readonly wordLines?: true;
}

/** A face to register: alias weight, alias italic, substitute weight, substitute
 * italic, local names, and Word's own ascent, descent and line gap (percent of
 * the font size) for a stand-in that is not metric-compatible, when known
 * (scripts/measure-word-vertical-metrics.py). */
export type RegistryAliasFace = [number, 0 | 1, number, 0 | 1, string[], ([string, string, string] | null)?, string?];
type RegistrySubstituteFace = [number, 0 | 1, string, number | null];
interface Registry {
  /** By index: the whole table's array, or the served parts' sparse record. */
  readonly ranges: Readonly<Record<number, string>>;
  readonly substitutes: Readonly<Record<string, readonly RegistrySubstituteFace[]>>;
  readonly families: Readonly<
    Record<string, readonly { substitute: string; faces: readonly RegistryAliasFace[] }[]>
  >;
}

const PUBLIC_PREFIX = "/fonts/office/";
const table = generated as Record<string, OfficeFontSubstitution>;

const normalize = (family: string) =>
  family.trim().replace(/^["']|["']$/g, "").trim().toLowerCase();

export function officeFontSubstitution(
  family: string | null | undefined,
): OfficeFontSubstitution | undefined {
  if (!family) return undefined;
  const own = table[normalize(family)];
  if (own) return own;
  // A family named with its style ("Helvetica Neue Light", "Arial Bold"):
  // the family the style belongs to, at that style's weight.
  for (const candidate of styledFamilyCandidates(family)) {
    const base = table[normalize(candidate.base)];
    if (base?.substitutes.length) {
      return { ...base, ...(candidate.weight !== 400 ? { weight: candidate.weight } : {}) };
    }
  }
  return undefined;
}

/** A family name read as a family plus a style: `base` drawn at `weight`. */
export interface StyledFamily {
  readonly base: string;
  readonly weight: number;
  readonly italic: boolean;
}

const STYLE_WORDS: Readonly<Record<string, { readonly weight?: number; readonly italic?: true }>> = {
  thin: { weight: 100 }, hairline: { weight: 100 },
  extralight: { weight: 200 }, ultralight: { weight: 200 },
  light: { weight: 300 }, semilight: { weight: 350 },
  book: { weight: 400 }, regular: { weight: 400 }, normal: { weight: 400 }, roman: { weight: 400 },
  medium: { weight: 500 },
  semibold: { weight: 600 }, demibold: { weight: 600 },
  bold: { weight: 700 },
  extrabold: { weight: 800 }, ultrabold: { weight: 800 },
  heavy: { weight: 900 }, black: { weight: 900 },
  italic: { italic: true }, italics: { italic: true }, oblique: { italic: true },
};
/** Producer suffixes that name no other design: Office's "(MS)" export name,
 * a Paneuropean or web cut of the same face. */
const FAMILY_NOISE = [/\s*\((?:ms|body|headings)\)$/i, /\s+paneuropean$/i, /\s+web\s*\([^)]*\)$/i];









export function styledFamilyCandidates(family: string): readonly StyledFamily[] {
  let name = family.trim().replace(/^["']|["']$/g, "").trim();
  let weight = 400;
  let italic = false;
  const out: StyledFamily[] = [];
  for (let guard = 0; guard < 6 && name; guard++) {
    const before = name;
    for (const noise of FAMILY_NOISE) name = name.replace(noise, "").trim();
    const words = name.split(/\s+/);
    // "Semi Bold", "Extra Light" and "Semi-Bold" are one style word.
    const joinedTwo = words.length > 2 ? (words.at(-2)! + words.at(-1)!).toLowerCase().replace(/-/g, "") : "";
    const one = (words.at(-1) ?? "").toLowerCase().replace(/-/g, "");
    const [style, cut] = STYLE_WORDS[joinedTwo] ? [STYLE_WORDS[joinedTwo], 2] : STYLE_WORDS[one] && words.length > 1 ? [STYLE_WORDS[one], 1] : [undefined, 0];
    if (style) {
      name = words.slice(0, -cut).join(" ");
      if (style.weight !== undefined && style.weight !== 400) weight = Math.max(weight === 400 ? 0 : weight, style.weight);
      if (style.italic) italic = true;
    }
    if (name === before) break;
    if (name) out.push({ base: name, weight, italic });
  }
  return out;
}

export function officeFontSubstitutes(
  family: string | null | undefined,
): readonly string[] {
  const own = officeFontSubstitution(family)?.substitutes ?? [];
  if (own.length || !family) return own;
  // A style-named open face ("Poppins Medium"): the face itself, which the
  // registry serves under its own name.
  const styled = styledFamilyCandidates(family).at(-1);
  if (styled && !powerPointFaces[normalize(family)]) return [styled.base];
  // A face neither Office nor this page has: the face PowerPoint draws in its
  
  const out: string[] = [];
  for (const face of powerPointSubstitutes(family)) {
    for (const name of [face, ...(officeFontSubstitution(face)?.substitutes ?? [])]) {
      if (!out.some((seen) => normalize(seen) === normalize(name))) out.push(name);
    }
  }
  return out;
}

interface PowerPointSubstitution {
  readonly face: string;
  /** By script (latin / ea / cs): the face PowerPoint drew, with its evidence. */
  readonly by: Readonly<Record<string, { readonly substitute: string; readonly decks: readonly string[] }>>;
}
const powerPointFaces = (powerPointTable as { faces: Record<string, PowerPointSubstitution> }).faces;







export function powerPointSubstitutes(family: string | null | undefined): readonly string[] {
  if (!family) return [];
  const entry = powerPointFaces[normalize(family)];
  if (!entry) return [];
  const out: string[] = [];
  for (const script of ["latin", "ea", "cs"]) {
    const face = entry.by[script]?.substitute;
    if (face && !out.includes(face)) out.push(face);
  }
  return out;
}

let officeFamilyList: readonly string[] | undefined;

/** Every Office family name the catalog covers, for pickers to list. The
 * table is static, so the sorted list is built once: a font picker asks for
 * it on every ribbon repaint, and the collation sort is not free. */
export function officeFontFamilies(): readonly string[] {
  officeFamilyList ??= Object.freeze(
    Object.values(table)
      .filter((entry) => entry.substitutes.length && !entry.unlisted)
      .map((entry) => entry.office)
      .filter((office, index, all) => all.indexOf(office) === index)
      .sort((a, b) => a.localeCompare(b)),
  );
  return officeFamilyList;
}

/**
 * The Office family a font program's own family name stands for: the name
 * itself when it is one (Arial), else the Office family whose first
 * substitute it is (a PDF set in a served face names "Noto Serif SC", which
 * SimSun and STZhongsong are drawn in), the one whose regular weight is
 * nearest `weight`. Undefined for a family the catalog has no face of.
 */
export function officeFamilyForFace(
  family: string | null | undefined,
  weight = 400,
): string | undefined {
  if (!family) return undefined;
  const own = officeFontSubstitution(family);
  if (own && own.substitutes.length) return own.office;
  const wanted = normalize(family);
  let best: OfficeFontSubstitution | undefined;
  for (const entry of Object.values(table)) {
    if (!entry.substitutes.length || normalize(entry.substitutes[0]) !== wanted) continue;
    const distance = Math.abs((entry.weight ?? 400) - weight);
    const bestDistance = best ? Math.abs((best.weight ?? 400) - weight) : Infinity;
    if (distance < bestDistance || (distance === bestDistance && best && entry.office.localeCompare(best.office) < 0)) {
      best = entry;
    }
  }
  return best?.office;
}

/**
 * True when a stylesheet or script has declared the family in the document's
 * FontFaceSet. `FontFaceSet.check` is not this: it answers false for a
 * declared face that has not loaded yet, and true for a family nothing
 * declares. A declared family is one the page will render with its own bytes,
 * loaded or not, so callers can treat it as present.
 */
export function officeFontDeclared(
  family: string,
  doc: Document | undefined = typeof document === "undefined"
    ? undefined
    : document,
): boolean {
  const fonts = doc?.fonts;
  if (!fonts) return false;
  const wanted = normalize(family);
  for (const face of fonts) {
    if (normalize(face.family) === wanted) return true;
  }
  return false;
}

const registeredNames = new WeakMap<Document, Set<string>>();
let indexed: { registry: Registry; size: number; index: Map<string, string> } | null = null;

/** The registry's families and substitutes by normalized name, rebuilt as
 * served parts add to it. */
function registryIndexOf(registry: Registry): Map<string, string> {
  const families = Object.keys(registry.families ?? {});
  const substitutes = Object.keys(registry.substitutes ?? {});
  const size = families.length + substitutes.length;
  if (!indexed || indexed.registry !== registry || indexed.size !== size) {
    indexed = {
      registry,
      size,
      index: new Map([...families, ...substitutes].map((name) => [normalize(name), name])),
    };
  }
  return indexed.index;
}

/**
 * Word's Thai faces for text in a run that names no complex-script font
 * (Word 16 for Mac, wp-fix-a/thai.docx): a serif run's Thai is set in
 * Angsana New, a sans run's in Browallia New. Their glyphs are small for
 * their em; the served stand-ins are scaled to the same advance. "ภาษาไทย"
 * at 11 pt, read from the advances Word printed (thai-word.pdf: the next
 * run starts 37.57px / 41.04px after the Thai one; Angsana New's own hmtx
 * gives 37.58px): Angsana New 37.57px, Noto Serif Thai 57.31px; Browallia
 * New 41.04px, Noto Sans Thai 52.22px. The scaled face is the stand-in itself
 * (no local() source), so a size-adjust never shrinks a real Angsana.
 */
const SCALED_THAI_FACES: Readonly<Record<string, string>> = {
  "angsana new": "65.56%",
  angsanaupc: "65.56%",
  // Fitted to the one word above, 78.59%; over running Thai the two faces
  
  // HarfBuzz in Browallia New's own file and in the served Noto Sans Thai:
  // 935.4 em against 1296.9 em). At 78.59% every Thai line of a Browallia
  // deck ran 9% wider than PowerPoint's and wrapped early.
  "browallia new": "72.12%",
  browalliaupc: "72.12%",
};









const SCALED_THAI_FACE_LATIN: Readonly<Record<string, string>> = {
  "browallia new": "65.41%",
  browalliaupc: "65.41%",
};
/** Whether a stand-in slice's range is a Thai one (U+0E01 .. U+0E5B). */
function thaiSlice(range: string | undefined): boolean {
  return typeof range === "string" && /U\+0?E01-0?E5B/i.test(range);
}
/**
 * The ASCII, Latin-1, quotes, ellipsis and em dash of the Chinese Office faces, drawn at Word's own advances
 * (scripts/build-cjk-latin-metric-fonts.py, public/fonts/cjk-latin): the
 * stand-ins' Latin is of other widths (Noto Serif SC sets a-z 9.6% wider
 * than SimSun, Noto Sans SC 8.1% narrower than Microsoft YaHei), so a mixed
 * Chinese and English line broke elsewhere than Word's; DengXian's quotes are
 * 0.356 em where the stand-in's are a full em, whatever the run's hint (Word
 * 16.86). Registered after the stand-in's slices, which it overrides for
 * those characters; the real face
 * still comes first through local(). By the Office family, per weight.
 */
// By the exact name a document gives the face: MS PGothic is not MS Gothic
// (its Latin is proportional), nor PMingLiU MingLiU.
const CJK_LATIN = cjkLatinFaces as {
  faces: Record<string, Record<string, string>>;
  names: Record<string, string>;
  /** A face's recut: its advances and box, drawn at the ink of Windows' own
   * face (scripts/build-cjk-latin-metric-fonts.py RECUTS). */
  recuts?: Record<string, string>;
  /** By face: exactly the characters it has (ASCII, the Latin-1 the Office
   * face has, the punctuation above), so an absent one stays the stand-in's. */
  ranges?: Record<string, string>;
};
const CJK_LATIN_PREFIX = "/fonts/cjk-latin/";
/**
 * The Latin faces of a family that has a recut (游ゴシック, its Medium and
 * Light): the recut, whose glyphs weigh what Windows' own Yu Gothic weighs.
 * The clone drawn from the catalog's slice at full weight inked Latin and
 * digits 1.37-1.42 times Word's (bold 1.10); the recut 1.01-1.03 (Word's PDF
 * of docx-look/probes/cjk15 y1, its embedded YuGothic-Regular/-Medium/-Light/
 * -Bold subsets against both, 27 characters at 256 px). Advances, cmap and
 * box are the clone's, so no line moves. `window.__officeYuLatinRecutOff =
 * true` asks the older reading (the clone), for an A/B of one build.
 */
function cjkLatinFacesOf(slug: string): Record<string, string> | undefined {
  const recut = CJK_LATIN.recuts?.[slug];
  const off = typeof window !== "undefined" && (window as Window & { __officeYuLatinRecutOff?: boolean }).__officeYuLatinRecutOff === true;
  return (recut && !off ? CJK_LATIN.faces[recut] : undefined) ?? CJK_LATIN.faces[slug];
}
const CJK_LATIN_RANGE = "U+0020-007E,U+00A0,U+00B7,U+2013-2014,U+2018-2019,U+201C-201D,U+2022,U+2026";

export const THAI_SERIF_FALLBACK_FAMILY = "Angsana New";
export const THAI_SANS_FALLBACK_FAMILY = "Browallia New";












const pendingFaces: (() => void)[] = [];
let drainingFaces = false;











const registrationsInFlight = new Map<Promise<unknown>, string>();

function trackRegistration<T>(work: Promise<T>, what: string): Promise<T> {
  registrationsInFlight.set(work, what);
  const done = () => {
    registrationsInFlight.delete(work);
  };
  work.then(done, done);
  return work;
}

/** What keeps `officeFontFacesPending` true, for a print's log: the registrations
 * in flight by what they ask for, and the faces still queued. */
export function officeFontRegistrationsPending(): string[] {
  const out = [...new Set(registrationsInFlight.values())].slice(0, 8);
  if (pendingFaces.length > 0 || madeFaces.length > 0 || drainingFaces) out.push(`faces queued ${pendingFaces.length + madeFaces.length}`);
  return out;
}

/** Whether a face registration is still under way (see `registrationsInFlight`). */
export function officeFontFacesPending(): boolean {
  return registrationsInFlight.size > 0 || pendingFaces.length > 0 || madeFaces.length > 0 || drainingFaces;
}

/**
 * Resolves once no face registration is under way (true), or at `timeoutMs`
 * with one still going (false). A registration that starts while it waits is
 * waited for too.
 */
export async function officeFontFacesSettled(timeoutMs: number): Promise<boolean> {
  const end = Date.now() + Math.max(0, timeoutMs);
  // A print's wait: faces queued as registrations answer are declared at once.
  flushOfficeFontFaces();
  while (officeFontFacesPending()) {
    flushOfficeFontFaces();
    const left = end - Date.now();
    if (left <= 0) return false;
    await Promise.race([
      Promise.allSettled([...registrationsInFlight.keys()]),
      new Promise<void>((resolve) => setTimeout(resolve, Math.min(25, left))),
    ]);
  }
  return true;
}
/** When the current task began declaring faces (-1: it has not). */
let sliceStart = -1;
const FACE_SLICE_MS = 6;

const now = (): number => (typeof performance === "undefined" ? Date.now() : performance.now());

function nextFaceSlice(run: () => void): void {
  const scheduler = (globalThis as { scheduler?: { postTask?: (task: () => void, options: { priority: string }) => unknown } }).scheduler;
  if (typeof scheduler?.postTask === "function") {
    void scheduler.postTask(run, { priority: "user-visible" });
    return;
  }
  setTimeout(run, 0);
}

/**
 * Faces made and not yet added. Making one is plain script; adding one to the
 * document invalidates the style and layout of everything on it, and the text
 * engine's index of faces. Added a slice at a time, a workbook's fonts (a CJK
 * family is a hundred slices at each weight, 1,960 faces in one open) kept the
 * whole page dirty for seconds: every read of a cell's box forced a layout of
 * the grid, on opening and on every page down. So they are made in slices and
 * added together, one invalidation.
 */
const madeFaces: {
  doc: Document;
  face: FontFace;
  src: string;
  built: Readonly<Record<string, string>>;
  declared: DeclaredFontFace;
}[] = [];

function addMadeFaces(): void {
  for (const { doc, face, src, built, declared } of madeFaces.splice(0)) {
    try {
      doc.fonts.add(face);
    } catch {
      // An engine that rejects a face leaves the family to its fallbacks.
      continue;
    }
    noteFontFaceSource(face, src, built, { doc, declared });
    if (spaceWanted.has(normalize(declared.family)) && holdsSpace(declared)) loadFace(face);
  }
}

/**
 * Families whose slice holding the space is fetched as soon as it is
 * declared ({@link loadOfficeFamilySpace}): a line's first available font is
 * the first family whose face for the space is loaded (CSS Fonts), and until
 * it is the browser goes on down the stack, resolving every face of every
 * family it passes, again for each new size, weight and style. A CJK sheet's
 * stack passes some 1,500 unloaded slices (Chromium on Linux: a font-service
 * lookup each, about 0.25 ms): its first measures of each size were 2 ms
 * apiece while the face's Latin slice, the one with the space, came last
 * (R116-X02: at 12 s).
 */
const spaceWanted = new Set<string>();

function loadFace(face: FontFace): void {
  try {
    if (typeof face.load === "function") void face.load().catch(() => {});
  } catch {
    // A face that cannot load leaves the line to the stack's next family.
  }
}

function holdsSpace(declared: DeclaredFontFace): boolean {
  const weight = declared.weight ?? "400";
  if ((declared.style ?? "normal") !== "normal" || (weight !== "400" && weight !== "700")) return false;
  if (!declared.unicodeRange) return true;
  return declared.unicodeRange.split(",").some((part) => {
    const match = /^\s*U\+([0-9A-F?]{1,6})(?:-([0-9A-F]{1,6}))?\s*$/i.exec(part);
    if (!match) return false;
    const low = parseInt(match[1].replace(/\?/g, "0"), 16);
    const high = match[2] ? parseInt(match[2], 16) : parseInt(match[1].replace(/\?/g, "F"), 16);
    return low <= 0x20 && high >= 0x20;
  });
}

/** Fetch the regular and bold slices holding the space of the faces declared
 * for `family` (now or later): it stands first in a cell's stack. */
export function loadOfficeFamilySpace(
  family: string | null | undefined,
  doc: Document | undefined = typeof document === "undefined" ? undefined : document,
): void {
  const key = family ? normalize(family) : "";
  if (!key || spaceWanted.has(key) || !doc?.fonts) return;
  spaceWanted.add(key);
  // Faces declared before it was asked for, from what was noted when they
  // were declared: the engine's getters restyle the page in Safari 16.0-16.3.
  for (const { face, declared } of notedFontFaces(doc) ?? []) {
    if (declared && normalize(declared.family) === key && holdsSpace(declared)) loadFace(face);
  }
}

/** Called once every face queued so far is in the document (`afterFacesDeclared`). */
const facesDeclaredWaiters: (() => void)[] = [];

/** Runs `done` once every face queued so far is added to its document: at
 * once when none waits, else after the drain's last slice. */
function afterFacesDeclared(done: () => void): void {
  // Only a drain under way adds faces later (the fast path adds a face as it
  // is made); a waiter queued with none would wait for good.
  if (drainingFaces || pendingFaces.length > 0) facesDeclaredWaiters.push(done);
  else done();
}

function drainFaces(): void {
  const end = now() + FACE_SLICE_MS;
  try {
    while (pendingFaces.length > 0) {
      pendingFaces.shift()!();
      if (now() >= end) break;
    }
  } finally {
    if (pendingFaces.length > 0) nextFaceSlice(drainFaces);
    else {
      drainingFaces = false;
      addMadeFaces();
      for (const done of facesDeclaredWaiters.splice(0)) done();
    }
  }
}









export function flushOfficeFontFaces(): void {
  if (!drainingFaces && pendingFaces.length === 0) return;
  while (pendingFaces.length > 0) pendingFaces.shift()!();
  // The drain's own next slice finds nothing left and ends.
  drainingFaces = false;
  addMadeFaces();
  for (const done of facesDeclaredWaiters.splice(0)) done();
}

function addFace(...args: Parameters<typeof declareFace>): void {
  if (!drainingFaces) {
    const at = now();
    if (sliceStart < 0) {
      sliceStart = at;
      // The task's slice ends with the task.
      nextFaceSlice(() => {
        sliceStart = -1;
      });
    }
    if (at - sliceStart < FACE_SLICE_MS) {
      declareFace(...args);
      addMadeFaces();
      return;
    }
    drainingFaces = true;
    nextFaceSlice(drainFaces);
  }
  pendingFaces.push(() => declareFace(...args));
}

/** Whether this engine applies a face's `font-variation-settings` descriptor
 * (Chromium does; WebKit has no such FontFace attribute and ignores it in
 * a rule too). Read once. */
let variationHonored: boolean | undefined;
function faceVariationHonored(): boolean {
  if (variationHonored === undefined) {
    variationHonored = typeof FontFace === "undefined" || "variationSettings" in FontFace.prototype;
  }
  return variationHonored;
}

function declareFace(
  doc: Document,
  family: string,
  weight: number,
  italic: 0 | 1,
  sources: readonly string[],
  range: string | undefined,
  sizeAdjust?: string,
  overrides?: readonly [string, string, string],
  drawnWeight?: number,
) {
  // A face declared at one weight and drawn from a stand-in at another
  // ("Avenir Heavy" regular = Nunito Sans 800): most stand-ins are variable
  // fonts, one file for every weight, and the browser sets their wght axis
  // to the asked weight clamped to the face's weight descriptor (CSS Fonts 4
  
  
  // own weight is stated as the face's wght; a static file has no such axis
  // and is drawn as it is.
  let variationSettings = drawnWeight !== undefined && drawnWeight !== weight ? `"wght" ${drawnWeight}` : undefined;
  let declaredWeight = weight;
  if (variationSettings && drawnWeight !== undefined && !faceVariationHonored()) {
    // WebKit (Safari, Playwright WebKit 26) neither exposes nor applies a
    // face's font-variation-settings descriptor, in a FontFace or an
    // @font-face rule: the stand-in drew at the asked weight, its axis
    
    // drawn at 400, the title 1251 px where Chromium draws it at 1320
    // px; Montserrat Semi-Bold Bold's regular is drawn at 700). It does
    // clamp the axis to the face's weight descriptor, so the face is
    // declared at the weight it is drawn at (CSS Fonts 4 §5.2: a regular
    // request finds the lightest face above 500 when none is at 400-500, so
    // a regular drawn at 600-900 is still found), a bold one at no less than
    // 700 so a bold request finds it. A regular drawn at 700 or more is then
    // also what a bold request finds first: such a family's bold runs draw at
    // the regular's weight (Montserrat Semi-Bold Bold's at 700, not 900) --
    // the runs that name the style are the regular ones. The text engine
    // reads the declared weight (no stated wght) and so measures what WebKit
    // draws.
    variationSettings = undefined;
    declaredWeight = weight >= 600 ? Math.max(drawnWeight, 700) : drawnWeight;
  }
  try {
    const face = new FontFace(family, sources.join(","), {
      weight: String(declaredWeight),
      style: italic ? "italic" : "normal",
      display: "swap",
      ...(range ? { unicodeRange: range } : {}),
      ...(sizeAdjust ? { sizeAdjust } : {}),
      // Word's line box rather than the stand-in's: a line set in it is as
      // tall as Word sets it.
      ...(overrides ? { ascentOverride: overrides[0], descentOverride: overrides[1], lineGapOverride: overrides[2] } : {}),
      ...(variationSettings ? { variationSettings } : {}),
    } as FontFaceDescriptors);
    // Added with the slice's other faces (addMadeFaces), and noted then with
    // what was declared, so no reader has to ask the face: its getters
    // restyle the page in Safari 16.0-16.3 (font-face-declarations).
    madeFaces.push({
      doc,
      face,
      src: sources.join(","),
      built: {
        ...(sizeAdjust ? { sizeAdjust } : {}),
        ...(overrides ? { ascentOverride: overrides[0], descentOverride: overrides[1], lineGapOverride: overrides[2] } : {}),
        ...(variationSettings ? { variationSettings } : {}),
      },
      declared: { family, weight: String(declaredWeight), style: italic ? "italic" : "normal", ...(range ? { unicodeRange: range } : {}) },
    });
  } catch {
    // An engine that rejects a descriptor leaves the family to its fallbacks.
  }
}

const localFaceProbes = new Map<string, Promise<boolean>>();
const LOCAL_FACE_PROBE_MS = 3000;
/** Whether any of these local() sources names a face this machine has: a
 * FontFace of local sources alone loads only when one of them is found. No
 * answer within LOCAL_FACE_PROBE_MS counts as not installed. */
function localFaceInstalled(sources: readonly string[]): Promise<boolean> {
  const key = sources.join(",");
  let probe = localFaceProbes.get(key);
  if (!probe) {
    try {
      const loaded = new FontFace("mrai-local-face-probe", key).load().then(() => true, () => false);
      const late = new Promise<boolean>((resolve) => setTimeout(() => resolve(false), LOCAL_FACE_PROBE_MS));
      probe = Promise.race([loaded, late]);
    } catch {
      probe = Promise.resolve(false);
    }
    localFaceProbes.set(key, probe);
  }
  return probe;
}

/**
 * Stand-ins the site's own stylesheet declares under the same name, with the
 * same slices at every weight the registry serves and more
 * (public/fonts/site/site-fonts.css: Noto Sans SC at 300-700, the same 101
 * ranges). Declared again here, a family had two sets of faces: Microsoft
 * YaHei's stand-in made Noto Sans SC 808 faces, every one of which Blink
 * consults, unloaded, at each new size, weight and style the page lays out
 * (R116-X02: 3,507 faces, a 2 s relayout after a font loaded).
 */
const SITE_DECLARED_STAND_INS = new Set(["noto sans sc"]);

function registerSubstitute(doc: Document, registry: Registry, name: string) {
  const names = registeredNames.get(doc)!;
  if (names.has(normalize(name))) return;
  names.add(normalize(name));
  if (SITE_DECLARED_STAND_INS.has(normalize(name))) return;
  for (const [weight, italic, out, range] of registry.substitutes[name] ?? []) {
    addFace(doc, name, weight, italic, [`url(${PUBLIC_PREFIX}${out}) format("woff2")`], range === null ? undefined : registry.ranges[range]);
  }
}

/** The slices of `substitute` nearest a weight and style. */
function nearestSlices(registry: Registry, substitute: string, weight: number, italic: boolean): RegistrySubstituteFace[] {
  const slices = registry.substitutes[substitute] ?? [];
  const styled = slices.filter((slice) => Boolean(slice[1]) === italic);
  const pool = styled.length ? styled : slices;
  if (!pool.length) return [];
  const best = pool.reduce((nearest, slice) => (Math.abs(slice[0] - weight) < Math.abs(nearest - weight) ? slice[0] : nearest), pool[0][0]);
  return pool.filter((slice) => slice[0] === best);
}

/**
 * Whether a style-named family ("Poppins SemiBold", "Roboto Condensed
 * Light") is drawn bold by emboldening its regular. Named for a
 * weight other than regular or bold, it is a family of its own whose one
 * style is that weight (its regular, and an italic): Windows lists every such
 * family it has with no bold face beside it (windows-family-styles.json: 163
 * of 163 on the truth machine), and Office emboldens the regular, at the
 * regular's advances (PowerPoint's PDF: "Poppins SemiBold,Bold").
 */
export function styledFamilyEmboldensRegular(styled: Pick<StyledFamily, "weight">): boolean {
  // Heavier than bold (ExtraBold, Black, Heavy: 800 and up), Windows draws the face as it is: no bold run of one in
  // its PDFs is emboldened (Lato Black, Poppins ExtraBold, Roboto Black ...), and its bold stays the heavy cut.
  return styled.weight !== 400 && styled.weight !== 700 && styled.weight < 800;
}

/**
 * A style-named family ("Poppins Medium") as a family of its own: its regular
 * drawn by the base family's face at the named weight. Its bold is that
 * regular emboldened where the family has no bold face
 * (styledFamilyEmboldensRegular: no bold face is declared, and the browser
 * synthesizes it), else the next heavier cut. The real face is tried first
 * through local().
 */
function registerStyledAlias(doc: Document, registry: Registry, family: string, base: string, styled: StyledFamily) {
  const local = `local("${family.replace(/["\\]/g, "\\$&")}")`;
  const heavier = styled.weight >= 600 ? 900 : 700;
  // An Office family: its first stand-in; an open face: itself.
  const substitute = registry.families[base]?.[0]?.substitute ?? (registry.substitutes[base] ? base : undefined);
  if (!substitute) return;
  registerSubstitute(doc, registry, substitute);
  const styles = styledFamilyEmboldensRegular(styled) ? [[400, styled.weight]] as const : [[400, styled.weight], [700, heavier]] as const;
  for (const [aliasWeight, weight] of styles) {
    for (const [sliceWeight, , out, range] of nearestSlices(registry, substitute, weight, styled.italic)) {
      addFace(doc, family, aliasWeight, styled.italic ? 1 : 0, [local, `url(${PUBLIC_PREFIX}${out}) format("woff2")`], range === null ? undefined : registry.ranges[range], undefined, undefined, sliceWeight);
    }
  }
}

/**
 * Moves on each time a family or alias family was let go to be asked for
 * again (not answered yet, refused): a caller that remembers the stacks it
 * built (`presentationFontFamilyWithAssets`) builds them afresh, so the
 * declaration they make is asked for again as before.
 */
let registrationGeneration = 0;
export function officeFontRegistrationGeneration(): number {
  return registrationGeneration;
}

const aliasFamilies = new WeakMap<Document, Set<string>>();

/**
 * One family drawn by several served faces, each for the characters its
 * slices cover (`script-fallbacks.ts`: "MRAI Script Fallback" is Noto Sans
 * for Latin, Greek and Cyrillic, Noto Sans Arabic for Arabic, ...). `faces`
 * in priority order: where two cover a character the earlier one draws it
 * (a family's faces are matched last-declared first, so they are declared in
 * reverse). Regular and bold, upright and italic, as each face serves them.
 */
export function ensureOfficeAliasFamily(
  alias: string,
  faces: readonly string[],
  doc: Document | undefined = typeof document === "undefined" ? undefined : document,
): void {
  if (!doc?.fonts || typeof FontFace === "undefined") return;
  let done = aliasFamilies.get(doc);
  if (!done) aliasFamilies.set(doc, (done = new Set()));
  if (done.has(alias)) return;
  done.add(alias);
  const settle = noteFamilyRegistration(alias);
  void trackRegistration(loadOfficeDataParts("font-registry", faces.map(normalize))
    .then((served) => {
      const registry = served as Registry | undefined;
      if (!registry) {
        done.delete(alias);
        registrationGeneration += 1;
        return;
      }
      const index = registryIndexOf(registry);
      for (const face of [...faces].reverse()) {
        const name = index.get(normalize(face));
        for (const [weight, italic, out, range] of (name && registry.substitutes[name]) || []) {
          if (weight !== 400 && weight !== 700) continue;
          addFace(doc, alias, weight, italic, [`url(${PUBLIC_PREFIX}${out}) format("woff2")`], range === null ? undefined : registry.ranges[range]);
        }
      }
    })
    .catch(() => {
      done.delete(alias);
      registrationGeneration += 1;
    })
    .finally(() => afterFacesDeclared(settle)), `registry ${alias}`);
}

/**
 * `alias`: the faces the document's own stylesheets declare for `families`
 * (in that order: a later family's face is tried first where both hold a
 * character, as CSS tries a family's faces), each declared again for the part
 * of its unicode-range inside `ranges` alone -- a family that draws those
 * characters and lets every other fall through to the stack's next face.
 * Registered once per document, when the rules are there to read (a sheet
 * not yet loaded is read on the next call).
 */
export function ensureOfficeRangedAliasFamily(
  alias: string,
  families: readonly string[],
  ranges: readonly (readonly [number, number])[],
  doc: Document | undefined = typeof document === "undefined" ? undefined : document,
): void {
  if (!doc?.fonts || typeof FontFace === "undefined") return;
  let done = aliasFamilies.get(doc);
  if (!done) aliasFamilies.set(doc, (done = new Set()));
  if (done.has(alias)) return;
  const wanted = families.map(normalize);
  const rules: { family: number; weight: number; italic: 0 | 1; src: string; range: string }[] = [];
  for (const sheet of Array.from(doc.styleSheets)) {
    let list: CSSRuleList;
    try {
      list = sheet.cssRules;
    } catch {
      continue; // another origin's sheet
    }
    for (const rule of Array.from(list)) {
      if (!(typeof CSSFontFaceRule !== "undefined" && rule instanceof CSSFontFaceRule)) continue;
      const family = wanted.indexOf(normalize(rule.style.getPropertyValue("font-family")));
      const src = rule.style.getPropertyValue("src");
      if (family < 0 || !src) continue;
      const own = parseUnicodeRange(rule.style.getPropertyValue("unicode-range")) ?? [[0, 0x10ffff]];
      const kept: string[] = [];
      for (const [a, b] of own) {
        for (const [c, d] of ranges) {
          const lo = Math.max(a, c), hi = Math.min(b, d);
          if (lo <= hi) kept.push(lo === hi ? `U+${lo.toString(16).toUpperCase()}` : `U+${lo.toString(16).toUpperCase()}-${hi.toString(16).toUpperCase()}`);
        }
      }
      if (!kept.length) continue;
      const weight = rule.style.getPropertyValue("font-weight").trim() === "bold" ? 700 : Number.parseInt(rule.style.getPropertyValue("font-weight"), 10) || 400;
      const italic = /italic|oblique/.test(rule.style.getPropertyValue("font-style")) ? 1 : 0;
      rules.push({ family, weight, italic, src, range: kept.join(", ") });
    }
  }
  if (!rules.length) return;
  done.add(alias);
  const settle = noteFamilyRegistration(alias);
  for (const rule of rules.sort((x, y) => x.family - y.family)) {
    addFace(doc, alias, rule.weight, rule.italic, [rule.src], rule.range);
  }
  afterFacesDeclared(settle);
}

/**
 * One script's letters in a served face at a size of its own: `alias` is the
 * slices of `face` that hold `letter`, declared for `range` alone and set at
 * `sizeAdjust` (regular and bold, as the face serves them). A family placed
 * in a stack ahead of faces without the script sets its letters at the
 * widths of the face an Office application draws them in (the sheet's
 * complex-script scales, lib/spreadsheet/office-complex-scales.ts). No
 * local() source: the size-adjust is the stand-in's alone.
 *
 * `unlessInstalled`: an Office family the reader may have (a PowerPoint
 * bullet's Cambria Math, TRIAGE 99): the scaled stand-in is declared only once
 * the probe finds no installed face of that family (`localFaceInstalled`, as
 * a Word-scaled stand-in waits for it), so a machine that has the real face
 * draws it at its own size -- the stack names the scaled family ahead of the
 * Office family, which draws wherever the scaled one is not declared.
 */
export function ensureOfficeScaledScriptFamily(
  alias: string,
  face: string,
  letter: number,
  range: string,
  sizeAdjust: string,
  doc: Document | undefined = typeof document === "undefined" ? undefined : document,
  unlessInstalled?: string,
): void {
  if (!doc?.fonts || typeof FontFace === "undefined") return;
  let done = aliasFamilies.get(doc);
  if (!done) aliasFamilies.set(doc, (done = new Set()));
  if (done.has(alias)) return;
  done.add(alias);
  const office = unlessInstalled ? normalize(unlessInstalled) : undefined;
  void loadOfficeDataParts("font-registry", office ? [normalize(face), office] : [normalize(face)])
    .then(async (served) => {
      const registry = served as Registry | undefined;
      const indexed = registry ? registryIndexOf(registry).get(normalize(face)) : undefined;
      // Only a served substitute's slices are the registry's to give: a face the static stylesheet declares
      // (styles/office-fonts.css: Noto Sans Math, STIX Two Math, Cambria Math's own rules) -- and one the registry
      // knows as an Office family only, its faces' local() names without a served file (DejaVu Sans, Noto Sans
      // Symbols) -- is declared from its rules there.
      const name = indexed && registry?.substitutes?.[indexed] ? indexed : undefined;
      const statics = name ? [] : staticFaceRules(doc, face);
      if (!name && !statics.length) {
        done.delete(alias);
        registrationGeneration += 1;
        return;
      }
      if (office) {
        const officeName = registry ? registryIndexOf(registry).get(office) : undefined;
        const locals = (registry && officeName ? registry.families[officeName] ?? [] : [])
          .flatMap(({ faces }) => faces.filter(([weight, italic]) => weight === 400 && !italic).flatMap(([, , , , names]) => names))
          .map((local) => `local("${local.replace(/["\\]/g, "\\$&")}")`);
        const sources = locals.length ? locals : staticFaceRules(doc, unlessInstalled as string).flatMap((rule) => rule.locals);
        if (sources.length && (await localFaceInstalled([...new Set(sources)]))) return;
      }
      if (!name) {
        for (const rule of statics) {
          if (rule.weight !== 400 && rule.weight !== 700) continue;
          if (rule.ranges && !rule.ranges.some(([start, end]) => letter >= start && letter <= end)) continue;
          if (rule.urls.length) addFace(doc, alias, rule.weight, rule.italic, rule.urls, range, sizeAdjust);
        }
        return;
      }
      if (!registry) return;
      for (const [weight, italic, out, sliceRange] of registry.substitutes[name] ?? []) {
        if (weight !== 400 && weight !== 700) continue;
        const ranges = parseUnicodeRange(sliceRange === null ? undefined : registry.ranges[sliceRange]);
        if (ranges && !ranges.some(([start, end]) => letter >= start && letter <= end)) continue;
        addFace(doc, alias, weight, italic, [`url(${PUBLIC_PREFIX}${out}) format("woff2")`], range, sizeAdjust);
      }
    })
    .catch(() => {
      done.delete(alias);
      registrationGeneration += 1;
    });
}

/** A family's @font-face rules in the document's own stylesheets: weight, slant, unicode-range, url() and local() sources apart. */
function staticFaceRules(
  doc: Document,
  family: string,
): { weight: number; italic: 0 | 1; ranges: readonly (readonly [number, number])[] | null; urls: string[]; locals: string[] }[] {
  const wanted = normalize(family);
  const out: { weight: number; italic: 0 | 1; ranges: readonly (readonly [number, number])[] | null; urls: string[]; locals: string[] }[] = [];
  for (const sheet of Array.from(doc.styleSheets ?? [])) {
    let list: CSSRuleList;
    try {
      list = sheet.cssRules;
    } catch {
      continue; // another origin's sheet
    }
    for (const rule of Array.from(list)) {
      if (!(typeof CSSFontFaceRule !== "undefined" && rule instanceof CSSFontFaceRule)) continue;
      if (normalize(rule.style.getPropertyValue("font-family")) !== wanted) continue;
      const src = rule.style.getPropertyValue("src");
      const parts = src.match(/(?:local|url)\([^)]*\)(?:\s*format\([^)]*\))?/g) ?? [];
      const weight = rule.style.getPropertyValue("font-weight").trim() === "bold" ? 700 : Number.parseInt(rule.style.getPropertyValue("font-weight"), 10) || 400;
      out.push({
        weight,
        italic: /italic|oblique/.test(rule.style.getPropertyValue("font-style")) ? 1 : 0,
        ranges: parseUnicodeRange(rule.style.getPropertyValue("unicode-range")),
        urls: parts.filter((part) => part.startsWith("url(")),
        locals: parts.filter((part) => part.startsWith("local(")),
      });
    }
  }
  return out;
}

/**
 * A sheet's cut family for `face` (excel-cut-families.ts): the faces the
 * registry declares for `face` -- the same local() names first, the same line
 * overrides, the same slices' ranges -- declared again under the family's own
 * name at the weight Windows Excel's ink weighs (the slices are variable:
 * one holding East Asian characters is set to the `cjk` weight, any other
 * to the `latin`), and the face's recut Latin clone over them. A face Windows has no bold of gets no bold
 * face: its bold is its regular emboldened, as Excel prints it. Nothing of
 * `face`'s own family changes. Until the registry has answered and the faces
 * are in the document the text engine waits for the family
 * (`noteFamilyRegistration`), as it does for `face`; a registry that cannot
 * be had leaves the family undeclared, and the stack's next family -- `face`
 * -- draws, as before.
 */
export function ensureExcelCutFamily(
  face: string | null | undefined,
  doc: Document | undefined = typeof document === "undefined" ? undefined : document,
): void {
  const cut = excelCutFamilyOf(face);
  if (!cut || !doc?.fonts || typeof FontFace === "undefined") return;
  let done = aliasFamilies.get(doc);
  if (!done) aliasFamilies.set(doc, (done = new Set()));
  const aliasKey = normalize(cut.alias);
  if (done.has(aliasKey)) return;
  done.add(aliasKey);
  const key = normalize(cut.face);
  const again = () => {
    done.delete(aliasKey);
    registrationGeneration += 1;
  };
  const settle = noteFamilyRegistration(cut.alias);
  void trackRegistration(loadOfficeDataParts("font-registry", [key])
    .then((served) => {
      const registry = served as Registry | undefined;
      const name = registry ? registryIndexOf(registry).get(key) : undefined;
      const families = registry && name ? registry.families[name] : undefined;
      if (!registry || !families?.length) {
        // Not answered yet (no connection, refused): a later call asks again.
        if (!registry || (officeDataObtainable("font-registry") && !officeDataPartKnown("font-registry", key))) again();
        return;
      }
      const clones = new Set<string>();
      for (const { substitute, faces } of families) {
        const slices = registry.substitutes[substitute] ?? [];
        for (const [weight, italic, subWeight, subItalic, locals, overrides, sizeAdjust] of faces) {
          const style = weight >= 600 ? cut.bold : cut.regular;
          // A stand-in scaled to Word's widths is declared apart from its
          // local() face (ensureOfficeFontFaces); no measured face is one.
          if (!style || sizeAdjust) continue;
          const localSources = locals.map((local) => `local("${local.replace(/["\\]/g, "\\$&")}")`);
          for (const [sliceWeight, sliceItalic, out, range] of slices) {
            if (sliceWeight !== subWeight || sliceItalic !== subItalic) continue;
            const sliceRange = range === null ? undefined : registry.ranges[range];
            const drawn = excelCutRangeIsEastAsian(parseUnicodeRange(sliceRange)) ? style.cjk : style.latin;
            addFace(doc, cut.alias, weight, italic, [...localSources, `url(${PUBLIC_PREFIX}${out}) format("woff2")`], sliceRange, undefined, overrides ?? undefined, drawn);
          }
          // The face's Latin at the Office face's advances, after the slices
          // it overrides for those characters.
          const file = cut.latin[weight >= 600 ? "700" : "400"];
          const id = `${weight}/${italic}`;
          if (!file || clones.has(id)) continue;
          clones.add(id);
          addFace(doc, cut.alias, weight, italic, [...localSources, `url(${CJK_LATIN_PREFIX}${file}) format("woff2")`], cut.latinRange ?? CJK_LATIN_RANGE, undefined, overrides ?? undefined);
        }
      }
    })
    .catch(again)
    .finally(() => afterFacesDeclared(settle)), `registry ${cut.alias}`);
}

/** By family (normalized): the regular face's Word line (ascent + descent +
 * line gap, em), the line box its overrides give a stand-in. */
const WORD_LINE_EM = new Map<string, number>();

/** Notes the family's Word line from its served faces (exported for tests). */
export function noteWordLineEm(faces: readonly RegistryAliasFace[], ...families: string[]): void {
  const regular = faces.find(([weight, italic, , , , overrides]) => weight === 400 && italic === 0 && overrides);
  const overrides = regular?.[5];
  if (!overrides) return;
  const em = overrides.reduce((sum, part) => sum + Number.parseFloat(part), 0) / 100;
  if (!(em > 0) || !Number.isFinite(em)) return;
  for (const family of families) WORD_LINE_EM.set(normalize(family), em);
}

let metricOverridesHonored: boolean | undefined;
/**
 * Whether the engine sets a line in a face's ascent-, descent- and
 * line-gap-override (Chromium, Firefox). WebKit (Safari, every iPad browser)
 * ignores the three descriptors: a stand-in keeps its own, taller line --
 * Noto Serif SC's for 宋体 is 1.5 em where Word's SimSun is 1 em -- and a
 * line of the font's normal height overflowed its cell above.
 */
export function fontMetricOverridesHonored(): boolean {
  if (metricOverridesHonored === undefined) {
    metricOverridesHonored = typeof FontFace === "undefined" || "ascentOverride" in FontFace.prototype;
  }
  return metricOverridesHonored;
}

/**
 * The height of a `line-height: normal` line in an Office family at
 * `fontSizePx`, as an engine that honours the stand-in's overrides sets it
 * (Word's line), for an engine that does not; null where the engine honours
 * them or the family has none (its own normal line is the right one).
 */
export function officeNormalLineHeightPx(fontFamily: string | null | undefined, fontSizePx: number): number | null {
  if (fontMetricOverridesHonored() || !fontFamily || !(fontSizePx > 0)) return null;
  // The family asked for: a Hangul clone in front of it (hangul-em.ts) has no line of its own.
  const em = WORD_LINE_EM.get(normalize(stackPrimaryFamily(fontFamily)));
  return em === undefined ? null : em * fontSizePx;
}

/**
 * The cut a stand-in draws its Office face's emboldened regular in, printed:
 * Noto Sans JP's regular strokes are lighter than ＭＳ Ｐゴシック's, so the
 * bold Excel prints (its regular, filled and stroked) is set in Noto's 500
 * under the same stroke, as the API's sheet writer sets it (office_pdf_fonts
 * EMBOLDENED_REGULAR_WEIGHTS; lookwin R121-X08's title in Excel's printed PDF:
 * 229.7 units of ink, Noto 400 stroked 184.6, 500 stroked 217.3).
 */
const EMBOLDENED_REGULAR_WEIGHTS: Readonly<Record<string, number>> = { "Noto Sans JP": 500 };
let emboldenedRegularCuts = false;
/** Whether a family with no bold face draws its bold from its stand-in's
 * heavier cut (EMBOLDENED_REGULAR_WEIGHTS): the print page, set with the
 * printer's resolution (native-typography setNativeTextPrinterDpi), before
 * any family is declared. */
export function setOfficeEmboldenedRegularCuts(on: boolean): void {
  emboldenedRegularCuts = on;
}

/** An alias family's faces, with a bold drawn from the stand-in's
 * emboldened cut where the family has no bold face and it prints. */
function withEmboldenedCuts(
  family: string,
  substitute: string,
  faces: readonly RegistryAliasFace[],
  slices: readonly RegistrySubstituteFace[],
): readonly RegistryAliasFace[] {
  const cut = EMBOLDENED_REGULAR_WEIGHTS[substitute];
  if (!emboldenedRegularCuts || cut === undefined || !officeFamilyHasNoBold(family)) return faces;
  if (!slices.some((slice) => slice[0] === cut)) return faces;
  const added: RegistryAliasFace[] = [];
  for (const face of faces) {
    if (face[0] !== 400 || faces.some((other) => other[0] >= 600 && other[1] === face[1])) continue;
    if (!slices.some((slice) => slice[0] === cut && slice[1] === face[3])) continue;
    added.push([700, face[1], cut, face[3], face[4], face[5], face[6]]);
  }
  return added.length ? [...faces, ...added] : faces;
}

/**
 * Declares the faces behind an Office family name (and its substitutes) in
 * the document, once, so text set in that family renders with the served
 * open face wherever the real one is not installed. Idempotent and cheap to
 * call from stack builders; a name the registry does not know is ignored.
 */
export function ensureOfficeFontFaces(
  family: string | null | undefined,
  doc: Document | undefined = typeof document === "undefined"
    ? undefined
    : document,
): void {
  if (!family || !doc?.fonts || typeof FontFace === "undefined") return;
  const key = normalize(family);
  // A sheet's cut family is declared with its face (ensureExcelCutFamily), not asked of the registry.
  if (!key || isExcelCutFamily(key)) return;
  let names = registeredNames.get(doc);
  if (!names) registeredNames.set(doc, (names = new Set()));
  if (names.has(key)) return;
  names.add(key);
  const styled = table[key] ? [] : styledFamilyCandidates(family);
  // Until the registry has answered and the faces it names are in the
  // document, the text engine waits for this family rather than measuring the
  // stack's next one (`noteFamilyRegistration`).
  const settle = noteFamilyRegistration(key);
  void trackRegistration(loadOfficeDataParts("font-registry", [key, ...styled.map((candidate) => normalize(candidate.base))])
    .then((served) => {
      const registry = served as Registry | undefined;
      const name = registry ? registryIndexOf(registry).get(key) : undefined;
      if (!name && registry) {
        const index = registryIndexOf(registry);
        const candidate = styled.find((entry) => index.has(normalize(entry.base)));
        if (candidate) {
          registerStyledAlias(doc, registry, family.trim().replace(/^["']|["']$/g, "").trim(), index.get(normalize(candidate.base))!, candidate);
          return;
        }
      }
      if (!name) {
        // Not answered yet (no connection, refused): a later call asks again.
        // A family the registry does not know stays settled.
        if (officeDataObtainable("font-registry") && !officeDataPartKnown("font-registry", key)) {
          names.delete(key);
          registrationGeneration += 1;
        }
        return;
      }
      if (!registry) return;
      if (registry.substitutes[name]) {
        names.delete(key);
        registerSubstitute(doc, registry, name);
        return;
      }
      const families = (registry.families[name] ?? []).map(({ substitute, faces }) => ({
        substitute,
        faces: withEmboldenedCuts(name, substitute, faces, registry.substitutes[substitute] ?? []),
      }));
      for (const { substitute, faces } of families) {
        noteWordLineEm(faces, key, name);
        registerSubstitute(doc, registry, substitute);
        const slices = registry.substitutes[substitute] ?? [];
        const scaled = SCALED_THAI_FACES[key];
        for (const [weight, italic, subWeight, subItalic, locals, overrides, sizeAdjust] of faces) {
          const localSources = scaled ? [] : locals.map((local) => `local("${local.replace(/["\\]/g, "\\$&")}")`);
          const addSlices = (sources: readonly string[]) => {
            for (const [sliceWeight, sliceItalic, out, range] of slices) {
              if (sliceWeight !== subWeight || sliceItalic !== subItalic) continue;
              // A scaled stand-in (the Thai faces) keeps its own scaled box;
              // its Latin slices take the face's own Latin scale where known.
              const sliceRange = range === null ? undefined : registry.ranges[range];
              const sliceScale = scaled && !thaiSlice(sliceRange) ? SCALED_THAI_FACE_LATIN[key] ?? scaled : scaled;
              addFace(doc, name, weight, italic, [...sources, `url(${PUBLIC_PREFIX}${out}) format("woff2")`], sliceRange, sliceScale ?? sizeAdjust, scaled ? undefined : overrides ?? undefined, subWeight);
            }
          };
          if (!sizeAdjust || !localSources.length) {
            addSlices(localSources);
            continue;
          }
          // A stand-in scaled to Word's widths (standInWidth) is for a machine
          // without the Office face: a size-adjust on a face that also lists
          // local() would shrink the real font too. The real face, where it is
          // installed, is the family alone, at its own size; elsewhere the
          // scaled stand-in. (Franklin Gothic Book scaled with no local()
          
          
          // at once, so a machine that has it never paints anything else; the
          // scaled stand-in joins it only once the probe finds no such face
          // (or gives no answer), and draws where the real face cannot.
          addFace(doc, name, weight, italic, localSources, undefined);
          void trackRegistration(localFaceInstalled(localSources).then((installed) => {
            if (!installed) addSlices([]);
          }), `local probe ${name}`);
        }
      }
      const slug = CJK_LATIN.names[key] ?? CJK_LATIN.names[normalize(name)];
      const latin = slug ? cjkLatinFacesOf(slug) : undefined;
      if (latin) {
        const declared = new Set<string>();
        for (const { faces } of families) {
          for (const [weight, italic, , , locals, overrides] of faces) {
            // A family Office ships without a bold face has its regular alone:
            // the browser emboldens it as Office does.
            const file = latin[weight >= 600 ? "700" : "400"] ?? latin["400"];
            const id = `${weight}/${italic}`;
            if (!file || declared.has(id)) continue;
            declared.add(id);
            const localSources = locals.map((local) => `local("${local.replace(/["\\]/g, "\\$&")}")`);
            addFace(doc, name, weight, italic, [...localSources, `url(${CJK_LATIN_PREFIX}${file}) format("woff2")`], CJK_LATIN.ranges?.[slug] ?? CJK_LATIN_RANGE, undefined, overrides ?? undefined);
          }
        }
      }
      // A Korean family's Hangul at its own full em (hangul-em.ts): the clone's
      // family, declared with this family's local() names and line overrides,
      // at the weights the Office family has (a bold of one without is
      // emboldened from the regular, as Word does).
      const hangul = hangulEmFamilyOf(key) ?? hangulEmFamilyOf(name);
      if (hangul) {
        const declared = new Set<number>();
        for (const { faces } of registry.families[name] ?? []) {
          for (const [weight, italic, , , locals, overrides] of faces) {
            if (italic || !hangul.weights.includes(weight) || declared.has(weight)) continue;
            declared.add(weight);
            const localSources = locals.map((local) => `local("${local.replace(/["\\]/g, "\\$&")}")`);
            for (const [file, range] of hangul.slices(weight)) {
              addFace(doc, hangul.alias, weight, 0, [...localSources, `url(${HANGUL_EM_PREFIX}${file}) format("woff2")`], range, undefined, overrides ?? undefined);
            }
          }
        }
      }
    })
    .catch(() => {
      names.delete(key);
      registrationGeneration += 1;
    })
    .finally(() => afterFacesDeclared(settle)), `registry ${key}`);
}

/**
 * A CSS font-family list with the substitutes for every Office name in it
 * placed right behind that name, so the stack keeps working where no alias
 * rule exists. Names already present are not repeated; the faces named are
 * registered in the document.
 */
export function withOfficeFontSubstitutes(stack: string): string {
  const parts = stack
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean);
  const seen = new Set(parts.map(normalize));
  const out: string[] = [];
  for (const part of parts) {
    out.push(part);
    ensureOfficeFontFaces(part);
    const powerPointFaces = powerPointSubstitutes(part);
    // A stand-in of a face further along the stack stands behind that face:
    // PowerPoint's Calibri for "San serifa" keeps the reader's own Calibri
    // ahead of Carlito.
    const behindLaterFace = new Set(
      powerPointFaces
        .filter((face) => seen.has(normalize(face)))
        .flatMap((face) => officeFontSubstitutes(face).map(normalize)),
    );
    for (const substitute of officeFontSubstitutes(part)) {
      if (seen.has(normalize(substitute)) || behindLaterFace.has(normalize(substitute))) continue;
      seen.add(normalize(substitute));
      // PowerPoint's own stand-in is itself an Office name: its faces too.
      if (powerPointFaces.includes(substitute)) ensureOfficeFontFaces(substitute);
      out.push(`"${substitute}"`);
    }
  }
  return out.join(", ");
}
