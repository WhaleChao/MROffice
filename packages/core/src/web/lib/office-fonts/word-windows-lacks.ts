/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import lacks from "./word-windows-lacks.json";




















type LacksTable = { faces: Record<string, string>; probed: Record<string, string>; aliases: { faces: Record<string, string>; exact?: Record<string, string> }; has: string[]; hasAdded?: string[] };
const table = lacks as LacksTable;
const lower = (entries: Record<string, string>) =>
  Object.entries(entries).map(([family, face]) => [family.toLowerCase(), face] as const);
/** The 182 families measured first (docx-look), the older reading's whole table. */
const MEASURED: ReadonlyMap<string, string> = new Map(lower(table.faces));
const FACES: ReadonlyMap<string, string> = new Map([...lower(table.probed), ...lower(table.faces)]);
/**
 * Word's own aliases for legacy and printer names and Windows'
 * FontSubstitutes, as Windows Word 16 draws them with no font-table entry
 * (docx-look/probes/cjk13 l1): Tms Rmn and CG Times are Times New Roman (the
 * registry's MS Serif for Tms Rmn is a bitmap face Word does not set: MS
 * Serif itself is Cambria), CG Omega Candara, Arabic Transparent Arial,
 * Tahoma Armenian and MS Shell Dlg 2 Tahoma, MS Shell Dlg Microsoft Sans
 * Serif. A name beginning with one is that face too ("CG Times Bold").
 */
const ALIASES: ReadonlyMap<string, string> = new Map(lower(table.aliases.faces));
/**
 * Word's aliases that hold for the whole name only, never as a prefix:
 * "System" is Calibri whatever the document's default face (docx-look/probes
 * cjk14 s1-s4), while "System Bold", Fixedsys and Terminal are Cambria.
 */
const EXACT_ALIASES: ReadonlyMap<string, string> = new Map(lower(table.aliases.exact ?? {}));

/** `window.__docsWordFontAliasesOff = true` asks the older reading (the
 * legacy names Word aliases set in Cambria), for an A/B of one build. */
function wordFontAliasesOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsWordFontAliasesOff?: boolean }).__docsWordFontAliasesOff === true;
}
const HAS_LISTED: ReadonlySet<string> = new Set(table.has.map((family) => family.toLowerCase()));





const HAS_ADDED: ReadonlyMap<string, string> = new Map((table.hasAdded ?? []).map((family) => [family.toLowerCase(), family] as const));
function wordHasAddedOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsWordHasAddedOff?: boolean }).__docsWordHasAddedOff === true;
}
const HAS = { has: (key: string): boolean => HAS_LISTED.has(key) || (HAS_ADDED.has(key) && !wordHasAddedOff()) };
/** The families Word has, as the table spells them, by their lower case. */
const HAS_LISTED_NAMES: ReadonlyMap<string, string> = new Map(table.has.map((family) => [family.toLowerCase(), family] as const));
const HAS_NAMES = { get: (key: string): string | undefined => HAS_LISTED_NAMES.get(key) ?? (wordHasAddedOff() ? undefined : HAS_ADDED.get(key)) };

const normalize = (family: string) => family.trim().replace(/^["']|["']$/g, "").trim().toLowerCase();

/** Whether Windows Word has the family (the table's `has`: Windows, Office
 * and Office cloud families, by their Latin names). */
export function wordWindowsHasFamily(family: string): boolean {
  return HAS.has(normalize(family));
}

/** `window.__docsWordLacksFixedOff = true` asks the older reading (the 182
 * measured families only, a font table's own pick before Word's Cambria), for
 * an A/B of one build. */
export function wordLacksFixedOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsWordLacksFixedOff?: boolean }).__docsWordLacksFixedOff === true;
}

const GENERIC = new Set(["serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui", "ui-monospace", "ui-serif", "ui-sans-serif", "math", "emoji"]);
/** A style word a name may carry after its family ("Lato-Regular", "Roboto Bold"). */
const STYLE_TAIL = /[\s-]+(?:regular|roman|book|normal|italic|oblique|bold|semibold|demibold|extrabold|ultrabold|black|heavy|light|extralight|ultralight|thin|medium)(?:[\s-]*(?:italic|oblique))?$/i;

/** `window.__docsWordFaceNamesOff = true` asks the older reading (a name
 * that is a CSS generic family drawn as that generic, a name with an accented
 * Latin letter not Word's to replace), for an A/B of one build. */
function wordFaceNamesOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsWordFaceNamesOff?: boolean }).__docsWordFaceNamesOff === true;
}

/** CSS's generic families, which a document may name as a face (WPS's
 * w:eastAsia="sans-serif", a web page's w:ascii="system-ui", corpus R403-D22,
 * R371-D27): faces Word does not have, so Windows Word 16 sets their Latin in
 * Cambria and their Han, kana and Hangul in SimSun and Batang
 * (docx-look/probes/cjk10 h3: system-ui, sans-serif, serif, monospace,
 * ui-monospace and cursive alike). */
const WORD_LACKED_GENERIC = new Set(["serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui", "ui-monospace", "ui-serif", "ui-sans-serif"]);

/** Whether Word's Cambria applies by default: a name in Latin letters, an
 * accented one among them ("HelveticaTürk", "Arial Kalın": Word's pick by
 * prefix, Arial, probes cjk10 g1 and f1), that names no family Word has, nor
 * one of them with a style word after it. Names in another script are East
 * Asian or localized faces, measured by script. */
function lacksByDefault(key: string): boolean {
  const letters = wordFaceNamesOff() ? /^[\x20-\x7e]+$/ : /^[\x20-\x7e\u00c0-\u024f\u1e00-\u1eff]+$/;
  if (!key || GENERIC.has(key) || HAS.has(key) || !letters.test(key) || !/[a-z]/.test(key)) return false;
  const stem = key.replace(STYLE_TAIL, "");
  return stem === key || !HAS.has(stem);
}

/** `window.__docsWordLacksPrefixOff = true` asks the older reading (a name
 * the table does not hold is Cambria), for an A/B of one build. */
function wordLacksPrefixOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsWordLacksPrefixOff?: boolean }).__docsWordLacksPrefixOff === true;
}

/** The names Word's own picks are keyed by, longest first; Arial, which Word
 * has, is one of them too ("Arialqzx" is set in Arial). */
const PREFIXES: readonly string[] = [...new Set([...FACES.keys(), ...ALIASES.keys(), "arial"])].sort((left, right) => right.length - left.length);

/**
 * Word picks a face for a name it lacks by the longest of its keys the name
 * begins with, a name its table does not hold included: "Liberation
 * SerifQzx", "Liberation Serif;Times New Roman" (LibreOffice's font list,
 * one name to Word; the page reads each semicolon as a space) and "Times;Qzx"
 * are Times New Roman, "Arialqzx", "Arial Narrow;Qzx" and "Liberation Sans
 * Narrow;Qzx" Arial, "Courier;Qzx" Courier New, "Segoe UI;Qzx" and "Myriad
 * Pro;Qzx" Segoe UI, "DejaVu Sans;Qzx" Verdana, while "Calibri;Qzx",
 * "Verdana;Qzx", "Tahoma;Qzx", "Liberation Mono;Qzx" and "DejaVu Serif;Qzx",
 * none of them beginning with a key, are Cambria (docx-look/probes/cjk2 t2,
 * t3: names Word had never met, Windows Word 16). Corpus R423-D06: a Normal
 * in "Liberation Serif;Times New Roman", which Word sets in Times New Roman.
 */
const prefixPicks = new Map<string, string | null>();
function lacksByPrefix(key: string): string | undefined {
  if (wordLacksPrefixOff()) return undefined;
  const aliasesOff = wordFontAliasesOff();
  const cacheKey = aliasesOff ? `\u0000${key}` : key;
  let pick = prefixPicks.get(cacheKey);
  if (pick === undefined) {
    const prefix = PREFIXES.find((candidate) => key.length > candidate.length && key.startsWith(candidate) &&
      (!aliasesOff || !ALIASES.has(candidate) || FACES.has(candidate)));
    pick = prefix === undefined ? null : FACES.get(prefix) ?? (aliasesOff ? undefined : ALIASES.get(prefix)) ?? "Arial";
    if (prefixPicks.size >= 2_000) prefixPicks.clear();
    prefixPicks.set(cacheKey, pick);
  }
  return pick ?? undefined;
}

/** `window.__docsWordFaceDashCutOff = true` asks the older reading (a name
 * with a "-" read whole: "Calibri-Light" a name the page could not draw),
 * for an A/B of one build. */
function wordFaceDashCutOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsWordFaceDashCutOff?: boolean }).__docsWordFaceDashCutOff === true;
}

/**
 * Word reads a name it does not have up to its first "-", as the reader
 * reads one up to its first "(" or "," (the API's `_word_face_name_cut`):
 * "Calibri-Light", "Calibri -x" and "Calibri - Body" are Calibri's regular,
 * "Lato-Regular" Lato, "Segoe UI-Semibold" Segoe UI, "Georgia-Bold" and
 * "Georgia-Ж" Georgia, "Arial Narrow-Bold" Arial Narrow, a head Word lacks
 * its pick for that head ("TimesNewRomanPS-BoldMT" Times New Roman), never a
 * full face ("Times New Roman-Bold" is the regular). Windows Word 16,
 * docx-look/probes/cjk10 g1 P45/P46, h1 P13, h2 P00-P04/P09/P10/P14/P20,
 * h3 P06/P08-P11 (names Word had never met). A name Word has whole
 * ("DFKai-SB", "SimSun-ExtB") is that face. The face Word draws for the
 * head, or undefined when there is no cut to make.
 */
function wordDashCutFace(family: string, key: string): string | undefined {
  if (wordFaceDashCutOff() || HAS.has(key)) return undefined;
  const dash = family.indexOf("-");
  const head = dash > 0 ? family.slice(0, dash).trim().replace(/^["']/, "").trim() : "";
  if (!head) return undefined;
  return HAS_NAMES.get(head.toLowerCase()) ?? wordWindowsLacksFace(head);
}

/** The face Windows Word draws `family`'s Latin text in, when Word lacks it. */
export function wordWindowsLacksFace(family: string | null | undefined): string | undefined {
  if (!family) return undefined;
  const key = normalize(family);
  if (wordLacksFixedOff()) return MEASURED.get(key);
  const known = FACES.get(key) ?? (wordFontAliasesOff() ? undefined : ALIASES.get(key) ?? EXACT_ALIASES.get(key));
  if (known !== undefined) return known;
  if (WORD_LACKED_GENERIC.has(key) && !wordFaceNamesOff()) return "Cambria";
  const cut = wordDashCutFace(family.trim(), key);
  if (cut !== undefined) return cut;
  if (!lacksByDefault(key)) return undefined;
  return lacksByPrefix(key) ?? "Cambria";
}
