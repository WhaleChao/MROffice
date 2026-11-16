/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import lacks from "./word-windows-lacks.json";




















type LacksTable = { faces: Record<string, string>; probed: Record<string, string>; has: string[] };
const table = lacks as LacksTable;
const lower = (entries: Record<string, string>) =>
  Object.entries(entries).map(([family, face]) => [family.toLowerCase(), face] as const);
/** The 182 families measured first (docx-look), the older reading's whole table. */
const MEASURED: ReadonlyMap<string, string> = new Map(lower(table.faces));
const FACES: ReadonlyMap<string, string> = new Map([...lower(table.probed), ...lower(table.faces)]);
const HAS: ReadonlySet<string> = new Set(table.has.map((family) => family.toLowerCase()));

const normalize = (family: string) => family.trim().replace(/^["']|["']$/g, "").trim().toLowerCase();

/** `window.__docsWordLacksFixedOff = true` asks the older reading (the 182
 * measured families only, a font table's own pick before Word's Cambria), for
 * an A/B of one build. */
export function wordLacksFixedOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsWordLacksFixedOff?: boolean }).__docsWordLacksFixedOff === true;
}

const GENERIC = new Set(["serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui", "ui-monospace", "ui-serif", "ui-sans-serif", "math", "emoji"]);
/** A style word a name may carry after its family ("Lato-Regular", "Roboto Bold"). */
const STYLE_TAIL = /[\s-]+(?:regular|roman|book|normal|italic|oblique|bold|semibold|demibold|extrabold|ultrabold|black|heavy|light|extralight|ultralight|thin|medium)(?:[\s-]*(?:italic|oblique))?$/i;

/** Whether Word's Cambria applies by default: a name in Latin letters that
 * names no family Word has, nor one of them with a style word after it. Names
 * in another script are East Asian or localized faces, measured by script. */
function lacksByDefault(key: string): boolean {
  if (!key || GENERIC.has(key) || HAS.has(key) || !/^[\x20-\x7e]+$/.test(key) || !/[a-z]/.test(key)) return false;
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
const PREFIXES: readonly string[] = [...new Set([...FACES.keys(), "arial"])].sort((left, right) => right.length - left.length);

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
  let pick = prefixPicks.get(key);
  if (pick === undefined) {
    const prefix = PREFIXES.find((candidate) => key.length > candidate.length && key.startsWith(candidate));
    pick = prefix === undefined ? null : FACES.get(prefix) ?? "Arial";
    if (prefixPicks.size >= 2_000) prefixPicks.clear();
    prefixPicks.set(key, pick);
  }
  return pick ?? undefined;
}

/** The face Windows Word draws `family`'s Latin text in, when Word lacks it. */
export function wordWindowsLacksFace(family: string | null | undefined): string | undefined {
  if (!family) return undefined;
  const key = normalize(family);
  if (wordLacksFixedOff()) return MEASURED.get(key);
  const known = FACES.get(key);
  if (known !== undefined) return known;
  if (!lacksByDefault(key)) return undefined;
  return lacksByPrefix(key) ?? "Cambria";
}
