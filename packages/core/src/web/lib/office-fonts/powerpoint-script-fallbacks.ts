/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import table from "./powerpoint-script-fallbacks.json";
import { ensureOfficeRangedAliasFamily } from "./substitutes";
import bulletTable from "./powerpoint-bullet-fallbacks.json";
import { POWERPOINT_WINDOWS_FACES, STYLE_TAIL, windowsOfficeFace } from "./windows-office-faces";

type FallbackFace = "Calibri" | "Cambria";
const FACES: ReadonlyMap<string, FallbackFace> = new Map(
  Object.entries((table as { faces: Record<string, { face: FallbackFace }> }).faces).map(([family, entry]) => [family.toLowerCase(), entry.face] as const),
);

/** The scripts the probe measured: Latin and its extensions (Vietnamese among them), combining marks, Greek, Cyrillic. */
export const POWERPOINT_SCRIPT_FALLBACK_RANGES: readonly (readonly [number, number])[] = [
  [0x0000, 0x024f],
  [0x0300, 0x036f],
  [0x0370, 0x03ff],
  [0x0400, 0x052f],
  [0x1e00, 0x1eff],
  [0x1f00, 0x1fff],
];

/** Each replacing face's served cuts, Office's own metrics last (tried first where both draw a letter). */
const SERVED: Readonly<Record<FallbackFace, { alias: string; families: readonly string[] }>> = {
  Calibri: { alias: "MRAI PowerPoint Fallback Calibri", families: ["Carlito"] },
  Cambria: { alias: "MRAI PowerPoint Fallback Cambria", families: ["Caladea", "MRAI-Metric-Cambria"] },
};

/** The face PowerPoint for Windows draws `family`'s missing letters in, when measured. */
export function powerPointScriptFallbackFace(family: string): FallbackFace | undefined {
  return FACES.get(family.trim().replace(/^["']|["']$/g, "").trim().toLowerCase());
}

/**
 * `stack` with, behind each measured Office face's metric clone
 * ("MRAI-Metric-…", as `withOfficeFontSubstitutes` lists it), the family that
 * draws the face's missing Latin, Greek and Cyrillic letters as PowerPoint
 * for Windows does. A stack without such a face is returned as it is.
 */
export function withPowerPointScriptFallbacks(stack: string): string {
  const parts = stack.split(",").map((part) => part.trim()).filter(Boolean);
  const unquote = (part: string) => part.replace(/^["']|["']$/g, "").trim();
  const out: string[] = [];
  let pending: FallbackFace | undefined;
  let changed = false;
  for (const part of parts) {
    const name = unquote(part);
    const face = powerPointScriptFallbackFace(name);
    // Cambria's own missing letters: no other face of PowerPoint's to fall to.
    if (face) pending = face.toLowerCase() === name.toLowerCase() ? undefined : face;
    out.push(part);
    if (pending && name.startsWith("MRAI-Metric-") && !face) {
      const served = SERVED[pending];
      ensureOfficeRangedAliasFamily(served.alias, served.families, POWERPOINT_SCRIPT_FALLBACK_RANGES);
      if (!out.some((existing) => unquote(existing) === served.alias)) out.push(`"${served.alias}"`);
      pending = undefined;
      changed = true;
    }
  }
  return changed ? out.join(", ") : stack;
}




























export interface PowerPointBulletFace {
  /** The Office face the bullet is drawn in. */
  readonly face: string;
  /** Its advance in that face (em), where the table records it; otherwise measured. */
  readonly advanceEm?: number;
  /** The bullet's own face may have the character after all (an embedded face, read on the page): tried first. */
  readonly ownFirst?: true;
  /**
   * Where the reader lacks `face` (Cambria Math, Segoe UI Symbol): the served
   * bullet face (lib/office-fonts/bullet-faces, `MRAI-Bullet-…`) whose glyph
   * is set at Windows' own ink box and advance for the character (TRIAGE 99,
   * `drawn` in the table) -- STIX Two Math, Cambria Math's stand-in, draws
   * U+25C6 half as large again as Windows does, centred on the math axis.
   */
  readonly drawnBy?: { readonly family: string };
}

interface BulletFaceRow {
  readonly covers: readonly string[];
  readonly panoseSerifStyle: number;
  readonly unitsPerEm?: number;
  readonly advances?: Readonly<Record<string, readonly number[]>>;
}
interface BulletTable {
  readonly blocks: readonly string[];
  readonly chain: { readonly serif: readonly string[]; readonly sans: readonly string[] };
  readonly faces: Readonly<Record<string, BulletFaceRow>>;
  /** Chain face -> served face -> [character (hex), size-adjust %] (scripts/build-powerpoint-bullet-fallbacks.py --glyphbox). */
  readonly drawn?: Readonly<Record<string, Readonly<Record<string, readonly string[]>>>>;
}

const BULLETS = bulletTable as unknown as BulletTable;
const span = (range: string): readonly [number, number] => {
  const [first, last] = range.split("-");
  return [parseInt(first, 16), parseInt(last ?? first, 16)];
};
const BULLET_BLOCKS = BULLETS.blocks.map(span);
const BULLET_FACES: ReadonlyMap<string, { readonly name: string; readonly row: BulletFaceRow; readonly covers: readonly (readonly [number, number])[] }> = new Map(
  Object.entries(BULLETS.faces).map(([name, row]) => [name.toLowerCase(), { name, row, covers: row.covers.map(span) }] as const),
);
/** Faces named for a symbol set: their characters are pictures, never fallen back from. */
const SYMBOL_SET_FACE = /^(wingdings|webdings|symbol$|marlett|ms outlook|bookshelf symbol|mt extra)/i;

/** Chain face (lower case) -> character -> the bullet face drawing it at Windows' ink box. */
const BULLET_DRAWN: ReadonlyMap<string, ReadonlyMap<number, { readonly family: string }>> = new Map(
  Object.entries(BULLETS.drawn ?? {}).map(([face, byFamily]) => [
    face.toLowerCase(),
    new Map(Object.entries(byFamily).flatMap(([family, codes]) => codes.map((hex) => [parseInt(hex, 16), { family }] as const))),
  ]),
);

const within = (ranges: readonly (readonly [number, number])[], code: number) => ranges.some(([first, last]) => first <= code && code <= last);

/** The recorded advance (em) of `code` in a chain face, if any. */
function bulletAdvanceEm(row: BulletFaceRow, code: number): number | undefined {
  if (!row.advances || !row.unitsPerEm) return undefined;
  for (const [start, widths] of Object.entries(row.advances)) {
    const index = code - parseInt(start, 16);
    if (index >= 0 && index < widths.length) return widths[index] / row.unitsPerEm;
  }
  return undefined;
}

/** Whether a recorded face has `code`; undefined outside the blocks the table records for it. */
function bulletFaceHas(face: string, code: number): boolean | undefined {
  const entry = BULLET_FACES.get(face.toLowerCase());
  if (!entry) return undefined;
  if (within(entry.covers, code)) return true;
  // Outside the shape blocks a face's own characters are not recorded (general punctuation: every face's),
  // except for the faces whose advances are (the chain faces, Calibri).
  return entry.row.advances ? false : within(BULLET_BLOCKS, code) ? false : undefined;
}






function bulletFaceEntry(face: string): (NonNullable<ReturnType<typeof BULLET_FACES.get>> & { readonly styleOf?: true }) | undefined {
  let name = face.toLowerCase();
  const own = BULLET_FACES.get(name);
  if (own) return own;
  for (let guard = 0; guard < 3; guard += 1) {
    const shorter = name.replace(STYLE_TAIL, "").trim();
    if (!shorter || shorter === name) return undefined;
    name = shorter;
    const family = BULLET_FACES.get(name);
    if (family) return { ...family, styleOf: true };
  }
  return undefined;
}

/** The first chain face behind a serif or sans face that has `code`. */
function bulletChainFace(serif: boolean, code: number): PowerPointBulletFace | undefined {
  for (const face of serif ? BULLETS.chain.serif : BULLETS.chain.sans) {
    const entry = BULLET_FACES.get(face.toLowerCase());
    if (!entry || !within(entry.covers, code)) continue;
    const advanceEm = bulletAdvanceEm(entry.row, code);
    const drawnBy = BULLET_DRAWN.get(entry.name.toLowerCase())?.get(code);
    return { face: entry.name, ...(advanceEm === undefined ? {} : { advanceEm }), ...(drawnBy ? { drawnBy } : {}) };
  }
  return undefined;
}










const bulletChainIsSans = (style: number) => style >= 11 && style <= 15;
















export function powerPointBulletFace(
  typeface: string,
  char: string,
  options: { readonly embedded?: boolean; readonly panoseSerifStyle?: number } = {},
): PowerPointBulletFace | undefined {
  const name = typeface.trim().replace(/^["']|["']$/g, "").trim();
  const chars = Array.from(char);
  // A name with a comma or parenthesis ("Arial,Sans-Serif") is cut by PowerPoint's own reading of it, not here;
  // a dotted name (".LucidaGrandeUI") is a Mac system face's.
  if (!name || chars.length !== 1 || SYMBOL_SET_FACE.test(name) || /[,(]/.test(name) || name.startsWith(".")) return undefined;
  const code = chars[0].codePointAt(0) as number;
  if (options.embedded && !bulletFaceEntry(name)) {
    if (!within(BULLET_BLOCKS, code)) return undefined;
    const fallback = bulletChainFace(!bulletChainIsSans(options.panoseSerifStyle ?? 0), code);
    return fallback ? { ...fallback, ownFirst: true } : undefined;
  }
  // The face PowerPoint has in its place: Office's own pick, else Calibri.
  const drawn = options.embedded ? undefined : windowsOfficeFace(name, POWERPOINT_WINDOWS_FACES)?.face;
  const entry = bulletFaceEntry(drawn ?? name);
  if (!entry) return undefined;
  const has = bulletFaceHas(entry.name, code);
  if (has !== false) {
    if (entry.name.toLowerCase() === name.toLowerCase() || (drawn === undefined && entry.styleOf)) return undefined;
    // The face drawn in its place has it (Calibri for a face Windows lacks).
    const advanceEm = has ? bulletAdvanceEm(entry.row, code) : undefined;
    return advanceEm === undefined ? { face: entry.name } : { face: entry.name, advanceEm };
  }
  return bulletChainFace(!bulletChainIsSans(entry.row.panoseSerifStyle), code);
}
