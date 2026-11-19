/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Hangul at the Korean Office fonts' own advances.
 *
 * Malgun Gothic, Gulim, Dotum, Batang and Gungsuh (and their -Che forms) set
 * every Hangul syllable and compatibility jamo a full em; the served stand-ins
 * (Noto Sans KR, Noto Serif KR) set them 0.92 em, so a Korean line held 8%
 * more text than Word's (Windows Word, corpus R284-D30: glyph origins 12.00 pt
 * apart at 12 pt; "동일/한" broken "동일한/"). The clones
 * (scripts/build-hangul-em-fonts.py, public/fonts/hangul-em) are the
 * stand-in's glyphs moved to the recorded advance, cut into the stand-in's own
 * unicode-range slices.
 *
 * They are a family of their own per Office family ("MRAI-Metric-Hangul-
 * MalgunGothic", the Office face first through local()), which a stack names
 * in front of that family (`withHangulEmFaces`): an editor opts in where it
 * draws Word's lines (Docs). The spreadsheet's grid keeps the stand-in's
 * widths, and nothing on the server changes. The faces are declared with the
 * Office family's (substitutes.ts, `ensureOfficeFontFaces`), with its line
 * overrides, so a Hangul glyph stands on the same line box as before.
 *
 * `window.__officeHangulEmOff = true` leaves every stack as it was.
 */
import { isExcelCutFamily } from "./excel-cut-families";
import hangulEmFaces from "./hangul-em-faces.json";

interface HangulEmManifest {
  families: Record<string, { clone: string; names: string[]; weights: number[] }>;
  clones: Record<string, { family: string; source: string }>;
  slices: Record<string, Record<string, Array<[string, string]>>>;
}

const MANIFEST = hangulEmFaces as unknown as HangulEmManifest;
export const HANGUL_EM_PREFIX = "/fonts/hangul-em/";
const ALIAS_PREFIX = "MRAI-Metric-Hangul-";
const ALIAS_KEY_PREFIX = ALIAS_PREFIX.toLowerCase();

export interface HangulEmFamily {
  /** The family a stack names in front of the Office family. */
  readonly alias: string;
  /** The weights the Office family has (a bold of a family without one is
   * emboldened from the regular, as Word does). */
  readonly weights: readonly number[];
  /** The slices of the clone drawn at `weight` (file, unicode-range). */
  slices(weight: number): ReadonlyArray<readonly [string, string]>;
}

const unquote = (name: string) => name.trim().replace(/^["']|["']$/g, "").trim();

const byName = new Map<string, HangulEmFamily>();
for (const [family, entry] of Object.entries(MANIFEST.families)) {
  const slices = MANIFEST.slices[entry.clone] ?? {};
  const found: HangulEmFamily = {
    alias: `${ALIAS_PREFIX}${family.replace(/\s+/g, "")}`,
    weights: entry.weights,
    slices: (weight) => slices[String(weight >= 600 && slices["700"] ? 700 : 400)] ?? [],
  };
  for (const name of [family, ...entry.names]) byName.set(name.toLowerCase(), found);
}

/** The Hangul clone of a Korean Office family (any of its names), or null. */
export function hangulEmFamilyOf(family: string): HangulEmFamily | null {
  return byName.get(unquote(family).toLowerCase()) ?? null;
}

/** Whether a stack token is one of the clones' families. */
export function isHangulEmAlias(family: string): boolean {
  return unquote(family).toLowerCase().startsWith(ALIAS_KEY_PREFIX);
}

/** A stack's first family that is not a Hangul clone, nor a sheet's cut
 * family (excel-cut-families.ts), unquoted: the family the run asked for.
 * What is kept by face (row heights, printed lines, digit widths) is read
 * under it. */
export function stackPrimaryFamily(stack: string): string {
  const families = stack.split(",");
  return unquote(families.find((family) => !isHangulEmAlias(family) && !isExcelCutFamily(family)) ?? families[0] ?? "");
}

function hangulEmOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __officeHangulEmOff?: boolean }).__officeHangulEmOff === true;
}

/**
 * The stack with each Korean Office family's Hangul clone named right in
 * front of it (once). Other characters fall through the clone (its slices
 * cover Hangul only) to the family as before.
 */
export function withHangulEmFaces(stack: string, ensure?: (family: string) => void): string {
  if (hangulEmOff() || !stack) return stack;
  const tokens = stack.split(",").map((token) => token.trim()).filter(Boolean);
  const seen = new Set(tokens.map((token) => unquote(token).toLowerCase()));
  const out: string[] = [];
  let changed = false;
  for (const token of tokens) {
    const clone = isHangulEmAlias(token) ? null : hangulEmFamilyOf(token);
    if (clone && !seen.has(clone.alias.toLowerCase())) {
      seen.add(clone.alias.toLowerCase());
      ensure?.(unquote(token));
      out.push(`"${clone.alias}"`);
      changed = true;
    }
    out.push(token);
  }
  return changed ? out.join(", ") : stack;
}
