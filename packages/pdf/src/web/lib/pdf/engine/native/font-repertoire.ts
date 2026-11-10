/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** The code points an installed font program always carries.
 *
 * A program narrowed to one edit's own glyphs would have to grow for the next
 * edit, and growing it means writing the whole program again -- so a document
 * would pay for the font once per edit. Narrowing to a fixed repertoire keeps
 * the installed program identical for every text inside it, which is what lets
 * a later edit reuse it without appending a single font byte. Text reaching
 * outside the repertoire installs a subset of its own glyphs instead
 * (font-extension-package.ts), so no character becomes uneditable; that font
 * is just not one a later edit can reuse. The whole canonical program is not
 * installed: its plan does not fit one engine answer.
 *
 * The repertoire is derived rather than listed: printable ASCII, plus the
 * symbol, fullwidth and level-1 hanzi rows of GB 2312, read back through the
 * platform's own GB18030 table.
 *
 * There are two such repertoires, each fixed, the smaller inside the larger.
 * Text in Latin, Greek or Cyrillic script -- one name fixed on an invoice --
 * installs the alphabetic one, a few hundred outlines; only text that needs a
 * hanzi pays for the GB 2312 rows. The larger one alone is four thousand
 * outlines, which would put a Latin edit's record past what a page may store.
 */

/** Blocks the alphabetic repertoire carries, beyond printable ASCII. */
const ALPHABETIC_BLOCKS: readonly (readonly [number, number])[] = [
  [0x00a0, 0x024f], // Latin-1 Supplement, Latin Extended-A and -B
  [0x02b0, 0x02ff], // spacing modifier letters (the accents typed alone)
  [0x0370, 0x03ff], // Greek
  [0x0400, 0x04ff], // Cyrillic
  [0x1e00, 0x1eff], // Latin Extended Additional (Vietnamese)
  [0x2000, 0x206f], // general punctuation
  [0x20a0, 0x20cf], // currency symbols
  [0x2100, 0x218f], // letterlike symbols and number forms
  [0x2190, 0x21ff], // arrows
];

export type PdfNativeFontRepertoireId = "alphabetic" | "cjk";

export interface PdfNativeFontRepertoireTier {
  readonly id: PdfNativeFontRepertoireId;
  readonly codePoints: ReadonlySet<number>;
}

const GB2312_ROWS: readonly (readonly [number, number])[] = [
  [0xa1a1, 0xa1fe], // symbols and punctuation
  [0xa3a1, 0xa3fe], // fullwidth alphanumerics
  [0xb0a1, 0xd7fe], // level-1 hanzi
];

let repertoire: ReadonlySet<number> | null | undefined;
let tiers: readonly PdfNativeFontRepertoireTier[] | undefined;

/** The largest repertoire: every code point an installed program can carry
 * without falling back to the whole canonical program. */
export function pdfNativeFontRepertoire(): ReadonlySet<number> | null {
  if (repertoire !== undefined) return repertoire;
  repertoire = buildRepertoire();
  return repertoire;
}

/** The fixed repertoires, smallest first, each containing the one before. */
export function pdfNativeFontRepertoireTiers(): readonly PdfNativeFontRepertoireTier[] {
  if (tiers !== undefined) return tiers;
  const alphabetic = alphabeticRepertoire();
  const cjk = pdfNativeFontRepertoire();
  tiers = Object.freeze([
    Object.freeze({ id: "alphabetic" as const, codePoints: alphabetic }),
    ...(cjk ? [Object.freeze({ id: "cjk" as const, codePoints: cjk })] : []),
  ]);
  return tiers;
}

function alphabeticRepertoire(): ReadonlySet<number> {
  const points = new Set<number>();
  for (let code = 0x20; code <= 0x7e; code += 1) points.add(code);
  for (const [first, last] of ALPHABETIC_BLOCKS) {
    for (let code = first; code <= last; code += 1) points.add(code);
  }
  return points;
}

function buildRepertoire(): ReadonlySet<number> | null {
  const points = new Set<number>(alphabeticRepertoire());
  let decoder: TextDecoder;
  try {
    decoder = new TextDecoder("gb18030", { fatal: false });
  } catch {
    // A runtime without the GB18030 table cannot derive the repertoire, and
    // guessing one would embed the wrong glyphs. Keep the whole program.
    return null;
  }
  for (const [first, last] of GB2312_ROWS) {
    for (let pair = first; pair <= last; pair += 1) {
      const lead = pair >> 8;
      const trail = pair & 0xff;
      if (trail < 0xa1 || trail > 0xfe) continue;
      const decoded = decoder.decode(Uint8Array.of(lead, trail));
      // An unassigned cell decodes to the replacement character; it carries no
      // glyph of its own and must not pull one in.
      if (decoded.length !== 1 || decoded === "�") continue;
      points.add(decoded.codePointAt(0)!);
    }
  }
  // A table that decoded almost nothing is a broken table, not a small one.
  return points.size > 3_000 + alphabeticRepertoire().size ? points : null;
}
