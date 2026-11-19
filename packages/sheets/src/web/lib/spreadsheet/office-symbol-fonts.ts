/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A legacy symbol font's codes as the Unicode characters Windows draws them
 * as (office-symbol-fonts.json; the API's office_symbol_fonts.py reads an
 * identical copy): a Wingdings cell's "ò ñ ó" are Windows Excel's ⇩ ⇧ ⬄
 * (lookwin R139-X01). Text whose every inked code the face's packaged
 * substitute draws itself (Wine Wingdings, Wine Symbol) keeps its glyphs.
 */
import TABLE from "./office-symbol-fonts.json";

interface SymbolTable {
  readonly faces: Readonly<Record<string, Readonly<Record<string, string>>>>;
  readonly substituteCodes: Readonly<Record<string, string>>;
}

const table = TABLE as unknown as SymbolTable;

function keyOf(family: string): string {
  return family.normalize("NFKC").toLowerCase().split(/\s+/).filter(Boolean).join(" ");
}

function codeOf(character: string): number {
  const code = character.codePointAt(0) ?? 0;
  return code >= 0xf020 && code <= 0xf0ff ? code - 0xf000 : code;
}

function hex(code: number): string {
  return code.toString(16).toUpperCase().padStart(2, "0");
}

/** `text` set in symbol font `family` as the characters its codes stand for,
 * or undefined: a family the table does not map, or text the substitute
 * draws itself (the API's office_symbol_text). */
export function officeSymbolText(family: string | null | undefined, text: string): string | undefined {
  if (!family) return undefined;
  const key = keyOf(family);
  const codes = table.faces[key];
  if (!codes) return undefined;
  const covered = new Set((table.substituteCodes[key] ?? "").split(" ").filter(Boolean));
  const inked = [...text].filter((character) => !/\s/u.test(character));
  if (!inked.length || inked.every((character) => covered.has(hex(codeOf(character))))) return undefined;
  return [...text].map((character) => {
    const code = codeOf(character);
    return code >= 0x20 && code <= 0xff ? codes[hex(code)] ?? character : character;
  }).join("");
}

/**
 * A text box's bullet as it is drawn: `a:buChar`'s character set in
 * `a:buFont`'s face. In a legacy symbol face the character is the symbol its
 * code stands for (Symbol's U+F0B7 the round bullet, Wingdings' U+F0A7 the
 * small square), and the page draws that symbol in the body's own font: no
 * substitute stands under the marker, so every code is mapped, and a symbol
 * code of a face without a table (Wingdings 2, Webdings) is a plain bullet
 * rather than the private-use box (lookwin R452-X59: the Notes box's twenty
 * Symbol bullets printed as boxes on taller lines). Undefined when the
 * character is drawn as written.
 */
export function officeSymbolMarker(family: string | null | undefined, text: string): string | undefined {
  if (!family || !text) return undefined;
  const codes = table.faces[keyOf(family)];
  const privateUse = (character: string) => {
    const code = character.codePointAt(0) ?? 0;
    return code >= 0xf020 && code <= 0xf0ff;
  };
  if (!codes) return [...text].some(privateUse) ? [...text].map((character) => privateUse(character) ? "\u2022" : character).join("") : undefined;
  const mapped = [...text].map((character) => {
    const code = codeOf(character);
    return code >= 0x20 && code <= 0xff ? codes[hex(code)] ?? (privateUse(character) ? "\u2022" : character) : character;
  }).join("");
  return mapped === text ? undefined : mapped;
}
