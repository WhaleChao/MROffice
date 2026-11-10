/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Office "Insert → Symbol" model, shared by every MRAI editor.
 *
 * This is the behavioural specification Word's Symbol dialog implements, read
 * off Microsoft's own documentation and ECMA-376 / ISO-IEC 29500-1 §17.3.3.30
 * (`w:sym`), and reproduced here as PURE DATA + PURE FUNCTIONS so it can be
 * unit-tested and reused by Docs / Presentation / Spreadsheet without dragging
 * a single editor dependency along.
 *
 * WHAT WORD DOES (the contract this module encodes)
 * -------------------------------------------------
 * 1. The dialog has two tabs: **Symbols** and **Special Characters**.
 * 2. The Symbols tab has a **font dropdown** whose first entry is
 *    `(normal text)`. `(normal text)` means "no symbol font": the picked
 *    character is inserted as an ORDINARY character in the run's own font, and
 *    is written to the file as ordinary `<w:t>` text.
 * 3. Choosing a real font instead makes the character a *symbol character*.
 *    ECMA-376 is explicit that such a character "does not use any of the run
 *    fonts specified in the rFonts element … this character shall be determined
 *    by pulling the character with the hexadecimal value specified in the char
 *    attribute from the font specified in the font attribute", and is written as
 *    `<w:sym w:font="…" w:char="…"/>` INSTEAD of `<w:t>` — inside a run that may
 *    also hold ordinary text.
 * 4. For the legacy non-Unicode "symbol fonts" (Symbol, Wingdings, Wingdings 2,
 *    Wingdings 3, Webdings) Word stores the character SHIFTED INTO THE UNICODE
 *    PRIVATE USE AREA by adding 0xF000, so Wingdings `0x3A` is written as
 *    `w:char="F03A"` (the ECMA-376 example itself). The offset exists so a
 *    machine without the font cannot silently font-match the codepoint to an
 *    unrelated Latin glyph. Both spellings are legal on READ — a producer may
 *    write the bare in-font value — so the importer accepts either and the
 *    exporter re-emits whatever it read, byte for byte.
 * 5. The grid is ordered by code point and scrolls; the readout under it names
 *    the character twice: its code IN THE CURRENT FONT and its Unicode code
 *    point. Under `(normal text)` those two are the same number.
 * 6. Word remembers **recently used symbols**, and a recent entry remembers its
 *    FONT as well as its character (a recent Wingdings glyph re-inserts as
 *    Wingdings, not as whatever font the dialog currently shows).
 *
 * WHAT THIS MODULE DELIBERATELY DOES NOT DO
 * -----------------------------------------
 * No font bytes, no glyph artwork, no font metrics: the repertoires below are
 * Unicode code-point RANGES, and the available font names come from MRAI's own
 * font catalogue (`officeTypographyCatalog`) through the caller. Nothing here is
 * derived from a proprietary font or from Word's binaries.
 */

// ---------------------------------------------------------------------------
// w:sym character-code encoding
// ---------------------------------------------------------------------------

/** Start of the Unicode Private Use Area block Word shifts symbol fonts into. */
export const SYMBOL_PUA_OFFSET = 0xf000;

/**
 * The legacy non-Unicode symbol fonts. A character picked from one of these is
 * stored PUA-shifted (`0xF000 + inFontCode`) exactly as Word writes it; every
 * other font stores its real Unicode code point.
 *
 * Lowercased for case-insensitive comparison — a DOCX in the wild spells these
 * "WINGDINGS" / "wingdings" as often as "Wingdings".
 */
const SYMBOL_FONT_NAMES: ReadonlySet<string> = new Set([
  "symbol",
  "wingdings",
  "wingdings 2",
  "wingdings 3",
  "webdings",
  "monotype sorts",
  "zapfdingbats",
  "zapf dingbats",
]);

/** Whether `font` is a legacy symbol font whose characters live in the PUA. */
export function isLegacySymbolFont(font: string): boolean {
  return SYMBOL_FONT_NAMES.has(font.trim().toLowerCase());
}

/** The longest font name `w:font` may carry (ECMA-376 ST_String, clamped). */
export const MAX_SYMBOL_FONT_NAME_LENGTH = 64;

/**
 * Normalize a `w:font` value: a trimmed, single-line, length-clamped font name,
 * or null when the value cannot be a font name at all. Fail-closed — a run
 * whose font attribute is junk degrades to an ordinary character rather than
 * carrying attacker-controlled text into `w:font` on the way back out.
 */
export function normalizeSymbolFont(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed || trimmed.length > MAX_SYMBOL_FONT_NAME_LENGTH) return null;
  // No control characters, no line breaks: a font name is one plain line.
  if (/[\u0000-\u001f\u007f]/.test(trimmed)) return null;
  return trimmed;
}

/**
 * Normalize a `w:char` value to the canonical 4-hex-digit UPPERCASE spelling
 * Word writes ("F0FC", "00A7"), or null when it is not a hex character code.
 *
 * ECMA-376 types `w:char` as ST_ShortHexNumber: exactly two bytes, i.e. four
 * hex digits. Producers do write fewer digits ("3a"), so a shorter value is
 * accepted and zero-padded; anything longer, non-hex, or outside the BMP is
 * rejected. PURE.
 */
export function normalizeSymbolCharCode(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!/^[0-9A-Fa-f]{1,4}$/.test(trimmed)) return null;
  return trimmed.toUpperCase().padStart(4, "0");
}

/** The canonical `w:char` spelling of a code point. PURE. */
export function symbolCharCodeFromCodePoint(codePoint: number): string | null {
  if (!Number.isInteger(codePoint) || codePoint < 0 || codePoint > 0xffff) {
    return null;
  }
  return codePoint.toString(16).toUpperCase().padStart(4, "0");
}

/**
 * The code point a `w:sym` renders as: the `w:char` value verbatim.
 *
 * A PUA-shifted value is NOT un-shifted here. U+F020..U+F0FF is exactly where
 * the legacy symbol fonts map their glyphs in their own `cmap`, so the shifted
 * code point is the one that actually renders the intended glyph when the font
 * is present — and keeping it is what makes re-export byte-identical. Callers
 * that want the in-font code for display use {@link symbolInFontCharCode}.
 */
export function symbolCodePointFromCharCode(charCode: string): number | null {
  const normalized = normalizeSymbolCharCode(charCode);
  return normalized === null ? null : Number.parseInt(normalized, 16);
}

/**
 * The character code Word's readout shows for a symbol: the code IN THE FONT,
 * i.e. the PUA offset removed when the value sits in U+F000..U+F0FF and the font
 * is a legacy symbol font. Word's own dialog shows `252` for the Wingdings
 * check mark stored as `F0FC`.
 */
export function symbolInFontCharCode(font: string, charCode: string): string | null {
  const codePoint = symbolCodePointFromCharCode(charCode);
  if (codePoint === null) return null;
  if (
    isLegacySymbolFont(font) &&
    codePoint >= SYMBOL_PUA_OFFSET &&
    codePoint <= SYMBOL_PUA_OFFSET + 0xff
  ) {
    return symbolCharCodeFromCodePoint(codePoint - SYMBOL_PUA_OFFSET);
  }
  return symbolCharCodeFromCodePoint(codePoint);
}

/**
 * The text a `w:sym` contributes to the document: the single character at its
 * code point, or "" when the char attribute is unusable.
 *
 * Note this is one UTF-16 unit by construction (`w:char` is capped at 0xFFFF),
 * so a symbol character is always exactly one character of document text — which
 * is what lets a symbol mark sit on a one-character text node.
 */
export function symbolCharText(charCode: string): string {
  const codePoint = symbolCodePointFromCharCode(charCode);
  return codePoint === null ? "" : String.fromCodePoint(codePoint);
}

/**
 * The `w:char` value to write for a character picked from `font`: the code point
 * PUA-shifted when the font is a legacy symbol font and the code point is a
 * bare in-font byte, otherwise the code point itself.
 *
 * Idempotent on an already-shifted value, so picking U+F0FC from Wingdings
 * writes `F0FC` rather than double-shifting to `E0FC`.
 */
export function symbolCharCodeForFont(
  font: string,
  codePoint: number,
): string | null {
  if (!Number.isInteger(codePoint) || codePoint < 0 || codePoint > 0xffff) {
    return null;
  }
  if (isLegacySymbolFont(font) && codePoint <= 0xff) {
    return symbolCharCodeFromCodePoint(SYMBOL_PUA_OFFSET + codePoint);
  }
  return symbolCharCodeFromCodePoint(codePoint);
}

// ---------------------------------------------------------------------------
// The character grid repertoires
// ---------------------------------------------------------------------------

/** An inclusive code-point range in the grid, in code-point order. */
export interface SymbolCodeRange {
  readonly first: number;
  readonly last: number;
}

/**
 * The `(normal text)` repertoire, in code-point order — the blocks Word's own
 * `(normal text)` grid walks. Deliberately excludes every range that holds only
 * unassigned, combining, format, or whitespace-only code points, so no cell in
 * the grid is a blank the reader cannot tell from a missing glyph.
 */
export const NORMAL_TEXT_SYMBOL_RANGES: readonly SymbolCodeRange[] = [
  { first: 0x0021, last: 0x007e }, // Basic Latin, printable (Word starts at "!")
  { first: 0x00a1, last: 0x00ff }, // Latin-1 Supplement
  { first: 0x0100, last: 0x017f }, // Latin Extended-A
  { first: 0x0180, last: 0x024f }, // Latin Extended-B
  { first: 0x02b0, last: 0x02ff }, // Spacing Modifier Letters
  { first: 0x0374, last: 0x03ff }, // Greek and Coptic (assigned span)
  { first: 0x0400, last: 0x04ff }, // Cyrillic
  { first: 0x2010, last: 0x2027 }, // General Punctuation, printable span
  { first: 0x2030, last: 0x205e }, // General Punctuation, printable span
  { first: 0x2070, last: 0x209c }, // Superscripts and Subscripts
  { first: 0x20a0, last: 0x20bf }, // Currency Symbols
  { first: 0x2100, last: 0x214f }, // Letterlike Symbols
  { first: 0x2150, last: 0x218b }, // Number Forms
  { first: 0x2190, last: 0x21ff }, // Arrows
  { first: 0x2200, last: 0x22ff }, // Mathematical Operators
  { first: 0x2300, last: 0x23ff }, // Miscellaneous Technical
  { first: 0x2460, last: 0x24ff }, // Enclosed Alphanumerics
  { first: 0x2500, last: 0x257f }, // Box Drawing
  { first: 0x2580, last: 0x259f }, // Block Elements
  { first: 0x25a0, last: 0x25ff }, // Geometric Shapes
  { first: 0x2600, last: 0x26ff }, // Miscellaneous Symbols
  { first: 0x2700, last: 0x27bf }, // Dingbats
  { first: 0x3001, last: 0x303f }, // CJK Symbols and Punctuation
  { first: 0x3041, last: 0x309f }, // Hiragana
  { first: 0x30a0, last: 0x30ff }, // Katakana
  { first: 0xff01, last: 0xff5e }, // Fullwidth forms
  { first: 0xffe0, last: 0xffe6 }, // Fullwidth currency signs
];

/**
 * A legacy symbol font's repertoire: the PUA window U+F020..U+F0FF those fonts
 * map their glyphs into (0xF000 + the printable single-byte range). Word's grid
 * for Wingdings / Symbol / Webdings shows exactly this window.
 */
export const LEGACY_SYMBOL_FONT_RANGES: readonly SymbolCodeRange[] = [
  { first: 0xf020, last: 0xf0ff },
];

/** The grid's column count — Word's Symbol grid is 31 cells wide. */
export const SYMBOL_GRID_COLUMNS = 31;

/** Hard ceiling on grid cells, so a pathological range list cannot hang a tab. */
const MAX_GRID_CELLS = 8192;

/**
 * Expand ranges into the grid's code points, in order, de-duplicated and
 * clamped to {@link MAX_GRID_CELLS}. PURE.
 */
export function expandSymbolRanges(
  ranges: readonly SymbolCodeRange[],
): number[] {
  const out: number[] = [];
  const seen = new Set<number>();
  for (const range of ranges) {
    if (
      !Number.isInteger(range.first) ||
      !Number.isInteger(range.last) ||
      range.first < 0 ||
      range.last > 0xffff ||
      range.last < range.first
    ) {
      continue;
    }
    for (let code = range.first; code <= range.last; code += 1) {
      if (seen.has(code)) continue;
      seen.add(code);
      out.push(code);
      if (out.length >= MAX_GRID_CELLS) return out;
    }
  }
  return out;
}

/**
 * The grid code points for a font selection: the legacy symbol-font PUA window
 * when the font is one of those fonts, otherwise the `(normal text)` repertoire.
 * `null` font = `(normal text)`.
 */
export function symbolGridCodePoints(font: string | null): number[] {
  return expandSymbolRanges(
    font !== null && isLegacySymbolFont(font)
      ? LEGACY_SYMBOL_FONT_RANGES
      : NORMAL_TEXT_SYMBOL_RANGES,
  );
}

// ---------------------------------------------------------------------------
// The Special Characters tab
// ---------------------------------------------------------------------------

/**
 * How a special character reaches the file. Word writes three of them as
 * dedicated run children rather than as literal text, because they are
 * line-breaking instructions rather than glyphs:
 *   - `noBreakHyphen` → `<w:noBreakHyphen/>` (a hyphen that never breaks)
 *   - `softHyphen`    → `<w:softHyphen/>`    (a break opportunity, usually invisible)
 * Everything else is an ordinary character in `<w:t>`.
 */
export type SpecialCharacterKind = "text" | "noBreakHyphen" | "softHyphen";

export interface OfficeSpecialCharacter {
  /** Stable id — React key, test id, and the i18n label key suffix. */
  readonly id: string;
  /** The character inserted into the document. Always exactly one character. */
  readonly char: string;
  /**
   * A legible stand-in shown in the "Character" column when {@link char} is
   * itself invisible (the spaces, the optional hyphen). Word shows the real
   * character; a blank cell in a web table reads as a rendering bug, so the
   * spaces get their open-box/space-marker faces.
   */
  readonly display?: string;
  /** How the character is written to DOCX. Defaults to "text". */
  readonly kind?: SpecialCharacterKind;
  /**
   * The keyboard shortcut MRAI ACTUALLY binds for this character, in the
   * platform-neutral ProseMirror keymap spelling ("Mod" = Cmd on macOS, Ctrl
   * elsewhere), or omitted when MRAI binds none.
   *
   * Word's own dialog lists OS-level dead-key sequences (Option+G for ©) in
   * this column. MRAI does not: those are the operating system's text input, not
   * this application's bindings, and claiming them would be a promise the editor
   * does not keep on a different keyboard layout. A blank cell here is the
   * honest answer, and the character is always one click away in the dialog.
   */
  readonly shortcut?: string;
}

/**
 * Word's Special Characters tab, in Word's own row order.
 *
 * The shortcut column is MRAI's real bindings only — see
 * {@link OfficeSpecialCharacter.shortcut}. The four that carry one are the four
 * Word treats as application commands rather than as OS text input, minus the
 * en dash: Word's en-dash chord is Ctrl/Cmd+Minus, which every browser owns for
 * page zoom, so binding it would be a shortcut that silently does not fire.
 */
export const OFFICE_SPECIAL_CHARACTERS: readonly OfficeSpecialCharacter[] = [
  { id: "emDash", char: "—", shortcut: "Mod-Alt-Minus" },
  { id: "enDash", char: "–" },
  {
    id: "nonBreakingHyphen",
    char: "\u2011",
    kind: "noBreakHyphen",
    shortcut: "Mod-Alt-Shift-Minus",
  },
  {
    id: "optionalHyphen",
    // U+00AD SOFT HYPHEN: invisible unless the line actually breaks there.
    char: "\u00ad",
    display: "‧",
    kind: "softHyphen",
    shortcut: "Mod-Alt-Period",
  },
  { id: "emSpace", char: "\u2003", display: "⌷" },
  { id: "enSpace", char: "\u2002", display: "⌷" },
  { id: "quarterEmSpace", char: "\u2005", display: "⌷" },
  {
    id: "nonBreakingSpace",
    char: "\u00a0",
    display: "␣",
    shortcut: "Ctrl-Shift-Space",
  },
  { id: "copyright", char: "©" },
  { id: "registered", char: "®" },
  { id: "trademark", char: "™" },
  { id: "section", char: "§" },
  { id: "paragraph", char: "¶" },
  { id: "ellipsis", char: "…" },
  { id: "singleOpeningQuote", char: "‘" },
  { id: "singleClosingQuote", char: "’" },
  { id: "doubleOpeningQuote", char: "“" },
  { id: "doubleClosingQuote", char: "”" },
];

/** Look up a special character by id. PURE; null for an unknown id. */
export function specialCharacterById(
  id: string,
): OfficeSpecialCharacter | null {
  return OFFICE_SPECIAL_CHARACTERS.find((entry) => entry.id === id) ?? null;
}

/**
 * The shortcut MRAI binds for a character, or null. Drives the dialog's
 * "Keyboard shortcut:" readout on the Symbols tab, which in Word shows the
 * binding of whichever character the grid has selected.
 */
export function symbolShortcutForChar(char: string): string | null {
  const entry = OFFICE_SPECIAL_CHARACTERS.find((item) => item.char === char);
  return entry?.shortcut ?? null;
}

/**
 * Render a keymap spelling as the label a reader recognises on THIS platform:
 * "Mod-Alt-Minus" → "⌘⌥-" on macOS, "Ctrl+Alt+-" elsewhere. PURE — the caller
 * passes the platform so this stays testable and SSR-safe.
 */
export function formatSymbolShortcut(
  shortcut: string,
  platform: "mac" | "other",
): string {
  const parts = shortcut.split("-");
  const keyToken = parts[parts.length - 1] ?? "";
  const modifiers = parts.slice(0, -1);
  const key =
    keyToken === "Minus"
      ? "-"
      : keyToken === "Period"
        ? "."
        : keyToken === "Space"
          ? "␣"
          : keyToken;
  if (platform === "mac") {
    const glyphs: Record<string, string> = {
      Mod: "⌘",
      Cmd: "⌘",
      Ctrl: "⌃",
      Alt: "⌥",
      Shift: "⇧",
    };
    return `${modifiers.map((mod) => glyphs[mod] ?? mod).join("")}${key}`;
  }
  const names: Record<string, string> = {
    Mod: "Ctrl",
    Cmd: "Ctrl",
    Ctrl: "Ctrl",
    Alt: "Alt",
    Shift: "Shift",
  };
  return [...modifiers.map((mod) => names[mod] ?? mod), key].join("+");
}

// ---------------------------------------------------------------------------
// Recently used symbols
// ---------------------------------------------------------------------------

/** Word's "Recently used symbols" row holds 16 entries. */
export const OFFICE_SYMBOL_RECENT_LIMIT = 16;

/**
 * One remembered symbol: its character AND the font it was picked from, because
 * a Wingdings glyph re-inserted from the recent row must come back as Wingdings.
 * `font: null` = `(normal text)`.
 */
export interface OfficeRecentSymbol {
  readonly char: string;
  readonly font: string | null;
}

/** The storage key for the recent row. Shared across editors — a user habit. */
export const OFFICE_SYMBOL_RECENT_STORAGE_KEY = "mrai.office.recentSymbols.v2";

function sameRecentSymbol(
  left: OfficeRecentSymbol,
  right: OfficeRecentSymbol,
): boolean {
  return left.char === right.char && left.font === right.font;
}

/**
 * Parse the persisted recent row. Fail-closed: every entry must be a single
 * character plus a normalizable font (or null), so a tampered localStorage value
 * can never inject arbitrary text — or an arbitrary `w:font` value — into a
 * document. Returns [] for any malformed input. PURE.
 */
export function parseRecentOfficeSymbols(raw: unknown): OfficeRecentSymbol[] {
  if (typeof raw !== "string" || !raw) return [];
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return [];
  }
  if (!Array.isArray(parsed)) return [];
  const out: OfficeRecentSymbol[] = [];
  for (const item of parsed) {
    if (!item || typeof item !== "object") continue;
    const candidate = item as { char?: unknown; font?: unknown };
    if (typeof candidate.char !== "string") continue;
    // Exactly ONE code point: a symbol is one character, and a longer string
    // would be arbitrary text smuggled in through the recent row.
    if (Array.from(candidate.char).length !== 1) continue;
    const codePoint = candidate.char.codePointAt(0);
    if (codePoint === undefined || codePoint > 0xffff) continue;
    const font =
      candidate.font === null || candidate.font === undefined
        ? null
        : normalizeSymbolFont(candidate.font);
    if (candidate.font != null && font === null) continue;
    const entry: OfficeRecentSymbol = { char: candidate.char, font };
    if (out.some((existing) => sameRecentSymbol(existing, entry))) continue;
    out.push(entry);
    if (out.length >= OFFICE_SYMBOL_RECENT_LIMIT) break;
  }
  return out;
}

/**
 * The recent row after inserting `entry`: fronted, de-duplicated on
 * (char, font), clamped. PURE, so the caller's storage write is a one-liner.
 */
export function pushRecentOfficeSymbol(
  recent: readonly OfficeRecentSymbol[],
  entry: OfficeRecentSymbol,
): OfficeRecentSymbol[] {
  if (Array.from(entry.char).length !== 1) return [...recent];
  const font = entry.font === null ? null : normalizeSymbolFont(entry.font);
  if (entry.font !== null && font === null) return [...recent];
  const normalized: OfficeRecentSymbol = { char: entry.char, font };
  return [
    normalized,
    ...recent.filter((item) => !sameRecentSymbol(item, normalized)),
  ].slice(0, OFFICE_SYMBOL_RECENT_LIMIT);
}

// ---------------------------------------------------------------------------
// The `symbolChar` mark's attributes
// ---------------------------------------------------------------------------

/**
 * The mark name every layer agrees on for a Word symbol character's identity.
 * The TipTap mark itself lives in
 * `components/console/editor/native-document/symbolCharMark`; the NAME and the
 * attribute validator live here so the importer, the exporter, and the paged
 * projection can all reach them without importing an editor extension.
 */
export const SYMBOL_CHAR_MARK = "symbolChar";

export interface SymbolCharMarkAttrs {
  /** `w:font` — the font the glyph is pulled from. */
  font: string;
  /** `w:char` — canonical 4-hex-digit UPPERCASE character code. */
  charCode: string;
}

/**
 * Validate a (font, charCode) pair into the exact attrs the exporter may write,
 * or null. FAIL-CLOSED — this is the only gate between a pasted, imported, or
 * persisted value and the `w:font` / `w:char` attributes of an exported
 * document, so a value that is not a plain font name plus a short hex number
 * never reaches them. PURE.
 */
export function normalizeSymbolCharAttrs(
  attrs: unknown,
): SymbolCharMarkAttrs | null {
  if (!attrs || typeof attrs !== "object") return null;
  const candidate = attrs as { font?: unknown; charCode?: unknown };
  const font = normalizeSymbolFont(candidate.font);
  const charCode = normalizeSymbolCharCode(candidate.charCode);
  if (font === null || charCode === null) return null;
  return { font, charCode };
}

// ---------------------------------------------------------------------------
// What the dialog hands back
// ---------------------------------------------------------------------------

/**
 * One insertion the dialog requests. `font === null` is `(normal text)`: insert
 * `char` as an ordinary character. Otherwise the character is a SYMBOL
 * character: insert `char` carrying the symbol identity (`font`, `charCode`) so
 * the exporter can write `<w:sym w:font w:char/>` instead of `<w:t>`.
 */
export interface OfficeSymbolInsertion {
  readonly char: string;
  readonly font: string | null;
  /** Canonical 4-hex `w:char` spelling. Present exactly when `font` is set. */
  readonly charCode: string | null;
  /** Non-"text" kinds write a dedicated run child instead of literal text. */
  readonly kind: SpecialCharacterKind;
}

/**
 * Build the insertion for a grid cell. Returns null when the code point cannot
 * be a symbol character (outside the BMP / not an integer), so the caller never
 * inserts something it cannot write back out.
 */
export function symbolInsertionForCodePoint(
  font: string | null,
  codePoint: number,
): OfficeSymbolInsertion | null {
  if (!Number.isInteger(codePoint) || codePoint < 0 || codePoint > 0xffff) {
    return null;
  }
  const char = String.fromCodePoint(codePoint);
  if (font === null) {
    return { char, font: null, charCode: null, kind: "text" };
  }
  const normalizedFont = normalizeSymbolFont(font);
  if (normalizedFont === null) return null;
  const charCode = symbolCharCodeForFont(normalizedFont, codePoint);
  if (charCode === null) return null;
  return { char, font: normalizedFont, charCode, kind: "text" };
}

/** Build the insertion for a Special Characters row. */
export function symbolInsertionForSpecialCharacter(
  entry: OfficeSpecialCharacter,
): OfficeSymbolInsertion {
  return {
    char: entry.char,
    font: null,
    charCode: null,
    kind: entry.kind ?? "text",
  };
}

/** Build the insertion for a recent-row entry. */
export function symbolInsertionForRecent(
  entry: OfficeRecentSymbol,
): OfficeSymbolInsertion | null {
  const codePoint = entry.char.codePointAt(0);
  if (codePoint === undefined) return null;
  // A remembered special character keeps its dedicated-element kind, so a
  // recent optional hyphen still exports as <w:softHyphen/>.
  const special = OFFICE_SPECIAL_CHARACTERS.find(
    (item) => item.char === entry.char,
  );
  if (entry.font === null && special) {
    return symbolInsertionForSpecialCharacter(special);
  }
  return symbolInsertionForCodePoint(entry.font, codePoint);
}

// ---------------------------------------------------------------------------
// Subset, Character code / from, and the default recent row
// ---------------------------------------------------------------------------

/**
 * Word's Symbol dialog, the parts that locate a character (Microsoft Support,
 * "Insert a symbol in Word"; "Insert ASCII or Unicode Latin-based symbols and
 * characters"):
 *   - Subset: for a Unicode font, the Unicode block the selected character
 *     belongs to; picking a block scrolls the grid to its first character.
 *     A legacy symbol font has no subsets (Word hides the box).
 *   - Character code + from: the selected character's code in the chosen
 *     notation. For a Unicode font "from" is Unicode (hex), ASCII (decimal)
 *     or ASCII (hex); ASCII is the Windows-1252 code page, so a character
 *     outside it has no ASCII code. For a symbol font it is Symbol (decimal)
 *     or Symbol (hex), the code in the font (0x20-0xFF). Typing a code selects
 *     that character.
 * The block names and bounds are the Unicode Character Database's
 * (Blocks.txt); only blocks the grid shows are listed.
 */
export interface OfficeSymbolSubset {
  readonly id: string;
  readonly first: number;
  readonly last: number;
}

export const OFFICE_SYMBOL_SUBSETS: readonly OfficeSymbolSubset[] = [
  { id: "basicLatin", first: 0x0000, last: 0x007f },
  { id: "latin1Supplement", first: 0x0080, last: 0x00ff },
  { id: "latinExtendedA", first: 0x0100, last: 0x017f },
  { id: "latinExtendedB", first: 0x0180, last: 0x024f },
  { id: "spacingModifierLetters", first: 0x02b0, last: 0x02ff },
  { id: "greekAndCoptic", first: 0x0370, last: 0x03ff },
  { id: "cyrillic", first: 0x0400, last: 0x04ff },
  { id: "generalPunctuation", first: 0x2000, last: 0x206f },
  { id: "superscriptsAndSubscripts", first: 0x2070, last: 0x209f },
  { id: "currencySymbols", first: 0x20a0, last: 0x20cf },
  { id: "letterlikeSymbols", first: 0x2100, last: 0x214f },
  { id: "numberForms", first: 0x2150, last: 0x218f },
  { id: "arrows", first: 0x2190, last: 0x21ff },
  { id: "mathematicalOperators", first: 0x2200, last: 0x22ff },
  { id: "miscellaneousTechnical", first: 0x2300, last: 0x23ff },
  { id: "enclosedAlphanumerics", first: 0x2460, last: 0x24ff },
  { id: "boxDrawing", first: 0x2500, last: 0x257f },
  { id: "blockElements", first: 0x2580, last: 0x259f },
  { id: "geometricShapes", first: 0x25a0, last: 0x25ff },
  { id: "miscellaneousSymbols", first: 0x2600, last: 0x26ff },
  { id: "dingbats", first: 0x2700, last: 0x27bf },
  { id: "cjkSymbolsAndPunctuation", first: 0x3000, last: 0x303f },
  { id: "hiragana", first: 0x3040, last: 0x309f },
  { id: "katakana", first: 0x30a0, last: 0x30ff },
  { id: "halfwidthAndFullwidthForms", first: 0xff00, last: 0xffef },
];

/** The Unicode block a code point belongs to, among the listed subsets. */
export function symbolSubsetForCodePoint(
  codePoint: number,
): OfficeSymbolSubset | null {
  return (
    OFFICE_SYMBOL_SUBSETS.find(
      (subset) => codePoint >= subset.first && codePoint <= subset.last,
    ) ?? null
  );
}

/** The subsets that hold at least one of the grid's code points, in order. */
export function symbolSubsetsForCodePoints(
  codePoints: readonly number[],
): OfficeSymbolSubset[] {
  return OFFICE_SYMBOL_SUBSETS.filter((subset) =>
    codePoints.some(
      (codePoint) => codePoint >= subset.first && codePoint <= subset.last,
    ),
  );
}

export type OfficeSymbolCodeBase =
  | "unicodeHex"
  | "asciiDecimal"
  | "asciiHex"
  | "symbolDecimal"
  | "symbolHex";

export const OFFICE_SYMBOL_UNICODE_CODE_BASES: readonly OfficeSymbolCodeBase[] =
  ["unicodeHex", "asciiDecimal", "asciiHex"];
export const OFFICE_SYMBOL_FONT_CODE_BASES: readonly OfficeSymbolCodeBase[] = [
  "symbolDecimal",
  "symbolHex",
];

/** Windows-1252 bytes 0x80-0x9F (the rest of 0x00-0xFF is Latin-1). */
const WINDOWS_1252_HIGH: Readonly<Record<number, number>> = {
  0x80: 0x20ac, 0x82: 0x201a, 0x83: 0x0192, 0x84: 0x201e, 0x85: 0x2026,
  0x86: 0x2020, 0x87: 0x2021, 0x88: 0x02c6, 0x89: 0x2030, 0x8a: 0x0160,
  0x8b: 0x2039, 0x8c: 0x0152, 0x8e: 0x017d, 0x91: 0x2018, 0x92: 0x2019,
  0x93: 0x201c, 0x94: 0x201d, 0x95: 0x2022, 0x96: 0x2013, 0x97: 0x2014,
  0x98: 0x02dc, 0x99: 0x2122, 0x9a: 0x0161, 0x9b: 0x203a, 0x9c: 0x0153,
  0x9e: 0x017e, 0x9f: 0x0178,
};

function windows1252ByteForCodePoint(codePoint: number): number | null {
  for (const [byte, mapped] of Object.entries(WINDOWS_1252_HIGH)) {
    if (mapped === codePoint) return Number(byte);
  }
  if (codePoint >= 0x80 && codePoint <= 0x9f) return null;
  return codePoint >= 0 && codePoint <= 0xff ? codePoint : null;
}

function codePointForWindows1252Byte(byte: number): number | null {
  if (byte >= 0x80 && byte <= 0x9f) return WINDOWS_1252_HIGH[byte] ?? null;
  return byte >= 0 && byte <= 0xff ? byte : null;
}

/** Whether a code point is one character the dialog may insert. */
function insertableSymbolCodePoint(codePoint: number): boolean {
  return (
    Number.isInteger(codePoint) &&
    codePoint >= 0x20 &&
    codePoint <= 0xffff &&
    !(codePoint >= 0x7f && codePoint <= 0x9f) &&
    !(codePoint >= 0xd800 && codePoint <= 0xdfff)
  );
}

/**
 * The Character code box's text for a grid code point in a notation, or ""
 * when the character has no code there (a character outside Windows-1252 in
 * ASCII). PURE.
 */
export function symbolCodeForBase(
  codePoint: number,
  base: OfficeSymbolCodeBase,
): string {
  switch (base) {
    case "unicodeHex":
      return symbolCharCodeFromCodePoint(codePoint) ?? "";
    case "asciiDecimal":
    case "asciiHex": {
      const byte = windows1252ByteForCodePoint(codePoint);
      if (byte === null) return "";
      return base === "asciiDecimal"
        ? String(byte)
        : byte.toString(16).toUpperCase().padStart(2, "0");
    }
    case "symbolDecimal":
    case "symbolHex": {
      const byte =
        codePoint >= SYMBOL_PUA_OFFSET && codePoint <= SYMBOL_PUA_OFFSET + 0xff
          ? codePoint - SYMBOL_PUA_OFFSET
          : codePoint <= 0xff
            ? codePoint
            : null;
      if (byte === null) return "";
      return base === "symbolDecimal"
        ? String(byte)
        : byte.toString(16).toUpperCase().padStart(2, "0");
    }
  }
}

/**
 * The code point a typed Character code names in a notation, or null when
 * the text is not a code there (fail closed: nothing is selected). For a
 * symbol font the result is the font's PUA code point, which is what its grid
 * holds. PURE.
 */
export function symbolCodePointFromCode(
  raw: string,
  base: OfficeSymbolCodeBase,
): number | null {
  const text = raw.trim();
  const hex = base === "unicodeHex" || base === "asciiHex" || base === "symbolHex";
  if (hex ? !/^[0-9A-Fa-f]{1,4}$/.test(text) : !/^\d{1,5}$/.test(text)) {
    return null;
  }
  const value = Number.parseInt(text, hex ? 16 : 10);
  let codePoint: number | null;
  if (base === "unicodeHex") codePoint = value;
  else if (base === "asciiDecimal" || base === "asciiHex") {
    codePoint = value <= 0xff ? codePointForWindows1252Byte(value) : null;
  } else {
    codePoint =
      value >= 0x20 && value <= 0xff ? SYMBOL_PUA_OFFSET + value : null;
  }
  return codePoint !== null && insertableSymbolCodePoint(codePoint)
    ? codePoint
    : null;
}

/**
 * The row a fresh Word install shows under "Recently used symbols" before
 * the user has inserted anything: € £ ¥ © ® ™ ± ≠ ≤ ≥ ÷ × ∞ µ α β, all in
 * (normal text). Shown until the user's own row exists.
 */
export const OFFICE_SYMBOL_DEFAULT_RECENT: readonly OfficeRecentSymbol[] = [
  "€", "£", "¥", "©", "®", "™", "±", "≠",
  "≤", "≥", "÷", "×", "∞", "µ", "α", "β",
].map((char) => ({ char, font: null }));
