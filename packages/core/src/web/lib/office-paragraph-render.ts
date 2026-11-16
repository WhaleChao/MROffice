/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How a DrawingML paragraph's MARKER is drawn — for any host that draws one.
 *
 * `a:pPr` is the same element in a slide's `p:sp`, a worksheet drawing's
 * `xdr:sp` and a document's `wps:wsp`, so the ladder of default glyphs, the
 * spelling of each `ST_TextAutonumberScheme` member and the rule that decides
 * which number a row shows are one thing, not three. They lived in
 * `presentation-paragraphs.ts` under a deck's name while the sheet
 * (`office-text-layout.ts`) and the document already called them; the audit
 * recorded that as naming debt twice (SHAPE-IMPORT-AUDIT.md §44.7). This module
 * is the name. `presentation-paragraphs.ts` re-exports every symbol, so the
 * deck's own import sites are unchanged.
 *
 * Dependency-free apart from the ONE list catalogue (no React, no DOM), so it
 * can be imported from the zod schema, the editor, the present overlay and the
 * sheet's shape text alike.
 *
 * What Microsoft publishes about the numbering: `a:buAutoNum` is
 * LEVEL-SENSITIVE (three markers at levels 0, 1, 0 number 1, 1, 2), `@startAt`
 * names the number the sequence starts at, and the alphabetic sequence is
 * 1 → `a`, 27 → `aa`, 53 → `aaa` (Office REPEATS the letter; it does not count
 * in base 26). What it does NOT publish: how any member is actually DRAWN — the
 * SDK enum page gives names only and MS-OODF3's ODF mapping is silent on the
 * `*Plain` delimiters and flattens the circled families to a plain `1`, which
 * is not what either app draws. Everything below that Microsoft does not state
 * was MEASURED, in real PowerPoint and in real Excel, and is recorded row by
 * row in `office-autonumber.measured.json`.
 */

import {
  OFFICE_BULLET_STYLE_GLYPHS,
  POWERPOINT_BULLET_CHARS,
  POWERPOINT_DEFAULT_BULLET_GLYPH_BY_LEVEL,
  POWERPOINT_NATIVE_BULLET_CHARS,
  POWERPOINT_WEB_BULLET_STYLE_TYPES,
  POWERPOINT_WEB_NUMBER_FORMATS,
} from "./office-list-catalog";

/** Indent bounds (PowerPoint shows 5 outline levels in its model: 0..4). */
export const MIN_PARAGRAPH_LEVEL = 0;
export const MAX_PARAGRAPH_LEVEL = 4;

/** Clamp an arbitrary value to a valid integer indent level. */
export function clampLevel(value: unknown): number {
  const numeric =
    typeof value === "number" && Number.isFinite(value) ? Math.round(value) : 0;
  return Math.min(MAX_PARAGRAPH_LEVEL, Math.max(MIN_PARAGRAPH_LEVEL, numeric));
}

/**
 * Per-level bullet glyph: solid → en-dash → middle-dot, repeating by depth.
 *
 * From the ONE list catalogue Docs and Presentation share.
 */
const BULLET_GLYPHS = POWERPOINT_DEFAULT_BULLET_GLYPH_BY_LEVEL;

/** The bullet glyph (•/–/·) shown for a bulleted paragraph at a given level. */
export function bulletGlyphForLevel(level: number): string {
  const index = clampLevel(level);
  return BULLET_GLYPHS[index] ?? BULLET_GLYPHS[0];
}

/**
 * Durable bullet glyphs accepted by the model. The final two are legacy MRAI
 * choices kept for existing saved pages; the visible Web gallery is the
 * separate {@link PRESENTATION_WEB_BULLET_STYLES} subset below.
 */
export const BULLET_CHARS = POWERPOINT_BULLET_CHARS;
export type BulletChar = (typeof BULLET_CHARS)[number];
const BULLET_CHAR_SET: ReadonlySet<string> = new Set(BULLET_CHARS);

/** The exact PowerPoint-for-the-web gallery, backed by the shared Docs data. */
export const PRESENTATION_WEB_BULLET_STYLES =
  POWERPOINT_WEB_BULLET_STYLE_TYPES.map((style) => ({
    style,
    glyph: OFFICE_BULLET_STYLE_GLYPHS[style] as BulletChar,
  }));

/**
 * Normalize an arbitrary persisted bullet-glyph choice into an allowed
 * {@link BulletChar}, or `undefined` when absent / unsupported so a block
 * WITHOUT an explicit glyph stays byte-clean and falls back to the per-level
 * default.
 */
export function normalizeBulletChar(value: unknown): BulletChar | undefined {
  if (typeof value !== "string") return undefined;
  const candidate = value.trim();
  return BULLET_CHAR_SET.has(candidate) ? (candidate as BulletChar) : undefined;
}

/**
 * The glyphs a text BLOCK's own `bulletChar` may hold: the gallery's, plus the
 * middle dot an imported box states for all its bulleted rows (reading row
 * 817). Rows keep the gallery vocabulary ({@link normalizeBulletChar}).
 */
export const BLOCK_BULLET_CHARS = POWERPOINT_NATIVE_BULLET_CHARS;
export type BlockBulletChar = (typeof BLOCK_BULLET_CHARS)[number];
const BLOCK_BULLET_CHAR_SET: ReadonlySet<string> = new Set(BLOCK_BULLET_CHARS);

/** {@link normalizeBulletChar} for a block's own marker. */
export function normalizeBlockBulletChar(value: unknown): BlockBulletChar | undefined {
  if (typeof value !== "string") return undefined;
  const candidate = value.trim();
  return BLOCK_BULLET_CHAR_SET.has(candidate) ? (candidate as BlockBulletChar) : undefined;
}

/**
 * The numbered-list formats an author may PICK for a whole text block
 * (PowerPoint / WPS "Numbering" gallery subset).
 */
export const NUMBER_FORMATS = POWERPOINT_WEB_NUMBER_FORMATS;
export type NumberFormat = (typeof NUMBER_FORMATS)[number];
const NUMBER_FORMAT_SET: ReadonlySet<string> = new Set(NUMBER_FORMATS);

/** Normalize an arbitrary persisted number-format choice, or drop it. */
export function normalizeNumberFormat(value: unknown): NumberFormat | undefined {
  if (typeof value !== "string") return undefined;
  const candidate = value.trim().toLowerCase();
  return NUMBER_FORMAT_SET.has(candidate)
    ? (candidate as NumberFormat)
    : undefined;
}

/**
 * The canonical auto-number format PowerPoint applies at an outline level when
 * the text body does not select one block-wide format.
 */
export function numberFormatForLevel(level: number): NumberFormat {
  switch (clampLevel(level) % 3) {
    case 1:
      return "lower-alpha";
    case 2:
      return "lower-roman";
    default:
      return "decimal";
  }
}

/**
 * Convert a 1-based ordinal to a lower-case alphabetic label.
 *
 * Office REPEATS the letter rather than counting in base 26, and Microsoft
 * states the sequence on the `AutoNumberedBullet` (`a:buAutoNum`) page itself:
 * 1 → `a`, 27 → `aa`, 53 → `aaa`. The repeat is bounded because
 * `a:buAutoNum@startAt` reaches 32767 and a label of a thousand letters is a
 * denial of service, not a number.
 */
const MAX_ALPHA_LABEL_REPEAT = 40;

function toAlpha(ordinal: number): string {
  const n = Math.max(1, Math.floor(ordinal));
  const letter = String.fromCharCode(97 + ((n - 1) % 26));
  const repeat = Math.min(
    MAX_ALPHA_LABEL_REPEAT,
    Math.floor((n - 1) / 26) + 1,
  );
  return letter.repeat(repeat);
}

/** Convert a 1-based ordinal to an UPPER-case alphabetic label (1→A, 27→AA). */
function toUpperAlpha(ordinal: number): string {
  return toAlpha(ordinal).toUpperCase();
}

/** Convert a 1-based ordinal to a lower-case Roman numeral (1→i, 4→iv). */
function toRoman(ordinal: number): string {
  const n = Math.max(1, Math.min(3999, Math.floor(ordinal)));
  const table: Array<[number, string]> = [
    [1000, "m"],
    [900, "cm"],
    [500, "d"],
    [400, "cd"],
    [100, "c"],
    [90, "xc"],
    [50, "l"],
    [40, "xl"],
    [10, "x"],
    [9, "ix"],
    [5, "v"],
    [4, "iv"],
    [1, "i"],
  ];
  let remaining = n;
  let out = "";
  for (const [value, symbol] of table) {
    while (remaining >= value) {
      out += symbol;
      remaining -= value;
    }
  }
  return out;
}

/**
 * Render a numbered label in an EXPLICITLY-chosen format (block-level override),
 * independent of indent level. Always returns a trailing-dot label so it matches
 * the per-level scheme's shape (`1.` / `a.` / `i.` / `A.`).
 */
export function numberLabelForFormat(
  format: NumberFormat,
  ordinal: number,
): string {
  const safeOrdinal = Math.max(1, Math.floor(ordinal));
  switch (format) {
    case "decimal-paren":
      return `${safeOrdinal})`;
    case "lower-alpha":
      return `${toAlpha(safeOrdinal)}.`;
    case "lower-alpha-paren":
      return `${toAlpha(safeOrdinal)})`;
    case "lower-roman":
      return `${toRoman(safeOrdinal)}.`;
    case "upper-alpha":
      return `${toUpperAlpha(safeOrdinal)}.`;
    case "upper-roman":
      return `${toRoman(safeOrdinal).toUpperCase()}.`;
    default:
      return `${safeOrdinal}.`;
  }
}

/**
 * The auto-number label for a numbered paragraph. The per-level scheme cycles
 * arabic → alpha → roman so nested numbered lists stay legible:
 *   level 0 → "1."   level 1 → "a."   level 2 → "i."   (then repeats)
 * `ordinal` is the 1-based position within the contiguous same-level run.
 */
export function autoNumberLabel(level: number, ordinal: number): string {
  return numberLabelForFormat(numberFormatForLevel(level), ordinal);
}

/** Fullwidth digits, for the `*Db*` (double-byte) schemes. */
function toFullWidthDigits(ordinal: number): string {
  return `${Math.max(1, Math.floor(ordinal))}`.replace(/[0-9]/g, (digit) =>
    String.fromCharCode(0xff10 + digit.charCodeAt(0) - 48),
  );
}

/**
 * The circled-number ladders, measured rather than guessed.
 *
 * Neither host's circled schemes are described anywhere Microsoft publishes —
 * MS-OODF3 flattens all three to a plain `1` for its ODF round trip, which is
 * not what either app draws. A probe of six rows per scheme, opened by real
 * PowerPoint and by real Excel and read out of each app's own PDF export
 * (SHAPE-IMPORT-AUDIT.md §47.1), drew:
 *
 * - `circleNumDbPlain` 1/10/11/20 → U+2460 ①, U+2469 ⑩, U+246A ⑪, U+2473 ⑳,
 *   and 21 → the bare characters `21`: the ladder simply ENDS at twenty.
 * - `circleNumWdWhitePlain` 1/2/9/10/11/12 → Wingdings 0x81/0x82/0x89/0x8A and
 *   then 0x81/0x82 again, i.e. it WRAPS every ten.
 * - `circleNumWdBlackPlain` at the same ordinals → 0x8C/0x8D/0x94/0x95/0x8C/0x8D.
 *
 * Those Wingdings code points are drawn here as their Unicode counterparts
 * rather than as private-use characters in a font the browser may not have.
 * That substitution is not a guess either: the system Wingdings `cmap` names
 * 0x81..0x8A `onesans`..`tensans` and 0x8C..0x95 `onesansinv`..`tensansinv`,
 * which is exactly U+2460..U+2469 (circled digits) and U+2776..U+277F (negative
 * circled digits). It is the same trade `SYMBOL_FONT_UNICODE_FALLBACK` makes
 * for `a:buChar`.
 */
const CIRCLED_LADDER_LIMIT = 20;
const WINGDINGS_CIRCLED_WRAP = 10;

function circledNumberDb(ordinal: number): string {
  const n = Math.max(1, Math.floor(ordinal));
  return n <= CIRCLED_LADDER_LIMIT
    ? String.fromCharCode(0x2460 + n - 1)
    : `${n}`;
}

function circledNumberWingdings(ordinal: number, base: number): string {
  const n = Math.max(1, Math.floor(ordinal));
  return String.fromCharCode(base + ((n - 1) % WINGDINGS_CIRCLED_WRAP));
}

const IDEOGRAPHIC_DIGITS = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九"];

/** 1..99 in Chinese numerals: 十, 十一, 二十, 九十九. */
function ideographicNumeral(ordinal: number): string {
  const n = Math.max(1, Math.floor(ordinal));
  if (n < 10) return IDEOGRAPHIC_DIGITS[n];
  const tens = Math.floor(n / 10);
  return `${tens === 1 ? "" : IDEOGRAPHIC_DIGITS[tens]}十${IDEOGRAPHIC_DIGITS[n % 10]}`;
}

/**
 * The `ST_TextAutonumberScheme` members this renderer can SPELL.
 *
 * DrawingML builds most of these names as alphabet + case + punctuation, so
 * `romanUcParenBoth` is `(I)` and `alphaLcPeriod` is `a.` with nothing left to
 * interpret. The editable model can NAME seven of the forty (the gallery), and
 * the `nativeBullet` receipt carries the rest verbatim (§31.2).
 *
 * Deliberately NOT a table of all forty. The East-Asian, Hebrew, Thai and Hindi
 * families pick digits out of another alphabet and only ordinal 1 of them was
 * ever measured; a ladder there would replace "the wrong number" with "the
 * wrong number in a different script". They fall back to the per-level scheme
 * and stay listed in the audit's remainder.
 */
const NATIVE_AUTONUMBER_LABEL: Readonly<
  Record<string, (ordinal: number) => string>
> = Object.freeze({
  // The `*Plain` members carry NO delimiter at all. MS-OODF3 is silent on them
  // and both hosts drew a bare number (§47.1).
  arabicPlain: (n) => `${n}`,
  arabicDbPlain: (n) => toFullWidthDigits(n),
  arabicDbPeriod: (n) => `${toFullWidthDigits(n)}．`,
  circleNumDbPlain: (n) => circledNumberDb(n),
  circleNumWdWhitePlain: (n) => circledNumberWingdings(n, 0x2460),
  circleNumWdBlackPlain: (n) => circledNumberWingdings(n, 0x2776),
  // PowerPoint's "Ideographic Numbering, wide period" (一． 二． 三．),
  // ECMA-376's Japanese / Simplified Chinese ideograph with a double-byte
  // period. Spelled to 99, past which it falls back to the level default.
  ea1JpnChsDbPeriod: (n) => (n <= 99 ? `${ideographicNumeral(n)}．` : `${n}.`),
  arabicPeriod: (n) => `${n}.`,
  arabicParenR: (n) => `${n})`,
  arabicParenBoth: (n) => `(${n})`,
  alphaLcPeriod: (n) => `${toAlpha(n)}.`,
  alphaLcParenR: (n) => `${toAlpha(n)})`,
  alphaLcParenBoth: (n) => `(${toAlpha(n)})`,
  alphaUcPeriod: (n) => `${toUpperAlpha(n)}.`,
  alphaUcParenR: (n) => `${toUpperAlpha(n)})`,
  alphaUcParenBoth: (n) => `(${toUpperAlpha(n)})`,
  romanLcPeriod: (n) => `${toRoman(n)}.`,
  romanLcParenR: (n) => `${toRoman(n)})`,
  romanLcParenBoth: (n) => `(${toRoman(n)})`,
  romanUcPeriod: (n) => `${toRoman(n).toUpperCase()}.`,
  romanUcParenR: (n) => `${toRoman(n).toUpperCase()})`,
  romanUcParenBoth: (n) => `(${toRoman(n).toUpperCase()})`,
});

/**
 * The label for one row the PACKAGE numbered, or undefined when this renderer
 * cannot spell that scheme and the per-level default should be used instead.
 */
export function nativeAutoNumberLabel(
  scheme: string | undefined,
  ordinal: number,
): string | undefined {
  if (scheme === undefined) return undefined;
  const spell = NATIVE_AUTONUMBER_LABEL[scheme];
  return spell === undefined ? undefined : spell(Math.max(1, Math.floor(ordinal)));
}

/** The shape one row must have for this renderer to number it. */
export interface RenderableParagraph {
  text: string;
  level: number;
  bullet: "none" | "bullet" | "number";
  /**
   * This row's own gallery choice, over the block's. PowerPoint states the
   * marker per paragraph (`a:pPr/a:buChar`, `a:buAutoNum`), so Bullets on one
   * item of a list changes that item alone.
   */
  bulletChar?: string;
  numberFormat?: string;
  /**
   * The row's OWN marker formatting (Bullets and Numbering), in the receipt's
   * keys. Where it is present it is what the row draws, over the positional
   * receipt below.
   */
  bulletFormat?: RenderableMarker;
  /**
   * A list row whose marker its layout does not draw (a SmartArt layout's
   * level-1 words, `SlideParagraph.markerHidden`): drawn as a row with none.
   */
  markerHidden?: true;
}

/** The marker identity this renderer reads off one row's receipt. */
export interface RenderableMarker {
  autoNum?: string;
  startAt?: number;
}

/**
 * A render-ready descriptor for one paragraph: its text, indent depth, and the
 * marker to show in the hanging-indent gutter (a bullet glyph, an auto-number
 * label, or `null` for an un-bulleted paragraph).
 */
export interface RenderedParagraph {
  text: string;
  level: number;
  marker: string | null;
}

/**
 * Block-level marker overrides (the author's picked bullet glyph / number
 * format). Both optional: an absent field keeps the per-level default scheme so
 * a block that never picked a glyph/format renders byte-for-byte as before.
 */
export interface MarkerOptions {
  bulletChar?: BlockBulletChar;
  numberFormat?: NumberFormat;
  /**
   * The `nativeBullet` receipt, aligned to the rows by index. Where a row's
   * entry carries an `autoNum` this renderer spells THAT scheme and starts the
   * run at its `startAt`, instead of the bounded gallery's per-level cycle.
   */
  nativeBullet?: readonly (RenderableMarker | null)[];
}

/** One level's live numbered run: its identity, and the last number drawn. */
interface NumberRunState {
  scheme: string;
  startAt: number;
  ordinal: number | undefined;
}










export function renderParagraphs(
  paragraphs: readonly RenderableParagraph[],
  options?: MarkerOptions,
): RenderedParagraph[] {
  const bulletChar = normalizeBlockBulletChar(options?.bulletChar);
  const numberFormat = normalizeNumberFormat(options?.numberFormat);
  // Read POSITIONALLY, so a length that no longer matches the rows means the
  // rows were rewritten under it and the receipt says nothing (the rule every
  // reader of it holds).
  const native =
    options?.nativeBullet?.length === paragraphs.length
      ? options.nativeBullet
      : undefined;
  // One run state per level. The run's IDENTITY is (scheme, startAt) and its
  // `ordinal` is the last number actually DRAWN at this level, or undefined
  // when the identity has not drawn anything yet.
  const runs = new Map<number, NumberRunState>();
  const clearDeeperThan = (level: number) => {
    for (const key of [...runs.keys()]) {
      if (key > level) runs.delete(key);
    }
  };
  return paragraphs.map((paragraph, index) => {
    const level = clampLevel(paragraph.level);
    if (paragraph.markerHidden === true && paragraph.bullet !== "none") {
      // Drawn as PowerPoint draws it: no marker, and the level's numbering
      // run broken, as a row without one breaks it.
      runs.delete(level);
      clearDeeperThan(level);
      return { text: paragraph.text, level, marker: null };
    }
    const rowBulletChar = normalizeBulletChar(paragraph.bulletChar) ?? bulletChar;
    const rowNumberFormat =
      normalizeNumberFormat(paragraph.numberFormat) ?? numberFormat;
    if (paragraph.bullet === "number") {
      const ownFormat = paragraph.bulletFormat;
      const nativeRow =
        ownFormat !== undefined ? ownFormat : (native?.[index] ?? null);
      // What a numbered row means was MEASURED, in both hosts, from probes
      // opened by real PowerPoint and real Excel and read out of each app's own
      // PDF export character by character (SHAPE-IMPORT-AUDIT.md §40.2, §47.1).
      // The two apps drew the same thing on every row of every case.
      //
      // 1. `a:buAutoNum@startAt` is stated PER PARAGRAPH. Its effective value
      //    (1 when absent) is compared with the previous row's at the same
      //    level, and the numbering RESTARTS where the value CHANGES:
      //      3,3,3 -> 3. 4. 5.   3,-,- -> 3. 1. 2.   -,5,- -> 1. 5. 1.
      //      -,1,- -> 1. 2. 3.   (so a stated `startAt="1"` needs no receipt)
      // 2. The SCHEME is part of the same identity. `arabicPeriod` x3 then
      //    `alphaUcPeriod` x2 drew `1. 2. 3. A. B.`, and changing back drew
      //    `1. 2. A. 1. 2.` -- a scheme change restarts just as a `startAt`
      //    change does. This is the rule the corpus text box S047 needed: it
      //    switches scheme half way down.
      // 3. A deeper run does not reset the parent: `5. 6. / 3. 4. / 7. 8.`
      //    with the level-0 rows all stating startAt 5, and `5. 6. / 3. /
      //    1. 2.` when the resumed level-0 rows state none.
      const scheme =
        nativeRow?.autoNum ??
        rowNumberFormat ??
        `level:${numberFormatForLevel(level)}`;
      // A row's own Start at applies to a gallery format too (Bullets and
      // Numbering > Start at); a receipt only ever states one with its scheme.
      const startAt =
        ownFormat !== undefined || nativeRow?.autoNum !== undefined
          ? (nativeRow?.startAt ?? 1)
          : 1;
      const previous = runs.get(level);
      const continuing =
        previous !== undefined &&
        previous.scheme === scheme &&
        previous.startAt === startAt;
      const carried = continuing ? previous.ordinal : undefined;
      const ordinal = carried === undefined ? startAt : carried + 1;
      // 4. An EMPTY numbered paragraph draws NOTHING and does NOT consume its
      //    ordinal: `1. 2. <empty> <empty>` continued at `3.`, and a row
      //    stating startAt 3 either side of an empty row stating the same drew
      //    `3.` then `4.`. It still carries the run identity forward, which is
      //    why an empty row stating a DIFFERENT startAt made the row after it
      //    restart (`1. <empty startAt=7> 1.`). Before this was measured the
      //    canvas numbered the empty rows too, so S047 drew `6.` where both
      //    apps draw `4.`.
      const draws = paragraph.text.trim().length > 0;
      runs.set(level, {
        scheme,
        startAt,
        ordinal: draws ? ordinal : carried,
      });
      clearDeeperThan(level);
      if (!draws) return { text: paragraph.text, level, marker: null };
      // The scheme the PACKAGE wrote wins where this renderer can spell it; a
      // picked format overrides the per-level cycling scheme; otherwise the
      // arabic→alpha→roman per-level default is used (back-compat).
      const marker =
        nativeAutoNumberLabel(nativeRow?.autoNum, ordinal) ??
        (rowNumberFormat
          ? numberLabelForFormat(rowNumberFormat, ordinal)
          : autoNumberLabel(level, ordinal));
      return { text: paragraph.text, level, marker };
    }
    // A non-number paragraph breaks its own level's run (and any deeper one) --
    // measured: `1. 2. <empty, a:buNone> 3.` drew `1. 2.` and then `1.`.
    runs.delete(level);
    clearDeeperThan(level);
    return {
      text: paragraph.text,
      level,
      // A picked glyph overrides the per-level default for a bulleted paragraph.
      marker:
        paragraph.bullet === "bullet"
          ? (rowBulletChar ?? bulletGlyphForLevel(level))
          : null,
    };
  });
}
