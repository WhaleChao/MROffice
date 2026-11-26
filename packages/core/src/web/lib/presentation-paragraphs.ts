/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Bulleted / numbered list paragraphs for presentation slide text blocks
 * (PowerPoint / WPS parity).
 *
 * A presentation text block historically carries a single flat `text` string.
 * This module adds an OPTIONAL, migration-safe `paragraphs` array: each
 * paragraph is `{ text, level, bullet }` where `level` is an indent depth
 * (0..4) and `bullet` is `"none" | "bullet" | "number"`. When a block carries no
 * `paragraphs` it renders exactly as before (flat text) — the helpers here
 * round-trip a flat string ↔ paragraphs so the editor (a single textarea) can
 * keep splitting on newlines while the structure is preserved.
 *
 * Keep this module dependency-free (no React / no DOM) so it can be unit tested
 * in isolation and imported from the zod schema, the editor, and the present
 * overlay alike.
 */

import {
  MAX_PARAGRAPH_LEVEL,
  MIN_PARAGRAPH_LEVEL,
  clampLevel,
  normalizeBulletChar,
  normalizeNumberFormat,
  type BulletChar,
  type NumberFormat,
} from "./office-paragraph-render";
import { isValidInlineImageSource } from "./inline-image-source";
import { normalizeDir } from "../components/console/editor/native-document/documentLayout";

/** The bullet style a paragraph carries. */
export type ParagraphBullet = "none" | "bullet" | "number";

/** DrawingML paragraph alignment values surfaced by the Presentation editor. */
export const PARAGRAPH_ALIGNS = [
  "left",
  "center",
  "right",
  "justify",
  "distribute",
] as const;
export type ParagraphAlign = (typeof PARAGRAPH_ALIGNS)[number];

/** Allowed bullet identifiers, shared with the schema enum. */
export const PARAGRAPH_BULLETS = ["none", "bullet", "number"] as const;

/** A single list paragraph. */
export interface SlideParagraph {
  text: string;
  /** Indent depth, clamped to `MIN_PARAGRAPH_LEVEL`..`MAX_PARAGRAPH_LEVEL`. */
  level: number;
  bullet: ParagraphBullet;
  /** Optional paragraph-local alignment; absent inherits the text-box default. */
  align?: ParagraphAlign;
  /**
   * The row's OWN indentation, in EMU, exactly as DrawingML states it:
   * `a:pPr@marL` is PowerPoint's "Before text" (the whole paragraph's left
   * margin) and `a:pPr@indent` its "Special" first line, NEGATIVE for a
   * hanging indent. Both optional — an absent one means "whatever the outline
   * `level` says", which is what every row carried before the import learned
   * to read them, so a stored page without them is an older reading and not an
   * edit. Exact quantities on purpose: the only thing near them was the
   * block-level `style.indentLevel`, which quantises a whole box to whole half
   * inches (SHAPE-IMPORT-AUDIT.md §19.3).
   */
  marL?: number;
  indent?: number;
  /**
   * PowerPoint's Right-to-Left paragraph direction (`a:pPr@rtl`): the line's
   * bidi base direction, the side its marker sits on and the side its own
   * indentation is measured from. Docs states the same per paragraph as
   * `dir` (`documentLayout`'s one whitelist); absent is left-to-right.
   */
  dir?: "rtl";
  /**
   * This row's own gallery marker, over the block-level `bulletChar` /
   * `numberFormat`. PowerPoint states the marker per paragraph
   * (`a:pPr/a:buChar`, `a:buAutoNum`), and Home > Bullets on one item of a
   * list changes that item alone (PPT-B084). Kept only while the row is of the
   * matching kind; absent follows the block's choice.
   */
  bulletChar?: BulletChar;
  numberFormat?: NumberFormat;
  /**
   * The row's own Bullets and Numbering formatting beyond the gallery
   * (PPT-A039/A040): a glyph in its symbol font (`a:buChar` + `a:buFont`),
   * any numbering scheme and its Start at (`a:buAutoNum@type/@startAt`), the
   * marker's colour (`a:buClr`) and size (`a:buSzPct`). The same keys and
   * units as the block's `nativeBullet` receipt, which the editor turns into
   * this field when it edits the rows, so the package's marker survives a
   * text edit. `{}` states "no formatting of its own".
   */
  bulletFormat?: ParagraphBulletFormat;
  







  markerHidden?: true;
  /**
   * The size, in points, of the row's end mark (`a:endParaRPr@sz`): how tall
   * a BLANK row is in PowerPoint, which no run states. Read by the import
   * (readings row 63, `emptyParagraphEndSize`) and written back by the
   * generated export; the editor only carries it, so an untouched page still
   * proves its original package.
   */
  endFontSize?: number;
  /**
   * The row's end mark set Superscript or Subscript (`a:endParaRPr@baseline`,
   * readings row 730): PowerPoint lays a BLANK row out at the raised size,
   * two thirds of `endFontSize` to a whole point (as it sets a raised run).
   * Carried like `endFontSize`; never written.
   */
  endScript?: "super" | "sub";
  /**
   * Paragraph > Asian Typography > Text alignment (`a:pPr@fontAlgn`,
   * ST_TextFontAlignType): where the characters of a line stand on it when
   * their sizes differ -- Auto, Top, Center, Baseline, Bottom. The row's own
   * value (readings row 513); `fontAlignInherited` is what its level
   * inherits from the list styles when not Auto, for the text engine only,
   * never written.
   */
  fontAlign?: PresentationFontAlign;
  fontAlignInherited?: PresentationFontAlign;
  








  runDefaults?: ParagraphRunDefaults;
  /**
   * The Latin face the row's paragraph mark is set in, on a row with East
   * Asian words (`a:endParaRPr`'s `a:latin`, else the last run's, else the
   * level's; readings row 899). PowerPoint lays the line that holds the mark
   * by it; the runs name only the faces their own script draws in. A flat
   * body says it on the block. For the text engine only; carried, never
   * written.
   */
  endLatinFace?: string;
  /**
   * The row's own kerning threshold in points (readings row 1009): PowerPoint
   * kerns its runs' character pairs from that size up (`a:rPr@kern` as the
   * row's first worded run resolves it; 0: not kerned). The row's share of
   * `source.nativeKerning`, carried with the row -- a row Enter splits off
   * takes it -- for the text engine, and written back by a download rebuilt
   * from the page; never authored. A flat body says it on the block.
   */
  nativeKerning?: number;
  /**
   * The row's own East Asian line rules (readings row 1009; truth look-fit
   * hang13): the `lang` / `altLang` its runs that state none inherit, and
   * `hangingPunct` false, `latinLnBrk` true, `eaLnBrk` false where it
   * resolves them so. The row's share of `source.nativeEastAsianLines`,
   * carried with the row as PowerPoint copies a row's `a:pPr` to both halves
   * of an Enter (office-truth lf13-enter); never authored. A flat body says
   * it on the block.
   */
  nativeEastAsianLines?: ParagraphEastAsianLines;
  /**
   * The row's right indent in EMU (`a:pPr@marR`, else its level's in the
   * shape's own list style; readings row 1029): PowerPoint ends every line of
   * the row that far inside the text area, centred and right-aligned lines
   * between both indents. Carried with the row -- PowerPoint copies a row's
   * `a:pPr` to both halves of an Enter, and has no control that sets it --
   * and written back by a download rebuilt from the page; never authored.
   */
  marR?: number;
  /**
   * The row ends at a line break of its paragraph (`a:br`), not at the
   * paragraph's mark: a row a one-paragraph flat body is split into at its
   * breaks (`paragraphBuildRows`). Built for drawing only, never stored.
   */
  endsInLineBreak?: true;
}

/** {@link SlideParagraph.nativeEastAsianLines}. */
export interface ParagraphEastAsianLines {
  lang?: string;
  altLang?: string;
  hangingPunct?: false;
  latinLnBrk?: true;
  eaLnBrk?: false;
  /** The row's Korean runs hang PowerPoint's Korean marks (readings row 1024). */
  koreanRules?: true;
}

const LINE_RULE_LANGUAGE = /^[A-Za-z]{2,8}(?:-[A-Za-z0-9]{1,8}){0,5}$/;

/** A row's stated kerning threshold (`nativeKerning`, points 0..4000), or undefined. */
export function normalizeParagraphNativeKerning(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) && value >= 0 && value <= 4000 ? value : undefined;
}

/** A row's stated East Asian line rules (`nativeEastAsianLines`), or undefined. */
export function normalizeParagraphEastAsianLines(value: unknown): ParagraphEastAsianLines | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const raw = value as Record<string, unknown>;
  const out: ParagraphEastAsianLines = {};
  if (typeof raw.lang === "string" && LINE_RULE_LANGUAGE.test(raw.lang)) out.lang = raw.lang;
  if (typeof raw.altLang === "string" && LINE_RULE_LANGUAGE.test(raw.altLang)) out.altLang = raw.altLang;
  if (raw.hangingPunct === false) out.hangingPunct = false;
  if (raw.latinLnBrk === true) out.latinLnBrk = true;
  if (raw.eaLnBrk === false) out.eaLnBrk = false;
  if (raw.koreanRules === true) out.koreanRules = true;
  return out;
}

/** A row's stated right indent (`marR`, EMU within ST_TextMargin, above 0), or undefined. */
export function normalizeParagraphRightIndent(value: unknown): number | undefined {
  return typeof value === "number" && Number.isInteger(value) && value > 0 && value <= 51206400 ? value : undefined;
}

/** A stated mark Latin face (`endLatinFace`), or undefined. */
export function normalizeParagraphEndLatinFace(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() && value.length <= 200 ? value.trim() : undefined;
}

/** {@link SlideParagraph.runDefaults}. */
export interface ParagraphRunDefaults {
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  fontFamily?: string;
  letterSpacing?: number;
}

/** A row's run defaults as stored, or undefined when none reads. */
export function normalizeParagraphRunDefaults(value: unknown): ParagraphRunDefaults | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const raw = value as Record<string, unknown>;
  const out: ParagraphRunDefaults = {};
  const finite = (item: unknown): item is number => typeof item === "number" && Number.isFinite(item);
  if (finite(raw.fontSize) && raw.fontSize >= 1 && raw.fontSize <= 4000) out.fontSize = raw.fontSize;
  if (typeof raw.bold === "boolean") out.bold = raw.bold;
  if (typeof raw.italic === "boolean") out.italic = raw.italic;
  if (typeof raw.color === "string" && raw.color) out.color = raw.color;
  if (typeof raw.fontFamily === "string" && raw.fontFamily) out.fontFamily = raw.fontFamily;
  if (finite(raw.letterSpacing)) out.letterSpacing = raw.letterSpacing;
  return Object.keys(out).length ? out : undefined;
}

/** ST_TextFontAlignType, in OOXML's own words. */
export type PresentationFontAlign = "auto" | "t" | "ctr" | "base" | "b";

export const PRESENTATION_FONT_ALIGNS: readonly PresentationFontAlign[] = ["auto", "t", "ctr", "base", "b"];

/** A stated Text alignment, or undefined for anything else. */
export function normalizeParagraphFontAlign(value: unknown): PresentationFontAlign | undefined {
  return typeof value === "string" && (PRESENTATION_FONT_ALIGNS as readonly string[]).includes(value)
    ? (value as PresentationFontAlign)
    : undefined;
}







export function paragraphEndLineFontSize(paragraph: Pick<SlideParagraph, "endFontSize" | "endScript"> | undefined): number | undefined {
  const size = paragraph?.endFontSize;
  if (typeof size !== "number" || !(size > 0)) return undefined;
  return paragraph?.endScript ? Math.max(1, Math.round((size * 2) / 3)) : size;
}

/** A row's stated end-mark size, in points (`a:endParaRPr@sz`, 1..4000). */
export function normalizeParagraphEndFontSize(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) && value >= 1 && value <= 4000
    ? value
    : undefined;
}

type CarriedParagraphKey =
  | "endFontSize"
  | "endScript"
  | "fontAlign"
  | "fontAlignInherited"
  | "runDefaults"
  | "endLatinFace"
  | "nativeKerning"
  | "nativeEastAsianLines"
  | "marR";

/** A row's own formatting the editor carries without drawing or authoring it. */
export function paragraphCarriedFields(
  record: unknown,
): Pick<SlideParagraph, CarriedParagraphKey> {
  if (!record || typeof record !== "object") return {};
  const raw = record as Record<string, unknown>;
  const out: Pick<SlideParagraph, CarriedParagraphKey> = {};
  // The row's own kerning and East Asian line rules go with it (readings row
  // 1009), as PowerPoint copies its `a:pPr` to both halves of an Enter.
  const nativeKerning = normalizeParagraphNativeKerning(raw.nativeKerning);
  if (nativeKerning !== undefined) out.nativeKerning = nativeKerning;
  const nativeEastAsianLines = normalizeParagraphEastAsianLines(raw.nativeEastAsianLines);
  if (nativeEastAsianLines !== undefined) out.nativeEastAsianLines = nativeEastAsianLines;
  // So does its right indent (readings row 1029).
  const marR = normalizeParagraphRightIndent(raw.marR);
  if (marR !== undefined) out.marR = marR;
  // The mark's Latin face goes with the row (readings row 899), as its mark does.
  const endLatinFace = normalizeParagraphEndLatinFace(raw.endLatinFace);
  if (endLatinFace !== undefined) out.endLatinFace = endLatinFace;
  // The row's own run defaults go with it through every rebuild, as its
  // `a:pPr` does in PowerPoint (readings row 885).
  const runDefaults = normalizeParagraphRunDefaults(raw.runDefaults);
  if (runDefaults !== undefined) out.runDefaults = runDefaults;
  const endFontSize = normalizeParagraphEndFontSize(raw.endFontSize);
  if (endFontSize !== undefined) out.endFontSize = endFontSize;
  if (raw.endScript === "super" || raw.endScript === "sub") out.endScript = raw.endScript;
  // A row's Text alignment goes with it through every rebuild, as its
  // direction does (the Paragraph dialog is what changes it).
  const fontAlign = normalizeParagraphFontAlign(raw.fontAlign);
  if (fontAlign !== undefined) out.fontAlign = fontAlign;
  const fontAlignInherited = normalizeParagraphFontAlign(raw.fontAlignInherited);
  if (fontAlignInherited !== undefined) out.fontAlignInherited = fontAlignInherited;
  return out;
}

/** A row's own marker formatting; see {@link SlideParagraph.bulletFormat}. */
export interface ParagraphBulletFormat {
  char?: string;
  typeface?: string;
  autoNum?: string;
  startAt?: number;
  color?: string;
  /** Thousandths of a percent of the text size, 25 000..400 000. */
  sizePct?: number;
  sizePts?: number;
}

const BULLET_FORMAT_SCHEME = /^[A-Za-z0-9]{1,32}$/;

/** Normalize a row's {@link ParagraphBulletFormat} for a row of kind `bullet`. */
export function normalizeParagraphBulletFormat(
  value: unknown,
  bullet: ParagraphBullet,
): ParagraphBulletFormat | undefined {
  if (bullet === "none" || !value || typeof value !== "object" || Array.isArray(value)) {
    return undefined;
  }
  const raw = value as Record<string, unknown>;
  const format: ParagraphBulletFormat = {};
  if (bullet === "bullet") {
    const char = boundedMarkerString(raw.char, 4);
    if (char !== undefined) format.char = char;
    const typeface = boundedMarkerString(raw.typeface, 128);
    if (typeface !== undefined && typeface.trim()) format.typeface = typeface;
  } else {
    if (typeof raw.autoNum === "string" && BULLET_FORMAT_SCHEME.test(raw.autoNum)) {
      format.autoNum = raw.autoNum;
    }
    const startAt = boundedMarkerInteger(raw.startAt, 1, 32_767);
    if (startAt !== undefined) format.startAt = startAt;
  }
  if (typeof raw.color === "string" && NATIVE_BULLET_HEX.test(raw.color)) {
    format.color = raw.color.toLowerCase();
  }
  const sizePct = boundedMarkerInteger(raw.sizePct, 25_000, 400_000);
  const sizePts = boundedMarkerInteger(raw.sizePts, 100, 400_000);
  if (sizePct !== undefined) format.sizePct = sizePct;
  else if (sizePts !== undefined) format.sizePts = sizePts;
  return format;
}

/**
 * The row's own marker choice, kept only for a row of the matching kind: a
 * glyph on an unbulleted row draws nothing and would be invisible state.
 */
export function paragraphMarkerFields(
  record: { bulletChar?: unknown; numberFormat?: unknown; bulletFormat?: unknown; markerHidden?: unknown },
  bullet: ParagraphBullet,
): Pick<SlideParagraph, "bulletChar" | "numberFormat" | "bulletFormat" | "markerHidden"> {
  const bulletFormat = normalizeParagraphBulletFormat(record.bulletFormat, bullet);
  const formatted = {
    ...(bulletFormat === undefined ? {} : { bulletFormat }),
    ...(record.markerHidden === true && bullet !== "none" ? { markerHidden: true as const } : {}),
  };
  if (bullet === "bullet") {
    const bulletChar = normalizeBulletChar(record.bulletChar);
    return bulletChar === undefined ? formatted : { bulletChar, ...formatted };
  }
  if (bullet === "number") {
    const numberFormat = normalizeNumberFormat(record.numberFormat);
    return numberFormat === undefined ? formatted : { numberFormat, ...formatted };
  }
  return {};
}

/** A row's stated paragraph direction, through Docs' own whitelist. */
export function normalizeParagraphDirection(value: unknown): "rtl" | undefined {
  return normalizeDir(value) === "rtl" ? "rtl" : undefined;
}

/**
 * The marker a row actually draws, where the bounded gallery cannot name it.
 *
 * The `bullet`/`bulletChar`/`numberFormat` model is a GALLERY: ten glyphs and
 * seven numbering formats, and that is what the editor lets anyone pick. A real
 * deck names whatever it likes, and more than half of the long tail measured
 * over 132 packages is a PRIVATE-USE code point — U+F06C, U+F02D, U+F096 — which
 * means nothing on its own. Real PowerPoint, exporting its own PDF of a probe
 * carrying exactly those, drew U+F06C as glyph 0x6C of an embedded
 * Wingdings-Regular and U+F02D as 0x2D of SymbolMT, with the paragraph TEXT
 * still in Calibri. ECMA-376 §21.1.2.4.4 states the same rule: the bullet font
 * "should only apply to the actual bullet character and not to the text within
 * the bullet". So the identity is `char` AND `typeface`, together.
 *
 * This is exactly the shape Docs already gives Word's "Define New Bullet"
 * (`bulletGlyph` + `bulletFont` on `DocumentBulletList`) — a marker outside the
 * gallery is drawn as itself, in its own font, and choosing a gallery style
 * clears it. Presentation reuses that reading rather than inventing a second.
 *
 * DISPLAY-ONLY. Nothing authors it, the canvas draws it, both writers put it
 * back unchanged, and it is dropped with its row on a free-text edit — the
 * disposal `nativeListLevels` already has, and for the same reason: it is read
 * POSITIONALLY against `paragraphs` (SHAPE-IMPORT-AUDIT.md §31.2).
 */
export interface NativeBulletMarker {
  /** `a:buChar@char`. Mutually exclusive with `autoNum` and `picture`. */
  char?: string;
  /**
   * `a:buBlip`: the picture a picture bullet draws, as an inline PNG/JPEG/GIF
   * (Bullets and Numbering > Picture). Drawn at the text's size, scaled by
   * `sizePct` like any marker; colour and font do not apply to it. Docs' own
   * picture bullet (`bulletPicture` on `DocumentBulletList`) states only that
   * a Word picture is missing, so there is no drawn picture model to share yet.
   */
  picture?: string;
  /** `a:buFont@typeface` — the font the glyph above belongs to. */
  typeface?: string;
  /** `a:buAutoNum@type`, an `ST_TextAutonumberScheme` the gallery cannot name. */
  autoNum?: string;
  /** `a:buAutoNum@startAt`, present only for a restart away from 1. */
  startAt?: number;
  /** `a:buClr` resolved to `#rrggbb`: the marker's own colour, not the text's. */
  color?: string;
  /** `a:buSzPct@val` in thousandths of a percent of the run size. */
  sizePct?: number;
  /** `a:buSzPts@val` in hundredths of a point. */
  sizePts?: number;
}

/** Bound for both indent quantities: DrawingML's own ST_TextMargin span. */
export const PARAGRAPH_INDENT_EMU_LIMIT = 51_206_400;

/** Normalize one authenticated EMU indent quantity, or drop it. */
export function normalizeParagraphIndentEmu(
  value: unknown,
  { signed }: { signed: boolean },
): number | undefined {
  if (typeof value !== "number" || !Number.isInteger(value)) return undefined;
  const minimum = signed ? -PARAGRAPH_INDENT_EMU_LIMIT : 0;
  return value >= minimum && value <= PARAGRAPH_INDENT_EMU_LIMIT
    ? value
    : undefined;
}

const NATIVE_BULLET_HEX = /^#[0-9a-fA-F]{6}$/;
/** The page validator's bound (`MAX_PRESENTATION_PICTURE_BULLET_CHARS`). */
const MAX_PICTURE_BULLET_CHARS = 180_000;

function boundedMarkerString(
  value: unknown,
  maximum: number,
): string | undefined {
  return typeof value === "string" && value.length >= 1 && value.length <= maximum
    ? value
    : undefined;
}

function boundedMarkerInteger(
  value: unknown,
  minimum: number,
  maximum: number,
): number | undefined {
  return typeof value === "number" &&
    Number.isInteger(value) &&
    value >= minimum &&
    value <= maximum
    ? value
    : undefined;
}

/** Normalize one persisted {@link NativeBulletMarker}, or drop it. */
export function normalizeNativeBulletMarker(
  value: unknown,
): NativeBulletMarker | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const raw = value as Record<string, unknown>;
  const char = boundedMarkerString(raw.char, 4);
  const autoNum = boundedMarkerString(raw.autoNum, 32);
  const picture =
    typeof raw.picture === "string" &&
    raw.picture.length <= MAX_PICTURE_BULLET_CHARS &&
    /^data:image\/(?:png|jpeg|gif);base64,/i.test(raw.picture) &&
    isValidInlineImageSource(raw.picture)
      ? raw.picture
      : undefined;
  // Exactly one marker identity, the same either/or the writers hold.
  if (
    [char, autoNum, picture].filter((value) => value !== undefined).length !== 1
  ) {
    return undefined;
  }
  const marker: NativeBulletMarker =
    char !== undefined ? { char } : autoNum !== undefined ? { autoNum } : { picture };
  if (picture !== undefined) {
    const pictureSizePct = boundedMarkerInteger(raw.sizePct, 25_000, 400_000);
    const pictureSizePts = boundedMarkerInteger(raw.sizePts, 100, 400_000);
    if (pictureSizePct !== undefined) marker.sizePct = pictureSizePct;
    else if (pictureSizePts !== undefined) marker.sizePts = pictureSizePts;
    return marker;
  }
  const typeface = boundedMarkerString(raw.typeface, 128);
  if (typeface !== undefined) marker.typeface = typeface;
  const startAt = boundedMarkerInteger(raw.startAt, 1, 32_767);
  if (autoNum !== undefined && startAt !== undefined) marker.startAt = startAt;
  if (typeof raw.color === "string" && NATIVE_BULLET_HEX.test(raw.color)) {
    marker.color = raw.color.toLowerCase();
  }
  const sizePct = boundedMarkerInteger(raw.sizePct, 25_000, 400_000);
  const sizePts = boundedMarkerInteger(raw.sizePts, 100, 400_000);
  if (sizePct !== undefined) marker.sizePct = sizePct;
  else if (sizePts !== undefined) marker.sizePts = sizePts;
  return marker;
}

/**
 * Normalize a persisted `nativeBullet` receipt against the rows it describes.
 *
 * Read POSITIONALLY, so a length that no longer matches the rows means the rows
 * were rewritten under it, and the only safe reading of a stale positional
 * receipt is none at all. An all-`null` array says nothing and is dropped too.
 */
export function normalizeNativeBulletReceipt(
  value: unknown,
  rowCount: number,
): (NativeBulletMarker | null)[] | undefined {
  if (!Array.isArray(value) || value.length !== rowCount || rowCount === 0) {
    return undefined;
  }
  const entries = value.map(
    (entry) => normalizeNativeBulletMarker(entry) ?? null,
  );
  return entries.some((entry) => entry !== null) ? entries : undefined;
}

/**
 * Drop every receipt a block reads POSITIONALLY against its rows.
 *
 * `nativeListLevels` and `nativeBullet` are both aligned to `paragraphs` by
 * index. Leaving either behind when the rows go hands the save a receipt that
 * aligns with nothing, and the page validator rejects exactly that — the edit
 * would fail closed on a label the user simply retyped. One helper so the next
 * receipt is free to add and no disposal site can be forgotten.
 */
/**
 * Hand a block's `nativeBullet` receipt over to its rows: each row whose
 * receipt entry states a character or scheme, a colour or a size takes it as
 * its own `bulletFormat`, which a text edit carries with the row. Only a
 * picture bullet stays a receipt. The save folds the two the same way, so
 * this hand-over is no change to the package (PPT-A039/A040).
 */
export function materializeNativeBulletReceipt<
  T extends { paragraphs?: unknown; nativeBullet?: unknown },
>(block: T): T {
  const rows = Array.isArray(block.paragraphs) ? block.paragraphs : undefined;
  const receipt = rows
    ? normalizeNativeBulletReceipt(block.nativeBullet, rows.length)
    : undefined;
  if (!rows || !receipt) return block;
  const paragraphs = rows.map((row, index) => {
    const entry = receipt[index];
    if (
      !entry ||
      entry.picture !== undefined ||
      !row ||
      typeof row !== "object" ||
      "bulletFormat" in row
    ) {
      return row;
    }
    const bullet = normalizeBullet((row as { bullet?: unknown }).bullet);
    const bulletFormat = normalizeParagraphBulletFormat(entry, bullet);
    return bulletFormat === undefined ? row : { ...row, bulletFormat };
  });
  const pictures = receipt.map((entry) => (entry?.picture !== undefined ? entry : null));
  const next = { ...block, paragraphs } as T & { nativeBullet?: unknown };
  if (pictures.some((entry) => entry !== null)) next.nativeBullet = pictures;
  else delete next.nativeBullet;
  return next;
}

export function dropNativeParagraphReceipts(block: object): void {
  const target = block as {
    nativeListLevels?: unknown;
    nativeBullet?: unknown;
  };
  delete target.nativeListLevels;
  delete target.nativeBullet;
}











export function drawnBulletChar(char: string): string {
  return char.length > 1 && char.startsWith("\u2022") ? "\u2022" : char;
}

/**
 * How to draw one native marker: the text to paint and the font to paint it in.
 *
 * A private-use code point renders as its symbol font's glyph only where that
 * font exists. Where it does not, the browser paints a notdef box, so the
 * mapped Unicode stand-in below is used instead — the same trade Word makes
 * when it substitutes a face. `fontAvailable` is injected rather than probed
 * here so this module stays DOM-free and the caller can memoize one
 * `document.fonts.check` per typeface.
 */
export function nativeBulletGlyph(
  marker: Pick<NativeBulletMarker, "char" | "typeface">,
  fontAvailable?: (typeface: string, text: string) => boolean,
): { text: string; fontFamily?: string } | undefined {
  if (marker.char === undefined) return undefined;
  const char = drawnBulletChar(marker.char);
  const typeface = marker.typeface;
  if (typeface === undefined) return { text: char };
  // Symbol's private-use bullets have a known Unicode stand-in, and the face
  // a browser finds under that name is rarely Office's: the Mac's own Symbol
  
  
  const standIn = SYMBOL_FONT_UNICODE_FALLBACK[`${typeface}:${char}`];
  if (standIn !== undefined && typeface.trim().toLowerCase() === "symbol") return { text: standIn };
  const drawable = fontAvailable === undefined || fontAvailable(typeface, char);
  if (drawable) return { text: char, fontFamily: typeface };
  const fallback = SYMBOL_FONT_UNICODE_FALLBACK[`${typeface}:${char}`];
  if (fallback !== undefined) return { text: fallback };
  // A symbol font's private-use code point with no known stand-in: a plain
  // bullet, never the notdef box the browser paints for it.
  return { text: isSymbolPrivateUse(char) ? "\u2022" : char };
}

/**
 * Whether `nativeBulletGlyph` would GUESS the marker where its face is not
 * drawn: a symbol-encoded face's character (Wingdings, Symbol, Webdings, or a
 * private-use code point) with no stand-in measured against PowerPoint. The
 * text engine's table cells refuse such a marker (readings row 1036) rather
 * than draw a plain bullet or the character's Latin letter.
 */
export function nativeBulletGlyphIsGuess(
  marker: Pick<NativeBulletMarker, "char" | "typeface">,
  glyph: { text: string; fontFamily?: string } | undefined,
): boolean {
  if (marker.char === undefined || marker.typeface === undefined || glyph?.fontFamily !== undefined) return false;
  const char = drawnBulletChar(marker.char);
  if (SYMBOL_FONT_UNICODE_FALLBACK[`${marker.typeface}:${char}`] !== undefined) return false;
  return isSymbolPrivateUse(char) || SYMBOL_ENCODED_FACE.test(marker.typeface.trim());
}

/** Faces whose characters are symbols stored at Latin code points. */
const SYMBOL_ENCODED_FACE = /^(wingdings( 2| 3)?|webdings|symbol|marlett|mt extra|monotype sorts|zapf ?dingbats|starsymbol|opensymbol)$/i;

/** One character in the range a symbol font's glyphs are stored shifted into (U+F020-U+F0FF). */
function isSymbolPrivateUse(text: string): boolean {
  const code = text.length === 1 ? text.charCodeAt(0) : -1;
  return code >= 0xf020 && code <= 0xf0ff;
}

/**
 * Unicode stand-ins for the symbol-font markers this corpus actually carries.
 *
 * Deliberately NOT a general Wingdings table. Each entry is a glyph measured in
 * a real PowerPoint PDF export of the package it came from, so a stand-in is
 * only offered where the shape it replaces is known.
 */
const SYMBOL_FONT_UNICODE_FALLBACK: Readonly<Record<string, string>> = {
  "Wingdings:": "■",
  "Wingdings:": "◆",
  "Wingdings:": "➢",
  "Wingdings:Ø": "➢",
  "Wingdings:ü": "✓",
  // PowerPoint's own Bullets gallery (PPT-A039): Wingdings 0x6C, 0x6E,
  // 0x71, 0x75, 0x76 per the Unicode Consortium's Wingdings mapping.
  "Wingdings:l": "●",
  "Wingdings:n": "■",
  "Wingdings:q": "❑",
  "Wingdings:u": "◆",
  "Wingdings:v": "❖",
  "Wingdings:§": "■",
  "Wingdings 2:": "▪",
  "Wingdings 2:": "▫",
  "Wingdings 2:": "●",
  "Wingdings 3:": "▶",
  // Office's Facet, Ion and Wisp themes (Wingdings 3 0x70-0x77, named
  // triangleup ... triangleopenrt in the font itself).
  "Wingdings 3:\uf070": "▲",
  "Wingdings 3:\uf071": "▼",
  "Wingdings 3:\uf072": "△",
  "Wingdings 3:\uf073": "▽",
  "Wingdings 3:\uf074": "◀",
  "Wingdings 3:\uf075": "▶",
  "Wingdings 3:\uf076": "◁",
  "Wingdings 3:\uf077": "▷",
  "Symbol:": "•",
  "Symbol:": "•",
  // The Office 2003 templates' `buFont Symbol` + `buChar ·`: the Symbol
  // font's code 0xB7 is its bullet, as PowerPoint draws it (symbol-bullet
  
  
  "Symbol:·": "•",
  // The same encoding's 0xA7 and 0xD8 (symbol-bullet probe R3, R4).
  "Symbol:§": "♣",
  "Symbol:Ø": "¬",
  "StarSymbol:": "■",
  "StarSymbol:": "•",
};

/** True when this row states an indentation of its own. */
export function paragraphStatesIndent(
  paragraph: Pick<SlideParagraph, "marL" | "indent"> | undefined,
): boolean {
  return (
    paragraph !== undefined &&
    (paragraph.marL !== undefined || paragraph.indent !== undefined)
  );
}

/**
 * A text block's whole text, as the server bounds it
 * (`MAX_PRESENTATION_TEXT_BLOCK_TEXT_LENGTH` in routers/notebooks.py).
 */
export const MAX_TEXT_BLOCK_TEXT_LENGTH = 20_000;
/**
 * Per-paragraph text cap + max paragraphs per block (mirrors the schema). A
 * PowerPoint paragraph has no length bound of its own, so a text box's row is
 * as long as its whole text (`MAX_PRESENTATION_TEXT_BLOCK_LINE_LENGTH`); the
 * 2,000 it held let a user type a line the save then refused (PPT-B392). A
 * shape label's rows keep `MAX_SHAPE_TEXT_LENGTH`.
 */
export const MAX_PARAGRAPH_TEXT_LENGTH = MAX_TEXT_BLOCK_TEXT_LENGTH;
export const MAX_PARAGRAPHS_PER_BLOCK = 200;

const BULLET_SET: ReadonlySet<string> = new Set(PARAGRAPH_BULLETS);
const ALIGN_SET: ReadonlySet<string> = new Set(PARAGRAPH_ALIGNS);









export function foldParagraphRecordsPastLimit(value: readonly unknown[]): readonly unknown[] {
  if (value.length <= MAX_PARAGRAPHS_PER_BLOCK) return value;
  const kept = value.slice(0, MAX_PARAGRAPHS_PER_BLOCK - 1);
  const rest = value
    .slice(MAX_PARAGRAPHS_PER_BLOCK - 1)
    .map((entry) => (entry && typeof entry === "object" ? (entry as Record<string, unknown>) : {}));
  const textOf = (row: Record<string, unknown>) => (typeof row.text === "string" ? row.text : "");
  const withRuns = rest.some((row) => Array.isArray(row.runs) && row.runs.length > 0);
  const runs: unknown[] = [];
  if (withRuns) {
    rest.forEach((row, index) => {
      if (index > 0) {
        const previous = runs[runs.length - 1] as { style?: unknown } | undefined;
        runs.push(previous?.style ? { text: "\n", style: previous.style } : { text: "\n" });
      }
      if (Array.isArray(row.runs) && row.runs.length > 0) runs.push(...row.runs);
      else if (textOf(row)) runs.push({ text: textOf(row) });
    });
  }
  const folded = { ...rest[0], text: rest.map(textOf).join("\n"), ...(withRuns ? { runs } : {}) };
  return [...kept, folded];
}

/** Normalize an arbitrary value into a valid {@link ParagraphBullet}. */
export function normalizeBullet(value: unknown): ParagraphBullet {
  return typeof value === "string" && BULLET_SET.has(value)
    ? (value as ParagraphBullet)
    : "none";
}

/** Normalize an authenticated paragraph-local alignment, or drop it. */
export function normalizeParagraphAlign(
  value: unknown,
): ParagraphAlign | undefined {
  return typeof value === "string" && ALIGN_SET.has(value)
    ? (value as ParagraphAlign)
    : undefined;
}

/**
 * Normalize an arbitrary persisted value into a bounded `SlideParagraph[]`, or
 * `undefined` when the value is not a usable paragraphs array. A non-array, an
 * empty array, or an array whose every entry is unusable all collapse to
 * `undefined` so a block without real list structure stays a flat-text block
 * (byte-stable). Each surviving paragraph has a string `text` (clamped), an
 * integer `level` in range, and a valid `bullet`.
 */
export function normalizeParagraphs(value: unknown): SlideParagraph[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const paragraphs: SlideParagraph[] = [];
  for (const entry of foldParagraphRecordsPastLimit(value)) {
    if (paragraphs.length >= MAX_PARAGRAPHS_PER_BLOCK) break;
    if (!entry || typeof entry !== "object") continue;
    const record = entry as Record<string, unknown>;
    const text = typeof record.text === "string" ? record.text.slice(0, MAX_PARAGRAPH_TEXT_LENGTH) : "";
    const paragraph: SlideParagraph = {
      text,
      level: clampLevel(record.level),
      bullet: normalizeBullet(record.bullet),
    };
    const align = normalizeParagraphAlign(record.align);
    if (align !== undefined) paragraph.align = align;
    const marL = normalizeParagraphIndentEmu(record.marL, { signed: false });
    if (marL !== undefined) paragraph.marL = marL;
    const indent = normalizeParagraphIndentEmu(record.indent, { signed: true });
    if (indent !== undefined) paragraph.indent = indent;
    const dir = normalizeParagraphDirection(record.dir);
    if (dir !== undefined) paragraph.dir = dir;
    Object.assign(paragraph, paragraphMarkerFields(record, paragraph.bullet));
    Object.assign(paragraph, paragraphCarriedFields(record));
    paragraphs.push(paragraph);
  }
  if (paragraphs.length === 0) return undefined;
  // A paragraphs array with NO bullets/numbers AND no indent is visually
  // indistinguishable from a flat-text block (the render adds a marker only for
  // a bullet/number and indents only for level > 0). Collapse it to `undefined`
  // so a flat block that was toggled into a list and back stays byte-for-byte
  // identical to its origin (backward-compat / no phantom `paragraphs` field).
  // A `none` paragraph that is still INDENTED (level > 0) is a real, visible
  // state and is preserved.
  return isFlatEquivalentParagraphs(paragraphs) ? undefined : paragraphs;
}

/**
 * True when these paragraphs carry no list structure the flat-text form cannot
 * represent: every paragraph is unbulleted, at the base indent level and
 * states no indentation of its own. Such a block renders identically to its
 * plain `text`, so it should not persist a `paragraphs` field.
 */
export function isFlatEquivalentParagraphs(
  paragraphs: readonly SlideParagraph[] | undefined,
): boolean {
  return (
    !!paragraphs &&
    paragraphs.length > 0 &&
    paragraphs.every(
      (p) =>
        p.bullet === "none" &&
        p.level === 0 &&
        p.align === undefined &&
        p.dir === undefined &&
        p.fontAlign === undefined &&
        p.fontAlignInherited === undefined &&
        // A row's own kerning and East Asian line rules (readings row 1009)
        // are the row's: a flat body keeps one set for all of its rows.
        p.nativeKerning === undefined &&
        p.nativeEastAsianLines === undefined &&
        // So is its right indent (readings row 1029): a flat body has none.
        p.marR === undefined &&
        !paragraphStatesIndent(p),
    )
  );
}

/**
 * Split a flat block string into one paragraph per line, preserving each line's
 * current bullet/level when a positionally-aligned previous paragraph exists.
 * This is how the single-textarea editor maps newline-separated text onto the
 * structured array: line N inherits the style of old paragraph N (so toggling a
 * list and then editing the text keeps the per-line bullets), and a brand-new
 * line defaults to `{ level: 0, bullet: "none" }`.
 */
export function paragraphsFromText(
  text: string,
  previous?: readonly SlideParagraph[],
): SlideParagraph[] {
  const lines = String(text ?? "").split("\n");
  return lines.map((line, index) => {
    const prior = previous?.[index];
    const paragraph: SlideParagraph = {
      text: line.slice(0, MAX_PARAGRAPH_TEXT_LENGTH),
      level: prior ? clampLevel(prior.level) : MIN_PARAGRAPH_LEVEL,
      bullet: prior ? normalizeBullet(prior.bullet) : "none",
    };
    const align = normalizeParagraphAlign(prior?.align);
    if (align !== undefined) paragraph.align = align;
    // Re-typed text keeps its row's direction, as it keeps its alignment.
    const dir = normalizeParagraphDirection(prior?.dir);
    if (dir !== undefined) paragraph.dir = dir;
    // Line N inherits row N's indentation for the same reason it inherits its
    // bullet and alignment: re-typing the text of an indented line is not a
    // request to move it back to the margin. A brand-new line has no prior row
    // and therefore no indent of its own, which is the inherited reading.
    const marL = normalizeParagraphIndentEmu(prior?.marL, { signed: false });
    if (marL !== undefined) paragraph.marL = marL;
    const indent = normalizeParagraphIndentEmu(prior?.indent, { signed: true });
    if (indent !== undefined) paragraph.indent = indent;
    if (prior) Object.assign(paragraph, paragraphMarkerFields(prior, paragraph.bullet));
    if (prior) Object.assign(paragraph, paragraphCarriedFields(prior));
    return paragraph;
  });
}

/** Join paragraphs back into the flat block `text` (newline per paragraph). */
export function textFromParagraphs(paragraphs: readonly SlideParagraph[] | undefined): string {
  if (!paragraphs || paragraphs.length === 0) return "";
  return paragraphs.map((paragraph) => paragraph.text).join("\n");
}

/**
 * Toggle a bullet style across EVERY paragraph in the block. Mirrors a list
 * button: pressing "bullet" turns the whole block into a bulleted list; pressing
 * it again (when already all that style) turns the list off (`none`). A mixed
 * block becomes uniformly the requested style.
 */
export function toggleBullet(
  paragraphs: readonly SlideParagraph[],
  bullet: Exclude<ParagraphBullet, "none">,
): SlideParagraph[] {
  const allAlready = paragraphs.length > 0 && paragraphs.every((p) => p.bullet === bullet);
  const next: ParagraphBullet = allAlready ? "none" : bullet;
  return paragraphs.map((paragraph) => ({ ...paragraph, bullet: next }));
}

/**
 * Shift the indent level of every paragraph by `delta` (e.g. +1 = indent, −1 =
 * outdent), clamped to the valid range. Returns a new array; paragraphs already
 * at a bound stay put.
 */
export function setLevel(paragraphs: readonly SlideParagraph[], delta: number): SlideParagraph[] {
  return paragraphs.map((paragraph) => ({
    ...paragraph,
    level: clampLevel(paragraph.level + delta),
  }));
}

/**
 * Map a textarea caret OFFSET (0-based char index into the flat `text`, where
 * paragraphs are joined by a single `\n`) to the index of the paragraph that
 * caret sits in. The textarea edits the joined string, so the structured
 * `paragraphs` array is addressed by counting `\n`s before the caret. A caret
 * exactly ON a newline is treated as the END of the paragraph BEFORE it (the
 * caret renders at the end of that visual line). Clamped to a valid index for a
 * non-empty array; returns 0 for an empty array. `paragraphs` is read only for
 * its length — the offset → line mapping uses each paragraph's text length.
 *
 * Pure; the inverse of `textFromParagraphs` for caret positioning. Used by the
 * editor to map a caret into a single paragraph for per-line operations
 * (Tab / Shift+Tab indent of just the current line).
 */
export function paragraphIndexAtOffset(
  paragraphs: readonly SlideParagraph[],
  offset: number,
): number {
  if (paragraphs.length === 0) return 0;
  const last = paragraphs.length - 1;
  if (!Number.isFinite(offset) || offset <= 0) return 0;
  const target = Math.round(offset);
  // Walk paragraphs, advancing `consumed` past each paragraph's text plus the
  // single `\n` separator that follows it. The caret belongs to the FIRST
  // paragraph whose end (text end, BEFORE its trailing newline) is >= the
  // offset — i.e. an offset that lands on a separator stays with the line it
  // terminates rather than jumping to the next line.
  let consumed = 0;
  for (let index = 0; index < paragraphs.length; index += 1) {
    const textEnd = consumed + paragraphs[index].text.length;
    if (target <= textEnd) return index;
    // Skip the text + the `\n` joiner to the start of the next paragraph.
    consumed = textEnd + 1;
  }
  return last;
}

/**
 * Shift the indent level of a SINGLE paragraph (by index) by `delta`, clamped to
 * the valid range. Returns a NEW array with only that paragraph replaced; every
 * other paragraph keeps its identity. An out-of-range index, a zero delta, or a
 * paragraph already at the relevant bound returns the SAME array reference (no
 * change), so a no-op never churns the model. Unlike {@link setLevel} (which
 * shifts every paragraph) this targets just the line the caret is on — Tab /
 * Shift+Tab demote/promote in a list.
 */
export function adjustParagraphLevelAt(
  paragraphs: readonly SlideParagraph[],
  index: number,
  delta: number,
  /**
   * The row's character formatting at its new level
   * (`presentationListLevelRestyle`: an imported body's runs re-inherit as
   * PowerPoint's do); null keeps the row where it is, as PowerPoint refuses
   * a move that would pass its largest type size. Absent: formatting stays.
   */
  restyle?: <P extends SlideParagraph>(paragraph: P, from: number, to: number) => P | null,
): SlideParagraph[] {
  const arr = paragraphs as SlideParagraph[];
  if (index < 0 || index >= paragraphs.length || delta === 0) return arr;
  const original = paragraphs[index];
  const nextLevel = clampLevel(original.level + delta);
  if (nextLevel === original.level) return arr;
  const current = restyle ? restyle(original, original.level, nextLevel) : original;
  if (current === null) return arr;
  // A row that states its own `a:pPr@marL` is drawn at that margin, not at
  // its level's, so the level alone would not move it: an imported text box
  // whose rows carry their exact indentation would ignore Increase / Decrease
  // Indent. The margin moves with the level, one half inch a step (the
  // ladder's rung and PowerPoint's text-box level step), within the schema's
  // span.
  const moved =
    current.marL === undefined
      ? {}
      : {
          marL: Math.min(
            PARAGRAPH_INDENT_EMU_LIMIT,
            Math.max(
              0,
              current.marL + (nextLevel - original.level) * PARAGRAPH_LEVEL_STEP_EMU,
            ),
          ),
        };
  return paragraphs.map((paragraph, position) =>
    position === index ? { ...current, level: nextLevel, ...moved } : paragraph,
  );
}

/** One outline level's step for a row that states its own margin (EMU). */
const PARAGRAPH_LEVEL_STEP_EMU = 457_200;

/**
 * The numbering engine lives in the host-neutral `office-paragraph-render`.
 *
 * `renderParagraphs` and the bounded marker vocabulary around it are not the
 * deck's: the sheet's shape text (`office-text-layout.ts`) and the document's
 * floating boxes call the same function, because `a:pPr` means the same thing
 * in `p:sp`, `xdr:sp` and `wps:wsp`. The audit recorded the deck-shaped NAME as
 * debt twice (SHAPE-IMPORT-AUDIT.md §44.7); §47.4 moved it. Everything is
 * re-exported here so the deck's own import sites are untouched.
 */
export {
  MIN_PARAGRAPH_LEVEL,
  MAX_PARAGRAPH_LEVEL,
  BLOCK_BULLET_CHARS,
  BULLET_CHARS,
  NUMBER_FORMATS,
  PRESENTATION_WEB_BULLET_STYLES,
  autoNumberLabel,
  bulletGlyphForLevel,
  clampLevel,
  nativeAutoNumberLabel,
  normalizeBlockBulletChar,
  normalizeBulletChar,
  normalizeNumberFormat,
  numberFormatForLevel,
  numberLabelForFormat,
  renderParagraphs,
  type BlockBulletChar,
  type BulletChar,
  type MarkerOptions,
  type NumberFormat,
  type RenderedParagraph,
} from "./office-paragraph-render";

/** True when at least one paragraph carries a visible bullet/number marker. */
export function hasAnyBullet(paragraphs: readonly SlideParagraph[] | undefined): boolean {
  return !!paragraphs?.some((paragraph) => paragraph.bullet !== "none");
}
