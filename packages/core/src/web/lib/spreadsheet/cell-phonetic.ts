/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ptToPx } from "../office-units";
/**
 * Phonetic guides (furigana) on a cell — Excel's Home ▸ Font ▸ Phonetic Guide.
 *
 * SpreadsheetML keeps them on the cell's string (ECMA-376 §18.4: `rPh` runs
 * over the base text, one `phoneticPr` with the font, character type and
 * alignment) and keeps "Show Phonetic Field" on the cell itself (`c@ph`). The
 * model mirrors that as one optional cell-level field, `cell.phonetic`:
 *
 *   runs       the furigana, each over [start, end) of `cell.value` in UTF-16
 *              code units (`rPh@sb` / `rPh@eb`, end exclusive)
 *   type       `phoneticPr@type` when it is not fullwidthKatakana (the default)
 *   alignment  `phoneticPr@alignment` when it is not left (the default)
 *   font       the phonetic font when it is not the default phonetic font
 *   visible    `c@ph`: Show Phonetic Field is on for this cell
 *
 * A plain cell carries no field at all, and a field never carries a default:
 * `{ value }` is a cell without furigana settings, byte-stable. The text and
 * the settings describe the cell's TEXT, so a number, a formula or an empty
 * cell keeps only `visible`.
 *
 * Import reads the package on the server (apps/api/app/services/file_adapters/
 * xlsx/phonetic.py); this module edits, draws and evaluates the field.
 */

export const SPREADSHEET_PHONETIC_TYPES = [
  "fullwidthKatakana",
  "halfwidthKatakana",
  "hiragana",
  "noConversion",
] as const;
export type SpreadsheetPhoneticType = (typeof SPREADSHEET_PHONETIC_TYPES)[number];

export const SPREADSHEET_PHONETIC_ALIGNMENTS = [
  "noControl",
  "left",
  "center",
  "distributed",
] as const;
export type SpreadsheetPhoneticAlignment =
  (typeof SPREADSHEET_PHONETIC_ALIGNMENTS)[number];

/** ECMA-376 CT_PhoneticPr defaults. */
export const DEFAULT_PHONETIC_TYPE = "fullwidthKatakana" satisfies SpreadsheetPhoneticType;
export const DEFAULT_PHONETIC_ALIGNMENT = "left" satisfies SpreadsheetPhoneticAlignment;
/**
 * The default phonetic font is the workbook's own face at 6 pt — what
 * Japanese Excel writes beside its 11 pt Normal font (`<phoneticPr
 * fontId="1"/>` naming `<sz val="6"/>`). The API importer and both writers
 * use the same default (app/services/file_adapters/xlsx/phonetic.py).
 */
export const DEFAULT_PHONETIC_FONT_PT = 6;

/** Bounds: a cell holds at most 32,767 characters, and each run covers one. */
export const MAX_PHONETIC_RUNS = 2_048;
export const MAX_PHONETIC_RUN_TEXT = 1_024;
const MAX_PHONETIC_TEXT_TOTAL = 32_767;
const MAX_PHONETIC_FONT_FAMILY = 200;

export interface SpreadsheetPhoneticRun {
  start: number;
  end: number;
  text: string;
}

export interface SpreadsheetPhoneticFont {
  /** The face name (or the font stack the grid's own font picker writes). */
  family: string;
  /** Points. */
  size: number;
  bold?: true;
  italic?: true;
  underline?: "single" | "double";
  strikethrough?: true;
  /** `#rrggbb`; absent = automatic. */
  color?: string;
}

export interface SpreadsheetCellPhonetic {
  runs?: SpreadsheetPhoneticRun[];
  type?: Exclude<SpreadsheetPhoneticType, "fullwidthKatakana">;
  alignment?: Exclude<SpreadsheetPhoneticAlignment, "left">;
  font?: SpreadsheetPhoneticFont;
  visible?: true;
}

/** The fields a Phonetic Settings OK writes (Settings tab + Font tab). */
export interface SpreadsheetPhoneticSettings {
  type: SpreadsheetPhoneticType;
  alignment: SpreadsheetPhoneticAlignment;
  /** null = the default phonetic font. */
  font: SpreadsheetPhoneticFont | null;
}

interface PhoneticCellLike {
  value?: unknown;
  kind?: unknown;
  formula?: unknown;
  phonetic?: unknown;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isPhoneticType(value: unknown): value is SpreadsheetPhoneticType {
  return (SPREADSHEET_PHONETIC_TYPES as readonly unknown[]).includes(value);
}

function isPhoneticAlignment(
  value: unknown,
): value is SpreadsheetPhoneticAlignment {
  return (SPREADSHEET_PHONETIC_ALIGNMENTS as readonly unknown[]).includes(value);
}

/** A valid phonetic font, or undefined. */
export function normalizePhoneticFont(
  raw: unknown,
): SpreadsheetPhoneticFont | undefined {
  if (!isRecord(raw)) return undefined;
  const family = typeof raw.family === "string" ? raw.family.trim() : "";
  const size = typeof raw.size === "number" ? raw.size : Number.NaN;
  if (
    !family ||
    family.length > MAX_PHONETIC_FONT_FAMILY ||
    !Number.isFinite(size) ||
    size < 1 ||
    size > 409
  ) {
    return undefined;
  }
  const font: SpreadsheetPhoneticFont = {
    family,
    size: Math.round(size * 100) / 100,
  };
  if (raw.bold === true) font.bold = true;
  if (raw.italic === true) font.italic = true;
  if (raw.underline === "single" || raw.underline === "double") {
    font.underline = raw.underline;
  }
  if (raw.strikethrough === true) font.strikethrough = true;
  if (typeof raw.color === "string" && /^#[0-9a-fA-F]{6}$/.test(raw.color)) {
    font.color = raw.color.toLowerCase();
  }
  return font;
}

export function phoneticFontsEqual(
  left: SpreadsheetPhoneticFont | undefined | null,
  right: SpreadsheetPhoneticFont | undefined | null,
): boolean {
  if (!left || !right) return !left && !right;
  return (
    left.family === right.family &&
    left.size === right.size &&
    left.bold === right.bold &&
    left.italic === right.italic &&
    left.underline === right.underline &&
    left.strikethrough === right.strikethrough &&
    left.color === right.color
  );
}

/** UTF-16 runs over `value`: sorted, inside it, never overlapping. */
export function normalizePhoneticRuns(
  raw: unknown,
  value: string,
): SpreadsheetPhoneticRun[] | undefined {
  if (!Array.isArray(raw) || raw.length === 0) return undefined;
  const candidates: SpreadsheetPhoneticRun[] = [];
  for (const item of raw.slice(0, MAX_PHONETIC_RUNS)) {
    if (!isRecord(item)) continue;
    const { start, end, text } = item;
    if (
      typeof start !== "number" ||
      typeof end !== "number" ||
      !Number.isInteger(start) ||
      !Number.isInteger(end) ||
      start < 0 ||
      end < start ||
      end > value.length ||
      typeof text !== "string" ||
      text.length === 0 ||
      text.length > MAX_PHONETIC_RUN_TEXT
    ) {
      continue;
    }
    candidates.push({ start, end, text });
  }
  candidates.sort((a, b) => a.start - b.start || a.end - b.end);
  const runs: SpreadsheetPhoneticRun[] = [];
  let reached = 0;
  let total = 0;
  for (const run of candidates) {
    if (runs.length > 0 && run.start < reached) continue;
    total += run.text.length;
    if (total > MAX_PHONETIC_TEXT_TOTAL) break;
    runs.push(run);
    reached = run.end;
  }
  return runs.length > 0 ? runs : undefined;
}

/**
 * The field as stored, normalized for a cell whose text is `value`, or
 * undefined when nothing is left. `textual` is false for a number or a
 * formula: only Show Phonetic Field belongs to such a cell.
 */
export function normalizeCellPhonetic(
  raw: unknown,
  value: string,
  textual = true,
): SpreadsheetCellPhonetic | undefined {
  if (!isRecord(raw)) return undefined;
  const next: SpreadsheetCellPhonetic = {};
  if (textual && value !== "") {
    const runs = normalizePhoneticRuns(raw.runs, value);
    if (runs) next.runs = runs;
    if (isPhoneticType(raw.type) && raw.type !== DEFAULT_PHONETIC_TYPE) {
      next.type = raw.type;
    }
    if (
      isPhoneticAlignment(raw.alignment) &&
      raw.alignment !== DEFAULT_PHONETIC_ALIGNMENT
    ) {
      next.alignment = raw.alignment;
    }
    const font = normalizePhoneticFont(raw.font);
    if (font) next.font = font;
  }
  if (raw.visible === true) next.visible = true;
  return Object.keys(next).length > 0 ? next : undefined;
}

function cellIsTextual(cell: PhoneticCellLike): boolean {
  return (
    typeof cell.value === "string" &&
    cell.kind !== "number" &&
    !(typeof cell.formula === "string" && cell.formula.trim() !== "")
  );
}

/** The cell's phonetic field, normalized against its own text. */
export function readCellPhonetic(
  cell: PhoneticCellLike | null | undefined,
): SpreadsheetCellPhonetic | undefined {
  if (!cell) return undefined;
  const value = typeof cell.value === "string" ? cell.value : "";
  return normalizeCellPhonetic(cell.phonetic, value, cellIsTextual(cell));
}

/** Show Phonetic Field is on for the cell. */
export function cellPhoneticVisible(
  cell: PhoneticCellLike | null | undefined,
): boolean {
  return isRecord(cell?.phonetic) && cell.phonetic.visible === true;
}

/** The settings the Phonetic Settings dialog shows for a cell. */
export function cellPhoneticSettings(
  cell: PhoneticCellLike | null | undefined,
): SpreadsheetPhoneticSettings {
  const phonetic = readCellPhonetic(cell);
  return {
    type: phonetic?.type ?? DEFAULT_PHONETIC_TYPE,
    alignment: phonetic?.alignment ?? DEFAULT_PHONETIC_ALIGNMENT,
    font: phonetic?.font ?? null,
  };
}

/**
 * The phonetic runs an in-place edit (F2) keeps: a run whose base characters
 * lie wholly inside the unchanged start or the unchanged end of the text keeps
 * its reading (shifted with that end); any run the edit touched goes. Excel's
 * own cell editor keeps furigana the same way, and drops them for a new entry.
 */
export function phoneticRunsForEditedText(
  runs: readonly SpreadsheetPhoneticRun[] | undefined,
  oldValue: string,
  value: string,
): SpreadsheetPhoneticRun[] | undefined {
  if (!runs || runs.length === 0) return undefined;
  if (oldValue === value) return [...runs];
  let prefix = 0;
  while (
    prefix < oldValue.length &&
    prefix < value.length &&
    oldValue.charCodeAt(prefix) === value.charCodeAt(prefix)
  ) {
    prefix += 1;
  }
  let suffix = 0;
  while (
    suffix < oldValue.length - prefix &&
    suffix < value.length - prefix &&
    oldValue.charCodeAt(oldValue.length - 1 - suffix) ===
      value.charCodeAt(value.length - 1 - suffix)
  ) {
    suffix += 1;
  }
  const shift = value.length - oldValue.length;
  const suffixStart = oldValue.length - suffix;
  const kept: SpreadsheetPhoneticRun[] = [];
  for (const run of runs) {
    if (run.end <= prefix && run.end > run.start) kept.push({ ...run });
    else if (run.start >= suffixStart && run.end > run.start) {
      kept.push({ start: run.start + shift, end: run.end + shift, text: run.text });
    }
  }
  return kept.length > 0 ? kept : undefined;
}

/**
 * The phonetic field a value write leaves (typing, paste, fill, F2).
 * Show Phonetic Field is the cell's and survives any write. A new entry
 * replaces the string, so its furigana go (Microsoft: "Excel automatically
 * adds phonetic symbols only to new data entered in the Japanese version");
 * the cell's character type, alignment and font stay with its text. An
 * in-place edit (`editedInPlace`) keeps the readings of the characters it
 * did not touch.
 */
export function phoneticAfterValueWrite(
  previous: unknown,
  previousValue: string,
  value: string,
  options: { editedInPlace?: boolean; textual?: boolean } = {},
): SpreadsheetCellPhonetic | undefined {
  if (!isRecord(previous)) return undefined;
  const textual = options.textual !== false;
  const current = normalizeCellPhonetic(previous, previousValue, true);
  if (!current) return undefined;
  if (value === previousValue && textual) return current;
  const next: SpreadsheetCellPhonetic = { ...current };
  delete next.runs;
  if (textual && options.editedInPlace) {
    const runs = phoneticRunsForEditedText(current.runs, previousValue, value);
    if (runs) next.runs = runs;
  }
  return normalizeCellPhonetic(next, value, textual);
}

// ── Character type conversion ────────────────────────────────────────────────

const HALFWIDTH_KATAKANA: Readonly<Record<string, string>> = (() => {
  // Fullwidth katakana → halfwidth forms (JIS X 0201). Voiced and
  // semi-voiced syllables are a base letter followed by ﾞ / ﾟ.
  const plain =
    "ァｧアｱィｨイｲゥｩウｳェｪエｴォｫオｵカｶキｷクｸケｹコｺサｻシｼスｽセｾソｿタﾀチﾁッｯツﾂテﾃトﾄナﾅニﾆヌﾇネﾈノﾉハﾊヒﾋフﾌヘﾍホﾎマﾏミﾐムﾑメﾒモﾓャｬヤﾔュｭユﾕョｮヨﾖラﾗリﾘルﾙレﾚロﾛワﾜヲｦンﾝーｰ・･。｡「｢」｣、､";
  const map: Record<string, string> = {};
  for (let index = 0; index + 1 < plain.length; index += 2) {
    map[plain[index]] = plain[index + 1];
  }
  const voiced = "ガカギキグクゲケゴコザサジシズスゼセゾソダタヂチヅツデテドトバハビヒブフベヘボホヴウ";
  for (let index = 0; index + 1 < voiced.length; index += 2) {
    map[voiced[index]] = `${map[voiced[index + 1]]}ﾞ`;
  }
  const semiVoiced = "パハピヒプフペヘポホ";
  for (let index = 0; index + 1 < semiVoiced.length; index += 2) {
    map[semiVoiced[index]] = `${map[semiVoiced[index + 1]]}ﾟ`;
  }
  return map;
})();

const FULLWIDTH_FROM_HALFWIDTH: Readonly<Record<string, string>> = (() => {
  const map: Record<string, string> = {};
  for (const [full, half] of Object.entries(HALFWIDTH_KATAKANA)) {
    if (!(half in map) || half.length === 2) map[half] = full;
  }
  return map;
})();

function hiraganaToKatakana(text: string): string {
  return text.replace(/[ぁ-ゖゝゞ]/g, (char) =>
    String.fromCharCode(char.charCodeAt(0) + 0x60),
  );
}

function katakanaToHiragana(text: string): string {
  return text.replace(/[ァ-ヶヽヾ]/g, (char) =>
    String.fromCharCode(char.charCodeAt(0) - 0x60),
  );
}

function halfwidthToFullwidthKatakana(text: string): string {
  let out = "";
  for (let index = 0; index < text.length; index += 1) {
    const pair = text.slice(index, index + 2);
    if (pair.length === 2 && FULLWIDTH_FROM_HALFWIDTH[pair]) {
      out += FULLWIDTH_FROM_HALFWIDTH[pair];
      index += 1;
      continue;
    }
    out += FULLWIDTH_FROM_HALFWIDTH[text[index]] ?? text[index];
  }
  return out;
}

/** Furigana shown in a character type (Phonetic Settings ▸ Type). */
export function convertPhoneticText(
  text: string,
  type: SpreadsheetPhoneticType,
): string {
  if (type === "noConversion") return text;
  const katakana = hiraganaToKatakana(halfwidthToFullwidthKatakana(text));
  if (type === "fullwidthKatakana") return katakana;
  if (type === "hiragana") return katakanaToHiragana(katakana);
  return Array.from(katakana, (char) => HALFWIDTH_KATAKANA[char] ?? char).join("");
}

// ── Reading and display ──────────────────────────────────────────────────────

/**
 * What PHONETIC(cell) returns: the cell's text with each run's base
 * characters replaced by their reading, in the cell's character type. A
 * cell without furigana reads as its own text.
 */
export function cellPhoneticReading(
  cell: PhoneticCellLike | null | undefined,
): string {
  const value = typeof cell?.value === "string" ? cell.value : "";
  const phonetic = readCellPhonetic(cell);
  if (!phonetic?.runs) return value;
  const type = phonetic.type ?? DEFAULT_PHONETIC_TYPE;
  let out = "";
  let position = 0;
  for (const run of phonetic.runs) {
    out += value.slice(position, run.start);
    out += convertPhoneticText(run.text, type);
    position = Math.max(position, run.end);
  }
  return out + value.slice(position);
}

export interface PhoneticSegment {
  base: string;
  /** The reading over `base`, as stored. */
  ruby?: string;
}









export function phoneticSegments(
  value: string,
  phonetic: SpreadsheetCellPhonetic | undefined,
): PhoneticSegment[] {
  const runs = phonetic?.runs ?? [];
  const segments: PhoneticSegment[] = [];
  let position = 0;
  for (const run of runs) {
    if (run.end <= run.start || run.start < position) continue;
    if (run.start > position) segments.push({ base: value.slice(position, run.start) });
    segments.push({
      base: value.slice(run.start, run.end),
      ruby: run.text,
    });
    position = run.end;
  }
  if (position < value.length || segments.length === 0) {
    segments.push({ base: value.slice(position) });
  }
  return segments;
}

/** The phonetic font a cell draws its guide in. */
export function effectivePhoneticFont(
  phonetic: SpreadsheetCellPhonetic | undefined,
  cellFontFamily: string | undefined,
): SpreadsheetPhoneticFont {
  return (
    phonetic?.font ?? {
      family: cellFontFamily || "",
      size: DEFAULT_PHONETIC_FONT_PT,
    }
  );
}

/** The guide line's height in px (the band Excel adds above the text). */
export function phoneticLineHeightPx(font: SpreadsheetPhoneticFont): number {
  return Math.ceil((ptToPx(font.size)) * 1.25);
}

/**
 * The guide's drawn size in px: the phonetic font's points in proportion to
 * the cell's own size (`cellFontSize` as stored, `cellFontPx` as drawn), or
 * its points as pixels where the cell's drawn size is unknown. Every surface
 * that shows the guide — the grid, Page Layout, Copy as Picture — sizes it
 * with this one rule (live audit B388).
 */
export function phoneticRubyFontPx(
  font: SpreadsheetPhoneticFont,
  cellFontSize: unknown,
  cellFontPx: number | undefined,
): number {
  const basePt = typeof cellFontSize === "number" && cellFontSize > 0 ? cellFontSize : 11;
  return cellFontPx !== undefined ? cellFontPx * (font.size / basePt) : ptToPx(font.size);
}

/** What a surface needs to draw a cell's shown guide, or null when the cell
 * shows none (Show Phonetic Field off, no readings, no text). */
export interface PhoneticGuideLayout {
  segments: PhoneticSegment[];
  alignment: SpreadsheetPhoneticAlignment;
  font: SpreadsheetPhoneticFont;
  /** The band's height in px, above the cell's text. */
  bandHeightPx: number;
}

export function cellPhoneticGuideLayout(
  cell: PhoneticCellLike | null | undefined,
  cellFontFamily: string | undefined,
): PhoneticGuideLayout | null {
  if (!cellPhoneticVisible(cell)) return null;
  const phonetic = readCellPhonetic(cell);
  const value = typeof cell?.value === "string" ? cell.value : "";
  if (!phonetic?.runs || value === "") return null;
  const font = effectivePhoneticFont(phonetic, cellFontFamily);
  return {
    segments: phoneticSegments(value, phonetic),
    alignment: phonetic.alignment ?? DEFAULT_PHONETIC_ALIGNMENT,
    font,
    bandHeightPx: phoneticLineHeightPx(font),
  };
}

// ── Commands ─────────────────────────────────────────────────────────────────

/** The cell with Show Phonetic Field set, or the same object when unchanged. */
export function cellWithPhoneticVisible<T extends PhoneticCellLike>(
  cell: T,
  visible: boolean,
): T {
  if (cellPhoneticVisible(cell) === visible) return cell;
  const value = typeof cell.value === "string" ? cell.value : "";
  const current = isRecord(cell.phonetic) ? { ...cell.phonetic } : {};
  if (visible) current.visible = true;
  else delete current.visible;
  const next = normalizeCellPhonetic(current, value, cellIsTextual(cell));
  const out = { ...cell };
  if (next) out.phonetic = next;
  else delete out.phonetic;
  return out;
}

/**
 * The cell with Phonetic Settings applied, or the same object when nothing
 * changes. Only a cell holding text has a string to carry them.
 */
export function cellWithPhoneticSettings<T extends PhoneticCellLike>(
  cell: T,
  settings: SpreadsheetPhoneticSettings,
): T {
  if (!cellIsTextual(cell) || cell.value === "") return cell;
  const value = cell.value as string;
  const current = readCellPhonetic(cell) ?? {};
  // The guide shows its readings as stored (phoneticSegments), so a new
  // character type rewrites them; an unchanged type leaves them as they are.
  const retyped = settings.type !== (current.type ?? DEFAULT_PHONETIC_TYPE);
  const next = normalizeCellPhonetic(
    {
      ...current,
      ...(retyped && current.runs
        ? { runs: current.runs.map((run) => ({ ...run, text: convertPhoneticText(run.text, settings.type) })) }
        : {}),
      type: settings.type,
      alignment: settings.alignment,
      font: settings.font ?? undefined,
    },
    value,
  );
  const unchanged =
    (next?.type ?? DEFAULT_PHONETIC_TYPE) ===
      (current.type ?? DEFAULT_PHONETIC_TYPE) &&
    (next?.alignment ?? DEFAULT_PHONETIC_ALIGNMENT) ===
      (current.alignment ?? DEFAULT_PHONETIC_ALIGNMENT) &&
    phoneticFontsEqual(next?.font, current.font) &&
    JSON.stringify(cell.phonetic ?? null) === JSON.stringify(next ?? null);
  if (unchanged) return cell;
  const out = { ...cell };
  if (next) out.phonetic = next;
  else delete out.phonetic;
  return out;
}

/** A deep copy safe to hand to another cell (clipboard, fill). */
export function clonePhonetic(
  phonetic: SpreadsheetCellPhonetic | undefined,
): SpreadsheetCellPhonetic | undefined {
  if (!phonetic) return undefined;
  return {
    ...phonetic,
    ...(phonetic.runs ? { runs: phonetic.runs.map((run) => ({ ...run })) } : {}),
    ...(phonetic.font ? { font: { ...phonetic.font } } : {}),
  };
}
