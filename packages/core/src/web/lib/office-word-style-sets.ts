/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word's built-in style sets (Design > Document Formatting), as Word for Mac
 * 16.86 applies them: the twelve sets Word ships (`QuickStyles/*.dotx`) plus
 * `default`, Word's Normal template. Applying one replaces the document's
 * defaults (`w:docDefaults`) and the set's Quick Styles; the theme stays,
 * and every font and colour the styles take from it stays a theme reference.
 *
 * The sets are what Word itself wrote after applying each one from the
 * ribbon (apps/api/scripts/generate_word_style_sets.py); the DOCX writer
 * has the same sets as Word's XML. Being measured from Word, they are not
 * in the browser bundle: the API serves them to a signed-in page as the
 * Office data table `word-style-sets` (lib/office-data), and until they are
 * here every reader below answers "not yet" (the Design commands refuse,
 * the galleries draw placeholders). Only the sets' ids ship with the app.
 * A page that can never have them (a guest's, the marketing site's
 * editors) reads MRAI's own recipe of the sets instead
 * (`office-word-style-set-fallback`), so its Design tab works.
 */

import { loadOfficeData, officeData, officeDataObtainable } from "./office-data/store";
import type { OfficeThemeLabel } from "./office-theme-catalog";
import { fallbackWordStyleSets } from "./office-word-style-set-fallback";

/** A colour as Word states it: `val` is the cached hex, the rest the theme reference. */
export interface WordStyleColor {
  readonly val?: string;
  readonly themeColor?: string;
  readonly themeTint?: string;
  readonly themeShade?: string;
  readonly fill?: string;
  readonly themeFill?: string;
  readonly themeFillTint?: string;
  readonly themeFillShade?: string;
  readonly color?: string;
}

export interface WordStyleBorder extends WordStyleColor {
  readonly sz?: number;
  readonly space?: number;
}

/** `w:pPr`, in Word's units (twentieths of a point; `line` in 240ths when `lineRule` is auto). */
export interface WordStyleParagraphProps {
  readonly before?: number;
  readonly after?: number;
  readonly line?: number;
  readonly lineRule?: string;
  readonly jc?: string;
  readonly ind?: Readonly<{ left?: number; right?: number; firstLine?: number; hanging?: number }>;
  readonly keepNext?: boolean;
  readonly keepLines?: boolean;
  readonly contextualSpacing?: boolean;
  readonly widowControl?: boolean;
  readonly outlineLvl?: number;
  readonly pBdr?: Readonly<Partial<Record<"top" | "left" | "bottom" | "right" | "between", WordStyleBorder>>>;
  readonly shd?: WordStyleColor;
}

/** `w:rPr`: `fontTheme` is the theme font the Latin text takes (`+Headings` / `+Body`). */
export interface WordStyleRunProps {
  readonly fontTheme?: "major" | "minor";
  readonly eastAsiaTheme?: "major" | "minor";
  readonly font?: string;
  /** Half-points. */
  readonly sz?: number;
  readonly b?: boolean;
  readonly i?: boolean;
  readonly caps?: boolean;
  readonly smallCaps?: boolean;
  readonly strike?: boolean;
  readonly dstrike?: boolean;
  readonly color?: WordStyleColor;
  /** Character spacing, twentieths of a point. */
  readonly spacing?: number;
  readonly kern?: number;
  readonly u?: WordStyleColor;
  readonly vertAlign?: string;
  readonly bdr?: WordStyleBorder;
  readonly shd?: WordStyleColor;
  readonly w14?: readonly string[];
}

export interface WordStyleDefinition {
  readonly type: "paragraph" | "character";
  readonly basedOn: string | null;
  readonly next: string | null;
  readonly pPr: WordStyleParagraphProps;
  readonly rPr: WordStyleRunProps;
}

/** The gallery styles a style set defines (the Home tab's Quick Styles). */
export type WordQuickStyleId =
  | "normal"
  | "noSpacing"
  | "heading1"
  | "heading2"
  | "heading3"
  | "heading4"
  | "heading5"
  | "heading6"
  | "heading7"
  | "heading8"
  | "heading9"
  | "title"
  | "subtitle"
  | "quote"
  | "intenseQuote"
  | "listParagraph"
  | "subtleEmphasis"
  | "emphasis"
  | "intenseEmphasis"
  | "strong"
  | "subtleReference"
  | "intenseReference"
  | "bookTitle";

export interface WordStyleSet {
  readonly id: string;
  readonly label: OfficeThemeLabel;
  readonly docDefaults: { readonly rPr: WordStyleRunProps; readonly pPr: WordStyleParagraphProps };
  readonly styles: Readonly<Partial<Record<WordQuickStyleId, WordStyleDefinition>>>;
}

/** Every style set's id, in the order Word's gallery lists them (`default` first). */
export const WORD_STYLE_SET_IDS = Object.freeze([
  "default",
  "distinctive",
  "elegant",
  "fancy",
  "formal",
  "manuscript",
  "modern",
  "newsprint",
  "perspective",
  "simple",
  "thatch",
  "traditional",
  "word2010",
] as const);

export type WordStyleSetId = (typeof WORD_STYLE_SET_IDS)[number];

export const WORD_DEFAULT_STYLE_SET_ID = "default";

const IDS: ReadonlySet<string> = new Set(WORD_STYLE_SET_IDS);

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

interface StyleSetTable {
  readonly raw: unknown;
  readonly list: readonly WordStyleSet[];
  readonly byId: ReadonlyMap<string, WordStyleSet>;
}

let served: StyleSetTable | null = null;

/**
 * The served table, checked: every set this app knows, in its order, each
 * with its label, document defaults and styles. Anything else is no table.
 */
function servedSets(): StyleSetTable | null {
  if (recipeForced) return fallbackSets();
  const raw = officeData("word-style-sets");
  if (raw === undefined) return officeDataObtainable() ? null : fallbackSets();
  if (served?.raw === raw) return served;
  const sets = isRecord(raw) ? raw.styleSets : undefined;
  if (!Array.isArray(sets) || sets.length !== WORD_STYLE_SET_IDS.length) return null;
  const ok = sets.every(
    (set, index) =>
      isRecord(set) &&
      set.id === WORD_STYLE_SET_IDS[index] &&
      isRecord(set.label) &&
      typeof set.label.en === "string" &&
      isRecord(set.docDefaults) &&
      isRecord(set.docDefaults.rPr) &&
      isRecord(set.docDefaults.pPr) &&
      isRecord(set.styles),
  );
  if (!ok) return null;
  const list = Object.freeze(sets as WordStyleSet[]);
  served = { raw, list, byId: new Map(list.map((set) => [set.id, set])) };
  return served;
}

let fallback: StyleSetTable | null = null;

/** The recipe's sets, for a page that cannot ask for Word's (a guest). */
function fallbackSets(): StyleSetTable {
  if (!fallback) {
    const list = fallbackWordStyleSets();
    fallback = { raw: list, list, byId: new Map(list.map((set) => [set.id, set])) };
  }
  return fallback;
}

let recipeForced = false;

/**
 * Runs `read` against MRAI's recipe of the sets, whatever this page has:
 * what a document styled on a guest page holds, for swapping it to Word's
 * own once they are here (`restyleFromRecipeStyleSets`). Synchronous.
 */
export function withRecipeWordStyleSets<T>(read: () => T): T {
  const before = recipeForced;
  recipeForced = true;
  try {
    return read();
  } finally {
    recipeForced = before;
  }
}

/** Whether Word's own sets are here (not the recipe a guest page reads). */
export function wordStyleSetsServed(): boolean {
  if (recipeForced) return false;
  const table = servedSets();
  return table !== null && table !== fallback;
}

/** Every style set, in Word's gallery order, or null until they are here. */
export function wordStyleSets(): readonly WordStyleSet[] | null {
  return servedSets()?.list ?? null;
}

/** Whether the style sets are here (the Design commands need them). */
export function wordStyleSetsAvailable(): boolean {
  return servedSets() !== null;
}

/** Asks for the style sets once (an editor opening); true when they are here. */
export async function loadWordStyleSets(): Promise<boolean> {
  await loadOfficeData("word-style-sets");
  return wordStyleSetsAvailable();
}

/** One style set, or undefined (an unknown id, or the sets not here yet). */
export function wordStyleSet(id: unknown): WordStyleSet | undefined {
  return typeof id === "string" ? servedSets()?.byId.get(id) : undefined;
}

export function isWordStyleSetId(id: unknown): id is WordStyleSetId {
  return typeof id === "string" && IDS.has(id);
}

/** Word's own names by gallery id (`heading 1` -> heading1), for style references in the data. */
const NAME_TO_ID: Readonly<Record<string, WordQuickStyleId>> = {
  Normal: "normal",
  "No Spacing": "noSpacing",
  "heading 1": "heading1",
  "heading 2": "heading2",
  "heading 3": "heading3",
  "heading 4": "heading4",
  "heading 5": "heading5",
  "heading 6": "heading6",
  "heading 7": "heading7",
  "heading 8": "heading8",
  "heading 9": "heading9",
  Title: "title",
  Subtitle: "subtitle",
  Quote: "quote",
  "Intense Quote": "intenseQuote",
  "List Paragraph": "listParagraph",
  "Subtle Emphasis": "subtleEmphasis",
  Emphasis: "emphasis",
  "Intense Emphasis": "intenseEmphasis",
  Strong: "strong",
  "Subtle Reference": "subtleReference",
  "Intense Reference": "intenseReference",
  "Book Title": "bookTitle",
};

/** The Quick Style Word's built-in style name stands for (`heading 1`, `Title`...). */
export function wordQuickStyleIdForName(name: unknown): WordQuickStyleId | null {
  if (typeof name !== "string") return null;
  const exact = NAME_TO_ID[name];
  if (exact) return exact;
  const lower = name.toLowerCase();
  const found = Object.entries(NAME_TO_ID).find(([key]) => key.toLowerCase() === lower);
  return found ? found[1] : null;
}

/**
 * A Quick Style's own properties merged down its `basedOn` chain, stopping
 * before Normal (Normal and the document defaults are the document's base,
 * which the editor keeps as document attributes, not per paragraph). A
 * property the style states wins over the one it inherits; `pBdr` and `ind`
 * merge edge by edge, as Word's style inheritance does.
 */
export function wordStyleChain(
  set: WordStyleSet,
  id: WordQuickStyleId,
): { pPr: WordStyleParagraphProps; rPr: WordStyleRunProps } {
  const chain: WordStyleDefinition[] = [];
  const seen = new Set<string>();
  let current: WordQuickStyleId | undefined = id;
  while (current && current !== "normal" && !seen.has(current)) {
    seen.add(current);
    const style: WordStyleDefinition | undefined = set.styles[current];
    if (!style) break;
    chain.push(style);
    current = style.basedOn ? NAME_TO_ID[style.basedOn] : undefined;
  }
  let pPr: WordStyleParagraphProps = {};
  let rPr: WordStyleRunProps = {};
  for (const style of chain.reverse()) {
    pPr = {
      ...pPr,
      ...style.pPr,
      ...(pPr.ind || style.pPr.ind ? { ind: { ...pPr.ind, ...style.pPr.ind } } : {}),
      ...(pPr.pBdr || style.pPr.pBdr ? { pBdr: { ...pPr.pBdr, ...style.pPr.pBdr } } : {}),
    };
    rPr = { ...rPr, ...style.rPr };
  }
  return { pPr, rPr };
}

/**
 * The document's base: `w:docDefaults` under the Normal style's own
 * properties (Normal is based on nothing; its text is the defaults').
 */
export function wordStyleSetBase(set: WordStyleSet): {
  pPr: WordStyleParagraphProps;
  rPr: WordStyleRunProps;
} {
  const normal = set.styles.normal;
  return {
    pPr: {
      ...set.docDefaults.pPr,
      ...normal?.pPr,
      ...(set.docDefaults.pPr.ind || normal?.pPr.ind
        ? { ind: { ...set.docDefaults.pPr.ind, ...normal?.pPr.ind } }
        : {}),
    },
    rPr: { ...set.docDefaults.rPr, ...normal?.rPr },
  };
}
