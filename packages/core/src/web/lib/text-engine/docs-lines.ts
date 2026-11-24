/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Docs line layout from the text engine (PLAN-phase2.md, Docs).
 *
 * A Word paragraph as the editor draws it -- a block whose inline content is
 * text runs, atoms (pictures, list labels, fields drawn as inline blocks),
 * tabs and hard breaks -- read into one sequence of pieces: every grapheme of
 * text with its advance from the served font file (no kerning unless the run
 * kerns, no ligatures, the run's letter and word spacing on top), every atom
 * with its drawn width. The lines are then filled the way Word fills them:
 * trailing spaces hang past the edge, a piece overflows only when the pen
 * passes the edge, a word too long for a line breaks between characters, and
 * a justified line in compatibility mode 15 keeps a word Word would keep by
 * narrowing its spaces (docs-word-rules.ts).
 *
 * The browser's own lines are read from the same pieces' drawn boxes, so the
 * two can be compared paragraph by paragraph (`auditDocsParagraph`), and,
 * where they differ, the browser can be told where to break.
 */
import { KINSOKU_EAST_ASIAN, OFFICE_KINSOKU, type KinsokuList } from "./kinsoku";
import { lineBreakOpportunities, wordJoins } from "./line-break";
import { measureText, type EngineFont } from "./measure";
import { pageEngineFont } from "./page-adapter";
import { wordKeepsJustifiedWord, wordTrailingPunctuationLength } from "./docs-word-rules";
import { freshRangesKept, releaseFreshRange } from "../dom-range-release";
import { WORD_DEFAULT_TAB_INTERVAL_TWIPS, wordTabAdvanceAt, wordTabStopAfter, type TabStop } from "../document-tabstops";

export type DocsPieceKind = "text" | "space" | "box" | "tab" | "break";

export interface DocsPiece {
  readonly kind: DocsPieceKind;
  /** The grapheme ("￼" for an atom, "\t", "\n"). */
  readonly text: string;
  /** Advance in CSS px: the engine's for text, the drawn width for others. */
  readonly advance: number;
  /** Text pieces: where the grapheme starts in the DOM. */
  readonly node?: Text;
  readonly offset?: number;
  /** Atoms, tabs drawn as boxes, hard breaks drawn as <br>. */
  readonly element?: Element;
  /** The run's font size in px (text pieces). */
  readonly fontPx?: number;
  /** The run's engine font (text pieces), for the line's box. */
  readonly font?: EngineFont;
  /** A space's natural advance in its run (text pieces). */
  readonly spacePx?: number;
  /** Letter spacing an engine correction span adds to the drawn grapheme, px. */
  readonly correction?: number;
  /** The run's own letter spacing, px (text pieces). */
  readonly letterSpacing?: number;
  /** The run's own word spacing, px (text pieces). */
  readonly wordSpacing?: number;
  /** The hyphen a line ending after this piece draws, px, when not measured
   * from `font` (hyphenAdvance). */
  readonly hyphenPx?: number;
  /** Text pieces: the run's computed line-height (px or "normal") and its
   * vertical-align as a length (px, up), for the line's CSS box; absent
   * when an inline around the run aligns it otherwise (super, middle...). */
  readonly lineHeight?: number | "normal";
  readonly shift?: number;
  /** Atoms whose box the page states (a formula's placeholder): its margin
   * box height and its vertical-align length (px, up). */
  readonly atom?: { readonly height: number; readonly shift: number };
  /** A box drawn as an inline-block of one line (a list or heading number):
   * the text boxes of that line, which raise the line it stands on as they
   * would standing on it themselves (its baseline is theirs). */
  readonly lineBoxes?: readonly { readonly font: EngineFont; readonly lineHeight: number | "normal" }[];
  /** A tab whose next stop lies past the line's right edge in a Word 2010
   * layout (`docsTabOverMarginLineOff`): the rest of its line runs on past
   * the edge. Set by the page (bandTabLayout's kept line) or as the engine
   * fills the line from the paragraph's stops. */
  readonly overMargin?: boolean;
  /** Text pieces: the run breaks between any two letters (`word-break:
   * break-all`, Word's "Allow Latin text to wrap in the middle of a word"
   * in compatibility mode 14 and later). */
  readonly anywhere?: boolean;
  /** A line break kept as text: its run's face, which sets a line of
   * nothing else (Word sets a break's line by the rest of it). */
  readonly breakFont?: EngineFont;
}

/** The page's kept line after a Word 2010 tab past the edge (bandTabLayout). */
const TAB_OVER_MARGIN_SELECTOR = '[data-tab-over-margin="true"]';













export function docsTabOverMarginLineOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsTabOverMarginLineOff?: boolean }).__docsTabOverMarginLineOff === true;
}















export function docsTabPastIndentOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsTabPastIndentOff?: boolean }).__docsTabPastIndentOff === true;
}

/** The block's own font and line-height: the strut every line starts from. */
export interface DocsStrut {
  readonly font: EngineFont;
  readonly lineHeight: number | "normal";
}

export interface DocsParagraph {
  readonly pieces: readonly DocsPiece[];
  /** Available width of the first line and of every later line, CSS px. */
  readonly firstWidth: number;
  readonly restWidth: number;
  readonly align: string;
  /** Why the engine cannot lay this paragraph out (the browser keeps it). */
  readonly unsupported?: string;
  /** Pending: a face is still loading. */
  readonly pending?: boolean;
  /** What the paragraph holds that may explain a difference. */
  readonly flags: readonly string[];
  /** CSS px per drawn (client rect) px: 1 / the page's scale. */
  readonly unscale: number;
  /** Word's "Allow punctuation to overflow the margin" is on for the
   * paragraph (`data-document-effective-overflow-punctuation`). */
  readonly overflowPunctuation?: boolean;
  /** The paragraph stands in a table cell, where only a comma or stop may hang. */
  readonly inTableCell?: boolean;
  /** false: the cell's column is the stored grid's, not one the page sized to
   * a word (CELL_FIT_EPSILON's allowance does not apply). Word wraps a word a
   * fraction of a twip wider than such a cell's text width (Windows Word,
   * corpus R417-D03 page 36: Tw Cen MT 8 pt "Mechanical harvester", 1349.14
   * twips in a 1549-twip fixed column with 100-twip margins, two lines). */
  readonly cellWidened?: boolean;
  /** How much wider than `firstWidth`/`restWidth` the box the browser lays
   * the lines out in is, px: a Word cell's text box taken from the model
   * (cellModelContentWidth). A line the engine ends there with room left in
   * the browser's box for its next word is broken by a break of its own
   * (planDocsLineCorrections' tabBreak). */
  readonly browserWiderBy?: number;
  /** The block's strut, when its font could be read. */
  readonly strut?: DocsStrut;
  /** The block's padding and border above its first line and below its
   * last, px: its box is its lines and these. */
  readonly chrome?: { readonly top: number; readonly bottom: number };
  /** A list paragraph's first-line strut (`p[data-document-list-marker-strut]`
   * ::before): an empty inline of the block's font, this tall and shifted,
   * which a larger number raises the line by. */
  readonly firstLineStrut?: { readonly lineHeight: number; readonly shift: number };
  /** Column mode with the paragraph's stops: its tabs advance to Word's stops
   * as its lines are filled (`docsEngineLineStarts` sets a tab piece's
   * advance from the pen), measured from the column's left edge. */
  readonly tabStops?: DocsTabLayout;
  /** A drop cap (`p[data-drop-cap]::first-letter`, drawn as a CSS initial
   * letter) stands beside the paragraph's first `count` lines, which are
   * `width` px wide at most: the room right of the letter and its distance
   * from text. The letter itself is a box of no width in `pieces`. */
  readonly leadingLines?: { readonly count: number; readonly width: number };
  /** Word's automatic hyphenation of the paragraph (docs-hyphenation.ts):
   * where its words may end a line in a hyphen, and Word's limits. */
  readonly hyphenation?: DocsHyphenation;
}

/** Automatic hyphenation as the engine breaks lines by it. */
export interface DocsHyphenation {
  /** Pieces a line may end after with a hyphen: the letter before each of
   * the language's hyphenation points. */
  readonly after: ReadonlySet<number>;
  /** Word's hyphenation zone (`w:hyphenationZone`), px: compatibility mode
   * 14 and earlier only. */
  readonly zonePx: number;
  /** Lines in a row that may end in an automatic hyphen (0: no limit). */
  readonly consecutiveLimit: number;
}

/** A soft (optional) hyphen: nothing on the line, a hyphen at its end. */
export const SOFT_HYPHEN = "\u00ad";







export function docsHyphenationOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsHyphenationOff?: boolean }).__docsHyphenationOff === true;
}

/** The soft hyphen the engine draws where it ends a line at an automatic
 * hyphenation point: no piece of the paragraph. */
export const DOCS_ENGINE_HYPHEN_CLASS = "docs-engine-hyphen";
/** The engine's word joiner (U+2060) between closing punctuation and an
 * opening bracket Word keeps together (docsBracketJoins): read as nothing. */
export const DOCS_ENGINE_JOINER_CLASS = "docs-engine-joiner";
/** The hyphen the engine draws at the end of a page cut's head that ends
 * inside a hyphenated word: read as text, the last piece of its line. */
export const DOCS_ENGINE_HEAD_HYPHEN_CLASS = "docs-engine-head-hyphen";

/** The room line `line` of the paragraph has, CSS px. */
export function docsLineWidth(paragraph: DocsParagraph, line: number): number {
  const width = line === 0 ? paragraph.firstWidth : paragraph.restWidth;
  const lead = paragraph.leadingLines;
  return lead && line < lead.count ? Math.min(width, lead.width) : width;
}

/** A paragraph whose first letter Word drops (Insert > Drop Cap): the
 * letter is drawn by the ::first-letter rule in note-editor.css. */
const DROP_CAP_SELECTOR = '[data-drop-cap="dropped"]:not([data-drop-cap-frame]):not([data-page-split-continuation])';

/** `initial-letter: <size> [<sink>]` as computed: the lines the letter
 * stands beside, or null where it is not an initial letter (`normal`, a
 * browser without the property). */
function initialLetterLines(value: string): number | null {
  const numbers = value.trim().split(/\s+/).map(Number).filter((n) => Number.isFinite(n) && n > 0);
  if (!numbers.length) return null;
  const [size, sink = Math.floor(size)] = numbers;
  return Math.max(Math.ceil(size - 1e-6), Math.round(sink));
}

/** The paragraph's own tab stops, as `ReadDocsParagraphOptions.tabStops`
 * gives them (the editor node's `tabStops`, normalized). */
export interface DocsTabStopsInput {
  readonly stops: readonly TabStop[] | null;
  /** The paragraph's tabs run past the right margin as Word 2010's did
   * (`wordTabOverMargin`): a centre or right stop there is not moved in. */
  readonly overMargin?: boolean;
  /** The room the boxes around the paragraph (a list and its items) take on
   * each side of the column the stops measure from, px: `columnWidthPx` is
   * what they leave. */
  readonly aroundLeftPx?: number;
  readonly aroundRightPx?: number;
}

export interface DocsTabLayout {
  readonly stops: readonly TabStop[] | null;
  readonly intervalTwips: number;
  /** Where the paragraph's text box starts in the column, and what its
   * right side leaves, px (its margin, padding and border). */
  readonly contentLeftPx: number;
  readonly indentRightPx: number;
  readonly columnPx: number;
  readonly stopsEndAtRightEdge: boolean;
  /** With `stopsEndAtRightEdge` (Word 2013 and later): a stop past the right
   * indent but inside the column stays where it is, and its line fills to
   * the column's edge (`docsTabPastIndentOff`). */
  readonly stopsReachColumn?: boolean;
}

/** A formula placeholder's box as the page states it (`--math-placeholder-*`,
 * note-editor.css: a hidden-overflow inline-block, so its baseline is its
 * bottom margin edge, lowered by the formula's depth); nothing for any other
 * atom. */
function placeholderAtom(node: Element, style: CSSStyleDeclaration): { atom?: { height: number; shift: number } } {
  if (!node.matches(".inline-math-preview[data-math-pending][data-math-layout-ready=\"inline\"]")) return {};
  const height = px(style.getPropertyValue("--math-placeholder-height"));
  const depth = px(style.getPropertyValue("--math-placeholder-depth") || "0px");
  if (!(height >= 0) || !(depth >= 0) || style.overflow === "visible") return {};
  return { atom: { height: height + px(style.marginTop) + px(style.marginBottom), shift: -depth } };
}

/** An atom's border-box width, CSS px: a formula placeholder's stated one
 * (no layout read), any other's drawn box. */
const STATED_WIDTH_ATOM = ".inline-math-preview[data-math-pending][data-math-layout-ready=\"inline\"]";

function atomWidthPx(node: Element, style: CSSStyleDeclaration, unscale: number): number {
  if (node.matches(STATED_WIDTH_ATOM)) {
    const stated = px(style.getPropertyValue("--math-placeholder-width"));
    if (stated >= 0) return stated;
  }
  return node.getBoundingClientRect().width * unscale;
}

/** A computed line-height: px, or "normal". */
function cssLineHeight(value: string): number | "normal" | null {
  if (value === "normal") return "normal";
  return value.endsWith("px") ? px(value) : null;
}

/** A computed vertical-align as a length (px, up), or null for a keyword
 * other than baseline (whose box the CSS model does not place). */
function cssShift(value: string): number | null {
  if (value === "baseline") return 0;
  return value.endsWith("px") ? px(value) : null;
}

export interface DocsLayoutRules {
  /** Word compatibility mode of the document (null: none recorded). */
  readonly compatibilityMode: number | null;
  readonly kinsoku: KinsokuList | null;
  /** Ideographs break as Latin letters (wordSetsIdeographsAsLetters). */
  readonly ideographsAsLetters?: boolean;
  /** False where the East Asian text is in a Western language
   * (wordEastAsianLineRulesApply): no kinsoku, a full-width mark ("）",
   * "。") breaks as an ideograph does and does not hang past the edge. */
  readonly eastAsianLineRules?: boolean;
}

const OUT_OF_FLOW_POSITIONS = new Set(["absolute", "fixed"]);
const ATOMIC_DISPLAYS = new Set(["inline-block", "inline-flex", "inline-grid", "inline-table"]);
const REPLACED = new Set(["IMG", "SVG", "svg", "CANVAS", "VIDEO", "INPUT", "IFRAME", "OBJECT", "math", "MATH"]);
const INLINE_TAGS_SAFE = new Set(["SPAN", "A", "B", "I", "U", "S", "STRONG", "EM", "SUB", "SUP", "MARK", "SMALL", "CODE", "INS", "DEL", "FONT", "ABBR", "BDI", "Q", "CITE", "DFN", "KBD", "SAMP", "VAR", "TIME", "DATA", "LABEL"]);
const FIT_EPSILON = 1 / 256;
/**
 * In a table cell: the cell's text width comes through the browser's layout,
 * which keeps the column, each padding and each border in 1/64 px units, so a
 * column Word widened to its longest word (wordTableAutofitDisplay) reads a few
 * units short of that word. Corpus R283-D20, R283-D23, R283-D28, R282-D08: a
 * 22 pt bold Arial heading ("SECURITY_S", 184.19 px) in a text width read as
 * 184.172 px was cut before its last letter; Word's PDF draws each on one line
 * (138.64 pt of ink at 22.08 pt). 1/16 px is under one twip (1/15 px).
 */
const CELL_FIT_EPSILON = 1 / 16;
function fitEpsilon(paragraph: DocsParagraph): number {
  return paragraph.inTableCell && (paragraph.cellWidened !== false || cellFitExactOff()) ? CELL_FIT_EPSILON : FIT_EPSILON;
}
/** On the spans the East Asian spacing projection draws (DocumentPageSetup). */
const WORD_CJK_SPACING_ATTR = "data-docs-word-cjk-spacing";
/** Harness switch (`window.__docsCellFitExactOff = true`): every cell gets the
 * widened column's allowance, as before (A/B). */
function cellFitExactOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsCellFitExactOff?: boolean }).__docsCellFitExactOff === true;
}
/** A table whose columns the page sized from their contents (Word's AutoFit,
 * or a stored grid widened to a word it must hold): wordTableAutofitDisplay. */
const WIDENED_TABLE_SELECTOR = "table[data-word-autofit-display], table[data-word-min-content-floor]";
/**
 * On a Word cell: its text box, px, as the model states it (the stored widths
 * of the grid columns it spans less its left and right margins; tableBorders).
 * The browser keeps the column, each padding and each border in 1/64 px units
 * and gives the cell's end padding back 1/32 px, so the text width it lays out
 * reads up to ~0.05 px wider than Word's. Windows Word wraps a word a fraction
 * of a twip wider than the model's (corpus R417-D03 page 36: Tw Cen MT 8 pt
 * "Mechanical harvester", 1349.14 twips, in a 1549-twip fixed column with
 * 100-twip margins -- 1349 twips of text, read 89.984 px for 89.933 -- is two
 * lines in Word).
 */
export const WORD_CELL_TEXT_PX_ATTR = "data-word-cell-text-px";
/** The model's width stands in for the browser's only within this of it: a
 * column the page drew other than at its stored width (shrunk to the text
 * column, a percentage share) keeps the width the browser laid it out at. */
const CELL_MODEL_WIDTH_TOLERANCE_PX = 0.25;
/** Harness switch (`window.__docsCellModelWidthOff = true`): a cell's text
 * width is the browser's, as before (A/B). */
function cellModelWidthOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsCellModelWidthOff?: boolean }).__docsCellModelWidthOff === true;
}
/** A cell paragraph's text width as Word's model states it, when its column is
 * the stored grid's and the browser's reading is that width rounded. */
function cellModelContentWidth(element: HTMLElement, style: CSSStyleDeclaration, browserContent: number): number | null {
  if (cellModelWidthOff()) return null;
  const cell = element.closest("td, th");
  const stated = cell?.getAttribute(WORD_CELL_TEXT_PX_ATTR);
  if (!cell || !stated || element.closest("table")?.matches(WIDENED_TABLE_SELECTOR)) return null;
  // Only a paragraph standing straight in the cell: a box between them (a
  // list, a nested block) takes room the model width does not know of.
  if (element.parentElement !== cell) return null;
  const textBox = Number(stated);
  if (!Number.isFinite(textBox) || textBox <= 0) return null;
  const model = textBox - px(style.marginLeft) - px(style.marginRight) - px(style.paddingLeft) - px(style.paddingRight)
    - px(style.borderLeftWidth) - px(style.borderRightWidth);
  return model > 0 && Math.abs(model - browserContent) < CELL_MODEL_WIDTH_TOLERANCE_PX ? model : null;
}
/** On the spans the engine's line corrections draw (documentEngineLineBreaks). */
export const DOCS_ENGINE_LINE_ATTR = "data-docs-engine-line";
/** The line break the engine draws before a tab that opens a line: no piece
 * of the paragraph (the engine breaks there by itself). */
export const DOCS_ENGINE_TAB_BREAK_CLASS = "docs-engine-tab-break";
/** Harness switch (`window.__docsEngineBreakStrutOff = true`): the engine's
 * line break revokes its paragraph's leaf line boxes, as before (A/B). */
export function engineBreakStrutOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsEngineBreakStrutOff?: boolean }).__docsEngineBreakStrutOff === true;
}
/** bandTabLayout's box around a tab it laid out. */
const LAID_TAB_CLASS = "sheet-band-story-tab";
/** The run's own letter spacing under a correction span, px. */
export const DOCS_ENGINE_BASE_SPACING_ATTR = "data-docs-engine-base-spacing";
/** The run's own word spacing under a correction span, px. */
export const DOCS_ENGINE_BASE_WORD_SPACING_ATTR = "data-docs-engine-base-word-spacing";
/** The marks Word lets hang past the right edge (overflowPunctuationHang.ts). */
const HANGING = new Set(["，", "。", "、", "．", "﹐", "﹒", "｡", "､", ",", "."]);
/** A mark kinsoku keeps off a line's start that Word hangs: punctuation
 * outside ASCII (not "々", "ゝ" or "℃"). */
const EAST_ASIAN_HANGING_MARK = /^[^\x00-\x7f]$/u;
const EAST_ASIAN_PUNCTUATION = /^\p{P}$/u;
function eastAsianHangOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsEastAsianHangOff?: boolean }).__docsEastAsianHangOff === true;
}
/** An East Asian (full-width or ideographic) mark. */
const FULL_WIDTH_MARK = /[\u3000-\u303f\ufe10-\ufe6f\uff00-\uffef]/;

function px(value: string | null | undefined): number {
  const parsed = Number.parseFloat(value ?? "");
  return Number.isFinite(parsed) ? parsed : 0;
}

function hasPseudoContent(view: Window, element: Element, which: "::before" | "::after"): boolean {
  const style = view.getComputedStyle(element, which);
  const content = style.content;
  if (!content || content === "none" || content === "normal" || content === '""' || content === "''") return false;
  // A label drawn out of the flow (a source numbering label) paints only.
  if (style.position === "absolute" || style.position === "fixed") return false;
  return style.display !== "none";
}

/** The text of a `content` value that is one quoted string, or null (a
 * counter, an attr(), several parts). */
function literalContent(content: string): string | null {
  const match = /^"((?:[^"\\]|\\.)*)"$/.exec(content.trim());
  return match ? match[1]!.replace(/\\(.)/g, "$1") : null;
}

/**
 * The room a list label drawn by `::before` takes at the start of the first
 * line, read with no layout: an inline-block of one literal glyph string
 * (a native list's marker) is the wider of its `min-width` and the glyph,
 * plus its margins (a negative start margin hangs it into the indent).
 * Null where only the page can say; "pending" while its face loads.
 */
function unlaidLabelAdvance(view: Window, element: Element): number | null | "pending" {
  const style = view.getComputedStyle(element, "::before");
  if (style.display !== "inline-block" || (style.float && style.float !== "none")) return null;
  const text = literalContent(style.content);
  if (text === null) return null;
  if (!/px$/.test(style.minWidth) && style.minWidth !== "auto" && style.minWidth !== "0px") return null;
  const minWidth = style.minWidth === "auto" ? 0 : px(style.minWidth);
  let glyph = 0;
  if (text) {
    const font = pageEngineFont(`${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`);
    const measured = font ? measureText(text, { ...font, followLocalOnly: true }) : null;
    if (measured?.status === "pending") return "pending";
    if (measured?.status === "ok") glyph = measured.width;
    // One symbol (a bullet in a symbol face the engine has no metrics for)
    // in a box at least one and a half of its em wide: no glyph is that
    // wide, so the box is its min-width.
    else if (!(Array.from(text).length === 1 && minWidth >= 1.5 * px(style.fontSize))) return null;
  }
  return Math.max(minWidth, glyph) + px(style.marginLeft) + px(style.marginRight);
}

/** Width a run's text-transform gives: what the browser draws. */
function transformed(text: string, transform: string): string {
  if (transform === "uppercase") return text.toUpperCase();
  if (transform === "lowercase") return text.toLowerCase();
  return text;
}

/**
 * Read a paragraph element into pieces. The block's own box gives the line
 * widths; `skip` drops editor chrome that takes no room (formatting marks).
 */
export interface ReadDocsParagraphOptions {
  /**
   * Read the paragraph without laying the page out: its width is this
   * column's (CSS px) less its own margins, padding and borders, read as
   * computed values (a px margin or padding is no layout read; `width` and
   * any box are). What only a laid-out page can say -- a tab's or an
   * object's drawn width, a CSS list label's room -- makes it unsupported.
   * The page ledger plans the pages it has not laid out with this.
   */
  readonly columnWidthPx?: number;
  /**
   * In column mode, a tab taken as wide as the interval the browser draws
   * the paragraph's tabs at (its `tab-size`, Word's default stops) instead
   * of leaving the paragraph to the page: its lines are then an estimate
   * (`flags` holds "approximate"), for a plan the page still checks.
   */
  readonly approximateTabs?: boolean;
  /**
   * In column mode, a run in a family the browser draws from an installed
   * font that differs from the served file read with the served file's
   * metrics instead of leaving the paragraph to the page: its lines are then
   * an estimate (`flags` holds "approximate").
   */
  readonly approximateFonts?: boolean;
  /**
   * In column mode, the paragraph's tab stops: each tab goes to its stop
   * (left, centre, right, decimal; past the last one Word's default stops at
   * the paragraph's `tab-size`) from where its line has got to, as Word and
   * the page's own tab layout (bandTabLayout) place it. Positional tabs are
   * not this: leave it out for a paragraph that holds one.
   */
  readonly tabStops?: DocsTabStopsInput;
  /**
   * In column mode, the box of the `ordinal`-th in-line formula preview the
   * paragraph draws (`.inline-math-preview`, in document order) when the
   * page states none (it has not been typeset or placed): the engine's
   * (paged/ledger/ledgerEngineMath). A paragraph read with a box that is not
   * `exact` holds "engine-formula" in its flags.
   */
  readonly formulaBox?: (ordinal: number, preview: HTMLElement) => DocsFormulaBox | null;
  /**
   * The element is the node drawn without the page's decorations (a hollow
   * sheet's probe): its `word-break: break-all` is the attribute's, which a
   * decoration may take back (an inert wordWrap off, Korean whole words), so
   * its runs break at words as before.
   */
  readonly undecorated?: boolean;
}

/** An in-line formula's box: its width, its height and its depth below the
 * baseline, CSS px. `exact`: a typeset box (a clone's), not the engine's. */
export interface DocsFormulaBox {
  readonly width: number;
  readonly height: number;
  readonly depth: number;
  readonly exact: boolean;
}

export function readDocsParagraph(element: HTMLElement, options: ReadDocsParagraphOptions = {}): DocsParagraph {
  const view = element.ownerDocument.defaultView;
  const flags = new Set<string>();
  const empty = (unsupported: string): DocsParagraph => ({ pieces: [], firstWidth: 0, restWidth: 0, align: "", unsupported, flags: [...flags], unscale: 1 });
  if (!view) return empty("no window");
  // A display equation on a line of its own within the paragraph: read from the page.
  if (element.querySelector("[data-math-display]")) return empty("display equation in a paragraph");
  const unlaid = options.columnWidthPx !== undefined;
  const approximateFonts = unlaid && options.approximateFonts === true;
  const style = view.getComputedStyle(element);
  if (style.writingMode && style.writingMode !== "horizontal-tb") return empty("vertical writing");
  if (style.direction === "rtl") return empty("rtl");
  // A paragraph the page does not wrap (a Word cell's `w:noWrap`) is one line
  // whatever its width: the browser draws it so.
  if (style.whiteSpace === "pre" || style.whiteSpace === "nowrap" || style.getPropertyValue("text-wrap-mode") === "nowrap") return empty("no wrap");
  const padL = px(style.paddingLeft), padR = px(style.paddingRight), borL = px(style.borderLeftWidth), borR = px(style.borderRightWidth);
  let rect: DOMRect | null = null;
  let outer: number;
  let content: number;
  let unscale: number;
  let browserWiderBy = 0;
  if (unlaid) {
    if (/%|auto/.test(`${style.marginLeft} ${style.marginRight} ${style.paddingLeft} ${style.paddingRight}`)) return empty("box not in px");
    outer = options.columnWidthPx! - px(style.marginLeft) - px(style.marginRight);
    content = outer - padL - padR - borL - borR;
    if (!(content > 0)) return empty("no room");
    unscale = 1;
  } else {
    rect = element.getBoundingClientRect();
    const borderBox = style.boxSizing === "border-box";
    const cssWidth = px(style.width);
    outer = borderBox ? cssWidth : cssWidth + padL + padR + borL + borR;
    content = borderBox ? cssWidth - padL - padR - borL - borR : cssWidth;
    // A paragraph its indents squeeze to no width (a cell narrower than they
    // are: R417-D03's "Asset’s ID") still draws, a character a line; its lines
    // are read at no width, scaled as its parent is, so a mode-14 space there
    // opens its own line as in Word (docsEngineLineStarts).
    const squeezed = (!(outer > 0) || !(rect.width > 0)) && cssWidth === 0 && rect.height > 0 && !spaceBreakOff();
    const parentElement = squeezed ? element.parentElement : null;
    const parentRect = parentElement?.getBoundingClientRect();
    if (parentElement && parentRect && parentElement.offsetWidth > 0 && parentRect.width > 0) {
      content = 0;
      unscale = parentElement.offsetWidth / parentRect.width;
    } else {
      if (!(outer > 0) || !(rect.width > 0)) return empty("not drawn");
      unscale = outer / rect.width;
      const model = cellModelContentWidth(element, style, content);
      if (model !== null) {
        browserWiderBy = content - model;
        content = model;
      }
    }
  }
  const textIndent = px(style.textIndent);
  const tabIntervalPx = style.tabSize?.endsWith("px") ? px(style.tabSize) : 0;
  const align = style.textAlign;
  const hyphenationOff = docsHyphenationOff();
  const pieces: DocsPiece[] = [];
  let unsupported: string | undefined;
  let pending = false;
  let formulaOrdinal = 0;
  if (hasPseudoContent(view, element, "::before")) flags.add("pseudo-before");
  if (hasPseudoContent(view, element, "::after")) flags.add("pseudo-after");

  // The run's line-height and its shift from the block's baseline: the
  // shifts of the inlines between it and the block add up.
  const runBox = (parent: Element, runStyle: CSSStyleDeclaration): { lineHeight?: number | "normal"; shift?: number } => {
    const lineHeight = cssLineHeight(runStyle.lineHeight);
    if (lineHeight === null) return {};
    let shift = 0;
    for (let inline: Element | null = parent; inline && inline !== element; inline = inline.parentElement) {
      const own = cssShift(inline === parent ? runStyle.verticalAlign : view.getComputedStyle(inline).verticalAlign);
      if (own === null) return {};
      shift += own;
    }
    return { lineHeight, shift };
  };

  const pushText = (node: Text, parent: Element) => {
    const raw = node.data;
    if (!raw) return;
    const runStyle = view.getComputedStyle(parent);
    if (runStyle.display === "none" || runStyle.visibility === "collapse") return;
    if (runStyle.fontVariantCaps && runStyle.fontVariantCaps !== "normal") { unsupported ??= "font-variant-caps"; return; }
    if (runStyle.fontVariantLigatures && runStyle.fontVariantLigatures !== "none" && runStyle.fontVariantLigatures !== "no-common-ligatures") flags.add("ligatures");
    if (runStyle.textTransform === "capitalize" || runStyle.textTransform === "full-width") { unsupported ??= `text-transform ${runStyle.textTransform}`; return; }
    const shrunk = parent.closest("[data-docs-justified-shrink]");
    if (shrunk) flags.add("justified-shrink");
    if (parent.closest("[data-docs-distributed-line]")) flags.add("distributed");
    const fontCss = `${runStyle.fontStyle} ${runStyle.fontWeight} ${runStyle.fontSize} ${runStyle.fontFamily}`;
    const base = pageEngineFont(fontCss);
    if (!base) { unsupported ??= "font not parsed"; return; }
    const font: EngineFont = { ...base, kerning: runStyle.fontKerning !== "none", followLocalOnly: true, ...(approximateFonts ? { approximateInstalled: true } : {}) };
    if (font.kerning) flags.add("kerning");
    // The engine's own line correction (DOCS_ENGINE_LINE_ATTR) states the
    // run's spacing without it.
    const corrected = parent.closest(`[${DOCS_ENGINE_LINE_ATTR}]`);
    const baseSpacing = corrected?.getAttribute(DOCS_ENGINE_BASE_SPACING_ATTR);
    const letterSpacing = baseSpacing != null ? px(baseSpacing) : runStyle.letterSpacing === "normal" ? 0 : px(runStyle.letterSpacing);
    const correction = corrected ? px(corrected.getAttribute(DOCS_ENGINE_LINE_ATTR)) : 0;
    if (corrected) flags.add("engine-corrected");
    // A shrink decoration's word spacing is the engine's own decision.
    const baseWordSpacing = corrected?.getAttribute(DOCS_ENGINE_BASE_WORD_SPACING_ATTR);
    const wordSpacing = baseWordSpacing != null ? px(baseWordSpacing) : shrunk || runStyle.wordSpacing === "normal" ? 0 : px(runStyle.wordSpacing);
    if (letterSpacing) flags.add("letter-spacing");
    if (wordSpacing) flags.add("word-spacing");
    const text = transformed(raw, runStyle.textTransform);
    if (text.length !== raw.length) { unsupported ??= "text-transform changes length"; return; }
    const spaceMeasure = measureText(" ", font);
    const spacePx = spaceMeasure.status === "ok" ? spaceMeasure.width + letterSpacing + wordSpacing : undefined;
    const fontPx = font.sizePx;
    const box = runBox(parent, runStyle);
    const anywhere = runStyle.wordBreak === "break-all" && !options.undecorated && !engineBreakAllOff();
    // Tabs and line feeds are pieces of their own; the text between them is
    // measured run by run.
    let at = 0;
    const flush = (end: number) => {
      if (end <= at) return;
      const chunk = text.slice(at, end);
      // An optional hyphen takes no room on its line (Word and the browser
      // draw it only where a line ends at it): the letters around it are
      // measured as one word, and it stands between them with no advance.
      const softHyphens = !hyphenationOff && chunk.includes(SOFT_HYPHEN);
      const measured = measureText(softHyphens ? chunk.split(SOFT_HYPHEN).join("") : chunk, font);
      if (measured.status === "pending") { pending = true; at = end; return; }
      if (measured.status === "unsupported") { unsupported ??= measured.reason; at = end; return; }
      if (measured.approximate) flags.add("approximate");
      let offset = at;
      const pushSoftHyphens = () => {
        while (softHyphens && offset < end && text[offset] === SOFT_HYPHEN) {
          pieces.push({ kind: "text", text: SOFT_HYPHEN, advance: 0, node, offset, fontPx, font, spacePx, letterSpacing, wordSpacing, ...(correction ? { correction } : {}), ...box });
          offset += 1;
        }
      };
      for (const grapheme of measured.graphemes) {
        pushSoftHyphens();
        const isSpace = grapheme.text === " " || grapheme.text === " ";
        const advance = grapheme.advance + (grapheme.advance > 0 || !isZeroWidth(grapheme.text) ? letterSpacing : 0) + (isSpace ? wordSpacing : 0);
        pieces.push({ kind: grapheme.text === " " ? "space" : "text", text: grapheme.text, advance, node, offset, fontPx, font, spacePx, letterSpacing, wordSpacing, ...(correction ? { correction } : {}), ...(anywhere ? { anywhere } : {}), ...box });
        offset += grapheme.text.length;
      }
      pushSoftHyphens();
      at = end;
    };
    for (let index = 0; index < text.length; index += 1) {
      const char = text[index];
      if (char === "\t" || char === "\n") {
        flush(index);
        if (char === "\n") pieces.push({ kind: "break", text: "\n", advance: 0, node, offset: index, ...(textBreakFontOff() ? {} : { breakFont: font }) });
        else if (unlaid && options.tabStops) {
          flags.add("tab");
          flags.add("tab-stops");
          // Its advance is the pen's: set as the line is filled.
          pieces.push({ kind: "tab", text: "\t", advance: 0, node, offset: index, fontPx });
        } else if (unlaid && options.approximateTabs && tabIntervalPx > 0) {
          flags.add("tab");
          flags.add("approximate");
          pieces.push({ kind: "tab", text: "\t", advance: tabIntervalPx, node, offset: index, fontPx });
        } else if (unlaid) {
          unsupported ??= "tab without layout";
        } else {
          flags.add("tab");
          // A tab on bandTabLayout's kept line is drawn as the character.
          const kept = node.parentElement?.closest(TAB_OVER_MARGIN_SELECTOR);
          const overMargin = kept && kept !== element && element.contains(kept) && !docsTabOverMarginLineOff();
          pieces.push({ kind: "tab", text: "\t", advance: drawnWidth(node, index, unscale), node, offset: index, fontPx, ...(overMargin ? { overMargin: true } : {}) });
        }
        at = index + 1;
      }
    }
    flush(text.length);
  };

  const walk = (node: Node) => {
    if (unsupported) return;
    if (node.nodeType === Node.TEXT_NODE) {
      pushText(node as Text, node.parentElement ?? element);
      return;
    }
    if (!(node instanceof Element)) return;
    if (node.tagName === "BR") {
      if (node.classList.contains("ProseMirror-trailingBreak") || node.classList.contains(DOCS_ENGINE_TAB_BREAK_CLASS)) return;
      pieces.push({ kind: "break", text: "\n", advance: 0, element: node });
      return;
    }
    if (node.tagName === "IMG" && node.classList.contains("ProseMirror-separator")) return;
    // The engine's own hyphen at a line it ends inside a word.
    if (node.classList.contains(DOCS_ENGINE_HYPHEN_CLASS)) return;
    if (node.classList.contains(DOCS_ENGINE_JOINER_CLASS)) return;
    // JustifiedLineBreaks' filler moves a manual break onto a line of its
    // own so the line before it is justified: it takes no room of a line.
    if (node.classList.contains("mrnote-justified-break-filler")) return;
    // A formatting mark shown on screen (the paragraph mark, a line-break
    // arrow): Word's marks move no line break, and the page draws them over
    // the line without taking room from it.
    if (unlaid && node.classList.contains("mrnote-formatting-mark")) return;
    const nodeStyle = view.getComputedStyle(node);
    if (nodeStyle.display === "none") return;
    if (OUT_OF_FLOW_POSITIONS.has(nodeStyle.position)) { flags.add("positioned"); return; }
    if (nodeStyle.float !== "none") { flags.add("float"); return; }
    // An in-line formula the page states no box for (typeset or waiting for
    // its clone): the engine's box stands for it.
    if (node.classList.contains("inline-math-preview")) {
      const ordinal = formulaOrdinal;
      formulaOrdinal += 1;
      if (unlaid && options.formulaBox && !node.matches(STATED_WIDTH_ATOM)) {
        const box = options.formulaBox(ordinal, node as HTMLElement);
        if (!box) { unsupported ??= "formula the engine cannot box"; return; }
        flags.add("atom");
        if (!box.exact) flags.add("engine-formula");
        pieces.push({ kind: "box", text: "\ufffc", advance: box.width, element: node, atom: { height: box.height, shift: -box.depth } });
        return;
      }
    }
    // The page's own tab layout draws a tab as a box its stop sizes: in
    // column mode it is the tab it holds, which goes to its stop here.
    if (unlaid && (options.tabStops || options.approximateTabs) && node.classList.contains(LAID_TAB_CLASS)) {
      node.childNodes.forEach(walk);
      return;
    }
    // A tab the page's tab layout has drawn (an inline-block its stop sizes)
    // is still a tab: it stays with what follows it, and what does not fit
    // after it takes it to the next line. Read as an atom it gave a break
    // after itself, and a tab that reached the right edge hung there with
    // the text after it at the margin (corpus R283-D12's references, a DOI
    // line under a typed-tab "hanging indent").
    if (!unlaid && node.classList.contains(LAID_TAB_CLASS) && ATOMIC_DISPLAYS.has(nodeStyle.display)) {
      const text = Array.from(node.childNodes).find(
        (child): child is Text => child.nodeType === Node.TEXT_NODE && (child as Text).data.includes("\t"),
      );
      flags.add("tab");
      const kept = node.closest(TAB_OVER_MARGIN_SELECTOR);
      pieces.push({
        kind: "tab",
        text: "\t",
        advance: atomWidthPx(node, nodeStyle, unscale) + px(nodeStyle.marginLeft) + px(nodeStyle.marginRight),
        element: node,
        ...(kept && kept !== element && element.contains(kept) && !docsTabOverMarginLineOff() ? { overMargin: true } : {}),
        ...(text ? { node: text, offset: text.data.indexOf("\t") } : {}),
        fontPx: px(nodeStyle.fontSize),
        ...runBox(node, nodeStyle),
      });
      return;
    }
    if (ATOMIC_DISPLAYS.has(nodeStyle.display) || REPLACED.has(node.tagName)) {
      // Only an atom's drawn box is read: reading every run's laid the page
      // out again between the editor's writes.
      if (unlaid && (node.classList.contains("docs-list-marker") || node.classList.contains("mrnote-heading-number--list"))) {
        const marker = unlaidMarkerWidth(node as HTMLElement, nodeStyle, view);
        if (marker === "pending") { pending = true; return; }
        if (marker === null) { unsupported ??= "list marker without layout"; return; }
        flags.add("list-marker");
        pieces.push({ kind: "box", text: "\ufffc", advance: marker.width, element: node, ...(marker.atom ? { atom: marker.atom } : {}), ...(marker.lineBoxes ? { lineBoxes: marker.lineBoxes } : {}) });
        return;
      }
      if (unlaid && node.matches(FIXED_INLINE_PICTURE)) {
        const picture = unlaidInlinePicture(node as HTMLElement, nodeStyle);
        if (!picture) { unsupported ??= "picture without a stated size"; return; }
        flags.add("atom");
        pieces.push({ kind: "box", text: "\ufffc", advance: picture.width, element: node, atom: picture.atom });
        return;
      }
      // A zero-size box the page states (a source page break's anchor) takes
      // no room on its line.
      if (unlaid && (node as HTMLElement).style?.width === "0px" && (node as HTMLElement).style?.height === "0px") return;
      if (unlaid && !node.matches(STATED_WIDTH_ATOM)) {
        unsupported ??= `object without layout (${node.tagName.toLowerCase()}${node.classList.length ? `.${node.classList[0]}` : ""})`;
        return;
      }
      const width = atomWidthPx(node, nodeStyle, unscale) + px(nodeStyle.marginLeft) + px(nodeStyle.marginRight);
      if (node.classList.contains("docs-list-marker")) flags.add("list-marker");
      else if (width > 0.01) flags.add("atom");
      pieces.push({ kind: "box", text: "￼", advance: width, element: node, ...placeholderAtom(node, nodeStyle) });
      return;
    }
    if (nodeStyle.display !== "inline" && nodeStyle.display !== "contents") { unsupported ??= `block ${node.tagName.toLowerCase()} inside`; return; }
    if (node.tagName === "RUBY" || node.tagName === "RT") { unsupported ??= "ruby"; return; }
    if (!INLINE_TAGS_SAFE.has(node.tagName)) flags.add(`tag-${node.tagName.toLowerCase()}`);
    if (hasPseudoContent(view, node, "::before") || hasPseudoContent(view, node, "::after")) { unsupported ??= "inline pseudo-element"; return; }
    const sideRoom = px(nodeStyle.marginLeft) + px(nodeStyle.paddingLeft) + px(nodeStyle.borderLeftWidth);
    if (sideRoom > 0.01) pieces.push({ kind: "box", text: "​", advance: sideRoom, element: node });
    const first = pieces.length;
    node.childNodes.forEach(walk);
    const endRoom = px(nodeStyle.marginRight) + px(nodeStyle.paddingRight) + px(nodeStyle.borderRightWidth);
    if (endRoom > 0.01) pieces.push({ kind: "box", text: "​", advance: endRoom, element: node });
    // The East Asian spacing projection takes a compressed opening mark's
    // room back with a negative start margin (and a centred mark's with
    // both): the run's own spacing, like the letter spacing it gives a
    // closing mark, and the browser breaks on it. Left out, every line with
    // such a mark measured up to a half-em wider here than drawn, the engine
    // broke it a character early, and its corrections never took (the
    
    if (node.hasAttribute(WORD_CJK_SPACING_ATTR) && (sideRoom < -0.01 || endRoom < -0.01)) {
      const own: number[] = [];
      for (let index = first; index < pieces.length; index += 1) {
        if (pieces[index].kind === "text" || pieces[index].kind === "space") own.push(index);
      }
      const take = (index: number | undefined, room: number) => {
        if (index === undefined || !(room < -0.01)) return;
        const piece = pieces[index];
        pieces[index] = { ...piece, advance: piece.advance + room };
      };
      take(own[0], sideRoom);
      take(own[own.length - 1], endRoom);
    }
  };
  element.childNodes.forEach(walk);

  let firstWidth = content - textIndent;
  // A leading ::before (a list label drawn by CSS) takes room on the first
  // line: read where the first piece stands.
  if (flags.has("pseudo-before") && unlaid) {
    const label = unlaidLabelAdvance(view, element);
    if (label === "pending") pending = true;
    else if (label === null) unsupported ??= "list label without layout";
    else {
      // The label opens the first line; its own line-height adds nothing.
      flags.add("list-marker");
      pieces.unshift({ kind: "box", text: "\ufffc", advance: label, atom: { height: 0, shift: 0 } });
    }
  }
  else if (flags.has("pseudo-before")) {
    if (align !== "left" && align !== "start" && align !== "justify") unsupported ??= "list label beside non-left text";
    else {
      const first = pieces.find((piece) => piece.kind !== "break");
      const firstRect = first ? pieceRect(first) : null;
      if (firstRect) {
        const contentLeft = rect!.left + (padL + borL) / unscale;
        firstWidth = content - (firstRect.left - contentLeft) * unscale;
      } else unsupported ??= "list label without text";
    }
  }
  const overflowPunctuation = element.getAttribute("data-document-effective-overflow-punctuation") === "true";
  const strutFont = pageEngineFont(`${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`);
  const strutLineHeight = cssLineHeight(style.lineHeight);
  const strut: DocsStrut | undefined = strutFont && strutLineHeight !== null
    ? { font: { ...strutFont, followLocalOnly: true, ...(approximateFonts ? { approximateInstalled: true } : {}) }, lineHeight: strutLineHeight }
    : undefined;
  const chrome = {
    top: px(style.paddingTop) + px(style.borderTopWidth),
    bottom: px(style.paddingBottom) + px(style.borderBottomWidth),
  };
  let firstLineStrut: DocsParagraph["firstLineStrut"];
  if (element.getAttribute("data-document-list-marker-strut") === "true") {
    const height = style.getPropertyValue("--document-list-marker-strut-height").trim();
    const shift = style.getPropertyValue("--document-list-marker-strut-shift").trim() || "0px";
    if (height.endsWith("px") && shift.endsWith("px")) firstLineStrut = { lineHeight: px(height), shift: px(shift) };
    else unsupported ??= "list strut not in px";
  }
  const tabStops: DocsTabLayout | undefined = unlaid && options.tabStops && flags.has("tab-stops")
    ? {
        stops: options.tabStops.stops,
        intervalTwips: tabIntervalPx > 0 ? tabIntervalPx * TWIPS_PER_CSS_PX : WORD_DEFAULT_TAB_INTERVAL_TWIPS,
        contentLeftPx: (options.tabStops.aroundLeftPx ?? 0) + px(style.marginLeft) + padL + borL,
        indentRightPx: (options.tabStops.aroundRightPx ?? 0) + px(style.marginRight) + padR + borR,
        columnPx: options.columnWidthPx! + (options.tabStops.aroundLeftPx ?? 0) + (options.tabStops.aroundRightPx ?? 0),
        stopsEndAtRightEdge: options.tabStops.overMargin !== true,
        ...(options.tabStops.overMargin !== true && !docsTabPastIndentOff() ? { stopsReachColumn: true } : {}),
      }
    : undefined;
  const inTableCell = element.closest("td, th") !== null;
  const cellWidened = inTableCell ? element.closest("table")?.matches(WIDENED_TABLE_SELECTOR) === true : undefined;
  // A drop cap: Word sets the frame's letter beside the paragraph's first
  // lines and fills each of them only as far as the room right of it
  // (Windows Word 16, probes/dropcap: R441-D01's "D", lines 2, the first two
  // lines start at the letter's advance plus its distance from text, the
  // third at the margin). The browser draws the letter as an initial letter
  // the same lines high; the engine reads its box and fills those lines
  // with that room, the letter taking none of theirs.
  // `window.__docsDropCapLinesOff = true`: the engine reads such a
  
  let leadingLines: DocsParagraph["leadingLines"];
  const dropCapOff = (view as Window & { __docsDropCapLinesOff?: boolean }).__docsDropCapLinesOff === true;
  if (!unsupported && !dropCapOff && element.matches(DROP_CAP_SELECTOR)) {
    const letter = view.getComputedStyle(element, "::first-letter");
    const lines = initialLetterLines(letter.getPropertyValue("initial-letter") || letter.getPropertyValue("-webkit-initial-letter"));
    const capIndex = pieces.findIndex((piece) => piece.kind !== "box" || piece.advance > 0.01);
    const cap = capIndex >= 0 ? pieces[capIndex] : undefined;
    const after = capIndex >= 0 ? pieces[capIndex + 1] : undefined;
    const capRect = !unlaid && lines !== null && cap?.kind === "text" ? pieceRect(cap) : null;
    if (unlaid) unsupported ??= "drop cap without layout";
    else if (lines === null) unsupported ??= "drop cap not an initial letter";
    // ::first-letter takes the punctuation around the letter with it.
    else if (!cap || cap.kind !== "text" || /\p{P}/u.test(cap.text) || (after?.kind === "text" && /\p{P}/u.test(after.text))) unsupported ??= "drop cap not one letter";
    // Drawn as an initial letter, the letter's box is wider than its own
    // advance (Chromium reports the box at the run's line height, not the
    // letter's: R441-D01's "D" 22.1 pt wide, 12 pt high at 11 pt).
    else if (!capRect || !rect || !(capRect.width * unscale > cap.advance * 1.4)) unsupported ??= "drop cap not drawn";
    else {
      const contentLeft = rect.left + (padL + borL) / unscale;
      const end = (capRect.right - contentLeft) * unscale + px(letter.marginRight);
      leadingLines = { count: lines, width: content - Math.max(0, end) };
      pieces[capIndex] = { kind: "box", text: "\ufffc", advance: 0, node: cap.node, offset: cap.offset };
      flags.add("drop-cap");
    }
  }
  return { pieces, firstWidth, restWidth: content, align, unsupported, pending, flags: [...flags], unscale, overflowPunctuation, inTableCell, ...(cellWidened === false ? { cellWidened } : {}), ...(browserWiderBy > 0 ? { browserWiderBy } : {}), ...(strut ? { strut } : {}), chrome, ...(firstLineStrut ? { firstLineStrut } : {}), ...(tabStops ? { tabStops } : {}), ...(leadingLines ? { leadingLines } : {}) };
}

/**
 * A paragraph's min-content and max-content widths as the browser lays its
 * margin box out in those modes (CSS Sizing 3 \u00a74; a table cell's
 * autofit reads them), from a paragraph read in column mode
 * (`columnWidthPx`): max-content is its longest line between hard breaks,
 * min-content its widest piece between two of the browser's break
 * opportunities, the first line's indent added to each and trailing spaces
 * hanging. Null where only the page can say (unsupported, pending, a
 * negative indent). `ideographsAsLetters` (wordSetsIdeographsAsLetters):
 * a run of ideographs and kana is one word of the min-content, as Word keeps
 * it (docx-look/probes/cjk3 tc12: mode 12, Calibri defaults, an AutoFit
 * table whose Japanese cell kept its whole run as its narrowest width and
 * crushed the "Item" column to 12.9 pt; with SimSun defaults or in mode 15
 * "Item" kept its width).
 */
export function docsContentWidths(
  paragraph: DocsParagraph,
  columnWidthPx: number,
  ideographsAsLetters = false,
): { min: number; max: number } | null {
  if (paragraph.unsupported || paragraph.pending) return null;
  const indent = paragraph.restWidth - paragraph.firstWidth;
  if (indent < 0) return null;
  const extra = columnWidthPx - paragraph.restWidth;
  const pieces = paragraph.pieces;
  const texts = pieces.map((piece) => (piece.kind === "box" ? "\ufffc" : piece.kind === "break" ? "\n" : piece.text));
  const opportunities = lineBreakOpportunities(texts, { kinsoku: null, browser: true, ideographsAsLetters });
  let max = 0;
  let min = 0;
  let line = indent;
  let lineTrailing = 0;
  let segment = indent;
  let segmentTrailing = 0;
  const closeSegment = () => {
    min = Math.max(min, segment - segmentTrailing);
    segment = 0;
    segmentTrailing = 0;
  };
  pieces.forEach((piece, index) => {
    if (piece.kind === "break") {
      max = Math.max(max, line - lineTrailing);
      closeSegment();
      line = 0;
      lineTrailing = 0;
      return;
    }
    if (index > 0 && opportunities[index] && pieces[index - 1]?.kind !== "break") closeSegment();
    const space = piece.kind === "space";
    line += piece.advance;
    segment += piece.advance;
    lineTrailing = space ? lineTrailing + piece.advance : 0;
    segmentTrailing = space ? segmentTrailing + piece.advance : 0;
  });
  max = Math.max(max, line - lineTrailing);
  closeSegment();
  return { min: extra + min, max: extra + max };
}

/**
 * A tab-suffixed list marker's width with no layout: its fitted minimum
 * (the `min-width` listDefinitionMarkers writes, in pt or px) or its label's
 * engine width if wider (an inline-block takes the larger). Its line-height
 * is 0 where the auto line-box plugin took the line over from it: then it
 * raises no line (a zero-high atom). Null for one the engine cannot read.
 */
function unlaidMarkerWidth(
  marker: HTMLElement,
  style: CSSStyleDeclaration,
  view: Window,
): { width: number; atom?: { height: number; shift: number }; lineBoxes?: DocsPiece["lineBoxes"] } | "pending" | null {
  // A heading number reaching a default tab stop is sized once laid out.
  if (marker.dataset.defaultTab === "true" && !marker.style.minWidth) return null;
  const minWidth = marker.style.minWidth;
  const minPx = !minWidth ? 0 : minWidth.endsWith("pt") ? (px(minWidth) * 4) / 3 : minWidth.endsWith("px") ? px(minWidth) : NaN;
  if (!Number.isFinite(minPx)) return null;
  const text = marker.textContent ?? "";
  const base = pageEngineFont(`${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`);
  let labelPx = 0;
  if (text) {
    if (!base) return null;
    const measured = measureText(text, { ...base, kerning: style.fontKerning !== "none", followLocalOnly: true });
    if (measured.status === "pending") return "pending";
    if (measured.status !== "ok") return null;
    const spacing = style.letterSpacing === "normal" ? 0 : px(style.letterSpacing);
    labelPx = measured.width + spacing * measured.graphemes.length;
  }
  const sidePad = px(style.paddingLeft) + px(style.paddingRight) + px(style.borderLeftWidth) + px(style.borderRightWidth);
  // `min-width` holds the padding in a border box, not in a content box.
  const box = style.boxSizing === "border-box" ? Math.max(minPx, labelPx + sidePad) : Math.max(minPx, labelPx) + sidePad;
  const width = box + px(style.marginLeft) + px(style.marginRight);
  const lineHeight = cssLineHeight(style.lineHeight);
  if (lineHeight === 0) return { width, atom: { height: 0, shift: 0 } };
  // Its one line: its own strut and each run inside that raises a line.
  if (!base || lineHeight === null) return null;
  const lineBoxes: { font: EngineFont; lineHeight: number | "normal" }[] = [{ font: { ...base, followLocalOnly: true }, lineHeight }];
  for (const child of Array.from(marker.querySelectorAll("*"))) {
    const childStyle = view.getComputedStyle(child);
    const childLineHeight = cssLineHeight(childStyle.lineHeight);
    if (childLineHeight === 0) continue;
    if (childLineHeight === null || childStyle.display !== "inline" || childStyle.verticalAlign !== "baseline") return null;
    const font = pageEngineFont(`${childStyle.fontStyle} ${childStyle.fontWeight} ${childStyle.fontSize} ${childStyle.fontFamily}`);
    if (!font) return null;
    lineBoxes.push({ font: { ...font, followLocalOnly: true }, lineHeight: childLineHeight });
  }
  return { width, lineBoxes };
}

/** An "In Line with Text" picture drawn at its stated px size (the image
 * node view's fixed-size frame), not one fitted to the page by CSS. */
const FIXED_INLINE_PICTURE = '.note-editor-image-resizable[data-inline-picture="true"][data-picture-fixed-size="true"]:not([data-page-fit-inline-image="true"])';

/**
 * Such a picture's box with no layout: the wrapper holds no line box (its
 * img is a block), so its baseline is its bottom margin edge and it sits on
 * the text baseline (note-editor.css) -- an atom as tall as the picture and
 * its margins, as wide as the picture and its side margins.
 */
function unlaidInlinePicture(wrapper: HTMLElement, style: CSSStyleDeclaration): { width: number; atom: { height: number; shift: number } } | null {
  const img = wrapper.querySelector<HTMLElement>(":scope > img");
  if (!img || !img.style.width.endsWith("px") || !img.style.height.endsWith("px")) return null;
  const width = px(img.style.width);
  const height = px(img.style.height);
  if (!(width > 0) || !(height > 0)) return null;
  if (/%|auto/.test(`${style.marginTop} ${style.marginBottom} ${style.marginLeft} ${style.marginRight}`)) return null;
  const chrome = px(style.paddingTop) + px(style.paddingBottom) + px(style.borderTopWidth) + px(style.borderBottomWidth);
  const sides = px(style.paddingLeft) + px(style.paddingRight) + px(style.borderLeftWidth) + px(style.borderRightWidth);
  return {
    width: width + sides + px(style.marginLeft) + px(style.marginRight),
    atom: { height: height + chrome + px(style.marginTop) + px(style.marginBottom), shift: 0 },
  };
}

function isZeroWidth(text: string): boolean {
  return /^[​‌‍⁠﻿­]+$/.test(text);
}

let range: Range | null = null;
function rangeOf(doc: Document): Range {
  if (!range || range.startContainer.ownerDocument !== doc) range = doc.createRange();
  return range;
}

function drawnWidth(node: Text, offset: number, unscale: number): number {
  const r = rangeOf(node.ownerDocument);
  r.setStart(node, offset);
  r.setEnd(node, offset + 1);
  const rects = r.getClientRects();
  return rects.length ? rects[0].width * unscale : 0;
}

/** The drawn box of a piece (client coordinates), or null. */
export function pieceRect(piece: DocsPiece): DOMRect | null {
  if (piece.element) {
    const rects = piece.element.getClientRects();
    return rects.length ? rects[0] : null;
  }
  if (!piece.node || piece.offset === undefined) return null;
  const r = rangeOf(piece.node.ownerDocument);
  r.setStart(piece.node, piece.offset);
  r.setEnd(piece.node, Math.min(piece.node.data.length, piece.offset + piece.text.length));
  const rects = r.getClientRects();
  for (const candidate of Array.from(rects)) if (candidate.height > 0) return candidate;
  return rects.length ? rects[0] : null;
}

const TWIPS_PER_CSS_PX = 15;

/**
 * How far the tab at `index` advances from `penPx` (column px): to the next of
 * the paragraph's stops past the pen, aligning the text up to its next tab or
 * line break there (bandTabLayout's segment).
 */
function docsTabAdvance(paragraph: DocsParagraph, tabs: DocsTabLayout, index: number, penPx: number): number {
  const pieces = paragraph.pieces;
  let widthPx = 0;
  let decimalPx: number | undefined;
  for (let at = index + 1; at < pieces.length; at += 1) {
    const piece = pieces[at];
    if (piece.kind === "tab" || piece.kind === "break") break;
    if (decimalPx === undefined && piece.text === ".") decimalPx = widthPx;
    widthPx += piece.advance;
  }
  return wordTabAdvanceAt({
    penPx,
    next: decimalPx === undefined ? { widthPx } : { widthPx, decimalPx },
    stops: tabs.stops,
    defaultIntervalTwips: tabs.intervalTwips,
    positional: null,
    line: { widthPx: tabs.columnPx, indentLeftPx: tabs.contentLeftPx, indentRightPx: tabs.indentRightPx },
    stopsEndAtRightEdge: tabs.stopsEndAtRightEdge,
    stopsReachColumn: tabs.stopsReachColumn === true,
  });
}

/**
 * Where each line starts (piece indices; the first is 0), filled the way Word
 * fills a line. A paragraph read with its stops (`tabStops`) gets each tab's
 * advance here, from where its line has got to.
 */


export function engineBreakAllOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsEngineBreakAllOff?: boolean }).__docsEngineBreakAllOff === true;
}




function textBreakFontOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsEngineBreakFontOff?: boolean }).__docsEngineBreakFontOff === true;
}

/** A letter or digit a `break-all` run may break beside (CSS Text 3: such a
 * run's letters and numbers break as ideographs do). */
const BREAK_ALL_LETTER = /^[\p{L}\p{N}]/u;

/** Harness switch (`window.__docsSpaceBreakOff = true`): a space after a
 * character wider than its line hangs in every mode, as before (A/B). */
function spaceBreakOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsSpaceBreakOff?: boolean }).__docsSpaceBreakOff === true;
}

/** Harness switch (`window.__docsLastLetterHangOff = true`): the paragraph's
 * last line of one letter wider than the measure is drawn as it comes, and
 * the paragraph's closing widgets wrap under it (A/B). */
function correctionPullUpOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsCorrectionPullUpOff?: boolean }).__docsCorrectionPullUpOff === true;
}

function pullUpTailSplitOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsPullUpTailSplitOff?: boolean }).__docsPullUpTailSplitOff === true;
}

function pullUpStickyOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsPullUpStickyOff?: boolean }).__docsPullUpStickyOff === true;
}

function lastLetterHangOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsLastLetterHangOff?: boolean }).__docsLastLetterHangOff === true;
}

/** Whether the paragraph's last line (of the engine's `starts`) is one
 * character wider than the measure: a cell narrower than a letter. Word
 * draws the letter past the edge and its paragraph mark with it, on that
 * line; a browser sends the editor's closing widgets (a cell-end mark, the
 * separator and trailing break after it) to a line of their own under it
 * (tables/probes/narrow: Word's "Asset’s ID" row is 10 lines, 108.1 pt; ours
 * drew an 11th). */
export function docsLastLineOverhangs(paragraph: DocsParagraph, starts: readonly number[]): boolean {
  const pieces = paragraph.pieces;
  if (lastLetterHangOff() || !starts.length || pieces[pieces.length - 1]?.kind === "break") return false;
  const line = starts.length - 1;
  const from = starts[line];
  const end = contentEnd(pieces, from, pieces.length);
  let engine = 0;
  let graphemes = 0;
  for (let at = from; at < end; at += 1) {
    const piece = pieces[at];
    if (piece.kind === "break") return false;
    engine += piece.advance;
    if (piece.kind === "text" || piece.kind === "space") graphemes += 1;
  }
  const width = line === 0 ? paragraph.firstWidth : paragraph.restWidth;
  return graphemes === 1 && engine > width + fitEpsilon(paragraph);
}

/** Whether line `line` of `starts` (the engine's raw starts) is opened by a
 * space, not after a hard break: a mode-14 space after a character wider than
 * the line (docsEngineLineStarts), drawn by a break before it. */
export function docsLineOpensWithSpace(paragraph: DocsParagraph, starts: readonly number[], line: number): boolean {
  if (line <= 0) return false;
  const start = starts[line];
  return paragraph.pieces[start]?.kind === "space" && paragraph.pieces[start - 1]?.kind !== "break" && paragraph.pieces[start - 1]?.kind !== "space";
}

export function docsEngineLineStarts(paragraph: DocsParagraph, rules: DocsLayoutRules): number[] {
  const pieces = paragraph.pieces;
  const count = pieces.length;
  const starts = [0];
  if (!count) return starts;
  // Opportunities over the visible sequence: zero-width atoms (editor
  // chrome) are transparent, an atom with width breaks on both sides as a
  // browser and Word break around an inline picture.
  // A letter of a break-all run is read as an ideograph, as the browser
  // breaks it: between any two, but not before a closing mark or after an
  // opening one (Windows Word, mode 14/15, print9/probes ww3: "xxx yyy..."
  // cut at the margin). Where the page sets ideographs as letters the run
  // keeps its words.
  const texts = pieces.map((piece) => (piece.kind === "box" ? "￼" : piece.kind === "break" ? "\n"
    : piece.anywhere && !rules.ideographsAsLetters && BREAK_ALL_LETTER.test(piece.text) ? "\u4e00" : piece.text));
  const visible: number[] = [];
  for (let index = 0; index < count; index += 1) {
    const piece = pieces[index];
    if (piece.kind === "box" && piece.advance <= 0.01) continue;
    if (piece.kind === "break") continue;
    visible.push(index);
  }
  const visibleOpportunities = lineBreakOpportunities(visible.map((index) => texts[index]), {
    kinsoku: rules.ideographsAsLetters || rules.eastAsianLineRules === false ? null : rules.kinsoku,
    word: true,
    ideographsAsLetters: rules.ideographsAsLetters === true,
    eastAsianMarksAsIdeographs: rules.eastAsianLineRules === false && rules.ideographsAsLetters !== true,
  });
  const opportunity = new Array<boolean>(count).fill(false);
  visible.forEach((index, order) => {
    const piece = pieces[index];
    const previous = order > 0 ? pieces[visible[order - 1]] : undefined;
    let may = visibleOpportunities[order];
    if (order > 0 && piece.kind === "box" && previous?.kind !== "space") may = true;
    if (order > 0 && previous?.kind === "box" && piece.kind !== "space") may = true;
    // A tab stays with what follows it: what does not fit after the tab takes
    // the tab to the next line, with or without a space before it, in modes
    // 14 and 15, justified or not (Word 16 for Mac, probe t-tabw, 128 lines;
    // R283-D12's references hang by exactly this).
    if (previous?.kind === "tab") may = false;
    if (order > 0 && piece.kind === "tab") may = true;
    opportunity[index] = may;
  });
  // An optional hyphen is a hyphenation point (hyphenCut), not a break
  // opportunity of its own: the line it ends draws a hyphen.
  const hyphenating = !docsHyphenationOff();
  if (hyphenating) {
    for (let index = 0; index + 1 < count; index += 1) {
      if (pieces[index].kind === "text" && pieces[index].text === SOFT_HYPHEN && pieces[index + 1].kind === "text") opportunity[index + 1] = false;
    }
  }
  /** Lines in a row ended by a hyphen so far. */
  let hyphenRun = 0;
  const justified = paragraph.align === "justify" && rules.compatibilityMode !== null && rules.compatibilityMode >= 15;
  // Word lets a stop or comma hang past the edge (wordPunctuationMayHang):
  // from mode 15 only in a justified or distributed paragraph.
  // In mode 14 and earlier every closing mark hangs, in any paragraph (Word 16.86
  // ladder: , . ; : ) ! ? % and the closing quote, justified or left), out of a table.
  const everyMarkHangs = rules.compatibilityMode !== null && rules.compatibilityMode < 15 && !paragraph.inTableCell;
  const hangs = Boolean(paragraph.overflowPunctuation) &&
    (rules.compatibilityMode === null || rules.compatibilityMode < 15 || paragraph.align === "justify" || paragraph.align === "distribute");
  const fit = fitEpsilon(paragraph);
  const spaceOpensLine = rules.compatibilityMode === null || rules.compatibilityMode < 15;
  // `window.__docsEastAsianHangOff = true` asks the older reading (only the
  // stops and commas hang, and in mode 14 a mark closing a Latin word).
  const eastAsianNoStart = rules.kinsoku && rules.eastAsianLineRules !== false && !rules.ideographsAsLetters && !eastAsianHangOff()
    ? OFFICE_KINSOKU[rules.kinsoku].noStart : null;
  let lineStart = 0;
  let lineIndex = 0;
  while (lineStart < count) {
    let width = docsLineWidth(paragraph, lineIndex);
    let used = 0;
    let pendingSpace = 0;
    let lastBreak = -1;
    let next = -1;
    let hard = false;
    let lastTab = -1;
    // Past a Word 2010 tab whose stop lies beyond the edge the line does not
    // break again (docsTabOverMarginLineOff).
    let runsOn = false;
    // Past a Word 2013 stop beyond the right indent the line fills to the
    // column's edge (docsTabPastIndentOff).
    let widened = false;
    for (let index = lineStart; index < count; index += 1) {
      const piece = pieces[index];
      if (piece.kind === "break") {
        next = index + 1;
        hard = true;
        break;
      }
      if (index > lineStart && opportunity[index]) lastBreak = index;
      if (piece.kind === "space") {
        // Mode 14 and earlier (or none): a space after a character that did
        // not fit the line by itself (a cell narrower than one letter) does
        // not hang; it opens a line of its own, and the next word another.
        // Windows Word, tables/probes/narrow n1-n6: "Asset’s ID" in a cell
        // whose text width is -2, 0 or 2 pt draws A|s|s|e|t|’|s| |I|D, ten
        // lines (R417-D03 page 9's header row, 108.3 pt); with 5 pt "s " hangs
        // as ever, and in mode 15 the space hangs at every width.
        if (!runsOn && spaceOpensLine && pendingSpace === 0 && index > lineStart && Number.isFinite(used) && used > width + fit && pieces[index - 1]?.kind === "text" && !spaceBreakOff()) {
          next = index;
          break;
        }
        pendingSpace += piece.advance;
        continue;
      }
      if (piece.kind === "tab") {
        lastTab = index;
        const tabs = paragraph.tabStops;
        if (tabs) {
          // The line's text starts at the paragraph's box, the first line at
          // its indent (a hanging one to the left).
          const lineLeft = tabs.contentLeftPx + (lineIndex === 0 ? paragraph.restWidth - paragraph.firstWidth : 0);
          const pen = lineLeft + used + pendingSpace;
          (piece as { advance: number }).advance = docsTabAdvance(paragraph, tabs, index, pen);
          if (!tabs.stopsEndAtRightEdge && !docsTabOverMarginLineOff()) {
            const stop = wordTabStopAfter(tabs.stops, pen);
            if (stop && stop.pos / TWIPS_PER_CSS_PX > tabs.columnPx - tabs.indentRightPx + 1 / TWIPS_PER_CSS_PX) {
              (piece as { overMargin?: boolean }).overMargin = true;
            }
          } else if (tabs.stopsReachColumn && !widened) {
            // Word 2013 and later: a stop past the right indent but inside the
            // column takes the line on to the column's edge (docsTabPastIndentOff).
            const stop = wordTabStopAfter(tabs.stops, pen);
            const at = stop ? stop.pos / TWIPS_PER_CSS_PX : 0;
            if (stop && tabs.indentRightPx > 0 && at > tabs.columnPx - tabs.indentRightPx + 1 / TWIPS_PER_CSS_PX && at <= tabs.columnPx + 1 / TWIPS_PER_CSS_PX) {
              width += tabs.indentRightPx;
              widened = true;
            }
          }
        }
        if (piece.overMargin) runsOn = true;
      }
      const nextUsed = used + pendingSpace + piece.advance;
      if (!runsOn && index > lineStart && nextUsed > width + fit && piece.advance > 0) {
        // Mode 14 and earlier: only the word's last character hangs, by its own
        // width, and a straight double quote not at all. Word 16 for Mac, left
        // Times New Roman 12 pt, 0.05 pt steps (fixtures-1110/hang-marks, 1 150
        // paragraphs): "data." and "data," stay 2.98 pt over and go at 3.03 (the
        // mark is 3.00); "data”" 5.31 / 5.36 (5.33); "data.”" the same
        // 5.31 / 5.36, "data)," 2.98 / 3.03 and "data”," 2.96 / 3.01 -- the
        // last mark's width, not the two marks' 8.33, 7.00 and 8.33; 'data."'
        // goes as soon as it is over. Corpus R280-D30 "(S01–S08),", R283-D02.
        const closesWord = everyMarkHangs && (() => {
          let after = index + 1;
          while (after < count && pieces[after].kind === "box" && pieces[after].advance <= 0.01) after += 1;
          return after >= count || pieces[after].kind !== "text";
        })();
        // An East Asian closing mark (one kinsoku keeps off a line's start)
        // hangs whatever follows it, in mode 14 and earlier and in a
        // justified mode-15 paragraph alike: Windows Word, zh-CN SimSun 11 pt,
        // 41 ideographs a line, each of ）。，、」”：；！？》 the 42nd, one
        // full width past the margin (docx-look/probes/cjk3 h14l/h14j/h15j;
        // h15l, left in mode 15, takes the 41st ideograph down with it).
        const eastAsianClosingMark = eastAsianNoStart !== null && piece.kind === "text" &&
          eastAsianNoStart.has(piece.text) && EAST_ASIAN_HANGING_MARK.test(piece.text) && EAST_ASIAN_PUNCTUATION.test(piece.text) &&
          (FULL_WIDTH_MARK.test(piece.text) || (index > 0 && pieces[index - 1].kind === "text" && KINSOKU_EAST_ASIAN.test(pieces[index - 1].text)));
        const hangingMark = (rules.eastAsianLineRules !== false || !FULL_WIDTH_MARK.test(piece.text)) && (eastAsianClosingMark || (everyMarkHangs
          ? closesWord && piece.kind === "text" && piece.text !== '"' && (HANGING.has(piece.text) || wordTrailingPunctuationLength(piece.text) === piece.text.length)
          : HANGING.has(piece.text)));
        if (hangs && hangingMark && (used === Number.POSITIVE_INFINITY || used + pendingSpace <= width + fit)) {
          // The mark hangs; whatever follows it opens the next line.
          used = Number.POSITIVE_INFINITY;
          pendingSpace = 0;
          continue;
        }
        // Word keeps the next word on a justified line when narrowing the
        // line's spaces is enough.
        // Only a word that reaches the piece that does not fit: past a space that
        // gives no opportunity (a closing mark, "words )", R329-D30) keeping the
        // word before it came back here for ever, and the page never answered.
        if (
          justified &&
          lastBreak > lineStart &&
          wordEnd(pieces, lastBreak, opportunity) > index &&
          keepsWord(pieces, lineStart, lastBreak, lastTab, width, opportunity, Boolean(paragraph.inTableCell))
        ) {
          const end = wordEnd(pieces, lastBreak, opportunity);
          // Continue the line past the kept word.
          let kept = 0;
          for (let at = lineStart; at < end; at += 1) if (pieces[at].kind !== "break") kept += pieces[at].advance;
          used = kept;
          pendingSpace = 0;
          index = end - 1;
          // The kept word is now part of the line: later words fit or not
          // against the line's natural width (Word narrows once per line).
          continue;
        }
        // The word that does not fit may end the line at a hyphenation
        // point (an optional hyphen, or automatic hyphenation's).
        if (hyphenating && piece.kind === "text" && Number.isFinite(used)) {
          const cut = hyphenCut(paragraph, rules, lineStart, lastBreak, lastTab, width, opportunity, justified, hyphenRun);
          if (cut >= 0) {
            next = cut + 1;
            break;
          }
        }
        next = lastBreak > lineStart ? lastBreak : index;
        break;
      }
      used = nextUsed;
      pendingSpace = 0;
    }
    if (next < 0 || next >= count) {
      if (hard && next === count) starts.push(count);
      break;
    }
    hyphenRun = hyphenating && hyphenPointKind(paragraph, next - 1) !== null && !hard ? hyphenRun + 1 : 0;
    starts.push(next);
    lineStart = next;
    lineIndex += 1;
    if (hard && next === count) break;
  }
  return starts;
}

function wordEnd(pieces: readonly DocsPiece[], from: number, opportunity: readonly boolean[]): number {
  let end = from + 1;
  while (end < pieces.length && pieces[end].kind !== "space" && pieces[end].kind !== "break" && !opportunity[end]) end += 1;
  return end;
}

function keepsWord(
  pieces: readonly DocsPiece[], lineStart: number, breakAt: number, lastTab: number, width: number, opportunity: readonly boolean[], inTableCell: boolean,
  fragment?: { readonly end: number; readonly hyphenPx: number },
): boolean {
  // A fragment: the word's letters up to a hyphenation point, the line
  // drawing the hyphen after them (judged as a whole word is, its letters
  // alone counting as the word: Windows Word, brk10/probes/hyph h1, mode 15
  // justified: "es-" of "establishment" kept 4.13 pt over and not 5.13).
  const end = fragment ? fragment.end : wordEnd(pieces, breakAt, opportunity);
  let natural = fragment ? fragment.hyphenPx : 0;
  for (let at = lineStart; at < end; at += 1) natural += pieces[at].advance;
  // The line's trailing spaces before the word count; its own spaces after
  // the word hang.
  const overflow = natural - width;
  let spaces = 0;
  let spacePx = 0;
  let fontPx = 0;
  // The spaces a line opens with take their room and narrow nothing: Word 16
  // for Mac, mode 15, Arial 12 pt, "lighthouse" after four spaces, a paragraph
  // opened by 0, 1 and 2 spaces (fixtures-1110/lead-spaces, 243 paragraphs):
  // kept to 4.118, 4.152 and 4.136 pt over and sent down 0.05 pt later, the
  // limit of the five spaces (4.167 pt) each time. Corpus R279-D13.
  let from = Math.max(lineStart, lastTab + 1);
  if (from === lineStart) while (from < breakAt && (pieces[from].kind === "space" || (pieces[from].kind === "box" && pieces[from].advance <= 0.01))) from += 1;
  for (let at = from; at < breakAt; at += 1) {
    const piece = pieces[at];
    if (piece.kind !== "space") continue;
    spaces += 1;
    spacePx = piece.spacePx ?? piece.advance;
    fontPx = piece.fontPx ?? 0;
  }
  let wordPx = 0;
  for (let at = breakAt; at < end; at += 1) wordPx += pieces[at].advance;
  let trailingPx = 0;
  // In a table cell the punctuation is paid for with the letters: Word's limit
  // there is the same with or without it (justifiedSpaceShrink: Word 16.86
  // ladder m10, corpus R280-D01; R280-D15's cell keeps "kn." off the line).
  for (let at = inTableCell ? breakAt - 1 : end - 1; at >= breakAt; at -= 1) {
    const piece = pieces[at];
    // What takes no room after the word (a bookmark's end, the paragraph's
    // closing chrome) does not hide its punctuation: corpus R280-D14, a
    // definition ending "EIR;" that Word keeps on its one line (PDF p. 3).
    if (piece.kind === "box" && piece.advance <= 0.01) continue;
    if (piece.kind !== "text" || wordTrailingPunctuationLength(piece.text) !== piece.text.length) break;
    trailingPx += piece.advance;
  }
  // Past a hyphen no space stands before the word (probe t-hy-keep).
  return wordKeepsJustifiedWord(overflow, spaces, pieces[breakAt - 1]?.kind === "space" ? fontPx : 0, spacePx, wordPx, trailingPx);
}

/** Word's room for automatic hyphenation in compatibility mode 15, px: the
 * word that does not fit is hyphenated only when moving it whole would leave
 * at least this much of a left-aligned line empty, whatever the document's
 * hyphenation zone (Windows Word, brk10/probes/hyph h2/h3: zones of 0, 283,
 * 425 and 1440 twips, 1 pt steps; brk12/probes/sq sq5, 0.1 pt steps: 17.6 pt,
 * between 17.58 and 17.64, with Calibri and Times New Roman 11 and 4 to 28
 * spaces on the line). A justified line weighs its spaces instead
 * (justifiedHyphenCut). */
const WORD_15_HYPHENATION_ROOM_PX = 17.6 * 4 / 3;
/** The rooms above are met within this, px: Word's 17.98 pt broke a word
 * (brk11/probes/h5), its widths a hair from the served faces'. */
const ROOM_TOLERANCE_PX = 0.05;

/** Whether a line may end in a hyphen after piece `at`: "soft" at an
 * optional hyphen, "auto" at an automatic hyphenation point. */
export function hyphenPointKind(paragraph: DocsParagraph, at: number): "soft" | "auto" | null {
  const piece = paragraph.pieces[at];
  if (!piece || paragraph.pieces[at + 1]?.kind !== "text" || docsHyphenationOff()) return null;
  if (piece.kind === "text" && piece.text === SOFT_HYPHEN) return "soft";
  return paragraph.hyphenation?.after.has(at) ? "auto" : null;
}

/** The hyphen a line ending after `piece` draws, px: its run's hyphen. */
export function hyphenAdvance(piece: DocsPiece): number {
  if (piece.hyphenPx !== undefined) return piece.hyphenPx;
  if (!piece.font) return 0;
  const measured = measureText("-", piece.font);
  return measured.status === "ok" ? measured.width + (piece.letterSpacing ?? 0) : 0;
}

/** The hyphen line `line` of `starts` ends in, px (0: it ends otherwise). */
export function docsLineHyphenPx(paragraph: DocsParagraph, starts: readonly number[], line: number): number {
  const next = starts[line + 1];
  if (next === undefined || next <= 0) return 0;
  return hyphenPointKind(paragraph, next - 1) ? hyphenAdvance(paragraph.pieces[next - 1]) : 0;
}

/**
 * Where a line that a word overflows ends inside that word, by Word's
 * hyphenation: the piece the line ends after, or -1 (the word goes whole).
 * The word runs from the line's last break opportunity.
 *
 * Windows Word (brk10/probes/hyph h1-h3; brk11/probes/h4, Calibri 11, 0.25
 * pt steps of the measure):
 * - Compatibility mode 15: on a justified line the last point whose
 *   fragment fits, or that narrowing the line's spaces keeps as it would a
 *   whole word (keepsWord); on any other line the last point whose
 *   fragment and hyphen fit, and an automatic point only when the line
 *   would leave at least WORD_15_HYPHENATION_ROOM_PX empty with the word
 *   moved whole.
 * - Mode 14 and earlier (12 alike), justified or not: the last point whose
 *   fragment fits with the wider of the hyphen and the letter after the point
 *   ("establish-" of "establishment" at 50.98 pt, not 45.49: the "m" after
 *   it is 8.79 pt to the hyphen's 3.37), and an automatic point only when the
 *   room after the space before the word is at least the document's
 *   hyphenation zone (0.25" unless stated: the first point at 20.73 pt, the
 *   space 2.49; zone 0 at 15.98, zone 0.5" at 38.98).
 * An optional hyphen (w:softHyphen) is a point in every paragraph: in mode
 * 15 with the same room on a line that is not justified, in mode 14 with
 * none and no letter after it to fit (brk11/probes/h5, autoHyphenation off:
 * "es-" at 17.98 pt in mode 15 left, 15.98 in mode 14; "establish-" at 45.98
 * in mode 14 where an automatic point needs 50.98).
 */
function hyphenCut(
  paragraph: DocsParagraph, rules: DocsLayoutRules, lineStart: number, lastBreak: number, lastTab: number, width: number,
  opportunity: readonly boolean[], justified: boolean, run: number,
): number {
  const pieces = paragraph.pieces;
  const wordStart = lastBreak > lineStart ? lastBreak : lineStart;
  if (pieces[wordStart]?.kind !== "text") return -1;
  const end = wordEnd(pieces, wordStart, opportunity);
  const fit = fitEpsilon(paragraph);
  const modern = rules.compatibilityMode !== null && rules.compatibilityMode >= 15;
  const hyphenation = paragraph.hyphenation;
  let content = 0;
  let all = 0;
  for (let at = lineStart; at < wordStart; at += 1) {
    const piece = pieces[at];
    if (piece.kind === "break") continue;
    all += piece.advance;
    if (piece.kind !== "space") content = all;
  }
  // Mode 15 asks the room of an optional hyphen as of an automatic one.
  const roomOf15 = justified || width - content >= WORD_15_HYPHENATION_ROOM_PX - ROOM_TOLERANCE_PX;
  const softAllowed = !modern || roomOf15;
  const autoAllowed = hyphenation !== undefined && (hyphenation.consecutiveLimit <= 0 || run < hyphenation.consecutiveLimit) &&
    (modern ? roomOf15 : width - all >= hyphenation.zonePx - ROOM_TOLERANCE_PX);
  let natural = 0;
  const through: number[] = [];
  for (let at = lineStart; at < end; at += 1) {
    if (pieces[at].kind !== "break") natural += pieces[at].advance;
    through.push(natural);
  }
  if (modern && justified && !justifiedHyphenChoiceOff()) {
    const allowed = (at: number) => {
      const kind = hyphenPointKind(paragraph, at);
      return kind !== null && (kind === "soft" || (hyphenation !== undefined && (hyphenation.consecutiveLimit <= 0 || run < hyphenation.consecutiveLimit)));
    };
    return justifiedHyphenCut(pieces, lineStart, wordStart, end, lastTab, width, content, through, fit, allowed);
  }
  for (let at = end - 2; at >= wordStart && at >= lineStart; at -= 1) {
    const kind = hyphenPointKind(paragraph, at);
    if (!kind || (kind === "auto" ? !autoAllowed : !softAllowed)) continue;
    const hyphenPx = hyphenAdvance(pieces[at]);
    const tail = modern || kind === "soft" ? hyphenPx : Math.max(hyphenPx, pieces[at + 1].advance);
    if (through[at - lineStart]! + tail <= width + fit) return at;
    if (justified && wordStart > lineStart &&
        keepsWord(pieces, lineStart, wordStart, lastTab, width, opportunity, Boolean(paragraph.inTableCell), { end: at + 1, hyphenPx })) {
      return at;
    }
  }
  return -1;
}

/**
 * Word's choice on a justified line in compatibility mode 15 when a word
 * does not fit (and narrowing the spaces does not keep it whole): end the
 * line at a hyphenation point, or move the word down. Windows Word weighs
 * how far each choice moves the line's spaces, per point of space width
 * (brk12/probes/sq sq2-sq4: Calibri 11 and 16, Arial 11, Times New Roman 11,
 * 4 to 28 spaces before the word, 0.1 pt steps of the measure, 160 measured
 * edges):
 * - moving the word down stretches the line's inner spaces (those after its
 *   last tab, up to the word, not the one before it) by r of their width;
 * - ending at a point that fits stretches all of them by r', one that does
 *   not fit narrows them by s (a quarter of a space each at most, as for a
 *   whole word);
 * - the choice costs r'^6 or (2s)^6, and a hyphen costs as much as stretching
 *   the spaces by 0.4875 of their width: Word hyphenates only when moving the
 *   word down would stretch them more than that ("es-" of "establishment"
 *   after 16 and 28 spaces went down until moving it stretched them 0.4865 to
 *   0.4880 of a space, whatever the font, the size or the stretch of the line
 *   with the hyphen), and keeps an overflowing piece against an earlier point
 *   while it narrows the spaces by at most half what the earlier point would
 *   stretch them ("estab-" against "es-": 0.499 to 0.507 of it).
 * These choose as Word did in 12,050 of 12,071 probe lines (the rest within a
 * step of 0.1 pt of the edge; the earlier reading 8,609), and in all 228
 * justified lines of optional hyphens (brk11/probes/h5; earlier 212).
 * The piece the line ends after, or -1.
 * `window.__docsJustifiedHyphenChoiceOff = true`: the earlier reading (no
 * room needed, the piece's letters kept as a whole word is).
 */
function justifiedHyphenCut(
  pieces: readonly DocsPiece[], lineStart: number, wordStart: number, end: number, lastTab: number, width: number, content: number,
  through: readonly number[], fit: number, allowed: (at: number) => boolean,
): number {
  let from = Math.max(lineStart, lastTab + 1);
  if (from === lineStart) while (from < wordStart && (pieces[from].kind === "space" || (pieces[from].kind === "box" && pieces[from].advance <= 0.01))) from += 1;
  let inner = 0;
  let pending = 0;
  for (let at = from; at < wordStart; at += 1) {
    const piece = pieces[at];
    if (piece.kind === "space") { pending += piece.spacePx ?? piece.advance; continue; }
    if (piece.kind === "box" && piece.advance <= 0.01) continue;
    inner += pending;
    pending = 0;
  }
  const all = inner + pending;
  const room = width - content;
  // A line of one word left behind stretches nothing: its word is broken.
  const down = inner > 0 ? Math.max(0, room) / inner : Number.POSITIVE_INFINITY;
  let best = -1;
  let bestCost = Number.POSITIVE_INFINITY;
  for (let at = end - 2; at >= wordStart && at >= lineStart; at -= 1) {
    if (!allowed(at)) continue;
    const over = through[at - lineStart]! + hyphenAdvance(pieces[at]) - width;
    let cost: number;
    if (over <= fit) {
      cost = all > 0 ? (Math.max(0, -over) / all) ** JUSTIFY_BADNESS_POWER : 0;
    } else {
      if (!(all > 0) || over > all * WORD_JUSTIFY_SPACE_SHARE_OF_HYPHEN + fit) continue;
      cost = (JUSTIFY_SHRINK_WEIGHT * over / all) ** JUSTIFY_BADNESS_POWER;
    }
    if (cost < bestCost) { best = at; bestCost = cost; }
    // Earlier points stretch the line more than this one that fits.
    if (over <= fit) break;
  }
  if (best < 0) return -1;
  return bestCost + JUSTIFY_HYPHEN_COST <= down ** JUSTIFY_BADNESS_POWER ? best : -1;
}

const JUSTIFY_BADNESS_POWER = 6;
const JUSTIFY_SHRINK_WEIGHT = 2;
/** What a hyphen costs: the spaces stretched by 0.4875 of their width. */
const JUSTIFY_HYPHEN_COST = 0.4875 ** JUSTIFY_BADNESS_POWER;
/** At most a quarter of each space narrowed, for a piece as for a whole word. */
const WORD_JUSTIFY_SPACE_SHARE_OF_HYPHEN = 0.25;

function justifiedHyphenChoiceOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsJustifiedHyphenChoiceOff?: boolean }).__docsJustifiedHyphenChoiceOff === true;
}

/**
 * Where the browser started each line: piece indices, from the drawn boxes.
 * A piece starts a line when its box shares no band with the line so far
 * (40 % of the smaller height, as the layout check counts lines).
 */
export function docsBrowserLineStarts(paragraph: DocsParagraph): number[] {
  const starts: number[] = [];
  let top = 0, bottom = 0;
  let open = false;
  paragraph.pieces.forEach((piece, index) => {
    if (piece.kind === "space" || piece.kind === "break") return;
    if (piece.kind === "box" && piece.advance <= 0.01) return;
    const r = pieceRect(piece);
    if (!r || !(r.height > 0)) return;
    if (!open) {
      starts.push(index);
      top = r.top; bottom = r.bottom; open = true;
      return;
    }
    const overlap = Math.min(bottom, r.bottom) - Math.max(top, r.top);
    if (overlap >= 0.4 * Math.min(bottom - top, r.bottom - r.top)) {
      top = Math.min(top, r.top); bottom = Math.max(bottom, r.bottom);
      return;
    }
    starts.push(index);
    top = r.top; bottom = r.bottom;
  });
  if (!starts.length || starts[0] !== 0) {
    // Leading spaces or chrome before the first drawn piece belong to line 1.
    if (starts.length) starts[0] = 0;
    else starts.push(0);
  }
  return starts;
}

/** Engine line starts mapped onto the drawn sequence: a start that falls on a
 * space or chrome moves to the next drawn piece, the way the browser's are
 * read. */
export function normalizeStarts(paragraph: DocsParagraph, starts: readonly number[]): number[] {
  const out: number[] = [];
  for (const start of starts) {
    let at = start;
    while (at < paragraph.pieces.length) {
      const piece = paragraph.pieces[at];
      if (piece.kind === "space" || piece.kind === "break" || (piece.kind === "box" && piece.advance <= 0.01)) at += 1;
      else break;
    }
    if (at >= paragraph.pieces.length) {
      // A line made only of a hard break's empty tail: not drawn by itself.
      continue;
    }
    if (!out.length) out.push(0);
    else if (at !== out[out.length - 1]) out.push(at);
  }
  return out.length ? out : [0];
}

/**
 * The paragraph's words (text between spaces, tabs and breaks; atoms are not
 * words) and, for each line start, the index of the word it starts with --
 * fractional when the line starts inside a word. For comparing lines with
 * another layout of the same words (Word's PDF).
 */
export function docsLineWords(paragraph: DocsParagraph, starts: readonly number[]): { words: string[]; lineWords: number[] } {
  const words: string[] = [];
  const wordOf = new Array<number>(paragraph.pieces.length).fill(-1);
  const charOf = new Array<number>(paragraph.pieces.length).fill(0);
  let open = false;
  paragraph.pieces.forEach((piece, index) => {
    if (piece.kind !== "text") {
      if (piece.kind !== "box" || piece.advance > 0.01) open = false;
      return;
    }
    if (isZeroWidth(piece.text)) return;
    if (!open) { words.push(""); open = true; }
    wordOf[index] = words.length - 1;
    charOf[index] = words[words.length - 1].length;
    words[words.length - 1] += piece.text;
  });
  const lineWords: number[] = [];
  for (const start of starts) {
    let at = start;
    while (at < paragraph.pieces.length && wordOf[at] < 0) at += 1;
    if (at >= paragraph.pieces.length) continue;
    const word = wordOf[at] + (charOf[at] > 0 ? 0.5 : 0);
    if (!lineWords.length || lineWords[lineWords.length - 1] !== word) lineWords.push(word);
  }
  return { words, lineWords };
}

export interface DocsParagraphAudit {
  readonly status: "same" | "differs" | "unsupported" | "pending";
  readonly reason?: string;
  readonly flags: readonly string[];
  readonly engineLines: number;
  readonly browserLines: number;
  /** First line (0-based) whose start differs. */
  readonly firstDiff?: number;
  /** At the first differing line: engine's start minus the browser's, in
   * pieces (negative: the engine breaks earlier). */
  readonly shift?: number;
  /** The engine's slack at that line end: width left on the line before the
   * word that went down (negative: the word overflowed by that much). */
  readonly margin?: number;
  /** Largest |drawn - engine| ink width of a line, px (text pieces only). */
  readonly inkDelta?: number;
  readonly text?: string;
  /** Text of each line by the engine and the browser at the first difference. */
  readonly at?: { engine: string; browser: string };
  /** With `detail`: the words and the word each line starts with. */
  readonly words?: string[];
  readonly engineWords?: number[];
  readonly browserWords?: number[];
}

function lineInk(paragraph: DocsParagraph, from: number, to: number): { engine: number; drawn: number } {
  let engine = 0;
  let drawn = 0;
  for (let at = from; at < to; at += 1) {
    const piece = paragraph.pieces[at];
    if (piece.kind !== "text") continue;
    const r = pieceRect(piece);
    if (!r) continue;
    engine += piece.advance;
    drawn += r.width * paragraph.unscale;
  }
  return { engine, drawn };
}

/** Engine against browser, one paragraph. */
export function auditDocsParagraph(
  element: HTMLElement, rules: DocsLayoutRules, detail = false, prepare?: (paragraph: DocsParagraph) => DocsParagraph,
): DocsParagraphAudit {
  const read = readDocsParagraph(element);
  // As the engine lays it out (its hyphenation points, docs-hyphenation).
  const paragraph = prepare ? prepare(read) : read;
  const base = { flags: paragraph.flags, engineLines: 0, browserLines: 0 };
  if (paragraph.unsupported) return { ...base, status: "unsupported", reason: paragraph.unsupported };
  if (paragraph.pending) return { ...base, status: "pending" };
  const engine = normalizeStarts(paragraph, docsEngineLineStarts(paragraph, rules));
  const browser = docsBrowserLineStarts(paragraph);
  let inkDelta = 0;
  for (let line = 0; line < browser.length; line += 1) {
    const ink = lineInk(paragraph, browser[line], browser[line + 1] ?? paragraph.pieces.length);
    inkDelta = Math.max(inkDelta, Math.abs(ink.drawn - ink.engine));
  }
  const textOf = (from: number, to: number) => paragraph.pieces.slice(from, to).map((piece) => (piece.kind === "box" ? "▫" : piece.text)).join("");
  let firstDiff = -1;
  for (let line = 0; line < Math.max(engine.length, browser.length); line += 1) {
    if (engine[line] !== browser[line]) { firstDiff = line; break; }
  }
  const words = detail ? docsLineWords(paragraph, engine) : null;
  const out = {
    ...base, engineLines: engine.length, browserLines: browser.length, inkDelta: Number(inkDelta.toFixed(3)),
    ...(words ? { words: words.words, engineWords: words.lineWords, browserWords: docsLineWords(paragraph, browser).lineWords } : {}),
  };
  if (firstDiff < 0) return { ...out, status: "same" };
  const e = engine[firstDiff] ?? paragraph.pieces.length;
  const b = browser[firstDiff] ?? paragraph.pieces.length;
  // The engine's slack where the line it ends at the earlier of the two starts.
  const lineFrom = firstDiff > 0 ? engine[firstDiff - 1] : 0;
  const width = docsLineWidth(paragraph, firstDiff > 0 ? firstDiff - 1 : 0);
  let used = 0;
  let pendingSpace = 0;
  const until = Math.max(e, b);
  for (let at = lineFrom; at < until; at += 1) {
    const piece = paragraph.pieces[at];
    if (piece.kind === "space") { pendingSpace += piece.advance; continue; }
    used += pendingSpace + piece.advance;
    pendingSpace = 0;
  }
  return {
    ...out,
    status: "differs",
    firstDiff,
    shift: e - b,
    margin: Number((width - used).toFixed(3)),
    text: textOf(0, paragraph.pieces.length).slice(0, 400),
    at: {
      engine: textOf(lineFrom, e).slice(-80),
      browser: textOf(firstDiff > 0 ? browser[firstDiff - 1] : 0, b).slice(-80),
    },
  };
}

/** Where a line's content ends: past its trailing spaces and breaks. */
function contentEnd(pieces: readonly DocsPiece[], from: number, to: number): number {
  let end = to;
  while (end > from && (pieces[end - 1].kind === "space" || pieces[end - 1].kind === "break")) end -= 1;
  return end;
}

/**
 * Each piece's natural width as the browser draws it, CSS px: text from its
 * drawn box, spaces at the engine's width (a justified line stretches them),
 * atoms and tabs as read. Read with the paragraph left-aligned when it is
 * justified, so no glyph box carries justification (callers arrange that).
 */
export function docsDrawnWidths(paragraph: DocsParagraph): number[] {
  return paragraph.pieces.map((piece) => {
    if (piece.kind !== "text") return piece.advance;
    // An optional hyphen: nothing mid-line, and at a line's end its hyphen
    // is the line's (planDocsLineCorrections adds it to both sides).
    if (piece.text === SOFT_HYPHEN && piece.advance === 0) return 0;
    const r = pieceRect(piece);
    return r ? r.width * paragraph.unscale - (piece.correction ?? 0) : piece.advance;
  });
}

/**
 * The width the browser draws pieces [from, to) at, read as ranges: each
 * stretch of text and spaces of one text node is one range (its own box, not
 * its letters' boxes summed: a letter's box is rounded out to the layout unit
 * on its own, so a line of letters summed reads wider than the browser lays
 * it out -- corpus R366-D13, ~1.1 px a line of 11 pt Calibri, 593.03 px for a
 * line of 591.93 at the engine's advances); atoms and tabs at their advance.
 * Null when a stretch is not drawn as one box (it wraps, or is not drawn).
 */
export function docsStretchDrawnWidth(paragraph: DocsParagraph, from: number, to: number): number | null {
  const pieces = paragraph.pieces;
  let total = 0;
  let at = from;
  while (at < to) {
    const first = pieces[at];
    if (first.kind === "break") { at += 1; continue; }
    if ((first.kind !== "text" && first.kind !== "space") || !first.node || first.offset === undefined) {
      total += first.advance;
      at += 1;
      continue;
    }
    let end = at + 1;
    let offsetEnd = first.offset + first.text.length;
    let correction = first.text === SOFT_HYPHEN && first.advance === 0 ? 0 : first.correction ?? 0;
    while (end < to) {
      const piece = pieces[end];
      if ((piece.kind !== "text" && piece.kind !== "space") || piece.node !== first.node || piece.offset !== offsetEnd) break;
      offsetEnd += piece.text.length;
      if (!(piece.text === SOFT_HYPHEN && piece.advance === 0)) correction += piece.correction ?? 0;
      end += 1;
    }
    const r = rangeOf(first.node.ownerDocument);
    r.setStart(first.node, first.offset);
    r.setEnd(first.node, Math.min(first.node.data.length, offsetEnd));
    const rects = Array.from(r.getClientRects()).filter((rect) => rect.height > 0);
    if (rects.length !== 1) return null;
    total += rects[0].width * paragraph.unscale - correction;
    at = end;
  }
  return total;
}

/**
 * Where the browser must be kept from breaking that the engine does not
 * break: closing punctuation before an opening bracket ("neutrino.[6]"), a
 * question mark or an ellipsis before what follows it, an em dash after what
 * is before it -- which Word keeps together and Chromium breaks (wordJoins).
 * Only in the words either side of each engine line end, where the browser
 * would break: the index of each such bracket's piece.
 */
export function docsBracketJoins(paragraph: DocsParagraph, starts: readonly number[]): number[] {
  const pieces = paragraph.pieces;
  const out: number[] = [];
  const letter = (at: number) => pieces[at]?.kind === "text";
  const scan = (from: number, to: number) => {
    for (let at = Math.max(from, 1); at < to; at += 1) {
      if (letter(at - 1) && letter(at) && wordJoins(pieces[at - 1].text, pieces[at].text, pieces[at - 2]?.text) && !out.includes(at)) out.push(at);
    }
  };
  for (let line = 1; line < starts.length; line += 1) {
    const start = starts[line];
    let end = start;
    while (end > 0 && (pieces[end - 1].kind === "space" || pieces[end - 1].kind === "break")) end -= 1;
    let wordFrom = end;
    while (wordFrom > 0 && letter(wordFrom - 1)) wordFrom -= 1;
    let wordTo = start;
    while (wordTo < pieces.length && letter(wordTo)) wordTo += 1;
    scan(wordFrom, end);
    scan(start, wordTo);
  }
  return out.sort((a, b) => a - b);
}

export interface DocsLineCorrection {
  /** Pieces [from, to) the correction spans: the line's text, without its
   * trailing spaces. */
  readonly from: number;
  readonly to: number;
  /** Extra letter spacing for every grapheme of the range, px: the drawn
   * width brought to the computed one. */
  readonly letterSpacing: number;
  /** Extra word spacing for the spaces after the line's last tab, px: a
   * justified line Word keeps by narrowing its spaces narrows them. */
  readonly wordSpacing: number;
  /** The first piece the word spacing applies from (after the last tab). */
  readonly spacesFrom: number;
  /** The tab that opens the next line (its piece index), or null. A browser
   * keeps a tab at the end of a line however full the line is (preserved
   * white space hangs past the measure, Chromium and WebKit alike), so the
   * line is ended before it by a break the engine draws there. The same
   * break goes before a space that opens the next line (a mode-14 cell
   * narrower than a letter, docsLineOpensWithSpace), and before the next
   * line's first word where the browser's box is wider than the measure
   * (`DocsParagraph.browserWiderBy`) and would keep that word. */
  readonly tabBreak: number | null;
}

/** Below this a line's drawn and computed widths count as equal, px. */
const SAME_WIDTH = 0.004;
/** What a correction leaves on the safe side of the edge, px. */
const CORRECTION_SLACK = 0.02;
/** Chromium sets each inline item's width up to a whole layout unit (1/64
 * px) before it fits items on a line, so a line of n items draws up to n/64
 * px wider than its glyphs add up to. */
const LAYOUT_UNIT = 1 / 64;

/** A grapheme's script class as the browser's run segmenter splits text
 * into items: Hangul, Han/Kana, and the rest (Latin, digits, punctuation
 * take the run they stand in). */
function scriptClass(text: string): number {
  const code = text.codePointAt(0) ?? 0;
  if ((code >= 0xac00 && code <= 0xd7af) || (code >= 0x1100 && code <= 0x11ff) || (code >= 0x3130 && code <= 0x318f)) return 1;
  if ((code >= 0x3040 && code <= 0x30ff) || (code >= 0x3400 && code <= 0x9fff) || (code >= 0xf900 && code <= 0xfaff) || (code >= 0x20000 && code <= 0x3ffff)) return 2;
  if ((code >= 0x41 && code <= 0x5a) || (code >= 0x61 && code <= 0x7a) || (code >= 0xc0 && code <= 0x24f)) return 3;
  return 0;
}

/**
 * How many inline items the browser fits a line's pieces [from, end) as: a
 * new one at every text node, every atom or tab, every script change (and
 * where a correction's narrowed spaces begin, a span of their own).
 */
export function docsLineItems(pieces: readonly DocsPiece[], from: number, end: number, spacesFrom = from): number {
  let items = 0;
  let node: Text | undefined | null = null;
  let script = 0;
  for (let at = from; at < end; at += 1) {
    const piece = pieces[at];
    if (piece.kind !== "text" && piece.kind !== "space") { items += 1; node = null; script = 0; continue; }
    const cls = piece.kind === "text" ? scriptClass(piece.text) : 0;
    const opens = node === null || piece.node !== node || (at === spacesFrom && at > from) || (cls && script && cls !== script);
    if (opens) { items += 1; script = cls; }
    else if (cls && !script) script = cls;
    node = piece.node;
  }
  return items;
}

/**
 * The corrections that make the browser break where the engine does: for
 * every engine line whose drawn width decides differently from the
 * engine's -- the line does not fit as drawn, or the next word would still
 * fit -- the drawn width is brought to the computed one by spreading the
 * difference over its graphemes' letter spacing, and a line the engine keeps
 * wider than the measure (Word's justified kept word) narrows its spaces to
 * fit. `force` corrects those lines whatever the widths say (a line the
 * browser still broke otherwise).
 */
export function planDocsLineCorrections(
  paragraph: DocsParagraph, starts: readonly number[], drawn: readonly number[], force: ReadonlySet<number> = new Set(),
  { roundingSlack = true, measure }: {
    roundingSlack?: boolean;
    /** The drawn width of pieces [from, to) read as ranges
     * (docsStretchDrawnWidth), null where it cannot be: whether the browser
     * pulls the next word up is read from it, and such a line is drawn to
     * its width from it. */
    measure?: (from: number, to: number) => number | null;
  } = {},
): DocsLineCorrection[] {
  const pieces = paragraph.pieces;
  const fit = fitEpsilon(paragraph);
  const out: DocsLineCorrection[] = [];
  for (let line = 0; line < starts.length; line += 1) {
    const from = starts[line];
    const to = starts[line + 1] ?? pieces.length;
    const end = contentEnd(pieces, from, to);
    if (end <= from) continue;
    let engine = 0;
    let browser = 0;
    let graphemes = 0;
    let lastTab = from - 1;
    let runsOn = false;
    for (let at = from; at < end; at += 1) {
      const piece = pieces[at];
      if (piece.kind === "break") continue;
      if (piece.kind === "tab") lastTab = at;
      if (piece.overMargin) runsOn = true;
      engine += piece.advance;
      browser += drawn[at];
      if (piece.kind === "text" || piece.kind === "space") graphemes += 1;
    }
    // A Word 2010 line run on past the edge by its tab: the page keeps it
    // whole as it stands, and nothing narrows it.
    if (!graphemes || runsOn) continue;
    // A line ended inside a word draws a hyphen there, as wide on both sides.
    const hyphenPx = docsLineHyphenPx(paragraph, starts, line);
    engine += hyphenPx;
    browser += hyphenPx;
    const width = docsLineWidth(paragraph, line);
    // Wider than the measure: a word kept by narrowing, on the paragraph's
    // last line too (corpus R279-D25: left as drawn, the browser sent the
    // word down and a two-line paragraph took three). That line, or one a
    // hard break ends, may overflow by a hanging mark alone: nothing narrows.
    const lastLine = to >= pieces.length || pieces[to - 1]?.kind === "break";
    const hangingPx = paragraph.overflowPunctuation && HANGING.has(pieces[end - 1].text) ? pieces[end - 1].advance : 0;
    const overflows = engine - (lastLine ? hangingPx : 0) > width + fit;
    // A line of one character wider than the measure (a cell narrower than a
    // letter) is not a kept word: Word draws the letter whole past the edge.
    // Squeezed, its letter spacing took the letter's whole width away and a
    // narrow column's letters stacked on one line (corpus R417-D03 "ID"
    // columns, tables/probes/narrow).
    // The paragraph's last such line is narrowed all the same: its letter
    // draws whole and nothing follows it but the closing widgets, which
    // then stay on its line as Word's paragraph mark does
    // (docsLastLineOverhangs).
    const closing = to >= pieces.length && pieces[to - 1]?.kind !== "break" && graphemes === 1 && !lastLetterHangOff();
    const squeeze = overflows && (graphemes > 1 || spaceBreakOff() || closing) ? engine - width : 0;
    // The browser leaves a tab on this line for as long as it fits. A justified
    // line takes up the room in its spaces, as Word's does; another is padded.
    const tabOpensNext = !lastLine && pieces[to]?.kind === "tab";
    // A browser hangs a space at a line's end however narrow the line: a
    // line the engine opens with one (mode 14, a cell narrower than a letter)
    // is ended before it by the same break.
    const spaceOpensNext = !lastLine && line + 1 < starts.length && docsLineOpensWithSpace(paragraph, starts, line + 1);
    let needed = force.has(line) || squeeze > 0 || tabOpensNext || spaceOpensNext;
    // The browser's box wider than the measure (a Word cell's text box from
    // the model): a line the engine ends whose next word the browser would
    // still fit is broken where the engine breaks it (corpus R417-D03 page
    // 36: "Mechanical harvester" 89.943 px in Word's 89.933 px, the browser's
    // 89.984 px box kept it on one line).
    let modelBreak = false;
    const wider = paragraph.browserWiderBy ?? 0;
    if (wider > 0 && !lastLine && !tabOpensNext && !spaceOpensNext && pieces[end]?.kind !== "break") {
      // Taken at the engine's advances as well as the drawn ones: each
      // grapheme's drawn box is a layout unit rounded, and summed over a word
      // they ran past the browser's box the word itself stays in (corpus
      // R417-D03: "Mechanical harvester" drawn 89.891 px whole, its boxes
      // summed past 89.984 px, so no break was planned).
      let trailing = 0;
      let trailingEngine = 0;
      for (let at = end; at < to; at += 1) {
        trailing += drawn[at];
        trailingEngine += pieces[at].advance;
      }
      let next = to;
      let nextWidth = 0;
      let nextEngine = 0;
      while (next < pieces.length && pieces[next].kind !== "space" && pieces[next].kind !== "break" && (next === to || !isOpportunityAfterSpace(pieces, next))) {
        nextWidth += drawn[next];
        nextEngine += pieces[next].advance;
        next += 1;
      }
      if (Math.min(engine + trailingEngine + nextEngine, browser + trailing + nextWidth) <= width + wider + fit) {
        modelBreak = true;
        needed = true;
      }
    }
    /** The drawn width this line must reach for the browser to leave the
     * next line's first word down, when it would pull that word up. */
    let pullUpTo = 0;
    /** The line's drawn width the correction is taken from. */
    let lineBrowser = browser;
    let rangedPull = false;
    if (!needed && Math.abs(engine - browser) > SAME_WIDTH) {
      // Does the browser keep this line whole?
      if (browser > width + fit) needed = true;
      // And leave the next line's first word down? (A hard break or the
      // paragraph's end decides by itself.)
      else if (!lastLine && pieces[end]?.kind !== "break") {
        let trailing = 0;
        for (let at = end; at < to; at += 1) trailing += drawn[at];
        let next = to;
        let nextWidth = 0;
        while (next < pieces.length && pieces[next].kind !== "space" && pieces[next].kind !== "break" && (next === to || !isOpportunityAfterSpace(pieces, next))) {
          nextWidth += drawn[next];
          next += 1;
        }
        // As ranges where the caller can read them: the letters' boxes
        // summed draw the line wider than the browser fits it. The next word
        // is read on its own, the trailing spaces at their advance: read as
        // one range with the spaces, the word was two boxes wherever the
        // browser broke before it -- as it does once this line is corrected
        // -- and the check fell back to the letters, so a line corrected on
        // one reading went uncorrected on the next, and back (corpus
        // R404-D02, page 1's seam 406 <-> 509 on every keystroke).
        // `window.__docsPullUpTailSplitOff = true`: spaces and word as one range.
        const lineDrawn = measure ? measure(from, end) : null;
        let tailDrawn: number | null = null;
        if (lineDrawn !== null && measure) {
          if (pullUpTailSplitOff()) tailDrawn = measure(end, next);
          else {
            const word = measure(to, next);
            tailDrawn = word === null ? null : trailing + word;
          }
        }
        const ranged = lineDrawn !== null && tailDrawn !== null;
        const line = ranged ? lineDrawn : browser;
        const tail = ranged ? tailDrawn ?? 0 : trailing + nextWidth;
        // A line read with its correction standing keeps it while the word
        // fits within what the reads round (each box read is up to a layout
        // unit out, and the correction's letter spacing moves where): decided
        // afresh at the edge, the reading with the correction and the one
        // without came out on either side and the line went back and forth.
        // `window.__docsPullUpStickyOff = true`: decided afresh each reading.
        let band = 0;
        if (ranged && !pullUpStickyOff() && pieces.slice(from, end).some((piece) => piece.correction)) {
          band = (docsLineItems(pieces, from, end) + 1) * LAYOUT_UNIT * Math.max(1, paragraph.unscale || 1);
        }
        if (line + tail <= width + fit + band) {
          needed = true;
          lineBrowser = line;
          if (ranged) rangedPull = true;
          if (!correctionPullUpOff()) pullUpTo = width + fit - tail;
        }
      }
    }
    if (!needed) continue;
    let spaces = 0;
    for (let at = lastTab + 1; at < end; at += 1) if (pieces[at].kind === "space") spaces += 1;
    // The safe side covers the browser's rounding of every item up to a
    // layout unit (CSS px; a page drawn smaller rounds coarser), less the
    // room the line already leaves: R342-D30's kept "no" stood 0.02 px
    // inside the measure over 16 spans and the browser sent it down.
    // `roundingSlack: false`: 0.02 px, as before.
    const rounding = docsLineItems(pieces, from, end, lastTab + 1) * LAYOUT_UNIT * Math.max(1, paragraph.unscale || 1);
    const slack = roundingSlack ? Math.max(CORRECTION_SLACK, rounding - Math.max(0, width - engine)) : CORRECTION_SLACK;
    const narrow = squeeze > 0 ? squeeze + slack : 0;
    // A line the browser would pull the next word up to: drawn at the
    // engine's width less the slack, a face the browser draws narrower than
    // the engine measures it (an installed Calibri, CoreText's advances
    // 0.14 % under the file's) still let the word up. Drawn instead past
    // where the word fits, by the slack, while the line itself still fits
    // with it: corpus R366-D13, "...to detect adrenal" (Word, 468.04 pt with
    // "insufficiency" in 468) drawn 623.77 px of 624 with the word.
    // `window.__docsCorrectionPullUpOff = true`: as before.
    // Read as ranges, the line is past that word by the browser's rounding
    // of the line's items and of the word too (each box read is a layout
    // unit out at most).
    const pullSlack = rangedPull ? Math.max(slack, rounding + LAYOUT_UNIT * Math.max(1, paragraph.unscale || 1)) : slack;
    const target = pullUpTo > 0 ? Math.min(Math.max(engine - slack, pullUpTo + pullSlack), width - slack) : engine - slack;
    // Word justifies the line a tab left: its spaces take up the room.
    const room = tabOpensNext ? Math.max(0, width - engine) : 0;
    const stretches = room > 0 && spaces > 0 && paragraph.align === "justify";
    out.push({
      from,
      to: end,
      letterSpacing: (narrow > 0 ? engine - browser - (spaces ? 0 : narrow + slack) : target - lineBrowser) / graphemes,
      wordSpacing: narrow > 0 && spaces ? -narrow / spaces : stretches ? room / spaces : 0,
      spacesFrom: lastTab + 1,
      tabBreak: tabOpensNext || spaceOpensNext || modelBreak ? to : null,
    });
  }
  return out;
}

/** A conservative word end for the next-word check: the word runs to the
 * next space (CJK text breaks at every character, so one grapheme). */
function isOpportunityAfterSpace(pieces: readonly DocsPiece[], at: number): boolean {
  return /[\u3040-\u30ff\u3400-\u9fff\uac00-\ud7af\uff00-\uffef]/.test(pieces[at].text);
}

/**
 * A quick yes to "does the browser break where the engine does?" without
 * reading every grapheme: the paragraph draws as many rows as the engine has
 * lines, and each engine line's first piece stands on a row below the piece
 * before it. False means "not shown cheaply" (read the full lines).
 */
export function docsBrowserAgreesQuickly(element: HTMLElement, paragraph: DocsParagraph, starts: readonly number[]): boolean {
  const doc = element.ownerDocument;
  const r = doc.createRange();
  r.selectNodeContents(element);
  const rects = Array.from(r.getClientRects());
  // Read: out of the page's range registry. Left there, the 14,000 of an
  // opening (R286) were visited by every node a page move took out, 1.6 ms a
  // removal and 4.5 s of the opening.
  if (!freshRangesKept()) releaseFreshRange(r);
  const rows: { top: number; bottom: number }[] = [];
  for (const rect of rects) {
    if (!(rect.height > 0.5) || !(rect.width > 0.5)) continue;
    const row = rows.find((band) => Math.min(band.bottom, rect.bottom) - Math.max(band.top, rect.top) >= 0.4 * Math.min(band.bottom - band.top, rect.bottom - rect.top));
    if (row) { row.top = Math.min(row.top, rect.top); row.bottom = Math.max(row.bottom, rect.bottom); }
    else rows.push({ top: rect.top, bottom: rect.bottom });
  }
  if (rows.length !== starts.length) return false;
  const drawnPiece = (index: number) => {
    const piece = paragraph.pieces[index];
    return piece && piece.kind !== "space" && piece.kind !== "break" && !(piece.kind === "box" && piece.advance <= 0.01);
  };
  for (let line = 1; line < starts.length; line += 1) {
    let before = starts[line] - 1;
    while (before >= 0 && !drawnPiece(before)) before -= 1;
    if (before < 0) return false;
    const a = pieceRect(paragraph.pieces[before]);
    const b = pieceRect(paragraph.pieces[starts[line]]);
    if (!a || !b) return false;
    if (Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top) >= 0.4 * Math.min(a.height, b.height)) return false;
  }
  return true;
}

/** The floats drawn in one page sheet (client rects), read once per pass. */
export type DocsFloatCache = Map<Element, DOMRect[]>;

/**
 * Whether a float (a picture or text box wrapped square or tight, a
 * float-both strip spacer's object, a floating table) narrows any of the
 * paragraph's lines: the engine fills full-width lines, so such a
 * paragraph is left to the browser's wrap around the float.
 */
export function docsParagraphBesideFloat(element: HTMLElement, cache: DocsFloatCache): boolean {
  if (element.querySelector('[data-float-wrap-spacer="true"]')) return true;
  const view = element.ownerDocument.defaultView;
  if (!view) return false;
  const sheet = element.closest(".page-sheet") ?? element.closest(".ProseMirror") ?? element.ownerDocument.body;
  let floats = cache.get(sheet);
  if (!floats) {
    floats = [];
    for (const candidate of Array.from(sheet.querySelectorAll<HTMLElement>("*"))) {
      if (candidate === element || element.contains(candidate)) continue;
      const style = view.getComputedStyle(candidate);
      if (style.float === "none" || style.display === "none") continue;
      const rect = candidate.getBoundingClientRect();
      if (rect.width > 0.5 && rect.height > 0.5) floats.push(rect);
    }
    cache.set(sheet, floats);
  }
  if (!floats.length) return false;
  const box = element.getBoundingClientRect();
  return floats.some((rect) => rect.bottom > box.top + 0.5 && rect.top < box.bottom - 0.5 && rect.right > box.left + 0.5 && rect.left < box.right - 0.5);
}
