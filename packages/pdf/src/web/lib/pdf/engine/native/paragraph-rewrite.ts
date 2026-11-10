/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The paragraph rewrite: the second tier of a native text edit.
 *
 * The surgical planner (list-text-fidelity) writes an edit into the shows the
 * source already has, and keeps every byte it can. Some edits it cannot
 * write at all: deleting a whole reader item (a trailing superscript `∗`, a
 * lone full stop in another font), growing a justified line past its frame,
 * a label drawn one glyph per show. For those, this tier saves the edit the
 * way a desktop editor does: it rewrites that one paragraph.
 *
 * - Every reader item of the paragraph is proved alone, as an unchanged edit,
 *   by the surgical planner. Those proofs name every source show that draws
 *   the paragraph: the paragraph owns exactly those shows.
 * - Every owned show is emptied where it stands, into the form of its own
 *   operator that draws nothing. A show the text after it draws on from is
 *   left as its own advance (`[n] TJ`), the way the delete command leaves it.
 * - One new text block, `q BT … ET Q`, typesets the new text right after the
 *   `ET` of the text object that holds the paragraph's last show.
 *
 * Every byte outside the owned shows' operands and the new block is kept.
 * The block keeps each character's font, size, spacing, scale, rise, slant
 * and fill colour (those of the source character it comes from; a typed character
 * takes its neighbour's), the first baseline, the line step, the first-line
 * indent, the frame's right edge and the alignment. It does not keep the
 * source's kerning inside a line: TeX's microtypography is lost here.
 *
 * Replay (native_paragraph_rewrite.py) proves every item again, empties the
 * same shows, and re-derives the layout from the same metrics.
 */
import { PDF_NATIVE_SYNTHETIC_ITALIC_SKEW } from "../../native-synthetic-style";
import { pdfNativeFaceNameStyle } from "../../native-font-identity";
import { LATIN1_EXACT } from "./latin1";
import { MRAI_NATIVE_BOX_MARK_CLOSE, MRAI_NATIVE_BOX_MARK_OPEN } from "./box-mark";
import { MRAI_NATIVE_GLYPHS_READ_AS_BLANKS_MESSAGE, MRAI_NATIVE_PARAGRAPH_READ_BACK_REFUSAL } from "./vocabulary";
import { summarizePdfNativeSelectionItems } from "./list-text-model";
import { MRAI_NATIVE_PARAGRAPH_REWRITE_SCHEMA_VERSION, readMraiNativeParagraphRewriteEvidence } from "./paragraph-rewrite-evidence";
import { MRAI_NATIVE_REORDERED_TEXT_MESSAGE, pdfTextHasRightToLeftRun } from "./reader-order";
import { PDF_SOURCE_NAME } from "./pdf-source-name";
import { pdfNativeSyntheticBoldStrokeWidth } from "../../native-synthetic-style";
import { pdfNativeTextRuleBlock, pdfNativeTextRuleBlocksBeside, pdfNativeTextRuleStandingAt } from "./text-rule-block";
import { pdfNativeTextRuleRangeIsWhole } from "../../native-text-rule-range";
export { MRAI_NATIVE_PARAGRAPH_REWRITE_SCHEMA_VERSION, readMraiNativeParagraphRewriteEvidence };
import {
  PdfNativeTransactionError,
  pdfNativeActualTextClusterSpans,
  sha256PdfBytes,
  type PdfNativeContentStreamTransactionRequest,
  type PdfNativeObjectPrecondition,
} from "./content-stream-transaction";
import {
  enclosingTextObject,
  mraiNativeEmptyChunkForm,
  mraiNativeEmptyShowPenForm,
  mraiNativeEmptyShowsForm,
  mraiNativeShowOperatorsIn,
  mraiNativeTextAfterShowIsPlacedAfresh,
} from "./text-object-delete";
import type {
  PdfNativeEditNotice,
  PdfNativeDeviceColor,
  PdfNativeFontEncoding,
  PdfNativeListTextFidelityEvidence,
  PdfNativeListTextHistoryReplacementInput,
  PdfNativeListTextHistoryReplacementPlan,
  PdfNativeListTextReplacementPlan,
  PdfNativeListTextSelectionEvidence,
  PdfNativeListTextSelectionItemEvidence,
  PdfNativeSourceTextAppearanceRun,
} from "./list-text-fidelity";
import { execBehind } from "@mrnote/core/web/lib/regexp-behind.ts";
import { pdfTextListMarkerOf } from "../../text-list-marker";


/** "distribute" (Word's 分散对齐): justified, its last line stretched too. */
export type PdfNativeParagraphRewriteAlignment = "left" | "center" | "right" | "justify" | "distribute";

export interface PdfNativeParagraphRewriteEvidence {
  readonly schemaVersion: typeof MRAI_NATIVE_PARAGRAPH_REWRITE_SCHEMA_VERSION;
  /** The content stream the paragraph is drawn in and rewritten in. */
  readonly target: Readonly<{ objectNumber: number; generation: number }>;
  /** Each reader item of the paragraph proved alone as an unchanged edit, in
   * the paragraph's order; null for a blank item the reader read into a gap
   * that no show draws. */
  readonly items: readonly (PdfNativeListTextFidelityEvidence | null)[];
  /** Every show those proofs name, in stream order: the shows the rewrite
   * empties. */
  readonly shows: readonly Readonly<{ start: number; end: number }>[];
  /** The shows the text after them in their object draws on from: each is
   * left as `[number] TJ`, its own advance, which its own proof measures. */
  readonly pens: readonly Readonly<{
    start: number;
    end: number;
    number: string;
    sourceFidelity: PdfNativeListTextFidelityEvidence;
    /** Set when the paragraph owns only this range of the show's codes: the
     * range is emptied to the number and the rest of the show stays. */
    chunk?: Readonly<{ codeStart: number; codeEnd: number }>;
    /** The character spacing the show is drawn under, when it is not zero:
     * the reader moves its pen by it once more for a number it reads with
     * no glyph before it, and the pen is written so that it moves as the
     * show did (mraiNativeEmptyShowPenForm, mraiNativeEmptyChunkForm). */
    characterSpacing?: number;
  }>[];
  




  readonly twins?: readonly Readonly<{
    start: number;
    end: number;
    owner: Readonly<{ start: number; end: number }>;
  }>[];
  /** The exact source bytes of the operator span; an undo writes them back. */
  readonly originalBytes: readonly number[];
  /** The objects the rewrite pins beside its stream (its items' fonts, the
   * page and its resource owners); an undo pins them as they stand then. */
  readonly pins: readonly Readonly<{ objectNumber: number; generation: number; role: PdfNativeObjectPrecondition["role"] }>[];
  /** The new text block and the layout it was set in; null on a read proof,
   * which proves only that the paragraph owns its shows. */
  readonly block: Readonly<{
    /** Where the block stands in the source stream: right after the `ET` of
     * the text object holding the last owned show. */
    start: number;
    bytes: readonly number[];
    alignment: PdfNativeParagraphRewriteAlignment;
    /** The paragraph's text direction in the user space the block is drawn
     * in: the unit vectors along and across its lines. */
    basis: readonly [number, number, number, number];
    frameLeft: number;
    frameRight: number;
    firstLeft: number;
    restLeft: number;
    firstBaseline: number;
    lineStep: number;
    originalLineCount: number;
    lineCount: number;
    /** Each new line's characters, as code point offsets into the new text. */
    lines: readonly Readonly<{ start: number; end: number }>[];
    /** Each new line's baseline across the lines, where a line set larger
     * than the paragraph stands lower than `firstBaseline - index *
     * lineStep` (absent when every line keeps the step). */
    lineBaselines?: readonly number[];
    /** Set when the frame's right edge is the one the reader dragged the
     * paragraph's frame to (WPS: a text block resized re-wraps its text),
     * rather than one read from the paragraph's own lines. */
    frameRequested?: true;
    /** Set when the line step is the one the reader set (line spacing)
     * rather than the paragraph's own; the evidence's lineSpacing stays the
     * paragraph's own. */
    lineStepRequested?: true;
    /** Set when the paragraph is set at the reader's size: every show's
     * size, spacing and rise by this factor (its line step too, unless one is
     * requested), in the frame the request pins. */
    sizeScale?: number;
  }> | null;
  /** Set on an exact byte restore of the named rewrite. */
  readonly history?: Readonly<{ originTransactionId: string; direction: "undo" | "redo" }>;
}

/** A paragraph's frame as the reader resized it: the right edge, as an x of
 * the page's own space (an upright paragraph; its left edge and indents
 * stay). */
export interface MraiNativeParagraphRewriteFrame {
  /** The right edge the reader dragged the frame to, an x of the page. */
  readonly right?: number;
  /** The baseline step the reader set (line spacing), in points of the
   * page across the lines. */
  readonly lineStep?: number;
  /** The factor the reader scaled the paragraph's text by (a font size set
   * on the whole paragraph); the frame is pinned by `right`. */
  readonly sizeScale?: number;
}

/** A line step set by the reader, in ems of the paragraph's largest size:
 * the bounds the replay holds it to as well. */
const REQUESTED_LINE_STEP_MIN_EM = 0.25;
const REQUESTED_LINE_STEP_MAX_EM = 20;
/** The bounds of a size the reader sets on a paragraph, as a factor of its
 * own (the replay holds it to them as well). */
const REQUESTED_SIZE_SCALE_MIN = 0.1;
const REQUESTED_SIZE_SCALE_MAX = 10;

/** A resized frame whose paragraph breaks where it did and stands where it
 * stood: nothing to write. */
export const MRAI_NATIVE_PARAGRAPH_FRAME_UNCHANGED_MESSAGE =
  "The resized frame leaves the paragraph's lines as they are.";

/** What this tier reads fonts with: the surgical planner's own metrics,
 * seen through the little this tier needs of them. */
export interface MraiNativeParagraphRewriteFont {
  readonly codeUnitBytes: 1 | 2;
  readonly decode: ReadonlyMap<number, string>;
  /** Codes the font's ToUnicode calls white space whose glyph draws
   * something (the planner's FontMetrics.inkedBlankCodes): the reader reads
   * them as spaces, so a paragraph that holds one is not the text it reads
   * as (refuseGlyphsReadAsBlanks). */
  readonly inkedBlankCodes?: ReadonlySet<number>;
  /** A simple font's subtype; a Type3 font draws glyphs as content of its
   * own and is never set again. */
  readonly simpleSubtype?: string;
  /** The face's base font name, its subset tag left off; what tells a bold
   * or italic face of the paragraph from its regular one. */
  readonly faceName?: string;
}

/** How a face is set, as its name says: bold (any heavier weight) and
 * italic (or oblique). Mirror of the replay's _face_style. */
function faceStyle(font: MraiNativeParagraphRewriteFont): readonly [number, number] {
  const style = pdfNativeFaceNameStyle(font.faceName ?? "");
  return [style.bold ? 1 : 0, style.italic ? 1 : 0];
}

/** The lean of a synthesised oblique: tan 12°, what browsers and Office
 * give a face with no italic of its own. */
// One lean for every synthesised italic, the editor's and both writers'
// (tan 12 degrees, native-synthetic-style): a rewrite that leaned a run by a
// rounder figure moved its matrix when the editor later read it upright.
const SYNTHETIC_OBLIQUE_SLANT = PDF_NATIVE_SYNTHETIC_ITALIC_SKEW;

/** How far two faces are from looking alike: one for weight, one for slant. */
function faceDistance(own: MraiNativeParagraphRewriteFont, other: MraiNativeParagraphRewriteFont): number {
  const [ownWeight, ownSlant] = faceStyle(own), [weight, slant] = faceStyle(other);
  return Math.abs(ownWeight - weight) + Math.abs(ownSlant - slant);
}

type Matrix = readonly [number, number, number, number, number, number];

/**
 * A face this tier may set a character in when no font of the paragraph has
 * it, the way a desktop editor substitutes a face for a glyph the document's
 * fonts lack. The writer installs it beside the page's own fonts; the rewrite
 * only needs its name and its metrics.
 */
export interface MraiNativeParagraphRewriteSubstitute {
  /** Called once with the characters no font of the paragraph writes, in
   * the order the new text first holds them; answers the installed face, or
   * null when it cannot draw them. */
  install(text: string): Promise<Readonly<{ resourceName: string; font: MraiNativeParagraphRewriteFont }> | null>;
  /** The face it installs (its base font name), when the writer knows it: a
   * paragraph drawn wholly in that face -- a new text box, written in a
   * subset holding only the letters typed so far -- is set wholly in it,
   * every character asked of it, rather than split across the subsets. */
  readonly faceName?: string;
}

/**
 * A face a character is set in by a format the reader asked for (a family
 * chosen for the paragraph or a part of it): a font of the paragraph, one of
 * the page's, or one the writer installs. Like the substitute, the block
 * names it and the read-back finds it by that name.
 */
export interface MraiNativeParagraphRewriteFace {
  readonly resourceName: string;
  readonly font: MraiNativeParagraphRewriteFont;
  /** The face has no bold (italic) of the weight (slant) asked: the
   * character is drawn stroked (leaned) in it, as any such face is. */
  readonly syntheticBold?: boolean;
  readonly slanted?: boolean;
}

/** How one character of the new text is set otherwise than its source
 * character: a format the reader asked for. A field left out keeps the
 * character's own. */
export interface MraiNativeParagraphRewriteCharacterFormat {
  readonly face?: MraiNativeParagraphRewriteFace;
  /** Its size as a factor of its own. */
  readonly sizeScale?: number;
  /** Its size on paper, in points (over sizeScale). */
  readonly sizePoints?: number;
  readonly color?: PdfNativeDeviceColor;
  /** Its rise (a superscript or subscript) as a share of its new size;
   * 0 sets it on the baseline. */
  readonly riseEm?: number;
  readonly syntheticBold?: boolean;
  /** Its lean (0 upright, SYNTHETIC_OBLIQUE_SLANT a synthesised italic). */
  readonly slant?: number;
  /** Its character spacing as a share of its new size. */
  readonly characterSpacingEm?: number;
  /** A rule under (through) it, in its colour (text-rule-block's rule). */
  readonly underline?: boolean;
  readonly strikethrough?: boolean;
}

/**
 * The formats an edit asks of the paragraph's new text, one entry for each
 * of its characters (null keeps the character as its source sets it), and
 * the paragraph's alignment. Set again in its frame with them, the
 * paragraph wraps where its text no longer fits the frame's width and grows
 * by lines, as a text being typed does: the frame keeps its width.
 */
/** The refusal for a bold or another face asked of a Type3 face's words. */
export const MRAI_NATIVE_TYPE3_FORMAT_MESSAGE =
  "This text is drawn in a Type3 font: its colour, size, slant, script and rules can be changed, but not its weight or font.";

export interface MraiNativeParagraphRewriteFormats {
  readonly characters: readonly (MraiNativeParagraphRewriteCharacterFormat | null)[];
  readonly alignment?: PdfNativeParagraphRewriteAlignment;
  /** Every line's left edge moved right by these points, the frame's right
   * edge kept (Increase Indent: the paragraph wraps in what is left). */
  readonly leftIndent?: number;
  /** The baseline step the lines are set at, in points (Line Spacing). */
  readonly lineStep?: number;
  /** The first line's left edge from the others', in points; negative hangs
   * it out (Word's First line / Hanging). */
  readonly firstLineIndent?: number;
}

/** What a caller resolves the formats with once the paragraph is read: the
 * proof the write starts from (the objects it adds are numbered after its
 * own) and whether the paragraph is page content. */
export type MraiNativeParagraphRewriteFormatsResolver = (context: Readonly<{
  request: PdfNativeListTextReplacementPlan["request"];
  /** The proof of the paragraph's first item (its font's page). */
  evidence: PdfNativeListTextReplacementPlan["evidence"];
  pageContent: boolean;
  text: string;
}>) => Promise<MraiNativeParagraphRewriteFormats>;

/** One show of the stream, as the surgical planner's walk reads it. */
export interface MraiNativeParagraphRewriteShow {
  readonly start: number;
  readonly end: number;
  readonly raw: Uint8Array;
  readonly codes: Uint8Array;
  readonly text: string | null;
  readonly fontName: string;
  readonly fontSize: number;
  readonly characterSpacing: number;
  readonly wordSpacing: number;
  readonly horizontalScale: number;
  readonly textRise: number;
  readonly renderingMode: number;
  readonly vertical: boolean;
  readonly deviceColor: PdfNativeDeviceColor;
  readonly deviceColorKnown: boolean;
  readonly graphicsMatrix: Matrix;
  readonly textMatrix: Matrix;
  readonly operatorKind: string;
  readonly segmentCodeCounts: readonly number[];
  readonly segmentAdjustments: readonly number[];
  readonly leadingKerningAdvance: number;
}

/**
 * The surgical planner's internals this tier reads the source with. The
 * planner hands them over at the call (it imports this module, never the
 * other way round), so the two read one source with one set of rules.
 */
export interface MraiNativeParagraphRewriteKernel {
  /** The selection in the shape a record carries: its identity checked against
   * the pinned source and each item reduced to the reader fields, without
   * whatever the box builder hung on the items for its own use. */
  exactSelection(selection: PdfNativeListTextSelectionEvidence): PdfNativeListTextSelectionEvidence;
  /** The surgical no-op proof of one reader item alone. */
  proveItem(selection: PdfNativeListTextSelectionEvidence): Promise<PdfNativeListTextReplacementPlan>;
  /** The bytes of the stream a plan edits, read the way the plan read them,
   * and whether it is one of the page's own content streams. */
  readStream(plan: PdfNativeListTextReplacementPlan): Promise<Readonly<{ streamBytes: Uint8Array; pageContent: boolean }>>;
  /** Every placed show of that stream, walked from where the plan began. */
  readShows(
    streamBytes: Uint8Array,
    plan: PdfNativeListTextReplacementPlan,
    fonts: ReadonlyMap<string, MraiNativeParagraphRewriteFont>,
  ): readonly MraiNativeParagraphRewriteShow[];
  resolveFont(
    resources: PdfNativeListTextFidelityEvidence["resources"],
    encoding: PdfNativeFontEncoding,
    witness: boolean,
  ): Promise<MraiNativeParagraphRewriteFont>;
  /** The codes that write `text` in `font`; refuses a character it lacks
   * with the planner's own missing-glyph refusal. `preferred` are the codes
   * the paragraph's own shows draw a character with, which is the answer
   * where a font names one character by several codes. */
  encodeText(
    text: string,
    font: MraiNativeParagraphRewriteFont,
    preferred?: ReadonlyMap<string, number> | null,
  ): Uint8Array;
  /** The glyphs' own widths at `fontSize`, spacing aside. */
  glyphAdvance(codes: Uint8Array, font: MraiNativeParagraphRewriteFont, fontSize: number): number;
  visualLineIndexes(items: readonly PdfNativeListTextSelectionItemEvidence[]): readonly number[];
  summarizeLines(texts: readonly string[], lineIndexes: readonly number[]): string;
  /** A glyph's characters as the reader shows them (pdf.js normalizeUnicode:
   * a font's micro sign reads as the Greek mu, a ligature as its letters). */
  normalizeUnicode(text: string): string;
  visiblePageBounds(pageNumber: number): Readonly<{ left: number; bottom: number; right: number; top: number }> | null;
  /** Where the stream a plan edits is seen, in page space (the visible page;
   * a form's box on it), with where the page's other streams start a show and
   * stand a picture: what a line grown into the room beside it may run into
   * besides its own stream's text. Null where they cannot be read. */
  streamWitness?(
    streamBytes: Uint8Array,
    plan: PdfNativeListTextReplacementPlan,
    fonts: ReadonlyMap<string, MraiNativeParagraphRewriteFont>,
  ): Readonly<{
    bounds: Readonly<{ left: number; bottom: number; right: number; top: number }>;
    others: readonly Readonly<{ x: number; y: number }>[];
    pictures: readonly Readonly<{ left: number; bottom: number; right: number; top: number }>[];
  }> | null;
  /** The numbers a show's `TJ` puts before its first glyph and after its
   * last, in thousandths of the size (the pen moves by them too). */
  showEdgeAdjustments(raw: string): number | null;
  /** The current bytes of one content stream, for a history restore. */
  inspectStream(target: Readonly<{ objectNumber: number; generation: number }>): Promise<Readonly<{
    identity: PdfNativeContentStreamTransactionRequest["identity"];
    objectSha256: string;
    streamSha256: string;
    streamBytes: Uint8Array;
  }>>;
  /** One object as it stands now, as a precondition. */
  pinObject(target: Readonly<{ objectNumber: number; generation: number }>): Promise<PdfNativeObjectPrecondition>;
}

const LATIN1 = LATIN1_EXACT;

/** A substitute face's own shows draw nothing of the paragraph yet. */
/** Computer Modern, Latin Modern, the AMS and Euler fonts by the names TeX
 * gives them, and the pack's own (lib/office-fonts/tex-pack). */
export function isTexDesignFont(faceName: string): boolean {
  return /^(?:[A-Z]{6}\+)?(?:cm[a-z]+\d|lm[a-z]+\d|ms[ab]m\d|eu[a-z]{2}\d|latinmodernmath|mrtex)/iu.test(faceName);
}

const NO_PREFERRED_CODES: ReadonlyMap<string, number> = new Map();

/** The text schema a rewritten paragraph's record carries (the planner's
 * MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION; named here, not imported, so the
 * planner stays the only module that imports the other). */
const PARAGRAPH_TEXT_SCHEMA_VERSION = "mrai.pdf-native-text.replace.v4" as const;

function refuse(message: string): never {
  throw new PdfNativeTransactionError("unsupported-document", message);
}

function asciiBytes(value: string): Uint8Array {
  return Uint8Array.from(value, (character) => character.charCodeAt(0) & 0xff);
}

/** A number as the block writes it: at most six decimals, no exponent. */
function formatNumber(value: number): string {
  if (!Number.isFinite(value) || Math.abs(value) > 1_000_000) {
    refuse("The rewritten paragraph needs a number this writer cannot write.");
  }
  const normalized = Math.abs(value) < 0.0000005 ? 0 : value;
  return normalized.toFixed(6).replace(/\.0+$/, "").replace(/(\.\d*?)0+$/, "$1");
}

/** How far apart two numbers of the page may be and still be one. */
const LAYOUT_TOLERANCE = 0.01;
/** The most a neighbouring show may already reach into the paragraph's own
 * ink and still count as abutting it, in points: two decimals of Tm rounding
 * on either side, well under a glyph (mirror of the replay's
 * _SOURCE_ABUTMENT_MAX). */
const SOURCE_ABUTMENT_MAX = 0.1;
/** Alignment is read off the lines' ends and starts to within a point. */
const ALIGNMENT_TOLERANCE = 1;
/** A space the font draws no glyph for is a third of the size, TeX's own. */
const NUMBER_SPACE_EM = 1 / 3;
/** A one-line paragraph that has to grow takes a step of 1.2 of its size. */
const SINGLE_LINE_STEP_EM = 1.2;
/** Of the room a line set larger takes, the share above its baseline (its
 * own step down grows by it; the next line's by the rest, its descent). */
const LINE_RISE_SHARE = 0.8;
const FIRST_LINE_LEADING_MAX = 1.25;
/** pdf.js reads a gap past this share of the size between two glyphs as a
 * space (pdfjs-dist legacy/build/pdf.worker.mjs TRACKING_SPACE_FACTOR; the
 * planner's own PDFJS_TRACKING_SPACE_FACTOR). */
const PDFJS_TRACKING_SPACE_FACTOR = 0.102;

/** The text characters a paragraph is laid out in: code points. */
function characters(text: string): string[] {
  return Array.from(text);
}

function isBlank(character: string): boolean {
  return character === " ";
}

/** The reader's white-space class (pdf.js `^(\s)`, written out so the
 * browser and the replay read the same set of code points). A Calibri
 * subset's word space is a tab to its ToUnicode, and the reader reads it as
 * the space between two words: matching an item to the shows that draw it
 * leaves every such glyph out, as the item's spaces are left out. What the
 * rewrite writes is another matter -- a no-break space typed into the text
 * is a character, and is set as one. */
const PDFJS_WHITESPACE_CHARACTER =
  /^[\t\n\v\f\r \u00a0\u1680\u2000-\u200a\u2028\u2029\u202f\u205f\u3000\ufeff]$/u;

function isReaderBlank(character: string): boolean {
  return PDFJS_WHITESPACE_CHARACTER.test(character);
}

const CJK = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u;

/** Full-width and CJK punctuation and forms: set like CJK characters, a
 * line may break before or after them (UAX #14's ideographic class). */
const WIDE_FORM = /[\u3000-\u303f\ufe30-\ufe4f\uff00-\uffef]/u;

/** Characters a line may not start with, and may not end with: Word's
 * default first/last-character sets for Simplified Chinese and Japanese
 * (ISO/IEC 29500 17.3.1.16 w:kinsoku), which WPS follows too. (Mirror of
 * the replay's _NO_LINE_START / _NO_LINE_END.) */
const NO_LINE_START = new Set(characters(
  "!%),.:;>?]}\u00a2\u00a8\u00b0\u00b7\u02c7\u02c9\u2015\u2016\u2019\u201d\u2026\u2030\u2032\u2033\u203a\u2103\u2236" +
  "\u3001\u3002\u3003\u3005\u3009\u300b\u300d\u300f\u3011\u3015\u3017\u301e\u309b\u309c\u309d\u309e\u30fb\u30fd\u30fe" +
  "\ufe36\ufe3a\ufe3e\ufe40\ufe44\ufe5a\ufe5c\ufe5e\uff01\uff02\uff05\uff07\uff09\uff0c\uff0e\uff1a\uff1b\uff1f" +
  "\uff3d\uff40\uff5c\uff5d\uff5e\uff61\uff63\uff64\uff65\uff9e\uff9f\uffe0"));
const NO_LINE_END = new Set(characters(
  "$([{\\\u00a3\u00a5\u00b7\u2018\u201c\u3008\u300a\u300c\u300e\u3010\u3014\u3016\u301d" +
  "\ufe59\ufe5b\ufe5d\uff04\uff08\uff0e\uff3b\uff5b\uff62\uffe1\uffe5"));

/** A combining mark: kept on the line of the character it marks. */
const COMBINING_MARK = /\p{M}/u;

/** Whether a line may break between two inked characters. The source's own
 * rule (`wide` false): only between two CJK characters. As WPS sets an
 * edited paragraph (`wide`): also where a CJK character or a full-width
 * form meets any other character, the first/last-character rules kept. */
function breaksBetween(previous: string, next: string, wide: boolean): boolean {
  // After a hyphen inside a word, as Word and WPS break "high-level": the
  // reader joins the two lines back without a space
  // (visualLineBoundaryNeedsSpace).
  if (/^[-\u2010]$/u.test(previous) && /^\p{L}$/u.test(next)) return true;
  if (!wide) return CJK.test(previous) && CJK.test(next);
  const ideographic = (character: string) => CJK.test(character) || WIDE_FORM.test(character);
  return (ideographic(previous) || ideographic(next)) && !NO_LINE_START.has(next) && !NO_LINE_END.has(previous);
}

/** Where a line of the paragraph starts (its own edge where its lines run
 * around something). */
function lineLeftOf(layout: ParagraphLayout, line: number): number {
  if (layout.lineLefts?.length) return layout.lineLefts[Math.min(line, layout.lineLefts.length - 1)];
  return line === 0 ? layout.firstLeft : layout.restLeft;
}

function applyMatrix(matrix: Matrix, x: number, y: number): readonly [number, number] {
  return [matrix[0] * x + matrix[2] * y + matrix[4], matrix[1] * x + matrix[3] * y + matrix[5]];
}

function invertMatrix(matrix: Matrix): Matrix {
  const determinant = matrix[0] * matrix[3] - matrix[1] * matrix[2];
  if (!Number.isFinite(determinant) || Math.abs(determinant) < 1e-12) {
    refuse("The paragraph is drawn under a matrix that cannot be inverted.");
  }
  const a = matrix[3] / determinant, b = -matrix[1] / determinant;
  const c = -matrix[2] / determinant, d = matrix[0] / determinant;
  return [a, b, c, d, -(a * matrix[4] + c * matrix[5]), -(b * matrix[4] + d * matrix[5])];
}

/** The user space of `to` a point of `from`'s user space lands in. */
function composeMatrix(first: Matrix, then: Matrix): Matrix {
  return [
    first[0] * then[0] + first[1] * then[2],
    first[0] * then[1] + first[1] * then[3],
    first[2] * then[0] + first[3] * then[2],
    first[2] * then[1] + first[3] * then[3],
    first[4] * then[0] + first[5] * then[2] + then[4],
    first[4] * then[1] + first[5] * then[3] + then[5],
  ];
}

/** Where one owned show is and how it is set, in the block's terms. */
/**
 * A rule another writer drew under a run (Chrome's print of a link: a thin
 * filled rectangle in the link's colour, beside the run's text object): its
 * colour, and where it lies and how thick it is across the line, as shares
 * of the run's size measured from the line's baseline (negative: below).
 */
type ParagraphLinkRule = Readonly<{ color: PdfNativeDeviceColor; offset: number; thickness: number }>;

interface OwnedShow {
  readonly start: number;
  readonly end: number;
  readonly show: MraiNativeParagraphRewriteShow;
  readonly font: MraiNativeParagraphRewriteFont;
  /** The item (paragraph order) that owns it and the visual line it is on. */
  readonly item: number;
  readonly line: number;
  /** A single-show proof of it, which a pen number is measured from. */
  readonly leaf: PdfNativeListTextFidelityEvidence | null;
  /** The range of the show's codes the paragraph owns, when the show also
   * draws text outside the paragraph (a table of contents line and its page
   * number, an equation and its number, share one show): whole segments,
   * emptied to a chunk pen while the rest of the show stays. */
  readonly partial: Readonly<{ codeStart: number; codeEnd: number }> | null;
  /** Size, spacing, scale and rise as the block writes them. */
  readonly size: number;
  readonly characterSpacing: number;
  readonly horizontalScale: number;
  readonly rise: number;
  /** How far its up axis leans along the line per unit of height: 0 upright,
   * the tangent of a synthetic oblique's angle otherwise. */
  readonly slant: number;
  readonly color: PdfNativeDeviceColor;
  /** Its origin along and across the paragraph's lines, and where its first
   * and last inked glyph start and end along them; null when it inks none. */
  readonly u: number;
  readonly v: number;
  readonly inked: Readonly<{ start: number; end: number; count: number }> | null;
  /** Where each of those inked glyphs starts and ends along the lines, in
   * the order the show draws them: the paragraph's own ink as the source
   * sets it, which the line check measures other text against. */
  readonly ink: readonly (readonly [number, number])[];
  /** Where the page draws this show again as a fake-bold pass, along and
   * across the lines from the show itself; null where it draws it once. */
  readonly twin: readonly [number, number] | null;
  /** Drawn filled and stroked (`2 Tr`): a bold made by stroking the glyphs
   * (this writer's own synthetic bold, a word processor's fake bold). The
   * rewrite draws it so again, its stroke in its fill colour. */
  readonly stroked: boolean;
  /** Ruled under (through) by a rule this writer drew beside its text
   * object (pdfNativeTextRuleBlock): the rewrite takes the rule away and
   * draws it again under the characters where they go. */
  readonly underline?: boolean;
  readonly strikethrough?: boolean;
  /** Underlined by a rule another writer drew beside it (a web print's link
   * underline: a thin filled rectangle of its own colour): taken away with
   * the run and drawn again under its characters where they go. */
  readonly linkRule?: ParagraphLinkRule;
}

interface ParagraphLayout {
  readonly alignment: PdfNativeParagraphRewriteAlignment;
  readonly frameLeft: number;
  readonly frameRight: number;
  readonly firstLeft: number;
  readonly restLeft: number;
  /** Each line's own left edge, where a flush-left paragraph's lines start
   * at more than one (it runs around a picture): a line set again keeps the
   * edge of the line it is, one past them the last line's. */
  readonly lineLefts?: readonly number[];
  readonly firstBaseline: number;
  readonly lineStep: number;
  readonly lineCount: number;
  readonly lineExtents: readonly Readonly<{ left: number; right: number }>[];
  /** The size of the paragraph's body text (its show that inks the most
   * glyphs): what a line's characters reaching higher or lower are measured
   * against (lineReachOf). */
  readonly bodySize: number;
  /** How far the source's first line's characters reached above its body
   * (lineReachOf): `firstBaseline` stands that much higher, as if they had
   * not lowered it, and the line keeps that lowering when set again. */
  readonly firstReach?: number;
}

/** A line's characters reach past its body's this much before the line is
 * set lower for them: a footnote mark (2/3 size, raised a third) does not
 * move its line. */
const LINE_REACH_FLOOR = 0.15;

/** How far a line's characters reach above its baseline and below it past
 * the body text's own (0.8 of a size above, 0.2 below, a rise moving both),
 * none under LINE_REACH_FLOOR of the body size. */
/**
 * The leading a first line's reach past the body text is spaced at. A paragraph
 * spaced widely (lines 40 pt apart at 14 pt) has that room between its lines,
 * not above them: at its own leading a word set at 72 stood a whole size under
 * the top of the box, with nothing above it.
 */
function firstLineLeading(lineStep: number, bodySize: number) {
  return bodySize > 0 ? Math.min(lineStep / bodySize, FIRST_LINE_LEADING_MAX) : 1;
}

function lineReachOf(characters: readonly Readonly<{ size: number; rise: number }>[], bodySize: number) {
  const above = Math.max(0, ...characters.map((character) => LINE_RISE_SHARE * character.size + character.rise)) -
    LINE_RISE_SHARE * bodySize;
  const below = Math.max(0, ...characters.map((character) => (1 - LINE_RISE_SHARE) * character.size - character.rise)) -
    (1 - LINE_RISE_SHARE) * bodySize;
  const floor = LINE_REACH_FLOOR * bodySize;
  return { above: above > floor ? above : 0, below: below > floor ? below : 0 };
}

interface ReadParagraph {
  readonly selection: PdfNativeListTextSelectionEvidence;
  readonly items: readonly PdfNativeListTextSelectionItemEvidence[];
  readonly plans: readonly (PdfNativeListTextReplacementPlan | null)[];
  readonly firstPlan: PdfNativeListTextReplacementPlan;
  readonly streamBytes: Uint8Array;
  readonly source: string;
  readonly pageContent: boolean;
  readonly shows: readonly OwnedShow[];
  /** For each character of the paragraph's text, the owned show it is drawn
   * by (an index into `shows`). */
  readonly characterShows: readonly number[];
  readonly basis: readonly [number, number, number, number];
  readonly blockMatrix: Matrix;
  readonly layout: ParagraphLayout;
  /** The paragraph's own line step, when `layout` is set at one the reader
   * asked for (paragraphAtRequestedLineStep). */
  readonly sourceLineStep?: number;
  /** The factor its shows are scaled by, when set at the reader's size
   * (paragraphAtRequestedSize). */
  readonly sizeScale?: number;
  readonly pens: PdfNativeParagraphRewriteEvidence["pens"];
  readonly spanStart: number;
  /** The rules readParagraph read as its shows' formats: taken away with
   * them (rewrittenSpan). */
  readonly rules: readonly Readonly<{ start: number; end: number }>[];
  readonly insertion: number;
  readonly pins: readonly PdfNativeObjectPrecondition[];
  readonly allShows: readonly MraiNativeParagraphRewriteShow[];
  /** The fake-bold passes of the owned shows (see the evidence's `twins`),
   * in stream order, and their spans as keys: emptied with their owners,
   * and no other text of the page. */
  readonly twins: readonly Readonly<{ start: number; end: number; owner: Readonly<{ start: number; end: number }> }>[];
  readonly twinKeys: ReadonlySet<string>;
  readonly fonts: ReadonlyMap<string, MraiNativeParagraphRewriteFont>;
  /** The paragraph's own fonts in the order its shows first use them: the
   * order a character no font of its own can write tries them in. */
  readonly fontOrder: readonly string[];
  /** Per font, the code each character the paragraph draws is drawn with. */
  readonly preferred: ReadonlyMap<string, ReadonlyMap<string, number>>;
  /** What the paragraph's lines leave between their words, as a share of
   * the line's size; null where no line of it has a gap to measure. */
  readonly gapRatio: number | null;
  /** By character of the paragraph's text: a blank standing between two shows
   * on one line (a TeX line's words and formulas are each their own), what the
   * line leaves there, along it. */
  readonly keptGaps: ReadonlyMap<number, number>;
  /** By character of the paragraph's text: a blank the reader read for a run
   * of more white space than one character that one show draws ("7  矢量",
   * "Fixture  |  07"), the length of that run (paragraphBlankRuns). */
  readonly blankRuns: ReadonlyMap<number, number>;
  /** Where the source inks a character of the paragraph's text along the lines. */
  readonly characterInk: ReadonlyMap<number, readonly [number, number]>;
  /** By character of the paragraph's text: the rules this writer drew under
   * (through) part of its run, read as those characters' own (readParagraph). */
  readonly characterRules: ReadonlyMap<number, Readonly<{ underline?: true; strikethrough?: true }>>;
  /** Where the reader left a gap on a line unread: for each, the index in
   * the paragraph's text of the character after the gap (see readParagraph). */
  readonly unreadGaps: readonly number[];
  /** Each of those gaps' width along the line, as the page draws it. */
  readonly unreadGapWidths: readonly number[];
  /** The frame is one the reader dragged (paragraphInRequestedFrame), not
   * the paragraph's own. */
  readonly frameRequested?: true;
}

/** The shows one item's proof names, each with the reader text it draws. */
function proofShows(
  evidence: PdfNativeListTextFidelityEvidence,
): readonly Readonly<{
  start: number;
  end: number;
  text: string;
  /** The show's text as its own codes spell it (its runs of blanks whole). */
  spelled: string;
  chunk: Readonly<{ codeStart: number; codeEnd: number }> | null;
  leaf: PdfNativeListTextFidelityEvidence | null;
  resourceName: string;
}>[] {
  if (evidence.positionedLines || evidence.paragraphReflow || evidence.consolidatedRuns ||
    evidence.propertyMutation || evidence.positionedRestore || evidence.paragraphRewrite) {
    refuse("A reader item of this paragraph is proved by another paragraph writer.");
  }
  const showEnd = (leaf: PdfNativeListTextFidelityEvidence) => {
    const span = leaf.textOperatorSpan, end = leaf.lineShift?.showEnd;
    return end !== undefined && span.start < end && end < span.end ? end : span.end;
  };
  if (evidence.positionedRunGroup) {
    return evidence.positionedRunGroup.runs.map((run) => Object.freeze({
      start: run.evidence.textOperatorSpan.start,
      end: showEnd(run.evidence),
      text: run.evidence.selection?.item.str || run.evidence.originalDisplayText || run.evidence.originalText,
      spelled: run.evidence.originalDisplayText || run.evidence.originalText,
      chunk: run.evidence.readerChunk ?? null,
      leaf: run.evidence.positionedRunGroup || run.evidence.linePlan ? null : run.evidence,
      resourceName: run.evidence.resourceName,
    }));
  }
  if (evidence.linePlan) {
    return evidence.linePlan.lines.map((line) => Object.freeze({
      start: line.operatorSpan.start,
      end: line.operatorSpan.end,
      text: line.originalText,
      spelled: line.originalText,
      chunk: null,
      leaf: null,
      resourceName: line.resourceName ?? evidence.resourceName,
    }));
  }
  return [Object.freeze({
    start: evidence.textOperatorSpan.start,
    end: showEnd(evidence),
    text: evidence.selection?.item.str || evidence.originalDisplayText || evidence.originalText,
    spelled: evidence.originalDisplayText || evidence.originalText,
    chunk: evidence.readerChunk ?? null,
    leaf: evidence,
    resourceName: evidence.resourceName,
  })];
}

/** Every font a proof reads, with whether it may be written in. */
function proofFonts(
  evidence: PdfNativeListTextFidelityEvidence,
  into: Map<string, { resources: PdfNativeListTextFidelityEvidence["resources"]; encoding: PdfNativeFontEncoding; editable: boolean }>,
): void {
  const add = (name: string, resources: PdfNativeListTextFidelityEvidence["resources"],
    encoding: PdfNativeFontEncoding, editable: boolean) => {
    const known = into.get(name);
    if (known) {
      if (JSON.stringify(known.resources) !== JSON.stringify(resources) ||
        JSON.stringify(known.encoding) !== JSON.stringify(encoding)) {
        refuse("The paragraph's proofs name one font resource two ways.");
      }
      known.editable ||= editable;
      return;
    }
    into.set(name, { resources, encoding, editable });
  };
  add(evidence.resourceName, evidence.resources, evidence.fontEncoding, true);
  for (const run of evidence.positionedRunGroup?.runs ?? []) proofFonts(run.evidence, into);
  for (const line of evidence.linePlan?.lines ?? []) {
    if (line.resourceName && line.resources && line.fontEncoding) {
      add(line.resourceName, line.resources, line.fontEncoding, true);
    }
  }
  for (const witness of evidence.layoutWitnesses ?? []) {
    add(witness.resourceName, witness.resources, witness.fontEncoding, false);
  }
}

/**
 * Along one show's own x axis (text space, before its scale): where each
 * glyph starts and where the next would, and the codes it draws. The pen
 * moves by every number of a `TJ`, the ones before the first glyph included.
 */
function showGlyphs(
  kernel: MraiNativeParagraphRewriteKernel,
  show: MraiNativeParagraphRewriteShow,
  font: MraiNativeParagraphRewriteFont,
): readonly Readonly<{ code: number; start: number; end: number; text: string | undefined }>[] {
  const unit = font.codeUnitBytes;
  const glyphs: { code: number; start: number; end: number; text: string | undefined }[] = [];
  // The segments are the strings the show holds, in bytes; a character code
  // is one or two of them.
  const counts = show.segmentCodeCounts.length ? show.segmentCodeCounts : [show.codes.length];
  let pen = show.leadingKerningAdvance;
  let offset = 0;
  for (const [segment, count] of counts.entries()) {
    if (count % unit !== 0) refuse("A show of this paragraph splits a character code.");
    for (let index = 0; index < count; index += unit) {
      const bytes = show.codes.slice(offset, offset + unit);
      if (bytes.length !== unit) refuse("A show of this paragraph splits a character code.");
      offset += unit;
      const code = unit === 1 ? bytes[0] : (bytes[0] << 8) | bytes[1];
      const width = kernel.glyphAdvance(bytes, font, show.fontSize) + show.characterSpacing +
        (unit === 1 && code === 0x20 ? show.wordSpacing : 0);
      glyphs.push({ code, start: pen, end: pen + width, text: font.decode.get(code) });
      pen += width;
    }
    if (segment < counts.length - 1) {
      pen += (-(show.segmentAdjustments[segment] ?? 0) * show.fontSize) / 1_000;
    }
  }
  if (offset !== show.codes.length) refuse("A show of this paragraph does not read as its codes.");
  return glyphs;
}

/**
 * Items no show proves alone, read as the shows of their font that stand on
 * their baseline and spell them, each a reader item of its own with the
 * origin and width the reader gives it; each is proved alone. An item that
 * cannot be read so is refused as it was.
 */
async function itemsAsTheirShows(
  kernel: MraiNativeParagraphRewriteKernel,
  selection: PdfNativeListTextSelectionEvidence,
  items: readonly PdfNativeListTextSelectionItemEvidence[],
  plans: readonly (PdfNativeListTextReplacementPlan | null)[],
  unproved: ReadonlyMap<number, string>,
): Promise<Readonly<{ items: PdfNativeListTextSelectionItemEvidence[]; plans: (PdfNativeListTextReplacementPlan | null)[] }>> {
  const [firstIndex, firstMessage] = [...unproved.entries()][0];
  const refused = (index: number, message: string): never =>
    refuse(`Reader item ${index + 1} of the paragraph does not prove alone: ${message}`);
  const proved = plans.find((plan) => plan !== null);
  if (!proved) return refused(firstIndex, firstMessage);
  const fontEvidence = new Map<string, { resources: PdfNativeListTextFidelityEvidence["resources"]; encoding: PdfNativeFontEncoding; editable: boolean }>();
  for (const plan of plans) if (plan) proofFonts(plan.evidence, fontEvidence);
  const fonts = new Map<string, MraiNativeParagraphRewriteFont>();
  for (const [name, entry] of fontEvidence) fonts.set(name, await kernel.resolveFont(entry.resources, entry.encoding, !entry.editable));
  const { streamBytes } = await kernel.readStream(proved);
  const shows = kernel.readShows(streamBytes, proved, fonts);
  const flat = (text: string) => kernel.normalizeUnicode(text).replace(/\s+/gu, "");
  const outItems: PdfNativeListTextSelectionItemEvidence[] = [];
  const outPlans: (PdfNativeListTextReplacementPlan | null)[] = [];
  for (const [index, item] of items.entries()) {
    const message = unproved.get(index);
    if (message === undefined) {
      outItems.push(item);
      outPlans.push(plans[index]);
      continue;
    }
    const [a, b] = item.transform;
    const scale = Math.hypot(a, b);
    if (!(scale > 0)) return refused(index, message);
    const along = (x: number, y: number) => ((x - item.transform[4]) * a + (y - item.transform[5]) * b) / scale;
    const across = (x: number, y: number) => ((y - item.transform[5]) * a - (x - item.transform[4]) * b) / scale;
    const tolerance = Math.max(0.5, 0.05 * scale);
    // Each show as the reader would read it alone: its first glyph's origin
    // and its glyphs' advance, on the page.
    const pieces: { at: number; item: PdfNativeListTextSelectionItemEvidence }[] = [];
    for (const show of shows) {
      const font = fonts.get(show.fontName);
      if (!font || show.vertical || show.text === null || !show.text.trim()) continue;
      let glyphs: ReturnType<typeof showGlyphs>;
      try { glyphs = showGlyphs(kernel, show, font); } catch { continue; }
      if (!glyphs.length) continue;
      const toPage = composeMatrix(show.textMatrix, show.graphicsMatrix);
      const horizontal = show.horizontalScale / 100;
      const [ox, oy] = [glyphs[0].start * horizontal, show.textRise];
      const x = ox * toPage[0] + oy * toPage[2] + toPage[4];
      const y = ox * toPage[1] + oy * toPage[3] + toPage[5];
      const at = along(x, y);
      const width = (glyphs.at(-1)!.end - glyphs[0].start) * horizontal * Math.hypot(toPage[0], toPage[1]);
      // On the item's baseline, starting inside it and ending by its end (a
      // show's blanks after its last ink may run past the item: the reader
      // puts them in the next, a line's ") " before a bullet).
      let inked = glyphs.length - 1;
      while (inked > 0 && !glyphs[inked].text?.trim()) inked -= 1;
      const inkedWidth = (glyphs[inked].end - glyphs[0].start) * horizontal * Math.hypot(toPage[0], toPage[1]);
      if (Math.abs(across(x, y)) > tolerance || at < -tolerance || at >= item.width - tolerance ||
        at + inkedWidth > item.width + tolerance) continue;
      const size = composeMatrix([show.fontSize * horizontal, 0, 0, show.fontSize, 0, show.textRise], toPage);
      // The reader's height is the length of the show's own up axis (a
      // slanted show's longer than its size).
      pieces.push({ at, item: Object.freeze({ ...item, str: kernel.normalizeUnicode(show.text), width,
        height: Math.hypot(size[2], size[3]),
        transform: Object.freeze([size[0], size[1], size[2], size[3], x, y]) as unknown as PdfNativeListTextSelectionItemEvidence["transform"],
        hasEOL: false }) });
    }
    pieces.sort((left, right) => left.at - right.at);
    if (pieces.length < 2 || flat(pieces.map((piece) => piece.item.str).join("")) !== flat(item.str)) {
      return refused(index, message);
    }
    for (const [at, piece] of pieces.entries()) {
      try {
        outPlans.push(await kernel.proveItem(Object.freeze({
          sourceSha256: selection.sourceSha256,
          selectionRevision: selection.selectionRevision,
          pageNumber: selection.pageNumber,
          startOffset: 0, endOffset: piece.item.str.length, item: piece.item,
        })));
      } catch (error) {
        if (!(error instanceof PdfNativeTransactionError)) throw error;
        return refused(index, `${message} (its show ${at + 1} of ${pieces.length}: ${error.message})`);
      }
      outItems.push(at === pieces.length - 1 ? Object.freeze({ ...piece.item, hasEOL: item.hasEOL }) : piece.item);
    }
  }
  return { items: outItems, plans: outPlans };
}

/**
 * The paragraph a selection names, read and owned: every reader item proved
 * alone, the shows those proofs name, where each is on the page and how it is
 * set, and the paragraph's original layout.
 */
async function readParagraph(
  kernel: MraiNativeParagraphRewriteKernel,
  named: PdfNativeListTextSelectionEvidence,
): Promise<ReadParagraph> {
  const selection = kernel.exactSelection(named);
  let items = selection.items ?? [selection.item];
  if (items.length < 1 || items.length > 256) refuse("Paragraph rewrite item count is unsupported.");
  let lineIndexes = kernel.visualLineIndexes(items);
  // Read as the box reads its items (the space a word gap stands for
  // before a run set off the line's baseline included).
  const originalText = selection.items
    ? summarizePdfNativeSelectionItems(items)
    : selection.item.str;
  if (originalText !== selection.item.str || selection.startOffset !== 0 ||
    selection.endOffset !== selection.item.str.length) {
    refuse("Paragraph rewrite needs the whole paragraph's text.");
  }
  const plans: (PdfNativeListTextReplacementPlan | null)[] = [];
  const unproved = new Map<number, string>();
  for (const [index, item] of items.entries()) {
    try {
      const plan = await kernel.proveItem(Object.freeze({
        sourceSha256: selection.sourceSha256,
        selectionRevision: selection.selectionRevision,
        pageNumber: selection.pageNumber,
        startOffset: 0,
        endOffset: item.str.length,
        item,
      }));
      plans.push(plan);
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
      // The reader reads a space into a wide gap no show draws; such an
      // item owns nothing and its space is laid out like any other.
      if (/^ +$/u.test(item.str)) {
        plans.push(null);
        continue;
      }
      unproved.set(index, error.message);
      plans.push(null);
    }
  }
  // An item the reader joined from several shows of one font -- an upright
  // run and a faux italic one set by its own slanted text matrix (this
  // writer's, WPS's), the reader breaking items only where the font or its
  // size changes -- proves no show alone: it is read as those shows, each
  
  // left the paragraph refused, and the typing after it was lost).
  if (unproved.size) {
    const split = await itemsAsTheirShows(kernel, selection, items, plans, unproved);
    items = split.items;
    plans.splice(0, plans.length, ...split.plans);
    lineIndexes = kernel.visualLineIndexes(items);
  }
  const firstPlan = plans.find((plan) => plan !== null);
  if (!firstPlan) refuse("No reader item of the paragraph draws text.");
  const sameWrite = (plan: PdfNativeListTextReplacementPlan) =>
    plan.request.identity.objectNumber === firstPlan.request.identity.objectNumber &&
    plan.request.identity.generation === firstPlan.request.identity.generation &&
    plan.request.expectedObjectSha256 === firstPlan.request.expectedObjectSha256 &&
    plan.request.expectedStreamSha256 === firstPlan.request.expectedStreamSha256 &&
    JSON.stringify(plan.request.pageContentConsolidation ?? null) ===
      JSON.stringify(firstPlan.request.pageContentConsolidation ?? null) &&
    JSON.stringify(plan.request.formInstanceIsolation ?? null) ===
      JSON.stringify(firstPlan.request.formInstanceIsolation ?? null);
  if (!plans.every((plan) => plan === null || sameWrite(plan))) {
    refuse("The paragraph's items are drawn in more than one content stream.");
  }
  if (firstPlan.request.formInstanceIsolation) {
    refuse("Paragraph rewrite does not write into a shared form.");
  }
  // The shows every proof names, and the item each belongs to.
  const claims = new Map<string, {
    start: number; end: number; item: number; leaf: PdfNativeListTextFidelityEvidence | null;
    chunks: Readonly<{ codeStart: number; codeEnd: number }>[] | null; resourceName: string;
  }>();
  const itemShowTexts: (readonly Readonly<{ key: string; text: string; spelled: string }>[])[] = [];
  // The shows each item's proof names, and the reader chunk of each it is.
  const itemParts: (readonly Readonly<{ key: string; chunk: Readonly<{ codeStart: number; codeEnd: number }> | null }>[])[] = [];
  const fontEvidence = new Map<string, { resources: PdfNativeListTextFidelityEvidence["resources"]; encoding: PdfNativeFontEncoding; editable: boolean }>();
  for (const [index, plan] of plans.entries()) {
    if (!plan) {
      itemShowTexts.push([]);
      itemParts.push([]);
      continue;
    }
    proofFonts(plan.evidence, fontEvidence);
    const shows = proofShows(plan.evidence);
    const texts: { key: string; text: string; spelled: string }[] = [];
    const parts: { key: string; chunk: Readonly<{ codeStart: number; codeEnd: number }> | null }[] = [];
    for (const show of shows) {
      if (!(show.start < show.end)) refuse("A reader item's proof names an empty show.");
      const key = `${show.start}:${show.end}`;
      const claimed = claims.get(key);
      if (claimed) {
        // One show the reader cut into several items at wide gaps: each
        // chunk item claims its codes, and every one is in the paragraph.
        if (!claimed.chunks || !show.chunk || claimed.item === index) {
          refuse("Two reader items of the paragraph claim one show.");
        }
        claimed.chunks.push(show.chunk);
      } else {
        claims.set(key, { start: show.start, end: show.end, item: index, leaf: show.leaf,
          chunks: show.chunk ? [show.chunk] : null, resourceName: show.resourceName });
      }
      // As the reader shows it: a show of one ligature glyph (Chromium sets
      // one per show) reads as its letters, U+FB01 as "fi".
      texts.push({ key, text: kernel.normalizeUnicode(show.text), spelled: kernel.normalizeUnicode(show.spelled) });
      parts.push({ key, chunk: show.chunk ?? null });
    }
    itemShowTexts.push(texts);
    itemParts.push(parts);
  }
  const ordered = [...claims.values()].sort((left, right) => left.start - right.start);
  for (const [index, show] of ordered.entries()) {
    if (index > 0 && ordered[index - 1].end > show.start) refuse("The paragraph's shows overlap.");
  }
  const { streamBytes, pageContent } = await kernel.readStream(firstPlan);
  const source = LATIN1.decode(streamBytes);
  const fonts = new Map<string, MraiNativeParagraphRewriteFont>();
  const editable = new Set<string>();
  for (const [name, entry] of fontEvidence) {
    fonts.set(name, await kernel.resolveFont(entry.resources, entry.encoding, !entry.editable));
    if (entry.editable) editable.add(name);
  }
  const allShows = kernel.readShows(streamBytes, firstPlan, fonts);
  const bySpan = new Map(allShows.map((show) => [`${show.start}:${show.end}`, show]));
  // The block is drawn where the last owned show's text object ends, in the
  // user space in force there.
  const lastClaim = ordered.at(-1)!;
  const lastShow = bySpan.get(`${lastClaim.start}:${lastClaim.end}`);
  if (!lastShow) refuse("The paragraph's last show is not where its proof puts it.");
  const blockMatrix = lastShow.graphicsMatrix;
  const toBlock = invertMatrix(blockMatrix);
  // The paragraph's text direction: its first item's first show's axes, in
  // the block's user space.
  const firstClaim = claims.get(itemShowTexts.find((texts) => texts.length > 0)![0].key)!;
  const firstShow = bySpan.get(`${firstClaim.start}:${firstClaim.end}`);
  if (!firstShow) refuse("The paragraph's first show is not where its proof puts it.");
  const linearOf = (show: MraiNativeParagraphRewriteShow): Matrix =>
    composeMatrix(composeMatrix(show.textMatrix, show.graphicsMatrix), toBlock);
  const firstLinear = linearOf(firstShow);
  const alongLength = Math.hypot(firstLinear[0], firstLinear[1]);
  const acrossLength = Math.hypot(firstLinear[2], firstLinear[3]);
  if (!(alongLength > 0) || !(acrossLength > 0)) refuse("The paragraph's text space is degenerate.");
  // Across the lines runs the first show's own up axis, or -- where a design
  // tool set that show slanted by its text matrix (a faux italic,
  // `9.5 0 2.0193 9.5 … Tm`) -- the direction square to its baseline: the
  // slant is that show's own, never the paragraph's.
  const firstSlanted = Math.abs(firstLinear[0] * firstLinear[2] + firstLinear[1] * firstLinear[3]) >
    1e-6 * alongLength * acrossLength;
  // The block's user space may itself be flipped: Chromium draws every page
  // under `.24 0 0 -.24 0 H cm` and sets its text upright on the page by a
  // text matrix flipped back (`1 0 0 -1 x y Tm`). Such text runs left-handed
  // in the block, and the rewrite sets its lines by the same basis, so its
  // Tm comes out flipped the same way. Only text the page shows mirrored is
  // refused.
  const firstHand = Math.sign(firstLinear[0] * firstLinear[3] - firstLinear[1] * firstLinear[2]);
  const basis = (firstSlanted
    ? [firstLinear[0] / alongLength, firstLinear[1] / alongLength,
      -firstHand * firstLinear[1] / alongLength, firstHand * firstLinear[0] / alongLength]
    : [firstLinear[0] / alongLength, firstLinear[1] / alongLength,
      firstLinear[2] / acrossLength, firstLinear[3] / acrossLength]) as readonly [number, number, number, number];
  if (Math.abs(basis[0] * basis[2] + basis[1] * basis[3]) > 1e-6 || firstHand === 0 ||
    Math.sign(basis[0] * basis[3] - basis[1] * basis[2]) !== firstHand ||
    firstHand * (blockMatrix[0] * blockMatrix[3] - blockMatrix[1] * blockMatrix[2]) <= 0) {
    refuse("The paragraph's text is mirrored.");
  }
  const lineOfItem = (item: number) => lineIndexes[item];
  const shows: OwnedShow[] = [];
  const indexByKey = new Map<string, number>();
  const itemInk = new Map<number, Readonly<{ start: number; pen: number; readerSize: number }>[]>();
  for (const claim of ordered) {
    const show = bySpan.get(`${claim.start}:${claim.end}`);
    if (!show) refuse("A show of this paragraph is not where its proof puts it.");
    if (show.text === null) refuse("A show of this paragraph stands under a replacement text.");
    if (show.vertical) refuse("Paragraph rewrite sets horizontal text only.");
    // Filled, or filled and stroked (a synthetic bold: a paragraph given one
    // by this writer could not be rewritten again -- every later edit of it,
    
    // clipped or invisible text is not set again.
    if (show.renderingMode !== 0 && show.renderingMode !== 2) refuse("A show of this paragraph is not drawn filled.");
    if (!show.deviceColorKnown) refuse("A show of this paragraph has a fill colour this writer cannot read.");
    if (!editable.has(show.fontName) || show.fontName !== claim.resourceName) {
      refuse("A show of this paragraph is drawn in a font its proof does not name.");
    }
    const font = fonts.get(show.fontName)!;
    const glyphs = showGlyphs(kernel, show, font);
    let partial: Readonly<{ codeStart: number; codeEnd: number }> | null = null;
    if (claim.chunks) {
      // Every glyph of a show the reader cut into chunks is in one of them
      // (their ranges are byte offsets into the show's codes), or is a blank
      // the reader left out between them -- or the paragraph owns only part
      // of the show: the reader cut it at wide gaps and only some pieces are
      // the paragraph's (a table of contents line and its page number, an
      // equation and its number, a manual's line and the word after an inline
      // icon). That part is one contiguous range of whole character codes;
      // the rest of the show stays, drawn on from where the part's pen would
      // have ended.
      const unit = font.codeUnitBytes;
      const covered = new Set<number>();
      for (const chunk of claim.chunks) {
        for (let byte = chunk.codeStart; byte < chunk.codeEnd; byte += 1) covered.add(byte);
      }
      const drawn = (index: number) => !/^\s*$/u.test(glyphs[index].text ?? "x");
      if (glyphs.some((_, index) => !covered.has(index * unit) && drawn(index))) {
        const codeStart = Math.min(...claim.chunks.map((chunk) => chunk.codeStart));
        const codeEnd = Math.max(...claim.chunks.map((chunk) => chunk.codeEnd));
        // The range need not be whole segments: the reader cuts a show at a
        // run of spaces inside one string too, and the emptied form keeps
        // the rest of that string (mraiNativeEmptyChunkForm). It is whole
        // character codes.
        if (codeStart % unit !== 0 || codeEnd % unit !== 0 || !(codeStart < codeEnd) ||
          glyphs.some((_, index) => index * unit >= codeStart && index * unit < codeEnd && !covered.has(index * unit) && drawn(index))) {
          refuse("A show of this paragraph also draws text outside it.");
        }
        partial = Object.freeze({ codeStart, codeEnd });
      }
    }
    // A mark the reader reads with the glyph before it (a zero-width
    // diacritic: TeX's `\not`, drawn a kern after the empty set and over the
    // equals sign after it, reads "∅̸ =") is set by this tier where that glyph
    // ends, as it drops the kerning inside a line. One the source draws
    // anywhere else -- past a number of its array, or opening a show of its
    // own -- would come off the glyph it is drawn over, and the "≠" would be
    // a slash beside an "=". (Mirror of the replay's owned shows.)
    for (const [index, glyph] of glyphs.entries()) {
      if (glyph.text === undefined || !PDFJS_ZERO_WIDTH_DIACRITIC.test(glyph.text)) continue;
      if (partial && !(index * font.codeUnitBytes >= partial.codeStart && index * font.codeUnitBytes < partial.codeEnd)) continue;
      if (index === 0 || Math.abs(glyph.start - glyphs[index - 1].end) > 1e-9 * show.fontSize) {
        refuse("A mark of this paragraph is drawn off the glyph the rewrite would set it after.");
      }
    }
    const linear = linearOf(show);
    const along = Math.hypot(linear[0], linear[1]), fullAcross = Math.hypot(linear[2], linear[3]);
    if (!(along > 0) || !(fullAcross > 0) ||
      Math.abs(linear[0] * basis[1] - linear[1] * basis[0]) > 1e-6 * along ||
      linear[0] * basis[0] + linear[1] * basis[1] <= 0 || linear[2] * basis[2] + linear[3] * basis[3] <= 0) {
      refuse("The paragraph's shows are not drawn along one direction.");
    }
    // A show set slanted (its up axis leaning along the line) keeps its
    // slant: the rewrite draws it under a text matrix leaning the same way,
    // and its size is its height square to the line. An upright show is
    // measured as it always was.
    const leaning = Math.abs(linear[2] * basis[3] - linear[3] * basis[2]) > 1e-6 * fullAcross;
    const across = leaning ? linear[2] * basis[2] + linear[3] * basis[3] : fullAcross;
    const slant = leaning ? (linear[2] * basis[0] + linear[3] * basis[1]) / across : 0;
    const [x, y] = applyMatrix(toBlock, ...applyMatrix(show.graphicsMatrix, ...applyMatrix(show.textMatrix, 0, 0)));
    const u = x * basis[0] + y * basis[1], v = x * basis[2] + y * basis[3];
    const scale = (show.horizontalScale / 100) * along;
    const inkedGlyphs = glyphs.filter((glyph, index) => !/^\s*$/u.test(glyph.text ?? "x") &&
      (!partial || (index * font.codeUnitBytes >= partial.codeStart && index * font.codeUnitBytes < partial.codeEnd)));
    const inked = inkedGlyphs.length
      ? Object.freeze({ start: u + inkedGlyphs[0].start * scale, end: u + inkedGlyphs.at(-1)!.end * scale,
        count: inkedGlyphs.length })
      : null;
    const ink = Object.freeze(inkedGlyphs.map((glyph) =>
      Object.freeze([u + glyph.start * scale, u + glyph.end * scale] as const)));
    // Where each item that claims the show inks it, as the reader measures
    // the gap after it: from its first inked glyph to the pen after its
    // last (the reader's prevTransform, short of the character spacing, and
    // the word spacing, that follow that glyph), with the show's size along
    // the line (the reader's tracking bound is the font size in the text
    // matrix's own units: pdf.js textAdvanceScale).
    for (const [itemIndex, parts] of itemParts.entries()) {
      for (const part of parts) {
        if (part.key !== `${claim.start}:${claim.end}`) continue;
        const own = glyphs.filter((glyph, index) => !/^\s*$/u.test(glyph.text ?? "x") &&
          (!part.chunk || (index * font.codeUnitBytes >= part.chunk.codeStart && index * font.codeUnitBytes < part.chunk.codeEnd)));
        if (!own.length) continue;
        const last = own.at(-1)!;
        const trailing = show.characterSpacing + (font.codeUnitBytes === 1 && last.code === 0x20 ? show.wordSpacing : 0);
        const ink = itemInk.get(itemIndex) ?? [];
        ink.push(Object.freeze({ start: u + own[0].start * scale, pen: u + (last.end - trailing) * scale,
          readerSize: show.fontSize * along }));
        itemInk.set(itemIndex, ink);
      }
    }
    indexByKey.set(`${claim.start}:${claim.end}`, shows.length);
    shows.push(Object.freeze({
      start: claim.start, end: claim.end, show, font, item: claim.item, line: lineOfItem(claim.item), leaf: claim.leaf,
      size: show.fontSize * across, characterSpacing: show.characterSpacing * across,
      horizontalScale: show.horizontalScale * (along / across), rise: show.textRise * across,
      slant, color: show.deviceColor, u, v, inked, ink, partial, twin: null, stroked: show.renderingMode === 2,
    }));
  }
  // A fake bold Qt writes draws a run again, byte for byte, half a point
  // along under a `/Span << /ActualText <> >> BDC`: the reader lists that
  // pass as an item of its own and the box builder drops it
  // (pdfJsOverprintItemIndexes), so it is no item of the paragraph and no
  // proof names it. Left standing, it would keep the old glyphs showing, in
  // bold, beside the new (the edit gates' third admission check). So a show
  // under such a span that draws an owned show's codes in its font within a
  // tenth of the size of it is that show's twin: emptied with it, and its
  // run set again by the block at the same offset. One owner has one twin.
  // (Mirror of the replay's twins.)
  const twins: { start: number; end: number; owner: { start: number; end: number } }[] = [];
  for (const candidate of allShows) {
    const key = `${candidate.start}:${candidate.end}`;
    if (indexByKey.has(key) || candidate.codes.length === 0 || !underEmptyActualText(source, candidate.start)) continue;
    const [x, y] = applyMatrix(toBlock, ...applyMatrix(candidate.graphicsMatrix, ...applyMatrix(candidate.textMatrix, 0, 0)));
    const u = x * basis[0] + y * basis[1], v = x * basis[2] + y * basis[3];
    let owner: number | null = null;
    let nearest = Number.POSITIVE_INFINITY;
    for (const [index, own] of shows.entries()) {
      if (own.partial || own.show.fontName !== candidate.fontName || !sameBytes(own.show.codes, candidate.codes)) continue;
      const tolerance = 0.1 * own.size;
      const du = u - own.u, dv = v - own.v;
      if (Math.abs(du) > tolerance || Math.abs(dv) > tolerance || Math.hypot(du, dv) >= nearest) continue;
      owner = index;
      nearest = Math.hypot(du, dv);
    }
    if (owner === null) continue;
    const own = shows[owner];
    if (own.twin) refuse(DUPLICATE_SHOWS_MESSAGE);
    shows[owner] = Object.freeze({ ...own, twin: Object.freeze([u - own.u, v - own.v] as const) });
    twins.push({ start: candidate.start, end: candidate.end, owner: { start: own.start, end: own.end } });
  }
  twins.sort((left, right) => left.start - right.start);
  const twinKeys = new Set(twins.map((twin) => `${twin.start}:${twin.end}`));
  // Each character of the paragraph's text and the show that draws it.
  const characterShows = paragraphCharacterShows(originalText, items, itemShowTexts, indexByKey);
  const blankRuns = paragraphBlankRuns(originalText, items, itemShowTexts);
  // The reader reads two glyphs of a show as touching only while the pen
  // between them stays inside its tracking bound (pdf.js
  // compareWithLastPosition: past TRACKING_SPACE_FACTOR of the size, the
  // character spacing and the kerning after a glyph read as a space). A
  // show whose character spacing alone is past that bound, and two of whose
  // letters the paragraph's text holds side by side, takes that spacing
  // back with its kerning: a producer spaces a table's figure out to its
  // column with Tc and cancels it between the glyphs it keeps together
  // (`8.54 Tc [(3.) 8540.2 (2)] TJ` reads ".2"). This tier drops the kerning
  // inside a line, and the spacing goes with it: those letters are set
  // touching, as the reader read them. (Mirror of the replay's owned show
  // spacing.)
  const paragraphCharacters = characters(originalText);
  const spacingTakenBack = new Set<number>();
  for (let index = 1; index < paragraphCharacters.length; index += 1) {
    const owner = characterShows[index];
    if (owner !== characterShows[index - 1] || isBlank(paragraphCharacters[index]) ||
      isBlank(paragraphCharacters[index - 1])) continue;
    const { show } = shows[owner];
    if (show.characterSpacing * (show.horizontalScale / 100) > PDFJS_TRACKING_SPACE_FACTOR * show.fontSize) {
      spacingTakenBack.add(owner);
    }
  }
  // Its text then ends, as this tier sets it, at its last glyph's own
  // advance: the spacing after that glyph is no part of the line.
  for (const owner of spacingTakenBack) {
    const own = shows[owner];
    const spacingAlong = own.characterSpacing * (own.horizontalScale / 100);
    shows[owner] = Object.freeze({ ...own, characterSpacing: 0,
      inked: own.inked && Object.freeze({ ...own.inked, end: own.inked.end - spacingAlong }),
      ink: Object.freeze(own.ink.map(([start, end]) => Object.freeze([start, end - spacingAlong] as const))) });
  }
  // The paragraph's lines, as the reader groups its items.
  const lineCount = Math.max(...lineIndexes) + 1;
  const lineBaselines: number[] = [];
  const lineExtents: { left: number; right: number }[] = [];
  for (let line = 0; line < lineCount; line += 1) {
    const members = shows.filter((show) => show.line === line);
    const inked = members.filter((show) => show.inked);
    if (!inked.length) refuse("A line of the paragraph draws nothing.");
    // The line's baseline is its main show's: the one that inks the most
    // glyphs, so a superscript or a subscript is the one that stands off it.
    // Of the shows set on the line's own baseline (no rise) where it has
    // any: this writer raises a part by its rise, and a last line whose
    // raised words came to outnumber its body's (a part made superscript at
    // every round) was read as standing on them, its body then written as a
    
    const level = inked.filter((show) => Math.abs(show.rise) <= LAYOUT_TOLERANCE);
    const main = (level.length ? level : inked)
      .reduce((best, show) => show.inked!.count > best.inked!.count ? show : best);
    lineBaselines.push(main.v + main.rise);
    lineExtents.push({ left: Math.min(...inked.map((show) => show.inked!.start)),
      right: Math.max(...inked.map((show) => show.inked!.end)) });
  }
  for (let line = 1; line < lineCount; line += 1) {
    if (!(lineBaselines[line - 1] - lineBaselines[line] > LAYOUT_TOLERANCE)) {
      refuse("The paragraph's lines do not run down its column.");
    }
  }
  // The body text's size, and how far each line's characters reach past it
  // (a line an earlier write set lower for a larger word: its step is read
  // back without that, and the first line's baseline too, so a paragraph set
  // again is not lowered twice).
  const bodySize = shows.filter((show) => show.inked)
    .reduce((best, show) => show.inked!.count > best.inked!.count ? show : best).size;
  const reach = lineBaselines.map((baseline, line) => lineReachOf(shows.filter((show) => show.line === line && show.inked)
    .map((show) => ({ size: show.size, rise: show.v + show.rise - baseline })), bodySize));
  const steps = lineBaselines.slice(1).map((baseline, index) => (lineBaselines[index] - baseline) /
    (1 + (reach[index + 1].above + reach[index].below) / bodySize)).sort((a, b) => a - b);
  const middle = Math.floor(steps.length / 2);
  // The median step; a one-line paragraph that has to grow takes 1.2 of its
  // body size.
  const lineStep = steps.length === 0 ? SINGLE_LINE_STEP_EM * bodySize
    : steps.length % 2 === 1 ? steps[middle] : (steps[middle - 1] + steps[middle]) / 2;
  const firstBaseline = lineBaselines[0] + firstLineLeading(lineStep, bodySize) * reach[0].above;
  const firstLeft = lineExtents[0].left;
  const restLeft = lineCount > 1 ? Math.min(...lineExtents.slice(1).map((extent) => extent.left)) : firstLeft;
  const frameLeft = Math.min(firstLeft, restLeft);
  const frameRight = Math.max(...lineExtents.map((extent) => extent.right));
  const alignment = paragraphAlignment(lineExtents, frameRight, restLeft);
  // A paragraph set flush left whose later lines start at more than one left
  // edge runs around something (a picture it wraps): this tier sets one
  // frame, and every later line at its left edge would run over it (mirror
  // of the replay).
  const lineLefts = alignment === "left" &&
    lineExtents.slice(1).some((extent) => Math.abs(extent.left - restLeft) > ALIGNMENT_TOLERANCE)
    ? Object.freeze(lineExtents.map((extent) => extent.left)) : undefined;
  // The reader ends its line inside a line -- where the next word leans
  // away in a text matrix of its own (a synthetic oblique after an upright
  // word: pdf.js measures the next glyph in that matrix's own frame), or
  // where TeX draws a subscript back under the superscript before it (u²₀:
  // the pen moves back and down by more than half the item's height, and
  // pdf.js compareWithLastPosition appends an EOL) -- and a gap the reader
  // would read as a space anywhere else on the line is read as no text at
  // all. The item before closes with an EOL, and the paragraph's text holds
  // no space between the two. This tier sets the text as the paragraph
  // reads it, dropping the kerning inside a line, and would close the gap
  // wherever the edit leaves characters on both sides of it; the gaps are
  // refused there by setParagraph (mirror of the replay).
  const unreadGaps: number[] = [];
  const unreadGapWidths: number[] = [];
  {
    const itemStarts: number[] = [];
    let offset = 0;
    for (const item of items) {
      const found = originalText.indexOf(item.str, offset);
      itemStarts.push(characters(originalText.slice(0, found)).length);
      offset = found + item.str.length;
    }
    const lastLine = Math.max(...lineIndexes);
    for (let index = 1; index < items.length; index += 1) {
      const before = items[index - 1], after = items[index];
      if (lineIndexes[index] !== lineIndexes[index - 1]) continue;
      // A justified line's gaps are its stretch, set again by this tier
      // wherever its words now fall: a Latin word in a Chinese line, its
      // own item, stood apart by the stretch this writer gave it, and every
      
      if (alignment === "justify" && lineIndexes[index] !== lastLine) continue;
      const beforeCharacters = characters(before.str), afterCharacters = characters(after.str);
      // The paragraph's text holds a space between the two: the viewer
      // joined them with one, or either item brings its own.
      if (!beforeCharacters.length || !afterCharacters.length ||
        itemStarts[index - 1] + beforeCharacters.length !== itemStarts[index] ||
        isReaderBlank(beforeCharacters.at(-1)!) || isReaderBlank(afterCharacters[0])) continue;
      // The gap as the page draws it, whatever the items say: from the pen
      // after the last glyph the item before inks (the reader's
      // prevTransform, short of the character spacing after it) to the first
      // glyph the item after inks, at the size of the show that ends the item
      // before. The reader reads a gap that wide as a space -- one it writes
      // into its item, or a white-space item of its own -- or ends its line
      // there; the text holding neither, the reader read no text in it, and
      // that is so whether or not the item before says it carries a line end
      // (a viewer that dropped the reader's white-space item, or the empty
      // item that carried its EOL, sends none). (Mirror of the replay.)
      const beforeInk = itemInk.get(index - 1), afterInk = itemInk.get(index);
      if (beforeInk?.length && afterInk?.length) {
        const ending = beforeInk.reduce((best, ink) => ink.pen > best.pen ? ink : best);
        const start = Math.min(...afterInk.map((ink) => ink.start));
        if (start - ending.pen >= PDFJS_TRACKING_SPACE_FACTOR * ending.readerSize) {
          unreadGaps.push(itemStarts[index]);
          unreadGapWidths.push(start - ending.pen);
          continue;
        }
      }
      if (!before.hasEOL) continue;
      const axis = Math.hypot(before.transform[0], before.transform[1]);
      if (!(axis > 0)) continue;
      const gap = ((after.transform[4] - before.transform[4]) * before.transform[0] +
        (after.transform[5] - before.transform[5]) * before.transform[1]) / axis - before.width;
      const size = Math.abs(before.transform[0] * before.transform[3] - before.transform[1] * before.transform[2]) / axis;
      if (gap >= PDFJS_TRACKING_SPACE_FACTOR * size) {
        unreadGaps.push(itemStarts[index]);
        unreadGapWidths.push(gap);
      }
    }
  }
  // The rules this writer drew under (through) the paragraph's runs, beside
  // their text objects: read as those runs' own formats and taken away with
  // their text, they are drawn again under the characters where they go.
  // Left, a rule stood where its run had been: a word underlined in one
  // round and set on another line by the next kept its old underline under
  // whatever came to stand there, and a paragraph moved showed it through
  
  // A rule under part of a run (WPS: two characters underlined) is those
  // characters' own: left, another part set larger slid its words out from
  // over it, and it stood under letters it never ruled.
  // Only where every such rule beside the paragraph's text objects is one
  // of its runs' (all or none: an undo finds them again by their place,
  // paragraphRuleSpans).
  const beside = paragraphRuleSpans(streamBytes, source, shows);
  const matched = new Map<number, { underline?: true; strikethrough?: true }>();
  const ruledCharacters = new Map<number, { underline?: true; strikethrough?: true }>();
  const ruledAt = new Set<number>();
  for (const [index, owned] of shows.entries()) {
    if (!beside.length || !owned.inked || !(owned.inked.end > owned.inked.start) ||
      !owned.show.graphicsMatrix.every((value, at) => Math.abs(value - blockMatrix[at]) <= 1e-9)) continue;
    const object = enclosingTextObject(source, owned.start);
    const across = owned.v + owned.rise;
    const geometry = { textMatrix: [basis[0], basis[1], basis[2], basis[3],
      owned.inked.start * basis[0] + across * basis[2], owned.inked.start * basis[1] + across * basis[3]] as
      [number, number, number, number, number, number],
    fontSize: owned.size, advance: owned.inked.end - owned.inked.start, color: owned.color };
    for (const style of ["underline", "strikethrough"] as const) {
      let standing: ReturnType<typeof pdfNativeTextRuleStandingAt> = null;
      try {
        standing = pdfNativeTextRuleStandingAt(streamBytes, style === "underline" ? object.start : object.end,
          style, geometry);
      } catch { standing = null; }
      if (!standing || ruledAt.has(standing.start)) continue;
      if (pdfNativeTextRuleRangeIsWhole(standing.range)) {
        ruledAt.add(standing.start);
        matched.set(index, { ...matched.get(index), [style]: true });
        continue;
      }
      // Its characters by where the source inks them, one to one (no
      // ligature); the blanks between them are under it too.
      const own = paragraphCharacters.flatMap((character, at) =>
        characterShows[at] === index && !isBlank(character) ? [at] : []);
      if (own.length !== owned.ink.length) continue;
      const from = owned.inked.start + standing.range.from * geometry.advance;
      const to = owned.inked.start + standing.range.to * geometry.advance;
      const under = own.filter((_, at) => {
        const middle = (owned.ink[at][0] + owned.ink[at][1]) / 2;
        return middle >= from && middle <= to;
      });
      if (!under.length) continue;
      ruledAt.add(standing.start);
      for (let at = under[0]; at <= under.at(-1)!; at += 1) {
        ruledCharacters.set(at, { ...ruledCharacters.get(at), [style]: true });
      }
    }
  }
  const ownRules = beside.length && beside.every((rule) => ruledAt.has(rule.start)) ? beside : [];
  if (!ownRules.length) ruledCharacters.clear();
  if (ownRules.length) for (const [index, marks] of matched) shows[index] = Object.freeze({ ...shows[index], ...marks });
  // And the rules another writer drew under them (a web print's links).
  const linked = paragraphLinkRules({ source, shows, basis, blockMatrix, skip: beside,
    lineSizes: lineBaselines.map((_, line) => Math.max(0, ...shows.filter((show) => show.line === line)
      .map((show) => show.size))) });
  if (linked) for (const [index, linkRule] of linked.marks) shows[index] = Object.freeze({ ...shows[index], linkRule });
  const rules = [...ownRules, ...(linked?.spans ?? [])].sort((left, right) => left.start - right.start);
  // Each line's rise is kept against its line's baseline.
  for (const [index, show] of shows.entries()) {
    shows[index] = Object.freeze({ ...show, rise: show.v + show.rise - lineBaselines[show.line] });
  }
  // The pen each emptied show leaves: its own advance, where the text after
  // it in its object draws on from where it ended.
  const pens: PdfNativeParagraphRewriteEvidence["pens"][number][] = [];
  for (const owned of shows) {
    const object = enclosingTextObject(source, owned.start);
    const objectShows = mraiNativeShowOperatorsIn(source, object.start + 2, object.end - 2);
    if (!objectShows) refuse("A text object of this paragraph cannot be read show by show.");
    // The shows emptied along with this one: every owned show after it up to
    // the first show that is not the paragraph's.
    let deletedEnd = owned.end;
    for (const other of objectShows) {
      if (other.start < owned.end) continue;
      if (!indexByKey.has(`${other.start}:${other.end}`)) break;
      deletedEnd = other.end;
    }
    if (!owned.partial && mraiNativeTextAfterShowIsPlacedAfresh(source, owned.end, deletedEnd, object.end)) continue;
    const leaf = owned.leaf;
    const scale = leaf?.sameLineLayoutState?.horizontalScale;
    if (!leaf || /^['"]/u.test(owned.show.operatorKind) || scale === undefined || !(scale > 0) ||
      !(leaf.fontSize > 0) || !(leaf.originalAdvance > 0) ||
      leaf.textOperatorSpan.start !== owned.start) {
      refuse("The text after a show of this paragraph draws on from where it ends.");
    }
    const edges = kernel.showEdgeAdjustments(LATIN1.decode(owned.show.raw));
    if (edges === null) refuse("A show of this paragraph cannot be read as one operator.");
    const { lineShift: _lineShift, sameLineCapacity: _capacity, ...own } = leaf;
    const sourceFidelity = Object.freeze({
      ...own,
      textOperatorSpan: Object.freeze({ start: owned.start, end: owned.end,
        sha256: await sha256PdfBytes(streamBytes.slice(owned.start, owned.end)) }),
    }) as PdfNativeListTextFidelityEvidence;
    const spacing = owned.show.characterSpacing !== 0 ? { characterSpacing: owned.show.characterSpacing } : {};
    if (owned.partial) {
      // The pen of the owned range: from its first glyph to its last, the
      // kerns between them included, in the show's own TJ units; the range's
      // segments become that number and the rest of the show draws on.
      const glyphs = showGlyphs(kernel, owned.show, owned.font);
      const unit = owned.font.codeUnitBytes;
      const first = glyphs[owned.partial.codeStart / unit], last = glyphs[owned.partial.codeEnd / unit - 1];
      if (!first || !last || !(owned.show.fontSize > 0)) refuse("A show of this paragraph cannot be emptied.");
      pens.push(Object.freeze({ start: owned.start, end: owned.end, sourceFidelity, chunk: owned.partial,
        number: formatNumber((-(last.end - first.start) * 1_000) / owned.show.fontSize), ...spacing }));
      continue;
    }
    pens.push(Object.freeze({ start: owned.start, end: owned.end, sourceFidelity,
      number: formatNumber((-leaf.originalAdvance * 1_000) / (leaf.fontSize * (scale / 100)) + edges), ...spacing }));
  }
  // Every item pins the objects its own proof read; one object is pinned
  // once, as a resource where any proof reads it as one.
  const pins = new Map<string, PdfNativeObjectPrecondition>();
  for (const plan of plans) {
    for (const pin of plan?.request.untouchedObjects ?? []) {
      const key = `${pin.objectNumber}:${pin.generation}`;
      if (pins.get(key)?.role !== "resource") pins.set(key, pin);
    }
  }
  const fontOrder: string[] = [];
  for (const show of shows) if (!fontOrder.includes(show.show.fontName)) fontOrder.push(show.show.fontName);
  // A font may name one character by several codes (a subset that keeps a
  // small-capital beside its plain glyph; a TeX face names each ligature
  // from two slots). The paragraph already says which of them it draws, and
  // a rewrite draws the same one -- a ligature's letters included.
  const preferred = new Map<string, Map<string, number>>();
  for (const name of fonts.keys()) preferred.set(name, new Map<string, number>());
  for (const show of shows) {
    const own = preferred.get(show.show.fontName)!;
    for (const glyph of showGlyphs(kernel, show.show, show.font)) {
      // Never a glyph the font only calls a space (inkedBlankCodes): taken
      // as the paragraph's blank, every space of the rewrite drew it (a
      // paragraph opening with a capital A named U+0020 came back with an
      
      if (glyph.text !== undefined && (characters(glyph.text).length === 1 || /^\p{L}+$/u.test(glyph.text)) &&
        !own.has(glyph.text) && !show.font.inkedBlankCodes?.has(glyph.code)) {
        own.set(glyph.text, glyph.code);
      }
    }
  }
  const styleOf = (index: number): CharacterStyle => {
    const show = shows[characterShows[index]];
    return Object.freeze({ fontName: show.show.fontName, font: show.font, preferred: preferred.get(show.show.fontName)!,
      size: show.size, characterSpacing: show.characterSpacing, horizontalScale: show.horizontalScale,
      rise: show.rise, slant: show.slant, color: show.color, twin: show.twin, syntheticBold: show.stroked });
  };
  const lineSizes = lineExtents.map((_, line) => Math.max(...shows.filter((show) => show.line === line).map((show) => show.size)));
  const gapRatio = paragraphGapRatio(kernel, selection.item.str, characterShows, shows, styleOf, lineExtents, lineSizes,
    blankRuns);
  const keptGaps = new Map<number, number>();
  const characterInk = new Map<number, readonly [number, number]>();
  {
    const all = characters(selection.item.str);
    // Each character's own inked glyph, in a show whose characters and ink go one
    // to one (no ligature): where the source set it.
    const kernOnly = new Set<number>();
    for (const [owner, own] of shows.entries()) {
      const mine = all.flatMap((character, index) => characterShows[index] === owner && !isBlank(character) ? [index] : []);
      if (mine.length !== own.ink.length) continue;
      mine.forEach((index, at) => characterInk.set(index, own.ink[at]));
      // A show that draws its blanks by kerning alone (TeX's "a, b" is one show
      // with a thin kern after the comma): the gap inside it is that kern. Set at
      // the one ratio, the line after it moved a point and a half.
      if (!showGlyphs(kernel, own.show, own.font).some((glyph) => glyph.text !== undefined && !glyph.text.trim())) kernOnly.add(owner);
    }
    for (let index = 1; index + 1 < all.length; index += 1) {
      if (!isBlank(all[index]) || isBlank(all[index - 1]) || isBlank(all[index + 1])) continue;
      const before = shows[characterShows[index - 1]], after = shows[characterShows[index + 1]];
      if (before.line !== after.line || !before.inked || !after.inked) continue;
      const inkBefore = characterInk.get(index - 1), inkAfter = characterInk.get(index + 1);
      // Inside one show the gap is the show's own blank, set at the one ratio,
      // unless the show draws it by kerning alone or by a run of blanks the
      // reader read as one: "7  矢量" set with one blank lost a space and
      
      if (before === after && !(inkBefore && inkAfter &&
        (kernOnly.has(characterShows[index]) || (blankRuns.get(index) ?? 1) > 1))) continue;
      const gap = before === after ? inkAfter![0] - inkBefore![1] : after.inked.start - before.inked.end;
      if (gap > 0 && gap <= 1.5 * Math.max(before.size, after.size)) keptGaps.set(index, gap);
    }
  }
  // A paragraph set flush left may grow to the right into the free space
  // beside it: up to the nearest other text of the stream that starts to the
  // right of its lines, on the band of any of them, or else the visible
  // page's edge. A form's frame is its own (mirror of the replay).
  let freeRight: number | null = null;
  const bounds = kernel.visiblePageBounds(selection.pageNumber);
  if (alignment === "left" && pageContent && bounds) {
    freeRight = Math.max(...[[bounds.left, bounds.bottom], [bounds.right, bounds.bottom],
      [bounds.left, bounds.top], [bounds.right, bounds.top]].map(([x, y]) => {
      const [bx, by] = applyMatrix(toBlock, x, y);
      return bx * basis[0] + by * basis[1];
    }));
    const others = otherInkedExtents(kernel, { allShows, shows, twinKeys, fonts, basis, blockMatrix });
    for (const [line, baseline] of lineBaselines.entries()) {
      const bottom = baseline - PDFJS_VERTICAL_SHIFT_RATIO * lineSizes[line];
      const top = baseline + PDFJS_VERTICAL_SHIFT_RATIO * lineSizes[line];
      for (const other of others) {
        if (other.across >= bottom && other.across <= top && other.start >= frameRight - LAYOUT_TOLERANCE) {
          freeRight = Math.min(freeRight, other.start);
        }
      }
    }
    if (!(freeRight > frameRight + LAYOUT_TOLERANCE)) freeRight = null;
  }
  const frame = paragraphFrameHoldingItsLines(kernel, selection.item.str, characterShows, shows, styleOf, gapRatio,
    { alignment, frameLeft, frameRight, firstLeft, restLeft }, freeRight);
  return Object.freeze({
    selection, items, plans, firstPlan, streamBytes, source, pageContent,
    shows: Object.freeze(shows), characterShows, basis, blockMatrix,
    layout: Object.freeze({ alignment, frameLeft: frame.left, frameRight: frame.right, firstLeft, restLeft,
      ...(lineLefts ? { lineLefts } : {}),
      firstBaseline, lineStep, lineCount, lineExtents: Object.freeze(lineExtents), bodySize,
      ...(reach[0].above > 0 ? { firstReach: reach[0].above } : {}) }),
    pens: Object.freeze(pens),
    spanStart: Math.min(shows[0].start, ...twins.map((twin) => twin.start), ...rules.map((rule) => rule.start)),
    rules: Object.freeze(rules.map((rule) => Object.freeze(rule))),
    insertion: Math.max(enclosingTextObject(source, Math.max(shows.at(-1)!.start, ...twins.map((twin) => twin.start))).end,
      ...rules.map((rule) => rule.end)),
    pins: Object.freeze([...pins.values()]),
    allShows, twins: Object.freeze(twins), twinKeys, fonts, fontOrder: Object.freeze(fontOrder), preferred, gapRatio, keptGaps,
    blankRuns, characterInk,
    characterRules: ruledCharacters,
    unreadGaps: Object.freeze(unreadGaps),
    unreadGapWidths: Object.freeze(unreadGapWidths),
  });
}

/** The message of the edit gates' third admission check for a copy of the
 * text left behind (list-text-fidelity's MRAI_NATIVE_DUPLICATE_SHOWS_MESSAGE;
 * this module is imported by that one, never the other way round). */
const DUPLICATE_SHOWS_MESSAGE =
  "The selected text is drawn twice in the same place; editing one of the shows would leave the other under it.";

/** A paragraph holds glyphs the reader read as spaces: what it reads as is
 * not what it draws, and it is not set again from that. */
export { MRAI_NATIVE_GLYPHS_READ_AS_BLANKS_MESSAGE };

/** Whether a show draws a glyph its font only calls white space. */
function drawsGlyphReadAsBlank(show: MraiNativeParagraphRewriteShow, font: MraiNativeParagraphRewriteFont | undefined): boolean {
  const inked = font?.inkedBlankCodes;
  if (!font || !inked?.size) return false;
  const unit = font.codeUnitBytes;
  for (let offset = 0; offset + unit <= show.codes.length; offset += unit) {
    if (inked.has(unit === 1 ? show.codes[offset] : (show.codes[offset] << 8) | show.codes[offset + 1])) return true;
  }
  return false;
}

/**
 * The paragraph is the text the reader read for it. A font whose ToUnicode
 * calls a drawn glyph a space (FontMetrics.inkedBlankCodes: Word's text of an
 * embedded picture, where the capital A, the colon, the quotes and the digits
 * all read U+0020) breaks that both ways. A show of the paragraph that draws
 * one reads as blanks where the page shows letters, and set again from what
 * was read it loses them. And a show made only of such glyphs -- the colon
 * and space after a line's last word, the quote a line opens with -- reads as
 * nothing at all: no item names it, so the paragraph does not own it, and it
 * would stay where it stood while the words around it moved. Either way the
 * rewrite is refused; nothing is written.
 */
function refuseGlyphsReadAsBlanks(paragraph: ReadParagraph): void {
  if (![...paragraph.fonts.values()].some((font) => font.inkedBlankCodes?.size)) return;
  if (paragraph.shows.some((own) => drawsGlyphReadAsBlank(own.show, own.font))) {
    refuse(MRAI_NATIVE_GLYPHS_READ_AS_BLANKS_MESSAGE);
  }
  const { basis, blockMatrix, layout } = paragraph;
  const toBlock = invertMatrix(blockMatrix);
  const owned = new Set([...paragraph.shows.map((show) => `${show.start}:${show.end}`), ...paragraph.twinKeys]);
  const reach = 2 * layout.bodySize;
  for (const show of paragraph.allShows) {
    if (owned.has(`${show.start}:${show.end}`) || ![0, 1, 2].includes(show.renderingMode) ||
      !drawsGlyphReadAsBlank(show, paragraph.fonts.get(show.fontName))) continue;
    const [x, y] = applyMatrix(composeMatrix(composeMatrix(show.textMatrix, show.graphicsMatrix), toBlock), 0, show.textRise);
    const along = x * basis[0] + y * basis[1], across = x * basis[2] + y * basis[3];
    // On a line of the paragraph, as far off its baseline as the reader
    // still reads a glyph on it, within two sizes of its frame.
    if (along < layout.frameLeft - reach || along > layout.frameRight + reach) continue;
    if (paragraph.shows.some((own) => Math.abs(across - own.v) <= PDFJS_VERTICAL_SHIFT_RATIO * Math.max(own.size, show.fontSize))) {
      refuse(MRAI_NATIVE_GLYPHS_READ_AS_BLANKS_MESSAGE);
    }
  }
}

function sameBytes(left: Uint8Array, right: Uint8Array): boolean {
  return left.length === right.length && left.every((byte, index) => byte === right[index]);
}

/** The marker of a pass that is no text: `/Span << /ActualText <> >> BDC`
 * (or `()`), which Qt writes over a fake bold's second pass. */
const EMPTY_ACTUAL_TEXT_MARK = new RegExp(String.raw`\/${PDF_SOURCE_NAME}\s*<<\s*\/ActualText\s*(?:<>|\(\))\s*>>\s*BDC\s*$`, "u");

/** Whether the operator at `position` stands inside an empty-/ActualText
 * marked-content span: the nearest marked-content operator before it opens
 * one (mirror of the replay's _under_empty_actual_text). */
function underEmptyActualText(source: string, position: number): boolean {
  const before = source.slice(Math.max(0, position - 4096), position);
  let last: RegExpExecArray | null = null;
  const markers = /(BDC|EMC)(?![A-Za-z0-9])/gu;
  const operator = { char: /[A-Za-z0-9]/, not: true } as const; // no lookbehind: Safari < 16.4
  for (let match = execBehind(markers, before, operator); match; match = execBehind(markers, before, operator)) last = match;
  return last !== null && last[1] === "BDC" && EMPTY_ACTUAL_TEXT_MARK.test(before.slice(0, last.index + 3));
}

/**
 * How the original lines sit in their frame. Justified: the lines before the
 * last end flush with the frame and start flush (the first may be
 * indented). Right: every line ends flush and they do not start flush.
 * Centred: the lines' middles agree and they do not start flush. Left
 * otherwise, and always for one line.
 */
function paragraphAlignment(
  extents: readonly Readonly<{ left: number; right: number }>[],
  frameRight: number,
  restLeft: number,
): PdfNativeParagraphRewriteAlignment {
  if (extents.length < 2) return "left";
  const flush = (value: number, target: number) => Math.abs(value - target) <= ALIGNMENT_TOLERANCE;
  const rest = extents.slice(1);
  const leftsFlush = rest.every((extent) => flush(extent.left, restLeft)) &&
    extents[0].left >= restLeft - ALIGNMENT_TOLERANCE;
  const rightsFlushBeforeLast = extents.slice(0, -1).every((extent) => flush(extent.right, frameRight));
  const rightsFlush = rightsFlushBeforeLast && flush(extents.at(-1)!.right, frameRight);
  if (rightsFlush && !leftsFlush) return "right";
  if (rightsFlushBeforeLast && leftsFlush) return "justify";
  const middles = extents.map((extent) => (extent.left + extent.right) / 2);
  const lefts = extents.map((extent) => extent.left);
  if (middles.every((middle) => flush(middle, middles[0])) &&
    Math.max(...lefts) - Math.min(...lefts) > ALIGNMENT_TOLERANCE) return "center";
  return "left";
}

/**
 * For each character of the paragraph's text, the owned show that draws it.
 * An item's text is found in the paragraph's text in order; an item drawn by
 * several shows is split across them by their inked characters, and a blank
 * goes with the character before it (the one after, at the very start).
 */
function paragraphCharacterShows(
  text: string,
  items: readonly PdfNativeListTextSelectionItemEvidence[],
  itemShowTexts: readonly (readonly Readonly<{ key: string; text: string }>[])[],
  indexByKey: ReadonlyMap<string, number>,
): readonly number[] {
  const all = characters(text);
  const owners: (number | null)[] = all.map(() => null);
  let offset = 0;
  for (const [index, item] of items.entries()) {
    const found = text.indexOf(item.str, offset);
    if (found < offset || !/^ *$/u.test(text.slice(offset, found))) {
      refuse("The paragraph's items cannot be found in its text.");
    }
    const first = characters(text.slice(0, found)).length;
    const own = characters(item.str);
    const shows = itemShowTexts[index];
    if (shows.length === 1) {
      const show = indexByKey.get(shows[0].key)!;
      // A blank the show draws itself (its text is the item's, blanks and
      // all) is its own: the space before 「2023」 in 「CLIMATE CHANGE
      // 2023」, with CHANGE set in another face, went with CHANGE, and NOW
      
      const drawsBlanks = shows[0].text === item.str;
      own.forEach((character, position) => { if (drawsBlanks || !isBlank(character)) owners[first + position] = show; });
    } else if (shows.length > 1) {
      const inked = own.map((character, position) => [character, position] as const)
        .filter(([character]) => !isBlank(character));
      const drawn = shows.flatMap((show) => characters(show.text).filter((character) => !isReaderBlank(character))
        .map((character) => [character, indexByKey.get(show.key)!] as const));
      if (inked.length !== drawn.length || inked.some(([character], position) => character !== drawn[position][0])) {
        refuse("A reader item's text does not match the shows that draw it.");
      }
      inked.forEach(([, position], rank) => { owners[first + position] = drawn[rank][1]; });
    }
    offset = found + item.str.length;
  }
  if (offset !== text.length && !/^ *$/u.test(text.slice(offset))) {
    refuse("The paragraph's items cannot be found in its text.");
  }
  const firstOwned = owners.find((owner) => owner !== null);
  if (firstOwned === undefined || firstOwned === null) refuse("The paragraph draws no character.");
  let previous: number = firstOwned;
  return Object.freeze(owners.map((owner) => {
    if (owner !== null) previous = owner;
    return owner ?? previous;
  }));
}

/**
 * The blanks of the paragraph's text that stand for more white space than one
 * character: an item one show draws whose text is the show's with each run
 * of white space read as one space (pdf.js: "Fixture  |  07" reads "Fixture |
 * 07"). By character of the text, the length of the run the show draws for
 * that blank; a blank the show draws as one character, or one no single show
 * spells out, is not listed.
 */
function paragraphBlankRuns(
  text: string,
  items: readonly PdfNativeListTextSelectionItemEvidence[],
  itemShowTexts: readonly (readonly Readonly<{ key: string; spelled: string }>[])[],
): ReadonlyMap<number, number> {
  const runs = new Map<number, number>();
  let offset = 0;
  for (const [index, item] of items.entries()) {
    const found = text.indexOf(item.str, offset);
    if (found < offset) break;
    offset = found + item.str.length;
    const shows = itemShowTexts[index];
    if (shows.length !== 1 || shows[0].spelled === item.str) continue;
    const drawn = characters(shows[0].spelled);
    const read = characters(item.str);
    const first = characters(text.slice(0, found)).length;
    const own = new Map<number, number>();
    // The show's blanks before the item's first character are the reader's
    // to drop; after its last, the next item's.
    let cursor = 0;
    if (read.length && !isReaderBlank(read[0])) while (cursor < drawn.length && isReaderBlank(drawn[cursor])) cursor += 1;
    let matched = true;
    for (const [position, character] of read.entries()) {
      if (cursor >= drawn.length) { matched = false; break; }
      if (isBlank(character)) {
        if (!isReaderBlank(drawn[cursor])) { matched = false; break; }
        let run = 0;
        while (cursor < drawn.length && isReaderBlank(drawn[cursor])) { run += 1; cursor += 1; }
        // A run the item ends on runs on into the next item's place.
        if (run > 1 && position + 1 < read.length) own.set(first + position, run);
        continue;
      }
      if (drawn[cursor] !== character) { matched = false; break; }
      cursor += 1;
    }
    if (matched) for (const [at, run] of own) runs.set(at, run);
  }
  return runs;
}

/**
 * Where the new text spells out, in the part an edit leaves alone, a run of
 * blanks the paragraph's text reads as one (`runOf` a blank's run length):
 * the positions in `replacement` of each such run's blanks after its first,
 * which the run is the paragraph's one blank without. A run the edit types
 * into, or one the new text spells as one blank, is not listed.
 */
function sourceSpelledRunsKept(
  original: readonly string[],
  replacement: readonly string[],
  runOf: (index: number) => number,
): readonly number[] {
  const spelled: string[] = [];
  const owner: number[] = [];
  for (const [index, character] of original.entries()) {
    const run = isBlank(character) ? Math.max(1, runOf(index)) : 1;
    for (let count = 0; count < run; count += 1) {
      spelled.push(character);
      owner.push(index);
    }
  }
  if (spelled.length === original.length) return [];
  const same = (spelt: string, typed: string) => spelt === typed || (isBlank(spelt) && isReaderBlank(typed));
  let head = 0;
  while (head < spelled.length && head < replacement.length && same(spelled[head], replacement[head])) head += 1;
  let tail = 0;
  while (tail < spelled.length - head && tail < replacement.length - head &&
    same(spelled[spelled.length - 1 - tail], replacement[replacement.length - 1 - tail])) tail += 1;
  const shift = replacement.length - spelled.length;
  const removed: number[] = [];
  for (let start = 0; start < spelled.length;) {
    let end = start + 1;
    while (end < spelled.length && owner[end] === owner[start]) end += 1;
    if (end - start > 1) {
      if (end <= head) for (let at = start + 1; at < end; at += 1) removed.push(at);
      else if (start >= spelled.length - tail) for (let at = start + 1; at < end; at += 1) removed.push(at + shift);
    }
    start = end;
  }
  return removed;
}

/** How one character of the new text is set. */
interface CharacterStyle {
  readonly fontName: string;
  readonly font: MraiNativeParagraphRewriteFont;
  /** The codes this font's own shows draw a character with. */
  readonly preferred: ReadonlyMap<string, number>;
  readonly size: number;
  readonly characterSpacing: number;
  readonly horizontalScale: number;
  readonly rise: number;
  /** The slant of the show it comes from (see OwnedShow). */
  readonly slant: number;
  readonly color: PdfNativeDeviceColor;
  /** The offset of the fake-bold pass its show is drawn again at (see
   * OwnedShow), which the block draws its run again at. */
  readonly twin: readonly [number, number] | null;
  /** Drawn filled and stroked in its colour (`2 Tr`): a character set in
   * the substitute face, which has no bold, where its own face is bold. */
  readonly syntheticBold: boolean;
  /** Ruled under (through), the rule drawn beside the block. */
  readonly underline?: boolean;
  readonly strikethrough?: boolean;
  /** Under another writer's rule, drawn again as it drew it. */
  readonly linkRule?: ParagraphLinkRule;
}

function styleKey(style: CharacterStyle): string {
  return JSON.stringify([style.fontName, style.size, style.characterSpacing, style.horizontalScale,
    style.rise, style.color.kind, style.color.components, style.slant, style.twin, style.syntheticBold,
    style.underline === true, style.strikethrough === true, style.linkRule ?? null]);
}

/** One piece line breaking moves as a whole: a run of blanks, or a word (a
 * run of CJK characters breaks between any two of them). */
interface Piece {
  readonly start: number;
  readonly end: number;
  readonly blank: boolean;
  readonly width: number;
}

interface SetLine {
  readonly start: number;
  readonly end: number;
  /** Where its last inked character ends: trailing blanks hang past it. */
  readonly inkEnd: number;
  readonly inkWidth: number;
}

/** The codes and width of one run of characters set in one style. */
function setRun(
  kernel: MraiNativeParagraphRewriteKernel,
  text: string,
  style: CharacterStyle,
): Readonly<{ codes: Uint8Array; width: number }> {
  const codes = kernel.encodeText(text, style.font, style.preferred);
  const count = codes.length / style.font.codeUnitBytes;
  const width = (kernel.glyphAdvance(codes, style.font, style.size) + style.characterSpacing * count) *
    (style.horizontalScale / 100);
  return Object.freeze({ codes, width });
}

/**
 * A blank: the font's own space glyph where it has one, and the gap the
 * paragraph's own lines leave between their words -- what the producer set,
 * justification aside -- which the glyph is moved on to by a number of its
 * own. A paragraph that leaves no gap to measure (one word, or CJK) takes
 * the font's space, or a third of the size where the font draws none, as
 * TeX's fonts do.
 */
function blankForm(
  kernel: MraiNativeParagraphRewriteKernel,
  style: CharacterStyle,
  gapRatio: number | null,
): Readonly<{ codes: Uint8Array | null; glyphWidth: number; width: number }> {
  let codes: Uint8Array | null = null;
  let glyphWidth = 0;
  try {
    const set = setRun(kernel, " ", style);
    const code = style.font.codeUnitBytes === 1 ? set.codes[0] : (set.codes[0] << 8) | set.codes[1];
    const decoded = style.font.decode.get(code);
    if (set.codes.length === style.font.codeUnitBytes && decoded === " " && !style.font.inkedBlankCodes?.has(code)) {
      codes = set.codes;
      glyphWidth = set.width;
    }
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError)) throw error;
  }
  const width = gapRatio !== null
    ? gapRatio * style.size * (style.horizontalScale / 100)
    : codes
      ? glyphWidth
      : style.size * NUMBER_SPACE_EM * (style.horizontalScale / 100);
  return Object.freeze({ codes, glyphWidth, width });
}

/**
 * What the paragraph's own lines leave between their words, as a share of
 * the line's size: each line's inked extent less the words it draws, over
 * the blanks between them. The tightest line's: the rewrite draws words
 * without the kerning the source set inside them, so a line kerned tighter
 * reads as narrower gaps, and at any wider gap that line no longer fits
 * where it stood and the paragraph wraps to a line more (a justified line's
 * stretch reads as wider gaps too, which the tightest line leaves out).
 * Null when no line has a blank, or none a gap the reader reads as a space.
 */
function paragraphGapRatio(
  kernel: MraiNativeParagraphRewriteKernel,
  text: string,
  characterShows: readonly number[],
  shows: readonly OwnedShow[],
  styleOf: (index: number) => CharacterStyle,
  lineExtents: readonly Readonly<{ left: number; right: number }>[],
  lineSizes: readonly number[],
  blankRuns: ReadonlyMap<number, number> = new Map(),
): number | null {
  const all = characters(text);
  const ratios: number[] = [];
  for (const [line, extent] of lineExtents.entries()) {
    const own = all.map((_, index) => index).filter((index) => shows[characterShows[index]].line === line);
    const first = own.findIndex((index) => !isBlank(all[index]));
    const last = own.length - 1 - [...own].reverse().findIndex((index) => !isBlank(all[index]));
    if (first < 0) continue;
    let words = 0;
    let blanks = 0;
    let cursor = first;
    while (cursor <= last) {
      if (isBlank(all[own[cursor]])) {
        // A blank the reader read for a run of them is that many blanks of
        // the gap ("Fixture  |  07": two spaces either side of the bar).
        blanks += blankRuns.get(own[cursor]) ?? 1;
        cursor += 1;
        continue;
      }
      const style = styleOf(own[cursor]);
      let stop = cursor + 1;
      while (stop <= last && !isBlank(all[own[stop]]) && styleKey(styleOf(own[stop])) === styleKey(style)) stop += 1;
      words += setRun(kernel, own.slice(cursor, stop).map((index) => all[index]).join(""), style).width;
      cursor = stop;
    }
    if (blanks < 1 || !(lineSizes[line] > 0)) continue;
    const ratio = (extent.right - extent.left - words) / (blanks * lineSizes[line]);
    // A gap no wider than the reader's tracking space is none it would read
    // as a space: the line's blanks are not word gaps -- a contents entry's
    // leader dots set apart by their character spacing, each dot's word
    // measured with the spacing after it, leave nothing between them -- and
    // blanks set that narrow would run the paragraph's words together.
    if (Number.isFinite(ratio) && ratio > PDFJS_TRACKING_SPACE_FACTOR && ratio <= 1.5) ratios.push(ratio);
  }
  return ratios.length === 0 ? null : Math.min(...ratios);
}

/** The width characters `start` to `end` take as this tier sets them: each
 * blank as its form makes it, the rest as runs of one style. */
function measureRange(
  kernel: MraiNativeParagraphRewriteKernel,
  text: readonly string[],
  styles: readonly CharacterStyle[],
  blanks: readonly (Readonly<{ width: number }> | null)[],
  start: number,
  end: number,
  kerns?: readonly number[],
): number {
  let width = 0;
  let cursor = start;
  while (cursor < end) {
    if (blanks[cursor]) {
      width += blanks[cursor]!.width;
      cursor += 1;
      continue;
    }
    if (cursor > start) width += kerns?.[cursor] ?? 0;
    const key = styleKey(styles[cursor]);
    let stop = cursor + 1;
    while (stop < end && !blanks[stop] && !kerns?.[stop] && styleKey(styles[stop]) === key) stop += 1;
    width += setRun(kernel, text.slice(cursor, stop).join(""), styles[cursor]).width;
    cursor = stop;
  }
  return width;
}

/** The pieces line breaking moves: runs of blanks, and words (split where
 * a line may break between two inked characters, breaksBetween). */
function textPieces(
  kernel: MraiNativeParagraphRewriteKernel,
  text: readonly string[],
  styles: readonly CharacterStyle[],
  blanks: readonly (Readonly<{ width: number }> | null)[],
  wide = false,
): Piece[] {
  return piecesOf(text, (start, end) => measureRange(kernel, text, styles, blanks, start, end), wide);
}

function piecesOf(text: readonly string[], measure: (start: number, end: number) => number, wide: boolean): Piece[] {
  const pieces: Piece[] = [];
  for (let index = 0; index < text.length;) {
    const blank = isBlank(text[index]);
    let end = index + 1;
    while (end < text.length && isBlank(text[end]) === blank &&
      (blank || !breaksBetween(text[end - 1], text[end], wide))) end += 1;
    pieces.push(Object.freeze({ start: index, end, blank, width: measure(index, end) }));
    index = end;
  }
  return pieces;
}

/**
 * The frame the paragraph is set in, grown to hold each of its own lines as
 * this tier sets them. The frame is read off the lines' inked extents, and
 * those are the source's, kerned: a producer that tightens a word with its
 * kerning (InDesign's `[(Commis) 10 (sioner) -25 (’) 100 (s)] TJ`) draws it
 * narrower than this tier, which drops the kerning inside a line, sets it.
 * In a frame as narrow as the kerned word the word no longer fits: a table
 * cell's heading of one word a line could not be edited at all, its widest
 * word "wider than the paragraph's frame". Each line is measured as the
 * rewrite sets it, from its first inked character to its last, and the
 * frame grows to hold it: to the right where lines start at the frame's
 * left, to the left where they end at its right, both ways where they are
 * centred. The grown frame is taken only where the paragraph's own text
 * still breaks in it where the source broke it; a frame that would pull a
 * word up onto the line above (a column of short rows, each narrower than
 * the widest) is not the source's, and the inked one stays. (Mirror of the
 * replay's frame.)
 */
/**
 * The layout of a paragraph indented by `points` on the page (Increase /
 * Decrease Indent): every line's left edge moved right by that much in the
 * paragraph's own space (its block matrix may scale it), the frame's right
 * edge kept, so the paragraph wraps in the width left (Word sets the left
 * indent and rewraps; the first line keeps its own indent beside it).
 * Refused where too little width would be left for its text.
 */
function indentedParagraphLayout(
  paragraph: Readonly<{ blockMatrix: Matrix; basis: readonly [number, number, number, number]; layout: ParagraphLayout }>,
  points: number,
): ParagraphLayout {
  const [a, b, c, d] = paragraph.blockMatrix;
  const [b0, b1] = paragraph.basis;
  // One unit along the paragraph's line direction, on the page.
  const scale = Math.hypot(b0 * a + b1 * c, b0 * b + b1 * d);
  if (!Number.isFinite(points) || !(scale > 1e-9)) refuse("The indent is not a length.");
  const shift = points / scale;
  const layout = paragraph.layout;
  const left = Math.max(layout.firstLeft, layout.restLeft, ...(layout.lineLefts ?? [])) + shift;
  if (layout.frameRight - left < 3 * Math.max(1, layout.bodySize)) {
    refuse("The paragraph has no room left for its text at that indent.");
  }
  return Object.freeze({
    ...layout,
    frameLeft: layout.frameLeft + shift,
    firstLeft: layout.firstLeft + shift,
    restLeft: layout.restLeft + shift,
    ...(layout.lineLefts ? { lineLefts: Object.freeze(layout.lineLefts.map((value) => value + shift)) } : {}),
  });
}

/**
 * The layout with its first line's left edge `points` from the others' (in
 * the paragraph's own space): Word's First line indent, or a hanging one
 * when negative. The others stay; the frame's left edge is the leftmost.
 */
function firstLineIndentedLayout(
  paragraph: Readonly<{ blockMatrix: Matrix; basis: readonly [number, number, number, number]; layout: ParagraphLayout }>,
  points: number,
): ParagraphLayout {
  const [a, b, c, d] = paragraph.blockMatrix;
  const [b0, b1] = paragraph.basis;
  const scale = Math.hypot(b0 * a + b1 * c, b0 * b + b1 * d);
  if (!Number.isFinite(points) || !(scale > 1e-9)) refuse("The first line indent is not a length.");
  const layout = paragraph.layout;
  const firstLeft = layout.restLeft + points / scale;
  if (layout.frameRight - firstLeft < 3 * Math.max(1, layout.bodySize)) {
    refuse("The first line has no room left for its text at that indent.");
  }
  return Object.freeze({
    ...layout,
    firstLeft,
    frameLeft: Math.min(firstLeft, layout.restLeft),
    ...(layout.lineLefts?.length ? { lineLefts: Object.freeze([firstLeft, ...layout.lineLefts.slice(1)]) } : {}),
  });
}










function centredParagraphFrame(
  kernel: Pick<MraiNativeParagraphRewriteKernel, "visiblePageBounds">,
  paragraph: Readonly<{ selection: Readonly<{ pageNumber: number }>; pageContent: boolean; blockMatrix: Matrix;
    basis: readonly [number, number, number, number];
    layout: Readonly<{ frameRight: number; lineExtents: readonly Readonly<{ left: number; right: number }>[] }> }>,
): Readonly<{ frameLeft: number; frameRight: number; firstLeft: number; restLeft: number }> {
  const left = Math.min(...paragraph.layout.lineExtents.map((extent) => extent.left));
  const right = Math.max(...paragraph.layout.lineExtents.map((extent) => extent.right));
  const bounds = paragraph.pageContent ? kernel.visiblePageBounds(paragraph.selection.pageNumber) : null;
  const pageLeft = bounds ? (() => {
    const toBlock = invertMatrix(paragraph.blockMatrix);
    return Math.min(...[[bounds.left, bounds.bottom], [bounds.right, bounds.bottom],
      [bounds.left, bounds.top], [bounds.right, bounds.top]].map(([x, y]) => {
      const [bx, by] = applyMatrix(toBlock, x, y);
      return bx * paragraph.basis[0] + by * paragraph.basis[1];
    }));
  })() : left;
  const room = Math.max(0, Math.min(paragraph.layout.frameRight - right, Number.isFinite(pageLeft) ? left - pageLeft : 0));
  return Object.freeze({ frameLeft: left - room, frameRight: right + room, firstLeft: left - room, restLeft: left - room });
}












function rightParagraphFrame(
  kernel: MraiNativeParagraphRewriteKernel,
  paragraph: ReadParagraph,
): Readonly<{ frameLeft: number; frameRight: number; firstLeft: number; restLeft: number }> {
  const left = Math.min(...paragraph.layout.lineExtents.map((extent) => extent.left));
  const right = Math.max(...paragraph.layout.lineExtents.map((extent) => extent.right));
  const own = Object.freeze({ frameLeft: left, frameRight: right, firstLeft: left, restLeft: left });
  const witness = kernel.streamWitness?.(paragraph.streamBytes, paragraph.firstPlan, paragraph.fonts) ?? null;
  const bounds = witness?.bounds ??
    (paragraph.pageContent ? kernel.visiblePageBounds(paragraph.selection.pageNumber) : null);
  if (!bounds) return own;
  const { basis, blockMatrix } = paragraph;
  const toBlock = invertMatrix(blockMatrix);
  const inBlock = (x: number, y: number) => {
    const [bx, by] = applyMatrix(toBlock, x, y);
    return [bx * basis[0] + by * basis[1], bx * basis[2] + by * basis[3]] as const;
  };
  const corners = (box: Readonly<{ left: number; bottom: number; right: number; top: number }>) =>
    [[box.left, box.bottom], [box.right, box.bottom], [box.left, box.top], [box.right, box.top]]
      .map(([x, y]) => inBlock(x, y));
  const boundsLeft = Math.min(...corners(bounds).map(([along]) => along));
  if (!Number.isFinite(boundsLeft) || boundsLeft >= left) return own;
  // The band each of its lines draws in, across them.
  const bands = paragraph.shows.filter((show) => show.inked).map((show) => [
    show.v + show.rise - PDFJS_VERTICAL_SHIFT_RATIO * show.size, show.v + show.rise + show.size] as const);
  const onBand = (low: number, high: number) => bands.some(([bottom, top]) => high >= bottom && low <= top);
  let limit = boundsLeft;
  for (const other of otherInkedExtents(kernel, paragraph)) {
    if (!onBand(other.across, other.across)) continue;
    const end = other.end ?? other.start;
    if (other.start >= left - LAYOUT_TOLERANCE) continue;
    limit = Math.max(limit, other.end === null ? left : end);
  }
  for (const origin of witness?.others ?? []) {
    const [along, across] = inBlock(origin.x, origin.y);
    if (onBand(across, across) && along < left - LAYOUT_TOLERANCE) limit = left;
  }
  for (const picture of witness?.pictures ?? []) {
    const placed = corners(picture);
    const alongs = placed.map(([along]) => along), acrosses = placed.map(([, across]) => across);
    if (!onBand(Math.min(...acrosses), Math.max(...acrosses)) || Math.min(...alongs) >= left - LAYOUT_TOLERANCE) continue;
    // One the line already stands over leaves its room as it was.
    if (Math.max(...alongs) > left + LAYOUT_TOLERANCE) continue;
    limit = Math.max(limit, Math.max(...alongs));
  }
  // A word's space short of what stands there.
  const gap = limit > boundsLeft ? 0.25 * paragraph.layout.bodySize : 0;
  const frameLeft = Math.min(left, limit + gap);
  return Object.freeze({ frameLeft, frameRight: right, firstLeft: frameLeft, restLeft: frameLeft });
}

function paragraphFrameHoldingItsLines(
  kernel: MraiNativeParagraphRewriteKernel,
  text: string,
  characterShows: readonly number[],
  shows: readonly OwnedShow[],
  styleOf: (index: number) => CharacterStyle,
  gapRatio: number | null,
  frame: Readonly<{ alignment: PdfNativeParagraphRewriteAlignment; frameLeft: number; frameRight: number;
    firstLeft: number; restLeft: number }>,
  freeRight: number | null = null,
): Readonly<{ left: number; right: number }> {
  const inked = Object.freeze({ left: frame.frameLeft, right: frame.frameRight });
  const all = characters(text);
  const opening = all.findIndex((character) => !isBlank(character));
  if (opening < 0) return inked;
  const own = all.slice(opening);
  const styles = own.map((_, index) => styleOf(index + opening));
  const lineOf = (index: number) => shows[characterShows[index + opening]].line;
  const lineCount = Math.max(...shows.map((show) => show.line)) + 1;
  let pieces: Piece[];
  let blanks: (Readonly<{ width: number }> | null)[];
  const lineWidths: number[] = [];
  const lineStarts: number[] = [];
  try {
    blanks = own.map((character, index) => isBlank(character) ? blankForm(kernel, styles[index], gapRatio) : null);
    pieces = textPieces(kernel, own, styles, blanks);
    for (let line = 0; line < lineCount; line += 1) {
      const members = own.map((_, index) => index).filter((index) => lineOf(index) === line && !isBlank(own[index]));
      if (!members.length) return inked;
      lineStarts.push(members[0]);
      lineWidths.push(measureRange(kernel, own, styles, blanks, members[0], members.at(-1)! + 1));
    }
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError)) throw error;
    return inked;
  }
  let left = frame.frameLeft;
  let right = frame.frameRight;
  for (const [line, width] of lineWidths.entries()) {
    if (frame.alignment === "right") {
      left = Math.min(left, right - width);
    } else if (frame.alignment === "center") {
      const excess = width - (right - left);
      if (excess > 0) {
        left -= excess / 2;
        right += excess / 2;
      }
    } else {
      right = Math.max(right, (line === 0 ? frame.firstLeft : frame.restLeft) + width);
    }
  }
  const availableIn = (frameRight: number) => (line: number) =>
    frame.alignment === "center" || frame.alignment === "right"
      ? frameRight - left
      : frameRight - (line === 0 ? frame.firstLeft : frame.restLeft);
  // The source's breaks are kept by either rule: the source's own (between
  // two CJK characters only) or WPS's (setParagraphLines), which a paragraph
  // a WPS or Word user wrote breaks by.
  const matches = (lines: readonly SetLine[]) =>
    lines.length === lineCount && lines.every((line, index) => line.start === lineStarts[index]);
  const holds = (frameRight: number) => matches(breakLines(pieces, availableIn(frameRight))) ||
    matches(setParagraphLines(kernel, own, styles, blanks, availableIn(frameRight)));
  // A paragraph set flush left grows to the right into the free space beside
  // it (readParagraph's freeRight), as far as a text box would: never so far
  // that a line takes up the first word of the line after it, which is where
  // the source broke. A line that has to run a hair wider than the source
  // set it -- a date whose new digits are wider -- then stays one line
  // instead of wrapping into a new one (mirror of the replay).
  if (freeRight !== null && frame.alignment === "left" && freeRight > right) {
    let grown = freeRight;
    for (let line = 0; line < lineCount - 1; line += 1) {
      const nextStart = lineStarts[line + 1];
      const next = pieces.find((piece) => piece.start === nextStart);
      if (!next) {
        grown = right;
        break;
      }
      const blank = pieces.find((piece) => piece.blank && piece.end === nextStart);
      const limit = (line === 0 ? frame.firstLeft : frame.restLeft) + lineWidths[line] + (blank?.width ?? 0) +
        next.width - 2 * LAYOUT_TOLERANCE;
      grown = Math.min(grown, limit);
    }
    if (grown > right && holds(grown)) right = grown;
  }
  if (left === inked.left && right === inked.right) return inked;
  if (!holds(right)) return inked;
  return Object.freeze({ left, right });
}

function colorOperator(color: PdfNativeDeviceColor): string {
  const components = color.components.map(formatNumber).join(" ");
  return `${components} ${color.kind === "gray" ? "g" : color.kind === "rgb" ? "rg" : "k"}`;
}

/** The same colour as the stroke colour (`G`, `RG`, `K`). */
function strokeColorOperator(color: PdfNativeDeviceColor): string {
  const components = color.components.map(formatNumber).join(" ");
  return `${components} ${color.kind === "gray" ? "G" : color.kind === "rgb" ? "RG" : "K"}`;
}

function hexCodes(codes: Uint8Array): string {
  return `<${Array.from(codes, (code) => code.toString(16).padStart(2, "0")).join("").toUpperCase()}>`;
}

/** A character's style set in the substitute face instead of its own. */
function inSubstitute(
  own: CharacterStyle,
  substitute: Readonly<{ resourceName: string; font: MraiNativeParagraphRewriteFont }>,
): CharacterStyle {
  // The substitute face has no bold of its own: a character whose own face
  // is bold is drawn stroked in its colour as well, the way a bold is
  
  // retyped in Chinese was saved regular).
  // Nor an italic: a character whose own face is italic is set leaning as a
  // synthesised oblique leans (a fifth of its height over its height, 12
  // degrees), the slant the block's Tm carries.
  const [ownBold, ownItalic] = faceStyle(own.font);
  const [substituteBold, substituteItalic] = faceStyle(substitute.font);
  return Object.freeze({ ...own, fontName: substitute.resourceName, font: substitute.font,
    preferred: NO_PREFERRED_CODES,
    syntheticBold: ownBold === 1 && substituteBold === 0,
    slant: ownItalic === 1 && substituteItalic === 0 && own.slant === 0 ? SYNTHETIC_OBLIQUE_SLANT : own.slant });
}

/** The paragraph's new text, set: its styles, its lines and the block. */
/** A gap between two shows of a line, along the lines, as the source sets it,
 * and how far the edit moved the text after it. */
type MovedGap = Readonly<{ from: number; to: number; across: number; size: number; dx: number }>;

function setParagraph(
  kernel: MraiNativeParagraphRewriteKernel,
  paragraph: ReadParagraph,
  replacementText: string,
  substitute: Readonly<{ resourceName: string; font: MraiNativeParagraphRewriteFont }> | null = null,
  formats: MraiNativeParagraphRewriteFormats | null = null,
  /** Every character is set in the substitute (its face is the paragraph's). */
  substituteWhole = false,
): Readonly<{
  bytes: Uint8Array;
  lines: readonly SetLine[];
  codes: readonly number[];
  advance: number;
  /** The layout the lines were set in: the paragraph's own, or with the
   * hanging indent a list marker given or taken off sets. */
  layout: ParagraphLayout;
  notices: readonly PdfNativeEditNotice[];
  fontsUsed: ReadonlySet<string>;
  /** The text is the source's and every line starts where the source's did. */
  sameBreaks: boolean;
  /** What moves the glyphs and rules standing in the gaps the edit moved. */
  carried: readonly Readonly<{ at: number; text: string }>[];
  /** Each line's baseline, where a line set larger stands lower than the
   * paragraph's step puts it. */
  lineBaselines?: readonly number[];
}> {
  const notices: PdfNativeEditNotice[] = [];
  // A line break typed in the text (Enter, as WPS takes it in a text block)
  // starts a new line of the paragraph there; the lines between breaks are
  // set as one run each. A break is no character the block draws.
  const typedLines = replacementText.split("\n");
  if (typedLines.length > 1 && typedLines.some((line) => !characters(line).some((character) => !isBlank(character)))) {
    refuse("A line break that leaves an empty line is set by the text-box lane.");
  }
  replacementText = typedLines.join("");
  // A Type3 face's glyphs are pictures: no new word can be asked of it, only
  
  // ignore the text rendering mode (ISO 32000-1 9.3.6), and no other face is
  // known to hold them.
  const type3Show = (show: Readonly<{ show: Readonly<{ fontName: string }> }>) =>
    paragraph.fonts.get(show.show.fontName)?.simpleSubtype === "Type3";
  if (paragraph.shows.some(type3Show)) {
    if (replacementText !== paragraph.selection.item.str.replace(/\n/gu, "")) {
      refuse("A paragraph drawn in a Type3 font is not rewritten.");
    }
    if (formats?.characters.some((format, index) => {
      if (!format) return false;
      const owner = paragraph.shows[paragraph.characterShows[index]];
      return Boolean(owner && type3Show(owner)) && (format.syntheticBold === true ||
        (format.face !== undefined && format.face.resourceName !== owner.show.fontName));
    })) refuse(MRAI_NATIVE_TYPE3_FORMAT_MESSAGE);
  }
  // A control character is one the reader reads where a glyph has no text of
  // its own (a bullet in a symbol face): kept where the source draws it --
  // the text before or after the edit, unchanged -- it is that glyph again;
  // typed, it is nothing a paragraph draws.
  if (/[\u0000-\u001f\u007f]/u.test(replacementText)) {
    const before = characters(paragraph.selection.item.str), after = characters(replacementText);
    let head = 0;
    while (head < before.length && head < after.length && before[head] === after[head]) head += 1;
    let tail = 0;
    while (tail < before.length - head && tail < after.length - head &&
      before[before.length - 1 - tail] === after[after.length - 1 - tail]) tail += 1;
    if (after.some((character, index) => /[\u0000-\u001f\u007f]/u.test(character) &&
      index >= head && index < after.length - tail)) {
      refuse("The new text holds a control character a paragraph cannot draw.");
    }
  }
  if (!characters(replacementText).some((character) => !isBlank(character))) {
    refuse("Emptying a paragraph is the delete command's job.");
  }
  // The substitute face reads back like any font of the paragraph: the block
  // names it, and the read-back finds it by that name.
  if (substitute) {
    paragraph = Object.freeze({ ...paragraph,
      fonts: new Map([...paragraph.fonts, [substitute.resourceName, substitute.font]]) });
  }
  // A blank run the paragraph's text opens with is the reader's, not the
  // page's: the gap the viewer folded into its first item, which no show
  // draws, and the paragraph's first line begins at its first inked glyph.
  // It is laid out as nothing, whether the edit keeps it or takes it away
  // (replay reads the texts the same way).
  const opening = characters(paragraph.selection.item.str).findIndex((character) => !isBlank(character));
  const openingKept = Math.min(opening, characters(replacementText).findIndex((character) => !isBlank(character)));
  const original = characters(paragraph.selection.item.str).slice(opening);
  const replacement = characters(replacementText).slice(openingKept);
  // The new text spelt the way the show draws a run of blanks the reader read
  // as one (the editor opens a line on the proof's own text, "Fixture  |  07"):
  // where the edit keeps such a run, it is the paragraph's one blank again,
  // which keeps the run's width (keptGaps). Laid out as typed, each of its
  // blanks took the gap of a whole word space, and a line retyped flush right
  // drew its words 8 px wider apart than the editor had shown them (stress
  
  const collapsed = sourceSpelledRunsKept(original, replacement,
    (index) => paragraph.blankRuns.get(index + opening) ?? 1);
  const formatOf = formats ? [...formats.characters] : null;
  for (const position of [...collapsed].reverse()) {
    replacement.splice(position, 1);
    formatOf?.splice(position + openingKept, 1);
  }
  // Where each typed break falls in the characters set (after the first).
  const forced: number[] = [];
  for (let count = 0, index = 0; index < typedLines.length - 1; index += 1) {
    count += characters(typedLines[index]).length;
    const at = count - openingKept - collapsed.filter((position) => position < count - openingKept).length;
    if (at > 0) forced.push(at);
  }
  const ownedBy = paragraph.characterShows.slice(opening);
  const characterShowAt = (index: number) => ownedBy[index];
  // Characters kept at either end come from their own source characters;
  // the ones typed between take the style of the character before them.
  let prefix = 0;
  while (prefix < original.length && prefix < replacement.length && original[prefix] === replacement[prefix]) prefix += 1;
  let suffix = 0;
  while (suffix < original.length - prefix && suffix < replacement.length - prefix &&
    original[original.length - 1 - suffix] === replacement[replacement.length - 1 - suffix]) suffix += 1;
  // A gap the reader read as no text is kept only by the kerning this tier
  // drops (readParagraph's unreadGaps). Where the edit leaves the place of
  // the gap between two characters of the new text that are not blanks --
  // the source's own, or typed next to it -- they would be set side by side
  // and the gap closed. An edit that takes away everything on one side of
  // it, puts a blank in it, or retypes the characters on both sides of it,
  // takes the gap with it. (Mirror of the replay's unread gaps.)
  const gapClosed = (gap: number) => {
    const boundary = gap - opening;
    const at = boundary <= prefix ? boundary
      : boundary >= original.length - suffix ? replacement.length - (original.length - boundary) : null;
    return at !== null && at > 0 && at < replacement.length && !isBlank(replacement[at - 1]) && !isBlank(replacement[at]);
  };
  // A gap the edit would close stays where the page draws it: a blank of its
  // width in both texts, which the reader then reads as a space. (Another box's
  // glyph stands in it in a display formula, the radical before a π: the box
  // read none only because that glyph was read between its two sides.)
  const heldGaps = new Map<number, number>();
  const closedGaps = paragraph.unreadGaps.flatMap((gap, index) =>
    gapClosed(gap) ? [{ boundary: gap - opening, width: paragraph.unreadGapWidths[index] }] : []);
  if (closedGaps.length) {
    if (forced.length) {
      refuse("A line of the paragraph leaves a gap between two words that its text does not read as a space.");
    }
    for (const [held, { boundary, width }] of closedGaps.entries()) {
      // Words typed at the gap go on from the word before it (x typed after the x
      // of dx is that math italic, not the upright = past the gap): the gap, and
      // what stands in it, after them.
      const typedAtGap = boundary === prefix && original.length - prefix - suffix === 0;
      const at = boundary <= prefix && !typedAtGap ? boundary : replacement.length - (original.length - boundary);
      original.splice(boundary, 0, " ");
      replacement.splice(at, 0, " ");
      formatOf?.splice(at + openingKept, 0, undefined as never);
      ownedBy.splice(boundary, 0, ownedBy[boundary - 1]);
      heldGaps.set(boundary, width);
      if (held < closedGaps.length - 1) closedGaps[held + 1] = { ...closedGaps[held + 1], boundary: closedGaps[held + 1].boundary + 1 };
      if (boundary <= prefix) prefix += 1;
    }
    prefix = 0;
    while (prefix < original.length && prefix < replacement.length && original[prefix] === replacement[prefix]) prefix += 1;
    suffix = 0;
    while (suffix < original.length - prefix && suffix < replacement.length - prefix &&
      original[original.length - 1 - suffix] === replacement[replacement.length - 1 - suffix]) suffix += 1;
  }
  /** A character of `original` as it stands in the paragraph's own text. */
  const inParagraphText = (index: number) => index + opening -
    [...heldGaps.keys()].filter((boundary) => boundary < index).length;
  // Where each character of the new text comes from: itself at either end,
  // and in between the characters it is written over, in their own order --
  // so a word in another colour keeps it. A character typed where the source
  // had none takes the one before it (the one after, at the very start).
  // A show set as a script on its line (a superscript or subscript: smaller
  // than the line's body and lifted or dropped off it) is not where words
  // typed after it go: they are set in the line's body. Taken from the
  // script, they kept its smaller size but not its lift -- a row of small
  // words on the baseline, which the reader could no longer tell from the
  // script above it, and the paragraph read back out of order (NIST
  
  const lineBody = new Map<number, OwnedShow>();
  for (const owned of paragraph.shows) {
    if (!owned.inked) continue;
    const body = lineBody.get(owned.line);
    if (!body || owned.size > body.size) lineBody.set(owned.line, owned);
  }
  const isScript = (owner: number) => {
    const owned = paragraph.shows[owner];
    const body = lineBody.get(owned.line);
    return Boolean(body && body !== owned && owned.size < body.size * 0.9 &&
      (Math.abs(owned.rise - body.rise) > 0.05 * body.size || Math.abs(owned.v - body.v) > 0.05 * body.size));
  };
  const typedAfter = (at: number): number => {
    const owner = characterShowAt(at);
    if (!isScript(owner)) return owner;
    // Inside a script the words are the script's: y typed between the x and the
    // 2 of e^{-x^2}, where the character after it is raised on this line too.
    if (at + 1 < original.length && !isBlank(original[at + 1])) {
      const next = characterShowAt(at + 1);
      if (paragraph.shows[next].line === paragraph.shows[owner].line && isScript(next)) return owner;
    }
    for (let before = at - 1; before >= 0; before -= 1) {
      const candidate = characterShowAt(before);
      if (paragraph.shows[candidate].line === paragraph.shows[owner].line && !isScript(candidate)) return candidate;
    }
    return owner;
  };
  const sourceOf = (index: number): number => {
    if (index < prefix) return characterShowAt(index);
    if (index >= replacement.length - suffix) {
      return characterShowAt(original.length - (replacement.length - index));
    }
    const replaced = original.length - suffix - prefix;
    const typed = replacement.length - suffix - prefix;
    if (replaced > 0 && typed > 0) {
      return characterShowAt(prefix + Math.floor(((index - prefix) * replaced) / typed));
    }
    // Typed against the word after a gap between two shows (a TeX line's 「. 若」,
    // the caret at the 若): that word's, not the gap's. Taken from the gap, it
    // took the bold full stop's weight.
    if (prefix > 0 && suffix > 0 && isBlank(original[prefix - 1]) && !isBlank(original[prefix]) &&
      characterShowAt(prefix) !== characterShowAt(prefix - 1)) return characterShowAt(prefix);
    if (prefix > 0) return typedAfter(prefix - 1);
    if (suffix > 0) return characterShowAt(original.length - suffix);
    return characterShowAt(0);
  };
  const canWrite = (character: string, font: MraiNativeParagraphRewriteFont, preferred: ReadonlyMap<string, number>) => {
    try {
      kernel.encodeText(character, font, preferred);
      return true;
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
      return false;
    }
  };
  // Letters the new text holds side by side are set touching, as the
  // paragraph's own are (readParagraph's spacingTakenBack): a show whose
  // character spacing alone the reader reads as a space, two of whose
  // characters the new text puts next to each other -- a page number typed
  // beside the one a tracked leader ends in -- is set without that spacing,
  // or the reader would read a space between them. (Mirror of the replay's
  // spacing taken back from the new text.)
  const spacingTakenBack = new Set<number>();
  for (let index = 1; index < replacement.length; index += 1) {
    const owner = sourceOf(index);
    if (owner !== sourceOf(index - 1) || isBlank(replacement[index]) || isBlank(replacement[index - 1])) continue;
    const { show } = paragraph.shows[owner];
    if (show.characterSpacing * (show.horizontalScale / 100) > PDFJS_TRACKING_SPACE_FACTOR * show.fontSize) {
      spacingTakenBack.add(owner);
    }
  }
  // The source character each character kept from it is (its common ends,
  // and between them the longest run of characters the two share in order):
  // the rules its run was drawn with stay on it, wherever the edit moves it.
  const keptSource = keptSourceCharacters(original, replacement, prefix, suffix);
  const styles: CharacterStyle[] = replacement.map((_, index) => {
    const owner = sourceOf(index);
    const show = paragraph.shows[owner];
    const kept = keptSource[index] >= 0 ? paragraph.shows[characterShowAt(keptSource[index])] : null;
    const keptRules = keptSource[index] >= 0 ? paragraph.characterRules.get(inParagraphText(keptSource[index])) : undefined;
    return Object.freeze({ fontName: show.show.fontName, font: show.font, preferred: paragraph.preferred.get(show.show.fontName)!,
      size: show.size, characterSpacing: spacingTakenBack.has(owner) ? 0 : show.characterSpacing,
      horizontalScale: show.horizontalScale, rise: show.rise, slant: show.slant, color: show.color, twin: show.twin,
      syntheticBold: show.stroked,
      // A rule the source drew under (through) its run, on the characters
      // kept from it; words typed take only the rules asked of them.
      ...(kept?.underline || keptRules?.underline ? { underline: true } : {}),
      ...(kept?.strikethrough || keptRules?.strikethrough ? { strikethrough: true } : {}),
      ...(kept?.linkRule ? { linkRule: kept.linkRule } : {}) });
  });
  // The formats the reader asked for, on the characters they name (the
  // edit's own text, the blanks it opens with included): each character set
  // in the face, size, colour, rise, weight and lean asked, its other
  // properties its source character's. A character the face asked cannot
  // write keeps its own face, as Word and WPS keep a Chinese character in
  // its East Asian face when a Latin face is chosen for it.
  if (formats) {
    for (let index = 0; index < replacement.length; index += 1) {
      const format = formatOf![index + openingKept];
      if (!format) continue;
      const own = styles[index];
      const face = format.face && canWrite(replacement[index], format.face.font, NO_PREFERRED_CODES)
        ? format.face : null;
      const size = format.sizePoints ?? own.size * (format.sizeScale ?? 1);
      if (!(size > 0) || !Number.isFinite(size)) refuse("A format asks a size the paragraph cannot set.");
      const scale = size / own.size;
      styles[index] = Object.freeze({ ...own,
        ...(face ? { fontName: face.resourceName, font: paragraph.fonts.get(face.resourceName) ?? face.font,
          preferred: paragraph.preferred.get(face.resourceName) ?? NO_PREFERRED_CODES } : {}),
        size,
        rise: format.riseEm !== undefined ? format.riseEm * size : own.rise * scale,
        color: format.color ?? own.color,
        characterSpacing: format.characterSpacingEm !== undefined ? format.characterSpacingEm * size
          : own.characterSpacing * scale,
        syntheticBold: format.syntheticBold ?? (face?.syntheticBold === true ? true : own.syntheticBold),
        ...(format.underline !== undefined ? { underline: format.underline } : {}),
        ...(format.strikethrough !== undefined ? { strikethrough: format.strikethrough } : {}),
        // An underline asked (or taken off) is this writer's: a link's own
        // rule goes with it.
        ...(format.underline !== undefined ? { linkRule: undefined } : {}),
        slant: format.slant ?? (face?.slanted ? SYNTHETIC_OBLIQUE_SLANT : own.slant),
        // A fake-bold pass is the source show's own second drawing: a
        // character set in another face or weight is drawn once.
        twin: face || format.syntheticBold !== undefined ? null : own.twin });
    }
  }
  // The gaps too: a gap left in the old subset switched the block back to it
  // between every two words.
  if (substitute && substituteWhole) {
    for (let position = 0; position < replacement.length; position += 1) {
      styles[position] = inSubstitute(styles[position], substitute);
    }
  }
  // A run of one style inside a word is written the way the planner writes
  // one (its ligatures included); where its font cannot write it, each
  // character its font lacks is set in the first of the paragraph's other
  // fonts that has it, the rest of the run staying in its own -- and where
  // none of them has it, in the substitute face the writer installed, at the
  // character's own size, spacing and colour.
  for (let index = 0; index < replacement.length;) {
    if (isBlank(replacement[index])) {
      index += 1;
      continue;
    }
    const key = styleKey(styles[index]);
    let end = index + 1;
    while (end < replacement.length && !isBlank(replacement[end]) && styleKey(styles[end]) === key) end += 1;
    const own = styles[index];
    if (!canWrite(replacement.slice(index, end).join(""), own.font, own.preferred)) {
      for (let position = index; position < end; position += 1) {
        const character = replacement[position];
        if (canWrite(character, own.font, own.preferred)) continue;
        // Of the paragraph's other fonts that have it, the one that looks most
        // like the character's own (same weight and slant), the first of them
        // in the paragraph's order.
        let other: string | undefined;
        let otherDistance = Infinity;
        for (const name of paragraph.fontOrder) {
          if (name === own.fontName || !canWrite(character, paragraph.fonts.get(name)!, paragraph.preferred.get(name)!)) continue;
          // A TeX font lends nothing: its x is a math italic or a bold of one
          // optical size, not this run's x. The refusal below finds the run's
          // own design (native-similar-face.ts).
          if (isTexDesignFont(paragraph.fonts.get(name)!.faceName ?? "")) continue;
          const distance = faceDistance(own.font, paragraph.fonts.get(name)!);
          if (distance < otherDistance) {
            other = name;
            otherDistance = distance;
          }
        }
        if (other === undefined && substitute && canWrite(character, substitute.font, NO_PREFERRED_CODES)) {
          styles[position] = inSubstitute(own, substitute);
          continue;
        }
        // No font of the paragraph has it: the planner's own missing-glyph
        // refusal, which the font extension answers.
        if (other === undefined) kernel.encodeText(character, own.font, own.preferred);
        styles[position] = Object.freeze({ ...own, fontName: other!, font: paragraph.fonts.get(other!)!,
          preferred: paragraph.preferred.get(other!)! });
      }
    }
    index = end;
  }
  // The lines: broken as WPS breaks an edited paragraph where that reads
  // back, else at blanks and between two CJK characters (setParagraphLines).
  const blanks = replacement.map((character, index) =>
    isBlank(character) ? blankForm(kernel, styles[index], paragraph.gapRatio) : null);
  let { layout } = paragraph;
  // 项目符号 / 编号 given to a left-set paragraph that had none: Word's
  // hanging indent. Its other lines stand in under its words -- the left
  // indent grown by the marker and the blank after it, the first line
  // hanging back as far -- not under the marker. A marker taken off a
  // paragraph that hung by one sets its lines back under its first.
  if (layout.alignment === "left" && !layout.lineLefts?.length) {
    const sourceMarker = pdfTextListMarkerOf(paragraph.selection.item.str);
    const marker = pdfTextListMarkerOf(replacement.join(""));
    const hung = layout.restLeft - layout.firstLeft;
    if (marker && !sourceMarker && Math.abs(hung) <= LAYOUT_TOLERANCE) {
      const count = characters(marker).length;
      const gap = blanks[count - 1];
      if (gap && count < replacement.length && !isBlank(replacement[count] ?? " ")) {
        let markerWidth: number | null = null;
        try {
          markerWidth = setRun(kernel, replacement.slice(0, count - 1).join(""), styles[0]).width;
        } catch (error) {
          if (!(error instanceof PdfNativeTransactionError)) throw error;
        }
        if (markerWidth !== null && markerWidth > 0) {
          layout = Object.freeze({ ...layout, restLeft: layout.firstLeft + markerWidth + gap.width });
        }
      }
    } else if (!marker && sourceMarker && hung > LAYOUT_TOLERANCE) {
      layout = Object.freeze({ ...layout, restLeft: layout.firstLeft });
    }
  }
  // A list item's marker and its text: the blank after the marker reaches
  // to the item's hanging indent, as a Word bullet's tab does. Set as a
  // word space, the text after the bullet moved in by most of the gap
  
  if (layout.alignment === "left" && layout.restLeft - layout.firstLeft >= 0.5 * layout.bodySize &&
    replacement.length > 2 && /^[•◦▪‣·\u2022\u25AA\u25CF\u2013-]$/u.test(replacement[0]) && blanks[1] && !isBlank(replacement[2] ?? "")) {
    let markerWidth: number | null = null;
    try {
      markerWidth = setRun(kernel, replacement[0], styles[0]).width;
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
    }
    const tab = markerWidth === null ? 0 : layout.restLeft - layout.firstLeft - markerWidth;
    if (tab > blanks[1].width) blanks[1] = Object.freeze({ ...blanks[1], width: tab });
  }
  // Each line's baseline. A line holding characters that reach higher than
  // its body text (set larger, or raised as a superscript) is set lower by
  // that much, the first line too, and the line after it by what reaches
  // lower than the body text; the other lines keep the paragraph's
  // step, as Word (and the editor's draft) sets a line by its own tallest
  // character. One step for the whole paragraph grew every line with one
  // large word, and a long paragraph ran down over the one below it (sweep
  
  // large word over the line above, where the draft had set it lower (sweep
  
  // The paragraph's leading as a share of its body size: what reaches past
  // the body text is spaced as the paragraph spaces its own.
  const leading = layout.bodySize > 0 ? layout.lineStep / layout.bodySize : 1;
  const baselinesOf = (set: readonly SetLine[]): number[] => {
    const reach = set.map((line) => lineReachOf(styles.slice(line.start, line.end), layout.bodySize));
    const baselines: number[] = [];
    for (let index = 0; index < set.length; index += 1) {
      baselines.push(index === 0 ? layout.firstBaseline - firstLineLeading(layout.lineStep, layout.bodySize) * reach[0].above
        : baselines[index - 1] - layout.lineStep - leading * (reach[index].above + reach[index - 1].below));
    }
    return baselines;
  };
  // A line the paragraph did not have ends at the page's right edge where
  // its frame runs past it (a source line with a URL past the margin): the
  // new lines are set on the page.
  const pageRight = (() => {
    const bounds = paragraph.pageContent ? kernel.visiblePageBounds(paragraph.selection.pageNumber) : null;
    if (!bounds) return Infinity;
    // The page's right edge along the lines (the block's own u, as
    // checkExtraLines reads other shows into it).
    const [x, y] = applyMatrix(invertMatrix(paragraph.blockMatrix), bounds.right, (bounds.bottom + bounds.top) / 2);
    const along = x * paragraph.basis[0] + y * paragraph.basis[1];
    return Number.isFinite(along) ? along : Infinity;
  })();
  const frameRightOf = (line: number) => line >= layout.lineCount ? Math.min(layout.frameRight, pageRight)
    : layout.frameRight;
  const available = (line: number) => layout.alignment === "center" || layout.alignment === "right"
    ? frameRightOf(line) - layout.frameLeft
    : frameRightOf(line) - lineLeftOf(layout, line);
  let lines = setParagraphLines(kernel, replacement, styles, blanks, available, forced);
  // The gaps kept where the line left them, its words staying where they stood
  // (one ratio for all moved a TeX line's formulas a point or two), unless that
  // takes a line more than the one ratio does: this tier sets words without the
  // source's kerning.
  {
    const gapsKept = blanks.map((blank, index) => {
      const source = keptSource[index];
      const gap = blank && source >= 0 && styles[index].size === paragraph.shows[sourceOf(index)].size
        ? heldGaps.get(source) ?? paragraph.keptGaps.get(inParagraphText(source)) : undefined;
      return gap === undefined ? blank : Object.freeze({ ...blank!, width: gap });
    });
    if (gapsKept.some((blank, index) => blank !== blanks[index])) {
      const kept = setParagraphLines(kernel, replacement, styles, gapsKept, available, forced);
      if (kept.length <= lines.length) {
        gapsKept.forEach((blank, index) => {
          if (blank === blanks[index]) return;
          blanks[index] = blank;
          // The gap is the word's after it, set in its style: left to a script
          // before it, the reader read the blank into the word from where the
          // script ended, no show stood there, and the line could not be opened
          // on its own text again.
          if (index + 1 < styles.length && !isBlank(replacement[index + 1])) {
            styles[index] = styles[index + 1];
            // (In that font's own space, or none.)
            blanks[index] = Object.freeze({ ...blankForm(kernel, styles[index], paragraph.gapRatio), width: blank!.width });
          }
        });
        lines = kept;
      }
    }
  }
  // The kerns between the characters the edit kept side by side stay too (a
  // TeX \\scriptspace before the stop after a square, a kerned pair): dropped, the
  // words after them moved a little on every edit. Unless that takes a line more.
  const kerns = replacement.map(() => 0);
  for (let index = 1; index < replacement.length; index += 1) {
    const source = keptSource[index];
    if (source < 1 || keptSource[index - 1] !== source - 1 || blanks[index] || blanks[index - 1]) continue;
    const ink = paragraph.characterInk.get(inParagraphText(source)), before = paragraph.characterInk.get(inParagraphText(source - 1));
    const show = paragraph.shows[sourceOf(index)];
    if (!ink || !before || styles[index].size !== show.size || styles[index].fontName !== show.show.fontName) continue;
    if (paragraph.shows[characterShowAt(source - 1)].line !== paragraph.shows[characterShowAt(source)].line) continue;
    const kern = ink[0] - before[1];
    if (Math.abs(kern) >= 0.01 && Math.abs(kern) < PDFJS_TRACKING_SPACE_FACTOR * show.size) kerns[index] = kern;
  }
  if (kerns.some(Boolean)) {
    const kept = setParagraphLines(kernel, replacement, styles, blanks, available, forced, false, kerns);
    if (kept.length <= lines.length) lines = kept;
    else kerns.fill(0);
  }
  // The room beside a flush-left paragraph (readParagraph's freeRight) keeps
  // a line that grew a hair from wrapping; it never pulls a word up onto the
  // line above. Set in the paragraph's own width (its widest line), an edit
  // that left fewer lines in the grown frame keeps that many: a table header
  // "Ballots / Completed" retyped "Votes" came back as one line run past its
  
  const ownRight = Math.max(...layout.lineExtents.map((extent) => extent.right));
  if (layout.alignment === "left" && !paragraph.frameRequested && layout.frameRight > ownRight + LAYOUT_TOLERANCE) {
    const own = setParagraphLines(kernel, replacement, styles, blanks,
      (line: number) => ownRight - lineLeftOf(layout, line), forced, false, kerns);
    if (lines.length < Math.min(own.length, layout.lineCount)) lines = own;
  }
  // A paragraph justified line by line by its horizontal scale (each of its
  // lines set at one Tz of its own, its main show's -- the one that inks the
  // most glyphs; a footnote mark keeps its own): the scale is the line's,
  // not its characters'. A word the edit moves onto another line takes that line's
  // scale, the line after the last the last line's. Carried with the word,
  // the line was drawn by two shows the reader reads as one item, which no
  // show alone proves, and the paragraph could not be set again (Attention
  
  // where the lines then break as they did.
  const lineScales = (() => {
    const main: (OwnedShow | null)[] = Array.from({ length: layout.lineCount }, () => null);
    for (const show of paragraph.shows) {
      if (!show.inked || show.line >= main.length) continue;
      const best = main[show.line];
      if (!best || show.inked.count > best.inked!.count) main[show.line] = show;
    }
    if (main.some((show) => show === null)) return null;
    const scales = main.map((show) => show!.horizontalScale);
    return new Set(scales).size < 2 ? null : scales;
  })();
  if (lineScales) {
    const restyled = styles.map((style, index) => {
      const line = lines.findIndex((set) => index >= set.start && index < set.end);
      const own = lineScales[paragraph.shows[sourceOf(index)].line];
      const scale = lineScales[Math.min(line < 0 ? lines.length - 1 : line, lineScales.length - 1)];
      return own === undefined || style.horizontalScale !== own || scale === own ? style
        : Object.freeze({ ...style, horizontalScale: scale });
    });
    if (restyled.some((style, index) => style !== styles[index])) {
      const restyledBlanks = replacement.map((character, index) => restyled[index] === styles[index] ? blanks[index]
        : isBlank(character) ? blankForm(kernel, restyled[index], paragraph.gapRatio) : null);
      const again = setParagraphLines(kernel, replacement, restyled, restyledBlanks, available, forced, false, kerns);
      if (again.length === lines.length && again.every((set, index) => set.start === lines[index].start &&
        set.end === lines[index].end)) {
        restyled.forEach((style, index) => { styles[index] = style; blanks[index] = restyledBlanks[index]; });
        lines = again;
      }
    }
  }
  // A word wider than its whole line runs on past the frame, the way a
  // desktop editor lets a label that gained a character grow into the room
  // beside it: on one of the paragraph's own lines of a page's content, the
  // line staying on the visible page (the other text on its line is checked
  // with every line's, by checkLinesClearOfOtherText). A line the paragraph
  // did not have, or a form's, still cannot hold it. (Mirror of the replay.)
  // A word that cannot run on so (on a line the paragraph did not have, or
  // off the page) breaks at a character instead, as Word breaks a word
  // wider than its line.
  const runsOnPastFrame = (set: readonly SetLine[], baselines = baselinesOf(set)) => set.every((line, index) => {
    if (!(line.inkWidth > available(index) + LAYOUT_TOLERANCE)) return true;
    const bounds = paragraph.pageContent ? kernel.visiblePageBounds(paragraph.selection.pageNumber) : null;
    if (index >= layout.lineCount || !bounds) return false;
    const width = line.inkWidth;
    const u = layout.alignment === "right" ? layout.frameRight - width
      : layout.alignment === "center" ? (layout.frameLeft + layout.frameRight) / 2 - width / 2
        : lineLeftOf(layout, index);
    const size = Math.max(...styles.slice(line.start, line.end).map((style) => style.size));
    const baseline = baselines[index];
    const [b0, b1, b2, b3] = paragraph.basis;
    return [[u, baseline - 0.25 * size], [u + width, baseline - 0.25 * size], [u, baseline + size],
      [u + width, baseline + size]].every(([along, across]) => {
      const [x, y] = applyMatrix(paragraph.blockMatrix, along * b0 + across * b2, along * b1 + across * b3);
      return x >= bounds.left - 0.5 && x <= bounds.right + 0.5 && y >= bounds.bottom - 0.5 && y <= bounds.top + 0.5;
    });
  });
  if (!runsOnPastFrame(lines)) {
    const split = setParagraphLines(kernel, replacement, styles, blanks, available, forced, true, kerns);
    if (split.every((line, index) => !(line.inkWidth > available(index) + LAYOUT_TOLERANCE))) lines = split;
  }
  // Set lower for its larger characters only as far as the paragraph's last
  // line then stays clear of the text below it: the page cannot push the
  // paragraph under it down, as Word would, and a lowered last line was drawn
  
  // larger word reaches up into its own paragraph's line above.
  const lineBaselines = (() => {
    const lowered = baselinesOf(lines);
    // Evenly stepped from where the source's own first line stands: its
    // characters' reach (a footnote mark raised a little past the floor)
    // lowered it in the source too. From the un-lowered baseline the
    // paragraph rose by that reach at every write (Attention p1's footnote
    
    const first = layout.firstBaseline - firstLineLeading(layout.lineStep, layout.bodySize) * (layout.firstReach ?? 0);
    const even = lines.map((_, index) => first - index * layout.lineStep);
    // A line that stays one line stays on its baseline, whatever is set larger
    
    // the preview that does not know which: a size with an underline moved the
    // line, the same size without it did not).
    if (lines.length === 1 && layout.lineCount === 1) return even;
    const last = lines.length - 1;
    if (lowered[last] >= even[last] - LAYOUT_TOLERANCE) return lowered;
    // Text under the last line's own ink, not under its frame: a page
    // number at the frame's far end kept a line its large word reached up
    
    const inkLeft = layout.alignment === "right" ? layout.frameRight - lines[last].inkWidth
      : layout.alignment === "center" ? (layout.frameLeft + layout.frameRight - lines[last].inkWidth) / 2
        : lineLeftOf(layout, last);
    const inkRight = inkLeft + lines[last].inkWidth;
    const below = otherInkedExtents(kernel, paragraph).filter((other) => other.across < even[last] - LAYOUT_TOLERANCE &&
      (other.end ?? Infinity) > inkLeft && other.start < inkRight)
      .reduce((top, other) => Math.max(top, other.across), -Infinity);
    if (!Number.isFinite(below)) return lowered;
    const reachDown = (1 - LINE_RISE_SHARE) * layout.bodySize, reachUp = LINE_RISE_SHARE * layout.bodySize;
    // As far down as the room under it goes: every line's lowering scaled
    // so the last line's stands just clear of the text below.
    const room = even[last] - reachDown - (below + reachUp);
    const drop = even[last] - lowered[last];
    if (room >= drop) return lowered;
    const at = (share: number) => lowered.map((baseline, index) => even[index] + share * (baseline - even[index]));
    // Never so little that a line's raised characters stand where the
    // reader reads them into the line above (within 0.55 of their size of
    // its baseline, native-paragraph-edit-boxes' rows): a new last line of
    // nothing but an enlarged superscript was read into the line above it,
    
    // last line then reaches into the text below it instead.
    const reachesUp = lines.map((line) => Math.max(0, ...styles.slice(line.start, line.end)
      .map((style) => style.rise + 0.6 * style.size)));
    const apart = (baselines: readonly number[]) => baselines.every((baseline, index) =>
      index === 0 || baselines[index - 1] - baseline > reachesUp[index]);
    let share = Math.max(0, room) / drop;
    if (!apart(at(share))) {
      if (!apart(lowered)) return lowered;
      let low = share, high = 1;
      for (let step = 0; step < 24; step += 1) {
        const middle = (low + high) / 2;
        if (apart(at(middle))) high = middle;
        else low = middle;
      }
      share = high;
    }
    return at(share);
  })();
  for (const [index, line] of lines.entries()) {
    if (line.inkWidth > available(index) + LAYOUT_TOLERANCE) {
      const bounds = paragraph.pageContent ? kernel.visiblePageBounds(paragraph.selection.pageNumber) : null;
      if (index >= layout.lineCount || !bounds) refuse("A word of the new text is wider than the paragraph's frame.");
      const width = line.inkWidth;
      const u = layout.alignment === "right" ? layout.frameRight - width
        : layout.alignment === "center" ? (layout.frameLeft + layout.frameRight) / 2 - width / 2
          : lineLeftOf(layout, index);
      const size = Math.max(...styles.slice(line.start, line.end).map((style) => style.size));
      const baseline = lineBaselines[index];
      const [b0, b1, b2, b3] = paragraph.basis;
      for (const [along, across] of [[u, baseline - 0.25 * size], [u + width, baseline - 0.25 * size],
        [u, baseline + size], [u + width, baseline + size]]) {
        const [x, y] = applyMatrix(paragraph.blockMatrix, along * b0 + across * b2, along * b1 + across * b3);
        if (x < bounds.left - 0.5 || x > bounds.right + 0.5 || y < bounds.bottom - 0.5 || y > bounds.top + 0.5) {
          refuse("A word of the new text is wider than the paragraph's frame, and it would run off the page.");
        }
      }
    }
  }
  if (lines.length > layout.lineCount) checkExtraLines(kernel, paragraph, lines, styles, notices, lineBaselines);
  // The block: every line placed by its own Tm, every run of one style one TJ.
  // A run set slanted is drawn under a Tm leaning as its source show did,
  // placed where the pen stands on its line; the run after it is placed back
  // upright the same way. Such a Tm stays on its line's baseline, which is
  // how a reader of the block tells it from the next line's.
  let block = "q BT\n0 Tw 0 Tr\n";
  const [b0, b1, b2, b3] = paragraph.basis;
  const placeAt = (u: number, v: number, slant: number) =>
    `${[b0, b1, b2 + slant * b0, b3 + slant * b1, u * b0 + v * b2, u * b1 + v * b3].map(formatNumber).join(" ")} Tm\n`;
  const state: { font: string | null; size: number | null; spacing: number | null;
    scale: number | null; rise: number | null; color: string | null; bold: boolean; strokeColor: string | null } =
    { font: null, size: null, spacing: null, scale: null, rise: null, color: null, bold: false, strokeColor: null };
  // The text state a run is drawn under: each part set where it changes.
  const applyState = (style: CharacterStyle) => {
    if (state.font !== style.fontName || state.size !== style.size) {
      block += `/${style.fontName} ${formatNumber(style.size)} Tf\n`;
      state.font = style.fontName;
      state.size = style.size;
    }
    if (state.spacing !== style.characterSpacing) {
      block += `${formatNumber(style.characterSpacing)} Tc\n`;
      state.spacing = style.characterSpacing;
    }
    if (state.scale !== style.horizontalScale) {
      block += `${formatNumber(style.horizontalScale)} Tz\n`;
      state.scale = style.horizontalScale;
    }
    if (state.rise !== style.rise) {
      block += `${formatNumber(style.rise)} Ts\n`;
      state.rise = style.rise;
    }
    const color = colorOperator(style.color);
    if (state.color !== color) {
      block += `${color}\n`;
      state.color = color;
    }
    // A synthetic bold: filled and stroked in the run's colour, the stroke a
    // fixed part of an em (native-synthetic-style) in the block's user space.
    // The block is drawn under its own q, so the width and colour it sets
    // end with it.
    if (style.syntheticBold) {
      const strokeColor = strokeColorOperator(style.color);
      if (!state.bold) block += "2 Tr\n";
      block += `${formatNumber(pdfNativeSyntheticBoldStrokeWidth(style.size, [b0, b1, b2, b3]))} w\n`;
      if (state.strokeColor !== strokeColor) {
        block += `${strokeColor}\n`;
        state.strokeColor = strokeColor;
      }
      state.bold = true;
    } else if (state.bold) {
      block += "0 Tr\n";
      state.bold = false;
    }
  };
  const allCodes: number[] = [];
  const fontsUsed = new Set<string>();
  let advance = 0;
  // The rules the formats ask (a run's characters underlined or struck
  // through), each under the run's inked stretch on its line, drawn as the
  // rule a press on one run draws: underlines before the block's text
  // object, strikethroughs after it.
  const underlineRules: string[] = [];
  const strikeRules: string[] = [];
  // Other writers' rules (a web print's links), each under its runs' inked
  // stretch on its line, drawn as they drew them: a rectangle filled in their
  // colour, after the strikethroughs.
  const linkRules: string[] = [];
  const ruleOf = (style: CharacterStyle, at: number, along: number, across: number, kind: "underline" | "strikethrough") =>
    LATIN1.decode(pdfNativeTextRuleBlock({
      textMatrix: [b0, b1, b2, b3, at * b0 + across * b2, at * b1 + across * b3],
      fontSize: style.size, advance: along, color: style.color,
    }, kind));
  // Where each line inks the page along the paragraph's lines: its runs of
  // inked characters, the blanks between them left out.
  const lineSpans: [number, number][][] = [];
  for (const [lineIndex, line] of lines.entries()) {
    const width = line.inkWidth;
    const u = layout.alignment === "right" ? layout.frameRight - width
      : layout.alignment === "center" ? (layout.frameLeft + layout.frameRight) / 2 - width / 2
        : lineLeftOf(layout, lineIndex);
    const v = lineBaselines[lineIndex];
    // A rise in a leaned frame moves the glyph along the leaned up axis, a
    // slant times the rise to the right of the pen: a run both leaned and
    // raised is placed back by that much, so its glyphs stand on from the
    // ones before. Placed at the pen, an italic superscript begun inside a
    // word stood 1.3 pt off it, and the next edit of the paragraph read
    
    const leanShift = (style: CharacterStyle) => style.slant * style.rise;
    let lineSlant = styles[line.start].slant;
    let lineLean = leanShift(styles[line.start]);
    block += placeAt(u - lineLean, v, lineSlant);
    // How far the pen has moved along this line.
    let pen = 0;
    const spans: [number, number][] = [];
    lineSpans.push(spans);
    advance += width;
    // The runs the line draws: where each starts along it, its style and
    // its array, for the fake-bold pass drawn again after the line.
    const runs: { penStart: number; style: CharacterStyle; array: string }[] = [];
    // Justified lines before the last stretch every blank between their
    // first and last inked characters by the same amount.
    let firstInk = line.start;
    while (firstInk < line.inkEnd && isBlank(replacement[firstInk])) firstInk += 1;
    const interior = replacement.slice(firstInk, line.inkEnd).filter(isBlank).length;
    const stretch = ((layout.alignment === "justify" && lineIndex < lines.length - 1) || layout.alignment === "distribute") &&
      interior > 0
      ? Math.max(0, available(lineIndex) - width) / interior : 0;
    // The rule a line draws under (through) consecutive runs of one rule
    // style is one rule: an underlined 「Period and」 whose space a part's
    // face had no glyph for (set in the run's own face, a run of its own)
    
    const reach: Record<"underline" | "strikethrough", { start: number; end: number; style: CharacterStyle } | null> =
      { underline: null, strikethrough: null };
    const ruleKey = (style: CharacterStyle) => JSON.stringify([style.size, style.rise, style.color]);
    let linkReach: { start: number; end: number; style: CharacterStyle } | null = null;
    const linkKey = (style: CharacterStyle) => JSON.stringify([style.size, style.rise, style.linkRule]);
    const closeLinkRule = () => {
      const open = linkReach;
      linkReach = null;
      const rule = open?.style.linkRule;
      if (!open || !rule || !(open.end > open.start)) return;
      const low = v + open.style.rise + rule.offset * open.style.size;
      const [x0, x1] = [open.start * b0, open.end * b0];
      const [y0, y1] = [low * b3, (low + rule.thickness * open.style.size) * b3];
      linkRules.push(`\n${colorOperator(rule.color)} ${[Math.min(x0, x1), Math.min(y0, y1), Math.abs(x1 - x0),
        Math.abs(y1 - y0)].map(formatNumber).join(" ")} re f\n`);
    };
    const closeRule = (kind: "underline" | "strikethrough") => {
      const open = reach[kind];
      if (open && open.end > open.start) {
        (kind === "underline" ? underlineRules : strikeRules)
          .push(ruleOf(open.style, open.start, open.end - open.start, v + open.style.rise, kind));
      }
      reach[kind] = null;
    };
    let cursor = line.start;
    while (cursor < line.end) {
      const style = styles[cursor];
      const key = styleKey(style);
      let stop = cursor + 1;
      while (stop < line.end && styleKey(styles[stop]) === key) stop += 1;
      if (style.slant !== lineSlant || leanShift(style) !== lineLean) {
        block += placeAt(u + pen - leanShift(style), v, style.slant);
        lineSlant = style.slant;
        lineLean = leanShift(style);
      }
      fontsUsed.add(style.fontName);
      applyState(style);
      const penStart = pen;
      // The array: strings of codes, a number for a blank drawn as a gap and
      // for the stretch after a blank on a justified line.
      const elements: string[] = [];
      let pendingCodes: number[] = [];
      let lastWasNumber = false;
      const flush = () => {
        if (pendingCodes.length) {
          elements.push(hexCodes(Uint8Array.from(pendingCodes)));
          lastWasNumber = false;
        }
        pendingCodes = [];
      };
      const pushNumber = (value: number) => {
        flush();
        if (lastWasNumber) elements.push("<>");
        elements.push(formatNumber(value));
        lastWasNumber = true;
      };
      const perUnit = 1_000 / (style.size * (style.horizontalScale / 100));
      let runInkStart: number | null = null;
      let runInkEnd = 0;
      for (let index = cursor; index < stop;) {
        const blank = blanks[index];
        if (blank) {
          const stretched = stretch > 0 && index >= firstInk && index < line.inkEnd;
          const extra = (stretched ? stretch : 0) + (blank.codes ? blank.width - blank.glyphWidth : blank.width);
          if (blank.codes) {
            pendingCodes.push(...blank.codes);
            allCodes.push(...blank.codes);
            if (Math.abs(extra) > 0.0005) pushNumber(-extra * perUnit);
          } else {
            pushNumber(-extra * perUnit);
          }
          pen += blank.width + (stretched ? stretch : 0);
          index += 1;
          continue;
        }
        const kern = index > line.start ? kerns[index] : 0;
        if (kern) {
          pushNumber(-kern * perUnit);
          pen += kern;
        }
        let end = index + 1;
        while (end < stop && !blanks[end] && !kerns[end]) end += 1;
        const set = setRun(kernel, replacement.slice(index, end).join(""), style);
        if (runInkStart === null) runInkStart = pen;
        runInkEnd = pen + set.width;
        pendingCodes.push(...set.codes);
        allCodes.push(...set.codes);
        if (spans.length && index > line.start && !blanks[index - 1]) spans[spans.length - 1][1] = u + pen + set.width;
        else spans.push([u + pen, u + pen + set.width]);
        pen += set.width;
        index = end;
      }
      flush();
      const array = elements.join(" ");
      block += `[${array}] TJ\n`;
      runs.push({ penStart, style, array });
      for (const kind of ["underline", "strikethrough"] as const) {
        const open = reach[kind];
        if (!style[kind] || (open && ruleKey(open.style) !== ruleKey(style))) closeRule(kind);
        if (!style[kind] || runInkStart === null || !(runInkEnd > runInkStart)) continue;
        if (reach[kind]) reach[kind]!.end = u + runInkEnd;
        else reach[kind] = { start: u + runInkStart, end: u + runInkEnd, style };
      }
      if (!style.linkRule || (linkReach && linkKey(linkReach.style) !== linkKey(style))) closeLinkRule();
      if (style.linkRule && runInkStart !== null && runInkEnd > runInkStart) {
        if (linkReach) linkReach.end = u + runInkEnd;
        else linkReach = { start: u + runInkStart, end: u + runInkEnd, style };
      }
      cursor = stop;
    }
    closeRule("underline");
    closeRule("strikethrough");
    closeLinkRule();
    // The fake-bold pass: each run whose show the page drew again is drawn
    // again, the same array at the same offset from where it stands, under
    // the empty /ActualText span the source marked its pass with, so a reader
    // that honours the mark reads it once (pdf.js reads both; the box
    // builder drops the pass, and the read-back below leaves it out).
    const twinRuns = runs.filter((run) => run.style.twin);
    if (twinRuns.length) {
      block += "/Span << /ActualText <> >> BDC\n";
      for (const run of twinRuns) {
        const [du, dv] = run.style.twin!;
        block += placeAt(u + run.penStart + du - leanShift(run.style), v + dv, run.style.slant);
        applyState(run.style);
        block += `[${run.array}] TJ\n`;
      }
      block += "EMC\n";
    }
  }
  block += "ET Q";
  checkLinesClearOfOtherText(kernel, paragraph, lines, styles, lineSpans, notices, lineBaselines);
  // The block reads back, as the reader reads it, as exactly the new text
  // (less the reader's opening blanks, which it lays out as nothing, and the
  // blanks at either end, which it never shows).
  const readBack = readBlockText(kernel, paragraph, block);
  // Wrapped, the lines read joined as the reader joins a paragraph's lines:
  // the blank a line breaks at is no character of either line, and between a
  // long CJK line and the next none is read (lineBoundaryNeedsSpace).
  // A typed break is one line boundary more, joined the same way: by the
  // lines as set (joined as typed segments, a segment holding a space was
  // never a long identifier, and 「…/final」 + 「https://…」 expected a space
  // the reader does not read).
  const expectedBack = readerSpelling(lines.length > 1
    ? joinLines(lines.map((line) => replacement.slice(line.start, line.end).join("").trim()))
    : forced.length
      ? joinLines(typedSegments(replacement, forced).map((segment) => segment.join("")))
      : replacement.join(""));
  if (readBack !== expectedBack) {
    if (process.env.NODE_ENV !== "production") {
      // Where the two first part (a dev build's probe: the refusal says only
      // that they differ).
      let at = 0;
      while (at < readBack.length && readBack[at] === expectedBack[at]) at += 1;
      try {
        console.debug("[pdf-paragraph] read-back differs " + JSON.stringify({ at,
          read: readBack.slice(Math.max(0, at - 20), at + 40), expected: expectedBack.slice(Math.max(0, at - 20), at + 40),
          readLength: readBack.length, expectedLength: expectedBack.length }));
      } catch { /* diagnostics never throw */ }
    }
    refuse(MRAI_NATIVE_PARAGRAPH_READ_BACK_REFUSAL);
  }
  // Where the source's own lines start, in the same characters (its first
  // inked character on each line).
  const sourceStarts: number[] = [];
  for (const [index, character] of original.entries()) {
    const line = paragraph.shows[characterShowAt(index)].line;
    if (!isBlank(character) && sourceStarts[line] === undefined) sourceStarts[line] = index;
  }
  // The gaps kept between two shows that the edit moved: what stands in them
  // (paragraphCarriedInGaps) goes with the text after them.
  const movedGaps: MovedGap[] = [];
  for (const [index, line] of lines.entries()) {
    const width = line.inkWidth;
    const left = layout.alignment === "right" ? layout.frameRight - width
      : layout.alignment === "center" ? (layout.frameLeft + layout.frameRight) / 2 - width / 2
        : lineLeftOf(layout, index);
    for (let at = line.start + 1; at + 1 < line.end; at += 1) {
      const source = keptSource[at];
      if (!isBlank(replacement[at]) || source < 1 || keptSource[at + 1] !== source + 1 || !isBlank(original[source])) continue;
      const before = paragraph.shows[characterShowAt(source - 1)], next = paragraph.shows[characterShowAt(source + 1)];
      if (before === next || before.line !== next.line || !before.inked || !next.inked) continue;
      const dx = left + measureRange(kernel, replacement, styles, blanks, line.start, at + 1, kerns) - next.inked.start;
      if (Math.abs(dx) > 0.01) {
        movedGaps.push(Object.freeze({ from: before.inked.end, to: next.inked.start, across: next.v, size: next.size, dx }));
      }
    }
  }
  const sameBreaks = original.join("") === replacement.join("") && lines.length === layout.lineCount &&
    lines.every((line, index) => {
      let first = line.start;
      while (first < line.end && isBlank(replacement[first])) first += 1;
      return sourceStarts[index] === first;
    });
  // The rules stand in the block's own q, beside its text object: the
  // underlines end where its BT starts, the strikethroughs start where its
  // ET ends (where a run's own rules stand, text-rule-block).
  // Each operator stands apart: with no underline before it, the block's BT
  // ran into the q ("qBT", no operator a reader knows), the text object
  // lost its start and the page's next paragraph write was refused as
  
  const after = [...strikeRules, ...linkRules];
  const ruledBlock = underlineRules.length || after.length
    ? `q${underlineRules.length ? underlineRules.join("") : " "}${block.slice(2, -2)}${after.join("")}${after.length ? "" : " "}Q`
    : block;
  // The baselines are said only where they are not the paragraph's step.
  const steppedEvenly = lineBaselines.every((baseline, index) =>
    Math.abs(baseline - (layout.firstBaseline - index * layout.lineStep)) <= 1e-9);
  return Object.freeze({ bytes: asciiBytes(ruledBlock), lines: Object.freeze(lines), codes: Object.freeze(allCodes), advance,
    layout, notices: Object.freeze(notices), fontsUsed: fontsUsed as ReadonlySet<string>, sameBreaks, carried: paragraphCarriedInGaps(kernel, paragraph, movedGaps),
    ...(steppedEvenly ? {} : { lineBaselines: Object.freeze(lineBaselines.map((baseline) => Math.round(baseline * 1e6) / 1e6)) }) });
}

/**
 * The paragraph laid out in the frame the reader resized it to: the right
 * edge moved, the left edge, the indents, the baselines and the alignment
 * kept. Only an upright paragraph (its lines along the page's x axis) has
 * its frame resized this way; the edge stays on the visible page and right
 * of where the lines start. (Mirror of the replay's requested frame.)
 */
function paragraphInRequestedFrame(
  kernel: MraiNativeParagraphRewriteKernel,
  paragraph: ReadParagraph,
  frame: Readonly<{ right: number }>,
): ReadParagraph {
  const { basis, blockMatrix, layout } = paragraph;
  const [ox, oy] = applyMatrix(blockMatrix, 0, 0);
  const [ax, ay] = applyMatrix(blockMatrix, basis[0], basis[1]);
  const alongX = ax - ox, alongY = ay - oy;
  if (!(alongX > 0) || Math.abs(alongY) > 1e-6 * alongX) {
    refuse("Only an upright paragraph is set again in a resized frame.");
  }
  const bounds = kernel.visiblePageBounds(paragraph.selection.pageNumber);
  if (!paragraph.pageContent || !bounds) refuse("A paragraph a form draws keeps the frame the form gives it.");
  if (!Number.isFinite(frame.right) || frame.right > bounds.right + 0.5 || frame.right < bounds.left - 0.5) {
    refuse("The resized frame runs off the page.");
  }
  // The edge, back along the lines: where it crosses the first baseline.
  const [, y] = applyMatrix(blockMatrix,
    layout.frameLeft * basis[0] + layout.firstBaseline * basis[2],
    layout.frameLeft * basis[1] + layout.firstBaseline * basis[3]);
  const [bx, by] = applyMatrix(invertMatrix(blockMatrix), frame.right, y);
  const right = bx * basis[0] + by * basis[1];
  if (!Number.isFinite(right) || right <= Math.max(layout.frameLeft, layout.firstLeft, layout.restLeft) + LAYOUT_TOLERANCE) {
    refuse("The resized frame leaves no room for the paragraph's text.");
  }
  return Object.freeze({ ...paragraph, frameRequested: true, layout: Object.freeze({ ...layout, frameRight: right }) });
}

/**
 * The paragraph set at the line spacing the reader chose (WPS: a paragraph's
 * line spacing re-sets its lines at that step from the first baseline down).
 * The step is given in points of the page across the lines and held to a
 * quarter of the paragraph's largest size and twenty times it. (Mirror of
 * the replay's requested line step.)
 */
function paragraphAtRequestedLineStep(paragraph: ReadParagraph, points: number): ReadParagraph {
  const { basis, blockMatrix, layout } = paragraph;
  const [ox, oy] = applyMatrix(blockMatrix, 0, 0);
  const [cx, cy] = applyMatrix(blockMatrix, basis[2], basis[3]);
  const scale = Math.hypot(cx - ox, cy - oy);
  const size = Math.max(...paragraph.shows.map((owned) => owned.size));
  if (!Number.isFinite(points) || !(scale > 0) || !(size > 0)) refuse("The line spacing cannot be read.");
  const step = points / scale;
  if (step < size * REQUESTED_LINE_STEP_MIN_EM || step > size * REQUESTED_LINE_STEP_MAX_EM) {
    refuse("The line spacing is out of range.");
  }
  if (Math.abs(step - layout.lineStep) <= LAYOUT_TOLERANCE) return paragraph;
  return Object.freeze({
    ...paragraph,
    sourceLineStep: paragraph.sourceLineStep ?? layout.lineStep,
    layout: Object.freeze({ ...layout, lineStep: step }),
  });
}

/**
 * The paragraph set at the size the reader chose for all of it (WPS: a font
 * size set on a text block re-sets its lines in its frame): every show's
 * size, spacing and rise scaled by one factor, and its line step with them
 * unless the reader set one too. Its frame is the one the request pins.
 * (Mirror of the replay's requested size.)
 */
function paragraphAtRequestedSize(paragraph: ReadParagraph, scale: number, lineStepRequested: boolean): ReadParagraph {
  if (!Number.isFinite(scale) || scale < REQUESTED_SIZE_SCALE_MIN || scale > REQUESTED_SIZE_SCALE_MAX) {
    refuse("The font size is out of range.");
  }
  if (Math.abs(scale - 1) <= 1e-6) return paragraph;
  const shows = Object.freeze(paragraph.shows.map((owned) => Object.freeze({
    ...owned, size: owned.size * scale, characterSpacing: owned.characterSpacing * scale, rise: owned.rise * scale,
  })));
  const layout = lineStepRequested ? paragraph.layout
    : Object.freeze({ ...paragraph.layout, lineStep: paragraph.layout.lineStep * scale });
  return Object.freeze({
    ...paragraph,
    shows,
    layout,
    sizeScale: scale,
    ...(lineStepRequested ? {} : { sourceLineStep: paragraph.sourceLineStep ?? paragraph.layout.lineStep }),
  });
}

/**
 * Greedy line breaking: a line takes pieces while they fit its width; a
 * word that does not fit starts the next line, and the blanks before it
 * stay at the end of the line they follow, hanging past its last inked
 * character. With `split`, a word wider than a whole line is broken where
 * the line is full, as WPS breaks a long run of letters, a combining mark
 * kept with its base.
 */
function breakLines(
  pieces: readonly Piece[],
  available: (line: number) => number,
  split?: Readonly<{ measure: (start: number, end: number) => number; splitsAt: (index: number) => boolean }>,
): SetLine[] {
  const lines: SetLine[] = [];
  let start = pieces[0]?.start ?? 0;
  let inkEnd = start;
  let inkWidth = 0;
  let pending = 0;
  let inked = false;
  let end = start;
  for (const piece of pieces) {
    if (piece.blank) {
      pending += piece.width;
      end = piece.end;
      continue;
    }
    if (inked && inkWidth + pending + piece.width > available(lines.length) + LAYOUT_TOLERANCE) {
      lines.push(Object.freeze({ start, end, inkEnd, inkWidth }));
      start = piece.start;
      inkWidth = 0;
      pending = 0;
      inked = false;
    }
    let from = piece.start;
    let width = piece.width;
    while (split && !inked && pending + width > available(lines.length) + LAYOUT_TOLERANCE) {
      const room = available(lines.length) + LAYOUT_TOLERANCE - pending;
      let cut = from + 1;
      while (cut + 1 < piece.end && split.measure(from, cut + 1) <= room) cut += 1;
      while (cut < piece.end && !split.splitsAt(cut)) cut += 1;
      if (cut >= piece.end) break;
      lines.push(Object.freeze({ start, end: cut, inkEnd: cut, inkWidth: pending + split.measure(from, cut) }));
      start = cut;
      from = cut;
      pending = 0;
      width = split.measure(from, piece.end);
    }
    inkWidth += pending + width;
    pending = 0;
    inked = true;
    inkEnd = piece.end;
    end = piece.end;
  }
  lines.push(Object.freeze({ start, end, inkEnd, inkWidth }));
  return lines;
}

/**
 * The lines a paragraph's text is set in: broken as WPS breaks it (where a
 * CJK character meets another, a run wider than the line at a character)
 * when those lines read back, joined as the reader joins them, as exactly
 * the text; otherwise only at blanks and between two CJK characters, which
 * always do. A break the reader would read a space into (between two words
 * of letters, in a paragraph with spaces) is never made. (Mirror of the
 * replay's _set_lines.)
 */
function setParagraphLines(
  kernel: MraiNativeParagraphRewriteKernel,
  text: readonly string[],
  styles: readonly CharacterStyle[],
  blanks: readonly (Readonly<{ width: number }> | null)[],
  available: (line: number) => number,
  forced: readonly number[] = [],
  splitWords = false,
  kerns?: readonly number[],
): SetLine[] {
  return paragraphLines(text, (start, end) => measureRange(kernel, text, styles, blanks, start, end, kerns), available, forced,
    splitWords);
}

/** The text between typed breaks (`forced`, positions in `text`). */
function typedSegments<T>(text: readonly T[], forced: readonly number[]): T[][] {
  const segments: T[][] = [];
  let start = 0;
  for (const stop of [...forced, text.length]) {
    segments.push(text.slice(start, stop));
    start = stop;
  }
  return segments;
}

/** A paragraph's lines: each run between typed breaks (`forced`) starts a
 * line and is broken on its own (mirror of the replay's _set_lines). */
function paragraphLines(
  text: readonly string[],
  measure: (start: number, end: number) => number,
  available: (line: number) => number,
  forced: readonly number[] = [],
  splitWords = false,
): SetLine[] {
  const lines: SetLine[] = [];
  let start = 0;
  for (const stop of [...forced, text.length]) {
    const offset = start, before = lines.length;
    const segment = text.slice(offset, stop);
    const local = (from: number, to: number) => measure(from + offset, to + offset);
    const room = (line: number) => available(before + line);
    const wide = breakLines(piecesOf(segment, local, true), room, {
      measure: local, splitsAt: (index) => !COMBINING_MARK.test(segment[index]),
    });
    // Broken only between words where the character breaks would read
    // otherwise -- unless asked to break a word wider than its whole line at
    // a character (`splitWords`, a long URL set larger that could not run on
    // past its frame), as Word breaks it, the reader reading the break as a
    // space (the read-back joins the lines so).
    const set = splitWords ||
      joinLines(wide.map((line) => segment.slice(line.start, line.end).join(""))) === segment.join("")
      ? wide : breakLines(piecesOf(segment, local, false), room);
    for (const line of set) {
      lines.push(Object.freeze({ start: line.start + offset, end: line.end + offset,
        inkEnd: line.inkEnd + offset, inkWidth: line.inkWidth }));
    }
    start = stop;
  }
  return lines;
}

/** Tests (and the replay's parity test): the lines `text` is set in, each
 * character `widths[i]` wide, `available(line)` the room of each line. */
export function paragraphLinesForTests(
  text: string,
  widths: readonly number[],
  available: (line: number) => number,
): readonly Readonly<{ start: number; end: number; inkEnd: number; inkWidth: number }>[] {
  const measure = (start: number, end: number) => widths.slice(start, end).reduce((sum, width) => sum + width, 0);
  return paragraphLines(characters(text), measure, available);
}

/** Tests (and the replay's parity test): as paragraphLinesForTests, with
 * the positions of typed breaks. */
export function paragraphLinesWithBreaksForTests(
  text: string,
  widths: readonly number[],
  available: (line: number) => number,
  forced: readonly number[],
): readonly Readonly<{ start: number; end: number; inkEnd: number; inkWidth: number }>[] {
  const measure = (start: number, end: number) => widths.slice(start, end).reduce((sum, width) => sum + width, 0);
  return paragraphLines(characters(text), measure, available, forced);
}

/**
 * A line the original paragraph did not have is drawn only where nothing
 * else of this stream draws text and inside the visible page: no other
 * show may start in its band across the frame. Form content is never given
 * a new line.
 */
function checkExtraLines(
  kernel: MraiNativeParagraphRewriteKernel,
  paragraph: ReadParagraph,
  lines: readonly SetLine[],
  styles: readonly CharacterStyle[],
  notices: PdfNativeEditNotice[],
  baselines: readonly number[],
): void {
  const { layout, basis, blockMatrix } = paragraph;
  if (!paragraph.pageContent) refuse("The new text needs more lines than the paragraph has, and a form gives it none.");
  const bounds = kernel.visiblePageBounds(paragraph.selection.pageNumber);
  if (!bounds) refuse("The new text needs more lines than the paragraph has, and the page's box is unknown.");
  const toBlock = invertMatrix(blockMatrix);
  const owned = new Set([...paragraph.shows.map((show) => `${show.start}:${show.end}`), ...paragraph.twinKeys]);
  const others = paragraph.allShows.filter((show) => show.codes.length > 0 && !owned.has(`${show.start}:${show.end}`))
    .map((show) => {
      const [x, y] = applyMatrix(toBlock, ...applyMatrix(show.graphicsMatrix, ...applyMatrix(show.textMatrix, 0, 0)));
      return [x * basis[0] + y * basis[1], x * basis[2] + y * basis[3]] as const;
    });
  for (let index = layout.lineCount; index < lines.length; index += 1) {
    const line = lines[index];
    const size = Math.max(...styles.slice(line.start, line.end).map((style) => style.size));
    const baseline = baselines[index];
    const bottom = baseline - 0.25 * size, top = baseline + size;
    // Other text standing where the new line goes is drawn over, as a text
    // box grown in Word or Acrobat draws over what lies below it: the write
    // goes through and the user is told (docs/pdf/edit-gates.md).
    if (others.some(([u, v]) => v >= bottom && v <= top && u >= layout.frameLeft - 0.5 && u <= layout.frameRight + 0.5) &&
      !notices.some((notice) => notice.code === "new-line-over-other-text")) {
      notices.push(Object.freeze({ code: "new-line-over-other-text" }));
    }
    // The new line as it is set (its own ink across the line), not the
    // frame: a frame the source's own lines run past the page with (a URL
    // past the right margin) would refuse every new line.
    const width = line.inkWidth;
    const left = layout.alignment === "right" ? layout.frameRight - width
      : layout.alignment === "center" ? (layout.frameLeft + layout.frameRight) / 2 - width / 2
        : lineLeftOf(layout, index);
    for (const [u, v] of [[left, bottom], [left + width, bottom], [left, top], [left + width, top]]) {
      const [x, y] = applyMatrix(blockMatrix, u * basis[0] + v * basis[2], u * basis[1] + v * basis[3]);
      if (x < bounds.left - 0.5 || x > bounds.right + 0.5 || y < bounds.bottom - 0.5 || y > bounds.top + 0.5) {
        refuse("The new text needs another line, and it would run off the page.");
      }
    }
  }
}

/** Below this share of the size a gap after a white-space glyph reads as no
 * space; past -0.2 of it the reader starts a new item (pdfjs-dist
 * legacy/build/pdf.worker.mjs NOT_A_SPACE_FACTOR, NEGATIVE_SPACE_FACTOR). */
const PDFJS_NOT_A_SPACE_FACTOR = 0.03;
const PDFJS_NEGATIVE_SPACE_FACTOR = -0.2;
const PDFJS_SPACE_IN_FLOW_MAX_FACTOR = 0.6;
/** A glyph raised or lowered by more than this share of its item's height
 * starts a new item (pdf.js VERTICAL_SHIFT_RATIO). */
const PDFJS_VERTICAL_SHIFT_RATIO = 0.25;
/** The most bytes one content-stream transaction writes (content-stream-transaction's MAX_REPLACEMENT_BYTES). */
const PARAGRAPH_REWRITE_MAX_REPLACEMENT_BYTES = 64 * 1024;
/** The reader's zero-width diacritics and invisible format marks (pdf.js
 * SpecialCharRegExp groups 2 and 3). */
const PDFJS_ZERO_WIDTH_DIACRITIC = /^\p{Mn}$/u;
const PDFJS_INVISIBLE_FORMAT_MARK = /^\p{Cf}$/u;

/**
 * The stream's other inked shows, along and across the paragraph's lines:
 * every show the paragraph does not own that draws a character other than
 * white space, as where its glyphs start and end along the lines and its
 * baseline (its rise included) across them. A show in a font the paragraph
 * does not name, or whose codes its font does not read, is taken as inked;
 * one whose widths the paragraph's fonts do not give has no known end
 * (mirror of the replay's _other_inked_extents).
 */
function otherInkedExtents(
  kernel: MraiNativeParagraphRewriteKernel,
  paragraph: Pick<ReadParagraph, "allShows" | "shows" | "twinKeys" | "fonts" | "basis" | "blockMatrix">,
): Readonly<{ start: number; end: number | null; across: number; show: MraiNativeParagraphRewriteShow }>[] {
  const { basis, blockMatrix } = paragraph;
  const toBlock = invertMatrix(blockMatrix);
  const owned = new Set(paragraph.shows.map((show) => `${show.start}:${show.end}`));
  return paragraph.allShows
    .filter((show) => show.codes.length > 0 && show.fontSize > 0 && [0, 1, 2].includes(show.renderingMode) &&
      !owned.has(`${show.start}:${show.end}`) && !paragraph.twinKeys.has(`${show.start}:${show.end}`) &&
      (show.text === null || !characters(show.text).every(isReaderBlank)))
    .map((show) => {
      const toLines = composeMatrix(composeMatrix(show.textMatrix, show.graphicsMatrix), toBlock);
      // A point of the show's baseline, its rise included, along and across
      // the paragraph's lines.
      const placed = (along: number) => {
        const [x, y] = applyMatrix(toLines, along, show.textRise);
        return [x * basis[0] + y * basis[1], x * basis[2] + y * basis[3]] as const;
      };
      const font = paragraph.fonts.get(show.fontName);
      const run = font && !show.vertical ? showRun(kernel, show, font) : null;
      if (!run) {
        const [start, across] = placed(0);
        return Object.freeze({ start, end: null, across, show });
      }
      const scale = show.horizontalScale / 100;
      const [start, across] = placed(run.start * scale), [end] = placed(run.end * scale);
      return Object.freeze({ start: Math.min(start, end), end: Math.max(start, end), across, show });
    });
}

/**
 * Where a show's glyphs start and end along its text space, before its
 * horizontal scale: from past the numbers before its first string for its
 * own advance (showGlyphs' measure). Null where its font cannot measure it.
 */
function showRun(
  kernel: MraiNativeParagraphRewriteKernel,
  show: MraiNativeParagraphRewriteShow,
  font: MraiNativeParagraphRewriteFont,
): Readonly<{ start: number; end: number }> | null {
  const counts = show.segmentCodeCounts.length ? show.segmentCodeCounts : [show.codes.length];
  if (counts.some((count) => count % font.codeUnitBytes !== 0) ||
    counts.reduce((total, count) => total + count, 0) !== show.codes.length) return null;
  try {
    const glyphs = showGlyphs(kernel, show, font);
    return glyphs.length ? Object.freeze({ start: glyphs[0].start, end: glyphs.at(-1)!.end }) : null;
  } catch (error) {
    if (error instanceof PdfNativeTransactionError) return null;
    throw error;
  }
}

/**
 * A line of the paragraph is set over no text the paragraph does not own:
 * no other inked show on its baseline band runs into a run of inked
 * characters the line draws (the blanks between words, however far a
 * justified line stretches them, draw nothing) further than it already ran
 * into the paragraph's own ink on that line in the source. How far a show
 * runs into a stretch of ink is measured from the nearer of the stretch's two
 * ends; the source's is the most it reaches into any one inked glyph the
 * paragraph's shows draw on the line, so only overlap the rewrite brings is
 * refused. Word places each span of a tagged line by a Tm of its own rounded
 * to two decimals, so the span before a line may end 0.04 pt past where the
 * line's first glyph starts; the line set again from that same start keeps
 * that overlap, and draws over nothing more. That allowance is for rounding
 * only (SOURCE_ABUTMENT_MAX): a show the source already sets over a whole
 * glyph of the paragraph is text drawn twice, which the rewrite still refuses
 * to set over again. A show whose end is unknown is
 * taken to run on from where it starts, so one starting from the line's
 * first inked glyph up to its last counts, whatever the source drew there. A
 * show that only touches a run -- starts where it ends, as TeX sets a big
 * operator after a bracket -- does not. The reader may leave a word of the
 * paragraph's line out of the paragraph (pdf.js ends its line where a word
 * is set slanted by a text matrix of its own, and the viewer then gives that
 * word a box of its own), and a line set over it would draw over it (mirror
 * of the replay's line check). Lines the paragraph did not have are checked
 * across the whole frame by checkExtraLines.
 */
function checkLinesClearOfOtherText(
  kernel: MraiNativeParagraphRewriteKernel,
  paragraph: ReadParagraph,
  lines: readonly SetLine[],
  styles: readonly CharacterStyle[],
  lineSpans: readonly (readonly (readonly [number, number])[])[],
  notices: PdfNativeEditNotice[],
  baselines: readonly number[],
): void {
  const { layout } = paragraph;
  const others = otherInkedExtents(kernel, paragraph);
  if (!others.length) return;
  for (let index = 0; index < Math.min(lines.length, layout.lineCount); index += 1) {
    const line = lines[index];
    const spans = lineSpans[index];
    if (!spans.length) continue;
    const size = Math.max(...styles.slice(line.start, line.end).map((style) => style.size));
    const baseline = baselines[index];
    // On the line: as far off its baseline as the reader still reads a glyph
    // on it (pdf.js VERTICAL_SHIFT_RATIO of the size).
    const bottom = baseline - PDFJS_VERTICAL_SHIFT_RATIO * size, top = baseline + PDFJS_VERTICAL_SHIFT_RATIO * size;
    // The paragraph's own ink on the line as the source sets it: every
    // inked glyph of its shows the reader groups on that line.
    const sourceInk = paragraph.shows.filter((show) => show.line === index).flatMap((show) => show.ink);
    // How far a show running from `start` to `end` reaches into a stretch of
    // ink, from the stretch's nearer end (at or below zero it does not).
    const depth = (start: number, end: number, [inkStart, inkEnd]: readonly [number, number]) =>
      Math.min(inkEnd - start, end - inkStart);
    const runsInto = ({ start, end, across }: (typeof others)[number]) => {
      if (!(across >= bottom && across <= top)) return false;
      if (end === null) return start >= spans[0][0] - LAYOUT_TOLERANCE && start < spans.at(-1)![1] - LAYOUT_TOLERANCE;
      const already = Math.min(SOURCE_ABUTMENT_MAX, Math.max(0, ...sourceInk.map((glyph) => depth(start, end, glyph))));
      return spans.some((span) => depth(start, end, span) > already + LAYOUT_TOLERANCE);
    };
    // A show drawn exactly over the paragraph's own ink on the line is a
    // second copy of the text (a fake bold, a printer driver's shadow): set
    // one copy again and the other keeps showing the old text under it, so
    
    // the message is list-text-fidelity's MRAI_NATIVE_DUPLICATE_SHOWS_MESSAGE).
    if (sourceInk.length) {
      const inkStart = Math.min(...sourceInk.map(([start]) => start));
      const inkEnd = Math.max(...sourceInk.map(([, end]) => end));
      if (others.some(({ start, end, across }) => across >= bottom && across <= top && end !== null &&
        Math.abs(start - inkStart) <= LAYOUT_TOLERANCE && Math.abs(end - inkEnd) <= LAYOUT_TOLERANCE)) {
        refuse("The selected text is drawn twice in the same place; editing one of the shows would leave the other under it.");
      }
    }
    // The line runs over other text of the page: drawn over, not lost; the
    // write goes through and the user is told (docs/pdf/edit-gates.md).
    if (others.some(runsInto) && !notices.some((notice) => notice.code === "line-over-other-text")) {
      notices.push(Object.freeze({ code: "line-over-other-text" }));
    }
  }
}

/**
 * The text a block reads as, read the way pdf.js reads it
 * (pdfjs-dist legacy/build/pdf.worker.mjs buildTextContentItem and
 * compareWithLastPosition), line by line, the lines joined the way the
 * reader joins a paragraph's lines. A white-space glyph is never written: it
 * only moves the pen. Between two inked glyphs the reader measures the gap
 * from the end of the first one's width -- the character spacing and the
 * number after it are part of the gap -- and reads a space past 0.102 of the
 * size, or past 0.03 of it when a white-space glyph stands in the gap. A
 * number with no glyph before it in its array moves the pen by the character
 * spacing too, as the reader's empty chunk does. A blank written as a space
 * glyph and pulled back to its neighbour, or letters set under a tracking
 * the reader reads as spaces, therefore read otherwise than the text they
 * were set for (replay reads the block the same way). A Tm on the baseline
 * of the line it is on places a slanted run, or the upright run after one,
 * where the pen stands: the reader reads on from an oblique run into an
 * upright one, and ends its line where an upright run is followed by one set
 * in a skewed matrix of its own (compareWithLastPosition measures the next
 * glyph through that matrix's inverse).
 */
function readBlockText(
  kernel: MraiNativeParagraphRewriteKernel,
  paragraph: ReadParagraph,
  block: string,
): string {
  const lines: string[] = [];
  let font: MraiNativeParagraphRewriteFont | null = null;
  let fontName = "";
  let size = 0, spacing = 0, scale = 100, rise = 0;
  // The reader's state along one line: the pen, where the last inked glyph
  // ended, the size and rise of the item it read that glyph into, and its
  // two last characters (pdf.js twoLastChars).
  let pen = 0;
  let inkEnd: number | null = null;
  let itemSize = 0, itemRise = 0, itemFont = "";
  let itemOpen = false;
  let lastTwo: [string, string] = [" ", " "];
  // Where the line's pen started along the paragraph's lines, the baseline
  // it runs on.
  let lineStart = 0;
  let lineBaseline: number | null = null;
  const save = (character: string): boolean => {
    const spaced = lastTwo[0] !== " " && lastTwo[1] === " ";
    lastTwo = [lastTwo[1], character];
    return spaced;
  };
  const reset = () => { lastTwo = [" ", " "]; };
  const read = (codes: readonly number[], extra: number) => {
    if (!font) refuse("The rewritten paragraph shows text before it sets a font.");
    if (!codes.length) {
      pen += (spacing + extra) * (scale / 100);
      return;
    }
    const unit = font.codeUnitBytes;
    for (let offset = 0; offset < codes.length; offset += unit) {
      const bytes = Uint8Array.from(codes.slice(offset, offset + unit));
      const code = unit === 1 ? bytes[0] : (bytes[0] << 8) | bytes[1];
      const text = font.decode.get(code);
      if (text === undefined) refuse("The rewritten paragraph draws a code its font does not read.");
      if (PDFJS_INVISIBLE_FORMAT_MARK.test(text)) continue;
      const width = kernel.glyphAdvance(bytes, font, size);
      const after = spacing + (offset + unit >= codes.length ? extra : 0);
      if (text && isReaderBlank(characters(text)[0])) {
        pen += (width + after) * (scale / 100);
        save(" ");
        continue;
      }
      const diacritic = PDFJS_ZERO_WIDTH_DIACRITIC.test(text);
      if (!diacritic && inkEnd !== null) {
        const gap = pen - inkEnd;
        if (Math.abs(gap) > 1e-9) {
          if (gap < PDFJS_NEGATIVE_SPACE_FACTOR * itemSize) {
            reset();
            itemOpen = false;
          } else {
            if (gap <= PDFJS_NOT_A_SPACE_FACTOR * itemSize) reset();
            if (gap <= PDFJS_TRACKING_SPACE_FACTOR * itemSize) {
              if (lastTwo[0] !== " " && lastTwo[1] === " ") {
                reset();
                itemOpen = false;
                lines[lines.length - 1] += " ";
              }
            } else {
              // Past 0.6 of the size the space is an item of its own.
              reset();
              if (gap > PDFJS_SPACE_IN_FLOW_MAX_FACTOR * itemSize) itemOpen = false;
              lines[lines.length - 1] += " ";
            }
            if (Math.abs(rise - itemRise) > PDFJS_VERTICAL_SHIFT_RATIO * itemSize) itemOpen = false;
          }
        }
      }
      if (!itemOpen || itemFont !== fontName) {
        itemOpen = true;
        itemSize = size;
        itemFont = fontName;
      }
      const drawn = diacritic ? 0 : width * (scale / 100);
      if (drawn) {
        inkEnd = pen + drawn;
        itemRise = rise;
      }
      if (save(text)) lines[lines.length - 1] += " ";
      // The reader shows the glyph's characters normalised (a TeX font's
      // micro sign, drawn for the μ the paragraph holds, reads as that μ).
      lines[lines.length - 1] += kernel.normalizeUnicode(text);
      pen += drawn + after * (scale / 100);
    }
  };
  const tokens = block.match(/\/[^\s/[\]<>]+|\[[^\]]*\]|[-+]?(?:\d+(?:\.\d*)?|\.\d+)|[A-Za-z]+/gu) ?? [];
  let operands: string[] = [];
  // The fake-bold pass, marked as no text: the box builder drops the items
  // the reader reads in it (pdfJsOverprintItemIndexes), so the paragraph
  // reads as if it were not there.
  let marked = false;
  for (const token of tokens) {
    if (!/^[A-Za-z]+$/u.test(token)) {
      operands.push(token);
      continue;
    }
    if (token === "BDC" || token === "EMC") {
      marked = token === "BDC";
      operands = [];
      continue;
    }
    if (marked) {
      operands = [];
      continue;
    }
    if (token === "Tf") {
      fontName = operands[0].slice(1);
      font = paragraph.fonts.get(fontName) ?? null;
      size = Number(operands[1]);
      // A change of font or size starts a new item (the reader flushes its
      // item before the next glyph and measures it at the new size).
      itemOpen = false;
    }
    if (token === "Tc") spacing = Number(operands[0]);
    if (token === "Tz") scale = Number(operands[0]);
    if (token === "Ts") rise = Number(operands[0]);
    if (token === "Tm") {
      const [b0, b1, b2, b3] = paragraph.basis;
      const [e, f] = [Number(operands[4]), Number(operands[5])];
      const along = e * b0 + f * b1, across = e * b2 + f * b3;
      if (lineBaseline !== null && Math.abs(across - lineBaseline) <= LAYOUT_TOLERANCE) {
        // A run placed on its line's baseline: the pen stands where the Tm
        // puts it, and the box reads on into it -- an upright run after a
        // slanted one, or a slanted one after an upright one, measured along
        // the baseline both share. (pdf.js itself ends its line before a run
        // set in a skewed matrix of its own, measured in that matrix's frame,
        // but the editor's boxes rejoin the two on their shared baseline by
        // the gap between them: an italic begun inside a word reads as the
        // word, "Attention" leaned from its "t", and its write was refused
        
        pen = along - lineStart;
        itemOpen = false;
      } else {
        // A line placed afresh, below the one before: the reader ends its
        // line there.
        lines.push("");
        pen = 0;
        inkEnd = null;
        itemOpen = false;
        reset();
        lineStart = along;
      }
      lineBaseline = across;
    }
    if (token === "TJ") {
      const elements = operands[0].slice(1, -1).match(/<[0-9A-F]*>|[-+]?(?:\d+(?:\.\d*)?|\.\d+)/gu) ?? [];
      let pending: number[] = [];
      for (const element of elements) {
        if (element.startsWith("<")) {
          const hex = element.slice(1, -1);
          for (let offset = 0; offset < hex.length; offset += 2) pending.push(Number.parseInt(hex.slice(offset, offset + 2), 16));
          continue;
        }
        const value = Number(element);
        if (value === 0) continue;
        read(pending, (-value * size) / 1_000);
        pending = [];
      }
      if (pending.length) read(pending, 0);
    }
    operands = [];
  }
  return joinLines(lines.map((line) => line.replace(/ +$/u, "")));
}

/** The new text as the reader shows it: each run of white space one space,
 * none at either end (the reader writes no white-space glyph, and reads a
 * space only between two it inks). */
function readerSpelling(text: string): string {
  return characters(text).map((character) => isReaderBlank(character) ? " " : character).join("")
    .replace(/ {2,}/gu, " ").replace(/^ +| +$/gu, "");
}

/** A paragraph's lines joined as the reader joins them. */
function joinLines(lines: readonly string[]): string {
  let text = "";
  for (const line of lines) {
    if (!line) continue;
    if (text && lineBoundaryNeedsSpace(text, line)) text += " ";
    text += line;
  }
  return text;
}

/** Mirror of the planner's visualLineBoundaryNeedsSpace. */
function lineBoundaryNeedsSpace(previous: string, current: string): boolean {
  if (/\s$/u.test(previous) || /^\s/u.test(current)) return false;
  const last = characters(previous).at(-1) ?? "", first = characters(current)[0] ?? "";
  if (CJK.test(last) && CJK.test(first)) return false;
  if (!/\s/u.test(previous) && !/\s/u.test(current) &&
    (characters(previous).length >= 24 || characters(current).length >= 24)) return false;
  return true;
}

/**
 * The bytes a rewrite leaves in its operator span: every owned show emptied
 * into its own operator's form (or into its advance where it keeps its pen),
 * every other byte of the span kept, and the new block after the span's last
 * `ET`. An /ActualText cluster span (a ligature's "fi") all of whose shows
 * the rewrite empties goes with them: its `BDC` and `EMC` become a space
 * each, so it no longer names letters the page does not draw there. Replay
 * builds the same bytes from the same span.
 */
/** For each character of `replacement`, the index of the character of
 * `original` it is kept as (-1 for one typed): the common `prefix` and
 * `suffix`, and between them the longest common subsequence. */
function keptSourceCharacters(
  original: readonly string[],
  replacement: readonly string[],
  prefix: number,
  suffix: number,
): number[] {
  const kept = replacement.map(() => -1);
  for (let index = 0; index < prefix; index += 1) kept[index] = index;
  for (let index = 0; index < suffix; index += 1) {
    kept[replacement.length - 1 - index] = original.length - 1 - index;
  }
  const from = original.slice(prefix, original.length - suffix);
  const to = replacement.slice(prefix, replacement.length - suffix);
  // Past this many cells a middle is taken as typed anew.
  if (!from.length || !to.length || from.length * to.length > 4_000_000) return kept;
  const width = to.length + 1;
  const lengths = new Uint16Array((from.length + 1) * width);
  for (let i = from.length - 1; i >= 0; i -= 1) {
    for (let j = to.length - 1; j >= 0; j -= 1) {
      lengths[i * width + j] = from[i] === to[j] ? lengths[(i + 1) * width + j + 1] + 1
        : Math.max(lengths[(i + 1) * width + j], lengths[i * width + j + 1]);
    }
  }
  for (let i = 0, j = 0; i < from.length && j < to.length;) {
    if (from[i] === to[j]) {
      kept[prefix + j] = prefix + i;
      i += 1;
      j += 1;
    } else if (lengths[(i + 1) * width + j] >= lengths[i * width + j + 1]) i += 1;
    else j += 1;
  }
  return kept;
}

/** The rule blocks of this writer's form stacked beside the text objects of
 * `shows` (pdfNativeTextRuleBlocksBeside), in stream order, each once. */
/** One operator of a content stream outside its text objects (a text
 * object is one entry, `BT`, spanning to just after its `ET`), with the
 * operands written before it. */
type OutsideTextOperator = Readonly<{ operator: string; operands: readonly string[]; start: number; end: number }>;

const OPERATOR_DELIMITER = /[\s()<>[\]{}/%]/u;

/** The operators of `source` from `from` on, outside its text objects; the
 * scan ends (null) at what it does not read: a string or comment outside a
 * text object, an inline image. */
function* operatorsOutsideText(source: string, from: number): Generator<OutsideTextOperator | null> {
  let operands: string[] = [];
  let operandStart = -1;
  let at = from;
  while (at < source.length) {
    const character = source[at];
    if (/\s/u.test(character)) { at += 1; continue; }
    const begin = at;
    if (character === "%" || character === "(") { yield null; return; }
    if (character === "/") {
      at += 1;
      while (at < source.length && !OPERATOR_DELIMITER.test(source[at])) at += 1;
    } else if (source.startsWith("<<", at)) {
      let depth = 0;
      while (at < source.length) {
        if (source.startsWith("<<", at)) { depth += 1; at += 2; continue; }
        if (source.startsWith(">>", at)) { depth -= 1; at += 2; if (depth === 0) break; continue; }
        if (source[at] === "(") { yield null; return; }
        at += 1;
      }
      if (depth !== 0) { yield null; return; }
    } else if (character === "<") {
      const close = source.indexOf(">", at);
      if (close < 0) { yield null; return; }
      at = close + 1;
    } else if (character === "[" || character === "]" || character === "{" || character === "}") {
      at += 1;
    } else if (/[-+.\d]/u.test(character)) {
      at += 1;
      while (at < source.length && /[\d.]/u.test(source[at])) at += 1;
    } else {
      at += 1;
      while (at < source.length && !OPERATOR_DELIMITER.test(source[at])) at += 1;
      const operator = source.slice(begin, at);
      if (operator === "BI") { yield null; return; }
      if (operator === "BT") {
        const object = enclosingTextObject(source, begin);
        yield Object.freeze({ operator, operands: Object.freeze([]), start: begin, end: object.end });
        at = object.end;
      } else {
        yield Object.freeze({ operator, operands: Object.freeze(operands), start: operandStart >= 0 ? operandStart : begin,
          end: at });
      }
      operands = [];
      operandStart = -1;
      continue;
    }
    if (operandStart < 0) operandStart = begin;
    operands.push(source.slice(begin, at));
  }
}

/** The operators that only set the fill colour, mark content or stroke
 * colour: a rule another writer drew may stand among them. */
const RULE_NEIGHBOUR_OPERATORS = new Set(["g", "rg", "k", "G", "RG", "K", "BDC", "BMC", "EMC", "MP", "DP",
  // The text state a run is drawn in, set outside its text object (a box a
  // group move wrote outside its clip sets each run so), and stroke widths.
  "Tc", "Tw", "Tz", "TL", "Tf", "Tr", "Ts", "w", "J", "j", "M", "d"]);
















function paragraphLinkRules(input: {
  readonly source: string;
  readonly shows: readonly OwnedShow[];
  readonly basis: readonly [number, number, number, number];
  readonly blockMatrix: Matrix;
  readonly lineSizes: readonly number[];
  /** This writer's own rule blocks beside the paragraph's text objects
   * (paragraphRuleSpans): read past, as the block draws other writers' rules
   * after its strikethroughs. */
  readonly skip: readonly Readonly<{ start: number; end: number }>[];
}): Readonly<{ spans: readonly Readonly<{ start: number; end: number }>[];
  marks: ReadonlyMap<number, ParagraphLinkRule> }> | null {
  const { source, shows, basis, blockMatrix, lineSizes, skip } = input;
  const [b0, b1, b2, b3] = basis;
  if (Math.abs(b1) > 1e-9 || Math.abs(b2) > 1e-9 || !(Math.abs(b0) > 1e-9) || !(Math.abs(b3) > 1e-9)) return null;
  if (!shows.every((owned) => owned.show.graphicsMatrix.every((value, at) => Math.abs(value - blockMatrix[at]) <= 1e-9))) {
    return null;
  }
  const objects = shows.map((owned) => enclosingTextObject(source, owned.start));
  const first = Math.min(...objects.map((object) => object.start));
  const last = Math.max(...objects.map((object) => object.end));
  const spans: Readonly<{ start: number; end: number }>[] = [];
  const marks = new Map<number, ParagraphLinkRule>();
  let fill: PdfNativeDeviceColor | null = null;
  let depth = 0;
  const fills: (PdfNativeDeviceColor | null)[] = [];
  let rect: Readonly<{ start: number; numbers: readonly number[] }> | null = null;
  // The runs a rule lies under, and its format on them; null when it is no
  // run's rule.
  const ruleOf = (numbers: readonly number[]) => {
    const [x, y, w, h] = numbers;
    const us = [x / b0, (x + w) / b0], vs = [y / b3, (y + h) / b3];
    const [ulo, uhi] = [Math.min(...us), Math.max(...us)], [vlo, vhi] = [Math.min(...vs), Math.max(...vs)];
    const thickness = vhi - vlo;
    if (!(thickness > 0) || !(uhi - ulo > thickness * 2)) return null;
    // Under the runs whose own baseline (a raised reference mark's, raised)
    // it lies just below, along their ink.
    const covered: number[] = [];
    for (const [index, owned] of shows.entries()) {
      if (!owned.inked || !(owned.inked.end > owned.inked.start)) continue;
      const baseline = owned.v + owned.rise;
      if (vhi > baseline + 0.02 * owned.size || vlo < baseline - 0.45 * owned.size || thickness > 0.2 * owned.size) continue;
      const overlap = Math.min(owned.inked.end, uhi) - Math.max(owned.inked.start, ulo);
      if (overlap > 0 && overlap >= 0.5 * (owned.inked.end - owned.inked.start)) covered.push(index);
    }
    if (!covered.length || covered.some((index) => shows[index].line !== shows[covered[0]].line)) return null;
    const slack = 0.5 * lineSizes[shows[covered[0]].line];
    const from = Math.min(...covered.map((index) => shows[index].inked!.start));
    const to = Math.max(...covered.map((index) => shows[index].inked!.end));
    if (ulo < from - slack || uhi > to + slack) return null;
    return { covered, vlo, thickness };
  };
  let past = false;
  for (const operator of operatorsOutsideText(source, first)) {
    // Among the paragraph's text objects anything unread reads none; after
    // the last, the rules end there.
    if (operator === null) { if (past) break; return null; }
    if (skip.some((span) => operator.start >= span.start && operator.end <= span.end)) continue;
    const among = operator.start < last;
    if (!among) past = true;
    const stop = () => among ? "none" as const : "end" as const;
    let outcome: "none" | "end" | null = null;
    const name = operator.operator;
    if (rect && name !== "f" && name !== "F" && name !== "f*") outcome = stop();
    else if (name === "BT") { if (!among) outcome = "end"; }
    else if (name === "q") { fills.push(fill); depth += 1; }
    // A matrix set among the paragraph's runs (each run drawn under its own
    // `q … cm`, as a group move outside its clip writes a box): a rule there
    // is still read only where it lies under a run's ink (ruleOf). After
    // the last run, the rules end.
    else if (name === "cm") { if (!among) outcome = "end"; }
    else if (name === "Q") {
      // Closing a state opened before the paragraph's first text object (a
      // writer that draws each run under its own q): among its runs the
      // matrix is the next run's (no matrix is set before it), the colour
      // unknown until set; after them the rules end.
      if (depth > 0) { depth -= 1; fill = fills.pop() ?? null; }
      else if (!among) outcome = "end";
      else fill = null;
    }
    else if (name === "re") {
      const numbers = operator.operands.map(Number);
      if (numbers.length !== 4 || numbers.some((value) => !Number.isFinite(value))) outcome = stop();
      else rect = { start: operator.start, numbers };
    } else if (name === "f" || name === "F" || name === "f*") {
      const read = rect && fill ? ruleOf(rect.numbers) : null;
      if (!read || !rect) outcome = stop();
      else {
        spans.push(Object.freeze({ start: rect.start, end: operator.end }));
        for (const index of read.covered) {
          if (marks.has(index)) continue;
          const size = shows[index].size;
          marks.set(index, Object.freeze({ color: fill!, offset: (read.vlo - shows[index].v - shows[index].rise) / size,
            thickness: read.thickness / size }));
        }
        rect = null;
      }
    } else if (name === "g" || name === "rg" || name === "k") {
      const components = operator.operands.map(Number);
      const kind = name === "g" ? "gray" as const : name === "rg" ? "rgb" as const : "cmyk" as const;
      fill = components.length === (kind === "gray" ? 1 : kind === "rgb" ? 3 : 4) &&
        components.every((value) => Number.isFinite(value) && value >= 0 && value <= 1)
        ? Object.freeze({ kind, components: Object.freeze(components) }) as PdfNativeDeviceColor : null;
    } else if (!RULE_NEIGHBOUR_OPERATORS.has(name)) outcome = stop();
    if (outcome === "none") return null;
    if (outcome === "end") break;
  }
  return spans.length ? Object.freeze({ spans: Object.freeze(spans), marks }) : null;
}

/** The rules another writer drew that a rewrite of a paragraph took away
 * from its span (paragraphLinkRules): every rectangle filled outside its text
 * objects -- the rewrite takes all of them or none. */
function paragraphLinkRuleSpans(source: string): Readonly<{ start: number; end: number }>[] {
  const spans: Readonly<{ start: number; end: number }>[] = [];
  // The span opens inside its first text object (at its first show): read
  // from that object's end on.
  let from = -1;
  for (let at = 0; at < source.length && from < 0;) {
    const character = source[at];
    if (character === "(") {
      let depth = 0;
      for (; at < source.length; at += 1) {
        if (source[at] === "\\") { at += 1; continue; }
        if (source[at] === "(") depth += 1;
        else if (source[at] === ")" && --depth === 0) { at += 1; break; }
      }
      continue;
    }
    if (character === "E" && source[at + 1] === "T" && (at === 0 || OPERATOR_DELIMITER.test(source[at - 1])) &&
      (at + 2 >= source.length || OPERATOR_DELIMITER.test(source[at + 2]))) from = at + 2;
    at += 1;
  }
  if (from < 0) return spans;
  let rect: number | null = null;
  for (const operator of operatorsOutsideText(source, from)) {
    if (operator === null) return [];
    if (operator.operator === "re") { rect = operator.start; continue; }
    if (rect !== null && (operator.operator === "f" || operator.operator === "F" || operator.operator === "f*")) {
      spans.push(Object.freeze({ start: rect, end: operator.end }));
    }
    rect = null;
  }
  return spans;
}

function paragraphRuleSpans(
  streamBytes: Uint8Array,
  source: string,
  shows: readonly Readonly<{ start: number }>[],
): Readonly<{ start: number; end: number }>[] {
  const objects = new Map<number, Readonly<{ start: number; end: number }>>();
  for (const show of shows) {
    const object = enclosingTextObject(source, show.start);
    objects.set(object.start, object);
  }
  const spans = new Map<number, Readonly<{ start: number; end: number }>>();
  for (const object of objects.values()) {
    for (const block of pdfNativeTextRuleBlocksBeside(streamBytes, object)) {
      spans.set(block.start, Object.freeze({ start: block.start, end: block.end }));
    }
  }
  return [...spans.values()].sort((left, right) => left.start - right.start);
}

const PDF_NUMBER = String.raw`-?(?:\d+\.?\d*|\.\d+)`;
/** A rule as a typesetter draws it: a stroked line, or a filled rectangle. */
const STROKED_RULE = new RegExp(String.raw`q\s+${PDF_NUMBER}\s+w\s+(${PDF_NUMBER})\s+(${PDF_NUMBER})\s+m\s+${PDF_NUMBER}\s+${PDF_NUMBER}\s+l\s+S\s+Q`, "gu");
const FILLED_RULE = new RegExp(String.raw`(${PDF_NUMBER})\s+(${PDF_NUMBER})\s+${PDF_NUMBER}\s+${PDF_NUMBER}\s+re\s+f\b`, "gu");
const shortNumber = (value: number) => {
  const text = value.toFixed(4).replace(/\.?0+$/u, "");
  return text === "-0" ? "0" : text;
};

/**
 * What stands in a gap the edit moved goes with the text after it: another
 * box's glyph (the radical before a π in a TeX formula, its own show above the
 * line), by a `Td` there and back around its show, and the rule over the words
 * after it (the radical's bar), under a translation of its own. They stood where
 * they were before, over the words the edit had moved. Left as it is where a
 * show runs on from its pen, or the span changes its space.
 */
function paragraphCarriedInGaps(
  kernel: MraiNativeParagraphRewriteKernel,
  paragraph: ReadParagraph,
  gaps: readonly MovedGap[],
): readonly Readonly<{ at: number; text: string }>[] {
  if (!gaps.length) return [];
  const { source, spanStart, insertion, basis, blockMatrix } = paragraph;
  const span = source.slice(spanStart, insertion);
  if (/\bcm\b/u.test(span)) return [];
  const writes: { at: number; text: string }[] = [];
  const within = (gap: MovedGap, along: number, across: number) => along >= gap.from - 1 && along <= gap.to + 1 &&
    Math.abs(across - gap.across) <= 1.2 * gap.size;
  const toBlock = invertMatrix(blockMatrix);
  for (const other of otherInkedExtents(kernel, paragraph)) {
    const gap = gaps.find((entry) => within(entry, other.start, other.across) && within(entry, other.end ?? other.start, other.across));
    if (!gap) continue;
    const after = source.slice(other.show.end, insertion);
    // The next show must place itself: this one's pen is not where it was.
    const nextShow = /\b(?:TJ|Tj)\b|'|"/u.exec(after);
    const placed = /\b(?:Td|TD|Tm|T\*|ET)\b/u.exec(after);
    if (nextShow && (!placed || placed.index > nextShow.index)) continue;
    const linear = composeMatrix(composeMatrix(other.show.textMatrix, other.show.graphicsMatrix), toBlock);
    const inverse = invertMatrix([linear[0], linear[1], linear[2], linear[3], 0, 0]);
    const [tx, ty] = applyMatrix(inverse, gap.dx * basis[0], gap.dx * basis[1]);
    writes.push({ at: other.show.start, text: ` ${shortNumber(tx)} ${shortNumber(ty)} Td ` },
      { at: other.show.end, text: ` ${shortNumber(-tx)} ${shortNumber(-ty)} Td` });
  }
  for (const pattern of [STROKED_RULE, FILLED_RULE]) {
    for (const match of span.matchAll(pattern)) {
      const x = Number(match[1]), y = Number(match[2]);
      const gap = gaps.find((entry) => within(entry, x * basis[0] + y * basis[1], x * basis[2] + y * basis[3]));
      if (!gap) continue;
      const at = spanStart + match.index;
      writes.push({ at, text: `q 1 0 0 1 ${shortNumber(gap.dx * basis[0])} ${shortNumber(gap.dx * basis[1])} cm ` },
        { at: at + match[0].length, text: " Q" });
    }
  }
  return writes;
}

function rewrittenSpan(
  original: string,
  spanStart: number,
  shows: readonly Readonly<{ start: number; end: number }>[],
  pens: PdfNativeParagraphRewriteEvidence["pens"],
  block: Uint8Array | null,
  rules: readonly Readonly<{ start: number; end: number }>[] = [],
  /** Written into the source between the shows: where another show or a rule
   * stands in a gap the edit moved, what moves it (absolute offsets). */
  insertions: readonly Readonly<{ at: number; text: string }>[] = [],
): Uint8Array {
  const penOf = new Map(pens.map((pen) => [`${pen.start}:${pen.end}`, pen]));
  const owned = new Set(shows.map((show) => `${show.start - spanStart}:${show.end - spanStart}`));
  // (And the rules the paragraph's runs were drawn with: the block draws
  // the ones its characters keep again.)
  const dropped = [...(pdfNativeActualTextClusterSpans(original) ?? [])
    .filter((span) => span.shows.length > 0 && span.shows.every((show) => owned.has(`${show.start}:${show.end}`)))
    .flatMap((span) => [[span.openStart, span.openEnd, " "], [span.closeStart, span.closeEnd, " "]] as const),
  ...rules.map((rule) => [rule.start - spanStart, rule.end - spanStart, " "] as const),
  ...insertions.map((insertion) => [insertion.at - spanStart, insertion.at - spanStart, insertion.text] as const)]
    .sort((left, right) => left[0] - right[0] || left[1] - right[1]);
  // The source between two shows, less the markers of the spans that go.
  const kept = (from: number, to: number) => {
    let text = "";
    let at = from;
    for (const [start, end, written] of dropped) {
      if (start === end ? start < from || start >= to : end <= from || start >= to) continue;
      text += `${original.slice(at, start)}${written}`;
      at = end;
    }
    return text + original.slice(at, to);
  };
  let out = "";
  let cursor = 0;
  for (const show of shows) {
    const start = show.start - spanStart, end = show.end - spanStart;
    if (start < cursor || end > original.length) refuse("The paragraph's shows are out of their span.");
    const pen = penOf.get(`${show.start}:${show.end}`);
    const spacing = pen?.characterSpacing ?? 0;
    const emptied = pen?.chunk
      ? mraiNativeEmptyChunkForm(original.slice(start, end), pen.chunk.codeStart, pen.chunk.codeEnd, pen.number,
        spacing)
      : pen !== undefined ? mraiNativeEmptyShowPenForm(original.slice(start, end), pen.number, spacing)
        : mraiNativeEmptyShowsForm(original.slice(start, end));
    if (emptied === null) refuse("A show of this paragraph cannot be emptied.");
    out += kept(cursor, start) + emptied;
    cursor = end;
  }
  out += kept(cursor, original.length);
  // The block set is the paragraph's box, named in the page (box-mark.ts):
  // its rows stay one box however the page reads them (they run onto the
  // next paragraph's baselines at a larger line spacing, a justified gap
  // reads as a line break). A bytes-published write carries no rows the
  // journal could keep it by, and the paragraph came back in pieces, one of
  
  return block
    ? Uint8Array.from([...asciiBytes(out), 0x0a, ...MRAI_NATIVE_BOX_MARK_OPEN, ...block, ...MRAI_NATIVE_BOX_MARK_CLOSE])
    : asciiBytes(out);
}

/** The request and evidence both answers of this tier share. */
async function paragraphEvidence(
  paragraph: ReadParagraph,
  replacementText: string,
  set: ReturnType<typeof setParagraph> | null,
  frameRequested = false,
): Promise<Readonly<{ request: PdfNativeContentStreamTransactionRequest; evidence: PdfNativeListTextFidelityEvidence }>> {
  const { streamBytes, source, spanStart, insertion, firstPlan, layout } = paragraph;
  const originalBytes = streamBytes.slice(spanStart, insertion);
  // The owned shows and their fake-bold passes, emptied in stream order.
  const emptied = [...paragraph.shows.map((show) => ({ start: show.start, end: show.end })), ...paragraph.twins]
    .sort((left, right) => left.start - right.start);
  const replacement = rewrittenSpan(source.slice(spanStart, insertion), spanStart, emptied,
    paragraph.pens, set?.bytes ?? null, paragraph.rules, set?.carried ?? []);
  // One edit writes at most as many bytes as a transaction carries; a
  // paragraph whose shows lie that far apart in its stream (a drawing set
  // between the lines of a CAD label) is left to the routes after this one.
  if (replacement.length > PARAGRAPH_REWRITE_MAX_REPLACEMENT_BYTES) {
    refuse("The paragraph's shows lie too far apart in their stream for one edit to rewrite.");
  }
  const sha256 = await sha256PdfBytes(originalBytes);
  const originalCodes = paragraph.shows.flatMap((show) => [...show.show.codes]);
  const originalAdvance = layout.lineExtents.reduce((sum, extent) => sum + (extent.right - extent.left), 0);
  const rewrite: PdfNativeParagraphRewriteEvidence = Object.freeze({
    schemaVersion: MRAI_NATIVE_PARAGRAPH_REWRITE_SCHEMA_VERSION,
    target: Object.freeze({ objectNumber: firstPlan.request.identity.objectNumber,
      generation: firstPlan.request.identity.generation }),
    items: Object.freeze(paragraph.plans.map((plan) => plan?.evidence ?? null)),
    shows: Object.freeze(paragraph.shows.map((show) => Object.freeze({ start: show.start, end: show.end }))),
    pens: paragraph.pens,
    ...(paragraph.twins.length ? { twins: paragraph.twins } : {}),
    originalBytes: Object.freeze([...originalBytes]),
    pins: Object.freeze(paragraph.pins.map((pin) => Object.freeze({ objectNumber: pin.objectNumber,
      generation: pin.generation, role: pin.role }))),
    block: set ? Object.freeze({
      start: insertion,
      bytes: Object.freeze([...set.bytes]),
      alignment: layout.alignment,
      basis: paragraph.basis,
      frameLeft: layout.frameLeft,
      frameRight: layout.frameRight,
      firstLeft: layout.firstLeft,
      restLeft: layout.restLeft,
      firstBaseline: layout.firstBaseline,
      lineStep: layout.lineStep,
      originalLineCount: layout.lineCount,
      lineCount: set.lines.length,
      lines: Object.freeze(set.lines.map((line) => Object.freeze({ start: line.start, end: line.end }))),
      ...(set.lineBaselines ? { lineBaselines: set.lineBaselines } : {}),
      ...(frameRequested ? { frameRequested: true as const } : {}),
      ...(paragraph.sourceLineStep !== undefined && paragraph.sizeScale === undefined
        ? { lineStepRequested: true as const } : {}),
      ...(paragraph.sizeScale !== undefined && paragraph.sourceLineStep !== undefined &&
        Math.abs(layout.lineStep - paragraph.sourceLineStep * paragraph.sizeScale) > LAYOUT_TOLERANCE
        ? { lineStepRequested: true as const } : {}),
      ...(paragraph.sizeScale !== undefined ? { sizeScale: paragraph.sizeScale } : {}),
    }) : null,
  });
  // The paragraph is described the way the in-place route describes one:
  // by its first item's font, at its first line, over the whole span.
  const {
    positionedRunGroup: _group, positionedLines: _lines, consolidatedRuns: _consolidated,
    positionedRestore: _restore, positionedFrameCapacity: _frame, sameLineLayoutState: _layout,
    sameLineCapacity: _capacity, lineShift: _shift, readerChunk: _chunk, propertyMutation: _property,
    paragraphReflow: _reflow, ...first
  } = firstPlan.evidence;
  const evidence: PdfNativeListTextFidelityEvidence = Object.freeze({
    ...first,
    schemaVersion: PARAGRAPH_TEXT_SCHEMA_VERSION,
    layoutKind: "paragraph-line" as const,
    selection: paragraph.selection,
    wrapping: "fixed-frame-multiline" as const,
    linePlan: null,
    listAnchor: null,
    originalText: paragraph.selection.item.str,
    originalDisplayText: paragraph.selection.item.str,
    replacementText,
    replacementDisplayText: replacementText,
    originalCodeBytes: Object.freeze(originalCodes),
    replacementCodeBytes: Object.freeze(set ? [...set.codes] : originalCodes),
    textX: layout.firstLeft,
    baseline: layout.firstBaseline,
    lineSpacing: paragraph.sourceLineStep ?? layout.lineStep,
    originalAdvance,
    replacementAdvance: set ? set.advance : originalAdvance,
    textOperatorSpan: Object.freeze({ start: spanStart, end: insertion, sha256 }),
    paragraphRewrite: rewrite,
  });
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: firstPlan.request.identity,
    expectedObjectSha256: firstPlan.request.expectedObjectSha256,
    expectedStreamSha256: firstPlan.request.expectedStreamSha256,
    operatorSpan: Object.freeze({ start: spanStart, end: insertion, expectedSha256: sha256 }),
    replacement,
    untouchedObjects: paragraph.pins,
    ...(firstPlan.request.pageContentConsolidation
      ? { pageContentConsolidation: firstPlan.request.pageContentConsolidation } : {}),
  });
  return Object.freeze({ request, evidence });
}

/**
 * The read proof of the paragraph rewrite: every item proved alone, so the
 * paragraph owns its shows and may open for editing. Its blanking request
 * empties those shows, the display intermediate the editor draws over.
 */
export async function proveMraiNativeParagraphRewriteOwnership(
  kernel: MraiNativeParagraphRewriteKernel,
  selection: PdfNativeListTextSelectionEvidence,
): Promise<Readonly<{ plan: PdfNativeListTextReplacementPlan; blankingRequest: PdfNativeContentStreamTransactionRequest }>> {
  const paragraph = await readParagraph(kernel, selection);
  const { request, evidence } = await paragraphEvidence(paragraph, selection.item.str, null);
  const sourceAppearance: PdfNativeSourceTextAppearanceRun[] =
    paragraph.plans.flatMap((plan) => plan?.sourceAppearance ?? []);
  return Object.freeze({
    plan: Object.freeze({
      request,
      evidence,
      properties: null,
      ...(sourceAppearance.length ? { sourceAppearance: Object.freeze(sourceAppearance) } : {}),
    }),
    blankingRequest: request,
  });
}

/** The characters of the new text no font of the paragraph writes, each once,
 * in the order the text first holds them: what a substitute face is asked
 * for (setParagraph sets exactly these in it). */
function charactersNoParagraphFontWrites(
  kernel: MraiNativeParagraphRewriteKernel,
  paragraph: ReadParagraph,
  replacementText: string,
): string {
  const missing: string[] = [];
  for (const character of characters(replacementText)) {
    if (isBlank(character) || missing.includes(character)) continue;
    const written = paragraph.fontOrder.some((name) => {
      try {
        kernel.encodeText(character, paragraph.fonts.get(name)!, paragraph.preferred.get(name)!);
        return true;
      } catch (error) {
        if (!(error instanceof PdfNativeTransactionError)) throw error;
        return false;
      }
    });
    if (!written) missing.push(character);
  }
  return missing.join("");
}

/** The paragraph rewrite of one edit: the plan a save writes. With a
 * substitute, a character no font of the paragraph has is set in the face
 * the substitute installs instead of refusing the edit; the plan then names
 * that face and the characters it draws. */
/**
 * The text with a line break where the source ends a line short of its frame
 * by more than the next line's first piece would take (up to its first
 * space, or past its first hyphen or slash): a break the writer typed
 * (Shift+Enter, a cell's second link), not one the measure made. Kept where
 * the edit leaves the text on both sides of it as it was; a frame the
 * reader resized keeps none (its lines are the reader's to break).
 */
function withSourceHardBreaks(paragraph: ReadParagraph, sourceText: string, replacementText: string,
  resized: boolean): string {
  const extents = paragraph.layout.lineExtents;
  if (resized || replacementText.includes("\n") || extents.length < 2) return replacementText;
  const before = characters(sourceText);
  const starts: number[] = [];
  const counts: number[] = [];
  for (const [index, character] of before.entries()) {
    const line = paragraph.shows[paragraph.characterShows[index] ?? -1]?.line;
    if (line === undefined) continue;
    counts[line] = (counts[line] ?? 0) + 1;
    if (starts[line] === undefined && !isBlank(character)) starts[line] = index;
  }
  const em = paragraph.layout.lineStep > 0 ? paragraph.layout.lineStep / 1.2 : 0;
  const breaks: number[] = [];
  for (let line = 0; line + 1 < extents.length; line += 1) {
    const next = starts[line + 1];
    const count = counts[line + 1] ?? 0;
    const width = extents[line + 1].right - extents[line + 1].left;
    if (next === undefined || !(count > 0) || !(width > 0) || !(em > 0)) continue;
    let piece = 0;
    while (next + piece < before.length && !isBlank(before[next + piece])) {
      piece += 1;
      if (/[-/\u2010\u2013]/u.test(before[next + piece - 1])) break;
    }
    const free = paragraph.layout.frameRight - extents[line].right;
    if (free > (width / count) * piece + em * 0.5) breaks.push(next);
  }
  if (!breaks.length) return replacementText;
  const after = characters(replacementText);
  let head = 0;
  while (head < before.length && head < after.length && before[head] === after[head]) head += 1;
  let tail = 0;
  while (tail < before.length - head && tail < after.length - head &&
    before[before.length - 1 - tail] === after[after.length - 1 - tail]) tail += 1;
  const at = breaks.flatMap((index) => index <= head ? [index]
    : index >= before.length - tail ? [index + after.length - before.length] : []);
  const kept = [...after];
  for (const index of [...new Set(at)].sort((left, right) => right - left)) {
    if (index <= 0 || index >= kept.length) continue;
    // A break that would leave a line with nothing on it is not kept.
    const line = kept.slice(index).join("").split("\n")[0];
    const lineBefore = kept.slice(0, index).join("").split("\n").at(-1) ?? "";
    if (!line.trim() || !lineBefore.trim()) continue;
    kept.splice(index, 0, "\n");
  }
  return kept.join("");
}

export async function planMraiNativeParagraphRewrite(
  kernel: MraiNativeParagraphRewriteKernel,
  selection: PdfNativeListTextSelectionEvidence,
  replacementText: string,
  substitute?: MraiNativeParagraphRewriteSubstitute,
  /** The frame the reader resized the paragraph to: its text re-wraps in it,
   * whether or not the text changed. */
  frame?: MraiNativeParagraphRewriteFrame,
  /** Formats the edit asks of the new text (a family, size, colour, rise,
   * weight or lean of the whole paragraph or of part of it), resolved once
   * the paragraph is read. */
  resolveFormats?: MraiNativeParagraphRewriteFormatsResolver,
): Promise<PdfNativeListTextReplacementPlan & Readonly<{ substitute?: Readonly<{ resourceName: string; text: string }> }>> {
  if (replacementText === selection.item.str && frame?.right === undefined && frame?.lineStep === undefined &&
    frame?.sizeScale === undefined && !resolveFormats) {
    refuse("Native text replacement is unchanged.");
  }
  if (frame?.sizeScale !== undefined && frame.right === undefined) {
    refuse("A paragraph set at another size keeps the frame its request names.");
  }
  if (!replacementText || replacementText.length > 2_000) refuse("Replacement text is empty or too long.");
  // The paragraph is set anew from the reader's text. A right-to-left run
  // arrives from the reader reordered, and set from that it would be drawn
  // in the reader's order: changing a number beside an Arabic word reversed
  
  // drawn order or refuse it; this one has no drawn order to keep.
  if (pdfTextHasRightToLeftRun(replacementText) || pdfTextHasRightToLeftRun(selection.item.str)) {
    refuse(MRAI_NATIVE_REORDERED_TEXT_MESSAGE);
  }
  const read = await readParagraph(kernel, selection);
  refuseGlyphsReadAsBlanks(read);
  const framed = frame?.right !== undefined ? paragraphInRequestedFrame(kernel, read, { right: frame.right }) : read;
  const spaced = frame?.lineStep !== undefined ? paragraphAtRequestedLineStep(framed, frame.lineStep) : framed;
  let paragraph = frame?.sizeScale !== undefined
    ? paragraphAtRequestedSize(spaced, frame.sizeScale, spaced.sourceLineStep !== undefined) : spaced;
  if (replacementText === selection.item.str && frame?.right === undefined && paragraph.sourceLineStep === undefined &&
    !resolveFormats) {
    refuse(MRAI_NATIVE_PARAGRAPH_FRAME_UNCHANGED_MESSAGE);
  }
  // A line the source ends short of its frame keeps that break (set again in
  // its frame, a table cell's two links on two lines ran into one line,
  
  replacementText = withSourceHardBreaks(paragraph, selection.item.str, replacementText, frame?.right !== undefined);
  // The formats, on the paragraph as read: the faces they name read back as
  // the paragraph's own fonts do; a larger size sets its own line lower, and
  // an alignment asked sets every line by it.
  const formats = resolveFormats
    ? await resolveFormats({ request: paragraph.firstPlan.request, evidence: paragraph.firstPlan.evidence,
      pageContent: paragraph.pageContent,
      text: replacementText.replace(/\n/gu, "") })
    : null;
  if (formats) {
    if (formats.characters.length !== characters(replacementText.replace(/\n/gu, "")).length) {
      refuse("The formats do not name the new text's characters.");
    }
    const faces = formats.characters.flatMap((format) => format?.face ? [format.face] : []);
    
    // 一样的逻辑」): a line the formats widen takes the room beside it the
    // paragraph's own lines leave free, and wraps onto the next line where
    // that room ends; the paragraph grows by lines, never wider.
    paragraph = Object.freeze({
      ...paragraph,
      // A face of the paragraph keeps its own metrics; one it did not use
      // reads back by the format's.
      fonts: new Map([...faces.map((face) => [face.resourceName, face.font] as const), ...paragraph.fonts]),
      // A larger size sets its own line lower (setParagraph's baselines):
      // the paragraph's step stays.
      layout: Object.freeze({ ...paragraph.layout,
        ...(formats.alignment ? { alignment: formats.alignment } : {}),
        ...(formats.alignment === "center" && paragraph.layout.alignment !== "center"
          ? centredParagraphFrame(kernel, paragraph) : {}),
        ...(formats.alignment === "right" && paragraph.layout.alignment !== "right"
          ? rightParagraphFrame(kernel, paragraph) : {}) }),
    });
    if (formats.leftIndent !== undefined && formats.leftIndent !== 0) {
      paragraph = Object.freeze({ ...paragraph, layout: indentedParagraphLayout(paragraph, formats.leftIndent) });
    }
    if (formats.firstLineIndent !== undefined) {
      paragraph = Object.freeze({ ...paragraph, layout: firstLineIndentedLayout(paragraph, formats.firstLineIndent) });
    }
    if (formats.lineStep !== undefined) paragraph = paragraphAtRequestedLineStep(paragraph, formats.lineStep);
  }
  // Kept split, a box's first letter stayed a run in its own subset beside
  
  const wholly = substitute?.faceName !== undefined && paragraph.fontOrder.length > 0 &&
    paragraph.fontOrder.every((name) => paragraph.fonts.get(name)!.faceName === substitute.faceName);
  const missing = !substitute ? "" : wholly
    ? [...new Set(characters(replacementText.replace(/\n/gu, "")).filter((character) => !isBlank(character)))].join("")
    : charactersNoParagraphFontWrites(kernel, paragraph, replacementText.replace(/\n/gu, ""));
  const installed = missing ? await substitute!.install(missing) : null;
  let set: ReturnType<typeof setParagraph>;
  try {
    set = setParagraph(kernel, paragraph, replacementText, installed, formats, wholly);
  } catch (error) {
    // A paragraph at the foot of its page that formats grow by a line has
    // no room below it: it grows upward instead, a line at a time, its last
    // line kept on the page (a footer set larger), as a text box anchored at
    // its bottom grows. Its first line must stay on the page too.
    if (!formats || !(error instanceof PdfNativeTransactionError) ||
      !/would run off the page/u.test(error.message) || !paragraph.pageContent) throw error;
    const bounds = kernel.visiblePageBounds(paragraph.selection.pageNumber);
    if (!bounds) throw error;
    let raised: ReturnType<typeof setParagraph> | null = null;
    for (let lines = 1; lines <= 60 && !raised; lines += 1) {
      const layout = Object.freeze({ ...paragraph.layout,
        firstBaseline: paragraph.layout.firstBaseline + lines * paragraph.layout.lineStep });
      const [b0, b1, b2, b3] = paragraph.basis;
      const [x, y] = applyMatrix(paragraph.blockMatrix, layout.firstLeft * b0 + (layout.firstBaseline +
        paragraph.layout.lineStep) * b2, layout.firstLeft * b1 + (layout.firstBaseline + paragraph.layout.lineStep) * b3);
      if (x < bounds.left - 0.5 || x > bounds.right + 0.5 || y < bounds.bottom - 0.5 || y > bounds.top + 0.5) break;
      try {
        const lifted = Object.freeze({ ...paragraph, layout });
        raised = setParagraph(kernel, lifted, replacementText, installed, formats);
        paragraph = lifted;
      } catch (liftError) {
        if (!(liftError instanceof PdfNativeTransactionError) || !/would run off the page/u.test(liftError.message)) {
          throw liftError;
        }
      }
    }
    if (!raised) throw error;
    set = raised;
  }
  // A frame that leaves a left-set paragraph breaking where it did, its text
  // unchanged, changes nothing on the page: there is nothing to write (a
  // paragraph set right, centred or justified follows its new edge).
  if (frame?.right !== undefined && paragraph.sourceLineStep === undefined && paragraph.sizeScale === undefined &&
    replacementText === selection.item.str && set.sameBreaks && paragraph.layout.alignment === "left" && !formats) {
    refuse(MRAI_NATIVE_PARAGRAPH_FRAME_UNCHANGED_MESSAGE);
  }
  // The block records the layout its lines were set in.
  if (set.layout !== paragraph.layout) paragraph = Object.freeze({ ...paragraph, layout: set.layout });
  // A substitute only answers characters no font of the paragraph has; a
  // rewrite that set none of them in it names no substitute.
  const used = installed !== null && set.fontsUsed.has(installed.resourceName);
  const { request, evidence } = await paragraphEvidence(paragraph, replacementText, set, frame?.right !== undefined);
  const sourceAppearance: PdfNativeSourceTextAppearanceRun[] =
    paragraph.plans.flatMap((plan) => plan?.sourceAppearance ?? []);
  return Object.freeze({
    request,
    evidence,
    properties: null,
    ...(set.notices.length ? { notices: set.notices } : {}),
    ...(sourceAppearance.length ? { sourceAppearance: Object.freeze(sourceAppearance) } : {}),
    ...(used && installed ? { substitute: Object.freeze({ resourceName: installed.resourceName, text: missing }) } : {}),
  });
}

function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

function codeBytesEqual(left: readonly number[], right: readonly number[]): boolean {
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

/**
 * Undo and redo of a rewrite: an exact byte restore. Undo writes the source
 * span back over the rewritten one, redo writes the rewritten span back over
 * the source; nothing is set again. The current bytes are rebuilt from the
 * origin's own evidence and must stand where the origin wrote them.
 */
export async function planMraiNativeParagraphRewriteHistory(
  kernel: MraiNativeParagraphRewriteKernel,
  input: PdfNativeListTextHistoryReplacementInput,
  prior: PdfNativeListTextFidelityEvidence,
  sourceSha256: string,
): Promise<PdfNativeListTextHistoryReplacementPlan> {
  const rewrite = prior.paragraphRewrite, id = input.positionedOriginTransactionId;
  if (!rewrite || !id || id.length > 200 || rewrite.history || !rewrite.block || !prior.selection) {
    refuse("Paragraph rewrite history origin is unavailable.");
  }
  const undo = codeBytesEqual(input.expectedCurrentCodeBytes, prior.replacementCodeBytes) &&
    codeBytesEqual(input.desiredCodeBytes, prior.originalCodeBytes);
  const redo = codeBytesEqual(input.expectedCurrentCodeBytes, prior.originalCodeBytes) &&
    codeBytesEqual(input.desiredCodeBytes, prior.replacementCodeBytes);
  if (undo === redo) refuse("Paragraph rewrite history has no exact opposite side.");
  const start = prior.textOperatorSpan.start;
  const original = Uint8Array.from(rewrite.originalBytes);
  const stream = await kernel.inspectStream(rewrite.target);
  // The rewrite took the rules its runs were drawn with away, or kept them
  // all (readParagraph): found again by their place beside its shows.
  const originalText = LATIN1.decode(original);
  const rewrittenWith = (rules: readonly Readonly<{ start: number; end: number }>[]) => rewrittenSpan(originalText,
    start, rewrite.shows, rewrite.pens, Uint8Array.from(rewrite.block!.bytes), rules);
  const sidesOf = (rewritten: Uint8Array) => undo ? [rewritten, original] as const : [original, rewritten] as const;
  const standsAs = (bytes: Uint8Array) => bytesEqual(stream.streamBytes.slice(start, start + bytes.length), bytes);
  let [current, restore] = sidesOf(rewrittenWith([]));
  if (!standsAs(current)) {
    // This writer's rules beside its shows, other writers' rules among them
    // (paragraphLinkRules), or both.
    const ruled = paragraphRuleSpans(original, originalText,
      rewrite.shows.map((show) => ({ start: show.start - start })))
      .map((rule) => Object.freeze({ start: rule.start + start, end: rule.end + start }));
    const linked = paragraphLinkRuleSpans(originalText).filter((rule) =>
      !ruled.some((own) => rule.start + start >= own.start && rule.end + start <= own.end))
      .map((rule) => Object.freeze({ start: rule.start + start, end: rule.end + start }));
    for (const rules of [ruled, linked, [...ruled, ...linked].sort((left, right) => left.start - right.start)]) {
      if (!rules.length) continue;
      [current, restore] = sidesOf(rewrittenWith(rules));
      if (standsAs(current)) break;
    }
  }
  if (!standsAs(current)) {
    refuse("Paragraph rewrite history source drifted from the origin receipt.");
  }
  // The objects the origin pinned stand as they stand now (the origin's
  // replay proves its fonts did not change).
  const pins: PdfNativeObjectPrecondition[] = [];
  for (const pin of rewrite.pins) {
    if (pin.objectNumber === rewrite.target.objectNumber && pin.generation === rewrite.target.generation) continue;
    pins.push(Object.freeze({ ...(await kernel.pinObject(pin)), role: pin.role }));
  }
  const direction = undo ? "undo" as const : "redo" as const;
  const currentText = undo ? prior.replacementDisplayText : prior.originalDisplayText;
  const selection: PdfNativeListTextSelectionEvidence = Object.freeze({
    sourceSha256, selectionRevision: input.selectionRevision, pageNumber: prior.selection.pageNumber,
    startOffset: 0, endOffset: currentText.length,
    item: Object.freeze({ ...prior.selection.item, str: currentText }),
  });
  const sha256 = await sha256PdfBytes(current);
  const evidence: PdfNativeListTextFidelityEvidence = Object.freeze({
    ...prior,
    selection,
    originalText: undo ? prior.replacementText : prior.originalText,
    originalDisplayText: currentText,
    originalCodeBytes: undo ? prior.replacementCodeBytes : prior.originalCodeBytes,
    originalAdvance: undo ? prior.replacementAdvance : prior.originalAdvance,
    replacementText: undo ? prior.originalText : prior.replacementText,
    replacementDisplayText: undo ? prior.originalDisplayText : prior.replacementDisplayText,
    replacementCodeBytes: undo ? prior.originalCodeBytes : prior.replacementCodeBytes,
    replacementAdvance: undo ? prior.originalAdvance : prior.replacementAdvance,
    textOperatorSpan: Object.freeze({ start, end: start + current.length, sha256 }),
    paragraphRewrite: Object.freeze({ ...rewrite, history: Object.freeze({ originTransactionId: id, direction }) }),
  });
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: stream.identity,
    expectedObjectSha256: stream.objectSha256,
    expectedStreamSha256: stream.streamSha256,
    operatorSpan: Object.freeze({ start, end: start + current.length, expectedSha256: sha256 }),
    replacement: restore,
    untouchedObjects: Object.freeze(pins),
  });
  return Object.freeze({ selection, plan: Object.freeze({ request, evidence, properties: null }) });
}





