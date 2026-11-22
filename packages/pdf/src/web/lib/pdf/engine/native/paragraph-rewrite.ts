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
import { pdfNativeFaceNameStyle, pdfNativeFontFamilyKey } from "../../native-font-identity";
import { LATIN1_EXACT } from "./latin1";
import { MRAI_NATIVE_BOX_MARK_CLOSE, MRAI_NATIVE_BOX_MARK_OPEN } from "./box-mark";
import { MRAI_NATIVE_BOLD_FACE_KEPT_REFUSAL, MRAI_NATIVE_GLYPHS_READ_AS_BLANKS_MESSAGE, MRAI_NATIVE_PARAGRAPH_READ_BACK_REFUSAL } from "./vocabulary";
import { readPdfSimpleTrueTypeGlyphExtents } from "./cid-glyph-map";
import { planPdfRegularFaceWords, type PdfRegularFaceGlyph } from "./regular-face-words";
import { summarizePdfNativeSelectionItems } from "./list-text-model";
import { MRAI_NATIVE_PARAGRAPH_REWRITE_SCHEMA_VERSION, readMraiNativeParagraphRewriteEvidence } from "./paragraph-rewrite-evidence";
import { MRAI_NATIVE_REORDERED_TEXT_MESSAGE, pdfTextHasRightToLeftRun } from "./reader-order";
import { PDF_SOURCE_NAME } from "./pdf-source-name";
import { PDF_NATIVE_TYPE3_BOLD_PASS_EM, pdfNativeSyntheticBoldStrokeWidth } from "../../native-synthetic-style";
import {
  MRAI_NATIVE_STANDING_STATE_UNKNOWN_MESSAGE,
  applyStandingEdits,
  composeStandingCutEdits,
  mraiNativeStandingBytesAreStateOrRules,
  mraiNativeStandingBytesAreANumber,
  mraiNativeStandingBytesCutAShow,
  mraiNativeStandingBytesOverstrikeAShow,
  mraiNativeStandingOverstrikeBytes,
  planStandingLeanEdits,
  normalizeStandingEdits,
  planStandingStateEdits,
  standingEditPieces,
  standingShowCuts,
  standingStateTraceDiffers,
  standingStrokeSpelling,
  walkStandingStream,
  type StandingEdit,
  type StandingTarget,
} from "./standing-state-edit";
import { pdfNativeStreamHoldsTextRule, pdfNativeTextRuleBlock, pdfNativeTextRuleBlockEnds, pdfNativeTextRuleBlocksBeside, pdfNativeTextRuleStandingAt } from "./text-rule-block";
import { pdfNativeTextRuleRangeIsWhole } from "../../native-text-rule-range";
export { MRAI_NATIVE_PARAGRAPH_REWRITE_SCHEMA_VERSION, readMraiNativeParagraphRewriteEvidence };
import {
  MAX_MRAI_NATIVE_PAGE_FONT_REFERENCES,
  PdfNativeTransactionError,
  pdfNativeActualTextClusterSpans,
  sha256PdfBytes,
  type PdfNativeContentStreamTransactionRequest,
  type PdfNativeObjectPrecondition,
  type PdfNativePageFontReferencePlan,
} from "./content-stream-transaction";
import {
  enclosingTextObject,
  mraiNativeEmptyChunkForm,
  mraiNativeEmptyShowPenForm,
  mraiNativeEmptyShowsForm,
  mraiNativeContentTokens,
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
  /** A format written into the stream by state alone, where the text stands
   * (planFormatByStateWhereTextStands): `shows` are the paragraph's shows,
   * none of them emptied or drawn again -- no pen, no block, no bytes of the
   * source carried. The record names what it writes as pieces of its span:
   * the span itself, whole and in order, and between its ranges only state
   * operators and rules (standing-state-edit.ts). */
  readonly inPlace?: true;
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

/** The refusal of a format the rewrite cannot set with a glyph the source
 * draws over the one before it (an accent over its letter) still over it. */
export const MRAI_NATIVE_PARAGRAPH_STACKED_GLYPH_REFUSAL =
  "A glyph of this text is drawn over the one before it (an accent over its letter), and the format cannot be set with it kept there.";






export const MRAI_NATIVE_PARAGRAPH_FORMAT_IN_PLACE_REFUSAL =
  "This format changes how the text is drawn, not where it stands, and the paragraph cannot be drawn again where it stands.";

/** The refusal of a face asked of the whole of a text for its style (the
 * italic, the bold of the text's own family) that would leave a character of
 * it drawn without that style (refuseStyleFaceThatLeavesTextUnstyled). */
export const MRAI_NATIVE_PARAGRAPH_STYLE_FACE_REFUSAL =
  "The face asked for its style cannot give it to all of this text, and a style is not written on a part of what it was asked of.";

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
  /** The codes whose glyph name no name list reads (read as the code itself):
   * what the reader reads for them is not known. A format written by state
   * counts them as glyphs no character is named for. */
  readonly nameUnreadCodes?: ReadonlySet<number>;
  /** A Type3 face every painted glyph of which is a picture that says its
   * own colours (`d0`: a colour emoji a browser prints as an image). Such a
   * glyph has no weight to take: a weight asked of the text around it is
   * written on that text, and the picture is drawn again as it stands. */
  readonly type3Pictures?: true;
  /** A Type3 face every painted glyph of which is a shape filled in the
   * text's colour (`d1`: the paths it fills, a stencil mask) -- a bitmap
   * face's letters, a plotting library's outlines. Nothing a stroked weight
   * sets reaches what it draws; its weight is a second pass of the show
   * (planFormatByStateWhereTextStands). Absent, the weight of a Type3 show
   * that is no picture is refused as it was. */
  readonly type3Shapes?: true;
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
  /** Its face is named for the weight or the slant it carries -- one press
   * of 加粗 or 倾斜 on the box -- not chosen for itself: the style is all of
   * the text's or is not written through that face
   * (refuseStyleFaceThatLeavesTextUnstyled). */
  readonly faceForStyle?: true;
  /** A bold face named so does not hold the character, which keeps its own
   * face: bold only if that face is, or is stroked. */
  readonly weightLeftToOwnFace?: true;
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

/** The refusal of an edit of a Type3 show whose weight is the show drawn
 * again after it (planFormatByStateWhereTextStands) that those passes do not
 * go with: one that would draw the show again or empty it (they would be
 * left standing where the show was), and a lean (the reader reads the
 * passes of a leaned show as lines of their own). The weight is taken off
 * first, which gives the page's own bytes back. */
export const MRAI_NATIVE_TYPE3_BOLD_STANDS_MESSAGE =
  "This Type3 text is bold by a second drawing of its glyphs, which this change cannot be made with; take the bold off first.";

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
export type MraiNativeParagraphRewriteFormatsResolver = ((context: Readonly<{
  request: PdfNativeListTextReplacementPlan["request"];
  /** The proof of the paragraph's first item (its font's page). */
  evidence: PdfNativeListTextReplacementPlan["evidence"];
  pageContent: boolean;
  text: string;
}>) => Promise<MraiNativeParagraphRewriteFormats>) & Readonly<{
  


  stateOnly?: boolean;
  


  leanReadByReader?: boolean;
  








  wordsReadByReader?: boolean;
}>;

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
  /** The pen move after its last glyph (a `TJ`'s last number), as `leadingKerningAdvance` is the one before its first. */
  readonly trailingKerningAdvance?: number;
  /** How much taller than its size the reader reports the show's item (a
   * Type3 face the reader measures by its glyph box); 1 where absent. */
  readonly readerVerticalScale?: number;
  /** A weight drew the show again right after it -- a Type3 show's passes
   * (standing-state-edit.ts mraiNativeStandingOverstrike), which `readShows`
   * reads as the show's weight and lists as no shows: where the passes end
   * in the stream, and how far along the line the first stands (text space,
   * before the horizontal scale). */
  readonly overstruck?: Readonly<{ end: number; offset: number }>;
  /** The show stands under an /ActualText: the reader shows that text in
   * place of its glyphs (naming a cluster of characters, or all of it). */
  readonly underActualText?: true;
  readonly underActualTextCluster?: true;
  /** The clip box the show is drawn under, when one was read. */
  readonly clipBox?: readonly [number, number, number, number] | null;
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
  /** Whether the clip a show is drawn under cuts through its ink (the item
   * proof's own check, for a show no item proof reached). */
  showCutByClip(show: MraiNativeParagraphRewriteShow, font: MraiNativeParagraphRewriteFont): boolean;
  /** The current bytes of one content stream, for a history restore. */
  inspectStream(target: Readonly<{ objectNumber: number; generation: number }>): Promise<Readonly<{
    identity: PdfNativeContentStreamTransactionRequest["identity"];
    objectSha256: string;
    streamSha256: string;
    streamBytes: Uint8Array;
  }>>;
  /** One object as it stands now, as a precondition. */
  pinObject(target: Readonly<{ objectNumber: number; generation: number }>): Promise<PdfNativeObjectPrecondition>;
  /** The simple TrueType fonts of the selection's page among the resource
   * names given, each with its base font name and its embedded program:
   * what a bold face is set again in the regular face of its family by
   * (planBoldFaceSetRegularWhereWordsStand). A name the page holds no such
   * font with a program for is left out. Read-only. */
  trueTypeFaces?(
    selection: PdfNativeListTextSelectionEvidence,
    names: readonly string[],
  ): Promise<ReadonlyMap<string, Readonly<{ baseFont: string; program: Uint8Array }>>>;
  /** The simple TrueType fonts with a program that other pages of the file
   * name and the selection's page does not, in the order of their object
   * numbers: where a face the page lacks is looked for (a bold face set
   * again in the regular face of its family, whose subset with the glyph
   * only another page names). `read` answers the font's metrics as a font
   * of the page's is read, its program, and the objects that reading pinned;
   * null where it cannot be proven. Read-only. */
  fileTrueTypeFaces?(
    selection: PdfNativeListTextSelectionEvidence,
  ): Promise<readonly Readonly<{
    font: Readonly<{ objectNumber: number; generation: number }>;
    baseFont: string;
    read(): Promise<Readonly<{
      font: MraiNativeParagraphRewriteFont;
      program: Uint8Array;
      pins: readonly PdfNativeObjectPrecondition[];
    }> | null>;
  }>[]>;
  /** The plan that names such fonts in the page's resources, each by a name
   * the page's Font category does not write (planMraiNativePageFontReference):
   * what a write that draws in them carries. Refuses where the page's
   * resources cannot take the entries. */
  referencePageFonts?(
    selection: PdfNativeListTextSelectionEvidence,
    fonts: readonly Readonly<{ objectNumber: number; generation: number }>[],
    stream: Readonly<{ objectNumber: number; generation: number }>,
  ): Promise<PdfNativePageFontReferencePlan>;
  /** The page's own content streams as a format written where the text
   * stands reads them (readParagraphWhereItStands): no item of the selection
   * is proved, so nothing here depends on the order the reader gives its
   * characters in. Absent where the caller reads no page. */
  standingPageStreams?(
    selection: PdfNativeListTextSelectionEvidence,
  ): Promise<readonly MraiNativeParagraphRewriteStandingStream[]>;
  /** The paint state a stream of the page starts in, for a format written
   * into it by state (standing-state-edit.ts): the page's streams drawn
   * before it, and the line width each graphics state dictionary the page
   * names sets (null: none; a name left out: not stated plainly). Null
   * where the stream is not one of the page's own. */
  standingStreamContext?(
    selection: PdfNativeListTextSelectionEvidence,
    plan: PdfNativeListTextReplacementPlan,
  ): Promise<Readonly<{
    prefix: readonly string[];
    lineWidths(names: readonly string[]): Promise<ReadonlyMap<string, string | null>>;
  }> | null>;
}

/**
 * One content stream of the selection's page, for a format written where the
 * text stands: a plan that names the stream (what `readStream` and
 * `readShows` take), every font of the page's resources that could be
 * proven, by its resource name, and the plan the write is carried by once
 * the shows it draws again are known -- it pins the fonts they are drawn in
 * and describes the paragraph by the first of them.
 */
export type MraiNativeParagraphRewriteStandingStream = Readonly<{
  plan: PdfNativeListTextReplacementPlan;
  fonts: ReadonlyMap<string, MraiNativeParagraphRewriteFont>;
  firstPlan(
    fontNames: readonly string[],
    first: Readonly<{ fontName: string; fontSize: number }>,
  ): Promise<PdfNativeListTextReplacementPlan>;
}>;

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











const CLEAR_OF_TEXT_BESIDE_EM = 1.6;

/** How near other text beside it a line may come where the alternative is a
 * line of its own drawn over other text: the reader's bound for reading on
 * in one item (PDFJS_SPACE_IN_FLOW_MAX_EM). The two may then read as one
 * line; neither is drawn over. */
const NEAR_TEXT_BESIDE_EM = 0.6;

/** How far past its own right edge a paragraph of several lines may set a
 * line, in ems: a hair (digits typed wider than the ones they replace). */
const PARAGRAPH_EDGE_HAIR_EM = 0.35;
/** A line that fills its paragraph's measure: at least this share of the
 * widest line's width (list-text-fidelity's lineRunsPastItsParagraph reads
 * the same). */
const PARAGRAPH_FLOWING_LINE_SHARE = 0.85;
/** Two rows on one axis: how far their middles may stand apart (a mark
 * raised after the first row's last word counts into its width), and how
 * far in from the first the second starts at least, in ems. */
const PARAGRAPH_ROWS_AXIS_EM = 0.25;
const PARAGRAPH_ROWS_MIN_INSET_EM = 0.1;
/** The most a neighbouring show may already reach into the paragraph's own
 * ink and still count as abutting it, in points: two decimals of Tm rounding
 * on either side, well under a glyph (mirror of the replay's
 * _SOURCE_ABUTMENT_MAX). */
const SOURCE_ABUTMENT_MAX = 0.1;
/** Alignment is read off the lines' ends and starts to within a point. */
const ALIGNMENT_TOLERANCE = 1;
/** How far a line's ink reaches under and over its baseline, in ems, when
 * text set off a line is asked whether it stands in that line's band
 * (list-text-fidelity's STACKED_INK_*). A line above or below, at any
 * leading a reader reads, stays clear of it. */
const STACKED_INK_UNDER_EM = 0.2;
const STACKED_INK_OVER_EM = 0.7;
/** A mark hanging after a line's last word: set at no more than this share
 * of the body's size, raised by at least this share of it. */
const HANGING_MARK_MAX_SIZE = 0.85;
const HANGING_MARK_MIN_RISE_EM = 0.15;
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
/** How far apart the baselines of a script and the word beside it stand at
 * most where the box reader reads back the space a leaned frame lost between
 * them, as a share of the larger size (native-paragraph-edit-boxes'
 * SKEWED_FRAME_SCRIPT_BASELINE_RATIO). */
const LEANED_SCRIPT_BASELINE_RATIO = 0.55;
/** A glyph the source starts this share of its size or more back from where
 * the glyph before it ended is set over that glyph, not after it: TeX's accent
 * over a letter (`[(F)100(\023)400(\020sic)] TJ` -- the acute, the pen taken
 * back across it, the dotless i under it), an overstrike. The reader starts a
 * new item there (PDFJS_NEGATIVE_SPACE_FACTOR) and reads no space. */
const STACKED_GLYPH_BACK_EM = 0.2;

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
  /** It draws glyphs and no character of the text: every one of them is a
   * glyph its font only calls white space (MraiNativeParagraphRewriteFont
   * .inkedBlankCodes: the marks an Urdu face stands over its letters), which
   * the reader reads as a blank. Owned by where it stands and among which
   * shows it is drawn, by the one read that owns such a show
   * (readParagraphWhereItStands). */
  readonly readsBlank?: true;
  /** How far above its line's baseline a rule under it runs, where that is
   * not its `rise`: a show the page stands off the baseline by where it
   * places it (the letters of a Nastaleeq word step up along the line) is
   * ruled on the line, with the shows beside it (readParagraphWhereItStands). */
  readonly ruleRise?: number;
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
  /** Each source line's baseline across the lines, as the page sets it (a
   * line under a fraction stands further down than the step): where a line
   * the edit leaves alone is drawn again. */
  readonly lineBaselines: readonly number[];
  /** The frame's right edge grown as near the text beside the paragraph as
   * NEAR_TEXT_BESIDE_EM allows, where that is further than the frame's own:
   * taken only instead of a new line over other text. */
  readonly nearFrameRight?: number;
  /** The paragraph's own line step, when `layout` is set at one the reader
   * asked for (paragraphAtRequestedLineStep). */
  readonly sourceLineStep?: number;
  /** Read where it stands with glyphs its fonts name no character for
   * (readParagraphWhereItStands, `unnamedGlyphs`): on a line that holds one,
   * how many characters a show draws is not known, and `characterShows` says
   * of such a line's characters only which line they are on (all of them are
   * dealt to its first show). Every show of the paragraph is asked the same
   * look, and its rules are read show by show. */
  readonly unreadGlyphs?: true;
  /** Read by where the reader's items start (ownShowsByPins): the shows of
   * glyphs read as blanks drawn at the edge of its shows that it does not
   * hold (`start:end` each). Of all such glyphs beside its lines, these
   * alone are not known to be other text's: each is drawn between a show of
   * the box and one that is not (refuseGlyphsReadAsBlanksBesideLines). */
  readonly edgeBlankShows?: ReadonlySet<string>;
  /** Written around the block where it is set inside the text object of the
   * paragraph's last show, which is ended before it and begun again after it
   * (standingBlockSplit). */
  readonly blockWrap?: Readonly<{ before: string; after: string }>;
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
  /** Why this tier's measure cannot set the paragraph anew, where the read
   * went on without it (readParagraph's `standing`): a format that moves no
   * glyph draws every show again by the array the page drew it with, and
   * needs none of the measure; anything that sets a line refuses with this. */
  readonly measureRefusal?: string;
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
  /** By character of the paragraph's text: a glyph the source sets back over
   * the glyph before it on its line (STACKED_GLYPH_BACK_EM), in shares of its
   * show's size along the line -- how far it starts from that glyph's end
   * (`back`, negative), how wide the source draws that glyph and this one,
   * and how far that glyph starts from the end of the glyph before it where
   * the text holds the three side by side (`lead`: TeX's kern before an
   * accent, which sets the letter where it would stand with none). */
  readonly stacked: ReadonlyMap<number, Readonly<{ back: number; beforeWidth: number; width: number; lead: number | null }>>;
  /** A show of the paragraph sets a glyph back so, and which characters of
   * the text its glyphs are is not known. */
  readonly stackedUnread: boolean;
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
    if (!pieces.length || flat(pieces.map((piece) => piece.item.str).join("")) !== flat(item.str)) {
      return refused(index, message);
    }
    // One show alone: the item is that show and the reader's gap space on
    // both sides of it -- the viewer folded the space before it in, and the
    // reader wrote the space after it before it ended the item (a show of
    // another frame comes next: in a leaned line, where TeX sets a formula
    // at another width beside the words; a script off the baseline). No
    // show reads as an item with a gap at both ends; the show does, proved
    // alone, and the gaps are laid out as any gap between two shows. An item
    // that is the show's own text was refused for what the show is.
    if (pieces.length === 1 && !(/^\s/u.test(item.str) && /\s$/u.test(item.str) &&
      !/^\s|\s$/u.test(pieces[0].item.str))) {
      return refused(index, message);
    }
    // (Its words as the reader read them, the spaces it read into the
    // show's own gaps among them -- TeX sets a line's words in one show, a
    // number of the array for each space: the show's codes alone spell no
    // space, and the paragraph's text holds the reader's.)
    if (pieces.length === 1) pieces[0] = { at: pieces[0].at, item: Object.freeze({ ...pieces[0].item, str: item.str.trim() }) };
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
 * The rules drawn under (through) a paragraph's shows, read as its shows'
 * own formats (readParagraph's; its comments are on the lines below): this
 * writer's rules beside the shows' text objects, marked on the shows they
 * rule whole (`shows` is updated in place) and returned by character where
 * one rules part of a show, and the rules another writer drew under them.
 * `rules` are the spans taken away with the text. `characterShows` may be
 * empty: no rule under part of a show is then read as the paragraph's.
 */
function paragraphOwnRules(input: Readonly<{
  streamBytes: Uint8Array;
  source: string;
  shows: OwnedShow[];
  basis: readonly [number, number, number, number];
  blockMatrix: Matrix;
  lineBaselines: readonly number[];
  paragraphCharacters: readonly string[];
  characterShows: readonly number[];
}>): { rules: Readonly<{ start: number; end: number }>[];
  ruledCharacters: Map<number, { underline?: true; strikethrough?: true }> } {
  const { streamBytes, source, shows, basis, blockMatrix, lineBaselines, paragraphCharacters, characterShows } = input;
  const beside = paragraphRuleSpansOnItsLines(streamBytes, source, shows, basis, lineBaselines);
  const matched = new Map<number, { underline?: true; strikethrough?: true }>();
  const ruledCharacters = new Map<number, { underline?: true; strikethrough?: true }>();
  const ruledAt = new Set<number>();
  // Where a rule under a show runs, across the lines.
  const ruleAcross = (owned: OwnedShow) =>
    owned.ruleRise !== undefined ? lineBaselines[owned.line] + owned.ruleRise : owned.v + owned.rise;
  for (const [index, owned] of shows.entries()) {
    if (!beside.length || !owned.inked || !(owned.inked.end > owned.inked.start) ||
      !owned.show.graphicsMatrix.every((value, at) => Math.abs(value - blockMatrix[at]) <= 1e-9)) continue;
    const object = enclosingTextObject(source, owned.start);
    const across = ruleAcross(owned);
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
  // A rule under consecutive runs of one line is one rule (setParagraph draws
  // it from where the first is inked to where the last is): the rule of
  // every run it runs under, whole. Read run by run it was no run's, the
  // paragraph then owned none of its rules (all or none), and an underline
  // set on a line of several runs -- a line drawn again show by show where it
  // stands, a title's words in two faces -- could not be taken off again,
  // nor go with its words.
  if (beside.length && !beside.every((rule) => ruledAt.has(rule.start))) {
    const determinant = basis[0] * basis[3] - basis[1] * basis[2];
    const along = (point: readonly [number, number]) => (point[0] * basis[3] - point[1] * basis[2]) / determinant;
    const ruled = shows.flatMap((owned, index) => owned.inked && owned.inked.end > owned.inked.start &&
      owned.show.graphicsMatrix.every((value, at) => Math.abs(value - blockMatrix[at]) <= 1e-9) ? [{ index, owned, inked: owned.inked }] : []);
    for (const rule of beside) {
      if (ruledAt.has(rule.start) || !(Math.abs(determinant) > 1e-12)) continue;
      const ends = pdfNativeTextRuleBlockEnds(streamBytes.slice(rule.start, rule.end));
      if (!ends) continue;
      const from = along(ends.from), to = along(ends.to);
      const near = (left: number, right: number) => Math.abs(left - right) <= 1e-3 + 1e-4 * Math.abs(to - from);
      for (const first of ruled) {
        if (ruledAt.has(rule.start)) break;
        if (!near(first.inked.start, from)) continue;
        // The runs of its line the rule runs under, from this one on.
        const under = ruled.filter((entry) => entry.owned.line === first.owned.line &&
          entry.inked.start >= from - 1e-3 && entry.inked.end <= to + 1e-3 + 1e-4 * Math.abs(to - from));
        if (under.length < 2 || !under.some((entry) => near(entry.inked.end, to)) ||
          under.some((entry) => entry.owned.size !== first.owned.size ||
            Math.abs(ruleAcross(entry.owned) - ruleAcross(first.owned)) > 1e-6 ||
            JSON.stringify(entry.owned.color) !== JSON.stringify(first.owned.color) ||
            enclosingTextObject(source, entry.owned.start).start !== enclosingTextObject(source, first.owned.start).start)) continue;
        const object = enclosingTextObject(source, first.owned.start);
        const across = ruleAcross(first.owned);
        const geometry = { textMatrix: [basis[0], basis[1], basis[2], basis[3],
          from * basis[0] + across * basis[2], from * basis[1] + across * basis[3]] as
          [number, number, number, number, number, number],
        fontSize: first.owned.size, advance: to - from, color: first.owned.color };
        for (const style of ["underline", "strikethrough"] as const) {
          let standing: ReturnType<typeof pdfNativeTextRuleStandingAt> = null;
          try {
            standing = pdfNativeTextRuleStandingAt(streamBytes, style === "underline" ? object.start : object.end, style, geometry);
          } catch { standing = null; }
          if (!standing || standing.start !== rule.start || !pdfNativeTextRuleRangeIsWhole(standing.range)) continue;
          ruledAt.add(rule.start);
          for (const entry of under) matched.set(entry.index, { ...matched.get(entry.index), [style]: true });
          break;
        }
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
  return { rules, ruledCharacters };
}

/**
 * The paragraph a selection names, read and owned: every reader item proved
 * alone, the shows those proofs name, where each is on the page and how it is
 * set, and the paragraph's original layout.
 */
async function readParagraph(
  kernel: MraiNativeParagraphRewriteKernel,
  named: PdfNativeListTextSelectionEvidence,
  /** The edit keeps the text and asks formats only (planMraiNativeParagraphRewrite):
   * what only the measure needs is not refused here but handed on as
   * `measureRefusal`, for setParagraph to refuse with once the formats turn
   * out to set a line after all. Ownership is read as strictly either way. */
  standing = false,
): Promise<ReadParagraph> {
  let measureRefusal: string | undefined;
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
    // a slash beside an "=". (Mirror of the replay's owned shows.) Drawn again
    // where it stands, by its own array, the mark stays over what it is drawn
    
    // an underline and a bold of a TeX paragraph holding one "≠").
    for (const [index, glyph] of glyphs.entries()) {
      if (glyph.text === undefined || !PDFJS_ZERO_WIDTH_DIACRITIC.test(glyph.text)) continue;
      if (partial && !(index * font.codeUnitBytes >= partial.codeStart && index * font.codeUnitBytes < partial.codeEnd)) continue;
      if (index === 0 || Math.abs(glyph.start - glyphs[index - 1].end) > 1e-9 * show.fontSize) {
        const message = "A mark of this paragraph is drawn off the glyph the rewrite would set it after.";
        if (!standing) refuse(message);
        measureRefusal ??= message;
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
  // Each line without the mark set raised and small after its last word (a
  // footnote star or number after a name): the mark hangs past a centred
  // line, which is centred without it.
  const unmarkedExtents = lineExtents.map((extent, line) => {
    const inked = shows.filter((show) => show.line === line && show.inked);
    const last = inked.reduce((best, show) => show.inked!.end > best.inked!.end ? show : best);
    const hangs = inked.length > 1 && last.size < bodySize * HANGING_MARK_MAX_SIZE &&
      last.v + last.rise - lineBaselines[line] > bodySize * HANGING_MARK_MIN_RISE_EM &&
      last.inked!.end - last.inked!.start < bodySize;
    if (!hangs) return extent;
    const body = inked.filter((show) => show !== last);
    return { left: Math.min(...body.map((show) => show.inked!.start)), right: Math.max(...body.map((show) => show.inked!.end)) };
  });
  const alignment = paragraphAlignment(lineExtents, frameRight, restLeft, unmarkedExtents);
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
  const { rules, ruledCharacters } = paragraphOwnRules({ streamBytes, source, shows, basis, blockMatrix, lineBaselines,
    paragraphCharacters, characterShows });
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
  // What the lines leave between their words is measured by setting the
  // words again in their fonts: a word the page draws with a glyph its font
  // gives no single reading (a subset's ligature, its letters nowhere else in
  // it) cannot be set, and neither can a line of it -- but it can be drawn
  
  // paragraph was refused for a width no lean needs).
  let gapRatio: number | null = null;
  try {
    gapRatio = paragraphGapRatio(kernel, selection.item.str, characterShows, shows, styleOf, lineExtents, lineSizes,
      blankRuns);
  } catch (error) {
    if (!standing || !(error instanceof PdfNativeTransactionError)) throw error;
    measureRefusal ??= error.message;
  }
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
  // The glyphs the source sets back over the one before them (an accent over
  // its letter). This tier sets a line's glyphs side by side: one it set so
  // came off the letter it was drawn over, "F´ısica" for "Física" (an arXiv
  // author line made upright, pdf-walk-1000 0686). Each is known by the
  // character of the text it draws, a ligature's glyph by its letters.
  const stacked = new Map<number, Readonly<{ back: number; beforeWidth: number; width: number; lead: number | null }>>();
  let stackedUnread = false;
  {
    const all = characters(selection.item.str);
    const glyphOf = new Map<number, Readonly<{ ink: readonly [number, number]; first: boolean; last: boolean }>>();
    const emOf = (own: OwnedShow) => own.size * (own.horizontalScale / 100);
    for (const [owner, own] of shows.entries()) {
      const mine = all.flatMap((character, index) => characterShows[index] === owner && !isBlank(character) ? [index] : []);
      let counts: readonly number[] | null = mine.length === own.ink.length ? own.ink.map(() => 1) : null;
      if (!counts) {
        try {
          const unit = own.font.codeUnitBytes;
          const read = showGlyphs(kernel, own.show, own.font).filter((glyph, index) => !/^\s*$/u.test(glyph.text ?? "x") &&
            (!own.partial || (index * unit >= own.partial.codeStart && index * unit < own.partial.codeEnd)))
            .map((glyph) => characters(kernel.normalizeUnicode(glyph.text ?? "x")).filter((character) => !isReaderBlank(character)).length);
          if (read.length === own.ink.length && read.every((count) => count > 0) &&
            read.reduce((sum, count) => sum + count, 0) === mine.length) counts = read;
        } catch (error) {
          if (!(error instanceof PdfNativeTransactionError)) throw error;
        }
      }
      if (!counts) {
        const em = emOf(own);
        if (em > 0 && own.ink.some((ink, at) => at > 0 && ink[0] - own.ink[at - 1][1] <= -STACKED_GLYPH_BACK_EM * em + 1e-9)) {
          stackedUnread = true;
        }
        continue;
      }
      let at = 0;
      for (const [glyph, count] of counts.entries()) {
        for (let part = 0; part < count; part += 1) {
          glyphOf.set(mine[at + part], Object.freeze({ ink: own.ink[glyph], first: part === 0, last: part === count - 1 }));
        }
        at += count;
      }
    }
    for (let index = 1; index < all.length; index += 1) {
      const glyph = glyphOf.get(index), before = glyphOf.get(index - 1);
      if (!glyph?.first || !before?.last) continue;
      const own = shows[characterShows[index]];
      const em = emOf(own);
      if (!(em > 0) || shows[characterShows[index - 1]].line !== own.line) continue;
      const back = glyph.ink[0] - before.ink[1];
      if (back > -STACKED_GLYPH_BACK_EM * em + 1e-9) continue;
      const led = index > 1 ? glyphOf.get(index - 2) : undefined;
      stacked.set(index, Object.freeze({ back: back / em, beforeWidth: (before.ink[1] - before.ink[0]) / em,
        width: (glyph.ink[1] - glyph.ink[0]) / em,
        lead: led?.last && before.first && shows[characterShows[index - 2]].line === own.line
          ? (before.ink[0] - led.ink[1]) / em : null }));
    }
  }
  // A paragraph set flush left may grow to the right into the free space
  // beside it: up to the nearest other text of the stream that starts to the
  // right of its lines, on the band of any of them, or else the visible
  // page's edge. A form's frame is its own (mirror of the replay).
  let freeRight: number | null = null;
  let nearRight: number | null = null;
  const bounds = kernel.visiblePageBounds(selection.pageNumber);
  if ((alignment === "left" || alignment === "justify") && pageContent && bounds) {
    freeRight = Math.max(...[[bounds.left, bounds.bottom], [bounds.right, bounds.bottom],
      [bounds.left, bounds.top], [bounds.right, bounds.top]].map(([x, y]) => {
      const [bx, by] = applyMatrix(toBlock, x, y);
      return bx * basis[0] + by * basis[1];
    }));
    nearRight = freeRight;
    const others = otherInkedExtents(kernel, { allShows, shows, twinKeys, fonts, basis, blockMatrix });
    for (const [line, baseline] of lineBaselines.entries()) {
      const bottom = baseline - PDFJS_VERTICAL_SHIFT_RATIO * lineSizes[line];
      const top = baseline + PDFJS_VERTICAL_SHIFT_RATIO * lineSizes[line];
      for (const other of others) {
        if (other.across >= bottom && other.across <= top && other.start >= frameRight - LAYOUT_TOLERANCE) {
          freeRight = Math.min(freeRight, other.start - CLEAR_OF_TEXT_BESIDE_EM * lineSizes[line]);
          nearRight = Math.min(nearRight, other.start - NEAR_TEXT_BESIDE_EM * lineSizes[line]);
        } else if (other.start >= frameRight - LAYOUT_TOLERANCE &&
          other.across + STACKED_INK_OVER_EM * lineSizes[line] > baseline - STACKED_INK_UNDER_EM * lineSizes[line] &&
          other.across - STACKED_INK_UNDER_EM * lineSizes[line] < baseline + STACKED_INK_OVER_EM * lineSizes[line]) {
          // Text set off the line whose ink stands in the line's band (a
          // fraction's numerator and denominator after "softmax(", a
          // script): the line stops where it starts. Grown past it, the
          // words ran between the two, through the fraction's bar.
          freeRight = Math.min(freeRight, other.start);
          nearRight = Math.min(nearRight, other.start);
        }
      }
    }
    // (A justified paragraph keeps its frame; the room beside it is read
    // only for the frame below.)
    if (alignment !== "left" || !(freeRight > frameRight + LAYOUT_TOLERANCE)) freeRight = null;
    if (!(nearRight > frameRight + LAYOUT_TOLERANCE)) nearRight = null;
  }
  // A centred paragraph's room: what is free on both its sides, clear of
  // the text beside it as a paragraph set left keeps clear.
  let centreRoom: number | null = null;
  if (alignment === "center" && pageContent && bounds) {
    const edges = [[bounds.left, bounds.bottom], [bounds.right, bounds.bottom], [bounds.left, bounds.top],
      [bounds.right, bounds.top]].map(([x, y]) => {
      const [bx, by] = applyMatrix(toBlock, x, y);
      return bx * basis[0] + by * basis[1];
    });
    let rightEdge = Math.max(...edges), leftEdge = Math.min(...edges);
    const others = otherInkedExtents(kernel, { allShows, shows, twinKeys, fonts, basis, blockMatrix });
    for (const [line, baseline] of lineBaselines.entries()) {
      const bottom = baseline - PDFJS_VERTICAL_SHIFT_RATIO * lineSizes[line];
      const top = baseline + PDFJS_VERTICAL_SHIFT_RATIO * lineSizes[line];
      for (const other of others) {
        if (other.across < bottom || other.across > top) continue;
        if (other.start >= frameRight - LAYOUT_TOLERANCE) {
          rightEdge = Math.min(rightEdge, other.start - CLEAR_OF_TEXT_BESIDE_EM * lineSizes[line]);
        } else if (other.end !== null && other.end <= frameLeft + LAYOUT_TOLERANCE) {
          leftEdge = Math.max(leftEdge, other.end + CLEAR_OF_TEXT_BESIDE_EM * lineSizes[line]);
        } else {
          // Text across the paragraph's own lines, or of unknown extent: no room is proven.
          rightEdge = frameRight;
          leftEdge = frameLeft;
        }
      }
    }
    centreRoom = Math.max(0, Math.min(rightEdge - frameRight, frameLeft - leftEdge));
  }
  const frame = paragraphFrameHoldingItsLines(kernel, selection.item.str, characterShows, shows, styleOf, gapRatio,
    { alignment, frameLeft, frameRight, firstLeft, restLeft }, freeRight, false, centreRoom);
  // (The frame with all the room beside the paragraph, as near the text
  // there as NEAR_TEXT_BESIDE_EM: the rewrite takes it only instead of a new
  // line drawn over other text, and only for the paragraph's last line.)
  const nearFrame = nearRight !== null
    ? paragraphFrameHoldingItsLines(kernel, selection.item.str, characterShows, shows, styleOf, gapRatio,
      { alignment: "left", frameLeft, frameRight, firstLeft, restLeft }, nearRight, true) : null;
  return Object.freeze({
    selection, items, plans, firstPlan, streamBytes, source, pageContent,
    shows: Object.freeze(shows), characterShows, basis, blockMatrix,
    layout: Object.freeze({ alignment, frameLeft: frame.left, frameRight: frame.right, firstLeft, restLeft,
      ...(lineLefts ? { lineLefts } : {}),
      firstBaseline, lineStep, lineCount, lineExtents: Object.freeze(lineExtents), bodySize,
      ...(reach[0].above > 0 ? { firstReach: reach[0].above } : {}) }),
    lineBaselines: Object.freeze([...lineBaselines]),
    ...(nearFrame && nearFrame.left === frame.left && nearFrame.right > frame.right + LAYOUT_TOLERANCE
      ? { nearFrameRight: nearFrame.right } : {}),
    pens: Object.freeze(pens),
    spanStart: Math.min(shows[0].start, ...twins.map((twin) => twin.start), ...rules.map((rule) => rule.start)),
    rules: Object.freeze(rules.map((rule) => Object.freeze(rule))),
    insertion: Math.max(blockInsertion({ source, allShows, toBlock, basis, blockMatrix, bodySize,
      spanStart: Math.min(shows[0].start, ...twins.map((twin) => twin.start), ...rules.map((rule) => rule.start)),
      objectEnd: enclosingTextObject(source, Math.max(shows.at(-1)!.start, ...twins.map((twin) => twin.start))).end }),
      ...rules.map((rule) => rule.end)),
    pins: Object.freeze([...pins.values()]),
    allShows, twins: Object.freeze(twins), twinKeys, fonts, fontOrder: Object.freeze(fontOrder), preferred, gapRatio, keptGaps,
    blankRuns, characterInk, stacked, stackedUnread,
    ...(measureRefusal !== undefined ? { measureRefusal } : {}),
    characterRules: ruledCharacters,
    unreadGaps: Object.freeze(unreadGaps),
    unreadGapWidths: Object.freeze(unreadGapWidths),
  });
}
















const BLOCK_INSERTION_REACH_BYTES = 8 * 1024;

function blockInsertion(input: Readonly<{
  source: string;
  allShows: readonly MraiNativeParagraphRewriteShow[];
  toBlock: Matrix;
  basis: readonly [number, number, number, number];
  blockMatrix: Matrix;
  bodySize: number;
  spanStart: number;
  objectEnd: number;
}>): number {
  const { source, toBlock, basis, blockMatrix } = input;
  const inked = input.allShows.filter((show) => show.codes.length > 0 && !show.vertical &&
    (show.text === null || !characters(show.text).every(isReaderBlank)));
  const placed = (show: MraiNativeParagraphRewriteShow) => {
    const [x, y] = applyMatrix(toBlock, ...applyMatrix(show.graphicsMatrix, ...applyMatrix(show.textMatrix, 0, show.textRise)));
    return Object.freeze({ along: x * basis[0] + y * basis[1], across: x * basis[2] + y * basis[3] });
  };
  let end = input.objectEnd;
  for (let hop = 0; hop < 16; hop += 1) {
    const before = inked.filter((show) => show.end <= end).reduce<MraiNativeParagraphRewriteShow | null>(
      (last, show) => !last || show.start > last.start ? show : last, null);
    const after = inked.filter((show) => show.start >= end).reduce<MraiNativeParagraphRewriteShow | null>(
      (first, show) => !first || show.start < first.start ? show : first, null);
    if (!before || !after) break;
    const from = placed(before), to = placed(after);
    // The same line: within a line's height across it (a fraction's parts
    // stand above and below its baseline, its denominator starting left of
    // where its numerator did), not back at the start of a line.
    if (Math.abs(to.across - from.across) > 0.9 * input.bodySize || to.along < from.along - 3 * input.bodySize) break;
    if (!after.graphicsMatrix.every((value, at) => Math.abs(value - blockMatrix[at]) <= 1e-9)) break;
    let next: ReturnType<typeof enclosingTextObject>;
    try {
      next = enclosingTextObject(source, after.start);
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
      break;
    }
    // Between the two: nothing but graphics drawn under a state of their own.
    const between = source.slice(end, next.start);
    const operators = between.split(/[^A-Za-z0-9*'"]+/u);
    // (No string, dictionary or inline image in it: paths, their state, a dash array.)
    if (!/^(?:\s*q\b[^()<>]*?\bQ\b)*\s*$/u.test(between) || /\b(?:BI|BDC|BMC|EMC|BT|ET)\b/u.test(between) ||
      operators.filter((token) => token === "q").length !== operators.filter((token) => token === "Q").length) break;
    // (The span the record keeps for an undo grows by what it writes past.)
    if (next.end - input.objectEnd > BLOCK_INSERTION_REACH_BYTES ||
      next.end - input.spanStart > PARAGRAPH_REWRITE_MAX_REPLACEMENT_BYTES / 2) break;
    end = next.end;
  }
  return end;
}

/** The refusal of a paragraph whose own shows are more bytes than one edit
 * writes (PARAGRAPH_REWRITE_MAX_REPLACEMENT_BYTES): a format on it, and words
 * typed into it, are not written. */
export const MRAI_NATIVE_PARAGRAPH_TOO_LARGE_MESSAGE =
  "The paragraph is drawn show by show in more bytes than one edit can write again.";

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
  refuseGlyphsReadAsBlanksBesideLines(paragraph);
}

/** The second of those two: a show the paragraph does not own, made of such
 * glyphs, on a line of it. (A format written where the text stands asks this
 * alone: it sets no text again, so the glyphs its own shows draw are not
 * lost, and readParagraphWhereItStands owns the shows of them that stand
 * with its lines.) */
function refuseGlyphsReadAsBlanksBesideLines(paragraph: ReadParagraph): void {
  if (![...paragraph.fonts.values()].some((font) => font.inkedBlankCodes?.size)) return;
  const { basis, blockMatrix, layout } = paragraph;
  const toBlock = invertMatrix(blockMatrix);
  const owned = new Set([...paragraph.shows.map((show) => `${show.start}:${show.end}`), ...paragraph.twinKeys]);
  const reach = 2 * layout.bodySize;
  for (const show of paragraph.allShows) {
    // (Read by pins, each such show was told the box's or other text's from
    // the stream: only those at the edge of the box's shows are not known.)
    if (paragraph.edgeBlankShows && !paragraph.edgeBlankShows.has(`${show.start}:${show.end}`)) continue;
    if (owned.has(`${show.start}:${show.end}`) || ![0, 1, 2].includes(show.renderingMode) ||
      !drawsGlyphReadAsBlank(show, paragraph.fonts.get(show.fontName))) continue;
    const [x, y] = applyMatrix(composeMatrix(composeMatrix(show.textMatrix, show.graphicsMatrix), toBlock), 0, show.textRise);
    const along = x * basis[0] + y * basis[1], across = x * basis[2] + y * basis[3];
    // On a line of the paragraph, as far off its baseline as the reader
    // still reads a glyph on it, within two sizes of its frame.
    if (along < layout.frameLeft - reach || along > layout.frameRight + reach) continue;
    if (paragraph.shows.some((own) => !own.readsBlank &&
      Math.abs(across - own.v) <= PDFJS_VERTICAL_SHIFT_RATIO * Math.max(own.size, show.fontSize))) {
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
  /** The lines without a mark hanging after their last word (see the
   * caller): centred lines agree by these where the mark puts one off. */
  unmarked: readonly Readonly<{ left: number; right: number }>[] = extents,
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
  // An author's name with its footnote star: "Ashish Vaswani*" over "Google
  // Brain" over an address stood two points off the others' axis by the
  // star, read as set flush left, and words typed in it were set from the
  
  const bodies = unmarked.map((extent) => (extent.left + extent.right) / 2);
  if (unmarked !== extents && bodies.every((middle) => flush(middle, bodies[0])) &&
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
  /** On the line the edit starts on: where what stays of it as the page
   * draws it ends, and this tier's own setting starts (setParagraph's kept
   * lines); unset where the tier sets the line from its start. */
  readonly keptUntil?: number;
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
  /** The room beside the paragraph is taken whole, a flowing paragraph's
   * too: asked only for the frame tried instead of a new line over other
   * text (readParagraph's nearFrameRight). */
  pastTheMeasure = false,
  /** The room free on both sides of a centred paragraph (readParagraph):
   * rows of their own lengths grow into it alike on both sides. */
  centreRoom: number | null = null,
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
    // A paragraph of several lines has a right edge its lines show: the room
    // beside it is for that hair, not for words. Grown as far as the next
    // line's first word allowed, the last line of a column typed onto ran
    // 12 pt out past the column's edge where a word typed belongs on a new
    
    // room beside it as before.
    // (A paragraph whose lines fill its measure, all but the last: rows of
    // their own lengths -- an address, a list of names -- are lines a break
    // ended, and grow as a single line does.)
    const widest = Math.max(...lineWidths);
    // (Two rows show no measure: the first is as wide as itself. Two rows on
    // one axis, the second starting in from the first -- a name over its
    // affiliation -- are rows of their own; read as a paragraph that flows,
    // the affiliation typed two letters longer was kept to the name's width
    // and its letters went down a line, onto the address under it
    
    const size = Math.max(...styles.map((style) => style.size));
    const rowsOnOneAxis = lineCount === 2 && frame.restLeft - frame.firstLeft >= PARAGRAPH_ROWS_MIN_INSET_EM * size &&
      Math.abs(frame.firstLeft + lineWidths[0] / 2 - (frame.restLeft + lineWidths[1] / 2)) <= PARAGRAPH_ROWS_AXIS_EM * size;
    const flows = !pastTheMeasure && lineCount > 1 && !rowsOnOneAxis &&
      lineWidths.slice(0, -1).every((width) => width >= PARAGRAPH_FLOWING_LINE_SHARE * widest);
    let grown = flows
      ? Math.min(freeRight, right + PARAGRAPH_EDGE_HAIR_EM * size)
      : freeRight;
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
  // A centred block of rows of their own lengths (a diagram's label, a name
  // over its affiliation) grows alike on both sides into the room free
  // beside it, as a text box set centred does: its widest row typed a
  // letter longer stays one row, on the block's axis. Kept to the widest
  // row's width, the letter went down to a row of its own -- or, written in
  // place, the row grew from its left edge off the axis. Never so far that a
  // row takes up the first word of the row after it; a centred paragraph
  // whose lines fill its measure keeps the measure.
  if (frame.alignment === "center" && centreRoom !== null && centreRoom > LAYOUT_TOLERANCE && lineCount > 1) {
    const widest = Math.max(...lineWidths);
    const flows = lineCount > 2 &&
      lineWidths.slice(0, -1).every((width) => width >= PARAGRAPH_FLOWING_LINE_SHARE * widest);
    if (!flows) {
      const size = Math.max(...styles.map((style) => style.size));
      let width = right - left + 2 * centreRoom;
      // A row the source ends short of the block's width by more than the
      // next row's first word would take is ended by a break of its own (a
      // label's "2100" over "& beyond"), not by the measure: it bounds no
      // growth, and is a break still in the frame grown (setParagraph keeps
      // it so: withSourceHardBreaks).
      const hardStarts: number[] = [];
      for (let line = 0; line < lineCount - 1; line += 1) {
        const nextStart = lineStarts[line + 1];
        const next = pieces.find((piece) => piece.start === nextStart);
        if (!next) {
          width = right - left;
          break;
        }
        const blank = pieces.find((piece) => piece.blank && piece.end === nextStart);
        if (right - left - lineWidths[line] > (blank?.width ?? 0) + next.width + 0.5 * size) {
          hardStarts.push(nextStart);
          continue;
        }
        width = Math.min(width, lineWidths[line] + (blank?.width ?? 0) + next.width - 2 * LAYOUT_TOLERANCE);
      }
      const grow = (width - (right - left)) / 2;
      if (grow > LAYOUT_TOLERANCE) {
        const [keptLeft, keptRight] = [left, right];
        left -= grow;
        right += grow;
        const held = hardStarts.length
          ? matches(setParagraphLines(kernel, own, styles, blanks, availableIn(right), hardStarts))
          : holds(right);
        if (!held) [left, right] = [keptLeft, keptRight];
        else if (hardStarts.length) return Object.freeze({ left, right });
      }
    }
  }
  if (left === inked.left && right === inked.right) return inked;
  if (!holds(right)) return inked;
  return Object.freeze({ left, right });
}

function colorOperator(color: PdfNativeDeviceColor): string {
  const components = color.components.map(formatNumber).join(" ");
  return `${components} ${color.kind === "gray" ? "g" : color.kind === "rgb" ? "rg" : "k"}`;
}

/** A colour as this writer spells it in a stream (colorOperator: each
 * component in formatNumber's digits) -- what the page then states, and
 * what a reader of the stream reads it back as. */
function colorAsWritten(color: PdfNativeDeviceColor): PdfNativeDeviceColor {
  return Object.freeze({ ...color, components: Object.freeze(color.components.map((component) => Number(formatNumber(component)))) }) as
    unknown as PdfNativeDeviceColor;
}













function formatsWithColoursAsWritten(resolve: MraiNativeParagraphRewriteFormatsResolver): MraiNativeParagraphRewriteFormatsResolver {
  return Object.assign(async (context: Parameters<MraiNativeParagraphRewriteFormatsResolver>[0]) => {
    const formats = await resolve(context);
    if (!formats.characters.some((format) => format?.color)) return formats;
    return Object.freeze({ ...formats, characters: Object.freeze(formats.characters.map((format) =>
      format?.color ? Object.freeze({ ...format, color: colorAsWritten(format.color) }) : format)) });
  }, { stateOnly: resolve.stateOnly, ...(resolve.leanReadByReader ? { leanReadByReader: true } : {}),
    ...(resolve.wordsReadByReader ? { wordsReadByReader: true } : {}) });
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

/**
 * Codes `from` to `to` of a source show as the block draws them again where
 * the page draws them: the array they were drawn with (its strings and the
 * numbers between them), the state the source drew it under said the way the
 * block says its own runs (setParagraph's kept shows), and the `Tm` that
 * places its first glyph of them -- leaning as asked, `du` along and `dv`
 * across the lines from where it stood. Null where it draws nothing there.
 */
function sourceShowPiece(
  kernel: MraiNativeParagraphRewriteKernel,
  paragraph: Pick<ReadParagraph, "basis" | "blockMatrix">,
  owned: OwnedShow,
  from: number,
  to: number,
): Readonly<{ style: CharacterStyle; word: number; array: string; codes: readonly number[];
  place: (slant: number, du: number, dv: number) => string }> | null {
  const { show, font } = owned;
  const [b0, b1, b2, b3] = paragraph.basis;
  const unit = font.codeUnitBytes;
  const first = showGlyphs(kernel, show, font)[from / unit];
  if (!first || !(to > from) || !(show.fontSize > 0)) return null;
  const counts = show.segmentCodeCounts.length ? show.segmentCodeCounts : [show.codes.length];
  const elements: string[] = [];
  const codes: number[] = [];
  let offset = 0;
  let pending = 0;
  for (const [segment, count] of counts.entries()) {
    const start = Math.max(offset, from), end = Math.min(offset + count, to);
    if (end > start) {
      if (pending !== 0 && elements.length) elements.push(formatNumber(pending));
      pending = 0;
      elements.push(hexCodes(show.codes.slice(start, end)));
      codes.push(...show.codes.slice(start, end));
    }
    offset += count;
    if (segment < counts.length - 1 && offset > from && offset < to) pending += show.segmentAdjustments[segment] ?? 0;
  }
  if (!elements.length) return null;
  // The show's text space in the block's: its size taken out of it.
  const across = owned.size / show.fontSize;
  const stretch = show.horizontalScale !== 0 ? owned.horizontalScale / show.horizontalScale : 1;
  if (!(across > 0) || !(stretch > 0) || !Number.isFinite(across) || !Number.isFinite(stretch)) return null;
  const style: CharacterStyle = Object.freeze({ fontName: show.fontName, font, preferred: NO_PREFERRED_CODES,
    size: owned.size, characterSpacing: show.characterSpacing * across, horizontalScale: show.horizontalScale,
    rise: show.textRise * across, slant: owned.slant, color: owned.color, twin: null, syntheticBold: owned.stroked });
  const linear = composeMatrix(composeMatrix(show.textMatrix, show.graphicsMatrix), invertMatrix(paragraph.blockMatrix));
  // (The numbers before its first glyph are in where it is placed.)
  const [x, y] = applyMatrix(linear, first.start * (show.horizontalScale / 100), 0);
  return Object.freeze({ style, word: show.wordSpacing * across, array: elements.join(" "), codes,
    place: (slant: number, du: number, dv: number) => `${[b0 * stretch, b1 * stretch, b2 + slant * b0, b3 + slant * b1,
      x + du * b0 + dv * b2, y + du * b1 + dv * b3].map(formatNumber).join(" ")} Tm\n` });
}

/** Whether two fonts are faces of one family, by their names (the Bold and
 * the Italic of one Times): what tells a face asked for its style from a
 * face of another family asked for itself. */
function facesOfOneFamily(left: MraiNativeParagraphRewriteFont, right: MraiNativeParagraphRewriteFont): boolean {
  if (!left.faceName || !right.faceName) return false;
  const key = pdfNativeFontFamilyKey(left.faceName);
  return key.length > 0 && key === pdfNativeFontFamilyKey(right.faceName);
}






















function formatsInFacesThatSetAll(
  kernel: MraiNativeParagraphRewriteKernel,
  text: readonly string[],
  formats: readonly (MraiNativeParagraphRewriteCharacterFormat | null | undefined)[],
  ownOf: (index: number) => Readonly<{ fontName: string; font: MraiNativeParagraphRewriteFont; syntheticBold: boolean; slant: number }>,
): (MraiNativeParagraphRewriteCharacterFormat | null | undefined)[] {
  const out = [...formats];
  const writes = (character: string, font: MraiNativeParagraphRewriteFont) => {
    try {
      kernel.encodeText(character, font, NO_PREFERRED_CODES);
      return true;
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
      return false;
    }
  };
  const faceAt = (index: number) => isBlank(text[index]) ? undefined : formats[index]?.face;
  for (let index = 0; index < text.length;) {
    const face = faceAt(index);
    if (!face) {
      index += 1;
      continue;
    }
    // The stretch asked of this face: on through the blanks between its words.
    let end = index + 1;
    for (let at = end; at < text.length; at += 1) {
      if (isBlank(text[at])) continue;
      if (faceAt(at)?.resourceName !== face.resourceName) break;
      end = at + 1;
    }
    const siblings: number[] = [];
    let lacking = -1;
    for (let at = index; at < end; at += 1) {
      if (faceAt(at)?.resourceName !== face.resourceName) continue;
      const own = ownOf(at);
      if (own.fontName === face.resourceName || !facesOfOneFamily(own.font, face.font)) continue;
      siblings.push(at);
      if (lacking < 0 && !writes(text[at], face.font)) lacking = at;
    }
    if (lacking >= 0) {
      const [faceBold, faceItalic] = faceStyle(face.font);
      let drawn: Readonly<{ syntheticBold?: boolean; slant?: number }> = {};
      for (const at of siblings) {
        const format = formats[at]!;
        const own = ownOf(at);
        const [ownBold, ownItalic] = faceStyle(own.font);
        const wantBold = faceBold === 1 || (format.syntheticBold ?? (face.syntheticBold === true ? true : own.syntheticBold));
        const wantItalic = faceItalic === 1 || (format.slant ?? (face.slanted ? SYNTHETIC_OBLIQUE_SLANT : own.slant)) !== 0;
        // (The face's own refusal of the letter it lacks.)
        if ((ownBold === 1 && !wantBold) || (ownItalic === 1 && !wantItalic)) {
          kernel.encodeText(text[lacking], face.font, NO_PREFERRED_CODES);
        }
        const { face: _face, syntheticBold: _bold, slant: _slant, ...rest } = format;
        void _face; void _bold; void _slant;
        drawn = Object.freeze({
          ...(ownBold === 1 || wantBold === own.syntheticBold ? {} : { syntheticBold: wantBold }),
          ...(ownItalic === 1 ? {} : { slant: wantItalic ? own.slant || SYNTHETIC_OBLIQUE_SLANT : 0 }),
        });
        out[at] = Object.freeze({ ...rest, ...drawn });
      }
      // The blanks between them as the words beside them are drawn.
      for (let at = index; at < end; at += 1) {
        const format = formats[at];
        if (!format || !isBlank(text[at])) continue;
        const { syntheticBold: _bold, slant: _slant, ...rest } = format;
        void _bold; void _slant;
        out[at] = Object.freeze({ ...rest, ...drawn });
      }
    }
    index = end;
  }
  return out;
}
































function refuseStyleFaceThatLeavesTextUnstyled(
  kernel: MraiNativeParagraphRewriteKernel,
  text: readonly string[],
  asked: readonly (MraiNativeParagraphRewriteCharacterFormat | null | undefined)[],
  set: readonly (MraiNativeParagraphRewriteCharacterFormat | null | undefined)[],
  ownOf: (index: number) => Readonly<{ fontName: string; font: MraiNativeParagraphRewriteFont; syntheticBold: boolean; slant: number }>,
): void {
  for (const [index, character] of text.entries()) {
    if (isBlank(character) || asked[index]?.weightLeftToOwnFace !== true) continue;
    const own = ownOf(index);
    if (faceStyle(own.font)[0] !== 1 && !(set[index]?.syntheticBold ?? own.syntheticBold)) refuse(MRAI_NATIVE_PARAGRAPH_STYLE_FACE_REFUSAL);
  }
  let face: MraiNativeParagraphRewriteFace | undefined;
  let ofTheFamily = false;
  for (const [index, character] of text.entries()) {
    if (isBlank(character)) continue;
    const named = asked[index]?.face;
    if (!named || (face && named.resourceName !== face.resourceName)) return;
    face = named;
    ofTheFamily ||= asked[index]?.faceForStyle === true ||
      (ownOf(index).fontName !== named.resourceName && facesOfOneFamily(ownOf(index).font, named.font));
  }
  if (!face || !ofTheFamily) return;
  const [faceBold, faceItalic] = faceStyle(face.font);
  if (!faceBold && !faceItalic) return;
  const writes = (character: string, font: MraiNativeParagraphRewriteFont) => {
    try {
      kernel.encodeText(character, font, NO_PREFERRED_CODES);
      return true;
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
      return false;
    }
  };
  for (const [index, character] of text.entries()) {
    if (isBlank(character)) continue;
    const format = set[index];
    const own = ownOf(index);
    const drawnIn = format?.face && writes(character, format.face.font) ? format.face : null;
    const [bold, italic] = faceStyle(drawnIn?.font ?? own.font);
    const stroked = format?.syntheticBold ?? (drawnIn?.syntheticBold === true ? true : own.syntheticBold);
    const leaned = (format?.slant ?? (drawnIn?.slanted ? SYNTHETIC_OBLIQUE_SLANT : own.slant)) !== 0;
    if ((faceItalic === 1 && italic !== 1 && !leaned) || (faceBold === 1 && bold !== 1 && !stroked)) {
      refuse(MRAI_NATIVE_PARAGRAPH_STYLE_FACE_REFUSAL);
    }
  }
}

/** The widest gap the editor's boxes read as one space before a run set in a
 * leaned frame of its own (native-paragraph-edit-boxes' folded spaces), in
 * shares of the size of the run before it. */
const READER_FOLDED_SPACE_MAX_EM = 1.5;

/** What draws a character otherwise without moving it. */
interface CharacterLook {
  readonly color: PdfNativeDeviceColor;
  readonly bold: boolean;
  readonly slant: number;
  readonly underline: boolean;
  readonly strikethrough: boolean;
  readonly linkRule: ParagraphLinkRule | null;
  readonly twin: readonly [number, number] | null;
}

function lookKey(look: CharacterLook): string {
  return JSON.stringify([look.color.kind, look.color.components, look.bold, look.slant, look.underline, look.strikethrough,
    look.linkRule, look.twin]);
}























function setParagraphWhereItStands(
  kernel: MraiNativeParagraphRewriteKernel,
  paragraph: ReadParagraph,
  /** The new text, a line for each break typed in it (setParagraph's). */
  typedLines: readonly string[],
  formats: MraiNativeParagraphRewriteFormats,
  /** Shows the block draws otherwise than by their own arrays (drawParagraphWhereItStands). */
  setAgain?: ShowSetAgain,
): ReturnType<typeof setParagraph> | null {
  const replacementText = typedLines.join("");
  if (formats.alignment !== undefined || formats.leftIndent || formats.lineStep !== undefined ||
    formats.firstLineIndent !== undefined || paragraph.frameRequested || paragraph.sourceLineStep !== undefined ||
    paragraph.sizeScale !== undefined || replacementText !== paragraph.selection.item.str) return null;
  const text = characters(replacementText);
  const { layout, shows } = paragraph;
  if (formats.characters.length !== text.length || text.some((_, index) => !shows[paragraph.characterShows[index]])) return null;
  const showOf = (index: number) => shows[paragraph.characterShows[index]];
  // A break in the text stands where the source starts a line (the break the
  // source's own short line keeps, withSourceHardBreaks): one typed anywhere
  // else is a new line, which the tier sets.
  const lineStarts = new Set<number>();
  const linesSeen = new Set<number>();
  for (const [index, character] of text.entries()) {
    if (isBlank(character) || linesSeen.has(showOf(index).line)) continue;
    linesSeen.add(showOf(index).line);
    lineStarts.add(index);
  }
  for (let count = 0, line = 0; line < typedLines.length - 1; line += 1) {
    count += characters(typedLines[line]).length;
    if (!lineStarts.has(count)) return null;
  }
  const ownOf = (index: number) => {
    const owned = showOf(index);
    return { fontName: owned.show.fontName, font: owned.font, syntheticBold: owned.stroked, slant: owned.slant };
  };
  const asked = formatsInFacesThatSetAll(kernel, text, formats.characters, ownOf);
  refuseStyleFaceThatLeavesTextUnstyled(kernel, text, formats.characters, asked, ownOf);
  const near = (left: number, right: number) => Math.abs(left - right) <= 1e-6 * Math.max(1, Math.abs(left), Math.abs(right));
  // What a show is drawn as, asked a format (or none) of.
  const lookOf = (owned: OwnedShow, format: (typeof asked)[number] | null,
    ruled: Readonly<{ underline?: true; strikethrough?: true }> | undefined): CharacterLook => Object.freeze({
    color: format?.color ?? owned.color,
    // (A bold asked of a show in a bold face: as it stands, strokeAskedChanges.)
    bold: strokeAskedChanges(format?.syntheticBold, owned) ? format!.syntheticBold! : owned.stroked,
    slant: format?.slant ?? owned.slant,
    underline: format?.underline ?? Boolean(owned.underline || ruled?.underline),
    strikethrough: format?.strikethrough ?? Boolean(owned.strikethrough || ruled?.strikethrough),
    // An underline asked (or taken off) is this writer's: a link's own rule goes with it.
    linkRule: format?.underline !== undefined ? null : owned.linkRule ?? null,
    // A fake-bold pass is the source show's own second drawing: not drawn where a weight is asked.
    twin: format?.syntheticBold !== undefined && !(format.syntheticBold && standsInBoldFace(owned)) ? null : owned.twin,
  });
  const looks: CharacterLook[] = [];
  for (const [index] of text.entries()) {
    const owned = showOf(index);
    const format = asked[index] ?? null;
    if (format) {
      const size = format.sizePoints ?? owned.size * (format.sizeScale ?? 1);
      if ((format.face && format.face.resourceName !== owned.show.fontName) || !near(size, owned.size) ||
        (format.riseEm !== undefined && !near(format.riseEm * size, owned.rise)) ||
        (format.characterSpacingEm !== undefined && !near(format.characterSpacingEm * size, owned.characterSpacing))) return null;
    }
    looks.push(lookOf(owned, format, paragraph.characterRules.get(index)));
  }
  // A show of glyphs the reader reads as blanks (OwnedShow.readsBlank) draws
  // no character of the text, so none is asked a format for it: it goes with
  // its line, in what is asked of the character drawn before it there (of the
  // first after it, where it is drawn first). Left as it stood, a line made
  // bold or red kept the marks over its letters thin and black.
  const lookBeside = (owner: number, line: number): CharacterLook | null => {
    let before: number | null = null, after: number | null = null;
    for (const [index, character] of text.entries()) {
      const at = paragraph.characterShows[index];
      if (isBlank(character) || shows[at].line !== line || at === owner) continue;
      if (at < owner) before = index;
      else if (after === null) after = index;
    }
    const beside = before ?? after;
    return beside === null ? null : lookOf(shows[owner], asked[beside] ?? null, undefined);
  };
  // From here the edit is of this kind: written where it stands, or not.
  const cannot = (): never => refuse(MRAI_NATIVE_PARAGRAPH_FORMAT_IN_PLACE_REFUSAL);
  const opening = text.findIndex((character) => !isBlank(character));
  if (opening < 0) cannot();
  // The source's own lines, as characters of the text (setParagraph's kept lines).
  const lineFirst: number[] = [];
  for (let index = opening; index < text.length; index += 1) {
    const line = showOf(index).line;
    if (index > opening && line < showOf(index - 1).line) cannot();
    if (lineFirst[line] === undefined) lineFirst[line] = index - opening;
  }
  if (lineFirst[0] !== 0 || lineFirst.length !== layout.lineCount || paragraph.lineBaselines.length < layout.lineCount ||
    layout.lineExtents.length < layout.lineCount) cannot();
  const lines: SetLine[] = [];
  for (let line = 0; line < layout.lineCount; line += 1) {
    if (lineFirst[line] === undefined) cannot();
    const end = lineFirst[line + 1] ?? text.length - opening;
    let inkEnd = end;
    while (inkEnd > lineFirst[line] && isBlank(text[inkEnd - 1 + opening])) inkEnd -= 1;
    const extent = layout.lineExtents[line];
    lines.push(Object.freeze({ start: lineFirst[line], end, inkEnd, inkWidth: extent.right - extent.left }));
  }
  // The characters each show draws, in the text.
  const showCharacters = new Map<number, number[]>();
  for (const [index] of text.entries()) {
    const owner = paragraph.characterShows[index];
    const own = showCharacters.get(owner);
    if (own) own.push(index);
    else showCharacters.set(owner, [index]);
  }
  const piecesOf = (owner: number, owned: OwnedShow): readonly StandingPiece[] => {
    const { show, font } = owned;
    const unit = font.codeUnitBytes;
    const from = owned.partial?.codeStart ?? 0, to = owned.partial?.codeEnd ?? show.codes.length;
    const inked = (showCharacters.get(owner) ?? []).filter((index) => !isBlank(text[index]));
    // The stretches of its characters drawn in one look.
    const runs: { first: number; last: number; look: CharacterLook }[] = [];
    for (const [at, index] of inked.entries()) {
      const last = runs.at(-1);
      if (last && lookKey(last.look) === lookKey(looks[index])) last.last = at;
      else runs.push({ first: at, last: at, look: looks[index] });
    }
    // Each stretch's codes of the show, where its ink starts and ends, and
    // its first and last character.
    const pieces: StandingPiece[] = [];
    if (runs.length <= 1) {
      pieces.push({ from, to, look: runs[0]?.look ?? (owned.readsBlank || paragraph.unreadGlyphs ? lookBeside(owner, owned.line) : null),
        ink: owned.inked ? [owned.inked.start, owned.inked.end] : null,
        characters: inked.length ? [inked[0], inked.at(-1)!] : null });
    } else {
      // The glyph that draws each of its characters: one each, or a
      // ligature's glyph its letters (as readParagraph reads a stack).
      const drawn = showGlyphs(kernel, show, font).flatMap((glyph, index) =>
        !/^\s*$/u.test(glyph.text ?? "x") && index * unit >= from && index * unit < to ? [{ glyph, index }] : []);
      if (drawn.length !== owned.ink.length) cannot();
      const counts = drawn.length === inked.length ? drawn.map(() => 1)
        : drawn.map(({ glyph }) => characters(kernel.normalizeUnicode(glyph.text ?? "x"))
          .filter((character) => !isReaderBlank(character)).length);
      if (counts.some((count) => count < 1) || counts.reduce((sum, count) => sum + count, 0) !== inked.length) cannot();
      const glyphOf: { at: number; first: boolean }[] = [];
      for (const [at, count] of counts.entries()) for (let part = 0; part < count; part += 1) glyphOf.push({ at, first: part === 0 });
      for (const [order, run] of runs.entries()) {
        // (A look that changes inside a ligature divides a glyph.)
        if (!glyphOf[run.first].first) cannot();
        const next = runs[order + 1];
        const lastGlyph = next ? glyphOf[next.first].at - 1 : drawn.length - 1;
        pieces.push({ from: order === 0 ? from : drawn[glyphOf[run.first].at].index * unit,
          to: next ? drawn[glyphOf[next.first].at].index * unit : to, look: run.look,
          ink: [owned.ink[glyphOf[run.first].at][0], owned.ink[lastGlyph][1]], characters: [inked[run.first], inked[run.last]] });
      }
    }
    return pieces;
  };
  return drawParagraphWhereItStands(kernel, paragraph, lines, piecesOf,
    (after, before) => text.slice(after + 1, before).some(isBlank), setAgain);
}

/** What the block draws in the place of one source show: runs of codes, each
 * placed by its own matrix and drawn in its own style. Null: the show is
 * drawn again by its own array, as every show is. */
type ShowSetAgain = (owner: number, owned: OwnedShow) => readonly Readonly<{
  place: string; style: CharacterStyle; word: number; array: string; codes: readonly number[] }>[] | null;

/** A stretch of a source show drawn again in one look: its codes of the show,
 * the look (null: the show's own, no rule), where its ink starts and ends
 * along the line, and its first and last character of the text (null where
 * the reader's characters are not what places it). */
type StandingPiece = Readonly<{ from: number; to: number; look: CharacterLook | null; ink: readonly [number, number] | null;
  characters: readonly [number, number] | null }>;

/** A show's frame as the reader compares in it: its axis along the line and
 * its up axis on the page, each at the size its glyphs are drawn. */
type ReaderFrame = readonly [number, number, number, number];

/**
 * Whether the reader ends its line between two glyphs (pdf.js
 * compareWithLastPosition): `from`, where the pen stood after the glyph
 * before, drawn in the frame `last`; `to`, the origin of the glyph after,
 * drawn in `current`. It compares on the page where the glyph after is set
 * along a page axis, and each point through its own frame where it is not --
 * so a leaned glyph is compared in a frame whose second coordinate grows
 * with the distance along the line.
 */
function pdfJsReaderEndsLine(last: ReaderFrame, from: readonly [number, number], current: ReaderFrame,
  to: readonly [number, number]): boolean {
  const tiny = 1e-9 * Math.hypot(...current);
  const zero = (value: number) => Math.abs(value) <= tiny;
  let [x, y] = to, [lastX, lastY] = from;
  if (!zero(current[0]) && zero(current[1]) && zero(current[2])) {
    if (current[0] < 0) [x, y, lastX, lastY] = [-x, -y, -lastX, -lastY];
  } else if (!zero(current[1]) && zero(current[0]) && zero(current[3])) {
    if (current[1] > 0) [x, y, lastX, lastY] = [y, x, lastY, lastX];
    else [x, y, lastX, lastY] = [-y, -x, -lastY, -lastX];
  } else {
    const scale = Math.hypot(current[0], current[1]), lastScale = Math.hypot(last[0], last[1]);
    [x, y] = [(current[0] * x + current[1] * y) / scale, (current[2] * x + current[3] * y) / scale];
    [lastX, lastY] = [(last[0] * lastX + last[1] * lastY) / lastScale, (last[2] * lastX + last[3] * lastY) / lastScale];
  }
  const drawnBack = x - lastX < PDFJS_NEGATIVE_SPACE_FACTOR * Math.hypot(last[0], last[1]);
  return Math.abs(y - lastY) > (drawnBack ? 0.5 : 1) * Math.hypot(last[2], last[3]);
}




const FINAL_STANDING_REFUSALS = new WeakSet<object>();



















function refuseLeanAmongOtherTextOnItsLines(
  paragraph: Pick<ReadParagraph, "allShows">,
  shearOf: ReadonlyMap<number, number>,
  /** The signs leaned with the lines they hang on (hangingSignLeans), by the
   * end of each one's span: neither the box's text nor a sign left upright. */
  leanedSigns: ReadonlySet<number> = new Set(),
): void {
  type Placed = Readonly<{ at: readonly [number, number]; along: readonly [number, number]; size: number }>;
  const placed = (show: MraiNativeParagraphRewriteShow): Placed => {
    const matrix = composeMatrix(show.textMatrix, show.graphicsMatrix);
    const length = Math.hypot(matrix[0], matrix[1]) || 1;
    return { at: applyMatrix(matrix, 0, show.textRise), along: [matrix[0] / length, matrix[1] / length],
      size: Math.abs(show.fontSize) * Math.hypot(matrix[2], matrix[3]) };
  };
  // How far off the baseline of `line` the point stands.
  const off = (line: Placed, point: readonly [number, number]) =>
    Math.abs((point[0] - line.at[0]) * line.along[1] - (point[1] - line.at[1]) * line.along[0]);
  let before: Placed | null = null;
  let between: Placed[] = [];
  for (const show of paragraph.allShows) {
    if (show.codes.length === 0 || show.vertical || leanedSigns.has(show.end)) continue;
    const here = placed(show);
    if (!shearOf.has(show.end)) {
      if (before) between.push(here);
      continue;
    }
    if (before && between.length && off(before, here.at) <= 0.25 * before.size &&
      between.some((other) => off(before!, other.at) > 0.25 * before!.size && off(before!, other.at) <= before!.size)) {
      const refusal = new PdfNativeTransactionError("unsupported-document", MRAI_NATIVE_PARAGRAPH_READ_BACK_REFUSAL);
      FINAL_STANDING_REFUSALS.add(refusal);
      throw refusal;
    }
    before = here;
    between = [];
  }
}

/**
 * The signs that hang on a line of the box and are no show of it, each with
 * the shear it takes when the box is leaned from `standing` to `slant`: a
 * show of other text drawn between two of the box's shows of one line --
 * after the one, before the other, in the stream and along the line -- off
 * that line's baseline by a quarter of a size to a whole one, in the space
 * and along the direction the box is drawn in, and leaning as the box leans
 * now. The radical sign TeX hangs over the line it stands on is one: drawn
 * after the «/» and before the «2» of «1/√2», 8 pt above the baseline, and
 * no character of the paragraph's text (the reader's boxes keep it out).
 *
 * Such a sign is a part of the line to the eye, and the box's by the same
 * rule that gives a mark drawn on a letter to that letter: it leans with the
 * line it hangs on. Left upright among leaned glyphs it also cut the
 * paragraph at every sign when the page was read again (pdf-walk-1000 0608:
 * 35 boxes read as 45), for which the lean was refused outright; leaned with
 * its line the page reads as it did. It is carried by the lean as the line
 * is: a point of the line that high above the baseline moves along it by the
 * lean times its height, so the sign is set that much further on (`hang`,
 * in the units of the matrix it is drawn under; `moved`, on the page) and
 * leans from there -- its foot stays by the glyph before it, its head meets
 * the rule TeX draws after it a little further in. Leaned about its own
 * origin where it stood, the sign's foot swung back across the «/».
 *
 * A show of other text on the baseline itself between two of the box's (a
 * word of the column beside it, 0601, 0661), further off than a size, or
 * drawn elsewhere along the line, is no such sign and stays as it stands.
 */
function hangingSignLeans(
  paragraph: Pick<ReadParagraph, "allShows" | "shows" | "basis" | "blockMatrix" | "lineBaselines">,
  standing: number,
  slant: number,
): ReadonlyMap<number, Readonly<{ shear: number; hang: number; moved: readonly [number, number] }>> {
  const { basis } = paragraph;
  const toBlock = invertMatrix(paragraph.blockMatrix);
  const ownedByEnd = new Map(paragraph.shows.map((show) => [show.end, show] as const));
  const leans = new Map<number, Readonly<{ shear: number; hang: number; moved: readonly [number, number] }>>();
  let before: OwnedShow | null = null;
  let between: MraiNativeParagraphRewriteShow[] = [];
  for (const show of paragraph.allShows) {
    if (show.codes.length === 0 || show.vertical) continue;
    const owned = ownedByEnd.get(show.end);
    if (!owned) {
      if (before) between.push(show);
      continue;
    }
    if (before && between.length && before.line === owned.line) {
      const baseline = paragraph.lineBaselines[owned.line];
      const size = before.size;
      const from = before.inked?.start ?? before.u, to = owned.inked?.start ?? owned.u;
      for (const sign of between) {
        if (!sign.graphicsMatrix.every((value, at) => Math.abs(value - before!.show.graphicsMatrix[at]) <= 1e-9 * Math.max(1, Math.abs(value))) ||
          sign.textRise !== 0 || !(sign.horizontalScale > 0)) continue;
        const linear = composeMatrix(composeMatrix(sign.textMatrix, sign.graphicsMatrix), toBlock);
        const along = Math.hypot(linear[0], linear[1]), fullAcross = Math.hypot(linear[2], linear[3]);
        if (!(along > 0) || !(fullAcross > 0) || Math.abs(linear[0] * basis[1] - linear[1] * basis[0]) > 1e-6 * along ||
          linear[0] * basis[0] + linear[1] * basis[1] <= 0 || linear[2] * basis[2] + linear[3] * basis[3] <= 0) continue;
        const leaning = Math.abs(linear[2] * basis[3] - linear[3] * basis[2]) > 1e-6 * fullAcross;
        const across = leaning ? linear[2] * basis[2] + linear[3] * basis[3] : fullAcross;
        const own = leaning ? (linear[2] * basis[0] + linear[3] * basis[1]) / across : 0;
        if (Math.abs(own - standing) > 1e-6) continue;
        const [x, y] = applyMatrix(toBlock, ...applyMatrix(sign.graphicsMatrix, ...applyMatrix(sign.textMatrix, 0, 0)));
        const u = x * basis[0] + y * basis[1], v = x * basis[2] + y * basis[3];
        const off = Math.abs(v - baseline);
        if (!(off > 0.25 * size && off <= size) || u < from - 0.25 * size || u > to + 0.25 * size) continue;
        // (Along its own line: the lean times how high it hangs, in the matrix's own units.)
        const hang = (slant - standing) * (v - baseline) / along;
        const onPage = composeMatrix(sign.textMatrix, sign.graphicsMatrix);
        leans.set(sign.end, Object.freeze({ shear: (slant - standing) * (across / along), hang,
          moved: Object.freeze([hang * onPage[0], hang * onPage[1]] as const) }));
      }
    }
    before = owned;
    between = [];
  }
  return leans;
}

/**
 * The guard of a lean's read-back, for a box whose shows stay where the page
 * draws them (drawn again in a block where they stand, or leaned in place by
 * their own matrices): asked stretch by stretch in the order the lines are
 * drawn, it refuses a lean the reader would read back as other text.
 */
function leanReadBackGuard(
  kernel: MraiNativeParagraphRewriteKernel,
  paragraph: Pick<ReadParagraph, "basis" | "blockMatrix" | "unreadGlyphs">,
): Readonly<{
  /** A stretch of a show drawn at `slant`, inked from `ink[0]` to `ink[1]` along the line, `across` it. */
  refuseLeanReadOtherwise: (owned: OwnedShow, slant: number, ink: readonly [number, number], across: number) => void;
  /** A stretch of white space alone was drawn. */
  blankDrawn: () => void;
}> {
  const [b0, b1, b2, b3] = paragraph.basis;
  // What a lean changes of how the page reads back. In a leaned frame the
  // reader measures a glyph against the visible glyph before it in each one's
  // own frame (pdf.js compareWithLastPosition: applyInverseRotation divides
  // both of a point's coordinates by the frame's length ALONG the line), and
  // ends its line where the two then stand further apart across the line
  // than the item is high (half that, drawn back from it). Ending its line
  // there, it reads no space there either -- and the editor's boxes join
  // what stands on one baseline as it comes. So where two glyphs of a line
  // the reader read a space between come to stand a line apart to it, the
  // box opens again with that space gone:
  //   - upright, a line whose shows are set at more than one width (pdfTeX's
  //     font expansion: `1.02 0 0 1 ... Tm` beside `1 0 0 1 ... Tm`) reads as
  //     one line; leaned, each change of width between two words read as a
  //     line end, «aboutRe bω= 3πjm/2, wherer0is» (pdf-walk-1000 0676);
  //   - a justified heading's two words set five sizes apart stand, leaned,
  //     more than a line apart in the reader's frame (the lean times the
  //     gap), «EXPRESSEMPLOYMENT» (0444).
  // Such a lean is a write that reads back otherwise: refused. A line end
  // the reader comes to see where it read no space -- a footnote mark set
  // against its word in a frame of another width (0527) -- changes no text
  // and no row of the box, and is no reason to refuse.
  const blockMatrix = paragraph.blockMatrix;
  const onPage = (u: number, across: number) => applyMatrix(blockMatrix, u * b0 + across * b2, u * b1 + across * b3);
  const readerEndsLine = pdfJsReaderEndsLine;
  const readerFrame = (owned: OwnedShow, slant: number): ReaderFrame => {
    const along = owned.size * (owned.horizontalScale / 100);
    const [ax, ay] = [b0 * along, b1 * along];
    const [ux, uy] = [(b2 + slant * b0) * owned.size, (b3 + slant * b1) * owned.size];
    return [ax * blockMatrix[0] + ay * blockMatrix[2], ax * blockMatrix[1] + ay * blockMatrix[3],
      ux * blockMatrix[0] + uy * blockMatrix[2], ux * blockMatrix[1] + uy * blockMatrix[3]];
  };
  // Where a show draws white space between its inked glyphs (`before`: by
  // inked glyph, ahead of it since the inked glyph before), before the first
  // and after the last. Null where its glyphs cannot be told as its ink was
  // read: then white space is taken to stand everywhere.
  const blanksOf = new Map<OwnedShow, Readonly<{ before: readonly boolean[]; leading: boolean; trailing: boolean }> | null>();
  const blanks = (owned: OwnedShow) => {
    if (blanksOf.has(owned)) return blanksOf.get(owned)!;
    let read: Readonly<{ before: readonly boolean[]; leading: boolean; trailing: boolean }> | null = null;
    try {
      const unit = owned.font.codeUnitBytes;
      const own = showGlyphs(kernel, owned.show, owned.font).filter((_, index) => !owned.partial ||
        (index * unit >= owned.partial.codeStart && index * unit < owned.partial.codeEnd));
      for (const marks of [false, true]) {
        const inks = own.map((glyph) => !/^\s*$/u.test(glyph.text ?? "x") ||
          (marks && Boolean(owned.font.inkedBlankCodes?.has(glyph.code))));
        if (inks.filter(Boolean).length !== owned.ink.length) continue;
        const before: boolean[] = [];
        let blank = false;
        for (const inked of inks) {
          if (inked) {
            before.push(blank);
            blank = false;
          } else blank = true;
        }
        read = Object.freeze({ before: Object.freeze(before), leading: before[0] === true, trailing: blank });
        break;
      }
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
    }
    blanksOf.set(owned, read);
    return read;
  };
  // Whether the reader reads a space between two glyphs of a line it reads
  // on from one another: a white-space glyph drawn between them, or a gap of
  // its tracking bound or more, at the size of the glyph before (taken a
  // tenth under it: what stands at the bound is read either way).
  const readsSpace = (gap: number, owned: OwnedShow, blank: boolean) => blank ||
    gap >= 0.9 * PDFJS_TRACKING_SPACE_FACTOR * owned.size *
      (owned.show.horizontalScale !== 0 ? owned.horizontalScale / owned.show.horizontalScale : 1);
  // Where the reader ends its line between two glyphs only because their
  // matrix leans, the boxes read the space back themselves (the box reader's
  // itemAfterSkewedFrameLineEnd: the item after stands on the item before's
  // baseline, in its size, its tracking bound or more further on -- taken a
  // tenth over the bound here, in the leaned item's own measure, since what
  // stands at the bound is read either way). Only between two glyphs that are
  // each a character of the text with a width of its own: a mark of no width
  // (the stroke of TeX's ≠, drawn over the sign after it) is no end of an
  // item the reader measures a gap from, and a paragraph that holds glyphs
  // its font names no character for keeps every such gap a spacer.
  // In one size on one baseline -- or a script beside a word of its line (a
  // subscript and the word after it, «r0 is»; a superscript and the word
  // after it): `apart`, how far apart across the line the two stand, is then
  // no more than the box reader's bound for a raised item of a line (0.55
  // of the larger size) nor than the glyph before is high, past which the
  // reader ends its line on the page too. The gap is measured in the size of
  // the glyph before, a script's as well (pdf-walk-1000 0676, 0861).
  // (`unnamedBefore`: whether the glyph before stands in a run of the box's
  // shows -- one line, one face, one size, as the reader makes an item --
  // that holds a glyph its font names no character for. The box reader keeps
  // the gap after such an item a spacer, and that one alone: a paragraph
  // with a formula's unnamed sign on one line reads the spaces between the
  // plain words of its other runs back as anywhere, 0676.)
  const spaceReadBack = (gap: number, before: OwnedShow, beforeGlyph: number, after: OwnedShow, afterGlyph: number,
    slant: number, apart = 0, unnamedBefore = false) => {
    if (before.readsBlank || after.readsBlank) return false;
    if (paragraph.unreadGlyphs && (unnamedBefore || holdsUnnamed(before) || holdsUnnamed(after))) return false;
    const [left, right] = [before.ink[beforeGlyph], after.ink[afterGlyph]];
    if (!left || !right || !(left[1] - left[0] > 1e-6) || !(right[1] - right[0] > 1e-6)) return false;
    const sameSet = Math.abs(before.size - after.size) <= 1e-6 * Math.max(1, before.size) &&
      Math.abs(before.rise - after.rise) <= 1e-6 && apart <= 1e-6;
    if (!sameSet && !(apart <= Math.min(LEANED_SCRIPT_BASELINE_RATIO * Math.max(before.size, after.size), before.size))) return false;
    return gap >= 1.1 * PDFJS_TRACKING_SPACE_FACTOR * before.size * Math.hypot(1, slant);
  };
  // Whether a show draws a glyph its font names no character for (or names
  // by a name no list reads); a show whose glyphs cannot be read counts.
  const unnamedOf = new Map<OwnedShow, boolean>();
  const holdsUnnamed = (owned: OwnedShow): boolean => {
    let known = unnamedOf.get(owned);
    if (known === undefined) {
      try {
        known = showGlyphs(kernel, owned.show, owned.font).some((glyph) => glyph.text === undefined ||
          Boolean(owned.font.nameUnreadCodes?.has(glyph.code)));
      } catch (error) {
        if (!(error instanceof PdfNativeTransactionError)) throw error;
        known = true;
      }
      unnamedOf.set(owned, known);
    }
    return known;
  };
  // The stretch drawn before, on this line or the one above, and whether a
  // show of white space alone was drawn since. (`unnamedRun`: see spaceReadBack.)
  let leanedBefore: Readonly<{ owned: OwnedShow; slant: number; end: number; across: number; last: number; unnamedRun: boolean }> | null = null;
  let blankSince = false;
  const refuseLeanReadOtherwise = (owned: OwnedShow, slant: number, ink: readonly [number, number], across: number) => {
    const changed = slant !== owned.slant;
    const last = leanedBefore;
    const own = blanks(owned);
    // (Its inked glyphs of this stretch, as the show's.)
    const first = owned.ink.findIndex(([start]) => start >= ink[0] - 1e-9);
    const inside = owned.ink.flatMap(([start, end], index) => start >= ink[0] - 1e-9 && end <= ink[1] + 1e-9 ? [index] : []);
    if (last && slant !== 0 && slant === last.slant && (changed || last.slant !== last.owned.slant)) {
      const from = onPage(last.end, last.across), to = onPage(ink[0], across);
      if (readerEndsLine(readerFrame(last.owned, last.owned.slant), from, readerFrame(owned, owned.slant), to) !==
        readerEndsLine(readerFrame(last.owned, slant), from, readerFrame(owned, slant), to)) {
        const lastBlanks = blanks(last.owned);
        const blank = blankSince || !own || !lastBlanks || first < 0 ||
          (first === 0 ? own.leading : own.before[first]) ||
          (last.last === last.owned.ink.length - 1 ? lastBlanks.trailing : lastBlanks.before[last.last + 1] ?? true);
        if (readsSpace(ink[0] - last.end, last.owned, blank) &&
          !(first >= 0 &&
            spaceReadBack(ink[0] - last.end, last.owned, last.last, owned, first, slant, Math.abs(across - last.across),
              last.unnamedRun))) {
          refuse(MRAI_NATIVE_PARAGRAPH_READ_BACK_REFUSAL);
        }
      }
    }
    // (And from glyph to glyph inside it: a gap its array opens.)
    if (slant !== 0 && changed) {
      const [was, now] = [readerFrame(owned, owned.slant), readerFrame(owned, slant)];
      for (let at = 1; at < inside.length; at += 1) {
        const [before, glyph] = [owned.ink[inside[at - 1]], owned.ink[inside[at]]];
        const from = onPage(before[1], across), to = onPage(glyph[0], across);
        if (readerEndsLine(was, from, was, to) !== readerEndsLine(now, from, now, to) &&
          readsSpace(glyph[0] - before[1], owned, !own || own.before[inside[at]] !== false) &&
          !spaceReadBack(glyph[0] - before[1], owned, inside[at - 1], owned, inside[at], slant)) {
          refuse(MRAI_NATIVE_PARAGRAPH_READ_BACK_REFUSAL);
        }
      }
    }
    // (A run: the shows of one line drawn one after the other in one face and size.)
    const runsOn = last !== null && last.owned.line === owned.line && last.owned.show.fontName === owned.show.fontName &&
      Math.abs(last.owned.size - owned.size) <= 1e-6 * Math.max(1, owned.size);
    leanedBefore = { owned, slant, end: ink[1], across, last: inside.at(-1) ?? owned.ink.length - 1,
      unnamedRun: (runsOn && last!.unnamedRun) || (paragraph.unreadGlyphs === true && holdsUnnamed(owned)) };
    blankSince = false;
  };
  return Object.freeze({ refuseLeanReadOtherwise, blankDrawn: () => { blankSince = true; } });
}

/**
 * The block that draws a paragraph again where it stands
 * (setParagraphWhereItStands): each show's pieces by the page's own array,
 * placed where the page placed them, in each piece's look; a rule under or
 * through consecutive pieces of one size, rise and colour.
 */
function drawParagraphWhereItStands(
  kernel: MraiNativeParagraphRewriteKernel,
  paragraph: Pick<ReadParagraph, "basis" | "blockMatrix" | "layout" | "shows" | "lineBaselines" | "unreadGlyphs">,
  lines: readonly SetLine[],
  piecesOf: (owner: number, owned: OwnedShow) => readonly StandingPiece[],
  /** Whether the text holds a blank between two of its characters. */
  blankBetween: (after: number, before: number) => boolean,
  setAgain?: ShowSetAgain,
): ReturnType<typeof setParagraph> {
  const { layout, shows } = paragraph;
  const [b0, b1, b2, b3] = paragraph.basis;
  let block = "q BT\n0 Tw 0 Tr\n";
  const state: { font: string | null; size: number | null; spacing: number | null; scale: number | null; rise: number | null;
    color: string | null; bold: boolean; strokeColor: string | null; word: number } =
    { font: null, size: null, spacing: null, scale: null, rise: null, color: null, bold: false, strokeColor: null, word: 0 };
  // (The text state a show is drawn under, each part said where it changes: setParagraph's.)
  const applyState = (style: CharacterStyle, word: number) => {
    if (state.word !== word) {
      block += `${formatNumber(word)} Tw\n`;
      state.word = word;
    }
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
  const underlineRules: string[] = [];
  const strikeRules: string[] = [];
  const linkRules: string[] = [];
  const { refuseLeanReadOtherwise, blankDrawn } = leanReadBackGuard(kernel, paragraph);
  for (let line = 0; line < layout.lineCount; line += 1) {
    const v = paragraph.lineBaselines[line];
    // The rules, as setParagraph draws a line's: one rule under consecutive
    // stretches of one size, rise and colour, from where the first is inked
    // to where the last is.
    type Ruled = Readonly<{ size: number; rise: number; color: PdfNativeDeviceColor; linkRule: ParagraphLinkRule | null }>;
    const reach: Record<"underline" | "strikethrough", { start: number; end: number; style: Ruled } | null> =
      { underline: null, strikethrough: null };
    const ruleKey = (style: Ruled) => JSON.stringify([style.size, style.rise, style.color]);
    let linkReach: { start: number; end: number; style: Ruled } | null = null;
    const linkKey = (style: Ruled) => JSON.stringify([style.size, style.rise, style.linkRule]);
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
        const across = v + open.style.rise;
        (kind === "underline" ? underlineRules : strikeRules).push(LATIN1.decode(pdfNativeTextRuleBlock({
          textMatrix: [b0, b1, b2, b3, open.start * b0 + across * b2, open.start * b1 + across * b3],
          fontSize: open.style.size, advance: open.end - open.start, color: open.style.color,
        }, kind)));
      }
      reach[kind] = null;
    };
    const ruleRun = (look: CharacterLook, style: Ruled, inkStart: number | null, inkEnd: number) => {
      for (const kind of ["underline", "strikethrough"] as const) {
        const open = reach[kind];
        if (!look[kind] || (open && ruleKey(open.style) !== ruleKey(style))) closeRule(kind);
        if (!look[kind] || inkStart === null || !(inkEnd > inkStart)) continue;
        // (Whichever way the page's shows follow one another along the line.)
        const standing = reach[kind];
        reach[kind] = standing
          ? { start: Math.min(standing.start, inkStart), end: Math.max(standing.end, inkEnd), style: standing.style }
          : { start: inkStart, end: inkEnd, style };
      }
      if (!style.linkRule || (linkReach && linkKey(linkReach.style) !== linkKey(style))) closeLinkRule();
      if (style.linkRule && inkStart !== null && inkEnd > inkStart) {
        if (linkReach) linkReach = { start: Math.min(linkReach.start, inkStart), end: Math.max(linkReach.end, inkEnd), style: linkReach.style };
        else linkReach = { start: inkStart, end: inkEnd, style };
      }
    };
    const twins: { place: string; style: CharacterStyle; word: number; array: string }[] = [];
    // The stretch drawn before on this line: where its ink ends, how it leans
    // (and leaned on the page), its size, and its last character.
    let before: Readonly<{ end: number; slant: number; changed: boolean; size: number; last: number }> | null = null;
    for (const [owner, owned] of shows.entries()) {
      if (owned.line !== line) continue;
      const { show } = owned;
      // (A show set again in another face, its glyphs placed anew.)
      const again = setAgain?.(owner, owned);
      if (again) {
        for (const run of again) {
          block += run.place;
          applyState(run.style, run.word);
          block += `[${run.array}] TJ\n`;
          allCodes.push(...run.codes);
          fontsUsed.add(run.style.fontName);
        }
        continue;
      }
      for (const { from: start, to: stop, look, ink, characters: drawnCharacters } of piecesOf(owner, owned)) {
        const piece = sourceShowPiece(kernel, paragraph, owned, start, stop);
        if (!piece) {
          if (ink) refuse(MRAI_NATIVE_PARAGRAPH_FORMAT_IN_PLACE_REFUSAL);
          continue;
        }
        const slant = look?.slant ?? owned.slant;
        // A stretch newly set in a leaned frame after text of its line: the
        // reader ends its item where the frame changes, whatever white space
        // the page draws there, and the editor's boxes read the two on from
        // one another by the gap between their ink alone -- a space where it
        // is a tracking's width or more and no wider than a folded space
        // (native-paragraph-edit-boxes), nothing otherwise. Where that is not
        // what the text holds there, the line would open again as other
        // words: «INTERNATIONALd/b/a» for a part leaned after a justified
        // line's wide gap. Refused as a write that reads back otherwise, which
        // a caller answers with the part in an italic face.
        if (ink && drawnCharacters) {
          if (before && slant !== 0 && slant !== before.slant && (slant !== owned.slant || before.changed) &&
            drawnCharacters[0] > before.last) {
            const gap = ink[0] - before.end;
            const reads = gap >= PDFJS_TRACKING_SPACE_FACTOR * before.size && gap <= READER_FOLDED_SPACE_MAX_EM * before.size;
            if (reads !== blankBetween(before.last, drawnCharacters[0])) refuse(MRAI_NATIVE_PARAGRAPH_READ_BACK_REFUSAL);
          }
          before = { end: ink[1], slant, changed: slant !== owned.slant, size: owned.size, last: drawnCharacters[1] };
        }
        if (ink && ink[1] > ink[0]) refuseLeanReadOtherwise(owned, slant, ink, v + owned.rise);
        else if (!ink) blankDrawn();
        const style: CharacterStyle = Object.freeze({ ...piece.style, slant, color: look?.color ?? owned.color,
          syntheticBold: look?.bold ?? owned.stroked });
        // A raised glyph stands along the leaned up axis: placed back by what
        // the new lean moves it, it stands where it stood.
        const place = (du: number, dv: number) => piece.place(slant, du - (slant - owned.slant) * piece.style.rise, dv);
        block += place(0, 0);
        applyState(style, piece.word);
        block += `[${piece.array}] TJ\n`;
        allCodes.push(...piece.codes);
        fontsUsed.add(show.fontName);
        const twin = look ? look.twin : owned.twin;
        if (twin) twins.push({ place: place(twin[0], twin[1]), style, word: piece.word, array: piece.array });
        // (A glyph of no width the reader reads as a blank -- a mark over a
        // letter -- is neither ruled nor an end of the rule under its line.)
        if (look && !(owned.readsBlank && !(ink && ink[1] > ink[0]))) {
          ruleRun(look, { size: owned.size, rise: owned.ruleRise ?? owned.rise, color: style.color, linkRule: look.linkRule },
            ink ? ink[0] : null, ink ? ink[1] : 0);
        }
      }
    }
    closeRule("underline");
    closeRule("strikethrough");
    closeLinkRule();
    // The fake-bold pass, under the empty /ActualText span the source marked it with.
    if (twins.length) {
      block += "/Span << /ActualText <> >> BDC\n";
      for (const twin of twins) {
        block += twin.place;
        applyState(twin.style, twin.word);
        block += `[${twin.array}] TJ\n`;
      }
      block += "EMC\n";
    }
  }
  block += "ET Q";
  const after = [...strikeRules, ...linkRules];
  const ruledBlock = underlineRules.length || after.length
    ? `q${underlineRules.length ? underlineRules.join("") : " "}${block.slice(2, -2)}${after.join("")}${after.length ? "" : " "}Q`
    : block;
  const baselines = paragraph.lineBaselines.slice(0, layout.lineCount);
  const steppedEvenly = baselines.every((baseline, index) =>
    Math.abs(baseline - (layout.firstBaseline - index * layout.lineStep)) <= 1e-9);
  return Object.freeze({ bytes: asciiBytes(ruledBlock), lines: Object.freeze(lines), codes: Object.freeze(allCodes),
    advance: lines.reduce((sum, line) => sum + line.inkWidth, 0), layout, notices: Object.freeze([]),
    fontsUsed: fontsUsed as ReadonlySet<string>, substituteCharacters: new Set<string>() as ReadonlySet<string>,
    sameBreaks: true, carried: Object.freeze([]),
    ...(steppedEvenly ? {} : { lineBaselines: Object.freeze(baselines.map((baseline) => Math.round(baseline * 1e6) / 1e6)) }) });
}

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
  /** The characters the block draws in the substitute face: those of the
   * lines it sets, not of the lines it draws again as the page drew them. */
  substituteCharacters: ReadonlySet<string>;
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
      // (A face of pictures -- colour emoji -- has no weight to take: the
      // weight is the words' around it, as a word processor sets bold text
      // that holds an emoji. Drawn under the stroke with them, a picture is
      // drawn as it was: the mode does not reach a Type3 procedure.)
      return Boolean(owner && type3Show(owner)) && ((format.syntheticBold === true && !owner.font.type3Pictures) ||
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
  // Formats that move no glyph of text that stays: drawn where it stands.
  if (formats && !substituteWhole) {
    const standing = setParagraphWhereItStands(kernel, paragraph, typedLines, formats);
    if (standing) return standing;
  }
  // Set by the measure from here on: what the read left to it is refused now.
  if (paragraph.measureRefusal !== undefined) refuse(paragraph.measureRefusal);
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
  // What the edit leaves alone stays as the page draws it. The lines above the
  // first word the edit touches, and what stands before that word on its own
  // line, are drawn again show by show: each where it stood, with the array it
  // was drawn with (its kerning, a fraction's numerator over its denominator).
  // This tier sets the text from that word on, and stops where its lines fall
  // back onto the source's (keptLines). Set whole, a paragraph typed onto at
  // its end came back with every line spaced again and the formula on its
  // third line taken apart, its radical 9 pt from the bar drawn over it
  
  // size or a line step asked of the paragraph sets every line of it.
  const lineOfCharacter = (index: number) => paragraph.shows[characterShowAt(index)].line;
  const anchor = ((): Readonly<{
    /** The first character this tier sets, and the source line it is on. */
    at: number;
    line: number;
    /** Where it stands along the line, when it does not start its line. */
    origin: number | null;
    /** The show it starts inside, and the code of that show it is drawn by:
     * the show's glyphs before it stay. */
    cut: Readonly<{ owner: number; code: number }> | null;
  }> | null => {
    if (formats || substituteWhole || paragraph.frameRequested || paragraph.sourceLineStep !== undefined ||
      paragraph.sizeScale !== undefined || opening !== openingKept || original.length === 0) return null;
    for (let index = 1; index < original.length; index += 1) {
      if (lineOfCharacter(index) < lineOfCharacter(index - 1)) return null;
    }
    const inkOf = (index: number) => paragraph.characterInk.get(index + opening);
    // A rule is drawn again under the characters it ruled, by where they are
    // inked: every character's place is then known.
    if ((paragraph.rules.length || paragraph.shows.some((show) => show.linkRule)) &&
      original.some((character, index) => !isBlank(character) && !inkOf(index))) return null;
    // The character the edit touches first: the one typed against, or the
    // first one changed.
    let touched = prefix > 0 && !isBlank(original[prefix - 1]) ? prefix - 1 : prefix;
    // Words typed at the end of a line, a break after them where the next
    // source line starts (the break the editor's rows put there, or one
    // typed): they are typed against that line's last word, not against the
    // line below, which stays as it is.
    // (Read from the break, not from the common prefix: words typed that
    // begin as the next line does -- "ab" before "avaswani" -- share a
    // letter with it.)
    const grown = replacement.length - original.length;
    const lineStartAbove = grown > 0 ? forced.map((at) => at - grown).find((start) =>
      start > 0 && start < original.length && lineOfCharacter(start) > lineOfCharacter(start - 1) &&
      original.slice(start).every((character, index) => character === replacement[start + grown + index]) &&
      original.slice(0, start - 1).every((character, index) => character === replacement[index])) : undefined;
    if (lineStartAbove !== undefined) {
      touched = lineStartAbove - 1;
      while (touched > 0 && isBlank(original[touched])) touched -= 1;
    }
    while (touched < original.length && isBlank(original[touched])) touched += 1;
    if (touched >= original.length) {
      touched = original.length - 1;
      while (touched > 0 && isBlank(original[touched])) touched -= 1;
    }
    const line = lineOfCharacter(touched);
    let lineStart = touched;
    while (lineStart > 0 && lineOfCharacter(lineStart - 1) === line) lineStart -= 1;
    // The word it is in, as far back as one show draws it: what another show
    // draws before it stays (the fraction a sentence ends with, before its
    // full stop: set again with the stop, its radical moved 9 pt).
    let word = touched;
    while (word > lineStart && !isBlank(original[word - 1]) &&
      characterShowAt(word - 1) === characterShowAt(touched)) word -= 1;
    const whole = Object.freeze({ at: lineStart, line, origin: null, cut: null });
    // A line set centred or flush right moves as a whole when its width does.
    if (word === lineStart || paragraph.layout.alignment === "center" || paragraph.layout.alignment === "right") return whole;
    const ink = inkOf(word);
    if (!ink) return whole;
    const owner = characterShowAt(word);
    const before = new Set<number>();
    for (let index = lineStart; index < word; index += 1) before.add(characterShowAt(index));
    // No show draws on both sides of the word's start but the word's own, and
    // nothing kept stands past where the word starts.
    for (let index = word; index < original.length && lineOfCharacter(index) === line; index += 1) {
      if (characterShowAt(index) !== owner && before.has(characterShowAt(index))) return whole;
    }
    for (const kept of before) {
      if (kept === owner) continue;
      const inked = paragraph.shows[kept].inked;
      if (inked && inked.end > ink[0] + LAYOUT_TOLERANCE) return whole;
    }
    if (!before.has(owner)) return Object.freeze({ at: word, line, origin: ink[0], cut: null });
    // The word starts inside a show: its glyphs and the text's characters go
    // one to one (characterInk), so the glyph the word starts at is known.
    const own = paragraph.shows[owner];
    const unit = own.font.codeUnitBytes;
    const drawn = showGlyphs(kernel, own.show, own.font).flatMap((glyph, index) =>
      !/^\s*$/u.test(glyph.text ?? "x") &&
      (!own.partial || (index * unit >= own.partial.codeStart && index * unit < own.partial.codeEnd)) ? [index] : []);
    let rank = 0;
    for (let index = lineStart; index < word; index += 1) {
      if (characterShowAt(index) === owner && !isBlank(original[index])) rank += 1;
    }
    if (rank < 1) return Object.freeze({ at: word, line, origin: ink[0], cut: null });
    if (drawn.length !== own.ink.length || rank >= drawn.length ||
      Math.abs(own.ink[rank][0] - ink[0]) > 1e-9 || own.ink[rank - 1][1] > ink[0] + LAYOUT_TOLERANCE) return whole;
    return Object.freeze({ at: word, line, origin: ink[0], cut: Object.freeze({ owner, code: drawn[rank] * unit }) });
  })();
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
  // (A gap in what stays as the page draws it stays with it.)
  const closedGaps = paragraph.unreadGaps.flatMap((gap, index) =>
    gapClosed(gap) && !(anchor && gap - opening <= anchor.at)
      ? [{ boundary: gap - opening, width: paragraph.unreadGapWidths[index] }] : []);
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
    // (A face of a character's own family sets all it is asked of, or the
    // style is drawn on the characters' own face: formatsInFacesThatSetAll.)
    const askedOfEach = replacement.map((_, index) => formatOf![index + openingKept]);
    const asked = formatsInFacesThatSetAll(kernel, replacement, askedOfEach, (index) => styles[index]);
    // (A style of the text's own family asked of all of it is all of it, or not written through that face.)
    if (replacementText === paragraph.selection.item.str) {
      refuseStyleFaceThatLeavesTextUnstyled(kernel, replacement, askedOfEach, asked, (index) => styles[index]);
    }
    for (let index = 0; index < replacement.length; index += 1) {
      const format = asked[index];
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
    // The page's edge along the lines (the block's own u, as checkExtraLines
    // reads other shows into it): the furthest of its corners that way, as
    // readParagraph reads the room beside the paragraph (its freeRight) --
    // the right edge of an upright line, the edge a turned line runs to.
    const toBlock = invertMatrix(paragraph.blockMatrix);
    const along = Math.max(...[[bounds.left, bounds.bottom], [bounds.right, bounds.bottom],
      [bounds.left, bounds.top], [bounds.right, bounds.top]].map(([px, py]) => {
      const [x, y] = applyMatrix(toBlock, px, py);
      return x * paragraph.basis[0] + y * paragraph.basis[1];
    }));
    return Number.isFinite(along) ? along : Infinity;
  })();
  // So does a line the paragraph has, set again. Its frame can stand past
  // the page: the frame is widened to hold the paragraph's lines by this
  // tier's own measure of them (paragraphFrameHoldingItsLines), and that
  // measure can be far wider than the page draws a line -- a contents line's
  // leader, dots under a character spacing the reader spells with a blank
  // after each ("Introduction . . . . 4", 255 pt on the page, 425 pt
  // measured, in a column that ends 42 pt from the paper's edge). In that
  // frame a word typed after the page number "fitted" on its line, the
  // check of a line against the page (below) is only made of a line wider
  // than its frame, and the word was written off the paper: saved, in the
  
  
  // word goes down a line as a text block's does (Acrobat, WPS) and a line
  // that still cannot stay on the page is refused (below) -- further only
  // as far as the paragraph's own lines already run past it, which stay
  // where the page draws them.
  const ownRight = Math.max(...layout.lineExtents.map((extent) => extent.right));
  const frameRightOf = (line: number) => Math.min(layout.frameRight,
    line >= layout.lineCount ? pageRight : Math.max(pageRight, ownRight));
  const available = (line: number) => layout.alignment === "center" || layout.alignment === "right"
    ? frameRightOf(line) - layout.frameLeft
    : frameRightOf(line) - lineLeftOf(layout, line);
  // The lines that stay as the page draws them (the anchor above): each
  // source line's first character in the text, and a source line as a set
  // line. None where the edit changed the paragraph's indents (a list marker
  // given or taken off), or the new text draws nothing from the anchor on.
  const kept = anchor !== null && layout === paragraph.layout &&
    replacement.slice(anchor.at).some((character) => !isBlank(character)) ? (() => {
      const lineFirst: number[] = [];
      for (let index = 0; index < original.length; index += 1) {
        const line = lineOfCharacter(index);
        if (lineFirst[line] === undefined) lineFirst[line] = index;
      }
      if (lineFirst[0] !== 0 || lineFirst.length !== layout.lineCount) return null;
      for (let line = 0; line < layout.lineCount; line += 1) if (lineFirst[line] === undefined) return null;
      const lineEnd = (line: number) => lineFirst[line + 1] ?? original.length;
      const sourceLine = (line: number, shift: number): SetLine => {
        let inkEnd = lineEnd(line);
        while (inkEnd > lineFirst[line] && isBlank(original[inkEnd - 1])) inkEnd -= 1;
        const extent = layout.lineExtents[line];
        return Object.freeze({ start: lineFirst[line] + shift, end: lineEnd(line) + shift, inkEnd: inkEnd + shift,
          inkWidth: extent.right - extent.left });
      };
      return Object.freeze({ ...anchor, lineFirst, sourceLine });
    })() : null;
  // Where set lines fall back onto the source's: the first of them, from the
  // edit's own line on, that starts where a source line starts with nothing
  // but kept text from there to the paragraph's end. It and the lines after
  // it are that source line and the ones after it, as the page draws them.
  const keptTail = (set: readonly SetLine[]): Readonly<{ at: number; line: number }> | null => {
    if (!kept || suffix <= 0) return null;
    const moved = replacement.length - original.length;
    const unchanged = replacement.length - suffix;
    for (let index = Math.max(1, kept.line); index < set.length; index += 1) {
      let first = set[index].start;
      while (first < set[index].end && isBlank(replacement[first])) first += 1;
      if (first < unchanged || first < kept.at || set[index].keptUntil !== undefined) continue;
      const line = kept.lineFirst.indexOf(first - moved);
      // A paragraph run round a picture keeps each line's own edge: its lines
      // stay only where they are.
      if (line < 1 || (layout.lineLefts?.length && line !== index)) continue;
      return Object.freeze({ at: index, line });
    }
    return null;
  };
  const withKeptTail = (set: SetLine[]): SetLine[] => {
    const tail = keptTail(set);
    if (!kept || !tail) return set;
    const moved = replacement.length - original.length;
    return [...set.slice(0, tail.at),
      ...Array.from({ length: layout.lineCount - tail.line }, (_, index) => kept.sourceLine(tail.line + index, moved))];
  };
  // The paragraph's lines: the source's own above the anchor, this tier's
  // from it on (in the room its line has left beside what stays on it; a word
  // that no longer fits there starts the next line), the source's again
  // where they fall back onto them.
  const setLines = (
    styleSet: readonly CharacterStyle[],
    blankSet: readonly (Readonly<{ width: number }> | null)[],
    room: (line: number) => number,
    splitWords = false,
    kernSet?: readonly number[],
  ): SetLine[] => {
    if (!kept) return setParagraphLines(kernel, replacement, styleSet, blankSet, room, forced, splitWords, kernSet);
    const { at, line: own } = kept;
    const head = Array.from({ length: own }, (_, line) => kept.sourceLine(line, 0));
    const from = (start: number, first: number, taken: number, breakWords: boolean) => setParagraphLines(kernel,
      replacement.slice(start), styleSet.slice(start), blankSet.slice(start),
      (line) => room(first + line) - (line === 0 ? taken : 0),
      forced.filter((stop) => stop > start).map((stop) => stop - start), breakWords, kernSet?.slice(start))
      .map((line) => Object.freeze({ start: line.start + start, end: line.end + start, inkEnd: line.inkEnd + start,
        inkWidth: line.inkWidth }));
    const lineStart = kept.lineFirst[own];
    if (kept.origin === null) {
      // The edit's line set whole (a centred or right-aligned one, which moves
      // as its width changes) -- unless every character of it is the source's
      // and it ends where the source's ended: what was typed went down to the
      // next line, and the line is the page's own, drawn again as it stands.
      // Set anew it was centred by this tier's measure of it, and a label's
      // last row moved 0.8 pt under the row typed after it (IPCC SPM.1's
      
      const set = from(lineStart, own, 0, splitWords);
      const source = kept.sourceLine(own, 0);
      const stays = set.length > 1 && prefix >= source.inkEnd && set[0].start === source.start &&
        set[0].inkEnd === source.inkEnd;
      return withKeptTail([...head,
        ...(stays ? [Object.freeze({ ...source, end: set[0].end, keptUntil: set[0].end })] : [set[0]]), ...set.slice(1)]);
    }
    const taken = kept.origin - lineLeftOf(layout, own);
    const beside = from(at, own, taken, false);
    const fits = !(beside[0].inkWidth > room(own) - taken + LAYOUT_TOLERANCE);
    const set = fits && splitWords ? from(at, own, taken, true) : beside;
    // A justified line is filled to its frame by its gaps, and what stays of
    // it keeps the gaps it has: where what this tier sets after it cannot
    // fill the line -- its first word no longer fits there, or it is one
    // word that ends short of the frame -- the whole line is set again, as
    // its paragraph's other lines reach the frame. (Left a word short, the
    // line read as the end of a paragraph inside one.)
    const filled = layout.alignment === "distribute" || (layout.alignment === "justify" &&
      (fits ? set.length > 1 : true) && own < layout.lineCount - 1);
    if (filled && (!fits || (!replacement.slice(at, set[0].inkEnd).some(isBlank) &&
      taken + set[0].inkWidth < room(own) - LAYOUT_TOLERANCE))) {
      return withKeptTail([...head, ...from(lineStart, own, 0, splitWords)]);
    }
    if (!fits) {
      // The word the edit starts in no longer fits beside what stays on its
      // line: it starts the next, and the line is what stays of it.
      let inkEnd = at;
      while (inkEnd > lineStart && isBlank(original[inkEnd - 1])) inkEnd -= 1;
      return withKeptTail([...head, Object.freeze({ start: lineStart, end: at, inkEnd, inkWidth: taken, keptUntil: at }),
        ...from(at, own + 1, 0, splitWords)]);
    }
    return withKeptTail([...head, Object.freeze({ start: lineStart, end: set[0].end, inkEnd: set[0].inkEnd,
      inkWidth: taken + set[0].inkWidth, keptUntil: at }), ...set.slice(1)]);
  };
  /** Where this tier's own setting of line `index` of `set` starts: the
   * line's end where the whole line stays as the page draws it. */
  const setFromIn = (set: readonly SetLine[], tail: ReturnType<typeof keptTail>, index: number): number => {
    const line = set[index];
    if (!kept) return line.start;
    if (index < kept.line || (tail && index >= tail.at)) return line.end;
    return index === kept.line && line.keptUntil !== undefined ? Math.min(line.keptUntil, line.end) : line.start;
  };
  /** A line of `set` that stays whole as the page draws it. */
  const staysIn = (set: readonly SetLine[]) => {
    const tail = keptTail(set);
    return (index: number) => setFromIn(set, tail, index) >= set[index].end;
  };
  let lines = setLines(styles, blanks, available);
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
      const gapLines = setLines(styles, gapsKept, available);
      if (gapLines.length <= lines.length) {
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
        lines = gapLines;
      }
    }
  }
  // A glyph the source sets back over the one before it (readParagraph's
  // stacked: TeX's accent, then the letter under it) is set over it again,
  // whatever face and size the two are set in now. Where the pen goes back
  // past the middle of the glyph before, the two keep their middles as far
  // apart as the source set them, in shares of the size -- the widths are the
  // new face's, so the number is the one TeX itself sets the accent by in that
  // face (CMTI12's 400 is CMR12's 381), and the source's own where the face
  // and size stay. The glyph before the two keeps the letter where the source
  // set it after it (TeX kerns before the accent so that the letter stands
  // where it would with none): carried as it was, a narrower accent drew the
  // letter into the one before it. Any other step back is carried as it is.
  const stackKerns = new Map<number, number>();
  const setWidth = (index: number): number | null => {
    try {
      return setRun(kernel, replacement[index], styles[index]).width;
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
      return null;
    }
  };
  for (let index = 1; index < replacement.length; index += 1) {
    const source = keptSource[index];
    if (source < 1 || keptSource[index - 1] !== source - 1 || blanks[index] || blanks[index - 1]) continue;
    const stack = paragraph.stacked.get(inParagraphText(source));
    if (!stack || inParagraphText(source - 1) !== inParagraphText(source) - 1) continue;
    const em = styles[index].size * (styles[index].horizontalScale / 100);
    const beforeWidth = setWidth(index - 1), width = setWidth(index);
    if (!(em > 0) || beforeWidth === null || width === null) continue;
    const over = -stack.back >= stack.beforeWidth / 2;
    const kern = over
      ? Math.min((stack.back + (stack.beforeWidth + stack.width) / 2) * em - (beforeWidth + width) / 2, -beforeWidth / 2)
      : stack.back * em;
    stackKerns.set(index, kern);
    if (over && stack.lead !== null && index > 1 && !stackKerns.has(index - 1) && keptSource[index - 2] === source - 2 &&
      !blanks[index - 2] && inParagraphText(source - 2) === inParagraphText(source) - 2) {
      stackKerns.set(index - 1, (stack.lead + stack.beforeWidth + stack.back) * em - beforeWidth - kern);
    }
  }
  const kerns = replacement.map((_, index) => stackKerns.get(index) ?? 0);
  if (stackKerns.size) lines = setLines(styles, blanks, available, false, kerns);
  // The kerns between the characters the edit kept side by side stay too (a
  // TeX \\scriptspace before the stop after a square, a kerned pair): dropped, the
  // words after them moved a little on every edit. Unless that takes a line more.
  const pairKerns = replacement.map(() => 0);
  for (let index = 1; index < replacement.length; index += 1) {
    const source = keptSource[index];
    if (source < 1 || keptSource[index - 1] !== source - 1 || blanks[index] || blanks[index - 1] || stackKerns.has(index)) continue;
    const ink = paragraph.characterInk.get(inParagraphText(source)), before = paragraph.characterInk.get(inParagraphText(source - 1));
    const show = paragraph.shows[sourceOf(index)];
    if (!ink || !before || styles[index].size !== show.size || styles[index].fontName !== show.show.fontName) continue;
    if (paragraph.shows[characterShowAt(source - 1)].line !== paragraph.shows[characterShowAt(source)].line) continue;
    const kern = ink[0] - before[1];
    if (Math.abs(kern) >= 0.01 && Math.abs(kern) < PDFJS_TRACKING_SPACE_FACTOR * show.size) pairKerns[index] = kern;
  }
  if (pairKerns.some(Boolean)) {
    const all = kerns.map((kern, index) => kern || pairKerns[index]);
    const kerned = setLines(styles, blanks, available, false, all);
    if (kerned.length <= lines.length) {
      lines = kerned;
      all.forEach((kern, index) => { kerns[index] = kern; });
    }
  }
  // The room beside a flush-left paragraph (readParagraph's freeRight) keeps
  // a line that grew a hair from wrapping; it never pulls a word up onto the
  // line above. Set in the paragraph's own width (its widest line), an edit
  // that left fewer lines in the grown frame keeps that many: a table header
  // "Ballots / Completed" retyped "Votes" came back as one line run past its
  
  if (layout.alignment === "left" && !paragraph.frameRequested && layout.frameRight > ownRight + LAYOUT_TOLERANCE) {
    const own = setLines(styles, blanks, (line: number) => ownRight - lineLeftOf(layout, line), false, kerns);
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
      const again = setLines(restyled, restyledBlanks, available, false, kerns);
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
  const runsOnPastFrame = (set: readonly SetLine[], baselines = baselinesOf(set), stays = staysIn(set)) => set.every((line, index) => {
    if (stays(index) || !(line.inkWidth > available(index) + LAYOUT_TOLERANCE)) return true;
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
    const split = setLines(styles, blanks, available, true, kerns);
    const stays = staysIn(split);
    if (split.every((line, index) => stays(index) || !(line.inkWidth > available(index) + LAYOUT_TOLERANCE))) lines = split;
  }
  // Set lower for its larger characters only as far as the paragraph's last
  // line then stays clear of the text below it: the page cannot push the
  // paragraph under it down, as Word would, and a lowered last line was drawn
  
  // larger word reaches up into its own paragraph's line above.
  // Where the lines stand with lines kept as the page draws them: a kept line
  // on its own baseline, a line set again on the baseline the source gave the
  // line it is (the page spaces a line under a fraction further down than
  // the step, and this tier's measure of what reaches past the body text is
  // not the producer's), lower only by what its characters now reach further
  // than the source's did there; a line the paragraph did not have a step
  // under the one before; and the lines fallen back onto the source's as far
  // under the line before them as the source set them.
  const keptBaselines = (): number[] => {
    const source = paragraph.lineBaselines;
    const tail = keptTail(lines);
    const reachOf = (line: number) => lineReachOf(paragraph.shows.filter((show) => show.line === line && show.inked)
      .map((show) => ({ size: show.size, rise: show.rise })), layout.bodySize);
    const reach = lines.map((line) => lineReachOf(styles.slice(line.start, line.end), layout.bodySize));
    const baselines: number[] = [];
    for (let index = 0; index < lines.length; index += 1) {
      if (tail && index >= tail.at) {
        baselines.push(source[tail.line + index - tail.at] + (baselines[tail.at - 1] - source[tail.line - 1]));
        continue;
      }
      if (index === 0 || index <= kept!.line) {
        baselines.push(source[index]);
        continue;
      }
      const reached = reach[index].above + reach[index - 1].below;
      if (index < source.length) {
        const grown = reached - (reachOf(index).above + reachOf(index - 1).below);
        baselines.push(baselines[index - 1] - (source[index - 1] - source[index]) - leading * (grown > LAYOUT_TOLERANCE ? grown : 0));
      } else {
        baselines.push(baselines[index - 1] - layout.lineStep - leading * reached);
      }
    }
    return baselines;
  };
  const lineBaselines = kept ? keptBaselines() : (() => {
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
  const keptLineTail = keptTail(lines);
  const setFromOf = (index: number) => setFromIn(lines, keptLineTail, index);
  for (const [index, line] of lines.entries()) {
    if (setFromOf(index) < line.end && line.inkWidth > available(index) + LAYOUT_TOLERANCE) {
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
  
  
  // two side by side and this tier cannot set one over the other -- which
  // characters the show's glyphs are is not known, or a line breaks between
  // the two.
  if (formats) {
    if (paragraph.stackedUnread) refuse(MRAI_NATIVE_PARAGRAPH_STACKED_GLYPH_REFUSAL);
    const setAt = new Map<number, number>();
    keptSource.forEach((source, index) => { if (source >= 0) setAt.set(source, index); });
    for (let index = 1; index < replacement.length; index += 1) {
      const source = keptSource[index];
      if (source < 1 || setAt.get(source - 1) !== index - 1 || !paragraph.stacked.has(inParagraphText(source)) ||
        inParagraphText(source - 1) !== inParagraphText(source) - 1) continue;
      if (!stackKerns.has(index) || lines.some((line) => line.start === index)) {
        refuse(MRAI_NATIVE_PARAGRAPH_STACKED_GLYPH_REFUSAL);
      }
    }
  }
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
    scale: number | null; rise: number | null; color: string | null; bold: boolean; strokeColor: string | null;
    word: number } =
    { font: null, size: null, spacing: null, scale: null, rise: null, color: null, bold: false, strokeColor: null,
      word: 0 };
  // The text state a run is drawn under: each part set where it changes. (A
  // word spacing only under a show kept as the page draws it.)
  const applyState = (style: CharacterStyle, word = 0) => {
    if (state.word !== word) {
      block += `${formatNumber(word)} Tw\n`;
      state.word = word;
    }
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
  const substituteCharacters = new Set<string>();
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
  // A show kept as the page draws it (codes `from` to `to` of it), its line
  // moved `drop` across the lines: as the array it was drawn with (its strings
  // and the numbers between them), from where its first glyph of them stood,
  // under the state the source drew it in -- said the way this block says its
  // own runs (the size in the font operator, along the paragraph's basis), so
  // the reader reads on from a kept show into a run set after it as it read
  // the line before: said in the source's own terms, a line whose text matrix
  // holds its size came back as three items, its last word a font size apart
  // from the words before it (a slide set sideways). What the reader measures
  // by stays the source's: a text matrix that stretches the line stretches it
  // here, not a Tz (the reader measures a gap against the matrix's scale, and
  // a line at 1.02 drawn again at 102 Tz read as two items at a gap of
  // 0.6 em: Attention p1's abstract), and a show placed off its line stays
  // placed there, not on the line under a rise (the reader takes a glyph
  // moved by a rise as still on its line: a subscript drawn again by a rise
  // read with a space after it that the page never read, p4's "dk.").
  const keptShow = (owned: OwnedShow, from: number, to: number, drop: number) => {
    const piece = sourceShowPiece(kernel, paragraph, owned, from, to);
    if (!piece) return null;
    allCodes.push(...piece.codes);
    fontsUsed.add(owned.show.fontName);
    return Object.freeze({ style: piece.style, word: piece.word, array: piece.array,
      place: (du: number, dv: number) => piece.place(owned.slant, du, dv + drop) });
  };
  // The characters each owned show draws, in the text.
  const showCharacters = new Map<number, number[]>();
  if (kept) {
    for (let index = 0; index < original.length; index += 1) {
      const own = showCharacters.get(characterShowAt(index));
      if (own) own.push(index);
      else showCharacters.set(characterShowAt(index), [index]);
    }
  }
  // The parts of the block that draw kept shows: no part of what this tier
  // set, and left out of what is read back.
  const keptRanges: [number, number][] = [];
  // Where each line inks the page along the paragraph's lines: its runs of
  // inked characters, the blanks between them left out.
  const lineSpans: [number, number][][] = [];
  for (const [lineIndex, line] of lines.entries()) {
    const width = line.inkWidth;
    const u = layout.alignment === "right" ? layout.frameRight - width
      : layout.alignment === "center" ? (layout.frameLeft + layout.frameRight) / 2 - width / 2
        : lineLeftOf(layout, lineIndex);
    const v = lineBaselines[lineIndex];
    // What of the line stays as the page draws it: the whole of a line above
    // the anchor's or fallen back onto the source's, the shows before the
    // anchor on its own line.
    const setFrom = setFromOf(lineIndex);
    const stays = setFrom >= line.end;
    const sourceLine = !kept || setFrom <= line.start ? null
      : keptLineTail && lineIndex >= keptLineTail.at ? keptLineTail.line + lineIndex - keptLineTail.at : lineIndex;
    if (kept && sourceLine !== null) {
      const sourceBaseline = paragraph.lineBaselines[sourceLine];
      // (The anchor's own line whole, where it stays whole: setLines.)
      const whole = sourceLine !== kept.line || (keptLineTail !== null && lineIndex >= keptLineTail.at) ||
        (kept.origin === null && stays);
      const drawn: NonNullable<ReturnType<typeof keptShow>>[] = [];
      const twins: { drawn: NonNullable<ReturnType<typeof keptShow>>; twin: readonly [number, number] }[] = [];
      for (const [owner, owned] of paragraph.shows.entries()) {
        if (owned.line !== sourceLine) continue;
        const from = owned.partial?.codeStart ?? 0;
        let to = owned.partial?.codeEnd ?? owned.show.codes.length;
        if (!whole) {
          const own = showCharacters.get(owner);
          if (kept.cut?.owner === owner) to = kept.cut.code;
          else if (own ? own.some((index) => index >= kept.at)
            : !(owned.inked && kept.origin !== null && owned.inked.end <= kept.origin + LAYOUT_TOLERANCE)) continue;
        }
        const show = keptShow(owned, from, to, v - sourceBaseline);
        if (!show) continue;
        drawn.push(show);
        if (owned.twin) twins.push({ drawn: show, twin: owned.twin });
      }
      if (drawn.length) {
        const rangeStart = block.length;
        for (const show of drawn) {
          block += show.place(0, 0);
          applyState(show.style, show.word);
          block += `[${show.array}] TJ\n`;
        }
        if (twins.length) {
          block += "/Span << /ActualText <> >> BDC\n";
          for (const { drawn: show, twin } of twins) {
            block += show.place(twin[0], twin[1]);
            applyState(show.style, show.word);
            block += `[${show.array}] TJ\n`;
          }
          block += "EMC\n";
        }
        // The state they were drawn under ends with them: what this tier sets
        // next says its own, whole.
        if (state.bold) block += "0 Tr\n";
        if (state.word !== 0) block += "0 Tw\n";
        Object.assign(state, { font: null, size: null, spacing: null, scale: null, rise: null, color: null, bold: false,
          strokeColor: null, word: 0 });
        keptRanges.push([rangeStart, block.length]);
      }
    }
    // A rise in a leaned frame moves the glyph along the leaned up axis, a
    // slant times the rise to the right of the pen: a run both leaned and
    // raised is placed back by that much, so its glyphs stand on from the
    // ones before. Placed at the pen, an italic superscript begun inside a
    // word stood 1.3 pt off it, and the next edit of the paragraph read
    
    const leanShift = (style: CharacterStyle) => style.slant * style.rise;
    const setStyle = styles[Math.min(setFrom, styles.length - 1)];
    let lineSlant = setStyle.slant;
    let lineLean = leanShift(setStyle);
    // How far the pen has moved along this line: past what stays on it.
    let pen = setFrom > line.start && kept && kept.origin !== null ? kept.origin - u : 0;
    if (!stays) block += placeAt(u + pen - lineLean, v, lineSlant);
    const spans: [number, number][] = [];
    lineSpans.push(spans);
    advance += width;
    // The runs the line draws: where each starts along it, its style and
    // its array, for the fake-bold pass drawn again after the line.
    const runs: { penStart: number; style: CharacterStyle; array: string }[] = [];
    // Justified lines before the last stretch every blank between their
    // first and last inked characters by the same amount.
    let firstInk = setFrom;
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
    // The rules a run's characters are drawn with, from where the run's ink
    // starts along the line to where it ends (null: it inks nothing).
    const ruleRun = (style: CharacterStyle, inkStart: number | null, inkEnd: number) => {
      for (const kind of ["underline", "strikethrough"] as const) {
        const open = reach[kind];
        if (!style[kind] || (open && ruleKey(open.style) !== ruleKey(style))) closeRule(kind);
        if (!style[kind] || inkStart === null || !(inkEnd > inkStart)) continue;
        if (reach[kind]) reach[kind]!.end = inkEnd;
        else reach[kind] = { start: inkStart, end: inkEnd, style };
      }
      if (!style.linkRule || (linkReach && linkKey(linkReach.style) !== linkKey(style))) closeLinkRule();
      if (style.linkRule && inkStart !== null && inkEnd > inkStart) {
        if (linkReach) linkReach.end = inkEnd;
        else linkReach = { start: inkStart, end: inkEnd, style };
      }
    };
    // What stays on the line as the page draws it keeps its rules, under where
    // the source inks its characters.
    for (let index = line.start; index < setFrom;) {
      const style = styles[index];
      const key = styleKey(style);
      let stop = index + 1;
      while (stop < setFrom && styleKey(styles[stop]) === key) stop += 1;
      let inkStart: number | null = null;
      let inkEnd = 0;
      if (style.underline || style.strikethrough || style.linkRule) {
        for (let at = index; at < stop; at += 1) {
          const ink = keptSource[at] >= 0 ? paragraph.characterInk.get(inParagraphText(keptSource[at])) : undefined;
          if (!ink) continue;
          if (inkStart === null) inkStart = ink[0];
          inkEnd = ink[1];
        }
      }
      ruleRun(style, inkStart, inkEnd);
      index = stop;
    }
    let cursor = setFrom;
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
      if (substitute && style.fontName === substitute.resourceName) {
        for (let at = cursor; at < stop; at += 1) {
          if (!isBlank(replacement[at])) substituteCharacters.add(replacement[at]);
        }
      }
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
        const kern = index > setFrom ? kerns[index] : 0;
        if (kern) {
          pushNumber(-kern * perUnit);
          pen += kern;
        }
        let end = index + 1;
        while (end < stop && !blanks[end] && !kerns[end]) end += 1;
        const set = setRun(kernel, replacement.slice(index, end).join(""), style);
        // (A glyph set back over the one before it may end before that one does.)
        runInkEnd = runInkStart === null ? pen + set.width : Math.max(runInkEnd, pen + set.width);
        if (runInkStart === null) runInkStart = pen;
        pendingCodes.push(...set.codes);
        allCodes.push(...set.codes);
        if (spans.length && index > setFrom && !blanks[index - 1]) {
          spans[spans.length - 1][1] = Math.max(spans[spans.length - 1][1], u + pen + set.width);
        } else spans.push([u + pen, u + pen + set.width]);
        pen += set.width;
        index = end;
      }
      flush();
      const array = elements.join(" ");
      block += `[${array}] TJ\n`;
      runs.push({ penStart, style, array });
      ruleRun(style, runInkStart === null ? null : u + runInkStart, u + runInkEnd);
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
  // (What stays as the page draws it reads as the page reads it: the read-back
  // is of what this tier set.)
  const readBack = readBlockText(kernel, paragraph, keptRanges.reduceRight((text, [from, to]) =>
    text.slice(0, from) + text.slice(to), block));
  // Wrapped, the lines read joined as the reader joins a paragraph's lines:
  // the blank a line breaks at is no character of either line, and between a
  // long CJK line and the next none is read (lineBoundaryNeedsSpace).
  // A typed break is one line boundary more, joined the same way: by the
  // lines as set (joined as typed segments, a segment holding a space was
  // never a long identifier, and 「…/final」 + 「https://…」 expected a space
  // the reader does not read).
  const expectedBack = readerSpelling(kept
    ? joinLines(lines.flatMap((line, index) => setFromOf(index) < line.end
      ? [replacement.slice(setFromOf(index), line.end).join("").trim()] : []))
    : lines.length > 1
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
    // (Nothing moved in what stays as the page draws it.)
    const from = setFromOf(index);
    if (from >= line.end) continue;
    const width = line.inkWidth;
    const left = from > line.start && kept && kept.origin !== null ? kept.origin
      : layout.alignment === "right" ? layout.frameRight - width
        : layout.alignment === "center" ? (layout.frameLeft + layout.frameRight) / 2 - width / 2
          : lineLeftOf(layout, index);
    for (let at = from + 1; at + 1 < line.end; at += 1) {
      const source = keptSource[at];
      if (!isBlank(replacement[at]) || source < 1 || keptSource[at + 1] !== source + 1 || !isBlank(original[source])) continue;
      const before = paragraph.shows[characterShowAt(source - 1)], next = paragraph.shows[characterShowAt(source + 1)];
      if (before === next || before.line !== next.line || !before.inked || !next.inked) continue;
      const dx = left + measureRange(kernel, replacement, styles, blanks, from, at + 1, kerns) - next.inked.start;
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
    layout, notices: Object.freeze(notices), fontsUsed: fontsUsed as ReadonlySet<string>,
    substituteCharacters: substituteCharacters as ReadonlySet<string>, sameBreaks, carried: paragraphCarriedInGaps(kernel, paragraph, movedGaps),
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
  // (And text that runs under the new line from a show starting outside the
  // frame: the next row of a table TeX sets as one show from its first
  // column. Read by where shows start alone, a figure wrapped down into the
  // row below its cell was drawn over that row's figure and no one was told.)
  const inked = otherInkedExtents(kernel, paragraph);
  for (let index = layout.lineCount; index < lines.length; index += 1) {
    const line = lines[index];
    const size = Math.max(...styles.slice(line.start, line.end).map((style) => style.size));
    const baseline = baselines[index];
    const bottom = baseline - 0.25 * size, top = baseline + size;
    // The new line as it is set (its own ink across the line), not the
    // frame: a frame the source's own lines run past the page with (a URL
    // past the right margin) would refuse every new line.
    const width = line.inkWidth;
    const left = layout.alignment === "right" ? layout.frameRight - width
      : layout.alignment === "center" ? (layout.frameLeft + layout.frameRight) / 2 - width / 2
        : lineLeftOf(layout, index);
    // Other text standing where the new line goes is drawn over, as a text
    // box grown in Word or Acrobat draws over what lies below it: the write
    // goes through and the user is told (docs/pdf/edit-gates.md).
    if ((others.some(([u, v]) => v >= bottom && v <= top && u >= layout.frameLeft - 0.5 && u <= layout.frameRight + 0.5) ||
      // (Text whose baseline lies under the new line's band still reaches
      // up into it with its own letters: a paragraph's ninth line set a
      // third of a line step above the line that followed the paragraph was
      // drawn through that line's capitals and nothing was said.)
      inked.some((other) => other.across <= top && other.across + OTHER_TEXT_ASCENT_RATIO * other.size >= bottom &&
        (other.end ?? Infinity) > left + LAYOUT_TOLERANCE && other.start < left + width - LAYOUT_TOLERANCE)) &&
      !notices.some((notice) => notice.code === "new-line-over-other-text")) {
      notices.push(Object.freeze({ code: "new-line-over-other-text" }));
    }
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
/** How far above its baseline text standing under a new line is taken to
 * draw, of its size: its capitals and ascenders. Lines set solid (a step of
 * the size) stay clear of each other by this. */
const OTHER_TEXT_ASCENT_RATIO = 0.7;
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
): Readonly<{ start: number; end: number | null; across: number; size: number; show: MraiNativeParagraphRewriteShow }>[] {
  const { basis, blockMatrix } = paragraph;
  const toBlock = invertMatrix(blockMatrix);
  const owned = new Set(paragraph.shows.map((show) => `${show.start}:${show.end}`));
  // The show's font size as it stands across the paragraph's lines.
  const sizeAcross = (show: MraiNativeParagraphRewriteShow, toLines: ReturnType<typeof composeMatrix>) => {
    const [x0, y0] = applyMatrix(toLines, 0, 0), [x1, y1] = applyMatrix(toLines, 0, show.fontSize);
    return Math.abs((x1 - x0) * basis[2] + (y1 - y0) * basis[3]);
  };
  // What a show the paragraph owns part of draws outside that part is other
  // text too (the next column of a table row TeX sets as one show: left out,
  // a figure typed longer had the room up to the page's edge, across the
  // columns beside it).
  const outside: Readonly<{ start: number; end: number | null; across: number; size: number; show: MraiNativeParagraphRewriteShow }>[] = [];
  for (const own of paragraph.shows) {
    if (!own.partial) continue;
    const { show, font } = own;
    let glyphs: ReturnType<typeof showGlyphs>;
    try {
      glyphs = showGlyphs(kernel, show, font);
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
      continue;
    }
    const unit = font.codeUnitBytes;
    const toLines = composeMatrix(composeMatrix(show.textMatrix, show.graphicsMatrix), toBlock);
    const placed = (along: number) => {
      const [x, y] = applyMatrix(toLines, along, show.textRise);
      return [x * basis[0] + y * basis[1], x * basis[2] + y * basis[3]] as const;
    };
    const scale = show.horizontalScale / 100;
    for (const [from, to] of [[0, own.partial.codeStart / unit], [own.partial.codeEnd / unit, glyphs.length]] as const) {
      const inked = glyphs.slice(from, to).filter((glyph) => !/^\s*$/u.test(glyph.text ?? "x"));
      if (!inked.length) continue;
      const [start, across] = placed(inked[0].start * scale), [end] = placed(inked.at(-1)!.end * scale);
      outside.push(Object.freeze({ start: Math.min(start, end), end: Math.max(start, end), across, size: sizeAcross(show, toLines), show }));
    }
  }
  return [...outside, ...paragraph.allShows
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
        return Object.freeze({ start, end: null, across, size: sizeAcross(show, toLines), show });
      }
      const scale = show.horizontalScale / 100;
      const [start, across] = placed(run.start * scale), [end] = placed(run.end * scale);
      return Object.freeze({ start: Math.min(start, end), end: Math.max(start, end), across, size: sizeAcross(show, toLines), show });
    })];
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

/**
 * This writer's rule blocks beside a paragraph's text objects that stand on
 * the paragraph's own lines. A text object may draw more than one paragraph
 * (a TeX page is one, or a few), and a rule of each stands beside it: a rule
 * under a line of another paragraph is none of this one's, and is neither
 * taken away with it nor a reason to own none of its own (paragraphOwnRules
 * owns all of them or none). A rule is on a line where its near edge runs
 * within the line's height of its baseline (an underline a little under it,
 * a strikethrough a third of a size over it, a raised or lowered run's with
 * the run) along a stretch the line's ink covers. One that cannot be read is
 * taken as the paragraph's, as every rule there was before.
 */
function paragraphRuleSpansOnItsLines(
  streamBytes: Uint8Array,
  source: string,
  shows: readonly OwnedShow[],
  basis: readonly [number, number, number, number],
  lineBaselines: readonly number[],
): Readonly<{ start: number; end: number }>[] {
  const lines = lineBaselines.map((baseline, line) => {
    const inked = shows.filter((show) => show.line === line && show.inked);
    return inked.length ? { baseline, left: Math.min(...inked.map((show) => show.inked!.start)),
      right: Math.max(...inked.map((show) => show.inked!.end)), size: Math.max(...inked.map((show) => show.size)) } : null;
  });
  return paragraphRuleSpans(streamBytes, source, shows).filter((rule) => {
    const ends = pdfNativeTextRuleBlockEnds(streamBytes.slice(rule.start, rule.end));
    if (!ends) return true;
    const from = ends.from[0] * basis[0] + ends.from[1] * basis[1], to = ends.to[0] * basis[0] + ends.to[1] * basis[1];
    const across = ends.from[0] * basis[2] + ends.from[1] * basis[3];
    return lines.some((line) => line !== null &&
      across - line.baseline >= -RULE_ON_LINE_BELOW_EM * line.size && across - line.baseline <= RULE_ON_LINE_ABOVE_EM * line.size &&
      Math.max(from, to) > line.left && Math.min(from, to) < line.right);
  });
}
/** How far under and over a line's baseline a rule of the line may run, in
 * shares of the line's largest size. */
const RULE_ON_LINE_BELOW_EM = 0.5;
const RULE_ON_LINE_ABOVE_EM = 0.75;

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
  /** What ends the text object the block is set in, and begins it again. */
  wrap?: Readonly<{ before: string; after: string }>,
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
    ? Uint8Array.from([...asciiBytes(out + (wrap?.before ?? "")), 0x0a, ...MRAI_NATIVE_BOX_MARK_OPEN, ...block,
      ...MRAI_NATIVE_BOX_MARK_CLOSE, ...asciiBytes(wrap?.after ?? "")])
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
  // (A write empties the shows it draws again. The passes a weight drew of a
  // Type3 show stand after the show, placed from where its pen ended:
  // emptied, the show would leave them drawn alone, somewhere else.)
  if (set && paragraph.shows.some((show) => show.show.overstruck)) refuse(MRAI_NATIVE_TYPE3_BOLD_STANDS_MESSAGE);
  const originalBytes = streamBytes.slice(spanStart, insertion);
  // The owned shows and their fake-bold passes, emptied in stream order.
  const emptied = [...paragraph.shows.map((show) => ({ start: show.start, end: show.end })), ...paragraph.twins]
    .sort((left, right) => left.start - right.start);
  const replacement = rewrittenSpan(source.slice(spanStart, insertion), spanStart, emptied,
    paragraph.pens, set?.bytes ?? null, paragraph.rules, set?.carried ?? [], paragraph.blockWrap);
  // One edit writes at most as many bytes as a transaction carries; a
  // paragraph whose shows lie that far apart in its stream (a drawing set
  // between the lines of a CAD label) is left to the routes after this one.
  if (replacement.length > PARAGRAPH_REWRITE_MAX_REPLACEMENT_BYTES) {
    // Said as it is: other drawing stands between its shows, or the
    // paragraph itself is drawn in that many bytes (a browser's print sets a
    // line glyph by glyph, a show and a matrix each: 1,923 shows and 79 KB
    // for one paragraph of pdf-walk-1000 0952, 179 KB to draw again).
    const apart = paragraph.shows.some((show, index) => index > 0 &&
      show.start - paragraph.shows[index - 1].end > BLOCK_INSERTION_REACH_BYTES);
    refuse(apart ? "The paragraph's shows lie too far apart in their stream for one edit to rewrite."
      : MRAI_NATIVE_PARAGRAPH_TOO_LARGE_MESSAGE);
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










function paragraphTextRules(paragraph: ReadParagraph): Readonly<{ underline?: true; strikethrough?: true }> | null {
  const text = characters(paragraph.selection.item.str);
  const inked = text.flatMap((character, index) => isBlank(character) ? [] : [index]);
  // (Read with glyphs no character is named for: which show draws a
  // character is not known, and its rules are those under every show of it
  // that takes one -- ReadParagraph.unreadGlyphs.)
  const ruledShows = paragraph.shows.filter((show) => show.inked && show.inked.end > show.inked.start && !show.readsBlank);
  const ruled = (style: "underline" | "strikethrough") => paragraph.unreadGlyphs
    ? ruledShows.length > 0 && ruledShows.every((show) => show[style] === true)
    : inked.length > 0 && inked.every((index) =>
      paragraph.shows[paragraph.characterShows[index]]?.[style] === true || paragraph.characterRules.get(index)?.[style] === true);
  const underline = ruled("underline"), strikethrough = ruled("strikethrough");
  return underline || strikethrough
    ? Object.freeze({ ...(underline ? { underline: true as const } : {}), ...(strikethrough ? { strikethrough: true as const } : {}) })
    : null;
}

/**
 * Those rules for a box that opens with another proof than this tier's (the
 * surgical routes prove a title's four lines, a line of fifteen shows): read
 * as this tier reads the box, since it is this tier that writes the box's
 * formats and takes its rules off. Asked only where the box's stream holds
 * a rule of this writer's at all. Null where it has none under all of its
 * text, or this tier does not read the box.
 */
export async function readMraiNativeParagraphTextRules(
  kernel: MraiNativeParagraphRewriteKernel,
  selection: PdfNativeListTextSelectionEvidence,
  plan: PdfNativeListTextReplacementPlan,
): Promise<Readonly<{ underline?: true; strikethrough?: true }> | null> {
  try {
    if (!pdfNativeStreamHoldsTextRule((await kernel.readStream(plan)).streamBytes)) return null;
    // Read as the write that draws them reads it: the rules of a paragraph
    // set where it stands are under its text whatever the rewrite's own
    // measure of that text would refuse.
    let rules: ReturnType<typeof paragraphTextRules> = null;
    try {
      rules = paragraphTextRules(await readParagraph(kernel, selection, true));
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
    }
    // (Or the rules a format written by state drew beside its text objects.)
    return rules ?? rulesOfLook(await readStandingStateLook(kernel, selection));
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError)) throw error;
    return null;
  }
}

function rulesOfLook(look: MraiNativeStandingTextLook | null): Readonly<{ underline?: true; strikethrough?: true }> | null {
  return look && (look.underline || look.strikethrough)
    ? Object.freeze({ ...(look.underline ? { underline: true as const } : {}), ...(look.strikethrough ? { strikethrough: true as const } : {}) })
    : null;
}













export async function readMraiNativeParagraphStandingLook(
  kernel: MraiNativeParagraphRewriteKernel,
  selection: PdfNativeListTextSelectionEvidence,
): Promise<Readonly<{
  textRules: Readonly<{ underline?: true; strikethrough?: true }> | null;
  syntheticBold: boolean;
  slanted: boolean;
}> | null> {
  try {
    const paragraph = await readParagraph(kernel, selection, true);
    const text = characters(paragraph.selection.item.str);
    const inked = text.flatMap((character, index) => isBlank(character) ? [] : [index]);
    const every = (test: (show: OwnedShow) => boolean) => inked.length > 0 && inked.every((index) => {
      const show = paragraph.shows[paragraph.characterShows[index]];
      return show !== undefined && test(show);
    });
    const textRules = paragraphTextRules(paragraph) ??
      // (A paragraph too large to draw again takes its rules by state: read there.)
      (pdfNativeStreamHoldsTextRule(paragraph.streamBytes) ? rulesOfLook(await readStandingStateLook(kernel, selection)) : null);
    return Object.freeze({
      textRules,
      syntheticBold: every((show) => show.stroked),
      slanted: every((show) => show.slant !== 0),
    });
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError)) throw error;
    
    // state alone, and is read the same way (planFormatByStateWhereTextStands).
    const byState = async () => {
      const look = await readStandingStateLook(kernel, selection);
      return look && Object.freeze({ textRules: rulesOfLook(look), syntheticBold: look.syntheticBold === true,
        slanted: look.slanted === true });
    };
    // A box no item proof reads whole (a TeX paragraph's unnamed symbols)
    // takes those formats by its shows alone, and is read the same way
    // (planFormatWhereUnprovedTextStands).
    if (!standsUnproved(kernel, selection, error)) return byState();
    try {
      const paragraph = await readParagraphWhereItStands(kernel, selection, true);
      const inked = paragraph.shows.filter((show) => show.inked);
      return Object.freeze({
        textRules: paragraphTextRules(paragraph),
        syntheticBold: inked.length > 0 && inked.every((show) => show.stroked),
        slanted: inked.length > 0 && inked.every((show) => show.slant !== 0),
      });
    } catch (placedError) {
      if (!(placedError instanceof PdfNativeTransactionError)) throw placedError;
      return byState();
    }
  }
}





function standsUnproved(
  kernel: MraiNativeParagraphRewriteKernel,
  selection: PdfNativeListTextSelectionEvidence,
  refusal: PdfNativeTransactionError,
): boolean {
  return Boolean(kernel.standingPageStreams) && /^Reader item \d+ of the paragraph does not prove alone: /u.test(refusal.message) &&
    !pdfTextHasRightToLeftRun(selection.item.str);
}

/**
 * The refusal of a format asked of the whole of a box whose text the reader
 * shows in another order than the page draws it (right-to-left script), where
 * the page's own shows cannot be proven to be that box and nothing else. The
 * reason follows it in brackets.
 */
export const MRAI_NATIVE_STANDING_FORMAT_REFUSAL =
  "This text is shown in another order than the page draws it, and its drawn text could not be proven to be all of this box and nothing else; the format was not applied.";

/** The refusal of a format asked of such a box that every show of it is
 * drawn in already: nothing is written. */
export const MRAI_NATIVE_STANDING_FORMAT_UNCHANGED_REFUSAL =
  "This text is drawn in the format asked already; nothing was written.";

/** Whether the page draws a show bold by its face (as its name says:
 * faceStyle) with no stroke on it. */
function standsInBoldFace(show: Pick<OwnedShow, "font" | "stroked">): boolean {
  return !show.stroked && faceStyle(show.font)[0] === 1;
}









function strokeAskedChanges(asked: boolean | undefined, show: Pick<OwnedShow, "font" | "stroked">): boolean {
  return asked !== undefined && asked !== show.stroked && !(asked && standsInBoldFace(show));
}

/** What a box's look says of the bold faces it is drawn in (MraiNativeStandingTextLook). */
function boldFaceLookOf(inked: readonly Pick<OwnedShow, "font" | "stroked">[]):
  Readonly<{ boldFaces?: "only" | "beside-strokes"; drawnBold?: true }> {
  if (!inked.some(standsInBoldFace)) return Object.freeze({});
  return Object.freeze({ boldFaces: inked.some((show) => show.stroked) ? "beside-strokes" as const : "only" as const,
    ...(inked.every((show) => show.stroked || standsInBoldFace(show)) ? { drawnBold: true as const } : {}) });
}

/**
 * Refuses a bold taken off a box whose only bold is its faces' own -- an
 * inked show of it stands in a bold face and none is stroked
 * (MRAI_NATIVE_BOLD_FACE_KEPT_REFUSAL): nothing would be written, and the
 * box would read bold still. Where a show of it is stroked (a regular comma
 * made bold beside words in a bold face, pdf-walk-1000 0444) the bold taken
 * off is that stroke taken off: the page's own bytes again, its bold words
 * bold as the page drew them.
 */
function refuseBoldTakenOffBoldFace(asked: boolean | undefined, inked: readonly Pick<OwnedShow, "font" | "stroked">[]): void {
  if (asked === false && !inked.some((show) => show.stroked) && inked.some(standsInBoldFace)) {
    refuse(MRAI_NATIVE_BOLD_FACE_KEPT_REFUSAL);
  }
}

function refuseStanding(reason: string): never {
  return refuse(`${MRAI_NATIVE_STANDING_FORMAT_REFUSAL} (${reason})`);
}





















function standingBlockSplit(
  source: string,
  object: Readonly<{ start: number; end: number }>,
  lastShowEnd: number,
): Readonly<{ at: number; lineMatrix: Matrix }> | null {
  const tokens = mraiNativeContentTokens(source, object.start + 2, object.end - 2);
  if (!tokens) return null;
  const number = (text: string) => /^[+-]?(?:\d+\.?\d*|\.\d+)$/u.test(text) ? Number(text) : null;
  let depth = 0;
  let line: Matrix | null = [1, 0, 0, 1, 0, 0];
  let operands: string[] = [];
  let at: number | null = null;
  for (const token of tokens) {
    const text = source.slice(token.start, token.end);
    if (!token.operator) {
      operands.push(text);
      continue;
    }
    const taken = operands;
    operands = [];
    if (at !== null) {
      // Other text after the cut: the object is worth cutting.
      if ((text === "Tj" || text === "TJ" || text === "'" || text === '"') &&
        !/^(?:\(\)|<>|\[\s*\])$/u.test(taken.at(-1) ?? "")) return line ? Object.freeze({ at, lineMatrix: line }) : null;
      continue;
    }
    if (token.start >= lastShowEnd) {
      if (text !== "EMC" || depth < 1) return null;
      depth -= 1;
      if (depth === 0) at = token.end;
      continue;
    }
    if (text === "BDC" || text === "BMC") depth += 1;
    else if (text === "EMC") {
      depth -= 1;
      if (depth < 0) return null;
    } else if (text === "Tm") {
      const values = taken.map(number);
      line = values.length === 6 && values.every((value) => value !== null) ? values as unknown as Matrix : null;
    } else if (text === "Td" || text === "TD") {
      const values = taken.map(number);
      line = line && values.length === 2 && values[0] !== null && values[1] !== null
        ? [line[0], line[1], line[2], line[3], values[0] * line[0] + values[1] * line[2] + line[4],
          values[0] * line[1] + values[1] * line[3] + line[5]]
        : null;
    } else if (text === "T*" || text === "'" || text === '"') line = null;
    if (token.end === lastShowEnd && depth === 0) at = token.end;
  }
  return null;
}

/** How far across a reader item's baseline a show of it may stand, as a share
 * of the item's height: the reader keeps a glyph on the item's line while it
 * steps no further than PDFJS_VERTICAL_SHIFT_RATIO of it. */
const STANDING_ACROSS_EM = 0.3;
/** How far past an item's ends along its line the middle of a glyph of it
 * may be (a mark of no width standing at its very end). */
const STANDING_ALONG_EM = 0.1;




const STANDING_MARK_BELOW_EM = 0.45;
const STANDING_MARK_ABOVE_EM = 1.25;

/**
 * A box read for a format written where its text stands, without the reader's
 * character order: what readParagraph answers, for a box of right-to-left (or
 * mixed) text, which the reader shows in logical order and the page draws in
 * visual order -- no item of it is proved alone, and none of its characters
 * is told to be a particular glyph.
 *
 * The box owns the shows of the page's own content whose inked glyphs all
 * stand in a reader item of the box (by the item's own place, direction and
 * width on the page, each glyph by its middle), line by line; and it owns
 * them only where, on every line, the characters those shows draw are the
 * line's characters as the reader read them -- the same characters as often,
 * in whatever order. A show that stands partly in the box, a show in a font
 * the page does not prove, text under a replacement text, shows of more than
 * one stream, or a line whose characters differ: refused, nothing is read.
 *
 * A glyph the font only calls white space (inkedBlankCodes) is ink here like
 * any other: the reader reads a blank of it and no item holds it, so no count
 * of characters proves it, and it is owned by its place alone -- in a reader
 * item of the box like every glyph, or, a show of nothing else that the page
 * stands off the baseline (the mark over a letter of an Urdu face), within
 * the reach of an item's line AND drawn between two shows the box owns.
 * Drawn at the edge of the box's shows it is neither proven the box's nor
 * another's: refused. Drawn among other text it is that text's.
 *
 * `characterShows` is then dealt by count, not by identity: each line's
 * characters go to that line's shows in the order both stand in, as many to
 * a show as it draws. It says which LINE and which SHOW COUNT a character
 * belongs to and nothing else, which is all a format that gives every
 * character the same change needs (setParagraphWhereItStands draws each show
 * whole, by its own operands, at its own place). Nothing that tells one
 * character from another may be read from it: the caller writes no text and
 * no format of a part with what this returns.
 */
async function readParagraphWhereItStands(
  kernel: MraiNativeParagraphRewriteKernel,
  named: PdfNativeListTextSelectionEvidence,
  /** The box may hold glyphs its fonts name no character for (a TeX
   * paragraph's symbols): see planFormatWhereUnprovedTextStands. Such a glyph
   * is owned by its place like every other, and counted, not read: a line
   * that holds some draws every character its named glyphs read as, and the
   * reader reads no more characters beside them than those glyphs could be
   * (UNNAMED_GLYPH_MAX_CHARACTERS each). What the item proofs this read does
   * without would have refused is then asked here: a show cut by its clip,
   * a glyph drawn twice in one place, text the box does not hold standing on
   * a line of it. */
  unnamedGlyphs = false,
  /** Read for a format written by state alone (planFormatByStateWhereTextStands):
   * no show is emptied, drawn again or moved among the others, so what only
   * a redrawn block needs is not asked -- that the box's shows stand one
   * after another in the stream with no other text between them, under one
   * clip, each cut by none, with the text after each placed afresh. A glyph
   * the page draws again at the same place (a fake bold's passes) is the
   * box's with the first, and counted once. */
  inPlace = false,
  /** Its shows told by where the reader's items start in the stream, not by
   * where their glyphs stand (ownShowsByPins below): asked where the read by
   * place refused the box (readParagraphForStateFormat). In place only. */
  pinned = false,
  /** The box's shows may have other text of the stream drawn between them:
   * the block then draws the box after that text, and the reader comes to it
   * in another order than it did. Asked only by a caller that reads the page
   * it wrote back with the reader itself (wordsReadByReader). */
  amongOtherText = false,
): Promise<ReadParagraph> {
  const streamsOf = kernel.standingPageStreams;
  if (!streamsOf) return refuse(MRAI_NATIVE_REORDERED_TEXT_MESSAGE);
  const selection = kernel.exactSelection(named);
  const items = selection.items ?? [selection.item];
  if (items.length < 1 || items.length > 256) refuse("Paragraph rewrite item count is unsupported.");
  const lineIndexes = kernel.visualLineIndexes(items);
  const originalText = selection.items ? summarizePdfNativeSelectionItems(items) : selection.item.str;
  if (originalText !== selection.item.str || selection.startOffset !== 0 ||
    selection.endOffset !== selection.item.str.length) {
    refuse("Paragraph rewrite needs the whole paragraph's text.");
  }
  const text = characters(originalText);
  const lineCount = Math.max(...lineIndexes) + 1;
  const inkedOf = (value: string) => characters(value).filter((character) => !isReaderBlank(character));
  const sameCharacters = (left: readonly string[], right: readonly string[]) => {
    if (left.length !== right.length) return false;
    const a = [...left].sort(), b = [...right].sort();
    return a.every((character, index) => character === b[index]);
  };
  // Each line's characters in the box's text: the text is its items line by
  // line, so the text of the lines up to one is where that line ends.
  const lineEnds: number[] = [];
  for (let line = 0; line < lineCount; line += 1) {
    const count = lineIndexes.filter((index) => index <= line).length;
    const head = items.slice(0, count);
    if (!lineIndexes.slice(0, count).includes(line) || lineIndexes.slice(0, count).some((index) => index > line)) {
      refuseStanding("its lines are not read one after another");
    }
    let end = text.length;
    if (line < lineCount - 1) {
      const headLines = kernel.visualLineIndexes(head);
      const prefix = selection.items ? summarizePdfNativeSelectionItems(head) : originalText;
      if (headLines.length !== count || headLines.some((index, at) => index !== lineIndexes[at]) ||
        !originalText.startsWith(prefix)) refuseStanding("its lines cannot be told apart in its text");
      end = characters(prefix).length;
    }
    const start = lineEnds[line - 1] ?? 0;
    if (!sameCharacters(text.slice(start, end).filter((character) => !isReaderBlank(character)),
      items.flatMap((item, index) => lineIndexes[index] === line ? inkedOf(item.str) : []))) {
      refuseStanding("its lines cannot be told apart in its text");
    }
    lineEnds.push(end);
  }
  // Where the reader stands each item on the page.
  const rects = items.flatMap((item, index) => {
    if (!inkedOf(item.str).length) return [];
    const [a, b, c, d, e, f] = item.transform;
    const length = Math.hypot(a, b), height = Math.hypot(c, d);
    if (!(length > 0) || !(height > 0) || !(item.width >= 0)) refuseStanding("an item of it has no place on the page");
    // (`up`: which side of its baseline its glyphs stand on.)
    return [{ item: index, line: lineIndexes[index], x: e, y: f, ux: a / length, uy: b / length, up: a * d - b * c < 0 ? -1 : 1,
      width: item.width, height }];
  });
  const rectAt = (x: number, y: number) => {
    let best: { item: number; line: number; across: number } | null = null;
    for (const rect of rects) {
      const dx = x - rect.x, dy = y - rect.y;
      const along = dx * rect.ux + dy * rect.uy, across = Math.abs(dy * rect.ux - dx * rect.uy);
      if (along < -STANDING_ALONG_EM * rect.height || along > rect.width + STANDING_ALONG_EM * rect.height ||
        across > STANDING_ACROSS_EM * rect.height) continue;
      if (!best || across < best.across) best = { item: rect.item, line: rect.line, across };
    }
    return best;
  };
  // The items within whose line's reach a point stands (a mark's): along the
  // item's baseline and square to it, however the item's glyphs lean.
  const reachAt = (x: number, y: number) => rects.filter((rect) => {
    const dx = x - rect.x, dy = y - rect.y;
    const along = dx * rect.ux + dy * rect.uy, up = (dy * rect.ux - dx * rect.uy) * rect.up;
    return along >= -STANDING_ALONG_EM * rect.height && along <= rect.width + STANDING_ALONG_EM * rect.height &&
      up >= -STANDING_MARK_BELOW_EM * rect.height && up <= STANDING_MARK_ABOVE_EM * rect.height;
  });
  const readsBlank = (glyph: Readonly<{ text: string | undefined }>) => /^\s*$/u.test(glyph.text ?? "x");
  type Found = Readonly<{ show: MraiNativeParagraphRewriteShow; font: MraiNativeParagraphRewriteFont; item: number; line: number;
    glyphs: ReturnType<typeof showGlyphs>;
    /** It inks, and every glyph it inks is one the reader reads as a blank. */
    marks: boolean;
    /** The range of its codes the box holds, where the show draws other
     * text too (in place only): `glyphs` are that range's. */
    partial?: Readonly<{ codeStart: number; codeEnd: number }> }>;
  /**
   * The box's shows told by the reader's own arithmetic, not by where glyphs
   * stand: a reader item is a run of glyphs one after another in the stream,
   * and its origin is the pen of the first of them. So each item of the box
   * is pinned to the glyph of the page that starts at its origin (or, an
   * item the reader cut from the one before it at a blank, to the glyph
   * after that one's run: its origin is the pen that run ended at), and the
   * glyphs after it in the stream are its own for as long as they draw the
   * characters it reads as -- the same characters as often, to the last one.
   * A glyph is then the box's because the reader read it into an item of the
   * box, wherever a face stands it: the letters of a Nastaleeq word step up
   * off the baseline, the middle of one stood outside its own item and the
   * middle of a letter of the next word inside it, and by place such a box
   * was never proven (pdf-walk-1000 0339: 25 of its boxes).
   *
   * What the reader reads of a glyph is not always where it draws it, and
   * each such case is told here from the stream, never guessed:
   * - A combining mark (the reader's zero-width diacritic) is read into the
   *   item before it in the stream, without a look at where it stands. One
   *   that does not stand on that item's line is not its own: the item ends
   *   before it and reads a mark its glyphs do not draw, and the mark is the
   *   box's that holds the glyph it is drawn with (a show's opening mark
   *   with the show; a mark alone in a show by the shows beside it, as a
   *   glyph read as a blank is).
   * - A glyph the reader reads as a blank is in no item. Drawn inside an
   *   item's run, or between two runs of the box, it is the box's; drawn
   *   between two shows of other text, that text's; drawn at the edge of the
   *   box's shows it is the box's only where its pen joins the box's show
   *   beside it, and where it stands within the reach of a line of the box
   *   without that, it is not known whose: refused, as it was.
   * - A glyph drawn again where a glyph of the box stands (a fake bold's
   *   second pass) is the box's with the first.
   * A glyph of no named character, a show under a replacement text, a show
   *   in a font the page does not prove: not read this way (refused).
   */
  const ownShowsByPins = (
    allShows: readonly MraiNativeParagraphRewriteShow[],
    fonts: ReadonlyMap<string, MraiNativeParagraphRewriteFont>,
  ): { found: Found[]; strays: number; edges: string[] } => {
    type PinGlyph = { show: number; index: number; code: number; start: readonly [number, number]; end: readonly [number, number];
      /** Where its own advance ends: `end` without the character spacing after it (the reader's width of it). */
      advanced: readonly [number, number];
      blank: boolean; ink: boolean; characters: readonly string[] | null; mark: boolean };
    const inOrder = allShows.filter((show) => show.codes.length > 0).sort((left, right) => left.start - right.start);
    const read = inOrder.map((show) => {
      const font = fonts.get(show.fontName);
      let glyphs: ReturnType<typeof showGlyphs> | null = null;
      if (font && !show.vertical) {
        try {
          glyphs = showGlyphs(kernel, show, font);
        } catch (error) {
          if (!(error instanceof PdfNativeTransactionError)) throw error;
        }
      }
      return { show, font, glyphs, first: 0, count: 0 };
    });
    const all: PinGlyph[] = [];
    for (const [index, entry] of read.entries()) {
      entry.first = all.length;
      if (!entry.glyphs) continue;
      const matrix = composeMatrix(entry.show.textMatrix, entry.show.graphicsMatrix), scale = entry.show.horizontalScale / 100;
      for (const [at, glyph] of entry.glyphs.entries()) {
        const blank = readsBlank(glyph);
        const drawn = glyph.text === undefined ? null : inkedOf(kernel.normalizeUnicode(glyph.text));
        all.push({ show: index, index: at, code: glyph.code,
          start: applyMatrix(matrix, glyph.start * scale, entry.show.textRise),
          end: applyMatrix(matrix, glyph.end * scale, entry.show.textRise),
          advanced: applyMatrix(matrix, (glyph.end - entry.show.characterSpacing) * scale, entry.show.textRise),
          blank, ink: !blank || Boolean(entry.font?.inkedBlankCodes?.has(glyph.code)), characters: drawn,
          mark: Boolean(drawn?.length) && drawn!.every((character) => /^\p{Mn}$/u.test(character)) });
      }
      entry.count = all.length - entry.first;
    }
    // A glyph that takes the pen on as the reader reads one: of a named
    // character that is not a blank and not a combining mark.
    const positioned = (glyph: PinGlyph) => !glyph.blank && glyph.characters !== null && !glyph.mark;
    const near = (left: readonly [number, number], right: readonly [number, number], reach: number) =>
      Math.hypot(left[0] - right[0], left[1] - right[1]) <= reach;
    const owner = new Map<number, (typeof rects)[number]>();
    const runs = new Map<(typeof rects)[number], { first: number; last: number }>();
    const walk = (rect: (typeof rects)[number], from: number): { first: number; last: number } => {
      const left = new Map<string, number>();
      for (const character of inkedOf(items[rect.item].str)) left.set(character, (left.get(character) ?? 0) + 1);
      let remaining = [...left.values()].reduce((sum, count) => sum + count, 0);
      const differs = (): never =>
        refuseStanding(`the page does not draw line ${rect.line + 1} of it with the characters it reads as`);
      const consumed: number[] = [];
      let first = -1, last = -1;
      for (let at = from; remaining > 0; at += 1) {
        const glyph = all[at];
        // (Nothing left to read but combining marks.)
        const onlyMarksLeft = () => last >= 0 && [...left].every(([character, count]) => count < 1 || /^\p{Mn}$/u.test(character));
        // The stream ends: the marks left were read from a show that cannot be read.
        if (!glyph) return onlyMarksLeft() ? { first, last } : differs();
        if (glyph.blank) continue;
        if (glyph.characters === null) return refuseStanding("a glyph of it names no character");
        const entry = read[glyph.show];
        // A combining mark the page does not draw on the item's line: the
        // reader's zero-width diacritic, which it reads into the item before
        // it in the stream without a look at where it stands (the mark a
        // show of other text opens with). Where marks are all that is left
        // to read, the item ends before it; the mark is drawn with the text
        // it stands on.
        if (glyph.mark && !reachAt(...glyph.start).some((reached) => reached === rect)) {
          if (first < 0) continue;
          return onlyMarksLeft() ? { first, last } : differs();
        }
        const asked = new Map<string, number>();
        for (const character of glyph.characters) asked.set(character, (asked.get(character) ?? 0) + 1);
        if ([...asked].every(([character, count]) => (left.get(character) ?? 0) >= count)) {
          for (const [character, count] of asked) left.set(character, left.get(character)! - count);
          remaining -= glyph.characters.length;
          if (first < 0) first = at;
          last = at;
          consumed.push(at);
          continue;
        }
        // A combining mark before the item's first glyph: the item before's, as the reader reads it.
        if (glyph.mark && first < 0) continue;
        // A glyph of the item drawn again where it stands.
        const space = composeMatrix(entry.show.textMatrix, entry.show.graphicsMatrix);
        const size = Math.abs(entry.show.fontSize) * Math.hypot(space[0], space[1]);
        if (consumed.some((other) => all[other].code === glyph.code && read[all[other].show].show.fontName === entry.show.fontName &&
          near(all[other].start, glyph.start, DUPLICATE_GLYPH_REACH_EM * size))) {
          last = at;
          continue;
        }
        return differs();
      }
      if (first < 0) return differs();
      return { first, last };
    };
    const pending = new Set(rects);
    const resolve = (rect: (typeof rects)[number], from: number) => {
      const run = walk(rect, from);
      for (let at = run.first; at <= run.last; at += 1) {
        if (owner.has(at) && !all[at].blank) refuseStanding("two items of it are read from the same glyphs");
        owner.set(at, rect);
      }
      // Where the reader ends the item, the page ends its glyphs (a run that
      // ran on into other text would end elsewhere).
      const tolerance = Math.max(STANDING_ITEM_END_POINTS, STANDING_ITEM_END_EM * rect.height);
      const along = ([x, y]: readonly [number, number]) => (x - rect.x) * rect.ux + (y - rect.y) * rect.uy;
      const starts: number[] = [], pens: number[] = [], advances: number[] = [];
      for (let at = run.first; at <= run.last; at += 1) {
        if (!positioned(all[at])) continue;
        starts.push(along(all[at].start));
        pens.push(along(all[at].end));
        advances.push(along(all[at].advanced));
      }
      // (The reader's width ends with the last glyph's own advance; its pen
      // goes on by the character spacing after it.)
      if (!starts.length || Math.min(...starts) < -tolerance || Math.max(...starts) > rect.width + tolerance ||
        (Math.abs(Math.max(...advances) - rect.width) > tolerance && Math.abs(Math.max(...pens) - rect.width) > tolerance)) {
        refuseStanding("an item of it does not stand where the page draws its glyphs");
      }
      runs.set(rect, run);
      pending.delete(rect);
    };
    let pinnedAny = false;
    for (let progress = true; pending.size > 0 && progress;) {
      progress = false;
      for (const rect of [...pending]) {
        // The glyph that starts at the item's origin.
        let from = all.findIndex((glyph) => positioned(glyph) && near(glyph.start, [rect.x, rect.y], STANDING_PIN_POINTS));
        if (from < 0) {
          // Or the item the reader cut it from: it starts where that one ends.
          for (const [other, run] of runs) {
            if (near([other.x + other.width * other.ux, other.y + other.width * other.uy], [rect.x, rect.y], STANDING_CHAIN_POINTS)) {
              from = run.last + 1;
            }
          }
        }
        if (from < 0) continue;
        resolve(rect, from);
        pinnedAny = true;
        progress = true;
      }
      if (!progress && pending.size > 0) {
        // Or the pen a glyph of the page ends at (an item cut from one the
        // box does not hold): the glyph after it. Along the item's line: the
        // reader stands a cut item on its own baseline, whatever the glyph
        // before it stood on (a letter stepped up, held by another box), as
        // far off it as the reader keeps a glyph in one item.
        for (const rect of [...pending]) {
          const before = all.findIndex((glyph) => {
            if (!positioned(glyph)) return false;
            const dx = glyph.end[0] - rect.x, dy = glyph.end[1] - rect.y;
            return Math.abs(dx * rect.ux + dy * rect.uy) <= STANDING_PIN_POINTS &&
              Math.abs(dy * rect.ux - dx * rect.uy) <= STANDING_ACROSS_EM * rect.height;
          });
          if (before < 0) continue;
          resolve(rect, before + 1);
          pinnedAny = true;
          progress = true;
          break;
        }
      }
    }
    if (!pinnedAny) return { found: [], strays: 0, edges: [] };
    if (pending.size > 0) refuseStanding("an item of it starts at no glyph the page draws");
    // Its glyphs drawn again in a pass of their own, after the first.
    const ownedGlyphs = [...owner.keys()].filter((at) => positioned(all[at]));
    for (const entry of read) {
      if (!entry.glyphs || !entry.count) continue;
      const own = all.slice(entry.first, entry.first + entry.count);
      const drawn = own.filter(positioned);
      if (!drawn.length || own.some((glyph, at) => owner.has(entry.first + at))) continue;
      const space = composeMatrix(entry.show.textMatrix, entry.show.graphicsMatrix);
      const size = Math.abs(entry.show.fontSize) * Math.hypot(space[0], space[1]);
      const twins = drawn.map((glyph) => ownedGlyphs.find((at) => all[at].code === glyph.code &&
        read[all[at].show].show.fontName === entry.show.fontName && near(all[at].start, glyph.start, DUPLICATE_GLYPH_REACH_EM * size)));
      if (twins.some((twin) => twin === undefined)) continue;
      for (const [at, glyph] of own.entries()) if (positioned(glyph)) owner.set(entry.first + at, owner.get(twins[drawn.indexOf(glyph)]!)!);
    }
    // Show by show: the blanks and marks a show draws with the box's glyphs,
    // the shows of nothing but those, and the shows the box holds a part of.
    const positionedOf = (index: number) => {
      const entry = read[index], out: number[] = [];
      for (let at = entry.first; at < entry.first + entry.count; at += 1) if (positioned(all[at])) out.push(at);
      return out;
    };
    const found: Found[] = [];
    let strays = 0;
    const edges: string[] = [];
    const alone: number[] = [];
    for (const [index, entry] of read.entries()) {
      if (!entry.glyphs || !entry.count) {
        // (A show that cannot be read, inside a run of the box.)
        if (!entry.glyphs && [...runs.values()].some((run) => entry.first > run.first && entry.first <= run.last)) {
          refuseStanding(`text of it is drawn in a font the page does not prove, /${entry.show.fontName}`);
        }
        continue;
      }
      const letters = positionedOf(index);
      if (!letters.length) {
        alone.push(index);
        continue;
      }
      for (let at = entry.first; at < entry.first + entry.count; at += 1) {
        if (owner.has(at) || positioned(all[at]) || (!all[at].blank && all[at].characters === null)) continue;
        // (A blank or a mark goes with the glyph it is drawn before; after
        // the show's last glyph, with that one.)
        const next = letters.find((letter) => letter > at) ?? letters.at(-1)!;
        const with_ = owner.get(next);
        const previous = [...letters].reverse().find((letter) => letter < at);
        if (with_ && (previous === undefined || owner.has(previous) || next < at)) owner.set(at, with_);
      }
      const inked = [];
      for (let at = entry.first; at < entry.first + entry.count; at += 1) if (all[at].ink) inked.push(at);
      const own = inked.filter((at) => owner.has(at));
      if (!letters.some((at) => owner.has(at))) continue;
      const { show, font, glyphs } = entry;
      if (!font) return refuseStanding(`text of it is drawn in a font the page does not prove, /${show.fontName}`);
      if (show.text === null) refuse("A show of this paragraph stands under a replacement text.");
      const rect = owner.get(letters.find((at) => owner.has(at))!)!;
      if (own.some((at) => owner.get(at)!.line !== rect.line)) refuseStanding("a show of it also draws text outside it");
      if (own.length === inked.length) {
        found.push({ show, font, item: rect.item, line: rect.line, glyphs, marks: false });
        continue;
      }
      // The one stretch of it the box holds (standingShowCuts).
      const [head, tail] = [inked.indexOf(own[0]), inked.indexOf(own.at(-1)!)];
      if (!inPlace || !/^T[jJ]/u.test(show.operatorKind) || tail - head + 1 !== own.length) {
        refuseStanding("a show of it also draws text outside it");
      }
      const partial = Object.freeze({ codeStart: (own[0] - entry.first) * font.codeUnitBytes,
        codeEnd: (own.at(-1)! - entry.first + 1) * font.codeUnitBytes });
      found.push({ show, font, item: rect.item, line: rect.line, marks: false, partial,
        glyphs: glyphs.slice(partial.codeStart / font.codeUnitBytes, partial.codeEnd / font.codeUnitBytes) });
    }
    // The shows of blanks and of glyphs read as blanks alone, each by the
    // shows it is drawn between.
    const lettered = read.flatMap((_, index) => positionedOf(index).length ? [index] : []);
    for (const index of alone) {
      const entry = read[index];
      const own = all.slice(entry.first, entry.first + entry.count);
      // (Marks alone in a show, read into an item of the box.)
      const readInto = own.flatMap((glyph, at) => !glyph.blank && owner.has(entry.first + at) ? [owner.get(entry.first + at)!] : []);
      if (readInto.length) {
        if (own.some((glyph, at) => glyph.ink && !glyph.blank && !owner.has(entry.first + at)) || !entry.font || entry.show.text === null) {
          refuseStanding("a show of it also draws text outside it");
        }
        found.push({ show: entry.show, font: entry.font!, item: readInto[0].item, line: readInto[0].line, glyphs: entry.glyphs!, marks: false });
        continue;
      }
      if (own.some((glyph) => glyph.characters === null)) continue;
      const before = [...lettered].reverse().find((other) => other < index);
      const after = lettered.find((other) => other > index);
      const previous = before === undefined ? undefined : owner.get(positionedOf(before).at(-1)!);
      const next = after === undefined ? undefined : owner.get(positionedOf(after)[0]);
      const inked = own.some((glyph) => glyph.ink);
      const take = (rect: (typeof rects)[number]) => {
        if (!entry.font || entry.show.text === null) return;
        found.push({ show: entry.show, font: entry.font, item: rect.item, line: rect.line, glyphs: entry.glyphs!,
          marks: inked && own.every((glyph) => !glyph.ink || glyph.blank) });
      };
      if (previous && next) {
        take(previous);
        continue;
      }
      if (!inked || (!previous && !next)) continue;
      // At the edge of the box's shows: the box's where its pen joins them
      // (or it starts where the box's glyph after it starts: stood on it).
      const joins = previous
        ? near(all[positionedOf(before!).at(-1)!].end, own[0].start, STANDING_PIN_POINTS)
        : near(own.at(-1)!.end, all[positionedOf(after!)[0]].start, STANDING_PIN_POINTS) ||
          near(own[0].start, all[positionedOf(after!)[0]].start, STANDING_PIN_POINTS);
      if (joins) {
        take((previous ?? next)!);
        continue;
      }
      const scale = entry.show.horizontalScale / 100, matrix = composeMatrix(entry.show.textMatrix, entry.show.graphicsMatrix);
      // Drawn on a letter: every glyph of it that draws starts at the
      // advance of a named glyph, no further across from it than a mark
      // stands from its letter -- a part of that letter set as a glyph of
      // its own, which the face's map reads as a space (the small tah of
      // 「ٹ」 over the letter's body, a point over 「ج」: pdf-walk-1000 0948),
      // or a combining mark the reader read into the item before it in the
      // stream (the 「ً」 of a line's last word, drawn first and read into
      // the list number at the line's other end: 0354).
      // It is that letter's: other text's where the letter is, and the
      // box's where the letter is the box's. Nearest across wins, and a
      // letter of the other side nearly as near leaves it unknown, as
      // before. Taken for a glyph of the box's line drawn outside its shows,
      // or for an edge nobody could tell the owner of, every format of the
      // box beside it was refused -- and a mark over the box's own first
      // letter, drawn after other text, with it.
      const size = Math.abs(entry.show.fontSize) * Math.hypot(matrix[0], matrix[1]);
      const letterOf = (glyph: PinGlyph): "box" | "other" | null => {
        let box = Number.POSITIVE_INFINITY, other = Number.POSITIVE_INFINITY;
        for (const [at, letter] of all.entries()) {
          if (!positioned(letter)) continue;
          const along = [letter.advanced[0] - letter.start[0], letter.advanced[1] - letter.start[1]] as const;
          const length = Math.hypot(along[0], along[1]);
          if (!(length > STANDING_PIN_POINTS)) continue;
          const from = [glyph.start[0] - letter.start[0], glyph.start[1] - letter.start[1]] as const;
          const reached = (from[0] * along[0] + from[1] * along[1]) / length;
          // (A mark is set from its letter's edge as often as from inside
          // it: 「ً」 0.7 pt before the 「ا」 it stands on, 0354.)
          if (!(reached > -LETTER_PART_ALONG_EM * size && reached < length + LETTER_PART_ALONG_EM * size)) continue;
          const off = Math.abs(from[0] * along[1] - from[1] * along[0]) / length;
          if (off > LETTER_PART_ACROSS_EM * size) continue;
          if (owner.has(at)) box = Math.min(box, off);
          else other = Math.min(other, off);
        }
        const margin = STANDING_ACROSS_EM * size;
        return box + margin < other ? "box" : other + margin < box ? "other" : null;
      };
      const drawing = own.filter((glyph) => glyph.ink);
      const whose = drawing.length > 0 ? letterOf(drawing[0]) : null;
      if (whose !== null && drawing.every((glyph) => letterOf(glyph) === whose)) {
        if (whose === "box") take((previous ?? next)!);
        continue;
      }
      if (entry.glyphs!.some((glyph) => reachAt(...applyMatrix(matrix, ((glyph.start + glyph.end) / 2) * scale, entry.show.textRise)).length > 0)) {
        strays += 1;
        continue;
      }
      edges.push(`${entry.show.start}:${entry.show.end}`);
    }
    return { found, strays, edges };
  };
  const streams = await streamsOf(selection);
  const held: { stream: MraiNativeParagraphRewriteStandingStream; streamBytes: Uint8Array; pageContent: boolean;
    allShows: readonly MraiNativeParagraphRewriteShow[]; found: Found[]; strays: number; edges?: string[] }[] = [];
  const unread: string[] = [];
  for (const stream of streams) {
    let streamBytes: Uint8Array, pageContent: boolean, allShows: readonly MraiNativeParagraphRewriteShow[];
    try {
      ({ streamBytes, pageContent } = await kernel.readStream(stream.plan));
      allShows = kernel.readShows(streamBytes, stream.plan, stream.fonts);
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
      unread.push(error.message);
      continue;
    }
    if (pinned) {
      held.push({ stream, streamBytes, pageContent, allShows, ...ownShowsByPins(allShows, stream.fonts) });
      continue;
    }
    const found: Found[] = [];
    const raised: { show: MraiNativeParagraphRewriteShow; font: MraiNativeParagraphRewriteFont;
      glyphs: ReturnType<typeof showGlyphs>; near: readonly (typeof rects)[number][] }[] = [];
    for (const show of allShows) {
      if (show.codes.length === 0) continue;
      const matrix = composeMatrix(show.textMatrix, show.graphicsMatrix);
      const font = stream.fonts.get(show.fontName);
      let glyphs: ReturnType<typeof showGlyphs> | null = null;
      if (font && !show.vertical) {
        try {
          glyphs = showGlyphs(kernel, show, font);
        } catch (error) {
          if (!(error instanceof PdfNativeTransactionError)) throw error;
        }
      }
      const scale = show.horizontalScale / 100;
      const inked = glyphs?.filter((glyph) => !readsBlank(glyph) || Boolean(font?.inkedBlankCodes?.has(glyph.code))) ?? null;
      const marks = Boolean(inked?.length) && inked!.every(readsBlank);
      // What of it the box is asked for: the middle of each glyph it inks;
      // its origin where it inks none, or its glyphs cannot be measured.
      const points = inked?.length
        ? inked.map((glyph) => applyMatrix(matrix, ((glyph.start + glyph.end) / 2) * scale, show.textRise))
        : [applyMatrix(matrix, 0, show.textRise)];
      const places = points.map(([x, y]) => rectAt(x, y));
      const inside = places.flatMap((place) => place ? [place] : []);
      if (!inside.length) {
        // (Marks stood off the baseline: the box's where they are drawn
        // among its shows, which is known once those are -- below.)
        if (marks && font && glyphs && show.text !== null) {
          const reaches = points.map(([x, y]) => reachAt(x, y));
          const near = reaches[0].filter((rect) => reaches.every((reach) => reach.some((other) => other.line === rect.line)));
          if (near.length) raised.push({ show, font, glyphs, near });
        }
        continue;
      }
      // (White space the box's lines hold, or that stands at their ends: its
      // own where it stands on one of them, and no ink either way.)
      if (inked && !inked.length) {
        if (font && glyphs && show.text !== null) found.push({ show, font, item: inside[0].item, line: inside[0].line, glyphs, marks });
        continue;
      }
      if (!font || !glyphs || !inked) return refuseStanding(`text of it is drawn in a font the page does not prove, /${show.fontName}`);
      // (A show that holds a glyph its font names no character for reads as
      // no text either: told from one under a replacement text by its mark.)
      if (show.text === null && (!unnamedGlyphs || show.underActualText || show.underActualTextCluster ||
        !glyphs.some((glyph) => glyph.text === undefined))) {
        refuse("A show of this paragraph stands under a replacement text.");
      }
      if (inside.length !== places.length || inside.some((place) => place.line !== inside[0].line)) {
        // In place, a show that draws the end of one of the box's lines and
        // then other text (a table row set as one `TJ`, a line and the page
        // number after it) is the box's in the one stretch of its glyphs
        // that stands in the box: the format is written on that stretch, the
        // show cut in two where it ends (standingShowCuts). Anything else --
        // glyphs of it on two lines of the box, glyphs of other text between
        // the box's -- is refused as it was.
        const first = places.findIndex((place) => place !== null);
        const last = places.length - 1 - [...places].reverse().findIndex((place) => place !== null);
        if (!inPlace || !/^T[jJ]/u.test(show.operatorKind) || inside.some((place) => place.line !== inside[0].line) ||
          places.slice(first, last + 1).some((place) => place === null)) {
          refuseStanding("a show of it also draws text outside it");
        }
        const at = (glyph: (typeof glyphs)[number]) => glyphs.indexOf(glyph) * font.codeUnitBytes;
        const partial = Object.freeze({ codeStart: at(inked[first]), codeEnd: at(inked[last]) + font.codeUnitBytes });
        found.push({ show, font, item: inside[0].item, line: inside[0].line, marks, partial,
          glyphs: glyphs.slice(partial.codeStart / font.codeUnitBytes, partial.codeEnd / font.codeUnitBytes) });
        continue;
      }
      found.push({ show, font, item: inside[0].item, line: inside[0].line, glyphs, marks });
    }
    // The marks stood off the baseline, each by the shows it is drawn between
    // (other such marks aside): two of the box's, and it is the box's, on the
    // line of the one before it where it stands within that line's reach;
    // two of other text, and it is that text's; one of each, and it is not
    // known whose.
    let strays = 0;
    if (raised.length) {
      const inOrder = allShows.filter((show) => show.codes.length > 0).sort((left, right) => left.start - right.start);
      const order = new Map(inOrder.map((show, index) => [show, index] as const));
      const ownOf = new Map(found.map((own) => [own.show, own] as const));
      const apart = new Set(raised.map((entry) => entry.show));
      for (const entry of raised) {
        let before = order.get(entry.show)! - 1, after = before + 2;
        while (before >= 0 && apart.has(inOrder[before])) before -= 1;
        while (after < inOrder.length && apart.has(inOrder[after])) after += 1;
        const [previous, next] = [ownOf.get(inOrder[before]), ownOf.get(inOrder[after])];
        if (!previous && !next) continue;
        const place = previous && next
          ? entry.near.find((rect) => rect.line === previous.line) ?? entry.near.find((rect) => rect.line === next.line) : undefined;
        if (!place) {
          strays += 1;
          continue;
        }
        found.push({ show: entry.show, font: entry.font, item: place.item, line: place.line, glyphs: entry.glyphs, marks: true });
      }
    }
    held.push({ stream, streamBytes, pageContent, allShows, found, strays });
  }
  const holding = held.filter((entry) => entry.found.some((own) => own.glyphs.some((glyph) => !readsBlank(glyph))));
  if (holding.length !== 1) {
    return refuseStanding(holding.length ? "its text is drawn by more than one content stream"
      : `no text of the page's own content stands where it does${unread.length ? `: ${unread[0]}` : ""}`);
  }
  const { stream, streamBytes, pageContent, allShows } = holding[0];
  if (holding[0].strays > 0) {
    refuseStanding("a glyph the reader reads as a blank stands on a line of it, and is not drawn among its shows");
  }
  const found = [...holding[0].found].sort((left, right) => left.show.start - right.show.start);
  // The block draws its shows again one after another, in the order the page
  // draws them: so the reader must come to each of them from the show it came
  // to it from before, or it reads the box otherwise with no glyph of it
  // moved. Text that is not the box's drawn between two shows of it would be
  // taken out from between them (pdf-walk-1000 0339: a mark raised off the
  // line, held by no item of the box, stood between a blank and the word
  // after it; without it the reader read that blank as a space and the word
  // as two), and lines drawn in another order than they stand would change
  // places.
  const ownShows = new Set(found.map((own) => own.show));
  const [firstStart, lastStart] = [found[0].show.start, found.at(-1)!.show.start];
  if (!inPlace && !amongOtherText &&
    allShows.some((show) => show.codes.length > 0 && show.start > firstStart && show.start < lastStart && !ownShows.has(show))) {
    refuseStanding("other text is drawn between its shows");
  }
  if (!inPlace && found.some((own, index) => index > 0 && own.line < found[index - 1].line)) {
    refuseStanding("its lines are not drawn one after another");
  }
  // (In place: the passes a page draws a glyph again by, a hair from the first.)
  const repeated = new Set<string>();
  if (inPlace) {
    const drawnAt = new Map<string, (readonly [number, number, number])[]>();
    for (const [index, own] of found.entries()) {
      const matrix = composeMatrix(own.show.textMatrix, own.show.graphicsMatrix), scale = own.show.horizontalScale / 100;
      const reach = DUPLICATE_GLYPH_REACH_EM * own.show.fontSize * Math.hypot(matrix[0], matrix[1]);
      for (const [at, glyph] of own.glyphs.entries()) {
        if (readsBlank(glyph)) continue;
        const [x, y] = applyMatrix(matrix, glyph.start * scale, own.show.textRise);
        const key = `${own.show.fontName}:${glyph.code}`;
        const list = drawnAt.get(key) ?? [];
        if (list.some(([ox, oy, oreach]) => Math.hypot(ox - x, oy - y) <= Math.max(reach, oreach))) repeated.add(`${index}:${at}`);
        else list.push([x, y, reach]);
        drawnAt.set(key, list);
      }
    }
  }
  // The characters each show draws, as the reader reads a glyph's.
  // (In place, a glyph whose name no list reads is one no character is known for.)
  const unnamed = (own: Found, glyph: Readonly<{ code: number; text: string | undefined }>) =>
    glyph.text === undefined || (inPlace && unnamedGlyphs && Boolean(own.font.nameUnreadCodes?.has(glyph.code)) && !readsBlank(glyph));
  const drawnOf = (own: Found, index: number) => own.glyphs.flatMap((glyph, at) => {
    if (repeated.has(`${index}:${at}`)) return [];
    if (unnamed(own, glyph)) return unnamedGlyphs ? [] : refuseStanding("a glyph of it names no character");
    return inkedOf(kernel.normalizeUnicode(glyph.text!));
  });
  const drawn = found.map(drawnOf);
  // The lines that hold a glyph no character is named for, and how many.
  const unnamedOnLine = Array.from({ length: lineCount }, (_, line) => found.reduce((sum, own, index) =>
    own.line === line ? sum + own.glyphs.filter((glyph, at) => unnamed(own, glyph) && !repeated.has(`${index}:${at}`)).length : sum, 0));
  // (By pins each item was read against its own glyphs: ownShowsByPins.)
  for (let line = 0; line < lineCount && !pinned; line += 1) {
    const read = text.slice(lineEnds[line - 1] ?? 0, lineEnds[line]).filter((character) => !isReaderBlank(character));
    const named = found.flatMap((own, index) => own.line === line ? drawn[index] : []);
    if (unnamedOnLine[line] > 0) {
      // Every character its named glyphs draw is read on the line, and what
      // the reader reads beside them is no more than its unnamed glyphs could be.
      const left = new Map<string, number>();
      for (const character of read) left.set(character, (left.get(character) ?? 0) + 1);
      const held = named.every((character) => {
        const count = left.get(character) ?? 0;
        left.set(character, count - 1);
        return count > 0;
      });
      if (!held || read.length - named.length > UNNAMED_GLYPH_MAX_CHARACTERS * unnamedOnLine[line]) {
        refuseStanding(`the page does not draw line ${line + 1} of it with the characters it reads as`);
      }
    } else if (!sameCharacters(read, named)) {
      refuseStanding(`the page does not draw line ${line + 1} of it with the characters it reads as`);
    }
  }
  if (unnamedGlyphs) {
    if (!inPlace) refuseShowsNoItemProofWouldOwn(kernel, found, allShows, stream.fonts);
    // The block draws every show again at one place of the stream, under the
    // clip in force there (its last show's): a show drawn under another clip
    // would be cut by one that was never its own.
    const clipOf = (show: MraiNativeParagraphRewriteShow) => JSON.stringify(show.clipBox ?? null);
    if (!inPlace && found.some((own) => clipOf(own.show) !== clipOf(found.at(-1)!.show))) {
      refuseStanding("its shows are drawn under more than one clip");
    }
    // Each reader item of the box is what the reader reads of those shows:
    // a glyph of them starts where it starts and one ends where it ends
    // (the reader's blanks at its ends aside), as the proof of an item alone
    // finds its show. Items that stand anywhere else are not the page's own,
    // whatever shows their place takes in. (A glyph may stand in two items:
    // an accent's and the letter's it is drawn over.)
    const drawnGlyphs = found.flatMap((own) => {
      const matrix = composeMatrix(own.show.textMatrix, own.show.graphicsMatrix), scale = own.show.horizontalScale / 100;
      return own.glyphs.flatMap((glyph) => readsBlank(glyph) && !own.font.inkedBlankCodes?.has(glyph.code) ? [] : [{
        start: applyMatrix(matrix, glyph.start * scale, own.show.textRise),
        end: applyMatrix(matrix, glyph.end * scale, own.show.textRise),
      }]);
    });
    for (const rect of rects) {
      const along = ([x, y]: readonly [number, number]) => (x - rect.x) * rect.ux + (y - rect.y) * rect.uy;
      const across = ([x, y]: readonly [number, number]) => Math.abs((y - rect.y) * rect.ux - (x - rect.x) * rect.uy);
      const tolerance = Math.max(STANDING_ITEM_END_POINTS, STANDING_ITEM_END_EM * rect.height);
      const on = drawnGlyphs.filter((glyph) => across(glyph.start) <= tolerance);
      const read = characters(items[rect.item].str);
      if ((!isReaderBlank(read[0]) && !on.some((glyph) => Math.abs(Math.min(along(glyph.start), along(glyph.end))) <= tolerance)) ||
        (!isReaderBlank(read.at(-1)!) &&
          !on.some((glyph) => Math.abs(Math.max(along(glyph.start), along(glyph.end)) - rect.width) <= tolerance))) {
        refuseStanding("an item of it does not stand where the page draws its glyphs");
      }
    }
  }
  // Its shows, each as the block draws it again (readParagraph's geometry).
  const blockMatrix = found.at(-1)!.show.graphicsMatrix;
  const toBlock = invertMatrix(blockMatrix);
  const linearOf = (show: MraiNativeParagraphRewriteShow) =>
    composeMatrix(composeMatrix(show.textMatrix, show.graphicsMatrix), toBlock);
  const head = found.find((own) => own.glyphs.some((glyph) => !/^\s*$/u.test(glyph.text ?? "x")))!;
  const first = linearOf(head.show);
  const alongLength = Math.hypot(first[0], first[1]), acrossLength = Math.hypot(first[2], first[3]);
  if (!(alongLength > 0) || !(acrossLength > 0)) refuse("The paragraph's text space is degenerate.");
  const firstSlanted = Math.abs(first[0] * first[2] + first[1] * first[3]) > 1e-6 * alongLength * acrossLength;
  const firstHand = Math.sign(first[0] * first[3] - first[1] * first[2]);
  const basis: readonly [number, number, number, number] = Object.freeze(firstSlanted
    ? [first[0] / alongLength, first[1] / alongLength, (-firstHand * first[1]) / alongLength, (firstHand * first[0]) / alongLength] as const
    : [first[0] / alongLength, first[1] / alongLength, first[2] / acrossLength, first[3] / acrossLength] as const);
  if (Math.abs(basis[0] * basis[2] + basis[1] * basis[3]) > 1e-6 || firstHand === 0 ||
    Math.sign(basis[0] * basis[3] - basis[1] * basis[2]) !== firstHand ||
    firstHand * (blockMatrix[0] * blockMatrix[3] - blockMatrix[1] * blockMatrix[2]) <= 0) {
    refuse("The paragraph's text is mirrored.");
  }
  const shows: OwnedShow[] = [];
  for (const own of found) {
    const { show, font, glyphs } = own;
    if (show.vertical) refuse("Paragraph rewrite does not support vertical text.");
    if (show.renderingMode !== 0 && show.renderingMode !== 2) refuse("A show of this paragraph is not drawn filled.");
    if (!show.deviceColorKnown) refuse("A show of this paragraph is drawn in a colour that cannot be read.");
    const linear = linearOf(show);
    const along = Math.hypot(linear[0], linear[1]), fullAcross = Math.hypot(linear[2], linear[3]);
    if (!(along > 0) || !(fullAcross > 0) ||
      Math.abs(linear[0] * basis[1] - linear[1] * basis[0]) > 1e-6 * along ||
      linear[0] * basis[0] + linear[1] * basis[1] <= 0 || linear[2] * basis[2] + linear[3] * basis[3] <= 0) {
      refuse("The paragraph's shows are not drawn along one direction.");
    }
    const leaning = Math.abs(linear[2] * basis[3] - linear[3] * basis[2]) > 1e-6 * fullAcross;
    const across = leaning ? linear[2] * basis[2] + linear[3] * basis[3] : fullAcross;
    const slant = leaning ? (linear[2] * basis[0] + linear[3] * basis[1]) / across : 0;
    const [x, y] = applyMatrix(toBlock, ...applyMatrix(show.graphicsMatrix, ...applyMatrix(show.textMatrix, 0, 0)));
    const u = x * basis[0] + y * basis[1], v = x * basis[2] + y * basis[3];
    const scale = (show.horizontalScale / 100) * along;
    const inkedGlyphs = glyphs.filter((glyph) => !readsBlank(glyph) || Boolean(font.inkedBlankCodes?.has(glyph.code)));
    // (Its ink from where it starts furthest back to where it ends furthest
    // on, whichever way its array steps along the line: a producer of
    // right-to-left text may draw a show's glyphs from the right.)
    const ink = Object.freeze(inkedGlyphs.map((glyph) =>
      Object.freeze([u + Math.min(glyph.start, glyph.end) * scale, u + Math.max(glyph.start, glyph.end) * scale] as const)));
    const inked = ink.length
      ? Object.freeze({ start: Math.min(...ink.map(([start]) => start)), end: Math.max(...ink.map(([, end]) => end)),
        count: ink.length })
      : null;
    shows.push(Object.freeze({
      start: show.start, end: show.end, show, font, item: own.item, line: own.line, leaf: null,
      size: show.fontSize * across, characterSpacing: show.characterSpacing * across,
      horizontalScale: show.horizontalScale * (along / across), rise: show.textRise * across,
      slant, color: show.deviceColor, u, v, inked, ink, partial: own.partial ?? null, twin: null, stroked: show.renderingMode === 2,
      // (Ruled on its line, wherever the page stands it across it: a word's
      // letters stepped up along the line are under one rule with it.)
      ruleRise: show.textRise * across, ...(own.marks ? { readsBlank: true as const } : {}),
    }));
  }
  // Its lines (readParagraph's): each on the baseline most of its glyphs
  // stand on (the letters of a Nastaleeq word step up from it, one show
  // each), the glyphs the reader reads as blanks aside.
  const lineBaselines: number[] = [];
  const lineExtents: { left: number; right: number }[] = [];
  for (let line = 0; line < lineCount; line += 1) {
    const inked = shows.filter((show) => show.line === line && show.inked);
    const letters = inked.filter((show) => !show.readsBlank);
    if (!letters.length) refuse("A line of the paragraph draws nothing.");
    const level = letters.filter((show) => Math.abs(show.rise) <= LAYOUT_TOLERANCE);
    const standing = new Map<number, { across: number; count: number }>();
    for (const show of level.length ? level : letters) {
      const across = show.v + show.rise, key = Math.round(across * 1e4);
      const there = standing.get(key);
      if (there) there.count += show.inked!.count;
      else standing.set(key, { across, count: show.inked!.count });
    }
    lineBaselines.push([...standing.values()].reduce((best, entry) => entry.count > best.count ? entry : best).across);
    lineExtents.push({ left: Math.min(...inked.map((show) => show.inked!.start)),
      right: Math.max(...inked.map((show) => show.inked!.end)) });
  }
  for (let line = 1; line < lineCount; line += 1) {
    if (!(lineBaselines[line - 1] - lineBaselines[line] > LAYOUT_TOLERANCE)) {
      refuse("The paragraph's lines do not run down its column.");
    }
  }
  const bodySize = shows.filter((show) => show.inked && !show.readsBlank)
    .reduce((best, show) => show.inked!.count > best.inked!.count ? show : best).size;
  const reach = lineBaselines.map((baseline, line) => lineReachOf(shows.filter((show) => show.line === line && show.inked && !show.readsBlank)
    .map((show) => ({ size: show.size, rise: show.v + show.rise - baseline })), bodySize));
  const steps = lineBaselines.slice(1).map((baseline, index) => (lineBaselines[index] - baseline) /
    (1 + (reach[index + 1].above + reach[index].below) / bodySize)).sort((a, b) => a - b);
  const middle = Math.floor(steps.length / 2);
  const lineStep = steps.length === 0 ? SINGLE_LINE_STEP_EM * bodySize
    : steps.length % 2 === 1 ? steps[middle] : (steps[middle - 1] + steps[middle]) / 2;
  const firstBaseline = lineBaselines[0] + firstLineLeading(lineStep, bodySize) * reach[0].above;
  const firstLeft = lineExtents[0].left;
  const restLeft = lineCount > 1 ? Math.min(...lineExtents.slice(1).map((extent) => extent.left)) : firstLeft;
  const frameLeft = Math.min(firstLeft, restLeft);
  const frameRight = Math.max(...lineExtents.map((extent) => extent.right));
  const alignment = paragraphAlignment(lineExtents, frameRight, restLeft);
  // Its characters dealt to its shows: line by line, as many to a show as it
  // draws (see above: by count, never by which character is which glyph).
  const owners: (number | null)[] = text.map(() => null);
  for (let line = 0; line < lineCount; line += 1) {
    const own = text.flatMap((character, index) =>
      index >= (lineEnds[line - 1] ?? 0) && index < lineEnds[line] && !isBlank(character) ? [index] : []);
    // (A show of glyphs the reader reads as blanks draws none of them.)
    const lineShows = shows.flatMap((show, index) => show.line === line && show.inked && drawn[index].length ? [index] : []);
    if (unnamedOnLine[line] > 0) {
      // (How many characters each show draws is not known: the line's first
      // inked show stands for the line; see ReadParagraph.unreadGlyphs.)
      const first = shows.findIndex((show) => show.line === line && show.inked);
      for (const index of own) owners[index] = first;
      continue;
    }
    let next = 0;
    for (const [order, index] of lineShows.entries()) {
      let count = order === lineShows.length - 1 ? own.length - next : drawn[index].length;
      // (By pins a show may draw a mark the reader read into another item,
      // and an item read one its shows do not draw: dealt as far as they go.)
      if (pinned) count = Math.max(0, Math.min(count, own.length - next));
      else if (count < 1 || next + count > own.length) refuseStanding("a show of it draws no character it reads as");
      for (let at = 0; at < count; at += 1) owners[own[next + at]] = index;
      next += count;
    }
  }
  const opening = owners.find((owner) => owner !== null);
  if (opening === undefined || opening === null) return refuse("The paragraph draws no character.");
  let previous: number = opening;
  const characterShows = Object.freeze(owners.map((owner) => {
    if (owner !== null) previous = owner;
    return owner ?? previous;
  }));
  // Its rules: those drawn under (through) whole shows of it only (no
  // character of it is told from another, so none is read as ruled alone).
  const source = LATIN1.decode(streamBytes);
  // (In place: its rules stand beside each text object of it, and are read there -- standingRuleGroups.)
  const { rules } = inPlace ? { rules: [] } : paragraphOwnRules({ streamBytes, source, shows, basis, blockMatrix, lineBaselines,
    paragraphCharacters: [], characterShows: [] });
  for (const [index, show] of shows.entries()) {
    shows[index] = Object.freeze({ ...show, rise: show.v + show.rise - lineBaselines[show.line] });
  }
  // No show of it is emptied to a pen (that is measured from a proof of the
  // show alone): the text after each must be placed afresh.
  const ownedKeys = new Set(shows.map((show) => `${show.start}:${show.end}`));
  for (const owned of shows) {
    const object = enclosingTextObject(source, owned.start);
    const objectShows = mraiNativeShowOperatorsIn(source, object.start + 2, object.end - 2);
    if (!objectShows) return refuse("A text object of this paragraph cannot be read show by show.");
    let deletedEnd = owned.end;
    for (const other of objectShows) {
      if (other.start < owned.end) continue;
      if (!ownedKeys.has(`${other.start}:${other.end}`)) break;
      deletedEnd = other.end;
    }
    // (Among other text, for a caller that reads back: what draws on from a
    // show of the box is that caller's to account for -- regularFaceWordsOf
    // lets only blanks do so, and reads them again where they then stand.)
    if (!inPlace && !amongOtherText && !mraiNativeTextAfterShowIsPlacedAfresh(source, owned.end, deletedEnd, object.end)) {
      refuse("The text after a show of this paragraph draws on from where it ends.");
    }
  }
  if (unnamedGlyphs) {
    // Text the box does not hold, standing on a line of it between the
    // line's ends: left as it was beside text in a new look, the line would
    // read as formatted in part. (Where it stands is all that is asked.)
    for (const show of allShows) {
      if (show.codes.length === 0 || ownShows.has(show)) continue;
      const font = stream.fonts.get(show.fontName);
      let glyphs: ReturnType<typeof showGlyphs> | null = null;
      if (font && !show.vertical) {
        try {
          glyphs = showGlyphs(kernel, show, font);
        } catch (error) {
          if (!(error instanceof PdfNativeTransactionError)) throw error;
        }
      }
      const inked = glyphs?.filter((glyph) => !readsBlank(glyph) || Boolean(font?.inkedBlankCodes?.has(glyph.code))) ?? null;
      if (inked && !inked.length) continue;
      const linear = linearOf(show), scale = show.horizontalScale / 100;
      const points = inked ? inked.map((glyph) => applyMatrix(linear, ((glyph.start + glyph.end) / 2) * scale, show.textRise))
        : [applyMatrix(linear, 0, show.textRise)];
      for (const [x, y] of points) {
        const u = x * basis[0] + y * basis[1], v = x * basis[2] + y * basis[3];
        if (lineBaselines.some((baseline, line) => u > lineExtents[line].left && u < lineExtents[line].right &&
          v - baseline > -UNNAMED_LINE_BELOW_EM * bodySize && v - baseline < UNNAMED_LINE_ABOVE_EM * bodySize)) {
          refuseStanding("text it does not hold stands on a line of it");
        }
      }
    }
  }
  const fontOrder: string[] = [];
  for (const show of shows) if (!fontOrder.includes(show.show.fontName)) fontOrder.push(show.show.fontName);
  const firstPlan = await stream.firstPlan(fontOrder, { fontName: head.show.fontName, fontSize: head.show.fontSize });
  const again = await kernel.readStream(firstPlan);
  if (!bytesEqual(again.streamBytes, streamBytes)) return refuseStanding("its content stream changed while it was read");
  const spanStart = Math.min(shows[0].start, ...rules.map((rule) => rule.start));
  const lastObject = enclosingTextObject(source, shows.at(-1)!.start);
  const split = rules.length || inPlace ? null : standingBlockSplit(source, lastObject, shows.at(-1)!.end);
  return Object.freeze({
    selection, items, plans: Object.freeze(items.map(() => null)), firstPlan, streamBytes, source, pageContent,
    shows: Object.freeze(shows), characterShows, basis, blockMatrix,
    layout: Object.freeze({ alignment, frameLeft, frameRight, firstLeft, restLeft, firstBaseline, lineStep, lineCount,
      lineExtents: Object.freeze(lineExtents), bodySize, ...(reach[0].above > 0 ? { firstReach: reach[0].above } : {}) }),
    lineBaselines: Object.freeze([...lineBaselines]),
    pens: Object.freeze([]),
    spanStart,
    rules: Object.freeze(rules.map((rule) => Object.freeze(rule))),
    insertion: inPlace ? shows.at(-1)!.end : split ? split.at : Math.max(blockInsertion({ source, allShows, toBlock, basis, blockMatrix, bodySize,
      spanStart, objectEnd: lastObject.end }), ...rules.map((rule) => rule.end)),
    ...(split ? { blockWrap: Object.freeze({ before: "\nET",
      after: `BT\n${split.lineMatrix.map((value) => formatNumber(value)).join(" ")} Tm\n` }) } : {}),
    pins: firstPlan.request.untouchedObjects ?? Object.freeze([]),
    allShows, twins: Object.freeze([]), twinKeys: new Set<string>(), fonts: stream.fonts, fontOrder: Object.freeze(fontOrder),
    preferred: new Map(fontOrder.map((name) => [name, NO_PREFERRED_CODES])), gapRatio: null,
    keptGaps: new Map(), blankRuns: new Map(), characterInk: new Map(), stacked: new Map(), stackedUnread: false,
    characterRules: new Map(), unreadGaps: Object.freeze([]), unreadGapWidths: Object.freeze([]),
    ...(unnamedOnLine.some((count) => count > 0) ? { unreadGlyphs: true as const } : {}),
    ...(pinned ? { edgeBlankShows: new Set(holding[0].edges ?? []) } : {}),
  });
}

/** How near a reader item's origin the glyph it is pinned to starts
 * (ownShowsByPins): the same pen, read by the reader and by this planner --
 * a twentieth of a point, far under the least a glyph takes the pen on. */
/** How far across a letter's baseline a part of it set as a glyph of its
 * own stands (em): a point or a small letter over or under its body. */
const LETTER_PART_ACROSS_EM = 1;
/** How far past either end of its letter's advance such a part may start (em). */
const LETTER_PART_ALONG_EM = 0.25;
const STANDING_PIN_POINTS = 0.05;
/** How near the end of one reader item the next one starts where the reader
 * cut the second from the first: its own sum of the same numbers. */
const STANDING_CHAIN_POINTS = 0.01;

/** How far from a reader item's ends the glyphs it holds may start and end
 * (itemsAsTheirShows' tolerance: half a point, a twentieth of its size). */
const STANDING_ITEM_END_POINTS = 0.5;
const STANDING_ITEM_END_EM = 0.05;

/** How many characters the reader may read for one glyph its font names no
 * character for (it reads the font program's own name for it: a ligature's
 * letters at most). */
const UNNAMED_GLYPH_MAX_CHARACTERS = 4;
/** How far below and above a line's baseline text counts as standing on the
 * line, as shares of the body size (readParagraphWhereItStands, `unnamedGlyphs`). */
const UNNAMED_LINE_BELOW_EM = 0.25;
const UNNAMED_LINE_ABOVE_EM = 0.75;

/** The item proof's refusal of a show its clip cuts through (list-text-fidelity's
 * MRAI_NATIVE_CLIPPED_SHOW_MESSAGE; named here, as DUPLICATE_SHOWS_MESSAGE is). */
const CLIPPED_SHOW_MESSAGE =
  "The selected text is drawn under a clip that cuts through its glyphs; the reader shows what the page does not.";

/**
 * What the proof of a reader item alone refuses of the show it names, asked
 * of the shows a box owns by their place with no item of it proved
 * (readParagraphWhereItStands, `unnamedGlyphs`): a show whose clip cuts
 * through its ink (the reader shows what the page does not), and a glyph the
 * page draws twice in one place (a format on one of the two would leave the
 * other under it, or over it).
 */
function refuseShowsNoItemProofWouldOwn(
  kernel: MraiNativeParagraphRewriteKernel,
  found: readonly Readonly<{ show: MraiNativeParagraphRewriteShow; font: MraiNativeParagraphRewriteFont;
    glyphs: ReturnType<typeof showGlyphs> }>[],
  allShows: readonly MraiNativeParagraphRewriteShow[],
  fonts: ReadonlyMap<string, MraiNativeParagraphRewriteFont>,
): void {
  for (const own of found) if (kernel.showCutByClip(own.show, own.font)) refuse(CLIPPED_SHOW_MESSAGE);
  // Every inked glyph of the stream, by its font and code: where it starts.
  const inks = (glyph: Readonly<{ text: string | undefined }>) => !/^\s*$/u.test(glyph.text ?? "x");
  const drawn = new Map<string, { x: number; y: number; show: MraiNativeParagraphRewriteShow; at: number }[]>();
  const owned = new Map(found.map((own) => [own.show, own.glyphs] as const));
  for (const show of allShows) {
    if (show.codes.length === 0 || show.vertical) continue;
    const font = fonts.get(show.fontName);
    let glyphs = owned.get(show) ?? null;
    if (!glyphs && font) {
      try {
        glyphs = showGlyphs(kernel, show, font);
      } catch (error) {
        if (!(error instanceof PdfNativeTransactionError)) throw error;
      }
    }
    if (!glyphs) continue;
    const matrix = composeMatrix(show.textMatrix, show.graphicsMatrix), scale = show.horizontalScale / 100;
    for (const [at, glyph] of glyphs.entries()) {
      if (!inks(glyph)) continue;
      const [x, y] = applyMatrix(matrix, glyph.start * scale, show.textRise);
      const key = `${show.fontName}:${glyph.code}`;
      const list = drawn.get(key);
      if (list) list.push({ x, y, show, at });
      else drawn.set(key, [{ x, y, show, at }]);
    }
  }
  for (const own of found) {
    const matrix = composeMatrix(own.show.textMatrix, own.show.graphicsMatrix), scale = own.show.horizontalScale / 100;
    const reach = DUPLICATE_GLYPH_REACH_EM * own.show.fontSize * Math.hypot(matrix[0], matrix[1]);
    for (const [at, glyph] of own.glyphs.entries()) {
      if (!inks(glyph)) continue;
      const [x, y] = applyMatrix(matrix, glyph.start * scale, own.show.textRise);
      if ((drawn.get(`${own.show.fontName}:${glyph.code}`) ?? []).some((other) => (other.show !== own.show || other.at !== at) &&
        Math.hypot(other.x - x, other.y - y) <= reach)) refuse(DUPLICATE_SHOWS_MESSAGE);
    }
  }
}

/** How near one another two drawings of one glyph are one glyph drawn twice,
 * as a share of its size (a fake bold's second pass stands a twentieth off). */
const DUPLICATE_GLYPH_REACH_EM = 0.1;

/**
 * The last check of a format written where the text stands: the stream as
 * the write leaves it is read again, and every glyph the page drew is drawn
 * still -- the same code of the same font, from the same point to the same
 * point (0.01 pt), at the same size -- and no other. One that is not: nothing
 * is written.
 */
function refuseMovedGlyphs(
  kernel: MraiNativeParagraphRewriteKernel,
  paragraph: ReadParagraph,
  replacement: Uint8Array,
  /** Shows the write carries along their line on purpose (a sign that hangs
   * on a leaned line: hangingSignLeans), by the end of each one's span in the
   * stream as it stood, and how far on the page: read again that far on. */
  carried: ReadonlyMap<number, readonly [number, number]> = new Map(),
  /** Glyphs the write sets again on purpose (a bold face set regular where
   * its words stand): by its show and its place in it, the font and code it
   * is drawn in then, how far along the page its origin moves, and its
   * advance in the show's own text space. Read again there, as that. */
  setAgain?: (show: MraiNativeParagraphRewriteShow, glyph: number) => Readonly<{
    fontName: string; code: number; move: readonly [number, number]; advance: number }> | null,
): void {
  const { streamBytes, spanStart, insertion } = paragraph;
  const candidate = new Uint8Array(streamBytes.length - (insertion - spanStart) + replacement.length);
  candidate.set(streamBytes.subarray(0, spanStart));
  candidate.set(replacement, spanStart);
  candidate.set(streamBytes.subarray(insertion), spanStart + replacement.length);
  const cannot = (): never => refuse(MRAI_NATIVE_PARAGRAPH_FORMAT_IN_PLACE_REFUSAL);
  type Drawn = { x0: number; y0: number; x1: number; y1: number; area: number; taken: boolean };
  const glyphsOf = (shows: readonly MraiNativeParagraphRewriteShow[], moved: typeof carried = new Map(), again = false) => {
    const byKey = new Map<string, Drawn[]>();
    let count = 0;
    for (const show of shows) {
      if (show.codes.length === 0) continue;
      const placed = composeMatrix(show.textMatrix, show.graphicsMatrix);
      const [dx, dy] = moved.get(show.end) ?? [0, 0];
      const matrix: Matrix = [placed[0], placed[1], placed[2], placed[3], placed[4] + dx, placed[5] + dy];
      const font = paragraph.fonts.get(show.fontName);
      const scale = show.horizontalScale / 100;
      const area = (matrix[0] * matrix[3] - matrix[1] * matrix[2]) * show.fontSize * show.fontSize * scale;
      const add = (key: string, from: number, to: number, move: readonly [number, number] = [0, 0]) => {
        const [fx, fy] = applyMatrix(matrix, from, show.textRise), [tx, ty] = applyMatrix(matrix, to, show.textRise);
        const [x0, y0, x1, y1] = [fx + move[0], fy + move[1], tx + move[0], ty + move[1]];
        const list = byKey.get(key);
        const entry = { x0, y0, x1, y1, area, taken: false };
        if (list) list.push(entry);
        else byKey.set(key, [entry]);
        count += 1;
      };
      // (A show in a font the page does not prove is no show of the
      // paragraph: it stands as one piece, by its bytes.)
      if (!font || show.vertical) add(`?${show.fontName}:${hexCodes(show.codes)}`, 0, 0);
      else {
        for (const [index, glyph] of showGlyphs(kernel, show, font).entries()) {
          const set = again ? setAgain?.(show, index) : null;
          if (set) add(`${set.fontName}:${set.code}`, glyph.start * scale, (glyph.start + set.advance) * scale, set.move);
          else add(`${show.fontName}:${glyph.code}`, glyph.start * scale, glyph.end * scale);
        }
      }
    }
    return { byKey, count };
  };
  let before: ReturnType<typeof glyphsOf>, after: ReturnType<typeof glyphsOf>;
  try {
    before = glyphsOf(paragraph.allShows, carried, true);
    after = glyphsOf(kernel.readShows(candidate, paragraph.firstPlan, paragraph.fonts));
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError)) throw error;
    return cannot();
  }
  if (before.count !== after.count) cannot();
  for (const [key, list] of before.byKey) {
    const others = after.byKey.get(key);
    if (!others || others.length !== list.length) return cannot();
    for (const glyph of list) {
      const same = others.find((other) => !other.taken &&
        Math.abs(other.x0 - glyph.x0) <= STANDING_GLYPH_TOLERANCE && Math.abs(other.y0 - glyph.y0) <= STANDING_GLYPH_TOLERANCE &&
        Math.abs(other.x1 - glyph.x1) <= STANDING_GLYPH_TOLERANCE && Math.abs(other.y1 - glyph.y1) <= STANDING_GLYPH_TOLERANCE &&
        Math.abs(other.area - glyph.area) <= 1e-4 * Math.max(1, Math.abs(glyph.area)));
      if (!same) return cannot();
      same.taken = true;
    }
  }
}

/** How far a glyph of a format-only edit may stand from where it stood, in
 * points: past it, nothing is written. */
const STANDING_GLYPH_TOLERANCE = 0.01;

/** The formats a character may be asked of where the text stands unread in
 * its order: those that move no glyph and need no character told from another. */
const STANDING_FORMAT_FIELDS: ReadonlySet<string> = new Set(["color", "syntheticBold", "slant", "underline", "strikethrough"]);




















async function planFormatWhereUnprovedTextStands(
  kernel: MraiNativeParagraphRewriteKernel,
  selection: PdfNativeListTextSelectionEvidence,
  resolveFormats: MraiNativeParagraphRewriteFormatsResolver,
  refusal: PdfNativeTransactionError,
): Promise<PdfNativeListTextReplacementPlan> {
  let paragraph: ReadParagraph;
  try {
    paragraph = await readParagraphWhereItStands(kernel, selection, true);
    refuseGlyphsReadAsBlanksBesideLines(paragraph);
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError)) throw error;
    // (Said as the page is: drawn twice, or cut by its clip.)
    if (error.message === DUPLICATE_SHOWS_MESSAGE || error.message === CLIPPED_SHOW_MESSAGE) throw error;
    throw refusal;
  }
  const text = paragraph.selection.item.str;
  const formats = await resolveFormats({ request: paragraph.firstPlan.request, evidence: paragraph.firstPlan.evidence,
    pageContent: paragraph.pageContent, text });
  const asked = formats.characters[0] ?? null;
  const key = JSON.stringify(asked);
  if (formats.characters.length !== characters(text).length || !asked ||
    formats.characters.some((format) => JSON.stringify(format ?? null) !== key) ||
    Object.entries(asked).some(([field, value]) => value !== undefined && !STANDING_FORMAT_FIELDS.has(field)) ||
    formats.alignment !== undefined || formats.leftIndent || formats.lineStep !== undefined ||
    formats.firstLineIndent !== undefined) {
    throw refusal;
  }
  const inked = paragraph.shows.filter((show) => show.inked);
  // A stroke draws no weight on a Type3 glyph (`Tr` does not reach its
  // procedure): said as that is, nothing written.
  if (asked.syntheticBold === true && inked.some((show) => show.font.simpleSubtype === "Type3" &&
    !show.font.type3Pictures && !show.stroked)) refuse(MRAI_NATIVE_TYPE3_FORMAT_MESSAGE);
  refuseBoldTakenOffBoldFace(asked.syntheticBold, inked);
  const sameColor = (left: unknown, right: unknown) => JSON.stringify(left) === JSON.stringify(right);
  if (!inked.some((show) =>
    (asked.color !== undefined && !sameColor(asked.color, show.color)) ||
    strokeAskedChanges(asked.syntheticBold, show) ||
    (asked.slant !== undefined && asked.slant !== show.slant) ||
    (show.inked!.end > show.inked!.start && asked.underline !== undefined && asked.underline !== (show.underline === true)) ||
    (show.inked!.end > show.inked!.start && asked.strikethrough !== undefined &&
      asked.strikethrough !== (show.strikethrough === true)))) {
    refuse(MRAI_NATIVE_STANDING_FORMAT_UNCHANGED_REFUSAL);
  }
  if (asked.slant !== undefined) refuseLeanReadAsOtherRows(paragraph, asked.slant);
  const set = setParagraphWhereItStands(kernel, paragraph, [text], formats);
  if (!set) throw refusal;
  const { request, evidence } = await paragraphEvidence(paragraph, text, set);
  refuseMovedGlyphs(kernel, paragraph, request.replacement);
  return Object.freeze({ request, evidence, properties: null });
}
















function refuseLeanReadAsOtherRows(paragraph: ReadParagraph, slant: number): void {
  const shows = paragraph.shows.filter((show) => show.inked);
  const baseline = (show: OwnedShow) => paragraph.lineBaselines[show.line] + show.rise;
  const heights = (show: OwnedShow) => [show.size * Math.hypot(1, show.slant), show.size * Math.hypot(1, slant)] as const;
  // Whether a distance lies on one side of a share of the height as it
  // stands and on the other as leaned (what stands at the bound, within half
  // a hundredth, is taken to cross it).
  const crosses = (distance: number, was: number, now: number) =>
    was !== now && distance > Math.min(was, now) * 0.995 && distance <= Math.max(was, now) * 1.005;
  for (const [index, show] of shows.entries()) {
    const [was, now] = heights(show);
    const next = shows[index + 1];
    if (next) {
      const [nextWas, nextNow] = heights(next);
      if (crosses(Math.abs(baseline(next) - baseline(show)), READER_LINE_APART_SHARE * Math.max(was, nextWas),
        READER_LINE_APART_SHARE * Math.max(now, nextNow))) refuse(MRAI_NATIVE_PARAGRAPH_READ_BACK_REFUSAL);
    }
    for (const other of shows.slice(index + 1)) {
      if (other.line !== show.line || Math.min(other.inked!.end, show.inked!.end) <= Math.max(other.inked!.start, show.inked!.start)) continue;
      const [otherWas, otherNow] = heights(other);
      if (!crosses(Math.abs(baseline(other) - baseline(show)), PDFJS_VERTICAL_SHIFT_RATIO * Math.min(was, otherWas),
        PDFJS_VERTICAL_SHIFT_RATIO * Math.min(now, otherNow))) continue;
      // Stacked or not, the two read in one order where the page draws the
      // shows that overlap there in the order they stand along the line (a
      // tilde after a word and a space, set a little right of its letter,
      // stands from that space on: «only eO(n)» either way). Only a run of
      // them drawn in another order reads otherwise (the same tilde opening
      // its line).
      const along = shows.filter((entry) => entry.line === show.line).sort((left, right) =>
        left.inked!.start - right.inked!.start || left.start - right.start);
      let from = along.indexOf(show), to = from;
      for (let grown = true; grown;) {
        grown = false;
        const reach = Math.max(...along.slice(from, to + 1).map((entry) => entry.inked!.end));
        const low = Math.min(...along.slice(from, to + 1).map((entry) => entry.inked!.start));
        if (to + 1 < along.length && along[to + 1].inked!.start < reach) { to += 1; grown = true; }
        if (from > 0 && along[from - 1].inked!.end > low) { from -= 1; grown = true; }
      }
      // (Where an item stands along the line to the boxes: where its show
      // starts, or -- the reader's space before it folded into it -- where
      // the show drawn before it on the line ends.)
      const place = (entry: OwnedShow) => {
        const before = shows[shows.indexOf(entry) - 1];
        return before && before.line === entry.line &&
          entry.inked!.start - before.inked!.end >= PDFJS_TRACKING_SPACE_FACTOR * before.size
          ? before.inked!.end : entry.inked!.start;
      };
      const run = along.slice(from, to + 1).sort((left, right) => place(left) - place(right) || left.start - right.start);
      if (run.some((entry, at) => at > 0 && entry.start < run[at - 1].start)) refuse(MRAI_NATIVE_PARAGRAPH_READ_BACK_REFUSAL);
    }
  }
}

/** How far off the baseline of the item before it, as a share of the taller
 * of the two, an item starts a line of its own (selectionVisualLineIndexes). */
const READER_LINE_APART_SHARE = 0.55;

/** The refusal of a rule that would be drawn under a clip that cuts it, or
 * one this writer cannot read: the box would read as ruled and show none. */
export const MRAI_NATIVE_STANDING_RULE_CLIP_MESSAGE =
  "This text is drawn under a clip that a rule under or through it would be cut by, or that cannot be read; the rule was not drawn.";

/** The most pieces one edit's replacement is written as (the transaction's
 * MAX_REPLACEMENT_PIECES; the server's _MAX_REPLACEMENT_PIECES). */
const STANDING_MAX_PIECES = 8192;

type StandingRuleGroup = Readonly<{
  line: number;
  /** The text object the group's shows are drawn in: its rules stand against it. */
  object: Readonly<{ start: number; end: number }>;
  /** Its first show along the line: its size, baseline and colour are the rule's. */
  body: OwnedShow;
  geometry(color?: PdfNativeDeviceColor): Parameters<typeof pdfNativeTextRuleBlock>[0];
}>;

/**
 * Where the rules of a box read where it stands are drawn and read: one under
 * (through) each run of its shows that stand one after another along a line
 * in one text object at one size, on one baseline and in one colour, from
 * where the first of them is inked to where the last is -- the rule the
 * paragraph rewrite draws under such shows and reads back as theirs
 * (paragraphOwnRules), so a paragraph ruled here and set again later takes
 * its rules with it. It stands beside that text object (pdfNativeTextRuleBlock:
 * an underline ends at its `BT`, a strikethrough starts after its `ET`), in
 * the user space and under the clip the text object is drawn in. Shows a hair
 * off one another's baseline (the passes of a fake bold) are one run, ruled
 * on the first's.
 *
 * A show the page stands off the line by where it places it -- the letters
 * of a Nastaleeq word step up along the line, each a show and a text object
 * of its own -- is ruled on the line, with the shows beside it (`ruleRise`:
 * its text rise alone, as the paragraph rewrite rules it). Ruled on its own
 * place, a word's underline climbed with its letters, a piece under each
 * (0339: eight pieces at six heights under 「داللتِ یضمنی ،طم」); the rule
 * of a line is one line under it. `stepped`: the groups as this writer made
 * them then, a show on its own place (steppedStandingRules).
 */
function standingRuleGroups(paragraph: ReadParagraph, stepped = false): readonly StandingRuleGroup[] {
  const riseOf = (show: OwnedShow) => (stepped ? show.rise : show.ruleRise ?? show.rise);
  const { source, shows, basis, blockMatrix, lineBaselines } = paragraph;
  const groups: { key: string; object: Readonly<{ start: number; end: number }>; shows: OwnedShow[] }[] = [];
  const ruled = shows.filter((show) => show.inked && show.inked.end > show.inked.start && !show.readsBlank);
  for (let line = 0; line < lineBaselines.length; line += 1) {
    const along = ruled.filter((show) => show.line === line).sort((left, right) =>
      left.inked!.start - right.inked!.start || left.start - right.start);
    let open: (typeof groups)[number] | null = null;
    for (const show of along) {
      const object = enclosingTextObject(source, show.start);
      // (Its baseline: the line's own where it stands within a twentieth of its size of it.)
      const rise = riseOf(show);
      const baseline = lineBaselines[line] + rise;
      const level = Math.abs(rise) <= STANDING_RULE_BASELINE_EM * show.size ? "line" : String(Math.round(baseline * 1e6));
      const key = JSON.stringify([object.start, Math.round(show.size * 1e6), level, show.color]);
      if (open && open.key === key) open.shows.push(show);
      else {
        open = { key, object: Object.freeze({ start: object.start, end: object.end }), shows: [show] };
        groups.push(open);
      }
    }
  }
  return Object.freeze(groups.map(({ object, shows: own }) => {
    const body = own[0];
    const left = Math.min(...own.map((show) => show.inked!.start)), right = Math.max(...own.map((show) => show.inked!.end));
    // (As the paragraph rewrite reads a rule's run: on its first show's own
    // baseline -- the line's, with its text rise.)
    const across = lineBaselines[body.line] + riseOf(body);
    const inBlock: Matrix = [basis[0], basis[1], basis[2], basis[3],
      left * basis[0] + across * basis[2], left * basis[1] + across * basis[3]];
    const space = body.show.graphicsMatrix;
    const textMatrix = space.every((value, at) => Math.abs(value - blockMatrix[at]) <= 1e-9) ? inBlock
      : composeMatrix(composeMatrix(inBlock, blockMatrix), invertMatrix(space));
    return Object.freeze({ line: body.line, object, body,
      geometry: (color?: PdfNativeDeviceColor) => ({ textMatrix: textMatrix as [number, number, number, number, number, number],
        fontSize: body.size, advance: right - left, color: color ?? body.color }) });
  }));
}

/** How far off its line's baseline a show may stand and be ruled with the
 * shows on it, as a share of its size. */
const STANDING_RULE_BASELINE_EM = 0.05;

/** The rule of a style standing whole under (through) a group, or null. */
function standingRuleOf(streamBytes: Uint8Array, group: StandingRuleGroup, style: "underline" | "strikethrough",
  edges: Readonly<{ start: number; end: number }> = group.object, color?: PdfNativeDeviceColor) {
  let standing: ReturnType<typeof pdfNativeTextRuleStandingAt> = null;
  try {
    standing = pdfNativeTextRuleStandingAt(streamBytes, style === "underline" ? edges.start : edges.end, style, group.geometry(color));
  } catch { standing = null; }
  return standing && pdfNativeTextRuleRangeIsWhole(standing.range) ? standing : null;
}

/**
 * The rules of a style this writer drew under the box before it ruled
 * stepped shows on their line -- a piece under each show where it stands
 * (standingRuleGroups, `stepped`) -- when every one of them stands: the box
 * is ruled so still, reads as ruled, and its next rule is written on the
 * line with those pieces taken off. Null where they do not all stand, or
 * where no show of the box is stepped (the two ways are one).
 */
function steppedStandingRules(paragraph: ReadParagraph, style: "underline" | "strikethrough") {
  if (!paragraph.shows.some((show) => show.ruleRise !== undefined && Math.abs(show.ruleRise - show.rise) > LAYOUT_TOLERANCE)) return null;
  const groups = standingRuleGroups(paragraph, true);
  const rules = groups.map((group) => standingRuleOf(paragraph.streamBytes, group, style));
  return groups.length > 0 && rules.every((rule) => rule !== null)
    ? rules.map((rule) => rule!) : null;
}

/** How a box read where it stands is drawn: stroked, leaned, ruled. */
function standingLookOfParagraph(paragraph: ReadParagraph): MraiNativeStandingTextLook {
  const inked = paragraph.shows.filter((show) => show.inked);
  const groups = standingRuleGroups(paragraph);
  const ruled = (style: "underline" | "strikethrough") => (groups.length > 0 &&
    groups.every((group) => standingRuleOf(paragraph.streamBytes, group, style) !== null)) ||
    steppedStandingRules(paragraph, style) !== null;
  return Object.freeze({
    ...(inked.length > 0 && inked.every((show) => show.stroked) ? { syntheticBold: true as const } : {}),
    ...boldFaceLookOf(inked),
    ...(inked.length > 0 && inked.every((show) => show.slant !== 0) ? { slanted: true as const } : {}),
    ...(ruled("underline") ? { underline: true as const } : {}),
    ...(ruled("strikethrough") ? { strikethrough: true as const } : {}),
  });
}

/** A box read for a format written by state (readParagraphWhereItStands, in
 * place), or null where it is not read so. */
async function readParagraphForStateFormat(
  kernel: MraiNativeParagraphRewriteKernel,
  selection: PdfNativeListTextSelectionEvidence,
  /** Told why the box is not read so. */
  why?: (refusal: PdfNativeTransactionError) => void,
  



  byPins = true,
): Promise<ReadParagraph | null> {
  if (!kernel.standingPageStreams || !kernel.standingStreamContext) return null;
  try {
    const paragraph = await readParagraphWhereItStands(kernel, selection, !pdfTextHasRightToLeftRun(selection.item.str), true);
    refuseGlyphsReadAsBlanksBesideLines(paragraph);
    return paragraph;
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError)) throw error;
    // Not proven the box's by where its glyphs stand (a face that steps its
    // letters off the baseline, a glyph read as a blank beside its line):
    // read by where the reader's items start in the stream instead
    // (ownShowsByPins), which also says whose each glyph read as a blank is.
    let said = error;
    if (byPins && (error.message.startsWith(MRAI_NATIVE_STANDING_FORMAT_REFUSAL) || error.message === MRAI_NATIVE_GLYPHS_READ_AS_BLANKS_MESSAGE)) {
      try {
        const paragraph = await readParagraphWhereItStands(kernel, selection, false, true, true);
        refuseGlyphsReadAsBlanksBesideLines(paragraph);
        return paragraph;
      } catch (pinError) {
        if (!(pinError instanceof PdfNativeTransactionError)) throw pinError;
        if (pinError.message.startsWith(MRAI_NATIVE_STANDING_FORMAT_REFUSAL) &&
          error.message.startsWith(MRAI_NATIVE_STANDING_FORMAT_REFUSAL)) said = pinError;
      }
    }
    
    // again read one, every item of it proved alone -- the same shows, owned
    // on stronger grounds (a paragraph whose letters the reader reads as
    // other characters than its fonts name: by place, a line of it does not
    // read as the characters it draws; pdf-walk-1000 0789, 0661). Its
    // formats are then written by state as any box's, where before they were
    // written by drawing it again. Not a box with fake-bold passes of its
    // own (twins): those are the block's to draw. And not text the reader
    // shows in another order than the page draws it: which glyph of a show
    // is which of the reader's characters that read does not say (a point
    // the reader read into the word before it in the stream would be left
    
    if (!pdfTextHasRightToLeftRun(selection.item.str)) {
      try {
        const proved = await readParagraph(kernel, selection, true);
        if (proved.twins.length === 0) {
          READ_BY_ITEM_PROOFS.add(proved);
          return proved;
        }
      } catch (again) {
        if (!(again instanceof PdfNativeTransactionError)) throw again;
      }
    }
    why?.(said);
    return null;
  }
}

/** The boxes readParagraphForStateFormat read by their items' own proofs
 * alone (its last way). An item proof reads a show under a replacement text
 * as no text of the item's: passes written on a show of such a box
 * (type3ShowPasses) would leave it a box no way of that read reads, and the
 * weight could not be taken off again. */
const READ_BY_ITEM_PROOFS = new WeakSet<object>();

/**
 * The passes that draw a weight on a Type3 show of filled shapes where it
 * stands (standing-state-edit.ts mraiNativeStandingOverstrike): the bytes
 * written right after the show, and how far along its line the first pass
 * stands (text space, before the horizontal scale) --
 * PDF_NATIVE_TYPE3_BOLD_PASS_EM of the size the reader gives the show.
 *
 * Null where the show takes none: its face is not known to draw filled
 * shapes only (`type3Shapes`), the box holds a stretch of it, it moves to its
 * line as it is drawn or runs down the page, or its advance cannot be read.
 * And where the reader would not read the passes back into the show
 * (show-passes' pdfTextItemsWithoutShowPasses, pdfReaderShowAndPasses): it reads a
 * pass as an item of its own only where the pass starts far enough back from
 * the glyph drawn before it -- past its negative-space bound, a fifth of the
 * size -- and reads each pass as the show's own item reads only where the
 * show is one run to it, no two glyphs of it further apart than it goes on
 * in one item (0.6 of the size) nor drawn back past that bound. A show of
 * one narrow glyph, or one set with a wide gap in it, keeps the refusal.
 */
function type3ShowPasses(
  kernel: MraiNativeParagraphRewriteKernel,
  source: string,
  owned: OwnedShow,
): Readonly<{ text: string; offset: number }> | null {
  const show = owned.show;
  if (owned.font.simpleSubtype !== "Type3" || !owned.font.type3Shapes || owned.partial || show.vertical ||
    !/^T[jJ]/u.test(show.operatorKind) || show.trailingKerningAdvance === undefined ||
    !(show.fontSize > 0) || !(show.horizontalScale > 0)) return null;
  // (Drawn along the page's x axis, upright: where the reader measures a
  // glyph against the one before it on the page itself. Turned or leaned, it
  // measures in the glyph's own frame and reads a pass that starts back from
  // the show as a line of its own -- pdf-walk-1000 0945 p2, a bold box
  // leaned: 33 boxes read as 35.)
  const placed = composeMatrix(show.textMatrix, show.graphicsMatrix);
  const along = Math.hypot(placed[0], placed[1]);
  if (owned.slant !== 0 || !(placed[0] > 0) || Math.abs(placed[1]) > 1e-6 * along || Math.abs(placed[2]) > 1e-6 * along) return null;
  let glyphs: ReturnType<typeof showGlyphs>;
  try {
    glyphs = showGlyphs(kernel, show, owned.font);
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError)) throw error;
    return null;
  }
  if (!glyphs.length) return null;
  // Where the pen stands after the show, from where it began: its glyphs,
  // their spacing, and every number of its array.
  const advance = glyphs.at(-1)!.end + show.trailingKerningAdvance;
  const offset = PDF_NATIVE_TYPE3_BOLD_PASS_EM * show.fontSize * Math.max(1, show.readerVerticalScale ?? 1);
  if (!Number.isFinite(advance) || !(offset > 0)) return null;
  // The glyphs the reader sees drawn (it passes over the ones it reads as
  // blanks, and one of no width leaves it where it was), each from its origin
  // to where the glyph alone ends.
  const scale = show.horizontalScale / 100;
  const unit = owned.font.codeUnitBytes;
  const drawn = glyphs.flatMap((glyph) => {
    if (/^\s*$/u.test(glyph.text ?? "x")) return [];
    const width = glyph.end - glyph.start - show.characterSpacing - (unit === 1 && glyph.code === 0x20 ? show.wordSpacing : 0);
    return width !== 0 ? [{ start: glyph.start * scale, end: (glyph.start + width) * scale }] : [];
  });
  if (!drawn.length) return null;
  const size = show.fontSize;
  for (let at = 1; at < drawn.length; at += 1) {
    const gap = drawn[at].start - drawn[at - 1].end;
    if (!(gap > TYPE3_PASS_READER_BACK * size) || !(gap < TYPE3_PASS_READER_ON * size)) return null;
  }
  // (The first pass starts the offset short of the show's whole run back from
  // its last glyph; the second, that run and the offset.)
  if (!(drawn[0].start + offset * scale - drawn.at(-1)!.end < PDFJS_NEGATIVE_SPACE_FACTOR * 1.1 * size)) return null;
  // (The numbers of a `TJ` are thousandths of the size, taken off the pen.)
  let back: string, again: string;
  try {
    back = formatNumber((1_000 * (advance - offset)) / show.fontSize);
    again = formatNumber((1_000 * (advance + offset)) / show.fontSize);
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError)) throw error;
    return null;
  }
  const text = mraiNativeStandingOverstrikeBytes(source.slice(owned.start, owned.end), back, again);
  return text ? Object.freeze({ text, offset }) : null;
}

/** Between two glyphs of a show the reader reads as one run (pdf.js
 * NEGATIVE_SPACE_FACTOR and SPACE_IN_FLOW_MAX_FACTOR, a tenth inside each):
 * past either, it ends the item it is building. */
const TYPE3_PASS_READER_BACK = -0.18;
const TYPE3_PASS_READER_ON = 0.54;





export async function readMraiNativeParagraphStateFormatLook(
  kernel: MraiNativeParagraphRewriteKernel,
  selection: PdfNativeListTextSelectionEvidence,
): Promise<Readonly<{ look: MraiNativeStandingTextLook; /** How many shows of the page are the box's. */ shows: number;
  /** Its Type3 glyphs are pictures all (paragraphHoldsType3PicturesOnly): a weight asked of it is written on its words. */
  type3Pictures: boolean;
  /** Its Type3 letters take a weight where they stand (paragraphType3LettersTakeWeight). */
  type3Weight: boolean;
  /** A weight drew a Type3 show of it again after it (the show's `overstruck`): what draws the box again is refused it. */
  type3Passes: boolean }> | null> {
  const paragraph = await readParagraphForStateFormat(kernel, selection);
  return paragraph ? Object.freeze({ look: standingLookOfParagraph(paragraph), shows: paragraph.shows.length,
    type3Pictures: paragraphHoldsType3PicturesOnly(paragraph), type3Weight: paragraphType3LettersTakeWeight(kernel, paragraph),
    type3Passes: paragraph.shows.some((show) => Boolean(show.show.overstruck)) }) : null;
}

async function readStandingStateLook(
  kernel: MraiNativeParagraphRewriteKernel,
  selection: PdfNativeListTextSelectionEvidence,
): Promise<MraiNativeStandingTextLook | null> {
  const paragraph = await readParagraphForStateFormat(kernel, selection);
  return paragraph ? standingLookOfParagraph(paragraph) : null;
}




function refusalOfRedrawOnly(message: string): boolean {
  return message === MRAI_NATIVE_PARAGRAPH_TOO_LARGE_MESSAGE ||
    message.startsWith("The paragraph's shows lie too far apart in their stream") ||
    message.startsWith(MRAI_NATIVE_STANDING_FORMAT_REFUSAL) ||
    /^Reader item \d+ of the paragraph does not prove alone: /u.test(message) ||
    message === DUPLICATE_SHOWS_MESSAGE || message === CLIPPED_SHOW_MESSAGE ||
    // (No text is set again from what the reader read: a glyph read as a
    
    // itself -- readParagraphForStateFormat.)
    message === MRAI_NATIVE_GLYPHS_READ_AS_BLANKS_MESSAGE;
}


























async function planFormatByStateWhereTextStands(
  kernel: MraiNativeParagraphRewriteKernel,
  selection: PdfNativeListTextSelectionEvidence,
  resolveFormats: MraiNativeParagraphRewriteFormatsResolver,
  refusal: PdfNativeTransactionError,
  


  first = false,
): Promise<PdfNativeListTextReplacementPlan> {
  // (Refused by both for want of the same proof -- that the page's shows are
  // this box and nothing else -- the reason said is this read's: what stood in
  // the way of drawing the box again does not stand in the way here.)
  let unread: PdfNativeTransactionError | null = null;
  const paragraph = await readParagraphForStateFormat(kernel, selection, (error) => { unread = error; }, !first);
  if (!paragraph) {
    const own = unread as PdfNativeTransactionError | null;
    throw own && own.message.startsWith(MRAI_NATIVE_STANDING_FORMAT_REFUSAL) &&
      refusal.message.startsWith(MRAI_NATIVE_STANDING_FORMAT_REFUSAL) ? own : refusal;
  }
  const context = await kernel.standingStreamContext!(selection, paragraph.firstPlan);
  if (!context) throw refusal;
  const text = paragraph.selection.item.str;
  const formats = await resolveFormats({ request: paragraph.firstPlan.request, evidence: paragraph.firstPlan.evidence,
    pageContent: paragraph.pageContent, text });
  const askedOfAll = formats.characters[0] ?? null;
  const key = JSON.stringify(askedOfAll);
  if (formats.characters.length !== characters(text).length || !askedOfAll ||
    formats.characters.some((format) => JSON.stringify(format ?? null) !== key) ||
    Object.entries(askedOfAll).some(([field, value]) => value !== undefined && !STANDING_FORMAT_FIELDS.has(field)) ||
    formats.alignment !== undefined || formats.leftIndent || formats.lineStep !== undefined ||
    formats.firstLineIndent !== undefined) {
    
    // no longer read because a weight drew a Type3 show of it again: said as
    // that, with what takes it off.)
    if (paragraph.shows.some((show) => show.show.overstruck)) refuse(MRAI_NATIVE_TYPE3_BOLD_STANDS_MESSAGE);
    throw refusal;
  }
  // The colour asked, as the stream will state it (colorAsWritten): it is
  // compared below with the colours the page states -- whether a show takes
  // it at all, and, read again, whether every glyph of the box is drawn in
  // it. Compared as asked, a component the stream spells in fewer digits (a
  // palette's 26/255, written 0.101961) never read back as itself, and every
  // colour but those whose components are all whole fifths was refused as not
  // written where the text stands (#cc1a1a on a paragraph a browser printed:
  
  const asked = askedOfAll.color ? Object.freeze({ ...askedOfAll, color: colorAsWritten(askedOfAll.color) }) : askedOfAll;
  const { source, streamBytes } = paragraph;
  const inked = paragraph.shows.filter((show) => show.inked);
  // (A weight that is a face's own is not taken off by state: said as that.)
  refuseBoldTakenOffBoldFace(asked.syntheticBold, inked);
  // A bold asked strokes the shows that are not bold as they stand -- where
  // one of them draws ink. With every inked show stroked or in a bold face
  // already, nothing is: the blanks between its words (a space shown in the
  // family's regular face) are not stroked for a weight no glyph would show.
  const strokeAsked = asked.syntheticBold === true && inked.some((show) => strokeAskedChanges(true, show));
  // A lean: the shear each show of the box takes, in the text space it is
  
  // that draw a box again where it stands lean it -- every show of it at one
  // lean now and at the lean asked after, whatever its face -- with no rise
  // of its own (a raised glyph stands along the leaned up axis, and would
  // move); and only where the reader reads the leaned box as it read it
  
  
  // answers. A show the box holds a stretch of (a line that runs on into the
  // column beside it) is cut where the stretch starts and ends
  // (standingShowCuts), and the stretch leans as a show of its own: the
  // matrix each piece starts under is set between the pieces.
  const leans = new Map<number, number>();
  /** The signs that hang on the box's lines and lean with them (by the end of each one's span). */
  const signs = new Map<number, Readonly<{ shear: number; hang: number; moved: readonly [number, number] }>>();
  const leanCuts: StandingEdit[] = [];
  const leanPieces = new Map<number, Readonly<{ pieces: number; own: number }>>();
  if (asked.slant !== undefined && paragraph.shows.some((show) => Math.abs(show.slant - asked.slant!) > 1e-9)) {
    const slant = asked.slant;
    const standing = paragraph.shows[0].slant;
    // (A box read by where the reader's items start, ownShowsByPins, is
    // leaned so as well: its shows are the ones its other formats are
    // written on where they stand, and what the reader would read of them
    // leaned is asked of the same guards.)
    if (paragraph.shows.some((show) => Math.abs(show.slant - standing) > 1e-9 || show.show.textRise !== 0 ||
      !(show.show.horizontalScale > 0) || !(show.horizontalScale > 0))) throw refusal;
    for (const show of paragraph.shows) {
      if (!show.partial) continue;
      // (Two stretches of one show: not cut so.)
      if (leanPieces.has(show.end)) throw refusal;
      const { codeStart, codeEnd } = show.partial;
      const cuts = standingShowCuts(source, show, [...(codeStart > 0 ? [codeStart] : []),
        ...(codeEnd < show.show.codes.length ? [codeEnd] : [])]);
      if (!cuts) throw refusal;
      for (const cut of cuts) leanCuts.push({ start: cut.offset, end: cut.offset, text: `${cut.close} ${cut.open}` });
      leanPieces.set(show.end, { pieces: 1 + cuts.length, own: codeStart > 0 ? 1 : 0 });
    }
    // (Read back by the caller with the reader itself, the forecasts are not asked.)
    const forecast = resolveFormats.leanReadByReader !== true;
    if (forecast) refuseLeanReadAsOtherRows(paragraph, slant);
    const guard = leanReadBackGuard(kernel, paragraph);
    for (let line = 0; forecast && line < paragraph.lineBaselines.length; line += 1) {
      for (const show of paragraph.shows) {
        if (show.line !== line) continue;
        if (show.inked && show.inked.end > show.inked.start) {
          guard.refuseLeanReadOtherwise(show, slant, [show.inked.start, show.inked.end], paragraph.lineBaselines[line] + show.rise);
        } else guard.blankDrawn();
      }
    }
    // (Its slant is read across its up axis: a show set wider than it is tall takes less shear for the same lean.)
    for (const show of paragraph.shows) leans.set(show.end, (slant - show.slant) * (show.show.horizontalScale / show.horizontalScale));
    // A sign that hangs on one of its lines leans with the line (hangingSignLeans).
    for (const [end, sign] of hangingSignLeans(paragraph, standing, slant)) {
      leans.set(end, sign.shear);
      signs.set(end, sign);
    }
  }
  // A weight asked of a Type3 show that is no picture. The stroke the other
  // shows take draws nothing on it (`Tr` does not reach its procedure): its
  // weight is the show drawn again a little along the line, right after it
  // (type3ShowPasses), and it is set under the stroke with the others so
  // that it reads as bold wherever a weight is read. Where the show takes
  // no such passes, said as it was, nothing written. Taken off, the passes
  // go whole and the stream is the bytes it was.
  const passEdits: StandingEdit[] = [];
  /** By the end of each show's span: the passes that stand after it once this is written (null: taken off). */
  const passAfter = new Map<number, Readonly<{ offset: number }> | null>();
  if (asked.syntheticBold === true) {
    for (const show of inked) {
      if (show.font.simpleSubtype !== "Type3" || show.font.type3Pictures || show.stroked || show.show.overstruck) continue;
      const pass = READ_BY_ITEM_PROOFS.has(paragraph) ? null : type3ShowPasses(kernel, source, show);
      if (!pass) refuse(MRAI_NATIVE_TYPE3_FORMAT_MESSAGE);
      passEdits.push({ start: show.end, end: show.end, text: pass!.text });
      passAfter.set(show.end, Object.freeze({ offset: pass!.offset }));
    }
  } else if (asked.syntheticBold === false) {
    for (const show of paragraph.shows) {
      if (!show.show.overstruck) continue;
      passEdits.push({ start: show.end, end: show.show.overstruck.end, text: "" });
      passAfter.set(show.end, null);
    }
  }
  // The stream's paint state, as the page's content before it leaves it.
  const names = new Set<string>();
  for (const stream of [...context.prefix, source]) {
    for (const match of stream.matchAll(/(\/[^\x00\t\n\f\r /<>[\]()%{}]+)[\x00\t\n\f\r ]+gs(?![A-Za-z0-9*'"])/gu)) names.add(match[1]);
  }
  const lineWidths = await context.lineWidths([...names]);
  const walked = (stream: string) => walkStandingStream({ source: stream, prefix: context.prefix,
    inherits: !paragraph.pageContent, lineWidths });
  const walk = walked(source);
  const byEnd = new Map(walk.shows.map((show) => [show.end, show] as const));
  // What each show of the box is asked to stand under.
  const sameColor = (left: PdfNativeDeviceColor, right: PdfNativeDeviceColor) => JSON.stringify(left) === JSON.stringify(right);
  const targets = new Map<number, StandingTarget>();
  const drawnAs = new Map<number, Readonly<{ mode: number; color: PdfNativeDeviceColor }>>();
  // A show the box holds a stretch of: cut in two where the stretch starts
  // and ends, the state written between the pieces (standingShowCuts).
  const cutEdits: StandingEdit[] = [];
  const cutOnly: StandingEdit[] = [];
  const cutTargets: { owned: OwnedShow; target: StandingTarget; pieces: number; own: number }[] = [];
  for (const owned of paragraph.shows) {
    const standing = byEnd.get(owned.end);
    if (!standing || standing.start !== owned.start) throw refusal;
    const target: { -readonly [K in keyof StandingTarget]: StandingTarget[K] } = {};
    let fill = standing.state.fill;
    const recolored = asked.color !== undefined && !sameColor(asked.color, owned.color);
    if (recolored) {
      fill = colorOperator(asked.color!);
      target.fill = fill;
    }
    const stroked = standing.state.tr === "2";
    // (A show in a bold face is bold as it stands: strokeAskedChanges.)
    const boldAsIs = asked.syntheticBold === true && !stroked && (!strokeAsked || standsInBoldFace(owned));
    if (asked.syntheticBold === true && !stroked && !boldAsIs) {
      if (fill === null) return refuse(MRAI_NATIVE_STANDING_STATE_UNKNOWN_MESSAGE);
      const matrix = owned.show.textMatrix;
      const width = pdfNativeSyntheticBoldStrokeWidth(Math.abs(owned.show.fontSize), [matrix[0], matrix[1], matrix[2], matrix[3]]);
      if (!(width > 0)) throw refusal;
      target.tr = "2";
      target.w = `${formatNumber(width)} w`;
      target.stroke = standingStrokeSpelling(fill);
    } else if (asked.syntheticBold === false && stroked) {
      target.tr = "0";
      target.w = "free";
      target.stroke = "free";
    } else if (stroked && recolored && standing.state.fill !== null &&
      standing.state.stroke === standingStrokeSpelling(standing.state.fill)) {
      // (A stroke in the fill's colour is this writer's weight: it keeps the text's colour.)
      target.stroke = standingStrokeSpelling(fill!);
    }
    // (A show with passes after it is one the box holds whole: cut, they would be the passes of neither piece.)
    if (owned.partial && owned.show.overstruck) throw refusal;
    if (owned.partial && Object.keys(target).length) {
      const length = owned.show.codes.length;
      const { codeStart, codeEnd } = owned.partial;
      const cuts = standingShowCuts(source, owned, [...(codeStart > 0 ? [codeStart] : []), ...(codeEnd < length ? [codeEnd] : [])]);
      if (!cuts) throw refusal;
      // (Its own state before and after: what stands at the show.)
      const set: string[] = [], put: string[] = [];
      for (const kind of ["tr", "w", "stroke", "fill"] as const) {
        const asked = target[kind];
        if (asked === undefined || asked === "free" || asked === standing.state[kind]) continue;
        const back = standing.state[kind];
        if (back === null) return refuse(MRAI_NATIVE_STANDING_STATE_UNKNOWN_MESSAGE);
        set.push(kind === "tr" ? `${asked} Tr` : asked);
        put.push(kind === "tr" ? `${back} Tr` : back);
      }
      if (set.length) {
        const [head, tail] = [codeStart > 0 ? cuts[0] : null, codeEnd < length ? cuts.at(-1)! : null];
        if (head) {
          cutEdits.push({ start: head.offset, end: head.offset, text: `${head.close} ${set.join(" ")} ${head.open}` });
          cutOnly.push({ start: head.offset, end: head.offset, text: `${head.close} ${head.open}` });
        } else cutEdits.push({ start: owned.start, end: owned.start, text: `${set.join(" ")} ` });
        if (tail) {
          cutEdits.push({ start: tail.offset, end: tail.offset, text: `${tail.close} ${put.join(" ")} ${tail.open}` });
          cutOnly.push({ start: tail.offset, end: tail.offset, text: `${tail.close} ${tail.open}` });
        } else cutEdits.push({ start: owned.end, end: owned.end, text: ` ${put.join(" ")}` });
        cutTargets.push({ owned, target: Object.freeze(target), pieces: 1 + (head ? 1 : 0) + (tail ? 1 : 0), own: head ? 1 : 0 });
      }
    } else if (Object.keys(target).length) targets.set(owned.end, Object.freeze(target));
    drawnAs.set(owned.end, Object.freeze({
      mode: asked.syntheticBold === undefined || boldAsIs ? owned.show.renderingMode : asked.syntheticBold ? 2 : 0,
      color: asked.color ?? owned.color }));
  }
  const stateEdits = normalizeStandingEdits([...(targets.size ? planStandingStateEdits(walk, targets) : []), ...cutEdits]);
  if (stateEdits.length) {
    // Read against the stream with its shows cut and nothing else written
    // (a cut writes no state): the same shows in the same order, each of the
    // box's -- a cut show's own piece among them -- under what was asked.
    const cut = walked(applyStandingEdits(source, normalizeStandingEdits(cutOnly)));
    const expected = new Map<number, StandingTarget>();
    const pieceOf = new Map(cutTargets.map((entry) => [entry.owned.end, entry] as const));
    let order = 0;
    for (const show of walk.shows) {
      const entry = pieceOf.get(show.end);
      const target = entry ? entry.target : targets.get(show.end);
      const at = cut.shows[order + (entry?.own ?? 0)];
      if (!at) return refuse(MRAI_NATIVE_PARAGRAPH_FORMAT_IN_PLACE_REFUSAL);
      if (target) expected.set(at.end, target);
      order += entry?.pieces ?? 1;
    }
    if (order !== cut.shows.length || cut.paints.length !== walk.paints.length + (cut.shows.length - walk.shows.length)) {
      refuse(MRAI_NATIVE_PARAGRAPH_FORMAT_IN_PLACE_REFUSAL);
    }
    const differs = standingStateTraceDiffers(cut, walked(applyStandingEdits(source, stateEdits)), expected);
    if (differs) refuse(MRAI_NATIVE_PARAGRAPH_FORMAT_IN_PLACE_REFUSAL);
  }
  // Its rules: beside each text object of each line, in the text's colour.
  const groups = standingRuleGroups(paragraph);
  const steppedRules = { underline: steppedStandingRules(paragraph, "underline"),
    strikethrough: steppedStandingRules(paragraph, "strikethrough") } as const;
  if (first) {
    
    // again read one (a later edit of its words is theirs). They own a
    // paragraph's rules only where every rule standing on its lines is one of
    // its runs' (paragraphOwnRules), and take a link's own underline away
    // with a rule asked: so every rule standing on the box's lines must be
    // one read here, and no other writer's rule may stand under a box a
    // rule is asked of.
    const standing = new Set([...groups.flatMap((group) => (["underline", "strikethrough"] as const).flatMap((style) =>
      standingRuleOf(streamBytes, group, style)?.start ?? [])),
      ...(steppedRules.underline ?? []).map((rule) => rule.start), ...(steppedRules.strikethrough ?? []).map((rule) => rule.start)]);
    const beside = paragraphRuleSpansOnItsLines(streamBytes, source, paragraph.shows, paragraph.basis, paragraph.lineBaselines);
    if (beside.some((rule) => !standing.has(rule.start))) throw refusal;
    if (asked.underline !== undefined || asked.strikethrough !== undefined || (asked.color !== undefined && beside.length)) {
      
      if (paragraph.shows.some((show) => !show.show.graphicsMatrix.every((value, at) =>
        Math.abs(value - paragraph.blockMatrix[at]) <= 1e-9))) throw refusal;
      const linked = paragraphLinkRules({ source, shows: paragraph.shows, basis: paragraph.basis, blockMatrix: paragraph.blockMatrix,
        skip: beside, lineSizes: paragraph.lineBaselines.map((_, line) => Math.max(0, ...paragraph.shows
          .filter((show) => show.line === line).map((show) => show.size))) });
      if (linked && linked.marks.size > 0) throw refusal;
    }
  }
  const ruleEdits: StandingEdit[] = [];
  const ruledAfter: { group: StandingRuleGroup; style: "underline" | "strikethrough"; on: boolean }[] = [];
  // Rules drawn a piece under each stepped show (steppedStandingRules) are
  // taken off with whatever is written, and the rule drawn on the line in
  // their place (or not, where it is taken off): the pieces that are a
  // group's rule on the line as well stay where they stand.
  for (const style of ["underline", "strikethrough"] as const) {
    const pieces = steppedRules[style];
    if (!pieces) continue;
    const onLine = new Set(groups.flatMap((group) => standingRuleOf(streamBytes, group, style)?.start ?? []));
    const taken = new Set<number>();
    for (const piece of pieces) {
      if (onLine.has(piece.start) || taken.has(piece.start)) continue;
      taken.add(piece.start);
      ruleEdits.push({ start: piece.start, end: piece.end, text: "" });
    }
  }
  for (const group of groups) {
    for (const style of ["underline", "strikethrough"] as const) {
      const standing = standingRuleOf(streamBytes, group, style);
      const on = asked[style] ?? (standing !== null || steppedRules[style] !== null);
      const recolored = asked.color !== undefined && !sameColor(asked.color, group.body.color);
      ruledAfter.push({ group, style, on });
      if (on === (standing !== null) && !(on && recolored)) continue;
      if (standing) ruleEdits.push({ start: standing.start, end: standing.end, text: "" });
      if (!on) continue;
      const block = LATIN1.decode(pdfNativeTextRuleBlock(group.geometry(asked.color), style));
      // Drawn in the user space and under the clip of the text object it stands against.
      const at = byEnd.get(group.body.end)!;
      const space = group.body.show.graphicsMatrix;
      const numbers = [...block.matchAll(/-?\d+(?:\.\d+)?/gu)].map((match) => Number(match[0])).slice(-8);
      if (at.clip === "unknown" || numbers.length !== 8 ||
        !space.every((value, index) => Math.abs(value - at.ctm[index]) <= 1e-6 * Math.max(1, Math.abs(value)))) {
        refuse(MRAI_NATIVE_STANDING_RULE_CLIP_MESSAGE);
      }
      const clip = at.clip;
      if (clip) {
        // The clip may cut the rule where it cuts the text over it -- along
        // the line (a printer driver draws a line again in every band it
        // crosses, each time under that band's clip: each pass takes its own
        // rule, cut as its glyphs are) -- and nowhere across it: a rule
        // outside the clip's height would be read as drawn and show nothing.
        const corners = [0, 2, 4, 6].map((corner) => applyMatrix(at.ctm, numbers[corner], numbers[corner + 1]));
        const xs = corners.map((point) => point[0]), ys = corners.map((point) => point[1]);
        const level = Math.abs(corners[0][1] - corners[1][1]) <= 0.01, upright = Math.abs(corners[0][0] - corners[1][0]) <= 0.01;
        const insideAcross = level ? Math.min(...ys) >= clip[1] - 0.01 && Math.max(...ys) <= clip[3] + 0.01
          : upright ? Math.min(...xs) >= clip[0] - 0.01 && Math.max(...xs) <= clip[2] + 0.01
            : corners.every(([x, y]) => x >= clip[0] - 0.01 && x <= clip[2] + 0.01 && y >= clip[1] - 0.01 && y <= clip[3] + 0.01);
        const shows = Math.max(...xs) > clip[0] && Math.min(...xs) < clip[2] && Math.max(...ys) > clip[1] && Math.min(...ys) < clip[3];
        if (!insideAcross || !shows) refuse(MRAI_NATIVE_STANDING_RULE_CLIP_MESSAGE);
      }
      const edge = style === "underline" ? group.object.start : group.object.end;
      ruleEdits.push({ start: edge, end: edge, text: block });
    }
  }
  let leanEdits: readonly StandingEdit[] = [];
  // The shows the lean is read back against, and the shear each takes: the
  // stream's as it stands, or as it stands with the box's stretches cut out
  // of the shows they are part of.
  let leanedShows: readonly MraiNativeParagraphRewriteShow[] = paragraph.allShows;
  let leanOf: ReadonlyMap<number, number> = leans;
  // The signs leaned with their lines, as the lean is written: the ones the
  // page then draws carried along (by the end of each one's span in the
  // stream as it stood, and in the shows the lean is read back against).
  let carried: ReadonlyMap<number, readonly [number, number]> = new Map();
  let leanedSigns: ReadonlySet<number> = new Set();
  if (leans.size) {
    const leadingKnown = context.prefix.length === 0 && Boolean(paragraph.pageContent);
    // Where the show before a matrix leaves the pen, read from the stream
    // without that matrix: what a matrix that may set nothing once the lean
    // is off is held against (planStandingLeanEdits).
    const penWithout = (stream: string) => (from: number, to: number): Matrix | null => {
      try {
        const without = kernel.readShows(asciiBytes(stream.slice(0, from) + stream.slice(to)), paragraph.firstPlan, paragraph.fonts);
        return without.find((show) => show.start >= from)?.textMatrix ?? null;
      } catch (error) {
        if (!(error instanceof PdfNativeTransactionError)) throw error;
        return null;
      }
    };
    /** The lean as this stream takes it, of the box's shows and (`withSigns`) of the signs that hang on its lines. */
    const leanPlanned = (withSigns: boolean) => {
      const taking = new Map([...leans].filter(([end]) => withSigns || !signs.has(end)));
      const hanging = new Map(withSigns ? [...signs].map(([end, sign]) => [end, sign.hang] as const) : []);
      if (leanCuts.length === 0) {
        return { shows: paragraph.allShows, leanOf: taking as ReadonlyMap<number, number>, signs: new Set(hanging.keys()),
          planned: planStandingLeanEdits({ walk, owned: taking, hangs: hanging,
            pens: new Map(paragraph.allShows.map((show) => [show.end, show.textMatrix] as const)), leadingKnown, format: formatNumber,
            penWithout: penWithout(source) }) };
      }
      // (A weight or a colour asked with it cuts the same shows for itself: the two are not written as one.)
      if (stateEdits.length) throw refusal;
      const cuts = normalizeStandingEdits(leanCuts);
      const cutSource = applyStandingEdits(source, cuts);
      const cutWalk = walked(cutSource);
      const cutShows = kernel.readShows(asciiBytes(cutSource), paragraph.firstPlan, paragraph.fonts);
      const cutLeans = new Map<number, number>();
      const cutHangs = new Map<number, number>();
      let order = 0;
      for (const show of walk.shows) {
        const piece = leanPieces.get(show.end);
        const own = cutWalk.shows[order + (piece?.own ?? 0)];
        if (!own) throw refusal;
        const lean = taking.get(show.end);
        if (lean !== undefined) cutLeans.set(own.end, lean);
        const hang = hanging.get(show.end);
        if (hang !== undefined) cutHangs.set(own.end, hang);
        order += piece?.pieces ?? 1;
      }
      if (order !== cutWalk.shows.length) throw refusal;
      const inCut = planStandingLeanEdits({ walk: cutWalk, owned: cutLeans, hangs: cutHangs,
        pens: new Map(cutShows.map((show) => [show.end, show.textMatrix] as const)), leadingKnown, format: formatNumber,
        penWithout: penWithout(cutSource) });
      return { shows: cutShows as readonly MraiNativeParagraphRewriteShow[], leanOf: cutLeans as ReadonlyMap<number, number>,
        signs: new Set(cutHangs.keys()),
        planned: typeof inCut === "string" ? inCut : composeStandingCutEdits(cuts, inCut) ?? "The cuts and the lean run into one another." };
    };
    // With the signs that hang on its lines where the stream places each by a
    // move or a matrix of its own; where it does not, without them (and the
    // lean is then refused for the sign left upright, below).
    let lean = leanPlanned(signs.size > 0);
    if (typeof lean.planned === "string" && signs.size > 0) lean = leanPlanned(false);
    
    if (typeof lean.planned === "string") throw refusal;
    leanEdits = lean.planned;
    leanedShows = lean.shows;
    leanOf = lean.leanOf;
    leanedSigns = lean.signs;
    if (lean.signs.size) carried = new Map([...signs].map(([end, sign]) => [end, sign.moved] as const));
    // (A lean taken off gives the page back as it was read before the lean: never refused for how it reads.)
    if (asked.slant !== 0 && resolveFormats.leanReadByReader !== true) {
      refuseLeanAmongOtherTextOnItsLines({ allShows: leanedShows }, leanOf, leanedSigns);
    }
  }
  // A lean and the passes of a show do not go together (type3ShowPasses): a
  // box that would stand leaned with passes after a show of it is refused,
  // and said as that -- the bold is taken off first.
  if (leans.size && paragraph.shows.some((show) =>
    passAfter.has(show.end) ? passAfter.get(show.end) !== null : Boolean(show.show.overstruck))) {
    const kept = new PdfNativeTransactionError("unsupported-document", MRAI_NATIVE_TYPE3_BOLD_STANDS_MESSAGE);
    FINAL_STANDING_REFUSALS.add(kept);
    throw kept;
  }
  // (The passes of a show are written right after it, before the state put back there.)
  if (passEdits.length && leanCuts.length) throw refusal;
  const edits = normalizeStandingEdits([...passEdits, ...stateEdits, ...leanEdits, ...ruleEdits]);
  if (!edits.length) refuse(MRAI_NATIVE_STANDING_FORMAT_UNCHANGED_REFUSAL);
  // (The span keeps a byte of its own: one that only took a rule away, or
  // only wrote one, is written as a piece of the stream beside it.)
  let spanStart = edits[0].start, spanEnd = edits.at(-1)!.end;
  // (And the show its passes are written after, or taken away from: the
  // record's contract reads them against the show they stand after.)
  const passShow = paragraph.shows.find((show) => passAfter.has(show.end) && show.end === edits[0].start);
  if (passShow) spanStart = passShow.start;
  if (!standingEditPieces(edits, spanStart, spanEnd).some((piece) => "spanStart" in piece)) {
    if (spanEnd < source.length) spanEnd += 1;
    else if (spanStart > 0) spanStart -= 1;
    else throw refusal;
  }
  // (A lean's last edit may be an operand of a positioning operator: the span
  // runs on to that operator, so what the operand is can be read from the
  // span alone -- the record's contract reads it there.)
  if (mraiNativeStandingBytesAreANumber(source.slice(edits.at(-1)!.start, edits.at(-1)!.end)) && edits.at(-1)!.end === spanEnd) {
    const rest = /^(?:[\x00\t\n\f\r ]+[+-]?(?:\d+\.?\d*|\.\d+)){1,3}[\x00\t\n\f\r ]+T[mdD]/u.exec(source.slice(spanEnd, spanEnd + 256));
    if (!rest) refuse(MRAI_NATIVE_PARAGRAPH_FORMAT_IN_PLACE_REFUSAL);
    spanEnd += rest![0].length;
  }
  // (What the record's own check holds it to: state operators and rules written, and nothing else taken away.)
  // (And for a lean: a number written for a number -- an operand of a positioning operator -- and whole
  // positioning operators, which the state grammar holds.)
  if (edits.some((edit) => {
    const taken = source.slice(edit.start, edit.end);
    if (mraiNativeStandingBytesAreANumber(taken) || mraiNativeStandingBytesAreANumber(edit.text)) {
      return !(mraiNativeStandingBytesAreANumber(taken) && mraiNativeStandingBytesAreANumber(edit.text));
    }
    // (A show cut in two for a lean is one show again with the lean off: the cut, and nothing in its place.)
    if (edit.text === "" && mraiNativeStandingBytesCutAShow(taken)) return false;
    // (The passes of a Type3 show: written whole after it, taken away whole.)
    if (edit.text === "" && mraiNativeStandingBytesOverstrikeAShow(taken)) return false;
    return !(mraiNativeStandingBytesAreStateOrRules(edit.text) || mraiNativeStandingBytesCutAShow(edit.text) ||
      mraiNativeStandingBytesOverstrikeAShow(edit.text)) || !mraiNativeStandingBytesAreStateOrRules(taken);
  })) refuse(MRAI_NATIVE_PARAGRAPH_FORMAT_IN_PLACE_REFUSAL);
  const pieces = standingEditPieces(edits, spanStart, spanEnd);
  const written = edits.reduce((sum, edit) => sum + edit.text.length, 0);
  if (written > PARAGRAPH_REWRITE_MAX_REPLACEMENT_BYTES || pieces.length > STANDING_MAX_PIECES) {
    refuse(MRAI_NATIVE_PARAGRAPH_TOO_LARGE_MESSAGE);
  }
  const candidate = applyStandingEdits(source, edits);
  const replacement = asciiBytes(candidate.slice(spanStart, candidate.length - (source.length - spanEnd)));
  const cannot = (): never => refuse(MRAI_NATIVE_PARAGRAPH_FORMAT_IN_PLACE_REFUSAL);
  // Read again: every glyph where it was; the box's shows drawn as asked and
  // no other show otherwise; each rule where it was drawn, and none left.
  refuseMovedGlyphs(kernel, { ...paragraph, spanStart, insertion: spanEnd }, replacement, carried);
  const candidateBytes = asciiBytes(candidate);
  const after = kernel.readShows(candidateBytes, paragraph.firstPlan, paragraph.fonts);
  // (Glyph by glyph, in the order the stream draws them: a show cut in two
  // draws the same glyphs in the same order.)
  const stretchOf = new Map(paragraph.shows.map((show) => [show.end, show.partial] as const));
  const drawnOf = (shows: readonly MraiNativeParagraphRewriteShow[], expected: boolean) => shows.flatMap((show) => {
    const unit = paragraph.fonts.get(show.fontName)?.codeUnitBytes ?? 0;
    const count = unit ? Math.floor(show.codes.length / unit) : 1;
    const want = expected ? drawnAs.get(show.end) : undefined;
    const stretch = expected ? stretchOf.get(show.end) : undefined;
    // (Drawn again: a show with passes after it -- the ones this wrote them
    // after, less the ones it took theirs from, and every other as it was.)
    const twice = expected && passAfter.has(show.end) ? passAfter.get(show.end) !== null : Boolean(show.overstruck);
    return Array.from({ length: count }, (_, glyph) => {
      const own = want && (!stretch || !unit || (glyph * unit >= stretch.codeStart && glyph * unit < stretch.codeEnd)) ? want : undefined;
      return `${own?.mode ?? show.renderingMode}|${show.deviceColorKnown ? JSON.stringify(own?.color ?? show.deviceColor) : "?"}|${twice ? 2 : 1}`;
    });
  });
  const [wanted, got] = [drawnOf(paragraph.allShows, true), drawnOf(after, false)];
  if (wanted.length !== got.length || wanted.some((value, glyph) => value !== got[glyph])) cannot();
  // The passes written stand as far along the line as they were set (read
  // from the bytes, the second on the show and the pen after it where the
  // show left it: `readShows` reads no others as passes), after the show
  // they were written for.
  // (The shows drawn twice, in the order the stream draws them: a show cut in two for its state is none of them.)
  if (passEdits.length) {
    const twice = paragraph.allShows.filter((show) => passAfter.has(show.end) ? passAfter.get(show.end) !== null : Boolean(show.overstruck));
    const read = after.filter((show) => show.overstruck);
    if (read.length !== twice.length) cannot();
    for (const [order, show] of twice.entries()) {
      const offset = passAfter.get(show.end)?.offset ?? show.overstruck!.offset;
      const now = read[order];
      if (now.fontName !== show.fontName || now.codes.length !== show.codes.length ||
        Math.abs(now.overstruck!.offset - offset) > 1e-6 * Math.max(1, Math.abs(show.fontSize))) cannot();
    }
  }
  // Every show under the matrix it had, but the box's under theirs leaned as
  // asked: its up axis moved along its own line by the shear, nothing else of
  // it. (Where each glyph stands was read above; a lean moves none.)
  // (Glyph by glyph in the order the stream draws them, as above: a show
  // that was cut in two for a lean and is one show again draws the same
  // glyphs under the same matrix.)
  if (leans.size) {
    const framesOf = (shows: readonly MraiNativeParagraphRewriteShow[], shearOf: ReadonlyMap<number, number> | null) => shows.flatMap((show) => {
      const unit = paragraph.fonts.get(show.fontName)?.codeUnitBytes ?? 0;
      const count = unit ? Math.floor(show.codes.length / unit) : 1;
      const shear = shearOf?.get(show.end) ?? 0;
      const [a, b, c, d] = show.textMatrix;
      const frame = Object.freeze({ linear: [a, b, c + shear * a, d + shear * b] as const, graphics: show.graphicsMatrix });
      return Array.from({ length: count }, () => frame);
    });
    const [expected, drawn] = [framesOf(leanedShows, leanOf), framesOf(after, null)];
    if (expected.length !== drawn.length) cannot();
    for (const [glyph, was] of expected.entries()) {
      const now = drawn[glyph];
      if (now === was) continue;
      const scale = Math.max(1e-9, Math.hypot(was.linear[0], was.linear[1]), Math.hypot(was.linear[2], was.linear[3]));
      if (was.linear.some((value, at) => Math.abs(value - now.linear[at]) > 2e-6 * scale) ||
        was.graphics.some((value, at) => Math.abs(value - now.graphics[at]) > 1e-9 * Math.max(1, Math.abs(value)))) cannot();
    }
  }
  const moved = (offset: number, inclusive: boolean) => offset + edits.reduce((sum, edit) =>
    edit.end < offset || (edit.end === offset && (inclusive || edit.start < edit.end)) ? sum + edit.text.length - (edit.end - edit.start) : sum, 0);
  for (const { group, style, on } of ruledAfter) {
    const edges = { start: moved(group.object.start, true), end: moved(group.object.end, false) };
    if ((standingRuleOf(candidateBytes, group, style, edges, asked.color) !== null) !== on) cannot();
  }
  
  // describe one, with no block, no pen and no source bytes.
  const sha256 = await sha256PdfBytes(streamBytes.slice(spanStart, spanEnd));
  const codes = paragraph.shows.flatMap((show) => [...show.show.codes]);
  const { layout, firstPlan } = paragraph;
  const advance = layout.lineExtents.reduce((sum, extent) => sum + (extent.right - extent.left), 0);
  const rewrite: PdfNativeParagraphRewriteEvidence = Object.freeze({
    schemaVersion: MRAI_NATIVE_PARAGRAPH_REWRITE_SCHEMA_VERSION,
    target: Object.freeze({ objectNumber: firstPlan.request.identity.objectNumber, generation: firstPlan.request.identity.generation }),
    items: Object.freeze(paragraph.plans.map((plan) => plan?.evidence ?? null)),
    shows: Object.freeze(paragraph.shows.map((show) => Object.freeze({ start: show.start, end: show.end }))),
    pens: Object.freeze([]),
    originalBytes: Object.freeze([]),
    pins: Object.freeze(paragraph.pins.map((pin) => Object.freeze({ objectNumber: pin.objectNumber, generation: pin.generation, role: pin.role }))),
    block: null,
    inPlace: true as const,
  });
  const evidence: PdfNativeListTextFidelityEvidence = Object.freeze({
    ...firstPlan.evidence,
    schemaVersion: PARAGRAPH_TEXT_SCHEMA_VERSION,
    layoutKind: "paragraph-line" as const,
    selection: paragraph.selection,
    wrapping: "fixed-frame-multiline" as const,
    linePlan: null,
    listAnchor: null,
    originalText: text,
    originalDisplayText: text,
    replacementText: text,
    replacementDisplayText: text,
    originalCodeBytes: Object.freeze(codes),
    replacementCodeBytes: Object.freeze(codes),
    textX: layout.firstLeft,
    baseline: layout.firstBaseline,
    lineSpacing: layout.lineStep,
    originalAdvance: advance,
    replacementAdvance: advance,
    textOperatorSpan: Object.freeze({ start: spanStart, end: spanEnd, sha256 }),
    paragraphRewrite: rewrite,
  });
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: firstPlan.request.identity,
    expectedObjectSha256: firstPlan.request.expectedObjectSha256,
    expectedStreamSha256: firstPlan.request.expectedStreamSha256,
    operatorSpan: Object.freeze({ start: spanStart, end: spanEnd, expectedSha256: sha256 }),
    replacement,
    replacementPieces: pieces,
    untouchedObjects: paragraph.pins,
    ...(firstPlan.request.pageContentConsolidation ? { pageContentConsolidation: firstPlan.request.pageContentConsolidation } : {}),
  });
  return Object.freeze({ request, evidence, properties: null });
}

/**
 * A format asked of the WHOLE of a box of right-to-left (or mixed) text, its
 * text unchanged: written where the text stands (readParagraphWhereItStands,
 * setParagraphWhereItStands), every show of the box drawn again by its own
 * operands at its own place in the look asked. Every character must be asked
 * the same change, and only one that moves no glyph: a part of the box, a
 * face, a size, a script, a spacing, an alignment keep the refusal of text
 * the reader reorders (which character of the reader's text a glyph is, this
 * tier does not know). What is written is read again glyph by glyph
 * (refuseMovedGlyphs).
 */
async function planFormatWhereReorderedTextStands(
  kernel: MraiNativeParagraphRewriteKernel,
  selection: PdfNativeListTextSelectionEvidence,
  resolveFormats: MraiNativeParagraphRewriteFormatsResolver,
): Promise<PdfNativeListTextReplacementPlan> {
  const paragraph = await readParagraphWhereItStands(kernel, selection);
  // (No text is set again: a glyph the reader reads as a blank is drawn again
  // with the show that draws it. One the box does not own, on a line of it,
  // would stay as it was beside text in a new look.)
  refuseGlyphsReadAsBlanksBesideLines(paragraph);
  const text = paragraph.selection.item.str;
  const formats = await resolveFormats({ request: paragraph.firstPlan.request, evidence: paragraph.firstPlan.evidence,
    pageContent: paragraph.pageContent, text });
  const asked = formats.characters[0] ?? null;
  const key = JSON.stringify(asked);
  if (formats.characters.length !== characters(text).length || !asked ||
    formats.characters.some((format) => JSON.stringify(format ?? null) !== key) ||
    Object.entries(asked).some(([field, value]) => value !== undefined && !STANDING_FORMAT_FIELDS.has(field)) ||
    formats.alignment !== undefined || formats.leftIndent || formats.lineStep !== undefined ||
    formats.firstLineIndent !== undefined) {
    refuse(MRAI_NATIVE_REORDERED_TEXT_MESSAGE);
  }
  // Asked what every show of it is drawn as already: nothing to write (the
  // block would draw the same page again), and said so.
  // A bold that is its faces' only: its words set again in the family's
  
  // is asked first (planMraiNativeParagraphRewrite); refused as the face's
  
  const inkedShows = paragraph.shows.filter((show) => show.inked);
  if (asked.syntheticBold === false && !inkedShows.some((show) => show.stroked) && inkedShows.some(standsInBoldFace)) {
    const set = await planBoldFaceSetRegularWhereWordsStand(kernel, selection, resolveFormats);
    if (set) return set;
  }
  refuseBoldTakenOffBoldFace(asked.syntheticBold, inkedShows);
  const sameColor = (left: unknown, right: unknown) => JSON.stringify(left) === JSON.stringify(right);
  if (!paragraph.shows.some((show) => show.inked &&
    ((asked.color !== undefined && !sameColor(asked.color, show.color)) ||
      strokeAskedChanges(asked.syntheticBold, show) ||
      (asked.slant !== undefined && asked.slant !== show.slant) ||
      // (A glyph of no width takes no rule: readParagraphWhereItStands.)
      (show.inked.end > show.inked.start && asked.underline !== undefined && asked.underline !== (show.underline === true)) ||
      (show.inked.end > show.inked.start && asked.strikethrough !== undefined &&
        asked.strikethrough !== (show.strikethrough === true))))) {
    refuse(MRAI_NATIVE_STANDING_FORMAT_UNCHANGED_REFUSAL);
  }
  const set = setParagraphWhereItStands(kernel, paragraph, [text], formats);
  if (!set) return refuse(MRAI_NATIVE_REORDERED_TEXT_MESSAGE);
  const { request, evidence } = await paragraphEvidence(paragraph, text, set);
  refuseMovedGlyphs(kernel, paragraph, request.replacement);
  return Object.freeze({ request, evidence, properties: null });
}

/** One glyph of a bold show as the regular face sets it again. */
type RegularFaceGlyph = Readonly<{ fontName: string; font: MraiNativeParagraphRewriteFont; code: number;
  /** How far its origin moves along its line on the page, and that as a vector of the page. */
  moved: number; move: readonly [number, number];
  /** Its advance in the regular face, in the show's text space and in points along the line. */
  advance: number; advancePoints: number }>;

/** A box's bold shows as the regular face sets them again (regularFaceWordsOf):
 * by the show (its end in the stream) and the glyph's place in it. */
type RegularFaceWords = Readonly<{ glyphs: ReadonlyMap<number, readonly RegularFaceGlyph[]>;
  /** Shows that are not the box's and draw on from a show of it (the pen not placed afresh between them): blanks
   * all, which stand that much back once the box's shows are emptied -- by the end of each one's span, as a vector
   * of the page (refuseMovedGlyphs' `carried`). */
  carried: ReadonlyMap<number, readonly [number, number]>;
  /** Where a glyph is set in a regular face the page does not name, which
   * another page of the file does: the plan that names those fonts in the
   * page's resources, each font by the name the plan gives it, and the
   * objects their reading pinned. Null where the page's own faces set all. */
  reference: Readonly<{ plan: PdfNativePageFontReferencePlan;
    fonts: ReadonlyMap<string, MraiNativeParagraphRewriteFont>;
    pins: readonly PdfNativeObjectPrecondition[] }> | null }>;

/**
 * Every glyph of a box's bold shows in the regular face of its family, each
 * word where it stood (regular-face-words.ts), or null where the box is not
 * one whose words can all be set so:
 *  - text of the page's own content, upright, with no rule under or through
 *    it, no fake-bold pass, no show it holds only part of, none stroked;
 *  - every inked show of it in a bold face that is a simple TrueType font
 *    with its program in the file, the others (a Latin word in another
 *    family) left as they stand;
 *  - for each glyph, one code of a regular face of the same family (a
 *    simple TrueType font with its program, named regular, upright as the
 *    bold face is or italic as it is) reading as the same character -- a
 *    presentation form, which is the glyph -- where the bold face names
 *    that character by one code alone and the regular face too: a face the
 *    PAGE names, else one another page of the FILE names, which the write
 *    then names in the page's resources (PdfNativePageFontReferencePlan);
 *    the fewest such faces that hold every glyph the page's lack, and no
 *    more than one write names;
 *  - the ink of both glyphs read from their programs, and the words set as
 *    planPdfRegularFaceWords holds them: joins kept, nothing run together,
 *    no edge moved more than two points, no gap closed to under half.
 * (pdf-walk-1000 0979 p.1: the title's «بعض» needs one glyph of a regular
 * subset only later pages name.) No font is written or changed: a family
 * whose regular face no page of the file holds the glyph in stays bold.
 */
/** Why the box last asked was not set again in its regular face
 * (regularFaceWordsOf), or null where it was: for the tests and the probes
 * that tell one refused box from another. Never read by a writer. */
let regularFaceWordsRefusal: string | null = null;
export function lastMraiNativeRegularFaceWordsRefusalForTests(): string | null {
  return regularFaceWordsRefusal;
}

async function regularFaceWordsOf(
  kernel: MraiNativeParagraphRewriteKernel,
  paragraph: ReadParagraph,
): Promise<RegularFaceWords | null> {
  const not = (why: string): null => { regularFaceWordsRefusal = why; return null; };
  if (!kernel.trueTypeFaces || !paragraph.pageContent || paragraph.twins.length || paragraph.rules.length) return not("not-plain-page-text");
  const { shows } = paragraph;
  const inked = shows.filter((show) => show.inked);
  if (!inked.length || !inked.some(standsInBoldFace)) return not("no-bold-face");
  if (shows.some((show) => show.partial || show.stroked || show.slant !== 0 || show.underline || show.strikethrough ||
    show.linkRule || show.twin || show.readsBlank || show.show.vertical || show.font.codeUnitBytes !== 1)) return not("show-not-plain");
  const bold = inked.filter(standsInBoldFace);
  const faces = await kernel.trueTypeFaces(paragraph.selection, [...paragraph.fonts.keys()]);
  const sansTag = (name: string) => name.replace(/^[A-Z]{6}\+/u, "");
  const extentsOf = new Map<string, ReturnType<typeof readPdfSimpleTrueTypeGlyphExtents>>();
  const extents = (name: string) => {
    if (!extentsOf.has(name)) {
      const face = faces.get(name);
      extentsOf.set(name, face ? readPdfSimpleTrueTypeGlyphExtents(face.program) : null);
    }
    return extentsOf.get(name) ?? null;
  };
  // The regular faces of a bold face's family that the page holds, in the
  // order of their names (one answer whatever order the page lists them in).
  const regularOf = new Map<string, readonly string[]>();
  const regularFaces = (boldName: string): readonly string[] => {
    const known = regularOf.get(boldName);
    if (known) return known;
    const own = faces.get(boldName);
    const family = own ? pdfNativeFontFamilyKey(sansTag(own.baseFont)) : "";
    const italic = own ? pdfNativeFaceNameStyle(sansTag(own.baseFont)).italic : false;
    const found = !own || !family ? [] : [...faces.entries()].filter(([name, face]) => {
      const style = pdfNativeFaceNameStyle(sansTag(face.baseFont));
      const font = paragraph.fonts.get(name);
      return name !== boldName && !style.bold && style.italic === italic && font?.codeUnitBytes === 1 &&
        font.simpleSubtype !== "Type3" && pdfNativeFontFamilyKey(sansTag(face.baseFont)) === family;
    }).map(([name]) => name).sort();
    regularOf.set(boldName, found);
    return found;
  };
  /** The one code of a font that reads as `text`, or null (none, or several). */
  const codeOf = (font: MraiNativeParagraphRewriteFont, text: string): number | null => {
    let found: number | null = null;
    for (const [code, read] of font.decode) {
      if (read !== text) continue;
      if (found !== null) return null;
      found = code;
    }
    return found;
  };
  const thousandths = (font: MraiNativeParagraphRewriteFont, code: number) => kernel.glyphAdvance(Uint8Array.of(code), font, 1000);
  // Along the line on the page: its direction, read off the first bold show.
  const placedOf = (show: MraiNativeParagraphRewriteShow) => composeMatrix(show.textMatrix, show.graphicsMatrix);
  const first = placedOf(bold[0].show);
  const length = Math.hypot(first[0], first[1]);
  if (!(length > 0)) return not("no-direction");
  const along = [first[0] / length, first[1] / length] as const;
  const dot = (point: readonly [number, number]) => point[0] * along[0] + point[1] * along[1];
  const across = (point: readonly [number, number]) => point[1] * along[0] - point[0] * along[1];
  const ownedAt = new Map(shows.map((show) => [show.end, show] as const));
  type RegularGlyph = Readonly<{ fontName: string; font: MraiNativeParagraphRewriteFont; code: number;
    ink: Readonly<{ left: number; right: number }> }>;
  /** The glyph of a regular face the page names that reads as `text`. */
  const regularOnPage = (boldName: string, text: string): RegularGlyph | null => {
    for (const name of regularFaces(boldName)) {
      const candidate = paragraph.fonts.get(name)!;
      const code = codeOf(candidate, text);
      const ink = code === null ? null : extents(name)?.(code) ?? null;
      if (code !== null && ink) return { fontName: name, font: candidate, code, ink };
    }
    return null;
  };
  // The glyphs no regular face of the page holds, by the bold face that
  // draws them: looked for in the regular faces of that family the other
  // pages of the file name. The fewest faces that hold them all are taken
  // (the one that holds most of what is still missing first; of equals, one
  // already taken for another face of the box, then the lower object), and
  // the page's resources name each for this write.
  const missing = new Map<string, Set<string>>();
  const referenced = new Map<string, Map<string, RegularGlyph>>();
  let reference: RegularFaceWords["reference"] = null;
  const planned: PdfRegularFaceGlyph[] = [];
  const source: { owned: OwnedShow; glyph: number; fontName: string; font: MraiNativeParagraphRewriteFont; code: number;
    advance: number; pointsPerUnit: number }[] = [];
  const lineAcross = new Map<number, { across: number; em: number }>();
  /** Every glyph of the bold shows in its regular face, or why not; a glyph
   * no face read so far holds is gathered in `missing` and passed over. */
  const setInRegularFaces = (): string | null => {
    planned.length = 0;
    source.length = 0;
    lineAcross.clear();
    missing.clear();
    for (const owned of bold) {
      const { show, font } = owned;
      const placed = placedOf(show);
      const stretch = Math.hypot(placed[0], placed[1]);
      // (One direction for the whole box, the text upright in it.)
      if (!(stretch > 0) || !(show.fontSize > 0) || !(show.horizontalScale > 0) ||
        Math.abs(placed[0] / stretch - along[0]) > 1e-6 || Math.abs(placed[1] / stretch - along[1]) > 1e-6) return "not-one-direction";
      const scale = show.horizontalScale / 100;
      const ownExtents = extents(show.fontName);
      if (!ownExtents) return "bold-program-unread";
      const em = show.fontSize * scale * stretch;
      for (const [index, glyph] of showGlyphs(kernel, show, font).entries()) {
        const text = glyph.text;
        if (text === undefined || /^\s*$/u.test(text) || codeOf(font, text) !== glyph.code) return "bold-glyph-unnamed";
        const boldInk = ownExtents(glyph.code);
        if (!boldInk) return "bold-glyph-no-ink";
        const regular = regularOnPage(show.fontName, text) ?? referenced.get(show.fontName)?.get(text) ?? null;
        if (!regular) {
          const texts = missing.get(show.fontName) ?? new Set<string>();
          texts.add(text);
          missing.set(show.fontName, texts);
          continue;
        }
        const origin = applyMatrix(placed, glyph.start * scale, show.textRise);
        const advance = thousandths(regular.font, regular.code);
        if (!(advance > 0)) return "regular-glyph-no-advance";
        planned.push({ line: owned.line, at: dot(origin), em,
          bold: { advance: thousandths(font, glyph.code), left: boldInk.left, right: boldInk.right },
          regular: { advance, left: regular.ink.left, right: regular.ink.right } });
        source.push({ owned, glyph: index, fontName: regular.fontName, font: regular.font, code: regular.code,
          advance: (advance / 1000) * show.fontSize, pointsPerUnit: scale * stretch });
        if (!lineAcross.has(owned.line)) lineAcross.set(owned.line, { across: across(origin), em });
      }
    }
    return null;
  };
  const refused = setInRegularFaces();
  if (refused) return not(refused);
  if (missing.size) {
    const onPage = (boldName: string) => (regularFaces(boldName).length ? "no-regular-glyph-on-the-page" : "no-regular-face-on-the-page");
    if (!kernel.fileTrueTypeFaces || !kernel.referencePageFonts) return not(onPage([...missing.keys()][0]));
    // (A page read as its streams joined into one is written as a new stream the page then names: a font
    // reference rides one plain write of a stream the page already draws, and is not asked there.)
    if (paragraph.firstPlan.request.pageContentConsolidation) return not("page-drawn-from-joined-streams");
    const fileFaces = await kernel.fileTrueTypeFaces(paragraph.selection);
    type Taken = { font: Readonly<{ objectNumber: number; generation: number }>; metrics: MraiNativeParagraphRewriteFont;
      pins: readonly PdfNativeObjectPrecondition[]; glyph(text: string): Omit<RegularGlyph, "fontName"> | null };
    const read = new Map<string, Promise<Taken | null>>();
    const readFace = (face: (typeof fileFaces)[number]) => {
      const key = `${face.font.objectNumber}:${face.font.generation}`;
      if (!read.has(key)) read.set(key, (async () => {
        const loaded = await face.read();
        if (!loaded || loaded.font.codeUnitBytes !== 1 || loaded.font.simpleSubtype === "Type3") return null;
        const ink = readPdfSimpleTrueTypeGlyphExtents(loaded.program);
        if (!ink) return null;
        return { font: face.font, metrics: loaded.font, pins: loaded.pins, glyph: (text: string) => {
          const code = codeOf(loaded.font, text);
          const own = code === null ? null : ink(code);
          return code !== null && own ? { font: loaded.font, code, ink: own } : null;
        } };
      })());
      return read.get(key)!;
    };
    const taken: Taken[] = [];
    const served = new Map<string, Map<string, Taken>>();
    for (const [boldName, texts] of [...missing].map(([name, texts]) => [name, new Set(texts)] as const)
      .sort(([left], [right]) => (left < right ? -1 : left > right ? 1 : 0))) {
      const own = faces.get(boldName);
      const family = own ? pdfNativeFontFamilyKey(sansTag(own.baseFont)) : "";
      const italic = own ? pdfNativeFaceNameStyle(sansTag(own.baseFont)).italic : false;
      const candidates: Taken[] = [];
      for (const face of fileFaces) {
        const style = pdfNativeFaceNameStyle(sansTag(face.baseFont));
        if (!family || style.bold || style.italic !== italic || pdfNativeFontFamilyKey(sansTag(face.baseFont)) !== family) continue;
        const loaded = await readFace(face);
        if (loaded) candidates.push(loaded);
      }
      if (!candidates.length) return not(regularFaces(boldName).length ? "no-regular-glyph-in-the-file" : "no-regular-face-in-the-file");
      const left = new Set(texts);
      const byText = new Map<string, Taken>();
      while (left.size) {
        let best: Taken | null = null, bestHolds = 0;
        for (const candidate of [...candidates].sort((a, b) =>
          Number(taken.includes(b)) - Number(taken.includes(a)) ||
          a.font.objectNumber - b.font.objectNumber || a.font.generation - b.font.generation)) {
          const holds = [...left].filter((text) => candidate.glyph(text)).length;
          if (holds > bestHolds) { best = candidate; bestHolds = holds; }
        }
        if (!best) return not("no-regular-glyph-in-the-file");
        if (!taken.includes(best)) taken.push(best);
        for (const text of [...left]) if (best.glyph(text)) { byText.set(text, best); left.delete(text); }
      }
      served.set(boldName, byText);
    }
    if (taken.length > MAX_MRAI_NATIVE_PAGE_FONT_REFERENCES) return not("too-many-faces-of-the-file");
    taken.sort((a, b) => a.font.objectNumber - b.font.objectNumber || a.font.generation - b.font.generation);
    let plan: PdfNativePageFontReferencePlan;
    try {
      plan = await kernel.referencePageFonts(paragraph.selection, taken.map((item) => item.font),
        paragraph.firstPlan.request.identity);
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
      return not(`page-resources-take-no-font: ${error.message}`);
    }
    const nameOf = new Map(taken.map((item, index) => [item, plan.fonts[index].resourceName] as const));
    for (const [boldName, byText] of served) {
      referenced.set(boldName, new Map([...byText].map(([text, item]) =>
        [text, { fontName: nameOf.get(item)!, ...item.glyph(text)! }] as const)));
    }
    const pins = new Map<string, PdfNativeObjectPrecondition>();
    for (const pin of taken.flatMap((item) => item.pins)) pins.set(`${pin.objectNumber}:${pin.generation}`, pin);
    reference = Object.freeze({ plan, fonts: new Map(taken.map((item) => [nameOf.get(item)!, item.metrics] as const)),
      pins: Object.freeze([...pins.values()]) });
    // Set again, the faces taken beside the page's own.
    const again = setInRegularFaces();
    if (again) return not(again);
    if (missing.size) return not("no-regular-glyph-in-the-file");
  }
  // What stays where it is on the box's lines: every other show of the
  // stream that draws ink there (the box's own in other faces among them),
  // and the blanks the box draws between its words.
  const blanks: { line: number; at: number }[] = [];
  const fixed: { line: number; left: number; right: number; joinsBefore: boolean; joinsAfter: boolean }[] = [];
  const setAgain = new Set(bold.map((show) => show.end));
  for (const show of paragraph.allShows) {
    if (setAgain.has(show.end) || show.codes.length === 0) continue;
    const font = paragraph.fonts.get(show.fontName);
    if (!font || show.vertical) continue;
    const placed = placedOf(show);
    const scale = show.horizontalScale / 100;
    let glyphs: ReturnType<typeof showGlyphs>;
    try { glyphs = showGlyphs(kernel, show, font); } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
      continue;
    }
    if (!glyphs.length) continue;
    const start = applyMatrix(placed, glyphs[0].start * scale, show.textRise);
    const end = applyMatrix(placed, glyphs.at(-1)!.end * scale, show.textRise);
    for (const [line, own] of lineAcross) {
      if (Math.abs(across(start) - own.across) > 0.5 * own.em) continue;
      // (A space of a font the file holds no program of, which names no character for it -- a standard
      // encoding's code 32 -- draws nothing; in a subset with a program, code 32 is a letter.)
      const draws = glyphs.some((glyph) => glyph.text === undefined
        ? !(glyph.code === 0x20 && !faces.has(show.fontName)) : !/^\s*$/u.test(glyph.text));
      if (!draws) { if (ownedAt.has(show.end)) blanks.push({ line, at: dot(start) }); continue; }
      // Whether its glyph at either end is a joined form: by its ink where its program is read; else by
      // what it reads as (a letter of a script that joins may be one; a full stop, a digit, a Latin letter not).
      const ink = extents(show.fontName);
      const joined = (glyph: (typeof glyphs)[number], side: "before" | "after") => {
        const own = ink?.(glyph.code);
        if (own) return side === "before" ? own.left <= -5 : own.right >= thousandths(font, glyph.code) - 2;
        return glyph.text === undefined || pdfTextHasRightToLeftRun(glyph.text);
      };
      const [low, high] = dot(start) <= dot(end) ? [glyphs[0], glyphs.at(-1)!] : [glyphs.at(-1)!, glyphs[0]];
      fixed.push({ line, left: Math.min(dot(start), dot(end)), right: Math.max(dot(start), dot(end)),
        joinsBefore: joined(low, "before"), joinsAfter: joined(high, "after") });
    }
  }
  const words = planPdfRegularFaceWords({ glyphs: planned, blanks, fixed });
  if (!words.ok) return not(words.detail ? `${words.reason} (${words.detail})` : words.reason);
  // What draws on from a show of the box. Every show of it is emptied, so a
  // show after one in its text object, with the pen not placed afresh between
  // them, stands back by what the emptied shows advanced. A producer pads a
  // column with spaces drawn on from one another (0979: runs of them after a
  // word): blanks may stand back, and are read again there; ink may not.
  const carried = new Map<number, readonly [number, number]>();
  const allAt = new Map(paragraph.allShows.map((show) => [show.start, show] as const));
  const textObjects = new Map<number, Readonly<{ start: number; end: number }>>();
  for (const owned of shows) {
    const object = enclosingTextObject(paragraph.source, owned.start);
    textObjects.set(object.start, object);
  }
  for (const object of textObjects.values()) {
    const operators = mraiNativeShowOperatorsIn(paragraph.source, object.start + 2, object.end - 2);
    if (!operators) return not("text-object-unread");
    let back: readonly [number, number] = [0, 0];
    for (const [index, operator] of operators.entries()) {
      if (index > 0 && mraiNativeTextAfterShowIsPlacedAfresh(paragraph.source, operators[index - 1].end, operators[index - 1].end, object.end)) {
        back = [0, 0];
      }
      const show = allAt.get(operator.start);
      const font = show ? paragraph.fonts.get(show.fontName) : undefined;
      if (ownedAt.has(operator.end)) {
        if (!show || !font || show.vertical) return not("text-object-unread");
        const drawn = showGlyphs(kernel, show, font);
        if (!drawn.length) continue;
        const placed = placedOf(show), scale = show.horizontalScale / 100;
        const from = applyMatrix(placed, drawn[0].start * scale, show.textRise), to = applyMatrix(placed, drawn.at(-1)!.end * scale, show.textRise);
        back = [back[0] - (to[0] - from[0]), back[1] - (to[1] - from[1])];
        continue;
      }
      if ((back[0] === 0 && back[1] === 0) || !show || show.codes.length === 0) continue;
      if (!font || show.vertical) return not("text-drawn-on-from-its-shows");
      if (showGlyphs(kernel, show, font).some((glyph) => glyph.text === undefined
        ? !(glyph.code === 0x20 && !faces.has(show.fontName)) : !/^\s*$/u.test(glyph.text))) return not("text-drawn-on-from-its-shows");
      carried.set(show.end, back);
    }
  }
  const glyphs = new Map<number, RegularFaceGlyph[]>();
  for (const [index, entry] of source.entries()) {
    const moved = words.at[index] - planned[index].at;
    const row = glyphs.get(entry.owned.end) ?? [];
    row[entry.glyph] = Object.freeze({ fontName: entry.fontName, font: entry.font, code: entry.code, moved,
      move: [moved * along[0], moved * along[1]] as const, advance: entry.advance, advancePoints: entry.advance * entry.pointsPerUnit });
    glyphs.set(entry.owned.end, row);
  }
  return Object.freeze({ glyphs, carried, reference });
}

/**
 * The write that sets a read box's bold shows again in the regular face of
 * their family where its words stand (planBoldFaceSetRegularWhereWordsStand),
 * its shows drawn in `color` where one is asked; null where the box is not
 * one whose every word can be set so. The whole of the write but the format
 * asked: the box's proof plans it too and drops it (regularFaceSetsBox), so
 * what the proof says of a box is what the press on it is answered.
 */
async function regularFaceWriteOf(
  kernel: MraiNativeParagraphRewriteKernel,
  paragraph: ReadParagraph,
  color: PdfNativeDeviceColor | undefined,
): Promise<Readonly<{ request: PdfNativeContentStreamTransactionRequest; evidence: PdfNativeListTextFidelityEvidence }> | null> {
  const text = paragraph.selection.item.str;
  const words = await regularFaceWordsOf(kernel, paragraph);
  if (!words) return null;
  const [b0, b1] = paragraph.basis;
  const block = paragraph.blockMatrix;
  // (A length along the lines of the block, on the page.)
  const pagePerBlock = Math.hypot(b0 * block[0] + b1 * block[2], b0 * block[1] + b1 * block[3]);
  if (!(pagePerBlock > 0)) return null;
  // (Its shows in other faces are drawn again as they stand, in the colour asked.)
  const set = setParagraphWhereItStands(kernel, paragraph, [text],
    Object.freeze({ characters: Object.freeze(characters(text).map(() => (color ? Object.freeze({ color }) : null))) }),
    (_owner, owned) => {
      const row = words.glyphs.get(owned.end);
      if (!row) return null;
      const drawn = showGlyphs(kernel, owned.show, owned.font);
      if (row.length !== drawn.length) return refuse(MRAI_NATIVE_PARAGRAPH_FORMAT_IN_PLACE_REFUSAL);
      const runs: { place: string; style: CharacterStyle; codes: number[]; fontName: string;
        /** How far the glyph after its last stands from its own origin if it is drawn on from it. */ next: number }[] = [];
      for (const [index, glyph] of row.entries()) {
        if (!glyph) return refuse(MRAI_NATIVE_PARAGRAPH_FORMAT_IN_PLACE_REFUSAL);
        const last = runs.at(-1);
        let run = last && last.fontName === glyph.fontName && Math.abs(last.next - glyph.moved) <= 1e-6 ? last : null;
        if (run) run.codes.push(glyph.code);
        else {
          const piece = sourceShowPiece(kernel, paragraph, owned, index, index + 1);
          if (!piece) return refuse(MRAI_NATIVE_PARAGRAPH_FORMAT_IN_PLACE_REFUSAL);
          run = { place: piece.place(owned.slant, glyph.moved / pagePerBlock, 0), codes: [glyph.code], fontName: glyph.fontName, next: 0,
            style: Object.freeze({ ...piece.style, fontName: glyph.fontName, font: glyph.font, characterSpacing: 0, syntheticBold: false,
              ...(color ? { color } : {}) }) };
          runs.push(run);
        }
        // Drawn on from this glyph, the next of the show stands the regular advance after it: that far
        // from its own origin, which the page set a bold advance (and its spacing) after this one's.
        const after = drawn[index + 1];
        run.next = after
          ? glyph.moved + glyph.advancePoints - (after.start - drawn[index].start) * (glyph.advancePoints / glyph.advance)
          : Number.NaN;
      }
      // (No word spacing: a code 32 of a subset is a letter.)
      return runs.map((run) => Object.freeze({ place: run.place, style: run.style, word: 0,
        array: hexCodes(Uint8Array.from(run.codes)), codes: Object.freeze(run.codes) }));
    });
  regularFaceWordsRefusal = "not-drawn-again-where-it-stands";
  if (!set) return null;
  const written = await paragraphEvidence(paragraph, text, set);
  const { evidence } = written;
  // Where a glyph is set in a face the page did not name: the write names
  // it in the page's resources, and pins the objects the face was read from.
  const reference = words.reference;
  let request = written.request;
  if (reference) {
    const pins = new Map<string, PdfNativeObjectPrecondition>((request.untouchedObjects ?? []).map((pin) => [`${pin.objectNumber}:${pin.generation}`, pin] as const));
    for (const pin of reference.pins) {
      const key = `${pin.objectNumber}:${pin.generation}`;
      if (pin.objectNumber === request.identity.objectNumber && pin.generation === request.identity.generation) {
        return refuse(MRAI_NATIVE_PARAGRAPH_FORMAT_IN_PLACE_REFUSAL);
      }
      if (!pins.has(key)) pins.set(key, pin);
    }
    if (pins.size > 256) return refuse("The page binds more objects than one edit can carry.");
    request = Object.freeze({ ...request, untouchedObjects: Object.freeze([...pins.values()]),
      pageFontReference: reference.plan });
  }
  regularFaceWordsRefusal = "written-glyphs-not-where-planned";
  // Read again: each glyph set again where it was planned, in the face and code planned; every other
  // glyph of the stream where it stood.
  refuseMovedGlyphs(kernel, reference ? Object.freeze({ ...paragraph, fonts: new Map([...paragraph.fonts, ...reference.fonts]) }) : paragraph,
    request.replacement, words.carried, (show, index) => {
      const glyph = words.glyphs.get(show.end)?.[index];
      return glyph ? { fontName: glyph.fontName, code: glyph.code, move: glyph.move, advance: glyph.advance } : null;
    });
  regularFaceWordsRefusal = null;
  return Object.freeze({ request, evidence });
}

/** Whether a bold taken off the box is written by setting its words again
 * in the regular face: the write itself, planned as the window's press is
 * (on the box read as that press reads it) and dropped. */
async function regularFaceSetsBox(
  kernel: MraiNativeParagraphRewriteKernel,
  selection: PdfNativeListTextSelectionEvidence,
): Promise<boolean> {
  regularFaceWordsRefusal = "box-not-read";
  try {
    const paragraph = await readParagraphWhereItStands(kernel, selection, false, false, false, true);
    refuseGlyphsReadAsBlanksBesideLines(paragraph);
    return (await regularFaceWriteOf(kernel, paragraph, undefined)) !== null;
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError)) throw error;
    if (regularFaceWordsRefusal !== null && !regularFaceWordsRefusal.includes(":")) regularFaceWordsRefusal += `: ${error.message}`;
    return false;
  }
}





























async function planBoldFaceSetRegularWhereWordsStand(
  kernel: MraiNativeParagraphRewriteKernel,
  selection: PdfNativeListTextSelectionEvidence,
  resolveFormats: MraiNativeParagraphRewriteFormatsResolver,
): Promise<PdfNativeListTextReplacementPlan | null> {
  regularFaceWordsRefusal = "not-read-back-by-its-caller";
  if (resolveFormats.wordsReadByReader !== true) return null;
  regularFaceWordsRefusal = "box-not-read";
  try {
    const paragraph = await readParagraphWhereItStands(kernel, selection, false, false, false, true);
    refuseGlyphsReadAsBlanksBesideLines(paragraph);
    const text = paragraph.selection.item.str;
    const formats = await resolveFormats({ request: paragraph.firstPlan.request, evidence: paragraph.firstPlan.evidence,
      pageContent: paragraph.pageContent, text });
    regularFaceWordsRefusal = "not-asked-of-the-whole-box-alone";
    // Asked of all of it, and nothing with it but a colour (which the block
    // draws its shows in; a rule or a lean pressed with it is asked again
    // once the box is regular).
    const askedOfAll = formats.characters[0] ?? null;
    const key = JSON.stringify(askedOfAll);
    if (formats.characters.length !== characters(text).length || formats.alignment !== undefined || formats.leftIndent ||
      formats.lineStep !== undefined || formats.firstLineIndent !== undefined || !askedOfAll || askedOfAll.syntheticBold !== false ||
      Object.entries(askedOfAll).some(([field, value]) => value !== undefined && field !== "syntheticBold" && field !== "color") ||
      formats.characters.some((format) => JSON.stringify(format ?? null) !== key)) return null;
    const write = await regularFaceWriteOf(kernel, paragraph, askedOfAll.color);
    return write && Object.freeze({ ...write, properties: null });
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError)) throw error;
    if (regularFaceWordsRefusal !== null && !regularFaceWordsRefusal.includes(":")) regularFaceWordsRefusal += `: ${error.message}`;
    return null;
  }
}











export async function readMraiNativeStandingTextLook(
  kernel: MraiNativeParagraphRewriteKernel,
  selection: PdfNativeListTextSelectionEvidence,
): Promise<MraiNativeStandingTextLook | null> {
  if (!kernel.standingPageStreams || !pdfTextHasRightToLeftRun(selection.item.str)) return null;
  try {
    const paragraph = await readParagraphWhereItStands(kernel, selection);
    const inked = paragraph.shows.filter((show) => show.inked);
    if (!inked.length) return null;
    const rules = paragraphTextRules(paragraph);
    const boldFaces = boldFaceLookOf(inked);
    // A bold that is its faces' only: whether its words can all be set
    // again in the family's regular face where they stand, as the press
    // that takes it off is then written (planBoldFaceSetRegularWhereWordsStand).
    const regularFace = boldFaces.boldFaces === "only" && await regularFaceSetsBox(kernel, selection);
    return Object.freeze({
      ...(inked.every((show) => show.stroked) ? { syntheticBold: true as const } : {}),
      ...boldFaces,
      ...(regularFace ? { regularFace: true as const } : {}),
      ...(inked.every((show) => show.slant !== 0) ? { slanted: true as const } : {}),
      ...(rules ?? {}),
    });
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError)) throw error;
    
    const look = await readStandingStateLook(kernel, selection);
    // ... and its bold, where that is its faces' only, may still be taken off
    // by its words set again in the regular face: read as the window asks
    // it, which reads the page it wrote back (the box among other text).
    if (look?.boldFaces !== "only") return look;
    return (await regularFaceSetsBox(kernel, selection)) ? Object.freeze({ ...look, regularFace: true as const }) : look;
  }
}

/** The look of a box read where its text stands (readMraiNativeStandingTextLook). */
export type MraiNativeStandingTextLook = Readonly<{
  syntheticBold?: true;
  /** An inked show of it stands in a bold face with no stroke on it: a
   * weight of the face's own, which a bold taken off where the text stands
   * does not take off. "only": no show of the box is stroked, and a bold
   * taken off it is refused (MRAI_NATIVE_BOLD_FACE_KEPT_REFUSAL);
   * "beside-strokes": other shows of it are stroked, and a bold taken off
   * takes those strokes off. */
  boldFaces?: "only" | "beside-strokes";
  /** With `boldFaces`: every other inked show of it is stroked or in a bold
   * face too -- the whole box reads bold, though not every show is stroked. */
  drawnBold?: true;
  /** With `boldFaces` "only": a bold taken off the box is written -- its
   * words set again in the regular face of the family, where they stand
   * (regularFaceWordsOf holds every word of it to its rules). Absent, that
   * press is refused (MRAI_NATIVE_BOLD_FACE_KEPT_REFUSAL). */
  regularFace?: true;
  slanted?: true;
  underline?: true;
  strikethrough?: true;
}>;

/**
 * Whether a box that holds Type3 glyphs holds none but pictures: every show
 * of it drawn in a Type3 face is drawn in a face of pictures (the font's
 * `type3Pictures`, read by the kernel's resolveFont). It is the test the
 * writers of a weight make of such a box (setParagraph's Type3 refusal,
 * planFormatByStateWhereTextStands): a weight asked of it is written on its
 * words, the pictures drawn as they stand. False where a show of it is drawn
 * in a Type3 face of letters -- a weight is refused there -- and where it
 * holds no Type3 glyph at all.
 */
function paragraphHoldsType3PicturesOnly(paragraph: ReadParagraph): boolean {
  const type3 = paragraph.shows.filter((show) => show.font.simpleSubtype === "Type3");
  return type3.length > 0 && type3.every((show) => show.font.type3Pictures === true);
}










function paragraphType3LettersTakeWeight(kernel: MraiNativeParagraphRewriteKernel, paragraph: ReadParagraph): boolean {
  const letters = paragraph.shows.filter((show) => show.font.simpleSubtype === "Type3" && !show.font.type3Pictures);
  return letters.length > 0 && !READ_BY_ITEM_PROOFS.has(paragraph) && letters.every((show) =>
    !show.inked || show.stroked || Boolean(show.show.overstruck) || type3ShowPasses(kernel, paragraph.source, show) !== null);
}



export async function readMraiNativeParagraphType3LettersTakeWeight(
  kernel: MraiNativeParagraphRewriteKernel,
  selection: PdfNativeListTextSelectionEvidence,
): Promise<boolean> {
  const standing = await readParagraphForStateFormat(kernel, selection);
  return standing !== null && paragraphType3LettersTakeWeight(kernel, standing);
}







export async function readMraiNativeParagraphType3PicturesOnly(
  kernel: MraiNativeParagraphRewriteKernel,
  selection: PdfNativeListTextSelectionEvidence,
): Promise<boolean> {
  try {
    return paragraphHoldsType3PicturesOnly(await readParagraph(kernel, selection, true));
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError)) throw error;
    const standing = await readParagraphForStateFormat(kernel, selection);
    return standing !== null && paragraphHoldsType3PicturesOnly(standing);
  }
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
  const textRules = paragraphTextRules(paragraph);
  
  
  // read of the box, by where its shows stand.)
  const type3Letters = paragraph.shows.some((show) => show.font.simpleSubtype === "Type3" && !show.font.type3Pictures);
  const type3Weight = type3Letters && await readMraiNativeParagraphType3LettersTakeWeight(kernel, selection);
  return Object.freeze({
    plan: Object.freeze({
      request,
      evidence,
      properties: null,
      ...(sourceAppearance.length ? { sourceAppearance: Object.freeze(sourceAppearance) } : {}),
      ...(textRules ? { textRules } : {}),
      ...(paragraphHoldsType3PicturesOnly(paragraph) ? { type3Pictures: true as const } : {}),
      ...(type3Weight ? { type3Weight: true as const } : {}),
      ...(type3Letters ? { type3Letters: true as const } : {}),
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
  resized: boolean, rowEdit?: MraiNativeParagraphRowEdit, decided?: { byRows: boolean }): string {
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
  const breakLines: number[] = [];
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
    if (free > (width / count) * piece + em * 0.5) { breaks.push(next); breakLines.push(line); }
  }
  const after = characters(replacementText);
  let head = 0;
  while (head < before.length && head < after.length && before[head] === after[head]) head += 1;
  let tail = 0;
  while (tail < before.length - head && tail < after.length - head &&
    before[before.length - 1 - tail] === after[after.length - 1 - tail]) tail += 1;
  // Words the editor shows at the end of a row, the row under it starting
  // where it did: that row's start is kept as a break too where the words
  // fit on their row -- whatever the break was. (A short middle row of a
  // centred block, the long row under it, reads as a line the measure broke:
  // "Google Brain" over an e-mail address. Typed after "Brain", the words
  // started a row of their own between the two.) Where they do not fit the
  // lines flow, as a paragraph's do.
  const typed = after.length - before.length;
  // Whether the text is `before` with words typed right before the line that
  // starts at `start`: all before the blank that line follows as it was, and
  // all from the line's start on. (Not read from the common ends: words that
  // begin as the line does share a letter with it.)
  const adjoins = (start: number) => typed > 0 && start > 0 &&
    before.slice(start).every((character, index) => character === after[start + typed + index]) &&
    before.slice(0, start - 1).every((character, index) => character === after[index]);
  if (rowEdit && rowEdit.rowStarts.length === extents.length - 1 && typed > 0) {
    for (let line = 0; line + 1 < extents.length; line += 1) {
      const next = starts[line + 1];
      const count = counts[line] ?? 0;
      const width = extents[line].right - extents[line].left;
      if (next === undefined || breaks.includes(next) || !(count > 0) || !(width > 0)) continue;
      if (!adjoins(next) || rowEdit.rowStarts[line] !== next + typed) continue;
      const wider = (width / count) * typed * 1.15;
      const room = paragraph.layout.alignment === "center"
        ? paragraph.layout.frameRight - paragraph.layout.frameLeft - width
        : paragraph.layout.frameRight - extents[line].right;
      if (wider <= room) { breaks.push(next); breakLines.push(line); }
    }
  }
  if (!breaks.length) return replacementText;
  // Words typed right at a kept break read the same whether they were typed
  // at the end of the line before it or at the start of the line after it
  // (the lines are joined by the one blank): the break then stands before
  // them, unless the editor says where its rows start (rowEdit) -- typed
  // after "2100" of a label "2100 / & beyond", the words stood on the row
  
  const delta = after.length - before.length;
  const rowStarts = rowEdit && rowEdit.rowStarts.length === extents.length - 1 ? rowEdit.rowStarts : null;
  const at = breaks.flatMap((index, order) => {
    const stays = index <= head, moves = index >= before.length - tail;
    if (rowStarts && adjoins(index)) {
      const start = rowStarts[breakLines[order]];
      if (start === index + delta || start === index) {
        if (decided) decided.byRows = true;
        return [start];
      }
    }
    return stays ? [index] : moves ? [index + delta] : [];
  });
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

/** Where the editor's rows start in the new text (an offset for each row
 * after the first, the rows being the box's lines as it showed them): what
 * tells words typed at the end of a line from words typed at the start of
 * the next, which read the same (withSourceHardBreaks). `mustDecide`: the
 * plan is refused unless the rows settled such a break -- the caller then
 * plans as it would without them. */
export type MraiNativeParagraphRowEdit = Readonly<{ rowStarts: readonly number[]; mustDecide?: boolean }>;

export async function planMraiNativeParagraphRewrite(
  kernel: MraiNativeParagraphRewriteKernel,
  selection: PdfNativeListTextSelectionEvidence,
  replacementText: string,
  substitute?: MraiNativeParagraphRewriteSubstitute,
  frame?: MraiNativeParagraphRewriteFrame,
  resolveFormatsAsAsked?: MraiNativeParagraphRewriteFormatsResolver,
  rowEdit?: MraiNativeParagraphRowEdit,
): Promise<PdfNativeListTextReplacementPlan & Readonly<{ substitute?: Readonly<{ resourceName: string; text: string }> }>> {
  const resolveFormats = resolveFormatsAsAsked && formatsWithColoursAsWritten(resolveFormatsAsAsked);
  // A colour, a stroked weight or a rule asked of the whole box with its
  // text as it stands is written by state alone where that is proven
  // (planFormatByStateWhereTextStands, asked first): no show is emptied or
  // drawn again, so nothing of the page is drawn in another order than it
  // was. A block that draws the paragraph again stands after the text
  // object it was drawn in, and the reader then comes to its text after
  // whatever that object drew next: the paragraph read the same, and the
  // boxes beside it were grouped otherwise on about one page in seven
  
  const standing = Boolean(resolveFormats) && replacementText === selection.item.str && frame === undefined &&
    !substitute && !rowEdit && replacementText.length <= 2_000;
  if (standing && resolveFormats!.stateOnly && kernel.standingPageStreams && kernel.standingStreamContext) {
    try {
      return await planFormatByStateWhereTextStands(kernel, selection, resolveFormats!,
        new PdfNativeTransactionError("unsupported-document", MRAI_NATIVE_PARAGRAPH_FORMAT_IN_PLACE_REFUSAL), true);
    } catch (error) {
      
      if (!(error instanceof PdfNativeTransactionError)) throw error;
      // Proven, and the box is drawn as asked already -- every show of it
      // under the weight and colour, every rule standing: that is the answer.
      
      // that changed nothing the reader sees and set the run apart from its
      // paragraph (下划线 asked again of an underlined run of pdf-walk-1000
      // 0294: 1.5 KB written, the page's three boxes read as seven).
      if (error.message === MRAI_NATIVE_STANDING_FORMAT_UNCHANGED_REFUSAL) throw error;
      // Or its bold is a face's own: the words set again in the family's
      // regular face where they stand, where every word of the box can be
      // (planBoldFaceSetRegularWhereWordsStand); refused as it was where not.
      if (error.message === MRAI_NATIVE_BOLD_FACE_KEPT_REFUSAL) {
        const set = await planBoldFaceSetRegularWhereWordsStand(kernel, selection, resolveFormats!);
        if (set) return set;
        throw error;
      }
      // (Or a lean it will not write because the page would read otherwise
      // with the box's shows where they stand: drawn again in a block they
      // would read otherwise too -- refuseLeanReadAsOtherLinesAmongOtherText.)
      if (FINAL_STANDING_REFUSALS.has(error)) throw error;
    }
  }
  try {
    return await planParagraphRewriteDrawnAgain(kernel, selection, replacementText, substitute, frame, resolveFormats, rowEdit);
  } catch (error) {
    // A format of the whole box with its text as it stands, refused only
    // because the box cannot be drawn again whole: written by state alone
    // (which refuses with this refusal whatever is not that). Every other
    // edit keeps its refusal.
    // (Or because a show of it is a Type3 show a weight was asked of: the
    
    // one writes a second pass of it where it stands.)
    if (!(error instanceof PdfNativeTransactionError) || !standing ||
      !(refusalOfRedrawOnly(error.message) || error.message === MRAI_NATIVE_TYPE3_FORMAT_MESSAGE)) throw error;
    try {
      return await planFormatByStateWhereTextStands(kernel, selection, resolveFormats!, error);
    } catch (kept) {
      // (Its bold a face's own, read here by where the reader's items start: the words set again in the
      // regular face where that writer reads the box -- planBoldFaceSetRegularWhereWordsStand -- or kept.)
      if (!(kept instanceof PdfNativeTransactionError) || kept.message !== MRAI_NATIVE_BOLD_FACE_KEPT_REFUSAL) throw kept;
      const set = await planBoldFaceSetRegularWhereWordsStand(kernel, selection, resolveFormats!);
      if (set) return set;
      throw kept;
    }
  }
}



async function planParagraphRewriteDrawnAgain(
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
  rowEdit?: MraiNativeParagraphRowEdit,
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
    // A format that moves no glyph, asked of the whole box with its text as
    // it stands, sets nothing from the reader's string: it is written where
    // the text stands, by the page's own shows (planFormatWhereReorderedTextStands,
    // which refuses whatever is not that). Everything else stays refused.
    if (replacementText !== selection.item.str || !resolveFormats || frame !== undefined || substitute || rowEdit ||
      !kernel.standingPageStreams) refuse(MRAI_NATIVE_REORDERED_TEXT_MESSAGE);
    return planFormatWhereReorderedTextStands(kernel, selection, resolveFormats!);
  }
  // Formats asked of text that stays may move no glyph of it, and then draw
  // it where it stands (setParagraphWhereItStands): read for that.
  const standing = replacementText === selection.item.str && Boolean(resolveFormats) &&
    frame === undefined && !substitute && !rowEdit;
  let read: ReadParagraph;
  try {
    read = await readParagraph(kernel, selection, standing);
  } catch (error) {
    // No reader item proof reads the box whole: such a format needs none
    // (planFormatWhereUnprovedTextStands, which refuses with this refusal
    // whatever is not that).
    if (!standing || !resolveFormats || !(error instanceof PdfNativeTransactionError) ||
      !standsUnproved(kernel, selection, error)) throw error;
    return planFormatWhereUnprovedTextStands(kernel, selection, resolveFormats, error);
  }
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
  
  const decided = { byRows: false };
  replacementText = withSourceHardBreaks(paragraph, selection.item.str, replacementText, frame?.right !== undefined,
    rowEdit, decided);
  if (rowEdit?.mustDecide && !decided.byRows) refuse("The editor's rows settle no line break of this edit.");
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
  let missing = !substitute ? "" : wholly
    ? [...new Set(characters(replacementText.replace(/\n/gu, "")).filter((character) => !isBlank(character)))].join("")
    : charactersNoParagraphFontWrites(kernel, paragraph, replacementText.replace(/\n/gu, ""));
  let installed = missing ? await substitute!.install(missing) : null;
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
  // A new line that would be drawn over other text, where the words fit on
  // their own line in the room up to the text beside it: they stay on that
  // line -- past the paragraph's measure, and nearer the text beside it than
  // a paragraph otherwise comes. An author's affiliation typed two letters
  // longer had its own e-mail address under it: kept to the width of the
  // name above, the letters went down a line, onto the address (Attention
  
  // on with it at worst; over it, both are spoilt.
  if (!frame && !formats && paragraph.nearFrameRight !== undefined &&
    set.notices.some((notice) => notice.code === "new-line-over-other-text")) {
    const nearer = Object.freeze({ ...paragraph, nearFrameRight: undefined,
      layout: Object.freeze({ ...paragraph.layout, frameRight: paragraph.nearFrameRight }) });
    try {
      const again = setParagraph(kernel, nearer, replacementText, installed, formats, wholly);
      const last = paragraph.layout.lineCount - 1;
      // Taken where it gives the words room on the paragraph's own last
      // line and changes no line before it.
      if (!again.notices.some((notice) => notice.code === "new-line-over-other-text" || notice.code === "line-over-other-text") &&
        again.lines.length === paragraph.layout.lineCount &&
        again.lines.slice(0, last).every((line, index) => line.start === set.lines[index]?.start && line.end === set.lines[index]?.end)) {
        set = again;
        paragraph = nearer;
      }
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError)) throw error;
    }
  }
  // The face is installed with the glyphs the block draws in it, no more.
  // The characters no font of the paragraph writes are those of its whole
  // text, and the lines the edit leaves alone are drawn again by the codes
  // the page drew them with: a word the page sets with a ligature ("flooding",
  // its font holding "fl" and no "f") is such a line's, and its "f" was asked
  // of the face, installed and never drawn -- which no write may carry (a
  // package is proven by the glyphs drawn in it). A label of that word typed
  // onto could not be written in place at all, and was taken apart as a text
  // box in another face and size (IPCC SPM.1's "Inland flooding and
  
  // what was drawn, the block set once more names the same glyphs by the
  // narrower face's own codes.
  if (installed && !wholly) {
    const drawn = characters(missing).filter((character) => set.substituteCharacters.has(character)).join("");
    if (drawn && drawn !== missing) {
      const narrowed = await substitute!.install(drawn);
      if (narrowed) {
        const again = setParagraph(kernel, paragraph, replacementText, narrowed, formats, wholly);
        if (again.substituteCharacters.size !== characters(drawn).length ||
          characters(drawn).some((character) => !again.substituteCharacters.has(character))) {
          refuse("The paragraph does not draw the characters its substitute face was installed for.");
        }
        set = again;
        installed = narrowed;
        missing = drawn;
      }
    }
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





