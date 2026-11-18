/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The deck editor's Font and Paragraph buttons on a text block, as pure
 * transforms of the block: character formatting over a run range (or the
 * whole box), Clear Formatting, Change Case, bullets and numbering, list
 * level, paragraph alignment and spacing, columns, autofit and the box's own
 * style. The ribbon (`PresentationWindow`) passes the person's live text
 * selection; an editor command passes none, which is the whole box. Both
 * commit what these return.
 */

import { withCommandStatedKeys } from "./presentation-command-stated-keys";
import { presentationTextUsesPoints } from "./presentation/text-units-block";
import {
  CSS_PX_PER_POINT,
  PRESENTATION_DEFAULT_TEXT_BOX_SIZE_PT,
  PRESENTATION_DEFAULT_TEXT_BOX_SIZE_PX,
} from "./presentation/text-units";
import {
  blockHoldsEquationRun,
  withEquationFormatMaterialized,
  withEquationFormatToggle,
} from "./presentation-equation-run-italic";
import { wordSegmentsOf } from "./document-word-at";
import {
  type PresentationBlock,
  type PresentationTextBlock,
  type PresentationTextStyle,
} from "../components/console/editor/native-document/documentCommands";
import {
  effectiveAutofitMode,
  nativeAutofitResizesShape,
  type AutofitMode,
} from "./presentation-autofit";
import {
  normalizeColumnSpacing,
  persistedColumnCount,
} from "./presentation-text-columns";
import type {
  PresentationChartBlock,
  PresentationShapeBlock,
} from "../components/console/editor/native-document/documentCommands";
import { stepOfficeFontSize } from "../components/office-ribbon/officeFontSizeCatalog";

/** Slide sizes are px at 96 dpi; the size box and the list are points. */
const SLIDE_PX_PER_POINT = CSS_PX_PER_POINT;

/**
 * The deck's Grow / Shrink Font: PowerPoint's own size list, in points, for a
 * size held in slide px; returns the next size as a "NNpx" field value (what
 * the size parsers read), or null. Replaces the old Docs px ladder, which
 * topped out at 48 px (36 pt), so one Shrink past it fell back to 36 pt.
 */
export function stepFontSize(currentPx: number | null | undefined, direction: 1 | -1): string | null {
  const px = typeof currentPx === "number" && Number.isFinite(currentPx) && currentPx > 0 ? currentPx : 24;
  const currentPt = Math.round((px / SLIDE_PX_PER_POINT) * 10) / 10;
  const nextPt = stepOfficeFontSize(currentPt, direction, "powerpoint");
  if (nextPt === currentPt) return null;
  return `${Math.round(nextPt * SLIDE_PX_PER_POINT * 100) / 100}px`;
}
import type { OfficeHomeTextCommand } from "./office-home-object-target";
import type { SmartArtNodeTextFormatPatch } from "./office-smartart/text-format";
import {
  applyPresentationChartHomeText,
  applyPresentationHomeObjectAnchor,
  applyPresentationHomeObjectRunMark,
  applyPresentationHomeObjectText,
  type PresentationHomeObjectRunMark,
  type PresentationHomeObjectTargets,
} from "./presentation-home-object-text";
import { normalizeLayoutId } from "./presentation-layout-id";
import type { NativeListIndent } from "./native-drawing-geometry";
import type { PresentationParagraphIndentChange } from "./presentation-paragraph-dialog";
import { PRESENTATION_SMARTART_LABEL_TEXT_ADAPTER } from "./presentation-smartart-text";
import {
  presentationLevelStylesOf,
  presentationListLevelRestyle,
  runStyleHandedToLevel,
  type PresentationLevelStyle,
} from "./presentation-level-styles";
import {
  applyCaseToRuns,
  applyCaseToText,
  type ChangeCaseMode,
} from "./presentation-text-case";
import {
  normalizeExactLineSpacingPoints,
  type ParagraphSpacing,
  readSpacingFromStyle,
} from "./presentation-line-spacing";
import {
  adjustParagraphLevelAt,
  type BulletChar,
  MAX_PARAGRAPH_LEVEL,
  normalizeBlockBulletChar,
  normalizeBulletChar,
  normalizeNumberFormat,
  type NumberFormat,
  clampLevel,
  dropNativeParagraphReceipts,
  materializeNativeBulletReceipt,
  hasAnyBullet,
  MAX_PARAGRAPH_TEXT_LENGTH,
  MAX_PARAGRAPHS_PER_BLOCK,
  foldParagraphRecordsPastLimit,
  normalizeBullet,
  normalizeParagraphAlign,
  normalizeParagraphDirection,
  normalizeParagraphIndentEmu,
  type ParagraphAlign,
  type ParagraphRunDefaults,
  type PresentationFontAlign,
  paragraphIndexAtOffset,
  paragraphMarkerFields,
  paragraphCarriedFields,
  paragraphsFromText,
  paragraphStatesIndent,
  type SlideParagraph,
  textFromParagraphs,
} from "./presentation-paragraphs";
import { rowsWithLayoutHiddenMarkers } from "./presentation-smartart-markers";
// A cycle (that module formats runs through this one); both are only called, never read at load.
import {
  presentationBlockWithCanonicalTextPaint,
  presentationBlockWithRunTextPaints,
} from "./presentation-run-text-paint";
import {
  applyRunFormat,
  applyStylePatch,
  normalizeRuns,
  normalizeRunStyle,
  paragraphRunSelectionRanges,
  parseRunFontSizeInput as parsePointRunFontSizeInput,
  resolveBooleanRunToggle,
  runCollectionsEqual,
  runFieldAtSelection,
  type ToggleableRunStyleKey,
  runsFromText,
  type RunStyle,
  textFromRuns,
  type TextRun,
} from "./presentation-point-runs";

function asRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

export const FONT_RUN_STYLE_KEYS = [
  "bold",
  "italic",
  "underline",
  "strike",
  "superscript",
  "subscript",
  "color",
  "highlight",
  "fontSize",
  "fontFamily",
  "textOutline",
  "textOutlineWidthPt",
  "textOutlineJoin",
  "textOutlineMiterLimit",
  "textOutlineCap",
  "textOutlineDash",
] as const satisfies readonly (keyof RunStyle)[];

export const FONT_RUN_CLEAR_PATCH = {
  ...Object.fromEntries(FONT_RUN_STYLE_KEYS.map((key) => [key, undefined])),
  // A run's own text shadow (readings row 808) is cleared with the rest.
  textShadow: undefined,
  textShadowEffect: undefined,
  // And its WordArt: Text Fill, Text Outline and Text Effects are `a:rPr`'s,
  // which Clear All Formatting empties (truth c64cleartx: fill, `a:ln`, glow
  // and shadow all gone; the body's Transform kept).
  textPaint: undefined,
} as Readonly<Partial<RunStyle>>;

// Only these fields are authenticated block-level inheritance in the current
// presentation renderer. Run-only fields (strike/baseline/highlight/outline)
// must never be promoted from an opaque legacy block.style record.
export const FONT_BLOCK_INHERITABLE_RUN_STYLE_KEYS = [
  "bold",
  "italic",
  "underline",
  "color",
  "fontSize",
  "fontFamily",
] as const satisfies readonly (keyof RunStyle)[];

export const CARET_WORD_CHARACTER = /^[\p{L}\p{M}\p{N}_]$/u;

/**
 * The word the insertion point is inside: word characters on both sides of
 * the caret, stretched to the word's ends. A caret at a word's edge, between
 * spaces or at either end of the text owns no word and yields the empty range
 * at the caret, which a writer treats as nothing to format.
 */
export function caretWordRange(text: string, caret: number): readonly [number, number] {
  const before = (index: number): number => {
    if (index <= 0) return 0;
    const low = text.charCodeAt(index - 1);
    const width = low >= 0xdc00 && low <= 0xdfff && index >= 2 ? 2 : 1;
    return CARET_WORD_CHARACTER.test(text.slice(index - width, index)) ? width : 0;
  };
  const after = (index: number): number => {
    if (index >= text.length) return 0;
    const point = text.codePointAt(index) ?? 0;
    return CARET_WORD_CHARACTER.test(String.fromCodePoint(point))
      ? point > 0xffff ? 2 : 1
      : 0;
  };
  if (before(caret) === 0 || after(caret) === 0) return [caret, caret];
  let from = caret;
  let to = caret;
  for (let width = before(from); width > 0; width = before(from)) from -= width;
  for (let width = after(to); width > 0; width = after(to)) to += width;
  // A run of letters is not always one word: Chinese and Japanese have no
  // spaces, and the run was the whole sentence (a caret in it bolded every
  // character). The run's own words decide, as Docs' word-at does.
  const run = text.slice(from, to);
  const at = caret - from;
  const word = wordSegmentsOf(run).find(
    (segment) => segment.index <= at && at <= segment.index + segment.segment.length && segment.isWordLike !== false,
  );
  if (!word || word.index === at || word.index + word.segment.length === at) return [caret, caret];
  return [from + word.index, from + word.index + word.segment.length];
}

/**
 * The characters a character command acts on. No text selection means the
 * whole object (the frame is selected); a highlighted range is itself; a bare
 * caret is the word it sits in and never the rest of the box, which is where
 * PowerPoint puts character formatting when nothing is highlighted (PPT-B006).
 */
export function selectedTextRange(
  blockId: string,
  text: string,
  selection: { blockId: string; start: number; end: number } | null,
): readonly [number, number] {
  const total = text.length;
  if (!selection || selection.blockId !== blockId) return [0, total];
  const start = Math.min(Math.max(Math.round(selection.start), 0), total);
  const end = Math.min(Math.max(Math.round(selection.end), 0), total);
  // A bare caret formats no existing character: the format is held for what
  
  // formatted with nothing highlighted). Not Office's word-at-caret rule.
  return end > start ? [start, end] : [start, start];
}

/** A text block's editable text, as its body's selection offsets address it. */
export function presentationEditableText(block: PresentationTextBlock): string {
  const paragraphs = normalizeParagraphsWithRuns(block.paragraphs);
  return paragraphs && ownsParagraphRichLayout(paragraphs)
    ? textFromParagraphs(paragraphs)
    : block.text;
}

/** The length of a text block's editable text, paragraphs included. */
export function presentationTextLength(block: PresentationTextBlock): number {
  return presentationEditableText(block).length;
}

/** Read autofit from the canonical block field, the imported style layer, or
 * the imported box's own `a:spAutoFit` (`effectiveAutofitMode`). */
export function autofitModeForBlock(block: PresentationBlock) {
  return effectiveAutofitMode(block);
}

export type PresentationParagraphSpacing = Pick<
  ParagraphSpacing,
  "lineSpacing" | "spaceBefore" | "spaceAfter"
> & {
  /**
   * The row's Line Spacing > Exactly, in points (`a:lnSpc/a:spcPts`): every
   * line of the row that tall whatever its type size. A row states a multiple
   * or an exact height, never both; picking a multiple replaces it.
   */
  lineSpacingPoints?: number;
};

export const PRESENTATION_PARAGRAPH_SPACING_FIELDS = [
  "lineSpacing",
  "spaceBefore",
  "spaceAfter",
] as const;

export type PresentationParagraphSpacingField =
  (typeof PRESENTATION_PARAGRAPH_SPACING_FIELDS)[number];

/** A list paragraph plus its optional per-run and per-paragraph formatting. */
export type SlideParagraphWithRuns = SlideParagraph &
  PresentationParagraphSpacing & {
    runs?: TextRun[];
    spacingFallbackBlocked?: PresentationParagraphSpacingField[];
  };

export function normalizedSpacingFallbackBlocked(
  value: unknown,
): PresentationParagraphSpacingField[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const selected = new Set(
    value
      .slice(0, PRESENTATION_PARAGRAPH_SPACING_FIELDS.length)
      .filter(
        (entry): entry is PresentationParagraphSpacingField =>
          typeof entry === "string" &&
          PRESENTATION_PARAGRAPH_SPACING_FIELDS.some(
            (field) => field === entry,
          ),
      ),
  );
  const normalized = PRESENTATION_PARAGRAPH_SPACING_FIELDS.filter((field) =>
    selected.has(field),
  );
  return normalized.length > 0 ? [...normalized] : undefined;
}

export function spacingFallbackBlocksEqual(left: unknown, right: unknown): boolean {
  const leftFields = normalizedSpacingFallbackBlocked(left) ?? [];
  const rightFields = normalizedSpacingFallbackBlocked(right) ?? [];
  return (
    leftFields.length === rightFields.length &&
    leftFields.every((field, index) => field === rightFields[index])
  );
}

export function normalizedParagraphSpacingFields(
  value: unknown,
): PresentationParagraphSpacing {
  if (!value || typeof value !== "object") return {};
  const normalized = readSpacingFromStyle(value as Record<string, unknown>);
  const spacing: PresentationParagraphSpacing = {};
  if (normalized?.lineSpacing !== undefined) {
    spacing.lineSpacing = normalized.lineSpacing;
  }
  if (normalized?.spaceBefore !== undefined) {
    spacing.spaceBefore = normalized.spaceBefore;
  }
  if (normalized?.spaceAfter !== undefined) {
    spacing.spaceAfter = normalized.spaceAfter;
  }
  const exact = normalizeExactLineSpacingPoints(
    (value as Record<string, unknown>).lineSpacingPoints,
  );
  if (exact !== undefined && spacing.lineSpacing === undefined) {
    spacing.lineSpacingPoints = exact;
  }
  return spacing;
}

export function hasParagraphSpacing(paragraph: PresentationParagraphSpacing): boolean {
  return (
    paragraph.lineSpacing !== undefined ||
    paragraph.lineSpacingPoints !== undefined ||
    paragraph.spaceBefore !== undefined ||
    paragraph.spaceAfter !== undefined
  );
}

export function hasParagraphFallbackBlock(
  paragraph: SlideParagraphWithRuns,
): boolean {
  return Boolean(
    normalizedSpacingFallbackBlocked(paragraph.spacingFallbackBlocked),
  );
}

export function ownsParagraphRichLayout(
  paragraphs: readonly SlideParagraphWithRuns[],
): boolean {
  return (
    hasAnyBullet(paragraphs) ||
    paragraphs.some(
      (paragraph) =>
        paragraph.align !== undefined ||
        hasParagraphSpacing(paragraph) ||
        hasParagraphFallbackBlock(paragraph) ||
        Boolean(normalizeRuns(paragraph.runs)),
    )
  );
}

/**
 * Normalize a block's extended paragraph records. The shared core owns list
 * semantics; imported PowerPoint paragraphs additionally carry authenticated
 * runs and their own line/before/after spacing. A flat-equivalent array remains
 * structured only when one of those paragraph-local fields would otherwise be
 * lost.
 */
export function normalizeParagraphsWithRuns(
  raw: unknown,
): SlideParagraphWithRuns[] | undefined {
  if (!Array.isArray(raw)) return undefined;
  const paragraphs: SlideParagraphWithRuns[] = [];
  for (const source of foldParagraphRecordsPastLimit(raw)) {
    if (paragraphs.length >= MAX_PARAGRAPHS_PER_BLOCK) break;
    if (!source || typeof source !== "object") continue;
    const record = source as Record<string, unknown>;
    const runs = normalizeRuns(record.runs);
    const rawText =
      typeof record.text === "string"
        ? record.text.slice(0, MAX_PARAGRAPH_TEXT_LENGTH)
        : "";
    const paragraph: SlideParagraphWithRuns = {
      text: runs ? textFromRuns(runs) : rawText,
      level: clampLevel(record.level),
      bullet: normalizeBullet(record.bullet),
      ...normalizedParagraphSpacingFields(record),
    };
    const align = normalizeParagraphAlign(record.align);
    if (align !== undefined) paragraph.align = align;
    // The row's own `a:pPr@marL`/`@indent`, in EMU. Kept here as well as in
    // `normalizeParagraphs` so the RENDER path reads the same row model the
    // save path bounds -- dropping it here would draw an indented label at the
    // margin while the package and the page both say otherwise.
    const marL = normalizeParagraphIndentEmu(record.marL, { signed: false });
    if (marL !== undefined) paragraph.marL = marL;
    const indent = normalizeParagraphIndentEmu(record.indent, { signed: true });
    if (indent !== undefined) paragraph.indent = indent;
    const dir = normalizeParagraphDirection(record.dir);
    if (dir !== undefined) paragraph.dir = dir;
    Object.assign(paragraph, paragraphMarkerFields(record, paragraph.bullet));
    Object.assign(paragraph, paragraphCarriedFields(record));
    if (runs) paragraph.runs = runs;
    const spacingFallbackBlocked = normalizedSpacingFallbackBlocked(
      record.spacingFallbackBlocked,
    );
    if (spacingFallbackBlocked) {
      paragraph.spacingFallbackBlocked = spacingFallbackBlocked;
    }
    paragraphs.push(paragraph);
  }
  if (paragraphs.length === 0) return undefined;
  const flatEquivalent = paragraphs.every(
    (paragraph) =>
      paragraph.bullet === "none" &&
      paragraph.level === 0 &&
      paragraph.align === undefined &&
      paragraph.dir === undefined &&
      paragraph.fontAlign === undefined &&
      paragraph.fontAlignInherited === undefined &&
      !paragraphStatesIndent(paragraph),
  );
  if (
    flatEquivalent &&
    !paragraphs.some(
      (paragraph) =>
        paragraph.align !== undefined ||
        hasParagraphSpacing(paragraph) ||
        hasParagraphFallbackBlock(paragraph) ||
        Boolean(normalizeRuns(paragraph.runs)),
    )
  ) {
    return undefined;
  }
  return paragraphs;
}

export type LiveTextSelection = {
  blockId: string;
  start: number;
  end: number;
} | null;

/** Split a flat block's authenticated runs across its newline paragraphs. */
export function paragraphsWithRunsFromTextBlock(
  block: PresentationTextBlock,
): SlideParagraphWithRuns[] {
  const existing = normalizeParagraphsWithRuns(block.paragraphs);
  if (existing) return existing;
  return withFlatRunsSplitAcrossParagraphs(
    paragraphsFromText(block.text),
    block.runs,
    block.text,
  );
}

/**
 * Give each row its share of the block's flat runs, cut at the row breaks of
 * ``text``. Rows keep everything else they state. A run collection that does
 * not spell ``text`` exactly is stale and splits nothing (fail closed).
 */
export function withFlatRunsSplitAcrossParagraphs(
  paragraphs: SlideParagraphWithRuns[],
  rawRuns: unknown,
  text: string,
): SlideParagraphWithRuns[] {
  const flatRuns = normalizeRuns(rawRuns);
  if (
    !flatRuns ||
    textFromRuns(flatRuns) !== text ||
    textFromParagraphs(paragraphs) !== text
  ) {
    return paragraphs;
  }
  let paragraphStart = 0;
  return paragraphs.map((paragraph, paragraphIndex) => {
    const paragraphEnd = paragraphStart + paragraph.text.length;
    const pieces: TextRun[] = [];
    let runStart = 0;
    for (const run of flatRuns) {
      const runEnd = runStart + run.text.length;
      const from = Math.max(paragraphStart, runStart);
      const to = Math.min(paragraphEnd, runEnd);
      if (from < to) {
        const text = run.text.slice(from - runStart, to - runStart);
        // The run keeps what it IS -- an equation, a field -- as well as its
        // style: an equation in a placeholder's inherited rows drew as its
        // bare characters (the owner's "(x+a)n=k=0n..." while editing). An
        // equation or a field is whole or not at all.
        const whole = from === runStart && to === runEnd;
        const piece: TextRun = { ...run, text, ...(run.style ? { style: { ...run.style } } : {}) };
        if (!whole) {
          delete (piece as { math?: unknown }).math;
          delete (piece as { field?: unknown }).field;
        }
        pieces.push(piece);
      }
      runStart = runEnd;
    }
    paragraphStart =
      paragraphEnd + (paragraphIndex < paragraphs.length - 1 ? 1 : 0);
    const runs = normalizeRuns(pieces);
    return runs ? { ...paragraph, runs } : paragraph;
  });
}

/** Rejoin per-paragraph runs when a list collapses back to flat text. */
export function flatRunsFromParagraphs(
  paragraphs: readonly SlideParagraphWithRuns[],
): TextRun[] | undefined {
  const runs: TextRun[] = [];
  paragraphs.forEach((paragraph, index) => {
    runs.push(...(normalizeRuns(paragraph.runs) ?? runsFromText(paragraph.text)));
    if (index < paragraphs.length - 1) runs.push({ text: "\n" });
  });
  return normalizeRuns(runs);
}

/** Resolve the exact paragraph span touched by the live textarea selection. */
export function paragraphSelectionSpan(
  blockId: string,
  paragraphs: readonly SlideParagraphWithRuns[],
  selection: LiveTextSelection,
): readonly [number, number] {
  if (!selection || selection.blockId !== blockId) {
    return [0, Math.max(0, paragraphs.length - 1)];
  }
  const total = textFromParagraphs(paragraphs).length;
  const rawStart = Math.min(Math.max(Math.round(selection.start), 0), total);
  const rawEnd = Math.min(Math.max(Math.round(selection.end), 0), total);
  const from = Math.min(rawStart, rawEnd);
  const to = Math.max(rawStart, rawEnd);
  const first = paragraphIndexAtOffset(paragraphs, from);
  const last = paragraphIndexAtOffset(
    paragraphs,
    to > from ? Math.max(from, to - 1) : from,
  );
  return [Math.min(first, last), Math.max(first, last)];
}

export const PRESENTATION_PARAGRAPH_FORMAT_FIELDS = [
  "align",
  ...PRESENTATION_PARAGRAPH_SPACING_FIELDS,
] as const;

export type PresentationParagraphFormatField =
  (typeof PRESENTATION_PARAGRAPH_FORMAT_FIELDS)[number];

export type PresentationParagraphFormatPatch = Partial<
  Pick<PresentationTextStyle, PresentationParagraphFormatField>
>;

export type PresentationParagraphFormatState = {
  align: ParagraphAlign | "mixed";
  lineSpacing: number | undefined | "mixed";
  spaceBefore: number | undefined | "mixed";
  spaceAfter: number | undefined | "mixed";
};

export function paragraphFormatValue(
  paragraph: SlideParagraphWithRuns,
  field: PresentationParagraphFormatField,
  fallback: PresentationTextStyle | undefined,
): ParagraphAlign | number | undefined {
  if (field === "align") {
    return (
      normalizeParagraphAlign(paragraph.align) ??
      normalizeParagraphAlign(fallback?.align) ??
      "left"
    );
  }
  const own = normalizedParagraphSpacingFields(paragraph)[field];
  if (own !== undefined) return own;
  if (
    normalizedSpacingFallbackBlocked(paragraph.spacingFallbackBlocked)?.includes(
      field,
    )
  ) {
    return undefined;
  }
  return normalizedParagraphSpacingFields(fallback)[field];
}

/** Derive one honest Ribbon state from the exact selected paragraph span. */
export function paragraphFormatSelectionState(
  block: PresentationTextBlock,
  selection: LiveTextSelection,
): PresentationParagraphFormatState {
  const paragraphs = paragraphsWithRunsFromTextBlock(block);
  const [first, last] = paragraphSelectionSpan(
    block.id,
    paragraphs,
    selection,
  );
  const selected = paragraphs.slice(first, last + 1);
  const stateFor = <T extends ParagraphAlign | number | undefined>(
    field: PresentationParagraphFormatField,
  ): T | "mixed" => {
    const values = selected.map((paragraph) =>
      paragraphFormatValue(paragraph, field, block.style),
    );
    const initial = values[0];
    return values.every((value) => Object.is(value, initial))
      ? (initial as T)
      : "mixed";
  };
  return {
    align: stateFor<ParagraphAlign>("align"),
    lineSpacing: stateFor<number | undefined>("lineSpacing"),
    spaceBefore: stateFor<number | undefined>("spaceBefore"),
    spaceAfter: stateFor<number | undefined>("spaceAfter"),
  };
}

/**
 * Move a text-box paragraph default onto every unaffected paragraph, then
 * change only the live selected span. This is the paragraph equivalent of
 * materialising inherited run formatting: removing a block-level value must
 * never restyle the unselected rows.
 */
export function applyParagraphFormatSelection(
  block: PresentationTextBlock,
  selection: LiveTextSelection,
  patch: PresentationParagraphFormatPatch,
): PresentationTextBlock | null {
  const fields = PRESENTATION_PARAGRAPH_FORMAT_FIELDS.filter((field) =>
    Object.prototype.hasOwnProperty.call(patch, field),
  );
  if (fields.length === 0) return null;
  const paragraphs = paragraphsWithRunsFromTextBlock(block);
  if (paragraphs.length === 0) return null;
  const [first, last] = paragraphSelectionSpan(
    block.id,
    paragraphs,
    selection,
  );

  const requestedValue = (
    field: PresentationParagraphFormatField,
  ): ParagraphAlign | number | undefined => {
    if (field === "align") {
      return normalizeParagraphAlign(patch.align) ?? "left";
    }
    return normalizedParagraphSpacingFields(patch)[field];
  };
  const selectionAlreadyMatches = paragraphs
    .slice(first, last + 1)
    .every((paragraph) =>
      fields.every(
        (field) =>
          // A row at Line Spacing > Exactly is at no multiple: any line
          // spacing asked of it is a change.
          !(
            field === "lineSpacing" &&
            normalizedParagraphSpacingFields(paragraph).lineSpacingPoints !==
              undefined
          ) &&
          Object.is(
            paragraphFormatValue(paragraph, field, block.style),
            requestedValue(field),
          ),
      ),
    );
  // Effective equality is the authority boundary. Inherited block defaults
  // must not be materialized merely because their paragraph-local XML/model
  // representation differs from the requested Ribbon value.
  if (selectionAlreadyMatches) return null;

  const blockStyle = { ...(block.style ?? {}) } as PresentationTextStyle;
  let blockStyleChanged = false;
  for (const field of fields) {
    if (Object.prototype.hasOwnProperty.call(blockStyle, field)) {
      delete blockStyle[field];
      blockStyleChanged = true;
    }
  }
  const migrationBlock = { ...block } as PresentationTextBlock;
  if (blockStyleChanged) {
    if (Object.keys(blockStyle).length > 0) migrationBlock.style = blockStyle;
    else delete migrationBlock.style;
  }

  let paragraphChanged = false;
  const nextParagraphs = paragraphs.map((paragraph, index) => {
    const next: SlideParagraphWithRuns = { ...paragraph };
    let changed = false;
    for (const field of fields) {
      const selected = index >= first && index <= last;
      if (field === "align") {
        const current = normalizeParagraphAlign(paragraph.align);
        const materialized = paragraphFormatValue(
          paragraph,
          field,
          block.style,
        ) as ParagraphAlign;
        const requested = selected
          ? normalizeParagraphAlign(patch.align)
          : materialized;
        // Left is written as left wherever someone chose it or something
        // stated it: dropped as "the default", a title whose box or layout
        // centres it took Left back to centred here and in PowerPoint, while
        // the Ribbon lit Left. A row nothing ever aligned stays unstated.
        const stated =
          selected ||
          current !== undefined ||
          normalizeParagraphAlign(block.style?.align) !== undefined;
        const persisted = requested === "left" && !stated ? undefined : requested;
        if (current !== persisted) changed = true;
        if (persisted === undefined) delete next.align;
        else next.align = persisted;
        continue;
      }

      const current = normalizedParagraphSpacingFields(paragraph)[field];
      const materialized = paragraphFormatValue(
        paragraph,
        field,
        block.style,
      ) as number | undefined;
      const requested = selected
        ? normalizedParagraphSpacingFields(patch)[field]
        : materialized;
      if (!Object.is(current, requested)) changed = true;
      if (requested === undefined) delete next[field];
      else next[field] = requested;
      if (
        field === "lineSpacing" &&
        selected &&
        next.lineSpacingPoints !== undefined
      ) {
        // One line spacing per row: the multiple (or the default) asked for
        // replaces the row's Exactly height.
        delete next.lineSpacingPoints;
        changed = true;
      }

      if (selected) {
        const blocked = (
          normalizedSpacingFallbackBlocked(next.spacingFallbackBlocked) ?? []
        ).filter((candidate) => candidate !== field);
        if (
          !spacingFallbackBlocksEqual(
            next.spacingFallbackBlocked,
            blocked.length > 0 ? blocked : undefined,
          )
        ) {
          changed = true;
        }
        if (blocked.length > 0) next.spacingFallbackBlocked = blocked;
        else delete next.spacingFallbackBlocked;
      }
    }
    if (changed) paragraphChanged = true;
    return changed ? next : paragraph;
  });
  const next = withPersistedParagraphs(migrationBlock, nextParagraphs);
  if (!paragraphChanged && !blockStyleChanged) return null;
  return next;
}

/**
 * The selected rows' Line Spacing > Exactly height (`lineSpacingPoints`), and
 * the first selected row's own indentation, for the Paragraph dialog.
 */
export function paragraphDialogSelectionExtras(
  block: PresentationTextBlock,
  selection: LiveTextSelection,
  inherited?: readonly NativeListIndent[],
): {
  lineSpacingPoints: number | undefined | "mixed";
  marL?: number;
  indent?: number;
  fontAlign?: PresentationFontAlign | "mixed";
} {
  const paragraphs = paragraphsWithRunsFromTextBlock(block);
  if (paragraphs.length === 0) return { lineSpacingPoints: undefined };
  const [first, last] = paragraphSelectionSpan(block.id, paragraphs, selection);
  const selected = paragraphs.slice(first, last + 1);
  const exact = selected.map(
    (paragraph) => normalizedParagraphSpacingFields(paragraph).lineSpacingPoints,
  );
  const head = selected[0];
  const { marL, indent } = head
    ? effectiveParagraphIndent(head, inherited)
    : { marL: undefined, indent: undefined };
  return {
    lineSpacingPoints: exact.every((value) => Object.is(value, exact[0])) ? exact[0] : "mixed",
    ...(marL === undefined ? {} : { marL }),
    ...(indent === undefined ? {} : { indent }),
    fontAlign: uniformFontAlign(selected),
  };
}

/** The Text alignment each row applies: its own, what its level inherits, else Auto. */
function effectiveParagraphFontAlign(
  paragraph: Pick<SlideParagraphWithRuns, "fontAlign" | "fontAlignInherited">,
): PresentationFontAlign {
  return paragraph.fontAlign ?? paragraph.fontAlignInherited ?? "auto";
}

function uniformFontAlign(
  rows: readonly Pick<SlideParagraphWithRuns, "fontAlign" | "fontAlignInherited">[],
): PresentationFontAlign | "mixed" {
  const values = new Set(rows.map(effectiveParagraphFontAlign));
  return values.size === 1 ? [...values][0]! : values.size === 0 ? "auto" : "mixed";
}

/**
 * Paragraph > Asian Typography > Text alignment on the selected rows: each
 * states it as its own `fontAlgn` (PowerPoint writes the attribute on every
 * row the dialog applies to, Auto included). Null when every selected row
 * already states that value.
 */
export function textBlockWithParagraphFontAlign(
  block: PresentationTextBlock,
  fontAlign: PresentationFontAlign,
  selection: LiveTextSelection,
): PresentationTextBlock | null {
  const paragraphs = paragraphsWithRunsFromTextBlock(block);
  if (paragraphs.length === 0) return null;
  const [first, last] = paragraphSelectionSpan(block.id, paragraphs, selection);
  let changed = false;
  const nextParagraphs = paragraphs.map((paragraph, index) => {
    if (index < first || index > last || paragraph.fontAlign === fontAlign) return paragraph;
    changed = true;
    return { ...paragraph, fontAlign };
  });
  return changed ? withPersistedParagraphs(block, nextParagraphs) : null;
}

/**
 * A row's indentation as PowerPoint applies it: its own `marL` / `indent`,
 * each falling back to the one its outline level inherits from the master
 * and list styles (`source.nativeListIndents`).
 */
function effectiveParagraphIndent(
  paragraph: Pick<SlideParagraphWithRuns, "marL" | "indent" | "level">,
  inherited: readonly NativeListIndent[] | undefined,
): { marL: number | undefined; indent: number | undefined } {
  const level = inherited?.[Math.max(0, Math.min(inherited.length - 1, paragraph.level))];
  return {
    marL: paragraph.marL ?? level?.marL,
    indent: paragraph.indent ?? level?.indent,
  };
}

/**
 * Paragraph > Indentation on the selected rows: Before text is each row's own
 * `marL`, Special / By its own `indent` (null for Special: none). The row
 * then states its indentation and no longer reads its level's inherited one,
 * so the half the person did not change is written as the value the row drew
 * with, and the other rows keep inheriting. Null when nothing changes.
 */
export function textBlockWithParagraphIndent(
  block: PresentationTextBlock,
  change: PresentationParagraphIndentChange,
  selection: LiveTextSelection,
  inherited?: readonly NativeListIndent[],
): PresentationTextBlock | null {
  const paragraphs = paragraphsWithRunsFromTextBlock(block);
  if (paragraphs.length === 0) return null;
  const [first, last] = paragraphSelectionSpan(block.id, paragraphs, selection);
  let changed = false;
  const nextParagraphs = paragraphs.map((paragraph, index) => {
    if (index < first || index > last) return paragraph;
    const effective = effectiveParagraphIndent(paragraph, inherited);
    const marL = change.marL !== undefined ? change.marL : effective.marL;
    const inheritedIndent = effectiveParagraphIndent({ level: paragraph.level }, inherited).indent;
    // Special: (none) is an indent of zero. A row whose level inherits a
    // hanging indent has to state the zero, or the cascade hands it back.
    const indent =
      change.indent === null
        ? inheritedIndent
          ? 0
          : undefined
        : change.indent !== undefined
          ? change.indent
          : effective.indent;
    const next: SlideParagraphWithRuns = { ...paragraph };
    if (marL === undefined) delete next.marL;
    else next.marL = marL;
    if (indent === undefined) delete next.indent;
    else next.indent = indent;
    if (next.marL === paragraph.marL && next.indent === paragraph.indent) return paragraph;
    changed = true;
    return next;
  });
  return changed ? withPersistedParagraphs(block, nextParagraphs) : null;
}

/**
 * Line Spacing > Exactly on the selected rows (`a:lnSpc/a:spcPts`): each takes
 * the exact height in place of any multiple, its own or the box's. The box's
 * multiple first moves onto the rows it still governs, so they keep theirs.
 * Null when every selected row already has that height.
 */
export function textBlockWithExactLineSpacing(
  block: PresentationTextBlock,
  points: number,
  selection: LiveTextSelection,
): PresentationTextBlock | null {
  const exact = normalizeExactLineSpacingPoints(points);
  if (exact === undefined) return null;
  // A row already at Exactly counts as "a change" to the paragraph writer,
  // so ask it to clear the multiple only where a selected row still has one.
  const rows = paragraphsWithRunsFromTextBlock(block);
  const [firstRow, lastRow] = paragraphSelectionSpan(block.id, rows, selection);
  const multipleSelected = rows
    .slice(firstRow, lastRow + 1)
    .some((row) => paragraphFormatValue(row, "lineSpacing", block.style) !== undefined);
  const base = multipleSelected
    ? applyParagraphFormatSelection(block, selection, { lineSpacing: undefined }) ?? block
    : block;
  const paragraphs = paragraphsWithRunsFromTextBlock(base);
  if (paragraphs.length === 0) return null;
  const [first, last] = paragraphSelectionSpan(base.id, paragraphs, selection);
  let changed = base !== block;
  const nextParagraphs = paragraphs.map((paragraph, index) => {
    if (index < first || index > last) return paragraph;
    if (paragraph.lineSpacingPoints === exact && paragraph.lineSpacing === undefined) {
      return paragraph;
    }
    changed = true;
    const next: SlideParagraphWithRuns = { ...paragraph, lineSpacingPoints: exact };
    delete next.lineSpacing;
    return next;
  });
  return changed ? withPersistedParagraphs(base, nextParagraphs) : null;
}

/** Why a block's text cannot be held as paragraph rows whole, if it cannot. */
export type ParagraphModelRefusal = "paragraphTooLong" | "tooManyParagraphs";

/**
 * Why a block's text cannot be held as paragraph rows whole, or null when it
 * can. A row holds MAX_PARAGRAPH_TEXT_LENGTH characters (a text box's whole
 * text, PPT-B392) and a block MAX_PARAGRAPHS_PER_BLOCK rows (the page schema
 * and the server refuse more), while a flat block's text is one string. Rows
 * already persisted have passed those bounds.
 */
export function textBlockParagraphModelRefusal(
  block: PresentationTextBlock,
): ParagraphModelRefusal | null {
  if (normalizeParagraphsWithRuns(block.paragraphs)) return null;
  const lines = String(block.text ?? "").split("\n");
  if (lines.length > MAX_PARAGRAPHS_PER_BLOCK) return "tooManyParagraphs";
  return lines.some((line) => line.length > MAX_PARAGRAPH_TEXT_LENGTH)
    ? "paragraphTooLong"
    : null;
}

/** Whether a block's text can be held as paragraph rows whole. */
export function textBlockFitsParagraphModel(block: PresentationTextBlock): boolean {
  return textBlockParagraphModelRefusal(block) === null;
}

/**
 * Persist list structure without losing run formatting or phantom defaults.
 * A block whose text does not fit the rows is returned as it is: seeding rows
 * from it cuts each line to the row bound (and the rows to their count), and
 * writing that back deleted the tail of a line on Bullets (PPT-B392).
 */
export function withPersistedParagraphs(
  block: PresentationTextBlock,
  paragraphs: readonly SlideParagraphWithRuns[],
): PresentationTextBlock {
  if (!textBlockFitsParagraphModel(block)) return block;
  const text = textFromParagraphs(paragraphs);
  const normalized = normalizeParagraphsWithRuns(paragraphs);
  const next = { ...block, text } as PresentationTextBlock;
  if (!normalized) {
    delete next.paragraphs;
    dropNativeParagraphReceipts(next);
    delete next.bulletChar;
    delete next.numberFormat;
    const runs = flatRunsFromParagraphs(paragraphs);
    if (runs) next.runs = runs;
    else delete next.runs;
    return next;
  }
  next.paragraphs = normalized;
  delete next.runs;
  if (!normalized.some((paragraph) => paragraph.bullet === "bullet")) {
    delete next.bulletChar;
  }
  if (!normalized.some((paragraph) => paragraph.bullet === "number")) {
    delete next.numberFormat;
  }
  return next;
}

export function paragraphPersistenceEqual(
  left: PresentationTextBlock,
  right: PresentationTextBlock,
): boolean {
  const leftParagraphs = normalizeParagraphsWithRuns(left.paragraphs);
  const rightParagraphs = normalizeParagraphsWithRuns(right.paragraphs);
  const paragraphsEqual =
    leftParagraphs === undefined && rightParagraphs === undefined
      ? true
      : !!leftParagraphs &&
        !!rightParagraphs &&
        leftParagraphs.length === rightParagraphs.length &&
        leftParagraphs.every((paragraph, index) => {
          const candidate = rightParagraphs[index];
          return (
            candidate !== undefined &&
            paragraph.text === candidate.text &&
            paragraph.level === candidate.level &&
            paragraph.bullet === candidate.bullet &&
            paragraph.bulletChar === candidate.bulletChar &&
            paragraph.numberFormat === candidate.numberFormat &&
            JSON.stringify(paragraph.bulletFormat) ===
              JSON.stringify(candidate.bulletFormat) &&
            paragraph.align === candidate.align &&
            paragraph.lineSpacing === candidate.lineSpacing &&
            paragraph.lineSpacingPoints === candidate.lineSpacingPoints &&
            paragraph.spaceBefore === candidate.spaceBefore &&
            paragraph.spaceAfter === candidate.spaceAfter &&
            spacingFallbackBlocksEqual(
              paragraph.spacingFallbackBlocked,
              candidate.spacingFallbackBlocked,
            ) &&
            runCollectionsEqual(paragraph.runs, candidate.runs)
          );
        });
  return (
    left.text === right.text &&
    paragraphsEqual &&
    runCollectionsEqual(left.runs, right.runs) &&
    left.bulletChar === right.bulletChar &&
    left.numberFormat === right.numberFormat
  );
}

/** Apply a character patch through the one newline-aware selection planner. */
/**
 * Whether the bounded run list lost words the formatted runs hold: the run
 * normalizer keeps at most `MAX_RUNS_PER_BLOCK` runs and drops the text past
 * them, and the save refuses a block whose runs no longer spell its text
 * ("runs must concatenate exactly to text") for the whole deck. An edit that
 * would split a block past the cap is refused instead of truncating it.
 */
function runsDropText(
  normalized: readonly TextRun[] | undefined,
  formatted: readonly TextRun[],
): boolean {
  return normalized !== undefined && textFromRuns(normalized) !== textFromRuns(formatted);
}

/** The character formats an empty text box's own style can carry. */
export const EMPTY_BOX_CHARACTER_KEYS = [
  "fontSize",
  "fontFamily",
  "color",
  "bold",
  "italic",
  "underline",
] as const;

export function formatTextBlockRunSelection(
  textBlock: PresentationTextBlock,
  blockId: string,
  selection: LiveTextSelection,
  patch: Partial<RunStyle> | ((style: RunStyle | undefined) => Partial<RunStyle>),
): PresentationTextBlock | null {
  // An empty box keeps a character format for the text typed into it, as
  // PowerPoint keeps the end-of-paragraph run's (`a:endParaRPr`): the box's
  // own style takes it and the first typed characters inherit it. A run of
  // no text is dropped, so the format was lost (PPT-B173).
  if (presentationEditableText(textBlock) === "") {
    const style = { ...((textBlock.style ?? {}) as Record<string, unknown>) };
    const current = Object.fromEntries(
      EMPTY_BOX_CHARACTER_KEYS.filter((key) => style[key] !== undefined).map((key) => [key, style[key]]),
    ) as RunStyle;
    const resolved = (typeof patch === "function" ? patch(current) : patch) as Record<string, unknown>;
    const next = { ...style };
    for (const key of EMPTY_BOX_CHARACTER_KEYS) {
      if (!Object.hasOwn(resolved, key)) continue;
      const value = resolved[key];
      if (value === undefined || value === null || value === "") delete next[key];
      else next[key] = value;
    }
    return JSON.stringify(next) === JSON.stringify(style)
      ? null
      : ({ ...textBlock, style: next } as PresentationTextBlock);
  }
  const paragraphs = normalizeParagraphsWithRuns(textBlock.paragraphs);
  if (paragraphs && ownsParagraphRichLayout(paragraphs)) {
    const [from, to] = selectedTextRange(
      blockId,
      textFromParagraphs(paragraphs),
      selection,
    );
    const ranges = paragraphRunSelectionRanges(paragraphs, from, to);
    if (ranges.length === 0) return null;
    const rangeByParagraph = new Map(
      ranges.map((range) => [range.paragraphIndex, range]),
    );
    let changed = false;
    let overflow = false;
    const nextParagraphs = paragraphs.map((paragraph, index) => {
      const range = rangeByParagraph.get(index);
      if (!range) return paragraph;
      const formatted = applyRunFormat(
        range.runs,
        range.from,
        range.to,
        patch,
      );
      const normalized = normalizeRuns(formatted);
      if (runsDropText(normalized, formatted)) overflow = true;
      if (runCollectionsEqual(paragraph.runs, normalized)) return paragraph;
      changed = true;
      const nextParagraph: SlideParagraph & { runs?: TextRun[] } = {
        ...paragraph,
        text: textFromRuns(formatted),
      };
      if (normalized) nextParagraph.runs = normalized;
      else delete nextParagraph.runs;
      return nextParagraph;
    });
    if (!changed || overflow) return null;
    return {
      ...textBlock,
      paragraphs: nextParagraphs,
      text: textFromParagraphs(nextParagraphs),
    };
  }

  const baseRuns = normalizeRuns(textBlock.runs) ?? runsFromText(textBlock.text);
  const [from, to] = selectedTextRange(blockId, textBlock.text, selection);
  const formatted = applyRunFormat(baseRuns, from, to, patch);
  const normalized = normalizeRuns(formatted);
  if (runsDropText(normalized, formatted)) return null;
  if (runCollectionsEqual(textBlock.runs, normalized)) return null;
  const next: PresentationTextBlock = {
    ...textBlock,
    text: textFromRuns(formatted),
  };
  if (normalized) next.runs = normalized;
  else delete next.runs;
  return next;
}











export function materializeBlockFontInheritance(
  textBlock: PresentationTextBlock,
  onlyKeys?: Iterable<string>,
): PresentationTextBlock {
  const blockStyle = asRecord(textBlock.style);
  if (!blockStyle) return textBlock;
  const wanted = onlyKeys === undefined ? null : new Set(onlyKeys);
  const inheritedCandidate: Record<string, unknown> = {};
  const ownedKeys: (keyof RunStyle)[] = [];
  for (const key of FONT_BLOCK_INHERITABLE_RUN_STYLE_KEYS) {
    if (wanted && !wanted.has(key)) continue;
    const raw = blockStyle[key];
    if (
      key === "fontSize" &&
      (typeof raw !== "number" ||
        !Number.isFinite(raw) ||
        raw < 1 ||
        raw > 400)
    ) {
      continue;
    }
    if (
      key === "fontFamily" &&
      (typeof raw !== "string" || !raw.trim() || raw.length > 200)
    ) {
      continue;
    }
    const normalized = normalizeRunStyle({ [key]: raw });
    const authenticated = normalized?.[key];
    if (authenticated === undefined) continue;
    inheritedCandidate[key] = authenticated;
    ownedKeys.push(key);
  }
  const inherited = normalizeRunStyle(inheritedCandidate);
  if (!inherited || ownedKeys.length === 0) return textBlock;

  const withInheritedStyle = (
    text: string,
    runs: readonly TextRun[] | undefined,
  ): TextRun[] =>
    (normalizeRuns(runs) ?? runsFromText(text)).map((run) => {
      const style = applyStylePatch(inherited, run.style ?? {});
      return style ? { text: run.text, style } : { text: run.text };
    });

  const next = { ...textBlock } as PresentationTextBlock & {
    style?: PresentationTextStyle & Record<string, unknown>;
  };
  let nextStyle = { ...blockStyle };
  for (const key of ownedKeys) delete nextStyle[key];
  // Moved onto the runs, the keys are no box's any more, and no whole-object
  // command's either (`style.commandStatedKeys`).
  nextStyle = withCommandStatedKeys(nextStyle, { remove: ownedKeys });
  if (Object.keys(nextStyle).length > 0) {
    next.style = nextStyle as PresentationTextStyle & Record<string, unknown>;
  } else {
    delete next.style;
  }

  const paragraphs = normalizeParagraphsWithRuns(textBlock.paragraphs);
  if (paragraphs && ownsParagraphRichLayout(paragraphs)) {
    const nextParagraphs = paragraphs.map((paragraph) => {
      const runs = normalizeRuns(
        withInheritedStyle(paragraph.text, paragraph.runs),
      );
      return runs ? { ...paragraph, runs } : paragraph;
    });
    next.paragraphs = nextParagraphs;
    next.text = textFromParagraphs(nextParagraphs);
    delete next.runs;
    return next;
  }

  const runs = normalizeRuns(withInheritedStyle(textBlock.text, textBlock.runs));
  if (runs) next.runs = runs;
  else delete next.runs;
  return next;
}

/**
 * The box's own style with `patch` merged in; a false, empty or absent value
 * deletes its key, and an empty style is dropped. Null when nothing changes.
 */
export function textBlockWithStyle(
  block: PresentationTextBlock,
  patch: Partial<PresentationTextStyle>,
): PresentationTextBlock | null {
  const nextStyle = { ...(block.style || {}), ...patch };
  for (const [key, value] of Object.entries(nextStyle)) {
    if (
      value === false ||
      value === undefined ||
      value === null ||
      value === ""
    ) {
      delete nextStyle[key as keyof PresentationTextStyle];
    }
  }
  const currentStyle = block.style ?? {};
  const currentKeys = Object.keys(currentStyle);
  const nextKeys = Object.keys(nextStyle);
  if (
    currentKeys.length === nextKeys.length &&
    nextKeys.every((key) =>
      Object.is(
        (currentStyle as Record<string, unknown>)[key],
        (nextStyle as Record<string, unknown>)[key],
      ),
    )
  ) {
    return null;
  }
  const next = { ...block } as PresentationTextBlock;
  if (nextKeys.length > 0) next.style = nextStyle;
  else delete next.style;
  // Text Shadow is the box's: on or off, no run keeps a shadow of its own
  // (readings row 808), as PowerPoint's button sets every selected run.
  if ("textShadow" in patch) withoutRunTextShadows(next);
  return next;
}

/** The block's runs without a text shadow of their own (readings row 808). */
function withoutRunTextShadows(block: PresentationTextBlock): void {
  const strip = <R extends { style?: RunStyle }>(runs: readonly R[] | undefined): R[] | undefined =>
    runs?.map((run) => {
      if (!run.style || (run.style.textShadow === undefined && run.style.textShadowEffect === undefined)) return run;
      const { textShadow: _shadow, textShadowEffect: _effect, ...rest } = run.style;
      void _shadow;
      void _effect;
      return { ...run, style: Object.keys(rest).length ? rest : undefined };
    });
  const record = block as unknown as { runs?: { style?: RunStyle }[]; paragraphs?: { runs?: { style?: RunStyle }[] }[] };
  if (Array.isArray(record.runs)) record.runs = strip(record.runs);
  if (Array.isArray(record.paragraphs)) {
    record.paragraphs = record.paragraphs.map((paragraph) =>
      Array.isArray(paragraph.runs) ? { ...paragraph, runs: strip(paragraph.runs) } : paragraph,
    );
  }
}

/**
 * AUTOFIT (PowerPoint "Shrink text on overflow"): the block-level `autofit`
 * is set ONLY to `shrink`; `none` deletes it, so a block that never opted in
 * stays byte-clean. Null when the mode is already in place.
 */
export function textBlockWithAutofit(
  block: PresentationTextBlock,
  mode: AutofitMode,
): PresentationTextBlock | null {
  if (autofitModeForBlock(block) === mode) return null;
  const next = { ...block } as PresentationTextBlock;
  const nextStyle = { ...(next.style ?? {}) } as PresentationTextStyle & {
    autofit?: unknown;
  };
  delete nextStyle.autofit;
  if (Object.keys(nextStyle).length > 0) next.style = nextStyle;
  else delete next.style;
  if (mode !== "none") {
    next.autofit = mode;
  } else if (nativeAutofitResizesShape(block)) {
    // Do not Autofit over an imported `a:spAutoFit` box is stated.
    next.autofit = "none";
  } else {
    delete next.autofit;
  }
  return next;
}

/**
 * Columns > More Columns...: the count and the Spacing between columns
 * (`columnSpacing`, EMU; PPT-A042). One column states neither. Null when
 * unchanged.
 */
export function textBlockWithColumns(
  block: PresentationTextBlock,
  count: number,
  spacingEmu: number,
): PresentationTextBlock | null {
  const nextCount = persistedColumnCount(count);
  const spacing = normalizeColumnSpacing(Math.round(spacingEmu));
  const next = { ...block } as PresentationTextBlock & { columnSpacing?: number };
  delete next.columnCount;
  delete next.columnSpacing;
  if (nextCount !== undefined) {
    next.columnCount = nextCount;
    if (spacing !== undefined) next.columnSpacing = spacing;
  }
  return JSON.stringify(next) === JSON.stringify(block) ? null : next;
}

/** Columns: 1 (the default) drops `columnCount`. Null when unchanged. */
export function textBlockWithColumnCount(
  block: PresentationTextBlock,
  raw: number,
): PresentationTextBlock | null {
  const current = persistedColumnCount(block.columnCount);
  const nextCount = persistedColumnCount(raw);
  if (current === nextCount) return null;
  const next = { ...block } as PresentationTextBlock;
  delete next.columnCount;
  if (nextCount !== undefined) next.columnCount = nextCount;
  return next;
}

/**
 * A list MARKER STYLE from the bullet / numbering galleries: the selected
 * paragraphs become that list kind with the chosen glyph or number format.
 * The whole box states its choice once, on the block, where the default
 * variant is stated by omission; some rows state it per paragraph
 * (PPT-B084). Null when nothing changes.
 */
export function textBlockWithListMarker(
  block: PresentationTextBlock,
  kind: "bullet" | "number",
  variant: string | undefined,
  selection: LiveTextSelection,
): PresentationTextBlock | null {
  const normalized =
    kind === "bullet"
      ? normalizeBulletChar(variant)
      : normalizeNumberFormat(variant);
  // The package's markers become the rows' own first, so a row the
  // command does not touch keeps the marker it draws (PPT-A039/A040).
  block = materializeNativeBulletReceipt(block);
  const paragraphs = paragraphsWithRunsFromTextBlock(block);
  if (paragraphs.length === 0) return null;
  const [first, last] = paragraphSelectionSpan(
    block.id,
    paragraphs,
    selection,
  );
  const field = kind === "bullet" ? "bulletChar" : "numberFormat";
  // The gallery's plain choice (the solid disc, "1. 2. 3.") is stated
  // by omission on the block, where it keeps the per-level ladder.
  const chosen =
    normalized === undefined || normalized === "decimal"
      ? undefined
      : normalized;
  // Every row: the block's choice changes and no row keeps its own.
  // Some rows: PowerPoint states the marker per paragraph, so only the
  // selected items change and the rest keep theirs (PPT-B084).
  const whole = first === 0 && last === paragraphs.length - 1;
  const blockChoice =
    kind === "bullet"
      ? normalizeBlockBulletChar(block.bulletChar)
      : normalizeNumberFormat(block.numberFormat);
  const rowChoice =
    whole || chosen === (blockChoice === "decimal" ? undefined : blockChoice)
      ? undefined
      : (chosen ?? (kind === "bullet" ? "•" : "decimal"));
  // SET (not toggle) the kind: the gallery states a target, and re-picking
  // a glyph on an existing list must never dissolve the list.
  const listed = paragraphs.map((paragraph, index) => {
    if (index < first || index > last) return paragraph;
    const next: SlideParagraphWithRuns = { ...paragraph, bullet: kind };
    delete next[field];
    if (next.bulletFormat) {
      // A gallery pick replaces the row's own glyph / scheme; its colour,
      // size and Start at stay, as PowerPoint keeps them.
      const kept = { ...next.bulletFormat };
      delete kept.char;
      delete kept.typeface;
      delete kept.autoNum;
      next.bulletFormat = kept;
    }
    if (rowChoice === undefined) return next;
    return kind === "bullet"
      ? { ...next, bulletChar: rowChoice as BulletChar }
      : { ...next, numberFormat: rowChoice as NumberFormat };
  });
  const next = withPersistedParagraphs(block, rowsWithLayoutHiddenMarkers(block, listed));
  if (whole) {
    if (kind === "bullet") {
      delete next.bulletChar;
      if (chosen !== undefined) next.bulletChar = chosen as BulletChar;
    } else {
      delete next.numberFormat;
      if (chosen !== undefined) next.numberFormat = chosen as NumberFormat;
    }
  }
  return paragraphPersistenceEqual(next, block) ? null : next;
}

/**
 * Bullets / Numbering: the selected paragraphs become that list kind, or
 * plain paragraphs when every one of them already is it. Flat runs are split
 * across newline paragraphs first, keeping character formatting.
 */
export function textBlockWithListToggle(
  block: PresentationTextBlock,
  kind: "bullet" | "number",
  selection: LiveTextSelection,
): PresentationTextBlock | null {
  // The package's markers become the rows' own first, so a row the
  // command does not touch keeps the marker it draws (PPT-A039/A040).
  block = materializeNativeBulletReceipt(block);
  const paragraphs = paragraphsWithRunsFromTextBlock(block);
  if (paragraphs.length === 0) return null;
  const [first, last] = paragraphSelectionSpan(
    block.id,
    paragraphs,
    selection,
  );
  const allSelected = paragraphs
    .slice(first, last + 1)
    .every((paragraph) => paragraph.bullet === kind);
  const bullet: SlideParagraph["bullet"] = allSelected ? "none" : kind;
  const nextParagraphs = paragraphs.map((paragraph, index) =>
    index >= first && index <= last
      ? { ...paragraph, bullet }
      : paragraph,
  );
  const next = withPersistedParagraphs(block, rowsWithLayoutHiddenMarkers(block, nextParagraphs));
  return paragraphPersistenceEqual(next, block) ? null : next;
}

/** None (the list galleries): the selected paragraphs lose their markers. */
export function textBlockWithListCleared(
  block: PresentationTextBlock,
  selection: LiveTextSelection,
): PresentationTextBlock | null {
  // The package's markers become the rows' own first, so a row the
  // command does not touch keeps the marker it draws (PPT-A039/A040).
  block = materializeNativeBulletReceipt(block);
  const paragraphs = paragraphsWithRunsFromTextBlock(block);
  if (paragraphs.length === 0) return null;
  const [first, last] = paragraphSelectionSpan(
    block.id,
    paragraphs,
    selection,
  );
  const nextParagraphs = paragraphs.map((paragraph, index) =>
    index >= first && index <= last && paragraph.bullet !== "none"
      ? { ...paragraph, bullet: "none" as const }
      : paragraph,
  );
  const next = withPersistedParagraphs(block, nextParagraphs);
  return paragraphPersistenceEqual(next, block) ? null : next;
}

/**
 * Increase / Decrease List Level over the selected paragraphs. A flat box's
 * shared indent moves onto every paragraph in the same change, so a selected
 * line neither loses a Decrease nor doubles an Increase.
 */
export function textBlockWithIndentShift(
  block: PresentationTextBlock,
  delta: -1 | 1,
  selection: LiveTextSelection,
): PresentationTextBlock | null {
  const existingParagraphs = normalizeParagraphsWithRuns(
    block.paragraphs,
  );
  let migrationBlock = block;
  let paragraphs = existingParagraphs;
  if (!paragraphs) {
    const rawIndent = (
      block.style as { indentLevel?: unknown } | undefined
    )?.indentLevel;
    const inheritedIndent =
      rawIndent === undefined
        ? 0
        : typeof rawIndent === "number" &&
            Number.isInteger(rawIndent) &&
            rawIndent >= 0 &&
            rawIndent <= MAX_PARAGRAPH_LEVEL
          ? rawIndent
          : null;
    // A block indent outside the paragraph model's authenticated range
    // cannot be migrated without changing its render/PPTX meaning.
    if (inheritedIndent === null) return null;
    paragraphs = paragraphsWithRunsFromTextBlock(block).map(
      (paragraph) =>
        paragraph.level === inheritedIndent
          ? paragraph
          : { ...paragraph, level: inheritedIndent },
    );

    // The inherited block indent affects every flat paragraph. Once a
    // selected line changes, move that shared base onto EVERY paragraph
    // and remove the block field in the same commit, preventing either a
    // lost Decrease or a double-applied Increase.
    if (rawIndent !== undefined) {
      const nextStyle = {
        ...(block.style ?? {}),
      } as PresentationTextStyle & { indentLevel?: unknown };
      delete nextStyle.indentLevel;
      migrationBlock = { ...block };
      if (Object.keys(nextStyle).length > 0) {
        migrationBlock.style = nextStyle;
      } else {
        delete migrationBlock.style;
      }
    }
  }
  if (paragraphs.length === 0) return null;
  const [first, last] = paragraphSelectionSpan(
    block.id,
    paragraphs,
    selection,
  );
  let nextParagraphs = paragraphs;
  // An imported body's words re-inherit at the new level, as PowerPoint's
  
  const restyle = presentationListLevelRestyle(block);
  for (let index = first; index <= last; index += 1) {
    nextParagraphs = adjustParagraphLevelAt(
      nextParagraphs,
      index,
      delta,
      restyle,
    ) as SlideParagraphWithRuns[];
  }
  if (nextParagraphs === paragraphs) return null;
  return withPersistedParagraphs(migrationBlock, nextParagraphs);
}

/**
 * Paragraph alignment and spacing on a text block. A flat box with no
 * selection keeps its box-wide style (Align Left is the default and is
 * stated by omission); paragraphs take the format on the selected rows.
 */
export function textBlockWithParagraphFormat(
  block: PresentationTextBlock,
  paragraphPatch: PresentationParagraphFormatPatch,
  selection: LiveTextSelection,
): PresentationTextBlock | null {
  if (
    selection === null &&
    normalizeParagraphsWithRuns(block.paragraphs) === undefined
  ) {
    const compatibilityPatch: PresentationParagraphFormatPatch = {
      ...paragraphPatch,
    };
    if (compatibilityPatch.align === "left") {
      compatibilityPatch.align = undefined;
    }
    return textBlockWithStyle(block, compatibilityPatch);
  }
  return applyParagraphFormatSelection(block, selection, paragraphPatch);
}

// A free text box's default size (`PRESENTATION_DEFAULT_TEXT_BOX_SIZE_PT`/
// `_PX`) is stated in text-units.ts, a module with no imports of its own:
// presentation-shape-buttons.ts derives its module-level
// `NEW_TEXT_BOX_FONT_SIZE_*` from it and is evaluated inside this module's
// import cycle (this -> run-text-paint -> shape-buttons -> this), so a read
// of this module's own constant there was a ReferenceError whenever this
// module was reached first -- the deck print route, every deck's PDF

export { PRESENTATION_DEFAULT_TEXT_BOX_SIZE_PT, PRESENTATION_DEFAULT_TEXT_BOX_SIZE_PX };

/** A text box the author drew: no layout placeholder, no role. */
function isFreeTextBoxBlock(block: PresentationTextBlock): boolean {
  const record = block as unknown as Record<string, unknown>;
  return record.kind === "textbox" && typeof record.placeholder !== "string" && typeof record.role !== "string";
}










export function textBlockWithFormattingCleared(
  textBlock: PresentationTextBlock,
  selection: LiveTextSelection,
): PresentationTextBlock | null {
  const cleared = formattingCleared(textBlock, selection);
  // A paint every character still wears is the box's again; one only some
  // keep stays on their runs (`presentationBlockWithCanonicalTextPaint`).
  const next = cleared && presentationBlockWithCanonicalTextPaint(cleared);
  // Clearing words already at the default (the size goes and comes back)
  // changes nothing: no save, no Undo step.
  return next && stableJson(next) === stableJson(textBlock) ? null : next;
}

/** A value's JSON with its keys sorted, to compare two blocks' content. */
function stableJson(value: unknown): string {
  return JSON.stringify(value, (_key, item: unknown) =>
    item && typeof item === "object" && !Array.isArray(item)
      ? Object.fromEntries(Object.entries(item as Record<string, unknown>).sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)))
      : item,
  );
}

function formattingCleared(
  textBlock: PresentationTextBlock,
  selection: LiveTextSelection,
): PresentationTextBlock | null {
  const blockId = textBlock.id;
  
  if (selection && selection.blockId === blockId && selection.end <= selection.start) return null;
  // The box's own WordArt paint is every character's: it goes onto the runs
  // first, so the cleared ones lose it and the rest keep theirs (PowerPoint
  // clears the selected characters' Text Fill, Outline and Effects alone,
  // truth c64cleartx C). Kept on the box, the paint was written back onto
  
  const unpainted = presentationBlockWithRunTextPaints(textBlock);
  const materialized = materializeBlockFontInheritance(unpainted);
  const cleared =
    formatTextBlockRunSelection(
      materialized,
      blockId,
      selection,
      FONT_RUN_CLEAR_PATCH,
    ) ?? (materialized === textBlock ? null : materialized);
  // Text Shadow and Character Spacing are kept for the whole box, so clearing
  // all of its text clears them too (PPT-B071: the shadow and the pressed
  // button stayed; the spacing stayed the same way). A partial selection
  // leaves the rest of the box's shadow and spacing alone.
  const text = presentationEditableText(textBlock);
  const [from, to] = selectedTextRange(blockId, text, selection);
  const wholeText = from === 0 && to >= text.length;
  // A body read with its list levels (reading row 518): the cleared words are
  // handed back to their level -- they state nothing of their own, and draw
  
  // save wrote b="0" where PowerPoint removes `b`).
  const ladder = presentationLevelStylesOf(textBlock);
  const handed = ladder && cleared ? withRunsHandedToLevel(cleared, from, to, ladder) : cleared;
  const base = ladder ? (handed ?? textBlock) : withDefaultTextBoxSize(cleared ?? textBlock, cleared !== null);
  const style = base.style as ({ textShadow?: unknown; textShadowEffect?: unknown; letterSpacing?: unknown } & Record<string, unknown>) | undefined;
  if (
    !wholeText ||
    (style?.textShadow === undefined && style?.textShadowEffect === undefined && style?.letterSpacing === undefined)
  ) {
    return base === textBlock ? cleared : base;
  }
  const { textShadow: _shadow, textShadowEffect: _effect, letterSpacing: _spacing, ...rest } = style!;
  void _shadow;
  void _effect;
  // Spacing every row's own run defaults state (`a:pPr/a:defRPr@spc`,
  // readings row 885) is the rows', not the words': it stays, as PowerPoint
  // keeps the paragraph's defRPr through Clear All Formatting.
  const rows = normalizeParagraphsWithRuns(base.paragraphs);
  const rowsSpacing =
    _spacing !== undefined &&
    Boolean(rows?.length) &&
    rows!.every((row) => row.runDefaults?.letterSpacing === _spacing);
  return { ...base, style: (rowsSpacing ? { ...rest, letterSpacing: _spacing } : rest) as PresentationTextBlock["style"] };
}

/** The runs of `block` inside [from, to) handed back to their rows' list level (`runStyleHandedToLevel`). */
function withRunsHandedToLevel(
  block: PresentationTextBlock,
  from: number,
  to: number,
  ladder: readonly PresentationLevelStyle[],
): PresentationTextBlock {
  // Words cleared of every format are kept as plain text (no runs): they get
  // one run again to carry the level's.
  const handRuns = (
    own: readonly TextRun[] | undefined,
    text: string,
    start: number,
    level: number,
    runDefaults?: ParagraphRunDefaults,
  ): TextRun[] | undefined => {
    const runs = own ?? (text ? runsFromText(text) : undefined);
    if (!runs) return undefined;
    let cursor = start;
    return runs.map((run) => {
      const runStart = cursor;
      cursor += run.text.length;
      if (run.text.length === 0 || runStart < from || cursor > to) return run;
      // Back to the level, under the row's own run defaults (readings row 885).
      return { ...run, style: runStyleHandedToLevel(run.style, ladder, level, runDefaults) };
    });
  };
  const paragraphs = normalizeParagraphsWithRuns(block.paragraphs);
  if (paragraphs && ownsParagraphRichLayout(paragraphs)) {
    let offset = 0;
    const next = paragraphs.map((paragraph) => {
      const start = offset;
      offset += paragraph.text.length + 1;
      const runs = handRuns(paragraph.runs, paragraph.text, start, Number(paragraph.level ?? 0), paragraph.runDefaults);
      return runs ? { ...paragraph, runs } : paragraph;
    });
    return { ...block, paragraphs: next };
  }
  const level = Number(((block.style ?? {}) as { indentLevel?: unknown }).indentLevel ?? 0);
  const runs = handRuns(normalizeRuns(block.runs) ?? undefined, block.text ?? "", 0, Number.isFinite(level) ? level : 0);
  return runs ? { ...block, runs } : block;
}

/** The cleared free text box stating the default size its words now take. */
function withDefaultTextBoxSize(block: PresentationTextBlock, changed: boolean): PresentationTextBlock {
  if (!changed || !isFreeTextBoxBlock(block)) return block;
  const style = (block.style ?? {}) as Record<string, unknown>;
  if (typeof style.fontSize === "number") return block;
  const size = presentationTextUsesPoints(block as unknown as PresentationBlock) ? PRESENTATION_DEFAULT_TEXT_BOX_SIZE_PT : PRESENTATION_DEFAULT_TEXT_BOX_SIZE_PX;
  return { ...block, style: { ...style, fontSize: size } as PresentationTextBlock["style"] };
}

/**
 * CHANGE CASE over the selection's runs; list rows are mapped across their
 * newline separators. Characters outside the range and every run style stay.
 */
export function textBlockWithCase(
  textBlock: PresentationTextBlock,
  mode: ChangeCaseMode,
  selection: LiveTextSelection,
): PresentationTextBlock | null {
  const blockId = textBlock.id;
  const paragraphs = normalizeParagraphsWithRuns(textBlock.paragraphs);
  if (paragraphs && ownsParagraphRichLayout(paragraphs)) {
    const [from, to] = selectedTextRange(
      blockId,
      textFromParagraphs(paragraphs),
      selection,
    );
    const ranges = paragraphRunSelectionRanges(paragraphs, from, to);
    if (ranges.length === 0) return null;
    const rangeByParagraph = new Map(
      ranges.map((range) => [range.paragraphIndex, range]),
    );
    let changed = false;
    const nextParagraphs = paragraphs.map((paragraph, index) => {
      const range = rangeByParagraph.get(index);
      if (!range) return paragraph;
      const recased = applyCaseToRuns(
        range.runs,
        range.from,
        range.to,
        mode,
      );
      const recasedText = textFromRuns(recased);
      const normalized = normalizeRuns(recased);
      if (
        recasedText === paragraph.text &&
        runCollectionsEqual(paragraph.runs, normalized)
      ) {
        return paragraph;
      }
      changed = true;
      const nextParagraph: SlideParagraph & { runs?: TextRun[] } = {
        ...paragraph,
        text: recasedText,
      };
      if (normalized) nextParagraph.runs = normalized;
      else delete nextParagraph.runs;
      return nextParagraph;
    });
    if (!changed) return null;
    return {
      ...textBlock,
      paragraphs: nextParagraphs,
      text: textFromParagraphs(nextParagraphs),
    };
  }

  const total = textBlock.text.length;
  const [from, to] = selectedTextRange(blockId, textBlock.text, selection);
  const hasSubRange = from > 0 || to < total;
  if (textBlock.runs?.length || hasSubRange) {
    const baseRuns =
      normalizeRuns(textBlock.runs) ?? runsFromText(textBlock.text);
    const recased = applyCaseToRuns(baseRuns, from, to, mode);
    const recasedText = textFromRuns(recased);
    const normalized = normalizeRuns(recased);
    if (
      recasedText === textBlock.text &&
      runCollectionsEqual(textBlock.runs, normalized)
    ) {
      return null;
    }
    const next: PresentationTextBlock = {
      ...textBlock,
      text: recasedText,
    };
    if (normalized) next.runs = normalized;
    else delete (next as { runs?: unknown }).runs;
    return next;
  }

  const recasedText = applyCaseToText(textBlock.text, mode);
  return recasedText === textBlock.text
    ? null
    : { ...textBlock, text: recasedText };
}

/**
 * Bold / Italic / Underline / Strikethrough / Superscript / Subscript as the
 * ribbon toggles them over the selection: off when every selected run has
 * it, else on. A flag the whole box states is first moved onto its runs,
 * since the runs alone could never turn it off (PPT-B096).
 */
export function textBlockWithRunToggle(
  block: PresentationTextBlock,
  key: ToggleableRunStyleKey,
  selection: LiveTextSelection,
): PresentationTextBlock | null {
  // Bold, Italic and Underline over an equation change its letters (math
  // italic / upright, bold, underlined), as Office Math does:
  // `lib/presentation-equation-run-italic`.
  if (key === "italic" || key === "bold" || key === "underline") {
    return withEquationFormatToggle(block, key, (materialized) =>
      textBlockWithRunToggleOfRuns(materialized, key, selection),
    );
  }
  return textBlockWithRunToggleOfRuns(block, key, selection);
}

function textBlockWithRunToggleOfRuns(
  block: PresentationTextBlock,
  key: ToggleableRunStyleKey,
  selection: LiveTextSelection,
): PresentationTextBlock | null {
  const blockId = block.id;
  if (
    (asRecord(block.style) as Record<string, unknown> | null)?.[key] === true
  ) {
    const materialized = materializeBlockFontInheritance(block, [key]);
    const text = presentationEditableText(materialized);
    const [from, to] = selectedTextRange(blockId, text, selection);
    const paragraphs = normalizeParagraphsWithRuns(materialized.paragraphs);
    const uniformlyOn =
      paragraphs && ownsParagraphRichLayout(paragraphs)
        ? paragraphRunSelectionRanges(paragraphs, from, to).every(
            (range) =>
              runFieldAtSelection(range.runs, range.from, range.to, key) ===
              true,
          )
        : runFieldAtSelection(
            normalizeRuns(materialized.runs) ?? runsFromText(materialized.text),
            from,
            to,
            key,
          ) === true;
    return (
      formatTextBlockRunSelection(materialized, blockId, selection, {
        [key]: !uniformlyOn,
      }) ?? materialized
    );
  }
  const paragraphs = normalizeParagraphsWithRuns(block.paragraphs);
  let patch: Partial<RunStyle>;
  if (paragraphs && ownsParagraphRichLayout(paragraphs)) {
    const [from, to] = selectedTextRange(
      blockId,
      textFromParagraphs(paragraphs),
      selection,
    );
    const ranges = paragraphRunSelectionRanges(paragraphs, from, to);
    if (ranges.length === 0) return null;
    const uniformlyOn = ranges.every(
      (range) =>
        runFieldAtSelection(range.runs, range.from, range.to, key) === true,
    );
    patch = { [key]: !uniformlyOn };
  } else {
    const baseRuns =
      normalizeRuns(block.runs) ?? runsFromText(block.text);
    const [from, to] = selectedTextRange(blockId, block.text, selection);
    patch = resolveBooleanRunToggle(baseRuns, from, to, key);
  }
  return formatTextBlockRunSelection(block, blockId, selection, patch);
}

/**
 * Whether one character flag is on over the selection (every character
 * has it, a flag the whole box states included), off (none has it), or
 * mixed: what a toggle button shows pressed, unpressed or neither.
 */
export function textBlockRunFlagState(
  block: PresentationTextBlock,
  key: ToggleableRunStyleKey,
  selection: LiveTextSelection,
): boolean | "mixed" {
  // An equation's Bold, Italic and Underline are its letters': the button
  // reads them (its letters are italic unless made upright).
  if ((key === "italic" || key === "bold" || key === "underline") && blockHoldsEquationRun(block)) {
    block = withEquationFormatMaterialized(block, key).block;
  }
  const source =
    (asRecord(block.style) as Record<string, unknown> | null)?.[key] === true
      ? materializeBlockFontInheritance(block, [key])
      : block;
  const [from, to] = selectedTextRange(block.id, presentationEditableText(source), selection);
  const paragraphs = normalizeParagraphsWithRuns(source.paragraphs);
  const runs =
    paragraphs && ownsParagraphRichLayout(paragraphs)
      ? paragraphRunSelectionRanges(paragraphs, from, to).flatMap((range) =>
          range.runs.length ? [{ runs: range.runs, from: range.from, to: range.to }] : [],
        )
      : [{ runs: normalizeRuns(source.runs) ?? runsFromText(source.text), from, to }];
  let on = 0;
  let off = 0;
  for (const range of runs) {
    let cursor = 0;
    for (const run of range.runs) {
      const runStart = cursor;
      cursor += run.text.length;
      if (cursor <= range.from || runStart >= range.to) continue;
      if (run.style?.[key] === true) on += 1;
      else off += 1;
    }
  }
  return off === 0 && on > 0 ? true : on === 0 ? false : "mixed";
}

/**
 * Grow / Shrink Font over a selection: every run steps from its own size
 * (PPT-B072: 12 and 48 both became 18), and a run with no size of its own
 * steps from the size it inherits. A run that cannot step keeps its size.
 */
export function steppedRunFontSizePatch(
  inheritedSize: number | undefined,
  direction: 1 | -1,
  /** The block keeps points (`presentationTextUsesPoints`): its sizes were
   * stepped as px, so 18 pt grew to 18.67. */
  sizesInPoints = false,
): (style: RunStyle | undefined) => Partial<RunStyle> {
  return (style) => {
    const current = typeof style?.fontSize === "number" ? style.fontSize : inheritedSize;
    const next = stepFontSize(
      sizesInPoints && current !== undefined ? current * SLIDE_PX_PER_POINT : current,
      direction,
    );
    const px = next ? parsePointRunFontSizeInput(next) : undefined;
    if (px === undefined) return {};
    return { fontSize: sizesInPoints ? Math.round((px / SLIDE_PX_PER_POINT) * 100) / 100 : px };
  };
}

/** Rendering units do not confer native package mutation authority. */
export { presentationTextUsesPoints };

export function emptyPlaceholderFontSize(
  block: PresentationTextBlock | null,
  layout: unknown,
): number | undefined {
  if (!block || block.text !== "") return undefined;
  const role = (block as Record<string, unknown>).role;
  if (role === "title") return normalizeLayoutId(layout) === "title" ? 60 : 44;
  if (role === "subtitle") return 32;
  if (role === "content" || role === "body" || role === "picture") return 28;
  return undefined;
}

/**
 * The size a placeholder's words draw at when nothing states one, in design
 * px: the stage's placeholder type (`presentation-window-shell.module.css`,
 * at the slide's design width) -- a title-slide title 56 px (42 pt), its
 * subtitle 28 px (21 pt), any other title 42 px, a content placeholder 30 px.
 * Other text draws at the slide's 17 px; undefined here.
 */
export function placeholderStageFontSize(
  block: PresentationTextBlock | null,
  layout: unknown,
): number | undefined {
  if (!block) return undefined;
  const role = (block as Record<string, unknown>).role;
  const titleSlide = normalizeLayoutId(layout) === "title";
  if (role === "title") return titleSlide ? 56 : 42;
  if (role === "subtitle") return titleSlide ? 28 : undefined;
  if (role === "content") return 30;
  return undefined;
}

/**
 * One Home command (Grow / Shrink, a toggle, Underline, Align, a size, a
 * face, a colour) over every selected object's words: text boxes, shapes,
 * SmartArt labels and charts' own text. The slide's next blocks, or null.
 */
export function homeObjectBlocksWithCommand(
  blocks: PresentationBlock[],
  targets: PresentationHomeObjectTargets,
  command: OfficeHomeTextCommand,
  fallbackSizePt: (block: PresentationTextBlock | PresentationShapeBlock) => number | undefined,
): PresentationBlock[] | null {
  const textIds = new Set(targets.text.map((block) => block.id));
  const chartIds = new Set(targets.charts.map((block) => block.id));
  const liveText = blocks.filter((block) => textIds.has(block.id)) as typeof targets.text;
  const written = liveText.length
    ? applyPresentationHomeObjectText(liveText, command, {
        fallbackSizePt,
        // The deck's own size presets, as the size box and the text
        // path step them.
        stepSizePx: (px, direction) => {
          const next = stepFontSize(px, direction);
          // To the hundredth: whole px stepped 24 pt to 27.75, not 28.
          return next ? parsePointRunFontSizeInput(next) : undefined;
        },
      })
    : null;
  const byId = new Map((written ?? []).map((block) => [block.id, block]));
  let changed = written !== null;
  const next = blocks.map((block) => {
    if (byId.has(block.id)) return byId.get(block.id) as PresentationBlock;
    if (!chartIds.has(block.id) || block.type !== "chart") return block;
    const chart = applyPresentationChartHomeText(block as PresentationChartBlock, command, targets.chartElement, targets.chartTitleRange);
    if (!chart || chart === block) return block;
    changed = true;
    return chart as PresentationBlock;
  });
  return changed ? next : null;
}

/** The selected text targets of `targets`, as they are on the live slide. */
function liveHomeObjectText(
  blocks: PresentationBlock[],
  targets: PresentationHomeObjectTargets,
): (PresentationTextBlock | PresentationShapeBlock)[] {
  const ids = new Set(targets.text.map((block) => block.id));
  return blocks.filter((block) => ids.has(block.id) && (block.type === "text" || block.type === "shape")) as (
    | PresentationTextBlock
    | PresentationShapeBlock
  )[];
}

function withBlocksReplaced(
  blocks: PresentationBlock[],
  written: readonly PresentationBlock[] | null,
): PresentationBlock[] | null {
  if (!written) return null;
  const byId = new Map(written.map((block) => [block.id, block]));
  return blocks.map((block) => byId.get(block.id) ?? block);
}

/** A script or a highlight on every selected object's words. The slide's next blocks, or null. */
export function homeObjectBlocksWithRunMark(
  blocks: PresentationBlock[],
  targets: PresentationHomeObjectTargets,
  mark: PresentationHomeObjectRunMark,
): PresentationBlock[] | null {
  return withBlocksReplaced(
    blocks,
    applyPresentationHomeObjectRunMark(liveHomeObjectText(blocks, targets), mark) as PresentationBlock[] | null,
  );
}

/** Align Text: the text anchored inside every selected object. The slide's next blocks, or null. */
export function homeObjectBlocksWithAnchor(
  blocks: PresentationBlock[],
  targets: PresentationHomeObjectTargets,
  anchor: "top" | "middle" | "bottom",
): PresentationBlock[] | null {
  return withBlocksReplaced(
    blocks,
    applyPresentationHomeObjectAnchor(liveHomeObjectText(blocks, targets), anchor) as PresentationBlock[] | null,
  );
}

/** An explicit format (the size box, a colour, a face) on every selected object's words. */
export function homeObjectBlocksWithPatch(
  blocks: PresentationBlock[],
  targets: PresentationHomeObjectTargets,
  patch: SmartArtNodeTextFormatPatch,
): PresentationBlock[] | null {
  const textIds = new Set(targets.text.map((block) => block.id));
  const chartIds = new Set(targets.charts.map((block) => block.id));
  let changed = false;
  const next = blocks.map((block) => {
    if (textIds.has(block.id) && (block.type === "text" || block.type === "shape")) {
      const written = PRESENTATION_SMARTART_LABEL_TEXT_ADAPTER.write(block as PresentationTextBlock, patch);
      if (written === block) return block;
      changed = true;
      return written as PresentationBlock;
    }
    if (chartIds.has(block.id) && block.type === "chart") {
      let chart: PresentationChartBlock | null = block as PresentationChartBlock;
      for (const command of [
        ...(Object.hasOwn(patch, "fontFamily") ? [{ kind: "family", family: patch.fontFamily ?? null } as const] : []),
        ...(patch.fontSizePt !== undefined ? [{ kind: "size", pt: patch.fontSizePt } as const] : []),
        ...(Object.hasOwn(patch, "color") ? [{ kind: "color", color: patch.color ?? null } as const] : []),
      ]) {
        chart = chart ? applyPresentationChartHomeText(chart, command, targets.chartElement, targets.chartTitleRange) ?? chart : chart;
      }
      if (!chart || chart === block) return block;
      changed = true;
      return chart as PresentationBlock;
    }
    return block;
  });
  return changed ? next : null;
}

/**
 * A character patch of the text path (sizes in the deck's px) as the
 * selected-object path's patch (sizes in points): what every selected
 * object's words take. Highlight and baseline shifts are run marks the
 * object path does not carry.
 */
export function homeObjectPatchFromRunPatch(patch: Partial<RunStyle>): SmartArtNodeTextFormatPatch {
  return {
    ...(Object.hasOwn(patch, "fontFamily") ? { fontFamily: patch.fontFamily || undefined } : {}),
    // Theme Fonts on the whole object: the face follows the theme (reading row 438).
    ...(patch.fontFamily && (patch.themeFontRole === "heading" || patch.themeFontRole === "body")
      ? { themeFontRole: patch.themeFontRole }
      : {}),
    ...(typeof patch.fontSize === "number" ? { fontSizePt: Math.round((patch.fontSize * 3) / 4 * 100) / 100 } : {}),
    ...(Object.hasOwn(patch, "bold") ? { bold: patch.bold === true } : {}),
    ...(Object.hasOwn(patch, "italic") ? { italic: patch.italic === true } : {}),
    ...(Object.hasOwn(patch, "underline") ? { underline: patch.underline ? ("single" as const) : undefined } : {}),
    ...(Object.hasOwn(patch, "strike") ? { strike: patch.strike === true } : {}),
    ...(Object.hasOwn(patch, "color") ? { color: patch.color || undefined } : {}),
  };
}

/**
 * A PowerPoint Bullets / Numbering library item, or the Bullets and Numbering
 * dialog's OK, on the selected rows (PPT-A039/A040). The block's one gallery
 * override never speaks for a selected row afterwards: where the whole box
 * is selected it goes, otherwise a selected row on the plain disc / "1." says
 * so itself.
 */
export function textBlockWithListRows(
  block: PresentationTextBlock,
  selection: LiveTextSelection,
  transform: (
    rows: readonly SlideParagraphWithRuns[],
    span: readonly [number, number],
  ) => SlideParagraphWithRuns[],
): PresentationTextBlock | null {
  const materialized = materializeNativeBulletReceipt(block);
  const paragraphs = paragraphsWithRunsFromTextBlock(materialized);
  if (paragraphs.length === 0) return null;
  const span = paragraphSelectionSpan(materialized.id, paragraphs, selection);
  const whole = span[0] === 0 && span[1] === paragraphs.length - 1;
  const rows = transform(paragraphs, span).map((row, index) => {
    if (whole || index < span[0] || index > span[1]) return row;
    if (row.bullet === "bullet" && materialized.bulletChar && row.bulletChar === undefined) {
      return { ...row, bulletChar: "•" as BulletChar };
    }
    if (row.bullet === "number" && materialized.numberFormat && row.numberFormat === undefined) {
      return { ...row, numberFormat: "decimal" as NumberFormat };
    }
    return row;
  });
  const next = withPersistedParagraphs(materialized, rowsWithLayoutHiddenMarkers(materialized, rows));
  if (whole) {
    delete next.bulletChar;
    delete next.numberFormat;
  }
  return paragraphPersistenceEqual(next, materialized) ? null : next;
}
