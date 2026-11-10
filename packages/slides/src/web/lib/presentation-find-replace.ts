/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isLayoutGraphicOnSlide } from "@mrnote/core/web/lib/presentation-master.ts";
import { mathTextReplaced } from "@mrnote/core/web/lib/office-math-text-replace.ts";
import { runCarriesMath, runIsAtomic, type TextRun } from "@mrnote/core/web/lib/presentation-runs.ts";
import { flatRunsFromParagraphs } from "@mrnote/core/web/lib/presentation-text-buttons.ts";
import { isTextEditLocked } from "@mrnote/core/web/lib/presentation-text-case.ts";
export * from "./presentation-run-ranges";
import { replaceRangeInRuns } from "./presentation-run-ranges";
import type { RunLike } from "./presentation-run-ranges";

/**
 * Cross-slide Find & Replace for the presentation editor (PowerPoint / WPS
 * Ctrl+H parity).
 *
 * Searches EVERY slide's user-visible text — the slide title, each text block's
 * flat `text` (and its per-run `runs` / list `paragraphs`), table cell text, and
 * a chart's title/caption — and replaces one or all matches. The replacement is
 * applied so the per-run invariant (`textFromRuns(runs) === text`) is preserved.
 * Matching runs over the CONTINUOUS text of a paragraph (or a run-formatted
 * block), so a phrase whose characters carry different formatting is found and
 * replaced as one: PowerPoint's Find and Replace works on text, not on the
 * formatting spans under it. The replacement takes the formatting of the first
 * character it replaces (PowerPoint's `TextRange.Replace`), every character
 * outside the match keeps its own run, and an unrelated bold phrase keeps its
 * style after a replace.
 *
 * Keep this module dependency-free (no React / no DOM) so it can be unit tested
 * in isolation and shared by the editor's Find & Replace dialog.
 */

/** Search options the dialog exposes. */
export interface FindOptions {
  /** Case-sensitive match (default: case-insensitive). */
  matchCase?: boolean;
  /** Match only complete Unicode words, never a substring inside a word. */
  wholeWord?: boolean;
  /**
   * Replace only: whether this Chart V2 chart may have its drawn title
   * rewritten (the host's chart-authority rule). Absent means yes. A chart
   * that refuses keeps its title and is not counted as replaced.
   */
  canReplaceChartTitle?: (block: Record<string, unknown>) => boolean;
}

/** PowerPoint Web searches the open presentation; it has no invented slide-only scope. */
export const PRESENTATION_FIND_SCOPE = "presentation" as const;

interface MatchRange {
  start: number;
  end: number;
}

const PRESENTATION_WORD_CHARACTER = /[\p{L}\p{M}\p{N}_]/u;

function hasWordCharacterBefore(text: string, index: number): boolean {
  if (index <= 0) return false;
  const previous = Array.from(text.slice(0, index)).at(-1) ?? "";
  return PRESENTATION_WORD_CHARACTER.test(previous);
}

function hasWordCharacterAfter(text: string, index: number): boolean {
  if (index >= text.length) return false;
  const next = Array.from(text.slice(index))[0] ?? "";
  return PRESENTATION_WORD_CHARACTER.test(next);
}

function isWholeWordRange(text: string, start: number, end: number): boolean {
  return (
    !hasWordCharacterBefore(text, start) &&
    !hasWordCharacterAfter(text, end)
  );
}

/**
 * `text` lower-cased character by character, with every folded position
 * mapped back to the ORIGINAL text. Lower-casing can change a character's
 * length (U+0130 "İ" lower-cases to "i" + U+0307), so an index found in the
 * folded string is never an index into the text (PPT-B045): `origin[k]` is
 * the original offset of the character whose folding holds position `k`,
 * and `origin[folded.length]` is the text's length. `boundary[k]` marks the
 * folded positions where a character's folding starts.
 */
function foldedText(text: string): {
  folded: string;
  origin: number[];
  boundary: boolean[];
} {
  let folded = "";
  const origin: number[] = [];
  const boundary: boolean[] = [];
  for (let index = 0; index < text.length; ) {
    const point = text.codePointAt(index) ?? 0;
    const character = String.fromCodePoint(point);
    // Up then down folds a final sigma (ς) to σ with its capital, so a word
    // and its upper-case spelling fold alike (PPT-B254).
    const lower = character.toUpperCase().toLowerCase();
    for (let unit = 0; unit < lower.length; unit += 1) {
      origin.push(index);
      boundary.push(unit === 0);
    }
    folded += lower;
    index += character.length;
  }
  origin.push(text.length);
  boundary.push(true);
  return { folded, origin, boundary };
}

function literalMatchRanges(
  text: string,
  query: string,
  options?: FindOptions,
): MatchRange[] {
  if (!query) return [];
  const ranges: MatchRange[] = [];
  const accept = (start: number, end: number) => {
    if (!options?.wholeWord || isWholeWordRange(text, start, end)) {
      ranges.push({ start, end });
    }
  };
  if (options?.matchCase) {
    for (let from = 0; ; ) {
      const start = text.indexOf(query, from);
      if (start < 0) break;
      accept(start, start + query.length);
      from = start + query.length;
    }
    return ranges;
  }
  // Case-insensitive: search the folded text, then map each hit back to the
  // original characters it covers. A hit that starts or ends inside one
  // character's folding (half of "İ") is not a whole-character match.
  const { folded, origin, boundary } = foldedText(text);
  // The query folds the way the text folds, character by character: whole-
  // string lower-casing gives a final sigma its word-final form (ς) while the
  // text's per-character fold gives σ, so ΟΣ never found ΟΣ (PPT-B254).
  const needle = foldedText(query).folded;
  for (let from = 0; ; ) {
    const start = folded.indexOf(needle, from);
    if (start < 0) break;
    const end = start + needle.length;
    if (boundary[start] && boundary[end]) {
      accept(origin[start], origin[end]);
      from = end;
    } else {
      from = start + 1;
    }
  }
  return ranges;
}

function replaceRanges(
  text: string,
  ranges: readonly MatchRange[],
  replacement: string,
): string {
  if (ranges.length === 0) return text;
  const pieces: string[] = [];
  let from = 0;
  for (const range of ranges) {
    pieces.push(text.slice(from, range.start), replacement);
    from = range.end;
  }
  pieces.push(text.slice(from));
  return pieces.join("");
}

/** Escape a string so it is treated literally inside a `RegExp`. */
export function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Count the non-overlapping occurrences of `query` in `text`. A blank `query`
 * matches nothing (so an empty search box never reports spurious hits). Pure.
 */
export function countMatches(
  text: string,
  query: string,
  options?: FindOptions,
): number {
  return literalMatchRanges(text, query, options).length;
}

/**
 * Replace EVERY non-overlapping occurrence of `query` with `replacement` in
 * `text`, returning the new string. Case-insensitive by default but preserves
 * the surrounding text exactly (only the matched spans are substituted). A blank
 * `query` is a no-op (returns `text` unchanged). Pure.
 */
export function replaceAllInString(
  text: string,
  query: string,
  replacement: string,
  options?: FindOptions,
): string {
  return replaceRanges(
    text,
    literalMatchRanges(text, query, options),
    replacement,
  );
}

function runsText(runs: readonly RunLike[]): string {
  return runs.map((run) => (typeof run.text === "string" ? run.text : "")).join("");
}








function rangeTouchesAtomicRun(runs: readonly RunLike[], start: number, end: number): boolean {
  let offset = 0;
  for (const run of runs) {
    const length = typeof run.text === "string" ? run.text.length : 0;
    if (offset < end && offset + length > start && runIsAtomic(run as unknown as TextRun)) return true;
    offset += length;
    if (offset >= end) break;
  }
  return false;
}

/**
 * `runs` with `[start, end)` replaced inside the one equation run holding it,
 * the equation's LaTeX changed to match (`mathTextReplaced`); null when the
 * range is not inside one equation or its LaTeX cannot follow.
 */
function equationRangeReplaced<T extends RunLike>(
  runs: readonly T[],
  start: number,
  end: number,
  replacement: string,
): T[] | null {
  let offset = 0;
  for (let index = 0; index < runs.length; index += 1) {
    const run = runs[index] as T & { math?: { latex?: unknown } };
    const length = typeof run.text === "string" ? run.text.length : 0;
    const runStart = offset;
    offset += length;
    if (offset <= start) continue;
    const latex = run.math?.latex;
    if (runStart > start || end > offset || typeof latex !== "string" || !runCarriesMath(run as unknown as TextRun)) return null;
    const edited = mathTextReplaced(String(run.text), latex, start - runStart, end - runStart, replacement);
    if (!edited) return null;
    const next = [...runs];
    next[index] = { ...run, text: edited.text, math: { ...run.math, latex: edited.latex } } as T;
    return next;
  }
  return null;
}

/** Matches in `runs` that reach into an equation or a field and Replace
 * leaves as they are: a field, or an equation whose LaTeX cannot follow. */
function atomicMatchesInRuns(
  runs: readonly RunLike[],
  query: string,
  options?: FindOptions,
  replacement?: string,
): number {
  if (!runs.some((run) => runIsAtomic(run as unknown as TextRun))) return 0;
  return literalMatchRanges(runsText(runs), query, options).filter(
    (range) =>
      rangeTouchesAtomicRun(runs, range.start, range.end) &&
      (replacement === undefined || equationRangeReplaced(runs, range.start, range.end, replacement) === null),
  ).length;
}

function replaceRangesInRuns<T extends RunLike>(
  runs: readonly T[],
  ranges: readonly MatchRange[],
  replacement: string,
): T[] {
  // Last range first, so the offsets of the earlier ones still hold.
  let next: T[] = [...runs];
  for (let index = ranges.length - 1; index >= 0; index -= 1) {
    if (rangeTouchesAtomicRun(runs, ranges[index].start, ranges[index].end)) {
      // Inside one equation: the same change in its LaTeX, else left whole.
      next = equationRangeReplaced(next, ranges[index].start, ranges[index].end, replacement) ?? next;
      continue;
    }
    next = replaceRangeInRuns(next, ranges[index].start, ranges[index].end, replacement);
  }
  return next;
}

/**
 * Replace every match inside a runs array. Matching reads the runs' continuous
 * text, so a match may span formatting runs; see {@link replaceRangeInRuns} for
 * where the replacement's formatting comes from. Returns the new runs
 * (untouched runs by reference; an equal array when nothing matched). Pure.
 */
export function replaceInRuns<T extends RunLike>(
  runs: readonly T[],
  query: string,
  replacement: string,
  options?: FindOptions,
): T[] {
  return replaceRangesInRuns(
    runs,
    literalMatchRanges(runsText(runs), query, options),
    replacement,
  );
}

/** Count matches across a runs array, reading its continuous text. */
export function countInRuns(
  runs: readonly RunLike[],
  query: string,
  options?: FindOptions,
): number {
  return literalMatchRanges(runsText(runs), query, options).length;
}

/* ------------------------------------------------------------------------- *
 * Deck-level traversal
 *
 * These helpers operate on the loose deck shape the editor holds (`SlideModel[]`)
 * without importing the editor's concrete types — every field is read
 * defensively so a malformed/partial slide can never throw. The result is a
 * NEW slides array (structurally cloned only where a replacement landed); slides
 * with no match are returned by reference so unrelated history/diff stays minimal.
 * ------------------------------------------------------------------------- */

/** The minimal slide/block shape this module reads. */
type AnyRecord = Record<string, unknown>;

function asArray(value: unknown): unknown[] {
  return Array.isArray(value) ? value : [];
}

function asString(value: unknown): string {
  return typeof value === "string" ? value : "";
}

/**
 * The title a chart DRAWS: a Chart V2 chart draws `chartV2.title`, which the
 * outer `title` only mirrors; a replace rewrites both, as the chart's own
 * title editing does (PPT-B031).
 */
function chartTitleSource(block: AnyRecord): { text: string; v2: AnyRecord | null } {
  const v2 =
    block.chartV2 && typeof block.chartV2 === "object" && !Array.isArray(block.chartV2)
      ? (block.chartV2 as AnyRecord)
      : null;
  return v2 && typeof v2.title === "string"
    ? { text: v2.title, v2 }
    : { text: asString(block.title), v2: null };
}

function withChartTitle(block: AnyRecord, title: string, v2: AnyRecord | null): AnyRecord {
  return v2
    ? { ...block, title, chartV2: { ...v2, title } }
    : { ...block, title };
}

function chartTitleReplaceable(block: AnyRecord, options?: FindOptions): boolean {
  return (
    !lockedAgainstReplace(block) &&
    (chartTitleSource(block).v2 === null ||
      options?.canReplaceChartTitle?.(block) !== false)
  );
}

/**
 * Lock fixes an object's place, never its words (PowerPoint's Lock: the
 * position is fixed, the content can still be changed), so Replace reaches
 * a locked text box, shape, table or chart alike. Only the file's own
 * `noTextEdit` (Designer's design elements) keeps its words.
 */
function lockedAgainstReplace(block: AnyRecord): boolean {
  return isTextEditLocked(block);
}

/** Matches in `block` that Replace may not change (locked or refused). */
function unreplaceableMatches(
  block: AnyRecord,
  query: string,
  options?: FindOptions,
  replacement?: string,
): number {
  if (block.type === "table" && lockedAgainstReplace(block)) {
    return countMatchesInBlock(block, query, options);
  }
  if (block.type === "table") {
    let atomic = 0;
    for (const row of asArray(block.rows)) {
      for (const cell of asArray(row)) atomic += atomicMatchesInRuns(cellRuns((cell ?? {}) as AnyRecord), query, options, replacement);
    }
    return atomic;
  }
  if (block.type === "text" || block.type === "shape") {
    const paragraphs = asArray(block.paragraphs);
    return (paragraphs.length > 0 ? paragraphs.map((p) => asArray((p as AnyRecord | null)?.runs)) : [asArray(block.runs)])
      .reduce<number>((total, runs) => total + atomicMatchesInRuns(runs as RunLike[], query, options, replacement), 0);
  }
  if (block.type === "chart") {
    return (
      (chartTitleReplaceable(block, options)
        ? 0
        : countMatches(chartTitleSource(block).text, query, options)) +
      (lockedAgainstReplace(block)
        ? countMatches(asString(block.caption), query, options)
        : 0)
    );
  }
  return 0;
}

function paragraphText(paragraph: AnyRecord): string {
  const runs = asArray(paragraph.runs) as RunLike[];
  return runs.length > 0
    ? runs.map((run) => asString(run.text)).join("")
    : asString(paragraph.text);
}

function textFromParagraphRecords(paragraphs: readonly unknown[]): string {
  return paragraphs
    .map((paragraph) => paragraphText((paragraph ?? {}) as AnyRecord))
    .join("\n");
}

/** A text block's words as one string: its rows, else its runs, else its text. */
function textBlockPlainText(block: AnyRecord): string {
  const paragraphs = asArray(block.paragraphs);
  if (paragraphs.length > 0) return textFromParagraphRecords(paragraphs);
  const runs = asArray(block.runs) as RunLike[];
  return runs.length > 0 ? runsText(runs) : asString(block.text);
}












function mirroredTitleBlockIndex(slide: AnyRecord): number {
  const title = asString(slide.title);
  if (!title) return -1;
  return asArray(slide.blocks).findIndex((value) => {
    const block = (value ?? {}) as AnyRecord;
    const role = asString(block.role) || asString(block.placeholder);
    return (
      (block.type === "text" || block.type === "shape") &&
      (role === "title" || role === "ctrTitle") &&
      textBlockPlainText(block) === title
    );
  });
}

/**
 * Count total matches of `query` across every searchable text field of one
 * block. Covers a text block's flat `text` (preferring its `runs` when present
 * so a styled block counts the same), its list `paragraphs` (each paragraph's
 * `text` or per-paragraph `runs`), table cell `text`, and a chart's
 * `title` / `caption`. Pure.
 */
export function countMatchesInBlock(
  block: AnyRecord,
  query: string,
  options?: FindOptions,
): number {
  if (!query) return 0;
  let total = 0;
  const type = block.type;
  if (type === "text" || (type === "shape" && asArray(block.paragraphs).length > 0)) {
    // A shape that holds paragraphs (a list box changed into a shape) is
    // drawn from them, so it is searched through them (PPT-B312).
    const paragraphs = asArray(block.paragraphs);
    if (paragraphs.length > 0) {
      for (const p of paragraphs) {
        const para = (p ?? {}) as AnyRecord;
        const runs = asArray(para.runs) as RunLike[];
        total +=
          runs.length > 0
            ? countInRuns(runs, query, options)
            : countMatches(asString(para.text), query, options);
      }
    } else {
      const runs = asArray(block.runs) as RunLike[];
      total +=
        runs.length > 0
          ? countInRuns(runs, query, options)
          : countMatches(asString(block.text), query, options);
    }
  } else if (type === "shape") {
    const runs = asArray(block.runs) as RunLike[];
    total +=
      runs.length > 0
        ? countInRuns(runs, query, options)
        : countMatches(asString(block.text), query, options);
  } else if (type === "table") {
    for (const row of asArray(block.rows)) {
      for (const cell of asArray(row)) {
        total += countMatches(asString((cell as AnyRecord)?.text), query, options);
      }
    }
  } else if (type === "chart") {
    total += countMatches(chartTitleSource(block).text, query, options);
    total += countMatches(asString(block.caption), query, options);
  }
  return total;
}

/** Count total matches across an entire deck (every slide title + block). */
export function countMatchesInDeck(
  slides: readonly AnyRecord[],
  query: string,
  options?: FindOptions,
): number {
  if (!query) return 0;
  let total = 0;
  for (const slide of slides) {
    if (mirroredTitleBlockIndex(slide) < 0) {
      total += countMatches(asString(slide.title), query, options);
    }
    for (const block of asArray(slide.blocks)) {
      // A layout's graphic drawn on the slide is the layout's: PowerPoint's
      // Find does not search it (Slide Master view does).
      if (isLayoutGraphicOnSlide(block)) continue;
      total += countMatchesInBlock(block as AnyRecord, query, options);
    }
  }
  return total;
}

export type PresentationFindField =
  | "slide-title"
  | "text"
  | "shape-text"
  | "table-cell"
  | "chart-title"
  | "chart-caption";

/**
 * One exact, replaceable occurrence in deck traversal order. The structural
 * coordinates are deliberately explicit: Replace-current reauthenticates the
 * live slide/block/path and matched substring before it mutates anything.
 * `textStart`/`textEnd` are flat text-block offsets used only to reveal the
 * occurrence in the canvas textarea.
 */
export interface PresentationFindMatch {
  readonly key: string;
  readonly slideId: string;
  readonly slideIndex: number;
  readonly blockId?: string;
  readonly blockIndex?: number;
  readonly field: PresentationFindField;
  readonly paragraphIndex?: number;
  readonly rowIndex?: number;
  readonly cellIndex?: number;
  readonly start: number;
  readonly end: number;
  readonly textStart?: number;
  readonly textEnd?: number;
}

function findMatchKey(
  match: Omit<PresentationFindMatch, "key">,
): string {
  return [
    match.slideIndex,
    match.slideId,
    match.blockIndex ?? -1,
    match.blockId ?? "",
    match.field,
    match.paragraphIndex ?? -1,
    match.rowIndex ?? -1,
    match.cellIndex ?? -1,
    match.start,
    match.end,
  ].join(":");
}

function appendStringMatches(
  destination: PresentationFindMatch[],
  text: string,
  query: string,
  options: FindOptions | undefined,
  structural: Omit<PresentationFindMatch, "key" | "start" | "end">,
  textOffset?: number,
): void {
  for (const range of literalMatchRanges(text, query, options)) {
    const match: Omit<PresentationFindMatch, "key"> = {
      ...structural,
      ...range,
      ...(textOffset === undefined
        ? {}
        : {
            textStart: textOffset + range.start,
            textEnd: textOffset + range.end,
          }),
    };
    destination.push({ ...match, key: findMatchKey(match) });
  }
}

/** Find's traversal: each searchable line with its structural coordinates. */
function visitSearchableTexts(
  slides: readonly AnyRecord[],
  visit: (
    text: string,
    structural: Omit<PresentationFindMatch, "key" | "start" | "end">,
    textOffset?: number,
    /** The runs `text` joins, when the line has them. */
    runs?: readonly RunLike[],
  ) => void,
): void {
  slides.forEach((slide, slideIndex) => {
    const slideId = asString(slide.id) || `slide-${slideIndex}`;
    if (mirroredTitleBlockIndex(slide) < 0) {
      visit(
        asString(slide.title),
        { slideId, slideIndex, field: "slide-title" },
      );
    }
    asArray(slide.blocks).forEach((value, blockIndex) => {
      const block = (value ?? {}) as AnyRecord;
      if (isLayoutGraphicOnSlide(block)) return;
      const blockId = asString(block.id) || `block-${blockIndex}`;
      const structural = { slideId, slideIndex, blockId, blockIndex };
      const shapeParagraphs =
        block.type === "shape" && asArray(block.paragraphs).length > 0;
      if (block.type === "text" || shapeParagraphs) {
        const paragraphs = asArray(block.paragraphs);
        if (paragraphs.length > 0) {
          let paragraphOffset = 0;
          paragraphs.forEach((paragraphValue, paragraphIndex) => {
            const paragraph = (paragraphValue ?? {}) as AnyRecord;
            const flatParagraphText = paragraphText(paragraph);
            const paragraphRuns = asArray(paragraph.runs) as RunLike[];
            visit(
              flatParagraphText,
              {
                ...structural,
                field: shapeParagraphs ? "shape-text" : "text",
                paragraphIndex,
              },
              paragraphOffset,
              paragraphRuns.length > 0 ? paragraphRuns : undefined,
            );
            paragraphOffset += flatParagraphText.length + 1;
          });
        } else {
          const runs = asArray(block.runs) as RunLike[];
          visit(
            runs.length > 0 ? runsText(runs) : asString(block.text),
            { ...structural, field: "text" },
            0,
            runs.length > 0 ? runs : undefined,
          );
        }
      } else if (block.type === "shape") {
        // A shape's own words (PPT-B027): searched like a text box's.
        const runs = asArray(block.runs) as RunLike[];
        visit(
          runs.length > 0 ? runsText(runs) : asString(block.text),
          { ...structural, field: "shape-text" },
          undefined,
          runs.length > 0 ? runs : undefined,
        );
      } else if (block.type === "table") {
        asArray(block.rows).forEach((row, rowIndex) => {
          asArray(row).forEach((cell, cellIndex) => {
            visit(
              asString((cell as AnyRecord)?.text),
              {
                ...structural,
                field: "table-cell",
                rowIndex,
                cellIndex,
              },
            );
          });
        });
      } else if (block.type === "chart") {
        visit(
          chartTitleSource(block).text,
          { ...structural, field: "chart-title" },
        );
        visit(
          asString(block.caption),
          { ...structural, field: "chart-caption" },
        );
      }
    });
  });
}

/**
 * Enumerate every replaceable match in deterministic deck order. A match may
 * span formatting runs but never a paragraph break, matching the writer.
 */
export function findMatchesInDeck(
  slides: readonly AnyRecord[],
  query: string,
  options?: FindOptions,
): PresentationFindMatch[] {
  if (!query) return [];
  const matches: PresentationFindMatch[] = [];
  visitSearchableTexts(slides, (text, structural, textOffset) =>
    appendStringMatches(matches, text, query, options, structural, textOffset),
  );
  return matches;
}

/**
 * Every line Find searches -- slide title, paragraph, shape words, table
 * cell, chart title and caption -- in the same deck order, with its slide.
 * Review ▸ Spelling walks these (PPT-A098) and then edits through
 * `replaceMatchInDeck`, so it changes exactly what Replace can change.
 */
export function presentationSearchableTexts(
  slides: readonly AnyRecord[],
): PresentationSearchableText[] {
  const lines: PresentationSearchableText[] = [];
  visitSearchableTexts(slides, (text, structural, _textOffset, runs) => {
    if (!text) return;
    lines.push({
      slideIndex: structural.slideIndex,
      text,
      lineKey: presentationFindLineKey(structural),
      ...(runs ? { runs } : {}),
    });
  });
  return lines;
}

export interface PresentationSearchableText {
  readonly slideIndex: number;
  readonly text: string;
  /** The line's place, as `presentationFindLineKey` gives it for a match. */
  readonly lineKey: string;
  /** The runs the text joins (their formatting), when the line has runs. */
  readonly runs?: readonly { readonly text: string; readonly style?: unknown }[];
}

/**
 * The line a match lies in, as a key: its slide, block, field, paragraph and
 * table cell, without its offsets. A match's `start` / `end` index into the
 * text of the searchable line with the same key.
 */
export function presentationFindLineKey(
  match: Omit<PresentationFindMatch, "key" | "start" | "end">,
): string {
  return [
    match.slideIndex,
    match.slideId,
    match.blockIndex ?? -1,
    match.blockId ?? "",
    match.field,
    match.paragraphIndex ?? -1,
    match.rowIndex ?? -1,
    match.cellIndex ?? -1,
  ].join(":");
}

/**
 * The one line of text a match's `start` / `end` index into: the slide
 * title, the paragraph (or the whole text, for a block without paragraphs),
 * the shape's words, the table cell, or the chart's title / caption. The Find
 * pane previews the hit inside it. Null when the deck no longer holds the
 * match's slide or block.
 */
export function presentationFindMatchLine(
  slides: readonly AnyRecord[],
  match: PresentationFindMatch,
): string | null {
  const slide = slides[match.slideIndex];
  if (!slide || (asString(slide.id) || `slide-${match.slideIndex}`) !== match.slideId) {
    return null;
  }
  if (match.field === "slide-title") return asString(slide.title);
  if (match.blockIndex === undefined) return null;
  const block = (asArray(slide.blocks)[match.blockIndex] ?? null) as AnyRecord | null;
  if (!block) return null;
  switch (match.field) {
    case "text": {
      const paragraphs = asArray(block.paragraphs);
      if (match.paragraphIndex !== undefined && paragraphs.length > 0) {
        return paragraphText((paragraphs[match.paragraphIndex] ?? {}) as AnyRecord);
      }
      const runs = asArray(block.runs) as RunLike[];
      return runs.length > 0 ? runsText(runs) : asString(block.text);
    }
    case "shape-text": {
      const runs = asArray(block.runs) as RunLike[];
      return runs.length > 0 ? runsText(runs) : asString(block.text);
    }
    case "table-cell": {
      const row = asArray(block.rows)[match.rowIndex ?? -1];
      return asString((asArray(row)[match.cellIndex ?? -1] as AnyRecord | undefined)?.text);
    }
    case "chart-title":
      return chartTitleSource(block).text;
    case "chart-caption":
      return asString(block.caption);
    default:
      return null;
  }
}

function replaceAuthenticatedSlice(
  text: string,
  match: PresentationFindMatch,
  query: string,
  replacement: string,
  options?: FindOptions,
): string | null {
  // The matched span need not be as long as the query: case folding can
  // change a character's length (PPT-B045).
  if (
    !query ||
    match.start < 0 ||
    match.end <= match.start ||
    match.end > text.length ||
    (options?.matchCase && match.end !== match.start + query.length)
  ) {
    return null;
  }
  const candidate = text.slice(match.start, match.end);
  const authenticated = options?.matchCase
    ? candidate === query
    : candidate.toLowerCase() === query.toLowerCase();
  if (
    !authenticated ||
    (options?.wholeWord &&
      !isWholeWordRange(text, match.start, match.end))
  ) {
    return null;
  }
  return `${text.slice(0, match.start)}${replacement}${text.slice(match.end)}`;
}

/**
 * A table cell's character runs when they spell exactly its `text` (the only
 * runs a cell draws; `presentation-cell-rich-text`), else none.
 */
function cellRuns(cell: AnyRecord): RunLike[] {
  const runs = asArray(cell.runs) as RunLike[];
  return runs.length > 0 && runsText(runs) === asString(cell.text) ? runs : [];
}

/**
 * Replace every match in a table cell. A cell with character runs is
 * rewritten run by run, so the characters outside the match keep their
 * formatting (PPT-B015); returns the SAME cell when nothing matched.
 */
function replaceAllInCell(
  cell: AnyRecord,
  query: string,
  replacement: string,
  options?: FindOptions,
): AnyRecord {
  const text = asString(cell.text);
  const ranges = literalMatchRanges(text, query, options);
  if (ranges.length === 0) return cell;
  const runs = cellRuns(cell);
  if (runs.length === 0) return { ...cell, text: replaceRanges(text, ranges, replacement) };
  const nextRuns = replaceRangesInRuns(runs, ranges, replacement);
  return { ...cell, runs: nextRuns, text: runsText(nextRuns) };
}

/**
 * Replace one authenticated match in a record holding `text` and optional
 * formatting `runs` (a paragraph, a run-formatted block). The match reads the
 * runs' continuous text; `null` when the live text no longer holds it.
 */
function replaceMatchInRunsRecord(
  record: AnyRecord,
  match: PresentationFindMatch,
  query: string,
  replacement: string,
  options?: FindOptions,
): AnyRecord | null {
  const runs = asArray(record.runs) as RunLike[];
  const text = runs.length > 0 ? runsText(runs) : asString(record.text);
  const nextText = replaceAuthenticatedSlice(
    text,
    match,
    query,
    replacement,
    options,
  );
  if (nextText === null || nextText === text) return null;
  if (runs.length === 0) return { ...record, text: nextText };
  if (rangeTouchesAtomicRun(runs, match.start, match.end)) {
    // Inside one equation: the same change in its LaTeX, else left whole.
    const edited = equationRangeReplaced(runs, match.start, match.end, replacement);
    return edited ? { ...record, runs: edited, text: runsText(edited) } : null;
  }
  const nextRuns = replaceRangeInRuns(runs, match.start, match.end, replacement);
  return { ...record, runs: nextRuns, text: runsText(nextRuns) };
}

/**
 * Replace one previously enumerated occurrence. Every structural coordinate and
 * the matched live substring are reauthenticated; any stale/malformed path is a
 * zero-write failure. The returned deck clones only the target ancestry.
 */
export function replaceMatchInDeck<S extends AnyRecord>(
  slides: readonly S[],
  match: PresentationFindMatch,
  query: string,
  replacement: string,
  options?: FindOptions,
): { slides: S[]; replaced: boolean; locked?: boolean } {
  const slide = slides[match.slideIndex];
  if (!slide || (asString(slide.id) || `slide-${match.slideIndex}`) !== match.slideId) {
    return { slides: [...slides], replaced: false };
  }
  if (match.field === "slide-title") {
    const nextTitle = replaceAuthenticatedSlice(
      asString(slide.title),
      match,
      query,
      replacement,
      options,
    );
    if (nextTitle === null || nextTitle === asString(slide.title)) {
      return { slides: [...slides], replaced: false };
    }
    const nextSlides = [...slides];
    nextSlides[match.slideIndex] = { ...slide, title: nextTitle } as S;
    return { slides: nextSlides, replaced: true };
  }

  const blocks = asArray(slide.blocks);
  const block = blocks[match.blockIndex ?? -1] as AnyRecord | undefined;
  if (
    !block ||
    (asString(block.id) || `block-${match.blockIndex}`) !== match.blockId
  ) {
    return { slides: [...slides], replaced: false };
  }
  if (lockedAgainstReplace(block)) {
    return { slides: [...slides], replaced: false, locked: true };
  }
  let nextBlock: AnyRecord | null = null;
  if (
    (match.field === "text" && block.type === "text") ||
    (match.field === "shape-text" &&
      block.type === "shape" &&
      match.paragraphIndex !== undefined)
  ) {
    const paragraphs = asArray(block.paragraphs);
    if (match.paragraphIndex !== undefined) {
      const paragraph = paragraphs[match.paragraphIndex] as AnyRecord | undefined;
      if (!paragraph) return { slides: [...slides], replaced: false };
      const nextParagraph = replaceMatchInRunsRecord(
        paragraph,
        match,
        query,
        replacement,
        options,
      );
      if (!nextParagraph) return { slides: [...slides], replaced: false };
      const nextParagraphs = [...paragraphs];
      nextParagraphs[match.paragraphIndex] = nextParagraph;
      nextBlock = {
        ...block,
        paragraphs: nextParagraphs,
        text: textFromParagraphRecords(nextParagraphs),
      };
    } else {
      nextBlock = replaceMatchInRunsRecord(
        block,
        match,
        query,
        replacement,
        options,
      );
    }
  } else if (match.field === "shape-text" && block.type === "shape") {
    nextBlock = replaceMatchInRunsRecord(block, match, query, replacement, options);
  } else if (match.field === "table-cell" && block.type === "table") {
    const rows = asArray(block.rows);
    const row = asArray(rows[match.rowIndex ?? -1]);
    const cell = row[match.cellIndex ?? -1] as AnyRecord | undefined;
    if (!cell) return { slides: [...slides], replaced: false };
    const runs = cellRuns(cell);
    const replaced = replaceMatchInRunsRecord(
      { text: cell.text, ...(runs.length > 0 ? { runs } : {}) },
      match,
      query,
      replacement,
      options,
    );
    if (!replaced) return { slides: [...slides], replaced: false };
    const nextRow = [...row];
    nextRow[match.cellIndex ?? -1] = {
      ...cell,
      text: replaced.text,
      ...(runs.length > 0 ? { runs: replaced.runs } : {}),
    };
    const nextRows = [...rows];
    nextRows[match.rowIndex ?? -1] = nextRow;
    nextBlock = { ...block, rows: nextRows };
  } else if (
    (match.field === "chart-title" || match.field === "chart-caption") &&
    block.type === "chart"
  ) {
    const title = match.field === "chart-title";
    const source = title ? chartTitleSource(block) : null;
    const current = source ? source.text : asString(block.caption);
    const nextText =
      title && !chartTitleReplaceable(block, options)
        ? null
        : replaceAuthenticatedSlice(current, match, query, replacement, options);
    if (nextText === null || nextText === current) {
      return { slides: [...slides], replaced: false };
    }
    nextBlock = source
      ? withChartTitle(block, nextText, source.v2)
      : { ...block, caption: nextText };
  }
  if (!nextBlock) return { slides: [...slides], replaced: false };
  const nextBlocks = [...blocks];
  nextBlocks[match.blockIndex ?? -1] = nextBlock;
  const nextSlides = [...slides];
  const renamed =
    mirroredTitleBlockIndex(slide) === match.blockIndex
      ? { title: textBlockPlainText(nextBlock) }
      : {};
  nextSlides[match.slideIndex] = { ...slide, blocks: nextBlocks, ...renamed } as S;
  return { slides: nextSlides, replaced: true };
}

/**
 * Replace every match in one block, keeping the per-run invariant. Returns a NEW
 * block when a replacement landed, or the SAME block reference when nothing
 * changed (so the caller can skip an unchanged slide). When a text block carries
 * `runs`/`paragraphs`, replacement goes through {@link replaceInRuns} so styled
 * spans keep their style AND the flat `text` is rebuilt from the new runs (the
 * invariant). Pure.
 */
export function replaceInBlock<T extends AnyRecord>(
  block: T,
  query: string,
  replacement: string,
  options?: FindOptions,
): T {
  if (!query) return block;
  const type = block.type;
  if (type === "text" || (type === "shape" && asArray(block.paragraphs).length > 0)) {
    const paragraphs = asArray(block.paragraphs);
    if (paragraphs.length > 0) {
      let changed = false;
      const nextParagraphs = paragraphs.map((p) => {
        const para = (p ?? {}) as AnyRecord;
        const runs = asArray(para.runs) as RunLike[];
        if (runs.length > 0) {
          const nextRuns = replaceInRuns(runs, query, replacement, options);
          const nextText = nextRuns.map((r) => r.text).join("");
          if (nextText !== asString(para.text)) {
            changed = true;
            return { ...para, runs: nextRuns, text: nextText };
          }
          return para;
        }
        const nextText = replaceAllInString(asString(para.text), query, replacement, options);
        if (nextText !== asString(para.text)) {
          changed = true;
          return { ...para, text: nextText };
        }
        return para;
      });
      if (!changed) return block;
      const nextText = textFromParagraphRecords(nextParagraphs);
      const next: AnyRecord = { ...block, paragraphs: nextParagraphs, text: nextText };
      // A shape label states its words twice, as rows and as flat runs, and
      
      
      // the save of the whole deck was refused). The flat runs take the same
      // replacement; where they would not spell the new words they are
      // rebuilt from the rows.
      const flat = asArray(block.runs) as RunLike[];
      if (flat.length > 0) {
        const nextFlat = replaceInRuns(flat, query, replacement, options);
        next.runs = runsText(nextFlat) === nextText ? nextFlat : flatRunsFromParagraphs(nextParagraphs as never);
      }
      return next as T;
    }
    const runs = asArray(block.runs) as RunLike[];
    if (runs.length > 0) {
      const nextRuns = replaceInRuns(runs, query, replacement, options);
      const nextText = nextRuns.map((r) => r.text).join("");
      if (nextText === asString(block.text)) return block;
      return { ...block, runs: nextRuns, text: nextText } as T;
    }
    const nextText = replaceAllInString(asString(block.text), query, replacement, options);
    return nextText === asString(block.text) ? block : ({ ...block, text: nextText } as T);
  }
  if (type === "shape") {
    const runs = asArray(block.runs) as RunLike[];
    if (runs.length > 0) {
      const nextRuns = replaceInRuns(runs, query, replacement, options);
      const nextText = runsText(nextRuns);
      return nextText === runsText(runs)
        ? block
        : ({ ...block, runs: nextRuns, text: nextText } as T);
    }
    if (typeof block.text !== "string") return block;
    const nextText = replaceAllInString(block.text, query, replacement, options);
    return nextText === block.text ? block : ({ ...block, text: nextText } as T);
  }
  if (type === "table") {
    if (lockedAgainstReplace(block)) return block;
    let changed = false;
    const nextRows = asArray(block.rows).map((row) =>
      asArray(row).map((cell) => {
        const c = (cell ?? {}) as AnyRecord;
        const next = replaceAllInCell(c, query, replacement, options);
        if (next !== c) changed = true;
        return next === c ? cell : next;
      }),
    );
    return changed ? ({ ...block, rows: nextRows } as T) : block;
  }
  if (type === "chart") {
    const source = chartTitleSource(block);
    const nextTitle = chartTitleReplaceable(block, options)
      ? replaceAllInString(source.text, query, replacement, options)
      : source.text;
    const nextCaption = lockedAgainstReplace(block)
      ? asString(block.caption)
      : replaceAllInString(asString(block.caption), query, replacement, options);
    if (nextTitle === source.text && nextCaption === asString(block.caption)) {
      return block;
    }
    let next: AnyRecord = { ...block };
    if (nextTitle !== source.text) next = withChartTitle(next, nextTitle, source.v2);
    if (typeof block.caption === "string") next.caption = nextCaption;
    return next as T;
  }
  return block;
}

/**
 * Replace every match across an ENTIRE deck. Returns `{ slides, replaced }`:
 * a NEW slides array (slides without a match returned by reference) and the
 * TOTAL number of occurrences substituted. A blank query yields the input
 * slides unchanged + `replaced: 0`. Pure — the editor commits the returned
 * slides through its normal undoable path.
 */
export function replaceAllInDeck<S extends AnyRecord>(
  slides: readonly S[],
  query: string,
  replacement: string,
  options?: FindOptions,
): { slides: S[]; replaced: number; lockedSkipped: number; equationSkipped: number } {
  if (!query) return { slides: [...slides], replaced: 0, lockedSkipped: 0, equationSkipped: 0 };
  let replaced = 0;
  let lockedSkipped = 0;
  // Matches inside an equation (or a field) Replace left whole: its LaTeX
  // could not follow the change (`mathTextReplaced`).
  let equationSkipped = 0;
  const nextSlides = slides.map((slide) => {
    let slideChanged = false;
    const mirrored = mirroredTitleBlockIndex(slide);
    const titleMatches =
      mirrored < 0 ? countMatches(asString(slide.title), query, options) : 0;
    const nextTitle =
      titleMatches > 0
        ? replaceAllInString(asString(slide.title), query, replacement, options)
        : asString(slide.title);
    if (titleMatches > 0) {
      replaced += titleMatches;
      slideChanged = true;
    }
    const nextBlocks = asArray(slide.blocks).map((block) => {
      const b = block as AnyRecord;
      if (isLayoutGraphicOnSlide(b)) return block;
      const skipped = unreplaceableMatches(b, query, options, replacement);
      if (lockedAgainstReplace(b)) lockedSkipped += skipped;
      else if (b.type === "text" || b.type === "shape" || b.type === "table") equationSkipped += skipped;
      const matches = countMatchesInBlock(b, query, options) - skipped;
      if (matches <= 0) return block;
      replaced += matches;
      slideChanged = true;
      return replaceInBlock(b, query, replacement, options);
    });
    if (!slideChanged) return slide;
    const next: AnyRecord = { ...slide, blocks: nextBlocks };
    if (mirrored >= 0) next.title = textBlockPlainText(nextBlocks[mirrored] as AnyRecord);
    else if (typeof slide.title === "string") next.title = nextTitle;
    return next as S;
  });
  return { slides: nextSlides, replaced, lockedSkipped, equationSkipped };
}
