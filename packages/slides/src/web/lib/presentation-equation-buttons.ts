/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Equations in a slide's text, as the editor's commands change them
 * (PowerPoint's Insert > Equation and the Equation tab, PPT-A052).
 *
 * An equation is one ATOMIC run (`TextRun.math`): its `text` is what the
 * server's Office Math draws in the line (`POST
 * /api/v1/presentation/equation-text`), its `math.latex` the equation. These
 * pure transforms put such a run into a text box the way PowerPoint places a
 * new equation:
 *
 * - at the caret (or over the selection) of a text box being edited -- a
 *   paragraph with nothing else in it takes the equation as its own line
 *   (`display`, Office's `m:oMathPara`), anything else inline;
 * - as a new last paragraph of a text box that is selected but not being
 *   edited, standing on its own line;
 * - in a new text box in the middle of the slide when nothing is selected.
 *
 * Every change goes through the editor's own text reconciliation
 * (`reconcilePresentationTextBlock`) so the box keeps its paragraphs and runs
 * exactly as typing would, and then marks the inserted characters as the one
 * equation run. No DOM, no network: the caller fetched the drawn characters.
 */

import { designPxToPoints } from "@mrnote/core/web/lib/presentation/text-units.ts";
import type { PresentationTextBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { textFromParagraphs } from "@mrnote/core/web/lib/presentation-paragraphs.ts";
import {
  runCarriesMath,
  runsFromText,
  textFromRuns,
  type RunMath,
  type TextRun,
} from "@mrnote/core/web/lib/presentation-runs.ts";
// Sizes to the hundredth (DrawingML's): whole px turned 10.5 pt into 11.
import { mergeRuns, normalizeRuns, textRunSlice } from "@mrnote/core/web/lib/presentation-point-runs.ts";
import { reconcilePresentationTextBlock, textBoxBlock } from "@mrnote/core/web/lib/presentation-shape-buttons.ts";
import { isTextEditLocked } from "@mrnote/core/web/lib/presentation-text-case.ts";
import {
  normalizeParagraphsWithRuns,
  ownsParagraphRichLayout,
} from "@mrnote/core/web/lib/presentation-text-buttons.ts";
export * from "./presentation-equation-box";
import { EQUATION_BOX_TEXT_INSETS, equationBoxSizeFor, presentationEquationTextIsUsable } from "./presentation-equation-box";
import type { PresentationEquation } from "./presentation-equation-box";

/** Where one equation run sits in a text box's flat text. */
export interface PresentationEquationSpan {
  readonly start: number;
  readonly end: number;
  readonly latex: string;
  readonly display: boolean;
}

/** The result of putting an equation in: the box, and the run's offsets. */
export interface PresentationEquationEdit<Block extends PresentationTextBlock = PresentationTextBlock> {
  readonly block: Block;
  readonly start: number;
  readonly end: number;
}

function runMathOf(equation: PresentationEquation): RunMath {
  return equation.display ? { latex: equation.latex, display: true } : { latex: equation.latex };
}

/** `runs` with `[from, to)` turned into ONE run carrying `math`. */
function runsWithEquation(
  runs: readonly TextRun[],
  from: number,
  to: number,
  math: RunMath,
): TextRun[] | null {
  const total = textFromRuns(runs).length;
  if (from < 0 || to > total || from >= to) return null;
  const inside = textRunSlice(runs, from, to);
  const text = inside.map((piece) => piece.text).join("");
  // The equation takes the format of its first character, as text typed at
  // that place would (Office writes the zone's own math font over it).
  const style = inside[0]?.style;
  const equation: TextRun = style ? { text, style: { ...style }, math } : { text, math };
  return mergeRuns([...textRunSlice(runs, 0, from), equation, ...textRunSlice(runs, to, total)]);
}

/** The box with `[from, to)` of its text marked as one equation run, or null. */
function textBlockWithEquationRun<Block extends PresentationTextBlock>(
  block: Block,
  from: number,
  to: number,
  math: RunMath,
): Block | null {
  const paragraphs = normalizeParagraphsWithRuns(block.paragraphs);
  if (paragraphs && ownsParagraphRichLayout(paragraphs)) {
    let base = 0;
    for (let index = 0; index < paragraphs.length; index += 1) {
      const paragraph = paragraphs[index]!;
      const end = base + paragraph.text.length;
      if (from >= base && to <= end) {
        const runs = normalizeRuns(paragraph.runs) ?? runsFromText(paragraph.text);
        const marked = runsWithEquation(runs, from - base, to - base, math);
        if (!marked) return null;
        const nextParagraphs = paragraphs.map((item, at) =>
          at === index ? { ...item, text: textFromRuns(marked), runs: marked } : item,
        );
        return { ...block, paragraphs: nextParagraphs, text: textFromParagraphs(nextParagraphs) };
      }
      base = end + 1;
    }
    return null;
  }
  const runs = normalizeRuns(block.runs) ?? runsFromText(block.text);
  const marked = runsWithEquation(runs, from, to, math);
  if (!marked) return null;
  return { ...block, text: textFromRuns(marked), runs: marked };
}

/** Every equation run of the box, in text order, with its flat offsets. */
export function presentationEquationSpans(block: PresentationTextBlock): PresentationEquationSpan[] {
  const spans: PresentationEquationSpan[] = [];
  const collect = (runs: readonly TextRun[] | undefined, base: number) => {
    let cursor = base;
    for (const run of runs ?? []) {
      const end = cursor + run.text.length;
      if (runCarriesMath(run) && run.math) {
        spans.push({ start: cursor, end, latex: run.math.latex, display: run.math.display === true });
      }
      cursor = end;
    }
  };
  const paragraphs = normalizeParagraphsWithRuns(block.paragraphs);
  if (paragraphs && ownsParagraphRichLayout(paragraphs)) {
    let base = 0;
    for (const paragraph of paragraphs) {
      collect(normalizeRuns(paragraph.runs), base);
      base += paragraph.text.length + 1;
    }
    return spans;
  }
  collect(normalizeRuns(block.runs), 0);
  return spans;
}

/**
 * The equation a selection of the box is ON: the selection lies within one
 * equation's characters, or covers exactly one equation (what a click on the
 * atom or a double-click selects). A caret beside an equation is not on it.
 */
export function presentationEquationSpanAt(
  block: PresentationTextBlock,
  start: number,
  end: number,
): PresentationEquationSpan | null {
  const from = Math.min(start, end);
  const to = Math.max(start, end);
  for (const span of presentationEquationSpans(block)) {
    if (from === to) {
      if (from > span.start && from < span.end) return span;
      continue;
    }
    if (from >= span.start && to <= span.end) return span;
  }
  return null;
}

/** The equation whose characters begin at `offset`, else null. */
export function presentationEquationSpanStartingAt(
  block: PresentationTextBlock,
  offset: number,
): PresentationEquationSpan | null {
  return presentationEquationSpans(block).find((span) => span.start === offset) ?? null;
}

/** Whether `[from, to)` of `text` is all of its paragraph once removed. */
function fillsItsParagraph(text: string, from: number, to: number): boolean {
  const before = from === 0 || text[from - 1] === "\n";
  const after = to === text.length || text[to] === "\n";
  return before && after;
}

/**
 * Whether an equation put over `[start, end)` of the box stands as its own
 * line: nothing else is left in its paragraph (PowerPoint writes such a zone
 * as `m:oMathPara`). Asked before the server is, which draws both the same.
 */
export function presentationEquationDisplayAt(
  block: PresentationTextBlock,
  selection: { readonly start: number; readonly end: number },
): boolean {
  const text = block.text ?? "";
  const clamp = (value: number) => Math.max(0, Math.min(text.length, Math.trunc(value)));
  const from = Math.min(clamp(selection.start), clamp(selection.end));
  const to = Math.max(clamp(selection.start), clamp(selection.end));
  return fillsItsParagraph(text, from, to);
}

/**
 * Put `equation` over `[start, end)` of the box's text (a caret when equal),
 * as typing would put characters there. The equation stands as its own line
 * when nothing else is left in its paragraph and the caller did not say
 * otherwise. Null when the box cannot take it.
 */
export function textBlockWithEquationAt<Block extends PresentationTextBlock>(
  block: Block,
  selection: { readonly start: number; readonly end: number },
  equation: PresentationEquation,
): PresentationEquationEdit<Block> | null {
  if (!presentationEquationTextIsUsable(equation.text)) return null;
  const text = block.text ?? "";
  const clamp = (value: number) => Math.max(0, Math.min(text.length, Math.trunc(value)));
  const from = Math.min(clamp(selection.start), clamp(selection.end));
  const to = Math.max(clamp(selection.start), clamp(selection.end));
  const display = equation.display ?? fillsItsParagraph(text, from, to);
  const nextText = `${text.slice(0, from)}${equation.text}${text.slice(to)}`;
  const reconciled = reconcilePresentationTextBlock(block, nextText, { start: from, end: to }) as Block;
  const end = from + equation.text.length;
  const marked = textBlockWithEquationRun(reconciled, from, end, runMathOf({ ...equation, display }));
  if (!marked) return null;
  const span = presentationEquationSpanStartingAt(marked, from);
  if (!span || span.end !== end) return null;
  return { block: marked, start: from, end };
}

/**
 * A selected box not being edited takes the equation as a new last paragraph
 * standing on its own line; an empty box takes it as its only paragraph.
 */
export function textBlockWithEquationAppended<Block extends PresentationTextBlock>(
  block: Block,
  equation: PresentationEquation,
): PresentationEquationEdit<Block> | null {
  const text = block.text ?? "";
  if (text === "") return textBlockWithEquationAt(block, { start: 0, end: 0 }, { ...equation, display: true });
  const withParagraph = reconcilePresentationTextBlock(block, `${text}\n`, {
    start: text.length,
    end: text.length,
  }) as Block;
  const at = withParagraph.text.length;
  return textBlockWithEquationAt(withParagraph, { start: at, end: at }, { ...equation, display: true });
}

/**
 * The equation over `span` replaced by `equation` (the editor's Done): the
 * new characters and LaTeX in the same place, the same line placement unless
 * the caller says otherwise.
 */
export function textBlockWithEquationReplaced<Block extends PresentationTextBlock>(
  block: Block,
  span: PresentationEquationSpan,
  equation: PresentationEquation,
): PresentationEquationEdit<Block> | null {
  const current = presentationEquationSpanStartingAt(block, span.start);
  if (!current || current.end !== span.end || current.latex !== span.latex) return null;
  return textBlockWithEquationAt(
    block,
    { start: span.start, end: span.end },
    { ...equation, display: equation.display ?? span.display },
  );
}

type Frame = { readonly left: number; readonly top: number; readonly width: number; readonly height: number };

/** The new-box frame PowerPoint uses before the equation has a size: mid-slide. */
export function defaultEquationBoxFrame(slideSize: { readonly width: number; readonly height: number }): Frame {
  const width = Math.round(slideSize.width * 0.34);
  const height = Math.round(slideSize.height * 0.14);
  return {
    left: Math.round((slideSize.width - width) / 2),
    top: Math.round((slideSize.height - height) / 2),
    width,
    height,
  };
}

/** A box narrower or shorter than this (EMU, about half a px) than its equation needs is refitted. */
export const EQUATION_BOX_FIT_TOLERANCE_EMU = 5_000;

/**
 * The equation box sized to its equation ("Resize shape to fit text", no
 * wrap): the equation as drawn (`contentPx`, design px) plus the box's
 * insets. A new box is centred on `around` (`anchor: "center"`, the default:
 * mid-slide, or where the ink was); a box already on the slide keeps its
 * top-left and grows to the right and down (`anchor: "start"`), as a
 * no-wrap box does when its text changes, so its equation never shifts.
 * Without a measurement (nothing laid out) the box keeps `around`.
 */
export function equationBoxFrame(
  contentPx: { readonly width: number; readonly height: number } | null,
  around: Frame,
  anchor: "center" | "start" = "center",
): Frame {
  if (!contentPx || !(contentPx.width > 0) || !(contentPx.height > 0)) {
    return {
      left: Math.round(around.left),
      top: Math.round(around.top),
      width: Math.max(1, Math.round(around.width)),
      height: Math.max(1, Math.round(around.height)),
    };
  }
  const { width, height } = equationBoxSizeFor(contentPx);
  if (anchor === "start") {
    return { left: Math.round(around.left), top: Math.round(around.top), width, height };
  }
  const centerX = around.left + around.width / 2;
  const centerY = around.top + around.height / 2;
  return {
    left: Math.round(centerX - width / 2),
    top: Math.round(centerY - height / 2),
    width,
    height,
  };
}

/**
 * The equation of an EQUATION BOX: a text box whose whole text is one
 * display equation. PowerPoint's Insert > Equation box is exactly that -- a
 * `txBox` shape whose one paragraph is an `a14:m`/`m:oMathPara` -- so an
 * imported one is an equation box too. Null for anything else (words beside
 * the equation, a shape's label, a SmartArt member).
 */
export function presentationEquationBoxSpan(block: unknown): PresentationEquationSpan | null {
  if (!block || typeof block !== "object") return null;
  const record = block as Record<string, unknown>;
  if (record.type !== "text" || typeof record.text !== "string" || record.text === "") return null;
  if (typeof record.smartArtRole === "string" && record.smartArtRole !== "") return null;
  const spans = presentationEquationSpans(block as PresentationTextBlock);
  if (spans.length !== 1) return null;
  const span = spans[0]!;
  return span.display && span.start === 0 && span.end === record.text.length ? span : null;
}

/**
 * Insert > Equation with nothing selected: PowerPoint makes a text box in the
 * middle of the slide holding the equation on its own line -- the equation
 * box, not wrapping (`wrap="none"`), at PowerPoint's insets.
 */
export function equationTextBoxBlock(
  id: string,
  slideSize: { readonly width: number; readonly height: number },
  equation: PresentationEquation,
  at?: { readonly left: number; readonly top: number; readonly width: number; readonly height: number },
  fontSizePx?: number,
): PresentationTextBlock | null {
  const frame = at ? equationBoxFrame(null, at) : defaultEquationBoxFrame(slideSize);
  // A free text box (`kind: "textbox"`) is a text block to every text command.
  const { autofit: _sized, ...box } = textBoxBlock(id, frame) as PresentationTextBlock;
  void _sized;
  // Its own fit sizes it to its equation (EQUATION_BOX_FIT_COMMIT).
  const equationBox: PresentationTextBlock = {
    ...box,
    style: {
      ...(box.style ?? {}),
      textWrap: false,
      // The box states points (`textBoxBlock`); the caller measured in px.
      ...(fontSizePx !== undefined && Number.isFinite(fontSizePx) && fontSizePx > 0 ? { fontSize: designPxToPoints(fontSizePx) } : {}),
    } as PresentationTextBlock["style"],
    textInsets: { ...EQUATION_BOX_TEXT_INSETS },
  };
  return textBlockWithEquationAt(equationBox, { start: 0, end: 0 }, { ...equation, display: true })?.block ?? null;
}

/** Where Insert > Equation puts the new equation, read from the selection. */
export type PresentationEquationTarget =
  /** The caret (or selection) of a text box being edited. */
  | { readonly kind: "caret"; readonly blockId: string; readonly start: number; readonly end: number }
  /** A text box selected by its border: a new last paragraph. */
  | { readonly kind: "append"; readonly blockId: string }
  /** Nothing that takes text: a new box in the middle of the slide, or in
   * `box` when the caller knows where it goes (Ink to Math: where the ink was). */
  | {
      readonly kind: "new-box";
      readonly box?: { readonly left: number; readonly top: number; readonly width: number; readonly height: number };
      /** The new box's font size (px); the default text box size when absent. */
      readonly fontSizePx?: number;
    };

type SlideLike = { readonly blocks: readonly unknown[]; readonly [key: string]: unknown };

/**
 * Which shapes hold words an equation can join (PowerPoint inserts into a
 * shape's text frame too). The editor knows its own rule
 * (`presentationShapeBlockTakesText`); without one only text boxes do.
 */
export type ShapeTakesEquation = (block: unknown) => boolean;

function asWritableTextBlock(block: unknown, shapeTakesEquation?: ShapeTakesEquation): PresentationTextBlock | null {
  if (!block || typeof block !== "object") return null;
  const record = block as Record<string, unknown>;
  // A shape's label has a text box's model (text, runs, paragraphs): the
  // same transforms apply to it.
  const shape = record.type === "shape" && shapeTakesEquation?.(block) === true;
  if (!shape && (record.type !== "text" || typeof record.text !== "string")) return null;
  if (shape && record.text !== undefined && typeof record.text !== "string") return null;
  // A SmartArt member's text belongs to the diagram's outline: PowerPoint
  // inserts no equation into it. Lock leaves the words open; the file's own
  // `noTextEdit` does not.
  if (typeof record.smartArtRole === "string" && record.smartArtRole !== "") return null;
  if (isTextEditLocked(record)) return null;
  return (shape && typeof record.text !== "string" ? { ...record, text: "" } : block) as PresentationTextBlock;
}

/**
 * PowerPoint's placement rule for a new equation: into the one selected text
 * box at its live caret, else as a new paragraph of that box, else into a new
 * box (several objects, a picture, a table or nothing selected).
 */
export function presentationEquationTarget(
  slide: SlideLike,
  selectedBlockIds: readonly string[],
  liveSelection: { readonly blockId: string; readonly start: number; readonly end: number } | null,
  shapeTakesEquation?: ShapeTakesEquation,
): PresentationEquationTarget {
  if (selectedBlockIds.length === 1) {
    const blockId = selectedBlockIds[0]!;
    const block = asWritableTextBlock(
      slide.blocks.find((item) => (item as { id?: unknown })?.id === blockId),
      shapeTakesEquation,
    );
    if (block) {
      return liveSelection?.blockId === blockId
        ? { kind: "caret", blockId, start: liveSelection.start, end: liveSelection.end }
        : { kind: "append", blockId };
    }
  }
  return { kind: "new-box" };
}

/** Whether the equation for `target` stands on its own line (asked before the server is). */
export function presentationEquationTargetDisplay(
  slide: SlideLike,
  target: PresentationEquationTarget,
  shapeTakesEquation?: ShapeTakesEquation,
): boolean {
  if (target.kind !== "caret") return true;
  const block = asWritableTextBlock(
    slide.blocks.find((item) => (item as { id?: unknown })?.id === target.blockId),
    shapeTakesEquation,
  );
  return block ? presentationEquationDisplayAt(block, target) : true;
}

/** The slide with the equation placed at `target`, the box it went into, and its offsets. */
export function slideWithEquation<Slide extends SlideLike>(
  slide: Slide,
  target: PresentationEquationTarget,
  equation: PresentationEquation,
  options: {
    readonly newId: () => string;
    readonly slideSize: { readonly width: number; readonly height: number };
    readonly maxBlocks?: number;
    readonly shapeTakesEquation?: ShapeTakesEquation;
  },
): { readonly slide: Slide; readonly blockId: string; readonly start: number; readonly end: number } | null {
  if (target.kind === "new-box") {
    if (options.maxBlocks !== undefined && slide.blocks.length >= options.maxBlocks) return null;
    const id = options.newId();
    const box = equationTextBoxBlock(id, options.slideSize, equation, target.box, target.fontSizePx);
    if (!box) return null;
    return { slide: { ...slide, blocks: [...slide.blocks, box] }, blockId: id, start: 0, end: box.text.length };
  }
  const index = slide.blocks.findIndex((item) => (item as { id?: unknown })?.id === target.blockId);
  const block = asWritableTextBlock(slide.blocks[index], options.shapeTakesEquation);
  if (index < 0 || !block) return null;
  const edit =
    target.kind === "caret"
      ? textBlockWithEquationAt(block, target, equation)
      : textBlockWithEquationAppended(block, equation);
  if (!edit) return null;
  const blocks = [...slide.blocks];
  blocks[index] = edit.block;
  return { slide: { ...slide, blocks }, blockId: target.blockId, start: edit.start, end: edit.end };
}
