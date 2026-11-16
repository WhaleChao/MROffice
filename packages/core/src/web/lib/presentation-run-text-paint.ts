/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * WordArt on the characters of a slide's text box or shape label: the text
 * paint (fill, outline, effects) as a property of each run, as DrawingML
 * states it on `a:rPr`, so one box holds several WordArt styles.
 *
 * PowerPoint applies a WordArt style, Text Fill, Text Outline and Text
 * Effects to the selected characters while the words are typed in, and to
 * every character when the object itself is selected (Microsoft Support,
 * "Add or delete a fill, outline, or effect for text or WordArt": each
 * procedure starts "select the text you want to change"). A caret alone
 * formats the word it is in, as the deck's Font commands do (PPT-B006).
 *
 * The model: a run's own `textPaint` is what its characters wear; a run
 * without one wears the block's `textPaint` (the paint every character of
 * the box shares). The writers here keep it canonical -- a paint every
 * character wears is the block's, and a block whose characters differ keeps
 * its paints on the runs and none of its own -- so a box styled whole reads,
 * draws and saves exactly as before.
 */

import type {
  PresentationBlock,
  PresentationShapeBlock,
  PresentationTextBlock,
} from "../components/console/editor/native-document/documentCommands";
import {
  normalizeOfficeTextPaint,
  officeTextPaintWithFontColor,
  type OfficeTextPaint,
} from "./office-text-paint";
import { shapeTextPaintPatch } from "./presentation-shape-buttons";
import {
  formatTextBlockRunSelection,
  normalizeParagraphsWithRuns,
  ownsParagraphRichLayout,
  presentationEditableText,
  selectedTextRange,
  type LiveTextSelection,
} from "./presentation-text-buttons";
import {
  normalizeRuns,
  runsFromText,
  sameRunTextPaint,
  type RunStyle,
  type TextRun,
} from "./presentation-point-runs";

/**
 * What a WordArt command does to one run's paint: the paint it wears now
 * (null: none) and its own style (for the colour a first paint starts from),
 * to the paint it wears after (null: none).
 */
export type RunTextPaintTransform = (
  paint: OfficeTextPaint | null,
  style: RunStyle | undefined,
) => OfficeTextPaint | null;

type TextFrame = PresentationTextBlock | PresentationShapeBlock;

function framePaint(block: PresentationBlock): OfficeTextPaint | null {
  const raw = (block as { textPaint?: unknown }).textPaint;
  if (!raw) return null;
  try {
    return normalizeOfficeTextPaint(raw) ?? null;
  } catch {
    return null;
  }
}

/** Whether `block` is a text box or a shape whose words can wear WordArt. */
function isTextFrame(block: PresentationBlock): block is TextFrame {
  return block.type === "text" || block.type === "shape";
}

/** The block's words as rows of runs, each row's offset in the editable text. */
function frameRows(block: TextFrame): { runs: TextRun[]; start: number }[] {
  const paragraphs = normalizeParagraphsWithRuns((block as { paragraphs?: unknown }).paragraphs);
  if (paragraphs && ownsParagraphRichLayout(paragraphs)) {
    let start = 0;
    return paragraphs.map((paragraph) => {
      const row = { runs: normalizeRuns(paragraph.runs) ?? runsFromText(paragraph.text), start };
      start += paragraph.text.length + 1;
      return row;
    });
  }
  const text = typeof block.text === "string" ? block.text : "";
  return [{ runs: normalizeRuns((block as { runs?: unknown }).runs) ?? runsFromText(text), start: 0 }];
}

/** Every run of the block's words with its own paint, in order. */
function frameRunPaints(block: TextFrame): Array<OfficeTextPaint | null> {
  const paints: Array<OfficeTextPaint | null> = [];
  for (const row of frameRows(block)) {
    for (const run of row.runs) {
      if (run.text.length > 0) paints.push(run.style?.textPaint ?? null);
    }
  }
  return paints;
}

/** Whether any run of the block's words wears WordArt of its own. */
export function presentationBlockHasRunTextPaint(block: PresentationBlock | null | undefined): boolean {
  return Boolean(block && isTextFrame(block) && frameRunPaints(block).some(Boolean));
}

/**
 * The block's runs rewritten by `map`: its rows' runs where the rows draw
 * the words (`ownsParagraphRichLayout`), else its flat runs. Everything else
 * a row or the block states is kept as it is.
 */
function withEveryRun(block: TextFrame, map: (run: TextRun) => TextRun): TextFrame {
  const next = { ...block } as TextFrame & { runs?: TextRun[]; paragraphs?: unknown };
  const rewrite = (runs: readonly TextRun[]) => normalizeRuns(runs.map(map));
  const paragraphs = normalizeParagraphsWithRuns((block as { paragraphs?: unknown }).paragraphs);
  const raw = (block as { paragraphs?: unknown }).paragraphs;
  if (paragraphs && ownsParagraphRichLayout(paragraphs) && Array.isArray(raw)) {
    next.paragraphs = raw.map((record, index) => {
      const paragraph = paragraphs[index];
      if (!paragraph || !record || typeof record !== "object") return record;
      const runs = rewrite(normalizeRuns(paragraph.runs) ?? runsFromText(paragraph.text));
      const row = { ...(record as Record<string, unknown>) };
      if (runs) row.runs = runs;
      else delete row.runs;
      return row;
    });
    return next;
  }
  const text = typeof block.text === "string" ? block.text : "";
  const runs = rewrite(normalizeRuns((block as { runs?: unknown }).runs) ?? runsFromText(text));
  if (runs) next.runs = runs;
  else delete next.runs;
  return next;
}

function runWithPaint(run: TextRun, paint: OfficeTextPaint | null): TextRun {
  const style = { ...(run.style ?? {}) } as RunStyle;
  if (paint) style.textPaint = paint;
  else delete style.textPaint;
  return Object.keys(style).length ? { ...run, style } : { ...run, style: undefined };
}

/**
 * The block's own paint moved onto every run that has none, and taken off
 * the block: a range is then formatted without the rest of the box changing.
 */
export function presentationBlockWithRunTextPaints<T extends PresentationBlock>(block: T): T {
  if (!isTextFrame(block)) return block;
  const own = framePaint(block);
  if (!own) return block;
  const next = withEveryRun(block, (run) =>
    run.style?.textPaint ? run : runWithPaint(run, own),
  ) as TextFrame & { textPaint?: unknown };
  delete next.textPaint;
  return next as unknown as T;
}

/**
 * The block kept canonical: a paint every character wears is the block's
 * (its runs wear none of their own); a block whose characters differ keeps
 * the paints on the runs and has none of its own.
 */
export function presentationBlockWithCanonicalTextPaint<T extends PresentationBlock>(block: T): T {
  if (!isTextFrame(block)) return block;
  const paints = frameRunPaints(block);
  if (paints.length === 0 || !paints.some(Boolean)) return block;
  const first = paints[0];
  const uniform = first !== null && paints.every((paint) => paint !== null && sameRunTextPaint(paint, first));
  if (!uniform) {
    const next = { ...block } as TextFrame & { textPaint?: unknown };
    delete next.textPaint;
    return next as unknown as T;
  }
  const bare = withEveryRun(block, (run) => runWithPaint(run, null));
  if (bare.type === "shape") {
    return { ...bare, ...shapeTextPaintPatch(bare as PresentationShapeBlock, first) } as unknown as T;
  }
  return { ...bare, textPaint: first } as unknown as T;
}

/**
 * The paint the characters a WordArt command acts on wear: one paint, null
 * (none), or undefined when they differ (the gallery and the colour buttons
 * then show no current value). `selection` null: every character.
 */
export function presentationTextPaintAt(
  block: PresentationBlock | null | undefined,
  selection: LiveTextSelection,
): OfficeTextPaint | null | undefined {
  if (!block || !isTextFrame(block)) return null;
  const own = framePaint(block);
  const text = presentationEditableText(block as PresentationTextBlock) ?? "";
  let [from, to] = selection && selection.blockId === block.id
    ? selectedTextRange(block.id, text, selection)
    : [0, text.length];
  if (from === to && selection && selection.blockId === block.id) {
    // A caret between words: what is typed there wears the paint before it.
    from = Math.max(0, from - 1);
    to = Math.max(from, Math.min(text.length, from + 1));
  }
  if (from === to) return own;
  let found: OfficeTextPaint | null | undefined;
  let seen = false;
  for (const row of frameRows(block)) {
    let at = row.start;
    for (const run of row.runs) {
      const start = at;
      at += run.text.length;
      if (run.text.length === 0 || at <= from || start >= to) continue;
      const paint = run.style?.textPaint ?? own;
      if (!seen) {
        found = paint;
        seen = true;
      } else if (!(paint === found || (paint && found && sameRunTextPaint(paint, found)))) {
        return undefined;
      }
    }
  }
  return seen ? (found ?? null) : own;
}

/**
 * The block with a WordArt command applied to the characters it acts on:
 * the selected ones while the words are typed in (a caret: the word it is
 * in), else every character. `runPatch` is what the command sets on those
 * characters besides their paint (a gallery style's bold). Null when nothing
 * changes or there is nothing to format (a caret between words).
 */
export function presentationBlockWithTextPaintIn(
  block: PresentationBlock,
  selection: LiveTextSelection,
  transform: RunTextPaintTransform,
  runPatch: Partial<RunStyle> = {},
): PresentationBlock | null {
  if (!isTextFrame(block)) return null;
  const text = presentationEditableText(block as PresentationTextBlock) ?? "";
  const typed = selection !== null && selection.blockId === block.id;
  if (text === "") return null;
  const range: LiveTextSelection = typed ? selection : { blockId: block.id, start: 0, end: text.length };
  const formatted = formatTextBlockRunSelection(
    presentationBlockWithRunTextPaints(block) as PresentationTextBlock,
    block.id,
    range,
    (style) => {
      const paint = transform(style?.textPaint ?? null, style);
      return {
        ...runPatch,
        textPaint: paint ?? undefined,
        // The paint owns the glyphs' fill and outline now (`a:rPr`'s fill
        // group and `a:ln` are the paint's), as the import reads them.
        ...(paint
          ? {
            color: undefined, textOutline: undefined, textOutlineWidthPt: undefined, textOutlineJoin: undefined,
            textOutlineMiterLimit: undefined, textOutlineCap: undefined, textOutlineDash: undefined,
          }
          : {}),
      };
    },
  );
  if (!formatted) return null;
  return presentationBlockWithCanonicalTextPaint(formatted as unknown as PresentationBlock);
}

/**
 * Font Color on characters that wear a paint: the colour is their fill,
 * outline and effects kept, a paint left with only a flat fill given way to
 * the plain colour (lib/office-text-paint `officeTextPaintWithFontColor`).
 */
export function presentationFontColorOnPaint(
  paint: OfficeTextPaint,
  color: string | undefined,
  automaticFill?: string,
): OfficeTextPaint | null {
  return officeTextPaintWithFontColor(paint, color, { automaticFill });
}

/** Whether a run patch (or what a patch function made of a run) states a font colour. */
function statesColor(patch: Partial<RunStyle>): boolean {
  return Object.hasOwn(patch, "color");
}

/**
 * {@link formatTextBlockRunSelection} with Font Color reaching the paint the
 * characters wear: every Home, Font dialog, typed-at-caret and paste path
 * that writes a run `color` goes through here. A patch that states no
 * colour, or a box that wears no paint, formats exactly as before.
 */
export function formatPresentationTextRuns(
  block: PresentationTextBlock,
  blockId: string,
  selection: LiveTextSelection,
  requested: Partial<RunStyle> | ((style: RunStyle | undefined) => Partial<RunStyle>),
  options: { readonly automaticFill?: string } = {},
): PresentationTextBlock | null {
  // A colour given to a hyperlink's words is the colour it shows, as
  // PowerPoint writes it (`ahyp:hlinkClr val="tx"`); taken away, the link
  // wears the theme's Hyperlink colour again. Without it the canvas kept the
  
  const linkAware = (made: Partial<RunStyle>, style: RunStyle | undefined): Partial<RunStyle> =>
    statesColor(made) && style?.href
      ? { ...made, linkColor: made.color ? "text" : undefined }
      : made;
  const patch =
    typeof requested === "function"
      ? (style: RunStyle | undefined) => linkAware(requested(style), style)
      : statesColor(requested)
        ? (style: RunStyle | undefined) => linkAware(requested, style)
        : requested;
  const frame = block as unknown as PresentationBlock;
  const worn = framePaint(frame) !== null || presentationBlockHasRunTextPaint(frame);
  if (!worn) return formatTextBlockRunSelection(block, blockId, selection, patch);
  let colored = typeof patch !== "function" && statesColor(patch);
  if (typeof patch === "function") {
    // A patch function shows what it states only per run: format once,
    // and go through the paints only when some run was given a colour.
    const plain = formatTextBlockRunSelection(block, blockId, selection, (style) => {
      const made = patch(style);
      if (statesColor(made)) colored = true;
      return made;
    });
    if (!colored) return plain;
  }
  if (!colored) return formatTextBlockRunSelection(block, blockId, selection, patch);
  const text = presentationEditableText(block) ?? "";
  if (text === "") {
    // An empty box: the colour is for the words typed into it, which wear
    // the box's paint.
    const styled = formatTextBlockRunSelection(block, blockId, selection, patch) ?? block;
    const own = framePaint(frame);
    if (!own) return styled === block ? null : styled;
    const resolved = typeof patch === "function" ? patch(undefined) : patch;
    const paint = presentationFontColorOnPaint(own, resolved.color ?? undefined, options.automaticFill);
    const next = { ...styled } as PresentationTextBlock & { textPaint?: unknown };
    if (paint) next.textPaint = paint;
    else delete next.textPaint;
    return next;
  }
  const formatted = formatTextBlockRunSelection(
    presentationBlockWithRunTextPaints(frame) as unknown as PresentationTextBlock,
    blockId,
    selection,
    (style) => {
      const made = typeof patch === "function" ? patch(style) : patch;
      const paint = style?.textPaint;
      if (!statesColor(made) || !paint) return made;
      return {
        ...made,
        textPaint: presentationFontColorOnPaint(paint, made.color ?? undefined, options.automaticFill) ?? undefined,
      };
    },
  );
  if (!formatted) return null;
  return presentationBlockWithCanonicalTextPaint(formatted as unknown as PresentationBlock) as unknown as PresentationTextBlock;
}

/**
 * Font Color over an object's words as a whole (the box or shape selected,
 * every SmartArt label of a graphic): the paint every character wears
 * takes the colour as {@link presentationFontColorOnPaint} does, after the
 * object's own colour has been written.
 */
export function presentationBlockWithFontColor<T extends PresentationBlock>(
  block: T,
  color: string | undefined,
  automaticFill?: string,
): T {
  if (!isTextFrame(block)) return block;
  if (framePaint(block) === null && !presentationBlockHasRunTextPaint(block)) return block;
  const spread = presentationBlockWithRunTextPaints(block) as TextFrame;
  const recolored = withEveryRun(spread, (run) =>
    run.style?.textPaint
      ? runWithPaint(run, presentationFontColorOnPaint(run.style.textPaint, color, automaticFill))
      : run,
  );
  return presentationBlockWithCanonicalTextPaint(recolored as PresentationBlock) as T;
}

/**
 * The selected objects' blocks after Home wrote a font colour on them
 * (`written`, null when their own colour already was it), with the paint
 * their words wear taking the colour too: null when nothing changed.
 */
export function presentationObjectBlocksWithFontColor(
  blocks: readonly PresentationBlock[],
  written: PresentationBlock[] | null,
  targetIds: ReadonlySet<string>,
  color: string | undefined,
  automaticFill?: string,
): PresentationBlock[] | null {
  let painted = false;
  const next = (written ?? blocks).map((block) => {
    if (!targetIds.has(block.id)) return block;
    const recolored = presentationBlockWithFontColor(block, color, automaticFill);
    if (recolored !== block) painted = true;
    return recolored;
  });
  return painted ? next : written;
}
