/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { presentationTextUsesPoints } from "./presentation/text-units-block";
import { CSS_PX_PER_POINT } from "./presentation/text-units";
import {
  blockHoldsEquationRun,
  withEquationFormatMaterialized,
  withEquationFormatSettled,
  withEquationFormatsRead,
  type EquationRunFormat,
} from "./presentation-equation-run-italic";
import type {
  PresentationBlock,
  PresentationChartBlock,
} from "../components/console/editor/native-document/documentCommands";
import {
  type OfficeChartText,
} from "./office-charts/chart-text";
import {
  officeHomeChartTextState,
  officeHomeObjectTextState,
  type OfficeHomeObjectTextState,
  type OfficeHomeTextCommand,
} from "./office-home-object-target";
import { readSmartArtBlockMarker } from "./office-smartart/model";
import { readSmartArtGroupMembers } from "./office-smartart/design-model";
import {
  applySmartArtTextFormat,
  stepOfficeFontSizePt,
  summarizeSmartArtTextFormat,
  type SmartArtNodeTextFormat,
  type SmartArtNodeTextFormatPatch,
  type SmartArtTextFormatAdapter,
  type SmartArtTextFormatSummary,
} from "./office-smartart/text-format";
import { readPresentationChartV2 } from "./presentation-chart-v2";
import {
  PRESENTATION_SMARTART_LABEL_TEXT_ADAPTER,
  type PresentationLabelledBlock,
} from "./presentation-smartart-text";
import { normalizeRuns, runsFromText, type RunStyle, type TextRun } from "./presentation-runs";
import { presentationTextRunsInRange } from "./presentation-text-clipboard";
import {
  officeChartElementText,
  officeChartTitleRangeText,
  officeChartWithElementTextCommand,
  officeChartWithTitleRangeCommand,
  type OfficeChartTextElement,
} from "./office-charts/chart-element-text";

const PX_PER_PT = CSS_PX_PER_POINT;







function storesPoints(block: PresentationLabelledBlock): boolean {
  return presentationTextUsesPoints(block as unknown as PresentationBlock);
}

/** The block with every stated size multiplied by `factor` (points <-> design px). */
function withSizesScaled<T extends PresentationLabelledBlock>(block: T, factor: number): T {
  const scale = (style: RunStyle | undefined): RunStyle | undefined =>
    style && typeof style.fontSize === "number"
      ? { ...style, fontSize: Math.round(style.fontSize * factor * 100) / 100 }
      : style;
  const run = (entry: TextRun): TextRun => (entry.style ? { ...entry, style: scale(entry.style) } : entry);
  const next = { ...block } as T & Record<string, unknown>;
  const style = (block as { style?: RunStyle }).style;
  if (style) next.style = scale(style);
  const runs = (block as { runs?: TextRun[] }).runs;
  if (Array.isArray(runs)) next.runs = runs.map(run);
  const paragraphs = (block as { paragraphs?: unknown }).paragraphs;
  if (Array.isArray(paragraphs)) {
    next.paragraphs = paragraphs.map((paragraph) => {
      const record = paragraph as Record<string, unknown>;
      return Array.isArray(record.runs) ? { ...record, runs: (record.runs as TextRun[]).map(run) } : record;
    });
  }
  return next;
}

/** Point-sized blocks in design px, for this module's reads and writes. */
function inDesignPx<T extends PresentationLabelledBlock>(blocks: readonly T[]): T[] {
  return blocks.map((block) => (storesPoints(block) ? withSizesScaled(block, PX_PER_PT) : block));
}

export interface PresentationHomeObjectTargets {
  /** Blocks whose words Home formats: text blocks and shape blocks. */
  readonly text: PresentationLabelledBlock[];
  /** Charts whose text Home formats. */
  readonly charts: PresentationChartBlock[];
  /**
   * What Home formats on those charts: the whole chart's text (the chart
   * area is selected), or the one element in edit -- the title being typed
   * (`officeChartTitleEditElement`). Absent is the chart.
   */
  readonly chartElement?: OfficeChartTextElement;
  /** The characters selected inside the chart title being typed: Home
   * formats them alone (`officeChartTitleEditRange`). */
  readonly chartTitleRange?: { start: number; end: number };
}

/**
 * The objects Home formats for a selection: every selected text or shape
 * block, every member of a SmartArt graphic one of them belongs to, and
 * every selected chart. Pictures, tables and media carry nothing Home's Font
 * group formats.
 */
export function presentationHomeObjectTargets(
  blocks: readonly PresentationBlock[],
  selectedIds: readonly string[],
  options: {
    /**
     * With ONE shape of a SmartArt graphic selected, PowerPoint formats that
     * shape's words only; the whole graphic's take the command when the
     * graphic itself is selected (its frame pressed, or Esc from a shape).
     * `item` names the selected shape: its own members (its label, its
     * picture) are the targets and the graphic is not expanded.
     */
    readonly item?: PresentationBlock | null;
    /** The chart element in edit, when one is (`officeChartTitleEditElement`). */
    readonly chartElement?: OfficeChartTextElement;
    readonly chartTitleRange?: { start: number; end: number } | null;
  } = {},
): PresentationHomeObjectTargets {
  const wanted = new Set(selectedIds);
  const item = options.item ? readSmartArtBlockMarker(options.item) : null;
  if (item?.itemId) {
    wanted.clear();
    for (const block of blocks) {
      const marker = readSmartArtBlockMarker(block);
      if (marker && marker.groupId === item.groupId && marker.itemId === item.itemId) wanted.add(block.id);
    }
  } else {
    for (const block of blocks) {
      if (!wanted.has(block.id) || !readSmartArtBlockMarker(block)) continue;
      const group = readSmartArtGroupMembers(blocks, block);
      for (const member of group?.members ?? []) wanted.add(member.id);
    }
  }
  const text: PresentationLabelledBlock[] = [];
  const charts: PresentationChartBlock[] = [];
  for (const block of blocks) {
    if (!wanted.has(block.id)) continue;
    if (block.type === "text" || block.type === "shape") {
      // A node box of a SmartArt graphic carries no words; its label does.
      if (readSmartArtBlockMarker(block)?.role === "node" && block.type === "shape") continue;
      if (readSmartArtBlockMarker(block)?.role === "connector") continue;
      text.push(block as PresentationLabelledBlock);
    } else if (block.type === "chart") {
      charts.push(block as PresentationChartBlock);
    }
  }
  return {
    text,
    charts,
    ...(options.chartElement && options.chartElement !== "chart" ? { chartElement: options.chartElement } : {}),
    ...(options.chartTitleRange ? { chartTitleRange: options.chartTitleRange } : {}),
  };
}

type LooseStyle = Record<string, unknown>;

function styleOf(block: PresentationLabelledBlock): LooseStyle {
  const style = (block as { style?: unknown }).style;
  return style && typeof style === "object" ? (style as LooseStyle) : {};
}

/**
 * Sizes kept to the hundredth: an imported box's points are read and written
 * here in design px, and whole px turned 32 pt (42.67 px) into 32.25.
 */
const RUN_SIZES = { fontSizePrecision: "hundredth" } as const;

/** Every run of the block, flat or inside its paragraphs, in order. */
function blockRuns(block: PresentationLabelledBlock): TextRun[] {
  const paragraphs = (block as { paragraphs?: unknown }).paragraphs;
  if (Array.isArray(paragraphs)) {
    const fromParagraphs = paragraphs.flatMap((paragraph) => {
      const runs = normalizeRuns((paragraph as { runs?: unknown })?.runs, RUN_SIZES);
      const text = typeof (paragraph as { text?: unknown })?.text === "string" ? (paragraph as { text: string }).text : "";
      return runs ?? runsFromText(text);
    });
    if (fromParagraphs.length) return fromParagraphs;
  }
  return normalizeRuns((block as { runs?: TextRun[] }).runs, RUN_SIZES) ?? runsFromText((block as { text?: string }).text ?? "");
}

/** The block with every run (flat and in its paragraphs) mapped through `map`. */
function withRunsMapped(
  block: PresentationLabelledBlock,
  map: (run: TextRun, index: number) => TextRun,
): PresentationLabelledBlock {
  const next = { ...block } as PresentationLabelledBlock & Record<string, unknown>;
  let index = 0;
  const paragraphs = (block as { paragraphs?: unknown }).paragraphs;
  if (Array.isArray(paragraphs) && paragraphs.length) {
    next.paragraphs = paragraphs.map((paragraph) => {
      const record = paragraph as Record<string, unknown>;
      const runs = normalizeRuns(record.runs, RUN_SIZES) ?? runsFromText(typeof record.text === "string" ? record.text : "");
      const mapped = runs.map((run) => map(run, index++));
      const normalized = normalizeRuns(mapped, RUN_SIZES);
      const out = { ...record };
      if (normalized) out.runs = normalized;
      else delete out.runs;
      return out;
    });
    return next;
  }
  const runs = normalizeRuns((block as { runs?: TextRun[] }).runs, RUN_SIZES) ?? runsFromText((block as { text?: string }).text ?? "");
  const normalized = normalizeRuns(runs.map((run) => map(run, index++)), RUN_SIZES);
  if (normalized) next.runs = normalized;
  else delete next.runs;
  return next;
}

/** One run of one block, with the format it is drawn with (style under run). */
interface RunLeaf {
  readonly block: PresentationLabelledBlock;
  readonly index: number;
  readonly format: SmartArtNodeTextFormat;
  readonly patch?: SmartArtNodeTextFormatPatch;
}

/**
 * The size a block's words draw at when neither the run nor the block states
 * one (the master's, the placeholder's, a shape label's 18 pt), in points;
 * the host knows its masters, so it supplies this.
 */
export type PresentationHomeFallbackSizePt = (block: PresentationLabelledBlock) => number | undefined;

function runFormat(
  block: PresentationLabelledBlock,
  run: TextRun,
  fallbackSizePt?: PresentationHomeFallbackSizePt,
): SmartArtNodeTextFormat {
  const style = styleOf(block);
  const own: RunStyle = run.style ?? {};
  const fallbackPx = fallbackSizePt ? (() => { const pt = fallbackSizePt(block); return pt === undefined ? undefined : pt * PX_PER_PT; })() : undefined;
  const size = typeof own.fontSize === "number" ? own.fontSize : typeof style.fontSize === "number" ? style.fontSize : fallbackPx;
  const family = typeof own.fontFamily === "string" && own.fontFamily ? own.fontFamily : typeof style.fontFamily === "string" && style.fontFamily ? style.fontFamily : undefined;
  const color = typeof own.color === "string" && own.color ? own.color : typeof style.color === "string" && style.color ? style.color : undefined;
  return {
    ...(family ? { fontFamily: family } : {}),
    ...(size === undefined ? {} : { fontSizePt: Math.round((size / PX_PER_PT) * 100) / 100 }),
    bold: own.bold ?? style.bold === true,
    italic: own.italic ?? style.italic === true,
    ...((own.underline ?? style.underline === true) ? { underline: "single" as const } : {}),
    strike: own.strike ?? style.strike === true,
    ...(color ? { color: color.toLowerCase() } : {}),
    ...(style.align === "left" || style.align === "center" || style.align === "right" || style.align === "justify"
      ? { align: style.align }
      : {}),
  };
}

function leavesOf(
  blocks: readonly PresentationLabelledBlock[],
  fallbackSizePt?: PresentationHomeFallbackSizePt,
): RunLeaf[] {
  return blocks.flatMap((block) =>
    blockRuns(block).map((run, index) => ({ block, index, format: runFormat(block, run, fallbackSizePt) })),
  );
}

const LEAF_ADAPTER: SmartArtTextFormatAdapter<RunLeaf> = {
  read: (leaf) => leaf.format,
  write: (leaf, patch) => ({ ...leaf, patch }),
};

/** What Home shows for the selected objects' words, "mixed" across runs and objects. */
export function summarizePresentationHomeObjectText(
  blocks: readonly PresentationLabelledBlock[],
  fallbackSizePt?: PresentationHomeFallbackSizePt,
): SmartArtTextFormatSummary | null {
  // An equation's Bold, Italic and Underline are its letters' (Office Math):
  // its letters read as Italic unless made upright.
  const read = blocks.some(blockHoldsEquationRun)
    ? blocks.map((block) => withEquationFormatsRead(block))
    : blocks;
  return summarizeSmartArtTextFormat(leavesOf(read, fallbackSizePt), LEAF_ADAPTER);
}

/**
 * One Home command over every selected object's words: the blocks
 * rewritten, or `null` when nothing changes. Grow / Shrink Font steps every
 * run from its own size (PowerPoint), every other command is decided over
 * all the runs (a toggle turns off only when every run has it) and written
 * on each object as a whole.
 */
export interface PresentationHomeObjectTextOptions {
  readonly fallbackSizePt?: PresentationHomeFallbackSizePt;
  /**
   * The host's own size ladder for Grow / Shrink, in the deck's px (the
   * ribbon's size presets); without it Office's point ladder is used.
   */
  readonly stepSizePx?: (currentPx: number, direction: 1 | -1) => number | undefined;
}

export function applyPresentationHomeObjectText<T extends PresentationLabelledBlock>(
  blocks: readonly T[],
  command: OfficeHomeTextCommand,
  options: PresentationHomeObjectTextOptions = {},
): T[] | null {
  // Bold, Italic and Underline over an equation change its letters (math
  // italic / upright, bold, underlined), as Office Math does:
  // `lib/presentation-equation-run-italic`.
  const kind: EquationRunFormat | null =
    command.kind === "toggle" && (command.field === "italic" || command.field === "bold")
      ? command.field
      : command.kind === "underline"
        ? "underline"
        : null;
  if (kind) {
    if (blocks.some(blockHoldsEquationRun)) {
      const materialized = blocks.map((block) => withEquationFormatMaterialized(block, kind));
      const applied = applyPresentationHomeObjectTextOfRuns(
        materialized.map((entry) => entry.block),
        command,
        options,
      );
      if (!applied) return null;
      return applied.map((block, index) =>
        block === materialized[index].block
          ? blocks[index]
          : withEquationFormatSettled(block, materialized[index].states, kind, true),
      );
    }
  }
  
  if (blocks.some(storesPoints)) {
    const viewed = inDesignPx(blocks);
    const applied = applyPresentationHomeObjectTextOfRuns(viewed, command, options);
    if (!applied) return null;
    return applied.map((block, index) =>
      block === viewed[index] ? blocks[index] : storesPoints(blocks[index]) ? withSizesScaled(block, 1 / PX_PER_PT) : block,
    );
  }
  return applyPresentationHomeObjectTextOfRuns(blocks, command, options);
}

function applyPresentationHomeObjectTextOfRuns<T extends PresentationLabelledBlock>(
  blocks: readonly T[],
  command: OfficeHomeTextCommand,
  options: PresentationHomeObjectTextOptions,
): T[] | null {
  if (command.kind === "step") {
    // An object with no words yet: Grow / Shrink sets the size its text
    // will be typed at (PowerPoint keeps it on the end-of-paragraph run),
    // stepped from the size it would draw at.
    const wordless = blocks.filter((block) => wordRuns([block]).length === 0);
    if (wordless.length) {
      const worded = blocks.filter((block) => !wordless.includes(block));
      const steppedWorded = worded.length ? applyPresentationHomeObjectText(worded, command, options) : null;
      let changed = steppedWorded !== null;
      const byBlock = new Map<PresentationLabelledBlock, T>();
      worded.forEach((block, index) => byBlock.set(block, (steppedWorded?.[index] ?? block) as T));
      const next = blocks.map((block) => {
        if (byBlock.has(block)) return byBlock.get(block)!;
        const style = styleOf(block);
        const currentPx = typeof style.fontSize === "number"
          ? style.fontSize
          : (() => { const pt = options.fallbackSizePt?.(block); return pt === undefined ? undefined : pt * PX_PER_PT; })();
        if (currentPx === undefined) return block;
        const nextPx = options.stepSizePx
          ? options.stepSizePx(currentPx, command.direction)
          : (() => {
              const pt = stepOfficeFontSizePt(currentPx / PX_PER_PT, command.direction);
              return pt === null ? undefined : Math.round(pt * PX_PER_PT * 100) / 100;
            })();
        if (nextPx === undefined || nextPx === style.fontSize) return block;
        changed = true;
        return { ...block, style: { ...style, fontSize: nextPx } } as T;
      });
      return changed ? next : null;
    }
  }
  const leaves = leavesOf(blocks, options.fallbackSizePt);
  if (leaves.length === 0) return null;
  const written = applySmartArtTextFormat(leaves, command, LEAF_ADAPTER);
  if (!written) return null;
  let changed = false;
  const next = blocks.map((block) => {
    const own = written.filter((leaf) => leaf.block === block && leaf.patch);
    if (own.length === 0) return block;
    if (command.kind === "step") {
      // Each run steps from its own size; the block's style stays.
      const byIndex = new Map(own.map((leaf) => [leaf.index, leaf] as const));
      const stepped = withRunsMapped(block, (run, index) => {
        const leaf = byIndex.get(index);
        if (!leaf) return run;
        const pt = leaf.patch!.fontSizePt;
        const currentPt = leaf.format.fontSizePt;
        const fontSize = options.stepSizePx && currentPt !== undefined
          ? options.stepSizePx(currentPt * PX_PER_PT, command.direction)
          : pt === undefined ? undefined : Math.round(pt * PX_PER_PT * 100) / 100;
        if (fontSize === undefined) return run;
        return { ...run, style: { ...(run.style ?? {}), fontSize } };
      });
      if (JSON.stringify(stepped) === JSON.stringify(block)) return block;
      changed = true;
      return stepped as T;
    }
    const patch = own[0].patch!;
    const result = PRESENTATION_SMARTART_LABEL_TEXT_ADAPTER.write(block, patch);
    if (result === block) return block;
    // A character format picked with the shape selected (not its text) is
    // every run's, as PowerPoint applies it: Bold, Italic, Underline,
    // Strikethrough, Font and Font Color set on every run of the shape --
    // a run's own value included (an explicit italic:false run turns
    // italic) -- and on the frame. The size keeps the frame's value showing
    // by taking each run's own off, as before.
    const resultStyle = ((result as { style?: Record<string, unknown> }).style ?? {}) as Record<string, unknown>;
    const patchValues = patch as unknown as Record<string, unknown>;
    // Turned off, the frame states nothing; each run states the off value
    // (PowerPoint writes b="0" on every run), so a master's bold title
    // does not show through.
    const stated = (["fontFamily", "bold", "italic", "underline", "strike", "color"] as const)
      .filter((key) => Object.hasOwn(patchValues, key))
      .map((key) => [key, Object.hasOwn(resultStyle, key) ? resultStyle[key] : patchValues[key]] as const)
      .filter(([, value]) => value !== undefined);
    const withRows = Array.isArray((result as { paragraphs?: unknown }).paragraphs);
    const sizeOff = withRows && Object.hasOwn(patchValues, "fontSizePt");
    const restate = (run: TextRun, sizeToo: boolean): TextRun => {
      const style: Record<string, unknown> = { ...(run.style ?? {}) };
      if (sizeToo) delete style.fontSize;
      for (const [key, value] of stated) style[key] = value;
      return Object.keys(style).length ? ({ ...run, style } as TextRun) : { text: run.text };
    };
    let cleaned = result as PresentationLabelledBlock;
    if (stated.length || sizeOff) {
      cleaned = withRunsMapped(cleaned, (run) => restate(run, sizeOff));
      // A block with rows also keeps a flat copy of its runs: the same marks.
      const flat = (cleaned as { runs?: unknown }).runs;
      if (withRows && Array.isArray(flat) && stated.length) {
        cleaned = { ...cleaned, runs: (flat as TextRun[]).map((run) => restate(run, false)) } as PresentationLabelledBlock;
      }
    }
    changed = true;
    return cleaned as T;
  });
  return changed ? next : null;
}

/**
 * A mark of the runs themselves that Home sets on a selected object's words:
 * a baseline shift or a highlight. PowerPoint formats all of an object's
 * text when its border is selected, these included (Microsoft Support,
 * "Format text as superscript or subscript"; "Highlight text").
 */
export type PresentationHomeObjectRunMark =
  | { readonly kind: "script"; readonly script: "superscript" | "subscript" }
  /** The Font dialog's OK: the baseline every run gets, stated outright (PPT-B425). */
  | { readonly kind: "scriptSet"; readonly superscript: boolean; readonly subscript: boolean }
  /** `undefined` is No Color. */
  | { readonly kind: "highlight"; readonly color: string | undefined };

function wordRuns(blocks: readonly PresentationLabelledBlock[]): TextRun[] {
  return blocks.flatMap(blockRuns).filter((run) => run.text !== "");
}

/**
 * The script / highlight Home shows for the selected objects' words: a
 * script reads on only when every run has it, `"mixed"` when some do; the
 * highlight is the runs' one colour, `"mixed"` when they differ.
 */
export function presentationHomeObjectRunMarks(blocks: readonly PresentationLabelledBlock[]): {
  readonly superscript: boolean | "mixed";
  readonly subscript: boolean | "mixed";
  readonly highlight: string | undefined | "mixed";
} {
  const runs = wordRuns(blocks);
  const agree = (read: (run: TextRun) => boolean): boolean | "mixed" => {
    if (runs.length === 0) return false;
    const on = runs.filter(read).length;
    return on === runs.length ? true : on === 0 ? false : "mixed";
  };
  const highlights = new Set(runs.map((run) => run.style?.highlight?.toLowerCase()));
  return {
    superscript: agree((run) => run.style?.superscript === true),
    subscript: agree((run) => run.style?.subscript === true),
    highlight: highlights.size > 1 ? "mixed" : [...highlights][0],
  };
}

/**
 * A script or a highlight on every run of every selected object: a script is
 * decided over all of them (it turns off only when every run already has it,
 * and setting one baseline clears the other); a highlight is set (or, with
 * No Color, taken off) everywhere. The blocks rewritten, or `null` when
 * nothing changes.
 */
export function applyPresentationHomeObjectRunMark<T extends PresentationLabelledBlock>(
  blocks: readonly T[],
  mark: PresentationHomeObjectRunMark,
): T[] | null {
  const runs = wordRuns(blocks);
  if (runs.length === 0) return null;
  let patch: Partial<Record<keyof RunStyle, unknown>>;
  if (mark.kind === "script") {
    const on = !runs.every((run) => run.style?.[mark.script] === true);
    patch = mark.script === "superscript"
      ? { superscript: on ? true : undefined, ...(on ? { subscript: undefined } : {}) }
      : { subscript: on ? true : undefined, ...(on ? { superscript: undefined } : {}) };
  } else if (mark.kind === "scriptSet") {
    patch = {
      superscript: mark.superscript ? true : undefined,
      subscript: mark.subscript && !mark.superscript ? true : undefined,
    };
  } else {
    patch = { highlight: mark.color ? mark.color.toLowerCase() : undefined };
  }
  let changed = false;
  const next = blocks.map((block) => {
    const mapped = withRunsMapped(block, (run) => {
      if (run.text === "") return run;
      const style = { ...(run.style ?? {}) } as Record<string, unknown>;
      for (const [key, value] of Object.entries(patch)) {
        if (value === undefined) delete style[key];
        else style[key] = value;
      }
      return Object.keys(style).length ? { ...run, style: style as RunStyle } : { text: run.text };
    });
    if (JSON.stringify(mapped) === JSON.stringify(block)) return block;
    changed = true;
    return mapped as T;
  });
  return changed ? next : null;
}

/**
 * Home's Align Text (top / middle / bottom) on every selected object: it
 * anchors the text inside each object (`a:bodyPr@anchor`), the object stays
 * where it is. A text box's default anchor is top and is left unstated; a
 * shape's absent anchor draws centred here, so a shape states every anchor.
 */
export function applyPresentationHomeObjectAnchor<T extends PresentationLabelledBlock>(
  blocks: readonly T[],
  anchor: "top" | "middle" | "bottom",
): T[] | null {
  let changed = false;
  const next = blocks.map((block) => {
    const style = { ...styleOf(block) };
    const value = block.type === "shape" ? anchor : anchor === "top" ? undefined : anchor;
    if (style.verticalAlign === value || (value === undefined && style.verticalAlign === undefined)) return block;
    if (value === undefined) delete style.verticalAlign;
    else style.verticalAlign = value;
    changed = true;
    return { ...block, style } as T;
  });
  return changed ? next : null;
}

/** A selected chart's text as Home shows it: the chart's own, or the
 * element in edit's (its own over the chart's). */
export function presentationChartHomeText(
  block: PresentationChartBlock,
  element: OfficeChartTextElement = "chart",
  titleRange?: { start: number; end: number },
): OfficeChartText | undefined {
  const attrs = readPresentationChartV2(block.chartV2);
  if (!attrs) return undefined;
  if (element === "title" && titleRange) return officeChartTitleRangeText(attrs, titleRange);
  return officeChartElementText(attrs, element);
}

/** The block with a Home command applied to the chart's text, or to the one
 * element in edit (Office formats the element selected, not the chart). */
export function applyPresentationChartHomeText(
  block: PresentationChartBlock,
  command: OfficeHomeTextCommand,
  element: OfficeChartTextElement = "chart",
  titleRange?: { start: number; end: number },
): PresentationChartBlock | null {
  const attrs = readPresentationChartV2(block.chartV2);
  if (!attrs) return null;
  const next =
    element === "title" && titleRange
      ? officeChartWithTitleRangeCommand(attrs, titleRange, command)
      : officeChartWithElementTextCommand(attrs, element, command);
  return next ? { ...block, chartV2: next } : null;
}

/**
 * The words selected in an object's text being typed in, as a block of their
 * own: Home reads their format, as its commands format just them. A caret
 * reads the character its typing continues (the one before it, else the one
 * after). Null when the range holds no words (a bare paragraph break).
 */
export function presentationHomeObjectWordsSlice<T extends PresentationLabelledBlock>(
  block: T,
  start: number,
  end: number,
): T | null {
  const from = Math.min(start, end);
  const to = Math.max(start, end);
  const words = (a: number, b: number) =>
    presentationTextRunsInRange(block, a, b).filter((run) => run.text.replace(/\n/g, "") !== "");
  const runs =
    from < to
      ? words(from, to)
      : (from > 0 ? words(from - 1, from) : []).length
        ? words(from - 1, from)
        : words(from, from + 1);
  if (runs.length === 0) return null;
  const { paragraphs: _paragraphs, ...rest } = block as T & { paragraphs?: unknown };
  void _paragraphs;
  return { ...rest, text: runs.map((run) => run.text).join(""), runs } as unknown as T;
}

/** Home's state for the selected objects, in the shared contract's shape. */
export function presentationHomeObjectTextState(
  targets: PresentationHomeObjectTargets,
  fallbackSizePt?: PresentationHomeFallbackSizePt,
): OfficeHomeObjectTextState {
  if (targets.text.length === 0 && targets.charts.length === 1) {
    return officeHomeChartTextState(
      presentationChartHomeText(targets.charts[0], targets.chartElement, targets.chartTitleRange),
    );
  }
  return officeHomeObjectTextState(
    summarizePresentationHomeObjectText(inDesignPx(targets.text), fallbackSizePt),
  );
}

/** Grow / Shrink of a chart's one size along Office's ladder. */
export { stepOfficeFontSizePt };
