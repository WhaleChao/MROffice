/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { CSS_PX_PER_POINT } from "./presentation/text-units";
import type { CSSProperties } from "react";

import {
  CELL_CHARACTER_FIELDS,
  CELL_RUNS_CLIPBOARD_TYPE,
  cellRunsFromClipboardHtml,
  cellTextRunCss,
  mergeCellTextRuns,
  parseCellRunsClipboard,
} from "./presentation-cell-rich-text";
import {
  normalizeRunStyle,
  normalizeRuns,
  runsFromText,
  splitRun,
  type RunStyle,
  type TextRun,
} from "./presentation-point-runs";
import { normalizeParagraphsWithRuns, ownsParagraphRichLayout } from "./presentation-text-buttons";

/** The clipboard flavour a deck text selection's runs travel in. */
export const PRESENTATION_TEXT_RUNS_CLIPBOARD_TYPE = "application/x-mrai-presentation-text-runs";

const MAX_CLIPBOARD_RUNS = 400;
const MAX_CLIPBOARD_JSON = 400_000;

/** The character keys a paste states on the pasted words, all of them: a key
 * the copied run does not set is cleared, so the words take the destination
 * box's format there and not the run the caret happened to be in. */
const PASTED_CHARACTER_KEYS = [
  "bold",
  "italic",
  "underline",
  "underlineStyle",
  "strike",
  "caps",
  "superscript",
  "subscript",
  "proofingLanguage",
  "alternateLanguage",
  "noProof",
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
  "href",
  "linkTooltip",
  "linkColor",
] as const satisfies readonly (keyof RunStyle)[];

type RunsSource = {
  readonly text?: unknown;
  readonly runs?: unknown;
  readonly paragraphs?: unknown;
};

function sliceRuns(runs: readonly TextRun[], from: number, to: number): TextRun[] {
  const out: TextRun[] = [];
  let offset = 0;
  for (const run of runs) {
    const start = Math.max(from - offset, 0);
    const end = Math.min(to - offset, run.text.length);
    if (start < end) out.push(splitRun(splitRun(run, start)[1], end - start)[0]);
    offset += run.text.length;
  }
  return out;
}

/** Sizes travel in design px, as an object copy's do (an imported box
 * states points, `presentationTextUsesPoints`). */
function withSizeScaled(run: TextRun, factor: number): TextRun {
  if (factor === 1 || typeof run.style?.fontSize !== "number") return run;
  return { ...run, style: { ...run.style, fontSize: Math.round(run.style.fontSize * factor * 100) / 100 } };
}

/**
 * The runs of `[start, end)` of a text box's (or a shape label's) words, as
 * its body's selection offsets address them: paragraphs joined by an
 * unformatted "\n", each character with the formatting set on it, sizes in
 * design px (`sizesInPoints`: the block states points).
 */
export function presentationTextRunsInRange(
  block: RunsSource,
  start: number,
  end: number,
  options: { readonly sizesInPoints?: boolean } = {},
): TextRun[] {
  // The words as the body addresses them (`presentationEditableText`): the
  // paragraphs when they own the layout, else the flat runs.
  const paragraphs = normalizeParagraphsWithRuns(block.paragraphs);
  const flatText = typeof block.text === "string" ? block.text : "";
  const runs: TextRun[] =
    paragraphs && ownsParagraphRichLayout(paragraphs)
      ? paragraphs.flatMap((paragraph, index) => [
          ...(index > 0 ? [{ text: "\n" }] : []),
          ...(normalizeRuns(paragraph.runs) ?? runsFromText(paragraph.text)),
        ])
      : (normalizeRuns(block.runs) ?? runsFromText(flatText));
  const total = runs.reduce((sum, run) => sum + run.text.length, 0);
  const from = Math.min(Math.max(Math.round(start), 0), total);
  const to = Math.min(Math.max(Math.round(end), 0), total);
  if (to <= from) return [];
  return sliceRuns(runs, from, to)
    .filter((run) => run.text)
    .map((run) => {
      const style = normalizeRunStyle(run.style);
      return withSizeScaled(style ? { text: run.text, style } : { text: run.text }, options.sizesInPoints ? CSS_PX_PER_POINT : 1);
    });
}







export function presentationRunsWithFieldsShown(
  runs: readonly TextRun[],
  shown: (field: NonNullable<TextRun["field"]>) => string | undefined,
): TextRun[] {
  return runs.map((run) => {
    if (!run.field) return run;
    const text = shown(run.field);
    return text && text !== run.text ? { ...run, text } : run;
  });
}

function cssText(css: CSSProperties): string {
  return Object.entries(css)
    .filter(([, value]) => value !== undefined && value !== null && value !== "")
    .map(([key, value]) => `${key.replace(/[A-Z]/g, (letter) => `-${letter.toLowerCase()}`)}:${String(value)}`)
    .join(";");
}

/**
 * Write a deck text selection's runs to a copy (or cut): the plain text,
 * HTML that states each run's own format for other programs, and the runs.
 * False when there is nothing to copy.
 */
export function writePresentationTextRunsClipboard(data: DataTransfer, runs: readonly TextRun[]): boolean {
  const text = runs.map((run) => run.text).join("");
  if (!text) return false;
  // A soft line break is a line break to every other reader.
  data.setData("text/plain", text.replace(/\v/g, "\n"));
  data.setData(PRESENTATION_TEXT_RUNS_CLIPBOARD_TYPE, JSON.stringify(runs));
  const escape = (value: string) => value.replace(/[&<>"]/g, (char) => `&#${char.charCodeAt(0)};`);
  data.setData(
    "text/html",
    runs
      .map((run) => {
        const style = run.style ? cssText(cellTextRunCss(run.style)) : "";
        return `<span${style ? ` style="${escape(style)}"` : ""}>${escape(run.text).replace(/[\n\v]/g, "<br>")}</span>`;
      })
      .join(""),
  );
  return true;
}

/**
 * The runs a text box's or a shape label's copy wrote, when they are valid
 * and spell exactly `plain` (the same copy).
 */
export function readPresentationTextRunsFlavour(data: DataTransfer, plain: string): TextRun[] | null {
  const expected = plain.replace(/\r\n?/g, "\n");
  const raw = data.getData(PRESENTATION_TEXT_RUNS_CLIPBOARD_TYPE);
  if (raw && raw.length <= MAX_CLIPBOARD_JSON) {
    let value: unknown;
    try {
      value = JSON.parse(raw);
    } catch {
      value = null;
    }
    if (
      Array.isArray(value) &&
      value.length > 0 &&
      value.length <= MAX_CLIPBOARD_RUNS &&
      value.every((run) => run && typeof run === "object" && typeof (run as TextRun).text === "string")
    ) {
      const runs = (value as TextRun[]).map((run) => {
        const style = normalizeRunStyle(run.style);
        return style ? { text: run.text, style } : { text: run.text };
      });
      if (runs.map((run) => run.text).join("").replace(/\v/g, "\n") === expected) return runs;
    }
  }
  return null;
}

/**
 * The runs a deck copy wrote -- a text box's or a shape label's, else a table
 * cell's -- when they are valid and spell exactly `plain` (the same copy).
 */
export function readPresentationTextRunsClipboard(data: DataTransfer, plain: string): TextRun[] | null {
  return (
    readPresentationTextRunsFlavour(data, plain) ??
    parseCellRunsClipboard(data.getData(CELL_RUNS_CLIPBOARD_TYPE), plain)
  );
}

/**
 * The formats a paste into a text box or a shape's label states on the
 * pasted words, by their offsets in the pasted text: a deck copy's runs, each
 * with every character key (`presentationPastedRunPatch`, its size in the
 * destination's unit), else what another
 * program's HTML states (bold, italic, underline, strikethrough). Null when
 * the paste states none.
 */
export function presentationPastedRunFormats(
  data: DataTransfer,
  plain: string,
  /** The destination states its sizes in points (an imported box). */
  options: { readonly sizesInPoints?: boolean } = {},
): { from: number; to: number; style: Partial<RunStyle> }[] | null {
  const own = readPresentationTextRunsClipboard(data, plain);
  const runs = own ?? cellRunsFromClipboardHtml(data.getData("text/html"), plain);
  if (!runs) return null;
  const out: { from: number; to: number; style: Partial<RunStyle> }[] = [];
  let offset = 0;
  for (const run of runs) {
    const length = run.text.replace(/\r\n?/g, "\n").length;
    // A paragraph break carries no character format.
    const style = own
      ? /^[\n\v]+$/.test(run.text)
        ? undefined
        : presentationPastedRunPatch(withSizeScaled(run, options.sizesInPoints ? 1 / CSS_PX_PER_POINT : 1).style)
      : run.style;
    if (style && length > 0) out.push({ from: offset, to: offset + length, style });
    offset += length;
  }
  return out.length ? out : null;
}

/**
 * The format a pasted run states on its words: every character key, the
 * run's own value or cleared (a cell run's `false` flags clear too).
 */
export function presentationPastedRunPatch(style: RunStyle | undefined): Partial<RunStyle> {
  const patch: Record<string, unknown> = {};
  for (const key of PASTED_CHARACTER_KEYS) {
    const value = style?.[key];
    patch[key] = value === undefined || value === false ? undefined : value;
  }
  return patch as Partial<RunStyle>;
}

/** A deck text copy's runs as a table cell takes them: the character
 * formats a cell has, adjacent equal runs merged. */
export function presentationTextRunsForCell(runs: readonly TextRun[]): TextRun[] {
  return mergeCellTextRuns(
    runs.map((run) => {
      const style = Object.fromEntries(
        Object.entries(run.style ?? {}).filter(([key]) => CELL_CHARACTER_FIELDS.has(key)),
      ) as RunStyle;
      return Object.keys(style).length ? { text: run.text, style } : { text: run.text };
    }),
  );
}
