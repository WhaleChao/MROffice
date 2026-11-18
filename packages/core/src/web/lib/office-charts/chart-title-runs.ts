/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A chart title's runs: the title's words in pieces, each with the text
 * format its characters carry over the title's own (`chartTitleStyle.text`)
 * and the chart's (`format.text`) -- the `a:r` runs of `c:title/c:tx/c:rich`.
 * Office formats the characters selected inside a title (Microsoft Support,
 * "Change the font and size of text in a chart": "select the individual
 * characters you want to format"); the runs are how a part keeps that. The
 * runs always spell the title exactly; a title with one plain run states no
 * runs at all.
 */
import {
  applyOfficeChartTextCommand,
  normalizeOfficeChartText,
  OFFICE_CHART_SUPERSCRIPT_SCALE,
  officeChartTextEChartsStyle,
  type OfficeChartText,
  type OfficeChartTextCommand,
} from "./chart-text";

export interface ChartTitleRun {
  readonly text: string;
  /** The run's own text format over the title's; absent is the title's. */
  readonly style?: OfficeChartText;
}

function sameStyle(a: OfficeChartText | undefined, b: OfficeChartText | undefined): boolean {
  return JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
}

/** Adjacent runs of one style joined; empty runs dropped; no runs when one plain run remains. */
export function mergeChartTitleRuns(runs: readonly ChartTitleRun[]): ChartTitleRun[] | undefined {
  const merged: ChartTitleRun[] = [];
  for (const run of runs) {
    if (!run.text) continue;
    const last = merged[merged.length - 1];
    if (last && sameStyle(last.style, run.style)) {
      merged[merged.length - 1] = { ...last, text: last.text + run.text };
    } else {
      merged.push(run.style ? { text: run.text, style: run.style } : { text: run.text });
    }
  }
  if (merged.length === 0) return undefined;
  if (merged.length === 1 && !merged[0].style) return undefined;
  return merged;
}

/** The words the runs spell. */
export function chartTitleRunsText(runs: readonly ChartTitleRun[] | undefined): string {
  return (runs ?? []).map((run) => run.text).join("");
}

/**
 * `value` as a title's runs, or undefined when it is not one: a list of
 * runs with their text and a valid style, spelling `title` exactly.
 */
export function normalizeChartTitleRuns(value: unknown, title: string): ChartTitleRun[] | undefined {
  if (!Array.isArray(value) || value.length === 0) return undefined;
  const runs: ChartTitleRun[] = [];
  for (const raw of value) {
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
    const record = raw as Record<string, unknown>;
    if (typeof record.text !== "string") return undefined;
    if (Object.keys(record).some((key) => key !== "text" && key !== "style")) return undefined;
    if (record.style !== undefined) {
      const style = normalizeOfficeChartText(record.style);
      if (!style) return undefined;
      runs.push({ text: record.text, style });
    } else {
      runs.push({ text: record.text });
    }
  }
  if (chartTitleRunsText(runs) !== title) return undefined;
  return mergeChartTitleRuns(runs);
}

/**
 * The runs after the title's words changed (typed over): the characters
 * both titles share at the start and at the end keep their runs, and what
 * was typed between takes the style of the run it was typed into (the run
 * before the change, as Office extends a run when typing within it).
 */
export function fitChartTitleRuns(
  runs: readonly ChartTitleRun[] | undefined,
  title: string,
): ChartTitleRun[] | undefined {
  if (!runs?.length) return undefined;
  const old = chartTitleRunsText(runs);
  if (old === title) return mergeChartTitleRuns(runs);
  let prefix = 0;
  while (prefix < old.length && prefix < title.length && old[prefix] === title[prefix]) prefix += 1;
  let suffix = 0;
  while (
    suffix < old.length - prefix &&
    suffix < title.length - prefix &&
    old[old.length - 1 - suffix] === title[title.length - 1 - suffix]
  )
    suffix += 1;
  const inserted = title.slice(prefix, title.length - suffix);
  const pieces: ChartTitleRun[] = [];
  let at = 0;
  let insertedStyle: OfficeChartText | undefined;
  let insertedPlaced = false;
  for (const run of runs) {
    const start = at;
    const end = at + run.text.length;
    at = end;
    // The run's part before the change.
    if (start < prefix) {
      const keep = run.text.slice(0, Math.min(run.text.length, prefix - start));
      if (keep) pieces.push({ ...run, text: keep });
      insertedStyle = run.style;
    }
    if (!insertedPlaced && end >= prefix) {
      pieces.push(insertedStyle || run.style ? { text: inserted, style: insertedStyle ?? run.style } : { text: inserted });
      insertedPlaced = true;
    }
    // The run's part after the change.
    const tailStart = old.length - suffix;
    if (end > tailStart) {
      const from = Math.max(0, tailStart - start);
      const keep = run.text.slice(from);
      if (keep) pieces.push({ ...run, text: keep });
    }
  }
  if (!insertedPlaced) pieces.push({ text: inserted });
  const fitted = mergeChartTitleRuns(pieces.map((run) => (run.style ? run : { text: run.text })));
  return fitted && chartTitleRunsText(fitted) === title ? fitted : undefined;
}

/**
 * A title typed in place, from the words as typed: each piece keeps the
 * format of the model run it was typed in (`run`, an index into the runs
 * the edit began with), so a run's format follows its characters exactly --
 * typed into, split by a new line, or deleted. `limit` caps the title's
 * characters (the model's).
 */
export function chartTitleRunsFromPieces(
  title: string,
  runs: readonly ChartTitleRun[] | undefined,
  pieces: readonly { text: string; run: number }[],
  limit = Infinity,
): { title: string; titleRuns?: ChartTitleRun[] } {
  const model = chartTitleRunsOf(title, runs);
  const typed: ChartTitleRun[] = [];
  let left = limit;
  for (const piece of pieces) {
    if (left <= 0) break;
    const text = piece.text.replace(/\r\n?/g, "\n").slice(0, left);
    left -= text.length;
    if (!text) continue;
    const style = model[piece.run]?.style;
    typed.push(style ? { text, style } : { text });
  }
  const words = typed.map((run) => run.text).join("");
  const titleRuns = mergeChartTitleRuns(typed);
  return titleRuns ? { title: words, titleRuns } : { title: words };
}

/** The runs of a title, one plain run when it states none. */
export function chartTitleRunsOf(title: string, runs: readonly ChartTitleRun[] | undefined): ChartTitleRun[] {
  const fitted = runs?.length && chartTitleRunsText(runs) === title ? [...runs] : [{ text: title }];
  return fitted;
}

/** The style the characters at `offset` carry (the run they sit in). */
export function chartTitleRunStyleAt(
  title: string,
  runs: readonly ChartTitleRun[] | undefined,
  offset: number,
): OfficeChartText | undefined {
  let at = 0;
  for (const run of chartTitleRunsOf(title, runs)) {
    if (offset < at + run.text.length) return run.style;
    at += run.text.length;
  }
  const last = chartTitleRunsOf(title, runs).at(-1);
  return last?.style;
}

/**
 * The runs after a Font command on the characters [start, end): the runs
 * are split at the range, and the command applied to each run inside it
 * against what that run shows (`shown(run)`), so Bold on a mixed range turns
 * every run bold the way Office does. Null when nothing changes.
 */
export function chartTitleRunsWithCommand(
  title: string,
  runs: readonly ChartTitleRun[] | undefined,
  start: number,
  end: number,
  command: OfficeChartTextCommand,
  inherited: OfficeChartText | undefined,
): ChartTitleRun[] | undefined | null {
  const from = Math.max(0, Math.min(start, end));
  const to = Math.min(title.length, Math.max(start, end));
  if (to <= from) return null;
  // A toggle acts on the whole selection: off only when every selected
  // character already has it, else on for all of them (Office's Bold on a
  // mixed selection makes it all bold).
  let resolved: OfficeChartTextCommand = command;
  if (command.kind === "toggle" || command.kind === "underline") {
    let cursor = 0;
    let allOn = true;
    for (const run of chartTitleRunsOf(title, runs)) {
      const runStart = cursor;
      const runEnd = cursor + run.text.length;
      cursor = runEnd;
      if (runEnd <= from || runStart >= to) continue;
      const shown = { ...(inherited ?? {}), ...(run.style ?? {}) } as OfficeChartText;
      const on = command.kind === "toggle" ? shown[command.field] === true : shown.underline === command.variant;
      if (!on) allOn = false;
    }
    resolved = { ...command, set: !allOn } as OfficeChartTextCommand & { set: boolean };
  }
  const pieces: ChartTitleRun[] = [];
  let at = 0;
  let changed = false;
  for (const run of chartTitleRunsOf(title, runs)) {
    const runStart = at;
    const runEnd = at + run.text.length;
    at = runEnd;
    const before = run.text.slice(0, Math.max(0, Math.min(run.text.length, from - runStart)));
    const after = run.text.slice(Math.min(run.text.length, Math.max(0, to - runStart)));
    const inside = run.text.slice(before.length, run.text.length - after.length);
    if (before) pieces.push({ ...run });
    if (before) pieces[pieces.length - 1] = { ...run, text: before };
    if (inside) {
      const shown = inherited || run.style ? { ...(inherited ?? {}), ...(run.style ?? {}) } : undefined;
      const next = applyRunCommand(run.style, shown, inherited, resolved);
      if (next !== null) changed = true;
      const style = next === null ? run.style : next;
      pieces.push(style ? { text: inside, style } : { text: inside });
    }
    if (after) pieces.push(run.style ? { text: after, style: run.style } : { text: after });
  }
  if (!changed) return null;
  return mergeChartTitleRuns(pieces);
}

function applyRunCommand(
  own: OfficeChartText | undefined,
  shown: OfficeChartText | undefined,
  inherited: OfficeChartText | undefined,
  command: OfficeChartTextCommand,
): OfficeChartText | null | undefined {
  const draft: Record<string, unknown> = { ...(own ?? {}) };
  const settle = (key: string, value: unknown) => {
    const below = (inherited as Record<string, unknown> | undefined)?.[key] ?? (typeof value === "boolean" ? false : undefined);
    if (value === undefined || value === below) delete draft[key];
    else draft[key] = value;
  };
  const set = (command as { set?: boolean }).set;
  switch (command.kind) {
    case "toggle":
      settle(command.field, set ?? (shown?.[command.field] === true ? false : true));
      break;
    case "underline":
      settle(
        "underline",
        set === undefined ? (shown?.underline === command.variant ? undefined : command.variant) : set ? command.variant : undefined,
      );
      break;
    default:
      return applyOfficeChartTextCommand(own, command);
  }
  const next = Object.keys(draft).length ? normalizeOfficeChartText(draft) : undefined;
  if (Object.keys(draft).length && !next) return null;
  return JSON.stringify(next ?? null) === JSON.stringify(own ?? null) ? null : next;
}

/**
 * Words drawn as rich pieces of the named style, every character as typed.
 * The engine's markup (`{name|words}`, zrender's `/\{([a-zA-Z0-9_]+)\|([^}]*)\}/`)
 * has no escape: a piece's words may hold `{` and `|` but end at the first
 * `}`. So a `}` is drawn between pieces, in the text's own style (a piece
 * never holds it), and the words stay the characters typed -- not their
 * full-width forms, another glyph of another width than Office draws.
 */
export function chartRichPiece(name: string, words: string): string {
  return words
    .split("}")
    .map((part) => (part ? `{${name}|${part}}` : ""))
    .join("}");
}

/**
 * Words drawn outside any piece, in a text that has rich styles, as typed:
 * only a `{` that would open a piece (`{name|...}`) is drawn as a piece of
 * no style of its own (`{_|{}`), which the engine draws in the text's style.
 */
export function chartRichLiteral(words: string): string {
  return words.replace(/\{(?=[a-zA-Z0-9_]+\|[^}]*\})/g, "{_|{}");
}

/** The characters a rich text draws, as the engine reads its markup. */
export function chartRichPlain(drawn: string): string {
  return drawn.replace(/\{[a-zA-Z0-9_]+\|([^}]*)\}/g, "$1");
}

/**
 * The title as ECharts draws rich text: `{rN|words}` pieces and a `rich`
 * map with each run's style over the title's base style, the words as typed
 * (`chartRichPiece`). Null when the runs carry no style of their own.
 */
/** One run's look over the base style, a raised or lowered run at two
 * thirds of its size against the top or the foot of the line. */
function chartRunLook(
  run: ChartTitleRun,
  base: Record<string, unknown>,
  scale: number,
): Record<string, unknown> {
  const look: Record<string, unknown> = { ...base, ...officeChartTextEChartsStyle(run.style, scale) };
  const baseline = run.style?.baseline;
  if (typeof baseline === "number" && baseline !== 0 && typeof look.fontSize === "number") {
    const size = look.fontSize as number;
    look.fontSize = size * OFFICE_CHART_SUPERSCRIPT_SCALE;
    look.verticalAlign = baseline > 0 ? "top" : "bottom";
    look.lineHeight = size * 1.2;
  }
  return look;
}








export function chartRunsRichTextOver(
  title: string,
  shown: string,
  runs: readonly ChartTitleRun[] | undefined,
  base: Record<string, unknown>,
  scale = 1,
): { text: string; rich: Record<string, Record<string, unknown>> } | null {
  const fitted = runs?.length && chartTitleRunsText(runs) === title ? runs : undefined;
  if (!fitted || !fitted.some((run) => run.style)) return null;
  const owner: number[] = [];
  fitted.forEach((run, index) => {
    for (let at = 0; at < run.text.length; at++) owner.push(index);
  });
  const pieces: { run: number; text: string }[] = [];
  const push = (run: number, character: string) => {
    const last = pieces[pieces.length - 1];
    if (last && last.run === run) last.text += character;
    else pieces.push({ run, text: character });
  };
  let at = 0;
  for (const character of shown) {
    if (character === "\n" && title[at] !== "\n") {
      // A wrap's break: the spaces it replaced are not drawn.
      while (at < title.length && /[ \t]/.test(title[at])) at++;
      push(owner[Math.min(at, owner.length - 1)] ?? 0, character);
      continue;
    }
    while (at < title.length && title[at] !== character && /[ \t]/.test(title[at])) at++;
    if (title[at] !== character) return null;
    push(owner[at], character);
    at++;
  }
  const rich: Record<string, Record<string, unknown>> = {};
  const text = pieces
    .map((piece, index) => {
      const name = `r${index}`;
      rich[name] = chartRunLook(fitted[piece.run], base, scale);
      return chartRichPiece(name, piece.text);
    })
    .join("");
  return { text, rich };
}

export function chartTitleRichText(
  title: string,
  runs: readonly ChartTitleRun[] | undefined,
  base: Record<string, unknown>,
  scale = 1,
): { text: string; rich: Record<string, Record<string, unknown>> } | null {
  const fitted = runs?.length && chartTitleRunsText(runs) === title ? runs : undefined;
  if (!fitted || !fitted.some((run) => run.style)) return null;
  const rich: Record<string, Record<string, unknown>> = {};
  const text = fitted
    .map((run, index) => {
      const name = `r${index}`;
      const look: Record<string, unknown> = { ...base, ...officeChartTextEChartsStyle(run.style, scale) };
      // A raised or lowered run (Font ▸ Offset, `baseline`) at two thirds of
      // its size against the top or the foot of the line, as PowerPoint draws
      // the CDC TB deck's "disease*": the mark drew full size on the line.
      const baseline = run.style?.baseline;
      if (typeof baseline === "number" && baseline !== 0 && typeof look.fontSize === "number") {
        const size = look.fontSize as number;
        look.fontSize = size * OFFICE_CHART_SUPERSCRIPT_SCALE;
        look.verticalAlign = baseline > 0 ? "top" : "bottom";
        look.lineHeight = size * 1.2;
      }
      rich[name] = look;
      return chartRichPiece(name, run.text);
    })
    .join("");
  return { text, rich };
}
