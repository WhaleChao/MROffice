/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ptToPx } from "./office-units";
import { pointsToCssPx } from "./presentation/text-units";
import { cssFontFamilyList } from "./css-font-family";
import { EMPTY_PRESENTATION_FONT_ALIASES, presentationFontFamilyWithAssets } from "./presentation-font-face-reads";
import type { CSSProperties } from "react";
import {
  normalizeRunMath,
  normalizeRunStyle,
  runCarriesMath,
  splitRun,
  type RunMath,
  type RunStyle,
  type TextRun,
} from "./presentation-runs";
import {
  normalizePresentationUnderlineStyle,
  presentationTextDecorationCss,
} from "./presentation-text-decoration";
import { officeTextPaintCss, officeTextPaintWithFontColor, type OfficeTextPaint } from "./office-text-paint";

/** A run's own WordArt as CSS (none, or not a valid paint: drawn plain). */
function cellRunTextPaintCss(style: RunStyle | undefined): CSSProperties {
  if (!style?.textPaint) return {};
  try {
    return officeTextPaintCss(style.textPaint);
  } catch {
    return {};
  }
}

export type CellTextRange = Readonly<{ start: number; end: number }>;
const FLAGS = [
  "bold",
  "italic",
  "underline",
  "strike",
  "superscript",
  "subscript",
] as const;
export const CELL_CHARACTER_FIELDS = new Set<string>([
  ...FLAGS,
  "underlineStyle",
  "strikeStyle",
  "fontFamily",
  "fontSize",
  "color",
  "highlight",
]);

// Unlike an ordinary slide run, false must survive here: it overrides an
// inherited cell/header format without changing the surrounding characters.
export function normalizeCellCharacterStyle(
  value: unknown,
): RunStyle | undefined {
  const style: RunStyle = {
    ...normalizeRunStyle(value, { fontSizePrecision: "hundredth" }),
  };
  if (value && typeof value === "object") {
    const raw = value as Record<string, unknown>;
    for (const key of FLAGS) if (raw[key] === false) style[key] = false;
    // Any DrawingML underline kind an imported cell run carries (the slide
    // run model keeps only the double line).
    const underlineStyle = normalizePresentationUnderlineStyle(raw.underlineStyle);
    if (style.underline === true && underlineStyle) style.underlineStyle = underlineStyle;
    else delete style.underlineStyle;
    if (style.strike === true && raw.strikeStyle === "double") style.strikeStyle = "double";
    else delete style.strikeStyle;
    // Cell sizes are points at the ribbon, CSS pixels inside character runs.
    if (
      typeof raw.fontSize === "number" &&
      Number.isFinite(raw.fontSize) &&
      raw.fontSize >= 1 &&
      raw.fontSize <= ptToPx(400)
    ) {
      style.fontSize = Math.round(raw.fontSize * 100) / 100;
    }
  }
  return Object.keys(style).length ? style : undefined;
}

/**
 * Adjacent runs of one format merged into one. An equation run (`math`,
 * PPT-A052) is ONE atom around the characters it draws, as in a text box: it
 * never merges with a neighbour, and its marker survives only on a run whose
 * text holds no line break (an `a14:m` holds none).
 */
export function mergeCellTextRuns(runs: readonly TextRun[]): TextRun[] {
  const out: TextRun[] = [];
  for (const run of runs) {
    if (!run.text) continue;
    const style = normalizeCellCharacterStyle(run.style);
    const math = normalizeRunMath(run.math, run.text);
    const last = out.at(-1);
    if (
      last &&
      !math &&
      !last.math &&
      JSON.stringify(last.style) === JSON.stringify(style)
    )
      last.text += run.text;
    else {
      const next: TextRun = style ? { text: run.text, style } : { text: run.text };
      if (math) next.math = math;
      out.push(next);
    }
  }
  return out;
}

export function cellCharacterRuns(text: string, value: unknown): TextRun[] {
  if (!Array.isArray(value) || value.length > 200)
    return text ? [{ text }] : [];
  if (value.some((run) => !run || typeof run.text !== "string"))
    return text ? [{ text }] : [];
  const runs = mergeCellTextRuns(value);
  return runs.map((run) => run.text).join("") === text
    ? runs
    : text
      ? [{ text }]
      : [];
}

/**
 * The `runs` a cell should PERSIST beside `text`, or `undefined` when the text
 * carries no character formatting (an unstyled cell stays `{ text }` clean).
 * Runs that do not concatenate exactly to `text` are dropped rather than kept:
 * the API refuses a cell whose runs disagree with its text, so every bulk edit
 * that rewrites a cell's text (paste, clear, merge) goes through here.
 */
export function persistedCellRuns(
  text: string,
  value: unknown,
): TextRun[] | undefined {
  const runs = cellCharacterRuns(text, value);
  // An equation is a run of its own even in the cell's one format: dropping
  // the runs would leave its characters as plain text.
  return runs.some((run) => run.style !== undefined || runCarriesMath(run))
    ? runs
    : undefined;
}

/**
 * A cell's own direct character format — the size, colour, font or flag a
 * row or cell operation wrote on the cell itself — spelled as a run style:
 * sizes are points on the cell and CSS pixels in a run.
 */
export function cellDirectCharacterStyle(
  cell: Record<string, unknown>,
): RunStyle | undefined {
  const style: Record<string, unknown> = {};
  for (const key of FLAGS) if (cell[key] === true) style[key] = true;
  if (cell.underline === true && typeof cell.underlineStyle === "string")
    style.underlineStyle = cell.underlineStyle;
  if (cell.strike === true && cell.strikeStyle === "double")
    style.strikeStyle = "double";
  for (const key of ["fontFamily", "color", "highlight"])
    if (typeof cell[key] === "string" && cell[key]) style[key] = cell[key];
  if (
    typeof cell.fontSize === "number" &&
    Number.isFinite(cell.fontSize) &&
    cell.fontSize > 0
  )
    style.fontSize = ptToPx(cell.fontSize);
  return normalizeCellCharacterStyle(style);
}

/**
 * Clear Formatting over part of a cell whose size, colour or flag sits on
 * the cell itself (a row or cell operation wrote it): the cell's direct
 * format is pressed into the characters outside the range, so they keep
 * their look, and dropped from the cell and from the characters inside it,
 * which return to the table style (PPT-B212). `null` when the cell holds
 * none of `keys` itself — the plain character clear is enough then.
 */
export function clearCellCharacterFormat(
  cell: { text: string; runs?: unknown } & Record<string, unknown>,
  range: CellTextRange,
  keys: readonly string[],
): { cleared: string[]; runs: TextRun[] | undefined } | null {
  const direct = (cellDirectCharacterStyle(cell) ?? {}) as Record<
    string,
    unknown
  >;
  const character = keys.filter((key) => CELL_CHARACTER_FIELDS.has(key));
  const cleared = character.filter((key) => direct[key] !== undefined);
  if (!cleared.length || !validCellTextRange(cell.text, range)) return null;
  const pressed = Object.fromEntries(cleared.map((key) => [key, direct[key]]));
  const baked = cellCharacterRuns(cell.text, cell.runs).map((run) => ({
    ...run,
    style: normalizeCellCharacterStyle({ ...pressed, ...run.style }),
  }));
  const patch = Object.fromEntries(
    character.map((key) => [key, undefined]),
  ) as Partial<RunStyle>;
  return {
    cleared,
    runs: persistedCellRuns(cell.text, formatCellTextRange(baked, range, patch)),
  };
}

export function validCellTextRange(
  text: string,
  range: CellTextRange,
): boolean {
  const boundary = (offset: number) =>
    !(
      offset > 0 &&
      offset < text.length &&
      /[\uD800-\uDBFF]/.test(text[offset - 1]) &&
      /[\uDC00-\uDFFF]/.test(text[offset])
    );
  return (
    Number.isInteger(range.start) &&
    Number.isInteger(range.end) &&
    range.start >= 0 &&
    range.start <= range.end &&
    range.end <= text.length &&
    boundary(range.start) &&
    boundary(range.end)
  );
}

/**
 * `cellPaint` is the WordArt the whole cell wears (Table Design > WordArt
 * Styles on the cell): characters without a paint of their own wear it, and
 * it draws over their plain colour.
 */
export function patchCellCharacterStyle(
  base: RunStyle | undefined,
  patch: Partial<RunStyle>,
  cellPaint?: OfficeTextPaint | null,
): RunStyle | undefined {
  const next: RunStyle = { ...base };
  for (const [key, value] of Object.entries(patch)) {
    if (!CELL_CHARACTER_FIELDS.has(key)) continue;
    if ((FLAGS as readonly string[]).includes(key))
      (next as Record<string, unknown>)[key] = value === true;
    else if (value === undefined || value === null || value === "")
      delete (next as Record<string, unknown>)[key];
    else (next as Record<string, unknown>)[key] = value;
  }
  if (Object.hasOwn(patch, "color")) {
    // Font Color on WordArt characters is their fill, outline and effects
    // kept: the paint drew them in the old colour beside the new one. Inside
    // a WordArt cell a colour needs a paint of the characters' own to show
    // over the cell's; Automatic gives them back the cell's.
    const color = patch.color || undefined;
    const worn = next.textPaint ?? (color ? cellPaint : null) ?? null;
    const paint = worn ? officeTextPaintWithFontColor(worn, color, { keepBare: Boolean(cellPaint) && color !== undefined }) : null;
    if (paint) next.textPaint = paint;
    else delete next.textPaint;
  }
  if (patch.superscript) next.subscript = false;
  if (patch.subscript) next.superscript = false;
  return normalizeCellCharacterStyle(next);
}

function sliceRuns(
  runs: readonly TextRun[],
  from: number,
  to: number,
): TextRun[] {
  const out: TextRun[] = [];
  let offset = 0;
  for (const run of runs) {
    const start = Math.max(from - offset, 0),
      end = Math.min(to - offset, run.text.length);
    if (start < end)
      out.push(splitRun(splitRun(run, start)[1], end - start)[0]);
    offset += run.text.length;
  }
  return out;
}

/**
 * A character format: one patch for every run, or one worked out from each
 * run's own style (Grow / Shrink Font steps each run from its own size, as
 * PowerPoint and Word do over words of mixed sizes).
 */
export type CellCharacterPatch =
  | Partial<RunStyle>
  | ((style: RunStyle | undefined) => Partial<RunStyle>);

export function formatCellTextRange(
  runs: readonly TextRun[],
  range: CellTextRange,
  patch: CellCharacterPatch,
  cellPaint?: OfficeTextPaint | null,
): TextRun[] {
  const text = runs.map((run) => run.text).join("");
  if (!validCellTextRange(text, range) || range.start === range.end)
    return [...runs];
  return mergeCellTextRuns([
    ...sliceRuns(runs, 0, range.start),
    // A whole equation inside the range keeps its `math`; its format rides
    // on the zone's own runs (`build_math_zone` takes the run's `a:rPr`).
    ...sliceRuns(runs, range.start, range.end).map((run) => ({
      ...run,
      style: patchCellCharacterStyle(
        run.style,
        typeof patch === "function" ? patch(run.style) : patch,
        cellPaint,
      ),
    })),
    ...sliceRuns(runs, range.end, text.length),
  ]);
}

/**
 * The runs with each run of `range` restyled by `map` (its style in, its
 * style out), the runs cut at the range's ends and merged after. What a
 * command that depends on each run's own value writes (WordArt: each run's
 * paint edited). An empty or invalid range changes nothing.
 */
export function mapCellTextRange(
  runs: readonly TextRun[],
  range: CellTextRange,
  map: (style: RunStyle | undefined) => RunStyle | undefined,
): TextRun[] {
  const text = runs.map((run) => run.text).join("");
  if (!validCellTextRange(text, range) || range.start === range.end) return [...runs];
  return mergeCellTextRuns([
    ...sliceRuns(runs, 0, range.start),
    ...sliceRuns(runs, range.start, range.end).map((run) => ({
      ...run,
      style: normalizeCellCharacterStyle(map(run.style)),
    })),
    ...sliceRuns(runs, range.end, text.length),
  ]);
}

export function cellCharacterStyleAt(
  runs: readonly TextRun[],
  offset: number,
): RunStyle | undefined {
  let end = 0;
  for (const [index, run] of runs.entries()) {
    end += run.text.length;
    if (offset <= end) {
      // Text typed beside an equation is not equation text: it takes the
      // format of the text around the equation (before it, else after it),
      // not the math face and italic of the zone's runs.
      if (runCarriesMath(run)) {
        const plain =
          runs.slice(0, index).reverse().find((item) => !runCarriesMath(item)) ??
          runs.slice(index + 1).find((item) => !runCarriesMath(item));
        if (plain) return plain.style;
      }
      return run.style;
    }
  }
  return runs.at(-1)?.style;
}

/** An equation as a cell inserts it: what it draws, and the equation itself. */
export type CellEquation = Readonly<{
  /**
   * The characters the server's Office Math draws for `latex` (`POST
   * /api/v1/presentation/equation-text`); never computed in the browser.
   */
  text: string;
  latex: string;
  display?: boolean;
}>;

/**
 * The equation run whose characters are exactly `range`, with its offsets,
 * else null: the selection a click on an equation makes, which "Edit
 * Equation" rewrites.
 */
export function cellEquationInRange(
  runs: readonly TextRun[],
  range: CellTextRange,
): { start: number; end: number; run: TextRun & { math: RunMath } } | null {
  let offset = 0;
  for (const run of runs) {
    const end = offset + run.text.length;
    if (offset === range.start && end === range.end && runCarriesMath(run))
      return { start: offset, end, run: run as TextRun & { math: RunMath } };
    if (offset > range.end) break;
    offset = end;
  }
  return null;
}

/**
 * Words typed beside a display equation make it an in-line one, as Word and
 * PowerPoint do: an equation is drawn on a line of its own only while it is
 * alone in its paragraph. Each paragraph (split at "\n") holding any other
 * character (a line break `\v` aside) turns its display equations in-line;
 * the runs come back unchanged (the same array) when none does.
 */
export function inlineEquationsBesideWords(runs: readonly TextRun[]): TextRun[] {
  const worded = new Set<number>();
  let paragraph = 0;
  for (const run of runs) {
    if (run.math) continue;
    const pieces = run.text.split("\n");
    pieces.forEach((piece, index) => {
      if (index > 0) paragraph += 1;
      if (piece.replace(/\v/g, "") !== "") worded.add(paragraph);
    });
  }
  if (worded.size === 0) return runs as TextRun[];
  paragraph = 0;
  let changed = false;
  const next = runs.map((run) => {
    if (!run.math) {
      paragraph += run.text.split("\n").length - 1;
      return run;
    }
    if (run.math.display !== true || !worded.has(paragraph)) return run;
    changed = true;
    return { ...run, math: { ...run.math, display: false } };
  });
  return changed ? next : (runs as TextRun[]);
}

/**
 * Insert an equation in place of `range` (Insert > Equation at the caret or
 * over a selection), or rewrite the equation `range` is exactly (Edit
 * Equation): one atomic run carrying `math`. The new run takes the format of
 * the text it replaces or stands beside, as typed text would. `null` when the
 * range or the equation is not one a cell can hold.
 */
export function insertCellEquation(
  runs: readonly TextRun[],
  range: CellTextRange,
  equation: CellEquation,
  style?: RunStyle,
): TextRun[] | null {
  const math = normalizeRunMath(
    { latex: equation.latex, display: equation.display === true },
    equation.text,
  );
  if (!equation.text || !math) return null;
  const replaced = cellEquationInRange(runs, range);
  const runStyle =
    style ??
    replaced?.run.style ??
    cellCharacterStyleAt(runs, range.end > range.start ? range.start + 1 : range.start);
  return replaceCellTextRangeWithRuns(runs, range, [
    { text: equation.text, ...(runStyle ? { style: runStyle } : {}), math },
  ]);
}

export function replaceCellTextRange(
  runs: readonly TextRun[],
  range: CellTextRange,
  inserted: string,
  typingStyle?: RunStyle,
): TextRun[] {
  const text = runs.map((run) => run.text).join("");
  if (
    !validCellTextRange(text, range) ||
    Array.from(inserted).some(
      (char) => char.length === 1 && /[\uD800-\uDFFF]/.test(char),
    )
  )
    return [...runs];
  // Typed at a caret, the new text takes the style before it; typed over a
  // selection, the style of the first character it replaces, as Word and
  // PowerPoint do (PPT-B317: a whole bold word retyped at a run boundary took
  // the plain style of the word before it).
  const style =
    typingStyle ??
    cellCharacterStyleAt(runs, range.end > range.start ? range.start + 1 : range.start);
  return mergeCellTextRuns([
    ...sliceRuns(runs, 0, range.start),
    ...(inserted ? [{ text: inserted, ...(style ? { style } : {}) }] : []),
    ...sliceRuns(runs, range.end, text.length),
  ]);
}







function cellRunScriptSize(style: RunStyle | undefined, size: CSSProperties["fontSize"]): CSSProperties["fontSize"] {
  if (!style?.superscript && !style?.subscript) return size;
  const px = typeof size === "number" ? size : typeof size === "string" ? /^([\d.]+)px$/.exec(size)?.[1] : undefined;
  const value = Number(px);
  if (!(value > 0)) return size === undefined ? "0.67em" : size;
  return `${(Math.max(1, Math.round((value * 0.75 * 2) / 3)) * 4) / 3}px`;
}

/**
 * The face a cell run's symbol-encoded characters (U+F000-U+F0FF) are drawn
 * from: its symbol font (`RunStyle.symbolFont`, `a:sym`) as a slide run's
 * resolves it (`runSymbolFontCss`: the served open clone behind the Office
 * name), read here without that module, which holds React state a server
 * component may not import. Undefined when the run names none.
 */
export function cellRunSymbolFontFamily(style: RunStyle | undefined): string | undefined {
  const face = typeof style?.symbolFont === "string" && style.symbolFont ? style.symbolFont : undefined;
  const served = face ? presentationFontFamilyWithAssets(face, EMPTY_PRESENTATION_FONT_ALIASES) : undefined;
  return served ? cssFontFamilyList(served) || undefined : undefined;
}








export function cellRunSymbolPieces(text: string): { readonly text: string; readonly symbol: boolean }[] {
  const pieces: { text: string; symbol: boolean }[] = [];
  for (const character of text) {
    const code = character.codePointAt(0) ?? 0;
    const symbol = code >= 0xf000 && code <= 0xf0ff;
    const last = pieces[pieces.length - 1];
    if (last && last.symbol === symbol) last.text += character;
    else pieces.push({ text: character, symbol });
  }
  return pieces;
}

export function cellTextRunCss(
  style: RunStyle | undefined,
  base: CSSProperties = {},
  inherited: RunStyle = {},
): CSSProperties {
  style = { ...inherited, ...style };
  const has = (key: keyof RunStyle) =>
    style && Object.prototype.hasOwnProperty.call(style, key);
  const decoration = [
    has("underline")
      ? style?.underline
      : String(base.textDecoration ?? "").includes("underline"),
    has("strike")
      ? style?.strike
      : String(base.textDecoration ?? "").includes("line-through"),
  ];
  // The decoration's kind: the run's own while it states its underline or
  // strikethrough, else the cell's (carried on the base CSS).
  const kindCss =
    has("underline") || has("strike")
      ? presentationTextDecorationCss(
          decoration[0]
            ? normalizePresentationUnderlineStyle(
                has("underline") ? style?.underlineStyle : undefined,
              )
            : undefined,
          decoration[1] && style?.strikeStyle === "double" ? "double" : undefined,
        )
      : decoration[0] || decoration[1]
        ? {
            textDecorationStyle: base.textDecorationStyle,
            textDecorationThickness: base.textDecorationThickness,
          }
        : {};
  // A run stated not bold takes the words' own regular weight (a Docs
  // shape's label is set in 500), not a fixed 400: unbolding a word drew it
  // lighter while it was typed than where it was drawn.
  const regularWeight =
    Number(base.fontWeight) >= 600 ? 400 : (base.fontWeight ?? 400);
  return {
    ...kindCss,
    fontWeight: has("bold") ? (style?.bold ? 700 : regularWeight) : base.fontWeight,
    // All caps / Small caps draw the words in capitals (the text is kept),
    // and a text outline draws round them, as the slide's own runs do.
    ...(style?.caps === "all"
      ? { textTransform: "uppercase" as const }
      : style?.caps === "small"
        ? { fontVariantCaps: "small-caps" as const }
        : {}),
    ...(typeof style?.textOutline === "string" && style.textOutline
      ? { WebkitTextStroke: `1px ${style.textOutline}` }
      : {}),
    fontStyle: has("italic")
      ? style?.italic
        ? "italic"
        : "normal"
      : base.fontStyle,
    // The longhand, never the `text-decoration` shorthand: written after the
    // kind above (as a hand-painted span or a React style writes keys in
    // order), the shorthand put the line back to solid, so a double
    // underline was drawn single.
    textDecorationLine:
      [decoration[0] ? "underline" : "", decoration[1] ? "line-through" : ""]
        .filter(Boolean)
        .join(" ") || "none",
    color: style?.color ?? base.color,
    backgroundColor: style?.highlight ?? base.backgroundColor,
    // A face this machine does not have falls to the words' own (the table's),
    
    // "Mukta Vaani Regular" drawn in Calibri).
    fontFamily:
      style?.fontFamily && typeof base.fontFamily === "string" && base.fontFamily && base.fontFamily !== style.fontFamily
        ? `${style.fontFamily}, ${base.fontFamily}`
        : (style?.fontFamily ?? base.fontFamily),
    fontSize: cellRunScriptSize(style, style?.fontSize !== undefined ? `${style.fontSize}px` : base.fontSize),
    // Character spacing on these characters (points to px), as a slide's own
    // run draws it (`runStyleCss`).
    ...(typeof style?.letterSpacing === "number"
      ? { letterSpacing: `${pointsToCssPx(style.letterSpacing)}px` }
      : {}),
    verticalAlign: style?.superscript
      ? "super"
      : style?.subscript
        ? "sub"
        : "baseline",
    // WordArt on these characters: drawn over their colour and outline, as
    // a slide's own run draws it (`runStyleCss`).
    ...cellRunTextPaintCss(style),
  };
}

/**
 * The clipboard flavour a cell's character selection carries beside
 * `text/plain` / `text/html`: its runs, so a paste into another cell keeps the
 * characters' formatting (PPT-B020). Read back only when it spells exactly the
 * plain text of the same copy.
 */
export const CELL_RUNS_CLIPBOARD_TYPE = "application/x-mrai-presentation-cell-runs";

/** The runs of `range` (formatting kept, adjacent equal runs merged). */
export function cellTextRunsInRange(
  runs: readonly TextRun[],
  range: CellTextRange,
): TextRun[] {
  const text = runs.map((run) => run.text).join("");
  if (!validCellTextRange(text, range)) return [];
  return mergeCellTextRuns(sliceRuns(runs, range.start, range.end));
}

/** Replace `range` with formatted runs (a rich paste), or `null` if refused. */
export function replaceCellTextRangeWithRuns(
  runs: readonly TextRun[],
  range: CellTextRange,
  inserted: readonly TextRun[],
): TextRun[] | null {
  const text = runs.map((run) => run.text).join("");
  if (!validCellTextRange(text, range)) return null;
  return mergeCellTextRuns([
    ...sliceRuns(runs, 0, range.start),
    ...inserted,
    ...sliceRuns(runs, range.end, text.length),
  ]);
}

/** The runs a cell copy wrote, when they are valid and spell `plain`. */
export function parseCellRunsClipboard(
  raw: string,
  plain: string,
): TextRun[] | null {
  if (!raw || raw.length > 200_000) return null;
  let value: unknown;
  try {
    value = JSON.parse(raw);
  } catch {
    return null;
  }
  if (
    !Array.isArray(value) ||
    value.length === 0 ||
    value.length > 200 ||
    value.some((run) => !run || typeof run.text !== "string")
  ) {
    return null;
  }
  const runs = mergeCellTextRuns(value as TextRun[]);
  return runs.map((run) => run.text).join("") === plain.replace(/\r\n?/g, "\n")
    ? runs
    : null;
}

/**
 * Character formatting another program's `text/html` states on the pasted
 * text, conservatively: bold, italic, underline and strikethrough from their
 * tags and inline styles, nothing it merely inherits (face, size and colour
 * stay the destination's). `null` unless the HTML's text is exactly `plain`
 * (whitespace or structure the plain flavour does not share means the HTML
 * cannot be trusted to line up) or nothing in it is formatted.
 */
export function cellRunsFromClipboardHtml(
  html: string,
  plain: string,
): TextRun[] | null {
  if (!html || html.length > 1_000_000 || typeof DOMParser === "undefined") {
    return null;
  }
  const body = new DOMParser().parseFromString(html, "text/html").body;
  const expected = plain.replace(/\r\n?/g, "\n");
  const out: TextRun[] = [];
  const flags = (element: Element | null): RunStyle => {
    const style: RunStyle = {};
    for (let node = element; node && node !== body; node = node.parentElement) {
      const tag = node.tagName;
      const css = (node as HTMLElement).style;
      const weight = css?.fontWeight ?? "";
      if (
        tag === "B" ||
        tag === "STRONG" ||
        weight === "bold" ||
        weight === "bolder" ||
        Number(weight) >= 600
      ) {
        style.bold ??= true;
      }
      if (tag === "I" || tag === "EM" || css?.fontStyle === "italic") style.italic ??= true;
      const decoration = `${css?.textDecoration ?? ""} ${css?.textDecorationLine ?? ""}`;
      if (tag === "U" || decoration.includes("underline")) style.underline ??= true;
      if (tag === "S" || tag === "STRIKE" || tag === "DEL" || decoration.includes("line-through")) {
        style.strike ??= true;
      }
    }
    return style;
  };
  const walker = body.ownerDocument.createTreeWalker(body, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node.textContent ?? "";
    if (!text) continue;
    const style = flags(node.parentElement);
    out.push(Object.keys(style).length ? { text, style } : { text });
  }
  const runs = mergeCellTextRuns(out);
  return runs.map((run) => run.text).join("") === expected &&
    runs.some((run) => run.style)
    ? runs
    : null;
}
