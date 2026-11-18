/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * In-cell rich text (Excel inline rich text): per-substring formatting within a
 * single cell. REUSES the presentation run model + its tested selection-aware
 * formatting (run split / merge / {@link applyRunFormat}) so cells and slide
 * text share one core. A cell persists `runs` ONLY when they carry formatting
 * the flat `value` + cell style cannot represent; otherwise it stays a plain
 * value cell (byte-stable, no phantom `runs`).
 *
 * Sync invariant: `textFromRuns(cell.runs) === cell.value`. The flat `value` is
 * always the concatenation of the runs (so formulas / sort / CSV export read it
 * unchanged); `runs` only adds per-span styling on top.
 */
import type { CSSProperties } from "react";

import {
  applyRunFormat as sharedApplyRunFormat,
  MAX_CELL_RUN_TEXT_LENGTH,
  MAX_CELL_RUNS,
  normalizeRuns,
  runFieldAtSelection,
  runsFromText,
  textFromRuns,
  type RunStyle as SharedRunStyle,
  type TextRun as SharedTextRun,
} from "../presentation-runs";

const SPREADSHEET_RUN_NORMALIZATION = Object.freeze({
  fontSizePrecision: "hundredth" as const,
  maxTextLength: MAX_CELL_RUN_TEXT_LENGTH,
  maxRuns: MAX_CELL_RUNS,
});

import { nativeFontSizeCss, normalizeNativeFontSize, type NativeFontSize } from "./native-font-size";
import { DOUBLE_UNDERLINE_THICKNESS, scriptBaselineCss } from "./cell-text-effects";
/** `themeFontRole`: the theme font a run's `<scheme>` binds it to, drawn in
 * the theme face for the UI's language whatever its `fontFamily` (the rFont
 * it also names) says, as a cell's own role is (X-328). */
/** `themeColors.color`: the theme slot (and tint) a run's colour was read from, which a save writes back as that
 * slot (`<color theme=… tint=…/>`) instead of a fixed RGB, as the cell's own `themeColors` are. */
export type RunThemeColors = { color?: { theme: number; tint?: number; rgb?: string } };
export type RunStyle = SharedRunStyle & {
  nativeFontSize?: NativeFontSize | null;
  themeFontRole?: "heading" | "body";
  themeColors?: RunThemeColors;
};
export type TextRun = Omit<SharedTextRun, "style"> & { style?: RunStyle };
export { runFieldAtSelection, runsFromText, textFromRuns };

/**
 * Normalize a persisted `runs` value for a cell whose flat text is `value`.
 * Drops flat-equivalent runs (≤1 unstyled run → `undefined`) so a plain cell
 * never carries a phantom array, AND discards runs whose concatenated text does
 * not match `value` (stale after a plain edit — the flat `value` wins). Returns
 * `undefined` for a plain cell.
 */
export function normalizeCellRuns(
  runs: unknown,
  value: string,
): TextRun[] | undefined {
  const normalized = normalizeRuns(runs, SPREADSHEET_RUN_NORMALIZATION);
  if (!normalized) return undefined;
  if (textFromRuns(normalized) !== value) return undefined;
  if (
    !Array.isArray(runs) ||
    !runs.some(run => run?.style && ("nativeFontSize" in run.style || "themeFontRole" in run.style || "themeColors" in run.style))
  ) {
    return normalized;
  }
  const preserved: TextRun[] = [];
  for (const raw of runs) {
    // One plain run normalizes to nothing on its own (a lone unstyled run is
    // flat text): it is still this cell's run, and skipping it made the kept
    // runs short of the value, so every run fell back to the shared form
    // without its native size or theme colour (R94-X08 E13: B2 pasted at G2).
    const one = normalizeRuns([raw], SPREADSHEET_RUN_NORMALIZATION)?.[0] ?? plainRunOf(raw);
    if (!one) continue;
    const pair = raw.style?.nativeFontSize === null ? null : normalizeNativeFontSize(raw.style?.nativeFontSize);
    const role = raw.style?.themeFontRole === "heading" || raw.style?.themeFontRole === "body" ? raw.style.themeFontRole : undefined;
    const themeColors = one.style?.color ? runThemeColors(raw.style?.themeColors) : undefined;
    preserved.push(
      pair !== undefined || role || themeColors
        ? {
            ...one,
            style: {
              ...one.style,
              ...(pair !== undefined ? { nativeFontSize: pair } : {}),
              ...(role ? { themeFontRole: role } : {}),
              ...(themeColors ? { themeColors } : {}),
            },
          }
        : one,
    );
  }
  return textFromRuns(preserved) === value ? preserved : normalized;
}

function plainRunOf(raw: unknown): TextRun | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const text = (raw as { text?: unknown }).text;
  return typeof text === "string" && text !== "" ? { text: text.slice(0, MAX_CELL_RUN_TEXT_LENGTH) } : undefined;
}

/** A run's theme colour reference as the page stores it, or undefined. */
function runThemeColors(raw: unknown): RunThemeColors | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const color = (raw as { color?: unknown }).color;
  if (!color || typeof color !== "object") return undefined;
  const { theme, tint, rgb } = color as { theme?: unknown; tint?: unknown; rgb?: unknown };
  if (typeof theme !== "number" || !Number.isInteger(theme) || theme < 0 || theme > 11) return undefined;
  return {
    color: {
      theme,
      ...(typeof tint === "number" && Number.isFinite(tint) && tint >= -1 && tint <= 1 ? { tint } : {}),
      ...(typeof rgb === "string" && /^#[0-9a-f]{6}$/i.test(rgb) ? { rgb } : {}),
    },
  };
}

/**
 * The runs a cell should KEEP after its flat `value` is (re)written (typing /
 * paste / fill / CSV import). Returns the existing runs ONLY when they still
 * concatenate to the new `value`; otherwise `undefined`, because a value write
 * invalidates stale per-span formatting. This preserves the sync invariant
 * `textFromRuns(runs) === value` on EVERY value write: a fresh value drops the
 * old runs (plain text), while an identical rewrite keeps valid runs. Rich
 * formatting is otherwise only (re)applied via the in-cell dialog, which holds
 * the value constant — never through the value-write path.
 */
export function runsForWrittenValue(
  prevRuns: TextRun[] | undefined,
  value: string,
): TextRun[] | undefined {
  if (!prevRuns || prevRuns.length === 0) return undefined;
  return textFromRuns(prevRuns) === value ? prevRuns : undefined;
}

/**
 * The runs of a cell after an in-place text edit (F2 / double-click / the
 * formula bar editing the cell's own text): the characters the edit kept keep
 * their formatting, and the characters it typed take the formatting of the
 * character before them (of the one after them at the very start), as Excel's
 * cell editor does. The edited span is the part between the longest common
 * prefix and suffix of the old and new text. Undefined when there is nothing
 * to keep (no runs, or every old character replaced), which a value write
 * reads as "plain text".
 */
export function runsForEditedText(
  prevRuns: TextRun[] | undefined,
  value: string,
): TextRun[] | undefined {
  return editedRunsKeepingFormat(prevRuns, value);
}

/**
 * {@link runsForEditedText} for any run shape `{ text, style? }` — the cell
 * runs above, and the header/footer section runs of Page Setup, whose style
 * vocabulary differs (points, double underline) but whose editing rule is the
 * same: kept characters keep their format, typed ones take the format of the
 * character before them.
 */
export function editedRunsKeepingFormat<S extends object>(
  prevRuns: ReadonlyArray<{ text: string; style?: S }> | undefined,
  value: string,
): Array<{ text: string; style?: S }> | undefined {
  type Run = { text: string; style?: S };
  if (!prevRuns || prevRuns.length === 0) return undefined;
  const oldText = prevRuns.map((run) => run.text).join("");
  // An unchanged text keeps the very runs it had (callers compare identity).
  if (oldText === value) return prevRuns as Run[];
  const oldChars = Array.from(oldText);
  const newChars = Array.from(value);
  let prefix = 0;
  while (
    prefix < oldChars.length &&
    prefix < newChars.length &&
    oldChars[prefix] === newChars[prefix]
  ) {
    prefix += 1;
  }
  let suffix = 0;
  while (
    suffix < oldChars.length - prefix &&
    suffix < newChars.length - prefix &&
    oldChars[oldChars.length - 1 - suffix] === newChars[newChars.length - 1 - suffix]
  ) {
    suffix += 1;
  }
  if (prefix === 0 && suffix === 0) return undefined;
  // The style of each old character, by code point.
  const styles: Array<Run["style"]> = [];
  for (const run of prevRuns) {
    const count = Array.from(run.text).length;
    for (let i = 0; i < count; i += 1) styles.push(run.style);
  }
  const typedStyle =
    prefix > 0 ? styles[prefix - 1] : styles[oldChars.length - suffix];
  const out: Run[] = [];
  const push = (text: string, style: Run["style"]) => {
    if (!text) return;
    const last = out.at(-1);
    if (last && JSON.stringify(last.style ?? {}) === JSON.stringify(style ?? {})) {
      last.text += text;
      return;
    }
    out.push(style ? { text, style: { ...style } } : { text });
  };
  for (let i = 0; i < prefix; i += 1) push(oldChars[i], styles[i]);
  push(newChars.slice(prefix, newChars.length - suffix).join(""), typedStyle);
  for (let i = oldChars.length - suffix; i < oldChars.length; i += 1) {
    push(oldChars[i], styles[i]);
  }
  return out.map((run) => run.text).join("") === value &&
    out.some((run) => run.style)
    ? out
    : undefined;
}

/**
 * The runs to FORMAT over for a cell: its existing runs, or the flat `value`
 * wrapped as a single unstyled run (so a first format on a plain cell works).
 */
export function cellRunsForEdit(
  runs: TextRun[] | undefined,
  value: string,
): TextRun[] {
  return runs && runs.length > 0 ? runs : runsFromText(value);
}

/**
 * Map a run's {@link RunStyle} to the CSS a cell text span needs: the font
 * effects + colour Excel inline rich text varies per run (bold / italic /
 * underline / strikethrough / colour / size / family). Cell-level concerns
 * (alignment, background, number format) stay on the cell, not the run.
 */
export function cellRunStyleCss(style: RunStyle | undefined): CSSProperties {
  const css: CSSProperties = {};
  if (!style) return css;
  if (style.bold) css.fontWeight = 700;
  if (style.italic) css.fontStyle = "italic";
  const decoration: string[] = [];
  if (style.underline) decoration.push("underline");
  if (style.strike) decoration.push("line-through");
  // Longhands, as the cell's own font is drawn (spreadsheetCellCss): the
  // `text-decoration` shorthand resets the line's style, and a run's Double
  // Underline (<rPr><u val="double"/>) was drawn as a single line.
  if (decoration.length > 0) css.textDecorationLine = decoration.join(" ");
  if (style.underline && style.underlineStyle === "double") {
    css.textDecorationStyle = "double";
    css.textDecorationThickness = DOUBLE_UNDERLINE_THICKNESS;
  }
  if (typeof style.color === "string" && style.color) css.color = style.color;
  if (typeof style.fontSize === "number" && Number.isFinite(style.fontSize)) {
    css.fontSize = nativeFontSizeCss(style.fontSize, style.nativeFontSize);
  }
  if (typeof style.fontFamily === "string" && style.fontFamily) {
    css.fontFamily = style.fontFamily;
  }
  // The 2 of H2O: raised or lowered off the baseline and drawn smaller — of
  // the run's own size, or of its line's when it has none (live audit B399).
  const script = style.superscript ? "superscript" : style.subscript ? "subscript" : null;
  if (script) Object.assign(css, scriptBaselineCss(script, css.fontSize));
  return css;
}

/** Keep spreadsheet point mirrors through the shared selection formatter.
 * The presentation formatter is still the authority for supported fields;
 * only unchanged spreadsheet size/point metadata is restored on its slices. */
export function applyRunFormat(runs: readonly TextRun[], start: number, end: number, patch: Partial<RunStyle>): TextRun[] {
  const total = textFromRuns(runs).length;
  const from = Math.min(total, Math.max(0, Math.round(start)));
  const to = Math.min(total, Math.max(0, Math.round(end)));
  const out: TextRun[] = [];
  let cursor = 0;
  for (const run of runs) {
    const low = Math.max(0, from - cursor), high = Math.min(run.text.length, to - cursor);
    if (from >= to || high <= low) {
      out.push({...run, ...(run.style ? {style:{...run.style}} : {})});
    } else {
      if (low > 0) out.push({...run,text:run.text.slice(0,low)});
      const text = run.text.slice(low,high);
      // In hundredths, as a cell's runs keep their size: rounded to whole
      // pixels, a 309.5 pt cell (412.67 px) gave its runs 413 px, 309.75 pt
      // (xlsx-rounds R443-X80 WALK2, Font Size on a rich cell).
      const formatted = sharedApplyRunFormat([{...run,text}],0,text.length,patch,SPREADSHEET_RUN_NORMALIZATION);
      for (const next of formatted) {
        const style: RunStyle = {...next.style};
        if (!Object.prototype.hasOwnProperty.call(patch,"fontSize")) {
          if (run.style?.fontSize !== undefined) style.fontSize = run.style.fontSize;
          if (run.style?.nativeFontSize !== undefined) style.nativeFontSize = run.style.nativeFontSize;
        } else if (run.style?.nativeFontSize !== undefined) style.nativeFontSize = null;
        // A face picked for the characters replaces the theme font they
        // were bound to; any other format keeps the binding.
        if (run.style?.themeFontRole && !Object.prototype.hasOwnProperty.call(patch,"fontFamily")) style.themeFontRole = run.style.themeFontRole;
        out.push({...next,style});
      }
      if (high < run.text.length) out.push({...run,text:run.text.slice(high)});
    }
    cursor += run.text.length;
  }
  const merged: TextRun[] = [];
  for (const run of out) {
    const last = merged.at(-1);
    if (last && JSON.stringify(last.style ?? {}) === JSON.stringify(run.style ?? {})) last.text += run.text;
    else merged.push({...run});
  }
  return merged;
}
