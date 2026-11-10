/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pdfTextEditCommonEnds } from "./text-edit-common-ends";
/**
 * Inline typography carried by an editable PDF text box.
 *
 * PDF text is emitted as positioned glyph runs, not semantic paragraphs.  A
 * visual paragraph can therefore contain several fonts / weights even when it
 * is edited as one box.  Keeping those runs alongside the flat `text` lets the
 * editor preserve untouched typography while ordinary typing changes only the
 * affected characters.
 *
 * Invariant: when `runs` is present, `textFromPdfTextStyleRuns(runs) === text`.
 */

export interface PdfTextStyleRun {
  text: string;
  /** Font size as a fraction of the rendered page height. */
  fontSize?: number;
  /** Stable, persistable CSS family (never a pdf.js synthetic g_d*_f* name). */
  fontFamily?: string;
  /** Original PDF/PostScript font name used by export face matching. */
  sourceFontName?: string;
  /** Live pdf.js family. Render-only: stripped before persistence. */
  displayFontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  /**
   * Word's character decorations and vertical position of these characters
   * only (a text box's own flags cover its whole text). Absent inherits the
   * box's.
   */
  underline?: boolean;
  strikethrough?: boolean;
  script?: "superscript" | "subscript";
  /**
   * An Office font family (a name of lib/office-fonts' catalog, e.g. "Calibri"
   * or "宋体") these characters are set in. The text-box writer embeds the
   * served open substitute's own font program for it (text-box-runs.ts);
   * `fontFamily` then holds the CSS stack that draws the same substitute.
   * Absent: the face is the portable class named by `sourceFontName` /
   * `fontFamily` (a base-14 face, and Noto Sans SC for East Asian text).
   */
  officeFamily?: string;
  /** Live extraction-only sample rect (page fractions); stripped on save. */
  sampleX?: number;
  sampleY?: number;
  sampleWidth?: number;
  sampleHeight?: number;
}

const MAX_PERSISTED_RUNS = 5_000;
const MAX_LIVE_RUNS = 20_000;
const MAX_RUN_TEXT = 20_000;
const PDFJS_SYNTHETIC_FONT_RE = /\bg_d\d+_f\d+\b/i;
const HEX_COLOR_RE = /^#[0-9a-fA-F]{6}$/;

type NormalizeOptions = {
  expectedText?: string;
  allowDisplayFont?: boolean;
};

function boundedString(value: unknown, max: number): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed ? trimmed.slice(0, max) : undefined;
}

function normalizeRun(
  value: unknown,
  allowDisplayFont: boolean,
): PdfTextStyleRun | null {
  if (!value || typeof value !== "object") return null;
  const record = value as Record<string, unknown>;
  if (typeof record.text !== "string" || record.text.length === 0) return null;

  const run: PdfTextStyleRun = {
    text: record.text.slice(0, MAX_RUN_TEXT),
  };
  const fontSize = Number(record.fontSize);
  if (Number.isFinite(fontSize) && fontSize > 0 && fontSize <= 1) {
    run.fontSize = fontSize;
  }
  const fontFamily = boundedString(record.fontFamily, 400);
  if (fontFamily && !PDFJS_SYNTHETIC_FONT_RE.test(fontFamily)) {
    run.fontFamily = fontFamily;
  }
  const sourceFontName = boundedString(record.sourceFontName, 200);
  if (sourceFontName) run.sourceFontName = sourceFontName;
  if (allowDisplayFont) {
    const displayFontFamily = boundedString(record.displayFontFamily, 600);
    if (displayFontFamily) run.displayFontFamily = displayFontFamily;
  }
  if (typeof record.bold === "boolean") run.bold = record.bold;
  if (typeof record.italic === "boolean") run.italic = record.italic;
  if (typeof record.color === "string" && HEX_COLOR_RE.test(record.color)) {
    run.color = record.color.toLowerCase();
  }
  if (typeof record.underline === "boolean") run.underline = record.underline;
  if (typeof record.strikethrough === "boolean") {
    run.strikethrough = record.strikethrough;
  }
  if (record.script === "superscript" || record.script === "subscript") {
    run.script = record.script;
  }
  const officeFamily = boundedString(record.officeFamily, 128);
  if (officeFamily) run.officeFamily = officeFamily;
  if (allowDisplayFont) {
    const sampleX = Number(record.sampleX);
    const sampleY = Number(record.sampleY);
    const sampleWidth = Number(record.sampleWidth);
    const sampleHeight = Number(record.sampleHeight);
    if (
      Number.isFinite(sampleX) &&
      Number.isFinite(sampleY) &&
      Number.isFinite(sampleWidth) &&
      Number.isFinite(sampleHeight) &&
      sampleX >= 0 &&
      sampleY >= 0 &&
      sampleWidth > 0 &&
      sampleHeight > 0 &&
      sampleX <= 1 &&
      sampleY <= 1 &&
      sampleWidth <= 1 &&
      sampleHeight <= 1
    ) {
      run.sampleX = sampleX;
      run.sampleY = sampleY;
      run.sampleWidth = sampleWidth;
      run.sampleHeight = sampleHeight;
    }
  }
  return run;
}

function runStyleEquals(a: PdfTextStyleRun, b: PdfTextStyleRun): boolean {
  return (
    a.fontSize === b.fontSize &&
    a.fontFamily === b.fontFamily &&
    a.sourceFontName === b.sourceFontName &&
    a.displayFontFamily === b.displayFontFamily &&
    a.bold === b.bold &&
    a.italic === b.italic &&
    a.color === b.color &&
    a.underline === b.underline &&
    a.strikethrough === b.strikethrough &&
    a.script === b.script &&
    a.officeFamily === b.officeFamily &&
    a.sampleX === b.sampleX &&
    a.sampleY === b.sampleY &&
    a.sampleWidth === b.sampleWidth &&
    a.sampleHeight === b.sampleHeight
  );
}

export function mergePdfTextStyleRuns(
  runs: readonly PdfTextStyleRun[],
): PdfTextStyleRun[] {
  const merged: PdfTextStyleRun[] = [];
  for (const raw of runs) {
    if (!raw.text) continue;
    const run = { ...raw };
    const previous = merged[merged.length - 1];
    if (previous && runStyleEquals(previous, run)) {
      previous.text += run.text;
    } else {
      merged.push(run);
    }
  }
  return merged;
}

export function textFromPdfTextStyleRuns(
  runs: readonly PdfTextStyleRun[] | undefined,
): string {
  return runs?.map((run) => run.text).join("") ?? "";
}

/**
 * Validate an untrusted/persisted run array and enforce the flat-text invariant.
 * A single run is intentionally retained: it may carry an exact source face
 * that is more precise than the box-level fallback.
 */
export function normalizePdfTextStyleRuns(
  value: unknown,
  options: NormalizeOptions = {},
): PdfTextStyleRun[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const cleaned: PdfTextStyleRun[] = [];
  const maxRuns = options.allowDisplayFont ? MAX_LIVE_RUNS : MAX_PERSISTED_RUNS;
  for (const entry of value.slice(0, maxRuns)) {
    const run = normalizeRun(entry, options.allowDisplayFont === true);
    if (run) cleaned.push(run);
  }
  const merged = mergePdfTextStyleRuns(cleaned);
  if (merged.length === 0) return undefined;
  if (
    typeof options.expectedText === "string" &&
    textFromPdfTextStyleRuns(merged) !== options.expectedText
  ) {
    return undefined;
  }
  return merged;
}

/** Strip render-only pdf.js font aliases before a run enters page JSON. */
export function persistablePdfTextStyleRuns(
  value: unknown,
  expectedText: string,
): PdfTextStyleRun[] | undefined {
  const normalized = normalizePdfTextStyleRuns(value, {
    expectedText,
    allowDisplayFont: true,
  });
  if (!normalized) return undefined;
  return normalizePdfTextStyleRuns(
    normalized.map((run) => {
      const persistable = { ...run };
      delete persistable.displayFontFamily;
      delete persistable.sampleX;
      delete persistable.sampleY;
      delete persistable.sampleWidth;
      delete persistable.sampleHeight;
      return persistable;
    }),
    { expectedText },
  );
}

function splitRunsAt(
  runs: readonly PdfTextStyleRun[],
  offset: number,
): [PdfTextStyleRun[], PdfTextStyleRun[]] {
  const before: PdfTextStyleRun[] = [];
  const after: PdfTextStyleRun[] = [];
  let cursor = 0;
  for (const run of runs) {
    const end = cursor + run.text.length;
    if (end <= offset) {
      before.push({ ...run });
    } else if (cursor >= offset) {
      after.push({ ...run });
    } else {
      const local = offset - cursor;
      if (local > 0) before.push({ ...run, text: run.text.slice(0, local) });
      if (local < run.text.length) {
        after.push({ ...run, text: run.text.slice(local) });
      }
    }
    cursor = end;
  }
  return [before, after];
}

function styleAt(
  runs: readonly PdfTextStyleRun[],
  offset: number,
): Omit<PdfTextStyleRun, "text"> {
  const total = textFromPdfTextStyleRuns(runs).length;
  const target = total === 0 ? 0 : Math.max(0, Math.min(total - 1, offset));
  let cursor = 0;
  for (const run of runs) {
    if (target < cursor + run.text.length) {
      const style: Partial<PdfTextStyleRun> = { ...run };
      delete style.text;
      return style;
    }
    cursor += run.text.length;
  }
  return {};
}

/**
 * Reconcile run boundaries after one free-text edit.
 *
 * The longest common prefix/suffix identifies the replaced character range.
 * Untouched runs on both sides survive byte-for-byte. New characters inherit
 * the style immediately before the insertion (or the first surviving style at
 * the start), matching Word/Acrobat typing behavior.
 */
export function reconcilePdfTextStyleRuns(
  value: unknown,
  previousText: string,
  nextText: string,
): PdfTextStyleRun[] | undefined {
  const runs = normalizePdfTextStyleRuns(value, {
    expectedText: previousText,
    allowDisplayFont: true,
  });
  if (!runs) return undefined;
  if (previousText === nextText)
    return persistablePdfTextStyleRuns(runs, nextText);

  const { prefix, suffix } = pdfTextEditCommonEnds(previousText, nextText);

  const previousEnd = previousText.length - suffix;
  const nextEnd = nextText.length - suffix;
  const inserted = nextText.slice(prefix, nextEnd);
  const [left] = splitRunsAt(runs, prefix);
  const [, right] = splitRunsAt(runs, previousEnd);
  const inheritedOffset = prefix > 0 ? prefix - 1 : previousEnd;
  const inheritedStyle = styleAt(runs, inheritedOffset);
  const nextRuns = [
    ...left,
    ...(inserted ? [{ text: inserted, ...inheritedStyle }] : []),
    ...right,
  ];
  return persistablePdfTextStyleRuns(nextRuns, nextText);
}

/** Apply a whole-box style control without flattening unrelated run fields. */
/**
 * Apply a style patch to the characters in `[start, end)` ONLY, splitting runs
 * at the boundaries so the rest of the box keeps its own styling.
 *
 * This is what separates a text box from a paragraph: WPS and Acrobat both let
 * you bold one word inside a box, and until now a style change here always
 * covered every run — bolding a word bolded the sentence.
 *
 * The range is clamped to the text, and an empty or inverted range is a no-op
 * (returns the runs unchanged) rather than an error: a toolbar click with a
 * collapsed caret should do nothing, not throw.
 */
export function patchPdfTextStyleRunsRange(
  value: unknown,
  text: string,
  start: number,
  end: number,
  patch: Partial<Omit<PdfTextStyleRun, "text" | "displayFontFamily">>,
): PdfTextStyleRun[] | undefined {
  const runs = normalizePdfTextStyleRuns(value, {
    expectedText: text,
    allowDisplayFont: true,
  });
  if (!runs) return undefined;
  const from = Math.max(0, Math.min(text.length, Math.trunc(start)));
  const to = Math.max(0, Math.min(text.length, Math.trunc(end)));
  if (!(to > from)) return persistablePdfTextStyleRuns(runs, text);

  const out: PdfTextStyleRun[] = [];
  let cursor = 0;
  for (const run of runs) {
    const runStart = cursor;
    const runEnd = cursor + run.text.length;
    cursor = runEnd;
    if (runEnd <= from || runStart >= to) {
      out.push(run);
      continue;
    }
    // Head: the part of this run before the selection.
    if (runStart < from) {
      out.push({ ...run, text: run.text.slice(0, from - runStart) });
    }
    // Body: the covered part, carrying the patch.
    const bodyStart = Math.max(runStart, from) - runStart;
    const bodyEnd = Math.min(runEnd, to) - runStart;
    out.push({
      ...run,
      ...patch,
      displayFontFamily: undefined,
      text: run.text.slice(bodyStart, bodyEnd),
    });
    // Tail: the part after the selection.
    if (runEnd > to) {
      out.push({ ...run, text: run.text.slice(to - runStart) });
    }
  }
  return persistablePdfTextStyleRuns(out, text);
}

export function patchPdfTextStyleRuns(
  value: unknown,
  text: string,
  patch: Partial<Omit<PdfTextStyleRun, "text" | "displayFontFamily">>,
): PdfTextStyleRun[] | undefined {
  const runs = normalizePdfTextStyleRuns(value, {
    expectedText: text,
    allowDisplayFont: true,
  });
  if (!runs) return undefined;
  return persistablePdfTextStyleRuns(
    runs.map((run) => ({ ...run, ...patch, displayFontFamily: undefined })),
    text,
  );
}
