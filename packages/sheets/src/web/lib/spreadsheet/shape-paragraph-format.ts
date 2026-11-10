/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A sheet shape's paragraphs formatted as Excel formats a shape's text:
 * Bullets, Numbering and the indent level on the paragraphs the selection
 * touches (all of them for a shape selected whole). The paragraphs are the
 * shape's `textLayout` -- one per line of its words, the same paragraph model
 * a slide text box and an imported label hold -- so the label draws them and
 * the workbook writes them as `a:pPr` (lvl, marL, indent, buChar /
 * buAutoNum).
 */

import {
  normalizeOfficeTextLayout,
  type OfficeTextLayout,
  type OfficeTextParagraph,
} from "@mrnote/core/web/lib/office-text-layout.ts";

import type {
  PresentationParagraphDialogResult,
  PresentationParagraphDialogSource,
} from "@mrnote/core/web/lib/presentation-paragraph-dialog.ts";

import type { SpreadsheetShapeStyle } from "@mrnote/core/web/lib/spreadsheet/shapes.ts";

/** Excel's Bullets on a shape: the marker hangs 0.1875 in (171450 EMU). */
const BULLET_HANG_EMU = 171_450;
/** One indent level: 0.5 in, as Office's text levels step. */
const LEVEL_STEP_EMU = 457_200;
/** The row model every host shares (and the page validator) stops at level 4. */
const MAX_LEVEL = 4;

/** The line indexes `[first, last]` a character range of `text` touches. */
function linesOf(text: string, range: { start: number; end: number } | null): [number, number] {
  const lines = text.split("\n");
  if (!range) return [0, lines.length - 1];
  let offset = 0;
  let first = -1;
  let last = lines.length - 1;
  for (let index = 0; index < lines.length; index += 1) {
    const end = offset + lines[index].length;
    if (first < 0 && range.start <= end) first = index;
    if (range.end <= end) {
      last = index;
      break;
    }
    offset = end + 1;
  }
  return [Math.max(0, first), Math.max(Math.max(0, first), last)];
}

function layoutFor(shape: SpreadsheetShapeStyle): OfficeTextLayout {
  const lines = (shape.text ?? "").split("\n").length;
  const own = normalizeOfficeTextLayout(shape.textLayout);
  if (own && own.paragraphs.length === lines) return { ...own, paragraphs: own.paragraphs.map((p) => ({ ...p })) };
  return {
    ...(own?.bodyPr ? { bodyPr: own.bodyPr } : {}),
    paragraphs: Array.from({ length: lines }, () => ({ level: 0, bullet: "none" as const })),
  };
}

function withMargins(paragraph: OfficeTextParagraph): OfficeTextParagraph {
  const next: OfficeTextParagraph = { ...paragraph };
  const hang = next.bullet === "none" ? 0 : BULLET_HANG_EMU;
  const margin = next.level * LEVEL_STEP_EMU + hang;
  if (margin > 0) next.marL = margin;
  else delete next.marL;
  if (hang > 0) next.indent = -hang;
  else delete next.indent;
  return next;
}

function withLayout(shape: SpreadsheetShapeStyle, layout: OfficeTextLayout): SpreadsheetShapeStyle {
  const plain = layout.paragraphs.every(
    (paragraph) =>
      paragraph.bullet === "none" &&
      paragraph.level === 0 &&
      !paragraph.align &&
      !paragraph.marL &&
      !paragraph.indent &&
      paragraph.lineSpacing === undefined &&
      !paragraph.spaceBefore &&
      !paragraph.spaceAfter,
  );
  const next = { ...shape } as SpreadsheetShapeStyle;
  if (plain && !layout.bodyPr) delete (next as { textLayout?: unknown }).textLayout;
  else next.textLayout = layout;
  return next;
}

/**
 * Bullets / Numbering on the paragraphs `range` touches (Excel's toggle: on
 * when any of them lacks the kind, off when all have it). Null when the
 * shape holds no words.
 */
export function spreadsheetShapeWithList(
  shape: SpreadsheetShapeStyle,
  kind: "bullet" | "number",
  range: { start: number; end: number } | null,
): SpreadsheetShapeStyle | null {
  if (!(shape.text ?? "").length) return null;
  const layout = layoutFor(shape);
  const [first, last] = linesOf(shape.text ?? "", range);
  const touched = layout.paragraphs.slice(first, last + 1);
  const off = touched.every((paragraph) => paragraph.bullet === kind);
  layout.paragraphs = layout.paragraphs.map((paragraph, index) => {
    if (index < first || index > last) return paragraph;
    const next: OfficeTextParagraph = { ...paragraph, bullet: off ? "none" : kind };
    delete next.bulletChar;
    delete next.numberFormat;
    delete next.nativeBullet;
    if (!off && kind === "bullet") next.bulletChar = "•";
    if (!off && kind === "number") next.numberFormat = "arabicPeriod";
    return withMargins(next);
  });
  return withLayout(shape, layout);
}

/** Increase / Decrease Indent on the paragraphs `range` touches. */
export function spreadsheetShapeWithIndent(
  shape: SpreadsheetShapeStyle,
  delta: 1 | -1,
  range: { start: number; end: number } | null,
): SpreadsheetShapeStyle | null {
  if (!(shape.text ?? "").length) return null;
  const layout = layoutFor(shape);
  const [first, last] = linesOf(shape.text ?? "", range);
  let changed = false;
  layout.paragraphs = layout.paragraphs.map((paragraph, index) => {
    if (index < first || index > last) return paragraph;
    const level = Math.max(0, Math.min(MAX_LEVEL, paragraph.level + delta));
    if (level === paragraph.level) return paragraph;
    changed = true;
    return withMargins({ ...paragraph, level });
  });
  return changed ? withLayout(shape, layout) : null;
}

/** One value the touched paragraphs share, "mixed" where they differ. */
function shared<T>(values: readonly T[]): T | "mixed" {
  return values.every((value) => value === values[0]) ? values[0] : "mixed";
}

/**
 * The Paragraph dialog's starting state for the paragraphs `range` touches
 * (Excel's shape Paragraph dialog is PowerPoint's). Null without words.
 */
export function spreadsheetShapeParagraphSource(
  shape: SpreadsheetShapeStyle,
  range: { start: number; end: number } | null,
  fallbackAlign: "left" | "center",
): PresentationParagraphDialogSource | null {
  if (!(shape.text ?? "").length) return null;
  const layout = layoutFor(shape);
  const [first, last] = linesOf(shape.text ?? "", range);
  const touched = layout.paragraphs.slice(first, last + 1);
  // The shape's own alignment (Home ▸ Alignment on it) is drawn over the rows'.
  const own = shape.textAlign;
  return {
    align: shared(touched.map((paragraph) => own ?? paragraph.align ?? fallbackAlign)),
    lineSpacing: shared(touched.map((paragraph) => paragraph.lineSpacing)),
    lineSpacingPoints: undefined,
    spaceBefore: shared(touched.map((paragraph) => paragraph.spaceBefore)),
    spaceAfter: shared(touched.map((paragraph) => paragraph.spaceAfter)),
    marL: touched[0]?.marL ?? 0,
    indent: touched[0]?.indent ?? 0,
  };
}

/**
 * The Paragraph dialog's OK on the paragraphs `range` touches: alignment,
 * Before text / Special, Spacing Before / After and Line Spacing (Single,
 * 1.5 lines, Double, Multiple). Only what the person changed is written.
 */
export function spreadsheetShapeWithParagraphFormat(
  shape: SpreadsheetShapeStyle,
  result: PresentationParagraphDialogResult,
  range: { start: number; end: number } | null,
): SpreadsheetShapeStyle | null {
  if (!(shape.text ?? "").length) return null;
  const layout = layoutFor(shape);
  const [first, last] = linesOf(shape.text ?? "", range);
  const { patch, indentation } = result;
  // The shape's own alignment is drawn over every row's: a paragraph's own
  // alignment takes over from it, the other rows keeping what they drew.
  const own = patch.align ? shape.textAlign : undefined;
  if (own) {
    layout.paragraphs = layout.paragraphs.map((paragraph) => ({ ...paragraph, align: own }));
  }
  layout.paragraphs = layout.paragraphs.map((paragraph, index) => {
    if (index < first || index > last) return paragraph;
    const next: OfficeTextParagraph = { ...paragraph };
    if (patch.align) next.align = patch.align;
    if ("lineSpacing" in patch) {
      if (patch.lineSpacing === undefined || patch.lineSpacing === 1) delete next.lineSpacing;
      else next.lineSpacing = patch.lineSpacing;
    }
    for (const key of ["spaceBefore", "spaceAfter"] as const) {
      if (!(key in patch)) continue;
      const value = patch[key];
      if (value === undefined || value === 0) delete next[key];
      else next[key] = value;
    }
    if (indentation?.marL !== undefined) {
      if (indentation.marL > 0) next.marL = indentation.marL;
      else delete next.marL;
    }
    if (indentation && "indent" in indentation) {
      if (indentation.indent) next.indent = indentation.indent;
      else delete next.indent;
    }
    return next;
  });
  const next = withLayout(shape, layout);
  if (own) delete (next as { textAlign?: unknown }).textAlign;
  return next;
}
