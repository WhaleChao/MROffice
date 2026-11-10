/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { equationBoxSizeFor } from "@mrnote/slides/web/lib/presentation-equation-box.ts";
import {
  anchorForObjectBounds,
  anchorWithDefaults,
  anchorWithExplicitSize,
  EMU_PER_CSS_PIXEL,
  objectBoundsPixels,
  type DrawingGrid,
  type SpreadsheetObjectAnchor,
} from "./object-buttons";
import { MAX_SHAPE_TEXT_SIZE, MIN_SHAPE_TEXT_SIZE, type SpreadsheetShapeStyle } from "@mrnote/core/web/lib/spreadsheet/shapes.ts";
import {
  normalizeOfficeTextLayout,
  type OfficeTextBodyProperties,
  type OfficeTextLayout,
} from "@mrnote/core/web/lib/office-text-layout.ts";

/** The size the box's equation is drawn at (the label's own text size). */
/** Excel's equation box insets: none (the deck's are `EQUATION_BOX_TEXT_INSETS`). */
export const SPREADSHEET_EQUATION_BOX_INSETS = { left: 0, top: 0, right: 0, bottom: 0 } as const;

/** The size (EMU) an equation box needs for its equation as drawn (`contentPx`, layout px). */
export function spreadsheetEquationBoxSizeFor(contentPx: {
  readonly width: number;
  readonly height: number;
}): { readonly width: number; readonly height: number } {
  return equationBoxSizeFor(contentPx, SPREADSHEET_EQUATION_BOX_INSETS);
}

export function spreadsheetEquationFontSizePx(shape: Pick<SpreadsheetShapeStyle, "textSize">): number {
  return Math.min(MAX_SHAPE_TEXT_SIZE, Math.max(MIN_SHAPE_TEXT_SIZE, shape.textSize ?? 12));
}

/**
 * `anchor` at `size` (EMU), its top-left and its placement kept: a box that
 * moves and sizes with cells (Excel's text box, an equation box included)
 * stays a two-cell anchor, its far marker moved to the new corner on `grid`.
 */
export function spreadsheetEquationAnchorAtSize(
  anchor: SpreadsheetObjectAnchor | undefined,
  size: { readonly width: number; readonly height: number },
  grid?: DrawingGrid,
): SpreadsheetObjectAnchor {
  const extent = { widthEmu: Math.max(1, Math.round(size.width)), heightEmu: Math.max(1, Math.round(size.height)) };
  if (anchor?.kind === "absolute") return { ...anchor, ...extent };
  if (anchor?.kind === "twoCell" && anchor.to && grid) {
    const bounds = objectBoundsPixels({ id: "", anchor }, grid.colWidth, grid.rowHeight);
    const resized = anchorForObjectBounds(
      anchor,
      { ...bounds, width: extent.widthEmu / EMU_PER_CSS_PIXEL, height: extent.heightEmu / EMU_PER_CSS_PIXEL },
      grid.rowCount,
      grid.columnCount,
      grid.colWidth,
      grid.rowHeight,
    );
    return { ...resized, from: anchor.from };
  }
  return anchorWithExplicitSize(anchorWithDefaults(anchor), extent);
}

/** The box's anchor for an equation drawn at `contentPx` (its insets added), its top-left kept. */
export function spreadsheetEquationBoxAnchor(
  anchor: SpreadsheetObjectAnchor | undefined,
  contentPx: { readonly width: number; readonly height: number },
  grid?: DrawingGrid,
): SpreadsheetObjectAnchor {
  return spreadsheetEquationAnchorAtSize(anchor, spreadsheetEquationBoxSizeFor(contentPx), grid);
}

/**
 * The equation a text box IS, or null: the deck's equation box model (a text
 * box whose words are one display equation run, `TextRun.math`, as Excel's
 * Insert > Equation box is one `m:oMathPara`), or a page written before it
 * (`shape.equation`, read as it was; it becomes the run on its next edit).
 */
export function spreadsheetEquationBoxLatex(
  shape: Pick<SpreadsheetShapeStyle, "kind" | "text" | "textRuns" | "equation"> | null | undefined,
): string | null {
  if (!shape || shape.kind !== "text-box") return null;
  if (shape.equation) return shape.equation.latex;
  const runs = shape.textRuns;
  const only = runs?.length === 1 ? runs[0] : undefined;
  return only?.math && only.math.display === true && only.text && only.text === shape.text ? only.math.latex : null;
}

/** `shape` as an equation box holding `equation` (the deck's model: one display run). */
export function spreadsheetEquationBoxShape<S extends SpreadsheetShapeStyle>(
  shape: S,
  equation: { readonly text: string; readonly latex: string },
): S {
  const next = { ...shape, text: equation.text, textRuns: [{ text: equation.text, math: { latex: equation.latex, display: true } }] } as S;
  delete (next as { equation?: unknown }).equation;
  return withSpreadsheetEquationBoxBody(next);
}







export const SPREADSHEET_EQUATION_BOX_BODY: Readonly<OfficeTextBodyProperties> = Object.freeze({
  anchor: "t",
  lIns: 0,
  tIns: 0,
  rIns: 0,
  bIns: 0,
  wrap: "none",
});

/**
 * `shape` with the equation box's own text body: Resize shape to fit text
 * (`textAutofit`) and the frame (`textLayout.bodyPr`, the envelope a file's
 * `a:bodyPr` is read into), a body the shape already states winning over the
 * equation box's defaults (an Excel file's own box). The frame rides on the
 * words' layout, which exists only while there are words; an empty box keeps
 * `textAutofit` and takes the frame with its first equation.
 */
export function withSpreadsheetEquationBoxBody<S extends SpreadsheetShapeStyle>(shape: S): S {
  const text = shape.text ?? "";
  if (!text) return { ...shape, textAutofit: shape.textAutofit ?? "shape" };
  const lines = text.split("\n").length;
  const own = normalizeOfficeTextLayout(shape.textLayout);
  const layout: OfficeTextLayout =
    own && own.paragraphs.length === lines
      ? own
      : { paragraphs: Array.from({ length: lines }, () => ({ level: 0, bullet: "none" as const })) };
  const bodyPr = { ...SPREADSHEET_EQUATION_BOX_BODY, ...(layout.bodyPr ?? {}) };
  // A body its file says wraps is that file's text box, not this one.
  if (bodyPr.wrap !== "none") return shape;
  return { ...shape, textAutofit: shape.textAutofit ?? "shape", textLayout: { ...layout, bodyPr } };
}

/**
 * Whether a shape's frame follows its words in both directions: Resize shape
 * to fit text on a body that does not wrap (an equation box, with or without
 * words beside its equation). Its one line sets its width, its lines its
 * height, its top-left kept (`a:spAutoFit`, ECMA-376: "the shape actually
 * resizes" as text is typed). A wrapped autofit body (height only) is not
 * drawn this way.
 */
export function spreadsheetShapeResizesToFitText(
  shape: Pick<SpreadsheetShapeStyle, "textAutofit" | "textLayout"> | null | undefined,
): boolean {
  return shape?.textAutofit === "shape" && normalizeOfficeTextLayout(shape.textLayout)?.bodyPr?.wrap === "none";
}
