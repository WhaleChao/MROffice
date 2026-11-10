/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure model + placement helpers for CREATING a brand-new PDF text box on a
 * blank area of a page (WPS / Acrobat "Add Text"). Unlike an edited-existing box
 * — which mirrors an extracted glyph run and pins a cover to that run's origin —
 * a created box has NO underlying glyphs, so its origin rect equals its own rect
 * (the flatten mask then paints the sampled page background only over the box's
 * own footprint, which is a no-op on already-blank paper).
 *
 * Geometry is page-relative FRACTIONS (0..1) to match the rest of the inline
 * editor. Framework-free so the model + clamp are unit-testable without a DOM.
 */

import { MIN_BOX_FRACTION } from "./text-box-geometry";

/**
 * The created-box model. A structural superset of the editor's persisted edit
 * shape (id / page / rect / text / font), plus `created: true` to distinguish a
 * user-authored box from an edited-existing one, and the `origin*` rect (here
 * equal to the box rect) so the shared flatten/cover path treats it uniformly.
 */
export interface CreatedTextBox {
  id: string;
  page: number;
  x: number;
  y: number;
  width: number;
  height: number;
  text: string;
  fontSize: number;
  color?: string;
  align: "left" | "center" | "right";
  bold: boolean;
  italic: boolean;
  underline: boolean;
  /** Marks a user-authored box (vs. an edited extracted line). */
  created: true;
  /** Origin === own rect: a new box has no source glyphs to keep covered. */
  originX: number;
  originY: number;
  originWidth: number;
  originHeight: number;
}

/** Default size (page fractions) + font of a freshly placed box. */
export const DEFAULT_NEW_BOX_WIDTH = 0.28;
export const DEFAULT_NEW_BOX_HEIGHT = 0.035;
/** Default font size as a fraction of page height (~14pt on Letter). */
export const DEFAULT_NEW_BOX_FONT_SIZE = 0.018;

function clamp(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) return min;
  return Math.min(max, Math.max(min, value));
}

export interface NewBoxPlacement {
  x: number;
  y: number;
  width: number;
  height: number;
}

/**
 * Place a box of the given size so its TOP-LEFT sits at (x, y) but the whole box
 * stays inside the page [0,1]. The size is clamped to a sane min and to the page;
 * then the top-left is clamped so right/bottom never exceed 1. Pure — drives both
 * the create model and any future drag-to-create rubber-band.
 */
export function clampNewBoxPlacement(
  x: number,
  y: number,
  width = DEFAULT_NEW_BOX_WIDTH,
  height = DEFAULT_NEW_BOX_HEIGHT,
): NewBoxPlacement {
  const w = clamp(width, MIN_BOX_FRACTION, 1);
  const h = clamp(height, MIN_BOX_FRACTION, 1);
  return {
    x: clamp(x, 0, 1 - w),
    y: clamp(y, 0, 1 - h),
    width: w,
    height: h,
  };
}

export interface CreateTextBoxOptions {
  /** Zero-based page index the box belongs to. */
  page: number;
  /** Unique id for the box (caller supplies; deterministic for tests). */
  id: string;
  /** Box size in page fractions (defaults applied + clamped). */
  width?: number;
  height?: number;
  /** Font size as a fraction of page height. */
  fontSize?: number;
  /** Initial text (defaults to ""). */
  text?: string;
  color?: string;
  align?: "left" | "center" | "right";
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
}

/**
 * Build a {@link CreatedTextBox} for a click at fractional (clickX, clickY). The
 * click is treated as the box's TOP-LEFT and the box is clamped on-page. The
 * origin rect equals the placed rect (no source glyphs). Empty/whitespace text
 * is allowed — the editor focuses the box for the user to type.
 */
export function createTextBox(
  clickX: number,
  clickY: number,
  opts: CreateTextBoxOptions,
): CreatedTextBox {
  const placed = clampNewBoxPlacement(
    clickX,
    clickY,
    opts.width ?? DEFAULT_NEW_BOX_WIDTH,
    opts.height ?? DEFAULT_NEW_BOX_HEIGHT,
  );
  const fontSize =
    Number.isFinite(opts.fontSize) && (opts.fontSize as number) > 0
      ? (opts.fontSize as number)
      : DEFAULT_NEW_BOX_FONT_SIZE;
  return {
    id: opts.id,
    page: Number.isFinite(opts.page) ? Math.max(0, Math.trunc(opts.page)) : 0,
    x: placed.x,
    y: placed.y,
    width: placed.width,
    height: placed.height,
    text: typeof opts.text === "string" ? opts.text : "",
    fontSize,
    color: typeof opts.color === "string" ? opts.color : undefined,
    align: opts.align === "center" || opts.align === "right" ? opts.align : "left",
    bold: Boolean(opts.bold),
    italic: Boolean(opts.italic),
    underline: Boolean(opts.underline),
    created: true,
    originX: placed.x,
    originY: placed.y,
    originWidth: placed.width,
    originHeight: placed.height,
  };
}

/** What the armed text tool places: a text box, one set as a formula
 * (插入 › 公式), or one painted with a WordArt style (插入 › 艺术字, its
 * gallery recipe's id). */
export type PdfArmedTextKind =
  | "text"
  | "formula"
  | `wordart:${string}`
  | `table:${number}x${number}`
  | `chart:${string}:${string}`
  | `smartart:${string}`;

/** The family and subtype of an armed chart (图表), or null. */
export function pdfArmedChartLeaf(kind: string): { family: string; subtype: string } | null {
  const match = /^chart:([A-Za-z]{1,40}):([A-Za-z0-9-]{1,60})$/.exec(kind);
  return match ? { family: match[1], subtype: match[2] } : null;
}

/** The layout (a catalog id) of an armed SmartArt (智能图形), or null. */
export function pdfArmedSmartArtLayout(kind: string): string | null {
  const match = /^smartart:([a-z0-9-]{1,80})$/.exec(kind);
  return match ? match[1] : null;
}

/** The rows and columns of an armed table (表格), or null. */
export function pdfArmedTableSize(kind: string): { rows: number; columns: number } | null {
  const match = /^table:(\d{1,3})x(\d{1,3})$/.exec(kind);
  return match ? { rows: Number(match[1]), columns: Number(match[2]) } : null;
}

/** WordArt's own size: Word sets its placeholder words at 36 pt. */
export const PDF_WORDART_FONT_SIZE_PT = 36;
