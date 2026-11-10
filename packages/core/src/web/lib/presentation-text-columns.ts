/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Multi-column TEXT BODY layout for a presentation text block / placeholder
 * (PowerPoint / WPS "Format Shape → Text Options → Columns").
 *
 * A text block may flow its body across 1–8 newspaper-style columns. The
 * column count is an ADDITIVE, OPTIONAL field (`columnCount`): a block that never
 * picked one omits it and renders exactly as before (one column, byte-for-byte
 * unchanged). The render edge projects the resolved count to CSS multi-column
 * (`column-count` + a gap); a single column resolves to an empty style object so
 * spreading it is a no-op.
 *
 * Pure, side-effect-free helpers shared by the presentation schema, the editor
 * canvas, and the present-mode overlay so the allowed range + render-ready CSS
 * live in one place. Keep this module dependency-free (no React / no DOM) so it
 * can be unit tested in isolation and imported from the zod schema.
 */

import type { CSSProperties } from "react";

/** The supported body column counts: PowerPoint's Columns dialog takes 1 to 16. */
export const MIN_TEXT_COLUMNS = 1;
export const MAX_TEXT_COLUMNS = 16;
/** Columns > Spacing between columns: 0" to 56" (PowerPoint's dialog), in EMU. */
export const MAX_TEXT_COLUMN_SPACING_EMU = 56 * 914_400;

/** The implicit default when a block carries no (or an invalid) column count. */
export const DEFAULT_TEXT_COLUMNS = 1;

/**
 * Normalize an arbitrary persisted value into an allowed column count in
 * [1, 8]. A non-integer is rounded; out-of-range values clamp to the nearest
 * bound; a missing / non-numeric / non-finite value collapses to the default
 * (1) so a stale or hand-edited deck can never request an unsupported count.
 */
export function normalizeColumnCount(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return DEFAULT_TEXT_COLUMNS;
  }
  const rounded = Math.round(value);
  if (rounded < MIN_TEXT_COLUMNS) return MIN_TEXT_COLUMNS;
  if (rounded > MAX_TEXT_COLUMNS) return MAX_TEXT_COLUMNS;
  return rounded;
}

/**
 * The CSS to layer onto the text-body CONTENT wrapper for the resolved column
 * count. A single column (the default) returns an EMPTY object so the historical
 * render stays byte-stable; 2–8 columns return `columnCount` + a comfortable
 * gap so the body text flows newspaper-style. `columnFill: "balance"` keeps the
 * columns roughly even (PowerPoint balances column content).
 */
export function textColumnsCss(
  value: unknown,
  native?: unknown,
  spacingEmu?: unknown,
): CSSProperties {
  const count = normalizeColumnCount(value);
  if (count <= 1) return {};
  const stated = normalizeColumnSpacing(spacingEmu);
  const imported =
    stated !== undefined ? stated / EMU_PER_PX : nativeTextColumnSpacingPx(native);
  // Right-to-left column order (`a:bodyPr@rtlCol`): the first column is the
  // rightmost and overflow moves left. It "in no way determines the direction
  // of text" (ECMA-376 §21.1.2.1.1), so the rows inside state their own
  // (`textColumnsRowDirection`).
  const order: CSSProperties = nativeTextColumnsRightToLeft(native) ? { direction: "rtl" } : {};
  if (imported !== undefined) {
    // An imported body's columns as PowerPoint lays them out
    // (`a:bodyPr@numCol/@spcCol`, ECMA-376 §21.1.2.1.1): the stated gap, and
    // each column filled before the next ("overflow containers"); a column
    // box with no height of its own still balances (CSS multicol).
    return {
      columnCount: count,
      columnGap: `${imported}px`,
      columnFill: "auto",
      ...order,
    };
  }
  return {
    columnCount: count,
    columnGap: "1.5em",
    columnFill: "balance",
    ...order,
  };
}

/**
 * Whether an imported body fills its columns right to left
 * (`source.nativeTextColumns.rightToLeft`, the import's `a:bodyPr@rtlCol`;
 * reading row 477). A Hebrew or Arabic PowerPoint states it on one-column
 * bodies too, so Columns picked in the editor flow in that order.
 */
export function nativeTextColumnsRightToLeft(native: unknown): boolean {
  return Boolean(native && typeof native === "object" && (native as { rightToLeft?: unknown }).rightToLeft === true);
}

/**
 * The base direction a body's rows state inside columns ordered right to
 * left: the column container's `direction: rtl` orders the columns only, so
 * every row says its own -- the box's (`dir`), else left to right. Undefined
 * (rows inherit, as before) for any other body.
 */
export function textColumnsRowDirection(
  columns: CSSProperties,
  boxDirection: unknown,
): "ltr" | "rtl" | undefined {
  if (columns.direction !== "rtl") return undefined;
  return boxDirection === "rtl" ? "rtl" : "ltr";
}

/** EMU per CSS px on the 96-dpi slide the canvas lays out. */
const EMU_PER_PX = 9525;

/**
 * The block's own Spacing between columns (`columnSpacing`, EMU, the
 * `a:bodyPr@spcCol` the writers put out; PPT-A042), or undefined.
 */
export function normalizeColumnSpacing(value: unknown): number | undefined {
  if (typeof value !== "number" || !Number.isInteger(value)) return undefined;
  if (value < 0 || value > MAX_TEXT_COLUMN_SPACING_EMU) return undefined;
  return value;
}

/** The spacing the box draws with: its own, else the imported body's, else 0. */
export function effectiveColumnSpacingEmu(block: {
  readonly columnSpacing?: unknown;
  readonly source?: unknown;
}): number {
  const stated = normalizeColumnSpacing(block.columnSpacing);
  if (stated !== undefined) return stated;
  const native = (block.source as { nativeTextColumns?: unknown } | undefined)?.nativeTextColumns;
  const px = nativeTextColumnSpacingPx(native);
  return px === undefined ? 0 : Math.round(px * EMU_PER_PX);
}

/**
 * The gap, in CSS px, an imported body states between its columns
 * (`source.nativeTextColumns.spacingEmu`, the import's `a:bodyPr@spcCol`),
 * or undefined for a block that is not an imported multi-column body.
 */
export function nativeTextColumnSpacingPx(native: unknown): number | undefined {
  if (!native || typeof native !== "object") return undefined;
  const spacing = (native as { spacingEmu?: unknown }).spacingEmu;
  if (typeof spacing !== "number" || !Number.isFinite(spacing) || spacing < 0) {
    return undefined;
  }
  return spacing / EMU_PER_PX;
}

/**
 * The column count to PERSIST for a text block: `undefined` for the default
 * single column (so an un-columned box serializes byte-for-byte as before), else
 * the normalized 2–8. Mirrors the "omit defaults" persistence the other text
 * features use.
 */
export function persistedColumnCount(value: unknown): number | undefined {
  const count = normalizeColumnCount(value);
  return count <= DEFAULT_TEXT_COLUMNS ? undefined : count;
}

/**
 * A vertical anchor for a box that ALSO flows in columns. The anchor's flex
 * column turned the box into a flex container, where `column-count` does
 * not apply, so Align Text > Middle put two columns back into one (PPT-B116).
 * A multi-column box keeps its block layout and is anchored with
 * `align-content` instead, which a multicol container takes; a
 * single-column box keeps the flex anchor unchanged.
 */
export function anchorCssForColumns(
  anchor: CSSProperties,
  columns: CSSProperties,
): CSSProperties {
  if (typeof columns.columnCount !== "number" || columns.columnCount <= 1) {
    return anchor;
  }
  if (anchor.display !== "flex") return anchor;
  const {
    display: _display,
    flexDirection: _direction,
    justifyContent,
    ...rest
  } = anchor;
  return {
    ...rest,
    alignContent:
      justifyContent === "center"
        ? "center"
        : justifyContent === "flex-end"
          ? "end"
          : "start",
  };
}
