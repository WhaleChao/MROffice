/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Per-text-box / shape VERTICAL text anchor (PowerPoint / WPS "Align Text" →
 * Top / Middle / Bottom parity).
 *
 * A text block, free text box, or shape label may anchor its text to the TOP,
 * MIDDLE, or BOTTOM of its box. This mirrors the spreadsheet cell
 * `verticalAlign` model (top/middle/bottom) so the same three options the user
 * already knows from sheets apply to slide text.
 *
 * Pure, side-effect-free helpers shared by the presentation schema, the editor
 * canvas, and the read-only present-mode overlay so the allowed set, the default
 * (TOP, matching the historical top-aligned render), and the render-ready CSS
 * live in exactly one place. Keep this module dependency-free (no React / no DOM)
 * so it can be unit tested in isolation and imported from the zod schema.
 *
 * Everything here is ADDITIVE and OPTIONAL: a block with no anchor resolves to
 * `top`, which renders byte-for-byte the same as a block did before this feature
 * (the wrapper was effectively top-anchored), so existing decks are unchanged.
 */

import type { CSSProperties } from "react";
import { writingModeCss, type WritingMode } from "./presentation-line-spacing";

/** The supported vertical anchors. `top` is the implicit (historical) default. */
export const VERTICAL_ANCHORS = ["top", "middle", "bottom"] as const;

export type VerticalAnchor = (typeof VERTICAL_ANCHORS)[number];

/** The anchor used when a block omits (or mis-specifies) one. */
export const DEFAULT_VERTICAL_ANCHOR: VerticalAnchor = "top";

const VERTICAL_ANCHOR_SET: ReadonlySet<string> = new Set(VERTICAL_ANCHORS);

/** Type guard: is `value` exactly one of the allowed anchors? */
export function isVerticalAnchor(value: unknown): value is VerticalAnchor {
  return typeof value === "string" && VERTICAL_ANCHOR_SET.has(value);
}

/**
 * Normalize an arbitrary persisted value into an allowed anchor. Unknown /
 * empty / non-string values collapse to {@link DEFAULT_VERTICAL_ANCHOR} so a
 * stale or hand-edited deck can never inject an unsupported value. Common
 * synonyms (`center`/`centre` → `middle`, `start` → `top`, `end` → `bottom`)
 * are accepted; the comparison is case-insensitive and trims whitespace.
 */
export function normalizeVerticalAnchor(value: unknown): VerticalAnchor {
  if (typeof value !== "string") return DEFAULT_VERTICAL_ANCHOR;
  const candidate = value.trim().toLowerCase();
  if (candidate === "center" || candidate === "centre") return "middle";
  if (candidate === "start" || candidate === "flex-start") return "top";
  if (candidate === "end" || candidate === "flex-end") return "bottom";
  return isVerticalAnchor(candidate) ? candidate : DEFAULT_VERTICAL_ANCHOR;
}

/**
 * Read a block-style record's optional `verticalAlign` field and normalize it.
 * Mirrors the spreadsheet cell `verticalAlign` accessor — a missing / invalid
 * value yields the default {@link DEFAULT_VERTICAL_ANCHOR}.
 */
export function readVerticalAnchor(
  style: { verticalAlign?: unknown } | null | undefined,
): VerticalAnchor {
  return normalizeVerticalAnchor(style?.verticalAlign);
}













export function shapeLabelVerticalAlign(
  style: { verticalAlign?: unknown } | null | undefined,
  source: unknown,
): VerticalAnchor | undefined {
  if (style?.verticalAlign !== undefined) {
    return isVerticalAnchor(style.verticalAlign) ? style.verticalAlign : undefined;
  }
  const layoutStyle =
    source && typeof source === "object" && !Array.isArray(source)
      ? (source as { layoutStyle?: unknown }).layoutStyle
      : undefined;
  return layoutStyle && typeof layoutStyle === "object" && !Array.isArray(layoutStyle)
    ? "top"
    : undefined;
}

/**
 * The CSS `flex` `justify-content` value that anchors a single text child to the
 * top / middle / bottom of a `display: flex; flex-direction: column` wrapper.
 */
export function verticalAnchorJustify(
  anchor: VerticalAnchor,
): "flex-start" | "center" | "flex-end" {
  if (anchor === "middle") return "center";
  if (anchor === "bottom") return "flex-end";
  return "flex-start";
}

/**
 * Render-ready CSS that turns a block wrapper into a vertical flex column and
 * anchors its (single) text child to the resolved position. Returns an EMPTY
 * object for the default `top` anchor so spreading it onto the historical
 * wrapper is a no-op — an un-anchored block renders byte-for-byte as before.
 *
 * Vertical text applies its writing mode to this wrapper too: the column
 * axis then follows DrawingML's block axis while paragraph alignment handles
 * the inline axis. This keeps rotated/vertical text centered across its box.
 *
 * The non-default anchors layer `display: flex` + `flex-direction: column` +
 * `justify-content`, which positions the child vertically WITHOUT changing the
 * child's own width / horizontal text alignment.
 */
export function verticalAnchorWrapperCss(
  anchor: VerticalAnchor,
  writingMode?: WritingMode,
): CSSProperties {
  const direction = writingModeCss(writingMode);
  if (anchor === DEFAULT_VERTICAL_ANCHOR && !direction.writingMode) return {};
  return {
    display: "flex",
    flexDirection: "column",
    justifyContent: verticalAnchorJustify(anchor),
    ...direction,
  };
}

/**
 * The `verticalAlign` value to PERSIST for a block style, or `undefined` when
 * the anchor is the default `top` (so an un-anchored block omits the field and
 * stays byte-clean). The editor write path uses this to drop a redundant
 * `verticalAlign: "top"`.
 */
export function persistedVerticalAnchor(
  value: unknown,
): VerticalAnchor | undefined {
  const anchor = normalizeVerticalAnchor(value);
  return anchor === DEFAULT_VERTICAL_ANCHOR ? undefined : anchor;
}

/**
 * `a:bodyPr@anchorCtr` as an imported slide block states it
 * (`style.anchorCenter`, hoisted from the block's top level at load). Stated
 * only when true; absent is the default. Drawn by the shared
 * `officeAnchorCenterBoundsCss` (`office-text-layout.ts`), which the sheet's
 * shape text uses too.
 */
export function readAnchorCenter(
  style: { anchorCenter?: unknown } | null | undefined,
): boolean {
  return style?.anchorCenter === true;
}
