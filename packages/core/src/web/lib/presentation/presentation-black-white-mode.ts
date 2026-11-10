/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Grayscale ▸ Change Selected Object (PPT-A109): how one object renders in
 * the Grayscale and Pure Black and White views (and when a deck is printed
 * without colour). PowerPoint keeps it on the object as `p:spPr@bwMode`
 * (ISO/IEC 29500-1 §20.1.10.10 ST_BlackWhiteMode); the deck's colours never
 * change. Stored on the block as `blackWhiteMode`; Automatic is the absence.
 *
 * The editor previews a mode as a filter over the object, on top of the
 * view's own filter: an approximation of PowerPoint's rendering, which maps
 * an object's fill and line separately (a CSS filter cannot), close enough to
 * tell the modes apart and to hide what Don't Show hides.
 */

import type { PresentationViewColorMode } from "./presentation-view-color-mode";

/** The modes of the gallery, in PowerPoint's order; `auto` is Automatic. */
export const PRESENTATION_BLACK_WHITE_MODES = [
  "auto",
  "gray",
  "ltGray",
  "invGray",
  "grayWhite",
  "blackGray",
  "blackWhite",
  "black",
  "white",
  "hidden",
] as const;
export type PresentationBlackWhiteMode = (typeof PRESENTATION_BLACK_WHITE_MODES)[number];

/** The modes a block stores (`blackWhiteMode`): all but Automatic, which is the absence. */
export const PRESENTATION_STORED_BLACK_WHITE_MODES = [
  "gray",
  "ltGray",
  "invGray",
  "grayWhite",
  "blackGray",
  "blackWhite",
  "black",
  "white",
  "hidden",
] as const satisfies readonly Exclude<PresentationBlackWhiteMode, "auto">[];

/** A stored mode, or undefined for Automatic and anything unknown. */
export function normalizePresentationBlackWhiteMode(
  value: unknown,
): Exclude<PresentationBlackWhiteMode, "auto"> | undefined {
  return typeof value === "string" &&
    (PRESENTATION_STORED_BLACK_WHITE_MODES as readonly string[]).includes(value)
    ? (value as Exclude<PresentationBlackWhiteMode, "auto">)
    : undefined;
}

const MODE_FILTERS: Readonly<Record<Exclude<PresentationBlackWhiteMode, "auto" | "hidden">, string>> = {
  gray: "grayscale(1)",
  ltGray: "grayscale(1) brightness(1.35)",
  invGray: "grayscale(1) invert(1)",
  grayWhite: "grayscale(1) contrast(0.6) brightness(1.5)",
  blackGray: "grayscale(1) contrast(1.6) brightness(0.65)",
  blackWhite: "grayscale(1) contrast(1000%)",
  black: "brightness(0)",
  white: "brightness(0) invert(1)",
};

/**
 * The style a block's frame takes for its mode in the current view: none in
 * Color view or for Automatic (the view's own filter renders it).
 */
export function presentationBlackWhiteStyle(
  viewMode: PresentationViewColorMode,
  block: { blackWhiteMode?: unknown },
): { filter?: string; visibility?: "hidden" } | undefined {
  if (viewMode === "color") return undefined;
  const mode = normalizePresentationBlackWhiteMode(block.blackWhiteMode);
  if (!mode) return undefined;
  if (mode === "hidden") return { visibility: "hidden" };
  return { filter: MODE_FILTERS[mode] };
}

/** `style` with the block's mode filter added (kept as is when it has none). */
export function withPresentationBlackWhiteStyle<S extends object | undefined>(
  viewMode: PresentationViewColorMode,
  block: { blackWhiteMode?: unknown },
  style: S,
): S {
  const extra = presentationBlackWhiteStyle(viewMode, block);
  return extra ? ({ ...(style ?? {}), ...extra } as S) : style;
}

/**
 * The blocks with the selected ones set to `mode` (Automatic removes the
 * setting); null when nothing changes.
 */
export function blocksWithBlackWhiteMode<B extends { id: string; blackWhiteMode?: unknown }>(
  blocks: readonly B[],
  selected: ReadonlySet<string>,
  mode: PresentationBlackWhiteMode,
): B[] | null {
  let changed = false;
  const next = blocks.map((block) => {
    if (!selected.has(block.id)) return block;
    const current = normalizePresentationBlackWhiteMode(block.blackWhiteMode) ?? "auto";
    if (current === mode) return block;
    changed = true;
    const copy = { ...block };
    if (mode === "auto") delete copy.blackWhiteMode;
    else copy.blackWhiteMode = mode;
    return copy;
  });
  return changed ? next : null;
}
