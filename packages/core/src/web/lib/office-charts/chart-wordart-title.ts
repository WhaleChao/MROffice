/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A WordArt chart title, drawn as a picture over the title's own (then
 * transparent) words: one entry in the chart option's `graphic`, built once
 * for every planner (Docs' shared planner, the deck's adapter) and placed by
 * the chart engine on the words where the layout drew them.
 */
import type { OfficeTextPaint, officeTextPaintSvg } from "../office-text-paint";

/** The `graphic` id of a WordArt title's picture. */
export const CHART_WORDART_TITLE_ID = "docs-wordart-title";

/** ECharts' own title padding: a title's words begin this far inside its `top`. */
const TITLE_PADDING_PX = 5;

type Painted = ReturnType<typeof officeTextPaintSvg>;

/**
 * The picture's `graphic` entry: centred over a title at `titleTop` drawn in
 * `fontSize` (the place a top title takes), with the words' centre in the
 * picture recorded (`info`) so the engine can stand the picture on the words
 * wherever the layout actually put them (a native title's own place, a band
 * a plan shares, a title moved off the top).
 */
export function chartWordArtTitleGraphic(
  painted: Painted,
  titleTop: number,
  fontSize: number,
  /** The paint itself: an editor that draws the title draws it with this. */
  paint?: OfficeTextPaint,
): Record<string, unknown> {
  return {
    id: CHART_WORDART_TITLE_ID,
    type: "image",
    silent: true,
    z: 100,
    left: "center",
    top: titleTop + TITLE_PADDING_PX + fontSize / 2 - painted.glyphCenter.y,
    info: { wordArtGlyph: painted.glyphCenter, ...(paint ? { paint } : {}) },
    style: { image: painted.uri, width: painted.width, height: painted.height },
  };
}

/**
 * The option's graphics with the WordArt title's picture in: any picture a
 * planner drew before is replaced, every other graphic (a map's scale, a
 * treemap's labels) is kept.
 */
export function withChartWordArtTitle(
  graphic: unknown,
  entry: Record<string, unknown>,
): Record<string, unknown>[] {
  const others = (Array.isArray(graphic) ? graphic : []).filter(
    (item): item is Record<string, unknown> =>
      Boolean(item) && typeof item === "object" && (item as { id?: unknown }).id !== CHART_WORDART_TITLE_ID,
  );
  return [...others, entry];
}

/** The words' centre in a WordArt title's picture, from its `graphic` entry. */
export function chartWordArtTitleGlyph(
  graphic: unknown,
): { entry: Record<string, unknown>; x: number; y: number } | null {
  const entry = (Array.isArray(graphic) ? graphic : []).find(
    (item): item is Record<string, unknown> =>
      Boolean(item) && typeof item === "object" && (item as { id?: unknown }).id === CHART_WORDART_TITLE_ID,
  );
  const info = entry?.info as { wordArtGlyph?: { x?: unknown; y?: unknown } } | undefined;
  const x = info?.wordArtGlyph?.x, y = info?.wordArtGlyph?.y;
  return entry && typeof x === "number" && typeof y === "number" && Number.isFinite(x) && Number.isFinite(y)
    ? { entry, x, y }
    : null;
}
