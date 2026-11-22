/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Table Design > Effects (Shadow, Reflection) and Shading > Table Background:
 * the page model, PowerPoint's galleries and the canvas rendering.
 *
 * PowerPoint's Effects menu on a table has three submenus (Microsoft Support,
 * "Change the look of a table"): Cell Bevel bevels each selected cell (the
 * cell's `bevel`), while Shadow and Reflection act on the WHOLE table whatever
 * cells are selected; Table Background paints the whole table under every
 * cell fill. The file side -- reader, from-scratch writer, in-place patch --
 * is the server's `app/services/office_table_effects.py` (`a:tblPr` fill and
 * `a:effectLst`). A table block keeps `tableEffects` (`shadow`, `reflection`,
 * each the effect's DrawingML parameters: EMU, 1/60000 degree, 1/1000 %) and
 * `tableBackground` (`#rrggbb`).
 *
 * Reflection's nine variations are the picture model's, measured in
 * PowerPoint for Mac (`lib/office-picture-effects`); Shadow is the shared
 * shadow model and its Outer / Inner / Perspective galleries
 * (`lib/office-shadow-effects`), each value classified by PowerPoint for Mac
 * as its gallery type.
 */

import type { CSSProperties } from "react";

import {
  normalizeOfficePictureEffects,
  type OfficeReflectionEffect,
} from "./office-picture-effects";
import {
  normalizeOfficeShadowEffect,
  officeShadowCss,
  officeShadowSilhouette,
  POWERPOINT_SHADOW_BLUR_DEVIATION,
  type OfficeShadowEffect,
  type OfficeShadowSilhouette,
} from "./office-shadow-effects";

/** The table's shadow: the shared model (`lib/office-shadow-effects`). */
export type TableShadowEffect = OfficeShadowEffect;

export interface PresentationTableEffects {
  readonly shadow?: TableShadowEffect;
  readonly reflection?: OfficeReflectionEffect;
}
export type PresentationTableEffectKind = keyof PresentationTableEffects;

const EMU_PER_PX = 9_525;

/** The page's `tableEffects`, or null when absent or not one (the server's rule). */
export function normalizePresentationTableEffects(value: unknown): PresentationTableEffects | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  const keys = Object.keys(raw);
  if (!keys.length || keys.some((key) => key !== "shadow" && key !== "reflection")) return null;
  const effects: { -readonly [K in PresentationTableEffectKind]?: PresentationTableEffects[K] } = {};
  if ("shadow" in raw) {
    const shadow = normalizeOfficeShadowEffect(raw.shadow);
    if (!shadow) return null;
    effects.shadow = shadow;
  }
  if ("reflection" in raw) {
    const reflection = normalizeOfficePictureEffects({ reflection: raw.reflection })?.reflection;
    if (!reflection || reflection === "none") return null;
    effects.reflection = reflection;
  }
  return effects;
}

/** The effects after one gallery pick (`value` null = "No <effect>"); undefined when none are left. */
export function withPresentationTableEffect<K extends PresentationTableEffectKind>(
  current: unknown,
  kind: K,
  value: NonNullable<PresentationTableEffects[K]> | null,
): PresentationTableEffects | undefined {
  const next: { -readonly [P in PresentationTableEffectKind]?: PresentationTableEffects[P] } = {
    ...(normalizePresentationTableEffects(current) ?? {}),
  };
  if (value === null) delete next[kind];
  else next[kind] = value;
  return Object.keys(next).length ? next : undefined;
}

/** Whether two shadow / reflection values are the same pick (a menu's checked state). */
export function samePresentationTableEffect(left: unknown, right: unknown): boolean {
  return JSON.stringify(left ?? null) === JSON.stringify(right ?? null);
}

/** The table's `tableBackground`, or null. */
export function normalizePresentationTableBackground(value: unknown): string | null {
  return typeof value === "string" && /^#[0-9a-fA-F]{6}$/.test(value) ? value.toLowerCase() : null;
}

function px(emu: number): number {
  return Math.round((emu / EMU_PER_PX) * 100) / 100;
}

/**
 * The table's Reflection on its frame, a faded mirror image below, as the
 * picture model draws it. Its Shadow is not the frame's
 * (`presentationTableShadowSilhouette`, `presentationTableShadowLayers`).
 */
export function presentationTableEffectsCss(value: unknown): CSSProperties | undefined {
  const effects = normalizePresentationTableEffects(value);
  if (!effects) return undefined;
  const css: CSSProperties = {};
  // The table's own shadow is not the frame's: an outer one is the silhouette
  // of what its cells paint (`presentationTableShadowSilhouette`), an inner
  // one its own layer (`presentationTableShadowLayers`).
  const reflection = effects.reflection;
  if (reflection) {
    const start = 100 - reflection.endPos / 1000;
    css.WebkitBoxReflect =
      `below ${px(reflection.dist)}px linear-gradient(to bottom, transparent ${start}%, ` +
      `rgb(0 0 0 / ${reflection.endA / 100_000}) ${start}%, rgb(0 0 0 / ${reflection.stA / 100_000}) 100%)`;
  }
  return Object.keys(css).length ? css : undefined;
}

/**
 * The table frame's own paint: Table Background under the cells, and the
 * Shadow and Reflection. `width` / `height` are the block's EMU frame.
 */
export function presentationTableFrameCss(block: {
  readonly tableBackground?: unknown;
  readonly tableEffects?: unknown;
  readonly width?: unknown;
  readonly height?: unknown;
}): CSSProperties {
  const background = normalizePresentationTableBackground(block.tableBackground);
  return {
    ...(background ? { background } : {}),
    ...(presentationTableEffectsCss(block.tableEffects) ?? {}),
  };
}

export function frameSize(block: { readonly width?: unknown; readonly height?: unknown }) {
  const extent = (value: unknown) => (typeof value === "number" && Number.isFinite(value) ? value / EMU_PER_PX : 0);
  return { width: extent(block.width), height: extent(block.height) };
}

/**
 * The layer an inner shadow needs over the table's cells (their fills would
 * hide an inset shadow on the table itself). Both null for an outer shadow:
 * that is the silhouette of the cells' paint (`presentationTableShadowSilhouette`).
 */
export function presentationTableShadowLayers(
  block: {
    readonly tableEffects?: unknown;
    readonly width?: unknown;
    readonly height?: unknown;
  },
  /**
   * The table style's background shadow (`tableStyleLookShadowEffect`), drawn
   * when the table states no shadow of its own: a scaled or skewed one needs
   * the layer under the table.
   */
  styleShadow?: unknown,
): { readonly under: CSSProperties | null; readonly over: CSSProperties | null } {
  const shadow = normalizePresentationTableEffects(block.tableEffects)?.shadow ?? normalizeOfficeShadowEffect(styleShadow);
  // An outer shadow is the silhouette of the cells' paint
  // (`presentationTableShadowSilhouette`), not a layer the frame's size.
  if (!shadow?.inner) return { under: null, over: null };
  const css = officeShadowCss(shadow, frameSize(block), { blurDeviation: POWERPOINT_SHADOW_BLUR_DEVIATION });
  return { under: css?.under ?? null, over: css?.over ?? null };
}
















export function presentationTableShadowSilhouette(
  block: { readonly tableEffects?: unknown },
  /** The table style's background shadow (`tableStyleLookShadowEffect`), when the table states none of its own. */
  styleShadow?: unknown,
): OfficeShadowSilhouette | null {
  const shadow = normalizePresentationTableEffects(block.tableEffects)?.shadow ?? normalizeOfficeShadowEffect(styleShadow);
  if (!shadow || shadow.inner) return null;
  return officeShadowSilhouette(shadow, { blurDeviation: POWERPOINT_SHADOW_BLUR_DEVIATION });
}

