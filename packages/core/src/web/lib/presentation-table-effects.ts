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
  type OfficeShadowEffect,
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
 * The table's Shadow and Reflection on the canvas, approximated with CSS on
 * the table's frame: an outer shadow a `box-shadow` at its offset and blur (a
 * scaled shadow grows by its spread), the reflection a faded mirror image
 * below, as the picture model draws it.
 */
export function presentationTableEffectsCss(
  value: unknown,
  size: { readonly width: number; readonly height: number },
): CSSProperties | undefined {
  const effects = normalizePresentationTableEffects(value);
  if (!effects) return undefined;
  const css: CSSProperties = {};
  const shadow = effects.shadow ? officeShadowCss(effects.shadow, size) : null;
  // A perspective or inner shadow draws on its own layer
  // (`presentationTableShadowLayers`).
  if (shadow?.boxShadow) css.boxShadow = shadow.boxShadow;
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
    ...(presentationTableEffectsCss(block.tableEffects, frameSize(block)) ?? {}),
  };
}

export function frameSize(block: { readonly width?: unknown; readonly height?: unknown }) {
  const extent = (value: unknown) => (typeof value === "number" && Number.isFinite(value) ? value / EMU_PER_PX : 0);
  return { width: extent(block.width), height: extent(block.height) };
}

/**
 * The layers a table's shadow needs beside its frame's own paint: under the
 * table for a perspective shadow, over the cells for an inner one (cell fills
 * would hide an inset shadow on the table itself). Both null for an outer
 * shadow, which is the frame's `box-shadow`.
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
  const shadow = normalizePresentationTableEffects(block.tableEffects)?.shadow ?? styleShadow;
  const css = shadow ? officeShadowCss(shadow, frameSize(block)) : null;
  return { under: css?.under ?? null, over: css?.over ?? null };
}

