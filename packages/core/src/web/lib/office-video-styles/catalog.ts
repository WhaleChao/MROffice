/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import data from "./native.json";

import type { NativePictureStyle } from "../office-picture-styles/native";
import type { OfficePictureEffects } from "../office-picture-effects";
import type { OfficeShadowAlignment } from "../office-shadow-effects";

/**
 * PowerPoint's 41 built-in Video Styles (Video Format > Video Styles), in
 * gallery order: PowerPoint for Mac 16.86's gallery, three groups (Subtle,
 * Moderate, Intense) of 10, 13 and 18 (its accessibility tree,
 * `VideoStylesGallery_Control_<group>_<n>`). Two items share a name (Beveled
 * Rounded Rectangle, items 14 and 25); the id follows the position.
 *
 * Like a picture style, a video style is a complete `spPr` recipe the server
 * writes (`app/services/file_adapters/office_video_styles.py`, the recipes
 * PowerPoint wrote); `native.json` is only what the editor draws, generated
 * from those recipes (`apps/api/scripts/generate_office_video_styles_web.py`),
 * in the picture styles' drawing shape so one renderer draws both galleries.
 */
export const OFFICE_VIDEO_STYLE_IDS = [
  "simple-beveled-rectangle",
  "center-shadow-rectangle",
  "soft-edge-rectangle",
  "beveled-frame-gradient",
  "simple-frame-black",
  "simple-frame-white",
  "slope-beveled-rectangle",
  "outer-shadow-rectangle",
  "glow-rectangle",
  "drop-shadow-rectangle",
  "compound-frame-black",
  "moderate-frame-black",
  "glow-rounded-rectangle",
  "beveled-rounded-rectangle",
  "rounded-diagonal-corner-white",
  "snip-diagonal-corner-white",
  "snip-diagonal-corner-gradient",
  "moderate-frame-gradient",
  "moderate-beveled-rectangle",
  "rotated-white",
  "rotated-gradient",
  "beveled-oval-black",
  "soft-edge-oval",
  "beveled-perspective-left",
  "beveled-rounded-rectangle-2",
  "reflected-rounded-rectangle",
  "reflected-bevel-white",
  "reflected-bevel-black",
  "reflected-perspective-right",
  "reflected-perspective-left",
  "bevel-reflected",
  "perspective-shadow-white",
  "perspective-shadow-gray",
  "bevel-perspective",
  "canvas-gray",
  "canvas-white",
  "relaxed-bevel-perspective",
  "monitor-gray",
  "bevel-rectangle",
  "metal-rounded-rectangle",
  "metal-frame",
] as const;
export type OfficeVideoStyleId = (typeof OFFICE_VIDEO_STYLE_IDS)[number];

export const OFFICE_VIDEO_STYLE_GROUPS: readonly {
  readonly id: "subtle" | "moderate" | "intense";
  readonly itemIds: readonly OfficeVideoStyleId[];
}[] = [
  { id: "subtle", itemIds: OFFICE_VIDEO_STYLE_IDS.slice(0, 10) },
  { id: "moderate", itemIds: OFFICE_VIDEO_STYLE_IDS.slice(10, 23) },
  { id: "intense", itemIds: OFFICE_VIDEO_STYLE_IDS.slice(23) },
];

/** One video style as the editor draws it (the picture styles' drawing shape). */
export type NativeVideoStyle = Omit<NativePictureStyle, "id" | "outlineIdentity" | "line"> & {
  readonly id: OfficeVideoStyleId;
  readonly index: number;
  readonly line?: NativePictureStyle["line"] & {
    /** `a:ln@cap` (rnd / sq / flat), as written. */
    readonly cap?: string;
    readonly gradient?: {
      readonly stops: readonly { readonly pos: number; readonly color: string; readonly opacity: number }[];
      readonly angle: number;
    };
  };
};

const STYLES = new Map(
  (data.styles as unknown as NativeVideoStyle[]).map((style) => [style.id, style]),
);

export function officeVideoStyle(id: unknown): NativeVideoStyle | null {
  return typeof id === "string" ? (STYLES.get(id as OfficeVideoStyleId) ?? null) : null;
}

/** The style's name as a `console-notebooks` key (PowerPoint's own names). */
export function officeVideoStyleLabelKey(id: OfficeVideoStyleId): string {
  return `office.videoStyle.${id}`;
}

/** The style in the picture-style renderer's shape (thumbnails, `OfficePictureStyleGraphic`). */
export function videoStyleAsPictureStyle(style: NativeVideoStyle): NativePictureStyle {
  return { ...style, outlineIdentity: null } as unknown as NativePictureStyle;
}

const num = (value: unknown, fallback = 0): number => {
  const parsed = typeof value === "number" ? value : typeof value === "string" ? Number(value) : NaN;
  return Number.isFinite(parsed) ? Math.round(parsed) : fallback;
};
const alpha = (opacity: unknown): number => Math.round(Math.max(0, Math.min(1, Number(opacity ?? 1))) * 100_000);
const ALIGNMENTS = new Set(["tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"]);

/**
 * The recipe's effects in the page's Picture Effects shape (`office-picture-
 * effects`), so the canvas draws a style's shadow, glow, soft edges,
 * reflection and bevel with the same frame a video's own effects use. The
 * camera is drawn by the style frame itself (a recipe may rotate an
 * orthographic camera, which Picture Effects' 3-D Rotation has no preset for).
 */
export function videoStyleEffects(style: NativeVideoStyle): OfficePictureEffects {
  const effects: Record<string, unknown> = {};
  for (const effect of style.effects ?? []) {
    const kind = String(effect.kind);
    if (kind === "outerShdw" || kind === "innerShdw") {
      if (effects.shadow) continue;
      const shadow: Record<string, unknown> = {
        ...(kind === "innerShdw" ? { inner: true } : {}),
        blurRad: num(effect.blurRad),
        dist: num(effect.dist),
        dir: num(effect.dir),
        color: String(effect.color ?? "#000000"),
        alpha: alpha(effect.opacity),
      };
      if (kind === "outerShdw") {
        shadow.algn = ALIGNMENTS.has(String(effect.algn)) ? String(effect.algn) as OfficeShadowAlignment : "b";
        for (const key of ["sx", "sy", "kx", "ky"] as const) {
          if (effect[key] !== undefined) shadow[key] = num(effect[key]);
        }
      }
      effects.shadow = shadow;
    } else if (kind === "glow") {
      effects.glow = { radius: num(effect.rad), color: String(effect.color ?? "#000000"), alpha: alpha(effect.opacity) };
    } else if (kind === "softEdge") {
      effects.softEdge = { radius: num(effect.rad) };
    } else if (kind === "reflection") {
      effects.reflection = {
        blurRad: num(effect.blurRad),
        stA: num(effect.stA, 100_000),
        endA: num(effect.endA, 0),
        endPos: num(effect.endPos, 100_000),
        dist: num(effect.dist),
      };
    }
  }
  const bevel = style.threeD?.bevel;
  if (bevel) {
    effects.bevel = { preset: String(bevel.prst ?? "circle"), width: num(bevel.w, 76_200), height: num(bevel.h, 76_200) };
  }
  return effects as OfficePictureEffects;
}
