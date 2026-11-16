/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { CSSProperties } from "react";

import { officeCameraEyePx } from "./office-camera-distance";
import { officeShadowCss, type OfficeShadowEffect } from "./office-shadow-effects";

export const OFFICE_PICTURE_EFFECT_PRESETS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12] as const;
export type OfficePictureEffectPreset = (typeof OFFICE_PICTURE_EFFECT_PRESETS)[number];

interface PresetLook {
  /** `a:outerShdw` of the preset's effect list, if any. */
  readonly shadow?: OfficeShadowEffect;
  /** `a:bevelT` width (EMU; DrawingML's default 76200 when the preset leaves `w` out). */
  readonly bevel: number;
  /** `a:sp3d@contourW` (EMU) in `a:contourClr`. */
  readonly contour?: { readonly width: number; readonly color: string };
  /** A perspective camera's `a:rot` (1/60000 degree) and `fov`. */
  readonly camera?: { readonly lat: number; readonly lon: number; readonly rev: number; readonly fov: number };
  /** Picture drawn over white at this opacity (< 1), or darkened to it (> 1 inverted: multiply). */
  readonly wash?: { readonly white?: number; readonly multiply?: number };
  /** The preset's `a:ln`, as the picture border fields. */
  readonly line: { readonly borderStyle: "none" } | { readonly borderStyle: "solid"; readonly borderColor: string; readonly borderWidth: number };
}

const black = (blurRad: number, dist: number, dir: number, alpha: number, scale?: number): OfficeShadowEffect => ({
  blurRad,
  dist,
  dir,
  algn: "ctr",
  color: "#000000",
  alpha,
  ...(scale ? { sx: scale, sy: scale } : {}),
});
const NO_LINE = { borderStyle: "none" } as const;

const LOOKS: Readonly<Record<OfficePictureEffectPreset, PresetLook>> = {
  1: { shadow: black(107_950, 12_700, 5_400_000, 100_000), bevel: 63_500, contour: { width: 44_450, color: "#ffffff" }, line: NO_LINE },
  2: { shadow: black(44_450, 27_940, 5_400_000, 32_000), bevel: 190_500, line: NO_LINE },
  3: { bevel: 139_700, line: NO_LINE },
  4: { shadow: black(190_500, 228_600, 2_700_000, 30_000), bevel: 127_000, line: NO_LINE },
  5: { shadow: black(149_987, 250_190, 8_460_000, 28_000), bevel: 88_900, line: NO_LINE },
  6: { shadow: black(57_785, 33_020, 3_180_000, 30_000), bevel: 38_100, line: NO_LINE },
  7: { bevel: 127_000, line: NO_LINE },
  8: { bevel: 76_200, wash: { white: 0.1 }, line: NO_LINE },
  9: {
    shadow: black(225_425, 50_800, 5_220_000, 33_000),
    bevel: 82_550,
    contour: { width: 19_050, color: "#ffffff" },
    camera: { lat: 486_000, lon: 19_530_000, rev: 174_000, fov: 3_300_000 },
    wash: { multiply: 0.87 },
    line: NO_LINE,
  },
  10: {
    shadow: black(317_500, 0, 2_700_000, 43_000),
    bevel: 260_350,
    camera: { lat: 19_086_000, lon: 19_067_999, rev: 3_108_000, fov: 2_700_000 },
    wash: { white: 0.9 },
    line: { borderStyle: "solid", borderColor: "#ffffff", borderWidth: 34_925 / 9_525 },
  },
  11: {
    shadow: black(127_000, 38_100, 2_700_000, 45_000),
    bevel: 203_200,
    camera: { lat: 20_376_000, lon: 1_938_000, rev: 20_112_001, fov: 2_700_000 },
    wash: { white: 0.84 },
    line: NO_LINE,
  },
  12: {
    shadow: black(184_150, 241_300, 11_520_000, 18_000, 110_000),
    bevel: 82_550,
    camera: { lat: 0, lon: 2_100_000, rev: 0, fov: 5_100_000 },
    wash: { multiply: 0.91 },
    line: NO_LINE,
  },
};

export function isOfficePictureEffectPreset(value: unknown): value is OfficePictureEffectPreset {
  return typeof value === "number" && (OFFICE_PICTURE_EFFECT_PRESETS as readonly number[]).includes(value);
}

/** The preset's line, as the picture border fields the border writer turns into its `a:ln`. */
export function officePictureEffectPresetLine(preset: OfficePictureEffectPreset): PresetLook["line"] {
  return LOOKS[preset].line;
}

/**
 * The preset's own effects as a shape draws them: its shadow, its top bevel
 * width (EMU), its contour and its perspective camera (`a:rot` in 1/60000
 * degree and the field of view). A picture adds its line and wash; a shape
 * has its own line, so it takes only these.
 */
export function officePictureEffectPresetEffects(preset: OfficePictureEffectPreset): {
  readonly shadow?: OfficeShadowEffect;
  readonly bevel: number;
  readonly contour?: { readonly width: number; readonly color: string };
  readonly camera?: { readonly lat: number; readonly lon: number; readonly rev: number; readonly fov: number };
} {
  const look = LOOKS[preset];
  return {
    ...(look.shadow ? { shadow: look.shadow } : {}),
    bevel: look.bevel,
    ...(look.contour ? { contour: look.contour } : {}),
    ...(look.camera ? { camera: look.camera } : {}),
  };
}

/** The preset's shadow (the shared shadow model), if it has one. */
export function officePictureEffectPresetShadow(preset: OfficePictureEffectPreset): OfficeShadowEffect | undefined {
  return LOOKS[preset].shadow;
}

const EMU_PER_PX = 9_525;

function degrees(units: number): number {
  const angle = units / 60_000;
  return angle > 180 ? angle - 360 : angle;
}

export interface OfficePictureEffectPresetCss {
  readonly wrapper: CSSProperties;
  readonly bevel?: CSSProperties;
  readonly shadowUnder?: CSSProperties;
  readonly shadowOver?: CSSProperties;
}

/**
 * The preset on the canvas (see the module comment for how close each part
 * is). `size` is the picture's frame in px; `scale` shrinks the lengths for a
 * gallery tile. `shadow: false` leaves the shadow to a Shadow picked over the
 * preset.
 */
export function officePictureEffectPresetCss(
  preset: OfficePictureEffectPreset,
  size: { readonly width: number; readonly height: number },
  options: { readonly scale?: number; readonly shadow?: boolean } = {},
): OfficePictureEffectPresetCss {
  const look = LOOKS[preset];
  const scale = options.scale ?? 1;
  const px = (emu: number) => (emu / EMU_PER_PX) * scale;
  const wrapper: CSSProperties = {};
  const filters: string[] = [];
  if (look.wash?.white) {
    // (1 - f) * picture + f * white, as contrast() then brightness().
    const f = 1 - look.wash.white;
    filters.push(`contrast(${((1 - f) / (1 + f)).toFixed(4)})`, `brightness(${(1 + f).toFixed(4)})`);
  } else if (look.wash?.multiply) {
    filters.push(`brightness(${look.wash.multiply})`);
  }
  if (filters.length) wrapper.filter = filters.join(" ");
  let shadow: ReturnType<typeof officeShadowCss> | null = null;
  if (look.shadow && options.shadow !== false) {
    const value = scale === 1 ? look.shadow : { ...look.shadow, blurRad: Math.round(look.shadow.blurRad * scale), dist: Math.round(look.shadow.dist * scale) };
    shadow = officeShadowCss(value, size);
  }
  const contour = look.contour ? `0 0 0 ${Math.max(0.5, px(look.contour.width)).toFixed(2)}px ${look.contour.color}` : "";
  const boxShadows = [contour, shadow?.boxShadow ?? ""].filter(Boolean);
  if (boxShadows.length) wrapper.boxShadow = boxShadows.join(", ");
  if (shadow?.under) wrapper.isolation = "isolate";
  if (look.camera) {
    // Office's camera stands at its own distance whatever the picture's size
    // (`officeCameraEyePx`, the shapes' and SmartArt's): probe
    // picture-preset-camera-probe p1-p3, Preset 12's far edge over its near
    // edge 0.859 at 2 in and 0.728 at 4 in (drawn 0.857 / 0.733; the 2 x
    // longer side / tan(fov / 2) this replaced drew 0.782 / 0.785 at both),
    // Preset 9's 0.726 / 0.606 (drawn 0.736 / 0.608; was 0.770 / 0.773).
    // A gallery tile (`scale`) stands it as much nearer as it shrinks.
    const distance = officeCameraEyePx(size, look.camera.fov / 60_000) * scale;
    wrapper.transform =
      `perspective(${distance.toFixed(1)}px) rotateZ(${(-degrees(look.camera.rev)).toFixed(2)}deg) ` +
      `rotateX(${(-degrees(look.camera.lat)).toFixed(2)}deg) rotateY(${(-degrees(look.camera.lon)).toFixed(2)}deg)`;
  }
  const rim = Math.max(1, Math.min(px(look.bevel), size.width / 4, size.height / 4));
  const bevel: CSSProperties = {
    position: "absolute",
    inset: 0,
    pointerEvents: "none",
    boxShadow: `inset ${rim / 2}px ${rim / 2}px ${rim}px rgb(255 255 255 / 0.55), inset ${-rim / 2}px ${-rim / 2}px ${rim}px rgb(0 0 0 / 0.35)`,
  };
  return {
    wrapper,
    bevel,
    ...(shadow?.under ? { shadowUnder: shadow.under } : {}),
    ...(shadow?.over ? { shadowOver: shadow.over } : {}),
  };
}
