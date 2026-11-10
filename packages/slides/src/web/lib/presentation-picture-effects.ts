/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Picture Format > Picture Effects (and the same menu as Graphics Effects on
 * an SVG graphic and Video Effects on a video) on a slide object: the
 * command behind the ribbon's menu (`OfficePictureEffectsMenu`), kept out of
 * the window so a context menu or a keyboard route runs the same step.
 */

import {
  withOfficePictureEffect,
  type OfficePictureEffectKind,
  type OfficePictureEffectValues,
  type OfficePictureEffects,
} from "@mrnote/core/web/lib/office-picture-effects.ts";
import {
  isOfficePictureEffectPreset,
  officePictureEffectPresetLine,
} from "@mrnote/core/web/lib/office-picture-effect-presets.ts";
import { officePictureStyleEffectKinds } from "@mrnote/core/web/lib/office-picture-styles/native.ts";

/** A picture, an SVG graphic or a video: whatever carries `pictureEffects`. */
export interface PresentationPictureEffectsHost {
  readonly pictureEffects?: OfficePictureEffects;
  /** A picture's older shadow switch (the exporter's own soft shadow). */
  readonly shadow?: boolean;
  readonly pictureStylePreset?: string;
  readonly source?: unknown;
}

/**
 * Why an object's effects cannot be edited, or null when they can. An SVG
 * graphic's Graphics Effects and a video's Video Effects are the same
 * `p:spPr` effects (PV #36 / #37).
 */
export function presentationPictureEffectsBlocker(
  block: PresentationPictureEffectsHost,
): "preserved" | null {
  const source = block.source as Record<string, unknown> | undefined;
  if (source?.pictureEffectsDisposition === "preserve-only") return "preserved";
  return null;
}

/** The effects the picture's Picture Style draws. */
export function presentationPictureStyleEffects(
  block: PresentationPictureEffectsHost,
): ReadonlySet<string> {
  return officePictureStyleEffectKinds(block.pictureStylePreset);
}

/**
 * The object after one effects pick (`value` null = "No <effect>"), or the
 * same block when nothing changes. Over a Picture Style, "No <effect>"
 * removes the style's own effect too. A Shadow pick (or No Shadow) also
 * clears the picture's older shadow switch: the object has one shadow.
 */
export function applyPresentationPictureEffect<
  T extends PresentationPictureEffectsHost,
  K extends OfficePictureEffectKind,
>(block: T, kind: K, value: OfficePictureEffectValues[K] | null): T {
  if (presentationPictureEffectsBlocker(block)) return block;
  const effects = withOfficePictureEffect(
    block.pictureEffects,
    kind,
    value,
    presentationPictureStyleEffects(block).has(kind),
  );
  if (kind === "preset" && isOfficePictureEffectPreset(value)) return withPreset(block, value, effects);
  const clearsSwitch = kind === "shadow" && block.shadow !== undefined;
  if (!clearsSwitch && JSON.stringify(effects ?? null) === JSON.stringify(block.pictureEffects ?? null)) return block;
  const next = { ...block } as T & { pictureEffects?: OfficePictureEffects; shadow?: boolean };
  if (effects) next.pictureEffects = effects;
  else delete next.pictureEffects;
  if (kind === "shadow") delete next.shadow;
  return next;
}

const BORDER_KEYS = ["borderStyle", "borderColor", "borderWidth"] as const;

/**
 * Preset N, as PowerPoint applies it: the picture's scene, 3-D shape and
 * shadow become the preset's (`effects`: the glow, reflection and soft edges
 * picked before it stay; the older shadow switch goes), and its line becomes
 * the preset's -- the border fields, which the border writer turns into the
 * preset's own `a:ln`. A Picture Style's recipe would write its own line and
 * effects over it, so the style goes too. An SVG graphic has no border
 * fields: its line stays. No Presets is the generic path: the line stays.
 */
function withPreset<T extends PresentationPictureEffectsHost>(
  block: T,
  preset: number,
  effects: OfficePictureEffects | undefined,
): T {
  if (!isOfficePictureEffectPreset(preset)) return block;
  const next = { ...block, pictureEffects: effects ?? { preset } } as T & Record<string, unknown>;
  // Reflect: the host type states these two read-only.
  Reflect.deleteProperty(next, "shadow");
  Reflect.deleteProperty(next, "pictureStylePreset");
  if ((block as Record<string, unknown>).officeGraphicKind !== "svg") {
    for (const key of BORDER_KEYS) delete next[key];
    Object.assign(next, officePictureEffectPresetLine(preset));
  }
  const same = Object.keys({ ...block, ...next }).every(
    (key) => JSON.stringify((block as Record<string, unknown>)[key]) === JSON.stringify(next[key]),
  );
  return same ? block : next;
}
