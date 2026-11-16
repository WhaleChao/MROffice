/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Format Shape > Effects on a slide's text box, as the slide draws them. A
 * text box is a shape (`p:sp txBox="1"`) and wears the shape's one model
 * (`shapeEffects`, `lib/office-shape-effects-paint`), drawn as PowerPoint
 * draws it on a text box:
 *
 * - its shadow, glow, soft edges and 3-D Format are the frame's when the box
 *   paints one (a fill or a line) -- the words on it cast none of their own --
 *   and the words' when it paints none, as PowerPoint casts a no-fill
 *   shape's effects from its text;
 * - its reflection mirrors the frame and the words together;
 * - its 3-D Rotation turns the frame and the words about the box's centre,
 *   the words staying flat on its front when Keep text flat is checked
 *   (`keepTextFlat`) or while they are typed into (`officeShapeTextTurn`).
 *
 * A box with no `shapeEffects` (its older soft-shadow switch, `shadow`, is
 * the frame's CSS shadow) draws exactly as before: null.
 */

import type { CSSProperties } from "react";

import { normalizeOfficePictureEffects } from "@mrnote/core/web/lib/office-picture-effects.ts";
import {
  officeEffectsUnderGroupCamera,
  officeShapeEffectsPaint,
  officeShapeTextTurn,
  type OfficeShapeEffectsPaint,
} from "@mrnote/core/web/lib/office-shape-effects-paint.ts";

const EMU_PER_PX = 9525;

export interface PresentationTextBoxEffectsLook {
  readonly paint: OfficeShapeEffectsPaint;
  /** The frame (fill and line) is painted on its own layer under the words. */
  readonly frameLayer: boolean;
  /** CSS `filter` on the frame layer ("" for none). */
  readonly frameFilter: string;
  /** The frame layer's own 3-D turn: the turn the words do not take ("" for none). */
  readonly frameTurn: string;
  /** CSS `filter` on the whole box: a box painting no frame casts its effects from its words. */
  readonly hostFilter: string;
  /** The 3-D turn the whole box takes, its words with it ("" for none). */
  readonly hostTurn: string;
  /** `-webkit-box-reflect` of the whole box ("" for none). */
  readonly reflect: string;
}

export function presentationTextBoxEffectsLook(
  block: unknown,
  options: {
    /** `url(#…)` of the paint's SVG filter where the host defines it. */
    readonly filterUrl: string;
    /** The box paints a frame (a fill or a line, `resolveTextBoxFill(...).hasBox`). */
    readonly hasFrame: boolean;
    /** The words are being typed into: they lie flat, as PowerPoint edits them. */
    readonly editing?: boolean;
  },
): PresentationTextBoxEffectsLook | null {
  if (!block || typeof block !== "object") return null;
  const record = block as Record<string, unknown>;
  // A text box in a group with a camera turns by the group's alone.
  const effects = officeEffectsUnderGroupCamera(normalizeOfficePictureEffects(record.shapeEffects), record.scene3d);
  if (!effects) return null;
  const length = (value: unknown) =>
    (typeof value === "number" && Number.isFinite(value) && value > 0 ? value : 952500) / EMU_PER_PX;
  const size = { width: length(record.width), height: length(record.height) };
  // The block wears its turn and flips outside the box's turned layers
  // (`buildBlockTransform`): their 3-D turn is taken through the flips
  // (`flippedCameraTurn`), the frame's and the words' alike.
  const frame = {
    rotation: typeof record.rotation === "number" ? record.rotation : 0,
    flipH: record.flipH === true,
    flipV: record.flipV === true,
  };
  const paint = officeShapeEffectsPaint(effects, size, {
    frame,
    gradientFill: typeof record.fillTo === "string",
  });
  if (!paint) return null;
  const filter = [paint.primitives.length ? options.filterUrl : "", ...paint.cssFilters].filter(Boolean).join(" ");
  const wordsTurn =
    officeShapeTextTurn(effects, size, { keepTextFlat: record.keepTextFlat, editing: options.editing, frame })?.transform ?? "";
  const frameTurn = paint.transform && !wordsTurn ? paint.transform : "";
  const frameLayer = options.hasFrame && Boolean(filter || frameTurn || paint.silhouette);
  return {
    paint,
    frameLayer,
    frameFilter: frameLayer ? filter : "",
    frameTurn: frameLayer ? frameTurn : "",
    hostFilter: options.hasFrame ? "" : filter,
    hostTurn: wordsTurn,
    reflect: paint.reflect,
  };
}

/** The frame properties of `textBoxFillCss` (its paint), as against the box's layout. */
const FRAME_PAINT_KEYS = [
  "background",
  "backgroundColor",
  "backgroundImage",
  "boxShadow",
  "outline",
  "borderColor",
  "borderImageSource",
  "borderImageSlice",
  "borderRadius",
] as const;

/**
 * `textBoxFillCss` split for a frame painted on its own layer: the host keeps
 * what lays the words out (padding, a border's width drawn transparent), the
 * layer takes the paint, reaching out over the host's border.
 */
export function presentationTextBoxFrameSplit(frameCss: CSSProperties): {
  readonly host: CSSProperties;
  readonly layer: CSSProperties;
} {
  const host = { ...frameCss } as Record<string, unknown>;
  const layer: Record<string, unknown> = {};
  for (const key of FRAME_PAINT_KEYS) {
    if (key in host) {
      layer[key] = host[key];
      delete host[key];
    }
  }
  if (host.borderWidth !== undefined) {
    layer.borderWidth = host.borderWidth;
    layer.borderStyle = host.borderStyle;
    host.borderColor = "transparent";
  }
  const border = typeof host.borderWidth === "number" ? host.borderWidth : 0;
  return {
    host: host as CSSProperties,
    layer: {
      ...(layer as CSSProperties),
      position: "absolute",
      inset: -border,
      boxSizing: "border-box",
      pointerEvents: "none",
      zIndex: -1,
    },
  };
}
