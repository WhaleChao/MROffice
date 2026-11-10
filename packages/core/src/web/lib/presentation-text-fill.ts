/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Per-text-box FILL colour + OUTLINE / border (PowerPoint / WPS "Shape Fill" /
 * "Shape Outline" applied to a text box).
 *
 * A text block (especially a free text box) may carry a background FILL (solid
 * or a two-stop gradient) and an OUTLINE (colour + weight + dash style), exactly
 * like a shape. Rather than duplicate the colour / stroke validation, this
 * module REUSES the shape normalizers (`lib/presentation-shapes`) — the text-box
 * fill IS the shape fill model, just attached to a text block.
 *
 * Pure, side-effect-free helpers shared by the presentation schema, the editor
 * canvas, and the present-mode overlay so the resolution + render-ready CSS live
 * in one place. Keep this module dependency-free (no React / no DOM) so it can be
 * unit tested in isolation and imported from the zod schema.
 *
 * Everything is ADDITIVE and OPTIONAL: a text block with NO fill / outline
 * resolves to `hasBox: false`, so the renderer keeps the historical plain text
 * render byte-for-byte and an un-styled deck is unchanged.
 */

import type { CSSProperties } from "react";

import {
  DEFAULT_GRADIENT_ANGLE,
  MAX_STROKE_WIDTH,
  SHAPE_SHADOW_CSS,
  normalizeGradientAngle,
  normalizeShapeColor,
  normalizeStrokeStyle,
  normalizeStrokeWidth,
  type StrokeStyle,
} from "./presentation-shapes";
import {
  officeStrokeAlign,
  type OfficeStrokeAlign,
} from "./office-shape-line-align";

/** The raw, optional fill / outline fields a text block may persist. */
export interface TextBoxFillInput {
  /** Background fill colour (hex). Absent ⇒ transparent (no fill). */
  fill?: unknown;
  /** The solid fill's Transparency, whole percent 1..100 (`a:alpha`), as a
   * shape's `fillTransparency`. Ignored with a gradient. */
  fillTransparency?: unknown;
  /** Second gradient stop; when present the fill renders as fill→fillTo. */
  fillTo?: unknown;
  /** Gradient sweep angle (deg, CSS convention). */
  gradientAngle?: unknown;
  /** Outline colour (hex). Only meaningful when `outlineWidth > 0`. */
  outlineColor?: unknown;
  /** Second outline stop; together with outlineColor makes a gradient line. */
  outlineColorTo?: unknown;
  outlineGradientAngle?: unknown;
  /** Outline weight (px). 0 ⇒ no outline. */
  outlineWidth?: unknown;
  /** Outline dash style. */
  outlineStyle?: unknown;
  /** `a:ln@cmpd`: a double stroke (PPT-B368). */
  outlineCompound?: unknown;
  /**
   * `a:ln@algn`. A text box IS a shape (SHAPE-IMPORT-AUDIT.md section 37), so
   * its border carries the same pen alignment a shape's outline carries.
   * Absent is a frame whose source line states nothing.
   */
  outlineAlign?: unknown;
  /** Soft drop shadow. */
  shadow?: unknown;
}

/** A resolved, render-ready text-box fill / outline descriptor. */
export interface ResolvedTextBoxFill {
  /** Background fill colour, or `null` for a transparent (fill-less) box. */
  fill: string | null;
  /** The solid fill's opacity, 0..1 (1 is opaque, and every gradient). */
  fillOpacity: number;
  /** The persisted Transparency behind `fillOpacity`, or `null` for opaque. */
  fillTransparency: number | null;
  /** Second gradient stop; non-null ⇒ the fill is a gradient from fill→fillTo. */
  fillTo: string | null;
  /** Gradient sweep angle in degrees (CSS convention). */
  gradientAngle: number;
  /** Outline weight in px (0 ⇒ no visible outline). */
  outlineWidth: number;
  /** Outline colour (only meaningful when `outlineWidth > 0`). */
  outlineColor: string;
  outlineColorTo: string | null;
  outlineGradientAngle: number;
  /** Outline dash style. */
  outlineStyle: StrokeStyle;
  /** The outline is drawn as two parallel strokes with a gap. */
  outlineCompound: boolean;
  /**
   * The frame's pen alignment, or `null` when its line states none. Carried so
   * a package that says `algn="in"` says it again on the way out; see
   * `@/lib/office-shape-line-align`.
   */
  outlineAlign: OfficeStrokeAlign | null;
  /** Whether a soft drop shadow is drawn under the box. */
  shadow: boolean;
  /** True iff at least one of fill / outline / shadow is active. */
  hasBox: boolean;
  




  square?: true;
}

/** Default outline colour when an outline is enabled without an explicit one. */
export const DEFAULT_TEXTBOX_OUTLINE_COLOR = "#1f2937";

/**
 * The weight an outline gets when the user turns one ON without naming a weight
 * -- picking a Shape Outline colour, or a dash, on a frame that had no line.
 *
 * A text frame's outline IS its width in this model: `persistedTextBoxFill`
 * drops the colour and the dash style whenever the width is 0, so a colour or a
 * dash applied alone would be silently discarded. PowerPoint turns the line on
 * in the same gesture, so this is the weight that makes the command visible.
 */
export const DEFAULT_TEXTBOX_OUTLINE_WIDTH = 2;

/**
 * Resolve a text block's persisted fill / outline fields into a normalized
 * descriptor. A gradient is active only when BOTH a base fill AND a distinct
 * second stop resolve. The outline colour is retained only when a positive width
 * makes it visible. `hasBox` is false exactly when no fill, no outline, and no
 * shadow apply — letting the renderer keep the plain text render unchanged.
 */
export function resolveTextBoxFill(
  input: TextBoxFillInput | undefined,
): ResolvedTextBoxFill {
  const fill = normalizeShapeColor(input?.fill, null);
  const fillTo = fill === null ? null : normalizeShapeColor(input?.fillTo, null);
  const outlineWidth = normalizeStrokeWidth(
    typeof input?.outlineWidth === "number" ? input.outlineWidth : 0,
  );
  const outlineColor =
    outlineWidth > 0
      ? normalizeShapeColor(input?.outlineColor, DEFAULT_TEXTBOX_OUTLINE_COLOR)
      : DEFAULT_TEXTBOX_OUTLINE_COLOR;
  const outlineColorTo =
    outlineWidth > 0
      ? normalizeShapeColor(input?.outlineColorTo, null)
      : null;
  const shadow = input?.shadow === true;
  const transparency = input?.fillTransparency;
  const fillTransparency =
    fill !== null &&
    fillTo === null &&
    typeof transparency === "number" &&
    Number.isInteger(transparency) &&
    transparency >= 1 &&
    transparency <= 100
      ? transparency
      : null;
  return {
    fill,
    fillOpacity: fillTransparency === null ? 1 : (100 - fillTransparency) / 100,
    fillTransparency,
    fillTo,
    gradientAngle:
      input?.gradientAngle === undefined
        ? DEFAULT_GRADIENT_ANGLE
        : normalizeGradientAngle(input.gradientAngle),
    outlineWidth,
    outlineColor,
    outlineColorTo,
    outlineGradientAngle:
      input?.outlineGradientAngle === undefined
        ? DEFAULT_GRADIENT_ANGLE
        : normalizeGradientAngle(input.outlineGradientAngle),
    outlineStyle: normalizeStrokeStyle(input?.outlineStyle),
    outlineCompound: input?.outlineCompound === "double",
    outlineAlign:
      outlineWidth > 0 ? officeStrokeAlign(input?.outlineAlign) : null,
    shadow,
    hasBox: fill !== null || outlineWidth > 0 || shadow,
    ...(textBoxFromFileShape(input) ? { square: true as const } : {}),
  };
}

/** Whether a text frame is a shape a file stated (its slide's shape id). */
export function textBoxFromFileShape(input: unknown): boolean {
  const source = (input as { source?: unknown } | undefined)?.source;
  return (
    typeof source === "object" &&
    source !== null &&
    typeof (source as { shapeId?: unknown }).shapeId === "string"
  );
}

/**
 * The `background` value for the text box: a solid colour, or a linear-gradient
 * `fill → fillTo` at the resolved angle when a gradient is active. Returns
 * `"transparent"` for a fill-less box.
 */
export function textBoxFillBackground(fill: ResolvedTextBoxFill): string {
  if (fill.fill === null) return "transparent";
  if (fill.fillTo === null) {
    if (fill.fillOpacity >= 1) return fill.fill;
    // A see-through solid fill (Fill > Transparency): the colour at its alpha.
    const alpha = Math.round(Math.max(0, fill.fillOpacity) * 255).toString(16).padStart(2, "0");
    return /^#[0-9a-fA-F]{6}$/.test(fill.fill) ? `${fill.fill}${alpha}` : fill.fill;
  }
  return `linear-gradient(${fill.gradientAngle}deg, ${fill.fill}, ${fill.fillTo})`;
}

/**
 * The CSS to layer onto the text-box WRAPPER for the resolved fill / outline.
 * Returns an empty object when nothing is active (so spreading it is a no-op and
 * the historical render stays byte-stable). A box with any fill/outline/shadow
 * also gets a little padding + rounded corners so the text doesn't touch the
 * edge (PowerPoint text-box inset), and `boxSizing: border-box` so the outline
 * doesn't grow the block geometry.
 */
/** The path a text-bearing shape's outline draws (`source.nativeTextGeometry`,
 * an imported freeform or oval with words), in its design-pixel view box, or
 * null for the plain box every other text frame is. */
export function textBoxNativeGeometry(
  block: unknown,
): { readonly paths: readonly { d: string; fill?: string; stroke?: boolean }[]; readonly viewBox: readonly [number, number] } | null {
  const source = (block as { source?: unknown } | undefined)?.source;
  const record =
    source && typeof source === "object" && !Array.isArray(source)
      ? (source as Record<string, unknown>).nativeTextGeometry
      : undefined;
  if (!record || typeof record !== "object" || (record as { schemaVersion?: unknown }).schemaVersion !== "pptx-text-geometry-v1") {
    return null;
  }
  const geometry = (record as { geometry?: unknown }).geometry as
    | { paths?: unknown; viewBox?: unknown }
    | undefined;
  const paths = Array.isArray(geometry?.paths)
    ? (geometry!.paths as unknown[]).filter(
        (path): path is { d: string; fill?: string; stroke?: boolean } =>
          !!path && typeof (path as { d?: unknown }).d === "string" && (path as { d: string }).d.length <= 200_000,
      )
    : [];
  const viewBox = geometry?.viewBox;
  if (
    !paths.length ||
    !Array.isArray(viewBox) ||
    viewBox.length !== 2 ||
    !viewBox.every((value) => typeof value === "number" && Number.isFinite(value) && value > 0)
  ) {
    return null;
  }
  return { paths, viewBox: [viewBox[0] as number, viewBox[1] as number] };
}

export function textBoxFillCss(
  fill: ResolvedTextBoxFill,
  /** The frame's fill and line are painted on its own path
   * (`textBoxNativeGeometry`), not as the box's background and border. */
  options?: { onGeometry?: boolean },
): CSSProperties {
  if (!fill.hasBox) return {};
  if (options?.onGeometry) {
    return { boxSizing: "border-box", padding: "0.4em 0.6em", isolation: "isolate" };
  }
  const css: CSSProperties = {
    background: textBoxFillBackground(fill),
    boxSizing: "border-box",
  };
  if (fill.outlineWidth > 0) {
    // WHERE the ink sits. Measured in real PowerPoint (SHAPE-IMPORT-AUDIT.md
    // §68.5): a 216 × 144 pt text frame with a 6 pt border and no stated
    // `algn` renders, in the PDF PowerPoint itself writes, with its STROKED
    // path at exactly 216 × 144 — the frame path, ink straddling it, the same
    // reading §59 took for a shape. An `algn="in"` frame is the exception the
    // file asks for and keeps the whole width inside.
    //
    // CSS can only paint a border inside a box, so a centred pen is painted
    // as two halves: the inner half as the border and the outer half as an
    // `outline`, which takes no layout space, so the block geometry is
    // untouched either way. A GRADIENT outline keeps the whole width inside —
    // `border-image` has no outer half to give — and is named in §68.7.
    const gradient = fill.outlineColorTo !== null;
    // A compound (double) stroke is one CSS `double` border over the whole
    // width: two lines and the gap between them (PPT-B368); it cannot be
    // split into an inner and an outer half.
    const straddle = fill.outlineAlign !== "inset" && !gradient && !fill.outlineCompound;
    const half = fill.outlineWidth / 2;
    if (straddle && fill.outlineStyle === "solid") {
      // A solid pen's inner half is an inset shadow, not a border: a border
      // narrows the box its words wrap in, and PowerPoint lays a frame's
      // words out from its own edge whatever its line (CDC p42: a 1.5 pt red
      // frame broke "National" into "Nationa / l", which fits by 0.2 px).
      css.boxShadow = `inset 0 0 0 ${half}px ${fill.outlineColor}`;
      css.outline = `${half}px solid ${fill.outlineColor}`;
      css.borderRadius = fill.square ? 0 : 6;
      css.padding = "0.4em 0.6em";
      if (fill.shadow) css.boxShadow = `${css.boxShadow}, ${SHAPE_SHADOW_CSS}`;
      return css;
    }
    css.borderWidth = straddle ? half : fill.outlineWidth;
    css.borderStyle = fill.outlineCompound ? "double" : fill.outlineStyle;
    css.borderColor = fill.outlineColor;
    if (gradient) {
      css.borderImageSource = `linear-gradient(${fill.outlineGradientAngle}deg, ${fill.outlineColor}, ${fill.outlineColorTo})`;
      css.borderImageSlice = 1;
    } else if (straddle) {
      css.outline = `${half}px ${fill.outlineStyle} ${fill.outlineColor}`;
    }
  }
  // Rounded corners + inset only when there's something visible to inset from.
  css.borderRadius = fill.square ? 0 : 6;
  css.padding = "0.4em 0.6em";
  if (fill.shadow) css.boxShadow = SHAPE_SHADOW_CSS;
  return css;
}

/** The persisted text-box fill / outline field names handled by the normalizers. */
export const TEXTBOX_FILL_KEYS = [
  "fill",
  "fillTransparency",
  "fillTo",
  "gradientAngle",
  "outlineColor",
  "outlineColorTo",
  "outlineGradientAngle",
  "outlineWidth",
  "outlineStyle",
  "outlineAlign",
  "shadow",
] as const;

/**
 * The fill / outline fields to PERSIST for a text block, omitting every field at
 * its default so an un-styled box serializes byte-for-byte as before. A gradient
 * second stop is kept only with a base fill; an outline colour / style only with
 * a positive width; the gradient angle only with an actual gradient.
 */
export function persistedTextBoxFill(
  input: TextBoxFillInput | undefined,
): Partial<{
  fill: string;
  fillTransparency: number;
  fillTo: string;
  gradientAngle: number;
  outlineColor: string;
  outlineColorTo: string;
  outlineGradientAngle: number;
  outlineWidth: number;
  outlineStyle: StrokeStyle;
  outlineAlign: OfficeStrokeAlign;
  shadow: true;
}> {
  const resolved = resolveTextBoxFill(input);
  const out: ReturnType<typeof persistedTextBoxFill> = {};
  if (resolved.fill !== null) {
    out.fill = resolved.fill;
    if (resolved.fillTransparency !== null) out.fillTransparency = resolved.fillTransparency;
    if (resolved.fillTo !== null) {
      out.fillTo = resolved.fillTo;
      // Only persist a non-default angle (keeps the common top→bottom byte-clean).
      if (resolved.gradientAngle !== DEFAULT_GRADIENT_ANGLE) {
        out.gradientAngle = resolved.gradientAngle;
      }
    }
  }
  if (resolved.outlineWidth > 0) {
    out.outlineWidth = Math.min(resolved.outlineWidth, MAX_STROKE_WIDTH);
    out.outlineColor = resolved.outlineColor;
    if (resolved.outlineColorTo !== null) {
      out.outlineColorTo = resolved.outlineColorTo;
      if (resolved.outlineGradientAngle !== DEFAULT_GRADIENT_ANGLE) {
        out.outlineGradientAngle = resolved.outlineGradientAngle;
      }
    }
    if (resolved.outlineStyle !== "solid") out.outlineStyle = resolved.outlineStyle;
    // Only a frame whose line STATED an alignment persists one: the `ln`
    // clause declares no default for `algn`, so writing one here would put an
    // attribute into every deck that never had it (audit 64/68).
    if (resolved.outlineAlign !== null) out.outlineAlign = resolved.outlineAlign;
  }
  if (resolved.shadow) out.shadow = true;
  return out;
}
