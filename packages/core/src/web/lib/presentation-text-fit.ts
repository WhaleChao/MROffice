/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PowerPoint's Resize shape to fit text (Format Shape > Text Options > Text
 * Box; `a:spAutoFit`): the frame of a text box follows its words.
 *
 * Microsoft's own statements (Support, "Fit text in a shape or text box":
 * Format Shape > Size & Properties > Text Box > Resize shape to fit text;
 * VBA `TextFrame.AutoSize` = `ppAutoSizeShapeToFitText`; ECMA-376 `spAutoFit`:
 * "the shape should be auto-fit to fully contain the text ... when the user
 * types more text within the shape, the shape actually resizes"):
 * - Insert > Text Box and Insert > WordArt boxes resize to their text; a
 *   drawn shape does not (its default is Do not Autofit).
 * - The box takes the height its text needs plus its top and bottom margins;
 *   the width stays the author's while the text wraps ("Wrap text in shape").
 * - PowerPoint resizes while the user types; opening a file keeps the file's
 *   extent (§35.1, measured on real PowerPoint), and the resize belongs to
 *   the edit that caused it (one Undo step).
 * How the frame moves is this editor's reading of the same rule -- the text
 * keeps its place while its box resizes:
 * - the edge the text is anchored to stays (top / bottom; a middle-anchored
 *   box resizes about its centre), and the box fits exactly, so it shrinks
 *   back when words are deleted;
 * - a box that does not wrap (`wrap="none"`) also takes its longest line's
 *   width plus its left and right margins, keeping its aligned side;
 * - a rotated box resizes in its own frame: the anchored edge stays where it
 *   is on the rotated box, so the centre moves along the box's own axes.
 */

export type TextFitAnchor = "top" | "middle" | "bottom";
export type TextFitAlign = "left" | "center" | "right";

export interface TextFitFrame {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
  /** Degrees, clockwise (the canvas's `rotate()`). */
  readonly rotation?: number;
}

/** A change smaller than this (EMU, about a design px) is no change. */
export const TEXT_FIT_TOLERANCE_EMU = 9_525;

/**
 * The frame a box takes to fit `needed` (EMU: the text's height and, for a
 * box that does not wrap, its width -- margins included), keeping its
 * anchored edge (and, unwrapped, its aligned side) where it is. Null when the
 * box already fits.
 */
export function presentationTextFitFrame(
  frame: TextFitFrame,
  needed: { readonly height: number; readonly width?: number },
  options: { readonly anchor: TextFitAnchor; readonly align: TextFitAlign; readonly wrap: boolean },
): { left: number; top: number; width: number; height: number } | null {
  const height = Math.max(1, Math.round(needed.height));
  const width =
    !options.wrap && typeof needed.width === "number" && needed.width > 0
      ? Math.max(1, Math.round(needed.width))
      : frame.width;
  const dh = height - frame.height;
  const dw = width - frame.width;
  if (Math.abs(dh) <= TEXT_FIT_TOLERANCE_EMU && Math.abs(dw) <= TEXT_FIT_TOLERANCE_EMU) return null;
  // The centre's move in the box's own frame.
  const localY = options.anchor === "top" ? dh / 2 : options.anchor === "bottom" ? -dh / 2 : 0;
  const localX = options.align === "left" ? dw / 2 : options.align === "right" ? -dw / 2 : 0;
  const angle = ((frame.rotation ?? 0) * Math.PI) / 180;
  const cos = Math.cos(angle);
  const sin = Math.sin(angle);
  const centerX = frame.left + frame.width / 2 + localX * cos - localY * sin;
  const centerY = frame.top + frame.height / 2 + localX * sin + localY * cos;
  return {
    left: Math.round(centerX - width / 2),
    top: Math.round(centerY - height / 2),
    width,
    height,
  };
}

/** A paragraph alignment as the side a no-wrap box keeps. */
export function textFitAlign(align: unknown): TextFitAlign {
  return align === "center" ? "center" : align === "right" ? "right" : "left";
}
