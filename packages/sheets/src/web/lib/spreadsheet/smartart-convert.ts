/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Reset > Convert > Convert to Shapes on a worksheet: the members the graphic
 * drew, as the ordinary shapes Excel leaves (one undo step in the window, one
 * command for the AI; both call this).
 *
 * A sheet member already carries its item's words (the engine host merges a
 * label into its shape, `lib/spreadsheet/smart-art`), laid out in the
 * diagram's own text box for it (`smartArtTextBox`, Office's `dsp:txXfrm`) or
 * over the member's whole box. Releasing the member drops every `smartArt…`
 * key, that box with them, and an ordinary shape lays its words out in its
 * preset's text rectangle (`a:rect`) inside its body insets: the words went
 * back to the middle of the shape (a Segmented Cycle's five labels on top of
 * each other in the disc). PowerPoint's and Excel's result keeps them where
 * they were drawn, as the shape's own `a:bodyPr` insets measured from that
 * rectangle (SmartArt probe B8: Basic Block List, 247650 EMU each side), the
 * deck's `presentationSmartArtMembersAsShapes` rule. The words therefore keep
 * their place through the release, in the shape's `textLayout.bodyPr`, which
 * the sheet draws, prints and writes as `a:bodyPr`.
 *
 * The quick-style marker on the shape (`shape.smartArtStyle`) is the
 * diagram's too: it is what told the sheet to lay a member's words over its
 * whole box rather than its preset's rectangle. A converted shape is an
 * ordinary one, so it goes, and its words are measured against the
 * rectangle Excel will lay them out in.
 *
 * Words standing OUTSIDE that rectangle (a pyramid layer's words wider than
 * its trapezoid's `a:rect`) cannot be put there by a body inset, which is
 * never negative; they are held at the rectangle's edge on that side.
 */

import { releaseSmartArtMembers } from "@mrnote/core/web/lib/office-smartart/design-model.ts";
import {
  EMU_PER_CSS_PIXEL,
  normalizeOfficeTextLayout,
  officeTextBodyInsetsPx,
  type OfficeTextBodyProperties,
  type OfficeTextLayout,
} from "@mrnote/core/web/lib/office-text-layout.ts";
import { spreadsheetObjectRotation, type SpreadsheetObjectAnchor } from "./object-buttons";
import { spreadsheetSmartArtTextBoxOf } from "@mrnote/core/web/lib/spreadsheet/smart-art.ts";
import { spreadsheetShapeTextInset } from "./shape-text-inset";
import type { SpreadsheetShapeStyle } from "@mrnote/core/web/lib/spreadsheet/shapes.ts";

type Json = Record<string, unknown>;

/** The member's box on the sheet, in CSS px (`objectBoundsPixels`). */
export type SpreadsheetMemberSize<T> = (member: T) => { width: number; height: number } | null;

function shapeOf(object: Json): SpreadsheetShapeStyle | null {
  const shape = object.shape;
  return shape && typeof shape === "object" && !Array.isArray(shape) ? (shape as SpreadsheetShapeStyle) : null;
}

/**
 * Where the words of `object` are drawn, in px from the element's edges: the
 * label's text area (`spreadsheetShapeTextInset`) less the body's padding
 * (`spreadsheetShapeTextPaintCss`: the stated insets, else none inside a
 * SmartArt text box, else the DrawingML defaults).
 */
function wordsArea(
  shape: SpreadsheetShapeStyle,
  width: number,
  height: number,
  textBox: ReturnType<typeof spreadsheetSmartArtTextBoxOf>,
) {
  const area = spreadsheetShapeTextInset(shape, width, height, textBox);
  const body = normalizeOfficeTextLayout(shape.textLayout)?.bodyPr;
  const padding = body || !textBox ? officeTextBodyInsetsPx(body) : { left: 0, top: 0, right: 0, bottom: 0 };
  return {
    left: area.left * width + padding.left,
    top: area.top * height + padding.top,
    right: area.right * width + padding.right,
    bottom: area.bottom * height + padding.bottom,
  };
}

/**
 * The shape's words held where the graphic drew them once it is an ordinary
 * shape: its text area is then the preset's rectangle, and the difference is
 * its body's four insets.
 */
function wordsKeptInPlace(original: Json, released: Json, size: (member: Json) => { width: number; height: number } | null): Json {
  const shape = shapeOf(released);
  if (!shape) return released;
  const { smartArtStyle: _style, ...ordinary } = shape;
  const next: Json = { ...released, shape: ordinary };
  if (typeof shape.text !== "string" || shape.text.trim() === "") return next;
  const drawn = size(original);
  if (!drawn || !(drawn.width > 0) || !(drawn.height > 0)) return next;
  // Words kept level on a shape turned a quarter are laid out in the box
  // the turned shape covers, its sides exchanged (the view's `upright`).
  const turn = ((spreadsheetObjectRotation(original.anchor as SpreadsheetObjectAnchor | undefined) % 360) + 360) % 360;
  const level = shape.textUpright === true || shape.textLayout?.bodyPr?.upright === true;
  const box =
    level && ((turn >= 45 && turn < 135) || (turn >= 225 && turn < 315))
      ? { width: drawn.height, height: drawn.width }
      : drawn;
  const before = wordsArea(shapeOf(original) ?? shape, box.width, box.height, spreadsheetSmartArtTextBoxOf(original));
  const after = spreadsheetShapeTextInset(ordinary, box.width, box.height);
  const emu = (px: number) => Math.max(0, Math.round(px * EMU_PER_CSS_PIXEL));
  const insets = {
    lIns: emu(before.left - after.left * box.width),
    tIns: emu(before.top - after.top * box.height),
    rIns: emu(before.right - after.right * box.width),
    bIns: emu(before.bottom - after.bottom * box.height),
  };
  const own = normalizeOfficeTextLayout(ordinary.textLayout);
  const lines = shape.text.split("\n").length;
  const layout: OfficeTextLayout =
    own && own.paragraphs.length === lines
      ? own
      : { paragraphs: Array.from({ length: lines }, () => ({ level: 0, bullet: "none" as const })) };
  const bodyPr: OfficeTextBodyProperties = { ...(layout.bodyPr ?? {}), ...insets };
  return { ...next, shape: { ...ordinary, textLayout: { ...layout, bodyPr } } };
}

/**
 * The diagram's members as the shapes Convert to Shapes leaves, in their
 * order: the drawn ones only (an item the layout does not draw, and the
 * graphic's own frame, have no shape to become), each without its SmartArt
 * identity, in the diagram's group, its words where they were drawn.
 */
export function spreadsheetSmartArtMembersAsShapes<T extends { id: string }>(
  members: readonly T[],
  size: SpreadsheetMemberSize<T>,
): T[] {
  const byId = new Map(members.map((member) => [member.id, member]));
  const sizeOf = (member: Json) => size(member as unknown as T);
  return releaseSmartArtMembers(members, { keepGroup: true }).map((released) => {
    const original = byId.get(released.id);
    return (
      original
        ? wordsKeptInPlace(original as unknown as Json, released as unknown as Json, sizeOf)
        : released
    ) as T;
  });
}
