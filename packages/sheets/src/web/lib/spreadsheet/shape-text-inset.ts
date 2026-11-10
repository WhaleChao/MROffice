/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where a worksheet shape's words are laid out inside its element: the view's
 * answer (SpreadsheetShapeView), and what Convert to Shapes reads to keep a
 * SmartArt member's words where the graphic drew them once the member is an
 * ordinary shape (`lib/spreadsheet/smartart-convert`).
 */

import { officePresetTextRect } from "@mrnote/core/web/lib/office-shape-adjust-handles.ts";
import { SMARTART_TEXT_BOX_UNIT, type SmartArtTextBox } from "@mrnote/core/web/lib/office-smartart/text-box.ts";
import { spreadsheetShapeDrawingPreset } from "./shape-preset";
import { spreadsheetShapeGeometryInset } from "./shape-edit-points";
import type { SpreadsheetShapeStyle } from "@mrnote/core/web/lib/spreadsheet/shapes.ts";

/**
 * Where a sheet shape's LABEL sits inside its element: the drawn geometry's
 * box (`spreadsheetShapeGeometryInset`) narrowed to the preset's own text
 * rectangle (`a:rect`, the one Docs and the deck lay text out in), evaluated
 * with the shape's adjust values at its aspect. Excel lays the words out in
 * that rectangle, the body insets measured from its edges.
 *
 * A SmartArt member with a text box of its own (`smartArtTextBox`, Office's
 * `dsp:txXfrm`, kept on the sheet object) lays its words out there instead.
 */
/** How much wider than its text box a SmartArt label may run before it wraps
 * (a share of the box; `spreadsheet_drawing_pdf.SMARTART_LABEL_WRAP_SLACK`). */
export const SMARTART_LABEL_WRAP_SLACK = 0.03;

export function spreadsheetShapeTextInset(
  shape: SpreadsheetShapeStyle,
  widthPx = 100,
  heightPx = 100,
  smartArtTextBox?: SmartArtTextBox | null,
) {
  const box = spreadsheetShapeGeometryInset(shape);
  if (smartArtTextBox) {
    // The diagram's own rectangle for these words, in the shape's box: a
    // Segmented Cycle wedge's box is the whole disc and its words sit in the
    // wedge. Centred in the box, five wedges drew five labels on top of each
    // other in the middle of the disc.
    const spanX = 1 - box.left - box.right;
    const spanY = 1 - box.top - box.bottom;
    const unit = SMARTART_TEXT_BOX_UNIT;
    // The engine fitted these words to the rectangle in Office's measures; the
    // faces drawn here run a hair wider (a digit in Tahoma 0.546 em against
    // Calibri's 0.507), so a line Office fills exactly -- Basic Venn's "项目"
    // in its 48 pt -- broke a character early. The words may reach a little
    // past it on the side their alignment leaves open (the print does too).
    const slack = spanX * (smartArtTextBox.width / unit) * SMARTART_LABEL_WRAP_SLACK;
    const align = shape.textAlign;
    const slackLeft = align === "right" ? slack : align === "left" || align === "justify" ? 0 : slack / 2;
    return {
      left: box.left + spanX * (smartArtTextBox.left / unit) - slackLeft,
      top: box.top + spanY * (smartArtTextBox.top / unit),
      right:
        box.right +
        spanX * (1 - (smartArtTextBox.left + smartArtTextBox.width) / unit) -
        (slack - slackLeft),
      bottom:
        box.bottom +
        spanY * (1 - (smartArtTextBox.top + smartArtTextBox.height) / unit),
    };
  }
  // A SmartArt member's words are laid out in the DIAGRAM's text box for that
  // shape (`dsp:txXfrm`), not its preset's `a:rect`: the engine places each
  // label over its whole node, the slide draws it there, MRAI's own export
  // writes the member's full box as its `dsp:txXfrm`, and the member's type
  // was fitted to that box. Narrowing an ellipse member to its inscribed
  // rectangle (§SH15) left a 100 px circle's 14 pt label about 50 px of
  // line, and "项目 1" broke into two lines. Every member carries the style
  // marker (§4), which is what says the shape is a diagram's.
  if (shape.smartArtStyle) return box;
  const preset = spreadsheetShapeDrawingPreset(shape);
  const innerWidth = widthPx * (1 - box.left - box.right);
  const innerHeight = heightPx * (1 - box.top - box.bottom);
  const rect = preset
    ? officePresetTextRect(preset, innerWidth, innerHeight, shape.adjustments)
    : null;
  if (!rect) return box;
  const spanX = 1 - box.left - box.right;
  const spanY = 1 - box.top - box.bottom;
  return {
    left: box.left + spanX * rect.left,
    top: box.top + spanY * rect.top,
    right: box.right + spanX * (1 - rect.right),
    bottom: box.bottom + spanY * (1 - rect.bottom),
  };
}
