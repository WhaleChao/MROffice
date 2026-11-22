/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The worksheet Draw tab's pens: Excel for Mac's gallery on the shared pen
 * model (`lib/office-ink-pens.ts`, the menus of `lib/document-ink.ts`).
 *
 * Excel for Mac (native ribbon, EXCEL_SPREADSHEET_FEATURE_REPLICATION_TREE.md
 * §15.2) opens its Draw tab with a black pen 0.35 mm, a red pen 0.35 mm, a
 * grey pencil 0.5 mm and a yellow highlighter 6 mm, then Add Pen and Draw
 * with Trackpad. A pen's menu is the shared one (six thickness dots from
 * 0.25 mm to 6 mm, sixteen colours, eight effects, More Colors, Delete); the
 * default pens keep Excel's own widths until the user picks a dot.
 */

import { ptToPxRounded } from "@mrnote/core/web/lib/office-units.ts";
import {
  DOCUMENT_INK_PRESETS,
  DOCUMENT_INK_THICKNESSES,
  type DocumentInkPreset,
} from "@mrnote/core/web/lib/document-ink.ts";
import type { OfficeInkPenSpec } from "@mrnote/core/web/lib/office-ink-pens.ts";

/** 1 mm is 72/25.4 pt; the outline states points to a tenth. */
function mmToPt(mm: number): number {
  return Math.round(((mm * 72) / 25.4) * 10) / 10;
}

 
const SPREADSHEET_INK_PREVIOUS_DEFAULT_PENS: readonly DocumentInkPreset[] = [
  { id: "pen-black-035", tool: "pen", color: "#000000", widthPt: mmToPt(0.35) },
  { id: "pen-red-035", tool: "pen", color: "#e12b2b", widthPt: mmToPt(0.35) },
  { id: "pencil-gray-05", tool: "pencil", color: "#8a8a8a", widthPt: mmToPt(0.5) },
  { id: "highlighter-yellow-6", tool: "highlighter", color: "#ffff00", widthPt: mmToPt(6) },
];







export const SPREADSHEET_INK_DEFAULT_PENS: readonly DocumentInkPreset[] = DOCUMENT_INK_PRESETS.map(
  (pen) => ({ ...pen }),
);

/** The native Excel control each default pen stands in for, by pen id (the
 * four Excel for Mac opens with); the galaxy and wide pens are MRAI's own
 * and name no Excel control. */
export const SPREADSHEET_INK_DEFAULT_PEN_CONTROL_IDS: Readonly<Record<string, string>> = {
  "pen-black": "XLSCTL.DRAW.PEN.BLACK_035",
  "pen-red": "XLSCTL.DRAW.PEN.RED_035",
  "pencil-gray": "XLSCTL.DRAW.PEN.PENCIL_GRAY_05",
  "highlighter-yellow": "XLSCTL.DRAW.PEN.HIGHLIGHT_YELLOW_6",
};

export const SPREADSHEET_INK_PEN_SPEC: OfficeInkPenSpec = {
  defaults: SPREADSHEET_INK_DEFAULT_PENS,
  previousDefaults: [SPREADSHEET_INK_PREVIOUS_DEFAULT_PENS],
  widthsPt: [
    ...new Set([
      ...DOCUMENT_INK_THICKNESSES.map(({ pt }) => pt),
      ...SPREADSHEET_INK_DEFAULT_PENS.map((pen) => pen.widthPt),
      // The widths of the gallery the worksheet opened with before, so a
      // stored one still reads (and, untouched, opens as today's).
      ...SPREADSHEET_INK_PREVIOUS_DEFAULT_PENS.map((pen) => pen.widthPt),
    ]),
  ],
};

export const SPREADSHEET_INK_PEN_PROFILE_STORAGE_KEY = "mrai:spreadsheet:ink-pens:v1";

/** A pen's width on the worksheet, whose ink strokes are measured in CSS pixels. */
export function spreadsheetInkWidthPx(widthPt: number): number {
  return Math.max(0.5, ptToPxRounded(widthPt, 3));
}
