/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ptToPxRounded } from "@mrnote/core/web/lib/office-units.ts";
import {
  DOCUMENT_INK_PRESETS,
  DOCUMENT_INK_THICKNESSES,
  type DocumentInkPreset,
  type DocumentInkTool,
} from "@mrnote/core/web/lib/document-ink.ts";
import type { OfficeInkPenSpec } from "@mrnote/ink/web/lib/office-ink-pens.ts";

/** 1 mm is 72/25.4 pt; widths are stated in points to a tenth. */
function mmToPt(mm: number): number {
  return Math.round(((mm * 72) / 25.4) * 10) / 10;
}

/** Pen, pencil and highlighter: the tools every MRAI Draw tab draws with. */
export const PRESENTATION_INK_TOOLS: readonly DocumentInkTool[] = ["pen", "pencil", "highlighter"];

 
const PRESENTATION_INK_PREVIOUS_DEFAULT_PENS: readonly DocumentInkPreset[] = [
  { id: "pen-black-05", tool: "pen", color: "#000000", widthPt: mmToPt(0.5) },
  { id: "pen-red-1", tool: "pen", color: "#ff0000", widthPt: mmToPt(1) },
  { id: "pen-blue-1", tool: "pen", color: "#0070c0", widthPt: mmToPt(1) },
  { id: "highlighter-yellow-5", tool: "highlighter", color: "#ffff00", widthPt: mmToPt(5) },
];







export const PRESENTATION_INK_DEFAULT_PENS: readonly DocumentInkPreset[] = DOCUMENT_INK_PRESETS.map(
  (pen) => ({ ...pen }),
);

export const PRESENTATION_INK_PEN_SPEC: OfficeInkPenSpec = {
  defaults: PRESENTATION_INK_DEFAULT_PENS,
  previousDefaults: [PRESENTATION_INK_PREVIOUS_DEFAULT_PENS],
  widthsPt: [
    ...new Set([
      ...DOCUMENT_INK_THICKNESSES.map(({ pt }) => pt),
      ...PRESENTATION_INK_DEFAULT_PENS.map((pen) => pen.widthPt),
      // The widths of the gallery the deck opened with before, so a stored
      // one still reads (and, untouched, opens as today's).
      ...PRESENTATION_INK_PREVIOUS_DEFAULT_PENS.map((pen) => pen.widthPt),
    ]),
  ],
  tools: PRESENTATION_INK_TOOLS,
};

/** v2: pens in points on the shared model (v1 kept integer pixel widths). */
export const PRESENTATION_INK_PEN_PROFILE_STORAGE_KEY = "mrai:presentation:ink-pens:v2";

/** A pen's stroke on the slide, in the slide's CSS pixels (96 per inch). */
export function presentationInkStrokePx(widthPt: number): number {
  return Math.max(0.5, ptToPxRounded(widthPt));
}

/** A width in points as the millimetres a pen's name states (0.5, 1, 5). */
export function presentationInkWidthMm(widthPt: number): number {
  const exact = DOCUMENT_INK_THICKNESSES.find((thickness) => thickness.pt === widthPt);
  if (exact) return exact.mm;
  // Points are stated to a tenth, so 5 mm reads back as 5.01: a value that
  // close to a tenth of a millimetre is that tenth.
  const mm = (widthPt * 25.4) / 72;
  const tenth = Math.round(mm * 10) / 10;
  return Math.abs(mm - tenth) < 0.02 ? tenth : Math.round(mm * 100) / 100;
}

/** The Standard Color names a pen colour carries, by `#rrggbb`. */
export const PRESENTATION_INK_COLOR_NAME_KEYS: Readonly<Record<string, string>> = {
  "#000000": "black",
  "#ffffff": "white",
  "#c00000": "darkRed",
  "#ff0000": "red",
  "#ffc000": "orange",
  "#ffff00": "yellow",
  "#92d050": "lightGreen",
  "#00b050": "green",
  "#00b0f0": "lightBlue",
  "#0070c0": "blue",
  "#002060": "darkBlue",
  "#7030a0": "purple",
};
