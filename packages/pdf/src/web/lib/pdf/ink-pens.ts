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
import { PDF_INK_MAX_WIDTH, type InkPenType } from "@mrnote/core/web/lib/pdf/ink.ts";

/** The tools PDF ink draws with. */
export const PDF_INK_PEN_TOOLS: readonly DocumentInkTool[] = ["pen", "highlighter"];

/** The gallery as the page window first shows it: the shared six pens less
 * the pencil and the effect pen (a black and a red pen, a yellow
 * highlighter and a heavier black pen). */
const PDF_INK_SHARED_PENS: readonly DocumentInkPreset[] = DOCUMENT_INK_PRESETS.filter(
  (pen) => PDF_INK_PEN_TOOLS.includes(pen.tool) && !pen.effect,
).map((pen) => ({ ...pen }));

/** The gallery's pens, the red pen first: a PDF is marked up in red (the
 * window's pen was red before the gallery), while a document's first pen is
 * black. */
export const PDF_INK_DEFAULT_PENS: readonly DocumentInkPreset[] = [
  ...PDF_INK_SHARED_PENS.filter((pen) => pen.id === "pen-red"),
  ...PDF_INK_SHARED_PENS.filter((pen) => pen.id !== "pen-red"),
];

export const PDF_INK_PEN_SPEC: OfficeInkPenSpec = {
  defaults: PDF_INK_DEFAULT_PENS,
  widthsPt: [
    ...new Set([
      ...DOCUMENT_INK_THICKNESSES.map(({ pt }) => pt),
      ...PDF_INK_DEFAULT_PENS.map((pen) => pen.widthPt),
    ]),
  ],
  tools: PDF_INK_PEN_TOOLS,
};

/** Kept in this browser, as the other editors keep theirs. */
export const PDF_INK_PEN_PROFILE_STORAGE_KEY = "mrai:pdf:ink-pens:v1";

/** A pen's stroke width on the page layer (CSS px, 96 per inch), capped at
 * {@link PDF_INK_MAX_WIDTH}. */
export function pdfInkPenWidthPx(widthPt: number): number {
  const px = ptToPxRounded(widthPt);
  return Math.min(PDF_INK_MAX_WIDTH, Math.max(0.5, px));
}

/** The page window's stroke kind for a pen. */
export function pdfInkPenType(tool: DocumentInkTool): InkPenType {
  return tool === "highlighter" ? "highlighter" : "pen";
}
