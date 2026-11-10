/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { TextBox } from "./text-boxes";
import type { PdfTextBoxTypographyFields } from "@mrnote/core/web/lib/pdf/text-box-typography.ts";
import {
  textBoxesFromPdfTextContent,
  type PdfTextContentLike,
} from "./text-box-extract";
import {
  normalizePdfTextStyleRuns,
  type PdfTextStyleRun,
} from "@mrnote/core/web/lib/pdf/text-style-runs.ts";
import {
  normalizePdfTextLineLayout,
  type PdfTextLineLayout,
} from "@mrnote/core/web/lib/pdf/text-line-layout.ts";

/**
 * Structural PDF text-box target shared by view mode, edit mode, and AI command
 * context. Keeping this adapter in one place prevents metadata drift: formula
 * boxes must carry both visible fallback text and semantic LaTeX everywhere the
 * assistant or editor can act on them.
 */
/** An edited box also carries its whole-box typography (text-box-typography.ts). */
export interface PdfTextBoxCommandTargetLike extends PdfTextBoxTypographyFields {
  id: string;
  page: number;
  kind?: "text" | "formula";
  latex?: string;
  text: string;
  x: number;
  y: number;
  width: number;
  height: number;
  /** Clockwise page-space rotation carried from the extracted PDF run. */
  angle?: number;
  fontSize?: number;
  fontFamily?: string;
  /** Current-document font face for rendering only; never persisted as a family. */
  displayFontFamily?: string;
  sourceFontName?: string;
  color?: string;
  align?: "left" | "center" | "right";
  bold?: boolean;
  italic?: boolean;
  /** Whole-box text-box formats (an edited box carries them). */
  underline?: boolean;
  strikethrough?: boolean;
  justify?: boolean;
  runs?: PdfTextStyleRun[];
  /** The Office family a saved text box is set in (text-box-runs.ts). */
  officeFamily?: string;
  lineLayout?: PdfTextLineLayout[];
  edited?: boolean;
  created?: boolean;
}

export function textBoxCommandTargetFromTextBox(
  box: TextBox,
  page: number,
): PdfTextBoxCommandTargetLike {
  return {
    id: box.id,
    page,
    text: box.text,
    x: box.x,
    y: box.y,
    width: box.width,
    height: box.height,
    ...(Number.isFinite(box.angle) && box.angle ? { angle: box.angle } : {}),
    fontSize: box.fontSize,
    fontFamily: box.fontFamily,
    displayFontFamily: box.displayFontFamily,
    sourceFontName: box.sourceFontName,
    color: box.color,
    bold: box.bold,
    italic: box.italic,
    runs: normalizePdfTextStyleRuns(box.runs, {
      expectedText: box.text,
      allowDisplayFont: true,
    }),
    lineLayout: normalizePdfTextLineLayout(box.lineLayout, box.text),
    kind: box.kind,
    latex: box.latex,
  };
}

interface PdfPageCommonObjectsLike {
  has?: (id: string) => boolean;
  get?: (id: string) => unknown;
}

interface PdfPageForTextBoxCommandTargets {
  getViewport: (args: { scale: number; rotation?: number }) => {
    width: number;
    height: number;
    transform: number[];
  };
  commonObjs?: PdfPageCommonObjectsLike;
}

export function textBoxCommandTargetsFromPdfPageTextContent(
  page: PdfPageForTextBoxCommandTargets,
  textContent: PdfTextContentLike,
  pageIndex: number,
  rotation = 0,
): PdfTextBoxCommandTargetLike[] {
  const viewport = page.getViewport({ scale: 1.0, rotation });
  return textBoxesFromPdfTextContent(textContent, viewport, `p${pageIndex}`, {
    commonObjs: page.commonObjs,
  }).map((box) => textBoxCommandTargetFromTextBox(box, pageIndex));
}
