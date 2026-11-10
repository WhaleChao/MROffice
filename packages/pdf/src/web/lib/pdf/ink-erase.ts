/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Draw tab's erasers over PDF ink (Word's Draw ▸ Eraser, "Draw and
 * write with ink in Office", Microsoft Support: three erasers). The Stroke
 * Eraser takes a whole stroke the pointer passes over; the Small and Medium
 * Erasers rub out the ink under the pointer and leave the rest of the stroke
 * as pieces, cut flat at the eraser's edge, exactly as the Docs editor's
 * erasers do (`documentInkEraserCut`, the one geometry every editor shares).
 *
 * PDF ink keeps a stroke's points as page fractions; the eraser measures in
 * the page layer's CSS pixels, where a stroke's `width` is painted.
 */

import {
  documentInkEraserCut,
  documentInkEraserRadiusPx,
  documentInkPolylineBetween,
  type DocumentInkEraser,
  type DocumentInkPoint,
} from "@mrnote/core/web/lib/document-ink.ts";

import type { PdfInkStroke } from "../../components/notebook/contents/pdf/PdfInkOverlay";

/** The erasers of the Draw tab, the shared kinds. */
export type PdfInkEraserKind = DocumentInkEraser;

export interface PdfInkEraseResult {
  /** Strokes the eraser took, whole or in part (their ids). */
  readonly removedIds: readonly string[];
  /** What the point erasers left of the strokes they cut, as new strokes. */
  readonly added: readonly PdfInkStroke[];
}

/** A page layer's painted size (CSS px): where the eraser's reach is measured. */
export interface PdfInkPageSize {
  readonly width: number;
  readonly height: number;
}

/** A stroke's fraction points as page-layer pixels. */
export function pdfInkStrokePointsPx(
  stroke: Pick<PdfInkStroke, "points">,
  page: PdfInkPageSize,
): DocumentInkPoint[] {
  const points: DocumentInkPoint[] = [];
  for (let index = 0; index + 1 < stroke.points.length; index += 2) {
    points.push({
      x: stroke.points[index] * page.width,
      y: stroke.points[index + 1] * page.height,
    });
  }
  return points;
}

function fractionsOf(points: readonly DocumentInkPoint[], page: PdfInkPageSize): number[] {
  const out: number[] = [];
  for (const point of points) {
    out.push(
      Math.min(1, Math.max(0, point.x / page.width)),
      Math.min(1, Math.max(0, point.y / page.height)),
    );
  }
  return out;
}

function pieceId(): string {
  return globalThis.crypto?.randomUUID?.() ?? `ink-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

/**
 * What an eraser dragged along `eraserPathPx` (page-layer pixels; one point
 * for a tap) does to `strokes` on that page: null when it touched no ink.
 * A third-party stroke the source PDF paints itself is erased whole only
 * (its pieces could not be repainted), with any eraser.
 */
export function erasePdfInkStrokes(
  strokes: readonly PdfInkStroke[],
  eraserPathPx: readonly DocumentInkPoint[],
  kind: PdfInkEraserKind,
  page: PdfInkPageSize,
): PdfInkEraseResult | null {
  if (eraserPathPx.length === 0 || !(page.width > 0) || !(page.height > 0)) return null;
  const radius = documentInkEraserRadiusPx(kind);
  const removedIds: string[] = [];
  const added: PdfInkStroke[] = [];
  for (const stroke of strokes) {
    if (stroke.deleted) continue;
    const points = pdfInkStrokePointsPx(stroke, page);
    if (points.length === 0) continue;
    const widthPx = Number.isFinite(stroke.width) && stroke.width > 0 ? stroke.width : 0;
    const pieces = documentInkEraserCut(points, eraserPathPx, radius, widthPx);
    if (pieces === null) continue;
    removedIds.push(stroke.id);
    if (kind === "stroke" || stroke.nativeSource) continue;
    for (const piece of pieces) {
      const line = documentInkPolylineBetween(points, piece.from, piece.to);
      if (line.length < 2) continue;
      const fractions = fractionsOf(line, page);
      if (fractions.length < 4) continue;
      added.push({
        id: pieceId(),
        page: stroke.page,
        color: stroke.color,
        width: stroke.width,
        points: fractions,
        ...(stroke.penType ? { penType: stroke.penType } : {}),
      });
    }
  }
  return removedIds.length === 0 ? null : { removedIds, added };
}

/** The erase result applied to a page's strokes: what the layer now shows. */
export function applyPdfInkErase(
  strokes: readonly PdfInkStroke[],
  result: PdfInkEraseResult,
): PdfInkStroke[] {
  const removed = new Set(result.removedIds);
  return [...strokes.filter((stroke) => !removed.has(stroke.id)), ...result.added];
}

/**
 * One eraser drag's change to the ink a page window writes into the PDF
 * (its managed strokes and their tombstones): the strokes it took leave as
 * tombstones -- an original /Ink path too, by its own path -- and the pieces
 * a point eraser left come in as new strokes. Written as one commit, so the
 * whole drag is one step to undo, as Docs' eraser is one transaction.
 * Ids the window does not hold (a reader row with no PDF copy) are no
 * change here; null when nothing changes.
 */
export function planPdfManagedInkErase(
  strokes: readonly PdfInkStroke[],
  tombstones: readonly PdfInkStroke[],
  erase: PdfInkEraseResult,
): { strokes: PdfInkStroke[]; tombstones: PdfInkStroke[] } | null {
  const removed = new Set(erase.removedIds);
  const taken = strokes.filter((entry) => removed.has(entry.id) && !entry.deleted);
  if (taken.length === 0 && erase.added.length === 0) return null;
  const addedIds = new Set(erase.added.map((entry) => entry.id));
  const live = (entry: PdfInkStroke): PdfInkStroke => {
    const copy = { ...entry };
    delete copy.deleted;
    return copy;
  };
  return {
    strokes: [
      ...strokes.filter((entry) => !removed.has(entry.id) && !addedIds.has(entry.id)),
      ...erase.added.map(live),
    ],
    tombstones: [
      ...tombstones.filter((entry) => !removed.has(entry.id) && !addedIds.has(entry.id)),
      ...taken.map((entry) => ({ ...entry, deleted: true })),
    ],
  };
}
