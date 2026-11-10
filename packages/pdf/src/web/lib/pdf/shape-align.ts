/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * 对齐 / 分布 for the objects selected on a PDF page (PowerPoint's, WPS's and
 * ONLYOFFICE's Align menu), in page fractions. The placing is the deck
 * editor's own (lib/presentation-align): two or more objects line up on the
 * box around them, three or more are spaced evenly between the outer two (by
 * centre order, as PowerPoint orders them), and one object lines up on the
 * page. Each object's move is a frame-to-frame transform for the page's
 * writer (PdfNativeVectorGroupTransform "bounds").
 */
import {
  alignBlocks,
  alignBlocksToBounds,
  distributeBlocks,
  type AlignEdge,
  type BlockGeometry,
  type DistributeAxis,
} from "@mrnote/core/web/lib/presentation-align.ts";
import type {
  PdfNativeVectorGroupTransform,
  PdfNativeVectorNormalizedBounds,
} from "./native-vector-selection-edit";

export type PdfShapeAlignCommand =
  | Readonly<{ kind: "align"; edge: AlignEdge }>
  | Readonly<{ kind: "distribute"; axis: DistributeAxis }>;

export type PdfShapeAlignFrame = PdfNativeVectorNormalizedBounds;

/** Page fractions in the integer units the deck's placing rounds to. */
const UNITS = 10_000_000;
/** A move smaller than this (a page fraction) is no move. */
const STILL = 1e-6;

const toBlock = (frame: PdfShapeAlignFrame): BlockGeometry => ({
  left: Math.round(frame.x * UNITS),
  top: Math.round(frame.y * UNITS),
  width: Math.round(frame.width * UNITS),
  height: Math.round(frame.height * UNITS),
});

/** Whether the command can act on this many objects (分布 needs three). */
export function pdfShapeAlignAvailable(command: PdfShapeAlignCommand, count: number): boolean {
  return command.kind === "align" ? count >= 1 : count >= 3;
}

/** One object's move: its place in the frames given, and its transform. */
export type PdfShapeAlignMove = Readonly<{ index: number; transform: PdfNativeVectorGroupTransform }>;

/**
 * The moves of the objects that move (the writer refuses an unchanged edit,
 * so one already in place is left out); empty when none does.
 */
export function planPdfShapeAlignment(
  frames: readonly PdfShapeAlignFrame[],
  command: PdfShapeAlignCommand,
): PdfShapeAlignMove[] {
  if (!pdfShapeAlignAvailable(command, frames.length)) return [];
  if (frames.some((frame) => ![frame.x, frame.y, frame.width, frame.height].every(Number.isFinite))) {
    return [];
  }
  const blocks = frames.map(toBlock);
  const placed =
    command.kind === "distribute"
      ? distributeBlocks(blocks, command.axis)
      : frames.length === 1
        ? alignBlocksToBounds(blocks, command.edge, { width: UNITS, height: UNITS })
        : alignBlocks(blocks, command.edge);
  const moves: PdfShapeAlignMove[] = [];
  frames.forEach((frame, index) => {
    const dx = placed[index].left / UNITS - frame.x;
    const dy = placed[index].top / UNITS - frame.y;
    if (Math.abs(dx) <= STILL && Math.abs(dy) <= STILL) return;
    moves.push({
      index,
      transform: {
        kind: "bounds",
        fromBounds: frame,
        toBounds: {
          x: Math.abs(dx) > STILL ? frame.x + dx : frame.x,
          y: Math.abs(dy) > STILL ? frame.y + dy : frame.y,
          width: frame.width,
          height: frame.height,
        },
      },
    });
  });
  return moves;
}
