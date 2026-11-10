/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Format Shape > Text Box margins on a text box: the text-to-frame distance
 * (`a:bodyPr@lIns/rIns/tIns/bIns`), stored in integer EMU on `textInsets`.
 * The window's margin fields and an editor command both commit what
 * `textBlockWithInset` returns.
 */

import type {
  PresentationTextBlock,
  PresentationTextInsets,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { PRESENTATION_GEOMETRY_MAX_EMU } from "@mrnote/core/web/lib/presentation-geometry-limit.ts";
import { isImportedPptxBlock } from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import { isTextBoxBlock } from "@mrnote/core/web/lib/presentation-slide-buttons.ts";

/** PowerPoint's default text box margins: 0.1" left/right, 0.05" top/bottom. */
export const DEFAULT_PRESENTATION_TEXT_INSETS: PresentationTextInsets = {
  left: 91_440,
  right: 91_440,
  top: 45_720,
  bottom: 45_720,
};
export const PRESENTATION_TEXT_INSET_EDGES = [
  "left",
  "right",
  "top",
  "bottom",
] as const;
export type PresentationTextInsetEdge = (typeof PRESENTATION_TEXT_INSET_EDGES)[number];
export const PRESENTATION_EMU_PER_POINT = 12_700;

function asRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

/** A box's four margins (absent: the defaults), or null when any is invalid. */
export function normalizedPresentationTextInsets(
  value: unknown,
): PresentationTextInsets | null {
  if (value === undefined) return { ...DEFAULT_PRESENTATION_TEXT_INSETS };
  const record = asRecord(value);
  if (!record) return null;
  const next = {} as PresentationTextInsets;
  for (const edge of PRESENTATION_TEXT_INSET_EDGES) {
    const raw = record[edge];
    if (
      typeof raw !== "number" ||
      !Number.isInteger(raw) ||
      raw < 0 ||
      raw > PRESENTATION_GEOMETRY_MAX_EMU
    ) {
      return null;
    }
    next[edge] = raw;
  }
  return next;
}

/** A margin in EMU as the field shows it, in points to one decimal. */
export function presentationTextInsetPoints(value: number): string {
  return String(Math.round((value / PRESENTATION_EMU_PER_POINT) * 10) / 10);
}

/** A typed margin in points (one decimal) as EMU, or null. */
export function parsePresentationTextInsetPoints(value: string): number | null {
  const trimmed = value.trim();
  if (!/^(?:\d+(?:\.\d)?|\.\d)$/.test(trimmed)) return null;
  const points = Number(trimmed);
  if (!Number.isFinite(points)) return null;
  const emu = Math.round(points * PRESENTATION_EMU_PER_POINT);
  return emu >= 0 && emu <= PRESENTATION_GEOMETRY_MAX_EMU ? emu : null;
}

/** An imported box takes margins only when it is the slide's own shape. */
export function presentationTextInsetsAreWritable(
  block: PresentationTextBlock,
): boolean {
  if (!isImportedPptxBlock(block)) return true;
  const source = asRecord(block.source);
  return Boolean(
    source &&
      typeof source.slideId === "string" &&
      source.slideId &&
      typeof source.shapeId === "string" &&
      source.shapeId &&
      source.inheritedFrom === undefined,
  );
}

/** The text box with one margin set, in EMU; null when it cannot take it or already has it. */
export function textBlockWithInset(
  block: PresentationTextBlock,
  edge: PresentationTextInsetEdge,
  emu: number,
): PresentationTextBlock | null {
  if (!isTextBoxBlock(block) || !presentationTextInsetsAreWritable(block)) return null;
  const current = normalizedPresentationTextInsets(block.textInsets);
  if (!current || current[edge] === emu) return null;
  return { ...block, textInsets: { ...current, [edge]: emu } };
}
