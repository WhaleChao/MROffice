/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { readSmartArtGraphicFormat, smartArtGraphicFormatStyle } from "@mrnote/core/web/lib/spreadsheet/smartart-background.ts";
import type { CSSProperties } from "react";

const EMU_PER_PX = 9525;

type Record_ = Record<string, unknown>;

function record(value: unknown): Record_ | null {
  return value && typeof value === "object" && !Array.isArray(value) ? (value as Record_) : null;
}

/** The plate's marker: a block carrying it is a graphic plate, drawn and never edited. */
const PLATE_MARK = "smartArtGraphicPlate";

/** Whether a block is a SmartArt graphic plate ({@link withNativeSmartArtGraphicPlates}). */
export function isSmartArtGraphicPlate(block: unknown): boolean {
  return record(block)?.[PLATE_MARK] === true;
}

/** The id of the member a plate shows and hides with; null for any other block. */
export function smartArtGraphicPlateMemberId(block: unknown): string | null {
  const id = record(block)?.plateOf;
  return isSmartArtGraphicPlate(block) && typeof id === "string" ? id : null;
}

function finiteRect(value: unknown): { left: number; top: number; width: number; height: number } | null {
  const rect = record(value);
  if (!rect) return null;
  const [left, top, width, height] = [rect.left, rect.top, rect.width, rect.height];
  return typeof left === "number" &&
    typeof top === "number" &&
    typeof width === "number" &&
    typeof height === "number" &&
    [left, top, width, height].every(Number.isFinite) &&
    width > 0 &&
    height > 0
    ? { left, top, width, height }
    : null;
}

/** The diagram a member of a file's drawing belongs to, and what its graphic wears; null for any other block. */
function nativeGraphicOf(block: unknown): { key: string; format: unknown; frame: NonNullable<ReturnType<typeof finiteRect>> } | null {
  const view = record(record(record(block)?.source)?.nativeSmartArtView);
  if (!view || view.schemaVersion !== "pptx-smartart-drawing-v1") return null;
  const format = view.graphicFormat;
  if (!readSmartArtGraphicFormat({ smartArtBackground: format })) return null;
  const frame = finiteRect(view.frameEmu);
  if (!frame) return null;
  const groupId = record(block)?.groupId;
  const key = JSON.stringify([typeof groupId === "string" ? groupId : null, view.frameShapeId ?? null]);
  return { key, format, frame };
}

/** One plate per first member object, so a slide drawn again hands the views the same plate. */
const plates = new WeakMap<object, Record_>();

function plateFor(member: Record_, graphic: NonNullable<ReturnType<typeof nativeGraphicOf>>): Record_ {
  const known = plates.get(member);
  if (known && known.smartArtBackground === graphic.format && JSON.stringify(known.frame) === JSON.stringify(graphic.frame)) {
    return known;
  }
  const groupId = typeof member.groupId === "string" ? member.groupId : String(member.id ?? "smartart");
  const plate: Record_ = {
    id: `${groupId}-graphic`,
    type: "shape",
    shape: "rectangle",
    ...graphic.frame,
    frame: graphic.frame,
    // Painted where the diagram is painted, as one graphic.
    ...(member.zIndex !== undefined ? { zIndex: member.zIndex } : {}),
    ...(member.sourceOrder !== undefined ? { sourceOrder: member.sourceOrder } : {}),
    smartArtBackground: graphic.format,
    // Shown and hidden with the diagram it belongs to (its play state).
    plateOf: member.id,
    [PLATE_MARK]: true,
  };
  plates.set(member, plate);
  return plate;
}

/**
 * The blocks with a plate before the first member of each file's SmartArt
 * graphic that wears a fill, line or effect of its own; the same array when
 * there is none. Only the members still in the list (a diagram the Selection
 * Pane hides is filtered out before) bring their plate.
 */
export function withNativeSmartArtGraphicPlates<T extends object>(blocks: readonly T[]): readonly T[] {
  let out: T[] | null = null;
  const seen = new Set<string>();
  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index];
    const graphic = nativeGraphicOf(block);
    if (graphic && !seen.has(graphic.key)) {
      seen.add(graphic.key);
      out ??= blocks.slice(0, index);
      out.push(plateFor(block as Record_, graphic) as unknown as T);
    }
    out?.push(block);
  }
  return out ?? blocks;
}

/**
 * The CSS of a plate on a slide `bounds` (EMU) across: where its frame
 * stands, painted as the graphic's own formatting is painted
 * ({@link smartArtGraphicFormatStyle}), never taking a press. Null for a
 * block that is no plate or wears nothing.
 */
export function smartArtGraphicPlateStyle(
  block: unknown,
  bounds: { readonly width: number; readonly height: number },
): CSSProperties | null {
  if (!isSmartArtGraphicPlate(block)) return null;
  const frame = finiteRect(record(block)?.frame);
  if (!frame || !(bounds.width > 0) || !(bounds.height > 0)) return null;
  const paint = smartArtGraphicFormatStyle(readSmartArtGraphicFormat(block), {
    width: frame.width / EMU_PER_PX,
    height: frame.height / EMU_PER_PX,
  });
  if (!paint) return null;
  return {
    position: "absolute",
    left: `${(frame.left / bounds.width) * 100}%`,
    top: `${(frame.top / bounds.height) * 100}%`,
    width: `${(frame.width / bounds.width) * 100}%`,
    height: `${(frame.height / bounds.height) * 100}%`,
    // A stacking layer of its own, in the slide's paint order, as every object.
    zIndex: 0,
    pointerEvents: "none",
    ...paint,
  };
}
