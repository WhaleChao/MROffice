/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { assignGroup, expandSelectionForGroup, normalizeGroupId, ungroup } from "@mrnote/core/web/lib/presentation-crop-group.ts";
import { selectionBoundingBox } from "@mrnote/core/web/lib/presentation-multi-transform.ts";

export interface ChartDrawingFrame { left: number; top: number; width: number; height: number }
export interface ChartDrawing {
  [key: string]: unknown;
  id: string;
  groupId?: string;
  frame: ChartDrawingFrame;
}

export function chartDrawingGroupId(value: unknown): string | undefined {
  const id = normalizeGroupId(value);
  return id && id.length <= 200 ? id : undefined;
}

export function selectedChartDrawings<T extends ChartDrawing>(drawings: readonly T[], id: string): T[] {
  const selected = expandSelectionForGroup(drawings, id);
  return drawings.filter(drawing => selected.has(drawing.id));
}

export function chartDrawingBounds(drawings: readonly ChartDrawing[]): ChartDrawingFrame | null {
  return selectionBoundingBox(drawings.map(drawing => ({ blockId: drawing.id, startGeom: drawing.frame })));
}

export function groupChartDrawings<T extends ChartDrawing>(drawings: readonly T[], ids: readonly string[], groupId?: string): T[] | null {
  const selected = new Set(ids.flatMap(id => [...expandSelectionForGroup(drawings, id)]));
  if (selected.size < 2 || selected.size > 64 || drawings.filter(item => selected.has(item.id)).length !== selected.size) return null;
  return assignGroup(drawings, selected, groupId);
}

export function ungroupChartDrawings<T extends ChartDrawing>(drawings: readonly T[], id: string): T[] {
  return ungroup(drawings, expandSelectionForGroup(drawings, id));
}

/** An existing member handle moves/resizes its whole group in one transaction.
 * Keep relative geometry, clamp the whole result together, never each member. */
export function transformChartDrawingGroup<T extends ChartDrawing>(drawings: readonly T[], id: string, frame: ChartDrawingFrame): T[] | null {
  const owner = drawings.find(item => item.id === id);
  // A caller may hand over a richer preview record (sheet/object ids, rotation);
  // only the four geometry fields decide whether the frame is usable.
  const { left, top, width, height } = frame;
  if (!owner || ![left, top, width, height].every(Number.isFinite) || width <= 0 || height <= 0) return null;
  frame = { left, top, width, height };
  const selected = selectedChartDrawings(drawings, id);
  const resizing = frame.width !== owner.frame.width || frame.height !== owner.frame.height;
  const sx = frame.width / owner.frame.width, sy = frame.height / owner.frame.height;
  const transformed = selected.map(item => ({ ...item, frame: {
    left: frame.left + (item.frame.left - owner.frame.left) * sx,
    top: frame.top + (item.frame.top - owner.frame.top) * sy,
    width: item.frame.width * sx,
    height: item.frame.height * sy,
  } }));
  // Imported small members must remain movable at their exact source size.
  // The authoring minimum applies only when this gesture changes dimensions.
  if (transformed.some(item => (resizing && (item.frame.width < 16 || item.frame.height < 16)) || item.frame.width > 4000 || item.frame.height > 4000)) return null;
  const box = chartDrawingBounds(transformed);
  if (!box || box.width > 100000 || box.height > 100000) return null;
  const dx = Math.max(0, -box.left), dy = Math.max(0, -box.top);
  const byId = new Map(transformed.map(item => [item.id, { ...item, frame: {
    ...item.frame, left: item.frame.left + dx, top: item.frame.top + dy,
  } }]));
  return drawings.map(item => byId.get(item.id) ?? item);
}
