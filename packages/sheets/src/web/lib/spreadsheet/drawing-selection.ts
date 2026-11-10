/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { expandSelectionForGroup } from "@mrnote/core/web/lib/presentation-crop-group.ts";
import { transformChartDrawingGroup, type ChartDrawing, type ChartDrawingFrame } from "./chart-groups";

/** Transient selection only; worksheet/native identities remain on the model. */
export function selectSpreadsheetDrawing<T extends ChartDrawing>(items: readonly T[], key: string, current: readonly string[], additive: boolean) {
  const visible = items.filter(item => !item.hidden);
  if (!visible.some(item => item.id === key) || new Set(items.map(item=>item.id)).size !== items.length) return null;
  const expanded = [...expandSelectionForGroup(visible, key)];
  const selected = new Set(additive ? current.filter(id => visible.some(item=>item.id === id)) : []);
  if (additive && expanded.every(id => selected.has(id))) expanded.forEach(id => selected.delete(id));
  else expanded.forEach(id => selected.add(id));
  return { ids: visible.filter(item=>selected.has(item.id)).map(item=>item.id), primary: selected.has(key) ? key : visible.find(item=>selected.has(item.id))?.id ?? null };
}

export function transformSpreadsheetDrawingSelection<T extends ChartDrawing>(items: readonly T[], ids: readonly string[], key: string, frame: ChartDrawingFrame): T[] | null {
  if (!ids.includes(key) || ids.length < 2 || ids.length > 64) return null;
  const selected = new Set(ids);
  const marker = "spreadsheet-transient-selection";
  const temporary = items.map(item=>({...item,groupId:selected.has(item.id)?marker:undefined}));
  const transformed = transformChartDrawingGroup(temporary,key,frame);
  return transformed?.map((item,index)=>selected.has(item.id)?{...items[index],frame:item.frame}:items[index]) ?? null;
}

export const SHEET_DRAWING_HANDLE = { "north-west":"nw", north:"n", "north-east":"ne", east:"e", "south-east":"se", south:"s", "south-west":"sw", west:"w" } as const;

/** The worksheet's handle name for a shared drawing handle (`SHEET_DRAWING_HANDLE` turned round). */
export const SHEET_HANDLE_NAME = { nw:"north-west", n:"north", ne:"north-east", e:"east", se:"south-east", s:"south", sw:"south-west", w:"west" } as const;

/** Keep the rotation target clear of the worksheet's sticky column headings. */
export function spreadsheetDrawingRotateAtBottom(top: number, height: number, rotation: number, visibleTop: number): boolean {
  const center = top + height / 2;
  const offset = (height / 2 + 15) * Math.cos(rotation * Math.PI / 180);
  return center - offset - 7 < visibleTop && center + offset - 7 >= visibleTop;
}
