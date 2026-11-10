/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A shape's polygon outline as the deck stores and draws it: "x,y x,y ..."
 * points in the shape's own 0-100 box. Reading and writing that string is
 * all the browser needs; the boolean geometry of Merge Shapes runs on the
 * server (`lib/presentation-shape-merge.ts`, engine `tools`).
 */

export interface Point {
  x: number;
  y: number;
}

/** Parse a "x,y x,y ..." polygon-points string (the shape render's format). */
export function parsePolygonString(value: string): Point[] {
  const out: Point[] = [];
  for (const pair of value.trim().split(/\s+/)) {
    const [x, y] = pair.split(",").map(Number);
    if (Number.isFinite(x) && Number.isFinite(y)) out.push({ x, y });
  }
  return out;
}

/** Format a polygon as the render's "x,y x,y ..." string (<= 2 decimals). */
export function formatPolygonString(points: Point[]): string {
  const n = (v: number) => (Number.isInteger(v) ? String(v) : String(Number(v.toFixed(2))));
  return points.map((p) => `${n(p.x)},${n(p.y)}`).join(" ");
}
