/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isLineCalloutShapeKind, type ShapeKind } from "@mrnote/core/web/lib/presentation-shapes.ts";

export interface LineCalloutDrawing {
  role: "shape";
  geometry: {
    schemaVersion: "drawingml-paths-v1";
    paths: Array<{ d: string; fill: "norm" | "none"; stroke: boolean }>;
  };
}

/**
 * Default DrawingML line callouts use the entire shape box for their text body;
 * leaders extend outside that box. Reuse the native path renderer so body fill
 * and outline remain independent (plain/accent callouts have no body outline).
 * Imported native geometry always takes precedence over these preset defaults.
 */
export function presentationLineCalloutDrawing(kind: ShapeKind): LineCalloutDrawing | null {
  if (!isLineCalloutShapeKind(kind)) return null;
  const segments = Number(kind.slice(-1));
  const leader = segments === 1
    ? "M-8.333 18.75 L-38.333 112.5"
    : segments === 2
      ? "M-8.333 18.75 L-16.667 18.75 L-46.667 112.5"
      : "M-8.333 18.75 L-16.667 18.75 L-16.667 100 L-8.333 112.963";
  const paths: LineCalloutDrawing["geometry"]["paths"] = [{
    d: "M0 0 L100 0 L100 100 L0 100 Z",
    fill: "norm",
    stroke: kind.includes("border"),
  }];
  if (kind.startsWith("accent")) {
    paths.push({ d: "M-8.333 0 Z L-8.333 100", fill: "none", stroke: true });
  }
  paths.push({ d: leader, fill: "none", stroke: true });
  return {
    role: "shape",
    geometry: { schemaVersion: "drawingml-paths-v1", paths },
  };
}
