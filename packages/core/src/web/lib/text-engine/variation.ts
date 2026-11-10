/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A glyph's advance in a variable face at the instance the browser draws.
 *
 * Google serves CJK faces as variable fonts whose default instance is Thin
 * (Noto Sans SC: wght 100..900, default 100); the hmtx advances are the
 * Thin ones, and both engines draw a 400 face at wght 400 -- "«" is 0.429 em
 * in hmtx and 0.479 em drawn. The browsers put the wght axis at the asked
 * weight clamped to the face's weight descriptor (CSS Fonts 4 §5.2), so the
 * engine applies the HVAR deltas at that location. A face whose advances
 * vary with an axis the engine does not place (optical size follows the
 * font size), or that has no HVAR, is not measured here.
 */
import { createSfntAdvanceDeltas, sfntVariationAxes } from "../pdf/engine/native/sfnt/sfnt-variation";
import type { SfntProgram } from "../pdf/engine/native/sfnt/sfnt-reader";
import type { FaceDeclaration } from "./face-index";
import type { FaceMetrics } from "./face-metrics";

const deltasByInstance = new Map<string, ((glyphId: number) => number) | null>();

/** The advance in font units, or NaN when the engine cannot place the face. */
export function glyphAdvanceUnits(face: FaceDeclaration, metrics: FaceMetrics, glyph: number, weight: number): number {
  const base = metrics.advances[glyph];
  const tables = metrics.variationTables;
  if (!tables) return base;
  // A face that states its own wght (an alias drawn at its stand-in's
  // weight) is drawn there whatever weight is asked.
  const wght = face.wght ?? Math.min(Math.max(weight, face.weightMin), face.weightMax);
  const key = `${face.url}\u0000${wght}`;
  let deltas = deltasByInstance.get(key);
  if (deltas === undefined) {
    deltas = null;
    try {
      const axes = sfntVariationAxes({ tables } as unknown as SfntProgram);
      // Only wght is placed; an axis the browser sets from something else
      // (opsz from the font size) leaves the face to the browser.
      if (axes.every((axis) => axis.tag === "wght" || axis.tag === "wdth" || axis.tag === "ital" || axis.tag === "slnt")) {
        deltas = createSfntAdvanceDeltas(tables, { wght });
      }
    } catch {
      deltas = null;
    }
    deltasByInstance.set(key, deltas);
  }
  if (!deltas) return Number.NaN;
  return base + deltas(glyph);
}
