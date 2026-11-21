/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How each path of a custom outline is painted, when its paths paint apart.
 *
 * `a:custGeom` holds a LIST of paths, each filled and outlined on its own
 * terms (`a:path@fill`, `a:path@stroke`, ECMA-376 §20.1.9.15): Google Docs
 * writes a bent arrow as its body, its fold in `darkenLess` and its outline
 * with `fill="none"`; Word's conversion of a legacy arc is a filled wedge that
 * is not outlined beside the arc that is. The persisted envelope
 * (`PresentationCustomPath`) is ONE path of several figures, so such a
 * geometry is the envelope of all its paths in order -- each opens with its
 * own `moveTo` -- plus this list, which says where each path starts and how
 * it is painted. The fill modes are `NativePathFill`, shaded by the one
 * `nativePathFillColor` the deck, the sheet's diagrams and the charts use.
 *
 * Server twin: `normalize_custom_path_paints` /
 * `custom_path_painted_segments` in `app/services/office_custom_geometry.py`.
 */
import { nativePathFillColor, type NativePathFill } from "./native-drawing-geometry";
import {
  presentationCustomPathCommandsToSvgPath,
  type PresentationCustomPath,
} from "./presentation-freeform";

export interface OfficeCustomPathPaint {
  /** How many of the envelope's commands this path holds, from its `moveTo`. */
  readonly commands: number;
  readonly fill: NativePathFill;
  readonly stroke: boolean;
}

const FILL_MODES: ReadonlySet<string> = new Set(["norm", "none", "darken", "darkenLess", "lighten", "lightenLess"]);
/** The most paths one geometry may paint apart (the drawing receipt's bound). */
export const OFFICE_CUSTOM_PATH_MAX_PAINTED_PATHS = 64;

/**
 * `value` as the paint of `path`'s paths, or null: every count lands on a
 * `moveTo`, the counts cover the commands exactly, and the paths do not all
 * paint alike (the envelope alone says that).
 */
export function normalizeOfficeCustomPathPaints(
  value: unknown,
  path: PresentationCustomPath | null | undefined,
): readonly OfficeCustomPathPaint[] | null {
  if (!path || !Array.isArray(value) || value.length < 2 || value.length > OFFICE_CUSTOM_PATH_MAX_PAINTED_PATHS) {
    return null;
  }
  const paints: OfficeCustomPathPaint[] = [];
  let index = 0;
  for (const raw of value) {
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
    const record = raw as Record<string, unknown>;
    const keys = Object.keys(record);
    if (keys.length !== 3 || !keys.every((key) => key === "commands" || key === "fill" || key === "stroke")) return null;
    const { commands, fill, stroke } = record;
    if (
      typeof commands !== "number" ||
      !Number.isInteger(commands) ||
      commands < 2 ||
      typeof fill !== "string" ||
      !FILL_MODES.has(fill) ||
      typeof stroke !== "boolean" ||
      path.commands[index]?.type !== "moveTo"
    ) {
      return null;
    }
    index += commands;
    if (index > path.commands.length) return null;
    paints.push({ commands, fill: fill as NativePathFill, stroke });
  }
  if (index !== path.commands.length) return null;
  if (new Set(paints.map((paint) => `${paint.fill}/${paint.stroke}`)).size < 2) return null;
  return paints;
}

export interface OfficeCustomPathPaintedSegment {
  /** SVG path data in the envelope's own grid. */
  readonly d: string;
  readonly fill: NativePathFill;
  readonly stroke: boolean;
}

/**
 * Each painted path of `path` as SVG path data with its paint. `paints` must
 * be {@link normalizeOfficeCustomPathPaints} of `path`. A path that ends with
 * a `close` is closed; the last one is closed when the envelope is.
 */
export function officeCustomPathPaintedSegments(
  path: PresentationCustomPath,
  paints: readonly OfficeCustomPathPaint[],
): OfficeCustomPathPaintedSegment[] {
  const segments: OfficeCustomPathPaintedSegment[] = [];
  let index = 0;
  paints.forEach((paint, number) => {
    let commands = path.commands.slice(index, index + paint.commands);
    index += paint.commands;
    let closed = path.closed;
    if (number < paints.length - 1) {
      closed = commands[commands.length - 1]?.type === "close";
      if (closed) commands = commands.slice(0, -1);
    }
    segments.push({ d: presentationCustomPathCommandsToSvgPath(commands, closed), fill: paint.fill, stroke: paint.stroke });
  });
  return segments;
}

/**
 * The colour one painted path is filled with: none, the shape's own fill, or
 * that fill darkened or lightened (a paint that is not a plain colour -- a
 * gradient, a pattern -- is the shape's own, unshaded).
 */
export function officeCustomPathSegmentFill(fill: string, mode: NativePathFill): string {
  if (mode === "none") return "transparent";
  return nativePathFillColor(fill, mode);
}
