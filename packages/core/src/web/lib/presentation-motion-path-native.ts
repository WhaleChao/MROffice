/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  MAX_MOTION_PATH_EMU,
  MAX_MOTION_PATH_WAYPOINTS,
  motionPathPresetOfWaypoints,
  polylinePointsAlongSmoothly,
  presetMotionPath,
  type MotionPath,
  type MotionWaypoint,
} from "./presentation-motion-path";
import { smoothShares } from "./presentation-smooth-timing";

type Point = readonly [number, number];

type Segment =
  | { readonly kind: "L"; readonly to: Point }
  | { readonly kind: "C"; readonly c1: Point; readonly c2: Point; readonly to: Point };

interface SubPath {
  readonly start: Point;
  readonly segments: Segment[];
  closed: boolean;
}

interface ParsedMotionPath {
  readonly subpaths: SubPath[];
  /** The path states its end (`E`). */
  readonly ended: boolean;
}

const COORDS: Readonly<Record<string, number>> = { M: 2, L: 2, C: 6 };

/** The path's subpaths in absolute shares of the slide, or null when it is no path. */
export function parseNativeMotionPath(path: string): ParsedMotionPath | null {
  const tokens = path.match(/[A-Za-z]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/g) ?? [];
  const subpaths: SubPath[] = [];
  let current: Point = [0, 0];
  let command: string | null = null;
  let ended = false;
  let at = 0;
  while (at < tokens.length) {
    const token = tokens[at];
    if (/^[A-Za-z]$/.test(token)) {
      at += 1;
      const upper = token.toUpperCase();
      if (upper === "E") {
        ended = true;
        break;
      }
      if (upper === "Z") {
        const open = subpaths.at(-1);
        if (!open) return null;
        open.closed = true;
        current = open.start;
        command = null;
        continue;
      }
      if (!(upper in COORDS)) return null;
      command = token;
      continue;
    }
    if (command === null) return null;
    const upper: string = command.toUpperCase();
    const relative: boolean = command !== upper;
    const count: number = COORDS[upper];
    if (at + count > tokens.length) return null;
    const values = tokens.slice(at, at + count).map(Number);
    if (values.some((value) => !Number.isFinite(value))) return null;
    at += count;
    const base = current;
    const point = (index: number): Point =>
      relative ? [base[0] + values[index], base[1] + values[index + 1]] : [values[index], values[index + 1]];
    if (upper === "M") {
      current = point(0);
      subpaths.push({ start: current, segments: [], closed: false });
      // Coordinates after a move are lines (SVG's rule, which PowerPoint's
      // paths follow).
      command = relative ? "l" : "L";
      continue;
    }
    let open = subpaths.at(-1);
    if (!open) return null;
    if (open.closed) {
      // A segment after a close starts a new subpath where the last began.
      open = { start: open.start, segments: [], closed: false };
      subpaths.push(open);
    }
    if (upper === "L") {
      current = point(0);
      open.segments.push({ kind: "L", to: current });
    } else {
      const segment: Segment = { kind: "C", c1: point(0), c2: point(2), to: point(4) };
      current = segment.to;
      open.segments.push(segment);
    }
  }
  return subpaths.length > 0 ? { subpaths, ended } : null;
}

/**
 * The points a path passes through, in shares of the slide from the
 * object's place: each command's end, a curve flattened into `curveSteps`
 * pieces, a close back to its subpath's start. Null when it is no path.
 */
export function nativeMotionPathPoints(path: string, curveSteps = 12): Array<[number, number]> | null {
  const parsed = parseNativeMotionPath(path);
  if (!parsed) return null;
  const points: Array<[number, number]> = [];
  for (const subpath of parsed.subpaths) {
    let current = subpath.start;
    points.push([current[0], current[1]]);
    for (const segment of subpath.segments) {
      if (segment.kind === "L") {
        points.push([segment.to[0], segment.to[1]]);
      } else {
        for (let step = 1; step <= curveSteps; step += 1) {
          const t = step / curveSteps;
          const u = 1 - t;
          points.push([
            u * u * u * current[0] + 3 * u * u * t * segment.c1[0] + 3 * u * t * t * segment.c2[0] + t * t * t * segment.to[0],
            u * u * u * current[1] + 3 * u * u * t * segment.c1[1] + 3 * u * t * t * segment.c2[1] + t * t * t * segment.to[1],
          ]);
        }
      }
      current = segment.to;
    }
    if (subpath.closed && (current[0] !== subpath.start[0] || current[1] !== subpath.start[1])) {
      points.push([subpath.start[0], subpath.start[1]]);
    }
  }
  return points;
}

/** `count + 1` points evenly along a polyline by length (the object moves at one speed). */
export function pointsEvenlyAlong(points: ReadonlyArray<readonly [number, number]>, count: number): Array<[number, number]> {
  if (points.length === 0) return [[0, 0]];
  const [x0, y0] = points[0]!;
  // A path with no length is one place.
  if (points.every(([x, y]) => x === x0 && y === y0)) return [[x0, y0]];
  return polylinePointsAlongSmoothly(points, count, 0, 0);
}

/** Where a path starts and ends, in shares of the slide from the object's place. */
export function nativeMotionPathEnds(path: string): { start: [number, number]; end: [number, number] } | null {
  const points = nativeMotionPathPoints(path, 1);
  return points && points.length > 0 ? { start: points[0]!, end: points[points.length - 1]! } : null;
}

const percent = (value: number) => {
  const rounded = Number((value * 100).toFixed(4));
  return Object.is(rounded, -0) ? 0 : rounded;
};

/**
 * A path played as PowerPoint plays it, for the slide drawn as a size
 * container (`cqw` / `cqh` are shares of the slide): its `@keyframes`, under
 * a name of its own, from the path's own start to its end -- from its end to
 * its start when it is played backwards (`reversed`, Reverse Path Direction)
 * -- and the `translate` the object holds where it stops. The frames are at
 * even times with Smooth start / end (`accel`, `decel` as shares) baked in,
 * so the keyframes run `linear`: an easing between keyframes would slow the
 * object down at each of them and leave the path as a whole at an even pace.
 * Null when it is no path.
 */
export function nativeMotionPathKeyframes(
  path: string,
  reversed = false,
  accel = 0,
  decel = 0,
): { name: string; css: string; endTranslate: string } | null {
  const points = nativeMotionPathPoints(path);
  if (!points || points.length === 0) return null;
  const shares = smoothShares(accel, decel);
  const frames = polylinePointsAlongSmoothly(points, 40, shares.accel, shares.decel, reversed);
  const translate = ([x, y]: readonly [number, number]) => `${percent(x)}cqw ${percent(y)}cqh`;
  const body = frames
    .map((point, index) => `${Number(((index / (frames.length - 1 || 1)) * 100).toFixed(4))}% { translate: ${translate(point)}; }`)
    .join(" ");
  let h1 = 0x811c9dc5;
  let h2 = 0x811c9dc5 ^ 0x9e3779b9;
  for (let index = 0; index < path.length; index += 1) {
    const c = path.charCodeAt(index);
    h1 = Math.imul(h1 ^ c, 0x01000193);
    h2 = Math.imul(h2 ^ c, 0x85ebca77);
  }
  const smooth =
    shares.accel > 0 || shares.decel > 0
      ? `s${Math.round(shares.accel * 100_000)}-${Math.round(shares.decel * 100_000)}-`
      : "";
  const name = `mpn-${reversed ? "r-" : ""}${smooth}${(h1 >>> 0).toString(36)}-${(h2 >>> 0).toString(36)}`;
  return { name, css: `@keyframes ${name} { ${body} }`, endTranslate: translate(frames[frames.length - 1]!) };
}

/** A coordinate as the path writes it: at most six decimals, no exponent. */
function coordinate(value: number): string {
  const rounded = Number(value.toFixed(6));
  return Object.is(rounded, -0) || rounded === 0 ? "0" : String(rounded);
}

function pointText(point: Point): string {
  return `${coordinate(point[0])} ${coordinate(point[1])}`;
}

/** The path text of parsed subpaths (absolute commands, as PowerPoint writes them). */
export function formatNativeMotionPath(parsed: ParsedMotionPath): string {
  const parts: string[] = [];
  for (const subpath of parsed.subpaths) {
    parts.push(`M ${pointText(subpath.start)}`);
    for (const segment of subpath.segments) {
      parts.push(
        segment.kind === "L"
          ? `L ${pointText(segment.to)}`
          : `C ${pointText(segment.c1)} ${pointText(segment.c2)} ${pointText(segment.to)}`,
      );
    }
    if (subpath.closed) parts.push("Z");
  }
  if (parsed.ended) parts.push("E");
  return parts.join(" ");
}


/** The path moved by (dx, dy) shares of the slide, or null when it is no path. */
export function translateNativeMotionPath(path: string, dx: number, dy: number): string | null {
  const parsed = parseNativeMotionPath(path);
  if (!parsed || !Number.isFinite(dx) || !Number.isFinite(dy)) return null;
  const move = (point: Point): Point => [point[0] + dx, point[1] + dy];
  return formatNativeMotionPath({
    ended: parsed.ended,
    subpaths: parsed.subpaths.map((subpath) => ({
      start: move(subpath.start),
      closed: subpath.closed,
      segments: subpath.segments.map((segment) =>
        segment.kind === "L"
          ? { kind: "L", to: move(segment.to) }
          : { kind: "C", c1: move(segment.c1), c2: move(segment.c2), to: move(segment.to) },
      ),
    })),
  });
}

/** The import's tolerance for reading a path as one of the editor's presets. */
const PRESET_MATCH_TOLERANCE_EMU = 1_000;

const clampEmu = (value: number) => Math.max(-MAX_MOTION_PATH_EMU, Math.min(MAX_MOTION_PATH_EMU, value));

/**
 * The editor's polyline of a PowerPoint path on a slide of this size, as the
 * import reads it: every point a command ends on (a curve by its end), from
 * the path's first point, a close back to it; one of the editor's presets
 * when it is that preset's geometry. Null when nothing is left to play.
 */
export function motionPathOfNativePath(
  path: string,
  slideWidthEmu: number,
  slideHeightEmu: number,
): MotionPath | null {
  const parsed = parseNativeMotionPath(path);
  if (!parsed || !(slideWidthEmu > 0) || !(slideHeightEmu > 0)) return null;
  const origin = parsed.subpaths[0].start;
  const points: Array<Point | null> = [];
  parsed.subpaths.forEach((subpath, index) => {
    if (index > 0) points.push(subpath.start);
    for (const segment of subpath.segments) points.push(segment.to);
    if (subpath.closed) points.push(null);
  });
  const waypoints: MotionWaypoint[] = [];
  for (const point of points) {
    if (waypoints.length >= MAX_MOTION_PATH_WAYPOINTS) break;
    waypoints.push(
      point === null
        ? { x: 0, y: 0 }
        : {
            x: clampEmu(Math.round((point[0] - origin[0]) * slideWidthEmu)),
            y: clampEmu(Math.round((point[1] - origin[1]) * slideHeightEmu)),
          },
    );
  }
  if (waypoints.length === 0) return null;
  const preset = motionPathPresetOfWaypoints(waypoints, PRESET_MATCH_TOLERANCE_EMU);
  // A preset keeps its canonical geometry, as the import stores it.
  return (preset ? presetMotionPath(preset) : null) ?? { preset: "custom", waypoints };
}
