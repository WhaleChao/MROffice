/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure geometry for the FREEFORM / Scribble shape tool (PowerPoint "Insert Shapes
 * ▸ Freeform"): turn a sequence of clicked canvas-space (slide-EMU) vertices into
 * a custom-outline shape — its bounding rectangle plus a 0–100 normalized
 * `points` string, exactly the shape model `resolveShapeStyle` honours for a
 * merged / custom shape. Pure + dependency-light so the normalization is
 * unit-testable apart from the canvas pointer plumbing.
 */

import { formatPolygonString, type Point } from "./presentation-polygon-points";

export const PRESENTATION_CUSTOM_GEOMETRY_TOOLS = [
  "curve",
  "freeform",
  "scribble",
] as const;

export type PresentationCustomGeometryTool =
  (typeof PRESENTATION_CUSTOM_GEOMETRY_TOOLS)[number];

export const PRESENTATION_CUSTOM_PATH_GRID = 100_000;
// Bezier handles may overhang the path viewport. Keep endpoints on the grid
// and bound handles separately instead of flattening imported native curves.
export const PRESENTATION_CUSTOM_PATH_CONTROL_LIMIT =
  4 * PRESENTATION_CUSTOM_PATH_GRID;
// The commands one stored outline may hold; the server's twin is
// `CUSTOM_PATH_STORED_MAX_COMMANDS` in `app/services/office_custom_geometry.py`.
// 4096: a slide's lettering converted to one outline (a white panel with the


export const PRESENTATION_CUSTOM_PATH_MAX_COMMANDS = 4096;
// What a Word document's stored outline may hold: twice a slide's (a Word
// crest or map converted to one outline, corpus R156-D28: 4,492 commands).
// The server's twin is `DOCUMENT_CUSTOM_PATH_STORED_MAX_COMMANDS`; slides
// and sheets keep theirs.
export const DOCUMENT_CUSTOM_PATH_MAX_COMMANDS = 8192;
// What a path DRAWN here may hold (Freeform, Scribble), and the reading bound
// the server's `CUSTOM_PATH_MAX_COMMANDS` keeps: raising what an imported
// outline may hold did not ask for longer authored paths.
export const PRESENTATION_CUSTOM_PATH_AUTHORING_MAX_COMMANDS = 1024;
// What the Scribble tool samples one stroke down to. Authoring density, not
// the envelope: raising what an imported outline may hold did not ask for
// denser pen strokes.
export const PRESENTATION_SCRIBBLE_MAX_POINTS = 512;
// The envelope of a slide ink stroke's stored `points`: what the canvas draws,
// the erasers read, the page schema and the server (`_validate_presentation_block`)
// admit. An imported open path keeps one point per command beside its
// `customPath`, so the polyline may hold as many points as the outline holds
// commands; every pen stroke drawn here keeps at most the authoring density.
export const PRESENTATION_INK_MAX_POINTS = PRESENTATION_CUSTOM_PATH_AUTHORING_MAX_COMMANDS;
// The `points` string's length cap, for every shape that carries one: the
// envelope's points at the widest spelling a point string may take
// (`-999.99,-999.99` and a space: 16 characters). The server's twin is
// `MAX_PRESENTATION_SHAPE_POINTS_LENGTH`.
export const PRESENTATION_SHAPE_POINTS_MAX_LENGTH = 16 * PRESENTATION_INK_MAX_POINTS;
export const PRESENTATION_CUSTOM_PATH_MAX_EXTERNAL_DIMENSION = 2_147_483_647;
// `a:arcTo` states its angles in 60000ths of a degree (`ST_Angle`), so one
// full turn is 21_600_000 and writers do NOT reduce a start angle to one turn:
// the largest `@stAng` measured across 550 real packages is 21_838_329. Four
// turns each way is the same kind of bound the Bezier handle gets.
export const PRESENTATION_CUSTOM_PATH_MAX_ANGLE = 4 * 21_600_000;

export interface PresentationCustomPathPoint {
  readonly x: number;
  readonly y: number;
}

export type PresentationCustomPathCommand =
  | ({ readonly type: "moveTo" } & PresentationCustomPathPoint)
  | ({ readonly type: "lineTo" } & PresentationCustomPathPoint)
  | ({
      readonly type: "quadBezierTo";
      readonly x1: number;
      readonly y1: number;
    } & PresentationCustomPathPoint)
  | ({
      readonly type: "cubicBezierTo";
      readonly x1: number;
      readonly y1: number;
      readonly x2: number;
      readonly y2: number;
    } & PresentationCustomPathPoint)
  // An elliptical arc, kept as the arc Office wrote rather than flattened into
  // Beziers, so an untouched shape exports the four numbers it arrived with.
  // `wR`/`hR` are radii on the shared grid; the angles are dimensionless.
  | {
      readonly type: "arcTo";
      readonly wR: number;
      readonly hR: number;
      readonly stAng: number;
      readonly swAng: number;
    }
  // Ends one closed figure inside a path that holds several. A path whose only
  // `a:close` is its last child still says so with `closed: true` — the
  // spelling every page stored before this command existed uses.
  | { readonly type: "close" };

/**
 * Durable custom geometry. This intentionally records geometry, not the
 * transient Curve / Freeform / Scribble tool that authored it: DrawingML
 * `custGeom` has no trustworthy source-tool identity.
 */
export interface PresentationCustomPath {
  readonly version: 1;
  readonly closed: boolean;
  readonly commands: readonly PresentationCustomPathCommand[];
}

export interface PresentationCustomShapeResult {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
  readonly customPath: PresentationCustomPath;
}

export type PresentationFreeformDraftSegment =
  | { readonly type: "lineTo"; readonly point: Point }
  | { readonly type: "curveTo"; readonly samples: readonly Point[] };

function isControlCoordinate(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isInteger(value) &&
    Math.abs(value) <= PRESENTATION_CUSTOM_PATH_CONTROL_LIMIT
  );
}

/** A command that states a point of its own. `close` and `arcTo` do not. */
type PresentationCustomPathPointCommand = Extract<
  PresentationCustomPathCommand,
  PresentationCustomPathPoint
>;

function endpoint(command: PresentationCustomPathPointCommand) {
  return { x: command.x, y: command.y };
}

function geometryPoints(
  command: PresentationCustomPathCommand,
): PresentationCustomPathPoint[] {
  if (command.type === "quadBezierTo") {
    return [{ x: command.x1, y: command.y1 }, endpoint(command)];
  }
  if (command.type === "cubicBezierTo") {
    return [
      { x: command.x1, y: command.y1 },
      { x: command.x2, y: command.y2 },
      endpoint(command),
    ];
  }
  if (command.type === "close" || command.type === "arcTo") return [];
  return [endpoint(command)];
}

function hasNonCollinearTriple(
  points: readonly PresentationCustomPathPoint[],
): boolean {
  const first = points[0];
  const second = points.find(
    (point) => point.x !== first.x || point.y !== first.y,
  );
  if (!second) return false;
  return points.some(
    (point) =>
      (second.x - first.x) * (point.y - first.y) -
        (second.y - first.y) * (point.x - first.x) !==
      0,
  );
}

/** Strict persisted-model validator; malformed paths fail closed. */
export function normalizePresentationCustomPath(
  value: unknown,
  maxCommands: number = PRESENTATION_CUSTOM_PATH_MAX_COMMANDS,
): PresentationCustomPath | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  if (
    Object.keys(record).length !== 3 ||
    !Object.hasOwn(record, "version") ||
    !Object.hasOwn(record, "closed") ||
    !Object.hasOwn(record, "commands")
  ) {
    return null;
  }
  if (record.version !== 1 || typeof record.closed !== "boolean") return null;
  if (!Array.isArray(record.commands)) return null;
  if (
    record.commands.length < 2 ||
    record.commands.length > maxCommands
  ) {
    return null;
  }

  const commands: PresentationCustomPathCommand[] = [];
  let hasArc = false;
  for (let index = 0; index < record.commands.length; index += 1) {
    const raw = record.commands[index];
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
    const command = raw as Record<string, unknown>;
    const type = command.type;
    if (
      type !== "moveTo" &&
      type !== "lineTo" &&
      type !== "quadBezierTo" &&
      type !== "cubicBezierTo" &&
      type !== "arcTo" &&
      type !== "close"
    ) {
      return null;
    }
    // A `moveTo` ALWAYS begins a figure — after a `close` or simply by lifting
    // the pen, exactly as in SVG. A `close` never leads, never doubles, never
    // follows a bare `moveTo` (no figure to close) and never trails: the
    // trailing one is the envelope's own `closed`.
    const previous = commands[commands.length - 1]?.type;
    if (index === 0 && type !== "moveTo") return null;
    if (type === "close") {
      if (
        previous === undefined ||
        previous === "close" ||
        previous === "moveTo" ||
        index === record.commands.length - 1
      ) {
        return null;
      }
    } else if (type !== "moveTo" && previous === "close") {
      return null;
    }
    const expectedKeys =
      type === "quadBezierTo"
        ? ["type", "x1", "y1", "x", "y"]
        : type === "cubicBezierTo"
          ? ["type", "x1", "y1", "x2", "y2", "x", "y"]
          : type === "arcTo"
            ? ["type", "wR", "hR", "stAng", "swAng"]
            : type === "close"
              ? ["type"]
              : ["type", "x", "y"];
    const actualKeys = Object.keys(command).sort();
    // Sorted once per command, not once per key compared.
    const sortedExpectedKeys = [...expectedKeys].sort();
    if (
      actualKeys.length !== expectedKeys.length ||
      actualKeys.some((key, keyIndex) => key !== sortedExpectedKeys[keyIndex])
    ) {
      return null;
    }
    if (type === "close") {
      commands.push({ type });
      continue;
    }
    if (type === "arcTo") {
      const radii = [command.wR, command.hR];
      const angles = [command.stAng, command.swAng];
      if (
        radii.some(
          (value) =>
            typeof value !== "number" ||
            !Number.isInteger(value) ||
            value <= 0 ||
            value > PRESENTATION_CUSTOM_PATH_CONTROL_LIMIT,
        ) ||
        angles.some(
          (value) =>
            typeof value !== "number" ||
            !Number.isInteger(value) ||
            Math.abs(value) > PRESENTATION_CUSTOM_PATH_MAX_ANGLE,
        ) ||
        command.swAng === 0
      ) {
        return null;
      }
      hasArc = true;
      commands.push({
        type,
        wR: command.wR as number,
        hR: command.hR as number,
        stAng: command.stAng as number,
        swAng: command.swAng as number,
      });
      continue;
    }
    // A point on the outline may lie past the box as far as a handle may:
    // an imported outline is drawn where its points lie (Office's
    
    // `office_custom_geometry.normalize_custom_path`.
    if (!isControlCoordinate(command.x) || !isControlCoordinate(command.y)) {
      return null;
    }
    if (type === "quadBezierTo") {
      if (
        !isControlCoordinate(command.x1) ||
        !isControlCoordinate(command.y1)
      ) {
        return null;
      }
      commands.push({
        type,
        x1: command.x1,
        y1: command.y1,
        x: command.x,
        y: command.y,
      });
    } else if (type === "cubicBezierTo") {
      if (
        !isControlCoordinate(command.x1) ||
        !isControlCoordinate(command.y1) ||
        !isControlCoordinate(command.x2) ||
        !isControlCoordinate(command.y2)
      ) {
        return null;
      }
      commands.push({
        type,
        x1: command.x1,
        y1: command.y1,
        x2: command.x2,
        y2: command.y2,
        x: command.x,
        y: command.y,
      });
    } else {
      commands.push({ type, x: command.x, y: command.y });
    }
  }

  if (hasArc) {
    // An arc with two positive radii and a non-zero sweep bounds area by
    // construction, and where it LANDS needs trigonometry. Deliberately not
    // done here: this validator is duplicated in Python and both halves have
    // to admit exactly the same envelope, which integer comparisons guarantee
    // and two implementations of sin/cos do not. The canvas does that
    // arithmetic, where a last-bit difference is invisible.
    return Object.freeze({
      version: 1,
      closed: record.closed,
      commands: Object.freeze(commands),
    });
  }
  // No arc reaches this point (an arc returned above); `close` has no anchor.
  const anchors = commands
    .filter(
      (command): command is PresentationCustomPathPointCommand =>
        command.type !== "close" && command.type !== "arcTo",
    )
    .map(endpoint);
  const uniqueAnchors = new Set(anchors.map(({ x, y }) => `${x},${y}`));
  if (!record.closed) {
    if (uniqueAnchors.size < 2) return null;
  } else {
    const geometry = commands.flatMap(geometryPoints);
    const uniqueGeometry = Array.from(
      new Map(
        geometry.map((point) => [`${point.x},${point.y}`, point]),
      ).values(),
    );
    if (uniqueGeometry.length < 3) return null;
    const xs = uniqueGeometry.map(({ x }) => x);
    const ys = uniqueGeometry.map(({ y }) => y);
    if (
      Math.max(...xs) <= Math.min(...xs) ||
      Math.max(...ys) <= Math.min(...ys)
    ) {
      return null;
    }
    // A closed Bezier may enclose a real region even when endpoints repeat; its
    // controls are therefore part of the dimensionality proof.
    if (!hasNonCollinearTriple(uniqueGeometry)) return null;
  }

  return Object.freeze({
    version: 1,
    closed: record.closed,
    commands: Object.freeze(commands),
  });
}

/** Whether the outline curves anywhere — an arc counts, it is not a line. */
export function presentationCustomPathHasBezier(
  path: PresentationCustomPath,
): boolean {
  return path.commands.some(
    (command) =>
      command.type === "quadBezierTo" ||
      command.type === "cubicBezierTo" ||
      command.type === "arcTo",
  );
}

/** Whether this one `a:path` holds more than one figure. */
export function presentationCustomPathHasSubPaths(
  path: PresentationCustomPath,
): boolean {
  return path.commands.some((command) => command.type === "close");
}

export function presentationCustomPathIsOpenLineOnly(
  path: PresentationCustomPath,
): boolean {
  return (
    !path.closed &&
    path.commands[0]?.type === "moveTo" &&
    path.commands.slice(1).every((command) => command.type === "lineTo")
  );
}

/** Whether the outline starts a second figure (a `moveTo` after the first command). */
export function presentationCustomPathLiftsThePen(
  path: PresentationCustomPath,
): boolean {
  return path.commands.slice(1).some((command) => command.type === "moveTo");
}









export function presentationCustomGeometryPathAdmitted(
  path: PresentationCustomPath,
): boolean {
  return (
    path.closed ||
    presentationCustomPathHasBezier(path) ||
    presentationCustomPathLiftsThePen(path)
  );
}

function formatGrid(value: number): string {
  return Number((value / 1_000).toFixed(3)).toString();
}

/** One 60000ths-of-a-degree `ST_Angle` as radians. */
const ANGLE_UNITS_PER_TURN = 21_600_000;

/**
 * Where an `a:arcTo` lands, given where the pen already is.
 *
 * ECMA-376 §20.1.9.4: the arc is the piece of an axis-aligned ellipse with
 * radii `wR`/`hR` that starts at the current point AT `stAng` and sweeps
 * `swAng` further. So the centre is fixed by the start — it is the current
 * point pushed back along the start angle — and the end follows from the
 * centre. Angles grow clockwise because y grows downward, which is also what
 * SVG's sweep flag means, so the two agree without a sign flip.
 *
 * This is the only floating-point arithmetic in the custom-path model, and it
 * lives here rather than in the validator on purpose: a last-bit difference
 * between the two language halves must never decide whether a shape is
 * admitted, only where a curve is drawn.
 */
export function presentationCustomPathArcEnd(
  from: PresentationCustomPathPoint,
  arc: { readonly wR: number; readonly hR: number; readonly stAng: number; readonly swAng: number },
): PresentationCustomPathPoint {
  const start = (arc.stAng / ANGLE_UNITS_PER_TURN) * 2 * Math.PI;
  const end = ((arc.stAng + arc.swAng) / ANGLE_UNITS_PER_TURN) * 2 * Math.PI;
  const centreX = from.x - arc.wR * Math.cos(start);
  const centreY = from.y - arc.hR * Math.sin(start);
  return {
    x: centreX + arc.wR * Math.cos(end),
    y: centreY + arc.hR * Math.sin(end),
  };
}

export function presentationCustomPathToSvgPath(value: unknown): string | null {
  const path = normalizePresentationCustomPath(value);
  if (!path) return null;
  return presentationCustomPathCommandsToSvgPath(path.commands, path.closed);
}

/**
 * The SVG path data of commands an already-normalized envelope holds -- the
 * whole of one, or the commands of ONE of its painted paths
 * (`lib/office-custom-path-paints`), which need not be an envelope alone.
 */
export function presentationCustomPathCommandsToSvgPath(
  commands: readonly PresentationCustomPathCommand[],
  closed: boolean,
): string {
  // An arc starts where the pen is, and a `close` puts the pen back at the
  // figure's first point, so drawing needs the position the commands imply.
  let pen: PresentationCustomPathPoint = { x: 0, y: 0 };
  let figureStart: PresentationCustomPathPoint = pen;
  const d = commands
    .map((command) => {
      if (command.type === "close") {
        pen = figureStart;
        return "Z";
      }
      if (command.type === "arcTo") {
        const to = presentationCustomPathArcEnd(pen, command);
        // SVG needs the large-arc flag the sweep implies, and its sweep flag
        // is 1 for the same direction DrawingML calls a positive angle.
        const large = Math.abs(command.swAng) > ANGLE_UNITS_PER_TURN / 2 ? 1 : 0;
        const sweep = command.swAng > 0 ? 1 : 0;
        pen = to;
        return (
          `A${formatGrid(command.wR)} ${formatGrid(command.hR)} 0 ${large} ${sweep}` +
          ` ${formatGrid(to.x)} ${formatGrid(to.y)}`
        );
      }
      pen = { x: command.x, y: command.y };
      if (command.type === "moveTo") {
        figureStart = pen;
        return `M${formatGrid(command.x)} ${formatGrid(command.y)}`;
      }
      if (command.type === "lineTo") {
        return `L${formatGrid(command.x)} ${formatGrid(command.y)}`;
      }
      if (command.type === "quadBezierTo") {
        return `Q${formatGrid(command.x1)} ${formatGrid(command.y1)} ${formatGrid(command.x)} ${formatGrid(command.y)}`;
      }
      return `C${formatGrid(command.x1)} ${formatGrid(command.y1)} ${formatGrid(command.x2)} ${formatGrid(command.y2)} ${formatGrid(command.x)} ${formatGrid(command.y)}`;
    })
    .join(" ");
  return closed ? `${d} Z` : d;
}

export function presentationCustomPathToLegacyPoints(
  value: unknown,
): string | null {
  const path = normalizePresentationCustomPath(value);
  if (!path || !presentationCustomPathIsOpenLineOnly(path)) return null;
  // `isOpenLineOnly` already proved every command is a `moveTo`/`lineTo`; the
  // filter is what tells the compiler the same thing.
  return path.commands
    .filter(
      (command): command is PresentationCustomPathPointCommand =>
        command.type !== "close" && command.type !== "arcTo",
    )
    .map((command) => `${formatGrid(command.x)},${formatGrid(command.y)}`)
    .join(" ");
}

/** DrawingML coordinate normalization, using deterministic exact half-up. */
export function normalizeExternalCustomPathCoordinate(
  coordinate: unknown,
  dimension: unknown,
  control = false,
): number | null {
  if (
    typeof coordinate !== "number" ||
    !Number.isSafeInteger(coordinate) ||
    typeof dimension !== "number" ||
    !Number.isInteger(dimension) ||
    dimension < 0 ||
    dimension > PRESENTATION_CUSTOM_PATH_MAX_EXTERNAL_DIMENSION
  ) {
    return null;
  }
  // A FLAT path declares that extent as 0: `<a:path w="523295" h="0">` is a
  // perfectly horizontal stroke, which is exactly what Office writes for the
  // straight connectors inside a diagram. Every coordinate on that axis is
  // then 0 and stays 0; there is nothing to divide by and nothing to scale.
  // The server twin (`office_custom_geometry.custom_path_coordinate`) has read
  // it that way since the sheet host was connected; this side refused it, so
  // the same package projected differently in the two chains
  // (SHAPE-IMPORT-AUDIT.md §8.2).
  //
  // Office rounds every point to its path's INTEGER grid, and that rounding
  // can put a point ONE unit outside the box it closes (`y="2538830"` in
  // `h="2538829"`, `y="1"` on a flat line). One unit, and only one, is read as
  // the edge it rounded past; refusing it took real freeforms off the slide
  
  // twin (`office_custom_geometry.custom_path_coordinate`).
  // A Bezier handle is bounded by the envelope's own +/-4x, not by the box,
  // so it takes no slack.
  if (dimension === 0) return Math.abs(coordinate) <= 1 ? 0 : null;
  const lower = control ? -dimension * 4 : 0;
  const upper = control ? dimension * 4 : dimension;
  // One unit of the path's grid, or of the envelope's when that is coarser.
  const slack = control
    ? 0
    : Math.max(1, Math.floor(dimension / PRESENTATION_CUSTOM_PATH_GRID));
  if (coordinate < lower - slack || coordinate > upper + slack) return null;
  const clamped = Math.min(Math.max(coordinate, lower), upper);
  const scaled = BigInt(clamped) * BigInt(PRESENTATION_CUSTOM_PATH_GRID);
  const divisor = BigInt(dimension);
  let quotient = scaled / divisor;
  let remainder = scaled % divisor;
  if (remainder < 0n) {
    quotient -= 1n;
    remainder += divisor;
  }
  return Number(quotient + (remainder * 2n >= divisor ? 1n : 0n));
}

type SlidePathCommand =
  | ({ readonly type: "moveTo" | "lineTo" } & Point)
  | ({
      readonly type: "cubicBezierTo";
      readonly c1: Point;
      readonly c2: Point;
    } & Point);

function resultFromSlideCommands(
  rawCommands: readonly SlidePathCommand[],
  closed: boolean,
  padEmu = 45_720,
): PresentationCustomShapeResult | null {
  if (
    rawCommands.length < 2 ||
    rawCommands.length > PRESENTATION_CUSTOM_PATH_AUTHORING_MAX_COMMANDS
  ) {
    return null;
  }
  const allPoints = rawCommands.flatMap((command) =>
    command.type === "cubicBezierTo"
      ? [command.c1, command.c2, { x: command.x, y: command.y }]
      : [{ x: command.x, y: command.y }],
  );
  if (
    allPoints.some(({ x, y }) => !Number.isFinite(x) || !Number.isFinite(y))
  ) {
    return null;
  }
  const xs = allPoints.map(({ x }) => x);
  const ys = allPoints.map(({ y }) => y);
  let minX = Math.min(...xs);
  let maxX = Math.max(...xs);
  let minY = Math.min(...ys);
  let maxY = Math.max(...ys);
  if (!closed) {
    if (maxX <= minX) {
      minX -= padEmu;
      maxX += padEmu;
    }
    if (maxY <= minY) {
      minY -= padEmu;
      maxY += padEmu;
    }
  }
  const width = maxX - minX;
  const height = maxY - minY;
  if (width <= 0 || height <= 0) return null;
  const grid = (value: number, min: number, extent: number) =>
    Math.max(
      0,
      Math.min(
        PRESENTATION_CUSTOM_PATH_GRID,
        Math.round(((value - min) / extent) * PRESENTATION_CUSTOM_PATH_GRID),
      ),
    );
  const commands: PresentationCustomPathCommand[] = rawCommands.map(
    (command) => {
      if (command.type === "cubicBezierTo") {
        return {
          type: "cubicBezierTo",
          x1: grid(command.c1.x, minX, width),
          y1: grid(command.c1.y, minY, height),
          x2: grid(command.c2.x, minX, width),
          y2: grid(command.c2.y, minY, height),
          x: grid(command.x, minX, width),
          y: grid(command.y, minY, height),
        };
      }
      return {
        type: command.type,
        x: grid(command.x, minX, width),
        y: grid(command.y, minY, height),
      };
    },
  );
  const customPath = normalizePresentationCustomPath({
    version: 1,
    closed,
    commands,
  });
  if (!customPath) return null;
  return { left: minX, top: minY, width, height, customPath };
}

function pointDistance(a: Point, b: Point): number {
  return Math.hypot(a.x - b.x, a.y - b.y);
}

export function isNearCustomPathStart(
  points: readonly Point[],
  thresholdEmu: number,
): boolean {
  return (
    points.length >= 3 &&
    thresholdEmu > 0 &&
    pointDistance(points[0], points[points.length - 1]) <= thresholdEmu
  );
}

export function buildCurveCustomGeometry(
  knots: readonly Point[],
  closed: boolean,
): PresentationCustomShapeResult | null {
  if (knots.length < (closed ? 3 : 2)) return null;
  const effective =
    closed && pointDistance(knots[0], knots[knots.length - 1]) === 0
      ? knots.slice(0, -1)
      : [...knots];
  if (effective.length < (closed ? 3 : 2)) return null;
  const commands: SlidePathCommand[] = [
    { type: "moveTo", x: effective[0].x, y: effective[0].y },
  ];
  const segmentCount = closed ? effective.length : effective.length - 1;
  for (let index = 0; index < segmentCount; index += 1) {
    const p1 = effective[index % effective.length];
    const p2 = effective[(index + 1) % effective.length];
    const p0 = closed
      ? effective[(index - 1 + effective.length) % effective.length]
      : effective[Math.max(0, index - 1)];
    const p3 = closed
      ? effective[(index + 2) % effective.length]
      : effective[Math.min(effective.length - 1, index + 2)];
    commands.push({
      type: "cubicBezierTo",
      c1: { x: p1.x + (p2.x - p0.x) / 6, y: p1.y + (p2.y - p0.y) / 6 },
      c2: { x: p2.x - (p3.x - p1.x) / 6, y: p2.y - (p3.y - p1.y) / 6 },
      x: p2.x,
      y: p2.y,
    });
  }
  return resultFromSlideCommands(commands, closed);
}

export function buildFreeformCustomGeometry(
  points: readonly Point[],
  closed: boolean,
  curved = false,
): PresentationCustomShapeResult | null {
  if (curved) return buildCurveCustomGeometry(points, closed);
  return resultFromSlideCommands(
    points.map((point, index) => ({
      type: index === 0 ? "moveTo" : "lineTo",
      x: point.x,
      y: point.y,
    })),
    closed,
  );
}

function appendCatmullRomCommands(
  commands: SlidePathCommand[],
  points: readonly Point[],
) {
  for (let index = 0; index < points.length - 1; index += 1) {
    const p1 = points[index];
    const p2 = points[index + 1];
    const p0 = points[Math.max(0, index - 1)];
    const p3 = points[Math.min(points.length - 1, index + 2)];
    commands.push({
      type: "cubicBezierTo",
      c1: { x: p1.x + (p2.x - p0.x) / 6, y: p1.y + (p2.y - p0.y) / 6 },
      c2: { x: p2.x - (p3.x - p1.x) / 6, y: p2.y - (p3.y - p1.y) / 6 },
      x: p2.x,
      y: p2.y,
    });
  }
}

/** Preserve ordered click-line and drag-curve segments in one Freeform path. */
export function buildMixedFreeformCustomGeometry(
  start: Point,
  segments: readonly PresentationFreeformDraftSegment[],
  closed: boolean,
): PresentationCustomShapeResult | null {
  const commands: SlidePathCommand[] = [
    { type: "moveTo", x: start.x, y: start.y },
  ];
  let current = start;
  for (const segment of segments) {
    if (segment.type === "lineTo") {
      commands.push({
        type: "lineTo",
        x: segment.point.x,
        y: segment.point.y,
      });
      current = segment.point;
      continue;
    }
    const samples = [
      current,
      ...simplifyScribblePoints(segment.samples, 2_000),
    ];
    if (samples.length < 2) continue;
    appendCatmullRomCommands(commands, samples);
    current = samples[samples.length - 1];
    if (commands.length > PRESENTATION_CUSTOM_PATH_AUTHORING_MAX_COMMANDS) return null;
  }
  return resultFromSlideCommands(commands, closed);
}

/**
 * At most `limit` of `items`, spread evenly along them: the first and the
 * last always, and between them the item nearest each even step (half up, in
 * integers, so the server's `evenly_thinned` picks the same ones). Fewer come
 * back as they are.
 */
export function evenlyThinned<T>(items: readonly T[], limit: number): T[] {
  const count = items.length;
  if (count <= limit) return items.slice();
  if (limit < 2) return items.slice(0, Math.max(0, limit));
  const span = limit - 1;
  return Array.from(
    { length: limit },
    (_, index) => items[Math.floor((2 * index * (count - 1) + span) / (2 * span))],
  );
}

/**
 * A stored ink stroke's `points` held to the envelope
 * ({@link PRESENTATION_INK_MAX_POINTS}). Only Record Studio ever stored more
 * (it kept every sample of a gesture), which the canvas could not draw as ink;
 * such a stroke is thinned to the authoring density, each point kept as it was
 * spelled. Anything else comes back unchanged.
 */
export function presentationInkPointsWithinEnvelope(value: unknown): unknown {
  if (typeof value !== "string") return value;
  const tokens = value.trim().split(/\s+/);
  if (tokens.length <= PRESENTATION_INK_MAX_POINTS) return value;
  return evenlyThinned(tokens, PRESENTATION_SCRIBBLE_MAX_POINTS).join(" ");
}

/** Deterministic distance simplifier used by continuous Scribble sampling. */
export function simplifyScribblePoints(
  points: readonly Point[],
  minimumDistance: number,
  limit = PRESENTATION_SCRIBBLE_MAX_POINTS,
): Point[] {
  const finite = points.filter(
    ({ x, y }) => Number.isFinite(x) && Number.isFinite(y),
  );
  if (finite.length <= 2) return finite.slice(0, limit);
  // A stroke longer than `limit` samples at `minimumDistance` keeps its
  // samples further apart, evenly along its whole length: stopping at the
  // limit cut every long stroke short and joined the cut to its end with a
  // straight line.
  let length = 0;
  for (let index = 1; index < finite.length; index += 1) {
    length += pointDistance(finite[index - 1], finite[index]);
  }
  const spacing = Math.max(minimumDistance, limit > 2 ? length / (limit - 2) : length);
  const result: Point[] = [finite[0]];
  for (let index = 1; index < finite.length - 1; index += 1) {
    if (
      pointDistance(result[result.length - 1], finite[index]) >= spacing
    ) {
      result.push(finite[index]);
      if (result.length === limit - 1) break;
    }
  }
  const last = finite[finite.length - 1];
  if (
    pointDistance(result[result.length - 1], last) > 0 &&
    result.length < limit
  ) {
    result.push(last);
  }
  return result;
}

export function buildScribbleCustomGeometry(
  samples: readonly Point[],
  minimumDistance = 9_144,
): PresentationCustomShapeResult | null {
  return buildFreeformCustomGeometry(
    simplifyScribblePoints(samples, minimumDistance),
    false,
  );
}

export interface FreeformShapeResult {
  left: number;
  top: number;
  width: number;
  height: number;
  /** The outline in the shape's own 0–100 space ("x,y x,y …"). */
  points: string;
}

/**
 * Build a custom-outline shape from freeform vertices in slide-EMU space. Returns
 * null for fewer than 3 vertices, a non-finite coordinate, or a degenerate
 * (zero-area — e.g. all-collinear) path, so a stray click or a straight line can
 * never produce an invisible / malformed shape.
 */
export function freeformToShape(
  points: readonly Point[],
): FreeformShapeResult | null {
  if (points.length < 3) return null;
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const p of points) {
    if (!Number.isFinite(p.x) || !Number.isFinite(p.y)) return null;
    if (p.x < minX) minX = p.x;
    if (p.y < minY) minY = p.y;
    if (p.x > maxX) maxX = p.x;
    if (p.y > maxY) maxY = p.y;
  }
  const width = maxX - minX;
  const height = maxY - minY;
  if (width <= 0 || height <= 0) return null;
  const normalized: Point[] = points.map((p) => ({
    x: ((p.x - minX) / width) * 100,
    y: ((p.y - minY) / height) * 100,
  }));
  return {
    left: minX,
    top: minY,
    width,
    height,
    points: formatPolygonString(normalized),
  };
}

/**
 * Build an INK STROKE shape from sampled pointer positions in slide-EMU space.
 * Unlike {@link freeformToShape} (closed polygons), a stroke may be straight —
 * zero width OR height is fine (a horizontal pen line is real ink). The box is
 * padded so a straight stroke still has a bounding box the geometry pipeline
 * (resize handles, marquee) can grab.
 *
 * A tap of the pen (every sample at one place) is a dot of the pen's own tip,
 * as a tap is in Office's ink: a box of the padding on each side of the tap,
 * and the shortest line the stored points state (one hundredth of a percent of
 * that box either side of its centre: the two decimals every point string
 * keeps, and less than any zoom can show), so the stroke's round tip paints
 * the dot exactly where the pen came down. Two equal samples would read back
 * as one point, which a file's outline cannot write. Only no sample at all,
 * or a non-finite one, is no stroke.
 *
 * A stroke keeps at most the authoring density
 * ({@link PRESENTATION_SCRIBBLE_MAX_POINTS}) of its samples, spread evenly
 * along it, whoever drew it: a Record Studio gesture keeps every sample the
 * pointer gave, and a stroke stored longer than the canvas draws fell back to
 * the kind's outline.
 */
export function inkStrokeToShape(
  samples: readonly Point[],
  padEmu = 45_720, // ~5px
): FreeformShapeResult | null {
  if (samples.length === 0) return null;
  if (samples.some((p) => !Number.isFinite(p.x) || !Number.isFinite(p.y))) return null;
  const points = evenlyThinned(samples, PRESENTATION_INK_MAX_POINTS);
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const p of points) {
    if (p.x < minX) minX = p.x;
    if (p.y < minY) minY = p.y;
    if (p.x > maxX) maxX = p.x;
    if (p.y > maxY) maxY = p.y;
  }
  if (maxX - minX <= 0 && maxY - minY <= 0) {
    return {
      left: minX - padEmu,
      top: minY - padEmu,
      width: Math.max(1, padEmu * 2),
      height: Math.max(1, padEmu * 2),
      points: "49.99,50 50.01,50",
    };
  }
  const left = minX - padEmu;
  const top = minY - padEmu;
  const width = Math.max(1, maxX - minX + padEmu * 2);
  const height = Math.max(1, maxY - minY + padEmu * 2);
  const normalized: Point[] = points.map((p) => ({
    x: ((p.x - left) / width) * 100,
    y: ((p.y - top) / height) * 100,
  }));
  // A thousandth of the box, as the curve is drawn from it (precision 3):
  // a hundredth was 0.13px across a stroke the width of the slide, and the
  // stroke moved by it as it landed.
  const n = (v: number) => String(Number(v.toFixed(3)));
  return {
    left,
    top,
    width,
    height,
    points: normalized.map((p) => `${n(p.x)},${n(p.y)}`).join(" "),
  };
}
