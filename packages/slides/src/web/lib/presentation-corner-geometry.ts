/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  readNativeDrawingGeometry,
  type NativeDrawingGeometry,
} from "@mrnote/core/web/lib/native-drawing-geometry.ts";

const KINDS = new Set([
  "rounded-rectangle",
  "single-round-corner",
  "same-side-corners-rounded",
  "diagonal-corners-rounded",
  "single-corner-snipped",
  "same-side-corners-snipped",
  "diagonal-corners-snipped",
  "single-corner-rounded-snipped",
]);
const number = (value: number) => String(Number(value.toFixed(7)));

/** The eight DrawingML corner programs share short-side radii but different
 * corner positions. Formula/command order matches the existing POI definitions. */
function pathFor(
  kind: string,
  w: number,
  h: number,
  a: number,
  b: number,
): string {
  const commands: string[] = [];
  const point = (op: string, x: number, y: number) =>
    commands.push(`${op}${number(x)} ${number(y)}`);
  const arc = (radius: number, x: number, y: number) => {
    if (radius > 0)
      commands.push(
        `A${number(radius)} ${number(radius)} 0 0 1 ${number(x)} ${number(y)}`,
      );
  };
  switch (kind) {
    case "rounded-rectangle":
      point("M", 0, a);
      arc(a, a, 0);
      point("L", w - a, 0);
      arc(a, w, a);
      point("L", w, h - a);
      arc(a, w - a, h);
      point("L", a, h);
      arc(a, 0, h - a);
      break;
    case "single-round-corner":
      point("M", 0, 0);
      point("L", w - a, 0);
      arc(a, w, a);
      point("L", w, h);
      point("L", 0, h);
      break;
    case "same-side-corners-rounded":
      point("M", a, 0);
      point("L", w - a, 0);
      arc(a, w, a);
      point("L", w, h - b);
      arc(b, w - b, h);
      point("L", b, h);
      arc(b, 0, h - b);
      point("L", 0, a);
      arc(a, a, 0);
      break;
    case "diagonal-corners-rounded":
      point("M", a, 0);
      point("L", w - b, 0);
      arc(b, w, b);
      point("L", w, h - a);
      arc(a, w - a, h);
      point("L", b, h);
      arc(b, 0, h - b);
      point("L", 0, a);
      arc(a, a, 0);
      break;
    case "single-corner-snipped":
      point("M", 0, 0);
      point("L", w - a, 0);
      point("L", w, a);
      point("L", w, h);
      point("L", 0, h);
      break;
    case "same-side-corners-snipped":
      point("M", a, 0);
      point("L", w - a, 0);
      point("L", w, a);
      point("L", w, h - b);
      point("L", w - b, h);
      point("L", b, h);
      point("L", 0, h - b);
      point("L", 0, a);
      break;
    case "diagonal-corners-snipped":
      point("M", b, 0);
      point("L", w - a, 0);
      point("L", w, a);
      point("L", w, h - b);
      point("L", w - b, h);
      point("L", a, h);
      point("L", 0, h - a);
      point("L", 0, b);
      break;
    case "single-corner-rounded-snipped":
      point("M", a, 0);
      point("L", w - b, 0);
      point("L", w, b);
      point("L", w, h);
      point("L", 0, h);
      point("L", 0, a);
      arc(a, a, 0);
      break;
  }
  return [...commands, "Z"].join(" ");
}

const tokens = (path: string) =>
  path.match(/[MLAZ]|[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[-+]?\d+)?/g) ?? [];

/** Recover only native corner programs, including authored adjustments. The
 * reconstructed original must match before its radii can be used for resizing;
 * arbitrary/custom paths retain their existing native rendering. */
function sourceRadii(
  kind: string,
  geometry: NativeDrawingGeometry,
): [number, number] | null {
  const [w, h] = geometry.viewBox ?? [100, 100];
  if (
    geometry.paths.length !== 1 ||
    geometry.paths[0].fill !== "norm" ||
    !geometry.paths[0].stroke
  )
    return null;
  const source = tokens(geometry.paths[0].d);
  if (source.join("") !== geometry.paths[0].d.replace(/[\s,]/g, ""))
    return null;
  let first: [number, number] | undefined;
  const lines: Array<[number, number]> = [];
  for (let i = 0; i < source.length;) {
    const op = source[i++];
    if (op === "M" || op === "L") {
      const p: [number, number] = [Number(source[i++]), Number(source[i++])];
      if (op === "M") {
        if (first) return null;
        first = p;
      } else lines.push(p);
    } else if (op === "A") i += 7;
    else if (op !== "Z") return null;
  }
  if (!first || !lines.length) return null;
  let a = 0,
    b = 0;
  if (kind === "rounded-rectangle") a = first[1];
  else if (kind === "single-round-corner" || kind === "single-corner-snipped")
    a = w - lines[0][0];
  else if (kind === "same-side-corners-rounded") {
    a = first[0];
    b = h - (lines[1]?.[1] ?? NaN);
  } else if (kind === "same-side-corners-snipped") {
    a = first[0];
    b = h - (lines[2]?.[1] ?? NaN);
  } else if (kind === "diagonal-corners-snipped") {
    a = w - lines[0][0];
    b = first[0];
  } else {
    a = first[0];
    b = w - lines[0][0];
  }
  const short = Math.min(w, h);
  if (
    [a, b].some(
      (value) =>
        !Number.isFinite(value) || value < 0 || value > short / 2 + 0.00001,
    )
  )
    return null;
  const expected = tokens(pathFor(kind, w, h, a, b));
  if (
    source.length !== expected.length ||
    source.some((value, i) =>
      Number.isNaN(Number(value)) || Number.isNaN(Number(expected[i]))
        ? value !== expected[i]
        : Math.abs(Number(value) - Number(expected[i])) > 0.00001,
    )
  )
    return null;
  return [a / short, b / short];
}

export function presentationCornerGeometry(
  kind: string,
  width: number,
  height: number,
  source?: unknown,
): NativeDrawingGeometry | null {
  if (
    !KINDS.has(kind) ||
    [width, height].some(
      (value) => !Number.isFinite(value) || value <= 0 || value > 100_000,
    )
  )
    return null;
  let radii: [number, number] = [
    0.16667,
    kind === "single-corner-rounded-snipped" ? 0.16667 : 0,
  ];
  if (source !== undefined) {
    const geometry = readNativeDrawingGeometry(source);
    const recovered = geometry && sourceRadii(kind, geometry);
    if (!recovered) return null;
    radii = recovered;
  }
  const short = Math.min(width, height);
  return {
    schemaVersion: "drawingml-paths-v1",
    viewBox: [width, height],
    paths: [
      {
        d: pathFor(kind, width, height, radii[0] * short, radii[1] * short),
        fill: "norm",
        stroke: true,
      },
    ],
  };
}
