/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { readNativeDrawingGeometry, type NativeDrawingGeometry, type NativePathFill } from "./native-drawing-geometry";

// Independent implementation of ISO/IEC 29500 DrawingML guide/path semantics.
// Programs come from the generated POI asset, never from page metadata. There
// is deliberately no custom-formula override or authoring entry point here.
const ARITY: Record<string, number> = { val: 1, abs: 1, sqrt: 1, max: 2, min: 2, at2: 2, cos: 2, sin: 2, tan: 2, "*/": 3, "+-": 3, "+/": 3, "?:": 3, pin: 3, mod: 3, cat2: 3, sat2: 3 };
const COMMAND_ARITY: Record<string, number> = { moveTo: 2, lnTo: 2, quadBezTo: 4, cubicBezTo: 6, arcTo: 4, close: 0 };
const ANGLE_UNIT = Math.PI / 10_800_000;
const owns = (value: object, key: string) => Object.prototype.hasOwnProperty.call(value, key);
function fail(): never { throw new Error("invalid_default_preset_program"); }
function record(value: unknown, keys: string[]): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value) || Object.keys(value).length !== keys.length || keys.some(key => !owns(value, key))) return fail();
  return value as Record<string, unknown>;
}
function finite(value: number): number {
  if (!Number.isFinite(value) || Math.abs(value) > 1e100) return fail();
  return value;
}
function coordinate(value: number): string {
  if (Math.abs(value) > 1e10) return fail();
  return String(Number(finite(value).toFixed(7)));
}
function variables(width: number, height: number): Record<string, number> {
  if (!Number.isFinite(width) || !Number.isFinite(height) || width < 0 || height < 0 || width > 1e8 || height > 1e8 || width + height <= 0) return fail();
  const values = Object.assign(Object.create(null), { l: 0, t: 0, r: width, b: height, w: width, h: height,
    hc: width / 2, vc: height / 2, ss: Math.min(width, height), ls: Math.max(width, height),
    cd2: 10_800_000, cd4: 5_400_000, cd8: 2_700_000, "3cd4": 16_200_000, "3cd8": 8_100_000, "5cd8": 13_500_000, "7cd8": 18_900_000 });
  for (const divisor of [2, 3, 4, 5, 6, 8, 10, 12, 16, 32]) {
    for (const [prefix, value] of [["wd", width], ["hd", height], ["ssd", Math.min(width, height)]] as const) values[prefix + divisor] = value / divisor;
  }
  return values;
}
function atom(value: unknown, vars: Record<string, number>): number {
  if (typeof value === "number") return finite(value);
  const ref = record(value, ["ref"]).ref;
  if (typeof ref !== "string" || !/^[A-Za-z0-9_]{1,81}$/.test(ref) || !owns(vars, ref)) return fail();
  return vars[ref];
}
function args(value: unknown, count: number, vars: Record<string, number>): number[] {
  if (!Array.isArray(value) || value.length !== count) return fail();
  return value.map(item => atom(item, vars));
}
function guide(op: string, operands: number[]): number {
  const [x, y, z] = operands;
  let result: number;
  switch (op) {
    case "val": result = x; break;
    case "abs": result = Math.abs(x); break;
    case "sqrt": result = Math.sqrt(Math.max(0, x)); break;
    case "max": result = Math.max(x, y); break;
    case "min": result = Math.min(x, y); break;
    case "at2": result = Math.atan2(y, x) / ANGLE_UNIT; break;
    case "cos": result = x * Math.cos(y * ANGLE_UNIT); break;
    case "sin": result = x * Math.sin(y * ANGLE_UNIT); break;
    case "tan": result = x * Math.tan(y * ANGLE_UNIT); break;
    case "*/": result = z ? x * y / z : 0; break;
    case "+-": result = x + y - z; break;
    case "+/": result = z ? (x + y) / z : 0; break;
    case "?:": result = x > 0 ? y : z; break;
    case "pin": result = Math.min(Math.max(y, x), z); break;
    case "mod": result = Math.hypot(x, y, z); break;
    case "cat2": result = x * Math.cos(Math.atan2(z, y)); break;
    case "sat2": result = x * Math.sin(Math.atan2(z, y)); break;
    default: return fail();
  }
  return finite(result);
}







function flatArcPoints(rx: number, ry: number, start: number, sweep: number): [number, number][] {
  const point = (angle: number): [number, number] => {
    let cos = Math.cos(angle), sin = Math.sin(angle);
    if (Math.abs(cos) < 1e-9) cos = 0;
    if (Math.abs(sin) < 1e-9) sin = 0;
    return ry === 0 ? [sin === 0 ? (cos < 0 ? -rx : rx) : 0, 0] : [0, cos === 0 ? (sin < 0 ? -ry : ry) : 0];
  };
  const quarter = Math.PI / 2, step = sweep > 0 ? 1 : -1, angles = [start];
  let k = sweep > 0 ? Math.floor(start / quarter + 1e-9) + 1 : Math.ceil(start / quarter - 1e-9) - 1;
  while ((k * quarter - (start + sweep)) * step < -1e-9) { angles.push(k * quarter); k += step; }
  angles.push(start + sweep);
  return angles.map(point);
}
function ellipsePoint(rx: number, ry: number, angle: number): [number, number] {
  // DrawingML uses the polar angle of a ray through an ellipse, not its
  // parameter angle. The distinction is visible when the shape is stretched.
  const theta = Math.atan2(rx * Math.sin(angle), ry * Math.cos(angle));
  return [rx * Math.cos(theta), ry * Math.sin(theta)];
}

/**
 * The guides one program takes an ADJUST VALUE for: the `a:avLst` handles, which
 * the catalog writes as a literal `val`. A `a:gdLst` guide never is one (its
 * `val` forms all reference another guide), so this is the same set the server's
 * `preset_adjust_names` returns from the same catalog.
 */
export function drawingmlDefaultPresetAdjustNames(value: unknown): readonly string[] | null {
  try {
    const program = record(value, ["dialect", "preset", "guides", "paths"]);
    if (program.dialect !== "drawingml-default-program-v1" || !Array.isArray(program.guides)) return null;
    const names: string[] = [];
    for (const raw of program.guides) {
      const item = record(raw, ["name", "op", "args"]);
      if (item.op !== "val" || !Array.isArray(item.args) || typeof item.args[0] !== "number") break;
      if (typeof item.name !== "string" || !/^[A-Za-z0-9_]{1,81}$/.test(item.name)) return null;
      names.push(item.name);
    }
    return names;
  } catch { return null; }
}

/**
 * Whether an `a:custGeom`'s `a:rect` written through `a:gdLst` guides is a
 * spelling of the full box or of the empty one -- the two text rectangles the
 * custom-geometry readers recognise. PowerPoint 2003's converter writes
 * `l="T15" t="T16" r="T17" b="T18"` over `*\/ 0 w 2021204` guides. Evaluated
 * with this module's one guide evaluator at two unrelated box sizes. The server
 * twin is `office_custom_geometry._guide_named_rect_is_recognisable`.
 */
export function drawingmlGuideNamedRectIsBoxOrEmpty(
  guides: readonly (readonly [string, string])[],
  rect: Readonly<Record<"l" | "t" | "r" | "b", string>>,
): boolean {
  for (const [width, height] of [[914_400, 685_800], [3_117_000, 211_000]] as const) {
    const vars = variables(width, height);
    for (const [name, formula] of guides) {
      if (!name || owns(vars, name)) continue;
      const [op = "", ...operands] = formula.trim().split(/\s+/);
      if (!owns(ARITY, op) || operands.length !== ARITY[op]) continue;
      try {
        vars[name] = guide(op, operands.map((token) => owns(vars, token)
          ? vars[token]
          : /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(token) ? finite(Number(token)) : fail()));
      } catch {
        continue;
      }
    }
    const values: number[] = [];
    for (const edge of ["l", "t", "r", "b"] as const) {
      const token = rect[edge];
      if (owns(vars, token)) values.push(vars[token]);
      else if (/^-?[0-9]{1,10}$/.test(token)) values.push(Number(token));
      else return false;
    }
    const full = [0, 0, width, height];
    if (!values.every((value, index) => Math.abs(value - full[index]) < 1) &&
      !values.every((value) => Math.abs(value) < 1)) return false;
  }
  return true;
}

function guideVariables(guides: unknown[], widthEmu: number, heightEmu: number,
  adjustments?: Readonly<Record<string, number>> | null): { vars: Record<string, number>; budget: number } | null {
  const vars = variables(widthEmu, heightEmu);
  let budget = 0;
  for (const raw of guides) {
    const item = record(raw, ["name", "op", "args"]), { name, op } = item;
    if (typeof name !== "string" || !/^[A-Za-z0-9_]{1,81}$/.test(name) || typeof op !== "string" || !owns(ARITY, op)) return null;
    const override = adjustments && owns(adjustments, name) ? adjustments[name] : undefined;
    vars[name] = override !== undefined && op === "val" ? finite(override) : guide(op, args(item.args, ARITY[op], vars));
    budget += ARITY[op];
  }
  return { vars, budget };
}

/**
 * Every guide of one trusted default program -- the built-ins (`l`, `hc`,
 * `ss`...), the `a:avLst` adjust values (with the caller's `adjustments`
 * substituted by name) and each `a:gdLst` formula -- evaluated in local shape
 * units. This is what an adjust handle's `a:pos` and its `min`/`max` bounds
 * are references into (ECMA-376 §20.1.9.14-15); the path evaluator below runs
 * the very same loop. Unlike the path evaluator it accepts ANY override of a
 * declared adjust name, because a handle drag probes candidate values.
 */
export function evaluateDrawingmlPresetGuides(value: unknown, widthEmu: number, heightEmu: number,
  adjustments?: Readonly<Record<string, number>> | null): Readonly<Record<string, number>> | null {
  try {
    const program = record(value, ["dialect", "preset", "guides", "paths"]);
    if (program.dialect !== "drawingml-default-program-v1" || !Array.isArray(program.guides) || program.guides.length > 2048) return null;
    return guideVariables(program.guides, widthEmu, heightEmu, adjustments)?.vars ?? null;
  } catch { return null; }
}

/** Evaluate a bounded trusted default program in local shape EMU.
 *
 * `adjustments` are the shape's own dragged handles (`a:avLst`), substituted for
 * the program's default adjust values by name — exactly what the server's
 * `_geometry_context` does before it evaluates the same catalog. Only a handle
 * this preset declares may be overridden; anything else refuses the evaluation
 * rather than drawing a different figure.
 */
export function evaluateDrawingmlDefaultPreset(value: unknown, widthEmu: number, heightEmu: number,
  adjustments?: Readonly<Record<string, number>> | null): NativeDrawingGeometry | null {
  try {
    const program = record(value, ["dialect", "preset", "guides", "paths"]);
    if (program.dialect !== "drawingml-default-program-v1" || typeof program.preset !== "string" || !/^[A-Za-z0-9]{1,81}$/.test(program.preset) ||
      !Array.isArray(program.guides) || program.guides.length > 2048 || !Array.isArray(program.paths) || !program.paths.length || program.paths.length > 64) return null;
    const overrides = adjustments ? Object.entries(adjustments) : [];
    if (overrides.length > 8 || overrides.some(([, amount]) => !Number.isSafeInteger(amount))) return null;
    const adjustNames = overrides.length ? drawingmlDefaultPresetAdjustNames(program) : [];
    if (adjustNames === null || overrides.some(([name]) => !adjustNames.includes(name))) return null;
    const evaluated = guideVariables(program.guides, widthEmu, heightEmu, adjustments);
    if (!evaluated) return null;
    const { vars } = evaluated;
    let { budget } = evaluated;
    const paths: NativeDrawingGeometry["paths"] = [];
    const targetWidth = Math.max(1, widthEmu / 9525), targetHeight = Math.max(1, heightEmu / 9525);
    for (const raw of program.paths) {
      const path = record(raw, ["w", "h", "fill", "stroke", "commands"]);
      if (!Array.isArray(path.commands) || !path.commands.length || path.commands.length > 4096 || typeof path.stroke !== "boolean" ||
        !["none", "norm", "darken", "darkenLess", "lighten", "lightenLess"].includes(String(path.fill))) return null;
      const pw = atom(path.w, vars), ph = atom(path.h, vars);
      if (pw < 0 || ph < 0) return null;
      const sx = targetWidth / Math.max(1, pw), sy = targetHeight / Math.max(1, ph);
      const xy = (x: number, y: number) => coordinate(x * sx) + " " + coordinate(y * sy);
      const commands: string[] = [];
      let x = 0, y = 0;
      for (let index = 0; index < path.commands.length; index++) {
        const command = record(path.commands[index], ["op", "args"]), op = command.op;
        if (typeof op !== "string" || !owns(COMMAND_ARITY, op) || (index === 0 && op !== "moveTo")) return null;
        budget += COMMAND_ARITY[op] + 1;
        if (budget > 32_768) return null;
        const values = args(command.args, COMMAND_ARITY[op], vars);
        if (op === "close") commands.push("Z");
        else if (op === "arcTo") {
          const [rx, ry, st, sw] = values, start = st * ANGLE_UNIT, sweep = sw * ANGLE_UNIT;
          if (rx < 0 || ry < 0 || !Number.isFinite(start) || !Number.isFinite(sweep) || Math.abs(sweep) > Math.PI * 20) return null;
          if ((rx === 0 && ry === 0) || sweep === 0) continue;
          if (rx === 0 || ry === 0) {
            const offsets = flatArcPoints(rx, ry, start, sweep), cx = x - offsets[0][0], cy = y - offsets[0][1];
            for (const [px, py] of offsets.slice(1)) {
              if (cx + px === x && cy + py === y) continue;
              x = cx + px; y = cy + py;
              commands.push("L" + xy(x, y));
            }
            continue;
          }
          const [dx, dy] = ellipsePoint(rx, ry, start), cx = x - dx, cy = y - dy;
          const segments = Math.max(1, Math.ceil(Math.abs(sweep) / Math.PI));
          for (let n = 1; n <= segments; n++) {
            const [px, py] = ellipsePoint(rx, ry, start + sweep * n / segments);
            x = cx + px; y = cy + py;
            commands.push(`A${coordinate(rx * sx)} ${coordinate(ry * sy)} 0 0 ${sweep > 0 ? 1 : 0} ${xy(x, y)}`);
          }
        } else {
          const points: string[] = [];
          for (let i = 0; i < values.length; i += 2) {
            if ((pw === 0 && values[i] !== 0) || (ph === 0 && values[i + 1] !== 0)) return null;
            points.push(xy(values[i], values[i + 1]));
          }
          x = values[values.length - 2]; y = values[values.length - 1];
          commands.push(({ moveTo: "M", lnTo: "L", quadBezTo: "Q", cubicBezTo: "C" } as Record<string, string>)[op] + points.join(" "));
        }
      }
      paths.push({ d: commands.join(" "), fill: path.fill as NativePathFill, stroke: path.stroke });
    }
    // Reuse the renderer's final coordinate, token, and path budgets.
    return readNativeDrawingGeometry({ schemaVersion: "drawingml-paths-v1", viewBox: [targetWidth, targetHeight], paths });
  } catch { return null; }
}
