/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { evaluateDrawingmlDefaultPreset } from "@mrnote/core/web/lib/drawingml-default-preset-program.ts";
import { DEFAULT_PRESET_PROGRAMS } from "@mrnote/core/web/lib/presentation-default-preset-programs.generated.ts";
import { readNativeDrawingGeometry, type NativeDrawingGeometry } from "@mrnote/core/web/lib/native-drawing-geometry.ts";
import { presentationShapeCatalogEntryForPreset } from "@mrnote/core/web/lib/presentation-shape-catalog.ts";

const owns = (value: object, key: string) => Object.prototype.hasOwnProperty.call(value, key);
function sameGeometry(actual: NativeDrawingGeometry, expected: NativeDrawingGeometry): boolean {
  if (!actual.viewBox || !expected.viewBox || actual.viewBox.some((n, i) => Math.abs(n - expected.viewBox![i]) > 1e-7) || actual.paths.length !== expected.paths.length) return false;
  return actual.paths.every((path, index) => {
    const other = expected.paths[index];
    if (path.fill !== other.fill || path.stroke !== other.stroke) return false;
    const tokens = (d: string) => d.match(/[MLQCAZ]|[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[-+]?\d+)?/g) ?? [];
    const left = tokens(path.d), right = tokens(other.d);
    return left.length === right.length && left.every((value, i) => value === right[i] ||
      (Number.isFinite(Number(value)) && Number.isFinite(Number(right[i])) && Math.abs(Number(value) - Number(right[i])) <= 1e-6));
  });
}

/** Only authenticated native defaults can opt into the shared template engine.
 * Source equality is a consistency check after the exact source reference; it
 * never establishes that a shape uses defaults. Missing legacy references keep
 * the existing path, including the separately authenticated six-arrow helpers.
 */
export function presentationDefaultPresetGeometry(kind: string, widthPixels: number, heightPixels: number, drawing?: Record<string, unknown>): NativeDrawingGeometry | null {
  if (!drawing || drawing.schemaVersion !== "pptx-shape-drawing-v1" || drawing.role !== "shape" || drawing.shapeKind !== kind ||
    owns(drawing, "customPathIdentity") || owns(drawing, "connectorIdentity")) return null;
  const raw = drawing.defaultProgram;
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const ref = raw as Record<string, unknown>;
  const fields = ["schemaVersion", "preset", "definitionHash", "sourceSizeEmu"];
  if (Object.keys(ref).length !== fields.length || fields.some(key => !owns(ref, key)) || ref.schemaVersion !== "pptx-default-preset-program-v1" ||
    typeof ref.preset !== "string" || !owns(DEFAULT_PRESET_PROGRAMS, ref.preset) || typeof ref.definitionHash !== "string" ||
    presentationShapeCatalogEntryForPreset(ref.preset)?.kind !== kind ||
    !Array.isArray(ref.sourceSizeEmu) || ref.sourceSizeEmu.length !== 2 || ref.sourceSizeEmu.some(n => typeof n !== "number" || !Number.isFinite(n) || n <= 0 || n > 1e8) ||
    !Number.isFinite(widthPixels) || !Number.isFinite(heightPixels) || widthPixels <= 0 || heightPixels <= 0) return null;
  const entry = DEFAULT_PRESET_PROGRAMS[ref.preset] as { definitionHash: string; program: unknown };
  if (ref.definitionHash !== entry.definitionHash) return null;
  const native = readNativeDrawingGeometry(drawing.geometry);
  const source = evaluateDrawingmlDefaultPreset(entry.program, ref.sourceSizeEmu[0], ref.sourceSizeEmu[1]);
  if (!native || !source || !sameGeometry(native, source)) return null;
  return evaluateDrawingmlDefaultPreset(entry.program, widthPixels * 9525, heightPixels * 9525);
}
