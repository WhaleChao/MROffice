/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The sheet's adapter for the shared arrow picker (`lib/office-line-ends`).
 *
 * A worksheet shape carries its `a:ln` ends on its own style, beside a `kind`
 * that — for a connector — names WHICH ends are decorated. So one pick may
 * move the kind, exactly as it moves a document shape's preset, and lands on
 * the kind this shape's own exported workbook imports back as
 * (docs/office-graphics/SHAPE-IMPORT-AUDIT.md §7.10).
 */

import {
  officeGeometryTakesLineEnds,
  officeShapeLineEndsPatch,
  readOfficeLineEnds,
  type OfficeLineEndsState,
} from "@mrnote/core/web/lib/office-line-ends.ts";
import { isOfficeShapePreset, officeShapeDefinition } from "@mrnote/core/web/lib/office-shape-catalog.ts";
import { presentationPresetGeometryForKind } from "@mrnote/core/web/lib/presentation-authored-preset-geometry.ts";
import {
  normalizePresentationConnectorSpec,
  normalizeShapeKind as normalizePresentationShapeKind,
  type ShapeKind,
} from "@mrnote/core/web/lib/presentation-shapes.ts";
import {
  isSpreadsheetPresetShapeKind,
  normalizeShapeStyle,
  officeShapePresetForSpreadsheetKind,
  spreadsheetShapeKindFromOfficePreset,
  type SpreadsheetShapeStyle,
} from "@mrnote/core/web/lib/spreadsheet/shapes.ts";

const LINE_END_KEYS = [
  "headEnd",
  "tailEnd",
  "headEndWidth",
  "headEndLength",
  "tailEndWidth",
  "tailEndLength",
] as const;

function previewKind(style: SpreadsheetShapeStyle): ShapeKind | null {
  if (!isSpreadsheetPresetShapeKind(style.kind) || style.kind === "text-box") {
    return null;
  }
  const kind: string = style.kind;
  return isOfficeShapePreset(kind)
    ? officeShapeDefinition(kind).previewKind
    : normalizePresentationShapeKind(kind);
}

/** The two ends as the picker reads them: the stated ones over the kind's own. */
export function spreadsheetShapeLineEnds(value: unknown): OfficeLineEndsState {
  const style = normalizeShapeStyle(value);
  const kind = previewKind(style);
  const implied = kind ? normalizePresentationConnectorSpec({ shape: kind }) : null;
  return readOfficeLineEnds(
    style as unknown as Record<string, unknown>,
    implied ?? undefined,
  );
}

/**
 * Whether this shape's outline has ends for an arrowhead to sit on — the rule
 * the canvas draws by, and where Excel greys its own Arrows gallery out.
 */
export function spreadsheetShapeTakesLineEnds(
  value: unknown,
  widthPx: number,
  heightPx: number,
): boolean {
  const style = normalizeShapeStyle(value);
  if (style.customPath) return style.customPath.closed === false;
  const kind = previewKind(style);
  if (!kind) return false;
  if (normalizePresentationConnectorSpec({ shape: kind })) return true;
  const geometry = presentationPresetGeometryForKind(
    kind,
    widthPx,
    heightPx,
    style.adjustments,
  );
  return geometry ? officeGeometryTakesLineEnds(geometry.paths) : false;
}

/**
 * One arrow pick → the next style, or `null` when nothing would change.
 *
 * Both types are written EXPLICITLY, `none` included: the workbook writer
 * leaves an end the page says nothing about with the source package, so
 * silence is not how a user takes an arrowhead off. An absent size stays
 * absent — `@w`/`@len` have no schema default.
 */
export function spreadsheetShapeLineEndsPatch(
  value: unknown,
  ends: OfficeLineEndsState,
  widthPx: number,
  heightPx: number,
): SpreadsheetShapeStyle | null {
  const current = normalizeShapeStyle(value);
  if (!spreadsheetShapeTakesLineEnds(current, widthPx, heightPx)) return null;
  let kind = current.kind;
  if (isSpreadsheetPresetShapeKind(current.kind) && !current.customPath) {
    const patch = officeShapeLineEndsPatch(
      officeShapePresetForSpreadsheetKind(current.kind),
      ends,
    );
    if (!patch) return null;
    kind = spreadsheetShapeKindFromOfficePreset(patch.preset);
  }
  const rest: Record<string, unknown> = { ...current };
  for (const key of LINE_END_KEYS) delete rest[key];
  const next = normalizeShapeStyle({ ...rest, kind, ...ends });
  return LINE_END_KEYS.every((key) => next[key] === current[key]) &&
    next.kind === current.kind
    ? null
    : next;
}
