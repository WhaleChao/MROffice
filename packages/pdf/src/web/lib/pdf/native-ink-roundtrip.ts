/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeInkPenType, type InkPenType } from "@mrnote/core/web/lib/pdf/ink.ts";
import {
  NATIVE_PDF_INK_SOURCE_KIND,
  sanitizeNativePdfInkAnnotationSource,
  type NativePdfInkAnnotationSource,
} from "./native-ink-annotations";

export const MANAGED_NATIVE_INK_PREFIX = "[[MRNOTE_NATIVE_INK_V1|";
export const IMPORTED_NATIVE_INK_PREFIX = "[[MRNOTE_NATIVE_INK_SOURCE_V1|";
/** A source /Ink restyled: its colour and width (_IMPORTED_INK_STYLE_RE in
 * native_edit_primitives.py). */
export const IMPORTED_NATIVE_INK_V2_PREFIX = "[[MRNOTE_NATIVE_INK_SOURCE_V2|";
export const MAX_MANAGED_NATIVE_INK_PATCHES = 400;
export const MAX_MANAGED_NATIVE_INK_POINT_VALUES = 4_000;

const MANAGED_NATIVE_INK_PATTERN =
  /^\[\[MRNOTE_NATIVE_INK_V1\|(upsert|delete)\|([A-Za-z0-9._~%\-]{1,384})(?:\|([0-9A-Fa-f]{6})\|(pen|highlighter)\|([0-9.]{1,20})\|(auto|0|90|180|270)\|([0-9.,-]{3,65536}))?\]\]$/;
const IMPORTED_NATIVE_INK_PATTERN =
  /^\[\[MRNOTE_NATIVE_INK_SOURCE_V1\|(delete_paths|translate_path|transform_paths)\|([1-9][0-9]{0,9}R(?:[0-9]{1,5})?)\|([1-9][0-9]{0,2})\|([0-9]{1,3}(?:,[0-9]{1,3})*)(?:\|(auto|0|90|180|270)\|(-?[0-9.]{1,20})\|(-?[0-9.]{1,20})(?:\|(-?[0-9.]{1,20})\|(-?[0-9.]{1,20}))?)?\]\]$/;

export interface ManagedNativeInkStroke {
  id: string;
  /** Zero-based original PDF page index. */
  page: number;
  color: string;
  width: number;
  /** Flattened page-relative fractions: [x0, y0, x1, y1, ...]. */
  points: number[];
  penType?: InkPenType;
  /** Retained deletion marker so immutable-source patchback can remove it. */
  deleted?: boolean;
  /** Exact third-party `/InkList` path identity discovered through pdf.js. */
  nativeSource?: NativePdfInkAnnotationSource;
  /** The page inserted in MR Office it is drawn on (`page` is then the page
   * before it): lib/pdf/inserted-page-identity. */
  insertedPageId?: string;
}

export interface ManagedNativeInkAnnotationShadow {
  id: string;
  text: string;
  page: number;
  rect: [number, number, number, number];
  /** Written on that inserted page, wherever it is in the file. */
  insertedPageId?: string;
}

export interface ManagedNativeInkPlan {
  annotations: ManagedNativeInkAnnotationShadow[];
  supported: boolean;
  reason?: "invalid_stroke" | "duplicate_id" | "patch_limit";
}

function encodeOpaqueId(value: string): string | null {
  const trimmed = value.trim();
  if (!trimmed) return null;
  const encoded = encodeURIComponent(trimmed).replace(
    /[!'()*]/g,
    (character) =>
      `%${character.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")}`,
  );
  const token = `ink.${encoded}`;
  return token.length <= 384 ? token : null;
}

function numberToken(value: number): string {
  return Number(value.toFixed(6)).toString();
}

function isValidStrokeGeometry(
  stroke: ManagedNativeInkStroke,
): stroke is ManagedNativeInkStroke & { points: number[] } {
  return (
    Number.isInteger(stroke.page) &&
    stroke.page >= 0 &&
    stroke.page <= 100_000 &&
    Number.isFinite(stroke.width) &&
    stroke.width > 0 &&
    stroke.width <= 100 &&
    Array.isArray(stroke.points) &&
    stroke.points.length >= 4 &&
    stroke.points.length <= MAX_MANAGED_NATIVE_INK_POINT_VALUES &&
    stroke.points.length % 2 === 0 &&
    stroke.points.every(
      (coordinate) =>
        typeof coordinate === "number" &&
        Number.isFinite(coordinate) &&
        coordinate >= 0 &&
        coordinate <= 1,
    )
  );
}

function nativeInkDirective(
  stroke: ManagedNativeInkStroke,
  rotation: 0 | 90 | 180 | 270 | "auto",
): string | null {
  const id = encodeOpaqueId(stroke.id);
  if (!id) return null;
  if (stroke.deleted) {
    return `${MANAGED_NATIVE_INK_PREFIX}delete|${id}]]`;
  }
  if (
    !isValidStrokeGeometry(stroke) ||
    !/^#[0-9A-Fa-f]{6}$/.test(stroke.color)
  ) {
    return null;
  }
  const color = stroke.color.slice(1).toUpperCase();
  const penType = normalizeInkPenType(stroke.penType);
  const points = stroke.points.map(numberToken).join(",");
  return (
    `${MANAGED_NATIVE_INK_PREFIX}upsert|${id}|${color}|${penType}|` +
    `${numberToken(stroke.width)}|${rotation}|${points}]]`
  );
}

function importedInkDeleteDirective(
  objectRef: string,
  pathCount: number,
  pathIndexes: readonly number[],
): string {
  return (
    `${IMPORTED_NATIVE_INK_PREFIX}delete_paths|${objectRef}|${pathCount}|` +
    `${pathIndexes.join(",")}]]`
  );
}

function importedInkTranslateDirective(
  objectRef: string,
  pathCount: number,
  pathIndexes: readonly number[],
  rotation: 0 | 90 | 180 | 270 | "auto",
  dx: number,
  dy: number,
): string {
  return (
    `${IMPORTED_NATIVE_INK_PREFIX}translate_path|${objectRef}|${pathCount}|` +
    `${pathIndexes.join(",")}|${rotation}|${numberToken(dx)}|${numberToken(dy)}]]`
  );
}

function importedInkTransformDirective(
  objectRef: string,
  pathCount: number,
  pathIndexes: readonly number[],
  rotation: 0 | 90 | 180 | 270 | "auto",
  scaleX: number,
  scaleY: number,
  translateX: number,
  translateY: number,
): string {
  return (
    `${IMPORTED_NATIVE_INK_PREFIX}transform_paths|${objectRef}|${pathCount}|` +
    `${pathIndexes.join(",")}|${rotation}|${numberToken(scaleX)}|` +
    `${numberToken(scaleY)}|${numberToken(translateX)}|${numberToken(translateY)}]]`
  );
}

/**
 * Build invisible annotation rows consumed by the existing PDF patchback
 * bridge. The rows remain in the page snapshot for deterministic replays but
 * are filtered from the comments UI by {@link isManagedNativeInkAnnotationText}.
 *
 * The plan is all-or-nothing. An invalid/oversized batch yields no native
 * annotation rows, allowing the caller to fail closed instead of writing only
 * a prefix of the user's strokes.
 */
export function planManagedNativeInkAnnotations(
  strokes: readonly ManagedNativeInkStroke[],
  rotations: ReadonlyMap<number, 0 | 90 | 180 | 270> = new Map(),
): ManagedNativeInkPlan {
  if (strokes.length === 0) {
    return { annotations: [], supported: true };
  }
  const ids = new Set<string>();
  const annotations: ManagedNativeInkAnnotationShadow[] = [];
  const importedDeletes = new Map<
    string,
    {
      page: number;
      pathCount: number;
      pathIndexes: Set<number>;
    }
  >();
  // One restyle per source object: all its paths share its /C and /BS.
  const importedStyles = new Map<
    string,
    { page: number; pathCount: number; color: string; width: number }
  >();
  const importedTransforms = new Map<
    string,
    {
      page: number;
      pathCount: number;
      rotation: 0 | 90 | 180 | 270 | "auto";
      paths: Map<number, { original: number[]; current: number[] }>;
    }
  >();
  for (const stroke of strokes) {
    if (ids.has(stroke.id)) {
      return { annotations: [], supported: false, reason: "duplicate_id" };
    }
    ids.add(stroke.id);
    if (stroke.nativeSource !== undefined) {
      const source = sanitizeNativePdfInkAnnotationSource(stroke.nativeSource);
      if (
        !source ||
        source.kind !== NATIVE_PDF_INK_SOURCE_KIND ||
        source.pathCount < 1
      ) {
        return { annotations: [], supported: false, reason: "invalid_stroke" };
      }
      if (!stroke.deleted) {
        const restyled =
          (source.originalColor !== undefined &&
            stroke.color.toUpperCase() !== source.originalColor) ||
          (source.originalWidth !== undefined &&
            Math.abs(stroke.width - source.originalWidth) > 0.01);
        if (restyled) {
          const style = importedStyles.get(source.objectRef);
          if (
            !/^#[0-9A-Fa-f]{6}$/.test(stroke.color) ||
            !Number.isFinite(stroke.width) ||
            stroke.width < 0.1 ||
            stroke.width > 100 ||
            (style &&
              (style.color !== stroke.color.toUpperCase() ||
                Math.abs(style.width - stroke.width) > 0.01 ||
                style.pathCount !== source.pathCount))
          ) {
            return { annotations: [], supported: false, reason: "invalid_stroke" };
          }
          importedStyles.set(source.objectRef, {
            page: stroke.page,
            pathCount: source.pathCount,
            color: stroke.color.toUpperCase(),
            width: stroke.width,
          });
        }
        const original = source.originalPoints;
        if (!original) continue;
        if (original.length !== stroke.points.length) {
          return { annotations: [], supported: false, reason: "invalid_stroke" };
        }
        const rotation = rotations.get(stroke.page + 1) ?? "auto";
        const current = importedTransforms.get(source.objectRef);
        if (
          current &&
          (current.page !== stroke.page ||
            current.pathCount !== source.pathCount ||
            current.rotation !== rotation)
        ) {
          return { annotations: [], supported: false, reason: "invalid_stroke" };
        }
        const group = current ?? {
          page: stroke.page,
          pathCount: source.pathCount,
          rotation,
          paths: new Map(),
        };
        if (group.paths.has(source.pathIndex)) {
          return { annotations: [], supported: false, reason: "invalid_stroke" };
        }
        group.paths.set(source.pathIndex, {
          original: [...original],
          current: [...stroke.points],
        });
        importedTransforms.set(source.objectRef, group);
        continue;
      }
      const current = importedDeletes.get(source.objectRef);
      if (
        current &&
        (current.page !== stroke.page || current.pathCount !== source.pathCount)
      ) {
        return { annotations: [], supported: false, reason: "invalid_stroke" };
      }
      const group = current ?? {
        page: stroke.page,
        pathCount: source.pathCount,
        pathIndexes: new Set<number>(),
      };
      if (group.pathIndexes.has(source.pathIndex)) {
        return { annotations: [], supported: false, reason: "invalid_stroke" };
      }
      group.pathIndexes.add(source.pathIndex);
      importedDeletes.set(source.objectRef, group);
      continue;
    }
    // An inserted page is not turned with the page before it.
    const rotation = stroke.insertedPageId ? "auto" : (rotations.get(stroke.page + 1) ?? "auto");
    const text = nativeInkDirective(stroke, rotation);
    if (!text) {
      return { annotations: [], supported: false, reason: "invalid_stroke" };
    }
    annotations.push({
      id: `__mrnote_native_ink__${stroke.id}`,
      text,
      page: stroke.page + 1,
      // The native primitive derives the exact Rect from InkList. This bounded
      // placeholder only satisfies the shared annotation transport contract.
      rect: [24, 24, 44, 44],
      ...(stroke.insertedPageId && !stroke.deleted ? { insertedPageId: stroke.insertedPageId } : {}),
    });
  }
  for (const [objectRef, group] of importedTransforms) {
    const pathIndexes = [...group.paths.keys()].sort((left, right) => left - right);
    if (
      pathIndexes.length !== group.pathCount ||
      pathIndexes.some((pathIndex, index) => pathIndex !== index)
    ) {
      return { annotations: [], supported: false, reason: "invalid_stroke" };
    }
    const original = pathIndexes.flatMap(
      (index) => group.paths.get(index)!.original,
    );
    const current = pathIndexes.flatMap(
      (index) => group.paths.get(index)!.current,
    );
    const axis = (offset: 0 | 1) => {
      const sourceValues = original.filter((_, index) => index % 2 === offset);
      const targetValues = current.filter((_, index) => index % 2 === offset);
      const sourceMin = Math.min(...sourceValues);
      const sourceMax = Math.max(...sourceValues);
      const targetMin = Math.min(...targetValues);
      const targetMax = Math.max(...targetValues);
      const sourceExtent = sourceMax - sourceMin;
      const targetExtent = targetMax - targetMin;
      if (sourceExtent <= 0.000001) {
        return targetExtent <= 0.000001
          ? { scale: 1, translate: targetMin - sourceMin }
          : null;
      }
      const scale = targetExtent / sourceExtent;
      return Number.isFinite(scale) && scale >= 0.01 && scale <= 100
        ? { scale, translate: targetMin - scale * sourceMin }
        : null;
    };
    const x = axis(0);
    const y = axis(1);
    if (
      !x ||
      !y ||
      current.some(
        (value, index) =>
          Math.abs(
            value -
              ((index % 2 === 0 ? x.scale : y.scale) * original[index] +
                (index % 2 === 0 ? x.translate : y.translate)),
          ) > 0.000001,
      )
    ) {
      return { annotations: [], supported: false, reason: "invalid_stroke" };
    }
    const translated =
      Math.abs(x.scale - 1) <= 0.000001 &&
      Math.abs(y.scale - 1) <= 0.000001;
    const unchanged =
      translated &&
      Math.abs(x.translate) <= 0.000001 &&
      Math.abs(y.translate) <= 0.000001;
    if (unchanged) continue;
    annotations.push({
      id: `__mrnote_native_ink_source__${objectRef}`,
      text: translated
        ? importedInkTranslateDirective(
            objectRef,
            group.pathCount,
            pathIndexes,
            group.rotation,
            x.translate,
            y.translate,
          )
        : importedInkTransformDirective(
            objectRef,
            group.pathCount,
            pathIndexes,
            group.rotation,
            x.scale,
            y.scale,
            x.translate,
            y.translate,
          ),
      page: group.page + 1,
      rect: [24, 24, 44, 44],
    });
  }
  for (const [objectRef, style] of importedStyles) {
    // A deleted object takes no restyle.
    if (importedDeletes.has(objectRef)) continue;
    annotations.push({
      id: `__mrnote_native_ink_source_style__${objectRef}`,
      text:
        `${IMPORTED_NATIVE_INK_V2_PREFIX}style|${objectRef}|${style.pathCount}|` +
        `${style.color.slice(1)}|${numberToken(style.width)}]]`,
      page: style.page + 1,
      rect: [24, 24, 44, 44],
    });
  }
  for (const [objectRef, group] of importedDeletes) {
    const pathIndexes = [...group.pathIndexes].sort(
      (left, right) => left - right,
    );
    if (
      pathIndexes.length !== group.pathCount ||
      pathIndexes.some((pathIndex, index) => pathIndex !== index)
    ) {
      return { annotations: [], supported: false, reason: "invalid_stroke" };
    }
    annotations.push({
      id: `__mrnote_native_ink_source__${objectRef}`,
      text: importedInkDeleteDirective(objectRef, group.pathCount, pathIndexes),
      page: group.page + 1,
      rect: [24, 24, 44, 44],
    });
  }
  if (annotations.length > MAX_MANAGED_NATIVE_INK_PATCHES) {
    return { annotations: [], supported: false, reason: "patch_limit" };
  }
  return { annotations, supported: true };
}

export function isManagedNativeInkAnnotationText(
  value: unknown,
): value is string {
  return (
    typeof value === "string" &&
    (value.startsWith(MANAGED_NATIVE_INK_PREFIX) ||
      value.startsWith(IMPORTED_NATIVE_INK_PREFIX) ||
      value.startsWith(IMPORTED_NATIVE_INK_V2_PREFIX))
  );
}

export function parseImportedNativeInkDirective(value: unknown): {
  operation: "delete_paths" | "translate_path" | "transform_paths";
  objectRef: string;
  pathCount: number;
  pathIndexes: number[];
  rotation?: 0 | 90 | 180 | 270 | "auto";
  dx?: number;
  dy?: number;
  scaleX?: number;
  scaleY?: number;
  translateX?: number;
  translateY?: number;
} | null {
  if (typeof value !== "string") return null;
  const match = IMPORTED_NATIVE_INK_PATTERN.exec(value);
  if (!match) return null;
  const operation = match[1] as
    | "delete_paths"
    | "translate_path"
    | "transform_paths";
  const pathCount = Number(match[3]);
  const pathIndexes = match[4].split(",").map(Number);
  if (
    !Number.isInteger(pathCount) ||
    pathIndexes.length !== pathCount ||
    pathIndexes.some(
      (pathIndex, index) =>
        !Number.isInteger(pathIndex) ||
        pathIndex < 0 ||
        pathIndex !== index,
    )
  ) {
    return null;
  }
  const rotationToken = match[5];
  const dx = match[6] === undefined ? undefined : Number(match[6]);
  const dy = match[7] === undefined ? undefined : Number(match[7]);
  const third = match[8] === undefined ? undefined : Number(match[8]);
  const fourth = match[9] === undefined ? undefined : Number(match[9]);
  if (
    (operation === "delete_paths" && rotationToken !== undefined) ||
    ((operation === "translate_path" || operation === "transform_paths") &&
      (rotationToken === undefined ||
        dx === undefined ||
        dy === undefined ||
        !Number.isFinite(dx) ||
        !Number.isFinite(dy) ||
        Math.abs(dx) > (operation === "translate_path" ? 1 : 100) ||
        Math.abs(dy) > (operation === "translate_path" ? 1 : 100))) ||
    (operation === "translate_path" &&
      (third !== undefined || fourth !== undefined)) ||
    (operation === "transform_paths" &&
      (third === undefined ||
        fourth === undefined ||
        !Number.isFinite(third) ||
        !Number.isFinite(fourth) ||
        dx! < 0.01 ||
        dy! < 0.01 ||
        dx! > 100 ||
        dy! > 100 ||
        Math.abs(third) > 100 ||
        Math.abs(fourth) > 100))
  ) {
    return null;
  }
  return {
    operation,
    objectRef: match[2],
    pathCount,
    pathIndexes,
    ...(operation === "translate_path"
      ? {
          rotation:
            rotationToken === "auto"
              ? "auto"
              : (Number(rotationToken) as 0 | 90 | 180 | 270),
          dx,
          dy,
        }
      : operation === "transform_paths"
        ? {
            rotation:
              rotationToken === "auto"
                ? "auto"
                : (Number(rotationToken) as 0 | 90 | 180 | 270),
            scaleX: dx,
            scaleY: dy,
            translateX: third,
            translateY: fourth,
          }
      : {}),
  };
}

export function parseManagedNativeInkDirective(value: unknown): {
  operation: "upsert" | "delete";
  id: string;
  color?: string;
  penType?: InkPenType;
  width?: number;
  rotation?: 0 | 90 | 180 | 270 | "auto";
  points?: number[];
} | null {
  if (typeof value !== "string") return null;
  const match = MANAGED_NATIVE_INK_PATTERN.exec(value);
  if (!match) return null;
  const operation = match[1] as "upsert" | "delete";
  if (operation === "delete") {
    return { operation, id: match[2] };
  }
  const width = Number(match[5]);
  const rotation =
    match[6] === "auto" ? "auto" : (Number(match[6]) as 0 | 90 | 180 | 270);
  const points = match[7].split(",").map(Number);
  if (
    !Number.isFinite(width) ||
    width <= 0 ||
    width > 100 ||
    points.length < 4 ||
    points.length > MAX_MANAGED_NATIVE_INK_POINT_VALUES ||
    points.length % 2 !== 0 ||
    points.some(
      (coordinate) =>
        !Number.isFinite(coordinate) || coordinate < 0 || coordinate > 1,
    )
  ) {
    return null;
  }
  return {
    operation,
    id: match[2],
    color: `#${match[3].toUpperCase()}`,
    penType: match[4] as InkPenType,
    width,
    rotation,
    points,
  };
}
