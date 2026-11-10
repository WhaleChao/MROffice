/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  MEASURE_UNITS,
  type MeasureMode,
  type MeasurePoint,
  type MeasureScale,
} from "./measure";

export const MANAGED_NATIVE_MEASUREMENT_PREFIX = "[[MRNOTE_NATIVE_MEASURE_V1|";
export const MAX_MANAGED_NATIVE_MEASUREMENTS = 400;

const MANAGED_NATIVE_MEASUREMENT_PATTERN =
  /^\[\[MRNOTE_NATIVE_MEASURE_V1\|(upsert|delete)\|([A-Za-z0-9._~%\-]{1,384})(?:\|(distance|perimeter|area)\|([0-9A-Fa-f]{6})\|(pt|in|cm|mm|ft|m)\|(pt|in|cm|mm|ft|m)\|([0-9.]{1,24})\|(auto|0|90|180|270)\|([0-9.,-]{7,16384})\|(v[A-Za-z0-9._~%\-]{1,4096})\|(none|v[A-Za-z0-9._~%\-]{1,2048})\|(none|v[A-Za-z0-9._~%:\-]{1,512}))?\]\]$/u;

/** A measurement the user made and wants kept. */
export interface MeasurementRecord {
  id: string;
  /** 1-based. */
  page: number;
  mode: MeasureMode;
  /** Click points in page-relative percent, the same space shapes use. */
  points: MeasurePoint[];
  scale: MeasureScale;
  /** The readout as it was shown. */
  label: string;
  /** `#rrggbb`; defaults to the shape default. */
  color?: string;
  author?: string;
  /** PDF date string, supplied rather than read from the clock. */
  created?: string;
}

export type ManagedNativeMeasurementRecord =
  | (MeasurementRecord & { deleted?: false })
  | {
      id: string;
      /** One-based original PDF page number. */
      page: number;
      deleted: true;
    };

export interface ManagedNativeMeasurementAnnotationShadow {
  id: string;
  text: string;
  page: number;
  rect: [number, number, number, number];
}

export interface ManagedNativeMeasurementPlan {
  annotations: ManagedNativeMeasurementAnnotationShadow[];
  supported: boolean;
  reason?: "invalid_measurement" | "duplicate_id" | "patch_limit";
}

export interface ParsedManagedNativeMeasurementDirective {
  operation: "upsert" | "delete";
  id: string;
  nativeId: string;
  rotation?: 0 | 90 | 180 | 270 | "auto";
  measurement?: Omit<MeasurementRecord, "id" | "page">;
}

function encodeOpaqueMeasurementId(value: string): string | null {
  const trimmed = value.trim();
  if (!trimmed) return null;
  const encoded = encodeURIComponent(trimmed).replace(
    /[!'()*]/g,
    (character) =>
      `%${character.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")}`,
  );
  const token = `measure.${encoded}`;
  return token.length <= 384 ? token : null;
}

function decodeOpaqueMeasurementId(value: string): string | null {
  if (!value.startsWith("measure.")) return null;
  try {
    const decoded = decodeURIComponent(value.slice("measure.".length));
    return decoded.trim() ? decoded : null;
  } catch {
    return null;
  }
}

function encodeMeasurementText(
  value: string | undefined,
  maxLength: number,
): string | null {
  if (value === undefined) return "none";
  if (!value || value.length > maxLength) return null;
  const encoded = encodeURIComponent(value).replace(
    /[!'()*]/g,
    (character) =>
      `%${character.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")}`,
  );
  return encoded.length <= maxLength * 8 ? `v${encoded}` : null;
}

function decodeMeasurementText(
  token: string,
  maxLength: number,
): string | null | undefined {
  if (token === "none") return undefined;
  if (!token.startsWith("v")) return null;
  try {
    const decoded = decodeURIComponent(token.slice(1));
    return decoded && decoded.length <= maxLength ? decoded : null;
  } catch {
    return null;
  }
}

function finitePercent(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isFinite(value) &&
    value >= 0 &&
    value <= 100
  );
}

function isValidMeasurementPage(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isInteger(value) &&
    value >= 1 &&
    value <= 10_000
  );
}

function nativeNumberToken(value: number): string {
  return Number(value.toFixed(6)).toString();
}

function normalizedMeasurementPoints(
  mode: MeasureMode,
  points: readonly MeasurePoint[],
): MeasurePoint[] | null {
  if (
    !Array.isArray(points) ||
    points.some((point) => !finitePercent(point.x) || !finitePercent(point.y))
  ) {
    return null;
  }
  const normalized =
    mode === "distance" && points.length >= 2
      ? [points[0], points[points.length - 1]]
      : [...points];
  const minimum = mode === "distance" ? 2 : 3;
  const maximum = mode === "distance" ? 2 : 200;
  if (normalized.length < minimum || normalized.length > maximum) return null;
  if (
    normalized.every(
      (point) =>
        Math.abs(point.x - normalized[0].x) < 0.000_001 &&
        Math.abs(point.y - normalized[0].y) < 0.000_001,
    )
  ) {
    return null;
  }
  return normalized.map((point) => ({ x: point.x, y: point.y }));
}

function nativeMeasurementDirective(
  record: ManagedNativeMeasurementRecord,
  rotation: 0 | 90 | 180 | 270 | "auto",
): string | null {
  const nativeId = encodeOpaqueMeasurementId(record.id);
  if (!nativeId || !isValidMeasurementPage(record.page)) return null;
  if (record.deleted) {
    return `${MANAGED_NATIVE_MEASUREMENT_PREFIX}delete|${nativeId}]]`;
  }
  if (!(["distance", "perimeter", "area"] as const).includes(record.mode)) {
    return null;
  }
  const points = normalizedMeasurementPoints(record.mode, record.points);
  const ratio = record.scale?.outputPerPage;
  if (
    !points ||
    !(MEASURE_UNITS as readonly string[]).includes(record.scale?.pageUnit) ||
    !(MEASURE_UNITS as readonly string[]).includes(record.scale?.outputUnit) ||
    typeof ratio !== "number" ||
    !Number.isFinite(ratio) ||
    ratio < 1e-6 ||
    ratio > 1e9
  ) {
    return null;
  }
  const colorMatch = /^#([0-9a-f]{6})$/i.exec(record.color ?? "#2563eb");
  const label = encodeMeasurementText(record.label, 500);
  const author = encodeMeasurementText(record.author, 200);
  const created = encodeMeasurementText(record.created, 64);
  if (!colorMatch || !label || !author || !created) return null;
  const pointToken = points
    .flatMap((point) => [point.x / 100, point.y / 100])
    .map(nativeNumberToken)
    .join(",");
  return (
    `${MANAGED_NATIVE_MEASUREMENT_PREFIX}upsert|${nativeId}|${record.mode}|` +
    `${colorMatch[1].toUpperCase()}|${record.scale.pageUnit}|` +
    `${record.scale.outputUnit}|${nativeNumberToken(ratio)}|${rotation}|` +
    `${pointToken}|${label}|${author}|${created}]]`
  );
}

/** Project stable measurement records into original-package patch rows. */
export function planManagedNativeMeasurementAnnotations(
  records: readonly ManagedNativeMeasurementRecord[],
  // Kept for the callers' shared signature; measurements do not use it (see
  // the directive below).
  _rotations: ReadonlyMap<number, 0 | 90 | 180 | 270> = new Map(),
): ManagedNativeMeasurementPlan {
  if (records.length === 0) return { annotations: [], supported: true };
  if (records.length > MAX_MANAGED_NATIVE_MEASUREMENTS) {
    return { annotations: [], supported: false, reason: "patch_limit" };
  }
  const ids = new Set<string>();
  const annotations: ManagedNativeMeasurementAnnotationShadow[] = [];
  for (const record of records) {
    if (ids.has(record.id)) {
      return { annotations: [], supported: false, reason: "duplicate_id" };
    }
    ids.add(record.id);
    // A record's points are already in the page's unrotated (user-space)
    // orientation: the reader unrotates the clicked points before it stores
    // them (pendingMeasurementDraft). Handing the page rotation on made the
    // writer unrotate them a second time -- on a /Rotate 90 page the saved
    
    const text = nativeMeasurementDirective(record, 0);
    if (!text) {
      return {
        annotations: [],
        supported: false,
        reason: "invalid_measurement",
      };
    }
    annotations.push({
      id: `__mrnote_native_measurement__${record.id}`,
      text,
      page: record.page,
      rect: [24, 24, 44, 44],
    });
  }
  return { annotations, supported: true };
}

export function isManagedNativeMeasurementAnnotationText(
  value: unknown,
): value is string {
  return (
    typeof value === "string" &&
    value.startsWith(MANAGED_NATIVE_MEASUREMENT_PREFIX)
  );
}

export function parseManagedNativeMeasurementDirective(
  value: unknown,
): ParsedManagedNativeMeasurementDirective | null {
  if (typeof value !== "string") return null;
  const match = MANAGED_NATIVE_MEASUREMENT_PATTERN.exec(value);
  if (!match) return null;
  const operation = match[1] as "upsert" | "delete";
  const nativeId = match[2];
  const id = decodeOpaqueMeasurementId(nativeId);
  if (!id) return null;
  if (operation === "delete") {
    return match[3] === undefined ? { operation, id, nativeId } : null;
  }
  const mode = match[3] as MeasureMode;
  const ratio = Number(match[7]);
  const rotation =
    match[8] === "auto" ? "auto" : (Number(match[8]) as 0 | 90 | 180 | 270);
  const pointValues = match[9].split(",").map(Number);
  const points: MeasurePoint[] = [];
  for (let index = 0; index < pointValues.length; index += 2) {
    points.push({
      x: Number((pointValues[index] * 100).toFixed(6)),
      y: Number((pointValues[index + 1] * 100).toFixed(6)),
    });
  }
  const normalizedPoints = normalizedMeasurementPoints(mode, points);
  const label = decodeMeasurementText(match[10], 500);
  const author = decodeMeasurementText(match[11], 200);
  const created = decodeMeasurementText(match[12], 64);
  if (
    !Number.isFinite(ratio) ||
    ratio < 1e-6 ||
    ratio > 1e9 ||
    pointValues.length % 2 !== 0 ||
    !normalizedPoints ||
    label === null ||
    label === undefined ||
    author === null ||
    created === null
  ) {
    return null;
  }
  return {
    operation,
    id,
    nativeId,
    rotation,
    measurement: {
      mode,
      points: normalizedPoints,
      scale: {
        pageUnit: match[5] as MeasureScale["pageUnit"],
        outputUnit: match[6] as MeasureScale["outputUnit"],
        outputPerPage: ratio,
      },
      label,
      color: `#${match[4].toUpperCase()}`,
      ...(author === undefined ? {} : { author }),
      ...(created === undefined ? {} : { created }),
    },
  };
}

export function parseManagedNativeMeasurementShadow(
  value: unknown,
): ManagedNativeMeasurementRecord | null {
  if (!value || typeof value !== "object") return null;
  const row = value as Record<string, unknown>;
  const page = Number(row.page);
  if (!isValidMeasurementPage(page)) return null;
  const parsed = parseManagedNativeMeasurementDirective(
    row.text ?? row.comment,
  );
  if (!parsed) return null;
  if (parsed.operation === "delete") {
    return { id: parsed.id, page, deleted: true };
  }
  return parsed.measurement
    ? { id: parsed.id, page, ...parsed.measurement }
    : null;
}
