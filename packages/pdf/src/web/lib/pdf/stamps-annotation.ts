/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for PDF rubber STAMP annotations (Acrobat "Stamp" parity).
 *
 * A stamp annotation is a single box on a page (page-relative *percentage*
 * coordinates [0..100], exactly like `pdf_signatures`) holding EITHER:
 *
 *   - a PRESET stamp: one of an allow-listed palette of styled-text stamps
 *     (`kind: "preset"`, `preset` key — e.g. APPROVED / 审阅 / 重要 / ✓),
 *     rendered as bold bordered text in the preset's color, or
 *   - an IMAGE stamp: a user-uploaded raster (`kind: "image"`, `image` data
 *     URL), sanitized through the SAME strict PNG/JPEG-only guard the drawn
 *     signature path uses (lib/pdf/signatures.ts `sanitizeSignatureImage`).
 *
 * Persisted as a SEPARATE `pdf_stamps` annotation type, additive to
 * `pdf_highlight` / `pdf_note` / `pdf_ink` / `pdf_signatures` — old PDFs keep
 * loading unchanged. Structurally this module mirrors `signatures.ts` (the
 * sanitizer / builder / clamp / per-page-cap quartet) so the viewer wires it
 * up identically (place via click, move/resize via PdfAnnotationFrame, bake at
 * export through the signature-image op with `source: "stamp"`).
 *
 * NOTE: distinct from lib/pdf/stamps.ts, which is the EXPORT-TIME watermark /
 * page-numbering ("Bates stamp") config — not a placed page annotation.
 */

import { isSafeHexColor } from "@mrnote/core/web/lib/pdf/safe-color.ts";
import { sanitizeSignatureImage } from "@mrnote/core/web/lib/pdf/signatures.ts";

/** Discriminates a preset styled-text stamp from an uploaded image stamp. */
export type StampAnnotationKind = "preset" | "image";

/**
 * One preset stamp: the persisted `key`, the exact `text` rendered/baked, and
 * a safe hex `color` for the text + border. Kept as an allow-list (like the
 * signature font allow-list) so a crafted payload can't inject arbitrary CSS —
 * the preset KEY is what is persisted; text/color always resolve through this
 * table.
 */
export interface StampPreset {
  key: string;
  text: string;
  color: string;
}

/**
 * The preset palette, in display order. Text is the literal stamped CONTENT
 * (it bakes into the exported PDF), so it is intentionally NOT localized —
 * both English and Chinese stamps are offered side by side, like a physical
 * stamp drawer.
 */
export const STAMP_PRESETS: readonly StampPreset[] = [
  { key: "approved", text: "APPROVED", color: "#15803d" },
  { key: "rejected", text: "REJECTED", color: "#b91c1c" },
  { key: "draft", text: "DRAFT", color: "#6b7280" },
  { key: "confidential", text: "CONFIDENTIAL", color: "#b91c1c" },
  { key: "final", text: "FINAL", color: "#1d4ed8" },
  { key: "reviewed", text: "已审阅", color: "#1d4ed8" },
  { key: "important", text: "重要", color: "#b91c1c" },
  { key: "urgent", text: "加急", color: "#c2410c" },
  { key: "check", text: "✓", color: "#15803d" },
  { key: "cross", text: "✗", color: "#b91c1c" },
];

const STAMP_PRESETS_BY_KEY = new Map(STAMP_PRESETS.map((p) => [p.key, p]));

/** Resolve a persisted preset key to its palette entry, or null when unknown. */
export function resolveStampPreset(key: unknown): StampPreset | null {
  if (typeof key !== "string") return null;
  return STAMP_PRESETS_BY_KEY.get(key) ?? null;
}

/**
 * Max stamps kept per page. Mirrors SIGNATURE_MAX_PER_PAGE (a
 * generous-but-bounded per-page cap).
 */
export const STAMP_MAX_PER_PAGE = 50;

/** Min / max box size in page-percentage units (resolution-independent). */
export const STAMP_MIN_SIZE = 1;
export const STAMP_MAX_SIZE = 100;
/** Default box size for a freshly placed stamp (page-percentage units). */
export const DEFAULT_STAMP_WIDTH = 18;
export const DEFAULT_STAMP_HEIGHT = 7;

/**
 * Size a newly uploaded image stamp without stretching its raster.
 *
 * Stamp boxes use page-relative percentages, so preserving a pixel aspect
 * ratio must account for the page's own rendered width/height. The result is
 * scaled uniformly to fit the page and the remaining space after the click.
 * Return null when the image/page geometry is invalid or the remaining area
 * cannot hold the minimum 1% box without distorting it.
 */
export function imageStampPlacementSize(input: {
  imageAspectRatio: number;
  pageWidth: number;
  pageHeight: number;
  x: number;
  y: number;
  preferredWidth?: number;
}): { width: number; height: number } | null {
  const {
    imageAspectRatio,
    pageWidth,
    pageHeight,
    x,
    y,
    preferredWidth = DEFAULT_STAMP_WIDTH,
  } = input;
  if (
    !Number.isFinite(imageAspectRatio) ||
    imageAspectRatio <= 0 ||
    !Number.isFinite(pageWidth) ||
    pageWidth <= 0 ||
    !Number.isFinite(pageHeight) ||
    pageHeight <= 0 ||
    !isFinitePercent(x) ||
    !isFinitePercent(y)
  ) {
    return null;
  }

  let width = clampSize(preferredWidth, DEFAULT_STAMP_WIDTH);
  let height = (width * pageWidth) / (pageHeight * imageAspectRatio);
  if (!Number.isFinite(height) || height <= 0) return null;

  // Find one uniform scale that keeps both percentage dimensions within the
  // persisted stamp contract. If no such scale exists, preserving the image
  // aspect ratio is impossible and placement fails closed.
  const minimumScale = Math.max(
    STAMP_MIN_SIZE / width,
    STAMP_MIN_SIZE / height,
  );
  const maximumScale = Math.min(
    STAMP_MAX_SIZE / width,
    STAMP_MAX_SIZE / height,
  );
  if (minimumScale > maximumScale) return null;
  const boundedScale = Math.min(maximumScale, Math.max(minimumScale, 1));
  width *= boundedScale;
  height *= boundedScale;

  const availableWidth = 100 - x;
  const availableHeight = 100 - y;
  const fitScale = Math.min(
    1,
    availableWidth / width,
    availableHeight / height,
  );
  width *= fitScale;
  height *= fitScale;
  if (
    !Number.isFinite(width) ||
    !Number.isFinite(height) ||
    width < STAMP_MIN_SIZE ||
    height < STAMP_MIN_SIZE
  ) {
    return null;
  }
  return { width, height };
}

export interface StampAnnotation {
  id: string;
  /** 1-based page number. */
  page: number;
  /** Box top-left + size in page-relative percent [0..100]. */
  x: number;
  y: number;
  width: number;
  height: number;
  /** "preset" (styled text) or "image" (uploaded raster). */
  kind: StampAnnotationKind;
  /**
   * For a PRESET stamp: the palette key (always resolvable via
   * {@link resolveStampPreset}). Empty string for an image stamp.
   */
  preset: string;
  /**
   * For an IMAGE stamp: a sanitized PNG/JPEG base64 data URL (render as an
   * `<img src>`; same guard as drawn signatures). Empty string for a preset.
   */
  image: string;
}

function isFinitePercent(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isFinite(value) &&
    value >= 0 &&
    value <= 100
  );
}

/** Clamp a single coordinate value into the on-page percentage range. */
function clampPercent(value: number): number {
  if (!Number.isFinite(value)) return 0;
  if (value < 0) return 0;
  if (value > 100) return 100;
  return value;
}

/** Clamp a box dimension into [{@link STAMP_MIN_SIZE}, {@link STAMP_MAX_SIZE}]. */
function clampSize(value: number, fallback: number): number {
  if (!Number.isFinite(value)) return fallback;
  if (value < STAMP_MIN_SIZE) return STAMP_MIN_SIZE;
  if (value > STAMP_MAX_SIZE) return STAMP_MAX_SIZE;
  return value;
}

/**
 * Clamp a stamp box so it stays fully on the page (mirrors
 * `clampSignatureBox`): top-left clamped into [0,100], size into the size
 * bounds, then shrunk so the box never spills past the right / bottom edge.
 */
export function clampStampBox(box: {
  x: number;
  y: number;
  width: number;
  height: number;
}): { x: number; y: number; width: number; height: number } {
  const x = clampPercent(box.x);
  const y = clampPercent(box.y);
  let width = clampSize(box.width, DEFAULT_STAMP_WIDTH);
  let height = clampSize(box.height, DEFAULT_STAMP_HEIGHT);
  if (x + width > 100) width = Math.max(STAMP_MIN_SIZE, 100 - x);
  if (y + height > 100) height = Math.max(STAMP_MIN_SIZE, 100 - y);
  return { x, y, width, height };
}

/**
 * The safe CSS color for a stamp annotation's text/border: the preset's
 * palette color when resolvable, else a neutral fallback. The extra
 * `isSafeHexColor` check is belt-and-braces — palette entries are static hex
 * literals — so a future palette edit can't accidentally introduce a
 * non-literal value into inline styles.
 */
export function stampAnnotationColor(
  stamp: Pick<StampAnnotation, "preset">,
): string {
  const preset = resolveStampPreset(stamp.preset);
  if (preset && isSafeHexColor(preset.color)) return preset.color;
  return "#b91c1c";
}

/**
 * Validate + normalize a persisted `pdf_stamps` payload into a usable
 * annotation, or return null when it is structurally invalid. Mirrors
 * `sanitizeSignaturePayload`: out-of-range positions are rejected rather than
 * snapped; box size clamped; a preset stamp needs a known palette key, an
 * image stamp a safe PNG/JPEG data URL (via the signature image guard); a
 * payload with neither usable content is rejected.
 */
export function sanitizeStampPayload(
  id: string,
  payload: Record<string, unknown> | null | undefined,
): StampAnnotation | null {
  if (!id || typeof id !== "string") return null;
  if (!payload || typeof payload !== "object") return null;
  const record = payload as Record<string, unknown>;
  const page = Number(record.page);
  if (!Number.isInteger(page) || page < 1 || page > 10000) return null;
  const { x, y } = record;
  if (!isFinitePercent(x) || !isFinitePercent(y)) return null;

  const rawKind = record.kind;
  const preset = resolveStampPreset(record.preset);
  const image = sanitizeSignatureImage(record.image);
  let kind: StampAnnotationKind;
  if (rawKind === "preset") kind = "preset";
  else if (rawKind === "image") kind = "image";
  else kind = preset ? "preset" : "image";

  if (kind === "preset" && !preset) return null;
  if (kind === "image" && !image) return null;

  const box = clampStampBox({
    x,
    y,
    width: Number(record.width),
    height: Number(record.height),
  });
  return {
    id,
    page,
    x: box.x,
    y: box.y,
    width: box.width,
    height: box.height,
    kind,
    // Only carry the content for the resolved kind (like the signature
    // sanitizer) so one payload can't smuggle both renderings.
    preset: kind === "preset" ? (preset as StampPreset).key : "",
    image: kind === "image" ? (image as string) : "",
  };
}

/**
 * Build the payload to POST for a new stamp. Position/size clamped, the
 * preset forced onto the palette (preset kind) and the image through the
 * signature guard (image kind) — persisted bytes match what we render.
 * Returns null when the input has no usable content for its kind.
 */
export function buildStampPayload(input: {
  page: number;
  x: number;
  y: number;
  width?: number;
  height?: number;
  kind: StampAnnotationKind;
  preset?: string;
  image?: string;
}): {
  page: number;
  x: number;
  y: number;
  width: number;
  height: number;
  kind: StampAnnotationKind;
  preset: string;
  image: string;
} | null {
  const box = clampStampBox({
    x: input.x,
    y: input.y,
    width: input.width ?? DEFAULT_STAMP_WIDTH,
    height: input.height ?? DEFAULT_STAMP_HEIGHT,
  });
  if (input.kind === "preset") {
    const preset = resolveStampPreset(input.preset);
    if (!preset) return null;
    return {
      page: input.page,
      x: box.x,
      y: box.y,
      width: box.width,
      height: box.height,
      kind: "preset",
      preset: preset.key,
      image: "",
    };
  }
  const image = sanitizeSignatureImage(input.image);
  if (!image) return null;
  return {
    page: input.page,
    x: box.x,
    y: box.y,
    width: box.width,
    height: box.height,
    kind: "image",
    preset: "",
    image,
  };
}

/**
 * Cap the number of stamps kept for a single page to
 * {@link STAMP_MAX_PER_PAGE}, dropping the OLDEST first. Returns a new array;
 * input untouched. Mirrors `capSignaturesPerPage`.
 */
export function capStampsPerPage<T>(
  stamps: ReadonlyArray<T>,
  max = STAMP_MAX_PER_PAGE,
): T[] {
  if (!Array.isArray(stamps)) return [];
  const cap = Number.isInteger(max) && max > 0 ? max : STAMP_MAX_PER_PAGE;
  if (stamps.length <= cap) return stamps.slice();
  return stamps.slice(stamps.length - cap);
}
