/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  watermarkPlacementFractions,
  type FlattenNumbering,
  type FlattenStampConfig,
  type FlattenWatermark,
} from "@mrnote/core/web/lib/pdf/flatten.ts";
import { resolveStampTokens } from "./stamps";

export const PDF_DECORATION_LEGACY_MODEL_VERSION = 1 as const;
export const PDF_DECORATION_MODEL_VERSION = 2 as const;
export const PDF_DECORATION_DEFAULT_MODEL_ID = "pdf-document-decorations";
export const PDF_DECORATION_NATIVE_DISABLED_REASON =
  "Managed PDF decorations are preserved by MRNote and baked into flattened output; native third-party decoration objects are not rewritten yet.";
export const PDF_WATERMARK_IMAGE_DISABLED_REASON =
  "Image watermarks require an editable notebook PDF with authenticated PNG/JPEG attachment storage, preview access, and export byte access.";

export type PdfDecorationKind = "watermark" | "page-number" | "header-footer";
export type PdfDecorationFont = "sans" | "serif" | "mono";
export type PdfDecorationHorizontal = "left" | "center" | "right";
export type PdfDecorationVertical = "header" | "footer";
export type PdfWatermarkVertical = "top" | "center" | "bottom";
export type PdfWatermarkRepeatRows = 1 | 2 | 3 | 6 | 8 | 12;
export type PdfDecorationRangeMode =
  "all" | "odd" | "even" | "current-onward" | "current-only" | "custom";

export interface PdfDecorationPageRange {
  mode: PdfDecorationRangeMode;
  /** 1-based page captured when a current-page range is committed. */
  anchorPage?: number;
  /** Comma/range grammar, for example `1,3,5-9`. */
  spec?: string;
}

interface PdfManagedDecorationBase {
  id: string;
  appliedAt: string;
  range: PdfDecorationPageRange;
  nativeDisposition: "flatten-only" | "managed-native";
  disabledReason?: string;
}

export interface PdfWatermarkDecoration extends PdfManagedDecorationBase {
  kind: "watermark";
  sourceKind: "text" | "image";
  /** Retained while switching source modes; rendered only for text source. */
  text: string;
  /** Durable authenticated attachment authority; never a blob/data/http URL. */
  image: PdfWatermarkImageSource | null;
  angle: number;
  opacity: number;
  color: string;
  font: PdfDecorationFont;
  /** Relative to the shorter page edge. */
  fontFraction: number;
  /** Raster width relative to the shorter page edge. */
  scaleFraction: number;
  repeatRows: PdfWatermarkRepeatRows;
  horizontal: PdfDecorationHorizontal;
  vertical: PdfWatermarkVertical;
  /** Page-relative offsets from the selected anchor, bounded to [-0.45, 0.45]. */
  offsetX: number;
  offsetY: number;
}

export interface PdfWatermarkImageSource {
  dataItemId: string;
  filename: string;
  mimeType: "image/png" | "image/jpeg";
  width: number;
  height: number;
}

export type PdfPageNumberFormat =
  | "number"
  | "page-of-total"
  | "number-of-total"
  | "page-label"
  | "page-total-label"
  // WPS's Chinese page-number styles: 第1/X页, 第1页, 第1页，共X页.
  | "zh-page-of-total"
  | "zh-page-label"
  | "zh-page-total-label";

const PDF_PAGE_NUMBER_FORMATS: ReadonlySet<string> = new Set([
  "number",
  "page-of-total",
  "number-of-total",
  "page-label",
  "page-total-label",
  "zh-page-of-total",
  "zh-page-label",
  "zh-page-total-label",
]);

export interface PdfPageNumberDecoration extends PdfManagedDecorationBase {
  kind: "page-number";
  format: PdfPageNumberFormat;
  start: number;
  vertical: PdfDecorationVertical;
  horizontal: PdfDecorationHorizontal;
  font: PdfDecorationFont;
  fontSize: number;
  color: string;
  /** Distance (pt) from the top (header) or bottom (footer) page edge. */
  margin: number;
  /**
   * Distance (pt) from the left or right page edge for a left/right number
   * (WPS's left/right page-number margins). Older records omit it and keep
   * using `margin` on both axes.
   */
  horizontalMargin?: number;
  





  template?: string;
}

/** The longest page-number wording a template may hold. */
export const PDF_PAGE_NUMBER_TEMPLATE_MAX = 60;

/** A template that names the page's number, within its bound. */
export function isPdfPageNumberTemplate(value: unknown): value is string {
  return typeof value === "string" && value.length <= PDF_PAGE_NUMBER_TEMPLATE_MAX &&
    value.includes("{page}") && !/[\u0000-\u001f]/u.test(value);
}

export interface PdfHeaderFooterSlots {
  header: Record<PdfDecorationHorizontal, string>;
  footer: Record<PdfDecorationHorizontal, string>;
}

export interface PdfHeaderFooterMargins {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

export interface PdfHeaderFooterDecoration extends PdfManagedDecorationBase {
  kind: "header-footer";
  slots: PdfHeaderFooterSlots;
  font: PdfDecorationFont;
  fontSize: number;
  color: string;
  margins: PdfHeaderFooterMargins;
}

export interface PdfDocumentDecorations {
  id: string;
  version:
    | typeof PDF_DECORATION_LEGACY_MODEL_VERSION
    | typeof PDF_DECORATION_MODEL_VERSION;
  watermark: PdfWatermarkDecoration | null;
  pageNumber: PdfPageNumberDecoration | null;
  headerFooter: PdfHeaderFooterDecoration | null;
}

export const EMPTY_PDF_DOCUMENT_DECORATIONS: PdfDocumentDecorations =
  Object.freeze({
    id: PDF_DECORATION_DEFAULT_MODEL_ID,
    version: PDF_DECORATION_MODEL_VERSION,
    watermark: null,
    pageNumber: null,
    headerFooter: null,
  });

const HEX_COLOR = /^#[0-9a-f]{6}$/i;
const MODEL_ID = /^[a-z0-9][a-z0-9._~-]{0,79}$/i;
const MAX_DECORATION_TEXT = 240;
const MAX_RANGE_SPEC = 200;
const DATA_ITEM_ID = /^[a-z0-9][a-z0-9._:-]{0,127}$/i;
const WATERMARK_REPEAT_ROWS = new Set<number>([1, 2, 3, 6, 8, 12]);

function finite(
  value: unknown,
  fallback: number,
  min: number,
  max: number,
): number {
  const parsed = Number(value);
  return Number.isFinite(parsed)
    ? Math.min(max, Math.max(min, parsed))
    : fallback;
}

function text(value: unknown, max = MAX_DECORATION_TEXT): string {
  return typeof value === "string"
    ? value
        .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, "")
        .slice(0, max)
    : "";
}

function modelId(value: unknown): string {
  const candidate = text(value, 80).trim();
  return MODEL_ID.test(candidate) ? candidate : PDF_DECORATION_DEFAULT_MODEL_ID;
}

function color(value: unknown, fallback: string): string {
  return typeof value === "string" && HEX_COLOR.test(value)
    ? value.toLowerCase()
    : fallback;
}

function font(value: unknown): PdfDecorationFont {
  return value === "serif" || value === "mono" ? value : "sans";
}

function range(value: unknown): PdfDecorationPageRange {
  const source =
    value && typeof value === "object"
      ? (value as Record<string, unknown>)
      : {};
  const mode: PdfDecorationRangeMode =
    source.mode === "odd" ||
    source.mode === "even" ||
    source.mode === "current-onward" ||
    source.mode === "current-only" ||
    source.mode === "custom"
      ? source.mode
      : "all";
  const anchor = Math.trunc(finite(source.anchorPage, 1, 1, 100_000));
  return {
    mode,
    ...(mode === "current-onward" || mode === "current-only"
      ? { anchorPage: anchor }
      : {}),
    ...(mode === "custom"
      ? { spec: text(source.spec, MAX_RANGE_SPEC).trim() }
      : {}),
  };
}

function base(
  source: Record<string, unknown>,
  id: string,
  nativeDisposition: PdfManagedDecorationBase["nativeDisposition"],
): PdfManagedDecorationBase {
  const appliedAt =
    typeof source.appliedAt === "string" &&
    !Number.isNaN(Date.parse(source.appliedAt))
      ? source.appliedAt
      : new Date(0).toISOString();
  return {
    id: text(source.id, 80).trim() || id,
    appliedAt,
    range: range(source.range),
    nativeDisposition,
    ...(nativeDisposition === "flatten-only"
      ? { disabledReason: PDF_DECORATION_NATIVE_DISABLED_REASON }
      : {}),
  };
}

function sanitizeWatermarkImage(value: unknown): PdfWatermarkImageSource | null {
  if (!value || typeof value !== "object") return null;
  const source = value as Record<string, unknown>;
  const dataItemId = text(source.dataItemId ?? source.data_item_id, 128).trim();
  const mimeType =
    source.mimeType === "image/png" || source.mime_type === "image/png"
      ? "image/png"
      : source.mimeType === "image/jpeg" || source.mime_type === "image/jpeg"
        ? "image/jpeg"
        : null;
  if (!DATA_ITEM_ID.test(dataItemId) || !mimeType) return null;
  return {
    dataItemId,
    filename: text(source.filename, 200).trim() || "watermark-image",
    mimeType,
    width: Math.round(finite(source.width, 1, 1, 100_000)),
    height: Math.round(finite(source.height, 1, 1, 100_000)),
  };
}

function watermarkRepeatRows(value: unknown): PdfWatermarkRepeatRows {
  const rows = Math.trunc(Number(value));
  return WATERMARK_REPEAT_ROWS.has(rows)
    ? (rows as PdfWatermarkRepeatRows)
    : 1;
}

function sanitizeWatermark(
  value: unknown,
  nativeDisposition: PdfManagedDecorationBase["nativeDisposition"],
): PdfWatermarkDecoration | null {
  if (!value || typeof value !== "object") return null;
  const source = value as Record<string, unknown>;
  const content = text(source.text).trim();
  const image = sanitizeWatermarkImage(source.image);
  const sourceKind = source.sourceKind === "image" ? "image" : "text";
  if ((sourceKind === "text" && !content) || (sourceKind === "image" && !image))
    return null;
  return {
    ...base(source, "pdf-watermark", nativeDisposition),
    kind: "watermark",
    sourceKind,
    text: content,
    image,
    angle: finite(source.angle, 45, -90, 90),
    opacity: finite(source.opacity, 0.5, 0.02, 1),
    color: color(source.color, "#64748b"),
    font: font(source.font),
    fontFraction: finite(source.fontFraction, 0.12, 0.02, 0.5),
    scaleFraction: finite(source.scaleFraction, 0.35, 0.08, 0.9),
    repeatRows: watermarkRepeatRows(source.repeatRows),
    horizontal:
      source.horizontal === "left" || source.horizontal === "right"
        ? source.horizontal
        : "center",
    vertical:
      source.vertical === "top" || source.vertical === "bottom"
        ? source.vertical
        : "center",
    offsetX: finite(source.offsetX, 0, -0.45, 0.45),
    offsetY: finite(source.offsetY, 0, -0.45, 0.45),
  };
}

function sanitizePageNumber(
  value: unknown,
  nativeDisposition: PdfManagedDecorationBase["nativeDisposition"],
): PdfPageNumberDecoration | null {
  if (!value || typeof value !== "object") return null;
  const source = value as Record<string, unknown>;
  const format: PdfPageNumberFormat =
    typeof source.format === "string" &&
    PDF_PAGE_NUMBER_FORMATS.has(source.format)
      ? (source.format as PdfPageNumberFormat)
      : "number";
  const margin = finite(source.margin, 24, 4, 144);
  return {
    ...base(source, "pdf-page-number", nativeDisposition),
    kind: "page-number",
    format,
    start: Math.trunc(finite(source.start, 1, 0, 1_000_000_000)),
    vertical: source.vertical === "header" ? "header" : "footer",
    horizontal:
      source.horizontal === "left" || source.horizontal === "right"
        ? source.horizontal
        : "center",
    font: font(source.font),
    fontSize: finite(source.fontSize, 10, 4, 72),
    color: color(source.color, "#334155"),
    margin,
    ...(source.horizontalMargin !== undefined
      ? { horizontalMargin: finite(source.horizontalMargin, margin, 4, 144) }
      : {}),
    ...(isPdfPageNumberTemplate(source.template) ? { template: source.template } : {}),
  };
}

function emptySlots(): PdfHeaderFooterSlots {
  return {
    header: { left: "", center: "", right: "" },
    footer: { left: "", center: "", right: "" },
  };
}

function sanitizeSlots(value: unknown): PdfHeaderFooterSlots {
  const output = emptySlots();
  if (!value || typeof value !== "object") return output;
  for (const vertical of ["header", "footer"] as const) {
    const row = (value as Record<string, unknown>)[vertical];
    if (!row || typeof row !== "object") continue;
    for (const horizontal of ["left", "center", "right"] as const) {
      output[vertical][horizontal] = text(
        (row as Record<string, unknown>)[horizontal],
      );
    }
  }
  return output;
}

function sanitizeHeaderFooterMargins(
  source: Record<string, unknown>,
): PdfHeaderFooterMargins {
  const legacyMargin =
    source.margin === undefined
      ? null
      : finite(source.margin, 24, 4, 144);
  const margins =
    source.margins && typeof source.margins === "object"
      ? (source.margins as Record<string, unknown>)
      : {};
  return {
    top: finite(margins.top, legacyMargin ?? 36, 4, 144),
    bottom: finite(margins.bottom, legacyMargin ?? 36, 4, 144),
    left: finite(margins.left, legacyMargin ?? 72, 4, 144),
    right: finite(margins.right, legacyMargin ?? 72, 4, 144),
  };
}

function sanitizeHeaderFooter(
  value: unknown,
  nativeDisposition: PdfManagedDecorationBase["nativeDisposition"],
): PdfHeaderFooterDecoration | null {
  if (!value || typeof value !== "object") return null;
  const source = value as Record<string, unknown>;
  const slots = sanitizeSlots(source.slots);
  if (
    !Object.values(slots.header).some((value) => value.trim()) &&
    !Object.values(slots.footer).some((value) => value.trim())
  )
    return null;
  return {
    ...base(source, "pdf-header-footer", nativeDisposition),
    kind: "header-footer",
    slots,
    font: font(source.font),
    fontSize: finite(source.fontSize, 8, 4, 72),
    color: color(source.color, "#334155"),
    margins: sanitizeHeaderFooterMargins(source),
  };
}

export function sanitizePdfDocumentDecorations(
  value: unknown,
): PdfDocumentDecorations {
  const source =
    value && typeof value === "object"
      ? (value as Record<string, unknown>)
      : {};
  const hasLegacyDecorationPayload = Boolean(
    source.watermark ||
      source.pageNumber ||
      source.page_number ||
      source.headerFooter ||
      source.header_footer,
  );
  const hasManagedNativeDecorationPayload = [
    source.watermark,
    source.pageNumber ?? source.page_number,
    source.headerFooter ?? source.header_footer,
  ].some(
    (decoration) =>
      decoration &&
      typeof decoration === "object" &&
      (decoration as Record<string, unknown>).nativeDisposition ===
        "managed-native",
  );
  const version =
    source.version === PDF_DECORATION_MODEL_VERSION ||
    (source.version === undefined && hasManagedNativeDecorationPayload)
      ? PDF_DECORATION_MODEL_VERSION
      : source.version === PDF_DECORATION_LEGACY_MODEL_VERSION ||
          hasLegacyDecorationPayload
        ? PDF_DECORATION_LEGACY_MODEL_VERSION
        : PDF_DECORATION_MODEL_VERSION;
  const nativeDisposition =
    version === PDF_DECORATION_MODEL_VERSION
      ? ("managed-native" as const)
      : ("flatten-only" as const);
  return {
    id: modelId(source.id ?? source.modelId ?? source.model_id),
    version,
    watermark: sanitizeWatermark(source.watermark, nativeDisposition),
    pageNumber: sanitizePageNumber(
      source.pageNumber ?? source.page_number,
      nativeDisposition,
    ),
    headerFooter: sanitizeHeaderFooter(
      source.headerFooter ?? source.header_footer,
      nativeDisposition,
    ),
  };
}

export function hasPdfDocumentDecorations(
  value: PdfDocumentDecorations,
): boolean {
  return Boolean(value.watermark || value.pageNumber || value.headerFooter);
}

export function createDefaultPdfDecoration(
  kind: PdfDecorationKind,
  currentPage: number,
  now = new Date(),
):
  PdfWatermarkDecoration | PdfPageNumberDecoration | PdfHeaderFooterDecoration {
  const common = {
    id: `pdf-${kind}`,
    appliedAt: now.toISOString(),
    range: { mode: "all" as const },
    nativeDisposition: "managed-native" as const,
  };
  if (kind === "watermark") {
    return {
      ...common,
      kind,
      sourceKind: "text",
      text: "",
      image: null,
      angle: 45,
      opacity: 0.5,
      color: "#64748b",
      font: "sans",
      fontFraction: 0.12,
      scaleFraction: 0.35,
      repeatRows: 1,
      horizontal: "center",
      vertical: "center",
      offsetX: 0,
      offsetY: 0,
    };
  }
  if (kind === "page-number") {
    return {
      ...common,
      kind,
      format: "number",
      start: 1,
      vertical: "footer",
      horizontal: "center",
      font: "sans",
      fontSize: 10,
      color: "#334155",
      margin: 24,
    };
  }
  return {
    ...common,
    kind,
    slots: emptySlots(),
    font: "sans",
    fontSize: 8,
    color: "#334155",
    margins: { top: 36, bottom: 36, left: 72, right: 72 },
    range: { mode: "all" },
  };
}

/** Returns null for invalid custom grammar; otherwise a sorted, de-duplicated set. */
export function parsePdfDecorationRangeSpec(
  spec: string,
  pageCount: number,
): number[] | null {
  const count = Math.max(0, Math.trunc(pageCount));
  const input = spec.trim();
  if (!input || count < 1) return null;
  const pages = new Set<number>();
  for (const raw of input.split(",")) {
    const part = raw.trim();
    const match = /^(\d+)(?:\s*-\s*(\d+))?$/.exec(part);
    if (!match) return null;
    const start = Number(match[1]);
    const end = Number(match[2] ?? match[1]);
    if (
      !Number.isSafeInteger(start) ||
      !Number.isSafeInteger(end) ||
      start < 1 ||
      end < start ||
      end > count
    )
      return null;
    for (let page = start; page <= end; page += 1) pages.add(page);
  }
  return [...pages].sort((a, b) => a - b);
}

export function pdfDecorationRangePages(
  rangeValue: PdfDecorationPageRange,
  pageCount: number,
): number[] {
  const count = Math.max(0, Math.trunc(pageCount));
  if (rangeValue.mode === "custom")
    return parsePdfDecorationRangeSpec(rangeValue.spec ?? "", count) ?? [];
  const anchor = Math.min(
    count,
    Math.max(1, Math.trunc(rangeValue.anchorPage ?? 1)),
  );
  const pages: number[] = [];
  for (let page = 1; page <= count; page += 1) {
    if (rangeValue.mode === "odd" && page % 2 === 0) continue;
    if (rangeValue.mode === "even" && page % 2 !== 0) continue;
    if (rangeValue.mode === "current-onward" && page < anchor) continue;
    if (rangeValue.mode === "current-only" && page !== anchor) continue;
    pages.push(page);
  }
  return pages;
}

function corner(
  vertical: PdfDecorationVertical,
  horizontal: PdfDecorationHorizontal,
): FlattenNumbering["corner"] {
  return `${vertical === "header" ? "top" : "bottom"}-${horizontal}`;
}

function pageNumberLabel(
  config: PdfPageNumberDecoration,
  value: number,
  total: number,
): string {
  if (config.template && isPdfPageNumberTemplate(config.template)) {
    return config.template.replaceAll("{pages}", String(total)).replaceAll("{page}", String(value));
  }
  if (config.format === "page-of-total") return `Page ${value} / ${total}`;
  if (config.format === "number-of-total") return `${value} / ${total}`;
  if (config.format === "page-label") return `Page ${value}`;
  if (config.format === "page-total-label")
    return `Page ${value}, ${total} pages`;
  if (config.format === "zh-page-of-total") return `第${value}/${total}页`;
  if (config.format === "zh-page-label") return `第${value}页`;
  if (config.format === "zh-page-total-label")
    return `第${value}页，共${total}页`;
  return String(value);
}

function resolvedSlot(
  template: string,
  page: number,
  total: number,
  appliedAt: string,
): string {
  return resolveStampTokens(template.replace(/\{pages\}/g, "{total}"), {
    number: page,
    pageNumber: page,
    totalPages: total,
    now: new Date(appliedAt),
  });
}

function flattenWatermark(
  watermark: PdfWatermarkDecoration,
  pageNumbers: readonly number[],
): FlattenWatermark {
  return {
    text: watermark.sourceKind === "text" ? watermark.text : "",
    imageDataItemId:
      watermark.sourceKind === "image"
        ? watermark.image?.dataItemId
        : undefined,
    imageAspectRatio:
      watermark.sourceKind === "image" && watermark.image
        ? watermark.image.width / watermark.image.height
        : undefined,
    angle: watermark.angle,
    opacity: watermark.opacity,
    color: watermark.color,
    fontFraction: watermark.fontFraction,
    scaleFraction: watermark.scaleFraction,
    repeatRows: watermark.repeatRows,
    horizontal: watermark.horizontal,
    vertical: watermark.vertical,
    offsetX: watermark.offsetX,
    offsetY: watermark.offsetY,
    font: watermark.font,
    pageNumbers,
  };
}

export function buildPdfDecorationFlattenConfig(
  value: PdfDocumentDecorations,
  sourcePageOrder: readonly number[],
): FlattenStampConfig | null {
  const model = sanitizePdfDocumentDecorations(value);
  if (!hasPdfDocumentDecorations(model)) return null;
  const order = sourcePageOrder.filter(
    (page) => Number.isInteger(page) && page >= 1,
  );
  const total = order.length;
  const sourcePageCount = order.reduce(
    (maximum, page) => Math.max(maximum, page),
    0,
  );
  const watermarks: FlattenWatermark[] = [];
  const numberings: FlattenNumbering[] = [];
  if (model.watermark) {
    watermarks.push(
      flattenWatermark(
        model.watermark,
        pdfDecorationRangePages(model.watermark.range, sourcePageCount),
      ),
    );
  }
  if (model.pageNumber) {
    const targetPages = new Set(
      pdfDecorationRangePages(model.pageNumber.range, sourcePageCount),
    );
    let sequence = model.pageNumber.start;
    const labels = order.map((sourcePage) => {
      if (!targetPages.has(sourcePage)) return "";
      const label = pageNumberLabel(
        model.pageNumber as PdfPageNumberDecoration,
        sequence,
        total,
      );
      sequence += 1;
      return label;
    });
    numberings.push({
      labels,
      corner: corner(model.pageNumber.vertical, model.pageNumber.horizontal),
      fontSize: model.pageNumber.fontSize,
      color: model.pageNumber.color,
      margin: model.pageNumber.margin,
      ...(model.pageNumber.horizontalMargin !== undefined
        ? {
            horizontalMargin:
              model.pageNumber.horizontal === "center"
                ? 0
                : model.pageNumber.horizontalMargin,
            verticalMargin: model.pageNumber.margin,
          }
        : {}),
      font: model.pageNumber.font,
    });
  }
  if (model.headerFooter) {
    const targetPages = new Set(
      pdfDecorationRangePages(model.headerFooter.range, sourcePageCount),
    );
    for (const vertical of ["header", "footer"] as const) {
      for (const horizontal of ["left", "center", "right"] as const) {
        const template = model.headerFooter.slots[vertical][horizontal].trim();
        if (!template) continue;
        numberings.push({
          labels: order.map((sourcePage, index) =>
            targetPages.has(sourcePage)
              ? resolvedSlot(
                  template,
                  index + 1,
                  total,
                  model.headerFooter!.appliedAt,
                )
              : "",
          ),
          corner: corner(vertical, horizontal),
          fontSize: model.headerFooter.fontSize,
          color: model.headerFooter.color,
          margin:
            vertical === "header"
              ? model.headerFooter.margins.top
              : model.headerFooter.margins.bottom,
          horizontalMargin:
            horizontal === "left"
              ? model.headerFooter.margins.left
              : horizontal === "right"
                ? model.headerFooter.margins.right
                : 0,
          verticalMargin:
            vertical === "header"
              ? model.headerFooter.margins.top
              : model.headerFooter.margins.bottom,
          font: model.headerFooter.font,
        });
      }
    }
  }
  return { watermarks, numberings };
}

export interface PdfDecorationPreviewEntry {
  key: string;
  text: string;
  kind: "watermark" | "watermark-image" | "slot";
  dataItemId?: string;
  imageAspectRatio?: number;
  scaleFraction?: number;
  imageWidthFraction?: number;
  imageHeightFraction?: number;
  x?: number;
  y?: number;
  vertical?: PdfDecorationVertical;
  horizontal?: PdfDecorationHorizontal;
  margin?: number;
  horizontalMargin?: number;
  verticalMargin?: number;
  angle?: number;
  opacity: number;
  color: string;
  font: PdfDecorationFont;
  fontSize: number;
}

export function pdfDecorationPreviewEntries(
  value: PdfDocumentDecorations,
  page: number,
  pageCount: number,
  pageWidth = 612,
  pageHeight = 792,
): PdfDecorationPreviewEntry[] {
  const model = sanitizePdfDocumentDecorations(value);
  const entries: PdfDecorationPreviewEntry[] = [];
  if (
    model.watermark &&
    pdfDecorationRangePages(model.watermark.range, pageCount).includes(page)
  ) {
    const watermark = flattenWatermark(model.watermark, [page]);
    const placements = watermarkPlacementFractions(
      watermark,
      pageWidth,
      pageHeight,
    );
    placements.forEach((placement, index) => {
      entries.push({
        key: `${model.watermark!.id}-${index}`,
        text:
          model.watermark!.sourceKind === "text"
            ? model.watermark!.text
            : "",
        kind:
          model.watermark!.sourceKind === "image"
            ? "watermark-image"
            : "watermark",
        dataItemId: model.watermark!.image?.dataItemId,
        imageAspectRatio:
          model.watermark!.image && model.watermark!.image.height > 0
            ? model.watermark!.image.width / model.watermark!.image.height
            : undefined,
        scaleFraction: placement.scaleFraction,
        imageWidthFraction: placement.imageWidthFraction,
        imageHeightFraction: placement.imageHeightFraction,
        x: placement.x,
        y: placement.y,
        angle: model.watermark!.angle,
        opacity: model.watermark!.opacity,
        color: model.watermark!.color,
        font: model.watermark!.font,
        fontSize: placement.fontFraction * 100,
      });
    });
  }
  if (
    model.pageNumber &&
    pdfDecorationRangePages(model.pageNumber.range, pageCount).includes(page)
  ) {
    const target = pdfDecorationRangePages(model.pageNumber.range, pageCount);
    const valueNumber =
      model.pageNumber.start + Math.max(0, target.indexOf(page));
    entries.push({
      key: model.pageNumber.id,
      text: pageNumberLabel(model.pageNumber, valueNumber, pageCount),
      kind: "slot",
      vertical: model.pageNumber.vertical,
      horizontal: model.pageNumber.horizontal,
      margin: model.pageNumber.margin,
      ...(model.pageNumber.horizontalMargin !== undefined
        ? {
            horizontalMargin: model.pageNumber.horizontalMargin,
            verticalMargin: model.pageNumber.margin,
          }
        : {}),
      opacity: 1,
      color: model.pageNumber.color,
      font: model.pageNumber.font,
      fontSize: model.pageNumber.fontSize,
    });
  }
  if (
    model.headerFooter &&
    pdfDecorationRangePages(model.headerFooter.range, pageCount).includes(page)
  ) {
    for (const vertical of ["header", "footer"] as const)
      for (const horizontal of ["left", "center", "right"] as const) {
        const template = model.headerFooter.slots[vertical][horizontal];
        if (!template.trim()) continue;
        entries.push({
          key: `${model.headerFooter.id}-${vertical}-${horizontal}`,
          text: resolvedSlot(
            template,
            page,
            pageCount,
            model.headerFooter.appliedAt,
          ),
          kind: "slot",
          vertical,
          horizontal,
          horizontalMargin:
            horizontal === "left"
              ? model.headerFooter.margins.left
              : horizontal === "right"
                ? model.headerFooter.margins.right
                : 0,
          verticalMargin:
            vertical === "header"
              ? model.headerFooter.margins.top
              : model.headerFooter.margins.bottom,
          opacity: 1,
          color: model.headerFooter.color,
          font: model.headerFooter.font,
          fontSize: model.headerFooter.fontSize,
        });
      }
  }
  return entries;
}
