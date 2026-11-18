/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  PDF_DECORATION_MODEL_VERSION,
  pdfDecorationPreviewEntries,
  sanitizePdfDocumentDecorations,
  type PdfDecorationFont,
  type PdfDecorationPreviewEntry,
  type PdfDocumentDecorations,
} from "./document-decorations";
import {
  MAX_MANAGED_NATIVE_DECORATION_LAYOUT_CHARS,
  MAX_MANAGED_NATIVE_STAMP_PATCHES,
  planManagedNativeStampAnnotations,
  type ManagedNativeStampRecord,
} from "./native-stamp-roundtrip";
import { sanitizeSignatureImage } from "@mrnote/core/web/lib/pdf/signatures.ts";

export const PDF_NATIVE_DECORATION_ID_PREFIX = "pdf-decoration:";











export const MAX_NATIVE_DECORATION_PAGES = 9_999;
export const MAX_NATIVE_DECORATION_CANVAS_DIMENSION = 1600;
const PDF_NATIVE_DECORATION_ID_PATTERN =
  /^pdf-decoration:[a-z0-9][a-z0-9._~-]{0,79}:page:([1-9][0-9]{0,3})$/i;

export interface PdfNativeDecorationPageSize {
  readonly width: number;
  readonly height: number;
  readonly rotation?: 0 | 90 | 180 | 270;
}

interface PdfDecorationCanvasContext {
  globalAlpha: number;
  fillStyle: string | CanvasGradient | CanvasPattern;
  font: string;
  textAlign: CanvasTextAlign;
  textBaseline: CanvasTextBaseline;
  save(): void;
  restore(): void;
  translate(x: number, y: number): void;
  rotate(angle: number): void;
  fillText(text: string, x: number, y: number, maxWidth?: number): void;
  drawImage(
    image: CanvasImageSource,
    dx: number,
    dy: number,
    dWidth: number,
    dHeight: number,
  ): void;
}

export interface PdfDecorationCanvas {
  width: number;
  height: number;
  getContext(contextId: "2d"): PdfDecorationCanvasContext | null;
  toDataURL(type?: string): string;
}

export type PdfDecorationCanvasFactory = (
  width: number,
  height: number,
) => PdfDecorationCanvas | null;

export type PdfDecorationImageResolver = (
  dataItemId: string,
) => Promise<CanvasImageSource | null>;

export type PdfNativeDecorationFailureReason =
  | "legacy_flatten_only"
  | "invalid_page_count"
  | "page_limit"
  | "invalid_page_geometry"
  | "invalid_page_rotation"
  | "duplicate_existing_id"
  | "reserved_id_collision"
  | "native_stamp_plan_refused"
  /** More pages to decorate than one write carries (with the page's other stamps). */
  | "stamp_limit"
  /** The decorated pages together (an image watermark's layers) are too large for one write. */
  | "total_size_limit"
  | "native_stamp_coverage_mismatch"
  | "canvas_unavailable"
  | "image_resolver_required"
  | "image_resolution_failed"
  | "render_failed"
  | "invalid_png"
  | "layout_too_large";






export type PdfNativeDecorationRenderResult =
  | Readonly<{
      supported: true;
      /** Base64 JSON {@link PdfNativeDecorationLayout}. */
      layout: string;
      /** Base64 PNG for an image watermark, or "" when the page has none. */
      raster: string;
    }>
  | Readonly<{
      supported: false;
      reason: PdfNativeDecorationFailureReason;
    }>;

export interface PdfNativeDecorationLayoutItem {
  readonly role: "header" | "footer" | "watermark";
  readonly text: string;
  /** Anchor point in rendered-page points, origin top-left, y down. */
  readonly x: number;
  readonly y: number;
  readonly size: number;
  readonly font: PdfDecorationFont;
  readonly bold: boolean;
  /** RRGGBB without `#`. */
  readonly color: string;
  readonly opacity: number;
  /** Counter-clockwise degrees about the anchor, as the preview draws it. */
  readonly angle: number;
  readonly align: "left" | "center" | "right";
  readonly baseline: "top" | "middle" | "bottom";
}

export interface PdfNativeDecorationLayout {
  readonly v: 1;
  /** Rendered page size in points (after the page's /Rotate). */
  readonly w: number;
  readonly h: number;
  readonly items: readonly PdfNativeDecorationLayoutItem[];
}

export interface RenderPdfDecorationNativeStampInput {
  readonly decorations: PdfDocumentDecorations;
  readonly page: number;
  readonly pageCount: number;
  readonly pageWidth: number;
  readonly pageHeight: number;
  readonly rotation?: 0 | 90 | 180 | 270;
  readonly resolveImage?: PdfDecorationImageResolver;
  readonly canvasFactory?: PdfDecorationCanvasFactory;
  /**
   * Image-watermark layers already painted in this reconciliation, by what
   * they show: pages of one size carry the same layer, painted once (a
   * canvas and a PNG encoding per page held the main thread for the whole
   * document).
   */
  readonly rasterLayers?: Map<string, Promise<PdfNativeDecorationRasterLayer>>;
}

type PdfNativeDecorationRasterLayer =
  | { supported: true; raster: string }
  | { supported: false; reason: PdfNativeDecorationFailureReason };

export type PdfDecorationPageRenderer = (
  input: RenderPdfDecorationNativeStampInput,
) => Promise<PdfNativeDecorationRenderResult>;

export type PdfNativeDecorationReconciliationResult =
  | Readonly<{
      supported: true;
      records: readonly ManagedNativeStampRecord[];
    }>
  | Readonly<{
      supported: false;
      reason: PdfNativeDecorationFailureReason;
      /** Exact pre-call state; callers must never persist a partial plan. */
      records: readonly ManagedNativeStampRecord[];
    }>;

export interface ReconcilePdfDecorationNativeStampsInput {
  readonly decorations: PdfDocumentDecorations;
  /** Original-PDF page order; array index + 1 is the native page number. */
  readonly pageSizes: readonly PdfNativeDecorationPageSize[];
  readonly currentRecords: readonly ManagedNativeStampRecord[];
  readonly resolveImage?: PdfDecorationImageResolver;
  /** Focused test seam; production callers use the bounded browser renderer. */
  readonly renderPage?: PdfDecorationPageRenderer;
}

function defaultCanvasFactory(
  width: number,
  height: number,
): PdfDecorationCanvas | null {
  if (typeof document === "undefined") return null;
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  return canvas as PdfDecorationCanvas;
}

function finitePositive(value: number): boolean {
  return Number.isFinite(value) && value > 0 && value <= 14_400;
}

/**
 * Pixel size of the image-watermark raster layer: up to 2x (144 DPI) and
 * never above {@link MAX_NATIVE_DECORATION_CANVAS_DIMENSION} on either side.
 */
function canvasGeometry(
  pageWidth: number,
  pageHeight: number,
): { width: number; height: number; scale: number } | null {
  if (!finitePositive(pageWidth) || !finitePositive(pageHeight)) return null;
  const scale = Math.min(
    2,
    MAX_NATIVE_DECORATION_CANVAS_DIMENSION / Math.max(pageWidth, pageHeight),
  );
  const width = Math.max(1, Math.round(pageWidth * scale));
  const height = Math.max(1, Math.round(pageHeight * scale));
  if (
    width > MAX_NATIVE_DECORATION_CANVAS_DIMENSION ||
    height > MAX_NATIVE_DECORATION_CANVAS_DIMENSION
  ) {
    return null;
  }
  return { width, height, scale };
}

function pageRotation(value: unknown): 0 | 90 | 180 | 270 | null {
  return value === undefined || value === 0
    ? 0
    : value === 90 || value === 180 || value === 270
      ? value
      : null;
}

function renderedPageGeometry(
  width: number,
  height: number,
  rotation: 0 | 90 | 180 | 270,
): { width: number; height: number } {
  return rotation === 90 || rotation === 270
    ? { width: height, height: width }
    : { width, height };
}

function layoutNumber(value: number): number {
  return Math.round(value * 1000) / 1000;
}

// The server writes one visible line per item; a stray control character in
// a template becomes the space the preview canvas draws for it.
const LAYOUT_CONTROL_CHARACTERS = /[\u0000-\u001f\u007f-\u009f\u2028\u2029]/g;

function layoutColor(color: string): string {
  return /^#[0-9a-f]{6}$/i.test(color) ? color.slice(1).toUpperCase() : "000000";
}

/**
 * The same geometry the decoration preview draws, as points on the rendered
 * page: a watermark is centred on its anchor and turned by its angle; a slot
 * hangs from the top margin (header) or sits on the bottom margin (footer).
 */
function layoutItem(
  entry: PdfDecorationPreviewEntry,
  width: number,
  height: number,
): PdfNativeDecorationLayoutItem | null {
  const text = entry.text.replace(LAYOUT_CONTROL_CHARACTERS, " ");
  if (!text.trim()) return null;
  if (entry.kind === "watermark") {
    return {
      role: "watermark",
      text,
      x: layoutNumber((entry.x ?? 0.5) * width),
      y: layoutNumber((entry.y ?? 0.5) * height),
      size: layoutNumber(
        Math.max(1, (entry.fontSize / 100) * Math.min(width, height)),
      ),
      font: entry.font,
      bold: true,
      color: layoutColor(entry.color),
      opacity: entry.opacity,
      angle: layoutNumber(Number(entry.angle ?? 0)),
      align: "center",
      baseline: "middle",
    };
  }
  const fallbackMargin = Math.max(0, Number(entry.margin ?? 24));
  const horizontalMargin = Math.max(
    0,
    Number(entry.horizontalMargin ?? fallbackMargin),
  );
  const verticalMargin = Math.max(
    0,
    Number(entry.verticalMargin ?? fallbackMargin),
  );
  const horizontal = entry.horizontal ?? "center";
  const vertical = entry.vertical ?? "footer";
  return {
    role: vertical,
    text,
    x: layoutNumber(
      horizontal === "left"
        ? horizontalMargin
        : horizontal === "right"
          ? width - horizontalMargin
          : width / 2,
    ),
    y: layoutNumber(
      vertical === "header" ? verticalMargin : height - verticalMargin,
    ),
    size: layoutNumber(Math.max(1, entry.fontSize)),
    font: entry.font,
    bold: false,
    color: layoutColor(entry.color),
    opacity: entry.opacity,
    angle: 0,
    align: horizontal,
    baseline: vertical === "header" ? "top" : "bottom",
  };
}

function base64Utf8(value: string): string {
  const bytes = new TextEncoder().encode(value);
  let binary = "";
  for (let index = 0; index < bytes.length; index += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(index, index + 0x8000));
  }
  return btoa(binary);
}

export function decodePdfDecorationNativeLayout(
  layout: string,
): PdfNativeDecorationLayout | null {
  try {
    const binary = atob(layout);
    const bytes = Uint8Array.from(binary, (character) =>
      character.charCodeAt(0),
    );
    const parsed = JSON.parse(new TextDecoder().decode(bytes)) as unknown;
    if (
      !parsed ||
      typeof parsed !== "object" ||
      (parsed as { v?: unknown }).v !== 1 ||
      !Array.isArray((parsed as { items?: unknown }).items)
    ) {
      return null;
    }
    return parsed as PdfNativeDecorationLayout;
  } catch {
    return null;
  }
}

async function drawWatermarkImage(
  context: PdfDecorationCanvasContext,
  entry: PdfDecorationPreviewEntry,
  width: number,
  height: number,
  resolveImage: PdfDecorationImageResolver | undefined,
): Promise<PdfNativeDecorationFailureReason | null> {
  if (!entry.dataItemId || !resolveImage) return "image_resolver_required";
  let image: CanvasImageSource | null;
  try {
    image = await resolveImage(entry.dataItemId);
  } catch {
    return "image_resolution_failed";
  }
  if (!image) return "image_resolution_failed";
  const imageWidth = entry.imageWidthFraction
    ? entry.imageWidthFraction * width
    : Math.min(width, height) * (entry.scaleFraction ?? 0.35);
  const imageHeight = entry.imageHeightFraction
    ? entry.imageHeightFraction * height
    : imageWidth / Math.max(0.02, entry.imageAspectRatio ?? 1);
  if (
    !Number.isFinite(imageWidth) ||
    !Number.isFinite(imageHeight) ||
    imageWidth <= 0 ||
    imageHeight <= 0
  ) {
    return "render_failed";
  }
  try {
    context.save();
    context.globalAlpha = entry.opacity;
    context.translate((entry.x ?? 0.5) * width, (entry.y ?? 0.5) * height);
    context.rotate((-Number(entry.angle ?? 0) * Math.PI) / 180);
    context.drawImage(
      image,
      -imageWidth / 2,
      -imageHeight / 2,
      imageWidth,
      imageHeight,
    );
    context.restore();
  } catch {
    return "render_failed";
  }
  return null;
}

/** Paint the image-watermark entries of one rendered page as a PNG layer. */
async function renderImageWatermarkLayer(
  entries: readonly PdfDecorationPreviewEntry[],
  renderedWidth: number,
  renderedHeight: number,
  input: RenderPdfDecorationNativeStampInput,
): Promise<
  | { supported: true; raster: string }
  | { supported: false; reason: PdfNativeDecorationFailureReason }
> {
  const geometry = canvasGeometry(renderedWidth, renderedHeight);
  if (!geometry) return { supported: false, reason: "invalid_page_geometry" };
  const canvas = (input.canvasFactory ?? defaultCanvasFactory)(
    geometry.width,
    geometry.height,
  );
  const context = canvas?.getContext("2d") ?? null;
  if (!canvas || !context) {
    return { supported: false, reason: "canvas_unavailable" };
  }
  try {
    for (const entry of entries) {
      const failure = await drawWatermarkImage(
        context,
        entry,
        geometry.width,
        geometry.height,
        input.resolveImage,
      );
      if (failure) return { supported: false, reason: failure };
    }
  } catch {
    return { supported: false, reason: "render_failed" };
  }
  let encoded: string;
  try {
    encoded = canvas.toDataURL("image/png");
  } catch {
    return { supported: false, reason: "invalid_png" };
  }
  const image = sanitizeSignatureImage(encoded);
  const prefix = "data:image/png;base64,";
  if (!image || !image.startsWith(prefix)) {
    return { supported: false, reason: "invalid_png" };
  }
  return { supported: true, raster: image.slice(prefix.length) };
}

export async function renderPdfDecorationNativeStamp(
  input: RenderPdfDecorationNativeStampInput,
): Promise<PdfNativeDecorationRenderResult> {
  const model = sanitizePdfDocumentDecorations(input.decorations);
  if (model.version !== PDF_DECORATION_MODEL_VERSION) {
    return { supported: false, reason: "legacy_flatten_only" };
  }
  if (
    !Number.isInteger(input.pageCount) ||
    input.pageCount < 1 ||
    !Number.isInteger(input.page) ||
    input.page < 1 ||
    input.page > input.pageCount
  ) {
    return { supported: false, reason: "invalid_page_count" };
  }
  if (input.pageCount > MAX_NATIVE_DECORATION_PAGES) {
    return { supported: false, reason: "page_limit" };
  }
  const rotation = pageRotation(input.rotation);
  if (rotation === null) {
    return { supported: false, reason: "invalid_page_rotation" };
  }
  if (!finitePositive(input.pageWidth) || !finitePositive(input.pageHeight)) {
    return { supported: false, reason: "invalid_page_geometry" };
  }
  const renderedPage = renderedPageGeometry(
    input.pageWidth,
    input.pageHeight,
    rotation,
  );
  const entries = pdfDecorationPreviewEntries(
    model,
    input.page,
    input.pageCount,
    renderedPage.width,
    renderedPage.height,
  );
  const items: PdfNativeDecorationLayoutItem[] = [];
  for (const entry of entries) {
    if (entry.kind === "watermark-image") continue;
    const item = layoutItem(entry, renderedPage.width, renderedPage.height);
    if (item) items.push(item);
  }
  const imageEntries = entries.filter(
    (entry) => entry.kind === "watermark-image",
  );
  let raster = "";
  if (imageEntries.length > 0) {
    const paint = () =>
      renderImageWatermarkLayer(
        imageEntries,
        renderedPage.width,
        renderedPage.height,
        input,
      );
    let painted: Promise<PdfNativeDecorationRasterLayer>;
    if (input.rasterLayers) {
      const key = JSON.stringify([
        renderedPage.width,
        renderedPage.height,
        imageEntries,
      ]);
      painted = input.rasterLayers.get(key) ?? paint();
      input.rasterLayers.set(key, painted);
    } else {
      painted = paint();
    }
    const layer = await painted;
    if (!layer.supported) return layer;
    raster = layer.raster;
  }
  const layout = base64Utf8(
    JSON.stringify({
      v: 1,
      w: layoutNumber(renderedPage.width),
      h: layoutNumber(renderedPage.height),
      items,
    } satisfies PdfNativeDecorationLayout),
  );
  if (layout.length > MAX_MANAGED_NATIVE_DECORATION_LAYOUT_CHARS) {
    return { supported: false, reason: "layout_too_large" };
  }
  return { supported: true, layout, raster };
}

export function pdfDecorationNativeStampId(
  modelId: string,
  page: number,
): string {
  return `${PDF_NATIVE_DECORATION_ID_PREFIX}${modelId}:page:${page}`;
}

/**
 * A live decoration record: the page-content layout this version writes, or
 * the full-page raster stamp older versions wrote (still a faithful native
 * copy until the decoration is next applied, which replaces it).
 */
function isManagedDecorationRecordShape(
  record: ManagedNativeStampRecord,
): boolean {
  if (record.deleted) return false;
  if (record.kind === "decoration") {
    return typeof record.layout === "string" && record.layout.length > 0;
  }
  return (
    record.kind === "image" &&
    record.preset === "" &&
    record.x === 0 &&
    record.y === 0 &&
    record.width === 100 &&
    record.height === 100 &&
    Boolean(
      sanitizeSignatureImage(record.image)?.startsWith(
        "data:image/png;base64,",
      ),
    )
  );
}

/** The pages (one-based) the decorations draw anything on. */
function decoratedPageNumbers(
  model: PdfDocumentDecorations,
  pageSizes: readonly PdfNativeDecorationPageSize[],
): number[] {
  const pages: number[] = [];
  const pageCount = pageSizes.length;
  for (let index = 0; index < pageCount; index += 1) {
    const size = pageSizes[index];
    const renderedPage = renderedPageGeometry(
      size.width,
      size.height,
      pageRotation(size.rotation) ?? 0,
    );
    if (
      pdfDecorationPreviewEntries(
        model,
        index + 1,
        pageCount,
        renderedPage.width,
        renderedPage.height,
      ).length > 0
    ) {
      pages.push(index + 1);
    }
  }
  return pages;
}

/**
 * Whether one write carries these decorations: a decorated page is one
 * stamp record, and a page's stamp records -- the reader's own stamps and
 * the records of pages no longer decorated included -- go to the PDF as one
 * plan of at most {@link MAX_MANAGED_NATIVE_STAMP_PATCHES}. `pages` is how
 * many pages the decorations draw on, `max` how many this page still takes.
 * Answered from the model and the page sizes alone, before anything is laid
 * out or sent, so a refusal names its reason at once. Null when the page
 * geometry is not readable (the reconciliation names that).
 */
export function pdfDecorationNativeStampBudget(input: {
  readonly decorations: PdfDocumentDecorations;
  readonly pageSizes: readonly PdfNativeDecorationPageSize[];
  readonly currentRecords: readonly ManagedNativeStampRecord[];
}): Readonly<{ pages: number; max: number; within: boolean }> | null {
  const model = sanitizePdfDocumentDecorations(input.decorations);
  if (
    input.pageSizes.length < 1 ||
    input.pageSizes.some(
      ({ width, height, rotation }) =>
        !finitePositive(width) ||
        !finitePositive(height) ||
        pageRotation(rotation) === null,
    )
  ) {
    return null;
  }
  const decorated = new Set(
    decoratedPageNumbers(model, input.pageSizes).map((page) =>
      pdfDecorationNativeStampId(model.id, page),
    ),
  );
  let others = 0;
  for (const record of input.currentRecords) {
    if (!decorated.has(record.id)) others += 1;
  }
  const max = Math.max(0, MAX_MANAGED_NATIVE_STAMP_PATCHES - others);
  return { pages: decorated.size, max, within: decorated.size <= max };
}

export function hasExactPdfDecorationNativeStampCoverage(input: {
  readonly decorations: PdfDocumentDecorations;
  readonly pageSizes: readonly PdfNativeDecorationPageSize[];
  readonly records: readonly ManagedNativeStampRecord[];
}): boolean {
  const model = sanitizePdfDocumentDecorations(input.decorations);
  if (
    model.version !== PDF_DECORATION_MODEL_VERSION ||
    input.pageSizes.length < 1 ||
    input.pageSizes.length > MAX_NATIVE_DECORATION_PAGES ||
    input.pageSizes.some(
      ({ width, height, rotation }) =>
        !finitePositive(width) ||
        !finitePositive(height) ||
        pageRotation(rotation) === null,
    )
  ) {
    return false;
  }

  const expected = new Set<string>(
    decoratedPageNumbers(model, input.pageSizes).map((page) =>
      pdfDecorationNativeStampId(model.id, page),
    ),
  );

  const active = new Set<string>();
  for (const record of input.records) {
    if (!record.id.startsWith(PDF_NATIVE_DECORATION_ID_PREFIX)) continue;
    if (record.deleted) continue;
    const match = PDF_NATIVE_DECORATION_ID_PATTERN.exec(record.id);
    const page = Number(match?.[1]);
    if (
      !match ||
      active.has(record.id) ||
      record.page !== page ||
      !isManagedDecorationRecordShape(record)
    ) {
      return false;
    }
    active.add(record.id);
  }
  if (active.size !== expected.size) return false;
  return [...expected].every((id) => active.has(id));
}

function unchangedFailure(
  currentRecords: readonly ManagedNativeStampRecord[],
  reason: PdfNativeDecorationFailureReason,
): PdfNativeDecorationReconciliationResult {
  return { supported: false, reason, records: currentRecords };
}

export async function reconcilePdfDecorationNativeStamps(
  input: ReconcilePdfDecorationNativeStampsInput,
): Promise<PdfNativeDecorationReconciliationResult> {
  const model = sanitizePdfDocumentDecorations(input.decorations);
  if (model.version !== PDF_DECORATION_MODEL_VERSION) {
    return unchangedFailure(input.currentRecords, "legacy_flatten_only");
  }
  if (input.pageSizes.length < 1) {
    return unchangedFailure(input.currentRecords, "invalid_page_count");
  }
  if (input.pageSizes.length > MAX_NATIVE_DECORATION_PAGES) {
    return unchangedFailure(input.currentRecords, "page_limit");
  }
  if (
    input.pageSizes.some(
      ({ width, height, rotation }) =>
        !finitePositive(width) ||
        !finitePositive(height) ||
        pageRotation(rotation) === null,
    )
  ) {
    return unchangedFailure(
      input.currentRecords,
      input.pageSizes.some(({ rotation }) => pageRotation(rotation) === null)
        ? "invalid_page_rotation"
        : "invalid_page_geometry",
    );
  }
  const currentIds = new Set<string>();
  for (const record of input.currentRecords) {
    if (currentIds.has(record.id)) {
      return unchangedFailure(input.currentRecords, "duplicate_existing_id");
    }
    currentIds.add(record.id);
    if (
      PDF_NATIVE_DECORATION_ID_PATTERN.test(record.id) &&
      !record.deleted &&
      !isManagedDecorationRecordShape(record)
    ) {
      return unchangedFailure(input.currentRecords, "reserved_id_collision");
    }
  }

  // Known before a page is laid out or an image layer painted: a document
  // with more pages to decorate than one write carries is refused here.
  const budget = pdfDecorationNativeStampBudget({
    decorations: model,
    pageSizes: input.pageSizes,
    currentRecords: input.currentRecords,
  });
  if (budget && !budget.within) {
    return unchangedFailure(input.currentRecords, "stamp_limit");
  }

  const pageCount = input.pageSizes.length;
  const renderPage = input.renderPage ?? renderPdfDecorationNativeStamp;
  const rasterLayers = new Map<
    string,
    Promise<PdfNativeDecorationRasterLayer>
  >();
  const desired = new Map<string, ManagedNativeStampRecord>();
  for (let index = 0; index < pageCount; index += 1) {
    const page = index + 1;
    const size = input.pageSizes[index];
    const rotation = pageRotation(size.rotation) ?? 0;
    const renderedPage = renderedPageGeometry(
      size.width,
      size.height,
      rotation,
    );
    if (
      pdfDecorationPreviewEntries(
        model,
        page,
        pageCount,
        renderedPage.width,
        renderedPage.height,
      ).length === 0
    ) {
      continue;
    }
    const rendered = await renderPage({
      decorations: model,
      page,
      pageCount,
      pageWidth: size.width,
      pageHeight: size.height,
      rotation,
      resolveImage: input.resolveImage,
      rasterLayers,
    });
    if (!rendered.supported) {
      return unchangedFailure(input.currentRecords, rendered.reason);
    }
    if (!rendered.layout) {
      return unchangedFailure(input.currentRecords, "render_failed");
    }
    const id = pdfDecorationNativeStampId(model.id, page);
    desired.set(id, {
      id,
      page,
      kind: "decoration",
      layout: rendered.layout,
      raster: rendered.raster,
    });
  }

  const records: ManagedNativeStampRecord[] = [];
  const consumed = new Set<string>();
  for (const current of input.currentRecords) {
    if (!PDF_NATIVE_DECORATION_ID_PATTERN.test(current.id)) {
      records.push(current);
      continue;
    }
    const replacement = desired.get(current.id);
    if (replacement) {
      records.push(replacement);
      consumed.add(current.id);
    } else if (current.deleted) {
      records.push(current);
    } else {
      records.push({ id: current.id, page: current.page, deleted: true });
    }
  }
  for (const [id, record] of desired) {
    if (!consumed.has(id)) records.push(record);
  }
  const plan = planManagedNativeStampAnnotations(records);
  if (!plan.supported) {
    return unchangedFailure(
      input.currentRecords,
      plan.reason === "patch_limit"
        ? "stamp_limit"
        : plan.reason === "image_stamp_total_limit"
          ? "total_size_limit"
          : "native_stamp_plan_refused",
    );
  }
  if (
    !hasExactPdfDecorationNativeStampCoverage({
      decorations: model,
      pageSizes: input.pageSizes,
      records,
    })
  ) {
    return unchangedFailure(
      input.currentRecords,
      "native_stamp_coverage_mismatch",
    );
  }
  return { supported: true, records };
}
