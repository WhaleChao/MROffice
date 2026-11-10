/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PDFPageProxy } from "pdfjs-dist";
import { createPdfRenderOperationsFilter } from "./pdfjs-render-operator-filter";
import type { TextBoxEditLike } from "./text-box-flatten";
import type { PdfTextBoxRect } from "@mrnote/core/web/lib/pdf/text-box-rect.ts";
import {
  PDF_TEXT_RASTER_MAX_DIMENSION,
  PDF_TEXT_RASTER_MAX_PIXELS,
} from "./text-run-layout";

export type PdfTextInkBounds = Readonly<PdfTextBoxRect>;

/** Appearance-local ink, not the wrapping container or a source/writer proof. */
export function pdfTextInkBoundsFromRgba(
  data: Uint8ClampedArray,
  width: number,
  height: number,
): PdfTextInkBounds | null {
  if (!validPdfTextSurfaceSize(width, height) || data.length !== width * height * 4)
    return null;
  let left = width,
    top = height,
    right = -1,
    bottom = -1;
  for (let y = 0; y < height; y += 1)
    for (let x = 0; x < width; x += 1) {
      if (data[(y * width + x) * 4 + 3] === 0) continue;
      left = Math.min(left, x);
      right = Math.max(right, x);
      top = Math.min(top, y);
      bottom = Math.max(bottom, y);
    }
  return right < left
    ? null
    : Object.freeze({
        x: left / width,
        y: top / height,
        width: (right + 1 - left) / width,
        height: (bottom + 1 - top) / height,
      });
}









export function pdfTextInkColorFromRgba(
  data: Uint8ClampedArray,
  width: number,
  height: number,
): string | null {
  if (!validPdfTextSurfaceSize(width, height) || data.length !== width * height * 4) return null;
  const counts = new Map<number, number>();
  const quantize = (value: number) => value & 0xf8;
  for (let index = 0; index < data.length; index += 4) {
    if (data[index + 3] < 128) continue;
    const key = (quantize(data[index]) << 16) | (quantize(data[index + 1]) << 8) | quantize(data[index + 2]);
    counts.set(key, (counts.get(key) ?? 0) + 1);
  }
  if (counts.size < 2) return null;
  const ranked = [...counts.entries()].sort((left, right) => right[1] - left[1]);
  const background = ranked[0][0];
  const distance = (key: number) =>
    Math.abs((key >> 16) - (background >> 16)) +
    Math.abs(((key >> 8) & 0xff) - ((background >> 8) & 0xff)) +
    Math.abs((key & 0xff) - (background & 0xff));
  let ink: [number, number] | null = null;
  let weight = 0;
  for (const entry of ranked.slice(1)) {
    const [key, count] = entry;
    const far = distance(key);
    if (far < 96 || count < 4) continue;
    if (count * far * far > weight) { weight = count * far * far; ink = entry; }
  }
  if (!ink) return null;
  const hex = (value: number) => (value | 0x04).toString(16).padStart(2, "0");
  return `#${hex(ink[0] >> 16)}${hex((ink[0] >> 8) & 0xff)}${hex(ink[0] & 0xff)}`;
}

/** `pdfTextInkColorFromRgba` over a canvas region given in CSS fractions of
 * the canvas's displayed box. Null when the canvas cannot be read. */
export function samplePdfTextInkColor(
  canvas: HTMLCanvasElement,
  region: Readonly<{ x: number; y: number; width: number; height: number }>,
): string | null {
  const x = Math.max(0, Math.floor(region.x * canvas.width));
  const y = Math.max(0, Math.floor(region.y * canvas.height));
  const width = Math.min(canvas.width - x, Math.ceil(region.width * canvas.width));
  const height = Math.min(canvas.height - y, Math.ceil(region.height * canvas.height));
  if (!(width > 0 && height > 0) || width * height > 4_000_000) return null;
  try {
    const context = canvas.getContext("2d", { willReadFrequently: true });
    const image = context?.getImageData(x, y, width, height);
    return image ? pdfTextInkColorFromRgba(image.data, width, height) : null;
  } catch {
    return null;
  }
}

export function validPdfTextSurfaceSize(width: number, height: number): boolean {
  return (
    Number.isSafeInteger(width) &&
    Number.isSafeInteger(height) &&
    width > 0 &&
    height > 0 &&
    width <= PDF_TEXT_RASTER_MAX_DIMENSION &&
    height <= PDF_TEXT_RASTER_MAX_DIMENSION &&
    width * height <= PDF_TEXT_RASTER_MAX_PIXELS
  );
}

/** Crop around the original rotation centre, not the smaller ink rectangle. */
export function pdfTextInkDisplayRect(
  box: PdfTextBoxRect,
  ink: PdfTextInkBounds,
  angle: number,
  pageWidth: number,
  pageHeight: number,
): PdfTextBoxRect {
  const radians = (angle * Math.PI) / 180;
  const dx = (ink.x + ink.width / 2 - 0.5) * box.width * pageWidth;
  const dy = (ink.y + ink.height / 2 - 0.5) * box.height * pageHeight;
  const width = box.width * ink.width,
    height = box.height * ink.height;
  return {
    x:
      box.x +
      box.width / 2 +
      (dx * Math.cos(radians) - dy * Math.sin(radians)) / pageWidth -
      width / 2,
    y:
      box.y +
      box.height / 2 +
      (dx * Math.sin(radians) + dy * Math.cos(radians)) / pageHeight -
      height / 2,
    width,
    height,
  };
}

/** UI-only association. Full text and the exact current viewport rectangle must
 * agree uniquely; never infer a source locator or a writer from this result. */
export function matchPdfTextInkAnnotation(
  edit: TextBoxEditLike,
  annotations: readonly unknown[],
  viewport: Pick<
    ReturnType<PDFPageProxy["getViewport"]>,
    "width" | "height" | "convertToViewportRectangle"
  >,
): string | null {
  const { x, y, width, height, text } = edit;
  if (
    [x, y, width, height, edit.angle ?? 0].some(
      (n) => typeof n !== "number" || !Number.isFinite(n),
    ) ||
    typeof text !== "string" ||
    !text.trim() ||
    !width ||
    !height ||
    width < 0 ||
    height < 0
  )
    return null;
  const angle = (Number(edit.angle ?? 0) * Math.PI) / 180;
  const w = width * viewport.width,
    h = height * viewport.height;
  const halfW =
    (Math.abs(w * Math.cos(angle)) + Math.abs(h * Math.sin(angle))) / 2;
  const halfH =
    (Math.abs(w * Math.sin(angle)) + Math.abs(h * Math.cos(angle))) / 2;
  const cx = (Number(x) + width / 2) * viewport.width;
  const cy = (Number(y) + height / 2) * viewport.height;
  const expected = [cx - halfW, cy - halfH, cx + halfW, cy + halfH];
  const matches = annotations.filter((value): value is { id: string } => {
    if (!value || typeof value !== "object") return false;
    const a = value as Record<string, unknown>;
    if (
      a.subtype !== "FreeText" ||
      a.hasAppearance !== true ||
      a.hasOwnCanvas ||
      typeof a.id !== "string" ||
      !/^[1-9]\d*R(?:\d+)?$/.test(a.id) ||
      (a.contentsObj as { str?: unknown } | undefined)?.str !== text ||
      !Array.isArray(a.rect) ||
      a.rect.length !== 4 ||
      !a.rect.every((n) => typeof n === "number" && Number.isFinite(n))
    )
      return false;
    const raw = viewport.convertToViewportRectangle(a.rect);
    const actual = [
      Math.min(raw[0], raw[2]),
      Math.min(raw[1], raw[3]),
      Math.max(raw[0], raw[2]),
      Math.max(raw[1], raw[3]),
    ];
    return actual.every((n, i) => Math.abs(n - expected[i]) <= 0.02);
  });
  return matches.length === 1 ? matches[0].id : null;
}

/** Render only the matched annotation's text/strokes from the CURRENT PDF.
 * The server writes portable vector text, so the draft PNG is not an oracle.
 * Reuses the loaded page/operator cache and the caller's single annotation pass.
 * No download, package mutation, new authority or main-thread PDF parsing. */
export async function inspectPdfNativeTextInk(input: {
  page: PDFPageProxy;
  annotations: readonly unknown[];
  edits: readonly TextBoxEditLike[];
  rotation: number;
  ops: Readonly<Record<string, number>>;
  signal: AbortSignal;
  optionalContentConfigPromise?: Parameters<
    PDFPageProxy["render"]
  >[0]["optionalContentConfigPromise"];
}): Promise<ReadonlyMap<TextBoxEditLike, PdfTextInkBounds>> {
  const { page, edits, annotations, rotation, ops, signal } = input;
  const result = new Map<TextBoxEditLike, PdfTextInkBounds>();
  if (signal.aborted || edits.length > 64 || annotations.length > 4_000)
    return result;
  const viewport = page.getViewport({ scale: 1, rotation });
  const targets = edits.map((edit) => ({
    edit,
    id: matchPdfTextInkAnnotation(edit, annotations, viewport),
  }));
  if (!targets.some((target) => target.id)) return result;
  const operators = await page.getOperatorList();
  if (signal.aborted || operators.fnArray.length > 200_000) return result;
  const spans = new Map<string, Array<[number, number]>>();
  let open: { id: string; start: number } | null = null;
  for (let i = 0; i < operators.fnArray.length; i += 1) {
    if (operators.fnArray[i] === ops.beginAnnotation) {
      if (open) return result;
      const args = operators.argsArray[i];
      if (
        !Array.isArray(args) ||
        args.length !== 5 ||
        typeof args[0] !== "string" ||
        typeof args[4] !== "boolean"
      )
        return result;
      open = { id: args[0], start: i };
    } else if (operators.fnArray[i] === ops.endAnnotation) {
      if (!open) return result;
      spans.set(open.id, [...(spans.get(open.id) ?? []), [open.start, i]]);
      open = null;
    }
  }
  if (open) return result;
  for (const { edit, id } of targets) {
    if (signal.aborted) break;
    if (
      !id ||
      targets.filter((target) => target.id === id).length !== 1 ||
      spans.get(id)?.length !== 1
    )
      continue;
    const [start, end] = spans.get(id)![0];
    const w = Number(edit.width) * viewport.width,
      h = Number(edit.height) * viewport.height;
    const density = 2;
    const width = Math.ceil(w * density),
      height = Math.ceil(h * density);
    if (!validPdfTextSurfaceSize(width, height)) continue;
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d", { willReadFrequently: true });
    if (!context) continue;
    const radians = (Number(edit.angle ?? 0) * Math.PI) / 180;
    const cos = Math.cos(radians),
      sin = Math.sin(radians);
    const cx = (Number(edit.x) + Number(edit.width) / 2) * viewport.width;
    const cy = (Number(edit.y) + Number(edit.height) / 2) * viewport.height;
    let task: ReturnType<PDFPageProxy["render"]> | undefined;
    let pixels: Uint8ClampedArray | undefined;
    const cancel = () => task?.cancel();
    const timer = setTimeout(cancel, 3_000);
    signal.addEventListener("abort", cancel, { once: true });
    try {
      // (Indices of getOperatorList's list; render runs pdf.js's folded one.)
      const inkFilter = createPdfRenderOperationsFilter({ list: operators, ops, keep: (index) => {
          const op = operators.fnArray[index];
          if (
            [
              ops.dependency,
              ops.beginMarkedContent,
              ops.beginMarkedContentProps,
              ops.endMarkedContent,
            ].includes(op)
          )
            return true;
          if (index < start || index > end) return false;
          // Skip the box fill, not glyphs. Keep underline strokes and clipping.
          if (op === ops.constructPath) {
            const paint = operators.argsArray[index]?.[0];
            return (
              paint === ops.stroke ||
              paint === ops.closeStroke ||
              paint === ops.clip ||
              paint === ops.eoClip
            );
          }
          return ![
            ops.fill,
            ops.eoFill,
            ops.fillStroke,
            ops.eoFillStroke,
            ops.closeFillStroke,
            ops.closeEOFillStroke,
          ].includes(op);
        } });
      task = page.render({
        canvas,
        canvasContext: context,
        viewport,
        background: "rgba(0,0,0,0)",
        optionalContentConfigPromise: input.optionalContentConfigPromise,
        transform: [
          density * cos,
          -density * sin,
          density * sin,
          density * cos,
          density * (w / 2 - cos * cx - sin * cy),
          density * (h / 2 + sin * cx - cos * cy),
        ],
        operationsFilter: inkFilter.filter,
      });
      inkFilter.bind(task);
      await task.promise;
      if (signal.aborted) break;
      if (!inkFilter.finish()) continue;
      pixels = context.getImageData(0, 0, width, height).data;
      const ink = pdfTextInkBoundsFromRgba(pixels, width, height);
      if (ink)
        result.set(
          edit,
          Object.freeze({
            x: (ink.x * width) / (w * density),
            y: (ink.y * height) / (h * density),
            width: (ink.width * width) / (w * density),
            height: (ink.height * height) / (h * density),
          }),
        );
    } catch {
      /* An unavailable appearance leaves the existing layout intact. */
    } finally {
      clearTimeout(timer);
      signal.removeEventListener("abort", cancel);
      pixels?.fill(0);
      canvas.width = 0;
      canvas.height = 0;
    }
  }
  return result;
}

/** Read-only editing background for an already saved, source-bound text box.
 * Reuse the loaded page, excluding one complete annotation scope. In particular,
 * never change the viewer's annotationStorage or hide neighbouring page ink.
 * The caller still authenticates source ownership independently before saving. */
export async function renderPdfManagedTextEditingSurface(input: {
  page: PDFPageProxy;
  annotations: readonly unknown[];
  edit: TextBoxEditLike;
  annotationId: string;
  rotation: number;
  ops: Readonly<Record<string, number>>;
  sourceCanvas: HTMLCanvasElement;
  signal: AbortSignal;
  isCurrent: () => boolean;
  optionalContentConfigPromise?: Parameters<PDFPageProxy["render"]>[0]["optionalContentConfigPromise"];
}): Promise<HTMLCanvasElement | null> {
  const { page, sourceCanvas, signal, ops } = input;
  const current = () => !signal.aborted && input.isCurrent();
  if (!current() || !validPdfTextSurfaceSize(sourceCanvas.width, sourceCanvas.height) ||
      input.annotations.length > 4_000 || ![0, 90, 180, 270].includes(input.rotation)) return null;
  const base = page.getViewport({ scale: 1, rotation: input.rotation });
  if (matchPdfTextInkAnnotation(input.edit, input.annotations, base) !== input.annotationId) return null;
  const operators = await page.getOperatorList();
  if (!current() || operators.fnArray.length > 200_000) return null;
  let open: { id: string; start: number } | null = null;
  let span: [number, number] | null = null;
  for (let i = 0; i < operators.fnArray.length; i += 1) {
    if (operators.fnArray[i] === ops.beginAnnotation) {
      const args = operators.argsArray[i];
      if (open || !Array.isArray(args) || args.length !== 5 ||
          typeof args[0] !== "string" || args[4] !== false) return null;
      open = { id: args[0], start: i };
    } else if (operators.fnArray[i] === ops.endAnnotation) {
      if (!open) return null;
      if (open.id === input.annotationId) {
        if (span) return null;
        span = [open.start, i];
      }
      open = null;
    }
  }
  if (open || !span) return null;
  const viewport = page.getViewport({ scale: sourceCanvas.width / base.width, rotation: input.rotation });
  if (!Number.isFinite(viewport.width) || !Number.isFinite(viewport.height) ||
      Math.abs(viewport.height - sourceCanvas.height) > 2) return null;
  const canvas = sourceCanvas.ownerDocument.createElement("canvas");
  canvas.width = sourceCanvas.width;
  canvas.height = sourceCanvas.height;
  for (const property of ["filter", "opacity", "transform", "transformOrigin"] as const)
    canvas.style[property] = sourceCanvas.style[property];
  let task: ReturnType<PDFPageProxy["render"]> | undefined;
  let published = false;
  let timedOut = false;
  const cancel = () => task?.cancel();
  const timer = setTimeout(() => { timedOut = true; cancel(); }, 3_000);
  signal.addEventListener("abort", cancel, { once: true });
  try {
    const context = canvas.getContext("2d");
    if (!context || !current()) return null;
    const [start, end] = span;
    // (Indices of getOperatorList's list; render runs pdf.js's folded one.)
    const spanFilter = createPdfRenderOperationsFilter({ list: operators, ops,
      keep: index => operators.fnArray[index] === ops.dependency || index < start || index > end });
    task = page.render({
      canvas, canvasContext: context, viewport,
      optionalContentConfigPromise: input.optionalContentConfigPromise,
      // Annotation begin/end must be removed together to keep graphics state
      // balanced. Dependencies may also be needed by later neighbouring ink.
      operationsFilter: spanFilter.filter,
    });
    spanFilter.bind(task);
    await task.promise;
    if (!current() || timedOut || !spanFilter.finish()) return null;
    published = true;
    return canvas;
  } catch {
    return null;
  } finally {
    clearTimeout(timer);
    signal.removeEventListener("abort", cancel);
    if (!published) { canvas.width = 0; canvas.height = 0; }
  }
}
