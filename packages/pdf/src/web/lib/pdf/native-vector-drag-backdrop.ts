/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What the page looks like without the vector being dragged.
 *
 * A drag ghost cut from the page canvas leaves the original in place and
 * carries whatever was painted over the object (text on a circle). The
 * client already knows how to write the page without the object -- the
 * delete plan -- so the page is rendered once from those bytes, off
 * screen, at the canvas's own pixel size. Cropping it under the object gives
 * a patch that erases the original; comparing it with the page as shown
 * leaves the ghost only the pixels the object itself contributes.
 */

import type { PdfNativePlanners } from "./remote/pdf-native-planners";
import { cropCanvasRegion } from "./canvas-crop";
import type { FractionRect } from "./embedded-image-edits";
import { traceNativeVectorOutline } from "./native-vector-drag-ghost";
import { createPdfNativeVectorHistoryEndpoint } from "./native-vector-commands";
import { MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION } from "./native-vector-commands";
import { PDFJS_ANNOTATION_MODE_ENABLE_STORAGE } from "./form-appearance-storage";
import { loadCompatiblePdfJs, PDF_JS_DOCUMENT_ASSETS } from "@mrnote/core/web/lib/pdf/pdfjs-runtime.ts";
import type { PdfNativeVectorSelectionRoute } from "./semantic-appearance-inspection";
import type { PdfVectorContentItem } from "./vector-content-inventory";
import type { PdfVectorPageGeometry } from "./vector-capability-reconciliation";

/** Channel difference up to which two renders show the same pixel. */
const SAME_PIXEL_TOLERANCE = 2;
/** The outline runs down the middle of a stroke; look this far outside it. */
const OUTLINE_RIM_PX = 3;
const MAX_FIT_SAMPLES = 20_000;
const FLAT_BACKDROP_VARIANCE = 4;
const LINEAR_FIT_RMS = 12;
/** Half the width up to which an unchanged streak is something painted over the object. */
const THIN_STROKE_PX = 5;

/** What the reader's own page render was given, so the two agree off the object. */
export interface PdfPageRenderMatch {
  /** The rotation the page canvas was rendered with. */
  readonly rotation: number;
  /** The exact viewport scale of the page canvas, when it is known. */
  readonly scale?: number;
  readonly width: number;
  readonly height: number;
  /** The live document's annotation storage (form values, hidden comments). */
  readonly annotationStorage?: Readonly<Record<string, unknown>> | null;
  /** Layers the reader shows or hides, by optional-content group id. */
  readonly layerVisibility?: ReadonlyArray<readonly [id: string, visible: boolean]>;
  readonly createCanvas?: () => HTMLCanvasElement;
}

type AnnotationStorageLike = {
  setValue(key: string, value: Record<string, unknown>): void;
  getValue?(key: string, defaultValue: unknown): unknown;
};

function intersectionOverSmaller(a: FractionRect, b: FractionRect): number {
  const width = Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x);
  const height = Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y);
  if (width <= 0 || height <= 0) return 0;
  const smaller = Math.min(a.width * a.height, b.width * b.height);
  return smaller > 0 ? (width * height) / smaller : 0;
}

/** Which annotation boxes are the object at `bounds` (all in page fractions). */
export function annotationsAtBounds(
  annotations: ReadonlyArray<{ id: string; box: FractionRect }>,
  bounds: FractionRect,
): string[] {
  // A line's own box has no area; give both a hair of it.
  const pad = 0.004;
  const grow = (box: FractionRect): FractionRect => ({
    x: box.x - pad,
    y: box.y - pad,
    width: box.width + 2 * pad,
    height: box.height + 2 * pad,
  });
  const target = grow(bounds);
  const targetArea = target.width * target.height;
  return annotations
    .filter(({ box }) => {
      const candidate = grow(box);
      const area = candidate.width * candidate.height;
      // The same object, not a page-wide one that merely contains it.
      return (
        intersectionOverSmaller(candidate, target) >= 0.6 &&
        area <= targetArea * 4 &&
        targetArea <= area * 4
      );
    })
    .map(({ id }) => id);
}

async function renderMatchedPage(
  bytes: Uint8Array,
  pageNumber: number,
  match: PdfPageRenderMatch,
  hideAnnotationsAt: FractionRect | null,
): Promise<{ canvas: HTMLCanvasElement; hidden: FractionRect | null } | null> {
  if (match.width <= 0 || match.height <= 0) return null;
  const pdfjs = await loadCompatiblePdfJs();
  const task = pdfjs.getDocument({ data: Uint8Array.from(bytes), ...PDF_JS_DOCUMENT_ASSETS });
  const document_ = await task.promise;
  try {
    const page = await document_.getPage(pageNumber);
    const base = page.getViewport({ scale: 1, rotation: match.rotation });
    if (base.width <= 0) return null;
    const viewport = page.getViewport({
      // The canvas's own scale keeps the two renders pixel-aligned; the
      // width ratio is within a pixel of it.
      scale:
        match.scale && Number.isFinite(match.scale) && match.scale > 0
          ? match.scale
          : match.width / base.width,
      rotation: match.rotation,
    });
    const storage = (document_ as { annotationStorage?: AnnotationStorageLike }).annotationStorage;
    if (storage && match.annotationStorage) {
      for (const [id, value] of Object.entries(match.annotationStorage)) {
        if (value && typeof value === "object") {
          storage.setValue(id, { ...(value as Record<string, unknown>) });
        }
      }
    }
    let hidden: FractionRect | null = null;
    if (hideAnnotationsAt) {
      const annotations = (await page.getAnnotations({ intent: "display" })) as Array<{
        id?: unknown;
        rect?: unknown;
      }>;
      const boxes: Array<{ id: string; box: FractionRect }> = [];
      for (const annotation of annotations) {
        const rect = annotation.rect;
        if (typeof annotation.id !== "string" || !Array.isArray(rect) || rect.length < 4) continue;
        const [x1, y1, x2, y2] = viewport.convertToViewportRectangle(
          rect.slice(0, 4).map(Number),
        );
        boxes.push({
          id: annotation.id,
          box: {
            x: Math.min(x1, x2) / viewport.width,
            y: Math.min(y1, y2) / viewport.height,
            width: Math.abs(x2 - x1) / viewport.width,
            height: Math.abs(y2 - y1) / viewport.height,
          },
        });
      }
      const ids = new Set(annotationsAtBounds(boxes, hideAnnotationsAt));
      // Nothing of the page's own painting is the object: no backdrop.
      if (ids.size === 0 || !storage) return null;
      for (const { id, box } of boxes) {
        if (!ids.has(id)) continue;
        // An object default: pdf.js assigns the entry into it (a null one
        // throws once the entry exists, PDF 0018).
        const previous = storage.getValue?.(id, {});
        storage.setValue(id, {
          ...(previous && typeof previous === "object"
            ? (previous as Record<string, unknown>)
            : {}),
          noView: true,
        });
        hidden = hidden
          ? {
              x: Math.min(hidden.x, box.x),
              y: Math.min(hidden.y, box.y),
              width:
                Math.max(hidden.x + hidden.width, box.x + box.width) -
                Math.min(hidden.x, box.x),
              height:
                Math.max(hidden.y + hidden.height, box.y + box.height) -
                Math.min(hidden.y, box.y),
            }
          : box;
      }
    }
    const canvas = (match.createCanvas ?? (() => document.createElement("canvas")))();
    canvas.width = match.width;
    canvas.height = match.height;
    const context = canvas.getContext("2d");
    if (!context) return null;
    // Painted the way the reader paints its pages (widgets and comments
    // from their appearance streams), so only the object differs.
    const parameters = {
      canvas,
      canvasContext: context,
      viewport,
      annotationMode: PDFJS_ANNOTATION_MODE_ENABLE_STORAGE,
    } as Parameters<typeof page.render>[0];
    if (match.layerVisibility && match.layerVisibility.length > 0) {
      try {
        const config = (await document_.getOptionalContentConfig()) as {
          setVisibility?: (id: string, visible: boolean) => void;
        };
        for (const [id, visible] of match.layerVisibility) config.setVisibility?.(id, visible);
        (parameters as { optionalContentConfigPromise?: unknown }).optionalContentConfigPromise =
          Promise.resolve(config);
      } catch {
        // A document without layers renders as it is.
      }
    }
    await page.render(parameters).promise;
    return { canvas, hidden };
  } finally {
    await document_.destroy().catch(() => undefined);
  }
}

/**
 * Any bytes' page rendered off screen the way the reader rendered its own:
 * the building block for a "page without this object" whose bytes the
 * caller prepared itself (an image removal plan, for one).
 */
export async function renderPdfPageLike(
  bytes: Uint8Array,
  pageNumber: number,
  match: PdfPageRenderMatch,
): Promise<HTMLCanvasElement | null> {
  return (await renderMatchedPage(bytes, pageNumber, match, null))?.canvas ?? null;
}

export async function renderNativeVectorAbsentPage(
  input: {
    /** The planners of the page the bytes are shown on. */
    readonly planners: PdfNativePlanners;
    readonly sourceBytes: Uint8Array;
    readonly route: PdfNativeVectorSelectionRoute;
    readonly pageGeometry: PdfVectorPageGeometry;
  } & PdfPageRenderMatch,
): Promise<HTMLCanvasElement | null> {
  const plan = await input.planners.planPdfNativeVectorBackdropDelete({
    sourceBytes: input.sourceBytes,
    route: input.route,
    pageGeometry: input.pageGeometry,
  });
  const candidate = plan.prepared.candidateBytes;
  try {
    return (await renderMatchedPage(candidate, input.route.pageNumber, input, null))?.canvas ?? null;
  } finally {
    candidate.fill(0);
  }
}

/** What the page paints over one object, across the whole page canvas. */
export interface PdfVectorDragAbove {
  readonly width: number;
  readonly height: number;
  /** Its coverage, 0-255, per page pixel. */
  readonly alpha: Uint8ClampedArray;
  /** Its colour times its coverage (premultiplied), RGB per page pixel. */
  readonly premultiplied: Uint8ClampedArray;
}

/**
 * The layer painted over a vector, from the page without it (`absent`) and
 * the page with a black and a white page-wide fill in its place.
 *
 * Over the fills the page shows `A + (1 - a) * F`, with `A` the layer
 * (premultiplied), `a` its coverage and `F` the fill as the vector's own
 * graphics state paints it: `k * g + (1 - k) * B`, `k` that state's opacity,
 * `g` black or white, `B` the page under. Their difference is
 * `(1 - a) * k * 255`; where nothing lies over the vector that is `k * 255`,
 * the value most of the page shows. Then `a` and `A` follow per pixel, with
 * `B` eliminated through the page without the vector. Null when the fills
 * paint nothing one can read (a fully transparent state, a blend mode).
 */
export function solveAboveLayer(
  absent: Uint8ClampedArray,
  black: Uint8ClampedArray,
  white: Uint8ClampedArray,
  width: number,
  height: number,
): PdfVectorDragAbove | null {
  const pixels = width * height;
  if (absent.length < pixels * 4 || black.length < pixels * 4 || white.length < pixels * 4) {
    return null;
  }
  const histogram = new Uint32Array(256);
  let counted = 0;
  const stride = Math.max(1, Math.floor(pixels / 400_000));
  for (let p = 0; p < pixels; p += stride) {
    const i = p * 4;
    const difference = Math.round(
      (white[i] - black[i] + white[i + 1] - black[i + 1] + white[i + 2] - black[i + 2]) / 3,
    );
    if (difference < 8) continue;
    histogram[difference] += 1;
    counted += 1;
  }
  let mode = 0;
  for (let value = 8; value < 256; value += 1) {
    if (histogram[value] > histogram[mode]) mode = value;
  }
  if (mode < 24 || histogram[mode] < counted * 0.2) return null;
  const opacity = mode / 255;
  const alpha = new Uint8ClampedArray(pixels);
  const premultiplied = new Uint8ClampedArray(pixels * 3);
  for (let p = 0; p < pixels; p += 1) {
    const i = p * 4;
    const difference = (white[i] - black[i] + white[i + 1] - black[i + 1] + white[i + 2] - black[i + 2]) / 3;
    const a = Math.min(1, Math.max(0, 1 - difference / mode));
    alpha[p] = Math.round(a * 255);
    if (a === 0) continue;
    for (let c = 0; c < 3; c += 1) {
      premultiplied[p * 3 + c] = (black[i + c] - (1 - opacity) * absent[i + c]) / opacity;
    }
  }
  return { width, height, alpha, premultiplied };
}

/**
 * The layer the page paints over the route's vector, rendered the way the
 * reader rendered the page; null when it cannot be read.
 */
export async function renderNativeVectorAbove(
  input: {
    /** The planners of the page the bytes are shown on. */
    readonly planners: PdfNativePlanners;
    readonly sourceBytes: Uint8Array;
    readonly route: PdfNativeVectorSelectionRoute;
    readonly pageGeometry: PdfVectorPageGeometry;
    readonly absent: HTMLCanvasElement;
  } & PdfPageRenderMatch,
): Promise<PdfVectorDragAbove | null> {
  const fills = await input.planners.planPdfNativeVectorBackdropFills({
    sourceBytes: input.sourceBytes,
    route: input.route,
    pageGeometry: input.pageGeometry,
  });
  try {
    const [black, white] = await Promise.all([
      renderMatchedPage(fills.black, input.route.pageNumber, input, null),
      renderMatchedPage(fills.white, input.route.pageNumber, input, null),
    ]);
    const read = (canvas: HTMLCanvasElement | undefined) =>
      canvas?.getContext("2d", { willReadFrequently: true })?.getImageData(0, 0, canvas.width, canvas.height).data ?? null;
    const width = input.absent.width;
    const height = input.absent.height;
    if (
      !black || !white ||
      black.canvas.width !== width || black.canvas.height !== height ||
      white.canvas.width !== width || white.canvas.height !== height
    ) {
      return null;
    }
    const absent = read(input.absent);
    const blackData = read(black.canvas);
    const whiteData = read(white.canvas);
    return absent && blackData && whiteData
      ? solveAboveLayer(absent, blackData, whiteData, width, height)
      : null;
  } finally {
    fills.black.fill(0);
    fills.white.fill(0);
  }
}

/**
 * The page without the annotation (or form widget) painted at `bounds`, and
 * the box those annotations really occupy; null when nothing the page paints
 * itself sits there (the object is drawn by the reader's own overlay).
 */
export async function renderPageWithoutAnnotationsAt(
  input: {
    readonly sourceBytes: Uint8Array;
    readonly pageNumber: number;
    /** The dragged frame's box, in page fractions as displayed. */
    readonly bounds: FractionRect;
  } & PdfPageRenderMatch,
): Promise<{ canvas: HTMLCanvasElement; hidden: FractionRect | null } | null> {
  return renderMatchedPage(input.sourceBytes, input.pageNumber, input, input.bounds);
}

/**
 * The page after each member of a multi-selection is removed in turn, top
 * of the paint order first: stage `k` is the page without the first `k + 1`
 * of them, so the last stage is the page without any. The removals chain the
 * way the batch delete's do (each member is found again in the bytes the
 * removal before it produced), so where members overlap every stage is the
 * real page, not an estimate.
 */
export async function renderNativeVectorGroupStages(
  input: {
    /** The planners of the page the bytes are shown on. */
    readonly planners: PdfNativePlanners;
    readonly sourceBytes: Uint8Array;
    readonly routes: readonly PdfNativeVectorSelectionRoute[];
    readonly pageGeometry: PdfVectorPageGeometry;
  } & PdfPageRenderMatch,
): Promise<Array<{ canvas: HTMLCanvasElement; route: PdfNativeVectorSelectionRoute }> | null> {
  if (input.routes.length === 0) return null;
  // A later paint sits on top; removing from the end keeps the earlier
  // members' operator spans where their proofs say they are.
  const ordered = [...input.routes].sort(
    (left, right) => right.operatorStart - left.operatorStart,
  );
  const selections = await input.planners.resolvePdfNativeVectorRouteSelections({
    sourceBytes: input.sourceBytes,
    routes: ordered,
    pageGeometry: input.pageGeometry,
  });
  const endpoints = selections.map(createPdfNativeVectorHistoryEndpoint);
  const stageBytes: Uint8Array[] = [];
  let current = input.sourceBytes;
  try {
    for (let index = 0; index < ordered.length; index += 1) {
      const selection = await input.planners.resolvePdfNativeVectorHistorySelection({
        sourceBytes: current,
        endpoint: endpoints[index],
      });
      const plan = await input.planners.planPdfNativeVectorBackdropSelectionDelete({
        sourceBytes: current,
        payload: {
          schemaVersion: MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION,
          selection,
          pageGeometry: input.pageGeometry,
        },
      });
      current = plan.prepared.candidateBytes;
      stageBytes.push(current);
    }
    const pageNumber = ordered[0].pageNumber;
    const canvases = await Promise.all(
      stageBytes.map((bytes) => renderMatchedPage(bytes, pageNumber, input, null)),
    );
    if (canvases.some((rendered) => rendered === null)) return null;
    return canvases.map((rendered, index) => ({
      canvas: (rendered as { canvas: HTMLCanvasElement }).canvas,
      route: ordered[index],
    }));
  } finally {
    for (const bytes of stageBytes) bytes.fill(0);
  }
}

/**
 * How much of what lies behind the object shows through it, 0 (opaque) to 1.
 *
 * Every pixel the object covers obeys `shown = t * absent + own`, whatever is
 * stacked above or below it, so `t` is the slope of `shown` against `absent`
 * over those pixels. Text painted over the object and the odd rim pixel do
 * not follow the line and are dropped once. Where the backdrop is flat, or
 * the object is not one evenly transparent paint, nothing can be told and
 * the object is taken as opaque.
 */
export function estimateTransmission(
  shown: Uint8ClampedArray,
  absent: Uint8ClampedArray,
  coverage: Uint8ClampedArray | null,
): number {
  const pixels = Math.min(shown.length, absent.length) >> 2;
  const stride = Math.max(1, Math.floor(pixels / MAX_FIT_SAMPLES));
  let samples: number[] = [];
  for (let pixel = 0; pixel < pixels; pixel += stride) {
    if (coverage && coverage[pixel] < 255) continue;
    samples.push(pixel * 4);
  }
  if (samples.length < 16) return 0;
  const fit = (indexes: readonly number[]) => {
    const lines: Array<{ slope: number | null; intercept: number }> = [];
    for (let channel = 0; channel < 3; channel += 1) {
      let sumX = 0;
      let sumY = 0;
      for (const i of indexes) {
        sumX += absent[i + channel];
        sumY += shown[i + channel];
      }
      const meanX = sumX / indexes.length;
      const meanY = sumY / indexes.length;
      let varX = 0;
      let cov = 0;
      for (const i of indexes) {
        const dx = absent[i + channel] - meanX;
        varX += dx * dx;
        cov += dx * (shown[i + channel] - meanY);
      }
      // A flat backdrop says nothing about how much of it shows through.
      const slope = varX / indexes.length >= FLAT_BACKDROP_VARIANCE ? cov / varX : null;
      lines.push({ slope, intercept: meanY - (slope ?? 0) * meanX });
    }
    const slopes = lines.flatMap((line) => (line.slope === null ? [] : [line.slope]));
    if (slopes.length === 0) return null;
    const shared = slopes.reduce((a, b) => a + b, 0) / slopes.length;
    const residual = (i: number) => {
      let worst = 0;
      for (let channel = 0; channel < 3; channel += 1) {
        const line = lines[channel];
        if (line.slope === null) continue;
        worst = Math.max(
          worst,
          Math.abs(shown[i + channel] - (line.slope * absent[i + channel] + line.intercept)),
        );
      }
      return worst;
    };
    let squares = 0;
    for (const i of indexes) squares += residual(i) ** 2;
    return {
      shared,
      spread: Math.max(...slopes) - Math.min(...slopes),
      residual,
      rms: Math.sqrt(squares / indexes.length),
    };
  };
  let model = fit(samples);
  if (!model) return 0;
  const limit = Math.max(4, 2.5 * model.rms);
  const first = model;
  const inliers = samples.filter((i) => first.residual(i) <= limit);
  if (inliers.length >= samples.length * 0.5 && inliers.length < samples.length) {
    samples = inliers;
    model = fit(samples);
    if (!model) return 0;
  }
  if (model.rms > LINEAR_FIT_RMS || model.spread > 0.25) return 0;
  return Math.min(1, Math.max(0, model.shared));
}

/** Which pixels two renders show differently. */
function changedMask(shown: Uint8ClampedArray, absent: Uint8ClampedArray): Uint8Array {
  const pixels = Math.min(shown.length, absent.length) >> 2;
  const changed = new Uint8Array(pixels);
  for (let pixel = 0; pixel < pixels; pixel += 1) {
    const i = pixel * 4;
    if (
      Math.max(
        Math.abs(shown[i] - absent[i]),
        Math.abs(shown[i + 1] - absent[i + 1]),
        Math.abs(shown[i + 2] - absent[i + 2]),
      ) > SAME_PIXEL_TOLERANCE
    ) {
      changed[pixel] = 1;
    }
  }
  return changed;
}

/**
 * The thin parts of a set of pixels: those no disc of `radius` that fits
 * inside the set reaches. The set is taken to continue beyond the crop, so
 * a narrow band along the crop's edge is not thin.
 *
 * Where removing an object leaves the page unchanged inside the object's own
 * footprint, something opaque lies over it there. Glyph strokes and hairlines
 * are thin; a part of the object that merely matches the page under it (the
 * white of a picture on white paper) is not.
 */
export function thinPixels(
  set: Uint8Array,
  width: number,
  height: number,
  radius: number,
): Uint8Array {
  const far = width + height + radius + 2;
  // City-block distance from each pixel of the set to the nearest pixel
  // outside it (two passes), the crop's edge counting as inside.
  const depth = new Int32Array(width * height);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const p = y * width + x;
      if (!set[p]) {
        depth[p] = 0;
        continue;
      }
      const up = y > 0 ? depth[p - width] : far;
      const left = x > 0 ? depth[p - 1] : far;
      depth[p] = Math.min(far, Math.min(up, left) + 1);
    }
  }
  for (let y = height - 1; y >= 0; y -= 1) {
    for (let x = width - 1; x >= 0; x -= 1) {
      const p = y * width + x;
      if (!set[p]) continue;
      const down = y < height - 1 ? depth[p + width] : far;
      const right = x < width - 1 ? depth[p + 1] : far;
      depth[p] = Math.min(depth[p], Math.min(down, right) + 1);
    }
  }
  // Along the crop's edge the set continues outside only where it runs
  // along that edge for longer than a stroke is wide: a band does, the end
  // of a glyph stroke that merely reaches the edge does not.
  const edgeDeep = new Uint8Array(width * height);
  const markRuns = (length: number, at: (index: number) => number) => {
    let start = -1;
    for (let index = 0; index <= length; index += 1) {
      const inSet = index < length && set[at(index)] === 1;
      if (inSet && start < 0) start = index;
      if (!inSet && start >= 0) {
        if (index - start > 2 * radius) {
          for (let k = start; k < index; k += 1) edgeDeep[at(k)] = 1;
        }
        start = -1;
      }
    }
  };
  markRuns(width, (x) => x);
  markRuns(width, (x) => (height - 1) * width + x);
  markRuns(height, (y) => y * width);
  markRuns(height, (y) => y * width + width - 1);
  // Distance from each pixel of the set to the nearest "core" pixel (one
  // deeper than the radius); anything within the radius of a core is thick.
  const reach = new Int32Array(width * height);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const p = y * width + x;
      if (!set[p]) {
        reach[p] = far;
        continue;
      }
      if (depth[p] > radius) {
        reach[p] = 0;
        continue;
      }
      // Beyond the crop a band is as deep as it likes: a core.
      const beyond = edgeDeep[p] ? 0 : far;
      const up = y > 0 ? reach[p - width] : beyond;
      const left = x > 0 ? reach[p - 1] : beyond;
      reach[p] = Math.min(far, Math.min(up, left) + 1);
    }
  }
  const thin = new Uint8Array(width * height);
  for (let y = height - 1; y >= 0; y -= 1) {
    for (let x = width - 1; x >= 0; x -= 1) {
      const p = y * width + x;
      if (!set[p]) continue;
      const beyond = edgeDeep[p] ? 0 : far;
      const down = y < height - 1 ? reach[p + width] : beyond;
      const right = x < width - 1 ? reach[p + 1] : beyond;
      reach[p] = Math.min(reach[p], Math.min(down, right) + 1);
      if (reach[p] > radius) thin[p] = 1;
    }
  }
  return thin;
}

/**
 * Which pixels of the crop are the object's, 0-255.
 *
 * Inside the proved outline every pixel is, by the outline's own
 * antialiasing; where the two renders agree there, something opaque lies
 * over the object and the pixel is marked hidden. In the rim just outside
 * (a stroke's outer half) and wherever no outline was proved, a pixel is the
 * object's only if removing the object changes it.
 */
export function buildDragCoverage(
  shown: Uint8ClampedArray,
  absent: Uint8ClampedArray,
  outlineAlpha: Uint8ClampedArray | null,
): { coverage: Uint8ClampedArray; hidden: Uint8Array } {
  const changed = changedMask(shown, absent);
  const pixels = changed.length;
  const coverage = new Uint8ClampedArray(pixels);
  const hidden = new Uint8Array(pixels);
  for (let pixel = 0; pixel < pixels; pixel += 1) {
    const inside = outlineAlpha ? outlineAlpha[pixel] : 0;
    if (inside > 0) {
      coverage[pixel] = inside;
      if (!changed[pixel] && inside === 255) hidden[pixel] = 1;
    } else if (changed[pixel]) {
      coverage[pixel] = 255;
    }
  }
  return { coverage, hidden };
}

/**
 * The pixels next to a hidden one carry the antialiased edge of whatever
 * hides it (the rim of a glyph): they are the object's, but what they show
 * is not. Mark them hidden too, so they are filled like the rest.
 */
export function spreadHidden(
  coverage: Uint8ClampedArray,
  hidden: Uint8Array,
  width: number,
  height: number,
): void {
  const seeds: number[] = [];
  for (let p = 0; p < hidden.length; p += 1) if (hidden[p]) seeds.push(p);
  for (const p of seeds) {
    const x = p % width;
    const y = (p - x) / width;
    for (let dy = -1; dy <= 1; dy += 1) {
      for (let dx = -1; dx <= 1; dx += 1) {
        const nx = x + dx;
        const ny = y + dy;
        if (nx < 0 || ny < 0 || nx >= width || ny >= height) continue;
        const n = ny * width + nx;
        if (coverage[n] > 0) hidden[n] = 1;
      }
    }
  }
}

/** Channel difference from the page under the object up to which a pixel of
 * the page without it shows only that page. */
const UNDER_TOLERANCE = 12;
/** Below this squared colour distance the object and the page under it look
 * alike, and what lies over the object cannot be told apart from it. */
const SEPARABLE_CONTRAST = 600;
/** Overlay coverage below which a pixel is the object's own. */
const MIN_OVERLAY_ALPHA = 0.04;
/** How far (per channel) the two renders may stray from one overlay over
 * the object and the page under it before the pixel is something else. */
const OVERLAY_RESIDUAL = 24;

export interface PdfVectorDragOverlay {
  /** What lies over the object, RGBA, at the object's start (crop pixels). */
  readonly overlay: Uint8ClampedArray;
  /** The page without the object, with that overlay taken off too, RGBA. */
  readonly under: Uint8ClampedArray;
  /** The pixels the overlay owns. */
  readonly mask: Uint8Array;
}

/**
 * What is painted over the object -- the text set on a filled box -- apart
 * from the object itself.
 *
 * Inside the object, the page as shown is `a * T + (1 - a) * S` and the page
 * without the object `a * T + (1 - a) * B`: `T` the overlay's colour, `a` its
 * coverage, `S` the object as it shows there, `B` the page under it. `B` is
 * the colour most of the object's footprint shows once the object is gone
 * (paper), `S` the object at the nearest pixel nothing covers; the two
 * renders then give `a` and `T` exactly, antialiased glyph edges included.
 * Small text on a box has almost no solid stroke cores, so a mask of
 * "unchanged" pixels missed most of it and the drag carried the glyphs'
 * edges along with the box. Null when the page under the object is not one
 * colour or the object looks like it.
 */
export function separateOverlay(
  shown: Uint8ClampedArray,
  absent: Uint8ClampedArray,
  coverage: Uint8ClampedArray,
  width: number,
  height: number,
): PdfVectorDragOverlay | null {
  const pixels = width * height;
  const bins = new Map<number, number>();
  let covered = 0;
  for (let p = 0; p < pixels; p += 1) {
    if (coverage[p] === 0) continue;
    covered += 1;
    const i = p * 4;
    const bin = ((absent[i] >> 3) << 10) | ((absent[i + 1] >> 3) << 5) | (absent[i + 2] >> 3);
    bins.set(bin, (bins.get(bin) ?? 0) + 1);
  }
  if (covered < 64) return null;
  let modeBin = -1;
  let modeCount = 0;
  for (const [bin, count] of bins) {
    if (count > modeCount) {
      modeBin = bin;
      modeCount = count;
    }
  }
  if (modeCount < covered * 0.5) return null;
  const sum = [0, 0, 0];
  for (let p = 0; p < pixels; p += 1) {
    if (coverage[p] === 0) continue;
    const i = p * 4;
    const bin = ((absent[i] >> 3) << 10) | ((absent[i + 1] >> 3) << 5) | (absent[i + 2] >> 3);
    if (bin !== modeBin) continue;
    sum[0] += absent[i];
    sum[1] += absent[i + 1];
    sum[2] += absent[i + 2];
  }
  const under = sum.map((value) => value / modeCount);
  // The nearest pixel the overlay leaves alone, for every covered pixel.
  const source = new Int32Array(pixels).fill(-1);
  let frontier: number[] = [];
  for (let p = 0; p < pixels; p += 1) {
    if (coverage[p] === 0) continue;
    const i = p * 4;
    if (
      Math.max(
        Math.abs(absent[i] - under[0]),
        Math.abs(absent[i + 1] - under[1]),
        Math.abs(absent[i + 2] - under[2]),
      ) <= UNDER_TOLERANCE
    ) {
      source[p] = p;
      frontier.push(p);
    }
  }
  while (frontier.length > 0) {
    const next: number[] = [];
    for (const p of frontier) {
      const x = p % width;
      const neighbours = [
        x > 0 ? p - 1 : -1,
        x < width - 1 ? p + 1 : -1,
        p >= width ? p - width : -1,
        p + width < pixels ? p + width : -1,
      ];
      for (const n of neighbours) {
        if (n < 0 || source[n] >= 0 || coverage[n] === 0) continue;
        source[n] = source[p];
        next.push(n);
      }
    }
    frontier = next;
  }
  const overlay = new Uint8ClampedArray(pixels * 4);
  const flat = new Uint8ClampedArray(pixels * 4);
  flat.set(absent.subarray(0, pixels * 4));
  const mask = new Uint8Array(pixels);
  let owned = 0;
  for (let p = 0; p < pixels; p += 1) {
    const s = source[p];
    if (coverage[p] === 0 || s < 0 || s === p) continue;
    const i = p * 4;
    const j = s * 4;
    const k = [shown[j] - under[0], shown[j + 1] - under[1], shown[j + 2] - under[2]];
    const kk = k[0] * k[0] + k[1] * k[1] + k[2] * k[2];
    if (kk < SEPARABLE_CONTRAST) continue;
    const d = [shown[i] - absent[i], shown[i + 1] - absent[i + 1], shown[i + 2] - absent[i + 2]];
    const objectShare = Math.min(1, Math.max(0, (d[0] * k[0] + d[1] * k[1] + d[2] * k[2]) / kk));
    const alpha = 1 - objectShare;
    // Something under the object it hides, not over it: the object shows
    // there as everywhere else, and only the page without it differs.
    if (alpha < MIN_OVERLAY_ALPHA) continue;
    if (
      Math.max(
        Math.abs(d[0] - objectShare * k[0]),
        Math.abs(d[1] - objectShare * k[1]),
        Math.abs(d[2] - objectShare * k[2]),
      ) > OVERLAY_RESIDUAL
    ) {
      continue;
    }
    for (let channel = 0; channel < 3; channel += 1) {
      overlay[i + channel] = (shown[i + channel] - objectShare * shown[j + channel]) / alpha;
      flat[i + channel] = under[channel];
    }
    overlay[i + 3] = Math.round(alpha * 255);
    mask[p] = 1;
    owned += 1;
  }
  return owned > 0 ? { overlay, under: flat, mask } : null;
}

/** Per-channel spread within which a pixel shows the object's one paint. */
const UNIFORM_PAINT_TOLERANCE = 8;
/** Per-channel distance within which a pixel is the page plus the object and
 * nothing over them. */
const CLEAN_PIXEL_TOLERANCE = 10;

/**
 * The one colour an object adds where it fully covers the page (`own` =
 * shown - transmission x page without it), when it is painted with one: a
 * filled or stroked shape is. Null for a gradient, a picture, or a sample
 * too small to tell.
 */
export function uniformPaintColour(
  shown: Uint8ClampedArray,
  absent: Uint8ClampedArray,
  coverage: Uint8ClampedArray,
  hidden: Uint8Array,
  transmission: number,
): readonly [number, number, number] | null {
  const pixels = coverage.length;
  const stride = Math.max(1, Math.floor(pixels / MAX_FIT_SAMPLES));
  const channels: number[][] = [[], [], []];
  for (let p = 0; p < pixels; p += stride) {
    if (coverage[p] !== 255 || hidden[p]) continue;
    for (let c = 0; c < 3; c += 1) {
      channels[c].push(shown[p * 4 + c] - transmission * absent[p * 4 + c]);
    }
  }
  const count = channels[0].length;
  if (count < 64) return null;
  const median = channels.map((values) => {
    const sorted = [...values].sort((a, b) => a - b);
    return sorted[Math.floor(sorted.length / 2)];
  }) as [number, number, number];
  let near = 0;
  for (let i = 0; i < count; i += 1) {
    if (channels.every((values, c) => Math.abs(values[i] - median[c]) <= UNIFORM_PAINT_TOLERANCE)) {
      near += 1;
    }
  }
  return near >= count * 0.6 ? median : null;
}

export interface PdfVectorUniformPaintLayers {
  /** Coverage, its antialiased rim measured where no outline was proved. */
  readonly coverage: Uint8ClampedArray;
  /** coverage x the paint: the object and nothing else. */
  readonly own: Int16Array;
  /** What lies over the object, as in PdfVectorDragOverlay; null if nothing. */
  readonly overlay: PdfVectorDragOverlay | null;
}

/**
 * An object of one paint, apart from what lies over it.
 *
 * Without the object a pixel shows the page `B`; with it, where nothing lies
 * over it, `B + c * (O - (1 - t) * B)` -- `c` the coverage, `O` the paint,
 * `t` the transmission. Every pixel that is not that holds something painted
 * over the object (text crossing a shape, its edge included): with `B` taken
 * from the nearest pixel that is, the two renders give that overlay's
 * coverage and colour. The object itself is then exactly `c * O`, so its
 * edge stays smooth where text crossed it -- the edge used to carry bites of
 * the glyphs and move them along like something seen under water. Where no
 * outline was proved, the rim's coverage is measured the same way.
 */
export function uniformPaintLayers(
  shown: Uint8ClampedArray,
  absent: Uint8ClampedArray,
  coverageIn: Uint8ClampedArray,
  width: number,
  height: number,
  transmission: number,
  paint: readonly [number, number, number],
  outlineProved: boolean,
): PdfVectorUniformPaintLayers {
  const pixels = width * height;
  const coverage = new Uint8ClampedArray(coverageIn);
  const step = (p: number, c: number) => paint[c] - (1 - transmission) * absent[p * 4 + c];
  const residualAt = (p: number, share: number) =>
    Math.max(
      Math.abs(shown[p * 4] - absent[p * 4] - share * step(p, 0)),
      Math.abs(shown[p * 4 + 1] - absent[p * 4 + 1] - share * step(p, 1)),
      Math.abs(shown[p * 4 + 2] - absent[p * 4 + 2] - share * step(p, 2)),
    );
  if (!outlineProved) {
    // The rim's own antialiasing, where the pixel is the page plus part of
    // the paint and nothing else.
    for (let p = 0; p < pixels; p += 1) {
      if (coverage[p] === 0) continue;
      const x = p % width;
      const y = (p - x) / width;
      const rim =
        x === 0 || y === 0 || x === width - 1 || y === height - 1 ||
        coverage[p - 1] === 0 || coverage[p + 1] === 0 ||
        coverage[p - width] === 0 || coverage[p + width] === 0;
      if (!rim) continue;
      let kk = 0;
      let dk = 0;
      for (let c = 0; c < 3; c += 1) {
        const k = step(p, c);
        kk += k * k;
        dk += (shown[p * 4 + c] - absent[p * 4 + c]) * k;
      }
      if (kk < SEPARABLE_CONTRAST) continue;
      const share = Math.min(1, Math.max(0, dk / kk));
      if (residualAt(p, share) <= CLEAN_PIXEL_TOLERANCE) coverage[p] = Math.round(share * 255);
    }
  }
  const own = new Int16Array(pixels * 3);
  const source = new Int32Array(pixels).fill(-1);
  let frontier: number[] = [];
  for (let p = 0; p < pixels; p += 1) {
    if (coverage[p] === 0) continue;
    const share = coverage[p] / 255;
    for (let c = 0; c < 3; c += 1) own[p * 3 + c] = Math.round(share * paint[c]);
    if (residualAt(p, share) <= CLEAN_PIXEL_TOLERANCE) {
      source[p] = p;
      frontier.push(p);
    }
  }
  // The page under each pixel something lies over: the nearest clean one's.
  while (frontier.length > 0) {
    const next: number[] = [];
    for (const p of frontier) {
      const x = p % width;
      const neighbours = [
        x > 0 ? p - 1 : -1,
        x < width - 1 ? p + 1 : -1,
        p >= width ? p - width : -1,
        p + width < pixels ? p + width : -1,
      ];
      for (const n of neighbours) {
        if (n < 0 || source[n] >= 0 || coverage[n] === 0) continue;
        source[n] = source[p];
        next.push(n);
      }
    }
    frontier = next;
  }
  const overlay = new Uint8ClampedArray(pixels * 4);
  const under = new Uint8ClampedArray(pixels * 4);
  under.set(absent.subarray(0, pixels * 4));
  const mask = new Uint8Array(pixels);
  let owned = 0;
  for (let p = 0; p < pixels; p += 1) {
    const s = source[p];
    if (coverage[p] === 0 || s < 0 || s === p) continue;
    const share = coverage[p] / 255;
    const i = p * 4;
    const b = [absent[s * 4], absent[s * 4 + 1], absent[s * 4 + 2]];
    const k = b.map((value, c) => share * (paint[c] - (1 - transmission) * value));
    const kk = k[0] * k[0] + k[1] * k[1] + k[2] * k[2];
    if (kk < SEPARABLE_CONTRAST) continue;
    const d = [shown[i] - absent[i], shown[i + 1] - absent[i + 1], shown[i + 2] - absent[i + 2]];
    const objectShare = Math.min(1, Math.max(0, (d[0] * k[0] + d[1] * k[1] + d[2] * k[2]) / kk));
    const alpha = 1 - objectShare;
    if (alpha < MIN_OVERLAY_ALPHA) continue;
    if (
      Math.max(
        Math.abs(d[0] - objectShare * k[0]),
        Math.abs(d[1] - objectShare * k[1]),
        Math.abs(d[2] - objectShare * k[2]),
      ) > OVERLAY_RESIDUAL
    ) {
      continue;
    }
    for (let c = 0; c < 3; c += 1) {
      const objectShown = b[c] + k[c];
      overlay[i + c] = (shown[i + c] - objectShare * objectShown) / alpha;
      under[i + c] = b[c];
    }
    overlay[i + 3] = Math.round(alpha * 255);
    mask[p] = 1;
    owned += 1;
  }
  return {
    coverage,
    own,
    overlay: owned > 0 ? { overlay, under, mask } : null,
  };
}

export interface PdfVectorDragModel {
  /** The object's crop, in page-canvas pixels. */
  readonly width: number;
  readonly height: number;
  /** The single object's transmission (0 for a model of several objects). */
  readonly transmission: number;
  /** Per pixel, how much of the backdrop shows there: the product of the
   * transmissions of every object covering it, all of it where none does. */
  readonly through: Float32Array;
  /** Per covered pixel, what the objects add: shown - through * absent. */
  readonly own: Int16Array;
  readonly coverage: Uint8ClampedArray;
  /** The whole page without the object. */
  readonly absent: {
    readonly data: Uint8ClampedArray;
    readonly width: number;
    readonly height: number;
  };
  /** The page without the object, cut under the object's start. */
  readonly patch: HTMLCanvasElement;
  /**
   * What lies over the object at its start (text on a filled box), the same
   * size and place as `patch`: it stays there, above the moving object, as
   * it will be painted once the object has moved. `patch` then leaves it
   * out. Null when nothing lies over the object or it cannot be told apart.
   */
  readonly overlay: HTMLCanvasElement | null;
  /** What the page paints over the object, everywhere: the ghost slides
   * under it wherever it lands. Null when not known (then `overlay`). */
  readonly above: PdfVectorDragAbove | null;
  /** How far the crop reaches past the object's frame, in page pixels. */
  readonly pad: {
    readonly left: number;
    readonly top: number;
    readonly right: number;
    readonly bottom: number;
  };
}

/**
 * How much of the backdrop shows at a pixel the object covers only in part
 * (an antialiased edge): the uncovered share shows all of it, the covered
 * share `transmission` of it. Composing an edge pixel with this, fully
 * opaque, is exact; blending a full-coverage value by the coverage instead
 * left a fringe of the old backdrop around an object dragged onto a
 * different one.
 */
function pixelTransmission(coverage: number, transmission: number): number {
  const covered = coverage / 255;
  return covered * transmission + (1 - covered);
}

/**
 * What the object adds at each covered pixel: shown minus the share of the
 * backdrop showing through. A hidden pixel shows something else instead of
 * the object; it takes the contribution of the nearest pixel that does show
 * it (with `width`), or of all of them on average (without).
 */
export function buildOwnContribution(
  shown: Uint8ClampedArray,
  absent: Uint8ClampedArray,
  coverage: Uint8ClampedArray,
  hidden: Uint8Array,
  transmission: number | Float32Array,
  width?: number,
): Int16Array {
  const pixels = coverage.length;
  const own = new Int16Array(pixels * 3);
  const sum = [0, 0, 0];
  let visible = 0;
  for (let pixel = 0; pixel < pixels; pixel += 1) {
    if (coverage[pixel] === 0) continue;
    const i = pixel * 4;
    const through =
      typeof transmission === "number"
        ? pixelTransmission(coverage[pixel], transmission)
        : transmission[pixel];
    for (let channel = 0; channel < 3; channel += 1) {
      own[pixel * 3 + channel] = Math.round(
        shown[i + channel] - through * absent[i + channel],
      );
    }
    if (!hidden[pixel] && coverage[pixel] === 255) {
      sum[0] += own[pixel * 3];
      sum[1] += own[pixel * 3 + 1];
      sum[2] += own[pixel * 3 + 2];
      visible += 1;
    }
  }
  if (width && width > 0) {
    // Grow the visible pixels into the hidden ones, nearest first.
    const height = pixels / width;
    const settled = new Uint8Array(pixels);
    let frontier: number[] = [];
    for (let pixel = 0; pixel < pixels; pixel += 1) {
      if (coverage[pixel] > 0 && !hidden[pixel]) {
        settled[pixel] = 1;
        frontier.push(pixel);
      }
    }
    while (frontier.length > 0) {
      const next: number[] = [];
      for (const pixel of frontier) {
        const x = pixel % width;
        const y = (pixel - x) / width;
        const neighbours = [
          x > 0 ? pixel - 1 : -1,
          x < width - 1 ? pixel + 1 : -1,
          y > 0 ? pixel - width : -1,
          y < height - 1 ? pixel + width : -1,
        ];
        for (const n of neighbours) {
          if (n < 0 || settled[n] || !hidden[n] || coverage[n] === 0) continue;
          settled[n] = 1;
          own[n * 3] = own[pixel * 3];
          own[n * 3 + 1] = own[pixel * 3 + 1];
          own[n * 3 + 2] = own[pixel * 3 + 2];
          next.push(n);
        }
      }
      frontier = next;
    }
    return own;
  }
  if (visible > 0) {
    const mean = sum.map((value) => Math.round(value / visible));
    for (let pixel = 0; pixel < pixels; pixel += 1) {
      if (!hidden[pixel]) continue;
      own[pixel * 3] = mean[0];
      own[pixel * 3 + 1] = mean[1];
      own[pixel * 3 + 2] = mean[2];
    }
  }
  return own;
}

/** The object turned about a point: where its unturned crop lies, in page pixels. */
export interface PdfVectorDragTurn {
  /** Clockwise, as the frame's rotation handle reports it. */
  readonly degrees: number;
  readonly centerX: number;
  readonly centerY: number;
  readonly crop: { x: number; y: number; width: number; height: number };
}

/**
 * The object as it would show with its top-left at (`targetX`, `targetY`)
 * and stretched to `outWidth` x `outHeight`, all in page-canvas pixels:
 * the backdrop it lands on showing through, plus its own contribution.
 */
export function composeDragGhostPixels(
  model: Pick<PdfVectorDragModel, "width" | "height" | "own" | "coverage" | "absent"> & {
    readonly through?: Float32Array;
    readonly transmission?: number;
    readonly above?: PdfVectorDragAbove | null;
  },
  out: Uint8ClampedArray,
  outWidth: number,
  outHeight: number,
  targetX: number,
  targetY: number,
  /** With a turn, the target is the turned crop's bounding box: each of its
   * pixels is traced back into the crop, and still composed with the
   * backdrop where it really lands -- the backdrop does not turn. */
  turn?: PdfVectorDragTurn,
  /** A resize carried across the opposite edge: the object reads mirrored
   * on that axis, the backdrop it lands on does not. */
  mirror?: Readonly<{ x: boolean; y: boolean }>,
): void {
  const { width, height, own, coverage, absent } = model;
  const radians = turn ? (turn.degrees * Math.PI) / 180 : 0;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  for (let y = 0; y < outHeight; y += 1) {
    const pageY = targetY + y;
    for (let x = 0; x < outWidth; x += 1) {
      const o = (y * outWidth + x) * 4;
      const pageX = targetX + x;
      let u: number;
      let v: number;
      if (turn) {
        // Undo the clockwise turn about the centre (y grows downwards).
        const dx = pageX + 0.5 - turn.centerX;
        const dy = pageY + 0.5 - turn.centerY;
        const backX = turn.centerX + cos * dx + sin * dy;
        const backY = turn.centerY - sin * dx + cos * dy;
        u = Math.floor(((backX - turn.crop.x) * width) / turn.crop.width);
        v = Math.floor(((backY - turn.crop.y) * height) / turn.crop.height);
        if (u < 0 || v < 0 || u >= width || v >= height) {
          out[o + 3] = 0;
          continue;
        }
      } else {
        u = Math.min(width - 1, Math.floor((x * width) / outWidth));
        v = Math.min(height - 1, Math.floor((y * height) / outHeight));
        if (mirror?.x) u = width - 1 - u;
        if (mirror?.y) v = height - 1 - v;
      }
      const source = v * width + u;
      const alpha = coverage[source];
      if (
        alpha === 0 ||
        pageX < 0 ||
        pageY < 0 ||
        pageX >= absent.width ||
        pageY >= absent.height
      ) {
        out[o + 3] = 0;
        continue;
      }
      const b = (pageY * absent.width + pageX) * 4;
      const through = model.through
        ? model.through[source]
        : pixelTransmission(alpha, model.transmission ?? 0);
      const above = model.above;
      if (above) {
        // Under what the page paints after the object, where it lands.
        const q = pageY * absent.width + pageX;
        const a = above.alpha[q] / 255;
        for (let c = 0; c < 3; c += 1) {
          out[o + c] =
            through * absent.data[b + c] +
            (1 - through) * above.premultiplied[q * 3 + c] +
            (1 - a) * own[source * 3 + c];
        }
        out[o + 3] = 255;
        continue;
      }
      out[o] = through * absent.data[b] + own[source * 3];
      out[o + 1] = through * absent.data[b + 1] + own[source * 3 + 1];
      out[o + 2] = through * absent.data[b + 2] + own[source * 3 + 2];
      // The pixel already holds its share of the backdrop it lands on.
      out[o + 3] = 255;
    }
  }
}

type DragModelItem = Pick<PdfVectorContentItem, "simpleSubpath" | "paintMode">;

function outlineAlphaFor(
  item: DragModelItem | null | undefined,
  bounds: FractionRect,
  width: number,
  height: number,
  createCanvas: () => HTMLCanvasElement,
): Uint8ClampedArray | null {
  const outline = item?.simpleSubpath;
  if (
    !outline ||
    outline.commands.length < 3 ||
    !(outline.closed || item?.paintMode !== "stroke")
  ) {
    return null;
  }
  const canvas = createCanvas();
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  if (!context) return null;
  traceNativeVectorOutline(context, outline, bounds, width, height);
  context.fillStyle = "#000";
  context.fill();
  const rgba = context.getImageData(0, 0, width, height).data;
  const alpha = new Uint8ClampedArray(width * height);
  for (let pixel = 0; pixel < alpha.length; pixel += 1) alpha[pixel] = rgba[pixel * 4 + 3];
  return alpha;
}

/** One object of a drag model: the page once it, too, has been removed. */
export interface PdfVectorDragStage {
  readonly canvas: HTMLCanvasElement;
  /** The renderer's outline evidence for a source vector; null for an
   * object with none (an annotation, a picture). */
  readonly item?: DragModelItem | null;
  /**
   * The object's own rectangle when it paints every pixel of it (an
   * unrotated opaque picture): the whole of it moves, including the parts
   * that match the page under them.
   */
  readonly footprint?: FractionRect | null;
}

/** One object's coverage of the crop, from the page before and after its removal. */
export function stageCoverage(
  before: Uint8ClampedArray,
  after: Uint8ClampedArray,
  width: number,
  height: number,
  shape: {
    outlineAlpha?: Uint8ClampedArray | null;
    /** In crop pixels. */
    footprint?: { x: number; y: number; width: number; height: number } | null;
    /** A source vector without a proved outline: what lies thinly over it
     * (glyphs) leaves holes in its difference; close them. */
    closeThinHoles?: boolean;
  },
): { coverage: Uint8ClampedArray; hidden: Uint8Array } {
  if (shape.footprint) {
    const changed = changedMask(before, after);
    const coverage = new Uint8ClampedArray(width * height);
    const unchangedInside = new Uint8Array(width * height);
    const left = Math.max(0, Math.round(shape.footprint.x));
    const top = Math.max(0, Math.round(shape.footprint.y));
    const right = Math.min(width, Math.round(shape.footprint.x + shape.footprint.width));
    const bottom = Math.min(height, Math.round(shape.footprint.y + shape.footprint.height));
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) {
        const p = y * width + x;
        const inside = x >= left && x < right && y >= top && y < bottom;
        if (inside) {
          coverage[p] = 255;
          if (!changed[p]) unchangedInside[p] = 1;
        } else if (changed[p]) {
          coverage[p] = 255;
        }
      }
    }
    // Thin unchanged strokes are something painted over the picture; a broad
    // unchanged area is the picture matching the page under it.
    return { coverage, hidden: thinPixels(unchangedInside, width, height, THIN_STROKE_PX) };
  }
  const { coverage, hidden } = buildDragCoverage(before, after, shape.outlineAlpha ?? null);
  if (!shape.outlineAlpha && shape.closeThinHoles) {
    const unchanged = new Uint8Array(width * height);
    for (let p = 0; p < unchanged.length; p += 1) unchanged[p] = coverage[p] === 0 ? 1 : 0;
    const thin = thinPixels(unchanged, width, height, THIN_STROKE_PX);
    for (let p = 0; p < thin.length; p += 1) {
      if (!thin[p]) continue;
      coverage[p] = 255;
      hidden[p] = 1;
    }
  }
  return { coverage, hidden };
}

/**
 * The drag model of one or several objects, from the page as shown and the
 * page after each of them is removed in turn (`stages`, the last being the
 * page without any of them).
 */
export function buildLayeredDragModel(
  pageCanvas: HTMLCanvasElement,
  stages: readonly PdfVectorDragStage[],
  bounds: FractionRect,
  options: {
    createCanvas?: () => HTMLCanvasElement;
    /** What the objects really occupy when that is more than their frame. */
    painted?: FractionRect | null;
    /** What the page paints over a single object (renderNativeVectorAbove). */
    above?: PdfVectorDragAbove | null;
  } = {},
): PdfVectorDragModel | null {
  if (stages.length === 0) return null;
  const createCanvas = options.createCanvas ?? (() => document.createElement("canvas"));
  // The rim is part of the crop so a stroke's outer half comes along.
  const rimX = OUTLINE_RIM_PX / Math.max(1, pageCanvas.width);
  const rimY = OUTLINE_RIM_PX / Math.max(1, pageCanvas.height);
  const painted = options.painted ?? bounds;
  const left = Math.max(0, Math.min(bounds.x, painted.x) - rimX);
  const top = Math.max(0, Math.min(bounds.y, painted.y) - rimY);
  const right = Math.min(
    1,
    Math.max(bounds.x + bounds.width, painted.x + painted.width) + rimX,
  );
  const bottom = Math.min(
    1,
    Math.max(bounds.y + bounds.height, painted.y + painted.height) + rimY,
  );
  const grown: FractionRect = { x: left, y: top, width: right - left, height: bottom - top };
  const shownCrop = cropCanvasRegion(pageCanvas, grown, { createCanvas });
  if (!shownCrop) return null;
  const width = shownCrop.canvas.width;
  const height = shownCrop.canvas.height;
  try {
    const shown = shownCrop.canvas.getContext("2d")?.getImageData(0, 0, width, height).data;
    if (!shown) return null;
    const through = new Float32Array(width * height).fill(1);
    const coverage = new Uint8ClampedArray(width * height);
    const hidden = new Uint8Array(width * height);
    let before: Uint8ClampedArray = shown;
    let lastPatch: HTMLCanvasElement | null = null;
    let lastTransmission = 0;
    let outlineProved = false;
    for (const stage of stages) {
      const crop = cropCanvasRegion(stage.canvas, grown, { createCanvas });
      if (!crop || crop.canvas.width !== width || crop.canvas.height !== height) return null;
      const after = crop.canvas.getContext("2d")?.getImageData(0, 0, width, height).data;
      if (!after) return null;
      const outlineAlpha = stage.footprint
        ? null
        : outlineAlphaFor(stage.item, grown, width, height, createCanvas);
      outlineProved = Boolean(stage.footprint) || outlineAlpha !== null;
      const layer = stageCoverage(before, after, width, height, {
        outlineAlpha,
        footprint: stage.footprint
          ? {
              x: ((stage.footprint.x - grown.x) / grown.width) * width,
              y: ((stage.footprint.y - grown.y) / grown.height) * height,
              width: (stage.footprint.width / grown.width) * width,
              height: (stage.footprint.height / grown.height) * height,
            }
          : null,
        // `item` undefined or a vector's: a source vector; null: not one.
        closeThinHoles: stage.item !== null,
      });
      // The line is fitted where this object plainly shows.
      const fitted = new Uint8ClampedArray(layer.coverage);
      for (let p = 0; p < fitted.length; p += 1) if (layer.hidden[p]) fitted[p] = 0;
      const transmission = estimateTransmission(before, after, fitted);
      for (let p = 0; p < through.length; p += 1) {
        if (layer.coverage[p] === 0) continue;
        through[p] *= pixelTransmission(layer.coverage[p], transmission);
        if (layer.coverage[p] > coverage[p]) coverage[p] = layer.coverage[p];
        // Hidden in one object's layer only matters if nothing else of the
        // selection shows there; settled below against the final page.
        if (layer.hidden[p]) hidden[p] = 1;
      }
      before = after;
      lastPatch = crop.canvas;
      lastTransmission = transmission;
    }
    const last = stages[stages.length - 1].canvas;
    const pageContext = last.getContext("2d");
    if (!pageContext || !lastPatch) return null;
    const page = pageContext.getImageData(0, 0, last.width, last.height);
    // A pixel some object of the selection visibly changes is not hidden,
    // whatever another object's layer said.
    const changedOverall = changedMask(shown, before);
    for (let p = 0; p < hidden.length; p += 1) {
      if (hidden[p] && changedOverall[p] && stages.length > 1) hidden[p] = 0;
    }
    // What the page paints over the object, known everywhere: taken off
    // both renders under its footprint, so what is left is the object over
    // the page beneath (`(x - A) / (1 - a)`); where it covers the object
    // all but entirely the object cannot be seen and is filled in.
    const above =
      stages.length === 1 &&
      options.above &&
      options.above.width === pageCanvas.width &&
      options.above.height === pageCanvas.height
        ? options.above
        : null;
    let objectShown: Uint8ClampedArray = shown;
    let objectAbsent: Uint8ClampedArray = before;
    if (above) {
      objectShown = new Uint8ClampedArray(shown);
      objectAbsent = new Uint8ClampedArray(before);
      const origin = shownCrop.rect;
      for (let p = 0; p < width * height; p += 1) {
        const x = p % width;
        const y = (p - x) / width;
        const q = (origin.y + y) * above.width + origin.x + x;
        const a = above.alpha[q] / 255;
        if (a === 0) continue;
        if (a >= 0.9) {
          if (coverage[p] > 0) hidden[p] = 1;
          continue;
        }
        for (let c = 0; c < 3; c += 1) {
          const lifted = above.premultiplied[q * 3 + c];
          objectShown[p * 4 + c] = (shown[p * 4 + c] - lifted) / (1 - a);
          objectAbsent[p * 4 + c] = (before[p * 4 + c] - lifted) / (1 - a);
        }
      }
    }
    // One object of one paint: the object is exactly coverage x paint, and
    // everything else under its footprint is what lies over it.
    const paint =
      stages.length === 1
        ? uniformPaintColour(objectShown, objectAbsent, coverage, hidden, lastTransmission)
        : null;
    const painted = paint
      ? uniformPaintLayers(objectShown, objectAbsent, coverage, width, height, lastTransmission, paint, outlineProved)
      : null;
    if (painted) {
      coverage.set(painted.coverage);
      for (let p = 0; p < through.length; p += 1) {
        through[p] = coverage[p] === 0 ? 1 : pixelTransmission(coverage[p], lastTransmission);
      }
    }
    // What is painted over the objects stays where it is and keeps out of
    // the ghost: the ghost is the objects alone, filled where it was.
    // (With the layer over the object known, the ghost is composed under it
    // wherever it lands, and nothing is lifted off here.)
    const separated = above
      ? null
      : painted
        ? painted.overlay
        : separateOverlay(shown, before, coverage, width, height);
    let patch = lastPatch;
    let overlay: HTMLCanvasElement | null = null;
    if (separated) {
      for (let p = 0; p < hidden.length; p += 1) if (separated.mask[p]) hidden[p] = 1;
      const toCanvas = (data: Uint8ClampedArray) => {
        const canvas = createCanvas();
        canvas.width = width;
        canvas.height = height;
        const context = canvas.getContext("2d");
        if (!context) return null;
        const image = context.createImageData(width, height);
        image.data.set(data);
        context.putImageData(image, 0, 0);
        return canvas;
      };
      const under = toCanvas(separated.under);
      overlay = under ? toCanvas(separated.overlay) : null;
      if (under && overlay) patch = under;
      else overlay = null;
    }
    spreadHidden(coverage, hidden, width, height);
    return {
      width,
      height,
      transmission: stages.length === 1 ? lastTransmission : 0,
      through,
      own:
        painted?.own ??
        buildOwnContribution(objectShown, objectAbsent, coverage, hidden, through, width),
      coverage,
      absent: { data: page.data, width: page.width, height: page.height },
      patch,
      overlay,
      above,
      pad: {
        left: (bounds.x - grown.x) * pageCanvas.width,
        top: (bounds.y - grown.y) * pageCanvas.height,
        right: (grown.x + grown.width - bounds.x - bounds.width) * pageCanvas.width,
        bottom: (grown.y + grown.height - bounds.y - bounds.height) * pageCanvas.height,
      },
    };
  } catch {
    // A canvas that refuses pixel reads has no model; the plain ghost stays.
    return null;
  }
}

/** The drag model of a single object. */
export function buildNativeVectorDragModel(
  pageCanvas: HTMLCanvasElement,
  absentCanvas: HTMLCanvasElement,
  bounds: FractionRect,
  item: DragModelItem | null | undefined,
  options: {
    createCanvas?: () => HTMLCanvasElement;
    /** What the object really occupies when that is more than its frame. */
    painted?: FractionRect | null;
    /** See `PdfVectorDragStage.footprint`. */
    footprint?: FractionRect | null;
    /** What the page paints over the object (renderNativeVectorAbove). */
    above?: PdfVectorDragAbove | null;
  } = {},
): PdfVectorDragModel | null {
  return buildLayeredDragModel(
    pageCanvas,
    [{ canvas: absentCanvas, item, footprint: options.footprint ?? null }],
    bounds,
    { createCanvas: options.createCanvas, painted: options.painted, above: options.above },
  );
}

/** Paint the object into `canvas` as it shows at the given page-pixel box. */
export function paintNativeVectorDragGhost(
  model: PdfVectorDragModel,
  canvas: HTMLCanvasElement,
  target: { x: number; y: number; width: number; height: number },
  turn?: PdfVectorDragTurn | null,
  mirror?: Readonly<{ x: boolean; y: boolean }> | null,
): boolean {
  const width = Math.max(1, Math.round(target.width));
  const height = Math.max(1, Math.round(target.height));
  if (canvas.width !== width) canvas.width = width;
  if (canvas.height !== height) canvas.height = height;
  const context = canvas.getContext("2d");
  if (!context) return false;
  const image = context.createImageData(width, height);
  composeDragGhostPixels(
    model,
    image.data,
    width,
    height,
    Math.round(target.x),
    Math.round(target.y),
    turn ?? undefined,
    mirror ?? undefined,
  );
  context.putImageData(image, 0, 0);
  return true;
}

/** The scratch canvas each stage composes its moving object in. */
const DRAG_STAGE_SCRATCH = new WeakMap<HTMLCanvasElement, HTMLCanvasElement>();

/**
 * Paint one whole frame of a drag into `stage`: a canvas the size of the
 * page canvas that lies exactly over it and never moves. The page without
 * the object where it started, the object composed where it now is, and what
 * the page paints over its start, each at its own page pixels.
 *
 * Nothing on screen moves, only pixels change, and they all change in this
 * one call. A floating canvas that carried the same pixels had to be moved
 * as well: the pixels were written at once and the box a render later, so
 * every pointer move showed a frame of new pixels in the old box and the
 * text that stands still inside the shape jumped by the move and back. The
 * stage also maps its pixels to the screen exactly as the page canvas does
 * (same size, same box), so the text it repeats lands on the page's own.
 */
export function paintNativeVectorDragStage(
  model: PdfVectorDragModel,
  stage: HTMLCanvasElement,
  /** Top-left of the object's crop where the gesture began, page pixels. */
  origin: { x: number; y: number },
  target: { x: number; y: number; width: number; height: number },
  turn?: PdfVectorDragTurn | null,
  mirror?: Readonly<{ x: boolean; y: boolean }> | null,
  createCanvas: () => HTMLCanvasElement = () => document.createElement("canvas"),
): boolean {
  const pageWidth = Math.max(1, model.absent.width);
  const pageHeight = Math.max(1, model.absent.height);
  // Sizing a canvas clears it; an unchanged one is cleared by hand.
  const resized = stage.width !== pageWidth || stage.height !== pageHeight;
  if (resized) {
    stage.width = pageWidth;
    stage.height = pageHeight;
  }
  const context = stage.getContext("2d");
  if (!context) return false;
  if (!resized) context.clearRect(0, 0, pageWidth, pageHeight);
  let scratch = DRAG_STAGE_SCRATCH.get(stage);
  if (!scratch) {
    scratch = createCanvas();
    DRAG_STAGE_SCRATCH.set(stage, scratch);
  }
  const originX = Math.round(origin.x);
  const originY = Math.round(origin.y);
  // Where the object was: the page without it.
  context.drawImage(model.patch, originX, originY);
  // The object where it is now, over whatever it lands on.
  if (!paintNativeVectorDragGhost(model, scratch, target, turn, mirror)) return false;
  context.drawImage(scratch, Math.round(target.x), Math.round(target.y));
  // What the page paints over the object's start stays above it.
  if (model.overlay) context.drawImage(model.overlay, originX, originY);
  return true;
}

/** How deep (in page pixels) an object may be on the axis a frame scales
 * alone for the scale to read as a stroke width, not a shape's size. */
const THIN_LAYER_PX = 48;

/**
 * The axis along which a lone-axis stretch of a thin layer repeats the
 * layer's middle strip: "y" for a horizontal line stretched taller, "x" for
 * a vertical one stretched wider; null when the gesture is anything else.
 */
export function thinAxisStretch(
  box: { x: number; y: number; width: number; height: number },
  rotation: number,
  mirrorX: boolean,
  mirrorY: boolean,
  scaleX: number,
  scaleY: number,
): "x" | "y" | null {
  if (rotation % 360 !== 0 || mirrorX || mirrorY) return null;
  const sameX = Math.abs(scaleX - 1) < 1e-6;
  const sameY = Math.abs(scaleY - 1) < 1e-6;
  if (sameX && !sameY && scaleY > 1 && box.height <= THIN_LAYER_PX && box.height * 4 <= box.width) {
    return "y";
  }
  if (sameY && !sameX && scaleX > 1 && box.width <= THIN_LAYER_PX && box.width * 4 <= box.height) {
    return "x";
  }
  return null;
}

/**
 * Paint the plain crop of an object into `stage` where its frame now is:
 * what the drag shows before the page without the object has been rendered.
 * The same stage as paintNativeVectorDragStage -- the size of the page
 * canvas, exactly over it, never moving -- so the crop is carried by pixels
 * written in the pointer's own task, not by a box a render later, and the
 * step to the composed ghost changes no element. The crop still holds
 * whatever is painted over the object; only the model can leave that behind.
 */
export function paintNativeVectorCropStage(
  stage: HTMLCanvasElement,
  crop: HTMLCanvasElement,
  page: { width: number; height: number },
  /** The object's frame, in page pixels. */
  target: { x: number; y: number; width: number; height: number },
  options: {
    /** Clockwise degrees about the frame's centre. */
    rotationDegrees?: number;
    mirror?: Readonly<{ x: boolean; y: boolean }> | null;
    /** Translucent while the original still shows underneath. */
    opacity?: number;
    /** The whole motion as one page-pixel affine, with where the crop was
     * cut (see paintNativeVectorLayerStage): wins over `target`. */
    matrix?: Readonly<{
      transform: readonly [number, number, number, number, number, number];
      cropAt: { x: number; y: number };
    }> | null;
  } = {},
): boolean {
  const pageWidth = Math.max(1, Math.round(page.width));
  const pageHeight = Math.max(1, Math.round(page.height));
  const resized = stage.width !== pageWidth || stage.height !== pageHeight;
  if (resized) {
    stage.width = pageWidth;
    stage.height = pageHeight;
  }
  const context = stage.getContext("2d");
  if (!context) return false;
  if (!resized) context.clearRect(0, 0, pageWidth, pageHeight);
  if (crop.width === 0 || crop.height === 0) return true;
  if (options.matrix) {
    const [a, b, c, d, e, f] = options.matrix.transform;
    context.save();
    context.globalAlpha = options.opacity ?? 1;
    context.transform(a, b, c, d, e, f);
    context.drawImage(crop, options.matrix.cropAt.x, options.matrix.cropAt.y);
    context.restore();
    return true;
  }
  // Unstretched, the crop keeps its own pixels: a move is a copy at whole
  // page pixels, never a resample.
  const width =
    Math.abs(target.width - crop.width) < 1 ? crop.width : Math.max(1, Math.round(target.width));
  const height =
    Math.abs(target.height - crop.height) < 1
      ? crop.height
      : Math.max(1, Math.round(target.height));
  const x = Math.round(target.x);
  const y = Math.round(target.y);
  const rotation = options.rotationDegrees ?? 0;
  const mirror = options.mirror ?? null;
  context.save();
  context.globalAlpha = options.opacity ?? 1;
  const strip = thinAxisStretch(
    { x, y, width: crop.width, height: crop.height },
    rotation,
    mirror?.x === true,
    mirror?.y === true,
    width / crop.width,
    height / crop.height,
  );
  if (strip === "y") {
    context.drawImage(crop, 0, Math.floor(crop.height / 2), crop.width, 1, x, y, width, height);
  } else if (strip === "x") {
    context.drawImage(crop, Math.floor(crop.width / 2), 0, 1, crop.height, x, y, width, height);
  } else if (rotation % 360 === 0 && !mirror?.x && !mirror?.y) {
    context.drawImage(crop, x, y, width, height);
  } else {
    context.translate(x + width / 2, y + height / 2);
    context.rotate((rotation * Math.PI) / 180);
    context.scale(mirror?.x ? -1 : 1, mirror?.y ? -1 : 1);
    context.drawImage(crop, -width / 2, -height / 2, width, height);
  }
  context.restore();
  return true;
}

/**
 * Where the ghost (or, with `rect === start`, the origin patch) sits once
 * the crop's rim is counted in: as a page-percent box for the DOM and as a
 * page-pixel box for painting.
 */
export function nativeVectorDragGhostBox(
  model: Pick<PdfVectorDragModel, "absent" | "pad"> &
    Partial<Pick<PdfVectorDragModel, "width" | "height">>,
  rect: { x: number; y: number; width: number; height: number },
  start: { width: number; height: number },
  /** Clockwise degrees the object is turned about its frame's centre. */
  rotationDegrees = 0,
): {
  percent: { x: number; y: number; width: number; height: number };
  pixels: { x: number; y: number; width: number; height: number };
  /** Set when turned: `pixels` is then the turned crop's bounding box. */
  turn: PdfVectorDragTurn | null;
  /** The axes `rect` carries mirrored (a handle pulled across). */
  mirror: { x: boolean; y: boolean } | null;
} {
  const pageWidth = Math.max(1, model.absent.width);
  const pageHeight = Math.max(1, model.absent.height);
  const flagged = rect as { mirrorX?: boolean; mirrorY?: boolean };
  const mirror =
    flagged.mirrorX || flagged.mirrorY
      ? { x: flagged.mirrorX === true, y: flagged.mirrorY === true }
      : null;
  // Mirrored, the crop's margin on one side lands on the other.
  const pad = {
    left: mirror?.x ? model.pad.right : model.pad.left,
    right: mirror?.x ? model.pad.left : model.pad.right,
    top: mirror?.y ? model.pad.bottom : model.pad.top,
    bottom: mirror?.y ? model.pad.top : model.pad.bottom,
  };
  // A resize stretches the crop's margins with the object.
  const scaleX = start.width > 0 ? rect.width / start.width : 1;
  const scaleY = start.height > 0 ? rect.height / start.height : 1;
  const percentX = (pixels: number) => ((pixels * scaleX) / pageWidth) * 100;
  const percentY = (pixels: number) => ((pixels * scaleY) / pageHeight) * 100;
  const percent = {
    x: rect.x - percentX(pad.left),
    y: rect.y - percentY(pad.top),
    width: rect.width + percentX(pad.left + pad.right),
    height: rect.height + percentY(pad.top + pad.bottom),
  };
  const crop = {
    x: (percent.x / 100) * pageWidth,
    y: (percent.y / 100) * pageHeight,
    width: (percent.width / 100) * pageWidth,
    height: (percent.height / 100) * pageHeight,
  };
  if (rotationDegrees % 360 === 0) {
    // On whole page-canvas pixels, as it is painted and as the crop was cut
    // (outward, cropPixelRect): a box between pixels made the browser
    // resample the ghost, and the text seen through a translucent shape
    // shimmered as the drag crossed pixel phases. Unstretched, the ghost is
    // exactly the crop's own size.
    const left = Math.floor(crop.x + 1e-6);
    const top = Math.floor(crop.y + 1e-6);
    const pixels = {
      x: left,
      y: top,
      width:
        scaleX === 1 && model.width
          ? model.width
          : Math.max(1, Math.ceil(crop.x + crop.width - 1e-6) - left),
      height:
        scaleY === 1 && model.height
          ? model.height
          : Math.max(1, Math.ceil(crop.y + crop.height - 1e-6) - top),
    };
    return {
      percent: {
        x: (pixels.x / pageWidth) * 100,
        y: (pixels.y / pageHeight) * 100,
        width: (pixels.width / pageWidth) * 100,
        height: (pixels.height / pageHeight) * 100,
      },
      pixels,
      turn: null,
      mirror,
    };
  }
  const centerX = ((rect.x + rect.width / 2) / 100) * pageWidth;
  const centerY = ((rect.y + rect.height / 2) / 100) * pageHeight;
  const radians = (rotationDegrees * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const corners = [
    [crop.x, crop.y],
    [crop.x + crop.width, crop.y],
    [crop.x, crop.y + crop.height],
    [crop.x + crop.width, crop.y + crop.height],
  ].map(([x, y]) => ({
    x: centerX + cos * (x - centerX) - sin * (y - centerY),
    y: centerY + sin * (x - centerX) + cos * (y - centerY),
  }));
  const left = Math.floor(Math.min(...corners.map((corner) => corner.x)));
  const top = Math.floor(Math.min(...corners.map((corner) => corner.y)));
  const right = Math.ceil(Math.max(...corners.map((corner) => corner.x)));
  const bottom = Math.ceil(Math.max(...corners.map((corner) => corner.y)));
  const pixels = { x: left, y: top, width: right - left, height: bottom - top };
  return {
    percent: {
      x: (pixels.x / pageWidth) * 100,
      y: (pixels.y / pageHeight) * 100,
      width: (pixels.width / pageWidth) * 100,
      height: (pixels.height / pageHeight) * 100,
    },
    pixels,
    turn: { degrees: rotationDegrees, centerX, centerY, crop },
    mirror: null,
  };
}
