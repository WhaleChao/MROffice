/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Dragged objects as real layers of their page, rendered from the page the
 * reader already holds. pdf.js runs the page's operator list with a filter
 * (`operationsFilter`): every state operation runs in every layer, and each
 * paint runs in exactly one -- the layer of the objects being moved, or the
 * layer of the page that is painted before, between or after them. No bytes
 * are rewritten and no file is parsed again; the renders together cost about
 * what one render of the page costs.
 *
 * A drag frame is then those layers in the page's own paint order, with the
 * moving ones carried by the frame: plain compositing of what the page really
 * paints, nothing estimated from pixel differences. It serves one source
 * vector, a selection of several (the page between them stays where it is,
 * between them) and the annotations and form widgets the page paints itself
 * (their appearance blocks move, the rest of the page and the other
 * annotations stay).
 */

import { createPdfRenderOperationsFilter } from "./pdfjs-render-operator-filter";
import type { FractionRect } from "./embedded-image-edits";
import { annotationsAtBounds, thinAxisStretch } from "./native-vector-drag-backdrop";

export { thinAxisStretch };

interface PixelBox {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/** One layer of the split, in paint order. */
export type PdfVectorDragLayer =
  | {
      readonly moving: false;
      /** The size of the page: what the page paints here, where it paints it. */
      readonly canvas: HTMLCanvasElement;
    }
  | {
      readonly moving: true;
      /** The moving objects of this layer, cut to their ink. */
      readonly canvas: HTMLCanvasElement;
      /** Where that cut sits on the page as the page paints it. */
      readonly box: PixelBox;
      /** How these objects blend with what lies under them. */
      readonly blend: GlobalCompositeOperation;
    };

export interface PdfVectorDragLayers {
  /** The page canvas's size: every layer is in its pixels. */
  readonly width: number;
  readonly height: number;
  /** Bottom first. The first is the page up to the first moving object and
   * is opaque; every other still layer is transparent where it paints
   * nothing. */
  readonly sequence: readonly PdfVectorDragLayer[];
  /** Where the moving objects stand, all together. */
  readonly objectBox: PixelBox;
}

type OperatorIds = Readonly<Record<string, number | undefined>>;

const PATH_PAINT_OPERATORS = [
  "stroke",
  "closeStroke",
  "fill",
  "eoFill",
  "fillStroke",
  "eoFillStroke",
  "closeFillStroke",
  "closeEOFillStroke",
] as const;

const OTHER_PAINT_OPERATORS = [
  "showText",
  "showSpacedText",
  "nextLineShowText",
  "nextLineSetSpacingShowText",
  "paintImageXObject",
  "paintImageXObjectRepeat",
  "paintImageMaskXObject",
  "paintImageMaskXObjectGroup",
  "paintImageMaskXObjectRepeat",
  "paintInlineImageXObject",
  "paintInlineImageXObjectGroup",
  "paintJpegXObject",
  "paintSolidColorImageMask",
  "shadingFill",
] as const;

function idSet(ops: OperatorIds, names: readonly string[]): Set<number> {
  const ids = new Set<number>();
  for (const name of names) {
    const id = ops[name];
    if (typeof id === "number") ids.add(id);
  }
  return ids;
}

/** What moves: one painted path of the page content (`start === end`), or
 * one annotation's appearance block (`beginAnnotation` to `endAnnotation`). */
export interface PdfVectorLayerMember {
  readonly start: number;
  readonly end: number;
  readonly blend: GlobalCompositeOperation;
}

/** The moving objects go into at most this many layers (each run of them
 * with nothing of the page painted in between is one); past that the page
 * is not split. Every layer in between is a canvas the size of the page. */
export const MAX_MOVING_LAYERS = 4;

const EVERY_LAYER = -1;

export interface PdfVectorLayerPlan {
  /** Per operation, the layer that runs it (`EVERY_LAYER` for the page
   * content's state). Layers are numbered in paint order: 0 still, 1 moving,
   * 2 still, ... An operation past this list belongs to the last layer. */
  readonly owner: Int16Array;
  /** How many layers there are: one still layer more than moving ones. */
  readonly layerCount: number;
  /** Per moving layer, bottom first: how it blends. */
  readonly blends: readonly GlobalCompositeOperation[];
  /** Per layer: whether any paint was given to it (a still layer with none
   * is not rendered; the first always is, it is the page's ground). */
  readonly painted: readonly boolean[];
}

/**
 * Split a page's operator list around the members: which layer runs which
 * operation. Everything of the page content that is not a paint (state,
 * transforms, clips, groups, marked content) runs in every layer, so each is
 * drawn in the state the page draws it in; an annotation's appearance block
 * is self-contained and runs whole in its one layer. Null when the page
 * cannot be split by leaving paints out: a path painted and clipped to at
 * once, text that clips (render mode 4 to 7), a member that is not a painted
 * path or an appearance block, or more runs of members than layers allowed.
 */
export function planNativeVectorLayers(
  operatorList: { readonly fnArray: readonly number[]; readonly argsArray: readonly unknown[] },
  ops: OperatorIds,
  members: readonly PdfVectorLayerMember[],
  /** Told why the page does not split, when it does not. */
  report?: (reason: string, detail?: unknown) => void,
): PdfVectorLayerPlan | null {
  const { fnArray, argsArray } = operatorList;
  const pathPaints = idSet(ops, PATH_PAINT_OPERATORS);
  const otherPaints = idSet(ops, OTHER_PAINT_OPERATORS);
  const constructPath = ops.constructPath;
  if (typeof constructPath !== "number" || pathPaints.size === 0) {
    report?.("no-operator-ids");
    return null;
  }
  if (members.length === 0) {
    report?.("no-members");
    return null;
  }
  let contentEnd = fnArray.length;
  if (typeof ops.beginAnnotation === "number") {
    const first = fnArray.indexOf(ops.beginAnnotation);
    if (first >= 0) contentEnd = first;
  }
  // The units of the page in paint order: each paint of the page content,
  // then each annotation's appearance block.
  const units: { start: number; end: number }[] = [];
  for (let index = 0; index < contentEnd; index += 1) {
    const fn = fnArray[index];
    if (fn === constructPath) {
      const args = argsArray[index];
      if (Array.isArray(args) && pathPaints.has(Number(args[0]))) {
        // Painted and made the clip at once: without the paint the clip
        // would fall to the next path.
        const before = index > 0 ? fnArray[index - 1] : -1;
        if (before === ops.clip || before === ops.eoClip) {
          report?.("painted-clip", { index });
          return null;
        }
        units.push({ start: index, end: index });
      }
    } else if (otherPaints.has(fn)) {
      units.push({ start: index, end: index });
    } else if (fn === ops.setTextRenderingMode) {
      const args = argsArray[index];
      if (Array.isArray(args) && Number(args[0]) >= 4) {
        report?.("clipping-text", { index });
        return null;
      }
    }
  }
  for (let index = contentEnd; index < fnArray.length; index += 1) {
    if (fnArray[index] !== ops.beginAnnotation) {
      // Between blocks: nothing the renderer writes, kept with what follows.
      units.push({ start: index, end: index });
      continue;
    }
    let end = index;
    while (end < fnArray.length - 1 && fnArray[end] !== ops.endAnnotation) end += 1;
    units.push({ start: index, end });
    index = end;
  }
  const memberAt = new Map<number, PdfVectorLayerMember>();
  for (const member of members) memberAt.set(member.start, member);
  const owner = new Int16Array(fnArray.length).fill(EVERY_LAYER);
  const blends: GlobalCompositeOperation[] = [];
  const painted: boolean[] = [true];
  let run = -1;
  let inRun = false;
  let matched = 0;
  for (const unit of units) {
    const member = memberAt.get(unit.start);
    let layer: number;
    if (member && member.end === unit.end) {
      matched += 1;
      // A member that blends on its own terms is a layer of its own; the
      // others share one while nothing of the page is painted between them.
      const plain = member.blend === "source-over";
      if (!inRun || !plain || blends[run] !== "source-over") {
        run += 1;
        blends.push(member.blend);
      }
      inRun = true;
      layer = 2 * run + 1;
    } else {
      inRun = false;
      layer = 2 * (run + 1);
    }
    painted[layer] = true;
    for (let index = unit.start; index <= unit.end; index += 1) owner[index] = layer;
  }
  if (matched !== members.length) {
    report?.("not-a-painted-object", { members: members.length, matched });
    return null;
  }
  const runs = run + 1;
  if (runs > MAX_MOVING_LAYERS) {
    report?.("too-many-runs", { runs });
    return null;
  }
  const layerCount = 2 * runs + 1;
  for (let layer = 0; layer < layerCount; layer += 1) painted[layer] = painted[layer] === true;
  return { owner, layerCount, blends, painted };
}

/** Whether the plan's layer `layer` runs operation `index`. */
export function nativeVectorLayerFilter(
  plan: PdfVectorLayerPlan,
  layer: number,
): (index: number) => boolean {
  const { owner, layerCount } = plan;
  const last = layerCount - 1;
  return (index) =>
    index >= owner.length
      ? // Annotation appearances the plan's list did not hold: after everything.
        layer === last
      : owner[index] === EVERY_LAYER || owner[index] === layer;
}

const BLEND_BY_PDF_MODE: Readonly<Record<string, GlobalCompositeOperation>> = {
  Normal: "source-over",
  Compatible: "source-over",
  Multiply: "multiply",
  Screen: "screen",
  Overlay: "overlay",
  Darken: "darken",
  Lighten: "lighten",
  ColorDodge: "color-dodge",
  ColorBurn: "color-burn",
  HardLight: "hard-light",
  SoftLight: "soft-light",
  Difference: "difference",
  Exclusion: "exclusion",
  Hue: "hue",
  Saturation: "saturation",
  Color: "color",
  Luminosity: "luminosity",
};

/** The canvas blend a PDF blend mode composites with. */
export function canvasBlendForPdfMode(mode: string | null | undefined): GlobalCompositeOperation {
  return (mode && BLEND_BY_PDF_MODE[mode]) || "source-over";
}

interface LayerViewport {
  readonly width: number;
  readonly height: number;
  convertToViewportRectangle?(rect: number[]): number[];
}

interface LayerPage {
  getViewport(params: { scale: number; rotation: number }): LayerViewport;
  render(params: Record<string, unknown>): { promise: Promise<unknown> };
  getAnnotations?(params: { intent: string }): Promise<unknown>;
}

/**
 * The appearance blocks of the annotations (or form widgets) the page paints
 * at `bounds`, as members to move. Null when the page paints nothing of its
 * own there (the object is drawn by the reader's overlay). `operatorList`
 * has to be the list the reader's render runs (same annotation mode and
 * storage), since blocks are found in it by annotation id.
 */
export async function annotationLayerMembersAt(input: {
  readonly page: LayerPage;
  readonly scale: number;
  readonly rotation: number;
  /** The dragged frame's box, in page fractions as displayed. */
  readonly bounds: FractionRect;
  readonly operatorList: { readonly fnArray: readonly number[]; readonly argsArray: readonly unknown[] };
  readonly ops: OperatorIds;
}): Promise<PdfVectorLayerMember[] | null> {
  const { page, operatorList, ops } = input;
  if (!page.getAnnotations || typeof ops.beginAnnotation !== "number") return null;
  const viewport = page.getViewport({ scale: input.scale, rotation: input.rotation });
  if (!viewport.convertToViewportRectangle || viewport.width <= 0 || viewport.height <= 0) {
    return null;
  }
  const annotations = (await page.getAnnotations({ intent: "display" })) as Array<{
    id?: unknown;
    rect?: unknown;
  }>;
  const boxes: Array<{ id: string; box: FractionRect }> = [];
  for (const annotation of annotations) {
    const rect = annotation.rect;
    if (typeof annotation.id !== "string" || !Array.isArray(rect) || rect.length < 4) continue;
    const [x1, y1, x2, y2] = viewport.convertToViewportRectangle(rect.slice(0, 4).map(Number));
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
  const ids = new Set(annotationsAtBounds(boxes, input.bounds));
  if (ids.size === 0) return null;
  const { fnArray, argsArray } = operatorList;
  const members: PdfVectorLayerMember[] = [];
  for (let index = 0; index < fnArray.length; index += 1) {
    if (fnArray[index] !== ops.beginAnnotation) continue;
    const args = argsArray[index];
    let end = index;
    while (end < fnArray.length - 1 && fnArray[end] !== ops.endAnnotation) end += 1;
    if (Array.isArray(args) && typeof args[0] === "string" && ids.has(args[0])) {
      members.push({ start: index, end, blend: "source-over" });
    }
    index = end;
  }
  return members.length > 0 ? members : null;
}

/** The box of everything not fully transparent in `data`, or null. */
export function inkBox(
  data: Uint8ClampedArray,
  width: number,
  height: number,
): { x: number; y: number; width: number; height: number } | null {
  let left = width;
  let top = height;
  let right = -1;
  let bottom = -1;
  for (let y = 0; y < height; y += 1) {
    const row = y * width * 4;
    for (let x = 0; x < width; x += 1) {
      if (data[row + x * 4 + 3] === 0) continue;
      if (x < left) left = x;
      if (x > right) right = x;
      if (y < top) top = y;
      if (y > bottom) bottom = y;
    }
  }
  if (right < left || bottom < top) return null;
  return { x: left, y: top, width: right - left + 1, height: bottom - top + 1 };
}

/**
 * Render the layers of `page` a plan describes, the way the reader rendered
 * the page (same viewport, same annotation mode, same layer visibility).
 * Null when a layer cannot be rendered, the moving objects leave no ink, or
 * the layers put back together are not the page the reader shows (a blend
 * this compositing cannot repeat): the caller then falls back to the model
 * built from rewritten bytes.
 */
export async function renderNativeVectorLayers(input: {
  readonly page: LayerPage;
  readonly plan: PdfVectorLayerPlan;
  /** The page canvas the reader shows: size, and what to check against. */
  readonly reference: HTMLCanvasElement;
  readonly scale: number;
  readonly rotation: number;
  readonly annotationMode: number;
  readonly optionalContentConfigPromise?: Promise<unknown>;
  readonly createCanvas?: () => HTMLCanvasElement;
  /** Told why the layers are not usable, when they are not. */
  readonly report?: (reason: string, detail?: unknown) => void;
  





  readonly extent?: Readonly<{ x: number; y: number; width: number; height: number }>;
  /** The list the plan was made from (getOperatorList) and its operator
   * ids: the plan's indices are run against pdf.js's folded render list
   * (createPdfRenderOperationsFilter). */
  readonly operatorList?: { readonly fnArray: readonly number[]; readonly argsArray: readonly unknown[] };
  readonly ops?: OperatorIds;
}): Promise<PdfVectorDragLayers | null> {
  const width = input.reference.width;
  const height = input.reference.height;
  if (width <= 0 || height <= 0) return null;
  const viewport = input.page.getViewport({ scale: input.scale, rotation: input.rotation });
  // The reader's canvas takes the viewport's whole pixels.
  if (
    (Math.floor(viewport.width) !== width || Math.floor(viewport.height) !== height) &&
    (Math.round(viewport.width) !== width || Math.round(viewport.height) !== height)
  ) {
    input.report?.("viewport-size", {
      viewport: [viewport.width, viewport.height],
      canvas: [width, height],
    });
    return null;
  }
  const make = input.createCanvas ?? (() => document.createElement("canvas"));
  const { plan } = input;
  // The moving layers' canvas: the page and the objects' whole extent.
  const firstReach = movingLayerReach(input.extent, width, height);
  const render = async (layer: number, reach: LayerReach) => {
    const moving = layer % 2 === 1;
    const canvas = make();
    canvas.width = moving ? reach.width : width;
    canvas.height = moving ? reach.height : height;
    const context = canvas.getContext("2d");
    if (!context) return null;
    const layerFilter = nativeVectorLayerFilter(plan, layer);
    const mapped = input.operatorList && input.ops
      ? createPdfRenderOperationsFilter({ list: input.operatorList, ops: input.ops, keep: layerFilter })
      : null;
    const task = input.page.render({
      canvas,
      canvasContext: context,
      viewport,
      annotationMode: input.annotationMode,
      operationsFilter: mapped ? mapped.filter : layerFilter,
      // Only the page's ground is opaque.
      ...(layer > 0 ? { background: "rgba(0,0,0,0)" } : {}),
      // Drawn shifted so what lies past the page's top or left edge lands on
      // the canvas too.
      ...(moving && (reach.x !== 0 || reach.y !== 0)
        ? { transform: [1, 0, 0, 1, -reach.x, -reach.y] }
        : {}),
      ...(input.optionalContentConfigPromise
        ? { optionalContentConfigPromise: input.optionalContentConfigPromise }
        : {}),
    });
    mapped?.bind(task);
    await task.promise;
    if (mapped && !mapped.finish()) {
      input.report?.("render-list-unaligned", { layer });
      return null;
    }
    return canvas;
  };
  // The extent is what the caller knows of the objects, and ink at the
  // canvas's edge says it knew too little: a shape turned by the user, just
  // written, is known by its unturned frame until the page's objects are read
  // again, and what stood past that frame was cut from the layer for good.
  const inkOf = (canvas: HTMLCanvasElement) => {
    const context = canvas.getContext("2d", { willReadFrequently: true });
    return context ? inkBox(context.getImageData(0, 0, canvas.width, canvas.height).data, canvas.width, canvas.height) : null;
  };
  const rendered = await Promise.all(
    Array.from({ length: plan.layerCount }, async (_unused, layer) => {
      if (!plan.painted[layer]) return null;
      const canvas = await render(layer, firstReach);
      if (!canvas || layer % 2 === 0) return canvas ? { canvas, reach: firstReach } : null;
      const ink = inkOf(canvas);
      const wider = ink ? movingLayerReachPastInk(firstReach, ink, width, height) : null;
      if (!wider) return { canvas, reach: firstReach };
      const again = await render(layer, wider);
      return again ? { canvas: again, reach: wider } : null;
    }),
  );
  const sequence: PdfVectorDragLayer[] = [];
  let union: { left: number; top: number; right: number; bottom: number } | null = null;
  for (let layer = 0; layer < plan.layerCount; layer += 1) {
    const drawn = rendered[layer];
    if (!drawn) {
      if (plan.painted[layer]) return null;
      continue;
    }
    const { canvas, reach } = drawn;
    if (layer % 2 === 0) {
      sequence.push({ moving: false, canvas });
      continue;
    }
    const context = canvas.getContext("2d", { willReadFrequently: true });
    if (!context) return null;
    const found = inkBox(
      context.getImageData(0, 0, canvas.width, canvas.height).data,
      canvas.width,
      canvas.height,
    );
    // Members that leave no ink (hidden, wholly off the page and past the
    // extent) move nothing.
    if (!found) continue;
    const cut = make();
    cut.width = found.width;
    cut.height = found.height;
    const cutContext = cut.getContext("2d");
    if (!cutContext) return null;
    cutContext.drawImage(canvas, found.x, found.y, found.width, found.height, 0, 0, found.width, found.height);
    // In page pixels: left of or above the page is negative.
    const box = { x: found.x + reach.x, y: found.y + reach.y, width: found.width, height: found.height };
    sequence.push({
      moving: true,
      canvas: cut,
      box,
      blend: plan.blends[(layer - 1) / 2] ?? "source-over",
    });
    union = union
      ? {
          left: Math.min(union.left, box.x),
          top: Math.min(union.top, box.y),
          right: Math.max(union.right, box.x + box.width),
          bottom: Math.max(union.bottom, box.y + box.height),
        }
      : { left: box.x, top: box.y, right: box.x + box.width, bottom: box.y + box.height };
  }
  if (!union) {
    input.report?.("no-ink");
    return null;
  }
  const layers: PdfVectorDragLayers = {
    width,
    height,
    sequence,
    objectBox: {
      x: union.left,
      y: union.top,
      width: union.right - union.left,
      height: union.bottom - union.top,
    },
  };
  if (!layersMatchPage(layers, input.reference, make)) {
    input.report?.("not-the-page", { box: layers.objectBox });
    return null;
  }
  return layers;
}

type LayerReach = Readonly<{ x: number; y: number; width: number; height: number }>;

/**
 * The canvas to render a moving layer on again when its ink ran to an edge of
 * `reach` that is not as far as a layer goes (a page's size past the page, as
 * movingLayerReach): that edge taken out that far. Null when no ink was cut.
 */
export function movingLayerReachPastInk(
  reach: LayerReach,
  ink: Readonly<{ x: number; y: number; width: number; height: number }>,
  width: number,
  height: number,
): LayerReach | null {
  const right = reach.x + reach.width, bottom = reach.y + reach.height;
  const left = ink.x <= 0 ? -width : reach.x;
  const top = ink.y <= 0 ? -height : reach.y;
  const farRight = ink.x + ink.width >= reach.width ? 2 * width : right;
  const farBottom = ink.y + ink.height >= reach.height ? 2 * height : bottom;
  if (left === reach.x && top === reach.y && farRight === right && farBottom === bottom) return null;
  return { x: left, y: top, width: farRight - left, height: farBottom - top };
}

/**
 * The canvas the moving layers are rendered on, in page pixels: the page,
 * grown to take in `extent` (page fractions) and a few pixels of
 * anti-aliasing, at most a page's size again on each side.
 */
export function movingLayerReach(
  extent: Readonly<{ x: number; y: number; width: number; height: number }> | undefined,
  width: number,
  height: number,
): { x: number; y: number; width: number; height: number } {
  if (!extent || ![extent.x, extent.y, extent.width, extent.height].every(Number.isFinite)) {
    return { x: 0, y: 0, width, height };
  }
  const margin = 4;
  const left = Math.max(-width, Math.min(0, Math.floor(extent.x * width) - margin));
  const top = Math.max(-height, Math.min(0, Math.floor(extent.y * height) - margin));
  const right = Math.min(2 * width, Math.max(width, Math.ceil((extent.x + extent.width) * width) + margin));
  const bottom = Math.min(2 * height, Math.max(height, Math.ceil((extent.y + extent.height) * height) + margin));
  return { x: left, y: top, width: right - left, height: bottom - top };
}

/** Put back together where the objects stand, the layers are the page. */
function layersMatchPage(
  layers: PdfVectorDragLayers,
  reference: HTMLCanvasElement,
  make: () => HTMLCanvasElement,
): boolean {
  // Only what the page shows can be compared: the part past its edge is the
  // page's to hide.
  const left = Math.max(0, layers.objectBox.x);
  const top = Math.max(0, layers.objectBox.y);
  const right = Math.min(layers.width, layers.objectBox.x + layers.objectBox.width);
  const bottom = Math.min(layers.height, layers.objectBox.y + layers.objectBox.height);
  if (right - left < 1 || bottom - top < 1) return true;
  const x = left;
  const y = top;
  const width = right - left;
  const height = bottom - top;
  const probe = make();
  probe.width = width;
  probe.height = height;
  const context = probe.getContext("2d", { willReadFrequently: true });
  const shown = reference.getContext("2d", { willReadFrequently: true });
  if (!context || !shown) return false;
  for (const layer of layers.sequence) {
    if (layer.moving) {
      context.globalCompositeOperation = layer.blend;
      context.drawImage(layer.canvas, layer.box.x - x, layer.box.y - y);
      context.globalCompositeOperation = "source-over";
    } else {
      context.drawImage(layer.canvas, x, y, width, height, 0, 0, width, height);
    }
  }
  let composed: Uint8ClampedArray;
  let page: Uint8ClampedArray;
  try {
    composed = context.getImageData(0, 0, width, height).data;
    page = shown.getImageData(x, y, width, height).data;
  } catch {
    return false;
  }
  return layerPixelsMatch(composed, page);
}

/** Whether two renders of one region are the same picture: all but a
 * hundredth of the pixels within a few levels (rounding of the split). */
export function layerPixelsMatch(composed: Uint8ClampedArray, page: Uint8ClampedArray): boolean {
  if (composed.length !== page.length || composed.length === 0) return false;
  const pixels = composed.length / 4;
  let off = 0;
  for (let p = 0; p < pixels; p += 1) {
    const i = p * 4;
    if (
      Math.abs(composed[i] - page[i]) > 12 ||
      Math.abs(composed[i + 1] - page[i + 1]) > 12 ||
      Math.abs(composed[i + 2] - page[i + 2]) > 12
    ) {
      off += 1;
    }
  }
  return off <= pixels * 0.01;
}

const thinLayerCores = new WeakMap<HTMLCanvasElement, Readonly<{ x: number; y: number }>>();

/**
 * The row ("y") or column ("x") of a thin layer that the stroke itself
 * paints: the most opaque one. Its middle was the faint anti-aliased row of
 * a hairline two pixels deep, and the bar it was stretched into came out
 * the page's colour.
 */
function thinLayerCore(canvas: HTMLCanvasElement, axis: "x" | "y"): number {
  const middle = axis === "y" ? Math.floor(canvas.height / 2) : Math.floor(canvas.width / 2);
  let known = thinLayerCores.get(canvas);
  if (!known) {
    let pixels: Uint8ClampedArray | null = null;
    try {
      pixels = canvas.getContext("2d", { willReadFrequently: true })?.getImageData(0, 0, canvas.width, canvas.height).data ?? null;
    } catch {
      pixels = null;
    }
    if (!pixels) return middle;
    const rows = new Float64Array(canvas.height);
    const columns = new Float64Array(canvas.width);
    for (let y = 0; y < canvas.height; y += 1) {
      for (let x = 0; x < canvas.width; x += 1) {
        const alpha = pixels[(y * canvas.width + x) * 4 + 3];
        rows[y] += alpha;
        columns[x] += alpha;
      }
    }
    const densest = (sums: Float64Array, fallback: number) => {
      let best = fallback;
      for (let index = 0; index < sums.length; index += 1) if (sums[index] > sums[best]) best = index;
      return best;
    };
    known = Object.freeze({
      x: densest(columns, Math.floor(canvas.width / 2)),
      y: densest(rows, Math.floor(canvas.height / 2)),
    });
    thinLayerCores.set(canvas, known);
  }
  return axis === "y" ? known.y : known.x;
}

/**
 * Paint one drag frame into `stage` (the page canvas's size, exactly over
 * it, never moving): where the objects were and where they now are, the
 * layers in the page's paint order, the moving ones carried by the frame. A
 * move copies them at whole pixels; a resize, a turn or a mirror goes
 * through the canvas's transform about the frame's centre.
 */
export function paintNativeVectorLayerStage(
  stage: HTMLCanvasElement,
  layers: PdfVectorDragLayers,
  /** The objects' frame when the gesture began, and now, in page pixels. */
  start: { x: number; y: number; width: number; height: number },
  target: { x: number; y: number; width: number; height: number },
  options: {
    rotationDegrees?: number;
    mirror?: Readonly<{ x: boolean; y: boolean }> | null;
    /** More places the page may paint the objects at (an earlier drop
     * whose write landed while this one waits), covered by the page
     * without them as well. */
    covers?: readonly { x: number; y: number; width: number; height: number }[];
    /** The whole motion as one page-pixel affine `[a, b, c, d, e, f]`
     * (x' = a x + c y + e): a turned frame resized along its own axes is a
     * scale no turn-and-scale about the centre can say. Wins over the rest. */
    matrix?: readonly [number, number, number, number, number, number] | null;
  } = {},
): boolean {
  const resized = stage.width !== layers.width || stage.height !== layers.height;
  if (resized) {
    stage.width = layers.width;
    stage.height = layers.height;
  }
  const context = stage.getContext("2d");
  if (!context) return false;
  if (!resized) context.clearRect(0, 0, layers.width, layers.height);
  const box = layers.objectBox;
  const rotation = options.rotationDegrees ?? 0;
  const mirrorX = options.mirror?.x === true;
  const mirrorY = options.mirror?.y === true;
  // A frame with no extent on one axis keeps that axis. A hairline's frame
  // is its stroke, half a pixel deep at 100%: read as no extent, a rule
  // pulled into a bar only moved on the stage, a thin line in the middle of
  // its new frame, until the page painted the bar.
  const scaleX = start.width > 1e-3 ? target.width / start.width : 1;
  const scaleY = start.height > 1e-3 ? target.height / start.height : 1;
  const matrix = options.matrix ?? null;
  const moved =
    !matrix &&
    rotation % 360 === 0 &&
    !mirrorX &&
    !mirrorY &&
    Math.abs(scaleX - 1) < 1e-6 &&
    Math.abs(scaleY - 1) < 1e-6;
  const startCentre = { x: start.x + start.width / 2, y: start.y + start.height / 2 };
  const targetCentre = { x: target.x + target.width / 2, y: target.y + target.height / 2 };
  const dx = Math.round(targetCentre.x - startCentre.x);
  const dy = Math.round(targetCentre.y - startCentre.y);
  const radians = (rotation * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const signedX = scaleX * (mirrorX ? -1 : 1);
  const signedY = scaleY * (mirrorY ? -1 : 1);
  // Where the objects land: their box's corners through the frame's transform.
  const landed = moved
    ? { x: box.x + dx, y: box.y + dy, width: box.width, height: box.height }
    : (() => {
        const corners = [
          [box.x, box.y],
          [box.x + box.width, box.y],
          [box.x, box.y + box.height],
          [box.x + box.width, box.y + box.height],
        ].map(([x, y]) => {
          if (matrix) {
            return {
              x: matrix[0] * x + matrix[2] * y + matrix[4],
              y: matrix[1] * x + matrix[3] * y + matrix[5],
            };
          }
          const u = (x - startCentre.x) * signedX;
          const v = (y - startCentre.y) * signedY;
          return {
            x: targetCentre.x + cos * u - sin * v,
            y: targetCentre.y + sin * u + cos * v,
          };
        });
        const xs = corners.map((corner) => corner.x);
        const ys = corners.map((corner) => corner.y);
        const left = Math.floor(Math.min(...xs)) - 1;
        const top = Math.floor(Math.min(...ys)) - 1;
        return {
          x: left,
          y: top,
          width: Math.ceil(Math.max(...xs)) + 1 - left,
          height: Math.ceil(Math.max(...ys)) + 1 - top,
        };
      })();
  // A thin layer stretched across itself is drawn over the frame's whole
  // depth (below): where it lands is that depth, not its own few rows carried
  // by the scale, which cut the bar short of its frame.
  const stretched = matrix ? null : thinAxisStretch(box, rotation, mirrorX, mirrorY, scaleX, scaleY);
  const reach = stretched === "y"
    ? { x: landed.x, y: Math.floor(target.y) - 1, width: landed.width, height: Math.ceil(target.height) + 2 }
    : stretched === "x"
      ? { x: Math.floor(target.x) - 1, y: landed.y, width: Math.ceil(target.width) + 2, height: landed.height }
      : null;
  context.save();
  // Only where the objects were and where they are: the page shows the rest.
  // Two pixels past the old ink: the browser resamples the stage and the
  // page canvas to the page's fractional CSS size, and at an edge where the
  // stage stops the old ink under it bled through as a grey outline.
  context.beginPath();
  context.rect(box.x - 2, box.y - 2, box.width + 4, box.height + 4);
  // A stretch lands on the frame's depth alone: the ink's rows carried by
  // the scale (a hundred times a hairline) reach far past it.
  if (reach) context.rect(reach.x, reach.y, reach.width, reach.height);
  else context.rect(landed.x, landed.y, landed.width, landed.height);
  for (const cover of options.covers ?? []) {
    // A frame box: the stroke and its anti-aliasing reach a little past it.
    context.rect(Math.floor(cover.x) - 2, Math.floor(cover.y) - 2, Math.ceil(cover.width) + 4, Math.ceil(cover.height) + 4);
  }
  context.clip();
  for (const layer of layers.sequence) {
    if (!layer.moving) {
      context.drawImage(layer.canvas, 0, 0);
      continue;
    }
    context.save();
    context.globalCompositeOperation = layer.blend;
    if (moved) {
      context.drawImage(layer.canvas, layer.box.x + dx, layer.box.y + dy);
    } else if (matrix) {
      context.transform(matrix[0], matrix[1], matrix[2], matrix[3], matrix[4], matrix[5]);
      context.drawImage(layer.canvas, layer.box.x, layer.box.y);
    } else {
      const strip = thinAxisStretch(layer.box, rotation, mirrorX, mirrorY, scaleX, scaleY);
      if (strip) {
        // A line pulled across itself grows thicker, not taller: the frame
        // scales one axis of an object only a few pixels deep on it. Scaling
        // its bitmap smeared the two anti-aliased edge rows into a soft
        
        // repeated over the new depth, is what a thicker stroke looks like.
        // The bar is as deep as the frame it is pulled to: the page paints
        // the stroke over exactly that depth. The layer's own depth (the
        // hairline and its anti-aliased rows, several times its stroke at a
        // high zoom) times the scale drew a bar taller than its frame.
        const x = targetCentre.x + (layer.box.x - startCentre.x) * scaleX;
        const y = targetCentre.y + (layer.box.y - startCentre.y) * scaleY;
        const core = thinLayerCore(layer.canvas, strip);
        if (strip === "y") {
          context.drawImage(
            layer.canvas,
            0, core, layer.box.width, 1,
            Math.round(x), target.y, layer.box.width * scaleX, target.height,
          );
        } else {
          context.drawImage(
            layer.canvas,
            core, 0, 1, layer.box.height,
            target.x, Math.round(y), target.width, layer.box.height * scaleY,
          );
        }
      } else {
        context.translate(targetCentre.x, targetCentre.y);
        context.rotate(radians);
        context.scale(signedX, signedY);
        context.translate(-startCentre.x, -startCentre.y);
        context.drawImage(layer.canvas, layer.box.x, layer.box.y);
      }
    }
    context.restore();
  }
  context.restore();
  return true;
}

/**
 * The page over `region` (page canvas pixels) as it will paint once its
 * moving objects stand `dx`, `dy` pixels away and are drawn `alphaScale`
 * times as opaque: what a picture's Transparency or Position write will
 * show, put together from the layers already in hand while the write saves.
 */
export function composeNativeVectorLayersRegion(
  layers: PdfVectorDragLayers,
  region: PixelBox,
  change: {
    readonly dx?: number;
    readonly dy?: number;
    readonly alphaScale?: number;
    /** A new drawing in place of the moving objects, filling `box` at
     * `alpha` (a format written on a translucent picture). */
    readonly replace?: Readonly<{ canvas: HTMLCanvasElement; box: PixelBox; alpha: number; turn?: number }>;
  },
  make: () => HTMLCanvasElement = () => document.createElement("canvas"),
): HTMLCanvasElement | null {
  const x = Math.floor(region.x);
  const y = Math.floor(region.y);
  const width = Math.ceil(region.x + region.width) - x;
  const height = Math.ceil(region.y + region.height) - y;
  if (width < 1 || height < 1 || width > 8192 || height > 8192) return null;
  const out = make();
  out.width = width;
  out.height = height;
  const context = out.getContext("2d");
  if (!context) return null;
  const dx = change.dx ?? 0;
  const dy = change.dy ?? 0;
  const alphaScale = change.alphaScale ?? 1;
  for (const layer of layers.sequence) {
    if (!layer.moving) {
      context.drawImage(layer.canvas, x, y, width, height, 0, 0, width, height);
      continue;
    }
    if (change.replace) {
      context.save();
      context.globalCompositeOperation = layer.blend;
      context.globalAlpha = Math.max(0, Math.min(1, change.replace.alpha));
      const box = change.replace.box;
      // Turned about its centre, as the frame shows it.
      context.translate(box.x - x + box.width / 2, box.y - y + box.height / 2);
      if (change.replace.turn) context.rotate((change.replace.turn * Math.PI) / 180);
      context.drawImage(change.replace.canvas, -box.width / 2, -box.height / 2, box.width, box.height);
      context.restore();
      continue;
    }
    let source: HTMLCanvasElement = layer.canvas;
    if (Math.abs(alphaScale - 1) > 1e-6) {
      // Scaled per pixel, not by globalAlpha: a picture made more opaque
      // than it is cannot be drawn by an alpha above one.
      const scaled = make();
      scaled.width = layer.canvas.width;
      scaled.height = layer.canvas.height;
      const scaledContext = scaled.getContext("2d", { willReadFrequently: true });
      const own = layer.canvas.getContext("2d", { willReadFrequently: true });
      if (!scaledContext || !own || scaled.width < 1 || scaled.height < 1) return null;
      let pixels: ImageData;
      try {
        pixels = own.getImageData(0, 0, scaled.width, scaled.height);
      } catch {
        return null;
      }
      const data = pixels.data;
      for (let i = 3; i < data.length; i += 4) data[i] = Math.min(255, Math.round(data[i] * alphaScale));
      scaledContext.putImageData(pixels, 0, 0);
      source = scaled;
    }
    context.save();
    context.globalCompositeOperation = layer.blend;
    context.drawImage(source, layer.box.x + dx - x, layer.box.y + dy - y);
    context.restore();
  }
  return out;
}
