/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { PdfNativeTransactionError } from "./transaction-error";
import type { PdfNativeTransactionErrorCode } from "./transaction-error";
import { MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_SCHEMA_VERSION } from "./vocabulary";
import type { PdfNativeEmbeddedImageFrameClip, PdfNativeEmbeddedImageObjectRef, PdfNativeEmbeddedImageTransformClip, PdfNativeEmbeddedImageTransformEvidence, PdfNativeEmbeddedImageTransformFrame, PdfNativeEmbeddedImageTransformMatrix } from "./content-stream-transaction";

export const PDF_RESOURCE_NAME = /^[A-Za-z][A-Za-z0-9_.+\-]{0,63}$/;
/** A name a page's own resources already use, read and written back as it
 * is: any PDF name of regular characters (TikZ names its graphics states
 * /pgf@ca1.0 -- every TeX paper with a pgf figure -- and the narrow form
 * above, for names this engine makes, refused the whole page). A name with
 * a #xx escape stays out. */
export const PDF_EXISTING_RESOURCE_NAME = /^[A-Za-z0-9!"$&'*+,\-.:;=?@^_`|~]{1,127}$/;
/** The same name inside a content-stream walk's pattern (no anchors): a
 * font a page's resources name `/TiRoARRN~1268702012` (a journal's
 * typesetter) is set with `Tf` under that name, and the narrow form left
 * every run of it out of the walk. */
export const PDF_EXISTING_RESOURCE_NAME_SOURCE = PDF_EXISTING_RESOURCE_NAME.source.slice(1, -1);

export function fail(code: PdfNativeTransactionErrorCode, message: string): never {
  throw new PdfNativeTransactionError(code, message);
}

export const EMBEDDED_IMAGE_TRANSFORM_NUMBER = "[-+]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)";

export const EMBEDDED_IMAGE_TRANSFORM_OPERATOR = new RegExp(
  `^q (?:/([A-Za-z][A-Za-z0-9_.+\\-]{0,63}) gs )?` +
    `((?:${EMBEDDED_IMAGE_TRANSFORM_NUMBER} ){6})cm ` +
    `(?:((?:${EMBEDDED_IMAGE_TRANSFORM_NUMBER} ){4})re W n )?` +
    `/([A-Za-z][A-Za-z0-9_.+\\-]{0,63}) Do Q$`,
);

export const EMBEDDED_IMAGE_TRANSFORM_MAX_OPERATOR_BYTES = 512;

// Google Docs (Skia) writes each picture as its own clipped block,
//   q <x y w h> re W* n q <a b c d e f> cm <colour / gs state> /X Do Q Q
// with the clip rectangle exactly the picture's box. Moving the draw alone
// leaves that clip where the picture was, so only the part still inside it

// the rectangle, which follows the picture; every byte between the rectangle
// and the draw stays as written.
export const EMBEDDED_IMAGE_FRAME_WS = "[ \\t\\r\\n\\f\\0]+";

export const EMBEDDED_IMAGE_FRAME_STATE =
  `(?:(?:${EMBEDDED_IMAGE_TRANSFORM_NUMBER}${EMBEDDED_IMAGE_FRAME_WS}){1,4}(?:RG|rg|g|G|k|K)` +
  `|/[A-Za-z][A-Za-z0-9_.+\\-]{0,63}${EMBEDDED_IMAGE_FRAME_WS}gs)`;

// A producer that cuts a picture to a shape (a tilted icon, Mozilla's
// images.pdf) clips its block once more, between the rectangle and the
// draw, by one or more paths: `x y m x y l ... h W n`. They are the picture's

// tilted clip where it was and showed only a sliver).
export const EMBEDDED_IMAGE_FRAME_PATHS =
  `((?:${EMBEDDED_IMAGE_FRAME_WS}(?:(?:${EMBEDDED_IMAGE_TRANSFORM_NUMBER}${EMBEDDED_IMAGE_FRAME_WS})*(?:m|l|c|v|y|h|re)` +
  `${EMBEDDED_IMAGE_FRAME_WS})+W\\*?${EMBEDDED_IMAGE_FRAME_WS}n)*)`;

// The block's state may be set before its `cm` as well as after it: the
// IPCC report's signatures are drawn `re W n q /GS0 gs <cm> /X3 Do Q Q`, and
// a frame not read as the picture's own stayed behind when the picture moved

export const EMBEDDED_IMAGE_FRAME_MIDDLE =
  `${EMBEDDED_IMAGE_FRAME_WS}re${EMBEDDED_IMAGE_FRAME_WS}W\\*?${EMBEDDED_IMAGE_FRAME_WS}n${EMBEDDED_IMAGE_FRAME_PATHS}` +
  `${EMBEDDED_IMAGE_FRAME_WS}q${EMBEDDED_IMAGE_FRAME_WS}(?:${EMBEDDED_IMAGE_FRAME_STATE}${EMBEDDED_IMAGE_FRAME_WS})*` +
  `((?:${EMBEDDED_IMAGE_TRANSFORM_NUMBER}${EMBEDDED_IMAGE_FRAME_WS}){6})cm` +
  `${EMBEDDED_IMAGE_FRAME_WS}(?:${EMBEDDED_IMAGE_FRAME_STATE}${EMBEDDED_IMAGE_FRAME_WS})*`;

export const EMBEDDED_IMAGE_FRAME_CLIP = new RegExp(
  `^(${EMBEDDED_IMAGE_TRANSFORM_NUMBER}${EMBEDDED_IMAGE_FRAME_WS}${EMBEDDED_IMAGE_TRANSFORM_NUMBER}` +
    `${EMBEDDED_IMAGE_FRAME_WS}${EMBEDDED_IMAGE_TRANSFORM_NUMBER}${EMBEDDED_IMAGE_FRAME_WS}${EMBEDDED_IMAGE_TRANSFORM_NUMBER})` +
    `(${EMBEDDED_IMAGE_FRAME_MIDDLE})([^\\s][\\s\\S]*)$`,
);

/** Split a transform span that starts with a Skia frame clip; null when it
 * is a bare draw or wrapper alone. */
export function parsePdfNativeEmbeddedImageFrameClip(
  text: string,
): PdfNativeEmbeddedImageFrameClip | null {
  const match = EMBEDDED_IMAGE_FRAME_CLIP.exec(text);
  if (!match) return null;
  const rect = match[1].split(/[ \t\r\n\f\0]+/).map(Number);
  const inner = match[4].trim().split(/[ \t\r\n\f\0]+/).map(Number);
  if (
    rect.length !== 4 ||
    inner.length !== 6 ||
    [...rect, ...inner].some((value) => !Number.isFinite(value)) ||
    Math.abs(inner[0] * inner[3] - inner[1] * inner[2]) < 1e-9
  ) {
    return null;
  }
  return Object.freeze({
    rect: Object.freeze([rect[0], rect[1], rect[2], rect[3]]) as readonly [number, number, number, number],
    rectText: match[1],
    middle: match[2],
    innerMatrix: Object.freeze([inner[0], inner[1], inner[2], inner[3], inner[4], inner[5]]) as PdfNativeEmbeddedImageTransformMatrix,
    ...(match[3] ? { clipPaths: match[3] } : {}),
    drawText: match[5],
  });
}

export function decodeLatin1(bytes: Uint8Array): string {
  let text = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    text += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return text;
}

export function exactEmbeddedImageTransformMatrix(
  value: unknown,
): PdfNativeEmbeddedImageTransformMatrix {
  if (
    !Array.isArray(value) ||
    value.length !== 6 ||
    value.some((item) => typeof item !== "number" || !Number.isFinite(item))
  ) {
    fail("invalid-request", "Embedded image transform matrix is invalid.");
  }
  const [a, b, c, d, e, f] = value as number[];
  if (Math.abs(a * d - b * c) < 1e-9) {
    fail("invalid-request", "Embedded image transform matrix is degenerate.");
  }
  return Object.freeze([a, b, c, d, e, f]) as PdfNativeEmbeddedImageTransformMatrix;
}

export function exactEmbeddedImageTransformClip(
  value: unknown,
): PdfNativeEmbeddedImageTransformClip | null {
  if (value === null || value === undefined) return null;
  if (
    typeof value !== "object" ||
    Array.isArray(value) ||
    ["x", "y", "width", "height"].some((key) => {
      const item = (value as Record<string, unknown>)[key];
      return typeof item !== "number" || !Number.isFinite(item);
    })
  ) {
    fail("invalid-request", "Embedded image transform clip is invalid.");
  }
  const clip = value as PdfNativeEmbeddedImageTransformClip;
  if (
    clip.x < 0 ||
    clip.y < 0 ||
    clip.width <= 0 ||
    clip.height <= 0 ||
    clip.x + clip.width > 1 + 1e-6 ||
    clip.y + clip.height > 1 + 1e-6 ||
    (clip.x === 0 && clip.y === 0 && clip.width >= 1 - 1e-9 && clip.height >= 1 - 1e-9)
  ) {
    fail("invalid-request", "Embedded image transform clip is out of range.");
  }
  return Object.freeze({
    x: clip.x,
    y: clip.y,
    width: clip.width,
    height: clip.height,
  });
}

export function parsePdfNativeEmbeddedImageTransformOperator(
  text: string,
): Readonly<{
  resourceName: string;
  matrix: PdfNativeEmbeddedImageTransformMatrix;
  clip: PdfNativeEmbeddedImageTransformClip | null;
  graphicsStateName: string | null;
}> | null {
  const match = EMBEDDED_IMAGE_TRANSFORM_OPERATOR.exec(text);
  if (!match) return null;
  const numbers = (raw: string) => raw.trim().split(" ").map(Number);
  const matrix = numbers(match[2]);
  const clipNumbers = match[3] ? numbers(match[3]) : null;
  if (
    matrix.some((value) => !Number.isFinite(value)) ||
    (clipNumbers && clipNumbers.some((value) => !Number.isFinite(value)))
  ) {
    return null;
  }
  return Object.freeze({
    resourceName: match[4],
    graphicsStateName: match[1] ?? null,
    matrix: Object.freeze(matrix) as unknown as PdfNativeEmbeddedImageTransformMatrix,
    clip: clipNumbers
      ? Object.freeze({
          x: clipNumbers[0],
          y: clipNumbers[1],
          width: clipNumbers[2],
          height: clipNumbers[3],
        })
      : null,
  });
}

export function readPdfNativeEmbeddedImageTransformEvidence(
  value: unknown,
): PdfNativeEmbeddedImageTransformEvidence | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const ref = (item: unknown): PdfNativeEmbeddedImageObjectRef | null => {
    if (!item || typeof item !== "object" || Array.isArray(item)) return null;
    const candidate = item as Record<string, unknown>;
    if (
      !Number.isSafeInteger(candidate.objectNumber) ||
      (candidate.objectNumber as number) < 1 ||
      !Number.isSafeInteger(candidate.generation) ||
      (candidate.generation as number) < 0
    ) {
      return null;
    }
    return Object.freeze({
      objectNumber: candidate.objectNumber as number,
      generation: candidate.generation as number,
    });
  };
  const bytes = (item: unknown): readonly number[] | null =>
    Array.isArray(item) &&
    item.length >= 4 &&
    item.length <= EMBEDDED_IMAGE_TRANSFORM_MAX_OPERATOR_BYTES &&
    item.every(
      (entry) => Number.isInteger(entry) && entry >= 0 && entry <= 255,
    )
      ? Object.freeze([...(item as number[])])
      : null;
  const page = ref(record.page);
  const resourceOwner = ref(record.resourceOwner);
  const image = ref(record.image);
  const contentStream = ref(record.contentStream);
  const originalOperatorBytes = bytes(record.originalOperatorBytes);
  const transformedOperatorBytes = bytes(record.transformedOperatorBytes);
  if (
    record.schemaVersion !== MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_SCHEMA_VERSION ||
    !Number.isSafeInteger(record.pageNumber) ||
    (record.pageNumber as number) < 1 ||
    !Number.isSafeInteger(record.imageOrder) ||
    (record.imageOrder as number) < 0 ||
    typeof record.resourceName !== "string" ||
    !PDF_RESOURCE_NAME.test(record.resourceName) ||
    !page ||
    !resourceOwner ||
    !image ||
    !contentStream ||
    !originalOperatorBytes ||
    !transformedOperatorBytes
  ) {
    return null;
  }
  let matrix: PdfNativeEmbeddedImageTransformMatrix;
  let clip: PdfNativeEmbeddedImageTransformClip | null;
  try {
    matrix = exactEmbeddedImageTransformMatrix(record.matrix);
    clip = exactEmbeddedImageTransformClip(record.clip ?? null);
  } catch {
    return null;
  }
  const transformedText = decodeLatin1(Uint8Array.from(transformedOperatorBytes));
  const parsed = parsePdfNativeEmbeddedImageTransformOperator(
    parsePdfNativeEmbeddedImageFrameClip(transformedText)?.drawText ?? transformedText,
  );
  if (!parsed || parsed.resourceName !== record.resourceName) return null;
  const graphicsStateName =
    record.graphicsStateName === undefined || record.graphicsStateName === null
      ? null
      : typeof record.graphicsStateName === "string" &&
          PDF_RESOURCE_NAME.test(record.graphicsStateName)
        ? record.graphicsStateName
        : undefined;
  const opacity =
    record.opacity === undefined || record.opacity === null
      ? null
      : typeof record.opacity === "number" &&
          Number.isFinite(record.opacity) &&
          record.opacity >= 0 &&
          record.opacity <= 1
        ? record.opacity
        : undefined;
  if (
    graphicsStateName === undefined ||
    opacity === undefined ||
    // The wrapper names exactly the state the evidence claims.
    parsed.graphicsStateName !== graphicsStateName ||
    (opacity !== null && graphicsStateName === null)
  ) {
    return null;
  }
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_SCHEMA_VERSION,
    pageNumber: record.pageNumber as number,
    imageOrder: record.imageOrder as number,
    resourceName: record.resourceName,
    page,
    resourceOwner,
    image,
    contentStream,
    originalOperatorBytes,
    transformedOperatorBytes,
    matrix,
    clip,
    ...(graphicsStateName ? { graphicsStateName } : {}),
    ...(opacity === null ? {} : { opacity }),
  });
}

/** `first` then `second`, in PDF's row-vector convention (`cm` order). */
export function concatPdfNativeEmbeddedImageTransformMatrices(
  first: PdfNativeEmbeddedImageTransformMatrix,
  second: PdfNativeEmbeddedImageTransformMatrix,
): PdfNativeEmbeddedImageTransformMatrix {
  const [a1, b1, c1, d1, e1, f1] = first;
  const [a2, b2, c2, d2, e2, f2] = second;
  return Object.freeze([
    a1 * a2 + b1 * c2,
    a1 * b2 + b1 * d2,
    c1 * a2 + d1 * c2,
    c1 * b2 + d1 * d2,
    e1 * a2 + f1 * c2 + e2,
    e1 * b2 + f1 * d2 + f2,
  ]) as PdfNativeEmbeddedImageTransformMatrix;
}

/**
 * Whether the CTM the native walk found at a paint is the one a reader
 * measured for the picture on screen. They are read from the same numbers,
 * but pdf.js hands a form's /Matrix over as a Float32Array, so they agree to
 * float32 precision rather than exactly: closer than a ten-thousandth of the
 * matrix's own size is the same placement (two different pictures that close
 * are drawn in the same place).
 */
export function pdfNativeEmbeddedImagePaintMatrixMatches(
  paint: readonly number[] | null | undefined,
  expected: readonly number[] | null | undefined,
): boolean {
  if (!paint || !expected || paint.length !== 6 || expected.length !== 6) return false;
  if (![...paint, ...expected].every((value) => Number.isFinite(value))) return false;
  const tolerance =
    1e-4 * Math.max(1, ...paint.map(Math.abs), ...expected.map(Math.abs));
  return paint.every((value, index) => Math.abs(value - expected[index]) <= tolerance);
}

/**
 * The writers' `imageOrder` for a picture a reader measured: its own order
 * when the paint there is drawn where the reader saw the picture, else the
 * one paint on the page drawn there. A reader counts paints the writers step
 * over or cannot reach (pdf.js lists an inline image, a picture inside a form
 * drawn on every page, a run of repeated stamps merged into one operator), so
 * the order alone can name another picture. Null when no single paint is
 * drawn where the picture is: the write is then refused, never aimed.
 */
export function resolvePdfNativeEmbeddedImageOrder(
  paints: readonly (readonly number[] | null)[],
  target: { readonly order: number; readonly matrix?: readonly number[] | null },
): number | null {
  const expected = target.matrix;
  if (!expected) return null;
  if (pdfNativeEmbeddedImagePaintMatrixMatches(paints[target.order], expected)) {
    return target.order;
  }
  const matches = paints.flatMap((paint, index) =>
    pdfNativeEmbeddedImagePaintMatrixMatches(paint, expected) ? [index] : [],
  );
  return matches.length === 1 ? matches[0] : null;
}

/**
 * The paint a queued delete or new image of a picture is made on when the
 * other window changed that very picture (no paint is drawn where the edit
 * found it): `onHead` are the page's paints on the version that window
 * saved, `onPlanned` the same page as this edit left it when it was first
 * planned, `order` the picture's order when the edit was made. It is the
 * paint at that order when every other paint of the page is where this edit
 * left it -- the lists differ in that paint alone, which a delete took away
 * and a new image left where it was. `already-there`: the page is as the
 * delete left it (the other window deleted the picture too). Where the
 * paints lie does not say which image a paint draws: the caller holds the
 * paint to the image the edit was made on before it writes
 * (readMraiNativeEmbeddedImageObject). Null whenever the lists differ
 * otherwise.
 */
export function resolvePdfNativeEmbeddedImageLaterSaveOrder(input: Readonly<{
  onHead: readonly (readonly number[] | null)[];
  onPlanned: readonly (readonly number[] | null)[];
  order: number;
  /** The edit took the paint away: the page as it left it has one fewer. */
  removes: boolean;
}>): number | "already-there" | null {
  const { onHead, onPlanned, order } = input;
  if (!Number.isSafeInteger(order) || order < 0) return null;
  const same = (left: readonly number[] | null | undefined, right: readonly number[] | null | undefined) =>
    left === null || left === undefined ? right === null || right === undefined
      : pdfNativeEmbeddedImagePaintMatrixMatches(right, left);
  if (input.removes && onPlanned.length === onHead.length) {
    return onHead.every((paint, index) => same(paint, onPlanned[index])) ? "already-there" : null;
  }
  if (onPlanned.length !== onHead.length - (input.removes ? 1 : 0) || !onHead[order]) return null;
  return onHead.every((paint, index) => index === order ||
    same(paint, onPlanned[input.removes && index > order ? index - 1 : index])) ? order : null;
}

export const UNIT_TURN_MATRICES: readonly PdfNativeEmbeddedImageTransformMatrix[] = [
  [1, 0, 0, 1, 0, 0],
  // 90° clockwise on screen: (u, v) -> (v, 1 - u)
  [0, -1, 1, 0, 0, 1],
  // 180°: (u, v) -> (1 - u, 1 - v)
  [-1, 0, 0, -1, 1, 1],
  // 270° clockwise (90° counter-clockwise): (u, v) -> (1 - v, u)
  [0, 1, -1, 0, 1, 0],
];

export const PAGE_FLIP_X_MATRIX: PdfNativeEmbeddedImageTransformMatrix = [-1, 0, 0, 1, 0, 0];

export const PAGE_FLIP_Y_MATRIX: PdfNativeEmbeddedImageTransformMatrix = [1, 0, 0, -1, 0, 0];

export function applyEmbeddedImageTransformMatrix(
  matrix: PdfNativeEmbeddedImageTransformMatrix,
  x: number,
  y: number,
): readonly [number, number] {
  return [
    matrix[0] * x + matrix[2] * y + matrix[4],
    matrix[1] * x + matrix[3] * y + matrix[5],
  ];
}

export function embeddedImageTransformBounds(
  matrix: PdfNativeEmbeddedImageTransformMatrix,
  region: PdfNativeEmbeddedImageTransformFrame,
): PdfNativeEmbeddedImageTransformFrame {
  const corners = [
    applyEmbeddedImageTransformMatrix(matrix, region.x, region.y),
    applyEmbeddedImageTransformMatrix(matrix, region.x + region.width, region.y),
    applyEmbeddedImageTransformMatrix(matrix, region.x, region.y + region.height),
    applyEmbeddedImageTransformMatrix(
      matrix,
      region.x + region.width,
      region.y + region.height,
    ),
  ];
  const xs = corners.map(([x]) => x);
  const ys = corners.map(([, y]) => y);
  const x = Math.min(...xs);
  const y = Math.min(...ys);
  const width = Math.max(...xs) - x;
  const height = Math.max(...ys) - y;
  if (![x, y, width, height].every(Number.isFinite) || !(width > 0) || !(height > 0)) {
    fail("invalid-request", "Embedded image transform bounds are degenerate.");
  }
  return Object.freeze({ x, y, width, height });
}

/**
 * Word's Crop > Fill, Fit and Reset Crop on a picture this editor cropped
 * (its wrapper's clip, in the image's unit square):
 * - fill: the picture scaled, proportions kept, until it covers the cropped
 *   area, centred on it; the area stays the crop;
 * - fit: the picture scaled, proportions kept, until all of it shows inside
 *   the cropped area, centred on it; nothing is cropped;
 * - reset: the crop taken away, the picture where and as large as it is.
 * Each is one placement of the whole picture: a picture nothing crops has
 * nothing to fill, fit or reset (the no-op is refused like any other).
 */
export type PdfNativeEmbeddedImageCropMode = "fill" | "fit" | "reset";

function pdfNativeEmbeddedImageCropModePlacement(input: {
  readonly ctm: PdfNativeEmbeddedImageTransformMatrix;
  readonly existingClip: PdfNativeEmbeddedImageTransformClip | null;
  readonly quarterTurns: 0 | 1 | 2 | 3;
  readonly flipX: boolean;
  readonly flipY: boolean;
  readonly crop?: PdfNativeEmbeddedImageTransformFrame | null;
  readonly pageTransform?: PdfNativeEmbeddedImageTransformMatrix | null;
  readonly cropMode?: PdfNativeEmbeddedImageCropMode | null;
}): Readonly<{
  placement: PdfNativeEmbeddedImageTransformMatrix;
  clip: PdfNativeEmbeddedImageTransformClip | null;
}> {
  const mode = input.cropMode;
  if (mode !== "fill" && mode !== "fit" && mode !== "reset") {
    fail("invalid-request", "Embedded image crop mode is invalid.");
  }
  if (input.crop || input.pageTransform || input.quarterTurns !== 0 || input.flipX || input.flipY) {
    fail("invalid-request", "Embedded image crop mode cannot be combined with another change.");
  }
  const ctm = exactEmbeddedImageTransformMatrix([...input.ctm]);
  const area = input.existingClip;
  if (mode === "reset" || !area) return Object.freeze({ placement: ctm, clip: null });
  // The picture is the unit square: scaled by `scale` about the area's
  // centre it covers (fill) or fits in (fit) the area, whatever the paint's
  // own proportions, since both axes scale alike.
  const scale = mode === "fill" ? Math.max(area.width, area.height) : Math.min(area.width, area.height);
  const offsetX = area.x + area.width / 2 - scale / 2;
  const offsetY = area.y + area.height / 2 - scale / 2;
  const placement = exactEmbeddedImageTransformMatrix([
    ...concatPdfNativeEmbeddedImageTransformMatrices([scale, 0, 0, scale, offsetX, offsetY], ctm),
  ]);
  if (mode === "fit") return Object.freeze({ placement, clip: null });
  const kept = {
    x: (area.x - offsetX) / scale,
    y: (area.y - offsetY) / scale,
    width: area.width / scale,
    height: area.height / scale,
  };
  const full = kept.x <= 1e-9 && kept.y <= 1e-9 && kept.width >= 1 - 1e-9 && kept.height >= 1 - 1e-9;
  return Object.freeze({ placement, clip: full ? null : exactEmbeddedImageTransformClip(kept) });
}

/**
 * The unit-square-to-user-space matrix and unit-space clip that land the
 * visible part of an image on `frame` with the requested turns and flips,
 * keeping whatever orientation the source paint already has: turns act in
 * the image's own space and flips on the page's (screen) axes, then the
 * page-space bounds of the
 * visible region are scaled and moved onto the frame. A new crop is
 * composed onto the existing clip (both in the image's unit square).
 *
 * With `pageTransform` the frame fit, the turns and the flips give way to
 * that one user-space affine, applied after the paint as the reader walked
 * it: the reader's frame composes a move, a resize along the picture's own
 * axes, a turn by any angle and a mirror into it, so a picture keeps its
 * proportions whatever angle it is drawn at (a fit of the bounding box
 * would shear a turned picture). The crop still composes onto the clip, in
 * the image's own square, so the part kept stays where it was on the page.
 */
export function pdfNativeEmbeddedImagePlacement(input: {
  readonly ctm: PdfNativeEmbeddedImageTransformMatrix;
  readonly existingClip: PdfNativeEmbeddedImageTransformClip | null;
  readonly frame: PdfNativeEmbeddedImageTransformFrame;
  readonly quarterTurns: 0 | 1 | 2 | 3;
  readonly flipX: boolean;
  readonly flipY: boolean;
  readonly crop?: PdfNativeEmbeddedImageTransformFrame | null;
  readonly pageTransform?: PdfNativeEmbeddedImageTransformMatrix | null;
  readonly cropMode?: PdfNativeEmbeddedImageCropMode | null;
}): Readonly<{
  placement: PdfNativeEmbeddedImageTransformMatrix;
  clip: PdfNativeEmbeddedImageTransformClip | null;
}> {
  if (input.cropMode) return pdfNativeEmbeddedImageCropModePlacement(input);
  const { frame } = input;
  if (
    ![frame.x, frame.y, frame.width, frame.height].every(Number.isFinite) ||
    !(frame.width > 0) ||
    !(frame.height > 0)
  ) {
    fail("invalid-request", "Embedded image transform frame is invalid.");
  }
  const pageTransform = input.pageTransform
    ? exactEmbeddedImageTransformMatrix([...input.pageTransform])
    : null;
  if (
    pageTransform &&
    (input.quarterTurns !== 0 || input.flipX || input.flipY)
  ) {
    fail(
      "invalid-request",
      "Embedded image page transform cannot be combined with turns or flips.",
    );
  }
  let clip = input.existingClip;
  const crop = input.crop ?? null;
  if (crop) {
    if (
      ![crop.x, crop.y, crop.width, crop.height].every(Number.isFinite) ||
      crop.x < 0 ||
      crop.y < 0 ||
      !(crop.width > 0) ||
      !(crop.height > 0) ||
      crop.x + crop.width > 1 + 1e-6 ||
      crop.y + crop.height > 1 + 1e-6
    ) {
      fail("invalid-request", "Embedded image crop is invalid.");
    }
    const base = clip ?? { x: 0, y: 0, width: 1, height: 1 };
    // `crop` is measured on the visible region with a top-left origin; the
    // unit square has a bottom-left origin.
    const composed = {
      x: base.x + crop.x * base.width,
      y: base.y + (1 - crop.y - crop.height) * base.height,
      width: crop.width * base.width,
      height: crop.height * base.height,
    };
    const full =
      composed.x <= 1e-9 &&
      composed.y <= 1e-9 &&
      composed.width >= 1 - 1e-9 &&
      composed.height >= 1 - 1e-9;
    clip = full ? null : exactEmbeddedImageTransformClip(composed);
  }
  if (pageTransform) {
    const placement = exactEmbeddedImageTransformMatrix([
      ...concatPdfNativeEmbeddedImageTransformMatrices(input.ctm, pageTransform),
    ]);
    // The picture must still cover some of the page's space: a transform
    // that collapses it, or flings it beyond any page, is not a move.
    embeddedImageTransformBounds(placement, clip ?? { x: 0, y: 0, width: 1, height: 1 });
    if (placement.some((value) => Math.abs(value) > 10_000_000)) {
      fail("invalid-request", "Embedded image page transform is out of range.");
    }
    return Object.freeze({ placement, clip });
  }
  let turned = concatPdfNativeEmbeddedImageTransformMatrices(
    UNIT_TURN_MATRICES[input.quarterTurns],
    input.ctm,
  );
  // Flips are the screen's: a horizontal flip mirrors left and right as the
  // picture is shown. Composed in the image's own space they followed its
  // local axes, so on a picture already drawn turned a quarter, "flip
  
  // after the paint's orientation is what is shown; the fit below moves the
  // mirrored bounds back onto the frame.
  if (input.flipX) {
    turned = concatPdfNativeEmbeddedImageTransformMatrices(turned, PAGE_FLIP_X_MATRIX);
  }
  if (input.flipY) {
    turned = concatPdfNativeEmbeddedImageTransformMatrices(turned, PAGE_FLIP_Y_MATRIX);
  }
  const bounds = embeddedImageTransformBounds(
    turned,
    clip ?? { x: 0, y: 0, width: 1, height: 1 },
  );
  const scaleX = frame.width / bounds.width;
  const scaleY = frame.height / bounds.height;
  const fit: PdfNativeEmbeddedImageTransformMatrix = [
    scaleX,
    0,
    0,
    scaleY,
    frame.x - bounds.x * scaleX,
    frame.y - bounds.y * scaleY,
  ];
  return Object.freeze({
    placement: exactEmbeddedImageTransformMatrix([
      ...concatPdfNativeEmbeddedImageTransformMatrices(turned, fit),
    ]),
    clip,
  });
}
