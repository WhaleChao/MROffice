/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Find the images a PDF already contains, and where they sit on the page.
 *
 * Only overlay images added in the editor are selectable today; a figure that
 * shipped inside the PDF cannot be moved, replaced or deleted. Swapping a
 * low-resolution scanned figure for a clean one is the most common image edit
 * there is, and it is currently impossible.
 *
 * The placement is not stored anywhere in the file — it is the transform in
 * force when the image is painted. pdf.js hands back the page's operator list;
 * walking it with a CTM stack recovers each `paintImageXObject`'s rectangle,
 * because that operator paints the UNIT SQUARE and lets the matrix do the
 * placing.
 *
 * Deliberately free of pdf.js: the operator ids come in as a parameter, so the
 * matrix arithmetic — which is the part that goes wrong — is testable without a
 * document.
 */

/** The subset of pdf.js's `OPS` enum this walker needs. */
export interface ImageOperatorIds {
  /** Clip operators; a following `constructPath` supplies the clip bounds. */
  clip?: number;
  eoClip?: number;
  constructPath?: number;
  /** Resource dependency markers pdf.js interleaves; they carry no drawing. */
  dependency?: number;
  /** pdf.js wraps an image paint in `save; setGState(SMask); ...; restore`. */
  setGState?: number;
  save: number;
  restore: number;
  transform: number;
  paintImageXObject: number;
  /** Optional: not every build emits these, and a missing id simply never
   *  matches rather than mis-matching operator 0. pdf.js merges three or more
   *  `q <scale+translate> cm /Same Do Q` in a row into one Repeat operator;
   *  given its id, each of those paints is listed (and counted) on its own,
   *  as the writers count them. */
  paintImageXObjectRepeat?: number;
  paintInlineImageXObject?: number;
  paintImageMaskXObject?: number;
  /** pdf.js brackets a transparency group -- and a soft mask, whose group
   *  options carry `smask` -- with these. What a soft mask draws only says how
   *  much of the page shows; its images are not the page's pictures. */
  beginGroup?: number;
  endGroup?: number;
  



  paintFormXObjectBegin?: number;
  paintFormXObjectEnd?: number;
  /** An annotation's appearance, which pdf.js appends after the page's own
   *  drawing. Its pictures are not page content (no page writer reaches
   *  them) and are measured in the annotation's own space, so given these
   *  ids they are neither listed nor counted. */
  beginAnnotation?: number;
  endAnnotation?: number;
  /** Marked content; the `Props` form with the tag "OC" carries an optional
   *  content group (a layer), see `PlacedImageOptions.isLayerVisible`. */
  beginMarkedContent?: number;
  beginMarkedContentProps?: number;
  endMarkedContent?: number;
}

export interface PlacedImageOptions {
  




  isLayerVisible?: (properties: unknown) => boolean;
}

/** A pdf.js operator list, as `getOperatorList()` returns it. */
export interface OperatorListLike {
  fnArray: ArrayLike<number>;
  argsArray: ArrayLike<unknown>;
}

export interface PlacedImage {
  /** The XObject's name in the page's resources, or "" for an inline image. */
  name: string;
  /** Page-relative fractions (0..1), top-left origin — the editor's own model. */
  x: number;
  y: number;
  width: number;
  height: number;
  





  sourceOpacity?: number;
  /** Position in the content stream; a later paint sits on top of an earlier. */
  order: number;
  /** The paint's index in the operator list, the order it shares with every
   *  other mark on the page (a vector's `operatorIndex` is the same count).
   *  Absent on a record that was not read from an operator list. */
  operatorIndex?: number;
  /**
   * True when the paint matrix rotates or skews the image. The rectangle is
   * then the bounding box rather than the image's own frame, so a replacement
   * fitted to it would not land the same way — the UI can say so instead of
   * quietly producing a crooked swap.
   */
  rotated: boolean;
  /** True for a stencil mask (`paintImageMaskXObject`), which carries no
   *  colour of its own and is usually a glyph-like decoration. */
  mask: boolean;
  /**
   * The CTM in force at the paint, in user space (`[a b c d e f]`): the unit
   * square maps through it onto the page. A native transform composes its
   * own matrix onto this instead of re-encoding pixels.
   */
  matrix: readonly [number, number, number, number, number, number];
  /**
   * The whole paint frame (the unit square through `matrix`) when a
   * rectangular clip in force at the paint hides part of it; then `x`, `y`,
   * `width` and `height` describe only the visible part. Absent when nothing
   * is clipped away.
   */
  frame?: { x: number; y: number; width: number; height: number };
  /**
   * The clips in force at the paint that are not the picture's own, as one
   * page-fraction box (their rectangles' intersection): a printed page's
   * printable area, a column. Absent when none cuts into the page. A move
   * past it takes the picture out from under it (the native reorder's
   * leaveClips), or the picture is drawn cut at its edge.
   */
  enclosingClip?: { x: number; y: number; width: number; height: number };
}

type Matrix = [number, number, number, number, number, number];

const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

/** `a` applied after `b` — the order pdf.js's `transform` op composes in. */
export function multiplyMatrix(a: Matrix, b: Matrix): Matrix {
  return [
    a[0] * b[0] + a[2] * b[1],
    a[1] * b[0] + a[3] * b[1],
    a[0] * b[2] + a[2] * b[3],
    a[1] * b[2] + a[3] * b[3],
    a[0] * b[4] + a[2] * b[5] + a[4],
    a[1] * b[4] + a[3] * b[5] + a[5],
  ];
}

function applyMatrix(m: Matrix, x: number, y: number): [number, number] {
  return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
}

function isMatrix(value: unknown): value is Matrix {
  return (
    Array.isArray(value) &&
    value.length >= 6 &&
    value.slice(0, 6).every((n) => typeof n === "number" && Number.isFinite(n))
  );
}

/** A matrix operand as an array, also from a typed array: pdf.js hands a
 *  form's /Matrix over as a Float32Array, which `Array.isArray` rejects, so
 *  a picture inside a scaled or moved form was measured as if the form sat
 *  untransformed. */
function matrixOperand(value: unknown): Matrix | null {
  const values =
    value && ArrayBuffer.isView(value) && !(value instanceof DataView)
      ? Array.from(value as unknown as ArrayLike<number>)
      : value;
  return isMatrix(values) ? (values.slice(0, 6) as Matrix) : null;
}














export function placedImagesFromOperatorList(
  operatorList: OperatorListLike,
  ops: ImageOperatorIds,
  page: PageBox,
  options: PlacedImageOptions = {},
): PlacedImage[] {
  const { fnArray, argsArray } = operatorList;
  const count = Math.min(fnArray.length, argsArray.length);
  if (count === 0 || !(page.width > 0) || !(page.height > 0)) return [];

  let ctm: Matrix = IDENTITY;
  const stack: Matrix[] = [];
  // The non-stroking alpha is graphics state too: `q` keeps it, `Q` brings
  // it back, and a `gs` naming an ExtGState with `ca` sets it.
  let alpha: number | null = null;
  const alphaStack: Array<number | null> = [];
  const out: PlacedImage[] = [];
  // An image-local clip: a rectangular clip set in the paint's own graphics
  // level as `q <cm> <rect clip> /Name Do Q`, the shape the native transform
  // writer produces. Clips from outer levels (a page box, a form frame) are
  // not the image's own and are ignored, so the frame the editor shows is
  // exactly the region the writer can re-map. pdf.js emits `clip` before the
  // path that defines it, so a clip is pending until its `constructPath`.
  type Box = [number, number, number, number];
  type Level = { ops: number[]; clip: Box | null; pendingClip: boolean };
  let level: Level = { ops: [], clip: null, pendingClip: false };
  const levels: Level[] = [];
  const isOwnClipLevel = (current: Level) => {
    // The writer's wrapper names its alpha before the `cm` when the paint
    // has one (`q /GS gs <cm> <rect> re W n /X Do Q`). A `gs` moves neither
    // the frame nor the clip; missing the clip behind it reported a cropped,
    // translucent picture as its whole frame, and the next crop, measured
    
    const seen = current.ops.filter(
      (op) => ops.setGState === undefined || op !== ops.setGState,
    );
    return (
      current.clip !== null &&
      seen.length === 3 &&
      seen[0] === ops.transform &&
      (seen[1] === ops.clip || seen[1] === ops.eoClip) &&
      seen[2] === ops.constructPath
    );
  };
  // pdf.js nests the paint itself in a synthetic `save; setGState; paint;
  // restore` level; look through such levels to the one holding the clip.
  const ownClipFor = (current: Level): Box | null => {
    let candidate = current;
    let depth = levels.length - 1;
    while (
      candidate.clip === null &&
      candidate.ops.every((op) => ops.setGState !== undefined && op === ops.setGState) &&
      depth >= 0
    ) {
      candidate = levels[depth];
      depth -= 1;
    }
    return isOwnClipLevel(candidate) ? candidate.clip : null;
  };

  // Open groups, true for a soft mask's; images inside any soft mask are the
  // mask's, which the native planner never lists as pictures either.
  const groups: boolean[] = [];
  let maskDepth = 0;
  let annotationDepth = 0;
  // Open marked-content sections, true for one whose layer is switched off.
  const markedContent: boolean[] = [];
  let hiddenDepth = 0;
  let paintOrder = 0;
  const pushLevel = () => {
    stack.push(ctm);
    alphaStack.push(alpha);
    levels.push(level);
    level = { ops: [], clip: null, pendingClip: false };
  };
  const popLevel = () => {
    // An unbalanced restore is a malformed stream, not a reason to give up on
    // the images already found.
    ctm = stack.pop() ?? IDENTITY;
    alpha = alphaStack.length > 0 ? (alphaStack.pop() ?? null) : null;
    level = levels.pop() ?? { ops: [], clip: null, pendingClip: false };
  };
  for (let i = 0; i < count; i += 1) {
    const fn = fnArray[i];
    if (
      (ops.beginMarkedContent !== undefined && fn === ops.beginMarkedContent) ||
      (ops.beginMarkedContentProps !== undefined && fn === ops.beginMarkedContentProps)
    ) {
      const args = argsArray[i] as unknown[] | undefined;
      const hidden =
        fn === ops.beginMarkedContentProps &&
        args?.[0] === "OC" &&
        options.isLayerVisible !== undefined &&
        !options.isLayerVisible(args[1]);
      markedContent.push(hidden);
      if (hidden) hiddenDepth += 1;
      continue;
    }
    if (ops.endMarkedContent !== undefined && fn === ops.endMarkedContent) {
      if (markedContent.pop()) hiddenDepth -= 1;
      continue;
    }
    if (ops.paintFormXObjectBegin !== undefined && fn === ops.paintFormXObjectBegin) {
      pushLevel();
      const matrix = matrixOperand((argsArray[i] as unknown[] | undefined)?.[0]);
      if (matrix) ctm = multiplyMatrix(ctm, matrix);
      continue;
    }
    if (ops.beginAnnotation !== undefined && fn === ops.beginAnnotation) {
      pushLevel();
      annotationDepth += 1;
      continue;
    }
    if (ops.endAnnotation !== undefined && fn === ops.endAnnotation) {
      popLevel();
      annotationDepth = Math.max(0, annotationDepth - 1);
      continue;
    }
    if (ops.paintFormXObjectEnd !== undefined && fn === ops.paintFormXObjectEnd) {
      popLevel();
      continue;
    }
    if (ops.beginGroup !== undefined && fn === ops.beginGroup) {
      const options = (argsArray[i] as unknown[] | undefined)?.[0] as { smask?: unknown } | undefined;
      const mask = Boolean(options?.smask);
      groups.push(mask);
      if (mask) maskDepth += 1;
      continue;
    }
    if (ops.endGroup !== undefined && fn === ops.endGroup) {
      if (groups.pop()) maskDepth -= 1;
      continue;
    }
    if (fn === ops.save) {
      pushLevel();
      continue;
    }
    if (fn === ops.restore) {
      popLevel();
      continue;
    }
    if (ops.setGState !== undefined && fn === ops.setGState) {
      // pdf.js hands the ExtGState as `[[key, value], ...]` pairs.
      const entries = (argsArray[i] as unknown[] | undefined)?.[0];
      if (Array.isArray(entries)) {
        for (const entry of entries) {
          if (
            Array.isArray(entry) &&
            entry[0] === "ca" &&
            typeof entry[1] === "number" &&
            Number.isFinite(entry[1])
          ) {
            alpha = Math.min(1, Math.max(0, entry[1]));
          }
        }
      }
      level.ops.push(fn);
      continue;
    }
    if (
      (ops.clip !== undefined && fn === ops.clip) ||
      (ops.eoClip !== undefined && fn === ops.eoClip)
    ) {
      level.pendingClip = true;
      level.ops.push(fn);
      continue;
    }
    if (ops.constructPath !== undefined && fn === ops.constructPath) {
      level.ops.push(fn);
      if (level.pendingClip) {
        level.pendingClip = false;
        const args = argsArray[i];
        const minMax = Array.isArray(args) ? args[2] : undefined;
        const local =
          Array.isArray(minMax) && minMax.length >= 4 && minMax.every((v) => Number.isFinite(v))
            ? (minMax as number[])
            : // Any typed array: one from another realm (pdf.js's worker)
              // fails an `instanceof` against this realm's constructors.
              ArrayBuffer.isView(minMax) && !(minMax instanceof DataView)
              ? Array.from(minMax as unknown as ArrayLike<number>)
              : null;
        if (local && local.length >= 4) {
          const corners = [
            applyMatrix(ctm, local[0], local[1]),
            applyMatrix(ctm, local[2], local[1]),
            applyMatrix(ctm, local[0], local[3]),
            applyMatrix(ctm, local[2], local[3]),
          ];
          level.clip = [
            Math.min(...corners.map(([x]) => x)),
            Math.min(...corners.map(([, y]) => y)),
            Math.max(...corners.map(([x]) => x)),
            Math.max(...corners.map(([, y]) => y)),
          ];
        } else {
          level.clip = null;
        }
      }
      continue;
    }
    if (fn === ops.transform) {
      const args = argsArray[i];
      const matrix = Array.isArray(args) && args.length === 1 ? args[0] : args;
      if (isMatrix(matrix)) {
        ctm = multiplyMatrix(ctm, matrix as Matrix);
      } else if (Array.isArray(args) && args.length >= 6 && isMatrix(args)) {
        ctm = multiplyMatrix(ctm, args as Matrix);
      }
      level.ops.push(fn);
      continue;
    }

    const repeat =
      ops.paintImageXObjectRepeat !== undefined && fn === ops.paintImageXObjectRepeat;
    const isPaint =
      fn === ops.paintImageXObject ||
      repeat ||
      (ops.paintInlineImageXObject !== undefined &&
        fn === ops.paintInlineImageXObject) ||
      (ops.paintImageMaskXObject !== undefined &&
        fn === ops.paintImageMaskXObject);
    if (!isPaint) {
      if (ops.dependency === undefined || fn !== ops.dependency) level.ops.push(fn);
      continue;
    }
    if (maskDepth > 0 || annotationDepth > 0) {
      level.ops.push(fn);
      continue;
    }
    const args = argsArray[i];
    const first = Array.isArray(args) ? args[0] : undefined;
    if (repeat) {
      // `[objId, scaleX, scaleY, positions]`: each merged paint was
      // `q <scaleX 0 0 scaleY x y> cm /Same Do Q`, in its own saved state.
      const scaleX = Array.isArray(args) ? args[1] : undefined;
      const scaleY = Array.isArray(args) ? args[2] : undefined;
      const positions = Array.isArray(args) ? args[3] : undefined;
      const xy =
        positions && (Array.isArray(positions) || ArrayBuffer.isView(positions))
          ? Array.from(positions as ArrayLike<number>)
          : [];
      level.ops.push(fn);
      if (typeof scaleX !== "number" || typeof scaleY !== "number") continue;
      for (let index = 0; index + 1 < xy.length; index += 2) {
        const at = multiplyMatrix(ctm, [scaleX, 0, 0, scaleY, xy[index], xy[index + 1]]);
        const rect = rectFromMatrix(at, page);
        if (!rect) continue;
        const order = paintOrder;
        paintOrder += 1;
        if (hiddenDepth > 0) continue;
        out.push({
          name: typeof first === "string" ? first : "",
          ...rect,
          order,
          operatorIndex: i,
          rotated: false,
          mask: false,
          matrix: [at[0], at[1], at[2], at[3], at[4], at[5]],
          ...(alpha !== null ? { sourceOpacity: alpha } : {}),
        });
      }
      continue;
    }

    // An inline image (BI ... EI) is no `Do`: the native planner, which
    // numbers the page's image invocations, never counts it, so it takes no
    // number here either. It did, and every picture painted after one was
    // offered under the number of another: a flip of the top-right picture of
    // pdf.js's images_1bit_grayscale.pdf was planned for the picture before it
    
    if (ops.paintInlineImageXObject !== undefined && fn === ops.paintInlineImageXObject) {
      level.ops.push(fn);
      continue;
    }
    const rect = rectFromMatrix(ctm, page);
    if (!rect) continue;
    // `order` is the picture's key toward the native planner, which counts
    // every painted image: one in a layer that is switched off keeps its
    // number and is simply not offered.
    const order = paintOrder;
    paintOrder += 1;
    if (hiddenDepth > 0) {
      level.ops.push(fn);
      continue;
    }
    // The image's own clip hides part of the frame; the editor works with what
    // is visible and keeps the whole frame for transforms.
    let visible = rect;
    let frame: PlacedImage["frame"];
    const clipBox = ownClipFor(level);
    // Every other clip in force, as one box.
    let enclosing: Box | null = null;
    for (const held of [...levels, level]) {
      if (!held.clip || held.clip === clipBox) continue;
      enclosing = enclosing
        ? [Math.max(enclosing[0], held.clip[0]), Math.max(enclosing[1], held.clip[1]),
          Math.min(enclosing[2], held.clip[2]), Math.min(enclosing[3], held.clip[3])]
        : held.clip;
    }
    const enclosingClip = enclosing && enclosing[2] > enclosing[0] && enclosing[3] > enclosing[1]
      ? fractionBox(page, [[enclosing[0], enclosing[1]], [enclosing[2], enclosing[3]]])
      : null;
    const cutsPage = enclosingClip !== null && (enclosingClip.x > 1e-4 || enclosingClip.y > 1e-4 ||
      enclosingClip.x + enclosingClip.width < 1 - 1e-4 || enclosingClip.y + enclosingClip.height < 1 - 1e-4);
    level.ops.push(fn);
    if (clipBox) {
      const clip = fractionBox(page, [
        [clipBox[0], clipBox[1]],
        [clipBox[2], clipBox[3]],
      ]);
      const left = Math.max(rect.x, clip.x);
      const right = Math.min(rect.x + rect.width, clip.x + clip.width);
      const top = Math.max(rect.y, clip.y);
      const bottom = Math.min(rect.y + rect.height, clip.y + clip.height);
      if (
        right - left > 1e-6 &&
        bottom - top > 1e-6 &&
        (right - left < rect.width - 1e-6 || bottom - top < rect.height - 1e-6)
      ) {
        visible = { x: left, y: top, width: right - left, height: bottom - top };
        frame = { ...rect };
      }
    }
    out.push({
      name: typeof first === "string" ? first : "",
      ...visible,
      ...(frame ? { frame } : {}),
      order,
      operatorIndex: i,
      // b and c carry rotation/skew; a pure scale-and-translate leaves them 0.
      rotated: Math.abs(ctm[1]) > 1e-6 || Math.abs(ctm[2]) > 1e-6,
      mask:
        ops.paintImageMaskXObject !== undefined &&
        fn === ops.paintImageMaskXObject,
      matrix: [ctm[0], ctm[1], ctm[2], ctm[3], ctm[4], ctm[5]],
      ...(alpha !== null ? { sourceOpacity: alpha } : {}),
      ...(cutsPage ? { enclosingClip: enclosingClip! } : {}),
    });
  }
  return out;
}

/**
 * The page-fraction rectangle a paint matrix covers.
 *
 * `paintImageXObject` paints the unit square, so the placement is the matrix
 * applied to (0,0)-(1,1). Its four corners are mapped and bounded, which stays
 * correct for a rotated or flipped matrix; PDF user space has a bottom-left
 * origin, so the top edge is the page height minus the highest corner.
 */
function rectFromMatrix(
  ctm: Matrix,
  page: PageBox,
): { x: number; y: number; width: number; height: number } | null {
  const corners = [
    applyMatrix(ctm, 0, 0),
    applyMatrix(ctm, 1, 0),
    applyMatrix(ctm, 0, 1),
    applyMatrix(ctm, 1, 1),
  ];
  const rect = fractionBox(page, corners);
  if (!(rect.width > 0) || !(rect.height > 0)) return null;
  if (![rect.x, rect.y, rect.width, rect.height].every(Number.isFinite)) return null;
  return rect;
}

/** The page box the walker measures against, see `placedImagesFromOperatorList`. */
export interface PageBox {
  width: number;
  height: number;
  /** pdf.js `viewport.transform` at rotation 0: user space to viewport px. */
  transform?: ArrayLike<number>;
}

/** User-space points as a top-left-origin page-fraction bounding box. */
function fractionBox(
  page: PageBox,
  points: ReadonlyArray<readonly [number, number]>,
): { x: number; y: number; width: number; height: number } {
  const transform =
    page.transform && page.transform.length >= 6
      ? (Array.from(page.transform).slice(0, 6) as Matrix)
      : null;
  const mapped = points.map(([x, y]) =>
    transform && transform.every(Number.isFinite)
      ? (() => {
          const [vx, vy] = applyMatrix(transform, x, y);
          return [vx / page.width, vy / page.height] as const;
        })()
      : ([x / page.width, (page.height - y) / page.height] as const),
  );
  const xs = mapped.map(([x]) => x);
  const ys = mapped.map(([, y]) => y);
  const left = Math.min(...xs);
  const top = Math.min(...ys);
  return {
    x: left,
    y: top,
    width: Math.max(...xs) - left,
    height: Math.max(...ys) - top,
  };
}

/**
 * Drop images that are too small to be a figure — a 1x1 spacer, a hairline
 * rule, a stencil glyph. Selecting those makes the picker useless on a page
 * that uses images for typography.
 */
export function significantImages(
  images: ReadonlyArray<PlacedImage>,
  minFraction = 0.01,
): PlacedImage[] {
  return images.filter(
    (image) =>
      !image.mask &&
      image.width >= minFraction &&
      image.height >= minFraction &&
      image.x + image.width > 0 &&
      image.y + image.height > 0 &&
      image.x < 1 &&
      image.y < 1,
  );
}
