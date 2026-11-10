/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativePlanners } from "./remote/pdf-native-planners";
import {
  renderPdfPageLike,
  solveAboveLayer,
  type PdfPageRenderMatch,
  type PdfVectorDragAbove,
} from "./native-vector-drag-backdrop";

/**
 * Whether a solved "painted above" layer can be believed: the page-wide fill
 * that stood in for the picture really reached the page outside it. A paint
 * under a clip (a frame clip, a form's box) fills only that clip, and the
 * rest of the page then reads as covered by something opaque -- a ghost
 * dragged there would vanish under it. Text and drawings over a page leave
 * far most of it uncovered.
 */
export function aboveLayerReachesPage(
  above: PdfVectorDragAbove,
  bounds: { x: number; y: number; width: number; height: number },
): boolean {
  const left = Math.floor(bounds.x * above.width);
  const top = Math.floor(bounds.y * above.height);
  const right = Math.ceil((bounds.x + bounds.width) * above.width);
  const bottom = Math.ceil((bounds.y + bounds.height) * above.height);
  const stride = Math.max(1, Math.floor((above.width * above.height) / 200_000));
  let outside = 0;
  let covered = 0;
  for (let p = 0; p < above.alpha.length; p += stride) {
    const x = p % above.width;
    const y = (p - x) / above.width;
    if (x >= left && x < right && y >= top && y < bottom) continue;
    outside += 1;
    if (above.alpha[p] >= 230) covered += 1;
  }
  return outside === 0 || covered / outside <= 0.5;
}

/**
 * Whether an image XObject paints every pixel of its frame: no soft mask,
 * no colour-key or stencil mask, no alpha inside its own JPEG 2000 data.
 * Such a picture is its whole rectangle, including the parts that happen to
 * match the page under them; one with a mask is only where it shows.
 *
 * An image is a stream, so its dictionary is never inside an object stream:
 * it is read from the file as written, the newest definition winning.
 * Anything that cannot be read says "not known to be opaque".
 */
export function imageXObjectIsOpaque(
  bytes: Uint8Array,
  objectNumber: number,
  generation: number,
): boolean {
  const header = new TextEncoder().encode(`${objectNumber} ${generation} obj`);
  let at = -1;
  for (let i = bytes.length - header.length; i >= 0; i -= 1) {
    let match = true;
    for (let k = 0; k < header.length; k += 1) {
      if (bytes[i + k] !== header[k]) {
        match = false;
        break;
      }
    }
    // "12 0 obj" must not be the tail of "112 0 obj".
    if (match && (i === 0 || bytes[i - 1] < 0x30 || bytes[i - 1] > 0x39)) {
      at = i;
      break;
    }
  }
  if (at < 0) return false;
  const window = bytes.subarray(at, Math.min(bytes.length, at + 4096));
  let dictionary = "";
  for (let i = 0; i < window.length; i += 1) dictionary += String.fromCharCode(window[i]);
  const end = dictionary.indexOf("stream");
  if (end < 0) return false;
  dictionary = dictionary.slice(0, end);
  if (!/\/Subtype\s*\/Image\b/.test(dictionary)) return false;
  return !/\/SMask\b|\/Mask\b|\/ImageMask\s+true|\/SMaskInData\s+[12]\b/.test(dictionary);
}

/**
 * The page `pageNumber` of `sourceBytes` without the image paint at
 * `imageOrder`, rendered at the match's size, and whether that image paints
 * its whole frame; throws when the removal cannot be planned (the caller
 * keeps its plain ghost then).
 */
export async function renderEmbeddedImageAbsentPage(
  input: {
    /** The planners of the page the bytes are shown on. */
    readonly planners: PdfNativePlanners;
    readonly sourceBytes: Uint8Array;
    readonly pageNumber: number;
    readonly imageOrder: number;
    /** The picture as displayed (page fractions): with it, what the page
     * paints over the picture is solved too (`above`). */
    readonly bounds?: { x: number; y: number; width: number; height: number };
  } & PdfPageRenderMatch,
): Promise<{
  canvas: HTMLCanvasElement;
  opaque: boolean;
  


  above: PdfVectorDragAbove | null;
} | null> {
  const { planners } = input;
  const plan = await planners.planMraiNativeEmbeddedImageRemoval({
    sourceBytes: input.sourceBytes,
    selection: { pageNumber: input.pageNumber, imageOrder: input.imageOrder },
  });
  const opaque = imageXObjectIsOpaque(
    input.sourceBytes,
    plan.evidence.image.objectNumber,
    plan.evidence.image.generation,
  );
  
  // by a page-wide black and white fill in the same graphics state.
  const fill = (gray: 0 | 1) =>
    planners.prepareMraiNativeContentStreamTransaction(input.sourceBytes, {
      ...plan.request,
      replacement: new TextEncoder().encode(
        `q ${gray} g -100000 -100000 200000 200000 re f Q`,
      ),
    })
      .then((prepared) => prepared.candidateBytes)
      .catch(() => null);
  const [prepared, blackBytes, whiteBytes] = await Promise.all([
    planners.prepareMraiNativeContentStreamTransaction(input.sourceBytes, plan.request),
    input.bounds ? fill(0) : Promise.resolve(null),
    input.bounds ? fill(1) : Promise.resolve(null),
  ]);
  const candidate = prepared.candidateBytes;
  try {
    const [canvas, black, white] = await Promise.all([
      renderPdfPageLike(candidate, input.pageNumber, input),
      blackBytes
        ? renderPdfPageLike(blackBytes, input.pageNumber, input).catch(() => null)
        : Promise.resolve(null),
      whiteBytes
        ? renderPdfPageLike(whiteBytes, input.pageNumber, input).catch(() => null)
        : Promise.resolve(null),
    ]);
    if (!canvas) return null;
    let above: PdfVectorDragAbove | null = null;
    if (
      input.bounds &&
      black &&
      white &&
      black.width === canvas.width &&
      black.height === canvas.height &&
      white.width === canvas.width &&
      white.height === canvas.height
    ) {
      try {
        const read = (target: HTMLCanvasElement) =>
          target
            .getContext("2d", { willReadFrequently: true })
            ?.getImageData(0, 0, target.width, target.height).data ?? null;
        const absentData = read(canvas);
        const blackData = read(black);
        const whiteData = read(white);
        const solved =
          absentData && blackData && whiteData
            ? solveAboveLayer(absentData, blackData, whiteData, canvas.width, canvas.height)
            : null;
        above = solved && aboveLayerReachesPage(solved, input.bounds) ? solved : null;
      } catch {
        above = null;
      }
    }
    return { canvas, opaque, above };
  } finally {
    candidate.fill(0);
    blackBytes?.fill(0);
    whiteBytes?.fill(0);
  }
}
