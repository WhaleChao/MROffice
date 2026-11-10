/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Office's "Save as Picture" on a shape (Word, Excel and PowerPoint put it on
 * the shape's context menu): the shape as the host draws it, as a PNG, or as
 * an SVG where the realm cannot rasterize (no canvas, a tainted picture fill).
 *
 * ONE rule for the three editors: the shape's own DOM root (the deck's block
 * container, the worksheet's object element, Docs' shape figure) is cloned
 * with its computed styles inlined and wrapped in an SVG `foreignObject`, so
 * the picture is exactly what the page shows -- preset geometry, fill,
 * outline, WordArt paint, rotation and all -- without a second renderer.
 */
import { downloadBlobFile, isBrowserDownloadAvailable } from "./download-file";
import {
  captureElementPicture,
  cloneWithComputedStyle,
  elementPictureSvg,
  rasterElementPicture,
} from "./export/element-picture";

export interface OfficeShapePicture {
  readonly svg: string;
  readonly width: number;
  readonly height: number;
}

/** Selection chrome and editing controls are the host's, not the shape's. */
const HOST_CONTROLS = "input, textarea, button";









interface PictureBox {
  width: number;
  height: number;
  /** Where the shape's own box starts inside the picture. */
  left: number;
  top: number;
  /** The shape's own layout size. */
  shapeWidth: number;
  shapeHeight: number;
}
function pictureSize(
  root: Element,
  size?: { width: number; height: number },
): PictureBox | null {
  const rect = root.getBoundingClientRect();
  const layoutWidth =
    size?.width ?? (root instanceof HTMLElement && root.offsetWidth > 0 ? root.offsetWidth : rect.width);
  const layoutHeight =
    size?.height ?? (root instanceof HTMLElement && root.offsetHeight > 0 ? root.offsetHeight : rect.height);
  if (!(layoutWidth > 0) || !(layoutHeight > 0)) return null;
  // Client pixels per layout pixel (the canvas zoom), from the box itself.
  const scale = rect.width > 0 ? layoutWidth / rect.width : 1;
  let minX = 0;
  let minY = 0;
  let maxX = layoutWidth;
  let maxY = layoutHeight;
  if (!size && rect.width > 0) {
    for (const node of root.querySelectorAll("*")) {
      if (node.matches(HOST_CONTROLS)) continue;
      const box = node.getBoundingClientRect();
      if (!(box.width > 0) || !(box.height > 0)) continue;
      minX = Math.min(minX, (box.left - rect.left) * scale);
      minY = Math.min(minY, (box.top - rect.top) * scale);
      maxX = Math.max(maxX, (box.right - rect.left) * scale);
      maxY = Math.max(maxY, (box.bottom - rect.top) * scale);
    }
  }
  const width = Math.ceil(maxX - minX);
  const height = Math.ceil(maxY - minY);
  if (!(width > 0) || !(height > 0) || width > 8192 || height > 8192) return null;
  return {
    width,
    height,
    left: Math.round(-minX),
    top: Math.round(-minY),
    shapeWidth: layoutWidth,
    shapeHeight: layoutHeight,
  };
}

/** The root sits wherever its host placed it; the picture is its own page,
 * the shape at its offset inside the box it draws over. */
function placeAtOrigin(clone: HTMLElement, box: PictureBox): void {
  clone.style.position = "relative";
  clone.style.left = `${box.left}px`;
  clone.style.top = `${box.top}px`;
  clone.style.margin = "0px";
  clone.style.width = `${box.shapeWidth}px`;
  clone.style.height = `${box.shapeHeight}px`;
}

/**
 * The shape as an SVG document. `size` overrides the measured box (a realm
 * without layout, such as a test, measures zero); null when there is no box.
 */
export function officeShapePictureSvg(
  root: Element,
  size?: { width: number; height: number },
): OfficeShapePicture | null {
  const box = pictureSize(root, size);
  if (!box) return null;
  const { width, height } = box;
  const { clone, pseudoRules } = cloneWithComputedStyle(root, HOST_CONTROLS);
  if (clone instanceof HTMLElement) placeAtOrigin(clone, box);
  const markup = new XMLSerializer().serializeToString(clone);
  const svg = elementPictureSvg({ markup, css: pseudoRules.join("\n"), width, height });
  return { svg, width, height };
}

/**
 * The shape as a PNG, with its pictures and faces carried in: a picture fill
 * or a served face would otherwise come out blank or in a fallback face.
 */
async function rasterizeShape(
  root: Element,
  size?: { width: number; height: number },
): Promise<Blob | null> {
  const box = pictureSize(root, size);
  if (!box) return null;
  const picture = await captureElementPicture(root, {
    width: box.width,
    height: box.height,
    omit: HOST_CONTROLS,
    edit: (clone) => placeAtOrigin(clone, box),
  });
  return rasterElementPicture(picture, { scale: 2 });
}

/**
 * The shape the host draws at `root` as a PNG, the same rendering Save as
 * Picture downloads, or null where the realm cannot rasterize it. The sheet
 * hands it to the system clipboard when an object is copied, so another
 * application pastes the object as it looks.
 */
export async function rasterizeOfficeShapeAsPng(
  root: Element,
  size?: { width: number; height: number },
): Promise<Blob | null> {
  return rasterizeShape(root, size).catch(() => null);
}

/**
 * Save the shape the host draws at `root` as `<baseName>.png`, falling back
 * to `<baseName>.svg` where the realm cannot rasterize. The result says which
 * (or that this realm cannot download at all).
 */
export async function saveOfficeShapeAsPicture(
  root: Element,
  baseName: string,
  size?: { width: number; height: number },
): Promise<"png" | "svg" | "unavailable"> {
  const picture = officeShapePictureSvg(root, size);
  if (!picture || !isBrowserDownloadAvailable()) return "unavailable";
  const name = baseName.replace(/[\\/:*?"<>|]+/g, "-").trim() || "shape";
  const png = await rasterizeShape(root, size).catch(() => null);
  if (png && downloadBlobFile(`${name}.png`, png)) return "png";
  const svg = new Blob([picture.svg], { type: "image/svg+xml;charset=utf-8" });
  return downloadBlobFile(`${name}.svg`, svg) ? "svg" : "unavailable";
}
