/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A page's new frame held for Save (AutoSave off): a crop or a page size the
 * window holds is shown before the saved file carries it. The page keeps the
 * coordinate frame of the file shown -- every layer on it (comments, shapes,
 * ink, text boxes, fields) is placed in fractions of that frame -- and the
 * held page is a window onto it: what lies outside the held page is clipped,
 * the page takes the held page's room in the reader, and paper fills what the
 * held page has beyond the page shown (an enlarged page, a crop let out).
 *
 * The geometry follows the server's writers exactly (native_edit_primitives
 * crop_page and resize_page): a crop is a top-left fraction of the page as
 * rendered; a page size keeps the lower-left corner in user space, and
 * fit-center scales the content by min(W/w, H/h) and centres it.
 */

/** What the window holds for a page, in the terms the page JSON keeps. */
export type PdfHeldPageGeometry =
  | Readonly<{
      kind: "crop";
      /** The held keep-rect and the one the shown file carries, both in
       * percent of the uncropped page as rendered (top-left origin). */
      held: Readonly<{ x: number; y: number; width: number; height: number }>;
      saved: Readonly<{ x: number; y: number; width: number; height: number }>;
    }>
  | Readonly<{
      kind: "page-size";
      /** The page's new size in points, and how its content is placed. */
      width: number;
      height: number;
      placement: "unchanged" | "fit-center";
    }>;

export interface PdfHeldPageGeometries {
  readonly key: string;
  readonly pages: Readonly<Record<number, PdfHeldPageGeometry>>;
  /** Save wrote them: still shown on the document on screen until the
   * saved one replaces it (no flash back to the old frame), then dropped. */
  readonly retired?: boolean;
}

/** The held page in fractions of the page shown (its rendered frame): it may
 * reach past it on any side. `scale` shrinks or grows the page shown inside
 * the held page (fit-center), about its top-left corner. */
export interface PdfHeldPageFrame {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly scale: number;
}

const finite = (value: number) => Number.isFinite(value);

/** A held crop as a frame of the page shown, whose own crop is `saved`. */
export function heldCropFrame(geometry: Extract<PdfHeldPageGeometry, { kind: "crop" }>): PdfHeldPageFrame | null {
  const { held, saved } = geometry;
  if (![held.x, held.y, held.width, held.height, saved.x, saved.y, saved.width, saved.height].every(finite) ||
    held.width <= 0 || held.height <= 0 || saved.width <= 0 || saved.height <= 0) return null;
  const frame = {
    x: (held.x - saved.x) / saved.width,
    y: (held.y - saved.y) / saved.height,
    width: held.width / saved.width,
    height: held.height / saved.height,
    scale: 1,
  };
  return isIdentityFrame(frame) ? null : Object.freeze(frame);
}

/**
 * A held page size as a frame of the page shown. `toRendered` maps a user-space
 * point to the page shown as rendered at scale 1 (its viewport), `view` is
 * the page's box in user space and `rendered` its rendered size.
 */
export function heldPageSizeFrame(input: Readonly<{
  geometry: Extract<PdfHeldPageGeometry, { kind: "page-size" }>;
  view: readonly [number, number, number, number];
  rendered: Readonly<{ width: number; height: number }>;
  toRendered: (x: number, y: number) => readonly [number, number];
}>): PdfHeldPageFrame | null {
  const { geometry, view, rendered } = input;
  const [x0, y0, x1, y1] = view;
  const sourceWidth = x1 - x0, sourceHeight = y1 - y0;
  if (![geometry.width, geometry.height, sourceWidth, sourceHeight, rendered.width, rendered.height].every(finite) ||
    geometry.width <= 0 || geometry.height <= 0 || sourceWidth <= 0 || sourceHeight <= 0 ||
    rendered.width <= 0 || rendered.height <= 0) return null;
  // The new page in user space: the lower-left corner stays (resize_page).
  const corners = [
    input.toRendered(x0, y0),
    input.toRendered(x0 + geometry.width, y0),
    input.toRendered(x0, y0 + geometry.height),
    input.toRendered(x0 + geometry.width, y0 + geometry.height),
  ];
  const xs = corners.map(([x]) => x), ys = corners.map(([, y]) => y);
  const left = Math.min(...xs), top = Math.min(...ys);
  const width = Math.max(...xs) - left, height = Math.max(...ys) - top;
  if (!(width > 0) || !(height > 0)) return null;
  if (geometry.placement === "fit-center") {
    // The content scaled into the new page and centred in it: the page shown,
    // scaled, sits in the middle of the held page.
    const scale = Math.min(geometry.width / sourceWidth, geometry.height / sourceHeight);
    if (!(scale > 0) || !finite(scale)) return null;
    const frame = {
      x: -(width - scale * rendered.width) / (2 * scale * rendered.width),
      y: -(height - scale * rendered.height) / (2 * scale * rendered.height),
      width: width / (scale * rendered.width),
      height: height / (scale * rendered.height),
      scale,
    };
    return isIdentityFrame(frame) ? null : Object.freeze(frame);
  }
  const frame = {
    x: left / rendered.width,
    y: top / rendered.height,
    width: width / rendered.width,
    height: height / rendered.height,
    scale: 1,
  };
  return isIdentityFrame(frame) ? null : Object.freeze(frame);
}

function isIdentityFrame(frame: PdfHeldPageFrame): boolean {
  const near = (a: number, b: number) => Math.abs(a - b) < 1e-6;
  return near(frame.x, 0) && near(frame.y, 0) && near(frame.width, 1) && near(frame.height, 1) && near(frame.scale, 1);
}

/** How the page shown (of `pageWidth` x `pageHeight` px) takes the frame. */
export interface PdfHeldPageFrameLayout {
  /** The held page in the page's own px, before `scale`. */
  readonly region: Readonly<{ left: number; top: number; width: number; height: number }>;
  /** Whether the held page reaches past the page shown (paper beyond it). */
  readonly extends: boolean;
  readonly style: Readonly<{
    clipPath: string;
    marginLeft: number;
    marginRight: number;
    marginTop: number;
    marginBottom: number;
    transform?: string;
    transformOrigin?: string;
  }>;
}

export function heldPageFrameLayout(
  frame: PdfHeldPageFrame,
  pageWidth: number,
  pageHeight: number,
): PdfHeldPageFrameLayout {
  const k = frame.scale;
  const left = frame.x * pageWidth, top = frame.y * pageHeight;
  const width = frame.width * pageWidth, height = frame.height * pageHeight;
  const right = pageWidth - left - width, bottom = pageHeight - top - height;
  const extendsPast = left < -0.5 || top < -0.5 || right < -0.5 || bottom < -0.5;
  const px = (value: number) => `${Math.round(value * 100) / 100}px`;
  return Object.freeze({
    region: Object.freeze({ left, top, width, height }),
    extends: extendsPast,
    style: Object.freeze({
      // In the page's own px: the clip turns with its transform.
      clipPath: `inset(${px(top)} ${px(right)} ${px(bottom)} ${px(left)})`,
      // The page's room in the reader is the held page's: the margins move
      // its box so that the held page starts where the page's room starts.
      marginLeft: -left * k,
      marginTop: -top * k,
      marginRight: width * k - pageWidth + left * k,
      marginBottom: height * k - pageHeight + top * k,
      ...(k !== 1 ? { transform: `scale(${k})`, transformOrigin: "0 0" } : {}),
    }),
  });
}
