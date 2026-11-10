/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Client-side "export as image (PNG)" — "输出为图片" — for the SVG-based
 * surfaces (a slide, a chart / diagram, or an HTML range wrapped into an SVG
 * `<foreignObject>`), with NO new dependency.
 *
 * Pipeline (PowerPoint / WPS "输出为图片" parity):
 *
 *   target SVG  →  serializeSvg()        // standalone XML w/ namespaces + size
 *               →  buildSvgDataUrl()     // data:image/svg+xml;base64,...
 *               →  <img>.src = dataUrl   // browser decodes the SVG
 *               →  canvas.drawImage(img) // raster at a chosen scale (high-DPI)
 *               →  canvas.toBlob(png)    // image/png blob
 *               →  downloadPngBlob()     // <a download> click
 *
 * The serialization + data-URL build + filename derivation are PURE (no canvas,
 * no network) so they are unit-tested directly. The single browser-only step is
 * `canvas.toBlob` (jsdom has no real 2D raster), which lives in
 * {@link svgElementToPngBlob} / {@link exportSvgElementToPng} behind an explicit
 * capability guard.
 *
 * SECURITY: serialization clones the live node and only ever reads attributes /
 * markup that the app itself rendered; the data URL is built from that markup.
 * No untrusted string is interpolated into executable positions.
 *
 * DEFER (noted, not implemented here):
 *   - a true high-DPI raster of the full spreadsheet DOM grid (cell borders,
 *     frozen panes, conditional formats) — that needs html2canvas-style DOM
 *     painting; the `<foreignObject>` wrapper here rasterizes inline-styled
 *     markup but does not resolve external stylesheet rules onto the clone;
 *   - 长图 / "long image" stitching of ALL slides into one tall PNG (this
 *     module exports ONE target — the current slide / the selected chart — per
 *     call; a caller could loop + stitch on a canvas as a follow-up).
 */

import {
  captureElementPicture,
  rasterElementPicture,
} from "./element-picture";

const SVG_NS = "http://www.w3.org/2000/svg";
const XLINK_NS = "http://www.w3.org/1999/xlink";

/** Options shared by the serialization + raster entry points. */
export interface SvgToPngOptions {
  /**
   * Pixel width to force onto the serialized `<svg>` (and the raster canvas,
   * times {@link scale}). Defaults to the element's own width — see
   * {@link resolveSvgSize}.
   */
  width?: number;
  /** Pixel height to force onto the serialized `<svg>`. */
  height?: number;
  /**
   * Device-pixel multiplier for the raster canvas, for a crisp ("high-DPI")
   * PNG. Clamped to `[1, 4]`; non-finite / <= 0 falls back to 1.
   */
  scale?: number;
  /**
   * Solid background painted under the image before the SVG is drawn (PNG keeps
   * alpha by default, but a slide usually wants an opaque white page). Omit /
   * pass `null` for a transparent PNG.
   */
  background?: string | null;
}

/** Clamp a requested raster scale into a sane `[1, 4]` integer-ish range. */
export function normalizeScale(scale: number | undefined): number {
  if (typeof scale !== "number" || !Number.isFinite(scale) || scale <= 0) {
    return 1;
  }
  return Math.min(4, Math.max(1, scale));
}

/**
 * Resolve the pixel `{ width, height }` to stamp onto a serialized SVG.
 *
 * Preference order, each axis independently:
 *   1. an explicit positive option (`opts.width` / `opts.height`);
 *   2. the element's own positive `width` / `height` attribute (numeric px);
 *   3. the third + fourth numbers of its `viewBox` (its intrinsic user units);
 *   4. a final `[1280, 720]` 16:9 fallback so the data URL is never zero-sized.
 *
 * Exposed for unit testing because a zero/absent size silently yields a blank
 * raster — the bug this guards against.
 */
export function resolveSvgSize(
  svg: SVGSVGElement,
  opts: SvgToPngOptions = {},
): { width: number; height: number } {
  const attrNum = (name: string): number | undefined => {
    const raw = svg.getAttribute(name);
    if (!raw) return undefined;
    // Accept a bare number or an explicit `px` length ONLY. A percentage /
    // `em` / `vh` etc. is CSS-relative and cannot define a raster size, so it
    // is ignored (we fall through to the viewBox). `parseFloat("100%")` would
    // otherwise return 100 and silently misuse a percentage as pixels.
    const trimmed = raw.trim();
    if (!/^[0-9]*\.?[0-9]+(?:px)?$/i.test(trimmed)) return undefined;
    const n = Number.parseFloat(trimmed);
    return Number.isFinite(n) && n > 0 ? n : undefined;
  };
  const viewBox = (svg.getAttribute("viewBox") || "")
    .split(/[\s,]+/)
    .map((part) => Number.parseFloat(part))
    .filter((n) => Number.isFinite(n));
  const vbWidth = viewBox.length === 4 && viewBox[2] > 0 ? viewBox[2] : undefined;
  const vbHeight =
    viewBox.length === 4 && viewBox[3] > 0 ? viewBox[3] : undefined;

  const optWidth =
    typeof opts.width === "number" && Number.isFinite(opts.width) && opts.width > 0
      ? opts.width
      : undefined;
  const optHeight =
    typeof opts.height === "number" &&
    Number.isFinite(opts.height) &&
    opts.height > 0
      ? opts.height
      : undefined;

  const width = optWidth ?? attrNum("width") ?? vbWidth ?? 1280;
  const height = optHeight ?? attrNum("height") ?? vbHeight ?? 720;
  return { width: Math.round(width), height: Math.round(height) };
}

/**
 * Serialize a live `<svg>` element into a STANDALONE XML string that a browser
 * can decode on its own:
 *
 *   - clones the node so the live DOM is untouched;
 *   - guarantees the SVG + xlink namespaces are present (a fragment rendered by
 *     React often inherits them from the document and omits them inline — a
 *     detached clone without `xmlns` will not parse as an image);
 *   - stamps explicit pixel `width`/`height` (see {@link resolveSvgSize}) AND a
 *     `viewBox` so the raster has a definite size even if the original sized
 *     itself purely by CSS (`width="100%"`);
 *   - prepends an `<?xml ...?>` declaration.
 *
 * PURE: no canvas / DOM mutation of the source / network. The argument must be
 * an `SVGSVGElement`; anything else throws (caller bug).
 */
export function serializeSvg(
  svg: SVGSVGElement,
  opts: SvgToPngOptions = {},
): string {
  if (!svg || typeof svg.cloneNode !== "function") {
    throw new TypeError("serializeSvg requires an <svg> element");
  }
  const { width, height } = resolveSvgSize(svg, opts);
  const clone = svg.cloneNode(true) as SVGSVGElement;

  clone.setAttribute("xmlns", SVG_NS);
  // Only declare xlink if the markup actually references it, but declaring it
  // unconditionally is harmless and avoids parse failures for legacy `xlink:href`.
  if (!clone.getAttribute("xmlns:xlink")) {
    clone.setAttribute("xmlns:xlink", XLINK_NS);
  }
  clone.setAttribute("width", String(width));
  clone.setAttribute("height", String(height));
  if (!clone.getAttribute("viewBox")) {
    clone.setAttribute("viewBox", `0 0 ${width} ${height}`);
  }

  const xml =
    typeof XMLSerializer !== "undefined"
      ? new XMLSerializer().serializeToString(clone)
      : // Extremely defensive: outerHTML is close enough for the (already
        // namespaced) clone in the rare environment without XMLSerializer.
        clone.outerHTML;
  return `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n${xml}`;
}

/**
 * UTF-8-safe base64 of an arbitrary string. `btoa` alone throws on multi-byte
 * characters (a slide title with CJK / emoji), so encode to UTF-8 bytes first.
 * Falls back to Node's `Buffer` when `btoa` is unavailable (SSR / tests).
 */
function base64Utf8(input: string): string {
  if (typeof TextEncoder !== "undefined" && typeof btoa === "function") {
    const bytes = new TextEncoder().encode(input);
    let binary = "";
    for (let i = 0; i < bytes.length; i += 1) {
      binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
  }
  // Node / SSR path.
  return Buffer.from(input, "utf-8").toString("base64");
}

/**
 * Build a `data:image/svg+xml;base64,...` URL from a serialized SVG string.
 *
 * Base64 (not percent-encoding) is used so the URL is robust to every byte the
 * markup may carry. PURE + deterministic — unit-tested for the prefix and a
 * round-trippable payload.
 */
export function buildSvgDataUrl(svgString: string): string {
  return `data:image/svg+xml;base64,${base64Utf8(svgString)}`;
}

/**
 * Convenience: serialize an element straight to a data URL (the two pure steps
 * the raster path runs back-to-back). Handy for callers and tests.
 */
export function svgElementToDataUrl(
  svg: SVGSVGElement,
  opts: SvgToPngOptions = {},
): string {
  return buildSvgDataUrl(serializeSvg(svg, opts));
}

/**
 * Wrap an arbitrary HTML element's inline-styled markup into a standalone
 * `<svg><foreignObject>` string of the given pixel size — the no-dependency way
 * to rasterize an HTML surface (e.g. a slide canvas of positioned `<div>`s, or
 * the clipboard `<table>` from `buildClipboardHtml`).
 *
 * The element's `outerHTML` is embedded inside an XHTML `<div xmlns="...">` so
 * the SVG parser accepts it. Inline styles + presentation attributes are
 * honored; rules that live only in an external stylesheet are NOT inlined here
 * (see the module DEFER note) — callers that need full fidelity should pass an
 * element whose styling is already inline, or accept the documented gap.
 *
 * PURE: reads `outerHTML` only; no mutation, no canvas.
 */
export function htmlMarkupToSvgString(
  innerHtml: string,
  width: number,
  height: number,
  background?: string | null,
): string {
  const w = Number.isFinite(width) && width > 0 ? Math.round(width) : 1280;
  const h = Number.isFinite(height) && height > 0 ? Math.round(height) : 720;
  const bg =
    background && /^#(?:[0-9a-fA-F]{3,8})$|^[a-zA-Z]+$/.test(background)
      ? `background:${background};`
      : "";
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n` +
    `<svg xmlns="${SVG_NS}" xmlns:xlink="${XLINK_NS}" ` +
    `width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">` +
    `<foreignObject x="0" y="0" width="${w}" height="${h}">` +
    `<div xmlns="http://www.w3.org/1999/xhtml" ` +
    `style="width:${w}px;height:${h}px;${bg}box-sizing:border-box;overflow:hidden">` +
    `${innerHtml}` +
    `</div></foreignObject></svg>`
  );
}

/**
 * Derive a safe PNG filename from a human label (a slide title / chart caption).
 * Strips characters illegal on common filesystems, collapses whitespace to `-`,
 * trims, caps length, and guarantees a `.png` suffix + a non-empty stem.
 *
 * PURE — unit-tested for the suffix, the fallback, and sanitization.
 */
export function toPngFilename(label: string | undefined, fallback = "image"): string {
  const cleaned = (label ?? "")
    .replace(/\.png$/i, "")
    // Drop path separators + characters Windows/macOS reject in names.
    .replace(/[\\/:*?"<>| -]/g, " ")
    .replace(/\s+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 120);
  const stem = cleaned.length > 0 ? cleaned : fallback;
  return `${stem}.png`;
}

/** Whether the current environment can actually raster + download a PNG. */
export function canRasterToPng(): boolean {
  return (
    typeof document !== "undefined" &&
    typeof Image !== "undefined" &&
    typeof URL !== "undefined" &&
    typeof URL.createObjectURL === "function" &&
    typeof document.createElement === "function"
  );
}

/**
 * Raster a serialized SVG data URL into a PNG `Blob` via an offscreen canvas.
 * BROWSER-ONLY (`canvas.toBlob`); resolves `null` when the environment lacks
 * the capability OR the decode/encode fails, so callers degrade gracefully.
 */
export function rasterDataUrlToPngBlob(
  dataUrl: string,
  width: number,
  height: number,
  options: { scale?: number; background?: string | null } = {},
): Promise<Blob | null> {
  if (!canRasterToPng()) return Promise.resolve(null);
  const scale = normalizeScale(options.scale);
  return new Promise<Blob | null>((resolve) => {
    const image = new Image();
    image.decoding = "async";
    image.onload = () => {
      try {
        const canvas = document.createElement("canvas");
        canvas.width = Math.max(1, Math.round(width * scale));
        canvas.height = Math.max(1, Math.round(height * scale));
        const ctx = canvas.getContext("2d");
        if (!ctx) {
          resolve(null);
          return;
        }
        if (options.background) {
          ctx.fillStyle = options.background;
          ctx.fillRect(0, 0, canvas.width, canvas.height);
        }
        ctx.drawImage(image, 0, 0, canvas.width, canvas.height);
        if (typeof canvas.toBlob === "function") {
          canvas.toBlob((blob) => resolve(blob), "image/png");
        } else {
          resolve(null);
        }
      } catch {
        // A tainted canvas (e.g. an SVG that pulled a cross-origin image)
        // throws on toBlob/export — fail soft rather than crash the editor.
        resolve(null);
      }
    };
    image.onerror = () => resolve(null);
    image.src = dataUrl;
  });
}

/**
 * Trigger a browser download of a PNG `Blob`. No-op (returns false) without a
 * DOM / object-URL support. Mirrors {@link downloadTextFile} in
 * `lib/download-file.ts` (object URL + `<a download>` + revoke).
 */
export function downloadPngBlob(filename: string, blob: Blob): boolean {
  if (
    typeof document === "undefined" ||
    typeof URL === "undefined" ||
    typeof URL.createObjectURL !== "function"
  ) {
    return false;
  }
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.rel = "noopener";
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  URL.revokeObjectURL(url);
  return true;
}

/**
 * End-to-end: serialize an `<svg>` element, raster it to PNG, and download it.
 * Returns `true` when a download was triggered, `false` when the environment
 * could not raster (so the caller can surface a fallback). Never throws on the
 * happy path — internal failures resolve to `false`.
 */
export async function exportSvgElementToPng(
  svg: SVGSVGElement,
  label: string | undefined,
  opts: SvgToPngOptions = {},
): Promise<boolean> {
  if (!canRasterToPng()) return false;
  const { width, height } = resolveSvgSize(svg, opts);
  const dataUrl = svgElementToDataUrl(svg, opts);
  const blob = await rasterDataUrlToPngBlob(dataUrl, width, height, {
    scale: opts.scale,
    background: opts.background,
  });
  if (!blob) return false;
  return downloadPngBlob(toPngFilename(label), blob);
}

/**
 * End-to-end for an HTML surface: the element as the page draws it (the
 * shared element picture: computed styles written in, pictures and faces
 * carried, editing controls left out), rastered and downloaded. Same return
 * contract as {@link exportSvgElementToPng}; never waits without end.
 */
export async function exportHtmlElementToPng(
  el: HTMLElement,
  label: string | undefined,
  opts: SvgToPngOptions = {},
): Promise<boolean> {
  if (!canRasterToPng()) return false;
  try {
    const picture = await captureElementPicture(el, {
      width: opts.width,
      height: opts.height,
      omit: "button, input, select, [role=\"toolbar\"]",
    });
    const blob = await rasterElementPicture(picture, {
      scale: normalizeScale(opts.scale),
      background: opts.background ?? "#ffffff",
    });
    return downloadPngBlob(toPngFilename(label), blob);
  } catch {
    return false;
  }
}
