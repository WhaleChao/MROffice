/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Format Shape > Fill > Picture or texture fill: the OPTIONS of a shape's
 * picture fill, one model for the deck, Docs and the worksheet.
 *
 * Office's pane (Word, Excel and PowerPoint draw the same page) offers, under
 * the picture source: Transparency (0..100 %), "Tile picture as texture",
 * then -- tiled -- Offset X / Offset Y (pt), Scale X / Scale Y (%),
 * Alignment (nine positions) and Mirror type (None / Horizontal / Vertical /
 * Both), or -- not tiled -- Offset left / right / top / bottom (% of the
 * shape's box), and "Rotate with shape".
 *
 * DrawingML (ECMA-376 §20.1.8.14 `a:blipFill`) holds them as:
 * - `a:blip/a:alphaModFix@amt` -- the picture's opacity in 1/1000 %;
 *   Transparency t is `amt = (100 - t) * 1000`. Absent: opaque.
 * - `a:tile` (§20.1.8.58) -- `tx` / `ty` an offset in EMU applied AFTER the
 *   alignment (default 0); `sx` / `sy` the scale of the picture's own size
 *   in 1/1000 % (default 100000 = 100 %); `algn` where the first tile sits
 *   in the box, after scaling and before the offset (ST_RectAlignment; Office
 *   writes `tl` and reads an absent one as top-left); `flip` none / x / y /
 *   xy, every other tile mirrored on that axis (default none).
 * - `a:stretch/a:fillRect` (§20.1.8.30) -- the rectangle the picture is
 *   stretched into, as insets of the shape's box in 1/1000 % (`l`, `t`, `r`,
 *   `b`, default 0; negative grows past the box). Office's Offset left is
 *   `l`, Offset right `r`, Offset top `t`, Offset bottom `b`.
 * - `a:blipFill@rotWithShape` -- whether the fill turns with the shape.
 *   Office writes `1`; `0` keeps the picture upright on a rotated shape.
 * `a:srcRect` is Crop (Picture Format > Crop), not part of this page.
 *
 * Each host keeps these beside its own picture record (`fillImage` on a
 * slide's shape, `fillPicture` on a Docs shape, `fillPictureOptions` on a
 * worksheet shape); an absent field is the host's behaviour before these
 * options existed. The server holds the same bounds
 * (`app/services/office_picture_fill.py`).
 */

/** ST_RectAlignment, in Office's pane order (row by row). */
export const OFFICE_PICTURE_TILE_ALIGNMENTS = ["tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"] as const;
export type OfficePictureTileAlignment = (typeof OFFICE_PICTURE_TILE_ALIGNMENTS)[number];

/** ST_TileFlipMode: Mirror type None / Horizontal / Vertical / Both. */
export const OFFICE_PICTURE_TILE_FLIPS = ["none", "x", "y", "xy"] as const;
export type OfficePictureTileFlip = (typeof OFFICE_PICTURE_TILE_FLIPS)[number];

/** `a:tile`, in the package's own units (EMU, 1/1000 %). */
export interface OfficePictureFillTile {
  readonly tx?: number;
  readonly ty?: number;
  readonly sx?: number;
  readonly sy?: number;
  readonly algn?: OfficePictureTileAlignment;
  readonly flip?: OfficePictureTileFlip;
}

/** `a:stretch/a:fillRect`: insets of the box in 1/1000 %. */
export interface OfficePictureFillRect {
  readonly l?: number;
  readonly t?: number;
  readonly r?: number;
  readonly b?: number;
}

export interface OfficePictureFillOptions {
  /** Present: "Tile picture as texture" (the stretch offsets do not apply). */
  readonly tile?: OfficePictureFillTile;
  /** The stretch's offsets (only while not tiled). */
  readonly fillRect?: OfficePictureFillRect;
  /** Whole percent 1..100; absent is opaque. */
  readonly transparency?: number;
  /** `false`: the picture stays upright on a rotated shape. Absent: turns. */
  readonly rotWithShape?: boolean;
}

/** ST_Coordinate's bound (EMU), as the package states an offset. */
export const OFFICE_PICTURE_TILE_OFFSET_LIMIT = 27_273_042_316_900;
/** ST_Percentage held to a signed 32-bit integer, as every reader parses it. */
export const OFFICE_PICTURE_PERCENT_LIMIT = 2_147_483_647;
/** The pane's bounds: Scale 1..1000 %, an offset of +/- 10000 % of the box. */
export const OFFICE_PICTURE_PANE_SCALE = { min: 1, max: 1000 } as const;
export const OFFICE_PICTURE_PANE_OFFSET_PERCENT = 10_000;
export const OFFICE_PICTURE_PANE_OFFSET_PT = 4000;

const EMU_PER_PT = 12_700;
const EMU_PER_PX = 9_525;

const integerWithin = (value: unknown, limit: number): number | null =>
  typeof value === "number" && Number.isSafeInteger(value) && Math.abs(value) <= limit ? value : null;

/** `value` as a tile, or null when any stated attribute is out of the model. */
export function normalizeOfficePictureFillTile(value: unknown): OfficePictureFillTile | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  const tile: Record<string, number | string> = {};
  for (const key of Object.keys(raw)) {
    if (!["tx", "ty", "sx", "sy", "algn", "flip"].includes(key)) return null;
  }
  for (const key of ["tx", "ty"] as const) {
    if (raw[key] === undefined) continue;
    const number = integerWithin(raw[key], OFFICE_PICTURE_TILE_OFFSET_LIMIT);
    if (number === null) return null;
    tile[key] = number;
  }
  for (const key of ["sx", "sy"] as const) {
    if (raw[key] === undefined) continue;
    const number = integerWithin(raw[key], OFFICE_PICTURE_PERCENT_LIMIT);
    if (number === null) return null;
    tile[key] = number;
  }
  if (raw.algn !== undefined) {
    if (!(OFFICE_PICTURE_TILE_ALIGNMENTS as readonly unknown[]).includes(raw.algn)) return null;
    tile.algn = raw.algn as string;
  }
  if (raw.flip !== undefined) {
    if (!(OFFICE_PICTURE_TILE_FLIPS as readonly unknown[]).includes(raw.flip)) return null;
    tile.flip = raw.flip as string;
  }
  return tile as OfficePictureFillTile;
}

/** `value` as a stretch's fill rectangle, or null. */
export function normalizeOfficePictureFillRect(value: unknown): OfficePictureFillRect | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  const rect: Record<string, number> = {};
  for (const key of Object.keys(raw)) {
    if (!["l", "t", "r", "b"].includes(key)) return null;
    const number = integerWithin(raw[key], OFFICE_PICTURE_PERCENT_LIMIT);
    if (number === null) return null;
    if (number !== 0) rect[key] = number;
  }
  return rect;
}

/**
 * The options a record states, held to the model; null when a stated option
 * is outside it (a host refuses the record rather than drawing a guess).
 * Default values collapse to absent so an untouched fill stays byte-clean.
 */
export function normalizeOfficePictureFillOptions(value: unknown): OfficePictureFillOptions | null {
  if (value === undefined || value === null) return {};
  if (typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  const out: { -readonly [K in keyof OfficePictureFillOptions]: OfficePictureFillOptions[K] } = {};
  if (raw.tile !== undefined) {
    const tile = normalizeOfficePictureFillTile(raw.tile);
    if (!tile) return null;
    out.tile = tile;
  }
  if (raw.fillRect !== undefined) {
    const rect = normalizeOfficePictureFillRect(raw.fillRect);
    if (!rect) return null;
    // A tiled picture has no stretch rectangle.
    if (!out.tile && Object.keys(rect).length) out.fillRect = rect;
  }
  if (raw.transparency !== undefined) {
    const t = raw.transparency;
    if (typeof t !== "number" || !Number.isInteger(t) || t < 0 || t > 100) return null;
    if (t > 0) out.transparency = t;
  }
  if (raw.rotWithShape !== undefined) {
    if (typeof raw.rotWithShape !== "boolean") return null;
    out.rotWithShape = raw.rotWithShape;
  }
  return out;
}

/** The option fields a host record carries, read off it (unknown keys ignored). */
export function officePictureFillOptionsOf(record: unknown): OfficePictureFillOptions {
  if (!record || typeof record !== "object") return {};
  const { tile, fillRect, transparency, rotWithShape } = record as Record<string, unknown>;
  return normalizeOfficePictureFillOptions({ tile, fillRect, transparency, rotWithShape }) ?? {};
}

/** Whether the options draw anything other than the host's plain fill. */
export function officePictureFillOptionsStated(options: OfficePictureFillOptions): boolean {
  return Boolean(options.tile || options.fillRect || options.transparency || options.rotWithShape === false);
}

/** `options` merged onto a host record: absent options leave no key behind. */
export function withOfficePictureFillOptions<T extends Record<string, unknown>>(
  record: T,
  options: OfficePictureFillOptions,
): T {
  const next: Record<string, unknown> = { ...record };
  for (const key of ["tile", "fillRect", "transparency", "rotWithShape"] as const) {
    if (options[key] === undefined) delete next[key];
    else next[key] = options[key];
  }
  return next as T;
}

// ---------------------------------------------------------------------------
// Drawing: one SVG <pattern> description every renderer paints.

export interface OfficePictureFillPatternImage {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly transform?: string;
}

/**
 * A picture fill as an SVG pattern in `objectBoundingBox` units (both the
 * tile and its content, so a host whose geometry draws in any viewBox paints
 * it over the shape's box): the tile rectangle, the images in it, the
 * pattern transform (Rotate with shape off), and the picture's opacity.
 */
export interface OfficePictureFillPattern {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly images: readonly OfficePictureFillPatternImage[];
  readonly transform?: string;
  readonly opacity?: number;
  









  readonly upright?: {
    readonly width: number;
    readonly height: number;
    readonly transform: string;
    readonly tile: { readonly x: number; readonly y: number; readonly width: number; readonly height: number };
    readonly images: readonly OfficePictureFillPatternImage[];
  };
}

const round = (value: number) => Math.round(value * 1e6) / 1e6;

/**
 * The pattern for `options` over a `width` x `height` box (px), the picture
 * `natural` px in size (96 dpi). A tile needs the natural size (null while it
 * is unknown); a stretch does not. `rotation` is the shape's own clockwise
 * rotation in degrees (only read when Rotate with shape is off).
 */
export function officePictureFillPattern(args: {
  readonly width: number;
  readonly height: number;
  readonly natural: { readonly width: number; readonly height: number } | null;
  readonly options: OfficePictureFillOptions;
  readonly rotation?: number;
}): OfficePictureFillPattern | null {
  const { width: W, height: H, options } = args;
  if (!(W > 0) || !(H > 0)) return null;
  const opacity = options.transparency ? round(1 - options.transparency / 100) : undefined;
  // Rotate with shape off keeps a TILED picture upright; a stretched one turns
  // with the shape whatever the flag says (PowerPoint's own drawing of both,
  
  const rotation = options.rotWithShape === false && options.tile ? -(args.rotation ?? 0) : 0;
  // A rotation in the box's fractional space: R conjugated by the box scale.
  const transform = rotation % 360
    ? (() => {
        const rad = (rotation * Math.PI) / 180;
        const cos = Math.cos(rad);
        const sin = Math.sin(rad);
        const a = cos, b = (sin * W) / H, c = (-sin * H) / W, d = cos;
        const e = 0.5 - a * 0.5 - c * 0.5;
        const f = 0.5 - b * 0.5 - d * 0.5;
        return `matrix(${[a, b, c, d, e, f].map(round).join(" ")})`;
      })()
    : undefined;
  if (options.tile) {
    const natural = args.natural;
    if (!natural || !(natural.width > 0) || !(natural.height > 0)) return null;
    const tile = options.tile;
    const tw = (natural.width * Math.abs(tile.sx ?? 100_000)) / 100_000 / W;
    const th = (natural.height * Math.abs(tile.sy ?? 100_000)) / 100_000 / H;
    if (!(tw > 0) || !(th > 0)) return null;
    const algn = tile.algn ?? "tl";
    const column = algn === "tr" || algn === "r" || algn === "br" ? 2 : algn === "t" || algn === "ctr" || algn === "b" ? 1 : 0;
    const row = algn === "bl" || algn === "b" || algn === "br" ? 2 : algn === "l" || algn === "ctr" || algn === "r" ? 1 : 0;
    const x = ((1 - tw) * column) / 2 + (tile.tx ?? 0) / EMU_PER_PX / W;
    const y = ((1 - th) * row) / 2 + (tile.ty ?? 0) / EMU_PER_PX / H;
    const flip = tile.flip ?? "none";
    const flipX = flip === "x" || flip === "xy";
    const flipY = flip === "y" || flip === "xy";
    const images: OfficePictureFillPatternImage[] = [{ x: 0, y: 0, width: round(tw), height: round(th) }];
    if (flipX) images.push({ x: 0, y: 0, width: round(tw), height: round(th), transform: `matrix(-1 0 0 1 ${round(2 * tw)} 0)` });
    if (flipY) images.push({ x: 0, y: 0, width: round(tw), height: round(th), transform: `matrix(1 0 0 -1 0 ${round(2 * th)})` });
    if (flipX && flipY) {
      images.push({ x: 0, y: 0, width: round(tw), height: round(th), transform: `matrix(-1 0 0 -1 ${round(2 * tw)} ${round(2 * th)})` });
    }
    const toPixels = (image: OfficePictureFillPatternImage): OfficePictureFillPatternImage => {
      const flips = /^matrix\((-?1) 0 0 (-?1) ([-\d.e]+) ([-\d.e]+)\)$/.exec(image.transform ?? "");
      return {
        x: round(image.x * W),
        y: round(image.y * H),
        width: round(image.width * W),
        height: round(image.height * H),
        ...(flips ? { transform: `matrix(${flips[1]} 0 0 ${flips[2]} ${round(Number(flips[3]) * W)} ${round(Number(flips[4]) * H)})` } : {}),
      };
    };
    return {
      x: round(x),
      y: round(y),
      width: round(tw * (flipX ? 2 : 1)),
      height: round(th * (flipY ? 2 : 1)),
      images,
      ...(transform ? { transform } : {}),
      ...(opacity !== undefined ? { opacity } : {}),
      ...(transform
        ? {
            upright: {
              width: round(W),
              height: round(H),
              transform: `rotate(${round(rotation)} ${round(W / 2)} ${round(H / 2)})`,
              tile: {
                x: round(x * W),
                y: round(y * H),
                width: round(tw * (flipX ? 2 : 1) * W),
                height: round(th * (flipY ? 2 : 1) * H),
              },
              images: images.map(toPixels),
            },
          }
        : {}),
    };
  }
  const rect = options.fillRect ?? {};
  const l = (rect.l ?? 0) / 100_000;
  const t = (rect.t ?? 0) / 100_000;
  const iw = 1 - l - (rect.r ?? 0) / 100_000;
  const ih = 1 - t - (rect.b ?? 0) / 100_000;
  if (!(iw > 0) || !(ih > 0)) return null;
  // One tile wide enough that the stretched picture never repeats, even
  // turned against the shape (a stretch is drawn once).
  const px = Math.min(0, l) - 2;
  const py = Math.min(0, t) - 2;
  const pw = Math.max(1, l + iw) - px + 2;
  const ph = Math.max(1, t + ih) - py + 2;
  return {
    x: round(px),
    y: round(py),
    width: round(pw),
    height: round(ph),
    images: [{ x: round(l - px), y: round(t - py), width: round(iw), height: round(ih) }],
    ...(transform ? { transform } : {}),
    ...(opacity !== undefined ? { opacity } : {}),
  };
}

/** A picture's pixel size in CSS px at its own resolution (96 dpi when it states none or a nonsensical one). */
function withResolution(size: { width: number; height: number }, dpiX: number, dpiY: number): { width: number; height: number } {
  const usable = (dpi: number) => Number.isFinite(dpi) && dpi >= 10 && dpi <= 3000;
  if (!usable(dpiX) || !usable(dpiY)) return size;
  return { width: (size.width * 96) / dpiX, height: (size.height * 96) / dpiY };
}

/**
 * A PNG's or JPEG's size read from its own bytes, in CSS px at the
 * resolution it states (PNG `pHYs`, JPEG JFIF density; 96 dpi when it states
 * none) -- the size Office tiles it at (PNG `IHDR`, JPEG `SOFn`) -- or null
 * for anything else. `byte(at)` is the byte at `at`, `count` how many there are.
 */
function pictureSizeFromBytes(byte: (at: number) => number, count: number): { width: number; height: number } | null {
  let size: { width: number; height: number } | null = null;
  if (count >= 24 && byte(0) === 0x89 && byte(1) === 0x50 && byte(2) === 0x4e && byte(3) === 0x47) {
    const int = (at: number) => ((byte(at) << 24) | (byte(at + 1) << 16) | (byte(at + 2) << 8) | byte(at + 3)) >>> 0;
    const type = (at: number) => String.fromCharCode(byte(at), byte(at + 1), byte(at + 2), byte(at + 3));
    if (type(12) !== "IHDR") return null;
    size = { width: int(16), height: int(20) };
    // `pHYs` (pixels per metre) before the image data: the picture's own
    // resolution, which Office sizes it by.
    let at = 8;
    while (at + 12 <= count) {
      const length = int(at);
      const kind = type(at + 4);
      if (kind === "IDAT" || kind === "IEND") break;
      if (kind === "pHYs" && length >= 9 && at + 17 <= count && byte(at + 16) === 1) {
        size = withResolution(size, int(at + 8) * 0.0254, int(at + 12) * 0.0254);
        break;
      }
      at += 12 + length;
    }
  } else if (count >= 4 && byte(0) === 0xff && byte(1) === 0xd8) {
    let jfifDpi: [number, number] | null = null;
    // JFIF's density (APP0): its units 1 are dots per inch, 2 per cm.
    if (
      count >= 18 && byte(2) === 0xff && byte(3) === 0xe0 &&
      String.fromCharCode(byte(6), byte(7), byte(8), byte(9), byte(10)) === "JFIF\0"
    ) {
      const units = byte(13);
      const x = (byte(14) << 8) | byte(15);
      const y = (byte(16) << 8) | byte(17);
      jfifDpi = units === 1 ? [x, y] : units === 2 ? [x * 2.54, y * 2.54] : null;
    }
    let at = 2;
    while (at + 9 < count) {
      if (byte(at) !== 0xff) break;
      const marker = byte(at + 1);
      const length = (byte(at + 2) << 8) | byte(at + 3);
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        size = { height: (byte(at + 5) << 8) | byte(at + 6), width: (byte(at + 7) << 8) | byte(at + 8) };
        break;
      }
      at += 2 + length;
    }
    if (size && jfifDpi) size = withResolution(size, jfifDpi[0], jfifDpi[1]);
  }
  return size && size.width > 0 && size.height > 0 ? size : null;
}

/**
 * The size of an inline PNG or JPEG picture read from its own header (see
 * `pictureSizeFromBytes`), or null. Cached per source: a tile needs it on
 * every paint and the inline spellings every host keeps are these two.
 */
const naturalSizeCache = new Map<string, { width: number; height: number } | null>();
export function officePictureNaturalSize(src: string | undefined | null): { width: number; height: number } | null {
  if (!src) return null;
  const cached = naturalSizeCache.get(src);
  if (cached !== undefined) return cached;
  let size: { width: number; height: number } | null = null;
  try {
    const match = /^data:image\/(png|jpe?g);base64,/i.exec(src);
    if (match && typeof atob === "function") {
      let body = src.slice(match[0].length);
      // A PNG states its size and resolution before its image data.
      if (match[1].toLowerCase() === "png") body = body.slice(0, 4096);
      const bytes = atob(body.slice(0, body.length - (body.length % 4)));
      size = pictureSizeFromBytes((at) => bytes.charCodeAt(at), bytes.length);
    }
  } catch {
    size = null;
  }
  if (naturalSizeCache.size > 64) naturalSizeCache.clear();
  naturalSizeCache.set(src, size);
  return size;
}







const loadedSizes = new Map<string, Promise<{ width: number; height: number } | null>>();
export function loadOfficePictureNaturalSize(src: string | undefined | null): Promise<{ width: number; height: number } | null> {
  if (!src) return Promise.resolve(null);
  if (src.startsWith("data:")) return Promise.resolve(officePictureNaturalSize(src));
  const known = loadedSizes.get(src);
  if (known) return known;
  const loading =
    typeof fetch !== "function"
      ? Promise.resolve(null)
      : fetch(src)
          .then((response) => (response.ok ? response.arrayBuffer() : null))
          .then((buffer) => {
            if (!buffer) return null;
            const bytes = new Uint8Array(buffer);
            return pictureSizeFromBytes((at) => bytes[at] ?? 0, bytes.length);
          })
          .catch(() => null);
  if (loadedSizes.size > 32) loadedSizes.clear();
  loadedSizes.set(src, loading);
  return loading;
}

// ---------------------------------------------------------------------------
// The pane's units: what Office's fields show, and back.

export const officePictureFillPaneUnits = {
  ptFromEmu: (emu: number | undefined) => Math.round(((emu ?? 0) / EMU_PER_PT) * 10) / 10,
  emuFromPt: (pt: number) => Math.round(pt * EMU_PER_PT),
  percentFromThousandths: (value: number | undefined, fallback = 0) => Math.round((value ?? fallback) / 10) / 100,
  thousandthsFromPercent: (percent: number) => Math.round(percent * 1000),
};
