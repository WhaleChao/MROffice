/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * An animated GIF (GIF89a) writer for File > Export > Create an Animated GIF.
 *
 * PowerPoint's GIF loops for ever (the NETSCAPE2.0 application extension,
 * repeat 0); each frame here carries its own 256-colour table made by median
 * cut over its pixels, LZW-coded as the GIF89a specification states. A frame
 * the same as the one before is not written again: the one before is held
 * longer instead (a slide that sits still for five seconds is one frame).
 * Delays are in hundredths of a second, rounded on the running total so a
 * long show does not drift.
 */

export type GifFrame = {
  /** RGBA, `width * height * 4` bytes. */
  readonly rgba: Uint8ClampedArray;
  /** How long the frame shows, in milliseconds. */
  readonly durationMs: number;
};

export type GifEncoderOptions = {
  readonly width: number;
  readonly height: number;
  /** Make background transparent: a pixel under this alpha is clear. */
  readonly transparent?: boolean;
};

/** The shortest delay browsers honour (shorter ones are drawn at 100 ms). */
const MIN_DELAY_CS = 2;
const ALPHA_CLEAR = 128;

type Pending = { indexes: Uint8Array; palette: Uint8Array; transparentIndex: number; rgba: Uint8ClampedArray; ms: number };

export class GifEncoder {
  private readonly options: GifEncoderOptions;
  private readonly parts: Uint8Array[] = [];
  private pending: Pending | null = null;
  private totalMs = 0;
  private writtenCs = 0;

  constructor(options: GifEncoderOptions) {
    this.options = options;
    const { width, height } = options;
    this.parts.push(
      ascii("GIF89a"),
      u16le(width),
      u16le(height),
      // No global colour table: every frame states its own.
      Uint8Array.of(0x00, 0x00, 0x00),
      // NETSCAPE2.0: loop for ever.
      Uint8Array.of(0x21, 0xff, 0x0b),
      ascii("NETSCAPE2.0"),
      Uint8Array.of(0x03, 0x01, 0x00, 0x00, 0x00),
    );
  }

  addFrame(frame: GifFrame): void {
    if (this.pending && sameImage(this.pending.rgba, frame.rgba)) {
      this.pending.ms += frame.durationMs;
      return;
    }
    this.flush();
    const { palette, indexes, transparentIndex } = quantize(frame.rgba, this.options.transparent === true);
    this.pending = { indexes, palette, transparentIndex, rgba: frame.rgba, ms: frame.durationMs };
  }

  /** The whole file. */
  finish(): Uint8Array {
    this.flush();
    this.parts.push(Uint8Array.of(0x3b));
    return concat(this.parts);
  }

  private flush(): void {
    const frame = this.pending;
    if (!frame) return;
    this.pending = null;
    this.totalMs += frame.ms;
    const delay = Math.max(MIN_DELAY_CS, Math.round(this.totalMs / 10) - this.writtenCs);
    this.writtenCs += delay;
    const { width, height } = this.options;
    const transparent = frame.transparentIndex >= 0;
    // Graphic control: dispose to background when clear pixels must not show
    // the frame before through them.
    const packed = ((transparent ? 2 : 1) << 2) | (transparent ? 1 : 0);
    this.parts.push(
      Uint8Array.of(0x21, 0xf9, 0x04, packed),
      u16le(Math.min(0xffff, delay)),
      Uint8Array.of(transparent ? frame.transparentIndex : 0, 0x00),
      Uint8Array.of(0x2c),
      u16le(0),
      u16le(0),
      u16le(width),
      u16le(height),
      // A local table of 256 entries (size field 7: 2^(7+1)).
      Uint8Array.of(0x80 | 0x07),
      frame.palette,
      lzwEncode(frame.indexes, 8),
    );
  }
}

function sameImage(a: Uint8ClampedArray, b: Uint8ClampedArray): boolean {
  if (a.length !== b.length) return false;
  for (let index = 0; index < a.length; index += 1) {
    if (a[index] !== b[index]) return false;
  }
  return true;
}

/**
 * Median cut to 255 colours (one entry kept for clear pixels when asked),
 * then every pixel to its nearest colour, remembered per 15-bit colour.
 */
function quantize(
  rgba: Uint8ClampedArray,
  transparent: boolean,
): { palette: Uint8Array; indexes: Uint8Array; transparentIndex: number } {
  const pixels = rgba.length / 4;
  // Per 15-bit colour: how many pixels, and the sum of their own colours
  // (a box's colour is its pixels' mean, not its buckets' centres).
  const histogram = new Map<number, Bucket>();
  // A sample is enough to place the boxes; every pixel is mapped after.
  const step = Math.max(1, Math.floor(pixels / 200_000));
  for (let pixel = 0; pixel < pixels; pixel += step) {
    const at = pixel * 4;
    if (transparent && rgba[at + 3]! < ALPHA_CLEAR) continue;
    const key = ((rgba[at]! >> 3) << 10) | ((rgba[at + 1]! >> 3) << 5) | (rgba[at + 2]! >> 3);
    const bucket = histogram.get(key);
    if (bucket) {
      bucket.count += 1;
      bucket.red += rgba[at]!;
      bucket.green += rgba[at + 1]!;
      bucket.blue += rgba[at + 2]!;
    } else {
      histogram.set(key, { count: 1, red: rgba[at]!, green: rgba[at + 1]!, blue: rgba[at + 2]! });
    }
  }
  const colors = medianCut(histogram, transparent ? 255 : 256);
  const palette = new Uint8Array(256 * 3);
  colors.forEach((color, index) => {
    palette[index * 3] = color[0];
    palette[index * 3 + 1] = color[1];
    palette[index * 3 + 2] = color[2];
  });
  const transparentIndex = transparent ? colors.length : -1;
  const nearest = new Map<number, number>();
  const indexes = new Uint8Array(pixels);
  for (let pixel = 0; pixel < pixels; pixel += 1) {
    const at = pixel * 4;
    if (transparent && rgba[at + 3]! < ALPHA_CLEAR) {
      indexes[pixel] = transparentIndex;
      continue;
    }
    const key = ((rgba[at]! >> 3) << 10) | ((rgba[at + 1]! >> 3) << 5) | (rgba[at + 2]! >> 3);
    let index = nearest.get(key);
    if (index === undefined) {
      index = closest(colors, rgba[at]!, rgba[at + 1]!, rgba[at + 2]!);
      nearest.set(key, index);
    }
    indexes[pixel] = index;
  }
  return { palette, indexes, transparentIndex };
}

type Rgb = [number, number, number];
type Bucket = { count: number; red: number; green: number; blue: number };
type ColorBox = { entries: Array<[key: number, bucket: Bucket]>; count: number };

function channel(key: number, which: 0 | 1 | 2): number {
  const value = which === 0 ? (key >> 10) & 31 : which === 1 ? (key >> 5) & 31 : key & 31;
  return (value << 3) | (value >> 2);
}

function medianCut(histogram: Map<number, Bucket>, limit: number): Rgb[] {
  if (histogram.size === 0) return [[0, 0, 0]];
  const all = Array.from(histogram.entries());
  const weight = (entries: ColorBox["entries"]) => entries.reduce((sum, [, bucket]) => sum + bucket.count, 0);
  const boxes: ColorBox[] = [{ entries: all, count: weight(all) }];
  while (boxes.length < limit) {
    // Split the box holding the most pixels that can still be split.
    let target = -1;
    for (let index = 0; index < boxes.length; index += 1) {
      if (boxes[index]!.entries.length > 1 && (target < 0 || boxes[index]!.count > boxes[target]!.count)) target = index;
    }
    if (target < 0) break;
    const { entries } = boxes[target]!;
    let widest: 0 | 1 | 2 = 0;
    let span = -1;
    for (const which of [0, 1, 2] as const) {
      let low = 255;
      let high = 0;
      for (const [key] of entries) {
        const value = channel(key, which);
        if (value < low) low = value;
        if (value > high) high = value;
      }
      if (high - low > span) {
        span = high - low;
        widest = which;
      }
    }
    entries.sort((a, b) => channel(a[0], widest) - channel(b[0], widest));
    const half = boxes[target]!.count / 2;
    let running = 0;
    let cut = 1;
    for (; cut < entries.length; cut += 1) {
      running += entries[cut - 1]![1].count;
      if (running >= half) break;
    }
    // The heaviest colour sorted last holds the half by itself: it goes on its
    // own. Cut past it, the box split into itself and an empty one 255 times
    
    cut = Math.min(cut, entries.length - 1);
    const left = entries.slice(0, cut);
    const right = entries.slice(cut);
    boxes.splice(
      target,
      1,
      { entries: left, count: weight(left) },
      { entries: right, count: weight(right) },
    );
  }
  return boxes.map(({ entries, count }) => {
    let red = 0;
    let green = 0;
    let blue = 0;
    for (const [, bucket] of entries) {
      red += bucket.red;
      green += bucket.green;
      blue += bucket.blue;
    }
    return [Math.round(red / count), Math.round(green / count), Math.round(blue / count)];
  });
}

function closest(colors: readonly Rgb[], red: number, green: number, blue: number): number {
  let best = 0;
  let bestDistance = Infinity;
  for (let index = 0; index < colors.length; index += 1) {
    const [r, g, b] = colors[index]!;
    const distance = (r - red) ** 2 * 2 + (g - green) ** 2 * 4 + (b - blue) ** 2 * 3;
    if (distance < bestDistance) {
      bestDistance = distance;
      best = index;
    }
  }
  return best;
}

/** Variable-length LZW (GIF89a appendix F), packed into 255-byte sub-blocks. */
function lzwEncode(indexes: Uint8Array, minCodeSize: number): Uint8Array {
  const clear = 1 << minCodeSize;
  const end = clear + 1;
  const out: number[] = [];
  let bitBuffer = 0;
  let bitCount = 0;
  let codeSize = minCodeSize + 1;
  const emit = (code: number) => {
    bitBuffer |= code << bitCount;
    bitCount += codeSize;
    while (bitCount >= 8) {
      out.push(bitBuffer & 0xff);
      bitBuffer >>>= 8;
      bitCount -= 8;
    }
  };
  let dictionary = new Map<number, number>();
  let next = end + 1;
  emit(clear);
  let prefix = indexes.length > 0 ? indexes[0]! : 0;
  for (let at = 1; at < indexes.length; at += 1) {
    const symbol = indexes[at]!;
    const key = (prefix << 8) | symbol;
    const found = dictionary.get(key);
    if (found !== undefined) {
      prefix = found;
      continue;
    }
    emit(prefix);
    if (next < 4096) {
      dictionary.set(key, next);
      if (next === 1 << codeSize && codeSize < 12) codeSize += 1;
      next += 1;
    } else {
      emit(clear);
      dictionary = new Map();
      next = end + 1;
      codeSize = minCodeSize + 1;
    }
    prefix = symbol;
  }
  emit(prefix);
  emit(end);
  if (bitCount > 0) out.push(bitBuffer & 0xff);
  const blocks: number[] = [minCodeSize];
  for (let at = 0; at < out.length; at += 255) {
    const size = Math.min(255, out.length - at);
    blocks.push(size);
    for (let index = 0; index < size; index += 1) blocks.push(out[at + index]!);
  }
  blocks.push(0);
  return Uint8Array.from(blocks);
}

function u16le(value: number): Uint8Array {
  return Uint8Array.of(value & 0xff, (value >>> 8) & 0xff);
}

function ascii(text: string): Uint8Array {
  return new TextEncoder().encode(text);
}

function concat(parts: readonly Uint8Array[]): Uint8Array {
  let length = 0;
  for (const part of parts) length += part.length;
  const out = new Uint8Array(length);
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}
