/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What Designer reads from a photo (SPEC designer-ms-parity §1.3 item 2):
 * where its subject is, so a crop keeps it, and the colour it is about, so
 * an idea can use it as an accent.
 *
 * PowerPoint uses face detection; MRAI has no vision model in the browser,
 * so the subject is found with a plain saliency heuristic on a 64-pixel
 * thumbnail: each pixel weighs its local contrast and saturation, skin
 * tones (YCbCr range) weigh three times as much as a face would, and a
 * gentle centre prior breaks ties. The weighted centroid is the focus. The
 * accent is the most-populated saturated hue bin, as its mean colour.
 */

export interface DesignerPictureReading {
  focus: { x: number; y: number };
  accent?: string;
}

const SIZE = 64;

/** The reading of an RGBA pixel buffer `width` x `height` (pure, for tests). */
export function readDesignerPicturePixels(data: Uint8ClampedArray | Uint8Array, width: number, height: number): DesignerPictureReading {
  const lum = new Float32Array(width * height);
  for (let i = 0; i < width * height; i += 1) {
    lum[i] = 0.299 * data[i * 4] + 0.587 * data[i * 4 + 1] + 0.114 * data[i * 4 + 2];
  }
  let sum = 0;
  let sx = 0;
  let sy = 0;
  const bins = new Map<number, { count: number; r: number; g: number; b: number }>();
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const i = y * width + x;
      const r = data[i * 4];
      const g = data[i * 4 + 1];
      const b = data[i * 4 + 2];
      const a = data[i * 4 + 3];
      if (a < 32) continue;
      const max = Math.max(r, g, b);
      const min = Math.min(r, g, b);
      const saturation = max === 0 ? 0 : (max - min) / max;
      const left = lum[y * width + Math.max(0, x - 1)];
      const right = lum[y * width + Math.min(width - 1, x + 1)];
      const up = lum[Math.max(0, y - 1) * width + x];
      const down = lum[Math.min(height - 1, y + 1) * width + x];
      const contrast = (Math.abs(right - left) + Math.abs(down - up)) / 255;
      // Skin: Cb 77..127, Cr 133..173 (the classic YCbCr skin box).
      const cb = 128 - 0.168736 * r - 0.331264 * g + 0.5 * b;
      const cr = 128 + 0.5 * r - 0.418688 * g - 0.081312 * b;
      const skin = cb >= 77 && cb <= 127 && cr >= 133 && cr <= 173 && max > 60 ? 1 : 0;
      const dx = x / (width - 1) - 0.5;
      const dy = y / (height - 1) - 0.5;
      const centre = 1 - Math.min(1, Math.sqrt(dx * dx + dy * dy) * 1.2);
      const weight = (contrast * 2 + saturation * 0.6 + skin * 2.4) * (0.6 + 0.4 * centre);
      sum += weight;
      sx += weight * (x + 0.5);
      sy += weight * (y + 0.5);
      if (saturation > 0.35 && max > 50 && max < 245) {
        let hue = 0;
        if (max === r) hue = ((g - b) / (max - min)) % 6;
        else if (max === g) hue = (b - r) / (max - min) + 2;
        else hue = (r - g) / (max - min) + 4;
        const bin = Math.floor(((hue * 60 + 360) % 360) / 30) * 2 + (max > 150 ? 1 : 0);
        const entry = bins.get(bin) ?? { count: 0, r: 0, g: 0, b: 0 };
        entry.count += 1;
        entry.r += r;
        entry.g += g;
        entry.b += b;
        bins.set(bin, entry);
      }
    }
  }
  const focus = sum > 0 ? { x: sx / sum / width, y: sy / sum / height } : { x: 0.5, y: 0.5 };
  let best: { count: number; r: number; g: number; b: number } | null = null;
  for (const entry of bins.values()) if (!best || entry.count > best.count) best = entry;
  const reading: DesignerPictureReading = {
    focus: { x: Math.round(focus.x * 1000) / 1000, y: Math.round(focus.y * 1000) / 1000 },
  };
  if (best && best.count >= width * height * 0.04) {
    const hex = (value: number) => Math.round(value / best!.count).toString(16).padStart(2, "0");
    reading.accent = `#${hex(best.r)}${hex(best.g)}${hex(best.b)}`;
  }
  return reading;
}

/** Read a picture the browser can load; null when it cannot be read (cross-origin, undecodable). */
export async function readDesignerPicture(src: string, timeoutMs = 2_000): Promise<DesignerPictureReading | null> {
  if (typeof document === "undefined" || typeof Image === "undefined") return null;
  const image = new Image();
  image.crossOrigin = "anonymous";
  const loaded = await new Promise<boolean>((resolve) => {
    const timer = setTimeout(() => resolve(false), timeoutMs);
    image.onload = () => {
      clearTimeout(timer);
      resolve(image.naturalWidth > 0 && image.naturalHeight > 0);
    };
    image.onerror = () => {
      clearTimeout(timer);
      resolve(false);
    };
    image.src = src;
  });
  if (!loaded) return null;
  try {
    const canvas = document.createElement("canvas");
    canvas.width = SIZE;
    canvas.height = SIZE;
    const context = canvas.getContext("2d", { willReadFrequently: true });
    if (!context) return null;
    context.drawImage(image, 0, 0, SIZE, SIZE);
    const { data } = context.getImageData(0, 0, SIZE, SIZE);
    return readDesignerPicturePixels(data, SIZE, SIZE);
  } catch {
    // A tainted canvas (a picture from another origin): no reading.
    return null;
  }
}
