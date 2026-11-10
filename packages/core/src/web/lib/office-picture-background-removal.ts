/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Picture Format ▸ Remove Background, shared by the Office hosts.
 *
 * PowerPoint (and Word/Excel) open a Background Removal session on a picture:
 * a marquee the foreground is looked for inside, an automatic estimate that
 * tints the background magenta, the pencils Mark Areas to Keep / Mark Areas to
 * Remove, then Keep Changes or Discard All Changes. The file keeps the result
 * as the picture and the session as parameters on the ORIGINAL
 * (`a:blip/a:extLst/a14:imgProps/a14:imgLayer/a14:imgEffect/a14:backgroundRemoval`,
 * MS-ODRAWXML 2.3.1.x): the marquee as `t b l r` and every pencil stroke as
 * line segments `x1 y1 x2 y2` (`a14:foregroundMark` / `a14:backgroundMark`),
 * all in thousandths of a percent of the picture. A point mark is a segment
 * whose two ends coincide. PowerPoint's default marquee is 10 % in from every
 * side (real decks: `t="10000" b="90000" l="10000" r="90000"`).
 *
 * This module is pure (no DOM): it takes RGBA pixels and returns the alpha
 * mask, so every host draws and stores the result its own way. The estimate is
 * a two-colour-model segmentation (the "GrabCut" family, without the graph
 * cut): pixels outside the marquee and under a Remove stroke are background,
 * pixels under a Keep stroke are foreground; colour histograms of the two
 * classes are re-estimated a few rounds, the label field is smoothed, and
 * foreground islands not joined to the main subject (or to a Keep stroke) are
 * dropped -- what PowerPoint's first estimate shows for a photo of one object.
 */

/** A pencil stroke segment in 0..1 picture fractions (MS-ODRAWXML marks). */
export interface BackgroundRemovalSegment {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

export interface BackgroundRemovalSpec {
  /** The marquee: edges as 0..1 fractions from the picture's top / left. */
  rect: { t: number; b: number; l: number; r: number };
  /** Mark Areas to Keep (`a14:foregroundMark`). */
  keep: BackgroundRemovalSegment[];
  /** Mark Areas to Remove (`a14:backgroundMark`). */
  remove: BackgroundRemovalSegment[];
}

export interface RgbaPixels {
  width: number;
  height: number;
  /** RGBA, 4 bytes per pixel, row-major. */
  data: Uint8ClampedArray | Uint8Array;
}

/** PowerPoint's opening marquee: 10 % in from every side. */
export const DEFAULT_BACKGROUND_REMOVAL_RECT = Object.freeze({
  t: 0.1,
  b: 0.9,
  l: 0.1,
  r: 0.9,
});

/** The most marks a session keeps (MS-ODRAWXML sets no bound; the page does). */
export const MAX_BACKGROUND_REMOVAL_MARKS = 2_000;

const clamp01 = (value: unknown): number | null =>
  typeof value === "number" && Number.isFinite(value)
    ? Math.min(1, Math.max(0, value))
    : null;

function normalizeSegments(value: unknown): BackgroundRemovalSegment[] {
  if (!Array.isArray(value)) return [];
  const out: BackgroundRemovalSegment[] = [];
  for (const raw of value.slice(0, MAX_BACKGROUND_REMOVAL_MARKS)) {
    if (!raw || typeof raw !== "object") continue;
    const r = raw as Record<string, unknown>;
    const x1 = clamp01(r.x1);
    const y1 = clamp01(r.y1);
    const x2 = clamp01(r.x2);
    const y2 = clamp01(r.y2);
    if (x1 === null || y1 === null || x2 === null || y2 === null) continue;
    out.push({ x1, y1, x2, y2 });
  }
  return out;
}

/** A stored session, or `null` when it is not one. */
export function normalizeBackgroundRemovalSpec(
  value: unknown,
): BackgroundRemovalSpec | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const rect = record.rect as Record<string, unknown> | undefined;
  if (!rect || typeof rect !== "object") return null;
  const t = clamp01(rect.t);
  const b = clamp01(rect.b);
  const l = clamp01(rect.l);
  const r = clamp01(rect.r);
  if (t === null || b === null || l === null || r === null || b <= t || r <= l) {
    return null;
  }
  return {
    rect: { t, b, l, r },
    keep: normalizeSegments(record.keep),
    remove: normalizeSegments(record.remove),
  };
}

export function defaultBackgroundRemovalSpec(): BackgroundRemovalSpec {
  return { rect: { ...DEFAULT_BACKGROUND_REMOVAL_RECT }, keep: [], remove: [] };
}

/** A freehand stroke (points in 0..1) as the file's line segments. */
export function strokeToSegments(
  points: ReadonlyArray<{ x: number; y: number }>,
): BackgroundRemovalSegment[] {
  const clean = points
    .map((p) => ({ x: clamp01(p.x), y: clamp01(p.y) }))
    .filter((p): p is { x: number; y: number } => p.x !== null && p.y !== null);
  if (clean.length === 0) return [];
  if (clean.length === 1) {
    return [{ x1: clean[0].x, y1: clean[0].y, x2: clean[0].x, y2: clean[0].y }];
  }
  const segments: BackgroundRemovalSegment[] = [];
  for (let i = 1; i < clean.length; i += 1) {
    segments.push({ x1: clean[i - 1].x, y1: clean[i - 1].y, x2: clean[i].x, y2: clean[i].y });
  }
  return segments;
}

const BINS = 16; // per channel: 4096 colour bins
const binOf = (r: number, g: number, b: number) =>
  ((r >> 4) * BINS + (g >> 4)) * BINS + (b >> 4);

/** Rasterize segments onto a label field with a brush radius (pixels). */
function paintSegments(
  labels: Int8Array,
  width: number,
  height: number,
  segments: readonly BackgroundRemovalSegment[],
  value: number,
  radius: number,
): void {
  for (const s of segments) {
    const x1 = s.x1 * (width - 1);
    const y1 = s.y1 * (height - 1);
    const x2 = s.x2 * (width - 1);
    const y2 = s.y2 * (height - 1);
    const steps = Math.max(1, Math.ceil(Math.hypot(x2 - x1, y2 - y1)));
    for (let i = 0; i <= steps; i += 1) {
      const cx = x1 + ((x2 - x1) * i) / steps;
      const cy = y1 + ((y2 - y1) * i) / steps;
      const minX = Math.max(0, Math.floor(cx - radius));
      const maxX = Math.min(width - 1, Math.ceil(cx + radius));
      const minY = Math.max(0, Math.floor(cy - radius));
      const maxY = Math.min(height - 1, Math.ceil(cy + radius));
      for (let y = minY; y <= maxY; y += 1) {
        for (let x = minX; x <= maxX; x += 1) {
          if ((x - cx) ** 2 + (y - cy) ** 2 <= radius * radius) {
            labels[y * width + x] = value;
          }
        }
      }
    }
  }
}

/** Grow mark seeds into similar-coloured neighbours (within the marquee). */
function growMarks(
  marks: Int8Array,
  hard: Int8Array,
  data: Uint8ClampedArray | Uint8Array,
  width: number,
  height: number,
): void {
  const limit = 48 * 48;
  for (const value of [-1, 1] as const) {
    const queue: number[] = [];
    const seen = new Uint8Array(marks.length);
    for (let i = 0; i < marks.length; i += 1) {
      if (marks[i] === value) {
        queue.push(i);
        seen[i] = 1;
      }
    }
    for (let head = 0; head < queue.length; head += 1) {
      const i = queue[head];
      const x = i % width;
      const y = (i - x) / width;
      for (const n of [
        x > 0 ? i - 1 : -1,
        x < width - 1 ? i + 1 : -1,
        y > 0 ? i - width : -1,
        y < height - 1 ? i + width : -1,
      ]) {
        if (n < 0 || seen[n]) continue;
        seen[n] = 1;
        const dr = data[n * 4] - data[i * 4];
        const dg = data[n * 4 + 1] - data[i * 4 + 1];
        const db = data[n * 4 + 2] - data[i * 4 + 2];
        if (dr * dr + dg * dg + db * db > limit) continue;
        // A Keep never reaches outside the marquee; a Remove may not undo
        // a Keep stroke.
        if (value === 1 && hard[n] === -1 && marks[n] !== 1) continue;
        if (value === -1 && marks[n] === 1) continue;
        hard[n] = value;
        queue.push(n);
      }
    }
  }
}

/**
 * The foreground mask for a session: one byte per pixel, 255 kept, 0
 * removed. Works on the pixels given; a host passes a reduced copy for the
 * live preview and scales the mask up (see {@link scaleMask}).
 */
export function computeBackgroundRemovalMask(
  pixels: RgbaPixels,
  spec: BackgroundRemovalSpec,
): Uint8Array {
  const { width, height, data } = pixels;
  const count = width * height;
  // Hard constraints: 1 foreground, -1 background, 0 open.
  const hard = new Int8Array(count);
  const top = Math.floor(spec.rect.t * height);
  const bottom = Math.ceil(spec.rect.b * height);
  const left = Math.floor(spec.rect.l * width);
  const right = Math.ceil(spec.rect.r * width);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      if (y < top || y >= bottom || x < left || x >= right) hard[y * width + x] = -1;
    }
  }
  // Transparent pixels are background already.
  for (let i = 0; i < count; i += 1) if (data[i * 4 + 3] < 16) hard[i] = -1;
  const brush = Math.max(1, Math.round(Math.min(width, height) * 0.015));
  paintSegments(hard, width, height, spec.remove, -1, brush);
  paintSegments(hard, width, height, spec.keep, 1, brush);
  // A mark claims the region it was drawn in, not just the pencil's width:
  // grow each mark through neighbours of a close colour (PowerPoint's marks
  // take the area they touch).
  const marks = new Int8Array(count);
  paintSegments(marks, width, height, spec.remove, -1, brush);
  paintSegments(marks, width, height, spec.keep, 1, brush);
  growMarks(marks, hard, data, width, height);

  const bins = new Int32Array(count);
  for (let i = 0; i < count; i += 1) {
    bins[i] = binOf(data[i * 4], data[i * 4 + 1], data[i * 4 + 2]);
  }
  // Initial labels: inside the marquee leans foreground toward its centre;
  // a band just inside its edge counts as background evidence too (the
  // marquee is drawn around the subject, so its rim is mostly backdrop).
  const cx = (left + right) / 2;
  const cy = (top + bottom) / 2;
  const hw = Math.max(1, (right - left) / 2);
  const hh = Math.max(1, (bottom - top) / 2);
  const prior = new Float32Array(count);
  const fg = new Uint8Array(count);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const i = y * width + x;
      const d = Math.hypot((x - cx) / hw, (y - cy) / hh); // 0 centre .. ~1.4 corner
      prior[i] = Math.log(Math.max(1e-3, 1.15 - d) / Math.max(1e-3, d + 0.15));
      fg[i] = hard[i] === 1 || (hard[i] === 0 && d < 0.7) ? 1 : 0;
    }
  }
  const fgHist = new Float64Array(BINS * BINS * BINS);
  const bgHist = new Float64Array(BINS * BINS * BINS);
  for (let round = 0; round < 4; round += 1) {
    fgHist.fill(1);
    bgHist.fill(1);
    let fgTotal = fgHist.length;
    let bgTotal = bgHist.length;
    for (let i = 0; i < count; i += 1) {
      if (fg[i]) {
        fgHist[bins[i]] += 1;
        fgTotal += 1;
      } else {
        bgHist[bins[i]] += hard[i] === -1 ? 2 : 1;
        bgTotal += hard[i] === -1 ? 2 : 1;
      }
    }
    for (let i = 0; i < count; i += 1) {
      if (hard[i] !== 0) {
        fg[i] = hard[i] === 1 ? 1 : 0;
        continue;
      }
      const ratio =
        Math.log(fgHist[bins[i]] / fgTotal) - Math.log(bgHist[bins[i]] / bgTotal);
      fg[i] = ratio + prior[i] * 0.6 > 0 ? 1 : 0;
    }
  }
  // Smooth the label field (majority of the 3x3 neighbourhood), twice.
  let current = fg;
  for (let pass = 0; pass < 2; pass += 1) {
    const next = new Uint8Array(count);
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) {
        const i = y * width + x;
        if (hard[i] !== 0) {
          next[i] = hard[i] === 1 ? 1 : 0;
          continue;
        }
        let votes = 0;
        let seen = 0;
        for (let dy = -1; dy <= 1; dy += 1) {
          const yy = y + dy;
          if (yy < 0 || yy >= height) continue;
          for (let dx = -1; dx <= 1; dx += 1) {
            const xx = x + dx;
            if (xx < 0 || xx >= width) continue;
            votes += current[yy * width + xx];
            seen += 1;
          }
        }
        next[i] = votes * 2 > seen ? 1 : 0;
      }
    }
    current = next;
  }
  // Keep the subject: the largest foreground component and every component a
  // Keep stroke touches; drop the specks a colour model leaves elsewhere.
  const component = new Int32Array(count).fill(-1);
  const sizes: number[] = [];
  const kept: boolean[] = [];
  const stack: number[] = [];
  for (let start = 0; start < count; start += 1) {
    if (!current[start] || component[start] !== -1) continue;
    const id = sizes.length;
    let size = 0;
    let touchesKeep = false;
    component[start] = id;
    stack.push(start);
    while (stack.length) {
      const i = stack.pop()!;
      size += 1;
      if (hard[i] === 1) touchesKeep = true;
      const x = i % width;
      const y = (i - x) / width;
      const neighbours = [
        x > 0 ? i - 1 : -1,
        x < width - 1 ? i + 1 : -1,
        y > 0 ? i - width : -1,
        y < height - 1 ? i + width : -1,
      ];
      for (const n of neighbours) {
        if (n >= 0 && current[n] && component[n] === -1) {
          component[n] = id;
          stack.push(n);
        }
      }
    }
    sizes.push(size);
    kept.push(touchesKeep);
  }
  const largest = sizes.length ? Math.max(...sizes) : 0;
  const mask = new Uint8Array(count);
  for (let i = 0; i < count; i += 1) {
    const id = component[i];
    if (id < 0) continue;
    if (kept[id] || sizes[id] === largest || sizes[id] >= largest * 0.25) mask[i] = 255;
  }
  return mask;
}

/** Nearest-neighbour scale of a mask to another size, with a 1-px feather. */
export function scaleMask(
  mask: Uint8Array,
  fromWidth: number,
  fromHeight: number,
  toWidth: number,
  toHeight: number,
): Uint8Array {
  const out = new Uint8Array(toWidth * toHeight);
  for (let y = 0; y < toHeight; y += 1) {
    const sy = Math.min(fromHeight - 1, Math.floor((y * fromHeight) / toHeight));
    for (let x = 0; x < toWidth; x += 1) {
      const sx = Math.min(fromWidth - 1, Math.floor((x * fromWidth) / toWidth));
      out[y * toWidth + x] = mask[sy * fromWidth + sx];
    }
  }
  // Feather: average with the 4-neighbourhood so the cut edge is not jagged.
  const feathered = new Uint8Array(out.length);
  for (let y = 0; y < toHeight; y += 1) {
    for (let x = 0; x < toWidth; x += 1) {
      const i = y * toWidth + x;
      let sum = out[i] * 4;
      let weight = 4;
      for (const [dx, dy] of [[-1, 0], [1, 0], [0, -1], [0, 1]] as const) {
        const xx = x + dx;
        const yy = y + dy;
        if (xx < 0 || yy < 0 || xx >= toWidth || yy >= toHeight) continue;
        sum += out[yy * toWidth + xx];
        weight += 1;
      }
      feathered[i] = Math.round(sum / weight);
    }
  }
  return feathered;
}

/** The picture with the mask applied to its alpha channel (a copy). */
export function applyBackgroundRemovalMask(
  pixels: RgbaPixels,
  mask: Uint8Array,
): Uint8ClampedArray {
  const out = new Uint8ClampedArray(pixels.data);
  for (let i = 0; i < mask.length; i += 1) {
    out[i * 4 + 3] = Math.round((out[i * 4 + 3] * mask[i]) / 255);
  }
  return out;
}

/**
 * Edit Picture ▸ Erase: brush over an object and the picture is filled in from
 * what surrounds it. The masked pixels are filled inward from their border by
 * repeated neighbour averaging (a diffusion inpaint) -- an honest local fill,
 * not a generated one.
 */
export function erasePictureRegion(
  pixels: RgbaPixels,
  segments: readonly BackgroundRemovalSegment[],
  radiusFraction = 0.03,
): Uint8ClampedArray {
  const { width, height } = pixels;
  const out = new Uint8ClampedArray(pixels.data);
  const hole = new Int8Array(width * height);
  const radius = Math.max(1, Math.round(Math.min(width, height) * radiusFraction));
  paintSegments(hole, width, height, segments, 1, radius);
  const known = new Uint8Array(width * height);
  for (let i = 0; i < known.length; i += 1) known[i] = hole[i] ? 0 : 1;
  // Peel the hole from its border inward: each pass fills the hole pixels
  // that touch a known one with the mean of their known neighbours.
  let remaining = known.reduce((sum, value) => sum + (value ? 0 : 1), 0);
  let guard = width + height;
  while (remaining > 0 && guard-- > 0) {
    const fill: Array<[number, number, number, number, number]> = [];
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) {
        const i = y * width + x;
        if (known[i]) continue;
        let r = 0;
        let g = 0;
        let b = 0;
        let a = 0;
        let n = 0;
        for (let dy = -1; dy <= 1; dy += 1) {
          for (let dx = -1; dx <= 1; dx += 1) {
            const xx = x + dx;
            const yy = y + dy;
            if (xx < 0 || yy < 0 || xx >= width || yy >= height) continue;
            const j = yy * width + xx;
            if (!known[j]) continue;
            r += out[j * 4];
            g += out[j * 4 + 1];
            b += out[j * 4 + 2];
            a += out[j * 4 + 3];
            n += 1;
          }
        }
        if (n) fill.push([i, r / n, g / n, b / n, a / n]);
      }
    }
    if (!fill.length) break;
    for (const [i, r, g, b, a] of fill) {
      out[i * 4] = r;
      out[i * 4 + 1] = g;
      out[i * 4 + 2] = b;
      out[i * 4 + 3] = a;
      known[i] = 1;
      remaining -= 1;
    }
  }
  return out;
}

/** The session as MS-ODRAWXML attributes (thousandths of a percent). */
export function backgroundRemovalSpecToOoxml(spec: BackgroundRemovalSpec): {
  t: number;
  b: number;
  l: number;
  r: number;
  foreground: Array<[number, number, number, number]>;
  background: Array<[number, number, number, number]>;
} {
  const pct = (value: number) => Math.round(Math.min(1, Math.max(0, value)) * 100_000);
  const segment = (s: BackgroundRemovalSegment): [number, number, number, number] => [
    pct(s.x1),
    pct(s.y1),
    pct(s.x2),
    pct(s.y2),
  ];
  return {
    t: pct(spec.rect.t),
    b: pct(spec.rect.b),
    l: pct(spec.rect.l),
    r: pct(spec.rect.r),
    foreground: spec.keep.map(segment),
    background: spec.remove.map(segment),
  };
}
