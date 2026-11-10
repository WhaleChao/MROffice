/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Format Shape > Fill > Picture or texture fill > Texture: the gallery's 24
 * kinds (Papyrus, Canvas, Denim, ... Medium wood), each a seamless tile of its
 * own drawn here from periodic noise -- MRAI's own pictures, not Office's
 * bitmaps -- applied as Office applies a texture: the picture tiled at 100 %
 * from the top left (`a:tile tx=0 ty=0 sx=100000 sy=100000 algn=tl
 * flip=none`), so it exports as an ordinary tiled picture fill that Word,
 * Excel and PowerPoint draw the same way.
 */

import type { OfficePictureFillTile } from "./office-picture-fill-options";

/** Office's own tile for a texture: 100 %, top left, no mirror. */
export const OFFICE_TEXTURE_TILE: OfficePictureFillTile = {
  tx: 0,
  ty: 0,
  sx: 100_000,
  sy: 100_000,
  algn: "tl",
  flip: "none",
};

/** The tile's side in pixels. */
export const OFFICE_TEXTURE_SIZE = 128;

type Rgb = readonly [number, number, number];

type TextureKind =
  | "fibers"
  | "weave"
  | "twill"
  | "basket"
  | "droplets"
  | "grain"
  | "marble"
  | "speckle"
  | "cloud"
  | "mesh"
  | "bouquet"
  | "wood";

interface TextureDefinition {
  readonly id: string;
  readonly zh: string;
  readonly en: string;
  readonly kind: TextureKind;
  /** Dark and light ends of the texture's colour ramp. */
  readonly dark: Rgb;
  readonly light: Rgb;
  readonly seed: number;
}

/** The gallery in Office's order. */
export const OFFICE_TEXTURES: readonly TextureDefinition[] = [
  { id: "papyrus", zh: "纸莎草纸", en: "Papyrus", kind: "fibers", dark: [196, 178, 132], light: [236, 224, 186], seed: 1 },
  { id: "canvas", zh: "画布", en: "Canvas", kind: "weave", dark: [200, 190, 166], light: [238, 232, 214], seed: 2 },
  { id: "denim", zh: "斜纹布", en: "Denim", kind: "twill", dark: [38, 62, 112], light: [104, 132, 182], seed: 3 },
  { id: "woven-mat", zh: "编织物", en: "Woven mat", kind: "basket", dark: [150, 118, 72], light: [222, 196, 140], seed: 4 },
  { id: "water-droplets", zh: "水滴", en: "Water droplets", kind: "droplets", dark: [104, 128, 148], light: [196, 214, 226], seed: 5 },
  { id: "paper-bag", zh: "纸袋", en: "Paper bag", kind: "fibers", dark: [150, 112, 70], light: [196, 158, 108], seed: 6 },
  { id: "fish-fossil", zh: "鱼类化石", en: "Fish fossil", kind: "grain", dark: [150, 138, 120], light: [214, 204, 186], seed: 7 },
  { id: "sand", zh: "沙滩", en: "Sand", kind: "grain", dark: [196, 170, 120], light: [236, 218, 172], seed: 8 },
  { id: "green-marble", zh: "绿色大理石", en: "Green marble", kind: "marble", dark: [30, 82, 58], light: [130, 182, 150], seed: 9 },
  { id: "white-marble", zh: "白色大理石", en: "White marble", kind: "marble", dark: [170, 170, 176], light: [246, 246, 248], seed: 10 },
  { id: "brown-marble", zh: "褐色大理石", en: "Brown marble", kind: "marble", dark: [92, 56, 36], light: [184, 140, 100], seed: 11 },
  { id: "granite", zh: "花岗岩", en: "Granite", kind: "speckle", dark: [70, 70, 74], light: [196, 194, 192], seed: 12 },
  { id: "newsprint", zh: "新闻纸", en: "Newsprint", kind: "grain", dark: [196, 196, 192], light: [236, 236, 232], seed: 13 },
  { id: "recycled-paper", zh: "再生纸", en: "Recycled paper", kind: "speckle", dark: [196, 184, 160], light: [240, 234, 220], seed: 14 },
  { id: "parchment", zh: "羊皮纸", en: "Parchment", kind: "cloud", dark: [214, 190, 132], light: [246, 234, 196], seed: 15 },
  { id: "stationery", zh: "信纸", en: "Stationery", kind: "cloud", dark: [206, 214, 224], light: [244, 246, 250], seed: 16 },
  { id: "blue-tissue-paper", zh: "蓝色面巾纸", en: "Blue tissue paper", kind: "cloud", dark: [150, 182, 222], light: [218, 232, 248], seed: 17 },
  { id: "pink-tissue-paper", zh: "粉色面巾纸", en: "Pink tissue paper", kind: "cloud", dark: [226, 168, 188], light: [248, 222, 232], seed: 18 },
  { id: "purple-mesh", zh: "紫色网格", en: "Purple mesh", kind: "mesh", dark: [86, 52, 120], light: [178, 146, 210], seed: 19 },
  { id: "bouquet", zh: "花束", en: "Bouquet", kind: "bouquet", dark: [120, 60, 90], light: [240, 220, 226], seed: 20 },
  { id: "cork", zh: "软木塞", en: "Cork", kind: "speckle", dark: [150, 96, 50], light: [214, 166, 112], seed: 21 },
  { id: "walnut", zh: "胡桃", en: "Walnut", kind: "wood", dark: [62, 36, 20], light: [132, 86, 52], seed: 22 },
  { id: "oak", zh: "栎木", en: "Oak", kind: "wood", dark: [168, 120, 70], light: [222, 184, 132], seed: 23 },
  { id: "medium-wood", zh: "中等木材", en: "Medium wood", kind: "wood", dark: [120, 72, 36], light: [190, 136, 82], seed: 24 },
];

const BY_ID = new Map(OFFICE_TEXTURES.map((texture) => [texture.id, texture]));

export function officeTextureById(id: string): TextureDefinition | null {
  return BY_ID.get(id) ?? null;
}

/** A deterministic random generator (mulberry32). */
function random(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4_294_967_296;
  };
}

/** Value noise on a lattice that wraps every `periodX` x `periodY` cells, so the tile is seamless. */
function periodicNoise(seed: number, periodX: number, periodY: number): (x: number, y: number) => number {
  const next = random(seed);
  const lattice = Array.from({ length: periodX * periodY }, () => next());
  const at = (i: number, j: number) =>
    lattice[(((j % periodY) + periodY) % periodY) * periodX + (((i % periodX) + periodX) % periodX)];
  const smooth = (t: number) => t * t * (3 - 2 * t);
  return (x, y) => {
    const i = Math.floor(x);
    const j = Math.floor(y);
    const fx = smooth(x - i);
    const fy = smooth(y - j);
    const top = at(i, j) * (1 - fx) + at(i + 1, j) * fx;
    const bottom = at(i, j + 1) * (1 - fx) + at(i + 1, j + 1) * fx;
    return top * (1 - fy) + bottom * fy;
  };
}

/**
 * Fractal noise over the tile, 0..1, seamless at `size`: `baseX` / `baseY`
 * lattice cells across and down (unequal for a grain that runs one way).
 */
function fbm(seed: number, size: number, baseX: number, octaves: number, baseY = baseX): (x: number, y: number) => number {
  const layers = Array.from({ length: octaves }, (_, octave) => ({
    noise: periodicNoise(seed * 31 + octave, baseX << octave, baseY << octave),
    scaleX: (baseX << octave) / size,
    scaleY: (baseY << octave) / size,
    weight: 1 / (1 << octave),
  }));
  const total = layers.reduce((sum, layer) => sum + layer.weight, 0);
  return (x, y) =>
    layers.reduce((sum, layer) => sum + layer.noise(x * layer.scaleX, y * layer.scaleY) * layer.weight, 0) / total;
}

/** The texture's brightness 0..1 at a pixel. */
function shade(texture: TextureDefinition, size: number): (x: number, y: number) => number {
  const n = fbm(texture.seed, size, 4, 4);
  const fine = fbm(texture.seed + 97, size, 32, 2);
  const tau = Math.PI * 2;
  switch (texture.kind) {
    case "fibers": {
      // Long fibres: few cells across, many down.
      const streak = fbm(texture.seed + 7, size, 1, 3, 24);
      return (x, y) => 0.45 * n(x, y) + 0.35 * streak(x, y) + 0.2 * fine(x, y);
    }
    case "weave": {
      return (x, y) => {
        const thread = 0.5 + 0.25 * Math.sin((x / size) * tau * 32) + 0.25 * Math.sin((y / size) * tau * 32);
        return 0.5 * thread + 0.3 * n(x, y) + 0.2 * fine(x, y);
      };
    }
    case "twill": {
      return (x, y) => 0.5 + 0.3 * Math.sin(((x + y) / size) * tau * 24) * (0.6 + 0.4 * fine(x, y)) + 0.2 * (n(x, y) - 0.5);
    }
    case "basket": {
      return (x, y) => {
        const cell = size / 8;
        const cx = Math.floor(x / cell);
        const cy = Math.floor(y / cell);
        const horizontal = (cx + cy) % 2 === 0;
        const across = horizontal ? (y % cell) / cell : (x % cell) / cell;
        const strand = Math.sin(across * Math.PI * 3) * 0.5 + 0.5;
        return 0.55 * strand + 0.25 * n(x, y) + 0.2 * fine(x, y);
      };
    }
    case "droplets": {
      const next = random(texture.seed);
      const drops = Array.from({ length: 22 }, () => ({ x: next() * size, y: next() * size, r: 3 + next() * 9 }));
      return (x, y) => {
        let value = 0.35 + 0.35 * n(x, y);
        for (const drop of drops) {
          for (const dx of [-size, 0, size]) {
            for (const dy of [-size, 0, size]) {
              const d = Math.hypot(x - drop.x - dx, y - drop.y - dy) / drop.r;
              if (d < 1) value = Math.max(value, 0.55 + 0.45 * (1 - d) * (y - drop.y - dy < 0 ? 1 : 0.4));
            }
          }
        }
        return value;
      };
    }
    case "grain":
      return (x, y) => 0.35 * n(x, y) + 0.65 * fine(x, y);
    case "marble": {
      return (x, y) => {
        const turbulence = n(x, y) * 6;
        const vein = Math.abs(Math.sin(((x + y) / size) * tau * 2 + turbulence));
        return 0.25 + 0.6 * Math.pow(vein, 0.35) + 0.15 * fine(x, y);
      };
    }
    case "speckle": {
      const coarse = fbm(texture.seed + 13, size, 32, 1);
      return (x, y) => {
        const s = coarse(x, y);
        const fleck = s > 0.72 ? 0 : s < 0.25 ? 1 : 0.5;
        return 0.45 * fleck + 0.35 * n(x, y) + 0.2 * fine(x, y);
      };
    }
    case "cloud":
      return (x, y) => 0.8 * n(x, y) + 0.2 * fine(x, y);
    case "mesh": {
      return (x, y) => {
        const gx = Math.abs(Math.sin((x / size) * Math.PI * 16));
        const gy = Math.abs(Math.sin((y / size) * Math.PI * 16));
        const wire = Math.max(1 - gx, 1 - gy) > 0.85 ? 1 : 0.15;
        return 0.6 * wire + 0.25 * n(x, y) + 0.15 * fine(x, y);
      };
    }
    case "bouquet": {
      const next = random(texture.seed);
      const flowers = Array.from({ length: 18 }, () => ({ x: next() * size, y: next() * size, r: 4 + next() * 6, v: next() }));
      return (x, y) => {
        let value = 0.75 + 0.2 * n(x, y);
        for (const flower of flowers) {
          for (const dx of [-size, 0, size]) {
            for (const dy of [-size, 0, size]) {
              const d = Math.hypot(x - flower.x - dx, y - flower.y - dy) / flower.r;
              if (d < 1) value = Math.min(value, 0.1 + 0.6 * flower.v + 0.2 * d);
            }
          }
        }
        return value;
      };
    }
    case "wood": {
      // The grain runs across: the warp and the streaks change slowly along x.
      const warp = fbm(texture.seed + 5, size, 1, 3, 2);
      const streaks = fbm(texture.seed + 11, size, 2, 2, 48);
      return (x, y) => {
        const ring = Math.sin(((y + warp(x, y) * size * 0.35) / size) * tau * 9);
        const streak = streaks(x, y);
        return 0.45 + 0.35 * ring * 0.5 + 0.2 * streak;
      };
    }
  }
}

/**
 * The texture's tile as a JPEG `data:` URL (browser only: drawn on a
 * canvas), or null when it cannot be drawn here.
 */
export function officeTextureDataUrl(id: string, size = OFFICE_TEXTURE_SIZE): string | null {
  const texture = officeTextureById(id);
  if (!texture || typeof document === "undefined") return null;
  const canvas = document.createElement("canvas");
  canvas.width = size;
  canvas.height = size;
  const context = canvas.getContext("2d");
  if (!context) return null;
  const image = context.createImageData(size, size);
  const brightness = shade(texture, size);
  for (let y = 0; y < size; y += 1) {
    for (let x = 0; x < size; x += 1) {
      const t = Math.min(1, Math.max(0, brightness(x, y)));
      const offset = (y * size + x) * 4;
      for (let channel = 0; channel < 3; channel += 1) {
        image.data[offset + channel] = Math.round(texture.dark[channel] + (texture.light[channel] - texture.dark[channel]) * t);
      }
      image.data[offset + 3] = 255;
    }
  }
  context.putImageData(image, 0, 0);
  try {
    return canvas.toDataURL("image/jpeg", 0.88);
  } catch {
    return null;
  }
}

/** `texture:<id>` as a picture source, or null. */
export function officeTextureSourceId(source: string): string | null {
  return source.startsWith("texture:") && BY_ID.has(source.slice(8)) ? source.slice(8) : null;
}
