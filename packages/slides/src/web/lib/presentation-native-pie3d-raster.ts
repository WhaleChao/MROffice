/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The deck's native 3-D pie, drawn: a bounded CPU ray cast of the scene the
 * chart engine planned (`ChartRenderPicture.pie3d`), laid over the engine's
 * picture. Drawing only; the scene's planning stays on the server
 * (`presentation-native-pie3d.ts`).
 */
export const NATIVE_PIE3D_PROFILE = "office-auto-pie-30-60-v1";

export interface NativePie3dScene {
  profile: typeof NATIVE_PIE3D_PROFILE;
  fingerprint: string;
  left: number;
  top: number;
  width: number;
  height: number;
  sin: number;
  cos: number;
  focal: number;
  distance: number;
  centerX: number;
  centerY: number;
  thickness: number;
  colors: number[][];
  ends: number[];
}

function pixel(
  scene: NativePie3dScene,
  px: number,
  py: number,
  output: Uint8ClampedArray,
  offset: number,
): void {
  const { sin: S, cos: C, distance: D, focal: f, thickness: h } = scene;
  const dx = (px - scene.centerX) / f,
    v = (py - scene.centerY) / f;
  const dy = v * S - C,
    dz = -v * C - S,
    ey = D * C,
    ez = D * S;
  const a = dx * dx + dy * dy,
    b = 2 * ey * dy,
    discriminant = b * b - 4 * a * (ey * ey - 1);
  let hit = Infinity,
    top = false;
  if (discriminant >= 0) {
    const root = Math.sqrt(discriminant);
    for (const t of [(-b - root) / (2 * a), (-b + root) / (2 * a)]) {
      const z = ez + dz * t;
      if (t > 0 && z >= 0 && z <= h) hit = Math.min(hit, t);
    }
  }
  const t = (h - ez) / dz,
    xt = dx * t,
    yt = ey + dy * t;
  if (t > 0 && t < hit && xt * xt + yt * yt <= 1) {
    hit = t;
    top = true;
  }
  if (!Number.isFinite(hit)) return;
  const x = dx * hit,
    y = ey + dy * hit;
  const nx = top ? 0 : x,
    ny = top ? -C : y * S,
    nz = top ? S : y * C;
  const diffuse = Math.min(1, 0.4 + Math.max(0, 0.45 * (-nx - ny + nz)));
  const vl = Math.hypot(dx, v, 1),
    light = 1 / Math.sqrt(3);
  const hx = -dx / vl - light,
    hy = -v / vl - light,
    hz = 1 / vl + light;
  const specular =
    0.25 *
    Math.max(0, (nx * hx + ny * hy + nz * hz) / Math.hypot(hx, hy, hz)) ** 32;
  const phase = (Math.atan2(y, x) + Math.PI / 2 + 2 * Math.PI) % (2 * Math.PI);
  let index = 0;
  while (index < scene.ends.length - 1 && phase >= scene.ends[index]) index++;
  for (let channel = 0; channel < 3; channel++)
    output[offset + channel] =
      scene.colors[index][channel] * diffuse + 255 * specular;
  output[offset + 3] = 255;
}

/** Bounded CPU ray casting. No imported images, workers, GPU, or persistent
 * cache; the existing picture cache owns the generated PNG/SVG lifetime. */
const ownRaster = Symbol("native-pie3d-raster");
interface NativePie3dRaster {
  url: string;
  scene: NativePie3dScene;
  [ownRaster]: true;
}

export async function renderNativePie3d(
  scene: NativePie3dScene,
  signal: AbortSignal,
): Promise<NativePie3dRaster> {
  signal.throwIfAborted();
  const scale = Math.min(
    2,
    Math.sqrt(2_000_000 / (scene.width * scene.height)),
  );
  const width = Math.max(1, Math.floor(scene.width * scale)),
    height = Math.max(1, Math.floor(scene.height * scale));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  try {
    const context = canvas.getContext("2d");
    if (!context) throw Error("native_pie3d_canvas_unavailable");
    const image = context.createImageData(width, height);
    for (let y = 0; y < height; y++) {
      if (y % 16 === 0) {
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
        signal.throwIfAborted();
      }
      for (let x = 0; x < width; x++)
        pixel(
          scene,
          ((x + 0.5) * scene.width) / width,
          ((y + 0.5) * scene.height) / height,
          image.data,
          (y * width + x) * 4,
        );
    }
    signal.throwIfAborted();
    context.putImageData(image, 0, 0);
    return { url: canvas.toDataURL("image/png"), scene, [ownRaster]: true };
  } finally {
    canvas.width = canvas.height = 0;
  }
}

/** Call only after the normal SVG sanitizer. Imported SVG images remain
 * forbidden; this exact transient result belongs to this computed scene. */
export function composeNativePie3dSvg(
  svg: string,
  scene: NativePie3dScene,
  raster: NativePie3dRaster,
): string {
  if (
    raster[ownRaster] !== true ||
    raster.scene !== scene ||
    !/^data:image\/png;base64,[A-Za-z0-9+/=]+$/.test(raster.url) ||
    raster.url.length > 1_500_000
  )
    throw Error("native_pie3d_raster_identity");
  const doc = new DOMParser().parseFromString(svg, "image/svg+xml");
  if (
    doc.documentElement.localName !== "svg" ||
    doc.querySelector("parsererror")
  )
    throw Error("native_pie3d_svg");
  const image = doc.createElementNS("http://www.w3.org/2000/svg", "image");
  image.setAttribute("href", raster.url);
  image.setAttribute("x", String(scene.left));
  image.setAttribute("y", String(scene.top));
  image.setAttribute("width", String(scene.width));
  image.setAttribute("height", String(scene.height));
  image.setAttribute("data-native-pie3d", scene.fingerprint);
  doc.documentElement.append(image);
  return new XMLSerializer().serializeToString(doc);
}
