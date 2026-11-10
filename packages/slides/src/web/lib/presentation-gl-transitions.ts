/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The slide transitions PowerPoint draws in 3-D, drawn with WebGL from
 * pictures of the two slides (Transitions > Exciting / Dynamic Content:
 * Vortex, Switch, Flip, Ripple, Honeycomb, Glitter-less Prism, Doors, Window,
 * Ferris Wheel, Gallery; and Dissolve / Random Bars, which are per-pixel).
 * The CSS show draws flat stand-ins for these; this draws them as PowerPoint
 * does: the slide as a surface turned, split, shattered or rippled in
 * perspective.
 *
 * One program draws textured quads (any number, each with its own corners in
 * 3-D and its own opacity) under a perspective camera; a second, full-slide
 * one, draws the per-pixel effects. `drawGlTransition` draws one frame of
 * `kind` at `progress`; the caller owns time.
 */

export const GL_TRANSITION_KINDS = [
  "vortex",
  "switch",
  "flip",
  "ripple",
  "honeycomb",
  "prism",
  "doors",
  "window",
  "ferris-wheel",
  "gallery",
  "dissolve",
  "random-bars",
] as const;
export type GlTransitionKind = (typeof GL_TRANSITION_KINDS)[number];

export function isGlTransitionKind(value: unknown): value is GlTransitionKind {
  return typeof value === "string" && (GL_TRANSITION_KINDS as readonly string[]).includes(value);
}

type Vec3 = [number, number, number];
type Quad = {
  /** Corners: top-left, top-right, bottom-right, bottom-left, in slide units (x right, y down, z towards the eye). */
  corners: [Vec3, Vec3, Vec3, Vec3];
  /** Texture rectangle [u0, v0, u1, v1]. */
  uv: [number, number, number, number];
  texture: "from" | "to";
  alpha: number;
};

export type GlTransitionRenderer = {
  draw(kind: GlTransitionKind, direction: string | null, progress: number): void;
  dispose(): void;
};

const QUAD_VERTEX = `
attribute vec3 aPosition;
attribute vec2 aUv;
attribute float aAlpha;
uniform vec2 uAspect;
varying vec2 vUv;
varying float vAlpha;
void main() {
  // Slide units: x in [-aspect, aspect], y in [-1, 1]; a camera 3.2 units away.
  float distance = 3.2;
  float depth = distance - aPosition.z;
  vec2 projected = aPosition.xy * distance / depth;
  gl_Position = vec4(projected.x / uAspect.x, -projected.y, 0.0, 1.0) * vec4(1.0, 1.0, 1.0, 1.0);
  gl_Position.z = clamp(-aPosition.z * 0.05, -0.99, 0.99);
  vUv = aUv;
  vAlpha = aAlpha;
}`;

const QUAD_FRAGMENT = `
precision mediump float;
uniform sampler2D uTexture;
varying vec2 vUv;
varying float vAlpha;
void main() {
  vec4 colour = texture2D(uTexture, vUv);
  gl_FragColor = vec4(colour.rgb, colour.a * vAlpha);
}`;

const PIXEL_VERTEX = `
attribute vec2 aPosition;
varying vec2 vUv;
void main() {
  vUv = vec2(aPosition.x * 0.5 + 0.5, 0.5 - aPosition.y * 0.5);
  gl_Position = vec4(aPosition, 0.0, 1.0);
}`;

const PIXEL_FRAGMENT = `
precision mediump float;
uniform sampler2D uFrom;
uniform sampler2D uTo;
uniform float uProgress;
uniform float uMode;
uniform float uAspect;
uniform vec2 uOrigin;
varying vec2 vUv;
float hash(vec2 p) { return fract(sin(dot(p, vec2(127.1, 311.7))) * 43758.5453); }
void main() {
  if (uMode < 0.5) {
    // Ripple: a ring spreading from the origin, the surface lensed by it.
    vec2 delta = (vUv - uOrigin) * vec2(uAspect, 1.0);
    float distance = length(delta);
    float front = uProgress * 1.6;
    float ring = exp(-pow((distance - front) * 9.0, 2.0));
    vec2 bent = vUv + normalize(delta + 1e-5) * ring * 0.035 * (1.0 - uProgress);
    float reached = smoothstep(front + 0.04, front - 0.04, distance);
    vec4 from = texture2D(uFrom, bent);
    vec4 to = texture2D(uTo, bent);
    gl_FragColor = mix(from, to, reached) + ring * 0.08;
  } else if (uMode < 1.5) {
    // Dissolve: the new slide shows through in random blocks.
    vec2 cell = floor(vUv * vec2(96.0 * uAspect, 96.0));
    float threshold = hash(cell);
    gl_FragColor = threshold < uProgress ? texture2D(uTo, vUv) : texture2D(uFrom, vUv);
  } else {
    // Random bars: lines across (uMode 2) or down (3) in random order.
    float coordinate = uMode < 2.5 ? vUv.y : vUv.x;
    float bar = floor(coordinate * 160.0);
    gl_FragColor = hash(vec2(bar, 7.0)) < uProgress ? texture2D(uTo, vUv) : texture2D(uFrom, vUv);
  }
}`;

function compile(gl: WebGLRenderingContext, vertex: string, fragment: string): WebGLProgram {
  const program = gl.createProgram()!;
  for (const [type, source] of [
    [gl.VERTEX_SHADER, vertex],
    [gl.FRAGMENT_SHADER, fragment],
  ] as const) {
    const shader = gl.createShader(type)!;
    gl.shaderSource(shader, source);
    gl.compileShader(shader);
    if (!gl.getShaderParameter(shader, gl.COMPILE_STATUS)) throw new Error(gl.getShaderInfoLog(shader) ?? "shader");
    gl.attachShader(program, shader);
  }
  gl.linkProgram(program);
  if (!gl.getProgramParameter(program, gl.LINK_STATUS)) throw new Error(gl.getProgramInfoLog(program) ?? "program");
  return program;
}

function texture(gl: WebGLRenderingContext, image: TexImageSource): WebGLTexture {
  const handle = gl.createTexture()!;
  gl.bindTexture(gl.TEXTURE_2D, handle);
  gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.CLAMP_TO_EDGE);
  gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.CLAMP_TO_EDGE);
  gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.LINEAR);
  gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.LINEAR);
  gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA, gl.RGBA, gl.UNSIGNED_BYTE, image);
  return handle;
}

const ease = (t: number) => (t < 0.5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2);
const clamp01 = (t: number) => Math.min(1, Math.max(0, t));

/** A quad turned by `angle` (radians) about a vertical axis at x = `pivotX`, z = `pivotZ`. */
function turnedY(
  rect: { left: number; right: number; top: number; bottom: number },
  angle: number,
  pivotX: number,
  pivotZ = 0,
  shift: Vec3 = [0, 0, 0],
): [Vec3, Vec3, Vec3, Vec3] {
  const point = (x: number, y: number): Vec3 => {
    const dx = x - pivotX;
    const dz = 0 - pivotZ;
    return [
      pivotX + dx * Math.cos(angle) - dz * Math.sin(angle) + shift[0],
      y + shift[1],
      pivotZ + dx * Math.sin(angle) + dz * Math.cos(angle) + shift[2],
    ];
  };
  return [point(rect.left, rect.top), point(rect.right, rect.top), point(rect.right, rect.bottom), point(rect.left, rect.bottom)];
}

/** Rotate a flat quad about its own centre in the slide plane and push it towards the eye. */
function spun(rect: { left: number; right: number; top: number; bottom: number }, angle: number, shift: Vec3, scale: number): [Vec3, Vec3, Vec3, Vec3] {
  const cx = (rect.left + rect.right) / 2;
  const cy = (rect.top + rect.bottom) / 2;
  const point = (x: number, y: number): Vec3 => {
    const dx = (x - cx) * scale;
    const dy = (y - cy) * scale;
    return [cx + dx * Math.cos(angle) - dy * Math.sin(angle) + shift[0], cy + dx * Math.sin(angle) + dy * Math.cos(angle) + shift[1], shift[2]];
  };
  return [point(rect.left, rect.top), point(rect.right, rect.top), point(rect.right, rect.bottom), point(rect.left, rect.bottom)];
}

function whole(aspect: number): { left: number; right: number; top: number; bottom: number } {
  return { left: -aspect, right: aspect, top: -1, bottom: 1 };
}

function seeded(index: number): number {
  const x = Math.sin(index * 12.9898 + 78.233) * 43758.5453;
  return x - Math.floor(x);
}

/** The quads of one frame of a turning / splitting transition. */
export function glTransitionQuads(kind: GlTransitionKind, direction: string | null, progress: number, aspect: number): Quad[] {
  const t = ease(clamp01(progress));
  const full: [number, number, number, number] = [0, 0, 1, 1];
  const rightward = direction === "right";
  const sign = rightward ? -1 : 1;
  const slide = whole(aspect);
  switch (kind) {
    case "flip": {
      // The slide turns over: the old face away, the new one round from behind.
      const angle = sign * t * Math.PI;
      const lift = Math.sin(t * Math.PI) * 0.6;
      if (t < 0.5) return [{ corners: turnedY(slide, angle, 0, 0, [0, 0, -lift]), uv: full, texture: "from", alpha: 1 }];
      return [{ corners: turnedY(slide, angle - sign * Math.PI, 0, 0, [0, 0, -lift]), uv: full, texture: "to", alpha: 1 }];
    }
    case "switch": {
      // The old slide turns aside and back as the new one turns in front.
      const outAngle = sign * t * 1.1;
      const inAngle = -sign * (1 - t) * 1.1;
      const away = Math.sin(t * Math.PI) * 1.4;
      return [
        { corners: turnedY(slide, outAngle, 0, 0, [-sign * t * aspect * 0.6, 0, -away - t]), uv: full, texture: "from", alpha: 1 - t * 0.5 },
        { corners: turnedY(slide, inAngle, 0, 0, [sign * (1 - t) * aspect * 0.6, 0, -away * 0.5]), uv: full, texture: "to", alpha: 0.5 + t * 0.5 },
      ];
    }
    case "prism": {
      // The two slides as two faces of a turning block.
      const angle = sign * t * (Math.PI / 2);
      const pivotZ = -aspect;
      const faces: Quad[] = [
        { corners: turnedY(slide, angle, 0, pivotZ), uv: full, texture: "from", alpha: 1 },
        { corners: turnedY(slide, angle - sign * (Math.PI / 2), 0, pivotZ), uv: full, texture: "to", alpha: 1 },
      ];
      return faces.sort((a, b) => a.corners[0][2] + a.corners[2][2] - (b.corners[0][2] + b.corners[2][2]));
    }
    case "doors":
    case "window": {
      // The old slide opens in two halves on their outer edges; the new one
      // comes forward behind them (Window from further away).
      const open = t * (kind === "doors" ? 1.6 : 1.3);
      const depth = kind === "doors" ? 0.6 : 1.4;
      const leftHalf = { ...slide, right: 0 };
      const rightHalf = { ...slide, left: 0 };
      return [
        { corners: spun(slide, 0, [0, 0, -(1 - t) * depth], 1), uv: full, texture: "to", alpha: 1 },
        { corners: turnedY(leftHalf, -open, -aspect), uv: [0, 0, 0.5, 1], texture: "from", alpha: 1 - t * 0.4 },
        { corners: turnedY(rightHalf, open, aspect), uv: [0.5, 0, 1, 1], texture: "from", alpha: 1 - t * 0.4 },
      ];
    }
    case "ferris-wheel": {
      // The slides ride one wheel: the old one down and out, the new one in.
      const radius = aspect * 2.4;
      const swing = (angle: number): [Vec3, Vec3, Vec3, Vec3] => {
        const cx = Math.sin(angle) * radius;
        const cy = radius - Math.cos(angle) * radius;
        return spun(slide, angle * 0.4, [cx, cy, -Math.abs(Math.sin(angle)) * 0.8], 0.92 + 0.08 * Math.cos(angle));
      };
      return [
        { corners: swing(-sign * t * 0.9), uv: full, texture: "from", alpha: 1 },
        { corners: swing(sign * (1 - t) * 0.9), uv: full, texture: "to", alpha: 1 },
      ];
    }
    case "gallery": {
      // A gallery wall: the old slide slides off turning away, the new one in.
      return [
        { corners: turnedY(slide, sign * t * 0.7, 0, 0, [-sign * t * aspect * 2.2, 0, -t * 0.6]), uv: full, texture: "from", alpha: 1 },
        { corners: turnedY(slide, -sign * (1 - t) * 0.7, 0, 0, [sign * (1 - t) * aspect * 2.2, 0, -(1 - t) * 0.6]), uv: full, texture: "to", alpha: 1 },
      ];
    }
    case "vortex": {
      // The old slide shatters into tiles that spin away round the centre,
      // the outer ones first.
      const columns = 24;
      const rows = 14;
      const quads: Quad[] = [{ corners: spun(slide, 0, [0, 0, 0], 1), uv: full, texture: "to", alpha: 1 }];
      for (let row = 0; row < rows; row += 1) {
        for (let column = 0; column < columns; column += 1) {
          const left = -aspect + (column / columns) * 2 * aspect;
          const right = -aspect + ((column + 1) / columns) * 2 * aspect;
          const top = -1 + (row / rows) * 2;
          const bottom = -1 + ((row + 1) / rows) * 2;
          const cx = (left + right) / 2;
          const cy = (top + bottom) / 2;
          const reach = Math.hypot(cx / aspect, cy) / Math.SQRT2;
          const start = (1 - reach) * 0.45 + seeded(row * columns + column) * 0.15;
          const local = clamp01((progress - start) / 0.4);
          if (local >= 1) continue;
          const swirl = sign * local * 5;
          const radius = Math.hypot(cx, cy) * (1 + local * 2.5);
          const angle = Math.atan2(cy, cx) + swirl;
          quads.push({
            corners: spun({ left, right, top, bottom }, swirl * 1.5, [Math.cos(angle) * radius - cx, Math.sin(angle) * radius - cy, local * 2], 1 - local * 0.5),
            uv: [column / columns, row / rows, (column + 1) / columns, (row + 1) / rows],
            texture: "from",
            alpha: 1 - local,
          });
        }
      }
      return quads;
    }
    case "honeycomb": {
      // Hexagonal cells of the old slide drop away towards the eye, scattered.
      const size = 0.22;
      const quads: Quad[] = [{ corners: spun(slide, 0, [0, 0, 0], 1), uv: full, texture: "to", alpha: 1 }];
      const width = Math.sqrt(3) * size;
      let index = 0;
      for (let y = -1 - size; y < 1 + size; y += 1.5 * size) {
        const odd = Math.round((y + 1 + size) / (1.5 * size)) % 2 === 1;
        for (let x = -aspect - width; x < aspect + width; x += width) {
          const cx = x + (odd ? width / 2 : 0);
          const start = seeded(index) * 0.6;
          index += 1;
          const local = clamp01((progress - start) / 0.4);
          if (local >= 1) continue;
          // A hexagon as its bounding quad (the cell's picture, square-cut).
          const rect = { left: cx - width / 2, right: cx + width / 2, top: y - size, bottom: y + size };
          quads.push({
            corners: spun(rect, local * 0.8, [0, local * 0.3, local * 2.2], 1 - local * 0.3),
            uv: [(rect.left + aspect) / (2 * aspect), (rect.top + 1) / 2, (rect.right + aspect) / (2 * aspect), (rect.bottom + 1) / 2],
            texture: "from",
            alpha: 1 - local,
          });
        }
      }
      return quads;
    }
    default:
      return [];
  }
}

export function createGlTransitionRenderer(
  canvas: HTMLCanvasElement,
  fromImage: TexImageSource,
  toImage: TexImageSource,
): GlTransitionRenderer | null {
  const gl = canvas.getContext("webgl", { premultipliedAlpha: false, preserveDrawingBuffer: true, antialias: true });
  if (!gl) return null;
  let quadProgram: WebGLProgram;
  let pixelProgram: WebGLProgram;
  try {
    quadProgram = compile(gl, QUAD_VERTEX, QUAD_FRAGMENT);
    pixelProgram = compile(gl, PIXEL_VERTEX, PIXEL_FRAGMENT);
  } catch {
    return null;
  }
  const from = texture(gl, fromImage);
  const to = texture(gl, toImage);
  const buffer = gl.createBuffer()!;
  const screen = gl.createBuffer()!;
  gl.bindBuffer(gl.ARRAY_BUFFER, screen);
  gl.bufferData(gl.ARRAY_BUFFER, new Float32Array([-1, -1, 1, -1, -1, 1, -1, 1, 1, -1, 1, 1]), gl.STATIC_DRAW);
  const aspect = canvas.width / Math.max(1, canvas.height);

  const drawPixels = (mode: number, progress: number, origin: [number, number]) => {
    gl.useProgram(pixelProgram);
    gl.bindBuffer(gl.ARRAY_BUFFER, screen);
    const position = gl.getAttribLocation(pixelProgram, "aPosition");
    gl.enableVertexAttribArray(position);
    gl.vertexAttribPointer(position, 2, gl.FLOAT, false, 0, 0);
    gl.activeTexture(gl.TEXTURE0);
    gl.bindTexture(gl.TEXTURE_2D, from);
    gl.uniform1i(gl.getUniformLocation(pixelProgram, "uFrom"), 0);
    gl.activeTexture(gl.TEXTURE1);
    gl.bindTexture(gl.TEXTURE_2D, to);
    gl.uniform1i(gl.getUniformLocation(pixelProgram, "uTo"), 1);
    gl.uniform1f(gl.getUniformLocation(pixelProgram, "uProgress"), progress);
    gl.uniform1f(gl.getUniformLocation(pixelProgram, "uMode"), mode);
    gl.uniform1f(gl.getUniformLocation(pixelProgram, "uAspect"), aspect);
    gl.uniform2f(gl.getUniformLocation(pixelProgram, "uOrigin"), origin[0], origin[1]);
    gl.drawArrays(gl.TRIANGLES, 0, 6);
  };

  const drawQuads = (quads: Quad[]) => {
    gl.useProgram(quadProgram);
    gl.uniform2f(gl.getUniformLocation(quadProgram, "uAspect"), aspect, 1);
    gl.enable(gl.BLEND);
    gl.blendFunc(gl.SRC_ALPHA, gl.ONE_MINUS_SRC_ALPHA);
    const position = gl.getAttribLocation(quadProgram, "aPosition");
    const uv = gl.getAttribLocation(quadProgram, "aUv");
    const alpha = gl.getAttribLocation(quadProgram, "aAlpha");
    gl.bindBuffer(gl.ARRAY_BUFFER, buffer);
    for (const side of ["to", "from"] as const) {
      // Draw back to front within each picture's quads.
      const own = quads.filter((quad) => quad.texture === side);
      if (own.length === 0) continue;
      const data = new Float32Array(own.length * 6 * 6);
      let at = 0;
      for (const quad of own) {
        const [u0, v0, u1, v1] = quad.uv;
        const uvs: Array<[number, number]> = [[u0, v0], [u1, v0], [u1, v1], [u0, v1]];
        for (const corner of [0, 1, 2, 0, 2, 3]) {
          const [x, y, z] = quad.corners[corner]!;
          data.set([x, y, z, uvs[corner]![0], uvs[corner]![1], quad.alpha], at);
          at += 6;
        }
      }
      gl.bufferData(gl.ARRAY_BUFFER, data, gl.DYNAMIC_DRAW);
      gl.enableVertexAttribArray(position);
      gl.vertexAttribPointer(position, 3, gl.FLOAT, false, 24, 0);
      gl.enableVertexAttribArray(uv);
      gl.vertexAttribPointer(uv, 2, gl.FLOAT, false, 24, 12);
      gl.enableVertexAttribArray(alpha);
      gl.vertexAttribPointer(alpha, 1, gl.FLOAT, false, 24, 20);
      gl.activeTexture(gl.TEXTURE0);
      gl.bindTexture(gl.TEXTURE_2D, side === "from" ? from : to);
      gl.uniform1i(gl.getUniformLocation(quadProgram, "uTexture"), 0);
      gl.drawArrays(gl.TRIANGLES, 0, own.length * 6);
    }
  };

  return {
    draw(kind, direction, progress) {
      gl.viewport(0, 0, canvas.width, canvas.height);
      gl.clearColor(0, 0, 0, 1);
      gl.clear(gl.COLOR_BUFFER_BIT | gl.DEPTH_BUFFER_BIT);
      const p = clamp01(progress);
      if (kind === "ripple") {
        const origin: [number, number] =
          direction === "left-up" ? [0, 0] : direction === "left-down" ? [0, 1] : direction === "right-up" ? [1, 0] : direction === "right-down" ? [1, 1] : [0.5, 0.5];
        drawPixels(0, p, origin);
      } else if (kind === "dissolve") {
        drawPixels(1, p, [0.5, 0.5]);
      } else if (kind === "random-bars") {
        drawPixels(direction === "vertical" ? 3 : 2, p, [0.5, 0.5]);
      } else {
        drawQuads(glTransitionQuads(kind, direction, p, aspect));
      }
    },
    dispose() {
      gl.deleteTexture(from);
      gl.deleteTexture(to);
      gl.deleteBuffer(buffer);
      gl.deleteBuffer(screen);
      gl.deleteProgram(quadProgram);
      gl.deleteProgram(pixelProgram);
    },
  };
}
