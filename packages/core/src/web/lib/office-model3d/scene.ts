/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A 3-D model's scene (Insert > 3D Models; MS-ODRAWXML 2.31 `am3d:model3d`)
 * as the page carries it (`source.nativeModel3d.scene`), and the arithmetic
 * that places its model, camera and lights the way PowerPoint draws them.
 *
 * No renderer here (the module is plain numbers, safe in any bundle); the
 * WebGL side is `./renderer`, loaded only when a model is drawn.
 *
 * What the file states, and what it means (MS-ODRAWXML 2.31.3, and
 * PowerPoint's own files, model3d-truth):
 *
 * - World space is in EMUs, 36,000,000 to the metre (`am3d:camera/am3d:up`
 *   is written `dy="36000000"`, one metre). The renderer works in metres.
 * - The model (`am3d:trans`, CT_Model3DTransform) is placed by, in order:
 *   its glTF scene in model units, times `meterPerModelUnit` (metres), in
 *   EMUs; plus `preTrans` (the centre of rotation); times `scale`; turned by
 *   `rot`; plus `postTrans`.
 * - The camera (CT_Model3DCamera) looks from `pos` at `lookAt`, `up` its up;
 *   a perspective camera's `fov` is the vertical field of view (60,000ths of
 *   a degree), over a square viewport.
 * - In object mode (`am3d:objViewport`, CT_ObjectViewport) that square view
 *   is `viewportSz` EMUs on the slide, centred on the frame's centre (the
 *   look-at point maps to it): the frame shows the part of the square it
 *   covers. In window mode (`am3d:winViewport`) the frame is the view.
 * - Lights are in world space (they stay as the model turns): an ambient
 *   light of `illuminance` lux, point and spot lights of `intensity`
 *   candela, directional lights of `illuminance` lux.
 */

/** EMUs to the metre: the unit PowerPoint states a model's world in. */
export const MODEL3D_EMU_PER_METRE = 36_000_000;

export type Model3dVector = readonly [number, number, number];

export interface Model3dPerspective {
  kind: "perspective";
  /** Vertical field of view, 60,000ths of a degree, (0, 180). */
  fov: number;
}

export interface Model3dOrthographic {
  kind: "orthographic";
  /** The size of one unit relative to the viewport's height. */
  size: number;
}

export interface Model3dCamera {
  position: Model3dVector;
  up: Model3dVector;
  lookAt: Model3dVector;
  projection: Model3dPerspective | Model3dOrthographic;
}

export interface Model3dTransform {
  /** Metres per model unit (absent: 1). */
  meterPerModelUnit: number;
  preTranslation: Model3dVector;
  scale: Model3dVector;
  /** `am3d:rot` ax / ay / az, 60,000ths of a degree. */
  rotation: Model3dVector;
  postTranslation: Model3dVector;
}

export interface Model3dAmbientLight {
  /** Linear RGB, 0-1 per channel. */
  color: Model3dVector;
  /** Lux. */
  illuminance: number;
}

export interface Model3dPointLight {
  kind: "point";
  color: Model3dVector;
  /** Candela. */
  intensity: number;
  position: Model3dVector;
  /** The light's radius, EMUs. */
  radius: number;
}

export interface Model3dSpotLight {
  kind: "spot";
  color: Model3dVector;
  intensity: number;
  position: Model3dVector;
  lookAt: Model3dVector;
  radius: number;
  /** The cone's angle, 60,000ths of a degree. */
  spotAngle: number;
}

export interface Model3dDirectionalLight {
  kind: "directional";
  color: Model3dVector;
  /** Lux. */
  illuminance: number;
  position: Model3dVector;
  lookAt: Model3dVector;
}

export type Model3dLight = Model3dPointLight | Model3dSpotLight | Model3dDirectionalLight;

export interface Model3dScene {
  camera: Model3dCamera;
  transform: Model3dTransform;
  /** Object mode: the square view's size on the slide, EMUs. Window mode: null. */
  viewportSize: number | null;
  /** The model frame (`am3d:spPr/a:xfrm/a:ext`), EMUs. */
  frameWidth: number;
  frameHeight: number;
  ambientLight: Model3dAmbientLight | null;
  lights: Model3dLight[];
}

const ANGLE_UNIT = 60000;

function finite(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

function vector(value: unknown): Model3dVector | null {
  if (!Array.isArray(value) || value.length !== 3 || !value.every(finite)) return null;
  return [value[0], value[1], value[2]];
}

function colour(value: unknown): Model3dVector | null {
  const rgb = vector(value);
  return rgb && rgb.every((channel) => channel >= 0) ? rgb : null;
}

/**
 * The scene the page carries for a model, or null when it is not one this
 * renderer can draw as PowerPoint does (the picture PowerPoint kept is then
 * drawn instead). Every value is checked: a page is data from the network.
 */
export function parseModel3dScene(raw: unknown): Model3dScene | null {
  if (!raw || typeof raw !== "object") return null;
  const scene = raw as Record<string, unknown>;
  const camera = scene.camera as Record<string, unknown> | undefined;
  const transform = scene.transform as Record<string, unknown> | undefined;
  if (!camera || typeof camera !== "object" || !transform || typeof transform !== "object") return null;
  const position = vector(camera.position);
  const up = vector(camera.up);
  const lookAt = vector(camera.lookAt);
  const projectionRaw = camera.projection as Record<string, unknown> | undefined;
  let projection: Model3dCamera["projection"] | null = null;
  if (projectionRaw && projectionRaw.kind === "perspective" && finite(projectionRaw.fov)
    && projectionRaw.fov > 0 && projectionRaw.fov < 180 * ANGLE_UNIT) {
    projection = { kind: "perspective", fov: projectionRaw.fov };
  } else if (projectionRaw && projectionRaw.kind === "orthographic" && finite(projectionRaw.size) && projectionRaw.size > 0) {
    projection = { kind: "orthographic", size: projectionRaw.size };
  }
  if (!position || !up || !lookAt || !projection) return null;
  const meterPerModelUnit = transform.meterPerModelUnit;
  const preTranslation = vector(transform.preTranslation);
  const scale = vector(transform.scale);
  const rotation = vector(transform.rotation);
  const postTranslation = vector(transform.postTranslation);
  if (!finite(meterPerModelUnit) || meterPerModelUnit <= 0 || !preTranslation || !scale || !rotation || !postTranslation) {
    return null;
  }
  const frameWidth = scene.frameWidth;
  const frameHeight = scene.frameHeight;
  if (!finite(frameWidth) || !finite(frameHeight) || frameWidth <= 0 || frameHeight <= 0) return null;
  const viewportSize = scene.viewportSize;
  if (viewportSize !== null && (!finite(viewportSize) || viewportSize <= 0)) return null;
  let ambientLight: Model3dAmbientLight | null = null;
  if (scene.ambientLight != null) {
    const ambient = scene.ambientLight as Record<string, unknown>;
    const color = colour(ambient.color);
    if (!color || !finite(ambient.illuminance) || ambient.illuminance < 0) return null;
    ambientLight = { color, illuminance: ambient.illuminance };
  }
  if (!Array.isArray(scene.lights)) return null;
  const lights: Model3dLight[] = [];
  for (const entry of scene.lights) {
    if (!entry || typeof entry !== "object") return null;
    const light = entry as Record<string, unknown>;
    const color = colour(light.color);
    const lightPosition = vector(light.position);
    if (!color || !lightPosition) return null;
    if (light.kind === "point" && finite(light.intensity) && light.intensity >= 0 && finite(light.radius)) {
      lights.push({ kind: "point", color, intensity: light.intensity, position: lightPosition, radius: light.radius });
      continue;
    }
    const target = vector(light.lookAt);
    if (light.kind === "spot" && target && finite(light.intensity) && light.intensity >= 0 && finite(light.radius)
      && finite(light.spotAngle) && light.spotAngle > 0) {
      lights.push({
        kind: "spot", color, intensity: light.intensity, position: lightPosition, lookAt: target,
        radius: light.radius, spotAngle: light.spotAngle,
      });
      continue;
    }
    if (light.kind === "directional" && target && finite(light.illuminance) && light.illuminance >= 0) {
      lights.push({ kind: "directional", color, illuminance: light.illuminance, position: lightPosition, lookAt: target });
      continue;
    }
    return null;
  }
  return {
    camera: { position, up, lookAt, projection },
    transform: { meterPerModelUnit, preTranslation, scale, rotation, postTranslation },
    viewportSize: viewportSize as number | null,
    frameWidth,
    frameHeight,
    ambientLight,
    lights,
  };
}

/** A 4x4 matrix, column-major (as WebGL and three take it). */
export type Model3dMatrix = number[];

function identity(): Model3dMatrix {
  return [1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1];
}

function multiply(a: Model3dMatrix, b: Model3dMatrix): Model3dMatrix {
  const out = new Array<number>(16).fill(0);
  for (let column = 0; column < 4; column += 1) {
    for (let row = 0; row < 4; row += 1) {
      let sum = 0;
      for (let k = 0; k < 4; k += 1) sum += a[k * 4 + row] * b[column * 4 + k];
      out[column * 4 + row] = sum;
    }
  }
  return out;
}

function translation(x: number, y: number, z: number): Model3dMatrix {
  return [1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, x, y, z, 1];
}

function scaling(x: number, y: number, z: number): Model3dMatrix {
  return [x, 0, 0, 0, 0, y, 0, 0, 0, 0, z, 0, 0, 0, 0, 1];
}

function rotationX(radians: number): Model3dMatrix {
  const c = Math.cos(radians);
  const s = Math.sin(radians);
  return [1, 0, 0, 0, 0, c, s, 0, 0, -s, c, 0, 0, 0, 0, 1];
}

function rotationY(radians: number): Model3dMatrix {
  const c = Math.cos(radians);
  const s = Math.sin(radians);
  return [c, 0, -s, 0, 0, 1, 0, 0, s, 0, c, 0, 0, 0, 0, 1];
}

function rotationZ(radians: number): Model3dMatrix {
  const c = Math.cos(radians);
  const s = Math.sin(radians);
  return [c, s, 0, 0, -s, c, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1];
}

/**
 * `am3d:rot` as a matrix. The order and sense the three angles compose in
 * are PowerPoint's, fitted to its own renderings (`MODEL3D_ROTATION_ORDER`).
 */
export function model3dRotationMatrix(rotation: Model3dVector): Model3dMatrix {
  const toRadians = (value: number) => (value / ANGLE_UNIT) * (Math.PI / 180) * MODEL3D_ROTATION_SENSE;
  const turns: Record<"x" | "y" | "z", Model3dMatrix> = {
    x: rotationX(toRadians(rotation[0])),
    y: rotationY(toRadians(rotation[1])),
    z: rotationZ(toRadians(rotation[2])),
  };
  let matrix = identity();
  for (const axis of MODEL3D_ROTATION_ORDER) matrix = multiply(matrix, turns[axis]);
  return matrix;
}

/**
 * The order the rotation's matrices multiply in, left to right (the last
 * one turns the model first). Fitted, see `model3dRotationMatrix`.
 */
export let MODEL3D_ROTATION_ORDER: ReadonlyArray<"x" | "y" | "z"> = ["z", "y", "x"];
/** 1: counter-clockwise looking down the axis (right-handed); -1: clockwise. */
export let MODEL3D_ROTATION_SENSE = 1;

/** For the fitting harness only: try another convention. */
export function setModel3dRotationConvention(order: ReadonlyArray<"x" | "y" | "z">, sense: number): void {
  MODEL3D_ROTATION_ORDER = order;
  MODEL3D_ROTATION_SENSE = sense;
}

/**
 * The model's matrix, glTF scene units to world metres:
 * postTrans . rot . scale . preTrans . (meterPerModelUnit, in EMUs), over
 * 36,000,000 EMUs to the metre.
 */
export function model3dModelMatrix(transform: Model3dTransform): Model3dMatrix {
  const emu = MODEL3D_EMU_PER_METRE;
  const [preX, preY, preZ] = transform.preTranslation;
  const [postX, postY, postZ] = transform.postTranslation;
  const [scaleX, scaleY, scaleZ] = transform.scale;
  let matrix = scaling(1 / emu, 1 / emu, 1 / emu);
  matrix = multiply(matrix, translation(postX, postY, postZ));
  matrix = multiply(matrix, model3dRotationMatrix(transform.rotation));
  matrix = multiply(matrix, scaling(scaleX, scaleY, scaleZ));
  matrix = multiply(matrix, translation(preX, preY, preZ));
  const unit = transform.meterPerModelUnit * emu;
  matrix = multiply(matrix, scaling(unit, unit, unit));
  return matrix;
}

/** A world point in metres. */
export function model3dWorldPoint(point: Model3dVector): Model3dVector {
  const emu = MODEL3D_EMU_PER_METRE;
  return [point[0] / emu, point[1] / emu, point[2] / emu];
}

/**
 * Which part of the camera's square view the frame shows, as the bounds of
 * the near plane's window in units of its half-size (-1..1 is the whole
 * square). Object mode: the frame's own size over the viewport's, centred.
 * Window mode: the frame is the view, the square's height its height.
 */
export function model3dViewWindow(scene: Pick<Model3dScene, "viewportSize" | "frameWidth" | "frameHeight">): {
  left: number;
  right: number;
  bottom: number;
  top: number;
} {
  if (scene.viewportSize !== null) {
    const halfWidth = scene.frameWidth / scene.viewportSize;
    const halfHeight = scene.frameHeight / scene.viewportSize;
    return { left: -halfWidth, right: halfWidth, bottom: -halfHeight, top: halfHeight };
  }
  const aspect = scene.frameWidth / scene.frameHeight;
  return { left: -aspect, right: aspect, bottom: -1, top: 1 };
}
