/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draws a 3-D model (`./scene`) as PowerPoint does: its glTF binary, placed,
 * lit and seen through its camera, into a canvas of the frame's size.
 *
 * Heavy (three and its glTF loader): only ever loaded with `import()` when a
 * model is on screen or its picture is drawn again, never in a first-screen
 * bundle.
 *
 * One WebGL context draws every model of the page (a deck of a dozen models
 * would otherwise exhaust the browser's contexts); each drawing is copied
 * out into the caller's own 2-D canvas at once.
 */

import {
  AmbientLight,
  Box3,
  DirectionalLight,
  Group,
  Matrix4,
  Mesh,
  MeshStandardMaterial,
  NoToneMapping,
  PCFSoftShadowMap,
  Object3D,
  PMREMGenerator,
  PerspectiveCamera,
  PointLight,
  SRGBColorSpace,
  Scene,
  Sphere,
  SpotLight,
  Vector3,
  WebGLRenderer,
  type Material,
  type Texture,
} from "three";
import { GLTFLoader } from "three/examples/jsm/loaders/GLTFLoader.js";
import { RoomEnvironment } from "three/examples/jsm/environments/RoomEnvironment.js";

import {
  model3dModelMatrix,
  model3dViewWindow,
  model3dWorldPoint,
  type Model3dScene,
} from "./scene";
import { SpecularGlossinessPlugin } from "./specGloss";

/** A model's glTF scene, loaded once and drawn as often as asked. */
export interface LoadedModel3d {
  root: Object3D;
  dispose(): void;
}




















export const MODEL3D_LIGHTING = {
  /** Every light's strength is multiplied by this. */
  exposure: 1.2,
  /** The ambient light's strength, over its illuminance. */
  ambient: 4.6416,
  /** The surroundings a specular-glossiness surface reflects (image-based light): strength. */
  environment: 0.3,
  /** three tone mapping: 0 none. */
  toneMapping: 0,
  toneMappingExposure: 1,
  /** The point and spot lights cast shadows (the model on itself): map size, 0 for none. */
  shadowMapSize: 1024,
  /** The shadows' softness (three's PCF radius, map texels). */
  shadowRadius: 3,
  shadowBias: -0.0005,
  /** Over the model's bounding radius: a metre-wide model and a 2 cm one shade alike. */
  shadowNormalBias: 0.01,
  /** The surroundings' strength on a metallic-roughness material. */
  environmentMetalRough: 0.6,
};

let sharedEnvironment: Texture | null = null;

/** The surroundings a model reflects: a neutral studio room, prefiltered once. */
function environment(gl: WebGLRenderer): Texture {
  if (sharedEnvironment) return sharedEnvironment;
  const generator = new PMREMGenerator(gl);
  sharedEnvironment = generator.fromScene(new RoomEnvironment(), 0.04).texture;
  generator.dispose();
  return sharedEnvironment;
}

/** Read a `.glb` (glTF 2.0 binary) into a scene. Rejects what it cannot read. */
export async function loadModel3d(bytes: ArrayBuffer): Promise<LoadedModel3d> {
  const loader = new GLTFLoader();
  loader.register((parser) => new SpecularGlossinessPlugin(parser));
  const gltf = await loader.parseAsync(bytes, "");
  const root = gltf.scene;
  return {
    root,
    dispose() {
      root.traverse((object) => {
        const mesh = object as Mesh;
        if (!mesh.isMesh) return;
        mesh.geometry?.dispose();
        const materials: Material[] = Array.isArray(mesh.material) ? mesh.material : [mesh.material];
        for (const material of materials) {
          for (const [key, value] of Object.entries(material)) {
            // The surroundings are the page's, shared by every model.
            if (key === "envMap") continue;
            if (value && typeof value === "object" && (value as Texture).isTexture) (value as Texture).dispose();
          }
          const own = (material.userData?.specGloss as { texture?: Texture | null } | undefined)?.texture;
          own?.dispose();
          material.dispose();
        }
      });
    },
  };
}

let sharedRenderer: WebGLRenderer | null = null;
let sharedRendererFailed = false;

/** The page's one WebGL renderer, or null where WebGL is not available. */
function renderer(): WebGLRenderer | null {
  if (sharedRenderer) return sharedRenderer;
  if (sharedRendererFailed || typeof document === "undefined") return null;
  try {
    const canvas = document.createElement("canvas");
    const next = new WebGLRenderer({ canvas, alpha: true, antialias: true, premultipliedAlpha: false, preserveDrawingBuffer: true });
    next.outputColorSpace = SRGBColorSpace;
    next.toneMapping = NoToneMapping;
    next.setClearColor(0x000000, 0);
    next.domElement.addEventListener("webglcontextlost", () => {
      // The next drawing starts a new context; what this one held (the
      // prefiltered surroundings) is gone with it.
      sharedRenderer = null;
      sharedEnvironment = null;
    });
    sharedRenderer = next;
    return next;
  } catch {
    sharedRendererFailed = true;
    return null;
  }
}

/** Whether this browser can draw a model at all (WebGL). */
export function model3dRenderingAvailable(): boolean {
  return renderer() !== null;
}

function lightColour(rgb: readonly [number, number, number]): [number, number, number, number] {
  // three takes a light's colour and its strength apart; the file's colour
  // is linear (scRGB), its strength the lux / candela it states.
  return [rgb[0], rgb[1], rgb[2], 1];
}

/**
 * Draw ``model`` under ``scene`` into ``target`` (resized to ``width`` x
 * ``height`` device pixels). Returns false when it cannot be drawn here (no
 * WebGL, an orthographic camera, an empty model): the caller keeps the
 * picture PowerPoint kept.
 */
export function renderModel3d(
  model: LoadedModel3d,
  scene: Model3dScene,
  target: HTMLCanvasElement,
  width: number,
  height: number,
): boolean {
  const gl = renderer();
  if (!gl || width < 1 || height < 1) return false;
  const { camera: cameraState } = scene;
  if (cameraState.projection.kind !== "perspective") return false;

  const world = new Scene();
  const placed = new Group();
  placed.matrixAutoUpdate = false;
  placed.matrix.fromArray(model3dModelMatrix(scene.transform));
  placed.add(model.root);
  world.add(placed);
  world.updateMatrixWorld(true);

  const bounds = new Box3().setFromObject(placed);
  if (bounds.isEmpty()) {
    placed.remove(model.root);
    return false;
  }
  const sphere = bounds.getBoundingSphere(new Sphere());

  const camera = new PerspectiveCamera();
  const eye = new Vector3(...model3dWorldPoint(cameraState.position));
  const at = new Vector3(...model3dWorldPoint(cameraState.lookAt));
  camera.position.copy(eye);
  camera.up.set(...cameraState.up).normalize();
  camera.lookAt(at);
  camera.updateMatrixWorld(true);
  const distance = eye.distanceTo(sphere.center);
  const near = Math.max(distance - sphere.radius * 1.5, distance * 0.001, 1e-4);
  const far = distance + sphere.radius * 1.5;
  const half = Math.tan(((cameraState.projection.fov / 60000) * Math.PI) / 360) * near;
  const window = model3dViewWindow(scene);
  camera.projectionMatrix.makePerspective(
    window.left * half,
    window.right * half,
    window.top * half,
    window.bottom * half,
    near,
    far,
  );
  camera.projectionMatrixInverse.copy(camera.projectionMatrix).invert();

  const { exposure, ambient } = MODEL3D_LIGHTING;
  if (MODEL3D_LIGHTING.environment > 0 || MODEL3D_LIGHTING.environmentMetalRough > 0) {
    const surroundings = environment(gl);
    model.root.traverse((object) => {
      const mesh = object as Mesh;
      if (!mesh.isMesh) return;
      for (const material of Array.isArray(mesh.material) ? mesh.material : [mesh.material]) {
        const standard = material as MeshStandardMaterial;
        if (!standard.isMeshStandardMaterial) continue;
        standard.envMap = surroundings;
        standard.envMapIntensity = standard.userData.specGloss
          ? MODEL3D_LIGHTING.environment
          : MODEL3D_LIGHTING.environmentMetalRough;
      }
    });
  }
  gl.toneMapping = MODEL3D_LIGHTING.toneMapping as typeof gl.toneMapping;
  gl.toneMappingExposure = MODEL3D_LIGHTING.toneMappingExposure;
  if (scene.ambientLight) {
    const [r, g, b] = lightColour(scene.ambientLight.color);
    const light = new AmbientLight(0xffffff, scene.ambientLight.illuminance * exposure * ambient);
    light.color.setRGB(r, g, b);
    world.add(light);
  }
  const shadows = MODEL3D_LIGHTING.shadowMapSize > 0;
  gl.shadowMap.enabled = shadows;
  gl.shadowMap.type = PCFSoftShadowMap;
  if (shadows) {
    model.root.traverse((object) => {
      const mesh = object as Mesh;
      if (mesh.isMesh) {
        mesh.castShadow = true;
        mesh.receiveShadow = true;
      }
    });
  }
  const castShadow = (light: PointLight | SpotLight | DirectionalLight) => {
    if (!shadows) return;
    light.castShadow = true;
    light.shadow.mapSize.set(MODEL3D_LIGHTING.shadowMapSize, MODEL3D_LIGHTING.shadowMapSize);
    light.shadow.radius = MODEL3D_LIGHTING.shadowRadius;
    light.shadow.bias = MODEL3D_LIGHTING.shadowBias;
    // In the model's own size: a metre-wide model and a 2 cm one shade alike.
    light.shadow.normalBias = MODEL3D_LIGHTING.shadowNormalBias * sphere.radius;
    const reach = light.position.distanceTo(sphere.center);
    const shadowCamera = light.shadow.camera as PerspectiveCamera;
    shadowCamera.near = Math.max(reach - sphere.radius * 1.2, 0.01);
    shadowCamera.far = reach + sphere.radius * 1.2;
    if ((light as DirectionalLight).isDirectionalLight) {
      const ortho = light.shadow.camera as unknown as { left: number; right: number; top: number; bottom: number };
      ortho.left = -sphere.radius; ortho.right = sphere.radius; ortho.top = sphere.radius; ortho.bottom = -sphere.radius;
    }
    light.shadow.camera.updateProjectionMatrix();
  };
  for (const entry of scene.lights) {
    const [r, g, b] = lightColour(entry.color);
    if (entry.kind === "point") {
      const light = new PointLight(0xffffff, entry.intensity * exposure, 0, 2);
      light.color.setRGB(r, g, b);
      light.position.set(...model3dWorldPoint(entry.position));
      castShadow(light);
      world.add(light);
    } else if (entry.kind === "spot") {
      const light = new SpotLight(0xffffff, entry.intensity * exposure, 0, ((entry.spotAngle / 60000) * Math.PI) / 360, 0, 2);
      light.color.setRGB(r, g, b);
      light.position.set(...model3dWorldPoint(entry.position));
      light.target.position.set(...model3dWorldPoint(entry.lookAt));
      world.add(light, light.target);
    } else {
      const light = new DirectionalLight(0xffffff, entry.illuminance * exposure);
      light.color.setRGB(r, g, b);
      light.position.set(...model3dWorldPoint(entry.position));
      light.target.position.set(...model3dWorldPoint(entry.lookAt));
      world.add(light, light.target);
    }
  }
  world.updateMatrixWorld(true);

  try {
    gl.setPixelRatio(1);
    gl.setSize(width, height, false);
    gl.clear();
    gl.render(world, camera);
    target.width = width;
    target.height = height;
    const context = target.getContext("2d");
    if (!context) return false;
    context.clearRect(0, 0, width, height);
    context.drawImage(gl.domElement, 0, 0, width, height);
    return true;
  } catch {
    return false;
  } finally {
    placed.remove(model.root);
  }
}

/** For tests: the matrix the model is placed by, as three holds it. */
export function model3dPlacement(scene: Model3dScene): Matrix4 {
  return new Matrix4().fromArray(model3dModelMatrix(scene.transform));
}

/**
 * How far the model reaches across the camera's square view as ``scene``
 * places it: the largest |x| and |y| of its vertices projected (-1..1 is the
 * square's edge), about the look-at point. PowerPoint fits a model's frame
 * to this, symmetric about the view's centre (`./turn`).
 */
export function model3dVisibleExtent(model: LoadedModel3d, scene: Model3dScene): { x: number; y: number } | null {
  const { camera: cameraState } = scene;
  if (cameraState.projection.kind !== "perspective") return null;
  const placed = new Group();
  placed.matrixAutoUpdate = false;
  placed.matrix.fromArray(model3dModelMatrix(scene.transform));
  placed.add(model.root);
  placed.updateMatrixWorld(true);
  const eye = new Vector3(...model3dWorldPoint(cameraState.position));
  const at = new Vector3(...model3dWorldPoint(cameraState.lookAt));
  const forward = at.clone().sub(eye).normalize();
  const right = forward.clone().cross(new Vector3(...cameraState.up).normalize()).normalize();
  const up = right.clone().cross(forward);
  const tangent = Math.tan(((cameraState.projection.fov / 60000) * Math.PI) / 360);
  let x = 0;
  let y = 0;
  const point = new Vector3();
  try {
    model.root.traverse((object) => {
      const mesh = object as Mesh;
      const position = mesh.isMesh ? mesh.geometry?.getAttribute("position") : null;
      if (!position) return;
      for (let index = 0; index < position.count; index += 1) {
        point.fromBufferAttribute(position, index).applyMatrix4(mesh.matrixWorld).sub(eye);
        const depth = point.dot(forward);
        if (depth <= 0) continue;
        x = Math.max(x, Math.abs(point.dot(right) / depth / tangent));
        y = Math.max(y, Math.abs(point.dot(up) / depth / tangent));
      }
    });
  } finally {
    placed.remove(model.root);
  }
  return x > 0 && y > 0 ? { x, y } : null;
}

/** ``model`` under ``scene`` as a PNG (a data URL) of ``width`` x ``height`` pixels, or null where it cannot be drawn. */
export function renderModel3dPicture(model: LoadedModel3d, scene: Model3dScene, width: number, height: number): string | null {
  if (typeof document === "undefined") return null;
  const canvas = document.createElement("canvas");
  if (!renderModel3d(model, scene, canvas, width, height)) return null;
  try {
    return canvas.toDataURL("image/png");
  } catch {
    return null;
  }
}
