/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A picture block that is an imported 3-D model (`source.nativeModel3d`,
 * reading 720/727): what the canvas needs to draw the model itself.
 *
 * Schema 1 carries nothing (the picture PowerPoint kept is all the page
 * has); schema 2 carries the scene (`./scene`) and the receipt of the
 * model's glTF part, which the page serves as
 * `GET /pages/{page}/presentation-native-media/{block}`.
 */

import { parseModel3dScene, type Model3dScene } from "./scene";

/** Larger glTF parts are not fetched (Office's own models are a few MB). */
export const MODEL3D_MAX_BYTES = 64 * 1024 * 1024;

export interface PresentationModel3d {
  scene: Model3dScene;
  /** SHA-256 of the glTF part, lower-case hex. */
  fingerprint: string;
  sizeBytes: number;
}

/** The model a block draws, or null for any block that is not one (or whose scene is not drawable). */
export function presentationModel3dOf(block: { source?: unknown } | null | undefined): PresentationModel3d | null {
  const source = block?.source;
  if (!source || typeof source !== "object") return null;
  const authority = (source as { nativeModel3d?: unknown }).nativeModel3d;
  if (!authority || typeof authority !== "object") return null;
  const { schemaVersion, scene: rawScene, model } = authority as {
    schemaVersion?: unknown;
    scene?: unknown;
    model?: unknown;
  };
  if (schemaVersion !== 2 || !model || typeof model !== "object") return null;
  const { fingerprint, sizeBytes, contentType } = model as {
    fingerprint?: unknown;
    sizeBytes?: unknown;
    contentType?: unknown;
  };
  if (
    typeof fingerprint !== "string" ||
    !/^[0-9a-f]{64}$/.test(fingerprint) ||
    typeof sizeBytes !== "number" ||
    !Number.isInteger(sizeBytes) ||
    sizeBytes <= 0 ||
    sizeBytes > MODEL3D_MAX_BYTES ||
    contentType !== "model/gltf-binary"
  ) {
    return null;
  }
  const scene = parseModel3dScene(rawScene);
  if (!scene || scene.camera.projection.kind !== "perspective") return null;
  return { scene, fingerprint, sizeBytes };
}

/** Whether bytes are a glTF 2.0 binary (`glTF`, version 2). */
export function isGlbBinary(bytes: Uint8Array): boolean {
  return (
    bytes.length >= 20 &&
    bytes[0] === 0x67 &&
    bytes[1] === 0x6c &&
    bytes[2] === 0x54 &&
    bytes[3] === 0x46 &&
    bytes[4] === 2 &&
    bytes[5] === 0 &&
    bytes[6] === 0 &&
    bytes[7] === 0
  );
}

/**
 * Whether the picture PowerPoint kept of the model (its Fallback / raster)
 * still shows the model as the page holds it: the model's own properties
 * (camera, transform, lights, viewport mode) are those the picture was drawn
 * with. An editor that changes them keeps the scene the picture shows as
 * `nativeModel3d.rasterScene` until a new picture is written; a move or a
 * resize changes none of them (the picture is stretched with its frame, as
 * PowerPoint stretches it). Read from the page, so a collaborator's change
 * counts as one's own.
 */
export function model3dRasterCurrent(block: { source?: unknown } | null | undefined): boolean {
  const authority = (block?.source as { nativeModel3d?: Record<string, unknown> } | undefined)?.nativeModel3d;
  if (!authority || typeof authority !== "object") return true;
  const shown = authority.rasterScene;
  if (shown === undefined) return true;
  return sceneProperties(shown) === sceneProperties(authority.scene);
}

function sceneProperties(scene: unknown): string {
  if (!scene || typeof scene !== "object") return "";
  const { camera, transform, ambientLight, lights, viewportSize } = scene as Record<string, unknown>;
  return JSON.stringify([camera, transform, ambientLight ?? null, lights ?? null, viewportSize === null]);
}

/** Device pixels over the picture's own pixels past which the picture looks soft: the model is drawn instead. */
export const MODEL3D_SHARPEN_ABOVE = 1.5;
/** ... and below which the picture is shown again (hysteresis: a zoom near the line does not flicker). */
export const MODEL3D_RASTER_BELOW = 1.25;

/**
 * Picture or drawing, for a model whose picture is current: the drawing once
 * the frame covers more than 1.5 device pixels per picture pixel, the
 * picture again below 1.25; between the two, as it was.
 */
export function model3dSharpnessMode(current: "raster" | "webgl", scale: number | null): "raster" | "webgl" {
  if (scale === null || !Number.isFinite(scale)) return current;
  if (current === "raster") return scale > MODEL3D_SHARPEN_ABOVE ? "webgl" : "raster";
  return scale < MODEL3D_RASTER_BELOW ? "raster" : "webgl";
}
