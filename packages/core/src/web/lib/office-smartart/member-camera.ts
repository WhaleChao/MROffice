/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isOfficePictureEffectPreset, officePictureEffectPresetEffects } from "../office-picture-effect-presets";

export const SMARTART_GRAPHIC_SCENE_KEY = "smartArtGraphicScene3d";

type Member = Record<string, unknown>;

const record = (value: unknown): Record<string, unknown> | null =>
  value && typeof value === "object" && !Array.isArray(value) ? (value as Record<string, unknown>) : null;

/** The turn a node's own Shape Effects give it, as a `rotation3d`; null for none. */
export function smartArtNodeCameraTurn(effects: unknown): Record<string, unknown> | null {
  const stated = record(effects);
  if (!stated) return null;
  const rotation = record(stated.rotation3d);
  if (rotation && typeof rotation.camera === "string") {
    const front = rotation.camera === "orthographicFront";
    return front && !rotation.lat && !rotation.lon && !rotation.rev ? null : { ...rotation };
  }
  if (stated.rotation3d === undefined && isOfficePictureEffectPreset(stated.preset)) {
    const camera = officePictureEffectPresetEffects(stated.preset).camera;
    return camera
      ? { camera: "perspectiveFront", lat: camera.lat, lon: camera.lon, rev: camera.rev, fov: camera.fov }
      : null;
  }
  return null;
}

/** Each label wearing its node's own turn (none where the node has none). */
export function withSmartArtLabelCameras<T>(blocks: readonly T[]): T[] {
  const turns = new Map<string, Record<string, unknown> | null>();
  for (const block of blocks as readonly Member[]) {
    if (block.smartArtRole !== "node" || typeof block.groupId !== "string" || typeof block.smartArtItemIndex !== "number") continue;
    const key = `${block.groupId}#${block.smartArtItemIndex}`;
    if (!turns.has(key)) turns.set(key, smartArtNodeCameraTurn(block.shapeEffects));
  }
  let changed = false;
  const next = (blocks as readonly Member[]).map((block) => {
    if (block.smartArtRole !== "label" || typeof block.groupId !== "string" || typeof block.smartArtItemIndex !== "number") {
      return block;
    }
    const turn = turns.get(`${block.groupId}#${block.smartArtItemIndex}`) ?? null;
    const effects = record(block.shapeEffects);
    const current = effects ? record(effects.rotation3d) : null;
    if (JSON.stringify(current ?? null) === JSON.stringify(turn)) return block;
    changed = true;
    const rest = { ...(effects ?? {}) };
    delete rest.rotation3d;
    const shapeEffects = turn ? { ...rest, rotation3d: turn } : rest;
    const out: Member = { ...block, shapeEffects };
    if (!Object.keys(shapeEffects).length) delete out.shapeEffects;
    return out;
  });
  return (changed ? next : blocks.slice()) as T[];
}

/**
 * Every member of each group whose first member keeps the graphic's own
 * camera (`smartArtGraphicScene3d`) drawn in that one scene (`scene3d`).
 */
export function withSmartArtGraphicCamera<T>(blocks: readonly T[]): T[] {
  const scenes = new Map<string, unknown>();
  for (const block of blocks as readonly Member[]) {
    const scene = record(block[SMARTART_GRAPHIC_SCENE_KEY]);
    if (scene && typeof block.groupId === "string" && !scenes.has(block.groupId)) scenes.set(block.groupId, scene);
  }
  if (!scenes.size) return blocks.slice();
  return (blocks as readonly Member[]).map((block) => {
    const scene = typeof block.groupId === "string" ? scenes.get(block.groupId) : undefined;
    if (!scene || JSON.stringify(block.scene3d) === JSON.stringify(scene)) return block;
    return { ...block, scene3d: JSON.parse(JSON.stringify(scene)) };
  }) as T[];
}

/** The graphic's own camera a group keeps, or null. */
export function smartArtGraphicSceneOf(members: readonly unknown[]): Record<string, unknown> | null {
  for (const member of members) {
    const scene = record(record(member)?.[SMARTART_GRAPHIC_SCENE_KEY]);
    if (scene) return scene;
  }
  return null;
}
