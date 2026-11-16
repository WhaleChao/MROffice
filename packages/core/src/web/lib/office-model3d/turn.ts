/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PresentationModel3d } from "./block";
import type { LoadedModel3d, renderModel3dPicture, model3dVisibleExtent } from "./renderer";
import type { Model3dScene, Model3dVector } from "./scene";

/** The frame over the model's reach (least squares over model3d-rot-A/B/D; +-3 %). */
export const MODEL3D_FRAME_FIT = 1.028;
/** The picture drawn for a turned model: pixels per inch of its frame (PowerPoint's own 144-192). */
export const MODEL3D_PICTURE_PX_PER_INCH = 144;
/** And never more than this many pixels along its longer side. */
export const MODEL3D_PICTURE_MAX_PX = 2048;
const EMU_PER_INCH = 914400;
const FULL_TURN = 21_600_000;

export interface Model3dFrame {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** A pane angle in degrees as `am3d:rot` holds it: 60,000ths, 0 <= value < 360 degrees. */
export function model3dRotationOf(degrees: number): number {
  const turn = Math.round(degrees * 60000) % FULL_TURN;
  return turn < 0 ? turn + FULL_TURN : turn;
}

/** The pane's X / Y / Z (degrees, 0..359.9) of a model's rotation. */
export function model3dPaneAngles(rotation: Model3dVector): [number, number, number] {
  const shown = (value: number) => {
    const turn = ((value % FULL_TURN) + FULL_TURN) % FULL_TURN;
    return Math.round((turn / 60000) * 10) / 10;
  };
  return [shown(rotation[0]), shown(rotation[1]), shown(rotation[2])];
}

/**
 * The frame of a turned model: about the old frame's centre, each side the
 * model's reach across the view (``extent``, -1..1 the view's edge) times
 * the view's size times MODEL3D_FRAME_FIT. Null in window mode (the frame
 * is the user's there and is not refitted) or for a reach of nothing.
 */
export function model3dTurnedFrame(
  frame: Model3dFrame,
  viewportSize: number | null,
  extent: { x: number; y: number } | null,
): Model3dFrame | null {
  if (viewportSize === null || !extent || !(extent.x > 0) || !(extent.y > 0)) return null;
  const width = Math.max(1, Math.round(extent.x * viewportSize * MODEL3D_FRAME_FIT));
  const height = Math.max(1, Math.round(extent.y * viewportSize * MODEL3D_FRAME_FIT));
  const centreX = frame.left + frame.width / 2;
  const centreY = frame.top + frame.height / 2;
  return { left: Math.round(centreX - width / 2), top: Math.round(centreY - height / 2), width, height };
}

/** The picture's pixel size for a frame of ``width`` x ``height`` EMUs. */
export function model3dPictureSize(width: number, height: number): { width: number; height: number } {
  let scale = MODEL3D_PICTURE_PX_PER_INCH / EMU_PER_INCH;
  const longest = Math.max(width, height) * scale;
  if (longest > MODEL3D_PICTURE_MAX_PX) scale *= MODEL3D_PICTURE_MAX_PX / longest;
  return { width: Math.max(1, Math.round(width * scale)), height: Math.max(1, Math.round(height * scale)) };
}

export interface Model3dTurnResult {
  frame: Model3dFrame;
  scene: Model3dScene;
  /** The turned model drawn, a PNG data URL. */
  picture: string;
}

/**
 * Turn a model to ``rotation`` (60,000ths of a degree, the file's ax / ay /
 * az): its new scene, its frame fitted about its centre, and its picture
 * drawn again. Null when it cannot be done here (no WebGL, a window-mode
 * view, a model with nothing to draw): the model is then left as it was.
 */
export function turnModel3d(
  renderer: { model3dVisibleExtent: typeof model3dVisibleExtent; renderModel3dPicture: typeof renderModel3dPicture },
  loaded: LoadedModel3d,
  model: PresentationModel3d,
  frame: Model3dFrame,
  rotation: Model3dVector,
): Model3dTurnResult | null {
  const turned: Model3dScene = { ...model.scene, transform: { ...model.scene.transform, rotation } };
  const next = model3dTurnedFrame(frame, turned.viewportSize, renderer.model3dVisibleExtent(loaded, turned));
  if (!next) return null;
  const scene: Model3dScene = { ...turned, frameWidth: next.width, frameHeight: next.height };
  const size = model3dPictureSize(next.width, next.height);
  const picture = renderer.renderModel3dPicture(loaded, scene, size.width, size.height);
  return picture ? { frame: next, scene, picture } : null;
}
