/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A group's 3-D Rotation (`grpSpPr/a:scene3d`) as Format Shape > 3-D Rotation
 * edits it with the group selected. Office states the camera once, on the
 * group, and turns every member by it alone (a member's own camera is kept
 * and not drawn, group-scene3d-probe V3); the page carries it on every member
 * as `scene3d`. The pane's 3-D Rotation reads and writes it in the one
 * rotation model every host's pane already edits (`OfficeRotation3dEffect`).
 */

import type { OfficeRotation3dEffect } from "./office-picture-effects";
import type { SmartArtCameraPreset, SmartArtScene3d } from "./office-smartart/styles-model";

/** The group camera as the pane's 3-D Rotation: its preset and the file's own X / Y / Z. */
export function officeGroupCameraRotation(scene: SmartArtScene3d | null | undefined): OfficeRotation3dEffect | undefined {
  if (!scene) return undefined;
  const { preset, latitude, longitude, revolution, fov } = scene.camera;
  if (preset === "orthographicFront" && !latitude && !longitude && !revolution) return undefined;
  // The file's own X / Y / Z come as one `a:rot`: all three, or none.
  const turned = [latitude, longitude, revolution].some((value) => typeof value === "number");
  return {
    camera: preset,
    ...(turned ? { lat: latitude ?? 0, lon: longitude ?? 0, rev: revolution ?? 0 } : {}),
    // Perspective (the camera's own field of view), as the pane edits it.
    ...(typeof fov === "number" ? { fov } : {}),
  };
}






export function officeGroupCameraScene(
  rotation: OfficeRotation3dEffect | "none" | null | undefined,
  previous: SmartArtScene3d | null | undefined,
): SmartArtScene3d | undefined {
  if (!rotation || rotation === "none") return undefined;
  if (rotation.camera === "orthographicFront" && !rotation.lat && !rotation.lon && !rotation.rev) return undefined;
  const turned = [rotation.lat, rotation.lon, rotation.rev].some((value) => typeof value === "number");
  return {
    camera: {
      preset: rotation.camera as SmartArtCameraPreset,
      ...(typeof previous?.camera.zoom === "number" ? { zoom: previous.camera.zoom } : {}),
      ...(typeof rotation.fov === "number" ? { fov: rotation.fov } : {}),
      ...(turned
        ? { latitude: rotation.lat ?? 0, longitude: rotation.lon ?? 0, revolution: rotation.rev ?? 0 }
        : {}),
    },
    lightRig: previous?.lightRig ?? { rig: "threePt", direction: "t" },
    ...(previous?.backdrop ? { backdrop: previous.backdrop } : {}),
  };
}

/** What the pane needs to edit a group's camera: the camera, and its writer. */
export interface OfficeGroupCameraControl {
  readonly scene3d: SmartArtScene3d | undefined;
  readonly write: (next: SmartArtScene3d | undefined) => boolean;
}
