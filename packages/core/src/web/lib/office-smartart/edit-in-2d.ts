/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useSyncExternalStore } from "react";

import type { SmartArtStyleId } from "./styles-model";

/** The editor a graphic lives in; part of its key, so ids never collide. */
export type SmartArtEditIn2dHost = "presentation" | "spreadsheet" | "document";

/**
 * The quick styles whose 3-D is a scene with a perspective view (a diagram
 * camera): the styles Edit in 2-D applies to (§3 of SMARTART-STYLES.md).
 */
const SCENE_STYLE_IDS: ReadonlySet<SmartArtStyleId> = new Set<SmartArtStyleId>([
  "3d5",
  "3d6",
  "3d7",
  "3d8",
  "3d9",
]);

/** A 3-D style with a perspective view: the scene styles `3d5`–`3d9`. */
export function smartArtStyleHasScene(styleId: string | null | undefined): boolean {
  return typeof styleId === "string" && SCENE_STYLE_IDS.has(styleId as SmartArtStyleId);
}

function sceneOf(value: unknown): unknown {
  if (!value || typeof value !== "object") return undefined;
  const record = value as Record<string, unknown>;
  // A worksheet member keeps its paint on `shape` (and an imported one its
  // display-only reading on `shape.nativeShape3d`).
  const shape = record.shape as Record<string, unknown> | undefined;
  const native = (shape?.nativeShape3d ?? record.nativeShape3d) as Record<string, unknown> | undefined;
  return record.scene3d ?? shape?.scene3d ?? native?.scene3d;
}

/** One member draws under a diagram camera (any but the flat front one). */
export function smartArtMemberHasScene(member: unknown): boolean {
  const scene = sceneOf(member) as { camera?: { preset?: unknown } } | undefined;
  const preset = scene?.camera?.preset;
  return typeof preset === "string" && preset !== "orthographicFront";
}

/**
 * Edit in 2-D applies to this graphic: its quick style is a scene style, or
 * (an imported graphic whose members disagree on a style id) its members are
 * drawn under a diagram camera.
 */
export function smartArtGraphicHasScene(
  styleId: string | null | undefined,
  members: readonly unknown[] = [],
): boolean {
  return smartArtStyleHasScene(styleId) || members.some(smartArtMemberHasScene);
}

/* ------------------------------------------------------------- the state */

const flatGraphics = new Set<string>();
const listeners = new Set<() => void>();
let version = 0;

/** The key of one graphic: its editor and its own group id. */
export function smartArtEditIn2dKey(host: SmartArtEditIn2dHost, groupId: string): string {
  return `${host}:${groupId}`;
}

/** The graphic is in Edit in 2-D. */
export function smartArtEditingIn2d(key: string | null | undefined): boolean {
  return typeof key === "string" && flatGraphics.has(key);
}

export function setSmartArtEditingIn2d(key: string, on: boolean): void {
  if (flatGraphics.has(key) === on) return;
  if (on) flatGraphics.add(key);
  else flatGraphics.delete(key);
  version += 1;
  for (const listener of [...listeners]) listener();
}

/** The ribbon button: one click in, one click back to 3-D. */
export function toggleSmartArtEditingIn2d(key: string): void {
  setSmartArtEditingIn2d(key, !flatGraphics.has(key));
}

export function subscribeSmartArtEditIn2d(listener: () => void): () => void {
  listeners.add(listener);
  return () => {
    listeners.delete(listener);
  };
}

const readVersion = () => version;

/** Re-renders the caller whenever any graphic enters or leaves 2-D. */
export function useSmartArtEditIn2dVersion(): number {
  return useSyncExternalStore(subscribeSmartArtEditIn2d, readVersion, readVersion);
}

/**
 * The graphic's SHAPES are locked: a scene style, not in Edit in 2-D. They
 * cannot be moved, resized or rotated one at a time (Larger / Smaller and the
 * shape's size included). The whole graphic still moves and resizes by its
 * frame, and every shape's text and formatting stay editable.
 */
export function smartArtShapesLockedIn3d(
  key: string | null | undefined,
  hasScene: boolean,
): boolean {
  return hasScene && !smartArtEditingIn2d(key);
}

/**
 * A graphic member as Edit in 2-D draws it: every piece of 3-D removed (the
 * camera, the solid's extrusion, bevels and contour, a label's text
 * extrusion), its fill, line and effects unchanged — PowerPoint's own flat
 * render of a scene style. The same object back when it has no 3-D.
 *
 * For drawing only: the page keeps the member's 3-D, so nothing built from
 * this copy may be written back.
 */
export function smartArtFlat2d<T>(member: T): T {
  if (!member || typeof member !== "object") return member;
  const record = member as Record<string, unknown>;
  const shape = record.shape as Record<string, unknown> | undefined;
  const has3d = (value: Record<string, unknown> | undefined) =>
    Boolean(value && (value.shape3d || value.scene3d || value.textShape3d || value.nativeShape3d));
  if (!has3d(record) && !has3d(shape)) return member;
  const strip = (value: Record<string, unknown>) => {
    const next = { ...value };
    delete next.shape3d;
    delete next.scene3d;
    delete next.textShape3d;
    delete next.nativeShape3d;
    return next;
  };
  const flat = strip(record);
  if (shape && has3d(shape)) flat.shape = strip(shape);
  return flat as T;
}
