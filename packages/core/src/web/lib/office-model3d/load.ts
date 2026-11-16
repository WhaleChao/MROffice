/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A page's 3-D model, loaded once: its glTF part fetched from the page
 * (`GET /pages/{page}/presentation-native-media/{block}`), checked against
 * its receipt (size, glTF header, SHA-256) and read by the renderer, which
 * is itself loaded with `import()` only here. Shared by the canvas's frame
 * (PresentationModel3dFrame) and the 3D Model Rotation command (./turn).
 */

import { getApiHttpBaseUrl } from "../env";

import { isGlbBinary, type PresentationModel3d } from "./block";
import type { LoadedModel3d } from "./renderer";

export type Model3dRendererModule = typeof import("./renderer");

let rendererModule: Promise<Model3dRendererModule> | null = null;

/** The renderer (three and its glTF loader), loaded on first use. */
export function loadModel3dRenderer(): Promise<Model3dRendererModule> {
  // By its full name: the document verification host aliases exactly this
  // specifier to renderer.verification.ts (no three.js in its bundle).
  rendererModule ??= import("./renderer");
  return rendererModule;
}

/** One load per glTF part per page: a deck repeats a model on many slides. */
const models = new Map<string, Promise<LoadedModel3d>>();
const MAX_MODELS = 24;

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", bytes as BufferSource);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

/** The model of ``blockId`` on ``pageId``, as its receipt names it. */
export function loadPageModel3d(pageId: string, blockId: string, model: PresentationModel3d): Promise<LoadedModel3d> {
  const key = `${pageId}:${model.fingerprint}`;
  const known = models.get(key);
  if (known) return known;
  const loading = (async () => {
    const response = await fetch(
      `${getApiHttpBaseUrl()}/api/v1/pages/${encodeURIComponent(pageId)}/presentation-native-media/${encodeURIComponent(blockId)}`,
      { credentials: "include" },
    );
    if (!response.ok) throw new Error(`model ${response.status}`);
    const bytes = new Uint8Array(await response.arrayBuffer());
    // Exactly the part the page names: its size, a glTF binary, its digest.
    if (bytes.length !== model.sizeBytes || !isGlbBinary(bytes) || (await sha256Hex(bytes)) !== model.fingerprint) {
      throw new Error("model bytes");
    }
    const { loadModel3d } = await loadModel3dRenderer();
    return loadModel3d(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
  })();
  loading.catch(() => {
    // A failed load is tried again by the next viewer of the model.
    if (models.get(key) === loading) models.delete(key);
  });
  if (models.size >= MAX_MODELS) {
    const oldest = models.keys().next().value;
    if (oldest !== undefined) {
      const evicted = models.get(oldest);
      models.delete(oldest);
      void evicted?.then((loaded) => loaded.dispose(), () => undefined);
    }
  }
  models.set(key, loading);
  return loading;
}
