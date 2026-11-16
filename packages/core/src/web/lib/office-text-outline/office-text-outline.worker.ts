/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// WordArt text outlines' glyph outlines off the page's thread
// (lib/office-text-outline/outline-client.ts): the face's bytes, the WOFF2
// decompression, the glyph decoding and the PathKit union of a glyph whose
// contours overlap. Answers path data in font units. A classic worker: it
// reads PathKit's script (pathkit-wasm, an Emscripten build with a Node
// branch no bundle could resolve) as the package ships it, with importScripts.
import { OutlineEngine, pathKitUnion, type FaceOutlines, type PathKitInit } from "./outline-engine";

type Request = { id: number; url: string; glyphs: number[] } | { id: number; union: string[] } | { release: true };
type Response = { ready: true } | { id: number; outlines?: FaceOutlines; union?: string[]; error?: string };

const scope = self as unknown as {
  postMessage(value: Response): void;
  onmessage: ((event: MessageEvent<Request>) => void) | null;
};

const engine = new OutlineEngine(
  async (url) => {
    const response = await fetch(url, { cache: "force-cache", credentials: "same-origin" });
    if (!response.ok) throw new Error(`font ${response.status}`);
    return response.arrayBuffer();
  },
  pathKitUnion(
    async () => {
      (self as unknown as { importScripts(url: string): void }).importScripts(new URL("pathkit-wasm/bin/pathkit.js", import.meta.url).href);
      const init = (self as unknown as { PathKitInit?: PathKitInit }).PathKitInit;
      if (typeof init !== "function") throw new Error("pathkit did not load");
      return init;
    },
    async () => {
      const response = await fetch(new URL("pathkit-wasm/bin/pathkit.wasm", import.meta.url));
      if (!response.ok) throw new Error(`pathkit ${response.status}`);
      return response.arrayBuffer();
    },
  ),
);

scope.postMessage({ ready: true });
scope.onmessage = (event) => {
  const data = event.data;
  if ("release" in data) {
    engine.release();
    return;
  }
  if ("union" in data) {
    void engine.unionPaths(data.union).then(
      (union) => scope.postMessage({ id: data.id, union }),
      (error: unknown) => scope.postMessage({ id: data.id, error: error instanceof Error ? error.message : String(error) }),
    );
    return;
  }
  void engine.outlines(data.url, data.glyphs).then(
    (outlines) => scope.postMessage({ id: data.id, outlines }),
    (error: unknown) => scope.postMessage({ id: data.id, error: error instanceof Error ? error.message : String(error) }),
  );
};
