/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Whether this browser can create a WebGL context at all.
 *
 * The 3-D chart kinds draw through echarts-gl, which assumes a context: on a
 * machine with WebGL disabled (a locked-down browser, a headless one, a GPU
 * blocklist) it throws from its own refresh loop, outside the chart host's
 * error handling. A host asks this first and shows its "cannot draw this
 * chart" state instead. The answer is cached for the page's lifetime; the
 * probe canvas is discarded and its context released.
 */
let cached: boolean | null = null;

export function browserSupportsWebGL(): boolean {
  if (cached !== null) return cached;
  if (typeof document === "undefined") return false;
  try {
    const canvas = document.createElement("canvas");
    const context =
      (canvas.getContext("webgl2") as WebGL2RenderingContext | null) ??
      (canvas.getContext("webgl") as WebGLRenderingContext | null);
    cached = context !== null;
    context?.getExtension("WEBGL_lose_context")?.loseContext();
  } catch {
    cached = false;
  }
  return cached;
}
