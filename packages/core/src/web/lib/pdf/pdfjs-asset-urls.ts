/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
const configuredBase = process.env.NEXT_PUBLIC_PDFJS_ASSET_BASE ?? "";
// Only the shape the routes have; anything else is not a base they serve.
const versionedBase = /^\/pdfjs\/[0-9A-Za-z][0-9A-Za-z.-]*\/$/.test(configuredBase)
  ? configuredBase
  : null;

/**
 * `GlobalWorkerOptions.workerSrc`. Under the versioned base the name ends in
 * `.js`: Cloudflare caches by extension alone and `.mjs` is not on its list.
 * It is the same module either way.
 */
export const PDF_JS_WORKER_URL = versionedBase
  ? `${versionedBase}pdf.worker.min.js`
  : "/pdf.worker.min.mjs";

/** `cMapUrl`, and where the native engine reads the same packed CMaps. */
export const PDF_JS_CMAP_BASE_URL = versionedBase ? `${versionedBase}cmaps/` : "/pdfjs-cmaps/";

/** `standardFontDataUrl`. */
export const PDF_JS_STANDARD_FONT_BASE_URL = versionedBase
  ? `${versionedBase}standard-fonts/`
  : "/pdfjs-standard-fonts/";

/** `wasmUrl`: PDF.js 5.7's decoder modules (`public/pdfjs-wasm/`). */
export const PDF_JS_WASM_BASE_URL = versionedBase ? `${versionedBase}wasm/` : "/pdfjs-wasm/";
