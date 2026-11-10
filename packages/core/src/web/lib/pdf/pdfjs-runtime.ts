/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import "../polyfills/readable-stream-async-iterator";
import { pdfJsFontDocument } from "./pdfjs-font-document";

export type CompatiblePdfJsApi = typeof import("pdfjs-dist");

type MathWithSumPrecise = typeof Math & {
  sumPrecise?: (values: Iterable<number>) => number;
};

const mathWithSumPrecise = Math as MathWithSumPrecise;

/**
 * PDF.js 5.7's modern build calls the Stage-4 `Math.sumPrecise` API while
 * laying out text. Safari releases that do not expose the API otherwise fail
 * before a PDF can render or enter the native text planner. Install a bounded,
 * spec-shaped exact-partials implementation before importing PDF.js.
 */
export function installPdfJsPlatformPolyfills(): void {
  if (typeof mathWithSumPrecise.sumPrecise === "function") return;

  Object.defineProperty(Math, "sumPrecise", {
    configurable: true,
    writable: true,
    value(values: Iterable<number>): number {
      const numbers: number[] = [];
      let sawPositiveValueOrZero = false;
      let infinitySign = 0;
      let count = 0;

      for (const value of values) {
        count += 1;
        if (count > Number.MAX_SAFE_INTEGER) {
          throw new RangeError("Maximum allowed index exceeded");
        }
        if (typeof value !== "number") {
          throw new TypeError("Value is not a number");
        }
        if (Number.isNaN(value)) return Number.NaN;
        if (value === Number.POSITIVE_INFINITY) {
          if (infinitySign < 0) return Number.NaN;
          infinitySign = 1;
          continue;
        }
        if (value === Number.NEGATIVE_INFINITY) {
          if (infinitySign > 0) return Number.NaN;
          infinitySign = -1;
          continue;
        }
        if (value !== 0 || Object.is(value, 0)) {
          sawPositiveValueOrZero = true;
        }
        if (value !== 0) numbers.push(value);
      }

      if (infinitySign > 0) return Number.POSITIVE_INFINITY;
      if (infinitySign < 0) return Number.NEGATIVE_INFINITY;
      if (numbers.length === 0) return sawPositiveValueOrZero ? 0 : -0;

      // Maintain non-overlapping partials (the same exact-summation family
      // used by Python's fsummation algorithm), then round only once.
      const partials: number[] = [];
      for (const number of numbers) {
        let x = number;
        let used = 0;
        for (const partial of partials) {
          let y = partial;
          if (Math.abs(x) < Math.abs(y)) {
            const swap = x;
            x = y;
            y = swap;
          }
          const high = x + y;
          const low = y - (high - x);
          if (low !== 0) partials[used++] = low;
          x = high;
        }
        partials.length = used;
        if (x !== 0) partials.push(x);
      }

      let high = 0;
      let low = 0;
      let index = partials.length - 1;
      while (index >= 0) {
        const next = partials[index--];
        const sum = high + next;
        low = next - (sum - high);
        high = sum;
        if (low !== 0) break;
      }
      if (
        index >= 0 &&
        ((low < 0 && partials[index] < 0) || (low > 0 && partials[index] > 0))
      ) {
        const doubled = low * 2;
        const rounded = high + doubled;
        if (doubled === rounded - high) high = rounded;
      }
      return high;
    },
  });
}

/**
 * Same-origin home of PDF.js 5.7's decoder modules (`public/pdfjs-wasm/`),
 * alongside the packed CMaps and standard fonts.
 *
 * PDF.js will not look for these modules at all unless `getDocument` is given
 * a `wasmUrl`: JBIG2, CCITTFax and JPEG 2000 image XObjects are then dropped
 * from the operator list with only a `getOperatorList - ignoring XObject`
 * warning, and ICCBased colour spaces fall back to an approximation. A scanned
 * page — CCITTFax is its usual encoding — renders as blank paper under the
 * editor's own selection handles. Loaders take it through
 * `PDF_JS_DOCUMENT_ASSETS` below.
 *
 * Passing it beside `cMapUrl`/`cMapPacked`/`standardFontDataUrl` also lets
 * PDF.js turn on `useWorkerFetch`, which is what its ICC (qcms) module
 * requires before it will initialise at all.
 */
export const PDF_JS_WASM_BASE_URL = "/pdfjs-wasm/";

/**
 * The same-origin assets every `getDocument` call hands PDF.js, as one value.
 *
 * They belong together: PDF.js turns on `useWorkerFetch` only when all of them
 * are given, and without the decoder URL it drops JBIG2, CCITTFax and JPEG
 * 2000 pictures outright. Spelled out loader by loader they drifted -- every
 * loader named the CMaps and fonts, none named the decoders, and scanned pages
 * rendered blank. Spread this, then override only what a loader genuinely
 * needs to differ.
 */
export const PDF_JS_DOCUMENT_ASSETS = Object.freeze({
  cMapUrl: "/pdfjs-cmaps/",
  cMapPacked: true,
  standardFontDataUrl: "/pdfjs-standard-fonts/",
  wasmUrl: PDF_JS_WASM_BASE_URL,
} as const);

/**
 * PDF.js whose `getDocument` loads a PDF's fonts through
 * {@link pdfJsFontDocument} (unless a caller names its own document): its
 * faces join the page's font set once loaded, a frame's worth at a time,
 * instead of restyling the whole page for each one.
 */
function withPdfJsFontDocument(pdfjs: CompatiblePdfJsApi): CompatiblePdfJsApi {
  const getDocument = pdfjs.getDocument;
  if (typeof getDocument !== "function" || typeof document === "undefined") return pdfjs;
  const wrapped: CompatiblePdfJsApi["getDocument"] = (src) => {
    if (
      src && typeof src === "object" && !(src instanceof URL) && !(src instanceof ArrayBuffer)
      && !ArrayBuffer.isView(src) && !src.ownerDocument && src.disableFontFace !== true
    ) {
      return getDocument({ ...src, ownerDocument: pdfJsFontDocument(document) });
    }
    return getDocument(src);
  };
  return { ...pdfjs, getDocument: wrapped };
}

let compatiblePdfJsPromise: Promise<CompatiblePdfJsApi> | null = null;












export function pdfJsModernBuildSupported(): boolean {
  const iteratorPrototype = (globalThis as { Iterator?: { prototype?: { some?: unknown; find?: unknown } } })
    .Iterator?.prototype;
  return (
    typeof (Map.prototype as { getOrInsertComputed?: unknown }).getOrInsertComputed === "function" &&
    typeof (Promise as { withResolvers?: unknown }).withResolvers === "function" &&
    typeof (Promise as { try?: unknown }).try === "function" &&
    typeof (URL as { parse?: unknown }).parse === "function" &&
    typeof AbortSignal !== "undefined" &&
    typeof (AbortSignal as { any?: unknown }).any === "function" &&
    typeof iteratorPrototype?.some === "function" &&
    typeof iteratorPrototype?.find === "function" &&
    typeof (Uint8Array.prototype as { toBase64?: unknown }).toBase64 === "function" &&
    typeof (Uint8Array as { fromBase64?: unknown }).fromBase64 === "function"
  );
}

export function loadCompatiblePdfJs(): Promise<CompatiblePdfJsApi> {
  if (compatiblePdfJsPromise) return compatiblePdfJsPromise;
  installPdfJsPlatformPolyfills();
  const build: Promise<unknown> = pdfJsModernBuildSupported()
    ? import("pdfjs-dist")
    : import("./pdfjs-legacy-build");
  compatiblePdfJsPromise = build.then((module) => {
    const pdfjs = ((module as { default?: CompatiblePdfJsApi }).default ??
      module) as CompatiblePdfJsApi;
    pdfjs.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";
    return withPdfJsFontDocument(pdfjs);
  });
  return compatiblePdfJsPromise;
}
