/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { measureOfficeFontLineBox } from "@mrnote/core/web/components/office-ribbon/officeFontLineBox.ts";

/** Native text-origin controls: regular faces at explicitly measured sizes,
 * horizontal, default line spacing. This changes paint/caret placement only;
 * font family, geometry and native OOXML stay in the source model. */
export type NativeBaselineRequest = Readonly<{
  fontFamily: string;
  /** Source witness only; the actual registered asset alias is always measured. */
  sourceFontFamily?: string;
  fontSizePx: number;
  lineHeight: number;
  profile?: "smartart-centered-90";
}>;

type FontWitness = Readonly<{ family: string; ascent: number; descent: number; width: number; baselines: readonly (readonly [sizePx: number, baselinePx: number])[] }>;
// Direct native PDF glyph origins relative to inner top, in 96dpi pixels.
// The Calibri/Dubai controls disproved a general capped-ascent formula.
// Font metrics below identify the installed face; they do not predict layout.
// PingFang 42/22px origins also cover the unchanged 31.5/16.5pt native title/footer.
const WITNESSES: readonly FontWitness[] = [
  { family: "PingFang SC", ascent: 1060, descent: 340, width: 2791, baselines: [[22, 20], [26.67, 24], [42, 38.72], [52, 46.72]] },
  { family: "Arial", ascent: 905, descent: 212, width: 2835.44921875, baselines: [[52, 50.56]] },
  { family: "Calibri", ascent: 952, descent: 269, width: 2540.52734375, baselines: [[26.67, 25.28], [52, 49.28]] },
  { family: "Dubai", ascent: 1129, descent: 559, width: 2644, baselines: [[52, 41.28]] },
];
const caches = new WeakMap<Document, Map<string, number>>();

// Native PDF glyph origins measured from the text-frame midpoint, not its top.
// The 65/53/51/40/24pt controls embed the same verified Calibri face. CSS rounds its
// line metrics differently at these sizes; a uniform -.06em shift is incorrect.
// No interpolation to unmeasured sizes or multiline SmartArt text is intended.
// Original flow pages 37/58: native PDF midpoint offsets 18.64px at 51pt
// and mean 14.720052px at 40pt (the four native offsets span 0.249px).
// Original matrix148 at53pt: mean19.792063px from five PDF glyph origins;
// quantized native origins span0.211px. No interpolation is used.
const SMARTART_CENTERED = [[86.67, 24.445], [70.67, 19.792063], [68, 18.64], [53.33, 14.72], [32, 8.98]] as const;

function faceFor(family: string): FontWitness | undefined {
  const primary = family.split(",")[0].trim().replace(/^["']|["']$/g, "").toLowerCase();
  return WITNESSES.find(face => face.family.toLowerCase() === primary);
}

/** Bounds deliberately match the native controls. Unmeasured faces, sizes and
 * custom spacing retain the prior layout rather than inheriting a guessed fix. */
export function nativeBaselineRequest(value: NativeBaselineRequest): NativeBaselineRequest | undefined {
  const primary = value.fontFamily.split(",")[0].trim().replace(/^["']|["']$/g, "");
  if (value.sourceFontFamily !== undefined && !/^MRAI-Font-[a-f0-9]{64}(?:-[a-f0-9]{64}){0,3}$/.test(primary)) return undefined;
  const face = faceFor(value.sourceFontFamily ?? value.fontFamily);
  if (value.profile === "smartart-centered-90") {
    return face?.family === "Calibri" && Number.isFinite(value.fontSizePx) &&
      SMARTART_CENTERED.some(([size]) => Math.abs(value.fontSizePx - size) < 0.005) &&
      Number.isFinite(value.lineHeight) && Math.abs(value.lineHeight - 1.08) < 0.0001 ? value : undefined;
  }
  if (!face || !Number.isFinite(value.fontSizePx) ||
    !face.baselines.some(([size]) => Math.abs(value.fontSizePx - size) < 0.005) ||
    !Number.isFinite(value.lineHeight) || Math.abs(value.lineHeight - 1.2) > 0.0001) return undefined;
  return value;
}

/** Measure the real CSS baseline with a temporary invisible, non-interactive
 * text probe. The source model and shared font assets are not changed. */
export function measureNativeBaselineOffset(doc: Document, request: NativeBaselineRequest): number | undefined {
  const eligible = nativeBaselineRequest(request), face = faceFor(request.sourceFontFamily ?? request.fontFamily);
  const actualFamily = request.fontFamily.split(",")[0].trim().replace(/^["']|["']$/g, "");
  if (!eligible || !face || !doc.body || !doc.fonts || doc.fonts.status !== "loaded") return undefined;
  const ctx = doc.createElement("canvas").getContext("2d");
  if (!ctx) return undefined;
  ctx.font = `400 1000px ${JSON.stringify(actualFamily)}`;
  const metric = ctx.measureText("Abc12");
  // Font names alone do not prove availability: an absent face silently falls
  // back in Canvas. Match scale-independent metrics and advances before use.
  if (![metric.fontBoundingBoxAscent, metric.fontBoundingBoxDescent, metric.width].every(Number.isFinite) ||
    Math.abs(metric.fontBoundingBoxAscent - face.ascent) > 0.1 ||
    Math.abs(metric.fontBoundingBoxDescent - face.descent) > 0.1 ||
    Math.abs(metric.width - face.width) > 0.05) return undefined;
  const key = `${request.profile ?? "ordinary"}:${face.family}:${actualFamily}:${request.fontSizePx}:${request.lineHeight}`;
  const cache = caches.get(doc) ?? new Map<string, number>();
  caches.set(doc, cache);
  if (cache.has(key)) return cache.get(key);
  const centered = request.profile === "smartart-centered-90";
  const measured = measureOfficeFontLineBox(doc, {
    font: `400 ${request.fontSizePx}px ${JSON.stringify(actualFamily)}`,
    lineHeight: String(request.lineHeight),
    measureHeight: centered,
  });
  if (!measured) return undefined;
  const cssBaseline = measured.baselinePx;
  const witness = (centered ? SMARTART_CENTERED : face.baselines).find(([size]) => Math.abs(request.fontSizePx - size) < 0.005);
  if (!witness) return undefined;
  const delta = witness[1] - cssBaseline + (centered ? measured.heightPx! / 2 : 0);
  if (!Number.isFinite(delta) || Math.abs(delta) > request.fontSizePx * 0.2) return undefined;
  cache.set(key, delta);
  return delta;
}
