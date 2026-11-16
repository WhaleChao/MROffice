/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A picture's Picture Format > Color and Corrections, drawn as Office draws
 * them, for every editor: the host-neutral `office-recolor-presets` steps in
 * one shared SVG filter per step list (defined once per document), and
 * Grayscale as the CSS term it is. Transparency is the picture's opacity.
 * Each host hands its own theme (the document's, the workbook's), against
 * which the Dark and Light tiles resolve.
 */

import {
  isOfficeRecolorPreset,
  officeAdjustFilterSteps,
  type OfficeAdjustFilterStep,
} from "./office-recolor-presets";
import type { OfficeThemeGridColors } from "./office-theme-color-grid";

/** Set Transparent Color (`a:clrChange` to alpha 0): every pixel of this
 * colour (0..255 channels) made clear, each channel within
 * OFFICE_CLEAR_COLOR_TOLERANCE of it. First in a step list, as Office applies
 * the colour change before the recolour and the corrections. */
export interface OfficeClearColorStep {
  readonly type: "clear";
  readonly color: readonly [number, number, number];
  /** How far each channel may stand off the colour (else
   * OFFICE_CLEAR_COLOR_TOLERANCE, a JPEG's). */
  readonly tolerance?: number;
}










export const OFFICE_CLEAR_COLOR_TOLERANCE = 16;
export const OFFICE_LOSSLESS_CLEAR_COLOR_TOLERANCE = 1;

/** The 256-entry discrete transfer table of one channel of a clear step:
 * 1 where a byte value lies within the tolerance of the colour's. */
export function officeClearColorTable(value: number, tolerance: number = OFFICE_CLEAR_COLOR_TOLERANCE): number[] {
  const centre = Math.max(0, Math.min(255, Math.round(value)));
  const within = Number.isFinite(tolerance) ? Math.max(0, Math.min(255, tolerance)) : OFFICE_CLEAR_COLOR_TOLERANCE;
  return Array.from({ length: 256 }, (_, index) => (Math.abs(index - centre) <= within ? 1 : 0));
}

/** The Set Transparent Color step of a `#RRGGBB` colour, or null; with the
 * picture's tolerance when it states one (`source.pictureClearTolerance`). */
export function officeClearColorStep(color: unknown, tolerance?: unknown): OfficeClearColorStep | null {
  if (typeof color !== "string" || !/^#[0-9A-Fa-f]{6}$/.test(color)) return null;
  const value = Number.parseInt(color.slice(1), 16);
  const stated = typeof tolerance === "number" && Number.isFinite(tolerance) && tolerance >= 0 && tolerance <= 255 ? tolerance : undefined;
  return { type: "clear", color: [(value >> 16) & 255, (value >> 8) & 255, value & 255], ...(stated !== undefined ? { tolerance: stated } : {}) };
}

/** A step this filter draws: the adjust steps and the clear colour. */
export type OfficePictureFilterStep = OfficeAdjustFilterStep | OfficeClearColorStep;

const SVG_NS = "http://www.w3.org/2000/svg";
const DEFS_ID = "mrnote-picture-recolor-filters";

export interface OfficePictureAdjust {
  readonly recolor?: unknown;
  readonly brightness?: number;
  readonly contrast?: number;
}

function defsHost(owner: Document): SVGSVGElement {
  const existing = owner.getElementById(DEFS_ID);
  if (existing instanceof SVGSVGElement) return existing;
  const svg = owner.createElementNS(SVG_NS, "svg");
  svg.id = DEFS_ID;
  svg.setAttribute("aria-hidden", "true");
  svg.setAttribute("width", "0");
  svg.setAttribute("height", "0");
  svg.setAttribute("style", "position:absolute;width:0;height:0;overflow:hidden;pointer-events:none");
  owner.body.appendChild(svg);
  return svg;
}

function filterId(steps: readonly OfficePictureFilterStep[]): string {
  const text = JSON.stringify(steps);
  let hash = 0;
  for (let index = 0; index < text.length; index += 1) {
    hash = (Math.imul(hash, 31) + text.charCodeAt(index)) | 0;
  }
  return `mrnote-picture-recolor-${(hash >>> 0).toString(36)}`;
}

function buildFilter(owner: Document, id: string, steps: readonly OfficePictureFilterStep[]): SVGFilterElement {
  const filter = owner.createElementNS(SVG_NS, "filter");
  filter.id = id;
  filter.setAttribute("color-interpolation-filters", "sRGB");
  for (const [key, value] of [["x", "0"], ["y", "0"], ["width", "1"], ["height", "1"]]) filter.setAttribute(key, value);
  for (const step of steps) {
    if (step.type === "clear") {
      // A pixel is the colour when each channel falls within the tolerance's
      // bins of a 256-step table; the three indicators summed less 2 are its
      // alpha, and the picture is drawn out of it.
      const key = owner.createElementNS(SVG_NS, "feComponentTransfer");
      key.setAttribute("in", "SourceGraphic");
      key.setAttribute("result", "mrnoteClearKey");
      ["feFuncR", "feFuncG", "feFuncB"].forEach((name, channel) => {
        const func = owner.createElementNS(SVG_NS, name);
        func.setAttribute("type", "discrete");
        func.setAttribute("tableValues", officeClearColorTable(step.color[channel] ?? 0, step.tolerance).join(" "));
        key.appendChild(func);
      });
      filter.appendChild(key);
      const hit = owner.createElementNS(SVG_NS, "feColorMatrix");
      hit.setAttribute("in", "mrnoteClearKey");
      hit.setAttribute("type", "matrix");
      hit.setAttribute("values", "0 0 0 0 0  0 0 0 0 0  0 0 0 0 0  1 1 1 0 -2");
      hit.setAttribute("result", "mrnoteClearHit");
      filter.appendChild(hit);
      const cleared = owner.createElementNS(SVG_NS, "feComposite");
      cleared.setAttribute("in", "SourceGraphic");
      cleared.setAttribute("in2", "mrnoteClearHit");
      cleared.setAttribute("operator", "out");
      filter.appendChild(cleared);
      continue;
    }
    if (step.type === "matrix") {
      const matrix = owner.createElementNS(SVG_NS, "feColorMatrix");
      matrix.setAttribute("type", "matrix");
      matrix.setAttribute("values", step.values.join(" "));
      filter.appendChild(matrix);
      continue;
    }
    const transfer = owner.createElementNS(SVG_NS, "feComponentTransfer");
    for (const name of ["feFuncR", "feFuncG", "feFuncB"]) {
      const func = owner.createElementNS(SVG_NS, name);
      if (step.type === "discrete") {
        func.setAttribute("type", "discrete");
        func.setAttribute("tableValues", step.table.join(" "));
      } else {
        func.setAttribute("type", "linear");
        func.setAttribute("slope", String(step.slope));
        func.setAttribute("intercept", String(step.intercept));
      }
      transfer.appendChild(func);
    }
    filter.appendChild(transfer);
  }
  return filter;
}

/** The SVG-filter term (`url(#…)`) for a step list, its filter defined once in `owner`. */
export function officeAdjustStepsFilterCss(
  steps: readonly OfficePictureFilterStep[],
  owner: Document | null = typeof document === "undefined" ? null : document,
): string {
  if (!owner || !owner.body || !steps.length) return "";
  const id = filterId(steps);
  if (!owner.getElementById(id)) defsHost(owner).appendChild(buildFilter(owner, id, steps));
  return `url(#${id})`;
}

/**
 * The CSS `filter` terms for a picture's Color and Corrections: Grayscale as
 * `grayscale(1)`, every other preset and the corrections as the shared SVG
 * filter; "" when the picture is drawn as it is.
 */
export function officePictureAdjustFilterCss(
  adjust: OfficePictureAdjust,
  theme: OfficeThemeGridColors | null | undefined,
  owner: Document | null = typeof document === "undefined" ? null : document,
): string {
  const recolor = isOfficeRecolorPreset(adjust.recolor) ? adjust.recolor : null;
  const terms: string[] = [];
  if (recolor === "grayscale") terms.push("grayscale(1)");
  const steps = officeAdjustFilterSteps(
    {
      recolor: recolor === "grayscale" ? null : recolor,
      brightness: adjust.brightness ?? 0,
      contrast: adjust.contrast ?? 0,
    },
    theme ?? undefined,
  );
  const svg = officeAdjustStepsFilterCss(steps, owner);
  if (svg) terms.push(svg);
  return terms.join(" ");
}

/**
 * Word's Washout as the CSS `filter` of a picture drawn over white: the
 * Washout correction (Brightness +70 %, Contrast -70 %: 0.3 c + 0.805 per
 * channel), the same for a header's VML picture watermark (`v:imagedata`
 * gain 19661f, black level 22938f) and a DrawingML picture's Recolor >
 * Washout (Windows Word, probe docx-look probes/washout w1/w3: black 205,
 * 25 % gray 224, 50 % gray 244, red 255/205/205). "" without a document.
 */
export function officeWashoutFilterCss(
  owner: Document | null = typeof document === "undefined" ? null : document,
): string {
  return officeAdjustStepsFilterCss(officeAdjustFilterSteps({ recolor: "washout" }), owner);
}
