/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Picture Format > Artistic Effects, shared by Docs, the deck and the
 * worksheet.
 *
 * Office semantics (Microsoft Support, "Apply an artistic effect to a
 * picture"; [MS-ODRAWXML] 2.3.3.19 CT_PictureEffect and the "Pictures"
 * section): a picture carries at most one artistic effect, chosen from
 * Office's gallery of 22 (plus None). It is stored in the picture's
 * `a:blip/a:extLst/a:ext[@uri="{BEBA8EAE-BF5A-486c-A8C5-ECC9F3942E4B}"]` as
 * `a14:imgProps/a14:imgLayer[@r:embed]/a14:imgEffect/a14:artistic*`, where the
 * `imgLayer` names the ORIGINAL picture and the blip's own `r:embed` is the
 * picture WITH the effect applied (the bitmap every other reader shows).
 * Choosing None removes the extension and puts the original back.
 *
 * MRAI draws an effect with an SVG filter (`OFFICE_ARTISTIC_EFFECT_FILTERS`);
 * the server writers render the same effect into the blip with Pillow
 * (`app/services/office_artistic_effects.py`). Five of Office's 22 have no
 * faithful drawing here (Chalk Sketch, Crisscross Etching, Light Screen, Paint
 * Brush, Paint Strokes): the gallery lists them disabled, and one read from a
 * file keeps Office's own rendered blip untouched.
 */

export const OFFICE_ARTISTIC_EFFECT_IDS = [
  "none",
  "artisticMarker",
  "artisticPencilGrayscale",
  "artisticPencilSketch",
  "artisticLineDrawing",
  "artisticChalkSketch",
  "artisticPaintStrokes",
  "artisticPaintBrush",
  "artisticGlowDiffused",
  "artisticBlur",
  "artisticLightScreen",
  "artisticWatercolorSponge",
  "artisticFilmGrain",
  "artisticMosiaicBubbles",
  "artisticGlass",
  "artisticCement",
  "artisticTexturizer",
  "artisticCrisscrossEtching",
  "artisticPastelsSmooth",
  "artisticPlasticWrap",
  "artisticCutout",
  "artisticPhotocopy",
  "artisticGlowEdges",
] as const;

/** Office's gallery order (the effect names are the `a14:` element names). */
export type OfficeArtisticEffectId = (typeof OFFICE_ARTISTIC_EFFECT_IDS)[number];

/** The effect's name as a `console-notebooks` catalogue key. */
export function officeArtisticEffectLabelKey(
  id: OfficeArtisticEffectId,
): string {
  return `office.artisticEffect.${id}`;
}

export function isOfficeArtisticEffectId(value: unknown): value is OfficeArtisticEffectId {
  return typeof value === "string" && (OFFICE_ARTISTIC_EFFECT_IDS as readonly string[]).includes(value);
}

/** The effect a picture stores, or null for none / an unknown value. */
export function normalizeOfficeArtisticEffect(value: unknown): Exclude<OfficeArtisticEffectId, "none"> | null {
  return isOfficeArtisticEffectId(value) && value !== "none" ? value : null;
}

const EDGE = '<feConvolveMatrix order="3" kernelMatrix="-1 -1 -1 -1 8 -1 -1 -1 -1" preserveAlpha="true"/>';
const GRAY = '<feColorMatrix type="saturate" values="0"/>';
const INVERT = '<feComponentTransfer><feFuncR type="table" tableValues="1 0"/><feFuncG type="table" tableValues="1 0"/><feFuncB type="table" tableValues="1 0"/></feComponentTransfer>';
const posterize = (levels: string) =>
  `<feComponentTransfer><feFuncR type="discrete" tableValues="${levels}"/><feFuncG type="discrete" tableValues="${levels}"/><feFuncB type="discrete" tableValues="${levels}"/></feComponentTransfer>`;

/**
 * The SVG filter primitives each drawable effect is drawn with (the inside of
 * one `<filter>`). An effect absent here is one MRAI cannot draw faithfully.
 */
export const OFFICE_ARTISTIC_EFFECT_FILTERS: Readonly<Partial<Record<OfficeArtisticEffectId, string>>> = {
  artisticBlur: '<feGaussianBlur stdDeviation="2.5"/>',
  artisticPencilGrayscale: `${GRAY}${EDGE}${INVERT}<feComponentTransfer><feFuncR type="gamma" exponent="2.2"/><feFuncG type="gamma" exponent="2.2"/><feFuncB type="gamma" exponent="2.2"/></feComponentTransfer>`,
  artisticPencilSketch: `${GRAY}${EDGE}${INVERT}<feComponentTransfer><feFuncR type="gamma" exponent="4"/><feFuncG type="gamma" exponent="4"/><feFuncB type="gamma" exponent="4"/></feComponentTransfer>`,
  artisticLineDrawing: `${GRAY}${EDGE}${INVERT}${posterize("0 1 1 1")}`,
  artisticPhotocopy: `${GRAY}${posterize("0 0 1")}`,
  artisticGlowEdges: `${EDGE}<feComponentTransfer><feFuncR type="linear" slope="3"/><feFuncG type="linear" slope="3"/><feFuncB type="linear" slope="3"/></feComponentTransfer>`,
  artisticCutout: posterize("0 0.33 0.66 1"),
  artisticMarker: `<feGaussianBlur stdDeviation="0.8"/>${posterize("0 0.25 0.5 0.75 1")}<feColorMatrix type="saturate" values="1.4"/>`,
  artisticFilmGrain:
    '<feTurbulence type="fractalNoise" baseFrequency="0.9" numOctaves="2" result="noise"/><feColorMatrix in="noise" type="saturate" values="0" result="grain"/><feComposite in="SourceGraphic" in2="grain" operator="arithmetic" k1="0" k2="1" k3="0.25" k4="-0.12" result="grained"/><feComposite in="grained" in2="SourceGraphic" operator="in"/>',
  artisticPastelsSmooth: '<feGaussianBlur stdDeviation="1.5"/><feColorMatrix type="saturate" values="1.3"/><feComponentTransfer><feFuncR type="linear" slope="0.85" intercept="0.15"/><feFuncG type="linear" slope="0.85" intercept="0.15"/><feFuncB type="linear" slope="0.85" intercept="0.15"/></feComponentTransfer>',
  artisticGlowDiffused: '<feGaussianBlur stdDeviation="3" result="soft"/><feBlend in="SourceGraphic" in2="soft" mode="screen"/>',
  artisticWatercolorSponge:
    '<feTurbulence type="fractalNoise" baseFrequency="0.05" numOctaves="2" result="t"/><feDisplacementMap in="SourceGraphic" in2="t" scale="6"/><feGaussianBlur stdDeviation="1"/><feColorMatrix type="saturate" values="1.5"/>',
  artisticTexturizer:
    '<feTurbulence type="fractalNoise" baseFrequency="0.6" numOctaves="2" result="t"/><feDiffuseLighting in="t" surfaceScale="2" lighting-color="#ffffff" result="light"><feDistantLight azimuth="45" elevation="55"/></feDiffuseLighting><feBlend in="SourceGraphic" in2="light" mode="multiply" result="lit"/><feComposite in="lit" in2="SourceGraphic" operator="in"/>',
  artisticCement:
    '<feTurbulence type="fractalNoise" baseFrequency="0.35" numOctaves="3" result="t"/><feDiffuseLighting in="t" surfaceScale="4" lighting-color="#ffffff" result="light"><feDistantLight azimuth="45" elevation="40"/></feDiffuseLighting><feColorMatrix in="SourceGraphic" type="saturate" values="0.35" result="dull"/><feBlend in="dull" in2="light" mode="multiply" result="lit"/><feComposite in="lit" in2="SourceGraphic" operator="in"/>',
  artisticGlass:
    '<feTurbulence type="fractalNoise" baseFrequency="0.08" numOctaves="1" result="t"/><feDisplacementMap in="SourceGraphic" in2="t" scale="10"/>',
  artisticMosiaicBubbles:
    '<feTurbulence type="turbulence" baseFrequency="0.12" numOctaves="1" result="t"/><feDisplacementMap in="SourceGraphic" in2="t" scale="14"/><feGaussianBlur stdDeviation="0.6"/>',
  // The five Office draws with brushes MRAI has no copy of, approximated the
  // way the others are (PPT-A133), the server drawing the same recipes.
  artisticChalkSketch:
    '<feColorMatrix type="saturate" values="0.5" result="dull"/><feComponentTransfer in="dull" result="dark"><feFuncR type="linear" slope="0.45"/><feFuncG type="linear" slope="0.45"/><feFuncB type="linear" slope="0.45"/></feComponentTransfer>' +
    `<feColorMatrix in="SourceGraphic" type="saturate" values="0"/>${EDGE}<feComponentTransfer result="chalk"><feFuncR type="linear" slope="3"/><feFuncG type="linear" slope="3"/><feFuncB type="linear" slope="3"/></feComponentTransfer><feBlend in="dark" in2="chalk" mode="screen"/>`,
  artisticPaintStrokes:
    '<feTurbulence type="fractalNoise" baseFrequency="0.02 0.15" numOctaves="2" result="t"/><feDisplacementMap in="SourceGraphic" in2="t" scale="8"/><feGaussianBlur stdDeviation="0.8"/><feColorMatrix type="saturate" values="1.2"/>',
  artisticPaintBrush:
    `<feTurbulence type="fractalNoise" baseFrequency="0.03" numOctaves="2" result="t"/><feDisplacementMap in="SourceGraphic" in2="t" scale="12"/>${posterize("0 0.2 0.4 0.6 0.8 1")}`,
  artisticLightScreen:
    '<feComponentTransfer result="light"><feFuncR type="linear" slope="0.7" intercept="0.3"/><feFuncG type="linear" slope="0.7" intercept="0.3"/><feFuncB type="linear" slope="0.7" intercept="0.3"/></feComponentTransfer><feTurbulence type="turbulence" baseFrequency="0 0.5" numOctaves="1" result="lines"/><feColorMatrix in="lines" type="saturate" values="0" result="grey"/><feBlend in="light" in2="grey" mode="screen" result="screened"/><feComposite in="screened" in2="SourceGraphic" operator="in"/>',
  artisticCrisscrossEtching:
    `${GRAY}${posterize("0 0.5 1")}<feComponentTransfer result="tone"><feFuncR type="linear" slope="1"/></feComponentTransfer><feTurbulence type="turbulence" baseFrequency="0.4 0.4" numOctaves="1" result="h"/><feColorMatrix in="h" type="saturate" values="0" result="hatch"/><feBlend in="tone" in2="hatch" mode="multiply" result="etched"/><feComposite in="etched" in2="SourceGraphic" operator="in"/>`,
  artisticPlasticWrap:
    '<feGaussianBlur stdDeviation="2" result="b"/><feSpecularLighting in="b" surfaceScale="5" specularConstant="1" specularExponent="20" lighting-color="#ffffff" result="spec"><feDistantLight azimuth="225" elevation="45"/></feSpecularLighting><feComposite in="spec" in2="SourceGraphic" operator="in" result="shine"/><feComposite in="SourceGraphic" in2="shine" operator="arithmetic" k1="0" k2="1" k3="0.6" k4="0"/>',
};

/** Whether MRAI can draw (and so offer) this effect. */
export function officeArtisticEffectDrawable(id: OfficeArtisticEffectId): boolean {
  return id === "none" || Boolean(OFFICE_ARTISTIC_EFFECT_FILTERS[id]);
}

/**
 * The Artistic Effects gallery's pick on one picture, as every editor applies
 * it: the picture with the effect (`artisticEffect`, removed for None), or
 * null when the pick does nothing — the effect came rendered from the file,
 * MRAI cannot draw it, or the picture already has it.
 */
export function withOfficeArtisticEffect<T extends object>(picture: T, id: OfficeArtisticEffectId): T | null {
  if (Reflect.get(picture, "artisticEffectRendered") === true || !officeArtisticEffectDrawable(id)) return null;
  const effect = normalizeOfficeArtisticEffect(id);
  if ((normalizeOfficeArtisticEffect(Reflect.get(picture, "artisticEffect")) ?? null) === effect) return null;
  const next = { ...picture };
  if (effect) Reflect.set(next, "artisticEffect", effect);
  else Reflect.deleteProperty(next, "artisticEffect");
  return next;
}

export function officeArtisticEffectFilterId(id: OfficeArtisticEffectId): string {
  return `office-artistic-${id}`;
}

let installed: Document | null = null;

/**
 * Put the filter definitions in the document once (a hidden SVG in `body`),
 * so any host's picture element can reference them by `filter: url(#…)`.
 */
export function ensureOfficeArtisticEffectFilters(doc: Document | undefined = globalThis.document): void {
  if (!doc?.body || installed === doc) return;
  if (doc.getElementById("office-artistic-effect-filters")) {
    installed = doc;
    return;
  }
  const holder = doc.createElementNS("http://www.w3.org/2000/svg", "svg");
  holder.id = "office-artistic-effect-filters";
  holder.setAttribute("aria-hidden", "true");
  holder.setAttribute("width", "0");
  holder.setAttribute("height", "0");
  holder.style.position = "absolute";
  holder.innerHTML = Object.entries(OFFICE_ARTISTIC_EFFECT_FILTERS)
    .map(
      ([id, body]) =>
        `<filter id="${officeArtisticEffectFilterId(id as OfficeArtisticEffectId)}" x="0" y="0" width="100%" height="100%" color-interpolation-filters="sRGB">${body}</filter>`,
    )
    .join("");
  doc.body.append(holder);
  installed = doc;
}

/**
 * The CSS `filter` that draws `effect` on a picture element, or "" -- none, an
 * effect MRAI cannot draw, or a picture whose source is ALREADY the rendered
 * result (an imported blip, or one the server rendered).
 */
export function officeArtisticEffectCss(effect: unknown, alreadyRendered = false): string {
  const id = normalizeOfficeArtisticEffect(effect);
  if (!id || alreadyRendered || !OFFICE_ARTISTIC_EFFECT_FILTERS[id]) return "";
  ensureOfficeArtisticEffectFilters();
  return `url(#${officeArtisticEffectFilterId(id)})`;
}
