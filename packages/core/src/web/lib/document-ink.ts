/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw-tab ink for Docs: the pure model.
 *
 * Word's Draw tab ("Draw and write with ink in Office"; Word for Mac: "Draw
 * with Trackpad") has a Draw toggle, an Eraser with three kinds (Stroke
 * Eraser, Small Eraser, Medium Eraser), a gallery of pens the user extends
 * with Add Pen (Pen, Highlighter, Pencil) and trims with Delete, and for each
 * pen a menu of six thicknesses, sixteen colours, eight effects and More
 * Colors (MRAI adds its own effects and custom ones, `lib/office-ink-effects`). A stroke here is an ordinary `documentShape` -- its own open
 * outline in the shared custom-geometry model, a line and no fill, floating
 * In Front Of Text -- with {@link DocumentShapeAttrs.inkTool} naming the tool
 * that drew it and {@link DocumentShapeAttrs.inkEffect} its effect. The DOCX
 * writers therefore need nothing new: the stroke leaves as the freeform
 * `wps:wsp` every host already writes (`a:custGeom`), which Word opens as a
 * drawn line; an effect leaves as its solid stand-in colour. A floating
 * stroke also leaves as Word's own ink (an InkML part, the API's
 * `document_docx_ink`), whose brush states the Office effect when there is one.
 *
 * The worksheet inks the same way (`lib/spreadsheet/shapes.ts`,
 * `createSpreadsheetInkShapeObject`); this module is its page-space twin.
 */

import { normalizeInkPenFeel, withInkPenFeelPatch, type InkPenFeel, type InkPenPatch } from "./office-ink-pen-dynamics";
import { inkLandedSampleIndexes, inkPressuresVary, inkSteadyPressureWidthFactor } from "./office-ink-smooth";
import type { FloatAnchor } from "./office-float-anchor";
import {
  DOCUMENT_INK_EFFECTS,
  DOCUMENT_INK_TOOLS,
  INK_CUT_ENDS,
  DOCUMENT_SHAPE_DEFAULT_RELATIVE_HEIGHT,
  DOCUMENT_SHAPE_EMU_PER_PX,
  DOCUMENT_SHAPE_MIN_EDGE_EMU,
  normalizeDocumentInkEffect,
  normalizeDocumentInkTool,
  normalizeDocumentShapeAttrs,
  normalizeInkCutEnds,
  normalizeInkPressure,
  normalizeInkAudio,
  normalizeInkTape,
  inkAudioPlace,
  INK_TAPE_STATES,
  MAX_INK_PRESSURE_CHARS,
  type DocumentInkEffect,
  type DocumentInkTool,
  type InkCutEnds,
  type InkTape,
  type DocumentShapeAttrs,
  type DocumentShapeAttrsInput,
} from "./document-shapes";
import { inkOpacityForTool } from "./presentation-annotation";
import {
  MRAI_INK_EFFECT_PAINTS,
  OFFICE_INK_EFFECT_PAINTS,
  inkEffectPaint,
  type InkEffectPaint,
} from "./office-ink-effects";
import {
  PRESENTATION_CUSTOM_PATH_GRID,
  PRESENTATION_CUSTOM_PATH_AUTHORING_MAX_COMMANDS,
  normalizePresentationCustomPath,
  type PresentationCustomPathCommand,
} from "./presentation-freeform";

export { DOCUMENT_INK_EFFECTS, DOCUMENT_INK_TOOLS, INK_CUT_ENDS, INK_TAPE_STATES, MAX_INK_PRESSURE_CHARS, inkAudioPlace, normalizeInkAudio, normalizeInkCutEnds, normalizeInkPressure, normalizeInkTape };
export type { DocumentInkEffect, DocumentInkTool, InkCutEnds, InkTape };

export interface DocumentInkPoint {
  x: number;
  y: number;
  /** The pen's pressure at this sample (0..1), when the pen reported one. */
  p?: number;
}

/** The stored pressures as fractions, or null when there are none / the count is off. */
export function documentInkPressures(value: unknown, count: number): number[] | null {
  const normalized = normalizeInkPressure(value);
  if (!normalized) return null;
  const tokens = normalized.split(" ");
  if (tokens.length !== count) return null;
  return tokens.map((token) => Number(token) / 100);
}

/** The samples' pressures as the shape stores them, or undefined when one width draws them. */
export function documentInkPressureOf(samples: readonly DocumentInkPoint[]): string | undefined {
  if (!inkPressuresVary(samples.map((sample) => sample.p))) return undefined;
  return samples.map((sample) => String(Math.round(Math.min(1, Math.max(0, sample.p ?? 0.5)) * 100))).join(" ");
}

/**
 * One pen of the Draw tab's gallery. Its identity is the user's (Add Pen
 * makes one, Delete removes one), as are its colour, thickness and effect,
 * chosen from the pen's own menu as in Word and kept in this browser.
 */
export interface DocumentInkPreset extends InkPenFeel {
  readonly id: string;
  readonly tool: DocumentInkTool;
  /** `#rrggbb`; with an effect, the effect's solid stand-in. */
  readonly color: string;
  /** Word states a pen's thickness in points; the outline carries it as is. */
  readonly widthPt: number;
  /** One of Word's eight effects, an MRAI preset or a custom effect (pens only). */
  readonly effect?: DocumentInkEffect;
  /**
   * Pressure off: the pen draws one width however hard it is pressed, as a
   * ballpoint does (a note-taking app's per-pen Pressure switch). Never a
   * highlighter's (it has no pressure).
   */
  readonly steady?: true;
  /**
   * Draw ▸ Tape: the pen lays strips of tape (`inkTape`), in the state a new
   * strip lands in. Only on the Tape tool's own pens, never a pen of the menu.
   */
  readonly tape?: InkTape;
}

/**
 * Word's six thicknesses (0.25 mm to 6 mm, the six dots of a pen's menu), in
 * the points the outline states: 1 mm is 72/25.4 pt, rounded to a tenth.
 */
export const DOCUMENT_INK_THICKNESSES: ReadonlyArray<{
  readonly mm: number;
  readonly pt: number;
}> = [
  { mm: 0.25, pt: 0.7 },
  { mm: 0.5, pt: 1.4 },
  { mm: 1, pt: 2.8 },
  { mm: 2, pt: 5.7 },
  { mm: 3.5, pt: 9.9 },
  { mm: 6, pt: 17 },
];

/** Word's sixteen ink colours (a pen's menu, four rows of four). */
export const DOCUMENT_INK_COLORS: readonly string[] = [
  "#f2b233",
  "#ea7325",
  "#e5397a",
  "#e12b2b",
  "#5d3a9b",
  "#9b2d8e",
  "#b8256e",
  "#2b57a5",
  "#3b8ed8",
  "#55c1f0",
  "#2f8f3c",
  "#7cc142",
  "#000000",
  "#404040",
  "#8a8a8a",
  "#ffffff",
];

/** A pen effect as it paints (`lib/office-ink-effects`): its layers, the
 * gradient its base wears, and the solid colour that stands in for it
 * wherever only one colour can be stated (the outline attribute, every file
 * writer). */
export type DocumentInkEffectPaint = InkEffectPaint;

/** Word's eight, as the pen menu's first row shows them. */
export const DOCUMENT_INK_EFFECT_PAINTS: readonly InkEffectPaint[] = OFFICE_INK_EFFECT_PAINTS;
/** MRAI's own presets, after Word's eight. */
export const DOCUMENT_INK_MRAI_EFFECT_PAINTS: readonly InkEffectPaint[] = MRAI_INK_EFFECT_PAINTS;

export function documentInkEffectPaint(effect: DocumentInkEffect): DocumentInkEffectPaint {
  return inkEffectPaint(effect);
}

/** The pen Add Pen makes for a tool: Word's defaults for a new pen. */
export function documentInkPenDefaults(tool: DocumentInkTool): Omit<DocumentInkPreset, "id"> {
  switch (tool) {
    case "highlighter":
      return { tool, color: "#ffff00", widthPt: 17 };
    case "pencil":
      return { tool, color: "#8a8a8a", widthPt: 2.8 };
    default:
      return { tool, color: "#000000", widthPt: 1.4 };
  }
}

/**
 * The gallery as Word first shows it: a black and a red pen, a pencil, a
 * yellow highlighter, a galaxy-effect pen and a heavier black pen.
 */
export const DOCUMENT_INK_PRESETS: readonly DocumentInkPreset[] = [
  { id: "pen-black", tool: "pen", color: "#000000", widthPt: 1.4 },
  { id: "pen-red", tool: "pen", color: "#e12b2b", widthPt: 1.4 },
  { id: "pencil-gray", tool: "pencil", color: "#8a8a8a", widthPt: 2.8 },
  { id: "highlighter-yellow", tool: "highlighter", color: "#ffff00", widthPt: 17 },
  { id: "pen-galaxy", tool: "pen", color: "#3b2a8a", widthPt: 2.8, effect: "galaxy" },
  { id: "pen-black-wide", tool: "pen", color: "#000000", widthPt: 2.8 },
];

/** The user's gallery of pens. */
export interface DocumentInkPenProfile {
  readonly version: 2;
  readonly pens: readonly DocumentInkPreset[];
}

export const DOCUMENT_INK_PEN_PROFILE_STORAGE_KEY = "mrai:document:ink-pens:v2";
export const DOCUMENT_INK_MAX_PENS = 24;

const INK_COLOR = /^#[0-9a-f]{6}$/;
const INK_PEN_ID = /^[a-z][a-z0-9-]{0,47}$/;
const INK_THICKNESS_PTS = new Set(DOCUMENT_INK_THICKNESSES.map(({ pt }) => pt));

function defaultDocumentInkPenProfile(): DocumentInkPenProfile {
  return { version: 2, pens: DOCUMENT_INK_PRESETS.map((pen) => ({ ...pen })) };
}

function normalizeDocumentInkPen(value: unknown): DocumentInkPreset | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const pen = value as Record<string, unknown>;
  const tool = normalizeDocumentInkTool(pen.tool);
  const color = typeof pen.color === "string" ? pen.color.toLowerCase() : "";
  if (typeof pen.id !== "string" || !INK_PEN_ID.test(pen.id) || !tool) return null;
  if (!INK_COLOR.test(color)) return null;
  const widthPt = typeof pen.widthPt === "number" && INK_THICKNESS_PTS.has(pen.widthPt)
    ? pen.widthPt
    : documentInkPenDefaults(tool).widthPt;
  const effect = tool === "pen" ? normalizeDocumentInkEffect(pen.effect) : undefined;
  return {
    id: pen.id,
    tool,
    color,
    widthPt,
    ...(effect ? { effect } : {}),
    ...(tool !== "highlighter" && pen.steady === true ? { steady: true as const } : {}),
    ...normalizeInkPenFeel(pen, tool),
  };
}

/**
 * A stored profile is a list of pens, each a known tool with a `#rrggbb`
 * colour, one of Word's thicknesses and (a pen) at most one of Word's
 * effects; a malformed pen is dropped, an empty or unreadable list falls
 * closed to Word's gallery, so a stale or forged value can never arm a tool
 * the tab does not offer.
 */
export function normalizeDocumentInkPenProfile(value: unknown): DocumentInkPenProfile {
  const fallback = defaultDocumentInkPenProfile();
  if (!value || typeof value !== "object" || Array.isArray(value)) return fallback;
  const record = value as Record<string, unknown>;
  if (record.version !== 2 || !Array.isArray(record.pens)) return fallback;
  const seen = new Set<string>();
  const pens: DocumentInkPreset[] = [];
  for (const entry of record.pens) {
    const pen = normalizeDocumentInkPen(entry);
    if (!pen || seen.has(pen.id)) continue;
    seen.add(pen.id);
    pens.push(pen);
    if (pens.length >= DOCUMENT_INK_MAX_PENS) break;
  }
  return pens.length ? { version: 2, pens } : fallback;
}

export function documentInkPenProfileFromStorage(
  serialized: string | null,
): DocumentInkPenProfile {
  if (!serialized) return defaultDocumentInkPenProfile();
  try {
    return normalizeDocumentInkPenProfile(JSON.parse(serialized));
  } catch {
    return defaultDocumentInkPenProfile();
  }
}

export function serializeDocumentInkPenProfile(profile: DocumentInkPenProfile): string {
  return JSON.stringify(normalizeDocumentInkPenProfile(profile));
}

/** The profile with one pen's colour, thickness or effect changed
 * (`effect: null` removes the effect; a colour pick does too, as in Word). */
export function withDocumentInkPen(
  profile: DocumentInkPenProfile,
  id: string,
  patch: InkPenPatch,
): DocumentInkPenProfile {
  return normalizeDocumentInkPenProfile({
    version: 2,
    pens: profile.pens.map((pen) => {
      if (pen.id !== id) return pen;
      const next: Record<string, unknown> = { ...pen };
      if (patch.widthPt !== undefined) next.widthPt = patch.widthPt;
      if (patch.color !== undefined) {
        next.color = patch.color;
        delete next.effect;
      }
      if (patch.effect === null) delete next.effect;
      else if (patch.effect !== undefined) {
        next.effect = patch.effect;
        next.color = documentInkEffectPaint(patch.effect).solid;
      }
      if (patch.steady === true) next.steady = true;
      else if (patch.steady === false) delete next.steady;
      return withInkPenFeelPatch(next, patch.feel);
    }),
  });
}

/** Add Pen: the profile with a new pen of `tool` at the end, and its id. */
export function withDocumentInkPenAdded(
  profile: DocumentInkPenProfile,
  tool: DocumentInkTool,
  seed = Date.now(),
): { profile: DocumentInkPenProfile; id: string } {
  const taken = new Set(profile.pens.map((pen) => pen.id));
  let id = `${tool}-${seed.toString(36)}`;
  for (let n = 1; taken.has(id); n += 1) id = `${tool}-${seed.toString(36)}-${n}`;
  const next = normalizeDocumentInkPenProfile({
    version: 2,
    pens: [...profile.pens, { id, ...documentInkPenDefaults(tool) }].slice(-DOCUMENT_INK_MAX_PENS),
  });
  return { profile: next, id: next.pens.some((pen) => pen.id === id) ? id : next.pens[0].id };
}

/** Delete: the profile without `id`; the last pen stays (Word keeps one). */
export function withDocumentInkPenRemoved(
  profile: DocumentInkPenProfile,
  id: string,
): DocumentInkPenProfile {
  if (profile.pens.length <= 1) return profile;
  return normalizeDocumentInkPenProfile({
    version: 2,
    pens: profile.pens.filter((pen) => pen.id !== id),
  });
}

export function documentInkPresetById(
  id: string,
  profile: DocumentInkPenProfile = defaultDocumentInkPenProfile(),
): DocumentInkPreset {
  return profile.pens.find((preset) => preset.id === id) ?? profile.pens[0];
}

/** A stroke keeps at most this many samples (the shared scribble bound). */
/** The most samples a stored stroke keeps: its custom path's command envelope. */
export const MAX_DOCUMENT_INK_POINTS = PRESENTATION_CUSTOM_PATH_AUTHORING_MAX_COMMANDS;

/** Word's three erasers: the Stroke Eraser removes whole strokes; the Small
 * and Medium Erasers rub out the part of a stroke under the pointer. */
export const DOCUMENT_INK_ERASERS = ["stroke", "small", "medium"] as const;
export type DocumentInkEraser = (typeof DOCUMENT_INK_ERASERS)[number];

/** An eraser's reach around the pointer, in CSS pixels. */
export function documentInkEraserRadiusPx(eraser: DocumentInkEraser): number {
  switch (eraser) {
    case "small":
      return 4;
    case "medium":
      return 11;
    default:
      return 7;
  }
}

/** The stroke eraser's reach around the pointer, in CSS pixels. */
export const DOCUMENT_INK_ERASER_RADIUS_PX = documentInkEraserRadiusPx("stroke");

/**
 * How a tool paints: a highlighter is translucent, a pen and a pencil opaque.
 * A pencil's lighter look is its texture ({@link documentInkPencilFilter}),
 * not a lower opacity: Word's black pencil has a fully black core.
 */
export function documentInkOpacity(tool: DocumentInkTool): number {
  if (tool === "pencil") return 1;
  return inkOpacityForTool(tool);
}

/**
 * Draw ▸ Tape: a strip covering the words is nearly opaque (they cannot be
 * read through it); one peeled back is a faint strip they show through, so
 * it can still be found and tapped closed again.
 */
export const DOCUMENT_INK_TAPE_OPACITY: Readonly<Record<InkTape, number>> = { covered: 0.96, revealed: 0.16 };

/** Draw ▸ Tape's strip widths (points): one line of body text, two, three. */
export const DOCUMENT_INK_TAPE_WIDTHS_PT: readonly number[] = [12, 18, 26];

/** How a stroke paints: its tool's opacity, or a strip of tape's state. */
export function documentInkStrokeOpacity(stroke: { readonly inkTool?: unknown; readonly inkTape?: unknown }): number {
  const tape = normalizeInkTape(stroke.inkTape);
  if (tape) return DOCUMENT_INK_TAPE_OPACITY[tape];
  return documentInkOpacity(normalizeDocumentInkTool(stroke.inkTool) ?? "pen");
}

/** The state a tap turns a strip of tape to. */
export function documentInkTapeFlipped(tape: InkTape): InkTape {
  return tape === "covered" ? "revealed" : "covered";
}

/** One primitive of an SVG filter, for both the DOM painter and React. */
export interface DocumentInkFilterPrimitive {
  readonly tag:
    | "feGaussianBlur"
    | "feTurbulence"
    | "feColorMatrix"
    | "feComposite"
    | "feMorphology";
  readonly attrs: Readonly<Record<string, string>>;
}

/** An SVG filter: its own attributes and its primitives, in order. */
export interface DocumentInkFilter {
  readonly attrs: Readonly<Record<string, string>>;
  readonly primitives: readonly DocumentInkFilterPrimitive[];
}

/** A region in the filtered element's user space (CSS pixels here). */
export interface DocumentInkFilterRegion {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

const round3 = (value: number) => String(Math.round(value * 1000) / 1000);

/**
 * Word's Pencil texture (MS-ODRAWXML `ST_InkEffectsType` "pencil"; Word
 * stores it as the InkML brush property `inkEffects`). The specification's
 * figure is a dense graphite core whose rim breaks up into grain, as a pencil
 * drags over paper tooth; the stroke's path itself stays smooth.
 *
 * Drawn as a stochastic threshold: the stroke's coverage is softened into a
 * profile (1 at the core, falling off at the rim), fine fractal noise is
 * subtracted, and what stays above the threshold is painted in the stroke's
 * own colour. The core stays nearly solid and the rim turns to specks, and
 * the specks may reach a little past the smooth outline, as they do in the
 * figure. The rim's width is the pencil's (a tip on paper), so it is capped
 * rather than growing with a thick stroke.
 *
 * The filter works in the user space of the element it is applied to, which
 * must be UNIFORM CSS pixels: under a stretched viewBox the grain would
 * stretch with it (and WebKit applies a CSS filter on an `<svg>` in its
 * viewBox space). The region is in the same pixels; an object-bounding-box
 * region would be empty for a flat stroke, whose path box has no height.
 *
 * `seed` fixes the grain, so a repaint draws the same specks.
 */
export function documentInkPencilFilter(
  strokeWidthPx: number,
  region: DocumentInkFilterRegion,
  seed: number,
): DocumentInkFilter {
  const width = Math.max(0.5, strokeWidthPx);
  const profileSigma = Math.min(2.2, Math.max(0.5, width * 0.32));
  const reach = Math.min(2, width * 0.2);
  const contrast = 3.5;
  const bias = -0.25;
  return {
    attrs: {
      filterUnits: "userSpaceOnUse",
      primitiveUnits: "userSpaceOnUse",
      x: round3(region.x),
      y: round3(region.y),
      width: round3(Math.max(1, region.width)),
      height: round3(Math.max(1, region.height)),
      "color-interpolation-filters": "sRGB",
    },
    primitives: [
      { tag: "feGaussianBlur", attrs: { in: "SourceAlpha", stdDeviation: round3(profileSigma), result: "profile" } },
      {
        tag: "feTurbulence",
        attrs: { type: "fractalNoise", baseFrequency: "1.1", numOctaves: "3", seed: String(seed), result: "noise" },
      },
      {
        tag: "feColorMatrix",
        attrs: {
          in: "noise",
          type: "matrix",
          values: "0 0 0 0 0  0 0 0 0 0  0 0 0 0 0  1 0 0 0 0",
          result: "grain",
        },
      },
      {
        tag: "feComposite",
        attrs: {
          in: "profile",
          in2: "grain",
          operator: "arithmetic",
          k1: "0",
          k2: String(contrast),
          k3: String(-contrast),
          k4: String(bias),
          result: "graphite",
        },
      },
      { tag: "feMorphology", attrs: { in: "SourceGraphic", operator: "dilate", radius: round3(reach), result: "reach" } },
      { tag: "feComposite", attrs: { in: "reach", in2: "graphite", operator: "in" } },
    ],
  };
}

/**
 * The pencil filter's region for a stroke being drawn: the points' box grown
 * by the stroke and the grain's reach, so the live preview filters only the
 * stroke and not the whole drawing surface.
 */
export function documentInkPencilRegionFor(
  points: readonly { readonly x: number; readonly y: number }[],
  strokeWidthPx: number,
): DocumentInkFilterRegion {
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const point of points) {
    minX = Math.min(minX, point.x);
    minY = Math.min(minY, point.y);
    maxX = Math.max(maxX, point.x);
    maxY = Math.max(maxY, point.y);
  }
  if (!Number.isFinite(minX)) return { x: 0, y: 0, width: 1, height: 1 };
  const pad = strokeWidthPx + 4;
  return { x: minX - pad, y: minY - pad, width: maxX - minX + 2 * pad, height: maxY - minY + 2 * pad };
}

/** A stable grain seed for one stroke (its path), so repaints match. */
export function documentInkPencilSeed(key: string): number {
  let hash = 2166136261;
  for (let index = 0; index < key.length; index += 1) {
    hash ^= key.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0) % 997;
}

/** The cut ends of a stroke whose first / last point is a cut. */
export function inkCutEndsOf(cutStart: boolean, cutEnd: boolean): InkCutEnds | undefined {
  if (cutStart) return cutEnd ? "both" : "start";
  return cutEnd ? "end" : undefined;
}

/**
 * How a stroke's ends are painted. A stroke with no cut end is one
 * round-capped line, as every stroke always was. A stroke with one is
 * butt-capped, so each cut end stops flat on the eraser's edge, and each end
 * the pen made gets its tip back as `tips`: a zero-length round-capped segment
 * there, a dot of the pen's width that with the line's flat end is exactly the
 * round tip. Paint the line and the tips in ONE group carrying the stroke's
 * opacity, so a highlighter's tip does not darken its line; and in a space
 * that is not scaled unevenly (a zero-length cap under `non-scaling-stroke`
 * in a stretched viewBox is an ellipse in WebKit).
 */
export function inkStrokeEndPaint<P>(
  points: readonly P[],
  cutEnds: InkCutEnds | undefined,
): { linecap: "round" | "butt"; tips: P[] } {
  if (!cutEnds || points.length === 0) return { linecap: "round", tips: [] };
  const tips: P[] = [];
  if (cutEnds === "end") tips.push(points[0]);
  if (cutEnds === "start") tips.push(points[points.length - 1]);
  return { linecap: "butt", tips };
}

/** The SVG `d` that paints `tips` as dots (zero-length round-capped segments). */
export function inkStrokeTipsPathData(tips: readonly DocumentInkPoint[]): string {
  return tips.map(({ x, y }) => `M${x} ${y}L${x} ${y}`).join("");
}

/** How much of its end segment a tip drawn along the stroke runs over. */
const INK_TIP_RUN = 0.01;

/**
 * The tips {@link inkStrokeEndPaint} gives back, for a space scaled unevenly
 * (a 0..100 box stretched to its frame, the width held by
 * `non-scaling-stroke`), where a zero-length cap is an ellipse in WebKit:
 * each tip is instead the last hundredth of the stroke's end segment, ending
 * at the tip, whose round cap is the pen's circle on screen. A dot (no
 * segment to run along) keeps the zero-length tip.
 */
export function inkStrokeTipsAlongPathData(
  points: readonly DocumentInkPoint[],
  cutEnds: InkCutEnds | undefined,
): string {
  const { tips } = inkStrokeEndPaint(points, cutEnds);
  return tips
    .map((tip) => {
      const ordered = tip === points[0] ? points : [...points].reverse();
      const toward = ordered.find((point) => point.x !== tip.x || point.y !== tip.y);
      if (!toward) return inkStrokeTipsPathData([tip]);
      const x = tip.x + (toward.x - tip.x) * INK_TIP_RUN;
      const y = tip.y + (toward.y - tip.y) * INK_TIP_RUN;
      return `M${x} ${y}L${tip.x} ${tip.y}`;
    })
    .join("");
}

function distanceToSegment(
  point: DocumentInkPoint,
  from: DocumentInkPoint,
  to: DocumentInkPoint,
): number {
  const dx = to.x - from.x;
  const dy = to.y - from.y;
  const lengthSquared = dx * dx + dy * dy;
  const t = lengthSquared > 0
    ? Math.max(0, Math.min(1, ((point.x - from.x) * dx + (point.y - from.y) * dy) / lengthSquared))
    : 0;
  return Math.hypot(point.x - (from.x + t * dx), point.y - (from.y + t * dy));
}

function cross(o: DocumentInkPoint, a: DocumentInkPoint, b: DocumentInkPoint): number {
  return (a.x - o.x) * (b.y - o.y) - (a.y - o.y) * (b.x - o.x);
}

/** The distance between segments a-b and c-d (0 where they cross). */
function segmentsDistance(
  a: DocumentInkPoint,
  b: DocumentInkPoint,
  c: DocumentInkPoint,
  d: DocumentInkPoint,
): number {
  const d1 = cross(c, d, a);
  const d2 = cross(c, d, b);
  const d3 = cross(a, b, c);
  const d4 = cross(a, b, d);
  if (((d1 > 0 && d2 < 0) || (d1 < 0 && d2 > 0)) && ((d3 > 0 && d4 < 0) || (d3 < 0 && d4 > 0))) {
    return 0;
  }
  return Math.min(
    distanceToSegment(a, c, d),
    distanceToSegment(b, c, d),
    distanceToSegment(c, a, b),
    distanceToSegment(d, a, b),
  );
}

/**
 * A piece of a stroke a point eraser leaves, from `from` to `to` along the
 * stroke's polyline (a position on it: the integer part is the segment, the
 * rest how far along that segment), and which of its ends are cuts.
 * Positions, not points, so each editor rebuilds the piece from its own stored
 * points ({@link documentInkPolylineBetween}): the space the eraser was
 * measured in is an affine image of them, and interpolation commutes with it.
 */
export interface DocumentInkStrokePiece {
  readonly from: number;
  readonly to: number;
  readonly cutEnds?: InkCutEnds;
}

/** A piece shorter than this (px) paints nothing worth keeping. */
const MIN_INK_PIECE_PX = 0.5;
/** The trace's smallest step, and how closely a cut is placed (px). */
const INK_CUT_STEP_PX = 0.25;
const INK_CUT_PRECISION_PX = 0.01;

/**
 * What a point eraser (the Small and Medium Erasers) leaves of a stroke: the
 * pieces outside the ink it rubbed out, `[]` when it rubbed out all of it,
 * or null when it touched none of it. Every argument is in one space of CSS
 * pixels: the stroke's samples, the eraser's path, its reach `radiusPx` and
 * the stroke's painted `strokeWidthPx`.
 *
 * The ink under the eraser is the painted line: wherever the line's cross
 * section (a pen width across the stroke) comes within the eraser's reach,
 * the stroke is cut there, across its whole width, and each cut is placed
 * where that stops. The piece left ends flat on the eraser's edge
 * ({@link InkCutEnds}): it neither paints back into the gap (a round tip
 * reaches half a pen width into it, and two translucent tips overlapping show
 * darker than the stroke) nor loses ink the eraser never touched (as a round
 * tip trimmed back to clear the gap does). An end the pen made keeps its
 * round tip unless the eraser touches that tip, which it then rubs off flat;
 * an end already cut stays flat.
 */
export function documentInkEraserCut(
  stroke: readonly DocumentInkPoint[],
  eraser: readonly DocumentInkPoint[],
  radiusPx: number,
  strokeWidthPx = 0,
  cutEnds?: InkCutEnds,
): DocumentInkStrokePiece[] | null {
  const count = stroke.length;
  if (count === 0 || eraser.length === 0) return null;
  const reach = Number.isFinite(radiusPx) && radiusPx > 0 ? radiusPx : 0;
  const half = Number.isFinite(strokeWidthPx) && strokeWidthPx > 0 ? strokeWidthPx / 2 : 0;
  type Segment = readonly [DocumentInkPoint, DocumentInkPoint];
  const eraserSegments: Segment[] =
    eraser.length === 1
      ? [[eraser[0], eraser[0]]]
      : eraser.slice(1).map((point, index) => [eraser[index], point] as const);
  const nearest = (from: DocumentInkPoint, to: DocumentInkPoint, segments: readonly Segment[]) => {
    let best = Infinity;
    for (const [a, b] of segments) {
      best = Math.min(best, segmentsDistance(from, to, a, b));
      if (best === 0) break;
    }
    return best;
  };
  const lengths: number[] = [];
  let total = 0;
  for (let index = 1; index < count; index += 1) {
    const length = Math.hypot(
      stroke[index].x - stroke[index - 1].x,
      stroke[index].y - stroke[index - 1].y,
    );
    lengths.push(length);
    total += length;
  }
  if (!(total > 0)) {
    // A dot: all of it or none of it.
    return nearest(stroke[0], stroke[0], eraserSegments) <= reach + half ? [] : null;
  }
  const last = count - 1;
  const hits: Array<[number, number]> = [];
  for (let index = 0; index < last; index += 1) {
    const length = lengths[index];
    if (!(length > 0)) continue;
    const a = stroke[index];
    const b = stroke[index + 1];
    const ux = (b.x - a.x) / length;
    const uy = (b.y - a.y) / length;
    const margin = reach + half + INK_CUT_STEP_PX;
    const near = eraserSegments.filter(
      ([p, q]) =>
        Math.max(p.x, q.x) >= Math.min(a.x, b.x) - margin &&
        Math.min(p.x, q.x) <= Math.max(a.x, b.x) + margin &&
        Math.max(p.y, q.y) >= Math.min(a.y, b.y) - margin &&
        Math.min(p.y, q.y) <= Math.max(a.y, b.y) + margin,
    );
    if (near.length === 0) continue;
    // How far the cross section at `t` is clear of the eraser's reach
    // (negative under it). The section moves `length` px per unit of `t`, so
    // this changes no faster, and the trace can step by its own value
    // without stepping over a cut.
    const clearance = (t: number) => {
      const x = a.x + (b.x - a.x) * t;
      const y = a.y + (b.y - a.y) * t;
      return (
        nearest({ x: x + uy * half, y: y - ux * half }, { x: x - uy * half, y: y + ux * half }, near) -
        reach
      );
    };
    let t = 0;
    let value = clearance(0);
    let inside = value <= 0;
    let enter = 0;
    while (t < 1) {
      const next = Math.min(1, t + Math.max(Math.abs(value), INK_CUT_STEP_PX) / length);
      const nextValue = clearance(next);
      if ((nextValue <= 0) !== inside) {
        // The cut, found to a hundredth of a pixel; what is kept always ends
        // on the side clear of the eraser.
        let lo = t;
        let hi = next;
        while ((hi - lo) * length > INK_CUT_PRECISION_PX) {
          const mid = (lo + hi) / 2;
          if ((clearance(mid) <= 0) === inside) lo = mid;
          else hi = mid;
        }
        if (inside) hits.push([index + enter, index + hi]);
        else enter = lo;
        inside = !inside;
      }
      t = next;
      value = nextValue;
    }
    if (inside) hits.push([index + enter, index + 1]);
  }
  // A round tip reaches half a pen width past its end, behind the end's cross
  // section: an eraser touching only that rubs the tip off, flat at the end.
  const tipTouched = (end: DocumentInkPoint, backX: number, backY: number) =>
    half > 0 &&
    eraserSegments.some(([p, q]) => {
      const sp = (p.x - end.x) * backX + (p.y - end.y) * backY;
      const sq = (q.x - end.x) * backX + (q.y - end.y) * backY;
      if (sp < 0 && sq < 0) return false;
      // The part of the eraser's segment behind the end.
      const lo = sp >= 0 ? 0 : sp / (sp - sq);
      const hi = sq >= 0 ? 1 : sp / (sp - sq);
      return (
        distanceToSegment(
          end,
          { x: p.x + (q.x - p.x) * lo, y: p.y + (q.y - p.y) * lo },
          { x: p.x + (q.x - p.x) * hi, y: p.y + (q.y - p.y) * hi },
        ) <=
        reach + half
      );
    });
  const startCut = cutEnds === "start" || cutEnds === "both";
  let endCut = cutEnds === "end" || cutEnds === "both";
  if (!startCut) {
    const segment = lengths.findIndex((length) => length > 0);
    const [a, b] = [stroke[segment], stroke[segment + 1]];
    const length = lengths[segment];
    if (tipTouched(stroke[0], (a.x - b.x) / length, (a.y - b.y) / length)) hits.push([0, 0]);
  }
  if (!endCut) {
    let segment = lengths.length - 1;
    while (!(lengths[segment] > 0)) segment -= 1;
    const [a, b] = [stroke[segment], stroke[segment + 1]];
    const length = lengths[segment];
    if (tipTouched(stroke[last], (b.x - a.x) / length, (b.y - a.y) / length)) {
      endCut = true;
      hits.push([last, last]);
    }
  }
  if (hits.length === 0) return null;
  hits.sort((left, right) => left[0] - right[0]);
  // Arc length at a position, to drop the slivers a cut can leave.
  const offsets = [0];
  for (const length of lengths) offsets.push(offsets[offsets.length - 1] + length);
  const arc = (position: number) => {
    const index = Math.min(last - 1, Math.floor(position));
    return offsets[index] + (position - index) * lengths[index];
  };
  const pieces: DocumentInkStrokePiece[] = [];
  const keep = (from: number, to: number, cutStart: boolean, cutEnd: boolean) => {
    if (arc(to) - arc(from) < MIN_INK_PIECE_PX) return;
    const ends = inkCutEndsOf(cutStart, cutEnd);
    pieces.push({ from, to, ...(ends ? { cutEnds: ends } : {}) });
  };
  let cursor = 0;
  let cursorCut = startCut;
  for (const [from, to] of hits) {
    if (from > cursor) keep(cursor, from, cursorCut, true);
    if (to >= cursor) {
      cursor = to;
      cursorCut = true;
    }
  }
  if (cursor < last) keep(cursor, last, cursorCut, endCut);
  return pieces;
}

/**
 * The polyline from position `from` to position `to` along `points` (see
 * {@link DocumentInkStrokePiece}): its two ends interpolated, the samples
 * between them as they are.
 */
/** The point `t` of the way from `a` to `b`, its pressure interpolated with it. */
function inkPointBetween(a: DocumentInkPoint, b: DocumentInkPoint, t: number): DocumentInkPoint {
  const point: DocumentInkPoint = { x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t };
  const pa = a.p;
  const pb = b.p;
  if (pa !== undefined || pb !== undefined) {
    const from = pa ?? pb ?? 0.5;
    const to = pb ?? pa ?? 0.5;
    point.p = from + (to - from) * t;
  }
  return point;
}

export function documentInkPolylineBetween(
  points: readonly DocumentInkPoint[],
  from: number,
  to: number,
): DocumentInkPoint[] {
  if (points.length < 2) return points.slice(0, 1).map(({ x, y }) => ({ x, y }));
  const at = (position: number) => {
    const index = Math.min(points.length - 2, Math.max(0, Math.floor(position)));
    const t = Math.min(1, Math.max(0, position - index));
    const a = points[index];
    const b = points[index + 1];
    return inkPointBetween(a, b, t);
  };
  const line = [at(from)];
  for (let index = Math.floor(from) + 1; index < to; index += 1) {
    line.push(inkPointBetween(points[index], points[index], 0));
  }
  line.push(at(to));
  return line;
}

/** A box ink is kept inside: the page a stroke starts on, the trackpad pad. */
export interface DocumentInkBounds {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

/** The part of segment a→b inside `bounds`, as fractions along it
 * (Liang-Barsky), or null when it misses the box. */
function inkSegmentInside(
  a: DocumentInkPoint,
  b: DocumentInkPoint,
  bounds: DocumentInkBounds,
): [number, number] | null {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  let enter = 0;
  let leave = 1;
  for (const [p, q] of [
    [-dx, a.x - bounds.left],
    [dx, bounds.right - a.x],
    [-dy, a.y - bounds.top],
    [dy, bounds.bottom - a.y],
  ] as const) {
    if (p === 0) {
      if (q < 0) return null;
      continue;
    }
    const t = q / p;
    if (p < 0) {
      if (t > leave) return null;
      enter = Math.max(enter, t);
    } else {
      if (t < enter) return null;
      leave = Math.min(leave, t);
    }
  }
  return [enter, leave];
}

/**
 * What a stroke leaves inside `bounds` (Word: the page edge cuts ink off):
 * each run of samples inside, ended where the stroke crosses out and started
 * again where it comes back, so the line reaches the edge and stops there.
 */
export function documentInkRunsInside(
  points: readonly DocumentInkPoint[],
  bounds: DocumentInkBounds,
): DocumentInkPoint[][] {
  const inside = (point: DocumentInkPoint) =>
    point.x >= bounds.left &&
    point.x <= bounds.right &&
    point.y >= bounds.top &&
    point.y <= bounds.bottom;
  if (points.length === 1) return inside(points[0]) ? [[points[0]]] : [];
  const at = (a: DocumentInkPoint, b: DocumentInkPoint, t: number) =>
    t === 0 ? a : t === 1 ? b : inkPointBetween(a, b, t);
  const runs: DocumentInkPoint[][] = [];
  let run: DocumentInkPoint[] | null = null;
  for (let index = 1; index < points.length; index += 1) {
    const a = points[index - 1];
    const b = points[index];
    const part = inkSegmentInside(a, b, bounds);
    if (!part) {
      if (run) runs.push(run);
      run = null;
      continue;
    }
    const [enter, leave] = part;
    if (!run || enter > 0) {
      if (run) runs.push(run);
      run = [at(a, b, enter)];
    }
    run.push(at(a, b, leave));
    if (leave < 1) {
      runs.push(run);
      run = null;
    }
  }
  if (run) runs.push(run);
  return runs;
}

/** The floating anchor an ink stroke is laid out and saved with. */
export type DocumentInkShapeAttrs = DocumentShapeAttrs & {
  readonly floatAnchor: FloatAnchor;
};

/**
 * The anchor of an ink stroke: In Front Of Text (`wp:wrapNone`, over the
 * words), its offsets measured from the column's left edge and the top of the
 * paragraph it is anchored to -- the same origins the offline DOCX writer
 * states for every floating shape, so the page and the file agree.
 */
export function documentInkFloatAnchor(attrs: {
  readonly xEmu: number;
  readonly yEmu: number;
  readonly widthEmu: number;
  readonly heightEmu: number;
  readonly relativeHeight?: number;
}): FloatAnchor {
  return {
    version: 1,
    horizontal: { relativeFrom: "column", offsetEmu: attrs.xEmu },
    vertical: { relativeFrom: "paragraph", offsetEmu: attrs.yEmu },
    widthEmu: attrs.widthEmu,
    heightEmu: attrs.heightEmu,
    wrap: { kind: "none" },
    behindDoc: false,
    locked: false,
    layoutInCell: true,
    allowOverlap: true,
    hidden: false,
    relativeHeight: attrs.relativeHeight ?? DOCUMENT_SHAPE_DEFAULT_RELATIVE_HEIGHT,
    simplePos: false,
    distTEmu: 0,
    distBEmu: 0,
    distLEmu: 0,
    distREmu: 0,
    effectLEmu: 0,
    effectTEmu: 0,
    effectREmu: 0,
    effectBEmu: 0,
    rotation: 0,
    flipH: false,
    flipV: false,
  };
}

/** What a stroke states about the pen that drew it. */
export type DocumentInkStrokeStyle = Pick<DocumentInkPreset, "tool" | "color" | "widthPt" | "effect" | "tape">;

/**
 * One pointer gesture as one ink stroke. `points` are EMU from the stroke's
 * anchor origins (column left, paragraph top). A tap is a dot; a straight
 * line is real ink; the box is padded by the stroke's own half width so the
 * paint never clips, and floored at the shape minimum so the selection grips
 * have something to hold. `cutEnds` is a piece a point eraser left.
 */
export function documentInkStrokeAttrs(
  points: readonly DocumentInkPoint[],
  preset: DocumentInkStrokeStyle,
  relativeHeight = DOCUMENT_SHAPE_DEFAULT_RELATIVE_HEIGHT,
  cutEnds?: InkCutEnds,
): DocumentInkShapeAttrs | null {
  // The samples as drawn, every one while they fit the stored stroke
  // (inkLandedSampleIndexes): the landed curve is the very curve the running
  // stroke drew. Thinning to an even spacing cut a long scribble's reversals
  // short and the whole stroke drew in on the lift.
  const finite = points.filter((point) => Number.isFinite(point.x) && Number.isFinite(point.y));
  const samples = inkLandedSampleIndexes(finite, DOCUMENT_SHAPE_EMU_PER_PX / 5, MAX_DOCUMENT_INK_POINTS).map(
    (index) => finite[index]!,
  );
  if (samples.length === 0) return null;
  if (samples.length === 1) samples.push({ ...samples[0] });
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const point of samples) {
    minX = Math.min(minX, point.x);
    minY = Math.min(minY, point.y);
    maxX = Math.max(maxX, point.x);
    maxY = Math.max(maxY, point.y);
  }
  const strokeEmu = preset.widthPt * 12_700;
  const padEmu = strokeEmu / 2 + 2 * DOCUMENT_SHAPE_EMU_PER_PX;
  let left = minX - padEmu;
  let top = minY - padEmu;
  let width = maxX - minX + 2 * padEmu;
  let height = maxY - minY + 2 * padEmu;
  if (width < DOCUMENT_SHAPE_MIN_EDGE_EMU) {
    left -= (DOCUMENT_SHAPE_MIN_EDGE_EMU - width) / 2;
    width = DOCUMENT_SHAPE_MIN_EDGE_EMU;
  }
  if (height < DOCUMENT_SHAPE_MIN_EDGE_EMU) {
    top -= (DOCUMENT_SHAPE_MIN_EDGE_EMU - height) / 2;
    height = DOCUMENT_SHAPE_MIN_EDGE_EMU;
  }
  const grid = (value: number, origin: number, extent: number) =>
    Math.max(
      0,
      Math.min(
        PRESENTATION_CUSTOM_PATH_GRID,
        Math.round(((value - origin) / extent) * PRESENTATION_CUSTOM_PATH_GRID),
      ),
    );
  const commands: PresentationCustomPathCommand[] = samples.map(
    (point, index) => ({
      type: index === 0 ? "moveTo" : "lineTo",
      x: grid(point.x, left, width),
      y: grid(point.y, top, height),
    }),
  );
  const customPath = normalizePresentationCustomPath({
    version: 1,
    closed: false,
    commands,
  });
  if (!customPath) return null;
  const effect = preset.tool === "pen" ? preset.effect : undefined;
  // A pen's pressure widens and thins the stroke as it is painted; a
  // highlighter's band is one width in Word, so its pressures are dropped.
  const inkPressure = preset.tool === "highlighter" ? undefined : documentInkPressureOf(samples);
  const attrs = normalizeDocumentShapeAttrs({
    // The nearest auto-shape for a host that draws only presets. Not a line
    // kind: both DOCX writers keep a line's own `prstGeom` and would drop the
    // stroke's `a:custGeom`; a rectangle takes the outline as its geometry.
    preset: "rect",
    customPath,
    xEmu: Math.round(left),
    yEmu: Math.round(top),
    widthEmu: Math.max(1, Math.round(width)),
    heightEmu: Math.max(1, Math.round(height)),
    fill: "none",
    outline: effect ? documentInkEffectPaint(effect).solid : preset.color,
    // A steady pressure is drawn at its width, as the running stroke was.
    outlineWidthPt:
      preset.tool === "highlighter"
        ? preset.widthPt
        : Math.round(preset.widthPt * inkSteadyPressureWidthFactor(samples.map((sample) => (sample as { p?: number }).p)) * 100) / 100,
    outlineStyle: "solid",
    wrapMode: "inFrontOfText",
    relativeHeight,
    zOrderStackComplete: true,
    inkTool: preset.tool,
    ...(effect ? { inkEffect: effect } : {}),
    ...(cutEnds ? { inkCutEnds: cutEnds } : {}),
    ...(inkPressure ? { inkPressure } : {}),
    // A strip of tape (Draw ▸ Tape), covering or peeled back.
    ...(preset.tape ? { inkTape: preset.tape } : {}),
  });
  if (!attrs.customPath || attrs.inkTool !== preset.tool) return null;
  return { ...attrs, floatAnchor: documentInkFloatAnchor(attrs) };
}

/** Whether these attrs are a stroke the Draw tab drew (an ink shape). */
export function isDocumentInkShapeAttrs(value: DocumentShapeAttrsInput): boolean {
  const attrs = (value ?? {}) as Readonly<Record<string, unknown>>;
  return (
    normalizeDocumentInkTool(attrs.inkTool) !== undefined &&
    normalizePresentationCustomPath(attrs.customPath) !== null
  );
}

/** The pen an existing stroke states, for a stroke rebuilt from its parts. */
export function documentInkStrokeStyle(value: DocumentShapeAttrsInput): DocumentInkStrokeStyle | null {
  const attrs = (value ?? {}) as Readonly<Record<string, unknown>>;
  const tool = normalizeDocumentInkTool(attrs.inkTool);
  if (!tool) return null;
  const color = typeof attrs.outline === "string" && INK_COLOR.test(attrs.outline.toLowerCase())
    ? attrs.outline.toLowerCase()
    : "#000000";
  const widthPt = typeof attrs.outlineWidthPt === "number" && attrs.outlineWidthPt > 0 ? attrs.outlineWidthPt : 1.4;
  const effect = tool === "pen" ? normalizeDocumentInkEffect(attrs.inkEffect) : undefined;
  const tape = normalizeInkTape(attrs.inkTape);
  return { tool, color, widthPt, ...(effect ? { effect } : {}), ...(tape ? { tape } : {}) };
}

/**
 * The stroke's samples as fractions of its own box (0..1 on each axis), in
 * drawing order -- what the erasers test against once the box is known on
 * screen.
 */
export function documentInkPathFractions(
  value: DocumentShapeAttrsInput,
): DocumentInkPoint[] {
  const attrs = (value ?? {}) as Readonly<Record<string, unknown>>;
  const path = normalizePresentationCustomPath(attrs.customPath);
  if (!path) return [];
  const points: DocumentInkPoint[] = [];
  for (const command of path.commands) {
    if (command.type === "close" || command.type === "arcTo") continue;
    points.push({
      x: command.x / PRESENTATION_CUSTOM_PATH_GRID,
      y: command.y / PRESENTATION_CUSTOM_PATH_GRID,
    });
  }
  const pressures = documentInkPressures(attrs.inkPressure, points.length);
  if (pressures) pressures.forEach((pressure, index) => { points[index]!.p = pressure; });
  return points;
}
