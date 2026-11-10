/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure geometry planner + interpolator for the present-mode MORPH transition
 * (PowerPoint / WPS "Morph" parity).
 *
 * A morph animates the step from the PREVIOUS slide INTO the slide that declares
 * `transition: "morph"`. Blocks present on BOTH slides (see {@link morphPlan} for
 * how an object on one slide is recognised on the other) are TWEENED from their slide-A geometry to their slide-B geometry; blocks only
 * on the new (`to`) slide FADE IN; blocks only on the previous (`from`) slide
 * FADE OUT.
 *
 * This module is intentionally dependency-light: it operates on plain `{ id, geo }`
 * inputs (no React, no DOM, no `PresentationBlock` import) so it can be unit-tested
 * in isolation and reused from any caller. The component layer is responsible for
 * extracting geometry off real blocks and for rendering the matched/entering blocks
 * with the existing present-mode block renderer; this module only decides the
 * matched / entering / leaving partition and the per-field interpolation.
 *
 * All geometry is in EMU (914,400 per inch) for position/size and DEGREES for
 * rotation — the same units the presentation blocks persist. The functions are
 * TOTAL and side-effect-free.
 */

/** A block's animatable geometry: position + size (EMU) + rotation (degrees). */
export interface MorphGeo {
  left: number;
  top: number;
  width: number;
  height: number;
  rotation: number;
}

/** The minimal shape the planner needs from each side's blocks. */
export interface MorphBlockGeo {
  id: string;
  geo: MorphGeo;
  /** The block's type; a copy or a `!!` name only matches its own type. */
  type?: string;
  /**
   * The block this one was copied from by Duplicate Slide or slide Paste
   * (`morphSourceId`), traced back to the first original.
   */
  sourceId?: string;
  /** The Selection Pane name. */
  name?: string;
}

/** One matched block, with its from/to geometry. */
export interface MorphMatch {
  /** The destination block's id. */
  id: string;
  /** The previous slide's block; equal to `id` unless matched as a copy or by name. */
  fromId: string;
  from: MorphGeo;
  to: MorphGeo;
}

/**
 * The full morph plan for advancing from `fromBlocks` into `toBlocks`.
 *
 * - `matched`: ids present on BOTH sides, in DOCUMENT ORDER of `toBlocks` (the
 *   destination slide drives ordering so the morph layer paints in the same
 *   z-order as the settled slide). Each carries `from`/`to` geometry.
 * - `entering`: ids present ONLY on `toBlocks` (fade in), in `toBlocks` order.
 * - `leaving`: ids present ONLY on `fromBlocks` (fade out), in `fromBlocks` order
 *   (the leftover order, after the `toBlocks`-driven matched set).
 *
 * Duplicate ids within a side are collapsed to the FIRST occurrence (a slide
 * should not carry two blocks with the same id; if it does, the planner is
 * deterministic rather than crashing). An id that appears on both sides is
 * matched even if the two blocks differ in TYPE — matching is purely by id, and
 * the component renders the destination block's content, so a type change simply
 * means the new content tweens in at the matched geometry.
 */
export interface MorphPlan {
  matched: MorphMatch[];
  entering: string[];
  leaving: string[];
}

/**
 * Build the matched / entering / leaving partition for a forward morph.
 *
 * An object is recognised on the next slide, in this order:
 *  1. the same block (`id`);
 *  2. the same object on a duplicated slide: a block and its copy made by
 *     Duplicate Slide or slide Paste share the original's id as `sourceId`;
 *  3. the same `!!` name on two blocks of the same type — Microsoft Support,
 *     "Morph transition: Tips and tricks": "If Morph sees two objects of the
 *     same type with the same name, starting with !!, across two slides, it'll
 *     morph one object to the other", and a name should be unique on a slide.
 * Every match is one to one. Order is deterministic: `matched` and `entering`
 * follow `toBlocks` document order; `leaving` follows `fromBlocks` document
 * order. Within a single side, only the FIRST block for a given id is
 * considered (later duplicates are ignored).
 */
export function morphPlan(
  fromBlocks: readonly MorphBlockGeo[],
  toBlocks: readonly MorphBlockGeo[],
): MorphPlan {
  const firstById = (blocks: readonly MorphBlockGeo[]) => {
    const byId = new Map<string, MorphBlockGeo>();
    for (const block of blocks) if (!byId.has(block.id)) byId.set(block.id, block);
    return byId;
  };
  // First-occurrence-wins maps so a duplicated id never yields two matches and
  // the geometry is taken from the first block carrying that id.
  const fromById = firstById(fromBlocks);
  const toById = firstById(toBlocks);
  const pairs = new Map<string, MorphBlockGeo>();
  const taken = new Set<string>();
  for (const [id, block] of toById) {
    const from = fromById.get(id);
    if (!from) continue;
    pairs.set(id, from);
    taken.add(id);
  }
  const pairBy = (key: (block: MorphBlockGeo) => string | undefined) => {
    const candidates = new Map<string, MorphBlockGeo>();
    for (const [id, from] of fromById) {
      const value = key(from);
      if (taken.has(id) || value === undefined) continue;
      const slot = `${from.type ?? ""}\u0000${value}`;
      if (!candidates.has(slot)) candidates.set(slot, from);
    }
    for (const [id, to] of toById) {
      const value = key(to);
      if (pairs.has(id) || value === undefined) continue;
      const slot = `${to.type ?? ""}\u0000${value}`;
      const from = candidates.get(slot);
      if (!from || taken.has(from.id)) continue;
      pairs.set(id, from);
      taken.add(from.id);
    }
  };
  pairBy((block) => block.sourceId ?? block.id);
  pairBy((block) =>
    typeof block.name === "string" && block.name.startsWith("!!") && block.name.length > 2
      ? block.name
      : undefined,
  );

  const matched: MorphMatch[] = [];
  const entering: string[] = [];
  for (const [id, block] of toById) {
    const from = pairs.get(id);
    if (from) matched.push({ id, fromId: from.id, from: from.geo, to: block.geo });
    else entering.push(id);
  }
  const leaving = [...fromById.keys()].filter((id) => !taken.has(id));

  return { matched, entering, leaving };
}

/**
 * Clamp `t` into the unit interval. NaN collapses to 0 (a non-finite progress is
 * treated as "not started" rather than propagating NaN into geometry).
 */
export function clampUnit(t: number): number {
  if (!Number.isFinite(t)) return 0;
  if (t <= 0) return 0;
  if (t >= 1) return 1;
  return t;
}

/**
 * Standard ease-in-out cubic on the unit interval. `t` is clamped to [0,1] first
 * so an out-of-range progress can never produce an out-of-range eased value. The
 * curve is monotonic non-decreasing with `f(0)=0`, `f(0.5)=0.5`, `f(1)=1`.
 */
export function easeInOutCubic(t: number): number {
  const x = clampUnit(t);
  return x < 0.5 ? 4 * x * x * x : 1 - Math.pow(-2 * x + 2, 3) / 2;
}

/** Linear interpolation of a single field. */
function lerp(from: number, to: number, eased: number): number {
  return from + (to - from) * eased;
}

/**
 * Interpolate every geometry field from `from` to `to` at progress `t∈[0,1]`,
 * applying {@link easeInOutCubic} easing. `t` is clamped, so `t≤0` returns the
 * `from` geometry and `t≥1` returns the `to` geometry exactly. Rotation is
 * interpolated as a plain numeric field (degrees) — callers that need shortest-arc
 * behavior should pre-normalize their angles; this keeps the function total and
 * faithful to the persisted numbers.
 */
export function lerpGeo(from: MorphGeo, to: MorphGeo, t: number): MorphGeo {
  const eased = easeInOutCubic(t);
  return {
    left: lerp(from.left, to.left, eased),
    top: lerp(from.top, to.top, eased),
    width: lerp(from.width, to.width, eased),
    height: lerp(from.height, to.height, eased),
    rotation: lerp(from.rotation, to.rotation, eased),
  };
}

/** What a block is apart from where it stands and who it is. */
const MORPH_FRAME_KEYS = new Set([
  "id",
  "left",
  "top",
  "width",
  "height",
  "rotation",
  "zIndex",
  "flipH",
  "flipV",
  "name",
  "sourceId",
  "locked",
  "source",
]);

/**
 * Whether a matched pair looks different apart from its frame: its colour,
 * formatting, shape, words or picture changed (Microsoft Support, "Morph
 * transition: Tips and tricks": those change smoothly too). The pair is then
 * drawn as both, the first giving way to the second, rather than as the
 * second throughout.
 */
export function morphContentChanged(from: object, to: object): boolean {
  const content = (block: object) => {
    const out: Record<string, unknown> = {};
    for (const [key, value] of Object.entries(block)) {
      if (!MORPH_FRAME_KEYS.has(key)) out[key] = value;
    }
    const drawing = (block as { source?: { nativeShapeDrawing?: { shapeKind?: unknown } } }).source?.nativeShapeDrawing;
    if (drawing?.shapeKind !== undefined) out.nativeShapeKind = drawing.shapeKind;
    return JSON.stringify(Object.fromEntries(Object.entries(out).sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))));
  };
  return content(from) !== content(to);
}
