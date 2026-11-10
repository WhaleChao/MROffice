/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * MOTION PATHS for the present-mode animation timeline (PowerPoint's fourth
 * animation class, alongside entrance / emphasis / exit).
 *
 * A motion path moves a block along a trajectory while its timeline step plays.
 * Unlike the entrance/emphasis/exit set — which are fixed named keyframes shared
 * by every block — a motion path is GEOMETRY: an ordered list of waypoints in EMU
 * RELATIVE to the block's start position (the first implicit waypoint is always
 * the origin `{x:0,y:0}` — the block's resting spot). Authors pick a NAMED PRESET
 * (`line-right`, `arc`, `loop`, …) which expands to a canonical waypoint list, or
 * a `custom` path of explicit waypoints.
 *
 * This module is the single source of truth for:
 *  - the allowed preset names + their canonical waypoint expansions,
 *  - the EMU clamp (mirrors the connectors-style `MAX_*_EMU` clamp so a hand-
 *    edited / stale path can never blow up coordinates),
 *  - normalization + validation of a persisted `motionPath` (bad preset → drop;
 *    NaN / huge / empty-custom → clamp or drop),
 *  - the PURE keyframe-CSS generator that maps a normalized motion path to the
 *    independent `translate:` keyframe percentages the renderer plays.
 *
 * Keep it dependency-free (no React / no DOM / no zod) so it can be unit tested
 * in isolation and imported from both the zod schema and the present-mode
 * renderer. Mirrors the conventions of the sibling presentation libs: a frozen
 * allowed list, a lenient normalizer that collapses stale values to a safe
 * default, and pure deterministic resolvers — no Date / no random anywhere, so a
 * given path always yields byte-identical CSS (SSR-safe).
 *
 * ── RENDER APPROACH (translate keyframes, NOT CSS `offset-path`) ──────────────
 * The present-mode block renderer positions every block with PERCENTAGE geometry
 * relative to the slide bounds (see `blockGeometryStyle` in PresentationWindow:
 * `left: (leftEmu / slideWidthEmu) * 100%`). CSS `offset-path: path(...)` takes
 * an SVG path in PIXELS relative to the element's OWN border-box — a reference
 * frame the renderer never computes (it is fully responsive, % of the slide). A
 * `translate(%)` is likewise BOX-relative (% of the element's own size), the
 * wrong frame for slide-relative EMU waypoints.
 *
 * So waypoints are emitted as a fraction of the SLIDE, realized with CONTAINER
 * QUERY length units (`cqw` / `cqh`): the present canvas slide is a size query
 * container, and a `translate: N cqw M cqh` moves the block by N% of the slide
 * width and M% of the slide height regardless of the block's own size — exactly
 * the slide-relative frame the percentage geometry already uses. This is
 * deterministic, SSR-stable, and needs only `container-type: size` on the slide
 * (added in styles/native-page-window.css). `offset-path` was rejected for the
 * box-vs-slide reference-frame mismatch above.
 *
 * ── WHY THE INDEPENDENT `translate:` PROPERTY, NOT `transform: translate()` ────
 * A block can carry an entrance/emphasis/exit animation (most of which animate
 * `transform` — fly-in, zoom, spin, pulse, bounce, grow-shrink) AND a motion path
 * at the same time, plus a static `transform: rotate()` for a rotated block. Two
 * @keyframes animating the SAME property do not compose — per CSS Animations L1
 * the one listed LAST wins that property — so a `transform: translate()` motion
 * keyframe would silently SUPPRESS a transform-based entrance/emphasis. The
 * individual `translate` / `rotate` / `scale` properties are applied separately
 * from (and before) the `transform` property and COMPOSE with it. Driving motion
 * via `translate:` therefore stacks cleanly on top of a transform-based entrance,
 * an emphasis effect, and a rotation — none clobbers the others.
 */

/**
 * The named motion-path PRESETS. `custom` is special: it carries no canonical
 * expansion — its waypoints come from the step's persisted `waypoints` list.
 *
 *  - `line-right` / `line-left` / `line-down` / `line-up`: a straight slide one
 *    "unit" in the named direction.
 *  - `arc`: a shallow rightward arc (up-and-over) ending one unit to the right.
 *  - `loop` / `circle`: a closed clockwise loop returning to the start (the block
 *    ends where it began — an emphasis-style round trip).
 *  - `custom`: an explicit author-supplied waypoint list.
 *
 * Every non-`custom` value has a deterministic waypoint expansion in
 * {@link PRESET_WAYPOINTS}; the keyframe generator turns any of them into CSS.
 */
export const MOTION_PATH_PRESETS = [
  "line-right",
  "line-left",
  "line-down",
  "line-up",
  "arc",
  "loop",
  "circle",
  "custom",
] as const;

export type MotionPathPreset = (typeof MOTION_PATH_PRESETS)[number];

const MOTION_PATH_PRESET_SET: ReadonlySet<string> = new Set(MOTION_PATH_PRESETS);

/**
 * Type guard: is `value` exactly one of the allowed preset names? Exact,
 * case-sensitive — use {@link normalizeMotionPathPreset} for lenient input.
 */
export function isMotionPathPreset(value: unknown): value is MotionPathPreset {
  return typeof value === "string" && MOTION_PATH_PRESET_SET.has(value);
}

/**
 * Normalize an arbitrary persisted value into an allowed preset, or `null` when
 * it cannot be salvaged. Unknown / empty / non-string values → `null` (so a
 * malformed motionPath is DROPPED, never silently rewritten to a wrong preset).
 * The comparison trims surrounding whitespace and is case-insensitive.
 */
export function normalizeMotionPathPreset(
  value: unknown,
): MotionPathPreset | null {
  if (typeof value !== "string") return null;
  const candidate = value.trim().toLowerCase();
  return isMotionPathPreset(candidate) ? candidate : null;
}

/**
 * A single waypoint: an EMU OFFSET from the block's start position. `{x:0,y:0}`
 * is the block's resting spot (the implicit first waypoint of every path).
 */
export interface MotionWaypoint {
  x: number;
  y: number;
}

/**
 * The EMU "unit" a preset moves by — one third of the default slide width, a
 * comfortable on-slide travel that scales with the slide because every offset is
 * ultimately emitted as a FRACTION of the slide (see the keyframe generator). A
 * fixed EMU keeps the presets deterministic + independent of the block's size.
 */
const PRESET_UNIT_EMU = 3_048_000; // 1/3 of the 9,144,000 EMU default slide width

/**
 * The schema's hard EMU bound for a waypoint offset (mirrors
 * PRESENTATION_GEOMETRY_EMU_LIMIT in native-page-schemas.ts and
 * MAX_PRESENTATION_GEOMETRY_EMU / MAX_CONNECTOR_EMU). A custom waypoint clamped
 * to this can never blow up the persisted payload or send the block arbitrarily
 * far off-slide. Symmetric ± so a path may move the block in any direction.
 */
export const MAX_MOTION_PATH_EMU = 20_000_000;

/** Max waypoints a custom path may carry (bounds the persisted payload + CSS). */
export const MAX_MOTION_PATH_WAYPOINTS = 60;

/** Clamp a single coordinate into the symmetric EMU window; NaN/∞ → 0. */
function clampMotionCoord(value: unknown): number {
  const raw = typeof value === "number" && Number.isFinite(value) ? value : 0;
  const rounded = Math.round(raw);
  if (rounded < -MAX_MOTION_PATH_EMU) return -MAX_MOTION_PATH_EMU;
  if (rounded > MAX_MOTION_PATH_EMU) return MAX_MOTION_PATH_EMU;
  return rounded;
}

/**
 * Canonical EMU waypoint expansion for each non-`custom` preset, relative to the
 * block's start (the implicit `{0,0}` origin is NOT repeated here — the keyframe
 * generator always prepends it). Deterministic + frozen.
 *
 *  - lines: a single endpoint one unit away.
 *  - arc: rise to a peak, then settle one unit to the right and back to baseline
 *    (a shallow up-and-over).
 *  - loop / circle: four quarter points of a clockwise square-ish loop that
 *    returns to the origin (the block ends where it began).
 */
const U = PRESET_UNIT_EMU;
const PRESET_WAYPOINTS: Readonly<
  Record<Exclude<MotionPathPreset, "custom">, readonly MotionWaypoint[]>
> = Object.freeze({
  "line-right": [{ x: U, y: 0 }],
  "line-left": [{ x: -U, y: 0 }],
  "line-down": [{ x: 0, y: U }],
  "line-up": [{ x: 0, y: -U }],
  arc: [
    { x: Math.round(U / 2), y: -Math.round(U / 2) },
    { x: U, y: 0 },
  ],
  loop: [
    { x: U, y: 0 },
    { x: U, y: U },
    { x: 0, y: U },
    { x: 0, y: 0 },
  ],
  circle: [
    { x: U, y: 0 },
    { x: U, y: U },
    { x: 0, y: U },
    { x: 0, y: 0 },
  ],
});

/**
 * A normalized, render-safe motion path on a timeline step. `preset` names the
 * kind; `waypoints` is the resolved EMU offset list (relative to the block's
 * start, the implicit origin excluded). For a non-`custom` preset the waypoints
 * are the canonical expansion; for `custom` they are the clamped author list.
 */
export interface MotionPath {
  preset: MotionPathPreset;
  /** EMU offsets from the block's start; never empty for a valid path. */
  waypoints: MotionWaypoint[];
}

function asRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

/**
 * Clamp + bound a RAW custom waypoint list into render-safe {@link MotionWaypoint}s.
 * Non-array → `[]`. Each entry must be an object with numeric-ish x/y; x/y are
 * clamped into the EMU window (NaN/∞ → 0). The list is capped at
 * {@link MAX_MOTION_PATH_WAYPOINTS}. Pure: never mutates its input.
 */
export function clampWaypoints(raw: unknown): MotionWaypoint[] {
  if (!Array.isArray(raw)) return [];
  const out: MotionWaypoint[] = [];
  for (const entry of raw) {
    if (out.length >= MAX_MOTION_PATH_WAYPOINTS) break;
    const rec = asRecord(entry);
    if (!rec) continue;
    out.push({ x: clampMotionCoord(rec.x), y: clampMotionCoord(rec.y) });
  }
  return out;
}

/**
 * Validate + normalize a RAW `motionPath` value into a {@link MotionPath}, or
 * `null` when it cannot be salvaged into a renderable path.
 *
 * Rules (each is exercised by the adversarial test-suite):
 *  - Non-object → `null`.
 *  - `preset` must normalize to an allowed name (bad/unknown/empty → `null`,
 *    the whole motionPath is DROPPED rather than guessed).
 *  - A NON-`custom` preset uses its canonical waypoint expansion (any persisted
 *    `waypoints` are ignored — the preset IS the geometry).
 *  - A `custom` preset's `waypoints` are clamped into the EMU window + capped;
 *    an EMPTY custom waypoint list (after clamping) is unrenderable → `null`.
 *
 * Pure: never mutates its input.
 */
export function normalizeMotionPath(raw: unknown): MotionPath | null {
  const record = asRecord(raw);
  if (!record) return null;
  const preset = normalizeMotionPathPreset(record.preset);
  if (preset === null) return null;
  if (preset !== "custom") {
    // The preset's canonical geometry is authoritative — fresh copies so the
    // caller can never mutate the frozen template.
    const waypoints = PRESET_WAYPOINTS[preset].map((p) => ({ x: p.x, y: p.y }));
    return { preset, waypoints };
  }
  const waypoints = clampWaypoints(record.waypoints);
  // A custom path with no usable waypoint is meaningless (the block would just
  // sit still) — drop it so the timeline + renderer never carry a no-op path.
  if (waypoints.length === 0) return null;
  return { preset, waypoints };
}

/**
 * Convenience builder: the canonical {@link MotionPath} for a named preset, or
 * `null` for `custom` (which has no canonical geometry) / an unknown name. Used
 * by the editor's motion-path picker so the persisted value always round-trips
 * through {@link normalizeMotionPath}.
 */
export function presetMotionPath(value: unknown): MotionPath | null {
  const preset = normalizeMotionPathPreset(value);
  if (preset === null || preset === "custom") return null;
  return normalizeMotionPath({ preset });
}

// ── CSS keyframe generation ──────────────────────────────────────────────────

/**
 * Format a number to at most 4 decimal places WITHOUT trailing zeros, so the
 * generated CSS is stable + compact (e.g. `33.3333`, `0`, `-66.6667`). Pure +
 * deterministic — identical input always yields identical output.
 */
function fmt(value: number): string {
  // toFixed(4) is deterministic across engines; strip trailing zeros + dot.
  const fixed = value.toFixed(4);
  return fixed.replace(/\.?0+$/, "");
}

/**
 * The container-query length units the keyframe translate uses. `cqw` = 1% of
 * the query container's (the present slide's) width; `cqh` = 1% of its height.
 * Moving by `(dxEmu / slideWidthEmu) * 100` cqw therefore moves the block by the
 * SLIDE-relative fraction, matching the percentage geometry the renderer uses.
 */
const X_UNIT = "cqw";
const Y_UNIT = "cqh";

/** The final slide-relative translate value held after a motion path completes. */
export function motionPathEndTranslate(
  path: MotionPath,
  slideWidthEmu: number,
  slideHeightEmu: number,
): string | null {
  const endpoint = path.waypoints.at(-1);
  if (
    !endpoint ||
    !Number.isFinite(slideWidthEmu) ||
    !Number.isFinite(slideHeightEmu) ||
    slideWidthEmu <= 0 ||
    slideHeightEmu <= 0
  ) {
    return null;
  }
  return `${fmt((endpoint.x / slideWidthEmu) * 100)}${X_UNIT} ${fmt(
    (endpoint.y / slideHeightEmu) * 100,
  )}${Y_UNIT}`;
}

/**
 * The CSS keyframe BODY (the `0% {…} 50% {…} 100% {…}` text, without the
 * `@keyframes <name> { … }` wrapper) for a motion path, given the slide's EMU
 * dimensions. Each waypoint's EMU offset is converted to a slide-relative
 * percentage and emitted as the independent `translate: x cqw y cqh` property
 * (NOT `transform: translate()`, which would clobber a transform-based entrance —
 * see the module header) so the block travels the SLIDE-relative trajectory
 * regardless of its own size while composing with entrance/emphasis/rotation.
 *
 * The frames are evenly spaced across `[0%, 100%]`: frame 0 is the origin
 * (`translate(0,0)` — the block's resting spot), and the N waypoints land at
 * `i/N` of the timeline. A single-waypoint path is a simple A→B slide.
 *
 * Deterministic: no Date / no random; identical (path, dims) → identical CSS.
 * Returns `""` for an empty path or non-positive dims (unrenderable).
 */
export function motionPathKeyframeBody(
  path: MotionPath,
  slideWidthEmu: number,
  slideHeightEmu: number,
): string {
  if (
    path.waypoints.length === 0 ||
    !Number.isFinite(slideWidthEmu) ||
    !Number.isFinite(slideHeightEmu) ||
    slideWidthEmu <= 0 ||
    slideHeightEmu <= 0
  ) {
    return "";
  }
  // The origin (resting spot) is the implicit first frame at 0%.
  const points: MotionWaypoint[] = [{ x: 0, y: 0 }, ...path.waypoints];
  const last = points.length - 1;
  const frames: string[] = [];
  for (let i = 0; i < points.length; i += 1) {
    const pct = last === 0 ? 0 : (i / last) * 100;
    const x = (points[i].x / slideWidthEmu) * 100;
    const y = (points[i].y / slideHeightEmu) * 100;
    frames.push(
      `${fmt(pct)}% { translate: ${fmt(x)}${X_UNIT} ${fmt(y)}${Y_UNIT}; }`,
    );
  }
  return frames.join(" ");
}

/**
 * A STABLE, deterministic keyframe NAME for a motion path at a given slide size.
 * Derived purely from the path geometry + dims (no Date / no random) so a given
 * path always produces the same `@keyframes` identifier — two steps with the same
 * path share one keyframe, and SSR + client render emit the identical name.
 * Always a valid CSS ident (`mpk-` + hex).
 *
 * Uses TWO independent FNV-1a passes (distinct offset basis + prime) folded into
 * one ident, i.e. an effective ~64-bit hash. A single 32-bit FNV-1a collides on
 * realistic custom-path inputs (~14 collisions across ~270k distinct paths) —
 * since the renderer keys its keyframe map by this name, two colliding paths on
 * one slide would have the second's CSS overwrite the first's, silently giving
 * both blocks the wrong trajectory. The second pass makes such a collision
 * ~2^-64 — negligible even for hand-authored / imported custom paths.
 */
export function motionPathKeyframeName(
  path: MotionPath,
  slideWidthEmu: number,
  slideHeightEmu: number,
): string {
  const seed = `${path.preset}|${slideWidthEmu}x${slideHeightEmu}|${path.waypoints
    .map((p) => `${p.x},${p.y}`)
    .join(";")}`;
  // Two independent FNV-1a passes (different basis + prime) → ~64-bit ident.
  let h1 = 0x811c9dc5;
  let h2 = 0x811c9dc5 ^ 0x9e3779b9;
  for (let i = 0; i < seed.length; i += 1) {
    const c = seed.charCodeAt(i);
    h1 ^= c;
    h1 = Math.imul(h1, 0x01000193);
    h2 ^= c;
    h2 = Math.imul(h2, 0x85ebca77);
  }
  // >>> 0 → unsigned; base-36 keeps each half short. The two halves together
  // make collisions on distinct paths vanishingly unlikely.
  return `mpk-${(h1 >>> 0).toString(36)}-${(h2 >>> 0).toString(36)}`;
}

/**
 * The full `@keyframes <name> { … }` CSS for a motion path, ready to inject into
 * a present-mode `<style>` tag. Returns `""` for an unrenderable path (empty /
 * bad dims). Pure + deterministic. The renderer collects these (deduped by name)
 * for every motion-path step on the current slide and injects them once.
 */
export function motionPathKeyframeCss(
  path: MotionPath,
  slideWidthEmu: number,
  slideHeightEmu: number,
): string {
  const body = motionPathKeyframeBody(path, slideWidthEmu, slideHeightEmu);
  if (!body) return "";
  const name = motionPathKeyframeName(path, slideWidthEmu, slideHeightEmu);
  return `@keyframes ${name} { ${body} }`;
}
