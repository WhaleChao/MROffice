/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Format Painter (PowerPoint 格式刷) for the native presentation editor — copy
 * ONE block's visual STYLE and paint it onto another block.
 *
 * This module is the PURE core: it knows how to extract a kind-agnostic style
 * SNAPSHOT from a source block and how to produce a style-only PATCH for a
 * target block of any kind. It contains no React, no DOM, and no time/random
 * source so it can be unit-tested in isolation and reasoned about deterministically.
 *
 * THE INVARIANT (what the painter is allowed to touch): only VISUAL STYLE
 * transfers. Geometry (`left`/`top`/`width`/`height`/`rotation`), identity
 * (`id`/`type`/`kind`), text CONTENT (`text`/`paragraphs`), grouping (`groupId`),
 * connector anchors (`startShapeId`/`endShapeId` and the `flipH`/`flipV` they
 * derive, plus the legacy `lineSlope`), media sources
 * (`src`/`url`/`alt`/`caption`), hyperlinks (`href`) and entrance animations
 * (`animation`) are NEVER part of a snapshot, so a paint can never move, rename,
 * regroup, reroute, relink or re-content a block — it only repaints it.
 *
 * GRACEFUL CROSS-KIND DEGRADATION: a snapshot is a UNION of every style family.
 * When painting onto a target, only the props applicable to the target's kind
 * are emitted; inapplicable props are silently dropped. Painting a shape's fill
 * onto an image keeps the image's geometry + src and applies only the
 * image-applicable visual props (shadow); painting a text style onto a shape
 * keeps the shape and applies only shape-applicable props (none from a text
 * snapshot) — never a crash, never a corrupt block.
 */

import {
  mergeRuns,
  normalizeRunStyle,
  type RunStyle,
  type TextRun,
} from "./presentation-point-runs";
import type {
  ParagraphAlign,
  ParagraphBullet,
} from "./presentation-paragraphs";
import type { OfficeShapeShadow } from "./office-shape-paint";

/** Shape Format > Shape Effects preset a shape can carry beside its shadow. */
type ShapeEffectPreset = "reflection" | "glow" | "softEdge" | "bevel" | "threeD";

/**
 * The minimal block shape this module reasons about. A presentation block is an
 * object with a `type` discriminator plus arbitrary additive style fields (the
 * editor model carries an index signature). We read style props off it loosely
 * and re-validate nothing here — the editor's existing style setters normalize
 * at the commit edge.
 */
export interface PaintableBlock {
  type?: unknown;
  kind?: unknown;
  style?: unknown;
  runs?: unknown;
  [key: string]: unknown;
}

/** A presentation block's per-block text style (PowerPoint parity). */
export interface TextStyleSnapshot {
  fontSize?: number;
  /** The box's own font: a font chosen before typing is stated here, not on a run (PPT-B342). */
  fontFamily?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  align?: "left" | "center" | "right";
}

/**
 * Character-format patch captured by Format Painter. `null` is intentional:
 * it means the selected source text uniformly inherits/defaults that field, so
 * painting must CLEAR an explicit target override. A missing key means the
 * source selection was mixed for that field and the target must retain it.
 */
export type FormatPainterRunStylePatch = {
  [K in keyof RunStyle]?: RunStyle[K] | null;
};

/** Paragraph formatting copied only from complete paragraph selections. */
export interface FormatPainterParagraphStylePatch {
  level?: number;
  bullet?: ParagraphBullet;
  align?: ParagraphAlign;
  /** `null` means the source uses the default and clears a target override. */
  lineSpacing?: number | null;
  spaceBefore?: number | null;
  spaceAfter?: number | null;
}

/**
 * A deep, kind-agnostic style SNAPSHOT captured at arm time. Every family is
 * optional; only the families the source block actually carried are present.
 * This object is a deep copy — mutating the source block after capture can never
 * change a pending paint.
 */
export interface FormatPainterSnapshot {
  /** Absent/`object` keeps the established whole-object painter semantics. */
  scope?: "object" | "text-selection";
  /** Shape visual style: fill / gradient / outline / shadow. */
  shape?: {
    fill?: string;
    fillTo?: string;
    gradientAngle?: number;
    stroke?: string;
    strokeWidth?: number;
    strokeStyle?: "solid" | "dashed" | "dotted";
    shadow?: boolean;
    /** The shadow's own offset, blur and colour, when it states them. */
    shapeShadow?: OfficeShapeShadow;
    /** The effects preset; "none" when the source shape has none. */
    effect?: ShapeEffectPreset | "none";
  };
  /** Image visual style: shadow / border / corner radius. */
  image?: {
    shadow?: boolean;
    borderColor?: string;
    borderWidth?: number;
    borderStyle?: "none" | "solid" | "dashed" | "dotted";
    /** The frame's compound line and corners (`borderCompound` / `borderJoin`). */
    borderCompound?: string;
    borderJoin?: string;
    cornerRadius?: number;
  };
  /** Table visual style: thin borders + header-row emphasis. */
  table?: FormatPainterTableLook;
  /** Block-level text style (font size / colour / bold / italic / underline / align). */
  text?: TextStyleSnapshot;
  /**
   * A uniform per-run style derived from the source's run formatting. The
   * painter applies this as a whole-block run style (mirroring the editor's
   * no-sub-selection run-style path). Absent when the source carried no usable
   * run formatting.
   */
  runStyle?: FormatPainterRunStylePatch;
  /** Paragraph-local formatting, absent for partial text selections. */
  paragraph?: FormatPainterParagraphStylePatch;
}

/** Keys of {@link FormatPainterSnapshot.shape}, used to copy only present fields. */
const SHAPE_STYLE_KEYS = [
  "fill",
  "fillTo",
  "gradientAngle",
  "stroke",
  "strokeWidth",
  "strokeStyle",
  "shadow",
] as const;

/** Keys of {@link FormatPainterSnapshot.image}. */
// The picture border is its colour, width AND dash (PPT-B211).
// Its compound lines and corners with it (readings row 856).
const IMAGE_STYLE_KEYS = [
  "shadow",
  "borderColor",
  "borderWidth",
  "borderStyle",
  "borderCompound",
  "borderJoin",
  "cornerRadius",
] as const;

/** The Table Style Options (and the thin-grid switch), each stated on or off. */
export const TABLE_LOOK_OPTION_KEYS = [
  "borders",
  "headerRow",
  "totalRow",
  "firstColumn",
  "lastColumn",
  "bandedRows",
  "bandedColumns",
] as const;

/**
 * A table's whole look as the painter carries it: its style (`a:tableStyleId`
 * and the resolved look it draws) and every Table Style Option, OFF ones too,
 * so a source with its header row turned off turns the target's off. It used
 * to carry only `borders` / `headerRow`, and only when on (PPT-B154).
 */
export type FormatPainterTableLook = {
  [K in (typeof TABLE_LOOK_OPTION_KEYS)[number]]?: boolean;
} & {
  /** The style, or null for a table that states none. */
  tableStyleId?: string | null;
  tableStyleLook?: Record<string, unknown> | null;
};

function captureTableLook(block: Record<string, unknown>): FormatPainterTableLook {
  const look: FormatPainterTableLook = {};
  for (const key of TABLE_LOOK_OPTION_KEYS) look[key] = block[key] === true;
  look.tableStyleId =
    typeof block.tableStyleId === "string" && block.tableStyleId ? block.tableStyleId : null;
  look.tableStyleLook =
    block.tableStyleLook && typeof block.tableStyleLook === "object"
      ? (JSON.parse(JSON.stringify(block.tableStyleLook)) as Record<string, unknown>)
      : null;
  return look;
}

/**
 * The target table with the source's look: its style and options, and -- as
 * applying a style from the gallery does -- no direct cell shading left over
 * a style that changed, so the new look shows. Text, borders drawn on cells,
 * merges and sizes stay the target's own.
 */
export function tableWithPaintedLook<T extends Record<string, unknown>>(
  table: T,
  look: FormatPainterTableLook,
): T {
  const next = { ...table } as Record<string, unknown>;
  for (const key of TABLE_LOOK_OPTION_KEYS) {
    if (look[key] === true) next[key] = true;
    else if (look[key] === false) delete next[key];
  }
  const styleChanged =
    ("tableStyleId" in look &&
      (look.tableStyleId ?? null) !==
        (typeof table.tableStyleId === "string" ? table.tableStyleId : null)) ||
    ("tableStyleLook" in look &&
      JSON.stringify(look.tableStyleLook ?? null) !==
        JSON.stringify(table.tableStyleLook ?? null));
  if ("tableStyleId" in look) {
    if (look.tableStyleId) next.tableStyleId = look.tableStyleId;
    else delete next.tableStyleId;
  }
  if ("tableStyleLook" in look) {
    if (look.tableStyleLook) next.tableStyleLook = JSON.parse(JSON.stringify(look.tableStyleLook));
    else delete next.tableStyleLook;
  }
  if (styleChanged && Array.isArray(table.rows)) {
    next.rows = (table.rows as unknown[][]).map((row) =>
      Array.isArray(row) &&
      row.some((cell) => cell && typeof cell === "object" && "fill" in (cell as object))
        ? row.map((cell) => {
            if (!cell || typeof cell !== "object" || !("fill" in (cell as object))) return cell;
            const { fill: _fill, ...rest } = cell as Record<string, unknown>;
            return rest;
          })
        : row,
    );
  }
  return next as T;
}

/** Keys of {@link TextStyleSnapshot}. */
const TEXT_STYLE_KEYS = [
  "fontSize",
  "fontFamily",
  "color",
  "bold",
  "italic",
  "underline",
  "align",
] as const;

/** Keys of {@link RunStyle}. */
const RUN_STYLE_KEYS = [
  "bold",
  "italic",
  "underline",
  "strike",
  "superscript",
  "subscript",
  "color",
  "highlight",
  "fontSize",
  "fontFamily",
  "textOutline",
  "textOutlineWidthPt",
  "textOutlineJoin",
  "textOutlineMiterLimit",
  "textOutlineCap",
  "textOutlineDash",
] as const;

/** The runtime value type each style key MUST carry to be copied. */
type StyleKeyType = "string" | "number" | "boolean";

/**
 * Per-key expected runtime type. A value whose type doesn't match its key is
 * dropped (e.g. a string in a numeric `gradientAngle`, a non-finite number in
 * `strokeWidth`), so a corrupt source field never poisons a snapshot.
 */
const STYLE_KEY_TYPES: Record<string, StyleKeyType> = {
  // shape
  fill: "string",
  fillTo: "string",
  gradientAngle: "number",
  stroke: "string",
  strokeWidth: "number",
  strokeStyle: "string",
  shadow: "boolean",
  // image
  borderColor: "string",
  borderWidth: "number",
  borderStyle: "string",
  borderCompound: "string",
  borderJoin: "string",
  cornerRadius: "number",
  // table
  borders: "boolean",
  headerRow: "boolean",
  // text style
  fontSize: "number",
  color: "string",
  bold: "boolean",
  italic: "boolean",
  underline: "boolean",
  align: "string",
  // run style (extra)
  strike: "boolean",
  superscript: "boolean",
  subscript: "boolean",
  highlight: "string",
  fontFamily: "string",
  textOutline: "string",
  textOutlineWidthPt: "number",
  textOutlineJoin: "string",
  textOutlineMiterLimit: "number",
  textOutlineCap: "string",
  textOutlineDash: "string",
};

function valueMatchesType(value: unknown, type: StyleKeyType): boolean {
  if (type === "string") return typeof value === "string";
  if (type === "boolean") return typeof value === "boolean";
  return typeof value === "number" && Number.isFinite(value);
}

/**
 * Copy only the present (defined, non-null, TYPE-VALID) keys from `source` into
 * a fresh object. Each key is validated against {@link STYLE_KEY_TYPES} so a
 * wrong-typed source field is dropped. Returns `undefined` when nothing usable
 * was found so an empty family is dropped rather than stored as `{}`.
 */
function pickStyle<K extends string>(
  source: Record<string, unknown>,
  keys: readonly K[],
): Record<K, unknown> | undefined {
  const out = {} as Record<K, unknown>;
  let has = false;
  for (const key of keys) {
    const value = source[key];
    if (value === undefined || value === null) continue;
    const expected = STYLE_KEY_TYPES[key];
    if (!expected || !valueMatchesType(value, expected)) continue;
    out[key] = value;
    has = true;
  }
  return has ? out : undefined;
}

/** The normalized block kind discriminator for painting decisions. */
export type PaintableKind = "shape" | "image" | "table" | "text" | "chart" | "embed" | "other";

/** Resolve a block's paint kind from its `type` discriminator. */
export function paintableKind(block: PaintableBlock | null | undefined): PaintableKind {
  const type = block?.type;
  if (type === "shape") return "shape";
  if (type === "image") return "image";
  if (type === "table") return "table";
  if (type === "text") return "text";
  if (type === "chart") return "chart";
  if (type === "embed") return "embed";
  return "other";
}

/**
 * Derive a single uniform {@link RunStyle} from a block's runs. The editor's
 * whole-block run-style path applies one style to the entire text, so for the
 * painter we capture the FIRST run's style as the representative uniform style
 * (a deep copy of its present, valid fields). Returns `undefined` when the block
 * has no runs or the first run carries no usable formatting.
 */
function captureRunStyle(block: PaintableBlock): RunStyle | undefined {
  const runs = block.runs;
  if (!Array.isArray(runs) || runs.length === 0) return undefined;
  const first = runs[0] as TextRun | undefined;
  const style = first?.style;
  if (!style || typeof style !== "object") return undefined;
  const picked = pickStyle(style as Record<string, unknown>, RUN_STYLE_KEYS);
  return picked as RunStyle | undefined;
}

/** One newline-aware selected run span supplied by the Presentation adapter. */
export interface FormatPainterRunRange {
  runs: readonly TextRun[];
  from: number;
  to: number;
  /**
   * Present only when this range includes the complete paragraph and its
   * paragraph mark. Missing means character-only Format Painter semantics.
   */
  paragraph?: FormatPainterParagraphStylePatch;
}

const PARAGRAPH_STYLE_KEYS = [
  "level",
  "bullet",
  "align",
  "lineSpacing",
  "spaceBefore",
  "spaceAfter",
] as const;

function validParagraphStyleValue(
  key: (typeof PARAGRAPH_STYLE_KEYS)[number],
  value: unknown,
): boolean {
  if (key === "level") {
    return Number.isInteger(value) && Number(value) >= 0 && Number(value) <= 4;
  }
  if (key === "bullet") {
    return value === "none" || value === "bullet" || value === "number";
  }
  if (key === "align") {
    return (
      value === "left" ||
      value === "center" ||
      value === "right" ||
      value === "justify" ||
      value === "distribute"
    );
  }
  return (
    value === null ||
    (typeof value === "number" && Number.isFinite(value) && value >= 0)
  );
}

/** Capture only fields that are present, valid, and uniform on every paragraph. */
function captureUniformParagraphStyle(
  paragraphs: readonly FormatPainterParagraphStylePatch[],
): FormatPainterParagraphStylePatch | undefined {
  if (paragraphs.length === 0) return undefined;
  const patch: FormatPainterParagraphStylePatch = {};
  for (const key of PARAGRAPH_STYLE_KEYS) {
    if (
      !paragraphs.every((paragraph) =>
        Object.prototype.hasOwnProperty.call(paragraph, key),
      )
    ) {
      continue;
    }
    const first = paragraphs[0]?.[key];
    if (
      !validParagraphStyleValue(key, first) ||
      !paragraphs.every((paragraph) => Object.is(paragraph[key], first))
    ) {
      continue;
    }
    (patch as Record<string, unknown>)[key] = first;
  }
  return Object.keys(patch).length > 0 ? patch : undefined;
}

/**
 * Capture character formatting from one or more selected run spans.
 *
 * Each field is evaluated independently across every covered run:
 * - one shared value -> copy it;
 * - uniformly absent -> capture `null` so the target override is cleared;
 * - disagreement -> omit the field so a mixed source never invents a value.
 *
 * The adapter materializes authenticated block-level inheritance before
 * calling this helper, so an absent value here is a genuine default rather
 * than an unresolved theme/block value. Separator-only or stale ranges yield
 * `null` and therefore cannot arm an inert painter.
 */
export function captureTextSelectionFormatSnapshot(
  ranges: readonly FormatPainterRunRange[],
): FormatPainterSnapshot | null {
  const styles: Array<RunStyle | undefined> = [];
  const paragraphStyles: FormatPainterParagraphStylePatch[] = [];
  let paragraphSelectionComplete = ranges.length > 0;
  for (const range of ranges) {
    const runs = mergeRuns(range.runs);
    const total = runs.reduce((length, run) => length + run.text.length, 0);
    const from = Math.min(Math.max(Math.round(range.from), 0), total);
    const to = Math.min(Math.max(Math.round(range.to), 0), total);
    if (from >= to) continue;
    if (range.paragraph) paragraphStyles.push(range.paragraph);
    else paragraphSelectionComplete = false;
    let cursor = 0;
    for (const run of runs) {
      const start = cursor;
      const end = cursor + run.text.length;
      cursor = end;
      if (end <= from || start >= to) continue;
      styles.push(normalizeRunStyle(run.style));
    }
  }
  const paragraph = paragraphSelectionComplete
    ? captureUniformParagraphStyle(paragraphStyles)
    : undefined;
  if (styles.length === 0 && !paragraph) return null;

  const patch: FormatPainterRunStylePatch = {};
  for (const key of RUN_STYLE_KEYS) {
    const first = styles[0]?.[key];
    if (!styles.every((style) => Object.is(style?.[key], first))) continue;
    (
      patch as Partial<
        Record<keyof RunStyle, RunStyle[keyof RunStyle] | null>
      >
    )[key] = first ?? null;
  }
  if (Object.keys(patch).length === 0 && !paragraph) return null;
  return {
    scope: "text-selection",
    ...(Object.keys(patch).length > 0 ? { runStyle: patch } : {}),
    ...(paragraph ? { paragraph } : {}),
  };
}

/**
 * Capture a deep style snapshot from the SOURCE block. Reads only style families
 * (never geometry / identity / content / grouping / anchors / media / href /
 * animation). The result is a fresh, deeply-copied object — every nested style
 * object is a new object containing copied primitives — so the source block can
 * be freely edited afterward without changing the snapshot.
 *
 * Returns `null` when the block is missing/invalid or carries NO copyable style
 * at all (so the caller can decline to arm with an empty clipboard).
 */
export function captureFormatSnapshot(
  block: PaintableBlock | null | undefined,
): FormatPainterSnapshot | null {
  if (!block || typeof block !== "object") return null;
  const snapshot: FormatPainterSnapshot = {};

  // Shape visual style lives directly on the (shape) block. We read it from ANY
  // block kind defensively — non-shape blocks simply won't carry these fields,
  // so `pickStyle` returns undefined and the family is dropped.
  const shape = pickStyle(block as Record<string, unknown>, SHAPE_STYLE_KEYS);
  if (shape) snapshot.shape = shape as FormatPainterSnapshot["shape"];
  if (
    block.type === "shape" &&
    (shape || (block as Record<string, unknown>).shapeShadow || (block as Record<string, unknown>).effect)
  ) {
    // A shape's effects are copied whether or not it has any, as PowerPoint's
    // painter copies them: a shape with no shadow paints the target's shadow
    // away. Only the present keys were copied, so the target kept its own
    // shadow under the source's fill (PPT-B200).
    const source = block as Record<string, unknown>;
    const shapeShadow = source.shapeShadow;
    const effect = source.effect;
    snapshot.shape = {
      ...snapshot.shape,
      shadow: source.shadow === true,
      ...(shapeShadow && typeof shapeShadow === "object"
        ? { shapeShadow: JSON.parse(JSON.stringify(shapeShadow)) as OfficeShapeShadow }
        : {}),
      effect: typeof effect === "string" ? (effect as ShapeEffectPreset) : "none",
    };
  }

  const image = pickStyle(block as Record<string, unknown>, IMAGE_STYLE_KEYS);
  if (image) snapshot.image = image as FormatPainterSnapshot["image"];
  if (block.type === "image" && image) {
    // The same for a picture's shadow (PPT-B200).
    snapshot.image = {
      ...snapshot.image,
      shadow: (block as Record<string, unknown>).shadow === true,
    };
  }

  if (block.type === "table") {
    snapshot.table = captureTableLook(block as Record<string, unknown>);
  }

  const styleObj = block.style;
  if (styleObj && typeof styleObj === "object") {
    const text = pickStyle(styleObj as Record<string, unknown>, TEXT_STYLE_KEYS);
    if (text) snapshot.text = text as TextStyleSnapshot;
  }

  const runStyle = captureRunStyle(block);
  if (runStyle) snapshot.runStyle = runStyle;

  // An entirely-empty snapshot means there was nothing to copy: decline to arm.
  if (
    !snapshot.shape &&
    !snapshot.image &&
    !snapshot.table &&
    !snapshot.text &&
    !snapshot.runStyle
  ) {
    return null;
  }
  return snapshot;
}

/**
 * The result of resolving a snapshot against a target block: the style-only
 * patches to apply, partitioned by setter. Each is `undefined` when the snapshot
 * carries nothing applicable to the target's kind — the caller skips that setter
 * entirely (so painting onto an incompatible kind is a graceful no-op rather
 * than a crash).
 */
export interface FormatPaintPlan {
  /** Shape-style patch for `updateShapeBlock` (only when target is a shape). */
  shapePatch?: {
    fill?: string;
    fillTo?: string;
    gradientAngle?: number;
    stroke?: string;
    strokeWidth?: number;
    strokeStyle?: "solid" | "dashed" | "dotted";
    shadow?: boolean;
    shapeShadow?: OfficeShapeShadow;
    /** Undefined clears the target's effects preset. */
    effect?: ShapeEffectPreset;
  };
  /** Image-style patch (only when target is an image). */
  imagePatch?: {
    shadow?: boolean;
    borderColor?: string;
    borderWidth?: number;
    borderStyle?: "none" | "solid" | "dashed" | "dotted";
    /** The frame's compound line and corners (`borderCompound` / `borderJoin`). */
    borderCompound?: string;
    borderJoin?: string;
    cornerRadius?: number;
  };
  /** Table look (only when target is a table): see {@link tableWithPaintedLook}. */
  tablePatch?: FormatPainterTableLook;
  /** Block text-style patch (only when target is a text/textbox block). */
  textPatch?: TextStyleSnapshot;
  /** Whole-block run-style patch (only when target is a text/textbox block). */
  runStylePatch?: FormatPainterRunStylePatch;
  /** Paragraph patch for a complete-paragraph text selection. */
  paragraphPatch?: FormatPainterParagraphStylePatch;
}

function cloneRecord<T extends object>(value: T): T {
  return { ...value };
}

/**
 * Resolve a snapshot into a kind-appropriate {@link FormatPaintPlan} for the
 * TARGET block. Only families applicable to the target's kind are returned;
 * everything else is dropped (graceful cross-kind degradation). The returned
 * patch objects are fresh copies (no aliasing back to the snapshot), so applying
 * one can never mutate the stored clipboard.
 *
 * - shape target  → shapePatch (fill/gradient/outline/shadow) + textPatch and
 *   runStylePatch for the words in the shape, when it has any
 * - image target  → imagePatch (shadow/border/cornerRadius)
 * - table target  → tablePatch (style + every Table Style Option)
 * - text target   → textPatch (block text style) + runStylePatch (uniform run style)
 * - chart/embed/other → empty plan (nothing applicable; paint is a no-op)
 *
 * Returns an empty plan (all fields undefined) — never null — so the caller can
 * always treat a paint as "apply whatever applies", and `isEmptyPlan` reports
 * when nothing was applicable.
 */
export function planFormatPaint(
  snapshot: FormatPainterSnapshot | null | undefined,
  target: PaintableBlock | null | undefined,
): FormatPaintPlan {
  const plan: FormatPaintPlan = {};
  if (!snapshot || !target) return plan;
  const kind = paintableKind(target);

  if (kind === "shape" && snapshot.shape) {
    const { effect, shapeShadow, ...rest } = snapshot.shape;
    plan.shapePatch = {
      ...rest,
      ...(shapeShadow ? { shapeShadow: JSON.parse(JSON.stringify(shapeShadow)) as OfficeShapeShadow } : {}),
      // Stated, so the target's own preset is cleared by a source without one.
      ...(effect !== undefined ? { effect: effect === "none" ? undefined : effect } : {}),
    };
  } else if (kind === "image" && snapshot.image) {
    plan.imagePatch = cloneRecord(snapshot.image);
  } else if (kind === "table" && snapshot.table) {
    plan.tablePatch = {
      ...snapshot.table,
      ...(snapshot.table.tableStyleLook
        ? { tableStyleLook: JSON.parse(JSON.stringify(snapshot.table.tableStyleLook)) }
        : {}),
    };
  }
  // A shape's own words take the source's text formatting as a text box's do
  // (PPT-B051): the painter copies a shape's look AND the formatting of the
  // words in it, keeping the target's words. A shape without words has
  // nothing for it to format.
  const shapeHasText =
    kind === "shape" && typeof target.text === "string" && target.text !== "";
  if (kind === "text" || shapeHasText) {
    if (snapshot.text) plan.textPatch = cloneRecord(snapshot.text);
    if (snapshot.runStyle) plan.runStylePatch = cloneRecord(snapshot.runStyle);
    if (kind === "text" && snapshot.paragraph) {
      plan.paragraphPatch = cloneRecord(snapshot.paragraph);
    }
  }
  return plan;
}

/** True when a paint plan carries no applicable patch (a no-op paint). */
export function isEmptyPlan(plan: FormatPaintPlan): boolean {
  return (
    !plan.shapePatch &&
    !plan.imagePatch &&
    !plan.tablePatch &&
    !plan.textPatch &&
    !plan.runStylePatch &&
    !plan.paragraphPatch
  );
}

/** The armed mode of the format painter. */
export type FormatPainterMode = "single" | "sticky";

/** The serializable armed state of the format painter (held in editor state). */
export interface PresentationFormatPainterState {
  snapshot: FormatPainterSnapshot;
  mode: FormatPainterMode;
  /** The id of the block the style was copied FROM (for the armed visual cue). */
  sourceId: string;
}

/**
 * Given the current painter state after a single paint, decide the next state.
 * Single-use disarms (returns null); sticky stays armed (returns the same
 * state). Pure — the state machine lives here, the editor just stores the result.
 */
export function nextStateAfterPaint(
  state: PresentationFormatPainterState,
): PresentationFormatPainterState | null {
  return state.mode === "sticky" ? state : null;
}
