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
import type { PresentationLevelStyle } from "./presentation-level-styles";
import { PRESENTATION_LEVEL_STYLE_KEYS, type PresentationLevelStyleKey } from "./presentation-runs";

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
  /** That font is the theme's heading or body font (a placeholder's, API readings row 843). */
  themeFontRole?: "heading" | "body";
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
  /**
   * The source's text frame (`a:bodyPr`), a whole-object paint only: Format
   * Painter replaces the target's whole frame with the source's (truth
   * c62bodypr / c62bodypr2: anchor, wrap, margins, direction, columns,
   * autofit, WordArt Transform -- what the source does not state is taken off
   * the target too). See {@link FormatPainterBodySnapshot}.
   */
  body?: FormatPainterBodySnapshot;
}

/**
 * A text frame's settings as the page states them, block-level and on its
 * style, and which imported shape they came from (`from`, its slide and shape
 * ids): the save then copies that shape's whole `a:bodyPr` -- with what the
 * page does not model, its rotation, column gap and order, overflow -- when
 * the original still holds it as `block` / `style` say (readings row 829,
 * `source.paintedBodyFrom`).
 */
export interface FormatPainterBodySnapshot {
  block: Record<string, unknown>;
  style: Record<string, unknown>;
  from?: { slideId: string; shapeId: string };
}

/** A text frame's block-level fields (`a:bodyPr`): margins, autofit, columns, WordArt Transform, keep text flat. */
export const FORMAT_PAINTER_BODY_BLOCK_KEYS = [
  "textInsets",
  "autofit",
  "columnCount",
  "columnSpacing",
  "textWarp",
  "keepTextFlat",
] as const;

/** A text frame's fields on the block's style (`a:bodyPr`): anchor, direction, wrap, overflow, 3-D text. */
export const FORMAT_PAINTER_BODY_STYLE_KEYS = [
  "verticalAlign",
  "anchorCenter",
  "writingMode",
  "textWrap",
  "verticalOverflow",
  "spaceFirstLastPara",
  "textWarp",
  "textShape3d",
  "autofit",
] as const;

function hasTextFrame(block: PaintableBlock): boolean {
  return block.type === "text" || (block.type === "shape" && typeof (block as { text?: unknown }).text === "string");
}

function captureBody(block: PaintableBlock): FormatPainterBodySnapshot | undefined {
  if (!hasTextFrame(block)) return undefined;
  const record = block as Record<string, unknown>;
  const style = record.style && typeof record.style === "object" && !Array.isArray(record.style)
    ? (record.style as Record<string, unknown>)
    : {};
  const pick = (from: Record<string, unknown>, keys: readonly string[]) =>
    Object.fromEntries(
      keys.filter((key) => from[key] !== undefined).map((key) => [key, JSON.parse(JSON.stringify(from[key]))]),
    );
  const source = record.source && typeof record.source === "object" ? (record.source as Record<string, unknown>) : null;
  const from =
    source && typeof source.slideId === "string" && typeof source.shapeId === "string"
      ? { slideId: source.slideId, shapeId: source.shapeId }
      : undefined;
  return {
    block: pick(record, FORMAT_PAINTER_BODY_BLOCK_KEYS),
    style: pick(style, FORMAT_PAINTER_BODY_STYLE_KEYS),
    ...(from ? { from } : {}),
  };
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
  "themeFontRole",
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
  themeFontRole: "string",
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
  // The first character's run: the block's flat runs, else its first row's
  // (a shape label held as rows states no flat runs).
  const rows = (block as { paragraphs?: unknown }).paragraphs;
  const firstRow = Array.isArray(rows) ? (rows[0] as { runs?: unknown } | undefined) : undefined;
  const runs = Array.isArray(block.runs) && block.runs.length > 0 ? block.runs : firstRow?.runs;
  if (!Array.isArray(runs) || runs.length === 0) return undefined;
  const first = runs[0] as TextRun | undefined;
  const style = first?.style;
  if (!style || typeof style !== "object") return undefined;
  const picked = pickStyle(style as Record<string, unknown>, RUN_STYLE_KEYS) as RunStyle | undefined;
  // A face that follows the theme's heading or body font goes as that
  // reference, so the painted words follow the theme as the source's do.
  const role = (style as { themeFontRole?: unknown }).themeFontRole;
  if (picked?.fontFamily && (role === "heading" || role === "body")) {
    return { ...picked, themeFontRole: role } as RunStyle;
  }
  return picked;
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
    // A placeholder's face from its layout (`source.layoutStyle`, API readings
    // row 843) says there whether it follows the theme's heading or body font.
    const layoutStyle = (block as { source?: { layoutStyle?: Record<string, unknown> } }).source?.layoutStyle;
    if (
      text &&
      text.themeFontRole === undefined &&
      layoutStyle &&
      typeof text.fontFamily === "string" &&
      text.fontFamily === layoutStyle.fontFamily &&
      (layoutStyle.themeFontRole === "heading" || layoutStyle.themeFontRole === "body")
    ) {
      (text as Record<string, unknown>).themeFontRole = layoutStyle.themeFontRole;
    }
    if (text && !(typeof text.fontFamily === "string" && (text.themeFontRole === "heading" || text.themeFontRole === "body"))) {
      // A role names a theme font only beside the face it resolved to.
      delete (text as Record<string, unknown>).themeFontRole;
    }
    if (text && Object.keys(text).length > 0) snapshot.text = text as TextStyleSnapshot;
  }

  const runStyle = captureRunStyle(block);
  if (runStyle) snapshot.runStyle = runStyle;

  const body = captureBody(block);
  if (body) snapshot.body = body;

  // An entirely-empty snapshot means there was nothing to copy: decline to arm.
  if (
    !snapshot.shape &&
    !snapshot.image &&
    !snapshot.table &&
    !snapshot.text &&
    !snapshot.runStyle &&
    // A frame the source states nothing of, on a box the page made, is no
    // format to arm with (an imported one names the frame to copy).
    !(snapshot.body && (snapshot.body.from || Object.keys(snapshot.body.block).length || Object.keys(snapshot.body.style).length))
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
  /**
   * A text box painted whole: its own first list level's run defaults
   * (`a:lstStyle/a:lvl1pPr/a:defRPr`, truth c64painter P2) -- what its words
   * now leave to their level (`inheritedKeys`) -- by level key, `null`
   * taking one off. The block's `source.nativeLevelStyles[0]` takes it.
   */
  levelOnePatch?: { [K in PresentationLevelStyleKey | "themeFontRole"]?: PresentationLevelStyle[K] | null };
  /** The source's whole text frame, onto a text box's or a shape's (whole-object paints only). */
  bodyPatch?: FormatPainterBodySnapshot;
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
    // A SHAPE's words: every character takes what the source's first
    // character wears -- its run's own format over its box's -- written on
    // the shape's runs, over their own formats; the shape's box keeps its own
    // (truth c64painter P1 / P4: a rectangle's and a content placeholder's
    // runs of 28 / 16 pt Cambria all became 24 pt bold, their `a:lstStyle`
    // untouched). Left on the box, a run that stated its own size and face
    // stayed as it was, and the save could not write a box-level change to a
    
    // its box (P2 / P3: its `a:lstStyle` level 1), as before.
    if (shapeHasText && snapshot.scope !== "text-selection" && (snapshot.text || snapshot.runStyle)) {
      plan.runStylePatch = sourceCharacterFormat(snapshot);
      plan.textPatch = snapshot.text?.align !== undefined ? { align: snapshot.text.align } : undefined;
    }
    if (kind === "text" && snapshot.paragraph) {
      plan.paragraphPatch = cloneRecord(snapshot.paragraph);
    }
    if (kind === "text" && snapshot.scope !== "text-selection" && typeof target.text === "string" && target.text !== "") {
      // A text box takes the format on its box (truth c64painter P2: the
      // box's `a:lstStyle` level 1 states the source's format, every run's
      // `a:rPr` emptied), so a run that stated its own size or italic no
      // longer keeps it, and words typed into the emptied box take it (P3).
      const effective = sourceCharacterFormat(snapshot);
      const box: Record<string, unknown> = { ...(plan.textPatch ?? {}) };
      for (const key of TEXT_STYLE_KEYS) {
        if (key === "align") continue;
        const value = effective[key as keyof FormatPainterRunStylePatch];
        if (value === null) delete box[key];
        else if (value !== undefined) box[key] = value;
      }
      plan.textPatch = Object.keys(box).length > 0 ? (box as TextStyleSnapshot) : undefined;
      // The runs leave the painted format to the box's first level, which
      // states it (`levelOnePatch`), as PowerPoint empties every run's
      // `a:rPr` (P2): they wear its values and say they are the level's
      // (`inheritedKeys`), so the save writes the level and hands the runs'
      // own back -- keeping them only where the file's chain would draw
      
      const level = levelOneOf(effective);
      const inherited = PRESENTATION_LEVEL_STYLE_KEYS.filter((key) => Object.hasOwn(level, key));
      plan.runStylePatch = {
        ...effective,
        ...(inherited.length > 0 ? { inheritedKeys: inherited } : {}),
        ...(inherited.includes("color") && typeof effective.color === "string" ? { colorFromList: true } : {}),
      };
      if (inherited.length > 0) plan.levelOnePatch = level;
    }
  }
  // The whole text frame (`a:bodyPr`), a whole-object paint onto a text box
  // or a shape that has one: PowerPoint replaces the target's frame with the
  // source's whole (truth c62bodypr: anchor, wrap, margins, direction,
  // columns and autofit; c62bodypr2: WordArt Transform -- a frame the source
  // does not warp loses its own warp).
  if (snapshot.body && snapshot.scope !== "text-selection" && hasTextFrame(target)) {
    plan.bodyPatch = JSON.parse(JSON.stringify(snapshot.body)) as FormatPainterBodySnapshot;
  }
  return plan;
}

/**
 * The first list level a text box painted whole takes (`levelOnePatch`): the
 * painted size, colour and face where the format states them, and bold and
 * italic always (a flag the source does not wear is off, P2's `b="0"`).
 */
function levelOneOf(effective: FormatPainterRunStylePatch): NonNullable<FormatPaintPlan["levelOnePatch"]> {
  const level: NonNullable<FormatPaintPlan["levelOnePatch"]> = {};
  if (typeof effective.fontSize === "number") level.fontSize = effective.fontSize;
  level.bold = effective.bold === true ? true : null;
  level.italic = effective.italic === true ? true : null;
  if (typeof effective.color === "string" && /^#[0-9a-f]{6}$/i.test(effective.color)) level.color = effective.color.toLowerCase();
  if (typeof effective.fontFamily === "string" && effective.fontFamily.trim()) {
    level.fontFamily = effective.fontFamily;
    level.themeFontRole =
      effective.themeFontRole === "heading" || effective.themeFontRole === "body" ? effective.themeFontRole : null;
  }
  return level;
}

/** The character format keys Format Painter carries whole. */
const SOURCE_CHARACTER_KEYS = [
  "fontSize",
  "fontFamily",
  "themeFontRole",
  "color",
  "bold",
  "italic",
  "underline",
  "strike",
  "superscript",
  "subscript",
] as const satisfies readonly (keyof FormatPainterRunStylePatch)[];

/**
 * The source's first character's whole format, as Format Painter carries it:
 * its run's own over its box's, a flag neither states off (`null`: the
 * painted run states none and draws it off, as PowerPoint writes it -- truth
 * c64painter P1: a target's `i="1"` gone under a source that states no
 * italic), and a face that names no theme font no longer following one.
 */
function sourceCharacterFormat(snapshot: FormatPainterSnapshot): FormatPainterRunStylePatch {
  const box = snapshot.text
    ? (Object.fromEntries(
        Object.entries(snapshot.text).filter(([key, value]) => key !== "align" && value !== undefined),
      ) as FormatPainterRunStylePatch)
    : {};
  const run = snapshot.runStyle ?? {};
  const effective: FormatPainterRunStylePatch = {
    bold: null,
    italic: null,
    underline: null,
    strike: null,
    superscript: null,
    subscript: null,
    ...box,
    ...run,
  };
  for (const flag of ["bold", "italic", "underline", "strike", "superscript", "subscript"] as const) {
    if (effective[flag] === false) effective[flag] = null;
  }
  if (effective.fontFamily !== undefined && effective.fontFamily !== null && !Object.hasOwn(run, "themeFontRole") && !Object.hasOwn(box, "themeFontRole")) {
    effective.themeFontRole = null;
  }
  return effective;
}

/** True when a paint plan carries no applicable patch (a no-op paint). */
export function isEmptyPlan(plan: FormatPaintPlan): boolean {
  return (
    !plan.shapePatch &&
    !plan.imagePatch &&
    !plan.tablePatch &&
    !plan.textPatch &&
    !plan.runStylePatch &&
    !plan.paragraphPatch &&
    !plan.bodyPatch
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
