/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for spreadsheet DRAWING SHAPES (Excel/WPS "形状" — rectangles,
 * ellipses, lines, arrows, triangles, optionally with a text label), plus the
 * bounded Web-ink paths authored by the Draw tab. Both use the existing
 * floating-object system alongside `image` and `chart`: object POSITION / SIZE /
 * anchor remain shared while this module owns the normalized `object.shape`
 * payload and the two object factories.
 */

import { officeFillPatternOf, type OfficeChartFillPattern } from "../office-chart-shape-paint";
import { normalizeOfficeTextWarp, type OfficeTextWarp } from "../office-text-warp";
import type { TextRun } from "../presentation-runs";
import { normalizeSpreadsheetShapeTextRuns } from "./shape-text-runs";
import {
  normalizeOfficeTextPaint,
  type OfficeTextPaint,
} from "../office-text-paint";
import type { InkShapeFit } from "../office-ink-shape-fit";
import {
  inkStrokeToShape,
  normalizePresentationCustomPath,
  PRESENTATION_SCRIBBLE_MAX_POINTS,
  type PresentationCustomPath,
} from "../presentation-freeform";
import {
  normalizeDocumentInkEffect,
  normalizeDocumentShapeAdjustments,
  normalizeInkCutEnds,
  type DocumentInkEffect,
  type InkCutEnds,
  normalizeInkPressure,
} from "../document-shapes";
import { documentInkPressureOf, documentInkPressures } from "../document-ink";
// One transparency model for every Office surface, rather than a sheet-only
// spelling of the same percentage.
import {
  DEFAULT_SHAPE_OPACITY,
  PRESENTATION_CONNECTOR_END_SIZES,
  PRESENTATION_CONNECTOR_ENDS,
  normalizeShapeOpacity,
  type PresentationConnectorEnd,
  type PresentationConnectorEndSize,
} from "../presentation-shapes";
import {
  OFFICE_SHAPE_DEFINITIONS,
  OFFICE_SHAPE_PRESETS,
  isOfficeShapePreset,
  officeShapeDefinition,
  type OfficeShapePreset,
} from "../office-shape-catalog";
// The label's FONT (family, bold, italic, underline, strikethrough) in the
// field names the Docs shape already persists, shared by every host.
import {
  normalizeOfficeShapeTextFont,
  type OfficeShapeTextFont,
} from "../office-shape-text-font";
// One text-body layout envelope for every host that draws an `a:txBody`,
// rather than a sheet-only spelling of the same paragraph properties.
import {
  normalizeOfficeTextLayout,
  type OfficeTextLayout,
} from "../office-text-layout";

export const SPREADSHEET_SHAPE_KINDS = [
  "rectangle",
  "rounded-rectangle",
  "ellipse",
  "triangle",
  "line",
  "arrow",
  "text-box",
  // The rest of the gallery WPS puts on its Insert tab. Each maps to one
  // OOXML preset geometry, so a shape drawn here is the same shape there and
  // a workbook that already had one is read rather than passed over.
  "right-triangle",
  "diamond",
  "pentagon",
  "hexagon",
  "star",
  "left-arrow",
  "up-arrow",
  "down-arrow",
  "chevron",
  "plus",
  "cloud-callout",
  "speech-bubble",
] as const;
export type SpreadsheetShapeKind = (typeof SPREADSHEET_SHAPE_KINDS)[number];
export type SpreadsheetPresetShapeKind =
  | OfficeShapePreset
  | SpreadsheetShapeKind
  | "text-box";
export const SPREADSHEET_INK_SHAPE_KIND = "ink" as const;
/**
 * A shape drawn by its OWN outline rather than by a preset.
 *
 * The SmartArt engine draws several layouts as real outlines — a segmented
 * cycle's wedges, a step-up's treads, a funnel's vessel, the pyramids — and a
 * sheet used to flatten every one of them to a rectangle, because a sheet shape
 * could only name a preset. This is the same escape hatch Draw ink already
 * uses, with the points in the shape's own space.
 */
export const SPREADSHEET_CUSTOM_SHAPE_KIND = "custom-geometry" as const;
export type SpreadsheetObjectShapeKind =
  | SpreadsheetPresetShapeKind
  | typeof SPREADSHEET_INK_SHAPE_KIND
  | typeof SPREADSHEET_CUSTOM_SHAPE_KIND;
const SHAPE_KIND_SET: ReadonlySet<string> = new Set([
  ...OFFICE_SHAPE_PRESETS,
  ...OFFICE_SHAPE_DEFINITIONS.map((definition) => definition.previewKind),
  ...SPREADSHEET_SHAPE_KINDS,
  SPREADSHEET_INK_SHAPE_KIND,
  SPREADSHEET_CUSTOM_SHAPE_KIND,
]);

export const DEFAULT_SHAPE_KIND: SpreadsheetPresetShapeKind = "rectangle";
export const DEFAULT_SHAPE_FILL = "#dbeafe";
export const DEFAULT_SHAPE_STROKE = "#2563eb";
export const DEFAULT_SHAPE_STROKE_WIDTH = 2;
export const MAX_SHAPE_STROKE_WIDTH = 24;


export const MAX_SHAPE_TEXT_LENGTH = 32_767;
/** XLS-A016: the most LaTeX one equation text box keeps (the server's bound). */
export const MAX_SHAPE_EQUATION_LATEX_LENGTH = 8_000;







export function spreadsheetEquationFits(latex: string, linear: string): boolean {
  return latex.length <= MAX_SHAPE_EQUATION_LATEX_LENGTH && linear.length <= MAX_SHAPE_TEXT_LENGTH;
}
/** The one freeform / ink vertex limit every Office host shares (the deck's
 * Scribble limit, `OFFICE_FREEFORM_MAX_POINTS`), not a sheet-only number. */
export const MAX_SPREADSHEET_INK_POINTS = PRESENTATION_SCRIBBLE_MAX_POINTS;
export const MAX_SPREADSHEET_INK_POINTS_CHARS = 16_384;
export const SPREADSHEET_HIGHLIGHTER_OPACITY = 0.38;
/** A pencil paints a little translucent, as its grain lets paper through. */
export const SPREADSHEET_PENCIL_OPACITY = 0.9;
export const SPREADSHEET_SHAPE_RECENT_STORAGE_KEY =
  "mrnote.spreadsheet.shape-recents.drawingml-prstgeom-v1";

/** The Draw tab's three kinds of pen (Office's Pen, Highlighter and Pencil). */
export type SpreadsheetInkTool = "pen" | "highlighter" | "pencil";

/** How a tool paints: a highlighter is translucent, a pencil a little, a pen opaque. */
export function spreadsheetInkOpacityForTool(tool: SpreadsheetInkTool): number {
  if (tool === "highlighter") return SPREADSHEET_HIGHLIGHTER_OPACITY;
  if (tool === "pencil") return SPREADSHEET_PENCIL_OPACITY;
  return 1;
}

export interface SpreadsheetInkPoint {
  x: number;
  y: number;
  /** The pen's pressure at this sample (0..1), when the pen reported one. */
  p?: number;
}

/** A shape's persisted style (the geometry lives on the object/anchor). */
import {
  normalizeOfficeShapeGradient,
  normalizeOfficeShapeShadow,
  type OfficeShapeGradient,
  type OfficeShapeShadow,
} from "../office-shape-paint";
import { officeStrokeAlign, type OfficeStrokeAlign } from "../office-shape-line-align";
import {
  normalizeOfficeOutlineDash,
  type OfficeMenuDash,
  type OfficeStrokeDash,
} from "../office-stroke-dash";
import { officeStrokeDetail, type OfficeStrokeDetail } from "../office-stroke-detail";
import { smartArtPictureSrc } from "../office-smartart/picture";
import {
  normalizeOfficePictureFillOptions,
  officePictureFillOptionsStated,
  type OfficePictureFillOptions,
} from "../office-picture-fill-options";
import {
  isNativeScene3d,
  isNativeShape3d,
  isSmartArtScene3d,
  isSmartArtShape3d,
  isSmartArtStyleId,
  isSmartArtTextShape3d,
  type SmartArtScene3d,
  type SmartArtShape3d,
  type SmartArtStyleId,
  type SmartArtTextShape3d,
} from "../office-smartart/styles-model";
/**
 * `OfficeShapeTextFont` is the label's font as Excel's Home ▸ Font group sets
 * it on a selected shape: `textFontFamily`, `textBold`, `textItalic`,
 * `textUnderline`, `textStrike`, written onto every `a:rPr` of the label.
 */
export type SpreadsheetShapeTextAlign = "left" | "center" | "right" | "justify";
export type SpreadsheetShapeTextVerticalAlign = "top" | "middle" | "bottom";
export const SPREADSHEET_SHAPE_TEXT_ALIGNS: readonly SpreadsheetShapeTextAlign[] = [
  "left",
  "center",
  "right",
  "justify",
];
export const SPREADSHEET_SHAPE_TEXT_VERTICAL_ALIGNS: readonly SpreadsheetShapeTextVerticalAlign[] =
  ["top", "middle", "bottom"];

export interface SpreadsheetShapeStyle extends OfficeShapeTextFont {
  fillGradient?: OfficeShapeGradient;
  /**
   * Format Shape > Fill & Line beyond the colours, as the server reads and
   * writes them (`a:pattFill`, the solid fill's and line's `a:alpha`, a
   * gradient line `a:ln/a:gradFill`). Absent: the page states none (the
   * file's is kept); null: none, written out.
   */
  fillPattern?: OfficeChartFillPattern | null;
  /** Whole percent 1..100 on a solid fill. */
  fillTransparency?: number | null;
  /** Whole percent 1..100 on a solid line. */
  outlineTransparency?: number | null;
  outlineGradient?: OfficeShapeGradient | null;
  kind: SpreadsheetObjectShapeKind;
  /**
   * XLS-A016: Insert ▸ Equation — the text box holds an equation (the shared
   * equation editor's LaTeX); its `text` carries the UnicodeMath linear form
   * a reader without the equation sees.
   */
  equation?: { latex: string };
  /**
   * Format Shape > Text Box > "Resize shape to fit text" (`a:spAutoFit` in
   * the label's `a:bodyPr`), in the words a Docs shape keeps it
   * (`DocumentShapeAttrs.textAutofit`): the frame follows its words. `"none"`
   * is an explicit off; absent states nothing. The sheet resizes a box that
   * also does not wrap (its `textLayout.bodyPr.wrap` is `"none"`) to its
   * longest line and its lines, as Excel's Insert > Equation box does
   * (`spreadsheetShapeResizesToFitText`).
   */
  textAutofit?: "none" | "shape";
  /**
   * Font > Character Spacing of the label (`a:rPr@spc`, the first run's), in
   * points: Expanded positive, Condensed negative (S30_textbox-
   * CharKerningSpace's "A V A I L" by 20 pt, X-162).
   */
  textCharacterSpacing?: number;
  /**
   * A `custom-geometry` shape's own outline, in the shared persisted envelope
   * every host stores (`normalizePresentationCustomPath`) — what an imported
   * `a:custGeom` becomes, beside the older polygon-only `customPoints`.
   */
  customPath?: PresentationCustomPath;
  /**
   * The dragged adjust handles of a preset shape (`a:prstGeom/a:avLst`), by
   * the preset's own guide names, exactly as a document shape carries them.
   * Absent means the preset's defaults.
   */
  adjustments?: Readonly<Record<string, number>>;
  /**
   * `a:ln/a:headEnd|a:tailEnd`. {@link kind} names WHICH ends are decorated;
   * these name WHICH decoration (all six `a:ST_LineEndType` members) and its
   * optional `@w`/`@len`. A value the kind already implies is absent, and an
   * absent size stays absent: `CT_LineEndProperties` declares no default for
   * any of them (docs/office-graphics/SHAPE-IMPORT-AUDIT.md §7.4).
   */
  headEnd?: PresentationConnectorEnd;
  tailEnd?: PresentationConnectorEnd;
  headEndWidth?: PresentationConnectorEndSize;
  headEndLength?: PresentationConnectorEndSize;
  tailEndWidth?: PresentationConnectorEndSize;
  tailEndLength?: PresentationConnectorEndSize;
  /** Interior fill: a #rrggbb hex, or "none" for an unfilled outline. */
  fill: string;
  /**
   * The picture in a SmartArt PICTURE SLOT, as the inline data URL every host
   * spells a SmartArt picture with (`lib/office-smartart/picture`).
   *
   * A sheet diagram is N ordinary shape objects, so a picture point is a shape
   * in the neutral placeholder colours — it has no picture of its own, which is
   * why the four picture layouts used to export with empty slots. It stays a
   * shape; this is the slot REMEMBERING what is in it, drawn over the preset by
   * `SpreadsheetShapeView` and written onto the data point as an `a:blipFill`
   * by the shared DiagramML writer (SMARTART-IMPORT.md §5.3).
   *
   * It lives on the STYLE rather than beside the object's identity markers
   * because that is where the sheet's own lift already looks for it
   * (`spreadsheet_xlsx_smartart_diagram` names `src` among the keys it takes
   * off the object's `shape` map), and because `fill` — the other thing that
   * says what the box looks like — is here too.
   */
  src?: string;
  /**
   * Format Shape > Fill > Picture or texture fill authored on the page:
   * `src` is the shape's OWN fill picture, stretched over its box inside the
   * geometry (`a:blipFill` + `a:stretch`), not a SmartArt slot's picture. An
   * imported fill says the same through `source.pictureFillPart` (the package
   * part it came from); this marker is the page's. Present only with `src`.
   */
  fillPicture?: true;
  /**
   * The picture fill's Format Shape options in the shared model
   * (`lib/office-picture-fill-options`): tiled with `a:tile`, or stretched
   * into `fillRect`; Transparency (`a:alphaModFix`); Rotate with shape. For
   * an authored or an imported fill alike. Absent: stretched over the box,
   * opaque, as every sheet stored before it.
   */
  fillPictureOptions?: OfficePictureFillOptions;
  /** Outline color: a #rrggbb hex, or "none" for no border. */
  stroke: string;
  /** Outline width in px (0..MAX_SHAPE_STROKE_WIDTH). */
  strokeWidth: number;
  /**
   * `a:ln@algn`, in the shared vocabulary (`@/lib/office-shape-line-align`).
   * `"inset"` is the `in` pen: the whole stroke belongs INSIDE the geometry
   * path (Office's `LineFormat.InsetPen`). ABSENT is an object whose line
   * states no alignment -- every shape drawn in this editor, and every
   * imported one whose `a:ln` carries no `algn` -- so it is drawn centred and
   * exported with no `algn` at all.
   */
  strokeAlign?: OfficeStrokeAlign;
  /**
   * Office's line detail, in the words a slide's and a Docs shape's line
   * keep: the exact dash (Shape Outline > Dashes), compound type, cap, join
   * and miter limit, and Sketched. Absent says nothing (a solid, single line).
   */
  outlineDash?: OfficeMenuDash | OfficeStrokeDash;
  outlineCompound?: OfficeStrokeDetail["outlineCompound"];
  outlineCap?: OfficeStrokeDetail["outlineCap"];
  outlineJoin?: OfficeStrokeDetail["outlineJoin"];
  outlineMiterLimit?: number;
  outlineSketch?: "curved" | "freehand" | "scribble";
  /** Optional centered text label. Absent when empty. */
  text?: string;
  /**
   * The character formatting of PART of {@link text}: the same runs a Docs
   * shape keeps (`./shape-text-runs`), each run's field overriding the
   * shape's own. Absent when every character wears the shape's format. Kept
   * only while the runs spell exactly `text`.
   */
  textRuns?: TextRun[];
  /**
   * The LAYOUT of those same characters — each line's numbering, outline level
   * and `a:pPr` margins, plus the `a:bodyPr` frame — as a display-only receipt
   * in the shared envelope every host reads (`lib/office-text-layout`).
   *
   * It is NOT a second source of truth for the characters: {@link text} is, and
   * the writer plans its edit against that string alone, so an untouched shape
   * is still written byte-for-byte. The receipt is read POSITIONALLY, one entry
   * per `\n`-separated line, and says nothing at all once the lines no longer
   * match it — the disposal `nativeBullet` already has
   * (docs/office-graphics/SHAPE-IMPORT-AUDIT.md §44).
   */
  textLayout?: OfficeTextLayout;
  /**
   * The words stay level while the shape turns (`a:bodyPr@upright`) without
   * a package layout to carry it: a SmartArt shape the layout turns under a
   * label it keeps level (Inverted Pyramid's layers). Only ever `true`.
   */
  textUpright?: true;
  /**
   * The label's WordArt Transform (`a:bodyPr/a:prstTxWarp`): the ONE warp
   * value of `@/lib/office-text-warp`, drawn by the shared painter. Display
   * only — the source package keeps the element.
   */
  textWarp?: OfficeTextWarp;
  /**
   * The label's own horizontal alignment (`a:pPr@algn` on every paragraph)
   * and vertical anchor (`a:bodyPr@anchor`), as Excel's Home ▸ Alignment
   * sets them with the shape selected. Absent draws what the shape always
   * drew (a text box top-left, any other shape centred) or what its imported
   * layout receipt says; a stated value wins over both.
   */
  textAlign?: SpreadsheetShapeTextAlign;
  textVerticalAlign?: SpreadsheetShapeTextVerticalAlign;
  /**
   * WordArt (艺术字): how the label is drawn. Absent means plain text at the
   * inherited size — the label a rectangle has always had. A preset is stored
   * by name so a workbook keeps meaning something if the presets are retuned.
   */
  textArt?: SpreadsheetTextArtPreset;
  /**
   * WordArt paint (the SHARED `OfficeTextPaint` Docs and the deck keep): text
   * fill, outline and effects, as Excel's WordArt Styles group writes them
   * onto the label's runs. When present it supersedes `textArt` / `textColor`.
   */
  textPaint?: OfficeTextPaint;
  /**
   * Label size in CSS px (`a:rPr@sz` / 75), kept to 0.01 px so a point size
   * such as 11 pt (14.67 px) survives the model. Absent inherits the shape
   * renderer's default.
   */
  textSize?: number;
  /** Label colour: a #rrggbb hex. Absent inherits the surrounding ink. */
  textColor?: string;
  /** Draw-tab ink kind. Absent on ordinary Insert-tab shapes. */
  inkTool?: SpreadsheetInkTool;
  /**
   * One of Office's eight pen effects (Rainbow, Galaxy, ...), pens only. The
   * stroke paints the effect's gradient; {@link stroke} holds its solid
   * stand-in, which is what every file writer states.
   */
  inkEffect?: DocumentInkEffect;
  /** Bounded points in this shape's own 0..100 coordinate space. */
  inkPoints?: string;
  /**
   * XLS-A032: Ink Replay — when each of {@link inkPoints} was drawn, in whole
   * milliseconds from the first, as `0 t1 t2 …` (one per point). Absent on
   * strokes drawn before timing was kept; those replay at a constant speed.
   */
  inkTimes?: string;
  /** A stroke a point eraser cut: which of its ends stop flat where the
   * eraser was (the shared {@link InkCutEnds}). Absent: both ends round. */
  inkCutEnds?: InkCutEnds;
  /**
   * A stroke a pressure pen drew: one integer 0..100 per {@link inkPoints}
   * point, space-separated; the painter widens and thins the stroke with
   * it. Absent on a stroke drawn at one width. Paint only: no file writer
   * reads it (a file line has one width).
   */
  inkPressure?: string;
  /**
   * A closed outline in this shape's own 0..100 space, as `x,y x,y …` — the
   * same spelling `inkPoints` uses, so one parser serves both. Present only on
   * a {@link SPREADSHEET_CUSTOM_SHAPE_KIND} shape.
   */
  customPoints?: string;
  /** Highlighter compositing opacity. Pen paths always normalize to 1. */
  inkOpacity?: number;
  /**
   * Transparency as a PERCENTAGE (100 = opaque), in the shared shape model
   * (`normalizeShapeOpacity`). Absent means opaque, so an ordinary shape
   * persists exactly what it did before transparency existed.
   *
   * A slide draws a SmartArt wash — a grouped list's plate behind its items, a
   * Venn's overlapping discs — by making the member itself translucent. A sheet
   * that dropped the value drew those as solid blocks, which is a different
   * picture, so the value crosses with the rest of the paint.
   */
  opacity?: number;
  /**
   * SmartArt quick-style paint (`docs/office-graphics/SMARTART-STYLES.md` §4),
   * in the SHARED vocabulary the engine resolves and the PPT blocks carry — not
   * a second, sheet-only spelling of the same DrawingML.
   *
   * A sheet diagram is its members as ordinary shape objects
   * (`lib/spreadsheet/smart-art`), so where a slide keeps these fields on the
   * shape BLOCK the worksheet keeps them here, next to `fillGradient`, which is
   * the same paint layer. The identity markers that say WHICH diagram a member
   * belongs to (`smartArtKind` / `groupId` / …) stay on the object itself; the
   * style id lives here because it is what the paint was resolved from.
   *
   * `shapeShadow` is the bounded outer shadow of `simple5` — the sheet's older
   * `object.shadow` preset stays the manual Shape-Effects vocabulary, and this
   * one wins wherever both appear, exactly as it does on a slide.
   */
  shapeShadow?: OfficeShapeShadow;
   
  shape3d?: SmartArtShape3d;
  /** The DIAGRAM camera (styles `3d5`–`3d9`), identical on every member. */
  scene3d?: SmartArtScene3d;
  /** 3-D on the label's own text body (`a:bodyPr/a:sp3d`); style `3d9` only. */
  textShape3d?: SmartArtTextShape3d;
  /** The one quick style this member's paint was resolved from. */
  smartArtStyle?: SmartArtStyleId;
  /**
   * An ORDINARY shape's own `a:sp3d`/`a:scene3d` (Excel: Format Shape ▸ Effects
   * ▸ 3-D Format / 3-D Rotation), as a DISPLAY-ONLY value.
   *
   * Deliberately NOT the editable {@link shape3d}: the solid belongs to the
   * source package and comes back byte for byte through a save, no writer reads
   * this, and nothing on the page edits it. It exists so the canvas can draw
   * the shape with the very bevel/extrusion/contour renderer a diagram member
   * already uses, instead of drawing a bevelled block as a flat rectangle —
   * the `nativeChartLook` arrangement, one layer of paint down
   * (docs/office-graphics/SHAPE-IMPORT-AUDIT.md §44.4).
   */
  nativeShape3d?: { shape3d: SmartArtShape3d; scene3d?: SmartArtScene3d };
}

/**
 * The WordArt presets, named for what they do rather than for a gallery
 * position, so the stored name survives a redesign of the gallery.
 */
export const SPREADSHEET_TEXT_ART_PRESETS = [
  "plain",
  "filled",
  "outlined",
  "shadowed",
  "gradient",
] as const;
export type SpreadsheetTextArtPreset =
  (typeof SPREADSHEET_TEXT_ART_PRESETS)[number];
const TEXT_ART_SET: ReadonlySet<string> = new Set(SPREADSHEET_TEXT_ART_PRESETS);

export const MIN_SHAPE_TEXT_SIZE = 8;
/**
 * 409 pt, the largest size Excel's Home ▸ Font Size box accepts, in px to the
 * model's 0.01 px. The old 96 px (72 pt) cap refused sizes Excel offers.
 */
export const MAX_SHAPE_TEXT_SIZE = 545.33;

/** Coerce to a known WordArt preset; anything else is plain text. */
export function normalizeTextArtPreset(
  value: unknown,
): SpreadsheetTextArtPreset | undefined {
  return typeof value === "string" && TEXT_ART_SET.has(value)
    ? (value as SpreadsheetTextArtPreset)
    : undefined;
}

/** Clamp a label size into the renderable range; junk drops the override. */
export function normalizeShapeTextSize(value: unknown): number | undefined {
  const size = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(size)) return undefined;
  return Math.max(
    MIN_SHAPE_TEXT_SIZE,
    Math.min(MAX_SHAPE_TEXT_SIZE, Math.round(size * 100) / 100),
  );
}

const HEX_RE = /^#[0-9a-fA-F]{6}$/;

/** Coerce to a whitelisted shape kind, defaulting to rectangle. */
export function normalizeShapeKind(
  value: unknown,
): SpreadsheetObjectShapeKind {
  return typeof value === "string" && SHAPE_KIND_SET.has(value)
    ? (value as SpreadsheetObjectShapeKind)
    : DEFAULT_SHAPE_KIND;
}

export function isSpreadsheetPresetShapeKind(
  value: unknown,
): value is SpreadsheetPresetShapeKind {
  return (
    typeof value === "string" &&
    value !== SPREADSHEET_INK_SHAPE_KIND &&
    value !== SPREADSHEET_CUSTOM_SHAPE_KIND &&
    SHAPE_KIND_SET.has(value)
  );
}

export function isSpreadsheetLineLikeShapeKind(
  kind: SpreadsheetPresetShapeKind,
): boolean {
  const kindValue: string = kind;
  return isOfficeShapePreset(kindValue)
    ? officeShapeDefinition(kindValue).lineLike === true
    : kindValue === "line" || kindValue === "arrow";
}

export function spreadsheetShapeKindFromOfficePreset(
  preset: OfficeShapePreset,
): SpreadsheetPresetShapeKind {
  return officeShapeDefinition(preset).previewKind;
}

export function officeShapePresetForSpreadsheetKind(
  kind: SpreadsheetPresetShapeKind,
): OfficeShapePreset {
  const kindValue: string = kind;
  if (isOfficeShapePreset(kindValue)) return kindValue;
  return (
    OFFICE_SHAPE_DEFINITIONS.find(
      (definition) => definition.previewKind === kindValue,
    )?.preset ?? "rect"
  );
}

/**
 * The six optional line-end fields, held to the shared vocabularies. A value
 * outside them is DROPPED rather than failing the shape: `kind` still names a
 * real geometry to draw.
 */
function normalizeSpreadsheetShapeLineEnds(
  o: Readonly<Record<string, unknown>>,
): Partial<SpreadsheetShapeStyle> {
  const fields: Record<string, string> = {};
  for (const [key, vocabulary] of [
    ["headEnd", PRESENTATION_CONNECTOR_ENDS],
    ["tailEnd", PRESENTATION_CONNECTOR_ENDS],
    ["headEndWidth", PRESENTATION_CONNECTOR_END_SIZES],
    ["headEndLength", PRESENTATION_CONNECTOR_END_SIZES],
    ["tailEndWidth", PRESENTATION_CONNECTOR_END_SIZES],
    ["tailEndLength", PRESENTATION_CONNECTOR_END_SIZES],
  ] as const) {
    const value = o[key];
    if (typeof value !== "string") continue;
    if (!(vocabulary as readonly string[]).includes(value)) continue;
    fields[key] = value;
  }
  return fields as Partial<SpreadsheetShapeStyle>;
}

/** Coerce to a #rrggbb hex (lower-cased) or the literal "none"; else fallback. */
export function normalizeShapeColor(value: unknown, fallback: string): string {
  if (value === "none") return "none";
  if (typeof value === "string" && HEX_RE.test(value))
    return value.toLowerCase();
  return fallback;
}

/** Clamp a stroke width into [0, MAX]; non-finite → the default. */
export function normalizeStrokeWidth(value: unknown): number {
  const n = Number(value);
  if (!Number.isFinite(n)) return DEFAULT_SHAPE_STROKE_WIDTH;
  return Math.max(0, Math.min(MAX_SHAPE_STROKE_WIDTH, Math.round(n*1000)/1000));
}

function formatInkCoordinate(value: number): string {
  return String(Math.round(value * 1_000) / 1_000);
}

/**
 * Parse one bounded persisted Draw path. Invalid/non-finite payloads produce an
 * empty list, and callers consequently drop the malformed ink object instead of
 * mounting an invisible or unbounded SVG path.
 */
export function parseSpreadsheetInkPoints(
  value: unknown,
): SpreadsheetInkPoint[] {
  if (
    typeof value !== "string" ||
    value.length === 0 ||
    value.length > MAX_SPREADSHEET_INK_POINTS_CHARS
  ) {
    return [];
  }
  const tokens = value.trim().split(/\s+/);
  if (tokens.length < 2 || tokens.length > MAX_SPREADSHEET_INK_POINTS) {
    return [];
  }
  const points: SpreadsheetInkPoint[] = [];
  for (const token of tokens) {
    const parts = token.split(",");
    if (parts.length !== 2) return [];
    const x = Number(parts[0]);
    const y = Number(parts[1]);
    if (!Number.isFinite(x) || !Number.isFinite(y)) return [];
    points.push({
      x: Math.max(0, Math.min(100, x)),
      y: Math.max(0, Math.min(100, y)),
    });
  }
  return points;
}

export function normalizeSpreadsheetInkPoints(
  value: unknown,
): string | undefined {
  const points = parseSpreadsheetInkPoints(value);
  if (points.length < 2) return undefined;
  return points
    .map(
      (point) =>
        `${formatInkCoordinate(point.x)},${formatInkCoordinate(point.y)}`,
    )
    .join(" ");
}

/** XLS-A032: the longest stretch one stroke's timing may state (10 minutes). */
export const MAX_SPREADSHEET_INK_TIME_MS = 600_000;

/**
 * Parse a stroke's point times (XLS-A032). Only a list of exactly `count`
 * whole, non-decreasing milliseconds starting at 0 is timing; anything else is
 * no timing at all (the stroke replays at a constant speed).
 */
export function parseSpreadsheetInkTimes(value: unknown, count: number): number[] | null {
  if (typeof value !== "string" || count < 2) return null;
  const tokens = value.trim().split(/\s+/);
  if (tokens.length !== count) return null;
  const times: number[] = [];
  for (const token of tokens) {
    if (!/^\d{1,7}$/.test(token)) return null;
    const time = Number(token);
    const previous = times[times.length - 1];
    if (time > MAX_SPREADSHEET_INK_TIME_MS || (previous !== undefined && time < previous)) return null;
    times.push(time);
  }
  return times[0] === 0 ? times : null;
}

export function normalizeSpreadsheetInkTimes(value: unknown, count: number): string | undefined {
  return parseSpreadsheetInkTimes(value, count)?.join(" ");
}

/** The stored pressures (`inkPressure`), one 0..1 per point, or null when
 * there are none or the count is off. */
export function parseSpreadsheetInkPressure(value: unknown, count: number): number[] | null {
  if (count < 2) return null;
  return documentInkPressures(value, count);
}

export function normalizeSpreadsheetInkPressure(value: unknown, count: number): string | undefined {
  return parseSpreadsheetInkPressure(value, count) ? normalizeInkPressure(value) : undefined;
}

/**
 * Same bounded 0..100 point list as ink, but an AREA rather than a stroke: it
 * takes at least three corners, because two points enclose nothing.
 */
export function normalizeSpreadsheetCustomPoints(
  value: unknown,
): string | undefined {
  const points = parseSpreadsheetInkPoints(value);
  if (points.length < 3) return undefined;
  return points
    .map(
      (point) =>
        `${formatInkCoordinate(point.x)},${formatInkCoordinate(point.y)}`,
    )
    .join(" ");
}

/** Sanitize a shape text label (strip control bytes, clamp); empty /
 *  whitespace-only / non-string → undefined so an unlabelled shape stores
 *  nothing.
 *
 *  The words are kept as typed, as Excel keeps a shape's text: runs of spaces,
 *  tabs, and spaces at either end of a line are content (user audit S-066:
 *  `A  B` came back `A B` once the box closed). A LINE BREAK is content too:
 *  the projection states a shape's paragraphs as lines, exactly as the package
 *  writer reads them (SHAPE-IMPORT-AUDIT.md §41.1), and the writer measures an
 *  edit against that reading; empty lines stay, trailing ones included, as the
 *  file's empty paragraphs. Only control bytes other than tab and line feed
 *  go, as the deck's shape text drops them (lib/presentation-text-format). */
const SHAPE_TEXT_STRIP_RE = /[\u0000-\u0008\u000b-\u001f\u007f-\u009f]/g;

export function normalizeShapeText(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const cleaned = value
    .replace(/\r\n?/g, "\n")
    .replace(SHAPE_TEXT_STRIP_RE, "")
    .slice(0, MAX_SHAPE_TEXT_LENGTH);
  return cleaned.trim() ? cleaned : undefined;
}

/**
 * Validate + normalize an arbitrary `object.shape` payload into a usable
 * {@link SpreadsheetShapeStyle}. Always returns a complete style (unknown fields
 * fall back to defaults) so a forged/partial value renders predictably.
 */
/** The line detail a shape style keeps, each field held to its vocabulary. */
function spreadsheetShapeLineDetail(
  o: Record<string, unknown>,
): Pick<
  SpreadsheetShapeStyle,
  "outlineDash" | "outlineCompound" | "outlineCap" | "outlineJoin" | "outlineMiterLimit" | "outlineSketch"
> {
  const dash = normalizeOfficeOutlineDash(o.outlineDash);
  const sketch =
    o.outlineSketch === "curved" || o.outlineSketch === "freehand" || o.outlineSketch === "scribble"
      ? o.outlineSketch
      : undefined;
  return {
    ...(dash ? { outlineDash: dash } : {}),
    ...officeStrokeDetail(o),
    ...(sketch ? { outlineSketch: sketch } : {}),
  };
}

/** A whole-percent Transparency 1..100, or null (none). */
function spreadsheetWholePercent(value: unknown): number | null {
  return typeof value === "number" && Number.isInteger(value) && value >= 1 && value <= 100 ? value : null;
}

export function normalizeShapeStyle(value: unknown): SpreadsheetShapeStyle {
  const o =
    value && typeof value === "object"
      ? (value as Record<string, unknown>)
      : {};
  const kind = normalizeShapeKind(o.kind);
  if (kind === SPREADSHEET_INK_SHAPE_KIND) {
    const inkTool: SpreadsheetInkTool =
      o.inkTool === "highlighter" || o.inkTool === "pencil" ? o.inkTool : "pen";
    const normalizedStroke = normalizeShapeColor(o.stroke, "#111827");
    const defaultWidth = inkTool === "highlighter" ? 8 : 2;
    const suppliedWidth = Number(o.strokeWidth);
    const strokeWidth = Math.max(
      1,
      normalizeStrokeWidth(
        Number.isFinite(suppliedWidth) ? suppliedWidth : defaultWidth,
      ),
    );
    const suppliedOpacity = Number(o.inkOpacity);
    const inkOpacity =
      inkTool === "highlighter"
        ? Number.isFinite(suppliedOpacity)
          ? Math.max(0.12, Math.min(0.6, suppliedOpacity))
          : SPREADSHEET_HIGHLIGHTER_OPACITY
        : spreadsheetInkOpacityForTool(inkTool);
    const inkPoints = normalizeSpreadsheetInkPoints(o.inkPoints);
    const inkEffect = inkTool === "pen" ? normalizeDocumentInkEffect(o.inkEffect) : undefined;
    const inkCutEnds = normalizeInkCutEnds(o.inkCutEnds);
    const inkTimes = inkPoints
      ? normalizeSpreadsheetInkTimes(o.inkTimes, parseSpreadsheetInkPoints(inkPoints).length)
      : undefined;
    const inkPressure = inkPoints
      ? normalizeSpreadsheetInkPressure(o.inkPressure, parseSpreadsheetInkPoints(inkPoints).length)
      : undefined;
    return {
      kind,
      fill: "none",
      stroke: normalizedStroke === "none" ? "#111827" : normalizedStroke,
      strokeWidth,
      inkTool,
      inkOpacity,
      ...(inkPoints ? { inkPoints } : {}),
      ...(inkTimes ? { inkTimes } : {}),
      ...(inkPressure ? { inkPressure } : {}),
      ...(inkEffect ? { inkEffect } : {}),
      ...(inkCutEnds ? { inkCutEnds } : {}),
    };
  }
  // A custom outline is an ordinary painted shape that happens to carry its own
  // geometry instead of naming a preset. With no usable path there is nothing to
  // draw, so it falls back to the rectangle rather than disappearing.
  const customPath =
    kind === SPREADSHEET_CUSTOM_SHAPE_KIND
      ? normalizePresentationCustomPath(o.customPath)
      : null;
  const customPoints =
    kind === SPREADSHEET_CUSTOM_SHAPE_KIND && !customPath
      ? normalizeSpreadsheetCustomPoints(o.customPoints)
      : undefined;
  // Handles only mean something on a preset; a custom outline, ink and a text
  // box have none. The names are held to the kind's preset by the API.
  const adjustments =
    kind !== SPREADSHEET_CUSTOM_SHAPE_KIND && kind !== "text-box"
      ? normalizeDocumentShapeAdjustments(o.adjustments)
      : null;
  const style: SpreadsheetShapeStyle = {
    kind:
      kind === SPREADSHEET_CUSTOM_SHAPE_KIND && !customPoints && !customPath
        ? DEFAULT_SHAPE_KIND
        : kind,
    fill: normalizeShapeColor(o.fill, DEFAULT_SHAPE_FILL),
    ...(o.fillGradient?{fillGradient:normalizeOfficeShapeGradient(o.fillGradient)}:{}),
    ...(o.fillPattern !== undefined ? { fillPattern: officeFillPatternOf(o.fillPattern) } : {}),
    ...(o.fillTransparency !== undefined ? { fillTransparency: spreadsheetWholePercent(o.fillTransparency) } : {}),
    ...(o.outlineTransparency !== undefined ? { outlineTransparency: spreadsheetWholePercent(o.outlineTransparency) } : {}),
    ...(o.outlineGradient !== undefined ? { outlineGradient: normalizeOfficeShapeGradient(o.outlineGradient) ?? null } : {}),
    stroke: normalizeShapeColor(o.stroke, DEFAULT_SHAPE_STROKE),
    strokeWidth: normalizeStrokeWidth(o.strokeWidth),
    ...(officeStrokeAlign(o.strokeAlign)
      ? { strokeAlign: officeStrokeAlign(o.strokeAlign) as OfficeStrokeAlign }
      : {}),
    ...(customPath ? { customPath } : {}),
    ...(customPoints ? { customPoints } : {}),
    ...(adjustments ? { adjustments } : {}),
    ...normalizeSpreadsheetShapeLineEnds(o),
    ...spreadsheetShapeLineDetail(o),
    // Opaque is the default, so it stays ABSENT rather than persisting as 100.
    ...(normalizeShapeOpacity(o.opacity) < DEFAULT_SHAPE_OPACITY
      ? { opacity: normalizeShapeOpacity(o.opacity) }
      : {}),
  };
  // A line / arrow has no interior fill — keep it "none" so it renders as a stroke.
  if (
    isSpreadsheetPresetShapeKind(style.kind) &&
    isSpreadsheetLineLikeShapeKind(style.kind)
  ) {
    style.fill = "none";
  }
  const text = normalizeShapeText(o.text);
  if (text) style.text = text;
  if (text) {
    const runs = normalizeSpreadsheetShapeTextRuns(o.textRuns, text);
    if (runs) style.textRuns = runs;
  }
  // XLS-A016: an equation lives only in a text box.
  const equation = o.equation as { latex?: unknown } | undefined;
  if (
    style.kind === "text-box" &&
    equation &&
    typeof equation === "object" &&
    typeof equation.latex === "string" &&
    equation.latex.length <= MAX_SHAPE_EQUATION_LATEX_LENGTH
  ) {
    style.equation = { latex: equation.latex };
  }
  // Resize shape to fit text: the frame's own setting, kept on an empty box
  // too (Excel keeps an empty equation box's `a:bodyPr` for its words).
  if (o.textAutofit === "shape" || o.textAutofit === "none") style.textAutofit = o.textAutofit;
  if (
    typeof o.textCharacterSpacing === "number" &&
    Number.isFinite(o.textCharacterSpacing) &&
    o.textCharacterSpacing !== 0 &&
    Math.abs(o.textCharacterSpacing) <= 1584
  ) {
    style.textCharacterSpacing = o.textCharacterSpacing;
  }
  // The label's own styling only means anything when there is a label.
  // The shared WordArt paint is kept even before there is a label: Excel's
  // default shape carries its text formatting for the text typed later.
  try {
    const paint = normalizeOfficeTextPaint(o.textPaint);
    if (paint) style.textPaint = paint;
  } catch {
    // A malformed paint is dropped, never a reason to lose the label.
  }
  // The label's font (Home ▸ Font on the selected shape), field by field. Like
  // the paint it is kept on an empty shape too: Office keeps the run
  // properties of a body with no text for the words typed into it later, and
  // a SmartArt graphic formatted as a whole includes its empty items.
  Object.assign(style, normalizeOfficeShapeTextFont(o));
  // The label's own alignment (Home ▸ Alignment with the shape selected),
  // kept on an empty shape too, as the font is.
  if ((SPREADSHEET_SHAPE_TEXT_ALIGNS as readonly unknown[]).includes(o.textAlign)) {
    style.textAlign = o.textAlign as SpreadsheetShapeTextAlign;
  }
  if (
    (SPREADSHEET_SHAPE_TEXT_VERTICAL_ALIGNS as readonly unknown[]).includes(
      o.textVerticalAlign,
    )
  ) {
    style.textVerticalAlign = o.textVerticalAlign as SpreadsheetShapeTextVerticalAlign;
  }
  // Its size and colour too: an equation box just inserted is empty until
  // its first commit, and the size Home gave it while it was open (or the
  // sheet's 11 pt it went in with) was dropped, the draft drawn at 9 pt and
  // the box value springing back (live audit S-172).
  const size = normalizeShapeTextSize(o.textSize);
  if (size !== undefined) style.textSize = size;
  const color =
    typeof o.textColor === "string" && HEX_RE.test(o.textColor)
      ? o.textColor.toLowerCase()
      : undefined;
  if (color) style.textColor = color;
  if (text) {
    const preset = normalizeTextArtPreset(o.textArt);
    if (preset && preset !== "plain") style.textArt = preset;
    // The label's LAYOUT, kept only while it still describes exactly these
    // lines. A stale positional receipt is dropped rather than applied to rows
    // it no longer matches, so a retyped label draws as bare lines instead of
    // carrying another paragraph's number.
    const layout = normalizeOfficeTextLayout(o.textLayout);
    if (layout && layout.paragraphs.length === text.split("\n").length) {
      style.textLayout = layout;
    }
    const warp = normalizeOfficeTextWarp(o.textWarp);
    if (warp) style.textWarp = warp;
    if (o.textUpright === true) style.textUpright = true;
  }
  // SmartArt quick-style paint. Bounded by the SHARED guards (the same bounds
  // the API validator enforces), and a value that fails them is dropped rather
  // than carried: a diagram then draws flat, which is the honest degradation
  // §5 allows, instead of a forged payload reaching a writer.
  if (o.shapeShadow !== undefined && o.shapeShadow !== null) {
    const shadow = normalizeOfficeShapeShadow(o.shapeShadow);
    if (shadow) style.shapeShadow = shadow;
  }
  if (isSmartArtShape3d(o.shape3d)) style.shape3d = o.shape3d;
  // The display-only twin, held to the SAME shared guards. A value that fails
  // them is dropped and the shape draws flat, exactly as it did before this
  // field existed — never a forged payload reaching the canvas.
  if (o.nativeShape3d && typeof o.nativeShape3d === "object") {
    const native = o.nativeShape3d as Record<string, unknown>;
    if (isNativeShape3d(native.shape3d)) {
      style.nativeShape3d = {
        shape3d: native.shape3d,
        ...(isNativeScene3d(native.scene3d)
          ? { scene3d: native.scene3d }
          : {}),
      };
    }
  }
  if (isSmartArtScene3d(o.scene3d)) style.scene3d = o.scene3d;
  if (isSmartArtTextShape3d(o.textShape3d)) style.textShape3d = o.textShape3d;
  if (isSmartArtStyleId(o.smartArtStyle)) style.smartArtStyle = o.smartArtStyle;
  // A picture slot's picture, in the one inline spelling all three hosts write.
  // Dropped rather than carried when it is not that spelling or is past the
  // bound the API enforces: the slot then draws and exports empty, which is
  // what every sheet stored before this field does, instead of a payload no
  // writer could turn into a blip reaching the workbook.
  const pictureSrc = smartArtPictureSrc(o.src);
  if (pictureSrc) style.src = pictureSrc;
  // A picture fill the page authored is only ever its picture: without one
  // the marker says nothing and is dropped.
  if (pictureSrc && o.fillPicture === true) style.fillPicture = true;
  // Its options only ever ride with a picture, and only inside the model.
  if (pictureSrc && o.fillPictureOptions !== undefined) {
    const options = normalizeOfficePictureFillOptions(o.fillPictureOptions);
    if (options && officePictureFillOptionsStated(options)) style.fillPictureOptions = options;
  }
  return style;
}

/**
 * Is this object's picture the SHAPE'S OWN FILL rather than a picture put in a
 * SmartArt slot? An imported fill says so by naming the package part its
 * `a:blipFill` points at (`source.pictureFillPart`); one picked on the page by
 * `shape.fillPicture`. A fill paints the geometry (clipped to the outline,
 * stretched to the box); a slot's picture covers the slot. The fill IS the
 * picture only while the shape holds one: another fill kind takes `src` off.
 */
export function spreadsheetShapeHasPictureFill(object: unknown): boolean {
  const shape = (object as { shape?: unknown } | null)?.shape;
  if (!shape || typeof shape !== "object") return false;
  const { src, fillPicture } = shape as { src?: unknown; fillPicture?: unknown };
  if (typeof src !== "string" || !src) return false;
  if (fillPicture === true) return true;
  const source = (object as { source?: unknown } | null)?.source;
  if (!source || typeof source !== "object") return false;
  const part = (source as { pictureFillPart?: unknown }).pictureFillPart;
  return typeof part === "string" && part.length > 0;
}

/**
 * `style` with its picture fill taken off, as picking No Fill, a colour, a
 * gradient or a pattern does in Office (the new fill replaces the picture). A
 * SmartArt slot's picture is not a fill and stays.
 */
export function spreadsheetShapeStyleWithoutPictureFill(
  object: unknown,
  style: SpreadsheetShapeStyle,
): SpreadsheetShapeStyle {
  if (!spreadsheetShapeHasPictureFill(object)) return style;
  const next = { ...style };
  delete next.src;
  delete next.fillPicture;
  delete next.fillPictureOptions;
  return next;
}

/**
 * Build the default shape style for a freshly-inserted shape of `kind`.
 *
 * A text box is the one kind that starts invisible: Excel and WPS both insert
 * it with no fill and no outline, because what the user is placing is the text,
 * not a box around it. Every other kind keeps the shared blue default.
 */
export function createShapeStyle(
  kind: SpreadsheetPresetShapeKind,
): SpreadsheetShapeStyle {
  if (kind === "text-box") {
    return normalizeShapeStyle({ kind, fill: "none", stroke: "none" });
  }
  return normalizeShapeStyle({ kind });
}

/**
 * A drawing-shape sheet object ready to drop onto the grid (Excel/WPS
 * "Insert ▸ Shapes"). Mirrors the image-object shape (id + type + anchor) so it
 * flows through the SAME objects[] render / select / persist path; `shape`
 * carries the normalized style. Anchored "from" the given cell; the object layer
 * supplies the default size via objectPositionStyle, exactly like an image.
 */
export interface SpreadsheetShapeObject {
  id: string;
  type: "shape";
  shape: SpreadsheetShapeStyle;
  /**
   * The shared object anchor. The key is `column`, not `col`, and `kind` is
   * required: this is the shape the page schema validates, and a shape written
   * with any other shape fails the whole workbook's parse -- which the editor
   * turns into an EMPTY workbook, not a missing shape.
   */
  anchor:
    | {
        kind: "oneCell";
        from: { row: number; column: number };
      }
    | {
        kind: "absolute";
        xEmu: number;
        yEmu: number;
        widthEmu: number;
        heightEmu: number;
      };
}

/** Shape produced by the cell-anchored insertion factory. */
export interface SpreadsheetOneCellShapeObject extends SpreadsheetShapeObject {
  anchor: {
    kind: "oneCell";
    from: { row: number; column: number };
  };
}

let shapeObjectSeq = 0;

/** Build a new shape object of `kind` anchored at `anchorCell` (clamped ≥ 0). */
export function createSpreadsheetShapeObject(
  kind: SpreadsheetPresetShapeKind,
  anchorCell: { row: number; col: number },
): SpreadsheetOneCellShapeObject {
  shapeObjectSeq += 1;
  const row = Math.max(0, Math.trunc(anchorCell.row));
  const col = Math.max(0, Math.trunc(anchorCell.col));
  return {
    id: `shape-${shapeObjectSeq}-${Date.now().toString(36)}`,
    type: "shape",
    shape: createShapeStyle(kind),
    anchor: { kind: "oneCell", from: { row, column: col } },
  };
}

/**
 * One shape drawn with Lines > Curve / Freeform: Shape / Freeform: Scribble:
 * its own outline (the shared custom-path envelope) in an absolute box of
 * stage CSS pixels -- the same anchor a Draw-tab stroke takes. A closed figure
 * takes the fill it was given; an open one is a line and takes none.
 */
export function createSpreadsheetCustomShapeObject(
  result: {
    readonly left: number;
    readonly top: number;
    readonly width: number;
    readonly height: number;
    readonly customPath: unknown;
  },
  paint: { readonly fill: string; readonly stroke: string; readonly strokeWidth: number },
): SpreadsheetShapeObject | null {
  const customPath = normalizePresentationCustomPath(result.customPath);
  if (
    !customPath ||
    ![result.left, result.top, result.width, result.height].every(Number.isFinite) ||
    result.width <= 0 ||
    result.height <= 0
  ) {
    return null;
  }
  shapeObjectSeq += 1;
  return {
    id: `shape-${shapeObjectSeq}-${Date.now().toString(36)}`,
    type: "shape",
    shape: normalizeShapeStyle({
      kind: SPREADSHEET_CUSTOM_SHAPE_KIND,
      customPath,
      fill: customPath.closed ? paint.fill : "none",
      stroke: paint.stroke,
      strokeWidth: paint.strokeWidth,
    }),
    anchor: {
      kind: "absolute",
      xEmu: Math.max(0, Math.round(result.left * 9_525)),
      yEmu: Math.max(0, Math.round(result.top * 9_525)),
      widthEmu: Math.max(1, Math.round(result.width * 9_525)),
      heightEmu: Math.max(1, Math.round(result.height * 9_525)),
    },
  };
}

export interface CreateSpreadsheetInkShapeInput {
  /** Stage-local CSS-pixel samples, already clamped to the worksheet surface. */
  points: readonly SpreadsheetInkPoint[];
  tool: SpreadsheetInkTool;
  color: string;
  width: number;
  /** A pen's effect; `color` is then the effect's solid stand-in. */
  effect?: DocumentInkEffect;
  /** A piece a point eraser left: the ends it cut. */
  cutEnds?: InkCutEnds;
  /** XLS-A032: when each of `points` was drawn (ms, any origin). */
  times?: readonly number[];
}

/**
 * Convert one pointer gesture into one bounded absolute-anchored ink object.
 * The reused Office freeform helper owns the padded bounding-box math; this
 * adapter only translates CSS pixels to EMUs and writes Spreadsheet semantics.
 */
export function createSpreadsheetInkShapeObject(
  input: CreateSpreadsheetInkShapeInput,
): SpreadsheetShapeObject | null {
  const points = input.points.slice(0, MAX_SPREADSHEET_INK_POINTS);
  if (points.length === 0) return null;
  const safeWidth = Math.max(1, normalizeStrokeWidth(input.width));
  const emuPoints = points.map((point) => ({
    x: point.x * 9_525,
    y: point.y * 9_525,
  }));
  const padEmu = Math.max(2, safeWidth / 2 + 2) * 9_525;
  // A tap is a dot of the pen's tip (the shared `inkStrokeToShape`).
  const geometry = inkStrokeToShape(emuPoints, padEmu);
  if (!geometry) return null;
  const inkPoints = normalizeSpreadsheetInkPoints(geometry.points);
  if (!inkPoints) return null;
  // The pen's timing rides along only while the stored points are the drawn
  // ones, one for one (a tap's dot states two points of its own).
  const times = input.times?.slice(0, points.length);
  const inkTimes =
    times && times.length === points.length && parseSpreadsheetInkPoints(inkPoints).length === points.length
      ? times.map((time) => Math.max(0, Math.round(time - times[0]!))).join(" ")
      : undefined;
  // So does a pressure pen's reading: one width per stored point, and none
  // when the pen pressed evenly (or is a mouse). A highlighter's band is one
  // width in Excel, so its pressures are dropped.
  const inkPressure =
    input.tool !== "highlighter" && parseSpreadsheetInkPoints(inkPoints).length === points.length
      ? documentInkPressureOf(points)
      : undefined;
  shapeObjectSeq += 1;
  return {
    id: `ink-${shapeObjectSeq}-${Date.now().toString(36)}`,
    type: "shape",
    shape: normalizeShapeStyle({
      kind: SPREADSHEET_INK_SHAPE_KIND,
      fill: "none",
      stroke: input.color,
      strokeWidth: safeWidth,
      inkTool: input.tool,
      inkOpacity: spreadsheetInkOpacityForTool(input.tool),
      inkPoints,
      ...(inkTimes ? { inkTimes } : {}),
      ...(inkPressure ? { inkPressure } : {}),
      ...(input.tool === "pen" && input.effect ? { inkEffect: input.effect } : {}),
      ...(input.cutEnds ? { inkCutEnds: input.cutEnds } : {}),
    }),
    anchor: {
      kind: "absolute",
      xEmu: Math.max(0, Math.round(geometry.left)),
      yEmu: Math.max(0, Math.round(geometry.top)),
      widthEmu: Math.max(1, Math.round(geometry.width)),
      heightEmu: Math.max(1, Math.round(geometry.height)),
    },
  };
}

/**
 * XLS-A032: Draw ▸ Ink to Shape — the Office shape a stroke was fitted to
 * (`fitInkStrokeToShape`, stage CSS pixels), outlined in the stroke's own
 * colour and width, unfilled, in the same absolute box a stroke takes.
 */
export function createSpreadsheetShapeFromInkFit(
  fit: InkShapeFit,
  paint: { readonly color: string; readonly width: number; readonly tool: SpreadsheetInkTool },
): SpreadsheetShapeObject | null {
  // `kind` is not a unit type on any member, so the members are told apart
  // by the field only one of them has.
  const line = "start" in fit ? fit : null;
  const lineLike = line !== null;
  const box =
    "start" in fit
      ? {
          left: Math.min(fit.start.x, fit.end.x),
          top: Math.min(fit.start.y, fit.end.y),
          width: Math.abs(fit.end.x - fit.start.x),
          height: Math.abs(fit.end.y - fit.start.y),
        }
      : fit.box;
  if (![box.left, box.top, box.width, box.height].every(Number.isFinite)) return null;
  const flipH = line
    ? line.start.x > line.end.x
    : "flipH" in fit && fit.flipH === true;
  const flipV = line !== null && line.start.y > line.end.y;
  const customPoints =
    fit.kind === "polygon"
      ? normalizeSpreadsheetCustomPoints(
          fit.points
            .map(
              (point) =>
                `${formatInkCoordinate(box.width > 0 ? ((point.x - box.left) / box.width) * 100 : 50)},${formatInkCoordinate(
                  box.height > 0 ? ((point.y - box.top) / box.height) * 100 : 50,
                )}`,
            )
            .join(" "),
        )
      : undefined;
  if (fit.kind === "polygon" && !customPoints) return null;
  shapeObjectSeq += 1;
  return {
    id: `shape-${shapeObjectSeq}-${Date.now().toString(36)}`,
    type: "shape",
    shape: normalizeShapeStyle({
      kind: fit.kind === "polygon" ? SPREADSHEET_CUSTOM_SHAPE_KIND : fit.kind,
      ...(customPoints ? { customPoints } : {}),
      fill: "none",
      stroke: paint.color,
      strokeWidth: Math.max(1, normalizeStrokeWidth(paint.width)),
      // A highlighter's shape keeps the highlighter's see-through ink.
      ...(paint.tool === "highlighter" ? { opacity: Math.round(SPREADSHEET_HIGHLIGHTER_OPACITY * 100) } : {}),
    }),
    anchor: {
      kind: "absolute",
      xEmu: Math.max(0, Math.round(box.left * 9_525)),
      yEmu: Math.max(0, Math.round(box.top * 9_525)),
      widthEmu: Math.max(1, Math.round(box.width * 9_525)),
      heightEmu: Math.max(1, Math.round(box.height * 9_525)),
      ...(flipH ? { flipH: true } : {}),
      ...(flipV ? { flipV: true } : {}),
    } as SpreadsheetShapeObject["anchor"],
  };
}

/** A normalized style is usable Draw ink only when it has a real bounded path. */
export function isSpreadsheetInkShapeStyle(
  shape: SpreadsheetShapeStyle | undefined,
): shape is SpreadsheetShapeStyle & {
  kind: typeof SPREADSHEET_INK_SHAPE_KIND;
  inkTool: SpreadsheetInkTool;
  inkPoints: string;
  inkOpacity: number;
} {
  return (
    shape?.kind === SPREADSHEET_INK_SHAPE_KIND &&
    (shape.inkTool === "pen" ||
      shape.inkTool === "highlighter" ||
      shape.inkTool === "pencil") &&
    typeof shape.inkPoints === "string" &&
    parseSpreadsheetInkPoints(shape.inkPoints).length >= 2 &&
    typeof shape.inkOpacity === "number"
  );
}

/** A normalized style is a real custom outline only when it encloses an area. */
export function isSpreadsheetCustomShapeStyle(
  shape: SpreadsheetShapeStyle | undefined,
): shape is SpreadsheetShapeStyle & {
  kind: typeof SPREADSHEET_CUSTOM_SHAPE_KIND;
  customPoints: string;
} {
  return (
    shape?.kind === SPREADSHEET_CUSTOM_SHAPE_KIND &&
    typeof shape.customPoints === "string" &&
    parseSpreadsheetInkPoints(shape.customPoints).length >= 3
  );
}

/**
 * Repair an object anchor written before the shape factory emitted the schema's
 * own shape: `{ from: { row, col } }` with no `kind`.
 *
 * A workbook holding one of those fails the page schema outright, and the editor
 * reads a failed parse as zero sheets -- so a single old shape blanks the whole
 * workbook. Repairing before validation turns that back into a workbook.
 * Returns the input unchanged when there is nothing to repair.
 */
export function repairShapeObjectAnchor<T>(object: T): T {
  if (!object || typeof object !== "object") return object;
  const record = object as Record<string, unknown>;
  if (record.type !== "shape") return object;
  const anchor = record.anchor;
  if (!anchor || typeof anchor !== "object") return object;
  const anchorRecord = anchor as Record<string, unknown>;
  const from = anchorRecord.from;
  const fromRecord =
    from && typeof from === "object" ? (from as Record<string, unknown>) : null;
  const needsColumn =
    fromRecord !== null &&
    typeof fromRecord.column !== "number" &&
    typeof fromRecord.col === "number";
  const needsKind = typeof anchorRecord.kind !== "string";
  if (!needsColumn && !needsKind) return object;
  return {
    ...record,
    anchor: {
      ...anchorRecord,
      ...(needsKind ? { kind: "oneCell" } : {}),
      ...(fromRecord
        ? {
            from: {
              ...fromRecord,
              ...(needsColumn ? { column: fromRecord.col } : {}),
            },
          }
        : {}),
    },
  } as T;
}
