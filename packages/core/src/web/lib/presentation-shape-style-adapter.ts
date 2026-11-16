/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { officeMenuDashFields, officeMenuDashOf } from "./office-stroke-dash";
import { MAX_SHAPE_TEXT_LENGTH } from "./presentation-text-format";
import { withChangedLevelKeysOwn, type RunStyle } from "./presentation-runs";
import { presentationLevelStylesOf } from "./presentation-level-styles";
import type {
  PresentationShapeBlock,
  PresentationTextBlock,
} from "../components/console/editor/native-document/documentCommands";
import type {
  DeckTheme,
  PresentationThemeColorSlot,
} from "./presentation-theme";
import { normalizeHexColor } from "./office-color-palette";
import {
  OFFICE_OUTLINE_SHOWN_WEIGHT_PT,
  officeLineWeightPtToPx,
  officeOutlineWeightWhenShown,
} from "./office-line-weight";
import { nativePictureStyle } from "./office-picture-styles/native";
import { normalizeOfficePictureEffects } from "./office-picture-effects";
import { resolveImageEffects } from "./presentation-image-effects";
import { officeVideoStyle } from "./office-video-styles/catalog";
import {
  officeShapeStyleRecipes,
  type OfficeShapeStyleRecipe,
} from "./office-shape-style-catalog";
import { officeShapeStylePresetRecipes } from "./office-shape-style-presets";
import { presentationShapeCatalogEntry } from "./presentation-shape-catalog";
import { presentationChangeShapeAutofitFrame } from "./presentation-change-shape-fit";
import { officeDefaultShapeNameAfterChange } from "./office-shape-default-names";
import { presentationThemeGridColors } from "./presentation-theme-color-binding";
import {
  DEFAULT_SHAPE_STROKE,
  isLineCalloutShapeKind,
  isOpenBasicShapeKind,
  normalizeStrokeStyle,
  normalizeStrokeWidth,
  type ShapeKind,
  type StrokeStyle,
} from "./presentation-shapes";
import {
  DEFAULT_TEXTBOX_OUTLINE_WIDTH,
  persistedTextBoxFill,
  resolveTextBoxFill,
  TEXTBOX_FILL_KEYS,
} from "./presentation-text-fill";

/**
 * The receipt that turns an INHERITED layout/master placeholder into a
 * slide-level object, written by the first Shape Format edit that touches one.
 *
 * PowerPoint lets you format a placeholder that exists only on the layout: the
 * slide gains a `p:sp` carrying that placeholder's identity and nothing but the
 * properties you changed. Measured against real PowerPoint 2026 on this machine
 * (probe `pptph-materialised-title.pptx`, opened with no repair prompt, read
 * back as `shape type place holder` / `placeholder type center title
 * placeholder`, then saved as Open XML): the shape PowerPoint keeps is
 *
 *   <p:sp>
 *     <p:nvSpPr>
 *       <p:cNvPr id="2" name="Title 1"/>
 *       <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
 *       <p:nvPr><p:ph type="ctrTitle"/></p:nvPr>
 *     </p:nvSpPr>
 *     <p:spPr>            <- ONLY the changed paint. No a:xfrm, no a:prstGeom.
 *       <a:solidFill>…</a:solidFill><a:ln w="47625">…</a:ln>
 *     </p:spPr>
 *     <p:txBody><a:bodyPr/><a:lstStyle/><a:p>…</a:p></p:txBody>
 *   </p:sp>
 *
 * and the companion probe `pptph-inherited-bare.pptx` -- the same deck with an
 * EMPTY `spTree` -- came back from PowerPoint's own save with the `spTree`
 * still empty: an untouched inherited placeholder is not written. The receipt
 * therefore only ever appears on a block the user actually edited.
 *
 * `idx` is carried exactly as the file states it: the probe's layout `ctrTitle`
 * has no `idx`, PowerPoint wrote none, and our reader reports that as
 * `placeholderNativeIndexPresent: false` beside a DERIVED `placeholderIndex`.
 * Writing the derived value would invent an attribute the package never had.
 */
export const PRESENTATION_MATERIALIZED_PLACEHOLDER_SCHEMA =
  "pptx-materialized-placeholder-v1";

export interface PresentationMaterializedPlaceholder {
  readonly schemaVersion: typeof PRESENTATION_MATERIALIZED_PLACEHOLDER_SCHEMA;
  /** `p:ph@type`, exactly as the layout/master states it. */
  readonly placeholderType: string;
  /** `p:ph@idx`, present only when the inherited placeholder really had one. */
  readonly placeholderIndex?: string;
  /** Which part the projection came from, for the writer's own record. */
  readonly inheritedFrom: "layout" | "master";
}

function asSourceRecord(
  block: PresentationShapeBlock | PresentationTextBlock,
): Record<string, unknown> | undefined {
  const source = (block as Record<string, unknown>).source;
  return source && typeof source === "object" && !Array.isArray(source)
    ? (source as Record<string, unknown>)
    : undefined;
}

function trimmedString(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}

/**
 * The placeholder identity of a block that is still a READING of the layout or
 * the master, or `null` when the block is not one.
 *
 * A layout/master projection has no `shapeId` -- those ids live in another part
 * and collide with slide-local ones -- which is precisely why nothing can be
 * patched onto it until it is materialised.
 */
export function presentationInheritedPlaceholderIdentity(
  block: PresentationShapeBlock | PresentationTextBlock,
): PresentationMaterializedPlaceholder | null {
  const source = asSourceRecord(block);
  const origin = source?.inheritedFrom;
  if (origin !== "layout" && origin !== "master") return null;
  if (trimmedString(source?.shapeId)) return null;
  const record = block as Record<string, unknown>;
  const placeholderType =
    trimmedString(record.placeholderType) || trimmedString(source?.placeholderType);
  if (!placeholderType) return null;
  const nativeIndexPresent =
    source?.placeholderNativeIndexPresent === true ||
    trimmedString(record.placeholderNativeIndex).length > 0;
  const rawIndex = nativeIndexPresent
    ? trimmedString(record.placeholderIndex) ||
      trimmedString(source?.placeholderIndex) ||
      trimmedString(record.placeholderNativeIndex)
    : "";
  return {
    schemaVersion: PRESENTATION_MATERIALIZED_PLACEHOLDER_SCHEMA,
    placeholderType,
    ...(/^[0-9]+$/.test(rawIndex) ? { placeholderIndex: rawIndex } : {}),
    inheritedFrom: origin,
  };
}

/**
 * The block a Shape Format edit leaves behind when its target was still an
 * inherited placeholder: the same block, now a SLIDE-LEVEL object carrying the
 * placeholder identity, with `inheritedFrom` gone so every downstream reader
 * (stacking, selection pane, the two writers) treats it as what it now is.
 *
 * Idempotent, and a no-op for anything that is not an inherited placeholder --
 * so it can sit on the single producer of each patched block rather than being
 * spelled out at every command, which is how the two halves of a rule drift.
 */
export function presentationBlockWithMaterializedPlaceholder<
  TBlock extends PresentationShapeBlock | PresentationTextBlock,
>(block: TBlock): TBlock {
  const identity = presentationInheritedPlaceholderIdentity(block);
  if (!identity) return block;
  const source = { ...(asSourceRecord(block) ?? {}) };
  delete source.inheritedFrom;
  source.materializedPlaceholder = identity;
  return { ...block, source } as TBlock;
}

/**
 * The deck's whole Shape Styles library: the 6 x 7 theme styles, then the
 * gallery's measured Presets rows (`lib/office-shape-style-presets`), drawn
 * with the deck's own Text 1 / Background 1 / accents so a pick binds to the
 * theme slot it paints with.
 */
export function presentationShapeStyleRecipes(
  theme: DeckTheme | null | undefined,
): readonly OfficeShapeStyleRecipe[] {
  const grid = presentationThemeGridColors(theme);
  return [
    ...presentationThemeShapeStyleRecipes(theme),
    ...officeShapeStylePresetRecipes(
      [
        grid.tx1,
        grid.accent1,
        grid.accent2,
        grid.accent3,
        grid.accent4,
        grid.accent5,
        grid.accent6,
      ],
      grid.bg1,
    ),
  ];
}

function presentationThemeShapeStyleRecipes(
  theme: DeckTheme | null | undefined,
) {
  const sourceSlots =
    theme?.identity?.selectionKind === "native" &&
    theme.identity.variantId === "base"
      ? theme.identity.source?.colorSlots
      : undefined;
  return officeShapeStyleRecipes(
    Array.from(
      { length: 6 },
      (_, index) =>
        normalizeHexColor(
          sourceSlots?.[`accent${index + 1}` as PresentationThemeColorSlot],
        ) ??
        theme?.accentColors?.[index] ??
        "",
    ),
  );
}

export function presentationCanChangeShape(
  block: PresentationShapeBlock,
): boolean {
  const kind = (block.shape ?? "rectangle") as ShapeKind;
  const entry = presentationShapeCatalogEntry(kind);
  // A freeform (a custom outline, no ink points) changes shape too, as in
  
  const freeform = kind === "custom-geometry" && Boolean(block.customPath) && !block.points;
  if (
    !freeform &&
    (!entry ||
      entry.geometry !== "preset" ||
      entry.category === "action-buttons" ||
      block.customPath ||
      block.points)
  )
    return false;
  const source = block.source as Record<string, unknown> | undefined;
  return !source || source.nativeShapeConversion === "preset-default";
}

export function presentationCanChangeToShape(kind: ShapeKind): boolean {
  const entry = presentationShapeCatalogEntry(kind);
  return entry?.geometry === "preset" && entry.category !== "action-buttons";
}

export function presentationShapeConversionPatch(
  kind: ShapeKind,
  // The shape as it stands: an autofit one (`a:spAutoFit`) takes the frame
  // the new geometry's text rectangle needs (`presentationChangeShapeAutofitFrame`).
  block?: PresentationShapeBlock,
): Partial<PresentationShapeBlock> {
  const noFill = isLineCalloutShapeKind(kind) || isOpenBasicShapeKind(kind);
  const fitted = block && block.shape !== kind ? presentationChangeShapeAutofitFrame(block, kind) : null;
  const renamed = block ? presentationDefaultNameAfterShapeChange(block, kind) : null;
  return {
    shape: kind,
    ...(fitted ?? {}),
    ...(renamed !== null ? { name: renamed } : {}),
    points: undefined,
    customPath: undefined,
    ...(noFill
      ? {
          fill: undefined,
          fillTo: undefined,
          gradientAngle: undefined,
          fillImage: undefined,
          fillPattern: undefined,
        }
      : {}),
  };
}

/**
 * The name a shape takes from Change Shape: PowerPoint's own default name
 * ("Rectangle 1", by its cNvPr id) becomes the new shape's ("Oval 1"); any
 * other name stays (`officeDefaultShapeNameAfterChange`). A text frame keeps
 * its name ("TextBox 5" changed to an oval is still "TextBox 5"), so only a
 * shape's conversion asks. A shape not yet in the file (no `shpN`) has no id
 * to number it by: the writer names it.
 */
function presentationDefaultNameAfterShapeChange(
  block: PresentationShapeBlock,
  kind: ShapeKind,
): string | null {
  const shapeId = (block.source as { shapeId?: unknown } | undefined)?.shapeId;
  const number = typeof shapeId === "string" && /^shp[1-9][0-9]*$/.test(shapeId) ? Number(shapeId.slice(3)) : NaN;
  const preset = (shapeKind: unknown) => {
    const entry = typeof shapeKind === "string" ? presentationShapeCatalogEntry(shapeKind as ShapeKind) : null;
    return entry?.geometry === "preset" ? (entry as { drawingMlPreset: string }).drawingMlPreset : null;
  };
  return officeDefaultShapeNameAfterChange((block as { name?: unknown }).name, number, preset(block.shape), preset(kind));
}














export function presentationRunTakesShapeStyleColor(
  style: RunStyle | undefined,
  block: { readonly source?: unknown; readonly style?: { readonly color?: string } },
): boolean {
  if (!style?.color) return true;
  if (style.inheritedKeys?.includes("color")) return !style.colorFromList;
  if (style.recipeColor) return true;
  const native = typeof (block.source as { shapeId?: unknown } | undefined)?.shapeId === "string";
  return !native && style.color === block.style?.color;
}

/** Whether the box's own colour (the default of its words) takes the style's: not where a list style states its first level's colour. */
function boxTakesShapeStyleColor(block: unknown): boolean {
  return !presentationLevelStylesOf(block)?.[0]?.colorFromList;
}

/** `runs` with the style's colour on each run it replaces, marked as the style's. */
function runsWithShapeStyleColor<R extends { style?: RunStyle }>(
  runs: readonly R[],
  block: { readonly source?: unknown; readonly style?: { readonly color?: string } },
  color: string | undefined,
): R[] {
  return runs.map((run) =>
    presentationRunTakesShapeStyleColor(run.style, block)
      ? ({ ...run, style: withShapeStyleColor(run.style, color) } as R)
      : run,
  );
}

function withShapeStyleColor(style: RunStyle | undefined, color: string | undefined): RunStyle {
  const next: Record<string, unknown> = { ...(style ?? {}) };
  if (color) {
    next.color = color;
    next.recipeColor = true;
  } else {
    delete next.color;
    delete next.recipeColor;
  }
  delete next.colorFromList;
  // The colour is the run's now (written on it), no longer its level's.
  return withChangedLevelKeysOwn(style, next) as RunStyle;
}

/** Preserve each object's text and formatting except the style's text colour. */
export function presentationShapeStylePatch(
  block: PresentationShapeBlock,
  recipe: OfficeShapeStyleRecipe,
): Partial<PresentationShapeBlock> {
  return {
    fill: recipe.fill,
    // The Semitransparent presets' 50 %; every other style is opaque.
    fillTransparency: recipe.fillTransparency,
    fillGradient:recipe.fillGradient,
    shapeShadow:recipe.shapeShadow,
    fillTo: recipe.fillTo,
    gradientAngle: recipe.gradientAngle,
    fillImage: undefined,
    fillPattern: undefined,
    stroke: recipe.stroke,
    strokeTo: undefined,
    strokeGradientAngle: undefined,
    strokeWidth: recipe.strokeWidth,
    strokeStyle: "solid",
    shadow: recipe.shadow || undefined,
    style: boxTakesShapeStyleColor(block) ? { ...block.style, color: recipe.textColor } : block.style,
    ...(block.runs ? { runs: runsWithShapeStyleColor(block.runs, block, recipe.textColor) } : {}),
  };
}

/**
 * Whether Edit Shape > Change Shape may re-author THIS text frame's geometry.
 *
 * The same fail-closed rule autoshapes use: a frame whose geometry the package
 * owns (anything but our own preset default) keeps the source authoritative, so
 * the command greys rather than writing a geometry we cannot honour on the way
 * back out. A scratch / authored frame carries no native source and is free.
 */
export function presentationTextFrameCanChangeShape(
  block: PresentationTextBlock,
): boolean {
  const source = block.source as Record<string, unknown> | undefined;
  return !source || source.nativeShapeConversion === "preset-default";
}

/**
 * A text frame whose text a shape's label cannot hold whole: a shape's label
 * is bounded (`MAX_SHAPE_TEXT_LENGTH`, the save contract's bound too), and
 * Change Shape must refuse such a frame rather than cut its text at the
 * first edit (PPT-B268).
 */
export function presentationTextFrameLabelTooLong(
  block: PresentationTextBlock,
): boolean {
  const text = typeof block.text === "string" ? block.text : "";
  return text.replace(/\r\n?/g, "\n").length > MAX_SHAPE_TEXT_LENGTH;
}

/**
 * The shape block a text frame becomes when the user picks a geometry.
 *
 * PowerPoint's Edit Shape > Change Shape keeps "all the existing attributes of
 * the original shape (such as the color, border, or any text in the shape)"
 * (Microsoft Support, "Change a shape into another shape"). At the file-format
 * level the change is confined to `p:spPr` -- `a:prstGeom@prst` plus `a:avLst`
 * -- while `p:nvSpPr/p:nvPr/p:ph` and `p:txBody` are SIBLINGS of `p:spPr`
 * (ECMA-376 §19.3.1.44), so a geometry edit structurally cannot touch the
 * placeholder marker or the text body.
 *
 * The MODEL -- a text block becoming a shape block that carries the placeholder
 * identity, rather than a text block growing a private geometry field -- is the
 * one our own importer already uses. A text-bearing placeholder whose label the
 * text gate admits is projected as a `shape` block stamped with
 * `_pptx_placeholder_editor_identity`, under the standing note that "a
 * placeholder that becomes a SHAPE block is still the slide's title / subtitle
 * / content". The Docs and Spreadsheet editors model a text box as a shape too
 * (`documentShape` with `isTextBox`, `SpreadsheetShapeObject` with
 * `shape.kind === "text-box"`). A private geometry field would make this editor
 * emit a block its own importer never produces, and the block would change type
 * on the first download and re-upload.
 *
 * Everything but the paint's spelling carries: the id, the placeholder identity
 * (`placeholder` / `placeholderType` / `role`), the stacking order, the
 * geometry rect and the text. A text frame's `outlineColor` / `outlineWidth` /
 * `outlineStyle` are renamed to a shape's `stroke` / `strokeWidth` /
 * `strokeStyle`, and `kind: "textbox"` is dropped because the block is no
 * longer discriminated that way.
 */
export function presentationTextFrameToShapeBlock(
  block: PresentationTextBlock,
  kind: ShapeKind,
): PresentationShapeBlock {
  const next = { ...block } as Record<string, unknown>;
  const outline = {
    stroke: next.outlineColor,
    strokeTo: next.outlineColorTo,
    strokeGradientAngle: next.outlineGradientAngle,
    strokeWidth: next.outlineWidth,
    strokeStyle: next.outlineStyle,
  };
  for (const key of [
    "outlineColor",
    "outlineColorTo",
    "outlineGradientAngle",
    "outlineWidth",
    "outlineStyle",
    // The text-box discriminator. A shape block is not a `kind: "textbox"`.
    "kind",
  ]) {
    delete next[key];
  }
  for (const [key, value] of Object.entries(outline)) {
    if (value !== undefined) next[key] = value;
  }
  next.type = "shape";
  next.shape = kind;
  // A box that resizes to its words takes the height the new geometry's text
  // rectangle needs for them, as PowerPoint's Change Shape does (its width,
  // top edge, `txBox` and `bodyPr` stay).
  const fitted = presentationChangeShapeAutofitFrame(block, kind);
  if (fitted) Object.assign(next, fitted);
  if (isLineCalloutShapeKind(kind) || isOpenBasicShapeKind(kind)) {
    // An open shape draws no interior, so a fill carried over from the box
    // would be paint nothing renders -- the same clearing the autoshape
    // conversion does in `presentationShapeConversionPatch`.
    for (const key of ["fill", "fillTo", "gradientAngle"]) delete next[key];
  }
  return presentationBlockWithMaterializedPlaceholder(
    next as unknown as PresentationShapeBlock,
  );
}

/**
 * The same quick style, in the field names a TEXT FRAME persists.
 *
 * PowerPoint offers the Shape Styles gallery for a text placeholder and a plain
 * text box exactly as it does for an autoshape: Microsoft documents the Shape
 * Styles group, Shape Fill, Shape Outline and Shape Effects against "a shape or
 * text box" throughout, never against autoshapes alone (Microsoft Support, "Add
 * a fill or effect to a shape or text box"; "Add or remove a border on a text
 * box, shape, or SmartArt graphic").
 *
 * A text block spells its stroke `outlineColor` / `outlineWidth` /
 * `outlineStyle` where a shape spells it `stroke` / `strokeWidth` /
 * `strokeStyle`, so the recipe is translated rather than re-authored -- the
 * gallery, the accents and the seven colour slots stay the shared ones.
 *
 * Two of the recipe's richer paints have no home in the text-box model and are
 * deliberately REDUCED rather than invented: `fillGradient` (three stops) lands
 * as the two-stop `fill` -> `fillTo` the text box stores, and `shapeShadow` (an
 * exact offset/blur/opacity) lands as the boolean `shadow`. The matcher below
 * compares this same reduced form, so the gallery re-selects the row it wrote;
 * a reader that cannot read back its own write is the defect shape recorded in
 * SHAPE-IMPORT-AUDIT.md (the writer re-reads, disagrees with itself, and the
 * edit is refused).
 */
export function presentationTextFrameStyleFill(
  recipe: OfficeShapeStyleRecipe,
): ReturnType<typeof persistedTextBoxFill> {
  const hasOutline = recipe.stroke !== "none" && recipe.strokeWidth > 0;
  return persistedTextBoxFill({
    fill: recipe.fill,
    fillTo: recipe.fillTo,
    gradientAngle: recipe.gradientAngle,
    outlineColor: hasOutline ? recipe.stroke : undefined,
    outlineWidth: hasOutline ? recipe.strokeWidth : 0,
    outlineStyle: "solid",
    shadow: recipe.shadow,
  });
}

/** The fields of a text frame's line among `TEXTBOX_FILL_KEYS`. */
const TEXTBOX_OUTLINE_KEYS: ReadonlySet<string> = new Set([
  "outlineColor",
  "outlineColorTo",
  "outlineGradientAngle",
  "outlineWidth",
  "outlineStyle",
  "outlineAlign",
]);

/** Recolour every run a text frame carries -- flat and per paragraph. */
function textFrameRunsRecoloured(
  block: PresentationTextBlock,
  color: string | undefined,
): Partial<PresentationTextBlock> {
  const patch: Partial<PresentationTextBlock> = {};
  if (block.runs) patch.runs = runsWithShapeStyleColor(block.runs, block, color);
  if (block.paragraphs) {
    patch.paragraphs = block.paragraphs.map((paragraph) =>
      paragraph.runs ? { ...paragraph, runs: runsWithShapeStyleColor(paragraph.runs, block, color) } : paragraph,
    );
  }
  return patch;
}

/**
 * The text frame this recipe produces: its box paint, plus the text colour
 * PowerPoint carries in the style's `a:fontRef` child colour (ECMA-376
 * §19.3.1.46 `p:style`; `CT_FontReference` "specifies which themed font to use
 * along with a choice of color"). Our deck writer states a shape's quick style
 * as EXPLICIT `spPr` paint and explicit run colours rather than as `p:style`
 * matrix references -- see `presentationShapeStylePatch` above and the export's
 * `apply_text_box_frame` -- so a text frame is written the same way, and one
 * reader serves both.
 *
 * Returns a whole block rather than a patch because the fill keys must be
 * DELETED, not set to `undefined`: `persistedTextBoxFill` omits every field at
 * its default so an unstyled box stays byte-stable, and a lingering
 * `gradientAngle: undefined` key would defeat that. This mirrors
 * `updateTextBoxFill`, the Shape Fill / Shape Outline precedent in the same
 * ribbon group.
 */
export function presentationTextFrameWithStyle(
  block: PresentationTextBlock,
  recipe: OfficeShapeStyleRecipe,
): PresentationTextBlock {
  const paint = presentationTextFrameStyleFill(recipe);
  // A style with no line over a frame that draws none keeps the frame's own
  // spelling of "no line": an imported box reads `a:ln` with no fill as
  // `outlineColor: "none"`, width 0, and a theme change that re-applied its
  // style dropped both, a change of the box the save could not write in
  
  const ownLine = paint.outlineWidth === undefined && resolveTextBoxFill(block).outlineWidth === 0;
  const base = { ...block } as Record<string, unknown>;
  for (const key of TEXTBOX_FILL_KEYS) {
    if (!(ownLine && TEXTBOX_OUTLINE_KEYS.has(key))) delete base[key];
  }
  return presentationBlockWithMaterializedPlaceholder({
    ...base,
    ...paint,
    style: boxTakesShapeStyleColor(block) ? { ...block.style, color: recipe.textColor } : block.style,
    ...textFrameRunsRecoloured(block, recipe.textColor),
  } as unknown as PresentationTextBlock);
}

/** Whether the gallery should show `recipe` as this text frame's current row. */
export function presentationTextFrameMatchesStyle(
  block: PresentationTextBlock,
  recipe: OfficeShapeStyleRecipe,
): boolean {
  const written = presentationTextFrameStyleFill(recipe);
  const current = persistedTextBoxFill(block);
  const sameBox = TEXTBOX_FILL_KEYS.every(
    (key) =>
      (current as Record<string, unknown>)[key] ===
      (written as Record<string, unknown>)[key],
  );
  // The style's fill transparency is the box's too: a Semitransparent preset
  
  // "Semitransparent, Accent 3" where the old scheme's Accent 3 was white).
  if (!sameBox || ((current as { fillTransparency?: number }).fillTransparency ?? undefined) !== recipe.fillTransparency) return false;
  const paragraphRuns = (block.paragraphs ?? []).flatMap(
    (paragraph) => paragraph.runs ?? [],
  );
  const runs = [...(block.runs ?? []), ...paragraphRuns];
  // Words whose colour the style leaves (a list style's, the person's) say
  // nothing about which style the frame wears.
  const styled = runs.filter((run) => presentationRunTakesShapeStyleColor(run.style, block) || run.style?.recipeColor);
  return runs.length
    ? styled.every((run) => (run.style?.color ?? block.style?.color) === recipe.textColor)
    : !boxTakesShapeStyleColor(block) || block.style?.color === recipe.textColor;
}

/**
 * ONE Shape Format paint edit, stated once for every object the tab governs.
 *
 * PowerPoint applies Shape Fill, Shape Outline (colour, Weight, Dashes) and
 * Shape Effects to EVERY selected object, autoshapes and text boxes together --
 * Microsoft documents all four against "a shape or text box" and names no
 * per-object-type restriction (Support: "Add a fill or effect to a shape or
 * text box"; "Add or remove a border on a text box, shape, or SmartArt
 * graphic"; "Change the color, style, or weight of a line"). The two kinds of
 * block spell the same paint differently (`stroke` / `strokeWidth` /
 * `strokeStyle` against `outlineColor` / `outlineWidth` / `outlineStyle`), so
 * the COMMAND is stated in neutral terms here and translated per kind, exactly
 * as `presentationTextFrameStyleFill` translates a quick-style recipe. A second
 * vocabulary per kind is what lets the two drift apart.
 *
 * `null` means "no fill" / "no outline"; an ABSENT key is not part of the edit.
 */
export interface PresentationShapeObjectPaint {
  fill?: string | null;
  outlineColor?: string | null;
  outlineWidth?: number;
  outlineStyle?: StrokeStyle;
  /** Shape Outline > Dashes, one of Office's eight (`@/lib/office-stroke-dash`). */
  outlineDash?: import("./office-stroke-dash").OfficeMenuDash;
  /** Shape Outline > Sketched ("none" draws the true outline). */
  outlineSketch?: "none" | "curved" | "freehand" | "scribble";
  shadow?: boolean;
  /** Shape Effects preset beside the shadow (`lib/office-shape-effects`);
   * `null` clears it. Autoshapes only. */
  effect?: import("./office-shape-effects").OfficeShapeEffectPreset | null;
}

/** The same edit in a text frame's own persisted field names. */
export type PresentationTextFramePaintPatch = Partial<
  Pick<PresentationTextBlock, (typeof TEXTBOX_FILL_KEYS)[number]>
>;

/**
 * The weight an autoshape's line shows at once a colour, a dash or a line
 * type turns it on: its own, or Office's 0.75 pt when it has none (an
 * Intense Effect style leaves none). Empty when the weight stands.
 */
export function presentationShapeShownOutlineWidth(
  block: Pick<PresentationShapeBlock, "strokeWidth">,
): { strokeWidth?: number } {
  const width = normalizeStrokeWidth(block.strokeWidth);
  const shown = officeOutlineWeightWhenShown(
    width,
    officeLineWeightPtToPx(OFFICE_OUTLINE_SHOWN_WEIGHT_PT),
  );
  return shown === width ? {} : { strokeWidth: shown };
}

/**
 * The paint in an AUTOSHAPE's field names. Keys absent from the edit stay absent.
 *
 * Given the shape it lands on, a Weight or a Dash picked for a shape with No
 * Outline brings the line back in the default outline colour, as PowerPoint
 * does and as Home > Shape Outline already did (PPT-B180): otherwise the
 * Shape Format menu showed the weight checked on a shape that drew no line.
 */
export function presentationShapeObjectShapePatch(
  paint: PresentationShapeObjectPaint,
  block?: Pick<PresentationShapeBlock, "stroke" | "strokeWidth">,
): Partial<PresentationShapeBlock> {
  const patch: Record<string, unknown> = {};
  if (
    block?.stroke === "none" &&
    !("outlineColor" in paint) &&
    ((paint.outlineWidth !== undefined &&
      normalizeStrokeWidth(paint.outlineWidth) > 0) ||
      paint.outlineStyle !== undefined ||
      paint.outlineDash !== undefined ||
      (paint.outlineSketch !== undefined && paint.outlineSketch !== "none"))
  ) {
    patch.stroke = DEFAULT_SHAPE_STROKE;
  }
  // A colour or a dash turns the line on: a shape left with no weight (an
  // Intense Effect style) shows it at Office's 0.75 pt instead of drawing
  // nothing, as Excel measured does; a hidden line keeps its own weight.
  if (
    block &&
    paint.outlineWidth === undefined &&
    (("outlineColor" in paint && paint.outlineColor !== null) ||
      paint.outlineStyle !== undefined ||
      paint.outlineDash !== undefined ||
      (paint.outlineSketch !== undefined && paint.outlineSketch !== "none"))
  ) {
    Object.assign(patch, presentationShapeShownOutlineWidth(block));
  }
  if ("fill" in paint) {
    // A colour supersedes the whole gradient / picture / pattern fill, the
    // same way Home > Shape Fill restates it: a shape shows one fill type.
    patch.fill = paint.fill ?? "none";
    patch.fillTo = undefined;
    patch.gradientAngle = undefined;
    patch.fillGradient = undefined;
    patch.fillImage = undefined;
    patch.fillPattern = undefined;
  }
  if ("outlineColor" in paint) {
    patch.stroke = paint.outlineColor ?? "none";
    patch.strokeTo = undefined;
    patch.strokeGradientAngle = undefined;
  }
  if (paint.outlineWidth !== undefined) {
    patch.strokeWidth = normalizeStrokeWidth(paint.outlineWidth);
  }
  if (paint.outlineStyle !== undefined) {
    patch.strokeStyle = normalizeStrokeStyle(paint.outlineStyle);
    patch.outlineDash = undefined;
  }
  if (paint.outlineDash !== undefined) {
    const fields = officeMenuDashFields(paint.outlineDash);
    patch.strokeStyle = fields.style;
    patch.outlineDash = fields.dash;
    patch.outlineCap = fields.cap;
  }
  if (paint.outlineSketch !== undefined) {
    patch.outlineSketch = paint.outlineSketch === "none" ? undefined : paint.outlineSketch;
  }
  if (paint.shadow !== undefined) {
    patch.shadow = paint.shadow ? true : undefined;
  }
  if (paint.effect !== undefined) {
    patch.effect = paint.effect ?? undefined;
  }
  return patch as Partial<PresentationShapeBlock>;
}

/**
 * The same paint in a TEXT FRAME's field names, read against the frame's own
 * current outline.
 *
 * The frame model has no separate "line on / off": an outline exists exactly
 * while its width is positive, so a colour or a dash applied to a frame that
 * had no line also turns the line on at {@link DEFAULT_TEXTBOX_OUTLINE_WIDTH}
 * -- otherwise `persistedTextBoxFill` would drop the edit and the command would
 * light up and do nothing. Each frame is read individually, so a multi-selection
 * recoloured together KEEPS each frame's own weight.
 */
export function presentationShapeObjectTextPatch(
  block: PresentationTextBlock,
  paint: PresentationShapeObjectPaint,
): PresentationTextFramePaintPatch {
  const patch: Record<string, unknown> = {};
  const currentWidth = resolveTextBoxFill(block).outlineWidth;
  const widthForVisibleOutline = currentWidth || DEFAULT_TEXTBOX_OUTLINE_WIDTH;
  if ("fill" in paint) {
    patch.fill = paint.fill ?? undefined;
    patch.fillTo = undefined;
    patch.gradientAngle = undefined;
  }
  if ("outlineColor" in paint) {
    patch.outlineColor = paint.outlineColor ?? undefined;
    patch.outlineColorTo = undefined;
    patch.outlineGradientAngle = undefined;
    patch.outlineWidth =
      paint.outlineColor === null ? undefined : widthForVisibleOutline;
  }
  if (paint.outlineStyle !== undefined) {
    patch.outlineStyle = normalizeStrokeStyle(paint.outlineStyle);
    patch.outlineWidth = widthForVisibleOutline;
  }
  // A text frame keeps the three families: an Office dash lands as its family.
  if (paint.outlineDash !== undefined) {
    patch.outlineStyle = officeMenuDashFields(paint.outlineDash).style;
    patch.outlineWidth = widthForVisibleOutline;
  }
  // An explicit weight is the user's last word, above the two defaults above.
  if (paint.outlineWidth !== undefined) {
    patch.outlineWidth = normalizeStrokeWidth(paint.outlineWidth);
  }
  if (paint.shadow !== undefined) {
    patch.shadow = paint.shadow ? true : undefined;
  }
  return patch as PresentationTextFramePaintPatch;
}

/**
 * The outline WEIGHT and DASH style a Shape Format selection currently shows,
 * for either kind of object, so the Weight field and the Dashes radio group
 * read one SELECTION rather than one kind of block.
 */
export function presentationShapeObjectOutlineWidth(
  block: PresentationShapeBlock | PresentationTextBlock,
): number {
  return block.type === "text"
    ? resolveTextBoxFill(block as PresentationTextBlock).outlineWidth
    : normalizeStrokeWidth((block as PresentationShapeBlock).strokeWidth);
}

/** The Dashes entry a Shape Format selection shows checked (Office's eight). */
export function presentationShapeObjectOutlineMenuDash(
  block: PresentationShapeBlock | PresentationTextBlock,
): import("./office-stroke-dash").OfficeMenuDash | null {
  return block.type === "text"
    ? officeMenuDashOf(resolveTextBoxFill(block as PresentationTextBlock).outlineStyle, undefined)
    : officeMenuDashOf(
        normalizeStrokeStyle((block as PresentationShapeBlock).strokeStyle),
        (block as PresentationShapeBlock).outlineDash,
      );
}

export function presentationShapeObjectOutlineStyle(
  block: PresentationShapeBlock | PresentationTextBlock,
): StrokeStyle {
  return block.type === "text"
    ? resolveTextBoxFill(block as PresentationTextBlock).outlineStyle
    : normalizeStrokeStyle((block as PresentationShapeBlock).strokeStyle);
}

/**
 * Whether a slide block's paint may land OUTSIDE the frame the file states.
 *
 * DrawingML's `a:xfrm/a:ext` is the GEOMETRY box, not the ink box, and a
 * shape's outline is CENTRED on the geometry path: `a:ln@w` is the line width
 * in EMU (ECMA-376 20.1.2.2.24, ST_LineWidth), and Office draws half of it on
 * each side of the path. A preset whose path touches its own box -- an ellipse
 * at four points, a rectangle along four edges -- therefore puts half the line
 * width outside the frame, and PowerPoint clips a shape's own outline nowhere:
 * only the slide edge cuts ink. A frame that clips at its geometry box cut
 * every thick outline flat exactly at those touch points.
 *
 * Both slide frames -- the editor canvas block and the present overlay block
 * -- used to spell out their own list (custom path / imported native drawing /
 * line callout), so the ellipse a title was re-shaped into fell through both.
 * Every shape answers the same way, from one place.
 */
/** Whether a picture's own line (Picture Border) is drawn, by the rule the
 * picture draws it by (`resolveImageEffects`). */
function presentationPictureLineDrawn(block: { readonly borderWidth?: unknown; readonly borderStyle?: unknown }): boolean {
  return resolveImageEffects({ borderWidth: block.borderWidth, borderStyle: block.borderStyle }).borderWidth > 0;
}

export function presentationBlockPaintsOutsideFrame(
  block:
    | {
        readonly type?: unknown;
        readonly pictureStylePreset?: unknown;
        readonly pictureEffects?: unknown;
        readonly shadow?: unknown;
        readonly kind?: unknown;
        readonly videoStyle?: unknown;
        readonly borderWidth?: unknown;
        readonly borderStyle?: unknown;
      }
    | null
    | undefined,
): boolean {
  if (block?.type === "shape") return true;
  if (block?.type === "image") {
    // A picture style's reflection, glow or outer shadow is drawn below or
    // around the picture's frame: clipping the frame hid the reflection as
    // soon as the picture was deselected, and in the slide show (PPT-B244).
    const style = nativePictureStyle(block.pictureStylePreset);
    if (style?.effects?.some((effect) => ["reflection", "glow", "outerShdw"].includes(String(effect.kind)))) {
      return true;
    }
    // A flat picture's line lies outside it (`lib/office-picture-line`, as
    // PowerPoint draws it): a frame cutting at the picture's edge drew none of
    
    if (style?.line && style.line.width > 0) return true;
    if (presentationPictureLineDrawn(block)) return true;
    // Picture Format > Picture Effects (a shadow, a reflection, a glow, a
    // 3-D rotation turning the picture past its frame) and the older shadow
    // switch paint around the frame the same way. The effects that stay
    // inside it (soft edges, a bevel, an inner shadow) lose nothing by it.
    return normalizeOfficePictureEffects(block.pictureEffects) !== null || block.shadow === true;
  }
  // Video Format > Video Styles and Video Effects frame the movie as Picture
  // Effects frame a picture.
  if (block?.type === "embed" && block.kind === "video") {
    return officeVideoStyle(block.videoStyle) !== null || normalizeOfficePictureEffects(block.pictureEffects) !== null;
  }
  return false;
}

export function presentationShapeMatchesStyle(
  block: PresentationShapeBlock,
  recipe: OfficeShapeStyleRecipe,
): boolean {
  return (
    !block.fillImage &&
    !block.fillPattern &&
    !block.strokeTo &&
    (block.strokeStyle ?? "solid") === "solid" &&
    block.fill === recipe.fill &&
    (block.fillTransparency ?? undefined) === recipe.fillTransparency &&
    JSON.stringify(block.fillGradient)===JSON.stringify(recipe.fillGradient) &&
    JSON.stringify(block.shapeShadow)===JSON.stringify(recipe.shapeShadow) &&
    (block.fillTo ?? undefined) === recipe.fillTo &&
    (block.gradientAngle ?? undefined) === recipe.gradientAngle &&
    block.stroke === recipe.stroke &&
    (recipe.stroke === "none" ||
      Math.abs((block.strokeWidth ?? 2) - recipe.strokeWidth) < 0.01) &&
    Boolean(block.shadow) === recipe.shadow &&
    (block.runs?.length
      ? block.runs
          .filter((run) => presentationRunTakesShapeStyleColor(run.style, block) || run.style?.recipeColor)
          .every((run) => (run.style?.color ?? block.style?.color) === recipe.textColor)
      : !boxTakesShapeStyleColor(block) || block.style?.color === recipe.textColor)
  );
}
