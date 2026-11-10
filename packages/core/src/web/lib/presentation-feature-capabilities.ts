/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  readFeatureCommandBindingManifest,
  type FeatureCommandBindingContract,
  type FeatureCommandBindingManifest,
  type FeatureCommandBindingRegistry,
  type FeatureCommandRegistryLineage,
} from "../components/office-ribbon/featureCommandRegistry";
import {
  PRESENTATION_NATIVE_AUTHORABLE_ANIMATION_EFFECTS,
  PRESENTATION_WEB_ANIMATION_EFFECTS,
} from "./presentation-anim-timeline";
import { PRESENTATION_GEOMETRY_MAX_EMU } from "./presentation-geometry-limit";

export type PresentationRibbonTabId =
  | "home"
  | "insert"
  | "draw"
  | "design"
  | "transitions"
  | "animations"
  | "slideshow"
  | "record"
  | "review"
  | "view"
  | "shape-format"
  | "camera-format"
  | "image-format"
  | "graphic-format"
  | "chart-design"
  | "chart-format"
  | "media-format"
  | "media-playback"
  | "smartart-design"
  | "smartart-format"
  | "table-design"
  | "table-layout";

export interface PresentationCapabilityContext {
  editable: boolean;
  /** Any block selected on the active slide. */
  hasSelection: boolean;
  /** The selection resolves to a TEXT block the ribbon can format. */
  hasTextSelection: boolean;
  /** The selection resolves to a SHAPE block (fill/outline/effects targets). */
  hasShapeSelection: boolean;
  /**
   * The selection resolves to words the paragraph commands format: a text
   * box's, or a shape's (PowerPoint's shape text is the same text body).
   * Absent: only a text box's (`hasTextSelection`).
   */
  hasParagraphSelection?: boolean;
  /** The block clipboard holds content (paste family). */
  hasClipboard: boolean;
  online: boolean;
}

export interface PresentationFeatureCapability {
  id: string;
  tabs: readonly PresentationRibbonTabId[];
  entry: "ribbon" | "menu" | "both";
  authoring: "native" | "projection" | "external" | "unsupported";
  pptxImport: "native" | "preserved" | "collapsed" | "none";
  fromScratchExport: "native" | "declared-loss" | "reject";
  patchbackExport: "native" | "declared-loss" | "reject";
  powerpoint: "editable" | "visible" | "unsupported";
  wps: "editable" | "visible" | "unsupported";
  /** Stable UI command/translation key; placement lives in `entry`, not here. */
  commandKey: string;
  /** Closed predicate id evaluated by `isPresentationFeatureAvailable`. */
  contextPredicate?:
    | "always"
    | "editable"
    | "selection"
    | "textSelection"
    | "paragraphSelection"
    | "shapeSelection"
    | "clipboard"
    | "online";
  /** Translation key explaining a disabled/contextual state. */
  reasonKey?: string;
}

/**
 * Versioned, evidence-bearing capability contract used by new presentation
 * slices. The legacy array below remains a compatibility view until each of
 * its 46 rows has real handler/domain/writer evidence; absence is represented
 * as `unknown`, never upgraded to `native` by a helper default.
 */
export const PRESENTATION_LEGACY_CAPABILITY_CATALOG_VERSION =
  "legacy-catalog-v1" as const;
export const PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION =
  "presentation-capability-catalog-v2" as const;

export type PresentationCapabilitySourceVersion =
  | typeof PRESENTATION_LEGACY_CAPABILITY_CATALOG_VERSION
  | typeof PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION
  | "unregistered";

export type PresentationCapabilityContextPredicate = NonNullable<
  PresentationFeatureCapability["contextPredicate"]
>;

export interface PresentationCapabilityOfficeBaseline {
  readonly platform: string | null;
  readonly version: string | null;
  readonly evidenceRefs: readonly string[];
}

export interface PresentationCapabilitySurfaceContract {
  readonly disposition: "cataloged" | "unknown";
  readonly tabs: readonly PresentationRibbonTabId[];
  readonly entry: PresentationFeatureCapability["entry"] | "unknown";
  readonly group: string | null;
  readonly order: number | null;
  readonly contextualPredicate:
    PresentationCapabilityContextPredicate | "unknown";
  readonly controlKind: string | null;
}

export interface PresentationCapabilityCommandContract {
  readonly disposition:
    "bound" | "parameterized" | "external" | "unsupported" | "unknown";
  readonly key: string | null;
  readonly handlerRef: string | null;
  readonly undoPolicy:
    "single-transaction" | "view-only" | "external" | "unknown";
  readonly availabilityPredicate:
    PresentationCapabilityContextPredicate | "unknown";
}

export interface PresentationCapabilitySemanticsContract {
  readonly defaults: Readonly<Record<string, unknown>> | "unknown";
  readonly limits: Readonly<Record<string, unknown>> | "unknown";
  readonly interactions: readonly string[];
  readonly shortcutIds: readonly string[];
}

export interface PresentationCapabilityDomainContract {
  readonly disposition: "typed" | "view-only" | "external" | "unknown";
  readonly schemaPaths: readonly string[];
  readonly version: string | null;
  readonly persistence: "document" | "session" | "external" | "unknown";
}

export interface PresentationCapabilityImportContract {
  readonly disposition: PresentationFeatureCapability["pptxImport"] | "unknown";
  readonly projectionRefs: readonly string[];
  readonly preservedUnknownParts: boolean | "unknown";
  readonly lossCodes: readonly string[];
  readonly flags: readonly string[];
}

export interface PresentationCapabilityExportContract {
  readonly disposition:
    PresentationFeatureCapability["fromScratchExport"] | "unknown";
  readonly serializerRefs: readonly string[];
  readonly lossCodes: readonly string[];
  readonly flags: readonly string[];
}

export interface PresentationCapabilityPatchbackContract extends PresentationCapabilityExportContract {
  readonly preservePolicy: string | null;
}

export interface PresentationCapabilityOfficeOutcomeContract {
  readonly powerpoint: PresentationFeatureCapability["powerpoint"] | "unknown";
  readonly wps: PresentationFeatureCapability["wps"] | "unknown";
}

export interface PresentationFeatureCapabilityContract {
  readonly featureId: string;
  readonly sourceVersion: PresentationCapabilitySourceVersion;
  readonly officeBaseline: PresentationCapabilityOfficeBaseline;
  readonly surface: PresentationCapabilitySurfaceContract;
  readonly command: PresentationCapabilityCommandContract;
  readonly authoring: {
    readonly disposition:
      PresentationFeatureCapability["authoring"] | "unknown";
  };
  readonly semantics: PresentationCapabilitySemanticsContract;
  readonly domain: PresentationCapabilityDomainContract;
  readonly import: PresentationCapabilityImportContract;
  readonly fromScratchExport: PresentationCapabilityExportContract;
  readonly patchbackExport: PresentationCapabilityPatchbackContract;
  readonly officeOutcome: PresentationCapabilityOfficeOutcomeContract;
  readonly acceptanceContractIds: readonly string[];
  readonly owner: string | null;
}

export interface PresentationFeatureCapabilityCatalog {
  readonly schemaVersion: typeof PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION;
  readonly sourceVersion: Exclude<
    PresentationCapabilitySourceVersion,
    "unregistered"
  >;
  readonly features: readonly PresentationFeatureCapabilityContract[];
  readonly featureIds: readonly string[];
  /** True only when a caller supplied and passed the complete handler set. */
  readonly handlerBindingsVerified: boolean;
}

export interface ReadPresentationFeatureCapabilityCatalogOptions {
  /** Reserved for a separately leased production command-registry adapter. */
  readonly handlerFeatureIds?: never;
}

const validatedPresentationFeatureCapabilityCatalogs =
  new WeakSet<PresentationFeatureCapabilityCatalog>();

function leaf(
  capability: PresentationFeatureCapability,
): PresentationFeatureCapability {
  return Object.freeze(capability);
}

/** A fully-native Home command: editor + both exports + Office all first-class. */
const nativeHome = (
  id: string,
  commandKey: string,
  overrides: Partial<PresentationFeatureCapability> = {},
): PresentationFeatureCapability =>
  leaf({
    id,
    tabs: ["home"],
    entry: "ribbon",
    authoring: "native",
    pptxImport: "native",
    fromScratchExport: "native",
    patchbackExport: "native",
    powerpoint: "editable",
    wps: "editable",
    commandKey,
    ...overrides,
  });

/** Editor-only state (never serialised): selection, panes, tool arming. */
const editorOnly = (
  id: string,
  commandKey: string,
  overrides: Partial<PresentationFeatureCapability> = {},
): PresentationFeatureCapability =>
  leaf({
    id,
    tabs: ["home"],
    entry: "ribbon",
    authoring: "native",
    pptxImport: "none",
    fromScratchExport: "native",
    patchbackExport: "native",
    powerpoint: "editable",
    wps: "editable",
    commandKey,
    ...overrides,
  });

export const PRESENTATION_FEATURE_CAPABILITY_REGISTRY: readonly PresentationFeatureCapability[] =
  Object.freeze([
    // ── 剪贴板 ──────────────────────────────────────────────────────────
    editorOnly("clipboard.paste", "presentation.command.paste", {
      contextPredicate: "editable",
    }),
    editorOnly(
      "clipboard.pasteUseDestinationTheme",
      "presentation.command.pasteUseDestinationTheme",
      {
        entry: "menu",
        contextPredicate: "clipboard",
        reasonKey: "presentation.capability.emptyClipboard",
      },
    ),
    editorOnly(
      "clipboard.pasteKeepSource",
      "presentation.command.pasteKeepSource",
      {
        entry: "menu",
        contextPredicate: "clipboard",
        reasonKey: "presentation.capability.emptyClipboard",
      },
    ),
    editorOnly(
      "clipboard.pastePicture",
      "presentation.command.pastePicture",
      {
        entry: "menu",
        contextPredicate: "clipboard",
        reasonKey: "presentation.capability.emptyClipboard",
      },
    ),
    editorOnly(
      "clipboard.pasteTextOnly",
      "presentation.command.pasteTextOnly",
      {
        entry: "menu",
        contextPredicate: "clipboard",
        reasonKey: "presentation.capability.emptyClipboard",
      },
    ),
    editorOnly("clipboard.cut", "presentation.command.cut", {
      contextPredicate: "selection",
      reasonKey: "presentation.capability.selectionRequired",
    }),
    editorOnly("clipboard.copy", "presentation.command.copy", {
      contextPredicate: "selection",
      reasonKey: "presentation.capability.selectionRequired",
    }),
    editorOnly(
      "clipboard.formatPainter",
      "presentation.command.formatPainter",
      {
        contextPredicate: "selection",
        reasonKey: "presentation.capability.selectionRequired",
      },
    ),
    editorOnly("editing.deleteSelection", "presentation.command.delete", {
      contextPredicate: "selection",
      reasonKey: "presentation.capability.deleteSelectionRequired",
    }),
    // ── 幻灯片 ──────────────────────────────────────────────────────────
    nativeHome("slides.new", "presentation.command.newSlide", {
      tabs: ["home", "insert"],
    }),
    nativeHome("slides.duplicate", "presentation.command.duplicateSlide", {
      entry: "menu",
    }),
    nativeHome("slides.layout", "presentation.command.layout", {
      entry: "both",
    }),
    nativeHome("slides.reset", "presentation.command.resetSlide"),
    nativeHome("slides.section", "presentation.command.section", {
      entry: "both",
    }),
    nativeHome("slides.fromOutline", "presentation.command.fromOutline", {
      entry: "menu",
    }),
    leaf({
      id: "slides.reuse",
      tabs: ["home"],
      entry: "menu",
      authoring: "native",
      pptxImport: "native",
      fromScratchExport: "reject",
      patchbackExport: "native",
      powerpoint: "editable",
      wps: "editable",
      commandKey: "presentation.command.reuseSlides",
      contextPredicate: "online",
      reasonKey: "presentation.capability.offline",
    }),
    leaf({
      id: "slides.setTransition",
      tabs: ["transitions"],
      entry: "ribbon",
      authoring: "native",
      pptxImport: "native",
      fromScratchExport: "native",
      patchbackExport: "native",
      powerpoint: "editable",
      wps: "unsupported",
      commandKey: "presentation.command.setSlideTransition",
      contextPredicate: "editable",
      reasonKey: "presentation.capability.selectionRequired",
    }),
    leaf({
      id: "slides.setTransitionDuration",
      tabs: ["transitions"],
      entry: "ribbon",
      authoring: "native",
      pptxImport: "native",
      fromScratchExport: "native",
      patchbackExport: "native",
      powerpoint: "editable",
      wps: "unsupported",
      commandKey: "presentation.command.setTransitionDuration",
      contextPredicate: "editable",
      reasonKey: "presentation.capability.selectionRequired",
    }),
    leaf({
      id: "slides.setTransitionAdvance",
      tabs: ["transitions"],
      entry: "ribbon",
      authoring: "native",
      pptxImport: "native",
      fromScratchExport: "native",
      patchbackExport: "native",
      powerpoint: "editable",
      wps: "unsupported",
      commandKey: "presentation.command.setTransitionAdvance",
      contextPredicate: "editable",
      reasonKey: "presentation.capability.selectionRequired",
    }),
    leaf({
      id: "slides.applyTransitionToAll",
      tabs: ["transitions"],
      entry: "ribbon",
      authoring: "native",
      pptxImport: "native",
      fromScratchExport: "native",
      patchbackExport: "native",
      powerpoint: "editable",
      wps: "unsupported",
      commandKey: "presentation.command.applyTransitionToAll",
      contextPredicate: "editable",
      reasonKey: "presentation.capability.selectionRequired",
    }),
    // ── 字体 ────────────────────────────────────────────────────────────
    nativeHome("font.clearFormatting", "presentation.command.clearFormatting", {
      contextPredicate: "textSelection",
      reasonKey: "presentation.capability.textRequired",
      pptxImport: "none",
    }),
    nativeHome("font.textShadow", "presentation.command.textShadow", {
      contextPredicate: "textSelection",
      reasonKey: "presentation.capability.textRequired",
    }),
    nativeHome("font.changeCase", "presentation.command.changeCase", {
      entry: "menu",
      contextPredicate: "paragraphSelection",
      reasonKey: "presentation.capability.textRequired",
      pptxImport: "none",
    }),
    nativeHome("font.letterSpacing", "presentation.command.letterSpacing", {
      entry: "menu",
      contextPredicate: "paragraphSelection",
      reasonKey: "presentation.capability.textRequired",
    }),
    // ── 段落 ────────────────────────────────────────────────────────────
    nativeHome("paragraph.bullets", "presentation.command.bullets", {
      contextPredicate: "paragraphSelection",
      reasonKey: "presentation.capability.textRequired",
    }),
    nativeHome("paragraph.numbering", "presentation.command.numbering", {
      contextPredicate: "paragraphSelection",
      reasonKey: "presentation.capability.textRequired",
    }),
    nativeHome("paragraph.indent", "presentation.command.indent", {
      contextPredicate: "paragraphSelection",
      reasonKey: "presentation.capability.listRequired",
    }),
    nativeHome("paragraph.lineSpacing", "presentation.command.lineSpacing", {
      entry: "menu",
      contextPredicate: "paragraphSelection",
      reasonKey: "presentation.capability.textRequired",
    }),
    nativeHome("paragraph.columns", "presentation.command.columns", {
      entry: "menu",
      contextPredicate: "paragraphSelection",
      reasonKey: "presentation.capability.textRequired",
    }),
    nativeHome(
      "paragraph.textDirection",
      "presentation.command.textDirection",
      {
        entry: "menu",
        contextPredicate: "paragraphSelection",
        reasonKey: "presentation.capability.textRequired",
      },
    ),
    nativeHome("paragraph.alignText", "presentation.command.alignText", {
      entry: "menu",
      contextPredicate: "textSelection",
      reasonKey: "presentation.capability.textRequired",
    }),
    nativeHome("paragraph.distribute", "presentation.command.distribute", {
      contextPredicate: "paragraphSelection",
      reasonKey: "presentation.capability.textRequired",
    }),
    nativeHome("paragraph.toSmartArt", "presentation.command.toSmartArt", {
      entry: "menu",
      contextPredicate: "textSelection",
      reasonKey: "presentation.capability.textWithContentRequired",
      pptxImport: "collapsed",
    }),
    // ── 绘图(Home 组) ──────────────────────────────────────────────────
    nativeHome("drawing.image", "presentation.command.insertImage", {
      tabs: ["home", "insert", "graphic-format"],
    }),
    leaf({
      id: "picture.altText",
      tabs: ["image-format", "graphic-format"],
      entry: "ribbon",
      authoring: "native",
      pptxImport: "native",
      fromScratchExport: "native",
      patchbackExport: "native",
      powerpoint: "editable",
      wps: "editable",
      commandKey: "presentation.command.editPictureAltText",
      contextPredicate: "selection",
      reasonKey: "presentation.capability.pictureSelectionRequired",
    }),
    nativeHome("drawing.shape", "presentation.command.insertShape", {
      entry: "both",
    }),
    nativeHome("drawing.textBox", "presentation.command.insertTextBox"),
    editorOnly("drawing.arrange", "presentation.command.arrange", {
      entry: "menu",
      contextPredicate: "selection",
      reasonKey: "presentation.capability.selectionRequired",
    }),
    nativeHome("drawing.quickStyles", "presentation.command.quickStyles", {
      entry: "menu",
      contextPredicate: "shapeSelection",
      reasonKey: "presentation.capability.shapeRequired",
    }),
    nativeHome("drawing.shapeFill", "presentation.command.shapeFill", {
      contextPredicate: "shapeSelection",
      reasonKey: "presentation.capability.shapeRequired",
    }),
    nativeHome("drawing.shapeOutline", "presentation.command.shapeOutline", {
      contextPredicate: "shapeSelection",
      reasonKey: "presentation.capability.shapeRequired",
    }),
    nativeHome("drawing.shapeEffects", "presentation.command.shapeEffects", {
      entry: "menu",
      contextPredicate: "shapeSelection",
      reasonKey: "presentation.capability.shapeRequired",
    }),
    // ── 编辑 ────────────────────────────────────────────────────────────
    editorOnly("editing.findReplace", "presentation.command.findReplace"),
    editorOnly("editing.selectAll", "presentation.command.selectAll", {
      entry: "menu",
    }),
    editorOnly("editing.selectObjects", "presentation.command.selectObjects", {
      entry: "menu",
    }),
    nativeHome("editing.selectionPane", "presentation.command.selectionPane", {
      entry: "menu",
      contextPredicate: "editable",
      reasonKey: "presentation.selectionPane.unavailable",
    }),
    // No command or surface exists for deck sensitivity. Keep the catalog row
    // as an explicit unsupported sentinel instead of inventing Home authority.
    leaf({
      id: "deck.sensitivity",
      tabs: [],
      entry: "menu",
      authoring: "unsupported",
      pptxImport: "none",
      fromScratchExport: "reject",
      patchbackExport: "reject",
      powerpoint: "unsupported",
      wps: "unsupported",
      commandKey: "presentation.command.sensitivity",
    }),
    // ── 设计器 ────────────────────────────────────────────────
    leaf({
      id: "deck.designer",
      tabs: ["design"],
      entry: "ribbon",
      authoring: "external",
      pptxImport: "none",
      // The editor's layout engine makes every idea; the external service
      // only ranks them, and without it the engine's own order shows. An
      // applied idea is committed through the ordinary native slide-object
      // writers, so the result round-trips in both package paths.
      fromScratchExport: "native",
      patchbackExport: "native",
      powerpoint: "editable",
      wps: "unsupported",
      commandKey: "presentation.command.designer",
      contextPredicate: "editable",
      reasonKey: "presentation.designIdeas.unavailable",
    }),
    // ── 绘图 tab(墨迹) ─────────────────────────────────────────────────
    leaf({
      id: "ink.drawMode",
      tabs: ["draw"],
      entry: "ribbon",
      authoring: "native",
      pptxImport: "native",
      fromScratchExport: "native",
      patchbackExport: "native",
      powerpoint: "editable",
      wps: "editable",
      commandKey: "presentation.command.inkDrawMode",
      contextPredicate: "editable",
    }),
    leaf({
      id: "ink.pen",
      tabs: ["draw"],
      entry: "ribbon",
      authoring: "native",
      pptxImport: "native",
      fromScratchExport: "native",
      patchbackExport: "native",
      powerpoint: "editable",
      wps: "editable",
      commandKey: "presentation.command.inkPen",
      contextPredicate: "editable",
    }),
    leaf({
      id: "ink.highlighter",
      tabs: ["draw"],
      entry: "ribbon",
      authoring: "native",
      pptxImport: "native",
      fromScratchExport: "native",
      patchbackExport: "native",
      powerpoint: "editable",
      wps: "editable",
      commandKey: "presentation.command.inkHighlighter",
      contextPredicate: "editable",
    }),
    leaf({
      id: "ink.erase",
      tabs: ["draw"],
      entry: "ribbon",
      authoring: "native",
      pptxImport: "none",
      fromScratchExport: "native",
      patchbackExport: "native",
      powerpoint: "editable",
      wps: "editable",
      commandKey: "presentation.command.inkErase",
      contextPredicate: "editable",
    }),
    leaf({
      id: "ink.trackpad",
      tabs: ["draw"],
      entry: "ribbon",
      authoring: "projection",
      pptxImport: "none",
      fromScratchExport: "reject",
      patchbackExport: "reject",
      powerpoint: "visible",
      wps: "unsupported",
      commandKey: "presentation.command.inkTrackpad",
      contextPredicate: "editable",
    }),
    // ── Slide Show tab ──────────────────────────────────────────────────
    // These are session/view commands. They intentionally carry no PPTX
    // mutation claim: the deck is byte-identical before and after starting a
    // show or opening Presenter View.
    editorOnly("slideshow.fromStart", "presentation.command.presentFromStart", {
      tabs: ["slideshow"],
      entry: "ribbon",
      authoring: "projection",
      fromScratchExport: "reject",
      patchbackExport: "reject",
      powerpoint: "visible",
      wps: "unsupported",
      contextPredicate: "always",
    }),
    editorOnly(
      "slideshow.fromCurrent",
      "presentation.command.presentFromCurrent",
      {
        tabs: ["slideshow"],
        entry: "both",
        authoring: "projection",
        fromScratchExport: "reject",
        patchbackExport: "reject",
        powerpoint: "visible",
        wps: "unsupported",
        contextPredicate: "always",
      },
    ),
    editorOnly("slideshow.presenterView", "presentation.command.openPresenterView", {
      tabs: ["slideshow"],
      contextPredicate: "always",
    }),
    editorOnly("slideshow.rehearse", "presentation.command.rehearse", {
      tabs: ["slideshow"],
      contextPredicate: "always",
    }),
    leaf({
      id: "record.narration",
      tabs: ["slideshow"],
      entry: "ribbon",
      authoring: "native",
      pptxImport: "none",
      // The recording lives on the slide; the package declares the drop until
      // media embedding lands (loss code slide_narration_not_exported).
      fromScratchExport: "declared-loss",
      patchbackExport: "declared-loss",
      powerpoint: "unsupported",
      wps: "unsupported",
      commandKey: "presentation.command.narration",
      contextPredicate: "editable",
      reasonKey: "presentation.capability.narrationUnsupported",
    }),
  ]);

export const PRESENTATION_TAB_IDS = Object.freeze([
  "home",
  "insert",
  "draw",
  "design",
  "transitions",
  "animations",
  "slideshow",
  "record",
  "review",
  "view",
  "shape-format",
  "camera-format",
  "image-format",
  "graphic-format",
  "chart-design",
  "chart-format",
  "media-format",
  "media-playback",
  "smartart-design",
  "smartart-format",
  "table-design",
  "table-layout",
] as const satisfies readonly PresentationRibbonTabId[]);

export interface PresentationProductionCommandEvidence {
  readonly featureId: string;
  readonly group: string;
  readonly action: string;
  readonly handlerRef: string;
  readonly controlId: string;
  readonly persistence: "document" | "session" | "external";
  readonly undoPolicy:
    | "single-transaction"
    | "view-only"
    | "external";
}

function productionCommandEvidence(
  featureId: string,
  group: string,
  handlerRef: string,
  persistence: PresentationProductionCommandEvidence["persistence"] =
    "document",
  undoPolicy: PresentationProductionCommandEvidence["undoPolicy"] =
    persistence === "document" ? "single-transaction" : "view-only",
): PresentationProductionCommandEvidence {
  return Object.freeze({
    featureId,
    group,
    action: `presentation.web.${featureId}`,
    handlerRef: `PresentationWindow#${handlerRef}`,
    controlId: `PPTCTL.WEB.${featureId.replace(/[^A-Za-z0-9]+/g, ".").toUpperCase()}`,
    persistence,
    undoPolicy,
  });
}

/**
 * Complete evidence ceiling for the commands installed by the production
 * PresentationWindow registry. The legacy catalog remains readable for old
 * receipts and tests, but it can no longer grant a live handler: every bound
 * command must have one of these explicit Web-guide identities and is migrated
 * to a strict v2 row before registry construction.
 */
export const PRESENTATION_PRODUCTION_COMMAND_EVIDENCE = Object.freeze([
  productionCommandEvidence("clipboard.paste", "clipboard", "pasteFromLiveClipboard"),
  productionCommandEvidence(
    "clipboard.pasteUseDestinationTheme",
    "clipboard",
    "pasteBlocksUseTheme",
  ),
  productionCommandEvidence(
    "clipboard.pasteKeepSource",
    "clipboard",
    "pasteBlocks",
  ),
  productionCommandEvidence(
    "clipboard.pastePicture",
    "clipboard",
    "pasteBlocksAsPicture",
  ),
  productionCommandEvidence(
    "clipboard.pasteTextOnly",
    "clipboard",
    "pasteBlocksTextOnly",
  ),
  productionCommandEvidence("clipboard.cut", "clipboard", "cutBlocks"),
  productionCommandEvidence("clipboard.copy", "clipboard", "copyBlocks"),
  productionCommandEvidence(
    "clipboard.formatPainter",
    "clipboard",
    "toggleFormatPainterSingle",
  ),
  productionCommandEvidence(
    "editing.deleteSelection",
    "slides",
    "deleteLiveHomeSelection",
  ),
  productionCommandEvidence("slides.new", "slides", "addSlideFromLayout"),
  productionCommandEvidence("slides.layout", "slides", "applyLayoutToActiveSlide"),
  productionCommandEvidence("slides.duplicate", "slides", "duplicateSlide"),
  productionCommandEvidence("slides.reset", "slides", "resetActiveSlide"),
  productionCommandEvidence("slides.section", "slides", "createSectionAtSlide"),
  productionCommandEvidence(
    "slides.fromOutline",
    "slides",
    "openOutlineImport",
    "session",
  ),
  productionCommandEvidence("slides.reuse", "slides", "openReuseSlides"),
  productionCommandEvidence("drawing.image", "drawing", "insertImageFile"),
  productionCommandEvidence("drawing.textBox", "drawing", "insertTextBox"),
  productionCommandEvidence("drawing.shape", "drawing", "insertShapeCatalogSelection"),
  productionCommandEvidence("paragraph.bullets", "paragraph", "toggleLiveRibbonListBullet"),
  productionCommandEvidence("paragraph.numbering", "paragraph", "toggleLiveRibbonListNumber"),
  productionCommandEvidence("paragraph.distribute", "paragraph", "toggleLiveRibbonDistribute"),
  productionCommandEvidence("font.clearFormatting", "font", "clearLiveRibbonTextFormatting"),
  productionCommandEvidence("font.textShadow", "font", "toggleLiveRibbonTextShadow"),
  productionCommandEvidence(
    "editing.findReplace",
    "editing",
    "openFindReplacePanel",
    "session",
  ),
  productionCommandEvidence("editing.selectAll", "editing", "selectAllBlocks", "session"),
  productionCommandEvidence(
    "editing.selectObjects",
    "editing",
    "activateSelectObjects",
    "session",
  ),
  productionCommandEvidence(
    "editing.selectionPane",
    "editing",
    "toggleSelectionPane",
    "session",
  ),
  productionCommandEvidence(
    "deck.designer",
    "designer",
    "toggleDesignIdeas",
    "external",
    "external",
  ),
  productionCommandEvidence(
    "slideshow.fromStart",
    "slide-show",
    "enterPresentModeFromStart",
    "session",
  ),
  productionCommandEvidence(
    "slideshow.fromCurrent",
    "slide-show",
    "enterPresentModeFromCurrent",
    "session",
  ),
  productionCommandEvidence(
    "slideshow.rehearse",
    "slide-show",
    "startRehearsing",
    "session",
  ),
  productionCommandEvidence("ink.drawMode", "draw", "writeInkInteractionMode", "session"),
  productionCommandEvidence("ink.erase", "draw", "eraseInk", "session"),
  productionCommandEvidence("ink.trackpad", "draw", "writeInkTrackpadEnabled", "session"),
  productionCommandEvidence("ink.pen", "draw", "writeInkPen", "session"),
  productionCommandEvidence("ink.highlighter", "draw", "writeInkHighlighter", "session"),
]);

const PRESENTATION_PRODUCTION_COMMAND_EVIDENCE_BY_ID = new Map(
  PRESENTATION_PRODUCTION_COMMAND_EVIDENCE.map((entry) => [
    entry.featureId,
    entry,
  ] as const),
);

export function presentationProductionCommandEvidence(
  featureId: string,
): PresentationProductionCommandEvidence | null {
  return PRESENTATION_PRODUCTION_COMMAND_EVIDENCE_BY_ID.get(featureId) ?? null;
}

const CONTEXT_PREDICATES = Object.freeze([
  "always",
  "editable",
  "selection",
  "textSelection",
  "paragraphSelection",
  "shapeSelection",
  "clipboard",
  "online",
] as const satisfies readonly PresentationCapabilityContextPredicate[]);

const ENTRY_DISPOSITIONS = ["ribbon", "menu", "both", "unknown"] as const;
const SURFACE_DISPOSITIONS = ["cataloged", "unknown"] as const;
const COMMAND_DISPOSITIONS = [
  "bound",
  "parameterized",
  "external",
  "unsupported",
  "unknown",
] as const;
const UNDO_POLICIES = [
  "single-transaction",
  "view-only",
  "external",
  "unknown",
] as const;
const AUTHORING_DISPOSITIONS = [
  "native",
  "projection",
  "external",
  "unsupported",
  "unknown",
] as const;
const DOMAIN_DISPOSITIONS = [
  "typed",
  "view-only",
  "external",
  "unknown",
] as const;
const PERSISTENCE_DISPOSITIONS = [
  "document",
  "session",
  "external",
  "unknown",
] as const;
const IMPORT_DISPOSITIONS = [
  "native",
  "preserved",
  "collapsed",
  "none",
  "unknown",
] as const;
const EXPORT_DISPOSITIONS = [
  "native",
  "declared-loss",
  "reject",
  "unknown",
] as const;
const OFFICE_OUTCOME_DISPOSITIONS = [
  "editable",
  "visible",
  "unsupported",
  "unknown",
] as const;

type DispositionValue<T extends readonly (string | boolean)[]> = T[number];

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function requireRecord(value: unknown, label: string): Record<string, unknown> {
  if (!isRecord(value)) {
    throw new Error(`Missing presentation capability ${label}`);
  }
  return value;
}

function requireString(value: unknown, label: string): string {
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`Missing presentation capability ${label}`);
  }
  return value;
}

function requireNullableString(value: unknown, label: string): string | null {
  if (value === null) return null;
  return requireString(value, label);
}

function requireNullableInteger(value: unknown, label: string): number | null {
  if (value === null) return null;
  if (!Number.isInteger(value) || (value as number) < 0) {
    throw new Error(`Invalid presentation capability ${label}`);
  }
  return value as number;
}

function requireStringList(value: unknown, label: string): readonly string[] {
  if (
    !Array.isArray(value) ||
    value.some((item) => typeof item !== "string" || item.trim() === "")
  ) {
    throw new Error(`Invalid presentation capability ${label}`);
  }
  return Object.freeze([...value]) as readonly string[];
}

function requireDisposition<T extends readonly (string | boolean)[]>(
  value: unknown,
  allowed: T,
  label: string,
): DispositionValue<T> {
  if (value === undefined) {
    throw new Error(`Missing presentation capability ${label} disposition`);
  }
  if (!(allowed as readonly unknown[]).includes(value)) {
    throw new Error(
      `Undeclared presentation capability ${label} disposition: ${String(value)}`,
    );
  }
  return value as DispositionValue<T>;
}

function legacyDisposition<T extends readonly string[]>(
  value: unknown,
  allowed: T,
): T[number] | "unknown" {
  return typeof value === "string" && allowed.includes(value)
    ? (value as T[number])
    : "unknown";
}

function requireTabs(
  value: unknown,
  label: string,
): readonly PresentationRibbonTabId[] {
  if (
    !Array.isArray(value) ||
    value.some(
      (tab) =>
        typeof tab !== "string" ||
        !PRESENTATION_TAB_IDS.includes(tab as PresentationRibbonTabId),
    )
  ) {
    throw new Error(`Invalid presentation capability ${label}`);
  }
  return Object.freeze([...value]) as readonly PresentationRibbonTabId[];
}

function legacyTabs(value: unknown): readonly PresentationRibbonTabId[] {
  if (!Array.isArray(value)) return Object.freeze([]);
  const tabs = value.filter(
    (tab): tab is PresentationRibbonTabId =>
      typeof tab === "string" &&
      PRESENTATION_TAB_IDS.includes(tab as PresentationRibbonTabId),
  );
  return tabs.length === value.length ? Object.freeze(tabs) : Object.freeze([]);
}

function requirePredicate(
  value: unknown,
  label: string,
): PresentationCapabilityContextPredicate | "unknown" {
  return requireDisposition(
    value,
    [...CONTEXT_PREDICATES, "unknown"] as const,
    label,
  );
}

function freezeUnknownRecord(
  value: unknown,
  label: string,
): Readonly<Record<string, unknown>> | "unknown" {
  if (value === "unknown") return value;
  if (!isRecord(value)) {
    throw new Error(`Invalid presentation capability ${label}`);
  }
  return Object.freeze({ ...value });
}

function normalizeLegacyCapability(
  value: unknown,
  index: number,
): PresentationFeatureCapabilityContract {
  const source = requireRecord(value, `features[${index}]`);
  const featureId = requireString(source.id, `features[${index}].id`);
  const tabs = legacyTabs(source.tabs);
  const entry = legacyDisposition(source.entry, ENTRY_DISPOSITIONS);
  const predicate = legacyDisposition(
    source.contextPredicate ?? "always",
    CONTEXT_PREDICATES,
  );
  const commandKey =
    typeof source.commandKey === "string" && source.commandKey.trim() !== ""
      ? source.commandKey
      : null;
  const surfaceDisposition =
    tabs.length > 0 && entry !== "unknown" ? "cataloged" : "unknown";

  return Object.freeze({
    featureId,
    sourceVersion: PRESENTATION_LEGACY_CAPABILITY_CATALOG_VERSION,
    officeBaseline: Object.freeze({
      platform: null,
      version: null,
      evidenceRefs: Object.freeze([]),
    }),
    surface: Object.freeze({
      disposition: surfaceDisposition,
      tabs,
      entry,
      group: null,
      order: null,
      contextualPredicate: predicate,
      controlKind: null,
    }),
    command: Object.freeze({
      // A translation/command key is not proof that a live handler exists.
      disposition: "unknown",
      key: commandKey,
      handlerRef: null,
      undoPolicy: "unknown",
      availabilityPredicate: predicate,
    }),
    authoring: Object.freeze({
      disposition: legacyDisposition(source.authoring, AUTHORING_DISPOSITIONS),
    }),
    semantics: Object.freeze({
      defaults: "unknown",
      limits: "unknown",
      interactions: Object.freeze([]),
      shortcutIds: Object.freeze([]),
    }),
    domain: Object.freeze({
      disposition: "unknown",
      schemaPaths: Object.freeze([]),
      version: null,
      persistence: "unknown",
    }),
    import: Object.freeze({
      disposition: legacyDisposition(source.pptxImport, IMPORT_DISPOSITIONS),
      projectionRefs: Object.freeze([]),
      preservedUnknownParts: "unknown",
      lossCodes: Object.freeze([]),
      flags: Object.freeze([]),
    }),
    fromScratchExport: Object.freeze({
      disposition: legacyDisposition(
        source.fromScratchExport,
        EXPORT_DISPOSITIONS,
      ),
      serializerRefs: Object.freeze([]),
      lossCodes: Object.freeze([]),
      flags: Object.freeze([]),
    }),
    patchbackExport: Object.freeze({
      disposition: legacyDisposition(
        source.patchbackExport,
        EXPORT_DISPOSITIONS,
      ),
      serializerRefs: Object.freeze([]),
      preservePolicy: null,
      lossCodes: Object.freeze([]),
      flags: Object.freeze([]),
    }),
    officeOutcome: Object.freeze({
      powerpoint: legacyDisposition(
        source.powerpoint,
        OFFICE_OUTCOME_DISPOSITIONS,
      ),
      wps: legacyDisposition(source.wps, OFFICE_OUTCOME_DISPOSITIONS),
    }),
    acceptanceContractIds: Object.freeze([]),
    owner: null,
  });
}

/**
 * The evidence half of a v2 row: where the behaviour is documented and
 * implemented (Office baseline, schema, projection and writer references,
 * loss codes, flags, acceptance ids, owner). It is audit truth: the ribbon
 * decides nothing from it, so the browser's catalogs do not carry it. It
 * lives in `./presentation-feature-capability-audit` (server / tests only),
 * keyed by feature id, and `readAuditedPresentationFeatureCapabilityCatalog`
 * reads a row with it whole.
 */
const PRESENTATION_CAPABILITY_AUDIT_LEAVES = Object.freeze({
  officeBaseline: null,
  acceptanceContractIds: null,
  owner: null,
  semantics: ["interactions"],
  domain: ["schemaPaths", "version"],
  import: ["projectionRefs", "preservedUnknownParts", "lossCodes", "flags"],
  fromScratchExport: ["serializerRefs", "lossCodes", "flags"],
  patchbackExport: ["serializerRefs", "preservePolicy", "lossCodes", "flags"],
} as const);

const NEUTRAL_AUDIT_LEAVES: Readonly<Record<string, unknown>> = Object.freeze({
  officeBaseline: { platform: null, version: null, evidenceRefs: [] },
  acceptanceContractIds: [],
  owner: null,
  "semantics.interactions": [],
  "domain.schemaPaths": [],
  "domain.version": null,
  "import.projectionRefs": [],
  "import.preservedUnknownParts": "unknown",
  "import.lossCodes": [],
  "import.flags": [],
  "fromScratchExport.serializerRefs": [],
  "fromScratchExport.lossCodes": [],
  "fromScratchExport.flags": [],
  "patchbackExport.serializerRefs": [],
  "patchbackExport.preservePolicy": null,
  "patchbackExport.lossCodes": [],
  "patchbackExport.flags": [],
});

/** The audit leaf paths a runtime row must not carry (`officeBaseline`, `import.lossCodes`, ...). */
export const PRESENTATION_CAPABILITY_AUDIT_LEAF_PATHS: readonly string[] =
  Object.freeze(Object.keys(NEUTRAL_AUDIT_LEAVES));

/**
 * A runtime row with its audit leaves set to the neutral values: none of them
 * may be present (a row that carries evidence belongs in the audit module).
 */
function withNeutralAuditLeaves(
  source: Record<string, unknown>,
  featureId: string,
): Record<string, unknown> {
  const result: Record<string, unknown> = { ...source };
  for (const [key, children] of Object.entries(
    PRESENTATION_CAPABILITY_AUDIT_LEAVES,
  )) {
    if (children === null) {
      if (key in source) {
        throw new Error(
          `Audit truth in a runtime presentation capability: ${featureId}.${key}`,
        );
      }
      result[key] = NEUTRAL_AUDIT_LEAVES[key];
      continue;
    }
    const owner = isRecord(source[key]) ? { ...source[key] } : source[key];
    if (!isRecord(owner)) continue;
    for (const child of children) {
      if (child in owner) {
        throw new Error(
          `Audit truth in a runtime presentation capability: ${featureId}.${key}.${child}`,
        );
      }
      owner[child] = NEUTRAL_AUDIT_LEAVES[`${key}.${child}`];
    }
    result[key] = owner;
  }
  return result;
}

function normalizeCurrentCapability(
  value: unknown,
  index: number,
  audited: boolean,
): PresentationFeatureCapabilityContract {
  const row = requireRecord(value, `features[${index}]`);
  const featureId = requireString(
    row.featureId,
    `features[${index}].featureId`,
  );
  // The browser's rows carry no evidence; their checks of it run where the
  // evidence is (the audited read, in the audit module's tests).
  const source = audited ? row : withNeutralAuditLeaves(row, featureId);
  const officeBaseline = requireRecord(
    source.officeBaseline,
    `${featureId}.officeBaseline`,
  );
  const surface = requireRecord(source.surface, `${featureId}.surface`);
  const command = requireRecord(source.command, `${featureId}.command`);
  const authoring = requireRecord(source.authoring, `${featureId}.authoring`);
  const semantics = requireRecord(source.semantics, `${featureId}.semantics`);
  const domain = requireRecord(source.domain, `${featureId}.domain`);
  const importContract = requireRecord(source.import, `${featureId}.import`);
  const fromScratchExport = requireRecord(
    source.fromScratchExport,
    `${featureId}.fromScratchExport`,
  );
  const patchbackExport = requireRecord(
    source.patchbackExport,
    `${featureId}.patchbackExport`,
  );
  const officeOutcome = requireRecord(
    source.officeOutcome,
    `${featureId}.officeOutcome`,
  );

  const surfaceDisposition = requireDisposition(
    surface.disposition,
    SURFACE_DISPOSITIONS,
    `${featureId}.surface`,
  );
  const tabs = requireTabs(surface.tabs, `${featureId}.surface.tabs`);
  const entry = requireDisposition(
    surface.entry,
    ENTRY_DISPOSITIONS,
    `${featureId}.surface.entry`,
  );
  if (
    surfaceDisposition === "cataloged" &&
    (tabs.length === 0 || entry === "unknown")
  ) {
    throw new Error(
      `Cataloged presentation feature lacks a surface: ${featureId}`,
    );
  }

  const commandDisposition = requireDisposition(
    command.disposition,
    COMMAND_DISPOSITIONS,
    `${featureId}.command`,
  );
  const commandKey = requireNullableString(
    command.key,
    `${featureId}.command.key`,
  );
  const handlerRef = requireNullableString(
    command.handlerRef,
    `${featureId}.command.handlerRef`,
  );
  if (
    (commandDisposition === "bound" ||
      commandDisposition === "parameterized") &&
    (!commandKey || !handlerRef)
  ) {
    throw new Error(
      `Executable presentation command lacks a handler: ${featureId}`,
    );
  }

  const domainDisposition = requireDisposition(
    domain.disposition,
    DOMAIN_DISPOSITIONS,
    `${featureId}.domain`,
  );
  const schemaPaths = requireStringList(
    domain.schemaPaths,
    `${featureId}.domain.schemaPaths`,
  );
  if (audited && domainDisposition === "typed" && schemaPaths.length === 0) {
    throw new Error(`Typed presentation feature lacks a schema: ${featureId}`);
  }

  const normalizedImport = Object.freeze({
    disposition: requireDisposition(
      importContract.disposition,
      IMPORT_DISPOSITIONS,
      `${featureId}.import`,
    ),
    projectionRefs: requireStringList(
      importContract.projectionRefs,
      `${featureId}.import.projectionRefs`,
    ),
    preservedUnknownParts: requireDisposition(
      importContract.preservedUnknownParts,
      [true, false, "unknown"] as const,
      `${featureId}.import.preservedUnknownParts`,
    ),
    lossCodes: requireStringList(
      importContract.lossCodes,
      `${featureId}.import.lossCodes`,
    ),
    flags: requireStringList(importContract.flags, `${featureId}.import.flags`),
  });
  if (
    audited &&
    normalizedImport.disposition === "native" &&
    normalizedImport.projectionRefs.length === 0
  ) {
    throw new Error(
      `Native presentation import lacks a projection: ${featureId}`,
    );
  }
  if (
    normalizedImport.preservedUnknownParts === false &&
    normalizedImport.lossCodes.length === 0
  ) {
    throw new Error(
      `Lossy presentation import lacks a loss code: ${featureId}`,
    );
  }

  const normalizedFromScratch = Object.freeze({
    disposition: requireDisposition(
      fromScratchExport.disposition,
      EXPORT_DISPOSITIONS,
      `${featureId}.fromScratchExport`,
    ),
    serializerRefs: requireStringList(
      fromScratchExport.serializerRefs,
      `${featureId}.fromScratchExport.serializerRefs`,
    ),
    lossCodes: requireStringList(
      fromScratchExport.lossCodes,
      `${featureId}.fromScratchExport.lossCodes`,
    ),
    flags: requireStringList(
      fromScratchExport.flags,
      `${featureId}.fromScratchExport.flags`,
    ),
  });
  const normalizedPatchback = Object.freeze({
    disposition: requireDisposition(
      patchbackExport.disposition,
      EXPORT_DISPOSITIONS,
      `${featureId}.patchbackExport`,
    ),
    serializerRefs: requireStringList(
      patchbackExport.serializerRefs,
      `${featureId}.patchbackExport.serializerRefs`,
    ),
    preservePolicy: requireNullableString(
      patchbackExport.preservePolicy,
      `${featureId}.patchbackExport.preservePolicy`,
    ),
    lossCodes: requireStringList(
      patchbackExport.lossCodes,
      `${featureId}.patchbackExport.lossCodes`,
    ),
    flags: requireStringList(
      patchbackExport.flags,
      `${featureId}.patchbackExport.flags`,
    ),
  });

  for (const [route, contract] of audited
    ? ([
        ["fromScratchExport", normalizedFromScratch],
        ["patchbackExport", normalizedPatchback],
      ] as const)
    : []) {
    if (
      contract.disposition === "native" &&
      contract.serializerRefs.length === 0
    ) {
      throw new Error(
        `Native presentation ${route} lacks a serializer: ${featureId}`,
      );
    }
    if (contract.disposition === "native" && contract.lossCodes.length > 0) {
      throw new Error(
        `Native presentation ${route} declares loss: ${featureId}`,
      );
    }
    if (
      contract.disposition === "declared-loss" &&
      contract.serializerRefs.length === 0
    ) {
      throw new Error(
        `Lossy presentation ${route} lacks a serializer: ${featureId}`,
      );
    }
    if (
      contract.disposition === "declared-loss" &&
      contract.lossCodes.length === 0
    ) {
      throw new Error(
        `Lossy presentation ${route} lacks a loss code: ${featureId}`,
      );
    }
  }
  if (
    audited &&
    (normalizedPatchback.disposition === "native" ||
      normalizedPatchback.disposition === "declared-loss") &&
    !normalizedPatchback.preservePolicy
  ) {
    throw new Error(
      `Writable presentation patchback lacks preserve policy: ${featureId}`,
    );
  }

  return Object.freeze({
    featureId,
    sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    officeBaseline: Object.freeze({
      platform: requireNullableString(
        officeBaseline.platform,
        `${featureId}.officeBaseline.platform`,
      ),
      version: requireNullableString(
        officeBaseline.version,
        `${featureId}.officeBaseline.version`,
      ),
      evidenceRefs: requireStringList(
        officeBaseline.evidenceRefs,
        `${featureId}.officeBaseline.evidenceRefs`,
      ),
    }),
    surface: Object.freeze({
      disposition: surfaceDisposition,
      tabs,
      entry,
      group: requireNullableString(surface.group, `${featureId}.surface.group`),
      order: requireNullableInteger(
        surface.order,
        `${featureId}.surface.order`,
      ),
      contextualPredicate: requirePredicate(
        surface.contextualPredicate,
        `${featureId}.surface.contextualPredicate`,
      ),
      controlKind: requireNullableString(
        surface.controlKind,
        `${featureId}.surface.controlKind`,
      ),
    }),
    command: Object.freeze({
      disposition: commandDisposition,
      key: commandKey,
      handlerRef,
      undoPolicy: requireDisposition(
        command.undoPolicy,
        UNDO_POLICIES,
        `${featureId}.command.undoPolicy`,
      ),
      availabilityPredicate: requirePredicate(
        command.availabilityPredicate,
        `${featureId}.command.availabilityPredicate`,
      ),
    }),
    authoring: Object.freeze({
      disposition: requireDisposition(
        authoring.disposition,
        AUTHORING_DISPOSITIONS,
        `${featureId}.authoring`,
      ),
    }),
    semantics: Object.freeze({
      defaults: freezeUnknownRecord(
        semantics.defaults,
        `${featureId}.semantics.defaults`,
      ),
      limits: freezeUnknownRecord(
        semantics.limits,
        `${featureId}.semantics.limits`,
      ),
      interactions: requireStringList(
        semantics.interactions,
        `${featureId}.semantics.interactions`,
      ),
      shortcutIds: requireStringList(
        semantics.shortcutIds,
        `${featureId}.semantics.shortcutIds`,
      ),
    }),
    domain: Object.freeze({
      disposition: domainDisposition,
      schemaPaths,
      version: requireNullableString(
        domain.version,
        `${featureId}.domain.version`,
      ),
      persistence: requireDisposition(
        domain.persistence,
        PERSISTENCE_DISPOSITIONS,
        `${featureId}.domain.persistence`,
      ),
    }),
    import: normalizedImport,
    fromScratchExport: normalizedFromScratch,
    patchbackExport: normalizedPatchback,
    officeOutcome: Object.freeze({
      powerpoint: requireDisposition(
        officeOutcome.powerpoint,
        OFFICE_OUTCOME_DISPOSITIONS,
        `${featureId}.officeOutcome.powerpoint`,
      ),
      wps: requireDisposition(
        officeOutcome.wps,
        OFFICE_OUTCOME_DISPOSITIONS,
        `${featureId}.officeOutcome.wps`,
      ),
    }),
    acceptanceContractIds: requireStringList(
      source.acceptanceContractIds,
      `${featureId}.acceptanceContractIds`,
    ),
    owner: requireNullableString(source.owner, `${featureId}.owner`),
  });
}

/**
 * Read either the frozen legacy-v1 table or a strict v2 catalog. Legacy rows
 * keep their old public values, while every dimension legacy never recorded is
 * explicitly normalized to `unknown`. V2 omissions are structural errors.
 */
export function readPresentationFeatureCapabilityCatalog(
  input: unknown,
  options: ReadPresentationFeatureCapabilityCatalogOptions = {},
): PresentationFeatureCapabilityCatalog {
  return readCatalog(input, options, false);
}

/**
 * The same read of a v2 catalog whose rows carry their evidence (the audit
 * module merges it in by feature id): every evidence field is required and
 * checked against the dispositions it backs.
 */
export function readAuditedPresentationFeatureCapabilityCatalog(
  input: unknown,
  options: ReadPresentationFeatureCapabilityCatalogOptions = {},
): PresentationFeatureCapabilityCatalog {
  return readCatalog(input, options, true);
}

function readCatalog(
  input: unknown,
  options: ReadPresentationFeatureCapabilityCatalogOptions,
  audited: boolean,
): PresentationFeatureCapabilityCatalog {
  const source = requireRecord(input, "catalog");
  const untrustedOptions = options as {
    readonly handlerFeatureIds?: unknown;
  };
  if (untrustedOptions.handlerFeatureIds !== undefined) {
    throw new Error(
      "Presentation production command-registry adapter is not installed",
    );
  }
  const sourceVersion = source.version;
  if (
    sourceVersion !== PRESENTATION_LEGACY_CAPABILITY_CATALOG_VERSION &&
    sourceVersion !== PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION
  ) {
    throw new Error(
      `Unknown presentation capability catalog version: ${String(sourceVersion)}`,
    );
  }
  if (!Array.isArray(source.features)) {
    throw new Error("Missing presentation capability catalog features");
  }

  const features = Object.freeze(
    source.features.map((feature, index) =>
      sourceVersion === PRESENTATION_LEGACY_CAPABILITY_CATALOG_VERSION
        ? normalizeLegacyCapability(feature, index)
        : normalizeCurrentCapability(feature, index, audited),
    ),
  );
  const featureIds = Object.freeze(
    features.map((feature) => feature.featureId),
  );
  const ids = new Set<string>();
  for (const featureId of featureIds) {
    if (ids.has(featureId)) {
      throw new Error(`Duplicate presentation feature id: ${featureId}`);
    }
    ids.add(featureId);
  }

  const catalog = Object.freeze({
    schemaVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    sourceVersion,
    features,
    featureIds,
    handlerBindingsVerified: false,
  });
  validatedPresentationFeatureCapabilityCatalogs.add(catalog);
  return catalog;
}

export interface VerifyPresentationFeatureCapabilityCatalogOptions {
  readonly expectedLineage: FeatureCommandRegistryLineage;
  readonly expectedRegistry: FeatureCommandBindingRegistry;
  readonly expectedBindingContract: FeatureCommandBindingContract;
}

/**
 * Canonical command semantics consumed by the shared v4 binding authority.
 * The normalized v2 row is fully data-only and frozen, so its JSON encoding
 * fingerprints every command/evidence/domain/native boundary without trusting
 * a caller-supplied list of feature IDs.
 */
export function presentationFeatureCommandSemanticContract(
  capability: PresentationFeatureCapabilityContract,
): string {
  return JSON.stringify(capability);
}

/**
 * Upgrade one parsed v2 catalog only after the shared command registry proves
 * the exact lineage, binding contract, registry generation, semantic row, and
 * execute-function identity. Structural clones and detached live registries
 * remain inert because both this module and the shared registry use private
 * WeakMap/WeakSet ownership.
 */
export function verifyPresentationFeatureCapabilityCatalogBindings(
  catalog: PresentationFeatureCapabilityCatalog,
  manifest: FeatureCommandBindingManifest,
  options: VerifyPresentationFeatureCapabilityCatalogOptions,
): PresentationFeatureCapabilityCatalog {
  if (
    !validatedPresentationFeatureCapabilityCatalogs.has(catalog) ||
    catalog.sourceVersion !== PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION ||
    catalog.handlerBindingsVerified
  ) {
    throw new Error("Untrusted presentation capability catalog");
  }
  const trusted = readFeatureCommandBindingManifest(manifest, options);
  const trustedFeatureIds = [...trusted.boundFeatureIds].sort();
  const catalogFeatureIds = [...catalog.featureIds].sort();
  if (
    trusted.editor !== "presentation" ||
    trustedFeatureIds.length !== catalogFeatureIds.length ||
    trustedFeatureIds.some(
      (featureId, index) => featureId !== catalogFeatureIds[index],
    )
  ) {
    throw new Error("Presentation capability binding identity mismatch");
  }
  const verified = Object.freeze({
    ...catalog,
    handlerBindingsVerified: true,
  });
  validatedPresentationFeatureCapabilityCatalogs.add(verified);
  return verified;
}

function productionCapabilitySource(
  evidence: PresentationProductionCommandEvidence,
  order: number,
): Record<string, unknown> {
  const legacy = PRESENTATION_FEATURE_CAPABILITY_REGISTRY.find(
    (entry) => entry.id === evidence.featureId,
  );
  if (!legacy) {
    throw new Error(
      `Missing legacy compatibility row for production feature: ${evidence.featureId}`,
    );
  }
  return {
    featureId: evidence.featureId,
    sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    surface: {
      disposition: "cataloged",
      tabs: [...legacy.tabs],
      entry: legacy.entry,
      group: evidence.group,
      order,
      contextualPredicate: legacy.contextPredicate ?? "always",
      controlKind: "command",
    },
    command: {
      disposition: "bound",
      key: legacy.commandKey,
      handlerRef: evidence.handlerRef,
      undoPolicy: evidence.undoPolicy,
      availabilityPredicate: legacy.contextPredicate ?? "always",
    },
    authoring: { disposition: legacy.authoring },
    semantics: {
      defaults: {},
      limits: {},
      shortcutIds: [],
    },
    domain: {
      disposition:
        evidence.persistence === "external"
          ? "external"
          : evidence.persistence === "session"
            ? "view-only"
            : "typed",
      persistence: evidence.persistence,
    },
    import: {
      disposition: legacy.pptxImport,
    },
    fromScratchExport: {
      disposition: legacy.fromScratchExport,
    },
    patchbackExport: {
      disposition: legacy.patchbackExport,
    },
    officeOutcome: {
      powerpoint: legacy.powerpoint,
      wps: legacy.wps,
    },
  };
}

const PRESENTATION_PRODUCTION_CAPABILITY_SOURCES = Object.freeze(
  PRESENTATION_PRODUCTION_COMMAND_EVIDENCE.map(productionCapabilitySource),
);

/** Strict v2 truth table consumed by the mounted production registry. */
export const PRESENTATION_FEATURE_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: PRESENTATION_PRODUCTION_CAPABILITY_SOURCES,
  });

export function presentationProductionCapabilityCatalogForFeatureIds(
  featureIds: readonly string[],
): PresentationFeatureCapabilityCatalog {
  const requested = new Set(featureIds);
  if (requested.size !== featureIds.length) {
    throw new Error("Duplicate presentation feature command");
  }
  const features = PRESENTATION_PRODUCTION_CAPABILITY_SOURCES.filter(
    (source) => requested.has(String(source.featureId)),
  );
  if (features.length !== requested.size) {
    const known = new Set(features.map((source) => String(source.featureId)));
    const missing = featureIds.find((featureId) => !known.has(featureId));
    throw new Error(`Missing production v2 capability evidence: ${missing}`);
  }
  return readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features,
  });
}

export function presentationRuntimeCapabilityFromContract(
  contract: PresentationFeatureCapabilityContract,
): PresentationFeatureCapability {
  const evidence = presentationProductionCommandEvidence(contract.featureId);
  const compatibility = PRESENTATION_FEATURE_CAPABILITY_REGISTRY.find(
    (entry) => entry.id === contract.featureId,
  );
  if (
    !evidence ||
    contract.sourceVersion !== PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION ||
    contract.surface.disposition !== "cataloged" ||
    contract.surface.entry === "unknown" ||
    contract.surface.contextualPredicate === "unknown" ||
    contract.command.key === null ||
    contract.command.availabilityPredicate === "unknown" ||
    contract.authoring.disposition === "unknown" ||
    contract.import.disposition === "unknown" ||
    contract.fromScratchExport.disposition === "unknown" ||
    contract.patchbackExport.disposition === "unknown" ||
    contract.officeOutcome.powerpoint === "unknown" ||
    contract.officeOutcome.wps === "unknown"
  ) {
    throw new Error(`Incomplete production v2 capability: ${contract.featureId}`);
  }
  return Object.freeze({
    id: contract.featureId,
    tabs: contract.surface.tabs,
    entry: contract.surface.entry,
    authoring: contract.authoring.disposition,
    pptxImport: contract.import.disposition,
    fromScratchExport: contract.fromScratchExport.disposition,
    patchbackExport: contract.patchbackExport.disposition,
    powerpoint: contract.officeOutcome.powerpoint,
    wps: contract.officeOutcome.wps,
    commandKey: contract.command.key,
    contextPredicate: contract.command.availabilityPredicate,
    ...(compatibility?.reasonKey
      ? { reasonKey: compatibility.reasonKey }
      : {}),
  });
}

export const PRESENTATION_IMAGE_FILE_FEATURE_ID = "drawing.image" as const;

/** Strict v2 truth row for PPT.WP.IMG01.FILE-INSERT. */
export const PRESENTATION_IMAGE_FILE_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_IMAGE_FILE_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["home", "insert", "graphic-format"],
          entry: "both",
          group: "Pictures",
          order: 0,
          contextualPredicate: "editable",
          controlKind: "menuitem-file-picker",
        },
        command: {
          disposition: "bound",
          key: "presentation.command.insertImage",
          handlerRef: "PresentationWindow#insertImageFile",
          undoPolicy: "single-transaction",
          availabilityPredicate: "editable",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: {
            source: "local-file",
            selectionCount: 1,
            placement: "centered-natural-96dpi-fit-slide",
          },
          limits: {
            mimeTypes: ["image/png", "image/jpeg", "image/svg+xml"],
            maxBytes: 50 * 1024 * 1024,
            maxDimension: 10_000,
            maxPixels: 40_000_000,
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: {
          powerpoint: "editable",
          wps: "editable",
        },
      },
    ],
  });

export const PRESENTATION_AUDIO_FILE_FEATURE_ID = "media.audioFile" as const;

/** Local audio file insertion with native scratch and imported-PPTX writers. */
export const PRESENTATION_AUDIO_FILE_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_AUDIO_FILE_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["insert"],
          entry: "ribbon",
          group: "Media",
          order: 1,
          contextualPredicate: "editable",
          controlKind: "menuitem-file-picker",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.insertAudioFile",
          handlerRef: "PresentationWindow#onInsertAudioFile",
          undoPolicy: "single-transaction",
          availabilityPredicate: "editable",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: {
            start: "click",
            volume: "high",
            placement: "centered-audio-control",
          },
          limits: {
            mimeTypes: [
              "audio/aac",
              "audio/mp4",
              "audio/mpeg",
              "audio/ogg",
              "audio/wav",
            ],
            maxBytes: 1_500_000,
            source: "one-local-file",
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: { powerpoint: "editable", wps: "unsupported" },
      },
    ],
  });

const PRESENTATION_AUDIO_FILE_FEATURE_CAPABILITY = leaf({
  id: PRESENTATION_AUDIO_FILE_FEATURE_ID,
  tabs: ["insert"],
  entry: "ribbon",
  authoring: "native",
  pptxImport: "native",
  fromScratchExport: "native",
  patchbackExport: "native",
  powerpoint: "editable",
  wps: "unsupported",
  commandKey: "presentation.command.insertAudioFile",
  contextPredicate: "editable",
  reasonKey: "presentation.capability.editableRequired",
});

export const PRESENTATION_RECORD_AUDIO_FEATURE_ID =
  "media.recordAudio" as const;

/** Microphone capture drafts until Insert, then uses the native audio writer. */
export const PRESENTATION_RECORD_AUDIO_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_RECORD_AUDIO_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["insert"],
          entry: "ribbon",
          group: "Media",
          order: 2,
          contextualPredicate: "editable",
          controlKind: "menuitem-dialog-recorder",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.recordAudio",
          handlerRef: "PresentationWindow#openRecordAudioDialog",
          undoPolicy: "single-transaction",
          availabilityPredicate: "editable",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: {
            start: "click",
            volume: "high",
            insertion: "explicit-after-preview",
          },
          limits: {
            captureMimeTypes: [
              "audio/mp4;codecs=mp4a.40.2",
              "audio/mp4",
              "audio/ogg;codecs=opus",
              "audio/ogg",
            ],
            maxBytes: 1_500_000,
            target: "one-live-current-slide",
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: { powerpoint: "editable", wps: "unsupported" },
      },
    ],
  });

const PRESENTATION_RECORD_AUDIO_FEATURE_CAPABILITY = leaf({
  id: PRESENTATION_RECORD_AUDIO_FEATURE_ID,
  tabs: ["insert"],
  entry: "ribbon",
  authoring: "native",
  pptxImport: "native",
  fromScratchExport: "native",
  patchbackExport: "native",
  powerpoint: "editable",
  wps: "unsupported",
  commandKey: "presentation.command.recordAudio",
  contextPredicate: "editable",
  reasonKey: "presentation.capability.editableRequired",
});

export const PRESENTATION_VIDEO_FILE_FEATURE_ID = "media.videoFile" as const;

/** One attested MP4 command: replace an exact native movie or append a new one. */
export const PRESENTATION_VIDEO_FILE_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_VIDEO_FILE_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["insert"],
          entry: "ribbon",
          group: "Media",
          order: 0,
          contextualPredicate: "editable",
          controlKind: "menuitem-file-picker",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.insertVideoFile",
          handlerRef: "PresentationWindow#videoFileCommand",
          undoPolicy: "single-transaction",
          availabilityPredicate: "editable",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: {
            format: "video/mp4",
            start: "click",
            volume: "high",
          },
          limits: {
            maxBytes: 256_000_000,
            maxAssetsPerDeck: 20,
            codec: "H.264/AVC",
            target:
              "replace-one-selected-native-video-otherwise-append-to-current-slide",
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: { powerpoint: "editable", wps: "unsupported" },
      },
    ],
  });

const PRESENTATION_VIDEO_FILE_FEATURE_CAPABILITY = leaf({
  id: PRESENTATION_VIDEO_FILE_FEATURE_ID,
  tabs: ["insert"],
  entry: "ribbon",
  authoring: "native",
  pptxImport: "native",
  fromScratchExport: "native",
  patchbackExport: "native",
  powerpoint: "editable",
  wps: "unsupported",
  commandKey: "presentation.command.insertVideoFile",
  contextPredicate: "editable",
  reasonKey: "presentation.capability.editableRequired",
});

export const PRESENTATION_PICTURE_REPLACE_FILE_FEATURE_ID =
  "picture.changePicture.file" as const;

/** Strict v2 truth row for PPT.WP.IMG02.REPLACE-PICTURE's local-file leaf. */
export const PRESENTATION_PICTURE_REPLACE_FILE_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_PICTURE_REPLACE_FILE_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["image-format"],
          entry: "both",
          group: "Adjust",
          order: 0,
          contextualPredicate: "selection",
          controlKind: "menuitem-file-picker",
        },
        command: {
          disposition: "bound",
          key: "presentation.command.changePictureFromFile",
          handlerRef: "PresentationWindow#replacePictureFromFile",
          undoPolicy: "single-transaction",
          availabilityPredicate: "selection",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: {
            source: "local-file",
            selectionCount: 1,
            preservePictureFormatting: true,
          },
          limits: {
            mimeTypes: ["image/png", "image/jpeg"],
            maxBytes: 50 * 1024 * 1024,
            maxDimension: 10_000,
            maxPixels: 40_000_000,
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: {
          powerpoint: "editable",
          wps: "editable",
        },
      },
    ],
  });
export const PRESENTATION_PICTURE_CROP_PRIMARY_CONTROL_ID =
  "PPTCTL.CONTEXT.PICTURE.CROP.PRIMARY" as const;
export const PRESENTATION_PICTURE_CROP_MENU_CONTROL_ID =
  "PPTCTL.CONTEXT.PICTURE.CROP.MENU" as const;
export const PRESENTATION_PICTURE_CROP_MENU_CROP_CONTROL_ID =
  "PPTCTL.CONTEXT.PICTURE.CROP.MENU.CROP" as const;
export const PRESENTATION_PICTURE_CROP_MENU_FILL_CONTROL_ID =
  "PPTCTL.CONTEXT.PICTURE.CROP.MENU.FILL" as const;
export const PRESENTATION_PICTURE_CROP_MENU_FIT_CONTROL_ID =
  "PPTCTL.CONTEXT.PICTURE.CROP.MENU.FIT" as const;
export const PRESENTATION_PICTURE_CROP_FEATURE_ID = "picture.crop" as const;
export const PRESENTATION_PICTURE_CROP_FILL_FEATURE_ID =
  "picture.crop.fill" as const;
export const PRESENTATION_PICTURE_CROP_FIT_FEATURE_ID =
  "picture.crop.fit" as const;

function pictureCropCapabilityRow(config: {
  readonly featureId: string;
  readonly commandKey: string;
  readonly handlerRef: string;
  readonly controlKind: string;
  readonly order: number;
  readonly mode: "crop" | "fill" | "fit";
  readonly controls: readonly string[];
}): Record<string, unknown> {
  return {
    featureId: config.featureId,
    sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    surface: {
      disposition: "cataloged",
      tabs: ["image-format"],
      entry: "both",
      group: "Size",
      order: config.order,
      contextualPredicate: "selection",
      controlKind: config.controlKind,
    },
    command: {
      disposition: "bound",
      key: config.commandKey,
      handlerRef: config.handlerRef,
      undoPolicy: "single-transaction",
      availabilityPredicate: "selection",
    },
    authoring: { disposition: "native" },
    semantics: {
      defaults: {
        mode: config.mode,
        anchor: "center",
        selectionCount: 1,
        preserveSourceBitmap: true,
      },
      limits: {
        target: "single-editable-native-picture",
        coordinateSpace: "signed-source-keep-rect",
      },
      shortcutIds: [],
    },
    domain: {
      disposition: "typed",
      persistence: "document",
    },
    import: {
      disposition: "native",
    },
    fromScratchExport: {
      disposition: "native",
    },
    patchbackExport: {
      disposition: "native",
    },
    officeOutcome: {
      powerpoint: "editable",
      wps: "editable",
    },
  };
}

/** Strict v2 rows for the implemented Crop split-button and Fill/Fit leaves. */
export const PRESENTATION_PICTURE_CROP_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      pictureCropCapabilityRow({
        featureId: PRESENTATION_PICTURE_CROP_FEATURE_ID,
        commandKey: "presentation.command.cropPicture",
        handlerRef: "PresentationWindow#cropPicture",
        controlKind: "split-button-primary-and-menuitem",
        order: 0,
        mode: "crop",
        controls: [
          PRESENTATION_PICTURE_CROP_PRIMARY_CONTROL_ID,
          PRESENTATION_PICTURE_CROP_MENU_CONTROL_ID,
          PRESENTATION_PICTURE_CROP_MENU_CROP_CONTROL_ID,
        ],
      }),
      pictureCropCapabilityRow({
        featureId: PRESENTATION_PICTURE_CROP_FILL_FEATURE_ID,
        commandKey: "presentation.command.cropPictureFill",
        handlerRef: "PresentationWindow#cropPictureFill",
        controlKind: "menuitem",
        order: 4,
        mode: "fill",
        controls: [PRESENTATION_PICTURE_CROP_MENU_FILL_CONTROL_ID],
      }),
      pictureCropCapabilityRow({
        featureId: PRESENTATION_PICTURE_CROP_FIT_FEATURE_ID,
        commandKey: "presentation.command.cropPictureFit",
        handlerRef: "PresentationWindow#cropPictureFit",
        controlKind: "menuitem",
        order: 5,
        mode: "fit",
        controls: [PRESENTATION_PICTURE_CROP_MENU_FIT_CONTROL_ID],
      }),
    ],
  });

export const PRESENTATION_PICTURE_ALT_TEXT_FEATURE_ID =
  "picture.altText" as const;

/** Strict native row for Picture Format → Alt Text. */
export const PRESENTATION_PICTURE_ALT_TEXT_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_PICTURE_ALT_TEXT_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["image-format"],
          entry: "ribbon",
          group: "Accessibility",
          order: 0,
          contextualPredicate: "selection",
          controlKind: "button-dialog",
        },
        command: {
          disposition: "bound",
          key: "presentation.command.editPictureAltText",
          handlerRef: "PresentationWindow#editPictureAltText",
          undoPolicy: "single-transaction",
          availabilityPredicate: "selection",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: {
            selectionCount: 1,
            decorativeWhenExtensionMissing: false,
            preserveLegacyTitle: true,
          },
          limits: {
            target: "single-editable-native-picture",
            descriptionCharacters: 1_000,
            onlineGeneration: "excluded",
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: {
          powerpoint: "editable",
          wps: "editable",
        },
      },
    ],
  });

export const PRESENTATION_PICTURE_BORDER_FEATURE_ID = "picture.border" as const;
export const PRESENTATION_PICTURE_BORDER_PRIMARY_CONTROL_ID =
  "PPTCTL.CONTEXT.PICTURE.STYLE.BORDER.PRIMARY" as const;
export const PRESENTATION_PICTURE_BORDER_MENU_CONTROL_ID =
  "PPTCTL.CONTEXT.PICTURE.STYLE.BORDER.MENU" as const;

/** Strict native row for the supported Picture Border split control. */
export const PRESENTATION_PICTURE_BORDER_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_PICTURE_BORDER_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["image-format"],
          entry: "both",
          group: "Picture Styles",
          order: 1,
          contextualPredicate: "selection",
          controlKind: "split-button-menu",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.editPictureBorder",
          handlerRef: "PresentationWindow#editPictureBorder",
          undoPolicy: "single-transaction",
          availabilityPredicate: "selection",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: {
            rememberedExplicitRgb: "#1f2937",
            omittedNativeWidthPreserved: true,
          },
          limits: {
            target: "single-editable-native-picture",
            colors: "explicit-rgb-only",
            styles: ["none", "solid", "dashed", "dotted"],
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: { powerpoint: "editable", wps: "editable" },
      },
    ],
  });

export const PRESENTATION_PICTURE_TRANSPARENCY_FEATURE_ID =
  "picture.transparency" as const;

/** Strict native row for Picture Transparency presets and detailed options. */
export const PRESENTATION_PICTURE_TRANSPARENCY_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_PICTURE_TRANSPARENCY_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["image-format"],
          entry: "both",
          group: "Adjust",
          order: 2,
          contextualPredicate: "selection",
          controlKind: "gallery-menu",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.editPictureTransparency",
          handlerRef: "PresentationWindow#editPictureTransparency",
          undoPolicy: "single-transaction",
          availabilityPredicate: "selection",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: {
            transparencyPercent: 0,
            visibleOpacityPercent: 100,
          },
          limits: {
            target: "single-editable-native-picture",
            presets: [0, 15, 30, 50, 65, 80, 95],
            mapping: "visible-opacity=100-transparency",
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: { powerpoint: "editable", wps: "editable" },
      },
    ],
  });

export const PRESENTATION_PICTURE_CORRECTIONS_FEATURE_ID =
  "picture.corrections" as const;

/** Strict native row for the bounded Picture Corrections gallery slice. */
export const PRESENTATION_PICTURE_CORRECTIONS_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_PICTURE_CORRECTIONS_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["image-format"],
          entry: "both",
          group: "Adjust",
          order: 1,
          contextualPredicate: "selection",
          controlKind: "gallery-menu",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.editPictureCorrections",
          handlerRef: "PresentationWindow#editPictureCorrections",
          undoPolicy: "single-transaction",
          availabilityPredicate: "selection",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: {
            brightnessDelta: 0,
            contrastDelta: 0,
            persistedNeutral: 100,
          },
          limits: {
            target: "single-editable-native-picture",
            signedDelta: [-100, 100],
            presets: [-40, -20, 0, 20, 40],
            storageMapping: "stored=delta+100",
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: { powerpoint: "editable", wps: "editable" },
      },
    ],
  });

export const PRESENTATION_PICTURE_COLOR_FEATURE_ID =
  "picture.color.recolor" as const;

/** Strict native row for the bounded Picture Color → Recolor slice. */
export const PRESENTATION_PICTURE_COLOR_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_PICTURE_COLOR_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["image-format"],
          entry: "menu",
          group: "Adjust",
          order: 2,
          contextualPredicate: "selection",
          controlKind: "menu",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.editPictureColor",
          handlerRef: "PresentationWindow#editPictureColor",
          undoPolicy: "single-transaction",
          availabilityPredicate: "selection",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: { recolor: "none" },
          limits: {
            target: "single-editable-native-picture",
            recolor: ["none", "grayscale"],
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: { powerpoint: "editable", wps: "editable" },
      },
    ],
  });

export const PRESENTATION_PICTURE_BRING_FORWARD_FEATURE_ID =
  "picture.arrange.bringForward" as const;
export const PRESENTATION_PICTURE_SEND_BACKWARD_FEATURE_ID =
  "picture.arrange.sendBackward" as const;
export const PRESENTATION_PICTURE_BRING_FORWARD_CONTROL_ID =
  "PPTCTL.CONTEXT.PICTURE.ARRANGE.BRING_FORWARD.PRIMARY" as const;
export const PRESENTATION_PICTURE_SEND_BACKWARD_CONTROL_ID =
  "PPTCTL.CONTEXT.PICTURE.ARRANGE.SEND_BACKWARD.PRIMARY" as const;

const presentationPictureArrangeStepCapability = (
  featureId:
    | typeof PRESENTATION_PICTURE_BRING_FORWARD_FEATURE_ID
    | typeof PRESENTATION_PICTURE_SEND_BACKWARD_FEATURE_ID,
  controlId:
    | typeof PRESENTATION_PICTURE_BRING_FORWARD_CONTROL_ID
    | typeof PRESENTATION_PICTURE_SEND_BACKWARD_CONTROL_ID,
  commandKey: string,
  handlerRef: string,
) => ({
  featureId,
  sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
  surface: {
    disposition: "cataloged" as const,
    tabs: ["image-format" as const],
    entry: "ribbon" as const,
    group: "Arrange",
    order: featureId === PRESENTATION_PICTURE_BRING_FORWARD_FEATURE_ID ? 0 : 1,
    contextualPredicate: "selection" as const,
    controlKind: "button-primary",
  },
  command: {
    disposition: "bound" as const,
    key: commandKey,
    handlerRef,
    undoPolicy: "single-transaction" as const,
    availabilityPredicate: "selection" as const,
  },
  authoring: { disposition: "native" as const },
  semantics: {
    defaults: { step: "one-adjacent-layer" },
    limits: {
      target: "single-live-unlocked-source-backed-picture",
      extremes: "disabled",
    },
    shortcutIds: [],
  },
  domain: {
    disposition: "typed" as const,
    persistence: "document" as const,
  },
  import: {
    disposition: "native" as const,
  },
  fromScratchExport: {
    disposition: "reject" as const,
  },
  patchbackExport: {
    disposition: "native" as const,
  },
  officeOutcome: {
    powerpoint: "editable" as const,
    wps: "unsupported" as const,
  },
});

/** Strict native rows for the two bounded Picture Arrange primary buttons. */
export const PRESENTATION_PICTURE_ARRANGE_STEP_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      presentationPictureArrangeStepCapability(
        PRESENTATION_PICTURE_BRING_FORWARD_FEATURE_ID,
        PRESENTATION_PICTURE_BRING_FORWARD_CONTROL_ID,
        "presentation.command.bringPictureForward",
        "PresentationWindow#bringPictureForward",
      ),
      presentationPictureArrangeStepCapability(
        PRESENTATION_PICTURE_SEND_BACKWARD_FEATURE_ID,
        PRESENTATION_PICTURE_SEND_BACKWARD_CONTROL_ID,
        "presentation.command.sendPictureBackward",
        "PresentationWindow#sendPictureBackward",
      ),
    ],
  });

export const PRESENTATION_PICTURE_SIZE_WIDTH_FEATURE_ID =
  "picture.size.width" as const;
export const PRESENTATION_PICTURE_SIZE_HEIGHT_FEATURE_ID =
  "picture.size.height" as const;
export const PRESENTATION_PICTURE_SIZE_HEIGHT_FIELD_CONTROL_ID =
  "PPTCTL.CONTEXT.PICTURE.SIZE.HEIGHT.FIELD" as const;
export const PRESENTATION_PICTURE_SIZE_HEIGHT_STEPPER_CONTROL_ID =
  "PPTCTL.CONTEXT.PICTURE.SIZE.HEIGHT.STEPPER" as const;
export const PRESENTATION_PICTURE_SIZE_WIDTH_FIELD_CONTROL_ID =
  "PPTCTL.CONTEXT.PICTURE.SIZE.WIDTH.FIELD" as const;
export const PRESENTATION_PICTURE_SIZE_WIDTH_STEPPER_CONTROL_ID =
  "PPTCTL.CONTEXT.PICTURE.SIZE.WIDTH.STEPPER" as const;

const presentationPictureSizeCapability = (
  featureId:
    | typeof PRESENTATION_PICTURE_SIZE_WIDTH_FEATURE_ID
    | typeof PRESENTATION_PICTURE_SIZE_HEIGHT_FEATURE_ID,
  fieldControlId:
    | typeof PRESENTATION_PICTURE_SIZE_WIDTH_FIELD_CONTROL_ID
    | typeof PRESENTATION_PICTURE_SIZE_HEIGHT_FIELD_CONTROL_ID,
  stepperControlId:
    | typeof PRESENTATION_PICTURE_SIZE_WIDTH_STEPPER_CONTROL_ID
    | typeof PRESENTATION_PICTURE_SIZE_HEIGHT_STEPPER_CONTROL_ID,
  commandKey: string,
  handlerRef: string,
  order: number,
) => ({
  featureId,
  sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
  surface: {
    disposition: "cataloged" as const,
    tabs: ["image-format" as const],
    entry: "ribbon" as const,
    group: "Size",
    order,
    contextualPredicate: "selection" as const,
    controlKind: "value-field-with-stepper",
  },
  command: {
    disposition: "parameterized" as const,
    key: commandKey,
    handlerRef,
    undoPolicy: "single-transaction" as const,
    availabilityPredicate: "selection" as const,
  },
  authoring: { disposition: "native" as const },
  semantics: {
    defaults: {
      displayUnit: "cm",
      displayPrecision: 2,
      stepEmu: 36_000,
    },
    limits: {
      target: "single-live-unlocked-authenticated-picture",
      extentEmu: [1, PRESENTATION_GEOMETRY_MAX_EMU],
      aspectLock: "authenticated-native-picLocks",
    },
    shortcutIds: [],
  },
  domain: {
    disposition: "typed" as const,
    persistence: "document" as const,
  },
  import: {
    disposition: "native" as const,
  },
  fromScratchExport: {
    disposition: "native" as const,
  },
  patchbackExport: {
    disposition: "native" as const,
  },
  officeOutcome: {
    powerpoint: "editable" as const,
    wps: "unsupported" as const,
  },
});

/** Strict native rows for Picture Size width/height fields and steppers. */
export const PRESENTATION_PICTURE_SIZE_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      presentationPictureSizeCapability(
        PRESENTATION_PICTURE_SIZE_HEIGHT_FEATURE_ID,
        PRESENTATION_PICTURE_SIZE_HEIGHT_FIELD_CONTROL_ID,
        PRESENTATION_PICTURE_SIZE_HEIGHT_STEPPER_CONTROL_ID,
        "presentation.command.resizePictureHeight",
        "PresentationWindow#resizePictureHeight",
        0,
      ),
      presentationPictureSizeCapability(
        PRESENTATION_PICTURE_SIZE_WIDTH_FEATURE_ID,
        PRESENTATION_PICTURE_SIZE_WIDTH_FIELD_CONTROL_ID,
        PRESENTATION_PICTURE_SIZE_WIDTH_STEPPER_CONTROL_ID,
        "presentation.command.resizePictureWidth",
        "PresentationWindow#resizePictureWidth",
        1,
      ),
    ],
  });

export const PRESENTATION_SLIDE_VISIBILITY_FEATURE_ID =
  "slides.toggleHidden" as const;
export const PRESENTATION_SLIDE_VISIBILITY_CONTROL_ID =
  "PPTCTL.RIBBON.SLIDESHOW.SETUP.HIDE_SLIDE" as const;
export const PRESENTATION_SLIDESHOW_START_FEATURE_IDS = Object.freeze([
  "slideshow.fromStart",
  "slideshow.fromCurrent",
] as const);
/** Current PowerPoint for the web exposes Presenter View. MRAI's session-only
 * authority opens a same-origin speaker console, keeps it synchronized with the
 * audience window, falls back to the in-window console when popups are blocked,
 * and never mutates the deck. Self-running/kiosk authoring remains unavailable. */
export const PRESENTATION_WEB_PRESENTER_VIEW_SUPPORTED = true as const;
export const PRESENTATION_WEB_SELF_RUNNING_SHOW_SUPPORTED = false as const;

export const PRESENTATION_SLIDE_TRANSITION_FEATURE_ID =
  "slides.setTransition" as const;

/** Strict native row for the bounded native transition gallery writers. */
export const PRESENTATION_SLIDE_TRANSITION_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_SLIDE_TRANSITION_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["transitions"],
          entry: "ribbon",
          group: "Transition to This Slide",
          order: 0,
          contextualPredicate: "editable",
          controlKind: "gallery",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.setSlideTransition",
          handlerRef: "PresentationWindow#setSlideTransition",
          undoPolicy: "single-transaction",
          availabilityPredicate: "editable",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: {
            kind: "none",
            selectedEffectDurationMs: 1000,
            nativeSpeed: "slow",
            pushDirection: "u",
            wipeDirection: "omitted",
            shapeOption: "circle",
            zoomOption: "in",
            orientationOption: "horz",
            splitOption: { orient: "horz", dir: "out" },
            combOption: "horz",
            p15PresetCarrier: "mc:AlternateContent",
            p15PresetFallback: "p:fade",
            // Fall Over, Drape, Peel Off and Page Curl are written on their
            // left side; Wind, Airplane and Origami on their right. `invX="1"`
            // is the other side of the same effect, never a second one.
            p15PresetSideDefault: "invX omitted",
            p15PresetAuthoredSide: "right",
            p14LeftRightCarrier: "mc:AlternateContent",
            p14LeftRightDefault: "r",
            p14LeftRightFallback: "p:fade",
            p14OrientationCarrier: "mc:AlternateContent",
            p14OrientationDefault: "horz",
            p14OrientationFallback: "p:fade",
            p14SideDirectionCarrier: "mc:AlternateContent",
            p14SideDirectionDefault: "l",
            p14SideDirectionFallback: "p:fade",
            p14EmptyCarrier: "mc:AlternateContent",
            p14EmptyFallback: "p:fade",
            p14TypedOptionsCarrier: "mc:AlternateContent",
            p14TypedOptionsFallback: "p:fade",
            revealOption: { dir: "l", thruBlk: false },
            rippleOption: "center",
            glitterOption: { dir: "l", pattern: "diamond" },
            shredOption: { pattern: "strip", dir: "in" },
            flyThroughOption: { dir: "in", hasBounce: false },
            clockOption: { spokes: 1 },
            eightDirectionCarrier: "direct-p:transition",
            eightDirectionDefault: "l",
          },
          limits: {
            kinds: [
              "none",
              "fade",
              "push",
              "wipe",
              "split",
              "dissolve",
              "cut",
              "shape",
              "uncover",
              "cover",
              "zoom",
              "random-bars",
              "checkerboard",
              "blinds",
              "clock",
              "comb",
              "fall-over",
              "drape",
              "curtains",
              "wind",
              "prestige",
              "fracture",
              "crush",
              "peel-off",
              "page-curl",
              "airplane",
              "origami",
              "switch",
              "flip",
              "gallery",
              "ferris-wheel",
              "cube", "box", "rotate", "orbit",
              "conveyor",
              "doors",
              "window",
              "vortex",
              "pan",
              "flash",
              "honeycomb",
              "reveal",
              "ripple",
              "glitter",
              "shred",
              "fly-through",
              "random",
            ],
            target: "one-live-active-editable-slide",
            legacyImportedWithoutReceipt: "preserve-only",
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: {
          powerpoint: "editable",
          wps: "unsupported",
        },
      },
    ],
  });

export const PRESENTATION_SLIDE_TRANSITION_DURATION_FEATURE_ID =
  "slides.setTransitionDuration" as const;

/** Strict native Duration row shared by the field and its number stepper. */
export const PRESENTATION_SLIDE_TRANSITION_DURATION_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_SLIDE_TRANSITION_DURATION_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["transitions"],
          entry: "ribbon",
          group: "Timing",
          order: 0,
          contextualPredicate: "editable",
          controlKind: "number-field-with-stepper",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.setTransitionDuration",
          handlerRef: "PresentationWindow#setTransitionDuration",
          undoPolicy: "single-transaction",
          availabilityPredicate: "editable",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: {
            selectedBasicEffectDurationMs: 1000,
            displayUnit: "seconds",
            displayDecimals: 2,
          },
          limits: {
            minimumMs: 0,
            maximumMs: 59990,
            stepMs: 10,
            target: "one-live-active-editable-basic-transition",
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: {
          powerpoint: "editable",
          wps: "unsupported",
        },
      },
    ],
  });

export const PRESENTATION_SLIDE_TRANSITION_ADVANCE_FEATURE_ID =
  "slides.setTransitionAdvance" as const;

/** Strict native authority shared by the On Click and After controls. */
export const PRESENTATION_SLIDE_TRANSITION_ADVANCE_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_SLIDE_TRANSITION_ADVANCE_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["transitions"],
          entry: "ribbon",
          group: "Timing",
          order: 1,
          contextualPredicate: "editable",
          controlKind: "checkboxes-and-number-field",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.setTransitionAdvance",
          handlerRef: "PresentationWindow#setSlideAdvance",
          undoPolicy: "single-transaction",
          availabilityPredicate: "editable",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: {
            advanceOnClick: true,
            advanceAfterEnabled: false,
            advanceAfterMs: 0,
            displayUnit: "seconds",
            displayDecimals: 2,
          },
          limits: {
            minimumAfterMs: 0,
            maximumAfterMs: 600000,
            stepMs: 10,
            importedMaximumAfterMs: 4294967295,
            target: "one-live-active-editable-slide",
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: {
          powerpoint: "editable",
          wps: "unsupported",
        },
      },
    ],
  });

export const PRESENTATION_SLIDE_TRANSITION_SOUND_FEATURE_ID =
  "slides.setTransitionSound" as const;

/** One bounded embedded-WAV transition sound, including None/clear. */
export const PRESENTATION_SLIDE_TRANSITION_SOUND_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_SLIDE_TRANSITION_SOUND_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["transitions"],
          entry: "ribbon",
          group: "Timing",
          order: 1,
          contextualPredicate: "editable",
          controlKind: "menuitem-file-picker",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.setTransitionSound",
          handlerRef: "PresentationWindow#setTransitionSound",
          undoPolicy: "single-transaction",
          availabilityPredicate: "editable",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: { sound: "none", loop: false },
          limits: {
            mimeTypes: ["audio/wav"],
            maxBytes: 1450000,
            target: "one-live-active-editable-slide",
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: { powerpoint: "editable", wps: "unsupported" },
      },
    ],
  });

export const PRESENTATION_SLIDE_TRANSITION_APPLY_ALL_FEATURE_ID =
  "slides.applyTransitionToAll" as const;

/** Atomic deck-wide authority for the complete TRN01–03 transition state. */
export const PRESENTATION_SLIDE_TRANSITION_APPLY_ALL_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_SLIDE_TRANSITION_APPLY_ALL_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["transitions"],
          entry: "ribbon",
          group: "Timing",
          order: 2,
          contextualPredicate: "editable",
          controlKind: "button",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.applyTransitionToAll",
          handlerRef: "PresentationWindow#applyTransitionToAll",
          undoPolicy: "single-transaction",
          availabilityPredicate: "editable",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: {
            target: "all-slides-including-hidden",
            copy: "basic-effect-duration-on-click-after",
            transaction: "one-deck-commit",
          },
          limits: {
            source: "one-live-active-editable-slide",
            preflight: "all-deck-effect-and-advance-receipts",
            failure: "atomic-zero-mutation",
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: {
          powerpoint: "editable",
          wps: "unsupported",
        },
      },
    ],
  });

export const PRESENTATION_ANIMATION_EFFECT_FEATURE_ID =
  "animations.addEffect" as const;
export const PRESENTATION_ANIMATION_EFFECT_CONTROL_IDS = [
  "PPTCTL.RIBBON.ANIMATIONS.ENTRANCE.GALLERY",
  ...PRESENTATION_WEB_ANIMATION_EFFECTS.filter(
    (entry) => entry.category === "entrance",
  ).map((entry) => entry.controlId),
  "PPTCTL.RIBBON.ANIMATIONS.EMPHASIS.GALLERY",
  ...PRESENTATION_WEB_ANIMATION_EFFECTS.filter(
    (entry) => entry.category === "emphasis",
  ).map((entry) => entry.controlId),
  "PPTCTL.RIBBON.ANIMATIONS.EXIT.GALLERY",
  ...PRESENTATION_WEB_ANIMATION_EFFECTS.filter(
    (entry) => entry.category === "exit",
  ).map((entry) => entry.controlId),
  "PPTCTL.RIBBON.ANIMATIONS.ADD",
] as const;

/** Strict native row; kept additive so the frozen legacy 46-row view stays stable. */
export const PRESENTATION_ANIMATION_EFFECT_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_ANIMATION_EFFECT_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["animations"],
          entry: "ribbon",
          group: "Animation",
          order: 0,
          contextualPredicate: "selection",
          controlKind: "gallery",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.addAnimationEffect",
          handlerRef: "PresentationWindow#addAnimationEffect",
          undoPolicy: "single-transaction",
          availabilityPredicate: "selection",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: {
            trigger: "on-click",
            delayMs: 0,
            durationMs: 380,
          },
          limits: {
            kinds: PRESENTATION_WEB_ANIMATION_EFFECTS.map((entry) => entry.id),
            authorableKinds: PRESENTATION_NATIVE_AUTHORABLE_ANIMATION_EFFECTS,
            categories: ["entrance", "emphasis", "exit"],
            nativePresetVariants: PRESENTATION_WEB_ANIMATION_EFFECTS.filter(
              (entry) => entry.authoring === "native",
            ).map((entry) => ({
              id: entry.id,
              presetClass: entry.nativePresetClass,
              presetId: entry.nativePresetId,
              presetSubtype: entry.nativePresetSubtype,
              behavior: entry.nativeBehavior,
              ...(entry.nativeFilter === null
                ? {}
                : { filter: entry.nativeFilter }),
              ...(entry.nativeTransition === null
                ? {}
                : { transition: entry.nativeTransition }),
              ...(entry.nativeRotationBy === null
                ? {}
                : { rotationBy: entry.nativeRotationBy }),
              ...(entry.nativeScaleX === null
                ? {}
                : { scaleX: entry.nativeScaleX }),
              ...(entry.nativeScaleY === null
                ? {}
                : { scaleY: entry.nativeScaleY }),
              ...(entry.nativeOpacity === null
                ? {}
                : { opacity: entry.nativeOpacity }),
              ...(entry.nativeColorAttribute === null
                ? {}
                : { colorAttribute: entry.nativeColorAttribute }),
              ...(entry.nativeDefaultColorScheme === null
                ? {}
                : { defaultColorScheme: entry.nativeDefaultColorScheme }),
              ...(entry.nativeColorCompanionSets.length === 0
                ? {}
                : { colorCompanionSets: entry.nativeColorCompanionSets }),
              ...(entry.nativeColorBehaviorOverride === null
                ? {}
                : {
                    colorBehaviorOverride:
                      entry.nativeColorBehaviorOverride,
                  }),
              ...(entry.nativeEffectGroupId === null
                ? {}
                : { effectGroupId: entry.nativeEffectGroupId }),
              ...(entry.nativeEffectAccel === null
                ? {}
                : { effectAccel: entry.nativeEffectAccel }),
              ...(entry.nativeParagraphBuild
                ? { paragraphBuild: true }
                : {}),
              ...(entry.nativeShapeBuild ? { shapeBuild: true } : {}),
              ...(entry.nativeDimensionWidthFrom === null
                ? {}
                : {
                    widthFrom: entry.nativeDimensionWidthFrom,
                    widthTo: entry.nativeDimensionWidthTo,
                    heightFrom: entry.nativeDimensionHeightFrom,
                    heightTo: entry.nativeDimensionHeightTo,
                  }),
            })),
            target: "one-live-selected-block",
            fillColorTargets: ["shape"],
            lineColorTargets: ["shape"],
            fontColorTargets: ["text", "shape-with-text"],
            maxStepsPerSlide: 200,
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: {
          powerpoint: "editable",
          wps: "unsupported",
        },
      },
    ],
  });

const PRESENTATION_ANIMATION_EFFECT_FEATURE_CAPABILITY = leaf({
  id: PRESENTATION_ANIMATION_EFFECT_FEATURE_ID,
  tabs: ["animations"],
  entry: "ribbon",
  authoring: "native",
  pptxImport: "native",
  fromScratchExport: "native",
  patchbackExport: "native",
  powerpoint: "editable",
  wps: "unsupported",
  commandKey: "presentation.command.addAnimationEffect",
  contextPredicate: "selection",
  reasonKey: "presentation.capability.selectionRequired",
});

export const PRESENTATION_ANIMATION_EFFECT_OPTION_FEATURE_ID =
  "animations.editEffectOption" as const;

/** One exact selected-occurrence authority, independent from canvas selection. */
export const PRESENTATION_ANIMATION_EFFECT_OPTION_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_ANIMATION_EFFECT_OPTION_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["animations"],
          entry: "ribbon",
          group: "Animation",
          order: 1,
          contextualPredicate: "editable",
          controlKind: "menu",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.editAnimationEffectOption",
          handlerRef: "PresentationWindow#editAnimationEffectOption",
          undoPolicy: "single-transaction",
          availabilityPredicate: "editable",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: { persistence: "omit-native-default" },
          limits: {
            target: "one-live-selected-animation-occurrence",
            effectOptions: PRESENTATION_WEB_ANIMATION_EFFECTS.filter(
              (entry) => entry.options.length > 0,
            ).map((entry) => ({
              id: entry.id,
              defaultOption: entry.defaultOption,
              options: entry.options,
            })),
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: { powerpoint: "editable", wps: "unsupported" },
      },
    ],
  });

const PRESENTATION_ANIMATION_EFFECT_OPTION_FEATURE_CAPABILITY = leaf({
  id: PRESENTATION_ANIMATION_EFFECT_OPTION_FEATURE_ID,
  tabs: ["animations"],
  entry: "ribbon",
  authoring: "native",
  pptxImport: "native",
  fromScratchExport: "native",
  patchbackExport: "native",
  powerpoint: "editable",
  wps: "unsupported",
  commandKey: "presentation.command.editAnimationEffectOption",
  contextPredicate: "editable",
  reasonKey: "presentation.capability.selectionRequired",
});

export const PRESENTATION_ANIMATION_SEQUENCE_FEATURE_ID =
  "animations.editSequence" as const;
export const PRESENTATION_ANIMATION_SEQUENCE_CONTROL_IDS = [
  "PPTCTL.RIBBON.ANIMATIONS.PAINTER",
  "PPTCTL.PANE.ANIMATION.MOVE_UP",
  "PPTCTL.PANE.ANIMATION.MOVE_DOWN",
  "PPTCTL.PANE.ANIMATION.DELETE",
  "PPTCTL.PANE.ANIMATION.LIST",
] as const;

/** ANI02 Pane visibility plus strict ordered move/remove authority. */
export const PRESENTATION_ANIMATION_SEQUENCE_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_ANIMATION_SEQUENCE_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["animations"],
          entry: "ribbon",
          group: "Advanced Animation",
          order: 1,
          contextualPredicate: "always",
          controlKind: "button",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.editAnimationSequence",
          handlerRef: "PresentationWindow#editAnimationSequence",
          undoPolicy: "single-transaction",
          availabilityPredicate: "editable",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: { paneOpen: false },
          limits: {
            operations: ["move-earlier", "move-later", "remove"],
            target: "one-live-exact-occurrence",
            maxStepsPerSlide: 200,
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: {
          powerpoint: "editable",
          wps: "unsupported",
        },
      },
    ],
  });

const PRESENTATION_ANIMATION_SEQUENCE_FEATURE_CAPABILITY = leaf({
  id: PRESENTATION_ANIMATION_SEQUENCE_FEATURE_ID,
  tabs: ["animations"],
  entry: "ribbon",
  authoring: "native",
  pptxImport: "native",
  fromScratchExport: "native",
  patchbackExport: "native",
  powerpoint: "editable",
  wps: "unsupported",
  commandKey: "presentation.command.editAnimationSequence",
  contextPredicate: "editable",
  reasonKey: "presentation.capability.editableRequired",
});

export const PRESENTATION_ANIMATION_TIMING_FEATURE_ID =
  "animations.editTiming" as const;
export const PRESENTATION_ANIMATION_TIMING_CONTROL_IDS = [
  "PPTCTL.RIBBON.ANIMATIONS.START",
  "PPTCTL.RIBBON.ANIMATIONS.DURATION.FIELD",
  "PPTCTL.RIBBON.ANIMATIONS.DURATION.STEPPER",
  "PPTCTL.PANE.ANIMATION.SETTINGS",
] as const;

/** ANI03 Start/Delay/Duration row with strict native timing custody. */
export const PRESENTATION_ANIMATION_TIMING_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_ANIMATION_TIMING_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["animations"],
          entry: "ribbon",
          group: "Timing",
          order: 2,
          contextualPredicate: "editable",
          controlKind: "field",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.editAnimationTiming",
          handlerRef: "PresentationWindow#editAnimationTiming",
          undoPolicy: "single-transaction",
          availabilityPredicate: "editable",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: { trigger: "on-click", delayMs: 0, durationMs: 380 },
          limits: {
            triggers: ["on-click", "with-previous", "after-previous"],
            delayMs: [0, 60000],
            durationMs: [0, 60000],
            stepMs: 50,
            target: "one-live-exact-occurrence",
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: { powerpoint: "editable", wps: "unsupported" },
      },
    ],
  });

const PRESENTATION_ANIMATION_TIMING_FEATURE_CAPABILITY = leaf({
  id: PRESENTATION_ANIMATION_TIMING_FEATURE_ID,
  tabs: ["animations"],
  entry: "ribbon",
  authoring: "native",
  pptxImport: "native",
  fromScratchExport: "native",
  patchbackExport: "native",
  powerpoint: "editable",
  wps: "unsupported",
  commandKey: "presentation.command.editAnimationTiming",
  contextPredicate: "editable",
  reasonKey: "presentation.capability.editableRequired",
});

export const PRESENTATION_ANIMATION_OBJECT_TRIGGER_FEATURE_ID =
  "animations.setObjectTrigger" as const;

/** Strict object-click trigger authority; media bookmark triggers remain open. */
export const PRESENTATION_ANIMATION_OBJECT_TRIGGER_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_ANIMATION_OBJECT_TRIGGER_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["animations"],
          entry: "ribbon",
          group: "Advanced Animation",
          order: 2,
          contextualPredicate: "editable",
          controlKind: "menu",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.setAnimationObjectTrigger",
          handlerRef: "PresentationWindow#setAnimationObjectTrigger",
          undoPolicy: "single-transaction",
          availabilityPredicate: "editable",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: { mode: "sequence" },
          limits: {
            modes: ["sequence", "on-click-of-object"],
            target: "one-live-exact-animation-occurrence",
            triggerObject:
              "one-distinct-live-same-slide-native-shape-identity",
            bookmarkTriggers: "unsupported-no-bookmark-writer",
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: { powerpoint: "editable", wps: "unsupported" },
      },
    ],
  });

const PRESENTATION_ANIMATION_OBJECT_TRIGGER_FEATURE_CAPABILITY = leaf({
  id: PRESENTATION_ANIMATION_OBJECT_TRIGGER_FEATURE_ID,
  tabs: ["animations"],
  entry: "ribbon",
  authoring: "native",
  pptxImport: "native",
  fromScratchExport: "native",
  patchbackExport: "native",
  powerpoint: "editable",
  wps: "unsupported",
  commandKey: "presentation.command.setAnimationObjectTrigger",
  contextPredicate: "editable",
  reasonKey: "presentation.capability.editableRequired",
});

/** Strict native row for the bounded Hide Slide / Show Slide toggle. */
export const PRESENTATION_SLIDE_VISIBILITY_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_SLIDE_VISIBILITY_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["slideshow"],
          entry: "ribbon",
          group: "Set Up",
          order: 0,
          contextualPredicate: "editable",
          controlKind: "toggle-button",
        },
        command: {
          disposition: "bound",
          key: "presentation.command.toggleSlideVisibility",
          handlerRef: "PresentationWindow#toggleSlideVisibility",
          undoPolicy: "single-transaction",
          availabilityPredicate: "editable",
        },
        authoring: { disposition: "native" },
        semantics: {
          defaults: {
            hidden: false,
            visibleNativeEncoding: "p:sld@show omitted",
            hiddenNativeEncoding: "p:sld@show=0",
          },
          limits: {
            target: "one-live-active-editable-slide",
            legacyImportedWithoutReceipt: "preserve-only",
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "typed",
          persistence: "document",
        },
        import: {
          disposition: "native",
        },
        fromScratchExport: {
          disposition: "native",
        },
        patchbackExport: {
          disposition: "native",
        },
        officeOutcome: {
          powerpoint: "editable",
          wps: "unsupported",
        },
      },
    ],
  });

export const PRESENTATION_ACCESSIBILITY_CHECK_FEATURE_ID =
  "review.accessibilityCheck" as const;
export const PRESENTATION_ACCESSIBILITY_CHECK_ACTION =
  "presentation.review.check-accessibility" as const;
export const PRESENTATION_ACCESSIBILITY_CHECK_HANDLER_REF =
  "PresentationWindow#checkAccessibility" as const;

/**
 * Strict view-only authority for Review → Check Accessibility. The scanner
 * consumes the live MRAI projection but owns no import or writer authority.
 */
export const PRESENTATION_ACCESSIBILITY_CHECK_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_ACCESSIBILITY_CHECK_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["review"],
          entry: "ribbon",
          group: "Accessibility",
          order: 0,
          contextualPredicate: "always",
          controlKind: "button-pane",
        },
        command: {
          disposition: "bound",
          key: "presentation.command.checkAccessibility",
          handlerRef: PRESENTATION_ACCESSIBILITY_CHECK_HANDLER_REF,
          undoPolicy: "view-only",
          availabilityPredicate: "always",
        },
        authoring: { disposition: "projection" },
        semantics: {
          defaults: {
            paneOpen: false,
            issueLimit: 500,
            scanMode: "manual",
          },
          limits: {
            deterministicRules: [
              "missing-alt-text",
              "missing-or-duplicate-slide-title",
              "table-header-or-merged-cells",
              "spatial-reading-order",
              "explicit-solid-color-contrast",
              "media-captions-manual-review",
            ],
            opaqueOrAmbiguousEvidence: "skip",
            repairAuthority: [
              "existing-slide-title-editor",
              "existing-picture-alt-text-command",
              "existing-table-header-row-writer",
              "existing-selection-pane-reading-order",
            ],
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "view-only",
          persistence: "session",
        },
        import: {
          disposition: "none",
        },
        fromScratchExport: {
          disposition: "reject",
        },
        patchbackExport: {
          disposition: "reject",
        },
        officeOutcome: {
          powerpoint: "visible",
          wps: "unsupported",
        },
      },
    ],
  });

const PRESENTATION_ACCESSIBILITY_CHECK_FEATURE_CAPABILITY = leaf({
  id: PRESENTATION_ACCESSIBILITY_CHECK_FEATURE_ID,
  tabs: ["review"],
  entry: "ribbon",
  authoring: "projection",
  pptxImport: "none",
  fromScratchExport: "reject",
  patchbackExport: "reject",
  powerpoint: "visible",
  wps: "unsupported",
  commandKey: "presentation.command.checkAccessibility",
  contextPredicate: "always",
});

export const PRESENTATION_VIEW_SURFACE_FEATURE_ID =
  "view.activateRibbonTab" as const;
export const PRESENTATION_VIEW_SURFACE_ACTION =
  "presentation.view.activate-ribbon-tab" as const;
export const PRESENTATION_VIEW_SURFACE_HANDLER_REF =
  "PresentationWindow#setActiveRibbonTab(view)" as const;

/**
 * Exact view-only authority for the View ribbon root. The five presentation
 * views and their status aliases are owned by PPT-VIEW-PRESENTATIONVIEWS-01;
 * this catalog row proves the frozen parent tab identity and delegates to that
 * already-authenticated dispatcher without introducing another view model.
 */
export const PRESENTATION_VIEW_SURFACE_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_VIEW_SURFACE_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["view"],
          entry: "ribbon",
          group: "Ribbon tabs",
          order: 9,
          contextualPredicate: "always",
          controlKind: "tab",
        },
        command: {
          disposition: "bound",
          key: PRESENTATION_VIEW_SURFACE_ACTION,
          handlerRef: PRESENTATION_VIEW_SURFACE_HANDLER_REF,
          undoPolicy: "view-only",
          availabilityPredicate: "always",
        },
        authoring: { disposition: "projection" },
        semantics: {
          defaults: { selected: false, defaultDocumentView: "normal" },
          limits: {
            dispatcher: "PPT-VIEW-PRESENTATIONVIEWS-01",
            ribbonModes: [
              "normal",
              "outline",
              "sorter",
              "notes-page",
              "reading",
            ],
            persistence: "session-only",
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "view-only",
          persistence: "session",
        },
        import: {
          disposition: "none",
        },
        fromScratchExport: {
          disposition: "reject",
        },
        patchbackExport: {
          disposition: "reject",
        },
        officeOutcome: {
          powerpoint: "visible",
          wps: "unsupported",
        },
      },
    ],
  });

const PRESENTATION_VIEW_SURFACE_FEATURE_CAPABILITY = leaf({
  id: PRESENTATION_VIEW_SURFACE_FEATURE_ID,
  tabs: ["view"],
  entry: "ribbon",
  authoring: "projection",
  pptxImport: "none",
  fromScratchExport: "reject",
  patchbackExport: "reject",
  powerpoint: "visible",
  wps: "unsupported",
  commandKey: PRESENTATION_VIEW_SURFACE_ACTION,
  contextPredicate: "always",
});

export const PRESENTATION_VIEW_AID_FEATURE_ID = "view.gridGuides" as const;
export const PRESENTATION_VIEW_AID_KINDS = Object.freeze([
  "ruler",
  "gridlines",
  "guides",
] as const);
export type PresentationViewAidKind =
  (typeof PRESENTATION_VIEW_AID_KINDS)[number];
export const PRESENTATION_VIEW_AID_ACTION =
  "presentation.view.toggle-view-aid" as const;
export const PRESENTATION_VIEW_AID_HANDLER_REF =
  "PresentationWindow#togglePresentationViewAid" as const;

const PRESENTATION_VIEW_AID_KIND_SET: ReadonlySet<string> = new Set(
  PRESENTATION_VIEW_AID_KINDS,
);

export function isPresentationViewAidKind(
  value: unknown,
): value is PresentationViewAidKind {
  return typeof value === "string" && PRESENTATION_VIEW_AID_KIND_SET.has(value);
}

/** Strict view-only authority for View → Show → Ruler/Gridlines/Guides. */
export const PRESENTATION_VIEW_AID_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_VIEW_AID_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["view"],
          entry: "both",
          group: "Show",
          order: 0,
          contextualPredicate: "always",
          controlKind: "parameterized-checkbox-triplet",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.toggleViewAid",
          handlerRef: PRESENTATION_VIEW_AID_HANDLER_REF,
          undoPolicy: "view-only",
          availabilityPredicate: "always",
        },
        authoring: { disposition: "projection" },
        semantics: {
          defaults: {
            ruler: false,
            gridlines: false,
            guides: false,
          },
          limits: {
            kinds: PRESENTATION_VIEW_AID_KINDS,
            availableView: "normal",
            gridMajorSpacingEmu: 914_400,
            guideGeometry: "one-horizontal-and-one-vertical-at-slide-center",
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "view-only",
          persistence: "session",
        },
        import: {
          disposition: "none",
        },
        fromScratchExport: {
          disposition: "reject",
        },
        patchbackExport: {
          disposition: "reject",
        },
        officeOutcome: {
          powerpoint: "visible",
          wps: "unsupported",
        },
      },
    ],
  });

const PRESENTATION_VIEW_AID_FEATURE_CAPABILITY = leaf({
  id: PRESENTATION_VIEW_AID_FEATURE_ID,
  tabs: ["view"],
  entry: "both",
  authoring: "projection",
  pptxImport: "none",
  fromScratchExport: "reject",
  patchbackExport: "reject",
  powerpoint: "visible",
  wps: "unsupported",
  commandKey: "presentation.command.toggleViewAid",
  contextPredicate: "always",
});

export const PRESENTATION_VIEW_NOTES_FEATURE_ID = "view.notesPane" as const;

const PRESENTATION_VIEW_NOTES_FEATURE_CAPABILITY = leaf({
  id: PRESENTATION_VIEW_NOTES_FEATURE_ID,
  tabs: ["view"],
  entry: "both",
  authoring: "projection",
  pptxImport: "none",
  fromScratchExport: "reject",
  patchbackExport: "reject",
  powerpoint: "visible",
  wps: "unsupported",
  commandKey: "presentation.command.toggleNotesPane",
  contextPredicate: "always",
});

export const PRESENTATION_VIEW_ZOOM_FEATURE_ID = "view.zoomFit" as const;
export const PRESENTATION_VIEW_ZOOM_FIT_CONTROL_ID =
  "PPTCTL.RIBBON.VIEW.ZOOM.FIT" as const;
export const PRESENTATION_VIEW_ZOOM_ACTION =
  "presentation.view.set-viewport-zoom" as const;
export const PRESENTATION_VIEW_ZOOM_HANDLER_REF =
  "PresentationWindow#setPresentationViewportZoom" as const;
export const PRESENTATION_VIEW_ZOOM_MIN_PERCENT = 10;
export const PRESENTATION_VIEW_ZOOM_MAX_PERCENT = 400;

export type PresentationViewportZoomRequest =
  Readonly<{ kind: "fit" }> | Readonly<{ kind: "percent"; percent: number }>;

export function isPresentationViewportZoomRequest(
  value: unknown,
): value is PresentationViewportZoomRequest {
  if (value === null || typeof value !== "object" || Array.isArray(value)) {
    return false;
  }
  const candidate = value as { kind?: unknown; percent?: unknown };
  const keys = Reflect.ownKeys(candidate);
  if (candidate.kind === "fit") {
    return keys.length === 1 && keys[0] === "kind";
  }
  return (
    candidate.kind === "percent" &&
    keys.length === 2 &&
    keys.includes("kind") &&
    keys.includes("percent") &&
    typeof candidate.percent === "number" &&
    Number.isFinite(candidate.percent) &&
    candidate.percent >= PRESENTATION_VIEW_ZOOM_MIN_PERCENT &&
    candidate.percent <= PRESENTATION_VIEW_ZOOM_MAX_PERCENT
  );
}

/** Strict view-only authority for View → Zoom → Fit to Window. */
export const PRESENTATION_VIEW_ZOOM_CAPABILITY_CATALOG =
  readPresentationFeatureCapabilityCatalog({
    version: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: [
      {
        featureId: PRESENTATION_VIEW_ZOOM_FEATURE_ID,
        sourceVersion: PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION,
        surface: {
          disposition: "cataloged",
          tabs: ["view"],
          entry: "both",
          group: "Zoom",
          order: 0,
          contextualPredicate: "always",
          controlKind: "button",
        },
        command: {
          disposition: "parameterized",
          key: "presentation.command.setViewportZoom",
          handlerRef: PRESENTATION_VIEW_ZOOM_HANDLER_REF,
          undoPolicy: "view-only",
          availabilityPredicate: "always",
        },
        authoring: { disposition: "projection" },
        semantics: {
          defaults: { mode: "percent", percent: 100 },
          limits: {
            minPercent: PRESENTATION_VIEW_ZOOM_MIN_PERCENT,
            maxPercent: PRESENTATION_VIEW_ZOOM_MAX_PERCENT,
            percentStep: 0.01,
            fitView: "normal",
          },
          shortcutIds: [],
        },
        domain: {
          disposition: "view-only",
          persistence: "session",
        },
        import: {
          disposition: "none",
        },
        fromScratchExport: {
          disposition: "reject",
        },
        patchbackExport: {
          disposition: "reject",
        },
        officeOutcome: {
          powerpoint: "visible",
          wps: "unsupported",
        },
      },
    ],
  });

const PRESENTATION_VIEW_ZOOM_FEATURE_CAPABILITY = leaf({
  id: PRESENTATION_VIEW_ZOOM_FEATURE_ID,
  tabs: ["view"],
  entry: "ribbon",
  authoring: "projection",
  pptxImport: "none",
  fromScratchExport: "reject",
  patchbackExport: "reject",
  powerpoint: "visible",
  wps: "unsupported",
  commandKey: "presentation.command.setViewportZoom",
  contextPredicate: "always",
});

export function findPresentationFeatureCapabilityContract(
  id: string,
  catalog: PresentationFeatureCapabilityCatalog = PRESENTATION_FEATURE_CAPABILITY_CATALOG,
): PresentationFeatureCapabilityContract | null {
  return catalog.features.find((feature) => feature.featureId === id) ?? null;
}

function unknownPresentationFeatureCapability(
  featureId: string,
): PresentationFeatureCapabilityContract {
  return Object.freeze({
    featureId,
    sourceVersion: "unregistered",
    officeBaseline: Object.freeze({
      platform: null,
      version: null,
      evidenceRefs: Object.freeze([]),
    }),
    surface: Object.freeze({
      disposition: "unknown",
      tabs: Object.freeze([]),
      entry: "unknown",
      group: null,
      order: null,
      contextualPredicate: "unknown",
      controlKind: null,
    }),
    command: Object.freeze({
      disposition: "unknown",
      key: null,
      handlerRef: null,
      undoPolicy: "unknown",
      availabilityPredicate: "unknown",
    }),
    authoring: Object.freeze({ disposition: "unknown" }),
    semantics: Object.freeze({
      defaults: "unknown",
      limits: "unknown",
      interactions: Object.freeze([]),
      shortcutIds: Object.freeze([]),
    }),
    domain: Object.freeze({
      disposition: "unknown",
      schemaPaths: Object.freeze([]),
      version: null,
      persistence: "unknown",
    }),
    import: Object.freeze({
      disposition: "unknown",
      projectionRefs: Object.freeze([]),
      preservedUnknownParts: "unknown",
      lossCodes: Object.freeze([]),
      flags: Object.freeze([]),
    }),
    fromScratchExport: Object.freeze({
      disposition: "unknown",
      serializerRefs: Object.freeze([]),
      lossCodes: Object.freeze([]),
      flags: Object.freeze([]),
    }),
    patchbackExport: Object.freeze({
      disposition: "unknown",
      serializerRefs: Object.freeze([]),
      preservePolicy: null,
      lossCodes: Object.freeze([]),
      flags: Object.freeze([]),
    }),
    officeOutcome: Object.freeze({
      powerpoint: "unknown",
      wps: "unknown",
    }),
    acceptanceContractIds: Object.freeze([]),
    owner: null,
  });
}

/** Missing catalog rows resolve to a complete, inert unknown contract. */
export function resolvePresentationFeatureCapabilityContract(
  id: string,
  catalog: PresentationFeatureCapabilityCatalog = PRESENTATION_FEATURE_CAPABILITY_CATALOG,
): PresentationFeatureCapabilityContract {
  return (
    findPresentationFeatureCapabilityContract(id, catalog) ??
    unknownPresentationFeatureCapability(id)
  );
}

function isCatalogOwnedPresentationFeatureCapability(
  capability: PresentationFeatureCapabilityContract,
  catalog: PresentationFeatureCapabilityCatalog,
): boolean {
  return (
    validatedPresentationFeatureCapabilityCatalogs.has(catalog) &&
    catalog.features.some((feature) => feature === capability)
  );
}

/**
 * A new-catalog feature is executable only after handler/domain verification.
 * (Its evidence -- Office baseline, acceptance ids -- is audit truth, checked
 * for every row by the audit module's tests, not carried by the browser.)
 */
export function isPresentationFeatureContractExecutable(
  capability: PresentationFeatureCapabilityContract,
  catalog: PresentationFeatureCapabilityCatalog = PRESENTATION_FEATURE_CAPABILITY_CATALOG,
): boolean {
  return (
    isCatalogOwnedPresentationFeatureCapability(capability, catalog) &&
    capability.sourceVersion ===
      PRESENTATION_CAPABILITY_CATALOG_SCHEMA_VERSION &&
    catalog.handlerBindingsVerified &&
    capability.surface.disposition === "cataloged" &&
    capability.surface.contextualPredicate !== "unknown" &&
    (capability.command.disposition === "bound" ||
      capability.command.disposition === "parameterized") &&
    capability.command.key !== null &&
    capability.command.handlerRef !== null &&
    capability.command.undoPolicy !== "unknown" &&
    capability.command.availabilityPredicate !== "unknown" &&
    capability.authoring.disposition !== "unknown" &&
    capability.authoring.disposition !== "unsupported" &&
    capability.domain.disposition !== "unknown"
  );
}

export function findPresentationFeatureCapability(
  id: string,
): PresentationFeatureCapability | null {
  return (
    PRESENTATION_FEATURE_CAPABILITY_REGISTRY.find((entry) => entry.id === id) ??
    (id === PRESENTATION_VIDEO_FILE_FEATURE_ID
      ? PRESENTATION_VIDEO_FILE_FEATURE_CAPABILITY
      : id === PRESENTATION_RECORD_AUDIO_FEATURE_ID
      ? PRESENTATION_RECORD_AUDIO_FEATURE_CAPABILITY
      : id === PRESENTATION_AUDIO_FILE_FEATURE_ID
      ? PRESENTATION_AUDIO_FILE_FEATURE_CAPABILITY
      : id === PRESENTATION_ANIMATION_EFFECT_FEATURE_ID
      ? PRESENTATION_ANIMATION_EFFECT_FEATURE_CAPABILITY
      : id === PRESENTATION_ANIMATION_EFFECT_OPTION_FEATURE_ID
        ? PRESENTATION_ANIMATION_EFFECT_OPTION_FEATURE_CAPABILITY
        : id === PRESENTATION_ANIMATION_SEQUENCE_FEATURE_ID
          ? PRESENTATION_ANIMATION_SEQUENCE_FEATURE_CAPABILITY
        : id === PRESENTATION_ANIMATION_TIMING_FEATURE_ID
          ? PRESENTATION_ANIMATION_TIMING_FEATURE_CAPABILITY
          : id === PRESENTATION_ANIMATION_OBJECT_TRIGGER_FEATURE_ID
            ? PRESENTATION_ANIMATION_OBJECT_TRIGGER_FEATURE_CAPABILITY
          : id === PRESENTATION_ACCESSIBILITY_CHECK_FEATURE_ID
            ? PRESENTATION_ACCESSIBILITY_CHECK_FEATURE_CAPABILITY
            : id === PRESENTATION_VIEW_SURFACE_FEATURE_ID
              ? PRESENTATION_VIEW_SURFACE_FEATURE_CAPABILITY
              : id === PRESENTATION_VIEW_AID_FEATURE_ID
                ? PRESENTATION_VIEW_AID_FEATURE_CAPABILITY
                : id === PRESENTATION_VIEW_NOTES_FEATURE_ID
                  ? PRESENTATION_VIEW_NOTES_FEATURE_CAPABILITY
                  : id === PRESENTATION_VIEW_ZOOM_FEATURE_ID
                    ? PRESENTATION_VIEW_ZOOM_FEATURE_CAPABILITY
                    : null) ??
    null
  );
}

export function isPresentationFeatureAvailable(
  capability: PresentationFeatureCapability,
  context: PresentationCapabilityContext,
): boolean {
  // Context can further restrict a real command, but it must never make a
  // catalog-only leaf executable merely because its predicate is `always`.
  if (capability.authoring === "unsupported") return false;
  switch (capability.contextPredicate ?? "always") {
    case "always":
      return true;
    case "editable":
      return context.editable;
    case "selection":
      return context.editable && context.hasSelection;
    case "textSelection":
      return context.editable && context.hasTextSelection;
    case "paragraphSelection":
      return context.editable && (context.hasTextSelection || context.hasParagraphSelection === true);
    case "shapeSelection":
      return context.editable && context.hasShapeSelection;
    case "clipboard":
      return context.editable && context.hasClipboard;
    case "online":
      return context.editable && context.online;
  }
}

/** One Ribbon command a catalog names, and the tabs that host it. */
export interface PresentationRibbonCommandEntry {
  readonly featureId: string;
  readonly commandKey: string;
  readonly tabs: readonly PresentationRibbonTabId[];
}

let presentationRibbonCommandEntriesCache:
  | readonly PresentationRibbonCommandEntry[]
  | null = null;

/**
 * Every Ribbon command the capability catalogs name, once per command. The
 * command search (Alt+Q) finds any of them, as PowerPoint's Search finds a
 * command on any tab; one without a direct handler there opens its tab.
 */
export function presentationRibbonCommandEntries(): readonly PresentationRibbonCommandEntry[] {
  if (presentationRibbonCommandEntriesCache) {
    return presentationRibbonCommandEntriesCache;
  }
  const entries: PresentationRibbonCommandEntry[] = [];
  const seen = new Set<string>();
  const add = (
    featureId: string,
    commandKey: string | null,
    tabs: readonly PresentationRibbonTabId[],
  ) => {
    if (!commandKey?.startsWith("presentation.command.") || tabs.length === 0) {
      return;
    }
    if (seen.has(featureId) || seen.has(commandKey)) return;
    seen.add(featureId);
    seen.add(commandKey);
    entries.push(Object.freeze({ featureId, commandKey, tabs }));
  };
  for (const legacy of [
    ...PRESENTATION_FEATURE_CAPABILITY_REGISTRY,
    PRESENTATION_AUDIO_FILE_FEATURE_CAPABILITY,
    PRESENTATION_RECORD_AUDIO_FEATURE_CAPABILITY,
    PRESENTATION_VIDEO_FILE_FEATURE_CAPABILITY,
    PRESENTATION_ANIMATION_EFFECT_FEATURE_CAPABILITY,
    PRESENTATION_ANIMATION_EFFECT_OPTION_FEATURE_CAPABILITY,
    PRESENTATION_ANIMATION_SEQUENCE_FEATURE_CAPABILITY,
    PRESENTATION_ANIMATION_TIMING_FEATURE_CAPABILITY,
    PRESENTATION_ANIMATION_OBJECT_TRIGGER_FEATURE_CAPABILITY,
    PRESENTATION_ACCESSIBILITY_CHECK_FEATURE_CAPABILITY,
    PRESENTATION_VIEW_SURFACE_FEATURE_CAPABILITY,
    PRESENTATION_VIEW_AID_FEATURE_CAPABILITY,
    PRESENTATION_VIEW_NOTES_FEATURE_CAPABILITY,
    PRESENTATION_VIEW_ZOOM_FEATURE_CAPABILITY,
  ]) {
    add(legacy.id, legacy.commandKey, legacy.tabs);
  }
  for (const catalog of [
    PRESENTATION_IMAGE_FILE_CAPABILITY_CATALOG,
    PRESENTATION_AUDIO_FILE_CAPABILITY_CATALOG,
    PRESENTATION_RECORD_AUDIO_CAPABILITY_CATALOG,
    PRESENTATION_VIDEO_FILE_CAPABILITY_CATALOG,
    PRESENTATION_PICTURE_REPLACE_FILE_CAPABILITY_CATALOG,
    PRESENTATION_PICTURE_CROP_CAPABILITY_CATALOG,
    PRESENTATION_PICTURE_ALT_TEXT_CAPABILITY_CATALOG,
    PRESENTATION_PICTURE_BORDER_CAPABILITY_CATALOG,
    PRESENTATION_PICTURE_TRANSPARENCY_CAPABILITY_CATALOG,
    PRESENTATION_PICTURE_CORRECTIONS_CAPABILITY_CATALOG,
    PRESENTATION_PICTURE_COLOR_CAPABILITY_CATALOG,
    PRESENTATION_PICTURE_ARRANGE_STEP_CAPABILITY_CATALOG,
    PRESENTATION_PICTURE_SIZE_CAPABILITY_CATALOG,
    PRESENTATION_SLIDE_TRANSITION_CAPABILITY_CATALOG,
    PRESENTATION_SLIDE_TRANSITION_DURATION_CAPABILITY_CATALOG,
    PRESENTATION_SLIDE_TRANSITION_ADVANCE_CAPABILITY_CATALOG,
    PRESENTATION_SLIDE_TRANSITION_SOUND_CAPABILITY_CATALOG,
    PRESENTATION_SLIDE_TRANSITION_APPLY_ALL_CAPABILITY_CATALOG,
    PRESENTATION_ANIMATION_EFFECT_CAPABILITY_CATALOG,
    PRESENTATION_ANIMATION_EFFECT_OPTION_CAPABILITY_CATALOG,
    PRESENTATION_ANIMATION_SEQUENCE_CAPABILITY_CATALOG,
    PRESENTATION_ANIMATION_TIMING_CAPABILITY_CATALOG,
    PRESENTATION_ANIMATION_OBJECT_TRIGGER_CAPABILITY_CATALOG,
    PRESENTATION_SLIDE_VISIBILITY_CAPABILITY_CATALOG,
    PRESENTATION_ACCESSIBILITY_CHECK_CAPABILITY_CATALOG,
    PRESENTATION_VIEW_SURFACE_CAPABILITY_CATALOG,
    PRESENTATION_VIEW_AID_CAPABILITY_CATALOG,
    PRESENTATION_VIEW_ZOOM_CAPABILITY_CATALOG,
  ]) {
    for (const feature of catalog.features) {
      add(feature.featureId, feature.command.key, feature.surface.tabs);
    }
  }
  presentationRibbonCommandEntriesCache = Object.freeze(entries);
  return presentationRibbonCommandEntriesCache;
}
