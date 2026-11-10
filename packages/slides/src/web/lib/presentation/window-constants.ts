/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Fixed values of the presentation editor window: history depth, geometry
 * and snap thresholds, z-indices, teleprompter font steps, icon glyph and
 * SmartArt class maps, pane chrome classes, the AI block selector, commit
 * markers and help links.
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */

import presentationChrome from "@mrnote/core/web/lib/presentation/presentation-window-shell.module.css";
import type {
  OfficeSmartArtDesignTabClasses,
} from "@mrnote/core/web/components/office-ribbon/OfficeSmartArtDesignTab.tsx";
import type {
} from "@mrnote/core/web/components/office-ribbon/OfficeSmartArtTextPane.tsx";
import { officeDialogClassName } from "@mrnote/core/web/components/office-ribbon/officeDialogChrome.tsx";
import { officeFontSizeOptions } from "@mrnote/core/web/components/office-ribbon/officeFontSizeCatalog.ts";
import { isEditorCopilotSurfaceEnabled } from "@mrnote/core/web/__cut__/mc2523d29ba.ts";

/** The shapes on the active slide and the thumbnails in the rail: what an AI write changes. */
export const PRESENTATION_AI_BLOCK_SELECTOR =
  "[data-testid='presentation-slide-canvas'] [data-block-id], [data-slide-id]";

/**
 * The chrome a self-drawn task pane of the deck wears (Format Background,
 * Master Properties), as the shared OfficeTaskPane does under its chrome
 * provider: the module keeps docking and size, the chrome draws the card.
 */
export const PRESENTATION_PANE_CHROME = officeDialogClassName({
  tone: "presentation",
  pane: true,
});

/** PowerPoint's Font Size list, from the shared Office catalog. */
export const PRESENTATION_FONT_SIZE_OPTIONS = officeFontSizeOptions("powerpoint");

/** A chart workbook's formula results filling in cells already committed:
 * never an undo step of its own (Undo returns to before the edit). */
export const CHART_WORKBOOK_RESULTS_COMMIT = "chart-workbook-results";

/** A page written again only so its media table holds a picture an Undo
 * restored as a token (lib/presentation-page-media): no undo step either. */
export const PAGE_MEDIA_REPAIR_COMMIT = "page-media-repair";



export const TEXT_LAYER_REPAIR_COMMIT = "text-layer-repair";

/** An equation box taking the size its equation needs (a derived correction
 * of its frame, never an edit): neither an undo step nor a redo breaker. */
export const EQUATION_BOX_FIT_COMMIT = "equation-box-fit";



export const TABLE_FRAME_FIT_COMMIT = "table-frame-fit";

/** A text box's frame following its words (Resize shape to fit text,
 * `presentation-text-fit`): part of the edit that changed the words, never an
 * undo step of its own. */
export const TEXT_BOX_FIT_COMMIT = "text-box-fit";

/** The slide's selection chrome (the selected block's frame, handles and
 * grips) drawn in a layer of its own over every block, so a block in front
 * of the selected one never hides them (PowerPoint draws them on top). Each
 * block is its own stacking context at z 0; this sits above all of them and
 * under the slide's overlays (comment markers 80, guides, previews). */
export const PRESENTATION_SELECTION_CHROME_LAYER = 20;

/** Erase's brush: a share of the picture's shorter side (Edit Picture ▸ Erase). */
export const PICTURE_ERASE_RADIUS = 0.04;

export const PRESENTATION_COPILOT_V1_ENABLED =
  isEditorCopilotSurfaceEnabled("presentation");

/** Bounded undo/redo history depth — caps in-memory snapshot retention. */
export const MAX_PRESENTATION_HISTORY = 50;

export const PRESENTATION_SHAPE_CATALOG_VIEWPORT_PADDING = 16;

export const PRESENTATION_SHAPE_CATALOG_ALIGN_OFFSET = -12;

/** On-screen snap tolerance (px) used to scale the EMU snap threshold to the
 * current canvas size, so the snap feels consistent regardless of slide scale. */
export const SNAP_THRESHOLD_PX = 6;

// The SmartArt diagram frame sits this many px outside its members' union rect
// so the frame's handles and drag edges never overlap a member's own handles.
export const SMARTART_FRAME_PAD_PX = 10;

/** Minimum pointer travel (px) before an empty-canvas press becomes a marquee
 * drag. A press that never moves this far is treated as a plain click (clear
 * selection), so a stray click never leaves a stuck zero-size selection box. */
export const MARQUEE_DRAG_THRESHOLD_PX = 4;

export const PRESENTATION_RECORD_HELP_URL =
  "https://support.microsoft.com/en-us/powerpoint/record-your-presentation";

const CHARACTER_SPACING_LABELS: Readonly<Record<number, string>> = {
  [-1]: "非常紧凑",
  [-0.5]: "紧凑",
  [0]: "常规",
  [0.5]: "松散",
  [1]: "非常松散",
  [2]: "加宽 2 磅",
  [4]: "加宽 4 磅",
};

/**
 * Hard cap for one slide's speaker notes — generous for a presenter script
 * (the presenter console renders it scrollable) while bounding the payload.
 */
export const TELEPROMPTER_FONT_SIZE_MIN = 16;

export const TELEPROMPTER_FONT_SIZE_MAX = 36;

export const TELEPROMPTER_FONT_SIZE_STEP = 2;

export const TELEPROMPTER_FONT_SIZE_DEFAULT = 24;



// Layout names, descriptions and categories come from the one shared catalog
// (`lib/office-smartart/catalog.ts`): `officeSmartArtLayout(id).nameKey` etc.
// Presentation keeps no bridge table of its own.

/**
 * Presentation's own SmartArt Design tab geometry, handed to the shared tab.
 *
 * The tab itself ships the same numbers in
 * `components/office-ribbon/OfficeSmartArtDesignTab.module.css` for a host that
 * has none; Presentation keeps passing its own because
 * `tests/unit/presentation-ribbon-more-css.test.ts` pins these rules in the
 * presentation shell stylesheet.
 */
export const PRESENTATION_SMARTART_DESIGN_TAB_CLASSES = {
  createGroup: presentationChrome.smartArtCreateGroup,
  createRow: presentationChrome.smartArtCreateRow,
  createHierarchyLane: presentationChrome.smartArtCreateHierarchyLane,
  layoutsGroup: presentationChrome.smartArtLayoutsGroup,
  layoutGallery: presentationChrome.smartArtLayoutGallery,
  stylesGroup: presentationChrome.smartArtStylesGroup,
  styleGallery: presentationChrome.smartArtStyleGallery,
  resetGroup: presentationChrome.smartArtResetGroup,
  propertiesGroup: presentationChrome.smartArtPropertiesGroup,
} satisfies OfficeSmartArtDesignTabClasses;


// Preview is a canvas-local isolation surface: persisted blocks top out below
// the editor's snap-guide chrome (9999), so these two layers mask the entire
// editable canvas while remaining below canvas-external menus/dialogs.
export const PRESENTATION_ANIMATION_PREVIEW_BACKDROP_Z_INDEX = 10_001;

export const PRESENTATION_ANIMATION_PREVIEW_LAYER_Z_INDEX = 10_002;

/** History key prefix of a content placeholder's picture insertion (PPT-B375). */
export const PLACEHOLDER_PICTURE_COMMIT = "placeholder-picture:";
/**
 * Designer ideas tried one after another on a slide are one history step:
 * the next replaces the last, and Undo returns the slide as it was (SPEC
 * designer-ms-parity §1.1). Followed by the slide id and the ideas' round.
 */
export const DESIGNER_IDEA_COMMIT = "designer-idea:";
/**
 * Home > Layout on an imported slide, and the new layout's graphics drawn on
 * it once its card arrives (`slideWithNativeLayoutGraphics`): one history
 * step. Followed by the slide id and the layout part.
 */
export const LAYOUT_GRAPHICS_COMMIT = "layout-graphics:";

/**
 * The Office themes Designer suggests for a new deck's title slide: ones
 * whose title slides carry a strong look of their own (artwork, colour
 * fields), as PowerPoint's first suggestions do.
 */
export const DESIGNER_TITLE_THEMES: readonly string[] = Object.freeze([
  "facet", "ion", "integral", "wisp", "dividend", "berlin", "celestial", "gallery", "circuit", "savon",
]);

/** The Alt Text pane's box: the longest description a picture keeps
 * (normalizeImageAltText), for every object the pane describes. */
export const PRESENTATION_ALT_TEXT_MAX_LENGTH = 1_000;
