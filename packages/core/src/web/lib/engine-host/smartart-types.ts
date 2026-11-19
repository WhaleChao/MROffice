/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The SmartArt engine's wire: what the browser sends to
 * `POST /api/v1/engine/smartart/{method}` and what comes back. Types only --
 * the engine (`engine-host/engines/smartart.ts`) runs on the server.
 *
 * Ids. A pure method cannot mint random ids, so where a result needs NEW ids
 * the engine returns placeholders ({@link SMARTART_ENGINE_ID_PREFIX} + a
 * number) and the browser replaces each distinct placeholder with a fresh id
 * (`lib/office-smartart/remote.ts`). A request may instead pass `idPrefix`
 * when the caller needs deterministic ids (the Copilot's `copilot-…-art-N`);
 * the engine then numbers from 0 in the order the layout asks, exactly as the
 * counter the caller used to pass did.
 */

import type { DocumentShapeAttrs, DocumentShapeAttrsInput } from "../document-shapes";
import type { DocumentSmartArtPatch } from "../office-smartart/document-design-model";
import type { DocumentSmartArtProjection } from "../office-smartart/document-shapes-model";
import type { SmartArtColorTheme } from "../office-smartart/colors";
import type { OfficeSmartArtFrame } from "../office-smartart/design-model";
import type {
  OfficeSmartArtHostNode,
  OfficeSmartArtNode,
  SmartArtBounds,
  SmartArtDirection,
  SmartArtKind,
  SmartArtNodePaint,
  SmartArtStyle,
} from "../office-smartart/model";
import type { NativeSmartArtFrame } from "../office-smartart/native-convert-model";
import type { NativeSmartArtOutlineItem } from "@/lib/office-smartart/native-outline";
import type { SmartArtOutlineItem } from "../office-smartart/outline";
import type { SmartArtIdentity } from "../office-smartart/smartart-identity";
import type { SmartArtStyleId } from "../office-smartart/styles-model";
import type { SheetSmartArtRebuild } from "../spreadsheet/smartart-design";
import type { SpreadsheetSmartArtMember } from "../spreadsheet/smart-art";

/**
 * A new id the browser must replace. It is spelled as an ordinary identifier
 * because the page models validate ids inside the engine's own pipeline (a
 * Docs member, an identity record), and it carries a fixed random tag so no
 * author text or real id is ever mistaken for one.
 */
export const SMARTART_ENGINE_ID_PREFIX = "smartart-engine-new-7c3f91e2-";

/** The host's localized placeholder words: `itemLabels[i - 1]` is item i's. */
export interface SmartArtPlaceholdersWire {
  itemLabels: string[];
  rootLabel: string;
}

export interface SmartArtBuildRequest {
  /** The workbook's / document's body face (`+mn-lt`), which sets how tall the labels' lines stand. */
  bodyFace?: string | null;
  kind: SmartArtKind;
  items: string[];
  levels?: number[];
  /** The box the layout is built in (EMU). */
  bounds: SmartArtBounds;
  style?: SmartArtStyle;
  styleId?: SmartArtStyleId;
  assistants?: boolean[];
  nodePaint?: (SmartArtNodePaint | null)[];
  direction?: SmartArtDirection;
  /**
   * Lay the diagram out in this box (Convert to SmartArt: the text block's
   * own place) instead of `bounds` at the origin; a box without a size falls
   * back to `bounds` (the deck's `smartArtFromLabels`).
   */
  geometry?: { left?: number; top?: number; width?: number; height?: number };
  idPrefix?: string;
  /** A Change Colors variation (Office's id, e.g. `colorful1`) the diagram is built in. */
  colorsId?: string | null;
  /** The theme colours `colorsId` resolves against (the deck's accents). */
  colorTheme?: Partial<SmartArtColorTheme> | null;
  /**
   * An Insert gallery alias card's Office alias of `kind` (Continuous Block
   * Process for a Basic Process): the diagram is laid out from that alias's
   * own definition and every member records it (`smartArtLayoutUrn`). A urn
   * that is no aligned alias of `kind` builds the catalog layout.
   */
  layoutUrn?: string | null;
}

export interface SmartArtPresentationRebuildRequest extends SmartArtLayoutDefinitionFields {
  members: OfficeSmartArtHostNode[];
  groupId: string;
  kind: SmartArtKind;
  /**
   * Change Layout's pick: an aligned Office alias of `kind` (its own gallery
   * entry), or `null` for the catalog layout itself. Omitted keeps the alias
   * the members record while `kind` stays theirs.
   */
  layoutUrn?: string | null;
  frame: OfficeSmartArtFrame;
  outline: SmartArtOutlineItem[];
  colors?: SmartArtStyle;
  styleId?: SmartArtStyleId;
  direction?: SmartArtDirection;
  allowItemChanges?: boolean;
  nodePaint?: (SmartArtNodePaint | null)[];
  keepAuthorPaint?: boolean;
  resetCustomGeometry?: boolean;
  /** `false`: the words drop the author's character formatting (Reset Graphic). */
  keepTextFormat?: boolean;
  /** The colour variation to rebuild in; `null` is Office's default, omitted keeps the members'. */
  colorsId?: string | null;
  /** Reset Graphic: restyle the rebuilt members at the catalog default after. */
  resetPaint?: boolean;
  /** The deck's theme colours, so a colour variation re-spreads over the new outline. */
  colorTheme?: Partial<SmartArtColorTheme> | null;
}

export interface SmartArtRestyleRequest extends SmartArtLayoutDefinitionFields {
  /** The diagram's members (only members of `groupId` change). */
  members: OfficeSmartArtHostNode[];
  groupId: string;
  colors?: SmartArtStyle | null;
  styleId?: SmartArtStyleId | null;
  authorPaint?: boolean;
  nodePaint?: (SmartArtNodePaint | null)[] | null;
  /** The colour variation applied (`smartArtColors`); null clears it. */
  colorsId?: string | null;
  /** The deck's theme colours a diagram drawn from Office's definitions is painted on. */
  colorTheme?: Partial<SmartArtColorTheme> | null;
}

/**
 * A diagram laid out from its own file's layout definition (SMARTART-GAP.md
 * §8.6). A caller names it by `layoutDefinitionRef` (or by the
 * `office-definition` members it carries); the engine client lends the part
 * itself as `layoutDefinition`, which is what the engine reads.
 */
export interface SmartArtLayoutDefinitionWire {
  urn: string;
  sha256: string;
  xml: string;
}

export interface SmartArtLayoutDefinitionFields {
  layoutDefinitionRef?: { urn: string; sha256: string };
  layoutDefinition?: SmartArtLayoutDefinitionWire;
}

export interface SmartArtNativeConvertRequest extends SmartArtLayoutDefinitionFields {
  layoutIdentity: string;
  outline: NativeSmartArtOutlineItem[];
  frameEmu: NativeSmartArtFrame;
  smartArtStyle?: unknown;
  direction?: unknown;
  catalogLayoutId?: string | null;
  /** MRAI's own identity record from the file; the engine proves it first. */
  sourceIdentity?: SmartArtIdentity | null;
  /** Group ids already on the page (the record's id is not reused then). */
  groupIdsInUse?: string[];
  images?: Record<string, string>;
  placeholders?: SmartArtPlaceholdersWire;
  /**
   * The diagram's colour variation as the file names it (the data part's
   * `csTypeId`, Office's id) and the deck's theme colours: the converted
   * diagram is painted the way PowerPoint paints it and keeps the variation.
   * Both omitted, the catalog palette, as before.
   */
  colorsId?: string | null;
  colorTheme?: Partial<SmartArtColorTheme> | null;
  /**
   * The diagram's file theme lists (`nativeSmartArtView.hostTheme`): what its
   * quick style's `a:fillRef` / `a:effectRef` draw when `colorTheme` brings
   * none of its own (a cache-less diagram printed on the server, TRIAGE S2b).
   */
  hostTheme?: unknown;
  /** The author's hand changes the projection read off the data part (checked by the engine). */
  customizations?: unknown;
  /** The graphic's own 3-D Rotation (`nativeSmartArtView.graphicScene3d`), checked by the engine. */
  graphicScene3d?: unknown;
  /** The diagram's own pictures (`nativeSmartArtView.docPictures`, readings row 1005), checked by the engine. */
  docPictures?: unknown;
}

export interface SmartArtSheetRebuildRequest extends SmartArtLayoutDefinitionFields {
  /** The workbook's / document's body face (`+mn-lt`), which sets how tall the labels' lines stand. */
  bodyFace?: string | null;
  /** The face the workbook theme gives Chinese words (`SmartArtColorTheme.eastAsianFace`). */
  eastAsianFace?: string | null;
  /** Change Layout's pick: an aligned Office alias, or `null` for the catalog layout. */
  layoutUrn?: string | null;
  members: OfficeSmartArtHostNode[];
  groupId: string;
  kind: SmartArtKind;
  frame: OfficeSmartArtFrame;
  outline: SmartArtOutlineItem[];
  colors?: SmartArtStyle;
  styleId?: SmartArtStyleId;
  direction?: SmartArtDirection;
  nodePaint?: (SmartArtNodePaint | null)[];
  allowItemChanges?: boolean;
  authorPaint?: boolean;
  /** Change Colors' variation (Office's id); `null` back to Office's default. */
  colorsId?: string | null;
  /** The workbook theme Office's variations are resolved against. */
  colorTheme?: { accents?: string[] } | null;
}

export interface SmartArtSheetConvertRequest extends SmartArtLayoutDefinitionFields {
  /** The workbook's / document's body face (`+mn-lt`), which sets how tall the labels' lines stand. */
  bodyFace?: string | null;
  /** The preserved sheet object (`type: "smartart"`). */
  object: Record<string, unknown>;
  layoutId: string;
  groupId: string;
  images?: Record<string, string>;
  /**
   * The workbook's theme colours and the diagram's colour variation (Office's
   * id; omitted, the one its data part names): the converted diagram is
   * painted the way Excel paints it, on the workbook's own theme (X-043).
   */
  colorsId?: string | null;
  colorTheme?: Partial<SmartArtColorTheme> | null;
  /** The conversion on open: a point the file leaves empty stays empty
   * (Excel draws no words there), rather than the seeded labels Insert and
   * an author's 转换为可编辑 write. */
  keepEmptyText?: boolean;
}

export interface SmartArtDocumentInsertRequest {
  layoutId: string;
  placeholders: SmartArtPlaceholdersWire;
  styleId?: SmartArtStyleId;
  /** The whole text pane (ids, texts, levels): the diagram is built from it rather than the gallery's sample. */
  outline?: SmartArtOutlineItem[];
  /** The diagram's frame on the page (EMU); the gallery's 6 x 3.5 in when absent. */
  frameEmu?: { width: number; height: number };
  /** A Change Colors variation (Office's id) the diagram is built in. */
  colorsId?: string | null;
  /** The theme colours `colorsId` resolves against. */
  colorTheme?: Partial<SmartArtColorTheme> | null;
  /** The prefix the built members' ids are numbered from. */
  idPrefix?: string;
}

export interface SmartArtDocumentRebuildRequest extends SmartArtLayoutDefinitionFields {
  /** Change Layout's pick: an aligned Office alias, or `null` for the catalog layout. */
  layoutUrn?: string | null;
  attrs: DocumentShapeAttrsInput;
  kind: SmartArtKind;
  outline: SmartArtOutlineItem[];
  colors?: SmartArtStyle;
  styleId?: SmartArtStyleId | null;
  keepAuthorPaint?: boolean;
  nodePaint?: (SmartArtNodePaint | null)[];
  resetCustomGeometry?: boolean;
  direction?: SmartArtDirection;
  /** Reset Shape: this item's members lose their own paint, move and resize first. */
  resetItemId?: string;
  /**
   * The host theme the diagram is drawn through (`documentSmartArtHostColorTheme`:
   * the document's colours and Design > Effects over the file's own theme
   * lists). Omitted, the file's lists its receipt carries, else Office's.
   */
  colorTheme?: Partial<SmartArtColorTheme> | null;
}

/**
 * Tidy / re-lay out one Docs SmartArt from its own outline (the Copilot's
 * `relayout_document_smartart`): the group's geometry is rebuilt by the
 * catalog layout, the texts, order, anchor and text formatting stay.
 * `layout` (a catalog id) switches layout, `items` replaces the item texts
 * in order (extra items are added, missing ones dropped), `styleId` /
 * `colorsId` restyle it. `resetGeometry` (default true) puts every moved or
 * resized shape back on the layout.
 */
export interface SmartArtDocumentRelayoutRequest extends SmartArtLayoutDefinitionFields {
  attrs: DocumentShapeAttrsInput;
  layout?: string | null;
  items?: string[] | null;
  styleId?: SmartArtStyleId | null;
  colorsId?: string | null;
  resetGeometry?: boolean;
  /** Every item's text back to the layout's own format (文字统一). */
  uniformText?: boolean;
  /**
   * The host theme the diagram is drawn through (`documentSmartArtHostColorTheme`:
   * the document's colours and Design > Effects over the file's own theme
   * lists). Omitted, the file's lists its receipt carries, else Office's.
   */
  colorTheme?: Partial<SmartArtColorTheme> | null;
}

export type SmartArtDocumentRelayoutRefusal = "not_smartart" | "rebuild_refused";

export interface SmartArtDocumentPaintRequest extends SmartArtLayoutDefinitionFields {
  attrs: DocumentShapeAttrsInput;
  colors?: SmartArtStyle | null;
  styleId?: SmartArtStyleId | null;
  authorPaint?: boolean;
  nodePaint?: SmartArtNodePaint[] | null;
  /** Change Colors' variation (Office's id); `null` back to Office's default. */
  colorsId?: string | null;
  /**
   * The host theme the diagram is drawn through (`documentSmartArtHostColorTheme`:
   * the document's colours and Design > Effects over the file's own theme
   * lists). Omitted, the file's lists its receipt carries, else Office's.
   */
  colorTheme?: Partial<SmartArtColorTheme> | null;
}

export interface SmartArtDocumentConvertRequest extends SmartArtLayoutDefinitionFields {
  dataXml: string;
  layoutIdentity?: unknown;
  layoutId?: string | null;
  frameEmu: { widthEmu: number; heightEmu: number };
  smartArtStyle?: unknown;
  groupIdsInUse?: string[];
  images?: Record<string, string>;
  /** WordArt each data point's words wear, by the point's `modelId`. */
  itemTextPaint?: Record<string, unknown>;
  /** The document's theme colours, which the diagram's colour variation is painted on. */
  colorTheme?: unknown;
}

/** A preserved diagram's data part (`dgm:dataModel` XML), as stored on its view. */
export interface SmartArtNativeOutlineRequest {
  dataXml: string;
}

export interface SmartArtPreviewRequest {
  previews: {
    kind: SmartArtKind;
    items: string[];
    levels?: number[];
    /** Which items are assistants (an organization chart's sample). */
    assistants?: boolean[];
    bounds: SmartArtBounds;
    style?: SmartArtStyle;
    styleId?: SmartArtStyleId;
    /** An aligned Office alias of `kind` to draw instead (`smartArtLayoutUrn`). */
    layoutUrn?: string;
  }[];
}

type Method<I, O> = { input: I; output: O };

export interface SmartArtEngineMethods {
  /** Insert / Convert to SmartArt on a slide (and the Copilot's). */
  "presentation.build": Method<SmartArtBuildRequest, { nodes: OfficeSmartArtNode[] }>;
  /** Change layout, Right to Left, any outline change, Reset Shape / Graphic. */
  "presentation.rebuild": Method<SmartArtPresentationRebuildRequest, { members: OfficeSmartArtHostNode[] | null }>;
  /** Change Colors / the Styles gallery. */
  "presentation.restyle": Method<SmartArtRestyleRequest, { members: OfficeSmartArtHostNode[] }>;
  /** A native Office diagram as MRAI's own (open conversion, 转换为可编辑). */
  "presentation.convert": Method<SmartArtNativeConvertRequest, { nodes: OfficeSmartArtNode[] | null }>;
  "sheet.build": Method<SmartArtBuildRequest, { members: SpreadsheetSmartArtMember[] }>;
  "sheet.rebuild": Method<SmartArtSheetRebuildRequest, { rebuild: SheetSmartArtRebuild | null }>;
  "sheet.convert": Method<SmartArtSheetConvertRequest, { members: Record<string, unknown>[] | null }>;
  "document.insert": Method<SmartArtDocumentInsertRequest, { projection: DocumentSmartArtProjection }>;
  "document.rebuild": Method<SmartArtDocumentRebuildRequest, { patch: DocumentSmartArtPatch | null }>;
  "document.relayout": Method<
    SmartArtDocumentRelayoutRequest,
    {
      patch: DocumentSmartArtPatch | null;
      refusal?: SmartArtDocumentRelayoutRefusal;
      layoutId?: string | null;
      items?: number;
    }
  >;
  "document.paint": Method<SmartArtDocumentPaintRequest, { patch: DocumentSmartArtPatch | null }>;
  "document.reset": Method<
    { attrs: DocumentShapeAttrsInput; colorTheme?: Partial<SmartArtColorTheme> | null },
    { patch: DocumentSmartArtPatch | null }
  >;
  "document.convert": Method<
    SmartArtDocumentConvertRequest,
    { converted: { attrs: DocumentShapeAttrs; kind: string; layoutId: string } | null }
  >;
  /**
   * A preserved diagram's outline, read from its data part (null: a data
   * model the reader cannot prove). The 转换为可编辑 gallery's capacity and
   * picture checks and the conversion's picture fetch read it.
   */
  "native.outline": Method<SmartArtNativeOutlineRequest, { outline: NativeSmartArtOutlineItem[] | null }>;
  /** Gallery thumbnails: the layouts, styles and colours a gallery shows. */
  previews: Method<SmartArtPreviewRequest, { previews: OfficeSmartArtNode[][] }>;
}
