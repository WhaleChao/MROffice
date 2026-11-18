/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeOfficeTextWarp } from "./office-text-warp";
import {
  CODE_CONTENT_TYPE,
  DOCUMENT_CONTENT_TYPE,
  PDF_CONTENT_TYPE,
  PRESENTATION_CONTENT_TYPE,
  SPREADSHEET_CONTENT_TYPE,
} from "./native-page-content-types";
import { CHART_COLOR_PALETTE_IDS } from "./office-charts/chart-color-palettes.generated";
import { isPictureNode } from "../components/console/editor/native-document/documentObjectNodes";
import { smartArtColorVariant } from "./office-smartart/colors";
import { isAliasNativeSmartArtLayout } from "./office-smartart/native-layouts";
import { PRESENTATION_STORED_BLACK_WHITE_MODES } from "./presentation/presentation-black-white-mode";
import {
  OFFICE_MASTER_BODY_LEVELS,
  OFFICE_MASTER_MAX_ADDED_PLACEHOLDERS,
  OFFICE_PLACEHOLDER_KINDS,
  OFFICE_ADDITIONAL_MASTER_KEY,
  OFFICE_CUSTOM_LAYOUT_ID,
  OFFICE_MASTER_MAX_ADDITIONAL,
  OFFICE_MASTER_MAX_CUSTOM_LAYOUTS,
  OFFICE_MASTER_PLACEHOLDER_FAMILIES,
} from "./presentation-office-master-model";
import {
  normalizeOfficeTrendline,
  type OfficeTrendline,
} from "./office-charts/trendline";
import { z } from "zod";
import { pageDocumentJsonEquals } from "./page-save-acknowledgement";
import { SPREADSHEET_MAX_OBJECTS } from "./spreadsheet/object-budget";
import { MAX_CHART_POINTS_PER_SERIES, MAX_CHART_SERIES } from "./office-charts/chart-limits.generated";
import { MAX_PRESENTATION_SLIDE_BLOCKS } from "./presentation/slide-block-budget";
import { MAX_PRESENTATION_SLIDES } from "./presentation/slide-limits";
import {
  isPresentationDeckReceiptRef,
  presentationDeckReceiptTokenKey,
} from "./presentation-deck-receipts";
import { MAX_TEXT_COLUMNS } from "./presentation-text-columns";
import { NATIVE_CELL_ALIGNMENTS } from "./presentation-table-style";
import { PRESENTATION_UNDERLINE_STYLES } from "./presentation-text-decoration";
import { nativeRowPixelPairMatches, normalizeNativeGrid } from "./spreadsheet/native-typography";
import { normalizeOfficeFontFace, type OfficeFontFace } from "../components/office-ribbon/officeFontFace";
import { presentationWebPictureStyle } from "./presentation-picture-style-catalog";
import { normalizeOfficeTextPaint } from "./office-text-paint";
import { normalizePresentationAnimatedBackground } from "./presentation-animated-background";
import { normalizeOfficePictureEffects } from "./office-picture-effects";
import { normalizeOfficeShadowEffect } from "./office-shadow-effects";
import { officePictureStyle } from "./office-picture-styles/catalog";
import { normalizeSpreadsheetChartFormat, type SpreadsheetChartFormat } from "./spreadsheet/chart-format";
import type { SpreadsheetChartBinding } from "./spreadsheet/chart";
import { normalizeChartNativeLook, type ChartNativeLook } from "./office-chart-native-look";
import { newSheetPrintSettings } from "./spreadsheet/print-settings";
import { parseDataQueryState, type DataQueryState } from "./spreadsheet/data-queries";
import { readPresentationChartV2 } from "./presentation-chart-v2";
import { readPresentationChartWorkbook } from "./presentation-chart-desktop";
import type { ChartBlockAttrsV2 } from "../components/console/editor/extensions/chartBlockShared";
import { isValidInlineImageSource } from "./inline-image-source";
import { isOfficeProvenanceRecord } from "./office-object-provenance";
import {normalizeOfficeShapeGradient,normalizeOfficeShapeShadow,type OfficeShapeGradient,type OfficeShapeShadow} from "./office-shape-paint";
import { normalizeOfficeChartShapePaint } from "./office-chart-shape-paint";
import {isNativeScene3d,isNativeShape3d,isSmartArtScene3d,isSmartArtShape3d,isSmartArtStyleId,isSmartArtTextShape3d,type SmartArtScene3d,type SmartArtShape3d,type SmartArtStyleId,type SmartArtTextShape3d} from "./office-smartart/styles-model";
import {
  MAX_PRESENTATION_IMAGE_DIMENSION,
  MAX_PRESENTATION_IMAGE_FILE_BYTES,
  MAX_PRESENTATION_IMAGE_PIXELS,
} from "./presentation-image-file";
import {
  PICTURE_CORRECTIONS_AMBIGUOUS_LEGACY_FLAG,
  PICTURE_CORRECTIONS_STORED_UNIT,
} from "./presentation-image-effects";
import {
  normalizeSignedCropRect,
  parsePresentationNativeGroupAuthority,
} from "./presentation-crop-group";
import { isPersistableDocumentMediaAttrs } from "./document-media";
import {
  DEFAULT_SPREADSHEET_COLUMN_COUNT,
  DEFAULT_SPREADSHEET_ROW_COUNT,
  NEW_SHEET_DEFAULT_COLUMN_WIDTH_PX,
  MAX_COLUMN_WIDTH_PX,
  MAX_GRID_COLUMN_SIZE_OVERRIDES,
  MAX_GRID_COLUMNS,
  MAX_GRID_ROW_SIZE_OVERRIDES,
  MAX_GRID_ROWS,
  MAX_ROW_HEIGHT_PX,
  MIN_COLUMN_WIDTH_PX,
  MIN_ROW_HEIGHT_PX,
} from "./spreadsheet/grid-sizing";
import {
  ICON_SET_FAMILIES,
  iconSetCount,
  MAX_RANGE_COL,
  MAX_RANGE_ROW,
  normalizeHexColor,
  sanitizeHyperlink,
} from "./spreadsheet/conditional-format";
import {
  DATA_TABLE_ICON_FILTER_SETS,
  DATA_TABLE_RELATIVE_DATE_PERIODS,
  normalizeDataTableColumnFilter,
  normalizeDataTableSortState,
  type DataTableColumnFilter,
  type DataTableSortState,
} from "./spreadsheet/data-table";
import {
  normalizeFilterDateGroups,
  type FilterDateGroup,
} from "./spreadsheet/date-filter-tree";
import { parseInternalLink, parseNamedLink } from "./spreadsheet/sheet-interactions";
import {
  blankCellFormatsArea,
  blankCellFormatsIssue,
  MAX_BLANK_CELL_FORMAT_GROUPS,
  MAX_BLANK_CELL_FORMAT_RANGES,
  MAX_BLANK_CELL_FORMAT_REF_LENGTH,
  MAX_BLANK_CELL_FORMAT_WORKBOOK_AREA,
} from "./spreadsheet/blank-cell-formats";
import { MAX_SPREADSHEET_INK_POINTS_CHARS } from "./spreadsheet/shapes";
import { SPARKLINE_NATIVE_GRAPH_BLOCKER } from "./spreadsheet/sparkline";
import spreadsheetPageEnvelope from "./spreadsheet/spreadsheet-page-envelope.shared.json";
import { MAX_SPREADSHEET_COMMENT_TEXT_LENGTH } from "./spreadsheet/comments";
import {
  normalizePageMargins,
  normalizePageSize,
} from "../components/console/editor/native-document/pageSetup";
import { normalizeColumns } from "../components/console/editor/native-document/documentLayout";
import { normalizeDocumentWatermark } from "../components/console/editor/native-document/watermark";
import {
  DOCUMENT_VIRTUAL_TABLE_NATIVE_XML_MAX_BYTES,
  estimateDocumentWithVirtualTablesNativeXmlBytes,
  DOCUMENT_VIRTUAL_TABLE_FORBIDDEN_ANCESTORS,
  isPersistableDocumentVirtualTableAttrs,
} from "../components/console/editor/native-document/documentVirtualTableModel";
import { normalizeSectionStoryGraph } from "../components/console/editor/native-document/sectionStoryGraph";
import { normalizeSectionLayouts } from "../components/console/editor/native-document/sectionLayouts";
import {
  FOOTNOTE_MODES,
  MAX_FOOTNOTE_BODY_LENGTH,
  MAX_FOOTNOTE_ID_LENGTH,
  MAX_FOOTNOTES,
} from "../components/console/editor/native-document/footnotes";
import {
  CITATION_STYLES,
  MAX_AUTHORS_PER_SOURCE,
  MAX_CSL_FIELD_LENGTH,
  MAX_SOURCE_ID_LENGTH,
  MAX_SOURCES,
} from "../components/console/editor/native-document/citations";
import {
  AUTHORABLE_SLIDE_TRANSITIONS,
  MAX_TRANSITION_ADVANCE_AFTER_MS,
  MAX_TRANSITION_DURATION_MS,
  MIN_TRANSITION_ADVANCE_AFTER_MS,
  MIN_TRANSITION_DURATION_MS,
  PRESENTATION_SLIDE_TRANSITION_SCHEMA_VERSION,
  SLIDE_TRANSITION_EFFECT_OPTIONS,
  SLIDE_TRANSITIONS,
  isDirectionalSlideTransition,
  isSlideTransitionDirectionForKind,
  isSlideTransitionP15LeftRightKind,
  normalizePresentationTransitionSound,
  SLIDE_TRANSITION_ZOOM_WARP_ELEMENT,
  type PresentationTransitionSound,
} from "./presentation-transitions";
import {
  LAYOUT_IDS,
  parseNativeLayoutCascadeReceipt,
} from "./presentation-layouts";
import { parseNativeSlideReuseReceipt } from "./presentation-reuse-slides";
import { OFFICE_RECOLOR_PRESET_IDS } from "./office-recolor-presets";
import { OFFICE_VIDEO_STYLE_IDS } from "./office-video-styles/catalog";
import { isVideoPosterSource } from "./presentation-video-poster-source";
import {
  nativeSlideSizeOriginalGeometry,
  parseNativeSlideSizeAuthority,
} from "./presentation-slide-size";
import {
  parseNativePresentationAudioAuthority,
  parseNativePresentationVideoAuthority,
} from "./presentation-video-file";
import {
  MAX_STROKE_WIDTH,
  PRESENTATION_CONNECTOR_END_SIZES,
  PRESENTATION_CONNECTOR_ENDS,
  PRESENTATION_CONNECTOR_PRESETS,
  PRESENTATION_CONNECTOR_ROUTES,
  SHAPE_KINDS,
  STROKE_STYLES,
  isPresentationConnectorShapeKind,
  normalizePresentationConnectorSpec,
  normalizePresentationOpenShapeLineEnds,
  presentationShapeTakesLineEnds,
  normalizeShapeKind,
  normalizeStrokeWidth,
  presentationInkStrokeWidthPx,
} from "./presentation-shapes";
import {
  PRESENTATION_SHAPE_POINTS_MAX_LENGTH,
  normalizePresentationCustomPath,
  presentationCustomGeometryPathAdmitted,
  presentationCustomPathIsOpenLineOnly,
  presentationInkPointsWithinEnvelope,
  type PresentationCustomPath,
} from "./presentation-freeform";
import {
  INK_EFFECT_MAX_CHARS,
  isInkEffect,
  INK_CUT_ENDS,
  MAX_INK_PRESSURE_CHARS,
  normalizeDocumentShapeAdjustments,
} from "./document-shapes";
import {
  normalizeOfficeTextLayout,
  type OfficeTextLayout,
} from "./office-text-layout";
import {
  normalizeOfficePictureFillOptions,
  type OfficePictureFillOptions,
} from "./office-picture-fill-options";
import {
  MAX_PRESENTATION_FIRST_SLIDE_NUMBER,
  MAX_PRESENTATION_DATE_LENGTH,
  MAX_PRESENTATION_FOOTER_LENGTH,
  MAX_PRESENTATION_NOTES_SIZE_EMU,
  MIN_PRESENTATION_FIRST_SLIDE_NUMBER,
} from "./presentation-header-footer";
import {
  MAX_PRESENTATION_MEDIA_BOOKMARKS,
  MAX_PRESENTATION_MEDIA_BOOKMARK_NAME_CHARS,
  MAX_PRESENTATION_MEDIA_FADE_MS,
  MAX_PRESENTATION_MEDIA_TIME_MS,
  MAX_PRESENTATION_INLINE_AUDIO_DATA_URL_CHARS,
  EMBED_KINDS,
  PRESENTATION_MEDIA_PLAYBACK_STARTS,
  PRESENTATION_MEDIA_VOLUMES,
  PRESENTATION_AUDIO_ASSET_MIME_TYPES,
  assetBackedPresentationMediaSource,
  isPersistablePresentationAudioDataUrl,
  isPersistableEmbedUrlOrEmpty,
} from "./presentation-embed";
import {
  MAX_PRESENTATION_MEDIA_CAPTION_LABEL_CHARS,
  MAX_PRESENTATION_MEDIA_CAPTION_TRACKS,
  normalizePresentationMediaCaptionTrack,
  normalizePresentationMediaCaptionTracks,
  parseNativePresentationCaptionAuthority,
} from "./presentation-media-captions";
import {
  BLOCK_ANIMATIONS,
  PRESENTATION_MEDIA_CLICK_ACTION_KINDS,
  PRESENTATION_RELATIVE_CLICK_ACTION_TARGETS,
  isPresentationActionButtonShapeKind,
  presentationHasAbsentNativeAction,
  isSafeBlockHref,
  parseSlideLinkTarget,
} from "./presentation-block-fx";
import { normalizeLinkTooltip } from "./presentation-link-settings";
import {
  MAX_TEXT_BLOCK_TEXT_LENGTH,
  BLOCK_BULLET_CHARS,
  BULLET_CHARS,
  MAX_PARAGRAPH_LEVEL,
  MAX_PARAGRAPH_TEXT_LENGTH,
  MAX_PARAGRAPHS_PER_BLOCK,
  MIN_PARAGRAPH_LEVEL,
  NUMBER_FORMATS,
  PARAGRAPH_ALIGNS,
  PARAGRAPH_BULLETS,
  PARAGRAPH_INDENT_EMU_LIMIT,
} from "./presentation-paragraphs";
import {
  MAX_EXACT_LINE_SPACING_PT,
  MAX_LINE_SPACING,
  MAX_PARAGRAPH_SPACING_PT,
  MIN_STORED_LINE_SPACING,
  MIN_PARAGRAPH_SPACING_PT,
} from "./presentation-line-spacing";
import {
  MAX_CELL_RUN_TEXT_LENGTH,
  MAX_RUN_FONT_SIZE,
  MAX_RUN_TEXT_LENGTH,
  MAX_RUNS_PER_BLOCK,
  MAX_CELL_RUNS,
  MIN_RUN_FONT_SIZE,
} from "./presentation-runs";
import {
  isEditableAnimationDurationMs,
  isEditableAnimationTimingMs,
  PRESENTATION_ANIMATION_EFFECT_OPTIONS,
  PRESENTATION_SLIDE_ANIMATION_SCHEMA_VERSION,
  presentationAnimationEffect,
  resolvePresentationAnimationEffectColor,
  resolvePresentationAnimationEffectOption,
  TIMELINE_ANIMATIONS,
  TIMELINE_TRIGGERS,
  MAX_STEP_DELAY_MS,
  MAX_STEP_DURATION_MS,
  MAX_TIMELINE_STEPS,
} from "./presentation-anim-timeline";
import {
  MOTION_PATH_PRESETS,
  MAX_MOTION_PATH_EMU,
  MAX_MOTION_PATH_WAYPOINTS,
} from "./presentation-motion-path";
import {
  isNativePresentationThemePartPath,
  isThemeDesignVariantId,
  PRESENTATION_THEME_COLOR_SLOTS,
  PRESENTATION_THEME_IDENTITY_SCHEMA_VERSION,
  PRESENTATION_THEME_VARIANT_IDS,
  THEME_FONT_FAMILIES,
} from "./presentation-theme";
import { presentationOfficeThemeVariant } from "./presentation-office-themes";
import { officeThemeFontGalleryEntry } from "./office-theme-fonts";
import { officeColorGalleryScheme, officeEffectGalleryScheme } from "./office-theme-catalog";
import {
  MAX_SHAPE_TEXT_LENGTH,
  normalizeShapeText,
} from "./presentation-text-format";
import type { PageType } from "./notebook-sdk";
import { PDF_NOTE_COLORS } from "./pdf/sticky-notes";
import { normalizePdfTextLineLayout } from "./pdf/text-line-layout";
import {
  PDF_INK_COLORS,
  PDF_INK_MAX_POINTS,
  PDF_INK_MAX_WIDTH,
} from "./pdf/ink";
import {
  ACROFORM_MAX_FIELD_NAME,
  ACROFORM_MAX_FIELDS,
  ACROFORM_MAX_OPTIONS,
  ACROFORM_MAX_TEXT,
} from "./pdf/acroform";
import {
  SIGNATURE_FONT_KEYS,
  SIGNATURE_MAX_IMAGE_BYTES,
  SIGNATURE_MAX_TEXT,
  sanitizeSignatureImage,
} from "./pdf/signatures";
import { REDACTION_MIN_SIZE } from "./pdf/redaction";
import {
  PAGE_LABEL_FIRST_VALUE_MAX,
  PAGE_LABEL_MAX_PAGES,
  PAGE_LABEL_PREFIX_MAX,
  isValidPageLabelPrefix,
} from "./pdf/page-labels";
import { PRESENTATION_GEOMETRY_MAX_EMU } from "./presentation-geometry-limit";
import { OFFICE_STROKE_DASH_FAMILY, type OfficeStrokeDash } from "./office-stroke-dash";
import {
  OFFICE_STROKE_CAPS,
  OFFICE_STROKE_COMPOUNDS,
  OFFICE_STROKE_JOINS,
  OFFICE_STROKE_MAX_MITER_LIMIT,
} from "./office-stroke-detail";
export * from "./native-spreadsheet-schemas";
import { PRESENTATION_GEOMETRY_EMU_LIMIT, PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT, imageSourceSchema, presentationBgHexSchema, presentationBgImageSchema, presentationBgPatternSchema, presentationBlockAnimationSchema, presentationBlockHrefSchema, presentationCropRectSchema, presentationGroupIdSchema, presentationLinkTooltipSchema, presentationParagraphsSchemaWithin, presentationRelativeClickActionSchema, presentationRunsSchema, presentationShapeBlockSchema, presentationTextPaintSchema, presentationTextWarpSchema, presentationThemeColorRefSchema, presentationThemeColorsSchema, presentationThemeColorsTransformedSchema, smartArtColorsIdSchema, smartArtLayoutUrnSchema, smartArtStyleIdSchema, smartArtTextShape3dSchema, spreadsheetPageContentSchema } from "./native-spreadsheet-schemas";

export {
  CODE_CONTENT_TYPE,
  DOCUMENT_CONTENT_TYPE,
  PDF_CONTENT_TYPE,
  PRESENTATION_CONTENT_TYPE,
  SPREADSHEET_CONTENT_TYPE,
};

const PDF_NOTE_COLOR_SET = new Set<string>(PDF_NOTE_COLORS);
const PDF_INK_COLOR_SET = new Set<string>(PDF_INK_COLORS);
const SIGNATURE_FONT_KEY_SET = new Set<string>(SIGNATURE_FONT_KEYS);
// Generous character cap for a base64 image data URL: 4/3 of the decoded-byte
// cap (base64 expands by ~4/3) plus the `data:image/...;base64,` prefix slack.
// The runtime sanitizer (lib/pdf/signatures.ts) enforces the real decoded-byte
// limit + the PNG/JPEG-only allow-list; this is a coarse upper bound so a
// crafted snapshot can't smuggle an unbounded string through the schema.
const SIGNATURE_MAX_IMAGE_DATA_URL_CHARS =
  Math.ceil((SIGNATURE_MAX_IMAGE_BYTES * 4) / 3) + 64;
// Drawn signature image: a base64 PNG/JPEG data URL only (never SVG — markup /
// script vector). Mirrors the runtime sanitizeSignatureImage allow-list so a
// page snapshot can't embed an `<img src>` the viewer would later reject.

// Optional RICH slide / theme background fill beyond a solid colour (PowerPoint
// / WPS "Format Background" parity): a two-stop GRADIENT, a PICTURE fill (image
// URL + sizing), or a repeating two-colour PATTERN. All additive + optional — a
// slide / theme with none of these validates byte-for-byte as before. Each is a
// generous structural guard; the render-time resolver (lib/presentation-slide-bg)
// re-normalizes / drops a malformed value, so this is not the render invariant.
const presentationBgGradientSchema = z
  .object({
    from: presentationBgHexSchema,
    to: presentationBgHexSchema,
    angle: z.number().min(0).max(360).optional(),
  })
  .passthrough();
// Document nodes are validated loosely (passthrough) so that ADDITIVE node
// types and attributes introduced over time keep existing pages valid. In
// particular: the `tableOfContents` node (which stores nothing — it derives its
// entries from the document's headings at render time), the additive
// per-heading `anchor_id` attribute (the stable Table-of-Contents link target),
// the additive per-paragraph/heading `dir` attribute (ltr/rtl text direction),
// and the additive per-table-cell `formula` attribute (a Word-style cell formula
// such as "=SUM(ABOVE)", whose computed value lives in the cell's text content)
// all validate here without a schema bump because `type` accepts any non-empty
// string and `attrs` is an optional loose record. A cell that carries no
// `formula` is byte-for-byte unchanged. Sanitization of those values happens at
// the editor's parse/render edge (see headingAnchors.ts / tableOfContents.ts /
// DocumentFormatting.ts whitelisting `dir` to ltr/rtl / tableCellStyle.ts
// normalizing `formula`), not at this persistence boundary.
export const documentNodeSchema: z.ZodType<Record<string, unknown>> = z.lazy(
  () =>
    z
      .object({
        type: z.string().min(1),
        attrs: z.record(z.unknown()).optional(),
        marks: z
          .array(
            z.object({
              type: z.string().min(1),
              attrs: z.record(z.unknown()).optional(),
            }),
          )
          .optional(),
        text: z.string().optional(),
        content: z.array(documentNodeSchema).optional(),
      })
      .passthrough(),
);

function documentNodeHasInvalidInlineImageSource(node: unknown): boolean {
  if (!node || typeof node !== "object") return false;
  const record = node as {
    attrs?: { src?: unknown };
    content?: unknown;
    type?: unknown;
  };
  if (
    typeof record.type === "string" &&
    isPictureNode(record.type) &&
    typeof record.attrs?.src === "string" &&
    !isValidInlineImageSource(record.attrs.src)
  ) {
    return true;
  }
  if (!Array.isArray(record.content)) return false;
  return record.content.some(documentNodeHasInvalidInlineImageSource);
}

const DOCUMENT_MEDIA_FORBIDDEN_ANCESTORS = new Set([
  "bulletList",
  "orderedList",
  "taskList",
  "listItem",
  "taskItem",
]);

function documentNodeHasInvalidMediaNode(
  node: unknown,
  ancestors: readonly string[] = [],
): boolean {
  if (!node || typeof node !== "object") return false;
  const record = node as {
    attrs?: unknown;
    content?: unknown;
    marks?: unknown;
    text?: unknown;
    type?: unknown;
  };
  if (
    record.type === "documentMedia" &&
    (ancestors.some((type) => DOCUMENT_MEDIA_FORBIDDEN_ANCESTORS.has(type)) ||
      !isPersistableDocumentMediaAttrs(record.attrs) ||
      record.text !== undefined ||
      record.content !== undefined ||
      (record.marks !== undefined &&
        (!Array.isArray(record.marks) || record.marks.length > 0)))
  ) {
    return true;
  }
  if (!Array.isArray(record.content)) return false;
  const nextAncestors =
    typeof record.type === "string" ? [...ancestors, record.type] : ancestors;
  return record.content.some((child) =>
    documentNodeHasInvalidMediaNode(child, nextAncestors),
  );
}

function documentNodeHasInvalidVirtualTable(
  node: unknown,
  ancestors: readonly string[] = [],
  tableIds: Set<string> = new Set(),
): boolean {
  if (!node || typeof node !== "object") return false;
  const record = node as {
    attrs?: unknown;
    content?: unknown;
    marks?: unknown;
    text?: unknown;
    type?: unknown;
  };
  if (record.type === "documentVirtualTable") {
    if (
      ancestors.some((type) =>
        DOCUMENT_VIRTUAL_TABLE_FORBIDDEN_ANCESTORS.has(type),
      ) ||
      !isPersistableDocumentVirtualTableAttrs(record.attrs) ||
      record.text !== undefined ||
      record.content !== undefined ||
      (record.marks !== undefined &&
        (!Array.isArray(record.marks) || record.marks.length > 0))
    ) {
      return true;
    }
    const tableId = record.attrs.tableId;
    if (tableIds.has(tableId)) return true;
    tableIds.add(tableId);
  }
  if (!Array.isArray(record.content)) return false;
  const nextAncestors =
    typeof record.type === "string" ? [...ancestors, record.type] : ancestors;
  return record.content.some((child) =>
    documentNodeHasInvalidVirtualTable(child, nextAncestors, tableIds),
  );
}

// Page setup (size + margins), the optional multi-column count, plus the
// optional document-level header/footer text are persisted on the top-level doc
// node's `attrs`. The `attrs` record stays OPTIONAL and otherwise loose
// (passthrough) so documents written before these features — which carry no
// `pageSize` / `pageMargins` / `columns` / `headerText` / `footerText` at all —
// still validate. But when those keys ARE present we validate them here (in
// addition to the rendering-edge normalizer/sanitizer), so a forged/out-of-
// whitelist value (e.g. a CSS-injection attempt smuggled in as `pageSize` or
// `columns`, or an over-long header) is rejected/bounded at parse time rather
// than trusted. Each key is independently `.optional()`, so a legacy doc that
// omits any of them is fine. Header/footer are plain strings clamped to a sane
// upper bound; markup/control-byte stripping happens at the editor edge
// (headerFooter.ts), so here we only bound the length additively.
const documentSourceListSchema = z
      .array(
        z
          .object({
            id: z.string().min(1).max(MAX_SOURCE_ID_LENGTH),
            type: z.string().max(MAX_CSL_FIELD_LENGTH).optional(),
            title: z.string().max(MAX_CSL_FIELD_LENGTH).optional(),
            author: z
              .array(
                z
                  .object({
                    family: z.string().max(MAX_CSL_FIELD_LENGTH).optional(),
                    given: z.string().max(MAX_CSL_FIELD_LENGTH).optional(),
                  })
                  .passthrough(),
              )
              .max(MAX_AUTHORS_PER_SOURCE)
              .optional(),
            year: z
              .union([z.string().max(MAX_CSL_FIELD_LENGTH), z.number()])
              .optional(),
            containerTitle: z.string().max(MAX_CSL_FIELD_LENGTH).optional(),
            publisher: z.string().max(MAX_CSL_FIELD_LENGTH).optional(),
            volume: z.string().max(MAX_CSL_FIELD_LENGTH).optional(),
            issue: z.string().max(MAX_CSL_FIELD_LENGTH).optional(),
            page: z.string().max(MAX_CSL_FIELD_LENGTH).optional(),
            url: z.string().max(MAX_CSL_FIELD_LENGTH).optional(),
          })
          .passthrough(),
      )
      .max(MAX_SOURCES);

const documentDocAttrsSchema = z
  .object({
    pageSize: z
      .unknown()
      .refine((value) => normalizePageSize(value) !== null, {
        message: "pageSize must be a whitelisted page size",
      })
      .optional(),
    pageMargins: z
      .unknown()
      .refine((value) => normalizePageMargins(value) !== null, {
        message: "pageMargins must be a whitelisted margin preset",
      })
      .optional(),
    columns: z
      .unknown()
      .refine((value) => normalizeColumns(value) !== null, {
        message: "columns must be 1, 2, or 3",
      })
      .optional(),
    headerText: z.string().max(5_000).optional(),
    footerText: z.string().max(5_000).optional(),
    // Word's document text watermark (Design > Watermark > Custom Watermark).
    // ADDITIVE/OPTIONAL and NULLABLE: documents written before the feature
    // carry no `watermark` at all, and clearing one persists `null`. When
    // present the bag must be one the shared model accepts, so a forged
    // colour, an unbounded stamp string, an off-whitelist font family or a
    // size outside Word's own ceiling is rejected at parse time rather than
    // reaching a painter or a VML writer. `normalizeWatermark` is the SAME
    // authority the editor, the offline DOCX writer and (restated, pinned
    // equal) the API's closed export validator apply.
    watermark: z
      .unknown()
      .refine(
        (value) => value == null || normalizeDocumentWatermark(value) !== null,
        { message: "watermark must be a valid text or picture watermark" },
      )
      .optional(),
    sectionStoryGraph: z.unknown().nullable().optional(),
    // Comment threads (Word/WPS "comments"). ADDITIVE/OPTIONAL: documents written
    // before the feature carry no `commentThreads` at all and still validate
    // (the key is `.optional()`). When present, each thread is bounded here
    // (id/comments/resolved) AND re-sanitized + capped at the editor edge by
    // normalizeCommentThreads (see native-document/comments.ts), so a forged
    // store can never smuggle markup or an unbounded blob past load. The bounds
    // mirror the runtime caps (MAX_COMMENT_THREADS / MAX_COMMENTS_PER_THREAD /
    // MAX_COMMENT_TEXT_LENGTH / MAX_COMMENT_AUTHOR_LENGTH). The anchoring
    // `comment` inline mark rides in the document body and validates through the
    // open `marks` passthrough (documentNodeSchema), so no node-schema bump is
    // needed.
    commentThreads: z
      .array(
        z
          .object({
            id: z.string().min(1).max(64),
            resolved: z.boolean().optional(),
            comments: z
              .array(
                z
                  .object({
                    text: z.string().max(4_000),
                    author: z.string().max(80).nullable().optional(),
                    timestamp: z.number().nullable().optional(),
                  })
                  .passthrough(),
              )
              .max(200),
          })
          .passthrough(),
      )
      .max(500)
      .optional(),
    // Footnotes / endnotes (Word/WPS parity). ADDITIVE/OPTIONAL: documents written
    // before the feature carry no `footnotes` / `footnoteMode` and still validate
    // (both keys are `.optional()`). When present, each note is bounded here
    // (id/body) AND re-sanitized + capped at the editor edge by normalizeFootnotes
    // (see native-document/footnotes.ts), so a forged store can never smuggle
    // markup or an unbounded blob past load. The bounds mirror the runtime caps
    // (MAX_FOOTNOTES / MAX_FOOTNOTE_ID_LENGTH / MAX_FOOTNOTE_BODY_LENGTH). The
    // anchoring `footnoteRef` inline atom node + the `footnoteList` block node both
    // ride in the document body and validate through the open `content`/`type`
    // passthrough (documentNodeSchema), so no node-schema bump is needed.
    footnotes: z
      .array(
        z
          .object({
            id: z.string().min(1).max(MAX_FOOTNOTE_ID_LENGTH),
            body: z.string().max(MAX_FOOTNOTE_BODY_LENGTH).optional(),
          })
          .passthrough(),
      )
      .max(MAX_FOOTNOTES)
      .optional(),
    footnoteMode: z.enum(FOOTNOTE_MODES).optional(),
    // Citations / bibliography (Word/WPS parity). ADDITIVE/OPTIONAL: documents
    // written before the feature carry no `sources` / `citationStyle` and still
    // validate. Sources are a CSL-JSON-like subset; each field is bounded here AND
    // re-sanitized + capped at the editor edge by normalizeSources (see
    // native-document/citations.ts). The bounds mirror the runtime caps (MAX_SOURCES
    // / MAX_SOURCE_ID_LENGTH / MAX_AUTHORS_PER_SOURCE / MAX_CSL_FIELD_LENGTH). The
    // anchoring `citation` inline atom node + the `bibliography` block node both
    // ride in the document body and validate through the open `content`/`type`
    // passthrough, so no node-schema bump is needed. Every object is `.passthrough()`
    // so additional CSL fields a future style needs round-trip untouched.
    // The editor keeps the library either as the legacy bare array or, since
    // the citation picker's source store, as the versioned envelope
    // `{ version: 1, items }` (CitationSourceStoreEnvelope in
    // native-document/citations.ts). Both are the same CSL subset; a
    // validator that knew only the array made Compare Documents (and the
    // HTML round trip) refuse every document that had cited a source.
    sources: z
      .union([
        documentSourceListSchema,
        z
          .object({ version: z.literal(1), items: documentSourceListSchema })
          .passthrough(),
      ])
      .optional(),
    citationStyle: z.enum(CITATION_STYLES).optional(),
  })
  .passthrough();

export const documentPageContentSchema = z
  .object({
    type: z.literal(DOCUMENT_CONTENT_TYPE),
    attrs: documentDocAttrsSchema.optional(),
    content: z.array(documentNodeSchema),
  })
  .passthrough()
  .superRefine((content, ctx) => {
    if (
      content.attrs?.sectionStoryGraph !== undefined &&
      content.attrs.sectionStoryGraph !== null
    ) {
      const sectionLayouts = normalizeSectionLayouts(
        content.attrs.sectionLayouts,
      );
      if (
        sectionLayouts === null ||
        normalizeSectionStoryGraph(
          content.attrs.sectionStoryGraph,
          sectionLayouts?.length,
        ) === null
      ) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["attrs", "sectionStoryGraph"],
          message: "Invalid section story graph",
        });
      }
    }
    if (content.content.some(documentNodeHasInvalidInlineImageSource)) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["content"],
        message: "Invalid inline image data URL",
      });
    }
    if (content.content.some((node) => documentNodeHasInvalidMediaNode(node))) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["content"],
        message: "Invalid document media attributes",
      });
    }
    const virtualTableIds = new Set<string>();
    if (
      content.content.some((node) =>
        documentNodeHasInvalidVirtualTable(node, [], virtualTableIds),
      )
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["content"],
        message: "Invalid sparse document table attributes",
      });
    }
    const sparseNativeXmlBytes =
      estimateDocumentWithVirtualTablesNativeXmlBytes(content);
    if (
      sparseNativeXmlBytes === null ||
      sparseNativeXmlBytes >= DOCUMENT_VIRTUAL_TABLE_NATIVE_XML_MAX_BYTES
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["content"],
        message: "Sparse document table cannot be serialized safely",
      });
    }
  });

// Allowed layout placeholder keys for a text block. Constrained to the same
// whitelist the editor resolves prompts from (LAYOUT_PLACEHOLDER_KEYS in
// PresentationWindow) so an imported/hand-edited deck can never inject an
// arbitrary placeholder value that would later index into an i18n key.
// (Comparison's headings and the caption layouts' caption and picture boxes
// were missing here, so a slide on those three layouts failed the page check.)
const presentationPlaceholderSchema = z.enum([
  "title",
  "subtitle",
  "content",
  "heading",
  "caption",
  "picture",
  "sectionTitle",
  "sectionText",
]);

function presentationPlaceholderSourceIsValid(
  source: Record<string, unknown>,
): boolean {
  const keys = [
    "placeholderType",
    "placeholderIndex",
    "placeholderNativeIndexPresent",
    "layoutPart",
    "masterPart",
  ];
  if (!keys.some((key) => Object.hasOwn(source, key))) return true;
  const index = source.placeholderIndex;
  const nativeIndexPresent = source.placeholderNativeIndexPresent;
  const layoutGeometry = source.layoutGeometry;
  const layoutStyle = source.layoutStyle;
  const geometryValid =
    layoutGeometry === undefined ||
    (layoutGeometry !== null &&
      typeof layoutGeometry === "object" &&
      !Array.isArray(layoutGeometry) &&
      Object.entries(layoutGeometry).every(
        ([key, value]) =>
          
          (["flipH", "flipV"].includes(key) && typeof value === "boolean") ||
          (["left", "top", "width", "height", "rotation"].includes(key) &&
            typeof value === "number" &&
            Number.isFinite(value)),
      ));
  const styleValid =
    layoutStyle !== null &&
    typeof layoutStyle === "object" &&
    !Array.isArray(layoutStyle) &&
    JSON.stringify(layoutStyle).length <= 4_096;
  return (
    typeof source.shapeId === "string" &&
    /^shp[1-9][0-9]*$/.test(source.shapeId) &&
    typeof source.placeholderType === "string" &&
    source.placeholderType.length > 0 &&
    source.placeholderType.length <= 32 &&
    typeof index === "string" &&
    /^(?:0|[1-9][0-9]{0,9})$/.test(index) &&
    Number(index) <= 4_294_967_295 &&
    typeof nativeIndexPresent === "boolean" &&
    (!nativeIndexPresent ? index === "0" : true) &&
    typeof source.layoutPart === "string" &&
    source.layoutPart.length > 0 &&
    source.layoutPart.length <= 512 &&
    typeof source.masterPart === "string" &&
    source.masterPart.length > 0 &&
    source.masterPart.length <= 512 &&
    geometryValid &&
    styleValid
  );
}

const presentationParagraphsSchema = presentationParagraphsSchemaWithin(
  MAX_PARAGRAPH_TEXT_LENGTH,
);

export const presentationTextBlockSchema = z
  .object({
    // Insert > WordArt / WordArt Styles on the box's words (PPT-A059).
    textPaint: presentationTextPaintSchema.optional(),
    textWarp: presentationTextWarpSchema.optional(),
    // SmartArt label 3-D; see `presentationShapeBlockSchema`.
    textShape3d:smartArtTextShape3dSchema.optional(),
    // The camera of the group the box sits in (`_pptx_group_scene_3d` reads
    // any `p:grpSp`'s, as PowerPoint draws it): held to the display-only
    // guard, as a shape block's is. The SmartArt Styles' six cameras refused
    
    // a group turned perspectiveContrastingLeftFacing).
    scene3d:z.custom<SmartArtScene3d>(isNativeScene3d).optional(),
    smartArtStyle:smartArtStyleIdSchema.optional(),
    // SmartArt Right to Left, on every member of a mirrored diagram.
    smartArtDirection:z.literal("rtl").optional(),
    // SmartArt Change Colors variation (Office's id), on every member.
    smartArtColors:smartArtColorsIdSchema.optional(),
    // The Office alias the diagram is drawn as (hProcess9), on every member.
    smartArtLayoutUrn:smartArtLayoutUrnSchema.optional(),
    id: z.string().min(1),
    type: z.literal("text"),
    text: z.string(),
    // Optional list structure; see `presentationParagraphsSchema`. Additive —
    // flat-text blocks omit it and still validate byte-for-byte.
    paragraphs: presentationParagraphsSchema.optional(),
    // Optional BLOCK-LEVEL bullet glyph / number format choice (PowerPoint /
    // WPS "Bullets & Numbering" gallery). Both additive + optional — a block
    // that never picked a glyph/format omits them and renders with the per-level
    // default scheme (byte-for-byte unchanged). Constrained to the shared
    // allowed sets so a hand-edited deck cannot inject an arbitrary glyph; the
    // render edge (`renderParagraphs`) re-normalizes via `normalizeBulletChar` /
    // `normalizeNumberFormat`, so this is a structural guard only.
    bulletChar: z.enum(BLOCK_BULLET_CHARS).optional(),
    numberFormat: z.enum(NUMBER_FORMATS).optional(),
    // Optional character-level (per-run) inline formatting for the flat block
    // text (when there are no `paragraphs`); see `presentationRunsSchema`.
    // Additive — blocks without per-run styling omit it and still validate.
    runs: presentationRunsSchema.optional(),
    // Exact DrawingML text-frame insets. The complete four-edge tuple is
    // required so no save path invents or silently borrows a missing edge.
    textInsets: z
      .object({
        left: z.number().int().min(0).max(PRESENTATION_GEOMETRY_EMU_LIMIT),
        right: z.number().int().min(0).max(PRESENTATION_GEOMETRY_EMU_LIMIT),
        top: z.number().int().min(0).max(PRESENTATION_GEOMETRY_EMU_LIMIT),
        bottom: z.number().int().min(0).max(PRESENTATION_GEOMETRY_EMU_LIMIT),
      })
      .strict()
      .optional(),
    // Optional TEXT-BOX FILL + OUTLINE (PowerPoint "Shape Fill / Outline" on a
    // text box), all additive + optional. `fill`(+`fillTo`/`gradientAngle`) is a
    // solid or two-stop gradient background; `outlineColor`/`outlineWidth`/
    // `outlineStyle` an optional border; `shadow` a soft drop shadow. A text
    // block with none of these renders byte-for-byte as before. Bounds mirror
    // the shape fill; the render edge re-normalizes via `resolveTextBoxFill`
    // (lib/presentation-text-fill), so this is a structural guard.
    fill: z.union([presentationBgHexSchema, z.literal("none")]).optional(),
    // The solid fill's Transparency, a whole percentage (`a:alpha`), named
    // and bounded as a shape's `fillTransparency`.
    fillTransparency: z.number().int().min(1).max(100).optional(),
    fillTo: presentationBgHexSchema.optional(),
    gradientAngle: z.number().min(0).max(360).optional(),
    outlineColor: z.union([presentationBgHexSchema, z.literal("none")]).optional(),
    themeColors: presentationThemeColorsSchema.optional(),
    themeColorsTransformed: presentationThemeColorsTransformedSchema.optional(),
    outlineColorTo: presentationBgHexSchema.optional(),
    outlineGradientAngle: z.number().min(0).max(360).optional(),
    outlineWidth: z.number().min(0).max(MAX_STROKE_WIDTH).optional(),
    outlineStyle: z.enum(STROKE_STYLES).optional(),
    // `a:ln@cmpd`: the border is two parallel strokes with a gap (PPT-B368).
    outlineCompound: z.literal("double").optional(),
    // A text box IS a shape, so its border states `a:ln@algn` the same way a
    // shape's outline does — the same two words, the same "absent states
    // nothing" rule (audit 68).
    outlineAlign: z.enum(["center", "inset"]).optional(),
    shadow: z.boolean().optional(),
    // Format Shape > Effects on the box (a text box is a shape): the shape's
    // one combinable model (`presentationShapeBlockSchema.shapeEffects`), and
    // 3-D Rotation > Keep text flat (`a:flatTx`), only when checked.
    shapeEffects: z
      .record(z.string(), z.unknown())
      .refine((value) => normalizeOfficePictureEffects(value) !== null, "Unsupported shape effects")
      .optional(),
    keepTextFlat: z.literal(true).optional(),
    // Optional BODY MULTI-COLUMN count (PowerPoint "Text Options → Columns"),
    // additive + optional: a one-column block omits it and validates byte-for-
    // byte. Bounded to the editor's Columns range (`MAX_TEXT_COLUMNS`, the
    
    // the render edge re-normalizes via `normalizeColumnCount`
    // (lib/presentation-text-columns), so this is a structural guard only.
    columnCount: z.number().int().min(1).max(MAX_TEXT_COLUMNS).optional(),
    // Columns > Spacing between columns (`a:bodyPr@spcCol`, EMU; PPT-A042).
    columnSpacing: z.number().int().min(0).max(56 * 914_400).optional(),
    left: z
      .number()
      .min(-PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    top: z
      .number()
      .min(-PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    width: z
      .number()
      .positive()
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    height: z
      .number()
      .positive()
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    rotation: z.number().min(-360).max(360).optional(),
    // Optional horizontal / vertical mirror flags (PowerPoint `flipH`/`flipV`
    // parity). Purely additive booleans: a block without either omits both and
    // validates byte-for-byte. Render-only — the flip is composed into the
    // block's CSS transform (flip THEN rotate, matching PowerPoint's xfrm) and
    // never mutates the geometry (left/top/width/height) or the text. A
    // hand-edited `false` is dropped to absent on the editor write path so the
    // serialized form stays minimal.
    flipH: z.boolean().optional(),
    flipV: z.boolean().optional(),
    // Non-visual accessibility metadata shared by ordinary shapes, text boxes,
    // and the Microsoft 365 Cameo placeholder. Empty/absent descriptions stay
    // byte-clean; decorative and visible description are mutually exclusive.
    alt: z.string().max(1_000).optional(),
    decorative: z.boolean().optional(),
    name: z.string().trim().min(1).max(200).optional(),
    hidden: z.boolean().optional(),
    // A slide holds up to MAX_PRESENTATION_SLIDE_BLOCKS objects, and its
    // paint order counts them all: a bound of 200 made the editor's commands
    
    
    sourceOrder: z.number().int().min(0).max(MAX_PRESENTATION_SLIDE_BLOCKS).optional(),
    zIndex: z.number().int().min(0).max(MAX_PRESENTATION_SLIDE_BLOCKS).optional(),
    animation: presentationBlockAnimationSchema.optional(),
    href: presentationBlockHrefSchema.optional(),
    linkTooltip: presentationLinkTooltipSchema.optional(),
    // Action Settings > Hyperlink to: Next Slide ... End Show (PPT-B012).
    clickAction: presentationRelativeClickActionSchema.optional(),
    clickHighlight: z.literal(true).optional(),
    // Action Settings > Mouse Over (`a:hlinkMouseOver`, PPT-B012).
    hoverHref: presentationBlockHrefSchema.optional(),
    hoverAction: presentationRelativeClickActionSchema.optional(),
    hoverHighlight: z.literal(true).optional(),
    // Grayscale > Change Selected Object (`p:spPr@bwMode`, PPT-A109); Automatic is absent.
    blackWhiteMode: z.enum(PRESENTATION_STORED_BLACK_WHITE_MODES).optional(),
    placeholder: presentationPlaceholderSchema.optional(),
    groupId: presentationGroupIdSchema.optional(),
    // The groups nested inside `groupId` this block also belongs to,
    // outermost first (presentation-crop-group `blockGroupPath`).
    groupPath: z.array(presentationGroupIdSchema).min(1).max(7).optional(),
    source: z.record(z.unknown()).optional(),
  })
  .passthrough()
  .superRefine((block, context) => {
    if (block.outlineColorTo && (!block.outlineColor || !block.outlineWidth)) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["outlineColorTo"],
        message: "gradient outline requires a visible first colour and width",
      });
    }
    if (block.outlineGradientAngle !== undefined && !block.outlineColorTo) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["outlineGradientAngle"],
        message: "gradient outline angle requires a second colour",
      });
    }
  })
  .superRefine((block, context) => {
    if (
      block.source &&
      !presentationPlaceholderSourceIsValid(block.source)
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["source"],
        message: "native placeholder source identity is malformed",
      });
    }
  });

export const presentationImageBlockSchema = z
  .object({
    id: z.string().min(1),
    type: z.literal("image"),
    pictureStylePreset: z.string().refine(value => presentationWebPictureStyle(value) !== null, "Unsupported picture style").optional(),
    // Picture Format > Picture Effects (Shadow, Glow, Reflection, Soft
    // Edges, Bevel, 3-D Rotation): `lib/office-picture-effects`, the server's
    // `office_picture_effects` rule.
    pictureEffects: z
      .record(z.string(), z.unknown())
      .refine((value) => normalizeOfficePictureEffects(value) !== null, "Unsupported picture effects")
      .optional(),
    // The camera of the group the picture sits in (`p:grpSpPr/a:scene3d`):
    // the group turns as one object, its pictures with it.
    scene3d: z.custom<SmartArtScene3d>(isNativeScene3d).optional(),
    // Picture Format > Animate as Background (PPT-A134).
    animatedBackground: z
      .unknown()
      .refine((value) => normalizePresentationAnimatedBackground(value) !== undefined, "Unsupported animated background")
      .optional(),
    pictureOriginalSource: z.object({
      slideId: z.string().regex(/^sld[1-9][0-9]*$/),
      // Bitmap lookup only: unlike positive native writer ids, OOXML permits 0.
      shapeId: z.string().regex(/^shp(?:0|[1-9][0-9]*)$/),
      pictureRenderCache: z.object({
        schemaVersion: z.literal("pptx-picture-render-cache-v1"),
        mediaPart: z.string().min(1).max(512),
        originalSha256: z.string().regex(/^[a-f0-9]{64}$/),
        previewSha256: z.string().regex(/^[a-f0-9]{64}$/),
        previewMimeType: z.literal("image/webp"),
        // A copy fitted to the deck's share of the page (readings row 860):
        // the long side it was made at, below the standard copy's 2048.
        previewMaxSide: z.number().int().min(16).max(2047).optional(),
      }).strict(),
    }).strict().optional(),
    src: imageSourceSchema,
    // Local-file insert persists an editor-specific asset-backed union member.
    // `src` remains a bounded, authorized renderer/patchback cache; `assetId`
    // is the durable source of truth used by native from-scratch export.
    sourceKind: z.literal("asset").optional(),
    assetId: z.string().min(1).max(200).optional(),
    mimeType: z.enum(["image/png", "image/jpeg", "image/svg+xml"]).optional(),
    officeGraphicKind: z.literal("svg").optional(),
    graphicFallbackSrc: imageSourceSchema.optional(),
    // Graphics Format > Graphics Fill / Graphics Outline (PPT-A161): the SVG
    // graphic's one fill / outline (`#rrggbb` or "none"), written by the
    // server into the SVG part as PowerPoint writes it; a theme swatch keeps
    // its reference in `themeColors` (graphicFill / graphicOutline).
    graphicFill: z.union([presentationBgHexSchema, z.literal("none")]).optional(),
    graphicOutline: z.union([presentationBgHexSchema, z.literal("none")]).optional(),
    // Graphics Outline > Weight in the SVG's own user units (the part's
    // `stroke-width`; the menu converts points at the picture's size) and
    // Dashes by DrawingML name.
    graphicOutlineWidth: z.number().positive().max(10_000).optional(),
    graphicOutlineDash: z
      .enum(["solid", "sysDot", "sysDash", "dash", "dashDot", "lgDash", "lgDashDot", "lgDashDotDot"])
      .optional(),
    themeColors: z
      .object({
        graphicFill: presentationThemeColorRefSchema.optional(),
        graphicOutline: presentationThemeColorRefSchema.optional(),
        // The picture's own frame fill / outline (an imported picture's
        // Format Picture > Fill, reading `pictureOwnFill`; the API's
        // `frame_theme_keys`).
        fill: presentationThemeColorRefSchema.optional(),
        stroke: presentationThemeColorRefSchema.optional(),
        outlineColor: presentationThemeColorRefSchema.optional(),
      })
      .strict()
      .optional(),
    // Format Picture > Fill: the solid colour drawn inside the picture's
    // outline, under the picture (where it is transparent or does not reach).
    fill: z.union([presentationBgHexSchema, z.literal("none")]).optional(),
    pixelWidth: z
      .number()
      .int()
      .min(1)
      .max(MAX_PRESENTATION_IMAGE_DIMENSION)
      .optional(),
    pixelHeight: z
      .number()
      .int()
      .min(1)
      .max(MAX_PRESENTATION_IMAGE_DIMENSION)
      .optional(),
    sizeBytes: z
      .number()
      .int()
      .min(1)
      .max(MAX_PRESENTATION_IMAGE_FILE_BYTES)
      .optional(),
    hasAlpha: z.boolean().optional(),
    exifOrientation: z.number().int().min(1).max(8).optional(),
    // PowerPoint keeps the legacy title and the current Alt Text description
    // as separate cNvPr attributes. `decorative` is intentionally optional:
    // absent means the Office extension was absent, while false means the
    // extension explicitly carried val="0".
    altTitle: z.string().max(2_000).optional(),
    alt: z.string().max(2_000).optional(),
    decorative: z.boolean().optional(),
    caption: z.string().optional(),
    left: z
      .number()
      .min(-PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    top: z
      .number()
      .min(-PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    width: z
      .number()
      .positive()
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    height: z
      .number()
      .positive()
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    rotation: z.number().min(-360).max(360).optional(),
    // Optional horizontal / vertical mirror flags (PowerPoint `flipH`/`flipV`
    // parity). Purely additive booleans: a block without either omits both and
    // validates byte-for-byte. Render-only — the flip is composed into the
    // block's CSS transform (flip THEN rotate, matching PowerPoint's xfrm) and
    // never mutates the geometry (left/top/width/height) or the text. A
    // hand-edited `false` is dropped to absent on the editor write path so the
    // serialized form stays minimal.
    flipH: z.boolean().optional(),
    flipV: z.boolean().optional(),
    name: z.string().trim().min(1).max(200).optional(),
    hidden: z.boolean().optional(),
    sourceOrder: z.number().int().min(0).max(MAX_PRESENTATION_SLIDE_BLOCKS).optional(),
    zIndex: z.number().int().min(0).max(MAX_PRESENTATION_SLIDE_BLOCKS).optional(),
    animation: presentationBlockAnimationSchema.optional(),
    href: presentationBlockHrefSchema.optional(),
    linkTooltip: presentationLinkTooltipSchema.optional(),
    // Action Settings > Hyperlink to: Next Slide ... End Show (PPT-B012).
    clickAction: presentationRelativeClickActionSchema.optional(),
    clickHighlight: z.literal(true).optional(),
    // Action Settings > Mouse Over (`a:hlinkMouseOver`, PPT-B012).
    hoverHref: presentationBlockHrefSchema.optional(),
    hoverAction: presentationRelativeClickActionSchema.optional(),
    hoverHighlight: z.literal(true).optional(),
    // Grayscale > Change Selected Object (`p:spPr@bwMode`, PPT-A109); Automatic is absent.
    blackWhiteMode: z.enum(PRESENTATION_STORED_BLACK_WHITE_MODES).optional(),
    groupId: presentationGroupIdSchema.optional(),
    // The groups nested inside `groupId` this block also belongs to,
    // outermost first (presentation-crop-group `blockGroupPath`).
    groupPath: z.array(presentationGroupIdSchema).min(1).max(7).optional(),
    // Optional signed source-space keep rectangle. Additive — image blocks
    // without it render the full source and still validate.
    cropRect: presentationCropRectSchema.optional(),
    // Optional image visual effects (PowerPoint parity): a soft drop shadow, a
    // coloured frame (hex colour + 0..24px width), and 0..64px rounded corners.
    // All additive — an image without effects omits every field and still
    // validates byte-for-byte. Mirrors the shape effect bounds; the hex colour
    // is validated the same way so a stale value can't inject arbitrary CSS.
    shadow: z.boolean().optional(),
    borderColor: z
      .string()
      .regex(/^#[0-9a-fA-F]{6}$/)
      .optional(),
    borderWidth: z.number().min(0).max(24).optional(),
    borderStyle: z.enum(["none", "solid", "dashed", "dotted"]).optional(),
    // The frame's compound line and corner join (`a:ln@cmpd`, `a:miter` /
    // `a:bevel`; the server's readings row 856): absent, one line, round corners.
    borderCompound: z.enum(OFFICE_STROKE_COMPOUNDS).optional(),
    borderJoin: z.enum(["miter", "bevel"]).optional(),
    cornerRadius: z.number().min(0).max(64).optional(),
    // Crop to Shape > Oval: the bitmap is clipped to the ellipse of its frame
    // (`a:prstGeom prst="ellipse"`, PPT-B364).
    // Crop to Shape: the picture's own closed preset (`a:prstGeom@prst`,
    // default adjustments; PPT-A132), the oval since PPT-B364.
    // The preset catalog itself is checked where it is drawn and on the
    // server (a name outside it draws the plain picture), not loaded here.
    clipShape: z
      .string()
      .regex(/^[A-Za-z0-9]{1,40}$/)
      .refine((value) => value !== "rect", "A rectangle is no crop shape")
      .optional(),
    // The crop shape's dragged handles (`a:avLst` literals by guide name).
    clipShapeAdjustments: z.record(z.string().regex(/^[A-Za-z0-9]{1,16}$/), z.number().int()).optional(),
    // A picture drawn inside a freeform (`a:custGeom`), the path model a
    // freeform shape's `customPath` carries.
    clipCustomPath: z
      .custom<PresentationCustomPath>(
        (value) => normalizePresentationCustomPath(value)?.closed === true,
        "A picture clips to one closed freeform outline",
      )
      .optional(),
    // Optional picture ADJUSTMENTS (PowerPoint "Color" / "Corrections" parity),
    // all additive: `opacity` (0..100 % visible), `recolor` (a whitelisted tint
    // mode), and `brightness` / `contrast` (0..200 %, 100 = unmodified). An
    // image with none of these renders byte-for-byte as before. The render edge
    // re-clamps via `resolveImageEffects` (lib/presentation-image-effects).
    opacity: z.number().min(0).max(100).optional(),
    recolor: z
      .enum([
        "none",
        "grayscale",
        "sepia",
        "washout",
        "invert",
        "duotone",
        "vivid",
      ])
      .optional(),
    // The temporary -100..-1 admission band exists only for source-addressed
    // pre-IMG07 imports. superRefine below rejects it everywhere else; the host
    // canonicalizes admitted legacy values to stored percent before saving.
    brightness: z.number().int().min(-100).max(200).optional(),
    contrast: z.number().int().min(-100).max(200).optional(),
    // Optional COMPOSITING mode (PowerPoint "Transparency / blend" parity → CSS
    // `mix-blend-mode`). Whitelisted + additive — a picture without it omits the
    // field and validates byte-for-byte. The render edge re-clamps the same way
    // via `resolveImageEffects` so a stale value can't inject arbitrary CSS.
    blendMode: z
      .enum(["normal", "multiply", "screen", "overlay", "darken", "lighten"])
      .optional(),
  })
  .passthrough()
  .superRefine((block, ctx) => {
    const source = block.source as Record<string, unknown> | undefined;
    const correctionsUnit = source?.pictureCorrectionsUnit;
    const hasNativePictureIdentity = [source?.slideId, source?.shapeId].every(
      (value) => typeof value === "string" && value.trim().length > 0,
    );
    const correctionValues = [block.brightness, block.contrast].filter(
      (value): value is number => value !== undefined,
    );
    const hasAmbiguousLegacyCorrectionsReceipt =
      source?.pictureCorrectionsDisposition === "preserve-only" &&
      source?.[PICTURE_CORRECTIONS_AMBIGUOUS_LEGACY_FLAG] === true &&
      correctionsUnit === undefined &&
      hasNativePictureIdentity &&
      correctionValues.length > 0 &&
      correctionValues.every(
        (value) => Number.isInteger(value) && value >= 0 && value <= 100,
      );
    if (
      correctionsUnit !== undefined &&
      correctionsUnit !== PICTURE_CORRECTIONS_STORED_UNIT
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["source", "pictureCorrectionsUnit"],
        message: "Unsupported picture corrections storage unit",
      });
    }
    for (const field of ["brightness", "contrast"] as const) {
      const value = block[field];
      if (
        typeof value === "number" &&
        value < 0 &&
        !(
          correctionsUnit === undefined &&
          hasNativePictureIdentity &&
          value >= -100
        )
      ) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: [field],
          message:
            "Negative picture corrections require legacy native identity",
        });
      }
    }
    const assetFields = [
      "sourceKind",
      "assetId",
      "mimeType",
      "pixelWidth",
      "pixelHeight",
      "sizeBytes",
      "hasAlpha",
      "exifOrientation",
    ] as const;
    // Imported inline pictures may carry only intrinsic pixel dimensions so
    // Fill/Fit can reproduce PowerPoint's aspect math. Those dimensions do not
    // imply asset custody. Any actual asset marker still requires the complete
    // authenticated union.
    const assetCustodyFields = [
      "sourceKind",
      "assetId",
      "mimeType",
      "sizeBytes",
      "hasAlpha",
      "exifOrientation",
    ] as const;
    const hasAssetCustody = assetCustodyFields.some(
      (field) => block[field] !== undefined,
    );
    const hasIntrinsicDimensions =
      block.pixelWidth !== undefined || block.pixelHeight !== undefined;
    if (hasAssetCustody) {
      for (const field of assetFields) {
        if (block[field] === undefined) {
          ctx.addIssue({
            code: z.ZodIssueCode.custom,
            path: [field],
            message: "Asset-backed image metadata must be complete",
          });
        }
      }
    } else if (hasIntrinsicDimensions) {
      for (const field of ["pixelWidth", "pixelHeight"] as const) {
        if (block[field] === undefined) {
          ctx.addIssue({
            code: z.ZodIssueCode.custom,
            path: [field],
            message: "Imported image dimensions must be complete",
          });
        }
      }
    }
    if (
      block.mimeType &&
      !block.src.toLowerCase().startsWith(`data:${block.mimeType};base64,`)
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["src"],
        message: "Asset-backed image cache MIME does not match metadata",
      });
    }
    const isOfficeGraphic = block.officeGraphicKind === "svg";
    if (isOfficeGraphic) {
      // A master's or layout's own graphic, drawn on the page as its slides
      // inherit it, has no slide or shape id of its own: the part it is read
      // from (`inheritedFrom`) is its identity, and the export keeps that
      
      // editor command refuse the page).
      const inheritedNativePicture =
        (source?.inheritedFrom === "master" || source?.inheritedFrom === "layout") &&
        typeof source?.shapeName === "string";
      const importedSvgWithoutFallback =
        block.sourceKind !== "asset" &&
        (hasNativePictureIdentity || inheritedNativePicture) &&
        source?.officeGraphicDisposition === "native-svg-no-raster-fallback";
      if (!block.graphicFallbackSrc && !importedSvgWithoutFallback) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["graphicFallbackSrc"],
          message:
            "Native SVG requires a compatibility PNG preview or authenticated no-fallback source",
        });
      }
      if (block.sourceKind === "asset" && block.mimeType !== "image/svg+xml") {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["mimeType"],
          message: "Asset-backed Office Graphic must keep SVG custody",
        });
      }
      // Its crop is not in this list: the import states it, drawn as PowerPoint
      // draws it and kept the package's (reading `svgGraphicCrop`, API rule); a
      
      // Nor are its Transparency, Brightness / Contrast and Color where the
      // import states them as the package's (preserve-only): PowerPoint draws
      
      const packageFormat = {
        opacity: source?.pictureTransparencyDisposition,
        brightness: source?.pictureCorrectionsDisposition,
        contrast: source?.pictureCorrectionsDisposition,
        recolor: source?.pictureColorDisposition,
      } as const;
      for (const field of [
        "borderColor",
        "borderWidth",
        "borderStyle",
        "cornerRadius",
        "opacity",
        "recolor",
        "brightness",
        "contrast",
        "blendMode",
        "shadow",
      ] as const) {
        if (field in packageFormat && packageFormat[field as keyof typeof packageFormat] === "preserve-only") continue;
        if (block[field] !== undefined) {
          ctx.addIssue({
            code: z.ZodIssueCode.custom,
            path: [field],
            message: "Office Graphic formatting is preserve-only",
          });
        }
      }
      // An SVG graphic's theme references are its two paints, and its frame's
      
      // picture placeholder with a fill of its own).
      for (const key of ["fill", "stroke", "outlineColor"] as const) {
        if (block.themeColors?.[key] !== undefined && (block as Record<string, unknown>)[key] === undefined) {
          ctx.addIssue({
            code: z.ZodIssueCode.custom,
            path: ["themeColors"],
            message: "An Office Graphic's theme colours are its graphic fill and outline, and its frame's where it states one",
          });
        }
      }
    } else if (
      block.graphicFallbackSrc !== undefined ||
      block.mimeType === "image/svg+xml" ||
      block.graphicFill !== undefined ||
      block.graphicOutline !== undefined ||
      block.graphicOutlineWidth !== undefined ||
      block.graphicOutlineDash !== undefined ||
      block.themeColors?.graphicFill !== undefined ||
      block.themeColors?.graphicOutline !== undefined
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["officeGraphicKind"],
        message: "SVG metadata requires Office Graphic identity",
      });
    }
    if (
      typeof block.pixelWidth === "number" &&
      typeof block.pixelHeight === "number" &&
      block.pixelWidth * block.pixelHeight > MAX_PRESENTATION_IMAGE_PIXELS
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["pixelWidth"],
        message: "Asset-backed image pixel budget exceeded",
      });
    }
    // An SVG graphic's own are drawn and stay the package's (row 857, above).
    const drawnGraphicFormats = block.officeGraphicKind === "svg";
    if (
      block.opacity !== undefined &&
      (block.source as Record<string, unknown> | undefined)
        ?.pictureTransparencyDisposition === "preserve-only" &&
      !drawnGraphicFormats
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["opacity"],
        message: "Opaque native picture transparency is preserve-only",
      });
    }
    if (
      block.recolor !== undefined &&
      source?.pictureColorDisposition === "preserve-only" &&
      !drawnGraphicFormats
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["recolor"],
        message: "Opaque native picture color is preserve-only",
      });
    }
    if (
      (block.brightness !== undefined || block.contrast !== undefined) &&
      (block.source as Record<string, unknown> | undefined)
        ?.pictureCorrectionsDisposition === "preserve-only" &&
      !hasAmbiguousLegacyCorrectionsReceipt &&
      !drawnGraphicFormats
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: [block.brightness !== undefined ? "brightness" : "contrast"],
        message: "Opaque native picture corrections are preserve-only",
      });
    }
  });

export const presentationTableCellSchema = z
  .object({
    text: z.string(),
    // A merge spans as far as its table does (MAX_TABLE_SPLIT_COLUMNS /
    
    colspan: z.number().int().min(1).max(1_024).optional(),
    rowspan: z.number().int().min(1).max(4_096).optional(),
    // Optional per-cell formatting (PowerPoint parity): a hex background fill,
    // a hex text colour, a bold flag, and a whitelisted horizontal alignment.
    // All additive — a cell without formatting omits every field and validates
    // byte-for-byte. The hex colours are validated the same way as shape / image
    // colours so a stale or hand-edited value can't inject arbitrary CSS.
    // "none" is an explicit No Fill over the table style's shading.
    // An imported fill with Transparency is `#rrggbbaa` (`a:alpha`).
    fill: z
      .union([
        presentationBgHexSchema,
        z.string().regex(/^#[0-9a-fA-F]{8}$/),
        z.literal("none"),
      ])
      .optional(),
    color: z
      .string()
      .regex(/^#[0-9a-fA-F]{6}$/)
      .optional(),
    bold: z.boolean().optional(),
    align: z.enum(NATIVE_CELL_ALIGNMENTS).optional(),
    // An imported underline's kind beyond the single line (`a:rPr@u` dbl,
    // heavy, dotted, dashed, wavy ...), kept while `underline` is on.
    underlineStyle: z.enum(PRESENTATION_UNDERLINE_STYLES).optional(),
    // Font > Double Strikethrough (`strike="dblStrike"`), kept while `strike`.
    strikeStyle: z.literal("double").optional(),
    // Table Design > Shading > Picture (`a:tcPr/a:blipFill`): the slide
    // background's picture model, its URL held to the same source predicate.
    fillImage: presentationBgImageSchema.optional(),
  })
  .passthrough();

export const presentationTableBlockSchema = z
  .object({
    id: z.string().min(1),
    type: z.literal("table"),
    
    
    // (`presentation-table-merge`) and the API's MAX_PRESENTATION_TABLE_*.
    rows: z.array(z.array(presentationTableCellSchema).max(1_024)).min(1).max(4_096),
    caption: z.string().optional(),
    // Table-level formatting (PowerPoint parity), both additive booleans: thin
    // cell borders, and a header-row emphasis that bolds + fills the first row
    // at render time. A table without either omits both and validates unchanged.
    borders: z.boolean().optional(),
    headerRow: z.boolean().optional(),
    // Table Style Options (PowerPoint/WPS): banded (zebra) body rows + an
    // emphasised total row. Additive + optional so existing tables validate
    // unchanged; absent/false means off.
    bandedRows: z.boolean().optional(),
    totalRow: z.boolean().optional(),
    // Column-axis Table Style Options (PowerPoint/WPS): emphasised first /
    // last column + banded (zebra) columns. Additive + optional so existing
    // tables validate unchanged; absent/false means off.
    firstColumn: z.boolean().optional(),
    lastColumn: z.boolean().optional(),
    bandedColumns: z.boolean().optional(),
    // `a:tblPr@rtl`: columns laid out right to left. Optional; absent = LTR.
    rightToLeft: z.boolean().optional(),
    // Table Design > Shading > Table Background: the whole table's colour
    // under every cell (`a:tblPr` solid fill). Optional; absent = none.
    tableBackground: z.string().regex(/^#[0-9a-fA-F]{6}$/).optional(),
    // Table Design > Effects > Shadow / Reflection on the whole table
    // (`a:tblPr/a:effectLst`, `lib/presentation-table-effects`): EMU,
    // 1/60000 degree and 1/1000 %. Optional; absent = none.
    tableEffects: z
      .object({
        // The shared shadow model (`lib/office-shadow-effects`): outer
        // (`algn`, optional scale and skew) or inner (`inner: true`).
        shadow: z
          .object({
            inner: z.literal(true).optional(),
            blurRad: z.number().int().min(0).max(1_270_000),
            dist: z.number().int().min(0).max(2_540_000),
            dir: z.number().int().min(0).max(21_599_999),
            algn: z.enum(["tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"]).optional(),
            color: z.string().regex(/^#[0-9a-fA-F]{6}$/),
            alpha: z.number().int().min(0).max(100_000),
            sx: z.number().int().min(-200_000).max(200_000).optional(),
            sy: z.number().int().min(-200_000).max(200_000).optional(),
            kx: z.number().int().min(-5_399_999).max(5_399_999).optional(),
            ky: z.number().int().min(-5_399_999).max(5_399_999).optional(),
          })
          .strict()
          .refine((value) => normalizeOfficeShadowEffect(value) !== null, "Unsupported table shadow")
          .optional(),
        reflection: z
          .object({
            blurRad: z.number().int().min(0).max(1_270_000),
            stA: z.number().int().min(0).max(100_000),
            endA: z.number().int().min(0).max(100_000),
            endPos: z.number().int().min(0).max(100_000),
            dist: z.number().int().min(0).max(1_270_000),
          })
          .strict()
          .optional(),
      })
      .strict()
      .optional(),
    left: z
      .number()
      .min(-PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    top: z
      .number()
      .min(-PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    width: z
      .number()
      .positive()
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    height: z
      .number()
      .positive()
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    rotation: z.number().min(-360).max(360).optional(),
    // Optional horizontal / vertical mirror flags (PowerPoint `flipH`/`flipV`
    // parity). Purely additive booleans: a block without either omits both and
    // validates byte-for-byte. Render-only — the flip is composed into the
    // block's CSS transform (flip THEN rotate, matching PowerPoint's xfrm) and
    // never mutates the geometry (left/top/width/height) or the text. A
    // hand-edited `false` is dropped to absent on the editor write path so the
    // serialized form stays minimal.
    flipH: z.boolean().optional(),
    flipV: z.boolean().optional(),
    name: z.string().trim().min(1).max(200).optional(),
    hidden: z.boolean().optional(),
    sourceOrder: z.number().int().min(0).max(MAX_PRESENTATION_SLIDE_BLOCKS).optional(),
    zIndex: z.number().int().min(0).max(MAX_PRESENTATION_SLIDE_BLOCKS).optional(),
    animation: presentationBlockAnimationSchema.optional(),
    href: presentationBlockHrefSchema.optional(),
    linkTooltip: presentationLinkTooltipSchema.optional(),
    groupId: presentationGroupIdSchema.optional(),
    // The groups nested inside `groupId` this block also belongs to,
    // outermost first (presentation-crop-group `blockGroupPath`).
    groupPath: z.array(presentationGroupIdSchema).min(1).max(7).optional(),
  })
  .passthrough()
  .superRefine((block, ctx) => {
    const cellCount = block.rows.reduce((count, row) => count + row.length, 0);
    
    // with the API's page validator and deck table reader
    // (MAX_PRESENTATION_TABLE_CELLS in `presentation-table-merge`).
    if (cellCount > 100_000) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["rows"],
        message: "Presentation table blocks may contain at most 100000 cells",
      });
    }
  });

export const presentationChartDataSchema = z
  .object({
    headers: z.array(z.string().max(500)).max(20).optional(),
    rows: z
      .array(z.array(z.string().max(2_000)).max(20))
      .max(200)
      .optional(),
  })
  .passthrough();

const presentationChartAreaPaintSchema = z.union([
  z.literal("none"),
  z.string().regex(/^#[0-9a-fA-F]{6}$/),
]);
// An element's Format pane detail the server reads and writes beside its fill
// and outline (`office_chart_shape_paint.chart_shape_paint`): line width,
// gradients, dash, effects. Without it, every deck whose chart states one
// (a 1.5 pt pie outline, AI70H H15 B08) failed this schema, and the editor
// commands refused the whole deck as invalid_content.
const presentationChartShapePaintSchema = z
  .record(z.string(), z.unknown())
  .refine((value) => {
    try {
      return normalizeOfficeChartShapePaint(value) !== undefined;
    } catch {
      // The normalizer throws on a key or value the paint cannot hold.
      return false;
    }
  });
const presentationChartAreaStyleSchema = z
  .object({
    fill: presentationChartAreaPaintSchema.optional(),
    outline: presentationChartAreaPaintSchema.optional(),
    shapePaint: presentationChartShapePaintSchema.optional(),
    // A chart's title as read on the block: its own font, its words' paint. A title
    
    text: z.record(z.string(), z.unknown()).optional(),
    textFill: presentationChartAreaPaintSchema.optional(),
  })
  .strict()
  .refine(
    (value) =>
      value.fill !== undefined || value.outline !== undefined || value.text !== undefined || value.textFill !== undefined,
  );

export const presentationChartBlockSchema = z
  .object({
    id: z.string().min(1),
    type: z.literal("chart"),
    chartV2: z.custom<ChartBlockAttrsV2>((value) => readPresentationChartV2(value) !== null).optional(),
    chartWorkbook: z.unknown().optional(),
    title: z.string().max(2_000).optional(),
    caption: z.string().max(2_000).optional(),
    chartType: z.string().max(200).optional(),
    ranges: z.array(z.string().max(500)).max(50).optional(),
    // Chart toggles (Excel/WPS parity): show/hide legend + axis labels.
    // Optional/additive so existing charts still validate; absent means shown
    // (defaults ON, matching Excel). MiniChart reads these via resolveChartToggles.
    showLegend: z.boolean().optional(),
    legendPosition: z.enum(["right", "left", "top", "bottom"]).optional(),
    // The six kinds of Office's Trendline menu; the one normaliser the three
    // editors share decides what is a trendline.
    trendline: z
      .custom<OfficeTrendline>(
        (value) => normalizeOfficeTrendline(value) !== undefined,
        "Invalid chart trendline",
      )
      .optional(),
    showAxisLabels: z.boolean().optional(),
    showDataLabels: z.boolean().optional(),
    showGridlines: z.boolean().optional(),
    showDataTable: z.boolean().optional(),
    axisTitleX: z.string().max(2_000).optional(),
    axisTitleY: z.string().max(2_000).optional(),
    chartStyle: z.number().int().min(1).max(48).optional(),
    chartAreaStyle: presentationChartAreaStyleSchema.optional(),
    plotAreaStyle: presentationChartAreaStyleSchema.optional(),
    chartTitleStyle: presentationChartAreaStyleSchema.optional(),
    legendStyle: presentationChartAreaStyleSchema.optional(),
    seriesStyles: z
      .array(
        z
          .object({
            seriesIndex: z.number().int().min(0).max(11),
            fill: z.string().regex(/^#[0-9a-fA-F]{6}$/).optional(),
            outline: presentationChartAreaPaintSchema.optional(),
            shapePaint: presentationChartShapePaintSchema.optional(),
          })
          .strict()
          .refine(
            (style) => style.fill !== undefined || style.outline !== undefined || style.shapePaint !== undefined,
          ),
      )
      .min(1)
      .max(12)
      .refine(
        (items) => new Set(items.map((item) => item.seriesIndex)).size === items.length,
      )
      .optional(),
    // An axis's line, what of its width, dash and 3-D format the import
    // authenticates beside it (`shapePaint`), and its labels' own text, as
    // the page validator admits (it refuses a fill on an axis): an axis line
    
    
    axisStyles: z
      .array(
        z
          .object({
            axis: z.enum(["category", "value", "secondaryValue"]),
            outline: presentationChartAreaPaintSchema.optional(),
            shapePaint: presentationChartShapePaintSchema.optional(),
            text: z.record(z.string(), z.unknown()).optional(),
          })
          .strict()
          .refine(
            (style) => style.outline !== undefined || style.shapePaint !== undefined || style.text !== undefined,
          ),
      )
      .min(1)
      .max(3)
      .refine(
        (items) => new Set(items.map((item) => item.axis)).size === items.length,
      )
      .optional(),
    axisTitleStyles: z
      .array(
        z
          .object({
            axis: z.enum(["category", "value"]),
            fill: presentationChartAreaPaintSchema.optional(),
            outline: presentationChartAreaPaintSchema.optional(),
          })
          .strict()
          .refine(
            (style) => style.fill !== undefined || style.outline !== undefined,
          ),
      )
      .min(1)
      .max(2)
      .refine(
        (items) => new Set(items.map((item) => item.axis)).size === items.length,
      )
      .optional(),
    // A line's paint, and what of its width, dash and effects the import
    // authenticates beside it (`shapePaint`, as the page validator admits:
    
    // command of the deck, "Unrecognized key(s) in object: 'shapePaint'").
    gridlineStyle: z
      .object({
        outline: presentationChartAreaPaintSchema,
        shapePaint: presentationChartShapePaintSchema.optional(),
      })
      .strict()
      .optional(),
    trendlineStyle: z
      .object({
        outline: presentationChartAreaPaintSchema,
        shapePaint: presentationChartShapePaintSchema.optional(),
      })
      .strict()
      .optional(),
    dataTableStyle: presentationChartAreaStyleSchema.optional(),
    dataLabelStyle: presentationChartAreaStyleSchema.optional(),
    sliceStyles: z
      .array(
        z
          .object({
            pointIndex: z.number().int().min(0).max(49),
            fill: presentationChartAreaPaintSchema.optional(),
            outline: presentationChartAreaPaintSchema.optional(),
            // A point's own 3-D Format, width and dash, as the page validator
            // admits beside (or without) its paint.
            shapePaint: presentationChartShapePaintSchema.optional(),
          })
          .strict()
          .refine(
            (style) => style.fill !== undefined || style.outline !== undefined || style.shapePaint !== undefined,
          ),
      )
      .min(1)
      .max(50)
      .refine(
        (items) => new Set(items.map((item) => item.pointIndex)).size === items.length,
      )
      .optional(),
    data: presentationChartDataSchema.optional(),
    // The picture PowerPoint saved of an imported chart this page cannot
    // plot (its frame's mc:Fallback), drawn in the frame instead of nothing.
    nativeChartFallbackSrc: z
      .string()
      .max(2_000_000)
      .regex(/^(?:data:image\/(?:png|jpeg|gif);base64,|mrai-media:[0-9a-f]{64}$)/i)
      .optional(),
    source: z.record(z.unknown()).optional(),
    left: z
      .number()
      .min(-PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    top: z
      .number()
      .min(-PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    width: z
      .number()
      .positive()
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    height: z
      .number()
      .positive()
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    rotation: z.number().min(-360).max(360).optional(),
    // Optional horizontal / vertical mirror flags (PowerPoint `flipH`/`flipV`
    // parity). Purely additive booleans: a block without either omits both and
    // validates byte-for-byte. Render-only — the flip is composed into the
    // block's CSS transform (flip THEN rotate, matching PowerPoint's xfrm) and
    // never mutates the geometry (left/top/width/height) or the text. A
    // hand-edited `false` is dropped to absent on the editor write path so the
    // serialized form stays minimal.
    flipH: z.boolean().optional(),
    flipV: z.boolean().optional(),
    name: z.string().trim().min(1).max(200).optional(),
    hidden: z.boolean().optional(),
    sourceOrder: z.number().int().min(0).max(MAX_PRESENTATION_SLIDE_BLOCKS).optional(),
    zIndex: z.number().int().min(0).max(MAX_PRESENTATION_SLIDE_BLOCKS).optional(),
    animation: presentationBlockAnimationSchema.optional(),
    href: presentationBlockHrefSchema.optional(),
    linkTooltip: presentationLinkTooltipSchema.optional(),
    groupId: presentationGroupIdSchema.optional(),
    // The groups nested inside `groupId` this block also belongs to,
    // outermost first (presentation-crop-group `blockGroupPath`).
    groupPath: z.array(presentationGroupIdSchema).min(1).max(7).optional(),
  })
  .passthrough()
  .refine((block) => block.chartWorkbook === undefined ||
    (block.chartV2 !== undefined && readPresentationChartWorkbook(block.chartWorkbook, block.chartV2) !== null),
  { message: "Chart workbook and chart data must agree" });

// Additive media / embed block (PowerPoint parity). `kind` selects how the URL
// is rendered: `iframe` (YouTube/Vimeo/generic sandboxed iframe), `video`
// (native `<video>`), or `audio` (native `<audio>`). The URL is validated
// through `isPersistableEmbedUrlOrEmpty`, which admits ONLY https URLs
// (rejecting `javascript:`/`data:`/`file:`/`blob:`/`http:` and protocol-relative
// refs) plus an empty draft string, so a hand-edited deck can never persist a
// dangerous src; the render-time resolver (lib/presentation-embed
// `resolveEmbed`) re-validates against the block's actual kind and always
// strict-sandboxes iframes. Existing decks contain no `embed` blocks, so this
// union member is purely additive and never changes validation of pre-existing
// content.
const presentationMediaPlaybackSettingsSchema = z
  .object({
    start: z.enum(PRESENTATION_MEDIA_PLAYBACK_STARTS).optional(),
    volume: z.enum(PRESENTATION_MEDIA_VOLUMES).optional(),
    loop: z.boolean().optional(),
    rewindAfterPlaying: z.boolean().optional(),
    playAcrossSlides: z.boolean().optional(),
    hideDuringShow: z.boolean().optional(),
    playFullScreen: z.boolean().optional(),
    trimStartMs: z
      .number()
      .int()
      .min(0)
      .max(MAX_PRESENTATION_MEDIA_TIME_MS)
      .optional(),
    trimEndMs: z
      .number()
      .int()
      .min(0)
      .max(MAX_PRESENTATION_MEDIA_TIME_MS)
      .nullable()
      .optional(),
    fadeInMs: z
      .number()
      .int()
      .min(0)
      .max(MAX_PRESENTATION_MEDIA_FADE_MS)
      .optional(),
    fadeOutMs: z
      .number()
      .int()
      .min(0)
      .max(MAX_PRESENTATION_MEDIA_FADE_MS)
      .optional(),
    // Playback > Add Bookmark (`p14:bmk`, PPT-A157).
    bookmarks: z
      .array(
        z
          .object({
            name: z
              .string()
              .min(1)
              .max(MAX_PRESENTATION_MEDIA_BOOKMARK_NAME_CHARS)
              .refine((name) => name.trim().length > 0),
            timeMs: z.number().int().min(0).max(MAX_PRESENTATION_MEDIA_TIME_MS),
          })
          .strict(),
      )
      .max(MAX_PRESENTATION_MEDIA_BOOKMARKS)
      .refine(
        (bookmarks) =>
          new Set(bookmarks.map((bookmark) => bookmark.name)).size ===
          bookmarks.length,
        { message: "Media bookmark names must be unique" },
      )
      .optional(),
  })
  .strict()
  .superRefine((settings, context) => {
    if (
      settings.trimEndMs !== undefined &&
      settings.trimEndMs !== null &&
      settings.trimEndMs <= (settings.trimStartMs ?? 0)
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["trimEndMs"],
        message: "Media trim end must follow trim start",
      });
    }
  });

const presentationMediaCaptionTrackSchema = z
  .object({
    id: z.string().min(1).max(64),
    label: z
      .string()
      .min(1)
      .max(MAX_PRESENTATION_MEDIA_CAPTION_LABEL_CHARS),
    language: z.string().min(2).max(64).optional(),
    vtt: z.string().min(1),
  })
  .strict()
  .superRefine((track, context) => {
    if (normalizePresentationMediaCaptionTrack(track) === null) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        message: "Media caption track must be canonical bounded WebVTT",
      });
    }
  })
  .transform((track) => normalizePresentationMediaCaptionTrack(track)!);

const presentationMediaCaptionTracksSchema = z
  .array(presentationMediaCaptionTrackSchema)
  .min(1)
  .max(MAX_PRESENTATION_MEDIA_CAPTION_TRACKS)
  .superRefine((tracks, context) => {
    if (normalizePresentationMediaCaptionTracks(tracks) === null) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        message: "Media caption track identities must be unique",
      });
    }
  });

export const presentationEmbedBlockSchema = z
  .object({
    id: z.string().min(1),
    type: z.literal("embed"),
    // Video Format > Video Shape, a video's own closed preset (`a:prstGeom`).
    clipShape: z
      .string()
      .regex(/^[A-Za-z0-9]{1,40}$/)
      .refine((value) => value !== "rect", "A rectangle is no video shape")
      .optional(),
    // Video Format > Video Border, a video's own p:spPr/a:ln (picture border fields).
    borderColor: z
      .string()
      .regex(/^#[0-9a-fA-F]{6}$/)
      .optional(),
    borderWidth: z.number().min(0).max(24).optional(),
    borderStyle: z.enum(["none", "solid", "dashed", "dotted"]).optional(),
    // Video Format > Video Effects, a video's own p:spPr effects.
    pictureEffects: z
      .record(z.string(), z.unknown())
      .refine((value) => normalizeOfficePictureEffects(value) !== null, "Unsupported video effects")
      .optional(),
    // Video Format > Corrections / Color (the poster blip's a:lum and
    // Recolor preset), Video Styles (a whole spPr recipe) and Poster Frame (a
    // poster the user set), PPT-A155: lib/presentation-video-format.
    videoCorrections: z
      .object({
        brightness: z.number().int().min(-100).max(100),
        contrast: z.number().int().min(-100).max(100),
      })
      .strict()
      .refine((value) => value.brightness !== 0 || value.contrast !== 0, "Neutral corrections are omitted")
      .optional(),
    videoRecolor: z.enum(OFFICE_RECOLOR_PRESET_IDS).optional(),
    videoStyle: z.enum(OFFICE_VIDEO_STYLE_IDS).optional(),
    posterSrc: z
      .string()
      .refine((value) => isVideoPosterSource(value), "A poster frame is an inline PNG or JPEG of at most 1 MiB")
      .optional(),
    kind: z.enum(EMBED_KINDS).optional(),
    // Required field. Empty drafts and safe https URLs remain supported; a
    // strict bounded audio data URL is admitted only for kind=audio.
    url: z.string().max(MAX_PRESENTATION_INLINE_AUDIO_DATA_URL_CHARS),
    // A local video or audio clip keeps the DataItem as source of truth. The
    // union is complete or absent; no signed/temporary URL is persisted in
    // the deck. A clip's size is bounded by the deck's media total only
    // (lib/document-size-limits); a video's by its 256 MB as well.
    sourceKind: z.literal("asset").optional(),
    assetId: z.string().min(1).max(200).optional(),
    mimeType: z
      .enum(["video/mp4", ...PRESENTATION_AUDIO_ASSET_MIME_TYPES])
      .optional(),
    sizeBytes: z
      .number()
      .int()
      .positive()
      .max(Number.MAX_SAFE_INTEGER)
      .optional(),
    title: z.string().max(2_000).optional(),
    playback: presentationMediaPlaybackSettingsSchema.optional(),
    captions: presentationMediaCaptionTracksSchema.optional(),
    left: z
      .number()
      .min(-PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    top: z
      .number()
      .min(-PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    width: z
      .number()
      .positive()
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    height: z
      .number()
      .positive()
      .max(PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT)
      .optional(),
    rotation: z.number().min(-360).max(360).optional(),
    // Optional horizontal / vertical mirror flags (PowerPoint `flipH`/`flipV`
    // parity). Purely additive booleans: a block without either omits both and
    // validates byte-for-byte. Render-only — the flip is composed into the
    // block's CSS transform (flip THEN rotate, matching PowerPoint's xfrm) and
    // never mutates the geometry (left/top/width/height) or the text. A
    // hand-edited `false` is dropped to absent on the editor write path so the
    // serialized form stays minimal.
    flipH: z.boolean().optional(),
    flipV: z.boolean().optional(),
    name: z.string().trim().min(1).max(200).optional(),
    hidden: z.boolean().optional(),
    sourceOrder: z.number().int().min(0).max(MAX_PRESENTATION_SLIDE_BLOCKS).optional(),
    zIndex: z.number().int().min(0).max(MAX_PRESENTATION_SLIDE_BLOCKS).optional(),
    animation: presentationBlockAnimationSchema.optional(),
    href: presentationBlockHrefSchema.optional(),
    linkTooltip: presentationLinkTooltipSchema.optional(),
    groupId: presentationGroupIdSchema.optional(),
    // The groups nested inside `groupId` this block also belongs to,
    // outermost first (presentation-crop-group `blockGroupPath`).
    groupPath: z.array(presentationGroupIdSchema).min(1).max(7).optional(),
  })
  .passthrough()
  .superRefine((block, context) => {
    const record = block as unknown as Record<string, unknown>;
    const url = typeof record.url === "string" ? record.url : "";
    const localAudio =
      record.kind === "audio" &&
      isPersistablePresentationAudioDataUrl(url);
    if (
      !localAudio &&
      (url.length > 4_000 || !isPersistableEmbedUrlOrEmpty(url))
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["url"],
        message: "Unsupported or unsafe embed URL",
      });
    }
    if (
      record.kind !== "video" &&
      (record.borderColor !== undefined ||
        record.borderWidth !== undefined ||
        record.borderStyle !== undefined)
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["borderStyle"],
        message: "Video Border is a video's only",
      });
    }
    const videoFormatField = ["videoCorrections", "videoRecolor", "videoStyle", "posterSrc"].find(
      (key) => record[key] !== undefined,
    );
    // An audio clip's own picture (its speaker icon) rides as `posterSrc`.
    const audioIcon = record.kind === "audio" && videoFormatField === "posterSrc" &&
      ["videoCorrections", "videoRecolor", "videoStyle"].every((key) => record[key] === undefined);
    if (record.kind !== "video" && videoFormatField && !audioIcon) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: [videoFormatField],
        message: "Video Format is a video's only",
      });
    }
    if (record.videoRecolor === "washout" && record.videoCorrections !== undefined) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["videoCorrections"],
        message: "Washout and Corrections share the video's one luminance",
      });
    }
    
    // stays the style's.
    if (record.videoStyle !== undefined && record.clipShape !== undefined) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["clipShape"],
        message: "A video style's geometry is its own",
      });
    }
    if (
      (record.kind === undefined || record.kind === "iframe") &&
      (record.playback !== undefined || record.captions !== undefined)
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: [record.captions !== undefined ? "captions" : "playback"],
        message:
          "Provider-owned iframe playback and caption settings are not persisted",
      });
    }
    const hasAssetField = ["sourceKind", "assetId", "mimeType", "sizeBytes"].some(
      (key) => key in record,
    );
    if (hasAssetField && assetBackedPresentationMediaSource(record) === null) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["sourceKind"],
        message: "Local media requires complete DataItem custody",
      });
    }
    const source =
      record.source &&
      typeof record.source === "object" &&
      !Array.isArray(record.source)
        ? (record.source as Record<string, unknown>)
        : null;
    if (
      source?.nativeVideoAuthority !== undefined &&
      (record.kind !== "video" ||
        parseNativePresentationVideoAuthority(source.nativeVideoAuthority) ===
          null)
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["source", "nativeVideoAuthority"],
        message: "native video authority is malformed",
      });
    }
    if (
      source?.nativeAudioAuthority !== undefined &&
      (record.kind !== "audio" ||
        parseNativePresentationAudioAuthority(source.nativeAudioAuthority) ===
          null)
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["source", "nativeAudioAuthority"],
        message: "native audio authority is malformed",
      });
    }
    if (
      source?.nativeCaptionAuthority !== undefined &&
      (() => {
        const authority = parseNativePresentationCaptionAuthority(
          source.nativeCaptionAuthority,
        );
        return (
          (record.kind !== "video" && record.kind !== "audio") ||
          authority === null ||
          source.slideId !== authority.slideId ||
          source.shapeId !== authority.shapeId
        );
      })()
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["source", "nativeCaptionAuthority"],
        message: "native media caption authority is malformed",
      });
    }
  });

export const presentationBlockSchema = z.union([
  presentationTextBlockSchema,
  presentationImageBlockSchema,
  presentationTableBlockSchema,
  presentationChartBlockSchema,
  presentationShapeBlockSchema,
  presentationEmbedBlockSchema,
]).superRefine((block, context) => {
  const source =
    block.source && typeof block.source === "object" && !Array.isArray(block.source)
      ? (block.source as Record<string, unknown>)
      : null;
  if (
    source &&
    "slideSizeOriginalGeometry" in source &&
    nativeSlideSizeOriginalGeometry(block) === null
  ) {
    context.addIssue({
      code: z.ZodIssueCode.custom,
      path: ["source", "slideSizeOriginalGeometry"],
      message: "native slide size geometry receipt is malformed",
    });
  }
});

// Optional ANIMATION TIMELINE step (PowerPoint animation-pane parity). One
// ordered entry on a slide's `animationSequence`: which block animates, the
// entrance animation (shared with the per-block fx set), how it STARTS relative
// to its neighbours (on-click / with-previous / after-previous), and its delay /
// duration in ms. Purely additive — existing decks omit `animationSequence` and
// still validate; the render-time normalizer (lib/presentation-anim-timeline)
// re-clamps/drops invalid steps, so this schema is a generous structural guard,
// not the render-time invariant. `id` is an optional editor key; `blockId`
// references a block on the same slide (a dangling ref is dropped at render).
// Optional MOTION PATH on a step (PowerPoint's fourth animation class): a named
// preset (line/arc/loop/…) or a `custom` list of EMU waypoint offsets from the
// block's start. Purely additive — a step without it is a plain
// entrance/emphasis/exit step. Bounded structurally here (preset enum, waypoint
// count + EMU window); the render-time normalizer (lib/presentation-motion-path)
// re-clamps/drops a malformed value, so this is a generous structural guard, not
// the render-time invariant.
const presentationMotionWaypointSchema = z.object({
  x: z.number().min(-MAX_MOTION_PATH_EMU).max(MAX_MOTION_PATH_EMU),
  y: z.number().min(-MAX_MOTION_PATH_EMU).max(MAX_MOTION_PATH_EMU),
});
const presentationMotionPathSchema = z
  .object({
    preset: z.enum(MOTION_PATH_PRESETS),
    waypoints: z
      .array(presentationMotionWaypointSchema)
      .max(MAX_MOTION_PATH_WAYPOINTS)
      .optional(),
  })
  .passthrough();

const presentationAnimationStepSchema = z
  .object({
    id: z.string().min(1).max(200).optional(),
    blockId: z.string().min(1).max(200),
    animation: z.enum(TIMELINE_ANIMATIONS),
    effectOption: z.enum(PRESENTATION_ANIMATION_EFFECT_OPTIONS).optional(),
    effectColor: z.string().regex(/^#[0-9a-f]{6}$/).optional(),
    trigger: z.enum(TIMELINE_TRIGGERS),
    delayMs: z.number().min(0).max(MAX_STEP_DELAY_MS),
    durationMs: z.number().min(0).max(MAX_STEP_DURATION_MS),
    motionPath: presentationMotionPathSchema.optional(),
    // Imported By Paragraph build (zero-based, inclusive). Read-only.
    paragraphs: z
      .object({
        start: z.number().int().min(0).max(9999),
        end: z.number().int().min(0).max(9999),
      })
      .strict()
      .refine((range) => range.start <= range.end)
      .optional(),
  })
  .passthrough()
  .superRefine((step, context) => {
    if (step.effectOption !== undefined) {
      const resolution = resolvePresentationAnimationEffectOption(
        step.animation,
        step.effectOption,
      );
      if (!resolution || resolution.effectOption !== step.effectOption) {
        context.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["effectOption"],
          message: "animation effectOption must be an exact non-default variant",
        });
      }
    }
    const effect = presentationAnimationEffect(step.animation);
    const color = resolvePresentationAnimationEffectColor(
      step.animation,
      step.effectColor,
    );
    if (
      (effect?.nativeBehavior === "color" &&
        (!color || color.effectColor !== step.effectColor)) ||
      (effect?.nativeBehavior !== "color" && step.effectColor !== undefined)
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["effectColor"],
        message: "animation effectColor must be canonical for a color effect",
      });
    }
  });

const presentationAnimationReceiptStepSchema = z
  .object({
    kind: z
      .enum(TIMELINE_ANIMATIONS)
      .refine(
        (kind) => presentationAnimationEffect(kind)?.authoring === "native",
      ),
    nativeShapeId: z.string().regex(/^shp[1-9][0-9]*$/),
    triggerNativeShapeId: z.string().regex(/^shp[1-9][0-9]*$/).optional(),
    effectOption: z.enum(PRESENTATION_ANIMATION_EFFECT_OPTIONS).optional(),
    effectColor: z.string().regex(/^#[0-9a-f]{6}$/).optional(),
    trigger: z.enum(TIMELINE_TRIGGERS),
    delayMs: z.number().refine(isEditableAnimationTimingMs),
    durationMs: z.number().refine(isEditableAnimationDurationMs),
  })
  .strict()
  .superRefine((step, context) => {
    if (step.triggerNativeShapeId === step.nativeShapeId) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["triggerNativeShapeId"],
        message: "animation trigger target must differ from effect target",
      });
    }
    const resolution = resolvePresentationAnimationEffectOption(
      step.kind,
      step.effectOption,
    );
    if (!resolution || resolution.effectOption !== step.effectOption) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["effectOption"],
        message: "animation receipt option is not canonical for its effect",
      });
    }
    const effect = presentationAnimationEffect(step.kind);
    const color = resolvePresentationAnimationEffectColor(
      step.kind,
      step.effectColor,
    );
    if (
      (effect?.nativeBehavior === "color" &&
        (!color || color.effectColor !== step.effectColor)) ||
      (effect?.nativeBehavior !== "color" && step.effectColor !== undefined)
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["effectColor"],
        message: "animation receipt color is not canonical for its effect",
      });
    }
  });

const presentationNativeSlideAnimationReceiptSchema = z
  .object({
    slideAnimationSchemaVersion: z.literal(
      PRESENTATION_SLIDE_ANIMATION_SCHEMA_VERSION,
    ),
    slideAnimationDisposition: z.literal("native"),
    slideAnimationContainerPresent: z.boolean(),
    slideAnimationOriginalEffectCount: z
      .number()
      .int()
      .min(0)
      .max(MAX_TIMELINE_STEPS),
    slideAnimationOriginalSequence: z
      .array(presentationAnimationReceiptStepSchema)
      .max(MAX_TIMELINE_STEPS),
    slideAnimationFingerprint: z.string().regex(/^[0-9a-f]{64}$/),
  })
  .passthrough()
  .superRefine((receipt, context) => {
    if (
      receipt.slideAnimationOriginalSequence.length !==
        receipt.slideAnimationOriginalEffectCount ||
      // An empty container is a real state: PowerPoint 2007-2010 keep
      // `<p:timing>` with only the root time node on a slide whose effects
      
      (!receipt.slideAnimationContainerPresent &&
        receipt.slideAnimationOriginalEffectCount > 0) ||
      Object.hasOwn(receipt, "slideAnimationRejectReason")
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        message: "native slide animation receipt is internally inconsistent",
      });
    }
  });

const presentationPreserveOnlySlideAnimationReceiptSchema = z
  .object({
    slideAnimationSchemaVersion: z.literal(
      PRESENTATION_SLIDE_ANIMATION_SCHEMA_VERSION,
    ),
    slideAnimationDisposition: z.literal("preserve-only"),
    slideAnimationContainerPresent: z.boolean(),
    slideAnimationOriginalEffectCount: z.null(),
    slideAnimationOriginalSequence: z.null(),
    slideAnimationFingerprint: z.string().regex(/^[0-9a-f]{64}$/),
    slideAnimationRejectReason: z.string().min(1).max(160),
  })
  .passthrough();

export const presentationSlideAnimationReceiptSchema = z.union([
  presentationNativeSlideAnimationReceiptSchema,
  presentationPreserveOnlySlideAnimationReceiptSchema,
]);

const PRESENTATION_SLIDE_ANIMATION_RECEIPT_KEYS = [
  "slideAnimationSchemaVersion",
  "slideAnimationDisposition",
  "slideAnimationContainerPresent",
  "slideAnimationOriginalEffectCount",
  "slideAnimationOriginalSequence",
  "slideAnimationFingerprint",
  "slideAnimationRejectReason",
] as const;

const presentationNativeSlideTransitionReceiptSchema = z
  .object({
    slideTransitionSchemaVersion: z.literal(
      PRESENTATION_SLIDE_TRANSITION_SCHEMA_VERSION,
    ),
    slideTransitionDisposition: z.literal("native"),
    slideTransitionContainerPresent: z.boolean(),
    slideTransitionEffectPresent: z.boolean(),
    // Morph is written back natively (its Effect Options too, readings row
    // 238): a deck with a Morph slide was refused to every editor command
    
    slideTransitionOriginalKind: z.enum(AUTHORABLE_SLIDE_TRANSITIONS),
    slideTransitionOriginalDurationMs: z
      .number()
      .int()
      .min(MIN_TRANSITION_DURATION_MS)
      .max(MAX_TRANSITION_DURATION_MS)
      .nullable(),
    slideTransitionOriginalDurationSource: z
      .enum(["exact", "speed", "default"])
      .nullable(),
    slideTransitionOriginalSpeed: z.enum(["fast", "med", "slow"]).nullable(),
    slideTransitionOriginalDirection: z
      .enum(SLIDE_TRANSITION_EFFECT_OPTIONS)
      .nullable()
      .optional(),
    slideTransitionOriginalAdvanceOnClick: z.boolean().optional(),
    slideTransitionOriginalAdvanceAfterMs: z
      .number()
      .int()
      .min(MIN_TRANSITION_ADVANCE_AFTER_MS)
      .max(MAX_TRANSITION_ADVANCE_AFTER_MS)
      .nullable()
      .optional(),
    slideTransitionFingerprint: z.string().regex(/^[0-9a-f]{64}$/),
  })
  .passthrough()
  .superRefine((receipt, context) => {
    const effectKind = receipt.slideTransitionOriginalKind;
    if (
      receipt.slideTransitionEffectPresent !== (effectKind !== "none") ||
      (receipt.slideTransitionEffectPresent &&
        (receipt.slideTransitionOriginalDurationMs === null ||
          !receipt.slideTransitionContainerPresent)) ||
      (!receipt.slideTransitionEffectPresent &&
        (receipt.slideTransitionOriginalDurationMs !== null ||
          receipt.slideTransitionOriginalDurationSource !== null)) ||
      (receipt.slideTransitionEffectPresent &&
        receipt.slideTransitionOriginalDurationSource === null) ||
      (receipt.slideTransitionOriginalDurationSource === "speed" &&
        receipt.slideTransitionOriginalSpeed === null) ||
      (receipt.slideTransitionOriginalDurationSource === "default" &&
        receipt.slideTransitionOriginalSpeed !== null) ||
      (!receipt.slideTransitionContainerPresent &&
        receipt.slideTransitionOriginalSpeed !== null) ||
      (isDirectionalSlideTransition(effectKind) &&
        effectKind !== "push" &&
        effectKind !== "wipe" &&
        !Object.hasOwn(receipt, "slideTransitionOriginalDirection")) ||
      (Object.hasOwn(receipt, "slideTransitionOriginalDirection") &&
        (isDirectionalSlideTransition(effectKind)
          ? // A preset receipt stamped before its Left/Right Effect Options
            // names no side; the one identity native then was the right-hand
            // one, so it still says which slide it came from.
            !(
              receipt.slideTransitionOriginalDirection === null &&
              isSlideTransitionP15LeftRightKind(effectKind)
            ) &&
            !isSlideTransitionDirectionForKind(
              effectKind,
              receipt.slideTransitionOriginalDirection,
            )
          : receipt.slideTransitionOriginalDirection !== null)) ||
      Object.hasOwn(receipt, "slideTransitionOriginalAdvanceOnClick") !==
        Object.hasOwn(receipt, "slideTransitionOriginalAdvanceAfterMs") ||
      Object.hasOwn(receipt, "slideTransitionRejectReason")
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        message: "native slide transition receipt is internally inconsistent",
      });
    }
  });

const presentationPreserveOnlySlideTransitionReceiptSchema = z
  .object({
    slideTransitionSchemaVersion: z.literal(
      PRESENTATION_SLIDE_TRANSITION_SCHEMA_VERSION,
    ),
    slideTransitionDisposition: z.literal("preserve-only"),
    slideTransitionContainerPresent: z.boolean(),
    slideTransitionEffectPresent: z.null(),
    slideTransitionOriginalKind: z.null(),
    slideTransitionOriginalDurationMs: z.null(),
    slideTransitionOriginalDurationSource: z.null(),
    slideTransitionOriginalSpeed: z.null(),
    slideTransitionOriginalDirection: z.null().optional(),
    slideTransitionOriginalAdvanceOnClick: z.null().optional(),
    slideTransitionOriginalAdvanceAfterMs: z.null().optional(),
    slideTransitionFingerprint: z.string().regex(/^[0-9a-f]{64}$/),
    slideTransitionRejectReason: z.string().min(1).max(160),
  })
  .passthrough()
  .superRefine((receipt, context) => {
    if (
      Object.hasOwn(receipt, "slideTransitionOriginalAdvanceOnClick") !==
      Object.hasOwn(receipt, "slideTransitionOriginalAdvanceAfterMs")
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        message: "preserve-only slide transition receipt is incomplete",
      });
    }
  });

export const presentationSlideTransitionReceiptSchema = z.union([
  presentationNativeSlideTransitionReceiptSchema,
  presentationPreserveOnlySlideTransitionReceiptSchema,
]);

const PRESENTATION_SLIDE_TRANSITION_RECEIPT_KEYS = [
  "slideTransitionSchemaVersion",
  "slideTransitionDisposition",
  "slideTransitionContainerPresent",
  "slideTransitionEffectPresent",
  "slideTransitionOriginalKind",
  "slideTransitionOriginalDurationMs",
  "slideTransitionOriginalDurationSource",
  "slideTransitionOriginalSpeed",
  "slideTransitionOriginalDirection",
  "slideTransitionOriginalAdvanceOnClick",
  "slideTransitionOriginalAdvanceAfterMs",
  "slideTransitionFingerprint",
  "slideTransitionRejectReason",
] as const;

const PRESENTATION_SLIDE_MASTER_SHAPES_SCHEMA_VERSION =
  "pptx-slide-master-shapes-v1";
export const presentationSlideMasterShapesReceiptSchema = z.union([
  z
    .object({
      slideMasterShapesSchemaVersion: z.literal(
        PRESENTATION_SLIDE_MASTER_SHAPES_SCHEMA_VERSION,
      ),
      slideMasterShapesDisposition: z.literal("native"),
      slideMasterShapesOriginalShow: z.boolean(),
    })
    .passthrough(),
  z
    .object({
      slideMasterShapesSchemaVersion: z.literal(
        PRESENTATION_SLIDE_MASTER_SHAPES_SCHEMA_VERSION,
      ),
      slideMasterShapesDisposition: z.literal("preserve-only"),
    })
    .passthrough(),
]);

const PRESENTATION_SLIDE_MASTER_SHAPES_RECEIPT_KEYS = [
  "slideMasterShapesSchemaVersion",
  "slideMasterShapesDisposition",
  "slideMasterShapesOriginalShow",
] as const;

const PRESENTATION_SLIDE_BACKGROUND_SCHEMA_VERSION = "pptx-slide-background-v1";
const presentationSlideBackgroundReceiptSchema = z.union([
  z
    .object({
      slideBackgroundSchemaVersion: z.literal(
        PRESENTATION_SLIDE_BACKGROUND_SCHEMA_VERSION,
      ),
      slideBackgroundDisposition: z.literal("native"),
      // No `p:bg` of its own: it shows its layout's (else its master's), and
      // the new layout's after Home > Layout (readings row 503).
      slideBackgroundFollowsLayout: z.literal(true).optional(),
    })
    .passthrough(),
  z
    .object({
      slideBackgroundSchemaVersion: z.literal(
        PRESENTATION_SLIDE_BACKGROUND_SCHEMA_VERSION,
      ),
      slideBackgroundDisposition: z.literal("preserve-only"),
      slideBackgroundRejectReason: z.string().min(1).max(160),
    })
    .passthrough(),
]);
const PRESENTATION_SLIDE_BACKGROUND_RECEIPT_KEYS = [
  "slideBackgroundSchemaVersion",
  "slideBackgroundDisposition",
  "slideBackgroundRejectReason",
  "slideBackgroundFollowsLayout",
] as const;

export const presentationSlideSourceSchema = z
  .record(z.unknown())
  .superRefine((source, context) => {
    // A stored page names the receipts every slide repeats
    // (`presentation-deck-receipts`); the loader writes them out again, and
    // they are checked in that form.
    const storedSlideSize = isPresentationDeckReceiptRef(source.slideSizeAuthority);
    if (
      source.slideSizeAuthority !== undefined &&
      !storedSlideSize &&
      parseNativeSlideSizeAuthority(source.slideSizeAuthority) === null
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["slideSizeAuthority"],
        message: "native slide size authority is malformed",
      });
    }
    const slideSizeRevision = source.slideSizeRevision;
    const slideSizeStrategy = source.slideSizeFitStrategy;
    if (
      slideSizeRevision !== undefined &&
      (typeof slideSizeRevision !== "number" ||
        !Number.isInteger(slideSizeRevision) ||
        slideSizeRevision < 0 ||
        slideSizeRevision > 2_147_483_647)
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["slideSizeRevision"],
        message: "native slide size revision is malformed",
      });
    }
    if (
      slideSizeStrategy !== undefined &&
      slideSizeStrategy !== "maximize" &&
      slideSizeStrategy !== "ensure-fit"
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["slideSizeFitStrategy"],
        message: "native slide size strategy is malformed",
      });
    }
    const slideSizeIntentPresent =
      slideSizeRevision !== undefined || slideSizeStrategy !== undefined;
    const slideSizeIntentValid =
      typeof slideSizeRevision === "number" &&
      Number.isInteger(slideSizeRevision) &&
      slideSizeRevision >= 0 &&
      (slideSizeRevision === 0
        ? slideSizeStrategy === undefined
        : slideSizeStrategy === "maximize" ||
          slideSizeStrategy === "ensure-fit");
    if (
      slideSizeIntentPresent &&
      (source.slideSizeAuthority === undefined || !slideSizeIntentValid)
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        message: "native slide size intent requires exact custody",
      });
    }
    const resetRevision = source.layoutResetRevision;
    if (
      resetRevision !== undefined &&
      (typeof resetRevision !== "number" ||
        !Number.isInteger(resetRevision) ||
        resetRevision < 0 ||
        resetRevision > 2_147_483_647)
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["layoutResetRevision"],
        message: "native layout reset revision is malformed",
      });
    }
    const layoutCascade = source.layoutCascade;
    if (resetRevision !== undefined && layoutCascade === undefined) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["layoutResetRevision"],
        message: "native layout reset revision requires layout custody",
      });
    }
    const storedLayoutCascade =
      presentationDeckReceiptTokenKey(
        layoutCascade && typeof layoutCascade === "object"
          ? (layoutCascade as Record<string, unknown>).layouts
          : undefined,
      ) !== null;
    if (layoutCascade !== undefined && !storedLayoutCascade) {
      const parsed = parseNativeLayoutCascadeReceipt(layoutCascade);
      if (
        !parsed ||
        source.slideId !== parsed.slideId ||
        (resetRevision !== undefined && !parsed.resetAuthority)
      ) {
        context.addIssue({
          code: z.ZodIssueCode.custom,
          message: "native layout cascade receipt is malformed",
        });
      }
    }
    if (
      source.nativeSlideReuse !== undefined &&
      (Object.keys(source).length !== 1 ||
        parseNativeSlideReuseReceipt(source.nativeSlideReuse) === null)
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["nativeSlideReuse"],
        message: "native reuse slide receipt is malformed",
      });
    }
    if (
      source.transitionSoundDisposition !== undefined &&
      !["absent", "native", "preserve-only"].includes(
        String(source.transitionSoundDisposition),
      )
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["transitionSoundDisposition"],
        message: "transition sound custody is malformed",
      });
    }
    const transitionKeys = Object.keys(source).filter((key) =>
      key.startsWith("slideTransition"),
    );
    if (transitionKeys.length > 0) {
      if (
        transitionKeys.some(
          (key) =>
            !PRESENTATION_SLIDE_TRANSITION_RECEIPT_KEYS.includes(
              key as (typeof PRESENTATION_SLIDE_TRANSITION_RECEIPT_KEYS)[number],
            ),
        ) ||
        (source.slideId !== undefined &&
          (typeof source.slideId !== "string" ||
            !/^sld[1-9][0-9]*$/.test(source.slideId)))
      ) {
        context.addIssue({
          code: z.ZodIssueCode.custom,
          message: "slide transition receipt has an unknown field or slide id",
        });
      } else if (
        !presentationSlideTransitionReceiptSchema.safeParse(source).success
      ) {
        context.addIssue({
          code: z.ZodIssueCode.custom,
          message: "slide transition receipt is malformed",
        });
      }
    }

    const animationKeys = Object.keys(source).filter((key) =>
      key.startsWith("slideAnimation"),
    );
    if (animationKeys.length > 0) {
      if (
        animationKeys.some(
          (key) =>
            !PRESENTATION_SLIDE_ANIMATION_RECEIPT_KEYS.includes(
              key as (typeof PRESENTATION_SLIDE_ANIMATION_RECEIPT_KEYS)[number],
            ),
        ) ||
        (source.slideId !== undefined &&
          (typeof source.slideId !== "string" ||
            !/^sld[1-9][0-9]*$/.test(source.slideId)))
      ) {
        context.addIssue({
          code: z.ZodIssueCode.custom,
          message: "slide animation receipt has an unknown field or slide id",
        });
      } else if (
        !presentationSlideAnimationReceiptSchema.safeParse(source).success
      ) {
        context.addIssue({
          code: z.ZodIssueCode.custom,
          message: "slide animation receipt is malformed",
        });
      }
    }

    const masterShapeKeys = Object.keys(source).filter((key) =>
      key.startsWith("slideMasterShapes"),
    );
    if (masterShapeKeys.length > 0) {
      if (
        masterShapeKeys.some(
          (key) =>
            !PRESENTATION_SLIDE_MASTER_SHAPES_RECEIPT_KEYS.includes(
              key as (typeof PRESENTATION_SLIDE_MASTER_SHAPES_RECEIPT_KEYS)[number],
            ),
        ) ||
        typeof source.slideId !== "string" ||
        !/^sld[1-9][0-9]*$/.test(source.slideId)
      ) {
        context.addIssue({
          code: z.ZodIssueCode.custom,
          message:
            "slide master-shape receipt has an unknown field or slide id",
        });
      } else if (
        !presentationSlideMasterShapesReceiptSchema.safeParse(source).success
      ) {
        context.addIssue({
          code: z.ZodIssueCode.custom,
          message: "slide master-shape receipt is malformed",
        });
      }
    }

    const backgroundKeys = Object.keys(source).filter((key) =>
      key.startsWith("slideBackground"),
    );
    if (backgroundKeys.length > 0) {
      if (
        backgroundKeys.some(
          (key) =>
            !PRESENTATION_SLIDE_BACKGROUND_RECEIPT_KEYS.includes(
              key as (typeof PRESENTATION_SLIDE_BACKGROUND_RECEIPT_KEYS)[number],
            ),
        )
      ) {
        context.addIssue({
          code: z.ZodIssueCode.custom,
          message: "slide background receipt has an unknown field",
        });
      } else if (
        !presentationSlideBackgroundReceiptSchema.safeParse(source).success
      ) {
        context.addIssue({
          code: z.ZodIssueCode.custom,
          message: "slide background receipt is malformed",
        });
      }
    }
  });

const PRESENTATION_COMMENT_UUID_RE =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;
const PRESENTATION_COMMENT_DATETIME_RE =
  /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?$/;
const PRESENTATION_COMMENT_CONTROL_RE = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/;
const presentationCommentXmlTextSchema = (maximum: number) =>
  z
    .string()
    .min(1)
    .max(maximum)
    .refine((value) => !PRESENTATION_COMMENT_CONTROL_RE.test(value));

const presentationCommentReplySchema = z
  .object({
    id: z.string().min(1).max(256),
    body: presentationCommentXmlTextSchema(4000),
    author: presentationCommentXmlTextSchema(120).optional(),
    authorId: presentationCommentXmlTextSchema(256).optional(),
    createdAt: z.string().regex(PRESENTATION_COMMENT_DATETIME_RE).optional(),
  })
  .strict();

const presentationCommentSchema = z
  .object({
    id: z.string().min(1).max(256),
    body: presentationCommentXmlTextSchema(4000),
    author: presentationCommentXmlTextSchema(120).optional(),
    authorId: presentationCommentXmlTextSchema(256).optional(),
    createdAt: z.string().regex(PRESENTATION_COMMENT_DATETIME_RE).optional(),
    resolved: z.boolean().optional(),
    blockId: z.string().min(1).max(256).optional(),
    replies: z.array(presentationCommentReplySchema).max(100).optional(),
  })
  .strict();

const presentationModernCommentRecordSchema = z
  .object({
    id: z.string().regex(PRESENTATION_COMMENT_UUID_RE),
    body: presentationCommentXmlTextSchema(4000),
    author: presentationCommentXmlTextSchema(120),
    authorId: presentationCommentXmlTextSchema(256),
    createdAt: z.string().regex(PRESENTATION_COMMENT_DATETIME_RE),
  })
  .passthrough();

const presentationCommentAuthorityPartSchema = z
  .record(z.string().min(1).max(80), z.string().min(1).max(512))
  .refine((value) => Object.keys(value).length <= 6);

export const presentationCommentAuthoritySchema = z
  .object({
    schemaVersion: z.literal("pptx-comments-v1"),
    mode: z.enum([
      "modern-editable",
      "legacy-preserve-only",
      "blocked-preserve-only",
    ]),
    graphDigest: z.string().regex(/^[0-9a-f]{64}$/),
    authorsPart: z.string().min(1).max(512).nullable(),
    authorsRelId: z.string().min(1).max(512).nullable(),
    commentParts: z.array(presentationCommentAuthorityPartSchema).max(512),
    contentTypes: z.array(presentationCommentAuthorityPartSchema).max(512),
    sourceHashes: z
      .record(z.string().min(1).max(512), z.string().regex(/^[0-9a-f]{64}$/))
      .refine((value) => Object.keys(value).length <= 512),
    blockerReason: z.string().max(160).optional(),
  })
  .strict();

const presentationTransitionSoundSchema = z
  .custom<PresentationTransitionSound>(
    (value) => normalizePresentationTransitionSound(value) !== null,
    "Transition sound must be one bounded embedded WAV",
  )
  .transform((value) => normalizePresentationTransitionSound(value)!);

function placeholderGeometryWithoutFlips(value: unknown): unknown {
  if (!value || typeof value !== "object" || Array.isArray(value)) return value ?? {};
  const rest = { ...(value as Record<string, unknown>) };
  delete rest.flipH;
  delete rest.flipV;
  return rest;
}

export const presentationSlideSchema = z
  .object({
    id: z.string().min(1),
    title: z.string(),
    widthEmu: z
      .number()
      .positive()
      .max(PRESENTATION_GEOMETRY_EMU_LIMIT)
      .optional(),
    heightEmu: z
      .number()
      .positive()
      .max(PRESENTATION_GEOMETRY_EMU_LIMIT)
      .optional(),
    // A bound on work only; what a slide may really carry is its share of the
    // page's bytes (`lib/presentation/slide-block-budget`), and the deck's own
    // whole-tree envelope has the last word. 200 used to sit here and cut real
    // slides — see the module's comment.
    blocks: z.array(presentationBlockSchema).max(MAX_PRESENTATION_SLIDE_BLOCKS),
    notes: z.string().optional(),
    backgroundColor: z
      .string()
      .regex(/^#[0-9a-fA-F]{6}$/)
      .optional(),
    backgroundTransparency: z.number().int().min(0).max(100).optional(),
    // Optional RICH background fill (gradient / picture / pattern), all additive.
    // A slide that carries one of these overrides its (and the theme's) solid
    // `backgroundColor` at render time; a slide with none renders byte-for-byte
    // as before. See `lib/presentation-slide-bg` for the resolution precedence.
    backgroundGradient: presentationBgGradientSchema.optional(),
    backgroundImage: presentationBgImageSchema.optional(),
    backgroundPattern: presentationBgPatternSchema.optional(),
    // PowerPoint Design → Hide Background Graphics. `true` maps to
    // `p:sld@showMasterSp="0"`; omission keeps the OOXML default (shown).
    hideBackgroundGraphics: z.boolean().optional(),
    // Optional HIDE / SKIP flag (PowerPoint / WPS "Hide Slide" parity). When
    // `true` the slide is kept in the deck (still edits + serializes) but is
    // SKIPPED during a slideshow. Additive — a visible slide omits it and
    // validates byte-for-byte; the editor write path drops a `false` flag to
    // absent so the serialized form stays minimal (see `normalizeHiddenFlag`).
    hidden: z.boolean().optional(),
    // Immutable native custody stamped by PPTX import. Legacy imported slides
    // without a transition receipt remain readable through passthrough, but the
    // live command adapter refuses to manufacture patchback authority for them.
    source: presentationSlideSourceSchema.optional(),
    // Optional present-mode entrance transition. Constrained to the shared
    // allowed set so an existing deck without this field still validates and a
    // stale value can never inject an unsupported animation.
    transition: z.enum(SLIDE_TRANSITIONS).optional(),
    // Optional per-slide transition DURATION (ms). Purely additive — a slide
    // without it renders with the CSS default duration (byte-for-byte as before)
    // and existing decks validate unchanged. Bounded to the shared
    // [MIN..MAX]_TRANSITION_DURATION_MS so a hand-edited value can never inject
    // a negative/absurd duration; the render edge re-clamps via
    // `clampTransitionDuration` (lib/presentation-transitions).
    transitionDurationMs: z
      .number()
      .int()
      .min(MIN_TRANSITION_DURATION_MS)
      .max(MAX_TRANSITION_DURATION_MS)
      .optional(),
    // Legacy field name; stores exact Push/Wipe directions, Shape geometry,
    // Zoom in/out, Split orientation/direction, or native orientation identity
    // (Random Bars, Checkerboard, Blinds, and Comb).
    transitionDirection: z.enum(SLIDE_TRANSITION_EFFECT_OPTIONS).optional(),
    // Which element spells a native Zoom. PowerPoint's own Zoom gallery entry
    // writes Office 2010's `p14:warp`, which carries the same In / Out
    // options as the 2006 `p:zoom`; a Zoom this editor authors new names
    // `p14:warp` too, and a slide keeps the element it came in with (absent
    // means `p:zoom`) so an edit never swaps one for the other.
    transitionZoomElement: z
      .literal(SLIDE_TRANSITION_ZOOM_WARP_ELEMENT)
      .optional(),
    // `advClick` omission is the OOXML default true; persist false only.
    transitionAdvanceOnClick: z.boolean().optional(),
    // `advTm` is an unsigned 32-bit millisecond value; omission disables After.
    transitionAdvanceAfterMs: z
      .number()
      .int()
      .min(MIN_TRANSITION_ADVANCE_AFTER_MS)
      .max(MAX_TRANSITION_ADVANCE_AFTER_MS)
      .optional(),
    transitionSound: presentationTransitionSoundSchema.optional(),
    // Read only: PowerPoint's rehearsed click timings (the slide's tag
    // TIMING), ms each after the click before; at most 1,000 clicks, each at
    // most `advTm`'s own bound.
    rehearsedClickTimingsMs: z
      .array(z.number().int().min(0).max(MAX_TRANSITION_ADVANCE_AFTER_MS))
      .max(1000)
      .optional(),
    // Optional master-slide layout id (Title / Title+Content / Two-Content /
    // Section Header / Blank), or a custom layout Slide Master view's Insert
    // Layout made on an Office theme deck (`custom-<letters>`, PPT-B030).
    // Purely additive — existing decks omit it and still validate; a stale
    // value is normalized to `blank` at render time.
    // A layout of an inserted master is `<key>--<layout>` (PPT-B030).
    layout: z
      .union([
        z.enum(LAYOUT_IDS),
        z.string().regex(OFFICE_CUSTOM_LAYOUT_ID),
        z.string().regex(/^master-[a-z]{1,2}--[a-z][a-z-]{0,40}$/),
      ])
      .optional(),
    // Optional reference to a deck-level master (see `presentationMasterSchema`).
    // A slide inherits the master's layout + style overrides as defaults. Purely
    // additive — existing decks omit it; a dangling id degrades to no master at
    // render time (lib/presentation-master).
    layoutId: z.string().min(1).max(512).optional(),
    masterId: z.string().min(1).max(512).optional(),
    // Optional present-mode ANIMATION TIMELINE: an ORDERED sequence of animation
    // steps (see `presentationAnimationStepSchema`). Purely additive — existing
    // decks omit it and still validate. Bounded length; the render-time
    // normalizer (lib/presentation-anim-timeline) drops/clamps invalid steps.
    animationSequence: z
      .array(presentationAnimationStepSchema)
      .max(MAX_TIMELINE_STEPS)
      .optional(),
    comments: z.array(presentationCommentSchema).max(200).optional(),
    commentAuthority: presentationCommentAuthoritySchema.optional(),
  })
  .passthrough()
  .superRefine((slide, context) => {
    for (let blockIndex = 0; blockIndex < slide.blocks.length; blockIndex += 1) {
      const block = slide.blocks[blockIndex] as Record<string, unknown>;
      const source = block.source;
      if (!source || typeof source !== "object" || Array.isArray(source)) continue;
      const sourceRecord = source as Record<string, unknown>;
      if (!Object.hasOwn(sourceRecord, "nativeGroupAuthority")) continue;
      // A stored page may name a group's custody, held once in
      // `deckReceipts` (`presentation-deck-receipts`); the loader writes it
      // out again, and it is checked in that form.
      if (presentationDeckReceiptTokenKey(sourceRecord.nativeGroupAuthority) !== null) continue;
      const authority = parsePresentationNativeGroupAuthority(
        sourceRecord.nativeGroupAuthority,
      );
      if (
        authority === null ||
        typeof sourceRecord.shapeId !== "string" ||
        !authority.members.some(
          (member) => member.shapeId === sourceRecord.shapeId,
        )
      ) {
        context.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["blocks", blockIndex, "source", "nativeGroupAuthority"],
          message: "native group authority is malformed or mismatched",
        });
      }
    }
    const reuseReceipt = parseNativeSlideReuseReceipt(
      slide.source?.nativeSlideReuse,
    );
    if (reuseReceipt && reuseReceipt.modelSlideId !== slide.id) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["source", "nativeSlideReuse", "modelSlideId"],
        message: "native reuse receipt does not address this model slide",
      });
    }
    const layoutCascade = parseNativeLayoutCascadeReceipt(slide.source?.layoutCascade);
    const selectedLayout = layoutCascade?.layouts.find(
      (entry) => entry.layoutPart === slide.layoutId,
    );
    if (
      layoutCascade &&
      (!selectedLayout || slide.masterId !== selectedLayout.masterPart)
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["layoutId"],
        message: "slide layout selection does not match native custody",
      });
    }
    if (layoutCascade) {
      const placeholders = new Map(
        layoutCascade.placeholders.map((placeholder) => [
          placeholder.shapeId,
          placeholder,
        ]),
      );
      slide.blocks.forEach((block, blockIndex) => {
        const blockRecord = block as unknown as Record<string, unknown>;
        const source =
          blockRecord.source &&
          typeof blockRecord.source === "object" &&
          !Array.isArray(blockRecord.source)
            ? (blockRecord.source as Record<string, unknown>)
            : null;
        if (!source || !("placeholderType" in source)) return;
        const placeholder =
          typeof source.shapeId === "string"
            ? placeholders.get(source.shapeId)
            : undefined;
        if (
          !placeholder ||
          source.placeholderType !== placeholder.placeholderType ||
          source.placeholderIndex !== placeholder.placeholderIndex ||
          source.placeholderNativeIndexPresent !==
            placeholder.nativeIndexPresent ||
          source.layoutPart !== layoutCascade.layoutPart ||
          source.masterPart !== layoutCascade.masterPart ||
          // The same receipt, whatever order its keys arrive in: a page read
          // back from storage and a plan the server just built order them
          // differently, and Duplicate Slide's plan was refused for it
          
          // A placeholder's flips (`flipH`/`flipV`) are no part of the box its
          // layout places, and the receipt reads none (`nativeLayoutGeometry`):
          // a flipped placeholder's own receipt refused the whole page
          
          !pageDocumentJsonEquals(
            placeholderGeometryWithoutFlips(source.layoutGeometry),
            placeholder.layoutGeometry ?? {},
          ) ||
          !pageDocumentJsonEquals(source.layoutStyle, placeholder.layoutStyle)
        ) {
          context.addIssue({
            code: z.ZodIssueCode.custom,
            path: ["blocks", blockIndex, "source"],
            message: "placeholder source does not match native layout custody",
          });
        }
      });
    }
    if (
      slide.backgroundTransparency !== undefined &&
      slide.backgroundColor === undefined
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["backgroundTransparency"],
        message: "solid background transparency requires a solid colour",
      });
    }
    if (
      slide.transitionDirection !== undefined &&
      !isSlideTransitionDirectionForKind(
        slide.transition,
        slide.transitionDirection,
      )
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["transitionDirection"],
        message: "transition Effect Options must match the native effect kind",
      });
    }
    if (
      slide.transitionZoomElement !== undefined &&
      slide.transition !== "zoom"
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["transitionZoomElement"],
        message: "only a Zoom transition names the element it is written with",
      });
    }
    const editableModern = slide.commentAuthority?.mode === "modern-editable";
    if (editableModern && slide.comments) {
      for (const [commentIndex, comment] of slide.comments.entries()) {
        const root = presentationModernCommentRecordSchema.safeParse(comment);
        if (!root.success) {
          context.addIssue({
            code: z.ZodIssueCode.custom,
            path: ["comments", commentIndex],
            message: "editable modern comment identity is malformed",
          });
        }
        for (const [replyIndex, reply] of (comment.replies ?? []).entries()) {
          const parsed = presentationModernCommentRecordSchema.safeParse(reply);
          if (!parsed.success) {
            context.addIssue({
              code: z.ZodIssueCode.custom,
              path: ["comments", commentIndex, "replies", replyIndex],
              message: "editable modern reply identity is malformed",
            });
          }
        }
      }
    }
  });

// Optional deck-level master/theme: the DEFAULT background, font family, and
// title / body text colours applied to every slide + block that does not supply
// its own override. Purely additive — existing decks omit it and still validate.
// Colours are strict six-digit hex (mirroring the slide background); the font is
// constrained to the shared whitelist so a hand-edited value can never inject an
// arbitrary `font-family`. The render-time resolver (lib/presentation-theme)
// re-normalizes each field, so this schema is a structural guard, not the
// render-time invariant.
const presentationThemeTokenSnapshotSchema = z
  .object({
    background: z
      .string()
      .regex(/^#[0-9a-fA-F]{6}$/)
      .optional(),
    backgroundGradient: presentationBgGradientSchema.optional(),
    backgroundPattern: presentationBgPatternSchema.optional(),
    accentColors: z
      .array(z.string().regex(/^#[0-9a-fA-F]{6}$/))
      .max(6)
      .optional(),
    fontFamily: z.enum(THEME_FONT_FAMILIES).optional(),
    titleColor: z
      .string()
      .regex(/^#[0-9a-fA-F]{6}$/)
      .optional(),
    bodyTextColor: z
      .string()
      .regex(/^#[0-9a-fA-F]{6}$/)
      .optional(),
    titleFontSize: z.number().min(6).max(400).optional(),
    bodyFontSize: z.number().min(6).max(400).optional(),
    titleBold: z.literal(true).optional(),
    bodyBold: z.literal(true).optional(),
    titleItalic: z.literal(true).optional(),
    bodyItalic: z.literal(true).optional(),
  })
  .strict();

const presentationThemeIdentitySchema = z
  .object({
    schemaVersion: z.literal(PRESENTATION_THEME_IDENTITY_SCHEMA_VERSION),
    selectionKind: z.enum(["catalog", "office", "native", "custom"]),
    themeId: z
      .string()
      .min(1)
      .max(160)
      .regex(/^[a-z0-9][a-z0-9._:-]*$/i),
    variantId: z.enum(PRESENTATION_THEME_VARIANT_IDS),
    source: z
      .object({
        part: z
          .string()
          .refine(isNativePresentationThemePartPath),
        fingerprint: z.string().regex(/^[0-9a-f]{64}$/),
        themeCount: z.number().int().min(1).max(64).optional(),
        name: z.string().min(1).max(200).optional(),
        colorSlots: z
          .record(z.string(), z.string().regex(/^#[0-9a-fA-F]{6}$/))
          .refine((value) =>
            Object.keys(value).every((key) =>
              (PRESENTATION_THEME_COLOR_SLOTS as readonly string[]).includes(
                key,
              ),
            ),
          )
          .optional(),
        majorLatinFont: z.string().min(1).max(200).optional(),
        minorLatinFont: z.string().min(1).max(200).optional(),
        masterDesignFingerprint: z
          .string()
          .regex(/^[0-9a-f]{64}$/)
          .optional(),
        // Every master theme of a deck of several (row 231).
        themeParts: z
          .array(
            z
              .object({
                part: z.string().refine(isNativePresentationThemePartPath),
                fingerprint: z.string().regex(/^[0-9a-f]{64}$/),
                masterDesignFingerprint: z.string().regex(/^[0-9a-f]{64}$/),
              })
              .strict(),
          )
          .min(2)
          .max(64)
          .optional(),
        baseTheme: presentationThemeTokenSnapshotSchema.optional(),
        // The Effects-gallery entry the theme's own scheme is (row 851).
        effectSchemeMatch: z
          .string()
          .min(1)
          .max(80)
          .regex(/^[a-z0-9][a-z0-9-]*$/)
          .optional(),
      })
      .strict()
      .optional(),
    // The theme came with the file (lib/presentation-theme). Every deck whose
    // own theme is an Office one holds it, and without it here the editor
    // buttons could not open such a deck at all (AI70H h50 C12).
    origin: z.literal("package").optional(),
  })
  .strict()
  .superRefine((value, context) => {
    if (value.selectionKind === "native" && !value.source) {
      context.addIssue({
        code: "custom",
        message: "native presentation theme identity requires source authority",
        path: ["source"],
      });
    }
    // An Office identity names a catalog theme and one of its own variants
    // (v1..v8); the MRAI kinds keep base / soft / deep / contrast.
    if (
      value.selectionKind === "office"
        ? !presentationOfficeThemeVariant(value.themeId, value.variantId)
        : !isThemeDesignVariantId(value.variantId)
    ) {
      context.addIssue({
        code: "custom",
        message: "presentation theme variant does not belong to the selected theme",
        path: ["variantId"],
      });
    }
  });

const presentationThemeHexSchema = z.string().regex(/^#[0-9a-fA-F]{6}$/);
const presentationThemeFaceSchema = z
  .string()
  .min(1)
  .max(200)
  .regex(/^[^<>&\u0000-\u001f\u007f]+$/);

const presentationOfficeMasterFrameSchema = z.tuple([
  z.number().int().min(-51_206_400).max(51_206_400),
  z.number().int().min(-51_206_400).max(51_206_400),
  z.number().int().min(1).max(51_206_400),
  z.number().int().min(1).max(51_206_400),
]);

const presentationOfficeMasterTextValuesSchema = z
  .object({
    fontSize: z.number().min(1).max(4000).optional(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    color: presentationThemeHexSchema.optional(),
  })
  .strict();

// Slide Master view's edits of one master of an Office theme deck (PPT-B030,
// lib/presentation-office-master-model): master one's, and each inserted master's.
const presentationOfficeMasterEditsShape = {
  textStyles: z
    .object({
      title: z.object({ "1": presentationOfficeMasterTextValuesSchema.optional() }).strict().optional(),
      body: z
        .object(
          Object.fromEntries(
            OFFICE_MASTER_BODY_LEVELS.map((level) => [
              level,
              presentationOfficeMasterTextValuesSchema.optional(),
            ]),
          ) as Record<
            (typeof OFFICE_MASTER_BODY_LEVELS)[number],
            z.ZodOptional<typeof presentationOfficeMasterTextValuesSchema>
          >,
        )
        .strict()
        .optional(),
    })
    .strict()
    .optional(),
  background: z.literal(true).optional(),
  backgroundStyle: z.number().int().min(1).max(12).optional(),
  masterFrames: z
    .object({
      title: presentationOfficeMasterFrameSchema.optional(),
      body: presentationOfficeMasterFrameSchema.optional(),
    })
    .strict()
    .optional(),
  layoutFrames: z
    .record(
      z.string().regex(/^[a-z][a-z-]{0,40}$/),
      z.record(z.string().regex(/^\d{1,2}$/), presentationOfficeMasterFrameSchema),
    )
    .optional(),
  name: z.string().min(1).max(200).optional(),
  layoutNames: z.record(z.string().regex(/^[a-z][a-z-]{0,40}$/), z.string().min(1).max(200)).optional(),
  preserve: z.literal(true).optional(),
  // Master Layout > Insert Placeholder / Title / Footers and Edit
  // Master > Delete on a layout (PPT-B030).
  addedPlaceholders: z
    .record(
      z.string().regex(/^[a-z][a-z-]{0,40}$/),
      z
        .array(
          z
            .object({
              kind: z.enum(OFFICE_PLACEHOLDER_KINDS),
              frame: presentationOfficeMasterFrameSchema,
            })
            .strict(),
        )
        .max(OFFICE_MASTER_MAX_ADDED_PLACEHOLDERS),
    )
    .optional(),
  layoutSlots: z
    .record(
      z.string().regex(/^[a-z][a-z-]{0,40}$/),
      z.array(z.string().regex(/^[ca]\d{1,2}$/)).max(32),
    )
    .optional(),
  deletedLayouts: z.array(z.string().regex(/^[a-z][a-z-]{0,40}$/)).max(64).optional(),
  layoutsWithoutFooters: z.array(z.string().regex(/^[a-z][a-z-]{0,40}$/)).max(64).optional(),
  // Background > Hide Background Graphics on a layout, Master Layout's
  // deleted master placeholders and Edit Master > Insert Layout (PPT-B030).
  layoutsHidingBackgroundGraphics: z.array(z.string().regex(/^[a-z][a-z-]{0,40}$/)).max(64).optional(),
  masterPlaceholdersOff: z.array(z.enum(OFFICE_MASTER_PLACEHOLDER_FAMILIES)).max(5).optional(),
  customLayouts: z.array(z.string().regex(OFFICE_CUSTOM_LAYOUT_ID)).max(OFFICE_MASTER_MAX_CUSTOM_LAYOUTS).optional(),
  layoutOrder: z.array(z.string().regex(/^[a-z][a-z-]{0,40}$/)).max(64).optional(),
};

export const presentationThemeSchema = z
  .object({
    identity: presentationThemeIdentitySchema.optional(),
    background: z
      .string()
      .regex(/^#[0-9a-fA-F]{6}$/)
      .optional(),
    // Optional deck-level RICH background fill (gradient / picture / pattern),
    // the DEFAULT for every slide that does not supply its own. Additive — a
    // theme without these is unchanged. Same shape + resolver as the per-slide
    // fields (lib/presentation-slide-bg).
    backgroundGradient: presentationBgGradientSchema.optional(),
    backgroundImage: presentationBgImageSchema.optional(),
    backgroundPattern: presentationBgPatternSchema.optional(),
    fontFamily: z.enum(THEME_FONT_FAMILIES).optional(),
    // Design > Variants > Fonts: one entry of Office's Fonts gallery (§75.1).
    fontSchemeId: z
      .string()
      .refine((value) => officeThemeFontGalleryEntry(value) !== undefined)
      .optional(),
    // Design > Variants > Colors: one entry of Office's Colors gallery (PPT-A073).
    colorSchemeId: z
      .string()
      .refine((value) => officeColorGalleryScheme(value) !== undefined)
      .optional(),
    // Design > Variants > Effects: one entry of Office's Effects gallery (PPT-A073).
    effectSchemeId: z
      .string()
      .refine((value) => officeEffectGalleryScheme(value) !== undefined)
      .optional(),
    // An Office theme's twelve colour slots and its font scheme by name
    // (lib/presentation-theme `colorSlots` / `fontScheme`).
    colorSlots: z
      .object(
        Object.fromEntries(
          PRESENTATION_THEME_COLOR_SLOTS.map((slot) => [
            slot,
            presentationThemeHexSchema,
          ]),
        ) as Record<
          (typeof PRESENTATION_THEME_COLOR_SLOTS)[number],
          typeof presentationThemeHexSchema
        >,
      )
      .strict()
      .optional(),
    // The first master's colour map (reading row 826, lib/presentation-theme
    // `colorMap`): bg1 / tx1 / bg2 / tx2 each to one of the twelve slots.
    colorMap: z
      .object({
        bg1: z.enum(PRESENTATION_THEME_COLOR_SLOTS),
        tx1: z.enum(PRESENTATION_THEME_COLOR_SLOTS),
        bg2: z.enum(PRESENTATION_THEME_COLOR_SLOTS),
        tx2: z.enum(PRESENTATION_THEME_COLOR_SLOTS),
      })
      .strict()
      .optional(),
    fontScheme: z
      .object({
        majorLatin: presentationThemeFaceSchema,
        minorLatin: presentationThemeFaceSchema,
        majorHans: presentationThemeFaceSchema.optional(),
        minorHans: presentationThemeFaceSchema.optional(),
        // Create New Theme Fonts' Name (`a:fontScheme@name`).
        name: z
          .string()
          .min(1)
          .max(80)
          .regex(/^[^\u0000-\u001f\u007f]+$/)
          .optional(),
      })
      .strict()
      .optional(),
    titleColor: z
      .string()
      .regex(/^#[0-9a-fA-F]{6}$/)
      .optional(),
    bodyTextColor: z
      .string()
      .regex(/^#[0-9a-fA-F]{6}$/)
      .optional(),
    // Slide Master view's edits of an Office theme deck's master (PPT-B030,
    // lib/presentation-office-master-model).
    officeMaster: z
      .object({
        ...presentationOfficeMasterEditsShape,
        // Edit Master > Insert Slide Master's masters (PPT-B030): each the
        // edits of one master, without a background of its own.
        additionalMasters: z
          .array(
            // An inserted master's background is its theme's own.
            z
              .object(presentationOfficeMasterEditsShape)
              .omit({ background: true, backgroundStyle: true })
              .extend({ key: z.string().regex(OFFICE_ADDITIONAL_MASTER_KEY) })
              .strict(),
          )
          .max(OFFICE_MASTER_MAX_ADDITIONAL)
          .optional(),
      })
      .strict()
      .optional(),
  })
  .passthrough();

// Optional deck-level MASTER layout. A master is a named bundle of an inherited
// layout id + style overrides (the same shape as the deck theme) that slides
// reference via `masterId` and INHERIT as defaults. Purely additive — existing
// decks carry no `masters`; the render-time resolver (lib/presentation-master)
// re-normalizes each field (id/name length, known layout id, valid theme), so
// this schema is a structural guard, not the render-time invariant.
export const presentationMasterSchema = z
  .object({
    id: z.string().min(1).max(200),
    name: z.string().max(200).optional(),
    layoutId: z.enum(LAYOUT_IDS).optional(),
    theme: presentationThemeSchema.optional(),
    // Placeholder blocks edited ON the master that propagate to every slide using
    // it (PowerPoint master text). A slide overrides an inherited placeholder by
    // providing its own block of the same id (see lib/presentation-master-cascade).
    // Optional + bounded — a master with no placeholder text validates unchanged.
    blocks: z.array(presentationBlockSchema).max(MAX_PRESENTATION_SLIDE_BLOCKS).optional(),
  })
  .passthrough();

// Optional deck-level SECTION (PowerPoint / WPS named, collapsible slide group).
// A section is a named anchor on the deck: every slide from `startSlideId`
// (inclusive) up to the next section's anchor belongs to it; slides before the
// first anchor form an implicit default group. Purely additive — existing decks
// carry no `sections` and still validate. This schema is a STRUCTURAL guard
// (id/name length, anchor present); the deeper invariants (anchor references an
// existing slide, no duplicate ids/anchors, order matches slide order) are
// enforced by `normalizeSections` (lib/presentation-sections) at the editor +
// serialize edges, so a hand-edited / forged value can never crash deck load.
const presentationSectionSchema = z
  .object({
    id: z.string().min(1).max(200),
    name: z.string().max(200).optional(),
    startSlideId: z.string().min(1).max(200),
  })
  .passthrough();

// Optional deck-level HEADER & FOOTER config (PowerPoint / WPS "Header & Footer"
// → Insert > Slide Number / Footer / Date). `slideNumber` shows each slide's
// 1-based index (bottom-right); `footerText` shows a fixed line (bottom-left);
// `dateField` shows a date (bottom-center), either `auto` (re-rendered to today)
// or `fixed` (a stored snapshot string). Purely additive — existing decks carry
// no `headerFooter` and still validate. This schema only BOUNDS a present value
// (footer / fixed-date length; the boolean + mode are coerced at the edges);
// `normalizeHeaderFooter` (lib/presentation-header-footer) strips control bytes +
// drops a `false` flag / empty footer / empty fixed date at the serialize edge,
// so a deck that never configured a header/footer persists byte-for-byte
// unchanged (no `headerFooter` key). Each field is independently `.optional()`.
const presentationDateFieldFormatSchema = z.enum([
  "datetime1",
  "datetime2",
  "datetime3",
  "datetime4",
  "datetime5",
  "datetime6",
  "datetime7",
  "datetime8",
  "datetime9",
  "datetime10",
  "datetime11",
  "datetime12",
  "datetime13",
]);
const presentationDateFieldSchema = z.discriminatedUnion("mode", [
  z
    .object({
      mode: z.literal("auto"),
      format: presentationDateFieldFormatSchema.optional(),
      
      locale: z.string().regex(/^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8}){0,3}$/).optional(),
    })
    .strict(),
  z
    .object({
      mode: z.literal("fixed"),
      text: z.string().max(MAX_PRESENTATION_DATE_LENGTH),
    })
    .strict(),
]);
const presentationHeaderFooterSlideOverrideSchema = z
  .object({
    slideNumber: z.boolean().optional(),
    footerText: z
      .string()
      .max(MAX_PRESENTATION_FOOTER_LENGTH)
      .nullable()
      .optional(),
    dateField: presentationDateFieldSchema.nullable().optional(),
  })
  .strict();
const presentationHeaderFooterSchema = z
  .object({
    firstSlideNumber: z
      .number()
      .int()
      .min(MIN_PRESENTATION_FIRST_SLIDE_NUMBER)
      .max(MAX_PRESENTATION_FIRST_SLIDE_NUMBER)
      .optional(),
    notesSize: z
      .object({
        widthEmu: z.number().int().positive().max(MAX_PRESENTATION_NOTES_SIZE_EMU),
        heightEmu: z.number().int().positive().max(MAX_PRESENTATION_NOTES_SIZE_EMU),
      })
      .strict()
      .optional(),
    slideNumber: z.boolean().optional(),
    footerText: z.string().max(MAX_PRESENTATION_FOOTER_LENGTH).optional(),
    dateField: presentationDateFieldSchema.optional(),
    excludeTitleSlide: z.literal(true).optional(),
    // Header and Footer > Notes and Handouts (PPT-B006).
    notesHandouts: z
      .object({
        pageNumber: z.literal(false).optional(),
        header: z.string().min(1).max(MAX_PRESENTATION_FOOTER_LENGTH).optional(),
        footer: z.string().min(1).max(MAX_PRESENTATION_FOOTER_LENGTH).optional(),
        dateField: presentationDateFieldSchema.optional(),
      })
      .strict()
      .optional(),
    slideOverrides: z
      .record(presentationHeaderFooterSlideOverrideSchema)
      .refine((overrides) => Object.keys(overrides).length <= 200, {
        message: "Header/footer overrides exceed the slide limit",
      })
      .optional(),
  })
  .strict();

// Deck-level PRINT / HANDOUT settings (PowerPoint "Print ▸ Handouts" parity):
// how many slides are tiled per printed page + a print scale percentage. All
// fields optional; absent means "use the export default" (handout-2). Schema
// mirrors the PresentationPrintSettings type in lib/presentation-print-settings.ts.
// `.passthrough()` so unknown future sub-fields survive round-trips.
export const presentationPrintSettingsSchema = z
  .object({
    // Slides tiled per printed page: 1 (one slide/page) / 2 / 3 / 4 / 6 / 9 (handout grid).
    slidesPerPage: z
      .union([z.literal(1), z.literal(2), z.literal(3), z.literal(4), z.literal(6), z.literal(9)])
      .optional(),
    // Print scale percentage 10–400 (100 = unscaled).
    scale: z.number().min(10).max(400).optional(),
    // Print Hidden Slides (PowerPoint's Print dialog): absent = on (its
    // default); an explicit on / off as the user or the file stated it.
    printHiddenSlides: z.boolean().optional(),
  })
  .passthrough();

export const presentationPageContentObjectSchema = z
  .object({
    type: z.literal(PRESENTATION_CONTENT_TYPE),
    fontFaces: z.array(z.custom<OfficeFontFace>((raw) => {
      const face = normalizeOfficeFontFace(raw);
      return Boolean(face?.asset && face.family.toLowerCase() === face.asset.family.toLowerCase() &&
        face.asset.files.every(file => file.attachmentId && !file.embeddedPart && !file.fontKey));
    })).max(128).optional(),
    slides: z.array(presentationSlideSchema).min(1).max(MAX_PRESENTATION_SLIDES),
    // Optional deck-level master/theme (see `presentationThemeSchema`).
    theme: presentationThemeSchema.optional(),
    // Optional deck-level masters (see `presentationMasterSchema`). Additive —
    // existing decks omit it and still validate.
    masters: z.array(presentationMasterSchema).max(50).optional(),
    // Optional deck-level SECTIONS (see `presentationSectionSchema`). Additive —
    // existing decks omit it and still validate. A malformed list is dropped /
    // repaired by `normalizeSections` on load, never crashing the editor.
    sections: z.array(presentationSectionSchema).max(200).optional(),
    // Optional deck-level slide number + footer (see
    // `presentationHeaderFooterSchema`). Additive — existing decks omit it.
    headerFooter: presentationHeaderFooterSchema.optional(),
    // Optional deck-level PRINT / HANDOUT settings (slides per page + scale).
    // Additive — existing decks omit it and still validate byte-for-byte.
    printSettings: presentationPrintSettingsSchema.optional(),
    // Home > Replace > Replace Fonts, in order (`lib/presentation-replace-font`):
    // the log a save of an imported deck replays on its original
    // (`pptx/replace_fonts`). Ids unique; each face one family, never a theme
    // reference.
    fontReplacements: z
      .array(
        z
          .object({
            id: z.string().min(1).max(64),
            from: z.string().trim().min(1).max(120).refine((face) => !/^\+m[nj]-(lt|ea|cs)$/i.test(face)),
            to: z.string().trim().min(1).max(120).refine((face) => !/^\+m[nj]-(lt|ea|cs)$/i.test(face)),
          })
          .strict(),
      )
      .max(200)
      .refine((entries) => new Set(entries.map((entry) => entry.id)).size === entries.length)
      .optional(),
  })
  .passthrough();
export type PresentationPageContentParsed = z.infer<typeof presentationPageContentObjectSchema>;

/** The deck-wide rules of a presentation page's content (ids, comment
 * authority, links between slides…), over its parsed fields. A reader that
 * parses the deck's slides one by one (`normalizeSlides`) applies them to the
 * whole deck so read, exactly as `presentationPageContentSchema` does. */
export function refinePresentationPageContent(content: PresentationPageContentParsed, ctx: z.RefinementCtx): void {
    if (content.fontFaces) {
      const faces = content.fontFaces.map(normalizeOfficeFontFace);
      const files = faces.flatMap(face => face?.asset?.files ?? []);
      if (new Set(faces.map(face => face?.family.toLowerCase())).size !== faces.length ||
        files.length > 128 || files.reduce((size, file) => size + file.sizeBytes, 0) > 128 * 1024 * 1024)
        ctx.addIssue({ code: z.ZodIssueCode.custom, path: ["fontFaces"], message: "Font bindings must be unique and bounded" });
    }
    const slideIds = new Set(content.slides.map((slide) => slide.id));
    if (slideIds.size !== content.slides.length) {
      content.slides.forEach((slide, slideIndex) => {
        if (content.slides.findIndex((candidate) => candidate.id === slide.id) !== slideIndex) {
          ctx.addIssue({
            code: z.ZodIssueCode.custom,
            path: ["slides", slideIndex, "id"],
            message: "Slide ids must be unique across the deck",
          });
        }
      });
    }
    const authorities = content.slides
      .map((slide) => slide.commentAuthority)
      .filter((value): value is z.infer<typeof presentationCommentAuthoritySchema> =>
        value !== undefined,
      );
    if (authorities.length !== 0 && authorities.length !== content.slides.length) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["slides"],
        message: "Comment authority must be present identically on every slide",
      });
    }
    const firstAuthority = authorities[0] ? JSON.stringify(authorities[0]) : null;
    // One receipt, whatever order its keys arrive in (a plan the server just
    // built orders them as the projection does).
    if (
      firstAuthority !== null &&
      authorities.some((authority) => !pageDocumentJsonEquals(authority, authorities[0]))
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["slides"],
        message: "Comment authority receipts disagree",
      });
    }
    const commentIds = new Set<string>();
    const enforceDocumentCommentIdentity = firstAuthority !== null;
    content.slides.forEach((slide, slideIndex) => {
      for (const [commentIndex, comment] of (slide.comments ?? []).entries()) {
        for (const [id, path] of [
          [comment.id, ["slides", slideIndex, "comments", commentIndex, "id"]],
          ...(comment.replies ?? []).map((reply, replyIndex) => [
            reply.id,
            ["slides", slideIndex, "comments", commentIndex, "replies", replyIndex, "id"],
          ] as const),
        ] as const) {
          if (enforceDocumentCommentIdentity && commentIds.has(id)) {
            ctx.addIssue({
              code: z.ZodIssueCode.custom,
              path: [...path],
              message: "Comment and reply ids must be unique across the deck",
            });
          }
          if (enforceDocumentCommentIdentity) commentIds.add(id);
        }
      }
    });
    for (const slideId of Object.keys(
      content.headerFooter?.slideOverrides ?? {},
    )) {
      if (!slideIds.has(slideId)) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["headerFooter", "slideOverrides", slideId],
          message: "Header/footer override must target an existing slide",
        });
      }
    }
    content.slides.forEach((slide, slideIndex) => {
      const blocks = new Map(slide.blocks.map((block) => [block.id, block]));
      slide.blocks.forEach((block, blockIndex) => {
        if (block.type === "text") {
          const runGroups = [
            { runs: block.runs, path: ["runs"] as const },
            ...(block.paragraphs ?? []).map((paragraph, paragraphIndex) => ({
              runs: paragraph.runs,
              path: ["paragraphs", paragraphIndex, "runs"] as const,
            })),
          ];
          for (const group of runGroups) {
            for (const [runIndex, run] of (group.runs ?? []).entries()) {
              const href = run.style?.href;
              if (
                href &&
                parseSlideLinkTarget(href) !== null &&
                parseSlideLinkTarget(href, slideIds) === null
              ) {
                ctx.addIssue({
                  code: z.ZodIssueCode.custom,
                  path: [
                    "slides",
                    slideIndex,
                    "blocks",
                    blockIndex,
                    ...group.path,
                    runIndex,
                    "style",
                    "href",
                  ],
                  message: "Internal run link must target an existing slide",
                });
              }
            }
          }
        }
        if (block.type !== "shape" || block.clickAction?.kind !== "play-media") {
          return;
        }
        const target = blocks.get(block.clickAction.targetBlockId);
        if (
          !target ||
          target.type !== "embed" ||
          target.kind !== block.clickAction.mediaKind ||
          !target.url
        ) {
          ctx.addIssue({
            code: z.ZodIssueCode.custom,
            path: ["slides", slideIndex, "blocks", blockIndex, "clickAction"],
            message: "Media Action Buttons require a playable same-slide media block",
          });
        }
      });
    });
}

export const presentationPageContentSchema = presentationPageContentObjectSchema.superRefine(refinePresentationPageContent);

// One reply in an annotation comment THREAD (Acrobat / WPS PDF "Reply"). Additive
// + optional — an annotation without a thread validates byte-for-byte as before.
export const pdfAnnotationReplySchema = z
  .object({
    id: z.string().min(1).max(120),
    text: z.string().max(2_000),
    createdAt: z.string().optional(),
    created_at: z.string().optional(),
  })
  .passthrough();

export const pdfAnnotationSchema = z
  .object({
    id: z.string().min(1),
    text: z.string(),
    quote: z.string().optional(),
    comment: z.string().optional(),
    summary: z.string().optional(),
    createdAt: z.string().optional(),
    created_at: z.string().optional(),
    replies: z.array(pdfAnnotationReplySchema).max(500).optional(),
  })
  .passthrough();

export const pdfFormFieldSchema = z
  .object({
    id: z.string().min(1),
    value: z.union([
      z.string().max(32_000),
      z.boolean(),
      z.array(z.string().max(ACROFORM_MAX_TEXT)).max(ACROFORM_MAX_OPTIONS),
    ]),
    label: z.string().max(500).optional(),
    page: z.number().int().min(1).optional(),
    updatedAt: z.string().optional(),
    updated_at: z.string().optional(),
  })
  .passthrough();

export const pdfPageTransformSchema = z
  .object({
    page: z.number().int().min(1),
    rotation: z
      .number()
      .int()
      .refine((value) => [0, 90, 180, 270].includes(value), {
        message: "rotation must be 0, 90, 180, or 270",
      }),
    updatedAt: z.string().optional(),
    updated_at: z.string().optional(),
  })
  .passthrough();

export const pdfDeletedPageSchema = z.number().int().min(1).max(10_000);

export const pdfInsertedPageSchema = z
  .object({
    id: z.string().min(1),
    afterPage: z.number().int().min(0).max(10_000),
    after_page: z.number().int().min(0).max(10_000).optional(),
    width: z.number().positive().max(5_000).optional(),
    height: z.number().positive().max(5_000).optional(),
    // When set, this inserted slot is a DUPLICATE of an original 1-based page
    // (not a blank). The flatten export copies that source page's content here;
    // the on-screen viewer renders the source page in the slot. Additive +
    // optional — a plain blank insert simply omits it.
    copyOfPage: z.number().int().min(1).max(10_000).optional(),
    copy_of_page: z.number().int().min(1).max(10_000).optional(),
    createdAt: z.string().optional(),
    created_at: z.string().optional(),
  })
  .passthrough();

export const pdfPageMoveSchema = z
  .object({
    page: z.number().int().min(1).max(10_000),
    afterPage: z.number().int().min(0).max(10_000).optional(),
    after_page: z.number().int().min(0).max(10_000).optional(),
    updatedAt: z.string().optional(),
    updated_at: z.string().optional(),
  })
  .passthrough()
  .superRefine((move, ctx) => {
    if (move.afterPage === undefined && move.after_page === undefined) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "afterPage or after_page is required",
        path: ["afterPage"],
      });
    }
  });

const pdfTextStyleRunSchema = z
  .object({
    text: z.string().min(1).max(20_000),
    fontSize: z.number().positive().max(1).optional(),
    fontFamily: z.string().min(1).max(400).optional(),
    sourceFontName: z.string().min(1).max(200).optional(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    color: z
      .string()
      .regex(/^#[0-9a-fA-F]{6}$/)
      .optional(),
  })
  .strict();

const pdfTextLineLayoutSchema = z
  .object({
    start: z.number().int().min(0).max(20_000),
    end: z.number().int().min(0).max(20_000),
    x: z.number().min(0).max(10),
    y: z.number().min(0).max(10),
    width: z.number().positive().max(10),
    height: z.number().positive().max(10),
  })
  .strict();

export const pdfTextBoxEditSchema = z
  .object({
    id: z.string().min(1).max(120),
    page: z.number().int().min(0).max(100_000),
    // Page-relative fractions (0..1) so the overlay is resolution-independent.
    x: z.number(),
    y: z.number(),
    width: z.number(),
    height: z.number(),
    angle: z.number().finite().optional(),
    writingMode: z.literal("vertical-rl").optional(),
    text: z.string().max(20_000),
    kind: z.enum(["text", "formula"]).optional(),
    latex: z.string().max(20_000).optional(),
    fontSize: z.number().optional(),
    color: z
      .string()
      .regex(/^#[0-9a-fA-F]{6}$/)
      .optional(),
    backgroundColor: z
      .string()
      .regex(/^#[0-9a-fA-F]{6}$/)
      .optional(),
    align: z.enum(["left", "center", "right"]).optional(),
    // Original run's font, carried so a flatten-export bakes the edited text in
    // a matching face (serif/sans/mono). Additive + optional (older edits omit).
    fontFamily: z.string().max(400).optional(),
    sourceFontName: z.string().max(200).optional(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    underline: z.boolean().optional(),
    runs: z.array(pdfTextStyleRunSchema).max(5_000).optional(),
    lineLayout: z.array(pdfTextLineLayoutSchema).max(2_000).optional(),
    // Marks a user-authored "Add Text" box (vs. an edited extracted line). A
    // created box has no underlying glyphs, so the export masks only its own
    // footprint. Additive — extracted-line edits omit it.
    created: z.boolean().optional(),
    // The user set the frame's width by hand: the text wraps inside it and the
    // frame grows downward (WPS). Additive — untouched lines omit it.
    fixedWidth: z.boolean().optional(),
    // The original glyph region (page fractions) the cover/mask stays pinned to
    // after the box is dragged. For a created box this equals the box rect.
    originX: z.number().optional(),
    originY: z.number().optional(),
    originWidth: z.number().optional(),
    originHeight: z.number().optional(),
  })
  .passthrough()
  .superRefine((edit, ctx) => {
    if (edit.runs && edit.runs.map((run) => run.text).join("") !== edit.text) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "runs must concatenate to text",
        path: ["runs"],
      });
    }
    if (
      edit.lineLayout &&
      !normalizePdfTextLineLayout(edit.lineLayout, edit.text)
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "lineLayout ranges must cover text lines exactly",
        path: ["lineLayout"],
      });
    }
  });

export const pdfImageOverlaySchema = z
  .object({
    id: z.string().min(1).max(120),
    page: z.number().int().min(0).max(100_000),
    dataItemId: z.string().min(1).max(200),
    // Page-relative fractions (0..1).
    x: z.number().min(0).max(1),
    y: z.number().min(0).max(1),
    width: z.number().min(Number.EPSILON).max(1),
    height: z.number().min(Number.EPSILON).max(1),
  })
  .passthrough();

export const pdfInkStrokeSchema = z
  .object({
    id: z.string().min(1).max(120),
    page: z.number().int().min(0).max(100_000),
    color: z.string().regex(/^#[0-9a-fA-F]{6}$/),
    width: z.number().min(0).max(100),
    // Flattened [x0,y0,x1,y1,…] page-relative fractions (0..1).
    points: z.array(z.number()).max(20_000),
  })
  .passthrough();

// Per-page CROP (Acrobat "Crop Pages" / WPS 去白边). The visible RECTANGLE OF
// THE PAGE TO KEEP, in page-relative percent [0..100] (top-left origin), baked
// as the page's PDF CropBox on export (everything outside is clipped; no content
// deleted). Additive — older PDFs omit it (no page cropped). Keyed by 1-based
// page; bounds are clamped by the runtime sanitizer (lib/pdf/crop-box.ts).
export const pdfCropBoxSchema = z
  .object({
    page: z.number().int().min(1).max(100_000),
    x: z.number(),
    y: z.number(),
    width: z.number(),
    height: z.number(),
  })
  .passthrough();

// Point-anchored text-comment ("sticky note") annotations. Persisted as
// `pdf_note` data-item annotations at runtime; mirrored here so a page snapshot
// that embeds them still validates. Additive — older PDFs omit the array.
export const pdfStickyNoteSchema = z
  .object({
    id: z.string().min(1).max(120),
    page: z.number().int().min(1).max(100_000),
    // Page-relative position in percent [0..100].
    x: z.number().min(0).max(100),
    y: z.number().min(0).max(100),
    text: z.string().max(2_000),
    // Restricted to the marker allow-list (PDF_NOTE_COLORS) so a crafted
    // snapshot can't inject arbitrary CSS into the inline marker style.
    color: z.string().refine((value) => PDF_NOTE_COLOR_SET.has(value), {
      message: "color must be one of PDF_NOTE_COLORS",
    }),
  })
  .passthrough();

// Freehand ink / drawing annotations. Persisted as `pdf_ink` data-item
// annotations at runtime (apps/web/lib/pdf/ink.ts: a polyline in page-relative
// percent [0..100] + an allow-listed stroke color + clamped width); mirrored
// here so a page snapshot that embeds them still validates. Additive — older
// PDFs omit the array. Distinct from `pdfInkStrokeSchema` above (which stores
// flattened 0..1 fraction points for the page-ops snapshot path); this mirrors
// the `pdf_ink` annotation contract exactly (percent coords + allow-list color).
export const pdfInkAnnotationSchema = z
  .object({
    id: z.string().min(1).max(120),
    page: z.number().int().min(1).max(100_000),
    // Restricted to the stroke allow-list (PDF_INK_COLORS) so a crafted
    // snapshot can't inject arbitrary CSS into the inline SVG stroke style.
    color: z.string().refine((value) => PDF_INK_COLOR_SET.has(value), {
      message: "color must be one of PDF_INK_COLORS",
    }),
    width: z.number().min(0).max(PDF_INK_MAX_WIDTH),
    // Polyline points in page-relative percent [0..100].
    points: z
      .array(
        z.object({
          x: z.number().min(0).max(100),
          y: z.number().min(0).max(100),
        }),
      )
      .min(2)
      .max(PDF_INK_MAX_POINTS),
  })
  .passthrough();

// VISUAL signature overlays (Acrobat "Sign" parity). Persisted as
// `pdf_signatures` data-item annotations at runtime (apps/web/lib/pdf/
// signatures.ts: a page-relative box [0..100] holding EITHER a drawn raster
// image data URL OR a typed line of text in an allow-listed font); mirrored here
// so a page snapshot that embeds them still validates. Additive — older PDFs
// omit the array. This is a VISUAL signature only, NOT a cryptographic/digital
// certificate signature. The `image` is restricted to a base64 PNG/JPEG data URL
// (never SVG) and length-capped, the `text` capped to SIGNATURE_MAX_TEXT and
// rendered as PLAIN TEXT (never innerHTML), and `font` restricted to the
// allow-list — so a crafted snapshot can't inject active content or arbitrary
// CSS, nor smuggle an unbounded blob.
export const pdfSignatureSchema = z
  .object({
    id: z.string().min(1).max(120),
    page: z.number().int().min(1).max(100_000),
    // Box top-left + size in page-relative percent [0..100].
    x: z.number().min(0).max(100),
    y: z.number().min(0).max(100),
    width: z.number().min(0).max(100),
    height: z.number().min(0).max(100),
    objectRotation: z.number().int().min(0).max(359).optional(),
    kind: z.enum(["drawn", "image", "typed"]),
    // Drawn: a base64 PNG/JPEG data URL only (allow-list mirrors the runtime
    // sanitizer); optional/empty for a typed signature.
    image: z
      .string()
      .max(SIGNATURE_MAX_IMAGE_DATA_URL_CHARS)
      .refine(
        (value) => value === "" || sanitizeSignatureImage(value) !== null,
        { message: "image must be a base64 PNG or JPEG data URL" },
      )
      .optional(),
    // Typed: the signature text (rendered as plain text); optional/empty for a
    // drawn signature.
    text: z.string().max(SIGNATURE_MAX_TEXT).optional(),
    // Typed: the font-family key, restricted to the allow-list.
    font: z
      .string()
      .refine((value) => SIGNATURE_FONT_KEY_SET.has(value), {
        message: "font must be one of SIGNATURE_FONT_KEYS",
      })
      .optional(),
  })
  .passthrough()
  .superRefine((sig, ctx) => {
    // A drawn signature must carry a non-empty image; a typed one non-empty text
    // — so a snapshot can't persist an empty (content-less) signature box.
    if ((sig.kind === "drawn" || sig.kind === "image") && !sig.image) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["image"],
        message: "a drawn or uploaded-image signature requires an image",
      });
    }
    if (sig.kind === "typed" && !sig.text) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["text"],
        message: "a typed signature requires text",
      });
    }
  });

// Redaction annotations (Acrobat "Redact" parity). Persisted as `pdf_redactions`
// data-item annotations at runtime (apps/web/lib/pdf/redaction.ts: a page-relative
// box [0..100] marking a region whose content is removed on export). Mirrored
// here so a page snapshot that embeds them still validates. Additive — older
// PDFs omit the array entirely; redaction is optional. A redaction carries NO
// free-form content (no text / image / color to smuggle) — only a bounded box —
// so it is rendered as a fixed opaque-black overlay and baked as a permanent
// opaque-black rectangle on export, with the underlying text omitted from the
// exported text layer. This is the strongest practical (NON-cryptographic)
// redaction: see the lib/pdf/redaction.ts docstring for the documented limit on
// byte-level content-stream removal.
export const pdfRedactionSchema = z
  .object({
    id: z.string().min(1).max(120),
    page: z.number().int().min(1).max(100_000),
    // Box top-left + size in page-relative percent [0..100]. A minimum size is
    // enforced so a degenerate (zero-area) redaction is never persisted.
    x: z.number().min(0).max(100),
    y: z.number().min(0).max(100),
    width: z.number().min(REDACTION_MIN_SIZE).max(100),
    height: z.number().min(REDACTION_MIN_SIZE).max(100),
  })
  .passthrough();

// Interactive AcroForm form-fill values. Persisted at runtime as a SINGLE
// additive `pdf_form_values` data-item annotation (apps/web/lib/pdf/acroform.ts:
// a flat `{ [fieldName]: value }` map where value is a length-capped string for
// text/single-choice fields, a string array for a multi-select list, or a
// boolean for checkboxes / the selected radio export value as a string).
// Mirrored here so a page snapshot that embeds the filled
// form still validates. Additive — older PDFs omit the field entirely. Field
// NAMES are capped to ACROFORM_MAX_FIELD_NAME, string VALUES to
// ACROFORM_MAX_TEXT, and selection arrays to ACROFORM_MAX_OPTIONS so a crafted
// snapshot can't smuggle unbounded data; the map itself is capped to
// ACROFORM_MAX_FIELDS entries. The values are ALWAYS placed
// into form-control values / rendered as plain text (never innerHTML), so this
// is bounding, not escaping.
export const pdfFormValuesSchema = z
  .record(
    z.string().min(1).max(ACROFORM_MAX_FIELD_NAME),
    z.union([
      z.string().max(ACROFORM_MAX_TEXT),
      z.boolean(),
      z.array(z.string().max(ACROFORM_MAX_TEXT)).max(ACROFORM_MAX_OPTIONS),
    ]),
  )
  .refine((map) => Object.keys(map).length <= ACROFORM_MAX_FIELDS, {
    message: `form values map exceeds ${ACROFORM_MAX_FIELDS} fields`,
  });

const pdfPageLabelRangeSchema = z
  .object({
    start: z
      .number()
      .int()
      .min(0)
      .max(PAGE_LABEL_MAX_PAGES - 1),
    style: z.enum([
      "decimal",
      "upper-roman",
      "lower-roman",
      "upper-alpha",
      "lower-alpha",
      "none",
    ]),
    prefix: z
      .string()
      .max(PAGE_LABEL_PREFIX_MAX)
      .refine(isValidPageLabelPrefix, {
        message:
          "page label prefix must contain bounded valid Unicode scalar values",
      })
      .optional(),
    firstValue: z
      .number()
      .int()
      .min(1)
      .max(PAGE_LABEL_FIRST_VALUE_MAX)
      .optional(),
  })
  .strict();

const pdfPageLabelRangesSchema = z
  .array(pdfPageLabelRangeSchema)
  .max(PAGE_LABEL_MAX_PAGES)
  .superRefine((ranges, ctx) => {
    if (ranges.length === 0) return;
    if (ranges[0]?.start !== 0) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: [0, "start"],
        message: "the first page label range must start at page index 0",
      });
    }
    for (let index = 1; index < ranges.length; index += 1) {
      if (ranges[index].start <= ranges[index - 1].start) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: [index, "start"],
          message: "page label ranges must be strictly ordered by start",
        });
      }
    }
  });

export const pdfPageContentSchema = z
  .object({
    type: z.literal(PDF_CONTENT_TYPE),
    dataItemId: z.string().min(1).nullable().optional(),
    data_item_id: z.string().min(1).nullable().optional(),
    annotations: z.array(pdfAnnotationSchema).max(1_000),
    textBoxEdits: z.array(pdfTextBoxEditSchema).max(5_000).optional(),
    text_box_edits: z.array(pdfTextBoxEditSchema).max(5_000).optional(),
    imageOverlays: z.array(pdfImageOverlaySchema).max(2_000).optional(),
    image_overlays: z.array(pdfImageOverlaySchema).max(2_000).optional(),
    inkStrokes: z.array(pdfInkStrokeSchema).max(10_000).optional(),
    ink_strokes: z.array(pdfInkStrokeSchema).max(10_000).optional(),
    formFields: z.array(pdfFormFieldSchema).max(16_384).optional(),
    form_fields: z.array(pdfFormFieldSchema).max(16_384).optional(),
    // Interactive AcroForm fill values keyed by field name (additive). Distinct
    // from `formFields` above (which is the native editor's per-field id+value
    // snapshot); this mirrors the runtime `pdf_form_values` annotation map.
    formValues: pdfFormValuesSchema.optional(),
    form_values: pdfFormValuesSchema.optional(),
    // Canonical native `/PageLabels` authority. Key absence preserves the
    // source tree, [] explicitly removes it, and a non-empty strict array
    // replaces it. No snake-case alias: one field owns this package contract.
    pageLabelRanges: pdfPageLabelRangesSchema.optional(),
    pageTransforms: z.array(pdfPageTransformSchema).max(10_000).optional(),
    page_transforms: z.array(pdfPageTransformSchema).max(10_000).optional(),
    deletedPages: z.array(pdfDeletedPageSchema).max(10_000).optional(),
    deleted_pages: z.array(pdfDeletedPageSchema).max(10_000).optional(),
    insertedPages: z.array(pdfInsertedPageSchema).max(10_000).optional(),
    inserted_pages: z.array(pdfInsertedPageSchema).max(10_000).optional(),
    pageMoves: z.array(pdfPageMoveSchema).max(10_000).optional(),
    page_moves: z.array(pdfPageMoveSchema).max(10_000).optional(),
    cropBoxes: z.array(pdfCropBoxSchema).max(10_000).optional(),
    crop_boxes: z.array(pdfCropBoxSchema).max(10_000).optional(),
    stickyNotes: z.array(pdfStickyNoteSchema).max(5_000).optional(),
    sticky_notes: z.array(pdfStickyNoteSchema).max(5_000).optional(),
    inkAnnotations: z.array(pdfInkAnnotationSchema).max(50_000).optional(),
    ink_annotations: z.array(pdfInkAnnotationSchema).max(50_000).optional(),
    // Visual signature overlays (additive). Mirrors the runtime `pdf_signatures`
    // annotation. Bounded generously; the runtime sanitizer + the backend
    // `_validate_pdf_content` cap counts + sizes.
    signatureAnnotations: z.array(pdfSignatureSchema).max(5_000).optional(),
    signature_annotations: z.array(pdfSignatureSchema).max(5_000).optional(),
    // Redaction boxes (additive). Mirrors the runtime `pdf_redactions`
    // annotation. Bounded generously; the runtime sanitizer + the backend
    // `_validate_pdf_content` cap counts + sizes. Baked as permanent opaque
    // rectangles on export (with underlying text omitted from the text layer).
    redactionAnnotations: z.array(pdfRedactionSchema).max(50_000).optional(),
    redaction_annotations: z.array(pdfRedactionSchema).max(50_000).optional(),
    source: z.record(z.unknown()).optional(),
  })
  .passthrough()
  .superRefine((content, ctx) => {
    if (Object.prototype.hasOwnProperty.call(content, "page_label_ranges")) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["page_label_ranges"],
        message: "use canonical pageLabelRanges",
      });
    }
    if (
      content.dataItemId &&
      content.data_item_id &&
      content.dataItemId !== content.data_item_id
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["dataItemId"],
        message: "dataItemId must match data_item_id",
      });
    }
  });

export const codePageContentSchema = z
  .object({
    type: z.literal(CODE_CONTENT_TYPE),
    language: z.string().max(64).optional(),
    content: z.string().max(1_000_000).optional(),
  })
  .passthrough();

export const nativePageContentSchemaByType = {
  document: documentPageContentSchema,
  spreadsheet: spreadsheetPageContentSchema,
  presentation: presentationPageContentSchema,
  pdf: pdfPageContentSchema,
  code: codePageContentSchema,
} as const;

export type NativeDocumentContent = z.infer<typeof documentPageContentSchema>;
export type NativeSpreadsheetContent = z.infer<
  typeof spreadsheetPageContentSchema
>;
export type NativePresentationContent = z.infer<
  typeof presentationPageContentSchema
>;
export type NativePdfContent = z.infer<typeof pdfPageContentSchema>;
export type NativeCodeContent = z.infer<typeof codePageContentSchema>;
export type NativePageContent =
  | NativeDocumentContent
  | NativeSpreadsheetContent
  | NativePresentationContent
  | NativePdfContent
  | NativeCodeContent;

export function createDefaultPageContent(
  pageType: PageType,
): NativePageContent {
  if (pageType === "spreadsheet") {
    return {
      type: SPREADSHEET_CONTENT_TYPE,
      sheets: [
        {
          id: "sheet-1",
          name: "Sheet1",
          rowCount: DEFAULT_SPREADSHEET_ROW_COUNT,
          columnCount: DEFAULT_SPREADSHEET_COLUMN_COUNT,
          // Excel's 8.43-character default column (XLS-A095).
          defaultColumnWidth: NEW_SHEET_DEFAULT_COLUMN_WIDTH_PX,
          cells: {},
          // A new workbook prints on A4 (print-settings newSheetPrintSettings).
          printSettings: newSheetPrintSettings(),
        },
      ],
    };
  }
  if (pageType === "presentation") {
    return {
      type: PRESENTATION_CONTENT_TYPE,
      slides: [
        {
          id: "slide-1",
          title: "",
          blocks: [{ id: "block-1", type: "text", text: "" }],
        },
      ],
    };
  }
  if (pageType === "pdf") {
    return {
      type: PDF_CONTENT_TYPE,
      dataItemId: null,
      annotations: [],
      pageTransforms: [],
      deletedPages: [],
      insertedPages: [],
      pageMoves: [],
    };
  }
  if (pageType === "code") {
    return {
      type: CODE_CONTENT_TYPE,
      language: "javascript",
      content: "",
    };
  }
  if (pageType === "document") {
    return {
      type: DOCUMENT_CONTENT_TYPE,
      content: [],
    };
  }
  throw new Error(`Unsupported page type: ${String(pageType)}`);
}

export function parseNativePageContent(
  pageType: PageType,
  content: unknown,
): NativePageContent {
  const schema = nativePageContentSchemaByType[pageType];
  if (!schema) {
    throw new Error(`Unsupported page type: ${String(pageType)}`);
  }
  return schema.parse(content) as NativePageContent;
}

export function buildPdfPageContent(
  dataItemId: string,
  source: { filename: string; sizeBytes: number; mediaType?: string | null },
): NativePdfContent {
  return {
    type: PDF_CONTENT_TYPE,
    dataItemId,
    annotations: [],
    pageTransforms: [],
    deletedPages: [],
    insertedPages: [],
    pageMoves: [],
    source: {
      filename: source.filename,
      sizeBytes: source.sizeBytes,
      mediaType: source.mediaType || "application/pdf",
    },
  };
}
