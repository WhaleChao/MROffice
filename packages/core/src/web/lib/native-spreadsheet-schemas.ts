/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeOfficeTextWarp } from "./office-text-warp";
import { SPREADSHEET_CONTENT_TYPE } from "./native-page-content-types";
import { CHART_COLOR_PALETTE_IDS } from "./office-charts/chart-color-palettes.generated";
import { smartArtColorVariant } from "./office-smartart/colors";
import { isAliasNativeSmartArtLayout } from "./office-smartart/native-layouts";
import { PRESENTATION_STORED_BLACK_WHITE_MODES } from "./presentation/presentation-black-white-mode";
import { z } from "zod";
import { SPREADSHEET_MAX_OBJECTS } from "./spreadsheet/object-budget";
import { MAX_PRESENTATION_SLIDE_BLOCKS } from "./presentation/slide-block-budget";
import { MAX_CHART_POINTS_PER_SERIES, MAX_CHART_SERIES } from "./office-charts/chart-limits.generated";
import { presentationDeckReceiptTokenKey } from "./presentation-deck-receipts";
import { nativeRowPixelPairMatches, normalizeNativeGrid } from "./spreadsheet/native-typography";
import { normalizeOfficeTextPaint } from "./office-text-paint";
import { normalizeOfficePictureEffects } from "./office-picture-effects";
import { officePictureStyle } from "./office-picture-styles/catalog";
import { normalizeSpreadsheetChartFormat, type SpreadsheetChartFormat } from "./spreadsheet/chart-format";
import type { SpreadsheetChartBinding } from "./spreadsheet/chart";
import { normalizeChartNativeLook, type ChartNativeLook } from "./office-chart-native-look";
import { parseDataQueryState, type DataQueryState } from "./spreadsheet/data-queries";
import type { ChartBlockAttrsV2 } from "../components/console/editor/extensions/chartBlockShared";
import { isValidInlineImageSource } from "./inline-image-source";
import { isOfficeProvenanceRecord } from "./office-object-provenance";
import { normalizeOfficeShapeGradient, normalizeOfficeShapeShadow, type OfficeShapeGradient, type OfficeShapeShadow } from "./office-shape-paint";
import { isNativeScene3d, isNativeShape3d, isSmartArtScene3d, isSmartArtShape3d, isSmartArtStyleId, isSmartArtTextShape3d, type SmartArtScene3d, type SmartArtShape3d, type SmartArtStyleId, type SmartArtTextShape3d } from "./office-smartart/styles-model";
import { normalizeSignedCropRect } from "./presentation-crop-group";
import { MAX_COLUMN_WIDTH_PX, MAX_DRAWING_LAYER_X_PX, MAX_DRAWING_LAYER_Y_PX, MAX_GRID_COLUMN_SIZE_OVERRIDES, MAX_GRID_COLUMNS, MAX_GRID_ROW_SIZE_OVERRIDES, MAX_GRID_ROWS, MAX_ROW_HEIGHT_PX, MIN_COLUMN_WIDTH_PX, MIN_ROW_HEIGHT_PX } from "./spreadsheet/grid-sizing";
import { ICON_SET_FAMILIES, iconSetCount, MAX_RANGE_COL, MAX_RANGE_ROW, normalizeHexColor, sanitizeHyperlink } from "./spreadsheet/conditional-format";
import { DATA_TABLE_ICON_FILTER_SETS, DATA_TABLE_RELATIVE_DATE_PERIODS, MAX_DATA_TABLES, normalizeDataTableColumnFilter, normalizeDataTableSortState, type DataTableColumnFilter, type DataTableSortState } from "./spreadsheet/data-table";
import { normalizeFilterDateGroups, type FilterDateGroup } from "./spreadsheet/date-filter-tree";
import { parseInternalLink, parseNamedLink } from "./spreadsheet/sheet-interactions";
import { blankCellFormatsArea, blankCellFormatsIssue, MAX_BLANK_CELL_FORMAT_GROUPS, MAX_BLANK_CELL_FORMAT_RANGES, MAX_BLANK_CELL_FORMAT_REF_LENGTH, MAX_BLANK_CELL_FORMAT_WORKBOOK_AREA } from "./spreadsheet/blank-cell-formats";
import { MAX_SPREADSHEET_INK_POINTS_CHARS } from "./spreadsheet/shapes";
import { SPARKLINE_NATIVE_GRAPH_BLOCKER } from "./spreadsheet/sparkline";
import spreadsheetPageEnvelope from "./spreadsheet/spreadsheet-page-envelope.shared.json";
import { MAX_SPREADSHEET_COMMENT_TEXT_LENGTH } from "./spreadsheet/comments";
import { MAX_STROKE_WIDTH, PRESENTATION_CONNECTOR_END_SIZES, PRESENTATION_CONNECTOR_ENDS, PRESENTATION_CONNECTOR_PRESETS, PRESENTATION_CONNECTOR_ROUTES, SHAPE_KINDS, STROKE_STYLES, isPresentationConnectorShapeKind, normalizePresentationConnectorSpec, normalizePresentationOpenShapeLineEnds, presentationShapeTakesLineEnds, normalizeShapeKind, normalizeStrokeWidth, presentationInkStrokeWidthPx } from "./presentation-shapes";
import { PRESENTATION_SHAPE_POINTS_MAX_LENGTH, normalizePresentationCustomPath, presentationCustomGeometryPathAdmitted, presentationCustomPathIsOpenLineOnly, presentationInkPointsWithinEnvelope, type PresentationCustomPath } from "./presentation-freeform";
import { INK_EFFECT_MAX_CHARS, isInkEffect, INK_CUT_ENDS, MAX_INK_PRESSURE_CHARS, normalizeDocumentShapeAdjustments } from "./document-shapes";
import { normalizeOfficeTextLayout, type OfficeTextLayout } from "./office-text-layout";
import { normalizeOfficePictureFillOptions, type OfficePictureFillOptions } from "./office-picture-fill-options";
import { BLOCK_ANIMATIONS, PRESENTATION_MEDIA_CLICK_ACTION_KINDS, PRESENTATION_RELATIVE_CLICK_ACTION_TARGETS, isPresentationActionButtonShapeKind, presentationHasAbsentNativeAction, isSafeBlockHref } from "./presentation-block-fx";
import { normalizeLinkTooltip } from "./presentation-link-settings";
import { MAX_TEXT_BLOCK_TEXT_LENGTH, BULLET_CHARS, MAX_PARAGRAPH_LEVEL, MAX_PARAGRAPH_TEXT_LENGTH, MAX_PARAGRAPHS_PER_BLOCK, MIN_PARAGRAPH_LEVEL, NUMBER_FORMATS, PARAGRAPH_ALIGNS, PARAGRAPH_BULLETS, PARAGRAPH_INDENT_EMU_LIMIT } from "./presentation-paragraphs";
import { MAX_EXACT_LINE_SPACING_PT, MAX_LINE_SPACING, MAX_PARAGRAPH_SPACING_PT, MIN_STORED_LINE_SPACING, MIN_PARAGRAPH_SPACING_PT } from "./presentation-line-spacing";
import { MAX_CELL_RUN_TEXT_LENGTH, MAX_RUN_FONT_SIZE, MAX_RUN_TEXT_LENGTH, MAX_RUNS_PER_BLOCK, MAX_CELL_RUNS, MIN_RUN_FONT_SIZE } from "./presentation-runs";
import { MAX_SHAPE_TEXT_LENGTH, normalizeShapeText } from "./presentation-text-format";
import { PRESENTATION_GEOMETRY_MAX_EMU, PRESENTATION_OBJECT_GEOMETRY_MAX_EMU } from "./presentation-geometry-limit";
import { OFFICE_STROKE_DASH_FAMILY, type OfficeStrokeDash } from "./office-stroke-dash";
import { OFFICE_STROKE_CAPS, OFFICE_STROKE_COMPOUNDS, OFFICE_STROKE_JOINS, OFFICE_STROKE_MAX_MITER_LIMIT } from "./office-stroke-detail";

// A spreadsheet cell / conditional-format color is stored as a literal hex
// string. We accept both the 6-digit (#rrggbb) and 3-digit (#rgb) shorthand the
// runtime normalizer (normalizeHexColor) understands, funnel the value through
// that same normalizer so the schema and the normalizer can never disagree about
// what a valid color is, and store the EXPANDED, lowercased #rrggbb form so the
// persisted value matches exactly what the renderer derives — a 3-digit "#abc"
// round-trips to "#aabbcc". Anything normalizeHexColor rejects (named colors,
// rgb(), CSS-injection attempts) fails to parse.
// Excel's ISO protection password hash (sheetProtection algorithmName/
// hashValue/saltValue/spinCount; workbookProtection workbook*): never a
// plaintext password. Bounds per MS-OFFCRYPTO §2.4.2.4.
export const protectionBase64Schema = z
  .string()
  .max(1_024)
  .regex(/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/);
export const protectionPasswordHashSchema = z
  .object({
    algorithmName: z.string().regex(/^[A-Za-z0-9-]{1,32}$/),
    hashValue: protectionBase64Schema.min(1),
    saltValue: protectionBase64Schema,
    spinCount: z.number().int().min(0).max(10_000_000),
  })
  .strict();
export const protectionPasswordVerifierSchema = z.discriminatedUnion("kind", [
  protectionPasswordHashSchema.extend({ kind: z.literal("hash") }).strict(),
  z
    .object({
      kind: z.literal("legacy"),
      verifier: z.string().regex(/^[0-9A-F]{4}$/),
    })
    .strict(),
]);
export const hexColorSchema = z
  .string()
  .max(7)
  .transform((value, ctx) => {
    const normalized = normalizeHexColor(value);
    if (normalized === null) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "Color must be a #rgb or #rrggbb hex value",
      });
      return z.NEVER;
    }
    return normalized;
  });

// A conditional-format range rectangle. Coordinates are bounded to the LAST
// addressable cell (MAX_RANGE_ROW / MAX_RANGE_COL = grid size - 1), never one
// cell past the grid edge, matching the runtime clamp in
// lib/spreadsheet/conditional-format (normalizeRange). Shared by every
// conditional-format rule kind.
export const conditionalRangeSchema = z
  .object({
    r1: z.number().int().min(0).max(MAX_RANGE_ROW),
    c1: z.number().int().min(0).max(MAX_RANGE_COL),
    r2: z.number().int().min(0).max(MAX_RANGE_ROW),
    c2: z.number().int().min(0).max(MAX_RANGE_COL),
  })
  .passthrough();

// Every built-in icon-set family name, as a z.enum tuple. Derived from the
// single source of truth (ICON_SET_FAMILIES) so adding a family there
// automatically widens the load-time schema — no second list to keep in sync.
export const iconSetNameEnum = ICON_SET_FAMILIES.map((f) => f.name) as [
  string,
  ...string[],
];


export const PRESENTATION_GEOMETRY_EMU_LIMIT = PRESENTATION_GEOMETRY_MAX_EMU;
/** An object frame's bound (`PRESENTATION_OBJECT_GEOMETRY_MAX_EMU`). */
export const PRESENTATION_OBJECT_GEOMETRY_EMU_LIMIT = PRESENTATION_OBJECT_GEOMETRY_MAX_EMU;

export const jsonScalarSchema = z.union([
  z.string(),
  z.number(),
  z.boolean(),
  z.null(),
]);
export const imageSourceSchema = z
  .string()
  .min(1)
  .max(2_000_000)
  .refine(isValidInlineImageSource, {
    message: "Invalid inline image data URL",
  });

// A worksheet picture that is not a floating object -- one placed IN a cell
// (Excel's Place in Cell) or a sheet's tiled background (Page Layout >
// Background). Both are packaged into the XLSX, so only an embedded raster
// data image within the page's single-string bound is admitted (the same
// rule lib/spreadsheet/sheet-image enforces on insert); the generic image
// source check above still authenticates the bytes against their MIME.
export const worksheetEmbeddedPictureSourceSchema = imageSourceSchema.refine(
  (value) =>
    // The API's MAX_JSON_STRING_CHARS (lib/office-picture-envelope).
    value.length <= 1_048_576 &&
    /^data:image\/(png|jpeg|gif|webp|bmp);base64,/i.test(value),
  { message: "A worksheet picture must be an embedded raster data image" },
);
// A workbook picture stored beside its page (lib/page-picture-objects): the
// digest of its bytes. The server requires it to be one of the page's own.
export const pictureObjectTokenSchema = z.string().regex(/^mrai-picture:[0-9a-f]{64}$/);
export const spreadsheetCellPictureSchema = z
  .object({
    // A picture read from a file past the embedded bound is stored beside the
    
    src: z.union([worksheetEmbeddedPictureSourceSchema, pictureObjectTokenSchema]),
    alt: z.string().max(2_000).optional(),
  })
  .strict();
export const spreadsheetSheetBackgroundPictureSchema = z
  .object({
    src: worksheetEmbeddedPictureSourceSchema,
  })
  .strict();

// Reusable strict six-digit hex colour (mirrors the slide background regex).
export const presentationBgHexSchema = z.string().regex(/^#[0-9a-fA-F]{6}$/);

export const presentationBgImageSchema = z
  .object({
    // The picture URL reuses the SAME validated inline-image-source predicate as
    // image blocks, so a stale / hand-edited deck can never persist a dangerous
    // background `url()`; the render edge re-checks too.
    url: imageSourceSchema,
    // `stretch`: an imported background PowerPoint stretches out of the
    // picture's proportions (`SlideBgPictureFit`).
    fit: z.enum(["cover", "contain", "tile", "stretch"]).optional(),
  })
  .passthrough();
// A shape's picture fill: the background's picture and fit, and Format
// Shape's picture options in the shared model (`a:tile`, the stretch's
// `a:fillRect`, Transparency, Rotate with shape), held to its bounds.
export const presentationShapeFillImageSchema = presentationBgImageSchema.superRefine((value, context) => {
  const { tile, fillRect, transparency, rotWithShape } = value as Record<string, unknown>;
  if (normalizeOfficePictureFillOptions({ tile, fillRect, transparency, rotWithShape }) === null) {
    context.addIssue({ code: z.ZodIssueCode.custom, message: "Unsupported picture fill options" });
  }
});
export const presentationBgPatternSchema = z
  .object({
    pattern: z.enum(["dots", "grid", "diagonal", "checker"]).optional(),
    color: presentationBgHexSchema.optional(),
    background: presentationBgHexSchema.optional(),
  })
  .passthrough();

// Inline rich-text run for a cell (Excel inline rich text): a text span + an
// optional per-run style. The style is whitelisted/validated at READ time by
// normalizeCellRuns (the shared presentation run model), so the schema only
// bounds the shape + text length here — the same normalize-on-read discipline
// the cell's other loose fields use. Reuses the presentation run bounds.
export const spreadsheetCellRunSchema = z.object({
  text: z.string().max(MAX_CELL_RUN_TEXT_LENGTH),
  style: z.record(z.unknown()).optional(),
});

export const officeShapeGradientSchema=z.custom<OfficeShapeGradient>(value=>{try{return Boolean(normalizeOfficeShapeGradient(value));}catch{return false;}});
export const officeShapeShadowSchema=z.custom<OfficeShapeShadow>(value=>{try{return Boolean(normalizeOfficeShapeShadow(value));}catch{return false;}});
// SmartArt quick-style 3-D (`docs/office-graphics/SMARTART-STYLES.md` §4). The
// guards enforce the same EMU / angle bounds the API validator does, so a
// malformed payload fails here instead of reaching a writer.
export const smartArtShape3dSchema=z.custom<SmartArtShape3d>(isSmartArtShape3d);
export const smartArtScene3dSchema=z.custom<SmartArtScene3d>(isSmartArtScene3d);
export const smartArtTextShape3dSchema=z.custom<SmartArtTextShape3d>(isSmartArtTextShape3d);
export const smartArtStyleIdSchema=z.custom<SmartArtStyleId>(isSmartArtStyleId);
export const smartArtColorsIdSchema=z.custom<string>((value) => smartArtColorVariant(value, { all: true }) !== null);
// The Office ALIAS a converted diagram is drawn as (`native-layouts`); the
// block's own kind is held to it by the reader (`smartArtAliasLayoutUrn`).
// A diagram laid out from its file's OWN definition (`office-definition`,
// SMARTART-GAP.md §8.6) records that definition's urn, whatever it is (a
// layout MRAI stocks none of, or a file's custom one): the server holds the
// kind, the urn and the part's digest to one another.
export const smartArtLayoutUrnSchema=z.custom<string>(
  (value) => isAliasNativeSmartArtLayout(value) || (typeof value === "string" && value.length >= 1 && value.length <= 512),
);
// Phonetic guide (furigana) of a cell: the readings over its text, its
// Phonetic Settings and Show Phonetic Field (lib/spreadsheet/cell-phonetic,
// which normalizes it at read time against the cell's text). Bounded here.
export const spreadsheetCellPhoneticSchema = z.object({
  runs: z
    .array(
      z.object({
        start: z.number().int().min(0).max(32_767),
        end: z.number().int().min(0).max(32_767),
        text: z.string().min(1).max(1_024),
      }),
    )
    .max(2_048)
    .optional(),
  type: z.enum(["halfwidthKatakana", "hiragana", "noConversion"]).optional(),
  alignment: z.enum(["noControl", "center", "distributed"]).optional(),
  font: z
    .object({
      family: z.string().min(1).max(200),
      size: z.number().finite().min(1).max(409),
      bold: z.literal(true).optional(),
      italic: z.literal(true).optional(),
      underline: z.enum(["single", "double"]).optional(),
      strikethrough: z.literal(true).optional(),
      color: hexColorSchema.optional(),
    })
    .optional(),
  visible: z.literal(true).optional(),
});

export const spreadsheetCellSchema = z
  .object({
    value: jsonScalarSchema,
    phonetic: spreadsheetCellPhoneticSchema.optional(),
    // Inline rich-text runs (Excel inline rich text): per-substring formatting.
    // The concatenated run text is kept == `value` so formulas / sort / CSV read
    // the flat value unchanged; the grid styles each run. Validated + collapsed
    // (flat-equivalent → dropped) at read via normalizeCellRuns. Bounded count.
    runs: z.array(spreadsheetCellRunSchema).max(MAX_CELL_RUNS).optional(),
    kind: z.enum(["text", "number"]).optional(),
    // An imported string whose characters read as TRUE, #N/A or a number
    // (SpreadsheetML t="s"/inlineStr): it stays a string (live audit B312).
    stringValue: z.literal(true).optional(),
    // Read-only DISPLAY CACHE (Option A): the cell's resolved on-screen string
    // (computed formula result and/or applied number format). Written by the
    // frontend serializer ONLY for cells where it adds information — a cell with
    // a `formula` or a `style.numberFormat` — so plain text/number cells stay
    // byte-stable (no `display` key). The CSV/plain-text export prefers this over
    // `value` so a formula cell exports "30" instead of "=A1+A2" and a formatted
    // cell exports "50.00%" instead of "0.5". Best-effort: absent for AI-written
    // or older content_json, in which case the backend falls back to `value`.
    display: z.string().max(4096).optional(),
    chartValueCache: z.object({
      kind: z.literal("number"),
      source: z.string().max(32_767).refine((source) => source.trimStart().startsWith("=")),
      value: z.number().finite(),
    }).strict().optional(),
    // Cell hyperlink (Excel/WPS parity): a clickable href stored on the cell.
    // Validated at PARSE time (not just re-sanitized at render) so an unsafe
    // persisted scheme — javascript:, data:, vbscript:, file:, tel:, relative,
    // protocol-relative — is rejected when the document loads, never silently
    // carried in the model. Funneled through the same sanitizeHyperlink used by
    // the renderer so schema and render can never disagree; only http(s)/mailto
    // survive. Bounded so a malicious file cannot persist a huge URL.
    //
    // A link to a place in THIS document (`#'Sheet'!A1`) is the other half:
    // the grid has always rendered those as an in-workbook jump, and the
    // importer reads them, but the schema knew only about URLs. A workbook
    // with a navigation sheet or a table of contents — which is what those
    // links are for — therefore failed to parse and rendered EMPTY.
    // Excel's dialog stops at 2,083 characters, its files do not, and the API
    // sets no bound: a link Excel saved made the whole workbook read-only here
    
    
    hyperlinkInArea: z.boolean().optional(),
    // The cell's own link is listed in the file after a link the sheet holds
    // over a large area that covers it (`hyperlinkAreas`): the place follows,
    // edits and copies the area's, the first over it (reading 950,
    // xlsxHyperlinkUnderArea; Windows Excel, xlsx-rounds probe/n35-p948).
    hyperlinkUnderArea: z.literal(true).optional(),
    // The link over several cells beneath this cell's own (the file lists the
    // cell's first): a sort that takes the cell's own away leaves this one on
    // the place (xlsxHyperlinkAreaUnder, R294-X08). Its address is checked as
    // `hyperlink` is.
    hyperlinkArea: z
      .object({
        hyperlink: z
          .string()
          .max(65_530)
          .refine(
            (value) =>
              sanitizeHyperlink(value) !== null || parseInternalLink(value) !== null || parseNamedLink(value) !== null,
          ),
        hyperlinkTooltip: z.string().max(255).optional(),
        hyperlinkDisplay: z.string().max(65_530).optional(),
        // Its place in the file's list of links (xlsxHyperlinkOwnUnder).
        hyperlinkRank: z.number().int().min(0).max(10_000_000).optional(),
      })
      .strict()
      .optional(),
    // The place in the file's list of links of the link the cell shows, on a
    // sheet with a link over several cells: a sort reads which of two links over
    // one place is first, as Excel's keeps their order (xlsxHyperlinkOwnUnder,
    
    hyperlinkRank: z.number().int().min(0).max(10_000_000).optional(),
    // The cell's own one-cell link the file lists after the area's link the cell
    // shows (`hyperlinkInArea`): a sort moves it with the cell and leaves the
    // area's on the place (xlsxHyperlinkOwnUnder, Windows Excel x35). Its address
    // is checked as `hyperlink` is.
    hyperlinkOwnUnder: z
      .object({
        hyperlink: z
          .string()
          .max(65_530)
          .refine(
            (value) =>
              sanitizeHyperlink(value) !== null || parseInternalLink(value) !== null || parseNamedLink(value) !== null,
          ),
        hyperlinkTooltip: z.string().max(255).optional(),
        hyperlinkDisplay: z.string().max(65_530).optional(),
        hyperlinkRank: z.number().int().min(0).max(10_000_000).optional(),
      })
      .strict()
      .optional(),
    hyperlink: z
      .string()
      .max(65_530)
      .refine(
        (value) =>
          sanitizeHyperlink(value) !== null ||
          parseInternalLink(value) !== null ||
          // A defined name is a place too (`#Budget`, live audit B359).
          parseNamedLink(value) !== null,
        {
          message: "Hyperlink must be a safe URL or a place in this document",
        },
      )
      .optional(),
    style: z
      .object({
        // The file's format record this cell came from (lib/spreadsheet/style-source-xf).
        sourceXf: z.number().int().min(0).max(65_489).optional(),
        bold: z.boolean().optional(),
        italic: z.boolean().optional(),
        underline: z.boolean().optional(),
        color: hexColorSchema.optional(),
        backgroundColor: hexColorSchema.optional(),
        // Excel's own limit on a number format code (B343).
        numberFormat: z.string().max(255).optional(),
        border: z.boolean().optional(),
        borderTop: z.boolean().optional(),
        borderRight: z.boolean().optional(),
        borderBottom: z.boolean().optional(),
        borderLeft: z.boolean().optional(),
        textAlign: z.enum(["left", "center", "right"]).optional(),
        verticalAlign: z.enum(["top", "middle", "bottom"]).optional(),
        wrapText: z.boolean().optional(),
      })
      .passthrough()
      .optional(),
    // A note or comment message holds up to Excel's cell text length (B315).
    comment: z
      .union([
        z
          .object({
            kind: z.literal("note"),
            text: z.string().min(1).max(MAX_SPREADSHEET_COMMENT_TEXT_LENGTH),
            author: z.string().max(200).optional(),
            // Show/Hide Note: only `true` is ever stored (absent = hidden).
            visible: z.literal(true).optional(),
            nativeBlocker: z.string().max(200).optional(),
          })
          .passthrough(),
        z
          .object({
            kind: z.literal("thread"),
            text: z.string().max(MAX_SPREADSHEET_COMMENT_TEXT_LENGTH),
            author: z.string().max(200).optional(),
            thread: z
              .array(
                z
                  .object({
                    text: z.string().min(1).max(MAX_SPREADSHEET_COMMENT_TEXT_LENGTH),
                    author: z.string().max(200).optional(),
                    ts: z.number().int().nonnegative().optional(),
                    nativeId: z.string().min(1).max(200).optional(),
                    personId: z.string().min(1).max(200).optional(),
                    dateTime: z.string().max(64).optional(),
                  })
                  .passthrough(),
              )
              .max(1_000),
            done: z.boolean().optional(),
            draft: z.string().max(MAX_SPREADSHEET_COMMENT_TEXT_LENGTH).optional(),
            nativeBlocker: z.string().max(200).optional(),
          })
          .passthrough(),
        // Backward-compatible receipt: normalizeCell assigns the discriminator.
        z
          .object({
            text: z.string().max(MAX_SPREADSHEET_COMMENT_TEXT_LENGTH),
            author: z.string().max(200).optional(),
            thread: z.array(z.unknown()).max(1_000).optional(),
            resolved: z.boolean().optional(),
          })
          .passthrough(),
      ])
      .optional(),
    // Cell lock flag (Excel/WPS sheet protection): every cell is "locked" by
    // default — the lock only takes effect once the sheet is protected. So only
    // an explicit `false` (unlocked) is meaningful; `true`/absent both mean
    // locked. Optional/additive so existing cells without it still validate, and
    // the editor persists ONLY `locked: false` so an unprotected / never-touched
    // sheet stays byte-for-byte. See lib/spreadsheet/protection. NOTE: this is a
    // workflow guard, not a security boundary (a hand-edited JSON can flip it).
    locked: z.boolean().optional(),
    // Format Cells > Protection > Hidden (cellXfs `protection hidden="1"`):
    // on a protected sheet the formula bar and edit box withhold the cell's
    // contents and a copy carries its value. Persisted only when true.
    formulaHidden: z.boolean().optional(),
    // A picture placed IN the cell (Excel's Place in Cell): the cell's value
    // while `value` holds Excel's `#VALUE!` fallback (lib/spreadsheet/
    // cell-picture). Optional/additive: absent on every other cell.
    picture: spreadsheetCellPictureSchema.optional(),
  })
  .passthrough();

// Formatted blank cells stored as ranges (lib/spreadsheet/blank-cell-formats).
// A group stands for the explicit blank `{ value: "", kind: "text", style,
// locked }` at every position its ranges cover and `cells` does not hold. The
// style goes through the exact cell style parser, so a range can never carry a
// format an explicit cell could not; references, grid bounds and the per-sheet
// budgets are checked by the sheet refinement below.
export const spreadsheetBlankCellFormatGroupSchema = z
  .object({
    style: spreadsheetCellSchema.shape.style
      .unwrap()
      .refine((style) => Object.keys(style).length > 0, {
        message: "A blank-cell format style cannot be empty",
      })
      .optional(),
    locked: z.literal(false).optional(),
    ranges: z
      .array(z.string().max(MAX_BLANK_CELL_FORMAT_REF_LENGTH))
      .min(1)
      .max(MAX_BLANK_CELL_FORMAT_RANGES),
  })
  .strict()
  .refine((group) => group.style !== undefined || group.locked === false, {
    message: "A blank-cell format needs a style or locked: false",
  });

export const spreadsheetObjectAnchorPointSchema = z
  .object({
    // Any cell of an Excel grid: a sheet may run far past 10,000 rows, and a
    // picture or chart far down one must not shut the whole workbook
    // read-only.
    row: z.number().int().min(0).max(MAX_GRID_ROWS - 1),
    column: z.number().int().min(0).max(MAX_GRID_COLUMNS - 1),
    // A frame turned a quarter about its Excel box may start before the
    // first row or column: a negative offset from line 0 (X-166).
    rowOffset: z.number().int().min(-1_000_000_000).optional(),
    columnOffset: z.number().int().min(-1_000_000_000).optional(),
  })
  .passthrough();

export const spreadsheetObjectAnchorSchema = z
  .object({
    kind: z.enum(["oneCell", "twoCell", "absolute"]),
    from: spreadsheetObjectAnchorPointSchema.optional(),
    to: spreadsheetObjectAnchorPointSchema.optional(),
    xEmu: z.number().int().min(0).optional(),
    yEmu: z.number().int().min(0).optional(),
    widthEmu: z.number().int().min(0).optional(),
    heightEmu: z.number().int().min(0).optional(),
    origin: z.literal("sheet").optional(),
  })
  .passthrough();

// Shadow effect for a floating object (Excel/WPS 形状效果/图片效果 → 阴影).
// Additive: absent means no shadow, which is what every object stored before
// this existed. Validated + clamped at the editor edge by
// lib/spreadsheet/shape-effects, which is also what turns the value into CSS,
// so a hand-edited file cannot push arbitrary text into a style attribute.
export const spreadsheetObjectShadowSchema = z
  .object({
    preset: z.string().max(40),
    // A loose string, deliberately: `hexColorSchema` here would make one bad
    // byte in a stored file fail the object union, then the sheet, then the
    // whole document — a malformed colour would blank the workbook rather than
    // lose a shadow. normalizeShadow refuses anything that is not #rrggbb
    // before it can reach a style attribute, the same way shape fill/stroke do.
    color: z.string().max(40).optional(),
    distance: z.number().optional(),
    blur: z.number().optional(),
    opacity: z.number().optional(),
  })
  .passthrough();

// Excel for the web Picture Format -> Crop persists the same DrawingML
// source-space keep rectangle used by the other Office editors.  Keep the
// Spreadsheet schema on the shared finite/overlap contract, while retaining a
// separate schema name so presentation-only fields never leak into worksheets.
export const spreadsheetPictureCropRectSchema = z
  .object({
    x: z.number(),
    y: z.number(),
    w: z.number(),
    h: z.number(),
  })
  .strict()
  .refine((value) => normalizeSignedCropRect(value) !== null, {
    message: "Invalid signed Spreadsheet picture crop keep rectangle",
  });


// Where a page-made object came from (an AI writer's workspace picture:
// `{ kind: "dataItem", dataItemId }`), beside its `source`, which is only
// ever an import identity (lib/office-object-provenance). A record names no
// package part, so no reader can take it for one.
export const officeObjectProvenanceSchema = z
  .object({ kind: z.string().min(1).max(64) })
  .catchall(z.union([z.string().max(500), z.number(), z.boolean(), z.null()]))
  .refine((value) => isOfficeProvenanceRecord(value), {
    message: "A provenance record names no package part",
  });

export const spreadsheetImageObjectSchema = z
  .object({
    id: z.string().min(1),
    type: z.literal("image"),
    // The inline source first: a malformed data URL reports why it is
    // refused, not the token pattern it does not match either.
    src: z.union([imageSourceSchema, pictureObjectTokenSchema]),
    alt: z.string().optional(),
    caption: z.string().optional(),
    hidden: z.boolean().optional(),
    nativeVisibilityBlocker: z.string().max(200).optional(),
    anchor: spreadsheetObjectAnchorSchema.optional(),
    source: z.record(z.unknown()).optional(),
    provenance: officeObjectProvenanceSchema.optional(),
    shadow: spreadsheetObjectShadowSchema.optional(),
    // Picture Effects beyond the Shadow (Reflection, Glow, Soft Edges, Bevel,
    // 3-D Format, 3-D Rotation): the deck's picture model
    // (`lib/office-picture-effects`, the server's `office_picture_effects`);
    // the Shadow stays `shadow`, and a Preset is not the sheet's.
    pictureEffects: z
      .record(z.string(), z.unknown())
      .refine(
        (value) =>
          normalizeOfficePictureEffects(value) !== null && !("preset" in value) && !("shadow" in value),
        "Unsupported picture effects",
      )
      .optional(),
    cropRect: spreadsheetPictureCropRectSchema.optional(),
    nativeImageCropBlocker: z.string().max(200).optional(),
    // Picture Format > Picture Styles (Office's shared 28-style catalogue).
    pictureStyle: z
      .string()
      .refine((value) => officePictureStyle(value) !== null, {
        message: "Unknown Office picture style",
      })
      .optional(),
    nativePictureStyleBlocker: z.string().max(200).optional(),
    nativeImageContentBlocker: z.literal("shared-media-part").optional(),
  })
  .passthrough();

export const spreadsheetChartDataSchema = z
  .object({
    headers: z.array(z.string().max(2_000)).max(20).optional(),
    rows: z
      .array(z.array(z.string().max(2_000)).max(20))
      .max(200)
      .optional(),
  })
  .passthrough();

export const spreadsheetChartValueAxisNumberFormatSchema = z
  .object({
    formatCode: z
      .string()
      .trim()
      .min(1)
      .max(255)
      .refine(
        (value) => !/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(value),
        { message: "Chart value-axis number format contains control characters" },
      ),
    sourceLinked: z.boolean(),
  })
  .strict();

export const spreadsheetChartSeriesRefSchema = z
  .object({
    title: z.string().min(1).max(500).optional(),
    titleText: z.string().min(1).max(500).optional(),
    categories: z.string().min(1).max(500).optional(),
    xValues: z.string().min(1).max(500).optional(),
    implicitX: z.literal(true).optional(),
    values: z.string().min(1).max(500),
    bubbleSizes: z.string().min(1).max(500).optional(),
    // A combination's series kind; "pie" / "doughnut" are the rings of a pie
    // and doughnut combination (a gauge, C20_tdf111173).
    renderType: z
      .enum(["column", "bar", "line", "area", "scatter", "bubble", "pie", "doughnut"])
      .optional(),
    axis: z.enum(["primary", "secondary"]).optional(),
    // The keys the chart part states as a defined name (the key holds its cells).
    names: z
      .object({
        title: z.string().min(1).max(500).optional(),
        categories: z.string().min(1).max(500).optional(),
        values: z.string().min(1).max(500).optional(),
        xValues: z.string().min(1).max(500).optional(),
        bubbleSizes: z.string().min(1).max(500).optional(),
      })
      .passthrough()
      .optional(),
  })
  .passthrough()
  .superRefine((value, context) => {
    if (value.xValues && value.implicitX) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["implicitX"],
        message: "seriesRefs require explicit xValues XOR implicitX",
      });
    }
    if (
      value.categories &&
      (value.xValues || value.implicitX || value.bubbleSizes)
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["categories"],
        message: "seriesRefs cannot mix category and XY authority",
      });
    }
    if (
      (value.renderType === "scatter" || value.renderType === "bubble") &&
      !value.xValues &&
      !value.implicitX
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["xValues"],
        message: "scatter series require xValues or authenticated implicitX",
      });
    }
    if (
      value.implicitX &&
      (value.renderType === "bubble" || value.bubbleSizes !== undefined)
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["implicitX"],
        message: "implicitX is supported only for scatter series",
      });
    }
    if (value.renderType === "bubble" && !value.bubbleSizes) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["bubbleSizes"],
        message: "bubble series require bubbleSizes",
      });
    }
  });

export const spreadsheetChartPaintSchema = z.union([
  z.literal("none"),
  z.string().regex(/^#[0-9a-fA-F]{6}$/),
]);
export const spreadsheetChartSeriesStyleSchema = z
  .array(
    z
      .object({
        seriesIndex: z.number().int().min(0).max(MAX_CHART_SERIES - 1),
        // Format Data Series > Fill > No fill (a waterfall's hidden base series).
        fill: spreadsheetChartPaintSchema.optional(),
        outline: spreadsheetChartPaintSchema.optional(),
      })
      .strict()
      .refine(
        (style) => style.fill !== undefined || style.outline !== undefined,
      ),
  )
  .min(1)
  .max(MAX_CHART_SERIES)
  .refine(
    (items) => new Set(items.map((item) => item.seriesIndex)).size === items.length,
  );
export const spreadsheetChartSeriesGradientStyleSchema = z
  .array(
    z
      .object({
        seriesIndex: z.number().int().min(0).max(MAX_CHART_SERIES - 1),
        from: z.string().regex(/^#[0-9a-f]{6}$/),
        to: z.string().regex(/^#[0-9a-f]{6}$/),
      })
      .strict()
      .refine((style) => style.from !== style.to),
  )
  .min(1)
  .max(MAX_CHART_SERIES)
  .refine(
    (items) => new Set(items.map((item) => item.seriesIndex)).size === items.length,
  );
export const spreadsheetChartSliceStyleSchema = z
  .array(
    z
      .object({
        pointIndex: z.number().int().min(0).max(MAX_CHART_POINTS_PER_SERIES - 1),
        fill: spreadsheetChartPaintSchema.optional(),
        outline: spreadsheetChartPaintSchema.optional(),
      })
      .strict()
      .refine(
        (style) => style.fill !== undefined || style.outline !== undefined,
      ),
  )
  .min(1)
  .max(MAX_CHART_POINTS_PER_SERIES)
  .refine(
    (items) =>
      new Set(items.map((item) => item.pointIndex)).size === items.length,
  );

export const spreadsheetChartLineSeriesStyleSchema = z
  .array(
    z
      .object({
        seriesIndex: z.number().int().min(0).max(MAX_CHART_SERIES - 1),
        markerSymbol: z
          .enum([
            "none",
            "circle",
            "dash",
            "diamond",
            "dot",
            "plus",
            "square",
            "star",
            "triangle",
            "x",
          ])
          .optional(),
        markerSize: z.number().int().min(2).max(72).optional(),
        smooth: z.boolean().optional(),
        lineWidthPt: z
          .number()
          .finite()
          .positive()
          .max(24)
          .refine(
            (value) => {
              const emu = Math.round(value * 12_700);
              return (
                emu >= 1 &&
                emu <= 24 * 12_700 &&
                Math.abs(value - emu / 12_700) <= 1e-9
              );
            },
          )
          .optional(),
      })
      .strict()
      .refine(
        (style) =>
          style.markerSymbol !== undefined ||
          style.smooth !== undefined ||
          style.lineWidthPt !== undefined,
      )
      .refine(
        (style) =>
          style.markerSize === undefined ||
          (style.markerSymbol !== undefined && style.markerSymbol !== "none"),
      ),
  )
  .min(1)
  .max(MAX_CHART_SERIES)
  .refine(
    (items) =>
      new Set(items.map((item) => item.seriesIndex)).size === items.length,
  );

// The shared chart model on a sheet chart, and its binding. The schema only
// asks for a record here: the model is a cache the readers (normalizeChart,
// the window's object normalizer) validate and drop when it is not one of
// theirs, so a projection from another reader can never cost the page its
// editability over a cache.
export const spreadsheetChartV2Schema = z.custom<ChartBlockAttrsV2>(
  (value) => value === undefined || (typeof value === "object" && value !== null && !Array.isArray(value)),
  "Invalid chart model",
);
export const spreadsheetChartBindingSchema = z.custom<SpreadsheetChartBinding>(
  (value) => value === undefined || (typeof value === "object" && value !== null && !Array.isArray(value)),
  "Invalid chart binding",
);

export const spreadsheetChartObjectSchema = z
  .object({
    id: z.string().min(1),
    type: z.literal("chart"),
    name: z.string().max(255).optional(),
    hidden: z.boolean().optional(),
    nativeVisibilityBlocker: z.string().max(200).optional(),
    title: z.string().max(2_000).optional(),
    caption: z.string().max(2_000).optional(),
    chartType: z.string().max(200).optional(),
    chartStyle: z.number().int().min(1).max(48).optional(),
    chartFormat: z.custom<SpreadsheetChartFormat>((value) => normalizeSpreadsheetChartFormat(value) !== undefined, "Invalid spreadsheet chart format").optional(),
    chartSubtype: z.string().min(1).max(100).optional(),
    desktopStyleId: z.string().min(1).max(100).optional(),
    sourceSheetId: z.string().min(1).max(200).optional(),
    nativeChartFormatBlocker: z.string().max(200).optional(),
    chartColorPalette: z.enum(CHART_COLOR_PALETTE_IDS).optional(),
    seriesStyles: spreadsheetChartSeriesStyleSchema.optional(),
    seriesGradientStyles: spreadsheetChartSeriesGradientStyleSchema.optional(),
    lineSeriesStyles: spreadsheetChartLineSeriesStyleSchema.optional(),
    sliceStyles: spreadsheetChartSliceStyleSchema.optional(),
    nativeChartColorBlocker: z.string().max(200).optional(),
    nativeChartColorEditBlocker: z.string().max(200).optional(),
    nativeChartLegendPositionBlocker: z.string().max(200).optional(),
    nativeChartLineBlocker: z.string().max(200).optional(),
    nativeChartLineWidthBlocker: z.string().max(200).optional(),
    nativeChartStyleBlocker: z.string().max(200).optional(),
    // The look an imported Office 2016 chart's style and colour parts give it;
    // display only, never written back, so any malformed reading is dropped.
    nativeLook: z
      .custom<ChartNativeLook>((value) => normalizeChartNativeLook(value) !== undefined, "Invalid chart look")
      .optional(),
    // The country an imported filled map's own geocoding cache binds each of
    // its cell labels to; display only, never written back.
    mapRegionCodes: z
      .array(
        z
          .object({
            label: z.string().min(1).max(200),
            id: z.string().regex(/^\d{3}$/),
          })
          .strict(),
      )
      .min(1)
      .max(100)
      .optional(),
    ranges: z.array(z.string().max(500)).max(50).optional(),
    // Chart editor (Excel/WPS parity): the SINGLE A1-style data range the chart
    // editor reads/writes (e.g. "A1:B10"). Optional/additive so existing charts
    // — which only ever had the plural `ranges` array — still validate; when
    // absent the editor falls back to the first usable `ranges` entry. The
    // value is funneled through lib/spreadsheet/chart-config (normalizeChartRange
    // -> chartRangeToA1) so a persisted value is a parseable, in-grid range.
    range: z.string().max(500).optional(),
    // Chart editor toggles (Excel/WPS parity): show/hide the legend and the
    // axis labels. Optional/additive so existing charts still validate; absent
    // means "shown" (the editor defaults both ON, matching Excel). MiniChart
    // reads these to draw or omit the legend / axis text.
    showLegend: z.boolean().optional(),
    // Legend placement (Excel "Legend position"); additive, normalized to "top"
    // by normalizeChartConfig when absent/forged.
    legendPosition: z.enum(["top", "bottom", "left", "right", "topRight"]).optional(),
    showAxisLabels: z.boolean().optional(),
    showDataLabels: z.boolean().optional(),
    axisTitleX: z.string().max(2_000).optional(),
    axisTitleY: z.string().max(2_000).optional(),
    valueAxisNumberFormat: spreadsheetChartValueAxisNumberFormatSchema.optional(),
    nativeValueAxisNumberFormatBlocker: z.string().max(200).optional(),
    showGridlines: z.boolean().optional(),
    showDataTable: z.boolean().optional(),
    seriesRefs: z.array(spreadsheetChartSeriesRefSchema).min(1).max(MAX_CHART_SERIES).optional(),
    chartV2: spreadsheetChartV2Schema.optional(),
    binding: spreadsheetChartBindingSchema.optional(),
    data: spreadsheetChartDataSchema.optional(),
    anchor: spreadsheetObjectAnchorSchema.optional(),
    source: z.record(z.unknown()).optional(),
    provenance: officeObjectProvenanceSchema.optional(),
    shadow: spreadsheetObjectShadowSchema.optional(),
  })
  .passthrough()
  .superRefine((chart, context) => {
    if (
      chart.seriesGradientStyles &&
      (chart.seriesStyles || chart.chartColorPalette)
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["seriesGradientStyles"],
        message: "series gradients cannot be combined with solid series paint or a palette",
      });
    }
  });

export const spreadsheetFloatingChartSchema = z
  .object({
    id: z.string().min(1).max(200),
    name: z.string().max(255).optional(),
    hidden: z.boolean().optional(),
    type: z.enum([
      "column",
      "bar",
      "line",
      "pie",
      "area",
      "scatter",
      "stacked-bar",
      "stacked-column",
      "combo",
      "doughnut",
      "radar",
      "bubble",
      "histogram",
      "waterfall",
      "box-and-whisker",
      "pareto",
      "treemap",
      "sunburst",
      "stock",
      "funnel",
      "map",
      "surface",
      "camera",
    ]),
    range: z.object({
      r1: z.number().int().min(0).max(MAX_RANGE_ROW),
      c1: z.number().int().min(0).max(MAX_RANGE_COL),
      r2: z.number().int().min(0).max(MAX_RANGE_ROW),
      c2: z.number().int().min(0).max(MAX_RANGE_COL),
    }),
    anchor: z.object({
      row: z.number().int().min(0).max(MAX_GRID_ROWS - 1),
      col: z.number().int().min(0).max(MAX_GRID_COLUMNS - 1),
    }),
    width: z.number().finite().min(80).max(4_000),
    height: z.number().finite().min(80).max(4_000),
    title: z.string().max(2_000).optional(),
    chartStyle: z.number().int().min(1).max(48).optional(),
    chartFormat: z.custom<SpreadsheetChartFormat>((value) => normalizeSpreadsheetChartFormat(value) !== undefined, "Invalid spreadsheet chart format").optional(),
    chartSubtype: z.string().min(1).max(100).optional(),
    desktopStyleId: z.string().min(1).max(100).optional(),
    sourceSheetId: z.string().min(1).max(200).optional(),
    nativeChartFormatBlocker: z.string().max(200).optional(),
    chartColorPalette: z.enum(CHART_COLOR_PALETTE_IDS).optional(),
    seriesStyles: spreadsheetChartSeriesStyleSchema.optional(),
    seriesGradientStyles: spreadsheetChartSeriesGradientStyleSchema.optional(),
    lineSeriesStyles: spreadsheetChartLineSeriesStyleSchema.optional(),
    sliceStyles: spreadsheetChartSliceStyleSchema.optional(),
    nativeChartColorEditBlocker: z.string().max(200).optional(),
    nativeChartLegendPositionBlocker: z.string().max(200).optional(),
    nativeChartLineBlocker: z.string().max(200).optional(),
    nativeChartLineWidthBlocker: z.string().max(200).optional(),
    seriesInColumns: z.boolean().optional(),
    // The range has no category line / no line of series names (Excel's
    // Insert Chart over numbers side by side, or without a heading).
    noCategories: z.literal(true).optional(),
    noSeriesNames: z.literal(true).optional(),
    
    nameRowsAbove: z.number().int().min(1).max(32).optional(),
    showLegend: z.boolean().optional(),
    legendPosition: z.enum(["top", "bottom", "left", "right", "topRight"]).optional(),
    showAxisLabels: z.boolean().optional(),
    showGridlines: z.boolean().optional(),
    showDataTable: z.boolean().optional(),
    showDataLabels: z.boolean().optional(),
    axisTitleX: z.string().max(2_000).optional(),
    axisTitleY: z.string().max(2_000).optional(),
    valueAxisNumberFormat: spreadsheetChartValueAxisNumberFormatSchema.optional(),
    nativeValueAxisNumberFormatBlocker: z.string().max(200).optional(),
    // The grid's own reach (lib/spreadsheet/grid-sizing), not 100,000 px.
    x: z.number().finite().min(0).max(MAX_DRAWING_LAYER_X_PX).optional(),
    y: z.number().finite().min(0).max(MAX_DRAWING_LAYER_Y_PX).optional(),
    // The chart's cell anchor (lib/spreadsheet/chart-anchor).
    drawingAnchor: spreadsheetObjectAnchorSchema.optional(),
    seriesRefs: z.array(spreadsheetChartSeriesRefSchema).min(1).max(MAX_CHART_SERIES).optional(),
    // The one chart model every editor shares, with where the chart's data
    // lives (lib/spreadsheet/chart-v2-model). Both or neither.
    chartV2: spreadsheetChartV2Schema.optional(),
    binding: spreadsheetChartBindingSchema.optional(),
    pivotSource: z
      .object({
        sheetId: z.string().trim().min(1).max(200),
        pivotName: z.string().trim().min(1).max(255).optional(),
      })
      .strict()
      .optional(),
    // A PivotChart's field buttons (PivotChart Analyze > Field Buttons).
    pivotOptions: z
      .object({
        visible: z.boolean().optional(),
        filter: z.boolean().optional(),
        legend: z.boolean().optional(),
        axis: z.boolean().optional(),
        values: z.boolean().optional(),
        expandCollapse: z.boolean().optional(),
      })
      .strict()
      .optional(),
  })
  .passthrough()
  .superRefine((chart, context) => {
    if (
      chart.seriesGradientStyles &&
      (chart.seriesStyles || chart.chartColorPalette)
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["seriesGradientStyles"],
        message: "series gradients cannot be combined with solid series paint or a palette",
      });
    }
  });

// Drawing shape (Excel/WPS 形状): a floating rectangle/ellipse/line/arrow/
// triangle, optionally labelled. The style sub-object is validated + defaulted
// at the editor edge via lib/spreadsheet/shapes; .passthrough() keeps it
// additive. Geometry rides the shared object anchor like image/chart.
export const spreadsheetShapeObjectSchema = z
  .object({
    id: z.string().min(1),
    type: z.literal("shape"),
    name: z.string().max(255).optional(),
    hidden: z.boolean().optional(),
    nativeVisibilityBlocker: z.string().max(200).optional(),
    // An ActiveX control's placeholder (`activex`): drawn where the control
    // sits, preserved as the file has it, never edited or removed here.
    nativeControlBlocker: z.string().max(200).optional(),
    shape: z
      .object({
        fillGradient:officeShapeGradientSchema.optional(),
        // SmartArt quick-style paint, the same fields a slide's shape block
        // carries (`presentationShapeBlockSchema`): a sheet diagram is its
        // members as ordinary shapes, so the paint rides on the style.
        shapeShadow:officeShapeShadowSchema.optional(),
        shape3d:smartArtShape3dSchema.optional(),
        // An ordinary shape's own `a:sp3d`/`a:scene3d`, DISPLAY-ONLY: the same
        // bounded values, in a differently named slot so nothing can mistake
        // the source package's solid for one the page owns and writes back.
        nativeShape3d: z
          .object({
            // Held to the display-only guards, not the SmartArt ones: an
            
            // refusing one here fails the whole workbook, not the one shape.
            shape3d: z.custom<SmartArtShape3d>(isNativeShape3d),
            scene3d: z.custom<SmartArtScene3d>(isNativeScene3d).optional(),
          })
          .optional(),
        // The camera of the group it came out of (`xdr:grpSpPr/a:scene3d`),
        // DISPLAY-ONLY like `nativeShape3d`: the group keeps it in the file.
        nativeGroupScene3d: z.custom<SmartArtScene3d>(isNativeScene3d).optional(),
        scene3d:smartArtScene3dSchema.optional(),
        textShape3d:smartArtTextShape3dSchema.optional(),
        smartArtStyle:smartArtStyleIdSchema.optional(),
        kind: z.string().max(40).optional(),
        // The shape's OWN outline (`a:custGeom`) and dragged adjust handles
        // (`a:prstGeom/a:avLst`): the same two envelopes the document and
        // slide blocks persist, bounded by the same shared normalizers. The
        // editor holds the handle NAMES to the kind's preset; the API does too.
        customPath: z
          .custom<PresentationCustomPath>(
            (value) => normalizePresentationCustomPath(value) !== null,
            "Unsupported custom path",
          )
          .optional(),
        adjustments: z
          .custom<Record<string, number>>(
            (value) => normalizeDocumentShapeAdjustments(value) !== null,
            "Unsupported shape adjustments",
          )
          .optional(),
        // `a:ln/a:headEnd|a:tailEnd`: `kind` names WHICH ends are decorated,
        // these name WHICH decoration (all six `a:ST_LineEndType` members) and
        // its optional `@w`/`@len`. A value the kind implies stays absent, and
        // an absent size stays absent — `CT_LineEndProperties` declares no
        // default for any of them (SHAPE-IMPORT-AUDIT.md §7.4).
        headEnd: z.enum(PRESENTATION_CONNECTOR_ENDS).optional(),
        tailEnd: z.enum(PRESENTATION_CONNECTOR_ENDS).optional(),
        headEndWidth: z.enum(PRESENTATION_CONNECTOR_END_SIZES).optional(),
        headEndLength: z.enum(PRESENTATION_CONNECTOR_END_SIZES).optional(),
        tailEndWidth: z.enum(PRESENTATION_CONNECTOR_END_SIZES).optional(),
        tailEndLength: z.enum(PRESENTATION_CONNECTOR_END_SIZES).optional(),
        fill: z.string().max(40).optional(),
        // Format Shape > Fill > Picture or texture fill authored on the page:
        // `src` (on the style, passthrough) is the shape's own fill picture.
        fillPicture: z.literal(true).optional(),
        // Its Format Shape options (tile, offsets, Transparency, Rotate with
        // shape), held to the shared model's bounds.
        fillPictureOptions: z
          .custom<OfficePictureFillOptions>(
            (value) => normalizeOfficePictureFillOptions(value) !== null,
            "Unsupported picture fill options",
          )
          .optional(),
        stroke: z.string().max(40).optional(),
        strokeWidth: z.number().optional(),
        // A text box's words, as lib/spreadsheet/shapes MAX_SHAPE_TEXT_LENGTH
        
        text: z.string().max(32_767).optional(),
        // The label's DISPLAY-ONLY layout receipt (`a:pPr` numbering/margins +
        // `a:bodyPr`), in the one envelope every host reads. Bounded by the
        // shared normalizer rather than restated here; a receipt that fails it
        // is dropped and the label draws as bare lines, which is what every
        // page stored before this field already does.
        textLayout: z
          .custom<OfficeTextLayout>(
            (value) => normalizeOfficeTextLayout(value) !== undefined,
            "Unsupported shape text layout",
          )
          .optional(),
        // Spreadsheet Draw-tab ink is persisted as one bounded shape per
        // stroke. The editor performs the stricter point/count normalization;
        // this schema cap prevents an oversized forged path reaching it.
        inkTool: z.enum(["pen", "highlighter", "pencil"]).optional(),
        inkEffect: z.string().max(INK_EFFECT_MAX_CHARS).refine(isInkEffect, "is not an ink effect").optional(),
        inkPoints: z.string().max(MAX_SPREADSHEET_INK_POINTS_CHARS).optional(),
        // XLS-A032: Ink Replay's per-point pen timing (ms from the first point).
        inkTimes: z.string().max(MAX_SPREADSHEET_INK_POINTS_CHARS).optional(),
        // A pressure pen's reading per point (paint only; the editor checks
        // the count against the points).
        inkPressure: z.string().max(MAX_INK_PRESSURE_CHARS).optional(),
        inkOpacity: z.number().min(0).max(1).optional(),
        // The ends a point eraser cut (flat); absent, both ends are round.
        inkCutEnds: z.enum(INK_CUT_ENDS).optional(),
      })
      .passthrough()
      .optional(),
    anchor: spreadsheetObjectAnchorSchema.optional(),
    source: z.record(z.unknown()).optional(),
    provenance: officeObjectProvenanceSchema.optional(),
    shadow: spreadsheetObjectShadowSchema.optional(),
  })
  .passthrough();

// PRESERVED Office SmartArt (DiagramML) anchored in a sheet drawing, in the
// exact shape the xlsx adapter produces
// (app/services/file_adapters/xlsx/native_smartart.py `as_object`). The five

// drawing of the native diagram and only its TEXT is editable; geometry never
// round-trips. A diagram whose layout is one of the 39 catalog layouts is NOT
// stored this way — it imports as ordinary `type: "shape"` members, identical
// to Insert ▸ SmartArt (docs/office-graphics/SMARTART-IMPORT.md).
//
// The member fields stay deliberately loose (`.passthrough()`, no colour
// regexes): one unexpected byte in a preserved drawing must degrade that
// member's paint at the render edge, never fail the object union — which
// would fail the sheet, then the whole workbook, and blank the document.
export const spreadsheetNativeSmartArtMemberSchema = z
  .object({
    memberId: z.string().min(1).max(64),
    role: z.enum(["shape", "text"]),
    zIndex: z.number().int().min(0).max(4_096).optional(),
    modelId: z.string().max(128).optional(),
    dataPointIds: z.array(z.string().max(128)).max(64).optional(),
    left: z.number().optional(),
    top: z.number().optional(),
    width: z.number().optional(),
    height: z.number().optional(),
    rotation: z.number().optional(),
    preset: z.string().max(64).optional(),
    // DrawingML paths, as `lib/native-drawing-geometry` reads them
    // (`{schemaVersion:"drawingml-paths-v1", paths, viewBox}`). Validated at the
    // render edge, not here, so one unreadable path degrades one member.
    customPath: z.record(z.unknown()).optional(),
    lineEnds: z.record(z.unknown()).optional(),
    paintOpacity: z.record(z.unknown()).optional(),
    strokeWidthEmu: z.number().optional(),
    fill: z.string().max(64).optional(),
    stroke: z.string().max(64).optional(),
    strokeWidth: z.number().optional(),
    text: z.string().max(2_000).optional(),
    textStyle: z.record(z.unknown()).optional(),
  })
  .passthrough();

export const spreadsheetSmartArtObjectSchema = z
  .object({
    id: z.string().min(1),
    type: z.literal("smartart"),
    name: z.string().max(255).optional(),
    hidden: z.boolean().optional(),
    groupId: z.string().max(200).optional(),
    anchor: spreadsheetObjectAnchorSchema.optional(),
    widthEmu: z.number().int().min(1).max(100_000_000),
    heightEmu: z.number().int().min(1).max(100_000_000),
    members: z.array(spreadsheetNativeSmartArtMemberSchema).max(400),
    // The one disposition this object supports. Geometry edits are refused by
    // the xlsx patchback, so the editor must never offer them either.
    nativeGroupDisposition: z.literal("preserve-only"),
    source: z.record(z.unknown()).optional(),
  })
  .passthrough();

export const spreadsheetSortRefSchema = z
  .string()
  .trim()
  .min(2)
  .max(64)
  .regex(/^\$?[A-Z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Z]{1,3}\$?[1-9][0-9]*)?$/i);

export const spreadsheetTextConditionFilterSchema = z
  .object({
    type: z.literal("text"),
    operator: z.enum([
      "contains",
      "does-not-contain",
      "begins-with",
      "ends-with",
      "equals",
      "not-equals",
      "matches",
      "does-not-match",
      // Custom AutoFilter comparisons against a value that is not a number:
      // the text of text cells is ordered against it.
      "greater-than",
      "greater-or-equal",
      "less-than",
      "less-or-equal",
    ]),
    operands: z.tuple([z.string().max(1_000)]),
  })
  .strict()
  .refine(
    (condition) =>
      condition.operands[0] !== "" ||
      ![
        "contains",
        "does-not-contain",
        "begins-with",
        "ends-with",
        "matches",
        "does-not-match",
        "greater-than",
        "greater-or-equal",
        "less-than",
        "less-or-equal",
      ].includes(condition.operator),
    { message: "wildcard condition operand must not be empty" },
  );

export const spreadsheetNumberConditionFilterSchema = z.discriminatedUnion(
  "operator",
  [
    z
      .object({
        type: z.literal("number"),
        operator: z.enum([
          "greater-than",
          "less-than",
          "greater-or-equal",
          "less-or-equal",
          "equals",
          "not-equals",
        ]),
        operands: z.tuple([z.number().finite()]),
      })
      .strict(),
    z
      .object({
        type: z.literal("number"),
        operator: z.literal("between"),
        operands: z.tuple([z.number().finite(), z.number().finite()]),
      })
      .strict(),
    z
      .object({
        type: z.literal("number"),
        // Legacy editor-state spelling. Runtime normalization migrates 1..255
        // to `top-items`; 256..500 keeps its old predicate and fails closed at
        // native Top10 export rather than silently disappearing or truncating.
        operator: z.literal("top-n"),
        operands: z.tuple([z.number().int().min(1).max(500)]),
      })
      .strict(),
    z
      .object({
        type: z.literal("number"),
        operator: z.enum(["top-items", "bottom-items"]),
        operands: z.tuple([z.number().int().min(1).max(255)]),
      })
      .strict(),
    z
      .object({
        type: z.literal("number"),
        operator: z.enum(["top-percent", "bottom-percent"]),
        operands: z.tuple([z.number().finite().min(0).max(100)]),
      })
      .strict(),
    z
      .object({
        type: z.literal("number"),
        operator: z.enum(["above-average", "below-average"]),
        operands: z.tuple([]),
      })
      .strict(),
  ],
);

// Excel's two-line Custom AutoFilter (`<customFilters and>` with two
// `<customFilter>`): two single lines joined by And or Or. A line is a text
// condition or one of the six number comparisons; ranked, average and between
// criteria are whole filters and never one line of a pair.
export const spreadsheetCustomCriterionConditionSchema = z.union([
  spreadsheetTextConditionFilterSchema,
  z
    .object({
      type: z.literal("number"),
      operator: z.enum([
        "greater-than",
        "less-than",
        "greater-or-equal",
        "less-or-equal",
        "equals",
        "not-equals",
      ]),
      operands: z.tuple([z.number().finite()]),
    })
    .strict(),
]);

export const spreadsheetCompoundConditionFilterSchema = z
  .object({
    type: z.literal("compound"),
    join: z.enum(["and", "or"]),
    conditions: z.tuple([
      spreadsheetCustomCriterionConditionSchema,
      spreadsheetCustomCriterionConditionSchema,
    ]),
  })
  .strict();

// Bounded migration-only shape emitted by the original Custom XML overlay.
// normalizeConditionFilter converts known aliases and drops unknown ones.
export const spreadsheetLegacyConditionFilterSchema = z
  .object({
    op: z.string().min(1).max(64),
    value: z.union([z.string().max(1_000), z.number().finite()]),
  })
  .strict();

export const spreadsheetConditionFiltersSchema = z
  .record(
    z.union([
      spreadsheetTextConditionFilterSchema,
      spreadsheetNumberConditionFilterSchema,
      spreadsheetCompoundConditionFilterSchema,
      spreadsheetLegacyConditionFilterSchema,
    ]),
  )
  .refine(
    (filters) =>
      Object.keys(filters).length <= MAX_GRID_COLUMN_SIZE_OVERRIDES &&
      Object.keys(filters).every((key) => /^(0|[1-9][0-9]*)$/.test(key)),
    { message: "conditionFilters must use bounded zero-based column keys" },
  );

export const spreadsheetSplitPaneNameSchema = z.enum([
  "topLeft",
  "topRight",
  "bottomLeft",
  "bottomRight",
]);
export const spreadsheetSplitCellRefSchema = z
  .string()
  .regex(/^[A-Z]{1,3}[1-9][0-9]*$/);
export const spreadsheetSplitSqrefSchema = z
  .string()
  .min(2)
  .max(4_096)
  .refine((value) => {
    const refs = value.split(" ");
    return (
      value === value.trim() &&
      refs.length >= 1 &&
      refs.length <= 64 &&
      refs.every((ref) =>
        /^(?:[A-Z]{1,3}[1-9][0-9]*)(?::[A-Z]{1,3}[1-9][0-9]*)?$/.test(
          ref,
        ),
      )
    );
  });
export const spreadsheetSplitPaneSelectionSchema = z
  .object({
    pane: spreadsheetSplitPaneNameSchema,
    activeCell: spreadsheetSplitCellRefSchema,
    sqref: spreadsheetSplitSqrefSchema,
    activeCellId: z.number().int().min(0).max(63),
  })
  .strict();
export const spreadsheetSplitPaneStateSchema = z
  .object({
    activePane: spreadsheetSplitPaneNameSchema,
    topLeftCell: spreadsheetSplitCellRefSchema,
    // sheetView/@topLeftCell: where the leading (top-left) pane is scrolled.
    leadingTopLeftCell: spreadsheetSplitCellRefSchema.optional(),
    selections: z.array(spreadsheetSplitPaneSelectionSchema).min(2).max(4),
  })
  .strict();

export function spreadsheetSplitRefBounds(
  value: string,
): { r1: number; c1: number; r2: number; c2: number } | null {
  const match = /^([A-Z]{1,3})([1-9][0-9]*)(?::([A-Z]{1,3})([1-9][0-9]*))?$/.exec(
    value,
  );
  if (!match) return null;
  const column = (label: string): number => {
    let result = 0;
    for (const character of label) {
      result = result * 26 + character.charCodeAt(0) - 64;
    }
    return result - 1;
  };
  const r1 = Number(match[2]) - 1;
  const c1 = column(match[1]);
  const r2 = Number(match[4] ?? match[2]) - 1;
  const c2 = column(match[3] ?? match[1]);
  return Number.isSafeInteger(r1) && Number.isSafeInteger(r2)
    ? { r1, c1, r2, c2 }
    : null;
}

export const spreadsheetPivotNativePartSchema = z
  .object({
    // OPC part names are relationship-resolved package-relative paths; their
    // physical folder is not guaranteed to be `xl/`. Keep a deliberately
    // narrow ASCII XML subset with canonical forward-slash segments so an
    // imported custody receipt can never name an absolute/traversal URI.
    path: z
      .string()
      .min(8)
      .max(512)
      .regex(
        /^[A-Za-z0-9_-][A-Za-z0-9_.-]*(?:\/[A-Za-z0-9_-][A-Za-z0-9_.-]*)*\.xml$/i,
      ),
    sha256: z.string().regex(/^[0-9a-f]{64}$/),
  })
  .strict();

// A native report's cache has a field per source column, and a sheet has
// 16,384 of them: a cache over 671 fields (xlsx-rounds R159-X06) failed the
// schema, and the whole workbook opened read-only for a report that is only
// preserved. What the editor can change is the reader's own limit, its blocker.
const SPREADSHEET_PIVOT_NATIVE_MAX_FIELDS = 16_384;
/** A PivotTable cache snapshot's cells, its header row included (lib/spreadsheet/pivot). */
const SPREADSHEET_PIVOT_SNAPSHOT_MAX_CELLS = 200_000;

export const spreadsheetPivotNativeFieldSchema = z
  .object({
    id: z.string().min(1).max(512),
    name: z.string().min(1).max(255),
    sourceIndex: z.number().int().min(0).max(SPREADSHEET_PIVOT_NATIVE_MAX_FIELDS - 1),
    cacheIndex: z.number().int().min(0).max(SPREADSHEET_PIVOT_NATIVE_MAX_FIELDS - 1),
    groupedFromCacheIndex: z.number().int().min(0).max(SPREADSHEET_PIVOT_NATIVE_MAX_FIELDS - 1).optional(),
  })
  .strict();

export const spreadsheetPivotNativeScalarSchema = z
  .object({
    kind: z.enum(["s", "n", "d", "b", "e", "m"]),
    raw: z.string().max(2_000),
  })
  .strict()
  .superRefine((scalar, context) => {
    if (
      (scalar.kind === "m" && scalar.raw !== "") ||
      (scalar.kind === "b" && scalar.raw !== "0" && scalar.raw !== "1")
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["raw"],
        message: "native Pivot scalar kind and lexical value do not agree",
      });
    }
  });

export const spreadsheetPivotNativeSnapshotSchema = z
  .object({
    revision: z.number().int().min(0).max(Number.MAX_SAFE_INTEGER),
    fingerprint: z.string().min(1).max(64),
    fieldIds: z.array(z.string().min(1).max(512)).min(1).max(SPREADSHEET_PIVOT_NATIVE_MAX_FIELDS),
    // Rows bounded by the 200,000 cells the editor and the save take
    
    matrix: z
      .array(z.array(z.string().max(2_000)).min(1).max(SPREADSHEET_PIVOT_NATIVE_MAX_FIELDS))
      .min(1)
      .max(SPREADSHEET_PIVOT_SNAPSHOT_MAX_CELLS),
    // Older editor snapshots omit this field and normalize preserve-only.
    // Every newly imported/editable native Pivot carries exact OOXML scalar
    // kind + lexical value so patchback never guesses from display strings.
    scalarMatrix: z
      .array(z.array(spreadsheetPivotNativeScalarSchema).min(1).max(SPREADSHEET_PIVOT_NATIVE_MAX_FIELDS))
      .min(1)
      .max(SPREADSHEET_PIVOT_SNAPSHOT_MAX_CELLS)
      .optional(),
  })
  .strict()
  .superRefine((snapshot, context) => {
    const width = snapshot.fieldIds.length;
    if (snapshot.matrix.some((row) => row.length !== width)) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["matrix"],
        message: "native Pivot cache rows must match the field identity width",
      });
    }
    if (snapshot.matrix.length * width > 200_000) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["matrix"],
        message: "native Pivot cache snapshot exceeds the editor cell boundary",
      });
    }
    if (
      snapshot.scalarMatrix &&
      (snapshot.scalarMatrix.length !== snapshot.matrix.length ||
        snapshot.scalarMatrix.some((row) => row.length !== width))
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["scalarMatrix"],
        message: "native Pivot scalar authority must match the display matrix",
      });
    }
  });

export const spreadsheetPivotAuthoredSnapshotSchema =
  spreadsheetPivotNativeSnapshotSchema.superRefine((snapshot, context) => {
    if (!snapshot.scalarMatrix) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["scalarMatrix"],
        message: "authored Pivot cache requires exact typed scalar authority",
      });
    }
  });

// Timeline-owned PivotTable date filters (Excel x15 `dateBetween`, whole
// days). See lib/spreadsheet/pivot-timeline.
export const spreadsheetPivotDateFilterSchema = z
  .object({
    colIndex: z.number().int().min(0).max(MAX_RANGE_COL),
    start: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
    end: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
    timeline: z.string().min(1).max(255),
    filterId: z.number().int().min(0).max(2_147_483_647).optional(),
    evalOrder: z.number().int().min(-1).max(2_147_483_647).optional(),
  })
  .strict();

export const spreadsheetPivotNativeAuthoritySchema = z
  .object({
    version: z.literal(1),
    disposition: z.enum(["editable", "preserve-only"]),
    blocker: z.string().min(1).max(200).optional(),
    pivotName: z.string().min(1).max(255),
    sourceSheetName: z.string().min(1).max(255),
    sourceRef: z.string().min(2).max(64),
    outputRef: z.string().min(2).max(64),
    styleName: z.string().min(1).max(64).optional(),
    worksheetRelationshipId: z.string().min(1).max(255),
    cacheId: z.number().int().min(0).max(2_147_483_647),
    cacheRelationshipId: z.string().min(1).max(255),
    recordsRelationshipId: z.string().min(1).max(255).optional(),
    pivotTable: spreadsheetPivotNativePartSchema,
    cacheDefinition: spreadsheetPivotNativePartSchema,
    cacheRecords: spreadsheetPivotNativePartSchema.optional(),
    fields: z.array(spreadsheetPivotNativeFieldSchema).min(1).max(SPREADSHEET_PIVOT_NATIVE_MAX_FIELDS),
    // The cache snapshot the PivotTable stands on: its own, or — when several
    // PivotTables read one cache — the workbook's `pivotCaches[cacheKey]`
    // entry they share (Excel keeps a cache once and refreshes them together).
    snapshot: spreadsheetPivotNativeSnapshotSchema.optional(),
    cacheKey: z.string().min(1).max(512).optional(),
    // Hidden items (item filters) the imported output cells were computed with;
    // `unit` names the level of a grouped date field the items belong to, and
    // `group` a manual grouping's group field.
    itemFilters: z
      .array(
        z
          .object({
            colIndex: z.number().int().min(0).max(MAX_RANGE_COL),
            selected: z.array(z.string().max(500)).max(10_000),
            unit: z.enum(["year", "quarter", "month"]).optional(),
            group: z.boolean().optional(),
          })
          .strict(),
      )
      .max(128)
      .optional(),
    // Timeline filters the imported output cells were computed with.
    dateFilters: z.array(spreadsheetPivotDateFilterSchema).max(16).optional(),
    // Label / Date / Value filters and Top 10 the imported output cells were
    // computed with (the importer's `fieldFilters`; normalizePivotNativeAuthority
    // reads them with lib/spreadsheet/pivot-field-filters). A report carrying
    // one failed this strict object and blanked the workbook (O04, X-235).
    fieldFilters: z
      .array(
        z
          .object({
            colIndex: z.number().int().min(0).max(MAX_RANGE_COL),
            unit: z.enum(["year", "quarter", "month"]).optional(),
            group: z.boolean().optional(),
            type: z.string().min(1).max(64),
            value1: z.union([z.string().max(255), z.number()]).optional(),
            value2: z.union([z.string().max(255), z.number()]).optional(),
            valueIndex: z.number().int().min(0).max(7).optional(),
            bottom: z.boolean().optional(),
            id: z.number().int().min(0).max(2_147_483_647).optional(),
          })
          .strict(),
      )
      .max(64)
      .optional(),
    // The cells the PivotTable occupied at import (pivotOutputFingerprint).
    outputFingerprint: z.string().regex(/^[0-9a-f]{16}$/).optional(),
    // The recipe's layout keys follow Excel's per-field reading.
    layoutReading: z.literal(2).optional(),
    // Classic PivotTable layout (the part's gridDropZones).
    classicLayout: z.literal(true).optional(),
    // Show the Values row cleared (the part's x14 hideValuesRow, X-266).
    hideValuesRow: z.literal(true).optional(),
    // The report's Values caption as its file states it (dataCaption, X-249).
    dataCaption: z.string().min(1).max(255).optional(),
    // The cache refreshes on open (pivotCacheDefinition@refreshOnLoad).
    refreshOnLoad: z.literal(true).optional(),
    // Where the cached report's areas lie (location, rowItems / colItems):
    // an undrawn report's style is painted from it (X-235). An import note
    // the normalizer drops when malformed, so it never blanks a workbook.
    savedLayout: z.unknown().optional(),
    // The report's formats on a field's item labels (format/pivotArea
    // labelOnly, X-283): an import note, dropped when malformed.
    labelFormats: z.unknown().optional(),
    // The report's formats on its grand totals, field buttons and chosen
    // items' values (X-294a): an import note, dropped when malformed.
    areaFormats: z.unknown().optional(),
    // Fields whose one Custom subtotal shows at the top of each group
    // (X-311): an import note, dropped when malformed.
    subtotalTopFields: z.unknown().optional(),
  })
  .strict()
  .superRefine((authority, context) => {
    if (
      (authority.disposition === "editable" && authority.blocker) ||
      (authority.disposition === "preserve-only" && !authority.blocker)
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["blocker"],
        message: "native Pivot blocker must match its preserve-only disposition",
      });
    }
    const ids = authority.fields.map((field) => field.id);
    const cacheIndexes = authority.fields.map((field) => field.cacheIndex);
    if (new Set(ids).size !== ids.length || new Set(cacheIndexes).size !== cacheIndexes.length) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["fields"],
        message: "native Pivot field identities must be unique",
      });
    }
    // Persisted, a shared cache's snapshot lives under `pivotCaches`; a
    // freshly imported or in-memory workbook still holds it on the PivotTable.
    if (authority.snapshot === undefined && authority.cacheKey === undefined) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["snapshot"],
        message: "native Pivot needs its own snapshot or a shared cache key",
      });
    }
    if (
      authority.snapshot &&
      (authority.snapshot.fieldIds.length !== ids.length ||
        authority.snapshot.fieldIds.some((id, index) => id !== ids[index]))
    ) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["snapshot", "fieldIds"],
        message: "native Pivot cache field identities are not aligned",
      });
    }
    const nativeParts: Array<
      readonly [
        "pivotTable" | "cacheDefinition" | "cacheRecords",
        { path: string },
      ]
    > = [
      ["pivotTable", authority.pivotTable],
      ["cacheDefinition", authority.cacheDefinition],
    ];
    if (authority.cacheRecords) {
      nativeParts.push(["cacheRecords", authority.cacheRecords]);
    }
    const paths = new Set<string>();
    for (const [name, part] of nativeParts) {
      const path = part.path.normalize("NFC").toLowerCase();
      if (paths.has(path)) {
        context.addIssue({
          code: z.ZodIssueCode.custom,
          path: [name, "path"],
          message: "native Pivot part paths must be case-insensitively unique",
        });
      }
      paths.add(path);
    }
  });

// One PivotTable recipe of a sheet. A sheet keeps its first PivotTable under
// `pivot` (the slot older pages know) and any others under `morePivots`.
export const spreadsheetSheetPivotSchema = z
      .object({
        source: z.string().max(500),
        name: z.string().min(1).max(255).optional(),
        // Pivot field references are SOURCE-COLUMN indices, so the max valid
        // index is the last grid column (MAX_RANGE_COL = MAX_GRID_COLUMNS - 1).
        // Using MAX_GRID_COLUMNS here would admit index 1000, which the runtime
        // clampIndex (lib/spreadsheet/pivot) silently drops -> data loss on load.
        rows: z.array(z.number().int().min(0).max(MAX_RANGE_COL)).max(64),
        columns: z.array(z.number().int().min(0).max(MAX_RANGE_COL)).max(64),
        values: z
          .array(
            z
              .object({
                colIndex: z.number().int().min(0).max(MAX_RANGE_COL),
                // Excel's eleven summary functions (XLS-A071).
                agg: z.enum([
                  "sum",
                  "count",
                  "average",
                  "max",
                  "min",
                  "product",
                  "countNums",
                  "stdDev",
                  "stdDevp",
                  "var",
                  "varp",
                ]),
                caption: z
                  .string()
                  .max(255)
                  .refine(
                    (caption) =>
                      !/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(
                        caption,
                      ),
                    {
                      message:
                        "PivotTable value caption contains XML control characters",
                    },
                  )
                  .optional(),
                numberFormat: z
                  .string()
                  .max(255)
                  .refine(
                    (format) =>
                      !/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(
                        format,
                      ),
                    {
                      message:
                        "PivotTable value number format contains XML control characters",
                    },
                  )
                  .optional(),
                // A value field may instead reference a CALCULATED field by name
                // (`calcRef`) and/or carry a "Show Values As" display transform
                // (`showAs`). Both are OPTIONAL + additive so an ordinary value
                // field persists byte-identically; the runtime re-validates via
                // normalizePivotConfig (an unknown calcRef / bad showAs is
                // dropped). See lib/spreadsheet/pivot.
                calcRef: z.string().max(128).optional(),
                showAs: z
                  .object({
                    // Every mode the pivot engine computes and the importer
                    // reads (lib/spreadsheet/pivot `PIVOT_SHOW_VALUES_AS`):
                    // a mode missing here (Excel's "% Of", showDataAs
                    // "percent", K24_BlankPivotTableField) made the whole
                    // workbook open blank and read-only.
                    mode: z.enum([
                      "normal",
                      "percentOfGrandTotal",
                      "percentOfRowTotal",
                      "percentOfColumnTotal",
                      "percentOfBaseItem",
                      "percentOfParentRowTotal",
                      "percentOfParentColumnTotal",
                      "percentOfParentTotal",
                      "runningTotalInRows",
                      "runningTotalInColumns",
                      "percentOfRunningTotalIn",
                      "rankSmallestToLargest",
                      "rankLargestToSmallest",
                      "differenceFrom",
                      "percentDifferenceFrom",
                      "index",
                    ]),
                    // A running total names only its base FIELD (Excel's
                    // dataField@baseField, no baseItem): `item` is the base
                    // value of the modes that measure against one (difference
                    // from, % of). Required here, a running total saved fine
                    // and its whole workbook failed to read back (S-114).
                    base: z
                      .object({
                        axis: z.enum(["row", "col"]),
                        pos: z.number().int().min(0).max(MAX_RANGE_COL),
                        item: z.string().max(256).optional(),
                      })
                      .passthrough()
                      .optional(),
                  })
                  .passthrough()
                  .optional(),
              })
              .passthrough(),
          )
          .max(64),
        // Calculated fields (Excel "Calculated Field"): user-defined formula
        // fields a value field may reference via `calcRef`. OPTIONAL + additive
        // (absent on existing sheets). Each is a unique name + an arithmetic
        // formula over other source field names; the list + lengths are bounded
        // so a hand-edited file cannot build a huge list / pathological formula
        // (DoS), and the runtime re-clamps via normalizePivotConfig.
        calculatedFields: z
          .array(
            z
              .object({
                name: z.string().max(128),
                formula: z.string().max(1024),
              })
              .passthrough(),
          )
          .max(32)
          .optional(),
        // Report / page filters (Excel PivotTable "Filters" area): each pares the
        // source rows to a SOURCE column's selected display values before
        // grouping. OPTIONAL + additive (absent on existing sheets). Bounds mirror
        // the runtime caps (lib/spreadsheet/pivot) so a hand-edited file cannot
        // smuggle an unbounded selection list (DoS); normalizePivotConfig
        // re-clamps + drops a filter on an axis column on load.
        filters: z
          .array(
            z
              .object({
                colIndex: z.number().int().min(0).max(MAX_RANGE_COL),
                selected: z.array(z.string().max(500)).max(10_000),
              })
              .passthrough(),
          )
          .max(16)
          .optional(),
        // Manual item filters (hidden PivotTable items): the visible values of a
        // non-report-filter field, written natively as `item h="1"`. Slicers
        // connected to the PivotTable read and write them. OPTIONAL + additive.
        itemFilters: z
          .array(
            z
              .object({
                colIndex: z.number().int().min(0).max(MAX_RANGE_COL),
                selected: z.array(z.string().max(500)).max(10_000),
              })
              .passthrough(),
          )
          .max(128)
          .optional(),
        // Field grouping (Excel "Group field"): bucket a ROW/COLUMN axis field by
        // date period or number bin. OPTIONAL + additive; normalizePivotConfig
        // drops a grouping on a non-axis column / invalid kind on load.
        grouping: z
          .array(
            z
              .object({
                colIndex: z.number().int().min(0).max(MAX_RANGE_COL),
                kind: z.enum(["number", "date", "discrete"]),
                size: z.number().optional(),
                start: z.number().optional(),
                unit: z.enum(["year", "quarter", "month"]).optional(),
                // Excel's manual groups (Group Selection): named groups of items.
                groups: z
                  .array(
                    z
                      .object({
                        name: z.string().min(1).max(255),
                        items: z.array(z.string().max(500)).min(1).max(10_000),
                      })
                      .strict(),
                  )
                  .max(1_000)
                  .optional(),
                display: z.enum(["groups+items", "groups", "items"]).optional(),
              })
              .passthrough(),
          )
          .max(64)
          .optional(),
        // Per-outer-group subtotal rows (Excel PivotTable subtotals). OPTIONAL +
        // additive; normalizePivotConfig keeps it only when explicitly true.
        rowSubtotals: z.boolean().optional(),
        // A newly-authored recommended Pivot owns a typed cache before XLSX
        // export allocates its native part and relationship identities.
        cacheSnapshot: spreadsheetPivotAuthoredSnapshotSchema.optional(),
        // Excel's Scenario PivotTable report: the cache above is the
        // scenarios' values (OOXML cacheSource type="scenario"), not a range.
        cacheSource: z.literal("scenario").optional(),
        native: spreadsheetPivotNativeAuthoritySchema.optional(),
        // Timeline (x15) date filters mirrored from connected timelines.
        dateFilters: z.array(spreadsheetPivotDateFilterSchema).max(16).optional(),
      })
      .passthrough();

export const spreadsheetSheetSchema = z
  .object({
    chartSheet: z.literal(true).optional(),
    nativeChartSheetBlocker: z.string().max(200).optional(),
    id: z.string().min(1),
    name: z.string(),
    rowCount: z.number().int().min(1).max(MAX_GRID_ROWS),
    columnCount: z.number().int().min(1).max(MAX_GRID_COLUMNS),
    // Sheet background (Excel's Page Layout > Background): one picture tiled
    // behind the cells on screen, never printed. Optional/additive; absent
    // means no background.
    backgroundPicture: spreadsheetSheetBackgroundPictureSchema.optional(),
    // Sheet tab color (Excel/WPS parity): an optional accent shown on the sheet
    // tab. Optional/additive so existing sheets without a color still validate;
    // absent means "no tab color". Funneled through the same hexColorSchema as
    // cell/conditional-format colors, so it stores the expanded, lowercased
    // #rrggbb form and a persisted value can never inject arbitrary CSS at
    // render (named colors, rgb(), CSS-injection attempts reject at parse time).
    tabColor: hexColorSchema.optional(),
    // View > Show > Gridlines / Headings off on this sheet (sheetView
    // @showGridLines / @showRowColHeaders). The editor stores only `true`;
    // absent means shown.
    hideGridlines: z.boolean().optional(),
    hideHeadings: z.boolean().optional(),
    // Show a zero in cells that have zero value, off (sheetView@showZeros).
    hideZeros: z.boolean().optional(),
    // Formulas > Show Formulas (sheetView@showFormulas) and the sheet's view
    // (sheetView@view, "pageLayout" / "pageBreakPreview"; absent = Normal).
    showFormulas: z.boolean().optional(),
    viewMode: z.string().max(32).optional(),
    // Sheet protection (Excel/WPS parity): when true, the sheet is in a read-only
    // guard mode — only cells explicitly unlocked (cell.locked === false) stay
    // editable; every other cell is blocked from edit/clear/paste. Optional/
    // additive so existing sheets validate; absent/false means "not protected".
    // The editor persists this ONLY when true so an unprotected sheet stays
    // byte-for-byte. See lib/spreadsheet/protection. NOTE: a WORKFLOW guard, not
    // a security boundary — a hand-edited JSON can flip it, and nothing enforces
    // it server-side. Mirrors Excel sheet protection, which is also bypassable.
    protected: z.boolean().optional(),
    // Native sheet-protection permissions. Persist only what differs from
    // Excel's Protect Sheet defaults; Allow Sort never overrides the
    // locked-cell requirement. `password` is the Unprotect verifier (ISO hash
    // or legacy 16-bit), never a plaintext password.
    protectionOptions: z
      .object({
        allowSort: z.boolean().optional(),
        // Excel's Format cells permission (live audit B138).
        allowFormatCells: z.boolean().optional(),
        allowFormatColumns: z.boolean().optional(),
        allowFormatRows: z.boolean().optional(),
        allowInsertColumns: z.boolean().optional(),
        allowInsertRows: z.boolean().optional(),
        allowInsertHyperlinks: z.boolean().optional(),
        allowDeleteColumns: z.boolean().optional(),
        allowDeleteRows: z.boolean().optional(),
        allowAutoFilter: z.boolean().optional(),
        allowPivotTables: z.boolean().optional(),
        allowEditObjects: z.boolean().optional(),
        allowEditScenarios: z.boolean().optional(),
        denySelectLockedCells: z.boolean().optional(),
        denySelectUnlockedCells: z.boolean().optional(),
        password: protectionPasswordVerifierSchema.optional(),
      })
      .optional(),
    // Hidden sheet (Excel/WPS parity): when true, the sheet is omitted from the
    // tab bar but its data is preserved; an "unhide" action restores it. Optional/
    // additive so existing sheets validate; the editor persists this ONLY when
    // true so a visible sheet stays byte-for-byte. At least one sheet must stay
    // visible (the hide action enforces it), so a workbook is never all-hidden.
    hidden: z.boolean().optional(),
    // Freeze panes (Excel parity): how many leading rows/columns stay pinned
    // while the rest of the grid scrolls. Optional so existing sheets without
    // freeze still validate; 0/absent means "no freeze".
    freezeRows: z.number().int().min(0).max(MAX_GRID_ROWS).optional(),
    freezeCols: z.number().int().min(0).max(MAX_GRID_COLUMNS).optional(),
    // True split panes are strict native view authority. Counts are divider
    // line indexes; the state must enumerate every physical pane exactly once.
    splitRow: z.number().int().min(0).max(MAX_GRID_ROWS - 1).optional(),
    splitCol: z.number().int().min(0).max(MAX_GRID_COLUMNS - 1).optional(),
    splitPaneState: spreadsheetSplitPaneStateSchema.optional(),
    // AutoFilter (Excel parity): per-column allowed display-value sets. A column
    // maps to the list of values that stay visible; rows whose value in that
    // column is not listed are hidden from the view (data is never deleted).
    // Keys are stringified column indices. Optional so existing sheets without
    // any filter still validate; absent/empty means "no filter". Both the
    // per-value length and the array length are bounded so a malicious file
    // cannot build a million-entry Set (DoS) when the filter is materialized.
    // A ticked value is a cell's whole display text (up to Excel's 32,767).
    filters: z.record(z.array(z.string().max(32_767)).max(10_000)).optional(),
    // Exact worksheet AutoFilter identity used by native Named Sheet Views.
    // OOXML stores this as xr:uid; keeping it separate from visible criteria
    // prevents an editor-created view from inventing or rebinding native state.
    autoFilterId: z
      .string()
      .regex(
        /^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$/,
      )
      .optional(),
    autoFilterRef: spreadsheetSortRefSchema.optional(),
    // The package's Excel `Database` range (_xlnm.Database), carried so a
    // row/column insert or delete moves it; never shown or edited.
    databaseRef: spreadsheetSortRefSchema.optional(),
    // Editor custody receipt: present only when creating the first local Sheet
    // View also had to synthesize the worksheet AutoFilter identity. Native
    // imports never infer this bit from an otherwise identical empty filter.
    namedSheetViewsOwnsAutoFilter: z.literal(true).optional(),
    // Excel Sheet Views ([MS-XLSX] 2.6): each view holds the worksheet
    // AutoFilter's criteria and sort keys and every Table AutoFilter's, in the
    // same grammars the worksheet and Tables use themselves. Foreign and
    // extension graphs remain preserve-only and carry a blocker instead of
    // being flattened.
    namedSheetViews: z
      .array(
        z
          .object({
            id: z
              .string()
              .regex(
                /^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$/,
              ),
            name: z
              .string()
              .min(1)
              .refine(
                (name) =>
                  name === name.trim() &&
                  Array.from(name).length <= 127 &&
                  !/[\u0000-\u001F\u007F]/.test(name) &&
                  !/[^\S ]/u.test(name) &&
                  !name.includes(";") &&
                  !/^_xlnsv\./i.test(name),
                "invalid named Sheet View name",
              ),
            filterId: z
              .string()
              .regex(
                /^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$/,
              )
              .optional(),
            ref: spreadsheetSortRefSchema.optional(),
            filters: z
              .record(z.array(z.string().max(32_767)).min(1).max(10_000))
              .optional(),
            // A view's typed AutoFilter conditions, one per absolute column,
            // in the same canonical shape as the worksheet's own.
            conditionFilters: spreadsheetConditionFiltersSchema.optional(),
            // Excel's date-tree criteria per absolute column.
            dateGroups: z
              .record(
                z.custom<FilterDateGroup[]>(
                  (value) => normalizeFilterDateGroups(value) !== null,
                  { message: "invalid Sheet View date groups" },
                ),
              )
              .optional(),
            // Filter by colour, normalized fills per absolute column; one
            // "#rrggbb" entry is a native colorFilter, others stay page-only.
            colorFilters: z
              .record(z.array(z.string().max(64)).min(1).max(256))
              .optional(),
            sortState: z
              .object({
                ref: spreadsheetSortRefSchema,
                conditions: z
                  .array(
                    z
                      .object({
                        ref: spreadsheetSortRefSchema,
                        descending: z.boolean().optional(),
                        sortBy: z
                          .enum(["cellColor", "fontColor", "icon"])
                          .optional(),
                        color: z
                          .string()
                          .regex(/^#[0-9a-fA-F]{6}$/)
                          .optional(),
                        icon: z
                          .object({
                            iconSet: z.enum(DATA_TABLE_ICON_FILTER_SETS),
                            iconId: z.number().int().min(0).max(4),
                          })
                          .strict()
                          .refine(
                            (icon) => icon.iconId < iconSetCount(icon.iconSet),
                            { message: "Sheet View sort icon is out of range" },
                          )
                          .optional(),
                        customList: z
                          .array(z.string().min(1).max(255))
                          .min(1)
                          .max(1_000)
                          .optional(),
                      })
                      .strict()
                      .refine(
                        (condition) =>
                          condition.sortBy === undefined
                            ? condition.color === undefined &&
                              condition.icon === undefined
                            : condition.customList === undefined &&
                              (condition.sortBy === "icon"
                                ? condition.icon !== undefined &&
                                  condition.color === undefined
                                : condition.color !== undefined &&
                                  condition.icon === undefined),
                        {
                          message:
                            "a Sheet View sort key is a value, colour or icon key",
                        },
                      ),
                  )
                  .min(1)
                  .max(64),
                caseSensitive: z.literal(true).optional(),
                sortMethod: z.enum(["pinYin", "stroke"]).optional(),
              })
              .strict()
              .optional(),
            // The view's Table AutoFilters, in the Tables' own grammar.
            tables: z
              .array(
                z
                  .object({
                    tableId: z.string().min(1).max(200),
                    filterId: z
                      .string()
                      .regex(
                        /^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$/,
                      ),
                    columnFilters: z
                      .record(
                        z.custom<DataTableColumnFilter>(
                          (value) => normalizeDataTableColumnFilter(value) !== null,
                          { message: "invalid Sheet View table criterion" },
                        ),
                      )
                      .optional(),
                    sortState: z
                      .custom<DataTableSortState>(
                        (value) =>
                          Array.isArray(
                            (value as { conditions?: unknown } | null)?.conditions,
                          ) &&
                          normalizeDataTableSortState(
                            value,
                            (value as { conditions: Array<{ columnId?: unknown }> })
                              .conditions.flatMap((condition) =>
                                typeof condition?.columnId === "number"
                                  ? [{ id: condition.columnId, name: "" }]
                                  : [],
                              ),
                          ) !== null,
                        { message: "invalid Sheet View table sort" },
                      )
                      .optional(),
                  })
                  .strict(),
              )
              .max(64)
              .optional(),
            nativeBlocker: z.string().min(1).max(255).optional(),
          })
          .strict()
          .superRefine((view, context) => {
            const hasFilterPayload =
              view.ref !== undefined ||
              view.filters !== undefined ||
              view.conditionFilters !== undefined ||
              view.dateGroups !== undefined ||
              view.colorFilters !== undefined ||
              view.sortState !== undefined;
            if (hasFilterPayload && view.filterId === undefined) {
              context.addIssue({
                code: z.ZodIssueCode.custom,
                message: "named Sheet View filterId is required",
              });
            }
          }),
      )
      .max(256)
      .superRefine((views, context) => {
        const ids = new Set<string>();
        const names = new Set<string>();
        for (const view of views) {
          const id = view.id.toUpperCase();
          const name = view.name.toLocaleLowerCase();
          if (ids.has(id)) {
            context.addIssue({
              code: z.ZodIssueCode.custom,
              message: "namedSheetViews ids must be unique",
            });
          }
          if (names.has(name)) {
            context.addIssue({
              code: z.ZodIssueCode.custom,
              message: "namedSheetViews names must be unique",
            });
          }
          ids.add(id);
          names.add(name);
        }
      })
      .optional(),
    namedSheetViewsNativeBlocker: z.string().min(1).max(255).optional(),
    // Set by an importer that read every note's Show/Hide Note state, so a
    // note without `visible` is hidden rather than unknown. Only `true`.
    noteVisibilityRead: z.literal(true).optional(),
    // Typed text/number AutoFilter criteria. Exactly one bounded descriptor is
    // persisted per absolute column; the small legacy union above is migration
    // input only and is normalized before runtime execution.
    conditionFilters: spreadsheetConditionFiltersSchema.optional(),
    // The worksheet's last supported value sort. References stay as bounded A1
    // ranges so Reapply, native import and original-package patchback all share
    // one authority; unsupported color/icon/custom-list sorts never enter this
    // shape and therefore cannot masquerade as a native value sort.
    sortState: z
      .object({
        ref: spreadsheetSortRefSchema,
        conditions: z
          .array(
            z
              .object({
                ref: spreadsheetSortRefSchema,
                descending: z.boolean().optional(),
              })
              .strict(),
          )
          .min(1)
          .max(64),
      })
      .strict()
      .optional(),
    // Advanced Filter's in-place result is native worksheet state, not a
    // manual row hide and not worksheet/table AutoFilter criteria. OOXML keeps
    // the list range in the sheet-scoped _FilterDatabase name and the current
    // result as hidden rows while sheetPr/filterMode is active, and the
    // criteria range in the sheet-scoped _xlnm.Criteria name (Excel writes it
    // with every Advanced Filter). Copy/unique are command inputs only.
    advancedFilterState: z
      .object({
        ref: spreadsheetSortRefSchema,
        criteria: spreadsheetSortRefSchema.optional(),
        hiddenRows: z
          .array(
            z
              .number()
              .int()
              .min(0)
              .max(MAX_GRID_ROWS - 1),
          )
          .max(MAX_GRID_ROW_SIZE_OVERRIDES)
          .refine(
            (rows) =>
              rows.every(
                (row, index) => index === 0 || row > rows[index - 1]!,
              ),
            {
              message:
                "advancedFilterState.hiddenRows must be sorted and unique",
            },
          ),
      })
      .strict()
      .optional(),
    // Column-width / row-height overrides (Excel/WPS parity). Keys are
    // stringified line indices; values are pixel sizes. Optional so existing
    // sheets without any manual resize still validate; a missing entry means
    // "use the default size". The per-value bounds match the resize clamp in
    // lib/spreadsheet/grid-sizing (columns clamp to
    // MIN_COLUMN_WIDTH_PX..MAX_COLUMN_WIDTH_PX = 1..4096 px; rows to
    // MIN_ROW_HEIGHT_PX..MAX_ROW_HEIGHT_PX = 18..512 px) so a persisted size can
    // stays within that range unless paired native row metadata proves a source
    // dimension outside it. Map entry counts use independent sparse
    // budgets; the full Excel address ceiling must never become an allocation
    // allowance.
    columnWidths: z
      .record(z.number().min(MIN_COLUMN_WIDTH_PX).max(MAX_COLUMN_WIDTH_PX))
      .refine(
        (map) =>
          Object.keys(map).length <= MAX_GRID_COLUMN_SIZE_OVERRIDES,
        {
          message: `columnWidths cannot have more than ${MAX_GRID_COLUMN_SIZE_OVERRIDES} entries`,
        },
      )
      .optional(),
    rowHeights: z
      // Native row point metadata supports up to 409.5pt/546px. Ordinary
      // authored dimensions retain the interactive limits in the sheet check.
      .record(z.number().min(1).max(546))
      .refine(
        (map) => Object.keys(map).length <= MAX_GRID_ROW_SIZE_OVERRIDES,
        {
          message: `rowHeights cannot have more than ${MAX_GRID_ROW_SIZE_OVERRIDES} entries`,
        },
      )
      .optional(),
    // Merged cells (Excel/WPS parity): each entry merges a rectangular block
    // into one visual cell, anchored at its top-left (row, col) with the given
    // rowSpan/colSpan. The anchor keeps its value; the covered cells are not
    // separately rendered and read as blank from formulas. Optional so existing
    // sheets without merges still validate; absent/empty means "no merges". The
    // array length is bounded so a malicious file cannot build a huge span list
    // (DoS on iteration); spans are clamped to the sheet's grid dimensions.
    merges: z
      .array(
        z
          .object({
            row: z.number().int().min(0).max(MAX_GRID_ROWS),
            col: z.number().int().min(0).max(MAX_GRID_COLUMNS),
            rowSpan: z.number().int().min(1).max(MAX_GRID_ROWS),
            colSpan: z.number().int().min(1).max(MAX_GRID_COLUMNS),
          })
          .passthrough(),
      )
      .max(MAX_GRID_ROW_SIZE_OVERRIDES)
      .optional(),
    // Outline groups (Excel/WPS parity): per-axis collapsible row/column
    // groupings. Each entry is a contiguous inclusive line range on one axis
    // plus a collapsed flag; collapsing hides the body lines (the renderer skips
    // them; SUBTOTAL 101-111 skips collapsed rows via isRowManuallyHidden).
    // Nesting is by containment and capped at 8 levels by the runtime
    // (normalizeOutlineGroups). Optional so existing sheets without outlines
    // still validate; absent/empty means "no groups". The array length is
    // bounded so a malicious file cannot build a huge list (DoS on iteration).
    outlineGroups: z
      .array(
        z
          .object({
            axis: z.enum(["row", "col"]),
            start: z.number().int().min(0).max(MAX_GRID_ROWS),
            end: z.number().int().min(0).max(MAX_GRID_ROWS),
            collapsed: z.boolean(),
          })
          .passthrough()
          // Axis-aware bounds: a column group's indices live in [0, cols) and a
          // row group's in [0, rows). Validating both against the row max would
          // let an out-of-range column group pass, then get silently clamped on
          // load (normalizeOutlineGroups) — a round-trip mutation. Reject it at
          // the gate instead. `start <= end` is also enforced so a reversed
          // range can't encode a phantom group.
          .refine(
            (g) => {
              const lineMax =
                g.axis === "col" ? MAX_GRID_COLUMNS : MAX_GRID_ROWS;
              return g.start <= g.end && g.start < lineMax && g.end < lineMax;
            },
            { message: "outlineGroups range is out of bounds for its axis" },
          ),
      )
      .max(MAX_GRID_ROW_SIZE_OVERRIDES)
      .optional(),
    // Excel's Outline Settings (Data > Outline > Settings; OOXML
    // sheetPr/outlinePr summaryBelow / summaryRight): false puts every group's
    // summary line, its toggle and its collapsed flag, before the group
    // instead of after it. Absent means Excel's default below / right.
    outlineSummaryBelow: z.boolean().optional(),
    outlineSummaryRight: z.boolean().optional(),
    /**
     * Preserve-only receipt set by an importer that could not read this
     * sheet's outline exactly as the file holds it (more than eight levels, a
     * malformed row/column outline attribute, lines outside the projection).
     * The native patchback writer would refuse any outline change there, so
     * the editor keeps Group / Ungroup / collapse off while it is present.
     */
    outlineGroupsNativeBlocker: z.string().min(1).max(220).optional(),
    // Manually hidden rows/columns (Excel "Hide rows/columns", DISTINCT from
    // outline grouping): arbitrary sets of absolute 0-based line indices the user
    // collapsed via the header context menu. Validated as in-range non-negative
    // integers and bounded to one entry per addressable line so a malicious file
    // cannot build a huge list (DoS on iteration) or inject an out-of-range index
    // that would be silently clamped on load (a round-trip mutation). Optional so
    // existing sheets without hidden lines validate BYTE-FOR-BYTE; absent/empty
    // means "none". See lib/spreadsheet/manual-hidden for the normalize/hide math.
    manuallyHiddenRows: z
      .array(
        z
          .number()
          .int()
          .min(0)
          .max(MAX_GRID_ROWS - 1),
      )
      .max(MAX_GRID_ROW_SIZE_OVERRIDES)
      .optional(),
    manuallyHiddenCols: z
      .array(
        z
          .number()
          .int()
          .min(0)
          .max(MAX_GRID_COLUMNS - 1),
      )
      .max(MAX_GRID_COLUMN_SIZE_OVERRIDES)
      .optional(),
    // Reading 998 (xlsx-rounds R474-X20): which hidden rows the file writes
    // 0 points high (ht="0"), a subset of manuallyHiddenRows. A filter that
    // shows its rows (a criterion met, Clear, Filter off) leaves them hidden,
    // as Windows Excel 16 does (lib/spreadsheet/filter-buttons).
    zeroHeightRows: z
      .array(
        z
          .number()
          .int()
          .min(0)
          .max(MAX_GRID_ROWS - 1),
      )
      .max(MAX_GRID_ROW_SIZE_OVERRIDES)
      .optional(),
    // The rows the sheet's filters hid when they were last applied, and what
    // they asked then (lib/spreadsheet/filter-result): an edited value does
    // not filter again, as in Excel.
    filterResult: z
      .object({
        rows: z
          .array(z.number().int().min(0).max(MAX_GRID_ROWS - 1))
          .max(spreadsheetPageEnvelope.maxArrayLength),
        criteria: z.string().min(1).max(64).optional(),
      })
      .strict()
      .optional(),
    // Rows hidden by default (sheetFormatPr/@zeroHeight, imported): the rows
    // past rowCount are hidden as well as those listed above. Absent = Excel's
    // default, every row shown unless hidden.
    rowsHiddenByDefault: z.literal(true).optional(),
    
    // a sheet hidden by default (row 1,048,576 of a run hidden to 1,048,575).
    // A row delete brings in shown rows at the grid's end, one per row deleted
    // (Excel shows them), so the list is bounded by the grid, not by a small
    // count: 64 made a sheet with 65 rows deleted open read-only.
    rowsShownPastEnd: z
      .array(z.number().int().min(0).max(MAX_GRID_ROWS - 1))
      .max(MAX_GRID_ROWS)
      .optional(),
    // Reading 945: the hidden default is the file's own run of hidden rows,
    // not zeroHeight; a row delete brings shown rows in at the grid's end.
    rowsHiddenByFileRun: z.literal(true).optional(),
    
    // large areas (a whole row, A7:XFD7), held as ranges of the sheet instead
    // of on each cell, in the file's order. A cell under one states only a
    // link of its own; the link a place follows is its cell's own, else the
    // first range over it.
    // The address is checked as a cell's `hyperlink` is.
    hyperlinkAreas: z
      .array(
        z
          .object({
            range: z
              .object({
                r1: z.number().int().min(0).max(MAX_RANGE_ROW),
                c1: z.number().int().min(0).max(MAX_RANGE_COL),
                r2: z.number().int().min(0).max(MAX_RANGE_ROW),
                c2: z.number().int().min(0).max(MAX_RANGE_COL),
              })
              .strict(),
            hyperlink: z
              .string()
              .min(1)
              .max(65_530)
              .refine(
                (value) =>
                  sanitizeHyperlink(value) !== null || parseInternalLink(value) !== null || parseNamedLink(value) !== null,
              ),
            hyperlinkTooltip: z.string().max(255).optional(),
            hyperlinkDisplay: z.string().max(65_530).optional(),
            
            // `hyperlinkRank` counts it (the cells' elements listed before it).
            hyperlinkRank: z.number().int().min(0).max(10_000_000).optional(),
          })
          .strict(),
      )
      .max(10_000)
      .optional(),
    // Named ranges (Excel/WPS parity): per-sheet defined names that formulas may
    // use in place of a cell/range reference (e.g. =SUM(Sales)). Each entry maps
    // a name to an A1 reference string ("B2" or "B2:B10"). Optional so existing
    // sheets without names still validate; absent/empty means "no names". The
    // array length and per-field lengths are bounded so a malicious file cannot
    // build a huge name table (DoS on iteration/resolution).
    namedRanges: z
      .array(
        z
          .object({
            name: z.string().min(1).max(255),
            // A cell/range ref is short, but a named FORMULA / LAMBDA ref (the
            // VBA-free reusable macro, e.g. "LAMBDA(price, qty, IF(...))") can be
            // long — bounded to MAX_FORMULA_REF_LENGTH so a real named formula
            // survives the reload round-trip (the runtime re-validates shape).
            ref: z.string().min(1).max(8_192),
            // Legacy entries without an explicit scope remain workbook names.
            // Worksheet names carry the exact persisted sheet identity so a
            // local name may truthfully shadow a workbook name (or the same
            // local name on another worksheet) without name-only collapse.
            scope: z.enum(["workbook", "worksheet"]).optional(),
            scopeSheetId: z.string().min(1).max(255).optional(),
            nativeBlocker: z
              .enum([
                "external-reference",
                "unsupported-attributes",
                "ambiguous-identity",
              ])
              .optional(),
            nativeSourceName: z.string().min(1).max(255).optional(),
            nativeSourceScopeSheetId: z
              .string()
              .min(1)
              .max(255)
              .optional(),
            // Excel's hidden flag: out of Name Manager, still in formulas.
            hidden: z.boolean().optional(),
            // Name Manager's Comment column (XLS-A066), Excel's 255 cap.
            comment: z.string().max(255).optional(),
          })
          .passthrough()
          .superRefine((namedRange, context) => {
            const scope = namedRange.scope ?? "workbook";
            if (scope === "worksheet" && !namedRange.scopeSheetId) {
              context.addIssue({
                code: z.ZodIssueCode.custom,
                path: ["scopeSheetId"],
                message: "worksheet named ranges require scopeSheetId",
              });
            }
            if (scope === "workbook" && namedRange.scopeSheetId !== undefined) {
              context.addIssue({
                code: z.ZodIssueCode.custom,
                path: ["scopeSheetId"],
                message: "workbook named ranges cannot carry scopeSheetId",
              });
            }
            if (
              namedRange.nativeSourceScopeSheetId !== undefined &&
              (scope !== "worksheet" || !namedRange.nativeSourceName)
            ) {
              context.addIssue({
                code: z.ZodIssueCode.custom,
                path: ["nativeSourceScopeSheetId"],
                message:
                  "native source sheet identity requires a worksheet-scoped source name",
              });
            }
            if (
              scope === "workbook" &&
              namedRange.nativeSourceScopeSheetId !== undefined
            ) {
              context.addIssue({
                code: z.ZodIssueCode.custom,
                path: ["nativeSourceScopeSheetId"],
                message:
                  "workbook named ranges cannot carry native source sheet identity",
              });
            }
          }),
      )
      .max(1_000)
      .optional(),
    // Recorded MACROS (Excel/WPS "Record Macro" — imperative automation; the
    // functional half is named LAMBDAs). Each macro is a name + a list of cell
    // edits RELATIVE to the recording anchor, replayable at any target. Only the
    // recipe is persisted; replaying writes through the normal edit path (undo +
    // protection). Bounded so a crafted file can't build a huge replay table.
    // Optional → existing sheets validate BYTE-FOR-BYTE. See lib/spreadsheet/macros.
    // A step is either the v1 entry shape ({dRow, dCol, value}) or a kinded
    // step (cells / lines / size / hidden / merges / freeze / sort / run ...),
    // all normalized by lib/spreadsheet/macros on read; the schema bounds the
    // list and the entry text only.
    macros: z
      .array(
        z
          .object({
            name: z.string().min(1).max(255),
            steps: z
              .array(
                z
                  .object({
                    kind: z.string().max(32).optional(),
                    dRow: z.number().int().optional(),
                    dCol: z.number().int().optional(),
                    value: z.string().max(32_000).optional(),
                  })
                  .passthrough(),
              )
              .max(5_000),
            mode: z.enum(["absolute", "relative"]).optional(),
            anchorRow: z.number().int().min(0).optional(),
            anchorCol: z.number().int().min(0).optional(),
            shortcut: z
              .object({ key: z.string().regex(/^[a-z]$/), shift: z.boolean().optional() })
              .optional(),
            description: z.string().max(1_000).optional(),
          })
          .passthrough(),
      )
      .max(100)
      .optional(),
    // Scenarios (Excel/WPS "Data ▸ What-If Analysis ▸ Scenario Manager"): each is
    // a NAMED set of {cell → value} overrides the user can save and re-apply to
    // compare what-if outcomes. Only the recipe (name + per-cell literal values)
    // is persisted — applying a scenario writes the values into the grid through
    // the normal edit path so dependent formulas recompute. Each `cells[].ref`
    // is a single A1 cell ("B2") that is REMAPPED on row/col insert/delete in
    // lib/spreadsheet/structural (a ref on a deleted line is dropped from that
    // scenario, the same as a deleted cell). Optional so existing sheets without
    // scenarios validate BYTE-FOR-BYTE; absent/empty means "none". The scenario
    // count, per-scenario cell count, and per-field lengths are bounded so a
    // malicious file cannot build a huge table (DoS on iteration/apply). See
    // lib/spreadsheet/scenarios for the capture/apply/normalize math (extra
    // fields pass through; the runtime re-normalizes via normalizeScenarios).
    scenarios: z
      .array(
        z
          .object({
            name: z.string().min(1).max(255),
            cells: z
              .array(
                z
                  .object({
                    ref: z.string().min(1).max(64),
                    value: z.string().max(32_767),
                  })
                  .passthrough(),
              )
              .max(256),
          })
          .passthrough(),
      )
      .max(256)
      .optional(),
    // What-If Data Tables (Data ▸ What-If Analysis ▸ Data Table): each keeps
    // its result block (`ref`, A1) and its row and/or column input cell; the
    // results hold their last values as ordinary cells and every calculation
    // recomputes them (Excel's {=TABLE(row,col)}). Optional so sheets without
    // one validate byte-for-byte; the runtime re-checks every entry and drops
    // a malformed one (lib/spreadsheet/what-if-table-definition).
    whatIfTables: z
      .array(
        z
          .object({
            ref: z.string().min(1).max(40),
            rowInput: z.string().min(1).max(16).optional(),
            colInput: z.string().min(1).max(16).optional(),
          })
          .strict(),
      )
      .max(64)
      .optional(),
    // Data validation rules (Excel/WPS parity): per-range constraints on what a
    // user may type into a cell. Each entry has a kind, an inclusive 0-based
    // range rectangle, and kind-specific params:
    //  - "list":      values[] (a fixed dropdown set).
    //  - "number":    min/max/integer (a numeric range).
    //  - "text":      minLength/maxLength (a character-count range).
    //  - "formula":   a per-cell predicate formula (Excel "Custom"); truthy ⇒ ok.
    //  - "date":      min/max ISO dates + optional weekdays / before-after-today.
    //  - "time":      min/max 24-hour times.
    //  - "dependent": a cascade dropdown whose options derive from another cell.
    // Optional so existing sheets without any validation still validate; absent/
    // empty means "no validation". The kind enum carries the three ADDITIVE
    // kinds; all new fields are optional and the object passes extra fields
    // through, so a persisted list/number/text rule validates byte-for-byte as
    // before. The array length, per-list value count and per-value length are
    // bounded so a malicious file cannot build a huge rule/value table (DoS on
    // iteration/lookup). See lib/spreadsheet/data-validation for the
    // validate/lookup/normalize math; the runtime re-normalizes via
    // normalizeRules so any out-of-shape field is dropped before use.
    validations: z
      .array(
        z
          .object({
            kind: z.enum([
              "any",
              "list",
              "number",
              "text",
              "formula",
              "date",
              "time",
              "dependent",
            ]),
            range: z
              .object({
                r1: z.number().int().min(0).max(MAX_GRID_ROWS),
                c1: z.number().int().min(0).max(MAX_GRID_COLUMNS),
                r2: z.number().int().min(0).max(MAX_GRID_ROWS),
                c2: z.number().int().min(0).max(MAX_GRID_COLUMNS),
              })
              .passthrough(),
            values: z.array(z.string().max(1_000)).max(1_000).optional(),
            sourceRange: z.string().max(255).optional(),
            showDropDown: z.boolean().optional(),
            // number: numeric bounds; date reuses min/max as ISO date STRINGS,
            // and time uses HH:MM[:SS], so accept either a number or a bounded
            // string here.
            min: z.union([z.number(), z.string().max(32)]).optional(),
            max: z.union([z.number(), z.string().max(32)]).optional(),
            operator: z
              .enum([
                "between",
                "notBetween",
                "equal",
                "notEqual",
                "greaterThan",
                "lessThan",
                "greaterThanOrEqual",
                "lessThanOrEqual",
              ])
              .optional(),
            integer: z.boolean().optional(),
            minLength: z.number().int().min(0).max(32_767).optional(),
            maxLength: z.number().int().min(0).max(32_767).optional(),
            // formula ("Custom"): a bounded per-cell predicate formula.
            formula: z.string().min(1).max(1_000).optional(),
            // date: optional weekday allow-list (0=Sun..6=Sat) + today bounds.
            weekdays: z.array(z.number().int().min(0).max(6)).max(7).optional(),
            notBefore: z.literal("today").optional(),
            notAfter: z.literal("today").optional(),
            // dependent (cascade): the controlling cell, how it is read, the
            // value→options map and a fallback option list.
            source: z
              .object({
                row: z.number().int().min(0).max(MAX_GRID_ROWS),
                col: z.number().int().min(0).max(MAX_GRID_COLUMNS),
              })
              .passthrough()
              .optional(),
            sourceMode: z.enum(["fixed", "perRow"]).optional(),
            options: z
              .record(z.array(z.string().max(1_000)).max(1_000))
              .optional(),
            optionsDefault: z
              .array(z.string().max(1_000))
              .max(1_000)
              .optional(),
            // Optional Excel-style INPUT MESSAGE (shown when a covered cell is
            // selected) and ERROR ALERT (title + body + style shown on an invalid
            // entry). All additive/optional so a persisted rule without any of
            // them validates byte-for-byte as before; `errorStyle` chooses
            // whether an invalid entry is blocked ("stop") or may be overridden
            // ("warning"/"information"). Lengths are bounded (title 255, body
            // 2000) so a hand-edited file cannot bloat the payload. The runtime
            // re-normalizes via normalizeRules/normalizeMessages.
            ignoreBlank: z.boolean().optional(),
            showInputMessage: z.boolean().optional(),
            showErrorMessage: z.boolean().optional(),
            inputTitle: z.string().max(255).optional(),
            inputMessage: z.string().max(2_000).optional(),
            errorTitle: z.string().max(255).optional(),
            errorMessage: z.string().max(2_000).optional(),
            errorStyle: z.enum(["stop", "warning", "information"]).optional(),
          })
          .passthrough(),
      )
      // MAX_VALIDATION_RULES (data-validation.ts): one a rectangle, as the
      // import splits a rule's sqref (xlsx-rounds R422-X46: 1,024 of one rule).
      .max(10_000)
      .optional(),
    // Conditional formatting (Excel/WPS parity): per-range rules that restyle
    // covered cells. A discriminated union on `kind` (additive — see
    // lib/spreadsheet/conditional-format):
    //  - comparison (DEFAULT; `kind` optional/absent so OLD data still
    //    validates): numeric operator + value + ≥1 hex color.
    //  - formula: a per-cell spreadsheet formula + ≥1 hex color.
    //  - colorScale: a 2-/3-point gradient (min/max + optional mid hex stops).
    //  - dataBar: a single hex bar color.
    //  - iconSet: a built-in icon set + two threshold fractions.
    //  - textContains: a text-match operator + needle + ≥1 hex color.
    //  - duplicateValues: optional `unique` flag + ≥1 hex color.
    //  - topBottom: a rank (+ optional bottom/percent flags) + ≥1 hex color.
    //  - aboveBelowAverage: optional below flag + optional stdDevs + ≥1 color.
    // Optional so existing sheets without any conditional formatting still
    // validate; absent/empty means "none". The array length is bounded so a
    // malicious file cannot build a huge rule table (DoS on iteration); every
    // color runs through normalizeHexColor (accepts #rgb / #rrggbb, stores the
    // expanded #rrggbb) so a persisted value can never inject arbitrary CSS at
    // render time. The comparison variant is listed FIRST so kind-less data
    // matches it (and serializes back byte-for-byte). Extra fields pass through;
    // the runtime re-normalizes via normalizeRules.
    conditionalFormats: z
      .array(
        z.union([
          // comparison (original shape; kind optional so old rules validate).
          z
            .object({
              kind: z.literal("comparison").optional(),
              range: conditionalRangeSchema,
              // Excel's two-threshold operators belong here too: the runtime
              // has always evaluated them (isRangeConditionalOperator reads
              // `value2`) and the server projects them from `cfRule/@operator`,
              // but leaving them out of this enum made the whole DOCUMENT fail
              // to parse, so a workbook with one "between" rule opened as a
              // single blank sheet. `value2` rides through on `.passthrough()`.
              operator: z.enum([">", "<", ">=", "<=", "=", "!=", "between", "notBetween"]),
              value: z.number(),
              value2: z.number().optional(),
              textColor: hexColorSchema.optional(),
              backgroundColor: hexColorSchema.optional(),
            })
            .passthrough(),
          // formula: a bounded per-cell predicate formula + ≥1 color.
          z
            .object({
              kind: z.literal("formula"),
              range: conditionalRangeSchema,
              formula: z.string().min(1).max(1_000),
              textColor: hexColorSchema.optional(),
              backgroundColor: hexColorSchema.optional(),
            })
            .passthrough(),
          // colorScale: 2-/3-point gradient stops.
          z
            .object({
              kind: z.literal("colorScale"),
              range: conditionalRangeSchema,
              min: z.object({ color: hexColorSchema }).passthrough(),
              // The midpoint takes every type an end does: Excel's
              // Formula midpoint (cfvo num/formula holding a reference,
              // S26_colorscale_num_with_ref, X-181) failed this enum and
              // set the whole rule aside.
              mid: z
                .object({
                  color: hexColorSchema,
                  type: z.enum(["percent", "percentile", "number", "formula", "min", "max"]).optional(),
                  value: z.number().finite().optional(),
                  formula: z.string().min(1).max(1_000).optional(),
                })
                .passthrough()
                .optional(),
              max: z.object({ color: hexColorSchema }).passthrough(),
            })
            .passthrough(),
          // dataBar: a single bar fill color.
          z
            .object({
              kind: z.literal("dataBar"),
              range: conditionalRangeSchema,
              color: hexColorSchema,
              minLength: z.number().int().min(0).max(100).optional(),
              maxLength: z.number().int().min(0).max(100).optional(),
              showValue: z.boolean().optional(),
              gradient: z.boolean().optional(),
              nativeDataBarId: z
                .string()
                .regex(/^\{[0-9A-Fa-f-]{36}\}$/)
                .optional(),
            })
            .passthrough(),
          // iconSet: one of the built-in icon-set families (3/4/5-icon, the full
          // Excel set; see ICON_SET_FAMILIES) + that family's (count-1) threshold
          // fractions in [0,1] + optional reverse / show-icon-only flags. The
          // thresholds array is accepted at 1–4 entries (a 5-icon set needs 4)
          // and the runtime normalizer (normalizeRules) re-derives the exact
          // count-1 length, sorts, and clamps; the schema only needs to admit a
          // valid new-family rule rather than drop it on load.
          z
            .object({
              kind: z.literal("iconSet"),
              range: conditionalRangeSchema,
              iconSet: z.enum(iconSetNameEnum),
              thresholds: z
                .array(z.number().min(0).max(1))
                .min(1)
                .max(4)
                .optional(),
              reverse: z.boolean().optional(),
              showIconOnly: z.boolean().optional(),
            })
            .passthrough(),
          // textContains: a text-match operator + non-empty bounded needle.
          z
            .object({
              kind: z.literal("textContains"),
              range: conditionalRangeSchema,
              text: z.string().min(1).max(1_000),
              // The rule's own formula (another program's, testing a cell
              // away from the one it paints), which Excel evaluates.
              formula: z.string().min(1).max(1_000).optional(),
              operator: z.enum([
                "contains",
                "notContains",
                "beginsWith",
                "endsWith",
              ]),
              textColor: hexColorSchema.optional(),
              backgroundColor: hexColorSchema.optional(),
            })
            .passthrough(),
          // duplicateValues: optional `unique` flag (true → highlight uniques).
          z
            .object({
              kind: z.literal("duplicateValues"),
              range: conditionalRangeSchema,
              unique: z.boolean().optional(),
              textColor: hexColorSchema.optional(),
              backgroundColor: hexColorSchema.optional(),
            })
            .passthrough(),
          // topBottom: a rank (≥1) + optional bottom / percent flags.
          z
            .object({
              kind: z.literal("topBottom"),
              range: conditionalRangeSchema,
              rank: z.number(),
              bottom: z.boolean().optional(),
              percent: z.boolean().optional(),
              textColor: hexColorSchema.optional(),
              backgroundColor: hexColorSchema.optional(),
            })
            .passthrough(),
          // aboveBelowAverage: optional below flag + optional std-dev offset.
          z
            .object({
              kind: z.literal("aboveBelowAverage"),
              range: conditionalRangeSchema,
              below: z.boolean().optional(),
              stdDevs: z.number().optional(),
              textColor: hexColorSchema.optional(),
              backgroundColor: hexColorSchema.optional(),
            })
            .passthrough(),
          // dateOccurring ("A Date Occurring…"): a relative period + ≥1 color.
          z
            .object({
              kind: z.literal("dateOccurring"),
              range: conditionalRangeSchema,
              period: z.enum([
                "yesterday",
                "today",
                "tomorrow",
                "last7Days",
                "lastWeek",
                "thisWeek",
                "nextWeek",
                "lastMonth",
                "thisMonth",
                "nextMonth",
              ]),
              textColor: hexColorSchema.optional(),
              backgroundColor: hexColorSchema.optional(),
            })
            .passthrough(),
        ]),
      )
      .max(spreadsheetPageEnvelope.maxConditionalFormatsPerSheet)
      .optional(),
    /**
     * Preserve-only receipt for an imported conditional-format graph whose
     * classic OOXML body is not wholly owned by the editor (x14/ext rules,
     * unknown attributes/elements, multi-area sqref, or partial dxf styles).
     * The Manage Rules writer must fail closed while this receipt is present.
     */
    conditionalFormatsNativeBlocker: z.string().min(1).max(220).optional(),
    // Pivot table (Excel/WPS parity): an OPTIONAL, additive recipe that
    // recomputes a read-only pivot grid from a source range on every load.
    // Absent means "no pivot". Only the RECIPE is persisted (source A1 range +
    // which source columns are row/column/value fields + each value field's
    // aggregation) — never the computed grid, so the pivot always reflects the
    // current data. Field indices and the per-axis array lengths are bounded so
    // a hand-edited file cannot build a huge field list (DoS on iteration); the
    // value-field aggregation is restricted to the five summary kinds. See
    // lib/spreadsheet/pivot for the compute/normalize math (extra fields pass
    // through; the runtime re-clamps via normalizePivotConfig).
    pivot: spreadsheetSheetPivotSchema.optional(),
    morePivots: z.array(spreadsheetSheetPivotSchema).max(63).optional(),
    // Native PivotTable timelines (Excel "Timeline", 日程表): imported views
    // whose settings are read-only and whose use-state (selection, time
    // level, scroll position) the editor writes together with the connected
    // PivotTables' `dateFilters`. See lib/spreadsheet/pivot-timeline.
    timelines: z
      .array(
        z
          .object({
            id: z.string().min(1).max(255),
            cacheName: z.string().min(1).max(255),
            sourceName: z.string().min(1).max(255),
            level: z.enum(["years", "quarters", "months", "days"]),
            selectionLevel: z
              .enum(["years", "quarters", "months", "days"])
              .optional(),
            bounds: z
              .object({
                start: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
                end: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
              })
              .strict(),
            selection: z
              .object({
                start: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
                end: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
              })
              .strict()
              .optional(),
            connections: z
              .array(
                z
                  .object({
                    sheetName: z.string().min(1).max(255),
                    pivotName: z.string().min(1).max(255),
                  })
                  .strict(),
              )
              .max(64),
          })
          .passthrough(),
      )
      .max(64)
      .optional(),
    // Native range-bound charts authored by the Spreadsheet surface. Imported
    // drawing charts use `objects[type=chart]`; both share the exact bounded
    // series-ref wire above when explicit native refs are present.
    charts: z.array(spreadsheetFloatingChartSchema).max(64).optional(),
    objects: z
      .array(
        z.union([
          spreadsheetImageObjectSchema,
          spreadsheetChartObjectSchema,
          spreadsheetShapeObjectSchema,
          spreadsheetSmartArtObjectSchema,
        ]),
      )
      // A bound on work, not the limit that binds: the objects' share of the
      // page is measured in bytes (lib/spreadsheet/object-budget.ts).
      .max(SPREADSHEET_MAX_OBJECTS)
      .optional(),
    // Sparklines (Excel/WPS parity): tiny in-cell charts keyed by the cell's
    // "row:col" ref (the same key shape as `cells`). Each entry is a chart kind
    // ("line"/"column"/"winloss"), an A1-style data range string, and optional
    // hex colors. A cell with a sparkline renders the mini chart instead of its
    // text (the underlying value is never deleted). Optional so existing sheets
    // without sparklines validate byte-for-byte; absent/empty means "none". The
    // map size and per-field lengths are bounded so a malicious file cannot
    // build a huge table (DoS on iteration); colors are re-validated through
    // normalizeSparklineColor at the runtime edge so a persisted value can never
    // inject arbitrary CSS. See lib/spreadsheet/sparkline for the normalize/
    // geometry math (extra fields pass through; the runtime re-clamps via
    // normalizeSparklines).
    sparklines: z
      .record(
        z
          .object({
            type: z.enum(["line", "column", "winloss"]),
            // Empty is a real Excel shape, not corruption: a sparkline group
            // whose `x14:sparkline` carries no `xm:f` member projects with no
            // range, and normalizeSparkline already treats that as "no data".
            // Requiring one character rejected the whole DOCUMENT, so such a
            // workbook opened as a single blank sheet.
            range: z.string().max(64),
            color: z.string().max(9).optional(),
            negativeColor: z.string().max(9).optional(),
            lineWeight: z.number().min(0.1).max(20).optional(),
          })
          .passthrough(),
      )
      .optional(),
    nativeSparklinesBlocker: z
      .literal(SPARKLINE_NATIVE_GRAPH_BLOCKER)
      .optional(),
    // Slicers (Excel/WPS parity): floating point-and-click filter panels bound
    // to one stable Table/Pivot field. `column` remains the worksheet placement
    // coordinate and legacy fallback, while `source` names the native filter
    // authority that export/patchback/reload must preserve. Each entry carries
    // an optional title, an OPTIONAL advisory copy of the chosen values
    // (empty/absent = all; the source filter wins), and optional panel geometry.
    // Optional so existing sheets without slicers validate byte-for-byte;
    // absent/empty means "none". The array length, per-field lengths, and the
    // cached selection size are bounded so a malicious file cannot build a huge
    // table (DoS on iteration). See lib/spreadsheet/slicer for the normalize +
    // selection<->filter math (extra fields pass through; the runtime re-clamps
    // via normalizeSlicers and drops a slicer on an out-of-range column).
    slicers: z
      .array(
        z
          .object({
            id: z.string().min(1).max(200),
            column: z.number().int().min(0).max(MAX_GRID_COLUMNS),
            source: z
              .discriminatedUnion("kind", [
                z.object({
                  kind: z.literal("table"),
                  tableId: z.string().min(1).max(200),
                  tableColumnId: z.number().int().positive(),
                }),
                z.object({
                  kind: z.literal("pivot"),
                  pivotOutput: z.string().min(1).max(64),
                  fieldId: z.string().min(1).max(200),
                  pivotName: z.string().min(1).max(255).optional(),
                  // Report Connections beyond the first PivotTable.
                  connections: z
                    .array(
                      z
                        .object({
                          pivotOutput: z.string().min(1).max(64),
                          pivotName: z.string().min(1).max(255),
                        })
                        .strict(),
                    )
                    .max(63)
                    .optional(),
                }),
              ])
              .optional(),
            title: z.string().max(200).optional(),
            selected: z.array(z.string().max(1_000)).max(10_000).optional(),
            x: z.number().optional(),
            y: z.number().optional(),
          })
          .passthrough(),
      )
      .max(64)
      .optional(),
    // Watch Window (Excel Formulas ▸ Watch Window parity): a small list of cells
    // the user pins to a side panel so their live value stays visible while
    // editing far-away cells in a large sheet. Stores only the coordinates
    // (row/col); the value + formula are re-derived at render from the evaluator,
    // never persisted. Optional so existing sheets validate byte-for-byte; the
    // runtime re-clamps via normalizeWatches and drops out-of-range coordinates.
    watches: z
      .array(
        z
          .object({
            row: z.number().int().min(0).max(MAX_GRID_ROWS),
            col: z.number().int().min(0).max(MAX_GRID_COLUMNS),
          })
          .passthrough(),
      )
      .max(64)
      .optional(),
    // Structured data tables (Excel ListObject / WPS "Format as Table" parity):
    // a NAMED rectangular range whose first row is the header, with optional
    // banded-row striping and an optional totals row. A table is a STYLE +
    // STRUCTURE overlay over the existing grid cells — it never owns cell content
    // (the cells stay authoritative); only the recipe is persisted. Each entry
    // carries an id, a name, the range (r1/c1/r2/c2, header at r1), `headerRow`
    // plus stable native tableColumn ids/names, optional header/style/totals
    // flags, and a `totals` map
    // keyed by in-range column OFFSET to a per-column aggregation. Optional so
    // existing sheets without tables validate byte-for-byte; absent/empty means
    // "none". The array length and per-field lengths are bounded so a malicious
    // file cannot build a huge table list (DoS on iteration / overlay render).
    // The range is REMAPPED on row/col insert/delete in lib/spreadsheet/structural
    // (alongside validations). See lib/spreadsheet/data-table for the normalize +
    // role + aggregate math (extra fields pass through; the runtime re-clamps via
    // normalizeDataTables and clamps a stale range into the current grid).
    dataQueries: z.custom<DataQueryState>((value) => {
      try { parseDataQueryState(value); return true; } catch { return false; }
    }, "Invalid or unsupported local query history").optional(),
    dataTables: z
      .array(
        z
          .object({
            id: z.string().min(1).max(200),
            name: z.string().min(1).max(200),
            range: z
              .object({
                r1: z.number().int().min(0).max(MAX_GRID_ROWS),
                c1: z.number().int().min(0).max(MAX_GRID_COLUMNS),
                r2: z.number().int().min(0).max(MAX_GRID_ROWS),
                c2: z.number().int().min(0).max(MAX_GRID_COLUMNS),
              })
              .passthrough(),
            headerRow: z.boolean().optional(),
            showFilterButton: z.boolean().optional(),
            columns: z
              .array(
                z
                  .object({
                    id: z.number().int().min(1).max(2_147_483_647),
                    // Excel writes at most 255; a file's may say more (a form's
                    // question as a header, 362 characters, R127-X09). Refused
                    // here, the whole table list was set aside and the table no
                    // longer followed its rows.
                    name: z.string().min(1).max(32_767),
                  })
                  .passthrough(),
              )
              // A table may span every column of the sheet (R240-X11's Table1,
              // A1:XFD272). At the old grid's 1,000 the table list was set
              // aside on open and a row put in inside the table left it short
              
              .max(MAX_GRID_COLUMNS)
              .optional(),
            columnFilters: z
              .record(
                z
                  .object({
                    values: z
                      .array(z.string().max(1_000))
                      .min(1)
                      .max(10_000)
                      .optional(),
                    dateGroups: z
                      .array(
                        z.discriminatedUnion("grouping", [
                          z
                            .object({
                              grouping: z.literal("year"),
                              year: z.number().int().min(1).max(9_999),
                            })
                            .strict(),
                          z
                            .object({
                              grouping: z.literal("month"),
                              year: z.number().int().min(1).max(9_999),
                              month: z.number().int().min(1).max(12),
                            })
                            .strict(),
                          z
                            .object({
                              grouping: z.literal("day"),
                              year: z.number().int().min(1).max(9_999),
                              month: z.number().int().min(1).max(12),
                              day: z.number().int().min(1).max(31),
                            })
                            .strict(),
                          z
                            .object({
                              grouping: z.literal("hour"),
                              year: z.number().int().min(1).max(9_999),
                              month: z.number().int().min(1).max(12),
                              day: z.number().int().min(1).max(31),
                              hour: z.number().int().min(0).max(23),
                            })
                            .strict(),
                          z
                            .object({
                              grouping: z.literal("minute"),
                              year: z.number().int().min(1).max(9_999),
                              month: z.number().int().min(1).max(12),
                              day: z.number().int().min(1).max(31),
                              hour: z.number().int().min(0).max(23),
                              minute: z.number().int().min(0).max(59),
                            })
                            .strict(),
                          z
                            .object({
                              grouping: z.literal("second"),
                              year: z.number().int().min(1).max(9_999),
                              month: z.number().int().min(1).max(12),
                              day: z.number().int().min(1).max(31),
                              hour: z.number().int().min(0).max(23),
                              minute: z.number().int().min(0).max(59),
                              second: z.number().int().min(0).max(59),
                            })
                            .strict(),
                        ]),
                      )
                      .min(1)
                      .max(10_000)
                      .optional(),
                    calendarType: z.literal("gregorian").optional(),
                    condition: z
                      .union([
                        spreadsheetTextConditionFilterSchema,
                        spreadsheetNumberConditionFilterSchema,
                        spreadsheetCompoundConditionFilterSchema,
                      ])
                      .optional(),
                    color: z
                      .object({
                        kind: z.enum(["cell", "font"]),
                        value: z.string().regex(/^#[0-9a-fA-F]{6}$/),
                      })
                      .strict()
                      .optional(),
                    icon: z
                      .object({
                        iconSet: z.enum(DATA_TABLE_ICON_FILTER_SETS),
                        iconId: z.number().int().min(0).max(4),
                      })
                      .strict()
                      .refine(
                        (icon) => icon.iconId < iconSetCount(icon.iconSet),
                        { message: "table icon id exceeds its icon family" },
                      )
                      .optional(),
                    relativeDate: z
                      .object({
                        period: z.enum(DATA_TABLE_RELATIVE_DATE_PERIODS),
                        val: z.number().finite().optional(),
                        maxVal: z.number().finite().optional(),
                      })
                      .strict()
                      .refine(
                        (criterion) =>
                          criterion.maxVal === undefined ||
                          (criterion.val !== undefined &&
                            criterion.maxVal >= criterion.val),
                        {
                          message:
                            "relative date maxVal requires an ordered val cache",
                        },
                      )
                      .optional(),
                  })
                  .strict()
                  .refine(
                    (filter) =>
                      filter.condition !== undefined
                        ? filter.values === undefined &&
                          filter.dateGroups === undefined &&
                          filter.calendarType === undefined &&
                          filter.color === undefined &&
                          filter.icon === undefined &&
                          filter.relativeDate === undefined
                        : filter.color !== undefined
                          ? filter.values === undefined &&
                            filter.dateGroups === undefined &&
                            filter.calendarType === undefined &&
                            filter.icon === undefined &&
                            filter.relativeDate === undefined
                          : filter.icon !== undefined
                            ? filter.values === undefined &&
                              filter.dateGroups === undefined &&
                              filter.calendarType === undefined &&
                              filter.relativeDate === undefined
                            : filter.relativeDate !== undefined
                              ? filter.values === undefined &&
                                filter.dateGroups === undefined &&
                                filter.calendarType === undefined
                            : (filter.values !== undefined ||
                                filter.dateGroups !== undefined) &&
                              (filter.calendarType === undefined ||
                                filter.dateGroups !== undefined),
                    {
                      message:
                        "a table column filter must use one condition, color, icon, relative-date, or value/date-group mode",
                    },
                  ),
              )
              .refine(
                // One filter a column, and a table may span every column of
                
                (filters) =>
                  Object.keys(filters).length <= MAX_GRID_COLUMNS &&
                  Object.keys(filters).every(
                    (key) =>
                      /^(0|[1-9][0-9]*)$/.test(key) &&
                      Number(key) >= 1 &&
                      Number(key) <= 2_147_483_647,
                  ),
                {
                  message:
                    "table column filters must use bounded stable column ids",
                },
              )
              .optional(),
            sortState: z
              .object({
                conditions: z
                  .array(
                    z
                      .object({
                        columnId: z
                          .number()
                          .int()
                          .min(1)
                          .max(2_147_483_647),
                        descending: z.boolean().optional(),
                        color: z
                          .object({
                            kind: z.enum(["cell", "font"]),
                            value: z.string().regex(/^#[0-9a-fA-F]{6}$/),
                          })
                          .strict()
                          .optional(),
                        icon: z
                          .object({
                            iconSet: z.enum(DATA_TABLE_ICON_FILTER_SETS),
                            iconId: z.number().int().min(0).max(4),
                          })
                          .strict()
                          .refine(
                            (icon) => icon.iconId < iconSetCount(icon.iconSet),
                            { message: "table sort icon index is out of range" },
                          )
                          .optional(),
                      })
                      .strict()
                      .refine(
                        (condition) =>
                          condition.color === undefined ||
                          condition.icon === undefined,
                        {
                          message:
                            "table sort condition cannot mix color and icon modes",
                        },
                      ),
                  )
                  .min(1)
                  .max(64)
                  .refine(
                    (conditions) =>
                      new Set(conditions.map((condition) => condition.columnId))
                        .size === conditions.length,
                    { message: "table sort keys must use unique stable column ids" },
                  ),
              })
              .strict()
              .optional(),
            styleId: z.string().min(1).max(40).optional(),
            nativeStyleName: z.string().min(1).max(255).optional(),
            bandedRows: z.boolean().optional(),
            bandedColumns: z.boolean().optional(),
            firstColumn: z.boolean().optional(),
            lastColumn: z.boolean().optional(),
            totalsRow: z.boolean().optional(),
            totals: z
              .record(
                z.enum([
                  "none",
                  "average",
                  "count",
                  "countNums",
                  "max",
                  "min",
                  "sum",
                  "stdDev",
                  "var",
                  "custom",
                ]),
              )
              .optional(),
            totalFormulas: z
              .record(
                z
                  .string()
                  .min(2)
                  .max(8_192)
                  .regex(/^=/)
                  .refine(
                    (formula) =>
                      !/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(
                        formula,
                      ),
                    { message: "table total formula contains XML control characters" },
                  ),
              )
              .optional(),
            calculatedFormulas: z
              .record(
                z
                  .string()
                  .min(2)
                  .max(8_192)
                  .regex(/^=/)
                  .refine(
                    (formula) =>
                      !/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(
                        formula,
                      ),
                    { message: "calculated column formula contains XML control characters" },
                  ),
              )
              .optional(),
            arrayCalculatedColumns: z.record(z.literal(true)).optional(),
          })
          .passthrough(),
      )
      // The page's own bound: a sheet past 64 tables had every table set aside, unmoved by line edits (R312-X25).
      .max(MAX_DATA_TABLES)
      .optional(),
    // Print / Page Setup (Excel/WPS parity): per-sheet print settings that control
    // what is exported to PDF.  All fields are optional; absent means "use the
    // export default".  Schema mirrors the PrintSettings type in
    // lib/spreadsheet/print-settings.ts. Using .passthrough() so unknown future
    // sub-fields survive round-trips without schema rejection.
    printSettings: z
      .object({
        // A1-notation print area, e.g. "A1:F40", or Excel's comma-separated
        // list of areas ("A1:F40,H1:J5"; MAX_PRINT_AREA_RANGES blocks of at
        // most 22 characters each). Absent = whole sheet.
        printArea: z.string().min(1).max(1500).optional(),
        // Row range to repeat at top of every page, e.g. "1:2".
        repeatTitleRows: z.string().min(1).max(50).optional(),
        // Column range to repeat at left of every page, e.g. "A:B".
        repeatTitleCols: z.string().min(1).max(50).optional(),
        orientation: z.enum(["portrait", "landscape"]).optional(),
        // Margins: a preset name OR an object with top/right/bottom/left in
        // points (0–288 = 0–4 inches).
        margins: z
          .union([
            z.enum(["normal", "narrow", "wide"]),
            z.object({
              top: z.number().min(0).max(288),
              right: z.number().min(0).max(288),
              bottom: z.number().min(0).max(288),
              left: z.number().min(0).max(288),
            }),
          ])
          .optional(),
        // Scaling: percent, legacy total-pages, or native independent axes.
        scaling: z
          .union([
            z
              .object({
                kind: z.literal("percent"),
                percent: z.number().min(10).max(400),
              })
              .strict(),
            z
              .object({
                kind: z.literal("fitToWidth"),
                pages: z.number().int().min(1).max(100),
              })
              .strict(),
            z
              .object({
                kind: z.literal("fitToPages"),
                pages: z.number().int().min(1).max(100),
              })
              .strict(),
            z
              .object({
                kind: z.literal("fitToPages"),
                widthPages: z.number().int().min(0).max(100),
                heightPages: z.number().int().min(0).max(100),
              })
              .strict()
              .refine(
                ({ widthPages, heightPages }) =>
                  widthPages > 0 || heightPages > 0,
                { message: "At least one fit dimension must be bounded" },
              ),
          ])
          .optional(),
        // Device print resolution and the printer configuration part, carried
        // as custody so an imported workbook keeps them on export.
        printerResolution: z
          .object({
            // ST_UnsignedInt: real dpi or a DEVMODE constant (4294967292+).
            horizontalDpi: z.number().int().min(1).max(4_294_967_295).optional(),
            verticalDpi: z.number().int().min(1).max(4_294_967_295).optional(),
          })
          .strict()
          .optional(),
        printerSettingsPart: z
          .string()
          .min(1)
          .max(87_384)
          .regex(/^[A-Za-z0-9+/]+={0,2}$/)
          .optional(),
      })
      .passthrough()
      .optional(),
    // Runs of identically formatted blank cells (Excel `<c s>` records with no
    // value, and row/column formats projected onto the data range) stored as
    // A1 ranges instead of one `cells` entry each. Optional/additive: a sheet
    // that still holds per-cell blank entries stays valid and renders the same,
    // and an explicit `cells` entry always wins over a range.
    blankCellFormats: z
      .array(spreadsheetBlankCellFormatGroupSchema)
      .max(MAX_BLANK_CELL_FORMAT_GROUPS)
      .optional(),
    cells: z.record(spreadsheetCellSchema).superRefine((cells, ctx) => {
      for (const key of Object.keys(cells)) {
        if (!/^\d+:\d+$/.test(key)) {
          ctx.addIssue({
            code: z.ZodIssueCode.custom,
            path: [key],
            message: "Cell keys must use row:column coordinates",
          });
          continue;
        }
        const [row, col] = key.split(":").map(Number);
        if (
          !Number.isSafeInteger(row) ||
          !Number.isSafeInteger(col) ||
          row < 0 ||
          col < 0 ||
          row >= MAX_GRID_ROWS ||
          col >= MAX_GRID_COLUMNS
        ) {
          ctx.addIssue({
            code: z.ZodIssueCode.custom,
            path: [key],
            message:
              "Cell coordinate is outside the Excel worksheet address space",
          });
        }
      }
    }),
  })
  .passthrough()
  .superRefine((sheet, ctx) => {
    const nativeRows = normalizeNativeGrid(sheet.nativeGrid)?.rowMetrics;
    for (const [index, pixels] of Object.entries(sheet.rowHeights ?? {})) {
      if ((pixels < MIN_ROW_HEIGHT_PX || pixels > MAX_ROW_HEIGHT_PX) &&
          !nativeRowPixelPairMatches(nativeRows, index, pixels)) {
        ctx.addIssue({ code: z.ZodIssueCode.custom, path: ["rowHeights", index],
          message: "Row height outside interactive limits requires a matching native point projection" });
      }
    }
    const splitRow = sheet.splitRow ?? 0;
    const splitCol = sheet.splitCol ?? 0;
    const splitActive = splitRow > 0 || splitCol > 0;
    if (splitRow >= sheet.rowCount) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["splitRow"],
        message: "splitRow must leave a trailing worksheet row",
      });
    }
    if (splitCol >= sheet.columnCount) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["splitCol"],
        message: "splitCol must leave a trailing worksheet column",
      });
    }
    if (
      splitActive &&
      ((sheet.freezeRows ?? 0) > 0 || (sheet.freezeCols ?? 0) > 0)
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["splitPaneState"],
        message: "freeze and split pane authority are mutually exclusive",
      });
    }
    // Counts-only split records predate pane-selection custody. They remain a
    // valid legacy input and the window deterministically hydrates the missing
    // selection authority. An explicit state without an active split is never
    // meaningful and remains invalid.
    if (!splitActive && sheet.splitPaneState) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["splitPaneState"],
        message: "an inactive split cannot retain pane authority",
      });
    }
    if (splitActive && sheet.splitPaneState) {
      const expectedPanes: readonly string[] =
        splitRow > 0 && splitCol > 0
          ? ["topLeft", "topRight", "bottomLeft", "bottomRight"]
          : splitRow > 0
            ? ["topLeft", "bottomLeft"]
            : ["topLeft", "topRight"];
      if (!expectedPanes.includes(sheet.splitPaneState.activePane)) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["splitPaneState", "activePane"],
          message: "activePane must name a physical split pane",
        });
      }
      if (
        sheet.splitPaneState.selections.length !== expectedPanes.length ||
        sheet.splitPaneState.selections.some(
          (selection, index) => selection.pane !== expectedPanes[index],
        )
      ) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["splitPaneState", "selections"],
          message: "split selections must cover physical panes in canonical order",
        });
      }
      const refs: Array<[Array<string | number>, string]> = [
        [["splitPaneState", "topLeftCell"], sheet.splitPaneState.topLeftCell],
      ];
      if (sheet.splitPaneState.leadingTopLeftCell !== undefined) {
        refs.push([
          ["splitPaneState", "leadingTopLeftCell"],
          sheet.splitPaneState.leadingTopLeftCell,
        ]);
      }
      for (const [index, selection] of sheet.splitPaneState.selections.entries()) {
        refs.push([
          ["splitPaneState", "selections", index, "activeCell"],
          selection.activeCell,
        ]);
        const ranges = selection.sqref.split(" ");
        for (const [rangeIndex, range] of ranges.entries()) {
          refs.push([
            ["splitPaneState", "selections", index, "sqref", rangeIndex],
            range,
          ]);
        }
        const active = spreadsheetSplitRefBounds(selection.activeCell);
        const indexed = spreadsheetSplitRefBounds(
          ranges[selection.activeCellId] ?? "",
        );
        if (
          !active ||
          !indexed ||
          active.r1 < Math.min(indexed.r1, indexed.r2) ||
          active.r1 > Math.max(indexed.r1, indexed.r2) ||
          active.c1 < Math.min(indexed.c1, indexed.c2) ||
          active.c1 > Math.max(indexed.c1, indexed.c2)
        ) {
          ctx.addIssue({
            code: z.ZodIssueCode.custom,
            path: ["splitPaneState", "selections", index, "activeCellId"],
            message: "the indexed sqref range must contain activeCell",
          });
        }
      }
      for (const [path, ref] of refs) {
        const bounds = spreadsheetSplitRefBounds(ref);
        if (
          !bounds ||
          bounds.r1 < 0 ||
          bounds.c1 < 0 ||
          bounds.r2 >= sheet.rowCount ||
          bounds.c2 >= sheet.columnCount
        ) {
          ctx.addIssue({
            code: z.ZodIssueCode.custom,
            path,
            message: "split pane reference exceeds the final worksheet grid",
          });
        }
      }
    }
    for (const [index, namedRange] of (sheet.namedRanges ?? []).entries()) {
      if (
        namedRange.scope === "worksheet" &&
        namedRange.scopeSheetId !== sheet.id
      ) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["namedRanges", index, "scopeSheetId"],
          message: "worksheet named range scope must match its containing sheet",
        });
      }
    }
    if (sheet.autoFilterId && !sheet.autoFilterRef) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["autoFilterRef"],
        message: "worksheet AutoFilter ref is required with its native uid",
      });
    }
    if (
      sheet.namedSheetViewsOwnsAutoFilter &&
      (!sheet.autoFilterId ||
        !sheet.autoFilterRef ||
        !(sheet.namedSheetViews ?? []).some(
          (view) =>
            view.filterId?.toUpperCase() === sheet.autoFilterId?.toUpperCase(),
        ))
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["namedSheetViewsOwnsAutoFilter"],
        message:
          "a Sheet Views-owned AutoFilter requires its uid, ref, and at least one view",
      });
    }
    for (const key of Object.keys(sheet.cells)) {
      const [row, col] = key.split(":").map((part) => Number(part));
      if (row >= sheet.rowCount || col >= sheet.columnCount) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["cells", key],
          message: "Cell is outside the declared sheet bounds",
        });
      }
    }
    const blankCellFormatsProblem = blankCellFormatsIssue(
      sheet.blankCellFormats,
      sheet.rowCount,
      sheet.columnCount,
    );
    if (blankCellFormatsProblem) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["blankCellFormats", ...blankCellFormatsProblem.path],
        message: blankCellFormatsProblem.message,
      });
    }
    for (const [index, view] of (sheet.namedSheetViews ?? []).entries()) {
      if (!view.filterId) continue;
      if (
        !sheet.autoFilterId ||
        view.filterId.toUpperCase() !== sheet.autoFilterId.toUpperCase()
      ) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["namedSheetViews", index, "filterId"],
          message: "named Sheet View filterId must match worksheet AutoFilter uid",
        });
      }
      if (
        view.ref !== undefined &&
        (!sheet.autoFilterRef ||
          view.ref.replace(/\$/g, "").toUpperCase() !==
            sheet.autoFilterRef.replace(/\$/g, "").toUpperCase())
      ) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["namedSheetViews", index, "ref"],
          message: "named Sheet View ref must match worksheet AutoFilter ref",
        });
      }
    }
  });

export const spreadsheetPageContentSchema = z
  .object({
    type: z.literal(SPREADSHEET_CONTENT_TYPE),
    // Mirrors the API's MAX_SPREADSHEET_WORKBOOK_SHEETS_TOTAL (the envelope's
    // maxSheetsTotal) so a workbook the server accepts is never rejected
    // client-side: a workbook imported with more sheets than a page adds sheets
    // up to (maxSheets, 200) keeps them all (xlsx-rounds R283-X12, 260 sheets).
    sheets: z.array(spreadsheetSheetSchema).min(1).max(spreadsheetPageEnvelope.maxSheetsTotal),
    // A save's own mark, never stored (the API takes it off): the window that
    // built it still held a sheet as its placeholder (deferred-sheets), so the
    // formulas it did not edit keep the file's values.
    sheetsUnfilled: z.literal(true).optional(),
    // Cache snapshots shared by several PivotTables, keyed by the cache part
    // (`native.cacheKey`). One PivotTable's cache stays on that PivotTable.
    // Stored compactly — one string per value, its OOXML kind then its lexical
    // value ("n12.5", "sBaden", "m") — since the display matrix follows from
    // the values (lib/spreadsheet/pivot `expandSharedPivotCache`).
    pivotCaches: z
      .record(
        z.string().min(1).max(512),
        z
          .object({
            revision: z.number().int().min(0).max(1_000_000_000),
            fingerprint: z.string().min(1).max(64),
            fieldIds: z.array(z.string().min(1).max(512)).min(1).max(SPREADSHEET_PIVOT_NATIVE_MAX_FIELDS),
            records: z
              .array(
                z.array(z.string().regex(/^[sndbem]/).max(2_001)).min(1).max(SPREADSHEET_PIVOT_NATIVE_MAX_FIELDS),
              )
              .min(1)
              .max(SPREADSHEET_PIVOT_SNAPSHOT_MAX_CELLS),
          })
          .strict()
          // The cache's cells are its budget, as a report's snapshot's are.
          .refine(
            (cache) => cache.records.length * cache.fieldIds.length <= SPREADSHEET_PIVOT_SNAPSHOT_MAX_CELLS,
            { message: "a pivot cache holds at most 200,000 cells" },
          ),
      )
      .optional(),
    activeSheetId: z.string().min(1).max(200).optional(),
    // File > Info > Protect Workbook > Encrypt with Password: the workbook's
    // files are written encrypted (lib/office/office-encryption). The password
    // itself is never stored.
    encryptWithPassword: z.literal(true).optional(),
    workbookStructureProtection: z
      .object({
        locked: z.boolean(),
        // Protect Workbook > Windows, as an imported file has it.
        windowsLocked: z.literal(true).optional(),
        passwordKind: z.enum(["none", "legacy", "modern", "unsupported"]),
        legacyVerifier: z.string().regex(/^[0-9A-F]{4}$/).optional(),
        passwordHash: protectionPasswordHashSchema.optional(),
        nativeBlocker: z.string().min(1).max(200).optional(),
      })
      .strict()
      .superRefine((protection, ctx) => {
        if (
          (protection.passwordKind === "legacy") !==
          (protection.legacyVerifier !== undefined)
        ) {
          ctx.addIssue({
            code: z.ZodIssueCode.custom,
            path: ["legacyVerifier"],
            message: "Legacy workbook protection requires exactly one verifier",
          });
        }
        if (
          (protection.passwordKind === "unsupported") !==
          (protection.nativeBlocker !== undefined)
        ) {
          ctx.addIssue({
            code: z.ZodIssueCode.custom,
            path: ["nativeBlocker"],
            message: "Unsupported workbook protection requires a native blocker",
          });
        }
        if (
          (protection.passwordKind === "modern") !==
            (protection.passwordHash !== undefined) ||
          (protection.passwordKind === "modern" &&
            !protection.locked &&
            protection.windowsLocked !== true)
        ) {
          ctx.addIssue({
            code: z.ZodIssueCode.custom,
            path: ["passwordHash"],
            message: "Modern workbook protection requires exactly one hash",
          });
        }
      })
      .optional(),
    // Workbook calculation mode. Automatic is omitted by the editor because it
    // is the native default; Manual is explicit and survives page/XLSX reload.
    calculationMode: z.enum(["auto", "manual"]).optional(),
    // The workbook's date system (Excel's "Use 1904 date system",
    // SpreadsheetML workbookPr@date1904). It decides how serials READ as
    // dates; the cells keep the serials the file wrote. Absent is the 1900
    // system; "1900" is written only by a page switched back from 1904, so an
    // original-package save knows its serials are already the page's own.
    dateSystem: z.enum(["1900", "1904", "1900-noncompat"]).optional(),
    // What the SERVER projection could not carry, in the browser importer's
    // own codes (`editable_page._xlsx_content`). The browser chain shows its
    // manifest before the user consents; this path — import batches and
    // attachments — has already imported, so the page carries the record and
    // its status bar reads it (`workbook-import-losses.ts` holds the words).
    // Bounded here because content_json is user-supplied at this boundary.
    importLosses: z
      .array(
        z
          .object({
            code: z.string().min(1).max(200),
            omitted: z.number().int().min(0).max(1_000_000),
          })
          .strict(),
      )
      .max(200)
      .optional(),
  })
  .passthrough()
  .superRefine((content, ctx) => {
    if (
      content.activeSheetId &&
      !content.sheets.some((sheet) => sheet.id === content.activeSheetId)
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["activeSheetId"],
        message: "Active sheet id must reference an existing sheet",
      });
    }
    // Covered blanks are materialized when a workbook opens, so their total
    // is an allocation budget across sheets, not only within one.
    const blankCellArea = content.sheets.reduce(
      (total, sheet) => total + blankCellFormatsArea(sheet.blankCellFormats),
      0,
    );
    if (blankCellArea > MAX_BLANK_CELL_FORMAT_WORKBOOK_AREA) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["sheets"],
        message: `blankCellFormats may cover at most ${MAX_BLANK_CELL_FORMAT_WORKBOOK_AREA} cells per workbook`,
      });
    }
  });

// Optional per-block present-mode entrance animation. Constrained to the shared
// allowed set so an existing deck without this field still validates and a stale
// value can never inject an unsupported animation. Purely additive — shared by
// every presentation block kind below.
export const presentationBlockAnimationSchema = z.enum(BLOCK_ANIMATIONS);

// Optional per-block hyperlink, validated against two target kinds: an external
// URL via the shared scheme whitelist (http/https/mailto only) OR an internal
// `slide://<slideId>` link for in-app navigation (strict id shape). Anything
// else — `javascript:`/`data:`/`vbscript:`/`file:` etc. — is rejected so a
// hand-edited deck can never persist an href that would render as a clickable
// link. See `isSafeBlockHref` / `sanitizeBlockHref` in presentation-block-fx.
export const presentationBlockHrefSchema = z
  .string()
  .refine(isSafeBlockHref, { message: "Unsupported or unsafe link scheme" });

// A link's ScreenTip (`a:hlinkClick@tooltip`, PPT-B010): plain text, bounded.
export const presentationLinkTooltipSchema = z
  .string()
  .refine((value) => normalizeLinkTooltip(value) === value, {
    message: "ScreenTip must be plain text",
  });

export const presentationRelativeClickActionSchema = z
  .object({
    kind: z.literal("relative-slide"),
    target: z.enum(PRESENTATION_RELATIVE_CLICK_ACTION_TARGETS),
  })
  .strict();

export const presentationMediaClickActionSchema = z
  .object({
    kind: z.literal("play-media"),
    mediaKind: z.enum(PRESENTATION_MEDIA_CLICK_ACTION_KINDS),
    targetBlockId: z.string().regex(/^[A-Za-z0-9_-]{1,200}$/),
  })
  .strict();

export const presentationClickActionSchema = z.union([
  presentationRelativeClickActionSchema,
  presentationMediaClickActionSchema,
]);

// Optional persistent group id shared by blocks moved/aligned together. Purely
// additive — existing decks omit it and still validate; selection grouping is a
// non-breaking, render-irrelevant attribute. Bounded length so a hand-edited
// value can't bloat the payload.
export const presentationGroupIdSchema = z.string().min(1).max(200);

// Optional signed source-space keep rectangle. Free crop occupies the 0..1
// subset; native PowerPoint Fit uses negative x/y and w/h > 1. The shared
// Presentation normalizer keeps this schema and the renderer on one bounded
// finite-overlap contract.
export const presentationCropRectSchema = z
  .object({
    x: z.number(),
    y: z.number(),
    w: z.number(),
    h: z.number(),
  })
  .passthrough()
  .refine((value) => normalizeSignedCropRect(value) !== null, {
    message: "Invalid signed Presentation crop keep rectangle",
  });

/** WordArt on a text box's or a shape's words: the shared text paint. */
export const presentationTextPaintSchema = z
  .unknown()
  .refine((value) => {
    try {
      return normalizeOfficeTextPaint(value) !== undefined;
    } catch {
      return false;
    }
  }, "Unsupported WordArt paint");

// Optional character-level (per-RUN) inline text formatting for a text block or
// one of its list paragraphs (PowerPoint parity). ADDITIVE — a block/paragraph
// written before this feature carries no `runs` at all and still validates (the
// field is `.optional()`); when absent the span renders exactly as the flat
// `text` with the block/paragraph style. When present each run is bounded here
// (text length, style subset) AND re-normalized at the editor + render edge by
// `normalizeRuns` (lib/presentation-runs.ts), which also enforces the invariant
// `textFromRuns(runs) === text` and collapses a flat-equivalent (single
// styleless run) to nothing. The bounds mirror the runtime caps
// (MAX_RUNS_PER_BLOCK / MAX_RUN_TEXT_LENGTH / MIN_RUN_FONT_SIZE..MAX_RUN_FONT_SIZE).
// `color` is the same strict #rrggbb hex used elsewhere; every style field is
// optional so a run with no overrides validates as plain text. A text box's
// run is bounded by its line (MAX_RUN_TEXT_LENGTH, PPT-B392); a shape label's
// keeps MAX_SHAPE_TEXT_LENGTH, as the server bounds each.
export const presentationRunsSchemaWithin = (maxTextLength: number) => z
  .array(
    z
      .object({
        text: z.string().max(maxTextLength),
        // The IDENTITY of the DrawingML text field (`a:fld`) this run came from
        // — a slide number, a date, a footer. ADDITIVE: a run written before
        // this feature carries none, and the text beside it is only
        // PowerPoint's cache of what the field last drew (it refreshes the
        // cache itself on open, docs/office-graphics/SHAPE-IMPORT-AUDIT.md
        // §17). Bounded to the importer's own grammar so a hand-edited deck
        // cannot smuggle an arbitrary attribute value into an exported package.
        field: z
          .object({
            type: z
              .string()
              .max(64)
              .regex(/^[A-Za-z][A-Za-z0-9._-]{0,63}$/),
            id: z
              .string()
              .regex(
                /^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$/,
              ),
          })
          .strict()
          .optional(),
        // The equation this run IS (Office Math `a14:m`, PPT-A052): its LaTeX
        // and whether it stands as its own line. ADDITIVE: the text beside it
        // is the characters the zone draws. Bounded as the server bounds it.
        math: z
          .object({
            latex: z.string().min(1).max(65_536),
            display: z.boolean().optional(),
          })
          .strict()
          .optional(),
        style: z
          .object({
            bold: z.boolean().optional(),
            italic: z.boolean().optional(),
            underline: z.boolean().optional(),
            // Display formats an imported run carries: a double underline
            // (`a:rPr@u="dbl"`, PPT-B363) and capitals (`a:rPr@cap`, PPT-B362).
            underlineStyle: z.literal("double").optional(),
            caps: z.enum(["all", "small"]).optional(),
            strike: z.boolean().optional(),
            // Baseline shifts (PowerPoint parity), mutually exclusive at the
            // normalize edge (`normalizeRunStyle` drops `subscript` when both
            // are set). Additive — a run without either omits both.
            superscript: z.boolean().optional(),
            subscript: z.boolean().optional(),
            proofingLanguage: z
              .string()
              .max(64)
              .regex(/^[A-Za-z]{2,8}(?:-[A-Za-z0-9]{1,8}){0,5}$/)
              .optional(),
            noProof: z.boolean().optional(),
            color: z
              .string()
              .regex(/^#[0-9a-fA-F]{6}$/)
              .optional(),
            // Text HIGHLIGHT (run background) colour — same strict #rrggbb hex
            // as `color` so a stale value can't inject arbitrary CSS. Additive.
            highlight: z
              .string()
              .regex(/^#[0-9a-fA-F]{6}$/)
              .optional(),
            fontSize: z
              .number()
              .min(MIN_RUN_FONT_SIZE)
              .max(MAX_RUN_FONT_SIZE)
              .optional(),
            fontFamily: z.string().max(200).optional(),
            // Character text OUTLINE (stroke) colour — same strict #rrggbb hex
            // (distinct from the `color` fill), rendered via -webkit-text-stroke.
            // Additive — older runs omit it (no stroke).
            textOutline: z
              .string()
              .regex(/^#[0-9a-fA-F]{6}$/)
              .optional(),
            href: z.union([
              presentationBlockHrefSchema,
              z.enum(["slide-action://first", "slide-action://last", "slide-action://next", "slide-action://previous"]),
            ]).optional(),
            linkTooltip: presentationLinkTooltipSchema.optional(),
            // The link draws in its text's colour (`ahyp:hlinkClr val="tx"`).
            linkColor: z.literal("text").optional(),
            // WordArt on these characters (`a:rPr`'s fill, `a:ln`, effects).
            textPaint: presentationTextPaintSchema.optional(),
            // Character Spacing on these characters (`a:rPr@spc`), points.
            letterSpacing: z.number().min(-200).max(200).optional(),
          })
          .passthrough()
          .optional(),
      })
      .passthrough(),
  )
  .max(MAX_RUNS_PER_BLOCK);
export const presentationRunsSchema = presentationRunsSchemaWithin(MAX_RUN_TEXT_LENGTH);
export const presentationShapeLabelRunsSchema =
  presentationRunsSchemaWithin(MAX_SHAPE_TEXT_LENGTH);

// Optional bulleted / numbered list paragraphs for a text block (PowerPoint /
// WPS parity). ADDITIVE — a text block written before this feature carries no
// `paragraphs` at all and still validates (the field is `.optional()`); when
// absent the block renders exactly as the flat `text` string. When present each
// paragraph is bounded here (text length / indent level / bullet enum) AND
// re-normalized at the editor + render edge by `normalizeParagraphs`
// (lib/presentation-paragraphs.ts), so a forged/hand-edited deck can never
// smuggle an unbounded list or an out-of-range indent level past load. The
// bounds mirror the runtime caps (MAX_PARAGRAPHS_PER_BLOCK /
// MAX_PARAGRAPH_TEXT_LENGTH / MIN_PARAGRAPH_LEVEL..MAX_PARAGRAPH_LEVEL). A
// paragraph may ALSO carry optional per-run `runs` (see
// `presentationRunsSchema`) for character-level formatting WITHIN the line.
// A text box's row is bounded by MAX_PARAGRAPH_TEXT_LENGTH, a shape label's by
// MAX_SHAPE_TEXT_LENGTH (PPT-B392).
export const presentationParagraphsSchemaWithin = (maxTextLength: number) => z
  .array(
    z
      .object({
        text: z.string().max(maxTextLength),
        level: z
          .number()
          .int()
          .min(MIN_PARAGRAPH_LEVEL)
          .max(MAX_PARAGRAPH_LEVEL),
        bullet: z.enum(PARAGRAPH_BULLETS),
        align: z.enum(PARAGRAPH_ALIGNS).optional(),
        lineSpacing: z
          .number()
          .min(MIN_STORED_LINE_SPACING)
          .max(MAX_LINE_SPACING)
          .optional(),
        // PowerPoint's Line Spacing > Exactly, in points (`a:lnSpc/a:spcPts`,
        // 0..1584): the row's lines are that tall whatever their type size.
        // ADDITIVE; a row states it OR a `lineSpacing` multiple, never both.
        lineSpacingPoints: z
          .number()
          .min(0)
          .max(MAX_EXACT_LINE_SPACING_PT)
          .optional(),
        spaceBefore: z
          .number()
          .min(MIN_PARAGRAPH_SPACING_PT)
          .max(MAX_PARAGRAPH_SPACING_PT)
          .optional(),
        spaceAfter: z
          .number()
          .min(MIN_PARAGRAPH_SPACING_PT)
          .max(MAX_PARAGRAPH_SPACING_PT)
          .optional(),
        spacingFallbackBlocked: z
          .array(z.enum(["lineSpacing", "spaceBefore", "spaceAfter"]))
          .max(3)
          .optional(),
        // The row's own indentation, in EMU: `a:pPr@marL` ("Before text",
        // ST_TextMargin 0..51206400) and `a:pPr@indent` (PowerPoint's
        // "Special" first line, NEGATIVE for a hanging indent, ST_TextIndent
        // over the same span either way) — ECMA-376 §20.1.10.65 / §20.1.10.62.
        // Both ADDITIVE and optional: a row stored before the import read them
        // carries neither and still validates, and an absent one keeps the
        // reading it always had (the outline `level`'s inherited margin).
        marL: z.number().int().min(0).max(PARAGRAPH_INDENT_EMU_LIMIT).optional(),
        indent: z
          .number()
          .int()
          .min(-PARAGRAPH_INDENT_EMU_LIMIT)
          .max(PARAGRAPH_INDENT_EMU_LIMIT)
          .optional(),
        // PowerPoint's Right-to-Left paragraph (`a:pPr@rtl`), in Docs' word for
        // it (`documentLayout.TEXT_DIRECTION_VALUES`); left-to-right is stated
        // by omission.
        dir: z.literal("rtl").optional(),
        // This row's own gallery marker, over the block-level `bulletChar` /
        // `numberFormat` below (PPT-B084): PowerPoint states the marker per
        // paragraph. Same allowed sets as the block fields.
        bulletChar: z.enum(BULLET_CHARS).optional(),
        numberFormat: z.enum(NUMBER_FORMATS).optional(),
        // The row's own Bullets and Numbering formatting beyond the gallery
        // (PPT-A039/A040), in the `nativeBullet` receipt's keys and units.
        bulletFormat: z
          .object({
            char: z.string().min(1).max(4).optional(),
            typeface: z.string().min(1).max(128).optional(),
            autoNum: z.string().regex(/^[A-Za-z0-9]{1,32}$/).optional(),
            startAt: z.number().int().min(1).max(32_767).optional(),
            color: z.string().regex(/^#[0-9a-fA-F]{6}$/).optional(),
            sizePct: z.number().int().min(25_000).max(400_000).optional(),
            sizePts: z.number().int().min(100).max(400_000).optional(),
          })
          .strict()
          .optional(),
        runs: presentationRunsSchemaWithin(maxTextLength).optional(),
      })
      .passthrough(),
  )
  .max(MAX_PARAGRAPHS_PER_BLOCK);
// A colour field's theme reference (`lib/presentation-theme-color-binding`):
// the `a:schemeClr` slot and luminance transform a theme swatch stands for,
// with the colour it drew as its authentication. Additive: a block without
// it validates byte-for-byte.
export const presentationThemeColorRefSchema = z
  .object({
    slot: z.enum([
      "bg1", "tx1", "bg2", "tx2",
      "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    ]),
    lumMod: z.number().int().min(0).max(100000).optional(),
    lumOff: z.number().int().min(0).max(100000).optional(),
    rgb: z.string().regex(/^#[0-9a-fA-F]{6}$/),
  })
  .strict();
export const presentationThemeColorsSchema = z
  .object({
    fill: presentationThemeColorRefSchema.optional(),
    stroke: presentationThemeColorRefSchema.optional(),
    outlineColor: presentationThemeColorRefSchema.optional(),
  })
  .strict();
// A reference with every transform in order (a shape's `p:style` line

// client before it refuses any key a swatch reference does not state
// (`presentation-theme-color-binding`'s THEMECOLORS_TRANSFORMED key).
const presentationThemeColorStepSchema = z.union([
  z.tuple([z.enum(["shade", "tint"]), z.number().int().min(0).max(100_000)]),
  z.tuple([z.enum(["lumMod", "satMod", "hueMod"]), z.number().int().min(0).max(10_000_000)]),
  z.tuple([z.enum(["lumOff", "satOff"]), z.number().int().min(-100_000).max(100_000)]),
  z.tuple([z.literal("hueOff"), z.number().int().min(-21_600_000).max(21_600_000)]),
]);
export const presentationThemeColorTransformedRefSchema = z
  .object({
    slot: presentationThemeColorRefSchema.shape.slot,
    transforms: z.array(presentationThemeColorStepSchema).min(1).max(8),
    rgb: z.string().regex(/^#[0-9a-fA-F]{6}$/),
  })
  .strict();
export const presentationThemeColorsTransformedSchema = z
  .object({
    fill: presentationThemeColorTransformedRefSchema.optional(),
    stroke: presentationThemeColorTransformedRefSchema.optional(),
    outlineColor: presentationThemeColorTransformedRefSchema.optional(),
  })
  .strict();

/** WordArt Transform on a text box's or a shape's words: the ONE warp value
 * of `@/lib/office-text-warp` (TEXT-WARP.md). */
export const presentationTextWarpSchema = z
  .unknown()
  .refine((value) => normalizeOfficeTextWarp(value) !== null, "Unsupported text warp");

// Optional six-digit hex colour for shape fill / stroke. Mirrors the slide
// background validation so a hand-edited or stale value can't inject arbitrary
// CSS.
export const presentationHexColorSchema = z.string().regex(/^#[0-9a-fA-F]{6}$/);

// Shape paint also has one bounded non-colour value: `none`, the persisted
// counterpart of DrawingML `a:noFill` for an explicit no-fill / no-outline
// choice. Keep gradient stops hex-only; only the primary fill/stroke paint may
// be absent this way.
export const presentationShapePaintSchema = z.union([
  presentationHexColorSchema,
  z.literal("none"),
]);
export const PRESENTATION_CAMEO_SHAPE_KINDS = new Set<string>(
  SHAPE_KINDS.filter(
    (kind) =>
      ![
        "line",
        "arrow",
        "double-arrow",
        "elbow-connector",
        "elbow-arrow-connector",
        "elbow-double-arrow-connector",
        "curved-connector",
        "curved-arrow-connector",
        "curved-double-arrow-connector",
        "ink",
        "custom-geometry",
        "flowchart-summing-junction",
      ].includes(kind) && !kind.startsWith("action-button-"),
  ),
);

export function presentationShapePayloadHasVisibleText(
  block: Record<string, unknown>,
): boolean {
  if (normalizeShapeText(block.text) !== undefined) return true;
  const hasVisibleRun = (value: unknown): boolean =>
    Array.isArray(value) &&
    value.some((run) => {
      if (!run || typeof run !== "object" || Array.isArray(run)) return false;
      return normalizeShapeText((run as Record<string, unknown>).text) !== undefined;
    });
  if (hasVisibleRun(block.runs)) return true;
  return (
    Array.isArray(block.paragraphs) &&
    block.paragraphs.some((paragraph) => {
      if (!paragraph || typeof paragraph !== "object" || Array.isArray(paragraph)) {
        return false;
      }
      const record = paragraph as Record<string, unknown>;
      return (
        normalizeShapeText(record.text) !== undefined ||
        hasVisibleRun(record.runs)
      );
    })
  );
}

// Additive shape block (rectangle / ellipse / line / arrow). Optional `shape`
// defaults to a rectangle at render time; `fill`/`stroke` are optional hex
// colours or the exact `none` sentinel above. Existing decks contain no `shape`
// blocks, so this union member is purely additive and never changes validation
// of pre-existing content.
//
// A legacy Record Studio ink stroke is read the way the editor reads it
// (`normalizePresentationBlock`) before it is validated: that studio stored
// every sample of a gesture, past the ink envelope `points` is held to, and its
// width in EMU, past the thickest outline. Each such stroke used to make the
// deck-wide parse that AI editor commands, Copilot selection and block paste
// rely on fail until the deck was next saved from the editor. Only values no
// current writer produces are rewritten (more points than the envelope holds,
// an ink wider than `MAX_STROKE_WIDTH`), with the shared thinning the server's
// slide validator applies too; the bounds below are unchanged for everything
// else, an ink within the envelope's point count and any other shape included.
export function presentationShapeBlockAsRead(value: unknown): unknown {
  if (!isPresentationShapeRecord(value) || value.shape !== "ink") return value;
  const points = presentationInkPointsWithinEnvelope(value.points);
  const strokeWidth = presentationInkStrokeWidthPx(value.strokeWidth);
  if (points === value.points && strokeWidth === value.strokeWidth) return value;
  return {
    ...value,
    ...(points !== value.points ? { points } : {}),
    ...(strokeWidth !== value.strokeWidth
      ? { strokeWidth: normalizeStrokeWidth(strokeWidth) }
      : {}),
  };
}

export function isPresentationShapeRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export const presentationShapeBlockSchema = z.preprocess(presentationShapeBlockAsRead, z
  .object({
    fillGradient:officeShapeGradientSchema.optional(),
    shapeShadow:officeShapeShadowSchema.optional(),
    
    // (`office_shape_paint.read_shape_3d` reads them all, as PowerPoint draws
    // them): held to the display-only guards, as a worksheet's native solid
    // is. The SmartArt Styles' narrower set refused the whole deck to every
    
    shape3d:z.custom<SmartArtShape3d>(isNativeShape3d).optional(),
    scene3d:z.custom<SmartArtScene3d>(isNativeScene3d).optional(),
    smartArtStyle:smartArtStyleIdSchema.optional(),
    // SmartArt Right to Left, on every member of a mirrored diagram.
    smartArtDirection:z.literal("rtl").optional(),
    // SmartArt Change Colors variation (Office's id), on every member.
    smartArtColors:smartArtColorsIdSchema.optional(),
    // The Office alias the diagram is drawn as (hProcess9), on every member.
    smartArtLayoutUrn:smartArtLayoutUrnSchema.optional(),
    id: z.string().min(1),
    type: z.literal("shape"),
    shape: z.enum(SHAPE_KINDS).optional(),
    // PowerPoint Cameo is a live-camera PLACEHOLDER, not persisted video.
    // The exact `true` marker keeps it inside the existing shape geometry /
    // arrange authority while the camera stream itself remains session-only.
    // Missing stays byte-identical for every ordinary shape.
    cameo: z.literal(true).optional(),
    alt: z.string().max(1_000).optional(),
    decorative: z.boolean().optional(),
    // The label's WordArt Transform (`a:prstTxWarp`).
    textWarp: presentationTextWarpSchema.optional(),
    // Format Shape > 3-D Rotation > Keep text flat (`a:flatTx`), only when checked.
    keepTextFlat: z.literal(true).optional(),
    // Optional CENTERED text label inside the shape (PowerPoint / WPS "add text
    // to shape" parity). Additive — pre-existing shape blocks carry no `text`
    // and still validate byte-for-byte. Bounded here; the render + serialize
    // edges re-sanitize via `normalizeShapeText` (lib/presentation-text-format),
    // which strips control bytes + drops an empty value so a label-less shape
    // stays byte-clean. May ALSO carry per-run `runs` for inline formatting
    // (same `presentationRunsSchema` as text blocks).
    // The label as a whole has the text box's bound, as the server's has none
    // of its own: a label of 2,159 characters imported from a real deck left
    // every button of that deck refused (AI70H h47 C13).
    text: z.string().max(MAX_TEXT_BLOCK_TEXT_LENGTH).optional(),
    runs: presentationShapeLabelRunsSchema.optional(),
    // A label of SEVERAL paragraphs, each with its own alignment, spacing,
    // bullet and runs — the same row model a text block carries, because it is
    // the same DrawingML `a:p`. Additive: a one-paragraph label omits it and
    // validates byte-for-byte. The importer emits it only above one row and the
    // exporter writes every row; before that the import gate refused a
    // multi-paragraph shape label outright (SHAPE-IMPORT-AUDIT.md §15.5).
    paragraphs: presentationParagraphsSchemaWithin(MAX_SHAPE_TEXT_LENGTH).optional(),
    bulletChar: z.enum(BULLET_CHARS).optional(),
    numberFormat: z.enum(NUMBER_FORMATS).optional(),
    fill: presentationShapePaintSchema.optional(),
    // The solid fill's Transparency (Format Shape > Fill > Transparency), a
    // whole percentage; PPTX `a:alpha` on the fill colour (PPT-A115).
    fillTransparency: z.number().int().min(1).max(100).optional(),
    // Second gradient stop; when present the fill renders as fill→fillTo.
    fillTo: presentationHexColorSchema.optional(),
    gradientAngle: z.number().min(0).max(360).optional(),
    // Optional PICTURE fill (PowerPoint "Picture or texture fill") — an image
    // url + sizing. Additive + optional: a colour / gradient shape omits it and
    // validates byte-for-byte. The url reuses the SAME validated inline-image
    // predicate as image blocks / slide backgrounds, so a stale / hand-edited
    // deck can never persist a dangerous `url()`; the render edge re-checks via
    // `normalizeShapeFillImage`, so this is a structural guard only.
    fillImage: presentationShapeFillImageSchema.optional(),
    fillPattern: presentationBgPatternSchema.optional(),
    stroke: presentationShapePaintSchema.optional(),
    // The solid line's Transparency (Format Shape > Line > Transparency), a
    // whole percentage; PPTX `a:alpha` on the line colour.
    outlineTransparency: z.number().int().min(1).max(100).optional(),
    themeColors: presentationThemeColorsSchema.optional(),
    themeColorsTransformed: presentationThemeColorsTransformedSchema.optional(),
    strokeTo: presentationHexColorSchema.optional(),
    // The OUTLINE's whole gradient -- the same model, reader and writer as
    // `fillGradient`. `strokeTo`/`strokeGradientAngle` stay beside it as the
    // two-colour projection an older client draws (audit SS32.4 -> SS36).
    strokeGradient: officeShapeGradientSchema.optional(),
    strokeGradientAngle: z.number().min(0).max(360).optional(),
    // Outline weight (px) and dash style; effects: a soft drop shadow.
    strokeWidth: z.number().min(0).max(MAX_STROKE_WIDTH).optional(),
    strokeStyle: z.enum(STROKE_STYLES).optional(),
    // `a:ln@algn`. Two words, no schema default: an absent field is a line
    // that states no alignment, which is every shape authored here and every
    // deck stored before audit 64. Declared so this schema carries the field
    // the projection reads and both writers write, instead of dropping it.
    strokeAlign: z.enum(["center", "inset"]).optional(),
    // Office's exact dash, compound type, cap and join (the fields Docs keeps).
    // The server writes every one of the Dash menu's eight it reads ("dash"
    // and "solid" included); the family alone draws those two.
    outlineDash: z
      .enum(["solid", "dash", ...(Object.keys(OFFICE_STROKE_DASH_FAMILY) as OfficeStrokeDash[])] as [string, ...string[]])
      .optional(),
    outlineCompound: z.enum(OFFICE_STROKE_COMPOUNDS).optional(),
    outlineCap: z.enum(OFFICE_STROKE_CAPS).optional(),
    outlineJoin: z.enum(OFFICE_STROKE_JOINS).optional(),
    outlineMiterLimit: z.number().int().min(0).max(OFFICE_STROKE_MAX_MITER_LIMIT).optional(),
    // Shape Outline > Sketched (PowerPoint's `ask:lineSketchStyleProps`).
    outlineSketch: z.enum(["curved", "freehand", "scribble"]).optional(),
    shadow: z.boolean().optional(),
    softEdge: z.boolean().optional(),
    // Shape Format > Shape Effects / Format Shape > Effects: Office's
    // combinable effects, the model pictures and Docs shapes keep
    // (`lib/office-picture-effects`, the server's `office_picture_effects`).
    shapeEffects: z
      .record(z.string(), z.unknown())
      .refine((value) => normalizeOfficePictureEffects(value) !== null, "Unsupported shape effects")
      .optional(),
    cropRect: presentationCropRectSchema.optional(),
    // Object transparency as a percentage (0–100; 100/absent = fully opaque).
    opacity: z.number().min(0).max(100).optional(),
    // Optional COMPOSITING mode (PowerPoint "Transparency / blend" parity → CSS
    // `mix-blend-mode`). Whitelisted + additive — a shape without it omits the
    // field and validates byte-for-byte. The render edge re-clamps via
    // `resolveShapeStyle` so a stale value can't inject arbitrary CSS.
    blendMode: z
      .enum(["normal", "multiply", "screen", "overlay", "darken", "lighten"])
      .optional(),
    // Anchored-connector fields (PowerPoint/WPS "connector" parity), all additive
    // + optional. On a `line`/`arrow` shape, `startShapeId`/`endShapeId` glue the
    // connector to two sibling shapes so it follows them on move/resize; the
    // direction that implies lives in the block's own `flipH`/`flipV` (set by
    // recomputeConnectors), because that is where DrawingML keeps it. A shape with
    // none of these validates exactly as before — pre-existing decks are untouched.
    // Ids reuse the groupId id-shape (non-empty, capped); a stale/hand-edited id
    // that no longer resolves degrades to an unanchored connector at render time
    // (no recompute; it keeps drawing its preset route in its own box).
    startShapeId: presentationGroupIdSchema.optional(),
    endShapeId: presentationGroupIdSchema.optional(),
    // LEGACY, READ-ONLY. Decks saved before the flips named the diagonal here
    // instead ("trbl" = the horizontal mirror of the preset's `(l,t)→(r,b)`
    // route, i.e. `flipH`). It is deliberately NOT normalized away: rewriting a
    // stored deck to drop it would change bytes no one asked to change, and an
    // old page must keep exporting the connector it always exported. The render
    // edge and the PPTX exporter both fold it into `flipH`; recomputeConnectors
    // deletes it the moment it recomputes that connector's flips.
    lineSlope: z.enum(["tlbr", "trbl"]).optional(),
    // Exact native connector identity. Explicit `straight` is a read-only
    // migration spelling and normalizes to absent so new JSON remains minimal.
    route: z
      .enum(PRESENTATION_CONNECTOR_ROUTES)
      .optional()
      .transform((value) => (value === "straight" ? undefined : value)),
    connectorPreset: z.enum(PRESENTATION_CONNECTOR_PRESETS).optional(),
    headEnd: z.enum(PRESENTATION_CONNECTOR_ENDS).optional(),
    tailEnd: z.enum(PRESENTATION_CONNECTOR_ENDS).optional(),
    // `a:ST_LineEndWidth` / `a:ST_LineEndLength`: the arrowhead's size relative
    // to the line weight ("Begin/End Arrow size" in PowerPoint). Optional with
    // NO schema default, so an absent one stays absent and the writer never
    // invents `med` on an untouched arrowhead (SHAPE-IMPORT-AUDIT.md §7.4).
    headEndWidth: z.enum(PRESENTATION_CONNECTOR_END_SIZES).optional(),
    headEndLength: z.enum(PRESENTATION_CONNECTOR_END_SIZES).optional(),
    tailEndWidth: z.enum(PRESENTATION_CONNECTOR_END_SIZES).optional(),
    tailEndLength: z.enum(PRESENTATION_CONNECTOR_END_SIZES).optional(),
    // ZOOM-tile markers (PowerPoint/WPS "Insert ▸ Zoom" parity), all additive +
    // optional. `zoomKind` selects the navigation intent and `zoomTargetId` names
    // the target slide / section (reuses the groupId id-shape: non-empty, capped).
    // A shape with neither validates exactly as before — pre-existing decks are
    // untouched. The render edge re-reads them via `readZoomTileMarker` and the
    // `href` (validated separately) drives the actual present-mode jump.
    zoomKind: z.enum(["slide", "section", "summary"]).optional(),
    zoomTargetId: presentationGroupIdSchema.optional(),
    // CUSTOM polygon outline (PowerPoint "Merge Shapes → Intersect"): the overlap
    // of two merged shapes as an "x,y x,y …" string in the shape's own 0–100
    // space; on ink, the stroke's polyline. Size-capped here at the ink
    // envelope (1024 points × 16 chars, the server's cap too: an imported path
    // keeps a point per command); the render edge does the real geometric
    // validation via `normalizeCustomPolygonPoints` (256 vertices for a
    // polygon) and falls back to the shape kind when malformed.
    // Optional/additive — a plain shape without it validates byte-for-byte as
    // before.
    points: z.string().max(PRESENTATION_SHAPE_POINTS_MAX_LENGTH).optional(),
    // Draw-tab ink: the tool that drew the stroke (a pencil wears its
    // texture, a highlighter its wash) and, for a pen, its effect: one of
    // Office's eight, an MRAI preset or a custom one (lib/office-ink-effects;
    // `stroke` holds the effect's solid stand-in, which is what the file
    // carries).
    inkTool: z.enum(["pen", "highlighter", "pencil"]).optional(),
    inkEffect: z.string().max(INK_EFFECT_MAX_CHARS).refine(isInkEffect, "is not an ink effect").optional(),
    // Draw-tab ink a point eraser cut: the ends that stop flat.
    inkCutEnds: z.enum(INK_CUT_ENDS).optional(),
    // Draw-tab ink from a pressure pen: one integer 0..100 per outline
    // sample (paint only; the editor checks the count against the outline).
    inkPressure: z.string().max(MAX_INK_PRESSURE_CHARS).optional(),
    customPath: z
      .custom<PresentationCustomPath>(
        // A stored page may name a shape's path, held once in
        // `deckReceipts` (`presentation-deck-receipts`); the loader writes it
        // out again, and it is checked in that form.
        (value) =>
          normalizePresentationCustomPath(value) !== null ||
          presentationDeckReceiptTokenKey(value) !== null,
        "Unsupported presentation custom path",
      )
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
    // Action Settings > Mouse Over (`a:hlinkMouseOver`, PPT-B012).
    hoverHref: presentationBlockHrefSchema.optional(),
    hoverAction: presentationRelativeClickActionSchema.optional(),
    hoverHighlight: z.literal(true).optional(),
    // Grayscale > Change Selected Object (`p:spPr@bwMode`, PPT-A109); Automatic is absent.
    blackWhiteMode: z.enum(PRESENTATION_STORED_BLACK_WHITE_MODES).optional(),
    clickAction: presentationClickActionSchema.optional(),
    groupId: presentationGroupIdSchema.optional(),
    // The groups nested inside `groupId` this block also belongs to,
    // outermost first (presentation-crop-group `blockGroupPath`).
    groupPath: z.array(presentationGroupIdSchema).min(1).max(7).optional(),
    // `a:grpFill` worn as the part of the group's own gradient, picture or
    // pattern under the block (`presentation-group-fill`).
    fillFromGroup: z.literal(true).optional(),
    groupFill: z.record(z.string(), z.unknown()).optional(),
    groupFillDepth: z.number().int().min(1).max(8).optional(),
  })
  .passthrough()
  .superRefine((block, ctx) => {
    if (block.fillPattern && (block.fillTo || block.fillImage)) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["fillPattern"],
        message: "pattern fill cannot share gradient or picture fill fields",
      });
    }
    if (block.strokeTo && (!block.stroke || block.stroke === "none")) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["strokeTo"],
        message: "gradient outline requires a visible first colour",
      });
    }
    if (block.strokeGradient && (!block.stroke || block.stroke === "none")) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["strokeGradient"],
        message: "gradient outline requires a visible first colour",
      });
    }
    if (block.strokeGradientAngle !== undefined && !block.strokeTo) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["strokeGradientAngle"],
        message: "gradient outline angle requires a second colour",
      });
    }
    const shape = normalizeShapeKind(block.shape);
    if (block.cameo === true && !PRESENTATION_CAMEO_SHAPE_KINDS.has(shape)) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["shape"],
        message: "Cameo requires a closed PowerPoint Web camera shape",
      });
    }
    if (
      block.cameo === true &&
      presentationShapePayloadHasVisibleText(
        block as unknown as Record<string, unknown>,
      )
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["text"],
        message: "Cameo cannot carry authored shape text",
      });
    }
    if (
      block.cameo === true &&
      [
        block.fill,
        block.fillTo,
        block.gradientAngle,
        block.fillImage,
        block.fillPattern,
        block.strokeTo,
        block.strokeGradient,
        block.strokeGradientAngle,
        block.opacity,
        block.blendMode,
        block.points,
        block.customPath,
      ].some((value) => value !== undefined)
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["cameo"],
        message:
          "Cameo cannot carry fill, gradient, transparency, or custom geometry",
      });
    }
    if (
      block.cameo === true &&
      block.shadow === true &&
      block.softEdge === true
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["softEdge"],
        message: "Cameo shadow and soft edge styles are mutually exclusive",
      });
    }
    if (block.decorative === true && block.alt?.trim()) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["alt"],
        message: "Decorative shapes cannot carry alternative text",
      });
    }
    const isActionButton = isPresentationActionButtonShapeKind(shape);
    const absentNativeAction = presentationHasAbsentNativeAction(block as unknown as Record<string, unknown>);
    const targetCount = Number(block.href !== undefined) + Number(block.clickAction !== undefined);
    if (
      (isActionButton && targetCount !== 1 && !absentNativeAction) ||
      // Any object's Action Settings may jump relative (PPT-B012); a media
      // action is an Action Button's own, and one click has one target.
      (!isActionButton &&
        block.clickAction !== undefined &&
        (block.clickAction.kind !== "relative-slide" || block.href !== undefined))
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["clickAction"],
        message: isActionButton
          ? "Action Buttons require exactly one safe click target"
          : "An object's click action is one relative slide jump",
      });
    }
    if (block.clickAction?.kind === "play-media") {
      const expectedKind =
        shape === "action-button-movie"
          ? "video"
          : shape === "action-button-sound"
            ? "audio"
            : null;
      if (expectedKind !== block.clickAction.mediaKind) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ["clickAction"],
          message: "Media actions require the matching Movie or Sound Action Button",
        });
      }
    }
    if (
      (shape === "action-button-movie" ||
        shape === "action-button-sound") &&
      block.clickAction?.kind !== "play-media" && !absentNativeAction
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["clickAction"],
        message: "Movie and Sound Action Buttons require a bound media action",
      });
    }
    const hasConnectorField = [
      block.route,
      block.connectorPreset,
      block.headEnd,
      block.tailEnd,
      block.headEndWidth,
      block.headEndLength,
      block.tailEndWidth,
      block.tailEndLength,
    ].some((value) => value !== undefined);
    const openShapeEnds =
      block.route === undefined &&
      block.connectorPreset === undefined &&
      !isPresentationConnectorShapeKind(shape) &&
      presentationShapeTakesLineEnds(block);
    if (openShapeEnds) {
      if (hasConnectorField && normalizePresentationOpenShapeLineEnds(block) === null) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: "Unsupported arrowhead on an open shape",
        });
      }
    } else if (
      (hasConnectorField || isPresentationConnectorShapeKind(shape)) &&
      normalizePresentationConnectorSpec(block) === null
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "Unsupported or contradictory presentation connector identity",
      });
    }
    const customPath = normalizePresentationCustomPath(block.customPath);
    const storedCustomPath = presentationDeckReceiptTokenKey(block.customPath) !== null;
    if (
      (shape === "ink" || shape === "custom-geometry") &&
      ((block.text?.length ?? 0) > 0 || (block.runs?.length ?? 0) > 0)
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "Ink and custom geometry do not support authored text",
      });
    }
    if (storedCustomPath) {
      // Named, not written out: checked once the loader writes it out.
    } else if (shape === "custom-geometry") {
      if (!customPath || !presentationCustomGeometryPathAdmitted(customPath)) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message:
            "Custom geometry requires a closed, Bezier-backed or multi-figure custom path",
        });
      }
    } else if (shape === "ink") {
      if (customPath && !presentationCustomPathIsOpenLineOnly(customPath)) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: "Ink custom paths must be open line-only geometry",
        });
      }
    } else if (block.customPath !== undefined) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "Custom path is not valid on a preset shape",
      });
    }
  }));
