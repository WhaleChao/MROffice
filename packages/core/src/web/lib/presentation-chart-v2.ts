/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  createChartBlockAttrsV2,
  officeInsertedChartLegendAndLabels,
  withOfficeInsertedChartFormat,
  defaultChartData,
  validateChartBlockAttrsV2,
  type ChartBlockAttrsV2,
  type ChartFamily,
  type ChartSubtype,
} from "../components/console/editor/extensions/chartBlockShared";
import {
  officeChartLeafForLegacyKind,
  officeChartLegacyKind,
  type OfficeChartLegacyKind,
} from "./office-charts/catalog";
import { officeChartPartKind } from "./office-charts/chart-type";
import type { PresentationChartBlock } from "../components/console/editor/native-document/documentCommands";
import { sameChartWorkbookDataShape } from "../components/console/editor/extensions/documentChartWorkbook";

export function readPresentationChartV2(
  value: unknown,
): ChartBlockAttrsV2 | null {
  // PowerPoint offers the same Insert > Chart catalog as Word, so every one of
  // the 61 catalog leaves is a legal owned V2 chart here. The shared validator
  // is the only gate: it already enforces the per-family data contracts,
  // including pie/funnel being single-series and the documented stock roles.
  return validateChartBlockAttrsV2(value).valid
    ? (value as ChartBlockAttrsV2)
    : null;
}

/** The label a chart block carries in `chartType`, straight from the catalog. */
export function presentationChartV2Kind(attrs: ChartBlockAttrsV2): string {
  return officeChartLegacyKind(attrs.chartType, attrs.chartSubtype);
}

/**
 * The attrs PowerPoint inserts for one catalog leaf.
 *
 * `title` is what the chart title reads. PowerPoint inserts a chart with its
 * title on, reading "Chart Title" -- the placeholder Word's inserted chart
 * shows and Docs inserts -- so the deck passes its localized placeholder; a
 * caller building a chart from its own data may leave it empty. `locale` is
 * the reader's interface language, which names the sample data
 * ("Category 1" / "Series 1" in English PowerPoint, "类别 1" / "系列 1" in
 * Chinese).
 */
export function createPresentationChartV2Leaf(
  family: ChartFamily,
  subtype: ChartSubtype,
  title = "",
  locale?: string | null,
): ChartBlockAttrsV2 {
  // PowerPoint's inserted chart shows its legend under the plot area, or
  // where its kind's own style puts it (a radar's, a waterfall's and a
  // treemap's at the top; a funnel's nowhere), and spaces it as Office does.
  return withOfficeInsertedChartFormat(
    createChartBlockAttrsV2(family, subtype, {
      title,
      format: { legend: officeInsertedChartLegendAndLabels(family, subtype).legend ?? "bottom" },
    }, { locale }),
  );
}

/** Same insert, addressed by the legacy PPT kind name a leaf may still carry. */
export function createPresentationChartV2(
  kind: OfficeChartLegacyKind,
): ChartBlockAttrsV2 {
  const leaf = officeChartLeafForLegacyKind(kind);
  if (!leaf) throw new Error(`Unknown presentation chart kind: ${kind}`);
  return createPresentationChartV2Leaf(leaf.family, leaf.subtype);
}

function record(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

/**
 * The owned chart authorities the deck's writers rewrite, with the part
 * fingerprints each states: a classic chart's (`chart_v2.py`), and an Office
 * 2016 chart's, whose cx part, classic `mc:Fallback` twin, workbooks, style,
 * colour and contract parts are rebuilt together (`chart_ex.py`
 * `owned_chart_ex_replacements`).
 */
const OWNED_CHART_AUTHORITY_FINGERPRINTS = {
  "pptx-owned-chart-v2": [
    "chartFingerprint",
    "chartRelationshipsFingerprint",
    "workbookFingerprint",
  ],
  // A classic chart another application wrote, read into the one chart
  // editor at import (`chart_v2.py` `project_imported_ppt_chart_v2`); its
  // edit is written into its own part in place. Its workbook, when the file
  // embeds one, is fingerprinted too (`workbookFingerprint`), and a chart
  // whose data is linked from outside the file has none.
  "pptx-imported-chart-v2": [
    "chartFingerprint",
    "chartRelationshipsFingerprint",
    "viewChartFingerprint",
  ],
  // An Office 2016 (ChartEx) chart another application wrote, read into the
  // one chart editor at import (`chart_ex.py` `project_imported_ppt_chart_ex`):
  // its cx part is written in place by the ChartEx writer, which proves the
  // part's bytes alone; the workbook it embeds (always one, or the projection
  // refuses the chart) is named but not fingerprinted.
  "pptx-imported-chart-ex-v1": ["chartFingerprint"],
  "pptx-owned-chart-ex-v2": [
    "chartPartFingerprint",
    "chartRelationshipsFingerprint",
    "workbookPartFingerprint",
    "fallbackChartPartFingerprint",
    "fallbackChartRelationshipsFingerprint",
    "contractPartFingerprint",
  ],
} as const;

/**
 * The kind of part an owned chart is in the deck's file, or `null` for a
 * chart not yet saved (the append writes it as either kind).
 */
export function presentationChartV2FilePartKind(
  block: PresentationChartBlock,
): "classic" | "chartEx" | null {
  const authority = record(record(block.source)?.nativeChartV2Authority);
  return authority?.schemaVersion === "pptx-owned-chart-ex-v2" ||
    authority?.schemaVersion === IMPORTED_CHART_EX_AUTHORITY
    ? "chartEx"
    : authority?.schemaVersion === "pptx-owned-chart-v2" ||
        authority?.schemaVersion === IMPORTED_CHART_AUTHORITY
      ? "classic"
      : null;
}

const IMPORTED_CHART_AUTHORITY = "pptx-imported-chart-v2";
const IMPORTED_CHART_EX_AUTHORITY = "pptx-imported-chart-ex-v1";

function importedChartAuthority(
  block: PresentationChartBlock,
): Record<string, unknown> | null {
  const authority = record(record(block.source)?.nativeChartV2Authority);
  return authority?.schemaVersion === IMPORTED_CHART_AUTHORITY ||
    authority?.schemaVersion === IMPORTED_CHART_EX_AUTHORITY
    ? authority
    : null;
}

/** A chart read from an imported part (its receipt, or a page's native view). */
export function presentationChartIsImported(block: PresentationChartBlock): boolean {
  return (
    importedChartAuthority(block) !== null ||
    record(record(block.source)?.nativeChartView) !== null
  );
}

/**
 * Where an imported chart's data lives, as its receipt classes it (`chart_v2.py`
 * `dataSource`: embedded, linked, ole, none, readonly). An editable ChartEx
 * receipt says `embedded` (its projection admits only a chart with one
 * embedded workbook); one stored before it said so has no class and means the
 * same.
 */
function importedChartDataSource(authority: Record<string, unknown>): string {
  return authority.schemaVersion === IMPORTED_CHART_EX_AUTHORITY &&
    typeof authority.dataSource !== "string"
    ? "embedded"
    : String(authority.dataSource);
}

/**
 * A chart the one model reads only in part (`chart_v2.py`
 * `project_preserved_ppt_chart_v2`, `editable_page.py` `_pptx_preserved_chart`):
 * its receipt says `editable: false` and names why (`preserveReasons`). It is
 * drawn from its attrs (or its native view), selected on the one chart
 * editor, and takes no edit -- the save declines every edit of it as
 * `preserved:<first reason>`. Its frame still moves, resizes and deletes.
 */
export function presentationChartV2PreserveOnly(block: PresentationChartBlock): boolean {
  return block.chartV2 !== undefined && importedChartAuthority(block)?.editable === false;
}

/** The reasons a preserve-only chart's receipt names, first the proof step. */
export function presentationChartV2PreserveReasons(
  block: PresentationChartBlock,
): readonly string[] {
  if (!presentationChartV2PreserveOnly(block)) return [];
  const reasons = importedChartAuthority(block)?.preserveReasons;
  return Array.isArray(reasons)
    ? reasons.filter((reason): reason is string => typeof reason === "string")
    : [];
}

/**
 * The reason codes a receipt (and a declined save, `preserved:<code>`) may
 * name that the deck has words for: the step of the writable proof that
 * refused the chart, then each construct the one model's reading left out
 * (`CHART_READ_GAPS`), then `model:<key>` for a format key the model refuses
 * for the chart's family. Each is `presentation.chartPreserved.reason.<name>`.
 */
const PRESERVE_REASON_MESSAGES: Readonly<Record<string, string>> = {
  "part-graph": "partGraph",
  "view-unread": "viewUnread",
  "view-differs": "viewDiffers",
  "chart-ex-part": "chartExPart",
  "view-only": "viewOnly",
  "no-model": "noModel",
  "area-paint": "areaPaint",
  "title-paint": "titlePaint",
  "legend-paint": "legendPaint",
  "series-paint": "seriesPaint",
  "axis-paint": "axisPaint",
  "data-labels": "dataLabels",
  trendlines: "trendlines",
  "error-bars": "errorBars",
  "up-down-bars": "upDownBars",
  "chart-lines": "chartLines",
  gridlines: "gridlines",
  "data-table": "dataTable",
  "axis-visibility": "axisVisibility",
  "axis-titles": "axisTitles",
  "series-name": "seriesName",
  categories: "categories",
  "category-levels": "categoryLevels",
  "points-beyond-model": "pointsBeyondModel",
  "series-beyond-model": "seriesBeyondModel",
  "text-beyond-model": "textBeyondModel",
  "point-paint": "pointPaint",
  "pie-series": "pieSeries",
  format: "format",
};

/**
 * The catalogue key naming one preserve reason, or null for a code this build
 * has no words for (the caller says the generic notice instead).
 */
export function presentationChartPreserveReasonKey(code: string): string | null {
  const name = code.startsWith("model:") ? "model" : PRESERVE_REASON_MESSAGES[code];
  return name ? `presentation.chartPreserved.reason.${name}` : null;
}

/** A translator over the deck's catalogue (`console-notebooks`). */
export type PresentationChartTranslate = (
  key: string,
  values?: Record<string, string>,
) => string;

/** The notice a preserve-only chart's reason code reads as. */
export function presentationChartPreserveNotice(
  code: string | undefined,
  t: PresentationChartTranslate,
): string {
  const reason = code === undefined ? null : presentationChartPreserveReasonKey(code);
  return reason
    ? t("presentation.chartPreserved.notice", { reason: t(reason) })
    : t("presentation.chartPreserved.generic");
}

/**
 * Why no chart command may write this chart, or null when the one editor
 * may: a preserve-only receipt's first reason, and for any other chart the
 * editor may not claim (a receipt still to come, one naming other parts) the
 * generic notice. The frame's own commands are not the chart's.
 */
export function presentationChartV2LockNotice(
  block: PresentationChartBlock,
  t: PresentationChartTranslate,
): string | null {
  if (presentationChartV2PreserveOnly(block))
    return presentationChartPreserveNotice(presentationChartV2PreserveReasons(block)[0], t);
  return presentationChartV2CanEdit(block) ? null : t("presentation.chartPreserved.generic");
}

/**
 * An imported chart stored before the server's V2 reading: the page holds
 * its native view (`source.nativeChartView`) and no receipt, and the deck
 * carries the view's composed attrs as `chartV2` for the canvas and the one
 * ribbon (`upgradePresentationChartBlock`). The receipt that lets the writer
 * edit the part is the server's to grant, so such a chart is read-only here
 * and the save sends its stored reading back (`presentationChartBlockForSave`).
 */
export function presentationChartV2AwaitsReceipt(block: PresentationChartBlock): boolean {
  const source = record(block.source);
  return (
    block.chartV2 !== undefined &&
    source !== null &&
    (typeof source.slideId === "string" ||
      typeof source.slidePart === "string" ||
      typeof source.part === "string") &&
    record(source.nativeChartV2Authority) === null
  );
}

/** A chart another application wrote, on the one chart editor since import. */
export function presentationChartIsImportedV2(block: PresentationChartBlock): boolean {
  return (
    block.chartV2 !== undefined &&
    (importedChartAuthority(block) !== null || presentationChartV2AwaitsReceipt(block))
  );
}












export function presentationChartV2DataWritable(block: PresentationChartBlock): boolean {
  const authority = importedChartAuthority(block);
  if (authority === null) return true;
  if (authority.editable === false) return false;
  const source = importedChartDataSource(authority);
  return source === "embedded" || source === "none" || source === "linked" || source === "ole";
}

/**
 * Whether a data edit of this chart breaks its link to a workbook outside
 * the file (or replaces a binary workbook): its data is then the values the
 * file stored, no longer the source workbook's -- what the data window says
 * while it is open.
 */
export function presentationChartV2BreaksLink(block: PresentationChartBlock): boolean {
  const authority = importedChartAuthority(block);
  if (authority === null || authority.editable === false) return false;
  const source = importedChartDataSource(authority);
  return source === "linked" || source === "ole";
}

/**
 * Whether the chart's type can be changed here. Change Chart Type keeps a
 * chart's data (Microsoft Support, "Change the chart type of an existing
 * chart"): an imported classic chart is rewritten around its embedded
 * workbook, or -- with no workbook this editor writes (values cached in the
 * part, a linked, binary or unwritable workbook) -- around the series data
 * its part states, references, caches and data link kept, as PowerPoint
 * rewrites such a chart from its cache (`office_chart_edit`, `part_data`).
 * An imported Office 2016 part is written in place only, never rebuilt, and
 * a preserve-only chart takes no edit. The save declines it otherwise
 * (`editable_page_patchback` `_pptx_chart_patch`: `type_writable`).
 */
export function presentationChartV2TypeWritable(block: PresentationChartBlock): boolean {
  const authority = importedChartAuthority(block);
  // An imported Office 2016 chart too: within its family its part is edited
  // in place (a treemap laid out as a sunburst), and another family -- a
  // classic type or another Office 2016 one -- is a new part compiled from
  // its attrs with its own workbook, as PowerPoint's Change Chart Type
  // rebuilds it (`pptx/chart_kind_conversion`).
  return authority === null || authority.editable !== false;
}

/**
 * Why the chart's type cannot be changed while the rest of it can, or null:
 * an Office 2016 part written in place.
 */
export function presentationChartV2TypeLock(
  block: PresentationChartBlock,
): "chartEx" | null {
  if (presentationChartV2TypeWritable(block)) return null;
  return importedChartAuthority(block)?.schemaVersion === IMPORTED_CHART_EX_AUTHORITY ? "chartEx" : null;
}

/**
 * Why this chart cannot become ``family`` while it can change type, or null.
 * An imported classic part is rewritten as a classic part (an Office 2016
 * type is another kind of part, written only for a chart this product
 * authored, `pptx/chart_kind_conversion.py`); one read around a look the
 * model has no words for (`nativeLook`) keeps its family, since another
 * family is a rebuilt part that would drop that look. The save declines
 * either by itself (`_pptx_chart_patch`).
 */
export function presentationChartTypeTargetLock(
  block: PresentationChartBlock,
  family: ChartFamily,
): "chartExTarget" | "nativeLook" | null {
  const authority = importedChartAuthority(block);
  if (authority === null) return null;
  if (authority.schemaVersion === IMPORTED_CHART_AUTHORITY && officeChartPartKind(family) === "chartEx")
    return "chartExTarget";
  if (authority.schemaVersion !== IMPORTED_CHART_AUTHORITY && authority.schemaVersion !== IMPORTED_CHART_EX_AUTHORITY)
    return null;
  const look = authority.nativeLook;
  if (Array.isArray(look) && look.length > 0 && family !== block.chartV2?.chartType)
    return "nativeLook";
  return null;
}

/**
 * Whether the chart's title text can be typed here: on every chart this
 * editor edits. A title that is a cell of the chart's workbook (`titleBound`,
 * `c:tx/c:strRef`) takes typed words as PowerPoint's does: typing over it
 * replaces the reference with the words and the cell keeps its text, however
 * the chart's data is held -- embedded, linked to a workbook outside the file,
 * or cached in the part (`office_chart_edit._replace_title_text`). It used to
 * be greyed on a linked chart, "edit it in the linked workbook".
 */
export function presentationChartV2TitleWritable(block: PresentationChartBlock): boolean {
  return block.chartV2 === undefined || presentationChartV2CanEdit(block);
}

/**
 * An imported chart still as its part reads: drawn from its native view
 * (the display profiles paint a 3-D pie, a bubble chart, a pie of pie as
 * PowerPoint laid them out). Once edited, the shared planner draws the
 * attrs, as it draws every chart this editor authored.
 */
export function presentationChartV2DrawsView(block: PresentationChartBlock): boolean {
  // A chart awaiting its receipt is unedited by construction: its attrs ARE
  // the view's composition, and the editor refuses every edit of them.
  if (presentationChartV2AwaitsReceipt(block)) return true;
  const authority = importedChartAuthority(block);
  if (authority === null || authority.viewAttrs === undefined) return false;
  return JSON.stringify(block.chartV2) === JSON.stringify(authority.viewAttrs);
}

/** A browser import may render V2 without claiming authority to rewrite it. */
export function presentationChartV2CanEdit(
  block: PresentationChartBlock,
): boolean {
  if (!readPresentationChartV2(block.chartV2)) return false;
  const source = record(block.source);
  if (!source || Object.keys(source).length === 0) return true;
  const authority = record(source.nativeChartV2Authority);
  // A preserve-only receipt keeps the writable receipts' keys (a linked
  // chart's names no workbook at all) and says it grants nothing.
  if (authority?.editable === false) return false;
  const fingerprints =
    authority &&
    OWNED_CHART_AUTHORITY_FINGERPRINTS[
      authority.schemaVersion as keyof typeof OWNED_CHART_AUTHORITY_FINGERPRINTS
    ];
  if (
    !authority ||
    !fingerprints ||
    authority.slideId !== source.slideId ||
    authority.shapeId !== source.shapeId ||
    authority.chartPart !== source.part ||
    authority.slidePart !== source.slidePart
  )
    return false;
  // An imported chart's workbook is in the receipt only when the file embeds
  // one; a chart with linked data has neither a workbook part nor its
  // relationship, and every other edit of it is still written in place.
  const embedded =
    authority.schemaVersion !== IMPORTED_CHART_AUTHORITY ||
    authority.dataSource === "embedded";
  // The ChartEx projection reads the slide's XML itself and names no slide
  // relationship (`chart_ex.py` `project_imported_ppt_chart_ex`); the writer
  // addresses the chart by slide, shape and part.
  const slideRelationship = authority.schemaVersion !== IMPORTED_CHART_EX_AUTHORITY;
  return (
    fingerprints.every(
      (key) =>
        typeof authority[key] === "string" &&
        /^[a-f0-9]{64}$/.test(authority[key] as string),
    ) &&
    (!embedded ||
      typeof authority.workbookFingerprint !== "string" ||
      /^[a-f0-9]{64}$/.test(authority.workbookFingerprint)) &&
    [
      ...(slideRelationship ? ["slideRelationshipId"] : []),
      ...(embedded ? ["workbookRelationshipId", "workbookPart"] : []),
    ].every(
      (key) =>
        typeof authority[key] === "string" &&
        (authority[key] as string).length > 0,
    )
  );
}

export function withPresentationChartV2(
  block: PresentationChartBlock,
  attrs: ChartBlockAttrsV2,
): PresentationChartBlock {
  return {
    ...block,
    chartV2: attrs,
    title: attrs.title,
    chartType: presentationChartV2Kind(attrs),
  };
}











export function presentationChartV2DataShapeWritable(block: PresentationChartBlock): boolean {
  const authority = importedChartAuthority(block);
  if (authority === null || importedChartDataSource(authority) !== "none") return true;
  return readPresentationChartV2(block.chartV2)?.data.kind === "category";
}

/** Whether ``data`` is an edit of the chart's data the save writes. */
export function presentationChartDataEditAllowed(
  block: PresentationChartBlock,
  data: ChartBlockAttrsV2["data"],
): boolean {
  if (!presentationChartV2DataWritable(block)) return JSON.stringify(data) === JSON.stringify(block.chartV2?.data);
  if (presentationChartV2DataShapeWritable(block) || block.chartV2 === undefined) return true;
  return sameChartWorkbookDataShape(block.chartV2.data, data);
}

/**
 * Why a chart cannot become this type because the type reads another kind
 * of data (X-Y values, stock prices, a hierarchy...) than the chart holds, or
 * null. Office's Change Chart Type keeps a chart's data; the one chart model
 * has no conversion between kinds (and an imported chart's workbook ranges
 * are of its own kind: the save declines a kind change as 'chartType').
 */
export function presentationChartTypeDataKindRefusal(
  attrs: ChartBlockAttrsV2,
  family: ChartFamily,
  subtype: ChartSubtype,
): "dataKind" | null {
  return defaultChartData(family, subtype).kind === attrs.data.kind ? null : "dataKind";
}
