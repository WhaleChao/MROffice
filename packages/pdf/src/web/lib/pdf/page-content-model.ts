/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The PDF editor's page-content model: how a stored PDF page
 * (`mrnote.pdf.v1` content_json) opens into the editor's collections, and how
 * those collections are written back as page content on save.
 *
 * Moved out of PdfPageWindow unchanged so the editor window and the headless
 * editor commands (lib/editor-commands/pdf) read and write a page through the
 * same functions. Pure: no React, no DOM, no pdf.js.
 */

import type { PdfImageOverlayItem } from "../../components/notebook/contents/pdf/PdfImageOverlay";
import type { PdfInkStroke } from "../../components/notebook/contents/pdf/PdfInkOverlay";
import type { PdfTextBoxEdit } from "../../components/notebook/contents/pdf/PdfTextBoxOverlay";
import {
  PDF_CONTENT_TYPE,
  pdfPageContentSchema,
} from "@mrnote/core/web/lib/native-page-schemas.ts";
import type { PdfAnnotationReply } from "@mrnote/core/web/lib/pdf/annotation-replies.ts";
import {
  type CropBoxEntry,
  sanitizeCropBox,
} from "./crop-box";
import {
  hasPdfDocumentDecorations,
  type PdfDocumentDecorations,
  sanitizePdfDocumentDecorations,
} from "./document-decorations";
import type { FlattenPageSize } from "@mrnote/core/web/lib/pdf/flatten.ts";
import {
  type NativePdfFormAuthoringOperation,
  readNativePdfFormAuthoringOperations,
} from "./form-field-authoring";
import type { NativePdfFormFieldEdit } from "./form-field-writeback";
import {
  type InkAnnotation,
  normalizeInkPenType,
} from "@mrnote/core/web/lib/pdf/ink.ts";
import {
  isManagedNativeMeasurementAnnotationText,
  type ManagedNativeMeasurementRecord,
  parseManagedNativeMeasurementShadow,
  planManagedNativeMeasurementAnnotations,
} from "./measure-annotations";
import {
  isNativePdfFreeTextRecord,
  type NativePdfFreeTextAnnotation,
  sanitizeNativePdfFreeTextAnnotationSource,
} from "./native-free-text-annotations";
import { sanitizeNativePdfInkAnnotationSource } from "./native-ink-annotations";
import {
  isManagedNativeInkAnnotationText,
  planManagedNativeInkAnnotations,
} from "./native-ink-roundtrip";
import type { NativePdfLinkRecord } from "./native-link-annotations";
import {
  isManagedNativeLinkAnnotationText,
  type ManagedNativeLinkRecord,
  parseImportedNativeLinkShadow,
  parseManagedNativeLinkShadow,
  planImportedNativeLinkAnnotations,
  planManagedNativeLinkAnnotations,
} from "./native-link-roundtrip";
import {
  partitionManagedNativeNotes,
  readNativeNoteReplyAuthority,
  serializeManagedNativeNotes,
} from "@mrnote/core/web/lib/pdf/native-note-roundtrip.ts";
import type { NativePdfShapeRecord } from "./native-shape-annotations";
import {
  isManagedNativeShapeAnnotationText,
  type ManagedNativeShapeRecord,
  parseImportedNativeShapeShadow,
  parseManagedNativeShapeShadow,
  planImportedNativeShapeAnnotations,
  planManagedNativeShapeAnnotations,
} from "./native-shape-roundtrip";
import {
  isManagedNativeSignatureAnnotationText,
  type ManagedNativeSignatureRecord,
  parseManagedNativeSignatureShadow,
  planManagedNativeSignatureAnnotations,
} from "./native-signature-roundtrip";
import type { NativePdfStampRecord } from "./native-stamp-annotations";
import {
  isManagedNativeStampAnnotationText,
  type ManagedNativeStampRecord,
  parseImportedNativeStampShadow,
  parseManagedNativeStampShadow,
  planImportedNativeStampAnnotations,
  planManagedNativeStampAnnotations,
} from "./native-stamp-roundtrip";
import {
  type NativePdfAnnotationSource,
  sanitizeNativePdfAnnotationSource,
  sanitizeNativePdfMarkupColor,
} from "@mrnote/core/web/lib/pdf/native-text-annotations.ts";
import {
  isManagedNativeTextBoxAnnotationText,
  type ManagedNativeTextBoxRecord,
  parseImportedNativeTextBoxShadow,
  parseManagedNativeTextBoxShadow,
  planImportedNativeTextBoxAnnotations,
  planManagedNativeTextBoxAnnotations,
  sanitizeManagedNativeTextBoxImage,
} from "./native-text-box-roundtrip";
import {
  type ManagedPdfProofingAnnotation,
  type ManagedPdfTextMarkup,
  sanitizeManagedPdfProofingAnnotation,
  sanitizeManagedPdfTextMarkup,
} from "@mrnote/core/web/lib/pdf/native-text-markup-writeback.ts";
import {
  effectiveRotationFor,
  normalizePageEditState,
  normalizeRotation,
  resolvePageItems,
} from "@mrnote/core/web/lib/pdf/page-ops.ts";
import { readPdfSourceTextBoxBinding } from "./source-text-box";
import { sanitizePdfTextBoxEditGeometry } from "@mrnote/core/web/lib/pdf/text-box-rect.ts";
import {
  MANAGED_NATIVE_CHART_PREFIXES,
  parseManagedNativeChartShadow,
  planManagedNativeChartAnnotations,
  type ManagedNativeChartRecord,
} from "./managed-chart";
import {
  MANAGED_NATIVE_TABLE_PREFIX,
  parseManagedNativeTableShadow,
  planManagedNativeTableAnnotations,
  type ManagedNativeTableRecord,
} from "./managed-table";
import { PDF_TEXT_BOX_MAX_COLUMNS, pdfTextBoxTabStops, pdfTextBoxTextPaint } from "@mrnote/core/web/lib/pdf/text-box-typography.ts";
import { normalizePdfTextLineLayout } from "@mrnote/core/web/lib/pdf/text-line-layout.ts";
import { sanitizeInsertedPageId } from "@mrnote/core/web/lib/pdf/inserted-page-identity.ts";
import {
  normalizePdfTextStyleRuns,
  persistablePdfTextStyleRuns,
} from "@mrnote/core/web/lib/pdf/text-style-runs.ts";

export interface PdfAnnotation {
  id: string;
  text: string;
  quote?: string;
  page?: number;
  /** The page inserted in MR Office the comment is on (`page` is then the
   * page before it): lib/pdf/inserted-page-identity. */
  insertedPageId?: string;
  rect?: [number, number, number, number];
  createdAt: string;
  author?: string;
  color?: string;
  nativeSource?: NativePdfAnnotationSource;
  managedMarkup?: ManagedPdfTextMarkup;
  managedProofing?: ManagedPdfProofingAnnotation;
  /** Comment-thread replies (Acrobat / WPS PDF "Reply"); absent for none. */
  replies?: PdfAnnotationReply[];
  /** Present stored reply metadata was lossy or malformed; native saves stop. */
  nativeReplyRoundTripBlocked?: true;
  


  modifiedAt?: string;
}

export interface PdfFormField {
  id: string;
  value: NativePdfFormFieldEdit["value"];
  label?: string;
  page?: number;
  updatedAt?: string;
}

export interface PdfPageTransform {
  page: number;
  rotation: 0 | 90 | 180 | 270;
  updatedAt?: string;
}

export interface PdfInsertedPage {
  id: string;
  afterPage: number;
  width?: number;
  height?: number;
  createdAt?: string;
}

export interface PdfPageMove {
  page: number;
  afterPage: number;
  updatedAt?: string;
}

export function getDataItemId(content: Record<string, unknown> | undefined): string {
  const value = content?.dataItemId || content?.data_item_id;
  return typeof value === "string" ? value : "";
}

export function getStoredAnnotations(
  content: Record<string, unknown> | undefined,
): PdfAnnotation[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  // A malformed reply payload can make the aggregate schema reject the page.
  // Still inspect the raw annotation rows so the reply authority becomes an
  // explicit blocker instead of disappearing from the next rewrite.
  const annotations: unknown[] = parsed.success
    ? parsed.data.annotations
    : Array.isArray(content?.annotations)
      ? content.annotations
      : [];
  const objectAnnotations = annotations.filter(
    (item): item is Record<string, unknown> =>
      !!item && typeof item === "object",
  );
  const managedLinkRecords = objectAnnotations
    .map(parseManagedNativeLinkShadow)
    .filter((record): record is ManagedNativeLinkRecord => Boolean(record));
  const managedLinkPlanSupported =
    planManagedNativeLinkAnnotations(managedLinkRecords).supported;
  const importedLinkRecords = objectAnnotations
    .map(parseImportedNativeLinkShadow)
    .filter((record): record is NativePdfLinkRecord => Boolean(record));
  const importedLinkPlanSupported =
    planImportedNativeLinkAnnotations(importedLinkRecords).supported;
  return objectAnnotations
    .filter((item) => {
      const text = item.text ?? item.comment;
      const managedLink = parseManagedNativeLinkShadow(item);
      const importedLink = parseImportedNativeLinkShadow(item);
      return (
        !isManagedNativeInkAnnotationText(text) &&
        !isManagedNativeShapeAnnotationText(text) &&
        !isManagedNativeMeasurementAnnotationText(text) &&
        !isManagedNativeTextBoxAnnotationText(text) &&
        !isManagedNativeSignatureAnnotationText(text) &&
        !isManagedNativeStampAnnotationText(text) &&
        // A table's, a chart's and a SmartArt's directive is that object, not
        // a comment (getManagedNativeTableRecords / getManagedNativeChartRecords).
        !(typeof text === "string" &&
          (text.startsWith(MANAGED_NATIVE_TABLE_PREFIX) ||
            MANAGED_NATIVE_CHART_PREFIXES.some((prefix) => text.startsWith(prefix)))) &&
        !(
          isManagedNativeLinkAnnotationText(text) &&
          ((managedLink && managedLinkPlanSupported) ||
            (importedLink && importedLinkPlanSupported))
        )
      );
    })
    .map((item) => {
      const managedProofing = sanitizeManagedPdfProofingAnnotation(
        item.managedProofing ?? item.managed_proofing,
      );
      const managedMarkup = sanitizeManagedPdfTextMarkup(
        item.managedMarkup ?? item.managed_markup,
      );
      const replyAuthority = readNativeNoteReplyAuthority(item);
      const insertedPageId = sanitizeInsertedPageId(item.insertedPageId);
      const rawRect = item.rect;
      const rect =
        Array.isArray(rawRect) &&
        rawRect.length === 4 &&
        rawRect.every(
          (coordinate): coordinate is number =>
            typeof coordinate === "number" && Number.isFinite(coordinate),
        ) &&
        rawRect[2] > rawRect[0] &&
        rawRect[3] > rawRect[1]
          ? ([rawRect[0], rawRect[1], rawRect[2], rawRect[3]] as [
              number,
              number,
              number,
              number,
            ])
          : undefined;
      return {
        id: String(item.id || crypto.randomUUID()),
        text: String(item.text || item.comment || ""),
        quote:
          typeof item.quote === "string"
            ? item.quote
            : (managedProofing?.quote ?? managedMarkup?.quote),
        page:
          typeof item.page === "number" && Number.isFinite(item.page)
            ? item.page
            : undefined,
        ...(insertedPageId ? { insertedPageId } : {}),
        ...(rect ? { rect } : {}),
        createdAt: String(
          item.createdAt || item.created_at || new Date().toISOString(),
        ),
        author:
          typeof item.author === "string" && item.author.trim()
            ? item.author.slice(0, 500)
            : undefined,
        color: sanitizeNativePdfMarkupColor(item.color),
        nativeSource: sanitizeNativePdfAnnotationSource(
          item.nativeSource ?? item.native_source,
        ),
        ...(typeof item.modifiedAt === "string" &&
        PDF_ANNOTATION_MODIFIED_AT_PATTERN.test(item.modifiedAt)
          ? { modifiedAt: item.modifiedAt }
          : {}),
        ...(managedProofing ? { managedProofing } : {}),
        ...(managedMarkup ? { managedMarkup } : {}),
        ...(replyAuthority.state === "complete"
          ? { replies: replyAuthority.replies }
          : {}),
        ...(replyAuthority.state === "malformed"
          ? { nativeReplyRoundTripBlocked: true as const }
          : {}),
      };
    })
    .filter(
      (item) =>
        item.text.trim() ||
        item.nativeSource?.kind === "pdf_text_markup_annotation" ||
        item.managedMarkup !== undefined ||
        item.managedProofing !== undefined,
    );
}

export const PDF_ANNOTATION_MODIFIED_AT_PATTERN =
  /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?Z$/;

export function getAnnotations(
  content: Record<string, unknown> | undefined,
): PdfAnnotation[] {
  return partitionManagedNativeNotes(getStoredAnnotations(content)).notes;
}

export function getAnnotationTombstones(
  content: Record<string, unknown> | undefined,
): PdfAnnotation[] {
  return partitionManagedNativeNotes(getStoredAnnotations(content)).tombstones;
}

/** The page's charts (图表, managed-chart.ts), read from their directives. */
export function getManagedNativeChartRecords(
  content: Record<string, unknown> | undefined,
): ManagedNativeChartRecord[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const annotations: unknown[] = parsed.success ? parsed.data.annotations : [];
  return annotations
    .map(parseManagedNativeChartShadow)
    .filter((record): record is ManagedNativeChartRecord => Boolean(record));
}

/** The page's tables (表格, managed-table.ts), read from their directives. */
export function getManagedNativeTableRecords(
  content: Record<string, unknown> | undefined,
): ManagedNativeTableRecord[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const annotations: unknown[] = parsed.success ? parsed.data.annotations : [];
  return annotations
    .map(parseManagedNativeTableShadow)
    .filter((record): record is ManagedNativeTableRecord => Boolean(record));
}

export function getManagedNativeShapeRecords(
  content: Record<string, unknown> | undefined,
): ManagedNativeShapeRecord[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const annotations: unknown[] = parsed.success ? parsed.data.annotations : [];
  return annotations
    .map(parseManagedNativeShapeShadow)
    .filter((record): record is ManagedNativeShapeRecord => Boolean(record));
}

export function partitionManagedNativeMeasurementRecords(
  content: Record<string, unknown> | undefined,
): {
  records: ManagedNativeMeasurementRecord[];
  malformedReservedRows: Record<string, unknown>[];
} {
  const parsed = pdfPageContentSchema.safeParse(content);
  const annotations: unknown[] = parsed.success ? parsed.data.annotations : [];
  const records: ManagedNativeMeasurementRecord[] = [];
  const malformedReservedRows: Record<string, unknown>[] = [];
  for (const annotation of annotations) {
    if (!annotation || typeof annotation !== "object") continue;
    const row = annotation as Record<string, unknown>;
    const text = row.text ?? row.comment;
    if (!isManagedNativeMeasurementAnnotationText(text)) continue;
    const record = parseManagedNativeMeasurementShadow(row);
    if (record) records.push(record);
    else malformedReservedRows.push(row);
  }
  return { records, malformedReservedRows };
}

export function getImportedNativeShapeRecords(
  content: Record<string, unknown> | undefined,
): NativePdfShapeRecord[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const annotations: unknown[] = parsed.success ? parsed.data.annotations : [];
  return annotations
    .map(parseImportedNativeShapeShadow)
    .filter((record): record is NativePdfShapeRecord => Boolean(record));
}

export function getManagedNativeSignatureRecords(
  content: Record<string, unknown> | undefined,
): ManagedNativeSignatureRecord[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const annotations: unknown[] = parsed.success ? parsed.data.annotations : [];
  return annotations
    .map(parseManagedNativeSignatureShadow)
    .filter((record): record is ManagedNativeSignatureRecord =>
      Boolean(record),
    );
}

export function getManagedNativeTextBoxRecords(
  content: Record<string, unknown> | undefined,
): ManagedNativeTextBoxRecord[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const annotations: unknown[] = parsed.success ? parsed.data.annotations : [];
  return annotations
    .map(parseManagedNativeTextBoxShadow)
    .filter((record): record is ManagedNativeTextBoxRecord => Boolean(record));
}

export function getManagedNativeStampRecords(
  content: Record<string, unknown> | undefined,
): ManagedNativeStampRecord[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const annotations: unknown[] = parsed.success ? parsed.data.annotations : [];
  return annotations
    .map(parseManagedNativeStampShadow)
    .filter((record): record is ManagedNativeStampRecord => Boolean(record));
}

export function getManagedNativeLinkRecords(
  content: Record<string, unknown> | undefined,
): ManagedNativeLinkRecord[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const annotations: unknown[] = parsed.success ? parsed.data.annotations : [];
  return annotations
    .map(parseManagedNativeLinkShadow)
    .filter((record): record is ManagedNativeLinkRecord => Boolean(record));
}

export function getImportedNativeLinkRecords(
  content: Record<string, unknown> | undefined,
): NativePdfLinkRecord[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const annotations: unknown[] = parsed.success ? parsed.data.annotations : [];
  return annotations
    .map(parseImportedNativeLinkShadow)
    .filter((record): record is NativePdfLinkRecord => Boolean(record));
}

export function getImportedNativeStampRecords(
  content: Record<string, unknown> | undefined,
): NativePdfStampRecord[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const annotations: unknown[] = parsed.success ? parsed.data.annotations : [];
  return annotations
    .map(parseImportedNativeStampShadow)
    .filter((record): record is NativePdfStampRecord => Boolean(record));
}

export function getFormFields(
  content: Record<string, unknown> | undefined,
): PdfFormField[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const fields: unknown[] = parsed.success
    ? parsed.data.formFields || parsed.data.form_fields || []
    : [];
  return fields
    .filter(
      (item): item is Record<string, unknown> =>
        !!item && typeof item === "object",
    )
    .map((item) => ({
      id: String(item.id || item.field_id || ""),
      value: Array.isArray(item.value)
        ? item.value.filter(
            (value): value is string => typeof value === "string",
          )
        : typeof item.value === "boolean"
          ? item.value
          : String(item.value ?? ""),
      // Keep the in-memory draft identical to its JSON wire shape. Own keys
      // set to undefined disappear over HTTP and falsely reject exact receipts.
      ...(typeof item.label === "string" ? { label: item.label } : {}),
      ...(typeof item.page === "number" && Number.isFinite(item.page)
        ? { page: item.page }
        : {}),
      updatedAt: String(item.updatedAt || item.updated_at || ""),
    }))
    .filter((item) => item.id.trim());
}

export function getNativePdfFormAuthoringOperations(
  content: Record<string, unknown> | undefined,
): readonly NativePdfFormAuthoringOperation[] | null {
  return readNativePdfFormAuthoringOperations(
    content?.pdfFormAuthoringOperations ??
      content?.pdf_form_authoring_operations,
  );
}

export function getPageTransforms(
  content: Record<string, unknown> | undefined,
): PdfPageTransform[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const transforms: unknown[] = parsed.success
    ? parsed.data.pageTransforms || parsed.data.page_transforms || []
    : [];
  return transforms
    .filter(
      (item): item is Record<string, unknown> =>
        !!item && typeof item === "object",
    )
    .map((item) => ({
      page: Number(item.page),
      rotation: normalizeRotation(item.rotation),
      updatedAt: String(item.updatedAt || item.updated_at || ""),
    }))
    .filter((item) => Number.isInteger(item.page) && item.page >= 1);
}

export function getDeletedPages(
  content: Record<string, unknown> | undefined,
): number[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const pages: unknown[] = parsed.success
    ? parsed.data.deletedPages || parsed.data.deleted_pages || []
    : [];
  return Array.from(
    new Set(
      pages
        .map((page) => Number(page))
        .filter(
          (page) => Number.isInteger(page) && page >= 1 && page <= 10_000,
        ),
    ),
  ).sort((a, b) => a - b);
}

export function getInsertedPages(
  content: Record<string, unknown> | undefined,
): PdfInsertedPage[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const pages: unknown[] = parsed.success
    ? parsed.data.insertedPages || parsed.data.inserted_pages || []
    : [];
  return pages
    .filter(
      (item): item is Record<string, unknown> =>
        !!item && typeof item === "object",
    )
    .map((item) => ({
      id: String(item.id || crypto.randomUUID()),
      afterPage: Number(item.afterPage ?? item.after_page ?? 0),
      width:
        typeof item.width === "number" && Number.isFinite(item.width)
          ? item.width
          : undefined,
      height:
        typeof item.height === "number" && Number.isFinite(item.height)
          ? item.height
          : undefined,
      createdAt: String(item.createdAt || item.created_at || ""),
    }))
    .filter((item) => Number.isInteger(item.afterPage) && item.afterPage >= 0);
}

export function getPageMoves(
  content: Record<string, unknown> | undefined,
): PdfPageMove[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const moves: unknown[] = parsed.success
    ? parsed.data.pageMoves || parsed.data.page_moves || []
    : [];
  return moves
    .filter(
      (item): item is Record<string, unknown> =>
        !!item && typeof item === "object",
    )
    .map((item) => ({
      page: Number(item.page),
      afterPage: Number(item.afterPage ?? item.after_page ?? 0),
      updatedAt: String(item.updatedAt || item.updated_at || ""),
    }))
    .filter(
      (item) =>
        Number.isInteger(item.page) &&
        item.page >= 1 &&
        Number.isInteger(item.afterPage) &&
        item.afterPage >= 0 &&
        item.afterPage !== item.page,
    );
}

export function getTextBoxEdits(
  content: Record<string, unknown> | undefined,
): PdfTextBoxEdit[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const edits: unknown[] = parsed.success
    ? parsed.data.textBoxEdits || parsed.data.text_box_edits || []
    : Array.isArray(content?.textBoxEdits)
      ? content.textBoxEdits
      : Array.isArray(content?.text_box_edits)
        ? content.text_box_edits
        : [];
  const normalized = edits
    .filter(
      (item): item is Record<string, unknown> =>
        !!item && typeof item === "object",
    )
    .map((item): PdfTextBoxEdit | null => {
      const align = item.align;
      // A source line's glyph whose ToUnicode is U+0000 (a bullet) is kept by
      // the API as U+FFFD (sanitize_jsonb_text). Read back as U+FFFD, the box
      // measured it as a visible character -- its row wrapped, its frame grew
      // -- and the writer refused the whole page's next save for a glyph no
      
      // source line reads it as it was written: U+0000, no glyph, no width.
      const restoreGlyphs = (value: string) =>
        item.sourceText && value.includes("\ufffd") ? value.replaceAll("\ufffd", "\u0000") : value;
      const text = typeof item.text === "string" ? restoreGlyphs(item.text) : "";
      const storedRuns = item.sourceText && Array.isArray(item.runs)
        ? item.runs.map((run: unknown) =>
          run && typeof run === "object" && typeof (run as { text?: unknown }).text === "string"
            ? { ...(run as Record<string, unknown>), text: restoreGlyphs((run as { text: string }).text) }
            : run)
        : item.runs;
      const num = (v: unknown): number | undefined =>
        Number.isFinite(Number(v)) ? Number(v) : undefined;
      const rawPage = item.page;
      const page = rawPage === undefined ? 0 : Number(rawPage);
      return normalizePdfTextBoxEdit({
        id: String(item.id || ""),
        sourceText: readPdfSourceTextBoxBinding(item.sourceText) ?? undefined,
        page: Number.isInteger(page) && page >= 0 ? page : -1,
        x: Number(item.x) || 0,
        y: Number(item.y) || 0,
        width: Number(item.width) || 0,
        height: Number(item.height) || 0,
        // Rotation is part of the managed text box's durable geometry. The
        // native V2 directive and PDF appearance preserve it, so dropping it
        // while rehydrating page JSON would make the editor reopen upright.
        angle: num(item.angle),
        writingMode:
          item.writingMode === "vertical-rl" ? "vertical-rl" : undefined,
        text,
        fontSize: typeof item.fontSize === "number" ? item.fontSize : undefined,
        color: typeof item.color === "string" ? item.color : undefined,
        backgroundColor:
          typeof item.backgroundColor === "string"
            ? item.backgroundColor
            : undefined,
        align:
          align === "left" || align === "center" || align === "right"
            ? align
            : undefined,
        kind:
          item.kind === "formula" || item.kind === "text"
            ? item.kind
            : undefined,
        latex: typeof item.latex === "string" ? item.latex : undefined,
        // Carry the full style + provenance set so bold/italic/underline, the
        // created flag, and the stationary cover origin survive a save/reload
        // (the prior subset map silently dropped them — bold text reverted to
        // normal, created boxes lost their identity, dragged covers re-revealed
        // the original glyphs).
        fontFamily:
          typeof item.fontFamily === "string" ? item.fontFamily : undefined,
        sourceFontName:
          typeof item.sourceFontName === "string"
            ? item.sourceFontName
            : undefined,
        // The Office family the box is set in (text-box-runs.ts); the writer
        // checks it against the catalog again.
        officeFamily:
          typeof item.officeFamily === "string" && item.officeFamily.trim()
            ? item.officeFamily.slice(0, 128)
            : undefined,
        bold: typeof item.bold === "boolean" ? item.bold : undefined,
        italic: typeof item.italic === "boolean" ? item.italic : undefined,
        underline:
          typeof item.underline === "boolean" ? item.underline : undefined,
        strikethrough: item.strikethrough === true ? true : undefined,
        justify: item.justify === true ? true : undefined,
        // Whole-box Word formats (text-box-typography.ts); the writer
        // validates their ranges again.
        textScript:
          item.textScript === "superscript" || item.textScript === "subscript"
            ? item.textScript
            : undefined,
        highlight:
          typeof item.highlight === "string" && /^#[0-9A-Fa-f]{6}$/u.test(item.highlight)
            ? item.highlight
            : undefined,
        lineSpacing: num(item.lineSpacing),
        characterSpacing: num(item.characterSpacing),
        wordSpacing: num(item.wordSpacing),
        horizontalScale: num(item.horizontalScale),
        leftIndent: num(item.leftIndent),
        rightIndent: num(item.rightIndent),
        firstLineIndent: num(item.firstLineIndent),
        distributed: item.distributed === true ? true : undefined,
        verticalAlign:
          item.verticalAlign === "middle" || item.verticalAlign === "bottom"
            ? item.verticalAlign
            : undefined,
        // 分栏 and 艺术字: dropped here, a reload set the box in one column
        // and plain ink, and its next save wrote it so.
        columnCount:
          typeof item.columnCount === "number" && Number.isInteger(item.columnCount) &&
            item.columnCount > 1 && item.columnCount <= PDF_TEXT_BOX_MAX_COLUMNS
            ? item.columnCount
            : undefined,
        columnSpacing: num(item.columnSpacing),
        // Word's paragraph Spacing Before / After (text-box-typography.ts).
        spaceBefore: num(item.spaceBefore),
        spaceAfter: num(item.spaceAfter),
        // Word's Tabs (text-box-typography.ts), kept with the box.
        tabStops: item.tabStops === undefined ? undefined
          : pdfTextBoxTabStops(item.tabStops, typeof item.fontSize === "number" ? item.fontSize : 12 / 792) ?? undefined,
        defaultTabStop: num(item.defaultTabStop),
        textPaint: item.textPaint === undefined ? undefined : pdfTextBoxTextPaint(item.textPaint) ?? undefined,
        // Preserve exact source-font spans across save/reload. This reader
        // used to reconstruct only the box-level style, silently discarding
        // schema-validated runs even though the writer persisted them.
        runs: normalizePdfTextStyleRuns(storedRuns, {
          expectedText: text,
        }),
        lineLayout: normalizePdfTextLineLayout(item.lineLayout, text),
        created: typeof item.created === "boolean" ? item.created : undefined,
        // Deleted extracted lines persist as mask-only records; dropping the
        // flag here made every deletion silently revert on reload (deleted
        // formulas visually resurrected from the latex fallback).
        deleted: typeof item.deleted === "boolean" ? item.deleted : undefined,
        originX: num(item.originX),
        originY: num(item.originY),
        originWidth: num(item.originWidth),
        originHeight: num(item.originHeight),
        nativeSource: sanitizeNativePdfFreeTextAnnotationSource(
          item.nativeSource ?? item.native_source,
        ),
        sourceAction:
          item.sourceAction === "content_style_geometry" ||
          item.source_action === "content_style_geometry"
            ? "content_style_geometry"
            : undefined,
        // Dropping this on reload would re-persist the overlay's fallback
        // alignment as a user choice and flatten the author's `/Q`.
        alignFromSource:
          item.alignFromSource === true || item.align_from_source === true
            ? true
            : undefined,
        sourceBorderWidth: num(
          item.sourceBorderWidth ?? item.source_border_width,
        ),
        nativeAppearanceImage:
          sanitizeManagedNativeTextBoxImage(
            item.nativeAppearanceImage ?? item.native_appearance_image,
          ) ?? undefined,
      });
    })
    .filter((item): item is PdfTextBoxEdit => Boolean(item));
  const sourceShadows = (parsed.success ? parsed.data.annotations : [])
    .map(parseImportedNativeTextBoxShadow)
    .filter((record): record is NativePdfFreeTextAnnotation => Boolean(record));
  if (sourceShadows.length === 0) return normalized;
  const shadowsByRef = new Map(
    sourceShadows.map((record) => [record.nativeSource.objectRef, record]),
  );
  const claimed = new Set<string>();
  const hydrated = normalized.map((edit) => {
    const source = sanitizeNativePdfFreeTextAnnotationSource(edit.nativeSource);
    if (!source || edit.deleted) return edit;
    const shadow = shadowsByRef.get(source.objectRef);
    if (!shadow) return edit;
    claimed.add(source.objectRef);
    return {
      ...edit,
      ...shadow,
      backgroundColor: edit.backgroundColor,
    };
  });
  for (const sourceShadow of sourceShadows) {
    if (!claimed.has(sourceShadow.nativeSource.objectRef)) {
      hydrated.push(sourceShadow);
    }
  }
  return hydrated;
}

export function getImageOverlays(
  content: Record<string, unknown> | undefined,
): PdfImageOverlayItem[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const overlays: unknown[] = parsed.success
    ? parsed.data.imageOverlays || parsed.data.image_overlays || []
    : [];
  return overlays
    .filter(
      (item): item is Record<string, unknown> =>
        !!item && typeof item === "object",
    )
    .map((item) => ({
      id: String(item.id || ""),
      page: Number(item.page) || 0,
      dataItemId: String(item.dataItemId || item.data_item_id || ""),
      x: Number(item.x) || 0,
      y: Number(item.y) || 0,
      width: Number(item.width) || 0,
      height: Number(item.height) || 0,
    }))
    .filter((item) => item.id.trim() && item.dataItemId.trim());
}

export function getStoredInkStrokes(
  content: Record<string, unknown> | undefined,
): PdfInkStroke[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const strokes: unknown[] = parsed.success
    ? parsed.data.inkStrokes || parsed.data.ink_strokes || []
    : [];
  return strokes
    .filter(
      (item): item is Record<string, unknown> =>
        !!item && typeof item === "object",
    )
    .map((item) => {
      const nativeSource = sanitizeNativePdfInkAnnotationSource(
        item.nativeSource ?? item.native_source,
      );
      const insertedPageId = sanitizeInsertedPageId(item.insertedPageId);
      return {
        id: String(item.id || ""),
        page: Number(item.page) || 0,
        ...(insertedPageId ? { insertedPageId } : {}),
        color: typeof item.color === "string" ? item.color : "#dc2626",
        width: typeof item.width === "number" ? item.width : 2,
        points: Array.isArray(item.points)
          ? item.points.filter((n): n is number => typeof n === "number")
          : [],
        ...(item.penType !== undefined || item.pen_type !== undefined
          ? { penType: normalizeInkPenType(item.penType ?? item.pen_type) }
          : {}),
        ...(item.deleted === true ? { deleted: true } : {}),
        ...(nativeSource ? { nativeSource } : {}),
      };
    })
    .filter((item) => item.id.trim() && item.points.length >= 4);
}

export function getInkStrokes(
  content: Record<string, unknown> | undefined,
): PdfInkStroke[] {
  return getStoredInkStrokes(content).filter((stroke) => !stroke.deleted);
}

export function getInkStrokeTombstones(
  content: Record<string, unknown> | undefined,
): PdfInkStroke[] {
  return getStoredInkStrokes(content).filter((stroke) => stroke.deleted);
}

export function asLiveInkStroke(stroke: PdfInkStroke): PdfInkStroke {
  const live = { ...stroke };
  delete live.deleted;
  return live;
}

export function readerInkToManagedStroke(stroke: InkAnnotation): PdfInkStroke {
  return {
    id: stroke.id,
    page: stroke.page - 1,
    color: stroke.color,
    width: stroke.width,
    points: stroke.points.flatMap((point) => [point.x / 100, point.y / 100]),
    penType: "pen",
    ...(stroke.insertedPageId ? { insertedPageId: stroke.insertedPageId } : {}),
  };
}

export function getCropBoxes(
  content: Record<string, unknown> | undefined,
): CropBoxEntry[] {
  const parsed = pdfPageContentSchema.safeParse(content);
  const crops: unknown[] = parsed.success
    ? parsed.data.cropBoxes || parsed.data.crop_boxes || []
    : [];
  return crops
    .filter(
      (item): item is Record<string, unknown> =>
        !!item && typeof item === "object",
    )
    .map((item) => sanitizeCropBox(item))
    .filter((item): item is CropBoxEntry => Boolean(item));
}

export interface PdfPageSizeEntry {
  page: number;
  sourceWidth: number;
  sourceHeight: number;
  width: number;
  height: number;
  contentPlacement: "unchanged" | "fit-center";
}

export function readPdfPageSizeEntries(content: Record<string, unknown> | undefined): {
  entries: PdfPageSizeEntry[];
  malformed: boolean;
} {
  const camelCase = content?.pageSizes;
  const snakeCase = content?.page_sizes;
  if (
    camelCase !== undefined &&
    snakeCase !== undefined &&
    JSON.stringify(camelCase) !== JSON.stringify(snakeCase)
  ) {
    return { entries: [], malformed: true };
  }
  const raw = camelCase ?? snakeCase;
  if (raw === undefined) return { entries: [], malformed: false };
  if (!Array.isArray(raw) || raw.length > 500) {
    return { entries: [], malformed: true };
  }
  const seen = new Set<number>();
  const entries: PdfPageSizeEntry[] = [];
  for (const value of raw) {
    if (!value || typeof value !== "object" || Array.isArray(value)) {
      return { entries: [], malformed: true };
    }
    const record = value as Record<string, unknown>;
    const keys = Object.keys(record).sort().join(",");
    if (
      (keys !== "height,page,sourceHeight,sourceWidth,width" &&
        keys !==
          "contentPlacement,height,page,sourceHeight,sourceWidth,width") ||
      !Number.isSafeInteger(record.page) ||
      (record.page as number) < 1 ||
      seen.has(record.page as number) ||
      typeof record.width !== "number" ||
      !Number.isFinite(record.width) ||
      record.width < 36 ||
      record.width > 14_400 ||
      typeof record.height !== "number" ||
      !Number.isFinite(record.height) ||
      record.height < 36 ||
      record.height > 14_400 ||
      typeof record.sourceWidth !== "number" ||
      !Number.isFinite(record.sourceWidth) ||
      record.sourceWidth < 36 ||
      record.sourceWidth > 14_400 ||
      typeof record.sourceHeight !== "number" ||
      !Number.isFinite(record.sourceHeight) ||
      record.sourceHeight < 36 ||
      record.sourceHeight > 14_400 ||
      (record.contentPlacement !== undefined &&
        record.contentPlacement !== "unchanged" &&
        record.contentPlacement !== "fit-center")
    ) {
      return { entries: [], malformed: true };
    }
    seen.add(record.page as number);
    entries.push({
      page: record.page as number,
      sourceWidth: record.sourceWidth,
      sourceHeight: record.sourceHeight,
      width: record.width,
      height: record.height,
      contentPlacement:
        record.contentPlacement === "fit-center" ? "fit-center" : "unchanged",
    });
  }
  entries.sort((left, right) => left.page - right.page);
  return { entries, malformed: false };
}

export function setPdfPageSizes(
  entries: readonly PdfPageSizeEntry[],
  pages: readonly number[],
  sourceSizes: Readonly<Record<number, FlattenPageSize>>,
  width: number | null,
  height: number | null,
  contentPlacement: "unchanged" | "fit-center",
  




  quarterTurnedPages?: ReadonlySet<number>,
): PdfPageSizeEntry[] | null {
  const byPage = new Map(entries.map((entry) => [entry.page, entry]));
  for (const page of pages) {
    const existing = byPage.get(page);
    const source = existing ?? sourceSizes[page];
    if (!source) return null;
    const sourceWidth = existing?.sourceWidth ?? source.width;
    const sourceHeight = existing?.sourceHeight ?? source.height;
    const turned = width !== null && height !== null && quarterTurnedPages?.has(page) === true;
    byPage.set(page, {
      page,
      sourceWidth,
      sourceHeight,
      width: (turned ? height : width) ?? sourceWidth,
      height: (turned ? width : height) ?? sourceHeight,
      contentPlacement:
        width === null || height === null ? "unchanged" : contentPlacement,
    });
  }
  return [...byPage.values()].sort((left, right) => left.page - right.page);
}

export function getPdfDocumentDecorations(
  content: Record<string, unknown> | undefined,
): PdfDocumentDecorations {
  return sanitizePdfDocumentDecorations(
    content?.pdfDecorations ?? content?.pdf_decorations,
  );
}

export function buildPdfPageEditContent(
  dataItemId: string,
  annotations: PdfAnnotation[],
  annotationTombstones: PdfAnnotation[],
  formFields: PdfFormField[],
  pageTransforms: PdfPageTransform[],
  deletedPages: number[],
  insertedPages: PdfInsertedPage[],
  pageMoves: PdfPageMove[],
  textBoxEdits: PdfTextBoxEdit[],
  nativeTextBoxes: ManagedNativeTextBoxRecord[],
  imageOverlays: PdfImageOverlayItem[],
  inkStrokes: PdfInkStroke[],
  inkStrokeTombstones: PdfInkStroke[],
  nativeShapes: ManagedNativeShapeRecord[],
  nativeMeasurements: ManagedNativeMeasurementRecord[],
  nativeSourceShapes: NativePdfShapeRecord[],
  nativeSignatures: ManagedNativeSignatureRecord[],
  nativeStamps: ManagedNativeStampRecord[],
  nativeSourceStamps: NativePdfStampRecord[],
  nativeLinks: ManagedNativeLinkRecord[],
  nativeSourceLinks: NativePdfLinkRecord[],
  nativeTables: readonly ManagedNativeTableRecord[],
  nativeCharts: readonly ManagedNativeChartRecord[],
  pageSizes: PdfPageSizeEntry[],
  cropBoxes: CropBoxEntry[],
  pdfDecorations: PdfDocumentDecorations,
  source?: Record<string, unknown>,
): Record<string, unknown> {
  const rotationMap = new Map(
    pageTransforms.map((transform) => [transform.page, transform.rotation]),
  );
  const nativeInkPlan = planManagedNativeInkAnnotations(
    [...inkStrokes, ...inkStrokeTombstones],
    rotationMap,
  );
  const nativeShapePlan = planManagedNativeShapeAnnotations(
    nativeShapes,
    rotationMap,
  );
  const nativeMeasurementPlan = planManagedNativeMeasurementAnnotations(
    nativeMeasurements,
    rotationMap,
  );
  const nativeSourceShapePlan = planImportedNativeShapeAnnotations(
    nativeSourceShapes,
    rotationMap,
  );
  const nativeSignaturePlan = planManagedNativeSignatureAnnotations(
    nativeSignatures,
    rotationMap,
  );
  const nativeTextBoxPlan = planManagedNativeTextBoxAnnotations(
    nativeTextBoxes,
    rotationMap,
  );
  const nativeSourceTextBoxPlan = planImportedNativeTextBoxAnnotations(
    textBoxEdits,
    rotationMap,
  );
  const nativeStampPlan = planManagedNativeStampAnnotations(
    nativeStamps,
    rotationMap,
  );
  const nativeSourceStampPlan = planImportedNativeStampAnnotations(
    nativeSourceStamps,
    rotationMap,
  );
  const nativeLinkPlan = planManagedNativeLinkAnnotations(
    nativeLinks,
    rotationMap,
  );
  const nativeSourceLinkPlan = planImportedNativeLinkAnnotations(
    nativeSourceLinks,
    rotationMap,
  );
  const nativeTablePlan = planManagedNativeTableAnnotations(nativeTables, rotationMap);
  const nativeChartPlan = planManagedNativeChartAnnotations(nativeCharts, rotationMap);
  const content: Record<string, unknown> = {
    type: PDF_CONTENT_TYPE,
    dataItemId,
    annotations: [
      ...serializeManagedNativeNotes(annotations, annotationTombstones),
      ...nativeInkPlan.annotations,
      ...nativeShapePlan.annotations,
      ...nativeMeasurementPlan.annotations,
      ...nativeSourceShapePlan.annotations,
      ...nativeTextBoxPlan.annotations,
      ...nativeSourceTextBoxPlan.annotations,
      ...nativeSignaturePlan.annotations,
      ...nativeStampPlan.annotations,
      ...nativeSourceStampPlan.annotations,
      ...nativeLinkPlan.annotations,
      ...nativeSourceLinkPlan.annotations,
      ...nativeTablePlan.annotations,
      ...nativeChartPlan.annotations,
    ],
    formFields,
    pageTransforms,
    deletedPages,
    insertedPages,
    pageMoves,
    textBoxEdits: normalizePdfTextBoxEdits(textBoxEdits),
    imageOverlays,
    inkStrokes: [...inkStrokes, ...inkStrokeTombstones],
    pageSizes,
    cropBoxes,
    source: source || {},
  };
  if (hasPdfDocumentDecorations(pdfDecorations)) {
    content.pdfDecorations = sanitizePdfDocumentDecorations(pdfDecorations);
  }
  return content;
}

export function normalizePdfTextBoxEdit(edit: PdfTextBoxEdit): PdfTextBoxEdit | null {
  const geometry = sanitizePdfTextBoxEditGeometry(edit);
  if (!geometry || !edit.id.trim() || edit.page < 0) return null;
  const base: PdfTextBoxEdit = { ...edit, ...geometry };
  const persistableRuns = persistablePdfTextStyleRuns(edit.runs, edit.text);
  if (persistableRuns) base.runs = persistableRuns;
  else delete base.runs;
  if (isNativePdfFreeTextRecord(base)) {
    // The inert native directive already carries the bounded PNG payload.
    // Avoid duplicating up to 1 MiB per box inside `textBoxEdits`; reload
    // hydrates it back from `parseImportedNativeTextBoxShadow`.
    delete base.nativeAppearanceImage;
  }
  if (geometry.originX === undefined) {
    delete base.originX;
    delete base.originY;
    delete base.originWidth;
    delete base.originHeight;
  }
  const kind = edit.kind ?? (edit.latex ? "formula" : undefined);
  if (kind === "text") {
    const normalized: PdfTextBoxEdit = { ...base, kind: "text" };
    delete normalized.latex;
    return normalized;
  }
  if (kind === "formula") {
    return { ...base, kind: "formula" };
  }
  return base;
}

export function normalizePdfTextBoxEdits(
  edits: readonly PdfTextBoxEdit[],
): PdfTextBoxEdit[] {
  return edits
    .map(normalizePdfTextBoxEdit)
    .filter((item): item is PdfTextBoxEdit => Boolean(item));
}












export function resolvePdfDecorationPageSizes(
  content: Record<string, unknown>,
  originalPageCount: number,
  nativePageSizes: Readonly<Record<number, FlattenPageSize>>,
): { width: number; height: number; rotation: 0 | 90 | 180 | 270 }[] {
  if (originalPageCount < 1 || originalPageCount > 10_000) return [];
  const state = normalizePageEditState({
    pageRotations: Object.fromEntries(
      getPageTransforms(content).map((entry) => [entry.page, entry.rotation]),
    ),
    deletedPages: getDeletedPages(content),
    pageMoves: getPageMoves(content),
    insertedPages: getInsertedPages(content),
  });
  const originalPages = Array.from(
    { length: originalPageCount },
    (_, index) => index + 1,
  );
  const sourceSizesByOriginal = new Map<
    number,
    { width: number; height: number; rotation: 0 | 90 | 180 | 270 }
  >();
  for (const originalPage of originalPages) {
    const source = nativePageSizes[originalPage];
    if (!source) continue;
    sourceSizesByOriginal.set(originalPage, {
      width: source.width,
      height: source.height,
      rotation: normalizeRotation(source.rotation),
    });
  }
  const pageSizeByOriginal = new Map(
    readPdfPageSizeEntries(content).entries.map((entry) => [
      entry.page,
      entry,
    ]),
  );
  return resolvePageItems(state, originalPages).map((item) => {
    if (item.kind === "inserted") {
      return {
        width: item.page.width,
        height: item.page.height,
        rotation: 0 as const,
      };
    }
    const source = sourceSizesByOriginal.get(item.page);
    const resized = pageSizeByOriginal.get(item.page);
    return {
      width: resized?.width ?? source?.width ?? 0,
      height: resized?.height ?? source?.height ?? 0,
      rotation: effectiveRotationFor(state, item.page, source?.rotation),
    };
  });
}
