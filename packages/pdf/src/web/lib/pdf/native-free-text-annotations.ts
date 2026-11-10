/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { TextBoxEditLike } from "./text-box-flatten";
import {
  nativePdfAnnotationIsInvisible,
  readNativePdfAnnotationFlags,
  sanitizePdfText,
  type NativePdfAnnotationFlags,
} from "@mrnote/core/web/lib/pdf/native-text-annotations.ts";

/**
 * Editable third-party `/FreeText` annotations discovered through pdf.js.
 *
 * The source object stays authoritative until the user changes it.  We retain
 * the exact indirect-object reference so native patchback updates or deletes
 * that object instead of flattening a visually similar replacement over it.
 */

export const NATIVE_PDF_FREE_TEXT_SOURCE_KIND = "pdf_free_text_annotation";
export const MAX_NATIVE_PDF_FREE_TEXT_ANNOTATIONS_PER_PAGE = 1_024;
export const MAX_NATIVE_PDF_FREE_TEXT_LENGTH = 20_000;

/** Matches the backend's `_MAX_FREE_TEXT_BORDER_WIDTH` (PDF points). */
export const MAX_NATIVE_PDF_FREE_TEXT_BORDER_WIDTH = 12;
/** pdf.js `AnnotationBorderStyleType.SOLID`. */
const SOLID_ANNOTATION_BORDER_STYLE = 1;

const PDF_OBJECT_REF_RE = /^[1-9][0-9]{0,9}R(?:[0-9]{1,5})?$/u;
const CONTROL_CHARS_RE = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/gu;
const LOCKED_ANNOTATION_FLAG = 1 << 7;
const LOCKED_CONTENTS_ANNOTATION_FLAG = 1 << 9;

export interface NativePdfFreeTextAnnotationSource {
  kind: typeof NATIVE_PDF_FREE_TEXT_SOURCE_KIND;
  objectRef: string;
  subtype: "FreeText";
}

export type NativePdfFreeTextSourceAction = "content_style_geometry";

export interface NativePdfFreeTextAnnotation extends TextBoxEditLike {
  id: string;
  /** Zero-based page index, matching the inline text editor. */
  page: number;
  x: number;
  y: number;
  width: number;
  height: number;
  text: string;
  fontSize: number;
  color: string;
  align: "left" | "center" | "right";
  /**
   * True while the source `/Q` is still the authority for the saved object.
   *
   * pdf.js does not report FreeText quadding, so `align` above is only the
   * overlay's layout fallback — never something the user confirmed. Saving it
   * would rewrite a centred or right-aligned author's box to left on the first
   * drag, so the native directive asks the backend to keep `/Q` untouched
   * until the user picks an alignment (which clears this flag).
   */
  alignFromSource?: true;
  /**
   * Solid source border, as a fraction of page height (same unit as
   * `fontSize`). Its colour lives in the source `/DA` stroke operators, which
   * pdf.js does not report, so the overlay draws it in the text colour while
   * the saved appearance is rebuilt from the object's own colour.
   */
  sourceBorderWidth?: number;
  





  sourceInsets?: { left: number; top: number; right: number; bottom: number };
  

  sourceOpacity?: number;
  fontFamily: string;
  sourceFontName: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  /**
   * Source annotations behave like additive boxes inside the shared overlay:
   * their original `/AP` is hidden through pdf.js annotation storage, so no
   * extracted-page glyph mask is needed.
   */
  created: true;
  originX: number;
  originY: number;
  originWidth: number;
  originHeight: number;
  nativeSource: NativePdfFreeTextAnnotationSource;
  sourceAction?: NativePdfFreeTextSourceAction;
  nativeAppearanceImage?: string;
  deleted?: false;
}

export interface NativePdfFreeTextTombstone extends Omit<
  NativePdfFreeTextAnnotation,
  "deleted" | "sourceAction"
> {
  deleted: true;
}

export type NativePdfFreeTextRecord =
  NativePdfFreeTextAnnotation | NativePdfFreeTextTombstone;

export type NativePdfFreeTextImportLossReason =
  | "invalid_object_ref"
  | "duplicate_object_ref"
  | "invalid_page_geometry"
  | "invalid_text_box_geometry"
  | "invalid_text"
  | "unsupported_intent"
  | "unsupported_rotation"
  | "unsupported_rich_text"
  | "unsupported_border_style"
  | "unsupported_typography"
  | "locked_annotation"
   
  | "hidden_annotation"
  /** An MRNote replacement text box (it covers source text it hides): only
   *  its source-text binding may write it, so it is not offered here. */
  | "source_text_owner"
  /** A table or chart made in MRAI: its own layer edits it. */
  | "managed_object"
  | "annotation_limit";

/**
 * What only the source bytes tell about a FreeText (pdf.js reports neither):
 * its /RD inner rectangle, and the BaseFont its /DA font resource names.
 */
export interface NativePdfFreeTextSourceFacts {
  /** /RD: left, top, right, bottom insets in points (ISO 32000-1 table 177). */
  rectangleDifferences?: readonly [number, number, number, number];
  


  appearanceOpacity?: number;
  /** The font dictionary's /BaseFont the /DA resource name points to. */
  baseFont?: string;
  




  sourceTextOwner?: boolean;
  /** A table, chart or SmartArt made in MRAI (/MRNoteTable, /MRNoteChart,
   *  /MRNoteSmartArt). */
  managedObject?: boolean;
  

  managedTextBox?: boolean;
}

export interface NativePdfFreeTextImportLoss {
  objectRef?: string;
  reason: NativePdfFreeTextImportLossReason;
}

export interface NativePdfFreeTextImportResult {
  textBoxes: NativePdfFreeTextAnnotation[];
  losses: NativePdfFreeTextImportLoss[];
}

export interface NativePdfFreeTextViewport {
  width: number;
  height: number;
  /** CSS pixels per PDF user-space point for the active viewport. */
  scale: number;
  convertToViewportRectangle: (
    rect: readonly number[],
  ) => readonly [number, number, number, number];
  /**
   * Width of one line in PDF points at the given size. With it the line's
   * place in the source appearance tells how the box is aligned; without it
   * the overlay falls back to left.
   */
  measureTextWidth?: (
    text: string,
    font: { fontFamily: string; bold: boolean; italic: boolean },
    fontSizePoints: number,
  ) => number;
}











export function freeTextAlignmentFromAppearance(input: {
  /** Box width in points. */
  boxWidth: number;
  /** Where the first line starts, in points from the box's left edge. */
  lineStart: number;
  /** The first line's width in points. */
  lineWidth: number;
  fontSize: number;
}): "left" | "center" | "right" | null {
  const { boxWidth, lineStart, lineWidth, fontSize } = input;
  if (
    ![boxWidth, lineStart, lineWidth, fontSize].every(Number.isFinite) ||
    boxWidth <= 0 ||
    lineWidth <= 0 ||
    fontSize <= 0
  ) {
    return null;
  }
  const slack = boxWidth - lineWidth;
  // The three placements are half the slack apart. Substitute-font metrics
  // and the box's padding move the start by a few points, so below this
  // much room the placements cannot be told apart, nor seen.
  if (slack < Math.max(12, fontSize)) return null;
  const distances = [
    ["left", Math.abs(lineStart)],
    ["center", Math.abs(lineStart - slack / 2)],
    ["right", Math.abs(lineStart - slack)],
  ] as const;
  const [nearest] = [...distances].sort((a, b) => a[1] - b[1]);
  // Halfway between two placements is neither; a start has to sit clearly
  // on one of them.
  return nearest[1] <= slack / 6 ? nearest[0] : null;
}

interface RawPdfString {
  str?: unknown;
}

interface RawDefaultAppearanceData {
  fontColor?: unknown;
  fontName?: unknown;
  fontSize?: unknown;
}

interface RawNativePdfFreeTextAnnotation {
  subtype?: unknown;
  id?: unknown;
  /** pdf.js normalizes the annotation's `/C` entry to RGB channels. */
  color?: unknown;
  rect?: unknown;
  contentsObj?: unknown;
  textContent?: unknown;
  defaultAppearanceData?: unknown;
  annotationFlags?: unknown;
  /** pdf.js `opacity`: the annotation's /CA. */
  opacity?: unknown;
  rotation?: unknown;
  it?: unknown;
  richText?: unknown;
  borderStyle?: unknown;
  /** First glyph of the source appearance, in points from the rect's corner. */
  textPosition?: unknown;
}

function isPdfObjectRef(value: unknown): value is string {
  return typeof value === "string" && PDF_OBJECT_REF_RE.test(value);
}

function safeText(value: unknown): string {
  const raw =
    typeof value === "string"
      ? value
      : value &&
          typeof value === "object" &&
          typeof (value as RawPdfString).str === "string"
        ? ((value as RawPdfString).str as string)
        : "";
  return raw
    .replace(/\r\n?/gu, "\n")
    .replace(CONTROL_CHARS_RE, "")
    .slice(0, MAX_NATIVE_PDF_FREE_TEXT_LENGTH);
}

function annotationText(annotation: RawNativePdfFreeTextAnnotation): string {
  const contents = safeText(annotation.contentsObj);
  if (contents.trim()) return contents;
  if (!Array.isArray(annotation.textContent)) return "";
  return annotation.textContent
    .map((line) => safeText(line))
    .join("\n")
    .slice(0, MAX_NATIVE_PDF_FREE_TEXT_LENGTH);
}

function finiteColorChannel(value: unknown): number | null {
  const channel = Number(value);
  return Number.isFinite(channel) && channel >= 0 && channel <= 255
    ? Math.round(channel)
    : null;
}

function optionalColorToHex(value: unknown): string | null {
  if (
    !value ||
    typeof value !== "object" ||
    !("length" in value) ||
    Number((value as { length?: unknown }).length) < 3
  ) {
    return null;
  }
  const channels = value as ArrayLike<unknown>;
  const red = finiteColorChannel(channels[0]);
  const green = finiteColorChannel(channels[1]);
  const blue = finiteColorChannel(channels[2]);
  if (red == null || green == null || blue == null) return null;
  return `#${[red, green, blue]
    .map((channel) => channel.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase()}`;
}

function colorToHex(value: unknown): string {
  return optionalColorToHex(value) ?? "#000000";
}

function fontNameText(value: unknown): string {
  if (typeof value === "string") return value;
  if (value && typeof value === "object") {
    const name = (value as { name?: unknown }).name;
    if (typeof name === "string") return name;
  }
  return "Helvetica";
}

function portableFont(
  value: unknown,
  baseFont?: string,
): {
  fontFamily: string;
  sourceFontName: string;
  bold: boolean;
  italic: boolean;
} {
  const sourceFontName = fontNameText(value).slice(0, 200) || "Helvetica";
  // The /DA names a resource, which is only a key: "FZ2572" can be Courier.
  // The font it points to decides the family when the bytes tell it
  
  const normalized = (baseFont?.trim() || sourceFontName)
    .replace(/^[A-Z]{6}\+/u, "")
    .toLowerCase();
  const mono = /cour|mono|code|typewriter/u.test(normalized);
  const serif =
    !mono &&
    /times|tiro|tibo|tiit|tibi|serif|georgia|garamond/u.test(normalized);
  return {
    fontFamily: mono ? "monospace" : serif ? "serif" : "sans-serif",
    sourceFontName,
    bold: /bold|hebo|tibo|tibi|cobo|cobold|black|semibold/u.test(normalized),
    italic: /italic|oblique|heob|tibi|tiit|coob/u.test(normalized),
  };
}

function normalizedRect(
  value: unknown,
): readonly [number, number, number, number] | null {
  if (!Array.isArray(value) || value.length < 4) return null;
  const values = value.slice(0, 4).map(Number);
  if (values.some((coordinate) => !Number.isFinite(coordinate))) return null;
  const x1 = Math.min(values[0], values[2]);
  const y1 = Math.min(values[1], values[3]);
  const x2 = Math.max(values[0], values[2]);
  const y2 = Math.max(values[1], values[3]);
  return x2 > x1 && y2 > y1 ? [x1, y1, x2, y2] : null;
}

function viewportFractions(
  rect: readonly number[],
  viewport: NativePdfFreeTextViewport,
): { x: number; y: number; width: number; height: number } | null {
  let converted: readonly [number, number, number, number];
  try {
    converted = viewport.convertToViewportRectangle(rect);
  } catch {
    return null;
  }
  const x1 = Math.min(converted[0], converted[2]) / viewport.width;
  const y1 = Math.min(converted[1], converted[3]) / viewport.height;
  const x2 = Math.max(converted[0], converted[2]) / viewport.width;
  const y2 = Math.max(converted[1], converted[3]) / viewport.height;
  if (
    ![x1, y1, x2, y2].every(Number.isFinite) ||
    x1 < -0.000_001 ||
    y1 < -0.000_001 ||
    x2 > 1.000_001 ||
    y2 > 1.000_001 ||
    x2 - x1 < 0.000_001 ||
    y2 - y1 < 0.000_001
  ) {
    return null;
  }
  return {
    x: Math.max(0, x1),
    y: Math.max(0, y1),
    width: Math.min(1, x2) - Math.max(0, x1),
    height: Math.min(1, y2) - Math.max(0, y1),
  };
}

export function sanitizeNativePdfFreeTextAnnotationSource(
  value: unknown,
): NativePdfFreeTextAnnotationSource | undefined {
  if (!value || typeof value !== "object") return undefined;
  const source = value as Record<string, unknown>;
  const objectRef = source.objectRef ?? source.object_ref;
  if (
    source.kind !== NATIVE_PDF_FREE_TEXT_SOURCE_KIND ||
    !isPdfObjectRef(objectRef) ||
    source.subtype !== "FreeText"
  ) {
    return undefined;
  }
  return {
    kind: NATIVE_PDF_FREE_TEXT_SOURCE_KIND,
    objectRef,
    subtype: "FreeText",
  };
}

export function isNativePdfFreeTextRecord(value: {
  nativeSource?: unknown;
}): value is NativePdfFreeTextRecord {
  return Boolean(sanitizeNativePdfFreeTextAnnotationSource(value.nativeSource));
}






export function pdfTextBoxLaneWrites(edit: {
  created?: boolean;
  nativeSource?: unknown;
}): boolean {
  return Boolean(edit.created) || isNativePdfFreeTextRecord(edit);
}

/**
 * Normalize standard typewriter-style source text boxes into the same fraction
 * model used by the MRNote inline text overlay.
 *
 * pdf.js exposes enough information to reject locked, rotated, callout, or
 * malformed boxes before they become interactive.  The backend repeats the
 * source-dictionary checks (including rich `/RC` content) before mutation.
 */
export function normalizeNativePdfFreeTextAnnotations(
  rawAnnotations: unknown,
  pageNumber: number,
  viewport: NativePdfFreeTextViewport,
  sourceFacts?: ReadonlyMap<string, NativePdfFreeTextSourceFacts> | null,
): NativePdfFreeTextImportResult {
  const textBoxes: NativePdfFreeTextAnnotation[] = [];
  const losses: NativePdfFreeTextImportLoss[] = [];
  if (
    !Array.isArray(rawAnnotations) ||
    !Number.isInteger(pageNumber) ||
    pageNumber < 1 ||
    !Number.isFinite(viewport.width) ||
    !Number.isFinite(viewport.height) ||
    !Number.isFinite(viewport.scale) ||
    viewport.width <= 0 ||
    viewport.height <= 0 ||
    viewport.scale <= 0
  ) {
    return { textBoxes, losses: [{ reason: "invalid_page_geometry" }] };
  }

  const referenceCounts = new Map<string, number>();
  for (const raw of rawAnnotations) {
    if (!raw || typeof raw !== "object") continue;
    const annotation = raw as RawNativePdfFreeTextAnnotation;
    if (annotation.subtype === "FreeText" && isPdfObjectRef(annotation.id)) {
      referenceCounts.set(
        annotation.id,
        (referenceCounts.get(annotation.id) ?? 0) + 1,
      );
    }
  }

  const reportedDuplicateRefs = new Set<string>();
  let sourceCount = 0;
  for (const raw of rawAnnotations) {
    if (!raw || typeof raw !== "object") continue;
    const annotation = raw as RawNativePdfFreeTextAnnotation;
    if (annotation.subtype !== "FreeText") continue;
    sourceCount += 1;
    const objectRef = isPdfObjectRef(annotation.id) ? annotation.id : undefined;
    if (sourceCount > MAX_NATIVE_PDF_FREE_TEXT_ANNOTATIONS_PER_PAGE) {
      losses.push({ objectRef, reason: "annotation_limit" });
      continue;
    }
    if (!objectRef) {
      losses.push({ reason: "invalid_object_ref" });
      continue;
    }
    if ((referenceCounts.get(objectRef) ?? 0) > 1) {
      if (!reportedDuplicateRefs.has(objectRef)) {
        losses.push({ objectRef, reason: "duplicate_object_ref" });
        reportedDuplicateRefs.add(objectRef);
      }
      continue;
    }

    const flags = Number(annotation.annotationFlags);
    // A Hidden (or NoView) source box is not drawn, so editing must not draw
    
    if (
      nativePdfAnnotationIsInvisible(
        readNativePdfAnnotationFlags(annotation.annotationFlags),
      )
    ) {
      losses.push({ objectRef, reason: "hidden_annotation" });
      continue;
    }
    if (
      Number.isInteger(flags) &&
      (flags & (LOCKED_ANNOTATION_FLAG | LOCKED_CONTENTS_ANNOTATION_FLAG)) !== 0
    ) {
      losses.push({ objectRef, reason: "locked_annotation" });
      continue;
    }
    if (sourceFacts?.get(objectRef)?.sourceTextOwner) {
      losses.push({ objectRef, reason: "source_text_owner" });
      continue;
    }
    if (sourceFacts?.get(objectRef)?.managedObject || sourceFacts?.get(objectRef)?.managedTextBox) {
      losses.push({ objectRef, reason: "managed_object" });
      continue;
    }
    // /IT FreeText is the default a plain text box has when /IT is left out
    
    // Callouts keep their leader line and stay read-only.
    if (
      annotation.it !== undefined &&
      annotation.it !== null &&
      annotation.it !== "FreeText" &&
      annotation.it !== "FreeTextTypeWriter"
    ) {
      losses.push({ objectRef, reason: "unsupported_intent" });
      continue;
    }
    const annotationRotation = Number(annotation.rotation ?? 0);
    if (
      !Number.isFinite(annotationRotation) ||
      !Number.isInteger(annotationRotation) ||
      annotationRotation < 0 ||
      annotationRotation >= 360 ||
      annotationRotation % 90 !== 0
    ) {
      losses.push({ objectRef, reason: "unsupported_rotation" });
      continue;
    }
    if (annotation.richText !== undefined && annotation.richText !== null) {
      losses.push({ objectRef, reason: "unsupported_rich_text" });
      continue;
    }
    // A plain solid rectangle is editable; rounded corners, dashes, and the
    // beveled/inset/underline styles change the drawn shape, so they stay
    // read-only rather than being approximated by the overlay.
    let borderWidth = 0;
    if (annotation.borderStyle && typeof annotation.borderStyle === "object") {
      const border = annotation.borderStyle as {
        width?: unknown;
        style?: unknown;
        dashArray?: unknown;
        horizontalCornerRadius?: unknown;
        verticalCornerRadius?: unknown;
      };
      const width = Number(border.width ?? 0);
      const radii = [
        Number(border.horizontalCornerRadius ?? 0),
        Number(border.verticalCornerRadius ?? 0),
      ];
      // `dashArray` is not consulted: pdf.js hands every annotation the same
      // `[3]` default, and a dash pattern only draws when the style is DASHED —
      // which the style check below already rejects.
      if (
        !Number.isFinite(width) ||
        width < 0 ||
        width > MAX_NATIVE_PDF_FREE_TEXT_BORDER_WIDTH ||
        !radii.every((radius) => Number.isFinite(radius) && radius === 0) ||
        (border.style !== undefined &&
          border.style !== null &&
          border.style !== SOLID_ANNOTATION_BORDER_STYLE)
      ) {
        losses.push({ objectRef, reason: "unsupported_border_style" });
        continue;
      }
      borderWidth = width;
    }

    const rect = normalizedRect(annotation.rect);
    const geometry = rect ? viewportFractions(rect, viewport) : null;
    if (!geometry) {
      losses.push({ objectRef, reason: "invalid_text_box_geometry" });
      continue;
    }
    // An empty box is a box: the author left it for someone to fill in, so
    
    const text = annotationText(annotation);
    const appearance =
      annotation.defaultAppearanceData &&
      typeof annotation.defaultAppearanceData === "object"
        ? (annotation.defaultAppearanceData as RawDefaultAppearanceData)
        : {};
    const sourceFontSize = Number(appearance.fontSize);
    const fontSize =
      Number.isFinite(sourceFontSize) && sourceFontSize > 0
        ? (sourceFontSize * viewport.scale) / viewport.height
        : Math.min(0.03, Math.max(0.012, geometry.height * 0.42));
    if (!Number.isFinite(fontSize) || fontSize <= 0 || fontSize > 1) {
      losses.push({ objectRef, reason: "unsupported_typography" });
      continue;
    }
    const facts = sourceFacts?.get(objectRef);
    const font = portableFont(appearance.fontName, facts?.baseFont);
    // The inner text area the author set (/RD), as edge insets of the box.
    let sourceInsets: NativePdfFreeTextAnnotation["sourceInsets"];
    const differences = facts?.rectangleDifferences;
    if (rect && differences && differences.some((value) => value > 0)) {
      const [left, top, right, bottom] = differences;
      const inner = normalizedRect([
        rect[0] + left,
        rect[1] + bottom,
        rect[2] - right,
        rect[3] - top,
      ]);
      const innerGeometry = inner ? viewportFractions(inner, viewport) : null;
      if (innerGeometry) {
        sourceInsets = {
          left: Math.max(0, innerGeometry.x - geometry.x),
          top: Math.max(0, innerGeometry.y - geometry.y),
          right: Math.max(
            0,
            geometry.x + geometry.width - (innerGeometry.x + innerGeometry.width),
          ),
          bottom: Math.max(
            0,
            geometry.y + geometry.height - (innerGeometry.y + innerGeometry.height),
          ),
        };
      }
    }
    // The source's own opacity: from its appearance when the bytes could be
    
    const opacityValue = facts ? facts.appearanceOpacity : Number(annotation.opacity);
    const sourceOpacity =
      opacityValue !== undefined &&
      Number.isFinite(opacityValue) &&
      opacityValue >= 0 &&
      opacityValue < 1
        ? opacityValue
        : undefined;
    // `/C` on a FreeText is the box's own fill, which the source appearance
    // paints behind the text. The overlay replaces that appearance, so it has
    // to paint the same fill or an author's coloured box would read as
    // transparent on screen while the saved PDF still fills it.
    const boxFill = optionalColorToHex(annotation.color);
    const firstLine =
      (Array.isArray(annotation.textContent)
        ? annotation.textContent.find(
            (line): line is string => typeof line === "string",
          )
        : undefined) ?? text.split(/\r\n?|\n/u)[0];
    const lineStart = Array.isArray(annotation.textPosition)
      ? Number(annotation.textPosition[0])
      : Number.NaN;
    const sourceAlign =
      rect &&
      !annotationRotation &&
      viewport.measureTextWidth &&
      firstLine &&
      Number.isFinite(lineStart) &&
      Number.isFinite(sourceFontSize) &&
      sourceFontSize > 0
        ? freeTextAlignmentFromAppearance({
            // Placement is judged inside the author's text area (/RD).
            boxWidth:
              rect[2] -
              rect[0] -
              (differences ? Math.max(0, differences[0]) + Math.max(0, differences[2]) : 0),
            lineStart: lineStart - (differences ? Math.max(0, differences[0]) : 0),
            lineWidth: viewport.measureTextWidth(firstLine, font, sourceFontSize),
            fontSize: sourceFontSize,
          })
        : null;
    const id = `pdf-native-freetext:${pageNumber}:${objectRef}`;
    textBoxes.push({
      id,
      page: pageNumber - 1,
      ...geometry,
      ...(boxFill ? { backgroundColor: boxFill } : {}),
      text,
      fontSize,
      color: colorToHex(appearance.fontColor),
      // Layout only: pdf.js never reports `/Q`, so the overlay lays the text
      // out as the source appearance places it (left when that cannot be
      // told), and the saved object keeps the author's quadding until the
      // user picks an alignment.
      align: sourceAlign ?? "left",
      alignFromSource: true,
      ...(borderWidth > 0
        ? {
            sourceBorderWidth: (borderWidth * viewport.scale) / viewport.height,
          }
        : {}),
      ...(sourceInsets ? { sourceInsets } : {}),
      ...(sourceOpacity !== undefined ? { sourceOpacity } : {}),
      ...font,
      underline: false,
      // /Rotate turns counterclockwise (as pdf.js and Acrobat draw it); the
      
      ...(annotationRotation ? { angle: (360 - annotationRotation) % 360 } : {}),
      created: true,
      originX: geometry.x,
      originY: geometry.y,
      originWidth: geometry.width,
      originHeight: geometry.height,
      nativeSource: {
        kind: NATIVE_PDF_FREE_TEXT_SOURCE_KIND,
        objectRef,
        subtype: "FreeText",
      },
    });
  }
  return { textBoxes, losses };
}

/**
 * Merge one freshly inspected page. Pending edits and deletion tombstones are
 * protected by immutable source ref; untouched discovery records are refreshed.
 */
export function mergeNativePdfFreeTextAnnotations<
  T extends TextBoxEditLike & {
    id?: string;
    page?: number;
    nativeSource?: unknown;
  },
>(
  current: readonly T[],
  incoming: readonly NativePdfFreeTextAnnotation[],
  pageNumber: number,
): T[] {
  const pageIndex = pageNumber - 1;
  const ordinary = current.filter(
    (record) =>
      Number(record.page) !== pageIndex || !isNativePdfFreeTextRecord(record),
  );
  const currentSourcePage = current.filter(
    (record) =>
      Number(record.page) === pageIndex && isNativePdfFreeTextRecord(record),
  );
  const protectedByRef = new Map(
    currentSourcePage
      .filter(
        (record) =>
          record.deleted === true ||
          record.sourceAction === "content_style_geometry",
      )
      .map((record) => [
        sanitizeNativePdfFreeTextAnnotationSource(record.nativeSource)!
          .objectRef,
        record,
      ]),
  );
  const nextPage: T[] = [...protectedByRef.values()];
  for (const discovered of incoming) {
    if (protectedByRef.has(discovered.nativeSource.objectRef)) continue;
    nextPage.push(discovered as T);
  }

  const next = [...ordinary, ...nextPage];
  if (
    next.length === current.length &&
    next.every((record, index) => record === current[index])
  ) {
    return current as T[];
  }
  return next;
}

/** One source /FreeText (a text box or a callout) as the annotation list
 * shows it. */
export interface NativePdfFreeTextListRow {
  readonly id: string;
  /** One-based page number. */
  readonly page: number;
  readonly text: string;
  readonly color?: string;
  readonly author?: string;
  readonly nativeSource: Readonly<{
    objectRef: string;
    subtype: "FreeText";
    flags: NativePdfAnnotationFlags;
  }>;
}










export function nativePdfFreeTextListRows(
  annotations: readonly unknown[],
  pageNumber: number,
  editableIds: ReadonlySet<string>,
  /** Tables and charts made in MRAI: no comment of the page. */
  managedObjectIds: ReadonlySet<string> = new Set(),
): NativePdfFreeTextListRow[] {
  const rows: NativePdfFreeTextListRow[] = [];
  for (const raw of annotations) {
    if (!raw || typeof raw !== "object") continue;
    const annotation = raw as {
      subtype?: unknown;
      id?: unknown;
      annotationFlags?: unknown;
      contentsObj?: unknown;
      titleObj?: unknown;
      color?: unknown;
    };
    if (annotation.subtype !== "FreeText" || typeof annotation.id !== "string" || !annotation.id) continue;
    if (managedObjectIds.has(annotation.id)) continue;
    const flags = readNativePdfAnnotationFlags(annotation.annotationFlags);
    if (nativePdfAnnotationIsInvisible(flags)) continue;
    const color = annotation.color && typeof annotation.color === "object" &&
      typeof (annotation.color as ArrayLike<number>).length === "number" &&
      (annotation.color as ArrayLike<number>).length === 3
      ? `#${Array.from(annotation.color as ArrayLike<number>, (value) =>
          Math.max(0, Math.min(255, Math.round(Number(value) || 0))).toString(16).padStart(2, "0")).join("")}`
      : undefined;
    const author = sanitizePdfText(annotation.titleObj, 200).trim();
    rows.push(Object.freeze({
      id: `pdf-native-freetext:${pageNumber}:${annotation.id}`,
      page: pageNumber,
      text: sanitizePdfText(annotation.contentsObj, MAX_NATIVE_PDF_FREE_TEXT_LENGTH).trim(),
      ...(color ? { color } : {}),
      ...(author ? { author } : {}),
      nativeSource: Object.freeze({
        objectRef: annotation.id,
        subtype: "FreeText" as const,
        flags: editableIds.has(annotation.id) ? flags : Object.freeze({ ...flags, readOnly: true }),
      }),
    }));
    if (rows.length >= MAX_NATIVE_PDF_FREE_TEXT_ANNOTATIONS_PER_PAGE) break;
  }
  return rows;
}
