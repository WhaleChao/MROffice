/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Adapter: convert persisted inline PDF text-box edits (the WPS-style editor's
 * `PdfTextBoxEdit`, stored as page FRACTIONS with a CSS font stack) into the
 * {@link FlattenTextBox} shape the flatten planner consumes (page PERCENT + a
 * resolved serif/sans/mono class). Pure + framework-free so the mapping is
 * unit-testable without a real PDF or DOM.
 */

import type { PdfTextBoxTypographyFields } from "@mrnote/core/web/lib/pdf/text-box-typography.ts";
import type { FlattenPageSize, FlattenTextBox } from "@mrnote/core/web/lib/pdf/flatten.ts";
import { tryClassifyFontName } from "./font-resolve";
import { DEFAULT_NEW_BOX_FONT_SIZE } from "@mrnote/core/web/lib/pdf/text-box-create.ts";
import { expandRectForGlyphCover, turnedRectBounds } from "@mrnote/core/web/lib/pdf/text-box-geometry.ts";
import { pdfSourceTextBoxInkCoverRect } from "./source-text-box";
import { effectiveTextBoxKind, latexForTextBoxKind } from "@mrnote/core/web/lib/pdf/text-box-kind.ts";
import { alignNativeTextBoxEditToSource, sanitizePdfTextBoxEditGeometry } from "@mrnote/core/web/lib/pdf/text-box-rect.ts";
import {
  normalizePdfTextStyleRuns,
  type PdfTextStyleRun,
} from "@mrnote/core/web/lib/pdf/text-style-runs.ts";
import {
  normalizePdfTextLineLayout,
  type PdfTextLineLayout,
} from "@mrnote/core/web/lib/pdf/text-line-layout.ts";
import type { NativePdfFreeTextAnnotationSource } from "./native-free-text-annotations";
import {
  hasPdfTextBoxTypography,
  readPdfTextBoxTypography,
  scalePdfTextBoxTypography,
} from "@mrnote/core/web/lib/pdf/text-box-typography.ts";

/**
 * Structural shape of one persisted edit (mirrors `PdfTextBoxEdit` without
 * importing the component, keeping this a leaf lib). Geometry is page-relative
 * FRACTIONS [0..1]; `page` is ZERO-based (the editor's page index); `fontSize`
 * is a fraction of page height; `fontFamily` is the original CSS font stack.
 */
export interface TextBoxEditLike extends PdfTextBoxTypographyFields {
  sourceText?: import("./source-text-box").PdfSourceTextBoxBinding;
  id?: string;
  page?: number;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  /** Clockwise object rotation around the text-box centre, in degrees. */
  angle?: number;
  /** Upright CJK glyphs in vertical columns; never approximate as horizontal. */
  writingMode?: "vertical-rl";
  text?: string;
  kind?: "text" | "formula";
  latex?: string;
  fontSize?: number;
  color?: string;
  align?: "left" | "center" | "right";
  fontFamily?: string;
  sourceFontName?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  /** A line through the middle of every rendered line (Word's Strikethrough). */
  strikethrough?: boolean;
  /** Justify every line but a paragraph's last to the box width (Word's Justify). */
  justify?: boolean;
  /** Mixed-format paragraph typography; concatenates exactly to `text`. */
  runs?: PdfTextStyleRun[];
  /**
   * The Office font family (lib/office-fonts catalog name) the whole box is
   * set in; the text-box writer embeds its served substitute (text-box-runs.ts).
   * A run naming a face of its own keeps it.
   */
  officeFamily?: string;
  /** Source line indentation and vertical placement within the paragraph box. */
  lineLayout?: PdfTextLineLayout[];
  /** True for text boxes created from blank space rather than extracted PDF text. */
  created?: boolean;
  /**
   * True when the user DELETED an extracted line as an object. The edit is
   * kept (with empty text) so the export still bakes the origin-pinned mask
   * over the original glyphs; any residual text/latex on the record must NOT
   * be drawn.
   */
  deleted?: boolean;
  /**
   * The ORIGINAL extracted glyph region (page FRACTIONS [0..1]) the cover/mask
   * must stay pinned to, independent of where the box was dragged. Carried
   * through to the flatten plan so the export masks the source location (not the
   * moved box). Additive — edits saved before origin tracking omit it and the
   * mask falls back to the box rect.
   */
  originX?: number;
  originY?: number;
  originWidth?: number;
  originHeight?: number;
  /** Sampled page background for the glyph cover/export mask. */
  backgroundColor?: string;
  /**
   * Exact source identity for a third-party `/FreeText` annotation. Source
   * records are patched through the native annotation bridge and must never be
   * flattened as a second visual text box over the original object.
   */
  nativeSource?: NativePdfFreeTextAnnotationSource;
  sourceAction?: "content_style_geometry";
  /**
   * Set on an imported `/FreeText` whose alignment the user has never picked:
   * `align` is the overlay's fallback, and the source `/Q` stays authoritative
   * in the saved PDF.
   */
  alignFromSource?: boolean;
  /** Solid imported-`/FreeText` border, as a fraction of page height. */
  sourceBorderWidth?: number;
  nativeAppearanceImage?: string;
}

export interface FlattenTextBoxEditOptions {
  pageSizes?: ReadonlyArray<FlattenPageSize>;
  /** Fresh extraction metadata used only to recover live pdf.js font aliases. */
  sourceBoxes?: ReadonlyArray<TextBoxEditLike>;
}

function runsWithLiveDisplayFonts(
  edit: TextBoxEditLike,
  runs: readonly PdfTextStyleRun[],
  sourceBoxes: ReadonlyArray<TextBoxEditLike> | undefined,
): PdfTextStyleRun[] {
  const source = sourceBoxes?.find(
    (box) => box.id === edit.id && Number(box.page) === Number(edit.page),
  );
  const sourceRuns = normalizePdfTextStyleRuns(source?.runs, {
    expectedText: typeof source?.text === "string" ? source.text : undefined,
    allowDisplayFont: true,
  });
  if (!sourceRuns) return [...runs];
  return runs.map((run) => {
    const matching = sourceRuns.find((candidate) => {
      if (run.sourceFontName && candidate.sourceFontName) {
        return run.sourceFontName === candidate.sourceFontName;
      }
      return Boolean(run.fontFamily && candidate.fontFamily === run.fontFamily);
    });
    return matching?.displayFontFamily
      ? { ...run, displayFontFamily: matching.displayFontFamily }
      : { ...run };
  });
}

/**
 * Resolve a CSS font-family stack to the base-14 class the bake can draw. Order
 * matters: "sans-serif" contains the substring "serif", so sans is matched
 * BEFORE serif. Unknown / empty -> "sans" (the safe default face).
 */
export function fontClassFromFamily(
  family?: string,
): "serif" | "sans" | "mono" {
  return tryClassifyFontName(family) ?? "sans";
}

function tryFontClassFromFamily(
  family?: string,
): "serif" | "sans" | "mono" | null {
  return tryClassifyFontName(family);
}

/**
 * Map persisted edits to flatten text-box inputs. Edits with non-finite or
 * zero-area geometry are skipped (never throw). An empty input yields `[]`, so a
 * PDF with no text edits adds nothing to the flatten plan (byte-for-byte).
 */
export function flattenTextBoxesFromEdits(
  edits: ReadonlyArray<TextBoxEditLike> | null | undefined,
  options: FlattenTextBoxEditOptions = {},
): FlattenTextBox[] {
  if (!Array.isArray(edits) || edits.length === 0) return [];
  const out: FlattenTextBox[] = [];
  for (const storedEdit of edits) {
    const edit = storedEdit && options.sourceBoxes
      ? alignNativeTextBoxEditToSource(storedEdit, options.sourceBoxes)
      : storedEdit;
    if (!edit || typeof edit !== "object") continue;
    if (edit.nativeSource?.kind === "pdf_free_text_annotation") continue;
    const geometry = sanitizePdfTextBoxEditGeometry(edit);
    if (!geometry) continue;
    const { x, y, width, height } = geometry;
    // A DELETED extracted line bakes as mask-only: force the text blank (and
    // drop latex below) so a record that still carries content draws nothing,
    // while its origin-pinned mask op still hides the original glyphs.
    const deleted = edit.deleted === true;
    // A deleted CREATED box has no underlying glyphs to mask — nothing to
    // bake at all (its record is normally removed outright by the host).
    if (deleted && edit.created) continue;
    const text = !deleted && typeof edit.text === "string" ? edit.text : "";
    const pageIndex = Number(edit.page);
    const pageNumber =
      (Number.isFinite(pageIndex) ? Math.max(0, Math.trunc(pageIndex)) : 0) + 1;
    const pageSize = options.pageSizes?.[pageNumber - 1];
    const fontSize = Number(edit.fontSize);
    // The original glyph region to mask (page fractions → percent). Only carried
    // when all four fields are finite & the area is positive; otherwise the
    // flatten mask falls back to the box rect (legacy edits).
    const hasOrigin =
      geometry.originX !== undefined &&
      geometry.originY !== undefined &&
      geometry.originWidth !== undefined &&
      geometry.originHeight !== undefined;
    const fontSizePercent =
      (Number.isFinite(fontSize) ? fontSize : DEFAULT_NEW_BOX_FONT_SIZE) * 100;
    // An OCR line's box moved off the scanned words it replaced: its fill is
    // a stationary cover at them, the box itself clear (PDF 0107), as the
    // writer paints it and the live view draws it.
    const inkCover = !deleted && edit.sourceText ? pdfSourceTextBoxInkCoverRect(edit) : null;
    const maskRect = inkCover ? turnedRectBounds(inkCover, edit.angle, pageSize) :
      (hasOrigin && !edit.created
        ? expandRectForGlyphCover(
            {
              x: geometry.originX as number,
              y: geometry.originY as number,
              width: geometry.originWidth as number,
              height: geometry.originHeight as number,
            },
            fontSizePercent / 100,
            pageSize,
          )
        : null);
    const fontClass =
      tryFontClassFromFamily(edit.sourceFontName) ??
      tryFontClassFromFamily(edit.fontFamily) ??
      "sans";
    const rawLatex =
      !deleted && typeof edit.latex === "string" ? edit.latex : undefined;
    const kind = effectiveTextBoxKind(edit);
    const latex = deleted ? undefined : latexForTextBoxKind(kind, rawLatex);
    if (edit.created && text.trim() === "" && !latex) {
      continue;
    }
    const normalizedRuns = normalizePdfTextStyleRuns(edit.runs, {
      expectedText: text,
    });
    const runs = normalizedRuns
      ? runsWithLiveDisplayFonts(edit, normalizedRuns, options.sourceBoxes)
      : undefined;
    const lineLayout = normalizePdfTextLineLayout(edit.lineLayout, text);
    out.push({
      // The editor stores a 0-based page index; flatten is 1-based.
      page: pageNumber,
      x: x * 100,
      y: y * 100,
      width: width * 100,
      height: height * 100,
      ...(maskRect
        ? {
            maskX: maskRect.x * 100,
            maskY: maskRect.y * 100,
            maskWidth: maskRect.width * 100,
            maskHeight: maskRect.height * 100,
          }
        : {}),
      text,
      ...(kind ? { kind } : {}),
      // A box the user made: its formula is drawn whenever its LaTeX parses.
      ...(edit.created === true ? { created: true } : {}),
      ...(latex !== undefined ? { latex } : {}),
      fontSizePercent,
      color: typeof edit.color === "string" ? edit.color : undefined,
      ...(typeof edit.backgroundColor === "string"
        ? { backgroundColor: edit.backgroundColor }
        : {}),
      align:
        edit.align === "center" || edit.align === "right" ? edit.align : "left",
      font: fontClass,
      // Preserve the raw CSS family so `standardFontFor` can try an exact
      // base-14 name match (e.g. "Helvetica" → exact Helvetica, not just
      // "sans → Helvetica via class"). Non-WinAnsi text keeps this source
      // identity but renders through the bundled unicode/raster path.
      ...(typeof edit.fontFamily === "string" && edit.fontFamily
        ? { fontFamily: edit.fontFamily }
        : {}),
      ...(typeof edit.sourceFontName === "string" && edit.sourceFontName
        ? { sourceFontName: edit.sourceFontName }
        : {}),
      bold: Boolean(edit.bold),
      italic: Boolean(edit.italic),
      underline: Boolean(edit.underline),
      ...(edit.strikethrough ? { strikethrough: true } : {}),
      ...(edit.justify ? { justify: true } : {}),
      ...(() => {
        const size = fontSizePercent / 100;
        const typography = readPdfTextBoxTypography(edit, size);
        // Carried in em so the planner turns it into points with the font.
        return hasPdfTextBoxTypography(typography)
          ? { typography: scalePdfTextBoxTypography(typography, 1 / size) }
          : {};
      })(),
      ...(runs
        ? {
            runs: runs.map((run) => ({
              text: run.text,
              fontSizePercent:
                (typeof run.fontSize === "number"
                  ? run.fontSize
                  : Number.isFinite(fontSize)
                    ? fontSize
                    : DEFAULT_NEW_BOX_FONT_SIZE) * 100,
              color: typeof run.color === "string" ? run.color : edit.color,
              font:
                tryFontClassFromFamily(run.sourceFontName) ??
                tryFontClassFromFamily(run.fontFamily) ??
                fontClass,
              bold: run.bold ?? Boolean(edit.bold),
              italic: run.italic ?? Boolean(edit.italic),
              ...(run.fontFamily ? { fontFamily: run.fontFamily } : {}),
              ...(run.displayFontFamily
                ? { displayFontFamily: run.displayFontFamily }
                : {}),
              ...(run.sourceFontName
                ? { sourceFontName: run.sourceFontName }
                : {}),
            })),
          }
        : {}),
      ...(lineLayout ? { lineLayout } : {}),
      // A turned box exports at the angle it was extracted at; the planner and
      // `drawTextBox` already know what to do with one.
      ...(Number.isFinite(edit.angle) && edit.angle
        ? { angle: edit.angle }
        : {}),
      ...(edit.writingMode === "vertical-rl"
        ? { writingMode: "vertical-rl" as const }
        : {}),
    });
  }
  return out;
}
