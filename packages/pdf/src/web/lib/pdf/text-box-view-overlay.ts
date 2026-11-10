/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  CJK_SANS_STACK,
  CJK_SERIF_STACK,
  FONT_STACK,
  fontFamilyForPdfText,
  resolvePdfFontFamily,
  tryClassifyFontName,
  normalizePdfSourceFontName,
  withEastAsianFallback,
} from "./font-resolve";
import { ensureOfficeFontFaces } from "@mrnote/core/web/lib/office-fonts/substitutes.ts";
import { isSafeCssColor } from "@mrnote/core/web/lib/pdf/safe-color.ts";
import { pdfSourceTextBoxCoversInk, pdfSourceTextBoxInkCoverRect } from "./source-text-box";
import {
  textBoxOverlayFill,
} from "./text-box-background";
import type { TextBoxEditLike } from "./text-box-flatten";
import { hasPdfTextBoxTypography, readPdfTextBoxTypography } from "@mrnote/core/web/lib/pdf/text-box-typography.ts";
import {
  expandRectForGlyphCover,
  type GlyphCoverOptions,
  turnedRectBounds,
} from "@mrnote/core/web/lib/pdf/text-box-geometry.ts";
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

export interface TextBoxViewRun {
  text: string;
  fontSize: number;
  color: string;
  fontFamily: string;
  fontWeight: "bold" | "normal";
  fontStyle: "italic" | "normal";
}

export interface TextBoxViewOverlay {
  id: string;
  /** 1-based PDF page number. */
  page: number;
  x: number;
  y: number;
  width: number;
  height: number;
  text: string;
  kind?: "text" | "formula";
  latex?: string;
  /** A box the user made: its formula is drawn whenever its LaTeX parses. */
  created?: true;
  fontSize: number;
  color: string;
  backgroundColor: string;
  fontFamily: string;
  fontWeight: "bold" | "normal";
  fontStyle: "italic" | "normal";
  /** CSS text-decoration-line: underline and/or line-through, or none. */
  textDecoration: string;
  textAlign: "left" | "center" | "right" | "justify";
  /** Word's remaining whole-box formats, lengths as page-height fractions. */
  typography?: import("@mrnote/core/web/lib/pdf/text-box-typography.ts").PdfTextBoxTypographyFields;
  /** Degrees the box is turned about its centre, as the editor frame and the
   * saved appearance turn it. Without it a turned box was drawn level from
   * the moment its editor closed until the page painted the save. */
  angle?: number;
  runs?: TextBoxViewRun[];
  lineLayout?: PdfTextLineLayout[];
  mask: {
    x: number;
    y: number;
    width: number;
    height: number;
  } | null;
  /** The mask's own fill when it is not the box's: an OCR line's box moved
   * off the scanned words leaves its fill at them and is clear itself
   * (PDF 0107, pdfSourceTextBoxInkCoverRect). */
  maskBackgroundColor?: string;
}

export interface TextBoxViewOverlayOptions {
  pageSizes?: ReadonlyArray<GlyphCoverOptions>;
  sourceBoxes?: Parameters<typeof alignNativeTextBoxEditToSource>[1];
}

function positiveFinite(value: unknown): number | null {
  const numeric = Number(value);
  return Number.isFinite(numeric) && numeric > 0 ? numeric : null;
}

function finite(value: unknown): number | null {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : null;
}

function resolveTextBoxFontFamily(edit: TextBoxEditLike): string {
  // A box set in an Office family (Brush Script MT, Kaushan Script...) is
  // drawn in the served face the editor set it in: the face is declared
  // here as well, since a page opened afresh draws its saved boxes before
  
  // title saved in Brush Script MT came back in a bold sans after a reload).
  if (edit.officeFamily) ensureOfficeFontFaces(edit.officeFamily);
  return withEastAsianFallback(resolveTextBoxLatinFontFamily(edit), String(edit.text || ""));
}

function resolveTextBoxLatinFontFamily(edit: TextBoxEditLike): string {
  const sourceResolved = resolvePdfFontFamily({
    fontName: edit.sourceFontName,
    styleFamily: edit.fontFamily,
  }).fontFamily;
  if (sourceResolved === CJK_SERIF_STACK || sourceResolved === CJK_SANS_STACK) {
    return (
      fontFamilyForPdfText(String(edit.text || ""), sourceResolved) ??
      sourceResolved
    );
  }
  const family = typeof edit.fontFamily === "string" ? edit.fontFamily : "";
  const trimmedFamily = family.trim();
  if (
    trimmedFamily &&
    !/\bg_d\d+_f\d+\b/i.test(trimmedFamily) &&
    !/[;{}]|\burl\s*\(|\bexpression\s*\(/i.test(trimmedFamily)
  ) {
    return (
      fontFamilyForPdfText(String(edit.text || ""), trimmedFamily) ??
      trimmedFamily
    ).slice(0, 400);
  }
  const fontClass =
    tryClassifyFontName(edit.sourceFontName) ??
    tryClassifyFontName(edit.fontFamily) ??
    "sans";
  return FONT_STACK[fontClass];
}

function viewRuns(
  edit: TextBoxEditLike,
  text: string,
  fallbackFontSize: number,
  fallbackColor: string,
): TextBoxViewRun[] | undefined {
  const runs = normalizePdfTextStyleRuns(edit.runs, { expectedText: text });
  if (!runs) return undefined;
  return runs.map((run: PdfTextStyleRun) => {
    if (run.officeFamily) ensureOfficeFontFaces(run.officeFamily);
    return {
    text: run.text,
    fontSize:
      typeof run.fontSize === "number" && run.fontSize > 0
        ? run.fontSize
        : fallbackFontSize,
    color: isSafeCssColor(run.color) ? run.color : fallbackColor,
    fontFamily: resolveTextBoxFontFamily({
      ...edit,
      text: run.text,
      fontFamily: run.fontFamily ?? edit.fontFamily,
      sourceFontName: run.sourceFontName ?? edit.sourceFontName,
    }),
    fontWeight: (run.bold ?? edit.bold) ? "bold" : "normal",
    fontStyle: (run.italic ?? edit.italic) ? "italic" : "normal",
    };
  });
}

export function textBoxViewOverlaysFromEdits(
  edits: ReadonlyArray<TextBoxEditLike> | null | undefined,
  options: TextBoxViewOverlayOptions = {},
): TextBoxViewOverlay[] {
  if (!Array.isArray(edits) || edits.length === 0) return [];

  const overlays: TextBoxViewOverlay[] = [];
  for (const storedEdit of edits) {
    const edit = storedEdit && options.sourceBoxes
      ? alignNativeTextBoxEditToSource(storedEdit, options.sourceBoxes)
      : storedEdit;
    if (!edit || typeof edit !== "object") continue;

    const geometry = sanitizePdfTextBoxEditGeometry(edit);
    if (!geometry) continue;
    const { x, y, width, height } = geometry;
    // A DELETED extracted line renders mask-only in the read view too: keep
    // the glyph cover, never its residual text/latex (mirrors the flatten
    // adapter). A deleted CREATED box has nothing to cover — skip entirely.
    const deleted = edit.deleted === true;
    if (deleted && edit.created) continue;
    const text = !deleted && typeof edit.text === "string" ? edit.text : "";

    const pageIndex = finite(edit.page);
    const page =
      (pageIndex === null ? 0 : Math.max(0, Math.trunc(pageIndex))) + 1;
    const pageSize = options.pageSizes?.[page - 1];
    const fontSize =
      positiveFinite(edit.fontSize) ?? Math.max(0.001, height * 0.8);
    // Opaque or it is not a cover: a translucent persisted value would show
    // the original glyphs beneath the replacement text (doubled text).
    // A box over source text is clear (the run is hidden in the bytes),
    // unless the run paints nothing and the box covers the scan under it in
    // the page's colour (PDF 0107): it is drawn on that fill, as written.
    const coversSourceInk = Boolean(edit.sourceText) && pdfSourceTextBoxCoversInk(edit.sourceText) &&
      isSafeCssColor(edit.backgroundColor);
    const fill = textBoxOverlayFill(edit);
    // Moved off the scanned words it replaced, the fill stays at them as a
    // stationary cover and the box is clear where it stands, as written.
    const inkCover = coversSourceInk && !deleted ? pdfSourceTextBoxInkCoverRect(edit) : null;
    const backgroundColor = inkCover ? "transparent" : fill;
    const color = isSafeCssColor(edit.color) ? edit.color : "#000000";
    const hasOrigin =
      geometry.originX !== undefined &&
      geometry.originY !== undefined &&
      geometry.originWidth !== undefined &&
      geometry.originHeight !== undefined;
    const baseMask = hasOrigin
      ? {
          x: geometry.originX as number,
          y: geometry.originY as number,
          width: geometry.originWidth as number,
          height: geometry.originHeight as number,
        }
      : { x, y, width, height };
    // The cover lies where the line lay: in the line's own axes, turned by
    // the box's angle (a line on a page turned a quarter carries that turn).
    const mask = inkCover ? turnedRectBounds(inkCover, edit.angle, pageSize) : (edit.created || edit.sourceText
      ? null
      : expandRectForGlyphCover(baseMask, fontSize, pageSize));
    const rawLatex =
      !deleted && typeof edit.latex === "string" ? edit.latex : undefined;
    const kind = effectiveTextBoxKind(edit);
    const latex = deleted ? undefined : latexForTextBoxKind(kind, rawLatex);
    if (edit.created && text.trim() === "" && !latex) {
      continue;
    }
    const runs = viewRuns(edit, text, fontSize, color);
    const lineLayout = normalizePdfTextLineLayout(edit.lineLayout, text);
    const exactSources = options.sourceBoxes?.filter(source =>
      source.id === edit.id && source.page === edit.page) ?? [];
    const liveSource = exactSources.length === 1 ? exactSources[0] : null;
    const sourceFontUnchanged = liveSource?.sourceFontName &&
      (!edit.fontFamily || normalizePdfSourceFontName(edit.fontFamily) ===
        normalizePdfSourceFontName(liveSource.sourceFontName)) &&
      normalizePdfSourceFontName(edit.sourceFontName ?? edit.fontFamily) ===
        normalizePdfSourceFontName(liveSource.sourceFontName);

    overlays.push({
      id: String(edit.id || `textbox-${page}-${overlays.length}`),
      page,
      x,
      y,
      width,
      height,
      text,
      kind,
      latex,
      ...(edit.created ? { created: true as const } : {}),
      fontSize,
      color,
      backgroundColor,
      fontFamily: sourceFontUnchanged && liveSource.displayFontFamily
        ? liveSource.displayFontFamily
        : resolveTextBoxFontFamily(edit),
      fontWeight: edit.bold ? "bold" : "normal",
      fontStyle: edit.italic ? "italic" : "normal",
      textDecoration:
        [
          ...(edit.underline ? ["underline"] : []),
          ...(edit.strikethrough ? ["line-through"] : []),
        ].join(" ") || "none",
      textAlign: edit.justify
        ? "justify"
        : edit.align === "center" || edit.align === "right"
          ? edit.align
          : "left",
      ...(runs ? { runs } : {}),
      ...(lineLayout ? { lineLayout } : {}),
      ...(() => {
        const typography = readPdfTextBoxTypography(edit, fontSize);
        return hasPdfTextBoxTypography(typography) ? { typography } : {};
      })(),
      ...(() => {
        const angle = finite(edit.angle);
        const turned = angle === null ? 0 : ((angle % 360) + 360) % 360;
        return turned > 0.001 && turned < 359.999 ? { angle: turned } : {};
      })(),
      mask,
      ...(inkCover ? { maskBackgroundColor: fill } : {}),
    });
  }

  return overlays;
}
