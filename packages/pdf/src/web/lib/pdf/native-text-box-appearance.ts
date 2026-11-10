/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  planFlattenOps,
  type FlattenPageSize,
  type RectOp,
  type TextOp,
} from "@mrnote/core/web/lib/pdf/flatten.ts";
import {
  isImportedNativeTextBoxEditEligible,
  isManagedNativeTextBoxEditEligible,
  sanitizeManagedNativeTextBoxImage,
} from "./native-text-box-roundtrip";
import {
  flattenTextBoxesFromEdits,
  type TextBoxEditLike,
} from "./text-box-flatten";
import { rasterBytesToDataUrl, renderPdfTextOpToImage } from "./text-op-raster";
import { pdfSourceTextBoxCoversInk, pdfSourceTextBoxInkCoverRect, pdfSourceTextBoxLineHeightRatio } from "./source-text-box";
import type { PdfPortableTextLayout } from "./portable-text-layout";
import type { PdfBrowserTextSpacing } from "./browser-text-layout";
import { readPdfTextBoxTypography } from "@mrnote/core/web/lib/pdf/text-box-typography.ts";
import { sanitizePdfTextBoxEditGeometry } from "@mrnote/core/web/lib/pdf/text-box-rect.ts";
import {
  pdfTextBoxDrawnRunsNeedRunWriter,
  pdfTextBoxRunCssFamily,
  resolvePdfTextBoxDrawnRuns,
} from "./text-box-runs";




export const nativeTextBoxAppearanceRefusals = new Map<string, string>();

/** Name why an appearance was not drawn (never throws). */
function refuseAppearance(id: string, reason: string, detail?: Record<string, unknown>): null {
  if (id) nativeTextBoxAppearanceRefusals.set(id, reason);
  if (process.env.NODE_ENV === "development") {
    try { console.warn("[pdf-text-box] appearance refused", { reason, ...detail }); } catch { /* never throws */ }
  }
  return null;
}

/**
 * Render the exact transparent PNG that becomes a managed `/FreeText` `/AP`.
 * The flatten planner remains the single geometry/style mapper: the edit is
 * temporarily addressed as page 0 against the supplied source page size, then
 * only its text op is rasterized. Unsupported complex edits return null and
 * remain on the explicit flattened-export path.
 */
export async function renderManagedNativeTextBoxAppearance(
  requestedEdit: TextBoxEditLike,
  pageSize: FlattenPageSize,
  displayFontFamily?: string,
  onPortableLayout?: (layout: PdfPortableTextLayout) => void,
  




  options: { allowImportedSource?: boolean;
    /** The frame grew down to hold its rows: the page fraction it needs. */
    onGrownHeight?: (height: number) => void } = {},
): Promise<string | null> {
  const refusalId = String(requestedEdit.id ?? "");
  if (refusalId) nativeTextBoxAppearanceRefusals.delete(refusalId);
  // The frame the page record keeps: content building clips a box to the
  // page (sanitizePdfTextBoxEditGeometry), and the writer checks the row
  // plan against that frame. A box widened past the page edge (a bolder
  
  
  const geometry = sanitizePdfTextBoxEditGeometry(requestedEdit);
  const edit: TextBoxEditLike = geometry ? { ...requestedEdit, ...geometry } : requestedEdit;
  if (
    !(options.allowImportedSource
      ? isImportedNativeTextBoxEditEligible(edit)
      : isManagedNativeTextBoxEditEligible(edit))
  ) {
    return refuseAppearance(refusalId, "ineligible", { imported: Boolean(options.allowImportedSource) });
  }
  if (
    !Number.isFinite(pageSize.width) ||
    !Number.isFinite(pageSize.height) ||
    pageSize.width <= 0 ||
    pageSize.height <= 0
  ) {
    return refuseAppearance(refusalId, "page-size");
  }
  // The flatten mapper skips imported FreeText (the file already paints
  // it); for its new appearance it is mapped like any other box.
  const imported = Boolean(edit.nativeSource);
  const textBoxes = flattenTextBoxesFromEdits(
    [{ ...edit, page: 0, ...(imported ? { nativeSource: undefined } : {}) }],
    { pageSizes: [pageSize] },
  );
  if (textBoxes.length !== 1) return refuseAppearance(refusalId, "flatten", { boxes: textBoxes.length });
  const plan = planFlattenOps({ textBoxes }, [pageSize]);
  const textOp = plan.ops.find(
    (operation): operation is TextOp => operation.kind === "text",
  );
  const backgroundOp = plan.ops.find(
    (operation): operation is RectOp => operation.kind === "rect",
  );
  if (!textOp || !backgroundOp) return refuseAppearance(refusalId, "plan", { text: Boolean(textOp), background: Boolean(backgroundOp) });
  const backgroundColor = `rgb(${Math.round(
    backgroundOp.color.r * 255,
  )}, ${Math.round(backgroundOp.color.g * 255)}, ${Math.round(
    backgroundOp.color.b * 255,
  )})`;
  const rotation = ((Number(pageSize.rotation ?? 0) % 360) + 360) % 360;
  const renderedWidth = rotation === 90 || rotation === 270 ? pageSize.height : pageSize.width;
  const renderedHeight = rotation === 90 || rotation === 270 ? pageSize.width : pageSize.height;
  // A box set in runs is drawn from its drawn runs (what the writer is
  // sent), sizes in the same points as the frame, and only on its row plan.
  const drawn = resolvePdfTextBoxDrawnRuns(edit);
  const runRows = drawn !== null && pdfTextBoxDrawnRunsNeedRunWriter(edit, drawn);
  if (runRows && !onPortableLayout) return refuseAppearance(refusalId, "runs-need-row-plan");
  // V5 rows are object-local, in the same point axes as the on-screen frame.
  // Page rotation belongs to placement, not to a second text reflow.
  const portableOp = onPortableLayout ? {...textOp,
    maxWidth:Number(edit.width)*renderedWidth,
    maxHeight:Number(edit.height)*renderedHeight,
    fontSize:Number(edit.fontSize)*renderedHeight,
    ...(runRows && drawn ? { runs: drawn.map((run) => ({
      text: run.text,
      fontSize: run.fontSize * renderedHeight,
      color: {
        r: Number.parseInt(run.color.slice(1, 3), 16) / 255,
        g: Number.parseInt(run.color.slice(3, 5), 16) / 255,
        b: Number.parseInt(run.color.slice(5, 7), 16) / 255,
      },
      font: run.font,
      bold: run.bold,
      italic: run.italic,
      underline: run.underline,
      strikethrough: run.strikethrough,
      ...(run.script ? { script: run.script } : {}),
      ...(run.officeFamily ? { officeFamily: run.officeFamily } : {}),
      fontFamily: pdfTextBoxRunCssFamily(run, run.text),
    })) } : {}),
  } : textOp;
  // An imported note keeps a clear background unless it states a fill. A
  // box over source text is clear, unless it covers the scan under a run
  // that paints nothing (PDF 0107): then it is drawn on its fill -- while its
  // frame holds the line it replaced. Moved off it, the fill stays at the
  // line (the writer paints it there, pdfSourceTextBoxInkCoverRect) and this
  // raster, the box's own part of the annotation, is clear.
  const clearBackground = edit.sourceText
    ? !(edit.backgroundColor && pdfSourceTextBoxCoversInk(edit.sourceText)) ||
      pdfSourceTextBoxInkCoverRect(edit) !== null
    : imported && !edit.backgroundColor;
  const rendered = await renderPdfTextOpToImage(portableOp, {
    backgroundColor: clearBackground ? undefined : backgroundColor,
    ...(edit.sourceText ? { browserLayout: true, displayFontFamily,
      browserLineHeightRatio: pdfSourceTextBoxLineHeightRatio(edit.sourceText),
      browserSpacing: pdfSourceTextBoxSpacing(edit) } : {}),
    ...(runRows ? { runRows: true, browserSpacing: pdfSourceTextBoxSpacing(edit) } : {}),
  });
  let drawn_ = rendered;
  // Rows the frame cannot hold (a line set bolder or larger wraps): the frame
  // grows down to hold them, to the page's bottom at most, as WPS grows a
  // text block; refused, the write retried in a loop and kept the draft
  
  if (drawn_ && onPortableLayout && !drawn_.portableLayout && options.onGrownHeight) {
    const own = Number(edit.height);
    // Down while the page has room; a box at the foot of the page grows
    // up past it (the record moves it up: pdfTextBoxGrownFrame), so what
    // was typed stays on the page instead of the save being refused.
    for (const factor of [1.5, 2, 3, 4, 6, 10, 16]) {
      const height = Math.min(1, own * factor);
      if (!(height > own)) break;
      const taller = await renderPdfTextOpToImage({ ...portableOp, maxHeight: height * renderedHeight }, {
        backgroundColor: clearBackground ? undefined : backgroundColor,
        ...(edit.sourceText ? { browserLayout: true, displayFontFamily,
          browserLineHeightRatio: pdfSourceTextBoxLineHeightRatio(edit.sourceText),
          browserSpacing: pdfSourceTextBoxSpacing(edit) } : {}),
        ...(runRows ? { runRows: true, browserSpacing: pdfSourceTextBoxSpacing(edit) } : {}),
      });
      if (taller?.portableLayout) {
        drawn_ = taller;
        options.onGrownHeight(height);
        break;
      }
      if (height >= 1) break;
    }
  }
  const renderedFinal = drawn_;
  if (!renderedFinal) return refuseAppearance(refusalId, "raster", { width: portableOp.maxWidth, height: portableOp.maxHeight });
  if (onPortableLayout) {
    if (!renderedFinal.portableLayout) return refuseAppearance(refusalId, "portable-layout");
    onPortableLayout(renderedFinal.portableLayout);
  }
  const dataUrl = rasterBytesToDataUrl(
    renderedFinal.bytes,
    renderedFinal.mimeType || "image/png",
  );
  return sanitizeManagedNativeTextBoxImage(dataUrl);
}



function pdfSourceTextBoxSpacing(edit: TextBoxEditLike): PdfBrowserTextSpacing | undefined {
  const size = Number(edit.fontSize);
  if (!(size > 0)) return undefined;
  const typography = readPdfTextBoxTypography(edit, size);
  const spacing: PdfBrowserTextSpacing = {
    ...(typography.characterSpacing ? { characterEm: typography.characterSpacing / size } : {}),
    ...(typography.wordSpacing ? { wordEm: typography.wordSpacing / size } : {}),
    ...(typography.horizontalScale ? { horizontalScale: typography.horizontalScale } : {}),
  };
  return Object.keys(spacing).length ? spacing : undefined;
}
