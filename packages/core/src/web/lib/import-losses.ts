/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What an imported file holds that its page does not.
 *
 * One reader for every host. The server projections record this under a single
 * key — `importLosses` on the page, `import_losses` on the imported resource's
 * own metadata — as `[{ code, omitted }]` sorted by code
 * (`app/services/resources/import_losses.py`). The workbook projection has
 * written that record since the drawing-shape work; the deck projection writes
 * it now; this module is the one place that reads it, so a notice never has to
 * be built twice.
 *
 * A loss entry is a statement about the PAGE, never about the file: the source
 * package is preserved and exported untouched in every case recorded here.
 */

export interface ImportLossEntry {
  readonly code: string;
  readonly omitted: number;
}

/** A page records a bounded vocabulary; anything longer is not a record. */
const MAX_IMPORT_LOSS_ENTRIES = 64;

/**
 * The message each code is named by. A code with no entry here still shows —
 * as its own code — because a projection that learns a new refusal must not go
 * quiet again while the translation catches up.
 */
const IMPORT_LOSS_MESSAGE_KEY: Readonly<Record<string, string>> = {
  pptx_slide_not_projected: "pptxSlide",
  pptx_diagram_not_projected: "pptxDiagram",
  pptx_unreadable_diagram_not_projected: "pptxUnreadableDiagram",
  
  pptx_grouped_diagram_not_projected: "pptxGroupedDiagram",
  // An equation text box PowerPoint keeps in mc:AlternateContent; the deck
  
  pptx_equation_not_projected: "pptxEquation",
  pptx_diagram_shape_geometry_not_projected: "pptxDiagramShapeGeometry",
  pptx_chart_not_projected: "pptxChart",
  pptx_picture_not_projected: "pptxPicture",
  pptx_slide_blocks_not_projected: "pptxSlideBlocks",
  // A layout thumbnail the master view draws as the silhouette: the layout
  // itself and every slide on it are whole (SHAPE-IMPORT-AUDIT.md §72).
  pptx_layout_preview_not_projected: "pptxLayoutPreview",
  // The master's own card, same rule, counted per master (§73).
  pptx_master_preview_not_projected: "pptxMasterPreview",
  // A slide whose design receipts (the file's layout catalog and Slide Size
  // authority) the page could not carry: the slide is whole and editable,
  
  pptx_design_authority_not_projected: "pptxDesignAuthority",
  // A picture drawn from a display copy because the deck's pictures took
  // the page past its size: on the slide and editable; the file and every
  // export keep the original.
  pptx_picture_display_copy: "pptxPictureDisplayCopy",
  
  
  // export keep the original.
  pptx_picture_src_refused: "pptxPictureSrcRefused",
  pptx_shape_not_projected: "pptxShape",
  // Distinct from `pptxShape` on purpose: that one is the BROWSER preview's
  // gap, and its message says the saved page reads those shapes in full. This
  // is the saved page's own -- a shape whose words are on no block of it.
  pptx_text_shape_not_projected: "pptxTextShape",
  // Everything else a slide DRAWS that reached no block: a line, a connector,
  // a filled freeform, a picture, a table, an embedded object's picture
  
  pptx_slide_object_not_projected: "pptxSlideObject",
  // A shape whose words are on the page as a rectangular text box while its
  // own outline (rounded corners, oval, callout, freeform) is not
  
  pptx_shape_geometry_not_projected: "pptxShapeGeometry",
  // A shape whose WordArt warp (`a:prstTxWarp`, anything but `textNoShape`)
  // bends its words in PowerPoint and not on this page (§11).
  pptx_text_warp_not_drawn: "pptxTextWarp",
  xlsx_drawing_shape_not_projected: "xlsxDrawingShape",
  xlsx_drawing_picture_not_projected: "xlsxDrawingPicture",
  xlsx_drawing_shape_adjustment_not_projected: "xlsxDrawingShapeAdjustment",
  xlsx_drawing_shape_picture_fill_not_drawn: "xlsxDrawingShapePictureFill",
  xlsx_drawing_shape_3d_not_drawn: "xlsxDrawingShape3d",
  xlsx_drawing_shape_text_layout_not_drawn: "xlsxDrawingShapeTextLayout",
  // Docs. The projection has recorded these as named receipts all along; what
  // they never had was a surface a reader sees
  // (docs/office-graphics/SHAPE-IMPORT-AUDIT.md §39).
  docx_shape_shadow_not_drawn: "docxShapeShadow",
  docx_text_box_not_projected: "docxTextBox",
  docx_header_footer_shape_not_projected: "docxHeaderFooterShape",
  docx_header_footer_placement_collapsed: "docxHeaderFooterPlacement",
  docx_header_footer_text_box_table_flattened: "docxHeaderFooterTextBoxTable",
  // A table in a text box or shape story the page does not hold (a box left
  // in the flow, a drawn shape's or group's story, a box in a cell).
  docx_text_box_table_not_projected: "docxTextBoxTable",
  docx_media_relationship_unresolved: "docxMediaRelationship",
  docx_paragraph_not_projected: "docxParagraph",
  // A drawing the projection read and the page builder could not publish
  // (a part of it fails the page's own model): said, never silent.
  docx_shape_not_projected: "docxShape",
  // An embedded object (Insert > Object) the page shows by its picture only:
  // linked, floating, or an embedding too large or of a kind the page cannot
  // carry (`document_docx_ole_object`). The source file keeps the object.
  docx_ole_object_not_carried: "docxOleObject",
  // A picture whose media is an EMF/WMF vector metafile. These three come from
  // the BROWSER importers (`officeCompiler.ts`), which cannot rasterise: the
  // server renders one to a PNG stand-in at the metafile's own frame and so
  // never reports it (SHAPE-IMPORT-AUDIT.md §53, §55). They live in this shared
  // table because the two import doors must name the same object with the same
  // words, and because a page that ever carries one must not show it raw.
  docx_picture_metafile_not_drawn: "docxPictureMetafile",
  xlsx_drawing_picture_metafile_not_drawn: "xlsxDrawingPictureMetafile",
  // A picture placed in a cell (Excel's Place in Cell) whose rich value or
  // media could not be read; the cell keeps its #VALUE! fallback.
  xlsx_cell_picture_not_projected: "xlsxCellPicture",
  // A header/footer picture (&G) whose VML shape or image could not be read;
  // its section keeps the &[Picture] code, so the source picture is kept.
  xlsx_header_footer_picture_not_projected: "xlsxHeaderFooterPicture",
  pptx_picture_metafile_not_drawn: "pptxPictureMetafile",
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** Parse one projection's loss record. Anything malformed reads as no record. */
export function readImportLosses(value: unknown): ImportLossEntry[] {
  if (!Array.isArray(value)) return [];
  const entries: ImportLossEntry[] = [];
  for (const raw of value.slice(0, MAX_IMPORT_LOSS_ENTRIES)) {
    if (!isRecord(raw)) continue;
    const code = raw.code;
    const omitted = raw.omitted;
    if (
      typeof code !== "string" ||
      !/^[a-z][a-z0-9_]{0,79}$/.test(code) ||
      typeof omitted !== "number" ||
      !Number.isSafeInteger(omitted) ||
      omitted <= 0
    )
      continue;
    entries.push({ code, omitted });
  }
  return entries;
}

/**
 * The record for one page, from whichever channel carries it.
 *
 * The resource's metadata is preferred because it survives editing: an editor
 * rebuilds the page body from its own model on save, so the copy on
 * `content_json` is the import's own and only lasts until the first save.
 */
export function readPageImportLosses(
  resourceMetadataValue: unknown,
  pageContentJson: unknown,
): ImportLossEntry[] {
  const fromMetadata = readImportLosses(resourceMetadataValue);
  if (fromMetadata.length > 0) return fromMetadata;
  return readImportLosses(
    isRecord(pageContentJson) ? pageContentJson.importLosses : undefined,
  );
}

/** The `importStatus.projectionLossItem.*` leaf naming this code, if any. */
export function importLossMessageKey(code: string): string | null {
  return IMPORT_LOSS_MESSAGE_KEY[code] ?? null;
}

/** How many objects a record accounts for in total. */
export function importLossTotal(entries: readonly ImportLossEntry[]): number {
  return entries.reduce((total, entry) => total + entry.omitted, 0);
}
