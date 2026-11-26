/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface WholeValuePathEntry {
  readonly pattern: string;
  readonly lane: string;
  readonly why: string;
}

export const WHOLE_VALUE_PATHS: readonly WholeValuePathEntry[] = [
  {
    pattern: "slides.*.blocks.*.chartV2.data",
    lane: "ppt",
    why: "a deck chart's categories, series names and values agree with each other and with the chart's workbook (3610 11-26: two windows' refreshes left [\"1\",\"1\"] beside one workbook row, every save refused)",
  },
  
  {
    pattern: "slides.*.blocks.*.data",
    lane: "ppt",
    why: "a chart's own data table: its headers and rows agree with each other (merged row by row a row grew past the headers)",
  },
  {
    pattern: "masters.*.blocks.*.data",
    lane: "ppt",
    why: "a master's chart's data table, as slides.*.blocks.*.data",
  },
  {
    pattern: "masters.*.layouts.*.blocks.*.data",
    lane: "ppt",
    why: "a layout's chart's data table, as slides.*.blocks.*.data",
  },
  {
    pattern: "slides.*.blocks.*.groupPath",
    lane: "ppt",
    why: "a shape's group path (outermost group to innermost) is one place in the nesting: merged two regroupings name groups that hold nothing",
  },
  // sheet: a workbook's room is one Y doc per shard (lib/collab/workbook-shards.ts), so a sheet chart's data sits at
  // these three paths of its shard's root, not under `sheets.*`: a native chart and an imported (drawing) chart in
  // a sheet's own shard, and an imported chart in a big sheet's object tile (`objects` of `{ index, value }`)
  {
    pattern: "charts.*.chartV2.data",
    lane: "sheet",
    why: "a sheet chart's categories, series names and values are one reading of its range (merged, two windows' rereads grew the list as a deck chart's did)",
  },
  {
    pattern: "objects.*.chartV2.data",
    lane: "sheet",
    why: "an imported sheet chart's data: one reading of its series references, kept together",
  },
  {
    pattern: "objects.*.value.chartV2.data",
    lane: "sheet",
    why: "the same chart in a big sheet's object tile shard",
  },
  
  // one tuple or one reading, at their shard paths (a workbook's own keys are under `workbook` in its manifest shard)
  {
    pattern: "workbook.nativeTableStyles.*.elements.*.backgroundColor",
    lane: "sheet",
    why: "a table style's colour is [theme slot, tint]: merged index by index it became three numbers, which the style reader drops",
  },
  {
    pattern: "workbook.nativeTableStyles.*.elements.*.color",
    lane: "sheet",
    why: "a table style's text colour, [theme slot, tint], as backgroundColor",
  },
  {
    pattern: "workbook.nativeTableStyles.*.elements.*.border.*.color",
    lane: "sheet",
    why: "a table style border's colour, [theme slot, tint], as backgroundColor",
  },
  {
    pattern: "hiddenRowRuns",
    lane: "sheet",
    why: "the hidden rows' runs ([first, last] each) are one reading of which rows are hidden: two ends from two writers name rows neither hid",
  },
  {
    pattern: "conditionalFormats.*.thresholds",
    lane: "sheet",
    why: "a colour scale's or data bar's thresholds are one scale (min, mid, max): merged they grow past the rule's stops",
  },
  {
    pattern: "conditionFilters.*.operands",
    lane: "sheet",
    why: "a filter condition's operands are its operator's arguments (one, or two for between)",
  },
  {
    pattern: "conditionFilters.*.conditions.*.operands",
    lane: "sheet",
    why: "the operands of one of a filter's two conditions, as conditionFilters.*.operands",
  },
  {
    pattern: "dataTables.*.columnFilters.*.condition.conditions.*.operands",
    lane: "sheet",
    why: "a table column filter condition's operands, as conditionFilters.*.operands",
  },
  {
    pattern: "objects.*.data",
    lane: "sheet",
    why: "a chart's own data table: its headers and rows agree with each other (a chart object in a sheet's shard)",
  },
  {
    pattern: "objects.*.value.data",
    lane: "sheet",
    why: "the same data table of a chart in a big sheet's object tile shard",
  },
  // pdf: a note's, a markup's or a stroke's coordinates are one shape. Written index by index, two people moving the
  // same note at once merged x of one with y of the other, or left eight numbers where four are, which the page's
  
  {
    pattern: "annotations.*.rect",
    lane: "pdf",
    why: "a note's rectangle [x1, y1, x2, y2] is one place: two concurrent moves give one person's rectangle, never x of one and y of the other, never eight numbers (a rect that is not four numbers refuses the page's save)",
  },
  {
    pattern: "annotations.*.quadPoints",
    lane: "pdf",
    why: "a markup's quadrilaterals (eight numbers each) outline the marked words together: merged index by index they outline no words",
  },
  {
    pattern: "annotations.*.managedMarkup.rect",
    lane: "pdf",
    why: "the window's own record of a markup keeps the note's rectangle: one place, as annotations.*.rect",
  },
  {
    pattern: "annotations.*.managedMarkup.quadPoints",
    lane: "pdf",
    why: "the window's own record of a markup keeps the marked words' quadrilaterals: one outline, as annotations.*.quadPoints",
  },
  {
    pattern: "inkStrokes.*.points",
    lane: "pdf",
    why: "a pen stroke's points are one line: two people's points merged index by index draw neither stroke",
  },
  {
    pattern: "textBoxEdits.*.sourceText.originalProof.evidence.selection.item.transform",
    lane: "pdf",
    why: "a page text item's transform [a, b, c, d, e, f] is one matrix: half of one and half of another places the text where it never was",
  },
  {
    pattern: "textBoxEdits.*.sourceText.pieces.*.codeRanges",
    lane: "pdf",
    why: "a text piece's code ranges ([start, end] each) are one reading of its text: an end from each writer names characters neither chose (registered as the whole list -- a range is an entry of a list, which the room writes in place)",
  },
  
  // footer's table states its column widths three times (percent, px, twips), in the document's attribute map
  // (collaborationDocumentAttrs: a header or footer at its top level, or in a section band, kept as a keyed list
  // entry whose value sits under __mrai_document_attr_crdt_value__)
  {
    pattern: "*.*.columnWidthsPct",
    lane: "docs",
    why: "a header or footer table's column widths in percent: one width per column, read with the other two units",
  },
  {
    pattern: "*.*.column_widths_px",
    lane: "docs",
    why: "the same widths in px, as columnWidthsPct",
  },
  {
    pattern: "*.*.nativeColumnWidthsTwips",
    lane: "docs",
    why: "the same widths in twips (as Word writes them), as columnWidthsPct",
  },
  {
    pattern: "sectionBands.*.__mrai_document_attr_crdt_value__.*.*.columnWidthsPct",
    lane: "docs",
    why: "a section band's header or footer table, as *.*.columnWidthsPct",
  },
  {
    pattern: "sectionBands.*.__mrai_document_attr_crdt_value__.*.*.column_widths_px",
    lane: "docs",
    why: "a section band's header or footer table, as *.*.column_widths_px",
  },
  {
    pattern: "sectionBands.*.__mrai_document_attr_crdt_value__.*.*.nativeColumnWidthsTwips",
    lane: "docs",
    why: "a section band's header or footer table, as *.*.nativeColumnWidthsTwips",
  },
];
