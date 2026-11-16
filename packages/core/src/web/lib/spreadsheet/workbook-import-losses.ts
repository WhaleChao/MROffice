/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What a workbook import could not carry, in words a reader understands.
 *
 * ONE table for BOTH import chains. The browser importer
 * (`workbook-import.ts`) shows these before the user consents to a lossy
 * import; the server projection (`editable_page._xlsx_content`, the path
 * import batches and attachments take) writes the same codes into
 * `content_json.importLosses`, and the page's own status bar reads them here.
 * A code with text in only one of those places would mean one door telling the
 * truth and the other one not.
 *
 * Lives apart from the importer so a page that only needs to NAME a loss does
 * not pull the whole xlsx reader into its bundle.
 */

export interface SpreadsheetImportLossEntry {
  readonly code: string;
  readonly affected: number;
}

/**
 * A viewer's locale, or the page's `console-notebooks` translator. The
 * translator reads `spreadsheet.importLoss.<code>` in every locale; a plain
 * locale string reads the English table below (the catalogue's `en` words).
 */
export type SpreadsheetImportLossLocale = string | ((key: string) => string);

const LOSS_LABELS_EN: Record<string, string> = {
  chart_subtype_unsupported: "Some chart subtypes are unsupported",
  xlsx_relationship_part_ambiguous:
    "Worksheet relationships cannot be identified safely",
  xlsx_chart_series_not_projected:
    "Some chart series cannot be converted safely",
  xlsx_chart_legend_position_not_projected:
    "A chart legend position cannot be converted safely",
  xlsx_chart_series_paint_not_projected:
    "Some chart series or slice colors cannot be converted safely",
  xlsx_chart_format_not_projected: "Some chart formatting cannot be converted safely",
  xlsx_chart_line_marker_not_projected:
    "Some line-chart markers cannot be converted safely",
  xlsx_chart_line_width_not_projected:
    "Some line-chart widths cannot be converted safely",
  xlsx_chart_value_axis_number_format_not_projected:
    "A chart primary value-axis number format cannot be converted safely",
  spreadsheet_cells_omitted: "The workbook exceeds the cell import limit",
  spreadsheet_objects_omitted: "The workbook exceeds the object import limit",
  xlsx_drawing_group_not_projected:
    "Grouped objects were not imported to avoid incorrect overlap",
  xlsx_drawing_text_box_format_not_projected:
    "Text box line breaks, alignment, or colors will be simplified",
  xlsx_drawing_shape_not_projected: "A drawing shape could not be recognized",
  xlsx_drawing_picture_metafile_not_drawn:
    "A picture is a vector metafile (EMF/WMF) the browser cannot draw (the original is kept)",
  xlsx_drawing_object_link_not_projected:
    "Drawing object links are not projected yet",
  worksheets_omitted: "The workbook exceeds the worksheet import limit",
  xlsx_cell_text_truncated:
    "Some cells hold more than Excel's 32,767 characters: as Excel does, only the first 32,767 are shown here; the file keeps the full text until you edit the cell",
  xlsx_defined_name_not_projected:
    "Some defined names are over the editor's limit: they stay in the file and keep working in formulas, but are not shown here and cannot be edited",
  xlsx_defined_name_unsupported:
    "Some defined names are written in a way the editor cannot read: they stay in the file and keep working in Excel, but are not shown here and cannot be edited",
  xlsx_line_sizes_not_projected:
    "Some row heights or hidden rows are over the editor's limit: they stay in the file, but here those rows show at the default height and are not hidden",
  xlsx_row_formats_not_projected:
    "Some row formats are over the editor's limit: they stay in the file, but here a cell typed on those rows past the used range does not take the row's format",
  xlsx_merges_not_projected:
    "Some merged cells are over the editor's limit: they stay merged in the file, but here those cells show unmerged",
  xlsx_validations_not_projected:
    "Some data validation rules are over the editor's limit: they stay in the file, but here those cells accept any entry and show no dropdown or input message",
  xlsx_autofilter_criteria_not_projected:
    "AutoFilter criteria are not projected yet",
  xlsx_autofilter_color_criteria_not_projected:
    "Color filter criteria are not projected yet",
  xlsx_autofilter_custom_criteria_not_projected:
    "Custom filter operators are not projected yet",
  xlsx_autofilter_dynamic_criteria_not_projected:
    "Dynamic date or average filters are not projected yet",
  xlsx_autofilter_icon_criteria_not_projected:
    "Icon filter criteria are not projected yet",
  xlsx_autofilter_top10_criteria_not_projected:
    "Top 10 item or percent filters are not projected yet",
  xlsx_autofilter_unknown_criteria_not_projected:
    "An unrecognized filter criterion is not projected",
  xlsx_autofilter_value_criteria_not_projected:
    "Value-list filter criteria are not projected yet",
  xlsx_conditional_format_rule_not_projected:
    "Some conditional formatting rules are not projected yet",
  xlsx_conditional_format_style_not_projected:
    "Conditional formatting differential styles are not projected yet",
  xlsx_data_validation_rule_not_projected:
    "Some data validation rules are not projected yet",
  xlsx_defined_names_not_projected:
    "Some workbook defined names cannot be projected safely",
  xlsx_formula_not_materialized: "Some special formulas cannot be materialized",
  xlsx_layout_enrichment_failed: "Some workbook layout could not be read",
  xlsx_merge_range_not_projected:
    "Some merged ranges exceed the editable workbook boundary",
  xlsx_outline_group_not_projected:
    "Outline groups will be reduced to hidden rows or columns",
  xlsx_additional_sheet_views_not_projected:
    "Additional worksheet views are not projected yet",
  xlsx_frozen_pane_not_projected:
    "The frozen-pane boundary could not be projected",
  xlsx_gradient_fill_not_projected: "Gradient fills are not projected yet",
  xlsx_connections_not_projected:
    "External data connections are not projected yet",
  xlsx_query_tables_not_projected: "Query tables are not projected yet",
  xlsx_xml_maps_not_projected: "XML maps are not projected yet",
  xlsx_custom_xml_not_projected: "Custom XML parts are not projected yet",
  xlsx_web_extensions_not_projected: "Office add-ins are not projected yet",
  xlsx_rich_data_not_projected:
    "Linked data types (stocks, currencies) are not projected yet",
  xlsx_cell_picture_not_projected:
    "A picture placed in a cell could not be read; the cell shows #VALUE!",
  xlsx_volatile_dependencies_not_projected:
    "Real-time data caches are not projected yet",
  xlsx_smartart_not_projected: "SmartArt graphics are not projected yet",
  xlsx_form_controls_not_projected: "Form controls are not projected yet",
  xlsx_ole_objects_not_projected: "Embedded objects are not projected yet",
  xlsx_macros_not_projected: "Macros (VBA) never run and are not projected",
  xlsx_native_slicers_not_projected:
    "Native slicer definitions are not projected yet",
  xlsx_native_timelines_not_projected:
    "Some timelines could not be read and are not projected",
  xlsx_external_links_not_projected:
    "External workbook links are not projected yet",
  xlsx_scenarios_not_projected:
    "Scenario Manager scenarios are not projected yet",
  xlsx_data_consolidate_not_projected:
    "Data consolidation settings are not projected yet",
  xlsx_smart_tags_not_projected: "Smart tags are not projected yet",
  xlsx_web_publish_items_not_projected:
    "Web publishing settings are not projected yet",
  xlsx_custom_properties_not_projected:
    "Worksheet custom properties are not projected yet",
  xlsx_ignored_errors_not_projected:
    "Some ignored-error markers are not projected yet",
  xlsx_phonetic_text_not_projected:
    "Phonetic (ruby) text is not projected yet",
  xlsx_vml_drawing_not_projected:
    "Legacy VML drawing objects are not projected yet",
  xlsx_dialog_sheet_not_projected: "A dialog sheet imports as a blank page",
  xlsx_header_footer_code_not_projected:
    "A header or footer contains an unrecognized control code",
  xlsx_header_footer_page_offset_not_projected:
    "A header or footer page-number offset is not projected yet",
  xlsx_header_footer_picture_not_projected:
    "A header or footer picture could not be read",
  xlsx_header_footer_truncated:
    "Over-long header or footer text was truncated",
  xlsx_manual_page_breaks_not_projected:
    "Some manual page breaks could not be projected",
  xlsx_page_margins_not_projected: "Page margins could not be projected",
  xlsx_page_setup_option_not_projected:
    "Some page-setup options are not projected yet",
  xlsx_paper_size_not_projected: "The paper size is not projected yet",
  xlsx_pattern_fill_approximated:
    "Pattern fills will be approximated with solid colors",
  xlsx_pattern_fill_not_projected:
    "An unrecognized pattern fill is not projected",
  xlsx_pivot_aggregation_not_projected:
    "Some PivotTable aggregations are not projected yet",
  xlsx_pivot_definition_not_projected:
    "A PivotTable definition could not be projected",
  xlsx_print_area_not_projected: "The print area could not be projected",
  xlsx_print_centering_not_projected: "Print centering is not projected yet",
  xlsx_print_fit_dimensions_not_projected:
    "The combined width/height fit policy is not projected yet",
  xlsx_print_options_not_projected: "Some print options could not be projected",
  xlsx_print_orientation_not_projected:
    "The print orientation could not be projected",
  xlsx_print_scaling_not_projected:
    "The print scaling policy could not be projected",
  xlsx_print_titles_not_projected:
    "Repeating print-title rows or columns could not be projected",
  xlsx_printer_resolution_not_projected:
    "Printer resolution is not projected yet",
  xlsx_printer_settings_not_projected:
    "Device-specific printer settings are not projected yet",
  xlsx_row_height_clamped:
    "Rows below the editable minimum height will be expanded",
  xlsx_shared_formula_anchor_missing: "A shared formula anchor is missing",
  xlsx_sort_state_not_projected:
    "Native worksheet sort conditions are not projected yet",
  xlsx_split_pane_not_projected: "A split-pane boundary could not be projected",
  xlsx_table_definition_not_projected:
    "Structured table definitions are not projected yet",
  xlsx_table_header_not_projected:
    "A header-less table will use its first row as the header",
  xlsx_table_totals_function_not_projected:
    "Some table total functions are not supported",
  xlsx_table_totals_label_not_projected:
    "Table total-row labels are not supported",
  xlsx_theme_color_unresolved: "Some theme colors could not be resolved",
  xlsx_worksheet_part_missing: "A worksheet XML part is missing",
};

/** The catalogue key of a code's words, or null when this table has none. */
export function spreadsheetImportLossKey(code: string): string | null {
  return Object.prototype.hasOwnProperty.call(LOSS_LABELS_EN, code)
    ? `spreadsheet.importLoss.${code}`
    : null;
}

/**
 * The reader's own words for one code, or null when this table has none.
 *
 * The page's shared import notice (`ImportedEditableStatusBar`, reading
 * `@/lib/import-losses`) names the codes it has a message for and falls back to
 * this table for every other workbook code, so a server-projected loss reaches
 * the user in the same words the browser importer uses before a lossy import.
 */
export function spreadsheetImportLossLabel(
  code: string,
  locale: SpreadsheetImportLossLocale,
): string | null {
  if (typeof locale === "function") {
    const key = spreadsheetImportLossKey(code);
    return key ? locale(key) : null;
  }
  const labels = LOSS_LABELS_EN;
  return labels[code] || null;
}

/** One line per loss: the reader's own words and how many objects it cost. */
export function describeSpreadsheetImportLosses(
  entries: readonly SpreadsheetImportLossEntry[],
  locale: SpreadsheetImportLossLocale,
): string[] {
  if (typeof locale === "function") {
    return entries.map(
      (entry) =>
        `${spreadsheetImportLossLabel(entry.code, locale) || entry.code} (${entry.affected})`,
    );
  }
  const labels = LOSS_LABELS_EN;
  return entries.map(
    (entry) => `${labels[entry.code] || entry.code} (${entry.affected})`,
  );
}
