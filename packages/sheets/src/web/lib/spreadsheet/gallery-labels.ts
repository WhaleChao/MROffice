/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  CELL_STYLE_GALLERY_GROUPS,
  CELL_STYLE_GALLERY_STYLES,
  type NamedCellStyle,
  type NamedCellStyleGroup,
} from "./cell-styles";
import {
  TABLE_STYLE_PRESETS,
  tableStyleLabelMetadata,
  type TableStylePreset,
} from "@mrnote/core/web/lib/spreadsheet/data-table.ts";

type Translate = (
  key: string,
  values?: Record<string, string | number>,
) => string;
const cellStyleIds = new Set(
  CELL_STYLE_GALLERY_STYLES.map((style) => style.id),
);
const cellStyleGroupIds = new Set(
  CELL_STYLE_GALLERY_GROUPS.map((group) => group.id),
);
const tableStyleIds = new Set(TABLE_STYLE_PRESETS.map((style) => style.id));

/** Only registry-owned labels are translated; user-authored names remain data. */
export function spreadsheetCellStyleLabel(
  style: NamedCellStyle,
  t: Translate,
): string {
  return cellStyleIds.has(style.id)
    ? t(`spreadsheet.dynamicCopy.cellStyle.${style.id}`)
    : style.label;
}

export function spreadsheetCellStyleGroupLabel(
  group: NamedCellStyleGroup,
  t: Translate,
): string {
  // The workbook's own styles head the gallery as Excel's "Custom" (XLS-A012).
  return cellStyleGroupIds.has(group.id) || group.id === "custom"
    ? t(`spreadsheet.dynamicCopy.cellStyleGroup.${group.id}`)
    : group.label;
}

/** Native recipe metadata owns colour/tone/ordinal; display text is never parsed. */
export function spreadsheetTableStyleLabel(
  style: TableStylePreset,
  t: Translate,
  accents?: readonly string[],
): string {
  if (!tableStyleIds.has(style.id)) return style.label.en;
  if (style.id === "none") return t("spreadsheet.dynamicCopy.tableStyleNone");
  const metadata = tableStyleLabelMetadata(style, accents);
  return t("spreadsheet.dynamicCopy.tableStyle", {
    color: t(`spreadsheet.dynamicCopy.tableStyleColor.${metadata.color}`),
    tone: t(`spreadsheet.dynamicCopy.tableStyleTone.${style.tone}`),
    index: metadata.ordinal,
  });
}
