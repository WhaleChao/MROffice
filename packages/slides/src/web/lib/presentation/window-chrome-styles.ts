/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Inline style objects shared by the presentation Ribbon and panes, and the
 * fixed command lists of the table and change-shape context menus.
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */

import type { CSSProperties } from "react";
import type { ShapeKind } from "@mrnote/core/web/lib/presentation-shapes.ts";

export const PRESENTATION_RIBBON_LABELED_BUTTON_STYLE: CSSProperties = {
  width: "auto",
  minWidth: 34,
  padding: "0 8px",
  gap: 5,
  whiteSpace: "nowrap",
  // A LABELLED ribbon button must never be narrower than its label.
  //
  // `width: "auto"` above looks like it says that and does not: the shared
  // toolbar class sizes ICON buttons with `flex-basis: 28px`, and in a flex row
  // flex-basis wins over width. So a labelled button rendered at 34px (the
  // min-width floor) while its `nowrap` text kept its natural 59px, spilled out
  // of the button, and printed on top of its neighbour — "幻灯片设置" overlapping
  // "幻灯片大小" in the Design tab, reported from real use.
  //
  // `flexBasis: "auto"` hands sizing back to the content; `flexShrink: 0` keeps a
  // crowded row from re-introducing the same crush. The ribbon already pages
  // horizontally when a row does not fit (`computeRibbonOverflow`) — which is
  // what PowerPoint and WPS do — so refusing to shrink costs nothing: the row
  // scrolls instead of deforming.
  flexBasis: "auto",
  flexShrink: 0,
};

/**
 * The table's structural commands a cell's right-click menu offers under its
 * Table ▸ row (PowerPoint for the web keeps them on the Table tab; the
 * product keeps them reachable here), in their order. Data only -- the
 * dispatcher lives in the component so the array stays a plain literal that
 * captures nothing. Split sits on the menu itself, as Office's does.
 */
export const TABLE_CONTEXT_COMMANDS = [
  { key: "insert-row-above", labelKey: "presentation.contextMenu.tableInsertRowAbove", icon: "insertRows" },
  { key: "insert-row-below", labelKey: "presentation.contextMenu.tableInsertRowBelow", icon: "insertRows" },
  { key: "insert-column-left", labelKey: "presentation.contextMenu.tableInsertColumnLeft", icon: "insertColumns" },
  { key: "insert-column-right", labelKey: "presentation.contextMenu.tableInsertColumnRight", icon: "insertColumns" },
  { key: "delete-row", labelKey: "presentation.contextMenu.tableDeleteRow", icon: "delete" },
  { key: "delete-column", labelKey: "presentation.contextMenu.tableDeleteColumn", icon: "delete" },
  { key: "merge", labelKey: "presentation.contextMenu.tableMergeCells", icon: "mergeCells" },
  { key: "split", labelKey: "presentation.contextMenu.split", icon: "splitCells" },
] as const;

export type TableContextCommand = (typeof TABLE_CONTEXT_COMMANDS)[number]["key"];

/**
 * Change Shape ▸ on the right-click menu: the presets PowerPoint for the web
 * reaches for first (its gallery's basic shapes, block arrows and a callout),
 * not the whole 190-kind catalog, which is the Shape Format tab's gallery.
 */
export const CONTEXT_MENU_CHANGE_SHAPE_KINDS: readonly ShapeKind[] = [
  "rectangle",
  "rounded-rectangle",
  "ellipse",
  "triangle",
  "right-triangle",
  "diamond",
  "parallelogram",
  "trapezoid",
  "pentagon",
  "hexagon",
  "octagon",
  "star",
  "chevron",
  "right-arrow",
  "left-arrow",
  "up-arrow",
  "down-arrow",
  "plus",
  "heart",
  "cloud",
  "callout-1",
];

/* blockOrderButtonStyle is a 22x20 ICON slot — buttons that also carry a
   text label (replace/reset/crop) rendered their CJK label vertically inside
   the 22px column. This variant sizes to content. */
export const labeledActionButtonStyle = {
  display: "inline-flex",
  alignItems: "center",
  gap: 4,
  height: 24,
  borderRadius: 4,
  border: "1px solid var(--border, #e5e7eb)",
  background: "#fff",
  color: "#475569",
  cursor: "pointer",
  padding: "0 8px",
  whiteSpace: "nowrap",
} as const;

export const blockOrderButtonStyle = {
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
  width: 22,
  height: 20,
  borderRadius: 4,
  border: "1px solid var(--border, #e5e7eb)",
  background: "#fff",
  color: "#475569",
  cursor: "pointer",
  padding: 0,
} as const;

/** Shared inline-row style for the shape inspector's labelled controls. */
export const inspectorChipStyle = {
  display: "inline-flex",
  alignItems: "center",
  // The format pane is narrow: a label and its controls wrap onto the next
  // line rather than squeezing the label to one character per line.
  flexWrap: "wrap",
  maxWidth: "100%",
  gap: 4,
  rowGap: 4,
  fontSize: 12,
} as const;
