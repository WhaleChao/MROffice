/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { createElement, memo, type ReactNode } from "react";

import type { OfficeSemanticIconId } from "./officeIconRegistry";
import {
  OFFICE_ICON_EMPHASIS_STROKE_WIDTH,
  OFFICE_ICON_STROKE_WIDTH,
} from "./officeIconTokens";

export type OfficeIconColorRole =
  | "ink"
  | "ink-soft"
  | "paper"
  | "accent"
  | "accent-fill"
  | "accent-soft-fill"
  | "warm"
  | "warm-fill";

interface OfficeIconPrimitiveBase {
  role: OfficeIconColorRole;
  emphasis?: true;
  strokeDasharray?: string;
}

type OfficeIconPrimitive =
  | (OfficeIconPrimitiveBase & {
      element: "path";
      d: string;
    })
  | (OfficeIconPrimitiveBase & {
      element: "rect";
      x: number;
      y: number;
      width: number;
      height: number;
      rx?: number;
    })
  | (OfficeIconPrimitiveBase & {
      element: "circle";
      cx: number;
      cy: number;
      r: number;
    })
  | (OfficeIconPrimitiveBase & {
      element: "line";
      x1: number;
      y1: number;
      x2: number;
      y2: number;
    })
  | (OfficeIconPrimitiveBase & {
      element: "polyline" | "polygon";
      points: string;
    });

type OfficeIconGeometry = readonly OfficeIconPrimitive[];

const TABLE_SURFACE: OfficeIconGeometry = [
  {
    element: "rect",
    role: "paper",
    x: 3,
    y: 4,
    width: 18,
    height: 16,
    rx: 1,
  },
];

const TABLE_GRID_LINES: OfficeIconGeometry = [
  {
    element: "path",
    role: "ink-soft",
    d: "M3.5 9.3h17M3.5 14.7h17M9 4.5v15M15 4.5v15",
  },
];

const COMPACT_TABLE: OfficeIconGeometry = [
  {
    element: "rect",
    role: "paper",
    x: 3,
    y: 3.5,
    width: 14,
    height: 14.5,
    rx: 1,
  },
  {
    element: "path",
    role: "ink-soft",
    d: "M3.5 8.3h13M3.5 13.1h13M8 4v13.5M12.5 4v13.5",
  },
];

const IMAGE_BASE: OfficeIconGeometry = [
  {
    element: "rect",
    role: "paper",
    x: 3,
    y: 4,
    width: 16,
    height: 14,
    rx: 1.3,
  },
  { element: "circle", role: "ink-soft", cx: 14.8, cy: 7.7, r: 1.25 },
  {
    element: "polyline",
    role: "accent",
    points: "4.2,16.7 8.8,11.7 11.5,14.2 13.6,12.3 17.8,16.7",
  },
];

const SHAPE_BASE: OfficeIconGeometry = [
  {
    element: "rect",
    role: "paper",
    x: 4,
    y: 4,
    width: 11,
    height: 11,
    rx: 1.5,
  },
  { element: "circle", role: "accent", cx: 16.5, cy: 16.2, r: 4.2 },
];

const FORMAT_BRUSH: OfficeIconGeometry = [
  {
    element: "path",
    role: "paper",
    d: "M14.8 3.2h5.6v4.3h-5.6zM16.2 7.5v2.2l-6.8 6.8-2.6-2.6 6.8-6.8",
  },
  {
    element: "path",
    role: "accent",
    emphasis: true,
    d: "m6.8 13.9-2.5 6.2 6.2-2.5",
  },
];

const FIND_LENS: OfficeIconGeometry = [
  { element: "circle", role: "paper", cx: 9.2, cy: 9.2, r: 5.4 },
  {
    element: "line",
    role: "accent",
    emphasis: true,
    x1: 13.1,
    y1: 13.1,
    x2: 20.4,
    y2: 20.4,
  },
  { element: "path", role: "ink-soft", d: "M6.6 8h5.2M6.6 10.7h3.5" },
];

function tableStyleGeometry(
  highlights: OfficeIconGeometry,
): OfficeIconGeometry {
  return [...TABLE_SURFACE, ...highlights, ...TABLE_GRID_LINES];
}

function paragraphGeometry(
  paths: string,
  accentPath: string,
): OfficeIconGeometry {
  return [
    { element: "path", role: "ink", d: paths },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: accentPath,
    },
  ];
}

/**
 * Original MRAI geometry authored on a 24px grid. Coordinates stay inside
 * the 20x20 optical safe area; no source asset, font glyph, or traced path is
 * embedded here.
 */
export const OFFICE_ICON_GEOMETRY = Object.freeze({
  // MRAI-internal derivatives adapted from the repository-owned Docs Insert
  // geometry at DocumentInsertRibbonIcons.tsx@ab80a559738c0f8a36053adf6304dc26899cdc8d:
  // page cover, blank page, page break, table, image, and shape.
  "page.cover.insert": [
    { element: "path", role: "paper", d: "M5.3 2.5h9.2l4.2 4.2v14.8H5.3z" },
    { element: "path", role: "ink-soft", d: "M14.5 2.5v4.2h4.2" },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 7.4,
      y: 9,
      width: 9.2,
      height: 3.2,
      rx: 0.7,
    },
    { element: "path", role: "accent", emphasis: true, d: "M7.4 12.2h9.2" },
    { element: "path", role: "ink-soft", d: "M7.8 15.2h8.4M7.8 18h5.4" },
  ],
  "page.blank.insert": [
    { element: "path", role: "paper", d: "M5.3 2.5h9.2l4.2 4.2v14.8H5.3z" },
    { element: "path", role: "ink-soft", d: "M14.5 2.5v4.2h4.2" },
    { element: "path", role: "accent", d: "M8 10.2h7.8M8 13.5h7.8M8 16.8h5" },
  ],
  "page.break.insert": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 2.7,
      width: 18,
      height: 6.2,
      rx: 0.8,
    },
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 15.1,
      width: 18,
      height: 6.2,
      rx: 0.8,
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M2.6 12h18.8",
      strokeDasharray: "2.2 2.2",
    },
    { element: "polyline", role: "accent", points: "9.7,9.8 12,12.2 14.3,9.8" },
  ],
  // Original page-orientation pair (not adapted from the Docs Insert geometry).
  // Portrait and Landscape are one 13 x 19 folded page given a quarter turn:
  // equal area, a tinted 4-unit flap, an accent heading and reading lines that
  // stay horizontal, so the pair reads as page orientation rather than as
  // rotated text. No emphasis strokes: menu surfaces that flatten stroke
  // weight render the pair exactly as drawn.
  "page.orientation.portrait": [
    { element: "path", role: "paper", d: "M5.5 2.5h9l4 4v15h-13z" },
    { element: "path", role: "accent-soft-fill", d: "M14.5 2.5v4h4z" },
    { element: "path", role: "ink-soft", d: "M14.5 2.5v4h4" },
    { element: "path", role: "accent", d: "M8 7h4.5" },
    { element: "path", role: "ink-soft", d: "M8 10.3h8M8 13.6h8M8 16.9h5.5" },
  ],
  "page.orientation.landscape": [
    { element: "path", role: "paper", d: "M2.5 5.5h15l4 4v9h-19z" },
    { element: "path", role: "accent-soft-fill", d: "M17.5 5.5v4h4z" },
    { element: "path", role: "ink-soft", d: "M17.5 5.5v4h4" },
    { element: "path", role: "accent", d: "M5.5 8.7h5.5" },
    { element: "path", role: "ink-soft", d: "M5.5 12h13M5.5 15.3h8.5" },
  ],
  "table.insert": [
    {
      element: "rect",
      role: "paper",
      x: 2.6,
      y: 4,
      width: 18.8,
      height: 16,
      rx: 1,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.1,
      y: 4.5,
      width: 17.8,
      height: 3.8,
      rx: 0.4,
    },
    { element: "path", role: "accent", emphasis: true, d: "M2.9 8.4h18.2" },
    {
      element: "path",
      role: "ink-soft",
      d: "M8.7 4.2v15.6M15.1 4.2v15.6M2.9 14h18.2",
    },
  ],
  // Editor-neutral table design and layout command faces. These are original
  // MRAI 24px constructions; table structure remains visible under each mark.
  "table.insertAbove": [
    ...COMPACT_TABLE,
    { element: "path", role: "accent", emphasis: true, d: "M3.2 2.7h13.6" },
    { element: "circle", role: "accent-soft-fill", cx: 20, cy: 5.1, r: 2.2 },
    { element: "path", role: "accent", d: "M20 3.8v2.6M18.7 5.1h2.6" },
  ],
  "table.insertBelow": [
    ...COMPACT_TABLE,
    { element: "path", role: "accent", emphasis: true, d: "M3.2 19.3h13.6" },
    { element: "circle", role: "accent-soft-fill", cx: 20, cy: 18.9, r: 2.2 },
    { element: "path", role: "accent", d: "M20 17.6v2.6M18.7 18.9h2.6" },
  ],
  "table.insertLeft": [
    ...COMPACT_TABLE,
    { element: "path", role: "accent", emphasis: true, d: "M1.8 3.8v13.9" },
    { element: "circle", role: "accent-soft-fill", cx: 4.5, cy: 21, r: 2.2 },
    { element: "path", role: "accent", d: "M4.5 19.7v2.6M3.2 21h2.6" },
  ],
  "table.insertRight": [
    ...COMPACT_TABLE,
    { element: "path", role: "accent", emphasis: true, d: "M18.2 3.8v13.9" },
    { element: "circle", role: "accent-soft-fill", cx: 20, cy: 20.2, r: 2.2 },
    { element: "path", role: "accent", d: "M20 18.9v2.6M18.7 20.2h2.6" },
  ],
  "table.select": [
    ...COMPACT_TABLE,
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 8.2,
      y: 8.5,
      width: 4.1,
      height: 4.4,
      rx: 0.35,
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M14.5 13.8v7.7l2.3-2.1 1.7 2.8 2-1.2-1.7-2.8 3.1-.5z",
    },
  ],
  "table.delete": [
    ...COMPACT_TABLE,
    { element: "circle", role: "warm-fill", cx: 18.3, cy: 18.2, r: 3.1 },
    {
      element: "path",
      role: "warm",
      emphasis: true,
      d: "m16.8 16.7 3 3m0-3-3 3",
    },
  ],
  "table.merge": [
    ...TABLE_SURFACE,
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.5,
      y: 9.5,
      width: 17,
      height: 5,
      rx: 0.35,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M3.5 9.3h17M3.5 14.7h17M9 4.5v4.8M15 4.5v4.8M9 14.7v4.8M15 14.7v4.8",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M5.4 12h4.3m-2-2 2.2 2-2.2 2M18.6 12h-4.3m2-2-2.2 2 2.2 2",
    },
  ],
  "table.split": [
    ...TABLE_SURFACE,
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.5,
      y: 9.5,
      width: 17,
      height: 5,
      rx: 0.35,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M3.5 9.3h17M3.5 14.7h17M9 4.5v4.8M15 4.5v4.8M9 14.7v4.8M15 14.7v4.8",
    },
    { element: "path", role: "accent", emphasis: true, d: "M12 9.5v5" },
    {
      element: "path",
      role: "accent",
      d: "M10 12H6m2-2-2.2 2L8 14M14 12h4m-2-2 2.2 2L16 14",
    },
  ],
  "table.rowHeight": [
    ...COMPACT_TABLE,
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.5,
      y: 8.5,
      width: 13,
      height: 4.4,
      rx: 0.3,
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M20 4v15.5M17.9 6.2 20 4l2.1 2.2M17.9 17.3l2.1 2.2 2.1-2.2",
    },
  ],
  "table.columnWidth": [
    {
      element: "rect",
      role: "paper",
      x: 4.5,
      y: 3,
      width: 15,
      height: 13.5,
      rx: 1,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 9.6,
      y: 3.5,
      width: 4.8,
      height: 12.5,
      rx: 0.3,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M9.5 3.5V16M14.5 3.5V16M5 8h14M5 12h14",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M4 20h16M6.2 17.9 4 20l2.2 2.1M17.8 17.9 20 20l-2.2 2.1",
    },
  ],
  "table.distributeRows": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 3.5,
      width: 14,
      height: 17,
      rx: 1,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M3.5 9.2h13M3.5 14.8h13M8 4v16M12.5 4v16",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M20 4.2v15.6M18.4 4.2h3.2M18.4 12h3.2M18.4 19.8h3.2",
    },
  ],
  "table.distributeColumns": [
    {
      element: "rect",
      role: "paper",
      x: 3.5,
      y: 3,
      width: 17,
      height: 14,
      rx: 1,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M9.2 3.5v13M14.8 3.5v13M4 8h16M4 12.5h16",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M4.2 20h15.6M4.2 18.4v3.2M12 18.4v3.2M19.8 18.4v3.2",
    },
  ],
  "table.sortAscending": [
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M5 3.8v16.4M2.8 18 5 20.2 7.2 18",
    },
    { element: "path", role: "ink", d: "M10 5.5h3.2M10 12h6.3M10 18.5h10.3" },
  ],
  "table.sortDescending": [
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M5 3.8v16.4M2.8 18 5 20.2 7.2 18",
    },
    { element: "path", role: "ink", d: "M10 5.5h10.3M10 12h6.3M10 18.5h3.2" },
  ],
  "table.pinRow": [
    ...COMPACT_TABLE,
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.5,
      y: 4,
      width: 13,
      height: 4.1,
      rx: 0.3,
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M16.2 4.5H21l-.8 3.1L22 9.4v1.4h-6.8V9.4L17 7.6zM18.6 10.8v6",
    },
  ],
  "table.unpinRow": [
    ...COMPACT_TABLE,
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.5,
      y: 4,
      width: 13,
      height: 4.1,
      rx: 0.3,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M16.2 4.5H21l-.8 3.1L22 9.4v1.4h-6.8V9.4L17 7.6zM18.6 10.8v6",
    },
    { element: "path", role: "warm", emphasis: true, d: "M15.1 3.8 22 10.7" },
  ],
  "table.altText": [
    ...COMPACT_TABLE,
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 12.2,
      y: 13,
      width: 9.2,
      height: 7.5,
      rx: 1,
    },
    {
      element: "rect",
      role: "accent",
      emphasis: true,
      x: 12.2,
      y: 13,
      width: 9.2,
      height: 7.5,
      rx: 1,
    },
    {
      element: "path",
      role: "accent",
      d: "M14.2 15.3h5.1M14.2 17.3h5.1M14.2 19.1h3.3",
    },
  ],
  "table.shading": [
    ...COMPACT_TABLE,
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 8.2,
      y: 8.5,
      width: 4.1,
      height: 4.4,
      rx: 0.3,
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "m15.2 8.4 3.1 3.1 2.7-2.7-3.1-3.1z",
    },
    {
      element: "path",
      role: "accent",
      d: "M20.1 14.3c0 1.3-1 2.3-2.2 2.3s-2.2-1-2.2-2.3c0-1.1 2.2-3.8 2.2-3.8s2.2 2.7 2.2 3.8z",
    },
  ],
  "table.borders": [
    ...TABLE_SURFACE,
    ...TABLE_GRID_LINES,
    {
      element: "rect",
      role: "accent",
      emphasis: true,
      x: 3,
      y: 4,
      width: 18,
      height: 16,
      rx: 1,
    },
  ],
  "table.cellMargins": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 3,
      width: 18,
      height: 18,
      rx: 1,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 7,
      y: 7,
      width: 10,
      height: 10,
      rx: 0.7,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M8.8 10h6.4M8.8 12.8h6.4M8.8 15.6h4.1",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M3 6h4M5 4v4M17 6h4M19 4v4M3 18h4M5 16v4M17 18h4M19 16v4",
    },
  ],
  "table.style.color": [
    ...COMPACT_TABLE,
    { element: "circle", role: "accent-fill", cx: 14.5, cy: 19.8, r: 1.5 },
    { element: "circle", role: "warm-fill", cx: 18.2, cy: 19.8, r: 1.5 },
    { element: "circle", role: "paper", cx: 21.3, cy: 17.4, r: 1.5 },
  ],
  "table.style.headerRow": tableStyleGeometry([
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.5,
      y: 4.5,
      width: 17,
      height: 4.6,
      rx: 0.3,
    },
    { element: "path", role: "accent", emphasis: true, d: "M3.5 9.3h17" },
  ]),
  "table.style.totalRow": tableStyleGeometry([
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.5,
      y: 14.8,
      width: 17,
      height: 4.7,
      rx: 0.3,
    },
    { element: "path", role: "accent", emphasis: true, d: "M3.5 14.7h17" },
  ]),
  "table.style.firstColumn": tableStyleGeometry([
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.5,
      y: 4.5,
      width: 5.3,
      height: 15,
      rx: 0.3,
    },
    { element: "path", role: "accent", emphasis: true, d: "M9 4.5v15" },
  ]),
  "table.style.lastColumn": tableStyleGeometry([
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 15.2,
      y: 4.5,
      width: 5.3,
      height: 15,
      rx: 0.3,
    },
    { element: "path", role: "accent", emphasis: true, d: "M15 4.5v15" },
  ]),
  "table.style.bandedRows": tableStyleGeometry([
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.5,
      y: 9.5,
      width: 17,
      height: 5,
      rx: 0.3,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.5,
      y: 17.2,
      width: 17,
      height: 2.3,
      rx: 0.3,
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M3.5 9.3h17M3.5 14.7h17",
    },
  ]),
  "table.style.bandedColumns": tableStyleGeometry([
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 9.2,
      y: 4.5,
      width: 5.6,
      height: 15,
      rx: 0.3,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 18,
      y: 4.5,
      width: 2.5,
      height: 15,
      rx: 0.3,
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M9 4.5v15M15 4.5v15",
    },
  ]),
  "table.gridlines": [
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.4,
      y: 4.4,
      width: 17.2,
      height: 15.2,
      rx: 0.8,
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      strokeDasharray: "1.8 1.5",
      d: "M3 4h18v16H3zM3 9.3h18M3 14.7h18M9 4v16M15 4v16",
    },
  ],
  // Shared table-cell text-flow command. Docs and Presentation can reuse the
  // geometry while retaining editor-specific command/model adapters.
  "table.cell.textDirection": [
    {
      element: "path",
      role: "ink",
      d: "M3.8 16.8 7.5 6h1.8L13 16.8M5.3 12.5h6.2",
    },
    { element: "path", role: "accent", emphasis: true, d: "M18 4.2v14.5" },
    {
      element: "polyline",
      role: "accent",
      emphasis: true,
      points: "14.9,15.6 18,19 21.1,15.6",
    },
  ],
  // Original, editor-neutral table command geometry. Product editors bind
  // their own command authority while sharing one compact ribbon language.
  "table.resize": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 3.5,
      width: 14.5,
      height: 14.5,
      rx: 1,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.5,
      y: 4,
      width: 13.5,
      height: 3.2,
      rx: 0.4,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M3.5 7.5H17M8 4v13.5M12.5 4v13.5M3.5 12h13.5",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M14.8 14.8 21 21M16.6 21H21v-4.4",
    },
  ],
  "table.slicer": [
    {
      element: "rect",
      role: "paper",
      x: 2.8,
      y: 3.5,
      width: 13.6,
      height: 16.8,
      rx: 1,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.3,
      y: 4,
      width: 12.6,
      height: 3.2,
      rx: 0.4,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M3.3 7.5h12.6M7.5 4v15.8M11.8 4v15.8M3.3 12h12.6M3.3 16h8.5",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M13.3 9.5h8l-3 3.5v5.2l-2.2 1.5V13z",
    },
  ],
  "table.removeDuplicates": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 5.5,
      width: 12.5,
      height: 14.5,
      rx: 1,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 7,
      y: 3,
      width: 12.5,
      height: 14.5,
      rx: 1,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M7.5 7h11.5M11.5 3.5V17M15.5 3.5V14M3.5 10h8M3.5 14h8",
    },
    { element: "circle", role: "warm-fill", cx: 18.3, cy: 18.2, r: 3.1 },
    {
      element: "path",
      role: "warm",
      emphasis: true,
      d: "m16.8 16.7 3 3m0-3-3 3",
    },
  ],
  "table.convertRange": [
    {
      element: "rect",
      role: "paper",
      x: 2.8,
      y: 3.5,
      width: 11.8,
      height: 14.2,
      rx: 1,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.3,
      y: 4,
      width: 10.8,
      height: 3.1,
      rx: 0.4,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M3.3 7.4h10.8M7 4v13.2M10.7 4v13.2M3.3 12h10.8",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M13.2 9.2c4.7.2 7.2 2.4 7.2 6.4m-2.8-2.5 2.8 2.8 2.3-3",
    },
    { element: "path", role: "ink-soft", d: "M15.8 18.3h6M18.8 15.3v6" },
  ],
  "data.refreshAll": [
    {
      element: "rect",
      role: "paper",
      x: 6.2,
      y: 6.2,
      width: 11.6,
      height: 11.6,
      rx: 1,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M6.7 10h10.6M10 6.7v10.6M14 6.7v10.6M6.7 14h10.6",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M4.2 10.3A8.2 8.2 0 0 1 18.5 5L21 7.4M19.8 4.3 21 7.4l-3.3.5M19.8 13.7A8.2 8.2 0 0 1 5.5 19L3 16.6M4.2 19.7 3 16.6l3.3-.5",
    },
  ],
  "image.insert": [
    {
      element: "rect",
      role: "paper",
      x: 2.8,
      y: 4.1,
      width: 18.4,
      height: 15.8,
      rx: 1.5,
    },
    { element: "circle", role: "ink-soft", cx: 16.6, cy: 8.3, r: 1.5 },
    {
      element: "polygon",
      role: "accent-soft-fill",
      points: "3.8,18.4 9.2,12.4 12.2,15.3 14.8,13.1 20.2,18.4",
    },
    {
      element: "polyline",
      role: "accent",
      emphasis: true,
      points: "3.8,18.4 9.2,12.4 12.2,15.3 14.8,13.1 20.2,18.4",
    },
  ],
  // Cross-editor image-format command faces. The landscape motif is shared,
  // while each overlay names one precise operation.
  "image.edit": [
    ...IMAGE_BASE,
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "m13.8 19.8 1-3.4 4.8-4.8 2.4 2.4-4.8 4.8zM18.5 12.7l2.4 2.4",
    },
  ],
  "image.removeBackground": [
    ...IMAGE_BASE,
    {
      element: "path",
      role: "warm",
      emphasis: true,
      strokeDasharray: "1.7 1.4",
      d: "M3.2 4.2h6.3v13.6H3.2z",
    },
    {
      element: "path",
      role: "warm",
      d: "m4.4 6 3.9 3.9M4.4 10.1l3.9 3.9M4.4 14.2l2.5 2.5",
    },
  ],
  "image.enhance": [
    ...IMAGE_BASE,
    {
      element: "path",
      role: "warm",
      emphasis: true,
      d: "M20 3v4M18 5h4M17.4 8.2v2.6M16.1 9.5h2.6",
    },
  ],
  "image.change": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 4,
      width: 11,
      height: 10,
      rx: 1,
    },
    { element: "circle", role: "ink-soft", cx: 10.7, cy: 7.2, r: 1 },
    {
      element: "polyline",
      role: "ink-soft",
      points: "4,12.5 7.2,9.3 9.2,11.2 11,9.8 13,12.5",
    },
    {
      element: "rect",
      role: "paper",
      x: 10,
      y: 10,
      width: 11,
      height: 10,
      rx: 1,
    },
    {
      element: "polyline",
      role: "ink-soft",
      points: "11,18.5 14.3,15.3 16.1,17.1 18,15.4 20,18.5",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M14.5 4.7h6.2m-2-2 2.2 2-2.2 2M9.5 19.3H3.3m2 2-2.2-2 2.2-2",
    },
  ],
  "image.reset": [
    ...IMAGE_BASE,
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M20.7 12.8a7.2 7.2 0 0 1-7.1 7.9 7.3 7.3 0 0 1-5.3-2.3M20.4 9v4h-4",
    },
  ],
  "image.corrections": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 4,
      width: 18,
      height: 16,
      rx: 1.3,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.5,
      y: 4.5,
      width: 8.5,
      height: 15,
      rx: 0.6,
    },
    { element: "path", role: "ink-soft", d: "M12 4.5v15" },
    { element: "circle", role: "accent", cx: 8, cy: 9.2, r: 2 },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M8 5v1.4M8 12v1.4M3.8 9.2h1.4M10.8 9.2h1.4",
    },
    {
      element: "polyline",
      role: "ink-soft",
      points: "13.2,17.8 16,14.8 18.2,16.6 20,14.7",
    },
  ],
  "image.color": [
    ...IMAGE_BASE,
    { element: "circle", role: "accent-fill", cx: 14.8, cy: 20.5, r: 1.4 },
    { element: "circle", role: "warm-fill", cx: 18.3, cy: 20.5, r: 1.4 },
    { element: "circle", role: "paper", cx: 21.2, cy: 18.3, r: 1.4 },
  ],
  "image.transparency": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 4,
      width: 18,
      height: 16,
      rx: 1.3,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M3.5 9.3h17M3.5 14.7h17M9.3 4.5v15M14.7 4.5v15",
    },
    {
      element: "path",
      role: "accent-soft-fill",
      d: "M3.5 4.5h5.8v4.8H3.5zM9.3 9.3h5.4v5.4H9.3zM14.7 14.7h5.8v4.8h-5.8z",
    },
    {
      element: "rect",
      role: "accent",
      emphasis: true,
      x: 3,
      y: 4,
      width: 18,
      height: 16,
      rx: 1.3,
    },
  ],
  "image.border": [
    ...IMAGE_BASE,
    {
      element: "rect",
      role: "accent",
      emphasis: true,
      x: 2.5,
      y: 3.5,
      width: 17,
      height: 15,
      rx: 1.6,
    },
  ],
  "image.crop": [
    ...IMAGE_BASE,
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M7 2.5v15.8a2 2 0 0 0 2 2h12.5M2.5 7H18a2 2 0 0 1 2 2v12.5",
    },
  ],
  "image.format": [
    ...IMAGE_BASE,
    ...FORMAT_BRUSH,
    { element: "path", role: "warm", d: "M4.4 20.8h7" },
  ],
  "image.compress": [
    ...IMAGE_BASE,
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M22 21.6l-4.4-4.4M17.6 20.9v-3.7h3.7",
    },
    { element: "path", role: "ink-soft", d: "M21.6 9.4v3.4M9.6 21.2h3.4" },
  ],
  "shape.insert": [
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3,
      y: 3.4,
      width: 10.2,
      height: 10.2,
      rx: 1.5,
    },
    {
      element: "rect",
      role: "accent",
      emphasis: true,
      x: 3,
      y: 3.4,
      width: 10.2,
      height: 10.2,
      rx: 1.5,
    },
    { element: "circle", role: "warm-fill", cx: 15.8, cy: 15.5, r: 5.2 },
    {
      element: "path",
      role: "warm",
      emphasis: true,
      d: "M15.8 10.3a5.2 5.2 0 1 1-3.7 8.9",
    },
  ],
  "shape.change": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 4,
      width: 8.5,
      height: 8.5,
      rx: 1.4,
    },
    { element: "circle", role: "paper", cx: 16.8, cy: 16.6, r: 4.2 },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M13.2 4.8h7.1m-2.2-2.1 2.3 2.1-2.3 2.1M10.8 19.2H3.7m2.2 2.1-2.3-2.1 2.3-2.1",
    },
  ],
  "shape.editPoints": [
    {
      element: "path",
      role: "ink",
      d: "M4.4 16.6C3.7 11.2 6.9 6 11.9 5.3c4.6-.6 8.4 2.4 8.1 6.6-.3 4.4-3.9 7.6-8.4 7.7-3.5.1-6.8-.3-7.2-3z",
    },
    { element: "path", role: "warm", d: "M8 6.7 15.8 3.9" },
    { element: "circle", role: "warm-fill", cx: 8, cy: 6.7, r: 1.2 },
    { element: "circle", role: "warm-fill", cx: 15.8, cy: 3.9, r: 1.2 },
    { element: "rect", role: "accent-fill", emphasis: true, x: 10.4, y: 3.8, width: 3, height: 3, rx: 0.5 },
    { element: "rect", role: "paper", x: 2.9, y: 15.1, width: 3, height: 3, rx: 0.5 },
    { element: "rect", role: "paper", x: 18.5, y: 10.4, width: 3, height: 3, rx: 0.5 },
  ],
  "connector.reroute": [
    { element: "rect", role: "paper", x: 2.5, y: 3, width: 7.5, height: 5.5, rx: 1.2 },
    { element: "rect", role: "paper", x: 14, y: 15.5, width: 7.5, height: 5.5, rx: 1.2 },
    { element: "path", role: "ink-soft", strokeDasharray: "1.6 1.6", d: "M6.25 8.5v9.75H14" },
    { element: "path", role: "accent", emphasis: true, d: "M10 5.75h2.25v12.5H14" },
    { element: "circle", role: "accent-fill", cx: 10, cy: 5.75, r: 1.1 },
    { element: "circle", role: "accent-fill", cx: 14, cy: 18.25, r: 1.1 },
  ],
  "table.align.left": [
    { element: "path", role: "ink-soft", strokeDasharray: "1.5 1.5", d: "M3.5 3.5v17M20.5 3.5v17" },
    { element: "rect", role: "paper", x: 4.5, y: 7, width: 9, height: 10, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M4.5 10.33h9M4.5 13.67h9M9 7v10" },
  ],
  "table.align.center": [
    { element: "path", role: "ink-soft", strokeDasharray: "1.5 1.5", d: "M3.5 3.5v17M20.5 3.5v17" },
    { element: "rect", role: "paper", x: 7.5, y: 7, width: 9, height: 10, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M7.5 10.33h9M7.5 13.67h9M12 7v10" },
  ],
  "table.align.right": [
    { element: "path", role: "ink-soft", strokeDasharray: "1.5 1.5", d: "M3.5 3.5v17M20.5 3.5v17" },
    { element: "rect", role: "paper", x: 10.5, y: 7, width: 9, height: 10, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M10.5 10.33h9M10.5 13.67h9M15 7v10" },
  ],
  "table.wrap.none": [
    { element: "path", role: "ink-soft", d: "M3.5 4h17M3.5 6.5h17" },
    { element: "rect", role: "paper", x: 6, y: 9, width: 12, height: 6.5, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M6 12.25h12M12 9v6.5" },
    { element: "path", role: "ink-soft", d: "M3.5 18h17M3.5 20.5h11" },
  ],
  "table.wrap.around": [
    { element: "path", role: "ink-soft", d: "M3.5 4h17" },
    { element: "rect", role: "paper", x: 3.5, y: 7, width: 9, height: 10, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M3.5 10.33h9M3.5 13.67h9M8 7v10" },
    { element: "path", role: "ink-soft", d: "M15 7.5h5.5M15 10.5h5.5M15 13.5h5.5M15 16.5h5.5M3.5 20h17" },
  ],
  "table.cell.alignTop": [
    { element: "rect", role: "paper", x: 4, y: 4, width: 16, height: 16, rx: 1.5 },
    { element: "path", role: "ink-soft", strokeDasharray: "1.5 1.5", d: "M4 12h16" },
    { element: "path", role: "accent", emphasis: true, d: "M7.5 7.5h9M7.5 10h6" },
  ],
  "table.cell.alignMiddle": [
    { element: "rect", role: "paper", x: 4, y: 4, width: 16, height: 16, rx: 1.5 },
    { element: "path", role: "accent", emphasis: true, d: "M7.5 10.75h9M7.5 13.25h6" },
  ],
  "table.cell.alignBottom": [
    { element: "rect", role: "paper", x: 4, y: 4, width: 16, height: 16, rx: 1.5 },
    { element: "path", role: "ink-soft", strokeDasharray: "1.5 1.5", d: "M4 12h16" },
    { element: "path", role: "accent", emphasis: true, d: "M7.5 14h9M7.5 16.5h6" },
  ],
  "page.border.none": [
    { element: "rect", role: "paper", x: 5, y: 3, width: 14, height: 18, rx: 1 },
    { element: "path", role: "ink-soft", d: "M8 7.5h8M8 10.5h8M8 13.5h8M8 16.5h5" },
  ],
  "page.border.box": [
    { element: "rect", role: "paper", x: 5, y: 3, width: 14, height: 18, rx: 1 },
    { element: "rect", role: "accent", emphasis: true, x: 6.75, y: 4.75, width: 10.5, height: 14.5, rx: 0.5 },
    { element: "path", role: "ink-soft", d: "M9 8h6M9 11h6M9 14h6M9 17h4" },
  ],
  "page.border.shadow": [
    { element: "rect", role: "paper", x: 5, y: 3, width: 14.5, height: 18.5, rx: 1 },
    { element: "path", role: "ink", emphasis: true, d: "M18.25 6v14.25h-10.5" },
    { element: "rect", role: "accent", emphasis: true, x: 6.75, y: 4.75, width: 10.5, height: 14.5, rx: 0.5 },
    { element: "path", role: "ink-soft", d: "M9 8h6M9 11h6M9 14h6M9 17h4" },
  ],
  "page.border.threeD": [
    { element: "rect", role: "paper", x: 5, y: 3, width: 14.5, height: 18.5, rx: 1 },
    { element: "polygon", role: "accent-soft-fill", points: "17.25,4.75 18.5,6 18.5,20.5 8,20.5 6.75,19.25 17.25,19.25" },
    { element: "path", role: "ink-soft", d: "M17.25 4.75l1.25 1.25v14.5h-10.5l-1.25-1.25" },
    { element: "rect", role: "accent", emphasis: true, x: 6.75, y: 4.75, width: 10.5, height: 14.5, rx: 0.5 },
    { element: "path", role: "ink-soft", d: "M9 8h6M9 11h6M9 14h6M9 17h4" },
  ],
  "page.border.custom": [
    { element: "rect", role: "paper", x: 5, y: 3, width: 14, height: 18, rx: 1 },
    { element: "path", role: "ink-soft", strokeDasharray: "1.5 1.5", d: "M17.25 4.75v14.5h-10.5" },
    { element: "path", role: "accent", emphasis: true, d: "M6.75 19.25V4.75h10.5" },
    { element: "path", role: "ink-soft", d: "M9 8h6M9 11h6M9 14h6M9 17h4" },
  ],
  "paragraph.tab.left": [
    { element: "path", role: "ink-soft", d: "M3.5 20h17" },
    { element: "path", role: "ink", d: "M9 6h9.5M9 9h6" },
    { element: "path", role: "accent", emphasis: true, d: "M9 11.5v6h5" },
  ],
  "paragraph.tab.center": [
    { element: "path", role: "ink-soft", d: "M3.5 20h17" },
    { element: "path", role: "ink", d: "M7 6h10M9 9h6" },
    { element: "path", role: "accent", emphasis: true, d: "M12 11.5v6M8 17.5h8" },
  ],
  "paragraph.tab.right": [
    { element: "path", role: "ink-soft", d: "M3.5 20h17" },
    { element: "path", role: "ink", d: "M5.5 6H15M9 9h6" },
    { element: "path", role: "accent", emphasis: true, d: "M15 11.5v6h-5" },
  ],
  "paragraph.tab.decimal": [
    { element: "path", role: "ink-soft", d: "M3.5 20h17" },
    { element: "path", role: "ink", d: "M6 6h4.5M13.5 6h3.5" },
    { element: "circle", role: "ink", cx: 12, cy: 6.5, r: 0.6 },
    { element: "path", role: "accent", emphasis: true, d: "M12 11.5v6M8 17.5h8" },
    { element: "circle", role: "accent-fill", cx: 14.75, cy: 14, r: 0.9 },
  ],
  "paragraph.tab.bar": [
    { element: "path", role: "ink-soft", d: "M3.5 20h17" },
    { element: "path", role: "ink", d: "M4.5 7h5.5M14 7h5.5M4.5 11h4M14 11h4" },
    { element: "path", role: "accent", emphasis: true, d: "M12 4v13.5" },
  ],
  "text.separator.tab": [
    { element: "path", role: "ink-soft", d: "M3.5 7h4M3.5 17h4M16.5 7h4M16.5 17h4" },
    { element: "path", role: "accent", emphasis: true, d: "M5 12h10.5M12.5 9l3 3-3 3M18.5 8.5v7" },
  ],
  "text.separator.comma": [
    { element: "path", role: "ink-soft", d: "M3.5 12h5.5M15 12h5.5" },
    { element: "circle", role: "accent-fill", cx: 12, cy: 12, r: 1.6 },
    { element: "path", role: "accent", emphasis: true, d: "M13.5 12.5c0 2.3-1 3.9-2.75 4.75" },
  ],
  "ink.toShape": [
    { element: "path", role: "ink", d: "M3.6 4.6c2.8-.9 5.6-.3 8.4-.6.4 2.7-.2 5.5.3 8.2-2.9.5-5.8-.3-8.6.3-.5-2.6.1-5.3-.1-7.9z" },
    { element: "rect", role: "paper", x: 10.4, y: 10.4, width: 10.8, height: 10.8, rx: 2.1 },
    { element: "rect", role: "accent", emphasis: true, x: 10.4, y: 10.4, width: 10.8, height: 10.8, rx: 2.1 },
    { element: "path", role: "warm", d: "M19.4 2.8v3M17.9 4.3h3" },
  ],
  "ink.toMath": [
    { element: "path", role: "ink", d: "M2.6 18.2c1.4-1.3 2.7-1.5 3.7-.6 1 1 1.8-3.2 2.4-7.3.6-3.8 1.9-6.1 3.7-6.6" },
    { element: "path", role: "accent", emphasis: true, d: "M21.4 9.2h-8.6l4.3 4.4-4.3 4.4h8.6" },
    { element: "path", role: "warm", d: "M17.2 2.3v3.2M15.6 3.9h3.2" },
  ],
  "ink.toText": [
    { element: "path", role: "ink", d: "M2.4 19.6c1-1.9 1.8-2.6 2.3-1.4.5 1.1 1 1.3 1.7-.3.6-1.4 1.2-1.5 1.6-.1.4 1.2 1.1 1.4 2.4.4" },
    { element: "path", role: "accent", emphasis: true, d: "M10.6 13.4 14 4.6l3.4 8.8M11.9 10.2h4.2" },
    { element: "path", role: "ink", d: "M20.4 7.8v6" },
    { element: "path", role: "warm", d: "M19.4 2.8v3M17.9 4.3h3" },
  ],
  "shape.fill": [
    ...SHAPE_BASE,
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "m15.2 7 3.2 3.2 2.8-2.8L18 4.2z",
    },
    {
      element: "path",
      role: "accent",
      d: "M21 14.1c0 1.4-1 2.4-2.3 2.4s-2.3-1-2.3-2.4c0-1.2 2.3-4 2.3-4s2.3 2.8 2.3 4z",
    },
  ],
  "shape.outline": [
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 4.2,
      y: 4.2,
      width: 15.6,
      height: 15.6,
      rx: 3,
    },
    {
      element: "rect",
      role: "accent",
      emphasis: true,
      x: 3,
      y: 3,
      width: 18,
      height: 18,
      rx: 3.4,
    },
    { element: "circle", role: "ink-soft", cx: 12, cy: 12, r: 4.2 },
  ],
  "shape.effects": [
    {
      element: "rect",
      role: "warm-fill",
      x: 6.3,
      y: 6.5,
      width: 13,
      height: 13,
      rx: 2.4,
    },
    {
      element: "rect",
      role: "accent",
      emphasis: true,
      x: 3.5,
      y: 3.7,
      width: 13,
      height: 13,
      rx: 2.4,
    },
    {
      element: "path",
      role: "warm",
      emphasis: true,
      d: "M19.7 3v4M17.7 5h4M18.9 8.7v2.6M17.6 10h2.6",
    },
  ],
  "object.altText": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 4,
      width: 12,
      height: 12,
      rx: 1.5,
    },
    { element: "circle", role: "ink-soft", cx: 11.5, cy: 7.5, r: 1.1 },
    {
      element: "polyline",
      role: "ink-soft",
      points: "4.2,14.5 7.4,11 9.4,12.7 11,11.3 13.8,14.5",
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 10.5,
      y: 12.5,
      width: 10.5,
      height: 7.5,
      rx: 1,
    },
    {
      element: "rect",
      role: "accent",
      emphasis: true,
      x: 10.5,
      y: 12.5,
      width: 10.5,
      height: 7.5,
      rx: 1,
    },
    {
      element: "path",
      role: "accent",
      d: "M12.5 15h6.5M12.5 17.2H19M12.5 19.2h4",
    },
  ],
  "object.bringForward": [
    {
      element: "rect",
      role: "ink-soft",
      x: 4,
      y: 9,
      width: 9,
      height: 9,
      rx: 1.3,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 9,
      y: 4,
      width: 9,
      height: 9,
      rx: 1.3,
    },
    {
      element: "rect",
      role: "accent",
      emphasis: true,
      x: 9,
      y: 4,
      width: 9,
      height: 9,
      rx: 1.3,
    },
    {
      element: "path",
      role: "accent",
      d: "M18.5 19.8v-5.5m-2.2 2.1 2.2-2.2 2.2 2.2",
    },
  ],
  "object.bringToFront": [
    {
      element: "rect",
      role: "ink-soft",
      x: 4,
      y: 9,
      width: 9,
      height: 9,
      rx: 1.3,
    },
    {
      element: "rect",
      role: "accent",
      emphasis: true,
      x: 9,
      y: 4,
      width: 9,
      height: 9,
      rx: 1.3,
    },
    {
      element: "path",
      role: "accent",
      d: "M20.2 18.5V8.8m-2.2 2.1 2.2-2.2 2.2 2.2M17.5 5.2h5.3",
    },
  ],
  "object.sendBackward": [
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 4,
      y: 9,
      width: 9,
      height: 9,
      rx: 1.3,
    },
    {
      element: "rect",
      role: "accent",
      emphasis: true,
      x: 4,
      y: 9,
      width: 9,
      height: 9,
      rx: 1.3,
    },
    {
      element: "rect",
      role: "ink-soft",
      x: 9,
      y: 4,
      width: 9,
      height: 9,
      rx: 1.3,
    },
    {
      element: "path",
      role: "accent",
      d: "M18.5 14.2v5.6m-2.2-2.2 2.2 2.2 2.2-2.2",
    },
  ],
  "object.sendToBack": [
    {
      element: "rect",
      role: "accent",
      emphasis: true,
      x: 4,
      y: 9,
      width: 9,
      height: 9,
      rx: 1.3,
    },
    {
      element: "rect",
      role: "ink-soft",
      x: 9,
      y: 4,
      width: 9,
      height: 9,
      rx: 1.3,
    },
    {
      element: "path",
      role: "accent",
      d: "M20.2 8.8v9.7M18 16.4l2.2 2.2 2.2-2.2M17.5 21h5.3",
    },
  ],
  "object.selectionPane": [
    { element: "rect", role: "paper", x: 3, y: 4, width: 11, height: 7, rx: 1 },
    { element: "circle", role: "accent-soft-fill", cx: 10.5, cy: 14.5, r: 4.5 },
    { element: "circle", role: "accent", cx: 10.5, cy: 14.5, r: 4.5 },
    {
      element: "rect",
      role: "paper",
      x: 15.5,
      y: 3.5,
      width: 6,
      height: 17,
      rx: 1,
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M17.2 7h2.6M17.2 11h2.6M17.2 15h2.6",
    },
  ],
  "object.group": [
    {
      element: "rect",
      role: "paper",
      x: 5,
      y: 6,
      width: 6.5,
      height: 6.5,
      rx: 1,
    },
    { element: "circle", role: "accent-soft-fill", cx: 15.5, cy: 14.8, r: 3.5 },
    {
      element: "rect",
      role: "accent",
      emphasis: true,
      strokeDasharray: "2 1.6",
      x: 2.8,
      y: 3.5,
      width: 18.4,
      height: 17,
      rx: 1,
    },
  ],
  "object.ungroup": [
    {
      element: "rect",
      role: "paper",
      x: 3.5,
      y: 4,
      width: 8,
      height: 8,
      rx: 1,
    },
    {
      element: "rect",
      role: "accent",
      emphasis: true,
      strokeDasharray: "1.8 1.5",
      x: 2.5,
      y: 3,
      width: 10,
      height: 10,
      rx: 1,
    },
    { element: "circle", role: "accent-soft-fill", cx: 16.2, cy: 15.8, r: 4 },
    {
      element: "circle",
      role: "accent",
      emphasis: true,
      cx: 16.2,
      cy: 15.8,
      r: 5,
    },
  ],
  "object.align": [
    { element: "path", role: "accent", emphasis: true, d: "M4 4v16" },
    {
      element: "rect",
      role: "paper",
      x: 6.5,
      y: 5,
      width: 11,
      height: 5,
      rx: 0.8,
    },
    {
      element: "rect",
      role: "paper",
      x: 6.5,
      y: 13.5,
      width: 15,
      height: 5,
      rx: 0.8,
    },
    { element: "path", role: "accent", d: "M4 7.5h2.5M4 16h2.5" },
  ],
  "object.rotate": [
    {
      element: "rect",
      role: "paper",
      x: 6,
      y: 8.15,
      width: 12,
      height: 12,
      rx: 1.5,
    },
    { element: "circle", role: "accent-fill", cx: 12, cy: 5.15, r: 1.3 },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M5 10.35a8.5 8.5 0 0 1 14.2-2.5L21 9.65M20.3 6.25l.7 3.4-3.4.2",
    },
  ],
  "object.flipHorizontal": [
    {
      element: "path",
      role: "accent",
      emphasis: true,
      strokeDasharray: "2 1.6",
      d: "M12 3v18",
    },
    { element: "polygon", role: "paper", points: "10,5 4,9 10,19" },
    { element: "polygon", role: "accent-soft-fill", points: "14,5 20,9 14,19" },
    { element: "polygon", role: "accent", points: "14,5 20,9 14,19" },
  ],
  "object.flipVertical": [
    {
      element: "path",
      role: "accent",
      emphasis: true,
      strokeDasharray: "2 1.6",
      d: "M3 12h18",
    },
    { element: "polygon", role: "paper", points: "5,10 9,4 19,10" },
    { element: "polygon", role: "accent-soft-fill", points: "5,14 9,20 19,14" },
    { element: "polygon", role: "accent", points: "5,14 9,20 19,14" },
  ],
  "object.height": [
    {
      element: "rect",
      role: "paper",
      x: 3.5,
      y: 5,
      width: 12.5,
      height: 14,
      rx: 1.5,
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M20 4v16M17.8 6.2 20 4l2.2 2.2M17.8 17.8 20 20l2.2-2.2",
    },
  ],
  "object.width": [
    {
      element: "rect",
      role: "paper",
      x: 5,
      y: 3.5,
      width: 14,
      height: 12.5,
      rx: 1.5,
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M4 20h16M6.2 17.8 4 20l2.2 2.2M17.8 17.8 20 20l-2.2 2.2",
    },
  ],
  "object.lockAspectRatio": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 4,
      width: 13,
      height: 16,
      rx: 1.5,
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M18 13.5v-2a2.7 2.7 0 0 1 5.4 0v2",
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 15.5,
      y: 13.5,
      width: 7,
      height: 6,
      rx: 1.2,
    },
    {
      element: "rect",
      role: "accent",
      x: 15.5,
      y: 13.5,
      width: 7,
      height: 6,
      rx: 1.2,
    },
    { element: "path", role: "ink-soft", d: "M5.5 7h8M5.5 17h6M6 4v3M13 17v3" },
  ],
  // Independently authored for this shared package; no external ancestry.
  "icon.insert": [
    {
      element: "rect",
      role: "accent",
      emphasis: true,
      x: 3,
      y: 3,
      width: 8,
      height: 8,
      rx: 2,
    },
    { element: "circle", role: "ink", cx: 17.3, cy: 6.8, r: 3.3 },
    {
      element: "polygon",
      role: "ink-soft",
      points: "3.7,20.5 7.1,14.2 10.5,20.5",
    },
    {
      element: "path",
      role: "ink",
      d: "M15.1 15.1h5.7v5.7h-5.7zM17.95 16.5v2.9M16.5 17.95h2.9",
    },
  ],
  // Independently authored for this shared package; no external ancestry.
  "diagram.smartArt.insert": [
    {
      element: "rect",
      role: "paper",
      x: 2.7,
      y: 3.8,
      width: 6.5,
      height: 4.7,
      rx: 0.9,
    },
    {
      element: "rect",
      role: "paper",
      x: 14.8,
      y: 3.8,
      width: 6.5,
      height: 4.7,
      rx: 0.9,
    },
    {
      element: "rect",
      role: "paper",
      x: 8.8,
      y: 15.3,
      width: 6.5,
      height: 4.7,
      rx: 0.9,
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M9.5 6.15h4.9m-1.8-1.8 1.8 1.8-1.8 1.8M18 8.8v1.7a5.7 5.7 0 0 1-5.7 5.7h-3",
    },
  ],
  // MRAI-internal derivatives from the same approved Docs Insert source:
  // chart, link, comment, and text box.
  "chart.insert": [
    { element: "path", role: "ink", d: "M3 3v18h18" },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 5.8,
      y: 13,
      width: 3.2,
      height: 7.4,
      rx: 0.5,
    },
    {
      element: "rect",
      role: "warm-fill",
      x: 10.6,
      y: 8.8,
      width: 3.2,
      height: 11.6,
      rx: 0.5,
    },
    {
      element: "rect",
      role: "accent-fill",
      x: 15.4,
      y: 4.6,
      width: 3.2,
      height: 15.8,
      rx: 0.5,
    },
  ],
  // Original shared Chart-tab geometry. These are visual semantics only;
  // chart-family support and save authority remain editor-owned adapters.
  "chart.switchRowColumn": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 3.5,
      width: 12.5,
      height: 12.5,
      rx: 1,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.5,
      y: 4,
      width: 11.5,
      height: 3,
      rx: 0.4,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M3.5 7.4H15M7.3 4v11.5M11.2 4v11.5M3.5 11.2H15",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M16.7 7.2h4.1l-1.7-1.7M20.8 7.2l-1.7 1.7M19.7 16.8v4.1l1.7-1.7M19.7 20.9 18 19.2",
    },
  ],
  "chart.selectData": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 3.5,
      width: 15.5,
      height: 15.5,
      rx: 1,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 7.2,
      y: 7.7,
      width: 7,
      height: 7,
      rx: 0.5,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M3.5 7.5H18M3.5 11.5H18M3.5 15.5H18M7.2 4v14.5M11.2 4v14.5M15.2 4v14.5",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M14.8 14.8v7l2.2-2 1.6 2.7 2-1.1-1.6-2.7 3-.5z",
    },
  ],
  "chart.type.column": [
    { element: "path", role: "ink", d: "M3 3.2v17.8h18" },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 5.5,
      y: 12.5,
      width: 3.3,
      height: 7.8,
      rx: 0.5,
    },
    {
      element: "rect",
      role: "accent-fill",
      x: 10.4,
      y: 8.3,
      width: 3.3,
      height: 12,
      rx: 0.5,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 15.3,
      y: 5,
      width: 3.3,
      height: 15.3,
      rx: 0.5,
    },
  ],
  "chart.type.line": [
    { element: "path", role: "ink", d: "M3 3.2v17.8h18" },
    {
      element: "polyline",
      role: "accent",
      emphasis: true,
      points: "4.5,16.5 8.2,11 12.1,14.3 16.1,7 20,9.5",
    },
    {
      element: "polyline",
      role: "warm",
      emphasis: true,
      points: "4.5,11.8 8.2,14.8 12.1,9.8 16.1,12.2 20,5.3",
    },
  ],
  "chart.type.scatter": [
    { element: "path", role: "ink", d: "M3 3.2v17.8h18" },
    { element: "circle", role: "accent-fill", cx: 6.1, cy: 15.2, r: 1.2 },
    { element: "circle", role: "accent-fill", cx: 9, cy: 11.3, r: 1.2 },
    { element: "circle", role: "accent-fill", cx: 12.4, cy: 13, r: 1.2 },
    { element: "circle", role: "accent-fill", cx: 15.2, cy: 7.6, r: 1.2 },
    { element: "circle", role: "accent-fill", cx: 18.5, cy: 5.2, r: 1.2 },
  ],
  "chart.type.pie": [
    {
      element: "path",
      role: "accent-soft-fill",
      d: "M11.5 3.2a8.8 8.8 0 1 0 8.8 8.8h-8.8z",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M13.2 2.8v7.7h7.7a7.7 7.7 0 0 0-7.7-7.7Z",
    },
    {
      element: "path",
      role: "warm-fill",
      d: "M13.2 12.2h8.2a8.2 8.2 0 0 1-3.8 6.9z",
    },
  ],
  "chart.type.bar": [
    { element: "path", role: "ink", d: "M3.2 3v18h17.8" },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 4,
      y: 5.2,
      width: 10,
      height: 3.2,
      rx: 0.5,
    },
    {
      element: "rect",
      role: "accent-fill",
      x: 4,
      y: 10.4,
      width: 15.8,
      height: 3.2,
      rx: 0.5,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 4,
      y: 15.6,
      width: 12.7,
      height: 3.2,
      rx: 0.5,
    },
  ],
  "chart.type.statistical": [
    { element: "path", role: "ink", d: "M3 3.2v17.8h18" },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 5,
      y: 14.5,
      width: 2.5,
      height: 5.8,
      rx: 0.3,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 7.8,
      y: 10.5,
      width: 2.5,
      height: 9.8,
      rx: 0.3,
    },
    {
      element: "rect",
      role: "accent-fill",
      x: 10.6,
      y: 7.2,
      width: 2.5,
      height: 13.1,
      rx: 0.3,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 13.4,
      y: 10.5,
      width: 2.5,
      height: 9.8,
      rx: 0.3,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 16.2,
      y: 14.5,
      width: 2.5,
      height: 5.8,
      rx: 0.3,
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M4.5 15.8c2.3 0 3.1-9.4 7.5-9.4s5.2 9.4 7.5 9.4",
    },
  ],
  "chart.type.combo": [
    { element: "path", role: "ink", d: "M3 3.2v17.8h18" },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 5.4,
      y: 13,
      width: 3,
      height: 7.3,
      rx: 0.4,
    },
    {
      element: "rect",
      role: "accent-fill",
      x: 10.5,
      y: 9.2,
      width: 3,
      height: 11.1,
      rx: 0.4,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 15.6,
      y: 6.2,
      width: 3,
      height: 14.1,
      rx: 0.4,
    },
    {
      element: "polyline",
      role: "warm",
      emphasis: true,
      points: "4.6,10.8 8.5,7.2 12.3,11.2 16.2,4.3 20,6.5",
    },
  ],
  "chart.type.other": [
    { element: "rect", role: "paper", x: 3, y: 3, width: 8, height: 8, rx: 1 },
    { element: "path", role: "accent", d: "M4.7 9.1 6.5 6.8 8.2 8 9.5 5" },
    { element: "rect", role: "paper", x: 13, y: 3, width: 8, height: 8, rx: 1 },
    {
      element: "path",
      role: "accent",
      d: "M14.6 9.3h4.8M14.6 7.3h3.4M14.6 5.3h2.1",
    },
    { element: "rect", role: "paper", x: 3, y: 13, width: 8, height: 8, rx: 1 },
    {
      element: "path",
      role: "accent",
      d: "M4.5 19.5h5M5 19.2v-2.5M7 19.2v-4M9 19.2v-5",
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 13,
      y: 13,
      width: 8,
      height: 8,
      rx: 1,
    },
    { element: "circle", role: "accent", cx: 17, cy: 17, r: 2.3 },
  ],
  "chart.title": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 3.2,
      width: 18,
      height: 17.6,
      rx: 1.2,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 6,
      y: 5.3,
      width: 12,
      height: 3,
      rx: 0.8,
    },
    { element: "path", role: "accent", emphasis: true, d: "M8 6.8h8" },
    {
      element: "path",
      role: "ink-soft",
      d: "M6 18V11m0 7h12M8.5 15l3-3 2.4 2 3.5-3.8",
    },
  ],
  "chart.axisTitles": [
    { element: "path", role: "ink", d: "M6 5v13h14" },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 10.5,
      y: 19,
      width: 9.5,
      height: 2.5,
      rx: 0.6,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 2.5,
      y: 6.5,
      width: 2.5,
      height: 9.5,
      rx: 0.6,
    },
    {
      element: "polyline",
      role: "accent",
      emphasis: true,
      points: "8,15 11,11 14,13 18,8",
    },
  ],
  "chart.legend": [
    { element: "path", role: "ink", d: "M3 3.2v17.8h18" },
    {
      element: "polyline",
      role: "accent",
      emphasis: true,
      points: "4.5,14.8 8,10 11.5,12.5 15,6.8",
    },
    { element: "circle", role: "accent-fill", cx: 17.8, cy: 7, r: 1.1 },
    { element: "path", role: "ink-soft", d: "M19.5 7h2M17.8 11.5h4M17.8 16h4" },
    { element: "circle", role: "warm-fill", cx: 16.2, cy: 11.5, r: 0.8 },
    { element: "circle", role: "ink-soft", cx: 16.2, cy: 16, r: 0.8 },
  ],
  "chart.dataLabels": [
    { element: "path", role: "ink", d: "M3 3.2v17.8h18" },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 5.3,
      y: 13,
      width: 3.5,
      height: 7.3,
      rx: 0.4,
    },
    {
      element: "rect",
      role: "accent-fill",
      x: 10.3,
      y: 9.4,
      width: 3.5,
      height: 10.9,
      rx: 0.4,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 15.3,
      y: 5.8,
      width: 3.5,
      height: 14.5,
      rx: 0.4,
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M5.6 11h3M10.6 7.4h3M15.6 3.8h3",
    },
  ],
  "chart.dataTable": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 3,
      width: 18,
      height: 18,
      rx: 1,
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M4.8 13V6m0 7h14M7.2 11l3-3 2.8 2 3.5-4",
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M3.5 14.5h17M3.5 17.5h17M9 14.5v6M15 14.5v6",
    },
  ],
  "chart.axes": [
    { element: "path", role: "accent", emphasis: true, d: "M5 3.5v15.5h15.5" },
    {
      element: "path",
      role: "ink-soft",
      d: "M3.5 7h3M3.5 11h3M3.5 15h3M9 17.5v3M13 17.5v3M17 17.5v3",
    },
    {
      element: "polyline",
      role: "accent",
      points: "5,3.5 3.7,5 6.3,5",
    },
    {
      element: "path",
      role: "accent",
      d: "M20.5 19l-1.5-1.3v2.6",
    },
  ],
  "chart.gridlines": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 3.5,
      width: 18,
      height: 17,
      rx: 1,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M3.5 7.5h17M3.5 11.7h17M3.5 15.9h17M7.5 4v16M12 4v16M16.5 4v16",
    },
    { element: "path", role: "accent", emphasis: true, d: "M3 20.5V3.5h18" },
  ],
  "chart.format": [
    { element: "path", role: "ink", d: "M3 3.2v17.8h15" },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 5,
      y: 13.5,
      width: 3,
      height: 6.8,
      rx: 0.4,
    },
    {
      element: "rect",
      role: "accent-fill",
      x: 9.5,
      y: 9.5,
      width: 3,
      height: 10.8,
      rx: 0.4,
    },
    {
      element: "path",
      role: "paper",
      d: "m14.1 14.4 4.8-4.8 2.8 2.8-4.8 4.8z",
    },
    {
      element: "path",
      role: "warm",
      emphasis: true,
      d: "m18.9 9.6 2.8 2.8M15.1 18.3c1.4-.4 2.3.2 2.7 1.5-.4 1.2-1.5 1.8-3.4 1.7",
    },
  ],
  "link.insert": [
    {
      element: "path",
      role: "ink",
      emphasis: true,
      d: "m9.7 14.3-1.9 1.9a3.8 3.8 0 0 1-5.4-5.4l3.1-3.1a3.8 3.8 0 0 1 5.4 0l.5.5",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "m14.3 9.7 1.9-1.9a3.8 3.8 0 1 1 5.4 5.4l-3.1 3.1a3.8 3.8 0 0 1-5.4 0l-.5-.5",
    },
    { element: "line", role: "ink-soft", x1: 8.4, y1: 15.6, x2: 15.6, y2: 8.4 },
  ],
  "comment.add": [
    {
      element: "path",
      role: "paper",
      d: "M3 3.8h18v13.4h-9.8l-4.7 3.1v-3.1H3z",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M12 7v7.2M8.4 10.6h7.2",
    },
  ],
  "textBox.insert": [
    {
      element: "rect",
      role: "paper",
      x: 2.8,
      y: 4,
      width: 18.4,
      height: 16,
      rx: 1.2,
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M5.7 17 9 7.6h1.7L14 17M7.1 13.2h5.5",
    },
    { element: "path", role: "ink-soft", d: "M15.8 9h3M15.8 12h3M15.8 15h3" },
  ],
  // Independently authored for this shared package; no external ancestry.
  "symbol.insert": [
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M7.1 19.3H3.4v-1.4h2.4c-1.7-1.7-2.7-4-2.7-6.4 0-5 3.9-8.6 8.9-8.6s8.9 3.6 8.9 8.6c0 2.4-1 4.7-2.7 6.4h2.4v1.4h-3.7v-3c1.4-1.2 2.3-2.9 2.3-4.8 0-3.5-2.8-5.9-7.2-5.9s-7.2 2.4-7.2 5.9c0 1.9.9 3.6 2.3 4.8z",
    },
  ],
  // Editor-neutral form-field faces. The glyphs describe durable form
  // semantics without PDF branding, font glyphs, or state-dependent artwork.
  "form.field": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 3,
      width: 18,
      height: 18,
      rx: 1.4,
    },
    { element: "path", role: "ink-soft", d: "M6 7h9M6 10.5h6.5M8.8 15.5h7.2" },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 5.5,
      y: 13.2,
      width: 2.4,
      height: 2.4,
      rx: 0.45,
    },
    { element: "path", role: "accent", emphasis: true, d: "M17.8 5.7v7" },
  ],
  "form.properties": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 3.5,
      width: 18,
      height: 17,
      rx: 1.4,
    },
    { element: "line", role: "ink-soft", x1: 6, y1: 8, x2: 18, y2: 8 },
    { element: "line", role: "ink-soft", x1: 6, y1: 12, x2: 18, y2: 12 },
    { element: "line", role: "ink-soft", x1: 6, y1: 16, x2: 18, y2: 16 },
    { element: "circle", role: "accent-fill", cx: 10, cy: 8, r: 1.35 },
    { element: "circle", role: "accent-fill", cx: 15, cy: 12, r: 1.35 },
    { element: "circle", role: "accent-fill", cx: 8, cy: 16, r: 1.35 },
  ],
  "form.appearance": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 3.5,
      width: 18,
      height: 17,
      rx: 1.4,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 5.2,
      y: 13.8,
      width: 13.6,
      height: 4.3,
      rx: 0.7,
    },
    { element: "path", role: "accent", emphasis: true, d: "M5.2 13.8h13.6" },
    { element: "circle", role: "accent-fill", cx: 7.3, cy: 8.3, r: 1.5 },
    { element: "circle", role: "warm-fill", cx: 12, cy: 8.3, r: 1.5 },
    { element: "circle", role: "ink-soft", cx: 16.7, cy: 8.3, r: 1.5 },
  ],
  "form.options": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 3.5,
      width: 18,
      height: 17,
      rx: 1.4,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 5.4,
      y: 6.2,
      width: 3,
      height: 3,
      rx: 0.5,
    },
    {
      element: "polyline",
      role: "accent",
      emphasis: true,
      points: "6,7.7 6.8,8.5 8.2,6.8",
    },
    { element: "line", role: "ink-soft", x1: 10.5, y1: 7.7, x2: 18.1, y2: 7.7 },
    {
      element: "rect",
      role: "paper",
      x: 5.4,
      y: 11,
      width: 3,
      height: 3,
      rx: 0.5,
    },
    {
      element: "line",
      role: "ink-soft",
      x1: 10.5,
      y1: 12.5,
      x2: 18.1,
      y2: 12.5,
    },
    {
      element: "rect",
      role: "paper",
      x: 5.4,
      y: 15.8,
      width: 3,
      height: 3,
      rx: 0.5,
    },
    {
      element: "line",
      role: "ink-soft",
      x1: 10.5,
      y1: 17.3,
      x2: 18.1,
      y2: 17.3,
    },
  ],
  "form.positionSize": [
    {
      element: "rect",
      role: "paper",
      x: 6.2,
      y: 6.2,
      width: 14.4,
      height: 14.4,
      rx: 1.2,
    },
    { element: "path", role: "ink-soft", d: "M9 10h8.8M9 13.2h6.2" },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M6.3 3.4h14.2M6.3 3.4l2-1.5M6.3 3.4l2 1.5M20.5 3.4l-2-1.5M20.5 3.4l-2 1.5",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M3.4 6.3v14.2M3.4 6.3l-1.5 2M3.4 6.3l1.5 2M3.4 20.5l-1.5-2M3.4 20.5l1.5-2",
    },
  ],
  "form.actions": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 3.5,
      width: 14.5,
      height: 17,
      rx: 1.4,
    },
    { element: "path", role: "ink-soft", d: "M6 8h7.7M6 11.5h5.5M6 15h6.8" },
    { element: "circle", role: "accent-soft-fill", cx: 18.1, cy: 14.8, r: 3.9 },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "m18.7 10.8-3 4.2h2.1l-.4 3.8 3.1-4.5h-2.1z",
    },
    { element: "path", role: "warm", d: "M17 7.3h4M19 5.3v4" },
  ],
  "edit.find": [...FIND_LENS],
  "edit.findPrevious": [
    ...FIND_LENS,
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M18.7 11.2V4.1m0 0-2.5 2.5m2.5-2.5 2.5 2.5",
    },
  ],
  "edit.findNext": [
    ...FIND_LENS,
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M18.7 4.1v7.1m0 0-2.5-2.5m2.5 2.5 2.5-2.5",
    },
  ],
  "edit.matchCase": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 4,
      width: 18,
      height: 16,
      rx: 1.3,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M6 16V7h3.8v9M6 11.2h3.8M13 16V9.5h3v6.5M13 12.7h3",
    },
    { element: "path", role: "accent", emphasis: true, d: "M18.4 7v9" },
  ],
  "edit.matchWholeWord": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 5,
      width: 18,
      height: 14,
      rx: 1.2,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M7.2 9h9.6M7.2 12h7.2M7.2 15h9.6",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M5.8 7H4.5v10h1.3M18.2 7h1.3v10h-1.3",
    },
  ],
  "edit.useRegex": [
    {
      element: "rect",
      role: "paper",
      x: 3,
      y: 4,
      width: 18,
      height: 16,
      rx: 1.3,
    },
    { element: "circle", role: "ink-soft", cx: 7, cy: 14.8, r: 1.1 },
    { element: "circle", role: "ink-soft", cx: 11, cy: 14.8, r: 1.1 },
    { element: "circle", role: "ink-soft", cx: 15, cy: 14.8, r: 1.1 },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M12 6.2v5.1M9.8 7.5l4.4 2.5M14.2 7.5 9.8 10",
    },
    { element: "path", role: "warm", d: "M18.1 7.2h2.1M19.2 6.1v2.1" },
  ],
  "edit.findFormat": [
    ...FIND_LENS,
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.5,
      y: 18.2,
      width: 10.5,
      height: 2.5,
      rx: 0.7,
    },
    { element: "path", role: "warm", emphasis: true, d: "M3.5 18.2H14" },
  ],
  "edit.replace": [
    { element: "path", role: "ink-soft", d: "M4 6.2h10M4 17.8h10" },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M6 9.4h13.2m0 0-3-3m3 3-3 3M18 14.6H4.8m0 0 3-3m-3 3 3 3",
    },
  ],
  "edit.close": [
    { element: "circle", role: "paper", cx: 12, cy: 12, r: 8.5 },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "m8.4 8.4 7.2 7.2m0-7.2-7.2 7.2",
    },
  ],
  // Generic editing, text-format, and paragraph command faces. Formatting is
  // expressed through abstract text strokes rather than font glyphs or SVG text.
  "edit.hand": [
    {
      element: "path",
      role: "paper",
      d: "M8.1 11.1V5.2a1.35 1.35 0 0 1 2.7 0v4.2-5.7a1.35 1.35 0 0 1 2.7 0v5.7-4.8a1.35 1.35 0 0 1 2.7 0v5.5-3.4a1.35 1.35 0 0 1 2.7 0v6.5c0 4.6-2.8 7.2-7.1 7.2h-.9c-2.7 0-4.3-1.3-5.7-3.4L3.4 14.2a1.5 1.5 0 0 1 2.4-1.8z",
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M10.8 9.4v2M13.5 9.4v2M16.2 10.1v1.8",
    },
    { element: "path", role: "accent-soft-fill", d: "M7.2 17.1h10.7v3.1H8.8z" },
    { element: "path", role: "accent", emphasis: true, d: "M7.2 17.1h10.7" },
  ],
  "edit.select": [
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M4 7V4h3M17 4h3v3M20 17v3h-3M7 20H4v-3",
    },
    { element: "path", role: "paper", d: "m7.1 6.1 10.2 7-4.2 1.1-2.2 4.1z" },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "m13.2 14.3 3.2 4.2",
    },
    { element: "circle", role: "accent-fill", cx: 7.1, cy: 6.1, r: 1.25 },
  ],
  "edit.exit": [
    { element: "path", role: "paper", d: "M4 4h9v16H4z" },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M10 12h11m0 0-3-3m3 3-3 3",
    },
  ],
  "edit.saveStatus": [
    { element: "path", role: "paper", d: "M4.5 2.8h9.2l4.1 4.1v14.3H4.5z" },
    {
      element: "path",
      role: "ink-soft",
      d: "M13.7 2.8v4.1h4.1M7.5 10h7M7.5 13h5",
    },
    { element: "circle", role: "accent-soft-fill", cx: 17.6, cy: 17.5, r: 4.2 },
    {
      element: "polyline",
      role: "accent",
      emphasis: true,
      points: "15.5,17.5 17,19 20.1,15.5",
    },
  ],
  "edit.undo": [
    { element: "path", role: "ink", d: "M9 14.5 4 9.5l5-5" },
    { element: "path", role: "accent", emphasis: true, d: "M4 9.5h10.2a5.3 5.3 0 0 1 0 10.6H10.5" },
  ],
  "edit.redo": [
    { element: "path", role: "ink", d: "m15 14.5 5-5-5-5" },
    { element: "path", role: "accent", emphasis: true, d: "M20 9.5H9.8a5.3 5.3 0 0 0 0 10.6h3.7" },
  ],
  "file.import": [
    { element: "path", role: "ink", d: "M4 14.5v4a1.5 1.5 0 0 0 1.5 1.5h13a1.5 1.5 0 0 0 1.5-1.5v-4" },
    { element: "path", role: "accent", emphasis: true, d: "M12 15.5V4M7.5 8.5 12 4l4.5 4.5" },
  ],
  "file.export": [
    { element: "path", role: "ink", d: "M4 14.5v4a1.5 1.5 0 0 0 1.5 1.5h13a1.5 1.5 0 0 0 1.5-1.5v-4" },
    { element: "path", role: "accent", emphasis: true, d: "M12 4v11.5M7.5 11 12 15.5l4.5-4.5" },
  ],
  "file.print": [
    { element: "path", role: "paper", d: "M7 8V3.5h10V8" },
    { element: "path", role: "ink", d: "M6.5 17H5a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-1.5" },
    { element: "path", role: "paper", d: "M7 13.5h10v7H7z" },
    { element: "path", role: "accent", emphasis: true, d: "M9.5 17h5" },
    { element: "circle", role: "warm-fill", cx: 17.5, cy: 11, r: 0.9 },
  ],
  "file.versionHistory": [
    { element: "path", role: "ink", d: "M3.6 13.2A8.5 8.5 0 1 0 5.9 6.1L3.5 8.5M3.5 4v4.5H8" },
    { element: "path", role: "accent", emphasis: true, d: "M12 7.5V12l3 1.8" },
  ],
  "ai.assistant": [
    { element: "path", role: "accent-soft-fill", d: "M10 4l1.6 4.6a2 2 0 0 0 1.2 1.2L17.4 11.4l-4.6 1.6a2 2 0 0 0-1.2 1.2L10 18.8l-1.6-4.6a2 2 0 0 0-1.2-1.2L2.6 11.4l4.6-1.6a2 2 0 0 0 1.2-1.2z" },
    { element: "path", role: "accent", emphasis: true, d: "M10 4l1.6 4.6a2 2 0 0 0 1.2 1.2L17.4 11.4l-4.6 1.6a2 2 0 0 0-1.2 1.2L10 18.8l-1.6-4.6a2 2 0 0 0-1.2-1.2L2.6 11.4l4.6-1.6a2 2 0 0 0 1.2-1.2z" },
    { element: "path", role: "ink", d: "M18.5 3v4M16.5 5h4M19.5 16.5v3M18 18h3" },
  ],
  "media.play": [
    { element: "path", role: "accent-soft-fill", d: "M7.5 5.2v13.6a.8.8 0 0 0 1.2.7l11-6.8a.8.8 0 0 0 0-1.4l-11-6.8a.8.8 0 0 0-1.2.7z" },
    { element: "path", role: "accent", emphasis: true, d: "M7.5 5.2v13.6a.8.8 0 0 0 1.2.7l11-6.8a.8.8 0 0 0 0-1.4l-11-6.8a.8.8 0 0 0-1.2.7z" },
  ],
  "media.stop": [
    { element: "rect", role: "accent-soft-fill", x: 6, y: 6, width: 12, height: 12, rx: 2.2 },
    { element: "rect", role: "accent", emphasis: true, x: 6, y: 6, width: 12, height: 12, rx: 2.2 },
  ],
  "media.record": [
    { element: "circle", role: "warm-fill", cx: 12, cy: 12, r: 7 },
  ],
  "window.close": [
    { element: "path", role: "ink", emphasis: true, d: "M6 6l12 12M18 6 6 18" },
  ],
  "action.add": [
    { element: "path", role: "ink", emphasis: true, d: "M12 5v14M5 12h14" },
  ],
  "action.remove": [
    { element: "path", role: "ink", emphasis: true, d: "M5 12h14" },
  ],
  "action.addBox": [
    { element: "rect", role: "paper", x: 4, y: 4, width: 16, height: 16, rx: 2.5 },
    { element: "path", role: "accent", emphasis: true, d: "M12 8v8M8 12h8" },
  ],
  "action.removeBox": [
    { element: "rect", role: "paper", x: 4, y: 4, width: 16, height: 16, rx: 2.5 },
    { element: "path", role: "accent", emphasis: true, d: "M8 12h8" },
  ],
  "menu.chevron.up": [
    { element: "path", role: "ink", emphasis: true, d: "M5.8 14.8 12 8.6l6.2 6.2" },
  ],
  "nav.chevrons.left": [
    { element: "path", role: "ink", emphasis: true, d: "M11.5 6.5 6 12l5.5 5.5M18 6.5 12.5 12l5.5 5.5" },
  ],
  "nav.chevrons.right": [
    { element: "path", role: "ink", emphasis: true, d: "M12.5 6.5 18 12l-5.5 5.5M6 6.5l5.5 5.5L6 17.5" },
  ],
  "nav.chevrons.upDown": [
    { element: "path", role: "ink", emphasis: true, d: "M7.5 9.5 12 5l4.5 4.5M7.5 14.5 12 19l4.5-4.5" },
  ],
  "nav.chevrons.collapse": [
    { element: "path", role: "ink", emphasis: true, d: "M7.5 4.5 12 9l4.5-4.5M7.5 19.5 12 15l4.5 4.5" },
  ],
  "overflow.moreVertical": [
    { element: "path", role: "ink", emphasis: true, d: "M12 5.2v.01M12 12v.01M12 18.8v.01" },
  ],
  "nav.arrow.up": [
    { element: "path", role: "ink", emphasis: true, d: "M12 19.5v-15M6 10.5l6-6 6 6" },
  ],
  "nav.arrow.down": [
    { element: "path", role: "ink", emphasis: true, d: "M12 4.5v15M6 13.5l6 6 6-6" },
  ],
  "nav.arrow.left": [
    { element: "path", role: "ink", emphasis: true, d: "M19.5 12h-15M10.5 6l-6 6 6 6" },
  ],
  "nav.arrow.right": [
    { element: "path", role: "ink", emphasis: true, d: "M4.5 12h15M13.5 6l6 6-6 6" },
  ],
  "nav.arrow.upRight": [
    { element: "path", role: "ink", emphasis: true, d: "M7 17 17 7M8.5 7H17v8.5" },
  ],
  "nav.arrow.toTop": [
    { element: "path", role: "ink", d: "M12 20.5V8M6.5 13.5 12 8l5.5 5.5" },
    { element: "path", role: "accent", emphasis: true, d: "M5 3.5h14" },
  ],
  "nav.arrow.toBottom": [
    { element: "path", role: "ink", d: "M12 3.5V16M6.5 10.5 12 16l5.5-5.5" },
    { element: "path", role: "accent", emphasis: true, d: "M5 20.5h14" },
  ],
  "nav.corner.upLeft": [
    { element: "path", role: "ink", d: "M19.5 19v-6a3 3 0 0 0-3-3H5" },
    { element: "path", role: "accent", emphasis: true, d: "M9 6 5 10l4 4" },
  ],
  "nav.corner.upRight": [
    { element: "path", role: "ink", d: "M4.5 19v-6a3 3 0 0 1 3-3H19" },
    { element: "path", role: "accent", emphasis: true, d: "m15 6 4 4-4 4" },
  ],
  "key.enter": [
    { element: "path", role: "ink", d: "M19.5 5v6a3 3 0 0 1-3 3H5" },
    { element: "path", role: "accent", emphasis: true, d: "M9 10l-4 4 4 4" },
  ],
  "nav.move": [
    { element: "path", role: "ink", d: "M12 3v18M3 12h18" },
    { element: "path", role: "accent", emphasis: true, d: "M9.5 5.5 12 3l2.5 2.5M9.5 18.5 12 21l2.5-2.5M5.5 9.5 3 12l2.5 2.5M18.5 9.5 21 12l-2.5 2.5" },
  ],
  "nav.move.horizontal": [
    { element: "path", role: "ink", d: "M3.5 12h17" },
    { element: "path", role: "accent", emphasis: true, d: "M7 8.5 3.5 12 7 15.5M17 8.5l3.5 3.5-3.5 3.5" },
  ],
  "nav.move.vertical": [
    { element: "path", role: "ink", d: "M12 3.5v17" },
    { element: "path", role: "accent", emphasis: true, d: "M8.5 7 12 3.5 15.5 7M8.5 17l3.5 3.5 3.5-3.5" },
  ],
  "nav.resize.diagonal": [
    { element: "path", role: "ink", d: "M5 5l14 14" },
    { element: "path", role: "accent", emphasis: true, d: "M5 10.5V5h5.5M19 13.5V19h-5.5" },
  ],
  "data.sort": [
    { element: "path", role: "ink", d: "M7.5 4.5v15M4 16l3.5 3.5L11 16" },
    { element: "path", role: "accent", emphasis: true, d: "M16.5 19.5v-15M13 8l3.5-3.5L20 8" },
  ],
  "data.swap": [
    { element: "path", role: "ink", d: "M4 8h15M15.5 4.5 19 8l-3.5 3.5" },
    { element: "path", role: "accent", emphasis: true, d: "M20 16H5M8.5 12.5 5 16l3.5 3.5" },
  ],
  "action.refresh": [
    { element: "path", role: "ink", d: "M4.5 10.5a7.8 7.8 0 0 1 13.7-4l1.8 2M19.5 13.5a7.8 7.8 0 0 1-13.7 4l-1.8-2" },
    { element: "path", role: "accent", emphasis: true, d: "M20 4.5v4h-4M4 19.5v-4h4" },
  ],
  "action.reset": [
    { element: "path", role: "accent", emphasis: true, d: "M4.6 13.5A7.5 7.5 0 1 0 6.6 6.8L4 9.5" },
    { element: "path", role: "ink", d: "M4 5v4.5h4.5" },
  ],
  "action.rotateClockwise": [
    { element: "path", role: "accent", emphasis: true, d: "M19.4 13.5A7.5 7.5 0 1 1 17.4 6.8L20 9.5" },
    { element: "path", role: "ink", d: "M20 5v4.5h-4.5" },
  ],
  "window.maximize": [
    { element: "path", role: "ink", d: "M14 4h6v6M10 20H4v-6" },
    { element: "path", role: "accent", emphasis: true, d: "M20 4l-6.5 6.5M4 20l6.5-6.5" },
  ],
  "window.restore": [
    { element: "path", role: "ink", d: "M20 4l-6 6M4 20l6-6" },
    { element: "path", role: "accent", emphasis: true, d: "M14 5.5V10h4.5M10 18.5V14H5.5" },
  ],
  "link.external": [
    { element: "path", role: "ink", d: "M18.5 13.5v5a1.5 1.5 0 0 1-1.5 1.5H5.5A1.5 1.5 0 0 1 4 18.5V7a1.5 1.5 0 0 1 1.5-1.5h5" },
    { element: "path", role: "accent", emphasis: true, d: "M14 4h6v6M20 4l-9 9" },
  ],
  "link.remove": [
    { element: "path", role: "ink", emphasis: true, d: "m9.7 14.3-1.9 1.9a3.8 3.8 0 0 1-5.4-5.4l3.1-3.1a3.8 3.8 0 0 1 5.4 0l.5.5" },
    { element: "path", role: "ink", emphasis: true, d: "m14.3 9.7 1.9-1.9a3.8 3.8 0 1 1 5.4 5.4l-3.1 3.1a3.8 3.8 0 0 1-5.4 0l-.5-.5" },
    { element: "path", role: "warm", emphasis: true, d: "M4.5 3.5l15 17" },
  ],
  "status.alert": [
    { element: "circle", role: "paper", cx: 12, cy: 12, r: 8.5 },
    { element: "path", role: "warm", emphasis: true, d: "M12 7.5v5.5M12 16.4h.01" },
  ],
  "status.warning": [
    { element: "path", role: "paper", d: "M12 3.5 21.5 20h-19z" },
    { element: "path", role: "warm", emphasis: true, d: "M12 9.5v5M12 17.4h.01" },
  ],
  "status.info": [
    { element: "circle", role: "paper", cx: 12, cy: 12, r: 8.5 },
    { element: "path", role: "accent", emphasis: true, d: "M12 11v5.5M12 7.6h.01" },
  ],
  "status.success": [
    { element: "circle", role: "paper", cx: 12, cy: 12, r: 8.5 },
    { element: "polyline", role: "accent", emphasis: true, points: "7.8,12.3 10.7,15.2 16.2,9.2" },
  ],
  "status.error": [
    { element: "circle", role: "paper", cx: 12, cy: 12, r: 8.5 },
    { element: "path", role: "warm", emphasis: true, d: "m9 9 6 6m0-6-6 6" },
  ],
  "status.blocked": [
    { element: "circle", role: "paper", cx: 12, cy: 12, r: 8.5 },
    { element: "path", role: "warm", emphasis: true, d: "M6 6l12 12" },
  ],
  "status.removeCircle": [
    { element: "circle", role: "paper", cx: 12, cy: 12, r: 8.5 },
    { element: "path", role: "warm", emphasis: true, d: "M8 12h8" },
  ],
  "status.checkDouble": [
    { element: "path", role: "ink", emphasis: true, d: "M2.5 12.5 7 17l8.5-9.5" },
    { element: "path", role: "accent", emphasis: true, d: "m11.5 16.2.8.8 9.2-9.5" },
  ],
  "status.activity": [
    { element: "path", role: "accent", emphasis: true, d: "M3 12h4l2.5-7 5 14 2.5-7h4" },
  ],
  "status.award": [
    { element: "path", role: "accent", emphasis: true, d: "M8.8 13.6 7.5 21l4.5-2.5 4.5 2.5-1.3-7.4" },
    { element: "circle", role: "paper", cx: 12, cy: 9, r: 5.5 },
  ],
  "status.verified": [
    { element: "path", role: "paper", d: "M12 2.8l2.4 1.9 3-.2.9 2.9 2.5 1.7-1 2.9 1 2.9-2.5 1.7-.9 2.9-3-.2-2.4 1.9-2.4-1.9-3 .2-.9-2.9-2.5-1.7 1-2.9-1-2.9 2.5-1.7.9-2.9 3 .2z" },
    { element: "polyline", role: "accent", emphasis: true, points: "8.5,12.2 11,14.7 15.5,9.8" },
  ],
  "status.flag": [
    { element: "path", role: "accent-soft-fill", d: "M5.5 4.5h12l-2.5 4 2.5 4h-12z" },
    { element: "path", role: "accent", emphasis: true, d: "M5.5 4.5h12l-2.5 4 2.5 4h-12" },
    { element: "path", role: "ink", d: "M5.5 21V4" },
  ],
  "status.gauge": [
    { element: "path", role: "ink", d: "M4.6 17.5a8.5 8.5 0 1 1 14.8 0" },
    { element: "path", role: "accent", emphasis: true, d: "M12 13.5l4-5" },
    { element: "circle", role: "accent-fill", cx: 12, cy: 13.5, r: 1 },
  ],
  "status.bolt": [
    { element: "path", role: "accent-soft-fill", d: "M13 2.5 4.5 13.5H11l-1 8 8.5-11H12z" },
    { element: "path", role: "accent", emphasis: true, d: "M13 2.5 4.5 13.5H11l-1 8 8.5-11H12z" },
  ],
  "status.streak": [
    { element: "path", role: "warm-fill", d: "M12 3c.5 3.5 5.5 5.5 5.5 11a5.5 5.5 0 0 1-11 0c0-2 .8-3.3 2-4.5.3 1.5 1 2.3 2 2.5C10 9 10.5 5.5 12 3z" },
  ],
  "status.frozen": [
    { element: "path", role: "ink", d: "M12 3v18M4.2 7.5l15.6 9M19.8 7.5l-15.6 9" },
    { element: "path", role: "accent", emphasis: true, d: "M9.5 4.5 12 6.5l2.5-2M9.5 19.5l2.5-2 2.5 2" },
  ],
  "status.health": [
    { element: "path", role: "paper", d: "M12 20.5S3.5 15.5 3.5 9.3A4.6 4.6 0 0 1 12 7a4.6 4.6 0 0 1 8.5 2.3c0 6.2-8.5 11.2-8.5 11.2z" },
    { element: "path", role: "accent", emphasis: true, d: "M3.8 12.5h4.4l1.8-3 2.6 5.5 1.6-2.5h6" },
  ],
  "status.trophy": [
    { element: "path", role: "ink", d: "M7.5 5.5H4.5V7a3 3 0 0 0 3 3M16.5 5.5h3V7a3 3 0 0 1-3 3M12 14v6.5" },
    { element: "path", role: "paper", d: "M7.5 4h9v5.5a4.5 4.5 0 0 1-9 0z" },
    { element: "path", role: "accent", emphasis: true, d: "M8 20.5h8" },
  ],
  "goal.target": [
    { element: "circle", role: "paper", cx: 12, cy: 12, r: 8.5 },
    { element: "circle", role: "ink-soft", cx: 12, cy: 12, r: 5 },
    { element: "circle", role: "accent-fill", cx: 12, cy: 12, r: 1.6 },
  ],
  "idea.lightbulb": [
    { element: "path", role: "paper", d: "M8.5 14.5a6 6 0 1 1 7 0c-.6.6-1 1.3-1 2v.5h-5v-.5c0-.7-.4-1.4-1-2z" },
    { element: "path", role: "accent", emphasis: true, d: "M9.8 20.5h4.4" },
  ],
  "selection.checkbox": [
    { element: "rect", role: "paper", x: 4, y: 4, width: 16, height: 16, rx: 2.5 },
    { element: "polyline", role: "accent", emphasis: true, points: "8,12.2 11,15.2 16.2,9.2" },
  ],
  "selection.radio": [
    { element: "circle", role: "paper", cx: 12, cy: 12, r: 8.5 },
    { element: "circle", role: "accent-fill", cx: 12, cy: 12, r: 3 },
  ],
  "selection.marquee": [
    { element: "rect", role: "accent", emphasis: true, strokeDasharray: "2.4 2.4", x: 4, y: 4, width: 16, height: 16, rx: 1.5 },
  ],
  "selection.text": [
    { element: "rect", role: "accent", strokeDasharray: "2.4 2.4", x: 3.5, y: 3.5, width: 17, height: 17, rx: 1.5 },
    { element: "path", role: "ink", d: "M8 9h8M8 12h8M8 15h5" },
  ],
  "shape.circle": [
    { element: "circle", role: "ink", cx: 12, cy: 12, r: 8.5 },
  ],
  "shape.circleDashed": [
    { element: "circle", role: "ink-soft", strokeDasharray: "2.6 2.6", cx: 12, cy: 12, r: 8.5 },
  ],
  "shape.square": [
    { element: "rect", role: "paper", x: 4.5, y: 4.5, width: 15, height: 15, rx: 2 },
  ],
  "shape.rectangle": [
    { element: "rect", role: "paper", x: 3, y: 6.5, width: 18, height: 11, rx: 1.5 },
  ],
  "shape.diamond": [
    { element: "path", role: "paper", d: "M12 3.5 20.5 12 12 20.5 3.5 12z" },
  ],
  "shape.triangle": [
    { element: "path", role: "paper", d: "M12 4.5 20.5 19.5h-17z" },
  ],
  "handle.grip.vertical": [
    { element: "path", role: "ink", emphasis: true, d: "M9 6v.01M15 6v.01M9 12v.01M15 12v.01M9 18v.01M15 18v.01" },
  ],
  "handle.grip.horizontal": [
    { element: "path", role: "ink", emphasis: true, d: "M6 9v.01M12 9v.01M18 9v.01M6 15v.01M12 15v.01M18 15v.01" },
  ],
  "symbol.asterisk": [
    { element: "path", role: "ink", emphasis: true, d: "M12 5v14M5.9 8.5l12.2 7M18.1 8.5 5.9 15.5" },
  ],
  "symbol.hash": [
    { element: "path", role: "ink", d: "M5 9h15M4 15h15" },
    { element: "path", role: "accent", emphasis: true, d: "M10.5 4 8.5 20M15.5 4l-2 16" },
  ],
  "symbol.slash": [
    { element: "path", role: "ink", emphasis: true, d: "M16.5 4 7.5 20" },
  ],
  "math.sigma": [
    { element: "path", role: "accent", emphasis: true, d: "M17.5 6.5v-2h-11l6 7.5-6 7.5h11v-2" },
  ],
  "math.radical": [
    { element: "path", role: "ink", d: "M3.5 12.5h3l3 7" },
    { element: "path", role: "accent", emphasis: true, d: "M9.5 19.5 14 4.5h7" },
  ],
  "object.align.left": [
    { element: "rect", role: "paper", x: 6.5, y: 6, width: 11, height: 4.5, rx: 1 },
    { element: "rect", role: "paper", x: 6.5, y: 13.5, width: 7, height: 4.5, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M3.5 3v18" },
  ],
  "object.align.center": [
    { element: "path", role: "accent", emphasis: true, d: "M12 3v18" },
    { element: "rect", role: "paper", x: 5, y: 6, width: 14, height: 4.5, rx: 1 },
    { element: "rect", role: "paper", x: 8, y: 13.5, width: 8, height: 4.5, rx: 1 },
  ],
  "object.align.right": [
    { element: "rect", role: "paper", x: 6.5, y: 6, width: 11, height: 4.5, rx: 1 },
    { element: "rect", role: "paper", x: 10.5, y: 13.5, width: 7, height: 4.5, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M20.5 3v18" },
  ],
  "object.align.top": [
    { element: "rect", role: "paper", x: 6, y: 6.5, width: 4.5, height: 11, rx: 1 },
    { element: "rect", role: "paper", x: 13.5, y: 6.5, width: 4.5, height: 7, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M3 3.5h18" },
  ],
  "object.align.middle": [
    { element: "path", role: "accent", emphasis: true, d: "M3 12h18" },
    { element: "rect", role: "paper", x: 6, y: 5, width: 4.5, height: 14, rx: 1 },
    { element: "rect", role: "paper", x: 13.5, y: 8, width: 4.5, height: 8, rx: 1 },
  ],
  "object.align.bottom": [
    { element: "rect", role: "paper", x: 6, y: 6.5, width: 4.5, height: 11, rx: 1 },
    { element: "rect", role: "paper", x: 13.5, y: 10.5, width: 4.5, height: 7, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M3 20.5h18" },
  ],
  "object.distribute.horizontal": [
    { element: "rect", role: "paper", x: 7, y: 7, width: 3.5, height: 10, rx: 1 },
    { element: "rect", role: "paper", x: 13.5, y: 7, width: 3.5, height: 10, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M3.5 3.5v17M20.5 3.5v17" },
  ],
  "object.distribute.vertical": [
    { element: "rect", role: "paper", x: 7, y: 7, width: 10, height: 3.5, rx: 1 },
    { element: "rect", role: "paper", x: 7, y: 13.5, width: 10, height: 3.5, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M3.5 3.5h17M3.5 20.5h17" },
  ],
  "object.layers": [
    { element: "path", role: "accent", emphasis: true, d: "M3.5 16 12 20.5l8.5-4.5" },
    { element: "path", role: "ink-soft", d: "M3.5 12.5 12 17l8.5-4.5" },
    { element: "path", role: "paper", d: "M12 3.5 21 8.5l-9 5-9-5z" },
  ],
  "object.boxes": [
    { element: "rect", role: "paper", x: 3, y: 12.5, width: 8, height: 8, rx: 1 },
    { element: "rect", role: "paper", x: 13, y: 12.5, width: 8, height: 8, rx: 1 },
    { element: "rect", role: "accent-soft-fill", x: 8, y: 3.5, width: 8, height: 8, rx: 1 },
    { element: "rect", role: "accent", emphasis: true, x: 8, y: 3.5, width: 8, height: 8, rx: 1 },
  ],
  "object.combine": [
    { element: "rect", role: "paper", x: 4, y: 4, width: 10.5, height: 10.5, rx: 1.5 },
    { element: "rect", role: "accent-soft-fill", x: 9.5, y: 9.5, width: 10.5, height: 10.5, rx: 1.5 },
    { element: "rect", role: "accent", emphasis: true, x: 9.5, y: 9.5, width: 10.5, height: 10.5, rx: 1.5 },
  ],
  "object.package": [
    { element: "path", role: "paper", d: "M12 3 20 7v10l-8 4-8-4V7z" },
    { element: "path", role: "ink-soft", d: "M4 7l8 4 8-4M12 11v10" },
    { element: "path", role: "accent", emphasis: true, d: "M8 5l8 4" },
  ],
  "object.package.verified": [
    { element: "path", role: "paper", d: "M12 3 20 7v10l-8 4-8-4V7z" },
    { element: "path", role: "ink-soft", d: "M4 7l8 4 8-4M12 11v10" },
    { element: "circle", role: "paper", cx: 18, cy: 18, r: 4 },
    { element: "polyline", role: "accent", emphasis: true, points: "16.2,18 17.5,19.3 19.9,16.7" },
  ],
  "object.package.search": [
    { element: "path", role: "paper", d: "M12 3 20 7v10l-8 4-8-4V7z" },
    { element: "path", role: "ink-soft", d: "M4 7l8 4 8-4M12 11v10" },
    { element: "circle", role: "paper", cx: 17.2, cy: 17.2, r: 3.2 },
    { element: "path", role: "accent", emphasis: true, d: "M19.6 19.6 22 22" },
  ],
  "object.scale": [
    { element: "rect", role: "ink-soft", strokeDasharray: "2.4 2.4", x: 3.5, y: 3.5, width: 17, height: 17, rx: 1.5 },
    { element: "rect", role: "paper", x: 3.5, y: 10.5, width: 10, height: 10, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M14 10l6-6M15.5 4H20v4.5" },
  ],
  "draw.curve": [
    { element: "path", role: "accent", emphasis: true, d: "M5 16.7C5 9 19 15 19 7.3" },
    { element: "circle", role: "paper", cx: 5, cy: 18.5, r: 1.8 },
    { element: "circle", role: "paper", cx: 19, cy: 5.5, r: 1.8 },
  ],
  "pointer.arrow": [
    { element: "path", role: "paper", d: "m5.5 4.5 14 5.5-6 2.3-2.3 6z" },
    { element: "path", role: "accent", emphasis: true, d: "m13.5 12.3 5.5 5.5" },
  ],
  "pointer.click": [
    { element: "path", role: "paper", d: "m8.5 8.5 11.5 4.5-4.9 1.8-1.8 4.9z" },
    { element: "path", role: "accent", emphasis: true, d: "M8.5 3v2.2M3 8.5h2.2M4.6 4.6l1.6 1.6" },
  ],
  "layout.columns.two": [
    { element: "rect", role: "paper", x: 3.5, y: 4, width: 17, height: 16, rx: 1.5 },
    { element: "path", role: "accent", emphasis: true, d: "M12 4v16" },
  ],
  "layout.columns.three": [
    { element: "rect", role: "paper", x: 3.5, y: 4, width: 17, height: 16, rx: 1.5 },
    { element: "path", role: "accent", emphasis: true, d: "M9.2 4v16M14.8 4v16" },
  ],
  "layout.rows": [
    { element: "rect", role: "paper", x: 3.5, y: 4, width: 17, height: 16, rx: 1.5 },
    { element: "path", role: "accent", emphasis: true, d: "M3.5 9.3h17M3.5 14.7h17" },
  ],
  "layout.panel.top": [
    { element: "rect", role: "paper", x: 3.5, y: 4, width: 17, height: 16, rx: 1.5 },
    { element: "rect", role: "accent-soft-fill", x: 4, y: 4.5, width: 16, height: 4.5, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M3.5 9h17" },
  ],
  "layout.panel.bottom": [
    { element: "rect", role: "paper", x: 3.5, y: 4, width: 17, height: 16, rx: 1.5 },
    { element: "rect", role: "accent-soft-fill", x: 4, y: 15, width: 16, height: 4.5, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M3.5 15h17" },
  ],
  "layout.panel.left": [
    { element: "rect", role: "paper", x: 3.5, y: 4, width: 17, height: 16, rx: 1.5 },
    { element: "rect", role: "accent-soft-fill", x: 4, y: 4.5, width: 5, height: 15, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M9 4v16" },
  ],
  "layout.panel.left.close": [
    { element: "rect", role: "paper", x: 3.5, y: 4, width: 17, height: 16, rx: 1.5 },
    { element: "path", role: "accent", emphasis: true, d: "M9 4v16" },
    { element: "path", role: "ink", d: "M16 9.5 13.5 12l2.5 2.5" },
  ],
  "layout.panel.left.open": [
    { element: "rect", role: "paper", x: 3.5, y: 4, width: 17, height: 16, rx: 1.5 },
    { element: "path", role: "accent", emphasis: true, d: "M9 4v16" },
    { element: "path", role: "ink", d: "M13.5 9.5 16 12l-2.5 2.5" },
  ],
  "layout.panel.right": [
    { element: "rect", role: "paper", x: 3.5, y: 4, width: 17, height: 16, rx: 1.5 },
    { element: "rect", role: "accent-soft-fill", x: 15, y: 4.5, width: 5, height: 15, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M15 4v16" },
  ],
  "layout.panel.right.close": [
    { element: "rect", role: "paper", x: 3.5, y: 4, width: 17, height: 16, rx: 1.5 },
    { element: "path", role: "accent", emphasis: true, d: "M15 4v16" },
    { element: "path", role: "ink", d: "M8 9.5l2.5 2.5L8 14.5" },
  ],
  "layout.template": [
    { element: "rect", role: "paper", x: 3.5, y: 4, width: 17, height: 16, rx: 1.5 },
    { element: "path", role: "ink-soft", d: "M9 9v11" },
    { element: "path", role: "accent", emphasis: true, d: "M3.5 9h17" },
  ],
  "layout.insertBetween": [
    { element: "rect", role: "paper", x: 6.5, y: 3.5, width: 14, height: 6, rx: 1 },
    { element: "rect", role: "paper", x: 6.5, y: 14.5, width: 14, height: 6, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M2.5 12h5M5.5 10l2 2-2 2" },
  ],
  "layout.separator.horizontal": [
    { element: "path", role: "ink", d: "M8.5 6.5 12 3l3.5 3.5M8.5 17.5 12 21l3.5-3.5" },
    { element: "path", role: "accent", emphasis: true, d: "M3 12h18" },
  ],
  "layout.separator.vertical": [
    { element: "path", role: "ink", d: "M6.5 8.5 3 12l3.5 3.5M17.5 8.5 21 12l-3.5 3.5" },
    { element: "path", role: "accent", emphasis: true, d: "M12 3v18" },
  ],
  "view.grid": [
    { element: "rect", role: "paper", x: 3.5, y: 3.5, width: 17, height: 17, rx: 1.5 },
    { element: "path", role: "accent", emphasis: true, d: "M3.5 12h17M12 3.5v17" },
  ],
  "view.grid.dense": [
    { element: "rect", role: "paper", x: 3.5, y: 3.5, width: 17, height: 17, rx: 1.5 },
    { element: "path", role: "accent", d: "M3.5 9.2h17M3.5 14.8h17M9.2 3.5v17M14.8 3.5v17" },
  ],
  "view.tiles": [
    { element: "rect", role: "paper", x: 3.5, y: 3.5, width: 7, height: 7, rx: 1.2 },
    { element: "rect", role: "paper", x: 13.5, y: 3.5, width: 7, height: 7, rx: 1.2 },
    { element: "rect", role: "paper", x: 3.5, y: 13.5, width: 7, height: 7, rx: 1.2 },
    { element: "rect", role: "accent-soft-fill", x: 13.5, y: 13.5, width: 7, height: 7, rx: 1.2 },
    { element: "rect", role: "accent", emphasis: true, x: 13.5, y: 13.5, width: 7, height: 7, rx: 1.2 },
  ],
  "view.dashboard": [
    { element: "rect", role: "paper", x: 3.5, y: 3.5, width: 7, height: 9.5, rx: 1.2 },
    { element: "rect", role: "paper", x: 13.5, y: 10.5, width: 7, height: 10, rx: 1.2 },
    { element: "rect", role: "paper", x: 3.5, y: 16, width: 7, height: 4.5, rx: 1.2 },
    { element: "rect", role: "accent-soft-fill", x: 13.5, y: 3.5, width: 7, height: 4, rx: 1.2 },
    { element: "rect", role: "accent", emphasis: true, x: 13.5, y: 3.5, width: 7, height: 4, rx: 1.2 },
  ],
  "view.show": [
    { element: "path", role: "paper", d: "M2.5 12s3.5-6.5 9.5-6.5 9.5 6.5 9.5 6.5-3.5 6.5-9.5 6.5S2.5 12 2.5 12z" },
    { element: "circle", role: "accent", emphasis: true, cx: 12, cy: 12, r: 2.8 },
  ],
  "view.hide": [
    { element: "path", role: "paper", d: "M2.5 12s3.5-6.5 9.5-6.5 9.5 6.5 9.5 6.5-3.5 6.5-9.5 6.5S2.5 12 2.5 12z" },
    { element: "circle", role: "ink-soft", cx: 12, cy: 12, r: 2.8 },
    { element: "path", role: "warm", emphasis: true, d: "M4.5 4l15 16" },
  ],
  "view.zoomOut": [
    { element: "circle", role: "paper", cx: 10, cy: 10, r: 6 },
    { element: "path", role: "accent", emphasis: true, d: "m14.5 14.5 6 6M7 10h6" },
  ],
  "view.lightMode": [
    { element: "path", role: "ink", d: "M12 2.5V5M12 19v2.5M2.5 12H5M19 12h2.5M5.3 5.3l1.8 1.8M16.9 16.9l1.8 1.8M5.3 18.7l1.8-1.8M16.9 7.1l1.8-1.8" },
    { element: "circle", role: "accent-soft-fill", cx: 12, cy: 12, r: 4 },
    { element: "circle", role: "accent", emphasis: true, cx: 12, cy: 12, r: 4 },
  ],
  "scan.line": [
    { element: "path", role: "ink", d: "M4 8V4h4M16 4h4v4M20 16v4h-4M8 20H4v-4" },
    { element: "path", role: "accent", emphasis: true, d: "M6.5 12h11" },
  ],
  "scan.search": [
    { element: "path", role: "ink", d: "M4 8V4h4M16 4h4v4M20 16v4h-4M8 20H4v-4" },
    { element: "circle", role: "paper", cx: 11.3, cy: 11.3, r: 3.2 },
    { element: "path", role: "accent", emphasis: true, d: "m13.7 13.7 2.8 2.8" },
  ],
  "edit.delete": [
    { element: "path", role: "paper", d: "M6.5 7.5h11l-.9 12a1.5 1.5 0 0 1-1.5 1.4H8.9a1.5 1.5 0 0 1-1.5-1.4z" },
    { element: "path", role: "ink-soft", d: "M10.2 11v6.5M13.8 11v6.5" },
    { element: "path", role: "warm", emphasis: true, d: "M4 7.5h16M9.5 7.5V5a1 1 0 0 1 1-1h3a1 1 0 0 1 1 1v2.5" },
  ],
  "edit.copy": [
    { element: "path", role: "ink-soft", d: "M8 16H5.5A1.5 1.5 0 0 1 4 14.5v-9A1.5 1.5 0 0 1 5.5 4h9A1.5 1.5 0 0 1 16 5.5V8" },
    { element: "rect", role: "paper", x: 8, y: 8, width: 12, height: 12, rx: 1.5 },
    { element: "path", role: "accent", emphasis: true, d: "M11 12.5h6M11 15.5h4" },
  ],
  "edit.duplicate": [
    { element: "path", role: "ink-soft", d: "M8 16H5.5A1.5 1.5 0 0 1 4 14.5v-9A1.5 1.5 0 0 1 5.5 4h9A1.5 1.5 0 0 1 16 5.5V8" },
    { element: "rect", role: "paper", x: 8, y: 8, width: 12, height: 12, rx: 1.5 },
    { element: "path", role: "accent", emphasis: true, d: "M14 11v6M11 14h6" },
  ],
  "edit.paste": [
    { element: "rect", role: "paper", x: 4, y: 4.5, width: 13, height: 16.5, rx: 2 },
    { element: "rect", role: "paper", x: 7.5, y: 2.8, width: 6, height: 3.4, rx: 1 },
    { element: "rect", role: "paper", x: 10.5, y: 11, width: 10, height: 10, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M13.5 14.5h4M13.5 17.5h2.5" },
  ],
  "edit.cut": [
    { element: "path", role: "accent", emphasis: true, d: "M8.6 8.6 20 18M8.6 15.4 20 6" },
    { element: "circle", role: "paper", cx: 6.5, cy: 7, r: 2.6 },
    { element: "circle", role: "paper", cx: 6.5, cy: 17, r: 2.6 },
  ],
  "edit.find.none": [
    { element: "circle", role: "paper", cx: 9.2, cy: 9.2, r: 5.4 },
    { element: "path", role: "ink", d: "M13.1 13.1l7.3 7.3" },
    { element: "path", role: "warm", emphasis: true, d: "m7.4 7.4 3.6 3.6m0-3.6-3.6 3.6" },
  ],
  "list.clipboard": [
    { element: "rect", role: "paper", x: 5, y: 4.5, width: 14, height: 16.5, rx: 2 },
    { element: "rect", role: "paper", x: 9, y: 2.8, width: 6, height: 3.4, rx: 1 },
    { element: "path", role: "ink-soft", emphasis: true, d: "M8.5 11v.01M8.5 15v.01" },
    { element: "path", role: "accent", emphasis: true, d: "M11.5 11h4M11.5 15h4" },
  ],
  "list.bullets": [
    { element: "path", role: "ink", d: "M9 6h11M9 12h11M9 18h11" },
    { element: "circle", role: "accent-fill", cx: 4.5, cy: 6, r: 0.7 },
    { element: "circle", role: "accent-fill", cx: 4.5, cy: 12, r: 0.7 },
    { element: "circle", role: "accent-fill", cx: 4.5, cy: 18, r: 0.7 },
  ],
  "list.numbered": [
    { element: "path", role: "ink", d: "M10 6h10M10 12h10M10 18h10" },
    { element: "path", role: "accent", d: "M4.4 5l1.2-.8V8M4 10.9c.4-.8 2.3-.8 2.3.3 0 1-2.3 1.6-2.3 2.6h2.4M4 16.5h2.3l-1.3 1.5c1.7 0 1.7 2.2 0 2.2-.5 0-.9-.2-1.1-.5" },
  ],
  "list.checks": [
    { element: "path", role: "ink", d: "M12 6h8.5M12 12h8.5M12 18h8.5" },
    { element: "path", role: "accent", emphasis: true, d: "m3.5 6 1.7 1.7 3.3-3.2M3.5 12l1.7 1.7 3.3-3.2M3.5 18l1.7 1.7 3.3-3.2" },
  ],
  "list.add": [
    { element: "path", role: "ink", d: "M4 6h12M4 12h9M4 18h9" },
    { element: "path", role: "accent", emphasis: true, d: "M18 12v6M15 15h6" },
  ],
  "list.tree": [
    { element: "path", role: "ink", d: "M4 5.5h9M10.5 12h9.5M10.5 18.5h9.5" },
    { element: "path", role: "accent", emphasis: true, d: "M5.5 8.5v8a2 2 0 0 0 2 2M5.5 12h2" },
  ],
  "list.media": [
    { element: "path", role: "ink", d: "M4 6h16M4 12h8M4 18h8" },
    { element: "path", role: "accent-soft-fill", d: "M15.5 13v6l5-3z" },
    { element: "path", role: "accent", emphasis: true, d: "M15.5 13v6l5-3z" },
  ],
  "text.type": [
    { element: "path", role: "ink", emphasis: true, d: "M12 5.5v14M9 19.5h6" },
    { element: "path", role: "accent", emphasis: true, d: "M5 8V5.5h14V8" },
  ],
  "text.heading": [
    { element: "path", role: "ink", emphasis: true, d: "M6 5v14M18 5v14" },
    { element: "path", role: "accent", emphasis: true, d: "M6 12h12" },
  ],
  "text.heading.1": [
    { element: "path", role: "ink", emphasis: true, d: "M4 6v12M12 6v12M4 12h8" },
    { element: "path", role: "accent", emphasis: true, d: "M16.5 11.5 19 10v8" },
  ],
  "text.heading.2": [
    { element: "path", role: "ink", emphasis: true, d: "M4 6v12M12 6v12M4 12h8" },
    { element: "path", role: "accent", emphasis: true, d: "M16 12a2.1 2.1 0 1 1 3.6 1.5L16 18h4.2" },
  ],
  "text.heading.3": [
    { element: "path", role: "ink", emphasis: true, d: "M4 6v12M12 6v12M4 12h8" },
    { element: "path", role: "accent", emphasis: true, d: "M16 10.5h3.8l-2.2 3a2.3 2.3 0 1 1-2 3.6" },
  ],
  "text.paragraphMark": [
    { element: "path", role: "ink", d: "M13 4v16M17 4v16" },
    { element: "path", role: "accent", emphasis: true, d: "M19 4h-9a4 4 0 0 0 0 8h3" },
  ],
  "text.quote": [
    { element: "path", role: "ink", emphasis: true, d: "M5 16.5c2.5-.5 4-2.5 4-6V7H4.5v5H9" },
    { element: "path", role: "accent", emphasis: true, d: "M13.5 16.5c2.5-.5 4-2.5 4-6V7H13v5h4.5" },
  ],
  "text.wrap": [
    { element: "path", role: "ink", d: "M3.5 6h17M3.5 18h7" },
    { element: "path", role: "accent", emphasis: true, d: "M3.5 12h14a3 3 0 0 1 0 6h-4M15.5 15.5l-2 2.5 2 2.5" },
  ],
  "file.document": [
    { element: "path", role: "paper", d: "M5.5 2.5h8.5l4.5 4.5v14.5h-13z" },
    { element: "path", role: "ink-soft", d: "M14 2.5V7h4.5" },
    { element: "path", role: "accent", emphasis: true, d: "M8.5 11h5" },
    { element: "path", role: "ink-soft", d: "M8.5 14h7M8.5 17h7" },
  ],
  "file.new": [
    { element: "path", role: "paper", d: "M5.5 2.5h8.5l4.5 4.5v14.5h-13z" },
    { element: "path", role: "ink-soft", d: "M14 2.5V7h4.5" },
    { element: "path", role: "accent", emphasis: true, d: "M12 11v6M9 14h6" },
  ],
  "file.multiple": [
    { element: "path", role: "ink-soft", d: "M6 6.5H4.5v14h10V19" },
    { element: "path", role: "paper", d: "M8 3h7.5l4 4v11H8z" },
    { element: "path", role: "ink-soft", d: "M15.5 3v4h4" },
    { element: "path", role: "accent", emphasis: true, d: "M11 11h5.5M11 14h4" },
  ],
  "file.download": [
    { element: "path", role: "paper", d: "M5.5 2.5h8.5l4.5 4.5v14.5h-13z" },
    { element: "path", role: "ink-soft", d: "M14 2.5V7h4.5" },
    { element: "path", role: "accent", emphasis: true, d: "M12 10.5v6.5M9.3 14.5 12 17.2l2.7-2.7" },
  ],
  "file.upload": [
    { element: "path", role: "paper", d: "M5.5 2.5h8.5l4.5 4.5v14.5h-13z" },
    { element: "path", role: "ink-soft", d: "M14 2.5V7h4.5" },
    { element: "path", role: "accent", emphasis: true, d: "M12 17.5V11M9.3 13.7 12 11l2.7 2.7" },
  ],
  "file.archive": [
    { element: "path", role: "paper", d: "M5.5 2.5h8.5l4.5 4.5v14.5h-13z" },
    { element: "path", role: "ink-soft", d: "M14 2.5V7h4.5" },
    { element: "path", role: "accent", emphasis: true, d: "M10.5 10h1.5M10.5 12.5h1.5M10.5 15h1.5M10.5 17.5h1.5" },
  ],
  "file.checked": [
    { element: "path", role: "paper", d: "M5.5 2.5h8.5l4.5 4.5v14.5h-13z" },
    { element: "path", role: "ink-soft", d: "M14 2.5V7h4.5" },
    { element: "polyline", role: "accent", emphasis: true, points: "8.8,14.5 11,16.7 15.2,12" },
  ],
  "file.code": [
    { element: "path", role: "paper", d: "M5.5 2.5h8.5l4.5 4.5v14.5h-13z" },
    { element: "path", role: "ink-soft", d: "M14 2.5V7h4.5" },
    { element: "path", role: "accent", emphasis: true, d: "M10.5 11.5l-2 2.5 2 2.5M13.5 11.5l2 2.5-2 2.5" },
  ],
  "file.diff": [
    { element: "path", role: "paper", d: "M5.5 2.5h8.5l4.5 4.5v14.5h-13z" },
    { element: "path", role: "ink-soft", d: "M14 2.5V7h4.5" },
    { element: "path", role: "accent", emphasis: true, d: "M12 10v4M10 12h4" },
    { element: "path", role: "warm", emphasis: true, d: "M10 17h4" },
  ],
  "file.image": [
    { element: "path", role: "paper", d: "M5.5 2.5h8.5l4.5 4.5v14.5h-13z" },
    { element: "path", role: "ink-soft", d: "M14 2.5V7h4.5" },
    { element: "circle", role: "ink-soft", cx: 10, cy: 11.5, r: 1.1 },
    { element: "polyline", role: "accent", points: "7.5,18 10.5,14.8 12.3,16.5 14,15 16.5,18" },
  ],
  "file.search": [
    { element: "path", role: "paper", d: "M5.5 2.5h8.5l4.5 4.5v14.5h-13z" },
    { element: "path", role: "ink-soft", d: "M14 2.5V7h4.5" },
    { element: "circle", role: "paper", cx: 11.5, cy: 14, r: 2.8 },
    { element: "path", role: "accent", emphasis: true, d: "m13.6 16.1 2.4 2.4" },
  ],
  "file.spreadsheet": [
    { element: "path", role: "paper", d: "M5.5 2.5h8.5l4.5 4.5v14.5h-13z" },
    { element: "path", role: "ink-soft", d: "M14 2.5V7h4.5" },
    { element: "path", role: "accent", d: "M8 11h8v7H8zM8 14.5h8M11.5 11v7" },
  ],
  "file.font": [
    { element: "path", role: "paper", d: "M5.5 2.5h8.5l4.5 4.5v14.5h-13z" },
    { element: "path", role: "ink-soft", d: "M14 2.5V7h4.5" },
    { element: "path", role: "accent", emphasis: true, d: "M9 11.5h6M12 11.5v6" },
  ],
  "file.warning": [
    { element: "path", role: "paper", d: "M5.5 2.5h8.5l4.5 4.5v14.5h-13z" },
    { element: "path", role: "ink-soft", d: "M14 2.5V7h4.5" },
    { element: "path", role: "warm", emphasis: true, d: "M12 10.5v4M12 17.4h.01" },
  ],
  "file.news": [
    { element: "path", role: "ink", d: "M5.5 9h-2v9a2 2 0 0 0 2 2" },
    { element: "rect", role: "paper", x: 5.5, y: 4, width: 15, height: 16, rx: 1.5 },
    { element: "rect", role: "accent-soft-fill", x: 8.5, y: 7, width: 4, height: 4, rx: 0.5 },
    { element: "path", role: "ink-soft", d: "M14.5 7.5h3M14.5 10.5h3M8.5 17h9" },
    { element: "path", role: "accent", emphasis: true, d: "M8.5 14h9" },
  ],
  "file.save": [
    { element: "path", role: "paper", d: "M4 5.5A1.5 1.5 0 0 1 5.5 4H16l4 4v10.5a1.5 1.5 0 0 1-1.5 1.5h-13A1.5 1.5 0 0 1 4 18.5z" },
    { element: "path", role: "ink-soft", d: "M8 4v4.5h6.5V4" },
    { element: "path", role: "accent", emphasis: true, d: "M7.5 20v-6.5h9V20" },
  ],
  "file.restore": [
    { element: "rect", role: "paper", x: 3, y: 4, width: 18, height: 5, rx: 1 },
    { element: "path", role: "ink", d: "M5 9v9.5A1.5 1.5 0 0 0 6.5 20h11a1.5 1.5 0 0 0 1.5-1.5V9" },
    { element: "path", role: "accent", emphasis: true, d: "M12 17.5v-6M9.5 13.5 12 11l2.5 2.5" },
  ],
  "file.attach": [
    { element: "path", role: "ink", d: "M19.5 11.5 12 19a4.6 4.6 0 0 1-6.5-6.5l8-8a3.1 3.1 0 0 1 4.4 4.4l-8 8a1.6 1.6 0 0 1-2.2-2.2l7.3-7.3" },
  ],
  "folder.closed": [
    { element: "path", role: "paper", d: "M3 6.5A1.5 1.5 0 0 1 4.5 5h4.6l2 2.5h8.4A1.5 1.5 0 0 1 21 9v9a1.5 1.5 0 0 1-1.5 1.5h-15A1.5 1.5 0 0 1 3 18z" },
    { element: "path", role: "accent", emphasis: true, d: "M3 10.5h18" },
  ],
  "folder.open": [
    { element: "path", role: "ink", d: "M3 18V6.5A1.5 1.5 0 0 1 4.5 5h4.6l2 2.5h7.4A1.5 1.5 0 0 1 20 9v1.5" },
    { element: "path", role: "paper", d: "M5.6 10.5h15.9L19 19.5H3z" },
    { element: "path", role: "accent", emphasis: true, d: "M5.6 10.5h15.9" },
  ],
  "folder.new": [
    { element: "path", role: "paper", d: "M3 6.5A1.5 1.5 0 0 1 4.5 5h4.6l2 2.5h8.4A1.5 1.5 0 0 1 21 9v9a1.5 1.5 0 0 1-1.5 1.5h-15A1.5 1.5 0 0 1 3 18z" },
    { element: "path", role: "accent", emphasis: true, d: "M12 11v5.5M9.25 13.75h5.5" },
  ],
  "folder.move": [
    { element: "path", role: "paper", d: "M3 6.5A1.5 1.5 0 0 1 4.5 5h4.6l2 2.5h8.4A1.5 1.5 0 0 1 21 9v9a1.5 1.5 0 0 1-1.5 1.5h-15A1.5 1.5 0 0 1 3 18z" },
    { element: "path", role: "accent", emphasis: true, d: "M8 14h7M12.5 11.5 15 14l-2.5 2.5" },
  ],
  "bookmark.mark": [
    { element: "path", role: "accent-soft-fill", d: "M6.5 3.5h11v17L12 16.5l-5.5 4z" },
    { element: "path", role: "accent", emphasis: true, d: "M6.5 3.5h11v17L12 16.5l-5.5 4z" },
  ],
  "bookmark.saved": [
    { element: "path", role: "paper", d: "M6.5 3.5h11v17L12 16.5l-5.5 4z" },
    { element: "polyline", role: "accent", emphasis: true, points: "9.2,10 11.2,12 14.8,8" },
  ],
  "bookmark.add": [
    { element: "path", role: "paper", d: "M6.5 3.5h11v17L12 16.5l-5.5 4z" },
    { element: "path", role: "accent", emphasis: true, d: "M12 7v5.5M9.25 9.75h5.5" },
  ],
  "book.closed": [
    { element: "path", role: "paper", d: "M5 4.5A1.5 1.5 0 0 1 6.5 3H19v15H6.5A1.5 1.5 0 0 0 5 19.5z" },
    { element: "path", role: "ink", d: "M5 19.5A1.5 1.5 0 0 0 6.5 21H19v-3" },
    { element: "path", role: "accent", emphasis: true, d: "M9 7.5h6" },
  ],
  "book.bookmarked": [
    { element: "path", role: "paper", d: "M5 4.5A1.5 1.5 0 0 1 6.5 3H19v15H6.5A1.5 1.5 0 0 0 5 19.5z" },
    { element: "path", role: "ink", d: "M5 19.5A1.5 1.5 0 0 0 6.5 21H19v-3" },
    { element: "path", role: "accent", emphasis: true, d: "M10 3v7l2.5-2 2.5 2V3" },
  ],
  "book.library": [
    { element: "rect", role: "paper", x: 3.5, y: 4, width: 5, height: 16, rx: 1 },
    { element: "rect", role: "paper", x: 8.5, y: 4, width: 5, height: 16, rx: 1 },
    { element: "path", role: "ink-soft", d: "M5 8h2M10 8h2" },
    { element: "path", role: "accent", emphasis: true, d: "m14.8 6.6 3.4-.9 3.4 13.4-3.4.9z" },
  ],
  "notebook.edit": [
    { element: "rect", role: "paper", x: 5, y: 3, width: 12.5, height: 18, rx: 1.5 },
    { element: "path", role: "ink", d: "M3.5 7h3M3.5 12h3M3.5 17h3" },
    { element: "path", role: "accent", emphasis: true, d: "m12.5 15.5.6-2.6 6-6 2 2-6 6z" },
  ],
  "note.sticky": [
    { element: "path", role: "paper", d: "M4 4h16v10.5L14.5 20H4z" },
    { element: "path", role: "ink-soft", d: "M7.5 8.5h9M7.5 12h5" },
    { element: "path", role: "accent-soft-fill", d: "M14.5 20v-5.5H20z" },
    { element: "path", role: "accent", emphasis: true, d: "M14.5 20v-5.5H20" },
  ],
  "sign.signature": [
    { element: "path", role: "ink-soft", d: "M3 19.5h18" },
    { element: "path", role: "accent", emphasis: true, d: "M4 15.5c2-5 4-8.5 5.5-8.5 2 0-1.5 8.5.5 8.5 1.5 0 2-3.5 3.5-3.5 1.3 0 .8 2.8 2.3 2.8 1 0 1.7-1 4.2-1" },
  ],
  "sign.stamp": [
    { element: "path", role: "paper", d: "M9.5 3.5h5l-1 7h3.5a2 2 0 0 1 2 2V15H5v-2.5a2 2 0 0 1 2-2h3.5z" },
    { element: "path", role: "accent", emphasis: true, d: "M4.5 19.5h15" },
  ],
  "label.tag": [
    { element: "path", role: "paper", d: "M3.5 4h8l9 9-7.5 7.5-9.5-9.5z" },
    { element: "circle", role: "accent", emphasis: true, cx: 8, cy: 8.5, r: 1.3 },
  ],
  "image.gallery": [
    { element: "path", role: "ink-soft", d: "M6.5 20.5h12a2 2 0 0 0 2-2v-10" },
    { element: "rect", role: "paper", x: 3, y: 3.5, width: 14.5, height: 13.5, rx: 1.3 },
    { element: "circle", role: "ink-soft", cx: 12.5, cy: 7.2, r: 1.1 },
    { element: "polyline", role: "accent", points: "4.2,15.5 8.2,11 10.6,13.2 12.4,11.6 16.3,15.6" },
  ],
  "image.download": [
    { element: "rect", role: "paper", x: 3, y: 4, width: 16, height: 14, rx: 1.3 },
    { element: "circle", role: "ink-soft", cx: 14.8, cy: 7.7, r: 1.25 },
    { element: "polyline", role: "ink-soft", points: "4.2,16.7 8.8,11.7 11.5,14.2 13.6,12.3 15.5,14.2" },
    { element: "path", role: "accent", emphasis: true, d: "M19.5 14v7M17 18.5l2.5 2.5 2.5-2.5" },
  ],
  "image.upload": [
    { element: "rect", role: "paper", x: 3, y: 4, width: 16, height: 14, rx: 1.3 },
    { element: "circle", role: "ink-soft", cx: 14.8, cy: 7.7, r: 1.25 },
    { element: "polyline", role: "ink-soft", points: "4.2,16.7 8.8,11.7 11.5,14.2 13.6,12.3 15.5,14.2" },
    { element: "path", role: "accent", emphasis: true, d: "M19.5 21v-7M17 16.5l2.5-2.5 2.5 2.5" },
  ],
  "image.missing": [
    { element: "rect", role: "paper", x: 3, y: 4, width: 16, height: 14, rx: 1.3 },
    { element: "circle", role: "ink-soft", cx: 14.8, cy: 7.7, r: 1.25 },
    { element: "polyline", role: "ink-soft", points: "4.2,16.7 8.8,11.7 11.5,14.2 13.6,12.3 17.8,16.7" },
    { element: "path", role: "warm", emphasis: true, d: "M3.5 2.5l16 17.5" },
  ],
  "media.pause": [
    { element: "rect", role: "accent-soft-fill", x: 6.5, y: 5, width: 4, height: 14, rx: 1 },
    { element: "rect", role: "accent", emphasis: true, x: 6.5, y: 5, width: 4, height: 14, rx: 1 },
    { element: "rect", role: "accent-soft-fill", x: 13.5, y: 5, width: 4, height: 14, rx: 1 },
    { element: "rect", role: "accent", emphasis: true, x: 13.5, y: 5, width: 4, height: 14, rx: 1 },
  ],
  "media.pauseCircle": [
    { element: "circle", role: "paper", cx: 12, cy: 12, r: 8.5 },
    { element: "path", role: "accent", emphasis: true, d: "M10 9v6M14 9v6" },
  ],
  "media.next": [
    { element: "path", role: "accent-soft-fill", d: "M5.5 5.5v13l10-6.5z" },
    { element: "path", role: "accent", emphasis: true, d: "M5.5 5.5v13l10-6.5z" },
    { element: "path", role: "ink", emphasis: true, d: "M18.5 5.5v13" },
  ],
  "media.camera": [
    { element: "path", role: "paper", d: "M3 8.5A1.5 1.5 0 0 1 4.5 7h3L9 4.5h6L16.5 7h3A1.5 1.5 0 0 1 21 8.5v9a1.5 1.5 0 0 1-1.5 1.5h-15A1.5 1.5 0 0 1 3 17.5z" },
    { element: "circle", role: "accent", emphasis: true, cx: 12, cy: 13, r: 3.4 },
  ],
  "media.film": [
    { element: "rect", role: "paper", x: 3.5, y: 4, width: 17, height: 16, rx: 2 },
    { element: "path", role: "ink-soft", d: "M8 4v16M16 4v16M3.5 9H8M3.5 15H8M16 9h4.5M16 15h4.5" },
    { element: "path", role: "accent", emphasis: true, d: "M8 12h8" },
  ],
  "media.mic": [
    { element: "rect", role: "paper", x: 9, y: 3, width: 6, height: 11, rx: 3 },
    { element: "path", role: "ink", d: "M5.5 11.5a6.5 6.5 0 0 0 13 0" },
    { element: "path", role: "accent", emphasis: true, d: "M12 18v3" },
  ],
  "media.music": [
    { element: "path", role: "accent", emphasis: true, d: "M9 18V5.5l10-2V16" },
    { element: "circle", role: "paper", cx: 6.5, cy: 18, r: 2.5 },
    { element: "circle", role: "paper", cx: 16.5, cy: 16, r: 2.5 },
  ],
  "media.volume": [
    { element: "path", role: "paper", d: "M3.5 9.5H7L12 5v14l-5-4.5H3.5z" },
    { element: "path", role: "accent", emphasis: true, d: "M15.5 9a4.2 4.2 0 0 1 0 6M18 6.5a7.8 7.8 0 0 1 0 11" },
  ],
  "present.screen": [
    { element: "rect", role: "paper", x: 3, y: 4.5, width: 18, height: 11.5, rx: 1 },
    { element: "path", role: "ink", d: "M2 4.5h20M12 16v2.5M8.5 21l3.5-2.5 3.5 2.5" },
    { element: "polyline", role: "accent", emphasis: true, points: "7,12.5 10.5,9.5 13,11.5 17,8" },
  ],
  "present.slideshow": [
    { element: "rect", role: "paper", x: 3, y: 4, width: 18, height: 12.5, rx: 1.5 },
    { element: "path", role: "ink", d: "M12 16.5V20M8 20.5h8" },
    { element: "path", role: "accent-soft-fill", d: "M10.3 7.5V13l4.6-2.75z" },
    { element: "path", role: "accent", emphasis: true, d: "M10.3 7.5V13l4.6-2.75z" },
  ],
  "present.share": [
    { element: "rect", role: "paper", x: 3, y: 4, width: 18, height: 12.5, rx: 1.5 },
    { element: "path", role: "ink", d: "M12 16.5V20M8 20.5h8" },
    { element: "path", role: "accent", emphasis: true, d: "M12 13V7.5M9.5 10 12 7.5l2.5 2.5" },
  ],
  "data.database": [
    { element: "path", role: "paper", d: "M4.5 6.5v11c0 1.7 3.4 3 7.5 3s7.5-1.3 7.5-3v-11" },
    { element: "path", role: "paper", d: "M4.5 6.5c0-1.7 3.4-3 7.5-3s7.5 1.3 7.5 3-3.4 3-7.5 3-7.5-1.3-7.5-3z" },
    { element: "path", role: "accent", emphasis: true, d: "M4.5 12c0 1.7 3.4 3 7.5 3s7.5-1.3 7.5-3" },
  ],
  "data.calculator": [
    { element: "rect", role: "paper", x: 5, y: 3, width: 14, height: 18, rx: 2 },
    { element: "rect", role: "accent-soft-fill", x: 7.5, y: 5.5, width: 9, height: 3.5, rx: 0.6 },
    { element: "path", role: "accent", emphasis: true, d: "M7.5 9h9" },
    { element: "path", role: "ink", emphasis: true, d: "M8 12.8v.01M12 12.8v.01M16 12.8v.01M8 16.8v.01M12 16.8v.01M16 16.8v.01" },
  ],
  "data.filter": [
    { element: "path", role: "paper", d: "M3.5 5h17l-6.5 7.8v5.2l-4 2v-7.2z" },
    { element: "path", role: "accent", emphasis: true, d: "M3.5 5h17" },
  ],
  "data.filter.clear": [
    { element: "path", role: "paper", d: "M3 4.5h15l-5.7 7v4.7l-3.6 1.8v-6.5z" },
    { element: "path", role: "accent", d: "M3 4.5h15" },
    { element: "path", role: "warm", emphasis: true, d: "m16.5 16.5 4.5 4.5m0-4.5-4.5 4.5" },
  ],
  "data.merge": [
    { element: "path", role: "ink", d: "M6 3.5v4.5c0 3 6 3.5 6 7M18 3.5v4.5c0 3-6 3.5-6 7" },
    { element: "path", role: "accent", emphasis: true, d: "M12 15v5.5M9.5 18l2.5 2.5 2.5-2.5" },
  ],
  "data.split": [
    { element: "path", role: "ink", d: "M12 3.5V9" },
    { element: "path", role: "accent", emphasis: true, d: "M12 9c0 3.5-6 4-6 7v4.5M12 9c0 3.5 6 4 6 7v4.5" },
  ],
  "storage.drive": [
    { element: "path", role: "paper", d: "M3.5 13.5l3-8.5h11l3 8.5V18a1.5 1.5 0 0 1-1.5 1.5H5A1.5 1.5 0 0 1 3.5 18z" },
    { element: "path", role: "ink-soft", d: "M3.5 13.5h17" },
    { element: "path", role: "accent", emphasis: true, d: "M7 16.5v.01M10.5 16.5v.01" },
  ],
  "cloud.storage": [
    { element: "path", role: "paper", d: "M7 18.5a4 4 0 0 1-.6-7.95 5.8 5.8 0 0 1 11.2-.75 4.4 4.4 0 0 1-.6 8.7z" },
    { element: "path", role: "accent", emphasis: true, d: "M9 15h6" },
  ],
  "cloud.upload": [
    { element: "path", role: "paper", d: "M7 18.5a4 4 0 0 1-.6-7.95 5.8 5.8 0 0 1 11.2-.75 4.4 4.4 0 0 1-.6 8.7z" },
    { element: "path", role: "accent", emphasis: true, d: "M12 16.5v-5.5M9.6 13.2l2.4-2.4 2.4 2.4" },
  ],
  "code.inline": [
    { element: "path", role: "ink", d: "m8 7.5-4.5 4.5L8 16.5M16 7.5l4.5 4.5-4.5 4.5" },
    { element: "path", role: "accent", emphasis: true, d: "M13.8 5l-3.6 14" },
  ],
  "code.braces": [
    { element: "path", role: "ink", d: "M8.5 3.5H8a2 2 0 0 0-2 2v4a2.5 2.5 0 0 1-2 2.5 2.5 2.5 0 0 1 2 2.5v4a2 2 0 0 0 2 2h.5" },
    { element: "path", role: "accent", emphasis: true, d: "M15.5 3.5h.5a2 2 0 0 1 2 2v4a2.5 2.5 0 0 0 2 2.5 2.5 2.5 0 0 0-2 2.5v4a2 2 0 0 1-2 2h-.5" },
  ],
  "code.brackets": [
    { element: "path", role: "ink", d: "M8.5 4H5v16h3.5" },
    { element: "path", role: "accent", emphasis: true, d: "M15.5 4H19v16h-3.5" },
  ],
  "code.terminal": [
    { element: "rect", role: "paper", x: 3.5, y: 4, width: 17, height: 16, rx: 2 },
    { element: "path", role: "ink-soft", d: "M12.5 15h4" },
    { element: "path", role: "accent", emphasis: true, d: "m7.5 9.5 3 2.5-3 2.5" },
  ],
  "code.bug": [
    { element: "path", role: "ink", d: "M9.5 6.5 8 4.5M14.5 6.5l1.5-2M8 12H4.5M8 16l-3 1.5M16 12h3.5M16 16l3 1.5" },
    { element: "path", role: "paper", d: "M8 10a4 4 0 0 1 8 0v5a4 4 0 0 1-8 0z" },
    { element: "path", role: "accent", emphasis: true, d: "M12 11v8" },
  ],
  "code.branch": [
    { element: "path", role: "ink", d: "M6.5 8v8" },
    { element: "path", role: "accent", emphasis: true, d: "M17.5 9.9c0 3.6-3.5 4.6-8.7 5.3" },
    { element: "circle", role: "paper", cx: 6.5, cy: 5.5, r: 2.4 },
    { element: "circle", role: "paper", cx: 6.5, cy: 18.5, r: 2.4 },
    { element: "circle", role: "paper", cx: 17.5, cy: 7.5, r: 2.4 },
  ],
  "code.compare": [
    { element: "path", role: "ink", d: "M6 8.4v7.1M18 15.6V8.5" },
    { element: "path", role: "accent", emphasis: true, d: "M11 6h4.5A2.5 2.5 0 0 1 18 8.5M13 3.5 10.5 6 13 8.5M13 18H8.5A2.5 2.5 0 0 1 6 15.5M11 15.5l2.5 2.5-2.5 2.5" },
    { element: "circle", role: "paper", cx: 6, cy: 6, r: 2.4 },
    { element: "circle", role: "paper", cx: 18, cy: 18, r: 2.4 },
  ],
  "code.pullRequest": [
    { element: "path", role: "ink", d: "M6 8.4v7.2" },
    { element: "path", role: "accent", emphasis: true, d: "M18 15.6V9.5A3.5 3.5 0 0 0 14.5 6H11M13.5 3.5 11 6l2.5 2.5" },
    { element: "circle", role: "paper", cx: 6, cy: 6, r: 2.4 },
    { element: "circle", role: "paper", cx: 6, cy: 18, r: 2.4 },
    { element: "circle", role: "paper", cx: 18, cy: 18, r: 2.4 },
  ],
  "diagram.network": [
    { element: "path", role: "accent", emphasis: true, d: "M12 8v4M6 16v-2.5A1.5 1.5 0 0 1 7.5 12h9a1.5 1.5 0 0 1 1.5 1.5V16" },
    { element: "rect", role: "paper", x: 9, y: 3, width: 6, height: 5, rx: 1 },
    { element: "rect", role: "paper", x: 3, y: 16, width: 6, height: 5, rx: 1 },
    { element: "rect", role: "paper", x: 15, y: 16, width: 6, height: 5, rx: 1 },
  ],
  "diagram.workflow": [
    { element: "rect", role: "paper", x: 3, y: 3.5, width: 7.5, height: 7.5, rx: 1.2 },
    { element: "rect", role: "paper", x: 13.5, y: 13, width: 7.5, height: 7.5, rx: 1.2 },
    { element: "path", role: "accent", emphasis: true, d: "M6.75 11v3a2.75 2.75 0 0 0 2.75 2.75h4" },
  ],
  "window.blocks": [
    { element: "rect", role: "paper", x: 3.5, y: 10.5, width: 10, height: 10, rx: 1.2 },
    { element: "path", role: "ink-soft", d: "M8.5 10.5v10M3.5 15.5h10" },
    { element: "rect", role: "accent-soft-fill", x: 14.5, y: 3.5, width: 6, height: 6, rx: 1.2 },
    { element: "rect", role: "accent", emphasis: true, x: 14.5, y: 3.5, width: 6, height: 6, rx: 1.2 },
  ],
  "extension.puzzle": [
    { element: "path", role: "paper", d: "M5 8h4V6.5a2 2 0 1 1 4 0V8h4v4h1.5a2 2 0 1 1 0 4H17v4H5z" },
    { element: "path", role: "accent", emphasis: true, d: "M5 16h1.5a2 2 0 1 0 0-4H5" },
  ],
  "connect.plug": [
    { element: "path", role: "ink", d: "M9 3.5v4M15 3.5v4" },
    { element: "path", role: "paper", d: "M6.5 7.5h11V11a5.5 5.5 0 0 1-11 0z" },
    { element: "path", role: "accent", emphasis: true, d: "M12 16.5V21" },
  ],
  "connect.unplug": [
    { element: "path", role: "ink", d: "M9 3.5v4M15 3.5v4" },
    { element: "path", role: "paper", d: "M6.5 7.5h11V11a5.5 5.5 0 0 1-11 0z" },
    { element: "path", role: "ink-soft", d: "M12 16.5V21" },
    { element: "path", role: "warm", emphasis: true, d: "M4 4l16 16.5" },
  ],
  "lab.flask": [
    { element: "path", role: "paper", d: "M9.5 3.5h5M10 3.5v6L4.8 18.6a1.3 1.3 0 0 0 1.1 1.9h12.2a1.3 1.3 0 0 0 1.1-1.9L14 9.5v-6" },
    { element: "path", role: "accent", emphasis: true, d: "M7.3 15h9.4" },
  ],
  "app.settings": [
    { element: "polygon", role: "paper", points: "10.6,3.1 13.4,3.1 13.6,5.2 15.7,6 17.3,4.7 19.3,6.7 18,8.3 18.8,10.4 20.9,10.6 20.9,13.4 18.8,13.6 18,15.7 19.3,17.3 17.3,19.3 15.7,18 13.6,18.8 13.4,20.9 10.6,20.9 10.4,18.8 8.3,18 6.7,19.3 4.7,17.3 6,15.7 5.2,13.6 3.1,13.4 3.1,10.6 5.2,10.4 6,8.3 4.7,6.7 6.7,4.7 8.3,6 10.4,5.2" },
    { element: "circle", role: "accent", emphasis: true, cx: 12, cy: 12, r: 3 },
  ],
  "app.sliders": [
    { element: "path", role: "ink", d: "M3.5 7.5h8M16.5 7.5h4M3.5 16.5h4M12.5 16.5h8" },
    { element: "circle", role: "accent", emphasis: true, cx: 14, cy: 7.5, r: 2.5 },
    { element: "circle", role: "accent", emphasis: true, cx: 10, cy: 16.5, r: 2.5 },
  ],
  "app.tools": [
    { element: "path", role: "paper", d: "M20.3 6.2 17.5 9 15 8.5 14.5 6l2.8-2.8a5 5 0 0 0-6.2 6.5L4 16.8A2.2 2.2 0 0 0 7.2 20l7.1-7.1a5 5 0 0 0 6-6.7z" },
    { element: "circle", role: "accent-fill", cx: 5.7, cy: 18.3, r: 0.5 },
  ],
  "nav.home": [
    { element: "path", role: "paper", d: "M4 10.5 12 4l8 6.5V20H4z" },
    { element: "path", role: "accent", emphasis: true, d: "M9.5 20v-6h5v6" },
  ],
  "nav.map": [
    { element: "path", role: "paper", d: "M3.5 6.5 9 4l6 2.5L20.5 4v13.5L15 20l-6-2.5-5.5 2.5z" },
    { element: "path", role: "ink-soft", d: "M15 6.5V20" },
    { element: "path", role: "accent", emphasis: true, d: "M9 4v13.5" },
  ],
  "nav.pin": [
    { element: "path", role: "paper", d: "M12 21s-6.5-6-6.5-11a6.5 6.5 0 0 1 13 0c0 5-6.5 11-6.5 11z" },
    { element: "circle", role: "accent", emphasis: true, cx: 12, cy: 10, r: 2.3 },
  ],
  "nav.locate": [
    { element: "path", role: "ink", d: "M12 2.5V6M12 18v3.5M2.5 12H6M18 12h3.5" },
    { element: "circle", role: "paper", cx: 12, cy: 12, r: 6 },
    { element: "circle", role: "accent-fill", cx: 12, cy: 12, r: 1.8 },
  ],
  "web.globe": [
    { element: "circle", role: "paper", cx: 12, cy: 12, r: 8.5 },
    { element: "path", role: "ink-soft", d: "M3.5 12h17" },
    { element: "path", role: "accent", emphasis: true, d: "M12 3.5c2.6 2.3 3.9 5.1 3.9 8.5s-1.3 6.2-3.9 8.5c-2.6-2.3-3.9-5.1-3.9-8.5s1.3-6.2 3.9-8.5z" },
  ],
  "notify.bell": [
    { element: "path", role: "paper", d: "M6 16.5V11a6 6 0 0 1 12 0v5.5l1.5 2h-15z" },
    { element: "path", role: "accent", emphasis: true, d: "M10 21a2.2 2.2 0 0 0 4 0" },
  ],
  "notify.bellOff": [
    { element: "path", role: "paper", d: "M6 16.5V11a6 6 0 0 1 12 0v5.5l1.5 2h-15z" },
    { element: "path", role: "ink-soft", d: "M10 21a2.2 2.2 0 0 0 4 0" },
    { element: "path", role: "warm", emphasis: true, d: "M4 3.5l16 17" },
  ],
  "mail.envelope": [
    { element: "rect", role: "paper", x: 3, y: 5, width: 18, height: 14, rx: 2 },
    { element: "path", role: "accent", emphasis: true, d: "m3.5 7 8.5 6.5L20.5 7" },
  ],
  "mail.inbox": [
    { element: "path", role: "paper", d: "M3.5 13l3-8h11l3 8v5.5A1.5 1.5 0 0 1 19 20H5a1.5 1.5 0 0 1-1.5-1.5z" },
    { element: "path", role: "accent", emphasis: true, d: "M3.5 13h5l1.2 2.5h4.6l1.2-2.5h5" },
  ],
  "contact.phone": [
    { element: "path", role: "paper", d: "M6.5 3.5h3l1.5 4-2 1.5a11 11 0 0 0 6 6l1.5-2 4 1.5v3a2 2 0 0 1-2.2 2A16.5 16.5 0 0 1 4.5 5.7a2 2 0 0 1 2-2.2z" },
    { element: "path", role: "accent", emphasis: true, d: "M14.5 4a5.5 5.5 0 0 1 5.5 5.5" },
  ],
  "chat.bubble": [
    { element: "path", role: "paper", d: "M12 4a8 8 0 0 1 6.9 12l1.1 4.5-4.6-1.2A8 8 0 1 1 12 4z" },
    { element: "path", role: "accent", emphasis: true, d: "M8.5 10.5h7M8.5 13.5h4.5" },
  ],
  "chat.conversation": [
    { element: "path", role: "paper", d: "M3 4h13v9.5H8L5 16v-2.5H3z" },
    { element: "path", role: "paper", d: "M10 10h11v8.5h-2V21l-3-2.5h-6z" },
    { element: "path", role: "accent", emphasis: true, d: "M13 14.2h5" },
  ],
  "chat.send": [
    { element: "path", role: "ink", emphasis: true, d: "M20.5 3.5 3.5 10l6.5 3.5 3.5 6.5zM10 13.5 20.5 3.5" },
  ],
  "comment.bubble": [
    { element: "path", role: "paper", d: "M3 4h18v13H9l-4 4v-4H3z" },
    { element: "path", role: "accent", emphasis: true, d: "M7 8.5h10" },
    { element: "path", role: "ink-soft", d: "M7 12.5h6" },
  ],
  "people.user": [
    { element: "circle", role: "paper", cx: 12, cy: 8, r: 4 },
    { element: "path", role: "accent", emphasis: true, d: "M4.5 20.5a7.5 7.5 0 0 1 15 0" },
  ],
  "people.user.check": [
    { element: "circle", role: "paper", cx: 10, cy: 8, r: 3.8 },
    { element: "path", role: "ink", d: "M3 20.5a7 7 0 0 1 12.5-4.3" },
    { element: "polyline", role: "accent", emphasis: true, points: "15.5,17.5 17.7,19.7 21.5,15.3" },
  ],
  "people.user.add": [
    { element: "circle", role: "paper", cx: 10, cy: 8, r: 3.8 },
    { element: "path", role: "ink", d: "M3 20.5a7 7 0 0 1 12.5-4.3" },
    { element: "path", role: "accent", emphasis: true, d: "M18.5 14.5v6M15.5 17.5h6" },
  ],
  "people.user.remove": [
    { element: "circle", role: "paper", cx: 10, cy: 8, r: 3.8 },
    { element: "path", role: "ink", d: "M3 20.5a7 7 0 0 1 12.5-4.3" },
    { element: "path", role: "warm", emphasis: true, d: "M15.5 17.5h6" },
  ],
  "people.user.settings": [
    { element: "circle", role: "paper", cx: 10, cy: 8, r: 3.8 },
    { element: "path", role: "ink", d: "M3 20.5a7 7 0 0 1 12.5-4.3" },
    { element: "path", role: "ink-soft", d: "M18.5 13.8v1.2M18.5 20v1.2M15.3 15.65l1 .6M20.7 18.75l1 .6M15.3 19.35l1-.6M20.7 16.25l1-.6" },
    { element: "circle", role: "accent", emphasis: true, cx: 18.5, cy: 17.5, r: 2 },
  ],
  "people.group": [
    { element: "circle", role: "paper", cx: 9, cy: 8.5, r: 3.6 },
    { element: "path", role: "ink", d: "M2.5 20a6.5 6.5 0 0 1 13 0" },
    { element: "path", role: "accent", emphasis: true, d: "M15.5 5.2a3.6 3.6 0 0 1 0 6.6M18 15a6.5 6.5 0 0 1 3.5 5" },
  ],
  "security.lock": [
    { element: "path", role: "ink", d: "M8 10.5v-3a4 4 0 0 1 8 0v3" },
    { element: "rect", role: "paper", x: 5, y: 10.5, width: 14, height: 10, rx: 2 },
    { element: "path", role: "accent", emphasis: true, d: "M12 14.5V17" },
  ],
  "security.key": [
    { element: "path", role: "accent", emphasis: true, d: "M11 12.7 19.5 4.2M16.5 7.2l2.3 2.3M14 9.7l1.8 1.8" },
    { element: "circle", role: "paper", cx: 8, cy: 15.5, r: 4 },
  ],
  "security.shield.check": [
    { element: "path", role: "paper", d: "M12 3 20 6v6c0 5-3.4 8-8 9-4.6-1-8-4-8-9V6z" },
    { element: "polyline", role: "accent", emphasis: true, points: "8.5,12 11,14.5 15.5,9.5" },
  ],
  "security.shield.alert": [
    { element: "path", role: "paper", d: "M12 3 20 6v6c0 5-3.4 8-8 9-4.6-1-8-4-8-9V6z" },
    { element: "path", role: "warm", emphasis: true, d: "M12 8v4.5M12 15.9h.01" },
  ],
  "security.shield.question": [
    { element: "path", role: "paper", d: "M12 3 20 6v6c0 5-3.4 8-8 9-4.6-1-8-4-8-9V6z" },
    { element: "path", role: "accent", emphasis: true, d: "M9.8 9.8a2.3 2.3 0 1 1 3.3 2.1c-.8.4-1.1.9-1.1 1.6M12 16.4h.01" },
  ],
  "security.shield.denied": [
    { element: "path", role: "paper", d: "M12 3 20 6v6c0 5-3.4 8-8 9-4.6-1-8-4-8-9V6z" },
    { element: "path", role: "warm", emphasis: true, d: "m9.5 9.5 5 5m0-5-5 5" },
  ],
  "billing.card": [
    { element: "rect", role: "paper", x: 2.5, y: 5, width: 19, height: 14, rx: 2 },
    { element: "path", role: "ink-soft", d: "M6 15h4" },
    { element: "path", role: "accent", emphasis: true, d: "M2.5 10h19" },
  ],
  "billing.coin": [
    { element: "circle", role: "paper", cx: 12, cy: 12, r: 8.5 },
    { element: "path", role: "accent", emphasis: true, d: "M14.8 9.2c-.5-.9-1.5-1.4-2.8-1.4-1.7 0-2.9.8-2.9 2.1 0 3 5.9 1.2 5.9 4.2 0 1.3-1.3 2.1-3 2.1-1.4 0-2.5-.5-3-1.4M12 6v1.8M12 16.2V18" },
  ],
  "billing.gift": [
    { element: "rect", role: "paper", x: 4.5, y: 11, width: 15, height: 9.5, rx: 1 },
    { element: "rect", role: "paper", x: 3.5, y: 7.5, width: 17, height: 3.5, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M12 7.5v13M12 7.5c-1.5-3.5-5.5-3.5-5-1 .3 1.2 2.5 1 5 1zm0 0c1.5-3.5 5.5-3.5 5-1-.3 1.2-2.5 1-5 1z" },
  ],
  "work.briefcase": [
    { element: "path", role: "ink", d: "M9 7.5v-2A1.5 1.5 0 0 1 10.5 4h3A1.5 1.5 0 0 1 15 5.5v2" },
    { element: "rect", role: "paper", x: 3, y: 7.5, width: 18, height: 12, rx: 1.8 },
    { element: "path", role: "accent", emphasis: true, d: "M3 12.5h18" },
  ],
  "learn.graduation": [
    { element: "path", role: "ink", d: "M6.5 11.8V16c1.5 1.5 3.3 2.2 5.5 2.2s4-.7 5.5-2.2v-4.2" },
    { element: "path", role: "paper", d: "M12 5 2.5 9.5 12 14l9.5-4.5z" },
    { element: "path", role: "accent", emphasis: true, d: "M21.5 9.5V15" },
  ],
  "social.heart": [
    { element: "path", role: "accent-soft-fill", d: "M12 20.5S3.5 15.5 3.5 9.3A4.6 4.6 0 0 1 12 7a4.6 4.6 0 0 1 8.5 2.3c0 6.2-8.5 11.2-8.5 11.2z" },
    { element: "path", role: "accent", emphasis: true, d: "M12 20.5S3.5 15.5 3.5 9.3A4.6 4.6 0 0 1 12 7a4.6 4.6 0 0 1 8.5 2.3c0 6.2-8.5 11.2-8.5 11.2z" },
  ],
  "social.star": [
    { element: "polygon", role: "accent-soft-fill", points: "12,3.6 14.3,9.4 20.6,9.8 15.7,13.8 17.3,19.9 12,16.5 6.7,19.9 8.3,13.8 3.4,9.8 9.7,9.4" },
    { element: "polygon", role: "accent", emphasis: true, points: "12,3.6 14.3,9.4 20.6,9.8 15.7,13.8 17.3,19.9 12,16.5 6.7,19.9 8.3,13.8 3.4,9.8 9.7,9.4" },
  ],
  "social.smile": [
    { element: "circle", role: "paper", cx: 12, cy: 12, r: 8.5 },
    { element: "path", role: "ink", emphasis: true, d: "M9 9.8v.01M15 9.8v.01" },
    { element: "path", role: "accent", emphasis: true, d: "M8.5 14a4.2 4.2 0 0 0 7 0" },
  ],
  "social.thumbsUp": [
    { element: "path", role: "paper", d: "M8 10.5l4-7a2.2 2.2 0 0 1 2.2 2.6l-.7 3.4h5a2 2 0 0 1 2 2.4l-1.2 6a2 2 0 0 1-2 1.6H8z" },
    { element: "rect", role: "paper", x: 3.5, y: 10.5, width: 4.5, height: 9.5, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M8 10.5V20" },
  ],
  "time.clock": [
    { element: "circle", role: "paper", cx: 12, cy: 12, r: 8.5 },
    { element: "path", role: "accent", emphasis: true, d: "M12 7.5V12l3 1.8" },
  ],
  "time.timer": [
    { element: "path", role: "ink", d: "M10 2.8h4" },
    { element: "circle", role: "paper", cx: 12, cy: 13.5, r: 7.5 },
    { element: "path", role: "accent", emphasis: true, d: "M12 13.5l2.8-3" },
  ],
  "calendar.date": [
    { element: "rect", role: "paper", x: 3.5, y: 5, width: 17, height: 15.5, rx: 2 },
    { element: "path", role: "ink", d: "M8 3v4M16 3v4" },
    { element: "path", role: "ink-soft", emphasis: true, d: "M8 14v.01M12 14v.01M16 14v.01M8 17.3v.01M12 17.3v.01" },
    { element: "path", role: "accent", emphasis: true, d: "M3.5 10h17" },
  ],
  "calendar.done": [
    { element: "rect", role: "paper", x: 3.5, y: 5, width: 17, height: 15.5, rx: 2 },
    { element: "path", role: "ink", d: "M8 3v4M16 3v4" },
    { element: "path", role: "ink-soft", d: "M3.5 10h17" },
    { element: "polyline", role: "accent", emphasis: true, points: "8.5,15 11,17.5 15.5,12.5" },
  ],
  "calendar.schedule": [
    { element: "path", role: "paper", d: "M20.5 11V7a2 2 0 0 0-2-2h-13a2 2 0 0 0-2 2v11.5a2 2 0 0 0 2 2H11" },
    { element: "path", role: "ink", d: "M8 3v4M16 3v4" },
    { element: "path", role: "ink-soft", d: "M3.5 10h17" },
    { element: "circle", role: "paper", cx: 17, cy: 17, r: 4.5 },
    { element: "path", role: "accent", emphasis: true, d: "M17 15v2l1.5 1" },
  ],
  "calendar.range": [
    { element: "rect", role: "paper", x: 3.5, y: 5, width: 17, height: 15.5, rx: 2 },
    { element: "path", role: "ink", d: "M8 3v4M16 3v4" },
    { element: "path", role: "ink-soft", d: "M3.5 10h17" },
    { element: "rect", role: "accent-soft-fill", x: 7, y: 13.2, width: 10, height: 3.2, rx: 1 },
    { element: "path", role: "accent", emphasis: true, d: "M7.5 14.8h9" },
  ],
  "ai.wand": [
    { element: "path", role: "paper", d: "m3.5 18.5 10.8-10.8 2 2L5.5 20.5z" },
    { element: "path", role: "accent", emphasis: true, d: "M17.5 3v3.5M15.75 4.75h3.5M20 11v2.5M18.75 12.25h2.5" },
    { element: "path", role: "ink", d: "M9 3.5v2M8 4.5h2" },
  ],
  "ai.bot": [
    { element: "path", role: "ink", d: "M12 4v3.5" },
    { element: "path", role: "ink-soft", d: "M2 13v3M22 13v3" },
    { element: "rect", role: "paper", x: 4, y: 7.5, width: 16, height: 12, rx: 2.5 },
    { element: "path", role: "accent", emphasis: true, d: "M9 12.5v2M15 12.5v2" },
    { element: "circle", role: "accent-fill", cx: 12, cy: 3.5, r: 0.6 },
  ],
  "ai.brain": [
    { element: "path", role: "paper", d: "M12 5.2A3.2 3.2 0 0 0 6.2 6 3.3 3.3 0 0 0 4 11a3.4 3.4 0 0 0 1.2 5.4A3.4 3.4 0 0 0 12 17.5z" },
    { element: "path", role: "paper", d: "M12 5.2A3.2 3.2 0 0 1 17.8 6 3.3 3.3 0 0 1 20 11a3.4 3.4 0 0 1-1.2 5.4A3.4 3.4 0 0 1 12 17.5z" },
    { element: "path", role: "ink-soft", d: "M8 10.5c1 0 2-.6 2-1.8M16 13.5c-1 0-2 .6-2 1.8" },
    { element: "path", role: "accent", emphasis: true, d: "M12 5.2v13.3" },
  ],
  "ai.orbit": [
    { element: "path", role: "ink-soft", d: "M3 12c0-2.5 4-4.5 9-4.5s9 2 9 4.5-4 4.5-9 4.5-9-2-9-4.5z" },
    { element: "circle", role: "paper", cx: 12, cy: 12, r: 3.2 },
    { element: "circle", role: "accent-fill", cx: 18.6, cy: 8.8, r: 1.2 },
  ],
  "action.launch": [
    { element: "path", role: "ink", d: "M7.6 12.5 5 15.5v3l3-1M16.4 12.5l2.6 3v3l-3-1" },
    { element: "path", role: "paper", d: "M12 3c3 2.2 4.5 5.3 4.5 9.3L15 15.5H9l-1.5-3.2C7.5 8.3 9 5.2 12 3z" },
    { element: "circle", role: "accent", emphasis: true, cx: 12, cy: 9.5, r: 1.6 },
    { element: "path", role: "warm", d: "M12 17.5V21" },
  ],
  "format.palette": [
    { element: "path", role: "paper", d: "M12 3.5a8.5 8.5 0 1 0 0 17c1.2 0 1.8-.9 1.8-1.8 0-1.4-1-1.7-1-2.9 0-1 .8-1.8 1.8-1.8h2.4a3.5 3.5 0 0 0 3.5-3.5c0-3.9-3.8-7-8.5-7z" },
    { element: "circle", role: "accent-fill", cx: 8.2, cy: 10.5, r: 0.9 },
    { element: "circle", role: "warm-fill", cx: 12, cy: 7.6, r: 0.9 },
    { element: "circle", role: "ink-soft", cx: 15.8, cy: 10, r: 0.9 },
  ],
  "format.eyedropper": [
    { element: "path", role: "paper", d: "m14 6 4 4-8.5 8.5-3.3.8-1.2 1.2-1.5-1.5 1.2-1.2.8-3.3z" },
    { element: "path", role: "accent", emphasis: true, d: "M12.5 4.5l7 7M15.5 7.5l2.2-2.2a2 2 0 0 1 2.9 2.9l-2.2 2.2" },
  ],
  "format.painter": [
    ...FORMAT_BRUSH,
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.2,
      y: 3,
      width: 9.4,
      height: 4.2,
      rx: 1,
    },
    { element: "path", role: "accent", d: "M4.6 5.1h6.6" },
  ],
  "format.clear": [
    {
      element: "path",
      role: "ink-soft",
      d: "M3.5 5.2h13M3.5 9.4h10M3.5 13.6h8",
    },
    {
      element: "path",
      role: "warm-fill",
      d: "m10.8 15.2 5.5-5.5 4 4-5.5 5.5h-4z",
    },
    { element: "path", role: "accent", emphasis: true, d: "M4 20h16" },
  ],
  "format.lineStyle": [
    {
      element: "path",
      role: "ink-soft",
      d: "M3 5.5h18M3 12h18",
      strokeDasharray: "4 2",
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M3 18.5h18",
      strokeDasharray: "1 2",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M3 12h18",
      strokeDasharray: "4 2",
    },
  ],
  "format.lineColor": [
    { element: "path", role: "ink", emphasis: true, d: "M3 17h12" },
    {
      element: "path",
      role: "accent",
      d: "M21 14.3c0 1.6-1.2 2.8-2.7 2.8s-2.7-1.2-2.7-2.8c0-1.4 2.7-4.7 2.7-4.7s2.7 3.3 2.7 4.7z",
    },
    {
      element: "rect",
      role: "accent-fill",
      x: 3,
      y: 19.2,
      width: 12,
      height: 1.8,
      rx: 0.5,
    },
  ],
  "format.lineWeight": [
    { element: "path", role: "ink-soft", d: "M3 5h18" },
    { element: "path", role: "ink", emphasis: true, d: "M3 11.5h18" },
    { element: "path", role: "accent", emphasis: true, d: "M3 18h18M3 20h18" },
  ],
  "text.fontSize.increase": [
    { element: "path", role: "ink", d: "M4 7h9M4 11h7M4 15h5" },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M18.5 18V6M15.8 8.7 18.5 6l2.7 2.7M15.5 18.5h6",
    },
  ],
  "text.fontSize.decrease": [
    { element: "path", role: "ink", d: "M4 7h9M4 11h7M4 15h5" },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M18.5 6v12M15.8 15.3l2.7 2.7 2.7-2.7M15.5 5.5h6",
    },
  ],
  "text.bold": [
    {
      element: "path",
      role: "ink",
      emphasis: true,
      d: "M6 4h6a4 4 0 0 1 0 8H6m0 0h7a4 4 0 0 1 0 8H6",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M6 4v16",
    },
  ],
  "text.italic": [
    {
      element: "path",
      role: "ink",
      emphasis: true,
      d: "M10 4h9M5 20h9",
    },
    { element: "path", role: "accent", emphasis: true, d: "M15 4 9 20" },
  ],
  "text.underline": [
    { element: "path", role: "ink", d: "M6 4v8a6 6 0 0 0 12 0V4" },
    {
      element: "line",
      role: "accent",
      emphasis: true,
      x1: 4,
      y1: 22,
      x2: 20,
      y2: 22,
    },
  ],
  "text.strikethrough": [
    { element: "path", role: "ink", d: "M10 10v8M10 12a3 3 0 1 0 0 4M14 5v13M14 12a3 3 0 1 1 0 4" },
    {
      element: "line",
      role: "accent",
      emphasis: true,
      x1: 2.5,
      y1: 12,
      x2: 21.5,
      y2: 12,
    },
  ],
  "text.subscript": [
    { element: "path", role: "ink", d: "M4 6l9 10M13 6 4 16" },
    {
      element: "path",
      role: "accent",
      d: "M16 17a2.5 2.5 0 0 1 5 0c0 1.5-2 2.5-5 4h5",
    },
  ],
  "text.superscript": [
    { element: "path", role: "ink", d: "M4 9l9 10M13 9 4 19" },
    {
      element: "path",
      role: "accent",
      d: "M16 5a2.5 2.5 0 0 1 5 0c0 1.5-2 2.5-5 4h5",
    },
  ],
  "text.highlight": [
    { element: "path", role: "ink", d: "M4 5h15M4 9.5h12M4 14h15" },
    {
      element: "rect",
      role: "warm-fill",
      x: 3,
      y: 16.8,
      width: 18,
      height: 3.2,
      rx: 0.6,
    },
    { element: "path", role: "warm", emphasis: true, d: "M4 18.4h16" },
  ],
  // Office's Font Color: the letter A over the colour it applies. Three text
  // strokes read as an alignment or a highlight, not as the colour of type.
  "text.fontColor": [
    { element: "path", role: "ink", d: "M5.2 15.8 10.6 3.6h1.8l5.4 12.2M7.4 11.2h8.2" },
    {
      element: "rect",
      role: "accent-fill",
      x: 3,
      y: 18,
      width: 18,
      height: 2.6,
      rx: 0.6,
    },
  ],
  "text.spacing.character": [
    { element: "path", role: "ink", d: "M3.5 3.75v8M7 3.75v8M17 3.75v8M20.5 3.75v8" },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M5 17.75h14M7.5 15.25 5 17.75l2.5 2.5M16.5 15.25 19 17.75l-2.5 2.5",
    },
  ],
  "text.input.keyboard": [
    {
      element: "rect",
      role: "paper",
      x: 2.5,
      y: 5,
      width: 19,
      height: 14,
      rx: 2,
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M5 8.2h2M9 8.2h2M13 8.2h2M17 8.2h2M5 11.7h2M9 11.7h2M13 11.7h2M17 11.7h2",
    },
    {
      element: "line",
      role: "accent",
      emphasis: true,
      x1: 7.5,
      y1: 15.5,
      x2: 16.5,
      y2: 15.5,
    },
  ],
  "paragraph.align.left": paragraphGeometry(
    "M5 5h16M5 9.7h11M5 14.3h16M5 19h9",
    "M3 4v16",
  ),
  "paragraph.align.center": paragraphGeometry(
    "M3 5h18M6 9.7h12M3 14.3h18M6 19h12",
    "M12 3v18",
  ),
  "paragraph.align.right": paragraphGeometry(
    "M3 5h16M8 9.7h11M3 14.3h16M10 19h9",
    "M21 4v16",
  ),
  // one unit higher than its siblings: its accent is a rule under the lines,
  // not a margin beside them, so this centres it on the canvas as they are
  "paragraph.align.justify": paragraphGeometry(
    "M3 4h18M3 8.7h18M3 13.3h18M3 18h18",
    "M3 20h18",
  ),
  "paragraph.align.distributed": paragraphGeometry(
    "M7 6h10M6 10h12M7 14h10M6 18h12",
    "M3 4v16M21 4v16M3 12h3M18 12h3M5 10l-2 2 2 2M19 10l2 2-2 2",
  ),
  "paragraph.indent.decrease": [
    { element: "path", role: "ink", d: "M10 5h11M10 10h11M10 15h11M10 20h8" },
    {
      element: "polyline",
      role: "accent",
      emphasis: true,
      points: "8,9 4,12 8,15",
    },
  ],
  "paragraph.indent.increase": [
    { element: "path", role: "ink", d: "M10 5h11M10 10h11M10 15h11M10 20h8" },
    {
      element: "polyline",
      role: "accent",
      emphasis: true,
      points: "4,9 8,12 4,15",
    },
  ],
  "paragraph.indent.firstLine": [
    { element: "path", role: "ink", d: "M9 5h12M3 10h18M3 15h18M3 20h14" },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M3 5h4M5 3l2 2-2 2",
    },
  ],
  "paragraph.spacing.line": [
    {
      element: "path",
      role: "ink",
      d: "M10 5h11M10 9.7h11M10 14.3h11M10 19h11",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M5 4v16M2.5 6.5 5 4l2.5 2.5M2.5 17.5 5 20l2.5-2.5",
    },
  ],
  // Independently authored for the shared quick-overflow control. The
  // round-capped micro-segments remain crisp at the frozen 16px render size.
  "overflow.more": [
    {
      element: "path",
      role: "ink",
      emphasis: true,
      d: "M5.2 12h.01M12 12h.01M18.8 12h.01",
    },
  ],
  // Independently authored for the shared ribbon's physical edge-scroll
  // controls. Separate non-mirrored glyphs keep action and direction aligned
  // in both LTR and RTL layouts.
  "ribbon.scroll.left": [
    {
      element: "path",
      role: "ink",
      emphasis: true,
      d: "M14.6 5.8 8.4 12l6.2 6.2",
    },
  ],
  "ribbon.scroll.right": [
    {
      element: "path",
      role: "ink",
      emphasis: true,
      d: "M9.4 5.8 15.6 12l-6.2 6.2",
    },
  ],
  // Independently authored for compact menu affordances in the shared font
  // picker. The centered geometry stays legible at the frozen 12px size.
  "menu.chevron.down": [
    {
      element: "path",
      role: "ink",
      emphasis: true,
      d: "M5.8 9.2 12 15.4l6.2-6.2",
    },
  ],
  // Independently authored for the shared font picker's selected menu row.
  // The asymmetric rise retains a clear check silhouette at 13px.
  "selection.check": [
    {
      element: "path",
      role: "ink",
      emphasis: true,
      d: "M4.8 12.3 9.6 17l9.6-10.2",
    },
  ],
  // Shared Review/View geometry. These metaphors are editor-neutral so Docs,
  // Sheets, Slides, and PDF can keep command state in adapters while sharing
  // one visual language.
  "review.spellcheck": [
    {
      element: "path",
      role: "ink",
      emphasis: true,
      d: "M4 16 9 4l5 12M6 11h6",
    },
    {
      element: "polyline",
      role: "accent",
      emphasis: true,
      points: "13,18 15.5,20.5 21,14",
    },
  ],
  "review.wordCount": [
    { element: "path", role: "paper", d: "M4 3h16v18H4z" },
    { element: "path", role: "ink-soft", d: "M7 7h10M7 11h7M7 15h5" },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M15 14h4M15 18h4M17 12v8",
    },
  ],
  "review.accessibility": [
    { element: "circle", role: "accent-fill", cx: 12, cy: 4, r: 2 },
    {
      element: "path",
      role: "ink",
      d: "M4 8h16M12 8v5M8 21l4-8 4 8M7 10l5 3 5-3",
    },
  ],
  "review.translate": [
    {
      element: "path",
      role: "ink",
      d: "M3 5h9M7.5 3v2M5 7c1.4 3 3.5 5.2 6.5 6.7M10.5 7C9 10.5 6.7 13 3.5 15",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "m13 20 3.5-9 3.5 9M14.5 17h4",
    },
  ],
  "review.language": [
    { element: "circle", role: "paper", cx: 9.5, cy: 9.5, r: 7 },
    { element: "path", role: "ink-soft", d: "M2.5 9.5h14M9.5 2.5c2.2 1.9 3.3 4.2 3.3 7s-1.1 5.1-3.3 7c-2.2-1.9-3.3-4.2-3.3-7s1.1-5.1 3.3-7z" },
    { element: "rect", role: "paper", x: 12.5, y: 12.5, width: 9.5, height: 9.5, rx: 2.2 },
    { element: "path", role: "accent", emphasis: true, d: "m15 20 2.25-5.6L19.5 20M15.85 17.9h2.8" },
  ],
  "review.trackChanges": [
    { element: "path", role: "paper", d: "M4 3h14v18H4z" },
    { element: "path", role: "ink-soft", d: "M7 7h8M7 11h8M7 15h5" },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "m13 18 7-7 2 2-7 7-3 1z",
    },
  ],
  "review.markup": [
    { element: "path", role: "paper", d: "M4 3h14v18H4z" },
    { element: "path", role: "ink-soft", d: "M7 7h8M7 11h8M7 15h5" },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M15 13h7l-2.7 3v3l-1.6 1v-4z",
    },
  ],
  "review.change.accept": [
    { element: "path", role: "paper", d: "M4 3h14v18H4z" },
    { element: "path", role: "ink-soft", d: "M7 7h8M7 11h6M7 15h5" },
    {
      element: "polyline",
      role: "accent",
      emphasis: true,
      points: "13,17 15.5,19.5 21,14",
    },
  ],
  "review.change.reject": [
    { element: "path", role: "paper", d: "M4 3h14v18H4z" },
    { element: "path", role: "ink-soft", d: "M7 7h8M7 11h6M7 15h5" },
    { element: "path", role: "warm", emphasis: true, d: "m15 15 5 5m0-5-5 5" },
  ],
  "review.comment": [
    { element: "path", role: "paper", d: "M3 4h18v13H9l-4 4v-4H3z" },
    { element: "path", role: "ink-soft", d: "M7 8h10M7 12h7" },
    { element: "path", role: "accent", emphasis: true, d: "M18 13v6M15 16h6" },
  ],
  "review.comment.delete": [
    { element: "path", role: "paper", d: "M3 4h18v13H9l-4 4v-4H3z" },
    { element: "path", role: "ink-soft", d: "M7 8h10M7 12h7" },
    { element: "path", role: "warm", emphasis: true, d: "M15 16h6" },
  ],
  "review.compare": [
    { element: "path", role: "paper", d: "M3 4h7v16H3zM14 4h7v16h-7z" },
    {
      element: "path",
      role: "ink-soft",
      d: "M5.5 8h2M5.5 12h2M16.5 12h2M16.5 16h2",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M8 16h8m-2.5-2.5L16 16l-2.5 2.5M16 8H8m2.5-2.5L8 8l2.5 2.5",
    },
  ],
  "review.protect": [
    {
      element: "path",
      role: "paper",
      d: "M12 3 20 6v6c0 5-3.4 8-8 9-4.6-1-8-4-8-9V6z",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M9 12v-2.2a3 3 0 0 1 6 0V12m-7 0h8v6H8z",
    },
  ],
  "review.readAloud": [
    { element: "path", role: "paper", d: "M3.5 3.5h10v17h-10z" },
    {
      element: "path",
      role: "ink-soft",
      d: "M6.5 7.5h4M6.5 11h4M6.5 14.5h2.5",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M14.5 10h2.2l3-2.5v9l-3-2.5h-2.2zM21 9.5c.8.7 1.2 1.5 1.2 2.5s-.4 1.8-1.2 2.5",
    },
  ],
  "view.document": [
    { element: "path", role: "paper", d: "M5 3h14v18H5z" },
    { element: "path", role: "ink-soft", d: "M8 7h8M8 11h8M8 15h5" },
    { element: "path", role: "accent", emphasis: true, d: "M5 6h14" },
  ],
  "view.reading": [
    {
      element: "path",
      role: "paper",
      d: "M3 5c3-1.5 6-1.5 9 1v14c-3-2-6-2-9-1zM21 5c-3-1.5-6-1.5-9 1v14c3-2 6-2 9-1z",
    },
    { element: "path", role: "ink-soft", d: "M6 9h3M6 12h3M15 9h3M15 12h3" },
    { element: "path", role: "accent", d: "M12 6v14" },
  ],
  "view.immersiveReader": [
    {
      element: "path",
      role: "paper",
      d: "M3 5c3-1.5 6-1.5 9 1v14c-3-2-6-2-9-1zM21 5c-3-1.5-6-1.5-9 1v14c3-2 6-2 9-1z",
    },
    { element: "path", role: "ink-soft", d: "M6 9h3M6 12h3M15 9h3M15 12h3" },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M2 3v5M2 3h5M22 3v5M22 3h-5M2 21v-5M2 21h5M22 21v-5M22 21h-5",
    },
  ],
  "view.navigationPane": [
    { element: "path", role: "paper", d: "M3 3h18v18H3z" },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3,
      y: 3,
      width: 6,
      height: 18,
    },
    { element: "path", role: "accent", emphasis: true, d: "M9 3v18" },
    {
      element: "path",
      role: "ink-soft",
      d: "M5 7h2M5 11h2M12 7h6M12 11h6M12 15h4",
    },
  ],
  "view.zoom": [
    { element: "circle", role: "paper", cx: 10, cy: 10, r: 6 },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "m14.5 14.5 6 6M7 10h6M10 7v6",
    },
  ],
  // Word's ruler: a horizontal rule with inch ticks and the indent marker
  // riding under it.
  "view.ruler": [
    {
      element: "rect",
      role: "paper",
      x: 2.5,
      y: 5.9,
      width: 19,
      height: 8,
      rx: 1.2,
    },
    {
      element: "path",
      role: "ink",
      d: "M5.5 5.9v3.4M8.5 5.9v2.2M11.5 5.9v3.4M14.5 5.9v2.2M17.5 5.9v3.4M20.5 5.9v2.2",
    },
    {
      element: "polygon",
      role: "accent-fill",
      emphasis: true,
      points: "6.2,18.1 10.2,18.1 8.2,14.9",
    },
    { element: "path", role: "accent", emphasis: true, d: "M8.2 13.9v1" },
  ],
  // A page whose text is crossed by a faint diagonal band: the watermark
  // reads behind the text, so the band is a soft fill under the lines.
  "page.watermark": [
    {
      element: "rect",
      role: "paper",
      x: 5,
      y: 3,
      width: 14,
      height: 18,
      rx: 1.3,
    },
    {
      element: "polygon",
      role: "accent-soft-fill",
      points: "5,14.6 17.4,4.2 19,6.4 6.6,16.8",
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M8 7.5h4.2M8 10.5h8M8 13.5h8M8 16.5h8",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "m7.2 16.2 11-9.2",
      strokeDasharray: "2.4 1.6",
    },
  ],
  // Square wrapping: the picture sits in the text column and the lines
  // break around it on both sides.
  "object.textWrap": [
    {
      element: "path",
      role: "ink-soft",
      d: "M3.5 4.5h17M3.5 19.5h17",
    },
    {
      element: "path",
      role: "ink",
      d: "M3.5 8.5h3.6M3.5 12h3.6M3.5 15.5h3.6M16.9 8.5h3.6M16.9 12h3.6M16.9 15.5h3.6",
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 9.3,
      y: 7.2,
      width: 5.4,
      height: 9.6,
      rx: 0.9,
    },
    {
      element: "rect",
      role: "accent",
      emphasis: true,
      x: 9.3,
      y: 7.2,
      width: 5.4,
      height: 9.6,
      rx: 0.9,
    },
    {
      element: "polyline",
      role: "accent",
      points: "9.9,15.4 11.3,13.4 12.3,14.6 14.1,12.2",
    },
  ],
  "view.pageGuides": [
    { element: "path", role: "paper", d: "M4 3h16v18H4z" },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M8 3v18M16 3v18M4 8h16M4 16h16",
      strokeDasharray: "2 2",
    },
  ],
  "view.darkMode": [
    {
      element: "path",
      role: "accent-soft-fill",
      d: "M18 17.5A8.5 8.5 0 0 1 7 6a8.5 8.5 0 1 0 11 11.5Z",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M18 17.5A8.5 8.5 0 0 1 7 6a8.5 8.5 0 1 0 11 11.5Z",
    },
  ],
  "view.window": [
    {
      element: "rect",
      role: "ink-soft",
      x: 3,
      y: 6.5,
      width: 13,
      height: 13,
      rx: 1.5,
    },
    {
      element: "rect",
      role: "paper",
      x: 8,
      y: 4.5,
      width: 13,
      height: 13,
      rx: 1.5,
    },
    { element: "path", role: "accent", emphasis: true, d: "M8 8.5h13" },
  ],
  "view.webLayout": [
    { element: "path", role: "paper", d: "M4 3h13v18H4z" },
    { element: "path", role: "ink-soft", d: "M7 7h7M7 10.5h7M7 14h3.5" },
    { element: "path", role: "accent", emphasis: true, d: "M17.5 12.5a4.5 4.5 0 1 0 0 9a4.5 4.5 0 1 0 0-9zM13 17h9M17.5 12.5c1.4 1.3 2.1 2.8 2.1 4.5s-.7 3.2-2.1 4.5c-1.4-1.3-2.1-2.8-2.1-4.5s.7-3.2 2.1-4.5z" },
  ],
  "view.focus": [
    { element: "path", role: "ink", d: "M4 8V4h4M16 4h4v4M20 16v4h-4M8 20H4v-4" },
    { element: "path", role: "accent", emphasis: true, d: "M8.5 9.5h7M8.5 12.5h7M8.5 15.5h4" },
  ],
  "view.onePage": [
    { element: "path", role: "paper", d: "M6 3h12v18H6z" },
    { element: "path", role: "ink-soft", d: "M9 11h6M9 14.5h6M9 18h3.5" },
    { element: "path", role: "accent", emphasis: true, d: "M9 7.5h6" },
  ],
  "view.multiplePages": [
    { element: "path", role: "paper", d: "M2.5 5h8.5v14H2.5zM13 5h8.5v14H13z" },
    { element: "path", role: "ink-soft", d: "M5 12h3.5M5 15h3.5M15.5 12H19M15.5 15H19" },
    { element: "path", role: "accent", emphasis: true, d: "M5 9h3.5M15.5 9H19" },
  ],
  "view.pageWidth": [
    { element: "path", role: "paper", d: "M5 3h14v12H5z" },
    { element: "path", role: "ink-soft", d: "M8 7h8M8 10.5h8" },
    { element: "path", role: "accent", emphasis: true, d: "M3.5 19.5h17M6.5 17l-3 2.5 3 2.5M17.5 17l3 2.5-3 2.5" },
  ],
  "help.search": [
    { element: "circle", role: "paper", cx: 12, cy: 12, r: 8.6 },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M9.3 9.2a2.9 2.9 0 1 1 4.1 2.7c-1 .5-1.4 1.1-1.4 2.1",
    },
    { element: "circle", role: "ink", cx: 12, cy: 17.3, r: 0.7 },
  ],
  "help.support": [
    {
      element: "path",
      role: "ink",
      d: "M4.2 13v-2a7.8 7.8 0 0 1 15.6 0v2",
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3,
      y: 11.5,
      width: 4,
      height: 6.2,
      rx: 1.2,
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 17,
      y: 11.5,
      width: 4,
      height: 6.2,
      rx: 1.2,
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M19 17.7v.8a2 2 0 0 1-2 2h-3",
    },
  ],
  "help.feedback": [
    {
      element: "path",
      role: "paper",
      d: "M3 4h18v13H10l-4.8 3.2V17H3z",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M7.2 10.6h.01M12 10.6h.01M16.8 10.6h.01",
    },
  ],
  "help.whatsNew": [
    {
      element: "path",
      role: "paper",
      d: "M3.2 10.2v3.6h3.3l8.2 4.2V6L6.5 10.2z",
    },
    { element: "path", role: "ink-soft", d: "m6.5 13.8 1.1 6h3.1l-.9-4.3" },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M18 7.2 20.6 5M18.8 11h3.4M18 14.8l2.6 2.2",
    },
  ],
  "font.search": [
    { element: "circle", role: "ink", cx: 10, cy: 10, r: 5.7 },
    {
      element: "path",
      role: "ink",
      emphasis: true,
      d: "m14.3 14.3 5 5",
    },
  ],
  "font.pin": [
    {
      element: "path",
      role: "ink",
      d: "m8 3 8 0-1.2 5 3.2 3v2H6v-2l3.2-3zM12 13v8",
    },
  ],
  "font.unpin": [
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "m8 3 8 0-1.2 5 3.2 3v2H6v-2l3.2-3zM12 13v8",
    },
  ],
  "font.browse": [
    {
      element: "path",
      role: "paper",
      d: "M3 6h7l2 2h9v11H3z",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "m8 17 2.5-6h1.7l2.5 6M9 14.8h4.7",
    },
  ],
  "font.variant.chevron": [
    {
      element: "path",
      role: "ink",
      emphasis: true,
      d: "m9.2 5.8 6.2 6.2-6.2 6.2",
    },
  ],
  // Original MRAI Draw geometry. The silhouettes share the same 24px safe
  // area and role tokens as Insert/Help, so Draw does not introduce a second
  // icon family or embed vendor artwork.
  "draw.select": [
    {
      element: "path",
      role: "ink",
      emphasis: true,
      d: "M5 3.2v14.6l3.8-3.6 2.7 6.2 2.5-1.1-2.7-6.1h5.4z",
    },
    {
      element: "rect",
      role: "accent",
      x: 15.7,
      y: 3.6,
      width: 4.6,
      height: 4.6,
      rx: 0.7,
    },
  ],
  "draw.mode": [
    {
      element: "path",
      role: "paper",
      d: "m5 17.6 1.1-4.4L15.8 3.5l4.7 4.7-9.7 9.7z",
    },
    {
      element: "path",
      role: "ink-soft",
      d: "m13.8 5.5 4.7 4.7M6.1 13.2l4.7 4.7",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M3.4 21c3.3-2 6.2.8 9.2-.7 2.2-1.1 4.1-.9 7.8.2",
    },
  ],
  "draw.lasso": [
    {
      element: "path",
      role: "ink",
      d: "M19.7 10.5c0 4.1-3.4 7-8.2 7s-8.2-2.9-8.2-7 3.4-7 8.2-7 8.2 2.9 8.2 7Z",
      strokeDasharray: "2.2 2.1",
    },
    {
      element: "path",
      role: "ink-soft",
      d: "M11.5 17.5c.2 2.2 1.6 3.4 4 3.4",
    },
    { element: "circle", role: "accent-fill", cx: 11.5, cy: 10.5, r: 1.5 },
  ],
  "draw.laser": [
    { element: "path", role: "ink", emphasis: true, d: "m3.6 18.2 8.3-8.3 2.2 2.2-8.3 8.3a1.55 1.55 0 0 1-2.2-2.2Z" },
    { element: "path", role: "ink-soft", d: "m9.3 12.5 2.2 2.2" },
    { element: "circle", role: "accent-fill", cx: 17.5, cy: 6.5, r: 2 },
    { element: "path", role: "accent", d: "M17.5 2.3v1.2M21.7 6.5h-1.2M20.5 3.5l-.85.85M14.5 3.5l.85.85M20.5 9.5l-.85-.85" },
  ],
  "draw.tape": [
    { element: "circle", role: "paper", cx: 8.5, cy: 8.8, r: 5.3 },
    { element: "circle", role: "ink-soft", cx: 8.5, cy: 8.8, r: 2 },
    { element: "path", role: "accent-soft-fill", d: "M8.5 14.1H20.6l-1.2 1.7 1.2 1.7H8.5Z" },
    { element: "path", role: "accent", emphasis: true, d: "M8.5 14.1H20.6l-1.2 1.7 1.2 1.7H8.5M12.5 14.1l-1.4 3.4M16.5 14.1l-1.4 3.4" },
    { element: "path", role: "ink-soft", d: "M3.5 20.8h17" },
  ],
  "draw.stickers": [
    { element: "path", role: "paper", d: "M6 3.5h12a2.5 2.5 0 0 1 2.5 2.5v8.5l-6 6H6A2.5 2.5 0 0 1 3.5 18V6A2.5 2.5 0 0 1 6 3.5Z" },
    { element: "path", role: "ink-soft", d: "M20.5 14.5H17a2.5 2.5 0 0 0-2.5 2.5v3.5" },
    { element: "path", role: "accent-fill", d: "m11 7.4 1.25 2.55 2.8.4-2.03 1.98.48 2.8L11 13.8l-2.5 1.33.48-2.8-2.03-1.98 2.8-.4Z" },
  ],
  "draw.eraser.stroke": [
    {
      element: "polygon",
      role: "accent-soft-fill",
      points: "4,15.5 13.7,5.8 20,12.1 12.1,20 8.5,20",
    },
    {
      element: "path",
      role: "ink",
      emphasis: true,
      d: "m4 15.5 9.7-9.7 6.3 6.3-7.9 7.9H8.5zM9.3 10.2l6.3 6.3",
    },
    { element: "path", role: "accent", d: "M3.2 21h17.6" },
  ],
  "draw.pen": [
    {
      element: "path",
      role: "paper",
      d: "m5.2 16.7 1.2-4.5L15.8 2.8l5.4 5.4-9.4 9.4-4.5 1.2z",
    },
    {
      element: "path",
      role: "ink-soft",
      d: "m13.6 5 5.4 5.4M6.4 12.2l5.4 5.4",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M3.4 21h17.2",
    },
  ],
  "draw.highlighter": [
    {
      element: "path",
      role: "paper",
      d: "m5.1 14.4 9.6-9.6 4.5 4.5-9.6 9.6H5.1z",
    },
    {
      element: "path",
      role: "ink-soft",
      d: "m12.6 6.9 4.5 4.5M5.1 14.4l4.5 4.5",
    },
    {
      element: "rect",
      role: "accent-soft-fill",
      x: 3.2,
      y: 20,
      width: 17.6,
      height: 2,
      rx: 1,
    },
  ],
  "draw.actionPen": [
    {
      element: "path",
      role: "paper",
      d: "m4.2 17.8 1.1-4.2 8.7-8.7 5.1 5.1-8.7 8.7-4.2 1.1z",
    },
    {
      element: "path",
      role: "ink-soft",
      d: "m12 6.9 5.1 5.1M5.3 13.6l5.1 5.1",
    },
    {
      element: "path",
      role: "accent",
      emphasis: true,
      d: "M18.6 2.5v3.3M17 4.1h3.3M20.1 7.1l1.5 1.5",
    },
    { element: "path", role: "warm", d: "M3.2 22h11.6" },
  ],
} satisfies Record<OfficeSemanticIconId, OfficeIconGeometry>);

export const OfficeIconGlyph = memo(function OfficeIconGlyph({
  semanticId,
}: {
  semanticId: OfficeSemanticIconId;
}): ReactNode {
  const geometry: OfficeIconGeometry = OFFICE_ICON_GEOMETRY[semanticId];
  return geometry.map((primitive, index) => {
    const { element, role, emphasis, ...geometry } = primitive;
    return createElement(element, {
      ...geometry,
      key: `${semanticId}-${index}`,
      "data-office-icon-role": role,
      "data-office-icon-emphasis": emphasis ? "true" : undefined,
      strokeWidth: emphasis
        ? OFFICE_ICON_EMPHASIS_STROKE_WIDTH
        : OFFICE_ICON_STROKE_WIDTH,
    });
  });
});
