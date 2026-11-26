/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** OS/2 win ascent, win descent (em) and the hhea ascent's share of the
 * hhea line (where the canvas puts the em box's top above the baseline), read
 * from Office's own font files (PowerPoint for Mac's DFonts, macOS's
 * Supplemental faces). Metric clones carry their face's box. */
const FACE_BOX: Readonly<Record<string, readonly [number, number, number]>> = {
  calibri: [0.9521, 0.2686, 0.78],
  "calibri light": [0.9521, 0.2686, 0.78],
  carlito: [0.9521, 0.2686, 0.78],
  arial: [0.9053, 0.2119, 0.8103],
  arimo: [0.9053, 0.2119, 0.8103],
  "liberation sans": [0.9053, 0.2119, 0.8103],
  "times new roman": [0.8911, 0.2163, 0.8047],
  tinos: [0.8911, 0.2163, 0.8047],
  "liberation serif": [0.8911, 0.2163, 0.8047],
  "trebuchet ms": [0.939, 0.2222, 0.8087],
  verdana: [1.0054, 0.21, 0.8272],
  tahoma: [1.0005, 0.2065, 0.8289],
  "century gothic": [0.9712, 0.2202, 0.8204],
  "gill sans mt": [0.9292, 0.2305, 0.8013],
  "gill sans": [0.918, 0.2305, 0.7993],
  corbel: [0.9521, 0.2686, 0.7437],
  candara: [0.9521, 0.2686, 0.7246],
  cambria: [0.9502, 0.2222, 0.8105],
  caladea: [0.9502, 0.2222, 0.8105],
  constantia: [0.9521, 0.2686, 0.751],
  georgia: [0.917, 0.2192, 0.807],
  helvetica: [0.9502, 0.2251, 0.77],
  "helvetica neue": [0.952, 0.213, 0.8172],
  garamond: [0.8618, 0.2632, 0.7661],
  "franklin gothic book": [0.9165, 0.2173, 0.8084],
  "franklin gothic medium": [0.9165, 0.2173, 0.8084],
  aptos: [1.0098, 0.2749, 0.7692],
  "arial narrow": [0.9219, 0.2104, 0.8153],
  "arial black": [1.1006, 0.3096, 0.7805],
  "tw cen mt": [0.856, 0.2329, 0.7861],
  "lucida sans": [0.9277, 0.2109, 0.8209],
  "book antiqua": [0.9604, 0.2822, 0.7659],
  "palatino linotype": [1.0498, 0.2993, 0.7781],
  century: [0.9858, 0.2158, 0.8201],
  consolas: [0.9199, 0.251, 0.7427],
  "courier new": [0.8325, 0.3003, 0.7349],
  cousine: [0.8325, 0.3003, 0.7349],
  rockwell: [0.9458, 0.2285, 0.8054],
  impact: [1.0088, 0.2109, 0.8271],
  "comic sans ms": [1.1021, 0.2915, 0.7908],
  "avenir next": [1.0, 0.366, 0.7321],
  "microsoft yahei": [1.0581, 0.2617, 0.8017],
  "microsoft jhenghei": [1.0757, 0.2544, 0.8087],
  dengxian: [0.8101, 0.2319, 0.7774],
  simsun: [0.8594, 0.1406, 0.8594],
  "ms gothic": [0.8594, 0.1406, 0.8594],
  "ms pgothic": [0.8594, 0.1406, 0.8594],
  "yu gothic": [0.9849, 0.3022, 0.7984],
  meiryo: [1.0601, 0.4399, 0.7067],
  "malgun gothic": [1.0884, 0.2417, 0.8183],
};

const DEFAULT_FACE = FACE_BOX.calibri;

/** The table's box for a CSS family list: its first face the table knows. */
function faceBox(family: unknown): readonly [number, number, number] {
  if (typeof family !== "string") return DEFAULT_FACE;
  for (const raw of family.split(",")) {
    const name = raw.replace(/["']/g, "").trim().toLowerCase();
    const box = FACE_BOX[name];
    if (box) return box;
  }
  return DEFAULT_FACE;
}

/**
 * Where Windows Excel stands a legend entry's baseline under its key's
 * middle, in em of the entry's face: (win ascent - 3 x win descent + 1) / 4
 * (lookwin lw124/fb, R480-X04's chart: Calibri 0.287 at 9 and 14 pt, Arial
 * 0.323, Times New Roman 0.319; this 0.286 / 0.317 / 0.311). An SVG's
 * central baseline stood the words 0.35 em (the API's PDF) or the face's
 * box middle (a browser's) under it: Calibri 0.57 pt low at 9 pt.
 */
export function officeLegendBaselineEm(family: unknown): number {
  const [ascent, descent] = faceBox(family);
  return (ascent - 3 * descent + 1) / 4;
}

/** A face's line (OS/2 win ascent + descent), in em: Calibri 1.221, Arial 1.117. */
export function officeFaceLineEm(family: unknown): number {
  const [ascent, descent] = faceBox(family);
  return ascent + descent;
}

/** The faces whose hhea line gap runs past their win box (GDI's external
 * leading, em), read from Windows' font files (xlsx backlog
 * bl12/bin/extlead.py; Times New Roman's line gap is the Windows face's 87
 * units, not 0 as in Office for Mac's copy); every other face in the table
 * has none. Metric clones carry their face's. */
const FACE_EXTERNAL_LEADING: Readonly<Record<string, number>> = {
  arial: 0.0327,
  arimo: 0.0327,
  "liberation sans": 0.0327,
  "times new roman": 0.0425,
  tinos: 0.0425,
  "liberation serif": 0.0425,
  "arial narrow": 0.0151,
  "century gothic": 0.0347,
  "lucida sans": 0.0391,
};

/**
 * A face's line with its external leading, in em (Arial 1.150, Calibri
 * 1.221): how far apart Windows Excel sets a worksheet chart's wrapped
 * category names (xlsx backlog bl10/x21, bl12/fonts, 9 pt names: Arial and
 * Times New Roman 10.34-10.35 pt apart, Calibri 10.97-11.0, Century Gothic
 * 11.01-11.04, Verdana 10.94-10.95).
 */
export function officeFaceLeadedLineEm(family: unknown): number {
  let leading = 0;
  if (typeof family === "string") {
    for (const raw of family.split(",")) {
      const name = raw.replace(/["']/g, "").trim().toLowerCase();
      if (FACE_BOX[name]) {
        leading = FACE_EXTERNAL_LEADING[name] ?? 0;
        break;
      }
    }
  }
  return officeFaceLineEm(family) + leading;
}

/** A major tick mark's length, in the unit of `size` (the axis's label size). */
export function officeTickLength(size: number, family?: unknown): number {
  return (faceBox(family)[0] / 3) * size;
}

/**
 * How far a value axis's labels stand off its line, in the unit of `size`:
 * beside an upright axis, the gap between the line and the labels' near
 * (right-aligned) edge; under or over a level axis, between the line and the
 * top of the labels' em box (where the engine sets a label's top).
 */
export function officeValueLabelGap(size: number, family: unknown, upright: boolean): number {
  const [ascent, descent, emAscent] = faceBox(family);
  if (upright) return ((ascent + descent) / 2 + ascent / 3) * size;
  return (ascent + ascent / 3 + (ascent + descent) / 4 - emAscent) * size;
}

















export function officeAxisTitleOffset(size: number, family: unknown, upright: boolean): number {
  if (upright) return 0.47 * size + 3.2;
  const [ascent, descent] = faceBox(family);
  return 0.55 * (ascent + descent) * size + 3.4;
}










export function officeAxisTitleInset(size: number, upright: boolean): number {
  return upright ? 0.7 * size + 12.9 : 0.6 * size + 10.8;
}

/**
 * Windows Excel's turned title beside a worksheet chart's left value axis
 * (lookwin lw124/fb: R480-X04's chart with its title at 8, 10, 14 and 20
 * pt, Calibri): the edge of the title's line nearer the axis stands 0.11 of
 * its size and 3.12 pt off the labels' outer edge (4.0 / 4.25 / 4.66 / 5.32
 * pt), and the line's middle 0.61 of its size and 12.4 pt in from the
 * chart's left edge (17.3 / 18.5 / 20.9 / 24.6 pt), in points. PowerPoint's
 * centre-to-labels measure (`officeAxisTitleOffset`) stood the title's line
 * half a line too far out: ECharts sets a turned name's line from its near
 * edge, not its middle (R480-X04: 3.9 pt left of Excel's).
 */
export const workbookTurnedAxisTitle = {
  gap: (size: number) => 0.11 * size + 3.12,
  inset: (size: number) => 0.61 * size + 12.4,
} as const;















export function officeTurnedCategoryNameGap(size: number, family: unknown, offset = 100): number {
  const [ascent, descent] = faceBox(family);
  return Math.max(0, ((ascent + descent) / 2 + (0.3 * (offset - 100)) / 100) * size);
}

/** The axis label key a plan leaves the gap its names take if they are
 * turned upright later (`officeTurnedCategoryNameGap`, px). */
export const OFFICE_TURNED_NAME_MARGIN = "officeTurnedNameMargin";
