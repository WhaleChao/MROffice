/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { CSSProperties } from "react";

import {
  clampLevel,
  drawnBulletChar,
  normalizeNativeBulletMarker,
  normalizeParagraphIndentEmu,
  PARAGRAPH_ALIGNS,
  renderParagraphs,
  type NativeBulletMarker,
  type ParagraphAlign,
  type SlideParagraph,
} from "./presentation-paragraphs";
import {
  POWERPOINT_AUTONUM_BY_NUMBER_FORMAT,
  POWERPOINT_BULLET_CHARS,
  POWERPOINT_WEB_NUMBER_FORMATS,
} from "./office-list-catalog";
import {
  MAX_LINE_SPACING as MAX_STORED_LINE_SPACING,
  MAX_PARAGRAPH_SPACING_PT,
  MIN_PARAGRAPH_SPACING_PT,
  MIN_STORED_LINE_SPACING,
  PX_PER_POINT,
} from "./presentation-line-spacing";

/** 914400 EMU to the inch, 96 CSS px to the inch (MS-ODRAW glossary). */
import { EMU_PER_PX as EMU_PER_CSS_PIXEL } from "./office-units";
export { EMU_PER_CSS_PIXEL };

/**
 * `a:bodyPr`'s four text insets, with the values implied when the attribute is
 * absent, as the Open XML SDK's BodyProperties page states them: the insets
 * "are used just as internal margins", `lIns`/`rIns` 91440 EMU = 0.1in and
 * `tIns`/`bIns` 45720 EMU = 0.05in.
 */
export const OFFICE_BODY_INSET_DEFAULTS_EMU = Object.freeze({
  lIns: 91_440,
  tIns: 45_720,
  rIns: 91_440,
  bIns: 45_720,
});

export type OfficeTextAnchor = "t" | "ctr" | "b";
export type OfficeTextVertOverflow = "overflow" | "ellipsis" | "clip";
export type OfficeTextHorzOverflow = "overflow" | "clip";
export type OfficeTextWrap = "square" | "none";

/** The STATED members of one `a:bodyPr`. Everything optional on purpose. */
export interface OfficeTextBodyProperties {
  anchor?: OfficeTextAnchor;
  anchorCtr?: boolean;
  lIns?: number;
  tIns?: number;
  rIns?: number;
  bIns?: number;
  vertOverflow?: OfficeTextVertOverflow;
  horzOverflow?: OfficeTextHorzOverflow;
  wrap?: OfficeTextWrap;
  /** Do not rotate text (`a:bodyPr@upright`): the words stay level while
   * the shape turns. */
  upright?: true;
  /** Text direction (`a:bodyPr@vert`) of a sheet shape's words: Rotate all
   * text 90 (`vert`) or 270 (`vert270`) degrees, turned with the box; East
   * Asian vertical (`eaVert`) and Stacked (`wordArtVert`, columns left to
   * right; `wordArtVertRtl`, right to left), set in columns. */
  vert?: "vert" | "vert270" | "eaVert" | "wordArtVert" | "wordArtVertRtl";
}

/** One `a:p`'s layout — or one LINE of it, when `a:br` split the paragraph. */
export interface OfficeTextParagraph {
  level: number;
  bullet: "none" | "bullet" | "number";
  bulletChar?: string;
  numberFormat?: string;
  /** `a:pPr@marL`, the paragraph's own left margin in EMU. */
  marL?: number;
  /** `a:pPr@indent`, the FIRST LINE's offset from it; negative = hanging. */
  indent?: number;
  align?: ParagraphAlign;
  nativeBullet?: NativeBulletMarker;
  /** `a:lnSpc/a:spcPct` as a multiple of single spacing. */
  lineSpacing?: number;
  /** `a:spcBef/a:spcPts` and `a:spcAft/a:spcPts`, in points. */
  spaceBefore?: number;
  spaceAfter?: number;
  


  defTabSz?: number;
}

/** Office's default tab interval when a paragraph states none: one inch. */
export const OFFICE_DEFAULT_TAB_EMU = 914_400;

export interface OfficeTextLayout {
  paragraphs: OfficeTextParagraph[];
  bodyPr?: OfficeTextBodyProperties;
}

/** The same ceiling the block paragraph model carries. */
export const MAX_OFFICE_TEXT_LAYOUT_PARAGRAPHS = 200;

const ANCHORS: ReadonlySet<string> = new Set(["t", "ctr", "b"]);
const VERT_OVERFLOWS: ReadonlySet<string> = new Set([
  "overflow",
  "ellipsis",
  "clip",
]);
const HORZ_OVERFLOWS: ReadonlySet<string> = new Set(["overflow", "clip"]);
const WRAPS: ReadonlySet<string> = new Set(["square", "none"]);
const ALIGNS: ReadonlySet<string> = new Set(PARAGRAPH_ALIGNS);
const BULLET_CHARS: ReadonlySet<string> = new Set(POWERPOINT_BULLET_CHARS);
const NUMBER_FORMATS: ReadonlySet<string> = new Set(
  POWERPOINT_WEB_NUMBER_FORMATS,
);

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function normalizeBodyProperties(
  value: unknown,
): OfficeTextBodyProperties | undefined {
  if (!isRecord(value)) return undefined;
  const body: OfficeTextBodyProperties = {};
  if (typeof value.anchor === "string" && ANCHORS.has(value.anchor)) {
    body.anchor = value.anchor as OfficeTextAnchor;
  }
  if (value.anchorCtr === true) body.anchorCtr = true;
  if (value.upright === true) body.upright = true;
  if (
    value.vert === "vert" ||
    value.vert === "vert270" ||
    value.vert === "eaVert" ||
    value.vert === "wordArtVert" ||
    value.vert === "wordArtVertRtl"
  ) {
    body.vert = value.vert;
  }
  for (const key of ["lIns", "tIns", "rIns", "bIns"] as const) {
    const inset = normalizeParagraphIndentEmu(value[key], { signed: false });
    if (inset !== undefined) body[key] = inset;
  }
  if (
    typeof value.vertOverflow === "string" &&
    VERT_OVERFLOWS.has(value.vertOverflow)
  ) {
    body.vertOverflow = value.vertOverflow as OfficeTextVertOverflow;
  }
  if (
    typeof value.horzOverflow === "string" &&
    HORZ_OVERFLOWS.has(value.horzOverflow)
  ) {
    body.horzOverflow = value.horzOverflow as OfficeTextHorzOverflow;
  }
  if (typeof value.wrap === "string" && WRAPS.has(value.wrap)) {
    body.wrap = value.wrap as OfficeTextWrap;
  }
  return Object.keys(body).length ? body : undefined;
}

function normalizeParagraph(value: unknown): OfficeTextParagraph | undefined {
  if (!isRecord(value)) return undefined;
  const bullet =
    value.bullet === "bullet" || value.bullet === "number"
      ? value.bullet
      : "none";
  const row: OfficeTextParagraph = {
    level: clampLevel(value.level),
    bullet,
  };
  if (
    bullet === "bullet" &&
    typeof value.bulletChar === "string" &&
    BULLET_CHARS.has(value.bulletChar)
  ) {
    row.bulletChar = value.bulletChar;
  }
  if (
    bullet === "number" &&
    typeof value.numberFormat === "string" &&
    NUMBER_FORMATS.has(value.numberFormat)
  ) {
    row.numberFormat = value.numberFormat;
  }
  const marL = normalizeParagraphIndentEmu(value.marL, { signed: false });
  if (marL !== undefined) row.marL = marL;
  const indent = normalizeParagraphIndentEmu(value.indent, { signed: true });
  if (indent !== undefined) row.indent = indent;
  if (typeof value.align === "string" && ALIGNS.has(value.align)) {
    row.align = value.align as ParagraphAlign;
  }
  const marker = normalizeNativeBulletMarker(value.nativeBullet);
  if (marker) row.nativeBullet = marker;
  if (
    typeof value.lineSpacing === "number" &&
    Number.isFinite(value.lineSpacing) &&
    value.lineSpacing >= MIN_STORED_LINE_SPACING &&
    value.lineSpacing <= MAX_STORED_LINE_SPACING
  ) {
    row.lineSpacing = value.lineSpacing;
  }
  for (const key of ["spaceBefore", "spaceAfter"] as const) {
    const raw = value[key];
    if (
      typeof raw === "number" &&
      Number.isFinite(raw) &&
      raw >= MIN_PARAGRAPH_SPACING_PT &&
      raw <= MAX_PARAGRAPH_SPACING_PT
    ) {
      row[key] = raw;
    }
  }
  const tab = normalizeParagraphIndentEmu(value.defTabSz, { signed: false });
  if (tab !== undefined && tab > 0) row.defTabSz = tab;
  return row;
}








export function officeTextDefaultTabPx(layout: OfficeTextLayout | undefined): number {
  const stated = layout?.paragraphs.find((paragraph) => paragraph.defTabSz)?.defTabSz;
  return (stated ?? OFFICE_DEFAULT_TAB_EMU) / EMU_PER_CSS_PIXEL;
}

/** Bound one persisted `textLayout` receipt, or drop it whole. */
export function normalizeOfficeTextLayout(
  value: unknown,
): OfficeTextLayout | undefined {
  if (!isRecord(value)) return undefined;
  if (!Array.isArray(value.paragraphs)) return undefined;
  if (
    value.paragraphs.length === 0 ||
    value.paragraphs.length > MAX_OFFICE_TEXT_LAYOUT_PARAGRAPHS
  ) {
    return undefined;
  }
  const paragraphs: OfficeTextParagraph[] = [];
  for (const entry of value.paragraphs) {
    const row = normalizeParagraph(entry);
    if (!row) return undefined;
    paragraphs.push(row);
  }
  const layout: OfficeTextLayout = { paragraphs };
  const body = normalizeBodyProperties(value.bodyPr);
  if (body) layout.bodyPr = body;
  return layout;
}

/** One drawn LINE: its characters, its marker, and where the two sit. */
export interface OfficeTextLine {
  text: string;
  /** The bullet glyph or auto-number label, or null when the line has none. */
  marker: string | null;
  /** `a:pPr@marL` as CSS px. 0 when the paragraph states none. */
  marginPx: number;
  /** `a:pPr@indent` as CSS px (negative = hanging). 0 when none is stated. */
  indentPx: number;
  align?: ParagraphAlign;
  /** `a:buClr`'s own colour, when the package painted the marker separately. */
  markerColor?: string;
  /** `a:buSzPct` as a multiple of the run size, when stated. */
  markerScale?: number;
  /** `a:buSzPts` in points, when stated (an ABSOLUTE size, not a multiple). */
  markerSizePt?: number;
  /** The paragraph's line spacing multiple, when stated (single otherwise). */
  lineSpacing?: number;
  /** Space before / after the paragraph in CSS px; only on its first / last line. */
  spaceBeforePx?: number;
  spaceAfterPx?: number;
}












export function officeTextLayoutLines(
  text: string,
  layout?: OfficeTextLayout,
): OfficeTextLine[] {
  const lines = text.split("\n");
  const rows = layout?.paragraphs;
  if (!rows || rows.length !== lines.length) {
    return lines.map((line) => ({
      text: line,
      marker: null,
      marginPx: 0,
      indentPx: 0,
    }));
  }
  const source: SlideParagraph[] = rows.map((row, index) => ({
    text: lines[index],
    level: row.level,
    bullet: row.bullet,
  }));
  // The renderer takes ONE block-level number format, and a text body states a
  // format PER PARAGRAPH — S047 alone switches from `arabicPeriod` to
  // `alphaUcPeriod` half way down. The per-row channel that already exists is
  // the `nativeBullet` receipt, so a row the gallery CAN name is handed to the
  // renderer as the scheme it names, through the one catalogue. Nothing is
  // persisted here: this array is an argument, not a page field.
  const rendered = renderParagraphs(source, {
    nativeBullet: rows.map((row) => {
      if (row.nativeBullet) return row.nativeBullet;
      if (row.bullet !== "number" || !row.numberFormat) return null;
      const autoNum =
        POWERPOINT_AUTONUM_BY_NUMBER_FORMAT[
          row.numberFormat as keyof typeof POWERPOINT_AUTONUM_BY_NUMBER_FORMAT
        ];
      return autoNum ? { autoNum } : null;
    }),
  });
  return rows.map((row, index) => {
    const marker =
      lines[index].trim().length === 0 ? null : markerFor(row, rendered[index]);
    const line: OfficeTextLine = {
      text: lines[index],
      marker,
      marginPx: row.marL ? row.marL / EMU_PER_CSS_PIXEL : 0,
      indentPx: row.indent ? row.indent / EMU_PER_CSS_PIXEL : 0,
    };
    if (row.align) line.align = row.align;
    if (row.lineSpacing !== undefined) line.lineSpacing = row.lineSpacing;
    if (row.spaceBefore) line.spaceBeforePx = row.spaceBefore * PX_PER_POINT;
    if (row.spaceAfter) line.spaceAfterPx = row.spaceAfter * PX_PER_POINT;
    if (marker !== null && row.nativeBullet) {
      if (row.nativeBullet.color) line.markerColor = row.nativeBullet.color;
      if (row.nativeBullet.sizePct !== undefined) {
        line.markerScale = row.nativeBullet.sizePct / 100_000;
      } else if (row.nativeBullet.sizePts !== undefined) {
        line.markerSizePt = row.nativeBullet.sizePts / 100;
      }
    }
    return line;
  });
}

/**
 * The layout of `before`'s paragraphs carried to the paragraphs of `after`
 * (the words as typed), as Office carries a paragraph's format: lines kept
 * at either end keep their own row, a paragraph split by Enter gives both
 * halves its row, and a paragraph added after the last one takes that one's.
 * Unchanged when the paragraph count is; undefined when there is no layout
 * or the words outgrow its ceiling.
 */
export function officeTextLayoutFollowingText(
  layout: OfficeTextLayout | undefined,
  before: string,
  after: string,
): OfficeTextLayout | undefined {
  const rows = layout?.paragraphs;
  const old = before.split("\n");
  const next = after.split("\n");
  if (!layout || !rows || rows.length !== old.length) return layout;
  if (next.length === old.length) return layout;
  if (next.length > MAX_OFFICE_TEXT_LAYOUT_PARAGRAPHS) return undefined;
  let prefix = 0;
  while (prefix < old.length && prefix < next.length && old[prefix] === next[prefix]) prefix += 1;
  let suffix = 0;
  while (
    suffix < old.length - prefix &&
    suffix < next.length - prefix &&
    old[old.length - 1 - suffix] === next[next.length - 1 - suffix]
  ) {
    suffix += 1;
  }
  const oldMiddle = rows.slice(prefix, rows.length - suffix);
  const fallback = rows[Math.max(0, prefix - 1)] ?? rows[rows.length - suffix] ?? rows[0]!;
  const middle = Array.from({ length: next.length - prefix - suffix }, (_, index) =>
    oldMiddle.length ? oldMiddle[Math.min(index, oldMiddle.length - 1)]! : fallback,
  );
  return {
    ...layout,
    paragraphs: [...rows.slice(0, prefix), ...middle, ...rows.slice(rows.length - suffix)],
  };
}

function markerFor(
  row: OfficeTextParagraph,
  rendered: { marker: string | null } | undefined,
): string | null {
  if (row.bullet === "none") return null;
  if (row.bullet === "bullet") {
    // The package's own glyph where the receipt carries one (a private-use code
    // point is nothing without its `a:buFont`, so the gallery glyph is what is
    // drawn when the receipt names no char of its own).
    const own = row.nativeBullet?.char;
    return (own === undefined ? undefined : drawnBulletChar(own)) ?? row.bulletChar ?? rendered?.marker ?? null;
  }
  return rendered?.marker ?? null;
}

/**
 * How a host draws one line's indents and marker: `a:pPr@marL` as the
 * paragraph's `margin-inline-start` and `@indent` as its `text-indent`, a
 * hanging indent opening the gutter the marker (an inline-block of at least
 * that width) fills, so the characters start at `marL` on every line. The
 * marker takes `a:buClr` and `a:buSzPct` (a multiple of the text size) or
 * `a:buSzPts`. Spacing and alignment stay the host's.
 */
export function officeTextLineIndentCss(line: OfficeTextLine): { block: CSSProperties; marker: CSSProperties } {
  const gutter = line.indentPx < 0 ? -line.indentPx : 0;
  return {
    block: {
      display: "block",
      marginInlineStart: line.marginPx ? `${line.marginPx}px` : undefined,
      textIndent: line.indentPx ? `${line.indentPx}px` : undefined,
    },
    marker: {
      display: "inline-block",
      minWidth: gutter ? `${gutter}px` : undefined,
      // `text-indent` is inherited and an inline-block is a line box of its
      // own: without this the marker's glyph moves out of its own box.
      textIndent: 0,
      color: line.markerColor,
      fontSize:
        line.markerSizePt !== undefined
          ? `${line.markerSizePt}pt`
          : line.markerScale !== undefined
            ? `${line.markerScale}em`
            : undefined,
    },
  };
}

/**
 * `a:bodyPr`'s insets as CSS px, falling back to the ECMA-376 schema defaults
 * for the attributes the package does not state.
 */
export function officeTextBodyInsetsPx(body?: OfficeTextBodyProperties): {
  left: number;
  top: number;
  right: number;
  bottom: number;
} {
  const pick = (key: keyof typeof OFFICE_BODY_INSET_DEFAULTS_EMU) =>
    (body?.[key] ?? OFFICE_BODY_INSET_DEFAULTS_EMU[key]) / EMU_PER_CSS_PIXEL;
  return {
    left: pick("lIns"),
    top: pick("tIns"),
    right: pick("rIns"),
    bottom: pick("bIns"),
  };
}

/**
 * `a:bodyPr@anchorCtr` -- "Top Centered" / "Middle Centered" / "Bottom
 * Centered" in PowerPoint's and Excel's Format Shape > Text Box: the text's
 * SMALLEST bounding box is centred across the frame, and each paragraph aligns
 * INSIDE that box, not inside the frame (ECMA-376 §21.1.2.1.1: different from
 * paragraph alignment, which aligns the text within the bounds box). The CSS
 * for that box: as wide as its longest line (at most the frame, where it
 * still wraps), centred by its auto side margins -- which centre it in a
 * block container and in a row flex container alike. Empty when the attribute
 * is off, so every other body renders exactly as before. Horizontal text
 * only: the readers state it only for a horizontal body.
 */













export function officeAnchorCenterShift(lineRights: readonly number[], width: number): number {
  const rights = lineRights.filter((right) => Number.isFinite(right));
  if (!rights.length || !(width > 0)) return 0;
  const right = Math.max(...rights);
  return right < width ? (width - right) / 2 : 0;
}

/**
 * {@link officeAnchorCenterShift} for a body the browser laid out: the
 * right edges of its drawn text, read off the layout in `element`'s own
 * (untransformed) CSS pixels, whatever scale the canvas is drawn at.
 */
export function measuredOfficeAnchorCenterShift(element: HTMLElement): number {
  const width = element.clientWidth;
  const box = element.getBoundingClientRect();
  if (!(width > 0) || !(box.width > 0)) return 0;
  const scale = box.width / element.offsetWidth || 1;
  // The text's own boxes, node by node: a range over the element would also
  // give each row element's whole-width border box.
  const doc = element.ownerDocument;
  const walker = doc.createTreeWalker(element, NodeFilter.SHOW_TEXT);
  const range = doc.createRange();
  const rights: number[] = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (!(node.textContent ?? "").trim()) continue;
    range.selectNodeContents(node);
    for (const rect of Array.from(range.getClientRects())) {
      if (rect.width > 0 && rect.height > 0) rights.push((rect.right - box.left) / scale);
    }
  }
  range.detach?.();
  return officeAnchorCenterShift(rights, width);
}

export function officeAnchorCenterBoundsCss(anchorCtr: boolean): CSSProperties {
  return anchorCtr
    ? {
        width: "fit-content",
        maxWidth: "100%",
        marginLeft: "auto",
        marginRight: "auto",
      }
    : {};
}
