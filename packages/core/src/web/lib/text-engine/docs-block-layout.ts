/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A Docs block's vertical layout from the engine: its lines' boxes and height,
 * with no reading of the page. The page ledger reads these instead of the
 * drawn boxes (L3-DESIGN): this module holds no DOM, so the same code runs in
 * a worker.
 *
 * Word's line: the highest ascent and the lowest descent of the faces on it,
 * with their line gap, times the auto multiple; an exact line is its value;
 * an at-least line the larger of its value and the single line.
 */
import { cssLineBox, type CssLineItem, type CssTextBox } from "./css-line-box";
import type { DocsParagraph, DocsPiece, DocsStrut } from "./docs-lines";
import { measureLineMetrics, type EngineFont, type EngineLineMetrics } from "./measure";

export type DocsLineRule =
  | { readonly kind: "multiple"; readonly value: number }
  | { readonly kind: "exact" | "atLeast"; readonly px: number };

export interface DocsLineBox {
  /** First piece of the line. */
  readonly start: number;
  readonly heightPx: number;
  /** What the line grid charges the line by, where it is not its box: the
   * page charges a row by its runs' line-height and its atoms' own heights
   * (pageRebalanceExtension measuredTextGridRows), not by the line box they
   * make together. */
  readonly gridPx?: number;
}

export type DocsBlockLayout =
  | { readonly status: "ok"; readonly lines: readonly DocsLineBox[]; readonly heightPx: number }
  | { readonly status: "pending" }
  | { readonly status: "unsupported"; readonly reason: string };

/** Why a line's faces leave its box to the browser: Word's rule for it here
 * is not the Latin one (the 1.3 East Asian line, a marker, a picture). */
const EAST_ASIAN = /[぀-ヿ㐀-鿿가-힯豈-﫿＀-￯]/;

type LineMetrics = (font: EngineFont) => EngineLineMetrics | "pending" | "unsupported";

function fontKey(font: EngineFont): string {
  return `${font.italic ? "i" : "n"}|${font.weight}|${font.sizePx}|${font.stack}`;
}

/** The single line of the fonts on it, or why it cannot be stated. */
function singleLinePx(
  fonts: ReadonlyMap<string, EngineFont>,
  metricsOf: LineMetrics,
  memo: Map<string, EngineLineMetrics | "pending" | "unsupported">,
): number | "pending" | "unsupported" {
  let top = 0;
  let bottom = 0;
  for (const [key, font] of fonts) {
    let metrics = memo.get(key);
    if (metrics === undefined) {
      metrics = metricsOf(font);
      memo.set(key, metrics);
    }
    if (typeof metrics === "string") return metrics;
    top = Math.max(top, metrics.ascent + metrics.lineGap / 2);
    bottom = Math.max(bottom, metrics.descent + metrics.lineGap / 2);
  }
  return top + bottom;
}

/**
 * A line as CSS draws it (css-line-box), for the lines whose box Word's rule
 * here does not give (East Asian text, an object on the line): the page has
 * already stated Word's rules as the block's and runs' line-heights and the
 * formula placeholders' boxes. Null when a piece on the line has no stated
 * box (an image, a run aligned by keyword) or the block no strut.
 */
function cssLinePx(
  pieces: readonly DocsPiece[],
  start: number,
  end: number,
  strut: DocsStrut | undefined,
  metricsOf: LineMetrics,
  memo: Map<string, EngineLineMetrics | "pending" | "unsupported">,
  /** An empty inline of the strut's font on this line (a list paragraph's
   * first-line strut). */
  extra?: { readonly lineHeight: number; readonly shift: number },
): { heightPx: number; gridPx: number } | "pending" | null {
  if (!strut) return null;
  const textBox = (font: EngineFont, lineHeight: number | "normal", shift: number): CssTextBox | "pending" | null => {
    const key = fontKey(font);
    let metrics = memo.get(key);
    if (metrics === undefined) {
      metrics = metricsOf(font);
      memo.set(key, metrics);
    }
    if (metrics === "pending") return "pending";
    if (metrics === "unsupported") return null;
    return { kind: "text", ascent: metrics.ascent, descent: metrics.descent, lineGap: metrics.lineGap, lineHeight, shift };
  };
  const strutBox = textBox(strut.font, strut.lineHeight, 0);
  if (strutBox === null || strutBox === "pending") return strutBox;
  const items: CssLineItem[] = [];
  const seen = new Set<string>();
  let gridPx = 0;
  if (extra) {
    const box = textBox(strut.font, extra.lineHeight, extra.shift);
    if (box === null || box === "pending") return box;
    items.push(box);
    gridPx = extra.lineHeight;
  }
  for (let at = start; at < end; at += 1) {
    const piece = pieces[at];
    if (piece.kind === "box" && piece.lineBoxes) {
      for (const inner of piece.lineBoxes) {
        const box = textBox(inner.font, inner.lineHeight, 0);
        if (box === null || box === "pending") return box;
        items.push(box);
        gridPx = Math.max(gridPx, inner.lineHeight === "normal" ? inner.font.sizePx * 1.2 : inner.lineHeight);
      }
      continue;
    }
    if (piece.kind === "box") {
      if (!piece.atom && (!(piece.advance > 0.01) || piece.text === "\u200b")) continue;
      if (!piece.atom) return null;
      items.push({ kind: "atom", height: piece.atom.height, depth: 0, shift: piece.atom.shift });
      gridPx = Math.max(gridPx, piece.atom.height);
      continue;
    }
    if (piece.kind === "break" || !piece.font) continue;
    if (piece.lineHeight === undefined || piece.shift === undefined) return null;
    gridPx = Math.max(gridPx, piece.lineHeight === "normal" ? piece.font.sizePx * 1.2 : piece.lineHeight);
    const key = `${fontKey(piece.font)}|${piece.lineHeight}|${piece.shift}`;
    if (seen.has(key)) continue;
    seen.add(key);
    const box = textBox(piece.font, piece.lineHeight, piece.shift);
    if (box === null || box === "pending") return box;
    items.push(box);
  }
  return { heightPx: cssLineBox(strutBox, items).height, gridPx };
}

/** The block's lines from a paragraph and its line starts (piece indices). */
export function layoutDocsBlockLines(
  paragraph: Pick<DocsParagraph, "pieces" | "unsupported" | "pending" | "strut" | "firstLineStrut">,
  starts: readonly number[],
  rule: DocsLineRule,
  strut: EngineFont | null,
  metricsOf: LineMetrics = measureLineMetrics,
): DocsBlockLayout {
  if (paragraph.unsupported) return { status: "unsupported", reason: paragraph.unsupported };
  if (paragraph.pending) return { status: "pending" };
  const memo = new Map<string, EngineLineMetrics | "pending" | "unsupported">();
  const lines: DocsLineBox[] = [];
  let heightPx = 0;
  const count = paragraph.pieces.length;
  const lineStarts = starts.length ? starts : [0];
  for (let index = 0; index < lineStarts.length; index += 1) {
    const start = lineStarts[index];
    const end = lineStarts[index + 1] ?? count;
    const fonts = new Map<string, EngineFont>();
    let boxed = false;
    let eastAsian = false;
    for (let at = start; at < end; at += 1) {
      const piece: DocsPiece = paragraph.pieces[at];
      if (piece.kind === "box") boxed ||= (piece.advance > 0.01 && piece.text !== "\u200b") || piece.atom !== undefined || piece.lineBoxes !== undefined;
      eastAsian ||= EAST_ASIAN.test(piece.text);
      if (piece.font) fonts.set(fontKey(piece.font), piece.font);
    }
    const extra = index === 0 ? paragraph.firstLineStrut : undefined;
    if (eastAsian || boxed || extra) {
      const css = cssLinePx(paragraph.pieces, start, end, paragraph.strut, metricsOf, memo, extra);
      if (css === "pending") return { status: "pending" };
      if (css === null) return { status: "unsupported", reason: eastAsian ? "east asian line" : "object on the line" };
      lines.push({ start, heightPx: css.heightPx, gridPx: css.gridPx });
      heightPx += css.heightPx;
      continue;
    }
    if (fonts.size === 0 && strut) fonts.set(fontKey(strut), strut);
    if (fonts.size === 0) return { status: "unsupported", reason: "no font on the line" };
    const single = singleLinePx(fonts, metricsOf, memo);
    if (single === "pending") return { status: "pending" };
    if (single === "unsupported") return { status: "unsupported", reason: "face without metrics" };
    const linePx =
      rule.kind === "multiple" ? single * rule.value : rule.kind === "exact" ? rule.px : Math.max(rule.px, single);
    lines.push({ start, heightPx: linePx });
    heightPx += linePx;
  }
  return { status: "ok", lines, heightPx };
}
