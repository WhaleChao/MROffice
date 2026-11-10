/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A KaTeX formula's width from KaTeX's own HTML tree and font metrics, with no
 * DOM (lib/office-equation-katex-layout.ts states what the walk adds up and how
 * it was checked against the browser). Shared by the equation boxes' layout
 * and the Word-width fit of an imported formula (officeMathFit), which reads a
 * symbol's own advance through `symbol`.
 */

import METRICS from "katex/src/fontMetricsData.js";

/** A node of KaTeX's HTML tree (`__renderToHTMLTree`): what it draws, before any DOM. */
export interface KatexTreeNode {
  readonly classes?: readonly string[];
  readonly style?: Readonly<Record<string, string | undefined>>;
  readonly children?: readonly KatexTreeNode[];
  readonly height?: number;
  readonly depth?: number;
  /** A symbol: its characters, advance width and italic correction (em). */
  readonly text?: string;
  readonly width?: number;
  readonly italic?: number;
}

/** A symbol's width (em of its own font, before its scale), given the node
 * drawn right after it, its next sibling (null for the last of its parent:
 * what a stylesheet's `+` sees). */
export type KatexSymbolWidth = (node: KatexTreeNode, next: KatexTreeNode | null) => number;

/** KaTeX's size multipliers, size1 .. size11 (`sizing reset-sizeN sizeM`). */
const SIZE_MULTIPLIERS = [0.5, 0.6, 0.7, 0.8, 0.9, 1, 1.2, 1.44, 1.728, 2.074, 2.488];

/** Horizontal space KaTeX's stylesheet adds to a class (em of the element's own font). */
const STYLESHEET_WIDTH: Readonly<Record<string, number>> = {
  // .sqrt > .root: \mkern 5mu before, \mkern -10mu after.
  root: 0.2777777778 - 0.5555555556,
  "x-arrow-pad": 1,
  "cd-arrow-pad": 0.55556 + 0.27778,
  boxpad: 0.6,
  "cancel-pad": 0.4,
  "cancel-lap": -0.4,
  angl: 0.03889 + 0.049,
  anglpad: 0.07778,
};

export function katexEm(value: string | undefined): number {
  const match = typeof value === "string" ? /^(-?[0-9.]+)em$/.exec(value.trim()) : null;
  return match ? Number(match[1]) : 0;
}

/** The KaTeX font a symbol is set in, read from its classes. */
function fontOf(classes: readonly string[]): string {
  const has = (name: string) => classes.includes(name);
  if (has("delimsizing")) {
    const size = classes.find((name) => /^size[1-4]$/.test(name));
    if (size) return `Size${size.slice(4)}-Regular`;
  }
  if (has("large-op")) return "Size2-Regular";
  if (has("small-op")) return "Size1-Regular";
  if (has("mathnormal")) return "Math-Italic";
  if (has("boldsymbol")) return "Math-BoldItalic";
  if (has("mathbf") || has("textbf")) return "Main-Bold";
  if (has("mathit") || has("textit")) return "Main-Italic";
  if (has("amsrm") || has("mathbb") || has("textbb")) return "AMS-Regular";
  if (has("mathcal")) return "Caligraphic-Regular";
  if (has("mathfrak") || has("textfrak")) return "Fraktur-Regular";
  if (has("mathscr") || has("textscr")) return "Script-Regular";
  if (has("mathsf") || has("textsf")) return "SansSerif-Regular";
  if (has("mathtt") || has("texttt")) return "Typewriter-Regular";
  return "Main-Regular";
}

/** A symbol's advance: every character's width in its font (KaTeX keeps only the first one's). */
export function katexSymbolAdvance(node: KatexTreeNode): number {
  const characters = [...(node.text ?? "")];
  const own = node.width ?? 0;
  if (characters.length <= 1) return own;
  const metrics = METRICS[fontOf(node.classes ?? [])];
  let total = 0;
  for (const character of characters) {
    const entry = metrics?.[String(character.codePointAt(0))];
    if (!entry) return own * characters.length;
    total += entry[4] ?? 0;
  }
  return total;
}

/** The scale a `sizing reset-sizeN sizeM` span draws its children at. */
function sizedScale(classes: readonly string[], scale: number): number {
  if (!classes.includes("sizing")) return scale;
  const reset = classes.find((name) => /^reset-size\d+$/.test(name));
  const size = classes.find((name) => /^size\d+$/.test(name));
  const from = reset ? SIZE_MULTIPLIERS[Number(reset.slice(10)) - 1] : undefined;
  const to = size ? SIZE_MULTIPLIERS[Number(size.slice(4)) - 1] : undefined;
  return from && to ? (scale * to) / from : scale;
}

/** KaTeX's own: the advance and the italic correction it draws as a right margin. */
const katexSymbolWidth: KatexSymbolWidth = (node) => katexSymbolAdvance(node) + Math.max(0, node.italic ?? 0);

/** A node's width, em of the .katex font, drawn at `scale`; a symbol's own
 * width is `symbol`'s (KaTeX's by default). */
export function katexNodeWidth(
  node: KatexTreeNode,
  scale: number,
  symbol: KatexSymbolWidth = katexSymbolWidth,
  next: KatexTreeNode | null = null,
): number {
  const classes = node.classes ?? [];
  const style = node.style ?? {};
  const own =
    katexEm(style.marginLeft) + katexEm(style.marginRight) + katexEm(style.paddingLeft) + katexEm(style.paddingRight) +
    classes.reduce((sum, name) => sum + (STYLESHEET_WIDTH[name] ?? 0), 0);
  if (typeof node.text === "string") return (symbol(node, next) + own) * scale;
  const children = node.children ?? [];
  // A strut sets the height, and a Safari table-cell fix is cancelled by its row.
  if (classes.includes("strut") || classes.includes("pstrut") || classes.includes("vlist-s")) return 0;
  if (classes.includes("nulldelimiter")) return (0.12 + own) * scale;
  if (katexEm(style.width) > 0) return (katexEm(style.width) + own) * scale;
  const inner = sizedScale(classes, scale);
  const stacked = classes.includes("vlist") || classes.includes("vlist-t");
  const content = stacked
    ? Math.max(0, ...children.map((child) => katexNodeWidth(child, inner, symbol)))
    : children.reduce((sum, child, index) => sum + katexNodeWidth(child, inner, symbol, children[index + 1] ?? null), 0);
  const minimum = katexEm(style.minWidth) * scale;
  return Math.max(minimum, content) + own * scale;
}

