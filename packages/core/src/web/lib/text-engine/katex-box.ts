/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A formula's box from KaTeX's own build, with no layout: the engine's
 * reading of a formula the page has not typeset yet.
 *
 * KaTeX lays a formula out from its glyph metrics before it draws anything:
 * every glyph of its build tree carries its advance, italic correction,
 * height and depth, every space its width in em, every script and fraction
 * its size change. The drawn width is those advances added up the way its
 * stylesheet stacks the boxes -- side by side, except a vlist's rows, which
 * stand over each other -- so the box is known as the tree is built
 * (`__renderToHTMLTree`: no DOM, no fonts loaded, no layout read). A page
 * waiting for its formulas to be typeset off-screen planned no further than
 * the first one (a 272-page book: four pages a pass, a few seconds each).
 * The page still typesets each formula when it reaches it; this box only
 * lets the plan go past.
 */
import katex from "katex";

export interface KatexBox {
  /** Advance width in em of the formula's font size, plus fixed px. */
  readonly widthEm: number;
  readonly widthPx: number;
  /** Height above and depth below the baseline, em. */
  readonly heightEm: number;
  readonly depthEm: number;
}

/** KaTeX's sizing multipliers, size1..size11 (katex.css `.sizing`). */
const SIZE_MULTIPLIERS = [0.5, 0.6, 0.7, 0.8, 0.9, 1, 1.2, 1.44, 1.728, 2.074, 2.488];

/** Fixed widths katex.css gives boxes the tree states none for. */
const NULL_DELIMITER_EM = 0.12;
const VLIST_SAFARI_PX = 2;

const MAX_KEPT = 20_000;
const kept = new Map<string, KatexBox | null>();

interface BuildNode {
  readonly classes?: readonly string[];
  readonly children?: readonly BuildNode[];
  readonly style?: Record<string, string | undefined>;
  readonly text?: string;
  readonly width?: number;
  readonly italic?: number;
  readonly height?: number;
  readonly depth?: number;
}

type RenderToHtmlTree = (expression: string, options: Record<string, unknown>) => BuildNode;

/**
 * The box of `latex` as KaTeX builds it, or null where KaTeX cannot (a parse
 * error, a construct whose width only the browser knows: a `\text` run in
 * the page's own font, an SVG stretched to its content).
 */
export function katexBox(latex: string, displayMode = false): KatexBox | null {
  const key = `${displayMode ? "D" : "I"}${latex}`;
  if (kept.has(key)) return kept.get(key)!;
  const box = buildBox(latex, displayMode);
  if (kept.size >= MAX_KEPT) kept.delete(kept.keys().next().value as string);
  kept.set(key, box);
  return box;
}

function buildBox(latex: string, displayMode: boolean): KatexBox | null {
  const render = (katex as unknown as { __renderToHTMLTree?: RenderToHtmlTree }).__renderToHTMLTree;
  if (typeof render !== "function") return null;
  let tree: BuildNode;
  try {
    tree = render(latex, { displayMode, throwOnError: true, strict: "ignore", trust: false });
  } catch {
    return null;
  }
  const html = findByClass(tree, "katex-html");
  if (!html) return null;
  const fixed = { px: 0, unknown: false };
  const widthEm = widthOf(html, fixed);
  if (fixed.unknown || !Number.isFinite(widthEm)) return null;
  const heightEm = Number(html.height ?? 0);
  const depthEm = Number(html.depth ?? 0);
  if (!Number.isFinite(heightEm) || !Number.isFinite(depthEm)) return null;
  return { widthEm, widthPx: fixed.px, heightEm, depthEm };
}

function findByClass(node: BuildNode, name: string): BuildNode | null {
  if (node.classes?.includes(name)) return node;
  for (const child of node.children ?? []) {
    const found = findByClass(child, name);
    if (found) return found;
  }
  return null;
}

/** A length in em (`0.2222em`), 0 for none; `unknown` for any other unit. */
function em(value: string | undefined, fixed: { px: number; unknown: boolean }): number {
  if (!value) return 0;
  const trimmed = value.trim();
  if (trimmed.endsWith("em")) {
    const n = Number.parseFloat(trimmed);
    return Number.isFinite(n) ? n : 0;
  }
  if (trimmed.endsWith("px")) {
    const n = Number.parseFloat(trimmed);
    if (Number.isFinite(n)) fixed.px += n;
    return 0;
  }
  if (trimmed === "0" || trimmed === "auto") return 0;
  fixed.unknown = true;
  return 0;
}

const ZERO_WIDTH = /[\u200b\u200c\u200d\ufeff]/;

/** The text faces KaTeX draws a `\text…` run in, by the glyph's classes. */
const TEXT_COMMANDS: readonly (readonly [string, string])[] = [
  ["textbf", "\\textbf"], ["textit", "\\textit"], ["textsf", "\\textsf"], ["texttt", "\\texttt"], ["textrm", "\\textrm"],
];
const charAdvances = new Map<string, number | null>();

/** One character's advance in a text run of that face, em; null when
 * KaTeX has no metrics for it. */
function textCharAdvance(char: string, classes: readonly string[]): number | null {
  const command = TEXT_COMMANDS.find(([name]) => classes.includes(name))?.[1] ?? "\\text";
  const key = `${command}${char}`;
  if (charAdvances.has(key)) return charAdvances.get(key)!;
  let advance: number | null = null;
  if (char === " ") {
    advance = SPACE_EM;
  } else if (WIDE.test(char)) {
    // KaTeX lends an East Asian character a Latin glyph's metrics; the
    // browser draws it in a CJK face, a full em wide.
    advance = 1;
  } else {
    const render = (katex as unknown as { __renderToHTMLTree?: RenderToHtmlTree }).__renderToHTMLTree;
    try {
      const tree = render?.(`${command}{${/[\\{}$&#^_%~]/.test(char) ? `\\${char}` : char}}`, { throwOnError: true, strict: "ignore", trust: false });
      const html = tree ? findByClass(tree, "katex-html") : null;
      const fixed = { px: 0, unknown: false };
      const width = html ? widthOf(html, fixed) : 0;
      advance = html && !fixed.unknown && width > 0 ? width : null;
    } catch {
      advance = null;
    }
  }
  charAdvances.set(key, advance);
  return advance;
}

/** East Asian wide characters (CJK, kana, hangul, full-width forms). */
const WIDE = /[\u1100-\u11ff\u2e80-\u9fff\uac00-\ud7af\uf900-\ufaff\ufe30-\ufe4f\uff00-\uff60\uffe0-\uffe6]/u;

/** The space in a `\text` run (KaTeX_Main's, 0.25 em). */
const SPACE_EM = 0.25;

/** The factor a `sizing` / `fontsize-ensurer` span scales its content by. */
function sizingScale(classes: readonly string[]): number {
  if (!classes.includes("sizing") && !classes.includes("fontsize-ensurer")) return 1;
  let from = 0, to = 0;
  for (const name of classes) {
    const reset = /^reset-size(\d+)$/.exec(name);
    if (reset) from = Number(reset[1]);
    const size = /^size(\d+)$/.exec(name);
    if (size) to = Number(size[1]);
  }
  const a = SIZE_MULTIPLIERS[from - 1], b = SIZE_MULTIPLIERS[to - 1];
  return a && b ? b / a : 1;
}

/**
 * A box's advance in em of its parent's font size. Glyphs add their advance
 * and italic correction (KaTeX draws the correction as a right margin);
 * a vlist stands its rows over each other (the widest is its width), any
 * other box sets its children side by side; a sizing span scales what it
 * holds; margins and stated widths are the box's own.
 */
function widthOf(node: BuildNode, fixed: { px: number; unknown: boolean }): number {
  const classes = node.classes ?? [];
  const style = node.style ?? {};
  const scale = sizingScale(classes);
  let own: number;
  if (typeof node.text === "string" && typeof node.width === "number") {
    const graphemes = Array.from(node.text).filter((char) => !ZERO_WIDTH.test(char));
    if (graphemes.length > 1) {
      // A text run KaTeX joined into one glyph keeps the first one's
      // advance: each character is read alone, in the run's face.
      let sum = 0;
      for (const char of graphemes) {
        const advance = textCharAdvance(char, classes);
        if (advance === null) { fixed.unknown = true; return 0; }
        sum += advance;
      }
      own = sum;
    } else if (graphemes.length === 1 && WIDE.test(graphemes[0]!)) {
      own = 1;
    } else if (node.width === 0 && graphemes.length === 1) {
      // A glyph KaTeX has no metrics for (a CJK character): only the
      // browser can say.
      fixed.unknown = true;
      return 0;
    } else {
      own = node.width + (typeof node.italic === "number" && node.italic > 0 ? node.italic : 0);
    }
  } else if (classes.includes("vlist-s")) {
    fixed.px += VLIST_SAFARI_PX;
    own = 0;
  } else if (classes.includes("nulldelimiter")) {
    own = NULL_DELIMITER_EM;
  } else if (classes.includes("strut") || classes.includes("pstrut") || classes.includes("rlap") || classes.includes("llap") || classes.includes("clap")) {
    own = 0;
  } else if (classes.includes("katex-error")) {
    fixed.unknown = true;
    return 0;
  } else {
    const widths = (node.children ?? []).map((child) => widthOf(child, fixed));
    const stacked = classes.includes("vlist") || classes.includes("vlist-t");
    own = stacked ? widths.reduce((a, b) => Math.max(a, b), 0) : widths.reduce((a, b) => a + b, 0);
    if (style.width !== undefined) {
      const stated = em(style.width, fixed);
      if (style.width.trim().endsWith("em")) own = stated;
    }
  }
  if (style.minWidth !== undefined) own = Math.max(own, em(style.minWidth, fixed));
  const margins = em(style.marginLeft, fixed) + em(style.marginRight, fixed);
  return (own + margins) * scale;
}
