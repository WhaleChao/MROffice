/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Breadcrumbs (VSCode parity): the thin strip above the code editor that shows
 * the document/cell name followed by the enclosing-symbol trail at the caret
 * (e.g. `file.ts › Calculator › add`), each crumb clickable to jump to that
 * symbol.
 *
 * This module is the PURE, framework-agnostic core: given the SAME flattened
 * {@link DocumentSymbol} list the Go-to-Symbol picker already computes (TS/JS via
 * the navigation tree, every other language via the regex outline) plus a caret
 * offset, it derives the deepest containing-span chain — ordered OUTERMOST first.
 * The editor renders the document/cell name as the leading crumb and this chain
 * as the trailing crumbs, reusing the existing `jumpToSymbol` dispatch when a
 * crumb is clicked.
 *
 * Why a flat list is enough: {@link import("./code-intelligence").documentSymbols}
 * (and the regex {@link import("./outline-symbols").outlineSymbols}) emit symbols
 * in document order with each span being a real sub-range of its container. Two
 * symbols that BOTH contain the same caret therefore can never be disjoint — one
 * must be nested inside the other — so the containing set, sorted by span size
 * (widest first), IS the ancestor chain from outermost to innermost. No tree walk
 * is required; the `depth` field is not even consulted (it can be stale relative
 * to skipped span-less ancestors, whereas the spans are authoritative).
 *
 * TOTAL + pure: a caret outside every symbol (or an empty list) yields `[]`, so
 * the editor renders just the file crumb. Never throws.
 */
import type { DocumentSymbol } from "./code-intelligence";

/**
 * The enclosing-symbol chain at `caret`, OUTERMOST → innermost.
 *
 * Returns every symbol whose span `[from, to]` contains `caret` (both bounds
 * INCLUSIVE — so a caret resting exactly at a span edge still counts as inside,
 * matching how the picker marks the active entry), ordered so the widest span
 * (the outermost container) comes first and the tightest (the symbol the caret
 * literally sits in) comes last. Ties on width are broken by `from` ascending
 * then `to` descending so the order is deterministic.
 *
 * The original array index of each symbol is preserved via the returned objects
 * being the SAME references from `symbols`, so a caller can map a crumb back to
 * its position with `symbols.indexOf(crumb)` to drive `jumpToSymbol`.
 *
 * @param symbols Flattened document outline (document order; may be empty).
 * @param caret   0-based buffer offset of the caret.
 * @returns The containing symbols outermost-first, or `[]` when the caret is
 *          outside every symbol.
 */
export function enclosingSymbolChain(
  symbols: readonly DocumentSymbol[],
  caret: number,
): DocumentSymbol[] {
  const containing = symbols.filter(
    (sym) => caret >= sym.from && caret <= sym.to,
  );
  // Widest span first = outermost ancestor first. Because the containing set is
  // strictly nested (see module doc), span width is a total order over it apart
  // from exact-duplicate spans, which the from/to tiebreak makes deterministic.
  containing.sort((a, b) => {
    const widthA = a.to - a.from;
    const widthB = b.to - b.from;
    if (widthA !== widthB) return widthB - widthA; // wider (outer) first
    if (a.from !== b.from) return a.from - b.from;
    return b.to - a.to;
  });
  return containing;
}

/** One rendered crumb: a symbol from the chain plus its index in the full list. */
export interface BreadcrumbCrumb {
  /** The symbol this crumb represents. */
  symbol: DocumentSymbol;
  /** Index of `symbol` within the original `symbols` list (drives jumpToSymbol). */
  index: number;
}

/**
 * The clickable symbol crumbs for the breadcrumb strip: {@link enclosingSymbolChain}
 * paired with each symbol's index in the original list (so a click can call
 * `jumpToSymbol(index)` against the SAME list the picker uses). The leading
 * document/cell-name crumb is added by the renderer, not here. Outermost first.
 *
 * Resolving the index here (rather than at render time) keeps the lookup against
 * the exact array the chain was derived from — a later `indexOf` on a copy would
 * be wrong if duplicate spans existed.
 */
export function breadcrumbCrumbs(
  symbols: readonly DocumentSymbol[],
  caret: number,
): BreadcrumbCrumb[] {
  const chain = enclosingSymbolChain(symbols, caret);
  return chain.map((symbol) => ({ symbol, index: symbols.indexOf(symbol) }));
}
