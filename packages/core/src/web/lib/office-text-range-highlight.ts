/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Paint ranges of text as part of the text: the CSS Custom Highlight API.
 *
 * A highlight drawn as boxes over the page, measured from the words'
 * client rects, is only right while the words are level and still: a turned
 * text box reports the upright square around each turned line, so the paint
 * stood as a square beside the words, and a box moved or turned after the
 * measure left the paint where the words had been. A named highlight is
 * painted by the browser on the range's own glyphs, behind them like the
 * selection: it turns, wraps, scales and moves with its characters under
 * any transform, sits under whatever covers the words, and goes with them.
 *
 * The ranges are StaticRanges: a live Range registered with the document
 * slows every DOM mutation in an open editor. A static range whose text
 * node leaves the page paints nothing; its owner measures again.
 *
 * Several owners (windows) can paint under one name: each sets its own
 * ranges and the name paints their union.
 */

export interface OfficeTextRangeHighlightPaint {
  /** The highlight's registered name (`::highlight(name)`). */
  readonly name: string;
  /** The declarations of its `::highlight()` rule. */
  readonly css: string;
}

const owners = new Map<string, Map<symbol, readonly AbstractRange[]>>();

/** Whether this browser paints named highlights (Chrome 105, Safari 17.2). */
export function officeTextRangeHighlightsSupported(): boolean {
  return (
    typeof CSS !== "undefined" &&
    "highlights" in CSS &&
    typeof Highlight !== "undefined" &&
    typeof StaticRange !== "undefined"
  );
}

// The build's CSS parser rejects `::highlight()`, so each rule is added to
// the document once, when its highlight is first painted.
function ensureRule(paint: OfficeTextRangeHighlightPaint): void {
  const id = `${paint.name}-style`;
  if (document.getElementById(id)) return;
  const style = document.createElement("style");
  style.id = id;
  // View state, never printed.
  style.textContent =
    `::highlight(${paint.name}) {\n  ${paint.css}\n}\n` +
    `@media print {\n  ::highlight(${paint.name}) {\n    background-color: transparent;\n  }\n}`;
  document.head.append(style);
}

/** Paint `owner`'s ranges under `paint` (an empty list takes them off). */
export function setOfficeTextRangeHighlight(
  paint: OfficeTextRangeHighlightPaint,
  owner: symbol,
  ranges: readonly AbstractRange[],
): void {
  if (!officeTextRangeHighlightsSupported()) return;
  let byOwner = owners.get(paint.name);
  if (!byOwner) {
    byOwner = new Map();
    owners.set(paint.name, byOwner);
  }
  if (ranges.length > 0) byOwner.set(owner, ranges);
  else byOwner.delete(owner);
  const all = [...byOwner.values()].flat();
  if (all.length === 0) {
    CSS.highlights.delete(paint.name);
    return;
  }
  ensureRule(paint);
  CSS.highlights.set(paint.name, new Highlight(...all));
}

/** Take `owner`'s ranges off `paint`. */
export function clearOfficeTextRangeHighlight(
  paint: OfficeTextRangeHighlightPaint,
  owner: symbol,
): void {
  setOfficeTextRangeHighlight(paint, owner, []);
}

/**
 * Word's Reading Highlight (Find > Highlight All): every match yellow, the
 * one the pane is on orange -- the colours of `.document-find-match`
 * (styles/note-editor.css), which the document editor paints inline.
 */
export const OFFICE_FIND_MATCH_PAINT: OfficeTextRangeHighlightPaint = {
  name: "mrai-find-match",
  css: "background-color: #fff200; color: inherit;",
};
export const OFFICE_FIND_CURRENT_PAINT: OfficeTextRangeHighlightPaint = {
  name: "mrai-find-current",
  css: "background-color: #ff9632; color: inherit;",
};
