/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A computed style property read once per element, class list and inline
 * style. `getComputedStyle` right after an edit forces the page's styles
 * to be recomputed, and the ribbon's typography resolvers (the drawn font
 * size, weight, slant, indent of the block at the caret) asked for one on
 * EVERY render -- on a long document tens of milliseconds a render, which
 * the ink pen felt as strokes landing late, since the ribbon renders after
 * each stroke lands. ProseMirror gives a run or block new DOM when what
 * decides its style changes (its marks, its attributes), so an element
 * that keeps its classes and inline style keeps its computed value.
 */
const CACHE = new WeakMap<Element, Map<string, { key: string; value: string }>>();

function elementKey(element: Element): string {
  return `${element.className}|${element.getAttribute("style") ?? ""}`;
}

export function cachedComputedStyle(element: Element, property: string): string {
  const key = elementKey(element);
  let byProperty = CACHE.get(element);
  const cached = byProperty?.get(property);
  if (cached && cached.key === key) return cached.value;
  const view = element.ownerDocument.defaultView ?? window;
  const value = view.getComputedStyle(element).getPropertyValue(property);
  if (!byProperty) {
    byProperty = new Map();
    CACHE.set(element, byProperty);
  }
  byProperty.set(property, { key, value });
  return value;
}
