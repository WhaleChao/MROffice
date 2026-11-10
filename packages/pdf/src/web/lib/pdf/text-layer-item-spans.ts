/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const PDF_TEXT_ITEM_INDEX_ATTR = "data-pdf-text-item-index";

/** A NodeList-like view of a text layer's spans indexed by TEXT ITEM. */
export interface PdfTextItemSpans {
  readonly length: number;
  item(itemIndex: number): HTMLElement | null;
  /** True when spans are resolved by their render-time item tag. */
  readonly tagged: boolean;
}

/** Tag pdf.js `TextLayer.textDivs` (1:1 with text items) with their index. */
export function tagTextLayerItemSpans(
  textDivs: ArrayLike<HTMLElement> | null | undefined,
): void {
  if (!textDivs) return;
  for (let index = 0; index < textDivs.length; index += 1) {
    textDivs[index]?.setAttribute(PDF_TEXT_ITEM_INDEX_ATTR, String(index));
  }
}

export function textLayerItemSpans(layer: ParentNode): PdfTextItemSpans {
  const tagged = layer.querySelectorAll<HTMLElement>(
    `span[${PDF_TEXT_ITEM_INDEX_ATTR}]`,
  );
  if (tagged.length > 0) {
    const byItem = new Map<number, HTMLElement>();
    tagged.forEach((span) => {
      const index = Number(span.getAttribute(PDF_TEXT_ITEM_INDEX_ATTR));
      if (Number.isSafeInteger(index) && index >= 0 && !byItem.has(index)) {
        byItem.set(index, span);
      }
    });
    return {
      length: byItem.size,
      item: (itemIndex) => byItem.get(itemIndex) ?? null,
      tagged: true,
    };
  }
  const spans = layer.querySelectorAll<HTMLElement>("span");
  return {
    length: spans.length,
    item: (itemIndex) => spans.item(itemIndex),
    tagged: false,
  };
}
