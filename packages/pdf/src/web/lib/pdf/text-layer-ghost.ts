/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * After a save the reader swaps to the saved revision: the page's text layer
 * is emptied and rebuilt (text content read, boxes laid out), and for a few
 * hundred milliseconds the page had no edit boxes at all -- every dashed
 * frame went out and came back after each edit, and a click in that moment
 * had no box to land on.
 *
 * The boxes the page showed are held meanwhile as a picture: a copy of the
 * layer with nothing in it the reader's own code reads (no layer marker, no
 * box, item or paragraph attributes; it takes no pointer and no focus) but
 * each box's line id and staged words, which name the line a press on it
 * is for (native-text-line-lineage.ts). It is paint only. A click on it falls through to the page and is kept for the
 * rebuilt layer (native-text-rebuild-click.ts), which marks the held box
 * selected. The copy goes the moment the layer has boxes again, or at its
 * deadline.
 */
import { PDF_NATIVE_TEXT_GHOST_KEPT_ATTRS } from "./native-text-line-lineage";

const EDIT_BOX = "data-pdf-native-text-edit-box";
export const PDF_TEXT_LAYER_GHOST_ATTR = "data-pdf-text-layer-ghost";
export const PDF_TEXT_LAYER_GHOST_BOX_ATTR = "data-pdf-native-text-edit-box-ghost";

/** Call before the layer is emptied. Returns a release (idempotent). */
export function holdPdfTextLayerGhost(layer: HTMLElement, deadlineMs = 3000): () => void {
  const parent = layer.parentElement;
  if (!parent || !layer.querySelector(`span[${EDIT_BOX}="true"]`)) return () => undefined;
  // One picture a page: a render superseded before it rebuilt the layer
  // leaves the first one in place (the layer it would copy is empty).
  for (const child of Array.from(parent.children)) {
    if (child.hasAttribute(PDF_TEXT_LAYER_GHOST_ATTR)) return () => undefined;
  }
  const ghost = layer.cloneNode(true) as HTMLElement;
  // A line's open editor stays on the page itself over the rebuild
  // (emptyPdfTextLayerForRebuild): a copy of it here drew its words twice.
  for (const open of Array.from(ghost.querySelectorAll('[data-pdf-native-text-inline-editing="true"]'))) open.remove();
  const strip = (element: Element, box: boolean) => {
    for (const name of element.getAttributeNames()) {
      if (name !== "class" && name !== "style" && !(box && PDF_NATIVE_TEXT_GHOST_KEPT_ATTRS.includes(name))) {
        element.removeAttribute(name);
      }
    }
    if (box) element.setAttribute(PDF_TEXT_LAYER_GHOST_BOX_ATTR, "true");
  };
  for (const element of Array.from(ghost.querySelectorAll("*"))) {
    strip(element, element.getAttribute(EDIT_BOX) === "true");
  }
  strip(ghost, false);
  ghost.setAttribute(PDF_TEXT_LAYER_GHOST_ATTR, "true");
  ghost.setAttribute("aria-hidden", "true");
  ghost.style.pointerEvents = "none";
  ghost.style.userSelect = "none";
  layer.after(ghost);

  let done = false;
  const observer = typeof MutationObserver === "function"
    ? new MutationObserver(() => { if (layer.querySelector(`span[${EDIT_BOX}="true"]`)) release(); })
    : null;
  const timer = setTimeout(() => release(), deadlineMs);
  function release() {
    if (done) return;
    done = true;
    observer?.disconnect();
    clearTimeout(timer);
    ghost.remove();
  }
  observer?.observe(layer, { childList: true, subtree: true, attributes: true, attributeFilter: [EDIT_BOX] });
  return release;
}
