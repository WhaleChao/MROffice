/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { OFFICE_CHROME_SELECTOR } from "@mrnote/core/web/lib/office-inactive-selection.ts";

/**
 * What keeps a PDF object selected when it is pressed: the hosts that say so,
 * and the editor chrome every editor shares (the ribbon and its tabs, menus,
 * panes, dialogs). The text editor's ribbon carried no retain marker, so a
 * press on any of its tabs deselected the text box and dropped the words
 * selected in it -- the command the user reached for then had nothing to act
 * on.
 */
const PDF_OBJECT_SELECTION_RETAIN_SELECTOR =
  '[data-pdf-object-selection-retain="true"],' +
  '[data-ribbon-object-selection-retain="true"],' +
  OFFICE_CHROME_SELECTOR;

function retainHostForTarget(target: EventTarget | null): Element | null {
  const element =
    target instanceof Element
      ? target
      : target instanceof Node
        ? target.parentElement
        : null;
  return element?.closest(PDF_OBJECT_SELECTION_RETAIN_SELECTOR) ?? null;
}

function pointerInsideRetainHost(event: PointerEvent): boolean {
  if (typeof document === "undefined") return false;
  const { clientX, clientY } = event;
  if (!Number.isFinite(clientX) || !Number.isFinite(clientY)) return false;
  return Array.from(
    document.querySelectorAll<HTMLElement>(
      PDF_OBJECT_SELECTION_RETAIN_SELECTOR,
    ),
  ).some((host) => {
    const rect = host.getBoundingClientRect();
    return (
      rect.width > 0 &&
      rect.height > 0 &&
      clientX >= rect.left &&
      clientX <= rect.right &&
      clientY >= rect.top &&
      clientY <= rect.bottom
    );
  });
}

/**
 * A real pointer event may target a Text node inside a ribbon button instead
 * of the button/SVG element itself. Walk both the target ancestry and the
 * composed path so PDF object selection survives every ribbon control hit.
 */
export function pointerRetainsPdfObjectSelection(event: PointerEvent): boolean {
  if (retainHostForTarget(event.target)) return true;
  if (
    event.composedPath().some((target) => retainHostForTarget(target) !== null)
  ) {
    return true;
  }
  return pointerInsideRetainHost(event);
}
