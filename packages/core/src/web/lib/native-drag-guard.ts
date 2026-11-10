/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const NATIVE_DRAG_SOURCE_SELECTOR =
  '[draggable="true"], input, textarea, [contenteditable="true"], [contenteditable=""], [contenteditable="plaintext-only"]';

type DragStartLike = {
  target: EventTarget | null;
  defaultPrevented?: boolean;
  preventDefault(): void;
};

function elementOf(target: EventTarget | null): Element | null {
  if (!target || typeof (target as Node).nodeType !== "number") return null;
  const node = target as Node;
  return node.nodeType === 1 ? (node as Element) : node.parentElement;
}

/** True when a native drag that starts at `target` is one a control or an open entry asked for. */
export function nativeDragAllowedAt(target: EventTarget | null, allow?: string): boolean {
  const element = elementOf(target);
  if (!element) return false;
  return Boolean(
    element.closest(allow ? `${NATIVE_DRAG_SOURCE_SELECTOR}, ${allow}` : NATIVE_DRAG_SOURCE_SELECTOR),
  );
}

/** `onDragStartCapture` for an editor root: refuses every native drag but the asked-for ones. */
export function refuseStrayNativeDrag(event: DragStartLike): void {
  if (event.defaultPrevented || nativeDragAllowedAt(event.target)) return;
  event.preventDefault();
}

/** The same guard with an editor's own extra drag sources (a CSS selector). */
export function refuseStrayNativeDragExcept(allow: string): (event: DragStartLike) => void {
  return (event) => {
    if (event.defaultPrevented || nativeDragAllowedAt(event.target, allow)) return;
    event.preventDefault();
  };
}
