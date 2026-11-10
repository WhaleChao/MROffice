/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
type Held =
  | Readonly<{ kind: "mouse"; ctor: "pointer" | "mouse"; type: string; clientX: number; clientY: number;
      button: number; buttons: number; detail: number; shiftKey: boolean; metaKey: boolean; ctrlKey: boolean; altKey: boolean }>
  | Readonly<{ kind: "key"; type: string; key: string; code: string; shiftKey: boolean; metaKey: boolean; ctrlKey: boolean; altKey: boolean }>;

const MOUSE = ["pointerdown", "mousedown", "pointerup", "mouseup", "click", "dblclick"] as const;
const KEYS = ["keydown", "keyup"] as const;

export type PdfNativeTextGestureHold = Readonly<{
  /** Plays back what was held; a second call does nothing. */
  release: () => void;
  /** How many events are held (for what reads them in tests and marks). */
  count: () => number;
}>;

export function holdPdfNativeTextGestures(ownerDocument: Document, options?: Readonly<{
  /** Whether a mouse event at this target is the page's: the app's other
   * surfaces (a menu, the title bar) go on as ever. */
  isPageEvent?: (event: Event) => boolean;
}>): PdfNativeTextGestureHold {
  const held: Held[] = [];
  let open = true;
  const take = (event: Event) => {
    if (!open || (options?.isPageEvent && !options.isPageEvent(event))) return;
    if (event instanceof KeyboardEvent) {
      if (event.isComposing) return;
      held.push({ kind: "key", type: event.type, key: event.key, code: event.code, shiftKey: event.shiftKey,
        metaKey: event.metaKey, ctrlKey: event.ctrlKey, altKey: event.altKey });
    } else if (event instanceof MouseEvent) {
      held.push({ kind: "mouse", ctor: event.type.startsWith("pointer") ? "pointer" : "mouse", type: event.type,
        clientX: event.clientX, clientY: event.clientY, button: event.button, buttons: event.buttons, detail: event.detail,
        shiftKey: event.shiftKey, metaKey: event.metaKey, ctrlKey: event.ctrlKey, altKey: event.altKey });
    } else return;
    event.preventDefault();
    event.stopImmediatePropagation();
  };
  // On the window: the page's React root is the document itself (Next), whose
  // capture listeners were there first, so held on the document a click still
  // reached the viewer's handler and took the kept drag's place.
  const listening: EventTarget = ownerDocument.defaultView ?? ownerDocument;
  for (const type of MOUSE) listening.addEventListener(type, take, true);
  for (const type of KEYS) listening.addEventListener(type, take, true);
  const stop = () => {
    for (const type of MOUSE) listening.removeEventListener(type, take, true);
    for (const type of KEYS) listening.removeEventListener(type, take, true);
  };
  return {
    count: () => held.length,
    release: () => {
      if (!open) return;
      open = false;
      stop();
      const view = ownerDocument.defaultView;
      if (!view) return;
      for (const event of held) {
        if (event.kind === "key") {
          const { kind: _kind, type, ...init } = event;
          (ownerDocument.activeElement ?? ownerDocument.body).dispatchEvent(new KeyboardEvent(type,
            { ...init, bubbles: true, cancelable: true, composed: true }));
          continue;
        }
        const { kind: _kind, ctor, type, ...init } = event;
        const target = ownerDocument.elementFromPoint?.(init.clientX, init.clientY) ?? ownerDocument.body;
        const Ctor = ctor === "pointer" && typeof view.PointerEvent === "function" ? view.PointerEvent : view.MouseEvent;
        target.dispatchEvent(new Ctor(type, { ...init, bubbles: true, cancelable: true, composed: true,
          ...(ctor === "pointer" ? { pointerId: 1, pointerType: "mouse", isPrimary: true } : {}) }));
      }
    },
  };
}
