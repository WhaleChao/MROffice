/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useRef } from "react";

let override: boolean | null = null;

/**
 * True on iPadOS WebKit (Safari and every other iPad browser): the one
 * platform with Scribble. iPadOS 13+ asks for desktop sites as "Macintosh";
 * its touch points (5) tell it from a Mac (0). Android has no Scribble.
 */
export function pencilScribblePossible(): boolean {
  if (override !== null) return override;
  if (typeof navigator === "undefined") return false;
  const ua = navigator.userAgent ?? "";
  return (
    (navigator.maxTouchPoints ?? 0) > 0 &&
    /AppleWebKit/.test(ua) &&
    /\b(iPad|iPhone|Macintosh)\b/.test(ua)
  );
}

export function setPencilScribblePossibleForTests(value: boolean | null): void {
  override = value;
}

/** Text inputs Scribble writes into (WebKit: text-like controls). */
const TEXT_INPUT_TYPES = new Set(["", "text", "search", "email", "url", "tel", "password", "number"]);

const EDITABLE_SELECTOR =
  '[contenteditable]:not([contenteditable="false"]), textarea, input';

type Lock =
  | { readonly kind: "editable"; readonly was: string; appSaysNo?: true }
  | { readonly kind: "readonly"; readonly was: boolean; appSaysNo?: true };

function isTextInput(element: Element): element is HTMLInputElement {
  return (
    element instanceof HTMLInputElement &&
    TEXT_INPUT_TYPES.has((element.getAttribute("type") ?? "").toLowerCase())
  );
}

/** An element Scribble could write into, by itself (not by inheritance). */
export function scribbleWritableElement(element: Element): boolean {
  if (element instanceof HTMLTextAreaElement) return !element.readOnly && !element.disabled;
  if (element instanceof HTMLInputElement) return isTextInput(element) && !element.readOnly && !element.disabled;
  const value = element.getAttribute("contenteditable");
  return value !== null && value.toLowerCase() !== "false";
}

/**
 * Make every Scribble-writable element under `roots` non-editable until the
 * returned release is called, and keep it so as the area changes. Release
 * gives each element back the state it had, unless the app changed it again
 * meanwhile in a way that is not ours (then the app's value stands).
 */
/**
 * Elements the writer chose to type into while Scribble is held off (a
 * finger or a key took them, not the pencil): every hold leaves them
 * editable until focus leaves them.
 */
const yieldedForTyping = new Set<Element>();








const drawingAreas = new Set<Element>();

export interface HoldScribbleOffOptions {
  /**
   * A finger, mouse or keyboard that takes a held field (a press on it, or
   * focus arriving from the keyboard) gets it back for typing until focus
   * leaves it; the pencil never does. For a hold over a whole page, where
   * a field the writer means to type in can be held.
   */
  readonly yieldToTyping?: boolean;
  /**
   * Never given back for typing, whatever takes them (read at the press):
   * the drawing area of an editor on its Draw tab, where the pencil only
   * ever draws.
   */
  readonly keepHeld?: () => readonly (Element | null | undefined)[];
}

function within(element: Element, area: Element): boolean {
  return area.contains(element) || element.contains(area);
}

/**
 * What Scribble could take near a point (client px), for a page gone deaf to
 * the pencil: whether a drawing area is held, and every element in the box
 * around the point it could still write into (a field outside the held
 * areas, or one given back for typing).
 */
export function scribbleWritableNear(doc: Document, x: number, y: number, radius: number): string {
  const writable: string[] = [];
  for (const element of Array.from(doc.querySelectorAll(EDITABLE_SELECTOR))) {
    if (!scribbleWritableElement(element)) continue;
    const rect = element.getBoundingClientRect();
    if (rect.right < x - radius || rect.left > x + radius || rect.bottom < y - radius || rect.top > y + radius) continue;
    if (rect.width === 0 && rect.height === 0) continue;
    writable.push(`${element.tagName.toLowerCase()}.${String(element.className).slice(0, 24)}${yieldedForTyping.has(element) ? "(typing)" : ""}`);
  }
  return `held=${drawingAreas.size} writable=${writable.length}${writable.length ? ` [${writable.slice(0, 3).join(" ")}]` : ""}`;
}

export function holdScribbleOff(roots: readonly Element[], options: HoldScribbleOffOptions = {}): () => void {
  const yields = options.yieldToTyping === true;
  const locks = new Map<Element, Lock>();
  const lock = (element: Element) => {
    // A hold that never yields takes back what one that does gave away.
    if (locks.has(element) || (yields && yieldedForTyping.has(element)) || !scribbleWritableElement(element)) {
      return;
    }
    if (element instanceof HTMLInputElement || element instanceof HTMLTextAreaElement) {
      locks.set(element, { kind: "readonly", was: element.readOnly });
      element.readOnly = true;
      return;
    }
    locks.set(element, { kind: "editable", was: element.getAttribute("contenteditable") ?? "true" });
    element.setAttribute("contenteditable", "false");
  };
  const sweep = (node: Node) => {
    if (!(node instanceof Element)) return;
    if (node.matches(EDITABLE_SELECTOR)) lock(node);
    for (const element of node.querySelectorAll(EDITABLE_SELECTOR)) lock(element);
    // The one shadow-DOM editor in the editors' pages: an equation field.
    for (const field of node.querySelectorAll("math-field")) {
      const shadow = (field as HTMLElement).shadowRoot;
      if (!shadow) continue;
      for (const element of shadow.querySelectorAll(EDITABLE_SELECTOR)) lock(element);
    }
  };
  for (const root of roots) sweep(root);
  // No text selection either while the tool is picked: the caret the
  // document kept (its editables now held non-editable) was the one thing an
  // Excalidraw page does not have under the Pencil, and with it Scribble in
  
  // Excalidraw never stalled, the document did; every stall report read
  // `selection=Caret`). The ranges come back when the tool is put down, if
  // nothing was selected meanwhile. ProseMirror reads a selection change only
  // while it has focus, which a picked ink tool has already taken.
  const view = roots[0]?.ownerDocument?.defaultView ?? null;
  const selection = view?.getSelection?.() ?? null;
  const keptRanges: Range[] = [];
  if (selection && selection.rangeCount > 0) {
    for (let index = 0; index < selection.rangeCount; index += 1) keptRanges.push(selection.getRangeAt(index).cloneRange());
    selection.removeAllRanges();
  }
  const restoreSelection = () => {
    if (!selection || keptRanges.length === 0 || selection.rangeCount > 0) return;
    for (const range of keptRanges) {
      if (range.startContainer.isConnected && range.endContainer.isConnected) selection.addRange(range);
    }
  };
  const handle = (records: readonly MutationRecord[]) => {
    for (const record of records) {
      if (record.type === "childList") {
        for (const added of record.addedNodes) sweep(added);
        continue;
      }
      const target = record.target;
      if (!(target instanceof Element)) continue;
      if (!scribbleWritableElement(target)) {
        // The app itself made a held element non-editable (a text box
        // left editing while the pen was picked): it wrote over our
        // value, so the old value was ours. Release leaves it so.
        const held = locks.get(target);
        const ours =
          record.attributeName === "contenteditable" ? record.oldValue === "false" : record.oldValue !== null;
        if (held && ours) held.appSaysNo = true;
        continue;
      }
      // Given back for typing: it stays editable until focus leaves it
      // (not in a drawing area: see `lock`).
      if (yields && yieldedForTyping.has(target)) continue;
      // The app (or ProseMirror, when its own value changed) made an
      // element editable again: it is taken again, its new state the
      // one to give back.
      locks.delete(target);
      lock(target);
    }
  };
  const observer = typeof MutationObserver === "undefined" ? null : new MutationObserver(handle);
  for (const root of roots) {
    observer?.observe(root, {
      subtree: true,
      childList: true,
      attributes: true,
      attributeOldValue: true,
      attributeFilter: ["contenteditable", "readonly"],
    });
  }
  // Typing: the held field a finger, the mouse or the keyboard takes is
  // given back, and held again once focus has left it.
  const keptHeld = (element: Element): boolean => {
    for (const area of drawingAreas) if (within(element, area)) return true;
    for (const area of options.keepHeld?.() ?? []) if (area && within(element, area)) return true;
    return false;
  };
  /** What this hold gave back, and how to stop waiting for its focus to leave. */
  const given = new Map<Element, () => void>();
  const give = (element: Element) => {
    const held = locks.get(element);
    if (!held || keptHeld(element)) return;
    locks.delete(element);
    yieldedForTyping.add(element);
    if (held.kind === "readonly") (element as HTMLInputElement | HTMLTextAreaElement).readOnly = held.was;
    else element.setAttribute("contenteditable", held.was);
    const ownerDocument = element.ownerDocument;
    const onFocusOut = () => {
      // After the focus has moved: still inside (a caret moving between
      // its own parts) keeps it.
      setTimeout(() => {
        const active = ownerDocument.activeElement;
        if (active && (active === element || element.contains(active))) return;
        element.removeEventListener("focusout", onFocusOut);
        given.delete(element);
        yieldedForTyping.delete(element);
        if (!released) lock(element);
      }, 0);
    };
    element.addEventListener("focusout", onFocusOut);
    given.set(element, () => element.removeEventListener("focusout", onFocusOut));
  };
  const heldAround = (target: EventTarget | null): Element | null => {
    for (let node = target instanceof Element ? target : null; node; node = node.parentElement) {
      if (locks.has(node)) return node;
    }
    return null;
  };
  let lastKeyAt = Number.NEGATIVE_INFINITY;
  const onKeyDown = (event: KeyboardEvent) => {
    lastKeyAt = event.timeStamp;
  };
  const onPointerDown = (event: PointerEvent) => {
    if (event.pointerType === "pen") return;
    const held = heldAround(event.target);
    if (held) give(held);
  };
  const onFocusIn = (event: FocusEvent) => {
    // Focus the keyboard moved (Tab) is the writer's; focus the app moved
    // itself (a panel opening, a field taking focus on its own) is not.
    if (event.timeStamp - lastKeyAt > 300) return;
    const held = heldAround(event.target);
    if (held) give(held);
  };
  const documents = new Set(roots.map((root) => root.ownerDocument));
  if (!yields) for (const root of roots) drawingAreas.add(root);
  if (options.yieldToTyping) {
    for (const doc of documents) {
      doc.addEventListener("pointerdown", onPointerDown, true);
      doc.addEventListener("keydown", onKeyDown, true);
      doc.addEventListener("focusin", onFocusIn, true);
    }
  }
  let released = false;
  return () => {
    if (released) return;
    released = true;
    if (!yields) for (const root of roots) drawingAreas.delete(root);
    if (options.yieldToTyping) {
      for (const doc of documents) {
        doc.removeEventListener("pointerdown", onPointerDown, true);
        doc.removeEventListener("keydown", onKeyDown, true);
        doc.removeEventListener("focusin", onFocusIn, true);
      }
    }
    // What the app did just before (not delivered yet) counts too.
    if (observer) handle(observer.takeRecords());
    observer?.disconnect();
    for (const [element, held] of locks) {
      if (held.appSaysNo) continue;
      if (held.kind === "readonly") {
        const input = element as HTMLInputElement | HTMLTextAreaElement;
        if (input.readOnly) input.readOnly = held.was;
      } else if (element.getAttribute("contenteditable") === "false") {
        element.setAttribute("contenteditable", held.was);
      }
    }
    locks.clear();
    restoreSelection();
    // A field given back that focus never entered (a finger that only
    // scrolled over it) is not left marked: the next hold takes it.
    for (const [element, stop] of given) {
      stop();
      yieldedForTyping.delete(element);
    }
    given.clear();
  };
}

/**
 * While `active` (an ink tool is picked) and Scribble exists, the elements
 * `roots` names hold no Scribble-writable element. `roots` is read when the
 * hold begins; `key` restarts it (a different drawing area).
 */
export function useInkScribbleGuard(
  roots: () => readonly (Element | null | undefined)[],
  active: boolean,
  key?: unknown,
): void {
  const rootsRef = useRef(roots);
  // (Effects run in order: the hold below reads this render's `roots`.)
  useEffect(() => {
    rootsRef.current = roots;
  });
  useEffect(() => {
    if (!active || !pencilScribblePossible()) return;
    const present = rootsRef.current().filter((root): root is Element => root instanceof Element);
    if (present.length === 0) return;
    return holdScribbleOff(present);
  }, [active, key]);
}












export function useDrawTabScribbleGuard(
  onDrawTab: boolean,
  drawingArea?: () => readonly (Element | null | undefined)[],
): void {
  const areaRef = useRef(drawingArea);
  useEffect(() => {
    areaRef.current = drawingArea;
  });
  useEffect(() => {
    if (!onDrawTab || !pencilScribblePossible() || typeof document === "undefined") return;
    return holdScribbleOff([document.body], { yieldToTyping: true, keepHeld: () => areaRef.current?.() ?? [] });
  }, [onDrawTab]);
}
