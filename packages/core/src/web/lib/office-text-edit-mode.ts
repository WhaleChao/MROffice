"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Which object's words are being typed, for the selection outline Office draws.
 *
 * Office draws a selected text box, shape or SmartArt shape in two states
 * ("Add or remove a border on a text box, shape, or SmartArt graphic",
 * Microsoft Support): a SOLID outline while the object itself is selected,
 * a DASHED one while the insertion point is in its text. The handles stay in
 * both. No inner text rectangle is drawn.
 *
 * Every editor marks the outline it already draws with
 * `data-text-editing="true"` (the shared rule in OfficeDrawingSelection.module.css
 * turns it dashed). An editor that keeps no explicit "editing the words of X"
 * state (the deck's text boxes are always-mounted text areas) reads it from
 * the focus with {@link useOfficeTextEditingOwner}:
 *
 * - the focus in a text entry inside an object: that object's words are typed;
 * - the focus moved into editor chrome (the ribbon, a menu, a task pane, a
 *   dialog) or out of the window: still typed, as Office keeps the text-edit
 *   state while a ribbon command acts on the words;
 * - anything else (the object's border, the canvas, another object): not.
 */

import { useEffect, useRef, useState, type RefObject } from "react";

import {
  OFFICE_CHROME_SELECTOR,
  officeRecentPressTarget,
} from "./office-inactive-selection";

/** The attribute a selection outline wears while its object's words are typed. */
export const OFFICE_TEXT_EDITING_ATTRIBUTE = "data-text-editing";

const TEXT_INPUT_TYPES = new Set(["text", "search", "url", "tel", "email", ""]);

/** Whether `element` takes typed text: a text area, a text input or an editable region. */
export function isOfficeTextEntry(element: Element | null): boolean {
  if (!element) return false;
  if (element instanceof HTMLTextAreaElement) return !element.readOnly && !element.disabled;
  if (element instanceof HTMLInputElement) {
    return TEXT_INPUT_TYPES.has(element.type) && !element.readOnly && !element.disabled;
  }
  if (!(element instanceof HTMLElement)) return false;
  if (element.isContentEditable) return true;
  // The attribute the browser derives `isContentEditable` from (the nearest
  // one decides), for runtimes that do not compute it.
  const editable = element.closest("[contenteditable]");
  return !!editable && editable.getAttribute("contenteditable") !== "false";
}

/**
 * The object a text entry types into: the value of `ownerAttribute` on its
 * closest ancestor inside `root`, or null when `element` is no text entry or
 * lies outside `root`.
 */
export function officeTextEntryOwner(
  element: Element | null,
  ownerAttribute: string,
  root: Element | null,
): string | null {
  if (!element || !root || !root.contains(element) || !isOfficeTextEntry(element)) return null;
  const owner = element.closest(`[${ownerAttribute}]`);
  return owner && root.contains(owner) ? owner.getAttribute(ownerAttribute) : null;
}

function inChrome(target: EventTarget | Element | null): boolean {
  return target instanceof Element && Boolean(target.closest(OFFICE_CHROME_SELECTOR));
}

/**
 * The id (`ownerAttribute`'s value) of the object under `rootRef` whose words
 * are being typed, or null. Follows the focus as described above; the value
 * changes only when the edit begins or ends, so a host re-renders only then.
 */
/**
 * Office ends the typing when a handle of the object is dragged: the words
 * are written and the OBJECT is selected (its outline solid), and the drag
 * resizes or turns the shape, not its text. A handle gesture cancels its
 * press (so the page selects no text), and a cancelled press leaves the
 * focus where it was: the words stayed open, dashed, through the resize.
 *
 * Takes the focus out of a text entry inside `scope` (the editor's canvas),
 * so each editor's own blur writes and closes the words. A focus in editor
 * chrome (a ribbon box) or outside `scope` is left alone.
 */
export function endOfficeTextEditingForHandle(scope: Element | null): void {
  const doc = scope?.ownerDocument ?? (typeof document === "undefined" ? null : document);
  const active = doc?.activeElement ?? null;
  if (!(active instanceof HTMLElement) || !isOfficeTextEntry(active)) return;
  if (scope && !scope.contains(active)) return;
  if (active.closest(OFFICE_CHROME_SELECTOR)) return;
  active.blur();
}

export function useOfficeTextEditingOwner(
  rootRef: RefObject<HTMLElement | null>,
  ownerAttribute: string,
): string | null {
  const [owner, setOwner] = useState<string | null>(null);
  // The owner last asked for: a focus that comes back to the same text (the
  // editor restores it after each commit) asks for nothing, since even a
  // same-value update re-renders the host before bailing out.
  const ownerRef = useRef<string | null>(null);
  useEffect(() => {
    if (typeof document === "undefined") return;
    const doc = document;
    const view = doc.defaultView ?? window;
    let timer: number | null = null;
    const publish = (next: string | null) => {
      if (ownerRef.current === next) return;
      ownerRef.current = next;
      setOwner(next);
    };
    const settle = () => {
      timer = null;
      const active = doc.activeElement;
      const next = officeTextEntryOwner(active, ownerAttribute, rootRef.current);
      if (next) {
        publish(next);
        return;
      }
      // Into the ribbon, a menu, a pane or a dialog (Safari gives a pressed
      // button no focus, so the press decides too), or to another window.
      if (inChrome(active) || inChrome(officeRecentPressTarget()) || !doc.hasFocus()) return;
      publish(null);
    };
    // Once the focus has settled: a focusout comes before the focusin it leads to.
    const schedule = () => {
      if (timer === null) timer = view.setTimeout(settle, 0);
    };
    // A focusin is where the focus settled: it is published at once, with
    // the press or key that moved it (one host render for the press or the
    // Escape, not a second one a task later). Only a focus that leaves for
    // nowhere waits for the focusin it may lead to.
    const arrive = () => {
      if (timer !== null) view.clearTimeout(timer);
      timer = null;
      settle();
    };
    doc.addEventListener("focusin", arrive, true);
    doc.addEventListener("focusout", schedule, true);
    settle();
    return () => {
      doc.removeEventListener("focusin", arrive, true);
      doc.removeEventListener("focusout", schedule, true);
      if (timer !== null) view.clearTimeout(timer);
    };
  }, [ownerAttribute, rootRef]);
  return owner;
}
