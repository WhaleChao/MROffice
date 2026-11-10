/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { EditorView } from "@tiptap/pm/view";

/**
 * What is typed into a document window before its editor exists goes in when the
 * pages show (Word's type-ahead): a window opens at once and its editor a moment
 * later, and the keys of that moment reached nothing.
 *
 * Only plain characters, only in the window in front, and only while nothing else
 * has the focus: a key meant for another field is that field's.
 */
const LONGEST_HOLD_MS = 30_000;

/** Whether `event` types a character, with nothing else (the holder's rule). */
export function keyTypesCharacter(event: KeyboardEvent): boolean {
  return (
    !event.defaultPrevented &&
    !event.isComposing &&
    !event.metaKey &&
    !event.ctrlKey &&
    !event.altKey &&
    event.key.length === 1
  );
}

/** Whether the keys of this moment belong to `host` (its window, or the page). */
export function hostHasTheKeyboard(host: Element | null | undefined, doc: Document): boolean {
  if (!host) return false;
  const active = doc.activeElement;
  return !active || active === doc.body || host.contains(active);
}

let held: { windowId: string; text: string; stop: () => void } | null = null;

function windowElement(doc: Document, windowId: string): HTMLElement | null {
  return doc.querySelector<HTMLElement>(`.wm-window[data-window-id="${CSS.escape(windowId)}"]`);
}

/** Starts holding the keys typed into window `windowId` until the editor takes them. */
export function holdKeysWhileWindowOpens(windowId: string): void {
  if (typeof document === "undefined" || held) return;
  const doc = document;
  const onKey = (event: KeyboardEvent) => {
    if (!held || !keyTypesCharacter(event)) return;
    const host = windowElement(doc, windowId);
    if (!host?.classList.contains("wm-window--focused") || !hostHasTheKeyboard(host, doc)) return;
    held.text += event.key;
  };
  const timer = setTimeout(() => held?.stop(), LONGEST_HOLD_MS);
  doc.addEventListener("keydown", onKey, true);
  held = {
    windowId,
    text: "",
    stop: () => {
      clearTimeout(timer);
      doc.removeEventListener("keydown", onKey, true);
      held = null;
    },
  };
}

/** The window around `element`, for `returnKeysToWindow`. */
export function windowIdAround(element: Element): string | null {
  return element.closest<HTMLElement>(".wm-window[data-window-id]")?.dataset.windowId ?? null;
}

/** An editor given keys it could not type before it went (a window's view is built
 * again as its document's room connects) hands them back for the next one. */
export function returnKeysToWindow(windowId: string | null, text: string): void {
  if (!windowId || !text) return;
  if (held?.windowId !== windowId) {
    held?.stop();
    holdKeysWhileWindowOpens(windowId);
  }
  if (held) held.text = text + held.text;
}

/** Whether the window around `element` is holding its keys already: they are not
 * to be held twice by what opens inside it. */
export function windowHoldsKeysFor(element: Element): boolean {
  const windowId = element.closest<HTMLElement>(".wm-window[data-window-id]")?.dataset.windowId;
  return Boolean(held && windowId && held.windowId === windowId);
}

/** Stops holding for the window around `element` and gives what was typed. */
export function takeKeysHeldForWindow(element: Element): string {
  const windowId = element.closest<HTMLElement>(".wm-window[data-window-id]")?.dataset.windowId;
  if (!held || !windowId || held.windowId !== windowId) return "";
  const { text } = held;
  held.stop();
  return text;
}

/** A key held while a document's view is handed over (`holdKeysThroughHandover`):
 * a character, or the editing keys a typist uses between them. */
export type HeldKey = { kind: "text"; text: string } | { kind: "enter"; shift: boolean } |
  { kind: "backspace" } | { kind: "delete" };

/** The key `event` stands for, while held, or null for one that is not held
 * (a shortcut, an arrow, a composition: they go where they would). */
export function heldKeyOf(event: KeyboardEvent): HeldKey | null {
  if (event.defaultPrevented || event.isComposing || event.metaKey || event.ctrlKey) return null;
  if (event.key.length === 1) return event.altKey ? null : { kind: "text", text: event.key };
  if (event.altKey) return null;
  if (event.key === "Enter") return { kind: "enter", shift: event.shiftKey };
  if (event.shiftKey) return null;
  if (event.key === "Backspace") return { kind: "backspace" };
  if (event.key === "Delete") return { kind: "delete" };
  return null;
}

/** Whether a key aimed at `target` is the document's, not a field's of its own
 * (a title, a find box, a dialog's input). */
function keyIsTheDocuments(target: EventTarget | null, host: Element): boolean {
  if (!(target instanceof Element)) return true;
  if (target === target.ownerDocument.body || target === target.ownerDocument.documentElement) return true;
  if (!host.contains(target)) return false;
  if (target.closest("input, textarea, select, [role='dialog']")) return false;
  // Another editable surface in the window (a comment, a field) keeps its keys;
  // the document's own editor (.ProseMirror) is the one being handed over.
  const editable = target.closest<HTMLElement>("[contenteditable='true']");
  return !editable || editable.classList.contains("ProseMirror");
}

/** `keys` as they read once typed: a Backspace takes back the character held
 * just before it (the typist corrected what the page did not show yet). */
function heldWithCorrection(keys: HeldKey[], key: HeldKey): void {
  const last = keys[keys.length - 1];
  if (key.kind === "backspace" && last?.kind === "text") {
    if (last.text.length > 1) last.text = last.text.slice(0, -1);
    else keys.pop();
    return;
  }
  if (key.kind === "text" && last?.kind === "text") {
    last.text += key.text;
    return;
  }
  keys.push(key);
}

/** The keys a document window holds while its view is handed over, until taken. */
export interface HandoverKeys {
  /** What was held, in the order typed (characters run together). */
  readonly keys: () => readonly HeldKey[];
  /** Stops holding and gives what was held. */
  readonly take: () => HeldKey[];
  /** Stops holding; what was held is dropped. */
  readonly stop: () => void;
  /** Whether it still holds what is typed (false once taken, stopped, past
   * its time or full: keys typed then go where they would). */
  readonly holding: () => boolean;
  /** Holds `event` as if it had reached the holder itself (a key whose own
   * dispatch is what began the hold: a listener added during it never hears
   * it); true when held. */
  readonly hold?: (event: KeyboardEvent) => boolean;
}

/** How many characters a handover holds at most (a typist's minutes of it):
 * past that the keys go where they would, the view refusing them in sight. */
export const HANDOVER_MOST_CHARACTERS = 20_000;

function heldCharacters(keys: readonly HeldKey[]): number {
  return keys.reduce((count, key) => count + (key.kind === "text" ? key.text.length : 1), 0);
}

/** Keys held before (a closed tab's, kept with its handover copy) given to
 * the handover that types them: they are taken as a live hold's are. */
export function heldKeysGiven(keys: readonly HeldKey[]): HandoverKeys {
  const held = keys.map((key) => ({ ...key }));
  return { keys: () => held, take: () => held.splice(0), stop: () => { held.length = 0; }, holding: () => false };
}

/** Whether `value` reads as held keys (one kept in storage, read back). */
export function heldKeysFrom(value: unknown): HeldKey[] | null {
  if (!Array.isArray(value)) return null;
  const keys: HeldKey[] = [];
  for (const item of value) {
    const kind = item && typeof item === "object" ? (item as { kind?: unknown }).kind : null;
    if (kind === "text" && typeof (item as { text?: unknown }).text === "string" && (item as { text: string }).text) {
      keys.push({ kind, text: (item as { text: string }).text });
    } else if (kind === "enter") keys.push({ kind, shift: (item as { shift?: unknown }).shift === true });
    else if (kind === "backspace" || kind === "delete") keys.push({ kind });
    else return null;
  }
  return keys;
}

/**
 * Holds what is typed into a document while its view is handed over to another
 * (a page read again as parts: its room retired, the view built anew), from
 * the moment the old view's edits are taken until the new view has its caret.
 * Word never drops a key it was typed: neither does this. Every held key is
 * kept from the old view too (prevented): what it had was taken already, so a
 * key that reached it afterwards would be in neither view.
 *
 * Held only while the keyboard is the document's: its window in front, nothing
 * else focused (`hostHasTheKeyboard`), the key not a field's of its own. After
 * `longestMs`, or HANDOVER_MOST_CHARACTERS, it holds no more; what it held is
 * kept until taken (`onHeld` is told of each key held, to keep them past the
 * tab: autorecover-store noteAutoRecoverHandoverKeys).
 */
export function holdKeysThroughHandover(
  host: () => Element | null,
  longestMs = 180_000,
  onHeld?: (keys: readonly HeldKey[]) => void,
): HandoverKeys {
  const held: HeldKey[] = [];
  if (typeof document === "undefined") {
    return { keys: () => held, take: () => held.splice(0), stop: () => undefined, holding: () => false };
  }
  const doc = document;
  let stopped = false;
  const stop = () => {
    if (stopped) return;
    stopped = true;
    clearTimeout(timer);
    doc.removeEventListener("keydown", onKey, true);
  };
  const onKey = (event: KeyboardEvent): boolean => {
    if (stopped) return false;
    const key = heldKeyOf(event);
    const element = host();
    if (!key || !element) return false;
    const frame = element.closest(".wm-window");
    if (frame && !frame.classList.contains("wm-window--focused")) return false;
    if (!hostHasTheKeyboard(frame ?? element, doc) || !keyIsTheDocuments(event.target, frame ?? element)) return false;
    if (heldCharacters(held) >= HANDOVER_MOST_CHARACTERS) {
      stop();
      return false;
    }
    event.preventDefault();
    event.stopImmediatePropagation();
    heldWithCorrection(held, key);
    onHeld?.(held);
    return true;
  };
  // Past its time it holds no more; what it held waits to be taken.
  const timer = setTimeout(stop, longestMs);
  doc.addEventListener("keydown", onKey, true);
  return {
    keys: () => held,
    take: () => { stop(); return held.splice(0); },
    stop: () => { stop(); held.length = 0; },
    holding: () => !stopped,
    hold: onKey,
  };
}

/** Whether `event` is a key a document's handover would hold for the editor
 * in `host` (its window in front, nothing else focused, not a field's own). */
export function keyHeldForDocument(event: KeyboardEvent, host: Element | null): boolean {
  if (!host || typeof document === "undefined" || !heldKeyOf(event)) return false;
  const frame = host.closest(".wm-window");
  if (frame && !frame.classList.contains("wm-window--focused")) return false;
  return hostHasTheKeyboard(frame ?? host, document) && keyIsTheDocuments(event.target, frame ?? host);
}

/** The view `typeHeldKeys` types into (a TipTap editor's, or a bare one). */
export interface HeldKeysTarget {
  readonly view: {
    readonly state: { readonly selection: { readonly from: number; readonly to: number }; readonly tr: { insertText(text: string, from?: number, to?: number): unknown } };
    dispatch(tr: never): void;
    someProp(name: "handleTextInput", f: (handler: (view: never, from: number, to: number, text: string, deflt: () => unknown) => boolean | void) => boolean | void): boolean | void;
  };
  readonly commands: { keyboardShortcut(name: string): boolean };
}

/** Types `keys` at the caret of `editor` as the keyboard would: a character
 * through the view's text input (input rules, tracked changes), the editing
 * keys through its key bindings. */
export function typeHeldKeys(editor: HeldKeysTarget, keys: readonly HeldKey[]): void {
  const { view } = editor;
  for (const key of keys) {
    if (key.kind !== "text") {
      editor.commands.keyboardShortcut(key.kind === "enter" ? (key.shift ? "Shift-Enter" : "Enter")
        : key.kind === "backspace" ? "Backspace" : "Delete");
      continue;
    }
    for (const character of Array.from(key.text)) {
      const { from, to } = view.state.selection;
      const deflt = () => view.state.tr.insertText(character, from, to);
      const handled = view.someProp("handleTextInput", (handler) => handler(view as never, from, to, character, deflt));
      if (!handled) view.dispatch(deflt() as never);
    }
  }
}








export function viewTakesInputNow(view: EditorView): boolean {
  return !view.isDestroyed && !view.someProp("editable", (editable) => editable(view.state) === false);
}
