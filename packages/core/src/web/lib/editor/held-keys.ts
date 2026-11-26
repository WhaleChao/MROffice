/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { Selection } from "@tiptap/pm/state";
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

/**
 * The window hold (a document window opening: a file double-clicked, an import's
 * window, a page's window before its editor exists). It starts as soon as the window
 * is up -- before the page exists, while the file is still being read -- and holds
 * the keys a typist uses (characters, Enter, Backspace, Delete, Tab, the arrows, the
 * formatting shortcuts) in the order typed, until the document's editor takes them
 * (`takeHeldKeysForWindow`) once it takes input. A window that becomes another (an
 * import's window swapped for its page's, `in_place_of`) passes its keys on.
 * Nothing it held is ever dropped silently: past its time it stops listening and
 * keeps what it has for the editor, or for the window to show (`heldKeysText`).
 * While nothing else has the focus the window gets a capture field of its own, so an
 * input method composes into it and its committed text is held whole.
 */
let held: {
  windowId: string;
  keys: HeldKey[];
  listening: boolean;
  stop: () => void;
  /** The capture field placed again in the window that holds now. */
  placeCapture: () => void;
} | null = null;

/** A held shortcut that could not be typed safely later (Undo, Save, ...): not
 * typed, and recorded here (and on `window.__heldKeysDropped`). */
const droppedShortcuts: string[] = [];
export function heldShortcutsDropped(): readonly string[] {
  return droppedShortcuts;
}

/** The shortcuts typed again after the hold: formatting at the caret, nothing that
 * reaches past it (Undo would undo the opening's own steps, Save has no page yet). */
const SAFE_HELD_SHORTCUTS = new Set(["Mod-b", "Mod-i", "Mod-u"]);
const ARROW_KEYS = new Set(["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown"]);

function shortcutName(event: KeyboardEvent): string {
  const mod = event.metaKey || event.ctrlKey ? "Mod-" : "";
  const alt = event.altKey ? "Alt-" : "";
  const shift = event.shiftKey && event.key.length !== 1 ? "Shift-" : "";
  return `${mod}${alt}${shift}${event.key.length === 1 ? event.key.toLowerCase() : event.key}`;
}

/** The key an opening window holds for `event`: "drop" for a shortcut it will not
 * type later (recorded), null for one it leaves alone (a composition's own keys, a
 * modifier alone, a function key). */
export function openingKeyOf(event: KeyboardEvent): HeldKey | "drop" | null {
  if (event.defaultPrevented || event.isComposing || event.key === "Process" || event.keyCode === 229) return null;
  if (event.metaKey || event.ctrlKey) {
    if (event.key.length !== 1 && !ARROW_KEYS.has(event.key) && event.key !== "Backspace" && event.key !== "Delete") return null;
    const name = shortcutName(event);
    return SAFE_HELD_SHORTCUTS.has(name) ? { kind: "shortcut", name } : "drop";
  }
  if (event.key === "Tab" && !event.altKey) return { kind: "tab", shift: event.shiftKey };
  if (ARROW_KEYS.has(event.key) && !event.altKey && !event.shiftKey) return { kind: "arrow", key: event.key as ArrowKey };
  return heldKeyOf(event);
}

function windowElement(doc: Document, windowId: string): HTMLElement | null {
  return doc.querySelector<HTMLElement>(`.wm-window[data-window-id="${CSS.escape(windowId)}"]`);
}

/** The window's capture field (see `held`): only while nothing else has the focus. */
const CAPTURE_ATTR = "data-held-keys-capture";

/** The window id of a hold started before its window exists. */
const OPENING_WINDOW = "";

/**
 * Starts holding the keys typed for a document handed to the app (a Word file
 * double-clicked) before any window shows it: the shell hands the file over,
 * and the window that will show it is made a moment later -- a key typed in
 * between reached the window in front (the home) or nothing (the 1-second
 * keygate under load: 5 of 40 typed in the ~0.3 s before the import's window had
 * mounted, all lost). The first window that opens and holds keys
 * (holdKeysWhileWindowOpens) takes these over with everything held since.
 */
export function holdKeysForOpeningDocument(): void {
  holdKeysWhileWindowOpens(OPENING_WINDOW);
}

/** Starts holding the keys typed into window `windowId` until the editor takes them.
 * Held already for a window that is gone (it became this one): its keys carry over. */
export function holdKeysWhileWindowOpens(windowId: string): void {
  if (typeof document === "undefined") return;
  const doc = document;
  if (held) {
    if (held.windowId !== windowId && !windowElement(doc, held.windowId)) {
      held.windowId = windowId;
      if (!held.listening) held.listening = true;
      // The field went with the window it stood in: an input method composing
      // in the new window composes into a field of its own again.
      held.placeCapture();
    }
    return;
  }
  let capture: HTMLTextAreaElement | null = null;
  let composing = false;
  const take = (key: HeldKey) => {
    if (!held) return;
    heldWithCorrection(held.keys, key);
  };
  const ours = (event: Event): HTMLElement | null => {
    // Held for a document whose window does not exist yet (holdKeysForOpeningDocument):
    // the keyboard is the opening's while nothing else has the focus.
    if ((held?.windowId ?? windowId) === OPENING_WINDOW) {
      const active = doc.activeElement;
      if (active && active !== doc.body && active !== capture) return null;
      const target = event.target;
      if (target instanceof Element && target !== capture && target !== doc.body) return null;
      return doc.body;
    }
    const host = windowElement(doc, held?.windowId ?? windowId);
    if (!host?.classList.contains("wm-window--focused") || !hostHasTheKeyboard(host, doc)) return null;
    const target = event.target;
    // A field of the window's own (a title, a find box) keeps its keys.
    if (target instanceof Element && target !== capture && target !== doc.body && target.closest("input, textarea, select, [contenteditable='true'], [role='dialog']")) return null;
    return host;
  };
  const placeCapture = () => {
    const opening = (held?.windowId ?? windowId) === OPENING_WINDOW;
    const host = opening ? doc.body : windowElement(doc, held?.windowId ?? windowId);
    if (!host) return;
    if (capture && capture.isConnected) {
      if (host.contains(capture)) return;
      // The opening's field (in the body) gives way to the window's own; a
      // composition still going in it ends there (its committed text is held).
      capture.remove();
      capture = null;
      composing = false;
    }
    const active = doc.activeElement;
    if (active && active !== doc.body) return;
    capture = doc.createElement("textarea");
    capture.setAttribute(CAPTURE_ATTR, "");
    capture.setAttribute("aria-hidden", "true");
    capture.tabIndex = -1;
    capture.style.cssText = "position:absolute;left:0;top:0;width:1px;height:1px;opacity:0;pointer-events:none;resize:none;border:0;padding:0;";
    capture.addEventListener("compositionstart", () => { composing = true; });
    capture.addEventListener("compositionend", (event) => {
      composing = false;
      const text = (event as CompositionEvent).data;
      if (held?.listening && text) take({ kind: "text", text });
      if (capture) capture.value = "";
    });
    capture.addEventListener("input", (event) => {
      const input = event as InputEvent;
      if (!held?.listening || composing || input.isComposing) return;
      if (input.inputType === "insertText" && input.data) take({ kind: "text", text: input.data });
      if (capture) capture.value = "";
    });
    host.appendChild(capture);
    capture.focus({ preventScroll: true });
  };
  const onKey = (event: KeyboardEvent) => {
    if (!held?.listening) return;
    if (!ours(event)) return;
    if (event.target === capture && (composing || event.isComposing || event.key === "Process" || event.keyCode === 229)) return;
    const key = openingKeyOf(event);
    if (key === null) return;
    if (key === "drop") {
      const name = shortcutName(event);
      droppedShortcuts.push(name);
      (window as Window & { __heldKeysDropped?: string[] }).__heldKeysDropped = droppedShortcuts.slice(-50);
      // Not prevented: the app's own shortcut (close the window, the menus) still acts.
      return;
    }
    // A character typed into the capture field arrives as its input (which an
    // input method's text does too): held there, once.
    if (event.target === capture && key.kind === "text") return;
    event.preventDefault();
    take(key);
  };
  const timer = setTimeout(() => {
    // Past its time it listens no more; what it held waits for the editor.
    if (held) held.listening = false;
    capture?.remove();
  }, LONGEST_HOLD_MS);
  doc.addEventListener("keydown", onKey, true);
  held = {
    windowId,
    keys: [],
    listening: true,
    stop: () => {
      clearTimeout(timer);
      doc.removeEventListener("keydown", onKey, true);
      const hadFocus = capture !== null && doc.activeElement === capture;
      capture?.remove();
      capture = null;
      if (hadFocus) (doc.activeElement as HTMLElement | null)?.blur?.();
      held = null;
    },
    placeCapture: () => placeCapture(),
  };
  placeCapture();
}

/** The window around `element`, for `returnKeysToWindow`. */
export function windowIdAround(element: Element): string | null {
  return element.closest<HTMLElement>(".wm-window[data-window-id]")?.dataset.windowId ?? null;
}

/** An editor given keys it could not type before it went (a window's view is built
 * again as its document's room connects) hands them back for the next one. */
export function returnHeldKeysToWindow(windowId: string | null, keys: readonly HeldKey[]): void {
  if (!windowId || keys.length === 0) return;
  if (held?.windowId !== windowId) {
    held?.stop();
    holdKeysWhileWindowOpens(windowId);
  }
  if (held) held.keys = [...keys.map((key) => ({ ...key })), ...held.keys];
}

/** `returnHeldKeysToWindow` for plain text. */
export function returnKeysToWindow(windowId: string | null, text: string): void {
  if (text) returnHeldKeysToWindow(windowId, [{ kind: "text", text }]);
}

/** Whether the window around `element` is holding its keys already: they are not
 * to be held twice by what opens inside it. */
export function windowHoldsKeysFor(element: Element): boolean {
  const windowId = element.closest<HTMLElement>(".wm-window[data-window-id]")?.dataset.windowId;
  return Boolean(held && held.listening && windowId && held.windowId === windowId);
}

/** Stops holding for the window around `element` and gives what was held. */
export function takeHeldKeysForWindow(element: Element): HeldKey[] {
  const windowId = element.closest<HTMLElement>(".wm-window[data-window-id]")?.dataset.windowId;
  if (!held || !windowId || held.windowId !== windowId) return [];
  const { keys } = held;
  held.stop();
  return keys;
}

/** `takeHeldKeysForWindow`, its characters only (older callers). */
export function takeKeysHeldForWindow(element: Element): string {
  return takeHeldKeysForWindow(element).map((key) => (key.kind === "text" ? key.text : "")).join("");
}

/** What window `windowId` holds, as text (to show a typist who cannot have it
 * typed: the page did not open). Enter reads as a line break. */
export function heldKeysText(windowId: string | null | undefined): string {
  if (!held || !windowId || held.windowId !== windowId) return "";
  return held.keys.map((key) => (key.kind === "text" ? key.text : key.kind === "enter" ? "\n" : key.kind === "tab" ? "\t" : "")).join("");
}

/** A key held while a document's view is handed over (`holdKeysThroughHandover`):
 * a character, or the editing keys a typist uses between them. */
type ArrowKey = "ArrowLeft" | "ArrowRight" | "ArrowUp" | "ArrowDown";
export type HeldKey = { kind: "text"; text: string } | { kind: "enter"; shift: boolean } |
  { kind: "backspace" } | { kind: "delete" } | { kind: "tab"; shift: boolean } |
  { kind: "arrow"; key: ArrowKey } | { kind: "shortcut"; name: string };

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
    else if (kind === "tab") keys.push({ kind, shift: (item as { shift?: unknown }).shift === true });
    else if (kind === "arrow" && ARROW_KEYS.has(String((item as { key?: unknown }).key))) keys.push({ kind, key: (item as { key: ArrowKey }).key });
    else if (kind === "shortcut" && SAFE_HELD_SHORTCUTS.has(String((item as { name?: unknown }).name))) keys.push({ kind, name: (item as { name: string }).name });
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
    if (key.kind === "arrow") {
      moveCaretFor(view as unknown as EditorView, key.key);
      continue;
    }
    if (key.kind !== "text") {
      editor.commands.keyboardShortcut(key.kind === "enter" ? (key.shift ? "Shift-Enter" : "Enter")
        : key.kind === "backspace" ? "Backspace" : key.kind === "delete" ? "Delete"
          : key.kind === "tab" ? (key.shift ? "Shift-Tab" : "Tab") : key.name);
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

/** An arrow key held while the window opened, typed again: the caret one place
 * along (Left / Right) or a line up or down from where it is drawn (Up / Down). */
function moveCaretFor(view: EditorView, key: ArrowKey): void {
  try {
    const { state } = view;
    const head = state.selection.head;
    let target: number | null = null;
    if (key === "ArrowLeft" || key === "ArrowRight") {
      target = Math.max(0, Math.min(state.doc.content.size, head + (key === "ArrowLeft" ? -1 : 1)));
    } else {
      const box = view.coordsAtPos(head);
      const line = Math.max(1, box.bottom - box.top);
      const found = view.posAtCoords({ left: box.left, top: key === "ArrowUp" ? box.top - line / 2 : box.bottom + line / 2 });
      target = found ? found.pos : null;
    }
    if (target === null || target === head) return;
    const next = Selection.near(state.doc.resolve(target), key === "ArrowLeft" || key === "ArrowUp" ? -1 : 1);
    view.dispatch(state.tr.setSelection(next).scrollIntoView());
  } catch {
    // A caret that cannot move there stays where it is.
  }
}








export function viewTakesInputNow(view: EditorView): boolean {
  return !view.isDestroyed && !view.someProp("editable", (editable) => editable(view.state) === false);
}
