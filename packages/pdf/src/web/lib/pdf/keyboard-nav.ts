/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure routing for the PDF viewer's GLOBAL keyboard shortcuts (Acrobat / WPS
 * parity): zoom in / out / reset and page navigation.
 *
 * The component owns a single `keydown` listener; this module decides, for a
 * given keystroke, which viewer ACTION (if any) it maps to — and, crucially,
 * whether the keystroke should be suppressed because focus is in a text field
 * (so typing a "+" into the page box, a note, or a form field never zooms the
 * document). Kept pure (the only DOM type touched is the read-only `target`
 * element, never mutated) so the routing table is unit-testable without a
 * mounted tree, mirroring `find-key.ts`.
 *
 * The find shortcut (Ctrl/Cmd+F) is intentionally NOT handled here — it has its
 * own dedicated handler in `find-key.ts` with bespoke hijack rules.
 */

/** The viewer actions a global shortcut can trigger. */
export type ViewerShortcutAction =
  | "zoom-in"
  | "zoom-out"
  | "zoom-reset"
  | "page-next"
  | "page-prev"
  


  | "screen-next"
  | "screen-prev"
  | "page-first"
  | "page-last"
  | "history-back"
  | "history-forward";

/** The keydown fields the router reads (a structural subset of KeyboardEvent). */
export interface ShortcutKeyEvent {
  key: string;
  metaKey: boolean;
  ctrlKey: boolean;
  altKey: boolean;
  shiftKey: boolean;
}

/**
 * Cmd+Up / Cmd+Down: macOS's keys for the start and the end of a text (what
 * Ctrl+Home / Ctrl+End are elsewhere). An editor on the page takes them as it
 * takes Home and End: left to the browser with the caret already at that end,
 * nothing moved and the key scrolled the file to its first or last page
 * instead, away from the line being typed in. Shift (extending a selection)
 * is the caller's to read.
 */
export function pdfTextEdgeKey(
  event: Pick<ShortcutKeyEvent, "key" | "metaKey" | "ctrlKey" | "altKey">,
): "start" | "end" | null {
  if (!event.metaKey || event.ctrlKey || event.altKey) return null;
  return event.key === "ArrowUp" ? "start" : event.key === "ArrowDown" ? "end" : null;
}

/** Optional per-call routing switches (all default off). */
export interface ViewerShortcutOptions {
  /**
   * Presentation (fullscreen) mode: Space pages forward and Shift+Space pages
   * back, the projector convention. Off outside presentation, where Space is
   * the hold-to-pan chord and must never be consumed as navigation.
   */
  spacePaging?: boolean;
}

/**
 * True when `target` is an element the user types into — a text input, a
 * textarea, a select, or any `contenteditable` host (including its descendants).
 * While such an element is focused, the viewer's letter/symbol/arrow shortcuts
 * must NOT fire so the keystroke reaches the field.
 *
 * Accepts an `EventTarget | null` (what `event.target` is typed as) and narrows
 * defensively: anything that is not an element with the inspected properties is
 * treated as "not a typing target" (false).
 */
/** Input types that never take character keys (their keys are activation). */
const NON_TEXT_INPUT_TYPES = new Set([
  "button",
  "checkbox",
  "color",
  "file",
  "hidden",
  "image",
  "radio",
  "range",
  "reset",
  "submit",
]);









export function isTextEditingTarget(target: EventTarget | null): boolean {
  if (!target || typeof target !== "object") return false;
  const node = target as {
    tagName?: unknown;
    type?: unknown;
    isContentEditable?: unknown;
    closest?: (selector: string) => Element | null;
  };
  if (node.isContentEditable === true) return true;
  const tag =
    typeof node.tagName === "string" ? node.tagName.toUpperCase() : "";
  if (tag === "TEXTAREA" || tag === "SELECT") return true;
  if (tag === "INPUT") {
    const type =
      typeof node.type === "string" ? node.type.toLowerCase() : "text";
    return !NON_TEXT_INPUT_TYPES.has(type);
  }
  if (typeof node.closest === "function") {
    try {
      const host = node.closest(
        'input, textarea, select, [contenteditable="true"]',
      );
      if (!host) return false;
      if (host.tagName.toUpperCase() !== "INPUT") return true;
      const type = (host as HTMLInputElement).type?.toLowerCase() ?? "text";
      return !NON_TEXT_INPUT_TYPES.has(type);
    } catch {
      /* closest with a bad selector — ignore, treat as non-typing */
    }
  }
  return false;
}

export function isTypingTarget(target: EventTarget | null): boolean {
  if (!target || typeof target !== "object") return false;
  const node = target as {
    tagName?: unknown;
    isContentEditable?: unknown;
    closest?: (selector: string) => unknown;
    readOnly?: unknown;
    disabled?: unknown;
    type?: unknown;
  };

  // contenteditable host or descendant.
  if (node.isContentEditable === true) return true;

  const tag =
    typeof node.tagName === "string" ? node.tagName.toUpperCase() : "";

  if (tag === "TEXTAREA" || tag === "SELECT") {
    // A disabled / read-only field can't receive typing, but it is rare and
    // treating it as a typing target is the safe (shortcut-suppressing) choice.
    return true;
  }
  if (tag === "INPUT") {
    // Non-text input types (button / checkbox / radio / range …) don't capture
    // character keys, but the conservative choice is to suppress for any input
    // so a focused field never has a shortcut stolen from under it. Buttons are
    // <button>, not <input type=button>, in this codebase.
    return true;
  }
  // A focusable element that delegates to a hidden contenteditable / input could
  // still be a typing surface; fall back to a closest() probe when available.
  if (typeof node.closest === "function") {
    try {
      if (node.closest('input, textarea, select, [contenteditable="true"]')) {
        return true;
      }
    } catch {
      /* closest with a bad selector — ignore, treat as non-typing */
    }
  }
  return false;
}

/**
 * Map a keydown to a {@link ViewerShortcutAction}, or null when the keystroke is
 * not a viewer shortcut (or must be ignored because focus is in a text field).
 *
 * Bindings (mirroring Acrobat / common readers):
 *   - Zoom: Ctrl/Cmd + "+" / "=" → zoom-in; Ctrl/Cmd + "-" → zoom-out;
 *     Ctrl/Cmd + "0" → zoom-reset (100%). These work even while NOT focused in
 *     the viewer body as long as focus is not in a text field, matching how a
 *     reader app's zoom keys are global; the text-field guard prevents stealing
 *     "+"/"-" from inputs.
 *   - Page nav (NO modifier): PageDown / ArrowDown / ArrowRight → page-next;
 *     PageUp / ArrowUp / ArrowLeft → page-prev; Home → page-first; End →
 *     page-last. Suppressed while a text field is focused (so arrows move the
 *     caret) and when any modifier is held (so Ctrl+Home etc. stay native).
 *
 * `target` is the focused node (event.target); pass it so the text-field guard
 * can run. When omitted, the guard is skipped (caller already vetted focus).
 */
export function resolveViewerShortcut(
  event: ShortcutKeyEvent,
  target?: EventTarget | null,
  options?: ViewerShortcutOptions,
): ViewerShortcutAction | null {
  if (!event || typeof event.key !== "string") return null;
  const typing = target !== undefined ? isTypingTarget(target) : false;
  const mod = event.metaKey || event.ctrlKey;

  // --- Zoom (requires the platform zoom modifier) ---
  if (mod && !event.altKey) {
    // A focused text field owns "+"/"-"/"0"; don't hijack them there.
    if (typing) return null;
    switch (event.key) {
      case "+":
      case "=": // unshifted "+" key on most layouts
        return "zoom-in";
      case "-":
      case "_":
        return "zoom-out";
      case "0":
        return "zoom-reset";
      default:
        return null;
    }
  }

  // --- View history (Alt+Arrow, browser-style back/forward) ---
  if (event.altKey && !mod && !event.shiftKey) {
    // A focused text field owns Alt+arrows (word-wise caret moves on macOS).
    if (typing) return null;
    switch (event.key) {
      case "ArrowLeft":
        return "history-back";
      case "ArrowRight":
        return "history-forward";
      default:
        return null;
    }
  }

  // --- Presentation-mode Space paging (Shift+Space = back) ---
  // Checked BEFORE the no-modifier gate because Shift is meaningful here; a
  // focused text field still owns the Space keystroke.
  if (options?.spacePaging && event.key === " " && !mod && !event.altKey) {
    if (typing) return null;
    return event.shiftKey ? "page-prev" : "page-next";
  }

  // --- Page navigation (no modifiers) ---
  if (mod || event.altKey || event.shiftKey) return null;
  // Arrows / Home / End belong to a focused text field's caret movement.
  if (typing) return null;
  switch (event.key) {
    case "PageDown":
      // A projector remote sends Page Down: in presentation it turns the page.
      return options?.spacePaging ? "page-next" : "screen-next";
    case "ArrowDown":
    case "ArrowRight":
      return "page-next";
    case "PageUp":
      return options?.spacePaging ? "page-prev" : "screen-prev";
    case "ArrowUp":
    case "ArrowLeft":
      return "page-prev";
    case "Home":
      return "page-first";
    case "End":
      return "page-last";
    default:
      return null;
  }
}
