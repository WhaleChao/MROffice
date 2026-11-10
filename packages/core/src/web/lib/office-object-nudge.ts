/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The keyboard on a selected drawing object, one reading for the three
 * editors (Docs' floating shape, the deck's block, the worksheet's object).
 *
 * Office semantics (Microsoft Support "Move a text box, WordArt, or shape",
 * "Use keyboard shortcuts to create PowerPoint presentations"; the Office
 * drawing team's "Nudging" note on Microsoft Learn): the OfficeArt layer
 * behaves the same in PowerPoint, Excel and Word's floating shapes --
 *
 *  - an arrow key nudges the object a set distance;
 *  - Ctrl + arrow nudges it ONE SCREEN PIXEL, so what that is on the page
 *    depends on the zoom (a smaller step at 400 %, a larger one at 25 %);
 *  - Shift + arrow resizes it (Right / Up enlarge, Left / Down reduce);
 *  - Alt + Left / Right rotates it 15 degrees counter-clockwise / clockwise;
 *  - a Word shape in line with text is not nudged (Word places it, not
 *    OfficeArt), which is the Docs adapter's to refuse.
 *
 * The set distance is MRAI's: 8 CSS px, a twelfth of an inch at 96 dpi,
 * PowerPoint's default grid pitch; the hosts' own snapping stays theirs.
 */

/** The plain arrow step, in the host's document pixels (1/12 in at 96 dpi). */
export const OFFICE_NUDGE_STEP_PX = 8;
/** Alt + Left / Right. */
export const OFFICE_ROTATE_STEP_DEGREES = 15;

export type OfficeObjectKeyAction =
  /** Move by (dx, dy) document pixels. */
  | { readonly kind: "move"; readonly dx: number; readonly dy: number }
  /** Grow (positive) or shrink (negative) by (dw, dh) document pixels,
   * keeping the top-left corner. */
  | { readonly kind: "resize"; readonly dw: number; readonly dh: number }
  /** Turn by `degrees` (clockwise positive). */
  | { readonly kind: "rotate"; readonly degrees: number };

export interface OfficeObjectKeyEvent {
  readonly key: string;
  readonly shiftKey?: boolean;
  readonly ctrlKey?: boolean;
  readonly metaKey?: boolean;
  readonly altKey?: boolean;
}

const ARROWS: Readonly<Record<string, readonly [number, number]>> = {
  ArrowLeft: [-1, 0],
  ArrowRight: [1, 0],
  ArrowUp: [0, -1],
  ArrowDown: [0, 1],
};












export function officeObjectKeyAction(
  event: OfficeObjectKeyEvent,
  zoom = 1,
  options: {
    readonly plainStep?: "set-distance" | "screen-pixel";
    

    readonly shiftArrow?: "resize" | "move";
  } = {},
): OfficeObjectKeyAction | null {
  const arrow = ARROWS[event.key];
  if (!arrow) return null;
  if (event.altKey) {
    if (arrow[0] === 0 || event.shiftKey || event.ctrlKey || event.metaKey) return null;
    return { kind: "rotate", degrees: arrow[0] * OFFICE_ROTATE_STEP_DEGREES };
  }
  const safeZoom = Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
  const step =
    event.ctrlKey || event.metaKey || options.plainStep === "screen-pixel"
      ? 1 / safeZoom
      : OFFICE_NUDGE_STEP_PX;
  if (event.shiftKey && options.shiftArrow !== "move") {
    // Right / Up enlarge, Left / Down reduce.
    return arrow[0] !== 0
      ? { kind: "resize", dw: arrow[0] * step, dh: 0 }
      : { kind: "resize", dw: 0, dh: -arrow[1] * step };
  }
  return { kind: "move", dx: arrow[0] * step, dy: arrow[1] * step };
}

/**
 * The keys that put a selected shape into text editing, one reading for the
 * three editors (Microsoft Support "Keyboard shortcuts for working with
 * shapes, text boxes, and WordArt": Enter selects the text within the
 * selected object; "Add text to a shape": select the shape and start
 * typing). Enter and F2 open the editor with the caret after the words; a
 * printable key opens it typing that key, which replaces the words as a
 * typed key replaces a selection. Null for every other key, for chords
 * (Ctrl / Cmd / Alt are shortcuts) and for a key an input method is still
 * composing -- the host leaves those to its own handling, but it must never
 * let a typed key REPLACE the selected object.
 */
export type OfficeObjectTextKey =
  | { readonly kind: "edit" }
  | { readonly kind: "type"; readonly seed: string };

export function officeObjectTextKey(
  event: OfficeObjectKeyEvent & { readonly isComposing?: boolean },
): OfficeObjectTextKey | null {
  if (event.ctrlKey || event.metaKey || event.altKey || event.isComposing) return null;
  if (event.key === "Enter" || event.key === "F2") {
    return event.shiftKey && event.key === "Enter" ? null : { kind: "edit" };
  }
  // One user-perceived character: `key` names a non-printing key by a word
  // ("Tab", "Escape", "Process"), a printable one by the character itself.
  return [...event.key].length === 1 ? { kind: "type", seed: event.key } : null;
}
